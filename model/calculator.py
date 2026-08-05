"""Direct Python translation of Huishoudprofiel verrekenprijs 2025.xlsx."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from .xlsx_reader import Workbook

HOURS = 8760
DEFAULT_RUNTIME_MODEL = Path(__file__).with_name("runtime_model.json")


@dataclass(frozen=True)
class Assumptions:
    annual_usage_kwh: float
    annual_solar_kwh: float
    energy_tax: float
    tax_reduction: float
    vat: float
    grid_costs: float
    supplier_markup: float
    fixed_supplier_costs: float
    feed_in_costs: float
    battery_capacity: float
    battery_power: float
    initial_charge: float
    battery_losses: float
    max_daily_battery_export: float
    lend_export_price: float
    solar_export_price: float | None


@dataclass(frozen=True)
class Hour:
    timestamp: float
    usage: float
    solar: float
    epex_mwh: float
    simple_peak: bool
    simple_day_max_price: float
    lend_transformer_power: float
    lend_peak: bool


@dataclass(frozen=True)
class ScenarioTotals:
    usage: float
    solar: float
    grid_import: float
    grid_export: float
    market_cost: float
    battery_charge_solar: float = 0.0
    battery_discharge_home: float = 0.0
    battery_export: float = 0.0
    battery_charge_grid: float = 0.0
    final_charge: float = 0.0


@dataclass(frozen=True)
class CostResult:
    gross_usage: float
    solar: float
    self_consumption: float
    grid_import: float
    grid_export: float
    taxable_electricity: float
    market_cost: float
    supplier_markup: float
    energy_tax: float
    fixed_supplier_costs: float
    feed_in_costs: float
    grid_costs: float
    tax_reduction: float
    opex_ex_vat: float
    vat: float
    opex_inc_vat: float
    average_monthly_opex: float


@dataclass(frozen=True)
class ModelResult:
    assumptions: Assumptions
    baseline: ScenarioTotals
    simple: ScenarioTotals
    lend: ScenarioTotals
    baseline_costs: CostResult
    simple_costs: CostResult
    lend_costs: CostResult


def _number(sheet: dict, address: str) -> float:
    return sheet[address].number()


def _coefficient(formula: str | None) -> float:
    if not formula:
        raise ValueError("Expected a profile formula")
    match = re.match(r"([0-9.Ee+-]+)\*Aannames!", formula)
    if not match:
        raise ValueError(f"Unexpected profile formula: {formula}")
    return float(match.group(1))


def load_model(path: str | Path, overrides: dict[str, float] | None = None) -> tuple[Assumptions, list[Hour]]:
    overrides = overrides or {}
    with Workbook(path) as workbook:
        assumptions_sheet = workbook.sheet("Aannames")
        base = workbook.sheet("Model zonder batterij")
        simple = workbook.sheet("Model met batterij (simpel)")
        lend = workbook.sheet("Model batterij LEND")

        values = {
            "annual_usage_kwh": _number(assumptions_sheet, "B5"),
            "annual_solar_kwh": _number(assumptions_sheet, "B6"),
            "energy_tax": _number(assumptions_sheet, "B20"),
            "tax_reduction": _number(assumptions_sheet, "B21"),
            "vat": _number(assumptions_sheet, "B22"),
            "grid_costs": _number(assumptions_sheet, "B23"),
            "supplier_markup": _number(assumptions_sheet, "B24"),
            "fixed_supplier_costs": _number(assumptions_sheet, "B25"),
            "feed_in_costs": _number(assumptions_sheet, "B26"),
            "battery_capacity": _number(assumptions_sheet, "B29"),
            "battery_power": _number(assumptions_sheet, "B30"),
            "initial_charge": _number(assumptions_sheet, "B31"),
            "battery_losses": _number(assumptions_sheet, "B32"),
            "max_daily_battery_export": _number(assumptions_sheet, "B33"),
            "lend_export_price": _number(assumptions_sheet, "B34"),
            # Older model versions have no B38 and use EPEX for solar exports.
            "solar_export_price": (
                _number(assumptions_sheet, "B38")
                if "B38" in assumptions_sheet
                else None
            ),
        }
        unknown = set(overrides) - set(values)
        if unknown:
            raise KeyError(f"Unknown assumptions: {sorted(unknown)}")
        values.update(overrides)
        assumptions = Assumptions(**values)

        hours = []
        for row in range(2, HOURS + 2):
            usage_base = _coefficient(base[f"B{row}"].formula)
            solar_base = _coefficient(base[f"C{row}"].formula)
            hours.append(
                Hour(
                    timestamp=_number(base, f"A{row}"),
                    usage=usage_base * assumptions.annual_usage_kwh / 4000.0,
                    solar=solar_base * assumptions.annual_solar_kwh / 4500.0,
                    epex_mwh=_number(base, f"D{row}"),
                    simple_peak=bool(_number(simple, f"O{row}")),
                    simple_day_max_price=_number(simple, f"P{row}"),
                    lend_transformer_power=_number(lend, f"O{row}"),
                    lend_peak=bool(_number(lend, f"P{row}")),
                )
            )
    return assumptions, hours


def load_standard_profile_model(
    path: str | Path,
    overrides: dict[str, float] | None = None,
) -> tuple[Assumptions, list[Hour]]:
    path = Path(path)
    if not path.exists() and DEFAULT_RUNTIME_MODEL.exists():
        return load_runtime_model(DEFAULT_RUNTIME_MODEL, overrides)

    assumptions, hours = load_model(path, overrides)
    with Workbook(path) as workbook:
        standard = workbook.sheet("Aannames stdprofielen")
        result = []
        for index, hour in enumerate(hours, start=16):
            result.append(
                Hour(
                    timestamp=hour.timestamp,
                    usage=_number(standard, f"E{index}")
                    * assumptions.annual_usage_kwh
                    / 4000.0,
                    solar=_number(standard, f"G{index}")
                    * assumptions.annual_solar_kwh
                    / 4500.0,
                    epex_mwh=hour.epex_mwh,
                    simple_peak=hour.simple_peak,
                    simple_day_max_price=hour.simple_day_max_price,
                    lend_transformer_power=hour.lend_transformer_power,
                    lend_peak=hour.lend_peak,
                )
            )
    return assumptions, result


def _standard_profile_shapes(path: str | Path) -> tuple[list[float], list[float]]:
    path = Path(path)
    if not path.exists() and DEFAULT_RUNTIME_MODEL.exists():
        assumptions, hours = load_runtime_model(DEFAULT_RUNTIME_MODEL)
        return (
            [hour.usage / assumptions.annual_usage_kwh for hour in hours],
            [hour.solar / assumptions.annual_solar_kwh for hour in hours],
        )

    with Workbook(path) as workbook:
        standard = workbook.sheet("Aannames stdprofielen")
        usage_shape = [
            _number(standard, f"E{row}") / 4000.0 for row in range(16, HOURS + 16)
        ]
        solar_shape = [
            _number(standard, f"G{row}") / 4500.0 for row in range(16, HOURS + 16)
        ]
    return usage_shape, solar_shape


def load_runtime_model(
    path: str | Path = DEFAULT_RUNTIME_MODEL,
    overrides: dict[str, float] | None = None,
) -> tuple[Assumptions, list[Hour]]:
    overrides = overrides or {}
    with Path(path).open(encoding="utf-8") as file:
        payload = json.load(file)

    values = dict(payload["assumptions"])
    base_annual_usage = values["annual_usage_kwh"]
    base_annual_solar = values["annual_solar_kwh"]
    unknown = set(overrides) - set(values)
    if unknown:
        raise KeyError(f"Unknown assumptions: {sorted(unknown)}")
    values.update(overrides)
    assumptions = Assumptions(**values)
    usage_factor = 0.0
    if base_annual_usage != 0.0:
        usage_factor = assumptions.annual_usage_kwh / base_annual_usage
    solar_factor = 0.0
    if base_annual_solar != 0.0:
        solar_factor = assumptions.annual_solar_kwh / base_annual_solar
    hours = [
        Hour(
            **(
                hour
                | {
                    "usage": hour["usage"] * usage_factor,
                    "solar": hour["solar"] * solar_factor,
                }
            )
        )
        for hour in payload["hours"]
    ]
    return assumptions, hours


def export_runtime_model(
    source: str | Path,
    target: str | Path = DEFAULT_RUNTIME_MODEL,
) -> None:
    assumptions, hours = load_standard_profile_model(source)
    payload = {
        "source": Path(source).name,
        "hours": [hour.__dict__ for hour in hours],
        "assumptions": assumptions.__dict__,
    }
    with Path(target).open("w", encoding="utf-8") as file:
        json.dump(payload, file, separators=(",", ":"))


def _grid_flows_from_shapes(
    usage_shape: list[float],
    solar_shape: list[float],
    annual_usage_kwh: float,
    annual_solar_kwh: float,
) -> tuple[float, float]:
    grid_import = grid_export = 0.0
    for usage_fraction, solar_fraction in zip(usage_shape, solar_shape):
        net = annual_usage_kwh * usage_fraction - annual_solar_kwh * solar_fraction
        grid_import += max(net, 0.0)
        grid_export += max(-net, 0.0)
    return grid_import, grid_export


def infer_annual_usage_solar_from_grid(
    path: str | Path,
    target_grid_import: float,
    target_grid_export: float,
) -> tuple[float, float]:
    if target_grid_import < 0.0 or target_grid_export < 0.0:
        raise ValueError("Grid import and export must be non-negative")
    if target_grid_import == 0.0 and target_grid_export == 0.0:
        return 0.0, 0.0

    usage_shape, solar_shape = _standard_profile_shapes(path)
    net_consumption = target_grid_import - target_grid_export
    lower_solar = max(-net_consumption, 0.0)

    def flows_for_solar(annual_solar_kwh: float) -> tuple[float, float]:
        annual_usage_kwh = annual_solar_kwh + net_consumption
        return _grid_flows_from_shapes(
            usage_shape,
            solar_shape,
            annual_usage_kwh,
            annual_solar_kwh,
        )

    lower_export = flows_for_solar(lower_solar)[1]
    if target_grid_export <= lower_export:
        return lower_solar + net_consumption, lower_solar

    upper_solar = max(lower_solar + target_grid_import + target_grid_export, 1.0)
    while flows_for_solar(upper_solar)[1] < target_grid_export:
        upper_solar *= 2.0

    for _ in range(80):
        solar = (lower_solar + upper_solar) / 2.0
        _, grid_export = flows_for_solar(solar)
        if grid_export < target_grid_export:
            lower_solar = solar
        else:
            upper_solar = solar

    annual_solar = (lower_solar + upper_solar) / 2.0
    return annual_solar + net_consumption, annual_solar


def _baseline(hours: Iterable[Hour], a: Assumptions) -> ScenarioTotals:
    usage = solar = grid_import = grid_export = market_cost = 0.0
    for hour in hours:
        net = hour.usage - hour.solar
        usage += hour.usage
        solar += hour.solar
        grid_import += max(net, 0.0)
        grid_export += max(-net, 0.0)
        epex = hour.epex_mwh / 1000.0
        market_cost += max(net, 0.0) * epex - max(-net, 0.0) * epex
    return ScenarioTotals(usage, solar, grid_import, grid_export, market_cost)


def _battery(hours: list[Hour], a: Assumptions, strategy: str) -> ScenarioTotals:
    soc = a.initial_charge
    usage = solar = grid_import = grid_export = market_cost = 0.0
    charged_solar = discharged_home = exported_battery = charged_grid = 0.0
    surpluses = [max(hour.solar - hour.usage, 0.0) for hour in hours]

    for index, hour in enumerate(hours):
        surplus = surpluses[index]
        deficit = max(hour.usage - hour.solar, 0.0)
        epex = hour.epex_mwh / 1000.0
        marginal_buy_price = (epex + a.supplier_markup + a.energy_tax) * (1.0 + a.vat)
        threshold = hour.simple_day_max_price if strategy == "simple" else a.lend_export_price

        discharge_home = (
            0.0
            if marginal_buy_price < threshold
            else min(deficit, a.battery_power, soc)
        )
        peak = hour.simple_peak if strategy == "simple" else hour.lend_peak
        battery_export = (
            min(
                a.max_daily_battery_export,
                max(a.battery_power - discharge_home, 0.0),
                max(soc - discharge_home, 0.0),
            )
            if peak
            else 0.0
        )
        solar_charge = (
            0.0
            if battery_export > 0.0
            else min(surplus, a.battery_power, max(a.battery_capacity - soc, 0.0))
        )

        grid_charge = 0.0
        if marginal_buy_price < threshold and battery_export == 0.0:
            future_surplus = sum(surpluses[index + 1 : index + 25])
            reserved_capacity = max(
                a.battery_capacity - min(a.battery_capacity, future_surplus), 0.0
            )
            grid_charge = min(
                max(a.battery_power - solar_charge, 0.0),
                max(
                    reserved_capacity
                    - (soc + solar_charge - discharge_home - battery_export),
                    0.0,
                ),
            )

        soc = soc + solar_charge + grid_charge - discharge_home - battery_export
        imported = deficit - discharge_home + grid_charge
        exported = surplus - solar_charge + battery_export
        if strategy == "simple":
            energy_cost = (imported - (surplus - solar_charge) - battery_export) * epex
        else:
            solar_export_price = (
                epex if a.solar_export_price is None else a.solar_export_price
            )
            energy_cost = (
                imported * epex
                - (surplus - solar_charge) * solar_export_price
                - battery_export * a.lend_export_price
            )

        usage += hour.usage
        solar += hour.solar
        grid_import += imported
        grid_export += exported
        market_cost += energy_cost
        charged_solar += solar_charge
        discharged_home += discharge_home
        exported_battery += battery_export
        charged_grid += grid_charge

    return ScenarioTotals(
        usage=usage,
        solar=solar,
        grid_import=grid_import,
        grid_export=grid_export,
        market_cost=market_cost,
        battery_charge_solar=charged_solar,
        battery_discharge_home=discharged_home,
        battery_export=exported_battery,
        battery_charge_grid=charged_grid,
        final_charge=soc,
    )


def _costs(totals: ScenarioTotals, a: Assumptions) -> CostResult:
    supplier_markup = totals.grid_import * a.supplier_markup
    energy_tax = totals.grid_import * a.energy_tax
    tax_reduction = -a.tax_reduction
    opex_ex_vat = sum(
        (
            totals.market_cost,
            supplier_markup,
            energy_tax,
            a.fixed_supplier_costs,
            a.feed_in_costs,
            a.grid_costs,
            tax_reduction,
        )
    )
    vat = opex_ex_vat * a.vat
    opex_inc_vat = opex_ex_vat + vat
    return CostResult(
        gross_usage=totals.usage,
        solar=totals.solar,
        self_consumption=totals.usage - totals.grid_import,
        grid_import=totals.grid_import,
        grid_export=totals.grid_export,
        taxable_electricity=totals.grid_import,
        market_cost=totals.market_cost,
        supplier_markup=supplier_markup,
        energy_tax=energy_tax,
        fixed_supplier_costs=a.fixed_supplier_costs,
        feed_in_costs=a.feed_in_costs,
        grid_costs=a.grid_costs,
        tax_reduction=tax_reduction,
        opex_ex_vat=opex_ex_vat,
        vat=vat,
        opex_inc_vat=opex_inc_vat,
        average_monthly_opex=opex_inc_vat / 12.0,
    )


def _calculate(assumptions: Assumptions, hours: list[Hour]) -> ModelResult:
    baseline = _baseline(hours, assumptions)
    simple = _battery(hours, assumptions, "simple")
    lend = _battery(hours, assumptions, "lend")
    return ModelResult(
        assumptions=assumptions,
        baseline=baseline,
        simple=simple,
        lend=lend,
        baseline_costs=_costs(baseline, assumptions),
        simple_costs=_costs(simple, assumptions),
        lend_costs=_costs(lend, assumptions),
    )


def calculate(path: str | Path, overrides: dict[str, float] | None = None) -> ModelResult:
    assumptions, hours = load_model(path, overrides)
    return _calculate(assumptions, hours)


def calculate_standard_profiles(
    path: str | Path,
    overrides: dict[str, float] | None = None,
) -> ModelResult:
    assumptions, hours = load_standard_profile_model(path, overrides)
    return _calculate(assumptions, hours)
