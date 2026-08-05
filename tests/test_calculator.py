from __future__ import annotations

import math
import unittest
from pathlib import Path

from model.calculator import (
    calculate,
    calculate_standard_profiles,
    infer_annual_usage_solar_from_grid,
)

MODEL = Path(__file__).parents[1] / "model" / "Huishoudprofiel verrekenprijs 2025.xlsx"


class CalculatorRegressionTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.result = calculate(MODEL)

    def assertClose(self, actual: float, expected: float) -> None:
        self.assertTrue(
            math.isclose(actual, expected, rel_tol=1e-12, abs_tol=1e-9),
            f"{actual!r} != {expected!r}",
        )

    def test_default_excel_totals(self) -> None:
        expected = {
            "baseline": (2484.5150765833932, 2984.515076583366, 157.34267970693105),
            "simple": (1412.6722499823229, 1912.6722499823052, 15.870222482427282),
            "lend": (1390.3262976183976, 1890.3262976183794, 40.792037257216926),
        }
        for name, values in expected.items():
            scenario = getattr(self.result, name)
            for actual, wanted in zip(
                (scenario.grid_import, scenario.grid_export, scenario.market_cost), values
            ):
                self.assertClose(actual, wanted)

    def test_default_excel_costs(self) -> None:
        self.assertClose(self.result.baseline_costs.opex_inc_vat, 527.68769317233034)
        self.assertClose(self.result.simple_costs.opex_inc_vat, 192.39252048418047)
        self.assertClose(self.result.lend_costs.opex_inc_vat, 219.12645161899727)

    def test_standard_profile_costs(self) -> None:
        result = calculate_standard_profiles(MODEL)
        self.assertClose(result.baseline_costs.opex_inc_vat, 541.4549443426316)
        self.assertClose(result.simple_costs.opex_inc_vat, 226.29473673553161)
        self.assertClose(result.lend_costs.opex_inc_vat, 230.10385318322193)

    def test_standard_profile_grid_inputs_infer_usage_and_solar(self) -> None:
        annual_usage, annual_solar = infer_annual_usage_solar_from_grid(
            MODEL,
            2438.665190356448,
            2938.66519035656,
        )
        self.assertClose(annual_usage, 4000.0)
        self.assertClose(annual_solar, 4500.0)

    def test_energy_balances(self) -> None:
        for scenario in (self.result.baseline, self.result.simple, self.result.lend):
            self.assertClose(
                scenario.usage + scenario.grid_export,
                scenario.solar + scenario.grid_import,
            )
        for scenario in (self.result.simple, self.result.lend):
            self.assertClose(
                scenario.battery_charge_solar + scenario.battery_charge_grid,
                scenario.battery_discharge_home
                + scenario.battery_export
                + scenario.final_charge,
            )

    def test_overrides_do_not_mutate_source(self) -> None:
        changed = calculate(
            MODEL,
            {"annual_usage_kwh": 5200.0, "annual_solar_kwh": 3000.0},
        )
        self.assertClose(changed.baseline.usage, 5200.0)
        self.assertClose(changed.baseline.solar, 3000.0)
        self.assertClose(self.result.baseline.usage, 4000.0)

    def test_fixed_solar_export_price(self) -> None:
        low = calculate(MODEL, {"solar_export_price": 0.05})
        high = calculate(MODEL, {"solar_export_price": 0.06})
        self.assertClose(low.baseline.market_cost, high.baseline.market_cost)
        self.assertClose(low.simple.market_cost, high.simple.market_cost)
        self.assertLess(high.lend.market_cost, low.lend.market_cost)
        self.assertLess(high.lend_costs.opex_inc_vat, low.lend_costs.opex_inc_vat)

    def test_fixed_battery_export_price_only_affects_lend(self) -> None:
        low = calculate(MODEL, {"lend_export_price": 0.10})
        high = calculate(MODEL, {"lend_export_price": 0.20})
        self.assertClose(low.baseline.market_cost, high.baseline.market_cost)
        self.assertClose(low.simple.market_cost, high.simple.market_cost)
        self.assertNotEqual(low.lend.market_cost, high.lend.market_cost)

    def test_external_battery_use_controls_both_battery_scenarios(self) -> None:
        disabled = calculate(MODEL, {"max_daily_battery_export": 0.0})
        enabled = calculate(MODEL, {"max_daily_battery_export": 2.5})
        self.assertClose(disabled.simple.battery_export, 0.0)
        self.assertClose(disabled.lend.battery_export, 0.0)
        self.assertGreater(enabled.simple.battery_export, 0.0)
        self.assertGreater(enabled.lend.battery_export, 0.0)


if __name__ == "__main__":
    unittest.main()
