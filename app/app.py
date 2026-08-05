from __future__ import annotations

import sys
from pathlib import Path

import streamlit as st

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from model.calculator import (
    CostResult,
    ModelResult,
    calculate_standard_profiles,
    infer_annual_usage_solar_from_grid,
)

MODEL = ROOT / "model" / "Huishoudprofiel verrekenprijs 2025.xlsx"


def apply_theme() -> None:
    st.markdown(
        """
        <style>
        :root {
            --endona-ink: #1d1d1b;
            --endona-orange: #ed7203;
            --endona-yellow: #fdc800;
        }

        .block-container {
            padding-top: 2.5rem;
        }

        h1, h2, h3 {
            color: var(--endona-orange) !important;
            font-weight: 700 !important;
        }

        section[data-testid="stSidebar"] {
            background: #ffffff;
            border-right: 4px solid var(--endona-yellow);
        }

        section[data-testid="stSidebar"] h2,
        section[data-testid="stSidebar"] h3 {
            color: var(--endona-ink) !important;
        }

        .stButton > button,
        div[data-baseweb="select"] > div,
        input {
            border-radius: 999px !important;
        }

        .stDataFrame {
            border: 1px solid rgba(29, 29, 27, 0.10);
            box-shadow: 0 10px 30px rgba(29, 29, 27, 0.06);
        }

        div[data-testid="stExpander"] {
            background: rgba(255, 255, 255, 0.78);
            border: 1px solid rgba(29, 29, 27, 0.10);
        }

        .scenario-card {
            min-height: 172px;
            height: 172px;
            border: 1px solid rgba(29, 29, 27, 0.10);
            border-radius: 0.5rem;
            padding: 1.1rem 1.25rem;
            box-shadow: 0 10px 30px rgba(29, 29, 27, 0.06);
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            background: #ffffff;
        }

        .scenario-card__label {
            color: rgba(29, 29, 27, 0.68);
            font-size: 0.9rem;
            margin-bottom: 0.35rem;
        }

        .scenario-card__value {
            color: var(--endona-ink);
            font-size: 2rem;
            font-weight: 700;
            line-height: 1.15;
        }

        .scenario-card__delta {
            color: #0f7a38;
            font-size: 0.95rem;
            margin-top: 0.45rem;
        }

        .scenario-card__monthly,
        .scenario-card__subsidy {
            color: rgba(29, 29, 27, 0.68);
            font-size: 0.9rem;
        }

        .scenario-card__subsidy {
            color: #0f7a38;
            font-weight: 600;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def euro(value: float) -> str:
    return f"€ {value:,.0f}".replace(",", ".")


def number(value: float, decimals: int = 0) -> str:
    formatted = f"{value:,.{decimals}f}"
    return formatted.replace(",", "X").replace(".", ",").replace("X", ".")


@st.cache_data(show_spinner=False)
def run_model(
    battery_capacity: float,
    grid_import: float,
    grid_export: float,
    model_revision: int,
) -> ModelResult:
    annual_usage, annual_solar = infer_annual_usage_solar_from_grid(
        MODEL,
        grid_import,
        grid_export,
    )
    return calculate_standard_profiles(
        MODEL,
        {
            "battery_capacity": battery_capacity,
            "max_daily_battery_export": battery_capacity / 3.0,
            "annual_usage_kwh": annual_usage,
            "annual_solar_kwh": annual_solar,
        },
    )


def scenario_card(
    title: str,
    costs: CostResult,
    baseline_cost: float,
    subsidy: float | None = None,
) -> None:
    saving = baseline_cost - costs.opex_inc_vat
    st.subheader(title)
    delta = "" if title == "Zonder batterij" else f"{euro(saving)} besparing"
    subsidy_text = "" if subsidy is None else f"Subsidie aanschaf batterij: {euro(subsidy)} per jaar"
    st.markdown(
        f"""
        <div class="scenario-card">
            <div>
                <div class="scenario-card__label">Totale energiekosten per jaar</div>
                <div class="scenario-card__value">{euro(costs.opex_inc_vat)}</div>
                <div class="scenario-card__delta">{delta}</div>
            </div>
            <div>
                <div class="scenario-card__monthly">{euro(costs.average_monthly_opex)} per maand</div>
                <div class="scenario-card__subsidy">{subsidy_text}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


st.set_page_config(
    page_title="Batterijvergelijking",
    page_icon="🔋",
    layout="wide",
)
apply_theme()

st.title("Batterijvergelijking")
st.caption(
    "Vergelijk de jaarlijkse energiekosten zonder batterij, met een "
    "privébatterij en met een LEND-batterij."
)

with st.sidebar:
    st.header("Uitgangspunten")
    grid_import = st.number_input(
        "Netafname (kWh/jaar)",
        min_value=0.0,
        max_value=50_000.0,
        value=6_000.0,
        step=100.0,
    )
    grid_export = st.number_input(
        "Teruglevering aan het net (kWh/jaar)",
        min_value=0.0,
        max_value=50_000.0,
        value=3_500.0,
        step=100.0,
    )
    annual_usage, annual_solar = infer_annual_usage_solar_from_grid(
        MODEL,
        grid_import,
        grid_export,
    )
    st.caption(
        "Omgerekend via de standaardprofielen: "
        f"{number(annual_usage)} kWh bruto verbruik en "
        f"{number(annual_solar)} kWh opwek/invoeding."
    )
    battery_capacity = st.number_input(
        "Grootte batterij (kWh)",
        min_value=0.0,
        max_value=100.0,
        value=9.3,
        step=0.5,
    )
    grid_operator_power = st.selectbox(
        "Vermogen beschikbaar voor Netbeheerder",
        options=(5, 10),
        index=0,
        format_func=lambda value: f"{value} kW",
        help=(
            "Voor LEND wordt de subsidie berekend als €50 per kW beschikbaar "
            "vermogen per jaar."
        ),
    )
if battery_capacity == 0:
    st.info("Bij een batterijgrootte van 0 kWh zijn beide batterijscenario’s gelijk aan elkaar.")

with st.spinner("Scenario’s berekenen…"):
    result = run_model(
        battery_capacity,
        grid_import,
        grid_export,
        MODEL.stat().st_mtime_ns,
    )

baseline_cost = result.baseline_costs.opex_inc_vat
lend_subsidy = grid_operator_power * 50.0
columns = st.columns(3)
with columns[0]:
    scenario_card(
        "Zonder batterij",
        result.baseline_costs,
        baseline_cost,
    )
with columns[1]:
    scenario_card(
        "Batterij privé",
        result.simple_costs,
        baseline_cost,
    )
with columns[2]:
    scenario_card(
        "Batterij LEND",
        result.lend_costs,
        baseline_cost,
        lend_subsidy,
    )

st.divider()
st.subheader("Kosten en besparingen")

rows = []
for name, costs, totals, subsidy in (
    ("Zonder batterij", result.baseline_costs, result.baseline, None),
    ("Batterij privé", result.simple_costs, result.simple, None),
    ("Batterij LEND", result.lend_costs, result.lend, lend_subsidy),
):
    rows.append(
        {
            "Scenario": name,
            "Kosten per jaar": euro(costs.opex_inc_vat),
            "Kosten per maand": euro(costs.average_monthly_opex),
            "Besparing per jaar": euro(baseline_cost - costs.opex_inc_vat),
            "Subsidie aanschaf batterij": "" if subsidy is None else euro(subsidy),
            "Netafname": f"{number(totals.grid_import)} kWh",
            "Teruglevering": f"{number(totals.grid_export)} kWh",
        }
    )
st.dataframe(rows, hide_index=True, width="stretch")

with st.expander("Bekijk kostenopbouw"):
    cost_rows = []
    labels = {
        "market_cost": "Inkoop en verkoop energie",
        "supplier_markup": "Leveranciersopslag",
        "energy_tax": "Energiebelasting",
        "fixed_supplier_costs": "Vaste leveringskosten",
        "feed_in_costs": "Terugleverkosten",
        "grid_costs": "Netwerkkosten",
        "tax_reduction": "Belastingvermindering",
        "vat": "Btw",
        "opex_inc_vat": "Totaal inclusief btw",
    }
    scenarios = {
        "Zonder batterij": result.baseline_costs,
        "Batterij privé": result.simple_costs,
        "Batterij LEND": result.lend_costs,
    }
    for field, label in labels.items():
        cost_rows.append(
            {"Kostenpost": label}
            | {name: euro(getattr(costs, field)) for name, costs in scenarios.items()}
        )
    st.dataframe(cost_rows, hide_index=True, width="stretch")

st.caption(f"Subsidie op aanschaf batterij bij LEND: {euro(lend_subsidy)} per jaar.")

st.caption(
    "De vergelijking gebruikt de standaardprofielen elektriciteit 2026 uit "
    "`Aannames stdprofielen` en de EPEX-prijzen uit 2025. "
    "De vaste verkoopprijzen voor batterij en zonnepanelen gelden alleen voor LEND. "
    "Vaste netwerkkosten, leverancierskosten, energiebelasting en btw zijn "
    "opgenomen; aanschaf- en financieringskosten van de batterij niet."
)
