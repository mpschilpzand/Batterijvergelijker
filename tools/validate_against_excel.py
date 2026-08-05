#!/usr/bin/env python3
"""Compare Python results with cached values from an Excel-recalculated workbook."""

from __future__ import annotations

import argparse
import math
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from model.calculator import calculate
from model.xlsx_reader import Workbook


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("workbooks", nargs="+", type=Path)
    parser.add_argument("--tolerance", type=float, default=1e-9)
    args = parser.parse_args()
    failed = False
    for path in args.workbooks:
        workbook_failed = False
        result = calculate(path)
        checks = {
            ("Model zonder batterij", "G8762"): result.baseline.grid_import,
            ("Model zonder batterij", "H8762"): result.baseline.grid_export,
            ("Model zonder batterij", "F8762"): result.baseline.market_cost,
            ("Model met batterij (simpel)", "L8762"): result.simple.grid_import,
            ("Model met batterij (simpel)", "M8762"): result.simple.grid_export,
            ("Model met batterij (simpel)", "N8762"): result.simple.market_cost,
            ("Model batterij LEND", "L8762"): result.lend.grid_import,
            ("Model batterij LEND", "M8762"): result.lend.grid_export,
            ("Model batterij LEND", "N8762"): result.lend.market_cost,
            ("Kostenoverzicht", "B18"): result.baseline_costs.opex_inc_vat,
            ("Kostenoverzicht", "C18"): result.simple_costs.opex_inc_vat,
            ("Kostenoverzicht", "D18"): result.lend_costs.opex_inc_vat,
        }
        max_difference = 0.0
        with Workbook(path) as workbook:
            for (sheet_name, address), python_value in checks.items():
                excel_value = workbook.sheet(sheet_name)[address].number()
                difference = abs(python_value - excel_value)
                max_difference = max(max_difference, difference)
                if not math.isclose(
                    python_value, excel_value, rel_tol=1e-12, abs_tol=args.tolerance
                ):
                    failed = True
                    workbook_failed = True
                    print(
                        f"FAIL {path.name} {sheet_name}!{address}: "
                        f"Python={python_value:.15g}, Excel={excel_value:.15g}, "
                        f"verschil={difference:.3g}"
                    )
        status = "FAIL" if workbook_failed else "OK"
        print(
            f"{status} {path.name}: {len(checks)} controles, "
            f"maximaal absoluut verschil {max_difference:.3g}"
        )
    raise SystemExit(1 if failed else 0)


if __name__ == "__main__":
    main()
