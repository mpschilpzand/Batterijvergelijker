from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from model.calculator import DEFAULT_RUNTIME_MODEL, export_runtime_model


def main() -> None:
    source = (
        Path(sys.argv[1])
        if len(sys.argv) > 1
        else ROOT / "model" / "Huishoudprofiel verrekenprijs 2025.xlsx"
    )
    target = Path(sys.argv[2]) if len(sys.argv) > 2 else DEFAULT_RUNTIME_MODEL
    export_runtime_model(source, target)
    print(f"Exported runtime model to {target}")


if __name__ == "__main__":
    main()
