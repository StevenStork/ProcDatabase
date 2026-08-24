#!/usr/bin/env python3
"""Sample rows mirroring linked-table columns for offline tests."""

from __future__ import annotations

import json
from pathlib import Path

SAMPLE = {
    "tblAssyStnd": [
        {
            "ASSEMBLY NO": "ABC123-001",
            "OPER SEQ": 10,
            "OPER CODE": "WELD",
            "ASSEMBLY DESCRIPTION": "Widget assembly",
            "OPER DESCRIPTION": "Weld station",
            "RUN TIME (HOURS)": 1.5,
            "FFA": "F1",
            "ORG CODE": "ORG1",
        },
        {
            "ASSEMBLY NO": "ABC123-002",
            "OPER SEQ": 20,
            "OPER CODE": "PAINT",
            "ASSEMBLY DESCRIPTION": "Widget assembly",
            "OPER DESCRIPTION": "Paint booth",
            "RUN TIME (HOURS)": 0.75,
            "FFA": "F2",
            "ORG CODE": "ORG1",
        },
    ],
    "tblRouteCard": [
        {
            "ASSEMBLY NO": "ABC123-001",
            "OPER SEQ": 10,
            "OPER CODE": "WELD",
            "ASSEMBLY DESCRIPTION": "Widget assembly",
            "OPER DESCRIPTION": "Weld station",
            "FFA": "F1",
            "ORG CODE": "ORG1",
        }
    ],
    "tblOperComps": [
        {
            "ASSEMBLY NO": "ABC123-001",
            "S/N": "SN001",
            "OPER SEQ": 10,
            "OPER CODE": "WELD",
            "OPER DESCRIPTION": "Weld station",
            "LABOR HPS (HOURS)": 2.0,
            "QTY": 1,
            "PROJECT": "P1",
            "PROGRAM FAMILY": "PF1",
            "FFA": "F1",
            "ORG CODE": "ORG1",
        }
    ],
    "tblProcTmYld": [
        {
            "Assembly No": "ABC123-001",
            "OPER SEQ": 10,
            "Avg 180 Day Ex": 8.0,
            "Avg 90 Day Ex": 5.0,
        }
    ],
    "parts": [
        {"BasePart": "ABC123", "Active": True, "Dash": "001", "DashActive": True},
        {"BasePart": "ABC123", "Active": True, "Dash": "002", "DashActive": False},
    ],
}


def active_assembly_numbers(parts: list[dict]) -> list[str]:
    out: list[str] = []
    for row in parts:
        if row.get("Active") and row.get("DashActive"):
            out.append(f"{row['BasePart']}-{row['Dash']}")
    return sorted(out)


def main() -> None:
    root = Path(__file__).resolve().parent
    payload = dict(SAMPLE)
    payload["activeAssemblyFilter"] = active_assembly_numbers(SAMPLE["parts"])
    out = root / "sample_data.json"
    out.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    print(f"Wrote {out}")


if __name__ == "__main__":
    main()
