"""Domain-rule tests mirroring Access VBA averages / assembly split / HPU."""

from __future__ import annotations


def split_assembly_no(assembly_no: str) -> tuple[str, str]:
    assembly_no = assembly_no.strip()
    dash_pos = assembly_no.find("-")
    if dash_pos > 0:
        return assembly_no[:dash_pos].strip(), assembly_no[dash_pos + 1 :].strip()
    return assembly_no, ""


def get_base_part(assembly_no: str) -> str:
    return split_assembly_no(assembly_no)[0]


def average_nonzero(values: list[float | None]) -> float | None:
    usable = [v for v in values if v is not None and v != 0]
    if not usable:
        return None
    return sum(usable) / len(usable)


def avg_labor_hours(
    base_part: str,
    op_seq: int,
    completions: list[dict],
    standards: list[dict],
) -> float | None:
    def match_hours(rows: list[dict], hours_key: str) -> float | None:
        matched = []
        for row in rows:
            if get_base_part(str(row["ASSEMBLY NO"])) != base_part:
                continue
            if int(row["OPER SEQ"]) != op_seq:
                continue
            matched.append(row.get(hours_key))
        return average_nonzero(matched)

    completions_avg = match_hours(completions, "LABOR HPS (HOURS)")
    if completions_avg is not None:
        return completions_avg
    return match_hours(standards, "RUN TIME (HOURS)")


def avg_proc_tm_yld(base_part: str, op_seq: int, rows: list[dict]) -> float | None:
    def match(col: str) -> float | None:
        matched = []
        for row in rows:
            if get_base_part(str(row["Assembly No"])) != base_part:
                continue
            if int(row["OPER SEQ"]) != op_seq:
                continue
            matched.append(row.get(col))
        return average_nonzero(matched)

    avg180 = match("Avg 180 Day Ex")
    if avg180 is not None:
        return avg180
    return match("Avg 90 Day Ex")


def hours_for_hpu(manual: float | None, imported: float | None, use_import: bool) -> float | None:
    if use_import and imported is not None:
        return imported
    return manual


def avg_hpu(hours: float | None, executions: float | None, batch: float | None) -> float | None:
    if hours is None or executions is None or not batch:
        return None
    return (hours * executions) / batch


def build_catalog(standards: list[dict]) -> dict[str, set[str]]:
    catalog: dict[str, set[str]] = {}
    for row in standards:
        base, dash = split_assembly_no(str(row["ASSEMBLY NO"]))
        catalog.setdefault(base, set())
        if dash:
            catalog[base].add(dash)
    return catalog


def test_split_assembly_no():
    assert split_assembly_no("ABC123-001") == ("ABC123", "001")
    assert split_assembly_no("NODASH") == ("NODASH", "")
    assert split_assembly_no("A-B-C") == ("A", "B-C")


def test_avg_labor_prefers_completions():
    completions = [
        {"ASSEMBLY NO": "P1-01", "OPER SEQ": 10, "LABOR HPS (HOURS)": 2.0},
        {"ASSEMBLY NO": "P1-02", "OPER SEQ": 10, "LABOR HPS (HOURS)": 4.0},
        {"ASSEMBLY NO": "P1-01", "OPER SEQ": 10, "LABOR HPS (HOURS)": 0.0},
    ]
    standards = [
        {"ASSEMBLY NO": "P1-01", "OPER SEQ": 10, "RUN TIME (HOURS)": 9.0},
    ]
    assert avg_labor_hours("P1", 10, completions, standards) == 3.0


def test_avg_labor_falls_back_to_standards():
    completions = [
        {"ASSEMBLY NO": "P1-01", "OPER SEQ": 10, "LABOR HPS (HOURS)": 0.0},
    ]
    standards = [
        {"ASSEMBLY NO": "P1-01", "OPER SEQ": 10, "RUN TIME (HOURS)": 1.5},
        {"ASSEMBLY NO": "P1-02", "OPER SEQ": 10, "RUN TIME (HOURS)": 2.5},
    ]
    assert avg_labor_hours("P1", 10, completions, standards) == 2.0


def test_avg_proc_prefers_180_then_90():
    rows = [
        {"Assembly No": "P1-01", "OPER SEQ": 20, "Avg 180 Day Ex": 0.0, "Avg 90 Day Ex": 5.0},
        {"Assembly No": "P1-02", "OPER SEQ": 20, "Avg 180 Day Ex": None, "Avg 90 Day Ex": 7.0},
    ]
    assert avg_proc_tm_yld("P1", 20, rows) == 6.0

    rows180 = [
        {"Assembly No": "P1-01", "OPER SEQ": 20, "Avg 180 Day Ex": 8.0, "Avg 90 Day Ex": 1.0},
    ]
    assert avg_proc_tm_yld("P1", 20, rows180) == 8.0


def test_hpu_and_import_overrides():
    # Manual process/ex with no overrides.
    hours = hours_for_hpu(1.0, 9.0, False)
    ex = hours_for_hpu(3.0, 8.0, False)
    assert hours == 1.0
    assert ex == 3.0
    assert avg_hpu(hours, ex, 2.0) == 1.5

    # Use Import Hrs / Use Import Ex swap in Imported values.
    hours = hours_for_hpu(1.0, 9.0, True)
    ex = hours_for_hpu(3.0, 8.0, True)
    assert hours == 9.0
    assert ex == 8.0
    assert avg_hpu(hours, ex, 2.0) == 36.0


def test_catalog_from_assy_stnd():
    standards = [
        {"ASSEMBLY NO": "ABC-001", "OPER SEQ": 10, "FFA": "F1"},
        {"ASSEMBLY NO": "ABC-002", "OPER SEQ": 20, "FFA": "F1"},
        {"ASSEMBLY NO": "XYZ-010", "OPER SEQ": 10, "FFA": "F2"},
    ]
    catalog = build_catalog(standards)
    assert catalog["ABC"] == {"001", "002"}
    assert catalog["XYZ"] == {"010"}


def active_assembly_numbers(parts: list[dict]) -> list[str]:
    out: list[str] = []
    for row in parts:
        if row.get("Active") and row.get("DashActive"):
            out.append(f"{row['BasePart']}-{row['Dash']}")
    return sorted(out)


def active_assembly_numbers_from_rccp(rccp_rows: list[dict]) -> list[str]:
    """Anything in tblRCCP is active; use ASSEMBLY NO (distinct)."""
    out: list[str] = []
    seen: set[str] = set()
    for row in rccp_rows:
        assembly_no = str(row.get("ASSEMBLY NO", "")).strip()
        if assembly_no and assembly_no not in seen:
            seen.add(assembly_no)
            out.append(assembly_no)
    return sorted(out)


def apply_rccp_dash_active(assembly_no: str, base_pn: str | None = None) -> tuple[str, str]:
    """Return (base_part, dash) for an RCCP row."""
    base, dash = split_assembly_no(assembly_no)
    if base_pn and str(base_pn).strip():
        base = str(base_pn).strip()
    return base, dash


def rag_band(days: int | None) -> str | None:
    if days is None:
        return None
    if days > 90:
        return "red"
    if days > 30:
        return "yellow"
    return "green"


def test_active_assembly_filter():
    parts = [
        {"BasePart": "ABC123", "Active": True, "Dash": "001", "DashActive": True},
        {"BasePart": "ABC123", "Active": True, "Dash": "002", "DashActive": False},
        {"BasePart": "XYZ", "Active": False, "Dash": "010", "DashActive": True},
    ]
    assert active_assembly_numbers(parts) == ["ABC123-001"]


def test_rccp_active_assemblies():
    rccp = [
        {"ASSEMBLY NO": "ABC123-001", "PRODUCT LINE: Text": "COM"},
        {"ASSEMBLY NO": "ABC123-001", "PRODUCT LINE: Text": "COM"},  # duplicate
        {"ASSEMBLY NO": "XYZ-010", "PRODUCT LINE: Text": "MIL"},
    ]
    assert active_assembly_numbers_from_rccp(rccp) == ["ABC123-001", "XYZ-010"]


def test_rccp_dash_from_assembly_no():
    assert apply_rccp_dash_active("ABC123-001") == ("ABC123", "001")
    assert apply_rccp_dash_active("ABC123-001", "ABC123") == ("ABC123", "001")
    assert apply_rccp_dash_active("ABC123-002", "OVERRIDE") == ("OVERRIDE", "002")


def test_rag_thresholds():
    assert rag_band(None) is None
    assert rag_band(10) == "green"
    assert rag_band(31) == "yellow"
    assert rag_band(91) == "red"


def test_linked_table_names():
    # Contract with the user's Access file.
    assert {
        "tblRouteCard",
        "tblAssyStnd",
        "tblOperComps",
        "tblRCCP",
    } == {"tblRouteCard", "tblAssyStnd", "tblOperComps", "tblRCCP"}
