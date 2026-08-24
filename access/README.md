# Access ProcDatabase

Access port of the Excel ProcDatabase concept. Source data comes from **linked tables** in your `.accdb`; SQL definitions live in same-folder **`dataQueries.xlsm`**.

## Prerequisites

Link these tables in Access before importing VBA:

| Access table   | Source                    |
|----------------|---------------------------|
| `tblRouteCard` | Route_Card                |
| `tblAssyStnd`  | Assembly_Standard         |
| `tblOperComps` | Oper_Completions          |

Optional: `tblProcTmYld` for 180/90-day execution averages.

See [`docs/DATA_SOURCES.md`](docs/DATA_SOURCES.md) for column lists and refresh behavior.

## First-time setup (Windows + Access)

1. Open your `.accdb` with the three linked tables.
2. Alt+F11 → **Import File** → import every module in [`vba/`](vba/).
3. Immediate window (`Ctrl+G`):

```vb
BootstrapProcDatabase
```

Or, if schema already exists:

```vb
BuildUi
```

4. Close/reopen (startup form `frmHome`) or run `StartProcDatabase`.
5. **Refresh Linked Data** on Home (or `RefreshAll`).

## Workflow

1. **Refresh Linked Data** — `RefreshLink` on linked tables → rebuild part/dash catalog from `tblAssyStnd` → seed FFAs/equipment → rebuild `tblActiveAssemblyFilter`.
2. Mark parts **Active** on Home; set **Date**, **Home FFA**, **Highlight**.
3. **Open Part** — activate dash conditions and product lines; **Seed Ops** from route card.
4. Edit ops (batch, export overrides); **Process Hours** / **Avg Ex** / **Avg HPU** follow Excel W/X/Y rules.
5. **Export** — FFA, product line, or all to `.xlsx` beside the database.

## Plan features implemented

| Plan item | Implementation |
|-----------|----------------|
| Local Parts / Dashes / ProductLines / Operations | `tblPart`, `tblPartDash`, `tblPartProductLine`, `tblOperation` |
| Linked standards / route / completions | Your `tblRouteCard`, `tblAssyStnd`, `tblOperComps` |
| Active assembly filter | `tblActiveAssemblyFilter` + `qryRouteCardActive` etc. |
| Home dashboard + RAG | `frmHome` bound to `tblPart`; Days >90 red, >30 yellow |
| Part form + ops subform | `frmPart`, `sfrmOperation` |
| References | `frmFFA`, `frmProductLine`, `frmEquipment` + seed from sources |
| Averages | `modAverages` (labor + yield fallbacks) |
| Exports | `modExport` — 10 columns matching Excel export PR |

## Module map

| Module | Role |
|--------|------|
| `modConstants` | Table/query/form names |
| `modLinkedData` | Refresh linked tables |
| `modActiveFilter` | Active assembly list + filtered queries |
| `modCatalog` | Build parts/dashes from standards |
| `modReferences` | Seed FFAs / equipment |
| `modAverages` | Labor / process-time averages |
| `modOperations` | Seed ops from route card |
| `modExport` | Excel exports |
| `modSchema` / `modUi` / `modApp` | Schema, forms, entry points |
| `modUtils` | Shared helpers |

## Offline tests

```bash
python3 -m pytest access/tests/test_domain_rules.py -v
python3 access/samples/generate_sample_data.py
```

## Note

This repo ships VBA as text modules to import into your Access file; it does not include a compiled `.accdb`.
