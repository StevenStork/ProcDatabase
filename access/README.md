# Access ProcDatabase

Access port of the Excel ProcDatabase concept. Source data comes from **linked tables** in your `.accdb`; SQL definitions live in same-folder **`dataQueries.xlsm`**.

## Prerequisites

Link these tables in Access before importing VBA:

| Access table   | Source                    |
|----------------|---------------------------|
| `tblRouteCard` | Route_Card                |
| `tblAssyStnd`  | Assembly_Standard         |
| `tblOperComps` | Oper_Completions          |
| `tblRCCP`      | RCCP (defines active assemblies) |

Optional: `tblProcTmYld` for 180/90-day execution averages.

See [`docs/DATA_SOURCES.md`](docs/DATA_SOURCES.md) for column lists and refresh behavior.

## First-time setup (Windows + Access)

1. Open your `.accdb` with the four required linked tables.
2. Alt+F11 → **Import File** → import every module in [`vba/`](vba/).
   - If re-importing after an update, **delete the old ProcDatabase modules first** (same names under *Modules*). Duplicate imports cause **Ambiguous name** compile errors.
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

### If Bootstrap fails

1. Run `DiagnoseSchema` in the Immediate window — it lists which linked/local tables Access sees.
2. Re-import the latest `modSchema.bas`, `modUtils.bas`, and `modApp.bas`, then **Debug → Compile**.
3. Run `BootstrapProcDatabase` again — errors now include the schema sub-step (for example `EnsurePartTables`).
4. If a previous attempt created broken local tables (`tblMeta`, `tblPart`, …), delete those **local** tables in the navigation pane and run bootstrap again. Do **not** delete your linked `tblRouteCard`, `tblAssyStnd`, `tblOperComps`, or `tblRCCP`.
5. If form creation fails after schema succeeds, run `BootstrapSchemaOnly`, then `BuildUi`.

## Workflow

1. **Refresh Linked Data** — refresh linked tables → rebuild catalog from `tblAssyStnd` → apply **tblRCCP** (Active parts/dashes + product-line UseFlags) → rebuild `tblActiveAssemblyFilter`.
2. On Home, find parts with **Search**, **Active only**, **FFA**, or **Jump to part** (no need to scroll the full list). Set **Date**, **Home FFA**, **Notes** as needed. Forms open maximized and size to the Access workspace (`UsableWidth` / `UsableHeight`); after UI changes, re-import `modUtils` / `modUi` / `modHome` / `modApp` and run `BuildUi`.
3. **Open Part** — confirm dashes/product lines; **Seed Ops** from route card (fills Imported Hours / Imported Ex only).
4. Enter **Process Hours**, **Avg Ex**, and **Batch Size** manually. **Avg HPU** = (hours × ex) / batch; check **Use Import Hrs** / **Use Import Ex** to substitute Imported values in that formula.
5. On each op, choose **Made In FFA** (FFA list), then **Equipment** (equipment linked to that FFA). **Equipment Type** fills from the selected equipment.
6. **Export** — FFA, product line, or all to `.xlsx` beside the database.

## Plan features implemented

| Plan item | Implementation |
|-----------|----------------|
| Local Parts / Dashes / ProductLines / Operations | `tblPart`, `tblPartDash`, `tblPartProductLine`, `tblOperation` |
| Linked standards / route / completions / RCCP | Your `tblRouteCard`, `tblAssyStnd`, `tblOperComps`, `tblRCCP` |
| Active assembly filter | From `tblRCCP` → `tblActiveAssemblyFilter` + filtered queries |
| Home dashboard + RAG | `frmHome` / `sfrmHomeList`; Days >90 red, >30 yellow |
| Part form + ops subform | `frmPart`, `sfrmOperation` |
| References | `frmFFA`, `frmProductLine` (+ PL Code), `frmEquipment` / `frmEquipmentFFA` / `frmEquipmentEntry` (manual; legacy `sfrmEquipmentFFA` is removed) |
| Averages | `modAverages` (labor + yield fallbacks) |
| Exports | `modExport` — includes Equipment + Equipment Type |

## Module map

| Module | Role |
|--------|------|
| `modConstants` | Table/query/form names |
| `modHome` | Home search/filter/jump controller |
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
