# Factory Capacity Database (VBA)

Paste-ready VBA for a new Excel workbook (`.xlsm`) that stores factory, equipment, process-type, and part-number master data in related tables for a future capacity model.

## Workbook sheets

### Capacity master data

| Sheet | Table | Purpose |
|---|---|---|
| Admin | — | Launch macros / buttons |
| Factories | `FactoriesTbl` | Factory master |
| Equipment | `EquipmentTbl` | Equipment master |
| ProcessTypes | `ProcessTypesTbl` | Process type master |
| FactoryEquipment | `FactoryEquipmentTbl` | Equipment at each factory |
| EquipmentProcesses | `EquipmentProcessTbl` | Process types per equipment |

### Part numbers (relational — no sheet per part)

| Sheet | Table | Purpose |
|---|---|---|
| **Parts** | `BasePartsTbl` | Master index of all base parts (`Name`, `FactoryCode`, `ProductLine`, …) |
| **PartEditor** | — | Load/edit workspace for one part at a time |
| PartDashConditions | `PartDashConditionsTbl` | Dash conditions per base part (`Separator`, `Active`) |
| PartOperations | `PartOperationsTbl` | Operations per base part (`OperSeq`, equipment, process type, avg toggles) |
| PartEditorCache | — | Hidden cache for sheet editor save diff (auto-created) |

### Linked source queries (connection-only)

Power Query connections — load as **connection only** when possible. Table names match connection names:

| Connection / table | Purpose |
|---|---|
| `tblRouteCard` | Route card operations |
| `tblAssyStnd` | Assembly standards (fallback process hours) |
| `tblOperComps` | Operation completions (primary process hours) |
| `tblTimeYield` | Process time / yield (`Avg 180 Day Ex`, `Avg 90 Day Ex`) |
| `tblRCCP` | Active assembly master (for future refresh pipeline) |

```mermaid
erDiagram
    FactoriesTbl ||--o{ BasePartsTbl : builds_at
    BasePartsTbl ||--o{ PartDashConditionsTbl : has
    BasePartsTbl ||--o{ PartOperationsTbl : defines
    FactoriesTbl ||--o{ FactoryEquipmentTbl : has
    EquipmentTbl ||--o{ FactoryEquipmentTbl : assigned
    EquipmentTbl ||--o{ EquipmentProcessTbl : supports
    ProcessTypesTbl ||--o{ EquipmentProcessTbl : assigned
    tblOperComps --> PartEditor : avg_hours
    tblAssyStnd --> PartEditor : avg_hours_fallback
    tblTimeYield --> PartEditor : avg_ex
```

## VBA modules to add

### Standard modules

| File | Module name |
|---|---|
| `modConstants.bas` | modConstants |
| `modTableIO.bas` | modTableIO |
| `modValidation.bas` | modValidation |
| `modBootstrap.bas` | modBootstrap |
| `modExcelOptimize.bas` | modExcelOptimize |
| `modFormUI.bas` | modFormUI |
| `modFormLauncher.bas` | modFormLauncher |
| `modPartIO.bas` | modPartIO |
| `modAverages.bas` | modAverages |
| `modPartSheetEditor.bas` | modPartSheetEditor |

### Class module

| File | Class name |
|---|---|
| `clsFormControlHandler.cls` | clsFormControlHandler |

### UserForms (optional — sheet editor is primary)

| UserForm name | Paste file | Purpose |
|---|---|---|
| `frmFactoryAdmin` | `frmFactoryAdmin.txt` | Factories |
| `frmEquipmentAdmin` | `frmEquipmentAdmin.txt` | Equipment |
| `frmProcessTypeAdmin` | `frmProcessTypeAdmin.txt` | Process types |
| `frmFactoryEquipmentAdmin` | `frmFactoryEquipmentAdmin.txt` | Factory-equipment links |
| `frmEquipmentProcessAdmin` | `frmEquipmentProcessAdmin.txt` | Equipment-process links |
| `frmPartEditor` | `frmPartEditor.txt` | Legacy popup part editor |
| `frmPartOperationsAdmin` | `frmPartOperationsAdmin.txt` | Operations per part |

Paste each `.txt` file into a blank UserForm with the matching `(Name)`.

### ThisWorkbook

Paste `ThisWorkbook.txt` into the ThisWorkbook code module.

## Setup steps

1. Save the workbook as **`FactoryCapacity.xlsm`**.
2. Import/paste all standard modules, class module, and UserForms (optional).
3. Paste `ThisWorkbook.txt`.
4. Add Power Query connections (`tblRouteCard`, `tblOperComps`, `tblRCCP`, `tblTimeYield`, `tblAssyStnd`) as connection-only.
5. Run **`BootstrapCapacityTables`** once.
6. Wire **Admin**, **Parts**, and **PartEditor** buttons:

| Button caption | Macro |
|---|---|
| Manage Factories | `ShowFactoryAdmin` |
| Manage Equipment | `ShowEquipmentAdmin` |
| Manage Process Types | `ShowProcessTypeAdmin` |
| Assign Equipment to Factories | `ShowFactoryEquipmentAdmin` |
| Assign Processes to Equipment | `ShowEquipmentProcessAdmin` |
| Open Part Editor sheet | `ShowPartEditor` |
| Load Part (PartEditor C3) | `LoadPartToEditor` |
| Save Part (PartEditor) | `SavePartFromEditor` |
| Clear Part Editor | `ClearPartEditor` |
| Open Part from Parts index | `OpenPartEditorFromPartsIndex` |
| Refresh RCCP | `RefreshRCCP` |
| Refresh Oper Completions | `RefreshOperComps` |
| Refresh Assembly Standards | `RefreshAssyStnd` |
| Refresh Route Card | `RefreshRouteCard` |
| Refresh all linked data | `RefreshAllLinkedData` |
| Part Operations (form) | `ShowPartOperationsAdmin` |
| Rebuild Tables | `BootstrapCapacityTables` |

## Part editor workflow (sheet-based)

1. Add factories and parts in **Parts** (`BasePartsTbl`) or create them via the editor on save.
2. Go to **PartEditor**, enter a base part or full assembly number in **C3**.
3. Click **Load Part** (created by bootstrap) — master fields, dash conditions, route-card rows, and operations load onto the sheet. **Avg Process Hours** and **Avg Ex** populate inline per `OperSeq` when that row’s **Show Avg Hours** / **Show Avg Ex** flags are TRUE.
4. Edit cells directly (name, factory, active, product line, notes in **C11:G16**, dash rows from column **I**, route card on the left of operations, operation rows from column **F**). Pick **Equipment** (factory equipment) and **Process Type** (processes for that equipment) from the dropdowns. Status messages appear in **C7**.
5. Click **Save Part** — changes write back to `BasePartsTbl`, `PartDashConditionsTbl`, and `PartOperationsTbl`. A hidden **PartEditorCache** sheet tracks the last loaded state for add/update/delete diffing.

Or select a row on **Parts** and run **`OpenPartEditorFromPartsIndex`**.

`BootstrapCapacityTables` formats PartEditor and creates the **Load Part**, **Save Part**, and **Clear** buttons on the sheet. It also drops the legacy **StatusDate** column from `BasePartsTbl` if present.

**Route Card** (columns B–D, from `tblRouteCard`): dash condition parsed from `ASSEMBLY NO`, plus `OPER SEQ` and `OPER CODE` for the loaded base part. Load `tblRouteCard` to a sheet as a ListObject.

### Average calculations (`modAverages`)

| Column | Source | Logic |
|---|---|---|
| **Avg Process Hours** | `tblOperComps` → `tblAssyStnd` | Average non-zero `LABOR HPS (HOURS)` for base part + `OPER SEQ`; fallback to average non-zero `RUN TIME (HOURS)` |
| **Avg Ex** | `tblTimeYield` | Average non-zero `Avg 180 Day Ex`; fallback to `Avg 90 Day Ex` |

Per-operation **Show Avg Hours** / **Show Avg Ex** columns control whether those values are filled. Matching uses `ASSEMBLY NO` (full dashed assembly numbers) and extracts the base part before `-`.

## Linked query refresh

### RCCP (`tblRCCP`)

The **#"Filtered FFAs"** step must read factory codes from `FactoriesTbl` instead of hard-coded FFAs. Paste the replacement step from [`PowerQuery/pqRCCP-FilteredFFAs.txt`](../PowerQuery/pqRCCP-FilteredFFAs.txt).

Run **`RefreshRCCP`**. The query filters `[FFA]` to active `FactoryCode` values. No VBA rewrites that filter on each refresh.

After the query refresh, VBA syncs **`PartDashConditionsTbl`**:

- Assemblies in `tblRCCP` that are missing from the dash table are **added** as Active (with `Separator` and leading zeros preserved in `DashCondition`).
- Dash rows present in the table but **not** in `tblRCCP` are marked **Inactive** (not deleted).

Parsing supports `BASE-DASH` and letter separators such as `BASEA01` (`Separator` = `A`, `DashCondition` = `01`). Uses `Base PN: Text` from RCCP when present.

**Important:** load `tblRCCP` to a sheet as a ListObject (hidden is fine). Connection-only alone cannot supply assembly numbers to the next step.

### Oper Completions (`tblOperComps`)

No full M rewrite required. Keep an `@ffa = '...'` parameter in the Source SQL. **`RefreshOperComps`** uses the same active `FactoryCode` list as RCCP, rewrites `@ffa` **only when that list changed**, then refreshes. See [`PowerQuery/pqOperComps-FFA.txt`](../PowerQuery/pqOperComps-FFA.txt).

`tblOperComps` may stay connection-only. Unchanged `@ffa` avoids Power Query permission prompts on every refresh.

### Assembly Standards (`tblAssyStnd`)

1. **`@ffa`** — same VBA update-on-change pattern as OperComps. See [`PowerQuery/pqAssyStnd-FFA.txt`](../PowerQuery/pqAssyStnd-FFA.txt).
2. **`#"Filter Assemblies"`** — after Source, keep only rows whose `ASSEMBLY NO` appears in `tblRCCP`. Paste from [`PowerQuery/pqAssyStnd-FilterAssemblies.txt`](../PowerQuery/pqAssyStnd-FilterAssemblies.txt). Requires `tblRCCP` as a workbook ListObject.

Run **`RefreshAssyStnd`** (or use combined refresh after RCCP).

### Route Card (`tblRouteCard`)

Same pattern as Assembly Standards:

1. **`@ffa`** — update-on-change from FactoriesTbl. See [`PowerQuery/pqRouteCard-FFA.txt`](../PowerQuery/pqRouteCard-FFA.txt).
2. **`#"Filter Assemblies"`** — after Source, filter to `tblRCCP` assemblies. Paste from [`PowerQuery/pqRouteCard-FilterAssemblies.txt`](../PowerQuery/pqRouteCard-FilterAssemblies.txt).

Run **`RefreshRouteCard`**.

### Combined

**`RefreshAllLinkedData`** currently runs RCCP → OperComps → AssyStnd → RouteCard. More queries will be added here later.

## Notes

- **One sheet per part is not used.** All parts live in tables; **PartEditor** is the edit workspace.
- Linked tables must exist as ListObjects on a sheet (visible or hidden) for averages to calculate. Connection-only queries need a refresh target sheet until parameterized refresh is implemented.
- Re-run **`BootstrapCapacityTables`** to migrate `BasePartsTbl` from PartEditor to Parts (legacy table renamed automatically).
- First data row is **row 4** on index sheets (headers on row 3).
