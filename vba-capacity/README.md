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
| **Parts** | `BasePartsTbl` | Master index of all base parts |
| **PartEditor** | — | Load/edit workspace for one part at a time |
| PartDashConditions | `PartDashConditionsTbl` | Dash conditions per base part |
| PartOperations | `PartOperationsTbl` | Operations (`OperSeq`) per base part |
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
| Refresh all linked data | `RefreshAllLinkedData` |
| Part Operations (form) | `ShowPartOperationsAdmin` |
| Rebuild Tables | `BootstrapCapacityTables` |

## Part editor workflow (sheet-based)

1. Add factories and parts in **Parts** (`BasePartsTbl`) or create them via the editor on save.
2. Go to **PartEditor**, enter a base part or full assembly number in **C3**.
3. Run **`LoadPartToEditor`** — master fields, dash conditions, and operations load onto the sheet. **Avg Process Hours** and **Avg Ex** populate inline per `OperSeq` from linked tables (when loaded).
4. Edit cells directly (factory, active, status date, notes, dash rows, operation rows).
5. Run **`SavePartFromEditor`** — changes write back to `BasePartsTbl`, `PartDashConditionsTbl`, and `PartOperationsTbl`. A hidden **PartEditorCache** sheet tracks the last loaded state for add/update/delete diffing.

Or select a row on **Parts** and run **`OpenPartEditorFromPartsIndex`**.

### Average calculations (`modAverages`)

| Column | Source | Logic |
|---|---|---|
| **Avg Process Hours** | `tblOperComps` → `tblAssyStnd` | Average non-zero `LABOR HPS (HOURS)` for base part + `OPER SEQ`; fallback to average non-zero `RUN TIME (HOURS)` |
| **Avg Ex** | `tblTimeYield` | Average non-zero `Avg 180 Day Ex`; fallback to `Avg 90 Day Ex` |

Matching uses `ASSEMBLY NO` (full dashed assembly numbers) and extracts the base part before `-`.

## RCCP refresh (`tblRCCP`)

The **#"Filtered FFAs"** step in the `tblRCCP` Power Query must read factory codes from `FactoriesTbl` instead of hard-coded FFAs. Paste the replacement step from [`PowerQuery/pqRCCP-FilteredFFAs.txt`](../PowerQuery/pqRCCP-FilteredFFAs.txt) into the query editor (Advanced Editor or replace that step only).

Then run **`RefreshRCCP`** from Admin. The query filters `[FFA]` to active `FactoryCode` values in `FactoriesTbl`. No VBA rewrites the M code on each refresh — update factories in the table and refresh.

## Notes

- **One sheet per part is not used.** All parts live in tables; **PartEditor** is the edit workspace.
- Linked tables must exist as ListObjects on a sheet (visible or hidden) for averages to calculate. Connection-only queries need a refresh target sheet until parameterized refresh is implemented.
- Re-run **`BootstrapCapacityTables`** to migrate `BasePartsTbl` from PartEditor to Parts (legacy table renamed automatically).
- First data row is **row 4** on index sheets (headers on row 3).
