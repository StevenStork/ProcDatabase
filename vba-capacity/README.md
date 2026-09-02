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
| **PartEditor** | `BasePartsTbl` | View all base parts; launch editor |
| FFAs | `FFAsTbl` | FFA master linked to `FactoryCode` |
| ProductLines | `ProductLinesTbl` | Product line master |
| PartDashConditions | `PartDashConditionsTbl` | Dash conditions per base part |
| PartFFAs | `PartFFAsTbl` | FFAs assigned to each base part |
| PartProductLines | `PartProductLinesTbl` | Product lines per base part |
| PartOperations | `PartOperationsTbl` | Operations (`OperSeq`) per base part |

```mermaid
erDiagram
    BasePartsTbl ||--o{ PartDashConditionsTbl : has
    BasePartsTbl ||--o{ PartFFAsTbl : uses
    BasePartsTbl ||--o{ PartProductLinesTbl : uses
    BasePartsTbl ||--o{ PartOperationsTbl : defines
    FFAsTbl ||--o{ PartFFAsTbl : assigned
    FFAsTbl }o--|| FactoriesTbl : located_in
    ProductLinesTbl ||--o{ PartProductLinesTbl : assigned
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

### Class module

| File | Class name |
|---|---|
| `clsFormControlHandler.cls` | clsFormControlHandler |

### UserForms

| UserForm name | Paste file | Purpose |
|---|---|---|
| `frmFactoryAdmin` | `frmFactoryAdmin.txt` | Factories |
| `frmEquipmentAdmin` | `frmEquipmentAdmin.txt` | Equipment |
| `frmProcessTypeAdmin` | `frmProcessTypeAdmin.txt` | Process types |
| `frmFactoryEquipmentAdmin` | `frmFactoryEquipmentAdmin.txt` | Factory-equipment links |
| `frmEquipmentProcessAdmin` | `frmEquipmentProcessAdmin.txt` | Equipment-process links |
| `frmFFAAdmin` | `frmFFAAdmin.txt` | FFA master |
| `frmProductLineAdmin` | `frmProductLineAdmin.txt` | Product line master |
| `frmPartEditor` | `frmPartEditor.txt` | **Main part editor** |
| `frmPartOperationsAdmin` | `frmPartOperationsAdmin.txt` | Operations per part |

Paste each `.txt` file into a blank UserForm with the matching `(Name)`.

### ThisWorkbook

Paste `ThisWorkbook.txt` into the ThisWorkbook code module.

## Setup steps

1. Save the workbook as **`FactoryCapacity.xlsm`**.
2. Import/paste all standard modules, `modPartIO`, and `clsFormControlHandler`.
3. Create nine blank UserForms and paste the matching form code.
4. Paste `ThisWorkbook.txt`.
5. Run **`BootstrapCapacityTables`** once.
6. Wire **Admin** and **PartEditor** buttons:

| Button caption | Macro |
|---|---|
| Manage Factories | `ShowFactoryAdmin` |
| Manage Equipment | `ShowEquipmentAdmin` |
| Manage Process Types | `ShowProcessTypeAdmin` |
| Assign Equipment to Factories | `ShowFactoryEquipmentAdmin` |
| Assign Processes to Equipment | `ShowEquipmentProcessAdmin` |
| Manage FFAs | `ShowFFAAdmin` |
| Manage Product Lines | `ShowProductLineAdmin` |
| Edit Parts | `ShowPartEditor` |
| Edit Selected Part (PartEditor sheet) | `EditSelectedPartFromSheet` |
| Part Operations | `ShowPartOperationsAdmin` |
| Rebuild Tables | `BootstrapCapacityTables` |

## Part number workflow

1. Add factories, then FFAs (`ShowFFAAdmin`) with a factory on each FFA row.
2. Add product lines if needed (`ShowProductLineAdmin`).
3. Open **`ShowPartEditor`** (or select a row on **PartEditor** and run **`EditSelectedPartFromSheet`**).
4. Create a base part, add dash conditions, FFAs, and product lines in the same form.
5. Use **Operations** to maintain `OperSeq` rows in `PartOperationsTbl`.

Factories for a part are derived from its active FFAs (shown read-only in the editor).

## Notes

- **One sheet per part is not used.** All parts live in tables; **PartEditor** is the single view/edit hub.
- Assembly numbers for external queries: `BuildActiveAssemblyNumberList()` in `modPartIO` joins active base parts with active dash conditions.
- First data row is **row 4** (headers on row 3).
- `Active` flags accept `True`/`False`, `1`/`0`, or `yes`/`no`.
