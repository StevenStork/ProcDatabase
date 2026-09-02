# Factory Capacity Database (VBA)

Paste-ready VBA for a new Excel workbook (`.xlsm`) that stores factory, equipment, and process-type master data plus many-to-many assignment tables for a future capacity model.

## Workbook sheets

| Sheet | Table | Purpose |
|---|---|---|
| Admin | — | Launch macros / buttons |
| Factories | `FactoriesTbl` | Factory master |
| Equipment | `EquipmentTbl` | Equipment master |
| ProcessTypes | `ProcessTypesTbl` | Process type master |
| FactoryEquipment | `FactoryEquipmentTbl` | Equipment at each factory |
| EquipmentProcesses | `EquipmentProcessTbl` | Process types per equipment |

## VBA modules to add

### Standard modules (Import or paste into Module objects)

| File | Module name |
|---|---|
| `modConstants.bas` | modConstants |
| `modTableIO.bas` | modTableIO |
| `modValidation.bas` | modValidation |
| `modBootstrap.bas` | modBootstrap |
| `modExcelOptimize.bas` | modExcelOptimize |
| `modFormUI.bas` | modFormUI |
| `modFormLauncher.bas` | modFormLauncher |

### Class modules

| File | Class name |
|---|---|
| `clsCommandButtonHandler.cls` | clsCommandButtonHandler |
| `clsListBoxHandler.cls` | clsListBoxHandler |
| `clsComboBoxHandler.cls` | clsComboBoxHandler |

Required for programmatic UserForm controls. Each handler wires only the events that control type supports (avoids run-time error 459).

### UserForms (manual shell + paste code)

For each form below:

1. VBA Editor → Insert → UserForm
2. Set the `(Name)` property in the Properties window
3. Paste the matching `.txt` file into the UserForm code window

| UserForm name | Paste file |
|---|---|
| `frmFactoryAdmin` | `frmFactoryAdmin.txt` |
| `frmEquipmentAdmin` | `frmEquipmentAdmin.txt` |
| `frmProcessTypeAdmin` | `frmProcessTypeAdmin.txt` |
| `frmFactoryEquipmentAdmin` | `frmFactoryEquipmentAdmin.txt` |
| `frmEquipmentProcessAdmin` | `frmEquipmentProcessAdmin.txt` |

### ThisWorkbook

Paste `ThisWorkbook.txt` into the ThisWorkbook code module.

## Setup steps

1. Create a new workbook and save as **`FactoryCapacity.xlsm`**.
2. Import or paste all standard modules and the three class modules.
3. Create five blank UserForms, rename each, paste the matching form code.
4. Paste `ThisWorkbook.txt` into ThisWorkbook.
5. Run **`BootstrapCapacityTables`** (Alt+F8) once to create sheets and tables.
6. On the **Admin** sheet, add Form Control buttons and assign macros:

| Button caption | Macro |
|---|---|
| Manage Factories | `ShowFactoryAdmin` |
| Manage Equipment | `ShowEquipmentAdmin` |
| Manage Process Types | `ShowProcessTypeAdmin` |
| Assign Equipment to Factories | `ShowFactoryEquipmentAdmin` |
| Assign Processes to Equipment | `ShowEquipmentProcessAdmin` |
| Rebuild Tables (dev) | `BootstrapCapacityTables` |

## Smoke test

1. Add two factories, three equipment records, and four process types via the master forms.
2. Assign equipment to both factories using **Assign Equipment to Factories**.
3. Assign multiple process types to one equipment record using **Assign Processes to Equipment**.
4. Confirm rows appear in the `FactoryEquipment` and `EquipmentProcesses` sheets.

## Notes

- All forms build their controls in `UserForm_Initialize`; no `.frm`/`.frx` import is required.
- Master-table deletes are blocked while junction rows still reference the record.
- `Active` flags accept `True`/`False`, `1`/`0`, or `yes`/`no` when reading existing data.
