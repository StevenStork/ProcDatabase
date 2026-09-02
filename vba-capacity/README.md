# Factory Capacity Database (VBA)

Paste-ready VBA for a new Excel workbook (`.xlsm`) that stores factory, equipment, and process-type master data plus many-to-many assignment tables for a future capacity model.

## Workbook sheets

| Sheet | Table | Purpose |
|---|---|---|
| Admin | — | Launch macros / buttons |
| Factories | `FactoriesTbl` | Factory master |
| Equipment | `EquipmentTbl` | Equipment master |
| ProcessTypes | `ProcessTypesTbl` | Process type master |
| FactoryEquipment | `FactoryEquipmentTbl` | Equipment at each factory (`FactoryCode`, `EquipmentCode`, `Notes`) |
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

### Class module (required)

| File | Class name |
|---|---|
| `clsFormControlHandler.cls` | clsFormControlHandler |

Required for programmatic UserForm button/list/combo events (avoids run-time error 459).

**How to add it (pick one method):**

- **Import (recommended):** VBA Editor → File → Import File → select `clsFormControlHandler.cls`
- **Paste:** Insert → Class Module, set `(Name)` = `clsFormControlHandler` in the Properties window, then paste everything from `Option Explicit` downward (do not paste the `VERSION` / `Attribute` header lines at the top of the `.cls` file)

If `(Name)` is anything else (e.g. `Class1`), you will get **Compile error: Type not defined** on `clsFormControlHandler`.

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
2. Import or paste all standard modules and **`clsFormControlHandler`** (see class module instructions above).
3. Create five blank UserForms, rename each, paste the matching form code.
4. Paste `ThisWorkbook.txt` into ThisWorkbook.
5. Run **`BootstrapCapacityTables`** (Alt+F8) once to create sheets and tables. Re-run after updating the VBA to remove any blank placeholder rows in row 4.
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
- First data row is **row 4** (headers on row 3). New records write into row 4 instead of leaving it blank.
- `FactoryEquipment` has no `Active` column — remove assignments with the **Remove** button instead.
- `Equipment` has no `EquipmentType` column — use **EquipmentProcesses** for process capabilities.
- Master-table deletes are blocked while junction rows still reference the record.
- `Active` flags accept `True`/`False`, `1`/`0`, or `yes`/`no` when reading existing data.
