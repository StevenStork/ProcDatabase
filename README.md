# ProcDatabase

Excel VBA database for factory part numbers, process hours, and FFA / product-line exports.

The workbook shell is [`ProcessDatabase/ProcessDatabase.xlsm`](ProcessDatabase/ProcessDatabase.xlsm). Macros live in [`vba/`](vba/) and must be imported into that workbook (the `.xlsm` in git has sheets and tables, not a compiled `vbaProject.bin`).

## Workbook layout

| Sheet | `A1` label | Role |
| --- | --- | --- |
| Home | `Home` | Navigator, part list `C:I`, Show/Hide buttons, form buttons |
| Part Number Template | `Part` | Template copied for each active part (`C2` = base part) |
| Part sheets | `Part` | Editor: FFA / dash / product-line checkboxes and the ops table |
| References | `Refs` | Very hidden. FFA+factory, product lines, equipment. Edit via **Update References** |
| Data | `Data` | Very hidden. `tblParts`, `tblOperations`, dirty flags and short hashes |
| `FFA - …` / `PL - …` | `Export` | Generated copy-out sheets. Show/Hide as the Export category |
| Assembly Standards | `Standards` | Source table `AssyStndTbl` for dash conditions |

Factory codes are derived from checked FFAs via References `B:C`. They are not a separate checklist.

## Part operation table (`PartOpsTbl`, columns M:Z)

| Header | Notes |
| --- | --- |
| Operation Sequence, Operation Code | Inputs |
| Imported Process Hours, Imported Average Executions, Batch Size | Inputs |
| Export Process Hours, Export Average Executions, Equipment Type | Inputs |
| Use Export Hours / Use Export Executions | In-cell checkboxes (`U`/`V`) |
| Process Hours | `=IF(AND(exportHours<>"",U=TRUE),exportHours,importedHours)` |
| Average Executions | Same pattern on executions |
| Average HPUs | `=(ProcessHours*AvgEx)/BatchSize` |
| FFA | Step the row is complete at |

## Canonical store (Data sheet)

- `tblParts` — base part, active, FFAs, factories, product lines, dashes, `ListSig`, `OpsDirty`, row count, sheet name
- `tblOperations` — every operation row across part sheets (full named fields)
- `B2` `HomeListHash`, `B3` `ExportOpsHash` — short checksums after a successful rebuild
- `B4` `RefsDirty` — set when References are saved

`Worksheet_Change` on a part ops table or checkbox list sets `OpsDirty`. **Refresh All**, **Update Exports**, and `Workbook_BeforeSave` sync dirty parts into the store, then rebuild Home `H:I` and export sheets. SheetActivate only does cheap UI (checkbox lists, buttons, gridlines). Do not put cache stamps in `A2`/`A3` on Home or Part sheets. Export sheets use `A2`/`A3` as type and key metadata only.

## Export sheets (8 columns)

`Part Number | Op Sequence | Op Code | Process Hours | Avg Ex | Batch Size | Avg HPU | Equipment Type`

- FFA sheets: ops rows whose FFA matches the sheet (filter only; FFA is not an export column)
- Product-line sheets: every ops row from parts that have that product line checked

## Home buttons

- Show/Hide by `A1` category, including **Export**. `Refs` and `Data` stay very hidden.
- Show/Hide part sheets by FFA
- **Update References**, **Update Exports**, **Refresh All** (S:Y from row 5)

## Importing VBA into the workbook

1. Open `ProcessDatabase/ProcessDatabase.xlsm`.
2. Enable macros. For `ImportSourceModules`, also enable **Trust access to the VBA project object model**.
3. Alt+F11 → File → Import File for each `vba/*.bas` and `vba/*.frm`, **or** run `ImportSourceModules` if the workbook sits next to the `vba` folder (repo root) or in `ProcessDatabase/` with `../vba`.
4. Paste [`vba/ThisWorkbook_SheetActivate.bas.txt`](vba/ThisWorkbook_SheetActivate.bas.txt) into **ThisWorkbook**.
5. Rebuild the workbook shell with `python3 ProcessDatabase/build_workbook.py` if the sheet layout needs regenerating (this does not write VBA).
