# ProcDatabase

Excel VBA process database, plus an **Access** port that uses linked source tables.

## Excel (existing)

VBA modules live under [`vba/`](vba/). Import into the workbook as before.

## Access (new)

See [`access/README.md`](access/README.md).

Your `.accdb` already links:

| Access table   | Source            |
|----------------|-------------------|
| `tblRouteCard` | Route_Card        |
| `tblAssyStnd`  | Assembly_Standard |
| `tblOperComps` | Oper_Completions  |

Import [`access/vba/*.bas`](access/vba/) into that database and run `BuildUi` from the Immediate window.
