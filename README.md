# ProcDatabase

Excel VBA process database, plus an **Access** port using linked source tables.

## Excel

VBA modules: [`vba/`](vba/). Import into the workbook as before.

## Access

Full port: [`access/README.md`](access/README.md).

Your `.accdb` links:

- `tblRouteCard` ← Route_Card  
- `tblAssyStnd` ← Assembly_Standard  
- `tblOperComps` ← Oper_Completions  

SQL lives in same-folder `dataQueries.xlsm`. Import [`access/vba/*.bas`](access/vba/) and run `BootstrapProcDatabase`.
