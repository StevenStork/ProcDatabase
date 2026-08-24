# Access ProcDatabase

Access port of the Excel ProcDatabase concept. Source manufacturing data comes from **linked tables** you already created in the `.accdb`.

## Linked source tables

| Access table   | Source                    |
|----------------|---------------------------|
| `tblRouteCard` | Route_Card                |
| `tblAssyStnd`  | Assembly_Standard         |
| `tblOperComps` | Oper_Completions          |

Shared columns use the same names/types as the Excel/query source (`ASSEMBLY NO`, `OPER SEQ`, `OPER CODE`, `FFA`, `ORG CODE`, etc.).

Optional: link or load process-time/yield data as `tblProcTmYld` (`Assembly No`, `OPER SEQ`, `Avg 180 Day Ex`, `Avg 90 Day Ex`). If missing, labor averages still work; execution averages return blank.

## Local tables (created by VBA)

`tblPart`, `tblPartDash`, `tblPartProductLine`, `tblOperation`, `tblFFA`, `tblProductLine`, `tblEquipment`, `tblMeta`.

## Setup on Windows (with Access)

1. Open your existing `.accdb` that already has the three linked tables.
2. Alt+F11 → File → Import File — import every module under [`access/vba/`](vba/).
3. Immediate window (`Ctrl+G`):

```vb
BuildUi
```

4. Close/reopen the database (startup form is `frmHome`), or run:

```vb
StartProcDatabase
```

5. On Home, click **Refresh Linked Data** (or run `RefreshAll`). That refreshes the linked table connections and rebuilds the part/dash catalog from `tblAssyStnd`.

## Workflow

1. **Refresh Linked Data** — `RefreshLink` on `tblRouteCard` / `tblAssyStnd` / `tblOperComps`, then rebuild catalog (preserves Active / Home FFA / dates).
2. Mark parts **Active** on Home; set Home FFA and status date.
3. **Open Part** — activate dash conditions and product lines; **Seed Ops** pulls route-card ops and fills imported hours/executions.
4. Edit batch size / export overrides on the ops subform (`Process Hours` / `Avg Ex` / `Avg HPU` follow the Excel W/X/Y rules).
5. **Export** — FFA, product line, or all to `.xlsx` beside the database.

## Domain rules (same as Excel VBA)

- Assembly number = `BasePart` + `-` + dash
- Labor hours: average non-zero `LABOR HPS (HOURS)` from `tblOperComps`, else non-zero `RUN TIME (HOURS)` from `tblAssyStnd`
- Executions: average non-zero 180-day from `tblProcTmYld`, else 90-day
- Ops selection: export override when flag is set, else imported; HPU = `(hours * executions) / batch`

## Module map

| Module | Role |
|--------|------|
| `modConstants` | Table/column/form names (`tblAssyStnd`, `tblOperComps`, …) |
| `modLinkedData` | Refresh linked sources |
| `modCatalog` | Build parts/dashes from standards |
| `modAverages` | Labor / process-time averages |
| `modOperations` | Seed ops from route card |
| `modExport` | Excel exports |
| `modSchema` / `modUi` / `modApp` | Local schema, forms, entry points |
| `modUtils` | Shared helpers |

## Note

This environment cannot compile a binary `.accdb`. Keep using your Access file; these modules are text to import into it.
