# Data sources for Access ProcDatabase

## Linked tables (required)

These must exist in your `.accdb` before running `BootstrapProcDatabase`:

| Access linked table | Source query / sheet in `dataQueries.xlsm` |
|---------------------|---------------------------------------------|
| `tblRouteCard` | Route_Card |
| `tblAssyStnd` | Assembly_Standard (Assy_Standard) |
| `tblOperComps` | Oper_Completions |

Column names and types match the Excel variant:

- `ASSEMBLY NO` (Text)
- `OPER SEQ` (Integer)
- `OPER CODE` (Text)
- `ASSEMBLY DESCRIPTION`, `OPER DESCRIPTION` (Text)
- `FFA`, `ORG CODE` (Text)
- `RUN TIME (HOURS)` / `LABOR HPS (HOURS)` (Decimal)
- `S/N`, `PROJECT`, `PROGRAM FAMILY` on completions

## dataQueries.xlsm

SQL for the three sources lives in **`dataQueries.xlsm`**, stored in the **same folder** as the Access database. Access linked tables should point at the same backend those queries use (SQL Server, ODBC, etc.) — Access does not read the `.xlsm` at runtime; the workbook is the query catalog your links were built from.

## Optional: process time / yield

Link or load `tblProcTmYld` with:

- `Assembly No`, `OPER SEQ`, `Avg 180 Day Ex`, `Avg 90 Day Ex`

If absent, labor averages still work; execution averages stay blank until yield data exists.

## Active assembly filter (replaces Excel parameter patching)

Excel `UpdateRouteCardConnection` rewrites `@assembly_number` in connection SQL from active Home parts × checked dashes.

Access equivalent:

1. `RebuildActiveAssemblyFilter` fills local `tblActiveAssemblyFilter` (`AssemblyNo`).
2. Saved queries narrow linked reads:
   - `qryRouteCardActive`
   - `qryAssyStndActive`
   - `qryOperCompsActive`

When no parts/dashes are active, seed/average code reads the full linked tables.

## Refresh workflow

1. `RefreshLink` on each linked table (pulls latest from backend).
2. `RebuildCatalogFromStandards` from `tblAssyStnd`.
3. `SeedReferencesFromSources` (FFAs, equipment from standards/route card).
4. `RebuildActiveAssemblyFilter`.

Run via Home → **Refresh Linked Data** or `RefreshAll`.
