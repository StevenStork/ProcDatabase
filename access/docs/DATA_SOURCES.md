# Data sources for Access ProcDatabase

## Linked tables (required)

These must exist in your `.accdb` before running `BootstrapProcDatabase` / `RefreshAll`:

| Access linked table | Source |
|---------------------|--------|
| `tblRouteCard` | Route_Card |
| `tblAssyStnd` | Assembly_Standard (Assy_Standard) |
| `tblOperComps` | Oper_Completions |
| `tblRCCP` | RCCP — **defines which assemblies are active** |

### Route / standards / completions columns

- `ASSEMBLY NO` (Text)
- `OPER SEQ` (Integer)
- `OPER CODE` (Text)
- `ASSEMBLY DESCRIPTION`, `OPER DESCRIPTION` (Text)
- `FFA`, `ORG CODE` (Text)
- `RUN TIME (HOURS)` / `LABOR HPS (HOURS)` (Decimal)
- `S/N`, `PROJECT`, `PROGRAM FAMILY` on completions

### tblRCCP columns

| Column | Role |
|--------|------|
| `ASSEMBLY NO` | Full assembly; dash portion becomes the active dash condition |
| `Base PN: Text` | Base part (preferred over parsing when present) |
| `ASSEMBLY DESCRIPTION` | Description |
| `PRODUCT LINE: Text` | **PL Code** — matched to `tblProductLine.[PL Code]` |
| `PROGRAM FAMILY` | Program family |
| `FFA` | Seeds Home FFA when blank |
| `ORG CODE` | Org code |

Every row in `tblRCCP` is treated as active. On refresh, Access:

1. Marks matching `tblPart` / `tblPartDash` rows Active
2. Sets `tblPartProductLine.UseFlag` from the RCCP product-line code
3. Fills `tblActiveAssemblyFilter` from distinct `ASSEMBLY NO` values

## dataQueries.xlsm

SQL for sources lives in **`dataQueries.xlsm`**, stored in the **same folder** as the Access database. Access linked tables should point at the same backend those queries use (SQL Server, ODBC, etc.) — Access does not read the `.xlsm` at runtime; the workbook is the query catalog your links were built from.

## Optional: process time / yield

Link `tblProcTmYld` (do not leave an empty local table with that name). Expected columns:

- `Assembly No`
- `OPER SEQ`
- `Avg 180 Day Ex`
- `Avg 90 Day Ex`

If absent, labor averages still work; execution averages stay blank until yield data exists.

## Active assembly filter

1. `RebuildActiveAssemblyFilter` fills local `tblActiveAssemblyFilter` (`AssemblyNo`) from `tblRCCP`.
2. Saved queries narrow linked reads:
   - `qryRouteCardActive`
   - `qryAssyStndActive`
   - `qryOperCompsActive`

If `tblRCCP` is missing, the filter falls back to Home Active × dash Active.

## Product lines

`tblProductLine` has:

- `ProductLine` — display / export name
- `PL Code` — matches `tblRCCP.[PRODUCT LINE: Text]`

Unknown PL codes from RCCP are inserted as new product-line rows (name = code) on refresh.

## Refresh workflow

1. `RefreshLink` on each linked table (including `tblRCCP`).
2. `RebuildCatalogFromStandards` from `tblAssyStnd`.
3. `ApplyRccpSelections` (Active flags + product-line UseFlags).
4. `SeedReferencesFromSources` (FFAs, equipment).
5. `RebuildActiveAssemblyFilter` from RCCP.

Run via Home → **Refresh Linked Data** or `RefreshAll`.
