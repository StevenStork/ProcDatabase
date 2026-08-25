-- Local tables (created by EnsureSchema)
-- Linked: tblRouteCard, tblAssyStnd, tblOperComps, tblRCCP (user-provided)

-- tblActiveAssemblyFilter (AssemblyNo TEXT PK) — filled from tblRCCP
-- tblPart (BasePart, Active, HomeFFA, StatusDate, Notes)
-- tblPartDash, tblPartProductLine, tblOperation
-- tblFFA, tblProductLine (+ PL Code), tblEquipment, tblEquipmentFFA, tblMeta
-- tblProcTmYld (optional)

-- Filtered views (saved queries):
-- qryRouteCardActive, qryAssyStndActive, qryOperCompsActive
-- qryOperations, qryHomeParts, qryExportOps
