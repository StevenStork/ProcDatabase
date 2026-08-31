-- Local tables (created by EnsureSchema)
-- Linked: tblRouteCard, tblAssyStnd, tblOperComps, tblRCCP (user-provided)

-- tblActiveAssemblyFilter (AssemblyNo TEXT PK) — filled from tblRCCP
-- tblPart (BasePart, Active, HomeFFA, StatusDate, Notes)
-- tblOperation (... OpSequence, OpLine, MadeInFFA, Equipment, EquipmentType, ProcessHours, AvgEx, ...)
-- tblPartDash, tblPartProductLine, tblOperation
-- tblFFA, tblProductLine (+ PL Code), tblEquipment (+ Equipment Type), tblEquipmentFFA, tblMeta
-- tblProcTmYld (optional)

-- Filtered views (saved queries):
-- qryRouteCardActive, qryAssyStndActive, qryOperCompsActive
-- qryOperations, qryHomeParts, qryExportOps
