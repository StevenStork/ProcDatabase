Attribute VB_Name = "modConstants"
Option Compare Database
Option Explicit

Public Const SCHEMA_VERSION As String = "4"

' Linked tables (user-created; SQL lives in same-folder dataQueries.xlsm).
Public Const TBL_ROUTE_CARD As String = "tblRouteCard"
Public Const TBL_ASSY_STANDARD As String = "tblAssyStnd"
Public Const TBL_OPER_COMPLETIONS As String = "tblOperComps"
Public Const TBL_PROC_TM_YLD As String = "tblProcTmYld"

Public Const SOURCE_WORKBOOK_XLSM As String = "dataQueries.xlsm"

' Local app tables.
Public Const TBL_META As String = "tblMeta"
Public Const TBL_FFA As String = "tblFFA"
Public Const TBL_PRODUCT_LINE As String = "tblProductLine"
Public Const TBL_EQUIPMENT As String = "tblEquipment"
Public Const TBL_PART As String = "tblPart"
Public Const TBL_PART_DASH As String = "tblPartDash"
Public Const TBL_PART_PL As String = "tblPartProductLine"
Public Const TBL_OPERATION As String = "tblOperation"
Public Const TBL_ACTIVE_FILTER As String = "tblActiveAssemblyFilter"

Public Const COL_SHEET_NAME As String = "SheetName"
Public Const COL_ASSEMBLY_NO_FILTER As String = "AssemblyNo"

Public Const QRY_HOME As String = "qryHomeParts"
Public Const QRY_OPERATIONS As String = "qryOperations"
Public Const QRY_EXPORT As String = "qryExportOps"
Public Const QRY_ROUTE_CARD_ACTIVE As String = "qryRouteCardActive"
Public Const QRY_ASSY_STND_ACTIVE As String = "qryAssyStndActive"
Public Const QRY_OPER_COMPS_ACTIVE As String = "qryOperCompsActive"

Public Const FRM_HOME As String = "frmHome"
Public Const FRM_PART As String = "frmPart"
Public Const FRM_FFA As String = "frmFFA"
Public Const FRM_PRODUCT_LINE As String = "frmProductLine"
Public Const FRM_EQUIPMENT As String = "frmEquipment"
Public Const FRM_EXPORT As String = "frmExport"
Public Const FRM_REFERENCES As String = "frmReferences"
Public Const SFRM_DASH As String = "sfrmPartDash"
Public Const SFRM_PL As String = "sfrmPartProductLine"
Public Const SFRM_OPS As String = "sfrmOperation"

Public Const COL_ASSEMBLY_NO As String = "ASSEMBLY NO"
Public Const COL_OPER_SEQ As String = "OPER SEQ"
Public Const COL_OPER_CODE As String = "OPER CODE"
Public Const COL_ASSY_DESC As String = "ASSEMBLY DESCRIPTION"
Public Const COL_OPER_DESC As String = "OPER DESCRIPTION"
Public Const COL_FFA As String = "FFA"
Public Const COL_ORG_CODE As String = "ORG CODE"
Public Const COL_RUN_TIME As String = "RUN TIME (HOURS)"
Public Const COL_SERIAL As String = "S/N"
Public Const COL_LABOR_HPS As String = "LABOR HPS (HOURS)"
Public Const COL_QTY As String = "QTY"
Public Const COL_PROJECT As String = "PROJECT"
Public Const COL_PROGRAM_FAMILY As String = "PROGRAM FAMILY"
Public Const COL_ASSEMBLY_NO_ALT As String = "Assembly No"
Public Const COL_AVG_180 As String = "Avg 180 Day Ex"
Public Const COL_AVG_90 As String = "Avg 90 Day Ex"

' RAG thresholds (same as Excel Home sheet).
Public Const RAG_YELLOW_DAYS As Long = 30
Public Const RAG_RED_DAYS As Long = 90
