Attribute VB_Name = "modConstants"
Option Compare Database
Option Explicit

Public Const SCHEMA_VERSION As String = "10"

' Linked tables (user-created; SQL lives in same-folder dataQueries.xlsm).
Public Const TBL_ROUTE_CARD As String = "tblRouteCard"
Public Const TBL_ASSY_STANDARD As String = "tblAssyStnd"
Public Const TBL_OPER_COMPLETIONS As String = "tblOperComps"
Public Const TBL_PROC_TM_YLD As String = "tblProcTmYld"
Public Const TBL_RCCP As String = "tblRCCP"

' Local app tables.
Public Const TBL_META As String = "tblMeta"
Public Const TBL_FFA As String = "tblFFA"
Public Const TBL_PRODUCT_LINE As String = "tblProductLine"
Public Const TBL_EQUIPMENT As String = "tblEquipment"
Public Const TBL_EQUIPMENT_FFA As String = "tblEquipmentFFA"
Public Const TBL_PART As String = "tblPart"
Public Const TBL_PART_DASH As String = "tblPartDash"
Public Const TBL_PART_PL As String = "tblPartProductLine"
Public Const TBL_OPERATION As String = "tblOperation"
Public Const TBL_ACTIVE_FILTER As String = "tblActiveAssemblyFilter"

Public Const COL_NOTES As String = "Notes"
Public Const COL_ASSEMBLY_NO_FILTER As String = "AssemblyNo"
Public Const COL_PL_CODE As String = "PL Code"
Public Const COL_EQUIP_TYPE As String = "Equipment Type"
Public Const COL_BASE_PART As String = "BasePart"
Public Const COL_ACTIVE As String = "Active"
Public Const COL_HOME_FFA As String = "HomeFFA"
Public Const COL_STATUS_DATE As String = "StatusDate"
Public Const COL_DASH As String = "Dash"
Public Const COL_PRODUCT_LINE As String = "ProductLine"
Public Const COL_FACTORY As String = "Factory"
Public Const COL_USE_FLAG As String = "UseFlag"
Public Const COL_EQUIPMENT As String = "Equipment"
Public Const COL_OP_SEQUENCE As String = "OpSequence"
Public Const COL_OP_CODE As String = "OpCode"

Public Const META_SCHEMA_VERSION As String = "SchemaVersion"
Public Const META_LAST_REFRESH As String = "LastRefresh"
Public Const META_ACTIVE_ASSEMBLY_LIST As String = "ActiveAssemblyList"

Public Const APP_TITLE As String = "ProcDatabase"
Public Const PROP_STARTUP_FORM As String = "StartupForm"
Public Const PROP_APP_TITLE As String = "AppTitle"

Public Const QRY_EXPORT_TEMP As String = "qryExportTemp"

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
Public Const FRM_EQUIPMENT_FFA As String = "frmEquipmentFFA"
Public Const FRM_EQUIPMENT_ENTRY As String = "frmEquipmentEntry"
Public Const FRM_EXPORT As String = "frmExport"
Public Const FRM_REFERENCES As String = "frmReferences"
Public Const SFRM_DASH As String = "sfrmPartDash"
Public Const SFRM_PL As String = "sfrmPartProductLine"
Public Const SFRM_OPS As String = "sfrmOperation"
Public Const SFRM_HOME_LIST As String = "sfrmHomeList"
' Legacy subform from older Equipment UI — deleted during BuildUi; do not recreate.
Public Const LEGACY_SFRM_EQUIPMENT_FFA As String = "sfrmEquipmentFFA"

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
Public Const COL_BASE_PN_TEXT As String = "Base PN: Text"
Public Const COL_PRODUCT_LINE_TEXT As String = "PRODUCT LINE: Text"

' RAG thresholds (same as Excel Home sheet).
Public Const RAG_YELLOW_DAYS As Long = 30
Public Const RAG_RED_DAYS As Long = 90
