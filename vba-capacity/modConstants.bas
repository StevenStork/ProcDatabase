Attribute VB_Name = "modConstants"
Option Explicit

'==============================================================================
' Factory Capacity Database — shared constants
'==============================================================================

Public Const ADMIN_SHEET_NAME As String = "Admin"
Public Const FACTORIES_SHEET_NAME As String = "Factories"
Public Const EQUIPMENT_SHEET_NAME As String = "Equipment"
Public Const PROCESS_TYPES_SHEET_NAME As String = "ProcessTypes"
Public Const FACTORY_EQUIPMENT_SHEET_NAME As String = "FactoryEquipment"
Public Const EQUIPMENT_PROCESSES_SHEET_NAME As String = "EquipmentProcesses"

Public Const PARTS_SHEET_NAME As String = "Parts"
Public Const PART_EDITOR_SHEET_NAME As String = "PartEditor"
Public Const PART_EDITOR_CACHE_SHEET_NAME As String = "PartEditorCache"
Public Const PART_DASH_CONDITIONS_SHEET_NAME As String = "PartDashConditions"
Public Const PART_OPERATIONS_SHEET_NAME As String = "PartOperations"

Public Const FACTORIES_TABLE_NAME As String = "FactoriesTbl"
Public Const EQUIPMENT_TABLE_NAME As String = "EquipmentTbl"
Public Const PROCESS_TYPES_TABLE_NAME As String = "ProcessTypesTbl"
Public Const FACTORY_EQUIPMENT_TABLE_NAME As String = "FactoryEquipmentTbl"
Public Const EQUIPMENT_PROCESSES_TABLE_NAME As String = "EquipmentProcessTbl"

Public Const BASE_PARTS_TABLE_NAME As String = "BasePartsTbl"
Public Const PART_DASH_CONDITIONS_TABLE_NAME As String = "PartDashConditionsTbl"
Public Const PART_OPERATIONS_TABLE_NAME As String = "PartOperationsTbl"

' Linked Power Query tables (connection-only; ListObject when loaded to a sheet).
Public Const LINKED_ROUTE_CARD_TABLE As String = "tblRouteCard"
Public Const LINKED_OPER_COMPS_TABLE As String = "tblOperComps"
Public Const LINKED_RCCP_TABLE As String = "tblRCCP"
Public Const LINKED_TIME_YIELD_TABLE As String = "tblTimeYield"
Public Const LINKED_ASSY_STND_TABLE As String = "tblAssyStnd"

Public Const COL_FFA As String = "FFA"

Public Const TABLE_HEADER_ROW As Long = 3
Public Const TABLE_FIRST_DATA_ROW As Long = 4

Public Const COL_FACTORY_CODE As String = "FactoryCode"
Public Const COL_FACTORY_NAME As String = "FactoryName"
Public Const COL_EQUIPMENT_CODE As String = "EquipmentCode"
Public Const COL_EQUIPMENT_NAME As String = "EquipmentName"
Public Const COL_PROCESS_TYPE_CODE As String = "ProcessTypeCode"
Public Const COL_PROCESS_TYPE_NAME As String = "ProcessTypeName"
Public Const COL_ACTIVE As String = "Active"
Public Const COL_NOTES As String = "Notes"

Public Const COL_BASE_PART_CODE As String = "BasePartCode"
Public Const COL_STATUS_DATE As String = "StatusDate"
Public Const COL_DASH_CONDITION As String = "DashCondition"
Public Const COL_SEPARATOR As String = "Separator"
Public Const COL_OPER_SEQ As String = "OperSeq"
Public Const COL_OPERATION_NAME As String = "OperationName"

' Linked source column headers.
Public Const COL_ASSEMBLY_NO As String = "ASSEMBLY NO"
Public Const COL_RCCP_BASE_PN As String = "Base PN: Text"
Public Const COL_OPER_SEQ_SOURCE As String = "OPER SEQ"
Public Const COL_LABOR_HPS As String = "LABOR HPS (HOURS)"
Public Const COL_RUN_TIME As String = "RUN TIME (HOURS)"
Public Const COL_AVG_180_DAY_EX As String = "Avg 180 Day Ex"
Public Const COL_AVG_90_DAY_EX As String = "Avg 90 Day Ex"
Public Const COL_AVG_PROCESS_HOURS As String = "Avg Process Hours"
Public Const COL_AVG_EX As String = "Avg Ex"

' PartEditor sheet layout (column B = labels, data from column C).
Public Const PE_LABEL_COL As Long = 2
Public Const PE_VALUE_COL As Long = 3
Public Const PE_INPUT_ROW As Long = 3
Public Const PE_BASE_PART_ROW As Long = 4
Public Const PE_ROW_FACTORY As Long = 6
Public Const PE_ROW_ACTIVE As Long = 7
Public Const PE_ROW_STATUS_DATE As Long = 8
Public Const PE_ROW_NOTES As Long = 9
Public Const PE_DASH_HEADER_ROW As Long = 12
Public Const PE_DASH_COL_START As Long = 3
Public Const PE_DASH_DATA_START_ROW As Long = 14
Public Const PE_DASH_MAX_ROWS As Long = 20
Public Const PE_OPS_HEADER_ROW As Long = 36
Public Const PE_OPS_COL_START As Long = 3
Public Const PE_OPS_DATA_START_ROW As Long = 38
Public Const PE_OPS_MAX_ROWS As Long = 30
Public Const PE_STATUS_ROW As Long = 70

Public Const PE_COL_OPER_SEQ As Long = 3
Public Const PE_COL_OPER_NAME As Long = 4
Public Const PE_COL_OPER_ACTIVE As Long = 5
Public Const PE_COL_OPER_NOTES As Long = 6
Public Const PE_COL_AVG_HOURS As Long = 7
Public Const PE_COL_AVG_EX As Long = 8

Public Const PE_COL_DASH As Long = 3
Public Const PE_COL_SEPARATOR As Long = 4
Public Const PE_COL_DASH_ACTIVE As Long = 5
Public Const PE_COL_DASH_NOTES As Long = 6

' PartEditorCache hidden sheet layout.
Public Const CACHE_BASE_PART_CELL As String = "A1"
Public Const CACHE_DASH_START_ROW As Long = 3
Public Const CACHE_OPS_START_ROW As Long = 30

Public Const FORM_MARGIN As Single = 12
Public Const FORM_BUTTON_HEIGHT As Single = 26
Public Const FORM_BUTTON_WIDTH As Single = 76
Public Const FORM_BUTTON_GAP As Single = 8
