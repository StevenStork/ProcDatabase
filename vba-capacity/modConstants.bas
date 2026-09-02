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

Public Const PART_EDITOR_SHEET_NAME As String = "PartEditor"
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
Public Const COL_OPER_SEQ As String = "OperSeq"
Public Const COL_OPERATION_NAME As String = "OperationName"

Public Const FORM_MARGIN As Single = 12
Public Const FORM_BUTTON_HEIGHT As Single = 26
Public Const FORM_BUTTON_WIDTH As Single = 76
Public Const FORM_BUTTON_GAP As Single = 8
