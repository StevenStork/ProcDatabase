Attribute VB_Name = "modSchema"
Option Explicit

'==============================================================================
' Workbook contract: sheet names, A1 labels, table names, and header strings.
' Callers look up operations by these header names, not by column letters.
'==============================================================================

Public Const HOME_SHEET_NAME As String = "Home"
Public Const TEMPLATE_SHEET_NAME As String = "Part Number Template"
Public Const REFERENCES_SHEET_NAME As String = "References"
Public Const DATA_SHEET_NAME As String = "Data"
Public Const ASSEMBLY_STANDARDS_SHEET_NAME As String = "Assembly Standards"

Public Const CATEGORY_CELL As String = "A1"
Public Const HOME_LABEL_VALUE As String = "Home"
Public Const PART_LABEL_VALUE As String = "Part"
Public Const REFS_LABEL_VALUE As String = "Refs"
Public Const EXPORT_LABEL_VALUE As String = "Export"
Public Const DATA_LABEL_VALUE As String = "Data"
Public Const STANDARDS_LABEL_VALUE As String = "Standards"

Public Const PART_NUMBER_CELL As String = "C2"
Public Const PART_LABEL_CELL As String = "A1"

Public Const LIST_HEADER_ROW As Long = 8
Public Const LIST_START_ROW As Long = 9
Public Const HOME_FFA_LABEL_CELL As String = "C8"
Public Const HOME_FFA_VALUE_CELL As String = "C9"
Public Const HOME_FFA_LABEL As String = "Home FFA"
Public Const DASH_VALUE_COLUMN As String = "E"
Public Const DASH_CHECKBOX_COLUMN As String = "F"
Public Const PRODUCT_LINE_VALUE_COLUMN As String = "G"
Public Const PRODUCT_LINE_CHECKBOX_COLUMN As String = "H"
Public Const INPUT_FILL_RGB As Long = 16377045 ' RGB(213, 229, 249)

Public Const OPS_FIRST_COLUMN As String = "M"
Public Const OPS_LAST_COLUMN As String = "Z"
Public Const OPS_HEADER_ROW As Long = 8
Public Const OPS_INPUT_LAST_COLUMN As String = "S"
Public Const PART_OPS_TABLE_NAME As String = "PartOpsTbl"

Public Const HDR_OP_SEQUENCE As String = "Operation Sequence"
Public Const HDR_OP_CODE As String = "Operation Code"
Public Const HDR_IMPORTED_HOURS As String = "Imported Process Hours"
Public Const HDR_IMPORTED_EX As String = "Imported Average Executions"
Public Const HDR_BATCH_SIZE As String = "Batch Size"
Public Const HDR_EXPORT_HOURS As String = "Export Process Hours"
Public Const HDR_EXPORT_EX As String = "Export Average Executions"
Public Const HDR_EQUIPMENT_TYPE As String = "Equipment Type"
Public Const HDR_USE_EXPORT_HOURS As String = "Use Export Hours"
Public Const HDR_USE_EXPORT_EX As String = "Use Export Executions"
Public Const HDR_PROCESS_HOURS As String = "Process Hours"
Public Const HDR_AVG_EX As String = "Average Executions"
Public Const HDR_AVG_HPU As String = "Average HPUs"
Public Const HDR_FFA As String = "FFA"

Public Const HOME_PART_TABLE_HEADER_ROW As Long = 5
Public Const HOME_PART_TABLE_FIRST_DATA_ROW As Long = 6
Public Const HOME_PART_TABLE_FIRST_COLUMN As String = "C"
Public Const HOME_PART_TABLE_LAST_COLUMN As String = "I"
Public Const HOME_PART_TABLE_BASE_PART_COLUMN As String = "C"
Public Const HOME_PART_TABLE_ACTIVE_COLUMN As String = "D"
Public Const HOME_PART_TABLE_DATE_COLUMN As String = "E"
Public Const HOME_PART_TABLE_DAYS_COLUMN As String = "F"
Public Const HOME_PART_TABLE_HIGHLIGHT_COLUMN_G As String = "G"
Public Const HOME_PART_TABLE_FFA_COLUMN As String = "H"
Public Const HOME_PART_TABLE_FACTORY_COLUMN As String = "I"
Public Const HEADER_BASE_PART As String = "Base Part Number"

Public Const TBL_PARTS_NAME As String = "tblParts"
Public Const TBL_OPS_NAME As String = "tblOperations"
Public Const COL_PARTS_BASE As String = "Base Part"
Public Const COL_PARTS_ACTIVE As String = "Active"
Public Const COL_PARTS_HOME_FFA As String = "Home FFA"
Public Const COL_PARTS_FFAS_LEGACY As String = "FFAs"
Public Const COL_PARTS_FACTORIES As String = "Factories"
Public Const COL_EXPORT_MADE_IN_FFA As String = "Made In FFA"
Public Const COL_PARTS_PRODUCT_LINES As String = "Product Lines"
Public Const COL_PARTS_DASHES As String = "Dashes"
Public Const COL_PARTS_UI_SCHEMA As String = "UiSchema"
Public Const COL_PARTS_LIST_SIG As String = "ListSig"
Public Const COL_PARTS_OPS_DIRTY As String = "OpsDirty"
Public Const COL_PARTS_OPS_ROW_COUNT As String = "OpsRowCount"
Public Const COL_PARTS_SHEET_NAME As String = "SheetName"
Public Const COL_OPS_PART_NUMBER As String = "Part Number"

Public Const DATA_HOME_HASH_CELL As String = "B2"
Public Const DATA_EXPORT_HASH_CELL As String = "B3"
Public Const DATA_REFS_DIRTY_CELL As String = "B4"
Public Const DATA_UI_SCHEMA_CELL As String = "B5"
Public Const UI_SCHEMA_VERSION As String = "2"
Public Const DATA_PARTS_HEADER_ROW As Long = 8
Public Const DATA_PARTS_FIRST_COL As Long = 1
Public Const DATA_OPS_HEADER_ROW As Long = 8
Public Const DATA_OPS_FIRST_COL As Long = 13

Public Const EXPORT_TYPE_CELL As String = "A2"
Public Const EXPORT_KEY_CELL As String = "A3"
Public Const EXPORT_TYPE_FFA As String = "FFA"
Public Const EXPORT_TYPE_PRODUCT_LINE As String = "ProductLine"
Public Const FFA_EXPORT_SHEET_NAME As String = "FFA Export"
Public Const FFA_EXPORT_KEY As String = "All"
Public Const FFA_SHEET_PREFIX As String = "FFA - "
Public Const PRODUCT_LINE_SHEET_PREFIX As String = "PL - "
Public Const EXPORT_HEADER_ROW As Long = 1
Public Const EXPORT_FIRST_DATA_ROW As Long = 2
Public Const EXPORT_FIRST_COLUMN As Long = 3
Public Const EXPORT_LAST_COLUMN As Long = 12
Public Const EXPORT_COLUMN_COUNT As Long = 10

Public Const REFS_FFA_COLUMN As String = "B"
Public Const REFS_FACTORY_COLUMN As String = "C"
Public Const REFS_PRODUCT_LINE_COLUMN As String = "D"
Public Const REFS_EQUIPMENT_COLUMN As String = "E"
Public Const REFS_EQUIPMENT_OWNERS_COLUMN As String = "F"
Public Const REFS_FIRST_DATA_ROW As Long = 2

Public Const ASSY_STANDARDS_TABLE_NAME As String = "AssyStndTbl"
Public Const COL_ASSEMBLY_NO As String = "ASSEMBLY NO"

Public Const XL_TYPE_CHECKBOX As Long = 2

Public Function OpsHeaderNames() As Variant
    OpsHeaderNames = Array( _
        HDR_OP_SEQUENCE, _
        HDR_OP_CODE, _
        HDR_IMPORTED_HOURS, _
        HDR_IMPORTED_EX, _
        HDR_BATCH_SIZE, _
        HDR_EXPORT_HOURS, _
        HDR_EXPORT_EX, _
        HDR_EQUIPMENT_TYPE, _
        HDR_USE_EXPORT_HOURS, _
        HDR_USE_EXPORT_EX, _
        HDR_PROCESS_HOURS, _
        HDR_AVG_EX, _
        HDR_AVG_HPU, _
        HDR_FFA)
End Function

Public Function ExportHeaderNames() As Variant
    ExportHeaderNames = Array( _
        COL_OPS_PART_NUMBER, _
        "Op Sequence", _
        "Op Code", _
        "Process Hours", _
        "Avg Ex", _
        "Batch Size", _
        "Avg HPU", _
        "Equipment Type", _
        COL_PARTS_HOME_FFA, _
        COL_EXPORT_MADE_IN_FFA)
End Function

Public Function PartsHeaderNames() As Variant
    PartsHeaderNames = Array( _
        COL_PARTS_BASE, _
        COL_PARTS_ACTIVE, _
        COL_PARTS_HOME_FFA, _
        COL_PARTS_FACTORIES, _
        COL_PARTS_PRODUCT_LINES, _
        COL_PARTS_DASHES, _
        COL_PARTS_UI_SCHEMA, _
        COL_PARTS_LIST_SIG, _
        COL_PARTS_OPS_DIRTY, _
        COL_PARTS_OPS_ROW_COUNT, _
        COL_PARTS_SHEET_NAME)
End Function

Public Function OpsStoreHeaderNames() As Variant
    OpsStoreHeaderNames = Array( _
        COL_OPS_PART_NUMBER, _
        HDR_OP_SEQUENCE, _
        HDR_OP_CODE, _
        HDR_IMPORTED_HOURS, _
        HDR_IMPORTED_EX, _
        HDR_BATCH_SIZE, _
        HDR_EXPORT_HOURS, _
        HDR_EXPORT_EX, _
        HDR_EQUIPMENT_TYPE, _
        HDR_USE_EXPORT_HOURS, _
        HDR_USE_EXPORT_EX, _
        HDR_PROCESS_HOURS, _
        HDR_AVG_EX, _
        HDR_AVG_HPU, _
        HDR_FFA)
End Function
