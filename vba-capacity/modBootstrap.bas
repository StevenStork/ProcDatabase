Attribute VB_Name = "modBootstrap"
Option Explicit

'==============================================================================
' BootstrapCapacityTables
' Creates missing sheets and ListObject tables with the expected headers.
' Safe to re-run; existing tables are validated but not cleared.
'==============================================================================

Public Sub BootstrapCapacityTables()
    On Error GoTo CleanUp

    OptimizeExcel True

    EnsureSheet ADMIN_SHEET_NAME, "Factory Capacity Database"
    EnsureSheet FACTORIES_SHEET_NAME, "Factories"
    EnsureSheet EQUIPMENT_SHEET_NAME, "Equipment"
    EnsureSheet PROCESS_TYPES_SHEET_NAME, "Process Types"
    EnsureSheet FACTORY_EQUIPMENT_SHEET_NAME, "Factory Equipment Assignments"
    EnsureSheet EQUIPMENT_PROCESSES_SHEET_NAME, "Equipment Process Capabilities"

    EnsureTable FACTORIES_SHEET_NAME, FACTORIES_TABLE_NAME, Array( _
        COL_FACTORY_CODE, COL_FACTORY_NAME, COL_ACTIVE, COL_NOTES)

    EnsureTable EQUIPMENT_SHEET_NAME, EQUIPMENT_TABLE_NAME, Array( _
        COL_EQUIPMENT_CODE, COL_EQUIPMENT_NAME, COL_ACTIVE, COL_NOTES)

    EnsureTable PROCESS_TYPES_SHEET_NAME, PROCESS_TYPES_TABLE_NAME, Array( _
        COL_PROCESS_TYPE_CODE, COL_PROCESS_TYPE_NAME, COL_ACTIVE, COL_NOTES)

    EnsureTable FACTORY_EQUIPMENT_SHEET_NAME, FACTORY_EQUIPMENT_TABLE_NAME, Array( _
        COL_FACTORY_CODE, COL_EQUIPMENT_CODE, COL_NOTES)

    EnsureTable EQUIPMENT_PROCESSES_SHEET_NAME, EQUIPMENT_PROCESSES_TABLE_NAME, Array( _
        COL_EQUIPMENT_CODE, COL_PROCESS_TYPE_CODE, COL_ACTIVE, COL_NOTES)

    EnsureSheet PART_EDITOR_SHEET_NAME, "Part Number Editor"
    EnsureSheet PART_DASH_CONDITIONS_SHEET_NAME, "Part Dash Conditions"
    EnsureSheet PART_OPERATIONS_SHEET_NAME, "Part Operations"

    EnsureTable PART_EDITOR_SHEET_NAME, BASE_PARTS_TABLE_NAME, Array( _
        COL_BASE_PART_CODE, COL_FACTORY_CODE, COL_ACTIVE, COL_STATUS_DATE, COL_NOTES)

    EnsureTable PART_DASH_CONDITIONS_SHEET_NAME, PART_DASH_CONDITIONS_TABLE_NAME, Array( _
        COL_BASE_PART_CODE, COL_DASH_CONDITION, COL_ACTIVE, COL_NOTES)

    EnsureTable PART_OPERATIONS_SHEET_NAME, PART_OPERATIONS_TABLE_NAME, Array( _
        COL_BASE_PART_CODE, COL_OPER_SEQ, COL_OPERATION_NAME, COL_ACTIVE, COL_NOTES)

    CompactAllCapacityTables

    FormatAdminSheet
    FormatPartEditorSheet

CleanUp:
    OptimizeExcel False
End Sub

Private Sub EnsureSheet(ByVal sheetName As String, ByVal titleText As String)
    Dim ws As Worksheet

    Set ws = FindWorksheetByName(sheetName)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    ws.Range("A1").Value = titleText
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
End Sub

Private Sub EnsureTable(ByVal sheetName As String, ByVal tableName As String, ByVal headers As Variant)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lastCol As Long
    Dim headerIndex As Long
    Dim tableRange As Range

    Set ws = FindWorksheetByName(sheetName)
    If ws Is Nothing Then Exit Sub

    lastCol = UBound(headers) - LBound(headers) + 1

    For headerIndex = LBound(headers) To UBound(headers)
        ws.Cells(TABLE_HEADER_ROW, headerIndex - LBound(headers) + 1).Value = CStr(headers(headerIndex))
    Next headerIndex

    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0

    If tbl Is Nothing Then
        Set tableRange = ws.Range(ws.Cells(TABLE_HEADER_ROW, 1), ws.Cells(TABLE_HEADER_ROW, lastCol))
        Set tbl = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
        tbl.Name = tableName
    Else
        EnsureTableHeaders tbl, headers
    End If

    ws.Rows(TABLE_HEADER_ROW).Font.Bold = True
    ws.Activate
    ws.Rows(TABLE_HEADER_ROW).Select
    ActiveWindow.FreezePanes = False
    ws.Rows(TABLE_HEADER_ROW + 1).Select
    ActiveWindow.FreezePanes = True
End Sub

Private Sub EnsureTableHeaders(ByVal tbl As ListObject, ByVal headers As Variant)
    Dim headerIndex As Long
    Dim expectedName As String
    Dim newCol As ListColumn

    For headerIndex = LBound(headers) To UBound(headers)
        expectedName = CStr(headers(headerIndex))
        If Not TableHasColumn(tbl, expectedName) Then
            Set newCol = tbl.ListColumns.Add
            newCol.Name = expectedName
            tbl.HeaderRowRange.Cells(1, newCol.Index).Value = expectedName
        End If
    Next headerIndex
End Sub

Private Sub FormatAdminSheet()
    Dim ws As Worksheet

    Set ws = FindWorksheetByName(ADMIN_SHEET_NAME)
    If ws Is Nothing Then Exit Sub

    ws.Range("A3").Value = "Use the buttons below or run these macros:"
    ws.Range("A4").Value = "ShowFactoryAdmin"
    ws.Range("A5").Value = "ShowEquipmentAdmin"
    ws.Range("A6").Value = "ShowProcessTypeAdmin"
    ws.Range("A7").Value = "ShowFactoryEquipmentAdmin"
    ws.Range("A8").Value = "ShowEquipmentProcessAdmin"
    ws.Range("A9").Value = "ShowPartEditor"
    ws.Range("A10").Value = "EditSelectedPartFromSheet"
    ws.Range("A11").Value = "ShowPartOperationsAdmin"
    ws.Range("A12").Value = "BootstrapCapacityTables"
    ws.Columns("A").ColumnWidth = 34
End Sub

Private Sub FormatPartEditorSheet()
    Dim ws As Worksheet

    Set ws = FindWorksheetByName(PART_EDITOR_SHEET_NAME)
    If ws Is Nothing Then Exit Sub

    ws.Range("A2").Value = "All base parts are listed in the table below. Each part has one factory. Use Edit Selected Part to manage dash conditions and operations."
    ws.Range("A2").WrapText = True
    ws.Rows("2").RowHeight = 30
End Sub

Private Function FindWorksheetByName(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set FindWorksheetByName = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Sub CompactAllCapacityTables()
    DeleteEmptyTableRows FindTable(FACTORIES_TABLE_NAME)
    DeleteEmptyTableRows FindTable(EQUIPMENT_TABLE_NAME)
    DeleteEmptyTableRows FindTable(PROCESS_TYPES_TABLE_NAME)
    DeleteEmptyTableRows FindTable(FACTORY_EQUIPMENT_TABLE_NAME)
    DeleteEmptyTableRows FindTable(EQUIPMENT_PROCESSES_TABLE_NAME)
    DeleteEmptyTableRows FindTable(BASE_PARTS_TABLE_NAME)
    DeleteEmptyTableRows FindTable(PART_DASH_CONDITIONS_TABLE_NAME)
    DeleteEmptyTableRows FindTable(PART_OPERATIONS_TABLE_NAME)
End Sub
