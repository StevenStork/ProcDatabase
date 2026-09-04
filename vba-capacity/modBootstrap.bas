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

    EnsureSheet PARTS_SHEET_NAME, "Parts Index"
    EnsureSheet PART_EDITOR_SHEET_NAME, "Part Number Editor"
    EnsureSheet PART_DASH_CONDITIONS_SHEET_NAME, "Part Dash Conditions"
    EnsureSheet PART_OPERATIONS_SHEET_NAME, "Part Operations"

    MigrateBasePartsTableToPartsSheet

    EnsureTable PARTS_SHEET_NAME, BASE_PARTS_TABLE_NAME, Array( _
        COL_BASE_PART_CODE, COL_FACTORY_CODE, COL_ACTIVE, COL_STATUS_DATE, COL_NOTES)

    EnsureTable PART_DASH_CONDITIONS_SHEET_NAME, PART_DASH_CONDITIONS_TABLE_NAME, Array( _
        COL_BASE_PART_CODE, COL_DASH_CONDITION, COL_SEPARATOR, COL_ACTIVE, COL_NOTES)

    EnsureTable PART_OPERATIONS_SHEET_NAME, PART_OPERATIONS_TABLE_NAME, Array( _
        COL_BASE_PART_CODE, COL_OPER_SEQ, COL_OPERATION_NAME, COL_ACTIVE, COL_NOTES)

    EnsureCacheSheet
    CompactAllCapacityTables

    FormatAdminSheet
    FormatPartsSheet
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

Private Sub MigrateBasePartsTableToPartsSheet()
    Dim editorWs As Worksheet
    Dim partsWs As Worksheet
    Dim tbl As ListObject
    Dim existingPartsTbl As ListObject

    Set editorWs = FindWorksheetByName(PART_EDITOR_SHEET_NAME)
    If editorWs Is Nothing Then Exit Sub

    On Error Resume Next
    Set tbl = editorWs.ListObjects(BASE_PARTS_TABLE_NAME)
    On Error GoTo 0

    If tbl Is Nothing Then Exit Sub

    Set partsWs = FindWorksheetByName(PARTS_SHEET_NAME)
    If partsWs Is Nothing Then Exit Sub

    Set existingPartsTbl = FindTable(BASE_PARTS_TABLE_NAME)
    If Not existingPartsTbl Is Nothing Then
        If existingPartsTbl.Parent.Name = PARTS_SHEET_NAME Then
            tbl.Delete
            Exit Sub
        End If
    End If

    tbl.Name = BASE_PARTS_TABLE_NAME & "_Legacy"
End Sub

Private Sub EnsureCacheSheet()
    Dim ws As Worksheet

    Set ws = FindWorksheetByName(PART_EDITOR_CACHE_SHEET_NAME)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = PART_EDITOR_CACHE_SHEET_NAME
    End If

    ws.Visible = xlSheetVeryHidden
    ws.Cells(CACHE_BASE_PART_CELL).Value = vbNullString
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

    If sheetName = PARTS_SHEET_NAME Or sheetName = FACTORIES_SHEET_NAME Then
        ws.Activate
        ws.Rows(TABLE_HEADER_ROW).Select
        ActiveWindow.FreezePanes = False
        ws.Rows(TABLE_HEADER_ROW + 1).Select
        ActiveWindow.FreezePanes = True
    End If
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
    ws.Range("A9").Value = "LoadPartToEditor"
    ws.Range("A10").Value = "SavePartFromEditor"
    ws.Range("A11").Value = "ClearPartEditor"
    ws.Range("A12").Value = "OpenPartEditorFromPartsIndex"
    ws.Range("A13").Value = "ShowPartEditor"
    ws.Range("A14").Value = "ShowPartOperationsAdmin"
    ws.Range("A15").Value = "RefreshRCCP"
    ws.Range("A16").Value = "RefreshOperComps"
    ws.Range("A17").Value = "RefreshAssyStnd"
    ws.Range("A18").Value = "RefreshRouteCard"
    ws.Range("A19").Value = "RefreshAllLinkedData"
    ws.Range("A20").Value = "BootstrapCapacityTables"
    ws.Columns("A").ColumnWidth = 36
End Sub

Private Sub FormatPartsSheet()
    Dim ws As Worksheet

    Set ws = FindWorksheetByName(PARTS_SHEET_NAME)
    If ws Is Nothing Then Exit Sub

    ws.Range("A2").Value = "All base parts are listed in the table below. Select a row and run Open Part Editor, or go to PartEditor and enter a part number."
    ws.Range("A2").WrapText = True
    ws.Rows("2").RowHeight = 30
End Sub

Private Sub FormatPartEditorSheet()
    Dim ws As Worksheet
    Dim inputRange As Range
    Dim avgHeaderRange As Range
    Dim dashInputRange As Range
    Dim opsInputRange As Range

    Set ws = FindWorksheetByName(PART_EDITOR_SHEET_NAME)
    If ws Is Nothing Then Exit Sub

    ' Title and guidance
    ws.Range("A1").Value = "Part Number Editor"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Color = RGB(32, 56, 100)

    ws.Range("A2").Value = "Enter a part or assembly number below, click Load Part, edit the fields, then click Save Part."
    ws.Range("A2").Font.Color = RGB(80, 80, 80)
    ws.Range("A2").WrapText = True
    ws.Rows("2").RowHeight = 28

    ' Load / identity
    StyleFieldLabel ws, PE_INPUT_ROW, "Part Number"
    StyleInputCell ws.Cells(PE_INPUT_ROW, PE_VALUE_COL), False
    ws.Cells(PE_INPUT_ROW, PE_HINT_COL).Value = "Base part or full assembly (e.g. 8545784-01 or 8545784A01)"
    StyleHintCell ws.Cells(PE_INPUT_ROW, PE_HINT_COL)

    StyleFieldLabel ws, PE_BASE_PART_ROW, "Base Part"
    StyleInputCell ws.Cells(PE_BASE_PART_ROW, PE_VALUE_COL), True
    ws.Cells(PE_BASE_PART_ROW, PE_HINT_COL).Value = "Filled by Load Part (read-only)"
    StyleHintCell ws.Cells(PE_BASE_PART_ROW, PE_HINT_COL)

    ' Master section
    ws.Cells(PE_MASTER_HEADER_ROW, PE_LABEL_COL).Value = "Master Record"
    StyleSectionHeader ws.Range(ws.Cells(PE_MASTER_HEADER_ROW, PE_LABEL_COL), ws.Cells(PE_MASTER_HEADER_ROW, 6))

    StyleFieldLabel ws, PE_ROW_FACTORY, "Factory"
    StyleInputCell ws.Cells(PE_ROW_FACTORY, PE_VALUE_COL), False
    ws.Cells(PE_ROW_FACTORY, PE_HINT_COL).Value = "FactoryCode from Factories sheet (required)"
    StyleHintCell ws.Cells(PE_ROW_FACTORY, PE_HINT_COL)

    StyleFieldLabel ws, PE_ROW_ACTIVE, "Active"
    StyleInputCell ws.Cells(PE_ROW_ACTIVE, PE_VALUE_COL), False
    ws.Cells(PE_ROW_ACTIVE, PE_HINT_COL).Value = "TRUE / FALSE (or Yes / No)"
    StyleHintCell ws.Cells(PE_ROW_ACTIVE, PE_HINT_COL)

    StyleFieldLabel ws, PE_ROW_STATUS_DATE, "Status Date"
    StyleInputCell ws.Cells(PE_ROW_STATUS_DATE, PE_VALUE_COL), False
    ws.Cells(PE_ROW_STATUS_DATE, PE_HINT_COL).Value = "Optional date (yyyy-mm-dd)"
    StyleHintCell ws.Cells(PE_ROW_STATUS_DATE, PE_HINT_COL)

    StyleFieldLabel ws, PE_ROW_NOTES, "Notes"
    StyleInputCell ws.Cells(PE_ROW_NOTES, PE_VALUE_COL), False
    ws.Cells(PE_ROW_NOTES, PE_HINT_COL).Value = "Optional free-text notes"
    StyleHintCell ws.Cells(PE_ROW_NOTES, PE_HINT_COL)

    ' Dash conditions
    ws.Cells(PE_DASH_HEADER_ROW - 1, PE_LABEL_COL).Value = "Dash Conditions"
    StyleSectionHeader ws.Range(ws.Cells(PE_DASH_HEADER_ROW - 1, PE_LABEL_COL), ws.Cells(PE_DASH_HEADER_ROW - 1, 6))
    ws.Cells(PE_DASH_HEADER_ROW - 1, 7).Value = "Add/edit rows below, then Save Part. Separator is usually - or a letter."
    StyleHintCell ws.Cells(PE_DASH_HEADER_ROW - 1, 7)

    ws.Cells(PE_DASH_HEADER_ROW, PE_COL_DASH).Value = "Dash Condition"
    ws.Cells(PE_DASH_HEADER_ROW, PE_COL_SEPARATOR).Value = "Separator"
    ws.Cells(PE_DASH_HEADER_ROW, PE_COL_DASH_ACTIVE).Value = "Active"
    ws.Cells(PE_DASH_HEADER_ROW, PE_COL_DASH_NOTES).Value = "Notes"
    StyleTableHeaderRow ws.Range(ws.Cells(PE_DASH_HEADER_ROW, PE_COL_DASH), ws.Cells(PE_DASH_HEADER_ROW, PE_COL_DASH_NOTES))

    Set dashInputRange = ws.Range( _
        ws.Cells(PE_DASH_DATA_START_ROW, PE_COL_DASH), _
        ws.Cells(PE_DASH_DATA_START_ROW + PE_DASH_MAX_ROWS - 1, PE_COL_DASH_NOTES))
    StyleEditableBlock dashInputRange
    ws.Range( _
        ws.Cells(PE_DASH_DATA_START_ROW, PE_COL_DASH), _
        ws.Cells(PE_DASH_DATA_START_ROW + PE_DASH_MAX_ROWS - 1, PE_COL_DASH)).NumberFormat = "@"

    ' Operations
    ws.Cells(PE_OPS_HEADER_ROW - 1, PE_LABEL_COL).Value = "Operations"
    StyleSectionHeader ws.Range(ws.Cells(PE_OPS_HEADER_ROW - 1, PE_LABEL_COL), ws.Cells(PE_OPS_HEADER_ROW - 1, 6))
    ws.Cells(PE_OPS_HEADER_ROW - 1, 7).Value = "Avg Process Hours / Avg Ex are calculated on Load (read-only)."
    StyleHintCell ws.Cells(PE_OPS_HEADER_ROW - 1, 7)

    ws.Cells(PE_OPS_HEADER_ROW, PE_COL_OPER_SEQ).Value = "Oper Seq"
    ws.Cells(PE_OPS_HEADER_ROW, PE_COL_OPER_NAME).Value = "Operation Name"
    ws.Cells(PE_OPS_HEADER_ROW, PE_COL_OPER_ACTIVE).Value = "Active"
    ws.Cells(PE_OPS_HEADER_ROW, PE_COL_OPER_NOTES).Value = "Notes"
    ws.Cells(PE_OPS_HEADER_ROW, PE_COL_AVG_HOURS).Value = "Avg Process Hours"
    ws.Cells(PE_OPS_HEADER_ROW, PE_COL_AVG_EX).Value = "Avg Ex"
    StyleTableHeaderRow ws.Range(ws.Cells(PE_OPS_HEADER_ROW, PE_COL_OPER_SEQ), ws.Cells(PE_OPS_HEADER_ROW, PE_COL_AVG_EX))

    Set opsInputRange = ws.Range( _
        ws.Cells(PE_OPS_DATA_START_ROW, PE_COL_OPER_SEQ), _
        ws.Cells(PE_OPS_DATA_START_ROW + PE_OPS_MAX_ROWS - 1, PE_COL_OPER_NOTES))
    StyleEditableBlock opsInputRange

    Set avgHeaderRange = ws.Range( _
        ws.Cells(PE_OPS_DATA_START_ROW, PE_COL_AVG_HOURS), _
        ws.Cells(PE_OPS_DATA_START_ROW + PE_OPS_MAX_ROWS - 1, PE_COL_AVG_EX))
    avgHeaderRange.Interior.Color = RGB(235, 238, 242)
    avgHeaderRange.Borders.Color = RGB(190, 198, 210)
    avgHeaderRange.Locked = True

    ' Status
    StyleFieldLabel ws, PE_STATUS_ROW, "Status"
    ws.Cells(PE_STATUS_ROW, PE_VALUE_COL).Font.Italic = True
    ws.Cells(PE_STATUS_ROW, PE_VALUE_COL).Font.Color = RGB(60, 60, 60)
    ws.Cells(PE_STATUS_ROW, PE_HINT_COL).Value = "Load/Save messages appear here"
    StyleHintCell ws.Cells(PE_STATUS_ROW, PE_HINT_COL)

    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 16
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 10
    ws.Columns("F").ColumnWidth = 18
    ws.Columns("G").ColumnWidth = 16
    ws.Columns("H").ColumnWidth = 10
    ws.Columns("I").ColumnWidth = 42

    FormatDashConditionTextColumn
    EnsurePartEditorButtons ws
End Sub

Private Sub StyleFieldLabel(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal captionText As String)
    With ws.Cells(rowIndex, PE_LABEL_COL)
        .Value = captionText
        .Font.Bold = True
        .Font.Color = RGB(40, 50, 65)
        .HorizontalAlignment = xlRight
    End With
End Sub

Private Sub StyleHintCell(ByVal cell As Range)
    cell.Font.Color = RGB(110, 110, 110)
    cell.Font.Italic = True
    cell.Font.Size = 9
End Sub

Private Sub StyleInputCell(ByVal cell As Range, ByVal readOnlyLook As Boolean)
    cell.Borders.Color = RGB(160, 170, 185)
    If readOnlyLook Then
        cell.Interior.Color = RGB(235, 238, 242)
    Else
        cell.Interior.Color = RGB(255, 252, 245)
    End If
End Sub

Private Sub StyleSectionHeader(ByVal headerRange As Range)
    headerRange.Merge
    headerRange.Font.Bold = True
    headerRange.Font.Size = 11
    headerRange.Font.Color = RGB(255, 255, 255)
    headerRange.Interior.Color = RGB(47, 84, 150)
    headerRange.HorizontalAlignment = xlLeft
    headerRange.VerticalAlignment = xlCenter
    headerRange.IndentLevel = 1
    headerRange.RowHeight = 20
End Sub

Private Sub StyleTableHeaderRow(ByVal headerRange As Range)
    headerRange.Font.Bold = True
    headerRange.Font.Color = RGB(40, 50, 65)
    headerRange.Interior.Color = RGB(210, 220, 232)
    headerRange.Borders.Color = RGB(160, 170, 185)
    headerRange.HorizontalAlignment = xlCenter
End Sub

Private Sub StyleEditableBlock(ByVal blockRange As Range)
    blockRange.Interior.Color = RGB(255, 252, 245)
    blockRange.Borders.Color = RGB(190, 198, 210)
End Sub

Private Sub EnsurePartEditorButtons(ByVal ws As Worksheet)
    Dim anchor As Range
    Dim buttonTop As Double
    Dim buttonHeight As Double
    Dim buttonWidth As Double
    Dim gap As Double

    DeleteWorksheetButton ws, PE_BTN_LOAD_NAME
    DeleteWorksheetButton ws, PE_BTN_SAVE_NAME
    DeleteWorksheetButton ws, PE_BTN_CLEAR_NAME

    ' Place buttons to the right of the Part Number input on row 3.
    Set anchor = ws.Cells(PE_INPUT_ROW, 6)
    buttonTop = anchor.Top
    buttonHeight = 24
    buttonWidth = 90
    gap = 8

    AddWorksheetButton ws, PE_BTN_LOAD_NAME, "Load Part", "LoadPartToEditor", _
        anchor.Left, buttonTop, buttonWidth, buttonHeight
    AddWorksheetButton ws, PE_BTN_SAVE_NAME, "Save Part", "SavePartFromEditor", _
        anchor.Left + buttonWidth + gap, buttonTop, buttonWidth, buttonHeight
    AddWorksheetButton ws, PE_BTN_CLEAR_NAME, "Clear", "ClearPartEditor", _
        anchor.Left + 2 * (buttonWidth + gap), buttonTop, buttonWidth, buttonHeight
End Sub

Private Sub AddWorksheetButton( _
    ByVal ws As Worksheet, _
    ByVal buttonName As String, _
    ByVal captionText As String, _
    ByVal onActionName As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal widthPos As Double, _
    ByVal heightPos As Double)

    Dim btn As Button

    Set btn = ws.Buttons.Add(leftPos, topPos, widthPos, heightPos)
    btn.Name = buttonName
    btn.Caption = captionText
    btn.OnAction = onActionName
    btn.Font.Bold = True
End Sub

Private Sub DeleteWorksheetButton(ByVal ws As Worksheet, ByVal buttonName As String)
    On Error Resume Next
    ws.Buttons(buttonName).Delete
    On Error GoTo 0
End Sub

Private Sub FormatDashConditionTextColumn()
    Dim tbl As ListObject
    Dim col As ListColumn

    Set tbl = FindTable(PART_DASH_CONDITIONS_TABLE_NAME)
    If tbl Is Nothing Then Exit Sub
    If Not TableHasColumn(tbl, COL_DASH_CONDITION) Then Exit Sub

    Set col = tbl.ListColumns(COL_DASH_CONDITION)
    col.Range.NumberFormat = "@"
    If Not col.DataBodyRange Is Nothing Then
        col.DataBodyRange.NumberFormat = "@"
    End If
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
