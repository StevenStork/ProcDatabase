Attribute VB_Name = "modBootstrap"
Option Explicit

'==============================================================================
' BootstrapCapacityTables
' Creates missing sheets and ListObject tables with the expected headers.
' Safe to re-run; existing tables are validated but not cleared.
'==============================================================================

Public Sub BootstrapCapacityTables()
    Dim formatError As String
    Dim bootstrapError As String
    Dim bootstrapErrorNumber As Long
    Dim currentStep As String

    On Error GoTo FailBootstrap

    currentStep = "OptimizeExcel True"
    OptimizeExcel True

    currentStep = "EnsureSheet Admin/Factories/Equipment/ProcessTypes"
    EnsureSheet ADMIN_SHEET_NAME, "Factory Capacity Database"
    EnsureSheet FACTORIES_SHEET_NAME, "Factories"
    EnsureSheet EQUIPMENT_SHEET_NAME, "Equipment"
    EnsureSheet PROCESS_TYPES_SHEET_NAME, "Process Types"
    EnsureSheet FACTORY_EQUIPMENT_SHEET_NAME, "Factory Equipment Assignments"
    EnsureSheet EQUIPMENT_PROCESSES_SHEET_NAME, "Equipment Process Capabilities"

    currentStep = "EnsureTable capacity masters"
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

    currentStep = "EnsureSheet Parts/PartEditor"
    EnsureSheet PARTS_SHEET_NAME, "Parts Index"
    EnsureSheet PART_EDITOR_SHEET_NAME, "Part Number Editor"
    EnsureSheet PART_DASH_CONDITIONS_SHEET_NAME, "Part Dash Conditions"
    EnsureSheet PART_OPERATIONS_SHEET_NAME, "Part Operations"

    currentStep = "MigrateBasePartsTableToPartsSheet"
    MigrateBasePartsTableToPartsSheet
    currentStep = "ClearPartEditorListObjects"
    ClearPartEditorListObjects

    currentStep = "EnsureTable Parts/Dash/Operations"
    EnsureTable PARTS_SHEET_NAME, BASE_PARTS_TABLE_NAME, Array( _
        COL_BASE_PART_CODE, COL_FACTORY_CODE, COL_ACTIVE, COL_STATUS_DATE, COL_NOTES)

    EnsureTable PART_DASH_CONDITIONS_SHEET_NAME, PART_DASH_CONDITIONS_TABLE_NAME, Array( _
        COL_BASE_PART_CODE, COL_DASH_CONDITION, COL_SEPARATOR, COL_ACTIVE, COL_NOTES)

    EnsureTable PART_OPERATIONS_SHEET_NAME, PART_OPERATIONS_TABLE_NAME, Array( _
        COL_BASE_PART_CODE, COL_OPER_SEQ, COL_OPERATION_NAME, COL_ACTIVE, COL_NOTES)

    currentStep = "EnsureCacheSheet"
    EnsureCacheSheet
    currentStep = "CompactAllCapacityTables"
    CompactAllCapacityTables

    currentStep = "FormatAdminSheet"
    FormatAdminSheet
    currentStep = "FormatPartsSheet"
    FormatPartsSheet

    currentStep = "FormatPartEditorSheet"
    On Error Resume Next
    FormatPartEditorSheet
    If Err.Number <> 0 Then
        formatError = "Error " & CStr(Err.Number) & ": " & Err.Description
        If Len(Trim$(Err.Description)) = 0 Then formatError = formatError & " (no description)"
        Err.Clear
    End If
    On Error GoTo FailBootstrap

    currentStep = "OptimizeExcel False"
    OptimizeExcel False

    If Len(formatError) > 0 Then
        MsgBox "Bootstrap finished, but PartEditor formatting failed:" & vbCrLf & formatError & vbCrLf & vbCrLf & _
            "Run FormatPartEditorLayout after fixing the issue.", vbExclamation
    End If
    Exit Sub

FailBootstrap:
    bootstrapErrorNumber = Err.Number
    bootstrapError = Err.Description
    If Len(Trim$(bootstrapError)) = 0 Then
        bootstrapError = "(no description)"
    End If

    On Error Resume Next
    OptimizeExcel False
    On Error GoTo 0

    MsgBox "BootstrapCapacityTables failed at step:" & vbCrLf & currentStep & vbCrLf & vbCrLf & _
        "Error " & CStr(bootstrapErrorNumber) & ": " & bootstrapError, vbExclamation
End Sub

' Standalone entry point if bootstrap formatting needs a re-run.
Public Sub FormatPartEditorLayout()
    Dim formatErrorNumber As Long
    Dim formatError As String

    On Error GoTo FailFormat
    OptimizeExcel True
    EnsureSheet PART_EDITOR_SHEET_NAME, "Part Number Editor"
    ClearPartEditorListObjects
    FormatPartEditorSheet
    OptimizeExcel False
    MsgBox "PartEditor layout and buttons updated.", vbInformation
    Exit Sub

FailFormat:
    formatErrorNumber = Err.Number
    formatError = Err.Description
    If Len(Trim$(formatError)) = 0 Then formatError = "(no description)"

    On Error Resume Next
    OptimizeExcel False
    On Error GoTo 0

    MsgBox "FormatPartEditorLayout failed:" & vbCrLf & _
        "Error " & CStr(formatErrorNumber) & ": " & formatError, vbExclamation
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
    ws.Range(CACHE_BASE_PART_CELL).Value = vbNullString
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
    ws.Range("A20").Value = "FormatPartEditorLayout"
    ws.Range("A21").Value = "BootstrapCapacityTables"
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

Private Sub ClearPartEditorListObjects()
    Dim ws As Worksheet
    Dim tableName As String

    Set ws = FindWorksheetByName(PART_EDITOR_SHEET_NAME)
    If ws Is Nothing Then Exit Sub

    Do While ws.ListObjects.Count > 0
        tableName = ws.ListObjects(1).Name
        On Error Resume Next
        ws.ListObjects(1).Unlist
        If Err.Number <> 0 Then
            Err.Clear
            ws.ListObjects(1).Delete
        End If
        On Error GoTo 0

        If ws.ListObjects.Count > 0 Then
            If StrComp(ws.ListObjects(1).Name, tableName, vbTextCompare) = 0 Then Exit Do
        End If
    Loop
End Sub

Private Sub FormatPartEditorSheet()
    Dim ws As Worksheet
    Dim avgRange As Range
    Dim dashInputRange As Range
    Dim opsInputRange As Range

    Set ws = FindWorksheetByName(PART_EDITOR_SHEET_NAME)
    If ws Is Nothing Then
        Err.Raise vbObjectError + 540, "FormatPartEditorSheet", "PartEditor sheet was not found."
    End If

    On Error Resume Next
    ws.Unprotect
    ws.Activate
    On Error GoTo 0

    ClearLegacyPartEditorLayout ws

    ' Title
    ws.Range("A1").Value = "Part Number Editor"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Color = RGB(32, 56, 100)
    ws.Range("A2").ClearContents

    ' Identity
    StyleFieldLabel ws, PE_INPUT_ROW, "Part Number"
    StyleInputCell ws.Cells(PE_INPUT_ROW, PE_VALUE_COL), False

    StyleFieldLabel ws, PE_BASE_PART_ROW, "Base Part"
    StyleInputCell ws.Cells(PE_BASE_PART_ROW, PE_VALUE_COL), True

    ' Master section (left)
    StyleSectionHeaderRange ws, PE_MASTER_HEADER_ROW, PE_LABEL_COL, PE_STATUS_VALUE_COL, "Master Record"

    StyleFieldLabel ws, PE_ROW_FACTORY, "Factory"
    StyleInputCell ws.Cells(PE_ROW_FACTORY, PE_VALUE_COL), False

    StyleFieldLabel ws, PE_ROW_ACTIVE, "Active"
    StyleInputCell ws.Cells(PE_ROW_ACTIVE, PE_VALUE_COL), False

    StyleFieldLabel ws, PE_ROW_STATUS_DATE, "Status Date"
    StyleInputCell ws.Cells(PE_ROW_STATUS_DATE, PE_VALUE_COL), False

    StyleFieldLabel ws, PE_ROW_NOTES, "Notes"
    StyleInputCell ws.Cells(PE_ROW_NOTES, PE_VALUE_COL), False

    ' Status under Master Record (F6 label, G6 value)
    ws.Cells(PE_STATUS_ROW, PE_STATUS_LABEL_COL).Value = "Status"
    ws.Cells(PE_STATUS_ROW, PE_STATUS_LABEL_COL).Font.Bold = True
    ws.Cells(PE_STATUS_ROW, PE_STATUS_LABEL_COL).Font.Color = RGB(40, 50, 65)
    ws.Cells(PE_STATUS_ROW, PE_STATUS_LABEL_COL).HorizontalAlignment = xlRight
    StyleInputCell ws.Cells(PE_STATUS_ROW, PE_STATUS_VALUE_COL), True
    ws.Cells(PE_STATUS_ROW, PE_STATUS_VALUE_COL).Font.Italic = True
    ws.Cells(PE_STATUS_ROW, PE_STATUS_VALUE_COL).Font.Color = RGB(60, 60, 60)

    ' Dash conditions (right side starting at J5)
    StyleSectionHeaderRange ws, PE_DASH_SECTION_ROW, PE_COL_DASH, PE_COL_DASH_NOTES, "Dash Conditions"

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

    ' Operations (where dash conditions used to be)
    StyleSectionHeaderRange ws, PE_OPS_SECTION_ROW, PE_LABEL_COL, PE_COL_AVG_EX, "Operations"

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

    Set avgRange = ws.Range( _
        ws.Cells(PE_OPS_DATA_START_ROW, PE_COL_AVG_HOURS), _
        ws.Cells(PE_OPS_DATA_START_ROW + PE_OPS_MAX_ROWS - 1, PE_COL_AVG_EX))
    avgRange.Interior.Color = RGB(235, 238, 242)
    avgRange.Borders.Color = RGB(190, 198, 210)

    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 18
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 10
    ws.Columns("F").ColumnWidth = 10
    ws.Columns("G").ColumnWidth = 28
    ws.Columns("H").ColumnWidth = 12
    ws.Columns("I").ColumnWidth = 3
    ws.Columns("J").ColumnWidth = 14
    ws.Columns("K").ColumnWidth = 10
    ws.Columns("L").ColumnWidth = 8
    ws.Columns("M").ColumnWidth = 18

    FormatDashConditionTextColumn
    EnsurePartEditorButtons ws
End Sub

Private Sub ClearLegacyPartEditorLayout(ByVal ws As Worksheet)
    ' Clear prior instruction/status/dash/ops areas from older layouts.
    ws.Range("A2").ClearContents
    ws.Range("I3:I70").ClearContents
    ws.Range("B11:H70").ClearContents
    ws.Range("B11:H70").Interior.ColorIndex = xlNone
    ws.Range("B11:H70").Borders.LineStyle = xlNone
    ws.Range("J5:M30").ClearContents
    ws.Range("J5:M30").Interior.ColorIndex = xlNone
    ws.Range("J5:M30").Borders.LineStyle = xlNone
    ws.Range("F6:G6").ClearContents
End Sub

Private Sub StyleFieldLabel(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal captionText As String)
    With ws.Cells(rowIndex, PE_LABEL_COL)
        .Value = captionText
        .Font.Bold = True
        .Font.Color = RGB(40, 50, 65)
        .HorizontalAlignment = xlRight
    End With
End Sub

Private Sub StyleInputCell(ByVal cell As Range, ByVal readOnlyLook As Boolean)
    cell.Borders.Color = RGB(160, 170, 185)
    If readOnlyLook Then
        cell.Interior.Color = RGB(235, 238, 242)
    Else
        cell.Interior.Color = RGB(255, 252, 245)
    End If
End Sub

Private Sub StyleSectionHeaderRange( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal startCol As Long, _
    ByVal endCol As Long, _
    ByVal captionText As String)

    Dim headerRange As Range
    Dim cell As Range

    Set headerRange = ws.Range(ws.Cells(rowIndex, startCol), ws.Cells(rowIndex, endCol))
    On Error Resume Next
    headerRange.UnMerge
    On Error GoTo 0

    For Each cell In headerRange.Cells
        cell.Value = vbNullString
        cell.Font.Bold = True
        cell.Font.Size = 11
        cell.Font.Color = RGB(255, 255, 255)
        cell.Interior.Color = RGB(47, 84, 150)
        cell.HorizontalAlignment = xlLeft
        cell.VerticalAlignment = xlCenter
    Next cell

    ws.Cells(rowIndex, startCol).Value = captionText
    ws.Rows(rowIndex).RowHeight = 20
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

    ' Place buttons on row 3 starting at column E (right of the part-number input).
    Set anchor = ws.Cells(PE_BUTTON_ROW, 5)
    buttonTop = anchor.Top
    buttonHeight = 22
    buttonWidth = 85
    gap = 6

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
    On Error Resume Next
    btn.Name = buttonName
    On Error GoTo 0
    btn.Caption = captionText
    btn.OnAction = onActionName
    btn.Font.Bold = True
End Sub

Private Sub DeleteWorksheetButton(ByVal ws As Worksheet, ByVal buttonName As String)
    Dim shp As Shape

    On Error Resume Next
    ws.Buttons(buttonName).Delete
    For Each shp In ws.Shapes
        If StrComp(shp.Name, buttonName, vbTextCompare) = 0 Then shp.Delete
    Next shp
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
