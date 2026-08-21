Attribute VB_Name = "modPartSheetActivate"
Option Explicit

Private Const PART_LABEL_CELL As String = "A1"
Private Const PART_LABEL_VALUE As String = "Part"
' Opaque activate cache (hidden via ;;;). Bump PART_CACHE_SCHEMA when
' activate formatting/formula behavior changes so stamps invalidate.
Private Const PART_CACHE_CELL As String = "A2"
Private Const PART_CACHE_SCHEMA As String = "1"
Private Const BASE_PART_CELL As String = "C2"

Private Const REFERENCES_SHEET_NAME As String = "References"
Private Const ASSEMBLY_STANDARDS_SHEET_NAME As String = "Assembly Standards"
Private Const ASSY_STANDARDS_TABLE_NAME As String = "AssyStndTbl"
Private Const COL_ASSEMBLY_NO As String = "ASSEMBLY NO"

Private Const LIST_START_ROW As Long = 9
Private Const FFA_VALUE_COLUMN As String = "C"
Private Const FFA_CHECKBOX_COLUMN As String = "D"
Private Const DASH_VALUE_COLUMN As String = "E"
Private Const DASH_CHECKBOX_COLUMN As String = "F"
Private Const PRODUCT_LINE_VALUE_COLUMN As String = "G"
Private Const PRODUCT_LINE_CHECKBOX_COLUMN As String = "H"

Private Const DATA_TABLE_FIRST_COLUMN As String = "M"
Private Const DATA_TABLE_LAST_COLUMN As String = "Z"
Private Const DATA_TABLE_HEADER_ROW As Long = 8
Private Const DATA_TABLE_INPUT_LAST_COLUMN As String = "S"
Private Const DATA_TABLE_WIDTH_N_TO_Q_PIXELS As Double = 68
Private Const DATA_TABLE_WIDTH_W_TO_Z_PIXELS As Double = 71
Private Const DATA_TABLE_WIDTH_SELECTED_PIXELS As Double = 90
Private Const DATA_TABLE_WIDTH_DEFAULT_PIXELS As Double = 96
Private Const PIXELS_TO_POINTS As Double = 72# / 96#

' Excel CellControl type for native in-cell checkboxes.
Private Const XL_TYPE_CHECKBOX As Long = 2

' Session memoization for source lists (cleared on VBA reset).
Private g_refBSig As String
Private g_refBValues() As String
Private g_refBCached As Boolean
Private g_refDSig As String
Private g_refDValues() As String
Private g_refDCached As Boolean
Private g_dashCacheBase As String
Private g_dashCacheSig As String
Private g_dashCacheValues() As String
Private g_dashCached As Boolean

' Call from ThisWorkbook.Workbook_SheetActivate:
'   HandlePartSheetActivate Sh
Public Sub HandlePartSheetActivate(ByVal Sh As Object)
    Dim ws As Worksheet
    Dim basePart As String
    Dim ffaValues() As String
    Dim dashConditions() As String
    Dim productLines() As String
    Dim currentSig As String
    Dim storedSig As String

    On Error GoTo CleanUp

    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    Set ws = Sh

    If Not IsPartSheet(ws) Then Exit Sub

    basePart = Trim$(CStr(ws.Range(BASE_PART_CELL).Value))
    If Len(basePart) = 0 Then Exit Sub

    EnsureDataSheet
    ClearLegacyPartCacheCells ws

    ffaValues = ReferenceColumnValues("B")
    dashConditions = DashConditionsForBasePart(basePart)
    productLines = ReferenceColumnValues("D")
    currentSig = BuildListSignature(basePart)
    storedSig = PartListSig(ws)

    If StrComp(storedSig, currentSig, vbBinaryCompare) = 0 Then
        If SpotCheckListCheckBoxes(ws, FFA_CHECKBOX_COLUMN, ArrayCount(ffaValues)) Then
            If SpotCheckListCheckBoxes(ws, DASH_CHECKBOX_COLUMN, ArrayCount(dashConditions)) Then
                If SpotCheckListCheckBoxes(ws, PRODUCT_LINE_CHECKBOX_COLUMN, ArrayCount(productLines)) Then
                    EnsurePartOpsTable ws
                    ActiveWindow.DisplayGridlines = False
                    Exit Sub
                End If
            End If
        End If
    End If

    OptimizeExcel True
    SyncValueCheckboxList ws, FFA_VALUE_COLUMN, FFA_CHECKBOX_COLUMN, ffaValues
    SyncValueCheckboxList ws, DASH_VALUE_COLUMN, DASH_CHECKBOX_COLUMN, dashConditions
    SyncValueCheckboxList ws, PRODUCT_LINE_VALUE_COLUMN, PRODUCT_LINE_CHECKBOX_COLUMN, productLines
    EnsurePartOpsTable ws
    FormatPartDataTable ws
    SyncPartToStore ws
    ActiveWindow.DisplayGridlines = False

CleanUp:
    OptimizeExcel False
End Sub

Private Sub ClearLegacyPartCacheCells(ByVal ws As Worksheet)
    On Error Resume Next
    If ws.Range("A2").NumberFormat = ";;;" Then
        ws.Range("A2").ClearContents
        ws.Range("A2").NumberFormat = "General"
    End If
    If ws.Range("A3").NumberFormat = ";;;" Then
        ws.Range("A3").ClearContents
        ws.Range("A3").NumberFormat = "General"
    End If
    On Error GoTo 0
End Sub

Private Function BuildPartActivateCacheKey( _
    ByVal basePart As String, _
    ByRef ffaValues() As String, _
    ByRef dashConditions() As String, _
    ByRef productLines() As String, _
    ByVal dataLastRow As Long) As String

    BuildPartActivateCacheKey = _
        PART_CACHE_SCHEMA & Chr$(31) & _
        UCase$(basePart) & Chr$(31) & _
        JoinStringArray(ffaValues) & Chr$(31) & _
        JoinStringArray(dashConditions) & Chr$(31) & _
        JoinStringArray(productLines) & Chr$(31) & _
        CStr(dataLastRow)
End Function

Private Function PartActivateCacheIsCurrent( _
    ByVal ws As Worksheet, _
    ByVal cacheKey As String, _
    ByRef ffaValues() As String, _
    ByRef dashConditions() As String, _
    ByRef productLines() As String, _
    ByVal dataLastRow As Long) As Boolean

    If StrComp(CStr(Nz(ws.Range(PART_CACHE_CELL).Value2)), cacheKey, vbBinaryCompare) <> 0 Then
        Exit Function
    End If

    If Not SpotCheckListCheckBoxes(ws, FFA_CHECKBOX_COLUMN, ArrayCount(ffaValues)) Then Exit Function
    If Not SpotCheckListCheckBoxes(ws, DASH_CHECKBOX_COLUMN, ArrayCount(dashConditions)) Then Exit Function
    If Not SpotCheckListCheckBoxes(ws, PRODUCT_LINE_CHECKBOX_COLUMN, ArrayCount(productLines)) Then Exit Function

    If dataLastRow >= LIST_START_ROW Then
        If Len(ws.Cells(LIST_START_ROW, "W").Formula) = 0 Then Exit Function
        If Len(ws.Cells(LIST_START_ROW, "X").Formula) = 0 Then Exit Function
        If Len(ws.Cells(LIST_START_ROW, "Y").Formula) = 0 Then Exit Function
        If Len(ws.Cells(dataLastRow, "Y").Formula) = 0 Then Exit Function
    End If

    PartActivateCacheIsCurrent = True
End Function

Private Function SpotCheckListCheckBoxes( _
    ByVal ws As Worksheet, _
    ByVal checkboxColumn As String, _
    ByVal itemCount As Long) As Boolean

    Dim endRow As Long

    If itemCount <= 0 Then
        SpotCheckListCheckBoxes = True
        Exit Function
    End If

    endRow = LIST_START_ROW + itemCount - 1
    If Not CellIsCheckBox(ws.Cells(LIST_START_ROW, checkboxColumn)) Then Exit Function
    If endRow > LIST_START_ROW Then
        If Not CellIsCheckBox(ws.Cells(endRow, checkboxColumn)) Then Exit Function
    End If

    SpotCheckListCheckBoxes = True
End Function

Private Function CellIsCheckBox(ByVal cell As Range) As Boolean
    Dim cellType As Variant

    On Error Resume Next
    cellType = cell.CellControl.Type
    If Err.Number <> 0 Then
        Err.Clear
        CellIsCheckBox = False
        Exit Function
    End If
    On Error GoTo 0

    CellIsCheckBox = (cellType = XL_TYPE_CHECKBOX)
End Function

Private Sub WritePartActivateCache(ByVal ws As Worksheet, ByVal cacheKey As String)
    With ws.Range(PART_CACHE_CELL)
        .NumberFormat = ";;;"
        .Value2 = cacheKey
    End With
End Sub

Private Function JoinStringArray(ByRef values() As String) As String
    Dim i As Long
    Dim parts() As String
    Dim count As Long

    count = ArrayCount(values)
    If count = 0 Then
        JoinStringArray = vbNullString
        Exit Function
    End If

    ReDim parts(0 To count - 1)
    For i = 0 To count - 1
        parts(i) = values(LBound(values) + i)
    Next i

    JoinStringArray = Join(parts, Chr$(30))
End Function

Private Function CloneStringArray(ByRef values() As String) As String()
    Dim result() As String
    Dim i As Long
    Dim count As Long

    count = ArrayCount(values)
    If count = 0 Then
        CloneStringArray = EmptyStringArray()
        Exit Function
    End If

    ReDim result(LBound(values) To UBound(values))
    For i = LBound(values) To UBound(values)
        result(i) = values(i)
    Next i

    CloneStringArray = result
End Function

' Writes values/checkboxes/borders only when the list or checkbox set needs updating.
Private Sub SyncValueCheckboxList( _
    ByVal ws As Worksheet, _
    ByVal valueColumn As String, _
    ByVal checkboxColumn As String, _
    ByRef sourceValues() As String)

    Dim currentValues() As String
    Dim sourceCount As Long
    Dim endRow As Long
    Dim checkboxRange As Range

    currentValues = ReadColumnList(ws, valueColumn)
    sourceCount = ArrayCount(sourceValues)

    If StringArraysEqual(currentValues, sourceValues) Then
        If sourceCount = 0 Then Exit Sub

        endRow = LIST_START_ROW + sourceCount - 1
        Set checkboxRange = ws.Range( _
            ws.Cells(LIST_START_ROW, checkboxColumn), _
            ws.Cells(endRow, checkboxColumn))

        If RangeIsAllCheckBoxes(checkboxRange) Then Exit Sub
    End If

    ClearValueCheckboxList ws, valueColumn, checkboxColumn

    If sourceCount = 0 Then Exit Sub

    PopulateValueCheckboxList ws, valueColumn, checkboxColumn, sourceValues
End Sub

Private Sub ClearValueCheckboxList( _
    ByVal ws As Worksheet, _
    ByVal valueColumn As String, _
    ByVal checkboxColumn As String)

    Dim lastRow As Long
    Dim clearRange As Range
    Dim checkboxRange As Range

    lastRow = Application.WorksheetFunction.Max( _
        FastLastUsedRowInColumn(ws, valueColumn), _
        FastLastUsedRowInColumn(ws, checkboxColumn), _
        LIST_START_ROW)

    If lastRow < LIST_START_ROW Then lastRow = LIST_START_ROW

    Set clearRange = ws.Range( _
        ws.Cells(LIST_START_ROW, valueColumn), _
        ws.Cells(lastRow, checkboxColumn))
    Set checkboxRange = ws.Range( _
        ws.Cells(LIST_START_ROW, checkboxColumn), _
        ws.Cells(lastRow, checkboxColumn))

    RemoveCellCheckBoxes checkboxRange
    clearRange.ClearContents
    clearRange.Interior.ColorIndex = xlNone
    clearRange.Borders.LineStyle = xlNone
End Sub

Private Sub PopulateValueCheckboxList( _
    ByVal ws As Worksheet, _
    ByVal valueColumn As String, _
    ByVal checkboxColumn As String, _
    ByRef values() As String)

    Dim i As Long
    Dim sourceCount As Long
    Dim endRow As Long
    Dim valueRange As Range
    Dim checkboxRange As Range
    Dim listRange As Range
    Dim outputValues As Variant
    Dim rowOffset As Long

    sourceCount = ArrayCount(values)
    If sourceCount = 0 Then Exit Sub

    endRow = LIST_START_ROW + sourceCount - 1
    Set valueRange = ws.Range(ws.Cells(LIST_START_ROW, valueColumn), ws.Cells(endRow, valueColumn))
    Set checkboxRange = ws.Range(ws.Cells(LIST_START_ROW, checkboxColumn), ws.Cells(endRow, checkboxColumn))
    Set listRange = ws.Range(ws.Cells(LIST_START_ROW, valueColumn), ws.Cells(endRow, checkboxColumn))

    ReDim outputValues(1 To sourceCount, 1 To 1)
    rowOffset = 1
    For i = LBound(values) To UBound(values)
        outputValues(rowOffset, 1) = values(i)
        rowOffset = rowOffset + 1
    Next i

    valueRange.Value = outputValues
    valueRange.HorizontalAlignment = xlLeft
    valueRange.VerticalAlignment = xlCenter

    checkboxRange.HorizontalAlignment = xlCenter
    checkboxRange.VerticalAlignment = xlCenter
    checkboxRange.Value = False
    ApplyCellCheckBoxes checkboxRange

    ApplyListBorders listRange
End Sub

Private Sub ApplyCellCheckBoxes(ByVal targetRange As Range)
    On Error GoTo FailSetCheckbox
    targetRange.CellControl.SetCheckbox
    Exit Sub

FailSetCheckbox:
    Err.Raise vbObjectError + 700, "ApplyCellCheckBoxes", _
        "Unable to add cell-control checkboxes. This requires Excel for Microsoft 365 with Insert > Checkbox support."
End Sub

Private Sub RemoveCellCheckBoxes(ByVal targetRange As Range)
    On Error Resume Next
    targetRange.CellControl.Clear
    On Error GoTo 0
End Sub

' Bulk check: CellControl.Type returns xlTypeCheckbox only when every cell matches.
Private Function RangeIsAllCheckBoxes(ByVal checkboxRange As Range) As Boolean
    Dim cellType As Variant

    On Error Resume Next
    cellType = checkboxRange.CellControl.Type
    If Err.Number <> 0 Then
        Err.Clear
        RangeIsAllCheckBoxes = False
        Exit Function
    End If
    On Error GoTo 0

    RangeIsAllCheckBoxes = (cellType = XL_TYPE_CHECKBOX)
End Function

Private Sub ApplyListBorders(ByVal listRange As Range)
    With listRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    listRange.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, ColorIndex:=xlAutomatic
End Sub

' Formats the Part sheet M:Z table: borders, column widths, U/V checkboxes,
' W:Y formulas, alignment, fills, and number formats.
Private Sub FormatPartDataTable(ByVal ws As Worksheet)
    EnsurePartOpsTable ws
    FormatPartDataTableBorders ws
    SetPartDataTableColumnWidths ws
    EnsureDataTableCheckBoxesAndFormulas ws
    ApplyPartDataTableStyles ws
End Sub

Private Sub FormatPartDataTableBorders(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim tableRange As Range

    lastRow = FastLastUsedRowInColumns(ws, DATA_TABLE_FIRST_COLUMN, DATA_TABLE_LAST_COLUMN)
    If lastRow < LIST_START_ROW Then Exit Sub

    Set tableRange = ws.Range( _
        ws.Cells(LIST_START_ROW, DATA_TABLE_FIRST_COLUMN), _
        ws.Cells(lastRow, DATA_TABLE_LAST_COLUMN))

    ' Skip expensive border redraw when the block already has an outer border.
    If tableRange.Borders(xlEdgeLeft).LineStyle <> xlNone Then
        If tableRange.Borders(xlEdgeLeft).Weight = xlMedium Then Exit Sub
    End If

    ApplyListBorders tableRange
End Sub

Private Sub SetPartDataTableColumnWidths(ByVal ws As Worksheet)
    Dim headerRange As Range

    If PartDataTableWidthsAreSet(ws) Then Exit Sub

    SetColumnWidthPixels ws.Range("N:Q"), DATA_TABLE_WIDTH_N_TO_Q_PIXELS
    SetColumnWidthPixels ws.Range("W:Z"), DATA_TABLE_WIDTH_W_TO_Z_PIXELS
    SetColumnWidthPixels ws.Range("M:M"), DATA_TABLE_WIDTH_SELECTED_PIXELS
    SetColumnWidthPixels ws.Range("R:S"), DATA_TABLE_WIDTH_SELECTED_PIXELS
    SetColumnWidthPixels ws.Range("U:V"), DATA_TABLE_WIDTH_SELECTED_PIXELS
    SetColumnWidthPixels ws.Range("T:T"), DATA_TABLE_WIDTH_DEFAULT_PIXELS

    Set headerRange = ws.Range( _
        ws.Cells(DATA_TABLE_HEADER_ROW, DATA_TABLE_FIRST_COLUMN), _
        ws.Cells(DATA_TABLE_HEADER_ROW, DATA_TABLE_LAST_COLUMN))
    headerRange.WrapText = True
    ws.Rows(DATA_TABLE_HEADER_ROW).AutoFit
End Sub

Private Function PartDataTableWidthsAreSet(ByVal ws As Worksheet) As Boolean
    Const WIDTH_TOLERANCE_POINTS As Double = 0.75

    If Abs(ws.Columns("N").Width - DATA_TABLE_WIDTH_N_TO_Q_PIXELS * PIXELS_TO_POINTS) >= WIDTH_TOLERANCE_POINTS Then Exit Function
    If Abs(ws.Columns("W").Width - DATA_TABLE_WIDTH_W_TO_Z_PIXELS * PIXELS_TO_POINTS) >= WIDTH_TOLERANCE_POINTS Then Exit Function
    If Abs(ws.Columns("M").Width - DATA_TABLE_WIDTH_SELECTED_PIXELS * PIXELS_TO_POINTS) >= WIDTH_TOLERANCE_POINTS Then Exit Function
    If Abs(ws.Columns("R").Width - DATA_TABLE_WIDTH_SELECTED_PIXELS * PIXELS_TO_POINTS) >= WIDTH_TOLERANCE_POINTS Then Exit Function
    If Abs(ws.Columns("U").Width - DATA_TABLE_WIDTH_SELECTED_PIXELS * PIXELS_TO_POINTS) >= WIDTH_TOLERANCE_POINTS Then Exit Function
    If Abs(ws.Columns("T").Width - DATA_TABLE_WIDTH_DEFAULT_PIXELS * PIXELS_TO_POINTS) >= WIDTH_TOLERANCE_POINTS Then Exit Function

    PartDataTableWidthsAreSet = True
End Function

' Sets column width so rendered width matches the requested pixel size at 96 DPI.
Private Sub SetColumnWidthPixels(ByVal columnRange As Range, ByVal pixelWidth As Double)
    Dim targetPoints As Double
    Dim sampleCol As Range

    targetPoints = pixelWidth * PIXELS_TO_POINTS
    Set sampleCol = columnRange.Columns(1)
    sampleCol.ColumnWidth = 10
    If sampleCol.Width = 0 Then Exit Sub

    columnRange.ColumnWidth = sampleCol.ColumnWidth * (targetPoints / sampleCol.Width)
End Sub

' Adds cell-control checkboxes in U:V and selection formulas in W:Y for each
' data row. Existing checkbox values are preserved.
Private Sub EnsureDataTableCheckBoxesAndFormulas(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim rowCount As Long
    Dim rowIndex As Long
    Dim checkboxRange As Range
    Dim savedValues As Variant
    Dim formulaW As Variant
    Dim formulaX As Variant
    Dim formulaY As Variant
    Dim expectedW As String
    Dim expectedX As String
    Dim expectedY As String
    Dim rowText As String

    lastRow = FastLastUsedRowInColumns(ws, DATA_TABLE_FIRST_COLUMN, DATA_TABLE_INPUT_LAST_COLUMN)
    If lastRow < LIST_START_ROW Then Exit Sub

    rowCount = lastRow - LIST_START_ROW + 1
    Set checkboxRange = ws.Range( _
        ws.Cells(LIST_START_ROW, "U"), _
        ws.Cells(lastRow, "V"))

    If Not RangeIsAllCheckBoxes(checkboxRange) Then
        savedValues = checkboxRange.Value2
        NormalizeCheckboxSeedValues savedValues
        checkboxRange.HorizontalAlignment = xlCenter
        checkboxRange.VerticalAlignment = xlCenter
        checkboxRange.Value = savedValues
        ApplyCellCheckBoxes checkboxRange
        checkboxRange.Value = savedValues
    End If

    rowText = CStr(LIST_START_ROW)
    expectedW = "=IF(AND(R" & rowText & "<>"""",U" & rowText & "=TRUE),R" & rowText & ",O" & rowText & ")"
    expectedX = "=IF(AND(S" & rowText & "<>"""",V" & rowText & "=TRUE),S" & rowText & ",P" & rowText & ")"
    expectedY = "=IF(OR(Q" & rowText & "="""",W" & rowText & "="""",X" & rowText & "=""""),"""",(W" & rowText & "*X" & rowText & ")/Q" & rowText & ")"

    If StrComp(ws.Cells(LIST_START_ROW, "W").Formula, expectedW, vbTextCompare) = 0 _
        And StrComp(ws.Cells(LIST_START_ROW, "X").Formula, expectedX, vbTextCompare) = 0 _
        And StrComp(ws.Cells(LIST_START_ROW, "Y").Formula, expectedY, vbTextCompare) = 0 _
        And Len(ws.Cells(lastRow, "W").Formula) > 0 _
        And Len(ws.Cells(lastRow, "X").Formula) > 0 _
        And Len(ws.Cells(lastRow, "Y").Formula) > 0 Then
        Exit Sub
    End If

    ReDim formulaW(1 To rowCount, 1 To 1)
    ReDim formulaX(1 To rowCount, 1 To 1)
    ReDim formulaY(1 To rowCount, 1 To 1)

    For rowIndex = LIST_START_ROW To lastRow
        rowText = CStr(rowIndex)
        formulaW(rowIndex - LIST_START_ROW + 1, 1) = _
            "=IF(AND(R" & rowText & "<>"""",U" & rowText & "=TRUE),R" & rowText & ",O" & rowText & ")"
        formulaX(rowIndex - LIST_START_ROW + 1, 1) = _
            "=IF(AND(S" & rowText & "<>"""",V" & rowText & "=TRUE),S" & rowText & ",P" & rowText & ")"
        formulaY(rowIndex - LIST_START_ROW + 1, 1) = _
            "=IF(OR(Q" & rowText & "="""",W" & rowText & "="""",X" & rowText & "=""""),"""",(W" & rowText & "*X" & rowText & ")/Q" & rowText & ")"
    Next rowIndex

    ws.Range(ws.Cells(LIST_START_ROW, "W"), ws.Cells(lastRow, "W")).Formula = formulaW
    ws.Range(ws.Cells(LIST_START_ROW, "X"), ws.Cells(lastRow, "X")).Formula = formulaX
    ws.Range(ws.Cells(LIST_START_ROW, "Y"), ws.Cells(lastRow, "Y")).Formula = formulaY
End Sub

Private Sub NormalizeCheckboxSeedValues(ByRef values As Variant)
    Dim rowIndex As Long
    Dim colIndex As Long

    If Not IsArray(values) Then
        values = CoerceCheckboxValue(values)
        Exit Sub
    End If

    For rowIndex = LBound(values, 1) To UBound(values, 1)
        For colIndex = LBound(values, 2) To UBound(values, 2)
            values(rowIndex, colIndex) = CoerceCheckboxValue(values(rowIndex, colIndex))
        Next colIndex
    Next rowIndex
End Sub

Private Function CoerceCheckboxValue(ByVal rawValue As Variant) As Boolean
    If IsError(rawValue) Then Exit Function
    If IsEmpty(rawValue) Or IsNull(rawValue) Then Exit Function
    If Len(Trim$(CStr(rawValue))) = 0 Then Exit Function

    On Error Resume Next
    CoerceCheckboxValue = CBool(rawValue)
    On Error GoTo 0
End Function

Private Sub ApplyPartDataTableStyles(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim tableRange As Range
    Dim headerRange As Range

    lastRow = FastLastUsedRowInColumns(ws, DATA_TABLE_FIRST_COLUMN, DATA_TABLE_LAST_COLUMN)
    If lastRow < LIST_START_ROW Then lastRow = LIST_START_ROW

    Set headerRange = ws.Range( _
        ws.Cells(DATA_TABLE_HEADER_ROW, DATA_TABLE_FIRST_COLUMN), _
        ws.Cells(DATA_TABLE_HEADER_ROW, DATA_TABLE_LAST_COLUMN))
    headerRange.Interior.ColorIndex = xlNone
    headerRange.HorizontalAlignment = xlCenter
    headerRange.VerticalAlignment = xlCenter

    Set tableRange = ws.Range( _
        ws.Cells(LIST_START_ROW, DATA_TABLE_FIRST_COLUMN), _
        ws.Cells(lastRow, DATA_TABLE_LAST_COLUMN))

    tableRange.HorizontalAlignment = xlCenter
    tableRange.VerticalAlignment = xlCenter

    ' Clear previous M:N highlighting if present, then apply the current fills.
    ws.Range(ws.Cells(LIST_START_ROW, "M"), ws.Cells(lastRow, "N")).Interior.ColorIndex = xlNone

    ws.Range(ws.Cells(LIST_START_ROW, "O"), ws.Cells(lastRow, "Q")).Interior.Color = RGB(213, 229, 249)
    ws.Range(ws.Cells(LIST_START_ROW, "T"), ws.Cells(lastRow, "T")).Interior.Color = RGB(213, 229, 249)
    ws.Range(ws.Cells(LIST_START_ROW, "U"), ws.Cells(lastRow, "V")).Interior.Color = RGB(213, 229, 249)
    ws.Range(ws.Cells(LIST_START_ROW, "Z"), ws.Cells(lastRow, "Z")).Interior.Color = RGB(213, 229, 249)

    ws.Range(ws.Cells(LIST_START_ROW, "O"), ws.Cells(lastRow, "S")).NumberFormat = "0.00"
    ws.Range(ws.Cells(LIST_START_ROW, "W"), ws.Cells(lastRow, "Y")).NumberFormat = "0.00"

    ' D/F/H lists can differ in length — clear shared leftover fills, then
    ' highlight only cells that actually have checkbox controls.
    ClearCheckboxColumnHighlights ws
    HighlightCheckboxColumn ws, FFA_VALUE_COLUMN, FFA_CHECKBOX_COLUMN
    HighlightCheckboxColumn ws, DASH_VALUE_COLUMN, DASH_CHECKBOX_COLUMN
    HighlightCheckboxColumn ws, PRODUCT_LINE_VALUE_COLUMN, PRODUCT_LINE_CHECKBOX_COLUMN
End Sub

' Clears D/F/H fills through the farthest list row so shorter columns do not
' keep highlight past their checkboxes.
Private Sub ClearCheckboxColumnHighlights(ByVal ws As Worksheet)
    Dim clearToRow As Long

    clearToRow = Application.WorksheetFunction.Max( _
        FastLastUsedRowInColumn(ws, FFA_VALUE_COLUMN), _
        FastLastUsedRowInColumn(ws, DASH_VALUE_COLUMN), _
        FastLastUsedRowInColumn(ws, PRODUCT_LINE_VALUE_COLUMN), _
        FastLastUsedRowInColumn(ws, FFA_CHECKBOX_COLUMN), _
        FastLastUsedRowInColumn(ws, DASH_CHECKBOX_COLUMN), _
        FastLastUsedRowInColumn(ws, PRODUCT_LINE_CHECKBOX_COLUMN), _
        LIST_START_ROW)

    If clearToRow < LIST_START_ROW Then Exit Sub

    ws.Range( _
        ws.Cells(LIST_START_ROW, FFA_CHECKBOX_COLUMN), _
        ws.Cells(clearToRow, FFA_CHECKBOX_COLUMN)).Interior.ColorIndex = xlNone
    ws.Range( _
        ws.Cells(LIST_START_ROW, DASH_CHECKBOX_COLUMN), _
        ws.Cells(clearToRow, DASH_CHECKBOX_COLUMN)).Interior.ColorIndex = xlNone
    ws.Range( _
        ws.Cells(LIST_START_ROW, PRODUCT_LINE_CHECKBOX_COLUMN), _
        ws.Cells(clearToRow, PRODUCT_LINE_CHECKBOX_COLUMN)).Interior.ColorIndex = xlNone
End Sub

Private Sub HighlightCheckboxColumn(ByVal ws As Worksheet, ByVal valueColumn As String, ByVal checkboxColumn As String)
    Dim lastRow As Long
    Dim checkboxRange As Range

    lastRow = FastLastUsedRowInColumn(ws, valueColumn)
    If lastRow < LIST_START_ROW Then Exit Sub

    Set checkboxRange = ws.Range( _
        ws.Cells(LIST_START_ROW, checkboxColumn), _
        ws.Cells(lastRow, checkboxColumn))

    If RangeIsAllCheckBoxes(checkboxRange) Then
        checkboxRange.Interior.Color = RGB(213, 229, 249)
    Else
        HighlightCheckboxCellsInRange checkboxRange
    End If
End Sub

Private Sub HighlightCheckboxCellsInRange(ByVal checkboxRange As Range)
    Dim cell As Range
    Dim cellType As Variant

    For Each cell In checkboxRange.Cells
        On Error Resume Next
        cellType = cell.CellControl.Type
        If Err.Number = 0 Then
            If cellType = XL_TYPE_CHECKBOX Then
                cell.Interior.Color = RGB(213, 229, 249)
            Else
                cell.Interior.ColorIndex = xlNone
            End If
        Else
            Err.Clear
            cell.Interior.ColorIndex = xlNone
        End If
        On Error GoTo 0
    Next cell
End Sub

Private Function ReadColumnList(ByVal ws As Worksheet, ByVal columnLetter As String) As String()
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim values() As String
    Dim valueCount As Long
    Dim rawValues As Variant
    Dim cellValue As String

    lastRow = FastLastUsedRowInColumn(ws, columnLetter)
    If lastRow < LIST_START_ROW Then
        ReadColumnList = EmptyStringArray()
        Exit Function
    End If

    rawValues = ws.Range(ws.Cells(LIST_START_ROW, columnLetter), ws.Cells(lastRow, columnLetter)).Value2
    valueCount = 0
    ReDim values(0 To 0)

    If Not IsArray(rawValues) Then
        cellValue = Trim$(CStr(Nz(rawValues)))
        If Len(cellValue) > 0 Then
            ReDim values(0 To 0)
            values(0) = cellValue
            ReadColumnList = values
        Else
            ReadColumnList = EmptyStringArray()
        End If
        Exit Function
    End If

    For rowIndex = 1 To UBound(rawValues, 1)
        cellValue = Trim$(CStr(Nz(rawValues(rowIndex, 1))))
        If Len(cellValue) = 0 Then Exit For

        ReDim Preserve values(0 To valueCount)
        values(valueCount) = cellValue
        valueCount = valueCount + 1
    Next rowIndex

    If valueCount = 0 Then
        ReadColumnList = EmptyStringArray()
    Else
        ReadColumnList = values
    End If
End Function

Private Function GetReferenceColumnValues(ByVal columnLetter As String) As String()
    Dim wsReferences As Worksheet
    Dim lastRow As Long
    Dim rawValues As Variant
    Dim sourceSig As String
    Dim values() As String

    Set wsReferences = ThisWorkbook.Worksheets(REFERENCES_SHEET_NAME)
    lastRow = FastLastUsedRowInColumn(wsReferences, columnLetter)

    If lastRow < 2 Then
        sourceSig = columnLetter & Chr$(31) & "0"
        rawValues = Empty
    Else
        rawValues = wsReferences.Range( _
            wsReferences.Cells(2, columnLetter), _
            wsReferences.Cells(lastRow, columnLetter)).Value2
        sourceSig = columnLetter & Chr$(31) & CStr(lastRow) & Chr$(31) & HashVariantColumn(rawValues)
    End If

    If StrComp(columnLetter, "B", vbTextCompare) = 0 Then
        If g_refBCached And StrComp(g_refBSig, sourceSig, vbBinaryCompare) = 0 Then
            GetReferenceColumnValues = CloneStringArray(g_refBValues)
            Exit Function
        End If
    ElseIf StrComp(columnLetter, "D", vbTextCompare) = 0 Then
        If g_refDCached And StrComp(g_refDSig, sourceSig, vbBinaryCompare) = 0 Then
            GetReferenceColumnValues = CloneStringArray(g_refDValues)
            Exit Function
        End If
    End If

    values = UniqueSortedValuesFromColumn(rawValues)

    If StrComp(columnLetter, "B", vbTextCompare) = 0 Then
        g_refBSig = sourceSig
        g_refBValues = CloneStringArray(values)
        g_refBCached = True
    ElseIf StrComp(columnLetter, "D", vbTextCompare) = 0 Then
        g_refDSig = sourceSig
        g_refDValues = CloneStringArray(values)
        g_refDCached = True
    End If

    GetReferenceColumnValues = values
End Function

Private Function UniqueSortedValuesFromColumn(ByVal rawValues As Variant) As String()
    Dim rowIndex As Long
    Dim cellValue As String
    Dim uniqueValues As Object

    Set uniqueValues = CreateObject("Scripting.Dictionary")
    uniqueValues.CompareMode = vbTextCompare

    If IsEmpty(rawValues) Then
        UniqueSortedValuesFromColumn = EmptyStringArray()
        Exit Function
    End If

    If Not IsArray(rawValues) Then
        cellValue = Trim$(CStr(Nz(rawValues)))
        If Len(cellValue) > 0 Then uniqueValues.Add cellValue, cellValue
        UniqueSortedValuesFromColumn = DictionaryKeysToSortedArray(uniqueValues)
        Exit Function
    End If

    For rowIndex = 1 To UBound(rawValues, 1)
        cellValue = Trim$(CStr(Nz(rawValues(rowIndex, 1))))
        If Len(cellValue) > 0 Then
            If Not uniqueValues.Exists(cellValue) Then
                uniqueValues.Add cellValue, cellValue
            End If
        End If
    Next rowIndex

    UniqueSortedValuesFromColumn = DictionaryKeysToSortedArray(uniqueValues)
End Function

Private Function GetDashConditionsForBasePart(ByVal basePart As String) As String()
    Dim wsStandards As Worksheet
    Dim tbl As ListObject
    Dim assemblyValues As Variant
    Dim sourceSig As String
    Dim values() As String
    Dim rowCount As Long

    Set wsStandards = ThisWorkbook.Worksheets(ASSEMBLY_STANDARDS_SHEET_NAME)
    Set tbl = wsStandards.ListObjects(ASSY_STANDARDS_TABLE_NAME)

    If tbl.DataBodyRange Is Nothing Then
        sourceSig = "0"
        assemblyValues = Empty
    Else
        rowCount = tbl.DataBodyRange.Rows.Count
        assemblyValues = ColumnValues(tbl.ListColumns(COL_ASSEMBLY_NO))
        sourceSig = CStr(rowCount) & Chr$(31) & HashVariantColumn(assemblyValues)
    End If

    If g_dashCached _
        And StrComp(g_dashCacheBase, basePart, vbTextCompare) = 0 _
        And StrComp(g_dashCacheSig, sourceSig, vbBinaryCompare) = 0 Then
        GetDashConditionsForBasePart = CloneStringArray(g_dashCacheValues)
        Exit Function
    End If

    values = ExtractDashConditions(assemblyValues, basePart)
    g_dashCacheBase = basePart
    g_dashCacheSig = sourceSig
    g_dashCacheValues = CloneStringArray(values)
    g_dashCached = True
    GetDashConditionsForBasePart = values
End Function

Private Function ExtractDashConditions(ByVal assemblyValues As Variant, ByVal basePart As String) As String()
    Dim rowIndex As Long
    Dim assemblyVal As String
    Dim rowBasePart As String
    Dim dashCondition As String
    Dim uniqueValues As Object

    Set uniqueValues = CreateObject("Scripting.Dictionary")
    uniqueValues.CompareMode = vbTextCompare

    If IsEmpty(assemblyValues) Then
        ExtractDashConditions = EmptyStringArray()
        Exit Function
    End If

    If Not IsArray(assemblyValues) Then
        assemblyVal = Trim$(CStr(Nz(assemblyValues)))
        If Len(assemblyVal) > 0 Then
            SplitAssemblyNo assemblyVal, rowBasePart, dashCondition
            If StrComp(rowBasePart, basePart, vbTextCompare) = 0 And Len(dashCondition) > 0 Then
                uniqueValues.Add dashCondition, dashCondition
            End If
        End If
        ExtractDashConditions = DictionaryKeysToSortedArray(uniqueValues)
        Exit Function
    End If

    For rowIndex = 1 To UBound(assemblyValues, 1)
        assemblyVal = Trim$(CStr(Nz(assemblyValues(rowIndex, 1))))
        If Len(assemblyVal) > 0 Then
            SplitAssemblyNo assemblyVal, rowBasePart, dashCondition
            If StrComp(rowBasePart, basePart, vbTextCompare) = 0 And Len(dashCondition) > 0 Then
                If Not uniqueValues.Exists(dashCondition) Then
                    uniqueValues.Add dashCondition, dashCondition
                End If
            End If
        End If
    Next rowIndex

    ExtractDashConditions = DictionaryKeysToSortedArray(uniqueValues)
End Function

' Lightweight content fingerprint — avoids building giant joined strings.
Private Function HashVariantColumn(ByVal values As Variant) As String
    Dim rowIndex As Long
    Dim cellValue As String
    Dim totalLen As Long
    Dim checkSum As Long
    Dim count As Long
    Dim firstValue As String
    Dim lastValue As String

    If IsEmpty(values) Then
        HashVariantColumn = "0:0:0:"
        Exit Function
    End If

    If Not IsArray(values) Then
        cellValue = Trim$(CStr(Nz(values)))
        HashVariantColumn = "1:" & CStr(Len(cellValue)) & ":" & CStr(Len(cellValue)) & ":" & cellValue
        Exit Function
    End If

    count = UBound(values, 1)
    For rowIndex = 1 To count
        cellValue = Trim$(CStr(Nz(values(rowIndex, 1))))
        totalLen = totalLen + Len(cellValue)
        If Len(cellValue) > 0 Then
            checkSum = (checkSum + AscW(Left$(cellValue, 1)) + AscW(Right$(cellValue, 1)) + Len(cellValue)) Mod 2147483647
            If Len(firstValue) = 0 Then firstValue = cellValue
            lastValue = cellValue
        End If
    Next rowIndex

    HashVariantColumn = CStr(count) & ":" & CStr(totalLen) & ":" & CStr(checkSum) & ":" & firstValue & Chr$(30) & lastValue
End Function

Private Function ColumnValues(ByVal col As ListColumn) As Variant
    Dim values As Variant
    Dim result(1 To 1, 1 To 1) As Variant

    values = col.DataBodyRange.Value2

    If IsArray(values) Then
        ColumnValues = values
    Else
        result(1, 1) = values
        ColumnValues = result
    End If
End Function

Private Function DictionaryKeysToSortedArray(ByVal valueMap As Object) As String()
    Dim keys() As String
    Dim key As Variant
    Dim index As Long

    If valueMap.Count = 0 Then
        DictionaryKeysToSortedArray = EmptyStringArray()
        Exit Function
    End If

    ReDim keys(0 To valueMap.Count - 1)
    index = 0
    For Each key In valueMap.Keys
        keys(index) = CStr(key)
        index = index + 1
    Next key

    SortStringArray keys
    DictionaryKeysToSortedArray = keys
End Function

Private Function StringArraysEqual(ByRef leftValues() As String, ByRef rightValues() As String) As Boolean
    Dim i As Long
    Dim leftCount As Long
    Dim rightCount As Long

    leftCount = ArrayCount(leftValues)
    rightCount = ArrayCount(rightValues)

    If leftCount <> rightCount Then Exit Function

    For i = 0 To leftCount - 1
        If StrComp(leftValues(i), rightValues(i), vbTextCompare) <> 0 Then Exit Function
    Next i

    StringArraysEqual = True
End Function

Private Function ArrayCount(ByRef values() As String) As Long
    If Not IsArrayInitialized(values) Then
        ArrayCount = 0
    Else
        ArrayCount = UBound(values) - LBound(values) + 1
    End If
End Function

Private Function EmptyStringArray() As String()
    Dim emptyKeys() As String
    EmptyStringArray = emptyKeys
End Function

Private Function IsArrayInitialized(ByRef values() As String) As Boolean
    Dim upperBound As Long

    On Error Resume Next
    upperBound = UBound(values)
    IsArrayInitialized = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Sub SplitAssemblyNo(ByVal assemblyNo As String, ByRef basePart As String, ByRef dashCondition As String)
    Dim dashPos As Long

    dashPos = InStr(1, assemblyNo, "-", vbBinaryCompare)

    If dashPos > 0 Then
        basePart = Trim$(Left$(assemblyNo, dashPos - 1))
        dashCondition = Trim$(Mid$(assemblyNo, dashPos + 1))
    Else
        basePart = assemblyNo
        dashCondition = vbNullString
    End If
End Sub

Private Sub SortStringArray(ByRef keys() As String)
    Dim i As Long
    Dim j As Long
    Dim tempKey As String

    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If CompareListValues(keys(i), keys(j)) > 0 Then
                tempKey = keys(i)
                keys(i) = keys(j)
                keys(j) = tempKey
            End If
        Next j
    Next i
End Sub

Private Function CompareListValues(ByVal leftValue As String, ByVal rightValue As String) As Long
    If IsNumeric(leftValue) And IsNumeric(rightValue) Then
        CompareListValues = Sgn(CDbl(leftValue) - CDbl(rightValue))
    Else
        CompareListValues = StrComp(leftValue, rightValue, vbTextCompare)
    End If
End Function

Private Function FastLastUsedRowInColumn(ByVal ws As Worksheet, ByVal columnLetter As String) As Long
    FastLastUsedRowInColumn = ws.Cells(ws.Rows.Count, columnLetter).End(xlUp).Row
End Function

Private Function FastLastUsedRowInColumns( _
    ByVal ws As Worksheet, _
    ByVal firstColumn As String, _
    ByVal lastColumn As String) As Long

    Dim colIndex As Long
    Dim firstColIndex As Long
    Dim lastColIndex As Long
    Dim columnLastRow As Long
    Dim maxRow As Long

    firstColIndex = ws.Columns(firstColumn).Column
    lastColIndex = ws.Columns(lastColumn).Column
    maxRow = 0

    For colIndex = firstColIndex To lastColIndex
        columnLastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row
        If columnLastRow > maxRow Then maxRow = columnLastRow
    Next colIndex

    FastLastUsedRowInColumns = maxRow
End Function

Private Function Nz(ByVal value As Variant) As Variant
    If IsError(value) Then
        Nz = vbNullString
    ElseIf IsEmpty(value) Or IsNull(value) Then
        Nz = vbNullString
    Else
        Nz = value
    End If
End Function
