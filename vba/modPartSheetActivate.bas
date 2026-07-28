Attribute VB_Name = "modPartSheetActivate"
Option Explicit

Private Const PART_LABEL_CELL As String = "A1"
Private Const PART_LABEL_VALUE As String = "Part"
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
Private Const DATA_TABLE_COLUMN_WIDTH_PIXELS As Double = 96

' Excel CellControl type for native in-cell checkboxes.
Private Const XL_TYPE_CHECKBOX As Long = 2

' Call from ThisWorkbook.Workbook_SheetActivate:
'   HandlePartSheetActivate Sh
Public Sub HandlePartSheetActivate(ByVal Sh As Object)
    Dim ws As Worksheet

    On Error GoTo CleanUp

    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    Set ws = Sh

    If StrComp(Trim$(CStr(ws.Range(PART_LABEL_CELL).Value)), PART_LABEL_VALUE, vbTextCompare) <> 0 Then
        Exit Sub
    End If

    OptimizeExcel True
    RefreshPartSheetLists ws
    FormatPartDataTable ws

CleanUp:
    OptimizeExcel False
End Sub

Private Sub RefreshPartSheetLists(ByVal ws As Worksheet)
    Dim basePart As String
    Dim ffaValues() As String
    Dim dashConditions() As String
    Dim productLines() As String

    basePart = Trim$(CStr(ws.Range(BASE_PART_CELL).Value))
    If Len(basePart) = 0 Then Exit Sub

    ffaValues = GetReferenceColumnValues("B")
    dashConditions = GetDashConditionsForBasePart(basePart)
    productLines = GetReferenceColumnValues("D")

    SyncValueCheckboxList ws, FFA_VALUE_COLUMN, FFA_CHECKBOX_COLUMN, ffaValues
    SyncValueCheckboxList ws, DASH_VALUE_COLUMN, DASH_CHECKBOX_COLUMN, dashConditions
    SyncValueCheckboxList ws, PRODUCT_LINE_VALUE_COLUMN, PRODUCT_LINE_CHECKBOX_COLUMN, productLines
End Sub

' Writes values/checkboxes/borders only when the list or checkbox set needs updating.
Private Sub SyncValueCheckboxList( _
    ByVal ws As Worksheet, _
    ByVal valueColumn As String, _
    ByVal checkboxColumn As String, _
    ByRef sourceValues() As String)

    Dim currentValues() As String
    Dim sourceCount As Long

    currentValues = ReadColumnList(ws, valueColumn)
    sourceCount = ArrayCount(sourceValues)

    If StringArraysEqual(currentValues, sourceValues) Then
        If CountCellCheckBoxes(ws, checkboxColumn, sourceCount) = sourceCount Then
            Exit Sub
        End If
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
        LastNonEmptyRowFrom(ws, valueColumn, LIST_START_ROW), _
        LastNonEmptyRowFrom(ws, checkboxColumn, LIST_START_ROW), _
        LIST_START_ROW)

    Set clearRange = ws.Range( _
        ws.Cells(LIST_START_ROW, valueColumn), _
        ws.Cells(lastRow, checkboxColumn))
    Set checkboxRange = ws.Range( _
        ws.Cells(LIST_START_ROW, checkboxColumn), _
        ws.Cells(lastRow, checkboxColumn))

    RemoveCellCheckBoxes checkboxRange
    clearRange.ClearContents
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

    ' Write all values first, then apply cell-control checkboxes to the whole block.
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
    Dim cell As Range

    On Error Resume Next
    ' Prefer a bulk clear when available; fall back to clearing cell by cell.
    targetRange.CellControl.Clear
    If Err.Number <> 0 Then
        Err.Clear
        For Each cell In targetRange.Cells
            cell.CellControl.Clear
            Err.Clear
        Next cell
    End If
    On Error GoTo 0
End Sub

Private Function CountCellCheckBoxes(ByVal ws As Worksheet, ByVal checkboxColumn As String, ByVal expectedCount As Long) As Long
    Dim rowIndex As Long
    Dim matchCount As Long
    Dim endRow As Long
    Dim cellType As Variant

    If expectedCount <= 0 Then
        CountCellCheckBoxes = 0
        Exit Function
    End If

    endRow = LIST_START_ROW + expectedCount - 1
    For rowIndex = LIST_START_ROW To endRow
        On Error Resume Next
        cellType = ws.Cells(rowIndex, checkboxColumn).CellControl.Type
        If Err.Number <> 0 Then
            Err.Clear
            CountCellCheckBoxes = 0
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0

        If cellType = XL_TYPE_CHECKBOX Then
            matchCount = matchCount + 1
        End If
    Next rowIndex

    CountCellCheckBoxes = matchCount
End Function

Private Sub ApplyListBorders(ByVal listRange As Range)
    With listRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    listRange.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, ColorIndex:=xlAutomatic
End Sub

' Formats the Part sheet M:Z table: borders on data rows, and fixed
' 96-pixel column widths for M:Z.
Private Sub FormatPartDataTable(ByVal ws As Worksheet)
    FormatPartDataTableBorders ws
    SetPartDataTableColumnWidths ws
End Sub

Private Sub FormatPartDataTableBorders(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim tableRange As Range

    lastRow = LastUsedRowInRangeColumns(ws, DATA_TABLE_FIRST_COLUMN, DATA_TABLE_LAST_COLUMN, LIST_START_ROW)
    If lastRow < LIST_START_ROW Then Exit Sub

    Set tableRange = ws.Range( _
        ws.Cells(LIST_START_ROW, DATA_TABLE_FIRST_COLUMN), _
        ws.Cells(lastRow, DATA_TABLE_LAST_COLUMN))

    ApplyListBorders tableRange
End Sub

Private Sub SetPartDataTableColumnWidths(ByVal ws As Worksheet)
    Dim colIndex As Long
    Dim firstColIndex As Long
    Dim lastColIndex As Long
    Dim headerRange As Range

    firstColIndex = ws.Columns(DATA_TABLE_FIRST_COLUMN).Column
    lastColIndex = ws.Columns(DATA_TABLE_LAST_COLUMN).Column

    For colIndex = firstColIndex To lastColIndex
        SetColumnWidthPixels ws.Columns(colIndex), DATA_TABLE_COLUMN_WIDTH_PIXELS
    Next colIndex

    Set headerRange = ws.Range( _
        ws.Cells(DATA_TABLE_HEADER_ROW, DATA_TABLE_FIRST_COLUMN), _
        ws.Cells(DATA_TABLE_HEADER_ROW, DATA_TABLE_LAST_COLUMN))
    headerRange.WrapText = True
    ws.Rows(DATA_TABLE_HEADER_ROW).AutoFit
End Sub

' Sets a column's width so its rendered width is the requested pixel size
' at 96 DPI (Excel's Width property is in points).
Private Sub SetColumnWidthPixels(ByVal columnRange As Range, ByVal pixelWidth As Double)
    Dim targetPoints As Double

    targetPoints = pixelWidth * 72# / 96#
    columnRange.ColumnWidth = 10
    If columnRange.Width <> 0 Then
        columnRange.ColumnWidth = columnRange.ColumnWidth * (targetPoints / columnRange.Width)
    End If
End Sub

Private Function LastUsedRowInRangeColumns( _
    ByVal ws As Worksheet, _
    ByVal firstColumn As String, _
    ByVal lastColumn As String, _
    ByVal startRow As Long) As Long

    Dim colIndex As Long
    Dim firstColIndex As Long
    Dim lastColIndex As Long
    Dim columnLastRow As Long
    Dim maxRow As Long

    firstColIndex = ws.Columns(firstColumn).Column
    lastColIndex = ws.Columns(lastColumn).Column
    maxRow = startRow - 1

    For colIndex = firstColIndex To lastColIndex
        columnLastRow = LastNonEmptyRowFrom(ws, ColumnLetter(colIndex), startRow)
        If columnLastRow > maxRow Then maxRow = columnLastRow
    Next colIndex

    LastUsedRowInRangeColumns = maxRow
End Function

Private Function ColumnLetter(ByVal columnIndex As Long) As String
    Dim dividend As Long
    Dim modulo As Long
    Dim result As String

    dividend = columnIndex
    Do
        modulo = (dividend - 1) Mod 26
        result = Chr$(65 + modulo) & result
        dividend = (dividend - modulo) \ 26
    Loop While dividend > 0

    ColumnLetter = result
End Function

Private Function ReadColumnList(ByVal ws As Worksheet, ByVal columnLetter As String) As String()
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim values() As String
    Dim valueCount As Long
    Dim cellValue As String

    lastRow = LastNonEmptyRowFrom(ws, columnLetter, LIST_START_ROW)
    If lastRow < LIST_START_ROW Then
        ReadColumnList = EmptyStringArray()
        Exit Function
    End If

    valueCount = 0
    ReDim values(0 To 0)

    For rowIndex = LIST_START_ROW To lastRow
        cellValue = Trim$(CStr(ws.Cells(rowIndex, columnLetter).Value))
        If Len(cellValue) > 0 Then
            ReDim Preserve values(0 To valueCount)
            values(valueCount) = cellValue
            valueCount = valueCount + 1
        Else
            Exit For
        End If
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
    Dim rowIndex As Long
    Dim cellValue As String
    Dim uniqueValues As Object

    Set wsReferences = ThisWorkbook.Worksheets(REFERENCES_SHEET_NAME)
    Set uniqueValues = CreateObject("Scripting.Dictionary")
    uniqueValues.CompareMode = vbTextCompare

    lastRow = LastUsedRowInColumn(wsReferences, columnLetter)
    For rowIndex = 2 To lastRow
        cellValue = Trim$(CStr(wsReferences.Cells(rowIndex, columnLetter).Value))
        If Len(cellValue) > 0 Then
            If Not uniqueValues.Exists(cellValue) Then
                uniqueValues.Add cellValue, cellValue
            End If
        End If
    Next rowIndex

    GetReferenceColumnValues = DictionaryKeysToSortedArray(uniqueValues)
End Function

Private Function GetDashConditionsForBasePart(ByVal basePart As String) As String()
    Dim wsStandards As Worksheet
    Dim tbl As ListObject
    Dim colAssembly As ListColumn
    Dim rowIndex As Long
    Dim assemblyVal As String
    Dim rowBasePart As String
    Dim dashCondition As String
    Dim uniqueValues As Object

    Set wsStandards = ThisWorkbook.Worksheets(ASSEMBLY_STANDARDS_SHEET_NAME)
    Set tbl = wsStandards.ListObjects(ASSY_STANDARDS_TABLE_NAME)
    Set colAssembly = tbl.ListColumns(COL_ASSEMBLY_NO)
    Set uniqueValues = CreateObject("Scripting.Dictionary")
    uniqueValues.CompareMode = vbTextCompare

    If tbl.DataBodyRange Is Nothing Then
        GetDashConditionsForBasePart = EmptyStringArray()
        Exit Function
    End If

    For rowIndex = 1 To tbl.ListRows.Count
        assemblyVal = Trim$(CStr(colAssembly.DataBodyRange.Cells(rowIndex, 1).Value))
        If Len(assemblyVal) > 0 Then
            SplitAssemblyNo assemblyVal, rowBasePart, dashCondition
            If StrComp(rowBasePart, basePart, vbTextCompare) = 0 And Len(dashCondition) > 0 Then
                If Not uniqueValues.Exists(dashCondition) Then
                    uniqueValues.Add dashCondition, dashCondition
                End If
            End If
        End If
    Next rowIndex

    GetDashConditionsForBasePart = DictionaryKeysToSortedArray(uniqueValues)
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

Private Function LastNonEmptyRowFrom(ByVal ws As Worksheet, ByVal columnLetter As String, ByVal startRow As Long) As Long
    Dim rowIndex As Long
    Dim lastRow As Long

    lastRow = LastUsedRowInColumn(ws, columnLetter)
    If lastRow < startRow Then
        LastNonEmptyRowFrom = startRow - 1
        Exit Function
    End If

    For rowIndex = lastRow To startRow Step -1
        If Len(Trim$(CStr(ws.Cells(rowIndex, columnLetter).Value))) > 0 Then
            LastNonEmptyRowFrom = rowIndex
            Exit Function
        End If
    Next rowIndex

    LastNonEmptyRowFrom = startRow - 1
End Function

Private Function LastUsedRowInColumn(ByVal ws As Worksheet, ByVal columnLetter As String) As Long
    Dim foundCell As Range
    Dim endUpRow As Long

    endUpRow = ws.Cells(ws.Rows.Count, columnLetter).End(xlUp).Row

    Set foundCell = ws.Columns(columnLetter).Find( _
        What:="*", _
        LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious)

    If foundCell Is Nothing Then
        LastUsedRowInColumn = endUpRow
    ElseIf foundCell.Row > endUpRow Then
        LastUsedRowInColumn = foundCell.Row
    Else
        LastUsedRowInColumn = endUpRow
    End If
End Function
