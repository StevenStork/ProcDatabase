Attribute VB_Name = "modExportSheets"
Option Explicit

' Builds one export sheet per FFA (References!B) and per product line
' (References!D). Export sheets are plain data for copying to another file —
' no formatting is applied.
'
' Column A holds export metadata (marker / type / key).
' Data table starts at column C:
'   Part Number | Op Sequence | Op Code | Process Hours | Avg Ex |
'   Batch Size  | Avg HPU     | Equipment Type
'
' Part Number is each Part sheet's base part (C2), with any dash condition
' stripped. Remaining fields come from that sheet's data table
' (M, N, W, X, Q, Y, T). FFA sheets only include rows whose column Z matches
' the FFA. Product-line sheets include every data row from Part sheets that
' have that product line checked (G/H).

Private Const PART_LABEL_CELL As String = "A1"
Private Const PART_LABEL_VALUE As String = "Part"
Private Const PART_NUMBER_CELL As String = "C2"

Private Const REFERENCES_SHEET_NAME As String = "References"
Private Const REFERENCES_FFA_COLUMN As String = "B"
Private Const REFERENCES_PRODUCT_LINE_COLUMN As String = "D"
Private Const REFERENCES_START_ROW As Long = 2

Private Const LIST_START_ROW As Long = 9
Private Const PRODUCT_LINE_VALUE_COLUMN As String = "G"
Private Const PRODUCT_LINE_CHECKBOX_COLUMN As String = "H"

Private Const DATA_FIRST_COLUMN As String = "M"
Private Const DATA_LAST_COLUMN As String = "Z"
Private Const COL_OP_SEQUENCE As String = "M"
Private Const COL_OP_CODE As String = "N"
Private Const COL_PROCESS_HOURS As String = "W"
Private Const COL_AVG_EX As String = "X"
Private Const COL_BATCH_SIZE As String = "Q"
Private Const COL_AVG_HPU As String = "Y"
Private Const COL_EQUIPMENT_TYPE As String = "T"
Private Const COL_FFA_MARK As String = "Z"

Private Const EXPORT_MARKER_CELL As String = "A1"
Private Const EXPORT_MARKER_VALUE As String = "Export"
Private Const EXPORT_TYPE_CELL As String = "A2"
Private Const EXPORT_KEY_CELL As String = "A3"
Private Const EXPORT_TYPE_FFA As String = "FFA"
Private Const EXPORT_TYPE_PRODUCT_LINE As String = "ProductLine"
' Legacy marker cells from the first export layout; cleared on rebuild.
Private Const LEGACY_EXPORT_MARKER_CELL As String = "AA1"
Private Const LEGACY_EXPORT_TYPE_CELL As String = "AA2"
Private Const LEGACY_EXPORT_KEY_CELL As String = "AA3"

Private Const FFA_SHEET_PREFIX As String = "FFA - "
Private Const PRODUCT_LINE_SHEET_PREFIX As String = "PL - "
Private Const EXPORT_HEADER_ROW As Long = 1
Private Const EXPORT_FIRST_DATA_ROW As Long = 2
Private Const EXPORT_FIRST_COLUMN As Long = 3   ' Column C
Private Const EXPORT_LAST_COLUMN As Long = 10  ' Column J
Private Const EXPORT_COLUMN_COUNT As Long = 8

Public Sub BuildExportSheets()
    Dim ffaValues() As String
    Dim productLines() As String
    Dim ffaRows As Object
    Dim productLineRows As Object
    Dim i As Long
    Dim keyName As String

    On Error GoTo CleanUp
    OptimizeExcel True

    ' Ensure W/X/Y formula results are current before exporting values.
    Application.Calculate

    ffaValues = GetReferenceColumnValues(REFERENCES_FFA_COLUMN)
    productLines = GetReferenceColumnValues(REFERENCES_PRODUCT_LINE_COLUMN)

    Set ffaRows = CreateObject("Scripting.Dictionary")
    ffaRows.CompareMode = vbTextCompare
    Set productLineRows = CreateObject("Scripting.Dictionary")
    productLineRows.CompareMode = vbTextCompare

    For i = 0 To ArrayCount(ffaValues) - 1
        keyName = ffaValues(LBound(ffaValues) + i)
        If Len(keyName) > 0 Then
            If Not ffaRows.Exists(keyName) Then
                ffaRows.Add keyName, New Collection
            End If
        End If
    Next i

    For i = 0 To ArrayCount(productLines) - 1
        keyName = productLines(LBound(productLines) + i)
        If Len(keyName) > 0 Then
            If Not productLineRows.Exists(keyName) Then
                productLineRows.Add keyName, New Collection
            End If
        End If
    Next i

    CollectExportRowsFromPartSheets ffaRows, productLineRows
    RemoveObsoleteExportSheets ffaRows, productLineRows

    For i = 0 To ArrayCount(ffaValues) - 1
        keyName = ffaValues(LBound(ffaValues) + i)
        If Len(keyName) > 0 Then
            WriteExportSheet _
                FFA_SHEET_PREFIX, _
                keyName, _
                EXPORT_TYPE_FFA, _
                ffaRows(keyName)
        End If
    Next i

    For i = 0 To ArrayCount(productLines) - 1
        keyName = productLines(LBound(productLines) + i)
        If Len(keyName) > 0 Then
            WriteExportSheet _
                PRODUCT_LINE_SHEET_PREFIX, _
                keyName, _
                EXPORT_TYPE_PRODUCT_LINE, _
                productLineRows(keyName)
        End If
    Next i

CleanUp:
    OptimizeExcel False
End Sub

Private Sub CollectExportRowsFromPartSheets(ByVal ffaRows As Object, ByVal productLineRows As Object)
    Dim ws As Worksheet
    Dim partNumber As String
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim ffaMark As String
    Dim markedProductLines As Collection
    Dim plIndex As Long
    Dim plName As String
    Dim exportRow As Variant
        Dim tableValues As Variant

    For Each ws In ThisWorkbook.Worksheets
        If Not IsPartSheet(ws) Then GoTo NextSheet

        partNumber = BasePartWithoutDash(CStr(Nz(ws.Range(PART_NUMBER_CELL).Value)))
        If Len(partNumber) = 0 Then partNumber = BasePartWithoutDash(ws.Name)
        lastRow = LastUsedRowInColumns(ws, DATA_FIRST_COLUMN, DATA_LAST_COLUMN)
        If lastRow < LIST_START_ROW Then GoTo NextSheet

        Set markedProductLines = MarkedProductLinesOnSheet(ws)

        tableValues = ws.Range( _
            ws.Cells(LIST_START_ROW, DATA_FIRST_COLUMN), _
            ws.Cells(lastRow, DATA_LAST_COLUMN)).Value2

        For rowIndex = 1 To UBound(tableValues, 1)
            If Not RowHasExportableData(tableValues, rowIndex) Then GoTo NextDataRow

            exportRow = BuildExportRow(partNumber, tableValues, rowIndex)

            ffaMark = Trim$(CStr(Nz(tableValues(rowIndex, ColumnOffset(COL_FFA_MARK)))))
            If Len(ffaMark) > 0 Then
                If ffaRows.Exists(ffaMark) Then
                    ffaRows(ffaMark).Add exportRow
                End If
            End If

            If Not markedProductLines Is Nothing Then
                For plIndex = 1 To markedProductLines.Count
                    plName = CStr(markedProductLines(plIndex))
                    If productLineRows.Exists(plName) Then
                        productLineRows(plName).Add exportRow
                    End If
                Next plIndex
            End If
NextDataRow:
        Next rowIndex
NextSheet:
    Next ws
End Sub

Private Function MarkedProductLinesOnSheet(ByVal ws As Worksheet) As Collection
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim lineName As String
    Dim marked As Collection

    Set marked = New Collection
    lastRow = LastUsedRowInColumn(ws, PRODUCT_LINE_VALUE_COLUMN)
    If lastRow < LIST_START_ROW Then
        Set MarkedProductLinesOnSheet = marked
        Exit Function
    End If

    For rowIndex = LIST_START_ROW To lastRow
        lineName = Trim$(CStr(Nz(ws.Cells(rowIndex, PRODUCT_LINE_VALUE_COLUMN).Value)))
        If Len(lineName) = 0 Then Exit For

        If IsActiveFlag(ws.Cells(rowIndex, PRODUCT_LINE_CHECKBOX_COLUMN).Value) Then
            marked.Add lineName
        End If
    Next rowIndex

    Set MarkedProductLinesOnSheet = marked
End Function

Private Function RowHasExportableData(ByVal tableValues As Variant, ByVal rowIndex As Long) As Boolean
    If Len(Trim$(CStr(Nz(tableValues(rowIndex, ColumnOffset(COL_OP_SEQUENCE)))))) > 0 Then
        RowHasExportableData = True
        Exit Function
    End If
    If Len(Trim$(CStr(Nz(tableValues(rowIndex, ColumnOffset(COL_OP_CODE)))))) > 0 Then
        RowHasExportableData = True
        Exit Function
    End If
    If Len(Trim$(CStr(Nz(tableValues(rowIndex, ColumnOffset(COL_FFA_MARK)))))) > 0 Then
        RowHasExportableData = True
    End If
End Function

Private Function BuildExportRow( _
    ByVal partNumber As String, _
    ByVal tableValues As Variant, _
    ByVal rowIndex As Long) As Variant

    Dim exportRow(1 To EXPORT_COLUMN_COUNT) As Variant

    exportRow(1) = partNumber
    exportRow(2) = Nz(tableValues(rowIndex, ColumnOffset(COL_OP_SEQUENCE)))
    exportRow(3) = Nz(tableValues(rowIndex, ColumnOffset(COL_OP_CODE)))
    exportRow(4) = Nz(tableValues(rowIndex, ColumnOffset(COL_PROCESS_HOURS)))
    exportRow(5) = Nz(tableValues(rowIndex, ColumnOffset(COL_AVG_EX)))
    exportRow(6) = Nz(tableValues(rowIndex, ColumnOffset(COL_BATCH_SIZE)))
    exportRow(7) = Nz(tableValues(rowIndex, ColumnOffset(COL_AVG_HPU)))
    exportRow(8) = Nz(tableValues(rowIndex, ColumnOffset(COL_EQUIPMENT_TYPE)))

    BuildExportRow = exportRow
End Function

Private Function ColumnOffset(ByVal columnLetter As String) As Long
    ' M:Z block starts at column M (=1 in the dumped array).
    ColumnOffset = ColumnIndexFromLetter(columnLetter) - ColumnIndexFromLetter(DATA_FIRST_COLUMN) + 1
End Function

Private Function ColumnIndexFromLetter(ByVal columnLetter As String) As Long
    Dim letters As String
    Dim i As Long
    Dim result As Long

    letters = UCase$(Trim$(columnLetter))
    result = 0
    For i = 1 To Len(letters)
        result = result * 26 + (Asc(Mid$(letters, i, 1)) - Asc("A") + 1)
    Next i

    ColumnIndexFromLetter = result
End Function

Private Sub WriteExportSheet( _
    ByVal namePrefix As String, _
    ByVal keyName As String, _
    ByVal exportType As String, _
    ByVal rows As Collection)

    Dim ws As Worksheet
    Dim output() As Variant
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim sourceRow As Variant
    Dim targetRange As Range

    Set ws = GetOrCreateExportSheet(namePrefix, keyName, exportType)

    ClearExportDataArea ws

    ws.Cells(EXPORT_HEADER_ROW, EXPORT_FIRST_COLUMN).Value = "Part Number"
    ws.Cells(EXPORT_HEADER_ROW, EXPORT_FIRST_COLUMN + 1).Value = "Op Sequence"
    ws.Cells(EXPORT_HEADER_ROW, EXPORT_FIRST_COLUMN + 2).Value = "Op Code"
    ws.Cells(EXPORT_HEADER_ROW, EXPORT_FIRST_COLUMN + 3).Value = "Process Hours"
    ws.Cells(EXPORT_HEADER_ROW, EXPORT_FIRST_COLUMN + 4).Value = "Avg Ex"
    ws.Cells(EXPORT_HEADER_ROW, EXPORT_FIRST_COLUMN + 5).Value = "Batch Size"
    ws.Cells(EXPORT_HEADER_ROW, EXPORT_FIRST_COLUMN + 6).Value = "Avg HPU"
    ws.Cells(EXPORT_HEADER_ROW, EXPORT_FIRST_COLUMN + 7).Value = "Equipment Type"

    If rows Is Nothing Then Exit Sub
    If rows.Count = 0 Then Exit Sub

    ReDim output(1 To rows.Count, 1 To EXPORT_COLUMN_COUNT)
    For rowIndex = 1 To rows.Count
        sourceRow = rows(rowIndex)
        For colIndex = 1 To EXPORT_COLUMN_COUNT
            output(rowIndex, colIndex) = sourceRow(colIndex)
        Next colIndex
    Next rowIndex

    Set targetRange = ws.Range( _
        ws.Cells(EXPORT_FIRST_DATA_ROW, EXPORT_FIRST_COLUMN), _
        ws.Cells(EXPORT_FIRST_DATA_ROW + rows.Count - 1, EXPORT_LAST_COLUMN))
    targetRange.Value2 = output
End Sub

Private Sub ClearExportDataArea(ByVal ws As Worksheet)
    Dim lastRow As Long

    ' Clear only the data table (C:J). Column A holds export metadata.
    lastRow = LastUsedRowInColumns(ws, "C", "J")
    If lastRow < EXPORT_HEADER_ROW Then lastRow = EXPORT_HEADER_ROW

    ws.Range( _
        ws.Cells(EXPORT_HEADER_ROW, EXPORT_FIRST_COLUMN), _
        ws.Cells(lastRow, EXPORT_LAST_COLUMN)).ClearContents

    ' Also clear any leftover table that was previously written in A:H.
    If Len(Trim$(CStr(Nz(ws.Cells(EXPORT_HEADER_ROW, 1).Value)))) > 0 Then
        If StrComp(Trim$(CStr(Nz(ws.Cells(EXPORT_HEADER_ROW, 1).Value))), EXPORT_MARKER_VALUE, vbTextCompare) <> 0 Then
            ws.Range(ws.Cells(EXPORT_HEADER_ROW, 1), ws.Cells(lastRow, 8)).ClearContents
        End If
    End If
End Sub

Private Function GetOrCreateExportSheet( _
    ByVal namePrefix As String, _
    ByVal keyName As String, _
    ByVal exportType As String) As Worksheet

    Dim ws As Worksheet
    Dim sheetName As String

    Set ws = FindExportSheet(exportType, keyName)
    sheetName = BuildExportSheetName(namePrefix, keyName)

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        On Error Resume Next
        ws.Name = sheetName
        If Err.Number <> 0 Then
            Err.Clear
            ws.Name = UniqueSheetName(sheetName)
        End If
        On Error GoTo 0
    ElseIf StrComp(ws.Name, sheetName, vbTextCompare) <> 0 Then
        On Error Resume Next
        ws.Name = sheetName
        If Err.Number <> 0 Then
            Err.Clear
            ws.Name = UniqueSheetName(sheetName)
        End If
        On Error GoTo 0
    End If

    ws.Range(EXPORT_MARKER_CELL).Value = EXPORT_MARKER_VALUE
    ws.Range(EXPORT_TYPE_CELL).Value = exportType
    ws.Range(EXPORT_KEY_CELL).Value = keyName
    ws.Range(LEGACY_EXPORT_MARKER_CELL).ClearContents
    ws.Range(LEGACY_EXPORT_TYPE_CELL).ClearContents
    ws.Range(LEGACY_EXPORT_KEY_CELL).ClearContents

    Set GetOrCreateExportSheet = ws
End Function

Private Function FindExportSheet(ByVal exportType As String, ByVal keyName As String) As Worksheet
    Dim ws As Worksheet
    Dim sheetType As String
    Dim sheetKey As String

    For Each ws In ThisWorkbook.Worksheets
        If IsExportSheet(ws) Then
            sheetType = ExportSheetType(ws)
            sheetKey = ExportSheetKey(ws)
            If StrComp(sheetType, exportType, vbTextCompare) = 0 Then
                If StrComp(sheetKey, keyName, vbTextCompare) = 0 Then
                    Set FindExportSheet = ws
                    Exit Function
                End If
            End If
        End If
    Next ws
End Function

Private Sub RemoveObsoleteExportSheets(ByVal ffaRows As Object, ByVal productLineRows As Object)
    Dim ws As Worksheet
    Dim sheetNames As Collection
    Dim sheetName As Variant
    Dim exportType As String
    Dim keyName As String
    Dim keepSheet As Boolean

    Set sheetNames = New Collection

    For Each ws In ThisWorkbook.Worksheets
        If IsExportSheet(ws) Then
            exportType = ExportSheetType(ws)
            keyName = ExportSheetKey(ws)
            keepSheet = False

            If StrComp(exportType, EXPORT_TYPE_FFA, vbTextCompare) = 0 Then
                keepSheet = ffaRows.Exists(keyName)
            ElseIf StrComp(exportType, EXPORT_TYPE_PRODUCT_LINE, vbTextCompare) = 0 Then
                keepSheet = productLineRows.Exists(keyName)
            End If

            If Not keepSheet Then
                sheetNames.Add ws.Name
            End If
        End If
    Next ws

    For Each sheetName In sheetNames
        ThisWorkbook.Worksheets(CStr(sheetName)).Delete
    Next sheetName
End Sub

Private Function IsExportSheet(ByVal ws As Worksheet) As Boolean
    If StrComp(Trim$(CStr(Nz(ws.Range(EXPORT_MARKER_CELL).Value))), EXPORT_MARKER_VALUE, vbTextCompare) = 0 Then
        IsExportSheet = True
        Exit Function
    End If

    ' Recognize sheets still marked in the legacy AA location.
    IsExportSheet = (StrComp(Trim$(CStr(Nz(ws.Range(LEGACY_EXPORT_MARKER_CELL).Value))), EXPORT_MARKER_VALUE, vbTextCompare) = 0)
End Function

Private Function ExportSheetType(ByVal ws As Worksheet) As String
    ExportSheetType = Trim$(CStr(Nz(ws.Range(EXPORT_TYPE_CELL).Value)))
    If Len(ExportSheetType) = 0 Then
        ExportSheetType = Trim$(CStr(Nz(ws.Range(LEGACY_EXPORT_TYPE_CELL).Value)))
    End If
End Function

Private Function ExportSheetKey(ByVal ws As Worksheet) As String
    ExportSheetKey = Trim$(CStr(Nz(ws.Range(EXPORT_KEY_CELL).Value)))
    If Len(ExportSheetKey) = 0 Then
        ExportSheetKey = Trim$(CStr(Nz(ws.Range(LEGACY_EXPORT_KEY_CELL).Value)))
    End If
End Function

Private Function IsPartSheet(ByVal ws As Worksheet) As Boolean
    IsPartSheet = (StrComp(Trim$(CStr(Nz(ws.Range(PART_LABEL_CELL).Value))), PART_LABEL_VALUE, vbTextCompare) = 0)
End Function

' Base part only — strip a trailing dash condition (e.g. "12345-AB" -> "12345").
Private Function BasePartWithoutDash(ByVal rawValue As String) As String
    Dim trimmed As String
    Dim dashPos As Long

    trimmed = Trim$(rawValue)
    If Len(trimmed) = 0 Then Exit Function

    dashPos = InStr(1, trimmed, "-", vbBinaryCompare)
    If dashPos > 1 Then
        BasePartWithoutDash = Trim$(Left$(trimmed, dashPos - 1))
    Else
        BasePartWithoutDash = trimmed
    End If
End Function

Private Function BuildExportSheetName(ByVal namePrefix As String, ByVal keyName As String) As String
    Dim cleanedKey As String
    Dim maxKeyLen As Long

    cleanedKey = SanitizeSheetNameFragment(keyName)
    maxKeyLen = 31 - Len(namePrefix)
    If maxKeyLen < 1 Then maxKeyLen = 1
    If Len(cleanedKey) > maxKeyLen Then
        cleanedKey = Left$(cleanedKey, maxKeyLen)
    End If

    BuildExportSheetName = SanitizeSheetNameFragment(namePrefix & cleanedKey)
End Function

Private Function UniqueSheetName(ByVal proposedName As String) As String
    Dim candidate As String
    Dim suffix As Long

    candidate = Left$(proposedName, 31)
    suffix = 1

    Do While SheetExists(candidate)
        candidate = Left$(proposedName, 31 - Len(CStr(suffix)) - 1) & "_" & CStr(suffix)
        suffix = suffix + 1
        If suffix > 9999 Then
            Err.Raise vbObjectError + 710, "UniqueSheetName", "Unable to invent a unique export sheet name."
        End If
    Loop

    UniqueSheetName = candidate
End Function

Private Function SanitizeSheetNameFragment(ByVal proposedName As String) As String
    Dim cleaned As String

    cleaned = proposedName
    cleaned = Replace$(cleaned, "\", "_")
    cleaned = Replace$(cleaned, "/", "_")
    cleaned = Replace$(cleaned, "?", "_")
    cleaned = Replace$(cleaned, "*", "_")
    cleaned = Replace$(cleaned, "[", "_")
    cleaned = Replace$(cleaned, "]", "_")
    cleaned = Replace$(cleaned, ":", "_")

    SanitizeSheetNameFragment = Trim$(cleaned)
End Function

Private Function GetReferenceColumnValues(ByVal columnLetter As String) As String()
    Dim wsReferences As Worksheet
    Dim lastRow As Long
    Dim rawValues As Variant
    Dim rowIndex As Long
    Dim cellValue As String
    Dim uniqueValues As Object

    Set uniqueValues = CreateObject("Scripting.Dictionary")
    uniqueValues.CompareMode = vbTextCompare

    Set wsReferences = ThisWorkbook.Worksheets(REFERENCES_SHEET_NAME)
    lastRow = LastUsedRowInColumn(wsReferences, columnLetter)

    If lastRow < REFERENCES_START_ROW Then
        GetReferenceColumnValues = EmptyStringArray()
        Exit Function
    End If

    rawValues = wsReferences.Range( _
        wsReferences.Cells(REFERENCES_START_ROW, columnLetter), _
        wsReferences.Cells(lastRow, columnLetter)).Value2

    If Not IsArray(rawValues) Then
        cellValue = Trim$(CStr(Nz(rawValues)))
        If Len(cellValue) > 0 Then uniqueValues.Add cellValue, cellValue
        GetReferenceColumnValues = DictionaryKeysToSortedArray(uniqueValues)
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

    GetReferenceColumnValues = DictionaryKeysToSortedArray(uniqueValues)
End Function

Private Function DictionaryKeysToSortedArray(ByVal dict As Object) As String()
    Dim keys() As String
    Dim key As Variant
    Dim i As Long

    If dict Is Nothing Then
        DictionaryKeysToSortedArray = EmptyStringArray()
        Exit Function
    End If

    If dict.Count = 0 Then
        DictionaryKeysToSortedArray = EmptyStringArray()
        Exit Function
    End If

    ReDim keys(0 To dict.Count - 1)
    i = 0
    For Each key In dict.Keys
        keys(i) = CStr(key)
        i = i + 1
    Next key

    QuickSortStrings keys, LBound(keys), UBound(keys)
    DictionaryKeysToSortedArray = keys
End Function

Private Sub QuickSortStrings(ByRef values() As String, ByVal firstIndex As Long, ByVal lastIndex As Long)
    Dim lowIndex As Long
    Dim highIndex As Long
    Dim pivot As String
    Dim swapValue As String

    lowIndex = firstIndex
    highIndex = lastIndex
    pivot = values((firstIndex + lastIndex) \ 2)

    Do While lowIndex <= highIndex
        Do While StrComp(values(lowIndex), pivot, vbTextCompare) < 0
            lowIndex = lowIndex + 1
        Loop
        Do While StrComp(values(highIndex), pivot, vbTextCompare) > 0
            highIndex = highIndex - 1
        Loop

        If lowIndex <= highIndex Then
            swapValue = values(lowIndex)
            values(lowIndex) = values(highIndex)
            values(highIndex) = swapValue
            lowIndex = lowIndex + 1
            highIndex = highIndex - 1
        End If
    Loop

    If firstIndex < highIndex Then QuickSortStrings values, firstIndex, highIndex
    If lowIndex < lastIndex Then QuickSortStrings values, lowIndex, lastIndex
End Sub

Private Function EmptyStringArray() As String()
    Dim values() As String
    EmptyStringArray = values
End Function

Private Function ArrayCount(ByRef values() As String) As Long
    On Error Resume Next
    ArrayCount = UBound(values) - LBound(values) + 1
    If Err.Number <> 0 Then
        Err.Clear
        ArrayCount = 0
    End If
    On Error GoTo 0
End Function

Private Function IsActiveFlag(ByVal activeValue As Variant) As Boolean
    If IsError(activeValue) Then Exit Function
    If IsEmpty(activeValue) Or IsNull(activeValue) Then Exit Function

    If VarType(activeValue) = vbBoolean Then
        IsActiveFlag = CBool(activeValue)
        Exit Function
    End If

    If IsNumeric(activeValue) Then
        IsActiveFlag = (CDbl(activeValue) <> 0)
        Exit Function
    End If

    Select Case LCase$(Trim$(CStr(activeValue)))
        Case "true", "yes", "y", "1"
            IsActiveFlag = True
    End Select
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

Private Function LastUsedRowInColumns( _
    ByVal ws As Worksheet, _
    ByVal firstColumn As String, _
    ByVal lastColumn As String) As Long

    Dim colIndex As Long
    Dim maxRow As Long
    Dim firstColIndex As Long
    Dim lastColIndex As Long

    firstColIndex = ws.Columns(firstColumn).Column
    lastColIndex = ws.Columns(lastColumn).Column
    maxRow = 1

    For colIndex = firstColIndex To lastColIndex
        If ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row > maxRow Then
            maxRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row
        End If
    Next colIndex

    LastUsedRowInColumns = maxRow
End Function

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    SheetExists = Not ws Is Nothing
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
