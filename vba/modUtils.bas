Attribute VB_Name = "modUtils"
Option Explicit

Public Function Nz(ByVal value As Variant) As Variant
    If IsError(value) Then
        Nz = vbNullString
    ElseIf IsEmpty(value) Or IsNull(value) Then
        Nz = vbNullString
    Else
        Nz = value
    End If
End Function

Public Function IsActiveFlag(ByVal activeValue As Variant) As Boolean
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

Public Function LastUsedRowInColumn(ByVal ws As Worksheet, ByVal columnLetter As String) As Long
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

Public Function FastLastUsedRowInColumn(ByVal ws As Worksheet, ByVal columnLetter As String) As Long
    FastLastUsedRowInColumn = ws.Cells(ws.Rows.Count, columnLetter).End(xlUp).Row
End Function

Public Function FastLastUsedRowInColumns( _
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

Public Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    SheetExists = Not ws Is Nothing
End Function

Public Function GetWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Public Function SanitizeSheetName(ByVal proposedName As String) As String
    Dim cleaned As String

    cleaned = proposedName
    cleaned = Replace$(cleaned, "\", "_")
    cleaned = Replace$(cleaned, "/", "_")
    cleaned = Replace$(cleaned, "?", "_")
    cleaned = Replace$(cleaned, "*", "_")
    cleaned = Replace$(cleaned, "[", "_")
    cleaned = Replace$(cleaned, "]", "_")
    cleaned = Replace$(cleaned, ":", "_")

    If Len(cleaned) > 31 Then cleaned = Left$(cleaned, 31)

    If Len(cleaned) = 0 Then
        Err.Raise vbObjectError + 600, "SanitizeSheetName", "Proposed name was empty after sanitizing."
    End If

    SanitizeSheetName = cleaned
End Function

Public Function UniqueSheetName(ByVal proposedName As String) As String
    Dim baseName As String
    Dim candidate As String
    Dim suffix As Long

    baseName = SanitizeSheetName(proposedName)
    candidate = baseName
    suffix = 1

    Do While SheetExists(candidate)
        suffix = suffix + 1
        candidate = Left$(baseName, 31 - Len(CStr(suffix)) - 1) & "-" & CStr(suffix)
    Loop

    UniqueSheetName = candidate
End Function

Public Function EmptyStringArray() As String()
    Dim values() As String
    EmptyStringArray = values
End Function

Public Function IsArrayInitialized(ByRef values() As String) As Boolean
    Dim upperBound As Long

    On Error Resume Next
    upperBound = UBound(values)
    IsArrayInitialized = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function ArrayCount(ByRef values() As String) As Long
    If Not IsArrayInitialized(values) Then
        ArrayCount = 0
    Else
        ArrayCount = UBound(values) - LBound(values) + 1
    End If
End Function

Public Function CloneStringArray(ByRef values() As String) As String()
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

Public Function JoinStringArray(ByRef values() As String) As String
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

Public Function StringArraysEqual(ByRef leftValues() As String, ByRef rightValues() As String) As Boolean
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

Public Function ArrayContainsValue(ByRef values() As String, ByVal itemName As String) As Boolean
    Dim i As Long

    For i = 0 To ArrayCount(values) - 1
        If StrComp(values(LBound(values) + i), itemName, vbTextCompare) = 0 Then
            ArrayContainsValue = True
            Exit Function
        End If
    Next i
End Function

Public Function DictionaryKeysToSortedArray(ByVal valueMap As Object) As String()
    Dim keys() As String
    Dim key As Variant
    Dim index As Long

    If valueMap Is Nothing Or valueMap.Count = 0 Then
        DictionaryKeysToSortedArray = EmptyStringArray()
        Exit Function
    End If

    ReDim keys(0 To valueMap.Count - 1)
    index = 0
    For Each key In valueMap.Keys
        keys(index) = CStr(key)
        index = index + 1
    Next key

    QuickSortStrings keys, LBound(keys), UBound(keys)
    DictionaryKeysToSortedArray = keys
End Function

Public Sub QuickSortStrings(ByRef values() As String, ByVal firstIndex As Long, ByVal lastIndex As Long)
    Dim lowIndex As Long
    Dim highIndex As Long
    Dim pivot As String
    Dim swapValue As String

    If Not IsArrayInitialized(values) Then Exit Sub
    If firstIndex >= lastIndex Then Exit Sub

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

Public Function HashVariantColumn(ByVal values As Variant) As String
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

Public Function HashString(ByVal value As String) As String
    Dim i As Long
    Dim checkSum As Long
    Dim ch As String

    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        checkSum = (checkSum + AscW(ch) * (i Mod 31 + 1)) Mod 2147483647
    Next i

    HashString = CStr(Len(value)) & ":" & CStr(checkSum) & ":" & Left$(value, 24) & Chr$(30) & Right$(value, 24)
End Function

Public Function JoinCheckedList(ByVal valueMap As Object) As String
    Dim keys() As String

    keys = DictionaryKeysToSortedArray(valueMap)
    If ArrayCount(keys) = 0 Then
        JoinCheckedList = vbNullString
    Else
        JoinCheckedList = Join(keys, ", ")
    End If
End Function

Public Function ListContainsToken(ByVal csvList As String, ByVal token As String) As Boolean
    If Len(csvList) = 0 Or Len(token) = 0 Then Exit Function
    ListContainsToken = (InStr(1, ", " & csvList & ", ", ", " & token & ", ", vbTextCompare) > 0)
End Function

Public Function IsPartSheet(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    IsPartSheet = (StrComp(Trim$(CStr(Nz(ws.Range(CATEGORY_CELL).Value))), PART_LABEL_VALUE, vbTextCompare) = 0)
End Function

Public Function IsExportSheet(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    IsExportSheet = (StrComp(Trim$(CStr(Nz(ws.Range(CATEGORY_CELL).Value))), EXPORT_LABEL_VALUE, vbTextCompare) = 0)
End Function

Public Function IsHiddenSystemSheet(ByVal ws As Worksheet) As Boolean
    Dim labelName As String

    If ws Is Nothing Then Exit Function
    labelName = Trim$(CStr(Nz(ws.Range(CATEGORY_CELL).Value)))
    If StrComp(labelName, REFS_LABEL_VALUE, vbTextCompare) = 0 Then
        IsHiddenSystemSheet = True
    ElseIf StrComp(labelName, DATA_LABEL_VALUE, vbTextCompare) = 0 Then
        IsHiddenSystemSheet = True
    End If
End Function

Public Sub SplitAssemblyNo(ByVal assemblyNo As String, ByRef basePart As String, ByRef dashCondition As String)
    Dim dashPos As Long

    dashPos = InStr(1, assemblyNo, "-", vbBinaryCompare)
    If dashPos > 0 Then
        basePart = Trim$(Left$(assemblyNo, dashPos - 1))
        dashCondition = Trim$(Mid$(assemblyNo, dashPos + 1))
    Else
        basePart = Trim$(assemblyNo)
        dashCondition = vbNullString
    End If
End Sub

Public Function UniqueSortedValuesFromColumn(ByVal rawValues As Variant) As String()
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
            If Not uniqueValues.Exists(cellValue) Then uniqueValues.Add cellValue, cellValue
        End If
    Next rowIndex

    UniqueSortedValuesFromColumn = DictionaryKeysToSortedArray(uniqueValues)
End Function

Public Function ReadColumnList(ByVal ws As Worksheet, ByVal columnLetter As String, ByVal startRow As Long) As String()
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim values() As String
    Dim valueCount As Long
    Dim rawValues As Variant
    Dim cellValue As String

    lastRow = FastLastUsedRowInColumn(ws, columnLetter)
    If lastRow < startRow Then
        ReadColumnList = EmptyStringArray()
        Exit Function
    End If

    rawValues = ws.Range(ws.Cells(startRow, columnLetter), ws.Cells(lastRow, columnLetter)).Value2
    valueCount = 0

    If Not IsArray(rawValues) Then
        cellValue = Trim$(CStr(Nz(rawValues)))
        If Len(cellValue) = 0 Then
            ReadColumnList = EmptyStringArray()
        Else
            ReDim values(0 To 0)
            values(0) = cellValue
            ReadColumnList = values
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

Public Function ListObjectByName(ByVal ws As Worksheet, ByVal tableName As String) As ListObject
    On Error Resume Next
    Set ListObjectByName = ws.ListObjects(tableName)
    On Error GoTo 0
End Function

Public Function ListColumnValues(ByVal col As ListColumn) As Variant
    Dim values As Variant
    Dim result(1 To 1, 1 To 1) As Variant

    If col Is Nothing Then
        ListColumnValues = Empty
        Exit Function
    End If
    If col.DataBodyRange Is Nothing Then
        ListColumnValues = Empty
        Exit Function
    End If

    values = col.DataBodyRange.Value2
    If IsArray(values) Then
        ListColumnValues = values
    Else
        result(1, 1) = values
        ListColumnValues = result
    End If
End Function
