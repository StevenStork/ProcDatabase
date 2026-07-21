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
Private Const FFA_COLUMN As String = "C"
Private Const DASH_COLUMN As String = "E"
Private Const PRODUCT_LINE_COLUMN As String = "G"

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

    SyncColumnList ws, FFA_COLUMN, ffaValues
    SyncColumnList ws, DASH_COLUMN, dashConditions
    SyncColumnList ws, PRODUCT_LINE_COLUMN, productLines
End Sub

' Writes values only when the existing column list differs from the source list.
Private Sub SyncColumnList(ByVal ws As Worksheet, ByVal columnLetter As String, ByRef sourceValues() As String)
    Dim currentValues() As String
    Dim lastDataRow As Long
    Dim clearRow As Long
    Dim i As Long
    Dim sourceCount As Long

    currentValues = ReadColumnList(ws, columnLetter)

    If StringArraysEqual(currentValues, sourceValues) Then Exit Sub

    sourceCount = ArrayCount(sourceValues)

    If sourceCount = 0 Then
        lastDataRow = LastUsedRowInColumn(ws, columnLetter)
        If lastDataRow >= LIST_START_ROW Then
            ws.Range(ws.Cells(LIST_START_ROW, columnLetter), ws.Cells(lastDataRow, columnLetter)).ClearContents
        End If
        Exit Sub
    End If

    For i = 0 To sourceCount - 1
        ws.Cells(LIST_START_ROW + i, columnLetter).Value = sourceValues(i)
    Next i

    lastDataRow = LastUsedRowInColumn(ws, columnLetter)
    clearRow = LIST_START_ROW + sourceCount
    If lastDataRow >= clearRow Then
        ws.Range(ws.Cells(clearRow, columnLetter), ws.Cells(lastDataRow, columnLetter)).ClearContents
    End If
End Sub

Private Function ReadColumnList(ByVal ws As Worksheet, ByVal columnLetter As String) As String()
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim values() As String
    Dim valueCount As Long
    Dim cellValue As String

    lastRow = LastUsedRowInColumn(ws, columnLetter)
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
            ' Stop at the first blank so trailing empties do not force a rewrite.
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
