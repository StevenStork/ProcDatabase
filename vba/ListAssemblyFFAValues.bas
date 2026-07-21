Attribute VB_Name = "ListAssemblyFFAValues"
Option Explicit

Private Const HEADER_BASE_PART As String = "Base Part Number"
Private Const HEADER_DASH_CONDITIONS As String = "Dash Conditions"
Private Const HEADER_FFA As String = "FFA"

' Builds a summary of unique base part numbers, their available dash
' conditions, and FFA values from AssyStndTbl on the Assembly Standards
' sheet, then writes the results to the Home sheet starting at C1.
Public Sub ListAssemblyFFAValues()
    Dim wsSource As Worksheet
    Dim wsHome As Worksheet
    Dim tbl As ListObject
    Dim colAssembly As ListColumn
    Dim colFFA As ListColumn
    Dim dashConditionsByBase As Object
    Dim ffaValuesByBase As Object
    Dim rowIndex As Long
    Dim assemblyVal As String
    Dim ffaVal As String
    Dim basePart As String
    Dim dashCondition As String
    Dim outputRow As Long
    Dim basePartKey As Variant
    Dim basePartKeys() As String
    Dim i As Long

    On Error GoTo CleanUp
    OptimizeExcel True

    Set wsSource = ThisWorkbook.Worksheets("Assembly Standards")
    Set wsHome = ThisWorkbook.Worksheets("Home")
    Set tbl = wsSource.ListObjects("AssyStndTbl")
    Set colAssembly = tbl.ListColumns("ASSEMBLY NO")
    Set colFFA = tbl.ListColumns("FFA")

    Set dashConditionsByBase = CreateObject("Scripting.Dictionary")
    dashConditionsByBase.CompareMode = vbTextCompare
    Set ffaValuesByBase = CreateObject("Scripting.Dictionary")
    ffaValuesByBase.CompareMode = vbTextCompare

    For rowIndex = 1 To tbl.ListRows.Count
        assemblyVal = Trim$(CStr(colAssembly.DataBodyRange.Cells(rowIndex, 1).Value))
        ffaVal = Trim$(CStr(colFFA.DataBodyRange.Cells(rowIndex, 1).Value))

        If Len(assemblyVal) > 0 Then
            SplitAssemblyNo assemblyVal, basePart, dashCondition
            EnsureBasePartEntry dashConditionsByBase, ffaValuesByBase, basePart

            If Len(dashCondition) > 0 Then
                AddUniqueValue dashConditionsByBase(basePart), dashCondition
            End If

            If Len(ffaVal) > 0 Then
                AddUniqueValue ffaValuesByBase(basePart), ffaVal
            End If
        End If
    Next rowIndex

    wsHome.Range("C1:E" & wsHome.Rows.Count).ClearContents
    wsHome.Cells(1, "C").Value = HEADER_BASE_PART
    wsHome.Cells(1, "D").Value = HEADER_DASH_CONDITIONS
    wsHome.Cells(1, "E").Value = HEADER_FFA

    If dashConditionsByBase.Count = 0 Then
        GoTo CleanUp
    End If

    ReDim basePartKeys(0 To dashConditionsByBase.Count - 1)
    i = 0
    For Each basePartKey In dashConditionsByBase.Keys
        basePartKeys(i) = CStr(basePartKey)
        i = i + 1
    Next basePartKey

    SortStringArray basePartKeys

    outputRow = 2
    For i = LBound(basePartKeys) To UBound(basePartKeys)
        wsHome.Cells(outputRow, "C").Value = basePartKeys(i)
        wsHome.Cells(outputRow, "D").Value = JoinDictionaryKeys(dashConditionsByBase(basePartKeys(i)))
        wsHome.Cells(outputRow, "E").Value = JoinDictionaryKeys(ffaValuesByBase(basePartKeys(i)))
        outputRow = outputRow + 1
    Next i

CleanUp:
    OptimizeExcel False
End Sub

Private Sub EnsureBasePartEntry(ByVal dashConditionsByBase As Object, ByVal ffaValuesByBase As Object, ByVal basePart As String)
    If Not dashConditionsByBase.Exists(basePart) Then
        Set dashConditionsByBase(basePart) = CreateObject("Scripting.Dictionary")
        dashConditionsByBase(basePart).CompareMode = vbTextCompare
        Set ffaValuesByBase(basePart) = CreateObject("Scripting.Dictionary")
        ffaValuesByBase(basePart).CompareMode = vbTextCompare
    End If
End Sub

Private Sub AddUniqueValue(ByVal valueMap As Object, ByVal value As String)
    If Not valueMap.Exists(value) Then
        valueMap.Add value, value
    End If
End Sub

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
            If StrComp(keys(i), keys(j), vbTextCompare) > 0 Then
                tempKey = keys(i)
                keys(i) = keys(j)
                keys(j) = tempKey
            End If
        Next j
    Next i
End Sub

Private Function JoinDictionaryKeys(ByVal valueMap As Object) As String
    Dim keys() As String
    Dim key As Variant
    Dim index As Long

    If valueMap.Count = 0 Then
        JoinDictionaryKeys = vbNullString
        Exit Function
    End If

    ReDim keys(0 To valueMap.Count - 1)
    index = 0
    For Each key In valueMap.Keys
        keys(index) = CStr(key)
        index = index + 1
    Next key

    SortStringArray keys
    JoinDictionaryKeys = Join(keys, ", ")
End Function
