Attribute VB_Name = "ListAssemblyFFAValues"
Option Explicit

' Builds a list of unique ASSEMBLY NO values from AssyStndTbl and, for each,
' writes the corresponding unique FFA values (comma-separated) on the Home sheet.
Public Sub ListAssemblyFFAValues()
    Dim wsSource As Worksheet
    Dim wsHome As Worksheet
    Dim tbl As ListObject
    Dim colAssembly As ListColumn
    Dim colFFA As ListColumn
    Dim assemblyMap As Object
    Dim rowIndex As Long
    Dim assemblyVal As String
    Dim ffaVal As String
    Dim outputRow As Long
    Dim assemblyKey As Variant
    Dim ffaKey As Variant
    Dim assemblyKeys() As String
    Dim i As Long
    Dim j As Long
    Dim tempKey As String

    Set wsSource = ThisWorkbook.Worksheets("Assembly Standard")
    Set wsHome = ThisWorkbook.Worksheets("Home")
    Set tbl = wsSource.ListObjects("AssyStndTbl")
    Set colAssembly = tbl.ListColumns("ASSEMBLY NO")
    Set colFFA = tbl.ListColumns("FFA")

    Set assemblyMap = CreateObject("Scripting.Dictionary")
    assemblyMap.CompareMode = vbTextCompare

    For rowIndex = 1 To tbl.ListRows.Count
        assemblyVal = Trim$(CStr(colAssembly.DataBodyRange.Cells(rowIndex, 1).Value))
        ffaVal = Trim$(CStr(colFFA.DataBodyRange.Cells(rowIndex, 1).Value))

        If Len(assemblyVal) > 0 Then
            If Not assemblyMap.Exists(assemblyVal) Then
                Set assemblyMap(assemblyVal) = CreateObject("Scripting.Dictionary")
                assemblyMap(assemblyVal).CompareMode = vbTextCompare
            End If

            If Len(ffaVal) > 0 Then
                If Not assemblyMap(assemblyVal).Exists(ffaVal) Then
                    assemblyMap(assemblyVal).Add ffaVal, ffaVal
                End If
            End If
        End If
    Next rowIndex

    wsHome.Range("C2:D" & wsHome.Rows.Count).ClearContents

    If assemblyMap.Count = 0 Then
        Exit Sub
    End If

    ReDim assemblyKeys(0 To assemblyMap.Count - 1)
    i = 0
    For Each assemblyKey In assemblyMap.Keys
        assemblyKeys(i) = CStr(assemblyKey)
        i = i + 1
    Next assemblyKey

    For i = LBound(assemblyKeys) To UBound(assemblyKeys) - 1
        For j = i + 1 To UBound(assemblyKeys)
            If StrComp(assemblyKeys(i), assemblyKeys(j), vbTextCompare) > 0 Then
                tempKey = assemblyKeys(i)
                assemblyKeys(i) = assemblyKeys(j)
                assemblyKeys(j) = tempKey
            End If
        Next j
    Next i

    outputRow = 2
    For i = LBound(assemblyKeys) To UBound(assemblyKeys)
        wsHome.Cells(outputRow, "C").Value = assemblyKeys(i)
        wsHome.Cells(outputRow, "D").Value = JoinDictionaryKeys(assemblyMap(assemblyKeys(i)))
        outputRow = outputRow + 1
    Next i
End Sub

Private Function JoinDictionaryKeys(ByVal valueMap As Object) As String
    Dim keys() As String
    Dim key As Variant
    Dim index As Long
    Dim i As Long
    Dim j As Long
    Dim tempKey As String

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

    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If StrComp(keys(i), keys(j), vbTextCompare) > 0 Then
                tempKey = keys(i)
                keys(i) = keys(j)
                keys(j) = tempKey
            End If
        Next j
    Next i

    JoinDictionaryKeys = Join(keys, ", ")
End Function
