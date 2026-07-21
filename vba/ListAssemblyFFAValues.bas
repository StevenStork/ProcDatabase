Attribute VB_Name = "ListAssemblyFFAValues"
Option Explicit

Private Const HEADER_BASE_PART As String = "Base Part Number"
Private Const HEADER_DASH_CONDITIONS As String = "Dash Conditions"

' Builds a summary of unique base part numbers and their available dash
' conditions from AssyStndTbl on the Assembly Standards sheet, then writes
' the results to the Home sheet starting at C1 (headers) and C2 (data).
Public Sub ListAssemblyFFAValues()
    Dim wsSource As Worksheet
    Dim wsHome As Worksheet
    Dim tbl As ListObject
    Dim colAssembly As ListColumn
    Dim basePartMap As Object
    Dim rowIndex As Long
    Dim assemblyVal As String
    Dim basePart As String
    Dim dashCondition As String
    Dim outputRow As Long
    Dim basePartKey As Variant
    Dim basePartKeys() As String
    Dim i As Long
    Dim j As Long
    Dim tempKey As String

    On Error GoTo CleanUp
    OptimizeExcel True

    Set wsSource = ThisWorkbook.Worksheets("Assembly Standards")
    Set wsHome = ThisWorkbook.Worksheets("Home")
    Set tbl = wsSource.ListObjects("AssyStndTbl")
    Set colAssembly = tbl.ListColumns("ASSEMBLY NO")

    Set basePartMap = CreateObject("Scripting.Dictionary")
    basePartMap.CompareMode = vbTextCompare

    For rowIndex = 1 To tbl.ListRows.Count
        assemblyVal = Trim$(CStr(colAssembly.DataBodyRange.Cells(rowIndex, 1).Value))

        If Len(assemblyVal) > 0 Then
            SplitAssemblyNo assemblyVal, basePart, dashCondition

            If Not basePartMap.Exists(basePart) Then
                Set basePartMap(basePart) = CreateObject("Scripting.Dictionary")
                basePartMap(basePart).CompareMode = vbTextCompare
            End If

            If Len(dashCondition) > 0 Then
                If Not basePartMap(basePart).Exists(dashCondition) Then
                    basePartMap(basePart).Add dashCondition, dashCondition
                End If
            End If
        End If
    Next rowIndex

    wsHome.Range("C1:D" & wsHome.Rows.Count).ClearContents
    wsHome.Cells(1, "C").Value = HEADER_BASE_PART
    wsHome.Cells(1, "D").Value = HEADER_DASH_CONDITIONS

    If basePartMap.Count = 0 Then
        GoTo CleanUp
    End If

    ReDim basePartKeys(0 To basePartMap.Count - 1)
    i = 0
    For Each basePartKey In basePartMap.Keys
        basePartKeys(i) = CStr(basePartKey)
        i = i + 1
    Next basePartKey

    SortStringArray basePartKeys

    outputRow = 2
    For i = LBound(basePartKeys) To UBound(basePartKeys)
        wsHome.Cells(outputRow, "C").Value = basePartKeys(i)
        wsHome.Cells(outputRow, "D").Value = JoinDictionaryKeys(basePartMap(basePartKeys(i)))
        outputRow = outputRow + 1
    Next i

CleanUp:
    OptimizeExcel False
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
