Attribute VB_Name = "ListAssemblyFFAValues"
Option Explicit

Private Const HEADER_BASE_PART As String = "Base Part Number"
Private Const HEADER_DASH_CONDITIONS As String = "Dash Conditions"
Private Const HEADER_FFA As String = "FFA"
Private Const HEADER_ACTIVE_PART As String = "Active Part"
Private Const SUMMARY_TABLE_NAME As String = "BasePartSummaryTbl"
Private Const SUMMARY_HEADER_ROW As Long = 2
Private Const SUMMARY_FIRST_DATA_ROW As Long = 3
Private Const CHECKBOX_NAME_PREFIX As String = "chkActivePart_"

' Builds a formatted summary table of unique base part numbers, dash
' conditions, and FFA values from AssyStndTbl on the Assembly Standards
' sheet, then writes the results to the Home sheet starting at C2.
Public Sub ListAssemblyFFAValues()
    Dim wsSource As Worksheet
    Dim wsHome As Worksheet
    Dim tbl As ListObject
    Dim colAssembly As ListColumn
    Dim colFFA As ListColumn
    Dim dashConditionsByBase As Object
    Dim ffaValuesByBase As Object
    Dim savedActiveStates As Object
    Dim rowIndex As Long
    Dim assemblyVal As String
    Dim ffaVal As String
    Dim basePart As String
    Dim dashCondition As String
    Dim outputRow As Long
    Dim lastRow As Long
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

    Set savedActiveStates = CaptureActivePartStates(wsHome)

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

    RemoveSummaryArtifacts wsHome

    wsHome.Cells(SUMMARY_HEADER_ROW, "C").Value = HEADER_BASE_PART
    wsHome.Cells(SUMMARY_HEADER_ROW, "D").Value = HEADER_DASH_CONDITIONS
    wsHome.Cells(SUMMARY_HEADER_ROW, "E").Value = HEADER_FFA
    wsHome.Cells(SUMMARY_HEADER_ROW, "F").Value = HEADER_ACTIVE_PART

    If dashConditionsByBase.Count = 0 Then
        FormatSummaryTable wsHome, SUMMARY_HEADER_ROW
        GoTo CleanUp
    End If

    ReDim basePartKeys(0 To dashConditionsByBase.Count - 1)
    i = 0
    For Each basePartKey In dashConditionsByBase.Keys
        basePartKeys(i) = CStr(basePartKey)
        i = i + 1
    Next basePartKey

    SortStringArray basePartKeys

    outputRow = SUMMARY_FIRST_DATA_ROW
    For i = LBound(basePartKeys) To UBound(basePartKeys)
        wsHome.Cells(outputRow, "C").Value = basePartKeys(i)
        wsHome.Cells(outputRow, "D").Value = JoinDictionaryKeys(dashConditionsByBase(basePartKeys(i)))
        wsHome.Cells(outputRow, "E").Value = JoinDictionaryKeys(ffaValuesByBase(basePartKeys(i)))
        wsHome.Cells(outputRow, "F").Value = False
        outputRow = outputRow + 1
    Next i

    lastRow = outputRow - 1
    FormatSummaryTable wsHome, lastRow
    AddActivePartCheckBoxes wsHome, SUMMARY_FIRST_DATA_ROW, lastRow, savedActiveStates

CleanUp:
    OptimizeExcel False
End Sub

Private Function CaptureActivePartStates(ByVal ws As Worksheet) As Object
    Dim states As Object
    Dim summaryTable As ListObject
    Dim rowIndex As Long
    Dim basePart As String
    Dim activeValue As Variant

    Set states = CreateObject("Scripting.Dictionary")
    states.CompareMode = vbTextCompare

    On Error Resume Next
    Set summaryTable = ws.ListObjects(SUMMARY_TABLE_NAME)
    On Error GoTo 0

    If summaryTable Is Nothing Then
        Set CaptureActivePartStates = states
        Exit Function
    End If

    If summaryTable.DataBodyRange Is Nothing Then
        Set CaptureActivePartStates = states
        Exit Function
    End If

    For rowIndex = 1 To summaryTable.DataBodyRange.Rows.Count
        basePart = Trim$(CStr(summaryTable.ListColumns(HEADER_BASE_PART).DataBodyRange.Cells(rowIndex, 1).Value))
        activeValue = summaryTable.ListColumns(HEADER_ACTIVE_PART).DataBodyRange.Cells(rowIndex, 1).Value

        If Len(basePart) > 0 Then
            states(basePart) = CBool(activeValue)
        End If
    Next rowIndex

    Set CaptureActivePartStates = states
End Function

Private Sub RemoveSummaryArtifacts(ByVal ws As Worksheet)
    Dim checkboxIndex As Long

    For checkboxIndex = ws.CheckBoxes.Count To 1 Step -1
        If Left$(ws.CheckBoxes(checkboxIndex).Name, Len(CHECKBOX_NAME_PREFIX)) = CHECKBOX_NAME_PREFIX Then
            ws.CheckBoxes(checkboxIndex).Delete
        End If
    Next checkboxIndex

    On Error Resume Next
    ws.ListObjects(SUMMARY_TABLE_NAME).Delete
    On Error GoTo 0

    ws.Range("C" & SUMMARY_HEADER_ROW & ":F" & ws.Rows.Count).Clear
End Sub

Private Sub FormatSummaryTable(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim summaryTable As ListObject
    Dim tableRange As Range

    On Error Resume Next
    ws.ListObjects(SUMMARY_TABLE_NAME).Delete
    On Error GoTo 0

    Set tableRange = ws.Range("C" & SUMMARY_HEADER_ROW & ":F" & lastRow)
    Set summaryTable = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    summaryTable.Name = SUMMARY_TABLE_NAME
    summaryTable.TableStyle = "TableStyleMedium2"

    With summaryTable.HeaderRowRange
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    If Not summaryTable.DataBodyRange Is Nothing Then
        summaryTable.ListColumns(HEADER_BASE_PART).DataBodyRange.HorizontalAlignment = xlLeft
        summaryTable.ListColumns(HEADER_DASH_CONDITIONS).DataBodyRange.HorizontalAlignment = xlLeft
        summaryTable.ListColumns(HEADER_FFA).DataBodyRange.HorizontalAlignment = xlLeft
    End If

    With summaryTable.ListColumns(HEADER_ACTIVE_PART).Range
        .HorizontalAlignment = xlCenter
        .ColumnWidth = 12
    End With
End Sub

Private Sub AddActivePartCheckBoxes( _
    ByVal ws As Worksheet, _
    ByVal firstDataRow As Long, _
    ByVal lastRow As Long, _
    ByVal savedActiveStates As Object)
    Dim outputRow As Long

    For outputRow = firstDataRow To lastRow
        AddActivePartCheckBox ws.Cells(outputRow, "F"), savedActiveStates
    Next outputRow
End Sub

Private Sub AddActivePartCheckBox(ByVal targetCell As Range, ByVal savedActiveStates As Object)
    Dim cb As CheckBox
    Dim cbWidth As Double
    Dim cbHeight As Double
    Dim basePart As String
    Dim isActive As Boolean

    basePart = Trim$(CStr(targetCell.Offset(0, -3).Value))

    If Not savedActiveStates Is Nothing Then
        If savedActiveStates.Exists(basePart) Then
            isActive = CBool(savedActiveStates(basePart))
        End If
    End If

    targetCell.Value = isActive
    targetCell.HorizontalAlignment = xlCenter
    targetCell.VerticalAlignment = xlCenter

    cbWidth = 14
    cbHeight = 14

    Set cb = targetCell.Worksheet.CheckBoxes.Add( _
        targetCell.Left + (targetCell.Width - cbWidth) / 2, _
        targetCell.Top + (targetCell.Height - cbHeight) / 2, _
        cbWidth, _
        cbHeight)
    cb.Name = CHECKBOX_NAME_PREFIX & targetCell.Row
    cb.Caption = vbNullString
    cb.LinkedCell = targetCell.Address(False, False)
    cb.Placement = xlMove
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
