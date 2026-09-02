Attribute VB_Name = "modPartIO"
Option Explicit

'==============================================================================
' Part-number relational data helpers.
'==============================================================================

Public Function BasePartHasAssignments(ByVal basePartCode As String) As Boolean
    BasePartHasAssignments = JunctionTableHasValue(FindTable(PART_DASH_CONDITIONS_TABLE_NAME), COL_BASE_PART_CODE, basePartCode) _
        Or JunctionTableHasValue(FindTable(PART_OPERATIONS_TABLE_NAME), COL_BASE_PART_CODE, basePartCode)
End Function

Public Function BasePartsReferenceFactory(ByVal factoryCode As String) As Boolean
    BasePartsReferenceFactory = JunctionTableHasValue(FindTable(BASE_PARTS_TABLE_NAME), COL_FACTORY_CODE, factoryCode)
End Function

Public Function ListPartAssignmentCodes( _
    ByVal tbl As ListObject, _
    ByVal basePartCode As String, _
    ByVal displayColumnName As String) As Variant

    ListPartAssignmentCodes = ListJunctionDisplayItems(tbl, COL_BASE_PART_CODE, basePartCode, displayColumnName)
End Function

Public Function SplitAssemblyNo(ByVal assemblyNo As String, ByRef basePartCode As String, ByRef dashCondition As String)
    Dim dashPos As Long

    assemblyNo = Trim$(assemblyNo)
    dashPos = InStr(1, assemblyNo, "-", vbBinaryCompare)

    If dashPos > 0 Then
        basePartCode = Trim$(Left$(assemblyNo, dashPos - 1))
        dashCondition = Trim$(Mid$(assemblyNo, dashPos + 1))
    Else
        basePartCode = assemblyNo
        dashCondition = vbNullString
    End If
End Function

Public Function BuildActiveAssemblyNumberList() As String
    Dim basePartsTbl As ListObject
    Dim dashTbl As ListObject
    Dim basePartCodes As Variant
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim basePartCode As String
    Dim assemblyList As Object
    Dim result As String
    Dim assemblyKey As Variant

    Set assemblyList = CreateObject("Scripting.Dictionary")
    assemblyList.CompareMode = vbTextCompare

    Set basePartsTbl = FindTable(BASE_PARTS_TABLE_NAME)
    Set dashTbl = FindTable(PART_DASH_CONDITIONS_TABLE_NAME)
    If basePartsTbl Is Nothing Or basePartsTbl.DataBodyRange Is Nothing Then Exit Function

    basePartCodes = basePartsTbl.ListColumns(COL_BASE_PART_CODE).DataBodyRange.Value2
    If Not IsArray(basePartCodes) Then
        basePartCodes = ArrayBasePartSingle(basePartsTbl)
        If Not IsArray(basePartCodes) Then Exit Function
    End If

    rowCount = UBound(basePartCodes, 1)
    For rowIndex = 1 To rowCount
        basePartCode = NormalizeCode(basePartCodes(rowIndex, 1))
        If Len(basePartCode) = 0 Then GoTo ContinueBasePart
        If Not IsActiveFlag(GetCellValueByListRow(basePartsTbl, basePartsTbl.ListRows(rowIndex).Index, COL_ACTIVE)) Then GoTo ContinueBasePart

        AppendActiveAssembliesForPart dashTbl, basePartCode, assemblyList

ContinueBasePart:
    Next rowIndex

    For Each assemblyKey In assemblyList.Keys
        If Len(result) > 0 Then result = result & ", "
        result = result & CStr(assemblyKey)
    Next assemblyKey

    BuildActiveAssemblyNumberList = result
End Function

Private Sub AppendActiveAssembliesForPart( _
    ByVal dashTbl As ListObject, _
    ByVal basePartCode As String, _
    ByVal assemblyList As Object)

    Dim dashCodes As Variant
    Dim dashRowIndex As Long
    Dim dashRowCount As Long
    Dim dashCondition As String
    Dim assemblyNo As String

    If dashTbl Is Nothing Or dashTbl.DataBodyRange Is Nothing Then
        assemblyList(basePartCode) = basePartCode
        Exit Sub
    End If

    dashCodes = dashTbl.ListColumns(COL_DASH_CONDITION).DataBodyRange.Value2
    If Not IsArray(dashCodes) Then Exit Sub

    dashRowCount = UBound(dashCodes, 1)
    For dashRowIndex = 1 To dashRowCount
        If Not ValuesMatchCode(dashTbl.ListColumns(COL_BASE_PART_CODE).DataBodyRange.Cells(dashRowIndex, 1).Value2, basePartCode) Then GoTo ContinueDash
        If Not IsActiveFlag(dashTbl.ListColumns(COL_ACTIVE).DataBodyRange.Cells(dashRowIndex, 1).Value2) Then GoTo ContinueDash

        dashCondition = NormalizeCode(dashCodes(dashRowIndex, 1))
        If Len(dashCondition) = 0 Then GoTo ContinueDash

        assemblyNo = basePartCode & "-" & dashCondition
        If Not assemblyList.Exists(assemblyNo) Then assemblyList.Add assemblyNo, assemblyNo

ContinueDash:
    Next dashRowIndex
End Sub

Private Function ArrayBasePartSingle(ByVal basePartsTbl As ListObject) As Variant
    Dim values(1 To 1, 1 To 1) As Variant
    values(1, 1) = basePartsTbl.ListColumns(COL_BASE_PART_CODE).DataBodyRange.Value2
    ArrayBasePartSingle = values
End Function

Public Function JunctionTableHasValue( _
    ByVal junctionTbl As ListObject, _
    ByVal filterColumnName As String, _
    ByVal filterValue As String) As Boolean

    Dim filterValues As Variant
    Dim rowIndex As Long

    JunctionTableHasValue = False
    If junctionTbl Is Nothing Then Exit Function
    If junctionTbl.DataBodyRange Is Nothing Then Exit Function

    filterValues = junctionTbl.ListColumns(filterColumnName).DataBodyRange.Value2
    If Not IsArray(filterValues) Then
        JunctionTableHasValue = ValuesMatchCode(filterValues, filterValue)
        Exit Function
    End If

    For rowIndex = 1 To UBound(filterValues, 1)
        If ValuesMatchCode(filterValues(rowIndex, 1), filterValue) Then
            JunctionTableHasValue = True
            Exit Function
        End If
    Next rowIndex
End Function

Private Function IsArrayAllocated(ByVal arr As Variant) As Boolean
    On Error Resume Next
    IsArrayAllocated = IsArray(arr) And (UBound(arr) >= LBound(arr))
    On Error GoTo 0
End Function
