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

' Splits ASSEMBLY NO on "-" or a letter separator (e.g. 8545784-01, 8545784A01).
' DashCondition keeps leading zeros as text.
Public Sub SplitAssemblyNo(ByVal assemblyNo As String, ByRef basePartCode As String, ByRef dashCondition As String)
    Dim separator As String
    SplitAssemblyNoWithSeparator assemblyNo, basePartCode, separator, dashCondition
End Sub

Public Sub SplitAssemblyNoWithSeparator( _
    ByVal assemblyNo As String, _
    ByRef basePartCode As String, _
    ByRef separator As String, _
    ByRef dashCondition As String)

    Dim dashPos As Long
    Dim charIndex As Long
    Dim character As String
    Dim trailingDigitsStart As Long

    assemblyNo = Trim$(assemblyNo)
    basePartCode = vbNullString
    separator = vbNullString
    dashCondition = vbNullString

    If Len(assemblyNo) = 0 Then Exit Sub

    dashPos = InStr(1, assemblyNo, "-", vbBinaryCompare)
    If dashPos > 0 Then
        basePartCode = Trim$(Left$(assemblyNo, dashPos - 1))
        separator = "-"
        dashCondition = Mid$(assemblyNo, dashPos + 1)
        Exit Sub
    End If

    ' Letter separator: BASE + Letter + digits (e.g. 8545784A01).
    trailingDigitsStart = 0
    For charIndex = Len(assemblyNo) To 1 Step -1
        character = Mid$(assemblyNo, charIndex, 1)
        If Not IsDigitCharacter(character) Then Exit For
        trailingDigitsStart = charIndex
    Next charIndex

    If trailingDigitsStart > 2 Then
        character = Mid$(assemblyNo, trailingDigitsStart - 1, 1)
        If IsLetterCharacter(character) Then
            basePartCode = Left$(assemblyNo, trailingDigitsStart - 2)
            separator = character
            dashCondition = Mid$(assemblyNo, trailingDigitsStart)
            Exit Sub
        End If
    End If

    basePartCode = assemblyNo
End Sub

Public Function BuildAssemblyNo( _
    ByVal basePartCode As String, _
    ByVal separator As String, _
    ByVal dashCondition As String) As String

    basePartCode = Trim$(basePartCode)
    separator = Trim$(separator)
    dashCondition = CStr(dashCondition)

    If Len(basePartCode) = 0 Then Exit Function
    If Len(dashCondition) = 0 Then
        BuildAssemblyNo = basePartCode
        Exit Function
    End If

    If Len(separator) = 0 Then separator = "-"
    BuildAssemblyNo = basePartCode & separator & dashCondition
End Function

' After tblRCCP refresh: add missing dash rows as Active, mark missing-from-RCCP rows Inactive.
Public Function SyncPartDashConditionsFromRCCP() As String
    Dim rccpTbl As ListObject
    Dim dashTbl As ListObject
    Dim rccpAssemblies As Object
    Dim assemblyValues As Variant
    Dim basePnValues As Variant
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim assemblyNo As String
    Dim basePartCode As String
    Dim separator As String
    Dim dashCondition As String
    Dim preferredBase As String
    Dim keyValue As String
    Dim listRowIndex As Long
    Dim fieldValues As Object
    Dim addedCount As Long
    Dim activatedCount As Long
    Dim inactivatedCount As Long
    Dim dashKey As Variant

    Set rccpAssemblies = CreateObject("Scripting.Dictionary")
    rccpAssemblies.CompareMode = vbBinaryCompare

    Set rccpTbl = FindTable(LINKED_RCCP_TABLE)
    Set dashTbl = FindTable(PART_DASH_CONDITIONS_TABLE_NAME)

    If rccpTbl Is Nothing Then
        SyncPartDashConditionsFromRCCP = "tblRCCP not found as a ListObject."
        Exit Function
    End If
    If dashTbl Is Nothing Then
        SyncPartDashConditionsFromRCCP = "PartDashConditionsTbl not found. Run BootstrapCapacityTables."
        Exit Function
    End If
    If Not TableHasColumn(rccpTbl, COL_ASSEMBLY_NO) Then
        SyncPartDashConditionsFromRCCP = "tblRCCP is missing the ASSEMBLY NO column."
        Exit Function
    End If
    If rccpTbl.DataBodyRange Is Nothing Then
        SyncPartDashConditionsFromRCCP = "tblRCCP has no rows."
        Exit Function
    End If

    EnsureDashConditionTextFormat dashTbl

    assemblyValues = ColumnTextValues(rccpTbl.ListColumns(COL_ASSEMBLY_NO))
    If TableHasColumn(rccpTbl, COL_RCCP_BASE_PN) Then
        basePnValues = ColumnTextValues(rccpTbl.ListColumns(COL_RCCP_BASE_PN))
    Else
        basePnValues = Empty
    End If

    rowCount = UBound(assemblyValues, 1)
    For rowIndex = 1 To rowCount
        assemblyNo = Trim$(CStr(assemblyValues(rowIndex, 1)))
        If Len(assemblyNo) = 0 Then GoTo ContinueRccp

        preferredBase = vbNullString
        If IsArray(basePnValues) Then
            preferredBase = Trim$(CStr(basePnValues(rowIndex, 1)))
        End If

        If Len(preferredBase) > 0 And Len(assemblyNo) >= Len(preferredBase) _
            And StrComp(Left$(assemblyNo, Len(preferredBase)), preferredBase, vbTextCompare) = 0 Then
            basePartCode = preferredBase
            ResolveSeparatorFromAssembly assemblyNo, preferredBase, separator, dashCondition
        Else
            SplitAssemblyNoWithSeparator assemblyNo, basePartCode, separator, dashCondition
        End If

        If Len(basePartCode) = 0 Or Len(dashCondition) = 0 Then GoTo ContinueRccp
        If Len(separator) = 0 Then separator = "-"

        keyValue = DashRowKey(basePartCode, dashCondition)
        If Not rccpAssemblies.Exists(keyValue) Then
            rccpAssemblies.Add keyValue, separator
        End If

ContinueRccp:
    Next rowIndex

    For Each dashKey In rccpAssemblies.Keys
        ParseDashRowKey CStr(dashKey), basePartCode, dashCondition
        separator = CStr(rccpAssemblies(dashKey))
        listRowIndex = FindJunctionListRow(dashTbl, COL_BASE_PART_CODE, basePartCode, COL_DASH_CONDITION, dashCondition)

        If listRowIndex = 0 Then
            Set fieldValues = NewFieldValuesDictionary()
            fieldValues(COL_BASE_PART_CODE) = basePartCode
            fieldValues(COL_DASH_CONDITION) = dashCondition
            fieldValues(COL_SEPARATOR) = separator
            fieldValues(COL_ACTIVE) = True
            UpsertJunctionRow dashTbl, COL_BASE_PART_CODE, basePartCode, COL_DASH_CONDITION, dashCondition, fieldValues
            WriteDashConditionText dashTbl, basePartCode, dashCondition
            addedCount = addedCount + 1
        Else
            If Not IsActiveFlag(GetCellValueByListRow(dashTbl, listRowIndex, COL_ACTIVE)) Then
                SetCellValueByListRow dashTbl, listRowIndex, COL_ACTIVE, True
                activatedCount = activatedCount + 1
            End If
            If TableHasColumn(dashTbl, COL_SEPARATOR) Then
                If Len(Trim$(CStr(Nz(GetCellValueByListRow(dashTbl, listRowIndex, COL_SEPARATOR))))) = 0 _
                    Or StrComp(CStr(Nz(GetCellValueByListRow(dashTbl, listRowIndex, COL_SEPARATOR))), separator, vbBinaryCompare) <> 0 Then
                    SetCellValueByListRow dashTbl, listRowIndex, COL_SEPARATOR, separator
                End If
            End If
            WriteDashConditionText dashTbl, basePartCode, dashCondition
        End If
    Next dashKey

    If Not dashTbl.DataBodyRange Is Nothing Then
        For listRowIndex = 1 To dashTbl.ListRows.Count
            basePartCode = Trim$(CStr(Nz(GetCellValueByListRow(dashTbl, listRowIndex, COL_BASE_PART_CODE))))
            dashCondition = DashConditionTextFromRow(dashTbl, listRowIndex)
            If Len(basePartCode) = 0 Or Len(dashCondition) = 0 Then GoTo ContinueExisting

            keyValue = DashRowKey(basePartCode, dashCondition)
            If Not rccpAssemblies.Exists(keyValue) Then
                If IsActiveFlag(GetCellValueByListRow(dashTbl, listRowIndex, COL_ACTIVE)) Then
                    SetCellValueByListRow dashTbl, listRowIndex, COL_ACTIVE, False
                    inactivatedCount = inactivatedCount + 1
                End If
            End If

ContinueExisting:
        Next listRowIndex
    End If

    SyncPartDashConditionsFromRCCP = _
        "Dash sync: added " & addedCount & _
        ", reactivated " & activatedCount & _
        ", inactivated " & inactivatedCount & "."
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
    Dim separator As String
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

        dashCondition = DashConditionTextFromRow(dashTbl, dashRowIndex)
        If Len(dashCondition) = 0 Then GoTo ContinueDash

        separator = "-"
        If TableHasColumn(dashTbl, COL_SEPARATOR) Then
            separator = Trim$(CStr(Nz(GetCellValueByListRow(dashTbl, dashRowIndex, COL_SEPARATOR))))
            If Len(separator) = 0 Then separator = "-"
        End If

        assemblyNo = BuildAssemblyNo(basePartCode, separator, dashCondition)
        If Not assemblyList.Exists(assemblyNo) Then assemblyList.Add assemblyNo, assemblyNo

ContinueDash:
    Next dashRowIndex
End Sub

Private Sub ResolveSeparatorFromAssembly( _
    ByVal assemblyNo As String, _
    ByVal basePartCode As String, _
    ByRef separator As String, _
    ByRef dashCondition As String)

    Dim remainder As String

    remainder = Mid$(assemblyNo, Len(basePartCode) + 1)
    If Len(remainder) = 0 Then Exit Sub

    If Left$(remainder, 1) = "-" Then
        separator = "-"
        dashCondition = Mid$(remainder, 2)
    ElseIf IsLetterCharacter(Left$(remainder, 1)) Then
        separator = Left$(remainder, 1)
        dashCondition = Mid$(remainder, 2)
    Else
        separator = "-"
        dashCondition = remainder
    End If
End Sub

Private Function DashRowKey(ByVal basePartCode As String, ByVal dashCondition As String) As String
    DashRowKey = basePartCode & vbTab & dashCondition
End Function

Private Sub ParseDashRowKey(ByVal keyValue As String, ByRef basePartCode As String, ByRef dashCondition As String)
    Dim tabPos As Long

    tabPos = InStr(1, keyValue, vbTab, vbBinaryCompare)
    If tabPos = 0 Then
        basePartCode = keyValue
        dashCondition = vbNullString
    Else
        basePartCode = Left$(keyValue, tabPos - 1)
        dashCondition = Mid$(keyValue, tabPos + 1)
    End If
End Sub

Private Function ColumnTextValues(ByVal col As ListColumn) As Variant
    Dim values As Variant
    Dim result() As Variant
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim cell As Range

    If col.DataBodyRange Is Nothing Then Exit Function

    rowCount = col.DataBodyRange.Rows.Count
    ReDim result(1 To rowCount, 1 To 1)

    ' Prefer displayed text so leading zeros survive numeric cells.
    For rowIndex = 1 To rowCount
        Set cell = col.DataBodyRange.Cells(rowIndex, 1)
        If Len(Trim$(cell.Text)) > 0 Then
            result(rowIndex, 1) = Trim$(cell.Text)
        Else
            result(rowIndex, 1) = Trim$(CStr(Nz(cell.Value2)))
        End If
    Next rowIndex

    ColumnTextValues = result
End Function

Private Function DashConditionTextFromRow(ByVal dashTbl As ListObject, ByVal listRowIndex As Long) As String
    Dim cell As Range

    Set cell = dashTbl.ListRows(listRowIndex).Range.Cells(1, TableColumnIndex(dashTbl, COL_DASH_CONDITION))
    If Len(Trim$(cell.Text)) > 0 Then
        DashConditionTextFromRow = CStr(cell.Text)
    Else
        DashConditionTextFromRow = CStr(Nz(cell.Value2))
    End If
End Function

Private Sub WriteDashConditionText( _
    ByVal dashTbl As ListObject, _
    ByVal basePartCode As String, _
    ByVal dashCondition As String)

    Dim listRowIndex As Long
    Dim cell As Range

    listRowIndex = FindJunctionListRow(dashTbl, COL_BASE_PART_CODE, basePartCode, COL_DASH_CONDITION, dashCondition)
    If listRowIndex = 0 Then Exit Sub

    Set cell = dashTbl.ListRows(listRowIndex).Range.Cells(1, TableColumnIndex(dashTbl, COL_DASH_CONDITION))
    cell.NumberFormat = "@"
    cell.Value = dashCondition
End Sub

Private Sub EnsureDashConditionTextFormat(ByVal dashTbl As ListObject)
    Dim col As ListColumn

    If Not TableHasColumn(dashTbl, COL_DASH_CONDITION) Then Exit Sub
    Set col = dashTbl.ListColumns(COL_DASH_CONDITION)
    col.Range.NumberFormat = "@"
    If Not col.DataBodyRange Is Nothing Then col.DataBodyRange.NumberFormat = "@"
End Sub

Private Function IsDigitCharacter(ByVal character As String) As Boolean
    If Len(character) = 0 Then Exit Function
    IsDigitCharacter = (character >= "0" And character <= "9")
End Function

Private Function IsLetterCharacter(ByVal character As String) As Boolean
    Dim upperChar As String

    If Len(character) = 0 Then Exit Function
    upperChar = UCase$(character)
    IsLetterCharacter = (upperChar >= "A" And upperChar <= "Z")
End Function

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
