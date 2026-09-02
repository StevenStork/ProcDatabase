Attribute VB_Name = "modTableIO"
Option Explicit

'==============================================================================
' Generic ListObject access for the capacity database tables.
'==============================================================================

Public Function FindTable(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim tbl As ListObject

    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set tbl = ws.ListObjects(tableName)
        On Error GoTo 0

        If Not tbl Is Nothing Then
            Set FindTable = tbl
            Exit Function
        End If
    Next ws
End Function

Public Function TableHasColumn(ByVal tbl As ListObject, ByVal columnName As String) As Boolean
    Dim col As ListColumn

    On Error Resume Next
    Set col = tbl.ListColumns(columnName)
    On Error GoTo 0

    TableHasColumn = Not col Is Nothing
End Function

Public Function TableColumnIndex(ByVal tbl As ListObject, ByVal columnName As String) As Long
    Dim col As ListColumn

    On Error GoTo Fail
    Set col = tbl.ListColumns(columnName)
    TableColumnIndex = col.Index
    Exit Function

Fail:
    TableColumnIndex = 0
End Function

Public Function FindListRowByKey( _
    ByVal tbl As ListObject, _
    ByVal keyColumnName As String, _
    ByVal keyValue As String) As Long

    Dim keyValues As Variant
    Dim rowIndex As Long
    Dim rowCount As Long

    FindListRowByKey = 0
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    If Not TableHasColumn(tbl, keyColumnName) Then Exit Function

    keyValues = tbl.ListColumns(keyColumnName).DataBodyRange.Value2
    If Not IsArray(keyValues) Then
        If ValuesMatchCode(keyValues, keyValue) Then
            FindListRowByKey = tbl.ListRows(1).Index
        End If
        Exit Function
    End If

    rowCount = UBound(keyValues, 1)
    For rowIndex = 1 To rowCount
        If Len(NormalizeCode(keyValues(rowIndex, 1))) = 0 Then GoTo ContinueKeyRow
        If ValuesMatchCode(keyValues(rowIndex, 1), keyValue) Then
            FindListRowByKey = tbl.ListRows(rowIndex).Index
            Exit Function
        End If
ContinueKeyRow:
    Next rowIndex
End Function

Public Function FindJunctionListRow( _
    ByVal tbl As ListObject, _
    ByVal key1ColumnName As String, _
    ByVal key1Value As String, _
    ByVal key2ColumnName As String, _
    ByVal key2Value As String) As Long

    Dim key1Values As Variant
    Dim key2Values As Variant
    Dim rowIndex As Long
    Dim rowCount As Long

    FindJunctionListRow = 0
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    key1Values = tbl.ListColumns(key1ColumnName).DataBodyRange.Value2
    key2Values = tbl.ListColumns(key2ColumnName).DataBodyRange.Value2

    If Not IsArray(key1Values) Then
        If ValuesMatchCode(key1Values, key1Value) And ValuesMatchCode(key2Values, key2Value) Then
            FindJunctionListRow = tbl.ListRows(1).Index
        End If
        Exit Function
    End If

    rowCount = UBound(key1Values, 1)
    For rowIndex = 1 To rowCount
        If ValuesMatchCode(key1Values(rowIndex, 1), key1Value) _
            And ValuesMatchCode(key2Values(rowIndex, 1), key2Value) Then
            FindJunctionListRow = tbl.ListRows(rowIndex).Index
            Exit Function
        End If
    Next rowIndex
End Function

Public Function JunctionExists( _
    ByVal tbl As ListObject, _
    ByVal key1ColumnName As String, _
    ByVal key1Value As String, _
    ByVal key2ColumnName As String, _
    ByVal key2Value As String) As Boolean

    JunctionExists = (FindJunctionListRow(tbl, key1ColumnName, key1Value, key2ColumnName, key2Value) > 0)
End Function

Public Function GetCellValueByListRow(ByVal tbl As ListObject, ByVal listRowIndex As Long, ByVal columnName As String) As Variant
    Dim lr As ListRow

    On Error GoTo Fail
    Set lr = tbl.ListRows(listRowIndex)
    GetCellValueByListRow = lr.Range.Cells(1, TableColumnIndex(tbl, columnName)).Value2
    Exit Function

Fail:
    GetCellValueByListRow = Empty
End Function

Public Sub SetCellValueByListRow( _
    ByVal tbl As ListObject, _
    ByVal listRowIndex As Long, _
    ByVal columnName As String, _
    ByVal newValue As Variant)

    Dim lr As ListRow

    Set lr = tbl.ListRows(listRowIndex)
    lr.Range.Cells(1, TableColumnIndex(tbl, columnName)).Value = newValue
End Sub

Public Sub UpsertRow( _
    ByVal tbl As ListObject, _
    ByVal keyColumnName As String, _
    ByVal keyValue As String, _
    ByVal fieldValues As Object)

    Dim listRowIndex As Long
    Dim colName As Variant

    listRowIndex = FindListRowByKey(tbl, keyColumnName, keyValue)
    If listRowIndex = 0 Then
        listRowIndex = GetOrCreateListRowIndex(tbl)
    End If

    For Each colName In fieldValues.Keys
        SetCellValueByListRow tbl, listRowIndex, CStr(colName), fieldValues(colName)
    Next colName
End Sub

Public Sub DeleteRowByKey( _
    ByVal tbl As ListObject, _
    ByVal keyColumnName As String, _
    ByVal keyValue As String)

    Dim listRowIndex As Long

    listRowIndex = FindListRowByKey(tbl, keyColumnName, keyValue)
    If listRowIndex > 0 Then
        tbl.ListRows(listRowIndex).Delete
    End If
End Sub

Public Sub DeleteJunctionRow( _
    ByVal tbl As ListObject, _
    ByVal key1ColumnName As String, _
    ByVal key1Value As String, _
    ByVal key2ColumnName As String, _
    ByVal key2Value As String)

    Dim listRowIndex As Long

    listRowIndex = FindJunctionListRow(tbl, key1ColumnName, key1Value, key2ColumnName, key2Value)
    If listRowIndex > 0 Then
        tbl.ListRows(listRowIndex).Delete
    End If
End Sub

Public Sub UpsertJunctionRow( _
    ByVal tbl As ListObject, _
    ByVal key1ColumnName As String, _
    ByVal key1Value As String, _
    ByVal key2ColumnName As String, _
    ByVal key2Value As String, _
    ByVal fieldValues As Object)

    Dim listRowIndex As Long
    Dim colName As Variant

    listRowIndex = FindJunctionListRow(tbl, key1ColumnName, key1Value, key2ColumnName, key2Value)
    If listRowIndex = 0 Then
        listRowIndex = GetOrCreateListRowIndex(tbl)
    End If

    For Each colName In fieldValues.Keys
        SetCellValueByListRow tbl, listRowIndex, CStr(colName), fieldValues(colName)
    Next colName
End Sub

Public Function ListAllDisplayItems( _
    ByVal tbl As ListObject, _
    ByVal keyColumnName As String, _
    ByVal displayColumnName As String, _
    ByVal activeOnly As Boolean) As Variant

    Dim items() As String
    Dim itemCount As Long
    Dim keyValues As Variant
    Dim displayValues As Variant
    Dim activeValues As Variant
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim displayText As String

    itemCount = 0
    ReDim items(0 To 0)

    If tbl Is Nothing Then
        ListAllDisplayItems = items
        Exit Function
    End If

    If tbl.DataBodyRange Is Nothing Then
        ListAllDisplayItems = items
        Exit Function
    End If

    keyValues = tbl.ListColumns(keyColumnName).DataBodyRange.Value2
    displayValues = tbl.ListColumns(displayColumnName).DataBodyRange.Value2
    If activeOnly And TableHasColumn(tbl, COL_ACTIVE) Then
        activeValues = tbl.ListColumns(COL_ACTIVE).DataBodyRange.Value2
    End If

    If Not IsArray(keyValues) Then
        If Len(NormalizeCode(keyValues)) = 0 Then
            ListAllDisplayItems = items
            Exit Function
        End If
        If Not activeOnly Or IsActiveFlag(activeValues) Then
            displayText = BuildDisplayItem(keyValues, displayValues)
            AppendDisplayItem items, itemCount, displayText
        End If
        ListAllDisplayItems = items
        Exit Function
    End If

    rowCount = UBound(keyValues, 1)
    For rowIndex = 1 To rowCount
        If Len(NormalizeCode(keyValues(rowIndex, 1))) = 0 Then GoTo ContinueRow
        If activeOnly Then
            If Not IsActiveFlag(activeValues(rowIndex, 1)) Then GoTo ContinueRow
        End If

        displayText = BuildDisplayItem(keyValues(rowIndex, 1), displayValues(rowIndex, 1))
        AppendDisplayItem items, itemCount, displayText

ContinueRow:
    Next rowIndex

    ListAllDisplayItems = items
End Function

Public Function ListJunctionDisplayItems( _
    ByVal tbl As ListObject, _
    ByVal filterColumnName As String, _
    ByVal filterValue As String, _
    ByVal displayColumnName As String) As Variant

    Dim items() As String
    Dim itemCount As Long
    Dim filterValues As Variant
    Dim displayValues As Variant
    Dim rowIndex As Long
    Dim rowCount As Long

    itemCount = 0
    ReDim items(0 To 0)

    If tbl Is Nothing Then
        ListJunctionDisplayItems = items
        Exit Function
    End If

    If tbl.DataBodyRange Is Nothing Then
        ListJunctionDisplayItems = items
        Exit Function
    End If

    filterValues = tbl.ListColumns(filterColumnName).DataBodyRange.Value2
    displayValues = tbl.ListColumns(displayColumnName).DataBodyRange.Value2

    If Not IsArray(filterValues) Then
        If ValuesMatchCode(filterValues, filterValue) Then
            AppendDisplayItem items, itemCount, CStr(Nz(displayValues))
        End If
        ListJunctionDisplayItems = items
        Exit Function
    End If

    rowCount = UBound(filterValues, 1)
    For rowIndex = 1 To rowCount
        If ValuesMatchCode(filterValues(rowIndex, 1), filterValue) Then
            AppendDisplayItem items, itemCount, CStr(Nz(displayValues(rowIndex, 1)))
        End If
    Next rowIndex

    ListJunctionDisplayItems = items
End Function

Public Function ExtractCodeFromDisplayItem(ByVal displayItem As String) As String
    Dim separatorPos As Long

    separatorPos = InStr(1, displayItem, " - ", vbBinaryCompare)
    If separatorPos > 0 Then
        ExtractCodeFromDisplayItem = Trim$(Left$(displayItem, separatorPos - 1))
    Else
        ExtractCodeFromDisplayItem = Trim$(displayItem)
    End If
End Function

Public Function LookupDisplayName( _
    ByVal tbl As ListObject, _
    ByVal keyColumnName As String, _
    ByVal keyValue As String, _
    ByVal displayColumnName As String) As String

    Dim listRowIndex As Long

    listRowIndex = FindListRowByKey(tbl, keyColumnName, keyValue)
    If listRowIndex = 0 Then
        LookupDisplayName = vbNullString
    Else
        LookupDisplayName = Trim$(CStr(Nz(GetCellValueByListRow(tbl, listRowIndex, displayColumnName))))
    End If
End Function

Public Function ValuesMatchCode(ByVal leftValue As Variant, ByVal rightValue As Variant) As Boolean
    ValuesMatchCode = (StrComp(NormalizeCode(leftValue), NormalizeCode(rightValue), vbTextCompare) = 0)
End Function

Private Function BuildDisplayItem(ByVal keyValue As Variant, ByVal displayValue As Variant) As String
    Dim codeText As String
    Dim nameText As String

    codeText = NormalizeCode(keyValue)
    nameText = Trim$(CStr(Nz(displayValue)))

    If Len(nameText) > 0 Then
        BuildDisplayItem = codeText & " - " & nameText
    Else
        BuildDisplayItem = codeText
    End If
End Function

Private Sub AppendDisplayItem(ByRef items() As String, ByRef itemCount As Long, ByVal displayText As String)
    If Len(displayText) = 0 Then Exit Sub

    If itemCount = 0 And (Not IsArrayAllocated(items) Or UBound(items) < LBound(items)) Then
        ReDim items(0 To 0)
    ElseIf itemCount > 0 Then
        ReDim Preserve items(0 To itemCount)
    End If

    items(itemCount) = displayText
    itemCount = itemCount + 1
End Sub

Private Function IsArrayAllocated(ByVal arr As Variant) As Boolean
    On Error Resume Next
    IsArrayAllocated = IsArray(arr) And (UBound(arr) >= LBound(arr))
    On Error GoTo 0
End Function

Public Sub BindComboBoxFromTable( _
    ByVal comboBox As Object, _
    ByVal tbl As ListObject, _
    ByVal keyColumnName As String, _
    ByVal displayColumnName As String, _
    ByVal activeOnly As Boolean)

    Dim items As Variant
    Dim itemIndex As Long

    comboBox.Clear
    items = ListAllDisplayItems(tbl, keyColumnName, displayColumnName, activeOnly)

    If Not IsArrayAllocated(items) Then Exit Sub
    If UBound(items) < LBound(items) Then Exit Sub

    For itemIndex = LBound(items) To UBound(items)
        If Len(items(itemIndex)) > 0 Then
            comboBox.AddItem items(itemIndex)
        End If
    Next itemIndex
End Sub

Public Sub BindListBoxFromArray(ByVal listBox As Object, ByVal items As Variant)
    Dim itemIndex As Long

    listBox.Clear

    If IsEmpty(items) Then Exit Sub
    If Not IsArrayAllocated(items) Then Exit Sub
    If UBound(items) < LBound(items) Then Exit Sub

    For itemIndex = LBound(items) To UBound(items)
        If Len(items(itemIndex)) > 0 Then
            listBox.AddItem items(itemIndex)
        End If
    Next itemIndex
End Sub

Public Function NewFieldValuesDictionary() As Object
    Set NewFieldValuesDictionary = CreateObject("Scripting.Dictionary")
    NewFieldValuesDictionary.CompareMode = vbTextCompare
End Function

Public Function GetOrCreateListRowIndex(ByVal tbl As ListObject) As Long
    If tbl Is Nothing Then Exit Function

    If tbl.DataBodyRange Is Nothing Then
        GetOrCreateListRowIndex = tbl.ListRows.Add.Index
        Exit Function
    End If

    If tbl.ListRows.Count = 1 And IsListRowEmpty(tbl, 1) Then
        GetOrCreateListRowIndex = tbl.ListRows(1).Index
        Exit Function
    End If

    GetOrCreateListRowIndex = tbl.ListRows.Add.Index
End Function

Public Sub DeleteEmptyTableRows(ByVal tbl As ListObject)
    Dim rowIndex As Long

    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = tbl.ListRows.Count To 1 Step -1
        If IsListRowEmpty(tbl, rowIndex) Then
            tbl.ListRows(rowIndex).Delete
        End If
    Next rowIndex
End Sub

Public Function IsListRowEmpty(ByVal tbl As ListObject, ByVal listRowIndex As Long) As Boolean
    Dim lr As ListRow
    Dim colIndex As Long
    Dim cellValue As Variant

    Set lr = tbl.ListRows(listRowIndex)

    For colIndex = 1 To tbl.ListColumns.Count
        cellValue = lr.Range.Cells(1, colIndex).Value2
        If Len(Trim$(CStr(Nz(cellValue)))) > 0 Then
            IsListRowEmpty = False
            Exit Function
        End If
    Next colIndex

    IsListRowEmpty = True
End Function
