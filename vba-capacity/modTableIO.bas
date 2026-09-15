Attribute VB_Name = "modTableIO"
Option Explicit

'==============================================================================
' Generic ListObject access for the capacity database tables.
'==============================================================================

Public Function FindTable(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim tbl As ListObject

    For Each ws In ThisWorkbook.Worksheets
        Set tbl = Nothing
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

Public Function FindWorksheetByName(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set FindWorksheetByName = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

' Always returns a 2-D array (rows x 1) or Empty when the table/column has no data.
Public Function ListColumnValues2D(ByVal tbl As ListObject, ByVal columnName As String) As Variant
    Dim values As Variant
    Dim wrapped(1 To 1, 1 To 1) As Variant

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    If Not TableHasColumn(tbl, columnName) Then Exit Function

    values = tbl.ListColumns(columnName).DataBodyRange.Value2
    If IsArray(values) Then
        ListColumnValues2D = values
    Else
        wrapped(1, 1) = values
        ListColumnValues2D = wrapped
    End If
End Function

Public Function TableColumnHasValue( _
    ByVal tbl As ListObject, _
    ByVal columnName As String, _
    ByVal matchValue As String) As Boolean

    Dim columnValues As Variant
    Dim rowIndex As Long

    columnValues = ListColumnValues2D(tbl, columnName)
    If Not IsArray(columnValues) Then Exit Function

    For rowIndex = 1 To UBound(columnValues, 1)
        If ValuesMatchCode(columnValues(rowIndex, 1), matchValue) Then
            TableColumnHasValue = True
            Exit Function
        End If
    Next rowIndex
End Function

Public Sub FormatListColumnAsText(ByVal tbl As ListObject, ByVal columnName As String)
    Dim col As ListColumn

    If tbl Is Nothing Then Exit Sub
    If Not TableHasColumn(tbl, columnName) Then Exit Sub

    Set col = tbl.ListColumns(columnName)
    col.Range.NumberFormat = "@"
    If Not col.DataBodyRange Is Nothing Then col.DataBodyRange.NumberFormat = "@"
End Sub

Public Function ListActiveKeyCodes(ByVal tbl As ListObject, ByVal keyColumnName As String) As Variant
    Dim items() As String
    Dim itemCount As Long
    Dim keyValues As Variant
    Dim rowIndex As Long
    Dim codeValue As String

    itemCount = 0
    ReDim items(0 To -1)

    keyValues = ListColumnValues2D(tbl, keyColumnName)
    If Not IsArray(keyValues) Then
        ListActiveKeyCodes = items
        Exit Function
    End If

    For rowIndex = 1 To UBound(keyValues, 1)
        codeValue = NormalizeCode(keyValues(rowIndex, 1))
        If Len(codeValue) = 0 Then GoTo ContinueActiveKey
        If TableHasColumn(tbl, COL_ACTIVE) Then
            If Not IsActiveFlag(GetCellValueByListRow(tbl, tbl.ListRows(rowIndex).Index, COL_ACTIVE)) Then GoTo ContinueActiveKey
        End If

        If itemCount = 0 Then
            ReDim items(0 To 0)
        Else
            ReDim Preserve items(0 To itemCount)
        End If
        items(itemCount) = codeValue
        itemCount = itemCount + 1

ContinueActiveKey:
    Next rowIndex

    ListActiveKeyCodes = items
End Function

Public Sub DeleteRowsMatchingKey( _
    ByVal tbl As ListObject, _
    ByVal keyColumnName As String, _
    ByVal keyValue As String)

    Dim rowIndex As Long

    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = tbl.ListRows.Count To 1 Step -1
        If ValuesMatchCode(GetCellValueByListRow(tbl, rowIndex, keyColumnName), keyValue) Then
            tbl.ListRows(rowIndex).Delete
        End If
    Next rowIndex
End Sub

Private Sub ApplyFieldValuesToListRow(ByVal tbl As ListObject, ByVal listRowIndex As Long, ByVal fieldValues As Object)
    Dim colName As Variant

    For Each colName In fieldValues.Keys
        SetCellValueByListRow tbl, listRowIndex, CStr(colName), fieldValues(colName)
    Next colName
End Sub

Private Sub BindControlFromItems(ByVal listControl As Object, ByVal items As Variant)
    Dim itemIndex As Long

    listControl.Clear
    If IsEmpty(items) Then Exit Sub
    If Not IsArrayAllocated(items) Then Exit Sub
    If UBound(items) < LBound(items) Then Exit Sub

    For itemIndex = LBound(items) To UBound(items)
        If Len(items(itemIndex)) > 0 Then
            listControl.AddItem items(itemIndex)
        End If
    Next itemIndex
End Sub

Public Function FindListRowByKey( _
    ByVal tbl As ListObject, _
    ByVal keyColumnName As String, _
    ByVal keyValue As String) As Long

    Dim keyValues As Variant
    Dim rowIndex As Long
    Dim rowCount As Long

    FindListRowByKey = 0
    keyValues = ListColumnValues2D(tbl, keyColumnName)
    If Not IsArray(keyValues) Then Exit Function

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
    key1Values = ListColumnValues2D(tbl, key1ColumnName)
    key2Values = ListColumnValues2D(tbl, key2ColumnName)
    If Not IsArray(key1Values) Or Not IsArray(key2Values) Then Exit Function

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

Public Function FindTripleJunctionListRow( _
    ByVal tbl As ListObject, _
    ByVal key1ColumnName As String, _
    ByVal key1Value As String, _
    ByVal key2ColumnName As String, _
    ByVal key2Value As String, _
    ByVal key3ColumnName As String, _
    ByVal key3Value As String) As Long

    Dim key1Values As Variant
    Dim key2Values As Variant
    Dim key3Values As Variant
    Dim rowIndex As Long
    Dim rowCount As Long

    FindTripleJunctionListRow = 0
    key1Values = ListColumnValues2D(tbl, key1ColumnName)
    key2Values = ListColumnValues2D(tbl, key2ColumnName)
    key3Values = ListColumnValues2D(tbl, key3ColumnName)
    If Not IsArray(key1Values) Or Not IsArray(key2Values) Or Not IsArray(key3Values) Then Exit Function

    rowCount = UBound(key1Values, 1)
    For rowIndex = 1 To rowCount
        If ValuesMatchCode(key1Values(rowIndex, 1), key1Value) _
            And ValuesMatchCode(key2Values(rowIndex, 1), key2Value) _
            And ValuesMatchCode(key3Values(rowIndex, 1), key3Value) Then
            FindTripleJunctionListRow = tbl.ListRows(rowIndex).Index
            Exit Function
        End If
    Next rowIndex
End Function

Public Sub DeleteTripleJunctionRow( _
    ByVal tbl As ListObject, _
    ByVal key1ColumnName As String, _
    ByVal key1Value As String, _
    ByVal key2ColumnName As String, _
    ByVal key2Value As String, _
    ByVal key3ColumnName As String, _
    ByVal key3Value As String)

    Dim listRowIndex As Long

    listRowIndex = FindTripleJunctionListRow( _
        tbl, key1ColumnName, key1Value, key2ColumnName, key2Value, key3ColumnName, key3Value)
    If listRowIndex > 0 Then
        tbl.ListRows(listRowIndex).Delete
    End If
End Sub

Public Sub UpsertTripleJunctionRow( _
    ByVal tbl As ListObject, _
    ByVal key1ColumnName As String, _
    ByVal key1Value As String, _
    ByVal key2ColumnName As String, _
    ByVal key2Value As String, _
    ByVal key3ColumnName As String, _
    ByVal key3Value As String, _
    ByVal fieldValues As Object)

    Dim listRowIndex As Long

    listRowIndex = FindTripleJunctionListRow( _
        tbl, key1ColumnName, key1Value, key2ColumnName, key2Value, key3ColumnName, key3Value)
    If listRowIndex = 0 Then
        listRowIndex = GetOrCreateListRowIndex(tbl)
    End If

    ApplyFieldValuesToListRow tbl, listRowIndex, fieldValues
End Sub

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

    listRowIndex = FindListRowByKey(tbl, keyColumnName, keyValue)
    If listRowIndex = 0 Then
        listRowIndex = GetOrCreateListRowIndex(tbl)
    End If

    ApplyFieldValuesToListRow tbl, listRowIndex, fieldValues
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

    listRowIndex = FindJunctionListRow(tbl, key1ColumnName, key1Value, key2ColumnName, key2Value)
    If listRowIndex = 0 Then
        listRowIndex = GetOrCreateListRowIndex(tbl)
    End If

    ApplyFieldValuesToListRow tbl, listRowIndex, fieldValues
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

    keyValues = ListColumnValues2D(tbl, keyColumnName)
    displayValues = ListColumnValues2D(tbl, displayColumnName)
    If Not IsArray(keyValues) Then
        ListAllDisplayItems = items
        Exit Function
    End If

    If activeOnly And TableHasColumn(tbl, COL_ACTIVE) Then
        activeValues = ListColumnValues2D(tbl, COL_ACTIVE)
    End If

    rowCount = UBound(keyValues, 1)
    For rowIndex = 1 To rowCount
        If Len(NormalizeCode(keyValues(rowIndex, 1))) = 0 Then GoTo ContinueRow
        If activeOnly And IsArray(activeValues) Then
            If Not IsActiveFlag(activeValues(rowIndex, 1)) Then GoTo ContinueRow
        End If

        If IsArray(displayValues) Then
            displayText = BuildDisplayItem(keyValues(rowIndex, 1), displayValues(rowIndex, 1))
        Else
            displayText = BuildDisplayItem(keyValues(rowIndex, 1), vbNullString)
        End If
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

    filterValues = ListColumnValues2D(tbl, filterColumnName)
    displayValues = ListColumnValues2D(tbl, displayColumnName)
    If Not IsArray(filterValues) Then
        ListJunctionDisplayItems = items
        Exit Function
    End If

    rowCount = UBound(filterValues, 1)
    For rowIndex = 1 To rowCount
        If ValuesMatchCode(filterValues(rowIndex, 1), filterValue) Then
            If IsArray(displayValues) Then
                AppendDisplayItem items, itemCount, CStr(Nz(displayValues(rowIndex, 1)))
            End If
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

Public Sub BindComboBoxFromTable( _
    ByVal comboBox As Object, _
    ByVal tbl As ListObject, _
    ByVal keyColumnName As String, _
    ByVal displayColumnName As String, _
    ByVal activeOnly As Boolean)

    BindControlFromItems comboBox, ListAllDisplayItems(tbl, keyColumnName, displayColumnName, activeOnly)
End Sub

Public Sub BindListBoxFromArray(ByVal listBox As Object, ByVal items As Variant)
    BindControlFromItems listBox, items
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
