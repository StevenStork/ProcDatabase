Attribute VB_Name = "modLinkedData"
Option Explicit

'==============================================================================
' Refresh linked Power Query connections (tblRCCP, etc.).
'
' tblRCCP FFA filter is driven inside Power Query by FactoriesTbl — see
' pqRCCP-FilteredFFAs.txt for the #"Filtered FFAs" M step to paste into the query.
' VBA only validates factory codes exist, then refreshes the connection.
'==============================================================================

Public Sub RefreshRCCP()
    Dim factoryCount As Long

    factoryCount = CountActiveFactoryCodes()
    If factoryCount = 0 Then
        MsgBox "Add at least one active FactoryCode to FactoriesTbl before refreshing tblRCCP.", vbExclamation
        Exit Sub
    End If

    On Error GoTo Fail
    OptimizeExcel True
    RefreshLinkedQuery LINKED_RCCP_TABLE
    OptimizeExcel False

    MsgBox "tblRCCP refreshed using " & factoryCount & " active factory code(s) from FactoriesTbl.", vbInformation
    Exit Sub

Fail:
    OptimizeExcel False
    MsgBox "Could not refresh tblRCCP: " & Err.Description, vbExclamation
End Sub

' Placeholder for a single Refresh Linked Data button (future: all queries).
Public Sub RefreshAllLinkedData()
    RefreshRCCP
End Sub

Public Function CountActiveFactoryCodes() As Long
    Dim tbl As ListObject
    Dim codes As Variant
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim codeValue As String

    Set tbl = FindTable(FACTORIES_TABLE_NAME)
    If tbl Is Nothing Or tbl.DataBodyRange Is Nothing Then Exit Function

    codes = tbl.ListColumns(COL_FACTORY_CODE).DataBodyRange.Value2
    If Not IsArray(codes) Then
        codeValue = NormalizeCode(codes)
        If Len(codeValue) > 0 Then
            If IsActiveFlag(GetCellValueByListRow(tbl, tbl.ListRows(1).Index, COL_ACTIVE)) Then
                CountActiveFactoryCodes = 1
            End If
        End If
        Exit Function
    End If

    rowCount = UBound(codes, 1)
    For rowIndex = 1 To rowCount
        codeValue = NormalizeCode(codes(rowIndex, 1))
        If Len(codeValue) = 0 Then GoTo ContinueRow
        If Not IsActiveFlag(GetCellValueByListRow(tbl, tbl.ListRows(rowIndex).Index, COL_ACTIVE)) Then GoTo ContinueRow
        CountActiveFactoryCodes = CountActiveFactoryCodes + 1

ContinueRow:
    Next rowIndex
End Function

Public Function BuildActiveFactoryCodeList() As String
    Dim tbl As ListObject
    Dim codes As Variant
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim codeValue As String
    Dim result As String

    Set tbl = FindTable(FACTORIES_TABLE_NAME)
    If tbl Is Nothing Or tbl.DataBodyRange Is Nothing Then Exit Function

    codes = tbl.ListColumns(COL_FACTORY_CODE).DataBodyRange.Value2
    If Not IsArray(codes) Then
        codeValue = NormalizeCode(codes)
        If Len(codeValue) > 0 Then
            If IsActiveFlag(GetCellValueByListRow(tbl, tbl.ListRows(1).Index, COL_ACTIVE)) Then
                BuildActiveFactoryCodeList = codeValue
            End If
        End If
        Exit Function
    End If

    rowCount = UBound(codes, 1)
    For rowIndex = 1 To rowCount
        codeValue = NormalizeCode(codes(rowIndex, 1))
        If Len(codeValue) = 0 Then GoTo ContinueRow
        If Not IsActiveFlag(GetCellValueByListRow(tbl, tbl.ListRows(rowIndex).Index, COL_ACTIVE)) Then GoTo ContinueRow

        If Len(result) > 0 Then result = result & ", "
        result = result & codeValue

ContinueRow:
    Next rowIndex

    BuildActiveFactoryCodeList = result
End Function

Public Sub RefreshLinkedQuery(ByVal queryOrConnectionName As String)
    Dim conn As WorkbookConnection
    Dim queryObj As Object

    Set conn = GetWorkbookConnection(queryOrConnectionName)
    If Not conn Is Nothing Then
        RefreshWorkbookConnection conn
        Exit Sub
    End If

    On Error Resume Next
    Set queryObj = ThisWorkbook.Queries(queryOrConnectionName)
    If Not queryObj Is Nothing Then
        queryObj.Refresh
        If Err.Number = 0 Then Exit Sub
    End If
    On Error GoTo 0

    Err.Raise vbObjectError + 530, "RefreshLinkedQuery", _
        "Connection or query '" & queryOrConnectionName & "' was not found."
End Sub

Private Function GetWorkbookConnection(ByVal connectionName As String) As WorkbookConnection
    On Error Resume Next
    Set GetWorkbookConnection = ThisWorkbook.Connections(connectionName)
    On Error GoTo 0
End Function

Private Sub RefreshWorkbookConnection(ByVal conn As WorkbookConnection)
    Select Case conn.Type
        Case xlConnectionTypeOLEDB
            conn.OLEDBConnection.BackgroundQuery = False
        Case xlConnectionTypeODBC
            conn.ODBCConnection.BackgroundQuery = False
    End Select

    conn.Refresh
End Sub
