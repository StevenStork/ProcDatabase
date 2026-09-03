Attribute VB_Name = "modLinkedData"
Option Explicit

'==============================================================================
' Refresh linked Power Query connections (tblRCCP, tblOperComps, etc.).
'
' tblRCCP FFA filter is driven inside Power Query by FactoriesTbl — see
' PowerQuery/pqRCCP-FilteredFFAs.txt for the #"Filtered FFAs" M step.
'
' tblOperComps: VBA rewrites only @assembly_number_list in the existing M Source
' SQL string using ASSEMBLY NO values from tblRCCP, then refreshes. The rest of
' the Power Query is left unchanged.
'==============================================================================

Private Const PARAM_ASSEMBLY_NUMBER_LIST As String = "@assembly_number_list"

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

Public Sub RefreshOperComps()
    Dim assemblyList As String
    Dim assemblyCount As Long

    assemblyList = BuildAssemblyNumberListFromRCCP()
    assemblyCount = CountCommaSeparatedItems(assemblyList)

    If assemblyCount = 0 Then
        MsgBox "No ASSEMBLY NO values found in tblRCCP." & vbCrLf & _
            "Refresh RCCP first and ensure tblRCCP is loaded to a sheet (ListObject), not connection-only.", _
            vbExclamation
        Exit Sub
    End If

    On Error GoTo Fail
    OptimizeExcel True
    UpdateQueryQuotedParameter LINKED_OPER_COMPS_TABLE, PARAM_ASSEMBLY_NUMBER_LIST, assemblyList
    RefreshLinkedQuery LINKED_OPER_COMPS_TABLE
    OptimizeExcel False

    MsgBox "tblOperComps refreshed with " & assemblyCount & " assembly number(s) from tblRCCP.", vbInformation
    Exit Sub

Fail:
    OptimizeExcel False
    MsgBox "Could not refresh tblOperComps: " & Err.Description, vbExclamation
End Sub

' Combined refresh button (extend as more queries are wired).
Public Sub RefreshAllLinkedData()
    RefreshRCCP
    RefreshOperComps
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

Public Function BuildAssemblyNumberListFromRCCP() As String
    Dim tbl As ListObject
    Dim assemblyValues As Variant
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim assemblyNo As String
    Dim uniqueAssemblies As Object
    Dim assemblyKey As Variant
    Dim result As String

    Set uniqueAssemblies = CreateObject("Scripting.Dictionary")
    uniqueAssemblies.CompareMode = vbTextCompare

    Set tbl = FindTable(LINKED_RCCP_TABLE)
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    If Not TableHasColumn(tbl, COL_ASSEMBLY_NO) Then Exit Function

    assemblyValues = tbl.ListColumns(COL_ASSEMBLY_NO).DataBodyRange.Value2
    If Not IsArray(assemblyValues) Then
        assemblyNo = Trim$(CStr(NzBlank(assemblyValues)))
        If Len(assemblyNo) > 0 Then uniqueAssemblies(assemblyNo) = assemblyNo
    Else
        rowCount = UBound(assemblyValues, 1)
        For rowIndex = 1 To rowCount
            assemblyNo = Trim$(CStr(NzBlank(assemblyValues(rowIndex, 1))))
            If Len(assemblyNo) = 0 Then GoTo ContinueAssembly
            If Not uniqueAssemblies.Exists(assemblyNo) Then uniqueAssemblies.Add assemblyNo, assemblyNo

ContinueAssembly:
        Next rowIndex
    End If

    For Each assemblyKey In uniqueAssemblies.Keys
        If Len(result) > 0 Then result = result & ", "
        result = result & EscapeSqlSingleQuotes(CStr(assemblyKey))
    Next assemblyKey

    BuildAssemblyNumberListFromRCCP = result
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

' Rewrites only parameterName = '...' inside the Power Query M Formula (or
' OLEDB/ODBC CommandText). Leaves the rest of the query unchanged.
Public Sub UpdateQueryQuotedParameter( _
    ByVal queryOrConnectionName As String, _
    ByVal parameterName As String, _
    ByVal parameterValue As String)

    Dim queryObj As Object
    Dim formulaText As String
    Dim conn As WorkbookConnection
    Dim commandText As String

    On Error Resume Next
    Set queryObj = ThisWorkbook.Queries(queryOrConnectionName)
    On Error GoTo 0

    If Not queryObj Is Nothing Then
        formulaText = CStr(queryObj.Formula)
        formulaText = ReplaceQuotedParameterValue(formulaText, parameterName, parameterValue, queryOrConnectionName)
        queryObj.Formula = formulaText
        Exit Sub
    End If

    Set conn = GetWorkbookConnection(queryOrConnectionName)
    If conn Is Nothing Then
        Err.Raise vbObjectError + 531, "UpdateQueryQuotedParameter", _
            "Query or connection '" & queryOrConnectionName & "' was not found."
    End If

    commandText = GetConnectionCommandText(conn)
    commandText = ReplaceQuotedParameterValue(commandText, parameterName, parameterValue, queryOrConnectionName)
    SetConnectionCommandText conn, commandText
End Sub

Private Function ReplaceQuotedParameterValue( _
    ByVal commandText As String, _
    ByVal parameterName As String, _
    ByVal parameterValue As String, _
    ByVal connectionName As String) As String

    Dim parameterPos As Long
    Dim equalsPos As Long
    Dim openingQuotePos As Long
    Dim closingQuotePos As Long

    ' Prefer the longest exact match so @assembly_number does not match
    ' @assembly_number_list when both appear in the same command text.
    parameterPos = FindParameterPosition(commandText, parameterName)
    If parameterPos = 0 Then
        Err.Raise vbObjectError + 516, "ReplaceQuotedParameterValue", _
            "Parameter '" & parameterName & "' was not found in " & connectionName & "."
    End If

    equalsPos = InStr(parameterPos, commandText, "=")
    If equalsPos = 0 Then
        Err.Raise vbObjectError + 517, "ReplaceQuotedParameterValue", _
            "Could not locate the assignment for '" & parameterName & "' in " & connectionName & "."
    End If

    openingQuotePos = InStr(equalsPos, commandText, "'")
    If openingQuotePos = 0 Then
        Err.Raise vbObjectError + 518, "ReplaceQuotedParameterValue", _
            "Could not locate the opening quote for '" & parameterName & "' in " & connectionName & "."
    End If

    closingQuotePos = InStr(openingQuotePos + 1, commandText, "'")
    If closingQuotePos = 0 Then
        Err.Raise vbObjectError + 519, "ReplaceQuotedParameterValue", _
            "Could not locate the closing quote for '" & parameterName & "' in " & connectionName & "."
    End If

    ReplaceQuotedParameterValue = Left$(commandText, openingQuotePos) & parameterValue & Mid$(commandText, closingQuotePos)
End Function

Private Function FindParameterPosition(ByVal commandText As String, ByVal parameterName As String) As Long
    Dim searchPos As Long
    Dim foundPos As Long
    Dim nextChar As String

    searchPos = 1

    Do
        foundPos = InStr(searchPos, commandText, parameterName, vbTextCompare)
        If foundPos = 0 Then Exit Do

        nextChar = Mid$(commandText, foundPos + Len(parameterName), 1)
        If Len(nextChar) = 0 Or Not IsParameterNameContinuation(nextChar) Then
            FindParameterPosition = foundPos
            Exit Function
        End If

        searchPos = foundPos + 1
    Loop
End Function

Private Function IsParameterNameContinuation(ByVal character As String) As Boolean
    Select Case LCase$(character)
        Case "a" To "z", "0" To "9", "_", "@"
            IsParameterNameContinuation = True
        Case Else
            IsParameterNameContinuation = False
    End Select
End Function

Private Function CountCommaSeparatedItems(ByVal csvText As String) As Long
    Dim parts As Variant
    Dim i As Long

    If Len(Trim$(csvText)) = 0 Then Exit Function

    parts = Split(csvText, ",")
    For i = LBound(parts) To UBound(parts)
        If Len(Trim$(CStr(parts(i)))) > 0 Then
            CountCommaSeparatedItems = CountCommaSeparatedItems + 1
        End If
    Next i
End Function

Private Function EscapeSqlSingleQuotes(ByVal textValue As String) As String
    EscapeSqlSingleQuotes = Replace(textValue, "'", "''")
End Function

Private Function NzBlank(ByVal value As Variant) As Variant
    If IsError(value) Then
        NzBlank = vbNullString
    ElseIf IsEmpty(value) Or IsNull(value) Then
        NzBlank = vbNullString
    Else
        NzBlank = value
    End If
End Function

Private Function GetWorkbookConnection(ByVal connectionName As String) As WorkbookConnection
    On Error Resume Next
    Set GetWorkbookConnection = ThisWorkbook.Connections(connectionName)
    On Error GoTo 0
End Function

Private Function GetConnectionCommandText(ByVal conn As WorkbookConnection) As String
    Select Case conn.Type
        Case xlConnectionTypeOLEDB
            GetConnectionCommandText = conn.OLEDBConnection.CommandText
        Case xlConnectionTypeODBC
            GetConnectionCommandText = conn.ODBCConnection.CommandText
        Case Else
            Err.Raise vbObjectError + 520, "GetConnectionCommandText", _
                "Connection '" & conn.Name & "' uses an unsupported connection type."
    End Select
End Function

Private Sub SetConnectionCommandText(ByVal conn As WorkbookConnection, ByVal commandText As String)
    Select Case conn.Type
        Case xlConnectionTypeOLEDB
            conn.OLEDBConnection.CommandText = commandText
        Case xlConnectionTypeODBC
            conn.ODBCConnection.CommandText = commandText
        Case Else
            Err.Raise vbObjectError + 521, "SetConnectionCommandText", _
                "Connection '" & conn.Name & "' uses an unsupported connection type."
    End Select
End Sub

Private Sub RefreshWorkbookConnection(ByVal conn As WorkbookConnection)
    Select Case conn.Type
        Case xlConnectionTypeOLEDB
            conn.OLEDBConnection.BackgroundQuery = False
        Case xlConnectionTypeODBC
            conn.ODBCConnection.BackgroundQuery = False
    End Select

    conn.Refresh
End Sub
