Attribute VB_Name = "modLinkedData"
Option Explicit

'==============================================================================
' Refresh linked Power Query connections (tblRCCP, tblOperComps, tblAssyStnd,
' tblRouteCard).
'
' tblRCCP FFA filter is driven inside Power Query by FactoriesTbl — see
' PowerQuery/pqRCCP-FilteredFFAs.txt for the #"Filtered FFAs" M step.
'
' tblOperComps / tblAssyStnd / tblRouteCard: VBA updates @ffa in the Source SQL
' only when the active factory code list from FactoriesTbl has changed. Refresh
' alone does not rewrite the query, so Power Query permission prompts stay rare.
'
' tblAssyStnd / tblRouteCard also filter ASSEMBLY NO in Power Query via
' #"Filter Assemblies" against tblRCCP — see PowerQuery/pqAssyStnd-FilterAssemblies.txt
' and PowerQuery/pqRouteCard-FilterAssemblies.txt.
'==============================================================================

Private Const PARAM_FFA As String = "@ffa"

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
    RefreshQueryWithFfaParameter LINKED_OPER_COMPS_TABLE, "tblOperComps"
End Sub

Public Sub RefreshAssyStnd()
    RefreshQueryWithFfaParameter LINKED_ASSY_STND_TABLE, "tblAssyStnd"
End Sub

Public Sub RefreshRouteCard()
    RefreshQueryWithFfaParameter LINKED_ROUTE_CARD_TABLE, "tblRouteCard"
End Sub

' Combined refresh button (extend as more queries are wired).
' RCCP first so #"Filter Assemblies" steps see current assemblies.
Public Sub RefreshAllLinkedData()
    RefreshRCCP
    RefreshOperComps
    RefreshAssyStnd
    RefreshRouteCard
End Sub

Private Sub RefreshQueryWithFfaParameter(ByVal queryName As String, ByVal displayName As String)
    Dim ffaList As String
    Dim currentFfa As String
    Dim ffaCount As Long
    Dim formulaUpdated As Boolean

    ffaList = BuildActiveFactoryCodeList()
    ffaCount = CountCommaSeparatedItems(ffaList)

    If ffaCount = 0 Then
        MsgBox "Add at least one active FactoryCode to FactoriesTbl before refreshing " & displayName & ".", vbExclamation
        Exit Sub
    End If

    On Error GoTo Fail
    OptimizeExcel True

    currentFfa = GetQuotedParameterValue(queryName, PARAM_FFA)
    If Not FfaListsMatch(currentFfa, ffaList) Then
        UpdateQueryQuotedParameter queryName, PARAM_FFA, ffaList
        formulaUpdated = True
    End If

    RefreshLinkedQuery queryName
    OptimizeExcel False

    If formulaUpdated Then
        MsgBox displayName & " @ffa updated to '" & ffaList & "' and refreshed.", vbInformation
    Else
        MsgBox displayName & " refreshed ( @ffa unchanged: '" & ffaList & "' ).", vbInformation
    End If
    Exit Sub

Fail:
    OptimizeExcel False
    MsgBox "Could not refresh " & displayName & ": " & Err.Description, vbExclamation
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

Public Function GetQuotedParameterValue( _
    ByVal queryOrConnectionName As String, _
    ByVal parameterName As String) As String

    Dim queryObj As Object
    Dim formulaText As String
    Dim conn As WorkbookConnection
    Dim commandText As String
    Dim parameterPos As Long
    Dim equalsPos As Long
    Dim openingQuotePos As Long
    Dim closingQuotePos As Long

    On Error Resume Next
    Set queryObj = ThisWorkbook.Queries(queryOrConnectionName)
    On Error GoTo 0

    If Not queryObj Is Nothing Then
        formulaText = CStr(queryObj.Formula)
    Else
        Set conn = GetWorkbookConnection(queryOrConnectionName)
        If conn Is Nothing Then
            Err.Raise vbObjectError + 532, "GetQuotedParameterValue", _
                "Query or connection '" & queryOrConnectionName & "' was not found."
        End If
        formulaText = GetConnectionCommandText(conn)
    End If

    parameterPos = FindParameterPosition(formulaText, parameterName)
    If parameterPos = 0 Then
        Err.Raise vbObjectError + 516, "GetQuotedParameterValue", _
            "Parameter '" & parameterName & "' was not found in " & queryOrConnectionName & "."
    End If

    equalsPos = InStr(parameterPos, formulaText, "=")
    openingQuotePos = InStr(equalsPos, formulaText, "'")
    closingQuotePos = InStr(openingQuotePos + 1, formulaText, "'")

    If equalsPos = 0 Or openingQuotePos = 0 Or closingQuotePos = 0 Then
        Err.Raise vbObjectError + 517, "GetQuotedParameterValue", _
            "Could not read quoted value for '" & parameterName & "' in " & queryOrConnectionName & "."
    End If

    GetQuotedParameterValue = Mid$(formulaText, openingQuotePos + 1, closingQuotePos - openingQuotePos - 1)
End Function

Private Function FfaListsMatch(ByVal leftList As String, ByVal rightList As String) As Boolean
    FfaListsMatch = (StrComp(CanonicalizeCodeList(leftList), CanonicalizeCodeList(rightList), vbTextCompare) = 0)
End Function

Private Function CanonicalizeCodeList(ByVal csvText As String) As String
    Dim parts As Variant
    Dim i As Long
    Dim j As Long
    Dim tempValue As String
    Dim codes() As String
    Dim codeCount As Long
    Dim result As String

    If Len(Trim$(csvText)) = 0 Then Exit Function

    parts = Split(csvText, ",")
    ReDim codes(0 To UBound(parts) - LBound(parts))
    codeCount = 0

    For i = LBound(parts) To UBound(parts)
        tempValue = NormalizeCode(parts(i))
        If Len(tempValue) = 0 Then GoTo ContinuePart
        codes(codeCount) = tempValue
        codeCount = codeCount + 1

ContinuePart:
    Next i

    If codeCount = 0 Then Exit Function

    ReDim Preserve codes(0 To codeCount - 1)

    For i = 0 To codeCount - 2
        For j = i + 1 To codeCount - 1
            If StrComp(codes(i), codes(j), vbTextCompare) > 0 Then
                tempValue = codes(i)
                codes(i) = codes(j)
                codes(j) = tempValue
            End If
        Next j
    Next i

    For i = 0 To codeCount - 1
        If Len(result) > 0 Then result = result & ", "
        result = result & codes(i)
    Next i

    CanonicalizeCodeList = result
End Function

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
