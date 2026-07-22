Attribute VB_Name = "modRouteCardConnection"
Option Explicit

Private Const ROUTE_CARD_CONNECTION_NAME As String = "RouteCard"
Private Const OPERATION_COMPLETIONS_CONNECTION_NAME As String = "OperationCompletions"
Private Const ASSEMBLY_NUMBER_PARAMETER As String = "@assembly_number"
Private Const ASSEMBLY_NUMBER_LIST_PARAMETER As String = "@assembly_number_list"

Private Const HOME_SHEET_NAME As String = "Home"
Private Const HOME_BASE_PART_COLUMN As String = "C"
Private Const HOME_ACTIVE_COLUMN As String = "D"
Private Const HOME_BASE_PART_START_ROW As Long = 3
Private Const HEADER_BASE_PART As String = "Base Part Number"

Private Const PART_SHEET_DASH_COLUMN As String = "E"
Private Const PART_SHEET_DASH_ACTIVE_COLUMN As String = "F"
Private Const PART_SHEET_LIST_START_ROW As Long = 9

' Updates the RouteCard connection command text so @assembly_number reflects
' active Home parts and the dash conditions checked on each part sheet,
' then refreshes the query.
Public Sub UpdateRouteCardConnection()
    Dim commandText As String
    Dim conn As WorkbookConnection
    Dim assemblyNumbers As String

    On Error GoTo CleanUp
    OptimizeExcel True

    Set conn = GetSummaryConnection(ROUTE_CARD_CONNECTION_NAME)
    assemblyNumbers = BuildActiveAssemblyNumberList()
    commandText = GetConnectionCommandText(conn)
    commandText = ReplaceQuotedParameterValue(commandText, ASSEMBLY_NUMBER_PARAMETER, assemblyNumbers, ROUTE_CARD_CONNECTION_NAME)
    SetConnectionCommandText conn, commandText
    RefreshConnection conn

CleanUp:
    OptimizeExcel False
End Sub

' Updates the OperationCompletions connection so @assembly_number is '%' and
' @assembly_number_list uses the same active Home parts / active dash conditions
' list as RouteCard, then refreshes the query.
Public Sub UpdateOperationCompletionsConnection()
    Dim commandText As String
    Dim conn As WorkbookConnection
    Dim assemblyNumbers As String

    On Error GoTo CleanUp
    OptimizeExcel True

    Set conn = GetSummaryConnection(OPERATION_COMPLETIONS_CONNECTION_NAME)
    assemblyNumbers = BuildActiveAssemblyNumberList()
    commandText = GetConnectionCommandText(conn)
    commandText = ReplaceQuotedParameterValue(commandText, ASSEMBLY_NUMBER_PARAMETER, "%", OPERATION_COMPLETIONS_CONNECTION_NAME)
    commandText = ReplaceQuotedParameterValue(commandText, ASSEMBLY_NUMBER_LIST_PARAMETER, assemblyNumbers, OPERATION_COMPLETIONS_CONNECTION_NAME)
    SetConnectionCommandText conn, commandText
    RefreshConnection conn

CleanUp:
    OptimizeExcel False
End Sub

Private Function GetSummaryConnection(ByVal connectionName As String) As WorkbookConnection
    Dim conn As WorkbookConnection

    Set conn = GetWorkbookConnection(connectionName)
    If conn Is Nothing Then
        Err.Raise vbObjectError + 515, "GetSummaryConnection", "Connection '" & connectionName & "' was not found."
    End If

    Set GetSummaryConnection = conn
End Function

' Builds assembly numbers from:
'   1) Home column C parts marked active in column D
'   2) Dash conditions on each part sheet (column E) that are checked in column F
Private Function BuildActiveAssemblyNumberList() As String
    Dim wsHome As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim basePart As String
    Dim assemblyNumbers() As String
    Dim assemblyCount As Long
    Dim activeDashes() As String
    Dim dashIndex As Long
    Dim assemblyNo As String

    Set wsHome = ThisWorkbook.Worksheets(HOME_SHEET_NAME)
    lastRow = LastUsedRowInColumn(wsHome, HOME_BASE_PART_COLUMN)

    assemblyCount = 0
    ReDim assemblyNumbers(0 To 0)

    If lastRow >= HOME_BASE_PART_START_ROW Then
        For rowIndex = HOME_BASE_PART_START_ROW To lastRow
            basePart = Trim$(CStr(wsHome.Cells(rowIndex, HOME_BASE_PART_COLUMN).Value))

            If Len(basePart) = 0 Then
                ' Skip blanks.
            ElseIf StrComp(basePart, HEADER_BASE_PART, vbTextCompare) = 0 Then
                ' Skip header text if present.
            ElseIf IsActiveFlag(wsHome.Cells(rowIndex, HOME_ACTIVE_COLUMN).Value) Then
                activeDashes = GetActiveDashConditionsFromPartSheet(basePart)

                If IsArrayInitialized(activeDashes) Then
                    For dashIndex = LBound(activeDashes) To UBound(activeDashes)
                        assemblyNo = basePart & "-" & activeDashes(dashIndex)
                        ReDim Preserve assemblyNumbers(0 To assemblyCount)
                        assemblyNumbers(assemblyCount) = assemblyNo
                        assemblyCount = assemblyCount + 1
                    Next dashIndex
                End If
            End If
        Next rowIndex
    End If

    If assemblyCount = 0 Then
        BuildActiveAssemblyNumberList = vbNullString
        Exit Function
    End If

    ReDim Preserve assemblyNumbers(0 To assemblyCount - 1)
    SortStringArray assemblyNumbers
    BuildActiveAssemblyNumberList = Join(assemblyNumbers, ", ")
End Function

Private Function GetActiveDashConditionsFromPartSheet(ByVal basePart As String) As String()
    Dim wsPart As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim dashCondition As String
    Dim activeDashes() As String
    Dim dashCount As Long

    If Not SheetExists(basePart) Then
        GetActiveDashConditionsFromPartSheet = EmptyStringArray()
        Exit Function
    End If

    Set wsPart = ThisWorkbook.Worksheets(basePart)
    lastRow = LastUsedRowInColumn(wsPart, PART_SHEET_DASH_COLUMN)
    dashCount = 0
    ReDim activeDashes(0 To 0)

    If lastRow < PART_SHEET_LIST_START_ROW Then
        GetActiveDashConditionsFromPartSheet = EmptyStringArray()
        Exit Function
    End If

    For rowIndex = PART_SHEET_LIST_START_ROW To lastRow
        dashCondition = Trim$(CStr(wsPart.Cells(rowIndex, PART_SHEET_DASH_COLUMN).Value))

        If Len(dashCondition) = 0 Then
            Exit For
        End If

        If IsActiveFlag(wsPart.Cells(rowIndex, PART_SHEET_DASH_ACTIVE_COLUMN).Value) Then
            ReDim Preserve activeDashes(0 To dashCount)
            activeDashes(dashCount) = dashCondition
            dashCount = dashCount + 1
        End If
    Next rowIndex

    If dashCount = 0 Then
        GetActiveDashConditionsFromPartSheet = EmptyStringArray()
    Else
        GetActiveDashConditionsFromPartSheet = activeDashes
    End If
End Function

Private Function IsActiveFlag(ByVal activeValue As Variant) As Boolean
    If IsError(activeValue) Then Exit Function
    If IsEmpty(activeValue) Or IsNull(activeValue) Then Exit Function

    If VarType(activeValue) = vbBoolean Then
        IsActiveFlag = CBool(activeValue)
        Exit Function
    End If

    If IsNumeric(activeValue) Then
        IsActiveFlag = (CDbl(activeValue) <> 0)
        Exit Function
    End If

    Select Case LCase$(Trim$(CStr(activeValue)))
        Case "true", "yes", "y", "1"
            IsActiveFlag = True
    End Select
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
        Err.Raise vbObjectError + 516, "ReplaceQuotedParameterValue", "Parameter '" & parameterName & "' was not found in the " & connectionName & " command text."
    End If

    equalsPos = InStr(parameterPos, commandText, "=")
    If equalsPos = 0 Then
        Err.Raise vbObjectError + 517, "ReplaceQuotedParameterValue", "Could not locate the assignment for '" & parameterName & "' in " & connectionName & "."
    End If

    openingQuotePos = InStr(equalsPos, commandText, "'")
    If openingQuotePos = 0 Then
        Err.Raise vbObjectError + 518, "ReplaceQuotedParameterValue", "Could not locate the opening quote for '" & parameterName & "' in " & connectionName & "."
    End If

    closingQuotePos = InStr(openingQuotePos + 1, commandText, "'")
    If closingQuotePos = 0 Then
        Err.Raise vbObjectError + 519, "ReplaceQuotedParameterValue", "Could not locate the closing quote for '" & parameterName & "' in " & connectionName & "."
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
            Err.Raise vbObjectError + 520, "GetConnectionCommandText", "Connection '" & conn.Name & "' uses an unsupported connection type."
    End Select
End Function

Private Sub SetConnectionCommandText(ByVal conn As WorkbookConnection, ByVal commandText As String)
    Select Case conn.Type
        Case xlConnectionTypeOLEDB
            conn.OLEDBConnection.CommandText = commandText
        Case xlConnectionTypeODBC
            conn.ODBCConnection.CommandText = commandText
        Case Else
            Err.Raise vbObjectError + 521, "SetConnectionCommandText", "Connection '" & conn.Name & "' uses an unsupported connection type."
    End Select
End Sub

Private Sub RefreshConnection(ByVal conn As WorkbookConnection)
    Select Case conn.Type
        Case xlConnectionTypeOLEDB
            conn.OLEDBConnection.BackgroundQuery = False
        Case xlConnectionTypeODBC
            conn.ODBCConnection.BackgroundQuery = False
    End Select

    conn.Refresh
End Sub

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    SheetExists = Not ws Is Nothing
End Function

Private Function LastUsedRowInColumn(ByVal ws As Worksheet, ByVal columnLetter As String) As Long
    Dim foundCell As Range
    Dim endUpRow As Long

    endUpRow = ws.Cells(ws.Rows.Count, columnLetter).End(xlUp).Row

    Set foundCell = ws.Columns(columnLetter).Find( _
        What:="*", _
        LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious)

    If foundCell Is Nothing Then
        LastUsedRowInColumn = endUpRow
    ElseIf foundCell.Row > endUpRow Then
        LastUsedRowInColumn = foundCell.Row
    Else
        LastUsedRowInColumn = endUpRow
    End If
End Function

Private Function EmptyStringArray() As String()
    Dim emptyKeys() As String
    EmptyStringArray = emptyKeys
End Function

Private Function IsArrayInitialized(ByRef values() As String) As Boolean
    Dim upperBound As Long

    On Error Resume Next
    upperBound = UBound(values)
    IsArrayInitialized = (Err.Number = 0)
    On Error GoTo 0
End Function

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
