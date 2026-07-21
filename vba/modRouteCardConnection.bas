Attribute VB_Name = "modRouteCardConnection"
Option Explicit

Private Const ROUTE_CARD_CONNECTION_NAME As String = "RouteCard"
Private Const SUMMARY_TABLE_NAME As String = "BasePartSummaryTbl"
Private Const HEADER_BASE_PART As String = "Base Part Number"
Private Const HEADER_DASH_CONDITIONS As String = "Dash Conditions"
Private Const HEADER_ACTIVE_PART As String = "Active Part"
Private Const ASSEMBLY_NUMBER_PARAMETER As String = "@assembly_number"

' Updates the RouteCard connection command text so @assembly_number reflects
' the active parts on the Home sheet summary table, then refreshes the query.
Public Sub UpdateRouteCardConnection()
    Dim wsHome As Worksheet
    Dim summaryTable As ListObject
    Dim conn As WorkbookConnection
    Dim assemblyNumbers As String
    Dim commandText As String
    Dim updatedCommandText As String

    On Error GoTo CleanUp
    OptimizeExcel True

    Set wsHome = ThisWorkbook.Worksheets("Home")
    Set summaryTable = wsHome.ListObjects(SUMMARY_TABLE_NAME)

    If summaryTable Is Nothing Then
        Err.Raise vbObjectError + 513, "UpdateRouteCardConnection", "Summary table '" & SUMMARY_TABLE_NAME & "' was not found on the Home sheet."
    End If

    If summaryTable.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 514, "UpdateRouteCardConnection", "No summary rows are available to build the RouteCard assembly list."
    End If

    Set conn = GetWorkbookConnection(ROUTE_CARD_CONNECTION_NAME)
    If conn Is Nothing Then
        Err.Raise vbObjectError + 515, "UpdateRouteCardConnection", "Connection '" & ROUTE_CARD_CONNECTION_NAME & "' was not found."
    End If

    assemblyNumbers = BuildActiveAssemblyNumberList(summaryTable)
    commandText = GetConnectionCommandText(conn)
    updatedCommandText = ReplaceAssemblyNumberParameter(commandText, assemblyNumbers)
    SetConnectionCommandText conn, updatedCommandText
    RefreshConnection conn

CleanUp:
    OptimizeExcel False
End Sub

Private Function BuildActiveAssemblyNumberList(ByVal summaryTable As ListObject) As String
    Dim rowIndex As Long
    Dim basePart As String
    Dim dashConditions As String
    Dim isActive As Boolean
    Dim assemblyNumbers() As String
    Dim assemblyCount As Long
    Dim assemblyNo As String

    assemblyCount = 0
    ReDim assemblyNumbers(0 To 0)

    For rowIndex = 1 To summaryTable.DataBodyRange.Rows.Count
        isActive = CBool(summaryTable.ListColumns(HEADER_ACTIVE_PART).DataBodyRange.Cells(rowIndex, 1).Value)

        If isActive Then
            basePart = Trim$(CStr(summaryTable.ListColumns(HEADER_BASE_PART).DataBodyRange.Cells(rowIndex, 1).Value))
            dashConditions = Trim$(CStr(summaryTable.ListColumns(HEADER_DASH_CONDITIONS).DataBodyRange.Cells(rowIndex, 1).Value))
            assemblyNo = BuildAssemblyNumber(basePart, dashConditions)

            If Len(assemblyNo) > 0 Then
                ReDim Preserve assemblyNumbers(0 To assemblyCount)
                assemblyNumbers(assemblyCount) = assemblyNo
                assemblyCount = assemblyCount + 1
            End If
        End If
    Next rowIndex

    If assemblyCount = 0 Then
        BuildActiveAssemblyNumberList = vbNullString
        Exit Function
    End If

    ReDim Preserve assemblyNumbers(0 To assemblyCount - 1)
    SortStringArray assemblyNumbers
    BuildActiveAssemblyNumberList = Join(assemblyNumbers, ", ")
End Function

Private Function BuildAssemblyNumber(ByVal basePart As String, ByVal dashConditions As String) As String
    Dim largestDashCondition As String

    If Len(basePart) = 0 Then Exit Function

    largestDashCondition = GetLargestDashCondition(dashConditions)

    If Len(largestDashCondition) > 0 Then
        BuildAssemblyNumber = basePart & "-" & largestDashCondition
    Else
        BuildAssemblyNumber = basePart
    End If
End Function

Private Function GetLargestDashCondition(ByVal dashConditions As String) As String
    Dim dashValues() As String
    Dim i As Long
    Dim currentValue As String
    Dim largestValue As String

    dashValues = SplitDelimitedList(dashConditions)
    If UBound(dashValues) < LBound(dashValues) Then Exit Function

    largestValue = Trim$(dashValues(LBound(dashValues)))
    If Len(largestValue) = 0 Then Exit Function

    For i = LBound(dashValues) + 1 To UBound(dashValues)
        currentValue = Trim$(dashValues(i))

        If Len(currentValue) > 0 Then
            If CompareDashCondition(currentValue, largestValue) > 0 Then
                largestValue = currentValue
            End If
        End If
    Next i

    GetLargestDashCondition = largestValue
End Function

Private Function CompareDashCondition(ByVal leftValue As String, ByVal rightValue As String) As Long
    If IsNumeric(leftValue) And IsNumeric(rightValue) Then
        CompareDashCondition = Sgn(CDbl(leftValue) - CDbl(rightValue))
    Else
        CompareDashCondition = StrComp(leftValue, rightValue, vbTextCompare)
    End If
End Function

Private Function SplitDelimitedList(ByVal value As String) As String()
    SplitDelimitedList = Split(Replace$(value, ", ", ","), ",")
End Function

Private Function ReplaceAssemblyNumberParameter(ByVal commandText As String, ByVal assemblyNumbers As String) As String
    Dim parameterPos As Long
    Dim equalsPos As Long
    Dim openingQuotePos As Long
    Dim closingQuotePos As Long

    parameterPos = InStr(1, commandText, ASSEMBLY_NUMBER_PARAMETER, vbTextCompare)
    If parameterPos = 0 Then
        Err.Raise vbObjectError + 516, "ReplaceAssemblyNumberParameter", "Parameter '" & ASSEMBLY_NUMBER_PARAMETER & "' was not found in the RouteCard command text."
    End If

    equalsPos = InStr(parameterPos, commandText, "=")
    If equalsPos = 0 Then
        Err.Raise vbObjectError + 517, "ReplaceAssemblyNumberParameter", "Could not locate the assignment for '" & ASSEMBLY_NUMBER_PARAMETER & "'."
    End If

    openingQuotePos = InStr(equalsPos, commandText, "'")
    If openingQuotePos = 0 Then
        Err.Raise vbObjectError + 518, "ReplaceAssemblyNumberParameter", "Could not locate the opening quote for '" & ASSEMBLY_NUMBER_PARAMETER & "'."
    End If

    closingQuotePos = InStr(openingQuotePos + 1, commandText, "'")
    If closingQuotePos = 0 Then
        Err.Raise vbObjectError + 519, "ReplaceAssemblyNumberParameter", "Could not locate the closing quote for '" & ASSEMBLY_NUMBER_PARAMETER & "'."
    End If

    ReplaceAssemblyNumberParameter = Left$(commandText, openingQuotePos) & assemblyNumbers & Mid$(commandText, closingQuotePos)
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
