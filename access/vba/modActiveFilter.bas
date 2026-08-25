Attribute VB_Name = "modActiveFilter"
Option Compare Database
Option Explicit

' Active assemblies come from linked tblRCCP (anything in that table is active).
' Falls back to Home Active x dash Active when tblRCCP is missing.
' tblActiveAssemblyFilter then drives qryRouteCardActive / qryAssyStndActive /
' qryOperCompsActive.

Public Sub RebuildActiveAssemblyFilter()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim assemblyNo As String
    Dim seen As Object
    Dim sql As String

    EnsureActiveAssemblyFilterTable
    Set db = CurrentDb
    db.Execute "DELETE FROM [" & TBL_ACTIVE_FILTER & "]", dbFailOnError

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    If TableExists(TBL_RCCP) Then
        sql = "SELECT [" & COL_ASSEMBLY_NO & "] FROM [" & TBL_RCCP & "]"
        Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
        Do Until rs.EOF
            assemblyNo = CoerceText(rs.Fields(COL_ASSEMBLY_NO).Value)
            If Len(assemblyNo) > 0 Then
                If Not seen.Exists(assemblyNo) Then
                    seen.Add assemblyNo, True
                    db.Execute "INSERT INTO [" & TBL_ACTIVE_FILTER & "] ([" & COL_ASSEMBLY_NO_FILTER & "]) VALUES (" & _
                        SqlText(assemblyNo) & ")", dbFailOnError
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
    Else
        sql = "SELECT p.BasePart, d.Dash FROM [" & TBL_PART & "] AS p " & _
            "INNER JOIN [" & TBL_PART_DASH & "] AS d ON p.BasePart = d.BasePart " & _
            "WHERE p.Active <> 0 AND d.Active <> 0"
        Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
        Do Until rs.EOF
            assemblyNo = CoerceText(rs!BasePart) & "-" & CoerceText(rs!Dash)
            If Len(assemblyNo) > 1 And Not seen.Exists(assemblyNo) Then
                seen.Add assemblyNo, True
                db.Execute "INSERT INTO [" & TBL_ACTIVE_FILTER & "] ([" & COL_ASSEMBLY_NO_FILTER & "]) VALUES (" & _
                    SqlText(assemblyNo) & ")", dbFailOnError
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If

    SetMeta "ActiveAssemblyList", ActiveAssemblyNumberList()
End Sub

Public Sub EnsureActiveAssemblyFilterTable()
    If TableExists(TBL_ACTIVE_FILTER) Then Exit Sub
    CurrentDb.Execute "CREATE TABLE [" & TBL_ACTIVE_FILTER & "] (" & _
        "[" & COL_ASSEMBLY_NO_FILTER & "] TEXT(50) CONSTRAINT PK_tblActiveAssemblyFilter PRIMARY KEY" & _
        ")", dbFailOnError
End Sub

Public Function ActiveFilterHasRows() As Boolean
    If Not TableExists(TBL_ACTIVE_FILTER) Then
        ActiveFilterHasRows = False
        Exit Function
    End If
    ActiveFilterHasRows = (DCount("*", TBL_ACTIVE_FILTER) > 0)
End Function

Public Function RouteCardSourceName() As String
    If ActiveFilterHasRows() And QueryExists(QRY_ROUTE_CARD_ACTIVE) Then
        RouteCardSourceName = QRY_ROUTE_CARD_ACTIVE
    Else
        RouteCardSourceName = TBL_ROUTE_CARD
    End If
End Function

Public Function AssyStndSourceName() As String
    If ActiveFilterHasRows() And QueryExists(QRY_ASSY_STND_ACTIVE) Then
        AssyStndSourceName = QRY_ASSY_STND_ACTIVE
    Else
        AssyStndSourceName = TBL_ASSY_STANDARD
    End If
End Function

Public Function OperCompsSourceName() As String
    If ActiveFilterHasRows() And QueryExists(QRY_OPER_COMPS_ACTIVE) Then
        OperCompsSourceName = QRY_OPER_COMPS_ACTIVE
    Else
        OperCompsSourceName = TBL_OPER_COMPLETIONS
    End If
End Function

Public Function ActiveAssemblyNumberList() As String
    Dim rs As DAO.Recordset
    Dim items As Collection
    Dim i As Long
    Dim names() As String

    Set items = New Collection
    If Not TableExists(TBL_ACTIVE_FILTER) Then
        ActiveAssemblyNumberList = vbNullString
        Exit Function
    End If

    Set rs = CurrentDb.OpenRecordset( _
        "SELECT [" & COL_ASSEMBLY_NO_FILTER & "] FROM [" & TBL_ACTIVE_FILTER & "] " & _
        "ORDER BY [" & COL_ASSEMBLY_NO_FILTER & "]", dbOpenSnapshot)
    Do Until rs.EOF
        items.Add CoerceText(rs.Fields(COL_ASSEMBLY_NO_FILTER).Value)
        rs.MoveNext
    Loop
    rs.Close

    If items.Count = 0 Then
        ActiveAssemblyNumberList = vbNullString
        Exit Function
    End If

    ReDim names(1 To items.Count)
    For i = 1 To items.Count
        names(i) = items(i)
    Next i
    ActiveAssemblyNumberList = Join(names, ", ")
End Function
