Attribute VB_Name = "modActiveFilter"
Option Compare Database
Option Explicit

' Active assemblies come from linked tblRCCP (anything in that table is active).
' Falls back to Home Active x dash Active when tblRCCP is missing.
' tblActiveAssemblyFilter then drives qryRouteCardActive / qryAssyStndActive /
' qryOperCompsActive.

Public Sub RebuildActiveAssemblyFilter()
    Dim db As DAO.Database
    Dim sql As String

    EnsureActiveAssemblyFilterTable
    Set db = CurrentDb
    db.Execute "DELETE FROM [" & TBL_ACTIVE_FILTER & "]", dbFailOnError

    If TableExists(TBL_RCCP) Then
        ' Set-based distinct load — avoids per-row INSERT Execute.
        sql = "INSERT INTO [" & TBL_ACTIVE_FILTER & "] ([" & COL_ASSEMBLY_NO_FILTER & "]) " & _
            "SELECT DISTINCT [" & COL_ASSEMBLY_NO & "] FROM [" & TBL_RCCP & "] " & _
            "WHERE Len(Nz([" & COL_ASSEMBLY_NO & "],'')) > 0"
        db.Execute sql, dbFailOnError
    Else
        sql = "INSERT INTO [" & TBL_ACTIVE_FILTER & "] ([" & COL_ASSEMBLY_NO_FILTER & "]) " & _
            "SELECT DISTINCT p.BasePart & '-' & d.Dash FROM [" & TBL_PART & "] AS p " & _
            "INNER JOIN [" & TBL_PART_DASH & "] AS d ON p.BasePart = d.BasePart " & _
            "WHERE p.Active <> 0 AND d.Active <> 0"
        db.Execute sql, dbFailOnError
    End If

    SetMeta META_ACTIVE_ASSEMBLY_LIST, ActiveAssemblyNumberList()
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
    Dim parts As String

    If Not TableExists(TBL_ACTIVE_FILTER) Then
        ActiveAssemblyNumberList = vbNullString
        Exit Function
    End If

    parts = vbNullString
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT [" & COL_ASSEMBLY_NO_FILTER & "] FROM [" & TBL_ACTIVE_FILTER & "] " & _
        "ORDER BY [" & COL_ASSEMBLY_NO_FILTER & "]", dbOpenSnapshot)
    Do Until rs.EOF
        If Len(parts) > 0 Then parts = parts & ", "
        parts = parts & CoerceText(rs.Fields(COL_ASSEMBLY_NO_FILTER).Value)
        rs.MoveNext
    Loop
    rs.Close
    ActiveAssemblyNumberList = parts
End Function
