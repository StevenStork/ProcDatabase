Attribute VB_Name = "modActiveFilter"
Option Compare Database
Option Explicit

' Replaces Excel modRouteCardConnection assembly-list patching.
' Active Home parts x checked dashes -> tblActiveAssemblyFilter, then
' qryRouteCardActive / qryAssyStndActive / qryOperCompsActive narrow reads.

Public Sub RebuildActiveAssemblyFilter()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim assemblyNo As String
    Dim sql As String

    EnsureActiveAssemblyFilterTable
    Set db = CurrentDb
    db.Execute "DELETE FROM [" & TBL_ACTIVE_FILTER & "]", dbFailOnError

    sql = "SELECT p.BasePart, d.Dash FROM [" & TBL_PART & "] AS p " & _
        "INNER JOIN [" & TBL_PART_DASH & "] AS d ON p.BasePart = d.BasePart " & _
        "WHERE p.Active <> 0 AND d.Active <> 0"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    Do Until rs.EOF
        assemblyNo = CoerceText(rs!BasePart) & "-" & CoerceText(rs!Dash)
        db.Execute "INSERT INTO [" & TBL_ACTIVE_FILTER & "] ([" & COL_ASSEMBLY_NO_FILTER & "]) VALUES (" & _
            SqlText(assemblyNo) & ")", dbFailOnError
        rs.MoveNext
    Loop
    rs.Close

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
