Attribute VB_Name = "modLinkedData"
Option Compare Database
Option Explicit

' Refreshes linked source tables already in the Access file:
'   tblRouteCard, tblAssyStnd, tblOperComps, tblRCCP (+ optional tblProcTmYld)
' Then rebuilds the local part catalog and applies RCCP active selections.

Public Sub RefreshSourceData()
    On Error GoTo Fail
    DoCmd.Hourglass True

    EnsureLinkedSourceTables
    RefreshLinkedTable TBL_ROUTE_CARD
    RefreshLinkedTable TBL_ASSY_STANDARD
    RefreshLinkedTable TBL_OPER_COMPLETIONS
    RefreshLinkedTable TBL_RCCP
    If TableExists(TBL_PROC_TM_YLD) Then
        If IsLinkedTable(TBL_PROC_TM_YLD) Then
            RefreshLinkedTable TBL_PROC_TM_YLD
        End If
    End If

    ' Catalog DELETE/INSERT invalidates open bound forms — close first.
    CloseProcDataForms
    RebuildCatalogFromStandards
    SeedReferencesFromSources
    SetMeta META_LAST_REFRESH, Format$(Now, "yyyy-mm-dd hh:nn:ss")
    DoCmd.Hourglass False
    Exit Sub

Fail:
    DoCmd.Hourglass False
    Err.Raise Err.Number, "RefreshSourceData", Err.Description
End Sub

Public Function LinkedSourceTablesReady() As Boolean
    LinkedSourceTablesReady = TableExists(TBL_ROUTE_CARD) _
        And TableExists(TBL_ASSY_STANDARD) _
        And TableExists(TBL_OPER_COMPLETIONS) _
        And TableExists(TBL_RCCP)
End Function

Public Sub EnsureLinkedSourceTables()
    Dim missing As String

    If Not TableExists(TBL_ROUTE_CARD) Then missing = missing & vbCrLf & "  - " & TBL_ROUTE_CARD
    If Not TableExists(TBL_ASSY_STANDARD) Then missing = missing & vbCrLf & "  - " & TBL_ASSY_STANDARD
    If Not TableExists(TBL_OPER_COMPLETIONS) Then missing = missing & vbCrLf & "  - " & TBL_OPER_COMPLETIONS
    If Not TableExists(TBL_RCCP) Then missing = missing & vbCrLf & "  - " & TBL_RCCP

    If Len(missing) > 0 Then
        Err.Raise vbObjectError + 620, "EnsureLinkedSourceTables", _
            "Required linked tables are missing:" & missing & vbCrLf & vbCrLf & _
            "Link Route_Card as tblRouteCard, Assembly_Standard as tblAssyStnd, " & _
            "Oper_Completions as tblOperComps, and RCCP as tblRCCP." & vbCrLf & vbCrLf & _
            "Run DiagnoseSchema for a full table report."
    End If
End Sub

Public Function IsLinkedTable(ByVal tableName As String) As Boolean
    Dim db As DAO.Database
    Dim td As DAO.TableDef
    On Error GoTo Fail
    ' Must keep a Database variable — CurrentDb alone returns a temp object
    ' and TableDef.Connect then raises 3420 (Object invalid or no longer set).
    Set db = CurrentDb
    Set td = db.TableDefs(tableName)
    IsLinkedTable = (Len(td.Connect) > 0)
    Exit Function
Fail:
    IsLinkedTable = False
End Function

Public Sub RefreshLinkedTable(ByVal tableName As String)
    Dim td As DAO.TableDef
    Dim db As DAO.Database

    Set db = CurrentDb
    Set td = db.TableDefs(tableName)

    If Len(td.Connect) = 0 Then
        ' Local table — nothing to refresh.
        Exit Sub
    End If

    On Error GoTo Fail
    td.RefreshLink
    Exit Sub

Fail:
    Err.Raise vbObjectError + 621, "RefreshLinkedTable", _
        "Could not refresh linked table '" & tableName & "': " & Err.Description
End Sub
