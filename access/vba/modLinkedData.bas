Attribute VB_Name = "modLinkedData"
Option Compare Database
Option Explicit

' Refreshes the three linked source tables already in the Access file:
'   tblRouteCard  <- Route_Card
'   tblAssyStnd   <- Assembly_Standard / Assy_Standard
'   tblOperComps  <- Oper_Completions
' Then rebuilds the local part catalog from standards.

Public Sub RefreshSourceData()
    On Error GoTo Fail
    DoCmd.Hourglass True

    EnsureLinkedSourceTables
    RefreshLinkedTable TBL_ROUTE_CARD
    RefreshLinkedTable TBL_ASSY_STANDARD
    RefreshLinkedTable TBL_OPER_COMPLETIONS
    If TableExists(TBL_PROC_TM_YLD) Then
        If IsLinkedTable(TBL_PROC_TM_YLD) Then
            RefreshLinkedTable TBL_PROC_TM_YLD
        End If
    End If

    RebuildCatalogFromStandards
    RequeryOpenForms
    SetMeta "LastRefresh", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    DoCmd.Hourglass False
    Exit Sub

Fail:
    DoCmd.Hourglass False
    Err.Raise Err.Number, "RefreshSourceData", Err.Description
End Sub

Public Sub EnsureLinkedSourceTables()
    Dim missing As String

    If Not TableExists(TBL_ROUTE_CARD) Then missing = missing & vbCrLf & "  - " & TBL_ROUTE_CARD
    If Not TableExists(TBL_ASSY_STANDARD) Then missing = missing & vbCrLf & "  - " & TBL_ASSY_STANDARD
    If Not TableExists(TBL_OPER_COMPLETIONS) Then missing = missing & vbCrLf & "  - " & TBL_OPER_COMPLETIONS

    If Len(missing) > 0 Then
        Err.Raise vbObjectError + 620, "EnsureLinkedSourceTables", _
            "Required linked tables are missing:" & missing & vbCrLf & vbCrLf & _
            "Link Route_Card as tblRouteCard, Assembly_Standard as tblAssyStnd, " & _
            "and Oper_Completions as tblOperComps."
    End If
End Sub

Public Function IsLinkedTable(ByVal tableName As String) As Boolean
    Dim td As DAO.TableDef
    On Error Resume Next
    Set td = CurrentDb.TableDefs(tableName)
    If Err.Number <> 0 Then
        IsLinkedTable = False
        Exit Function
    End If
    On Error GoTo 0
    IsLinkedTable = (Len(td.Connect) > 0)
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

Private Sub RequeryOpenForms()
    Dim frm As Form
    For Each frm In Forms
        On Error Resume Next
        frm.Requery
        On Error GoTo 0
    Next frm
End Sub
