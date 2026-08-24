Attribute VB_Name = "modApp"
Option Compare Database
Option Explicit

Private BootstrapStep As String

' Entry points callable from startup, ribbon, or Immediate window.

Public Sub BootstrapProcDatabase()
    On Error GoTo Fail
    DoCmd.Hourglass True
    BootstrapStep = "EnsureSchema"
    EnsureSchema
    BootstrapStep = "EnsureQueries"
    EnsureQueries
    BootstrapStep = "EnsureUi"
    EnsureUi
    BootstrapStep = "EnsureStartup"
    EnsureStartup
    BootstrapStep = "SetMeta"
    SetMeta "SchemaVersion", SCHEMA_VERSION
    DoCmd.Hourglass False
    MsgBox "ProcDatabase is ready." & vbCrLf & vbCrLf & _
        "Linked: " & TBL_ROUTE_CARD & ", " & TBL_ASSY_STANDARD & ", " & TBL_OPER_COMPLETIONS & vbCrLf & _
        "Run RefreshAll to pull data and rebuild the catalog.", vbInformation, "ProcDatabase"
    Exit Sub
Fail:
    DoCmd.Hourglass False
    MsgBox "Bootstrap failed during " & BootstrapStep & ":" & vbCrLf & vbCrLf & _
        Err.Description & " (" & Err.Number & ")", vbCritical, "ProcDatabase"
End Sub

' Schema + queries only (skip form build). Use if EnsureUi fails.
Public Sub BootstrapSchemaOnly()
    On Error GoTo Fail
    BootstrapStep = "EnsureSchema"
    EnsureSchema
    BootstrapStep = "EnsureQueries"
    EnsureQueries
    SetMeta "SchemaVersion", SCHEMA_VERSION
    MsgBox "Schema and queries are ready. Run BuildUi to create forms.", vbInformation, "ProcDatabase"
    Exit Sub
Fail:
    MsgBox "BootstrapSchemaOnly failed during " & BootstrapStep & ":" & vbCrLf & vbCrLf & _
        Err.Description & " (" & Err.Number & ")", vbCritical, "ProcDatabase"
End Sub

Public Sub StartProcDatabase()
    On Error GoTo Fail
    EnsureSchema
    EnsureQueries
    EnsureStartup
    If Not ObjectExists(acForm, FRM_HOME) Then
        EnsureUi
    End If
    DoCmd.OpenForm FRM_HOME
    Exit Sub
Fail:
    MsgBox "Start failed: " & Err.Description, vbCritical, "ProcDatabase"
End Sub

Public Sub RefreshAll()
    On Error GoTo Fail
    RefreshSourceData
    RebuildActiveAssemblyFilter
    MsgBox "Linked tables refreshed, catalog rebuilt, and active assembly filter updated.", vbInformation, "ProcDatabase"
    Exit Sub
Fail:
    MsgBox "Refresh failed: " & Err.Description, vbCritical, "ProcDatabase"
End Sub

Public Sub BuildUi()
    On Error GoTo Fail
    EnsureSchema
    EnsureQueries
    EnsureUi
    EnsureStartup
    MsgBox "Schema, queries, and forms are ready. Startup form is " & FRM_HOME & ".", vbInformation, "ProcDatabase"
    Exit Sub
Fail:
    MsgBox "UI build failed: " & Err.Description & " (" & Err.Number & ")", vbCritical, "ProcDatabase"
End Sub

Public Sub RebuildFilterOnly()
    RebuildActiveAssemblyFilter
    MsgBox "Active assembly filter rebuilt: " & ActiveAssemblyNumberList(), vbInformation, "ProcDatabase"
End Sub
