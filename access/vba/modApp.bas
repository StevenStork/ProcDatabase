Attribute VB_Name = "modApp"
Option Compare Database
Option Explicit

Private BootstrapStep As String

' Entry points callable from startup, ribbon, or Immediate window.

Public Sub BootstrapProcDatabase()
    On Error GoTo Fail
    DoCmd.Hourglass True
    CloseProcDataForms
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
    On Error Resume Next
    OpenFormSized FRM_HOME
    On Error GoTo 0
    MsgBox "ProcDatabase is ready." & vbCrLf & vbCrLf & _
        "Linked: " & TBL_ROUTE_CARD & ", " & TBL_ASSY_STANDARD & ", " & TBL_OPER_COMPLETIONS & ", " & TBL_RCCP & vbCrLf & _
        "Run RefreshAll to pull data and rebuild the catalog.", vbInformation, "ProcDatabase"
    Exit Sub
Fail:
    DoCmd.Hourglass False
    Dim detail As String
    detail = Err.Description & " (" & Err.Number & ")"
    If BootstrapStep = "EnsureSchema" And Len(SchemaSubStep) > 0 Then
        detail = detail & vbCrLf & vbCrLf & "Schema sub-step: " & SchemaSubStep
    End If
    If BootstrapStep = "EnsureUi" And Len(UiSubStep) > 0 Then
        detail = detail & vbCrLf & vbCrLf & "UI sub-step: " & UiSubStep
    End If
    MsgBox "Bootstrap failed during " & BootstrapStep & ":" & vbCrLf & vbCrLf & detail, vbCritical, "ProcDatabase"
End Sub

' Schema + queries only (skip form build). Use if EnsureUi fails.
Public Sub BootstrapSchemaOnly()
    On Error GoTo Fail
    CloseProcDataForms
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
        CloseProcDataForms
        EnsureUi
    End If
    OpenFormSized FRM_HOME
    Exit Sub
Fail:
    MsgBox "Start failed: " & Err.Description, vbCritical, "ProcDatabase"
End Sub

Public Sub RefreshAll()
    On Error GoTo Fail
    RefreshSourceData
    RebuildActiveAssemblyFilter
    On Error Resume Next
    OpenFormSized FRM_HOME
    On Error GoTo Fail
    MsgBox "Linked tables refreshed, catalog rebuilt, and active assembly filter updated.", vbInformation, "ProcDatabase"
    Exit Sub
Fail:
    MsgBox "Refresh failed: " & Err.Description & " (" & Err.Number & ")", vbCritical, "ProcDatabase"
End Sub

Public Sub BuildUi()
    On Error GoTo Fail
    Dim buildStep As String

    ' Rebuild forms/queries only. Skip EnsureSchema — ALTER TABLE on tblPart
    ' needs an exclusive lock and fails (3211) while Home/startup UI has it open.
    buildStep = "CloseForms"
    On Error Resume Next
    SetDbProperty "StartupForm", dbText, vbNullString
    On Error GoTo Fail
    CloseProcDataForms
    DBEngine.Idle dbRefreshCache
    DoEvents

    buildStep = "EnsureQueries"
    EnsureQueries
    buildStep = "EnsureUi"
    EnsureUi
    buildStep = "EnsureStartup"
    EnsureStartup
    On Error Resume Next
    OpenFormSized FRM_HOME
    On Error GoTo Fail
    MsgBox "Queries and forms are ready. Startup form is " & FRM_HOME & ".", vbInformation, "ProcDatabase"
    Exit Sub
Fail:
    Dim uiDetail As String
    uiDetail = Err.Description & " (" & Err.Number & ")"
    If Len(UiSubStep) > 0 Then
        uiDetail = uiDetail & vbCrLf & vbCrLf & "UI sub-step: " & UiSubStep
    End If
    MsgBox "UI build failed during " & buildStep & ":" & vbCrLf & vbCrLf & uiDetail, vbCritical, "ProcDatabase"
End Sub

Public Sub RebuildFilterOnly()
    RebuildActiveAssemblyFilter
    MsgBox "Active assembly filter rebuilt: " & ActiveAssemblyNumberList(), vbInformation, "ProcDatabase"
End Sub
