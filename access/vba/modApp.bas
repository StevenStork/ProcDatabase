Attribute VB_Name = "modApp"
Option Compare Database
Option Explicit

' Entry points callable from AutoExec, ribbon, or immediate window.

Public Sub StartProcDatabase()
    On Error GoTo Fail
    EnsureSchema
    EnsureQueries
    EnsureStartup
    DoCmd.OpenForm FRM_HOME
    Exit Sub
Fail:
    MsgBox "Start failed: " & Err.Description, vbCritical, "ProcDatabase"
End Sub

Public Sub RefreshAll()
    On Error GoTo Fail
    RefreshSourceData
    MsgBox "Linked tables refreshed and part catalog rebuilt.", vbInformation, "ProcDatabase"
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
    MsgBox "UI build failed: " & Err.Description, vbCritical, "ProcDatabase"
End Sub
