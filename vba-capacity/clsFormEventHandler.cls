VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Target As MSForms.Control

Private mCallback As Object
Private mEventName As String
Private mEventKind As String

Public Sub Init( _
    ByVal ctl As MSForms.Control, _
    ByVal callback As Object, _
    ByVal eventName As String, _
    ByVal eventKind As String)

    Set Target = ctl
    Set mCallback = callback
    mEventName = eventName
    mEventKind = eventKind
End Sub

Private Sub Target_Click()
    If mEventKind = "Click" Then
        mCallback.HandleFormEvent mEventName
    End If
End Sub

Private Sub Target_Change()
    If mEventKind = "Change" Then
        mCallback.HandleFormEvent mEventName
    End If
End Sub

Private Sub Target_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If mEventKind = "DblClick" Then
        mCallback.HandleFormEvent mEventName
    End If
End Sub
