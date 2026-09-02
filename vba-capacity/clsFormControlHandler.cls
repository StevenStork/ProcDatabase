VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormControlHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' One handler class with typed WithEvents for each supported control.
' Only the control passed to Init is wired; the others stay Nothing.

Public WithEvents Btn As MSForms.CommandButton
Public WithEvents Lst As MSForms.ListBox
Public WithEvents Cbo As MSForms.ComboBox

Private mCallback As Object
Private mEventName As String

Public Sub Init( _
    ByVal ctl As Object, _
    ByVal callback As Object, _
    ByVal eventName As String, _
    ByVal eventKind As String)

    Dim controlType As String

    Set mCallback = callback
    mEventName = eventName
    controlType = TypeName(ctl)

    Select Case LCase$(eventKind)
        Case "click"
            If controlType = "CommandButton" Then
                Set Btn = ctl
            ElseIf controlType = "ListBox" Then
                Set Lst = ctl
            Else
                Err.Raise vbObjectError + 514, "clsFormControlHandler.Init", _
                    "Control type '" & controlType & "' does not support Click."
            End If
        Case "change"
            If controlType = "ComboBox" Then
                Set Cbo = ctl
            Else
                Err.Raise vbObjectError + 515, "clsFormControlHandler.Init", _
                    "Control type '" & controlType & "' does not support Change."
            End If
        Case Else
            Err.Raise vbObjectError + 516, "clsFormControlHandler.Init", _
                "Unsupported event kind '" & eventKind & "'."
    End Select
End Sub

Private Sub Btn_Click()
    mCallback.HandleFormEvent mEventName
End Sub

Private Sub Lst_Click()
    mCallback.HandleFormEvent mEventName
End Sub

Private Sub Cbo_Change()
    mCallback.HandleFormEvent mEventName
End Sub
