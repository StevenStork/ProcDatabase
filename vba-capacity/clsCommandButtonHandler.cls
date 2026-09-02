VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommandButtonHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Btn As MSForms.CommandButton

Private mCallback As Object
Private mEventName As String

Public Sub Init( _
    ByVal btn As MSForms.CommandButton, _
    ByVal callback As Object, _
    ByVal eventName As String)

    Set Btn = btn
    Set mCallback = callback
    mEventName = eventName
End Sub

Private Sub Btn_Click()
    mCallback.HandleFormEvent mEventName
End Sub
