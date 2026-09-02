VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsComboBoxHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Cbo As MSForms.ComboBox

Private mCallback As Object
Private mEventName As String

Public Sub Init( _
    ByVal cbo As MSForms.ComboBox, _
    ByVal callback As Object, _
    ByVal eventName As String)

    Set Cbo = cbo
    Set mCallback = callback
    mEventName = eventName
End Sub

Private Sub Cbo_Change()
    mCallback.HandleFormEvent mEventName
End Sub
