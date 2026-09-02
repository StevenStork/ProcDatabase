VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsListBoxHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Lst As MSForms.ListBox

Private mCallback As Object
Private mEventName As String

Public Sub Init( _
    ByVal lst As MSForms.ListBox, _
    ByVal callback As Object, _
    ByVal eventName As String)

    Set Lst = lst
    Set mCallback = callback
    mEventName = eventName
End Sub

Private Sub Lst_Click()
    mCallback.HandleFormEvent mEventName
End Sub
