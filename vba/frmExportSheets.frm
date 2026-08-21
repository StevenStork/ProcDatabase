VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExportSheets 
   Caption         =   "Update Export Sheets"
   ClientHeight    =   4215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExportSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Layout is built at runtime so this .frm does not need a .frx designer blob.

Private lblScope As MSForms.Label
Private WithEvents cboScope As MSForms.ComboBox
Private lblItem As MSForms.Label
Private WithEvents cboItem As MSForms.ComboBox
Private lblAllConfirm As MSForms.Label
Private chkAllConfirm As MSForms.CheckBox
Private WithEvents cmdRun As MSForms.CommandButton
Private WithEvents cmdCancel As MSForms.CommandButton

Private Sub UserForm_Initialize()
    Me.Caption = "Update Export Sheets"
    Me.Width = 330
    Me.Height = 250

    Set lblScope = AddLabel("lblScope", 12, 12, 294, 16, "Export type")
    Set cboScope = AddCombo("cboScope", 12, 28, 294, 22)
    Set lblItem = AddLabel("lblItem", 12, 58, 294, 16, "Select item")
    Set cboItem = AddCombo("cboItem", 12, 74, 294, 22)
    Set lblAllConfirm = AddLabel("lblAllConfirm", 12, 58, 294, 48, _
        "This will rebuild the FFA export sheet and every product-line export sheet from the current Part data.")
    lblAllConfirm.WordWrap = True
    Set chkAllConfirm = AddCheckBox("chkAllConfirm", 12, 110, 294, 24, _
        "I confirm I want to update all export sheets")
    Set cmdRun = AddButton("cmdRun", 132, 170, 80, 28, "Run")
    Set cmdCancel = AddButton("cmdCancel", 226, 170, 80, 28, "Cancel")

    With cboScope
        .Clear
        .AddItem EXPORT_SCOPE_FFA
        .AddItem EXPORT_SCOPE_PRODUCT_LINE
        .AddItem EXPORT_SCOPE_ALL
        .ListIndex = 0
    End With

    ApplyScopeUi
End Sub

Private Function AddLabel( _
    ByVal controlName As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal widthPos As Single, _
    ByVal heightPos As Single, _
    ByVal captionText As String) As MSForms.Label

    Dim ctl As MSForms.Label
    Set ctl = Me.Controls.Add("Forms.Label.1", controlName, True)
    ctl.Left = leftPos
    ctl.Top = topPos
    ctl.Width = widthPos
    ctl.Height = heightPos
    ctl.Caption = captionText
    Set AddLabel = ctl
End Function

Private Function AddCombo( _
    ByVal controlName As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal widthPos As Single, _
    ByVal heightPos As Single) As MSForms.ComboBox

    Dim ctl As MSForms.ComboBox
    Set ctl = Me.Controls.Add("Forms.ComboBox.1", controlName, True)
    ctl.Left = leftPos
    ctl.Top = topPos
    ctl.Width = widthPos
    ctl.Height = heightPos
    ctl.Style = fmStyleDropDownList
    Set AddCombo = ctl
End Function

Private Function AddCheckBox( _
    ByVal controlName As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal widthPos As Single, _
    ByVal heightPos As Single, _
    ByVal captionText As String) As MSForms.CheckBox

    Dim ctl As MSForms.CheckBox
    Set ctl = Me.Controls.Add("Forms.CheckBox.1", controlName, True)
    ctl.Left = leftPos
    ctl.Top = topPos
    ctl.Width = widthPos
    ctl.Height = heightPos
    ctl.Caption = captionText
    Set AddCheckBox = ctl
End Function

Private Function AddButton( _
    ByVal controlName As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal widthPos As Single, _
    ByVal heightPos As Single, _
    ByVal captionText As String) As MSForms.CommandButton

    Dim ctl As MSForms.CommandButton
    Set ctl = Me.Controls.Add("Forms.CommandButton.1", controlName, True)
    ctl.Left = leftPos
    ctl.Top = topPos
    ctl.Width = widthPos
    ctl.Height = heightPos
    ctl.Caption = captionText
    Set AddButton = ctl
End Function

Private Sub cboScope_Change()
    ApplyScopeUi
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRun_Click()
    Dim scopeName As String
    Dim itemName As String
    Dim scopeDescription As String

    On Error GoTo Fail

    scopeName = Trim$(CStr(cboScope.Value))
    If Len(scopeName) = 0 Then
        MsgBox "Select an export type.", vbExclamation, Me.Caption
        Exit Sub
    End If

    If StrComp(scopeName, EXPORT_SCOPE_ALL, vbTextCompare) = 0 Then
        If Not chkAllConfirm.Value Then
            MsgBox "Check the confirmation box before updating all exports.", vbExclamation, Me.Caption
            Exit Sub
        End If
        itemName = vbNullString
        scopeDescription = "the FFA export sheet and every product-line export sheet"
    ElseIf StrComp(scopeName, EXPORT_SCOPE_FFA, vbTextCompare) = 0 Then
        itemName = vbNullString
        scopeDescription = "the FFA export sheet (every part number)"
    Else
        itemName = Trim$(CStr(cboItem.Value))
        If Len(itemName) = 0 Then
            MsgBox "Select an item from the second list.", vbExclamation, Me.Caption
            Exit Sub
        End If
        scopeDescription = "the " & scopeName & " export sheet for """ & itemName & """"
    End If

    If Not ConfirmExportOverwrite(scopeDescription) Then Exit Sub

    Me.Hide
    BuildExportSheetsCore scopeName, itemName, True
    MsgBox "Export update finished for " & scopeDescription & ".", vbInformation, Me.Caption
    Unload Me
    Exit Sub

Fail:
    MsgBox "Export update failed:" & vbCrLf & vbCrLf & Err.Description, vbCritical, Me.Caption
    On Error Resume Next
    Me.Show
End Sub

Private Sub ApplyScopeUi()
    Dim scopeName As String
    Dim values() As String
    Dim i As Long

    If cboScope Is Nothing Then Exit Sub
    scopeName = Trim$(CStr(cboScope.Value))

    If StrComp(scopeName, EXPORT_SCOPE_ALL, vbTextCompare) = 0 Then
        lblItem.Visible = False
        cboItem.Visible = False
        cboItem.Clear

        lblAllConfirm.Caption = "This will rebuild the FFA export sheet and every product-line export sheet from the current Part data."
        lblAllConfirm.Visible = True
        chkAllConfirm.Visible = True
        chkAllConfirm.Value = False
        Exit Sub
    End If

    If StrComp(scopeName, EXPORT_SCOPE_FFA, vbTextCompare) = 0 Then
        lblItem.Visible = False
        cboItem.Visible = False
        cboItem.Clear

        lblAllConfirm.Caption = "Creates one FFA export sheet with every part number, including Home FFA and Made In FFA."
        lblAllConfirm.Visible = True
        chkAllConfirm.Visible = False
        chkAllConfirm.Value = False
        Exit Sub
    End If

    lblAllConfirm.Visible = False
    chkAllConfirm.Visible = False
    chkAllConfirm.Value = False

    lblItem.Visible = True
    cboItem.Visible = True
    cboItem.Clear

    If StrComp(scopeName, EXPORT_SCOPE_PRODUCT_LINE, vbTextCompare) = 0 Then
        lblItem.Caption = "Select product line"
        values = ListExportProductLines()
    Else
        Exit Sub
    End If

    For i = 0 To ExportFormArrayCount(values) - 1
        cboItem.AddItem values(LBound(values) + i)
    Next i

    If cboItem.ListCount > 0 Then cboItem.ListIndex = 0
End Sub

Private Function ExportFormArrayCount(ByRef values() As String) As Long
    On Error Resume Next
    ExportFormArrayCount = UBound(values) - LBound(values) + 1
    If Err.Number <> 0 Then
        Err.Clear
        ExportFormArrayCount = 0
    End If
    On Error GoTo 0
End Function
