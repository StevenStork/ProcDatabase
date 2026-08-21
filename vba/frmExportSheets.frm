VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExportSheets 
   Caption         =   "Update Export Sheets"
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "frmExportSheets.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExportSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Manual control setup (if importing the .frm without .frx):
'   lblScope     Label        Caption "Export type"
'   cboScope     ComboBox     Style 2-fmStyleDropDownList
'   lblItem      Label        Caption "Select item"
'   cboItem      ComboBox     Style 2-fmStyleDropDownList
'   lblAllConfirm Label       Caption (see code); WordWrap True; Visible False
'   chkAllConfirm CheckBox    Caption "I confirm I want to update all export sheets"
'                             Visible False
'   cmdRun       CommandButton Caption "Run"
'   cmdCancel    CommandButton Caption "Cancel"
'
' Suggested layout (twips approx): form 5760x3720
'   lblScope 120,120, 2000x240
'   cboScope 120,360, 5400x240
'   lblItem 120,720, 5400x240
'   cboItem 120,960, 5400x240
'   lblAllConfirm 120,720, 5400x600
'   chkAllConfirm 120,1400, 5400x360
'   cmdRun 2400,3000, 1400x400
'   cmdCancel 4000,3000, 1400x400

Private Sub UserForm_Initialize()
    With cboScope
        .Clear
        .AddItem EXPORT_SCOPE_FFA
        .AddItem EXPORT_SCOPE_PRODUCT_LINE
        .AddItem EXPORT_SCOPE_ALL
        .ListIndex = 0
    End With

    lblAllConfirm.Caption = _
        "This will rebuild every FFA and product-line export sheet from the current Part data."
    lblAllConfirm.WordWrap = True

    ApplyScopeUi
End Sub

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
        scopeDescription = "all FFA and product-line export sheets"
    Else
        itemName = Trim$(CStr(cboItem.Value))
        If Len(itemName) = 0 Then
            MsgBox "Select an item from the second list.", vbExclamation, Me.Caption
            Exit Sub
        End If
        scopeDescription = "the " & scopeName & " export sheet for """ & itemName & """"
    End If

    ' Always confirm before overwriting — old export data cannot be recovered.
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

    scopeName = Trim$(CStr(cboScope.Value))

    If StrComp(scopeName, EXPORT_SCOPE_ALL, vbTextCompare) = 0 Then
        lblItem.Visible = False
        cboItem.Visible = False
        cboItem.Clear

        lblAllConfirm.Visible = True
        chkAllConfirm.Visible = True
        chkAllConfirm.Value = False
        Exit Sub
    End If

    lblAllConfirm.Visible = False
    chkAllConfirm.Visible = False
    chkAllConfirm.Value = False

    lblItem.Visible = True
    cboItem.Visible = True
    cboItem.Clear

    If StrComp(scopeName, EXPORT_SCOPE_FFA, vbTextCompare) = 0 Then
        lblItem.Caption = "Select FFA"
        values = ListExportFfas()
    ElseIf StrComp(scopeName, EXPORT_SCOPE_PRODUCT_LINE, vbTextCompare) = 0 Then
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
