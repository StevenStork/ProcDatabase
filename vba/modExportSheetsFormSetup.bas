Attribute VB_Name = "modExportSheetsFormSetup"
Option Explicit

' Creates frmExportSheets with the required controls and code when it is not
' already in the workbook. Requires:
'   File → Options → Trust Center → Trust Center Settings → Macro Settings →
'   "Trust access to the VBA project object model"
'
' After install, run ShowUpdateExportSheets.

Private Const FORM_NAME As String = "frmExportSheets"

Public Sub InstallExportSheetsForm()
    Dim vbProj As Object
    Dim vbComp As Object
    Dim designer As Object
    Dim codeMod As Object
    Dim existing As Object

    On Error GoTo Fail

    Set vbProj = ThisWorkbook.VBProject

    On Error Resume Next
    Set existing = vbProj.VBComponents(FORM_NAME)
    On Error GoTo Fail

    If Not existing Is Nothing Then
        MsgBox _
            """" & FORM_NAME & """ already exists in this workbook." & vbCrLf & vbCrLf & _
            "Run ShowUpdateExportSheets to open it.", _
            vbInformation, "Install Export Form"
        Exit Sub
    End If

    Set vbComp = vbProj.VBComponents.Add(3) ' vbext_ct_MSForm
    vbComp.Name = FORM_NAME
    vbComp.Properties("Caption") = "Update Export Sheets"
    vbComp.Properties("Width") = 320
    vbComp.Properties("Height") = 240

    Set designer = vbComp.Designer
    AddFormControl designer, "Forms.Label.1", "lblScope", 12, 12, 280, 18, "Export type"
    AddFormControl designer, "Forms.ComboBox.1", "cboScope", 12, 30, 280, 22, vbNullString
    AddFormControl designer, "Forms.Label.1", "lblItem", 12, 60, 280, 18, "Select item"
    AddFormControl designer, "Forms.ComboBox.1", "cboItem", 12, 78, 280, 22, vbNullString
    AddFormControl designer, "Forms.Label.1", "lblAllConfirm", 12, 60, 280, 48, vbNullString
    AddFormControl designer, "Forms.CheckBox.1", "chkAllConfirm", 12, 112, 280, 24, _
        "I confirm I want to update all export sheets"
    AddFormControl designer, "Forms.CommandButton.1", "cmdRun", 120, 160, 80, 28, "Run"
    AddFormControl designer, "Forms.CommandButton.1", "cmdCancel", 210, 160, 80, 28, "Cancel"

    designer.Controls("cboScope").Style = 2 ' fmStyleDropDownList
    designer.Controls("cboItem").Style = 2
    designer.Controls("lblAllConfirm").WordWrap = True
    designer.Controls("lblAllConfirm").Visible = False
    designer.Controls("chkAllConfirm").Visible = False

    Set codeMod = vbComp.CodeModule
    If codeMod.CountOfLines > 0 Then
        codeMod.DeleteLines 1, codeMod.CountOfLines
    End If
    codeMod.AddFromString ExportSheetsFormCode()

    MsgBox _
        """" & FORM_NAME & """ was created." & vbCrLf & vbCrLf & _
        "Run ShowUpdateExportSheets to open the form.", _
        vbInformation, "Install Export Form"
    Exit Sub

Fail:
    MsgBox _
        "Could not install the export form." & vbCrLf & vbCrLf & _
        "Enable ""Trust access to the VBA project object model"" in Excel Trust Center, " & _
        "or import vba/frmExportSheets.frm and add the controls listed in " & _
        "vba/frmExportSheets_Setup.txt." & vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical, "Install Export Form"
End Sub

Private Sub AddFormControl( _
    ByVal designer As Object, _
    ByVal progId As String, _
    ByVal controlName As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal widthPos As Single, _
    ByVal heightPos As Single, _
    ByVal captionText As String)

    Dim ctl As Object

    Set ctl = designer.Controls.Add(progId, controlName, True)
    ctl.Left = leftPos
    ctl.Top = topPos
    ctl.Width = widthPos
    ctl.Height = heightPos

    On Error Resume Next
    ctl.Caption = captionText
    On Error GoTo 0
End Sub

Private Function ExportSheetsFormCode() As String
    Dim s As String

    s = s & "Option Explicit" & vbCrLf
    s = s & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    With cboScope" & vbCrLf
    s = s & "        .Clear" & vbCrLf
    s = s & "        .AddItem EXPORT_SCOPE_FFA" & vbCrLf
    s = s & "        .AddItem EXPORT_SCOPE_PRODUCT_LINE" & vbCrLf
    s = s & "        .AddItem EXPORT_SCOPE_ALL" & vbCrLf
    s = s & "        .ListIndex = 0" & vbCrLf
    s = s & "    End With" & vbCrLf
    s = s & "    lblAllConfirm.Caption = ""This will rebuild every FFA and product-line export sheet from the current Part data.""" & vbCrLf
    s = s & "    lblAllConfirm.WordWrap = True" & vbCrLf
    s = s & "    ApplyScopeUi" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & vbCrLf
    s = s & "Private Sub cboScope_Change()" & vbCrLf
    s = s & "    ApplyScopeUi" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & vbCrLf
    s = s & "Private Sub cmdCancel_Click()" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & vbCrLf
    s = s & "Private Sub cmdRun_Click()" & vbCrLf
    s = s & "    Dim scopeName As String" & vbCrLf
    s = s & "    Dim itemName As String" & vbCrLf
    s = s & "    Dim scopeDescription As String" & vbCrLf
    s = s & "    On Error GoTo Fail" & vbCrLf
    s = s & "    scopeName = Trim$(CStr(cboScope.Value))" & vbCrLf
    s = s & "    If Len(scopeName) = 0 Then" & vbCrLf
    s = s & "        MsgBox ""Select an export type."", vbExclamation, Me.Caption" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If StrComp(scopeName, EXPORT_SCOPE_ALL, vbTextCompare) = 0 Then" & vbCrLf
    s = s & "        If Not chkAllConfirm.Value Then" & vbCrLf
    s = s & "            MsgBox ""Check the confirmation box before updating all exports."", vbExclamation, Me.Caption" & vbCrLf
    s = s & "            Exit Sub" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        itemName = vbNullString" & vbCrLf
    s = s & "        scopeDescription = ""all FFA and product-line export sheets""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        itemName = Trim$(CStr(cboItem.Value))" & vbCrLf
    s = s & "        If Len(itemName) = 0 Then" & vbCrLf
    s = s & "            MsgBox ""Select an item from the second list."", vbExclamation, Me.Caption" & vbCrLf
    s = s & "            Exit Sub" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        scopeDescription = ""the "" & scopeName & "" export sheet for """""" & itemName & """"""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Not ConfirmExportOverwrite(scopeDescription) Then Exit Sub" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "    BuildExportSheetsCore scopeName, itemName, True" & vbCrLf
    s = s & "    MsgBox ""Export update finished for "" & scopeDescription & ""."", vbInformation, Me.Caption" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "Fail:" & vbCrLf
    s = s & "    MsgBox ""Export update failed:"" & vbCrLf & vbCrLf & Err.Description, vbCritical, Me.Caption" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Me.Show" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & vbCrLf
    s = s & "Private Sub ApplyScopeUi()" & vbCrLf
    s = s & "    Dim scopeName As String" & vbCrLf
    s = s & "    Dim values() As String" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    scopeName = Trim$(CStr(cboScope.Value))" & vbCrLf
    s = s & "    If StrComp(scopeName, EXPORT_SCOPE_ALL, vbTextCompare) = 0 Then" & vbCrLf
    s = s & "        lblItem.Visible = False" & vbCrLf
    s = s & "        cboItem.Visible = False" & vbCrLf
    s = s & "        cboItem.Clear" & vbCrLf
    s = s & "        lblAllConfirm.Visible = True" & vbCrLf
    s = s & "        chkAllConfirm.Visible = True" & vbCrLf
    s = s & "        chkAllConfirm.Value = False" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    lblAllConfirm.Visible = False" & vbCrLf
    s = s & "    chkAllConfirm.Visible = False" & vbCrLf
    s = s & "    chkAllConfirm.Value = False" & vbCrLf
    s = s & "    lblItem.Visible = True" & vbCrLf
    s = s & "    cboItem.Visible = True" & vbCrLf
    s = s & "    cboItem.Clear" & vbCrLf
    s = s & "    If StrComp(scopeName, EXPORT_SCOPE_FFA, vbTextCompare) = 0 Then" & vbCrLf
    s = s & "        lblItem.Caption = ""Select FFA""" & vbCrLf
    s = s & "        values = ListExportFfas()" & vbCrLf
    s = s & "    ElseIf StrComp(scopeName, EXPORT_SCOPE_PRODUCT_LINE, vbTextCompare) = 0 Then" & vbCrLf
    s = s & "        lblItem.Caption = ""Select product line""" & vbCrLf
    s = s & "        values = ListExportProductLines()" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    For i = 0 To ExportFormArrayCount(values) - 1" & vbCrLf
    s = s & "        cboItem.AddItem values(LBound(values) + i)" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    If cboItem.ListCount > 0 Then cboItem.ListIndex = 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & vbCrLf
    s = s & "Private Function ExportFormArrayCount(ByRef values() As String) As Long" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    ExportFormArrayCount = UBound(values) - LBound(values) + 1" & vbCrLf
    s = s & "    If Err.Number <> 0 Then" & vbCrLf
    s = s & "        Err.Clear" & vbCrLf
    s = s & "        ExportFormArrayCount = 0" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "End Function" & vbCrLf

    ExportSheetsFormCode = s
End Function
