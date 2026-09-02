Attribute VB_Name = "modFormUI"
Option Explicit

'==============================================================================
' Helpers for building UserForm controls at runtime (paste-in forms).
'==============================================================================

Public Function CreateControlHandlers() As Collection
    Set CreateControlHandlers = New Collection
End Function

Public Const FORM_MARGIN As Single = 12
Public Const FORM_BUTTON_HEIGHT As Single = 26
Public Const FORM_BUTTON_WIDTH As Single = 76
Public Const FORM_BUTTON_GAP As Single = 8

Public Sub ConfigureFormSize(ByVal frm As Object, ByVal widthPoints As Single, ByVal heightPoints As Single)
    frm.Width = widthPoints
    frm.Height = heightPoints
End Sub

Public Function FormButtonRowTop(ByVal contentBottom As Single) As Single
    FormButtonRowTop = contentBottom + 16
End Function

Public Function FormButtonLeft(ByVal buttonIndex As Long, Optional ByVal startLeft As Single = FORM_MARGIN) As Single
    FormButtonLeft = startLeft + (buttonIndex * (FORM_BUTTON_WIDTH + FORM_BUTTON_GAP))
End Function

Public Function FormCloseButtonLeft(ByVal formWidth As Single) As Single
    FormCloseButtonLeft = formWidth - FORM_MARGIN - FORM_BUTTON_WIDTH
End Function

    Set CreateControlHandlers = New Collection
End Function

Public Function AddLabel( _
    ByVal frm As Object, _
    ByVal caption As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal widthPos As Single) As Object

    Dim lbl As Object

    Set lbl = frm.Controls.Add("Forms.Label.1", UniqueControlName("lbl"))
    lbl.Caption = caption
    lbl.Left = leftPos
    lbl.Top = topPos
    lbl.Width = widthPos
    lbl.Height = 15

    Set AddLabel = lbl
End Function

Public Function AddTextBox( _
    ByVal frm As Object, _
    ByVal controlName As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal widthPos As Single, _
    Optional ByVal heightPos As Single = 21) As Object

    Dim txt As Object

    Set txt = frm.Controls.Add("Forms.TextBox.1", controlName)
    txt.Left = leftPos
    txt.Top = topPos
    txt.Width = widthPos
    txt.Height = heightPos

    Set AddTextBox = txt
End Function

Public Function AddCheckBox( _
    ByVal frm As Object, _
    ByVal controlName As String, _
    ByVal caption As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single) As Object

    Dim chk As Object

    Set chk = frm.Controls.Add("Forms.CheckBox.1", controlName)
    chk.Caption = caption
    chk.Left = leftPos
    chk.Top = topPos
    chk.Width = 80
    chk.Height = 18

    Set AddCheckBox = chk
End Function

Public Function AddComboBox( _
    ByVal frm As Object, _
    ByVal controlName As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal widthPos As Single) As Object

    Dim cbo As Object

    Set cbo = frm.Controls.Add("Forms.ComboBox.1", controlName)
    cbo.Left = leftPos
    cbo.Top = topPos
    cbo.Width = widthPos
    cbo.Height = 21
    cbo.Style = fmStyleDropDownList

    Set AddComboBox = cbo
End Function

Public Function AddListBox( _
    ByVal frm As Object, _
    ByVal controlName As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal widthPos As Single, _
    ByVal heightPos As Single, _
    Optional ByVal multiSelect As Boolean = False) As Object

    Dim lst As Object

    Set lst = frm.Controls.Add("Forms.ListBox.1", controlName)
    lst.Left = leftPos
    lst.Top = topPos
    lst.Width = widthPos
    lst.Height = heightPos

    If multiSelect Then
        lst.MultiSelect = fmMultiSelectMulti
    Else
        lst.MultiSelect = fmMultiSelectSingle
    End If

    Set AddListBox = lst
End Function

Public Function AddCommandButton( _
    ByVal frm As Object, _
    ByVal controlName As String, _
    ByVal caption As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal widthPos As Single) As Object

    Dim cmd As Object

    Set cmd = frm.Controls.Add("Forms.CommandButton.1", controlName)
    cmd.Caption = caption
    cmd.Left = leftPos
    cmd.Top = topPos
    cmd.Width = widthPos
    cmd.Height = 24

    Set AddCommandButton = cmd
End Function

Public Sub WireControlEvent( _
    ByVal handlers As Collection, _
    ByVal ctl As Object, _
    ByVal callback As Object, _
    ByVal eventName As String, _
    ByVal eventKind As String)

    Dim handler As clsFormControlHandler

    Set handler = New clsFormControlHandler
    handler.Init ctl, callback, eventName, eventKind
    handlers.Add handler
End Sub

Public Function GetControl(ByVal frm As Object, ByVal controlName As String) As Object
    Set GetControl = frm.Controls(controlName)
End Function

Private Function UniqueControlName(ByVal prefix As String) As String
    Static counter As Long
    counter = counter + 1
    UniqueControlName = prefix & CStr(counter)
End Function

Public Function SelectedComboCode(ByVal comboBox As Object) As String
    If comboBox.ListIndex < 0 Then
        SelectedComboCode = vbNullString
    Else
        SelectedComboCode = ExtractCodeFromDisplayItem(CStr(comboBox.List(comboBox.ListIndex)))
    End If
End Function

Public Function SelectedListCode(ByVal listBox As Object) As String
    If listBox.ListIndex < 0 Then
        SelectedListCode = vbNullString
    Else
        SelectedListCode = ExtractCodeFromDisplayItem(CStr(listBox.List(listBox.ListIndex)))
    End If
End Function

Public Function SelectedListText(ByVal listBox As Object) As String
    If listBox.ListIndex < 0 Then
        SelectedListText = vbNullString
    Else
        SelectedListText = CStr(listBox.List(listBox.ListIndex))
    End If
End Function

Public Sub SelectListItemByCode(ByVal listBox As Object, ByVal codeValue As String)
    Dim itemIndex As Long

    For itemIndex = 0 To listBox.ListCount - 1
        If ValuesMatchCode(ExtractCodeFromDisplayItem(CStr(listBox.List(itemIndex))), codeValue) Then
            listBox.ListIndex = itemIndex
            Exit Sub
        End If
    Next itemIndex

    listBox.ListIndex = -1
End Sub

Public Sub SelectComboByCode(ByVal comboBox As Object, ByVal codeValue As String)
    Dim itemIndex As Long

    For itemIndex = 0 To comboBox.ListCount - 1
        If ValuesMatchCode(ExtractCodeFromDisplayItem(CStr(comboBox.List(itemIndex))), codeValue) Then
            comboBox.ListIndex = itemIndex
            Exit Sub
        End If
    Next itemIndex

    comboBox.ListIndex = -1
End Sub

Public Function SelectedListBoxCodes(ByVal listBox As Object) As Variant
    Dim selectedCodes() As String
    Dim selectedCount As Long
    Dim itemIndex As Long

    selectedCount = 0
    ReDim selectedCodes(0 To 0)

    For itemIndex = 0 To listBox.ListCount - 1
        If listBox.Selected(itemIndex) Then
            If selectedCount > 0 Then
                ReDim Preserve selectedCodes(0 To selectedCount)
            End If
            selectedCodes(selectedCount) = ExtractCodeFromDisplayItem(CStr(listBox.List(itemIndex)))
            selectedCount = selectedCount + 1
        End If
    Next itemIndex

    If selectedCount = 0 Then
        ReDim selectedCodes(0 To -1)
    End If

    SelectedListBoxCodes = selectedCodes
End Function

Public Function IsArrayAllocated(ByVal arr As Variant) As Boolean
    On Error Resume Next
    IsArrayAllocated = IsArray(arr) And (UBound(arr) >= LBound(arr))
    On Error GoTo 0
End Function
