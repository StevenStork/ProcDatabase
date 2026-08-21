VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReferences 
   Caption         =   "Update References"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Layout is built at runtime so this .frm does not need a .frx designer blob.

Private lblCategory As MSForms.Label
Private WithEvents cboCategory As MSForms.ComboBox
Private lblItems As MSForms.Label
Private WithEvents lstItems As MSForms.ListBox
Private lblName As MSForms.Label
Private txtName As MSForms.TextBox
Private lblFactory As MSForms.Label
Private txtFactory As MSForms.TextBox
Private lblOwners As MSForms.Label
Private lstOwners As MSForms.ListBox
Private WithEvents cmdAdd As MSForms.CommandButton
Private WithEvents cmdUpdate As MSForms.CommandButton
Private WithEvents cmdDelete As MSForms.CommandButton
Private WithEvents cmdSave As MSForms.CommandButton
Private WithEvents cmdCancel As MSForms.CommandButton

Private m_ffas As Object
Private m_productLines As Object
Private m_equipment As Object
Private m_dirty As Boolean
Private m_suppress As Boolean

Private Sub UserForm_Initialize()
    Dim data As Object

    Me.Caption = "Update References"
    Me.Width = 520
    Me.Height = 400

    Set lblCategory = AddLabel("lblCategory", 12, 12, 200, 16, "Category")
    Set cboCategory = AddCombo("cboCategory", 12, 28, 200, 22)
    Set lblItems = AddLabel("lblItems", 12, 58, 200, 16, "Items")
    Set lstItems = AddList("lstItems", 12, 74, 200, 220, False)

    Set lblName = AddLabel("lblName", 228, 58, 268, 16, "Name")
    Set txtName = AddText("txtName", 228, 74, 268, 22)
    Set lblFactory = AddLabel("lblFactory", 228, 104, 268, 16, "Factory")
    Set txtFactory = AddText("txtFactory", 228, 120, 268, 22)
    Set lblOwners = AddLabel("lblOwners", 228, 104, 268, 16, "Owning FFAs")
    Set lstOwners = AddList("lstOwners", 228, 120, 268, 174, True)

    Set cmdAdd = AddButton("cmdAdd", 228, 304, 84, 28, "Add")
    Set cmdUpdate = AddButton("cmdUpdate", 320, 304, 84, 28, "Update")
    Set cmdDelete = AddButton("cmdDelete", 412, 304, 84, 28, "Delete")
    Set cmdSave = AddButton("cmdSave", 320, 340, 84, 28, "Save")
    Set cmdCancel = AddButton("cmdCancel", 412, 340, 84, 28, "Cancel")

    Set data = LoadReferenceData()
    Set m_ffas = data("Ffas")
    Set m_productLines = data("ProductLines")
    Set m_equipment = data("Equipment")
    m_dirty = False

    With cboCategory
        .Clear
        .AddItem REFS_CATEGORY_FFA
        .AddItem REFS_CATEGORY_PRODUCT_LINE
        .AddItem REFS_CATEGORY_EQUIPMENT
        .ListIndex = 0
    End With
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

Private Function AddText( _
    ByVal controlName As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal widthPos As Single, _
    ByVal heightPos As Single) As MSForms.TextBox

    Dim ctl As MSForms.TextBox
    Set ctl = Me.Controls.Add("Forms.TextBox.1", controlName, True)
    ctl.Left = leftPos
    ctl.Top = topPos
    ctl.Width = widthPos
    ctl.Height = heightPos
    Set AddText = ctl
End Function

Private Function AddList( _
    ByVal controlName As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal widthPos As Single, _
    ByVal heightPos As Single, _
    ByVal multiSelect As Boolean) As MSForms.ListBox

    Dim ctl As MSForms.ListBox
    Set ctl = Me.Controls.Add("Forms.ListBox.1", controlName, True)
    ctl.Left = leftPos
    ctl.Top = topPos
    ctl.Width = widthPos
    ctl.Height = heightPos
    If multiSelect Then
        ctl.MultiSelect = fmMultiSelectMulti
    Else
        ctl.MultiSelect = fmMultiSelectSingle
    End If
    Set AddList = ctl
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

Private Sub cboCategory_Change()
    ApplyCategoryUi
End Sub

Private Sub lstItems_Click()
    If m_suppress Then Exit Sub
    LoadSelectedItem
End Sub

Private Sub cmdAdd_Click()
    Dim itemName As String
    Dim owners As String

    itemName = Trim$(txtName.Value)
    If Len(itemName) = 0 Then
        MsgBox "Enter a name.", vbExclamation, Me.Caption
        Exit Sub
    End If

    Select Case CurrentCategory()
        Case REFS_CATEGORY_FFA
            If m_ffas.Exists(itemName) Then
                MsgBox "That FFA already exists.", vbExclamation, Me.Caption
                Exit Sub
            End If
            m_ffas.Add itemName, Trim$(txtFactory.Value)
        Case REFS_CATEGORY_PRODUCT_LINE
            If m_productLines.Exists(itemName) Then
                MsgBox "That product line already exists.", vbExclamation, Me.Caption
                Exit Sub
            End If
            m_productLines.Add itemName, itemName
        Case REFS_CATEGORY_EQUIPMENT
            If m_equipment.Exists(itemName) Then
                MsgBox "That equipment already exists.", vbExclamation, Me.Caption
                Exit Sub
            End If
            owners = SelectedOwnerList()
            m_equipment.Add itemName, NormalizeOwnerList(owners, m_ffas)
        Case Else
            Exit Sub
    End Select

    m_dirty = True
    RefreshItemList itemName
End Sub

Private Sub cmdUpdate_Click()
    Dim oldName As String
    Dim newName As String
    Dim owners As String

    oldName = SelectedItemName()
    newName = Trim$(txtName.Value)
    If Len(oldName) = 0 Then
        MsgBox "Select an item to update.", vbExclamation, Me.Caption
        Exit Sub
    End If
    If Len(newName) = 0 Then
        MsgBox "Enter a name.", vbExclamation, Me.Caption
        Exit Sub
    End If

    Select Case CurrentCategory()
        Case REFS_CATEGORY_FFA
            If StrComp(oldName, newName, vbTextCompare) <> 0 Then
                If m_ffas.Exists(newName) Then
                    MsgBox "That FFA already exists.", vbExclamation, Me.Caption
                    Exit Sub
                End If
                m_ffas.Remove oldName
                Set m_equipment = RenameFfaInEquipment(m_equipment, oldName, newName)
            Else
                m_ffas.Remove oldName
            End If
            m_ffas.Add newName, Trim$(txtFactory.Value)
        Case REFS_CATEGORY_PRODUCT_LINE
            If StrComp(oldName, newName, vbTextCompare) <> 0 Then
                If m_productLines.Exists(newName) Then
                    MsgBox "That product line already exists.", vbExclamation, Me.Caption
                    Exit Sub
                End If
            End If
            m_productLines.Remove oldName
            m_productLines.Add newName, newName
        Case REFS_CATEGORY_EQUIPMENT
            If StrComp(oldName, newName, vbTextCompare) <> 0 Then
                If m_equipment.Exists(newName) Then
                    MsgBox "That equipment already exists.", vbExclamation, Me.Caption
                    Exit Sub
                End If
                m_equipment.Remove oldName
            Else
                m_equipment.Remove oldName
            End If
            owners = SelectedOwnerList()
            m_equipment.Add newName, NormalizeOwnerList(owners, m_ffas)
        Case Else
            Exit Sub
    End Select

    m_dirty = True
    RefreshItemList newName
End Sub

Private Sub cmdDelete_Click()
    Dim itemName As String
    Dim response As VbMsgBoxResult

    itemName = SelectedItemName()
    If Len(itemName) = 0 Then
        MsgBox "Select an item to delete.", vbExclamation, Me.Caption
        Exit Sub
    End If

    response = MsgBox("Delete """ & itemName & """ from " & CurrentCategory() & "?", _
        vbExclamation + vbYesNo + vbDefaultButton2, Me.Caption)
    If response <> vbYes Then Exit Sub

    Select Case CurrentCategory()
        Case REFS_CATEGORY_FFA
            m_ffas.Remove itemName
            Set m_equipment = RemoveFfaFromEquipment(m_equipment, itemName)
        Case REFS_CATEGORY_PRODUCT_LINE
            m_productLines.Remove itemName
        Case REFS_CATEGORY_EQUIPMENT
            m_equipment.Remove itemName
        Case Else
            Exit Sub
    End Select

    m_dirty = True
    RefreshItemList vbNullString
    ClearEditor
End Sub

Private Sub cmdSave_Click()
    Dim data As Object

    Set data = CreateObject("Scripting.Dictionary")
    data.CompareMode = vbTextCompare
    data.Add "Ffas", m_ffas
    data.Add "ProductLines", m_productLines
    data.Add "Equipment", m_equipment
    SaveReferenceData data
    m_dirty = False
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Dim response As VbMsgBoxResult

    If m_dirty Then
        response = MsgBox("Discard unsaved reference changes?", _
            vbExclamation + vbYesNo + vbDefaultButton2, Me.Caption)
        If response <> vbYes Then Exit Sub
    End If
    Unload Me
End Sub

Private Sub ApplyCategoryUi()
    Dim isFfa As Boolean
    Dim isEquipment As Boolean

    isFfa = (StrComp(CurrentCategory(), REFS_CATEGORY_FFA, vbTextCompare) = 0)
    isEquipment = (StrComp(CurrentCategory(), REFS_CATEGORY_EQUIPMENT, vbTextCompare) = 0)

    lblFactory.Visible = isFfa
    txtFactory.Visible = isFfa
    lblOwners.Visible = isEquipment
    lstOwners.Visible = isEquipment

    If isFfa Then
        lblItems.Caption = "FFAs"
        lblName.Caption = "FFA name"
    ElseIf isEquipment Then
        lblItems.Caption = "Equipment"
        lblName.Caption = "Equipment name"
        RefreshOwnerList vbNullString
    Else
        lblItems.Caption = "Product lines"
        lblName.Caption = "Product line name"
    End If

    RefreshItemList vbNullString
    ClearEditor
End Sub

Private Sub RefreshItemList(ByVal selectName As String)
    Dim keys() As String
    Dim i As Long
    Dim selectIndex As Long

    m_suppress = True
    lstItems.Clear
    selectIndex = -1
    keys = SortedDictionaryKeys(CurrentMap())

    On Error Resume Next
    For i = 0 To UBound(keys)
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
        lstItems.AddItem keys(i)
        If Len(selectName) > 0 Then
            If StrComp(keys(i), selectName, vbTextCompare) = 0 Then selectIndex = i
        End If
    Next i
    On Error GoTo 0

    If selectIndex >= 0 Then lstItems.ListIndex = selectIndex
    m_suppress = False
    If selectIndex >= 0 Then LoadSelectedItem
End Sub

Private Sub RefreshOwnerList(ByVal selectedOwners As String)
    Dim keys() As String
    Dim i As Long
    Dim ownerName As String

    lstOwners.Clear
    keys = SortedDictionaryKeys(m_ffas)

    On Error Resume Next
    For i = 0 To UBound(keys)
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
        ownerName = keys(i)
        lstOwners.AddItem ownerName
        If InStr(1, ", " & selectedOwners & ", ", ", " & ownerName & ", ", vbTextCompare) > 0 Then
            lstOwners.Selected(lstOwners.ListCount - 1) = True
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub LoadSelectedItem()
    Dim itemName As String

    itemName = SelectedItemName()
    If Len(itemName) = 0 Then
        ClearEditor
        Exit Sub
    End If

    txtName.Value = itemName

    Select Case CurrentCategory()
        Case REFS_CATEGORY_FFA
            txtFactory.Value = CStr(m_ffas(itemName))
        Case REFS_CATEGORY_EQUIPMENT
            RefreshOwnerList CStr(m_equipment(itemName))
        Case Else
            txtFactory.Value = vbNullString
    End Select
End Sub

Private Sub ClearEditor()
    txtName.Value = vbNullString
    txtFactory.Value = vbNullString
    If CurrentCategory() = REFS_CATEGORY_EQUIPMENT Then RefreshOwnerList vbNullString
End Sub

Private Function CurrentCategory() As String
    If cboCategory Is Nothing Then Exit Function
    CurrentCategory = Trim$(CStr(cboCategory.Value))
End Function

Private Function CurrentMap() As Object
    Select Case CurrentCategory()
        Case REFS_CATEGORY_FFA
            Set CurrentMap = m_ffas
        Case REFS_CATEGORY_PRODUCT_LINE
            Set CurrentMap = m_productLines
        Case Else
            Set CurrentMap = m_equipment
    End Select
End Function

Private Function SelectedItemName() As String
    If lstItems.ListIndex < 0 Then Exit Function
    SelectedItemName = Trim$(CStr(lstItems.Value))
End Function

Private Function SelectedOwnerList() As String
    Dim i As Long
    Dim names As Collection
    Dim parts() As String

    Set names = New Collection
    For i = 0 To lstOwners.ListCount - 1
        If lstOwners.Selected(i) Then names.Add CStr(lstOwners.List(i))
    Next i

    If names.Count = 0 Then
        SelectedOwnerList = vbNullString
        Exit Function
    End If

    ReDim parts(0 To names.Count - 1)
    For i = 1 To names.Count
        parts(i - 1) = CStr(names(i))
    Next i
    SelectedOwnerList = Join(parts, ", ")
End Function
