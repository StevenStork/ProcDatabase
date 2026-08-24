Attribute VB_Name = "modUi"
Option Compare Database
Option Explicit

' Creates (or recreates) the Home / Part / reference / export forms.
' Safe to re-run: existing forms are deleted first.

Public Sub EnsureUi()
    CreateReferenceForms
    CreateHomeForm
    CreatePartDashSubform
    CreatePartProductLineSubform
    CreateOperationSubform
    CreatePartForm
    CreateReferencesForm
    CreateExportForm
End Sub

Private Sub CreateReferenceForms()
    CreateSimpleTableForm FRM_FFA, TBL_FFA, "FFAs"
    CreateSimpleTableForm FRM_PRODUCT_LINE, TBL_PRODUCT_LINE, "Product Lines"
    CreateSimpleTableForm FRM_EQUIPMENT, TBL_EQUIPMENT, "Equipment"
End Sub

Private Sub CreateSimpleTableForm(ByVal formName As String, ByVal tableName As String, ByVal caption As String)
    Dim frm As Form
    DeleteObjectIfExists acForm, formName
    Set frm = CreateForm()
    frm.RecordSource = tableName
    frm.Caption = caption
    frm.DefaultView = 2
    frm.AllowAdditions = True
    frm.AllowDeletions = True
    frm.AllowEdits = True
    DoCmd.Close acForm, frm.Name, acSaveYes
    RenameLastForm formName
End Sub

Private Sub CreateHomeForm()
    Dim frm As Form
    Dim ctl As Control

    DeleteObjectIfExists acForm, FRM_HOME
    Set frm = CreateForm()
    frm.RecordSource = TBL_PART
    frm.Caption = "ProcDatabase Home"
    frm.DefaultView = 2 ' Datasheet
    frm.AllowAdditions = False
    frm.AllowDeletions = False
    frm.AllowEdits = True

    AddBoundTextBox frm, "BasePart", "Base Part", 0, 0, 2000
    AddBoundCheckBox frm, "Active", "Active", 2100, 0
    AddBoundTextBox frm, "StatusDate", "Date", 2600, 0, 1400

    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , 4100, 0, 900, 300)
    ctl.Name = "txtDays"
    ctl.ControlSource = "=IIf(IsNull([StatusDate]),Null,DateDiff(""d"",[StatusDate],Date()))"
    ApplyDaysRagFormat ctl

    AddBoundTextBox frm, "Highlight", "Highlight", 5100, 0, 1600
    AddBoundComboBox frm, "HomeFFA", "Home FFA", 6800, 0, 1600, _
        "SELECT FFA FROM [" & TBL_FFA & "] ORDER BY FFA"

    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , 8500, 0, 2000, 300)
    ctl.Name = "txtFactories"
    ctl.ControlSource = "=DLookup(""Factory"",""" & TBL_FFA & """,""FFA='"" & Replace(Nz([HomeFFA],""),""'"",""''"") & ""'"")"
    ctl.Enabled = False

    AddBoundTextBox frm, COL_SHEET_NAME, "Sheet", 10600, 0, 1600

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 0, 400, 1800, 360)
    ctl.Name = "btnRefresh"
    ctl.Caption = "Refresh Linked Data"
    ctl.OnClick = "=UiRefreshLinkedData()"

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 1900, 400, 1800, 360)
    ctl.Name = "btnOpenPart"
    ctl.Caption = "Open Part"
    ctl.OnClick = "=OpenSelectedPart()"

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 3800, 400, 1800, 360)
    ctl.Name = "btnSeedOps"
    ctl.Caption = "Seed Active Ops"
    ctl.OnClick = "=UiSeedActiveOps()"

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 5700, 400, 1600, 360)
    ctl.Name = "btnReferences"
    ctl.Caption = "References"
    ctl.OnClick = "=UiOpenReferences()"

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 7400, 400, 1400, 360)
    ctl.Name = "btnExport"
    ctl.Caption = "Export"
    ctl.OnClick = "=UiOpenExport()"

    DoCmd.Close acForm, frm.Name, acSaveYes
    RenameLastForm FRM_HOME
End Sub

Private Sub CreatePartDashSubform()
    Dim frm As Form
    DeleteObjectIfExists acForm, SFRM_DASH
    Set frm = CreateForm()
    frm.RecordSource = "SELECT * FROM [" & TBL_PART_DASH & "]"
    frm.DefaultView = 2
    frm.AllowAdditions = False
    AddBoundTextBox frm, "Dash", "Dash", 0, 0, 1400
    AddBoundCheckBox frm, "Active", "Active", 1500, 0
    DoCmd.Close acForm, frm.Name, acSaveYes
    RenameLastForm SFRM_DASH
End Sub

Private Sub CreatePartProductLineSubform()
    Dim frm As Form
    DeleteObjectIfExists acForm, SFRM_PL
    Set frm = CreateForm()
    frm.RecordSource = "SELECT * FROM [" & TBL_PART_PL & "]"
    frm.DefaultView = 2
    frm.AllowAdditions = False
    AddBoundTextBox frm, "ProductLine", "Product Line", 0, 0, 2000
    AddBoundCheckBox frm, "UseFlag", "Use", 2100, 0
    DoCmd.Close acForm, frm.Name, acSaveYes
    RenameLastForm SFRM_PL
End Sub

Private Sub CreateOperationSubform()
    Dim frm As Form
    DeleteObjectIfExists acForm, SFRM_OPS
    Set frm = CreateForm()
    frm.RecordSource = QRY_OPERATIONS
    frm.DefaultView = 2
    frm.AllowAdditions = True
    frm.AllowDeletions = True
    AddBoundTextBox frm, "OpSequence", "Op Seq", 0, 0, 900
    AddBoundTextBox frm, "OpCode", "Op Code", 1000, 0, 1100
    AddBoundTextBox frm, "ImportedHours", "Imported Hours", 2200, 0, 1200
    AddBoundTextBox frm, "ImportedEx", "Imported Ex", 3500, 0, 1100
    AddBoundTextBox frm, "BatchSize", "Batch", 4700, 0, 900
    AddBoundTextBox frm, "ExportHours", "Export Hours", 5700, 0, 1200
    AddBoundTextBox frm, "ExportEx", "Export Ex", 7000, 0, 1100
    AddBoundTextBox frm, "EquipmentType", "Equipment", 8200, 0, 1400
    AddBoundCheckBox frm, "UseExportHours", "Use Hrs", 9700, 0
    AddBoundCheckBox frm, "UseExportEx", "Use Ex", 10400, 0
    AddBoundTextBox frm, "ProcessHours", "Process Hours", 11100, 0, 1200
    AddBoundTextBox frm, "AvgEx", "Avg Ex", 12400, 0, 1000
    AddBoundTextBox frm, "AvgHPU", "Avg HPU", 13500, 0, 1000
    AddBoundTextBox frm, "MadeInFFA", "Made In FFA", 14600, 0, 1400
    DoCmd.Close acForm, frm.Name, acSaveYes
    RenameLastForm SFRM_OPS
End Sub

Private Sub CreatePartForm()
    Dim frm As Form
    Dim ctl As Control

    DeleteObjectIfExists acForm, FRM_PART
    Set frm = CreateForm()
    frm.RecordSource = "SELECT * FROM [" & TBL_PART & "]"
    frm.Caption = "Part"
    frm.DefaultView = 0 ' Single Form
    frm.AllowAdditions = False

    AddBoundTextBox frm, "BasePart", "Base Part", 1200, 200, 2000
    AddBoundCheckBox frm, "Active", "Active", 3400, 200
    AddBoundComboBox frm, "HomeFFA", "Home FFA", 1200, 600, 2000, _
        "SELECT FFA FROM [" & TBL_FFA & "] ORDER BY FFA"
    AddBoundTextBox frm, "StatusDate", "Status Date", 1200, 1000, 1600

    Set ctl = CreateControl(frm.Name, acSubform, acDetail, , , 200, 1600, 3200, 2400)
    ctl.Name = "subDashes"
    ctl.SourceObject = SFRM_DASH
    ctl.LinkMasterFields = "BasePart"
    ctl.LinkChildFields = "BasePart"

    Set ctl = CreateControl(frm.Name, acSubform, acDetail, , , 3600, 1600, 3600, 2400)
    ctl.Name = "subProductLines"
    ctl.SourceObject = SFRM_PL
    ctl.LinkMasterFields = "BasePart"
    ctl.LinkChildFields = "BasePart"

    Set ctl = CreateControl(frm.Name, acSubform, acDetail, , , 200, 4200, 9000, 3600)
    ctl.Name = "subOperations"
    ctl.SourceObject = SFRM_OPS
    ctl.LinkMasterFields = "BasePart"
    ctl.LinkChildFields = "BasePart"

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 7400, 200, 1600, 360)
    ctl.Name = "btnSeed"
    ctl.Caption = "Seed Ops"
    ctl.OnClick = "=SeedOperationsForCurrentPart()"

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 7400, 600, 1600, 360)
    ctl.Name = "btnClose"
    ctl.Caption = "Close"
    ctl.OnClick = "=UiCloseCurrentForm()"

    DoCmd.Close acForm, frm.Name, acSaveYes
    RenameLastForm FRM_PART
End Sub

Private Sub CreateReferencesForm()
    Dim frm As Form
    Dim ctl As Control

    DeleteObjectIfExists acForm, FRM_REFERENCES
    Set frm = CreateForm()
    frm.Caption = "References"
    frm.RecordSource = vbNullString
    frm.DefaultView = 0

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 200, 200, 2000, 400)
    ctl.Name = "btnFFA"
    ctl.Caption = "Edit FFAs"
    ctl.OnClick = "=UiOpenFfaForm()"

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 200, 700, 2000, 400)
    ctl.Name = "btnPL"
    ctl.Caption = "Edit Product Lines"
    ctl.OnClick = "=UiOpenProductLineForm()"

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 200, 1200, 2000, 400)
    ctl.Name = "btnEquip"
    ctl.Caption = "Edit Equipment"
    ctl.OnClick = "=UiOpenEquipmentForm()"

    DoCmd.Close acForm, frm.Name, acSaveYes
    RenameLastForm FRM_REFERENCES
End Sub

Private Sub CreateExportForm()
    Dim frm As Form
    Dim ctl As Control

    DeleteObjectIfExists acForm, FRM_EXPORT
    Set frm = CreateForm()
    frm.Caption = "Export"
    frm.RecordSource = vbNullString

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 200, 200, 2200, 400)
    ctl.Name = "btnFfa"
    ctl.Caption = "Export FFA"
    ctl.OnClick = "=UiExportFfa()"

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 200, 700, 2200, 400)
    ctl.Name = "btnAll"
    ctl.Caption = "Export All"
    ctl.OnClick = "=UiExportAll()"

    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, , , 200, 1300, 2200, 360)
    ctl.Name = "cboProductLine"
    ctl.RowSource = "SELECT ProductLine FROM [" & TBL_PRODUCT_LINE & "] ORDER BY ProductLine"
    ctl.RowSourceType = "Table/Query"
    ctl.LimitToList = True

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 2600, 1300, 1800, 400)
    ctl.Name = "btnPl"
    ctl.Caption = "Export Product Line"
    ctl.OnClick = "=ExportSelectedProductLine()"

    DoCmd.Close acForm, frm.Name, acSaveYes
    RenameLastForm FRM_EXPORT
End Sub

Private Sub AddBoundTextBox(ByVal frm As Form, ByVal fieldName As String, ByVal caption As String, _
    ByVal leftPos As Long, ByVal topPos As Long, ByVal widthPos As Long)
    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , fieldName, leftPos, topPos, widthPos, 300)
    ctl.Name = "txt" & fieldName
    ctl.ControlSource = fieldName
End Sub

Private Sub AddBoundCheckBox(ByVal frm As Form, ByVal fieldName As String, ByVal caption As String, _
    ByVal leftPos As Long, ByVal topPos As Long)
    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acCheckBox, acDetail, , fieldName, leftPos, topPos, 300, 300)
    ctl.Name = "chk" & fieldName
    ctl.ControlSource = fieldName
End Sub

Private Sub AddBoundComboBox(ByVal frm As Form, ByVal fieldName As String, ByVal caption As String, _
    ByVal leftPos As Long, ByVal topPos As Long, ByVal widthPos As Long, ByVal rowSource As String)
    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, , fieldName, leftPos, topPos, widthPos, 300)
    ctl.Name = "cbo" & fieldName
    ctl.ControlSource = fieldName
    ctl.RowSource = rowSource
    ctl.RowSourceType = "Table/Query"
    ctl.LimitToList = False
End Sub

Private Sub RenameLastForm(ByVal desiredName As String)
    Dim i As Long
    Dim createdName As String
    Dim frm As Object

    ' CreateForm leaves a FormN object; rename the most recently created unmatched Form*.
    For i = CurrentProject.AllForms.Count - 1 To 0 Step -1
        createdName = CurrentProject.AllForms(i).Name
        If StrComp(createdName, desiredName, vbTextCompare) = 0 Then Exit Sub
        If Left$(createdName, 4) = "Form" And IsNumeric(Mid$(createdName, 5)) Then
            DoCmd.Rename desiredName, acForm, createdName
            Exit Sub
        End If
    Next i
End Sub

Public Function OpenSelectedPart() As Boolean
    Dim basePart As String
    On Error GoTo Fail
    If Not CurrentProject.AllForms(FRM_HOME).IsLoaded Then
        OpenSelectedPart = False
        Exit Function
    End If
    basePart = CoerceText(Forms(FRM_HOME)!BasePart)
    If Len(basePart) = 0 Then
        MsgBox "Select a part row first.", vbExclamation, "Open Part"
        OpenSelectedPart = False
        Exit Function
    End If
    DoCmd.OpenForm FRM_PART, , , "BasePart = " & SqlText(basePart)
    OpenSelectedPart = True
    Exit Function
Fail:
    OpenSelectedPart = False
End Function

Public Function SeedOperationsForCurrentPart() As Boolean
    Dim basePart As String
    On Error GoTo Fail
    basePart = CoerceText(Forms(FRM_PART)!BasePart)
    SeedOperationsForPart basePart
    Forms(FRM_PART)!subOperations.Requery
    SeedOperationsForCurrentPart = True
    Exit Function
Fail:
    SeedOperationsForCurrentPart = False
End Function

Public Function ExportSelectedProductLine() As Boolean
    Dim productLine As String
    On Error GoTo Fail
    productLine = CoerceText(Forms(FRM_EXPORT)!cboProductLine)
    If Len(productLine) = 0 Then
        MsgBox "Choose a product line.", vbExclamation, "Export"
        ExportSelectedProductLine = False
        Exit Function
    End If
    ExportOps "PRODUCTLINE", productLine
    ExportSelectedProductLine = True
    Exit Function
Fail:
    ExportSelectedProductLine = False
End Function

Public Function UiRefreshLinkedData() As Boolean
    On Error GoTo Fail
    RefreshAll
    UiRefreshLinkedData = True
    Exit Function
Fail:
    UiRefreshLinkedData = False
End Function

Public Function UiSeedActiveOps() As Boolean
    On Error GoTo Fail
    SeedOperationsForActiveParts
    UiSeedActiveOps = True
    Exit Function
Fail:
    UiSeedActiveOps = False
End Function

Public Function UiOpenReferences() As Boolean
    On Error GoTo Fail
    DoCmd.OpenForm FRM_REFERENCES
    UiOpenReferences = True
    Exit Function
Fail:
    UiOpenReferences = False
End Function

Public Function UiOpenExport() As Boolean
    On Error GoTo Fail
    DoCmd.OpenForm FRM_EXPORT
    UiOpenExport = True
    Exit Function
Fail:
    UiOpenExport = False
End Function

Public Function UiCloseCurrentForm() As Boolean
    On Error Resume Next
    DoCmd.Close acForm
    UiCloseCurrentForm = True
End Function

Public Function UiOpenFfaForm() As Boolean
    On Error GoTo Fail
    DoCmd.OpenForm FRM_FFA
    UiOpenFfaForm = True
    Exit Function
Fail:
    UiOpenFfaForm = False
End Function

Public Function UiOpenProductLineForm() As Boolean
    On Error GoTo Fail
    DoCmd.OpenForm FRM_PRODUCT_LINE
    UiOpenProductLineForm = True
    Exit Function
Fail:
    UiOpenProductLineForm = False
End Function

Public Function UiOpenEquipmentForm() As Boolean
    On Error GoTo Fail
    DoCmd.OpenForm FRM_EQUIPMENT
    UiOpenEquipmentForm = True
    Exit Function
Fail:
    UiOpenEquipmentForm = False
End Function

Public Function UiOpenFfaTable() As Boolean
    UiOpenFfaTable = UiOpenFfaForm()
End Function

Public Function UiOpenProductLineTable() As Boolean
    UiOpenProductLineTable = UiOpenProductLineForm()
End Function

Public Function UiOpenEquipmentTable() As Boolean
    UiOpenEquipmentTable = UiOpenEquipmentForm()
End Function

Public Function UiExportFfa() As Boolean
    On Error GoTo Fail
    ExportOps "FFA"
    UiExportFfa = True
    Exit Function
Fail:
    UiExportFfa = False
End Function

Public Function UiExportAll() As Boolean
    On Error GoTo Fail
    ExportOps "ALL"
    UiExportAll = True
    Exit Function
Fail:
    UiExportAll = False
End Function
