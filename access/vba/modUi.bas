Attribute VB_Name = "modUi"
Option Compare Database
Option Explicit

' Creates (or recreates) the Home / Part / reference / export forms.
' Safe to re-run: existing forms are deleted first.

Public UiSubStep As String

Public Sub EnsureUi()
    On Error GoTo Fail
    UiSubStep = "CreateReferenceForms"
    CreateReferenceForms
    UiSubStep = "CreatePartDashSubform"
    CreatePartDashSubform
    UiSubStep = "CreatePartProductLineSubform"
    CreatePartProductLineSubform
    UiSubStep = "CreateOperationSubform"
    CreateOperationSubform
    UiSubStep = "CreatePartForm"
    CreatePartForm
    UiSubStep = "CreateHomeForm"
    CreateHomeForm
    UiSubStep = "CreateReferencesForm"
    CreateReferencesForm
    UiSubStep = "CreateExportForm"
    CreateExportForm
    UiSubStep = vbNullString
    Exit Sub
Fail:
    Err.Raise Err.Number, "EnsureUi." & UiSubStep, Err.Description
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
    SaveAndRenameForm frm, formName
End Sub

Private Sub CreateHomeForm()
    Dim frm As Form

    DeleteObjectIfExists acForm, FRM_HOME
    Set frm = CreateForm()
    frm.RecordSource = QRY_HOME
    frm.Caption = "ProcDatabase Home"
    frm.DefaultView = 1
    frm.AllowAdditions = False
    frm.AllowDeletions = False
    frm.AllowEdits = True
    frm.ScrollBars = 2
    frm.RecordSelectors = True
    frm.NavigationButtons = True
    On Error Resume Next
    frm.Section(acHeader).Visible = True
    On Error GoTo 0

    AddHomeField frm, "BasePart", 120, 600, 1800
    AddHomeField frm, "Active", 2040, 600, 900, True
    AddHomeField frm, "StatusDate", 3000, 600, 1200
    AddHomeField frm, "Days", 4320, 600, 720
    frm.Controls("txtDays").Enabled = False
    AddHomeField frm, "Highlight", 5100, 600, 1200
    AddHomeCombo frm, "HomeFFA", 6360, 600, 1440
    AddHomeField frm, "Factories", 7860, 600, 1800
    frm.Controls("txtFactories").Enabled = False
    AddHomeField frm, COL_SHEET_NAME, 9720, 600, 1200

    AddHeaderButton frm, "btnRefresh", "Refresh Linked Data", 120, 120, "=UiRefreshLinkedData()", 2160
    AddHeaderButton frm, "btnOpenPart", "Open Part", 2400, 120, "=OpenSelectedPart()", 1440
    AddHeaderButton frm, "btnSeedOps", "Seed Active Ops", 3960, 120, "=UiSeedActiveOps()", 1800
    AddHeaderButton frm, "btnReferences", "References", 5880, 120, "=UiOpenReferences()", 1440
    AddHeaderButton frm, "btnExport", "Export", 7440, 120, "=UiOpenExport()", 1200

    SaveAndRenameForm frm, FRM_HOME
    ApplyHomeDaysRagDesign FRM_HOME
End Sub

Private Sub AddHomeField(ByVal frm As Form, ByVal fieldName As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthPos As Long, Optional ByVal isCheckBox As Boolean = False)
    Dim ctl As Control
    If isCheckBox Then
        Set ctl = CreateControl(frm.Name, acCheckBox, acDetail, , fieldName, leftPos, topPos, 300, 300)
        ctl.Name = "chk" & fieldName
    Else
        Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , fieldName, leftPos, topPos, widthPos, 300)
        ctl.Name = "txt" & fieldName
    End If
    ctl.ControlSource = fieldName
End Sub

Private Sub AddHomeCombo(ByVal frm As Form, ByVal fieldName As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthPos As Long)
    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, , fieldName, leftPos, topPos, widthPos, 300)
    ctl.Name = "cbo" & fieldName
    ctl.ControlSource = fieldName
    ctl.RowSource = "SELECT FFA FROM [" & TBL_FFA & "] ORDER BY FFA"
    ctl.RowSourceType = "Table/Query"
    ctl.LimitToList = False
End Sub

Private Sub AddHeaderButton(ByVal frm As Form, ByVal ctlName As String, ByVal caption As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal onClick As String, ByVal widthPos As Long)
    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acCommandButton, acHeader, , , leftPos, topPos, widthPos, 360)
    ctl.Name = ctlName
    ctl.Caption = caption
    ctl.OnClick = onClick
End Sub

Private Sub ApplyHomeDaysRagDesign(ByVal formName As String)
    On Error GoTo CleanUp
    DoCmd.OpenForm formName, acDesign
    ApplyDaysRagFormat Forms(formName).Controls("txtDays")
CleanUp:
    On Error Resume Next
    DoCmd.Close acForm, formName, acSaveYes
    On Error GoTo 0
End Sub

Private Sub CreatePartDashSubform()
    Dim frm As Form
    DeleteObjectIfExists acForm, SFRM_DASH
    Set frm = CreateForm()
    frm.RecordSource = TBL_PART_DASH
    frm.DefaultView = 2
    frm.AllowAdditions = False
    AddDetailField frm, "Dash", 0, 0, 1400
    AddDetailCheck frm, "Active", 1500, 0
    SaveAndRenameForm frm, SFRM_DASH
End Sub

Private Sub CreatePartProductLineSubform()
    Dim frm As Form
    DeleteObjectIfExists acForm, SFRM_PL
    Set frm = CreateForm()
    frm.RecordSource = TBL_PART_PL
    frm.DefaultView = 2
    frm.AllowAdditions = False
    AddDetailField frm, "ProductLine", 0, 0, 2000
    AddDetailCheck frm, "UseFlag", 2100, 0
    SaveAndRenameForm frm, SFRM_PL
End Sub

Private Sub CreateOperationSubform()
    Dim frm As Form
    DeleteObjectIfExists acForm, SFRM_OPS
    Set frm = CreateForm()
    frm.RecordSource = QRY_OPERATIONS
    frm.DefaultView = 2
    frm.AllowAdditions = True
    frm.AllowDeletions = True
    AddDetailField frm, "OpSequence", 0, 0, 900
    AddDetailField frm, "OpCode", 1000, 0, 1100
    AddDetailField frm, "ImportedHours", 2200, 0, 1200
    AddDetailField frm, "ImportedEx", 3500, 0, 1100
    AddDetailField frm, "BatchSize", 4700, 0, 900
    AddDetailField frm, "ExportHours", 5700, 0, 1200
    AddDetailField frm, "ExportEx", 7000, 0, 1100
    AddDetailField frm, "EquipmentType", 8200, 0, 1400
    AddDetailCheck frm, "UseExportHours", 9700, 0
    AddDetailCheck frm, "UseExportEx", 10400, 0
    AddDetailField frm, "ProcessHours", 11100, 0, 1200
    AddDetailField frm, "AvgEx", 12400, 0, 1000
    AddDetailField frm, "AvgHPU", 13500, 0, 1000
    AddDetailField frm, "MadeInFFA", 14600, 0, 1400
    SaveAndRenameForm frm, SFRM_OPS
End Sub

Private Sub CreatePartForm()
    Dim frm As Form
    Dim ctl As Control

    DeleteObjectIfExists acForm, FRM_PART
    Set frm = CreateForm()
    frm.RecordSource = TBL_PART
    frm.Caption = "Part"
    frm.DefaultView = 0
    frm.AllowAdditions = False

    AddDetailField frm, "BasePart", 1200, 200, 2000
    AddDetailCheck frm, "Active", 3400, 200
    AddDetailCombo frm, "HomeFFA", 1200, 600, 2000
    AddDetailField frm, "StatusDate", 1200, 1000, 1600

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

    SaveAndRenameForm frm, FRM_PART
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

    SaveAndRenameForm frm, FRM_REFERENCES
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

    SaveAndRenameForm frm, FRM_EXPORT
End Sub

Private Sub AddDetailField(ByVal frm As Form, ByVal fieldName As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthPos As Long)
    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , fieldName, leftPos, topPos, widthPos, 300)
    ctl.Name = "txt" & fieldName
    ctl.ControlSource = fieldName
End Sub

Private Sub AddDetailCheck(ByVal frm As Form, ByVal fieldName As String, ByVal leftPos As Long, ByVal topPos As Long)
    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acCheckBox, acDetail, , fieldName, leftPos, topPos, 300, 300)
    ctl.Name = "chk" & fieldName
    ctl.ControlSource = fieldName
End Sub

Private Sub AddDetailCombo(ByVal frm As Form, ByVal fieldName As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthPos As Long)
    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, , fieldName, leftPos, topPos, widthPos, 300)
    ctl.Name = "cbo" & fieldName
    ctl.ControlSource = fieldName
    ctl.RowSource = "SELECT FFA FROM [" & TBL_FFA & "] ORDER BY FFA"
    ctl.RowSourceType = "Table/Query"
    ctl.LimitToList = False
End Sub

Private Sub SaveAndRenameForm(ByRef frm As Form, ByVal desiredName As String)
    Dim savedName As String
    savedName = frm.Name
    DoCmd.Close acForm, savedName, acSaveYes
    Set frm = Nothing
    If StrComp(savedName, desiredName, vbTextCompare) <> 0 Then
        DoCmd.Rename desiredName, acForm, savedName
    End If
End Sub

Public Function OpenSelectedPart() As Boolean
    Dim basePart As String
    On Error GoTo Fail
    If Not CurrentProject.AllForms(FRM_HOME).IsLoaded Then
        OpenSelectedPart = False
        Exit Function
    End If
    basePart = CoerceText(Forms(FRM_HOME)!txtBasePart)
    If Len(basePart) = 0 Then
        basePart = CoerceText(Forms(FRM_HOME)!BasePart)
    End If
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
    basePart = CoerceText(Forms(FRM_PART)!txtBasePart)
    If Len(basePart) = 0 Then basePart = CoerceText(Forms(FRM_PART)!BasePart)
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
