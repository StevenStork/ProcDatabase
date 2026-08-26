Attribute VB_Name = "modUi"
Option Compare Database
Option Explicit

' Creates (or recreates) the Home / Part / reference / export forms.
' Safe to re-run: existing forms are deleted first.

Public UiSubStep As String

Public Sub EnsureUi()
    On Error GoTo Fail
    CloseProcDataForms
    UiSubStep = "CreateReferenceForms"
    CreateReferenceForms
    UiSubStep = "CreateEquipmentForm"
    CreateEquipmentForm
    UiSubStep = "CreateEquipmentFfaForm"
    CreateEquipmentFfaForm
    UiSubStep = "CreateEquipmentEntryForm"
    CreateEquipmentEntryForm
    UiSubStep = "CreatePartDashSubform"
    CreatePartDashSubform
    UiSubStep = "CreatePartProductLineSubform"
    CreatePartProductLineSubform
    UiSubStep = "CreateOperationSubform"
    CreateOperationSubform
    UiSubStep = "CreatePartForm"
    CreatePartForm
    UiSubStep = "CreateHomeListSubform"
    CreateHomeListSubform
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
    Dim frm As Form

    ' Datasheet view only shows bound controls — empty forms look blank.
    DeleteObjectIfExists acForm, FRM_FFA
    Set frm = CreateForm()
    frm.RecordSource = TBL_FFA
    frm.Caption = "FFAs"
    frm.DefaultView = 2
    frm.AllowAdditions = True
    frm.AllowDeletions = True
    frm.AllowEdits = True
    AddDetailField frm, "FFA", 0, 0, 1800
    AddDetailField frm, "Factory", 1900, 0, 3600
    SaveAndRenameForm frm, FRM_FFA

    DeleteObjectIfExists acForm, FRM_PRODUCT_LINE
    Set frm = CreateForm()
    frm.RecordSource = TBL_PRODUCT_LINE
    frm.Caption = "Product Lines"
    frm.DefaultView = 2
    frm.AllowAdditions = True
    frm.AllowDeletions = True
    frm.AllowEdits = True
    AddDetailField frm, "ProductLine", 0, 0, 3600
    AddDetailField frm, COL_PL_CODE, 3700, 0, 1800
    SaveAndRenameForm frm, FRM_PRODUCT_LINE
End Sub

Private Sub CreateEquipmentForm()
    Dim frm As Form

    ' Continuous form: fill a row, then use the blank row below for the next.
    DeleteObjectIfExists acForm, FRM_EQUIPMENT
    DeleteObjectIfExists acForm, "sfrmEquipmentFFA"
    Set frm = CreateForm()
    frm.RecordSource = TBL_EQUIPMENT
    frm.Caption = "Equipment — Tab or ↓ for a new row"
    frm.DefaultView = 1
    frm.AllowAdditions = True
    frm.AllowDeletions = True
    frm.AllowEdits = True
    frm.RecordSelectors = True
    frm.NavigationButtons = True
    frm.ScrollBars = 2
    On Error Resume Next
    frm.Section(acDetail).Height = 360
    On Error GoTo 0

    AddDetailField frm, "Equipment", 120, 30, 3600
    AddDetailField frm, COL_EQUIP_TYPE, 3840, 30, 2400

    SaveAndRenameForm frm, FRM_EQUIPMENT
End Sub

Private Sub CreateEquipmentFfaForm()
    Dim frm As Form
    Dim ctl As Control

    ' Continuous form: each row is one Equipment↔FFA link; blank row at bottom for next.
    DeleteObjectIfExists acForm, FRM_EQUIPMENT_FFA
    Set frm = CreateForm()
    frm.RecordSource = TBL_EQUIPMENT_FFA
    frm.Caption = "Equipment FFAs — Tab or ↓ for a new row"
    frm.DefaultView = 1
    frm.AllowAdditions = True
    frm.AllowDeletions = True
    frm.AllowEdits = True
    frm.RecordSelectors = True
    frm.NavigationButtons = True
    frm.ScrollBars = 2
    On Error Resume Next
    frm.Section(acDetail).Height = 360
    On Error GoTo 0

    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, , "Equipment", 120, 30, 3600, 300)
    ctl.Name = "cboEquipment"
    ctl.ControlSource = "Equipment"
    ctl.RowSource = "SELECT Equipment FROM [" & TBL_EQUIPMENT & "] ORDER BY Equipment"
    ctl.RowSourceType = "Table/Query"
    ctl.LimitToList = True

    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, , "FFA", 3840, 30, 2400, 300)
    ctl.Name = "cboFFA"
    ctl.ControlSource = "FFA"
    ctl.RowSource = "SELECT FFA FROM [" & TBL_FFA & "] ORDER BY FFA"
    ctl.RowSourceType = "Table/Query"
    ctl.LimitToList = True

    SaveAndRenameForm frm, FRM_EQUIPMENT_FFA
End Sub

Private Sub CreateEquipmentEntryForm()
    Dim frm As Form
    Dim ctl As Control
    Dim lbl As Control

    ' Unbound entry form: fill fields → Add → clears for the next piece.
    DeleteObjectIfExists acForm, FRM_EQUIPMENT_ENTRY
    Set frm = CreateForm()
    frm.RecordSource = vbNullString
    frm.Caption = "Add Equipment"
    frm.DefaultView = 0
    frm.AllowAdditions = False
    frm.AllowDeletions = False
    frm.AllowEdits = True
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.ScrollBars = 0

    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, 200, 2000, 300)
    lbl.Name = "lblEquipment"
    lbl.Caption = "Equipment name"

    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , 2200, 200, 3600, 300)
    ctl.Name = "txtEquipment"
    ctl.Enabled = True
    ctl.Locked = False

    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, 600, 2000, 300)
    lbl.Name = "lblEquipType"
    lbl.Caption = "Equipment type"

    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , 2200, 600, 3600, 300)
    ctl.Name = "txtEquipmentType"
    ctl.Enabled = True
    ctl.Locked = False

    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, 1000, 3600, 300)
    lbl.Name = "lblFfas"
    lbl.Caption = "FFAs where this equipment exists (click to select)"

    Set ctl = CreateControl(frm.Name, acListBox, acDetail, , , 200, 1300, 5600, 2400)
    ctl.Name = "lstFfas"
    ctl.RowSource = "SELECT FFA FROM [" & TBL_FFA & "] ORDER BY FFA"
    ctl.RowSourceType = "Table/Query"
    ctl.MultiSelect = 1

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 200, 3900, 1800, 400)
    ctl.Name = "btnAdd"
    ctl.Caption = "Add Equipment"
    ctl.OnClick = "=UiAddEquipmentEntry()"

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 2200, 3900, 1400, 400)
    ctl.Name = "btnClear"
    ctl.Caption = "Clear"
    ctl.OnClick = "=UiClearEquipmentEntry()"

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 3800, 3900, 1400, 400)
    ctl.Name = "btnClose"
    ctl.Caption = "Close"
    ctl.OnClick = "=UiCloseCurrentForm()"

    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, 4400, 5600, 300)
    lbl.Name = "lblStatus"
    lbl.Caption = "Enter a name, optional type, select FFAs, then Add."

    SaveAndRenameForm frm, FRM_EQUIPMENT_ENTRY
End Sub

Private Sub CreateHomeListSubform()
    Dim frm As Form

    DeleteObjectIfExists acForm, SFRM_HOME_LIST
    Set frm = CreateForm()
    frm.RecordSource = QRY_HOME
    frm.DefaultView = 1
    frm.AllowAdditions = False
    frm.AllowDeletions = False
    frm.AllowEdits = True
    frm.ScrollBars = 2
    frm.RecordSelectors = True
    frm.NavigationButtons = True

    AddHomeField frm, "BasePart", 120, 0, 1800
    AddHomeField frm, "Active", 2040, 0, 900, True
    AddHomeField frm, "StatusDate", 3000, 0, 1200
    AddHomeField frm, "Days", 4320, 0, 720
    frm.Controls("txtDays").Enabled = False
    AddHomeField frm, COL_NOTES, 5100, 0, 2400
    AddHomeCombo frm, "HomeFFA", 7560, 0, 1440
    AddHomeField frm, "Factories", 9060, 0, 1800
    frm.Controls("txtFactories").Enabled = False

    SaveAndRenameForm frm, SFRM_HOME_LIST
    ApplyHomeDaysRagDesign SFRM_HOME_LIST
End Sub

Private Sub CreateHomeForm()
    Dim frm As Form
    Dim ctl As Control
    Dim lbl As Control

    DeleteObjectIfExists acForm, FRM_HOME
    Set frm = CreateForm()
    frm.RecordSource = vbNullString
    frm.Caption = "ProcDatabase Home"
    frm.DefaultView = 0
    frm.AllowAdditions = False
    frm.AllowDeletions = False
    frm.AllowEdits = True
    frm.ScrollBars = 2
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.OnLoad = "=HomeForm_Load()"

    ' Row 1 — actions
    AddToolbarButton frm, "btnRefresh", "Refresh Linked Data", 120, 120, "=UiRefreshLinkedData()", 2160
    AddToolbarButton frm, "btnSeedOps", "Seed Active Ops", 2400, 120, "=UiSeedActiveOps()", 1800
    AddToolbarButton frm, "btnReferences", "References", 4320, 120, "=UiOpenReferences()", 1440
    AddToolbarButton frm, "btnExport", "Export", 5880, 120, "=UiOpenExport()", 1200

    ' Row 2 — find / filter (avoid scrolling an infinite list)
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 120, 560, 900, 300)
    lbl.Name = "lblSearch"
    lbl.Caption = "Search"

    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , 1000, 540, 2400, 360)
    ctl.Name = "txtSearch"
    ctl.OnChange = "=HomeSearchChanged()"
    ctl.OnExit = "=HomeApplyFilter()"

    Set ctl = CreateControl(frm.Name, acCheckBox, acDetail, , , 3600, 560, 300, 300)
    ctl.Name = "chkActiveOnly"
    ctl.DefaultValue = "True"
    ctl.AfterUpdate = "=HomeApplyFilter()"

    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 3960, 560, 1200, 300)
    lbl.Name = "lblActiveOnly"
    lbl.Caption = "Active only"

    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 5400, 560, 600, 300)
    lbl.Name = "lblFilterFfa"
    lbl.Caption = "FFA"

    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, , , 6000, 540, 1800, 360)
    ctl.Name = "cboFilterFfa"
    ctl.RowSource = "SELECT FFA FROM [" & TBL_FFA & "] ORDER BY FFA"
    ctl.RowSourceType = "Table/Query"
    ctl.LimitToList = False
    ctl.AfterUpdate = "=HomeApplyFilter()"

    AddToolbarButton frm, "btnClearFilter", "Clear filters", 8000, 540, "=HomeClearFilter()", 1600

    ' Row 3 — jump to part
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 120, 1020, 1200, 300)
    lbl.Name = "lblJump"
    lbl.Caption = "Jump to part"

    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, , , 1400, 1000, 3600, 360)
    ctl.Name = "cboJumpPart"
    ctl.RowSource = "SELECT BasePart FROM [" & QRY_HOME & "] ORDER BY BasePart"
    ctl.RowSourceType = "Table/Query"
    ctl.LimitToList = False
    ctl.AfterUpdate = "=HomeJumpToPart()"

    AddToolbarButton frm, "btnJump", "Go", 5200, 1000, "=HomeJumpToPart()", 900
    AddToolbarButton frm, "btnOpenPart", "Open Part", 6240, 1000, "=HomeOpenSelectedPart()", 1440

    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 120, 1440, 10000, 300)
    lbl.Name = "lblStatus"
    lbl.Caption = "Use Search / Active only / FFA / Jump to find parts without scrolling."

    Set ctl = CreateControl(frm.Name, acSubform, acDetail, , , 120, 1800, 11040, 4200)
    ctl.Name = "subParts"
    ctl.SourceObject = SFRM_HOME_LIST

    SaveAndRenameForm frm, FRM_HOME
    EnsureHomeFormClassModule
End Sub

' Best-effort: install Form_frmHome class module when VBA project access is trusted.
Public Sub EnsureHomeFormClassModule()
    Dim comp As Object
    Dim codeMod As Object
    Dim src As String
    On Error GoTo Fail

    DoCmd.OpenForm FRM_HOME, acDesign
    On Error Resume Next
    Forms(FRM_HOME).HasModule = True
    On Error GoTo Fail
    DoCmd.Close acForm, FRM_HOME, acSaveYes

    Set comp = Application.VBE.ActiveVBProject.VBComponents("Form_" & FRM_HOME)
    Set codeMod = comp.CodeModule
    src = HomeFormClassModuleSource()
    If codeMod.CountOfLines > 0 Then
        codeMod.DeleteLines 1, codeMod.CountOfLines
    End If
    codeMod.AddFromString src
    Exit Sub
Fail:
    ' Trust Center may block VBIDE — OnLoad/=HomeForm_Load() still works.
    On Error Resume Next
    DoCmd.Close acForm, FRM_HOME, acSaveYes
    On Error GoTo 0
End Sub

Private Function HomeFormClassModuleSource() As String
    Dim s As String
    s = "Option Compare Database" & vbCrLf
    s = s & "Option Explicit" & vbCrLf & vbCrLf
    s = s & "' Form class module — forwards to modHome controller." & vbCrLf
    s = s & "Private Sub Form_Load()" & vbCrLf
    s = s & "    HomeForm_Load" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub txtSearch_Change()" & vbCrLf
    s = s & "    HomeSearchChanged" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub txtSearch_Exit(Cancel As Integer)" & vbCrLf
    s = s & "    HomeApplyFilter" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub chkActiveOnly_AfterUpdate()" & vbCrLf
    s = s & "    HomeApplyFilter" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub cboFilterFfa_AfterUpdate()" & vbCrLf
    s = s & "    HomeApplyFilter" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub cboJumpPart_AfterUpdate()" & vbCrLf
    s = s & "    HomeJumpToPart" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub btnJump_Click()" & vbCrLf
    s = s & "    HomeJumpToPart" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub btnOpenPart_Click()" & vbCrLf
    s = s & "    HomeOpenSelectedPart" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub btnClearFilter_Click()" & vbCrLf
    s = s & "    HomeClearFilter" & vbCrLf
    s = s & "End Sub" & vbCrLf
    HomeFormClassModuleSource = s
End Function

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

Private Sub AddToolbarButton(ByVal frm As Form, ByVal ctlName As String, ByVal caption As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal onClick As String, ByVal widthPos As Long)
    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , leftPos, topPos, widthPos, 360)
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
    ' RAG colors are optional; ignore design-time failures.
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
    Dim ctl As Control
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

    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, , "EquipmentType", 8200, 0, 1800, 300)
    ctl.Name = "cboEquipmentType"
    ctl.ControlSource = "EquipmentType"
    ctl.RowSource = "SELECT Equipment FROM [" & TBL_EQUIPMENT & "] ORDER BY Equipment"
    ctl.RowSourceType = "Table/Query"
    ctl.LimitToList = False

    AddDetailCheck frm, "UseExportHours", 10100, 0
    AddDetailCheck frm, "UseExportEx", 10800, 0
    AddDetailField frm, "ProcessHours", 11500, 0, 1200
    AddDetailField frm, "AvgEx", 12800, 0, 1000
    AddDetailField frm, "AvgHPU", 13900, 0, 1000
    AddDetailField frm, "MadeInFFA", 15000, 0, 1400
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

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 200, 1200, 2400, 400)
    ctl.Name = "btnEquipAdd"
    ctl.Caption = "Add Equipment"
    ctl.OnClick = "=UiOpenEquipmentEntryForm()"

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 200, 1700, 2400, 400)
    ctl.Name = "btnEquip"
    ctl.Caption = "View Equipment List"
    ctl.OnClick = "=UiOpenEquipmentForm()"

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 200, 2200, 2400, 400)
    ctl.Name = "btnEquipFfa"
    ctl.Caption = "View Equipment FFAs"
    ctl.OnClick = "=UiOpenEquipmentFfaForm()"

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
    OpenSelectedPart = HomeOpenSelectedPart()
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
    HomeRefreshAfterDataChange
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

Public Function UiOpenEquipmentEntryForm() As Boolean
    On Error GoTo Fail
    DoCmd.OpenForm FRM_EQUIPMENT_ENTRY
    UiClearEquipmentEntry
    UiOpenEquipmentEntryForm = True
    Exit Function
Fail:
    UiOpenEquipmentEntryForm = False
End Function

Public Function UiOpenEquipmentFfaForm() As Boolean
    On Error GoTo Fail
    DoCmd.OpenForm FRM_EQUIPMENT_FFA
    UiOpenEquipmentFfaForm = True
    Exit Function
Fail:
    UiOpenEquipmentFfaForm = False
End Function

Public Function UiAddEquipmentEntry() As Boolean
    Dim frm As Form
    Dim lst As ListBox
    Dim equipName As String
    Dim equipType As String
    Dim ffaValue As String
    Dim i As Long
    Dim addedFfas As Long
    Dim db As DAO.Database

    On Error GoTo Fail
    If Not CurrentProject.AllForms(FRM_EQUIPMENT_ENTRY).IsLoaded Then
        UiAddEquipmentEntry = False
        Exit Function
    End If

    Set frm = Forms(FRM_EQUIPMENT_ENTRY)
    equipName = Trim$(CoerceText(frm!txtEquipment))
    equipType = Trim$(CoerceText(frm!txtEquipmentType))
    If Len(equipName) = 0 Then
        frm!lblStatus.Caption = "Enter an equipment name first."
        MsgBox "Enter an equipment name.", vbExclamation, "Add Equipment"
        UiAddEquipmentEntry = False
        Exit Function
    End If

    Set db = CurrentDb
    If IsNull(DLookup("Equipment", TBL_EQUIPMENT, "Equipment = " & SqlText(equipName))) Then
        db.Execute "INSERT INTO [" & TBL_EQUIPMENT & "] (Equipment, [" & COL_EQUIP_TYPE & "]) VALUES (" & _
            SqlText(equipName) & ", " & SqlNullableText(equipType) & ")", dbFailOnError
    ElseIf Len(equipType) > 0 Then
        db.Execute "UPDATE [" & TBL_EQUIPMENT & "] SET [" & COL_EQUIP_TYPE & "] = " & SqlText(equipType) & _
            " WHERE Equipment = " & SqlText(equipName), dbFailOnError
    End If

    addedFfas = 0
    Set lst = frm!lstFfas
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            ffaValue = CoerceText(lst.ItemData(i))
            If Len(ffaValue) > 0 Then
                If IsNull(DLookup("FFA", TBL_EQUIPMENT_FFA, _
                    "Equipment = " & SqlText(equipName) & " AND FFA = " & SqlText(ffaValue))) Then
                    db.Execute "INSERT INTO [" & TBL_EQUIPMENT_FFA & "] (Equipment, FFA) VALUES (" & _
                        SqlText(equipName) & ", " & SqlText(ffaValue) & ")", dbFailOnError
                    addedFfas = addedFfas + 1
                End If
            End If
        End If
    Next i

    UiClearEquipmentEntry
    frm!lblStatus.Caption = "Added '" & equipName & "' (" & CStr(addedFfas) & " new FFA link(s)). Ready for next."
    On Error Resume Next
    frm!txtEquipment.SetFocus
    On Error GoTo 0
    UiAddEquipmentEntry = True
    Exit Function
Fail:
    On Error Resume Next
    If CurrentProject.AllForms(FRM_EQUIPMENT_ENTRY).IsLoaded Then
        Forms(FRM_EQUIPMENT_ENTRY)!lblStatus.Caption = "Add failed: " & Err.Description
    End If
    MsgBox "Could not add equipment: " & Err.Description, vbCritical, "Add Equipment"
    UiAddEquipmentEntry = False
End Function

Public Function UiClearEquipmentEntry() As Boolean
    Dim frm As Form
    Dim lst As ListBox
    Dim i As Long
    On Error GoTo Fail
    If Not CurrentProject.AllForms(FRM_EQUIPMENT_ENTRY).IsLoaded Then
        UiClearEquipmentEntry = False
        Exit Function
    End If
    Set frm = Forms(FRM_EQUIPMENT_ENTRY)
    frm!txtEquipment = Null
    frm!txtEquipmentType = Null
    Set lst = frm!lstFfas
    For i = 0 To lst.ListCount - 1
        lst.Selected(i) = False
    Next i
    UiClearEquipmentEntry = True
    Exit Function
Fail:
    UiClearEquipmentEntry = False
End Function

Private Function SqlNullableText(ByVal value As String) As String
    If Len(Trim$(value)) = 0 Then
        SqlNullableText = "Null"
    Else
        SqlNullableText = SqlText(value)
    End If
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
