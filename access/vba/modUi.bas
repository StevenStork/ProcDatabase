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

    If Not TableExists(TBL_FFA) Then
        Err.Raise vbObjectError + 1, "CreateReferenceForms", _
            "Table " & TBL_FFA & " is missing. Run BootstrapProcDatabase / EnsureSchema first."
    End If
    If Not TableExists(TBL_PRODUCT_LINE) Then
        Err.Raise vbObjectError + 1, "CreateReferenceForms", _
            "Table " & TBL_PRODUCT_LINE & " is missing. Run BootstrapProcDatabase / EnsureSchema first."
    End If

    ' Continuous single-line rows (same pattern as Equipment) — CreateForm defaults
    ' leave a tall Detail section so one record fills the window.
    Set frm = CreateForm()
    frm.RecordSource = TBL_FFA
    frm.Caption = "FFAs — Tab or ↓ for a new row"
    frm.DefaultView = 1
    frm.AllowAdditions = True
    frm.AllowDeletions = True
    frm.AllowEdits = True
    frm.RecordSelectors = True
    frm.NavigationButtons = True
    frm.ScrollBars = 2
    AddDetailField frm, "FFA", 120, 30, 1800, False
    AddDetailField frm, "Factory", 2040, 30, 3600, False
    SetCompactDetailHeight frm
    AddContinuousColumnHeaders frm, Array( _
        Array("FFA", 120&, 1800&), _
        Array("Factory", 2040&, 3600&))
    SaveAndRenameForm frm, FRM_FFA

    Set frm = CreateForm()
    frm.RecordSource = TBL_PRODUCT_LINE
    frm.Caption = "Product Lines — Tab or ↓ for a new row"
    frm.DefaultView = 1
    frm.AllowAdditions = True
    frm.AllowDeletions = True
    frm.AllowEdits = True
    frm.RecordSelectors = True
    frm.NavigationButtons = True
    frm.ScrollBars = 2
    AddDetailField frm, "ProductLine", 120, 30, 3600, False
    AddDetailField frm, COL_PL_CODE, 3840, 30, 1800, False
    SetCompactDetailHeight frm
    AddContinuousColumnHeaders frm, Array( _
        Array("Product Line", 120&, 3600&), _
        Array("PL Code", 3840&, 1800&))
    SaveAndRenameForm frm, FRM_PRODUCT_LINE
End Sub

Private Sub CreateEquipmentForm()
    Dim frm As Form

    ' Continuous form: fill a row, then use the blank row below for the next.
    ' Remove legacy subform (older UI embedded Equipment↔FFA here). Current UI uses
    ' frmEquipmentFFA / frmEquipmentEntry instead — do not recreate sfrmEquipmentFFA.
    DeleteObjectIfExists acForm, LEGACY_SFRM_EQUIPMENT_FFA

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

    AddDetailField frm, "Equipment", 120, 30, 3600, False
    AddDetailField frm, COL_EQUIP_TYPE, 3840, 30, 2400, False
    SetCompactDetailHeight frm
    AddContinuousColumnHeaders frm, Array( _
        Array("Equipment", 120&, 3600&), _
        Array("Equipment Type", 3840&, 2400&))

    SaveAndRenameForm frm, FRM_EQUIPMENT
End Sub

Private Sub CreateEquipmentFfaForm()
    Dim frm As Form
    Dim ctl As Control

    If Not TableExists(TBL_EQUIPMENT_FFA) Then
        Err.Raise vbObjectError + 1, "CreateEquipmentFfaForm", _
            "Table " & TBL_EQUIPMENT_FFA & " is missing. Run BootstrapProcDatabase / EnsureSchema first."
    End If

    ' Continuous form: each row is one Equipment↔FFA link; blank row at bottom for next.
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
    SetCompactDetailHeight frm
    AddContinuousColumnHeaders frm, Array( _
        Array("Equipment", 120&, 3600&), _
        Array("FFA", 3840&, 2400&))

    SaveAndRenameForm frm, FRM_EQUIPMENT_FFA
End Sub

Private Sub CreateEquipmentEntryForm()
    Dim frm As Form
    Dim ctl As Control
    Dim lbl As Control

    ' Unbound entry form: fill fields → Add → clears for the next piece.
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
    ' Leave room below for Add/Clear/Close + status (~900 twips).
    SizeFillControl ctl, 200, 1300, 240, 1000
    AnchorStretch ctl, True, True

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 200, LayoutUsableHeight() - 900, 1800, 400)
    ctl.Name = "btnAdd"
    ctl.Caption = "Add Equipment"
    ctl.OnClick = "=UiAddEquipmentEntry()"
    AnchorBottom ctl

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 2200, LayoutUsableHeight() - 900, 1400, 400)
    ctl.Name = "btnClear"
    ctl.Caption = "Clear"
    ctl.OnClick = "=UiClearEquipmentEntry()"
    AnchorBottom ctl

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 3800, LayoutUsableHeight() - 900, 1400, 400)
    ctl.Name = "btnClose"
    ctl.Caption = "Close"
    ctl.OnClick = "=UiCloseForm(""" & FRM_EQUIPMENT_ENTRY & """)"
    AnchorBottom ctl

    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, LayoutUsableHeight() - 480, LayoutUsableWidth() - 440, 300)
    lbl.Name = "lblStatus"
    lbl.Caption = "Enter a name, optional type, select FFAs, then Add."
    AnchorBottom lbl, True

    SaveAndRenameForm frm, FRM_EQUIPMENT_ENTRY
End Sub

Private Sub CreateHomeListSubform()
    Dim frm As Form

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
    SetCompactDetailHeight frm
    AddContinuousColumnHeaders frm, Array( _
        Array("Base Part", 120&, 1800&), _
        Array("Active", 2040&, 900&), _
        Array("Status Date", 3000&, 1200&), _
        Array("Days", 4320&, 720&), _
        Array("Notes", 5100&, 2400&), _
        Array("Home FFA", 7560&, 1440&), _
        Array("Factories", 9060&, 1800&))

    SaveAndRenameForm frm, SFRM_HOME_LIST
    ApplyHomeDaysRagDesign SFRM_HOME_LIST
End Sub

Private Sub CreateHomeForm()
    Dim frm As Form
    Dim ctl As Control
    Dim lbl As Control

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

    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 120, 1440, LayoutUsableWidth() - 360, 300)
    lbl.Name = "lblStatus"
    lbl.Caption = "Use Search / Active only / FFA / Jump to find parts without scrolling."
    AnchorStretch lbl, True, False

    Set ctl = CreateControl(frm.Name, acSubform, acDetail, , , 120, 1800, 11040, 4200)
    ctl.Name = "subParts"
    ctl.SourceObject = SFRM_HOME_LIST
    SizeFillControl ctl, 120, 1800, 240, 240
    AnchorStretch ctl, True, True

    SaveAndRenameForm frm, FRM_HOME
End Sub

Private Sub AddHomeField(ByVal frm As Form, ByVal fieldName As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthPos As Long, Optional ByVal isCheckBox As Boolean = False)
    Dim ctl As Control
    If isCheckBox Then
        Set ctl = CreateControl(frm.Name, acCheckBox, acDetail, , fieldName, leftPos, topPos, 300, 300)
        ctl.Name = "chk" & ControlBaseName(fieldName)
    Else
        Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , fieldName, leftPos, topPos, widthPos, 300)
        ctl.Name = "txt" & ControlBaseName(fieldName)
    End If
    ctl.ControlSource = fieldName
End Sub

Private Sub AddHomeCombo(ByVal frm As Form, ByVal fieldName As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthPos As Long)
    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, , fieldName, leftPos, topPos, widthPos, 300)
    ctl.Name = "cbo" & ControlBaseName(fieldName)
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
    AttachFieldLabel frm, ctl, FriendlyFieldCaption("EquipmentType"), 8200, 0, 1800

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
    Dim contentW As Long
    Dim contentH As Long
    Dim halfW As Long
    Dim midH As Long
    Dim opsTop As Long
    Dim opsH As Long

    Set frm = CreateForm()
    frm.RecordSource = TBL_PART
    frm.Caption = "Part"
    frm.DefaultView = 0
    frm.AllowAdditions = False

    contentW = LayoutUsableWidth() - 400
    contentH = LayoutUsableHeight() - 400
    If contentW < 10000 Then contentW = 10000
    If contentH < 7000 Then contentH = 7000
    halfW = (contentW - 200) \ 2
    midH = CLng(contentH * 0.32)
    If midH < 2200 Then midH = 2200
    opsTop = 1600 + midH + 200
    opsH = contentH - opsTop - 200
    If opsH < 2800 Then opsH = 2800

    AddStandaloneLabel frm, "Base Part", 200, 200, 900
    AddDetailField frm, "BasePart", 1200, 200, 2000, False
    AddDetailCheck frm, "Active", 3400, 200, False
    AddStandaloneLabel frm, "Active", 3700, 200, 800
    AddStandaloneLabel frm, "Home FFA", 200, 600, 900
    AddDetailCombo frm, "HomeFFA", 1200, 600, 2000, False
    AddStandaloneLabel frm, "Status Date", 200, 1000, 900
    AddDetailField frm, "StatusDate", 1200, 1000, 1600, False

    Set ctl = CreateControl(frm.Name, acSubform, acDetail, , , 200, 1600, halfW, midH)
    ctl.Name = "subDashes"
    ctl.SourceObject = SFRM_DASH
    ctl.LinkMasterFields = "BasePart"
    ctl.LinkChildFields = "BasePart"
    AnchorStretch ctl, True, False

    Set ctl = CreateControl(frm.Name, acSubform, acDetail, , , 200 + halfW + 200, 1600, halfW, midH)
    ctl.Name = "subProductLines"
    ctl.SourceObject = SFRM_PL
    ctl.LinkMasterFields = "BasePart"
    ctl.LinkChildFields = "BasePart"
    AnchorStretch ctl, True, False

    Set ctl = CreateControl(frm.Name, acSubform, acDetail, , , 200, opsTop, contentW, opsH)
    ctl.Name = "subOperations"
    ctl.SourceObject = SFRM_OPS
    ctl.LinkMasterFields = "BasePart"
    ctl.LinkChildFields = "BasePart"
    AnchorStretch ctl, True, True

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , LayoutUsableWidth() - 1840, 200, 1600, 360)
    ctl.Name = "btnSeed"
    ctl.Caption = "Seed Ops"
    ctl.OnClick = "=SeedOperationsForCurrentPart()"
    AnchorTopRight ctl

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , LayoutUsableWidth() - 1840, 600, 1600, 360)
    ctl.Name = "btnClose"
    ctl.Caption = "Close"
    ctl.OnClick = "=UiCloseForm(""" & FRM_PART & """)"
    AnchorTopRight ctl

    SaveAndRenameForm frm, FRM_PART
End Sub

Private Sub CreateReferencesForm()
    Dim frm As Form
    Dim ctl As Control

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

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 200, 2800, 1400, 400)
    ctl.Name = "btnClose"
    ctl.Caption = "Close"
    ctl.OnClick = "=UiCloseForm(""" & FRM_REFERENCES & """)"

    SaveAndRenameForm frm, FRM_REFERENCES
End Sub

Private Sub CreateExportForm()
    Dim frm As Form
    Dim ctl As Control

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

    Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 200, 1900, 1400, 400)
    ctl.Name = "btnClose"
    ctl.Caption = "Close"
    ctl.OnClick = "=UiCloseForm(""" & FRM_EXPORT & """)"

    SaveAndRenameForm frm, FRM_EXPORT
End Sub

Private Function ControlBaseName(ByVal fieldName As String) As String
    ControlBaseName = Replace(Replace(Replace(fieldName, " ", vbNullString), ":", vbNullString), "/", vbNullString)
End Function

' Datasheet column headers use the attached label Caption — not txtOpSequence.
Private Function FriendlyFieldCaption(ByVal fieldName As String) As String
    Select Case fieldName
        Case "OpSequence": FriendlyFieldCaption = "Op Sequence"
        Case "OpCode": FriendlyFieldCaption = "Op Code"
        Case "ImportedHours": FriendlyFieldCaption = "Imported Hours"
        Case "ImportedEx": FriendlyFieldCaption = "Imported Ex"
        Case "BatchSize": FriendlyFieldCaption = "Batch Size"
        Case "ExportHours": FriendlyFieldCaption = "Export Hours"
        Case "ExportEx": FriendlyFieldCaption = "Export Ex"
        Case "EquipmentType": FriendlyFieldCaption = "Equipment Type"
        Case "UseExportHours": FriendlyFieldCaption = "Use Exp Hrs"
        Case "UseExportEx": FriendlyFieldCaption = "Use Exp Ex"
        Case "ProcessHours": FriendlyFieldCaption = "Process Hours"
        Case "AvgEx": FriendlyFieldCaption = "Avg Ex"
        Case "AvgHPU": FriendlyFieldCaption = "Avg HPU"
        Case "MadeInFFA": FriendlyFieldCaption = "Made In FFA"
        Case "BasePart": FriendlyFieldCaption = "Base Part"
        Case "StatusDate": FriendlyFieldCaption = "Status Date"
        Case "HomeFFA": FriendlyFieldCaption = "Home FFA"
        Case "ProductLine": FriendlyFieldCaption = "Product Line"
        Case "UseFlag": FriendlyFieldCaption = "Use"
        Case "Active": FriendlyFieldCaption = "Active"
        Case "Dash": FriendlyFieldCaption = "Dash"
        Case "Days": FriendlyFieldCaption = "Days"
        Case "Factories": FriendlyFieldCaption = "Factories"
        Case "FFA": FriendlyFieldCaption = "FFA"
        Case "Factory": FriendlyFieldCaption = "Factory"
        Case "Equipment": FriendlyFieldCaption = "Equipment"
        Case COL_NOTES: FriendlyFieldCaption = "Notes"
        Case COL_PL_CODE: FriendlyFieldCaption = "PL Code"
        Case COL_EQUIP_TYPE: FriendlyFieldCaption = "Equipment Type"
        Case Else: FriendlyFieldCaption = fieldName
    End Select
End Function

' Attached label Caption becomes the datasheet column heading.
Private Sub AttachFieldLabel(ByVal frm As Form, ByVal ctl As Control, ByVal caption As String, _
    ByVal leftPos As Long, ByVal topPos As Long, ByVal widthPos As Long)

    Dim lbl As Control
    On Error Resume Next
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, ctl.Name, , leftPos, topPos, widthPos, 240)
    If Not lbl Is Nothing Then
        lbl.Name = "lbl" & ControlBaseName(ctl.Name)
        lbl.Caption = caption
        lbl.Width = widthPos
    End If
    On Error GoTo 0
End Sub

Private Sub AddStandaloneLabel(ByVal frm As Form, ByVal caption As String, _
    ByVal leftPos As Long, ByVal topPos As Long, ByVal widthPos As Long)

    Dim lbl As Control
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , leftPos, topPos, widthPos, 300)
    lbl.Name = "lbl" & ControlBaseName(caption) & CStr(leftPos)
    lbl.Caption = caption
End Sub

' CreateForm() leaves Header/Footer off. Toggle them on before referencing acHeader
' (otherwise Section/CreateControl raises 2148).
Private Function EnsureFormHeader(ByVal frm As Form, Optional ByVal headerHeight As Long = 360) As Boolean
    Dim probe As Long
    On Error Resume Next
    probe = frm.Section(acHeader).Height
    If Err.Number <> 0 Then
        Err.Clear
        ' Form is already open in Design from CreateForm — toggle Header/Footer on.
        DoCmd.SelectObject acForm, frm.Name, False
        DoCmd.RunCommand acCmdFormHdrFtr
    End If
    Err.Clear
    frm.Section(acHeader).Height = headerHeight
    EnsureFormHeader = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' Column titles once in the Form Header for continuous (multi-row) forms.
Private Sub AddContinuousColumnHeaders(ByVal frm As Form, ByVal headers As Variant)
    Dim i As Long
    Dim lbl As Control
    Dim caption As String
    Dim leftPos As Long
    Dim widthPos As Long

    On Error GoTo FailSoft
    If Not EnsureFormHeader(frm) Then Exit Sub

    For i = LBound(headers) To UBound(headers)
        caption = CStr(headers(i)(0))
        leftPos = CLng(headers(i)(1))
        widthPos = CLng(headers(i)(2))
        Set lbl = CreateControl(frm.Name, acLabel, acHeader, , , leftPos, 40, widthPos, 280)
        lbl.Name = "hdr" & ControlBaseName(caption) & CStr(i)
        lbl.Caption = caption
        lbl.FontBold = True
    Next i
    Exit Sub
FailSoft:
    ' Headers are optional — do not fail BuildUi (2148 if Header still missing).
End Sub

Private Sub AddDetailField(ByVal frm As Form, ByVal fieldName As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthPos As Long, _
    Optional ByVal attachLabel As Boolean = True)

    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , fieldName, leftPos, topPos, widthPos, 300)
    ctl.Name = "txt" & ControlBaseName(fieldName)
    ctl.ControlSource = fieldName
    If attachLabel Then
        AttachFieldLabel frm, ctl, FriendlyFieldCaption(fieldName), leftPos, topPos, widthPos
    End If
End Sub

Private Sub AddDetailCheck(ByVal frm As Form, ByVal fieldName As String, ByVal leftPos As Long, ByVal topPos As Long, _
    Optional ByVal attachLabel As Boolean = True)

    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acCheckBox, acDetail, , fieldName, leftPos, topPos, 300, 300)
    ctl.Name = "chk" & ControlBaseName(fieldName)
    ctl.ControlSource = fieldName
    If attachLabel Then
        AttachFieldLabel frm, ctl, FriendlyFieldCaption(fieldName), leftPos, topPos, 1200
    End If
End Sub

Private Sub AddDetailCombo(ByVal frm As Form, ByVal fieldName As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthPos As Long, _
    Optional ByVal attachLabel As Boolean = True)

    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acComboBox, acDetail, , fieldName, leftPos, topPos, widthPos, 300)
    ctl.Name = "cbo" & ControlBaseName(fieldName)
    ctl.ControlSource = fieldName
    ctl.RowSource = "SELECT FFA FROM [" & TBL_FFA & "] ORDER BY FFA"
    ctl.RowSourceType = "Table/Query"
    ctl.LimitToList = False
    If attachLabel Then
        AttachFieldLabel frm, ctl, FriendlyFieldCaption(fieldName), leftPos, topPos, widthPos
    End If
End Sub

Private Sub SaveAndRenameForm(ByRef frm As Form, ByVal desiredName As String)
    Dim savedName As String
    Dim errNum As Long
    Dim errDesc As String

    On Error GoTo Fail

    ' Single-form shells: grow detail to fill the Access workspace.
    ' Continuous/datasheet forms: wide window only — keep row height intact.
    Select Case desiredName
        Case FRM_HOME, FRM_PART, FRM_EQUIPMENT_ENTRY, FRM_REFERENCES, FRM_EXPORT
            ApplyLargeFormLayout frm, 0.96, 0.92, True
        Case FRM_EQUIPMENT, FRM_EQUIPMENT_FFA, FRM_FFA, FRM_PRODUCT_LINE, SFRM_HOME_LIST
            ApplyLargeFormLayout frm, 0.96, 0.92, False
            SetCompactDetailHeight frm
        Case SFRM_OPS, SFRM_DASH, SFRM_PL
            ApplyLargeFormLayout frm, 0.9, 0.5, False
    End Select

    savedName = frm.Name
    DoCmd.Close acForm, savedName, acSaveYes
    Set frm = Nothing
    DoEvents

    If StrComp(savedName, desiredName, vbTextCompare) <> 0 Then
        ' Replace any prior copy only after the new form saved successfully.
        DeleteObjectIfExists acForm, desiredName
        DoEvents
        DoCmd.Rename desiredName, acForm, savedName
    End If

    If Not ObjectExists(acForm, desiredName) Then
        Err.Raise vbObjectError + 1, "SaveAndRenameForm", _
            "Form '" & desiredName & "' was not created (temp name '" & savedName & "')."
    End If
    Exit Sub

Fail:
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    If Not frm Is Nothing Then
        DoCmd.Close acForm, frm.Name, acSaveNo
        Set frm = Nothing
    End If
    On Error GoTo 0
    Err.Raise errNum, "SaveAndRenameForm." & desiredName, errDesc
End Sub

' One-line continuous-form rows (CreateForm defaults to a tall Detail section).
Private Sub SetCompactDetailHeight(ByVal frm As Form, Optional ByVal rowTwips As Long = 360)
    On Error Resume Next
    frm.Section(acDetail).Height = rowTwips
    On Error GoTo 0
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
    OpenFormSized FRM_REFERENCES
    UiOpenReferences = True
    Exit Function
Fail:
    UiOpenReferences = False
End Function

Public Function UiOpenExport() As Boolean
    On Error GoTo Fail
    OpenFormSized FRM_EXPORT
    UiOpenExport = True
    Exit Function
Fail:
    UiOpenExport = False
End Function

Public Function UiCloseCurrentForm() As Boolean
    Dim formName As String
    On Error Resume Next
    ' Expression-button clicks often have no default "current" object for
    ' DoCmd.Close acForm with no name — close the active form explicitly.
    formName = Screen.ActiveForm.Name
    If Len(formName) = 0 Then
        formName = ParentFormName(Screen.ActiveControl)
    End If
    If Len(formName) > 0 Then
        DoCmd.Close acForm, formName, acSaveYes
    End If
    UiCloseCurrentForm = True
End Function

' Prefer this from OnClick when the form name is known.
Public Function UiCloseForm(ByVal formName As String) As Boolean
    On Error Resume Next
    DoCmd.Close acForm, formName, acSaveYes
    UiCloseForm = True
End Function

Private Function ParentFormName(ByVal ctl As Control) As String
    Dim p As Object
    On Error Resume Next
    Set p = ctl
    Do While Not p Is Nothing
        If TypeOf p Is Form Then
            ParentFormName = p.Name
            Exit Function
        End If
        Set p = p.Parent
    Loop
End Function

Public Function UiOpenFfaForm() As Boolean
    On Error GoTo Fail
    OpenFormSized FRM_FFA
    UiOpenFfaForm = True
    Exit Function
Fail:
    MsgBox Err.Description, vbExclamation, "Edit FFAs"
    UiOpenFfaForm = False
End Function

Public Function UiOpenProductLineForm() As Boolean
    On Error GoTo Fail
    OpenFormSized FRM_PRODUCT_LINE
    UiOpenProductLineForm = True
    Exit Function
Fail:
    MsgBox Err.Description, vbExclamation, "Edit Product Lines"
    UiOpenProductLineForm = False
End Function

Public Function UiOpenEquipmentForm() As Boolean
    On Error GoTo Fail
    OpenFormSized FRM_EQUIPMENT
    UiOpenEquipmentForm = True
    Exit Function
Fail:
    MsgBox Err.Description, vbExclamation, "Equipment"
    UiOpenEquipmentForm = False
End Function

Public Function UiOpenEquipmentEntryForm() As Boolean
    On Error GoTo Fail
    OpenFormSized FRM_EQUIPMENT_ENTRY
    UiClearEquipmentEntry
    UiOpenEquipmentEntryForm = True
    Exit Function
Fail:
    MsgBox Err.Description, vbExclamation, "Add Equipment"
    UiOpenEquipmentEntryForm = False
End Function

Public Function UiOpenEquipmentFfaForm() As Boolean
    On Error GoTo Fail
    OpenFormSized FRM_EQUIPMENT_FFA
    UiOpenEquipmentFfaForm = True
    Exit Function
Fail:
    MsgBox Err.Description, vbExclamation, "Equipment FFAs"
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
