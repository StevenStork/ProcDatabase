Attribute VB_Name = "modSchema"
Option Compare Database
Option Explicit

Public SchemaSubStep As String

Public Sub EnsureSchema()
    On Error GoTo Fail
    SchemaSubStep = "EnsureMetaTable"
    EnsureMetaTable
    SchemaSubStep = "EnsureLookupTables"
    EnsureLookupTables
    SchemaSubStep = "EnsurePartTables"
    EnsurePartTables
    SchemaSubStep = "EnsureActiveAssemblyFilterTable"
    EnsureActiveAssemblyFilterTable
    SchemaSubStep = "UpgradeExistingSchema"
    UpgradeExistingSchema
    SchemaSubStep = vbNullString
    Exit Sub
Fail:
    Err.Raise Err.Number, "EnsureSchema." & SchemaSubStep, Err.Description
End Sub

Public Sub DiagnoseSchema()
    Dim report As String
    report = "ProcDatabase schema diagnosis:" & vbCrLf & vbCrLf
    report = report & DescribeLinkedTables()
    report = report & vbCrLf & "Local tables:" & vbCrLf
    report = report & DescribeLocalTables()
    MsgBox report, vbInformation, "DiagnoseSchema"
End Sub

Private Function DescribeLinkedTables() As String
    Dim lines As String
    lines = "Linked source tables:" & vbCrLf
    lines = lines & DescribeTable(TBL_ROUTE_CARD)
    lines = lines & DescribeTable(TBL_ASSY_STANDARD)
    lines = lines & DescribeTable(TBL_OPER_COMPLETIONS)
    lines = lines & DescribeTable(TBL_RCCP)
    lines = lines & DescribeTable(TBL_PROC_TM_YLD)
    DescribeLinkedTables = lines
End Function

Private Function DescribeLocalTables() As String
    Dim names As Variant
    Dim i As Long
    Dim lines As String
    names = Array(TBL_META, TBL_FFA, TBL_PRODUCT_LINE, TBL_EQUIPMENT, TBL_EQUIPMENT_FFA, TBL_PART, _
        TBL_PART_DASH, TBL_PART_PL, TBL_OPERATION, TBL_ACTIVE_FILTER)
    For i = LBound(names) To UBound(names)
        lines = lines & DescribeTable(CStr(names(i)))
    Next i
    DescribeLocalTables = lines
End Function

Private Function DescribeTable(ByVal tableName As String) As String
    If Not TableExists(tableName) Then
        DescribeTable = "  - " & tableName & ": missing" & vbCrLf
        Exit Function
    End If
    If IsLinkedTable(tableName) Then
        DescribeTable = "  - " & tableName & ": linked" & vbCrLf
    Else
        DescribeTable = "  - " & tableName & ": local" & vbCrLf
    End If
End Function

Private Sub EnsureMetaTable()
    If TableExists(TBL_META) Then Exit Sub
    ExecuteDDL "CREATE TABLE [" & TBL_META & "] (" & _
        "[MetaKey] TEXT(50) CONSTRAINT PK_tblMeta PRIMARY KEY, " & _
        "[MetaValue] MEMO" & _
        ")"
End Sub

Private Sub EnsureLookupTables()
    If Not TableExists(TBL_FFA) Then
        ExecuteDDL "CREATE TABLE [" & TBL_FFA & "] (" & _
            "[FFA] TEXT(50) CONSTRAINT PK_tblFFA PRIMARY KEY, " & _
            "[Factory] TEXT(100)" & _
            ")"
    End If

    If Not TableExists(TBL_PRODUCT_LINE) Then
        ExecuteDDL "CREATE TABLE [" & TBL_PRODUCT_LINE & "] (" & _
            "[ProductLine] TEXT(100) CONSTRAINT PK_tblProductLine PRIMARY KEY, " & _
            "[" & COL_PL_CODE & "] TEXT(50)" & _
            ")"
        SeedDefaultProductLines
    End If

    If Not TableExists(TBL_EQUIPMENT) Then
        ExecuteDDL "CREATE TABLE [" & TBL_EQUIPMENT & "] (" & _
            "[Equipment] TEXT(100) CONSTRAINT PK_tblEquipment PRIMARY KEY, " & _
            "[" & COL_EQUIP_TYPE & "] TEXT(100)" & _
            ")"
    End If

    EnsureEquipmentFfaTable
    EnsureEquipmentTypeColumn
End Sub

Public Sub EnsureEquipmentFfaTable()
    If TableExists(TBL_EQUIPMENT_FFA) Then Exit Sub
    ExecuteDDL "CREATE TABLE [" & TBL_EQUIPMENT_FFA & "] (" & _
        "[Equipment] TEXT(100), " & _
        "[FFA] TEXT(50), " & _
        "CONSTRAINT PK_tblEquipmentFFA PRIMARY KEY ([Equipment], [FFA])" & _
        ")"
End Sub

Public Sub EnsureEquipmentTypeColumn()
    If Not TableExists(TBL_EQUIPMENT) Then Exit Sub
    AddTextColumnIfMissing TBL_EQUIPMENT, COL_EQUIP_TYPE, 100
End Sub

Private Sub EnsurePartTables()
    If Not TableExists(TBL_PART) Then
        ExecuteDDL "CREATE TABLE [" & TBL_PART & "] (" & _
            "[BasePart] TEXT(50) CONSTRAINT PK_tblPart PRIMARY KEY, " & _
            "[Active] YESNO, " & _
            "[HomeFFA] TEXT(50), " & _
            "[StatusDate] DATETIME, " & _
            "[" & COL_NOTES & "] TEXT(255)" & _
            ")"
    End If

    If Not TableExists(TBL_PART_DASH) Then
        ExecuteDDL "CREATE TABLE [" & TBL_PART_DASH & "] (" & _
            "[BasePart] TEXT(50), " & _
            "[Dash] TEXT(50), " & _
            "[Active] YESNO, " & _
            "CONSTRAINT PK_tblPartDash PRIMARY KEY ([BasePart], [Dash])" & _
            ")"
    End If

    If Not TableExists(TBL_PART_PL) Then
        ExecuteDDL "CREATE TABLE [" & TBL_PART_PL & "] (" & _
            "[BasePart] TEXT(50), " & _
            "[ProductLine] TEXT(100), " & _
            "[UseFlag] YESNO, " & _
            "CONSTRAINT PK_tblPartPL PRIMARY KEY ([BasePart], [ProductLine])" & _
            ")"
    End If

    If Not TableExists(TBL_OPERATION) Then
        ExecuteDDL "CREATE TABLE [" & TBL_OPERATION & "] (" & _
            "[OperationID] COUNTER CONSTRAINT PK_tblOperation PRIMARY KEY, " & _
            "[BasePart] TEXT(50), " & _
            "[OpSequence] LONG, " & _
            "[OpCode] TEXT(50), " & _
            "[ProcessHours] DOUBLE, " & _
            "[AvgEx] DOUBLE, " & _
            "[BatchSize] DOUBLE, " & _
            "[ImportedHours] DOUBLE, " & _
            "[ImportedEx] DOUBLE, " & _
            "[ExportHours] DOUBLE, " & _
            "[ExportEx] DOUBLE, " & _
            "[Equipment] TEXT(100), " & _
            "[EquipmentType] TEXT(100), " & _
            "[UseExportHours] YESNO, " & _
            "[UseExportEx] YESNO, " & _
            "[MadeInFFA] TEXT(50)" & _
            ")"
        ExecuteDDL "CREATE UNIQUE INDEX ux_ops_part_seq ON [" & TBL_OPERATION & "] ([BasePart], [OpSequence])"
    End If
End Sub

Private Sub UpgradeExistingSchema()
    If Not TableExists(TBL_PART) Then Exit Sub
    On Error Resume Next
    CurrentDb.TableDefs.Refresh
    On Error GoTo 0
    MigratePartColumns
    EnsureOperationManualColumns
    EnsureProductLinePlCodeColumn
    EnsureEquipmentFfaTable
    EnsureEquipmentTypeColumn
    MigrateEquipmentOwningFfas
    MigrateLegacyMetaColumns
End Sub

' Process Hours / Avg Ex are manual inputs (not derived from Import/Export overrides).
' Equipment is chosen from equipment linked to MadeInFFA; EquipmentType is derived.
Public Sub EnsureOperationManualColumns()
    If Not TableExists(TBL_OPERATION) Then Exit Sub
    AddDoubleColumnIfMissing TBL_OPERATION, "ProcessHours"
    AddDoubleColumnIfMissing TBL_OPERATION, "AvgEx"
    AddTextColumnIfMissing TBL_OPERATION, COL_EQUIPMENT, 100
    AddTextColumnIfMissing TBL_OPERATION, "EquipmentType", 100
    AddTextColumnIfMissing TBL_OPERATION, "MadeInFFA", 50
End Sub

' Move legacy OwningFFAs memo into tblEquipmentFFA, then drop the memo column.
Public Sub MigrateEquipmentOwningFfas()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim equipName As String
    Dim owners As String
    Dim parts() As String
    Dim i As Long
    Dim ffaValue As String

    If Not TableExists(TBL_EQUIPMENT) Then Exit Sub
    EnsureEquipmentFfaTable
    If Not FieldExists(TBL_EQUIPMENT, "OwningFFAs") Then Exit Sub

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT Equipment, OwningFFAs FROM [" & TBL_EQUIPMENT & "]", dbOpenSnapshot)
    Do Until rs.EOF
        equipName = CoerceText(rs!Equipment)
        owners = CoerceText(rs!OwningFFAs)
        If Len(equipName) > 0 And Len(owners) > 0 Then
            parts = Split(owners, ",")
            For i = LBound(parts) To UBound(parts)
                ffaValue = Trim$(parts(i))
                If Len(ffaValue) > 0 Then
                    If IsNull(DLookup("FFA", TBL_EQUIPMENT_FFA, _
                        "Equipment = " & SqlText(equipName) & " AND FFA = " & SqlText(ffaValue))) Then
                        On Error Resume Next
                        db.Execute "INSERT INTO [" & TBL_EQUIPMENT_FFA & "] (Equipment, FFA) VALUES (" & _
                            SqlText(equipName) & ", " & SqlText(ffaValue) & ")", dbFailOnError
                        On Error GoTo 0
                    End If
                End If
            Next i
        End If
        rs.MoveNext
    Loop
    rs.Close

    DropColumnIfExists TBL_EQUIPMENT, "OwningFFAs"
End Sub

Public Sub EnsureProductLinePlCodeColumn()
    If Not TableExists(TBL_PRODUCT_LINE) Then Exit Sub
    AddTextColumnIfMissing TBL_PRODUCT_LINE, COL_PL_CODE, 50
End Sub

' Highlight -> Notes; drop unused SheetName. Safe to re-run.
Public Sub MigratePartColumns()
    Dim db As DAO.Database
    If Not TableExists(TBL_PART) Then Exit Sub
    Set db = CurrentDb

    If Not FieldExists(TBL_PART, COL_NOTES) Then
        AddTextColumnIfMissing TBL_PART, COL_NOTES, 255
        If FieldExists(TBL_PART, "Highlight") Then
            On Error Resume Next
            db.Execute "UPDATE [" & TBL_PART & "] SET [" & COL_NOTES & "] = [Highlight] " & _
                "WHERE [" & COL_NOTES & "] IS NULL OR [" & COL_NOTES & "] = ''", dbFailOnError
            On Error GoTo 0
        End If
    End If

    DropColumnIfExists TBL_PART, "Highlight"
    DropColumnIfExists TBL_PART, "SheetName"
End Sub

Private Sub MigrateLegacyMetaColumns()
    If Not TableExists(TBL_META) Then Exit Sub
    If FieldExists(TBL_META, "MetaKey") Then Exit Sub
    If Not FieldExists(TBL_META, "Key") Then Exit Sub
    AddTextColumnIfMissing TBL_META, "MetaKey", 50
    AddMemoColumnIfMissing TBL_META, "MetaValue"
    On Error Resume Next
    CurrentDb.Execute "UPDATE [" & TBL_META & "] SET [MetaKey]=[Key], [MetaValue]=[Value]", dbFailOnError
    On Error GoTo 0
End Sub

Private Sub EnsureOptionalProcTmYldTable()
    ' Optional linked yield source — intentionally not auto-created.
End Sub

Private Sub SeedDefaultProductLines()
    InsertProductLineIfMissing "Commercial"
    InsertProductLineIfMissing "Military"
    InsertProductLineIfMissing "Spare"
End Sub

Private Sub InsertProductLineIfMissing(ByVal productLine As String)
    If IsNull(DLookup("ProductLine", TBL_PRODUCT_LINE, "ProductLine = " & SqlText(productLine))) Then
        CurrentDb.Execute "INSERT INTO [" & TBL_PRODUCT_LINE & "] (ProductLine) VALUES (" & SqlText(productLine) & ")", dbFailOnError
    End If
End Sub

Private Sub ExecuteDDL(ByVal sql As String)
    On Error GoTo Fail
    CurrentDb.Execute sql, dbFailOnError
    Exit Sub
Fail:
    Select Case Err.Number
        Case 3010, 3012, 3029, 3191, 3288
            Err.Clear
        Case Else
            Err.Raise Err.Number, "ExecuteDDL", Err.Description & vbCrLf & vbCrLf & sql
    End Select
End Sub

Public Sub EnsureQueries()
    MigratePartColumns
    EnsureOperationManualColumns
    EnsureProductLinePlCodeColumn
    EnsureEquipmentFfaTable
    EnsureEquipmentTypeColumn
    MigrateEquipmentOwningFfas

    ' ProcessHours / AvgEx are stored manual fields. AvgHPU uses Import overrides when checked
    ' and the imported value is present; otherwise falls back to the manual field:
    ' hours = IIf(UseImportHrs And ImportedHours Not Null, ImportedHours, ProcessHours)
    ' ex    = IIf(UseImportEx  And ImportedEx  Not Null, ImportedEx,    AvgEx)
    ' HPU   = (hours * ex) / BatchSize
    ReplaceQuery QRY_OPERATIONS, _
        "SELECT o.*, " & _
        "IIf(Nz(o.BatchSize,0)=0 OR " & _
        "IIf(o.UseExportHours<>0 And o.ImportedHours Is Not Null, o.ImportedHours, o.ProcessHours) IS NULL OR " & _
        "IIf(o.UseExportEx<>0 And o.ImportedEx Is Not Null, o.ImportedEx, o.AvgEx) IS NULL, Null, " & _
        "(IIf(o.UseExportHours<>0 And o.ImportedHours Is Not Null, o.ImportedHours, o.ProcessHours) * " & _
        "IIf(o.UseExportEx<>0 And o.ImportedEx Is Not Null, o.ImportedEx, o.AvgEx)) / o.BatchSize) AS AvgHPU " & _
        "FROM [" & TBL_OPERATION & "] AS o"

    ReplaceQuery QRY_HOME, _
        "SELECT p.BasePart, p.Active, p.StatusDate, " & _
        "IIf(p.StatusDate IS NULL, Null, DateDiff('d', p.StatusDate, Date())) AS Days, " & _
        "p.[" & COL_NOTES & "], p.HomeFFA, f.Factory AS Factories " & _
        "FROM [" & TBL_PART & "] AS p LEFT JOIN [" & TBL_FFA & "] AS f ON p.HomeFFA = f.FFA"

    ReplaceQuery QRY_EXPORT, _
        "SELECT q.BasePart AS [Part Number], q.OpSequence AS [Op Sequence], q.OpCode AS [Op Code], " & _
        "q.ProcessHours AS [Process Hours], q.AvgEx AS [Avg Ex], q.BatchSize AS [Batch Size], " & _
        "q.AvgHPU AS [Avg HPU], q.Equipment AS [Equipment], q.EquipmentType AS [Equipment Type], " & _
        "p.HomeFFA AS [Home FFA], q.MadeInFFA AS [Made In FFA] " & _
        "FROM [" & QRY_OPERATIONS & "] AS q INNER JOIN [" & TBL_PART & "] AS p ON q.BasePart = p.BasePart " & _
        "WHERE p.Active <> 0 AND EXISTS (" & _
        "SELECT 1 FROM [" & TBL_PART_DASH & "] AS d WHERE d.BasePart = p.BasePart AND d.Active <> 0)"

    EnsureFilteredSourceQueries
End Sub

Private Sub EnsureFilteredSourceQueries()
    If Not LinkedSourceTablesReady() Then Exit Sub

    ReplaceQuery QRY_ROUTE_CARD_ACTIVE, _
        "SELECT rc.* FROM [" & TBL_ROUTE_CARD & "] AS rc " & _
        "INNER JOIN [" & TBL_ACTIVE_FILTER & "] AS f ON rc.[" & COL_ASSEMBLY_NO & "] = f.[" & COL_ASSEMBLY_NO_FILTER & "]"

    ReplaceQuery QRY_ASSY_STND_ACTIVE, _
        "SELECT st.* FROM [" & TBL_ASSY_STANDARD & "] AS st " & _
        "INNER JOIN [" & TBL_ACTIVE_FILTER & "] AS f ON st.[" & COL_ASSEMBLY_NO & "] = f.[" & COL_ASSEMBLY_NO_FILTER & "]"

    ReplaceQuery QRY_OPER_COMPS_ACTIVE, _
        "SELECT oc.* FROM [" & TBL_OPER_COMPLETIONS & "] AS oc " & _
        "INNER JOIN [" & TBL_ACTIVE_FILTER & "] AS f ON oc.[" & COL_ASSEMBLY_NO & "] = f.[" & COL_ASSEMBLY_NO_FILTER & "]"
End Sub

Public Sub EnsureStartup()
    On Error Resume Next
    If ObjectExists(acForm, FRM_HOME) Then
        SetDbProperty PROP_STARTUP_FORM, dbText, FRM_HOME
    Else
        SetDbProperty PROP_STARTUP_FORM, dbText, vbNullString
    End If
    SetDbProperty PROP_APP_TITLE, dbText, APP_TITLE
    Application.RefreshTitleBar
    On Error GoTo 0
End Sub
