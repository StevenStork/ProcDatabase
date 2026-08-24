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
    SchemaSubStep = "EnsureOptionalProcTmYldTable"
    EnsureOptionalProcTmYldTable
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
    lines = lines & DescribeTable(TBL_PROC_TM_YLD)
    DescribeLinkedTables = lines
End Function

Private Function DescribeLocalTables() As String
    Dim names As Variant
    Dim i As Long
    Dim lines As String
    names = Array(TBL_META, TBL_FFA, TBL_PRODUCT_LINE, TBL_EQUIPMENT, TBL_PART, _
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
            "[ProductLine] TEXT(100) CONSTRAINT PK_tblProductLine PRIMARY KEY" & _
            ")"
        SeedDefaultProductLines
    End If

    If Not TableExists(TBL_EQUIPMENT) Then
        ExecuteDDL "CREATE TABLE [" & TBL_EQUIPMENT & "] (" & _
            "[Equipment] TEXT(100) CONSTRAINT PK_tblEquipment PRIMARY KEY, " & _
            "[OwningFFAs] MEMO" & _
            ")"
    End If
End Sub

Private Sub EnsurePartTables()
    If Not TableExists(TBL_PART) Then
        ExecuteDDL "CREATE TABLE [" & TBL_PART & "] (" & _
            "[BasePart] TEXT(50) CONSTRAINT PK_tblPart PRIMARY KEY, " & _
            "[Active] YESNO, " & _
            "[HomeFFA] TEXT(50), " & _
            "[StatusDate] DATETIME, " & _
            "[Highlight] TEXT(255), " & _
            "[" & COL_SHEET_NAME & "] TEXT(50)" & _
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
            "[ImportedHours] DOUBLE, " & _
            "[ImportedEx] DOUBLE, " & _
            "[BatchSize] DOUBLE, " & _
            "[ExportHours] DOUBLE, " & _
            "[ExportEx] DOUBLE, " & _
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
    AddTextColumnIfMissing TBL_PART, COL_SHEET_NAME, 50
    On Error Resume Next
    CurrentDb.Execute "UPDATE [" & TBL_PART & "] SET [" & COL_SHEET_NAME & "] = [BasePart] " & _
        "WHERE [" & COL_SHEET_NAME & "] IS NULL OR [" & COL_SHEET_NAME & "] = ''", dbFailOnError
    On Error GoTo 0
    MigrateLegacyMetaColumns
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
    If TableExists(TBL_PROC_TM_YLD) Then Exit Sub
    ExecuteDDL "CREATE TABLE [" & TBL_PROC_TM_YLD & "] (" & _
        "[" & COL_ASSEMBLY_NO_ALT & "] TEXT(50), " & _
        "[" & COL_OPER_SEQ & "] LONG, " & _
        "[" & COL_AVG_180 & "] DOUBLE, " & _
        "[" & COL_AVG_90 & "] DOUBLE" & _
        ")"
    ExecuteDDL "CREATE INDEX ix_pty_assy ON [" & TBL_PROC_TM_YLD & "] ([" & COL_ASSEMBLY_NO_ALT & "], [" & COL_OPER_SEQ & "])"
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
    ReplaceQuery QRY_OPERATIONS, _
        "SELECT q.*, " & _
        "IIf(Nz(q.BatchSize,0)=0 OR q.ProcessHours IS NULL OR q.AvgEx IS NULL, Null, (q.ProcessHours * q.AvgEx) / q.BatchSize) AS AvgHPU " & _
        "FROM (" & _
        "SELECT o.*, " & _
        "IIf(o.UseExportHours <> 0 AND o.ExportHours IS NOT NULL, o.ExportHours, o.ImportedHours) AS ProcessHours, " & _
        "IIf(o.UseExportEx <> 0 AND o.ExportEx IS NOT NULL, o.ExportEx, o.ImportedEx) AS AvgEx " & _
        "FROM [" & TBL_OPERATION & "] AS o" & _
        ") AS q"

    ReplaceQuery QRY_HOME, _
        "SELECT p.BasePart, p.Active, p.StatusDate, " & _
        "IIf(p.StatusDate IS NULL, Null, DateDiff('d', p.StatusDate, Date())) AS Days, " & _
        "p.Highlight, p.HomeFFA, p.[" & COL_SHEET_NAME & "], f.Factory AS Factories " & _
        "FROM [" & TBL_PART & "] AS p LEFT JOIN [" & TBL_FFA & "] AS f ON p.HomeFFA = f.FFA"

    ReplaceQuery QRY_EXPORT, _
        "SELECT q.BasePart AS [Part Number], q.OpSequence AS [Op Sequence], q.OpCode AS [Op Code], " & _
        "q.ProcessHours AS [Process Hours], q.AvgEx AS [Avg Ex], q.BatchSize AS [Batch Size], " & _
        "q.AvgHPU AS [Avg HPU], q.EquipmentType AS [Equipment Type], " & _
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
        SetDbProperty "StartupForm", dbText, FRM_HOME
    Else
        SetDbProperty "StartupForm", dbText, vbNullString
    End If
    SetDbProperty "AppTitle", dbText, "ProcDatabase"
    Application.RefreshTitleBar
    On Error GoTo 0
End Sub
