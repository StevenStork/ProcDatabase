Attribute VB_Name = "modSchema"
Option Compare Database
Option Explicit

Public Sub EnsureSchema()
    EnsureMetaTable
    EnsureLookupTables
    EnsurePartTables
    EnsureActiveAssemblyFilterTable
    EnsureOptionalProcTmYldTable
    EnsureLinkedSourceTables
    UpgradeExistingSchema
End Sub

Private Sub EnsureMetaTable()
    Dim td As DAO.TableDef
    If TableExists(TBL_META) Then Exit Sub
    Set td = CurrentDb.CreateTableDef(TBL_META)
    AddTextField td, "Key", 50
    AddMemoField td, "Value"
    CurrentDb.TableDefs.Append td
    CurrentDb.Execute "CREATE UNIQUE INDEX PrimaryKey ON [" & TBL_META & "] ([Key])", dbFailOnError
End Sub

Private Sub EnsureLookupTables()
    Dim td As DAO.TableDef

    If Not TableExists(TBL_FFA) Then
        Set td = CurrentDb.CreateTableDef(TBL_FFA)
        AddTextField td, "FFA", 50
        AddTextField td, "Factory", 100
        CurrentDb.TableDefs.Append td
        CurrentDb.Execute "CREATE UNIQUE INDEX PrimaryKey ON [" & TBL_FFA & "] ([FFA])", dbFailOnError
    End If

    If Not TableExists(TBL_PRODUCT_LINE) Then
        Set td = CurrentDb.CreateTableDef(TBL_PRODUCT_LINE)
        AddTextField td, "ProductLine", 100
        CurrentDb.TableDefs.Append td
        CurrentDb.Execute "CREATE UNIQUE INDEX PrimaryKey ON [" & TBL_PRODUCT_LINE & "] ([ProductLine])", dbFailOnError
        SeedDefaultProductLines
    End If

    If Not TableExists(TBL_EQUIPMENT) Then
        Set td = CurrentDb.CreateTableDef(TBL_EQUIPMENT)
        AddTextField td, "Equipment", 100
        AddMemoField td, "OwningFFAs"
        CurrentDb.TableDefs.Append td
        CurrentDb.Execute "CREATE UNIQUE INDEX PrimaryKey ON [" & TBL_EQUIPMENT & "] ([Equipment])", dbFailOnError
    End If
End Sub

Private Sub EnsurePartTables()
    Dim td As DAO.TableDef

    If Not TableExists(TBL_PART) Then
        Set td = CurrentDb.CreateTableDef(TBL_PART)
        AddTextField td, "BasePart", 50
        AddBooleanField td, "Active"
        AddTextField td, "HomeFFA", 50
        AddDateField td, "StatusDate"
        AddTextField td, "Highlight", 255
        AddTextField td, COL_SHEET_NAME, 50
        CurrentDb.TableDefs.Append td
        CurrentDb.Execute "CREATE UNIQUE INDEX PrimaryKey ON [" & TBL_PART & "] ([BasePart])", dbFailOnError
    End If

    If Not TableExists(TBL_PART_DASH) Then
        Set td = CurrentDb.CreateTableDef(TBL_PART_DASH)
        AddTextField td, "BasePart", 50
        AddTextField td, "Dash", 50
        AddBooleanField td, "Active"
        CurrentDb.TableDefs.Append td
        CurrentDb.Execute "CREATE UNIQUE INDEX PrimaryKey ON [" & TBL_PART_DASH & "] ([BasePart], [Dash])", dbFailOnError
    End If

    If Not TableExists(TBL_PART_PL) Then
        Set td = CurrentDb.CreateTableDef(TBL_PART_PL)
        AddTextField td, "BasePart", 50
        AddTextField td, "ProductLine", 100
        AddBooleanField td, "UseFlag"
        CurrentDb.TableDefs.Append td
        CurrentDb.Execute "CREATE UNIQUE INDEX PrimaryKey ON [" & TBL_PART_PL & "] ([BasePart], [ProductLine])", dbFailOnError
    End If

    If Not TableExists(TBL_OPERATION) Then
        Set td = CurrentDb.CreateTableDef(TBL_OPERATION)
        AddAutoField td, "OperationID"
        AddTextField td, "BasePart", 50
        AddLongField td, "OpSequence"
        AddTextField td, "OpCode", 50
        AddDoubleField td, "ImportedHours"
        AddDoubleField td, "ImportedEx"
        AddDoubleField td, "BatchSize"
        AddDoubleField td, "ExportHours"
        AddDoubleField td, "ExportEx"
        AddTextField td, "EquipmentType", 100
        AddBooleanField td, "UseExportHours"
        AddBooleanField td, "UseExportEx"
        AddTextField td, "MadeInFFA", 50
        CurrentDb.TableDefs.Append td
        CurrentDb.Execute "CREATE UNIQUE INDEX PrimaryKey ON [" & TBL_OPERATION & "] ([OperationID])", dbFailOnError
        CurrentDb.Execute "CREATE UNIQUE INDEX ux_ops_part_seq ON [" & TBL_OPERATION & "] ([BasePart], [OpSequence])", dbFailOnError
    End If
End Sub

Private Sub UpgradeExistingSchema()
    If TableExists(TBL_PART) Then
        AddTextFieldIfMissing TBL_PART, COL_SHEET_NAME, 50
        CurrentDb.Execute "UPDATE [" & TBL_PART & "] SET [" & COL_SHEET_NAME & "] = [BasePart] " & _
            "WHERE [" & COL_SHEET_NAME & "] IS NULL OR [" & COL_SHEET_NAME & "] = ''", dbFailOnError
    End If
End Sub

Private Sub EnsureOptionalProcTmYldTable()
    Dim td As DAO.TableDef
    If TableExists(TBL_PROC_TM_YLD) Then Exit Sub
    Set td = CurrentDb.CreateTableDef(TBL_PROC_TM_YLD)
    AddTextField td, COL_ASSEMBLY_NO_ALT, 50
    AddLongField td, COL_OPER_SEQ
    AddDoubleField td, COL_AVG_180
    AddDoubleField td, COL_AVG_90
    CurrentDb.TableDefs.Append td
    CurrentDb.Execute "CREATE INDEX ix_pty_assy ON [" & TBL_PROC_TM_YLD & "] ([" & COL_ASSEMBLY_NO_ALT & "], [" & COL_OPER_SEQ & "])", dbFailOnError
End Sub

Private Sub SeedDefaultProductLines()
    CurrentDb.Execute "INSERT INTO [" & TBL_PRODUCT_LINE & "] (ProductLine) VALUES ('Commercial')", dbFailOnError
    CurrentDb.Execute "INSERT INTO [" & TBL_PRODUCT_LINE & "] (ProductLine) VALUES ('Military')", dbFailOnError
    CurrentDb.Execute "INSERT INTO [" & TBL_PRODUCT_LINE & "] (ProductLine) VALUES ('Spare')", dbFailOnError
End Sub

Private Sub AddTextField(ByVal td As DAO.TableDef, ByVal fieldName As String, ByVal size As Long)
    Dim fld As DAO.Field
    Set fld = td.CreateField(fieldName, dbText, size)
    fld.AllowZeroLength = True
    fld.Required = False
    td.Fields.Append fld
End Sub

Private Sub AddMemoField(ByVal td As DAO.TableDef, ByVal fieldName As String)
    Dim fld As DAO.Field
    Set fld = td.CreateField(fieldName, dbMemo)
    fld.AllowZeroLength = True
    td.Fields.Append fld
End Sub

Private Sub AddLongField(ByVal td As DAO.TableDef, ByVal fieldName As String)
    td.Fields.Append td.CreateField(fieldName, dbLong)
End Sub

Private Sub AddDoubleField(ByVal td As DAO.TableDef, ByVal fieldName As String)
    td.Fields.Append td.CreateField(fieldName, dbDouble)
End Sub

Private Sub AddBooleanField(ByVal td As DAO.TableDef, ByVal fieldName As String)
    Dim fld As DAO.Field
    Set fld = td.CreateField(fieldName, dbBoolean)
    fld.DefaultValue = "False"
    td.Fields.Append fld
End Sub

Private Sub AddDateField(ByVal td As DAO.TableDef, ByVal fieldName As String)
    td.Fields.Append td.CreateField(fieldName, dbDate)
End Sub

Private Sub AddAutoField(ByVal td As DAO.TableDef, ByVal fieldName As String)
    Dim fld As DAO.Field
    Set fld = td.CreateField(fieldName, dbLong)
    fld.Attributes = fld.Attributes Or dbAutoIncrField
    td.Fields.Append fld
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
    SetDbProperty "StartupForm", dbText, FRM_HOME
    SetDbProperty "AppTitle", dbText, "ProcDatabase"
    Application.RefreshTitleBar
End Sub
