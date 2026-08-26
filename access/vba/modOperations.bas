Attribute VB_Name = "modOperations"
Option Compare Database
Option Explicit

' Seeds tblOperation from Route_Card for a base part's active dash conditions.
' Imported hours / executions use the same average fallbacks as the Excel UDFs.
' Existing override columns are preserved when an op row already exists.

Public Sub SeedOperationsForPart(ByVal basePart As String, Optional ByVal rebuildFilter As Boolean = True)
    Dim db As DAO.Database
    Dim rsDash As DAO.Recordset
    Dim rsRoute As DAO.Recordset
    Dim existingOps As Object
    Dim assemblyNo As String
    Dim opSeq As Variant
    Dim opCode As String
    Dim madeInFfa As String
    Dim importedHours As Variant
    Dim importedEx As Variant
    Dim routeSource As String
    Dim sql As String
    Dim ownCache As Boolean

    basePart = Trim$(basePart)
    If Len(basePart) = 0 Then Exit Sub

    If rebuildFilter Then RebuildActiveAssemblyFilter

    ' Single-part seed from the UI owns its average cache; batch seed wraps outside.
    ownCache = Not AverageCacheIsActive()
    If ownCache Then BeginAverageCache

    On Error GoTo CleanUp
    Set db = CurrentDb
    routeSource = RouteCardSourceName()
    Set existingOps = LoadExistingOpMap(db, basePart)

    sql = "SELECT Dash FROM [" & TBL_PART_DASH & "] WHERE BasePart = " & SqlText(basePart) & _
        " AND Active <> 0"
    Set rsDash = db.OpenRecordset(sql, dbOpenSnapshot)
    Do Until rsDash.EOF
        assemblyNo = basePart & "-" & CoerceText(rsDash!Dash)
        sql = "SELECT [" & COL_OPER_SEQ & "], [" & COL_OPER_CODE & "], [" & COL_FFA & "] " & _
            "FROM [" & routeSource & "] WHERE [" & COL_ASSEMBLY_NO & "] = " & SqlText(assemblyNo) & _
            " ORDER BY [" & COL_OPER_SEQ & "]"
        Set rsRoute = db.OpenRecordset(sql, dbOpenSnapshot)
        Do Until rsRoute.EOF
            opSeq = rsRoute.Fields(COL_OPER_SEQ).Value
            If Not IsNull(opSeq) Then
                opCode = CoerceText(rsRoute.Fields(COL_OPER_CODE).Value)
                madeInFfa = CoerceText(rsRoute.Fields(COL_FFA).Value)
                importedHours = AvgLaborHoursByBasePartAndOp(basePart, opSeq)
                importedEx = AvgProcTmYldByBasePartAndOp(basePart, opSeq)
                UpsertOperation db, existingOps, basePart, CLng(opSeq), opCode, importedHours, importedEx, madeInFfa
            End If
            rsRoute.MoveNext
        Loop
        rsRoute.Close
        rsDash.MoveNext
    Loop
    rsDash.Close

CleanUp:
    If ownCache Then EndAverageCache
    If Err.Number <> 0 Then Err.Raise Err.Number, "SeedOperationsForPart", Err.Description
End Sub

Public Sub SeedOperationsForActiveParts()
    Dim rs As DAO.Recordset
    Dim db As DAO.Database

    RebuildActiveAssemblyFilter
    BeginAverageCache
    On Error GoTo CleanUp
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT BasePart FROM [" & TBL_PART & "] WHERE Active <> 0", dbOpenSnapshot)
    Do Until rs.EOF
        SeedOperationsForPart CStr(rs!BasePart), False
        rs.MoveNext
    Loop
    rs.Close

CleanUp:
    EndAverageCache
    If Err.Number <> 0 Then Err.Raise Err.Number, "SeedOperationsForActiveParts", Err.Description
End Sub

Private Function LoadExistingOpMap(ByVal db As DAO.Database, ByVal basePart As String) As Object
    Dim rs As DAO.Recordset
    Dim map As Object
    Dim key As String

    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = vbTextCompare
    Set rs = db.OpenRecordset( _
        "SELECT OperationID, OpSequence FROM [" & TBL_OPERATION & "] WHERE BasePart = " & SqlText(basePart), _
        dbOpenSnapshot)
    Do Until rs.EOF
        key = CStr(CLng(rs!OpSequence))
        If Not map.Exists(key) Then map.Add key, CLng(rs!OperationID)
        rs.MoveNext
    Loop
    rs.Close
    Set LoadExistingOpMap = map
End Function

Private Sub UpsertOperation( _
    ByVal db As DAO.Database, _
    ByVal existingOps As Object, _
    ByVal basePart As String, _
    ByVal opSeq As Long, _
    ByVal opCode As String, _
    ByVal importedHours As Variant, _
    ByVal importedEx As Variant, _
    ByVal madeInFfa As String)

    Dim key As String
    Dim sql As String

    key = CStr(opSeq)
    If Not existingOps.Exists(key) Then
        sql = "INSERT INTO [" & TBL_OPERATION & "] " & _
            "(BasePart, OpSequence, OpCode, ImportedHours, ImportedEx, BatchSize, " & _
            "ExportHours, ExportEx, EquipmentType, UseExportHours, UseExportEx, MadeInFFA) VALUES (" & _
            SqlText(basePart) & ", " & opSeq & ", " & SqlText(opCode) & ", " & SqlNullableNumber(importedHours) & ", " & _
            SqlNullableNumber(importedEx) & ", Null, Null, Null, Null, False, False, " & SqlNullableText(madeInFfa) & ")"
        db.Execute sql, dbFailOnError
        existingOps.Add key, True
    Else
        sql = "UPDATE [" & TBL_OPERATION & "] SET " & _
            "OpCode = " & SqlText(opCode) & ", " & _
            "ImportedHours = " & SqlNullableNumber(importedHours) & ", " & _
            "ImportedEx = " & SqlNullableNumber(importedEx) & ", " & _
            "MadeInFFA = " & SqlNullableText(madeInFfa) & " " & _
            "WHERE BasePart = " & SqlText(basePart) & " AND OpSequence = " & opSeq
        db.Execute sql, dbFailOnError
    End If
End Sub
