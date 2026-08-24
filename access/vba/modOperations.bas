Attribute VB_Name = "modOperations"
Option Compare Database
Option Explicit

' Seeds tblOperation from Route_Card for a base part's active dash conditions.
' Imported hours / executions use the same average fallbacks as the Excel UDFs.
' Existing override columns are preserved when an op row already exists.

Public Sub SeedOperationsForPart(ByVal basePart As String)
    RebuildActiveAssemblyFilter
    Dim db As DAO.Database
    Dim rsDash As DAO.Recordset
    Dim rsRoute As DAO.Recordset
    Dim assemblyNo As String
    Dim opSeq As Variant
    Dim opCode As String
    Dim madeInFfa As String
    Dim importedHours As Variant
    Dim importedEx As Variant
    Dim sql As String

    basePart = Trim$(basePart)
    If Len(basePart) = 0 Then Exit Sub

    Set db = CurrentDb
    sql = "SELECT Dash FROM [" & TBL_PART_DASH & "] WHERE BasePart = " & SqlText(basePart) & " AND Active <> 0"
    Set rsDash = db.OpenRecordset(sql, dbOpenSnapshot)
    Do Until rsDash.EOF
        assemblyNo = basePart & "-" & CoerceText(rsDash!Dash)
        sql = "SELECT [" & COL_OPER_SEQ & "], [" & COL_OPER_CODE & "], [" & COL_FFA & "] " & _
            "FROM [" & RouteCardSourceName() & "] WHERE [" & COL_ASSEMBLY_NO & "] = " & SqlText(assemblyNo) & _
            " ORDER BY [" & COL_OPER_SEQ & "]"
        Set rsRoute = db.OpenRecordset(sql, dbOpenSnapshot)
        Do Until rsRoute.EOF
            opSeq = rsRoute.Fields(COL_OPER_SEQ).Value
            If Not IsNull(opSeq) Then
                opCode = CoerceText(rsRoute.Fields(COL_OPER_CODE).Value)
                madeInFfa = CoerceText(rsRoute.Fields(COL_FFA).Value)
                importedHours = AvgLaborHoursByBasePartAndOp(basePart, opSeq)
                importedEx = AvgProcTmYldByBasePartAndOp(basePart, opSeq)
                UpsertOperation db, basePart, CLng(opSeq), opCode, importedHours, importedEx, madeInFfa
            End If
            rsRoute.MoveNext
        Loop
        rsRoute.Close
        rsDash.MoveNext
    Loop
    rsDash.Close
End Sub

Public Sub SeedOperationsForActiveParts()
    RebuildActiveAssemblyFilter
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT BasePart FROM [" & TBL_PART & "] WHERE Active <> 0", dbOpenSnapshot)
    Do Until rs.EOF
        SeedOperationsForPart CStr(rs!BasePart)
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Sub UpsertOperation( _
    ByVal db As DAO.Database, _
    ByVal basePart As String, _
    ByVal opSeq As Long, _
    ByVal opCode As String, _
    ByVal importedHours As Variant, _
    ByVal importedEx As Variant, _
    ByVal madeInFfa As String)

    Dim existingId As Variant
    Dim sql As String

    existingId = DLookup("OperationID", TBL_OPERATION, _
        "BasePart = " & SqlText(basePart) & " AND OpSequence = " & opSeq)

    If IsNull(existingId) Then
        sql = "INSERT INTO [" & TBL_OPERATION & "] " & _
            "(BasePart, OpSequence, OpCode, ImportedHours, ImportedEx, BatchSize, " & _
            "ExportHours, ExportEx, EquipmentType, UseExportHours, UseExportEx, MadeInFFA) VALUES (" & _
            SqlText(basePart) & ", " & opSeq & ", " & SqlText(opCode) & ", " & SqlNullableNumber(importedHours) & ", " & _
            SqlNullableNumber(importedEx) & ", Null, Null, Null, Null, False, False, " & SqlNullableText(madeInFfa) & ")"
        db.Execute sql, dbFailOnError
    Else
        sql = "UPDATE [" & TBL_OPERATION & "] SET " & _
            "OpCode = " & SqlText(opCode) & ", " & _
            "ImportedHours = " & SqlNullableNumber(importedHours) & ", " & _
            "ImportedEx = " & SqlNullableNumber(importedEx) & ", " & _
            "MadeInFFA = " & SqlNullableText(madeInFfa) & " " & _
            "WHERE OperationID = " & CLng(existingId)
        db.Execute sql, dbFailOnError
    End If
End Sub

Private Function SqlNullableNumber(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        SqlNullableNumber = "Null"
    ElseIf Not IsNumeric(value) Then
        SqlNullableNumber = "Null"
    Else
        SqlNullableNumber = Str$(CDbl(value))
    End If
End Function

Private Function SqlNullableText(ByVal value As String) As String
    If Len(value) = 0 Then
        SqlNullableText = "Null"
    Else
        SqlNullableText = SqlText(value)
    End If
End Function
