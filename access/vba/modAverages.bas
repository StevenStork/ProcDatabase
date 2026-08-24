Attribute VB_Name = "modAverages"
Option Compare Database
Option Explicit

' Ports AvgLaborHoursByBasePartAndOp and AvgProcTmYldByBasePartAndOp.
' Zero / null / blank values are excluded. Labor falls back to standards
' run time. Process-time executions prefer 180-day then 90-day.

Public Function AvgLaborHoursByBasePartAndOp(ByVal BasePartNumber As Variant, ByVal OpSequence As Variant) As Variant
    Dim basePart As String
    Dim opSeq As String
    Dim completionsAvg As Variant
    Dim standardsAvg As Variant

    On Error GoTo Fail
    basePart = Trim$(CStr(Nz(BasePartNumber, vbNullString)))
    opSeq = Trim$(CStr(Nz(OpSequence, vbNullString)))
    If Len(basePart) = 0 Or Len(opSeq) = 0 Then
        AvgLaborHoursByBasePartAndOp = Null
        Exit Function
    End If

    completionsAvg = AverageHoursForMatch(TBL_OPER_COMPLETIONS, COL_ASSEMBLY_NO, COL_OPER_SEQ, COL_LABOR_HPS, basePart, opSeq)
    If IsNumeric(completionsAvg) Then
        If CDbl(completionsAvg) <> 0 Then
            AvgLaborHoursByBasePartAndOp = CDbl(completionsAvg)
            Exit Function
        End If
    End If

    standardsAvg = AverageHoursForMatch(TBL_ASSY_STANDARD, COL_ASSEMBLY_NO, COL_OPER_SEQ, COL_RUN_TIME, basePart, opSeq)
    If IsNumeric(standardsAvg) Then
        AvgLaborHoursByBasePartAndOp = CDbl(standardsAvg)
    Else
        AvgLaborHoursByBasePartAndOp = Null
    End If
    Exit Function
Fail:
    AvgLaborHoursByBasePartAndOp = Null
End Function

Public Function AvgProcTmYldByBasePartAndOp(ByVal BasePartNumber As Variant, ByVal OpSequence As Variant) As Variant
    Dim basePart As String
    Dim opSeq As String
    Dim avg180 As Variant
    Dim avg90 As Variant

    On Error GoTo Fail
    basePart = Trim$(CStr(Nz(BasePartNumber, vbNullString)))
    opSeq = Trim$(CStr(Nz(OpSequence, vbNullString)))
    If Len(basePart) = 0 Or Len(opSeq) = 0 Then
        AvgProcTmYldByBasePartAndOp = Null
        Exit Function
    End If
    If Not TableExists(TBL_PROC_TM_YLD) Then
        AvgProcTmYldByBasePartAndOp = Null
        Exit Function
    End If
    If DCount("*", TBL_PROC_TM_YLD) = 0 Then
        AvgProcTmYldByBasePartAndOp = Null
        Exit Function
    End If

    avg180 = AverageHoursForMatch(TBL_PROC_TM_YLD, COL_ASSEMBLY_NO_ALT, COL_OPER_SEQ, COL_AVG_180, basePart, opSeq)
    If IsNumeric(avg180) Then
        If CDbl(avg180) <> 0 Then
            AvgProcTmYldByBasePartAndOp = CDbl(avg180)
            Exit Function
        End If
    End If

    avg90 = AverageHoursForMatch(TBL_PROC_TM_YLD, COL_ASSEMBLY_NO_ALT, COL_OPER_SEQ, COL_AVG_90, basePart, opSeq)
    If IsNumeric(avg90) Then
        AvgProcTmYldByBasePartAndOp = CDbl(avg90)
    Else
        AvgProcTmYldByBasePartAndOp = Null
    End If
    Exit Function
Fail:
    AvgProcTmYldByBasePartAndOp = Null
End Function

Private Function AverageHoursForMatch( _
    ByVal tableName As String, _
    ByVal assemblyColumnName As String, _
    ByVal opSequenceColumnName As String, _
    ByVal hoursColumnName As String, _
    ByVal basePart As String, _
    ByVal opSequence As String) As Variant

    Dim rs As DAO.Recordset
    Dim assemblyNo As String
    Dim rowBasePart As String
    Dim rowOpSequence As String
    Dim hoursValue As Double
    Dim totalHours As Double
    Dim matchCount As Long
    Dim sql As String

    If Not TableExists(tableName) Then
        AverageHoursForMatch = Null
        Exit Function
    End If

    sql = "SELECT [" & assemblyColumnName & "], [" & opSequenceColumnName & "], [" & hoursColumnName & "] FROM [" & tableName & "]"
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    totalHours = 0
    matchCount = 0
    Do Until rs.EOF
        assemblyNo = CoerceText(rs.Fields(assemblyColumnName).Value)
        rowOpSequence = CoerceText(rs.Fields(opSequenceColumnName).Value)
        If Len(assemblyNo) > 0 And ValuesMatch(rowOpSequence, opSequence) Then
            rowBasePart = GetBasePartNumber(assemblyNo)
            If ValuesMatch(rowBasePart, basePart) Then
                If TryGetNonZeroNumeric(rs.Fields(hoursColumnName).Value, hoursValue) Then
                    totalHours = totalHours + hoursValue
                    matchCount = matchCount + 1
                End If
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close

    If matchCount = 0 Then
        AverageHoursForMatch = Null
    Else
        AverageHoursForMatch = totalHours / matchCount
    End If
End Function
