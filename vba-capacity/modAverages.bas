Attribute VB_Name = "modAverages"
Option Explicit

'==============================================================================
' Average process hours and execution factor from linked query tables.
'
' Avg Process Hours: average LABOR HPS (HOURS) in tblOperComps for the base
'   part and oper seq (zero values excluded). Falls back to average non-zero
'   RUN TIME (HOURS) in tblAssyStnd.
'
' Avg Ex: average Avg 180 Day Ex in tblTimeYield for the base part and oper
'   seq (zero values excluded). Falls back to Avg 90 Day Ex.
'==============================================================================

Public Function AvgProcessHoursByBasePartAndOp(ByVal basePartNumber As Variant, ByVal opSequence As Variant) As Variant
    Dim basePart As String
    Dim opSeq As String
    Dim operCompsAvg As Variant
    Dim standardsAvg As Variant

    On Error GoTo Fail

    basePart = Trim$(CStr(basePartNumber))
    opSeq = Trim$(CStr(opSequence))

    If Len(basePart) = 0 Or Len(opSeq) = 0 Then
        AvgProcessHoursByBasePartAndOp = Empty
        Exit Function
    End If

    operCompsAvg = AverageNumericForMatch( _
        LINKED_OPER_COMPS_TABLE, _
        COL_LABOR_HPS, _
        basePart, _
        opSeq)

    If IsNumeric(operCompsAvg) Then
        If CDbl(operCompsAvg) <> 0 Then
            AvgProcessHoursByBasePartAndOp = CDbl(operCompsAvg)
            Exit Function
        End If
    End If

    standardsAvg = AverageNumericForMatch( _
        LINKED_ASSY_STND_TABLE, _
        COL_RUN_TIME, _
        basePart, _
        opSeq)

    If IsNumeric(standardsAvg) Then
        AvgProcessHoursByBasePartAndOp = CDbl(standardsAvg)
    Else
        AvgProcessHoursByBasePartAndOp = Empty
    End If

    Exit Function

Fail:
    AvgProcessHoursByBasePartAndOp = Empty
End Function

Public Function AvgExByBasePartAndOp(ByVal basePartNumber As Variant, ByVal opSequence As Variant) As Variant
    Dim basePart As String
    Dim opSeq As String
    Dim avg180 As Variant
    Dim avg90 As Variant

    On Error GoTo Fail

    basePart = Trim$(CStr(basePartNumber))
    opSeq = Trim$(CStr(opSequence))

    If Len(basePart) = 0 Or Len(opSeq) = 0 Then
        AvgExByBasePartAndOp = Empty
        Exit Function
    End If

    avg180 = AverageNumericForMatch( _
        LINKED_TIME_YIELD_TABLE, _
        COL_AVG_180_DAY_EX, _
        basePart, _
        opSeq)

    If IsNumeric(avg180) Then
        If CDbl(avg180) <> 0 Then
            AvgExByBasePartAndOp = CDbl(avg180)
            Exit Function
        End If
    End If

    avg90 = AverageNumericForMatch( _
        LINKED_TIME_YIELD_TABLE, _
        COL_AVG_90_DAY_EX, _
        basePart, _
        opSeq)

    If IsNumeric(avg90) Then
        AvgExByBasePartAndOp = CDbl(avg90)
    Else
        AvgExByBasePartAndOp = Empty
    End If

    Exit Function

Fail:
    AvgExByBasePartAndOp = Empty
End Function

Public Function FormatAverageDisplay(ByVal avgValue As Variant) As String
    If IsEmpty(avgValue) Then
        FormatAverageDisplay = vbNullString
    ElseIf IsNumeric(avgValue) Then
        FormatAverageDisplay = Format$(CDbl(avgValue), "0.####")
    Else
        FormatAverageDisplay = vbNullString
    End If
End Function

Private Function AverageNumericForMatch( _
    ByVal tableName As String, _
    ByVal valueColumnName As String, _
    ByVal basePart As String, _
    ByVal opSequence As String) As Variant

    Dim tbl As ListObject
    Dim assemblyValues As Variant
    Dim opSequenceValues As Variant
    Dim metricValues As Variant
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim assemblyNo As String
    Dim rowBasePart As String
    Dim rowOpSequence As String
    Dim metricValue As Double
    Dim totalValue As Double
    Dim matchCount As Long

    Set tbl = FindTable(tableName)
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    If Not TableHasColumn(tbl, COL_ASSEMBLY_NO) _
        Or Not TableHasColumn(tbl, COL_OPER_SEQ_SOURCE) _
        Or Not TableHasColumn(tbl, valueColumnName) Then
        Exit Function
    End If

    assemblyValues = ListColumnValues2D(tbl, COL_ASSEMBLY_NO)
    opSequenceValues = ListColumnValues2D(tbl, COL_OPER_SEQ_SOURCE)
    metricValues = ListColumnValues2D(tbl, valueColumnName)
    If Not IsArray(assemblyValues) Or Not IsArray(opSequenceValues) Or Not IsArray(metricValues) Then Exit Function
    rowCount = UBound(assemblyValues, 1)

    For rowIndex = 1 To rowCount
        assemblyNo = Trim$(CStr(Nz(assemblyValues(rowIndex, 1))))
        rowOpSequence = Trim$(CStr(Nz(opSequenceValues(rowIndex, 1))))

        If Len(assemblyNo) > 0 And OpSequencesMatch(rowOpSequence, opSequence) Then
            rowBasePart = BasePartFromAssemblyNo(assemblyNo)

            If ValuesMatchCode(rowBasePart, basePart) Then
                If TryGetNonZeroNumeric(metricValues(rowIndex, 1), metricValue) Then
                    totalValue = totalValue + metricValue
                    matchCount = matchCount + 1
                End If
            End If
        End If
    Next rowIndex

    If matchCount > 0 Then
        AverageNumericForMatch = totalValue / matchCount
    End If
End Function

Private Function BasePartFromAssemblyNo(ByVal assemblyNo As String) As String
    SplitAssemblyNo assemblyNo, BasePartFromAssemblyNo, vbNullString
End Function

Private Function TryGetNonZeroNumeric(ByVal rawValue As Variant, ByRef numericValue As Double) As Boolean
    If IsError(rawValue) Then Exit Function
    If IsEmpty(rawValue) Or IsNull(rawValue) Then Exit Function
    If Len(Trim$(CStr(rawValue))) = 0 Then Exit Function
    If Not IsNumeric(rawValue) Then Exit Function

    numericValue = CDbl(rawValue)
    If numericValue = 0 Then Exit Function

    TryGetNonZeroNumeric = True
End Function
