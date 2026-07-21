Attribute VB_Name = "modAvgLaborHours"
Option Explicit

'==============================================================================
' AvgLaborHoursByBasePartAndOp
'
' Worksheet usage:
'   =AvgLaborHoursByBasePartAndOp(A2, B2)
'
' Looks up the average LABOR HPS (HOURS) in the Operation Completions table for
' the given base part number and op sequence. If that average does not exist or
' is 0, falls back to the average RUN TIME (HOURS) in AssyStndTbl.
'==============================================================================

Private Const OP_COMPLETIONS_TABLE_NAME As String = "OperationCompletions"
Private Const ASSY_STANDARDS_TABLE_NAME As String = "AssyStndTbl"

Private Const COL_ASSEMBLY_NO As String = "ASSEMBLY NO"
Private Const COL_OP_SEQUENCE As String = "OP SEQUENCE"
Private Const COL_LABOR_HPS As String = "LABOR HPS (HOURS)"
Private Const COL_RUN_TIME As String = "RUN TIME (HOURS)"

Public Function AvgLaborHoursByBasePartAndOp(ByVal BasePartNumber As Variant, ByVal OpSequence As Variant) As Variant
    Dim basePart As String
    Dim opSeq As String
    Dim opCompletionsAvg As Variant
    Dim standardsAvg As Variant

    On Error GoTo Fail

    basePart = Trim$(CStr(BasePartNumber))
    opSeq = Trim$(CStr(OpSequence))

    If Len(basePart) = 0 Or Len(opSeq) = 0 Then
        AvgLaborHoursByBasePartAndOp = CVErr(xlErrNA)
        Exit Function
    End If

    opCompletionsAvg = AverageHoursForMatch( _
        OP_COMPLETIONS_TABLE_NAME, _
        COL_ASSEMBLY_NO, _
        COL_OP_SEQUENCE, _
        COL_LABOR_HPS, _
        basePart, _
        opSeq)

    If IsNumeric(opCompletionsAvg) Then
        If CDbl(opCompletionsAvg) <> 0 Then
            AvgLaborHoursByBasePartAndOp = CDbl(opCompletionsAvg)
            Exit Function
        End If
    End If

    standardsAvg = AverageHoursForMatch( _
        ASSY_STANDARDS_TABLE_NAME, _
        COL_ASSEMBLY_NO, _
        COL_OP_SEQUENCE, _
        COL_RUN_TIME, _
        basePart, _
        opSeq)

    If IsNumeric(standardsAvg) Then
        AvgLaborHoursByBasePartAndOp = CDbl(standardsAvg)
    Else
        AvgLaborHoursByBasePartAndOp = CVErr(xlErrNA)
    End If

    Exit Function

Fail:
    AvgLaborHoursByBasePartAndOp = CVErr(xlErrValue)
End Function

Private Function AverageHoursForMatch( _
    ByVal tableName As String, _
    ByVal assemblyColumnName As String, _
    ByVal opSequenceColumnName As String, _
    ByVal hoursColumnName As String, _
    ByVal basePart As String, _
    ByVal opSequence As String) As Variant

    Dim tbl As ListObject
    Dim assemblyValues As Variant
    Dim opSequenceValues As Variant
    Dim hoursValues As Variant
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim assemblyNo As String
    Dim rowBasePart As String
    Dim rowOpSequence As String
    Dim hoursValue As Double
    Dim totalHours As Double
    Dim matchCount As Long

    Set tbl = FindListObjectByName(tableName)
    If tbl Is Nothing Then
        AverageHoursForMatch = CVErr(xlErrNA)
        Exit Function
    End If

    If tbl.DataBodyRange Is Nothing Then
        AverageHoursForMatch = CVErr(xlErrNA)
        Exit Function
    End If

    If Not TableHasColumn(tbl, assemblyColumnName) _
        Or Not TableHasColumn(tbl, opSequenceColumnName) _
        Or Not TableHasColumn(tbl, hoursColumnName) Then
        AverageHoursForMatch = CVErr(xlErrName)
        Exit Function
    End If

    assemblyValues = ColumnValues(tbl.ListColumns(assemblyColumnName))
    opSequenceValues = ColumnValues(tbl.ListColumns(opSequenceColumnName))
    hoursValues = ColumnValues(tbl.ListColumns(hoursColumnName))
    rowCount = UBound(assemblyValues, 1)

    totalHours = 0
    matchCount = 0

    For rowIndex = 1 To rowCount
        assemblyNo = Trim$(CStr(Nz(assemblyValues(rowIndex, 1))))
        rowOpSequence = Trim$(CStr(Nz(opSequenceValues(rowIndex, 1))))

        If Len(assemblyNo) > 0 And ValuesMatch(rowOpSequence, opSequence) Then
            rowBasePart = GetBasePartNumber(assemblyNo)

            If ValuesMatch(rowBasePart, basePart) Then
                If TryGetNumericHours(hoursValues(rowIndex, 1), hoursValue) Then
                    totalHours = totalHours + hoursValue
                    matchCount = matchCount + 1
                End If
            End If
        End If
    Next rowIndex

    If matchCount = 0 Then
        AverageHoursForMatch = CVErr(xlErrNA)
    Else
        AverageHoursForMatch = totalHours / matchCount
    End If
End Function

Private Function ColumnValues(ByVal col As ListColumn) As Variant
    Dim values As Variant
    Dim result(1 To 1, 1 To 1) As Variant

    values = col.DataBodyRange.Value2

    If IsArray(values) Then
        ColumnValues = values
    Else
        result(1, 1) = values
        ColumnValues = result
    End If
End Function

Private Function FindListObjectByName(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim tbl As ListObject

    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set tbl = ws.ListObjects(tableName)
        On Error GoTo 0

        If Not tbl Is Nothing Then
            Set FindListObjectByName = tbl
            Exit Function
        End If
    Next ws
End Function

Private Function TableHasColumn(ByVal tbl As ListObject, ByVal columnName As String) As Boolean
    Dim col As ListColumn

    On Error Resume Next
    Set col = tbl.ListColumns(columnName)
    On Error GoTo 0

    TableHasColumn = Not col Is Nothing
End Function

Private Function GetBasePartNumber(ByVal assemblyNo As String) As String
    Dim dashPos As Long

    dashPos = InStr(1, assemblyNo, "-", vbBinaryCompare)

    If dashPos > 0 Then
        GetBasePartNumber = Trim$(Left$(assemblyNo, dashPos - 1))
    Else
        GetBasePartNumber = Trim$(assemblyNo)
    End If
End Function

Private Function ValuesMatch(ByVal leftValue As String, ByVal rightValue As String) As Boolean
    If IsNumeric(leftValue) And IsNumeric(rightValue) Then
        ValuesMatch = (CDbl(leftValue) = CDbl(rightValue))
    Else
        ValuesMatch = (StrComp(leftValue, rightValue, vbTextCompare) = 0)
    End If
End Function

Private Function TryGetNumericHours(ByVal rawValue As Variant, ByRef hoursValue As Double) As Boolean
    If IsError(rawValue) Then Exit Function
    If IsEmpty(rawValue) Then Exit Function
    If Len(Trim$(CStr(rawValue))) = 0 Then Exit Function
    If Not IsNumeric(rawValue) Then Exit Function

    hoursValue = CDbl(rawValue)
    TryGetNumericHours = True
End Function

Private Function Nz(ByVal value As Variant) As Variant
    If IsError(value) Then
        Nz = vbNullString
    ElseIf IsEmpty(value) Or IsNull(value) Then
        Nz = vbNullString
    Else
        Nz = value
    End If
End Function
