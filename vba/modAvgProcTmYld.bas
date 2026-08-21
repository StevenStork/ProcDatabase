Attribute VB_Name = "modAvgProcTmYld"
Option Explicit

'==============================================================================
' AvgProcTmYldByBasePartAndOp
'
' Worksheet usage:
'   =AvgProcTmYldByBasePartAndOp(A2, B2)
'
' Averages Avg 180 Day Ex from ProcTmYldTbl for the given base part number and
' OPER SEQ. Null/blank/zero values are excluded. If no usable 180-day values
' exist, averages Avg 90 Day Ex with the same matching rules.
'==============================================================================

Private Const PROC_TM_YLD_TABLE_NAME As String = "ProcTmYldTbl"
Private Const COL_ASSEMBLY_NO As String = "Assembly No"
Private Const COL_OP_SEQUENCE As String = "OPER SEQ"
Private Const COL_AVG_180_DAY_EX As String = "Avg 180 Day Ex"
Private Const COL_AVG_90_DAY_EX As String = "Avg 90 Day Ex"

Public Function AvgProcTmYldByBasePartAndOp(ByVal BasePartNumber As Variant, ByVal OpSequence As Variant) As Variant
    Dim basePart As String
    Dim opSeq As String
    Dim avg180 As Variant
    Dim avg90 As Variant

    On Error GoTo Fail

    basePart = Trim$(CStr(BasePartNumber))
    opSeq = Trim$(CStr(OpSequence))

    If Len(basePart) = 0 Or Len(opSeq) = 0 Then
        AvgProcTmYldByBasePartAndOp = CVErr(xlErrNA)
        Exit Function
    End If

    avg180 = AverageColumnForMatch(COL_AVG_180_DAY_EX, basePart, opSeq)

    If IsNumeric(avg180) Then
        If CDbl(avg180) <> 0 Then
            AvgProcTmYldByBasePartAndOp = CDbl(avg180)
            Exit Function
        End If
    End If

    avg90 = AverageColumnForMatch(COL_AVG_90_DAY_EX, basePart, opSeq)

    If IsNumeric(avg90) Then
        AvgProcTmYldByBasePartAndOp = CDbl(avg90)
    Else
        AvgProcTmYldByBasePartAndOp = CVErr(xlErrNA)
    End If

    Exit Function

Fail:
    AvgProcTmYldByBasePartAndOp = CVErr(xlErrValue)
End Function

Private Function AverageColumnForMatch( _
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

    Set tbl = FindListObjectByName(PROC_TM_YLD_TABLE_NAME)
    If tbl Is Nothing Then
        AverageColumnForMatch = CVErr(xlErrNA)
        Exit Function
    End If

    If tbl.DataBodyRange Is Nothing Then
        AverageColumnForMatch = CVErr(xlErrNA)
        Exit Function
    End If

    If Not TableHasColumn(tbl, COL_ASSEMBLY_NO) _
        Or Not TableHasColumn(tbl, COL_OP_SEQUENCE) _
        Or Not TableHasColumn(tbl, hoursColumnName) Then
        AverageColumnForMatch = CVErr(xlErrName)
        Exit Function
    End If

    assemblyValues = ColumnValues(tbl.ListColumns(COL_ASSEMBLY_NO))
    opSequenceValues = ColumnValues(tbl.ListColumns(COL_OP_SEQUENCE))
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
                If TryGetNonZeroNumeric(hoursValues(rowIndex, 1), hoursValue) Then
                    totalHours = totalHours + hoursValue
                    matchCount = matchCount + 1
                End If
            End If
        End If
    Next rowIndex

    If matchCount = 0 Then
        AverageColumnForMatch = CVErr(xlErrNA)
    Else
        AverageColumnForMatch = totalHours / matchCount
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

Private Function TryGetNonZeroNumeric(ByVal rawValue As Variant, ByRef numericValue As Double) As Boolean
    If IsError(rawValue) Then Exit Function
    If IsEmpty(rawValue) Then Exit Function
    If IsNull(rawValue) Then Exit Function
    If Len(Trim$(CStr(rawValue))) = 0 Then Exit Function
    If Not IsNumeric(rawValue) Then Exit Function

    numericValue = CDbl(rawValue)
    If numericValue = 0 Then Exit Function

    TryGetNonZeroNumeric = True
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
