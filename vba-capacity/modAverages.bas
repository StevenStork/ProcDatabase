Attribute VB_Name = "modAverages"
Option Explicit

'==============================================================================
' Average process hours and execution factor.
'
' Values are stored on the hidden PartAverages sheet (PartAveragesTbl) and
' rebuilt whenever linked OperComps / AssyStnd / TimeYield data is refreshed.
' PartEditor looks up that table instead of scanning the source queries per row.
'
' Avg Process Hours: average non-zero LABOR HPS (HOURS) in tblOperComps for
'   base part + oper seq; fallback to average non-zero RUN TIME in tblAssyStnd.
' Avg Ex: average non-zero Avg 180 Day Ex in tblTimeYield; fallback to Avg 90.
'==============================================================================

Private mLookupLoaded As Boolean
Private mAvgHours As Object
Private mAvgEx As Object

Public Function AvgProcessHoursByBasePartAndOp(ByVal basePartNumber As Variant, ByVal opSequence As Variant) As Variant
    Dim lookupKey As String

    On Error GoTo Fail

    lookupKey = AverageLookupKey(basePartNumber, opSequence)
    If Len(lookupKey) = 0 Then
        AvgProcessHoursByBasePartAndOp = Empty
        Exit Function
    End If

    EnsureAverageLookupReady
    If mAvgHours Is Nothing Then
        AvgProcessHoursByBasePartAndOp = Empty
        Exit Function
    End If
    If mAvgHours.Exists(lookupKey) Then
        AvgProcessHoursByBasePartAndOp = mAvgHours(lookupKey)
    Else
        AvgProcessHoursByBasePartAndOp = Empty
    End If
    Exit Function

Fail:
    AvgProcessHoursByBasePartAndOp = Empty
End Function

Public Function AvgExByBasePartAndOp(ByVal basePartNumber As Variant, ByVal opSequence As Variant) As Variant
    Dim lookupKey As String

    On Error GoTo Fail

    lookupKey = AverageLookupKey(basePartNumber, opSequence)
    If Len(lookupKey) = 0 Then
        AvgExByBasePartAndOp = Empty
        Exit Function
    End If

    EnsureAverageLookupReady
    If mAvgEx Is Nothing Then
        AvgExByBasePartAndOp = Empty
        Exit Function
    End If
    If mAvgEx.Exists(lookupKey) Then
        AvgExByBasePartAndOp = mAvgEx(lookupKey)
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

' Button-safe entry point. Rebuilds PartAveragesTbl from the linked source tables.
Public Sub RebuildPartAverages()
    Dim rowCount As Long
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo FailRebuild
    OptimizeExcel True
    rowCount = RebuildPartAveragesTable()
    OptimizeExcel False
    MsgBox "Part averages rebuilt (" & CStr(rowCount) & " rows) from tblOperComps, tblAssyStnd, and tblTimeYield.", vbInformation
    Exit Sub

FailRebuild:
    errNumber = Err.Number
    errDescription = Err.Description
    If Len(Trim$(errDescription)) = 0 Then errDescription = "(no description)"
    On Error Resume Next
    OptimizeExcel False
    On Error GoTo 0
    MsgBox "RebuildPartAverages failed:" & vbCrLf & "Error " & CStr(errNumber) & ": " & errDescription, vbExclamation
End Sub

' Rebuild hidden PartAveragesTbl. Returns the number of stored rows.
Public Function RebuildPartAveragesTable() As Long
    Dim hoursOperSum As Object
    Dim hoursOperCount As Object
    Dim hoursStndSum As Object
    Dim hoursStndCount As Object
    Dim ex180Sum As Object
    Dim ex180Count As Object
    Dim ex90Sum As Object
    Dim ex90Count As Object
    Dim allKeys As Object
    Dim lookupKey As Variant
    Dim rowCount As Long
    Dim values() As Variant
    Dim keyParts As Variant
    Dim avgHours As Variant
    Dim avgEx As Variant

    InvalidateAverageLookup
    EnsurePartAveragesInfrastructure

    Set hoursOperSum = NewLookupDictionary()
    Set hoursOperCount = NewLookupDictionary()
    Set hoursStndSum = NewLookupDictionary()
    Set hoursStndCount = NewLookupDictionary()
    Set ex180Sum = NewLookupDictionary()
    Set ex180Count = NewLookupDictionary()
    Set ex90Sum = NewLookupDictionary()
    Set ex90Count = NewLookupDictionary()
    Set allKeys = NewLookupDictionary()

    AccumulateAverages LINKED_OPER_COMPS_TABLE, COL_LABOR_HPS, hoursOperSum, hoursOperCount, allKeys
    AccumulateAverages LINKED_ASSY_STND_TABLE, COL_RUN_TIME, hoursStndSum, hoursStndCount, allKeys
    AccumulateAverages LINKED_TIME_YIELD_TABLE, COL_AVG_180_DAY_EX, ex180Sum, ex180Count, allKeys
    AccumulateAverages LINKED_TIME_YIELD_TABLE, COL_AVG_90_DAY_EX, ex90Sum, ex90Count, allKeys

    Set mAvgHours = NewLookupDictionary()
    Set mAvgEx = NewLookupDictionary()

    If allKeys.Count = 0 Then
        WritePartAveragesRows Empty, 0
        mLookupLoaded = True
        RebuildPartAveragesTable = 0
        Exit Function
    End If

    ReDim values(1 To allKeys.Count, 1 To 4)
    rowCount = 0

    For Each lookupKey In allKeys.Keys
        avgHours = Empty
        If hoursOperCount.Exists(lookupKey) Then
            If CLng(hoursOperCount(lookupKey)) > 0 Then
                avgHours = CDbl(hoursOperSum(lookupKey)) / CLng(hoursOperCount(lookupKey))
            End If
        End If
        If Not IsNumeric(avgHours) Then
            If hoursStndCount.Exists(lookupKey) Then
                If CLng(hoursStndCount(lookupKey)) > 0 Then
                    avgHours = CDbl(hoursStndSum(lookupKey)) / CLng(hoursStndCount(lookupKey))
                End If
            End If
        End If

        avgEx = Empty
        If ex180Count.Exists(lookupKey) Then
            If CLng(ex180Count(lookupKey)) > 0 Then
                avgEx = CDbl(ex180Sum(lookupKey)) / CLng(ex180Count(lookupKey))
            End If
        End If
        If Not IsNumeric(avgEx) Then
            If ex90Count.Exists(lookupKey) Then
                If CLng(ex90Count(lookupKey)) > 0 Then
                    avgEx = CDbl(ex90Sum(lookupKey)) / CLng(ex90Count(lookupKey))
                End If
            End If
        End If

        If Not IsNumeric(avgHours) And Not IsNumeric(avgEx) Then GoTo ContinueKey

        keyParts = Split(CStr(lookupKey), vbTab)
        If UBound(keyParts) < 1 Then GoTo ContinueKey

        rowCount = rowCount + 1
        values(rowCount, 1) = CStr(keyParts(0))
        values(rowCount, 2) = CStr(keyParts(1))
        If IsNumeric(avgHours) Then
            values(rowCount, 3) = CDbl(avgHours)
            mAvgHours(CStr(lookupKey)) = CDbl(avgHours)
        Else
            values(rowCount, 3) = Empty
        End If
        If IsNumeric(avgEx) Then
            values(rowCount, 4) = CDbl(avgEx)
            mAvgEx(CStr(lookupKey)) = CDbl(avgEx)
        Else
            values(rowCount, 4) = Empty
        End If

ContinueKey:
    Next lookupKey

    If rowCount = 0 Then
        WritePartAveragesRows Empty, 0
    Else
        If rowCount < allKeys.Count Then
            values = TrimValuesRows(values, rowCount, 4)
        End If
        WritePartAveragesRows values, rowCount
    End If

    mLookupLoaded = True
    RebuildPartAveragesTable = rowCount
End Function

Private Sub EnsureAverageLookupReady()
    If mLookupLoaded Then Exit Sub

    LoadAverageLookupFromTable
    If mAvgHours Is Nothing Then Set mAvgHours = NewLookupDictionary()
    If mAvgEx Is Nothing Then Set mAvgEx = NewLookupDictionary()

    If mAvgHours.Count = 0 And mAvgEx.Count = 0 Then
        RebuildPartAveragesTable
    End If

    mLookupLoaded = True
End Sub

Private Sub LoadAverageLookupFromTable()
    Dim tbl As ListObject
    Dim baseParts As Variant
    Dim operSeqs As Variant
    Dim hoursValues As Variant
    Dim exValues As Variant
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim lookupKey As String
    Dim numericValue As Double

    InvalidateAverageLookup
    Set mAvgHours = NewLookupDictionary()
    Set mAvgEx = NewLookupDictionary()

    Set tbl = FindTable(PART_AVERAGES_TABLE_NAME)
    If tbl Is Nothing Or tbl.DataBodyRange Is Nothing Then Exit Sub
    If Not TableHasColumn(tbl, COL_BASE_PART_CODE) Or Not TableHasColumn(tbl, COL_OPER_SEQ) Then Exit Sub

    baseParts = ListColumnValues2D(tbl, COL_BASE_PART_CODE)
    operSeqs = ListColumnValues2D(tbl, COL_OPER_SEQ)
    If Not IsArray(baseParts) Or Not IsArray(operSeqs) Then Exit Sub

    hoursValues = Empty
    exValues = Empty
    If TableHasColumn(tbl, COL_AVG_PROCESS_HOURS) Then
        hoursValues = ListColumnValues2D(tbl, COL_AVG_PROCESS_HOURS)
    End If
    If TableHasColumn(tbl, COL_AVG_EX) Then
        exValues = ListColumnValues2D(tbl, COL_AVG_EX)
    End If

    rowCount = UBound(baseParts, 1)
    For rowIndex = 1 To rowCount
        lookupKey = AverageLookupKey(baseParts(rowIndex, 1), operSeqs(rowIndex, 1))
        If Len(lookupKey) = 0 Then GoTo ContinueLoad

        If IsArray(hoursValues) Then
            If TryGetNonZeroNumeric(hoursValues(rowIndex, 1), numericValue) Then
                mAvgHours(lookupKey) = numericValue
            End If
        End If
        If IsArray(exValues) Then
            If TryGetNonZeroNumeric(exValues(rowIndex, 1), numericValue) Then
                mAvgEx(lookupKey) = numericValue
            End If
        End If

ContinueLoad:
    Next rowIndex
End Sub

Private Sub InvalidateAverageLookup()
    mLookupLoaded = False
    Set mAvgHours = Nothing
    Set mAvgEx = Nothing
End Sub

Private Sub AccumulateAverages( _
    ByVal tableName As String, _
    ByVal valueColumnName As String, _
    ByVal sums As Object, _
    ByVal counts As Object, _
    ByVal allKeys As Object)

    Dim tbl As ListObject
    Dim assemblyValues As Variant
    Dim opSequenceValues As Variant
    Dim metricValues As Variant
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim assemblyNo As String
    Dim rowBasePart As String
    Dim dashCondition As String
    Dim lookupKey As String
    Dim metricValue As Double

    Set tbl = FindTable(tableName)
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    If Not TableHasColumn(tbl, COL_ASSEMBLY_NO) Then Exit Sub
    If Not TableHasColumn(tbl, COL_OPER_SEQ_SOURCE) Then Exit Sub
    If Not TableHasColumn(tbl, valueColumnName) Then Exit Sub

    assemblyValues = ListColumnValues2D(tbl, COL_ASSEMBLY_NO)
    opSequenceValues = ListColumnValues2D(tbl, COL_OPER_SEQ_SOURCE)
    metricValues = ListColumnValues2D(tbl, valueColumnName)
    If Not IsArray(assemblyValues) Or Not IsArray(opSequenceValues) Or Not IsArray(metricValues) Then Exit Sub

    rowCount = UBound(assemblyValues, 1)
    For rowIndex = 1 To rowCount
        assemblyNo = Trim$(CStr(Nz(assemblyValues(rowIndex, 1))))
        If Len(assemblyNo) = 0 Then GoTo ContinueRow
        If Not TryGetNonZeroNumeric(metricValues(rowIndex, 1), metricValue) Then GoTo ContinueRow

        SplitAssemblyNo assemblyNo, rowBasePart, dashCondition
        rowBasePart = NormalizeCode(rowBasePart)
        lookupKey = AverageLookupKey(rowBasePart, opSequenceValues(rowIndex, 1))
        If Len(lookupKey) = 0 Then GoTo ContinueRow

        If Not allKeys.Exists(lookupKey) Then allKeys.Add lookupKey, True
        If sums.Exists(lookupKey) Then
            sums(lookupKey) = CDbl(sums(lookupKey)) + metricValue
            counts(lookupKey) = CLng(counts(lookupKey)) + 1
        Else
            sums.Add lookupKey, metricValue
            counts.Add lookupKey, 1
        End If

ContinueRow:
    Next rowIndex
End Sub

Private Sub WritePartAveragesRows(ByVal values As Variant, ByVal rowCount As Long)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lastRow As Long
    Dim tableRange As Range

    EnsurePartAveragesInfrastructure
    Set ws = FindWorksheetByName(PART_AVERAGES_SHEET_NAME)
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Visible = xlSheetHidden
    Set tbl = ws.ListObjects(PART_AVERAGES_TABLE_NAME)
    If Not tbl Is Nothing Then tbl.Delete
    Set tbl = Nothing
    On Error GoTo 0

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < TABLE_HEADER_ROW Then lastRow = TABLE_HEADER_ROW
    ws.Range(ws.Cells(TABLE_HEADER_ROW, 1), ws.Cells(lastRow, 4)).Clear

    ws.Cells(TABLE_HEADER_ROW, 1).Value = COL_BASE_PART_CODE
    ws.Cells(TABLE_HEADER_ROW, 2).Value = COL_OPER_SEQ
    ws.Cells(TABLE_HEADER_ROW, 3).Value = COL_AVG_PROCESS_HOURS
    ws.Cells(TABLE_HEADER_ROW, 4).Value = COL_AVG_EX
    ws.Rows(TABLE_HEADER_ROW).Font.Bold = True

    If rowCount > 0 And IsArray(values) Then
        ws.Cells(TABLE_FIRST_DATA_ROW, 1).Resize(rowCount, 4).Value2 = values
        ws.Range( _
            ws.Cells(TABLE_FIRST_DATA_ROW, 1), _
            ws.Cells(TABLE_FIRST_DATA_ROW + rowCount - 1, 2)).NumberFormat = "@"
        ws.Range( _
            ws.Cells(TABLE_FIRST_DATA_ROW, 3), _
            ws.Cells(TABLE_FIRST_DATA_ROW + rowCount - 1, 4)).NumberFormat = "0.####"
        Set tableRange = ws.Range( _
            ws.Cells(TABLE_HEADER_ROW, 1), _
            ws.Cells(TABLE_FIRST_DATA_ROW + rowCount - 1, 4))
    Else
        Set tableRange = ws.Range(ws.Cells(TABLE_HEADER_ROW, 1), ws.Cells(TABLE_HEADER_ROW, 4))
    End If

    Set tbl = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    tbl.Name = PART_AVERAGES_TABLE_NAME

    ws.Range("A2").Value = "Last rebuilt " & Format$(Now, "yyyy-mm-dd hh:nn") & _
        " from tblOperComps, tblAssyStnd, and tblTimeYield."
    ws.Visible = xlSheetVeryHidden
End Sub

Private Function TrimValuesRows(ByVal values As Variant, ByVal rowCount As Long, ByVal columnCount As Long) As Variant
    Dim trimmed() As Variant
    Dim rowIndex As Long
    Dim colIndex As Long

    ReDim trimmed(1 To rowCount, 1 To columnCount)
    For rowIndex = 1 To rowCount
        For colIndex = 1 To columnCount
            trimmed(rowIndex, colIndex) = values(rowIndex, colIndex)
        Next colIndex
    Next rowIndex
    TrimValuesRows = trimmed
End Function

Private Function AverageLookupKey(ByVal basePartNumber As Variant, ByVal opSequence As Variant) As String
    Dim basePart As String
    Dim opSeq As String

    basePart = NormalizeCode(basePartNumber)
    opSeq = NormalizeOperSeqKey(opSequence)
    If Len(basePart) = 0 Or Len(opSeq) = 0 Then
        AverageLookupKey = vbNullString
    Else
        AverageLookupKey = basePart & vbTab & opSeq
    End If
End Function

Private Function NewLookupDictionary() As Object
    Set NewLookupDictionary = CreateObject("Scripting.Dictionary")
    NewLookupDictionary.CompareMode = vbTextCompare
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
