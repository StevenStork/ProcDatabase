Attribute VB_Name = "modAverages"
Option Compare Database
Option Explicit

' Ports AvgLaborHoursByBasePartAndOp and AvgProcTmYldByBasePartAndOp.
' Zero / null / blank values are excluded. Labor falls back to standards
' run time. Process-time executions prefer 180-day then 90-day.
'
' Seed Ops calls BeginAverageCache once so each source table is scanned
' a single time instead of once per operation row.

Private gLaborCache As Object
Private gExCache As Object
Private gCacheActive As Boolean
Private gProcTmMissing As Boolean

Public Sub BeginAverageCache()
    Set gLaborCache = Nothing
    Set gExCache = Nothing
    gCacheActive = True
    gProcTmMissing = Not TableExists(TBL_PROC_TM_YLD)
End Sub

Public Sub EndAverageCache()
    Set gLaborCache = Nothing
    Set gExCache = Nothing
    gCacheActive = False
End Sub

Public Function AverageCacheIsActive() As Boolean
    AverageCacheIsActive = gCacheActive
End Function

Public Function AvgLaborHoursByBasePartAndOp(ByVal BasePartNumber As Variant, ByVal OpSequence As Variant) As Variant
    Dim basePart As String
    Dim opSeq As String
    Dim key As String
    Dim completionsAvg As Variant
    Dim standardsAvg As Variant

    On Error GoTo Fail
    basePart = Trim$(CStr(Nz(BasePartNumber, vbNullString)))
    opSeq = Trim$(CStr(Nz(OpSequence, vbNullString)))
    If Len(basePart) = 0 Or Len(opSeq) = 0 Then
        AvgLaborHoursByBasePartAndOp = Null
        Exit Function
    End If

    If gCacheActive Then
        EnsureLaborCache
        key = CacheKey(basePart, opSeq)
        If gLaborCache.Exists(key) Then
            AvgLaborHoursByBasePartAndOp = gLaborCache(key)
        Else
            AvgLaborHoursByBasePartAndOp = Null
        End If
        Exit Function
    End If

    completionsAvg = AverageHoursForMatch(OperCompsSourceName(), COL_ASSEMBLY_NO, COL_OPER_SEQ, COL_LABOR_HPS, basePart, opSeq)
    If IsNumeric(completionsAvg) Then
        If CDbl(completionsAvg) <> 0 Then
            AvgLaborHoursByBasePartAndOp = CDbl(completionsAvg)
            Exit Function
        End If
    End If

    standardsAvg = AverageHoursForMatch(AssyStndSourceName(), COL_ASSEMBLY_NO, COL_OPER_SEQ, COL_RUN_TIME, basePart, opSeq)
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
    Dim key As String
    Dim avg180 As Variant
    Dim avg90 As Variant

    On Error GoTo Fail
    basePart = Trim$(CStr(Nz(BasePartNumber, vbNullString)))
    opSeq = Trim$(CStr(Nz(OpSequence, vbNullString)))
    If Len(basePart) = 0 Or Len(opSeq) = 0 Then
        AvgProcTmYldByBasePartAndOp = Null
        Exit Function
    End If

    If gCacheActive Then
        If gProcTmMissing Then
            AvgProcTmYldByBasePartAndOp = Null
            Exit Function
        End If
        EnsureExCache
        key = CacheKey(basePart, opSeq)
        If gExCache.Exists(key) Then
            AvgProcTmYldByBasePartAndOp = gExCache(key)
        Else
            AvgProcTmYldByBasePartAndOp = Null
        End If
        Exit Function
    End If

    If Not TableExists(TBL_PROC_TM_YLD) Then
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

Private Sub EnsureLaborCache()
    If Not gLaborCache Is Nothing Then Exit Sub
    Set gLaborCache = CreateObject("Scripting.Dictionary")
    gLaborCache.CompareMode = vbTextCompare
    ' Completions win; standards fill gaps only.
    AccumulateHoursCache OperCompsSourceName(), COL_ASSEMBLY_NO, COL_OPER_SEQ, COL_LABOR_HPS, gLaborCache, False
    AccumulateHoursCache AssyStndSourceName(), COL_ASSEMBLY_NO, COL_OPER_SEQ, COL_RUN_TIME, gLaborCache, True
End Sub

Private Sub EnsureExCache()
    If Not gExCache Is Nothing Then Exit Sub
    Set gExCache = CreateObject("Scripting.Dictionary")
    gExCache.CompareMode = vbTextCompare
    If gProcTmMissing Then Exit Sub
    ' 180-day wins; 90-day fills gaps only.
    AccumulateHoursCache TBL_PROC_TM_YLD, COL_ASSEMBLY_NO_ALT, COL_OPER_SEQ, COL_AVG_180, gExCache, False
    AccumulateHoursCache TBL_PROC_TM_YLD, COL_ASSEMBLY_NO_ALT, COL_OPER_SEQ, COL_AVG_90, gExCache, True
End Sub

' One pass over a source: group non-zero hours by base part + op sequence.
' fillGapsOnly:=True skips keys already present (fallback chain).
Private Sub AccumulateHoursCache( _
    ByVal tableName As String, _
    ByVal assemblyColumnName As String, _
    ByVal opSequenceColumnName As String, _
    ByVal hoursColumnName As String, _
    ByVal cache As Object, _
    ByVal fillGapsOnly As Boolean)

    Dim rs As DAO.Recordset
    Dim sums As Object
    Dim counts As Object
    Dim assemblyNo As String
    Dim rowBasePart As String
    Dim rowOpSequence As String
    Dim hoursValue As Double
    Dim key As String
    Dim k As Variant

    If Not TableExists(tableName) And Not QueryExists(tableName) Then Exit Sub

    Set sums = CreateObject("Scripting.Dictionary")
    sums.CompareMode = vbTextCompare
    Set counts = CreateObject("Scripting.Dictionary")
    counts.CompareMode = vbTextCompare

    Set rs = CurrentDb.OpenRecordset( _
        "SELECT [" & assemblyColumnName & "], [" & opSequenceColumnName & "], [" & hoursColumnName & "] " & _
        "FROM [" & tableName & "]", dbOpenSnapshot)
    Do Until rs.EOF
        assemblyNo = CoerceText(rs.Fields(assemblyColumnName).Value)
        rowOpSequence = CoerceText(rs.Fields(opSequenceColumnName).Value)
        If Len(assemblyNo) > 0 And Len(rowOpSequence) > 0 Then
            If TryGetNonZeroNumeric(rs.Fields(hoursColumnName).Value, hoursValue) Then
                rowBasePart = GetBasePartNumber(assemblyNo)
                If Len(rowBasePart) > 0 Then
                    key = CacheKey(rowBasePart, rowOpSequence)
                    If Not (fillGapsOnly And cache.Exists(key)) Then
                        If sums.Exists(key) Then
                            sums(key) = CDbl(sums(key)) + hoursValue
                            counts(key) = CLng(counts(key)) + 1
                        Else
                            sums.Add key, hoursValue
                            counts.Add key, 1&
                        End If
                    End If
                End If
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close

    For Each k In sums.Keys
        If Not (fillGapsOnly And cache.Exists(CStr(k))) Then
            If cache.Exists(CStr(k)) Then
                cache(CStr(k)) = CDbl(sums(k)) / CLng(counts(k))
            Else
                cache.Add CStr(k), CDbl(sums(k)) / CLng(counts(k))
            End If
        End If
    Next k
End Sub

Private Function CacheKey(ByVal basePart As String, ByVal opSequence As String) As String
    ' Normalize numeric op sequences so "10" and 10 share a key.
    If IsNumeric(opSequence) Then
        CacheKey = basePart & vbTab & CStr(CLng(CDbl(opSequence)))
    Else
        CacheKey = basePart & vbTab & opSequence
    End If
End Function

' Single-lookup path (no batch cache): restrict rows by base-part prefix.
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

    If Not TableExists(tableName) And Not QueryExists(tableName) Then
        AverageHoursForMatch = Null
        Exit Function
    End If

    sql = "SELECT [" & assemblyColumnName & "], [" & opSequenceColumnName & "], [" & hoursColumnName & "] " & _
        "FROM [" & tableName & "] WHERE [" & assemblyColumnName & "] Like " & SqlText(basePart & "-*") & _
        " OR [" & assemblyColumnName & "] = " & SqlText(basePart)
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
