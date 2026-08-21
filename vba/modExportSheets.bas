Attribute VB_Name = "modExportSheets"
Option Explicit

' Builds one export sheet per FFA (References!B) and per product line
' (References!D). Export sheets are plain data for copying to another file —
' no formatting is applied.
'
' Column A holds export metadata (marker / type / key).
' Data table starts at column C:
'   Part Number | Op Sequence | Op Code | Process Hours | Avg Ex |
'   Batch Size  | Avg HPU     | Equipment Type
'
' Part Number is each Part sheet's base part (C2), with any dash condition
' stripped. Remaining fields come from that sheet's data table
' (M, N, W, X, Q, Y, T). FFA sheets only include rows whose column Z matches
' the FFA. Product-line sheets include every data row from Part sheets that
' have that product line checked (G/H).
'
' Speed: Part A2 (activate cache) is not enough — it ignores Z marks and
' checkbox values. Part A3 stores an export source stamp (hidden ;;;). A
' workbook-level build cache skips rewriting export sheets when every Part
' A3 cheap-check still matches and References are unchanged.

Private Const PART_LABEL_CELL As String = "A1"
Private Const PART_LABEL_VALUE As String = "Part"
Private Const PART_NUMBER_CELL As String = "C2"
Private Const PART_EXPORT_CACHE_CELL As String = "A3"
Private Const PART_EXPORT_CACHE_SCHEMA As String = "1"

Private Const REFERENCES_SHEET_NAME As String = "References"
Private Const REFERENCES_FFA_COLUMN As String = "B"
Private Const REFERENCES_PRODUCT_LINE_COLUMN As String = "D"
Private Const REFERENCES_START_ROW As Long = 2

Private Const LIST_START_ROW As Long = 9
Private Const PRODUCT_LINE_VALUE_COLUMN As String = "G"
Private Const PRODUCT_LINE_CHECKBOX_COLUMN As String = "H"

Private Const DATA_FIRST_COLUMN As String = "M"
Private Const DATA_LAST_COLUMN As String = "Z"
' Offsets within an M:Z Value2 dump (M = 1).
Private Const OFF_OP_SEQUENCE As Long = 1          ' M
Private Const OFF_OP_CODE As Long = 2               ' N
Private Const OFF_BATCH_SIZE As Long = 5            ' Q
Private Const OFF_EQUIPMENT_TYPE As Long = 8        ' T
Private Const OFF_PROCESS_HOURS As Long = 11        ' W
Private Const OFF_AVG_EX As Long = 12               ' X
Private Const OFF_AVG_HPU As Long = 13              ' Y
Private Const OFF_FFA_MARK As Long = 14             ' Z

Private Const EXPORT_MARKER_CELL As String = "A1"
Private Const EXPORT_MARKER_VALUE As String = "Export"
Private Const EXPORT_TYPE_CELL As String = "A2"
Private Const EXPORT_KEY_CELL As String = "A3"
Private Const EXPORT_TYPE_FFA As String = "FFA"
Private Const EXPORT_TYPE_PRODUCT_LINE As String = "ProductLine"
Private Const LEGACY_EXPORT_MARKER_CELL As String = "AA1"
Private Const LEGACY_EXPORT_TYPE_CELL As String = "AA2"
Private Const LEGACY_EXPORT_KEY_CELL As String = "AA3"

Private Const FFA_SHEET_PREFIX As String = "FFA - "
Private Const PRODUCT_LINE_SHEET_PREFIX As String = "PL - "
Private Const EXPORT_HEADER_ROW As Long = 1
Private Const EXPORT_FIRST_DATA_ROW As Long = 2
Private Const EXPORT_FIRST_COLUMN As Long = 3   ' Column C
Private Const EXPORT_LAST_COLUMN As Long = 10  ' Column J
Private Const EXPORT_COLUMN_COUNT As Long = 8

Private Const EXPORT_BUILD_CACHE_NAME As String = "ProcDb_ExportBuildCache"
Private Const EXPORT_BUILD_CACHE_SCHEMA As String = "1"

Public Sub BuildExportSheets()
    Dim ffaValues() As String
    Dim productLines() As String
    Dim ffaRows As Object
    Dim productLineRows As Object
    Dim buildKey As String
    Dim i As Long
    Dim keyName As String

    On Error GoTo CleanUp
    OptimizeExcel True

    ffaValues = GetReferenceColumnValues(REFERENCES_FFA_COLUMN)
    productLines = GetReferenceColumnValues(REFERENCES_PRODUCT_LINE_COLUMN)

    ' Cheap pass: verify Part A3 stamps + References. If nothing changed,
    ' skip Calculate / row collection / sheet writes entirely.
    If ExportBuildIsCurrent(ffaValues, productLines, buildKey) Then
        GoTo CleanUp
    End If

    ' Formula columns W/X/Y need current values only when rebuilding.
    Application.Calculate

    Set ffaRows = CreateObject("Scripting.Dictionary")
    ffaRows.CompareMode = vbTextCompare
    Set productLineRows = CreateObject("Scripting.Dictionary")
    productLineRows.CompareMode = vbTextCompare

    For i = 0 To ArrayCount(ffaValues) - 1
        keyName = ffaValues(LBound(ffaValues) + i)
        If Len(keyName) > 0 Then
            If Not ffaRows.Exists(keyName) Then ffaRows.Add keyName, New Collection
        End If
    Next i

    For i = 0 To ArrayCount(productLines) - 1
        keyName = productLines(LBound(productLines) + i)
        If Len(keyName) > 0 Then
            If Not productLineRows.Exists(keyName) Then productLineRows.Add keyName, New Collection
        End If
    Next i

    CollectExportRowsFromPartSheets ffaRows, productLineRows
    RemoveObsoleteExportSheets ffaRows, productLineRows

    For i = 0 To ArrayCount(ffaValues) - 1
        keyName = ffaValues(LBound(ffaValues) + i)
        If Len(keyName) > 0 Then
            WriteExportSheet FFA_SHEET_PREFIX, keyName, EXPORT_TYPE_FFA, ffaRows(keyName)
        End If
    Next i

    For i = 0 To ArrayCount(productLines) - 1
        keyName = productLines(LBound(productLines) + i)
        If Len(keyName) > 0 Then
            WriteExportSheet PRODUCT_LINE_SHEET_PREFIX, keyName, EXPORT_TYPE_PRODUCT_LINE, productLineRows(keyName)
        End If
    Next i

    ' Rebuild key from the A3 stamps just written during collection.
    WriteExportBuildCache BuildExportBuildCacheKey(ffaValues, productLines)

CleanUp:
    OptimizeExcel False
End Sub

' Returns True when Part A3 cheap-checks and the workbook build cache all match.
Private Function ExportBuildIsCurrent( _
    ByRef ffaValues() As String, _
    ByRef productLines() As String, _
    ByRef buildKey As String) As Boolean

    Dim ws As Worksheet
    Dim stamps As Object
    Dim sheetKey As Variant
    Dim stampParts() As String
    Dim i As Long
    Dim partNumber As String
    Dim lastRow As Long
    Dim cheapSig As String
    Dim storedStamp As String

    If Not ExportTargetsExist(ffaValues, productLines) Then Exit Function

    Set stamps = CreateObject("Scripting.Dictionary")
    stamps.CompareMode = vbTextCompare

    For Each ws In ThisWorkbook.Worksheets
        If Not IsPartSheet(ws) Then GoTo NextSheet

        partNumber = BasePartWithoutDash(CStr(Nz(ws.Range(PART_NUMBER_CELL).Value)))
        If Len(partNumber) = 0 Then partNumber = BasePartWithoutDash(ws.Name)

        lastRow = FastLastUsedRowInColumns(ws, DATA_FIRST_COLUMN, DATA_LAST_COLUMN)
        cheapSig = BuildCheapPartExportSig(ws, partNumber, lastRow)
        storedStamp = CStr(Nz(ws.Range(PART_EXPORT_CACHE_CELL).Value2))

        If Not PartExportStampMatchesCheap(storedStamp, cheapSig) Then Exit Function

        If Not stamps.Exists(ws.Name) Then
            stamps.Add ws.Name, storedStamp
        End If
NextSheet:
    Next ws

    If stamps.Count = 0 Then
        stampParts = EmptyStringArray()
    Else
        ReDim stampParts(0 To stamps.Count - 1)
        i = 0
        For Each sheetKey In stamps.Keys
            stampParts(i) = CStr(sheetKey) & Chr$(30) & CStr(stamps(sheetKey))
            i = i + 1
        Next sheetKey
        QuickSortStrings stampParts, LBound(stampParts), UBound(stampParts)
    End If

    buildKey = EXPORT_BUILD_CACHE_SCHEMA & Chr$(31) & _
        PART_EXPORT_CACHE_SCHEMA & Chr$(31) & _
        JoinStringArray(ffaValues) & Chr$(31) & _
        JoinStringArray(productLines) & Chr$(31) & _
        JoinStringArray(stampParts)

    If StrComp(GetExportBuildCache(), buildKey, vbBinaryCompare) <> 0 Then Exit Function

    ExportBuildIsCurrent = True
End Function

Private Function ExportTargetsExist( _
    ByRef ffaValues() As String, _
    ByRef productLines() As String) As Boolean

    Dim i As Long
    Dim keyName As String

    For i = 0 To ArrayCount(ffaValues) - 1
        keyName = ffaValues(LBound(ffaValues) + i)
        If Len(keyName) > 0 Then
            If FindExportSheet(EXPORT_TYPE_FFA, keyName) Is Nothing Then Exit Function
        End If
    Next i

    For i = 0 To ArrayCount(productLines) - 1
        keyName = productLines(LBound(productLines) + i)
        If Len(keyName) > 0 Then
            If FindExportSheet(EXPORT_TYPE_PRODUCT_LINE, keyName) Is Nothing Then Exit Function
        End If
    Next i

    ExportTargetsExist = True
End Function

Private Function BuildExportBuildCacheKey( _
    ByRef ffaValues() As String, _
    ByRef productLines() As String) As String

    Dim ws As Worksheet
    Dim stamps As Object
    Dim sheetKey As Variant
    Dim stampParts() As String
    Dim i As Long
    Dim storedStamp As String

    Set stamps = CreateObject("Scripting.Dictionary")
    stamps.CompareMode = vbTextCompare

    For Each ws In ThisWorkbook.Worksheets
        If IsPartSheet(ws) Then
            storedStamp = CStr(Nz(ws.Range(PART_EXPORT_CACHE_CELL).Value2))
            If Not stamps.Exists(ws.Name) Then
                stamps.Add ws.Name, storedStamp
            End If
        End If
    Next ws

    If stamps.Count = 0 Then
        stampParts = EmptyStringArray()
    Else
        ReDim stampParts(0 To stamps.Count - 1)
        i = 0
        For Each sheetKey In stamps.Keys
            stampParts(i) = CStr(sheetKey) & Chr$(30) & CStr(stamps(sheetKey))
            i = i + 1
        Next sheetKey
        QuickSortStrings stampParts, LBound(stampParts), UBound(stampParts)
    End If

    BuildExportBuildCacheKey = EXPORT_BUILD_CACHE_SCHEMA & Chr$(31) & _
        PART_EXPORT_CACHE_SCHEMA & Chr$(31) & _
        JoinStringArray(ffaValues) & Chr$(31) & _
        JoinStringArray(productLines) & Chr$(31) & _
        JoinStringArray(stampParts)
End Function

Private Sub CollectExportRowsFromPartSheets(ByVal ffaRows As Object, ByVal productLineRows As Object)
    Dim ws As Worksheet
    Dim partNumber As String
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim ffaMark As String
    Dim markedProductLines As Collection
    Dim plIndex As Long
    Dim plName As String
    Dim exportRow As Variant
    Dim tableValues As Variant
    Dim plSig As String
    Dim fullStamp As String

    For Each ws In ThisWorkbook.Worksheets
        If Not IsPartSheet(ws) Then GoTo NextSheet

        partNumber = BasePartWithoutDash(CStr(Nz(ws.Range(PART_NUMBER_CELL).Value)))
        If Len(partNumber) = 0 Then partNumber = BasePartWithoutDash(ws.Name)

        lastRow = FastLastUsedRowInColumns(ws, DATA_FIRST_COLUMN, DATA_LAST_COLUMN)
        Set markedProductLines = MarkedProductLinesOnSheet(ws, plSig)

        If lastRow < LIST_START_ROW Then
            fullStamp = BuildPartExportStamp( _
                BuildCheapPartExportSig(ws, partNumber, lastRow), _
                "empty")
            WritePartExportCache ws, fullStamp
            GoTo NextSheet
        End If

        tableValues = ws.Range( _
            ws.Cells(LIST_START_ROW, DATA_FIRST_COLUMN), _
            ws.Cells(lastRow, DATA_LAST_COLUMN)).Value2

        fullStamp = BuildPartExportStamp( _
            BuildCheapPartExportSigFromParts(partNumber, lastRow, HashColumnFromTable(tableValues, OFF_FFA_MARK), plSig), _
            HashExportRelevantBlock(tableValues))
        WritePartExportCache ws, fullStamp

        If Not IsArray(tableValues) Then GoTo NextSheet

        For rowIndex = 1 To UBound(tableValues, 1)
            If Not RowHasExportableData(tableValues, rowIndex) Then GoTo NextDataRow

            exportRow = BuildExportRow(partNumber, tableValues, rowIndex)

            ffaMark = Trim$(CStr(Nz(tableValues(rowIndex, OFF_FFA_MARK))))
            If Len(ffaMark) > 0 Then
                If ffaRows.Exists(ffaMark) Then ffaRows(ffaMark).Add exportRow
            End If

            If Not markedProductLines Is Nothing Then
                For plIndex = 1 To markedProductLines.Count
                    plName = CStr(markedProductLines(plIndex))
                    If productLineRows.Exists(plName) Then
                        productLineRows(plName).Add exportRow
                    End If
                Next plIndex
            End If
NextDataRow:
        Next rowIndex
NextSheet:
    Next ws
End Sub

Private Function MarkedProductLinesOnSheet( _
    ByVal ws As Worksheet, _
    ByRef plSig As String) As Collection

    Dim lastRow As Long
    Dim rowIndex As Long
    Dim lineName As String
    Dim marked As Collection
    Dim sigParts As Collection

    Set marked = New Collection
    Set sigParts = New Collection
    plSig = vbNullString

    lastRow = FastLastUsedRowInColumn(ws, PRODUCT_LINE_VALUE_COLUMN)
    If lastRow < LIST_START_ROW Then
        Set MarkedProductLinesOnSheet = marked
        Exit Function
    End If

    For rowIndex = LIST_START_ROW To lastRow
        lineName = Trim$(CStr(Nz(ws.Cells(rowIndex, PRODUCT_LINE_VALUE_COLUMN).Value2)))
        If Len(lineName) = 0 Then Exit For

        If IsActiveFlag(ws.Cells(rowIndex, PRODUCT_LINE_CHECKBOX_COLUMN).Value2) Then
            marked.Add lineName
            sigParts.Add lineName & "=1"
        Else
            sigParts.Add lineName & "=0"
        End If
    Next rowIndex

    plSig = JoinCollection(sigParts)
    Set MarkedProductLinesOnSheet = marked
End Function

Private Function BuildCheapPartExportSig( _
    ByVal ws As Worksheet, _
    ByVal partNumber As String, _
    ByVal lastRow As Long) As String

    Dim zHash As String
    Dim plSig As String
    Dim marked As Collection

    If lastRow >= LIST_START_ROW Then
        zHash = HashVariantColumn( _
            ws.Range(ws.Cells(LIST_START_ROW, "Z"), ws.Cells(lastRow, "Z")).Value2)
    Else
        zHash = "0"
    End If

    Set marked = MarkedProductLinesOnSheet(ws, plSig)
    BuildCheapPartExportSig = BuildCheapPartExportSigFromParts(partNumber, lastRow, zHash, plSig)
End Function

Private Function BuildCheapPartExportSigFromParts( _
    ByVal partNumber As String, _
    ByVal lastRow As Long, _
    ByVal zHash As String, _
    ByVal plSig As String) As String

    BuildCheapPartExportSigFromParts = _
        UCase$(partNumber) & Chr$(30) & _
        CStr(lastRow) & Chr$(30) & _
        zHash & Chr$(30) & _
        plSig
End Function

Private Function BuildPartExportStamp(ByVal cheapSig As String, ByVal fullHash As String) As String
    BuildPartExportStamp = PART_EXPORT_CACHE_SCHEMA & Chr$(31) & cheapSig & Chr$(31) & fullHash
End Function

Private Function PartExportStampMatchesCheap(ByVal storedStamp As String, ByVal cheapSig As String) As Boolean
    Dim parts As Variant

    If Len(storedStamp) = 0 Then Exit Function
    parts = Split(storedStamp, Chr$(31))
    If UBound(parts) < 2 Then Exit Function
    If StrComp(CStr(parts(0)), PART_EXPORT_CACHE_SCHEMA, vbBinaryCompare) <> 0 Then Exit Function
    PartExportStampMatchesCheap = (StrComp(CStr(parts(1)), cheapSig, vbBinaryCompare) = 0)
End Function

Private Sub WritePartExportCache(ByVal ws As Worksheet, ByVal stamp As String)
    With ws.Range(PART_EXPORT_CACHE_CELL)
        .NumberFormat = ";;;"
        .Value2 = stamp
    End With
End Sub

Private Sub WriteExportBuildCache(ByVal buildKey As String)
    Dim nm As Name
    Dim escaped As String

    escaped = Replace$(buildKey, """", """""")

    On Error Resume Next
    ThisWorkbook.Names(EXPORT_BUILD_CACHE_NAME).Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add _
        Name:=EXPORT_BUILD_CACHE_NAME, _
        RefersTo:="=""" & escaped & """"
End Sub

Private Function GetExportBuildCache() As String
    Dim nm As Name
    Dim refersTo As String

    On Error Resume Next
    Set nm = ThisWorkbook.Names(EXPORT_BUILD_CACHE_NAME)
    On Error GoTo 0
    If nm Is Nothing Then Exit Function

    On Error Resume Next
    GetExportBuildCache = CStr(Application.Evaluate(nm.RefersTo))
    If Err.Number <> 0 Then
        Err.Clear
        refersTo = nm.RefersTo
        ' Fallback parse for ="value" style names.
        If Left$(refersTo, 2) = "=""" And Right$(refersTo, 1) = """" Then
            GetExportBuildCache = Replace$(Mid$(refersTo, 3, Len(refersTo) - 3), """""", """")
        End If
    End If
    On Error GoTo 0
End Function

Private Function RowHasExportableData(ByVal tableValues As Variant, ByVal rowIndex As Long) As Boolean
    If Len(Trim$(CStr(Nz(tableValues(rowIndex, OFF_OP_SEQUENCE))))) > 0 Then
        RowHasExportableData = True
        Exit Function
    End If
    If Len(Trim$(CStr(Nz(tableValues(rowIndex, OFF_OP_CODE))))) > 0 Then
        RowHasExportableData = True
        Exit Function
    End If
    If Len(Trim$(CStr(Nz(tableValues(rowIndex, OFF_FFA_MARK))))) > 0 Then
        RowHasExportableData = True
    End If
End Function

Private Function BuildExportRow( _
    ByVal partNumber As String, _
    ByVal tableValues As Variant, _
    ByVal rowIndex As Long) As Variant

    Dim exportRow(1 To EXPORT_COLUMN_COUNT) As Variant

    exportRow(1) = partNumber
    exportRow(2) = Nz(tableValues(rowIndex, OFF_OP_SEQUENCE))
    exportRow(3) = Nz(tableValues(rowIndex, OFF_OP_CODE))
    exportRow(4) = Nz(tableValues(rowIndex, OFF_PROCESS_HOURS))
    exportRow(5) = Nz(tableValues(rowIndex, OFF_AVG_EX))
    exportRow(6) = Nz(tableValues(rowIndex, OFF_BATCH_SIZE))
    exportRow(7) = Nz(tableValues(rowIndex, OFF_AVG_HPU))
    exportRow(8) = Nz(tableValues(rowIndex, OFF_EQUIPMENT_TYPE))

    BuildExportRow = exportRow
End Function

Private Function HashExportRelevantBlock(ByVal tableValues As Variant) As String
    Dim rowIndex As Long
    Dim totalLen As Long
    Dim checkSum As Long
    Dim count As Long
    Dim cellValue As String
    Dim offsets As Variant
    Dim i As Long

    If Not IsArray(tableValues) Then
        HashExportRelevantBlock = "0"
        Exit Function
    End If

    offsets = Array(OFF_OP_SEQUENCE, OFF_OP_CODE, OFF_BATCH_SIZE, OFF_EQUIPMENT_TYPE, _
        OFF_PROCESS_HOURS, OFF_AVG_EX, OFF_AVG_HPU, OFF_FFA_MARK)

    count = UBound(tableValues, 1)
    For rowIndex = 1 To count
        For i = LBound(offsets) To UBound(offsets)
            cellValue = Trim$(CStr(Nz(tableValues(rowIndex, CLng(offsets(i))))))
            totalLen = totalLen + Len(cellValue)
            If Len(cellValue) > 0 Then
                checkSum = (checkSum + AscW(Left$(cellValue, 1)) + AscW(Right$(cellValue, 1)) + Len(cellValue)) Mod 2147483647
            End If
        Next i
    Next rowIndex

    HashExportRelevantBlock = CStr(count) & ":" & CStr(totalLen) & ":" & CStr(checkSum)
End Function

Private Function HashColumnFromTable(ByVal tableValues As Variant, ByVal colOffset As Long) As String
    Dim rowIndex As Long
    Dim values() As Variant
    Dim count As Long

    If Not IsArray(tableValues) Then
        HashColumnFromTable = HashVariantColumn(tableValues)
        Exit Function
    End If

    count = UBound(tableValues, 1)
    ReDim values(1 To count, 1 To 1)
    For rowIndex = 1 To count
        values(rowIndex, 1) = tableValues(rowIndex, colOffset)
    Next rowIndex

    HashColumnFromTable = HashVariantColumn(values)
End Function

Private Function HashVariantColumn(ByVal values As Variant) As String
    Dim rowIndex As Long
    Dim cellValue As String
    Dim totalLen As Long
    Dim checkSum As Long
    Dim count As Long
    Dim firstValue As String
    Dim lastValue As String

    If IsEmpty(values) Then
        HashVariantColumn = "0:0:0:"
        Exit Function
    End If

    If Not IsArray(values) Then
        cellValue = Trim$(CStr(Nz(values)))
        HashVariantColumn = "1:" & CStr(Len(cellValue)) & ":" & CStr(Len(cellValue)) & ":" & cellValue
        Exit Function
    End If

    count = UBound(values, 1)
    For rowIndex = 1 To count
        cellValue = Trim$(CStr(Nz(values(rowIndex, 1))))
        totalLen = totalLen + Len(cellValue)
        If Len(cellValue) > 0 Then
            checkSum = (checkSum + AscW(Left$(cellValue, 1)) + AscW(Right$(cellValue, 1)) + Len(cellValue)) Mod 2147483647
            If Len(firstValue) = 0 Then firstValue = cellValue
            lastValue = cellValue
        End If
    Next rowIndex

    HashVariantColumn = CStr(count) & ":" & CStr(totalLen) & ":" & CStr(checkSum) & ":" & firstValue & Chr$(30) & lastValue
End Function

Private Sub WriteExportSheet( _
    ByVal namePrefix As String, _
    ByVal keyName As String, _
    ByVal exportType As String, _
    ByVal rows As Collection)

    Dim ws As Worksheet
    Dim output() As Variant
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim sourceRow As Variant
    Dim targetRange As Range

    Set ws = GetOrCreateExportSheet(namePrefix, keyName, exportType)

    ClearExportDataArea ws

    ws.Cells(EXPORT_HEADER_ROW, EXPORT_FIRST_COLUMN).Value = "Part Number"
    ws.Cells(EXPORT_HEADER_ROW, EXPORT_FIRST_COLUMN + 1).Value = "Op Sequence"
    ws.Cells(EXPORT_HEADER_ROW, EXPORT_FIRST_COLUMN + 2).Value = "Op Code"
    ws.Cells(EXPORT_HEADER_ROW, EXPORT_FIRST_COLUMN + 3).Value = "Process Hours"
    ws.Cells(EXPORT_HEADER_ROW, EXPORT_FIRST_COLUMN + 4).Value = "Avg Ex"
    ws.Cells(EXPORT_HEADER_ROW, EXPORT_FIRST_COLUMN + 5).Value = "Batch Size"
    ws.Cells(EXPORT_HEADER_ROW, EXPORT_FIRST_COLUMN + 6).Value = "Avg HPU"
    ws.Cells(EXPORT_HEADER_ROW, EXPORT_FIRST_COLUMN + 7).Value = "Equipment Type"

    If rows Is Nothing Then Exit Sub
    If rows.Count = 0 Then Exit Sub

    ReDim output(1 To rows.Count, 1 To EXPORT_COLUMN_COUNT)
    For rowIndex = 1 To rows.Count
        sourceRow = rows(rowIndex)
        For colIndex = 1 To EXPORT_COLUMN_COUNT
            output(rowIndex, colIndex) = sourceRow(colIndex)
        Next colIndex
    Next rowIndex

    Set targetRange = ws.Range( _
        ws.Cells(EXPORT_FIRST_DATA_ROW, EXPORT_FIRST_COLUMN), _
        ws.Cells(EXPORT_FIRST_DATA_ROW + rows.Count - 1, EXPORT_LAST_COLUMN))
    targetRange.Value2 = output
End Sub

Private Sub ClearExportDataArea(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim lastRowA As Long

    lastRow = FastLastUsedRowInColumns(ws, "C", "J")
    If lastRow < EXPORT_HEADER_ROW Then lastRow = EXPORT_HEADER_ROW
    ws.Range( _
        ws.Cells(EXPORT_HEADER_ROW, EXPORT_FIRST_COLUMN), _
        ws.Cells(lastRow, EXPORT_LAST_COLUMN)).ClearContents

    lastRow = FastLastUsedRowInColumn(ws, "B")
    If lastRow >= EXPORT_HEADER_ROW Then
        ws.Range(ws.Cells(EXPORT_HEADER_ROW, 2), ws.Cells(lastRow, 2)).ClearContents
    End If

    lastRowA = FastLastUsedRowInColumn(ws, "A")
    If lastRowA > 3 Then
        ws.Range(ws.Cells(4, 1), ws.Cells(lastRowA, 1)).ClearContents
    End If
End Sub

Private Function GetOrCreateExportSheet( _
    ByVal namePrefix As String, _
    ByVal keyName As String, _
    ByVal exportType As String) As Worksheet

    Dim ws As Worksheet
    Dim sheetName As String

    Set ws = FindExportSheet(exportType, keyName)
    sheetName = BuildExportSheetName(namePrefix, keyName)

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        On Error Resume Next
        ws.Name = sheetName
        If Err.Number <> 0 Then
            Err.Clear
            ws.Name = UniqueSheetName(sheetName)
        End If
        On Error GoTo 0
    ElseIf StrComp(ws.Name, sheetName, vbTextCompare) <> 0 Then
        On Error Resume Next
        ws.Name = sheetName
        If Err.Number <> 0 Then
            Err.Clear
            ws.Name = UniqueSheetName(sheetName)
        End If
        On Error GoTo 0
    End If

    ws.Range(EXPORT_MARKER_CELL).Value = EXPORT_MARKER_VALUE
    ws.Range(EXPORT_TYPE_CELL).Value = exportType
    ws.Range(EXPORT_KEY_CELL).Value = keyName
    ws.Range(LEGACY_EXPORT_MARKER_CELL).ClearContents
    ws.Range(LEGACY_EXPORT_TYPE_CELL).ClearContents
    ws.Range(LEGACY_EXPORT_KEY_CELL).ClearContents

    Set GetOrCreateExportSheet = ws
End Function

Private Function FindExportSheet(ByVal exportType As String, ByVal keyName As String) As Worksheet
    Dim ws As Worksheet
    Dim sheetType As String
    Dim sheetKey As String

    For Each ws In ThisWorkbook.Worksheets
        If IsExportSheet(ws) Then
            sheetType = ExportSheetType(ws)
            sheetKey = ExportSheetKey(ws)
            If StrComp(sheetType, exportType, vbTextCompare) = 0 Then
                If StrComp(sheetKey, keyName, vbTextCompare) = 0 Then
                    Set FindExportSheet = ws
                    Exit Function
                End If
            End If
        End If
    Next ws
End Function

Private Sub RemoveObsoleteExportSheets(ByVal ffaRows As Object, ByVal productLineRows As Object)
    Dim ws As Worksheet
    Dim sheetNames As Collection
    Dim sheetName As Variant
    Dim exportType As String
    Dim keyName As String
    Dim keepSheet As Boolean

    Set sheetNames = New Collection

    For Each ws In ThisWorkbook.Worksheets
        If IsExportSheet(ws) Then
            exportType = ExportSheetType(ws)
            keyName = ExportSheetKey(ws)
            keepSheet = False

            If StrComp(exportType, EXPORT_TYPE_FFA, vbTextCompare) = 0 Then
                keepSheet = ffaRows.Exists(keyName)
            ElseIf StrComp(exportType, EXPORT_TYPE_PRODUCT_LINE, vbTextCompare) = 0 Then
                keepSheet = productLineRows.Exists(keyName)
            End If

            If Not keepSheet Then sheetNames.Add ws.Name
        End If
    Next ws

    For Each sheetName In sheetNames
        ThisWorkbook.Worksheets(CStr(sheetName)).Delete
    Next sheetName
End Sub

Private Function IsExportSheet(ByVal ws As Worksheet) As Boolean
    If StrComp(Trim$(CStr(Nz(ws.Range(EXPORT_MARKER_CELL).Value))), EXPORT_MARKER_VALUE, vbTextCompare) = 0 Then
        IsExportSheet = True
        Exit Function
    End If
    IsExportSheet = (StrComp(Trim$(CStr(Nz(ws.Range(LEGACY_EXPORT_MARKER_CELL).Value))), EXPORT_MARKER_VALUE, vbTextCompare) = 0)
End Function

Private Function ExportSheetType(ByVal ws As Worksheet) As String
    ExportSheetType = Trim$(CStr(Nz(ws.Range(EXPORT_TYPE_CELL).Value)))
    If Len(ExportSheetType) = 0 Then
        ExportSheetType = Trim$(CStr(Nz(ws.Range(LEGACY_EXPORT_TYPE_CELL).Value)))
    End If
End Function

Private Function ExportSheetKey(ByVal ws As Worksheet) As String
    ExportSheetKey = Trim$(CStr(Nz(ws.Range(EXPORT_KEY_CELL).Value)))
    If Len(ExportSheetKey) = 0 Then
        ExportSheetKey = Trim$(CStr(Nz(ws.Range(LEGACY_EXPORT_KEY_CELL).Value)))
    End If
End Function

Private Function IsPartSheet(ByVal ws As Worksheet) As Boolean
    IsPartSheet = (StrComp(Trim$(CStr(Nz(ws.Range(PART_LABEL_CELL).Value))), PART_LABEL_VALUE, vbTextCompare) = 0)
End Function

Private Function BasePartWithoutDash(ByVal rawValue As String) As String
    Dim trimmed As String
    Dim dashPos As Long

    trimmed = Trim$(rawValue)
    If Len(trimmed) = 0 Then Exit Function

    dashPos = InStr(1, trimmed, "-", vbBinaryCompare)
    If dashPos > 1 Then
        BasePartWithoutDash = Trim$(Left$(trimmed, dashPos - 1))
    Else
        BasePartWithoutDash = trimmed
    End If
End Function

Private Function BuildExportSheetName(ByVal namePrefix As String, ByVal keyName As String) As String
    Dim cleanedKey As String
    Dim maxKeyLen As Long

    cleanedKey = SanitizeSheetNameFragment(keyName)
    maxKeyLen = 31 - Len(namePrefix)
    If maxKeyLen < 1 Then maxKeyLen = 1
    If Len(cleanedKey) > maxKeyLen Then cleanedKey = Left$(cleanedKey, maxKeyLen)

    BuildExportSheetName = SanitizeSheetNameFragment(namePrefix & cleanedKey)
End Function

Private Function UniqueSheetName(ByVal proposedName As String) As String
    Dim candidate As String
    Dim suffix As Long

    candidate = Left$(proposedName, 31)
    suffix = 1

    Do While SheetExists(candidate)
        candidate = Left$(proposedName, 31 - Len(CStr(suffix)) - 1) & "_" & CStr(suffix)
        suffix = suffix + 1
        If suffix > 9999 Then
            Err.Raise vbObjectError + 710, "UniqueSheetName", "Unable to invent a unique export sheet name."
        End If
    Loop

    UniqueSheetName = candidate
End Function

Private Function SanitizeSheetNameFragment(ByVal proposedName As String) As String
    Dim cleaned As String

    cleaned = proposedName
    cleaned = Replace$(cleaned, "\", "_")
    cleaned = Replace$(cleaned, "/", "_")
    cleaned = Replace$(cleaned, "?", "_")
    cleaned = Replace$(cleaned, "*", "_")
    cleaned = Replace$(cleaned, "[", "_")
    cleaned = Replace$(cleaned, "]", "_")
    cleaned = Replace$(cleaned, ":", "_")

    SanitizeSheetNameFragment = Trim$(cleaned)
End Function

Private Function GetReferenceColumnValues(ByVal columnLetter As String) As String()
    Dim wsReferences As Worksheet
    Dim lastRow As Long
    Dim rawValues As Variant
    Dim rowIndex As Long
    Dim cellValue As String
    Dim uniqueValues As Object

    Set uniqueValues = CreateObject("Scripting.Dictionary")
    uniqueValues.CompareMode = vbTextCompare

    Set wsReferences = ThisWorkbook.Worksheets(REFERENCES_SHEET_NAME)
    lastRow = FastLastUsedRowInColumn(wsReferences, columnLetter)

    If lastRow < REFERENCES_START_ROW Then
        GetReferenceColumnValues = EmptyStringArray()
        Exit Function
    End If

    rawValues = wsReferences.Range( _
        wsReferences.Cells(REFERENCES_START_ROW, columnLetter), _
        wsReferences.Cells(lastRow, columnLetter)).Value2

    If Not IsArray(rawValues) Then
        cellValue = Trim$(CStr(Nz(rawValues)))
        If Len(cellValue) > 0 Then uniqueValues.Add cellValue, cellValue
        GetReferenceColumnValues = DictionaryKeysToSortedArray(uniqueValues)
        Exit Function
    End If

    For rowIndex = 1 To UBound(rawValues, 1)
        cellValue = Trim$(CStr(Nz(rawValues(rowIndex, 1))))
        If Len(cellValue) > 0 Then
            If Not uniqueValues.Exists(cellValue) Then uniqueValues.Add cellValue, cellValue
        End If
    Next rowIndex

    GetReferenceColumnValues = DictionaryKeysToSortedArray(uniqueValues)
End Function

Private Function DictionaryKeysToSortedArray(ByVal dict As Object) As String()
    Dim keys() As String
    Dim key As Variant
    Dim i As Long

    If dict Is Nothing Or dict.Count = 0 Then
        DictionaryKeysToSortedArray = EmptyStringArray()
        Exit Function
    End If

    ReDim keys(0 To dict.Count - 1)
    i = 0
    For Each key In dict.Keys
        keys(i) = CStr(key)
        i = i + 1
    Next key

    QuickSortStrings keys, LBound(keys), UBound(keys)
    DictionaryKeysToSortedArray = keys
End Function

Private Sub QuickSortStrings(ByRef values() As String, ByVal firstIndex As Long, ByVal lastIndex As Long)
    Dim lowIndex As Long
    Dim highIndex As Long
    Dim pivot As String
    Dim swapValue As String

    lowIndex = firstIndex
    highIndex = lastIndex
    pivot = values((firstIndex + lastIndex) \ 2)

    Do While lowIndex <= highIndex
        Do While StrComp(values(lowIndex), pivot, vbTextCompare) < 0
            lowIndex = lowIndex + 1
        Loop
        Do While StrComp(values(highIndex), pivot, vbTextCompare) > 0
            highIndex = highIndex - 1
        Loop

        If lowIndex <= highIndex Then
            swapValue = values(lowIndex)
            values(lowIndex) = values(highIndex)
            values(highIndex) = swapValue
            lowIndex = lowIndex + 1
            highIndex = highIndex - 1
        End If
    Loop

    If firstIndex < highIndex Then QuickSortStrings values, firstIndex, highIndex
    If lowIndex < lastIndex Then QuickSortStrings values, lowIndex, lastIndex
End Sub

Private Function JoinStringArray(ByRef values() As String) As String
    Dim i As Long
    Dim parts() As String
    Dim count As Long

    count = ArrayCount(values)
    If count = 0 Then
        JoinStringArray = vbNullString
        Exit Function
    End If

    ReDim parts(0 To count - 1)
    For i = 0 To count - 1
        parts(i) = values(LBound(values) + i)
    Next i

    JoinStringArray = Join(parts, Chr$(30))
End Function

Private Function JoinCollection(ByVal values As Collection) As String
    Dim i As Long
    Dim parts() As String

    If values Is Nothing Or values.Count = 0 Then
        JoinCollection = vbNullString
        Exit Function
    End If

    ReDim parts(0 To values.Count - 1)
    For i = 1 To values.Count
        parts(i - 1) = CStr(values(i))
    Next i

    JoinCollection = Join(parts, Chr$(30))
End Function

Private Function EmptyStringArray() As String()
    Dim values() As String
    EmptyStringArray = values
End Function

Private Function ArrayCount(ByRef values() As String) As Long
    On Error Resume Next
    ArrayCount = UBound(values) - LBound(values) + 1
    If Err.Number <> 0 Then
        Err.Clear
        ArrayCount = 0
    End If
    On Error GoTo 0
End Function

Private Function IsActiveFlag(ByVal activeValue As Variant) As Boolean
    If IsError(activeValue) Then Exit Function
    If IsEmpty(activeValue) Or IsNull(activeValue) Then Exit Function

    If VarType(activeValue) = vbBoolean Then
        IsActiveFlag = CBool(activeValue)
        Exit Function
    End If

    If IsNumeric(activeValue) Then
        IsActiveFlag = (CDbl(activeValue) <> 0)
        Exit Function
    End If

    Select Case LCase$(Trim$(CStr(activeValue)))
        Case "true", "yes", "y", "1"
            IsActiveFlag = True
    End Select
End Function

Private Function FastLastUsedRowInColumn(ByVal ws As Worksheet, ByVal columnLetter As String) As Long
    FastLastUsedRowInColumn = ws.Cells(ws.Rows.Count, columnLetter).End(xlUp).Row
End Function

Private Function FastLastUsedRowInColumns( _
    ByVal ws As Worksheet, _
    ByVal firstColumn As String, _
    ByVal lastColumn As String) As Long

    Dim colIndex As Long
    Dim maxRow As Long
    Dim firstColIndex As Long
    Dim lastColIndex As Long

    firstColIndex = ws.Columns(firstColumn).Column
    lastColIndex = ws.Columns(lastColumn).Column
    maxRow = 1

    For colIndex = firstColIndex To lastColIndex
        If ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row > maxRow Then
            maxRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row
        End If
    Next colIndex

    FastLastUsedRowInColumns = maxRow
End Function

Private Function LastUsedRowInColumn(ByVal ws As Worksheet, ByVal columnLetter As String) As Long
    LastUsedRowInColumn = FastLastUsedRowInColumn(ws, columnLetter)
End Function

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    SheetExists = Not ws Is Nothing
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
