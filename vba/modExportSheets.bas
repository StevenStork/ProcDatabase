Attribute VB_Name = "modExportSheets"
Option Explicit

' Builds one export sheet per FFA and per product line from tblOperations /
' tblParts. Copy-out columns only:
'   Part Number | Op Sequence | Op Code | Process Hours | Avg Ex |
'   Batch Size  | Avg HPU     | Equipment Type

Public Const EXPORT_SCOPE_FFA As String = "FFA"
Public Const EXPORT_SCOPE_PRODUCT_LINE As String = "Product Line"
Public Const EXPORT_SCOPE_ALL As String = "All"

Public Sub ShowUpdateExportSheets()
    frmExportSheets.Show vbModal
End Sub

Public Sub BuildExportSheets()
    BuildExportSheetsCore EXPORT_SCOPE_ALL, vbNullString, False
End Sub

Public Sub BuildExportSheetsCore( _
    ByVal scopeName As String, _
    ByVal itemName As String, _
    ByVal forceRebuild As Boolean)

    Dim ffaValues() As String
    Dim productLines() As String
    Dim ffaRows As Object
    Dim productLineRows As Object
    Dim i As Long
    Dim keyName As String
    Dim scopeIsAll As Boolean
    Dim scopeIsFfa As Boolean
    Dim scopeIsProductLine As Boolean

    scopeIsAll = (StrComp(scopeName, EXPORT_SCOPE_ALL, vbTextCompare) = 0)
    scopeIsFfa = (StrComp(scopeName, EXPORT_SCOPE_FFA, vbTextCompare) = 0)
    scopeIsProductLine = (StrComp(scopeName, EXPORT_SCOPE_PRODUCT_LINE, vbTextCompare) = 0)

    If Not scopeIsAll And Not scopeIsFfa And Not scopeIsProductLine Then
        Err.Raise vbObjectError + 720, "BuildExportSheetsCore", "Unknown export scope: " & scopeName
    End If
    If Not scopeIsAll Then
        If Len(Trim$(itemName)) = 0 Then
            Err.Raise vbObjectError + 721, "BuildExportSheetsCore", "Select an item to export."
        End If
    End If

    On Error GoTo CleanUp
    OptimizeExcel True

    EnsureDataSheet
    EnsureReferencesSheet
    SyncDirtyPartsToStore

    ffaValues = ReferenceColumnValues(REFS_FFA_COLUMN)
    productLines = ReferenceColumnValues(REFS_PRODUCT_LINE_COLUMN)

    If scopeIsFfa Then
        If Not ArrayContainsValue(ffaValues, itemName) Then
            Err.Raise vbObjectError + 722, "BuildExportSheetsCore", _
                "FFA """ & itemName & """ was not found on the References sheet."
        End If
    ElseIf scopeIsProductLine Then
        If Not ArrayContainsValue(productLines, itemName) Then
            Err.Raise vbObjectError + 723, "BuildExportSheetsCore", _
                "Product line """ & itemName & """ was not found on the References sheet."
        End If
    End If

    If scopeIsAll And Not forceRebuild Then
        If ExportOpsHashIsCurrent() Then GoTo CleanUp
    End If

    Application.Calculate

    Set ffaRows = CreateObject("Scripting.Dictionary")
    ffaRows.CompareMode = vbTextCompare
    Set productLineRows = CreateObject("Scripting.Dictionary")
    productLineRows.CompareMode = vbTextCompare

    If scopeIsAll Or scopeIsFfa Then
        If scopeIsFfa Then
            ffaRows.Add Trim$(itemName), New Collection
        Else
            For i = 0 To ArrayCount(ffaValues) - 1
                keyName = ffaValues(LBound(ffaValues) + i)
                If Len(keyName) > 0 Then
                    If Not ffaRows.Exists(keyName) Then ffaRows.Add keyName, New Collection
                End If
            Next i
        End If
    End If

    If scopeIsAll Or scopeIsProductLine Then
        If scopeIsProductLine Then
            productLineRows.Add Trim$(itemName), New Collection
        Else
            For i = 0 To ArrayCount(productLines) - 1
                keyName = productLines(LBound(productLines) + i)
                If Len(keyName) > 0 Then
                    If Not productLineRows.Exists(keyName) Then productLineRows.Add keyName, New Collection
                End If
            Next i
        End If
    End If

    CollectExportRowsFromStore ffaRows, productLineRows

    If scopeIsAll Then RemoveObsoleteExportSheets ffaRows, productLineRows

    If scopeIsAll Or scopeIsFfa Then
        For i = 0 To ArrayCount(ffaValues) - 1
            keyName = ffaValues(LBound(ffaValues) + i)
            If Len(keyName) = 0 Then GoTo NextFfa
            If scopeIsFfa And StrComp(keyName, itemName, vbTextCompare) <> 0 Then GoTo NextFfa
            WriteExportSheet FFA_SHEET_PREFIX, keyName, EXPORT_TYPE_FFA, ffaRows(keyName)
NextFfa:
        Next i
    End If

    If scopeIsAll Or scopeIsProductLine Then
        For i = 0 To ArrayCount(productLines) - 1
            keyName = productLines(LBound(productLines) + i)
            If Len(keyName) = 0 Then GoTo NextPl
            If scopeIsProductLine And StrComp(keyName, itemName, vbTextCompare) <> 0 Then GoTo NextPl
            WriteExportSheet PRODUCT_LINE_SHEET_PREFIX, keyName, EXPORT_TYPE_PRODUCT_LINE, productLineRows(keyName)
NextPl:
        Next i
    End If

    If scopeIsAll Then
        WriteExportOpsHash
    End If

CleanUp:
    OptimizeExcel False
End Sub

Public Function ListExportFfas() As String()
    ListExportFfas = ReferenceColumnValues(REFS_FFA_COLUMN)
End Function

Public Function ListExportProductLines() As String()
    ListExportProductLines = ReferenceColumnValues(REFS_PRODUCT_LINE_COLUMN)
End Function

Public Function ConfirmExportOverwrite(ByVal scopeDescription As String) As Boolean
    Dim response As VbMsgBoxResult

    response = MsgBox( _
        "This will overwrite " & scopeDescription & "." & vbCrLf & vbCrLf & _
        "You cannot get the old export back." & vbCrLf & vbCrLf & _
        "Do you want to continue?", _
        vbExclamation + vbYesNo + vbDefaultButton2, _
        "Confirm Export Update")

    ConfirmExportOverwrite = (response = vbYes)
End Function

Private Sub CollectExportRowsFromStore(ByVal ffaRows As Object, ByVal productLineRows As Object)
    Dim ops As ListObject
    Dim parts As ListObject
    Dim rowIndex As Long
    Dim partNumber As String
    Dim ffaValue As String
    Dim exportRow As Variant
    Dim plMap As Object
    Dim plKey As Variant
    Dim plList As String

    Set ops = OpsTable()
    Set parts = PartsTable()
    Set plMap = ProductLinesByPart(parts)

    If ops.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = 1 To ops.DataBodyRange.Rows.Count
        partNumber = Trim$(CStr(Nz(ListColumnByHeader(ops, COL_OPS_PART_NUMBER).DataBodyRange.Cells(rowIndex, 1).Value)))
        If Len(partNumber) = 0 Then GoTo NextRow
        If Len(Trim$(CStr(Nz(ListColumnByHeader(ops, HDR_OP_SEQUENCE).DataBodyRange.Cells(rowIndex, 1).Value)))) = 0 Then
            If Len(Trim$(CStr(Nz(ListColumnByHeader(ops, HDR_OP_CODE).DataBodyRange.Cells(rowIndex, 1).Value)))) = 0 Then
                GoTo NextRow
            End If
        End If

        exportRow = BuildExportRow(ops, rowIndex)
        ffaValue = Trim$(CStr(Nz(ListColumnByHeader(ops, HDR_FFA).DataBodyRange.Cells(rowIndex, 1).Value)))

        If Len(ffaValue) > 0 Then
            If ffaRows.Exists(ffaValue) Then ffaRows(ffaValue).Add exportRow
        End If

        If plMap.Exists(partNumber) Then
            plList = CStr(plMap(partNumber))
            For Each plKey In productLineRows.Keys
                If ListContainsToken(plList, CStr(plKey)) Then
                    productLineRows(plKey).Add exportRow
                End If
            Next plKey
        End If
NextRow:
    Next rowIndex
End Sub

Private Function ProductLinesByPart(ByVal parts As ListObject) As Object
    Dim map As Object
    Dim rowIndex As Long
    Dim basePart As String

    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = vbTextCompare
    If parts.DataBodyRange Is Nothing Then
        Set ProductLinesByPart = map
        Exit Function
    End If

    For rowIndex = 1 To parts.DataBodyRange.Rows.Count
        basePart = Trim$(CStr(Nz(parts.ListColumns(COL_PARTS_BASE).DataBodyRange.Cells(rowIndex, 1).Value)))
        If Len(basePart) > 0 Then
            If Not map.Exists(basePart) Then
                map.Add basePart, CStr(Nz(parts.ListColumns(COL_PARTS_PRODUCT_LINES).DataBodyRange.Cells(rowIndex, 1).Value))
            End If
        End If
    Next rowIndex

    Set ProductLinesByPart = map
End Function

Private Function BuildExportRow(ByVal ops As ListObject, ByVal rowIndex As Long) As Variant
    Dim values(1 To 8) As Variant

    values(1) = ListColumnByHeader(ops, COL_OPS_PART_NUMBER).DataBodyRange.Cells(rowIndex, 1).Value
    values(2) = ListColumnByHeader(ops, HDR_OP_SEQUENCE).DataBodyRange.Cells(rowIndex, 1).Value
    values(3) = ListColumnByHeader(ops, HDR_OP_CODE).DataBodyRange.Cells(rowIndex, 1).Value
    values(4) = ListColumnByHeader(ops, HDR_PROCESS_HOURS).DataBodyRange.Cells(rowIndex, 1).Value
    values(5) = ListColumnByHeader(ops, HDR_AVG_EX).DataBodyRange.Cells(rowIndex, 1).Value
    values(6) = ListColumnByHeader(ops, HDR_BATCH_SIZE).DataBodyRange.Cells(rowIndex, 1).Value
    values(7) = ListColumnByHeader(ops, HDR_AVG_HPU).DataBodyRange.Cells(rowIndex, 1).Value
    values(8) = ListColumnByHeader(ops, HDR_EQUIPMENT_TYPE).DataBodyRange.Cells(rowIndex, 1).Value
    BuildExportRow = values
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
    Dim headers As Variant
    Dim i As Long

    Set ws = GetOrCreateExportSheet(namePrefix, keyName, exportType)
    ClearExportDataArea ws

    headers = ExportHeaderNames()
    For i = LBound(headers) To UBound(headers)
        ws.Cells(EXPORT_HEADER_ROW, EXPORT_FIRST_COLUMN + i - LBound(headers)).Value = headers(i)
    Next i

    If rows Is Nothing Then Exit Sub
    If rows.Count = 0 Then Exit Sub

    ReDim output(1 To rows.Count, 1 To EXPORT_COLUMN_COUNT)
    For rowIndex = 1 To rows.Count
        sourceRow = rows(rowIndex)
        For colIndex = 1 To EXPORT_COLUMN_COUNT
            output(rowIndex, colIndex) = sourceRow(colIndex)
        Next colIndex
    Next rowIndex

    ws.Range( _
        ws.Cells(EXPORT_FIRST_DATA_ROW, EXPORT_FIRST_COLUMN), _
        ws.Cells(EXPORT_FIRST_DATA_ROW + rows.Count - 1, EXPORT_LAST_COLUMN)).Value2 = output
End Sub

Private Sub ClearExportDataArea(ByVal ws As Worksheet)
    Dim lastRow As Long

    lastRow = FastLastUsedRowInColumns(ws, "C", "J")
    If lastRow < EXPORT_HEADER_ROW Then lastRow = EXPORT_HEADER_ROW
    ws.Range( _
        ws.Cells(EXPORT_HEADER_ROW, EXPORT_FIRST_COLUMN), _
        ws.Cells(lastRow, EXPORT_LAST_COLUMN)).ClearContents
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

    ws.Range(CATEGORY_CELL).Value = EXPORT_LABEL_VALUE
    ws.Range(EXPORT_TYPE_CELL).Value = exportType
    ws.Range(EXPORT_KEY_CELL).Value = keyName
    Set GetOrCreateExportSheet = ws
End Function

Private Function FindExportSheet(ByVal exportType As String, ByVal keyName As String) As Worksheet
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If IsExportSheet(ws) Then
            If StrComp(Trim$(CStr(Nz(ws.Range(EXPORT_TYPE_CELL).Value))), exportType, vbTextCompare) = 0 Then
                If StrComp(Trim$(CStr(Nz(ws.Range(EXPORT_KEY_CELL).Value))), keyName, vbTextCompare) = 0 Then
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
            exportType = Trim$(CStr(Nz(ws.Range(EXPORT_TYPE_CELL).Value)))
            keyName = Trim$(CStr(Nz(ws.Range(EXPORT_KEY_CELL).Value)))
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

Private Function BuildExportSheetName(ByVal namePrefix As String, ByVal keyName As String) As String
    Dim cleanedKey As String
    Dim maxKeyLen As Long

    cleanedKey = SanitizeSheetName(keyName)
    maxKeyLen = 31 - Len(namePrefix)
    If maxKeyLen < 1 Then maxKeyLen = 1
    If Len(cleanedKey) > maxKeyLen Then cleanedKey = Left$(cleanedKey, maxKeyLen)
    BuildExportSheetName = SanitizeSheetName(namePrefix & cleanedKey)
End Function
