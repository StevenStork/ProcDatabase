Attribute VB_Name = "modPartSheetEditor"
Option Explicit

'==============================================================================
' Sheet-based part editor on PartEditor with cache on PartEditorCache.
'==============================================================================

Public Sub LoadPartToEditor(Optional ByVal partInput As String = vbNullString)
    Dim ws As Worksheet
    Dim basePartCode As String
    Dim dashCondition As String
    Dim inputValue As String

    Set ws = GetPartEditorWorksheet()
    If ws Is Nothing Then
        MsgBox "PartEditor sheet was not found. Run BootstrapCapacityTables first.", vbExclamation
        Exit Sub
    End If

    If Len(partInput) = 0 Then
        inputValue = Trim$(CStr(ws.Cells(PE_INPUT_ROW, PE_VALUE_COL).Value2))
    Else
        inputValue = Trim$(partInput)
        ws.Cells(PE_INPUT_ROW, PE_VALUE_COL).Value = inputValue
    End If

    If Len(inputValue) = 0 Then
        MsgBox "Enter a base part or assembly number in cell C3.", vbExclamation
        Exit Sub
    End If

    SplitAssemblyNo inputValue, basePartCode, dashCondition
    basePartCode = NormalizeCode(basePartCode)
    If Len(basePartCode) = 0 Then
        MsgBox "Could not resolve a base part number.", vbExclamation
        Exit Sub
    End If

    On Error GoTo CleanUp
    OptimizeExcel True

    EnsureCacheSheetExists
    ClearEditorDataRanges ws
    ws.Cells(PE_BASE_PART_ROW, PE_VALUE_COL).Value = basePartCode

    LoadMasterFields ws, basePartCode
    LoadDashRows ws, basePartCode
    LoadOperationRows ws, basePartCode
    WriteEditorCache basePartCode
    ApplyFactoryValidation ws

    SetEditorStatus ws, "Loaded " & basePartCode & "."

CleanUp:
    OptimizeExcel False
End Sub

Public Sub SavePartFromEditor()
    Dim ws As Worksheet
    Dim basePartCode As String
    Dim factoryCode As String
    Dim factoryTbl As ListObject
    Dim fieldValues As Object

    Set ws = GetPartEditorWorksheet()
    If ws Is Nothing Then Exit Sub

    basePartCode = NormalizeCode(CStr(ws.Cells(PE_BASE_PART_ROW, PE_VALUE_COL).Value2))
    If Not ValidateRequiredCode(basePartCode, "Base Part") Then Exit Sub

    factoryCode = NormalizeCode(CStr(ws.Cells(PE_ROW_FACTORY, PE_VALUE_COL).Value2))
    If Not ValidateRequiredCode(factoryCode, "Factory") Then Exit Sub

    Set factoryTbl = FindTable(FACTORIES_TABLE_NAME)
    If Not ValidateForeignKeyExists(factoryTbl, COL_FACTORY_CODE, factoryCode, "Factory") Then Exit Sub

    On Error GoTo CleanUp
    OptimizeExcel True

    Set fieldValues = NewFieldValuesDictionary()
    fieldValues(COL_BASE_PART_CODE) = basePartCode
    fieldValues(COL_FACTORY_CODE) = factoryCode
    fieldValues(COL_ACTIVE) = ActiveFlagToCellValue(ReadEditorActiveFlag(ws))
    fieldValues(COL_STATUS_DATE) = ParseEditorStatusDate(ws.Cells(PE_ROW_STATUS_DATE, PE_VALUE_COL).Value2)
    fieldValues(COL_NOTES) = Trim$(CStr(ws.Cells(PE_ROW_NOTES, PE_VALUE_COL).Value2))

    UpsertRow FindTable(BASE_PARTS_TABLE_NAME), COL_BASE_PART_CODE, basePartCode, fieldValues
    SyncDashAssignments basePartCode, ws
    SyncOperationAssignments basePartCode, ws
    WriteEditorCache basePartCode

    SetEditorStatus ws, "Saved " & basePartCode & "."

CleanUp:
    OptimizeExcel False
End Sub

Public Sub ClearPartEditor()
    Dim ws As Worksheet

    Set ws = GetPartEditorWorksheet()
    If ws Is Nothing Then Exit Sub

    On Error GoTo CleanUp
    OptimizeExcel True

    ws.Cells(PE_INPUT_ROW, PE_VALUE_COL).ClearContents
    ClearEditorDataRanges ws
    ClearEditorCache
    SetEditorStatus ws, vbNullString

CleanUp:
    OptimizeExcel False
End Sub

Public Sub OpenPartEditorFromPartsIndex()
    Dim wsParts As Worksheet
    Dim wsEditor As Worksheet
    Dim selectedCode As String
    Dim tbl As ListObject

    On Error GoTo Fail

    Set wsParts = ThisWorkbook.Worksheets(PARTS_SHEET_NAME)
    If TypeName(Selection) <> "Range" Then GoTo Fail
    If Selection.ListObject Is Nothing Then GoTo Fail

    Set tbl = Selection.ListObject
    If tbl.Name <> BASE_PARTS_TABLE_NAME Then GoTo Fail
    If tbl.DataBodyRange Is Nothing Then GoTo Fail
    If Intersect(Selection, tbl.DataBodyRange) Is Nothing Then GoTo Fail

    selectedCode = NormalizeCode(Selection.Cells(1, 1).Value2)
    If Len(selectedCode) = 0 Then GoTo Fail

    Set wsEditor = GetPartEditorWorksheet()
    If wsEditor Is Nothing Then GoTo Fail

    wsEditor.Activate
    wsEditor.Cells(PE_INPUT_ROW, PE_VALUE_COL).Value = selectedCode
    LoadPartToEditor selectedCode
    Exit Sub

Fail:
    MsgBox "Select a row in the Parts index table (BasePartsTbl), then run this macro.", vbExclamation
End Sub

Private Sub LoadMasterFields(ByVal ws As Worksheet, ByVal basePartCode As String)
    Dim tbl As ListObject
    Dim listRowIndex As Long

    Set tbl = FindTable(BASE_PARTS_TABLE_NAME)
    listRowIndex = FindListRowByKey(tbl, COL_BASE_PART_CODE, basePartCode)

    If listRowIndex = 0 Then
        ws.Cells(PE_ROW_FACTORY, PE_VALUE_COL).ClearContents
        ws.Cells(PE_ROW_ACTIVE, PE_VALUE_COL).Value = True
        ws.Cells(PE_ROW_STATUS_DATE, PE_VALUE_COL).ClearContents
        ws.Cells(PE_ROW_NOTES, PE_VALUE_COL).ClearContents
        Exit Sub
    End If

    ws.Cells(PE_ROW_FACTORY, PE_VALUE_COL).Value = CStr(NzBlank(GetCellValueByListRow(tbl, listRowIndex, COL_FACTORY_CODE)))
    ws.Cells(PE_ROW_ACTIVE, PE_VALUE_COL).Value = IsActiveFlag(GetCellValueByListRow(tbl, listRowIndex, COL_ACTIVE))
    ws.Cells(PE_ROW_STATUS_DATE, PE_VALUE_COL).Value = FormatEditorStatusDate(GetCellValueByListRow(tbl, listRowIndex, COL_STATUS_DATE))
    ws.Cells(PE_ROW_NOTES, PE_VALUE_COL).Value = CStr(NzBlank(GetCellValueByListRow(tbl, listRowIndex, COL_NOTES)))
End Sub

Private Sub LoadDashRows(ByVal ws As Worksheet, ByVal basePartCode As String)
    Dim tbl As ListObject
    Dim dashCodes As Variant
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim sheetRow As Long
    Dim loadedCount As Long
    Dim dashCell As Range

    Set tbl = FindTable(PART_DASH_CONDITIONS_TABLE_NAME)
    If tbl Is Nothing Or tbl.DataBodyRange Is Nothing Then Exit Sub

    dashCodes = tbl.ListColumns(COL_DASH_CONDITION).DataBodyRange.Value2
    If Not IsArray(dashCodes) Then Exit Sub

    rowCount = UBound(dashCodes, 1)
    sheetRow = PE_DASH_DATA_START_ROW
    loadedCount = 0

    For rowIndex = 1 To rowCount
        If Not ValuesMatchCode(GetCellValueByListRow(tbl, rowIndex, COL_BASE_PART_CODE), basePartCode) Then GoTo ContinueDash
        If loadedCount >= PE_DASH_MAX_ROWS Then Exit For

        Set dashCell = ws.Cells(sheetRow, PE_COL_DASH)
        dashCell.NumberFormat = "@"
        dashCell.Value = CStr(NzBlank(GetCellValueByListRow(tbl, rowIndex, COL_DASH_CONDITION)))
        If TableHasColumn(tbl, COL_SEPARATOR) Then
            ws.Cells(sheetRow, PE_COL_SEPARATOR).Value = CStr(NzBlank(GetCellValueByListRow(tbl, rowIndex, COL_SEPARATOR)))
        End If
        ws.Cells(sheetRow, PE_COL_DASH_ACTIVE).Value = IsActiveFlag(GetCellValueByListRow(tbl, rowIndex, COL_ACTIVE))
        ws.Cells(sheetRow, PE_COL_DASH_NOTES).Value = CStr(NzBlank(GetCellValueByListRow(tbl, rowIndex, COL_NOTES)))

        sheetRow = sheetRow + 1
        loadedCount = loadedCount + 1

ContinueDash:
    Next rowIndex
End Sub

Private Sub LoadOperationRows(ByVal ws As Worksheet, ByVal basePartCode As String)
    Dim tbl As ListObject
    Dim operSeqValues As Variant
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim sheetRow As Long
    Dim loadedCount As Long
    Dim operSeq As String
    Dim avgHours As Variant
    Dim avgEx As Variant

    Set tbl = FindTable(PART_OPERATIONS_TABLE_NAME)
    If tbl Is Nothing Or tbl.DataBodyRange Is Nothing Then Exit Sub

    operSeqValues = tbl.ListColumns(COL_OPER_SEQ).DataBodyRange.Value2
    If Not IsArray(operSeqValues) Then Exit Sub

    rowCount = UBound(operSeqValues, 1)
    sheetRow = PE_OPS_DATA_START_ROW
    loadedCount = 0

    For rowIndex = 1 To rowCount
        If Not ValuesMatchCode(GetCellValueByListRow(tbl, rowIndex, COL_BASE_PART_CODE), basePartCode) Then GoTo ContinueOp
        If loadedCount >= PE_OPS_MAX_ROWS Then Exit For

        operSeq = Trim$(CStr(NzBlank(GetCellValueByListRow(tbl, rowIndex, COL_OPER_SEQ))))
        ws.Cells(sheetRow, PE_COL_OPER_SEQ).Value = operSeq
        ws.Cells(sheetRow, PE_COL_OPER_NAME).Value = CStr(NzBlank(GetCellValueByListRow(tbl, rowIndex, COL_OPERATION_NAME)))
        ws.Cells(sheetRow, PE_COL_OPER_ACTIVE).Value = IsActiveFlag(GetCellValueByListRow(tbl, rowIndex, COL_ACTIVE))
        ws.Cells(sheetRow, PE_COL_OPER_NOTES).Value = CStr(NzBlank(GetCellValueByListRow(tbl, rowIndex, COL_NOTES)))

        avgHours = AvgProcessHoursByBasePartAndOp(basePartCode, operSeq)
        avgEx = AvgExByBasePartAndOp(basePartCode, operSeq)
        ws.Cells(sheetRow, PE_COL_AVG_HOURS).Value = FormatAverageDisplay(avgHours)
        ws.Cells(sheetRow, PE_COL_AVG_EX).Value = FormatAverageDisplay(avgEx)

        sheetRow = sheetRow + 1
        loadedCount = loadedCount + 1

ContinueOp:
    Next rowIndex
End Sub

Private Sub SyncDashAssignments(ByVal basePartCode As String, ByVal ws As Worksheet)
    Dim tbl As ListObject
    Dim cacheRows As Object
    Dim currentRows As Object
    Dim dashKey As Variant
    Dim rowData As Variant
    Dim fieldValues As Object
    Dim dashCell As Range
    Dim listRowIndex As Long

    Set tbl = FindTable(PART_DASH_CONDITIONS_TABLE_NAME)
    If tbl Is Nothing Then Exit Sub

    Set cacheRows = ReadCachedDashRows()
    Set currentRows = ReadSheetDashRows(ws)

    For Each dashKey In cacheRows.Keys
        If Not currentRows.Exists(dashKey) Then
            DeleteJunctionRow tbl, COL_BASE_PART_CODE, basePartCode, COL_DASH_CONDITION, CStr(dashKey)
        End If
    Next dashKey

    For Each dashKey In currentRows.Keys
        rowData = currentRows(dashKey)
        Set fieldValues = NewFieldValuesDictionary()
        fieldValues(COL_BASE_PART_CODE) = basePartCode
        fieldValues(COL_DASH_CONDITION) = CStr(dashKey)
        fieldValues(COL_SEPARATOR) = CStr(rowData(1))
        fieldValues(COL_ACTIVE) = ActiveFlagToCellValue(CBool(rowData(2)))
        fieldValues(COL_NOTES) = CStr(rowData(3))

        UpsertJunctionRow tbl, COL_BASE_PART_CODE, basePartCode, COL_DASH_CONDITION, CStr(dashKey), fieldValues

        listRowIndex = FindJunctionListRow(tbl, COL_BASE_PART_CODE, basePartCode, COL_DASH_CONDITION, CStr(dashKey))
        If listRowIndex > 0 Then
            Set dashCell = tbl.ListRows(listRowIndex).Range.Cells(1, TableColumnIndex(tbl, COL_DASH_CONDITION))
            dashCell.NumberFormat = "@"
            dashCell.Value = CStr(dashKey)
        End If
    Next dashKey
End Sub

Private Sub SyncOperationAssignments(ByVal basePartCode As String, ByVal ws As Worksheet)
    Dim tbl As ListObject
    Dim cacheRows As Object
    Dim currentRows As Object
    Dim operKey As Variant
    Dim rowData As Variant
    Dim fieldValues As Object

    Set tbl = FindTable(PART_OPERATIONS_TABLE_NAME)
    If tbl Is Nothing Then Exit Sub

    Set cacheRows = ReadCachedOperationRows()
    Set currentRows = ReadSheetOperationRows(ws)

    For Each operKey In cacheRows.Keys
        If Not currentRows.Exists(operKey) Then
            DeleteJunctionRow tbl, COL_BASE_PART_CODE, basePartCode, COL_OPER_SEQ, CStr(operKey)
        End If
    Next operKey

    For Each operKey In currentRows.Keys
        rowData = currentRows(operKey)
        Set fieldValues = NewFieldValuesDictionary()
        fieldValues(COL_BASE_PART_CODE) = basePartCode
        fieldValues(COL_OPER_SEQ) = CStr(operKey)
        fieldValues(COL_OPERATION_NAME) = CStr(rowData(1))
        fieldValues(COL_ACTIVE) = ActiveFlagToCellValue(CBool(rowData(2)))
        fieldValues(COL_NOTES) = CStr(rowData(3))

        UpsertJunctionRow tbl, COL_BASE_PART_CODE, basePartCode, COL_OPER_SEQ, CStr(operKey), fieldValues
    Next operKey
End Sub

Private Sub WriteEditorCache(ByVal basePartCode As String)
    Dim wsCache As Worksheet
    Dim wsEditor As Worksheet
    Dim rowIndex As Long
    Dim sheetRow As Long
    Dim dashRows As Object
    Dim opRows As Object
    Dim dashKey As Variant
    Dim operKey As Variant
    Dim rowData As Variant

    Set wsCache = GetCacheWorksheet()
    Set wsEditor = GetPartEditorWorksheet()
    If wsCache Is Nothing Or wsEditor Is Nothing Then Exit Sub

    wsCache.Cells.Clear
    wsCache.Range(CACHE_BASE_PART_CELL).Value = basePartCode

    wsCache.Cells(CACHE_DASH_START_ROW - 1, 1).Value = COL_DASH_CONDITION
    wsCache.Cells(CACHE_DASH_START_ROW - 1, 2).Value = COL_SEPARATOR
    wsCache.Cells(CACHE_DASH_START_ROW - 1, 3).Value = COL_ACTIVE
    wsCache.Cells(CACHE_DASH_START_ROW - 1, 4).Value = COL_NOTES

    Set dashRows = ReadSheetDashRows(wsEditor)
    sheetRow = CACHE_DASH_START_ROW
    For Each dashKey In dashRows.Keys
        rowData = dashRows(dashKey)
        wsCache.Cells(sheetRow, 1).NumberFormat = "@"
        wsCache.Cells(sheetRow, 1).Value = CStr(dashKey)
        wsCache.Cells(sheetRow, 2).Value = rowData(1)
        wsCache.Cells(sheetRow, 3).Value = rowData(2)
        wsCache.Cells(sheetRow, 4).Value = rowData(3)
        sheetRow = sheetRow + 1
    Next dashKey

    wsCache.Cells(CACHE_OPS_START_ROW - 1, 1).Value = COL_OPER_SEQ
    wsCache.Cells(CACHE_OPS_START_ROW - 1, 2).Value = COL_OPERATION_NAME
    wsCache.Cells(CACHE_OPS_START_ROW - 1, 3).Value = COL_ACTIVE
    wsCache.Cells(CACHE_OPS_START_ROW - 1, 4).Value = COL_NOTES

    Set opRows = ReadSheetOperationRows(wsEditor)
    sheetRow = CACHE_OPS_START_ROW
    For Each operKey In opRows.Keys
        rowData = opRows(operKey)
        wsCache.Cells(sheetRow, 1).Value = CStr(operKey)
        wsCache.Cells(sheetRow, 2).Value = rowData(1)
        wsCache.Cells(sheetRow, 3).Value = rowData(2)
        wsCache.Cells(sheetRow, 4).Value = rowData(3)
        sheetRow = sheetRow + 1
    Next operKey
End Sub

Private Sub ClearEditorCache()
    Dim wsCache As Worksheet

    Set wsCache = GetCacheWorksheet()
    If wsCache Is Nothing Then Exit Sub

    wsCache.Cells.Clear
End Sub

Private Function ReadCachedDashRows() As Object
    Set ReadCachedDashRows = ReadCacheSection(CACHE_DASH_START_ROW, 1, 4)
End Function

Private Function ReadCachedOperationRows() As Object
    Set ReadCachedOperationRows = ReadCacheSection(CACHE_OPS_START_ROW, 1, 4)
End Function

Private Function ReadCacheSection(ByVal startRow As Long, ByVal keyCol As Long, ByVal valueColCount As Long) As Object
    Dim wsCache As Worksheet
    Dim rows As Object
    Dim rowIndex As Long
    Dim keyValue As String
    Dim rowData As Variant

    Set rows = CreateObject("Scripting.Dictionary")
    rows.CompareMode = vbTextCompare

    Set wsCache = GetCacheWorksheet()
    If wsCache Is Nothing Then
        Set ReadCacheSection = rows
        Exit Function
    End If

    rowIndex = startRow
    Do While Len(Trim$(CStr(wsCache.Cells(rowIndex, keyCol).Value2))) > 0 Or Len(Trim$(wsCache.Cells(rowIndex, keyCol).Text)) > 0
        If Len(Trim$(wsCache.Cells(rowIndex, keyCol).Text)) > 0 Then
            keyValue = Trim$(wsCache.Cells(rowIndex, keyCol).Text)
        Else
            keyValue = NormalizeCode(CStr(wsCache.Cells(rowIndex, keyCol).Value2))
        End If
        If Len(keyValue) > 0 Then
            ReDim rowData(0 To valueColCount - 1)
            If valueColCount >= 2 Then rowData(1) = wsCache.Cells(rowIndex, keyCol + 1).Value2
            If valueColCount >= 3 Then rowData(2) = wsCache.Cells(rowIndex, keyCol + 2).Value2
            If valueColCount >= 4 Then rowData(3) = wsCache.Cells(rowIndex, keyCol + 3).Value2
            rows(keyValue) = rowData
        End If
        rowIndex = rowIndex + 1
        If rowIndex > startRow + 100 Then Exit Do
    Loop

    Set ReadCacheSection = rows
End Function

Private Function ReadSheetDashRows(ByVal ws As Worksheet) As Object
    Dim rows As Object
    Dim rowIndex As Long
    Dim dashCode As String
    Dim separator As String
    Dim rowData As Variant
    Dim dashCell As Range

    Set rows = CreateObject("Scripting.Dictionary")
    rows.CompareMode = vbBinaryCompare

    For rowIndex = PE_DASH_DATA_START_ROW To PE_DASH_DATA_START_ROW + PE_DASH_MAX_ROWS - 1
        Set dashCell = ws.Cells(rowIndex, PE_COL_DASH)
        If Len(Trim$(dashCell.Text)) > 0 Then
            dashCode = CStr(dashCell.Text)
        Else
            dashCode = Trim$(CStr(dashCell.Value2))
        End If
        If Len(dashCode) = 0 Then GoTo ContinueDash

        separator = Trim$(CStr(ws.Cells(rowIndex, PE_COL_SEPARATOR).Value2))
        If Len(separator) = 0 Then separator = "-"

        ReDim rowData(0 To 3)
        rowData(1) = separator
        rowData(2) = IsActiveFlag(ws.Cells(rowIndex, PE_COL_DASH_ACTIVE).Value2)
        rowData(3) = Trim$(CStr(ws.Cells(rowIndex, PE_COL_DASH_NOTES).Value2))
        rows(dashCode) = rowData

ContinueDash:
    Next rowIndex

    Set ReadSheetDashRows = rows
End Function

Private Function ReadSheetOperationRows(ByVal ws As Worksheet) As Object
    Dim rows As Object
    Dim rowIndex As Long
    Dim operSeq As String
    Dim rowData As Variant

    Set rows = CreateObject("Scripting.Dictionary")
    rows.CompareMode = vbTextCompare

    For rowIndex = PE_OPS_DATA_START_ROW To PE_OPS_DATA_START_ROW + PE_OPS_MAX_ROWS - 1
        operSeq = Trim$(CStr(ws.Cells(rowIndex, PE_COL_OPER_SEQ).Value2))
        If Len(operSeq) = 0 Then GoTo ContinueOp

        ReDim rowData(0 To 3)
        rowData(1) = Trim$(CStr(ws.Cells(rowIndex, PE_COL_OPER_NAME).Value2))
        rowData(2) = IsActiveFlag(ws.Cells(rowIndex, PE_COL_OPER_ACTIVE).Value2)
        rowData(3) = Trim$(CStr(ws.Cells(rowIndex, PE_COL_OPER_NOTES).Value2))
        rows(NormalizeOperSeqKey(operSeq)) = rowData

ContinueOp:
    Next rowIndex

    Set ReadSheetOperationRows = rows
End Function

Private Sub ClearEditorDataRanges(ByVal ws As Worksheet)
    ws.Cells(PE_BASE_PART_ROW, PE_VALUE_COL).ClearContents
    ws.Range(ws.Cells(PE_ROW_FACTORY, PE_VALUE_COL), ws.Cells(PE_ROW_NOTES, PE_VALUE_COL)).ClearContents
    ws.Range(ws.Cells(PE_DASH_DATA_START_ROW, PE_COL_DASH), _
        ws.Cells(PE_DASH_DATA_START_ROW + PE_DASH_MAX_ROWS - 1, PE_COL_DASH_NOTES)).ClearContents
    ws.Range(ws.Cells(PE_DASH_DATA_START_ROW, PE_COL_DASH), _
        ws.Cells(PE_DASH_DATA_START_ROW + PE_DASH_MAX_ROWS - 1, PE_COL_DASH)).NumberFormat = "@"
    ws.Range(ws.Cells(PE_OPS_DATA_START_ROW, PE_COL_OPER_SEQ), _
        ws.Cells(PE_OPS_DATA_START_ROW + PE_OPS_MAX_ROWS - 1, PE_COL_AVG_EX)).ClearContents
End Sub

Private Sub ApplyFactoryValidation(ByVal ws As Worksheet)
    Dim factoryCodes As String
    Dim validationRange As Range

    factoryCodes = BuildFactoryValidationList()
    Set validationRange = ws.Cells(PE_ROW_FACTORY, PE_VALUE_COL)

    On Error Resume Next
    validationRange.Validation.Delete
    On Error GoTo 0

    If Len(factoryCodes) = 0 Then Exit Sub

    validationRange.Validation.Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, _
        Formula1:=factoryCodes
End Sub

Private Function BuildFactoryValidationList() As String
    Dim tbl As ListObject
    Dim codes As Variant
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim codeValue As String
    Dim result As String

    Set tbl = FindTable(FACTORIES_TABLE_NAME)
    If tbl Is Nothing Or tbl.DataBodyRange Is Nothing Then Exit Function

    codes = tbl.ListColumns(COL_FACTORY_CODE).DataBodyRange.Value2
    If Not IsArray(codes) Then
        codeValue = NormalizeCode(codes)
        If Len(codeValue) > 0 And IsActiveFlag(GetCellValueByListRow(tbl, tbl.ListRows(1).Index, COL_ACTIVE)) Then
            BuildFactoryValidationList = codeValue
        End If
        Exit Function
    End If

    rowCount = UBound(codes, 1)
    For rowIndex = 1 To rowCount
        codeValue = NormalizeCode(codes(rowIndex, 1))
        If Len(codeValue) = 0 Then GoTo ContinueFactory
        If Not IsActiveFlag(GetCellValueByListRow(tbl, tbl.ListRows(rowIndex).Index, COL_ACTIVE)) Then GoTo ContinueFactory

        If Len(result) > 0 Then result = result & ","
        result = result & codeValue

ContinueFactory:
    Next rowIndex

    BuildFactoryValidationList = result
End Function

Private Function ReadEditorActiveFlag(ByVal ws As Worksheet) As Boolean
    ReadEditorActiveFlag = IsActiveFlag(ws.Cells(PE_ROW_ACTIVE, PE_VALUE_COL).Value2)
End Function

Private Function FormatEditorStatusDate(ByVal rawValue As Variant) As String
    If IsDate(rawValue) Then
        FormatEditorStatusDate = Format$(CDate(rawValue), "yyyy-mm-dd")
    Else
        FormatEditorStatusDate = Trim$(CStr(NzBlank(rawValue)))
    End If
End Function

Private Function ParseEditorStatusDate(ByVal rawValue As Variant) As Variant
    If IsEmpty(rawValue) Or Len(Trim$(CStr(rawValue))) = 0 Then
        ParseEditorStatusDate = Empty
    ElseIf IsDate(rawValue) Then
        ParseEditorStatusDate = CDate(rawValue)
    Else
        ParseEditorStatusDate = Trim$(CStr(rawValue))
    End If
End Function

Private Function NormalizeOperSeqKey(ByVal operSeq As String) As String
    operSeq = Trim$(operSeq)
    If IsNumeric(operSeq) Then
        NormalizeOperSeqKey = CStr(CLng(CDbl(operSeq)))
    Else
        NormalizeOperSeqKey = operSeq
    End If
End Function

Private Sub SetEditorStatus(ByVal ws As Worksheet, ByVal statusText As String)
    ws.Cells(PE_STATUS_ROW, PE_STATUS_VALUE_COL).Value = statusText
End Sub

Private Sub EnsureCacheSheetExists()
    Dim ws As Worksheet

    Set ws = GetCacheWorksheet()
    If ws Is Nothing Then
        BootstrapCapacityTables
    End If
End Sub

Private Function GetPartEditorWorksheet() As Worksheet
    On Error Resume Next
    Set GetPartEditorWorksheet = ThisWorkbook.Worksheets(PART_EDITOR_SHEET_NAME)
    On Error GoTo 0
End Function

Private Function GetCacheWorksheet() As Worksheet
    On Error Resume Next
    Set GetCacheWorksheet = ThisWorkbook.Worksheets(PART_EDITOR_CACHE_SHEET_NAME)
    On Error GoTo 0
End Function

Private Function NzBlank(ByVal value As Variant) As Variant
    If IsError(value) Then
        NzBlank = vbNullString
    ElseIf IsEmpty(value) Or IsNull(value) Then
        NzBlank = vbNullString
    Else
        NzBlank = value
    End If
End Function
