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
    ApplyOperationDropdowns ws

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
    fieldValues(COL_NOTES) = ReadEditorNotes(ws)

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
        ClearEditorNotes ws
        Exit Sub
    End If

    ws.Cells(PE_ROW_FACTORY, PE_VALUE_COL).Value = CStr(NzBlank(GetCellValueByListRow(tbl, listRowIndex, COL_FACTORY_CODE)))
    ws.Cells(PE_ROW_ACTIVE, PE_VALUE_COL).Value = IsActiveFlag(GetCellValueByListRow(tbl, listRowIndex, COL_ACTIVE))
    WriteEditorNotes ws, CStr(NzBlank(GetCellValueByListRow(tbl, listRowIndex, COL_NOTES)))
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
    Dim showAvgHours As Boolean
    Dim showAvgEx As Boolean
    Dim equipmentCode As String
    Dim processTypeCode As String

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

        equipmentCode = vbNullString
        If TableHasColumn(tbl, COL_EQUIPMENT_CODE) Then
            equipmentCode = NormalizeCode(CStr(NzBlank(GetCellValueByListRow(tbl, rowIndex, COL_EQUIPMENT_CODE))))
        End If
        ws.Cells(sheetRow, PE_COL_EQUIPMENT).Value = equipmentCode

        processTypeCode = vbNullString
        If TableHasColumn(tbl, COL_PROCESS_TYPE_CODE) Then
            processTypeCode = NormalizeCode(CStr(NzBlank(GetCellValueByListRow(tbl, rowIndex, COL_PROCESS_TYPE_CODE))))
        End If
        ws.Cells(sheetRow, PE_COL_PROCESS_TYPE).Value = processTypeCode

        ws.Cells(sheetRow, PE_COL_OPER_ACTIVE).Value = IsActiveFlag(GetCellValueByListRow(tbl, rowIndex, COL_ACTIVE))
        ws.Cells(sheetRow, PE_COL_OPER_NOTES).Value = CStr(NzBlank(GetCellValueByListRow(tbl, rowIndex, COL_NOTES)))

        showAvgHours = True
        If TableHasColumn(tbl, COL_SHOW_AVG_HOURS) Then
            If Not IsBlankCellValue(GetCellValueByListRow(tbl, rowIndex, COL_SHOW_AVG_HOURS)) Then
                showAvgHours = IsActiveFlag(GetCellValueByListRow(tbl, rowIndex, COL_SHOW_AVG_HOURS))
            End If
        End If
        ws.Cells(sheetRow, PE_COL_SHOW_AVG_HOURS).Value = showAvgHours

        showAvgEx = True
        If TableHasColumn(tbl, COL_SHOW_AVG_EX) Then
            If Not IsBlankCellValue(GetCellValueByListRow(tbl, rowIndex, COL_SHOW_AVG_EX)) Then
                showAvgEx = IsActiveFlag(GetCellValueByListRow(tbl, rowIndex, COL_SHOW_AVG_EX))
            End If
        End If
        ws.Cells(sheetRow, PE_COL_SHOW_AVG_EX).Value = showAvgEx

        ApplyAveragesForOperationRow ws, sheetRow, basePartCode

        sheetRow = sheetRow + 1
        loadedCount = loadedCount + 1

ContinueOp:
    Next rowIndex
End Sub

' Called from ThisWorkbook SheetChange for cascading dropdowns and avg toggles.
Public Sub HandlePartEditorSheetChange(ByVal Target As Range)
    Dim ws As Worksheet
    Dim factoryCell As Range
    Dim opsEquipment As Range
    Dim opsShowToggle As Range
    Dim opsSeq As Range
    Dim changedRow As Long
    Dim basePartCode As String

    If Target Is Nothing Then Exit Sub
    Set ws = Target.Worksheet
    If StrComp(ws.Name, PART_EDITOR_SHEET_NAME, vbTextCompare) <> 0 Then Exit Sub

    Set factoryCell = ws.Cells(PE_ROW_FACTORY, PE_VALUE_COL)
    Set opsEquipment = ws.Range( _
        ws.Cells(PE_OPS_DATA_START_ROW, PE_COL_EQUIPMENT), _
        ws.Cells(PE_OPS_DATA_START_ROW + PE_OPS_MAX_ROWS - 1, PE_COL_EQUIPMENT))
    Set opsShowToggle = ws.Range( _
        ws.Cells(PE_OPS_DATA_START_ROW, PE_COL_SHOW_AVG_HOURS), _
        ws.Cells(PE_OPS_DATA_START_ROW + PE_OPS_MAX_ROWS - 1, PE_COL_SHOW_AVG_EX))
    Set opsSeq = ws.Range( _
        ws.Cells(PE_OPS_DATA_START_ROW, PE_COL_OPER_SEQ), _
        ws.Cells(PE_OPS_DATA_START_ROW + PE_OPS_MAX_ROWS - 1, PE_COL_OPER_SEQ))

    On Error GoTo CleanUp
    Application.EnableEvents = False

    If Not Intersect(Target, factoryCell) Is Nothing Then
        ApplyOperationDropdowns ws
        ClearInvalidEquipmentAndProcess ws
        GoTo CleanUp
    End If

    If Not Intersect(Target, opsEquipment) Is Nothing Then
        changedRow = Intersect(Target, opsEquipment).Row
        ApplyProcessTypeValidationForRow ws, changedRow
        ClearInvalidProcessTypeForRow ws, changedRow
        GoTo CleanUp
    End If

    If Not Intersect(Target, opsShowToggle) Is Nothing Or Not Intersect(Target, opsSeq) Is Nothing Then
        basePartCode = NormalizeCode(CStr(ws.Cells(PE_BASE_PART_ROW, PE_VALUE_COL).Value2))
        changedRow = Target.Row
        If changedRow >= PE_OPS_DATA_START_ROW And changedRow <= PE_OPS_DATA_START_ROW + PE_OPS_MAX_ROWS - 1 Then
            ApplyAveragesForOperationRow ws, changedRow, basePartCode
        End If
    End If

CleanUp:
    Application.EnableEvents = True
End Sub

Private Sub ApplyAveragesForOperationRow(ByVal ws As Worksheet, ByVal sheetRow As Long, ByVal basePartCode As String)
    Dim operSeq As String
    Dim avgHours As Variant
    Dim avgEx As Variant

    If Len(basePartCode) = 0 Then
        basePartCode = NormalizeCode(CStr(ws.Cells(PE_BASE_PART_ROW, PE_VALUE_COL).Value2))
    End If

    operSeq = Trim$(CStr(ws.Cells(sheetRow, PE_COL_OPER_SEQ).Value2))
    If Len(operSeq) = 0 Or Len(basePartCode) = 0 Then
        ws.Cells(sheetRow, PE_COL_AVG_HOURS).ClearContents
        ws.Cells(sheetRow, PE_COL_AVG_EX).ClearContents
        Exit Sub
    End If

    If IsActiveFlag(ws.Cells(sheetRow, PE_COL_SHOW_AVG_HOURS).Value2) Then
        avgHours = AvgProcessHoursByBasePartAndOp(basePartCode, operSeq)
        ws.Cells(sheetRow, PE_COL_AVG_HOURS).Value = FormatAverageDisplay(avgHours)
    Else
        ws.Cells(sheetRow, PE_COL_AVG_HOURS).ClearContents
    End If

    If IsActiveFlag(ws.Cells(sheetRow, PE_COL_SHOW_AVG_EX).Value2) Then
        avgEx = AvgExByBasePartAndOp(basePartCode, operSeq)
        ws.Cells(sheetRow, PE_COL_AVG_EX).Value = FormatAverageDisplay(avgEx)
    Else
        ws.Cells(sheetRow, PE_COL_AVG_EX).ClearContents
    End If
End Sub

Public Sub ApplyOperationDropdowns(ByVal ws As Worksheet)
    Dim factoryCode As String
    Dim equipmentList As String
    Dim rowIndex As Long
    Dim equipmentRange As Range

    If ws Is Nothing Then Set ws = GetPartEditorWorksheet()
    If ws Is Nothing Then Exit Sub

    factoryCode = NormalizeCode(CStr(ws.Cells(PE_ROW_FACTORY, PE_VALUE_COL).Value2))
    equipmentList = BuildEquipmentValidationList(factoryCode)

    Set equipmentRange = ws.Range( _
        ws.Cells(PE_OPS_DATA_START_ROW, PE_COL_EQUIPMENT), _
        ws.Cells(PE_OPS_DATA_START_ROW + PE_OPS_MAX_ROWS - 1, PE_COL_EQUIPMENT))
    ApplyListValidation equipmentRange, equipmentList

    For rowIndex = PE_OPS_DATA_START_ROW To PE_OPS_DATA_START_ROW + PE_OPS_MAX_ROWS - 1
        ApplyProcessTypeValidationForRow ws, rowIndex
    Next rowIndex
End Sub

Private Sub ApplyProcessTypeValidationForRow(ByVal ws As Worksheet, ByVal sheetRow As Long)
    Dim equipmentCode As String
    Dim processList As String

    equipmentCode = NormalizeCode(CStr(ws.Cells(sheetRow, PE_COL_EQUIPMENT).Value2))
    processList = BuildProcessTypeValidationList(equipmentCode)
    ApplyListValidation ws.Cells(sheetRow, PE_COL_PROCESS_TYPE), processList
End Sub

Private Sub ClearInvalidProcessTypeForRow(ByVal ws As Worksheet, ByVal sheetRow As Long)
    Dim processTypeCode As String
    Dim processList As String

    processTypeCode = NormalizeCode(CStr(ws.Cells(sheetRow, PE_COL_PROCESS_TYPE).Value2))
    If Len(processTypeCode) = 0 Then Exit Sub

    processList = "," & UCase$(BuildProcessTypeValidationList( _
        NormalizeCode(CStr(ws.Cells(sheetRow, PE_COL_EQUIPMENT).Value2)))) & ","
    If InStr(1, processList, "," & UCase$(processTypeCode) & ",", vbTextCompare) = 0 Then
        ws.Cells(sheetRow, PE_COL_PROCESS_TYPE).ClearContents
    End If
End Sub

Private Sub ClearInvalidEquipmentAndProcess(ByVal ws As Worksheet)
    Dim factoryCode As String
    Dim equipmentList As String
    Dim rowIndex As Long
    Dim equipmentCode As String

    factoryCode = NormalizeCode(CStr(ws.Cells(PE_ROW_FACTORY, PE_VALUE_COL).Value2))
    equipmentList = "," & UCase$(BuildEquipmentValidationList(factoryCode)) & ","

    For rowIndex = PE_OPS_DATA_START_ROW To PE_OPS_DATA_START_ROW + PE_OPS_MAX_ROWS - 1
        equipmentCode = NormalizeCode(CStr(ws.Cells(rowIndex, PE_COL_EQUIPMENT).Value2))
        If Len(equipmentCode) > 0 Then
            If InStr(1, equipmentList, "," & UCase$(equipmentCode) & ",", vbTextCompare) = 0 Then
                ws.Cells(rowIndex, PE_COL_EQUIPMENT).ClearContents
                ws.Cells(rowIndex, PE_COL_PROCESS_TYPE).ClearContents
            Else
                ClearInvalidProcessTypeForRow ws, rowIndex
            End If
        End If
        ApplyProcessTypeValidationForRow ws, rowIndex
    Next rowIndex
End Sub

Private Function IsBlankCellValue(ByVal rawValue As Variant) As Boolean
    If IsError(rawValue) Then
        IsBlankCellValue = True
    ElseIf IsEmpty(rawValue) Or IsNull(rawValue) Then
        IsBlankCellValue = True
    Else
        IsBlankCellValue = (Len(Trim$(CStr(rawValue))) = 0)
    End If
End Function

Private Sub ApplyListValidation(ByVal targetRange As Range, ByVal csvList As String)
    On Error Resume Next
    targetRange.Validation.Delete
    On Error GoTo 0

    If Len(csvList) = 0 Then Exit Sub

    targetRange.Validation.Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, _
        Formula1:=csvList
End Sub

Private Function BuildEquipmentValidationList(ByVal factoryCode As String) As String
    Dim junctionTbl As ListObject
    Dim equipmentTbl As ListObject
    Dim rowIndex As Long
    Dim equipmentCode As String
    Dim result As String
    Dim seen As Object

    factoryCode = NormalizeCode(factoryCode)
    If Len(factoryCode) = 0 Then Exit Function

    Set junctionTbl = FindTable(FACTORY_EQUIPMENT_TABLE_NAME)
    If junctionTbl Is Nothing Or junctionTbl.DataBodyRange Is Nothing Then Exit Function

    Set equipmentTbl = FindTable(EQUIPMENT_TABLE_NAME)
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    For rowIndex = 1 To junctionTbl.ListRows.Count
        If Not ValuesMatchCode(GetCellValueByListRow(junctionTbl, rowIndex, COL_FACTORY_CODE), factoryCode) Then GoTo ContinueEquip
        equipmentCode = NormalizeCode(CStr(NzBlank(GetCellValueByListRow(junctionTbl, rowIndex, COL_EQUIPMENT_CODE))))
        If Len(equipmentCode) = 0 Then GoTo ContinueEquip
        If seen.Exists(equipmentCode) Then GoTo ContinueEquip

        If Not equipmentTbl Is Nothing Then
            If FindListRowByKey(equipmentTbl, COL_EQUIPMENT_CODE, equipmentCode) > 0 Then
                If Not IsActiveFlag(GetCellValueByListRow(equipmentTbl, _
                    FindListRowByKey(equipmentTbl, COL_EQUIPMENT_CODE, equipmentCode), COL_ACTIVE)) Then GoTo ContinueEquip
            End If
        End If

        seen.Add equipmentCode, True
        If Len(result) > 0 Then result = result & ","
        result = result & equipmentCode

ContinueEquip:
    Next rowIndex

    BuildEquipmentValidationList = result
End Function

Private Function BuildProcessTypeValidationList(ByVal equipmentCode As String) As String
    Dim junctionTbl As ListObject
    Dim processTbl As ListObject
    Dim rowIndex As Long
    Dim processCode As String
    Dim result As String
    Dim seen As Object
    Dim processRow As Long

    equipmentCode = NormalizeCode(equipmentCode)
    If Len(equipmentCode) = 0 Then Exit Function

    Set junctionTbl = FindTable(EQUIPMENT_PROCESSES_TABLE_NAME)
    If junctionTbl Is Nothing Or junctionTbl.DataBodyRange Is Nothing Then Exit Function

    Set processTbl = FindTable(PROCESS_TYPES_TABLE_NAME)
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    For rowIndex = 1 To junctionTbl.ListRows.Count
        If Not ValuesMatchCode(GetCellValueByListRow(junctionTbl, rowIndex, COL_EQUIPMENT_CODE), equipmentCode) Then GoTo ContinueProcess
        If TableHasColumn(junctionTbl, COL_ACTIVE) Then
            If Not IsActiveFlag(GetCellValueByListRow(junctionTbl, rowIndex, COL_ACTIVE)) Then GoTo ContinueProcess
        End If

        processCode = NormalizeCode(CStr(NzBlank(GetCellValueByListRow(junctionTbl, rowIndex, COL_PROCESS_TYPE_CODE))))
        If Len(processCode) = 0 Then GoTo ContinueProcess
        If seen.Exists(processCode) Then GoTo ContinueProcess

        If Not processTbl Is Nothing Then
            processRow = FindListRowByKey(processTbl, COL_PROCESS_TYPE_CODE, processCode)
            If processRow > 0 Then
                If Not IsActiveFlag(GetCellValueByListRow(processTbl, processRow, COL_ACTIVE)) Then GoTo ContinueProcess
            End If
        End If

        seen.Add processCode, True
        If Len(result) > 0 Then result = result & ","
        result = result & processCode

ContinueProcess:
    Next rowIndex

    BuildProcessTypeValidationList = result
End Function

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
        fieldValues(COL_EQUIPMENT_CODE) = CStr(rowData(4))
        fieldValues(COL_PROCESS_TYPE_CODE) = CStr(rowData(5))
        fieldValues(COL_SHOW_AVG_HOURS) = ActiveFlagToCellValue(CBool(rowData(6)))
        fieldValues(COL_SHOW_AVG_EX) = ActiveFlagToCellValue(CBool(rowData(7)))

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
    wsCache.Cells(CACHE_OPS_START_ROW - 1, 5).Value = COL_EQUIPMENT_CODE
    wsCache.Cells(CACHE_OPS_START_ROW - 1, 6).Value = COL_PROCESS_TYPE_CODE
    wsCache.Cells(CACHE_OPS_START_ROW - 1, 7).Value = COL_SHOW_AVG_HOURS
    wsCache.Cells(CACHE_OPS_START_ROW - 1, 8).Value = COL_SHOW_AVG_EX

    Set opRows = ReadSheetOperationRows(wsEditor)
    sheetRow = CACHE_OPS_START_ROW
    For Each operKey In opRows.Keys
        rowData = opRows(operKey)
        wsCache.Cells(sheetRow, 1).Value = CStr(operKey)
        wsCache.Cells(sheetRow, 2).Value = rowData(1)
        wsCache.Cells(sheetRow, 3).Value = rowData(2)
        wsCache.Cells(sheetRow, 4).Value = rowData(3)
        wsCache.Cells(sheetRow, 5).Value = rowData(4)
        wsCache.Cells(sheetRow, 6).Value = rowData(5)
        wsCache.Cells(sheetRow, 7).Value = rowData(6)
        wsCache.Cells(sheetRow, 8).Value = rowData(7)
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
    Set ReadCachedOperationRows = ReadCacheSection(CACHE_OPS_START_ROW, 1, CACHE_OPS_VALUE_COL_COUNT)
End Function

Private Function ReadCacheSection(ByVal startRow As Long, ByVal keyCol As Long, ByVal valueColCount As Long) As Object
    Dim wsCache As Worksheet
    Dim rows As Object
    Dim rowIndex As Long
    Dim keyValue As String
    Dim rowData As Variant
    Dim valueIndex As Long

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
            For valueIndex = 1 To valueColCount - 1
                rowData(valueIndex) = wsCache.Cells(rowIndex, keyCol + valueIndex).Value2
            Next valueIndex
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

        ReDim rowData(0 To CACHE_OPS_VALUE_COL_COUNT - 1)
        rowData(1) = Trim$(CStr(ws.Cells(rowIndex, PE_COL_OPER_NAME).Value2))
        rowData(2) = IsActiveFlag(ws.Cells(rowIndex, PE_COL_OPER_ACTIVE).Value2)
        rowData(3) = Trim$(CStr(ws.Cells(rowIndex, PE_COL_OPER_NOTES).Value2))
        rowData(4) = NormalizeCode(CStr(ws.Cells(rowIndex, PE_COL_EQUIPMENT).Value2))
        rowData(5) = NormalizeCode(CStr(ws.Cells(rowIndex, PE_COL_PROCESS_TYPE).Value2))
        rowData(6) = IsActiveFlag(ws.Cells(rowIndex, PE_COL_SHOW_AVG_HOURS).Value2)
        rowData(7) = IsActiveFlag(ws.Cells(rowIndex, PE_COL_SHOW_AVG_EX).Value2)
        ' Empty toggle cells default to True on save.
        If IsEmpty(ws.Cells(rowIndex, PE_COL_SHOW_AVG_HOURS).Value2) _
            Or Len(Trim$(CStr(NzBlank(ws.Cells(rowIndex, PE_COL_SHOW_AVG_HOURS).Value2)))) = 0 Then
            rowData(6) = True
        End If
        If IsEmpty(ws.Cells(rowIndex, PE_COL_SHOW_AVG_EX).Value2) _
            Or Len(Trim$(CStr(NzBlank(ws.Cells(rowIndex, PE_COL_SHOW_AVG_EX).Value2)))) = 0 Then
            rowData(7) = True
        End If
        rows(NormalizeOperSeqKey(operSeq)) = rowData

ContinueOp:
    Next rowIndex

    Set ReadSheetOperationRows = rows
End Function

Private Sub ClearEditorDataRanges(ByVal ws As Worksheet)
    ws.Cells(PE_BASE_PART_ROW, PE_VALUE_COL).ClearContents
    ws.Cells(PE_ROW_FACTORY, PE_VALUE_COL).ClearContents
    ws.Cells(PE_ROW_ACTIVE, PE_VALUE_COL).ClearContents
    ClearEditorNotes ws
    ws.Range(ws.Cells(PE_DASH_DATA_START_ROW, PE_COL_DASH), _
        ws.Cells(PE_DASH_DATA_START_ROW + PE_DASH_MAX_ROWS - 1, PE_COL_DASH_NOTES)).ClearContents
    ws.Range(ws.Cells(PE_DASH_DATA_START_ROW, PE_COL_DASH), _
        ws.Cells(PE_DASH_DATA_START_ROW + PE_DASH_MAX_ROWS - 1, PE_COL_DASH)).NumberFormat = "@"
    ws.Range(ws.Cells(PE_OPS_DATA_START_ROW, PE_COL_OPER_SEQ), _
        ws.Cells(PE_OPS_DATA_START_ROW + PE_OPS_MAX_ROWS - 1, PE_OPS_LAST_COL)).ClearContents
End Sub

Private Function EditorNotesRange(ByVal ws As Worksheet) As Range
    Set EditorNotesRange = ws.Range( _
        ws.Cells(PE_NOTES_VALUE_ROW, PE_NOTES_VALUE_COL_START), _
        ws.Cells(PE_NOTES_VALUE_ROW_END, PE_NOTES_VALUE_COL_END))
End Function

Private Function ReadEditorNotes(ByVal ws As Worksheet) As String
    ReadEditorNotes = Trim$(CStr(EditorNotesRange(ws).Cells(1, 1).Value2))
End Function

Private Sub WriteEditorNotes(ByVal ws As Worksheet, ByVal notesText As String)
    EditorNotesRange(ws).Cells(1, 1).Value = notesText
End Sub

Private Sub ClearEditorNotes(ByVal ws As Worksheet)
    EditorNotesRange(ws).ClearContents
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
