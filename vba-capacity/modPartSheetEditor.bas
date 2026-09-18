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
    Dim errNumber As Long
    Dim errDescription As String
    Dim currentStep As String

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

    On Error GoTo FailLoad
    currentStep = "OptimizeExcel"
    OptimizeExcel True

    currentStep = "EnsureCacheSheetExists"
    EnsureCacheSheetExists
    currentStep = "ClearEditorDataRanges"
    ClearEditorDataRanges ws
    ws.Cells(PE_BASE_PART_ROW, PE_VALUE_COL).Value = basePartCode

    currentStep = "LoadMasterFields"
    LoadMasterFields ws, basePartCode
    currentStep = "LoadDashRows"
    LoadDashRows ws, basePartCode
    currentStep = "LoadRouteCardRows"
    LoadRouteCardRows ws, basePartCode
    currentStep = "LoadOperationRows"
    LoadOperationRows ws, basePartCode
    currentStep = "WriteEditorCache"
    WriteEditorCache basePartCode
    currentStep = "ApplyFactoryValidation"
    ApplyFactoryValidation ws
    currentStep = "ApplyOperationDropdowns"
    ApplyOperationDropdowns ws

    SetEditorStatus ws, "Loaded " & basePartCode & "."
    OptimizeExcel False
    Exit Sub

FailLoad:
    errNumber = Err.Number
    errDescription = Err.Description
    If Len(Trim$(errDescription)) = 0 Then errDescription = "(no description)"
    On Error Resume Next
    OptimizeExcel False
    On Error GoTo 0
    MsgBox "LoadPartToEditor failed at step:" & vbCrLf & currentStep & vbCrLf & vbCrLf & _
        "Error " & CStr(errNumber) & ": " & errDescription, vbExclamation
End Sub

' Button-safe entry point (OnAction can be unreliable with Optional-parameter Subs).
Public Sub LoadPartToEditorButton()
    LoadPartToEditor
End Sub

Public Sub SavePartFromEditor()
    Dim ws As Worksheet
    Dim basePartCode As String
    Dim factoryCode As String
    Dim factoryTbl As ListObject
    Dim fieldValues As Object
    Dim errNumber As Long
    Dim errDescription As String
    Dim currentStep As String

    Set ws = GetPartEditorWorksheet()
    If ws Is Nothing Then Exit Sub

    basePartCode = NormalizeCode(CStr(ws.Cells(PE_BASE_PART_ROW, PE_VALUE_COL).Value2))
    If Not ValidateRequiredCode(basePartCode, "Base Part") Then Exit Sub

    factoryCode = NormalizeCode(CStr(ws.Cells(PE_ROW_FACTORY, PE_VALUE_COL).Value2))
    If Not ValidateRequiredCode(factoryCode, "Factory") Then Exit Sub

    Set factoryTbl = FindTable(FACTORIES_TABLE_NAME)
    If Not ValidateForeignKeyExists(factoryTbl, COL_FACTORY_CODE, factoryCode, "Factory") Then Exit Sub

    On Error GoTo FailSave
    currentStep = "OptimizeExcel"
    OptimizeExcel True

    currentStep = "BuildFieldValues"
    Set fieldValues = NewFieldValuesDictionary()
    fieldValues(COL_BASE_PART_CODE) = basePartCode
    fieldValues(COL_PART_NAME) = Trim$(CStr(ws.Cells(PE_ROW_NAME, PE_VALUE_COL).Value2))
    fieldValues(COL_FACTORY_CODE) = factoryCode
    fieldValues(COL_ACTIVE) = ActiveFlagToCellValue(ReadEditorActiveFlag(ws))
    fieldValues(COL_PRODUCT_LINE) = Trim$(CStr(ws.Cells(PE_ROW_PRODUCT_LINE, PE_VALUE_COL).Value2))
    fieldValues(COL_NOTES) = ReadEditorNotes(ws)

    currentStep = "UpsertBasePart"
    UpsertRow FindTable(BASE_PARTS_TABLE_NAME), COL_BASE_PART_CODE, basePartCode, fieldValues
    currentStep = "SyncDashAssignments"
    SyncDashAssignments basePartCode, ws
    currentStep = "SyncOperationAssignments"
    SyncOperationAssignments basePartCode, ws
    SyncOperationRowDefaults ws
    currentStep = "WriteEditorCache"
    WriteEditorCache basePartCode

    SetEditorStatus ws, "Saved " & basePartCode & "."
    OptimizeExcel False
    Exit Sub

FailSave:
    errNumber = Err.Number
    errDescription = Err.Description
    If Len(Trim$(errDescription)) = 0 Then errDescription = "(no description)"
    On Error Resume Next
    OptimizeExcel False
    On Error GoTo 0
    MsgBox "SavePartFromEditor failed at step:" & vbCrLf & currentStep & vbCrLf & vbCrLf & _
        "Error " & CStr(errNumber) & ": " & errDescription, vbExclamation
End Sub

Public Sub SavePartFromEditorButton()
    SavePartFromEditor
End Sub

Public Sub ClearPartEditor()
    Dim ws As Worksheet
    Dim errNumber As Long
    Dim errDescription As String
    Dim currentStep As String

    Set ws = GetPartEditorWorksheet()
    If ws Is Nothing Then Exit Sub

    On Error GoTo FailClear
    currentStep = "OptimizeExcel"
    OptimizeExcel True

    currentStep = "ClearInput"
    ws.Cells(PE_INPUT_ROW, PE_VALUE_COL).ClearContents
    currentStep = "ClearEditorDataRanges"
    ClearEditorDataRanges ws
    currentStep = "ResetEditorCapacity"
    EnsurePartEditorOpsCapacity ws, PE_OPS_MAX_ROWS, True
    EnsurePartEditorRouteCapacity ws, PE_ROUTE_MAX_ROWS, True
    currentStep = "ClearEditorCache"
    ClearEditorCache
    currentStep = "SetEditorStatus"
    SetEditorStatus ws, vbNullString

    OptimizeExcel False
    Exit Sub

FailClear:
    errNumber = Err.Number
    errDescription = Err.Description
    If Len(Trim$(errDescription)) = 0 Then errDescription = "(no description)"
    On Error Resume Next
    OptimizeExcel False
    On Error GoTo 0
    MsgBox "ClearPartEditor failed at step:" & vbCrLf & currentStep & vbCrLf & vbCrLf & _
        "Error " & CStr(errNumber) & ": " & errDescription, vbExclamation
End Sub

Public Sub ClearPartEditorButton()
    ClearPartEditor
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
        SafeClearCellOrMerge ws.Cells(PE_ROW_NAME, PE_VALUE_COL)
        SafeClearCellOrMerge ws.Cells(PE_ROW_FACTORY, PE_VALUE_COL)
        ws.Cells(PE_ROW_ACTIVE, PE_VALUE_COL).Value = True
        SafeClearCellOrMerge ws.Cells(PE_ROW_PRODUCT_LINE, PE_VALUE_COL)
        ClearEditorNotes ws
        Exit Sub
    End If

    If TableHasColumn(tbl, COL_PART_NAME) Then
        SafeClearCellOrMerge ws.Cells(PE_ROW_NAME, PE_VALUE_COL)
        ws.Cells(PE_ROW_NAME, PE_VALUE_COL).Value = CStr(Nz(GetCellValueByListRow(tbl, listRowIndex, COL_PART_NAME)))
    Else
        SafeClearCellOrMerge ws.Cells(PE_ROW_NAME, PE_VALUE_COL)
    End If
    ws.Cells(PE_ROW_FACTORY, PE_VALUE_COL).Value = CStr(Nz(GetCellValueByListRow(tbl, listRowIndex, COL_FACTORY_CODE)))
    ws.Cells(PE_ROW_ACTIVE, PE_VALUE_COL).Value = IsActiveFlag(GetCellValueByListRow(tbl, listRowIndex, COL_ACTIVE))
    If TableHasColumn(tbl, COL_PRODUCT_LINE) Then
        ws.Cells(PE_ROW_PRODUCT_LINE, PE_VALUE_COL).Value = CStr(Nz(GetCellValueByListRow(tbl, listRowIndex, COL_PRODUCT_LINE)))
    Else
        SafeClearCellOrMerge ws.Cells(PE_ROW_PRODUCT_LINE, PE_VALUE_COL)
    End If
    WriteEditorNotes ws, CStr(Nz(GetCellValueByListRow(tbl, listRowIndex, COL_NOTES)))
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
        dashCell.Value = CStr(Nz(GetCellValueByListRow(tbl, rowIndex, COL_DASH_CONDITION)))
        If TableHasColumn(tbl, COL_SEPARATOR) Then
            ws.Cells(sheetRow, PE_COL_SEPARATOR).Value = CStr(Nz(GetCellValueByListRow(tbl, rowIndex, COL_SEPARATOR)))
        End If
        ws.Cells(sheetRow, PE_COL_DASH_ACTIVE).Value = IsActiveFlag(GetCellValueByListRow(tbl, rowIndex, COL_ACTIVE))
        ws.Cells(sheetRow, PE_COL_DASH_NOTES).Value = CStr(Nz(GetCellValueByListRow(tbl, rowIndex, COL_NOTES)))

        sheetRow = sheetRow + 1
        loadedCount = loadedCount + 1

ContinueDash:
    Next rowIndex
End Sub

Private Sub LoadRouteCardRows(ByVal ws As Worksheet, ByVal basePartCode As String)
    Dim tbl As ListObject
    Dim rowIndex As Long
    Dim sheetRow As Long
    Dim loadedCount As Long
    Dim assemblyNo As String
    Dim rowBasePart As String
    Dim dashCondition As String
    Dim separator As String
    Dim operSeq As String
    Dim operCode As String
    Dim matchCount As Long
    Dim matchIndex As Long
    Dim matchRows() As Variant
    Dim sortIndex As Long
    Dim swapIndex As Long
    Dim tempDash As String
    Dim tempSeq As String
    Dim tempCode As String
    Dim hasOperCode As Boolean

    ClearRouteCardRange ws

    Set tbl = FindTable(LINKED_ROUTE_CARD_TABLE)
    If tbl Is Nothing Or tbl.DataBodyRange Is Nothing Then
        EnsurePartEditorRouteCapacity ws, PE_ROUTE_MAX_ROWS, True
        Exit Sub
    End If
    If Not TableHasColumn(tbl, COL_ASSEMBLY_NO) Then
        EnsurePartEditorRouteCapacity ws, PE_ROUTE_MAX_ROWS, True
        Exit Sub
    End If
    If Not TableHasColumn(tbl, COL_OPER_SEQ_SOURCE) Then
        EnsurePartEditorRouteCapacity ws, PE_ROUTE_MAX_ROWS, True
        Exit Sub
    End If

    hasOperCode = TableHasColumn(tbl, COL_OPER_CODE_SOURCE)

    ' Count matching rows first so we can allocate once.
    matchCount = 0
    For rowIndex = 1 To tbl.ListRows.Count
        assemblyNo = Trim$(CStr(Nz(GetCellValueByListRow(tbl, rowIndex, COL_ASSEMBLY_NO))))
        If Len(assemblyNo) = 0 Then GoTo ContinueCount
        SplitAssemblyNoWithSeparator assemblyNo, rowBasePart, separator, dashCondition
        If ValuesMatchCode(NormalizeCode(rowBasePart), basePartCode) Then matchCount = matchCount + 1
ContinueCount:
    Next rowIndex

    If matchCount = 0 Then
        EnsurePartEditorRouteCapacity ws, PE_ROUTE_MAX_ROWS, True
        Exit Sub
    End If

    ReDim matchRows(1 To matchCount, 1 To 3)

    matchIndex = 0
    For rowIndex = 1 To tbl.ListRows.Count
        assemblyNo = Trim$(CStr(Nz(GetCellValueByListRow(tbl, rowIndex, COL_ASSEMBLY_NO))))
        If Len(assemblyNo) = 0 Then GoTo ContinueFill

        SplitAssemblyNoWithSeparator assemblyNo, rowBasePart, separator, dashCondition
        rowBasePart = NormalizeCode(rowBasePart)
        If Not ValuesMatchCode(rowBasePart, basePartCode) Then GoTo ContinueFill

        operSeq = Trim$(CStr(Nz(GetCellValueByListRow(tbl, rowIndex, COL_OPER_SEQ_SOURCE))))
        operCode = vbNullString
        If hasOperCode Then
            operCode = GetListRowCellText(tbl, rowIndex, COL_OPER_CODE_SOURCE)
        End If

        matchIndex = matchIndex + 1
        matchRows(matchIndex, 1) = dashCondition
        matchRows(matchIndex, 2) = operSeq
        matchRows(matchIndex, 3) = operCode

ContinueFill:
    Next rowIndex

    matchCount = matchIndex
    If matchCount = 0 Then
        EnsurePartEditorRouteCapacity ws, PE_ROUTE_MAX_ROWS, True
        Exit Sub
    End If

    ' Sort by dash condition, then OPER SEQ.
    For sortIndex = 1 To matchCount - 1
        For swapIndex = sortIndex + 1 To matchCount
            If RouteRowSortKey2(CStr(matchRows(swapIndex, 1)), CStr(matchRows(swapIndex, 2))) < _
               RouteRowSortKey2(CStr(matchRows(sortIndex, 1)), CStr(matchRows(sortIndex, 2))) Then
                tempDash = CStr(matchRows(sortIndex, 1))
                tempSeq = CStr(matchRows(sortIndex, 2))
                tempCode = CStr(matchRows(sortIndex, 3))
                matchRows(sortIndex, 1) = matchRows(swapIndex, 1)
                matchRows(sortIndex, 2) = matchRows(swapIndex, 2)
                matchRows(sortIndex, 3) = matchRows(swapIndex, 3)
                matchRows(swapIndex, 1) = tempDash
                matchRows(swapIndex, 2) = tempSeq
                matchRows(swapIndex, 3) = tempCode
            End If
        Next swapIndex
    Next sortIndex

    EnsurePartEditorRouteCapacity ws, matchCount + PE_BLOCK_SPARE_ROWS, True

    sheetRow = PE_ROUTE_DATA_START_ROW
    loadedCount = 0
    For sortIndex = 1 To matchCount
        If sheetRow > PartEditorRouteLastRow() Then Exit For
        ws.Cells(sheetRow, PE_COL_ROUTE_DASH).NumberFormat = "@"
        ws.Cells(sheetRow, PE_COL_ROUTE_DASH).Value = CStr(matchRows(sortIndex, 1))
        ws.Cells(sheetRow, PE_COL_ROUTE_OPER_SEQ).Value = CStr(matchRows(sortIndex, 2))
        ws.Cells(sheetRow, PE_COL_ROUTE_OPER_CODE).NumberFormat = "@"
        ws.Cells(sheetRow, PE_COL_ROUTE_OPER_CODE).Value = CStr(matchRows(sortIndex, 3))
        sheetRow = sheetRow + 1
        loadedCount = loadedCount + 1
    Next sortIndex
End Sub

Private Function RouteRowSortKey2(ByVal dashCondition As String, ByVal operSeq As String) As String
    Dim seqNumber As Double
    Dim dashKey As String

    dashCondition = Trim$(dashCondition)
    operSeq = Trim$(operSeq)

    ' Keep dash text sortable with leading zeros preserved as text.
    dashKey = dashCondition
    If IsNumeric(dashCondition) Then
        If Len(dashCondition) < 10 Then
            dashKey = String$(10 - Len(dashCondition), "0") & dashCondition
        Else
            dashKey = dashCondition
        End If
    End If

    If IsNumeric(operSeq) Then
        seqNumber = CDbl(operSeq)
        RouteRowSortKey2 = dashKey & "|" & Format$(seqNumber, "0000000000.0000")
    Else
        RouteRowSortKey2 = dashKey & "|" & operSeq
    End If
End Function

Private Function GetListRowCellText( _
    ByVal tbl As ListObject, _
    ByVal listRowIndex As Long, _
    ByVal columnName As String) As String

    Dim cell As Range
    Dim textValue As String

    On Error GoTo FailText
    Set cell = tbl.ListRows(listRowIndex).Range.Cells(1, TableColumnIndex(tbl, columnName))
    textValue = Trim$(cell.Text)
    If Len(textValue) > 0 Then
        GetListRowCellText = textValue
    Else
        GetListRowCellText = Trim$(CStr(Nz(cell.Value2)))
    End If
    Exit Function

FailText:
    GetListRowCellText = vbNullString
End Function

Private Sub ClearRouteCardRange(ByVal ws As Worksheet)
    SafeClearRange ws.Range( _
        ws.Cells(PE_ROUTE_DATA_START_ROW, PE_COL_ROUTE_DASH), _
        ws.Cells(PartEditorRouteLastRow(), PE_COL_ROUTE_OPER_CODE))
    SafeNumberFormat ws.Range( _
        ws.Cells(PE_ROUTE_DATA_START_ROW, PE_COL_ROUTE_DASH), _
        ws.Cells(PartEditorRouteLastRow(), PE_COL_ROUTE_DASH)), "@"
    SafeNumberFormat ws.Range( _
        ws.Cells(PE_ROUTE_DATA_START_ROW, PE_COL_ROUTE_OPER_CODE), _
        ws.Cells(PartEditorRouteLastRow(), PE_COL_ROUTE_OPER_CODE)), "@"
End Sub

Private Sub LoadOperationRows(ByVal ws As Worksheet, ByVal basePartCode As String)
    Dim tbl As ListObject
    Dim operSeqValues As Variant
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim sheetRow As Long
    Dim loadedCount As Long
    Dim matchCount As Long
    Dim sortIndex As Long
    Dim matchRows() As Variant
    Dim operSeq As String
    Dim opLine As Long
    Dim showAvgHours As Boolean
    Dim showAvgEx As Boolean
    Dim equipmentCode As String
    Dim processTypeCode As String
    Dim madeInFfa As String
    Dim partFactory As String

    Set tbl = FindTable(PART_OPERATIONS_TABLE_NAME)
    If tbl Is Nothing Or tbl.DataBodyRange Is Nothing Then
        EnsurePartEditorOpsCapacity ws, PE_OPS_MAX_ROWS, True
        SyncOperationRowDefaults ws
        Exit Sub
    End If

    operSeqValues = ListColumnValues2D(tbl, COL_OPER_SEQ)
    If Not IsArray(operSeqValues) Then
        EnsurePartEditorOpsCapacity ws, PE_OPS_MAX_ROWS, True
        SyncOperationRowDefaults ws
        Exit Sub
    End If

    rowCount = UBound(operSeqValues, 1)
    partFactory = NormalizeCode(CStr(ws.Cells(PE_ROW_FACTORY, PE_VALUE_COL).Value2))
    matchCount = 0

    For rowIndex = 1 To rowCount
        If ValuesMatchCode(GetCellValueByListRow(tbl, rowIndex, COL_BASE_PART_CODE), basePartCode) Then
            matchCount = matchCount + 1
        End If
    Next rowIndex

    If matchCount = 0 Then
        EnsurePartEditorOpsCapacity ws, PE_OPS_MAX_ROWS, True
        SyncOperationRowDefaults ws
        Exit Sub
    End If

    ReDim matchRows(1 To matchCount, 1 To 1)
    matchCount = 0
    For rowIndex = 1 To rowCount
        If Not ValuesMatchCode(GetCellValueByListRow(tbl, rowIndex, COL_BASE_PART_CODE), basePartCode) Then GoTo ContinueCollect
        matchCount = matchCount + 1
        matchRows(matchCount, 1) = rowIndex

ContinueCollect:
    Next rowIndex

    If matchCount = 0 Then
        EnsurePartEditorOpsCapacity ws, PE_OPS_MAX_ROWS, True
        SyncOperationRowDefaults ws
        Exit Sub
    End If

    ' Sort by Oper Seq then Op Line.
    SortOperationListRows tbl, matchRows, matchCount
    EnsurePartEditorOpsCapacity ws, matchCount + PE_BLOCK_SPARE_ROWS, True

    sheetRow = PE_OPS_DATA_START_ROW
    loadedCount = 0

    For sortIndex = 1 To matchCount
        rowIndex = CLng(matchRows(sortIndex, 1))
        If sheetRow > PartEditorOpsLastRow() Then Exit For

        operSeq = Trim$(CStr(Nz(GetCellValueByListRow(tbl, rowIndex, COL_OPER_SEQ))))
        opLine = ReadOpLineValue(GetCellValueByListRow(tbl, rowIndex, COL_OP_LINE))

        ws.Cells(sheetRow, PE_COL_OPER_SEQ).Value = operSeq
        ws.Cells(sheetRow, PE_COL_OP_LINE).Value = opLine

        ws.Cells(sheetRow, PE_COL_OPER_CODE).NumberFormat = "@"
        ws.Cells(sheetRow, PE_COL_OPER_CODE).Value = ReadOperationCodeFromRow(tbl, rowIndex)

        madeInFfa = vbNullString
        If TableHasColumn(tbl, COL_MADE_IN_FFA) Then
            madeInFfa = NormalizeCode(CStr(Nz(GetCellValueByListRow(tbl, rowIndex, COL_MADE_IN_FFA))))
        End If
        If Len(madeInFfa) = 0 Then madeInFfa = partFactory
        ws.Cells(sheetRow, PE_COL_MADE_IN_FFA).Value = madeInFfa

        equipmentCode = vbNullString
        If TableHasColumn(tbl, COL_EQUIPMENT_CODE) Then
            equipmentCode = NormalizeCode(CStr(Nz(GetCellValueByListRow(tbl, rowIndex, COL_EQUIPMENT_CODE))))
        End If
        ws.Cells(sheetRow, PE_COL_EQUIPMENT).Value = equipmentCode

        processTypeCode = vbNullString
        If TableHasColumn(tbl, COL_PROCESS_TYPE_CODE) Then
            processTypeCode = NormalizeCode(CStr(Nz(GetCellValueByListRow(tbl, rowIndex, COL_PROCESS_TYPE_CODE))))
        End If
        ws.Cells(sheetRow, PE_COL_PROCESS_TYPE).Value = processTypeCode

        If TableHasColumn(tbl, COL_PROCESS_HOURS) Then
            ws.Cells(sheetRow, PE_COL_PROCESS_HOURS).Value = GetCellValueByListRow(tbl, rowIndex, COL_PROCESS_HOURS)
        End If
        If TableHasColumn(tbl, COL_MANUAL_AVG_EX) Then
            ws.Cells(sheetRow, PE_COL_MANUAL_AVG_EX).Value = GetCellValueByListRow(tbl, rowIndex, COL_MANUAL_AVG_EX)
        End If
        If TableHasColumn(tbl, COL_BATCH_SIZE) Then
            ws.Cells(sheetRow, PE_COL_BATCH_SIZE).Value = GetCellValueByListRow(tbl, rowIndex, COL_BATCH_SIZE)
        End If

        showAvgHours = True
        If TableHasColumn(tbl, COL_USE_AVG_HOURS) Then
            If Not IsBlankCellValue(GetCellValueByListRow(tbl, rowIndex, COL_USE_AVG_HOURS)) Then
                showAvgHours = IsActiveFlag(GetCellValueByListRow(tbl, rowIndex, COL_USE_AVG_HOURS))
            End If
        ElseIf TableHasColumn(tbl, COL_SHOW_AVG_HOURS) Then
            If Not IsBlankCellValue(GetCellValueByListRow(tbl, rowIndex, COL_SHOW_AVG_HOURS)) Then
                showAvgHours = IsActiveFlag(GetCellValueByListRow(tbl, rowIndex, COL_SHOW_AVG_HOURS))
            End If
        End If
        ws.Cells(sheetRow, PE_COL_USE_AVG_HOURS).Value = showAvgHours

        showAvgEx = True
        If TableHasColumn(tbl, COL_USE_AVG_EX) Then
            If Not IsBlankCellValue(GetCellValueByListRow(tbl, rowIndex, COL_USE_AVG_EX)) Then
                showAvgEx = IsActiveFlag(GetCellValueByListRow(tbl, rowIndex, COL_USE_AVG_EX))
            End If
        ElseIf TableHasColumn(tbl, COL_SHOW_AVG_EX) Then
            If Not IsBlankCellValue(GetCellValueByListRow(tbl, rowIndex, COL_SHOW_AVG_EX)) Then
                showAvgEx = IsActiveFlag(GetCellValueByListRow(tbl, rowIndex, COL_SHOW_AVG_EX))
            End If
        End If
        ws.Cells(sheetRow, PE_COL_USE_AVG_EX).Value = showAvgEx

        ws.Cells(sheetRow, PE_COL_OPER_ACTIVE).Value = IsActiveFlag(GetCellValueByListRow(tbl, rowIndex, COL_ACTIVE))
        ws.Cells(sheetRow, PE_COL_OPER_NOTES).Value = CStr(Nz(GetCellValueByListRow(tbl, rowIndex, COL_NOTES)))

        ApplyAveragesForOperationRow ws, sheetRow, basePartCode

        sheetRow = sheetRow + 1
        loadedCount = loadedCount + 1
    Next sortIndex

    SyncOperationRowDefaults ws
End Sub

Private Function ReadOperationCodeFromRow(ByVal tbl As ListObject, ByVal rowIndex As Long) As String
    Dim codeText As String

    codeText = vbNullString
    If TableHasColumn(tbl, COL_OPER_CODE) Then
        codeText = GetListRowCellText(tbl, rowIndex, COL_OPER_CODE)
    End If
    If Len(Trim$(codeText)) = 0 And TableHasColumn(tbl, COL_OPERATION_NAME) Then
        codeText = GetListRowCellText(tbl, rowIndex, COL_OPERATION_NAME)
    End If
    ReadOperationCodeFromRow = codeText
End Function

Private Sub SortOperationListRows(ByVal tbl As ListObject, ByRef matchRows As Variant, ByVal matchCount As Long)
    Dim i As Long
    Dim j As Long
    Dim leftRow As Long
    Dim rightRow As Long
    Dim leftSeq As String
    Dim rightSeq As String
    Dim leftLine As Long
    Dim rightLine As Long
    Dim swapValue As Variant

    For i = 1 To matchCount - 1
        For j = i + 1 To matchCount
            leftRow = CLng(matchRows(i, 1))
            rightRow = CLng(matchRows(j, 1))
            leftSeq = NormalizeOperSeqKey(CStr(Nz(GetCellValueByListRow(tbl, leftRow, COL_OPER_SEQ))))
            rightSeq = NormalizeOperSeqKey(CStr(Nz(GetCellValueByListRow(tbl, rightRow, COL_OPER_SEQ))))
            leftLine = ReadOpLineValue(GetCellValueByListRow(tbl, leftRow, COL_OP_LINE))
            rightLine = ReadOpLineValue(GetCellValueByListRow(tbl, rightRow, COL_OP_LINE))

            If CompareOperSeqThenLine(leftSeq, leftLine, rightSeq, rightLine) > 0 Then
                swapValue = matchRows(i, 1)
                matchRows(i, 1) = matchRows(j, 1)
                matchRows(j, 1) = swapValue
            End If
        Next j
    Next i
End Sub

Private Function CompareOperSeqThenLine( _
    ByVal leftSeq As String, _
    ByVal leftLine As Long, _
    ByVal rightSeq As String, _
    ByVal rightLine As Long) As Long

    If IsNumeric(leftSeq) And IsNumeric(rightSeq) Then
        If CDbl(leftSeq) < CDbl(rightSeq) Then
            CompareOperSeqThenLine = -1
            Exit Function
        ElseIf CDbl(leftSeq) > CDbl(rightSeq) Then
            CompareOperSeqThenLine = 1
            Exit Function
        End If
    Else
        CompareOperSeqThenLine = StrComp(leftSeq, rightSeq, vbTextCompare)
        If CompareOperSeqThenLine <> 0 Then Exit Function
    End If

    If leftLine < rightLine Then
        CompareOperSeqThenLine = -1
    ElseIf leftLine > rightLine Then
        CompareOperSeqThenLine = 1
    Else
        CompareOperSeqThenLine = 0
    End If
End Function

Private Function ReadOpLineValue(ByVal rawValue As Variant) As Long
    If IsError(rawValue) Then
        ReadOpLineValue = 1
    ElseIf IsEmpty(rawValue) Or IsNull(rawValue) Then
        ReadOpLineValue = 1
    ElseIf Len(Trim$(CStr(rawValue))) = 0 Then
        ReadOpLineValue = 1
    ElseIf IsNumeric(rawValue) Then
        ReadOpLineValue = CLng(CDbl(rawValue))
        If ReadOpLineValue < 1 Then ReadOpLineValue = 1
    Else
        ReadOpLineValue = 1
    End If
End Function

' Called from ThisWorkbook SheetChange for cascading dropdowns and avg toggles.
Public Sub HandlePartEditorSheetChange(ByVal Target As Range)
    Dim ws As Worksheet
    Dim factoryCell As Range
    Dim opsEquipment As Range
    Dim opsMadeInFfa As Range
    Dim opsSeq As Range
    Dim opsBlock As Range
    Dim routeBlock As Range
    Dim changedRow As Long
    Dim basePartCode As String
    Dim opsLastRow As Long
    Dim routeLastRow As Long

    If Target Is Nothing Then Exit Sub
    Set ws = Target.Worksheet
    If StrComp(ws.Name, PART_EDITOR_SHEET_NAME, vbTextCompare) <> 0 Then Exit Sub

    opsLastRow = PartEditorOpsLastRow()
    routeLastRow = PartEditorRouteLastRow()

    Set factoryCell = ws.Cells(PE_ROW_FACTORY, PE_VALUE_COL)
    Set opsEquipment = ws.Range( _
        ws.Cells(PE_OPS_DATA_START_ROW, PE_COL_EQUIPMENT), _
        ws.Cells(opsLastRow, PE_COL_EQUIPMENT))
    Set opsMadeInFfa = ws.Range( _
        ws.Cells(PE_OPS_DATA_START_ROW, PE_COL_MADE_IN_FFA), _
        ws.Cells(opsLastRow, PE_COL_MADE_IN_FFA))
    Set opsSeq = ws.Range( _
        ws.Cells(PE_OPS_DATA_START_ROW, PE_COL_OPER_SEQ), _
        ws.Cells(opsLastRow, PE_COL_OPER_SEQ))
    Set opsBlock = ws.Range( _
        ws.Cells(PE_OPS_DATA_START_ROW, PE_COL_OPER_SEQ), _
        ws.Cells(opsLastRow, PE_OPS_LAST_COL))
    Set routeBlock = ws.Range( _
        ws.Cells(PE_ROUTE_DATA_START_ROW, PE_COL_ROUTE_DASH), _
        ws.Cells(routeLastRow, PE_COL_ROUTE_OPER_CODE))

    On Error GoTo CleanUp
    Application.EnableEvents = False

    If Not Intersect(Target, factoryCell) Is Nothing Then
        ApplyOperationDropdowns ws
        ClearInvalidEquipmentAndProcess ws
        GoTo CleanUp
    End If

    If Not Intersect(Target, opsMadeInFfa) Is Nothing Then
        changedRow = Intersect(Target, opsMadeInFfa).Row
        ApplyEquipmentValidationForRow ws, changedRow
        ClearInvalidEquipmentAndProcessForRow ws, changedRow
        SyncOperationRowDefaultsForRow ws, changedRow
        GrowPartEditorOpsIfNeeded ws, changedRow
        GoTo CleanUp
    End If

    If Not Intersect(Target, opsEquipment) Is Nothing Then
        changedRow = Intersect(Target, opsEquipment).Row
        ApplyProcessTypeValidationForRow ws, changedRow
        ClearInvalidProcessTypeForRow ws, changedRow
        SyncOperationRowDefaultsForRow ws, changedRow
        GrowPartEditorOpsIfNeeded ws, changedRow
        GoTo CleanUp
    End If

    If Not Intersect(Target, opsBlock) Is Nothing Then
        basePartCode = NormalizeCode(CStr(ws.Cells(PE_BASE_PART_ROW, PE_VALUE_COL).Value2))
        For changedRow = Intersect(Target, opsBlock).Row To Intersect(Target, opsBlock).Row + Intersect(Target, opsBlock).Rows.Count - 1
            If changedRow >= PE_OPS_DATA_START_ROW And changedRow <= opsLastRow Then
                SyncOperationRowDefaultsForRow ws, changedRow
                If Not Intersect(Target, opsSeq) Is Nothing Then
                    If Not Intersect(ws.Cells(changedRow, PE_COL_OPER_SEQ), opsSeq) Is Nothing Then
                        ApplyAveragesForOperationRow ws, changedRow, basePartCode
                    End If
                End If
                GrowPartEditorOpsIfNeeded ws, changedRow
            End If
        Next changedRow
        GoTo CleanUp
    End If

    If Not Intersect(Target, routeBlock) Is Nothing Then
        For changedRow = Intersect(Target, routeBlock).Row To Intersect(Target, routeBlock).Row + Intersect(Target, routeBlock).Rows.Count - 1
            If changedRow >= PE_ROUTE_DATA_START_ROW And changedRow <= routeLastRow Then
                GrowPartEditorRouteIfNeeded ws, changedRow
            End If
        Next changedRow
    End If

CleanUp:
    Application.EnableEvents = True
End Sub

Private Sub SyncOperationRowDefaults(ByVal ws As Worksheet)
    Dim rowIndex As Long

    If ws Is Nothing Then Exit Sub
    For rowIndex = PE_OPS_DATA_START_ROW To PartEditorOpsLastRow()
        SyncOperationRowDefaultsForRow ws, rowIndex
    Next rowIndex
End Sub

Private Sub SyncOperationRowDefaultsForRow(ByVal ws As Worksheet, ByVal sheetRow As Long)
    If OperationEditorRowHasData(ws, sheetRow) Then
        If IsBlankCellValue(ws.Cells(sheetRow, PE_COL_OP_LINE).Value2) Then
            ws.Cells(sheetRow, PE_COL_OP_LINE).Value = 1
        End If
        If IsBlankCellValue(ws.Cells(sheetRow, PE_COL_USE_AVG_HOURS).Value2) Then
            ws.Cells(sheetRow, PE_COL_USE_AVG_HOURS).Value = True
        End If
        If IsBlankCellValue(ws.Cells(sheetRow, PE_COL_USE_AVG_EX).Value2) Then
            ws.Cells(sheetRow, PE_COL_USE_AVG_EX).Value = True
        End If
    Else
        ws.Cells(sheetRow, PE_COL_USE_AVG_HOURS).ClearContents
        ws.Cells(sheetRow, PE_COL_USE_AVG_EX).ClearContents
        If IsBlankCellValue(ws.Cells(sheetRow, PE_COL_OPER_SEQ).Value2) Then
            ws.Cells(sheetRow, PE_COL_OP_LINE).ClearContents
        End If
    End If
End Sub

Private Function OperationEditorRowHasData(ByVal ws As Worksheet, ByVal sheetRow As Long) As Boolean
    OperationEditorRowHasData = _
        Not IsBlankCellValue(ws.Cells(sheetRow, PE_COL_OPER_SEQ).Value2) _
        Or Not IsBlankCellValue(ws.Cells(sheetRow, PE_COL_OPER_CODE).Value2) _
        Or Not IsBlankCellValue(ws.Cells(sheetRow, PE_COL_MADE_IN_FFA).Value2) _
        Or Not IsBlankCellValue(ws.Cells(sheetRow, PE_COL_EQUIPMENT).Value2) _
        Or Not IsBlankCellValue(ws.Cells(sheetRow, PE_COL_PROCESS_TYPE).Value2) _
        Or Not IsBlankCellValue(ws.Cells(sheetRow, PE_COL_PROCESS_HOURS).Value2) _
        Or Not IsBlankCellValue(ws.Cells(sheetRow, PE_COL_MANUAL_AVG_EX).Value2) _
        Or Not IsBlankCellValue(ws.Cells(sheetRow, PE_COL_BATCH_SIZE).Value2) _
        Or Not IsBlankCellValue(ws.Cells(sheetRow, PE_COL_OPER_NOTES).Value2)
End Function

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

    ' Calculated averages are always shown; Use Avg Hours/Ex are preference flags only.
    avgHours = AvgProcessHoursByBasePartAndOp(basePartCode, operSeq)
    ws.Cells(sheetRow, PE_COL_AVG_HOURS).Value = FormatAverageDisplay(avgHours)

    avgEx = AvgExByBasePartAndOp(basePartCode, operSeq)
    ws.Cells(sheetRow, PE_COL_AVG_EX).Value = FormatAverageDisplay(avgEx)
End Sub

Public Sub ApplyOperationDropdowns(ByVal ws As Worksheet)
    If ws Is Nothing Then Set ws = GetPartEditorWorksheet()
    If ws Is Nothing Then Exit Sub

    ApplyOperationDropdownsForRows ws, PE_OPS_DATA_START_ROW, PartEditorOpsLastRow()
End Sub

Private Sub ApplyOperationDropdownsForRows(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim rowIndex As Long
    Dim madeInRange As Range
    Dim factoryCodes As String

    If lastRow < firstRow Then Exit Sub

    factoryCodes = BuildFactoryValidationList()
    Set madeInRange = ws.Range( _
        ws.Cells(firstRow, PE_COL_MADE_IN_FFA), _
        ws.Cells(lastRow, PE_COL_MADE_IN_FFA))
    ApplyListValidation madeInRange, factoryCodes

    For rowIndex = firstRow To lastRow
        ApplyEquipmentValidationForRow ws, rowIndex
        ApplyProcessTypeValidationForRow ws, rowIndex
    Next rowIndex
End Sub

Private Function OperationFactoryCodeForRow(ByVal ws As Worksheet, ByVal sheetRow As Long) As String
    Dim madeInFfa As String

    madeInFfa = NormalizeCode(CStr(ws.Cells(sheetRow, PE_COL_MADE_IN_FFA).Value2))
    If Len(madeInFfa) > 0 Then
        OperationFactoryCodeForRow = madeInFfa
    Else
        OperationFactoryCodeForRow = NormalizeCode(CStr(ws.Cells(PE_ROW_FACTORY, PE_VALUE_COL).Value2))
    End If
End Function

Private Sub ApplyEquipmentValidationForRow(ByVal ws As Worksheet, ByVal sheetRow As Long)
    Dim factoryCode As String
    Dim equipmentList As String

    factoryCode = OperationFactoryCodeForRow(ws, sheetRow)
    equipmentList = BuildEquipmentValidationList(factoryCode)
    ApplyListValidation ws.Cells(sheetRow, PE_COL_EQUIPMENT), equipmentList
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

Private Sub ClearInvalidEquipmentAndProcessForRow(ByVal ws As Worksheet, ByVal sheetRow As Long)
    Dim factoryCode As String
    Dim equipmentList As String
    Dim equipmentCode As String

    factoryCode = OperationFactoryCodeForRow(ws, sheetRow)
    equipmentList = "," & UCase$(BuildEquipmentValidationList(factoryCode)) & ","
    equipmentCode = NormalizeCode(CStr(ws.Cells(sheetRow, PE_COL_EQUIPMENT).Value2))

    If Len(equipmentCode) > 0 Then
        If InStr(1, equipmentList, "," & UCase$(equipmentCode) & ",", vbTextCompare) = 0 Then
            ws.Cells(sheetRow, PE_COL_EQUIPMENT).ClearContents
            ws.Cells(sheetRow, PE_COL_PROCESS_TYPE).ClearContents
        Else
            ClearInvalidProcessTypeForRow ws, sheetRow
        End If
    End If

    ApplyEquipmentValidationForRow ws, sheetRow
    ApplyProcessTypeValidationForRow ws, sheetRow
End Sub

Private Sub ClearInvalidEquipmentAndProcess(ByVal ws As Worksheet)
    Dim rowIndex As Long

    For rowIndex = PE_OPS_DATA_START_ROW To PartEditorOpsLastRow()
        ClearInvalidEquipmentAndProcessForRow ws, rowIndex
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
    ' Excel list validation Formula1 max length is 255 characters.
    If Len(csvList) > 255 Then csvList = Left$(csvList, 255)

    On Error Resume Next
    targetRange.Validation.Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, _
        Formula1:=csvList
    On Error GoTo 0
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
        equipmentCode = NormalizeCode(CStr(Nz(GetCellValueByListRow(junctionTbl, rowIndex, COL_EQUIPMENT_CODE))))
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

        processCode = NormalizeCode(CStr(Nz(GetCellValueByListRow(junctionTbl, rowIndex, COL_PROCESS_TYPE_CODE))))
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
    Dim currentRows As Object
    Dim seenKeys As Object
    Dim operKey As Variant
    Dim rowData As Variant
    Dim fieldValues As Object
    Dim operSeq As String
    Dim opLine As Long
    Dim listRowIndex As Long
    Dim rowIndex As Long
    Dim storedSeq As String
    Dim storedLine As Long
    Dim storedKey As String

    Set tbl = FindTable(PART_OPERATIONS_TABLE_NAME)
    If tbl Is Nothing Then Exit Sub

    Set currentRows = ReadSheetOperationRows(ws)
    Set seenKeys = CreateObject("Scripting.Dictionary")
    seenKeys.CompareMode = vbTextCompare

    ' Drop stale rows and extra copies for this part, keeping one row per current key.
    If Not tbl.DataBodyRange Is Nothing Then
        For rowIndex = tbl.ListRows.Count To 1 Step -1
            If Not ValuesMatchCode(GetCellValueByListRow(tbl, rowIndex, COL_BASE_PART_CODE), basePartCode) Then GoTo ContinueDeleteOp

            storedSeq = Trim$(CStr(Nz(GetCellValueByListRow(tbl, rowIndex, COL_OPER_SEQ))))
            storedLine = 1
            If TableHasColumn(tbl, COL_OP_LINE) Then
                storedLine = ReadOpLineValue(GetCellValueByListRow(tbl, rowIndex, COL_OP_LINE))
            End If
            storedKey = BuildOperationCacheKey(storedSeq, storedLine)

            If Len(storedSeq) = 0 Or Not currentRows.Exists(storedKey) Then
                tbl.ListRows(rowIndex).Delete
            ElseIf seenKeys.Exists(storedKey) Then
                tbl.ListRows(rowIndex).Delete
            Else
                seenKeys.Add storedKey, True
            End If

ContinueDeleteOp:
        Next rowIndex
    End If

    For Each operKey In currentRows.Keys
        rowData = currentRows(operKey)
        operSeq = CStr(rowData(1))
        opLine = ReadOpLineValue(rowData(2))

        Set fieldValues = NewFieldValuesDictionary()
        fieldValues(COL_BASE_PART_CODE) = basePartCode
        fieldValues(COL_OPER_SEQ) = operSeq
        If TableHasColumn(tbl, COL_OP_LINE) Then fieldValues(COL_OP_LINE) = opLine
        If TableHasColumn(tbl, COL_OPER_CODE) Then
            fieldValues(COL_OPER_CODE) = CStr(rowData(3))
        ElseIf TableHasColumn(tbl, COL_OPERATION_NAME) Then
            fieldValues(COL_OPERATION_NAME) = CStr(rowData(3))
        End If
        If TableHasColumn(tbl, COL_MADE_IN_FFA) Then fieldValues(COL_MADE_IN_FFA) = CStr(rowData(4))
        fieldValues(COL_EQUIPMENT_CODE) = CStr(rowData(5))
        fieldValues(COL_PROCESS_TYPE_CODE) = CStr(rowData(6))
        If TableHasColumn(tbl, COL_PROCESS_HOURS) Then fieldValues(COL_PROCESS_HOURS) = rowData(7)
        If TableHasColumn(tbl, COL_MANUAL_AVG_EX) Then fieldValues(COL_MANUAL_AVG_EX) = rowData(8)
        If TableHasColumn(tbl, COL_BATCH_SIZE) Then fieldValues(COL_BATCH_SIZE) = rowData(9)
        If TableHasColumn(tbl, COL_USE_AVG_HOURS) Then
            fieldValues(COL_USE_AVG_HOURS) = ActiveFlagToCellValue(IsActiveFlag(rowData(10)))
        ElseIf TableHasColumn(tbl, COL_SHOW_AVG_HOURS) Then
            fieldValues(COL_SHOW_AVG_HOURS) = ActiveFlagToCellValue(IsActiveFlag(rowData(10)))
        End If
        If TableHasColumn(tbl, COL_USE_AVG_EX) Then
            fieldValues(COL_USE_AVG_EX) = ActiveFlagToCellValue(IsActiveFlag(rowData(11)))
        ElseIf TableHasColumn(tbl, COL_SHOW_AVG_EX) Then
            fieldValues(COL_SHOW_AVG_EX) = ActiveFlagToCellValue(IsActiveFlag(rowData(11)))
        End If
        fieldValues(COL_ACTIVE) = ActiveFlagToCellValue(IsActiveFlag(rowData(12)))
        fieldValues(COL_NOTES) = CStr(rowData(13))

        listRowIndex = FindOperationListRow(tbl, basePartCode, operSeq, opLine)
        If listRowIndex = 0 Then listRowIndex = GetOrCreateListRowIndex(tbl)
        ApplyFieldValuesToListRow tbl, listRowIndex, fieldValues
        WriteOperCodeText tbl, basePartCode, operSeq, opLine, CStr(rowData(3))
    Next operKey
End Sub

Private Function FindOperationListRow( _
    ByVal tbl As ListObject, _
    ByVal basePartCode As String, _
    ByVal operSeq As String, _
    ByVal opLine As Long) As Long

    Dim rowIndex As Long
    Dim storedLine As Long

    FindOperationListRow = 0
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    For rowIndex = 1 To tbl.ListRows.Count
        If Not ValuesMatchCode(GetCellValueByListRow(tbl, rowIndex, COL_BASE_PART_CODE), basePartCode) Then GoTo ContinueFindOp
        If Not OpSequencesMatch(GetCellValueByListRow(tbl, rowIndex, COL_OPER_SEQ), operSeq) Then GoTo ContinueFindOp

        storedLine = 1
        If TableHasColumn(tbl, COL_OP_LINE) Then
            storedLine = ReadOpLineValue(GetCellValueByListRow(tbl, rowIndex, COL_OP_LINE))
        End If
        If storedLine = opLine Then
            FindOperationListRow = tbl.ListRows(rowIndex).Index
            Exit Function
        End If

ContinueFindOp:
    Next rowIndex
End Function

Private Sub WriteOperCodeText( _
    ByVal tbl As ListObject, _
    ByVal basePartCode As String, _
    ByVal operSeq As String, _
    ByVal opLine As Long, _
    ByVal operCode As String)

    Dim listRowIndex As Long
    Dim codeCell As Range
    Dim columnName As String

    If TableHasColumn(tbl, COL_OPER_CODE) Then
        columnName = COL_OPER_CODE
    ElseIf TableHasColumn(tbl, COL_OPERATION_NAME) Then
        columnName = COL_OPERATION_NAME
    Else
        Exit Sub
    End If

    If TableHasColumn(tbl, COL_OP_LINE) Then
        listRowIndex = FindOperationListRow(tbl, basePartCode, operSeq, opLine)
    Else
        listRowIndex = FindJunctionListRow(tbl, COL_BASE_PART_CODE, basePartCode, COL_OPER_SEQ, operSeq)
    End If
    If listRowIndex <= 0 Then Exit Sub

    Set codeCell = tbl.ListRows(listRowIndex).Range.Cells(1, TableColumnIndex(tbl, columnName))
    codeCell.NumberFormat = "@"
    codeCell.Value = operCode
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
    wsCache.Cells(CACHE_OPS_START_ROW - 1, 2).Value = COL_OP_LINE
    wsCache.Cells(CACHE_OPS_START_ROW - 1, 3).Value = COL_OPER_CODE
    wsCache.Cells(CACHE_OPS_START_ROW - 1, 4).Value = COL_MADE_IN_FFA
    wsCache.Cells(CACHE_OPS_START_ROW - 1, 5).Value = COL_EQUIPMENT_CODE
    wsCache.Cells(CACHE_OPS_START_ROW - 1, 6).Value = COL_PROCESS_TYPE_CODE
    wsCache.Cells(CACHE_OPS_START_ROW - 1, 7).Value = COL_PROCESS_HOURS
    wsCache.Cells(CACHE_OPS_START_ROW - 1, 8).Value = COL_MANUAL_AVG_EX
    wsCache.Cells(CACHE_OPS_START_ROW - 1, 9).Value = COL_BATCH_SIZE
    wsCache.Cells(CACHE_OPS_START_ROW - 1, 10).Value = COL_USE_AVG_HOURS
    wsCache.Cells(CACHE_OPS_START_ROW - 1, 11).Value = COL_USE_AVG_EX
    wsCache.Cells(CACHE_OPS_START_ROW - 1, 12).Value = COL_ACTIVE
    wsCache.Cells(CACHE_OPS_START_ROW - 1, 13).Value = COL_NOTES

    Set opRows = ReadSheetOperationRows(wsEditor)
    sheetRow = CACHE_OPS_START_ROW
    For Each operKey In opRows.Keys
        rowData = opRows(operKey)
        wsCache.Cells(sheetRow, 1).Value = CStr(rowData(1))
        wsCache.Cells(sheetRow, 2).Value = rowData(2)
        wsCache.Cells(sheetRow, 3).NumberFormat = "@"
        wsCache.Cells(sheetRow, 3).Value = CStr(rowData(3))
        wsCache.Cells(sheetRow, 4).Value = rowData(4)
        wsCache.Cells(sheetRow, 5).Value = rowData(5)
        wsCache.Cells(sheetRow, 6).Value = rowData(6)
        wsCache.Cells(sheetRow, 7).Value = rowData(7)
        wsCache.Cells(sheetRow, 8).Value = rowData(8)
        wsCache.Cells(sheetRow, 9).Value = rowData(9)
        wsCache.Cells(sheetRow, 10).Value = rowData(10)
        wsCache.Cells(sheetRow, 11).Value = rowData(11)
        wsCache.Cells(sheetRow, 12).Value = rowData(12)
        wsCache.Cells(sheetRow, 13).Value = rowData(13)
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
        If rowIndex > startRow + PE_EDITOR_ABSOLUTE_MAX_ROWS Then Exit Do
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

Private Function ReadCachedOperationRows() As Object
    Dim wsCache As Worksheet
    Dim rows As Object
    Dim rowIndex As Long
    Dim operSeq As String
    Dim opLine As Long
    Dim rowData As Variant
    Dim valueIndex As Long

    Set rows = CreateObject("Scripting.Dictionary")
    rows.CompareMode = vbTextCompare

    Set wsCache = GetCacheWorksheet()
    If wsCache Is Nothing Then
        Set ReadCachedOperationRows = rows
        Exit Function
    End If

    rowIndex = CACHE_OPS_START_ROW
    Do While Len(Trim$(CStr(Nz(wsCache.Cells(rowIndex, 1).Value2)))) > 0
        operSeq = Trim$(CStr(Nz(wsCache.Cells(rowIndex, 1).Value2)))
        opLine = ReadOpLineValue(wsCache.Cells(rowIndex, 2).Value2)

        ReDim rowData(0 To CACHE_OPS_VALUE_COL_COUNT - 1)
        rowData(1) = operSeq
        rowData(2) = opLine
        For valueIndex = 3 To CACHE_OPS_VALUE_COL_COUNT - 1
            rowData(valueIndex) = wsCache.Cells(rowIndex, valueIndex).Value2
        Next valueIndex

        rows(BuildOperationCacheKey(operSeq, opLine)) = rowData
        rowIndex = rowIndex + 1
        If rowIndex > CACHE_OPS_START_ROW + PE_EDITOR_ABSOLUTE_MAX_ROWS Then Exit Do
    Loop

    Set ReadCachedOperationRows = rows
End Function

Private Function ReadSheetOperationRows(ByVal ws As Worksheet) As Object
    Dim rows As Object
    Dim rowIndex As Long
    Dim operSeq As String
    Dim opLine As Long
    Dim rowData As Variant

    Set rows = CreateObject("Scripting.Dictionary")
    rows.CompareMode = vbTextCompare

    For rowIndex = PE_OPS_DATA_START_ROW To PartEditorOpsLastRow()
        operSeq = Trim$(CStr(ws.Cells(rowIndex, PE_COL_OPER_SEQ).Value2))
        If Len(operSeq) = 0 Then GoTo ContinueOp

        opLine = ReadOpLineValue(ws.Cells(rowIndex, PE_COL_OP_LINE).Value2)

        ReDim rowData(0 To CACHE_OPS_VALUE_COL_COUNT - 1)
        rowData(1) = operSeq
        rowData(2) = opLine
        If Len(Trim$(ws.Cells(rowIndex, PE_COL_OPER_CODE).Text)) > 0 Then
            rowData(3) = Trim$(ws.Cells(rowIndex, PE_COL_OPER_CODE).Text)
        Else
            rowData(3) = Trim$(CStr(ws.Cells(rowIndex, PE_COL_OPER_CODE).Value2))
        End If
        rowData(4) = NormalizeCode(CStr(ws.Cells(rowIndex, PE_COL_MADE_IN_FFA).Value2))
        rowData(5) = NormalizeCode(CStr(ws.Cells(rowIndex, PE_COL_EQUIPMENT).Value2))
        rowData(6) = NormalizeCode(CStr(ws.Cells(rowIndex, PE_COL_PROCESS_TYPE).Value2))
        rowData(7) = ReadOptionalNumericCell(ws.Cells(rowIndex, PE_COL_PROCESS_HOURS))
        rowData(8) = ReadOptionalNumericCell(ws.Cells(rowIndex, PE_COL_MANUAL_AVG_EX))
        rowData(9) = ReadOptionalNumericCell(ws.Cells(rowIndex, PE_COL_BATCH_SIZE))
        rowData(10) = IsActiveFlag(ws.Cells(rowIndex, PE_COL_USE_AVG_HOURS).Value2)
        rowData(11) = IsActiveFlag(ws.Cells(rowIndex, PE_COL_USE_AVG_EX).Value2)
        rowData(12) = IsActiveFlag(ws.Cells(rowIndex, PE_COL_OPER_ACTIVE).Value2)
        rowData(13) = Trim$(CStr(ws.Cells(rowIndex, PE_COL_OPER_NOTES).Value2))

        If IsEmpty(ws.Cells(rowIndex, PE_COL_USE_AVG_HOURS).Value2) _
            Or Len(Trim$(CStr(Nz(ws.Cells(rowIndex, PE_COL_USE_AVG_HOURS).Value2)))) = 0 Then
            rowData(10) = True
        End If
        If IsEmpty(ws.Cells(rowIndex, PE_COL_USE_AVG_EX).Value2) _
            Or Len(Trim$(CStr(Nz(ws.Cells(rowIndex, PE_COL_USE_AVG_EX).Value2)))) = 0 Then
            rowData(11) = True
        End If

        rows(BuildOperationCacheKey(operSeq, opLine)) = rowData

ContinueOp:
    Next rowIndex

    Set ReadSheetOperationRows = rows
End Function

Private Function BuildOperationCacheKey(ByVal operSeq As String, ByVal opLine As Long) As String
    BuildOperationCacheKey = NormalizeOperSeqKey(operSeq) & vbTab & CStr(opLine)
End Function

Private Sub ParseOperationCacheKey(ByVal cacheKey As String, ByRef operSeq As String, ByRef opLine As Long)
    Dim parts As Variant

    parts = Split(cacheKey, vbTab)
    If UBound(parts) >= 0 Then
        operSeq = CStr(parts(0))
    Else
        operSeq = vbNullString
    End If
    If UBound(parts) >= 1 Then
        opLine = ReadOpLineValue(parts(1))
    Else
        opLine = 1
    End If
End Function

Private Function ReadOptionalNumericCell(ByVal cell As Range) As Variant
    Dim rawValue As Variant

    rawValue = cell.Value2
    If IsError(rawValue) Then
        ReadOptionalNumericCell = Empty
    ElseIf IsEmpty(rawValue) Or IsNull(rawValue) Then
        ReadOptionalNumericCell = Empty
    ElseIf Len(Trim$(CStr(rawValue))) = 0 Then
        ReadOptionalNumericCell = Empty
    ElseIf IsNumeric(rawValue) Then
        ReadOptionalNumericCell = CDbl(rawValue)
    Else
        ReadOptionalNumericCell = Empty
    End If
End Function

Private Sub ClearEditorDataRanges(ByVal ws As Worksheet)
    SafeClearCellOrMerge ws.Cells(PE_BASE_PART_ROW, PE_VALUE_COL)
    SafeClearCellOrMerge ws.Cells(PE_ROW_NAME, PE_VALUE_COL)
    SafeClearCellOrMerge ws.Cells(PE_STATUS_ROW, PE_VALUE_COL)
    SafeClearCellOrMerge ws.Cells(PE_ROW_FACTORY, PE_VALUE_COL)
    SafeClearCellOrMerge ws.Cells(PE_ROW_ACTIVE, PE_VALUE_COL)
    SafeClearCellOrMerge ws.Cells(PE_ROW_PRODUCT_LINE, PE_VALUE_COL)
    ClearEditorNotes ws

    SafeClearRange ws.Range( _
        ws.Cells(PE_DASH_DATA_START_ROW, PE_COL_DASH), _
        ws.Cells(PE_DASH_DATA_START_ROW + PE_DASH_MAX_ROWS - 1, PE_COL_DASH_NOTES))
    SafeNumberFormat ws.Range( _
        ws.Cells(PE_DASH_DATA_START_ROW, PE_COL_DASH), _
        ws.Cells(PE_DASH_DATA_START_ROW + PE_DASH_MAX_ROWS - 1, PE_COL_DASH)), "@"

    ClearRouteCardRange ws

    SafeClearRange ws.Range( _
        ws.Cells(PE_OPS_DATA_START_ROW, PE_COL_OPER_SEQ), _
        ws.Cells(PartEditorOpsLastRow(), PE_OPS_LAST_COL))
    SafeNumberFormat ws.Range( _
        ws.Cells(PE_OPS_DATA_START_ROW, PE_COL_OPER_CODE), _
        ws.Cells(PartEditorOpsLastRow(), PE_COL_OPER_CODE)), "@"

    ' Master Active stays True; Use Avg flags are filled only on rows that have operation data.
    ws.Cells(PE_ROW_ACTIVE, PE_VALUE_COL).Value = True
    SyncOperationRowDefaults ws
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
    Dim notesRange As Range

    Set notesRange = EditorNotesRange(ws)
    SafeClearRange notesRange
End Sub

Private Sub SafeClearCellOrMerge(ByVal cell As Range)
    On Error Resume Next
    If cell.MergeCells Then
        cell.MergeArea.ClearContents
    Else
        cell.ClearContents
    End If
    If Err.Number <> 0 Then
        Err.Clear
        cell.Value = vbNullString
    End If
    On Error GoTo 0
End Sub

Private Sub SafeClearRange(ByVal targetRange As Range)
    Dim cell As Range

    On Error Resume Next
    targetRange.ClearContents
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    Err.Clear

    ' Fall back cell-by-cell / merge-area clears when a partial merge blocks the range clear.
    For Each cell In targetRange.Cells
        If cell.MergeCells Then
            If cell.Address = cell.MergeArea.Cells(1, 1).Address Then
                cell.MergeArea.ClearContents
            End If
        Else
            cell.ClearContents
        End If
        Err.Clear
    Next cell
    On Error GoTo 0
End Sub

Private Sub SafeNumberFormat(ByVal targetRange As Range, ByVal formatText As String)
    On Error Resume Next
    targetRange.NumberFormat = formatText
    On Error GoTo 0
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
    If Len(factoryCodes) > 255 Then factoryCodes = Left$(factoryCodes, 255)

    On Error Resume Next
    validationRange.Validation.Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, _
        Formula1:=factoryCodes
    On Error GoTo 0
End Sub

Private Function BuildFactoryValidationList() As String
    BuildFactoryValidationList = JoinStringArray(ListActiveKeyCodes(FindTable(FACTORIES_TABLE_NAME), COL_FACTORY_CODE), ",")
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
    Set GetPartEditorWorksheet = FindWorksheetByName(PART_EDITOR_SHEET_NAME)
End Function

Private Function GetCacheWorksheet() As Worksheet
    Set GetCacheWorksheet = FindWorksheetByName(PART_EDITOR_CACHE_SHEET_NAME)
End Function

'==============================================================================
' Operations / Route Card block capacity (grows past the default formatted size)
'==============================================================================

Public Function PartEditorOpsCapacity() As Long
    PartEditorOpsCapacity = GetNamedCapacity(PE_OPS_CAPACITY_NAME, PE_OPS_MAX_ROWS)
End Function

Public Function PartEditorOpsLastRow() As Long
    PartEditorOpsLastRow = PE_OPS_DATA_START_ROW + PartEditorOpsCapacity() - 1
End Function

Public Function PartEditorRouteCapacity() As Long
    PartEditorRouteCapacity = GetNamedCapacity(PE_ROUTE_CAPACITY_NAME, PE_ROUTE_MAX_ROWS)
End Function

Public Function PartEditorRouteLastRow() As Long
    PartEditorRouteLastRow = PE_ROUTE_DATA_START_ROW + PartEditorRouteCapacity() - 1
End Function

Public Sub EnsurePartEditorOpsCapacity(ByVal ws As Worksheet, ByVal neededRows As Long, Optional ByVal allowShrink As Boolean = False)
    EnsurePartEditorBlockCapacity ws, True, neededRows, allowShrink
End Sub

Public Sub EnsurePartEditorRouteCapacity(ByVal ws As Worksheet, ByVal neededRows As Long, Optional ByVal allowShrink As Boolean = False)
    EnsurePartEditorBlockCapacity ws, False, neededRows, allowShrink
End Sub

Private Sub EnsurePartEditorBlockCapacity( _
    ByVal ws As Worksheet, _
    ByVal isOps As Boolean, _
    ByVal neededRows As Long, _
    ByVal allowShrink As Boolean)

    Dim minRows As Long
    Dim startRow As Long
    Dim currentCapacity As Long
    Dim targetCapacity As Long
    Dim oldLast As Long
    Dim newLast As Long
    Dim nameText As String

    If ws Is Nothing Then Exit Sub

    If isOps Then
        minRows = PE_OPS_MAX_ROWS
        startRow = PE_OPS_DATA_START_ROW
        currentCapacity = PartEditorOpsCapacity()
        nameText = PE_OPS_CAPACITY_NAME
    Else
        minRows = PE_ROUTE_MAX_ROWS
        startRow = PE_ROUTE_DATA_START_ROW
        currentCapacity = PartEditorRouteCapacity()
        nameText = PE_ROUTE_CAPACITY_NAME
    End If

    If neededRows < minRows Then neededRows = minRows
    If neededRows > PE_EDITOR_ABSOLUTE_MAX_ROWS Then neededRows = PE_EDITOR_ABSOLUTE_MAX_ROWS

    targetCapacity = currentCapacity
    If neededRows > currentCapacity Then
        targetCapacity = neededRows
    ElseIf allowShrink And neededRows < currentCapacity Then
        targetCapacity = neededRows
    End If

    If targetCapacity = currentCapacity Then
        SetNamedCapacity nameText, currentCapacity
        Exit Sub
    End If

    oldLast = startRow + currentCapacity - 1
    newLast = startRow + targetCapacity - 1
    SetNamedCapacity nameText, targetCapacity

    If targetCapacity > currentCapacity Then
        If isOps Then
            FormatPartEditorOpsRowRange ws, oldLast + 1, newLast
            ApplyPartEditorOpsCheckboxesForRows ws, oldLast + 1, newLast
            ApplyOperationDropdownsForRows ws, oldLast + 1, newLast
        Else
            FormatPartEditorRouteRowRange ws, oldLast + 1, newLast
        End If
    Else
        If isOps Then
            ClearPartEditorExtraOpsRows ws, newLast + 1, oldLast
        Else
            ClearPartEditorExtraRouteRows ws, newLast + 1, oldLast
        End If
    End If
End Sub

Private Function GetNamedCapacity(ByVal nameText As String, ByVal defaultValue As Long) As Long
    Dim nm As Name
    Dim rawValue As String
    Dim parsedValue As Long

    GetNamedCapacity = defaultValue
    On Error Resume Next
    Set nm = ThisWorkbook.Names(nameText)
    On Error GoTo 0
    If nm Is Nothing Then Exit Function

    rawValue = Trim$(Replace(Replace(nm.RefersTo, "=", ""), ",", ""))
    If Len(rawValue) = 0 Then Exit Function
    If Not IsNumeric(rawValue) Then Exit Function

    parsedValue = CLng(Val(rawValue))
    If parsedValue < defaultValue Then parsedValue = defaultValue
    If parsedValue > PE_EDITOR_ABSOLUTE_MAX_ROWS Then parsedValue = PE_EDITOR_ABSOLUTE_MAX_ROWS
    GetNamedCapacity = parsedValue
End Function

Private Sub SetNamedCapacity(ByVal nameText As String, ByVal capacityValue As Long)
    On Error Resume Next
    ThisWorkbook.Names(nameText).Delete
    If capacityValue < 1 Then capacityValue = 1
    ThisWorkbook.Names.Add Name:=nameText, RefersTo:="=" & CStr(capacityValue)
    On Error GoTo 0
End Sub

Private Sub GrowPartEditorOpsIfNeeded(ByVal ws As Worksheet, ByVal changedRow As Long)
    Dim lastRow As Long
    Dim triggerFrom As Long

    lastRow = PartEditorOpsLastRow()
    triggerFrom = lastRow - PE_BLOCK_GROW_TRIGGER_ROWS + 1
    If triggerFrom < PE_OPS_DATA_START_ROW Then triggerFrom = PE_OPS_DATA_START_ROW
    If changedRow < triggerFrom Or changedRow > lastRow Then Exit Sub
    If Not OperationEditorRowHasData(ws, changedRow) Then Exit Sub

    EnsurePartEditorOpsCapacity ws, PartEditorOpsCapacity() + PE_BLOCK_GROW_BY, False
End Sub

Private Sub GrowPartEditorRouteIfNeeded(ByVal ws As Worksheet, ByVal changedRow As Long)
    Dim lastRow As Long
    Dim triggerFrom As Long

    lastRow = PartEditorRouteLastRow()
    triggerFrom = lastRow - PE_BLOCK_GROW_TRIGGER_ROWS + 1
    If triggerFrom < PE_ROUTE_DATA_START_ROW Then triggerFrom = PE_ROUTE_DATA_START_ROW
    If changedRow < triggerFrom Or changedRow > lastRow Then Exit Sub
    If Not RouteEditorRowHasData(ws, changedRow) Then Exit Sub

    EnsurePartEditorRouteCapacity ws, PartEditorRouteCapacity() + PE_BLOCK_GROW_BY, False
End Sub

Private Function RouteEditorRowHasData(ByVal ws As Worksheet, ByVal sheetRow As Long) As Boolean
    RouteEditorRowHasData = _
        Not IsBlankCellValue(ws.Cells(sheetRow, PE_COL_ROUTE_DASH).Value2) _
        Or Not IsBlankCellValue(ws.Cells(sheetRow, PE_COL_ROUTE_OPER_SEQ).Value2) _
        Or Not IsBlankCellValue(ws.Cells(sheetRow, PE_COL_ROUTE_OPER_CODE).Value2)
End Function
