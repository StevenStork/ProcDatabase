Attribute VB_Name = "modDataStore"
Option Explicit

' Canonical store on the very-hidden Data sheet:
'   tblParts       membership + dirty/list stamps
'   tblOperations  every operation row across part sheets
'   B2/B3          short hashes after the last Home / export rebuild
'   B4             RefsDirty

Public Sub EnsureDataSheet()
    Dim ws As Worksheet
    Dim created As Boolean
    Dim errNum As Long
    Dim errDesc As String
    Static inEnsure As Boolean

    If inEnsure Then Exit Sub
    inEnsure = True
    On Error GoTo CleanUp

    Set ws = GetWorksheet(DATA_SHEET_NAME)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        On Error Resume Next
        ws.Name = DATA_SHEET_NAME
        Err.Clear
        On Error GoTo CleanUp
        created = True
    End If

    ws.Range(CATEGORY_CELL).Value = DATA_LABEL_VALUE
    ws.Range("A2").Value = "HomeListHash"
    ws.Range("A3").Value = "ExportOpsHash"
    ws.Range("A4").Value = "RefsDirty"
    ws.Range("A5").Value = "UiSchema"
    If Len(Trim$(CStr(Nz(ws.Range(DATA_UI_SCHEMA_CELL).Value)))) = 0 Then
        ws.Range(DATA_UI_SCHEMA_CELL).Value = UI_SCHEMA_VERSION
    End If

    RelocateCanonicalTables ws
    EnsureNamedTable ws, TBL_PARTS_NAME, DATA_PARTS_HEADER_ROW, DATA_PARTS_FIRST_COL, PartsHeaderNames()
    EnsureNamedTable ws, TBL_OPS_NAME, DATA_OPS_HEADER_ROW, DATA_OPS_FIRST_COL, OpsStoreHeaderNames()
    ConvertLegacyHomeFfaColumn ListObjectByName(ws, TBL_PARTS_NAME)

    If StrComp(Trim$(CStr(Nz(ws.Range(DATA_UI_SCHEMA_CELL).Value))), UI_SCHEMA_VERSION, vbTextCompare) <> 0 Then
        ws.Range(DATA_UI_SCHEMA_CELL).Value = UI_SCHEMA_VERSION
        MarkAllListSigsStale
        ws.Range(DATA_HOME_HASH_CELL).Value = vbNullString
        ws.Range(DATA_EXPORT_HASH_CELL).Value = vbNullString
    End If

    On Error Resume Next
    ws.Visible = xlSheetVeryHidden
    Err.Clear
    On Error GoTo CleanUp

CleanUp:
    errNum = Err.Number
    errDesc = Err.Description
    inEnsure = False
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise errNum, "EnsureDataSheet", errDesc
End Sub

' Creates the References sheet if missing, stamps A1/headers, and keeps it
' very hidden. Lives here (not in modReferences) so export/Home/DataStore
' compile even when frmReferences was not imported.
Public Sub EnsureReferencesSheet()
    Dim ws As Worksheet
    Dim created As Boolean

    Set ws = GetWorksheet(REFERENCES_SHEET_NAME)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        On Error Resume Next
        ws.Name = REFERENCES_SHEET_NAME
        On Error GoTo 0
        created = True
    End If

    ws.Range(CATEGORY_CELL).Value = REFS_LABEL_VALUE
    ws.Cells(1, REFS_FFA_COLUMN).Value = HDR_FFA
    ws.Cells(1, REFS_FACTORY_COLUMN).Value = "Factory"
    ws.Cells(1, REFS_PRODUCT_LINE_COLUMN).Value = "Product Line"
    ws.Cells(1, REFS_EQUIPMENT_COLUMN).Value = "Equipment"
    ws.Cells(1, REFS_EQUIPMENT_OWNERS_COLUMN).Value = "Owning FFAs"

    If created Then
        ws.Range("A1:F1").Font.Bold = True
    End If

    HideReferencesSheet ws
End Sub

Public Sub HideReferencesSheet(Optional ByVal ws As Worksheet)
    If ws Is Nothing Then
        Set ws = GetWorksheet(REFERENCES_SHEET_NAME)
        If ws Is Nothing Then Exit Sub
    End If

    On Error Resume Next
    ws.Visible = xlSheetVeryHidden
    On Error GoTo 0
End Sub

Public Sub RefreshAllProcDatabase()
    On Error GoTo CleanUp
    OptimizeExcel True

    EnsureDataSheet
    EnsureReferencesSheet
    SyncAllPartsToStore True
    RebuildHomeFromStore
    BuildExportSheetsCore EXPORT_SCOPE_ALL, vbNullString, True

CleanUp:
    OptimizeExcel False
End Sub

Public Sub HandleWorkbookBeforeSave()
    On Error GoTo CleanUp
    OptimizeExcel True
    EnsureDataSheet
    SyncDirtyPartsToStore
    RebuildHomeFromStore
    BuildExportSheetsCore EXPORT_SCOPE_ALL, vbNullString, False

CleanUp:
    OptimizeExcel False
End Sub

Public Sub HandleWorkbookSheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim ws As Worksheet

    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    Set ws = Sh

    If IsPartSheet(ws) Then
        If PartChangeTouchesOpsOrLists(ws, Target) Then
            MarkPartOpsDirty ws
        End If
        Exit Sub
    End If

    If StrComp(ws.Name, HOME_SHEET_NAME, vbTextCompare) = 0 Then
        If Not Intersect(Target, ws.Range( _
            ws.Cells(HOME_PART_TABLE_FIRST_DATA_ROW, HOME_PART_TABLE_ACTIVE_COLUMN), _
            ws.Cells(ws.Rows.Count, HOME_PART_TABLE_BASE_PART_COLUMN))) Is Nothing Then
            MarkHomeListStale
        End If
    End If
End Sub

Public Sub MarkPartOpsDirty(ByVal ws As Worksheet)
    Dim tbl As ListObject
    Dim rowIndex As Long
    Dim basePart As String

    If Not IsPartSheet(ws) Then Exit Sub
    EnsureDataSheet

    basePart = Trim$(CStr(Nz(ws.Range(PART_NUMBER_CELL).Value)))
    If Len(basePart) = 0 Then basePart = ws.Name

    Set tbl = PartsTable()
    rowIndex = FindPartsRow(tbl, basePart)
    If rowIndex = 0 Then
        rowIndex = AppendPartsRow(tbl, basePart, ws.Name)
    End If

    tbl.ListColumns(COL_PARTS_OPS_DIRTY).DataBodyRange.Cells(rowIndex, 1).Value = True
    tbl.ListColumns(COL_PARTS_SHEET_NAME).DataBodyRange.Cells(rowIndex, 1).Value = ws.Name
End Sub

Public Sub MarkReferencesStale()
    EnsureDataSheet
    DataSheet().Range(DATA_REFS_DIRTY_CELL).Value = True
    MarkAllListSigsStale
End Sub

Public Sub SyncDirtyPartsToStore()
    SyncAllPartsToStore False
End Sub

Public Sub SyncAllPartsToStore(ByVal forceAll As Boolean)
    Dim ws As Worksheet
    Dim refsDirty As Boolean

    EnsureDataSheet
    refsDirty = IsActiveFlag(DataSheet().Range(DATA_REFS_DIRTY_CELL).Value)

    For Each ws In ThisWorkbook.Worksheets
        If IsPartSheet(ws) Then
            If forceAll Or refsDirty Or PartIsDirty(ws) Then
                SyncPartToStore ws
            End If
        End If
    Next ws

    If refsDirty Then DataSheet().Range(DATA_REFS_DIRTY_CELL).Value = False
    RemoveMissingPartRows
End Sub

Public Sub SyncPartToStore(ByVal ws As Worksheet)
    Dim tblParts As ListObject
    Dim basePart As String
    Dim rowIndex As Long
    Dim homeFfa As String
    Dim productLines As String
    Dim dashes As String
    Dim factories As String
    Dim listSig As String
    Dim opsCount As Long

    If Not IsPartSheet(ws) Then Exit Sub
    EnsureDataSheet
    EnsureHomeFfaField ws
    EnsurePartOpsTable ws

    basePart = Trim$(CStr(Nz(ws.Range(PART_NUMBER_CELL).Value)))
    If Len(basePart) = 0 Then basePart = ws.Name

    homeFfa = PartHomeFfaValue(ws)
    productLines = JoinCheckedList(CheckedValuesOnPart(ws, PRODUCT_LINE_VALUE_COLUMN, PRODUCT_LINE_CHECKBOX_COLUMN))
    dashes = JoinCheckedList(CheckedValuesOnPart(ws, DASH_VALUE_COLUMN, DASH_CHECKBOX_COLUMN))
    factories = FactoriesForFfaList(homeFfa)
    listSig = BuildListSignature(basePart)

    Set tblParts = PartsTable()
    rowIndex = FindPartsRow(tblParts, basePart)
    If rowIndex = 0 Then rowIndex = AppendPartsRow(tblParts, basePart, ws.Name)

    With tblParts
        .ListColumns(COL_PARTS_BASE).DataBodyRange.Cells(rowIndex, 1).Value = basePart
        .ListColumns(COL_PARTS_ACTIVE).DataBodyRange.Cells(rowIndex, 1).Value = HomePartIsActive(basePart)
        .ListColumns(COL_PARTS_HOME_FFA).DataBodyRange.Cells(rowIndex, 1).Value = homeFfa
        .ListColumns(COL_PARTS_FACTORIES).DataBodyRange.Cells(rowIndex, 1).Value = factories
        .ListColumns(COL_PARTS_PRODUCT_LINES).DataBodyRange.Cells(rowIndex, 1).Value = productLines
        .ListColumns(COL_PARTS_DASHES).DataBodyRange.Cells(rowIndex, 1).Value = dashes
        .ListColumns(COL_PARTS_UI_SCHEMA).DataBodyRange.Cells(rowIndex, 1).Value = UI_SCHEMA_VERSION
        .ListColumns(COL_PARTS_LIST_SIG).DataBodyRange.Cells(rowIndex, 1).Value = listSig
        .ListColumns(COL_PARTS_OPS_DIRTY).DataBodyRange.Cells(rowIndex, 1).Value = False
        .ListColumns(COL_PARTS_SHEET_NAME).DataBodyRange.Cells(rowIndex, 1).Value = ws.Name
    End With

    opsCount = ReplaceOperationsForPart(basePart, ws)
    tblParts.ListColumns(COL_PARTS_OPS_ROW_COUNT).DataBodyRange.Cells(rowIndex, 1).Value = opsCount
End Sub

Public Function PartListSig(ByVal ws As Worksheet) As String
    Dim tbl As ListObject
    Dim rowIndex As Long
    Dim basePart As String

    EnsureDataSheet
    basePart = Trim$(CStr(Nz(ws.Range(PART_NUMBER_CELL).Value)))
    If Len(basePart) = 0 Then basePart = ws.Name

    Set tbl = PartsTable()
    rowIndex = FindPartsRow(tbl, basePart)
    If rowIndex = 0 Then Exit Function
    PartListSig = CStr(Nz(tbl.ListColumns(COL_PARTS_LIST_SIG).DataBodyRange.Cells(rowIndex, 1).Value))
End Function

Public Function BuildListSignature(ByVal basePart As String) As String
    BuildListSignature = UI_SCHEMA_VERSION & Chr$(31) & _
        JoinStringArray(ReferenceColumnValues(REFS_FFA_COLUMN)) & Chr$(31) & _
        JoinStringArray(ReferenceColumnValues(REFS_PRODUCT_LINE_COLUMN)) & Chr$(31) & _
        JoinStringArray(DashConditionsForBasePart(basePart))
    BuildListSignature = HashString(BuildListSignature)
End Function

Public Function PartsTable() As ListObject
    EnsureDataSheet
    Set PartsTable = DataSheet().ListObjects(TBL_PARTS_NAME)
End Function

Public Function OpsTable() As ListObject
    EnsureDataSheet
    Set OpsTable = DataSheet().ListObjects(TBL_OPS_NAME)
End Function

Public Function DataSheet() As Worksheet
    EnsureDataSheet
    Set DataSheet = ThisWorkbook.Worksheets(DATA_SHEET_NAME)
End Function

Public Function HomeListHashIsCurrent() As Boolean
    Dim currentHash As String

    EnsureDataSheet
    currentHash = HashString(PartsTableFingerprint())
    HomeListHashIsCurrent = (StrComp(CStr(Nz(DataSheet().Range(DATA_HOME_HASH_CELL).Value)), currentHash, vbBinaryCompare) = 0)
End Function

Public Sub WriteHomeListHash()
    EnsureDataSheet
    DataSheet().Range(DATA_HOME_HASH_CELL).Value = HashString(PartsTableFingerprint())
End Sub

Public Function ExportOpsHashIsCurrent() As Boolean
    Dim currentHash As String

    EnsureDataSheet
    currentHash = HashString(OpsTableFingerprint())
    ExportOpsHashIsCurrent = (StrComp(CStr(Nz(DataSheet().Range(DATA_EXPORT_HASH_CELL).Value)), currentHash, vbBinaryCompare) = 0)
End Function

Public Sub WriteExportOpsHash()
    EnsureDataSheet
    DataSheet().Range(DATA_EXPORT_HASH_CELL).Value = HashString(OpsTableFingerprint())
End Sub

Public Sub RebuildHomeFromStore()
    Dim wsHome As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim basePart As String
    Dim tbl As ListObject
    Dim storeRow As Long

    Set wsHome = ThisWorkbook.Worksheets(HOME_SHEET_NAME)
    lastRow = HomePartTableLastRow(wsHome)
    EnsureDataSheet
    Set tbl = PartsTable()

    If lastRow < HOME_PART_TABLE_FIRST_DATA_ROW Then
        WriteHomeListHash
        Exit Sub
    End If

    For rowIndex = HOME_PART_TABLE_FIRST_DATA_ROW To lastRow
        basePart = Trim$(CStr(Nz(wsHome.Cells(rowIndex, HOME_PART_TABLE_BASE_PART_COLUMN).Value)))
        If Len(basePart) = 0 Or Not IsActiveFlag(wsHome.Cells(rowIndex, HOME_PART_TABLE_ACTIVE_COLUMN).Value) Then
            wsHome.Cells(rowIndex, HOME_PART_TABLE_FFA_COLUMN).Value = vbNullString
            wsHome.Cells(rowIndex, HOME_PART_TABLE_FACTORY_COLUMN).Value = vbNullString
        Else
            storeRow = FindPartsRow(tbl, basePart)
            If storeRow = 0 Then
                wsHome.Cells(rowIndex, HOME_PART_TABLE_FFA_COLUMN).Value = vbNullString
                wsHome.Cells(rowIndex, HOME_PART_TABLE_FACTORY_COLUMN).Value = vbNullString
            Else
                wsHome.Cells(rowIndex, HOME_PART_TABLE_FFA_COLUMN).Value = _
                    tbl.ListColumns(COL_PARTS_HOME_FFA).DataBodyRange.Cells(storeRow, 1).Value
                wsHome.Cells(rowIndex, HOME_PART_TABLE_FACTORY_COLUMN).Value = _
                    tbl.ListColumns(COL_PARTS_FACTORIES).DataBodyRange.Cells(storeRow, 1).Value
            End If
        End If
    Next rowIndex

    WriteHomeListHash
End Sub

Public Function HomePartTableLastRow(ByVal ws As Worksheet) As Long
    Dim columnLetter As Variant
    Dim columnLastRow As Long
    Dim maxRow As Long

    maxRow = HOME_PART_TABLE_HEADER_ROW - 1
    For Each columnLetter In Array( _
        HOME_PART_TABLE_BASE_PART_COLUMN, _
        HOME_PART_TABLE_ACTIVE_COLUMN, _
        HOME_PART_TABLE_DATE_COLUMN, _
        HOME_PART_TABLE_HIGHLIGHT_COLUMN_G)

        columnLastRow = LastUsedRowInColumn(ws, CStr(columnLetter))
        If columnLastRow > maxRow Then maxRow = columnLastRow
    Next columnLetter

    If maxRow < HOME_PART_TABLE_HEADER_ROW Then
        HomePartTableLastRow = HOME_PART_TABLE_HEADER_ROW - 1
    Else
        HomePartTableLastRow = maxRow
    End If
End Function

Public Function HomePartIsActive(ByVal basePart As String) As Boolean
    Dim wsHome As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long

    Set wsHome = ThisWorkbook.Worksheets(HOME_SHEET_NAME)
    lastRow = HomePartTableLastRow(wsHome)
    If lastRow < HOME_PART_TABLE_FIRST_DATA_ROW Then Exit Function

    For rowIndex = HOME_PART_TABLE_FIRST_DATA_ROW To lastRow
        If StrComp(Trim$(CStr(Nz(wsHome.Cells(rowIndex, HOME_PART_TABLE_BASE_PART_COLUMN).Value))), basePart, vbTextCompare) = 0 Then
            HomePartIsActive = IsActiveFlag(wsHome.Cells(rowIndex, HOME_PART_TABLE_ACTIVE_COLUMN).Value)
            Exit Function
        End If
    Next rowIndex
End Function

Public Function ReferenceColumnValues(ByVal columnLetter As String) As String()
    Dim wsReferences As Worksheet
    Dim lastRow As Long
    Dim rawValues As Variant

    Set wsReferences = ThisWorkbook.Worksheets(REFERENCES_SHEET_NAME)
    lastRow = FastLastUsedRowInColumn(wsReferences, columnLetter)
    If lastRow < 2 Then
        ReferenceColumnValues = EmptyStringArray()
        Exit Function
    End If

    rawValues = wsReferences.Range(wsReferences.Cells(2, columnLetter), wsReferences.Cells(lastRow, columnLetter)).Value2
    ReferenceColumnValues = UniqueSortedValuesFromColumn(rawValues)
End Function

Public Function DashConditionsForBasePart(ByVal basePart As String) As String()
    Dim wsStandards As Worksheet
    Dim tbl As ListObject
    Dim assemblyValues As Variant
    Dim rowIndex As Long
    Dim assemblyVal As String
    Dim rowBasePart As String
    Dim dashCondition As String
    Dim uniqueValues As Object

    Set uniqueValues = CreateObject("Scripting.Dictionary")
    uniqueValues.CompareMode = vbTextCompare

    Set wsStandards = GetWorksheet(ASSEMBLY_STANDARDS_SHEET_NAME)
    If wsStandards Is Nothing Then
        DashConditionsForBasePart = EmptyStringArray()
        Exit Function
    End If

    Set tbl = ListObjectByName(wsStandards, ASSY_STANDARDS_TABLE_NAME)
    If tbl Is Nothing Then
        DashConditionsForBasePart = EmptyStringArray()
        Exit Function
    End If
    If tbl.DataBodyRange Is Nothing Then
        DashConditionsForBasePart = EmptyStringArray()
        Exit Function
    End If

    assemblyValues = ListColumnValues(tbl.ListColumns(COL_ASSEMBLY_NO))
    If IsEmpty(assemblyValues) Then
        DashConditionsForBasePart = EmptyStringArray()
        Exit Function
    End If

    For rowIndex = 1 To UBound(assemblyValues, 1)
        assemblyVal = Trim$(CStr(Nz(assemblyValues(rowIndex, 1))))
        If Len(assemblyVal) > 0 Then
            SplitAssemblyNo assemblyVal, rowBasePart, dashCondition
            If StrComp(rowBasePart, basePart, vbTextCompare) = 0 And Len(dashCondition) > 0 Then
                If Not uniqueValues.Exists(dashCondition) Then uniqueValues.Add dashCondition, dashCondition
            End If
        End If
    Next rowIndex

    DashConditionsForBasePart = DictionaryKeysToSortedArray(uniqueValues)
End Function

Public Function FactoriesForFfaList(ByVal ffaList As String) As String
    Dim wsReferences As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim ffaValue As String
    Dim factoryName As String
    Dim factoryMap As Object
    Dim uniqueFactories As Object
    Dim parts As Variant
    Dim i As Long

    Set uniqueFactories = CreateObject("Scripting.Dictionary")
    uniqueFactories.CompareMode = vbTextCompare
    Set factoryMap = CreateObject("Scripting.Dictionary")
    factoryMap.CompareMode = vbTextCompare

    Set wsReferences = GetWorksheet(REFERENCES_SHEET_NAME)
    If wsReferences Is Nothing Then
        FactoriesForFfaList = vbNullString
        Exit Function
    End If

    lastRow = FastLastUsedRowInColumn(wsReferences, REFS_FFA_COLUMN)
    For rowIndex = 2 To lastRow
        ffaValue = Trim$(CStr(Nz(wsReferences.Cells(rowIndex, REFS_FFA_COLUMN).Value)))
        factoryName = Trim$(CStr(Nz(wsReferences.Cells(rowIndex, REFS_FACTORY_COLUMN).Value)))
        If Len(ffaValue) > 0 And Not factoryMap.Exists(ffaValue) Then
            factoryMap.Add ffaValue, factoryName
        End If
    Next rowIndex

    parts = Split(ffaList, ",")
    For i = LBound(parts) To UBound(parts)
        ffaValue = Trim$(CStr(parts(i)))
        If Len(ffaValue) > 0 Then
            If factoryMap.Exists(ffaValue) Then
                factoryName = CStr(factoryMap(ffaValue))
                If Len(factoryName) > 0 Then
                    If Not uniqueFactories.Exists(factoryName) Then uniqueFactories.Add factoryName, factoryName
                End If
            End If
        End If
    Next i

    FactoriesForFfaList = JoinCheckedList(uniqueFactories)
End Function

Public Function CheckedValuesOnPart( _
    ByVal ws As Worksheet, _
    ByVal valueColumn As String, _
    ByVal checkboxColumn As String) As Object

    Dim lastRow As Long
    Dim rowIndex As Long
    Dim itemName As String
    Dim checked As Object

    Set checked = CreateObject("Scripting.Dictionary")
    checked.CompareMode = vbTextCompare

    lastRow = FastLastUsedRowInColumn(ws, valueColumn)
    If lastRow < LIST_START_ROW Then
        Set CheckedValuesOnPart = checked
        Exit Function
    End If

    For rowIndex = LIST_START_ROW To lastRow
        itemName = Trim$(CStr(Nz(ws.Cells(rowIndex, valueColumn).Value)))
        If Len(itemName) = 0 Then Exit For
        If IsActiveFlag(ws.Cells(rowIndex, checkboxColumn).Value) Then
            If Not checked.Exists(itemName) Then checked.Add itemName, itemName
        End If
    Next rowIndex

    Set CheckedValuesOnPart = checked
End Function

Public Function PartHomeFfaValue(ByVal ws As Worksheet) As String
    Dim homeFfa As String

    homeFfa = Trim$(CStr(Nz(ws.Range(HOME_FFA_VALUE_CELL).Value)))
    If StrComp(homeFfa, HOME_FFA_LABEL, vbTextCompare) = 0 Then Exit Function
    If StrComp(homeFfa, HDR_FFA, vbTextCompare) = 0 Then Exit Function
    PartHomeFfaValue = homeFfa
End Function

' Writes the Home FFA label and blue dropdown. Converts leftover FFA checkbox
' lists by taking the first checked FFA. Returns True when the layout changed.
Public Function EnsureHomeFfaField(ByVal ws As Worksheet) As Boolean
    Dim savedValue As String
    Dim alreadyHome As Boolean
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim itemName As String

    alreadyHome = (StrComp(Trim$(CStr(Nz(ws.Range(HOME_FFA_LABEL_CELL).Value))), HOME_FFA_LABEL, vbTextCompare) = 0)
    If alreadyHome Then
        savedValue = PartHomeFfaValue(ws)
    Else
        EnsureHomeFfaField = True
        lastRow = FastLastUsedRowInColumn(ws, "C")
        If lastRow >= LIST_START_ROW Then
            For rowIndex = LIST_START_ROW To lastRow
                itemName = Trim$(CStr(Nz(ws.Cells(rowIndex, "C").Value)))
                If Len(itemName) = 0 Then Exit For
                If IsActiveFlag(ws.Cells(rowIndex, "D").Value) Then
                    savedValue = itemName
                    Exit For
                End If
            Next rowIndex
        End If
        If Len(savedValue) = 0 Then savedValue = PartHomeFfaValue(ws)
        ClearLegacyFfaChecklist ws
    End If

    On Error Resume Next
    ws.Range("C8:D8").UnMerge
    On Error GoTo 0

    With ws.Range(HOME_FFA_LABEL_CELL)
        .Value = HOME_FFA_LABEL
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With ws.Range(HOME_FFA_VALUE_CELL)
        .Interior.Color = INPUT_FILL_RGB
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        If Len(savedValue) > 0 Then .Value = savedValue
    End With

    ApplyHomeFfaValidation ws
End Function

' Keeps M:Z as a formatted range. Converts leftover PartOpsTbl ListObjects
' back to a normal range so the part data entry block is not a table.
Public Sub EnsurePartOpsTable(ByVal ws As Worksheet)
    Dim lo As ListObject
    Dim doomed As Collection
    Dim headers As Variant
    Dim i As Long
    Dim firstCol As Long

    firstCol = ws.Columns(OPS_FIRST_COLUMN).Column
    Set doomed = New Collection
    For Each lo In ws.ListObjects
        If StrComp(lo.Name, PART_OPS_TABLE_NAME, vbTextCompare) = 0 Then
            doomed.Add lo
        ElseIf Not lo.HeaderRowRange Is Nothing Then
            If lo.HeaderRowRange.Row = OPS_HEADER_ROW Then
                If lo.HeaderRowRange.Column = firstCol Then
                    doomed.Add lo
                End If
            End If
        End If
    Next lo

    For i = 1 To doomed.Count
        On Error Resume Next
        doomed(i).Unlist
        On Error GoTo 0
    Next i

    headers = OpsHeaderNames()
    For i = LBound(headers) To UBound(headers)
        With ws.Cells(OPS_HEADER_ROW, firstCol + i - LBound(headers))
            .Value = headers(i)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next i
End Sub

Public Function ListColumnByHeader(ByVal tbl As ListObject, ByVal headerName As String) As ListColumn
    Dim col As ListColumn

    If tbl Is Nothing Then Exit Function
    For Each col In tbl.ListColumns
        If StrComp(Trim$(col.Name), headerName, vbTextCompare) = 0 Then
            Set ListColumnByHeader = col
            Exit Function
        End If
    Next col
End Function

Private Function ReplaceOperationsForPart(ByVal basePart As String, ByVal ws As Worksheet) As Long
    Dim dest As ListObject
    Dim keepers As Collection
    Dim rowIndex As Long
    Dim partCol As ListColumn
    Dim newRows() As Variant
    Dim srcRow As Long
    Dim colIndex As Long
    Dim destHeaders As Variant
    Dim keepCount As Long
    Dim i As Long
    Dim output() As Variant
    Dim totalRows As Long
    Dim lastRow As Long
    Dim headerCols As Object
    Dim seqCol As Long
    Dim codeCol As Long
    Dim destColCount As Long
    Dim headerName As String
    Dim srcCol As Long

    Set dest = OpsTable()
    destHeaders = OpsStoreHeaderNames()
    destColCount = UBound(destHeaders) - LBound(destHeaders) + 1
    Set headerCols = PartOpsHeaderMap(ws)

    Set keepers = New Collection
    If Not dest.DataBodyRange Is Nothing Then
        Set partCol = ListColumnByHeader(dest, COL_OPS_PART_NUMBER)
        For rowIndex = 1 To dest.DataBodyRange.Rows.Count
            If StrComp(Trim$(CStr(Nz(partCol.DataBodyRange.Cells(rowIndex, 1).Value))), basePart, vbTextCompare) <> 0 Then
                keepers.Add dest.DataBodyRange.Rows(rowIndex).Value
            End If
        Next rowIndex
    End If

    lastRow = FastLastUsedRowInColumns(ws, OPS_FIRST_COLUMN, OPS_LAST_COLUMN)
    If lastRow >= LIST_START_ROW Then
        If headerCols.Exists(HDR_OP_SEQUENCE) Then seqCol = CLng(headerCols(HDR_OP_SEQUENCE))
        If headerCols.Exists(HDR_OP_CODE) Then codeCol = CLng(headerCols(HDR_OP_CODE))
        If seqCol > 0 And codeCol > 0 Then
            ReDim newRows(1 To lastRow - LIST_START_ROW + 1, 1 To destColCount)
            For srcRow = LIST_START_ROW To lastRow
                If Len(Trim$(CStr(Nz(ws.Cells(srcRow, seqCol).Value)))) = 0 Then
                    If Len(Trim$(CStr(Nz(ws.Cells(srcRow, codeCol).Value)))) = 0 Then
                        GoTo NextSrc
                    End If
                End If
                ReplaceOperationsForPart = ReplaceOperationsForPart + 1
                For colIndex = LBound(destHeaders) To UBound(destHeaders)
                    headerName = CStr(destHeaders(colIndex))
                    If StrComp(headerName, COL_OPS_PART_NUMBER, vbTextCompare) = 0 Then
                        newRows(ReplaceOperationsForPart, colIndex - LBound(destHeaders) + 1) = basePart
                    ElseIf headerCols.Exists(headerName) Then
                        srcCol = CLng(headerCols(headerName))
                        newRows(ReplaceOperationsForPart, colIndex - LBound(destHeaders) + 1) = ws.Cells(srcRow, srcCol).Value
                    Else
                        newRows(ReplaceOperationsForPart, colIndex - LBound(destHeaders) + 1) = vbNullString
                    End If
                Next colIndex
NextSrc:
            Next srcRow
        End If
    End If

    keepCount = keepers.Count
    totalRows = keepCount + ReplaceOperationsForPart
    ClearTableBody dest

    If totalRows = 0 Then Exit Function

    ReDim output(1 To totalRows, 1 To destColCount)
    For i = 1 To keepCount
        CopyRowInto output, i, keepers(i)
    Next i
    For i = 1 To ReplaceOperationsForPart
        For colIndex = 1 To UBound(output, 2)
            output(keepCount + i, colIndex) = newRows(i, colIndex)
        Next colIndex
    Next i

    dest.Resize dest.HeaderRowRange.Resize(totalRows + 1, UBound(output, 2))
    dest.DataBodyRange.Value2 = output
End Function

Private Function PartOpsHeaderMap(ByVal ws As Worksheet) As Object
    Dim map As Object
    Dim colIndex As Long
    Dim headerName As String

    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = vbTextCompare
    For colIndex = ws.Columns(OPS_FIRST_COLUMN).Column To ws.Columns(OPS_LAST_COLUMN).Column
        headerName = Trim$(CStr(Nz(ws.Cells(OPS_HEADER_ROW, colIndex).Value)))
        If Len(headerName) > 0 Then
            If Not map.Exists(headerName) Then map.Add headerName, colIndex
        End If
    Next colIndex
    Set PartOpsHeaderMap = map
End Function

Private Sub CopyRowInto(ByRef output() As Variant, ByVal destRow As Long, ByVal sourceRow As Variant)
    Dim colIndex As Long

    If Not IsArray(sourceRow) Then
        output(destRow, 1) = sourceRow
        Exit Sub
    End If

    For colIndex = 1 To UBound(output, 2)
        output(destRow, colIndex) = sourceRow(1, colIndex)
    Next colIndex
End Sub

Private Sub ClearTableBody(ByVal tbl As ListObject)
    Dim colCount As Long

    colCount = tbl.ListColumns.Count
    If Not tbl.DataBodyRange Is Nothing Then
        If tbl.ListRows.Count > 1 Then
            tbl.DataBodyRange.Delete
        Else
            tbl.DataBodyRange.ClearContents
        End If
    End If

    If tbl.ListRows.Count = 0 Then
        tbl.ListRows.Add
        tbl.DataBodyRange.ClearContents
    End If
End Sub

Private Function FindPartsRow(ByVal tbl As ListObject, ByVal basePart As String) As Long
    Dim rowIndex As Long
    Dim col As ListColumn

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    Set col = tbl.ListColumns(COL_PARTS_BASE)
    For rowIndex = 1 To tbl.DataBodyRange.Rows.Count
        If StrComp(Trim$(CStr(Nz(col.DataBodyRange.Cells(rowIndex, 1).Value))), basePart, vbTextCompare) = 0 Then
            FindPartsRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function AppendPartsRow(ByVal tbl As ListObject, ByVal basePart As String, ByVal sheetName As String) As Long
    Dim lr As ListRow

    If tbl.DataBodyRange Is Nothing Then
        tbl.ListRows.Add
    ElseIf tbl.ListRows.Count = 1 Then
        If Len(Trim$(CStr(Nz(tbl.ListColumns(COL_PARTS_BASE).DataBodyRange.Cells(1, 1).Value)))) = 0 Then
            AppendPartsRow = 1
            tbl.ListColumns(COL_PARTS_BASE).DataBodyRange.Cells(1, 1).Value = basePart
            tbl.ListColumns(COL_PARTS_SHEET_NAME).DataBodyRange.Cells(1, 1).Value = sheetName
            Exit Function
        End If
        tbl.ListRows.Add
    Else
        tbl.ListRows.Add
    End If

    AppendPartsRow = tbl.ListRows.Count
    tbl.ListColumns(COL_PARTS_BASE).DataBodyRange.Cells(AppendPartsRow, 1).Value = basePart
    tbl.ListColumns(COL_PARTS_SHEET_NAME).DataBodyRange.Cells(AppendPartsRow, 1).Value = sheetName
End Function

Private Function PartIsDirty(ByVal ws As Worksheet) As Boolean
    Dim tbl As ListObject
    Dim rowIndex As Long
    Dim basePart As String

    basePart = Trim$(CStr(Nz(ws.Range(PART_NUMBER_CELL).Value)))
    If Len(basePart) = 0 Then basePart = ws.Name
    Set tbl = PartsTable()
    rowIndex = FindPartsRow(tbl, basePart)
    If rowIndex = 0 Then
        PartIsDirty = True
        Exit Function
    End If
    PartIsDirty = IsActiveFlag(tbl.ListColumns(COL_PARTS_OPS_DIRTY).DataBodyRange.Cells(rowIndex, 1).Value)
End Function

Private Sub MarkAllListSigsStale()
    Dim tbl As ListObject
    Dim rowIndex As Long

    Set tbl = PartsTable()
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    For rowIndex = 1 To tbl.DataBodyRange.Rows.Count
        tbl.ListColumns(COL_PARTS_LIST_SIG).DataBodyRange.Cells(rowIndex, 1).Value = vbNullString
        tbl.ListColumns(COL_PARTS_OPS_DIRTY).DataBodyRange.Cells(rowIndex, 1).Value = True
    Next rowIndex
End Sub

Private Sub MarkHomeListStale()
    EnsureDataSheet
    DataSheet().Range(DATA_HOME_HASH_CELL).Value = vbNullString
End Sub

Private Sub RemoveMissingPartRows()
    Dim tbl As ListObject
    Dim rowIndex As Long
    Dim sheetName As String
    Dim ws As Worksheet

    Set tbl = PartsTable()
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = tbl.ListRows.Count To 1 Step -1
        sheetName = Trim$(CStr(Nz(tbl.ListColumns(COL_PARTS_SHEET_NAME).DataBodyRange.Cells(rowIndex, 1).Value)))
        If Len(sheetName) > 0 Then
            Set ws = GetWorksheet(sheetName)
            If ws Is Nothing Then
                If tbl.ListRows.Count = 1 Then
                    tbl.DataBodyRange.ClearContents
                Else
                    tbl.ListRows(rowIndex).Delete
                End If
            End If
        End If
    Next rowIndex
End Sub

Private Function PartChangeTouchesOpsOrLists(ByVal ws As Worksheet, ByVal Target As Range) As Boolean
    Dim watchRange As Range

    If Not Intersect(Target, ws.Range(HOME_FFA_VALUE_CELL)) Is Nothing Then
        PartChangeTouchesOpsOrLists = True
        Exit Function
    End If

    Set watchRange = ws.Range( _
        ws.Cells(LIST_HEADER_ROW, DASH_VALUE_COLUMN), _
        ws.Cells(ws.Rows.Count, PRODUCT_LINE_CHECKBOX_COLUMN))
    If Not Intersect(Target, watchRange) Is Nothing Then
        PartChangeTouchesOpsOrLists = True
        Exit Function
    End If

    Set watchRange = ws.Range( _
        ws.Cells(OPS_HEADER_ROW, OPS_FIRST_COLUMN), _
        ws.Cells(ws.Rows.Count, OPS_LAST_COLUMN))
    PartChangeTouchesOpsOrLists = Not Intersect(Target, watchRange) Is Nothing
End Function

Private Function PartsTableFingerprint() As String
    Dim tbl As ListObject

    Set tbl = PartsTable()
    If tbl.DataBodyRange Is Nothing Then
        PartsTableFingerprint = "0"
    Else
        PartsTableFingerprint = HashVariantColumn(tbl.ListColumns(COL_PARTS_BASE).DataBodyRange.Value2) & Chr$(31) & _
            HashVariantColumn(tbl.ListColumns(COL_PARTS_HOME_FFA).DataBodyRange.Value2) & Chr$(31) & _
            HashVariantColumn(tbl.ListColumns(COL_PARTS_FACTORIES).DataBodyRange.Value2)
    End If
End Function

Private Function OpsTableFingerprint() As String
    Dim tbl As ListObject

    Set tbl = OpsTable()
    If tbl.DataBodyRange Is Nothing Then
        OpsTableFingerprint = "0"
    Else
        OpsTableFingerprint = HashVariantColumn(tbl.ListColumns(COL_OPS_PART_NUMBER).DataBodyRange.Value2) & Chr$(31) & _
            HashVariantColumn(ListColumnByHeader(tbl, HDR_OP_SEQUENCE).DataBodyRange.Value2) & Chr$(31) & _
            HashVariantColumn(ListColumnByHeader(tbl, HDR_PROCESS_HOURS).DataBodyRange.Value2) & Chr$(31) & _
            HashVariantColumn(ListColumnByHeader(tbl, HDR_FFA).DataBodyRange.Value2)
    End If
End Function

Private Sub RelocateCanonicalTables(ByVal ws As Worksheet)
    Dim loParts As ListObject
    Dim loOps As ListObject
    Dim partsBody As Variant
    Dim opsBody As Variant
    Dim partsHdrs As Variant
    Dim opsHdrs As Variant
    Dim moveTables As Boolean

    Set loParts = ListObjectByName(ws, TBL_PARTS_NAME)
    Set loOps = ListObjectByName(ws, TBL_OPS_NAME)

    If Not loParts Is Nothing Then
        If loParts.HeaderRowRange.Row <> DATA_PARTS_HEADER_ROW Or loParts.HeaderRowRange.Column <> DATA_PARTS_FIRST_COL Then
            moveTables = True
        End If
    End If
    If Not loOps Is Nothing Then
        If loOps.HeaderRowRange.Row <> DATA_OPS_HEADER_ROW Or loOps.HeaderRowRange.Column <> DATA_OPS_FIRST_COL Then
            moveTables = True
        End If
    End If
    If Not moveTables Then Exit Sub

    If Not loParts Is Nothing Then
        partsHdrs = TableHeaderNames(loParts)
        If Not loParts.DataBodyRange Is Nothing Then partsBody = loParts.DataBodyRange.Value
    End If
    If Not loOps Is Nothing Then
        opsHdrs = TableHeaderNames(loOps)
        If Not loOps.DataBodyRange Is Nothing Then opsBody = loOps.DataBodyRange.Value
    End If

    If Not loParts Is Nothing Then loParts.Delete
    If Not loOps Is Nothing Then loOps.Delete

    ws.Range(ws.Cells(DATA_PARTS_HEADER_ROW, 1), ws.Cells(ws.Rows.Count, DATA_OPS_FIRST_COL + 20)).Clear

    EnsureNamedTable ws, TBL_PARTS_NAME, DATA_PARTS_HEADER_ROW, DATA_PARTS_FIRST_COL, PartsHeaderNames()
    EnsureNamedTable ws, TBL_OPS_NAME, DATA_OPS_HEADER_ROW, DATA_OPS_FIRST_COL, OpsStoreHeaderNames()
    WriteMappedTableBody ListObjectByName(ws, TBL_PARTS_NAME), PartsHeaderNames(), partsHdrs, partsBody
    WriteMappedTableBody ListObjectByName(ws, TBL_OPS_NAME), OpsStoreHeaderNames(), opsHdrs, opsBody
End Sub

Private Function TableHeaderNames(ByVal lo As ListObject) As Variant
    Dim i As Long
    Dim result() As String

    ReDim result(1 To lo.ListColumns.Count)
    For i = 1 To lo.ListColumns.Count
        result(i) = Trim$(lo.ListColumns(i).Name)
    Next i
    TableHeaderNames = result
End Function

Private Sub WriteMappedTableBody( _
    ByVal lo As ListObject, _
    ByVal destHeaders As Variant, _
    ByVal srcHeaders As Variant, _
    ByVal srcBody As Variant)

    Dim rowCount As Long
    Dim destColCount As Long
    Dim r As Long
    Dim c As Long
    Dim srcCol As Long
    Dim output() As Variant
    Dim destName As String

    If IsEmpty(srcBody) Then Exit Sub
    If Not IsArray(srcHeaders) Then Exit Sub
    If Not IsArray(srcBody) Then Exit Sub

    rowCount = UBound(srcBody, 1)
    destColCount = UBound(destHeaders) - LBound(destHeaders) + 1
    If rowCount < 1 Then Exit Sub

    ReDim output(1 To rowCount, 1 To destColCount)
    For r = 1 To rowCount
        For c = LBound(destHeaders) To UBound(destHeaders)
            destName = CStr(destHeaders(c))
            srcCol = IndexOfHeaderName(srcHeaders, destName)
            If srcCol = 0 And StrComp(destName, COL_PARTS_HOME_FFA, vbTextCompare) = 0 Then
                srcCol = IndexOfHeaderName(srcHeaders, COL_PARTS_FFAS_LEGACY)
                If srcCol > 0 Then
                    output(r, c - LBound(destHeaders) + 1) = FirstCommaToken(CStr(Nz(srcBody(r, srcCol))))
                    GoTo NextDestCol
                End If
            End If
            If srcCol > 0 Then output(r, c - LBound(destHeaders) + 1) = srcBody(r, srcCol)
NextDestCol:
        Next c
    Next r

    lo.Resize lo.HeaderRowRange.Resize(rowCount + 1, destColCount)
    lo.DataBodyRange.Value2 = output
End Sub

Private Function IndexOfHeaderName(ByVal headers As Variant, ByVal headerName As String) As Long
    Dim i As Long

    For i = LBound(headers) To UBound(headers)
        If StrComp(Trim$(CStr(headers(i))), headerName, vbTextCompare) = 0 Then
            IndexOfHeaderName = i
            Exit Function
        End If
    Next i
End Function

Private Function FirstCommaToken(ByVal textValue As String) As String
    Dim parts As Variant

    If Len(textValue) = 0 Then Exit Function
    parts = Split(textValue, ",")
    FirstCommaToken = Trim$(CStr(parts(LBound(parts))))
End Function

Private Sub ConvertLegacyHomeFfaColumn(ByVal tbl As ListObject)
    Dim col As ListColumn
    Dim rowIndex As Long
    Dim rawValue As String

    If tbl Is Nothing Then Exit Sub
    Set col = ListColumnByHeader(tbl, COL_PARTS_HOME_FFA)
    If col Is Nothing Then Set col = ListColumnByHeader(tbl, COL_PARTS_FFAS_LEGACY)
    If col Is Nothing Then Exit Sub
    If col.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = 1 To col.DataBodyRange.Rows.Count
        rawValue = Trim$(CStr(Nz(col.DataBodyRange.Cells(rowIndex, 1).Value)))
        If InStr(1, rawValue, ",") > 0 Then
            col.DataBodyRange.Cells(rowIndex, 1).Value = FirstCommaToken(rawValue)
        End If
    Next rowIndex
End Sub

Private Sub ClearLegacyFfaChecklist(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim leftover As Range

    lastRow = Application.WorksheetFunction.Max( _
        FastLastUsedRowInColumn(ws, "C"), _
        FastLastUsedRowInColumn(ws, "D"), _
        LIST_START_ROW)

    On Error Resume Next
    ws.Range(ws.Cells(LIST_HEADER_ROW, "D"), ws.Cells(lastRow, "D")).CellControl.Clear
    On Error GoTo 0

    If lastRow > LIST_START_ROW Then
        Set leftover = ws.Range(ws.Cells(LIST_START_ROW + 1, "C"), ws.Cells(lastRow, "D"))
        leftover.ClearContents
        leftover.Interior.ColorIndex = xlNone
        leftover.Borders.LineStyle = xlNone
    End If

    ws.Range("D8").ClearContents
    ws.Range("D8").Font.Bold = False
    ws.Range("D9").ClearContents
    ws.Range("D9").Interior.ColorIndex = xlNone
    ws.Range("D9").Borders.LineStyle = xlNone
End Sub

Private Sub ApplyHomeFfaValidation(ByVal ws As Worksheet)
    Dim formulaText As String

    EnsureReferencesSheet
    formulaText = "=OFFSET(" & REFERENCES_SHEET_NAME & "!$" & REFS_FFA_COLUMN & "$" & CStr(REFS_FIRST_DATA_ROW) & _
        ",0,0,MAX(1,COUNTA(" & REFERENCES_SHEET_NAME & "!$" & REFS_FFA_COLUMN & ":$" & REFS_FFA_COLUMN & ")-1),1)"

    On Error Resume Next
    ws.Range(HOME_FFA_VALUE_CELL).Validation.Delete
    On Error GoTo 0

    On Error Resume Next
    ws.Range(HOME_FFA_VALUE_CELL).Validation.Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, _
        Formula1:=formulaText
    ws.Range(HOME_FFA_VALUE_CELL).Validation.IgnoreBlank = True
    ws.Range(HOME_FFA_VALUE_CELL).Validation.InCellDropdown = True
    On Error GoTo 0
End Sub

Private Sub EnsureNamedTable( _
    ByVal ws As Worksheet, _
    ByVal tableName As String, _
    ByVal headerRow As Long, _
    ByVal firstCol As Long, _
    ByVal headers As Variant)

    Dim lo As ListObject
    Dim colCount As Long
    Dim i As Long
    Dim tableRange As Range

    colCount = UBound(headers) - LBound(headers) + 1
    For i = LBound(headers) To UBound(headers)
        ws.Cells(headerRow, firstCol + i - LBound(headers)).Value = headers(i)
        ws.Cells(headerRow, firstCol + i - LBound(headers)).Font.Bold = True
    Next i

    Set lo = ListObjectByName(ws, tableName)
    If lo Is Nothing Then
        Set tableRange = ws.Range( _
            ws.Cells(headerRow, firstCol), _
            ws.Cells(headerRow + 1, firstCol + colCount - 1))
        Set lo = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
        lo.Name = tableName
    End If
End Sub
