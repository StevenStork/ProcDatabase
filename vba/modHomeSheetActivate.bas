Attribute VB_Name = "modHomeSheetActivate"
Option Explicit

Private Const HOME_SHEET_NAME As String = "Home"
Private Const CATEGORY_CELL As String = "A1"
Private Const PART_LABEL_VALUE As String = "Part"

Private Const CATEGORY_BUTTON_ANCHOR_CELL As String = "K5"
Private Const CATEGORY_BUTTON_NAME_PREFIX As String = "btnSheetCat_"

Private Const FFA_BUTTON_ANCHOR_CELL As String = "O5"
Private Const FFA_BUTTON_NAME_PREFIX As String = "btnPartFFA_"
Private Const REFERENCES_SHEET_NAME As String = "References"
Private Const REFERENCES_FFA_COLUMN As String = "B"
Private Const REFERENCES_FACTORY_COLUMN As String = "C"
Private Const REFERENCES_FFA_START_ROW As Long = 2

Private Const BUTTON_WIDTH As Double = 140
Private Const BUTTON_HEIGHT As Double = 24
Private Const BUTTON_VERTICAL_GAP As Double = 8
Private Const BUTTON_HORIZONTAL_GAP As Double = 8
Private Const BUTTON_CAPTION_SHOW_PREFIX As String = "Show "
Private Const BUTTON_CAPTION_HIDE_PREFIX As String = "Hide "

Private Const FORM_BUTTON_NAME_PREFIX As String = "btnHomeForm_"
Private Const FORM_BUTTON_ANCHOR_CELL As String = "S5"
Private Const FORM_BUTTON_LAST_COLUMN As String = "Y"

Private Const HOME_PART_TABLE_HEADER_ROW As Long = 5
Private Const HOME_PART_TABLE_FIRST_DATA_ROW As Long = 6
Private Const HOME_PART_TABLE_FIRST_COLUMN As String = "C"
Private Const HOME_PART_TABLE_LAST_COLUMN As String = "I"
Private Const HOME_PART_TABLE_BASE_PART_COLUMN As String = "C"
Private Const HOME_PART_TABLE_ACTIVE_COLUMN As String = "D"
Private Const HOME_PART_TABLE_DATE_COLUMN As String = "E"
Private Const HOME_PART_TABLE_DAYS_COLUMN As String = "F"
Private Const HOME_PART_TABLE_HIGHLIGHT_COLUMN_G As String = "G"
Private Const HOME_PART_TABLE_FFA_COLUMN As String = "H"
Private Const HOME_PART_TABLE_FACTORY_COLUMN As String = "I"
Private Const PART_SHEET_BASE_PART_CELL As String = "C2"

' Opaque Home activate cache (hidden via ;;;). Bump schema when formatting logic changes.
Private Const HOME_CACHE_CELL As String = "A2"
Private Const HOME_CACHE_SCHEMA As String = "3"
Private Const REFS_LABEL_VALUE As String = "Refs"
Private Const EXPORT_LABEL_VALUE As String = "Export"

' Per-activate session caches (reset at the start of each Home activate).
Private g_partSheetsByBase As Object
Private g_ffaCheckedMap As Object
Private g_ffaVisibleCheckedMap As Object
Private g_categoryVisibleMap As Object
Private g_markedFfasByBasePartSession As Object

' Call from ThisWorkbook.Workbook_SheetActivate:
'   HandleHomeSheetActivate Sh
Public Sub HandleHomeSheetActivate(ByVal Sh As Object)
    Dim wsHome As Worksheet
    Dim lastRow As Long
    Dim needsFormat As Boolean

    On Error GoTo CleanUp

    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    If StrComp(Sh.Name, HOME_SHEET_NAME, vbTextCompare) <> 0 Then Exit Sub

    Set wsHome = Sh
    EnsureReferencesSheet
    EnsureDataSheet
    ClearLegacyCacheCell wsHome
    ResetHomeSessionCaches

    lastRow = HomePartTableLastRow(wsHome)
    needsFormat = Not HomeListHashIsCurrent()
    If lastRow >= HOME_PART_TABLE_FIRST_DATA_ROW Then
        If InStr(1, wsHome.Cells(HOME_PART_TABLE_FIRST_DATA_ROW, HOME_PART_TABLE_DAYS_COLUMN).Formula, "TODAY()", vbTextCompare) = 0 Then
            needsFormat = True
        End If
    End If

    OptimizeExcel True
    SyncCategoryToggleButtons wsHome
    SyncFfaToggleButtons wsHome
    SyncFormActionButtons wsHome

    If needsFormat Then
        FormatHomePartTable wsHome, lastRow
        RebuildHomeFromStore
    End If

CleanUp:
    OptimizeExcel False
End Sub

Private Sub ClearLegacyCacheCell(ByVal ws As Worksheet)
    On Error Resume Next
    If ws.Range("A2").NumberFormat = ";;;" Then
        ws.Range("A2").ClearContents
        ws.Range("A2").NumberFormat = "General"
    End If
    On Error GoTo 0
End Sub

Private Sub ResetHomeSessionCaches()
    Set g_partSheetsByBase = Nothing
    Set g_ffaCheckedMap = Nothing
    Set g_ffaVisibleCheckedMap = Nothing
    Set g_categoryVisibleMap = Nothing
    Set g_markedFfasByBasePartSession = Nothing
End Sub

Private Function BuildHomeActivateCacheKey( _
    ByVal ws As Worksheet, _
    ByVal lastRow As Long, _
    ByVal markedFfasByBasePart As Object) As String

    BuildHomeActivateCacheKey = _
        HOME_CACHE_SCHEMA & Chr$(31) & _
        BuildHomeTableInputSignature(ws, lastRow) & Chr$(31) & _
        CollectAndSignMarkedFfas(ws, lastRow, markedFfasByBasePart) & Chr$(31) & _
        BuildReferencesFactorySignature() & Chr$(31) & _
        BuildHomeButtonSignature()
End Function

Private Function HomeActivateCacheIsCurrent( _
    ByVal ws As Worksheet, _
    ByVal cacheKey As String, _
    ByVal lastRow As Long) As Boolean

    If StrComp(CStr(Nz(ws.Range(HOME_CACHE_CELL).Value2)), cacheKey, vbBinaryCompare) <> 0 Then
        Exit Function
    End If

    If lastRow >= HOME_PART_TABLE_FIRST_DATA_ROW Then
        If InStr(1, ws.Cells(HOME_PART_TABLE_FIRST_DATA_ROW, HOME_PART_TABLE_DAYS_COLUMN).Formula, "TODAY()", vbTextCompare) = 0 Then
            Exit Function
        End If
    End If

    HomeActivateCacheIsCurrent = True
End Function

Private Sub WriteHomeActivateCache(ByVal ws As Worksheet, ByVal cacheKey As String)
    With ws.Range(HOME_CACHE_CELL)
        .NumberFormat = ";;;"
        .Value2 = cacheKey
    End With
End Sub

Private Function BuildHomeTableInputSignature(ByVal ws As Worksheet, ByVal lastRow As Long) As String
    Dim rowIndex As Long
    Dim parts() As String
    Dim partCount As Long

    If lastRow < HOME_PART_TABLE_FIRST_DATA_ROW Then
        BuildHomeTableInputSignature = CStr(lastRow)
        Exit Function
    End If

    partCount = lastRow - HOME_PART_TABLE_FIRST_DATA_ROW + 1
    ReDim parts(1 To partCount)

    For rowIndex = HOME_PART_TABLE_FIRST_DATA_ROW To lastRow
        parts(rowIndex - HOME_PART_TABLE_FIRST_DATA_ROW + 1) = _
            Trim$(CStr(ws.Cells(rowIndex, HOME_PART_TABLE_BASE_PART_COLUMN).Value)) & Chr$(30) & _
            CStr(IsActiveFlag(ws.Cells(rowIndex, HOME_PART_TABLE_ACTIVE_COLUMN).Value)) & Chr$(30) & _
            CStr(ws.Cells(rowIndex, HOME_PART_TABLE_DATE_COLUMN).Value2) & Chr$(30) & _
            Trim$(CStr(ws.Cells(rowIndex, HOME_PART_TABLE_HIGHLIGHT_COLUMN_G).Value))
    Next rowIndex

    BuildHomeTableInputSignature = CStr(lastRow) & Chr$(31) & Join(parts, Chr$(29))
End Function

Private Function CollectAndSignMarkedFfas( _
    ByVal ws As Worksheet, _
    ByVal lastRow As Long, _
    ByVal markedFfasByBasePart As Object) As String

    Dim rowIndex As Long
    Dim basePart As String
    Dim ffaList As String
    Dim parts As Collection
    Dim sigParts() As String
    Dim i As Long

    Set parts = New Collection
    If lastRow < HOME_PART_TABLE_FIRST_DATA_ROW Then
        CollectAndSignMarkedFfas = vbNullString
        Exit Function
    End If

    EnsurePartSheetIndex
    EnsureFfaPresenceMaps

    For rowIndex = HOME_PART_TABLE_FIRST_DATA_ROW To lastRow
        basePart = Trim$(CStr(ws.Cells(rowIndex, HOME_PART_TABLE_BASE_PART_COLUMN).Value))
        If Len(basePart) = 0 Then GoTo NextRow
        If Not IsActiveFlag(ws.Cells(rowIndex, HOME_PART_TABLE_ACTIVE_COLUMN).Value) Then GoTo NextRow

        If markedFfasByBasePart.Exists(basePart) Then
            ffaList = CStr(markedFfasByBasePart(basePart))
        Else
            ffaList = GetMarkedFfaListForBasePart(basePart)
            markedFfasByBasePart.Add basePart, ffaList
        End If

        parts.Add basePart & "=" & ffaList
NextRow:
    Next rowIndex

    If parts.Count = 0 Then
        CollectAndSignMarkedFfas = vbNullString
        Exit Function
    End If

    ReDim sigParts(1 To parts.Count)
    For i = 1 To parts.Count
        sigParts(i) = CStr(parts(i))
    Next i

    CollectAndSignMarkedFfas = Join(sigParts, Chr$(29))
End Function

Private Function BuildReferencesFactorySignature() As String
    Dim wsReferences As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim parts() As String
    Dim partCount As Long

    Set wsReferences = ThisWorkbook.Worksheets(REFERENCES_SHEET_NAME)
    lastRow = LastUsedRowInColumn(wsReferences, REFERENCES_FFA_COLUMN)
    If lastRow < REFERENCES_FFA_START_ROW Then
        BuildReferencesFactorySignature = "0"
        Exit Function
    End If

    partCount = lastRow - REFERENCES_FFA_START_ROW + 1
    ReDim parts(1 To partCount)

    For rowIndex = REFERENCES_FFA_START_ROW To lastRow
        parts(rowIndex - REFERENCES_FFA_START_ROW + 1) = _
            Trim$(CStr(wsReferences.Cells(rowIndex, REFERENCES_FFA_COLUMN).Value)) & Chr$(30) & _
            Trim$(CStr(wsReferences.Cells(rowIndex, REFERENCES_FACTORY_COLUMN).Value))
    Next rowIndex

    BuildReferencesFactorySignature = CStr(lastRow) & Chr$(31) & Join(parts, Chr$(29))
End Function

Private Function BuildHomeButtonSignature() As String
    Dim categories As Object
    Dim ffaValues As Object
    Dim categoryKeys() As String
    Dim ffaKeys() As String
    Dim i As Long
    Dim labelName As String
    Dim parts As Collection
    Dim sigParts() As String

    Set categories = CollectSheetCategories(HomeCategory())
    Set ffaValues = CollectReferenceFfas()
    EnsureCategoryVisibilityMap
    EnsureFfaPresenceMaps

    Set parts = New Collection
    categoryKeys = DictionaryKeysToSortedArray(categories)
    If IsArrayInitialized(categoryKeys) Then
        For i = LBound(categoryKeys) To UBound(categoryKeys)
            labelName = categoryKeys(i)
            parts.Add "C:" & labelName & "=" & CStr(CategoryHasVisibleSheets(labelName))
        Next i
    End If

    ffaKeys = DictionaryKeysToSortedArray(ffaValues)
    If IsArrayInitialized(ffaKeys) Then
        For i = LBound(ffaKeys) To UBound(ffaKeys)
            labelName = ffaKeys(i)
            parts.Add "F:" & labelName & "=" & CStr(FfaHasVisiblePartSheets(labelName))
        Next i
    End If

    If parts.Count = 0 Then
        BuildHomeButtonSignature = vbNullString
        Exit Function
    End If

    ReDim sigParts(1 To parts.Count)
    For i = 1 To parts.Count
        sigParts(i) = CStr(parts(i))
    Next i

    BuildHomeButtonSignature = Join(sigParts, Chr$(29))
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

' OnAction handler for sheet-category toggle buttons on the Home sheet.
Public Sub ToggleSheetCategoryVisibility()
    Dim callerName As String
    Dim category As String
    Dim ws As Worksheet
    Dim wsHome As Worksheet
    Dim targetState As XlSheetVisibility

    On Error GoTo CleanUp
    OptimizeExcel True
    ResetHomeSessionCaches

    Set wsHome = ThisWorkbook.Worksheets(HOME_SHEET_NAME)

    callerName = CStr(Application.Caller)
    category = LabelFromButtonName(callerName, CATEGORY_BUTTON_NAME_PREFIX, CollectSheetCategories(HomeCategory()))
    If Len(category) = 0 Then
        category = LabelFromButtonCaption(wsHome.Buttons(callerName).Caption)
    End If
    If Len(category) = 0 Then GoTo CleanUp

    If CategoryHasVisibleSheets(category) Then
        targetState = xlSheetHidden
    Else
        targetState = xlSheetVisible
    End If

    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, HOME_SHEET_NAME, vbTextCompare) <> 0 Then
            If IsToggleableSheetCategory(Trim$(CStr(ws.Range(CATEGORY_CELL).Value)), HomeCategory()) Then
                If StrComp(Trim$(CStr(ws.Range(CATEGORY_CELL).Value)), category, vbTextCompare) = 0 Then
                    ws.Visible = targetState
                End If
            End If
        End If
    Next ws

    UpdateToggleButtonCaption wsHome, CATEGORY_BUTTON_NAME_PREFIX, category, (targetState = xlSheetVisible)

CleanUp:
    On Error Resume Next
    ThisWorkbook.Worksheets(HOME_SHEET_NAME).Activate
    On Error GoTo 0
    OptimizeExcel False
End Sub

' OnAction handler for Part-sheet FFA toggle buttons on the Home sheet.
Public Sub TogglePartFfaVisibility()
    Dim callerName As String
    Dim ffaValue As String
    Dim ws As Worksheet
    Dim wsHome As Worksheet
    Dim targetState As XlSheetVisibility

    On Error GoTo CleanUp
    OptimizeExcel True
    ResetHomeSessionCaches

    Set wsHome = ThisWorkbook.Worksheets(HOME_SHEET_NAME)

    callerName = CStr(Application.Caller)
    ffaValue = LabelFromButtonName(callerName, FFA_BUTTON_NAME_PREFIX, CollectReferenceFfas())
    If Len(ffaValue) = 0 Then
        ffaValue = LabelFromButtonCaption(wsHome.Buttons(callerName).Caption)
    End If
    If Len(ffaValue) = 0 Then GoTo CleanUp

    If Not AnyPartSheetHasActiveFfa(ffaValue) Then
        OptimizeExcel False
        MsgBox "No parts have the FFA """ & ffaValue & """ assigned to them.", _
            vbInformation, "FFA Visibility"
        On Error Resume Next
        wsHome.Activate
        On Error GoTo 0
        Exit Sub
    End If

    If FfaHasVisiblePartSheets(ffaValue) Then
        targetState = xlSheetHidden
    Else
        targetState = xlSheetVisible
    End If

    EnsurePartSheetIndex
    EnsureFfaPresenceMaps
    For Each ws In ThisWorkbook.Worksheets
        If IsPartSheet(ws) Then
            If MarkedFfaListContains(GetMarkedFfaListForBasePart( _
                Trim$(CStr(ws.Range(PART_SHEET_BASE_PART_CELL).Value))), ffaValue) Then
                ws.Visible = targetState
            End If
        End If
    Next ws

    UpdateToggleButtonCaption wsHome, FFA_BUTTON_NAME_PREFIX, ffaValue, (targetState = xlSheetVisible)

CleanUp:
    On Error Resume Next
    ThisWorkbook.Worksheets(HOME_SHEET_NAME).Activate
    On Error GoTo 0
    OptimizeExcel False
End Sub

Private Sub SyncCategoryToggleButtons(ByVal wsHome As Worksheet)
    Dim categories As Object

    Set categories = CollectSheetCategories(HomeCategory())
    SyncToggleButtons _
        wsHome, _
        categories, _
        CATEGORY_BUTTON_NAME_PREFIX, _
        CATEGORY_BUTTON_ANCHOR_CELL, _
        "ToggleSheetCategoryVisibility", _
        True
End Sub

Private Sub SyncFfaToggleButtons(ByVal wsHome As Worksheet)
    Dim ffaValues As Object

    Set ffaValues = CollectReferenceFfas()
    SyncToggleButtons _
        wsHome, _
        ffaValues, _
        FFA_BUTTON_NAME_PREFIX, _
        FFA_BUTTON_ANCHOR_CELL, _
        "TogglePartFfaVisibility", _
        False
End Sub

Private Sub SyncFormActionButtons(ByVal wsHome As Worksheet)
    Dim captions() As String
    Dim actions() As String
    Dim wantedNames As Object
    Dim btn As Button
    Dim buttonsToDelete As Collection
    Dim buttonKey As Variant
    Dim i As Long
    Dim buttonName As String

    FormActionButtonDefs captions, actions
    Set wantedNames = CreateObject("Scripting.Dictionary")
    wantedNames.CompareMode = vbTextCompare

    For i = LBound(captions) To UBound(captions)
        buttonName = FORM_BUTTON_NAME_PREFIX & MakeNameSafe(captions(i))
        wantedNames.Add buttonName, i

        If Not ButtonExists(wsHome, buttonName) Then
            AddFormActionButton wsHome, buttonName, captions(i), actions(i)
        Else
            wsHome.Buttons(buttonName).OnAction = actions(i)
            wsHome.Buttons(buttonName).Caption = captions(i)
        End If
    Next i

    Set buttonsToDelete = New Collection
    For Each btn In wsHome.Buttons
        If Left$(btn.Name, Len(FORM_BUTTON_NAME_PREFIX)) = FORM_BUTTON_NAME_PREFIX Then
            If Not wantedNames.Exists(btn.Name) Then buttonsToDelete.Add btn.Name
        End If
    Next btn

    For Each buttonKey In buttonsToDelete
        wsHome.Buttons(CStr(buttonKey)).Delete
    Next buttonKey

    LayoutFormActionButtons wsHome, captions
End Sub

' Captions/OnAction pairs for Home form launchers. Append new entries here;
' layout fills S:Y left-to-right from row 5, then wraps to the next row.
Private Sub FormActionButtonDefs(ByRef captions() As String, ByRef actions() As String)
    ReDim captions(0 To 2)
    ReDim actions(0 To 2)
    captions(0) = "Update References"
    actions(0) = "ShowUpdateReferences"
    captions(1) = "Update Exports"
    actions(1) = "ShowUpdateExportSheets"
    captions(2) = "Refresh All"
    actions(2) = "RefreshAllProcDatabase"
End Sub

Private Sub AddFormActionButton( _
    ByVal wsHome As Worksheet, _
    ByVal buttonName As String, _
    ByVal captionText As String, _
    ByVal onActionName As String)

    Dim btn As Button
    Dim anchor As Range

    Set anchor = wsHome.Range(FORM_BUTTON_ANCHOR_CELL)
    Set btn = wsHome.Buttons.Add(anchor.Left, anchor.Top, BUTTON_WIDTH, BUTTON_HEIGHT)
    btn.Name = buttonName
    btn.OnAction = onActionName
    btn.Caption = captionText
End Sub

Private Sub LayoutFormActionButtons(ByVal wsHome As Worksheet, ByRef captions() As String)
    Dim i As Long
    Dim buttonName As String
    Dim btn As Button
    Dim anchor As Range
    Dim lastCol As Range
    Dim leftPos As Double
    Dim topPos As Double
    Dim rightLimit As Double
    Dim colIndex As Long
    Dim rowIndex As Long

    Set anchor = wsHome.Range(FORM_BUTTON_ANCHOR_CELL)
    Set lastCol = wsHome.Cells(anchor.Row, FORM_BUTTON_LAST_COLUMN)
    rightLimit = lastCol.Left + lastCol.Width

    colIndex = 0
    rowIndex = 0
    For i = LBound(captions) To UBound(captions)
        buttonName = FORM_BUTTON_NAME_PREFIX & MakeNameSafe(captions(i))
        If ButtonExists(wsHome, buttonName) Then
            leftPos = anchor.Left + colIndex * (BUTTON_WIDTH + BUTTON_HORIZONTAL_GAP)
            If colIndex > 0 And (leftPos + BUTTON_WIDTH) > rightLimit Then
                colIndex = 0
                rowIndex = rowIndex + 1
                leftPos = anchor.Left
            End If

            topPos = anchor.Top + rowIndex * (BUTTON_HEIGHT + BUTTON_VERTICAL_GAP)
            Set btn = wsHome.Buttons(buttonName)
            btn.Left = leftPos
            btn.Top = topPos
            btn.Width = BUTTON_WIDTH
            btn.Height = BUTTON_HEIGHT
            colIndex = colIndex + 1
        End If
    Next i
End Sub

' Formats the Home part table (C:I from header row 5 through the last used row):
' thin + medium borders (including medium header border), center alignment, D/G fill,
' E short dates, F days-since formulas with RAG colors. H/I come from tblParts.
Private Sub FormatHomePartTable(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim tableRange As Range
    Dim headerRange As Range
    Dim dataLastRow As Long
    Dim rowIndex As Long
    Dim formulaDays As Variant

    If lastRow < HOME_PART_TABLE_HEADER_ROW Then Exit Sub

    ws.Cells(HOME_PART_TABLE_HEADER_ROW, HOME_PART_TABLE_BASE_PART_COLUMN).Value = HEADER_BASE_PART
    ws.Cells(HOME_PART_TABLE_HEADER_ROW, HOME_PART_TABLE_ACTIVE_COLUMN).Value = "Active Part"
    ws.Cells(HOME_PART_TABLE_HEADER_ROW, HOME_PART_TABLE_DATE_COLUMN).Value = "Date"
    ws.Cells(HOME_PART_TABLE_HEADER_ROW, HOME_PART_TABLE_DAYS_COLUMN).Value = "Days"
    ws.Cells(HOME_PART_TABLE_HEADER_ROW, HOME_PART_TABLE_HIGHLIGHT_COLUMN_G).Value = "Highlight"
    ws.Cells(HOME_PART_TABLE_HEADER_ROW, HOME_PART_TABLE_FFA_COLUMN).Value = HOME_FFA_LABEL
    ws.Cells(HOME_PART_TABLE_HEADER_ROW, HOME_PART_TABLE_FACTORY_COLUMN).Value = "Factories"

    Set tableRange = ws.Range( _
        ws.Cells(HOME_PART_TABLE_HEADER_ROW, HOME_PART_TABLE_FIRST_COLUMN), _
        ws.Cells(lastRow, HOME_PART_TABLE_LAST_COLUMN))
    Set headerRange = ws.Range( _
        ws.Cells(HOME_PART_TABLE_HEADER_ROW, HOME_PART_TABLE_FIRST_COLUMN), _
        ws.Cells(HOME_PART_TABLE_HEADER_ROW, HOME_PART_TABLE_LAST_COLUMN))

    ApplyHomeListBorders tableRange
    headerRange.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, ColorIndex:=xlAutomatic

    tableRange.HorizontalAlignment = xlCenter
    tableRange.VerticalAlignment = xlCenter
    headerRange.Interior.ColorIndex = xlNone

    If lastRow < HOME_PART_TABLE_FIRST_DATA_ROW Then Exit Sub

    dataLastRow = lastRow

    ws.Range( _
        ws.Cells(HOME_PART_TABLE_FIRST_DATA_ROW, HOME_PART_TABLE_ACTIVE_COLUMN), _
        ws.Cells(dataLastRow, HOME_PART_TABLE_ACTIVE_COLUMN)).Interior.Color = RGB(213, 229, 249)

    ws.Range( _
        ws.Cells(HOME_PART_TABLE_FIRST_DATA_ROW, HOME_PART_TABLE_HIGHLIGHT_COLUMN_G), _
        ws.Cells(dataLastRow, HOME_PART_TABLE_HIGHLIGHT_COLUMN_G)).Interior.Color = RGB(213, 229, 249)

    ws.Range( _
        ws.Cells(HOME_PART_TABLE_FIRST_DATA_ROW, HOME_PART_TABLE_DATE_COLUMN), _
        ws.Cells(dataLastRow, HOME_PART_TABLE_DATE_COLUMN)).NumberFormat = "m/d/yyyy"

    ReDim formulaDays(1 To dataLastRow - HOME_PART_TABLE_FIRST_DATA_ROW + 1, 1 To 1)
    For rowIndex = HOME_PART_TABLE_FIRST_DATA_ROW To dataLastRow
        formulaDays(rowIndex - HOME_PART_TABLE_FIRST_DATA_ROW + 1, 1) = _
            "=IF(E" & CStr(rowIndex) & "="""","""",TODAY()-E" & CStr(rowIndex) & ")"
    Next rowIndex

    With ws.Range( _
        ws.Cells(HOME_PART_TABLE_FIRST_DATA_ROW, HOME_PART_TABLE_DAYS_COLUMN), _
        ws.Cells(dataLastRow, HOME_PART_TABLE_DAYS_COLUMN))
        .Formula = formulaDays
        .NumberFormat = "0"
    End With

    ApplyHomeStatusColumnFormats ws.Range( _
        ws.Cells(HOME_PART_TABLE_FIRST_DATA_ROW, HOME_PART_TABLE_DAYS_COLUMN), _
        ws.Cells(dataLastRow, HOME_PART_TABLE_DAYS_COLUMN))
End Sub

Private Function HomePartTableLastRow(ByVal ws As Worksheet) As Long
    Dim columnLetter As Variant
    Dim columnLastRow As Long
    Dim maxRow As Long

    maxRow = HOME_PART_TABLE_HEADER_ROW - 1

    ' Size from input columns only so formula/populated columns do not stretch the table.
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

Private Sub ApplyHomeListBorders(ByVal listRange As Range)
    With listRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    listRange.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, ColorIndex:=xlAutomatic
End Sub

' Column F: red when > 90, yellow when > 30, green otherwise.
Private Sub ApplyHomeStatusColumnFormats(ByVal targetRange As Range)
    Dim formatCondition As FormatCondition

    targetRange.FormatConditions.Delete

    Set formatCondition = targetRange.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlGreater, _
        Formula1:="90")
    formatCondition.StopIfTrue = True
    formatCondition.Interior.Color = RGB(255, 102, 102)

    Set formatCondition = targetRange.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlGreater, _
        Formula1:="30")
    formatCondition.StopIfTrue = True
    formatCondition.Interior.Color = RGB(255, 235, 132)

    Set formatCondition = targetRange.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlLessEqual, _
        Formula1:="30")
    formatCondition.StopIfTrue = True
    formatCondition.Interior.Color = RGB(146, 208, 80)
End Sub

Private Function GetMarkedFfaListForBasePart(ByVal basePart As String) As String
    EnsureFfaPresenceMaps

    If Len(basePart) = 0 Then
        GetMarkedFfaListForBasePart = vbNullString
        Exit Function
    End If

    If g_markedFfasByBasePartSession.Exists(basePart) Then
        GetMarkedFfaListForBasePart = CStr(g_markedFfasByBasePartSession(basePart))
    Else
        GetMarkedFfaListForBasePart = vbNullString
    End If
End Function

Private Function MarkedFfaListContains(ByVal ffaList As String, ByVal ffaValue As String) As Boolean
    If Len(ffaList) = 0 Or Len(ffaValue) = 0 Then Exit Function
    MarkedFfaListContains = (InStr(1, ", " & ffaList & ", ", ", " & ffaValue & ", ", vbTextCompare) > 0)
End Function

Private Function FindPartSheetByBasePart(ByVal basePart As String) As Worksheet
    EnsurePartSheetIndex

    If g_partSheetsByBase.Exists(basePart) Then
        Set FindPartSheetByBasePart = g_partSheetsByBase(basePart)
    End If
End Function

Private Sub EnsurePartSheetIndex()
    Dim ws As Worksheet
    Dim basePart As String

    If Not g_partSheetsByBase Is Nothing Then Exit Sub

    Set g_partSheetsByBase = CreateObject("Scripting.Dictionary")
    g_partSheetsByBase.CompareMode = vbTextCompare

    For Each ws In ThisWorkbook.Worksheets
        If IsPartSheet(ws) Then
            basePart = Trim$(CStr(ws.Range(PART_SHEET_BASE_PART_CELL).Value))
            If Len(basePart) = 0 Then basePart = ws.Name

            If Not g_partSheetsByBase.Exists(basePart) Then
                g_partSheetsByBase.Add basePart, ws
            End If

            If Not g_partSheetsByBase.Exists(ws.Name) Then
                g_partSheetsByBase.Add ws.Name, ws
            End If
        End If
    Next ws
End Sub

Private Sub EnsureFfaPresenceMaps()
    Dim sheetKey As Variant
    Dim visitedSheets As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim ffaValue As String
    Dim basePart As String
    Dim markedFfas As Object
    Dim ffaKeys() As String
    Dim ffaKey As Variant
    Dim keyIndex As Long
    Dim ffaList As String

    If Not g_ffaCheckedMap Is Nothing Then Exit Sub

    EnsurePartSheetIndex
    Set g_ffaCheckedMap = CreateObject("Scripting.Dictionary")
    g_ffaCheckedMap.CompareMode = vbTextCompare
    Set g_ffaVisibleCheckedMap = CreateObject("Scripting.Dictionary")
    g_ffaVisibleCheckedMap.CompareMode = vbTextCompare
    Set g_markedFfasByBasePartSession = CreateObject("Scripting.Dictionary")
    g_markedFfasByBasePartSession.CompareMode = vbTextCompare
    Set visitedSheets = CreateObject("Scripting.Dictionary")
    visitedSheets.CompareMode = vbTextCompare

    For Each sheetKey In g_partSheetsByBase.Keys
        Set ws = g_partSheetsByBase(sheetKey)
        If visitedSheets.Exists(ws.Name) Then GoTo NextSheet
        visitedSheets.Add ws.Name, True

        basePart = Trim$(CStr(ws.Range(PART_SHEET_BASE_PART_CELL).Value))
        If Len(basePart) = 0 Then basePart = ws.Name

        Set markedFfas = CreateObject("Scripting.Dictionary")
        markedFfas.CompareMode = vbTextCompare

        lastRow = 0
        ffaValue = PartHomeFfaValue(ws)
        If Len(ffaValue) > 0 Then
            If Not markedFfas.Exists(ffaValue) Then
                markedFfas.Add ffaValue, ffaValue
            End If
            If Not g_ffaCheckedMap.Exists(ffaValue) Then
                g_ffaCheckedMap.Add ffaValue, True
            End If
            If ws.Visible = xlSheetVisible Then
                If Not g_ffaVisibleCheckedMap.Exists(ffaValue) Then
                    g_ffaVisibleCheckedMap.Add ffaValue, True
                End If
            End If
        End If

        If markedFfas.Count = 0 Then
            ffaList = vbNullString
        Else
            ReDim ffaKeys(0 To markedFfas.Count - 1)
            keyIndex = 0
            For Each ffaKey In markedFfas.Keys
                ffaKeys(keyIndex) = CStr(ffaKey)
                keyIndex = keyIndex + 1
            Next ffaKey
            ffaList = Join(ffaKeys, ", ")
        End If

        If Not g_markedFfasByBasePartSession.Exists(basePart) Then
            g_markedFfasByBasePartSession.Add basePart, ffaList
        End If
NextSheet:
    Next sheetKey
End Sub

Private Sub EnsureCategoryVisibilityMap()
    Dim ws As Worksheet
    Dim categoryName As String
    Dim homeCategoryName As String

    If Not g_categoryVisibleMap Is Nothing Then Exit Sub

    homeCategoryName = HomeCategory()
    Set g_categoryVisibleMap = CreateObject("Scripting.Dictionary")
    g_categoryVisibleMap.CompareMode = vbTextCompare

    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, HOME_SHEET_NAME, vbTextCompare) <> 0 Then
            categoryName = Trim$(CStr(ws.Range(CATEGORY_CELL).Value))
            If IsToggleableSheetCategory(categoryName, homeCategoryName) Then
                If Not g_categoryVisibleMap.Exists(categoryName) Then
                    g_categoryVisibleMap.Add categoryName, False
                End If
                If ws.Visible = xlSheetVisible Then
                    g_categoryVisibleMap(categoryName) = True
                End If
            End If
        End If
    Next ws
End Sub

Private Function BuildFfaFactoryMap() As Object
    Dim wsReferences As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim ffaValue As String
    Dim factoryName As String
    Dim factoryMap As Object

    Set factoryMap = CreateObject("Scripting.Dictionary")
    factoryMap.CompareMode = vbTextCompare

    Set wsReferences = ThisWorkbook.Worksheets(REFERENCES_SHEET_NAME)
    lastRow = LastUsedRowInColumn(wsReferences, REFERENCES_FFA_COLUMN)
    If lastRow < REFERENCES_FFA_START_ROW Then
        Set BuildFfaFactoryMap = factoryMap
        Exit Function
    End If

    For rowIndex = REFERENCES_FFA_START_ROW To lastRow
        ffaValue = Trim$(CStr(wsReferences.Cells(rowIndex, REFERENCES_FFA_COLUMN).Value))
        factoryName = Trim$(CStr(wsReferences.Cells(rowIndex, REFERENCES_FACTORY_COLUMN).Value))
        If Len(ffaValue) > 0 Then
            If Not factoryMap.Exists(ffaValue) Then
                factoryMap.Add ffaValue, factoryName
            End If
        End If
    Next rowIndex

    Set BuildFfaFactoryMap = factoryMap
End Function

Private Function FactoriesForFfaList(ByVal ffaList As String, ByVal ffaFactoryMap As Object) As String
    Dim ffaParts As Variant
    Dim i As Long
    Dim ffaValue As String
    Dim factoryName As String
    Dim factories As Collection
    Dim factoryItems() As String
    Dim factoryIndex As Long

    If Len(Trim$(ffaList)) = 0 Then
        FactoriesForFfaList = vbNullString
        Exit Function
    End If

    Set factories = New Collection
    ffaParts = Split(ffaList, ",")

    For i = LBound(ffaParts) To UBound(ffaParts)
        ffaValue = Trim$(CStr(ffaParts(i)))
        If Len(ffaValue) > 0 Then
            factoryName = vbNullString
            If Not ffaFactoryMap Is Nothing Then
                If ffaFactoryMap.Exists(ffaValue) Then
                    factoryName = CStr(ffaFactoryMap(ffaValue))
                End If
            End If
            factories.Add factoryName
        End If
    Next i

    If factories.Count = 0 Then
        FactoriesForFfaList = vbNullString
        Exit Function
    End If

    ReDim factoryItems(0 To factories.Count - 1)
    For factoryIndex = 1 To factories.Count
        factoryItems(factoryIndex - 1) = CStr(factories(factoryIndex))
    Next factoryIndex

    FactoriesForFfaList = Join(factoryItems, ", ")
End Function

Private Sub SyncToggleButtons( _
    ByVal wsHome As Worksheet, _
    ByVal labels As Object, _
    ByVal namePrefix As String, _
    ByVal anchorCell As String, _
    ByVal onActionName As String, _
    ByVal isCategoryGroup As Boolean)

    Dim labelKeys() As String
    Dim btn As Button
    Dim i As Long
    Dim labelName As String
    Dim buttonName As String
    Dim buttonsToDelete As Collection
    Dim buttonKey As Variant
    Dim isVisible As Boolean

    Set buttonsToDelete = New Collection
    For Each btn In wsHome.Buttons
        If Left$(btn.Name, Len(namePrefix)) = namePrefix Then
            labelName = LabelFromButtonName(btn.Name, namePrefix, labels)
            If Len(labelName) = 0 Then labelName = LabelFromButtonCaption(btn.Caption)
            If Not labels.Exists(labelName) Then
                buttonsToDelete.Add btn.Name
            End If
        End If
    Next btn

    For Each buttonKey In buttonsToDelete
        wsHome.Buttons(CStr(buttonKey)).Delete
    Next buttonKey

    labelKeys = DictionaryKeysToSortedArray(labels)

    If IsArrayInitialized(labelKeys) Then
        For i = LBound(labelKeys) To UBound(labelKeys)
            labelName = labelKeys(i)
            buttonName = BuildButtonName(namePrefix, labelName)

            If isCategoryGroup Then
                isVisible = CategoryHasVisibleSheets(labelName)
            Else
                isVisible = FfaHasVisiblePartSheets(labelName)
            End If

            If Not ButtonExists(wsHome, buttonName) Then
                AddToggleButton wsHome, buttonName, labelName, anchorCell, onActionName, isVisible
            Else
                wsHome.Buttons(buttonName).OnAction = onActionName
                wsHome.Buttons(buttonName).Caption = BuildButtonCaption(labelName, isVisible)
            End If
        Next i
    End If

    LayoutToggleButtons wsHome, labelKeys, namePrefix, anchorCell
End Sub

Private Function CollectSheetCategories(ByVal homeCategory As String) As Object
    Dim categories As Object
    Dim ws As Worksheet
    Dim categoryName As String

    Set categories = CreateObject("Scripting.Dictionary")
    categories.CompareMode = vbTextCompare

    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, HOME_SHEET_NAME, vbTextCompare) <> 0 Then
            categoryName = Trim$(CStr(ws.Range(CATEGORY_CELL).Value))
            If IsToggleableSheetCategory(categoryName, homeCategory) Then
                If Not categories.Exists(categoryName) Then
                    categories.Add categoryName, categoryName
                End If
            End If
        End If
    Next ws

    Set CollectSheetCategories = categories
End Function

Private Function CollectReferenceFfas() As Object
    Dim ffaValues As Object
    Dim wsReferences As Worksheet
    Dim rowIndex As Long
    Dim lastRow As Long
    Dim ffaValue As String

    Set ffaValues = CreateObject("Scripting.Dictionary")
    ffaValues.CompareMode = vbTextCompare

    Set wsReferences = ThisWorkbook.Worksheets(REFERENCES_SHEET_NAME)
    lastRow = LastUsedRowInColumn(wsReferences, REFERENCES_FFA_COLUMN)

    For rowIndex = REFERENCES_FFA_START_ROW To lastRow
        ffaValue = Trim$(CStr(wsReferences.Cells(rowIndex, REFERENCES_FFA_COLUMN).Value))
        If Len(ffaValue) > 0 Then
            If Not ffaValues.Exists(ffaValue) Then
                ffaValues.Add ffaValue, ffaValue
            End If
        End If
    Next rowIndex

    Set CollectReferenceFfas = ffaValues
End Function

Private Function AnyPartSheetHasActiveFfa(ByVal ffaValue As String) As Boolean
    EnsureFfaPresenceMaps
    AnyPartSheetHasActiveFfa = g_ffaCheckedMap.Exists(ffaValue)
End Function

Private Sub AddToggleButton( _
    ByVal wsHome As Worksheet, _
    ByVal buttonName As String, _
    ByVal labelName As String, _
    ByVal anchorCell As String, _
    ByVal onActionName As String, _
    ByVal isVisible As Boolean)

    Dim btn As Button
    Dim anchor As Range

    Set anchor = wsHome.Range(anchorCell)
    Set btn = wsHome.Buttons.Add(anchor.Left, anchor.Top, BUTTON_WIDTH, BUTTON_HEIGHT)
    btn.Name = buttonName
    btn.OnAction = onActionName
    btn.Caption = BuildButtonCaption(labelName, isVisible)
End Sub

Private Sub UpdateToggleButtonCaption( _
    ByVal wsHome As Worksheet, _
    ByVal namePrefix As String, _
    ByVal labelName As String, _
    ByVal isVisible As Boolean)

    Dim buttonName As String

    buttonName = BuildButtonName(namePrefix, labelName)
    If ButtonExists(wsHome, buttonName) Then
        wsHome.Buttons(buttonName).Caption = BuildButtonCaption(labelName, isVisible)
    End If
End Sub

Private Function BuildButtonCaption(ByVal labelName As String, ByVal isVisible As Boolean) As String
    If isVisible Then
        BuildButtonCaption = BUTTON_CAPTION_HIDE_PREFIX & labelName
    Else
        BuildButtonCaption = BUTTON_CAPTION_SHOW_PREFIX & labelName
    End If
End Function

Private Function CategoryHasVisibleSheets(ByVal categoryName As String) As Boolean
    EnsureCategoryVisibilityMap
    If g_categoryVisibleMap.Exists(categoryName) Then
        CategoryHasVisibleSheets = CBool(g_categoryVisibleMap(categoryName))
    End If
End Function

Private Function FfaHasVisiblePartSheets(ByVal ffaValue As String) As Boolean
    EnsureFfaPresenceMaps
    FfaHasVisiblePartSheets = g_ffaVisibleCheckedMap.Exists(ffaValue)
End Function

Private Function PartSheetHasActiveFfa(ByVal ws As Worksheet, ByVal ffaValue As String) As Boolean
    If Len(ffaValue) = 0 Then Exit Function
    PartSheetHasActiveFfa = (StrComp(PartHomeFfaValue(ws), ffaValue, vbTextCompare) = 0)
End Function

Private Function IsToggleableSheetCategory(ByVal categoryName As String, ByVal homeCategory As String) As Boolean
    If Len(categoryName) = 0 Then Exit Function
    If StrComp(categoryName, homeCategory, vbTextCompare) = 0 Then Exit Function
    If StrComp(categoryName, REFS_LABEL_VALUE, vbTextCompare) = 0 Then Exit Function
    If StrComp(categoryName, DATA_LABEL_VALUE, vbTextCompare) = 0 Then Exit Function
    IsToggleableSheetCategory = True
End Function

Private Function IsPartSheet(ByVal ws As Worksheet) As Boolean
    IsPartSheet = (StrComp(Trim$(CStr(ws.Range(CATEGORY_CELL).Value)), PART_LABEL_VALUE, vbTextCompare) = 0)
End Function

Private Function HomeCategory() As String
    HomeCategory = Trim$(CStr(ThisWorkbook.Worksheets(HOME_SHEET_NAME).Range(CATEGORY_CELL).Value))
End Function

Private Function LabelFromButtonCaption(ByVal caption As String) As String
    Dim trimmedCaption As String

    trimmedCaption = Trim$(caption)

    If StrComp(Left$(trimmedCaption, Len(BUTTON_CAPTION_SHOW_PREFIX)), BUTTON_CAPTION_SHOW_PREFIX, vbTextCompare) = 0 Then
        LabelFromButtonCaption = Trim$(Mid$(trimmedCaption, Len(BUTTON_CAPTION_SHOW_PREFIX) + 1))
    ElseIf StrComp(Left$(trimmedCaption, Len(BUTTON_CAPTION_HIDE_PREFIX)), BUTTON_CAPTION_HIDE_PREFIX, vbTextCompare) = 0 Then
        LabelFromButtonCaption = Trim$(Mid$(trimmedCaption, Len(BUTTON_CAPTION_HIDE_PREFIX) + 1))
    Else
        LabelFromButtonCaption = trimmedCaption
    End If
End Function

Private Sub LayoutToggleButtons( _
    ByVal wsHome As Worksheet, _
    ByRef labelKeys() As String, _
    ByVal namePrefix As String, _
    ByVal anchorCell As String)

    Dim i As Long
    Dim buttonName As String
    Dim btn As Button
    Dim anchor As Range
    Dim topPos As Double

    If Not IsArrayInitialized(labelKeys) Then Exit Sub

    Set anchor = wsHome.Range(anchorCell)

    For i = LBound(labelKeys) To UBound(labelKeys)
        buttonName = BuildButtonName(namePrefix, labelKeys(i))
        If ButtonExists(wsHome, buttonName) Then
            Set btn = wsHome.Buttons(buttonName)
            topPos = anchor.Top + (i - LBound(labelKeys)) * (BUTTON_HEIGHT + BUTTON_VERTICAL_GAP)
            btn.Left = anchor.Left
            btn.Top = topPos
            btn.Width = BUTTON_WIDTH
            btn.Height = BUTTON_HEIGHT
        End If
    Next i
End Sub

Private Function BuildButtonName(ByVal namePrefix As String, ByVal labelName As String) As String
    BuildButtonName = namePrefix & MakeNameSafe(labelName)
End Function

Private Function LabelFromButtonName(ByVal buttonName As String, ByVal namePrefix As String, ByVal labels As Object) As String
    Dim encoded As String
    Dim labelKey As Variant

    If Left$(buttonName, Len(namePrefix)) <> namePrefix Then Exit Function
    encoded = Mid$(buttonName, Len(namePrefix) + 1)

    For Each labelKey In labels.Keys
        If StrComp(MakeNameSafe(CStr(labelKey)), encoded, vbTextCompare) = 0 Then
            LabelFromButtonName = CStr(labelKey)
            Exit Function
        End If
    Next labelKey
End Function

Private Function ButtonExists(ByVal ws As Worksheet, ByVal buttonName As String) As Boolean
    Dim btn As Button

    On Error Resume Next
    Set btn = ws.Buttons(buttonName)
    On Error GoTo 0

    ButtonExists = Not btn Is Nothing
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

Private Function MakeNameSafe(ByVal textValue As String) As String
    Dim cleaned As String
    Dim i As Long
    Dim character As String

    cleaned = vbNullString
    For i = 1 To Len(textValue)
        character = Mid$(textValue, i, 1)
        Select Case True
            Case character Like "[A-Za-z0-9]"
                cleaned = cleaned & character
            Case Else
                cleaned = cleaned & "_"
        End Select
    Next i

    If Len(cleaned) = 0 Then cleaned = "Label"
    If Len(cleaned) > 20 Then cleaned = Left$(cleaned, 20)
    MakeNameSafe = cleaned
End Function

Private Function LastUsedRowInColumn(ByVal ws As Worksheet, ByVal columnLetter As String) As Long
    Dim foundCell As Range
    Dim endUpRow As Long

    endUpRow = ws.Cells(ws.Rows.Count, columnLetter).End(xlUp).Row

    Set foundCell = ws.Columns(columnLetter).Find( _
        What:="*", _
        LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious)

    If foundCell Is Nothing Then
        LastUsedRowInColumn = endUpRow
    ElseIf foundCell.Row > endUpRow Then
        LastUsedRowInColumn = foundCell.Row
    Else
        LastUsedRowInColumn = endUpRow
    End If
End Function

Private Function DictionaryKeysToSortedArray(ByVal valueMap As Object) As String()
    Dim keys() As String
    Dim key As Variant
    Dim index As Long

    If valueMap.Count = 0 Then
        DictionaryKeysToSortedArray = EmptyStringArray()
        Exit Function
    End If

    ReDim keys(0 To valueMap.Count - 1)
    index = 0
    For Each key In valueMap.Keys
        keys(index) = CStr(key)
        index = index + 1
    Next key

    SortStringArray keys
    DictionaryKeysToSortedArray = keys
End Function

Private Function EmptyStringArray() As String()
    Dim emptyKeys() As String
    EmptyStringArray = emptyKeys
End Function

Private Function IsArrayInitialized(ByRef values() As String) As Boolean
    Dim upperBound As Long

    On Error Resume Next
    upperBound = UBound(values)
    IsArrayInitialized = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Sub SortStringArray(ByRef keys() As String)
    Dim i As Long
    Dim j As Long
    Dim tempKey As String

    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If StrComp(keys(i), keys(j), vbTextCompare) > 0 Then
                tempKey = keys(i)
                keys(i) = keys(j)
                keys(j) = tempKey
            End If
        Next j
    Next i
End Sub
