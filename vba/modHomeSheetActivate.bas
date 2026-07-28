Attribute VB_Name = "modHomeSheetActivate"
Option Explicit

Private Const HOME_SHEET_NAME As String = "Home"
Private Const CATEGORY_CELL As String = "A1"
Private Const PART_LABEL_VALUE As String = "Part"

Private Const CATEGORY_BUTTON_ANCHOR_CELL As String = "K5"
Private Const CATEGORY_BUTTON_NAME_PREFIX As String = "btnSheetCat_"

Private Const FFA_BUTTON_ANCHOR_CELL As String = "O5"
Private Const FFA_BUTTON_NAME_PREFIX As String = "btnPartFFA_"
Private Const PART_FFA_VALUE_COLUMN As String = "C"
Private Const PART_FFA_ACTIVE_COLUMN As String = "D"
Private Const PART_LIST_START_ROW As Long = 9
Private Const REFERENCES_SHEET_NAME As String = "References"
Private Const REFERENCES_FFA_COLUMN As String = "B"
Private Const REFERENCES_FFA_START_ROW As Long = 2

Private Const BUTTON_WIDTH As Double = 140
Private Const BUTTON_HEIGHT As Double = 24
Private Const BUTTON_VERTICAL_GAP As Double = 8
Private Const BUTTON_CAPTION_SHOW_PREFIX As String = "Show "
Private Const BUTTON_CAPTION_HIDE_PREFIX As String = "Hide "

Private Const HOME_PART_TABLE_HEADER_ROW As Long = 5
Private Const HOME_PART_TABLE_FIRST_DATA_ROW As Long = 6
Private Const HOME_PART_TABLE_FIRST_COLUMN As String = "C"
Private Const HOME_PART_TABLE_LAST_COLUMN As String = "F"
Private Const HOME_PART_TABLE_HIGHLIGHT_COLUMN As String = "D"
Private Const HOME_PART_TABLE_DATE_COLUMN As String = "E"
Private Const HOME_PART_TABLE_STATUS_COLUMN As String = "F"

' Call from ThisWorkbook.Workbook_SheetActivate:
'   HandleHomeSheetActivate Sh
Public Sub HandleHomeSheetActivate(ByVal Sh As Object)
    Dim wsHome As Worksheet

    On Error GoTo CleanUp

    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    If StrComp(Sh.Name, HOME_SHEET_NAME, vbTextCompare) <> 0 Then Exit Sub

    Set wsHome = Sh

    OptimizeExcel True
    SyncCategoryToggleButtons wsHome
    SyncFfaToggleButtons wsHome
    FormatHomePartTable wsHome

CleanUp:
    OptimizeExcel False
End Sub

' OnAction handler for sheet-category toggle buttons on the Home sheet.
Public Sub ToggleSheetCategoryVisibility()
    Dim callerName As String
    Dim category As String
    Dim ws As Worksheet
    Dim wsHome As Worksheet
    Dim targetState As XlSheetVisibility

    On Error GoTo CleanUp
    OptimizeExcel True

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
            If StrComp(Trim$(CStr(ws.Range(CATEGORY_CELL).Value)), category, vbTextCompare) = 0 Then
                ws.Visible = targetState
            End If
        End If
    Next ws

    UpdateToggleButtonCaption wsHome, CATEGORY_BUTTON_NAME_PREFIX, category, CategoryHasVisibleSheets(category)

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

    For Each ws In ThisWorkbook.Worksheets
        If IsPartSheet(ws) Then
            If PartSheetHasActiveFfa(ws, ffaValue) Then
                ws.Visible = targetState
            End If
        End If
    Next ws

    UpdateToggleButtonCaption wsHome, FFA_BUTTON_NAME_PREFIX, ffaValue, FfaHasVisiblePartSheets(ffaValue)

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

' Formats the Home part table (C:F from header row 5 through the last used row):
' thin + medium borders, center alignment, D fill, E short dates, F RAG colors.
Private Sub FormatHomePartTable(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim tableRange As Range
    Dim dataLastRow As Long

    lastRow = HomePartTableLastRow(ws)
    If lastRow < HOME_PART_TABLE_HEADER_ROW Then Exit Sub

    Set tableRange = ws.Range( _
        ws.Cells(HOME_PART_TABLE_HEADER_ROW, HOME_PART_TABLE_FIRST_COLUMN), _
        ws.Cells(lastRow, HOME_PART_TABLE_LAST_COLUMN))

    ApplyHomeListBorders tableRange
    tableRange.HorizontalAlignment = xlCenter
    tableRange.VerticalAlignment = xlCenter

    ' Keep the header row unfilled; style data rows when present.
    ws.Cells(HOME_PART_TABLE_HEADER_ROW, HOME_PART_TABLE_HIGHLIGHT_COLUMN).Interior.ColorIndex = xlNone

    If lastRow < HOME_PART_TABLE_FIRST_DATA_ROW Then Exit Sub

    dataLastRow = lastRow

    ws.Range( _
        ws.Cells(HOME_PART_TABLE_FIRST_DATA_ROW, HOME_PART_TABLE_HIGHLIGHT_COLUMN), _
        ws.Cells(dataLastRow, HOME_PART_TABLE_HIGHLIGHT_COLUMN)).Interior.Color = RGB(213, 229, 249)

    ws.Range( _
        ws.Cells(HOME_PART_TABLE_FIRST_DATA_ROW, HOME_PART_TABLE_DATE_COLUMN), _
        ws.Cells(dataLastRow, HOME_PART_TABLE_DATE_COLUMN)).NumberFormat = "m/d/yyyy"

    ApplyHomeStatusColumnFormats ws.Range( _
        ws.Cells(HOME_PART_TABLE_FIRST_DATA_ROW, HOME_PART_TABLE_STATUS_COLUMN), _
        ws.Cells(dataLastRow, HOME_PART_TABLE_STATUS_COLUMN))
End Sub

Private Function HomePartTableLastRow(ByVal ws As Worksheet) As Long
    Dim columnLetter As Variant
    Dim columnLastRow As Long
    Dim maxRow As Long

    maxRow = HOME_PART_TABLE_HEADER_ROW - 1

    For Each columnLetter In Array( _
        HOME_PART_TABLE_FIRST_COLUMN, _
        HOME_PART_TABLE_HIGHLIGHT_COLUMN, _
        HOME_PART_TABLE_DATE_COLUMN, _
        HOME_PART_TABLE_LAST_COLUMN)

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

            If Len(categoryName) > 0 Then
                If StrComp(categoryName, homeCategory, vbTextCompare) <> 0 Then
                    If Not categories.Exists(categoryName) Then
                        categories.Add categoryName, categoryName
                    End If
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
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If IsPartSheet(ws) Then
            If PartSheetHasActiveFfa(ws, ffaValue) Then
                AnyPartSheetHasActiveFfa = True
                Exit Function
            End If
        End If
    Next ws
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
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, HOME_SHEET_NAME, vbTextCompare) <> 0 Then
            If StrComp(Trim$(CStr(ws.Range(CATEGORY_CELL).Value)), categoryName, vbTextCompare) = 0 Then
                If ws.Visible = xlSheetVisible Then
                    CategoryHasVisibleSheets = True
                    Exit Function
                End If
            End If
        End If
    Next ws
End Function

Private Function FfaHasVisiblePartSheets(ByVal ffaValue As String) As Boolean
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If IsPartSheet(ws) Then
            If PartSheetHasActiveFfa(ws, ffaValue) Then
                If ws.Visible = xlSheetVisible Then
                    FfaHasVisiblePartSheets = True
                    Exit Function
                End If
            End If
        End If
    Next ws
End Function

Private Function PartSheetHasActiveFfa(ByVal ws As Worksheet, ByVal ffaValue As String) As Boolean
    Dim rowIndex As Long
    Dim lastRow As Long
    Dim cellFfa As String

    lastRow = LastUsedRowInColumn(ws, PART_FFA_VALUE_COLUMN)
    If lastRow < PART_LIST_START_ROW Then Exit Function

    For rowIndex = PART_LIST_START_ROW To lastRow
        cellFfa = Trim$(CStr(ws.Cells(rowIndex, PART_FFA_VALUE_COLUMN).Value))
        If Len(cellFfa) = 0 Then Exit For

        If StrComp(cellFfa, ffaValue, vbTextCompare) = 0 Then
            If IsActiveFlag(ws.Cells(rowIndex, PART_FFA_ACTIVE_COLUMN).Value) Then
                PartSheetHasActiveFfa = True
                Exit Function
            End If
        End If
    Next rowIndex
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
