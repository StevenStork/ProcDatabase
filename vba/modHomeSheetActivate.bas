Attribute VB_Name = "modHomeSheetActivate"
Option Explicit

Private Const HOME_SHEET_NAME As String = "Home"
Private Const CATEGORY_CELL As String = "A1"
Private Const BUTTON_ANCHOR_CELL As String = "K5"
Private Const BUTTON_NAME_PREFIX As String = "btnSheetCat_"
Private Const BUTTON_WIDTH As Double = 140
Private Const BUTTON_HEIGHT As Double = 24
Private Const BUTTON_VERTICAL_GAP As Double = 8
Private Const BUTTON_CAPTION_SHOW_PREFIX As String = "Show "
Private Const BUTTON_CAPTION_HIDE_PREFIX As String = "Hide "

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

CleanUp:
    OptimizeExcel False
End Sub

' OnAction handler for category toggle buttons created on the Home sheet.
Public Sub ToggleSheetCategoryVisibility()
    Dim callerName As String
    Dim category As String
    Dim ws As Worksheet
    Dim wsHome As Worksheet
    Dim anyVisible As Boolean
    Dim targetState As XlSheetVisibility

    On Error GoTo CleanUp
    OptimizeExcel True

    Set wsHome = ThisWorkbook.Worksheets(HOME_SHEET_NAME)

    callerName = CStr(Application.Caller)
    category = CategoryFromButtonName(callerName)
    If Len(category) = 0 Then
        category = CategoryFromButtonCaption(wsHome.Buttons(callerName).Caption)
    End If
    If Len(category) = 0 Then GoTo CleanUp

    anyVisible = CategoryHasVisibleSheets(category)

    If anyVisible Then
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

    UpdateCategoryButtonCaption wsHome, category

CleanUp:
    On Error Resume Next
    ThisWorkbook.Worksheets(HOME_SHEET_NAME).Activate
    On Error GoTo 0
    OptimizeExcel False
End Sub

Private Sub SyncCategoryToggleButtons(ByVal wsHome As Worksheet)
    Dim homeCategory As String
    Dim categories As Object
    Dim categoryKeys() As String
    Dim btn As Button
    Dim i As Long
    Dim categoryName As String
    Dim buttonName As String
    Dim buttonsToDelete As Collection
    Dim buttonKey As Variant

    homeCategory = Trim$(CStr(wsHome.Range(CATEGORY_CELL).Value))
    Set categories = CollectSheetCategories(homeCategory)

    ' Remove buttons for categories that no longer exist.
    Set buttonsToDelete = New Collection
    For Each btn In wsHome.Buttons
        If Left$(btn.Name, Len(BUTTON_NAME_PREFIX)) = BUTTON_NAME_PREFIX Then
            categoryName = CategoryFromButtonName(btn.Name)
            If Len(categoryName) = 0 Then categoryName = CategoryFromButtonCaption(btn.Caption)
            If Not categories.Exists(categoryName) Then
                buttonsToDelete.Add btn.Name
            End If
        End If
    Next btn

    For Each buttonKey In buttonsToDelete
        wsHome.Buttons(CStr(buttonKey)).Delete
    Next buttonKey

    categoryKeys = DictionaryKeysToSortedArray(categories)

    ' Create any missing category buttons and refresh Show/Hide captions.
    If IsArrayInitialized(categoryKeys) Then
        For i = LBound(categoryKeys) To UBound(categoryKeys)
            categoryName = categoryKeys(i)
            buttonName = BuildButtonName(categoryName)

            If Not ButtonExists(wsHome, buttonName) Then
                AddCategoryButton wsHome, buttonName, categoryName
            Else
                wsHome.Buttons(buttonName).OnAction = "ToggleSheetCategoryVisibility"
                UpdateCategoryButtonCaption wsHome, categoryName
            End If
        Next i
    End If

    LayoutCategoryButtons wsHome, categoryKeys
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

Private Sub AddCategoryButton(ByVal wsHome As Worksheet, ByVal buttonName As String, ByVal categoryName As String)
    Dim btn As Button
    Dim anchor As Range

    Set anchor = wsHome.Range(BUTTON_ANCHOR_CELL)
    Set btn = wsHome.Buttons.Add(anchor.Left, anchor.Top, BUTTON_WIDTH, BUTTON_HEIGHT)
    btn.Name = buttonName
    btn.OnAction = "ToggleSheetCategoryVisibility"
    btn.Caption = BuildButtonCaption(categoryName)
End Sub

Private Sub UpdateCategoryButtonCaption(ByVal wsHome As Worksheet, ByVal categoryName As String)
    Dim buttonName As String

    buttonName = BuildButtonName(categoryName)
    If ButtonExists(wsHome, buttonName) Then
        wsHome.Buttons(buttonName).Caption = BuildButtonCaption(categoryName)
    End If
End Sub

Private Function BuildButtonCaption(ByVal categoryName As String) As String
    If CategoryHasVisibleSheets(categoryName) Then
        BuildButtonCaption = BUTTON_CAPTION_HIDE_PREFIX & categoryName
    Else
        BuildButtonCaption = BUTTON_CAPTION_SHOW_PREFIX & categoryName
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

Private Function CategoryFromButtonCaption(ByVal caption As String) As String
    Dim trimmedCaption As String

    trimmedCaption = Trim$(caption)

    If StrComp(Left$(trimmedCaption, Len(BUTTON_CAPTION_SHOW_PREFIX)), BUTTON_CAPTION_SHOW_PREFIX, vbTextCompare) = 0 Then
        CategoryFromButtonCaption = Trim$(Mid$(trimmedCaption, Len(BUTTON_CAPTION_SHOW_PREFIX) + 1))
    ElseIf StrComp(Left$(trimmedCaption, Len(BUTTON_CAPTION_HIDE_PREFIX)), BUTTON_CAPTION_HIDE_PREFIX, vbTextCompare) = 0 Then
        CategoryFromButtonCaption = Trim$(Mid$(trimmedCaption, Len(BUTTON_CAPTION_HIDE_PREFIX) + 1))
    Else
        CategoryFromButtonCaption = trimmedCaption
    End If
End Function

Private Sub LayoutCategoryButtons(ByVal wsHome As Worksheet, ByRef categoryKeys() As String)
    Dim i As Long
    Dim buttonName As String
    Dim btn As Button
    Dim anchor As Range
    Dim topPos As Double

    If Not IsArrayInitialized(categoryKeys) Then Exit Sub

    Set anchor = wsHome.Range(BUTTON_ANCHOR_CELL)

    For i = LBound(categoryKeys) To UBound(categoryKeys)
        buttonName = BuildButtonName(categoryKeys(i))
        If ButtonExists(wsHome, buttonName) Then
            Set btn = wsHome.Buttons(buttonName)
            topPos = anchor.Top + (i - LBound(categoryKeys)) * (BUTTON_HEIGHT + BUTTON_VERTICAL_GAP)
            btn.Left = anchor.Left
            btn.Top = topPos
            btn.Width = BUTTON_WIDTH
            btn.Height = BUTTON_HEIGHT
        End If
    Next i
End Sub

Private Function BuildButtonName(ByVal categoryName As String) As String
    BuildButtonName = BUTTON_NAME_PREFIX & MakeNameSafe(categoryName)
End Function

Private Function CategoryFromButtonName(ByVal buttonName As String) As String
    Dim encoded As String
    Dim categories As Object
    Dim categoryKey As Variant
    Dim homeCategory As String

    If Left$(buttonName, Len(BUTTON_NAME_PREFIX)) <> BUTTON_NAME_PREFIX Then Exit Function
    encoded = Mid$(buttonName, Len(BUTTON_NAME_PREFIX) + 1)

    homeCategory = Trim$(CStr(ThisWorkbook.Worksheets(HOME_SHEET_NAME).Range(CATEGORY_CELL).Value))
    Set categories = CollectSheetCategories(homeCategory)

    For Each categoryKey In categories.Keys
        If StrComp(MakeNameSafe(CStr(categoryKey)), encoded, vbTextCompare) = 0 Then
            CategoryFromButtonName = CStr(categoryKey)
            Exit Function
        End If
    Next categoryKey
End Function

Private Function ButtonExists(ByVal ws As Worksheet, ByVal buttonName As String) As Boolean
    Dim btn As Button

    On Error Resume Next
    Set btn = ws.Buttons(buttonName)
    On Error GoTo 0

    ButtonExists = Not btn Is Nothing
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

    If Len(cleaned) = 0 Then cleaned = "Category"
    If Len(cleaned) > 20 Then cleaned = Left$(cleaned, 20)
    MakeNameSafe = cleaned
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
