Attribute VB_Name = "modPartNumberSheets"
Option Explicit

Private Const HOME_SHEET_NAME As String = "Home"
Private Const TEMPLATE_SHEET_NAME As String = "Part Number Template"
Private Const REFERENCES_SHEET_NAME As String = "References"
Private Const ASSEMBLY_STANDARDS_SHEET_NAME As String = "Assembly Standards"
Private Const ASSY_STANDARDS_TABLE_NAME As String = "AssyStndTbl"
Private Const COL_ASSEMBLY_NO As String = "ASSEMBLY NO"

Private Const BASE_PART_COLUMN As String = "C"
Private Const BASE_PART_START_ROW As Long = 3
Private Const HEADER_BASE_PART As String = "Base Part Number"
Private Const TEMPLATE_BASE_PART_CELL As String = "C2"
Private Const PART_LABEL_CELL As String = "A1"
Private Const PART_LABEL_VALUE As String = "Part"

Private Const FFA_VALUE_COLUMN As String = "C"
Private Const FFA_CHECKBOX_COLUMN As String = "D"
Private Const DASH_VALUE_COLUMN As String = "E"
Private Const DASH_CHECKBOX_COLUMN As String = "F"
Private Const LIST_START_ROW As Long = 9

Private Const FFA_CHECKBOX_PREFIX As String = "chkFFA_"
Private Const DASH_CHECKBOX_PREFIX As String = "chkDash_"

' For each base part number listed in Home column C, creates a worksheet
' named after that part (copied from Part Number Template) when one does
' not already exist, then populates FFA and dash-condition lists.
Public Sub CreateMissingPartNumberSheets()
    Dim wsHome As Worksheet
    Dim wsTemplate As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim basePart As String

    On Error GoTo CleanUp
    OptimizeExcel True

    Set wsHome = ThisWorkbook.Worksheets(HOME_SHEET_NAME)
    Set wsTemplate = ThisWorkbook.Worksheets(TEMPLATE_SHEET_NAME)

    lastRow = LastUsedRowInColumn(wsHome, BASE_PART_COLUMN)
    If lastRow < BASE_PART_START_ROW Then
        GoTo CleanUp
    End If

    For rowIndex = BASE_PART_START_ROW To lastRow
        basePart = Trim$(CStr(wsHome.Cells(rowIndex, BASE_PART_COLUMN).Value))

        If Len(basePart) = 0 Then
            ' Skip blank cells in the list.
        ElseIf StrComp(basePart, HEADER_BASE_PART, vbTextCompare) = 0 Then
            ' Skip the column header if present.
        ElseIf Not SheetExists(basePart) Then
            CreatePartNumberSheetFromTemplateSafe wsTemplate, basePart
        End If
    Next rowIndex

CleanUp:
    OptimizeExcel False
End Sub

' Creates one part sheet; errors for a single part do not stop the remaining parts.
Private Sub CreatePartNumberSheetFromTemplateSafe(ByVal wsTemplate As Worksheet, ByVal basePart As String)
    On Error Resume Next
    CreatePartNumberSheetFromTemplate wsTemplate, basePart
    On Error GoTo 0
End Sub

Private Sub CreatePartNumberSheetFromTemplate(ByVal wsTemplate As Worksheet, ByVal basePart As String)
    Dim wsNew As Worksheet
    Dim safeName As String
    Dim ffaValues() As String
    Dim dashConditions() As String

    safeName = SanitizeSheetName(basePart)

    wsTemplate.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set wsNew = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    wsNew.Name = safeName

    wsNew.Range(PART_LABEL_CELL).Value = PART_LABEL_VALUE
    wsNew.Range(TEMPLATE_BASE_PART_CELL).Value = basePart

    ffaValues = GetReferenceFFAValues()
    dashConditions = GetDashConditionsForBasePart(basePart)

    PopulateValueCheckboxList _
        wsNew, _
        ffaValues, _
        FFA_VALUE_COLUMN, _
        FFA_CHECKBOX_COLUMN, _
        LIST_START_ROW, _
        FFA_CHECKBOX_PREFIX

    PopulateValueCheckboxList _
        wsNew, _
        dashConditions, _
        DASH_VALUE_COLUMN, _
        DASH_CHECKBOX_COLUMN, _
        LIST_START_ROW, _
        DASH_CHECKBOX_PREFIX
End Sub

Private Sub PopulateValueCheckboxList( _
    ByVal ws As Worksheet, _
    ByRef values() As String, _
    ByVal valueColumn As String, _
    ByVal checkboxColumn As String, _
    ByVal startRow As Long, _
    ByVal checkboxNamePrefix As String)

    Dim i As Long
    Dim outputRow As Long
    Dim valueCell As Range
    Dim checkboxCell As Range
    Dim listRange As Range

    If Not IsArrayInitialized(values) Then Exit Sub

    outputRow = startRow
    For i = LBound(values) To UBound(values)
        Set valueCell = ws.Cells(outputRow, valueColumn)
        Set checkboxCell = ws.Cells(outputRow, checkboxColumn)

        valueCell.Value = values(i)
        valueCell.HorizontalAlignment = xlLeft
        valueCell.VerticalAlignment = xlCenter

        checkboxCell.HorizontalAlignment = xlCenter
        checkboxCell.VerticalAlignment = xlCenter

        AddActiveXCheckBox ws, checkboxCell, BuildControlName(checkboxNamePrefix, ws.Name, outputRow)

        If listRange Is Nothing Then
            Set listRange = ws.Range(valueCell, checkboxCell)
        Else
            Set listRange = Union(listRange, ws.Range(valueCell, checkboxCell))
        End If

        outputRow = outputRow + 1
    Next i

    If Not listRange Is Nothing Then
        ApplyListBorders listRange
    End If
End Sub

Private Sub AddActiveXCheckBox(ByVal ws As Worksheet, ByVal targetCell As Range, ByVal checkboxName As String)
    Dim ole As OLEObject
    Dim boxSize As Double
    Dim leftPos As Double
    Dim topPos As Double

    boxSize = Application.WorksheetFunction.Min(targetCell.Width, targetCell.Height) - 4
    If boxSize < 10 Then boxSize = 12

    leftPos = targetCell.Left + (targetCell.Width - boxSize) / 2
    topPos = targetCell.Top + (targetCell.Height - boxSize) / 2

    Set ole = ws.OLEObjects.Add( _
        ClassType:="Forms.CheckBox.1", _
        Link:=False, _
        DisplayAsIcon:=False, _
        Left:=leftPos, _
        Top:=topPos, _
        Width:=boxSize, _
        Height:=boxSize)

    ole.Name = checkboxName
    ole.Placement = xlMoveAndSize

    With ole.Object
        .Caption = vbNullString
        .Value = False
        .SpecialEffect = 0
        .BackStyle = 0
    End With
End Sub

Private Sub ApplyListBorders(ByVal listRange As Range)
    With listRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    listRange.BorderAround LineStyle:=xlContinuous, Weight:=xlThick, ColorIndex:=xlAutomatic
End Sub

Private Function GetReferenceFFAValues() As String()
    Dim wsReferences As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim ffaValue As String
    Dim uniqueValues As Object
    Dim keys() As String

    Set wsReferences = ThisWorkbook.Worksheets(REFERENCES_SHEET_NAME)
    Set uniqueValues = CreateObject("Scripting.Dictionary")
    uniqueValues.CompareMode = vbTextCompare

    lastRow = LastUsedRowInColumn(wsReferences, "B")
    For rowIndex = 2 To lastRow
        ffaValue = Trim$(CStr(wsReferences.Cells(rowIndex, "B").Value))
        If Len(ffaValue) > 0 Then
            If Not uniqueValues.Exists(ffaValue) Then
                uniqueValues.Add ffaValue, ffaValue
            End If
        End If
    Next rowIndex

    keys = DictionaryKeysToSortedArray(uniqueValues)
    GetReferenceFFAValues = keys
End Function

Private Function GetDashConditionsForBasePart(ByVal basePart As String) As String()
    Dim wsStandards As Worksheet
    Dim tbl As ListObject
    Dim colAssembly As ListColumn
    Dim rowIndex As Long
    Dim assemblyVal As String
    Dim rowBasePart As String
    Dim dashCondition As String
    Dim uniqueValues As Object
    Dim keys() As String

    Set wsStandards = ThisWorkbook.Worksheets(ASSEMBLY_STANDARDS_SHEET_NAME)
    Set tbl = wsStandards.ListObjects(ASSY_STANDARDS_TABLE_NAME)
    Set colAssembly = tbl.ListColumns(COL_ASSEMBLY_NO)
    Set uniqueValues = CreateObject("Scripting.Dictionary")
    uniqueValues.CompareMode = vbTextCompare

    If tbl.DataBodyRange Is Nothing Then
        GetDashConditionsForBasePart = EmptyStringArray()
        Exit Function
    End If

    For rowIndex = 1 To tbl.ListRows.Count
        assemblyVal = Trim$(CStr(colAssembly.DataBodyRange.Cells(rowIndex, 1).Value))
        If Len(assemblyVal) > 0 Then
            SplitAssemblyNo assemblyVal, rowBasePart, dashCondition
            If StrComp(rowBasePart, basePart, vbTextCompare) = 0 And Len(dashCondition) > 0 Then
                If Not uniqueValues.Exists(dashCondition) Then
                    uniqueValues.Add dashCondition, dashCondition
                End If
            End If
        End If
    Next rowIndex

    keys = DictionaryKeysToSortedArray(uniqueValues)
    GetDashConditionsForBasePart = keys
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

Private Sub SplitAssemblyNo(ByVal assemblyNo As String, ByRef basePart As String, ByRef dashCondition As String)
    Dim dashPos As Long

    dashPos = InStr(1, assemblyNo, "-", vbBinaryCompare)

    If dashPos > 0 Then
        basePart = Trim$(Left$(assemblyNo, dashPos - 1))
        dashCondition = Trim$(Mid$(assemblyNo, dashPos + 1))
    Else
        basePart = assemblyNo
        dashCondition = vbNullString
    End If
End Sub

Private Sub SortStringArray(ByRef keys() As String)
    Dim i As Long
    Dim j As Long
    Dim tempKey As String

    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If CompareListValues(keys(i), keys(j)) > 0 Then
                tempKey = keys(i)
                keys(i) = keys(j)
                keys(j) = tempKey
            End If
        Next j
    Next i
End Sub

Private Function CompareListValues(ByVal leftValue As String, ByVal rightValue As String) As Long
    If IsNumeric(leftValue) And IsNumeric(rightValue) Then
        CompareListValues = Sgn(CDbl(leftValue) - CDbl(rightValue))
    Else
        CompareListValues = StrComp(leftValue, rightValue, vbTextCompare)
    End If
End Function

Private Function BuildControlName(ByVal prefix As String, ByVal sheetName As String, ByVal rowIndex As Long) As String
    ' ActiveX control names must be unique across the whole workbook.
    BuildControlName = prefix & MakeControlSafe(sheetName) & "_" & CStr(rowIndex)
End Function

Private Function MakeControlSafe(ByVal textValue As String) As String
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

    If Len(cleaned) = 0 Then cleaned = "Part"
    MakeControlSafe = cleaned
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

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    SheetExists = Not ws Is Nothing
End Function

Private Function SanitizeSheetName(ByVal proposedName As String) As String
    Dim cleaned As String

    cleaned = proposedName
    cleaned = Replace$(cleaned, "\", "_")
    cleaned = Replace$(cleaned, "/", "_")
    cleaned = Replace$(cleaned, "?", "_")
    cleaned = Replace$(cleaned, "*", "_")
    cleaned = Replace$(cleaned, "[", "_")
    cleaned = Replace$(cleaned, "]", "_")
    cleaned = Replace$(cleaned, ":", "_")

    If Len(cleaned) > 31 Then
        cleaned = Left$(cleaned, 31)
    End If

    If Len(cleaned) = 0 Then
        Err.Raise vbObjectError + 600, "SanitizeSheetName", "Base part number produced an empty sheet name."
    End If

    SanitizeSheetName = cleaned
End Function
