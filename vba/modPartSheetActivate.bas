Attribute VB_Name = "modPartSheetActivate"
Option Explicit

Private Const PART_LABEL_CELL As String = "A1"
Private Const PART_LABEL_VALUE As String = "Part"
Private Const BASE_PART_CELL As String = "C2"

Private Const REFERENCES_SHEET_NAME As String = "References"
Private Const ASSEMBLY_STANDARDS_SHEET_NAME As String = "Assembly Standards"
Private Const ASSY_STANDARDS_TABLE_NAME As String = "AssyStndTbl"
Private Const COL_ASSEMBLY_NO As String = "ASSEMBLY NO"

Private Const LIST_START_ROW As Long = 9
Private Const FFA_VALUE_COLUMN As String = "C"
Private Const FFA_CHECKBOX_COLUMN As String = "D"
Private Const DASH_VALUE_COLUMN As String = "E"
Private Const DASH_CHECKBOX_COLUMN As String = "F"
Private Const PRODUCT_LINE_VALUE_COLUMN As String = "G"
Private Const PRODUCT_LINE_CHECKBOX_COLUMN As String = "H"

Private Const FFA_CHECKBOX_PREFIX As String = "chkFFA_"
Private Const DASH_CHECKBOX_PREFIX As String = "chkDash_"
Private Const PRODUCT_LINE_CHECKBOX_PREFIX As String = "chkProd_"

' Call from ThisWorkbook.Workbook_SheetActivate:
'   HandlePartSheetActivate Sh
Public Sub HandlePartSheetActivate(ByVal Sh As Object)
    Dim ws As Worksheet

    On Error GoTo CleanUp

    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    Set ws = Sh

    If StrComp(Trim$(CStr(ws.Range(PART_LABEL_CELL).Value)), PART_LABEL_VALUE, vbTextCompare) <> 0 Then
        Exit Sub
    End If

    OptimizeExcel True
    RefreshPartSheetLists ws

CleanUp:
    OptimizeExcel False
End Sub

Private Sub RefreshPartSheetLists(ByVal ws As Worksheet)
    Dim basePart As String
    Dim ffaValues() As String
    Dim dashConditions() As String
    Dim productLines() As String

    basePart = Trim$(CStr(ws.Range(BASE_PART_CELL).Value))
    If Len(basePart) = 0 Then Exit Sub

    ffaValues = GetReferenceColumnValues("B")
    dashConditions = GetDashConditionsForBasePart(basePart)
    productLines = GetReferenceColumnValues("D")

    SyncValueCheckboxList ws, FFA_VALUE_COLUMN, FFA_CHECKBOX_COLUMN, ffaValues, FFA_CHECKBOX_PREFIX
    SyncValueCheckboxList ws, DASH_VALUE_COLUMN, DASH_CHECKBOX_COLUMN, dashConditions, DASH_CHECKBOX_PREFIX
    SyncValueCheckboxList ws, PRODUCT_LINE_VALUE_COLUMN, PRODUCT_LINE_CHECKBOX_COLUMN, productLines, PRODUCT_LINE_CHECKBOX_PREFIX
End Sub

' Writes values/checkboxes/borders only when the list or checkbox set needs updating.
Private Sub SyncValueCheckboxList( _
    ByVal ws As Worksheet, _
    ByVal valueColumn As String, _
    ByVal checkboxColumn As String, _
    ByRef sourceValues() As String, _
    ByVal checkboxNamePrefix As String)

    Dim currentValues() As String
    Dim sourceCount As Long
    Dim expectedPrefix As String

    currentValues = ReadColumnList(ws, valueColumn)
    sourceCount = ArrayCount(sourceValues)
    expectedPrefix = ControlNamePrefix(checkboxNamePrefix, ws.Name)

    If StringArraysEqual(currentValues, sourceValues) Then
        If CountCheckBoxesWithPrefix(ws, expectedPrefix) = sourceCount Then
            Exit Sub
        End If
    End If

    ClearValueCheckboxList ws, valueColumn, checkboxColumn, expectedPrefix

    If sourceCount = 0 Then Exit Sub

    PopulateValueCheckboxList ws, valueColumn, checkboxColumn, sourceValues, checkboxNamePrefix
End Sub

Private Sub ClearValueCheckboxList( _
    ByVal ws As Worksheet, _
    ByVal valueColumn As String, _
    ByVal checkboxColumn As String, _
    ByVal expectedPrefix As String)

    Dim lastRow As Long
    Dim clearRange As Range

    DeleteCheckBoxesWithPrefix ws, expectedPrefix

    lastRow = Application.WorksheetFunction.Max( _
        LastUsedRowInColumn(ws, valueColumn), _
        LastUsedRowInColumn(ws, checkboxColumn), _
        LIST_START_ROW)

    If lastRow >= LIST_START_ROW Then
        Set clearRange = ws.Range( _
            ws.Cells(LIST_START_ROW, valueColumn), _
            ws.Cells(lastRow, checkboxColumn))
        clearRange.ClearContents
        clearRange.Borders.LineStyle = xlNone
    End If
End Sub

Private Sub PopulateValueCheckboxList( _
    ByVal ws As Worksheet, _
    ByVal valueColumn As String, _
    ByVal checkboxColumn As String, _
    ByRef values() As String, _
    ByVal checkboxNamePrefix As String)

    Dim i As Long
    Dim outputRow As Long
    Dim valueCell As Range
    Dim checkboxCell As Range
    Dim listRange As Range

    outputRow = LIST_START_ROW
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

    listRange.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, ColorIndex:=xlAutomatic
End Sub

Private Sub DeleteCheckBoxesWithPrefix(ByVal ws As Worksheet, ByVal expectedPrefix As String)
    Dim oleIndex As Long
    Dim ole As OLEObject

    For oleIndex = ws.OLEObjects.Count To 1 Step -1
        Set ole = ws.OLEObjects(oleIndex)
        If Left$(ole.Name, Len(expectedPrefix)) = expectedPrefix Then
            ole.Delete
        End If
    Next oleIndex
End Sub

Private Function CountCheckBoxesWithPrefix(ByVal ws As Worksheet, ByVal expectedPrefix As String) As Long
    Dim ole As OLEObject
    Dim matchCount As Long

    For Each ole In ws.OLEObjects
        If Left$(ole.Name, Len(expectedPrefix)) = expectedPrefix Then
            matchCount = matchCount + 1
        End If
    Next ole

    CountCheckBoxesWithPrefix = matchCount
End Function

Private Function ReadColumnList(ByVal ws As Worksheet, ByVal columnLetter As String) As String()
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim values() As String
    Dim valueCount As Long
    Dim cellValue As String

    lastRow = LastUsedRowInColumn(ws, columnLetter)
    If lastRow < LIST_START_ROW Then
        ReadColumnList = EmptyStringArray()
        Exit Function
    End If

    valueCount = 0
    ReDim values(0 To 0)

    For rowIndex = LIST_START_ROW To lastRow
        cellValue = Trim$(CStr(ws.Cells(rowIndex, columnLetter).Value))
        If Len(cellValue) > 0 Then
            ReDim Preserve values(0 To valueCount)
            values(valueCount) = cellValue
            valueCount = valueCount + 1
        Else
            Exit For
        End If
    Next rowIndex

    If valueCount = 0 Then
        ReadColumnList = EmptyStringArray()
    Else
        ReadColumnList = values
    End If
End Function

Private Function GetReferenceColumnValues(ByVal columnLetter As String) As String()
    Dim wsReferences As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim cellValue As String
    Dim uniqueValues As Object

    Set wsReferences = ThisWorkbook.Worksheets(REFERENCES_SHEET_NAME)
    Set uniqueValues = CreateObject("Scripting.Dictionary")
    uniqueValues.CompareMode = vbTextCompare

    lastRow = LastUsedRowInColumn(wsReferences, columnLetter)
    For rowIndex = 2 To lastRow
        cellValue = Trim$(CStr(wsReferences.Cells(rowIndex, columnLetter).Value))
        If Len(cellValue) > 0 Then
            If Not uniqueValues.Exists(cellValue) Then
                uniqueValues.Add cellValue, cellValue
            End If
        End If
    Next rowIndex

    GetReferenceColumnValues = DictionaryKeysToSortedArray(uniqueValues)
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

    GetDashConditionsForBasePart = DictionaryKeysToSortedArray(uniqueValues)
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

Private Function StringArraysEqual(ByRef leftValues() As String, ByRef rightValues() As String) As Boolean
    Dim i As Long
    Dim leftCount As Long
    Dim rightCount As Long

    leftCount = ArrayCount(leftValues)
    rightCount = ArrayCount(rightValues)

    If leftCount <> rightCount Then Exit Function

    For i = 0 To leftCount - 1
        If StrComp(leftValues(i), rightValues(i), vbTextCompare) <> 0 Then Exit Function
    Next i

    StringArraysEqual = True
End Function

Private Function ArrayCount(ByRef values() As String) As Long
    If Not IsArrayInitialized(values) Then
        ArrayCount = 0
    Else
        ArrayCount = UBound(values) - LBound(values) + 1
    End If
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
    BuildControlName = ControlNamePrefix(prefix, sheetName) & CStr(rowIndex)
End Function

Private Function ControlNamePrefix(ByVal prefix As String, ByVal sheetName As String) As String
    ControlNamePrefix = prefix & MakeControlSafe(sheetName) & "_"
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
