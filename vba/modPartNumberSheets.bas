Attribute VB_Name = "modPartNumberSheets"
Option Explicit

Private Const HOME_SHEET_NAME As String = "Home"
Private Const TEMPLATE_SHEET_NAME As String = "Part Number Template"
Private Const BASE_PART_COLUMN As String = "C"
Private Const ACTIVE_PART_COLUMN As String = "D"
Private Const BASE_PART_START_ROW As Long = 3
Private Const HEADER_BASE_PART As String = "Base Part Number"
Private Const TEMPLATE_BASE_PART_CELL As String = "C2"
Private Const PART_LABEL_CELL As String = "A1"
Private Const PART_LABEL_VALUE As String = "Part"

' For each active base part number listed in Home column C (starting at row 3),
' creates a worksheet named after that part (copied from Part Number Template)
' when one does not already exist, and writes the base part number into C2.
' A part is active when the same row in column D is True.
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
        ElseIf Not IsActivePart(wsHome.Cells(rowIndex, ACTIVE_PART_COLUMN).Value) Then
            ' Skip parts that are not marked active in column D.
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

    safeName = SanitizeSheetName(basePart)

    wsTemplate.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set wsNew = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    wsNew.Name = safeName

    wsNew.Range(PART_LABEL_CELL).Value = PART_LABEL_VALUE
    wsNew.Range(TEMPLATE_BASE_PART_CELL).Value = basePart
End Sub

Private Function IsActivePart(ByVal activeValue As Variant) As Boolean
    If IsError(activeValue) Then Exit Function
    If IsEmpty(activeValue) Or IsNull(activeValue) Then Exit Function

    If VarType(activeValue) = vbBoolean Then
        IsActivePart = CBool(activeValue)
        Exit Function
    End If

    If IsNumeric(activeValue) Then
        IsActivePart = (CDbl(activeValue) <> 0)
        Exit Function
    End If

    Select Case LCase$(Trim$(CStr(activeValue)))
        Case "true", "yes", "y", "1"
            IsActivePart = True
    End Select
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
