Attribute VB_Name = "modPartNumberSheets"
Option Explicit

Private Const HOME_SHEET_NAME As String = "Home"
Private Const TEMPLATE_SHEET_NAME As String = "Part Number Template"
Private Const BASE_PART_COLUMN As String = "C"
Private Const BASE_PART_START_ROW As Long = 2
Private Const HEADER_BASE_PART As String = "Base Part Number"
Private Const TEMPLATE_BASE_PART_CELL As String = "C2"

' For each base part number listed in Home column C, creates a worksheet
' named after that part (copied from Part Number Template) when one does
' not already exist, and writes the base part number into C2.
Public Sub CreateMissingPartNumberSheets()
    Dim wsHome As Worksheet
    Dim wsTemplate As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim basePart As String
    Dim createdCount As Long
    Dim skippedCount As Long

    On Error GoTo CleanUp
    OptimizeExcel True

    Set wsHome = ThisWorkbook.Worksheets(HOME_SHEET_NAME)
    Set wsTemplate = ThisWorkbook.Worksheets(TEMPLATE_SHEET_NAME)

    lastRow = wsHome.Cells(wsHome.Rows.Count, BASE_PART_COLUMN).End(xlUp).Row
    If lastRow < BASE_PART_START_ROW Then
        GoTo CleanUp
    End If

    For rowIndex = BASE_PART_START_ROW To lastRow
        basePart = Trim$(CStr(wsHome.Cells(rowIndex, BASE_PART_COLUMN).Value))

        If Len(basePart) = 0 Then
            ' Skip blank cells in the list.
        ElseIf StrComp(basePart, HEADER_BASE_PART, vbTextCompare) = 0 Then
            ' Skip the column header if present.
        ElseIf SheetExists(basePart) Then
            skippedCount = skippedCount + 1
        Else
            CreatePartNumberSheetFromTemplate wsTemplate, basePart
            createdCount = createdCount + 1
        End If
    Next rowIndex

CleanUp:
    OptimizeExcel False
End Sub

Private Sub CreatePartNumberSheetFromTemplate(ByVal wsTemplate As Worksheet, ByVal basePart As String)
    Dim wsNew As Worksheet
    Dim safeName As String

    safeName = SanitizeSheetName(basePart)

    wsTemplate.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set wsNew = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    wsNew.Name = safeName
    wsNew.Range(TEMPLATE_BASE_PART_CELL).Value = basePart
End Sub

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
