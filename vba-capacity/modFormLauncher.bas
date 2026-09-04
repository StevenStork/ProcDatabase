Attribute VB_Name = "modFormLauncher"
Option Explicit

'==============================================================================
' Public entry points unique to this module (form show helpers and aliases).
' LoadPartToEditor / SavePartFromEditor / ClearPartEditor / OpenPartEditorFromPartsIndex
' live only in modPartSheetEditor. Refresh* macros live only in modLinkedData.
' Do not duplicate those Public Sub names here — Excel raises Ambiguous name.
'==============================================================================

Public Sub ShowFactoryAdmin()
    frmFactoryAdmin.Show
End Sub

Public Sub ShowEquipmentAdmin()
    frmEquipmentAdmin.Show
End Sub

Public Sub ShowProcessTypeAdmin()
    frmProcessTypeAdmin.Show
End Sub

Public Sub ShowFactoryEquipmentAdmin()
    frmFactoryEquipmentAdmin.Show
End Sub

Public Sub ShowEquipmentProcessAdmin()
    frmEquipmentProcessAdmin.Show
End Sub

Public Sub ShowPartEditor()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PART_EDITOR_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        BootstrapCapacityTables
        Set ws = ThisWorkbook.Worksheets(PART_EDITOR_SHEET_NAME)
    End If

    ws.Activate
    ws.Cells(PE_INPUT_ROW, PE_VALUE_COL).Select
End Sub

Public Sub ShowPartOperationsAdmin()
    frmPartOperationsAdmin.Show
End Sub

' Legacy alias for Parts index row selection.
Public Sub EditSelectedPartFromSheet()
    OpenPartEditorFromPartsIndex
End Sub
