Attribute VB_Name = "modFormLauncher"
Option Explicit

'==============================================================================
' Public entry points for Admin sheet buttons and macros.
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
    frmPartEditor.Show
End Sub

Public Sub ShowPartOperationsAdmin()
    frmPartOperationsAdmin.Show
End Sub

Public Sub EditSelectedPartFromSheet()
    Dim ws As Worksheet
    Dim selectedCode As String
    Dim tbl As ListObject

    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets(PART_EDITOR_SHEET_NAME)
    If TypeName(Selection) <> "Range" Then GoTo Fail
    If Selection.ListObject Is Nothing Then GoTo Fail

    Set tbl = Selection.ListObject
    If tbl.Name <> BASE_PARTS_TABLE_NAME Then GoTo Fail
    If tbl.DataBodyRange Is Nothing Then GoTo Fail
    If Intersect(Selection, tbl.DataBodyRange) Is Nothing Then GoTo Fail

    selectedCode = NormalizeCode(Selection.Cells(1, 1).Value2)
    If Len(selectedCode) = 0 Then GoTo Fail

    frmPartEditor.PreselectBasePart selectedCode
    frmPartEditor.Show
    Exit Sub

Fail:
    ShowPartEditor
End Sub
