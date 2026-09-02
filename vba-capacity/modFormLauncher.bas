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
