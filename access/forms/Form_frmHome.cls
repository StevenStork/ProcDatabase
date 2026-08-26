Option Compare Database
Option Explicit

' Form_frmHome class module (source of truth for event handlers).
' BuildUi attempts to install this into Form_frmHome via VBIDE when trusted.
' Control OnLoad/=HomeForm_Load() properties also call modHome directly.

Private Sub Form_Load()
    HomeForm_Load
End Sub

Private Sub txtSearch_Change()
    HomeSearchChanged
End Sub

Private Sub txtSearch_Exit(Cancel As Integer)
    HomeApplyFilter
End Sub

Private Sub chkActiveOnly_AfterUpdate()
    HomeApplyFilter
End Sub

Private Sub cboFilterFfa_AfterUpdate()
    HomeApplyFilter
End Sub

Private Sub cboJumpPart_AfterUpdate()
    HomeJumpToPart
End Sub

Private Sub btnJump_Click()
    HomeJumpToPart
End Sub

Private Sub btnOpenPart_Click()
    HomeOpenSelectedPart
End Sub

Private Sub btnClearFilter_Click()
    HomeClearFilter
End Sub
