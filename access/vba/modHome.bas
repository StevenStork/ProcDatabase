Attribute VB_Name = "modHome"
Option Compare Database
Option Explicit

' Home page controller — search/filter/jump without scrolling the full catalog.
' Wired from frmHome control events (=HomeForm_Load(), =HomeApplyFilter(), etc.).

Public Function HomeForm_Load() As Boolean
    On Error GoTo Fail
    If Not CurrentProject.AllForms(FRM_HOME).IsLoaded Then
        HomeForm_Load = False
        Exit Function
    End If
    On Error Resume Next
    DoCmd.Maximize
    On Error GoTo Fail
    With Forms(FRM_HOME)
        .AllowEdits = True
        On Error Resume Next
        !chkActiveOnly = True
        !txtSearch = Null
        !cboFilterFfa = Null
        !cboJumpPart = Null
        On Error GoTo Fail
    End With
    HomeRefreshFilterFfaCombo
    HomeApplyFilter
    HomeForm_Load = True
    Exit Function
Fail:
    HomeForm_Load = False
End Function

Public Function HomeApplyFilter() As Boolean
    Dim frm As Form
    Dim listForm As Form
    Dim criteria As String
    Dim searchText As String
    Dim ffaValue As String
    Dim activeOnly As Boolean

    On Error GoTo Fail
    If Not HomeIsReady(frm, listForm) Then
        HomeApplyFilter = False
        Exit Function
    End If

    searchText = Trim$(CoerceText(frm!txtSearch))
    ffaValue = Trim$(CoerceText(frm!cboFilterFfa))
    activeOnly = CBool(Nz(frm!chkActiveOnly, True))

    If activeOnly Then
        criteria = AppendCriteria(criteria, "Active <> 0")
    End If
    If Len(ffaValue) > 0 Then
        criteria = AppendCriteria(criteria, "HomeFFA = " & SqlText(ffaValue))
    End If
    If Len(searchText) > 0 Then
        criteria = AppendCriteria(criteria, _
            "(BasePart Like " & SqlText("*" & searchText & "*") & _
            " Or Notes Like " & SqlText("*" & searchText & "*") & _
            " Or HomeFFA Like " & SqlText("*" & searchText & "*") & _
            " Or Factories Like " & SqlText("*" & searchText & "*") & ")")
    End If

    If Len(criteria) = 0 Then
        listForm.FilterOn = False
        listForm.Filter = vbNullString
    Else
        listForm.Filter = criteria
        listForm.FilterOn = True
    End If

    HomeRefreshJumpCombo
    HomeUpdateStatus
    HomeApplyFilter = True
    Exit Function
Fail:
    HomeApplyFilter = False
End Function

Public Function HomeClearFilter() As Boolean
    On Error GoTo Fail
    If Not CurrentProject.AllForms(FRM_HOME).IsLoaded Then
        HomeClearFilter = False
        Exit Function
    End If
    With Forms(FRM_HOME)
        !txtSearch = Null
        !cboFilterFfa = Null
        !chkActiveOnly = True
        !cboJumpPart = Null
    End With
    HomeApplyFilter
    HomeClearFilter = True
    Exit Function
Fail:
    HomeClearFilter = False
End Function

Public Function HomeSearchChanged() As Boolean
    On Error Resume Next
    If CurrentProject.AllForms(FRM_HOME).IsLoaded Then
        ' OnChange fires before Value updates — use Text for live filtering.
        Forms(FRM_HOME)!txtSearch.Value = Forms(FRM_HOME)!txtSearch.Text
    End If
    On Error GoTo 0
    HomeSearchChanged = HomeApplyFilter()
End Function

Public Function HomeJumpToPart() As Boolean
    Dim frm As Form
    Dim listForm As Form
    Dim basePart As String
    Dim rs As DAO.Recordset

    On Error GoTo Fail
    If Not HomeIsReady(frm, listForm) Then
        HomeJumpToPart = False
        Exit Function
    End If

    basePart = Trim$(CoerceText(frm!cboJumpPart))
    If Len(basePart) = 0 Then
        HomeJumpToPart = False
        Exit Function
    End If

    Set rs = listForm.RecordsetClone
    rs.FindFirst "BasePart = " & SqlText(basePart)
    If rs.NoMatch Then
        ' Try partial match within the filtered list.
        rs.FindFirst "BasePart Like " & SqlText(basePart & "*")
    End If
    If Not rs.NoMatch Then
        listForm.Bookmark = rs.Bookmark
        frm!lblStatus.Caption = "Selected " & CoerceText(listForm!BasePart) & "."
        HomeJumpToPart = True
    Else
        frm!lblStatus.Caption = "No match for '" & basePart & "' in the current filter."
        HomeJumpToPart = False
    End If
    rs.Close
    Exit Function
Fail:
    HomeJumpToPart = False
End Function

Public Function HomeOpenSelectedPart() As Boolean
    Dim frm As Form
    Dim listForm As Form
    Dim basePart As String

    On Error GoTo Fail
    If Not HomeIsReady(frm, listForm) Then
        HomeOpenSelectedPart = False
        Exit Function
    End If

    basePart = Trim$(CoerceText(frm!cboJumpPart))
    If Len(basePart) = 0 Then
        basePart = CoerceText(listForm!txtBasePart)
    End If
    If Len(basePart) = 0 Then
        basePart = CoerceText(listForm!BasePart)
    End If
    If Len(basePart) = 0 Then
        MsgBox "Select or jump to a part first.", vbExclamation, "Open Part"
        HomeOpenSelectedPart = False
        Exit Function
    End If

    OpenFormSized FRM_PART, "BasePart = " & SqlText(basePart)
    HomeOpenSelectedPart = True
    Exit Function
Fail:
    HomeOpenSelectedPart = False
End Function

Public Function HomeRefreshAfterDataChange() As Boolean
    On Error GoTo Fail
    If Not CurrentProject.AllForms(FRM_HOME).IsLoaded Then
        HomeRefreshAfterDataChange = False
        Exit Function
    End If
    Forms(FRM_HOME)!subParts.Requery
    HomeRefreshFilterFfaCombo
    HomeApplyFilter
    HomeRefreshAfterDataChange = True
    Exit Function
Fail:
    HomeRefreshAfterDataChange = False
End Function

Private Function HomeIsReady(ByRef frm As Form, ByRef listForm As Form) As Boolean
    On Error GoTo Fail
    If Not CurrentProject.AllForms(FRM_HOME).IsLoaded Then
        HomeIsReady = False
        Exit Function
    End If
    Set frm = Forms(FRM_HOME)
    Set listForm = frm!subParts.Form
    HomeIsReady = True
    Exit Function
Fail:
    HomeIsReady = False
End Function

Private Sub HomeRefreshJumpCombo()
    Dim frm As Form
    Dim listForm As Form
    Dim sql As String
    Dim whereClause As String

    On Error GoTo CleanUp
    If Not HomeIsReady(frm, listForm) Then Exit Sub

    whereClause = vbNullString
    If listForm.FilterOn And Len(Nz(listForm.Filter, vbNullString)) > 0 Then
        whereClause = " WHERE " & listForm.Filter
    End If
    sql = "SELECT BasePart FROM [" & QRY_HOME & "]" & whereClause & " ORDER BY BasePart"
    frm!cboJumpPart.RowSource = sql
    frm!cboJumpPart.Requery
CleanUp:
End Sub

Private Sub HomeRefreshFilterFfaCombo()
    On Error Resume Next
    If Not CurrentProject.AllForms(FRM_HOME).IsLoaded Then Exit Sub
    With Forms(FRM_HOME)!cboFilterFfa
        .RowSource = "SELECT FFA FROM [" & TBL_FFA & "] ORDER BY FFA"
        .Requery
    End With
End Sub

Private Sub HomeUpdateStatus()
    Dim frm As Form
    Dim listForm As Form
    Dim rs As DAO.Recordset
    Dim n As Long

    On Error GoTo CleanUp
    If Not HomeIsReady(frm, listForm) Then Exit Sub
    Set rs = listForm.RecordsetClone
    If rs.BOF And rs.EOF Then
        n = 0
    Else
        rs.MoveLast
        n = rs.RecordCount
    End If
    rs.Close
    frm!lblStatus.Caption = CStr(n) & " part(s) shown. Search, filter by FFA, or jump to a part — Active only is on by default."
CleanUp:
End Sub

Private Function AppendCriteria(ByVal existing As String, ByVal piece As String) As String
    If Len(existing) = 0 Then
        AppendCriteria = piece
    Else
        AppendCriteria = existing & " And " & piece
    End If
End Function
