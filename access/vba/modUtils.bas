Attribute VB_Name = "modUtils"
Option Compare Database
Option Explicit

' Large form defaults in twips (~12.5" x ~7.6"). Used when UsableWidth is unavailable.
Private Const LAYOUT_DEFAULT_WIDTH As Long = 18000
Private Const LAYOUT_DEFAULT_HEIGHT As Long = 11000

Public Sub SplitAssemblyNo(ByVal assemblyNo As String, ByRef basePart As String, ByRef dashCondition As String)
    Dim dashPos As Long

    assemblyNo = Trim$(assemblyNo)
    dashPos = InStr(1, assemblyNo, "-", vbBinaryCompare)

    If dashPos > 0 Then
        basePart = Trim$(Left$(assemblyNo, dashPos - 1))
        dashCondition = Trim$(Mid$(assemblyNo, dashPos + 1))
    Else
        basePart = assemblyNo
        dashCondition = vbNullString
    End If
End Sub

Public Function GetBasePartNumber(ByVal assemblyNo As String) As String
    Dim basePart As String
    Dim dashCondition As String
    SplitAssemblyNo assemblyNo, basePart, dashCondition
    GetBasePartNumber = basePart
End Function

Public Function SqlText(ByVal value As String) As String
    SqlText = "'" & Replace(value, "'", "''") & "'"
End Function

Public Function SqlNullableText(ByVal value As String) As String
    value = Trim$(value)
    If Len(value) = 0 Then
        SqlNullableText = "Null"
    Else
        SqlNullableText = SqlText(value)
    End If
End Function

Public Function SqlNullableNumber(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        SqlNullableNumber = "Null"
    ElseIf Not IsNumeric(value) Then
        SqlNullableNumber = "Null"
    Else
        SqlNullableNumber = Str$(CDbl(value))
    End If
End Function

Public Function SqlNullableDate(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        SqlNullableDate = "Null"
    ElseIf Not IsDate(value) Then
        SqlNullableDate = "Null"
    Else
        SqlNullableDate = "#" & Format$(CDate(value), "yyyy-mm-dd") & "#"
    End If
End Function

Public Function SqlBool(ByVal value As Boolean) As String
    If value Then
        SqlBool = "True"
    Else
        SqlBool = "False"
    End If
End Function

Public Function TableExists(ByVal tableName As String) As Boolean
    Dim td As DAO.TableDef
    On Error Resume Next
    Set td = CurrentDb.TableDefs(tableName)
    TableExists = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function QueryExists(ByVal queryName As String) As Boolean
    Dim qdf As DAO.QueryDef
    On Error Resume Next
    Set qdf = CurrentDb.QueryDefs(queryName)
    QueryExists = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function ObjectExists(ByVal objectType As AcObjectType, ByVal objectName As String) As Boolean
    Dim i As Long

    Select Case objectType
        Case acForm
            For i = 0 To CurrentProject.AllForms.Count - 1
                If StrComp(CurrentProject.AllForms(i).Name, objectName, vbTextCompare) = 0 Then
                    ObjectExists = True
                    Exit Function
                End If
            Next i
        Case acReport
            For i = 0 To CurrentProject.AllReports.Count - 1
                If StrComp(CurrentProject.AllReports(i).Name, objectName, vbTextCompare) = 0 Then
                    ObjectExists = True
                    Exit Function
                End If
            Next i
        Case acMacro
            For i = 0 To CurrentProject.AllMacros.Count - 1
                If StrComp(CurrentProject.AllMacros(i).Name, objectName, vbTextCompare) = 0 Then
                    ObjectExists = True
                    Exit Function
                End If
            Next i
        Case Else
            ObjectExists = False
    End Select
End Function

Public Sub DeleteTableIfExists(ByVal tableName As String)
    If TableExists(tableName) Then
        CurrentDb.Execute "DROP TABLE [" & tableName & "]", dbFailOnError
    End If
End Sub

Public Sub ReplaceQuery(ByVal queryName As String, ByVal sql As String)
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef

    Set db = CurrentDb
    If QueryExists(queryName) Then
        db.QueryDefs.Delete queryName
    End If
    Set qdf = db.CreateQueryDef(queryName, sql)
End Sub

Public Sub SetMeta(ByVal metaKey As String, ByVal metaValue As String)
    Dim db As DAO.Database
    Dim keyColumn As String
    Dim valueColumn As String
    Set db = CurrentDb
    keyColumn = MetaKeyColumnName()
    valueColumn = MetaValueColumnName()
    db.Execute "DELETE FROM [" & TBL_META & "] WHERE [" & keyColumn & "] = " & SqlText(metaKey), dbFailOnError
    db.Execute "INSERT INTO [" & TBL_META & "] ([" & keyColumn & "], [" & valueColumn & "]) VALUES (" & _
        SqlText(metaKey) & ", " & SqlText(metaValue) & ")", dbFailOnError
End Sub

Public Function GetMeta(ByVal metaKey As String) As String
    GetMeta = Nz(DLookup("[" & MetaValueColumnName() & "]", TBL_META, _
        "[" & MetaKeyColumnName() & "] = " & SqlText(metaKey)), vbNullString)
End Function

Private Function MetaKeyColumnName() As String
    If TableExists(TBL_META) And FieldExists(TBL_META, "MetaKey") Then
        MetaKeyColumnName = "MetaKey"
    Else
        MetaKeyColumnName = "Key"
    End If
End Function

Private Function MetaValueColumnName() As String
    If TableExists(TBL_META) And FieldExists(TBL_META, "MetaValue") Then
        MetaValueColumnName = "MetaValue"
    Else
        MetaValueColumnName = "Value"
    End If
End Function

Public Function ValuesMatch(ByVal leftValue As String, ByVal rightValue As String) As Boolean
    If IsNumeric(leftValue) And IsNumeric(rightValue) Then
        ValuesMatch = (CDbl(leftValue) = CDbl(rightValue))
    Else
        ValuesMatch = (StrComp(leftValue, rightValue, vbTextCompare) = 0)
    End If
End Function

Public Function TryGetNonZeroNumeric(ByVal rawValue As Variant, ByRef numericValue As Double) As Boolean
    If IsError(rawValue) Then Exit Function
    If IsEmpty(rawValue) Then Exit Function
    If IsNull(rawValue) Then Exit Function
    If Len(Trim$(CStr(rawValue))) = 0 Then Exit Function
    If Not IsNumeric(rawValue) Then Exit Function

    numericValue = CDbl(rawValue)
    If numericValue = 0 Then Exit Function

    TryGetNonZeroNumeric = True
End Function

Public Function CoerceText(ByVal rawValue As Variant) As String
    If IsError(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then
        CoerceText = vbNullString
    Else
        CoerceText = Trim$(CStr(rawValue))
    End If
End Function

Public Sub DeleteObjectIfExists(ByVal objectType As AcObjectType, ByVal objectName As String)
    On Error Resume Next
    DoCmd.Close objectType, objectName, acSaveNo
    DoCmd.DeleteObject objectType, objectName
    On Error GoTo 0
End Sub

' Workspace size in twips. Do NOT use Application.UsableWidth/Height —
' many Access builds lack those members (compile: method or data member not found).
' Prefer late-bound CallByName; otherwise LAYOUT_DEFAULT_* . Maximize fills the window.

Private Function LayoutLateBoundTwips(ByVal propName As String) As Long
    Dim raw As Variant
    On Error Resume Next
    raw = CallByName(Application, propName, VbGet)
    If Err.Number = 0 And IsNumeric(raw) Then
        If CLng(raw) > 1000 Then LayoutLateBoundTwips = CLng(raw)
    End If
    Err.Clear
    On Error GoTo 0
End Function

Public Function LayoutUsableWidth(Optional ByVal frac As Double = 0.96) As Long
    Dim raw As Long
    raw = LayoutLateBoundTwips("UsableWidth")
    If raw <= 0 Then raw = LAYOUT_DEFAULT_WIDTH
    LayoutUsableWidth = CLng(raw * frac)
    If LayoutUsableWidth < 14000 Then LayoutUsableWidth = 14000
End Function

Public Function LayoutUsableHeight(Optional ByVal frac As Double = 0.92) As Long
    Dim raw As Long
    raw = LayoutLateBoundTwips("UsableHeight")
    If raw <= 0 Then raw = LAYOUT_DEFAULT_HEIGHT
    LayoutUsableHeight = CLng(raw * frac)
    If LayoutUsableHeight < 9000 Then LayoutUsableHeight = 9000
End Function

' Size a form in design to ~full Access workspace (twips).
' expandDetail:=False for continuous/datasheet forms (keeps row height normal).
Public Sub ApplyLargeFormLayout(ByVal frm As Form, _
    Optional ByVal widthFrac As Double = 0.96, _
    Optional ByVal heightFrac As Double = 0.92, _
    Optional ByVal expandDetail As Boolean = True)

    Dim targetW As Long
    Dim targetH As Long

    On Error Resume Next
    targetW = LayoutUsableWidth(widthFrac)
    targetH = LayoutUsableHeight(heightFrac)

    frm.Width = targetW
    If expandDetail Then
        ' Single-form shells: grow the detail section to fill the window.
        frm.InsideHeight = targetH
        frm.Section(acDetail).Height = targetH
    End If
    ' List/datasheet forms keep their own row height; Maximize at open fills the screen.
    frm.AutoCenter = True
    frm.MinMaxButtons = True
    frm.CloseButton = True
    frm.BorderStyle = 2 ' sizable
    frm.PopUp = False
    frm.Modal = False
    On Error GoTo 0
End Sub

' Place a control so it fills the remaining client area (design-time).
Public Sub SizeFillControl(ByVal ctl As Control, _
    ByVal leftMargin As Long, _
    ByVal topMargin As Long, _
    Optional ByVal rightMargin As Long = 240, _
    Optional ByVal bottomMargin As Long = 240)

    Dim w As Long
    Dim h As Long

    w = LayoutUsableWidth() - leftMargin - rightMargin
    h = LayoutUsableHeight() - topMargin - bottomMargin
    If w < 6000 Then w = 6000
    If h < 2400 Then h = 2400
    ctl.Left = leftMargin
    ctl.Top = topMargin
    ctl.Width = w
    ctl.Height = h
End Sub

' Stretch with the form window when the user resizes (Access 2010+).
Public Sub AnchorStretch(ByVal ctl As Control, _
    Optional ByVal stretchHorizontal As Boolean = True, _
    Optional ByVal stretchVertical As Boolean = False)

    On Error Resume Next
    If stretchHorizontal Then ctl.HorizontalAnchor = 2 ' acHorizontalAnchorBoth
    If stretchVertical Then ctl.VerticalAnchor = 2     ' acVerticalAnchorBoth
    On Error GoTo 0
End Sub

Public Sub AnchorTopRight(ByVal ctl As Control)
    On Error Resume Next
    ctl.HorizontalAnchor = 1 ' acHorizontalAnchorRight
    ctl.VerticalAnchor = 0   ' acVerticalAnchorTop
    On Error GoTo 0
End Sub

Public Sub AnchorBottom(ByVal ctl As Control, Optional ByVal stretchHorizontal As Boolean = False)
    On Error Resume Next
    If stretchHorizontal Then
        ctl.HorizontalAnchor = 2 ' acHorizontalAnchorBoth
    Else
        ctl.HorizontalAnchor = 0 ' acHorizontalAnchorLeft
    End If
    ctl.VerticalAnchor = 1 ' acVerticalAnchorBottom
    On Error GoTo 0
End Sub

' Open a form and expand it to the workspace (maximize main windows).
Public Sub OpenFormSized(ByVal formName As String, _
    Optional ByVal whereCondition As String = vbNullString, _
    Optional ByVal maximizeWindow As Boolean = True)

    On Error GoTo Fail
    If Not ObjectExists(acForm, formName) Then
        Err.Raise vbObjectError + 2102, "OpenFormSized", _
            "Form '" & formName & "' does not exist. Close open forms and run BuildUi " & _
            "(or BootstrapProcDatabase) from the Immediate window."
    End If
    If Len(whereCondition) > 0 Then
        DoCmd.OpenForm formName, , , whereCondition
    Else
        DoCmd.OpenForm formName
    End If
    On Error Resume Next
    If maximizeWindow Then
        DoCmd.Maximize
    Else
        DoCmd.MoveSize 0, 0, LayoutUsableWidth(1#), LayoutUsableHeight(1#)
    End If
    On Error GoTo 0
    Exit Sub
Fail:
    Err.Raise Err.Number, "OpenFormSized", Err.Description
End Sub

' Close open forms/tables/queries so schema DDL / form rebuild can lock tables (avoids 3211).
Public Sub CloseProcDataForms()
    Dim safety As Long
    On Error Resume Next
    DoCmd.Close acForm, FRM_PART, acSaveNo
    DoCmd.Close acForm, FRM_HOME, acSaveNo
    DoCmd.Close acForm, FRM_REFERENCES, acSaveNo
    DoCmd.Close acForm, FRM_EXPORT, acSaveNo
    DoCmd.Close acForm, FRM_FFA, acSaveNo
    DoCmd.Close acForm, FRM_PRODUCT_LINE, acSaveNo
    DoCmd.Close acForm, FRM_EQUIPMENT, acSaveNo
    DoCmd.Close acForm, FRM_EQUIPMENT_FFA, acSaveNo
    DoCmd.Close acForm, FRM_EQUIPMENT_ENTRY, acSaveNo
    DoCmd.Close acForm, SFRM_HOME_LIST, acSaveNo
    DoCmd.Close acForm, SFRM_DASH, acSaveNo
    DoCmd.Close acForm, SFRM_PL, acSaveNo
    DoCmd.Close acForm, SFRM_OPS, acSaveNo
    DoCmd.Close acForm, LEGACY_SFRM_EQUIPMENT_FFA, acSaveNo
    ' Datasheet views also hold locks and are not in the Forms collection.
    DoCmd.Close acTable, TBL_PART, acSaveNo
    DoCmd.Close acTable, TBL_PART_DASH, acSaveNo
    DoCmd.Close acTable, TBL_PART_PL, acSaveNo
    DoCmd.Close acTable, TBL_OPERATION, acSaveNo
    DoCmd.Close acQuery, QRY_HOME, acSaveNo
    DoCmd.Close acQuery, QRY_OPERATIONS, acSaveNo
    DoCmd.Close acQuery, QRY_EXPORT, acSaveNo
    ' Also clear any leftover Design-view FormN windows from a prior CreateForm.
    safety = 0
    Do While Forms.Count > 0 And safety < 50
        DoCmd.Close acForm, Forms(0).Name, acSaveNo
        safety = safety + 1
    Loop
    DoEvents
    DBEngine.Idle dbRefreshCache
    On Error GoTo 0
End Sub

Public Sub SetDbProperty(ByVal propName As String, ByVal propType As Integer, ByVal propValue As Variant)
    Dim db As DAO.Database
    Dim prp As DAO.Property

    Set db = CurrentDb
    On Error Resume Next
    db.Properties(propName) = propValue
    If Err.Number <> 0 Then
        Err.Clear
        Set prp = db.CreateProperty(propName, propType, propValue)
        db.Properties.Append prp
    End If
    On Error GoTo 0
End Sub

Public Function FieldExists(ByVal tableName As String, ByVal fieldName As String) As Boolean
    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Dim fld As DAO.Field

    On Error GoTo Fail
    Set db = CurrentDb
    Set td = db.TableDefs(tableName)
    For Each fld In td.Fields
        If StrComp(fld.Name, fieldName, vbTextCompare) = 0 Then
            FieldExists = True
            Exit Function
        End If
    Next fld
    FieldExists = False
    Exit Function
Fail:
    FieldExists = False
End Function

Public Sub AddTextFieldIfMissing(ByVal tableName As String, ByVal fieldName As String, ByVal size As Long)
    AddTextColumnIfMissing tableName, fieldName, size
End Sub

Public Sub AddTextColumnIfMissing(ByVal tableName As String, ByVal fieldName As String, ByVal size As Long)
    If FieldExists(tableName, fieldName) Then Exit Sub
    On Error GoTo Fail
    CurrentDb.Execute "ALTER TABLE [" & tableName & "] ADD COLUMN [" & fieldName & "] TEXT(" & size & ")", dbFailOnError
    Exit Sub
Fail:
    ' 3380 = field already exists; 3211 = table locked by open form — skip and continue.
    If Err.Number = 3029 Or Err.Number = 3191 Or Err.Number = 3380 Or Err.Number = 3211 Then
        Err.Clear
    Else
        Err.Raise Err.Number, "AddTextColumnIfMissing", Err.Description
    End If
End Sub

Public Sub AddMemoColumnIfMissing(ByVal tableName As String, ByVal fieldName As String)
    If FieldExists(tableName, fieldName) Then Exit Sub
    On Error GoTo Fail
    CurrentDb.Execute "ALTER TABLE [" & tableName & "] ADD COLUMN [" & fieldName & "] MEMO", dbFailOnError
    Exit Sub
Fail:
    If Err.Number = 3029 Or Err.Number = 3191 Or Err.Number = 3380 Or Err.Number = 3211 Then
        Err.Clear
    Else
        Err.Raise Err.Number, "AddMemoColumnIfMissing", Err.Description
    End If
End Sub

Public Sub DropColumnIfExists(ByVal tableName As String, ByVal fieldName As String)
    If Not FieldExists(tableName, fieldName) Then Exit Sub
    On Error GoTo Fail
    CurrentDb.Execute "ALTER TABLE [" & tableName & "] DROP COLUMN [" & fieldName & "]", dbFailOnError
    Exit Sub
Fail:
    ' 3211 = locked; ignore so UI rebuild can continue.
    If Err.Number = 3211 Or Err.Number = 3029 Or Err.Number = 3191 Then
        Err.Clear
    Else
        Err.Raise Err.Number, "DropColumnIfExists", Err.Description
    End If
End Sub

Public Sub ApplyDaysRagFormat(ByVal daysControl As Control)
    Dim fc As FormatCondition

    On Error GoTo CleanUp
    daysControl.FormatConditions.Delete

    Set fc = daysControl.FormatConditions.Add(acExpression, , "Nz([txtDays],0)>" & RAG_RED_DAYS)
    fc.BackColor = RGB(255, 199, 206)
    fc.ForeColor = RGB(156, 0, 6)

    Set fc = daysControl.FormatConditions.Add(acExpression, , _
        "Nz([txtDays],0)>" & RAG_YELLOW_DAYS & " And Nz([txtDays],0)<=" & RAG_RED_DAYS)
    fc.BackColor = RGB(255, 235, 156)
    fc.ForeColor = RGB(156, 101, 0)
CleanUp:
End Sub

Public Function DaysSinceDate(ByVal statusDate As Variant) As Variant
    If IsError(statusDate) Or IsNull(statusDate) Or IsEmpty(statusDate) Then
        DaysSinceDate = Null
    ElseIf Not IsDate(statusDate) Then
        DaysSinceDate = Null
    Else
        DaysSinceDate = DateDiff("d", CDate(statusDate), Date())
    End If
End Function
