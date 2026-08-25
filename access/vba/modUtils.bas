Attribute VB_Name = "modUtils"
Option Compare Database
Option Explicit

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
    Dim containerName As String

    Select Case objectType
        Case acForm
            containerName = "Forms"
        Case acReport
            containerName = "Reports"
        Case acMacro
            containerName = "Scripts"
        Case Else
            ObjectExists = False
            Exit Function
    End Select

    For i = 0 To CurrentProject.AllForms.Count - 1
        ' Walked below by type.
    Next i

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

Public Function CoerceLong(ByVal rawValue As Variant) As Variant
    If IsError(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then
        CoerceLong = Null
    ElseIf Len(Trim$(CStr(rawValue))) = 0 Then
        CoerceLong = Null
    ElseIf IsNumeric(rawValue) Then
        CoerceLong = CLng(rawValue)
    Else
        CoerceLong = Null
    End If
End Function

Public Function CoerceDouble(ByVal rawValue As Variant) As Variant
    If IsError(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then
        CoerceDouble = Null
    ElseIf Len(Trim$(CStr(rawValue))) = 0 Then
        CoerceDouble = Null
    ElseIf IsNumeric(rawValue) Then
        CoerceDouble = CDbl(rawValue)
    Else
        CoerceDouble = Null
    End If
End Function

Public Sub DeleteObjectIfExists(ByVal objectType As AcObjectType, ByVal objectName As String)
    On Error Resume Next
    DoCmd.Close objectType, objectName, acSaveNo
    DoCmd.DeleteObject objectType, objectName
    On Error GoTo 0
End Sub

' Close open forms so schema DDL / form rebuild can lock local tables (avoids 3211).
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
    DoCmd.Close acForm, SFRM_HOME_LIST, acSaveNo
    DoCmd.Close acForm, SFRM_DASH, acSaveNo
    DoCmd.Close acForm, SFRM_PL, acSaveNo
    DoCmd.Close acForm, SFRM_OPS, acSaveNo
    ' Also clear any leftover Design-view FormN windows from a prior CreateForm.
    safety = 0
    Do While Forms.Count > 0 And safety < 50
        DoCmd.Close acForm, Forms(0).Name, acSaveNo
        safety = safety + 1
    Loop
    DoEvents
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
    If Err.Number = 3029 Or Err.Number = 3191 Or Err.Number = 3380 Then
        Err.Clear
    ElseIf Err.Number = 3211 Then
        Err.Raise Err.Number, "AddTextColumnIfMissing", _
            Err.Description & vbCrLf & vbCrLf & _
            "Close all open forms (especially Home) and run BuildUi / Bootstrap again."
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
    If Err.Number = 3029 Or Err.Number = 3191 Or Err.Number = 3380 Then
        Err.Clear
    Else
        Err.Raise Err.Number, "AddMemoColumnIfMissing", Err.Description
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
