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
    Set db = CurrentDb
    db.Execute "DELETE FROM [" & TBL_META & "] WHERE [Key] = " & SqlText(metaKey), dbFailOnError
    db.Execute "INSERT INTO [" & TBL_META & "] ([Key], [Value]) VALUES (" & _
        SqlText(metaKey) & ", " & SqlText(metaValue) & ")", dbFailOnError
End Sub

Public Function GetMeta(ByVal metaKey As String) As String
    GetMeta = Nz(DLookup("[Value]", TBL_META, "[Key] = " & SqlText(metaKey)), vbNullString)
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
    Dim td As DAO.TableDef
    Dim fld As DAO.Field
    On Error Resume Next
    Set td = CurrentDb.TableDefs(tableName)
    Set fld = td.Fields(fieldName)
    FieldExists = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Sub AddTextFieldIfMissing(ByVal tableName As String, ByVal fieldName As String, ByVal size As Long)
    Dim td As DAO.TableDef
    Dim fld As DAO.Field
    If FieldExists(tableName, fieldName) Then Exit Sub
    Set td = CurrentDb.TableDefs(tableName)
    Set fld = td.CreateField(fieldName, dbText, size)
    fld.AllowZeroLength = True
    td.Fields.Append fld
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
