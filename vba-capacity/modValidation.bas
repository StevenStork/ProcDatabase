Attribute VB_Name = "modValidation"
Option Explicit

'==============================================================================
' Validation helpers for capacity database keys and foreign keys.
'==============================================================================

Public Function NormalizeCode(ByVal rawValue As Variant) As String
    NormalizeCode = Trim$(CStr(Nz(rawValue)))
End Function

Public Function IsActiveFlag(ByVal rawValue As Variant) As Boolean
    Dim textValue As String

    If IsError(rawValue) Then Exit Function
    If IsEmpty(rawValue) Or IsNull(rawValue) Then Exit Function

    Select Case VarType(rawValue)
        Case vbBoolean
            IsActiveFlag = rawValue
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            IsActiveFlag = (CDbl(rawValue) <> 0)
        Case Else
            textValue = LCase$(Trim$(CStr(rawValue)))
            IsActiveFlag = (textValue = "true" Or textValue = "yes" Or textValue = "y" Or textValue = "1")
    End Select
End Function

Public Function ActiveFlagToCellValue(ByVal isActive As Boolean) As Variant
    If isActive Then
        ActiveFlagToCellValue = True
    Else
        ActiveFlagToCellValue = False
    End If
End Function

Public Function ValidateRequiredCode(ByVal codeValue As String, ByVal fieldLabel As String) As Boolean
    If Len(NormalizeCode(codeValue)) = 0 Then
        MsgBox fieldLabel & " is required.", vbExclamation, "Validation"
        ValidateRequiredCode = False
    Else
        ValidateRequiredCode = True
    End If
End Function

Public Function ValidateUniqueKey( _
    ByVal tbl As ListObject, _
    ByVal keyColumnName As String, _
    ByVal keyValue As String, _
    ByVal excludeListRow As Long) As Boolean

    Dim existingRow As Long

    existingRow = FindListRowByKey(tbl, keyColumnName, keyValue)
    If existingRow = 0 Then
        ValidateUniqueKey = True
        Exit Function
    End If

    If excludeListRow > 0 And existingRow = excludeListRow Then
        ValidateUniqueKey = True
    Else
        MsgBox "A record with " & keyColumnName & " '" & keyValue & "' already exists.", _
            vbExclamation, "Validation"
        ValidateUniqueKey = False
    End If
End Function

Public Function ValidateForeignKeyExists( _
    ByVal tbl As ListObject, _
    ByVal keyColumnName As String, _
    ByVal keyValue As String, _
    ByVal entityLabel As String) As Boolean

    If Len(NormalizeCode(keyValue)) = 0 Then
        MsgBox entityLabel & " is required.", vbExclamation, "Validation"
        ValidateForeignKeyExists = False
        Exit Function
    End If

    If FindListRowByKey(tbl, keyColumnName, keyValue) = 0 Then
        MsgBox entityLabel & " '" & keyValue & "' was not found.", vbExclamation, "Validation"
        ValidateForeignKeyExists = False
    Else
        ValidateForeignKeyExists = True
    End If
End Function

Public Function ValidateJunctionUnique( _
    ByVal tbl As ListObject, _
    ByVal key1ColumnName As String, _
    ByVal key1Value As String, _
    ByVal key2ColumnName As String, _
    ByVal key2Value As String, _
    ByVal excludeListRow As Long) As Boolean

    Dim existingRow As Long

    existingRow = FindJunctionListRow(tbl, key1ColumnName, key1Value, key2ColumnName, key2Value)
    If existingRow = 0 Then
        ValidateJunctionUnique = True
        Exit Function
    End If

    If excludeListRow > 0 And existingRow = excludeListRow Then
        ValidateJunctionUnique = True
    Else
        MsgBox "That assignment already exists.", vbExclamation, "Validation"
        ValidateJunctionUnique = False
    End If
End Function

Public Function Nz(ByVal value As Variant) As Variant
    If IsError(value) Then
        Nz = vbNullString
    ElseIf IsEmpty(value) Or IsNull(value) Then
        Nz = vbNullString
    Else
        Nz = value
    End If
End Function
