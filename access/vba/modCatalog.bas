Attribute VB_Name = "modCatalog"
Option Compare Database
Option Explicit

' Rebuilds tblPart / tblPartDash from Assy_Standard the same way
' ListAssemblyFFAValues builds the Home catalog: split ASSEMBLY NO on the
' first dash, keep unique base parts and dash conditions, preserve Active.

Public Sub RebuildCatalogFromStandards()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim savedParts As Object
    Dim savedDashes As Object
    Dim assemblyNo As String
    Dim basePart As String
    Dim dashCondition As String
    Dim ffaValue As String
    Dim parts As Object
    Dim dashes As Object
    Dim ffas As Object
    Dim partKey As Variant
    Dim dashKey As Variant
    Dim ffaKey As Variant

    Set db = CurrentDb
    Set savedParts = SnapshotParts()
    Set savedDashes = SnapshotDashes()

    Set parts = CreateObject("Scripting.Dictionary")
    parts.CompareMode = vbTextCompare
    Set dashes = CreateObject("Scripting.Dictionary")
    dashes.CompareMode = vbTextCompare
    Set ffas = CreateObject("Scripting.Dictionary")
    ffas.CompareMode = vbTextCompare

    Set rs = db.OpenRecordset("SELECT [" & COL_ASSEMBLY_NO & "], [" & COL_FFA & "] FROM [" & TBL_ASSY_STANDARD & "]", dbOpenSnapshot)
    Do Until rs.EOF
        assemblyNo = CoerceText(rs(COL_ASSEMBLY_NO))
        ffaValue = CoerceText(rs(COL_FFA))
        If Len(assemblyNo) > 0 Then
            SplitAssemblyNo assemblyNo, basePart, dashCondition
            If Len(basePart) > 0 Then
                If Not parts.Exists(basePart) Then
                    parts.Add basePart, basePart
                    Set dashes(basePart) = CreateObject("Scripting.Dictionary")
                    dashes(basePart).CompareMode = vbTextCompare
                End If
                If Len(dashCondition) > 0 Then
                    If Not dashes(basePart).Exists(dashCondition) Then
                        dashes(basePart).Add dashCondition, dashCondition
                    End If
                End If
            End If
            If Len(ffaValue) > 0 Then
                If Not ffas.Exists(ffaValue) Then
                    ffas.Add ffaValue, ffaValue
                End If
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close

    For Each ffaKey In ffas.Keys
        If IsNull(DLookup("FFA", TBL_FFA, "FFA = " & SqlText(CStr(ffaKey)))) Then
            db.Execute "INSERT INTO [" & TBL_FFA & "] (FFA, Factory) VALUES (" & _
                SqlText(CStr(ffaKey)) & ", '')", dbFailOnError
        End If
    Next ffaKey

    db.Execute "DELETE FROM [" & TBL_PART_DASH & "]", dbFailOnError
    db.Execute "DELETE FROM [" & TBL_PART & "]", dbFailOnError

    For Each partKey In parts.Keys
        db.Execute "INSERT INTO [" & TBL_PART & "] (BasePart, Active, HomeFFA, StatusDate, [" & COL_NOTES & "]) VALUES (" & _
            SqlText(CStr(partKey)) & ", False, Null, Null, Null)", dbFailOnError
        RestorePartState db, CStr(partKey), savedParts
        If dashes.Exists(partKey) Then
            For Each dashKey In dashes(partKey).Keys
                db.Execute "INSERT INTO [" & TBL_PART_DASH & "] (BasePart, Dash, Active) VALUES (" & _
                    SqlText(CStr(partKey)) & ", " & SqlText(CStr(dashKey)) & ", False)", dbFailOnError
                RestoreDashState db, CStr(partKey), CStr(dashKey), savedDashes
            Next dashKey
        End If
        EnsureProductLineRows db, CStr(partKey)
    Next partKey

    ApplyRccpSelections
    RebuildActiveAssemblyFilter
End Sub

' tblRCCP drives which parts/dashes are Active and which product lines are used.
Public Sub ApplyRccpSelections()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim assemblyNo As String
    Dim basePart As String
    Dim dashCondition As String
    Dim baseFromCol As String
    Dim ffaValue As String
    Dim plCode As String
    Dim productLine As String
    Dim sql As String

    If Not TableExists(TBL_RCCP) Then Exit Sub
    EnsureProductLinePlCodeColumn

    Set db = CurrentDb
    db.Execute "UPDATE [" & TBL_PART & "] SET Active = False", dbFailOnError
    db.Execute "UPDATE [" & TBL_PART_DASH & "] SET Active = False", dbFailOnError
    If TableExists(TBL_PART_PL) Then
        db.Execute "UPDATE [" & TBL_PART_PL & "] SET UseFlag = False", dbFailOnError
    End If

    sql = "SELECT [" & COL_ASSEMBLY_NO & "]"
    If FieldExists(TBL_RCCP, COL_BASE_PN_TEXT) Then
        sql = sql & ", [" & COL_BASE_PN_TEXT & "]"
    End If
    If FieldExists(TBL_RCCP, COL_FFA) Then
        sql = sql & ", [" & COL_FFA & "]"
    End If
    If FieldExists(TBL_RCCP, COL_PRODUCT_LINE_TEXT) Then
        sql = sql & ", [" & COL_PRODUCT_LINE_TEXT & "]"
    End If
    sql = sql & " FROM [" & TBL_RCCP & "]"

    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    Do Until rs.EOF
        assemblyNo = CoerceText(rs.Fields(COL_ASSEMBLY_NO).Value)
        If Len(assemblyNo) > 0 Then
            SplitAssemblyNo assemblyNo, basePart, dashCondition
            If FieldExists(TBL_RCCP, COL_BASE_PN_TEXT) Then
                baseFromCol = CoerceText(rs.Fields(COL_BASE_PN_TEXT).Value)
                If Len(baseFromCol) > 0 Then basePart = baseFromCol
            End If

            If Len(basePart) > 0 Then
                EnsurePartRow db, basePart
                db.Execute "UPDATE [" & TBL_PART & "] SET Active = True WHERE BasePart = " & _
                    SqlText(basePart), dbFailOnError

                If FieldExists(TBL_RCCP, COL_FFA) Then
                    ffaValue = CoerceText(rs.Fields(COL_FFA).Value)
                    If Len(ffaValue) > 0 Then
                        EnsureFfaRow db, ffaValue
                        db.Execute "UPDATE [" & TBL_PART & "] SET HomeFFA = " & SqlText(ffaValue) & _
                            " WHERE BasePart = " & SqlText(basePart) & _
                            " AND (HomeFFA IS NULL OR HomeFFA = '')", dbFailOnError
                    End If
                End If

                If Len(dashCondition) > 0 Then
                    EnsureDashRow db, basePart, dashCondition
                    db.Execute "UPDATE [" & TBL_PART_DASH & "] SET Active = True WHERE BasePart = " & _
                        SqlText(basePart) & " AND Dash = " & SqlText(dashCondition), dbFailOnError
                End If

                If FieldExists(TBL_RCCP, COL_PRODUCT_LINE_TEXT) Then
                    plCode = CoerceText(rs.Fields(COL_PRODUCT_LINE_TEXT).Value)
                    If Len(plCode) > 0 Then
                        productLine = ResolveProductLineFromPlCode(db, plCode)
                        If Len(productLine) > 0 Then
                            EnsurePartProductLineRow db, basePart, productLine
                            db.Execute "UPDATE [" & TBL_PART_PL & "] SET UseFlag = True WHERE BasePart = " & _
                                SqlText(basePart) & " AND ProductLine = " & SqlText(productLine), dbFailOnError
                        End If
                    End If
                End If
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Sub EnsurePartRow(ByVal db As DAO.Database, ByVal basePart As String)
    If Not IsNull(DLookup("BasePart", TBL_PART, "BasePart = " & SqlText(basePart))) Then Exit Sub
    db.Execute "INSERT INTO [" & TBL_PART & "] (BasePart, Active, HomeFFA, StatusDate, [" & COL_NOTES & "]) VALUES (" & _
        SqlText(basePart) & ", False, Null, Null, Null)", dbFailOnError
    EnsureProductLineRows db, basePart
End Sub

Private Sub EnsureDashRow(ByVal db As DAO.Database, ByVal basePart As String, ByVal dashCondition As String)
    If Not IsNull(DLookup("Dash", TBL_PART_DASH, _
        "BasePart = " & SqlText(basePart) & " AND Dash = " & SqlText(dashCondition))) Then Exit Sub
    db.Execute "INSERT INTO [" & TBL_PART_DASH & "] (BasePart, Dash, Active) VALUES (" & _
        SqlText(basePart) & ", " & SqlText(dashCondition) & ", False)", dbFailOnError
End Sub

Private Sub EnsureFfaRow(ByVal db As DAO.Database, ByVal ffaValue As String)
    If Not IsNull(DLookup("FFA", TBL_FFA, "FFA = " & SqlText(ffaValue))) Then Exit Sub
    db.Execute "INSERT INTO [" & TBL_FFA & "] (FFA, Factory) VALUES (" & _
        SqlText(ffaValue) & ", '')", dbFailOnError
End Sub

Private Function ResolveProductLineFromPlCode(ByVal db As DAO.Database, ByVal plCode As String) As String
    Dim existing As Variant
    existing = DLookup("ProductLine", TBL_PRODUCT_LINE, "[" & COL_PL_CODE & "] = " & SqlText(plCode))
    If Not IsNull(existing) Then
        ResolveProductLineFromPlCode = CoerceText(existing)
        Exit Function
    End If
    ' Also allow ProductLine name to match the RCCP code directly.
    existing = DLookup("ProductLine", TBL_PRODUCT_LINE, "ProductLine = " & SqlText(plCode))
    If Not IsNull(existing) Then
        On Error Resume Next
        db.Execute "UPDATE [" & TBL_PRODUCT_LINE & "] SET [" & COL_PL_CODE & "] = " & SqlText(plCode) & _
            " WHERE ProductLine = " & SqlText(plCode) & _
            " AND ([" & COL_PL_CODE & "] IS NULL OR [" & COL_PL_CODE & "] = '')", dbFailOnError
        On Error GoTo 0
        ResolveProductLineFromPlCode = plCode
        Exit Function
    End If
    db.Execute "INSERT INTO [" & TBL_PRODUCT_LINE & "] (ProductLine, [" & COL_PL_CODE & "]) VALUES (" & _
        SqlText(plCode) & ", " & SqlText(plCode) & ")", dbFailOnError
    ResolveProductLineFromPlCode = plCode
End Function

Private Sub EnsurePartProductLineRow(ByVal db As DAO.Database, ByVal basePart As String, ByVal productLine As String)
    If Not IsNull(DLookup("ProductLine", TBL_PART_PL, _
        "BasePart = " & SqlText(basePart) & " AND ProductLine = " & SqlText(productLine))) Then Exit Sub
    db.Execute "INSERT INTO [" & TBL_PART_PL & "] (BasePart, ProductLine, UseFlag) VALUES (" & _
        SqlText(basePart) & ", " & SqlText(productLine) & ", False)", dbFailOnError
End Sub

Private Function SnapshotParts() As Object
    Dim rs As DAO.Recordset
    Dim map As Object
    Dim key As String
    Dim notes As String
    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = vbTextCompare
    If Not TableExists(TBL_PART) Then
        Set SnapshotParts = map
        Exit Function
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM [" & TBL_PART & "]", dbOpenSnapshot)
    Do Until rs.EOF
        key = CoerceText(rs!BasePart)
        If Len(key) > 0 Then
            If FieldExists(TBL_PART, COL_NOTES) Then
                notes = CoerceText(rs.Fields(COL_NOTES).Value)
            ElseIf FieldExists(TBL_PART, "Highlight") Then
                notes = CoerceText(rs.Fields("Highlight").Value)
            Else
                notes = vbNullString
            End If
            map.Add key, Array( _
                CBool(Nz(rs!Active, False)), _
                CoerceText(rs!HomeFFA), _
                rs!StatusDate, _
                notes)
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set SnapshotParts = map
End Function

Private Function SnapshotDashes() As Object
    Dim rs As DAO.Recordset
    Dim map As Object
    Dim key As String
    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = vbTextCompare
    If Not TableExists(TBL_PART_DASH) Then
        Set SnapshotDashes = map
        Exit Function
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM [" & TBL_PART_DASH & "]", dbOpenSnapshot)
    Do Until rs.EOF
        key = CoerceText(rs!BasePart) & vbTab & CoerceText(rs!Dash)
        If Not map.Exists(key) Then
            map.Add key, CBool(Nz(rs!Active, False))
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set SnapshotDashes = map
End Function

Private Sub RestorePartState(ByVal db As DAO.Database, ByVal basePart As String, ByVal savedParts As Object)
    Dim state As Variant
    If Not savedParts.Exists(basePart) Then Exit Sub
    state = savedParts(basePart)
    db.Execute "UPDATE [" & TBL_PART & "] SET " & _
        "Active = " & SqlBool(state(0)) & ", " & _
        "HomeFFA = " & SqlNullableText(CStr(state(1))) & ", " & _
        "StatusDate = " & SqlNullableDate(state(2)) & ", " & _
        "[" & COL_NOTES & "] = " & SqlNullableText(CStr(state(3))) & " " & _
        "WHERE BasePart = " & SqlText(basePart), dbFailOnError
End Sub

Private Sub RestoreDashState(ByVal db As DAO.Database, ByVal basePart As String, ByVal dashCondition As String, ByVal savedDashes As Object)
    Dim key As String
    key = basePart & vbTab & dashCondition
    If Not savedDashes.Exists(key) Then Exit Sub
    db.Execute "UPDATE [" & TBL_PART_DASH & "] SET Active = " & SqlBool(savedDashes(key)) & _
        " WHERE BasePart = " & SqlText(basePart) & " AND Dash = " & SqlText(dashCondition), dbFailOnError
End Sub

Private Sub EnsureProductLineRows(ByVal db As DAO.Database, ByVal basePart As String)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT ProductLine FROM [" & TBL_PRODUCT_LINE & "]", dbOpenSnapshot)
    Do Until rs.EOF
        If IsNull(DLookup("ProductLine", TBL_PART_PL, _
            "BasePart = " & SqlText(basePart) & " AND ProductLine = " & SqlText(CStr(rs!ProductLine)))) Then
            db.Execute "INSERT INTO [" & TBL_PART_PL & "] (BasePart, ProductLine, UseFlag) VALUES (" & _
                SqlText(basePart) & ", " & SqlText(CStr(rs!ProductLine)) & ", False)", dbFailOnError
        End If
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Function SqlBool(ByVal value As Boolean) As String
    If value Then
        SqlBool = "True"
    Else
        SqlBool = "False"
    End If
End Function

Private Function SqlNullableText(ByVal value As String) As String
    If Len(value) = 0 Then
        SqlNullableText = "Null"
    Else
        SqlNullableText = SqlText(value)
    End If
End Function

Private Function SqlNullableDate(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        SqlNullableDate = "Null"
    ElseIf Not IsDate(value) Then
        SqlNullableDate = "Null"
    Else
        SqlNullableDate = "#" & Format$(CDate(value), "yyyy-mm-dd") & "#"
    End If
End Function
