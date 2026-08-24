Attribute VB_Name = "modReferences"
Option Compare Database
Option Explicit

' Seeds lookup tables from linked source data (replaces hidden References sheet).

Public Sub SeedReferencesFromSources()
    SeedFfasFromStandards
    SeedEquipmentFromRouteCard
End Sub

Public Sub SeedFfasFromStandards()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim ffaValue As String
    Dim sourceName As String

    If Not TableExists(TBL_ASSY_STANDARD) Then Exit Sub
    Set db = CurrentDb
    sourceName = AssyStndSourceName()
    Set rs = db.OpenRecordset("SELECT DISTINCT [" & COL_FFA & "] FROM [" & sourceName & "] WHERE Len(Nz([" & COL_FFA & "],''))>0", dbOpenSnapshot)
    Do Until rs.EOF
        ffaValue = CoerceText(rs(COL_FFA))
        If Len(ffaValue) > 0 Then
            If IsNull(DLookup("FFA", TBL_FFA, "FFA = " & SqlText(ffaValue))) Then
                db.Execute "INSERT INTO [" & TBL_FFA & "] (FFA, Factory) VALUES (" & _
                    SqlText(ffaValue) & ", '')", dbFailOnError
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
End Sub

Public Sub SeedEquipmentFromRouteCard()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim equipName As String
    Dim ffaValue As String
    Dim owners As String
    Dim sourceName As String

    If Not TableExists(TBL_ROUTE_CARD) Then Exit Sub
    Set db = CurrentDb
    sourceName = RouteCardSourceName()
    Set rs = db.OpenRecordset( _
        "SELECT DISTINCT [" & COL_OPER_DESC & "], [" & COL_FFA & "] FROM [" & sourceName & "] " & _
        "WHERE Len(Nz([" & COL_OPER_DESC & "],''))>0", dbOpenSnapshot)
    Do Until rs.EOF
        equipName = CoerceText(rs(COL_OPER_DESC))
        ffaValue = CoerceText(rs(COL_FFA))
        If Len(equipName) > 0 Then
            If IsNull(DLookup("Equipment", TBL_EQUIPMENT, "Equipment = " & SqlText(equipName))) Then
                owners = ffaValue
                db.Execute "INSERT INTO [" & TBL_EQUIPMENT & "] (Equipment, OwningFFAs) VALUES (" & _
                    SqlText(equipName) & ", " & SqlText(owners) & ")", dbFailOnError
            ElseIf Len(ffaValue) > 0 Then
                owners = CoerceText(Nz(DLookup("OwningFFAs", TBL_EQUIPMENT, "Equipment = " & SqlText(equipName)), vbNullString))
                If Not OwnerListContains(owners, ffaValue) Then
                    If Len(owners) > 0 Then owners = owners & ", "
                    owners = owners & ffaValue
                    db.Execute "UPDATE [" & TBL_EQUIPMENT & "] SET OwningFFAs = " & SqlText(owners) & _
                        " WHERE Equipment = " & SqlText(equipName), dbFailOnError
                End If
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Function OwnerListContains(ByVal ownerList As String, ByVal ffaValue As String) As Boolean
    Dim parts() As String
    Dim i As Long
    If Len(ownerList) = 0 Then Exit Function
    parts = Split(ownerList, ",")
    For i = LBound(parts) To UBound(parts)
        If StrComp(Trim$(parts(i)), ffaValue, vbTextCompare) = 0 Then
            OwnerListContains = True
            Exit Function
        End If
    Next i
End Function
