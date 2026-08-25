Attribute VB_Name = "modReferences"
Option Compare Database
Option Explicit

' Seeds lookup tables from linked source data.
' Equipment is user-maintained (name + FFA assignments) — not seeded from sources.

Public Sub SeedReferencesFromSources()
    SeedFfasFromStandards
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
