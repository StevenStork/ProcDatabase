Attribute VB_Name = "modReferences"
Option Compare Database
Option Explicit

' Seeds lookup tables from linked source data.
' Equipment is user-maintained (name, type, and FFA assignments) — not seeded from sources.
' Catalog rebuild already inserts FFAs from standards; this fills any extras from the
' active (filtered) standards query without per-row DLookup.

Public Sub SeedReferencesFromSources()
    SeedFfasFromStandards
End Sub

Public Sub SeedFfasFromStandards()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim existing As Object
    Dim ffaValue As String
    Dim sourceName As String

    If Not TableExists(TBL_ASSY_STANDARD) Then Exit Sub
    Set db = CurrentDb
    sourceName = AssyStndSourceName()

    Set existing = CreateObject("Scripting.Dictionary")
    existing.CompareMode = vbTextCompare
    Set rs = db.OpenRecordset("SELECT FFA FROM [" & TBL_FFA & "]", dbOpenSnapshot)
    Do Until rs.EOF
        ffaValue = CoerceText(rs!FFA)
        If Len(ffaValue) > 0 And Not existing.Exists(ffaValue) Then existing.Add ffaValue, True
        rs.MoveNext
    Loop
    rs.Close

    Set rs = db.OpenRecordset( _
        "SELECT DISTINCT [" & COL_FFA & "] FROM [" & sourceName & "] WHERE Len(Nz([" & COL_FFA & "],''))>0", _
        dbOpenSnapshot)
    Do Until rs.EOF
        ffaValue = CoerceText(rs(COL_FFA))
        If Len(ffaValue) > 0 And Not existing.Exists(ffaValue) Then
            db.Execute "INSERT INTO [" & TBL_FFA & "] (FFA, Factory) VALUES (" & _
                SqlText(ffaValue) & ", '')", dbFailOnError
            existing.Add ffaValue, True
        End If
        rs.MoveNext
    Loop
    rs.Close
End Sub
