Attribute VB_Name = "modExport"
Option Compare Database
Option Explicit

Public Sub ExportOps(ByVal exportType As String, Optional ByVal exportKey As String = vbNullString)
    Dim destPath As String
    Dim folder As String
    Dim stamp As String

    stamp = Format$(Now, "yyyymmdd_hhnn")
    folder = CurrentProject.Path
    exportType = Trim$(exportType)

    Select Case UCase$(exportType)
        Case "FFA", "ALL"
            destPath = folder & "\FFA_Export_" & stamp & ".xlsx"
            ExportQuery "SELECT * FROM [" & QRY_EXPORT & "]", destPath, "FFA Export"
            If StrComp(exportType, "ALL", vbTextCompare) <> 0 Then
                MsgBox "Wrote " & destPath, vbInformation, "Export"
                Exit Sub
            End If
            ExportEachProductLine folder, stamp
            MsgBox "Wrote FFA and product-line workbooks in " & folder, vbInformation, "Export"
        Case "PRODUCTLINE", "PRODUCT LINE", "PL"
            If Len(Trim$(exportKey)) = 0 Then
                Err.Raise vbObjectError + 610, "ExportOps", "Select a product line to export."
            End If
            destPath = folder & "\PL_" & MakeFileSafe(exportKey) & "_" & stamp & ".xlsx"
            ExportQuery ProductLineSql(exportKey), destPath, Left$("PL - " & exportKey, 31)
            MsgBox "Wrote " & destPath, vbInformation, "Export"
        Case Else
            Err.Raise vbObjectError + 611, "ExportOps", "Unknown export type: " & exportType
    End Select
End Sub

Private Sub ExportEachProductLine(ByVal folder As String, ByVal stamp As String)
    Dim rs As DAO.Recordset
    Dim productLine As String
    Dim destPath As String
    Set rs = CurrentDb.OpenRecordset("SELECT ProductLine FROM [" & TBL_PRODUCT_LINE & "] ORDER BY ProductLine", dbOpenSnapshot)
    Do Until rs.EOF
        productLine = CoerceText(rs!ProductLine)
        If Len(productLine) > 0 Then
            If DCount("*", TBL_PART_PL, "ProductLine = " & SqlText(productLine) & " AND UseFlag <> 0") > 0 Then
                destPath = folder & "\PL_" & MakeFileSafe(productLine) & "_" & stamp & ".xlsx"
                ExportQuery ProductLineSql(productLine), destPath, Left$("PL - " & productLine, 31)
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Function ProductLineSql(ByVal productLine As String) As String
    ProductLineSql = "SELECT e.* FROM [" & QRY_EXPORT & "] AS e " & _
        "INNER JOIN [" & TBL_PART_PL & "] AS pl ON e.[Part Number] = pl.BasePart " & _
        "WHERE pl.UseFlag <> 0 AND pl.ProductLine = " & SqlText(productLine)
End Function

Private Sub ExportQuery(ByVal sql As String, ByVal destPath As String, ByVal sheetName As String)
    Dim qdf As DAO.QueryDef
    Dim tempName As String
    tempName = "qryExportTemp"
    ReplaceQuery tempName, sql
    If Len(Dir$(destPath)) > 0 Then
        Kill destPath
    End If
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, tempName, destPath, True, sheetName
    CurrentDb.QueryDefs.Delete tempName
End Sub

Private Function MakeFileSafe(ByVal textValue As String) As String
    Dim i As Long
    Dim ch As String
    Dim result As String
    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        Select Case ch
            Case "\", "/", ":", "*", "?", """", "<", ">", "|", " "
                result = result & "_"
            Case Else
                result = result & ch
        End Select
    Next i
    If Len(result) = 0 Then result = "export"
    MakeFileSafe = result
End Function
