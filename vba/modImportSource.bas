Attribute VB_Name = "modImportSource"
Option Explicit

' Re-imports vba/*.bas and vba/*.frm from a folder next to this workbook.
' Requires Excel Trust Center → Macro Settings → "Trust access to the VBA project object model".
Public Sub ImportSourceModules()
    Dim folderPath As String
    Dim fileName As String
    Dim vbProj As Object
    Dim vbComp As Object
    Dim fullPath As String

    folderPath = ThisWorkbook.Path
    If Right$(folderPath, 1) <> Application.PathSeparator Then
        folderPath = folderPath & Application.PathSeparator
    End If
    If Len(Dir$(folderPath & "vba", vbDirectory)) > 0 Then
        folderPath = folderPath & "vba" & Application.PathSeparator
    ElseIf Len(Dir$(ThisWorkbook.Path & Application.PathSeparator & ".." & Application.PathSeparator & "vba", vbDirectory)) > 0 Then
        folderPath = ThisWorkbook.Path & Application.PathSeparator & ".." & Application.PathSeparator & "vba" & Application.PathSeparator
    End If

    On Error GoTo Fail
    Set vbProj = ThisWorkbook.VBProject

    fileName = Dir$(folderPath & "*.bas")
    Do While Len(fileName) > 0
        If StrComp(Right$(fileName, 4), ".txt", vbTextCompare) <> 0 Then
            fullPath = folderPath & fileName
            RemoveComponentIfExists vbProj, ModuleNameFromFile(fileName)
            vbProj.VBComponents.Import fullPath
        End If
        fileName = Dir$
    Loop

    fileName = Dir$(folderPath & "*.frm")
    Do While Len(fileName) > 0
        fullPath = folderPath & fileName
        RemoveComponentIfExists vbProj, ModuleNameFromFile(fileName)
        vbProj.VBComponents.Import fullPath
        fileName = Dir$
    Loop

    MsgBox "VBA source import finished from:" & vbCrLf & folderPath, vbInformation, "Import Source Modules"
    Exit Sub

Fail:
    MsgBox "Could not import VBA source. Enable Trust access to the VBA project object model." & vbCrLf & vbCrLf & Err.Description, _
        vbCritical, "Import Source Modules"
End Sub

Private Sub RemoveComponentIfExists(ByVal vbProj As Object, ByVal componentName As String)
    Dim vbComp As Object

    On Error Resume Next
    Set vbComp = vbProj.VBComponents(componentName)
    If Not vbComp Is Nothing Then vbProj.VBComponents.Remove vbComp
    On Error GoTo 0
End Sub

Private Function ModuleNameFromFile(ByVal fileName As String) As String
    Dim dotPos As Long

    dotPos = InStrRev(fileName, ".")
    If dotPos > 1 Then
        ModuleNameFromFile = Left$(fileName, dotPos - 1)
    Else
        ModuleNameFromFile = fileName
    End If
End Function
