Attribute VB_Name = "modReferences"
Option Explicit

'==============================================================================
' References sheet schema (sheet is always xlSheetVeryHidden; A1 = "Refs")
'
'   A1  Refs          sheet type — excluded from Home Show/Hide buttons
'   B1  FFA           header
'   C1  Factory       header (factory for the FFA on the same row)
'   D1  Product Line  header
'   E1  Equipment     header
'   F1  Owning FFAs   header (comma-separated FFA names from column B)
'
' Data starts at row 2. B/C are paired (one FFA + factory per row).
' D and E/F are independent lists and may be different lengths.
' Equipment ownership is many-to-many: one equipment row can list several
' FFAs, and one FFA can appear on several equipment rows.
'==============================================================================

Public Const REFS_SHEET_NAME As String = "References"
Public Const REFS_LABEL_CELL As String = "A1"
Public Const REFS_HEADER_ROW As Long = 1

Public Const REFS_HEADER_FFA As String = "FFA"
Public Const REFS_HEADER_FACTORY As String = "Factory"
Public Const REFS_HEADER_PRODUCT_LINE As String = "Product Line"
Public Const REFS_HEADER_EQUIPMENT As String = "Equipment"
Public Const REFS_HEADER_OWNERS As String = "Owning FFAs"

Public Const REFS_CATEGORY_FFA As String = "FFA"
Public Const REFS_CATEGORY_PRODUCT_LINE As String = "Product Line"
Public Const REFS_CATEGORY_EQUIPMENT As String = "Equipment"

Public Sub ShowUpdateReferences()
    Dim frm As Object

    EnsureReferencesSheet
    On Error GoTo FailForm
    Set frm = UserForms.Add("frmReferences")
    frm.Show vbModal
    Exit Sub

FailForm:
    MsgBox "The Update References form is not in this workbook." & vbCrLf & vbCrLf & _
        "Import vba/frmReferences.frm to edit FFA, product line, and equipment lists.", _
        vbExclamation, "Update References"
End Sub

Public Function LoadReferenceData() As Object
    Dim data As Object
    Dim ws As Worksheet

    EnsureReferencesSheet
    Set ws = ThisWorkbook.Worksheets(REFS_SHEET_NAME)

    Set data = CreateObject("Scripting.Dictionary")
    data.CompareMode = vbTextCompare
    data.Add "Ffas", ReadFfaFactoryMap(ws)
    data.Add "ProductLines", ReadNameMap(ws, REFS_PRODUCT_LINE_COLUMN)
    data.Add "Equipment", ReadEquipmentMap(ws)

    Set LoadReferenceData = data
End Function

Public Sub SaveReferenceData(ByVal data As Object)
    Dim ws As Worksheet
    Dim ffas As Object
    Dim productLines As Object
    Dim equipment As Object

    EnsureReferencesSheet
    Set ws = ThisWorkbook.Worksheets(REFS_SHEET_NAME)

    Set ffas = data("Ffas")
    Set productLines = data("ProductLines")
    Set equipment = data("Equipment")

    WritePairedColumn ws, REFS_FFA_COLUMN, REFS_FACTORY_COLUMN, ffas
    WriteNameColumn ws, REFS_PRODUCT_LINE_COLUMN, productLines
    WriteEquipmentColumns ws, equipment
    HideReferencesSheet ws
    MarkReferencesStale
End Sub

Public Function NormalizeOwnerList(ByVal rawList As String, ByVal knownFfas As Object) As String
    Dim parts As Variant
    Dim i As Long
    Dim token As String
    Dim ffaName As String
    Dim uniqueNames As Object
    Dim keys() As String
    Dim key As Variant
    Dim index As Long

    Set uniqueNames = CreateObject("Scripting.Dictionary")
    uniqueNames.CompareMode = vbTextCompare

    parts = Split(rawList, ",")
    For i = LBound(parts) To UBound(parts)
        token = Trim$(CStr(parts(i)))
        If Len(token) = 0 Then GoTo NextPart

        ffaName = ResolveOwnerToFfa(token, knownFfas)
        If Len(ffaName) = 0 Then GoTo NextPart
        If Not uniqueNames.Exists(ffaName) Then uniqueNames.Add ffaName, ffaName
NextPart:
    Next i

    If uniqueNames.Count = 0 Then
        NormalizeOwnerList = vbNullString
        Exit Function
    End If

    ReDim keys(0 To uniqueNames.Count - 1)
    index = 0
    For Each key In uniqueNames.Keys
        keys(index) = CStr(key)
        index = index + 1
    Next key
    QuickSortStrings keys, LBound(keys), UBound(keys)
    NormalizeOwnerList = Join(keys, ", ")
End Function

' Keep FFA values (column B). Tokens that already match an FFA are kept;
' factory names (column C) are mapped back to their FFA.
Private Function ResolveOwnerToFfa(ByVal token As String, ByVal knownFfas As Object) As String
    Dim key As Variant

    If knownFfas Is Nothing Then
        ResolveOwnerToFfa = token
        Exit Function
    End If

    For Each key In knownFfas.Keys
        If StrComp(CStr(key), token, vbTextCompare) = 0 Then
            ResolveOwnerToFfa = CStr(key)
            Exit Function
        End If
    Next key

    For Each key In knownFfas.Keys
        If StrComp(CStr(knownFfas(key)), token, vbTextCompare) = 0 Then
            ResolveOwnerToFfa = CStr(key)
            Exit Function
        End If
    Next key
End Function

Public Function RenameFfaInEquipment(ByVal equipment As Object, ByVal oldName As String, ByVal newName As String) As Object
    Dim equipName As Variant
    Dim owners As String
    Dim parts As Variant
    Dim i As Long
    Dim ownerName As String
    Dim uniqueNames As Object
    Dim keys() As String
    Dim key As Variant
    Dim index As Long

    For Each equipName In equipment.Keys
        owners = CStr(equipment(equipName))
        If InStr(1, ", " & owners & ", ", ", " & oldName & ", ", vbTextCompare) = 0 Then GoTo NextEquip

        Set uniqueNames = CreateObject("Scripting.Dictionary")
        uniqueNames.CompareMode = vbTextCompare
        parts = Split(owners, ",")
        For i = LBound(parts) To UBound(parts)
            ownerName = Trim$(CStr(parts(i)))
            If StrComp(ownerName, oldName, vbTextCompare) = 0 Then ownerName = newName
            If Len(ownerName) > 0 Then
                If Not uniqueNames.Exists(ownerName) Then uniqueNames.Add ownerName, ownerName
            End If
        Next i

        If uniqueNames.Count = 0 Then
            equipment(equipName) = vbNullString
        Else
            ReDim keys(0 To uniqueNames.Count - 1)
            index = 0
            For Each key In uniqueNames.Keys
                keys(index) = CStr(key)
                index = index + 1
            Next key
            QuickSortStrings keys, LBound(keys), UBound(keys)
            equipment(equipName) = Join(keys, ", ")
        End If
NextEquip:
    Next equipName

    Set RenameFfaInEquipment = equipment
End Function

Public Function RemoveFfaFromEquipment(ByVal equipment As Object, ByVal ffaName As String) As Object
    Dim equipName As Variant
    Dim parts As Variant
    Dim i As Long
    Dim ownerName As String
    Dim keepers As Collection
    Dim kept() As String
    Dim index As Long

    For Each equipName In equipment.Keys
        Set keepers = New Collection
        parts = Split(CStr(equipment(equipName)), ",")
        For i = LBound(parts) To UBound(parts)
            ownerName = Trim$(CStr(parts(i)))
            If Len(ownerName) > 0 Then
                If StrComp(ownerName, ffaName, vbTextCompare) <> 0 Then keepers.Add ownerName
            End If
        Next i

        If keepers.Count = 0 Then
            equipment(equipName) = vbNullString
        Else
            ReDim kept(0 To keepers.Count - 1)
            For index = 1 To keepers.Count
                kept(index - 1) = CStr(keepers(index))
            Next index
            equipment(equipName) = Join(kept, ", ")
        End If
    Next equipName

    Set RemoveFfaFromEquipment = equipment
End Function

Private Function ReadFfaFactoryMap(ByVal ws As Worksheet) As Object
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim ffaName As String
    Dim factoryName As String
    Dim ffas As Object

    Set ffas = CreateObject("Scripting.Dictionary")
    ffas.CompareMode = vbTextCompare

    lastRow = LastUsedRowInColumn(ws, REFS_FFA_COLUMN)
    If lastRow < REFS_FIRST_DATA_ROW Then
        Set ReadFfaFactoryMap = ffas
        Exit Function
    End If

    For rowIndex = REFS_FIRST_DATA_ROW To lastRow
        ffaName = Trim$(CStr(Nz(ws.Cells(rowIndex, REFS_FFA_COLUMN).Value)))
        factoryName = Trim$(CStr(Nz(ws.Cells(rowIndex, REFS_FACTORY_COLUMN).Value)))
        If Len(ffaName) > 0 Then
            If Not ffas.Exists(ffaName) Then ffas.Add ffaName, factoryName
        End If
    Next rowIndex

    Set ReadFfaFactoryMap = ffas
End Function

Private Function ReadNameMap(ByVal ws As Worksheet, ByVal columnLetter As String) As Object
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim itemName As String
    Dim names As Object

    Set names = CreateObject("Scripting.Dictionary")
    names.CompareMode = vbTextCompare

    lastRow = LastUsedRowInColumn(ws, columnLetter)
    If lastRow < REFS_FIRST_DATA_ROW Then
        Set ReadNameMap = names
        Exit Function
    End If

    For rowIndex = REFS_FIRST_DATA_ROW To lastRow
        itemName = Trim$(CStr(Nz(ws.Cells(rowIndex, columnLetter).Value)))
        If Len(itemName) > 0 Then
            If Not names.Exists(itemName) Then names.Add itemName, itemName
        End If
    Next rowIndex

    Set ReadNameMap = names
End Function

Private Function ReadEquipmentMap(ByVal ws As Worksheet) As Object
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim equipName As String
    Dim owners As String
    Dim equipment As Object
    Dim ffas As Object

    Set equipment = CreateObject("Scripting.Dictionary")
    equipment.CompareMode = vbTextCompare
    Set ffas = ReadFfaFactoryMap(ws)

    lastRow = LastUsedRowInColumn(ws, REFS_EQUIPMENT_COLUMN)
    If lastRow < REFS_FIRST_DATA_ROW Then
        Set ReadEquipmentMap = equipment
        Exit Function
    End If

    For rowIndex = REFS_FIRST_DATA_ROW To lastRow
        equipName = Trim$(CStr(Nz(ws.Cells(rowIndex, REFS_EQUIPMENT_COLUMN).Value)))
        owners = Trim$(CStr(Nz(ws.Cells(rowIndex, REFS_EQUIPMENT_OWNERS_COLUMN).Value)))
        If Len(equipName) > 0 Then
            If Not equipment.Exists(equipName) Then
                equipment.Add equipName, NormalizeOwnerList(owners, ffas)
            End If
        End If
    Next rowIndex

    Set ReadEquipmentMap = equipment
End Function

Private Sub WritePairedColumn( _
    ByVal ws As Worksheet, _
    ByVal keyColumn As String, _
    ByVal valueColumn As String, _
    ByVal pairs As Object)

    Dim keys() As String
    Dim i As Long
    Dim lastOldRow As Long
    Dim values() As Variant

    keys = SortedDictionaryKeys(pairs)
    lastOldRow = Application.WorksheetFunction.Max( _
        LastUsedRowInColumn(ws, keyColumn), _
        LastUsedRowInColumn(ws, valueColumn), _
        REFS_FIRST_DATA_ROW)

    If lastOldRow >= REFS_FIRST_DATA_ROW Then
        ws.Range(ws.Cells(REFS_FIRST_DATA_ROW, keyColumn), ws.Cells(lastOldRow, valueColumn)).ClearContents
    End If

    If ArrayCount(keys) = 0 Then Exit Sub

    ReDim values(1 To ArrayCount(keys), 1 To 2)
    For i = 0 To ArrayCount(keys) - 1
        values(i + 1, 1) = keys(i)
        values(i + 1, 2) = pairs(keys(i))
    Next i

    ws.Range( _
        ws.Cells(REFS_FIRST_DATA_ROW, keyColumn), _
        ws.Cells(REFS_FIRST_DATA_ROW + ArrayCount(keys) - 1, valueColumn)).Value2 = values
End Sub

Private Sub WriteNameColumn(ByVal ws As Worksheet, ByVal columnLetter As String, ByVal names As Object)
    Dim keys() As String
    Dim i As Long
    Dim lastOldRow As Long
    Dim values() As Variant

    keys = SortedDictionaryKeys(names)
    lastOldRow = LastUsedRowInColumn(ws, columnLetter)
    If lastOldRow >= REFS_FIRST_DATA_ROW Then
        ws.Range(ws.Cells(REFS_FIRST_DATA_ROW, columnLetter), ws.Cells(lastOldRow, columnLetter)).ClearContents
    End If

    If ArrayCount(keys) = 0 Then Exit Sub

    ReDim values(1 To ArrayCount(keys), 1 To 1)
    For i = 0 To ArrayCount(keys) - 1
        values(i + 1, 1) = keys(i)
    Next i

    ws.Range( _
        ws.Cells(REFS_FIRST_DATA_ROW, columnLetter), _
        ws.Cells(REFS_FIRST_DATA_ROW + ArrayCount(keys) - 1, columnLetter)).Value2 = values
End Sub

Private Sub WriteEquipmentColumns(ByVal ws As Worksheet, ByVal equipment As Object)
    Dim keys() As String
    Dim i As Long
    Dim lastOldRow As Long
    Dim values() As Variant

    keys = SortedDictionaryKeys(equipment)
    lastOldRow = Application.WorksheetFunction.Max( _
        LastUsedRowInColumn(ws, REFS_EQUIPMENT_COLUMN), _
        LastUsedRowInColumn(ws, REFS_EQUIPMENT_OWNERS_COLUMN), _
        REFS_FIRST_DATA_ROW)

    If lastOldRow >= REFS_FIRST_DATA_ROW Then
        ws.Range( _
            ws.Cells(REFS_FIRST_DATA_ROW, REFS_EQUIPMENT_COLUMN), _
            ws.Cells(lastOldRow, REFS_EQUIPMENT_OWNERS_COLUMN)).ClearContents
    End If

    If ArrayCount(keys) = 0 Then Exit Sub

    ReDim values(1 To ArrayCount(keys), 1 To 2)
    For i = 0 To ArrayCount(keys) - 1
        values(i + 1, 1) = keys(i)
        values(i + 1, 2) = equipment(keys(i))
    Next i

    ws.Range( _
        ws.Cells(REFS_FIRST_DATA_ROW, REFS_EQUIPMENT_COLUMN), _
        ws.Cells(REFS_FIRST_DATA_ROW + ArrayCount(keys) - 1, REFS_EQUIPMENT_OWNERS_COLUMN)).Value2 = values
End Sub

Public Function SortedDictionaryKeys(ByVal dict As Object) As String()
    SortedDictionaryKeys = DictionaryKeysToSortedArray(dict)
End Function
