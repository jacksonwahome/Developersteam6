Attribute VB_Name = "Unit_macros"
Option Explicit
Dim UnitRow As Long, UnitCol As Long, UnitID As Long, AttRow As Long, LastAttRow As Long, LastUnitRow As Long

Dim FileFldr As FileDialog
Dim filePath As String

Sub Unit_saveupdate()
With Sheet3
If Range("g5").Value = Empty Or Range("i5").Value = Empty Then 'missing Unit or tenant name
MsgBox "please make sure to add a Unit & Tenant name"
Exit Sub
End If
If Range("b4").Value = Empty Then 'new record
UnitRow = Sheet8.Range("A99999").End(xlUp).Row + 1 'first available Unitrow
.Range("k3").Value = .Range("b5").Value 'next Unit id
Sheet8.Range("a" & UnitRow).Value = .Range("b5").Value 'next Unit ID

Else 'existing
UnitRow = Range("b4").Value 'Unit row
End If
 
 For UnitCol = 2 To 12
 Sheet8.Cells(UnitRow, UnitCol).Value = .Range(Sheet8.Cells(1, UnitCol).Value).Value 'add or update Unit data
 Next UnitCol
 .Shapes("exist group").Visible = msoCTrue 'show existing group
 .Shapes("NewGrp").Visible = msoFalse 'hide new group
 MsgBox "Unit saved/updated"
End With


End Sub

Sub Unit_AddNew()
With Sheet3
    .Range("d4:d11,i3,g3,e5,i5,k3,k5,g7:k7,g5,g9,i9,k9,g11,i11,m4:m11,k11").ClearContents
    .Shapes("exist group").Visible = msoFalse 'hide existing group
    .Shapes("NewGrp").Visible = msoCTrue 'show new group

End With
End Sub

Sub Unit_Load()
With Sheet3
    If .Range("b4").Value = Empty Then
    MsgBox "please select the correct Unit from the list"
    Exit Sub
    End If
    UnitRow = .Range("b4").Value
    For UnitCol = 3 To 12 ' u don't wonna load prop id and unit id sinc those are formulas
    .Range(Sheet8.Cells(1, UnitCol).Value).Value = Sheet8.Cells(UnitRow, UnitCol).Value 'add or update Unit data
    Next UnitCol
    .Shapes("exist group").Visible = msoCTrue 'show existing group
    .Shapes("NewGrp").Visible = msoFalse 'hide new group
    Unit_AttachLoad
End With

End Sub
Sub UnitLoadUnits()
Sheet3.Range("c4:d11").ClearContents
With Sheet8
LastUnitRow = .Range("A99999").End(xlUp).Row 'last Unit Row
If LastUnitRow < 4 Then Exit Sub
.Range("o3").Value = Sheet3.Range("b3").Value 'property ID
.Range("a3:d" & LastUnitRow).AdvancedFilter xlFilterCopy, .Range("o2:o3"), .Range("q2:r2"), Unique:=True
Sheet3.Range("c4:d11") = .Range("q3:r10").Value 'bring in unit details
End With
End Sub


Sub Unit_Delete()
    If MsgBox("Do you want to delete this Unit?", vbYesNo, "Delete Unit!") = vbNo Then Exit Sub
    If Sheet3.Range("b4").Value = Empty Then
    MsgBox "please select a Unit to delete"
    Exit Sub
    End If
    UnitRow = Sheet3.Range("b4").Value ' Unit row
    Sheet8.Range(UnitRow & ":" & UnitRow).EntireRow.Delete
    Unit_AddNew
End Sub

Sub Unit_AddAttach()
    With Sheet11
    Set FileFldr = Application.FileDialog(msoFileDialogFilePicker)
        With FileFldr
        .Title = "Select A file to attach"
        .Filters.Add "All Files", "*.*", 1
        If .Show <> -1 Then GoTo Noselection
        filePath = .SelectedItems(1)
        End With
        AttRow = .Range("A99999").End(xlUp).Row + 1 'first available attach row
        .Range("a" & AttRow).Value = Sheet2.Range("b6").Value 'attch ID
        .Range("b" & AttRow).Value = Sheet3.Range("b3").Value 'Property id
        .Range("c" & AttRow).Value = Sheet3.Range("k3").Value 'Unit id
        .Range("e" & AttRow).Value = filePath 'filepath
Noselection:
    End With
    Unit_AttachLoad
End Sub

Sub Unit_AttachLoad()
Sheet3.Range("m4:m11").ClearContents
    With Sheet11
    .Range("L3").Value = Sheet3.Range("i3").Value 'Unit id (Criteria)
    LastAttRow = .Range("a99999").End(xlUp).Row + 1 'last attach row
    .Range("a2:e" & LastAttRow).AdvancedFilter xlFilterCopy, .Range("L2:L3"), .Range("o2"), Unique:=True
    Sheet3.Range("M4:M11").Value = .Range("o3:o10").Value 'bring over file attachments
            
    End With

End Sub

