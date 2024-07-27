Attribute VB_Name = "Prop_macros"
Option Explicit
Dim PropRow As Long, PropCol As Long, PropID As Long, AttRow As Long, LastAttRow As Long

Dim FileFldr As FileDialog
Dim filePath As String

Sub property_saveupdate()
With Sheet2
If Range("e5").Value = Empty Then 'missing property name
MsgBox "please make sure to add a property name"
Exit Sub
End If
If Range("b4").Value = Empty Then 'new record
PropRow = Sheet7.Range("A99999").End(xlUp).Row + 1 'first available Proprow
.Range("i3").Value = .Range("b5").Value 'next prop id
Sheet7.Range("a" & PropRow).Value = .Range("b5").Value 'next prop ID
.Range("e3").Value = .Range("e5").Value 'prop name

Else 'existing
PropRow = Range("b4").Value 'property row
End If
 
 For PropCol = 2 To 11
 Sheet7.Cells(PropRow, PropCol).Value = .Range(Sheet7.Cells(1, PropCol).Value).Value 'add or update prop data
 Next PropCol
 .Shapes("exist group").Visible = msoCTrue 'show existing group
 .Shapes("NewGroup").Visible = msoFalse 'hide new group
 MsgBox "property saved/updated"
End With


End Sub

Sub Property_AddNew()
With Sheet2
    .Range("e3,i3,g3,e5,i5,e7:i7,e9,g5,g9,i9,e11,g11,i11,k4:k11").ClearContents
    .Shapes("exist group").Visible = msoFalse 'hide existing group
    .Shapes("NewGroup").Visible = msoCTrue 'show new group

End With
End Sub

Sub Property_load()
With Sheet2
    If .Range("b3").Value = Empty And .Range("e3") <> Empty Then
    MsgBox "please select the correct property from the list"
    Exit Sub
    End If
    PropRow = .Range("b3").Value
    For PropCol = 1 To 11
    .Range(Sheet7.Cells(1, PropCol).Value).Value = Sheet7.Cells(PropRow, PropCol).Value 'add or update prop data
    Next PropCol
    .Shapes("exist group").Visible = msoCTrue 'show existing group
    .Shapes("NewGroup").Visible = msoFalse 'hide new group
    Property_AttachLoad
End With

End Sub

Sub Property_CancelNew()
    If Sheet7.Range("b4") <> Empty Then Sheet2.Range("e3").Value = Sheet7.Range("b4").Value 'add in the first property
End Sub

Sub Property_Delete()
    If MsgBox("Do you want to delete this property?", vbYesNo, "Delete property!") = vbNo Then Exit Sub
    If Sheet2.Range("b3").Value = Empty Then
    MsgBox "please select a property to delete"
    Exit Sub
    End If
    PropRow = Sheet2.Range("b3").Value ' property row
    Sheet7.Range(PropRow & ":" & PropRow).EntireRow.Delete
    Property_AddNew
End Sub

Sub Property_AddAttach()
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
        .Range("b" & AttRow).Value = Sheet2.Range("i3").Value 'prop id
        .Range("e" & AttRow).Value = filePath 'filepath
Noselection:
    End With
    Property_AttachLoad
End Sub

Sub Property_AttachLoad()
Sheet2.Range("k4:k11").ClearContents
    With Sheet11
    .Range("k3").Value = Sheet2.Range("i3").Value 'property id (Criteria)
    LastAttRow = .Range("a99999").End(xlUp).Row + 1 'last attach row
    .Range("a2:e" & LastAttRow).AdvancedFilter xlFilterCopy, .Range("k2:k3"), .Range("o2"), Unique:=True
    Sheet2.Range("k4:k11").Value = .Range("o3:o10").Value 'bring over file attachments
            
    End With

End Sub
