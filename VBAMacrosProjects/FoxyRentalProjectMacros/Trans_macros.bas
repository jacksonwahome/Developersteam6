Attribute VB_Name = "Trans_macros"
Option Explicit
Dim TransRow As Long, TransCol As Long, TransID As Long, AttRow As Long, LastAttRow As Long, LastTransRow As Long, LastUnitRow As Long

Dim FileFldr As FileDialog
Dim filePath As String
Sub Trans_GetUniqueUnits()
With Sheet8
LastUnitRow = .Range("A99999").End(xlUp).Row 'Last Unit row
If LastUnitRow < 4 Or Sheet4.Range("b3").Value = Empty Then Exit Sub
.Range("o3").Value = Sheet4.Range("b3").Value 'Property ID
.Range("a3:d" & LastUnitRow).AdvancedFilter xlFilterCopy, .Range("o2:o3"), .Range("q2:r2"), Unique:=True
Sheet4.Range("I5").ClearContents
End With

End Sub
Sub Trans_saveupdate()
With Sheet4
If Range("g3").Value = Empty Or Range("g5").Value = Empty Or Range("i9").Value = Empty Or Range("k9").Value = Empty Then 'missing Trans or tenant name
MsgBox "please make sure to add all required fields( Property, name Type & amount)"
Exit Sub
End If
If Range("b5").Value = Empty Then 'new record
TransRow = Sheet9.Range("A99999").End(xlUp).Row + 1 'first available Transrow
.Range("k3").Value = .Range("b6").Value 'next Trans id
Sheet9.Range("a" & TransRow).Value = .Range("b6").Value 'next Trans ID

Else 'existing
TransRow = Range("b5").Value 'Trans row
End If
 
 For TransCol = 2 To 10
 Sheet9.Cells(TransRow, TransCol).Value = .Range(Sheet9.Cells(1, TransCol).Value).Value 'add or update Trans data
 Next TransCol
 .Shapes("exist group").Visible = msoCTrue 'show existing group
 .Shapes("NewGroup").Visible = msoFalse 'hide new group
 MsgBox "Trans saved/updated"
End With


End Sub

Sub Trans_AddNew()
With Sheet4
    .Range("g3,k3,g5,i5,k5,g7:k7,g5,g9,i9,k9,m4:m11").ClearContents
    .Shapes("exist group").Visible = msoFalse 'hide existing group
    .Shapes("NewGroup").Visible = msoCTrue 'show new group

End With
End Sub

Sub Trans_Load()
Dim FoundProp As Range, FoundUnit As Range
With Sheet4
    If .Range("b5").Value = Empty Then
    MsgBox "please select the correct Transaction from the list"
    Exit Sub
    End If
    TransRow = .Range("b5").Value
    Set FoundProp = Sheet7.Range("Prop_ID").Find(Sheet9.Range("b" & TransRow).Value)
    If Not FoundProp Is Nothing Then .Range("g3").Value = Sheet7.Range("b" & FoundProp.Row).Value 'property name
    Set FoundUnit = Sheet8.Range("Unit_ID").Find(Sheet9.Range("c" & TransRow).Value)
    If Not FoundUnit Is Nothing Then .Range("i5").Value = Sheet8.Range("d" & FoundUnit.Row).Value 'unit name
    
    For TransCol = 4 To 10
    .Range(Sheet9.Cells(1, TransCol).Value).Value = Sheet9.Cells(TransRow, TransCol).Value 'add or update Trans data
    Next TransCol
    .Shapes("exist group").Visible = msoCTrue 'show existing group
    .Shapes("NewGroup").Visible = msoFalse 'hide new group
    Trans_AttachLoad
End With

End Sub
Sub Trans_LoadRecentTrans()
Sheet4.Range("c4:d11").ClearContents
With Sheet9
LastTransRow = .Range("A99999").End(xlUp).Row 'last Trans Row
If LastTransRow < 4 Then Exit Sub
On Error Resume Next
.Names("Criteria").Delete ' get rid of the criteria we do not need it
On Error GoTo 0
.Range("a3:f" & LastTransRow).AdvancedFilter xlFilterCopy, , .Range("n2:p2"), Unique:=True
With .Sort 'sort by date
.SortFields.Clear
.SortFields.Add Key:=Sheet9.Range("n3"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
.SetRange Sheet9.Range("N3:p" & LastTransRow - 1)
.Apply
End With
Sheet4.Range("c4:d" & LastTransRow).Value = .Range("o3:p" & LastTransRow - 1).Value 'bring in Trans details

End With
End Sub


Sub Trans_Delete()
    If MsgBox("Do you want to delete this Transaction?", vbYesNo, "Delete Trans!") = vbNo Then Exit Sub
    If Sheet4.Range("b5").Value = Empty Then
    MsgBox "please select a Transaction to delete"
    Exit Sub
    End If
    TransRow = Sheet4.Range("b5").Value ' Trans row
    Sheet9.Range(TransRow & ":" & TransRow).EntireRow.Delete
    Trans_AddNew
End Sub
Sub Trans_CancelNew()
If Sheet4.Range("d4").Value <> Empty Then Sheet4.Range("d4").Select
End Sub
Sub Trans_AddAttach()
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
        .Range("b" & AttRow).Value = Sheet4.Range("b3").Value 'Property id
        .Range("c" & AttRow).Value = Sheet4.Range("b4").Value 'Unit id
        .Range("d" & AttRow).Value = Sheet4.Range("k3").Value 'Trans id
        .Range("e" & AttRow).Value = filePath 'filepath
Noselection:
    End With
    Trans_AttachLoad
End Sub

Sub Trans_AttachLoad()
Sheet4.Range("m4:m11").ClearContents
    With Sheet11 'attachmens database
    .Range("m3").Value = Sheet4.Range("k3").Value 'Trans id (Criteria)
    LastAttRow = .Range("a99999").End(xlUp).Row + 1 'last attach row
    .Range("a2:e" & LastAttRow).AdvancedFilter xlFilterCopy, .Range("m2:m3"), .Range("o2"), Unique:=True
    Sheet4.Range("M4:M11").Value = .Range("o3:o10").Value 'bring over file attachments
            
    End With

End Sub


