Attribute VB_Name = "Alerts_macros"
Option Explicit
Dim AlertRow As Long, AlertCol As Long, AlertID As Long, AttRow As Long, LastAttRow As Long, LastAlertRow As Long, LastUnitRow As Long


Sub Alert_saveupdate()
With Sheet5
If Range("g5").Value = Empty Or Range("i5").Value = Empty Or Range("k5").Value = Empty Then 'missing Alert name, date time
MsgBox "please make sure to add all required fields( Name, Date and time)"
Exit Sub
End If
If Range("b4").Value = Empty Then 'new record
AlertRow = Sheet10.Range("A99999").End(xlUp).Row + 1 'first available Alertrow
.Range("k3").Value = .Range("b5").Value 'next Alert id
Sheet10.Range("a" & AlertRow).Value = .Range("b5").Value 'next Alert ID

Else 'existing
AlertRow = .Range("b4").Value 'Alert row
End If
 
 For AlertCol = 2 To 5
 Sheet10.Cells(AlertRow, AlertCol).Value = .Range(Sheet10.Cells(1, AlertCol).Value).Value 'add or update Alert data
 Next AlertCol
 Sheet10.Range("f" & AlertRow).Value = .Range("i5").Value + .Range("k5").Value ' date and time
 Sheet10.Range("h" & AlertRow).Value = "=Row()"
 .Shapes("exist group").Visible = msoCTrue 'show existing group
 .Shapes("NewGroup").Visible = msoFalse 'hide new group
 MsgBox "Alert saved/updated"
 Alert_LoadRecentAlerts
End With


End Sub

Sub Alert_AddNew()
With Sheet5
    .Range("k3,g5,i5,k5,g7:k7").ClearContents
    .Shapes("exist group").Visible = msoFalse 'hide existing group
    .Shapes("NewGroup").Visible = msoCTrue 'show new group

End With
End Sub

Sub Alert_Load()
With Sheet5
    If .Range("b4").Value = Empty Then
    MsgBox "please select the correct Alert from the list"
    Exit Sub
    End If
    AlertRow = .Range("b4").Value
    
    For AlertCol = 2 To 5
    .Range(Sheet10.Cells(1, AlertCol).Value).Value = Sheet10.Cells(AlertRow, AlertCol).Value 'add or update Alert data
    Next AlertCol
    .Shapes("exist group").Visible = msoCTrue 'show existing group
    .Shapes("NewGroup").Visible = msoFalse 'hide new group
    
End With

End Sub
Sub Alert_LoadRecentAlerts()
Sheet5.Range("c4:d11").ClearContents
With Sheet10
LastAlertRow = .Range("A99999").End(xlUp).Row 'last Alert Row
If LastAlertRow < 4 Then Exit Sub

.Range("a3:h" & LastAlertRow).AdvancedFilter xlFilterCopy, .Range("L2:L3"), .Range("n2:u2"), Unique:=True
LastAlertRow = .Range("N99999").End(xlUp).Row 'last Alert Row
If LastAlertRow < 3 Then Exit Sub
With .Sort 'sort by date
.SortFields.Clear
.SortFields.Add Key:=Sheet10.Range("p3"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
.SetRange Sheet10.Range("N3:u" & LastAlertRow)
.Apply
End With
Sheet5.Range("c4:d" & LastAlertRow + 1).Value = .Range("n3:o" & LastAlertRow).Value 'bring in Alert details

End With
End Sub


Sub Alert_Delete()
    If MsgBox("Do you want to delete this Alert?", vbYesNo, "Delete Alert!") = vbNo Then Exit Sub
    If Sheet5.Range("b4").Value = Empty Then
    MsgBox "please select a Alert to delete"
    Exit Sub
    End If
    AlertRow = Sheet5.Range("b4").Value ' Alert row
    Sheet10.Range(AlertRow & ":" & AlertRow).EntireRow.Delete
    Alert_LoadRecentAlerts
    Alert_AddNew
End Sub
Sub Alert_CancelNew()
If Sheet5.Range("d4").Value <> Empty Then Sheet5.Range("d4").Select
End Sub






