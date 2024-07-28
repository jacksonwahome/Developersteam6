Attribute VB_Name = "BankersChequesModule"

Sub CopyRows()
    Dim activeRow As Range
    Dim numRows As Long
    Dim i As Long
    
    numRows = ActiveSheet.Cells(Rows.Count, "k").End(xlUp).Row
    
    For i = numRows To 1 Step -1
        If Cells(i, "K").Value = 2 Then
            Rows(i + 1).EntireRow.Insert
            Rows(i).Copy Rows(i + 1)
        ElseIf Cells(i, "K").Value = 3 Then
            Rows(i + 1).Resize(2).EntireRow.Insert
            Rows(i).Copy Rows(i + 1).Resize(2)
        ElseIf Cells(i, "K").Value = 4 Then
            Rows(i + 1).Resize(3).EntireRow.Insert
            Rows(i).Copy Rows(i + 1).Resize(3)
        ElseIf Cells(i, "K").Value = 5 Then
            Rows(i + 1).Resize(4).EntireRow.Insert
            Rows(i).Copy Rows(i + 1).Resize(4)
        End If
    Next i
    Range("K2:K200").Clear
End Sub






   


