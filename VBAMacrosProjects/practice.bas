Attribute VB_Name = "practice"
 Sub MultipleReports()
Dim lastRowG As Long
lastRowG = Cells(Rows.Count, "G").End(xlUp).Row
    If isEmptyRange And Range("k2").Interior.Color <> vbYellow Then
        MsgBox "type in the number cheques for each loan account!", vbExclamation, "Empty Cell Warning"
    Else
    Range("K2:K" & lastRowG).Interior.Color = vbYellow
    Filecheckoff.CopyRows
    End If
If Range("K2").Interior.Color = vbYellow Then
CheckPayeeAndAmount

End If
End Sub

Sub CheckPayeeAndAmount()

    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRng As Range
    Dim cell As Range
    Dim emptyCells As Boolean ' Flag to track empty cells

    ' Find the last row and last column in column G with data on the active sheet
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "G").End(xlUp).Row
    lastCol = 9

    ' Define the range starting from H2
    Set dataRng = ActiveSheet.Range("H2", ActiveSheet.Cells(lastRow, lastCol))

    ' Add borders to the data range
    dataRng.BorderAround ColorIndex:=1, Weight:=xlThin
    dataRng.Borders.LineStyle = xlContinuous

    ' Initialize the emptyCells flag to False
    emptyCells = False

    ' Loop through each cell in the range and check for blanks
    For Each cell In dataRng
        If cell.Value = "" Then
            ' Color the cell red
            cell.Interior.Color = RGB(255, 0, 0) ' Red color
            emptyCells = True ' Set the flag to True if any cell is empty
        End If
    Next cell

    ' Display message if there are empty cells
    If emptyCells = True Then
    MsgBox ("Payee and amount cannot be empty")
    Else
    
Filecheckoff.DuplicateAndRenameSheet
    End If
End Sub


