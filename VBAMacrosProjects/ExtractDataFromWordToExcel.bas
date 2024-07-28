Attribute VB_Name = "loopfiles"

Sub table1and3()
Dim y As Long
Dim Wd As New Word.Application
Dim Doc As Word.Document
Dim Sh As Worksheet
Dim Filepath As String

For y = 2 To 100
Set Wd = New Word.Application
Filepath = Sheets("files").Cells(y, 1).Value
Wd.Visible = True
Set Doc = Wd.Documents.Open(Filepath)
Set tbls = Doc.Tables
Set Sh = ActiveSheet
 tbls(1).Range.Copy
Sheets("rawdata").Activate
Sheets("rawdata").Range("a1").Select
ActiveSheet.Paste

Selection.UnMerge
Selection.WrapText = False
Selection.Columns.AutoFit

Sheets("data").Select
Dim x As Integer
Dim searchValue As String
Dim searchRange As Range
Dim foundCell As Range
Dim copyRange As Range
Dim copyValue As Variant
Dim SearchValue2 As String
Dim FoundCell2 As Range
Dim searchRange2 As Range
Dim result As String 'to be used from table 2 and so on
Dim cell As Range 'to be used from table 2 and so on
x = 2
For x = 2 To 19
Cells(2, x).Select
    If x = 3 Then
    ActiveCell.FormulaR1C1 = "=VLOOKUP(R[-1]C,rawdata!R1C1:R13C2,2,FALSE)"
    ActiveCell.Copy
    ActiveCell.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Else
searchValue = ActiveCell.Offset(-1, 0).Value ' Set the value to search for
        If x = 4 Then
        Set searchRange = Sheets("rawdata").Range("d1:E10") ' Set the range to search within
        ElseIf (x >= 8 And x <= 19) Then
        Set searchRange = Sheets("rawdata").Range("b4:b20")
    
        Else
        Set searchRange = Sheets("rawdata").Range("A1:A10")
        End If
        
    Set foundCell = searchRange.Find(searchValue, LookIn:=xlValues, LookAt:=xlWhole) ' Search for the value in the range
    If Not foundCell Is Nothing Then ' Check if a match was found
        Set copyRange = foundCell.Offset(0, 1).Resize(1, 5) ' Set the range to copy
        Set copyValue = copyRange.Find("*", LookIn:=xlValues, LookAt:=xlWhole) ' Find the first non-empty cell in the copy range
        If Not copyValue Is Nothing Then ' Check if a value was found in the copy range
            'copyValue.Copy ' Copy the value to the clipboard
            Sheets("data").Cells(2, x).Value = copyValue
        Else
            Sheets("data").Cells(2, x).Value = "blank"
        End If
    Else
        Sheets("data").Cells(2, x).Value = "missing"
    End If
    End If
    
Next x


For x = 21 To 26
SearchValue2 = "Security Offered"
Set searchRange2 = Sheets("rawdata").Range("a13:a17")
Set FoundCell2 = searchRange2.Find(SearchValue2, LookIn:=xlValues, LookAt:=xlWhole) ' Search for the value in the range
If Not FoundCell2 Is Nothing Then ' Check if a match was found

If FoundCell2.Offset(3, 1).Value <> Empty And FoundCell2.Offset(3, 2) <> Empty Then
Cells(2, x).Select
searchValue = ActiveCell.Offset(-1, 0).Value ' Set the value to search for
Set searchRange = Sheets("rawdata").Range("a13:j17") ' Set the range to search within
Set foundCell = searchRange.Find(searchValue, LookIn:=xlValues, LookAt:=xlWhole) ' Search for the value in the range
If Not foundCell Is Nothing Then ' Check if a match was found
Cells(2, x).Value = foundCell.Offset(1, 0).Value
Cells(2, x + 14).Value = foundCell.Offset(2, 0).Value
Cells(2, x + 20).Value = foundCell.Offset(3, 0).Value
End If

ElseIf FoundCell2.Offset(2, 1).Value <> Empty And FoundCell2.Offset(2, 2) <> Empty Then
Cells(2, x).Select
searchValue = ActiveCell.Offset(-1, 0).Value ' Set the value to search for
Set searchRange = Sheets("rawdata").Range("a13:j17") ' Set the range to search within
Set foundCell = searchRange.Find(searchValue, LookIn:=xlValues, LookAt:=xlWhole) ' Search for the value in the range
If Not foundCell Is Nothing Then ' Check if a match was found
Cells(2, x).Value = foundCell.Offset(1, 0).Value
Cells(2, x + 14).Value = foundCell.Offset(2, 0).Value
End If



Else
Cells(2, x).Select
searchValue = ActiveCell.Offset(-1, 0).Value ' Set the value to search for
Set searchRange = Sheets("rawdata").Range("a13:j17") ' Set the range to search within
Set foundCell = searchRange.Find(searchValue, LookIn:=xlValues, LookAt:=xlWhole) ' Search for the value in the range
If Not foundCell Is Nothing Then ' Check if a match was found
Cells(2, x).Value = foundCell.Offset(1, 0).Value
End If
End If
End If
Next x

For x = 27 To 27
    searchValue = "Comment On Security" ' Set the value to search for
    Set searchRange = Sheets("rawdata").Range("A10:A35") ' Set the range to search within
    Set foundCell = searchRange.Find(searchValue, LookIn:=xlValues, LookAt:=xlWhole) ' Search for the value in the range
    If Not foundCell Is Nothing Then ' Check if a match was found
        ' Perform TEXTJOIN with the three cells below the found value, separated by line breaks
        For Each cell In foundCell.Offset(0, 1).Resize(3)
            result = result & cell.Value & vbCrLf
        Next cell
        result = Left(result, Len(result) - 2) ' Remove the extra line break at the end
        Cells(2, x) = result ' Put the result in the active cell
    End If
Next x

For x = 28 To 28 'client historical information
        Sheets("rawdata").Columns("a:j").Delete
        tbls(2).Range.Copy
        Sheets("rawdata").Activate
        Sheets("rawdata").Range("a1").Select
        ActiveSheet.Paste
        Selection.UnMerge
        Selection.WrapText = False
        Selection.Columns.AutoFit
        result = ""
        ' Perform TEXTJOIN with the three cells below the found value, separated by line breaks
        For Each cell In Sheets("rawdata").Range("a1").Resize(7)
            result = result & cell.Value & vbCrLf
        Next cell
        result = Left(result, Len(result) - 2) ' Remove the extra line break at the end
        Sheets("data").Cells(2, x) = result ' Put the result in the active cell
   
Next x

For x = 29 To 29 'loan purpose
        Sheets("rawdata").Columns("a:j").Delete
        tbls(3).Range.Copy
        Sheets("rawdata").Activate
        Sheets("rawdata").Range("a1").Select
        ActiveSheet.Paste
        Selection.UnMerge
        Selection.WrapText = False
        Selection.Columns.AutoFit
        result = ""
        ' Perform TEXTJOIN with the three cells below the found value, separated by line breaks
        For Each cell In Sheets("rawdata").Range("a1").Resize(7)
            result = result & cell.Value & vbCrLf
        Next cell
        result = Left(result, Len(result) - 2) ' Remove the extra line break at the end
        Sheets("data").Cells(2, x) = result ' Put the result in the active cell
   
Next x

For x = 31 To 31 'crb report
    
        Sheets("rawdata").Columns("a:j").Delete
        tbls(5).Range.Copy
        Sheets("rawdata").Activate
        Sheets("rawdata").Range("a1").Select
        ActiveSheet.Paste
        Selection.UnMerge
        Selection.WrapText = False
        Selection.Columns.AutoFit
        result = ""
        ' Perform TEXTJOIN with the three cells below the found value, separated by line breaks
        For Each cell In Sheets("rawdata").Range("a1").Resize(5, 1)
            result = result & cell.Value & vbCrLf
        Next cell
        result = Left(result, Len(result) - 2) ' Remove the extra line break at the end
        Sheets("data").Cells(2, x) = result ' Put the result in the active cell
   
Next x

For x = 34 To 34
        Sheets("rawdata").Columns("a:j").Delete
        On Error Resume Next ' Enable error handling
        tbls(8).Range.Copy ' Attempt to copy the range
        If Err.Number <> 0 Then ' Check if an error occurred
            Err.Clear ' Clear the error object
            ' Alternative code to execute if an error occurred
            Sheets("data").Cells(2, 31).ClearContents
            tbls(7).Range.Copy ' Copy the range from tbls(7) instead
    
        End If
        
        ' Rest of your code...
        Sheets("rawdata").Activate
        Sheets("rawdata").Range("A1").Select
        ActiveSheet.Paste
        Selection.UnMerge
        Selection.WrapText = False
        Selection.Columns.AutoFit
        result = ""
        ' Perform TEXTJOIN with the three cells below the found value, separated by line breaks
        
Dim preparedByRange As Range
Set preparedByRange = Sheets("rawdata").Range("A1:A7")
For Each cell In preparedByRange
    If cell.Value = "Prepared By" Then
        result = Application.WorksheetFunction.TextJoin(vbCrLf, True, preparedByRange.Resize(cell.Row - 1).Value)
        Exit For
    End If
Next cell

        Sheets("data").Cells(2, x) = result ' Put the result in the active cell
        Sheets("rawdata").Columns("a:j").Delete
        Sheets("data").Cells(2, 1).Value = Sheets("files").Cells(y, 1).Value
        Sheets("data").Rows(2).Insert Shift:=xlDown
        Wd.Quit
Next x
Next y
End Sub








