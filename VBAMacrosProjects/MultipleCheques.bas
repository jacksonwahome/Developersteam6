Attribute VB_Name = "MultipleCheques"
Sub MultipleBankersCheques()
    Dim rng As Range
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim currentValue As Double
    Dim i As Integer
    
    ' Create a new instance of Word using late binding
    Set wordApp = GetObject(, "Word.Application")

    wordApp.Visible = True
    
    ' Add a new document
    Set wordDoc = wordApp.Documents.Add

 
     
            Range("L2").Select 'copy values in the amount in words columns
            If Range("L3").Value <> Empty Then
        Range(Selection, Selection.End(xlDown)).Select
        End If
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Application.CutCopyMode = False
        
      Range("M2").Select 'convert number to text and preserve the commas
       Do While ActiveCell.Offset(0, -4) <> Empty
    ActiveCell.FormulaR1C1 = "=TEXT(RC[-4],""#,##0"")"
    ActiveCell.Offset(1, 0).Select
    Loop
    Range("m1").Value = Range("i2").Value
    Columns("m:M").Select
Selection.Copy
Columns("I:I").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Application.CutCopyMode = False
  Columns("m:M").Delete
  
    Range("h2").Select 'add square brackets to placeholders
    Do While ActiveCell.Offset(0, -1) <> Empty
    ActiveCell.Value = "[" & ActiveCell.Value & "]"
    ActiveCell.Offset(0, 1) = "[" & ActiveCell.Offset(0, 1) & "]"
    ActiveCell.Offset(0, 4) = "[" & ActiveCell.Offset(0, 4) & "]"
    ActiveCell.Offset(0, -3) = "[" & ActiveCell.Offset(0, -3) & "]"
    ActiveCell.Offset(0, -7) = "[" & ActiveCell.Offset(0, -7) & "]"
    ActiveCell.Offset(1, 0).Select
    Loop
    
    
    
    

 
Range("g2").Select
Do While ActiveCell.Offset(0, -1) <> Empty
Sheets("Template").Range("a16").Value = "Kindly issue us a banker’s cheque of Ksh:" & " " & ActiveCell.Offset(0, 2).Value _
& " " & "(" & ActiveCell.Offset(0, 5).Value & ")" & " " & "in favor of" & " " & ActiveCell.Offset(0, 1).Value & "." & " " & _
"Being loan buyoff for" & " " & ActiveCell.Offset(0, -6).Value & " " & "of ID:" & " " & ActiveCell.Offset(0, -2).Value

BoldBrackets

        ' Copy the range A1:A23 to Word and set font size to 12
        Sheets("Template").Range("A2:A24").Copy
        wordApp.Selection.Paste
        wordApp.Selection.Font.Size = 12
        
        ' Insert a page break
        wordApp.Selection.InsertBreak Type:=7
        

ActiveCell.Offset(1, 0).Select
Loop
' Select the entire document
    wordDoc.Content.Select
    
    ' Set font size to 12
    wordApp.Selection.Font.Size = 12
    
    ' Find and replace all occurrences of square brackets
    wordApp.Selection.Find.ClearFormatting
    wordApp.Selection.Find.text = "["
    wordApp.Selection.Find.Replacement.ClearFormatting
    wordApp.Selection.Find.Replacement.text = ""
    wordApp.Selection.Find.Execute Replace:=2, Forward:=True, Wrap:=wdFindContinue
    
    wordApp.Selection.Find.text = "]"
    wordApp.Selection.Find.Execute Replace:=2, Forward:=True, Wrap:=wdFindContinue
    
' Delete the last page break if present
With wordDoc.Range.Find 'use the Find method on the document range
    .text = "^m" 'search for the page break character
    .Replacement.text = "" 'replace it with nothing
    .Forward = False 'search backwards from the end of the document
    .Wrap = wdFindStop 'stop after one search
    .Execute Replace:=wdReplaceOne 'replace only one occurrence
  End With
  
With wordDoc.Content 'Get the content of the document
    .Collapse Direction:=wdCollapseEnd 'Move the range to the end of the document
    .MoveStart Unit:=wdCharacter, Count:=-1 'Move the start of the range back by one character
    If .text = vbCr Then 'Check if the range contains a paragraph mark
      .Delete 'Delete the range
    End If
  End With


   ActiveWorkbook.Close SaveChanges:=False

        ' Activate the Word window
    wordApp.Activate
    'Selection.Collapse Direction:=wdCollapseEnd 'Move the selection to the end of the current selection


    wordDoc.SaveAs2 fileName:="C:\Users\User1\Desktop\Bankers\BankersCheques_" & Format(Now, "ddmmyyyy_hhmmss") & ".docx" 'Save the document with a new name that includes the current date and time
    Set wordDoc = Nothing
    Set wordApp = Nothing
End Sub
Sub MultipleChequesReport()
 'this is where we determine if there are duplicates
 Range("d2").Select
Dim rng As Range
Set rng = Range(Cells(2, 4), Cells(Rows.Count, 4).End(xlUp))

Do While ActiveCell.Offset(0, -1) <> Empty
If WorksheetFunction.CountIf(rng, ActiveCell.Value) = 1 Then

Sheets("Template").Range("a16").Value = "Kindly issue us a banker’s cheque of Ksh:" & " " & Range("g2").Offset(0, 2).Value _
& " " & "(" & Range("g2").Offset(0, 5).Value & ")" & " " & "in favor of" & " " & Range("g2").Offset(0, 1).Value & "." & " " & _
"Being loan buyoff for" & " " & Range("g2").Offset(0, -6).Value & " " & "of ID:" & " " & Range("g2").Offset(0, -2).Value

ActiveCell.Offset(1, 0).Select

ElseIf WorksheetFunction.CountIf(rng, ActiveCell.Value) = 2 Then
Sheets("Template").Range("a16").Value = "Kindly issue us the following bankers cheques:" & Chr(10) _
& "1. " & "Ksh " & ActiveCell.Offset(0, 5).Value & " " & "(" & ActiveCell.Offset(0, 8).Value & ")" & " " & "in favor of" & " " & ActiveCell.Offset(0, 4).Value & "." & " " & Chr(10) _
& "2. " & "Ksh " & ActiveCell.Offset(1, 5).Value & " " & "(" & ActiveCell.Offset(1, 8).Value & ")" & " " & "in favor of" & " " & ActiveCell.Offset(1, 4).Value & "." & " " & Chr(10) _
& "Being loan buyoff for" & " " & ActiveCell.Offset(0, -3).Value & " " & "of ID:" & " " & ActiveCell.Offset(0, 1).Value
ActiveCell.Offset(2, 0).Select

ElseIf WorksheetFunction.CountIf(rng, ActiveCell.Value) = 3 Then
Sheets("Template").Range("a16").Value = "Kindly issue us the following bankers cheques:" & Chr(10) _
& "1. " & "Ksh " & ActiveCell.Offset(0, 5).Value & " " & "(" & ActiveCell.Offset(0, 8).Value & ")" & " " & "in favor of" & " " & ActiveCell.Offset(0, 4).Value & "." & " " & Chr(10) _
& "2. " & "Ksh " & ActiveCell.Offset(1, 5).Value & " " & "(" & ActiveCell.Offset(1, 8).Value & ")" & " " & "in favor of" & " " & ActiveCell.Offset(1, 4).Value & "." & " " & Chr(10) _
& "3. " & "Ksh " & ActiveCell.Offset(2, 5).Value & " " & "(" & ActiveCell.Offset(2, 8).Value & ")" & " " & "in favor of" & " " & ActiveCell.Offset(2, 4).Value & "." & " " & Chr(10) _
& "Being loan buyoff for" & " " & ActiveCell.Offset(0, -3).Value & " " & "of ID:" & " " & ActiveCell.Offset(0, 1).Value
ActiveCell.Offset(3, 0).Select

ElseIf WorksheetFunction.CountIf(rng, ActiveCell.Value) = 4 Then
Sheets("Template").Range("a16").Value = "Kindly issue us the following bankers cheques:" & Chr(10) _
& "1. " & "Ksh " & ActiveCell.Offset(0, 5).Value & " " & "(" & ActiveCell.Offset(0, 8).Value & ")" & " " & "in favor of" & " " & ActiveCell.Offset(0, 4).Value & "." & " " & Chr(10) _
& "2. " & "Ksh " & ActiveCell.Offset(1, 5).Value & " " & "(" & ActiveCell.Offset(1, 8).Value & ")" & " " & "in favor of" & " " & ActiveCell.Offset(1, 4).Value & "." & " " & Chr(10) _
& "3. " & "Ksh " & ActiveCell.Offset(2, 5).Value & " " & "(" & ActiveCell.Offset(2, 8).Value & ")" & " " & "in favor of" & " " & ActiveCell.Offset(2, 4).Value & "." & " " & Chr(10) _
& "4. " & "Ksh " & ActiveCell.Offset(3, 5).Value & " " & "(" & ActiveCell.Offset(3, 8).Value & ")" & " " & "in favor of" & " " & ActiveCell.Offset(3, 4).Value & "." & " " & Chr(10) _
& "Being loan buyoff for" & " " & ActiveCell.Offset(0, -3).Value & " " & "of ID:" & " " & ActiveCell.Offset(0, 1).Value
ActiveCell.Offset(4, 0).Select
End If
Loop
End Sub

