Attribute VB_Name = "Filecheckoff"
Sub TrackCheckoff()
Dim renge As Range
Range("j6").Select
    If Range("j7").Value = Empty Then
        If Range("j6").Value = "Check-Off Loan" Then
            If MsgBox("There are CheckOffs in your report. Do you want to track them?", vbYesNo, "Track checkoff") = vbYes Then
             Track
            End If
        End If
    Else: Range(Selection, Selection.End(xlDown)).Select
        Set renge = Selection
        If WorksheetFunction.CountIf(renge, "Check-Off Loan") > 0 Then
            If MsgBox("There are CheckOffs in your report. Do you want to track them and/or generate letters?", vbYesNo, "Track checkoff") = vbYes Then
             Track
            End If
        End If

    End If
End Sub


Sub Track()

    Sheets("EFT").Copy After:=Sheets("EFT")
    ActiveSheet.Name = "TrackCheckoff"
    Range("j6").Select
    
   Do While ActiveCell.Offset(0, -1) <> ""
   If ActiveCell.Value = "Check-Off Loan" And ActiveCell.Offset(0, 5) = "Completed" Then
   ActiveCell.Offset(1, 0).Select
   Else
   ActiveCell.EntireRow.Delete
   End If
   Loop
   

   
   Columns("A").Delete
   Columns("c").Delete
   Columns("d").Delete
   Columns("g:l").Delete
   Rows("1:4").Delete
   
   Range("h1").Value = "payee"
   Range("i1").Value = "Bankers Cheque Amount"
   Range("k1").Value = "No of Cheques"
   Range("j1").Value = "Limuru Cheque No"
   Columns("h").ColumnWidth = 40
   Columns("i:K").AutoFit
   
    Range("e:e").Cut
    Range("A1").Insert Shift:=xlToRight
     Range("d:d").Cut
    Range("b1").Insert Shift:=xlToRight
        Range("e:e").Cut
    Range("d1").Insert Shift:=xlToRight
           Range("f:f").Cut
    Range("e1").Insert Shift:=xlToRight

ApplyValidationList
ApplyDataValidationWithCustomMessage
End Sub
Sub ApplyValidationList()
    Dim itemList As Variant
    Dim validationRange As Range

    ' List of items as a comma-separated string
    Dim items As String
    'items = "African Capital Limited,Citizen Credit Limited,Cooperative Bank of Kenya Limited,East African Futures Company Limited,Equity Bank Kenya Limited,Faulu Microfinance Bank Limited,Gemini Ventures Limited,Harambee Sacco Limited,Hela Capital Limited,Izwe Loans Kenya Limited,Kenya Women Microfinance Bank,Letshego Kenya Limited,LockBx Limited,Maxxton Enterprises Limited,Platinum Credit Limited,pomelo credit services Limited,Premier Credit Limited,Progressive Credit Limited,Select management services limited,Trustgro Sca Limited"
    items = "African Capital Limited," & _
        "Citizen Credit Limited," & _
        "Cooperative Bank of Kenya Limited," & _
        "East African Futures Company Limited," & _
        "Equity Bank Kenya Limited," & _
        "Faulu Microfinance Bank Limited," & _
        "Gemini Ventures Limited," & _
        "Harambee Sacco Limited," & _
        "Hela Capital Limited," & _
        "Izwe Loans Kenya Limited," & _
        "Kenya Women Microfinance Bank," & _
        "Letshego Kenya Limited," & _
        "LockBx Limited," & _
        "Maxxton Enterprises Limited," & _
        "Platinum Credit Limited," & _
        "pomelo credit services Limited," & _
        "Premier Kenya Limited," & _
        "Progressive Credit Limited," & _
        "Select management services limited," & _
        "Shirika Deposit Taking Sacco," & _
        "Mwito DT Sacco LTD," & _
        "Mwananchi Credit LTD," & "Rafiki Microfinance Bank LTD," & "Kcb Bank Kenya LTD," & "Kenya National Police Dt Sacco," & "Higher Education Loans Board," & "Jafari Credit Limited," & "Family Bank Limited," & "Shirika Deposit Taking Sacco Limited," & "Unifi Credit Limited," & "Micromart Africa Limited," & "Kifedha Limited," & "Hazina Sacco Society LTD," & "Jamii Sacco Society Limited," & "Arthi Sacco Society LTD," & "Trans National Times Sacco Society Limited," & "Magereza Sacco Society LTD," & "Absa Bank Kenya Plc," & _
        "Trustgro Sca Limited"

    ' Split the string into an array based on the comma and sort it alphabetically
    itemList = Split(items, ",")
    SortArray itemList

    ' Apply data validation to Range H1:H50 in the active worksheet
    Set validationRange = ActiveSheet.Range("H1:H50")
    With validationRange.Validation
        .Delete ' Remove any existing validation in the range
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:=Join(itemList, ",")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Choose a Payee"
        .InputMessage = "Choose a Payee from the list. You can also add a new one."
        .ErrorMessage = ""
    End With

    
    CreateRoundedRectanglesWithMacros 'run macro to create buttons
End Sub

Sub SortArray(arr As Variant)
    Dim temp As Variant
    Dim i As Long, j As Long
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

Sub CreateRoundedRectanglesWithMacros()
    Dim lastRow As Long
    Dim buttonTop As Double
    Dim generateShape As shape
    Dim trackShape As shape
    Dim Multiple As shape

    ' Find the last used row in Column C
    lastRow = Cells(Rows.Count, "C").End(xlUp).Row

    ' Calculate the top position for the shapes (just below the data)
    buttonTop = Cells(lastRow, "C").Top + Cells(lastRow, "C").Height + 15

    ' Create the "Generate" shape (rounded rectangle) in Column E
    Set generateShape = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, _
        Left:=Cells(lastRow, "E").Left, Top:=buttonTop, Width:=120, Height:=30)
    With generateShape
        .Name = "GenerateShape"
        .TextFrame.Characters.text = "Generate Reports"
        .Fill.ForeColor.RGB = RGB(0, 0, 0) ' Set the shape fill color to green (RGB: 0, 255, 0)
        .Shadow.Visible = msoTrue ' Add shadow effect
        .Shadow.Type = msoShadow25 ' Set shadow type
        .Shadow.OffsetX = 1 ' Set shadow offset along X-axis
        .Shadow.OffsetY = 1 ' Set shadow offset along Y-axis
        .Shadow.Transparency = 0.5 ' Set shadow transparency (0 to 1)
        
    
    End With

    
    ' Create the "Track" shape (rounded rectangle) in Column H
    Set trackShape = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, _
        Left:=Cells(lastRow, "H").Left, Top:=buttonTop, Width:=120, Height:=30)
    With trackShape
        .Name = "TrackShape"
        .TextFrame.Characters.text = "Track Checkoff"
        .Fill.ForeColor.RGB = RGB(0, 112, 192) ' Set the shape fill color to blue
        .Shadow.Visible = msoTrue ' Add shadow effect
        .Shadow.Type = msoShadow25 ' Set shadow type
        .Shadow.OffsetX = 1 ' Set shadow offset along X-axis
        .Shadow.OffsetY = 1 ' Set shadow offset along Y-axis
        .Shadow.Transparency = 0.5 ' Set shadow transparency (0 to 1)
    End With


        
        
         ' Create the "Multiple" shape (rounded rectangle) in Column i
    Set Multiple = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, _
        Left:=Cells(lastRow, "I").Left, Top:=buttonTop, Width:=150, Height:=30)
    With Multiple
        .Name = "Multiple"
        .TextFrame.Characters.text = "Multiple Cheques"
        .Fill.ForeColor.RGB = RGB(128, 0, 128) ' Set the shape fill color to purple
        .Shadow.Visible = msoTrue ' Add shadow effect
        .Shadow.Type = msoShadow25 ' Set shadow type
        .Shadow.OffsetX = 1 ' Set shadow offset along X-axis
        .Shadow.OffsetY = 1 ' Set shadow offset along Y-axis
        .Shadow.Transparency = 0.5 ' Set shadow transparency (0 to 1)
    End With
        ActiveSheet.Shapes.Range(Array("GenerateShape", "TrackShape", "Multiple")).Select
    With Selection.ShapeRange.ThreeD
        .BevelTopType = msoBevelCircle
        .BevelTopInset = 6
        .BevelTopDepth = 6
    End With
    
    ActiveSheet.Shapes.Range(Array("GenerateShape", "TrackShape", "Multiple")).Select
      With Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 12
    Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = _
        msoAlignCenter
    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
        End With
        
    ' Assign macros to the shapes
    AssignMacrosToShapes
    Range("i2").Select
End Sub

' Macro to be executed when "Generate" shape is clicked
Sub GenerateButton_Click()
AmountInWords
End Sub

' Macro to be executed when "Track" shape is clicked
Sub TrackButton_Click()
ConcatenateDuplicatesAndDelete 'in case of multiple cheques
CheckAndColorCells
End Sub
' Macro to be executed when "Multiple" shape is clicked
Sub MultipleButton_Click()
CheckColumnKBasedOnColumnG
'the following are commented just to show what runs
        'CopyRows
'ActiveSheet.Shapes.Range(Array("Multiple")).Delete
'DuplicateAndRenameSheet

End Sub

Sub AssignMacrosToShapes()
    ' Assign macros to the shapes
    Dim ws As Worksheet
    Dim shape As shape
    Dim macroName As String

    ' Set the worksheet where the shapes were created
    Set ws = ActiveSheet

    ' Assign macros to each shape
    For Each shape In ws.Shapes
        If shape.Name = "GenerateShape" Then
            macroName = "GenerateButton_Click"
        ElseIf shape.Name = "TrackShape" Then
            macroName = "TrackButton_Click"
        ElseIf shape.Name = "Multiple" Then
            macroName = "MultipleButton_Click"
        End If

        On Error Resume Next ' Ignore errors if the shape doesn't have a macro assigned
        shape.OnAction = macroName
        On Error GoTo 0
    Next shape
End Sub
Sub CheckAndColorCells()
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRng As Range
    Dim cell As Range
    Dim emptyCells As Boolean ' Flag to track empty cells

    ' Find the last row and last column in column G with data on the active sheet
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "G").End(xlUp).Row
    lastCol = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column

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
        If MsgBox("Some information is missing. Do you still want to copy the data?", vbYesNo, "Track checkoff") = vbYes Then
            If Range("a3").Value = Empty Then
                Range("a2").Select
                Range(Selection, Selection.End(xlToRight)).Select
                Selection.Copy
            Else
        
            Range("a2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Range(Selection, Selection.End(xlToRight)).Select
            Selection.Copy
            End If
         End If
        
        Else
                If Range("a3").Value = Empty Then
                Range("a2").Select
                Range(Selection, Selection.End(xlToRight)).Select
                Selection.Copy
                Else
                Range("a2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Range(Selection, Selection.End(xlToRight)).Select
                Selection.Copy
                End If
        End If
   
End Sub

Sub AmountInWords()
    Range("L1").Value = "Amount In Words"
    Range("L2").Select
    Do While ActiveCell.Offset(-0, -3) <> Empty
        Application.CutCopyMode = False
        ActiveCell.Formula2R1C1 = "=Jackson(RC[-3])"
        ActiveCell.Offset(1, 0).Select
    Loop
    
    
        Columns("H:H").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    ActiveSheet.Cells.Copy ' Copying data from Sheet2
    
    Dim wb As Workbook
    Dim filePath As String
    filePath = "C:\Users\Jack\OneDrive - foxyevents\templates\ChequesTemplate.xlsx"
    
    On Error Resume Next
    Set wb = Workbooks(Dir(filePath))
    On Error GoTo 0
    
    If wb Is Nothing Then
        Set wb = Workbooks.Open(filePath)
    End If
    
    wb.Sheets(1).Activate ' Activating Sheet1 in the destination workbook
    ActiveSheet.Cells.PasteSpecial Paste:=xlPasteAll ' Pasting data into Sheet1 of the destination workbook
    Application.CutCopyMode = False ' Clear the clipboard after paste
    Columns("a:L").AutoFit
        ActiveSheet.Shapes.Range(Array("WordShape")).Select
    Selection.Delete
    
    Range("h2").Select

GenerateWordDocuments
End Sub

Sub GenerateWordDocuments()
    Dim lastRow As Long
    Dim shapeTop As Double
    Dim WordShape As shape

    ' Find the last used row in Column C
    lastRow = Cells(Rows.Count, "C").End(xlUp).Row

    ' Calculate the top position for the shape (just below the data)
    shapeTop = Cells(lastRow, "C").Top + Cells(lastRow, "C").Height + 15

    ' Create the "word" shape (rounded rectangle) in Column E
    Set WordShape = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, _
        Left:=Cells(lastRow, "E").Left, Top:=shapeTop, Width:=200, Height:=30)
    With WordShape
        .Name = "WordShape"
        .TextFrame.Characters.text = "Export to Microsoft Word"
        .Fill.ForeColor.RGB = RGB(0, 0, 0) ' Set the shape fill color to green (RGB: 0, 255, 0)
        .Shadow.Visible = msoTrue ' Add shadow effect
        .Shadow.Type = msoShadow25 ' Set shadow type
        .Shadow.OffsetX = 1 ' Set shadow offset along X-axis
        .Shadow.OffsetY = 1 ' Set shadow offset along Y-axis
        .Shadow.Transparency = 0.5 ' Set shadow transparency (0 to 1)
        .TextFrame.HorizontalAlignment = xlHAlignCenter ' Align text to the center horizontally
        .TextFrame.VerticalAlignment = xlVAlignCenter ' Align text to the center vertically
    End With

    ' Add 3D effect (bevel) to the shape
    WordShape.Select
    With Selection.ShapeRange.ThreeD
        .BevelTopType = msoBevelCircle
        .BevelTopInset = 6
        .BevelTopDepth = 6
    End With

    ' Set font size for the shape's text
    WordShape.Select
    With Selection.ShapeRange.TextFrame2.TextRange.Font
        .Size = 12
    End With

    ' Assign the "GenerateWordButton_Click" macro to the shape
    On Error Resume Next ' Ignore errors if the shape doesn't have a macro assigned
    WordShape.OnAction = "GenerateWordButton_Click"
    On Error GoTo 0

    ' Move the selection to cell H2 after creating the shape and assigning the macro
    Range("H2").Select
End Sub

' Macro to be executed when "Generate" shape is clicked
Sub GenerateWordButton_Click()
   bankersCheques
End Sub

Sub bankersCheques()
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

 
     
            Range("L2").Select ' copy values in the amount in words columns
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
    
 'this is where we determine if there are duplicates
 Range("d2").Select
Dim ThisRange As Range
Set ThisRange = Range(Cells(2, 4), Cells(Rows.Count, 4).End(xlUp))

Do While ActiveCell.Offset(0, -1) <> Empty
If WorksheetFunction.CountIf(ThisRange, ActiveCell.Value) = 1 Then

Sheets("Template").Range("a16").Value = "Kindly issue us a banker’s cheque of Ksh:" & " " & ActiveCell.Offset(0, 5).Value _
& " " & "(" & ActiveCell.Offset(0, 8).Value & ")" & " " & "in favor of" & " " & ActiveCell.Offset(0, 4).Value & "." & " " & _
"Being loan buyoff for" & " " & ActiveCell.Offset(0, -3).Value & " " & "of ID:" & " " & ActiveCell.Offset(0, 1).Value

ActiveCell.Offset(1, 0).Select

ElseIf WorksheetFunction.CountIf(ThisRange, ActiveCell.Value) = 2 Then
Sheets("Template").Range("a16").Value = "Kindly issue us the following banker's cheques:" & Chr(10) _
& "1. " & "Ksh " & ActiveCell.Offset(0, 5).Value & " " & "(" & ActiveCell.Offset(0, 8).Value & ")" & " " & "in favor of" & " " & ActiveCell.Offset(0, 4).Value & "." & " " & Chr(10) _
& "2. " & "Ksh " & ActiveCell.Offset(1, 5).Value & " " & "(" & ActiveCell.Offset(1, 8).Value & ")" & " " & "in favor of" & " " & ActiveCell.Offset(1, 4).Value & "." & " " & Chr(10) _
& "Being loan buyoff for" & " " & ActiveCell.Offset(0, -3).Value & " " & "of ID:" & " " & ActiveCell.Offset(0, 1).Value
ActiveCell.Offset(2, 0).Select

ElseIf WorksheetFunction.CountIf(ThisRange, ActiveCell.Value) = 3 Then
Sheets("Template").Range("a16").Value = "Kindly issue us the following banker's cheques:" & Chr(10) _
& "1. " & "Ksh " & ActiveCell.Offset(0, 5).Value & " " & "(" & ActiveCell.Offset(0, 8).Value & ")" & " " & "in favor of" & " " & ActiveCell.Offset(0, 4).Value & "." & " " & Chr(10) _
& "2. " & "Ksh " & ActiveCell.Offset(1, 5).Value & " " & "(" & ActiveCell.Offset(1, 8).Value & ")" & " " & "in favor of" & " " & ActiveCell.Offset(1, 4).Value & "." & " " & Chr(10) _
& "3. " & "Ksh " & ActiveCell.Offset(2, 5).Value & " " & "(" & ActiveCell.Offset(2, 8).Value & ")" & " " & "in favor of" & " " & ActiveCell.Offset(2, 4).Value & "." & " " & Chr(10) _
& "Being loan buyoff for" & " " & ActiveCell.Offset(0, -3).Value & " " & "of ID:" & " " & ActiveCell.Offset(0, 1).Value
ActiveCell.Offset(3, 0).Select

ElseIf WorksheetFunction.CountIf(ThisRange, ActiveCell.Value) = 4 Then
Sheets("Template").Range("a16").Value = "Kindly issue us the following banker's cheques:" & Chr(10) _
& "1. " & "Ksh " & ActiveCell.Offset(0, 5).Value & " " & "(" & ActiveCell.Offset(0, 8).Value & ")" & " " & "in favor of" & " " & ActiveCell.Offset(0, 4).Value & "." & " " & Chr(10) _
& "2. " & "Ksh " & ActiveCell.Offset(1, 5).Value & " " & "(" & ActiveCell.Offset(1, 8).Value & ")" & " " & "in favor of" & " " & ActiveCell.Offset(1, 4).Value & "." & " " & Chr(10) _
& "3. " & "Ksh " & ActiveCell.Offset(2, 5).Value & " " & "(" & ActiveCell.Offset(2, 8).Value & ")" & " " & "in favor of" & " " & ActiveCell.Offset(2, 4).Value & "." & " " & Chr(10) _
& "4. " & "Ksh " & ActiveCell.Offset(3, 5).Value & " " & "(" & ActiveCell.Offset(3, 8).Value & ")" & " " & "in favor of" & " " & ActiveCell.Offset(3, 4).Value & "." & " " & Chr(10) _
& "Being loan buyoff for" & " " & ActiveCell.Offset(0, -3).Value & " " & "of ID:" & " " & ActiveCell.Offset(0, 1).Value
ActiveCell.Offset(4, 0).Select
End If


BoldBrackets

        ' Copy the range A1:A23 to Word and set font size to 12
        Sheets("Template").Range("A2:A24").Copy
        wordApp.Selection.Paste
        wordApp.Selection.Font.Size = 12
        
        ' Insert a page break
        wordApp.Selection.InsertBreak Type:=7
        


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
    
     ActiveWorkbook.Close SaveChanges:=False
    wordDoc.SaveAs2 fileName:="C:\Users\Jack\OneDrive - foxyevents\templates\Bankers\BankersCheques_" & Format(Now, "ddmmmyyyy_hhmmss") & ".docx"
    
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
    .MoveEnd Unit:=wdCharacter, Count:=-1 'Move the end of the range back by one character
    If .text = vbCr Then 'Check if the range contains a paragraph mark
        .Delete 'Delete the range
    End If
End With

  

        ' Activate the Word window
    wordApp.Activate
    'Selection.Collapse Direction:=wdCollapseEnd 'Move the selection to the end of the current selection


    'wordDoc.SaveAs2 fileName:="C:\Users\User1\Desktop\Bankers\BankersCheques_" & Format(Now, "ddmmyyyy_hhmmss") & ".docx" 'Save the document with a new name that includes the current date and time
    
    Set wordDoc = Nothing
    Set wordApp = Nothing
End Sub

Sub BoldBrackets()
    'Declare variables
    Dim cell As Range
    Dim text As String
    Dim start As Long
    Dim finish As Long
    
    'Set the cell to a16
    Set cell = Sheets("Template").Range("a16")
    
    'Get the text from the cell
    text = cell.Value
    
    'Loop through the text to find square brackets
    For i = 1 To Len(text)
        'If the character is an opening bracket, mark the start position
        If Mid(text, i, 1) = "[" Then
            start = i
        End If
        
        'If the character is a closing bracket, mark the finish position
        If Mid(text, i, 1) = "]" Then
            finish = i
        End If
        
        'If both start and finish positions are found, apply bold formatting to the text between them
        If start > 0 And finish > 0 Then
            cell.Characters(start, finish - start + 1).Font.Bold = True
            
            'Reset the start and finish positions to zero for the next pair of brackets
            start = 0
            finish = 0
        End If
    Next i
   'Replace the square brackets with empty strings
'cell.Replace "[", "", xlPart
'cell.Replace "]", "", xlPart
 
End Sub
Sub CopyRows()
    Dim activeRow As Range
    Dim numRows As Long
    Dim i As Long
    Dim lastRowG As Long
lastRowG = Cells(Rows.Count, "G").End(xlUp).Row
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
    Range("K2:K" & lastRowG).Interior.Color = vbYellow
End Sub
Sub CheckColumnKBasedOnColumnG()
    Dim lastRowG As Long
    Dim i As Long
    Dim isEmptyRange As Boolean
    
    ' Find the last row in column G
    lastRowG = Cells(Rows.Count, "G").End(xlUp).Row
    
    ' Initialize the flag for empty range
    isEmptyRange = True
    
    ' Loop through each row in column G
    For i = 2 To lastRowG ' Assuming row 1 is for headers
        If Not IsEmpty(Cells(i, "K").Value) Then
            isEmptyRange = False
            Exit For ' Exit the loop if a non-empty cell is found
        End If
    Next i
    
    If isEmptyRange And Range("k2").Interior.Color <> vbYellow Then
        MsgBox "type in the number cheques for each loan account!", vbExclamation, "Empty Cell Warning"
    Else
    Range("K2:K" & lastRowG).Interior.Color = vbYellow
        MultipleReports
    End If
End Sub


 Sub DuplicateAndRenameSheet()
    On Error Resume Next
    ActiveSheet.Shapes.Range(Array("Multiple")).Delete
    Sheets("TrackCheckoff").Copy After:=Sheets(Sheets.Count)
    On Error GoTo 0
    
    On Error Resume Next
    ActiveSheet.Name = "GenerateReports"
    ActiveSheet.Shapes.Range(Array("TrackShape")).Delete
        Sheets("Trackcheckoff").Select
    ActiveSheet.Shapes.Range(Array("GenerateShape")).Select
    Selection.Delete
    On Error GoTo 0
End Sub
Sub ConcatenateDuplicatesAndDelete()
    Dim lastRow As Long
    Dim currentRow As Long
    Dim bankNames As String
    Dim amounts As String
    
    lastRow = Cells(Rows.Count, "D").End(xlUp).Row
    currentRow = 2 ' Start from the second row
    
    Do While currentRow <= lastRow
        Dim currentValue As String
        Dim newRow As Long
        currentValue = Cells(currentRow, "D").Value
        bankNames = Cells(currentRow, "H").Value & "/"
        amounts = Cells(currentRow, "I").Value & "/"
        newRow = currentRow + 1
        
        Do While newRow <= lastRow And Cells(newRow, "D").Value = currentValue
            bankNames = bankNames & Cells(newRow, "H").Value & "/"
            amounts = amounts & Cells(newRow, "I").Value & "/"
            Rows(newRow).Delete
            lastRow = lastRow - 1
        Loop
        
        bankNames = Left(bankNames, Len(bankNames) - 1)
        amounts = Left(amounts, Len(amounts) - 1)
        Cells(currentRow, "H").Value = bankNames
        Cells(currentRow, "I").Value = amounts
        currentRow = currentRow + 1
    Loop
End Sub
Sub ApplyDataValidationWithCustomMessage()
    Dim validationRange As Range
    Dim validationFormula As String
    
    ' Set the validation range
    Set validationRange = Range("K2:K200")
    
    ' Define the data validation formula
    validationFormula = "=AND(ISNUMBER(K2), K2>=1, K2<=4)"
    
    ' Apply data validation
    With validationRange.Validation
        .Delete
        .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=validationFormula
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True ' Enable custom error message
        .ErrorTitle = "Validation Error"
        .ErrorMessage = "You can only have a maximum of 4 cheques per loan account."
    End With
End Sub
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
    If cell.Value <> Empty Then
    cell.Interior.Color = vbWhite
        Else
            cell.Interior.Color = RGB(255, 0, 0) ' Red color
            emptyCells = True ' Set the flag to True if any cell is empty
        End If
    Next cell

    ' Display message if there are empty cells
    If emptyCells = True Then
    MsgBox ("Payee and amount cannot be empty")
    Else
    
DuplicateAndRenameSheet
    End If
End Sub


