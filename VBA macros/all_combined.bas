Attribute VB_Name = "all_combined"
Sub AllinOneMacro()
Attribute AllinOneMacro.VB_ProcData.VB_Invoke_Func = "A\n14"

Range("D6").Select

If ActiveCell.Offset(0, -1) = "M-Pesa" And Range("a6").Value <> 1 And Range("A1").Value <> "Refunds" Then
    Mpesa
ElseIf ActiveCell.Offset(0, -1).Value = "EFT" And Range("a6").Value <> 1 Then
    Eft
ElseIf ActiveCell.Offset(0, -1).Value = "Cheque" And Range("a6").Value <> 1 Then
    Cheques
ElseIf Range("a1").Value = "Bulk Plan ID" Then
    MpesaValidationCleanup
ElseIf Range("H5").Value = "Customer" And Range("a5").Value = "NO" Then
    MpesaValidation
ElseIf Range("a1").Value = "Refunds" And Range("m5").Value = "From Mobile" Then
    refundsCleanup
ElseIf Range("A6").Value = 1 And Range("C5").Value = "Group Name" Then
    RefundsValidation
ElseIf Range("d5").Value = "Loan Officer" Then
    ReadyToDisburse
Else: MsgBox "oops! unable run macros on this Sheet", vbCritical

End If
End Sub

Sub Mpesa()
    Columns("A:U").Select
    Selection.Columns.AutoFit
    Range("A5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.ColumnWidth = 5
    Columns("B:R").AutoFit
    'Columns("B:B").Delete
    Columns("e:e").Select
    Selection.Delete
    Columns("q:q").Delete
    Columns("R:R").Delete
    Range("A5").Value = "NO"
    Range("a5").Font.Bold = True
    Cells(1, 1).Select

'this do while loop will input numbers for both supafast and others

Dim i As Integer
Dim supa As Integer
Dim a As Integer
i = 6
a = 0
supa = 1



Do Until Cells(i, 2).Value = ""
If Cells(i, 11) <> "Supafast Loan" Then
    Cells(i, 1) = a + 1
    a = a + 1
    i = i + 1
Else
    Cells(i, 1) = supa
    i = i + 1
    supa = supa + 1
End If
Loop
'This do while loop will add 254 to phone numbers

Range("R6").Select
Do Until ActiveCell.Offset(0, -1) = ""
    ActiveCell.FormulaR1C1 = "=concatenate(254,RC[-8])"
    ActiveCell.Offset(1, 0).Select
Loop

'This will copy phone numbers to column j


Range("R6").Select
If Range("R7").Value = "" Then
        Selection.Copy
        Range("J6").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Columns("J:J").AutoFit
        Application.CutCopyMode = False
Else
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        Range("J6").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Selection.Columns.AutoFit
       Application.CutCopyMode = False
End If
        Columns("R:R").Delete

 'this will autosum the loan amounts

Range("M6").Select
If Range("m7").Value = "" Then
        ActiveCell.Offset(1, 0).Select
        ActiveCell.FormulaR1C1 = "=SUM(R6C12:R[-1]C)"
        ActiveCell.Offset(0, 0).Select
        Selection.Font.Bold = True
        Selection.Font.Size = 14
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
Else
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Select
        ActiveCell.FormulaR1C1 = "=SUM(R6C12:R[-1]C)"
        ActiveCell.Offset(0, 0).Select
        Selection.Font.Bold = True
        Selection.Font.Size = 14
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
End If

    'this will add the signatory string.

   Range("B5").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(2, 0).Select
        Selection.Value = "SIGNATORY A"
        Selection.Font.Bold = True
        Selection.Font.Size = 12
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
        ActiveCell.Offset(0, 9).Select
        ActiveCell.Value = "SIGNATORY B"
        Selection.Font.Bold = True
        Selection.Font.Size = 12
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With

      'this macro will add borders

        Range("A5").Select
        Range(Selection, Selection.End(xlDown)).Select
        Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Borders
        .LineStyle = xlcontinous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

'reject or accept based on loan product

Range("K6").Select
Do Until ActiveCell.Offset(0, -5) = ""
    If ActiveCell.Value = "Group - Agribusiness" Or ActiveCell.Value = "Group - Agri Asset Acquisition" Or ActiveCell.Value = "Group - Animal Farming" Or ActiveCell.Value = "Check-Off Loan" Or ActiveCell.Value = "Individual Loan" Or ActiveCell.Value = "Group - Non Agribusiness" Then
        ActiveCell.Offset(1, 0).Select
    ElseIf ActiveCell.Value = "Supafast Loan" Then
        ActiveCell.Interior.Color = vbYellow
        ActiveCell.Offset(1, 0).Select
    Else: ActiveCell.Interior.Color = vbRed
        ActiveCell.EntireRow.Font.Strikethrough = True
        ActiveCell.Offset(0, 7).Value = "reject"
        ActiveCell.Offset(0, 7).Font.Strikethrough = False
        ActiveCell.Offset(1, 0).Select
    End If
Loop
        Columns("N:N").Delete
        Columns("P:P").AutoFit
        Columns("q:q").Select
        Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        ActiveSheet.Name = "validation"
    
'This procedure will sort the sheet by product
        Range("A5").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.AutoFilter
        ActiveWorkbook.Worksheets("validation").AutoFilter.Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("validation").AutoFilter.Sort.SortFields.Add(Range( _
        "K5"), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB( _
        255, 255, 0)
    With ActiveWorkbook.Worksheets("validation").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


'creating 1 additional worksheets

        Cells.Select
        Selection.Copy
        Sheets.Add.Name = "Disbursals"
        Worksheets("Disbursals").Select
        ActiveSheet.Paste
        Range("a1").Select

'copy the worksheets for validation purposes
On Error Resume Next
Sheets(ActiveSheet.index + 1).Activate
    If Err.Number <> 0 Then Sheets(1).Activate
        Range("a1").Select
        Application.CutCopyMode = False
        Range("j6").Select
      If Range("J7") = "" Then
        Selection.Copy
      Else
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    End If

End Sub

'Eft module
Sub Eft()
        Columns("A:U").Select
        Selection.Columns.AutoFit
        Columns("A:A").ColumnWidth = 15
        Columns("A:A").Insert
        Columns("A:A").ColumnWidth = 5
        Columns("t:t").Delete
        Columns("R:R").Delete
        Columns("k:k").Delete
        Range("A5").Value = "NO"
        Range("a5").Font.Bold = True
        Cells(1, 1).Select

'this do while loop will input column a data

Dim x As Integer
Dim y As Integer
x = 6
y = 0

Do Until Cells(x, 2).Value = ""
        Cells(x, 1) = y + 1
        y = y + 1
        x = x + 1
Loop
        Range("A5").Select
        Range(Selection, Selection.End(xlDown)).Select
        Range(Selection, Selection.End(xlToRight)).Select
With Selection.Borders

        .LineStyle = xlcontinous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin

End With
        Range("g6").Select
Do Until ActiveCell.Offset(0, -5) = ""
    If ActiveCell.Value = "KCB Limuru" Then
        ActiveCell.Offset(1, 0).Select
    Else: ActiveCell.Interior.Color = vbYellow
        ActiveCell.Offset(1, 0).Select
    End If
Loop
        Columns("Q:Q").Select
        Range(Selection, Selection.End(xlToRight)).Select

With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
End With

        Range("R6").Select
Do Until ActiveCell.Offset(0, -5) = ""
        ActiveCell.FormulaR1C1 = "=RC[-5]&"".00'"""
        ActiveCell.Offset(1, 0).Select
Loop
        Range("R6").Select
If Range("R7").Value = "" Then
        Selection.Copy
Else
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
End If
        Range("M6").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Application.CutCopyMode = False
        Columns("r:r").Delete
        Range("M6").Select
If Range("M7").Value = "" Then
        Selection.Interior.Color = vbRed
Else
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Interior.Color = vbRed
End If

        Columns("M:M").AutoFit
        Columns("C:C").Delete
        Range("M6").Select
If Range("M7").Value = "" Then
ActiveSheet.Name = "EFT"

        Selection.Copy
Else
    ActiveSheet.Name = "EFT"
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
End If
Range("a1").Select
        Filecheckoff.TrackCheckoff 'run macro to track checkoff

End Sub


Sub Cheques()
'cheque module
        Columns("A:U").Select
        Selection.Columns.AutoFit
        Columns("A:A").Insert
        Columns("A:A").ColumnWidth = 5
        Columns("B:B").ColumnWidth = 15
        Columns("R:R").Delete
        Columns("S:S").Delete
        Columns("K:K").Delete
        Range("A5").Value = "NO"
        Range("a5").Font.Bold = True
        Cells(1, 1).Select

'this do while loop will input column a data

Dim z As Integer
Dim q As Integer
z = 6
q = 0

Do Until Cells(z, 2).Value = ""
Cells(z, 1) = q + 1
q = q + 1
z = z + 1

Loop
        Range("A5").Select
        Range(Selection, Selection.End(xlDown)).Select
        Range(Selection, Selection.End(xlToRight)).Select
With Selection.Borders
        .LineStyle = xlcontinous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    
End With
'find the difference between amount disbursed and amount sent
        Range("R6").Select
Do Until ActiveCell.Offset(0, -5).Value = ""
        ActiveCell.FormulaR1C1 = "=RC[-5]-RC[-4]"
        ActiveCell.Offset(1, 0).Range("A1").Select
Loop
        Range("R6").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Application.CutCopyMode = False
        Range("R6").Select
If ActiveCell.Offset(1, -1).Value = "" Then
        ActiveCell.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Application.CutCopyMode = False
Else
Do Until ActiveCell.Offset(0, -1).Value = ""
    If ActiveCell.Value = 0 Then
        ActiveCell.Offset(1, 0).Select
    Else: ActiveCell.EntireRow.Interior.Color = vbCyan
        ActiveCell.Offset(0, -4).Font.Bold = True
        ActiveCell.Offset(0, -4).Interior.Color = vbYellow
        ActiveCell.Offset(1, 0).Select
    End If
Loop
End If
        Columns("R:R").Delete
        Columns("R:R").Select
        Range(Selection, Selection.End(xlToRight)).Select
With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
End With

If MsgBox("Do you want to generate bank letters?", vbYesNo, "Generate Reports") = vbYes Then
MsgBox ("Type in the payees in the selected column"), vbOKOnly

 BankLetters
Else
 
        Range("B6").Select
        Range(Selection, Selection.End(xlDown)).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
         End If
End Sub

Sub MpesaValidationCleanup()

'validation clean up module

        Rows("1:3").Delete
        Rows("3:7").Delete
        Columns("E:G").Select
        Selection.Delete
        Columns("D:D").Select
        Selection.NumberFormat = "0"
        Columns("A:h").AutoFit
        Range("a4").Value = Range("b2").Value
        Columns("B:C").Delete
        Columns("A:A").ColumnWidth = 10
        Columns("E:E").Copy
End Sub

Sub MpesaValidation()

'validation/check names module
        Worksheets("Disbursals").Select
        Columns("I:I").Select
On Error Resume Next
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("I5").Value = "Mpesa name"
        Columns("I:I").AutoFit
        Columns("I:I").Select
        Selection.Copy
        Worksheets("validation").Select
'Columns("I:I").Select
On Error Resume Next
        Columns("I:I").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
If Err.Number <> 0 Then Range("a1").Select
        Application.CutCopyMode = False
    If Range("H5").Value = "Customer" And Range("i5").Value = "Mpesa name" Then
        Columns("A:f").Select
        Selection.Delete
        Columns("d:k").Select
        Selection.Delete
        Columns("c:c").Select
        Selection.TextToColumns Destination:=Range("c1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1)), TrailingMinusNumbers:=True
'this do while loop will return a value if the search is true
        Range("f6").Select
Do Until ActiveCell.Offset(0, -5) = ""
        ActiveCell.FormulaR1C1 = "=IF(ISNUMBER(SEARCH(RC[-3],RC[-4])),""yes"",""no"")"
        ActiveCell.Offset(1, 0).Select
Loop
        Range("g6").Select
Do Until ActiveCell.Offset(0, -5) = ""
        ActiveCell.FormulaR1C1 = "=IF(ISNUMBER(SEARCH(RC[-3],RC[-5])),""yes"",""no"")"
        ActiveCell.Offset(1, 0).Select
Loop
 ' pastvalues Macro
        Range("f6").Select
        If Range("f7") = "" Then
            Range(Selection, Selection.End(xlToRight)).Select
        Else
            Range(Selection, Selection.End(xlToRight)).Select
            Range(Selection, Selection.End(xlDown)).Select
        End If
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Application.CutCopyMode = False
'this do while loop will return accept or reject
        Range("h6").Select
Do Until ActiveCell.Offset(0, -7) = ""
    If ActiveCell.Offset(0, -2).Value = "no" And ActiveCell.Offset(0, -1).Value = "no" Then
    ActiveCell.Value = "reject"
    ElseIf ActiveCell.Offset(0, -2).Value = "yes" And ActiveCell.Offset(0, -1).Value = "no" Or ActiveCell.Offset(0, -1).Value = "yes" And ActiveCell.Offset(0, -2).Value = "no" Then
    ActiveCell.Value = "confirm"
    Else
    ActiveCell.Value = "accept"
    End If
    ActiveCell.Offset(1, 0).Select
Loop
     Range("d6").Select
Do Until ActiveCell.Offset(0, -2) = ""
    If ActiveCell.Value = "" And ActiveCell.Offset(0, -1).Value = "" Then
    ActiveCell.Offset(0, 4).Value = "reject"
    ElseIf Len(ActiveCell) = "2" Or Len(ActiveCell.Offset(0, -1)) = "2" Or Len(ActiveCell) = "1" Or Len(ActiveCell.Offset(0, -1)) = "1" Then
    ActiveCell.Offset(0, 4).Value = "confirm"
    Else: ActiveCell.Select
    End If
    ActiveCell.Offset(1, 0).Select
Loop
    Columns("f:g").Select
    Selection.Delete
    Columns("c:e").Select
    Selection.EntireColumn.AutoFit
    Range("f6").Select
Do Until ActiveCell.Offset(0, -4) = ""
    If ActiveCell.Value = "reject" Then
    ActiveCell.EntireRow.Interior.Color = vbRed
    ElseIf ActiveCell.Value = "confirm" Then
    ActiveCell.EntireRow.Interior.Color = vbYellow
    Else
    ActiveCell.Select
    End If
    ActiveCell.Offset(1, 0).Select
Loop
        Rows("1:5").Delete
        Range("g1").Select
        Range(Selection, Selection.End(xlDown)).Select
        Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
   ' this module will copy the rows with issues.
    Dim j As Integer
    Dim g As Integer
    g = 1
    j = 2
   Cells(g, 2).Select
Do Until ActiveCell.Offset(0, -1) = ""
        Cells(g, 2).Select
    If ActiveCell.Interior.Color = vbRed Or ActiveCell.Interior.Color = vbYellow Then
        ActiveCell.Offset(0, -1).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
        Cells(j, 9).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Selection.PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
        ActiveCell.Offset(1, 0).Select
        j = j + 1
        g = g + 1
        Cells(g, 2).Select
    Else: ActiveCell.Offset(1, 0).Select
        g = g + 1
    End If
Loop

  'this module will input the branch name even when result is one
        Range("H2").Select
If Range("I3").Value = "" Then
    ActiveCell.FormulaR1C1 = _
    "=INDEX(Disbursals!C[-3],MATCH(validation!RC[1],Disbursals!C[-1],0))"
ElseIf Range("I2").Value = "" Then
    ActiveCell.Offset(1, 0).Select
Else
    Do Until ActiveCell.Offset(0, 1) = ""
        ActiveCell.FormulaR1C1 = _
        "=INDEX(Disbursals!C[-3],MATCH(validation!RC[1],Disbursals!C[-1],0))"
        ActiveCell.Offset(1, 0).Select
    Loop
End If
    Range("H2").Select
If Range("h3").Value = "" Then
    Selection.Copy
Else
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
End If
    ActiveCell.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
  'this module will fill reg as and unregistered
Columns("K:K").Insert
Range("K2").Select
Do Until ActiveCell.Offset(0, -1) = ""
    If ActiveCell.Offset(0, 1) = "" Then
        ActiveCell.Value = "Unregistered"
        ActiveCell.Offset(1, 0).Select
    Else: ActiveCell.FormulaR1C1 = "=""Reg as""&"" ""&RC[1]&"" ""&RC[2]&"" ""&RC[3]"
        ActiveCell.Offset(1, 0).Select
    End If
Loop
Range("K2").Select
If Range("I2").Value = "" Then
    Selection.Copy
ElseIf Range("I3").Value = "" Then
    Selection.Copy
Else
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
End If
ActiveCell.Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Columns("h:N").AutoFit
'This module will return the row value of the customer in the disbursals sheet
Range("g2").Select
Do Until ActiveCell.Offset(0, 2) = ""
    ActiveCell.FormulaR1C1 = "=MATCH(RC[2],Disbursals!C,0)"
    ActiveCell.Offset(1, 0).Select
Loop
'This module will sort the reds and the yellows.
    Range("H1:M1").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("validation").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("validation").AutoFilter.Sort.SortFields.Add(Range( _
    "I1"), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB( _
    255, 0, 0)
    With ActiveWorkbook.Worksheets("validation").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
  'this module will make highlights in the disbursals sheet
Dim p As Integer
Range("g2").Select
Do Until ActiveCell.Offset(0, 2) = ""
    If ActiveCell.Offset(0, 2).Interior.Color = vbRed Then
        p = ActiveCell.Value
        Worksheets("Disbursals").Rows(p).EntireRow.Font.Strikethrough = True
        Worksheets("Disbursals").Cells(p, 8).Interior.Color = vbRed
        Worksheets("Disbursals").Cells(p, 17).Value = "reject"
        Worksheets("Disbursals").Cells(p, 17).Font.Strikethrough = False
        ActiveCell.Offset(1, 0).Select
        p = ActiveCell.Value
    Else:
        p = ActiveCell.Value
        Worksheets("Disbursals").Cells(p, 8).Interior.Color = vbYellow
        Worksheets("Disbursals").Cells(p, 8).Value = Worksheets("Disbursals").Cells(p, 8).Value & "**"
        Worksheets("Disbursals").Cells(p, 8).Font.Bold = True
        ActiveCell.Offset(1, 0).Select
        p = ActiveCell.Value
    End If
Loop
Range("g1").Value = "Row num"
Range("h1").Value = "Branch"
Range("i1").Value = "Loan Account"
Range("J1").Value = "system Name"
Range("K1").Value = "Mpesa Name"
Columns("L:O").Delete
Columns("a:f").Delete
Columns("c;c").AutoFit
Range("A1:G1").AutoFilter.Clear
Worksheets("Disbursals").Columns("h:h").AutoFit
Else: MsgBox "Failed!! Validation Names not found." & vbCrLf & "Copy Names column from MPesa validation Sheet and Try again.", vbExclamation
 End If
End Sub

Sub refundsCleanup()
Columns("O:O").Delete
Columns("m:n").Delete
Range("A5").Select
Columns("N:N").Delete
Columns("a:a").Insert
Columns("b:b").ColumnWidth = 13
Columns("c:n").AutoFit
Columns("j:j").ColumnWidth = 15
'Inserting column a data
Range("a1").Select
Dim num As Integer
Dim dig As Integer
num = 6
dig = 0
Do Until Cells(num, 2).Value = ""
    Cells(num, 1) = dig + 1
    num = num + 1
    dig = dig + 1
Loop
Columns("a:a").ColumnWidth = 4
Range("a5").Value = "NO"

'adding 254 to phone numbers
Range("O6").Select
Do Until ActiveCell.Offset(0, -1) = ""
    ActiveCell.FormulaR1C1 = "=concatenate(254,RC[-5])"
    ActiveCell.Offset(1, 0).Select
Loop
Range("O6").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Range("J6").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
Selection.Columns.AutoFit
Application.CutCopyMode = False
Columns("O:O").Delete

'adding boarders
Range("A5").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select
   With Selection.Borders
        .LineStyle = xlcontinous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
Dim rng As Range
    Range("e6").Select
    Range(Selection, Selection.End(xlDown)).Select
Set rng = Selection 'Change range as required
If rng.Count = WorksheetFunction.CountIf(rng, "MPESA Account Safaricom") Then
    Columns("f:f").Delete
    Range("I6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
ElseIf rng.Count = WorksheetFunction.CountIf(rng, "KCB Kapsowar") Then
    Range("a1").Select
Else: MsgBox "Separate Mpesa and EFT refunds"
Exit Sub
End If
ActiveSheet.Name = "Refunds Validation"
Sheets.Add
ActiveSheet.Name = "Customer refunds"
Worksheets("Refunds Validation").Select
Cells.Select
Selection.Copy
Worksheets("Customer refunds").Select
ActiveSheet.Paste
Worksheets("Refunds Validation").Activate
Columns("B:C").Delete
Columns("C:C").Delete
Columns("E:E").Delete
Columns("F:I").Delete
Range("e6").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
End Sub

Sub RefundsValidation()
'Refunds validation/check names module
Worksheets("Customer refunds").Activate
Columns("h:h").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Range("H5").Value = "MPesa Name"
Columns("H:H").AutoFit
If Err.Number <> 0 Then Range("a1").Select
    Worksheets("Refunds Validation").Activate
    Columns("f:f").Select
    On Error Resume Next
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        If Err.Number <> 0 Then Range("a1").Select
            Application.CutCopyMode = False
            Columns("e:e").Delete
            If Range("E5").Value <> "Customer Name" Then
                MsgBox (" Refund Validation names not found")
            Else
                 Columns("E:E").Select
                Selection.TextToColumns Destination:=Range("E1"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
                Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
                :=Array(Array(1, 1), Array(2, 1), Array(3, 1)), TrailingMinusNumbers:=True
 End If
'this do while loop will return a value if the search is true
        Range("H6").Select
Do Until ActiveCell.Offset(0, -7) = ""
    ActiveCell.FormulaR1C1 = "=IF(ISNUMBER(SEARCH(RC[-3],RC[-4])),""yes"",""no"")"
    ActiveCell.Offset(1, 0).Select
Loop
    Range("I6").Select
Do Until ActiveCell.Offset(0, -5) = ""
    ActiveCell.FormulaR1C1 = "=IF(ISNUMBER(SEARCH(RC[-3],RC[-5])),""yes"",""no"")"
    ActiveCell.Offset(1, 0).Select
Loop
 ' pastvalues Macro
    Range("H6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
'this do while loop will return accept or reject
      Range("J6").Select
Do Until ActiveCell.Offset(0, -7) = ""
    If ActiveCell.Offset(0, -2).Value = "no" And ActiveCell.Offset(0, -1).Value = "no" Then
        ActiveCell.Value = "reject"
    ElseIf ActiveCell.Offset(0, -2).Value = "yes" And ActiveCell.Offset(0, -1).Value = "no" Or ActiveCell.Offset(0, -1).Value = "yes" And ActiveCell.Offset(0, -2).Value = "no" Then
     ActiveCell.Value = "confirm"
    Else
        ActiveCell.Value = "accept"
    End If
    ActiveCell.Offset(1, 0).Select
Loop
'This do while loop will return accept, reject or confirm
     Range("F6").Select
Do Until ActiveCell.Offset(0, -2) = ""
    If ActiveCell.Value = "" And ActiveCell.Offset(0, -1).Value = "" Then
        ActiveCell.Offset(0, 4).Value = "reject"
    ElseIf Len(ActiveCell) = "2" Or Len(ActiveCell.Offset(0, -1)) = "2" Or Len(ActiveCell) = "1" Or Len(ActiveCell.Offset(0, -1)) = "1" Then
        ActiveCell.Offset(0, 4).Value = "confirm"
    Else: ActiveCell.Select
    End If
        ActiveCell.Offset(1, 0).Select
Loop
Columns("H:I").Delete
Columns("E:H").EntireColumn.AutoFit
Range("H6").Select
Do Until ActiveCell.Offset(0, -4) = ""
    If ActiveCell.Value = "reject" Then
        ActiveCell.EntireRow.Interior.Color = vbRed
    ElseIf ActiveCell.Value = "confirm" Then
        ActiveCell.EntireRow.Interior.Color = vbYellow
    Else
        ActiveCell.Select
    End If
        ActiveCell.Offset(1, 0).Select
Loop
Rows("1:5").Delete
Range("I1").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select
        With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
        End With
Dim ref As Integer
Range("A1").Select
Do Until ActiveCell.Offset(0, 1) = ""
    If ActiveCell.Interior.Color = vbRed Then
        ref = ActiveCell.Value + 5
        Worksheets("Customer refunds").Cells(ref, 7).Interior.Color = vbRed
        Worksheets("Customer refunds").Cells(ref, 7).EntireRow.Font.Strikethrough = True
        Worksheets("Customer refunds").Cells(ref, 14).Value = "reject"
        Worksheets("Customer refunds").Cells(ref, 14).Font.Strikethrough = False
        ActiveCell.Offset(1, 0).Select
    ElseIf ActiveCell.Interior.Color = vbYellow Then
        ref = ActiveCell.Value + 5
        Worksheets("Customer refunds").Cells(ref, 7).Interior.Color = vbYellow
        Worksheets("Customer refunds").Cells(ref, 7).Value = Worksheets("Customer refunds").Cells(ref, 7).Value & "**"
        Worksheets("Customer refunds").Cells(ref, 7).Font.Bold = True
        ActiveCell.Offset(1, 0).Select
    Else
    ActiveCell.Offset(1, 0).Select
    End If
Loop

End Sub

Sub ReadyToDisburse()
ActiveSheet.Name = "ReadyToDisburse"
   Range("a5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
       ActiveSheet.Range("$A$5:$O$2000").AutoFilter Field:=3, Criteria1:=Array( _
        "Homa Bay", "Kenyenya", "Kerugoya", "Meru", "Migori", "Nyamira", "Thika", "Chuka", "Embu", "Maua", "Oyugis", "Nyeri", "Rongo", "CBD"), Operator:=xlFilterValues

        Columns("a:b").ColumnWidth = 15
        
       Range("a6").Select
       Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("ReadyToDisburse").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ReadyToDisburse").Sort.SortFields.Add2 Key:= _
        ActiveCell.Offset(-1, 9).Range("A1:A2000"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("ReadyToDisburse").Sort.SortFields.Add2 Key:= _
        ActiveCell.Offset(-1, 2).Range("A1:A2000"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ReadyToDisburse").Sort
        .SetRange ActiveCell.Offset(-1, 0).Range("A1:O2000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    
              Range("a5").Select
        Range(Selection, Selection.End(xlDown)).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
        
End Sub

Sub BankLetters()
ActiveSheet.Name = "ChqDisbursals"
Cells.Copy
Sheets.Add
Worksheets("Sheet1").Select
ActiveSheet.Paste
ActiveSheet.Name = "reports"
Rows("1:4").Delete
Range("r1").Value = "PAYEE"
Range("r1").Select
Columns("r:r").ColumnWidth = 50

If Range("a3").Value = Empty Then
Range("r2").Select
Else

    Range("a2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("reports").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("reports").Sort.SortFields.Add2 Key:=ActiveCell. _
        Range("A1:A200"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("reports").Sort
        .SetRange ActiveCell.Offset(-1, 0).Range("A1:R39")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End If

'Columns("f:j").Insert
'Dim strokes As Integer
'Range("f2").Select
'Do Until ActiveCell.Offset(0, -1) = Empty
' ActiveCell.FormulaR1C1 = "=LEN(RC[-1])-LEN(SUBSTITUTE(RC[-1],""/"",""""))"
'ActiveCell.Offset(1, 0).Select
'Loop
End Sub
