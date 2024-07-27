Attribute VB_Name = "RiskAutomations"


Sub GroupAutomation()
If Range("o2").Value = "" Then
If MsgBox("Do you want to continue with a new report?", vbYesNo, "Generate Reports") = vbYes Then
Dim ws As Worksheet
    Dim lastRow As Long
    Dim LastCell As Long
  
    Worksheets("Groups").Range("A2:H1000").ClearContents
    Worksheets("Loans").Range("J2:J1000").ClearContents
    Worksheets("Loans").Range("O2:R1000").ClearContents
    Worksheets("Loans").Range("T2:Z1000").ClearContents
    LastCell = Cells(Rows.Count, "B").End(xlUp).Row
    Range("e2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("E2").Select
    'ActiveCell.Formula2R1C1 = "=RC[-1]:R[26]C[-1]&RC[-4]:R[26]C[-4]"
    ActiveCell.Formula2R1C1 = "=RC[-1]:R[" & LastCell - 2 & "]C[-1]&RC[-4]:R[" & LastCell - 2 & "]C[-4]"
    Range("j2").Select
     Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("J2").Select
    'ActiveCell.Formula2R1C1 = _
        "=HYPERLINK(""https://jk.mfiexpert.app/loans/loanaccountsne.aspx?id=""&RC[-8]:R[" & LastCell & "]C[-8])"
        ActiveCell.Formula2R1C1 = "=HYPERLINK(""https://jk.mfiexpert.app/search/search.aspx?search=""&RC[-8]:R[" & LastCell - 2 & "]C[-8])"
Range("D2:I2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Worksheets("Groups").Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
             Application.CutCopyMode = False
Worksheets("Groups").Activate
Set ws = ActiveSheet
 lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
Worksheets("Groups").Range("A1:F1").Resize(lastRow).RemoveDuplicates Columns:=2, Header:=xlYes
Worksheets("Groups").Range("G2:H1000").ClearContents
End If
End If
End Sub
Sub UIpathGraduation()
Range("j1").Interior.Color = vbRed
End Sub

