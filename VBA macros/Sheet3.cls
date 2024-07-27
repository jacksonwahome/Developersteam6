VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub UIpath()
Dim ws As Worksheet
Dim lastRow As Long
Set ws = ActiveWorkbook.Sheets("Graduation")
Columns("J:O").Delete
  
' Clear any existing filters
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If

    ' Apply filter to show only rows where column c is "Group - Agribusiness"
    ws.Range("A1:F" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=3, Criteria1:=Array("Group - Agribusiness", "Group - Agri Asset Acquisition", "Group - Animal Farming"), Operator:=xlFilterValues
    Range("A1:F1").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Range("J1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
             Application.CutCopyMode = False
               If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    Columns("H:M").AutoFit
    Worksheets("Loans").Cells(ws.Range("I2").Value, 21).Value = Range("N3").Value
    Worksheets("Loans").Cells(ws.Range("I2").Value, 22).Value = Range("O3").Value
    Worksheets("Loans").Cells(ws.Range("I2").Value, 24).Value = Range("K3").Value
    ws.Columns("A:F").ClearContents
    ws.Columns("J:P").ClearContents
    
End Sub

Sub UIPathGrad()
Dim ws As Worksheet
Dim lastRow As Long
Set ws = ActiveWorkbook.Sheets("Graduation")
ws.Range("H2:M500").ClearContents
   ' Find the last row with data in the specified range
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    ' Delete the last two rows
    If lastRow > 1 Then
        ws.Rows(lastRow).Delete
        ws.Rows(lastRow - 1).Delete
    ElseIf lastRow = 1 Then
        ws.Rows(lastRow).Delete
    End If
' Clear any existing filters
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If

    ' Apply filter to show only rows where column c is "Group - Agribusiness"
    ws.Range("A1:F" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=3, Criteria1:=Array("Group - Agribusiness", "E-Loan", "Clean Energy", "Elimika", "Simu Yangu", "Group - Water Tank"), Operator:=xlFilterValues
    
End Sub
