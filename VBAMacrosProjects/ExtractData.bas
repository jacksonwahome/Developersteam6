Attribute VB_Name = "ExtractData"
Sub ExtractData()
Dim Wd As New Word.Application
Dim Doc As Word.Document
Dim Sh As Worksheet
Wd.Visible = True
'Set Doc = Wd.Documents.Open(ActiveWorkbook.Path & "\form.docx")
Set Doc = Wd.Documents.Open("C:\Users\User1\Desktop\Risk sanctions\Danson\DANIEL KEIGE GACHUE-MINI SUPERMARKET.docx")
Set tbls = Doc.Tables
Set Sh = ActiveSheet
 tbls(1).Range.Copy
Range("a1").Select
ActiveSheet.Paste
Wd.Quit
Selection.UnMerge
Selection.WrapText = False
Selection.Columns.AutoFit
End Sub
