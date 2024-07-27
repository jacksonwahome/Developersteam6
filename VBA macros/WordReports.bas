Attribute VB_Name = "WordReports"
'this code will only work if you have checked the microsoft word object library
'this method is called early binding
Sub CreateWordReport()
Dim wordApp As Word.Application
Set wordApp = New Word.Application
wordApp.Visible = True
wordApp.Activate

End Sub
'But what if you want to share with someone who has not activated the word object library?
'or if someone has a different version of office?

Sub CreateWordReporUsingCreateObject()
'this method is called late binding
'The big downside of this method is that it has not intellisense and means you will not be able to use the child objects
'instead you can code using early binding and then share it with late binding.
Dim wordApp As Object
Set wordApp = CreateObject("Word.Application")
wordApp.Visible = True
wordApp.Activate
End Sub
'How to copy and excel table into a word document, close and save that document
Sub CreateWordReport1()
Dim wordApp As Word.Application
Set wordApp = New Word.Application
With wordApp
    .Visible = True
    .Activate
    .Documents.Add
    With .Selection
        .BoldRun
        .Paragraphs.Alignment = wdAlignParagraphCenter
        .Font.Size = 14
        .TypeText "Quartery Report"
        .BoldRun
        .TypeParagraph
        .Font.Size = 11
        Range("a1", Range("a1").End(xlDown).End(xlToRight)).Copy
        .Paste
    End With

End With
'add a now functiont add date to you documents so that its not overwritten when saved again
'however office would not allow a colon as part of a file name
'use the format ?format(now,"yyyy-mm-dd hh-mm-ss")
wordApp.ActiveDocument.SaveAs2 Environ("userprofile") & "\Desktop\quartery report " & Format(Now, "yyyy-mm-dd hh-mm-ss") & ".docx"
wordApp.ActiveDocument.Close
'you can add wordapp.quit to quit the application
'but you even don't need to make it visible in the first place
'finally when converting early binding to late binding, change all the variables to the numbers representing those variables so that
'they are understood by createobject
End Sub
