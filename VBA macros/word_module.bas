Attribute VB_Name = "word_module"

' this would not work for someone who doesnt have the microsoft 16 oject library
'that is why the block has been commented
'This method is called early binding

Sub BankLetters()
Dim wordApp As Word.Application
Set wordApp = New Word.Application


With wordApp
.Visible = True  'you can uncomment this if you want to see the report
.Activate ' up to this point there is no document
.Documents.Add 'to add a new doc

With .Selection
.Font.Size = 12
.BoldRun
.TypeText "KCB Kapsowar"
.Font.Size = 12
.TypeParagraph
.TypeText "Kapsowar Branch "
.TypeParagraph
.BoldRun
.TypeText "Dear Sir,"
.TypeParagraph
.Font.Underline = wdUnderlineSingle
.TypeText "SUBJECT:  CHEQUES DISBURSEMENT:"
.Font.Underline = none
.TypeParagraph
.TypeText "Attached Please, find a list of cheque (s) issued to our client (s) from our Account number 1160345643 "
.TypeText "for verification before you make any payments to them. In case of any clarification, please call the undersigned on:"
.TypeParagraph
.TypeText "0709692000 or 0709692030"
.TypeParagraph
.Font.Size = 20
.BoldRun
.TypeText "BRANCH: KERUGOYA"
.Font.Size = 11


End With
End With
'wordapp.ActiveDocument.SaveAs2 Environ("userprofile") & "\Desktop\AUTOBANKLETTERSREPORT.DOCX"
'this is the same as saveas
'open the immediate window so that you can get the location to save the document
'run the ?environ("userprofile")command
'However if you save it that way it will continue overwriting the existing files. you can instead add date and time to separate the files
'run the "now" command in immediate window but note that windows wont allow colons in file names so format using dashes
'Sub bankletters()
'wordapp.ActiveDocument.SaveAs2 Environ("userprofile") & "\Desktop\AUTOBANKLETTERS\REPORT " & Format(Now, "yyyy-mm-dd hh-mm-ss") & ".DOCX"
'wordapp.ActiveDocument.Close
'wordapp.Quit
End Sub

' now lets use another method that doesn't use the microsoft office 16 library
'This method is called late binding meaning that you do it when you need it.
'the downside with late binding is that when you start typing wordapp. it will not autocomplete the objects( intellisense)
'Hence it is a good ideas to use early binding when you are coding and then use late binding when you are done

'Dim wordapp As Object
'Set wordapp = CreateObject("word.application")
'wordapp.Visible = True
'wordapp.Activate
'End Sub

