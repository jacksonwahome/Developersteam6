Attribute VB_Name = "filesandfolders"
Sub filesandFolders()
Dim fso As Scripting.FileSystemObject
Set fso = New Scripting.FileSystemObject
'creating a folder using vba
If Not fso.FolderExists("C:\Users\User1\Desktop\Risk and sanctions2") Then
fso.CreateFolder ("C:\Users\User1\Desktop\Risk and sanctions2")
End If

End Sub
'OR
Sub filesandFolders2()
Dim fso As Scripting.FileSystemObject
Set fso = New Scripting.FileSystemObject
Dim folderpath As String
folderpath = "C:\Users\User1\Desktop\"
'creating a folder using vba
If Not fso.FolderExists("C:\Users\User1\Desktop\Risk and sanctions2") Then
fso.CreateFolder (folderpath & "MIMI")
End If

End Sub
'The problems again with this code is that it is hardcoded. eg. users, users1 etc or someones name
'you want to make sure that it automatically picks the folder name from the system
'You will need to use the environ function

Sub filesandFolders3()
'this code uses the environ function to pick the user automatically
'use the immediate window to get the user
Dim fso As Scripting.FileSystemObject
Set fso = New Scripting.FileSystemObject
Dim folderpath As String
folderpath = Environ("userprofile") & "\Desktop\" 'here is the environ function +the desktop
'creating a folder using vba
If Not fso.FolderExists("C:\Users\User1\Desktop\Risk and sanctions2") Then
fso.CreateFolder (folderpath & "wewe")
End If

End Sub



'copying a file from one folder to another folder using VBA

Sub filesandFolders4()
Dim fso As Scripting.FileSystemObject
Set fso = New Scripting.FileSystemObject
Dim folderpath As String
folderpath = Environ("userprofile") & "\Desktop\"
'using the copyfile function
fso.CopyFile Source:=folderpath & "MIMI\Gabriel Namasaka Barasa.docx", _
Destination:=folderpath & "wewe\Gabriel Namasaka Barasa.docx"
'if the fil is not found, it gives and error "file not found"
'else you can use if statement to check if file exist in the source folder
End Sub
Sub filesandFolders5()
Dim fso As Scripting.FileSystemObject
Set fso = New Scripting.FileSystemObject
Dim folderpath As String
folderpath = Environ("userprofile") & "\Desktop\"

If fso.FileExists(folderpath & "MIMI\Gabriel Namasaka Barasa.docx") Then
'using the copyfile function
fso.CopyFile Source:=folderpath & "MIMI\Gabriel Namasaka Barasa.docx", _
Destination:=folderpath & "wewe\Gabriel Namasaka Barasa.docx"
End If
End Sub


'How to copy multiple files from one folder to another
Sub filesandFolders6()
Dim fso As Scripting.FileSystemObject
Dim fileRef As Scripting.file
Dim folderRef As Scripting.Folder
Set fso = New Scripting.FileSystemObject
Dim folderpath As String
folderpath = Environ("userprofile") & "\Desktop\"

Set folderRef = fso.GetFolder(folderpath & "MIMI")
For Each fileRef In folderRef.Files
'Debug.Print fileRef.Name 'print each name of the files in the selected folder
'what if you wanted only excel files?
    If fso.GetExtensionName(fileRef.Name) = "xlsx" Then
    Debug.Print fileRef.Name
    End If
Next fileRef
End Sub

'copying all the files
Sub filesandFolders7()
Dim fso As Scripting.FileSystemObject
Dim fileRef As Scripting.file
Dim folderRef As Scripting.Folder
Set fso = New Scripting.FileSystemObject
Dim folderpath As String
folderpath = Environ("userprofile") & "\Desktop\"

Set folderRef = fso.GetFolder(folderpath & "MIMI")
For Each fileRef In folderRef.Files
fileRef.Copy (folderpath & "wewe\" & fileRef.Name)
Next fileRef
End Sub

'copying only excel files

Sub filesandFolders8()
Dim fso As Scripting.FileSystemObject
Dim fileRef As Scripting.file
Dim folderRef As Scripting.Folder

Set fso = New Scripting.FileSystemObject
Dim folderpath As String
folderpath = Environ("userprofile") & "\Desktop\"

Set folderRef = fso.GetFolder(folderpath & "MIMI")
For Each fileRef In folderRef.Files
    If fso.GetExtensionName(fileRef.Name) = "xlsx" Then
    fileRef.Copy (folderpath & "wewe\" & fileRef.Name)
    End If
Next fileRef
End Sub
'what if you have folders and subfolders?

Sub filesandFolders9()
Dim fso As Scripting.FileSystemObject
Dim fileRef As Scripting.file
Dim folderRef As Scripting.Folder
Dim subfol As Scripting.Folder
Set fso = New Scripting.FileSystemObject
Dim folderpath As String
folderpath = Environ("userprofile") & "\Desktop\"
'you need to have two loops. one to copy the files and the other one to copy the folders.
Set folderRef = fso.GetFolder(folderpath & "MIMI")
For Each fileRef In folderRef.Files
    fileRef.Copy (folderpath & "wewe\" & fileRef.Name)
Next fileRef

For Each subfol In folderRef.SubFolders
    subfol.Copy (folderpath & "wewe\" & subfol.Name)
Next subfol


End Sub
