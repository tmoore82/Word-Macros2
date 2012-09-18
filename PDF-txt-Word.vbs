'PDF-txt-Word 
'1.1 
'9/6/2012 by tmoore82

'This script converts PDFs to plain text, then copies and pastes that plain text to a new Word document
'This can be used when you want the text from a PDF without any of the formatting, 
'especially when you suspect the document is corrupted.

'Users should save the pdf to their Desktop before running the script.

MsgBox ("When I convert the PDF to Word, you may not see anything on the screen. I'll let you know when I'm done, though, K?")

Dim objFSO
Set objFSO = CreateObject ("Scripting.FileSystemObject")

'string to hold file path (DMM is short for Don't Mind Me, as this is a space for temporary files the user can ignore.)
Dim strDMM
strDMM = "C:\dmm\"

'make this directory if it doesn't exits
On Error Resume Next
objFSO.CreateFolder(strDMM)
On Error GoTo 0

'get the username to go to the right filepath
Dim strUser
strUser = InputBox("What is your username?" & chr(13) & chr(13) & "(Example: mooret)", "Username")

'get the file name to process
Dim TheFile
TheFile = InputBox("What is the file name?" & chr(13) & chr(13) & "(Example: 703582663_2.pdf)", "Name of File")

'declare some acrobat variables
Dim AcroXApp
Dim AcroXAVDoc
Dim AcroXPDDoc

'open acrobat
Set AcroXApp = CreateObject("AcroExch.App")
AcroXApp.Hide

'open the document we want
Set AcroXAVDoc = CreateObject("AcroExch.AVDoc")
AcroXAVDoc.Open "c:\Users\" & strUser & "\Desktop\" & TheFile, "Acrobat"

'make sure the acrobat window is active
AcroXAVDoc.BringToFront

'I don't know what this line does. As with a lot of this, I copied it from code online.
Set AcroXPDDoc = AcroXAVDoc.GetPDDoc

'activate JavaScript commands w/Acrobat
Dim jsObj
Set jsObj = AcroXPDDoc.GetJSObject

'save the file as plain text
jsObj.SaveAs "C:\dmm\" & "pdf-plain-text.txt", "com.adobe.acrobat.plain-text"

'close the file and exit acrobat
AcroXAVDoc.Close False
AcroXApp.Hide
AcroXApp.Exit

'Create a Word Object
Dim objWord
set objWord = CreateObject("Word.Application")

'insert the file into Word
With objWord
	.Visible = False
	.Documents.Add()
	.Selection.InsertFile "C:\dmm\pdf-plain-text.txt"
	'the next three lines lines are to change everything to our default style at work.
	'if you just want the plain text and want to go from there, you don't need them.
	.Selection.WholeStory
	.Selection.Style="Body Text"
	.Selection.Homekey
End With

'make Word visible
With objWord
	.Visible = True
End With

'Remind the user to save
MsgBox ("I'm all done! Save the document.")