'PDF-txt-Word 
'1.0 
'9/5/2012 by tmoore82

'This script converts PDFs to plain text, then copies and pastes that plain text to a new Word document
'This can be used when you want the text from a PDF without any of the formatting, 
'especially when you suspect the document is corrupted.

'Users should save the pdf to their Desktop before running the script.

'string to hold file path (DMM is short for Don't Mind Me, as this is a space for temporary files the user can ignore.)
Dim strDMM
strDMM = "C:\dmm"

'make this directory if it doesn't exits
On Error Resume Next
MkDir strDMM
On Error GoTo 0

'get the username to go to the right filepath
Dim strUser
strUser = InputBox("What is your username?" & chr(13) & chr(13) & "(Example: [myname])", "Username")

'get the file name to process
Dim TheFile
TheFile = InputBox("What is the file name?" & chr(13) & chr(13) & "(Example: [file.pdf])", "Name of File")

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

'declare constants for manipulating the text files
Const ForReading = 1
Const ForWriting = 2

'Create a File System Object
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

'open the text file
dim objFile
set objFile=objFSO.OpenTextFile("C:\dmm\pdf-plain-text.txt", ForReading)

'Create a Word Object
Dim objWord
set objWord = CreateObject("Word.Application")

'make Word hidden
With objWord
	.Visible = False
End With

'create a blank document
Dim objDoc
Set objDoc=objWord.Documents.Add()

'create a shorter variable to pass commands to Word
Dim objSelection
set objSelection=objWord.Selection

'Thanks to Kurt for the following!
'Read one line at a time from the text file and 
'type that line into Word until the end of the file is reached 
Dim strLine 
Do Until objFile.AtEndOfStream    
	strLine = objFile.ReadLine    
	objSelection.TypeText strLine 
	objSelection.TypeParagraph
Loop

objFile.Close

'make Word visible
With objWord
	.Visible = True
End With

'Remind the user to save
MsgBox ("Save the document.")