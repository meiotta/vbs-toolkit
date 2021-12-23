Const FOR_READING = 1 
Const FOR_WRITING = 2 

'set up the file path and number of lines to remove
strFileName = "\\netdrive\netfolder\netsubfolder\file.csv" 
iNumberOfLinesToDelete = 1




Set objFS = CreateObject("Scripting.FileSystemObject") 
fsize = objFS.GetFile(strFilename).size
'check if empty
If fsize < 200 Then

WScript.Quit
End if

'put the whole thing in memory, read only
Set objTS = objFS.OpenTextFile(strFileName, FOR_READING) 
strContents = objTS.ReadAll 
objTS.Close 

'open a new copy
arrLines = Split(strContents, vbNewLine) 
Set objTS = objFS.OpenTextFile(strFileName, FOR_WRITING) 

'if the line number is above the number we're removing, write it off the open memory copy
For i=0 To UBound(arrLines) 
If i > (iNumberOfLinesToDelete - 1) Then 
  objTS.WriteLine arrLines(i) 
End If 
Next 


WScript.Quit
