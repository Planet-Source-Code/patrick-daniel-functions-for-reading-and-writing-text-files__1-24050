<div align="center">

## Functions For Reading and Writing Text Files


</div>

### Description

This code consists of 2 Function (ReadFile and WriteFile). All you have to do is Point ReadFile to a filepath and it will return the text within that file. Write file recieves a filepath and the string value to save to it. They include simple error handling..and can easily and quickly be advanced to handle open and save dialogs.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Patrick Daniel](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/patrick-daniel.md)
**Level**          |Beginner
**User Rating**    |4.9 (39 globes from 8 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/patrick-daniel-functions-for-reading-and-writing-text-files__1-24050/archive/master.zip)





### Source Code

```
'-- If you have any problems with this code please contact me
'-- at patrick1@mediaone.net. Feel free to drop me a line
'-- letting me know you are using this or if this code is
'-- helpfull to you. Enjoy!!
Public Function ReadFile(strPath As String) As Variant
On Error GoTo eHandler
  Dim iFileNumber As Integer
  Dim blnOpen As Boolean
  iFileNumber = FreeFile
  Open strPath For Input As #iFileNumber
  blnOpen = True
  ReadFile = Input(LOF(iFileNumber), iFileNumber)
eHandler:
  If blnOpen Then Close #iFileNumber
  If Err Then MsgBox Err.Description, vbOKOnly + vbExclamation, Err.Number & " - " & Err.Source
End Function
Public Function WriteFile(strPath As String, strValue As String) As Boolean
On Error GoTo eHandler
  Dim iFileNumber As Integer
  Dim blnOpen As Boolean
  iFileNumber = FreeFile
  Open strPath For Output As #iFileNumber
  blnOpen = True
  Print #iFileNumber, strValue
eHandler:
  If blnOpen Then Close #iFileNumber
  If Err Then
   MsgBox Err.Description, vbOKOnly + vbExclamation, Err.Number & " - " & Err.Source
  Else
   WriteFile = True
  End If
End Function
```

