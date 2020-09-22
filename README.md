<div align="center">

## Create a Internet Shorcut on Anyone's Computer


</div>

### Description

This code will create an internet shortcut on someone's computer. All you have to do it call it with a path and hyperlink!
 
### More Info
 
Path- path should be complete with filename ending in url (e.g. "C:\Shortcut.url")

Hyperlink- this is the complete website address

So easy I think you can handle it.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Beginner
**User Rating**    |4.0 (20 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/create-a-internet-shorcut-on-anyone-s-computer__1-12248/archive/master.zip)





### Source Code

```
Private Sub CreateHyperlink(Path As String, Hyperlink as String)
  Open Path For Output As #1 'open file access
  Print #1, "[Internetshortcut]" 'print on first line
  Print #1, "URL=" & Hyperlink 'print url on second line
  Close #1 'close it
End Sub
```

