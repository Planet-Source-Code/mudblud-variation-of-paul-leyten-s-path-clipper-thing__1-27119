<div align="center">

## Variation of Paul Leyten's Path Clipper thing\.


</div>

### Description

splits a path up a path and shortens parts to make it shorter.

e.g.

c:\program files\blingblongblu\bah\bleh.exe

c:\program...\blingbl...\bah\bleh.exe
 
### More Info
 
path

shorten'd path


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[MudBlud](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mudblud.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mudblud-variation-of-paul-leyten-s-path-clipper-thing__1-27119/archive/master.zip)





### Source Code

```
Private Function ShortenPath(Path As String, MaxLen As Integer) As String
Dim bleh() As String
bleh = Split(Path, "\")
For x = 0 To UBound(bleh)
  If Not x = UBound(bleh) Then
    If Len(bleh(x)) > MaxLen Then
      bleh(x) = Mid$(bleh(x), 1, MaxLen - 3) & "..."
    End If
    tmp = tmp & bleh(x) & "\"
  Else
    tmp = tmp & bleh(x)
  End If
Next
ShortenPath = tmp
End Function
```

