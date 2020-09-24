<div align="center">

## PadString Any Side


</div>

### Description

Pads a string with any character you like. I usually use it to pad numbers with leading zeros. But you can use it for other things.
 
### More Info
 
Input ltring, length of return string, pad character, and side to pad

a string

No matter howmany characters in the pstrChar your pad character is the first character of pstrChar.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VisualProgman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/visualprogman.md)
**Level**          |Beginner
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VB Script
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/visualprogman-padstring-any-side__1-33748/archive/master.zip)

### API Declarations

```
Public Enum enPadString
 pdLeft
 pdRight
End Enum
```


### Source Code

```
Public Function PadString(pstrInput As String, _
 pintWidth As Integer, _
 pstrChar As String, _
 Optional penSidetoPad As enPadString = pdLeft) As String
 'Returns
 '-------
 'PadString("12345", 10, "0") = "0000012345"
 'PadString("12345", 10, "0", pdRight)) = "1234500000"
 'Declare Variables
 '-----------------
 Dim strTemp As String
 '-----------------
 'End Declares
 'Creates a string to the length of
 'pintWidth of the first character
 'of pstrChar.
 strTemp = String$(pintWidth, pstrChar)
 'Check to see what side to pad?
 If penSidetoPad = pdRight Then
  PadString = Left$(pstrInput & strTemp, pintWidth)
 Else
  PadString = Right$(strTemp & pstrInput, pintWidth)
 End If
End Function 'PadString
```

