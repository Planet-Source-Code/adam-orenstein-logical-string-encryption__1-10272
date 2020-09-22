<div align="center">

## Logical String Encryption


</div>

### Description

This will logically encrypt a string by deviating its ASC() value. Works perfectly. Its just a VERY Short Algoritm the String gets passed through Charachter by Character. To decrypt just redo the encryption. just like a double negative.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Adam Orenstein](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/adam-orenstein.md)
**Level**          |Advanced
**User Rating**    |3.2 (16 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Encryption](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/encryption__1-48.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/adam-orenstein-logical-string-encryption__1-10272/archive/master.zip)





### Source Code

```
Public Function EncodeText(TheText As String) As String
Dim Letter As String
Dim TextLen As Integer
Dim Crypt As Double
  TextLen = Len(TheText)
  For Crypt = 1 To TextLen
    Letter = Asc(Mid(TheText, Crypt, 1))
    Letter = Letter Xor 255
    Result$ = Result$ & Chr(Letter)
  Next Crypt
  EncodeText = Result$
End Function
```

