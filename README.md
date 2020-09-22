<div align="center">

## Dials a telephone number w/o using APIs


</div>

### Description

Uses the MSComm control to call a telephone number using your modem WITHOUT HAVEING DIALER.EXE! :)
 
### More Info
 
num - the telephone number

Assumes you have a MSComm control on your form named "Communications"


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[alex](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/alex.md)
**Level**          |Unknown
**User Rating**    |4.7 (28 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/alex-dials-a-telephone-number-w-o-using-apis__1-2120/archive/master.zip)





### Source Code

```
Private Sub Dial(num As String)
 ' Open the com port.
 Communications.PortOpen = True
 ' Send the attention command to the modem.
 Communications.Output = "AT" + Chr$(13)
 ' Wait for processing.
 Do
  DoEvents
  Loop Until Communications.InBufferCount >= 2
  ' Dial the number.
  Communications.Output = "ATDT " + num + Chr$(13)
  ' Takes about 47 sec. to dial
  wait = Timer + 47
  Do
   DoEvents
   Loop While Timer <= wait
   ' Uncomment to disconnect after dialing.
   'Communications.PortOpen = False
  End Sub
```

