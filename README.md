<div align="center">

## Timeout/Pause


</div>

### Description

You can pause execution of code for the specified duration. Different from "Sleep" api in that it will not lock up the whole program.
 
### More Info
 
Duration - Specify the seconds you want to pause execution for

GetTickCount(api), like Timer resets at some point. Timer resets at midnight, and returns the seconds since midnight. GetTickCount returns the ticks (milliseconds) since the o/s was started. I remember reading that it resets to 0 after 49.7 days (different o/s may vary [windows]). I don't think it should cause a problem, but it may result in the loop, looping indefinitely until the program is shut down. This is slightly more accurate than timeout/pause routines that use Timer.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Xeek](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/xeek.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/xeek-timeout-pause__1-9633/archive/master.zip)

### API Declarations

```
Declare Function GetTickCount Lib "kernel32" Alias "GetTickCount" () As Long
```


### Source Code

```
Sub Pause(Duration As Double)
'example: Pause (0.8) 'pause for .8 seconds
Dim start As Double 'declare variable
  start# = GetTickCount 'store milliseconds since boot
  Do: DoEvents 'start loop
On Error Resume Next 'dunno, kept giving me an error once. so i put this here and it stopped giving me the error
  Loop Until GetTickCount - start# >= (Duration# * 1000) 'loop until the actual time (minus stored time) is greater than or equal to the duration (seconds * 1000 = milliseconds)
End Sub
```

