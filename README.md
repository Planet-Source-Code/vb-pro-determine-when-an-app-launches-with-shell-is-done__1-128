<div align="center">

## Determine when an app launches with SHELL is done


</div>

### Description

In VB3, you call GetModuleUsage() to determine when an app you started with the Shell command was complete. However, this call does not work correctly in the 32-bit arena of Windows NT and Windows 95.

To overcome this obstacle, use a routine in both 16- and 32- bit environments that will tell you when a program has finished, even if it does not create a window.

The IsInst() routine uses the TaskFirst and TaskNext functions defined in the TOOLHELP.DLL to see if the instance handle returned by the Shell function is still valid. When IsInst() returns False, the command has finished.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VB Pro](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vb-pro.md)
**Level**          |Unknown
**User Rating**    |3.8 (15 globes from 4 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vb-pro-determine-when-an-app-launches-with-shell-is-done__1-128/archive/master.zip)





### Source Code

```
hInst = Shell("foobar.exe")
Do While IsInst(hInst)
DoEvents
Loop
Function IsInst(hInst As Integer) As Boolean
Dim taskstruct As TaskEntry
Dim retc As Boolean
IsInst = False
taskstruct.dwSize = Len(taskstruct)
retc = TaskFirst(taskstruct)
Do While retc
If taskstruct.hInst = hInst Then
' note: the task handle is: taskstruct.hTask
IsInst = True
Exit Function
End If
retc = TaskNext(taskstruct)
Loop
End Function
```

