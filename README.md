<div align="center">

## Close a window


</div>

### Description

Close a window when you know the title of

this window.

Uses the API FindWindow and PostMessage.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Marcus Schmitt](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/marcus-schmitt.md)
**Level**          |Intermediate
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/marcus-schmitt-close-a-window__1-35934/archive/master.zip)





### Source Code

```
Option Explicit
Private Declare Function PostMessage Lib "user32" _
  Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindow Lib "user32" _
  Alias "FindWindowA" (ByVal szClass$, ByVal szTitle$) As Long
Private Const WM_CLOSE = &H10
Private Sub Command1_Click()
  Dim hWnd, retval As Long
  Dim WinTitle As String
  WinTitle = "Recycle Bin" '<- Title of Window
  hWnd = FindWindow(vbNullString, WinTitle)
  retval = PostMessage(hWnd, WM_CLOSE, 0&, 0&)
End Sub
```

