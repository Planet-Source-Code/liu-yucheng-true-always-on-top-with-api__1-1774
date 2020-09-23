<div align="center">

## True Always On top With API


</div>

### Description

True window always on top. Like MS toolbar.

It floats over any active application is inactive but is still always on top.

Allows you to switch between applications and window is still always on top!

Much better than Jake McCurry's lousy Always on top code.
 
### More Info
 
Window Handle of Window to stay on top.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Liu Yucheng](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/liu-yucheng.md)
**Level**          |Unknown
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/liu-yucheng-true-always-on-top-with-api__1-1774/archive/master.zip)

### API Declarations

```
Public Const HWND_BOTTOM = 1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
```


### Source Code

```
Public Sub MakeWindowAlwaysTop(hwnd As Long)
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub
Public Sub MakeWindowNotTop(hwnd As Long)
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub
```

