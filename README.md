# KeyboardCOM
This application serves as a simple example of global keyboard hooking within COM rules. This is a multithreaded solution to the single threaded Visual Basic for Applications (VBA) implementation. This class library is built specifically for Excel with the goal of allowing users to run subroutines (macros) when a user-determined virtual-keycode is pressed.

## Using
- To use this application in Excel, import the type library 'KeyboardCOM.tlb' in VBA References
- While not illegal, multiple `KeyboardHook` objects should be avoided
- Virtual keycodes are found here: https://docs.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes
- Declaration, initialization, and usage is below:

```vb.net
Dim kbHook As New KeyboardHook

Sub startKeyboardHook()
    kbHook.addHotkey ActiveWorkbook, &H52, "getSetCase" 'r
    kbHook.addHotkey ActiveWorkbook, &H51, "updateCWS" 'q
    kbHook.addHotkey ActiveWorkbook, &H58, "getSetCaseX" 'x
    kbHook.addHotkey ActiveWorkbook, &H26, "cwsNavTool_moveUp" 'up arrow
    kbHook.addHotkey ActiveWorkbook, &H28, "cwsNavTool_moveDown" 'down arrow
    kbHook.addHotkey ActiveWorkbook, &H27, "cwsNavTool_GetSetCase" 'right arrow

    kbHook.StartHook
End Sub

Sub stopKeyboardHook()
    kbHook.StopHook
End Sub

Sub getSetCase()
  ...
End Sub
```

## Building
- This application is built using .NET Framework 3.5
- This application has the following references:
  - WindowsBase
  - System.Windows.Forms
  - Microsoft Office 16.0 Object Library (COM)
  - Microsoft Visual Basic for Applications Extensibility 5.3 (COM)
  - Microsoft Excel 16.0 Object Library (COM)
