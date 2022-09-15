Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Interop
Imports Microsoft.Office.Interop.Excel

Friend Class MethodsClass
    Public Const ClassId As String = "2adca793-deab-4725-a2d0-a0c3aa168847"
    Public Const InterfaceId As String = "5018979f-94fd-4c92-a878-70f2550e9b9f"
    Public Const EventsId As String = "22c3305d-8d09-43ad-a629-7a496fca75e4"

    Public Const TID_REGISTRY_DIR = "Software\KeyboardCOM\PIDs"

    Public Const WM_HOTKEY As Integer = &H312
    Public Const WM_QUIT As Integer = &H12

    <DllImport("User32.dll")> Public Shared Function RegisterHotKey(ByVal hWnd As IntPtr, ByVal id As Integer, ByVal fsModifiers As Integer, ByVal vk As Integer) As Integer : End Function
    <DllImport("User32.dll")> Public Shared Function GetMessage(ByRef lpMsg As MSG, ByVal hWnd As IntPtr, ByVal wMsgFilterMin As UInteger, ByVal wMsgFilterMax As UInteger) As <MarshalAs(UnmanagedType.Bool)> Boolean : End Function
    <DllImport("User32.dll")> Public Shared Function PostThreadMessage(ByVal id As Integer, ByVal msg As Integer, ByVal wparam As IntPtr, ByVal lparam As IntPtr) As Integer : End Function
    <DllImport("user32.dll")> Public Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As IntPtr, ByVal lParam As StringBuilder) As IntPtr : End Function
    <DllImport("kernel32.dll")> Public Shared Function GetCurrentThreadId() As UInteger : End Function

    Public Structure HotKey
        Dim workbook As Workbook
        Dim vk As Integer
        Dim functionName As String
        Dim argsSeperatedByComma As String
        Dim id As Integer
    End Structure
End Class