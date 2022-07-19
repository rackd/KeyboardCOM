Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows.Forms
Imports System.Windows.Interop
Imports Microsoft.Office.Interop.Excel

<ComClass(KeyboardHook.ClassId, KeyboardHook.InterfaceId, KeyboardHook.EventsId)>
Public Class KeyboardHook

#Region "COM GUIDs"
    Public Const ClassId As String = "2adca793-deab-4725-a2d0-a0c3aa168847"
    Public Const InterfaceId As String = "5018979f-94fd-4c92-a878-70f2550e9b9f"
    Public Const EventsId As String = "22c3305d-8d09-43ad-a629-7a496fca75e4"
#End Region

#Region "Windows API functions"
    <DllImport("User32.dll")>
    Public Shared Function RegisterHotKey(ByVal hwnd As IntPtr, ByVal id As Integer, ByVal fsModifiers As Integer, ByVal vk As Integer) As Integer : End Function
    <DllImport("User32.dll")>
    Public Shared Function GetMessage(ByRef lpMsg As MSG, ByVal hWnd As IntPtr, ByVal wMsgFilterMin As UInteger, ByVal wMsgFilterMax As UInteger) As <MarshalAs(UnmanagedType.Bool)> Boolean : End Function
    <DllImport("User32.dll")>
    Public Shared Function PostThreadMessage(ByVal id As Integer, ByVal msg As Integer, ByVal wparam As IntPtr, ByVal lparam As IntPtr) As Integer : End Function
    <DllImport("kernel32.dll")>
    Public Shared Function GetCurrentThreadId() As UInteger : End Function
#End Region

    ReadOnly keyConverter As New KeysConverter

    Private Const WM_HOTKEY As Integer = &H312
    Private Const WM_QUIT As Integer = &H12

    Private hotKeyID As Integer
    Private childThreadID As Integer
    Private hotKeyList As New List(Of HotKey)
    Private hookStarted As Boolean

    Private Structure HotKey
        Dim workbook As Workbook
        Dim vk As Integer
        Dim functionName As String
        Dim id As Integer
    End Structure

    Public Sub AddHotkey(workbook As Workbook, vk As Integer, functionName As String)
        Dim hotkey As New HotKey With {
            .workbook = workbook,
            .vk = vk,
            .functionName = functionName,
            .id = hotKeyID
        }

        hotKeyList.Add(hotkey)
        hotKeyID += 1
    End Sub

    Public Sub StartHook()
        If hookStarted Then
            Debug.WriteLine("[ERROR] Could not start... keyboard hook already started")
            Exit Sub
        End If

        If hotKeyList Is Nothing Then
            Debug.WriteLine("[ERROR] Could not start... no hotkeys added")
            Exit Sub
        End If

        Call New Thread(Sub() Start()).Start()
    End Sub

    Public Sub StopHook()
        If Not hookStarted Then
            Debug.WriteLine("[ERROR] Could not stop... hooked not started")
            Exit Sub
        End If

        If childThreadID = 0 Then
            Debug.WriteLine("[ERROR] Could not stop... child thread ID uninitialized")
            Exit Sub
        End If

        PostThreadMessage(childThreadID, WM_QUIT, 0, 0)
        Debug.WriteLine("Sent terminatation request to child thread message queue with ID: " & childThreadID)
    End Sub

    Private Sub Start()
        Dim _workbook As Workbook
        Dim _vk As Integer
        Dim _functionName As String
        Dim _id As Integer
        Dim msg As New MSG

        For Each hotkey In hotKeyList
            _workbook = hotkey.workbook
            _vk = hotkey.vk
            _functionName = hotkey.functionName
            _id = hotkey.id

            If RegisterHotKey(IntPtr.Zero, _id, 0, _vk) Then
                Debug.WriteLine("Hotkey registered, Key: " & keyConverter.ConvertToString(_vk) & ", ID: " & _id)
            Else
                Debug.WriteLine("[ERROR] On hotkey register, Error number: " & Err.LastDllError & ", Key: " & keyConverter.ConvertToString(_vk) & ", ID: " & _id)
                Debug.WriteLine("Could not start...")
                resetHook()
                Exit Sub
            End If
        Next

        childThreadID = GetCurrentThreadId()
        hookStarted = True
        Debug.WriteLine(vbNewLine & "Successfully registered all hotkeys on thread ID: " & childThreadID)

        Do While GetMessage(msg, IntPtr.Zero, 0, 0) <> 0
            If msg.message = WM_HOTKEY Then
                For Each hotkey In hotKeyList
                    If msg.wParam = hotkey.id Then
                        Try
                            hotkey.workbook.Application.Run(hotkey.functionName)
                            Debug.WriteLine("Ran function '" & hotkey.functionName & "'")
                        Catch ex As Exception
                            Debug.WriteLine("[Error] VBA error detected... stopping hook, Error message: " & ex.Message)
                            Exit Do
                        End Try
                    End If
                Next
            End If
        Loop

        resetHook()
        Debug.WriteLine("Keyboard hook thread successfully exited")
    End Sub

    Private Sub resetHook()
        hotKeyID = 0
        childThreadID = 0
        hotKeyList = Nothing
        hookStarted = False
    End Sub
End Class
