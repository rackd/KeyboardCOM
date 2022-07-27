Imports System.Threading
Imports System.Windows.Forms
Imports System.Windows.Interop
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Win32
Imports KeyboardCOM.MethodsClass

<ComClass(ClassId, InterfaceId, EventsId)>
Public Class KeyboardHook
    ReadOnly keyConverter As New KeysConverter
    Private hotKeyID As Integer
    Private childThreadID As Integer
    Private hotKeyList As New List(Of HotKey)

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
        If hotKeyList Is Nothing Then
            Debug.WriteLine("[ERROR] Could not start... no hotkeys added")
            Exit Sub
        End If

        SendQuitMessagesFromReg()

        Call New Thread(Sub() Start()).Start()
    End Sub

    Public Sub StopHook()
        SendQuitMessagesFromReg()
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
        Debug.WriteLine(vbNewLine & "Successfully registered all hotkeys on thread ID: " & childThreadID)
        AddTIDToReg(childThreadID)

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
        Debug.WriteLine("Keyboard hook thread with ID " & childThreadID & " successfully exited")
    End Sub

    Private Sub resetHook()
        hotKeyID = 0
        childThreadID = 0
        hotKeyList = Nothing
        SendQuitMessagesFromReg()
    End Sub

    Private Sub SendQuitMessagesFromReg()
        Registry.CurrentUser.CreateSubKey(TID_REGISTRY_DIR)

        For Each TID As String In Registry.CurrentUser.OpenSubKey(TID_REGISTRY_DIR).GetValueNames
            Debug.WriteLine("Found thread, sending quit request to thread with ID: " & TID)

            PostThreadMessage(CInt(TID), WM_QUIT, 0, 0)

            Registry.CurrentUser.CreateSubKey(TID_REGISTRY_DIR).DeleteValue(CInt(TID))
        Next
    End Sub

    Private Sub AddTIDToReg(TID As String)
        Registry.CurrentUser.CreateSubKey(TID_REGISTRY_DIR).SetValue(TID, "")
    End Sub
End Class
