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

    Public Sub AddHotkey(workbook As Workbook, vk As Integer, functionName As String, argsSeperatedByComma As String)
        Dim hotkey As New HotKey With {
            .workbook = workbook,
            .vk = vk,
            .functionName = functionName,
            .argsSeperatedByComma = argsSeperatedByComma,
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
        Dim msg As New MSG

        For Each hotkey In hotKeyList
            If RegisterHotKey(IntPtr.Zero, hotkey.id, 0, hotkey.vk) Then
                Debug.WriteLine("Hotkey with id" & hotkey.id & " successfully registerd.")
                Debug.WriteLine("\t- Key: " & keyConverter.ConvertToString(hotkey.vk))
                Debug.WriteLine("\t- Function name: " & hotkey.functionName)
                Debug.WriteLine("\t- Arg string: " & hotkey.functionName)
            Else
                Debug.WriteLine("[ERROR] On hotkey register, Error number: " & Err.LastDllError & ", Key: " & keyConverter.ConvertToString(hotkey.vk) & ", ID: " & hotkey.id)
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
                        Dim argCount As Integer
                        Dim args(29) As String
                        Dim i As Integer
                        i = 0

                        If hotkey.argsSeperatedByComma = "" Then
                            argCount = 0
                        Else
                            If Not hotkey.argsSeperatedByComma.Contains(","c) Then
                                argCount = 0
                                args(0) = hotkey.argsSeperatedByComma
                            Else
                                Try
                                    args = Split(hotkey.argsSeperatedByComma, ",", -1)
                                    argCount = args.Length - 1

                                    For Each arg As String In args
                                        args(i) = arg
                                        i += 1
                                    Next
                                Catch ex As Exception
                                    Debug.WriteLine("Error on arg parse")
                                    Throw New Exception(ex.Message)
                                End Try
                            End If
                        End If

                        Try
                            Select Case argCount
                                Case 0
                                    hotkey.workbook.Application.Run(hotkey.functionName)
                                Case 1
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0))
                                Case 2
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1))
                                Case 3
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2))
                                Case 4
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3))
                                Case 5
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4))
                                Case 6
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5))
                                Case 7
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5), args(6))
                                Case 8
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7))
                                Case 9
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8))
                                Case 10
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9))
                                Case 11
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10))
                                Case 12
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11))
                                Case 13
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12))
                                Case 14
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13))
                                Case 15
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14))
                                Case 16
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15))
                                Case 17
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16))
                                Case 18
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17))
                                Case 19
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18))
                                Case 20
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19))
                                Case 21
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20))
                                Case 22
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21))
                                Case 23
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22))
                                Case 24
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23))
                                Case 25
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24))
                                Case 26
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25))
                                Case 27
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26))
                                Case 28
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26), args(27))
                                Case 29
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26), args(27), args(28))
                                Case 30
                                    hotkey.workbook.Application.Run(hotkey.functionName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26), args(27), args(28), args(29))

                            End Select

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