Attribute VB_Name = "ShutSystem"
Option Explicit

Public ex As String
Public ex2 As String
Public ex3 As String
Public ex4 As String

Private Type LUID
    UsedPart As Long
    IgnoredForNowHigh32BitPart As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Const VER_PLATFORM_WIN32_NT = 2&
Private Const STATUS_TIMEOUT = &H102&
Private Const INFINITE = -1&
Private Const QS_KEY = &H1&
Private Const QS_MOUSEMOVE = &H2&
Private Const QS_MOUSEBUTTON = &H4&
Private Const QS_POSTMESSAGE = &H8&
Private Const QS_TIMER = &H10&
Private Const QS_PAINT = &H20&
Private Const QS_SENDMESSAGE = &H40&
Private Const QS_HOTKEY = &H80&
Private Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT Or _
                            QS_TIMER Or QS_POSTMESSAGE Or _
                            QS_MOUSEBUTTON Or QS_MOUSEMOVE Or _
                            QS_HOTKEY Or QS_KEY)
                            

Private Const EWX_FORCESHUTDOWN = 5&
Private Const WM_CLOSE = &H10&
Private Const WM_QUERYENDSESSION = &H11&
Private Const WM_ENDSESSION = &H16&
Private Const ENDSESSION_LOGOFF = &H80000000
Private Const SMTO_ABORTIFHUNG = &H2&
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PROCESS_TERMINATE = 1&
Private Const GWL_STYLE = -16&
Private Const GWL_EXSTYLE = -20&
Private Const WS_POPUP = &H80000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_DISABLED = &H8000000
Private Const WS_EX_TOOLWINDOW = &H80&
Private Const TOKEN_ADJUST_PRIVILEGES = &H20&
Private Const TOKEN_QUERY = &H8&
Private Const SE_PRIVILEGE_ENABLED = &H2&
Private Const SE_SHUTDOWN_NAME = "seShutdownPrivilege"

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
                (ByVal hWnd As Long, ByVal wMsg As Long, _
                ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function EnumWindows Lib "user32" _
                (ByVal lpEnumFunc As Long, _
                ByVal lParam As Long) As Long
Private Declare Function SendMessageTimeout Lib "user32" _
                Alias "SendMessageTimeoutA" (ByVal hWnd As Long, _
                ByVal msg As Long, ByVal wParam As Long, _
                ByVal lParam As Long, ByVal fuFlags As Long, _
                ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" _
                (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" _
                (ByVal dwDesiredAccess As Long, _
                ByVal bInheritHandle As Long, _
                ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" _
                (ByVal hProcess As Long, _
                ByVal uExitCode As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function CloseHandle Lib "kernel32" _
                (ByVal hObject As Long) As Long
Private Declare Function GetWindowLong Lib "user32" _
                Alias "GetWindowLongA" (ByVal hWnd As Long, _
                ByVal nIndex As Long) As Long
Private Declare Function GetClassName Lib "user32" _
                Alias "GetClassNameA" (ByVal hWnd As Long, _
                ByVal lpClassName As String, _
                ByVal nMaxCount As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32" _
                (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
                (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function MsgWaitForMultipleObjects Lib "user32" _
                (ByVal nCount As Long, pHandles As Long, _
                ByVal fWaitAll As Long, ByVal dwMilliseconds As Long, _
                ByVal dwWakeMask As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" _
                (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, _
                TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" _
                Alias "LookupPrivilegeValueA" _
                (ByVal lpSystemName As String, _
                ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" _
                (ByVal TokenHandle As Long, _
                ByVal DisableAllPrivileges As Long, _
                NewState As TOKEN_PRIVILEGES, _
                ByVal BufferLength As Long, _
                PreviousState As TOKEN_PRIVILEGES, _
                ReturnLength As Long) As Long

Enum WindowsToClose
    wtcAll
    wtcVisible
    wtcProgram
    wtcPopup
End Enum

Enum ShutError
    seNone
    sePrivileges
    seShutdown
End Enum

Dim PIds() As Long, PIdsCount As Long

' Shutdown function performs OS shutting down.
' Return values:
' sePrivileges1 - insufficient rights,
' seShutdown1 - unable to shut down,
' seNone - without errors (never returns)
 
Public Function ShutDown() As ShutError
Dim i As Long, hProcess As Long
Dim OSVer As OSVERSIONINFO
ShutDown = seNone
' Checking OS type
OSVer.dwOSVersionInfoSize = 148&
GetVersionEx OSVer
If OSVer.dwPlatformId = VER_PLATFORM_WIN32_NT Then
' Adjusting privileges
    AdjustToken
    If Err.LastDllError <> 0& Then
        ShutDown = sePrivileges
        Exit Function
    End If
End If
PIdsCount = 0&
ReDim PIds(1& To 1&)
' Closing all PopUp-windows
EnumWindows AddressOf EnumWindows0Proc, 0&
MsgWaitObj 1000&
' Closing all program windows
EnumWindows AddressOf EnumWindows1Proc, 0&
MsgWaitObj 8000&
' Notifying all remaining windows about OS shutting down
EnumWindows AddressOf EnumWindows2Proc, 0&
MsgWaitObj 6000&
' Closing all appeared PopUp-windows
EnumWindows AddressOf EnumWindows0Proc, 0&
MsgWaitObj 1000&
' Terminating all remaining applications
EnumWindows AddressOf EnumWindows3Proc, 0&
MsgWaitObj 4000&
' Shutting down the system
ExitWindowsEx EWX_FORCESHUTDOWN, 0&
' This function never returns
If Err.LastDllError <> 0& Then ShutDown = seShutdown
End Function

' EnumWindows0Proc function closes specified window
' by sending it WM_CLOSE message, if it is PopUp-window

Private Function EnumWindows0Proc(ByVal hWnd As Long, _
                ByVal lParam As Long) As Long
If WindowToClose(hWnd, wtcPopup) Then
    PostMessage hWnd, WM_CLOSE, 0&, 0&
End If
EnumWindows0Proc = 1&
End Function

' EnumWindows1Proc function closes specified window
' by sending it WM_CLOSE message, if it is main window of application

Private Function EnumWindows1Proc(ByVal hWnd As Long, _
                ByVal lParam As Long) As Long
If WindowToClose(hWnd, wtcProgram) Then
    PostMessage hWnd, WM_CLOSE, 0&, 0&
End If
EnumWindows1Proc = 1&
End Function

' EnumWindows2Proc function sends to specified window, if it is main
' window of application, WM_QUERYENDSESSION message. If it returns
' True, it sends also WM_ENDSESSION message.

Private Function EnumWindows2Proc(ByVal hWnd As Long, ByVal lParam As Long) As Long
Dim r As Long, r1 As Long
If WindowToClose(hWnd, wtcProgram) Then
    r = SendMessageTimeout(hWnd, WM_QUERYENDSESSION, 0&, ENDSESSION_LOGOFF, SMTO_ABORTIFHUNG, 5000&, r1)
    If r <> 0& Then PostMessage hWnd, WM_ENDSESSION, r1, ENDSESSION_LOGOFF
End If
EnumWindows2Proc = 1&
End Function

' EnumWindows3Proc terminates process, which is binded with window,
' if it is visible. One process never terminates twice.

Private Function EnumWindows3Proc(ByVal hWnd As Long, ByVal lParam As Long) As Long
Dim hProcess As Long, PId As Long, i As Long
EnumWindows3Proc = 1&
If WindowToClose(hWnd, wtcVisible) Then
    GetWindowThreadProcessId hWnd, PId
    For i = 1& To PIdsCount
        If PId = PIds(i) Then
            Exit Function
        End If
    Next
    PIdsCount = PIdsCount + 1&
    ReDim Preserve PIds(1& To PIdsCount)
    PIds(PIdsCount) = PId
    hProcess = OpenProcess(STANDARD_RIGHTS_REQUIRED Or PROCESS_TERMINATE, 0&, PId)
    If Err.LastDllError = 0& Then
        TerminateProcess hProcess, 0&
        CloseHandle hProcess
    End If
End If
End Function

' Ôóíêöèÿ WindowClassname âîçâðàùàåò èìÿ êëàññà îêíà.

' WindowClassname function returns window's class name.

Private Function WindowClassname(hWnd As Long) As String
WindowClassname = String$(255&, 0)
WindowClassname = Left$(WindowClassname, GetClassName(hWnd, WindowClassname, 255&))
End Function

' WindowToClose function returns True if specified window
' corresponds to specified criteria: it is visible or it is main
' application's window or it is PopUp-window.
' For some window classes (which belong to Explorer, Progman,
' WinPopup, Task Bar) this function always returns False.
' It is because the safest way to terminate these applications
' is to allow them to terminate with operating system.

Private Function WindowToClose(hWnd As Long, wtc As WindowsToClose) As Boolean
Dim PId As Long, v As Boolean, St As Long, ExSt As Long
St = GetWindowLong(hWnd, GWL_STYLE)
ExSt = GetWindowLong(hWnd, GWL_EXSTYLE)
Select Case wtc
    Case wtcVisible
        v = (St And WS_VISIBLE) <> 0&
    Case wtcProgram
        v = ((ExSt And WS_EX_TOOLWINDOW) = 0&) And ((St And WS_POPUP) = 0) And ((St And WS_VISIBLE) <> 0& And ((St And WS_DISABLED) = 0&))
    Case wtcPopup
        v = ((St And WS_POPUP) <> 0&) And ((St And WS_VISIBLE) <> 0) And ((St And WS_DISABLED) = 0&) And ((ExSt And WS_EX_TOOLWINDOW) = 0&)
    Case Else
        v = True
End Select
If v Then
    Select Case WindowClassname(hWnd)
        Case "ExploreWClass", "IEFrame", "IEDummyFrame", "Shell_TrayWnd", "Progman", "CabinetWClass", "BaseBar", "WinPopup"
        Case Else
            GetWindowThreadProcessId hWnd, PId
            If PId <> GetCurrentProcessId() Then
                WindowToClose = True
            End If
    End Select
End If
End Function

' AdjustToken() procedure adjusts privileges
' to shut down Windows NT/2000.

Private Sub AdjustToken()
Dim hdlProcessHandle As Long
Dim hdlTokenHandle As Long
Dim tmpLuid As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long

    hdlProcessHandle = GetCurrentProcess()
    OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
    TOKEN_QUERY), hdlTokenHandle

    'Get the LUID for shutdown privilege.
    LookupPrivilegeValue vbNullString, SE_SHUTDOWN_NAME, tmpLuid

    'One privilege to set
    tkp.PrivilegeCount = 1&
    tkp.TheLuid = tmpLuid
    tkp.Attributes = SE_PRIVILEGE_ENABLED

    'Enable the shutdown privilege in the access token of this
    'process.

    AdjustTokenPrivileges hdlTokenHandle, 0&, tkp, _
    Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded

End Sub

' MsgWaitObj function replaces Sleep, WaitForSingleObject,
' WaitForMultipleObjects API functions. Unlike mentioned functions,
' this function doesn't block processing the thread messages,
' which allows to work with COM objects, in particular forms,
' during wait period.
' Using instead Sleep:
'    MsgWaitObj dwMilliseconds
' Using instead WaitForSingleObject:
'    retval = MsgWaitObj(dwMilliseconds, hObj, 1&)
' Using instead WaitForMultipleObjects:
'    retval = MsgWaitObj(dwMilliseconds, hObj(0&), n), where n - number
' of waitable objects, hObj() - array of their handles.

Public Function MsgWaitObj(Interval As Long, Optional hObj As Long = 0, Optional nObj As Long = 0&) As Long
Dim T As Long, T1 As Long
If Interval <> INFINITE Then
    T = GetTickCount()
    On Error Resume Next
    T = T + Interval
' Overflow correction
    If Err <> 0& Then
        If T > 0& Then
            T = ((T + &H80000000) + Interval) + &H80000000
        Else
            T = ((T - &H80000000) + Interval) - &H80000000
        End If
    End If
    On Error GoTo 0
Else
    T1 = INFINITE
End If
Do
    If Interval <> INFINITE Then
        T1 = GetTickCount()
        On Error Resume Next
        T1 = T - T1
' Overflow correction
        If Err <> 0& Then
            If T > 0& Then
                T1 = ((T + &H80000000) - (T1 - &H80000000))
            Else
                T1 = ((T - &H80000000) - (T1 + &H80000000))
            End If
        End If
        On Error GoTo 0
        If IIf((T1 Xor Interval) > 0&, T1 > Interval, T1 < 0&) Then
            MsgWaitObj = STATUS_TIMEOUT
            Exit Function
        End If
    End If
    MsgWaitObj = MsgWaitForMultipleObjects(nObj, hObj, 0&, T1, QS_ALLINPUT)
    DoEvents
    If MsgWaitObj <> nObj Then Exit Function
Loop
End Function
