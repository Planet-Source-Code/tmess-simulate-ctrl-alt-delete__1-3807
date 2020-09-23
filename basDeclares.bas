Attribute VB_Name = "basDeclares"
Option Explicit

Private PID As Long
Public IsResond As String

Public Const PROCESS_ALL_ACCESS = &H1F0FFF

Public Declare Function OpenProcess Lib "kernel32" _
    (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long
    
Public Declare Function TerminateProcess Lib "kernel32" _
    (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
    
Private Const WM_NULL = &H0
Private Const SMTO_BLOCK = &H1
Private Const SMTO_ABORTIFHUNG = &H2

Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, _
    ByVal lParam As Long) As Long
    
Private Declare Function GetWindowThreadProcessId Lib "user32" _
    (ByVal hwnd As Long, lpdwProcessId As Long) As Long
    
Private Declare Function SendMessageTimeout Lib "user32" _
    Alias "SendMessageTimeoutA" (ByVal hwnd As Long, _
    ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, _
    ByVal fuFlags As Long, ByVal uTimeout As Long, _
    pdwResult As Long) As Long

Private Function fEnumWindowsCallBack(ByVal hwnd As Long, ByVal lpData As Long) As Long
Dim lThreadId  As Long
Dim lProcessId As Long
'
' This callback function is called by Windows (from the EnumWindows
' API call) for EVERY window that exists until fEnumWindowsCallBack
' is set False.
'
fEnumWindowsCallBack = 1
lThreadId = GetWindowThreadProcessId(hwnd, lProcessId)

If lProcessId = PID Then
    Call strCheck(hwnd)
    fEnumWindowsCallBack = 0
End If

End Function

Public Function fEnumWindows(clsPID As Long) As Boolean
Dim hwnd As Long

PID = clsPID

' The EnumWindows function enumerates all top-level windows
' on the screen by passing the handle of each window, in turn,
' to an application-defined callback function. EnumWindows
' continues until the last top-level window is enumerated or
' the callback function returns FALSE.
'
 Call EnumWindows(AddressOf fEnumWindowsCallBack, hwnd)
End Function
    

Private Function strCheck(ByVal lhwnd As Long)
Dim lResult As Long
Dim lReturn As Long
Dim strRunning As String

' If no app started, get out.
'
If lhwnd = 0 Then Exit Function
'
' Check the status of the application specifying
' a timeout period of 1 second (1000 miliseconds).
'
' SMTO_ABORTIFHUNG Returns without waiting for the
'       time-out period to elapse if the receiving
'       process appears to be in a "hung" state.
'
' SMTO_BLOCK Prevents the calling thread from processing
'       any other requests until the function returns.
'
lReturn = SendMessageTimeout(lhwnd, WM_NULL, 0&, 0&, SMTO_ABORTIFHUNG And SMTO_BLOCK, 1000, lResult)

If lReturn Then
    IsResond = "Responding"
Else
    IsResond = "Not Responding"
End If
End Function


