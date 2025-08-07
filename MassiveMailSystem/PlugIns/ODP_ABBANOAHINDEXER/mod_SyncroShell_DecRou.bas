Attribute VB_Name = "mod_SyncroShell_DecRou"
Option Explicit

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
Private Const STARTF_USESHOWNWINDOW = &H1&
Private Const SW_HIDE = 0

Public Function ExecuteAndWait(ByVal CmdLine As String) As Boolean

    Dim proc    As PROCESS_INFORMATION
    Dim start   As STARTUPINFO
    Dim ret     As Long

    ' Initialize the STARTUPINFO structure:
    start.cb = Len(start)
    start.dwFlags = STARTF_USESHOWNWINDOW 'Necessary for wShowWindow to work
    start.wShowWindow = SW_HIDE 'Hide window

    ' Start the shelled application:
    ret& = CreateProcessA(0&, CmdLine$, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)

    ' Wait for the shelled application to finish:
    ret& = WaitForSingleObject(proc.hProcess, INFINITE)
    ret& = CloseHandle(proc.hProcess)

End Function

'Public Function Get_ShortPath(LongPath As String) As String
'
'    Dim S As String
'    Dim I As Long
'    Dim PathLength As Long
'
'    I = Len(LongPath) + 1
'    S = String$(I, 0)
'    PathLength = GetShortPathName(LongPath, S, I)
'
'    Get_ShortPath = Left$(S, PathLength)
'
'End Function

'Public Sub Kill_Application(ByVal AppName As String)
'
'    Dim hWindow         As Long
'    Dim lngResult       As Long
'    Dim lngReturnValue  As Long
'
'    hWindow = FindWindow(vbNullString, AppName)
'    lngReturnValue = PostMessage(hWindow, WM_CLOSE, vbNull, vbNull)
'    lngResult = WaitForSingleObject(hWindow, INFINITE)
'
'    Wait 2
'
'    hWindow = FindWindow(vbNullString, "Adobe Reader")
'
'    If IsWindow(hWindow) = 1 Then
'        MsgBox AppName & " handle still exists."
'    'Else
'    '    MsgBox "Program " & AppName & " closed."
'    End If
'
'End Sub

