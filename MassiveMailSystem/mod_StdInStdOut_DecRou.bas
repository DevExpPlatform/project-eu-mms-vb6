Attribute VB_Name = "mod_StdInStdOut_DecRou"
Option Explicit

Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As Any, lpProcessInformation As Any) As Long
Private Declare Function GetNamedPipeInfo Lib "kernel32" (ByVal hNamedPipe As Long, lType As Long, lLenOutBuf As Long, lLenInBuf As Long, lMaxInstances As Long) As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
 
Private Type SECURITY_ATTRIBUTES
    nLength                     As Long
    lpSecurityDescriptor        As Long
    bInheritHandle              As Long
End Type

Private Type STARTUPINFO
    cb                          As Long
    lpReserved                  As Long
    lpDesktop                   As Long
    lpTitle                     As Long
    dwX                         As Long
    dwY                         As Long
    dwXSize                     As Long
    dwYSize                     As Long
    dwXCountChars               As Long
    dwYCountChars               As Long
    dwFillAttribute             As Long
    dwFlags                     As Long
    wShowWindow                 As Integer
    cbReserved2                 As Integer
    lpReserved2                 As Long
    hStdInput                   As Long
    hStdOutput                  As Long
    hStdError                   As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess                    As Long
    hThread                     As Long
    dwProcessID                 As Long
    dwThreadID                  As Long
End Type

'Private Const STD_ERROR_HANDLE  As Long = -12&
'Private Const STD_INPUT_HANDLE  As Long = -10&
Private Const STD_OUTPUT_HANDLE As Long = -11&

Function StdOutRead(sCommandLine As String, Optional bShowWindow As Boolean = False) As String
    
    'Const clReadBytes               As Long = 256
    Const INFINITE As Long = &HFFFFFFFF
    Const STARTF_USESHOWWINDOW = &H1, STARTF_USESTDHANDLES = &H100&
    Const SW_HIDE = 0, SW_NORMAL = 1
    Const NORMAL_PRIORITY_CLASS = &H20&
    
    'Const PIPE_CLIENT_END = &H0
    'Const PIPE_SERVER_END = &H1
    Const PIPE_TYPE_BYTE = &H0
    'Const PIPE_TYPE_MESSAGE = &H4
    
    Dim tProcInfo                   As PROCESS_INFORMATION, lRetVal As Long, lSuccess As Long
    Dim tStartupInf                 As STARTUPINFO
    Dim tSecurAttrib                As SECURITY_ATTRIBUTES, lhwndReadPipe As Long, lhwndWritePipe As Long
    Dim lBytesRead                  As Long, sBuffer As String
    Dim lPipeOutLen                 As Long, lPipeInLen As Long, lMaxInst As Long
    
    tSecurAttrib.nLength = Len(tSecurAttrib)
    tSecurAttrib.bInheritHandle = 1&
    tSecurAttrib.lpSecurityDescriptor = 0&

    lRetVal = CreatePipe(lhwndReadPipe, lhwndWritePipe, tSecurAttrib, 0)
    
    If lRetVal = 0 Then Exit Function

    tStartupInf.cb = Len(tStartupInf)
    tStartupInf.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    tStartupInf.hStdOutput = lhwndWritePipe
    
    If bShowWindow Then
        tStartupInf.wShowWindow = SW_NORMAL
    Else
        tStartupInf.wShowWindow = SW_HIDE
    End If

    lRetVal = CreateProcessA(0&, sCommandLine, tSecurAttrib, tSecurAttrib, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, tStartupInf, tProcInfo)
    
    If lRetVal <> 1 Then Exit Function
    
    WaitForSingleObject tProcInfo.hProcess, INFINITE
    
    lSuccess = GetNamedPipeInfo(lhwndReadPipe, PIPE_TYPE_BYTE, lPipeOutLen, lPipeInLen, lMaxInst)
    
    If lSuccess Then
        sBuffer = String(lPipeOutLen, 0)
        lSuccess = ReadFile(lhwndReadPipe, sBuffer, lPipeOutLen, lBytesRead, 0&)
        
        If lSuccess = 1 Then StdOutRead = Left$(sBuffer, lBytesRead)
    End If
    
    Call CloseHandle(tProcInfo.hProcess)
    Call CloseHandle(tProcInfo.hThread)
    Call CloseHandle(lhwndReadPipe)
    Call CloseHandle(lhwndWritePipe)

End Function

Public Sub StdOutWrite(ByVal Text As String)
    
    Dim hSTDOUT         As Long
    Dim lngBytesWritten As Long
    Dim retval          As Long
    
    hSTDOUT = GetStdHandle(STD_OUTPUT_HANDLE)
    retval = WriteFile(hSTDOUT, ByVal Text, Len(Text), lngBytesWritten, ByVal 0&)
    retval = CloseHandle(hSTDOUT)

End Sub

