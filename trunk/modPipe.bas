Attribute VB_Name = "modPipe"
Private Declare Function GetExitCodeProcess Lib "KERNEL32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Option Explicit

Private Declare Function CreatePipe Lib "KERNEL32" ( _
    phReadPipe As Long, _
    phWritePipe As Long, _
    lpPipeAttributes As Any, _
    ByVal nSize As Long) As Long

Private Declare Function ReadFile Lib "KERNEL32" ( _
    ByVal hFile As Long, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
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
    dwProcessId As Long
    dwThreadId As Long
End Type

Private Declare Function CreateProcessA Lib "KERNEL32" (ByVal _
   lpApplicationName As Long, ByVal lpCommandLine As String, _
   lpProcessAttributes As Any, lpThreadAttributes As Any, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As Any, lpProcessInformation As Any) As Long

Private Declare Function WaitForSingleObject Lib "KERNEL32" _
    (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "KERNEL32" (ByVal _
   hObject As Long) As Long

Const SW_SHOWMINNOACTIVE = 7
Const STARTF_USESHOWWINDOW = &H1
Const INFINITE = -1&
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const STARTF_USESTDHANDLES = &H100&

Private Const STILL_ACTIVE = &H103

' to execute the browser
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOWDEFAULT = 10
'delay function
Private Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)

Public strPipeData As String

Public Function ExecCmdPipe(ByVal CmdLine As String) As String
    'Executes the command, and when it finish returns value to VB

    Dim proc As PROCESS_INFORMATION, ret As Long, bSuccess As Long
    Dim start As STARTUPINFO
    Dim sa As SECURITY_ATTRIBUTES
    Dim hReadPipe As Long, hWritePipe As Long
    Dim bytesread As Long, mybuff As String
    Dim i As Integer
    
    Dim sReturnStr As String
    
    ' the lenght of the string must be 10 * 1024
    
    mybuff = String(10 * 1024, Chr$(65))
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1&
    sa.lpSecurityDescriptor = 0&
    ret = CreatePipe(hReadPipe, hWritePipe, sa, 0)
    If ret = 0 Then
        '===Error
        ExecCmdPipe = "Error: CreatePipe failed. " & Err.LastDllError
        Exit Function
    End If
    start.cb = Len(start)
    start.hStdOutput = hWritePipe
    start.dwFlags = STARTF_USESTDHANDLES + STARTF_USESHOWWINDOW
    start.wShowWindow = 0
    
    ' Start the shelled application:
    ret& = CreateProcessA(0&, CmdLine$, sa, sa, 1&, _
        NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    If ret <> 1 Then
        '===Error
        sReturnStr = "Error: CreateProcess failed. " & Err.LastDllError
    End If

'Dim nRet
'    Do
'        Sleep (100)
'        GetExitCodeProcess proc.hProcess, nRet
'        Sleep (100)
'        DoEvents
'        bSuccess = ReadFile(hReadPipe, mybuff, Len(mybuff), bytesread, 0&)

'        If bSuccess Then frmMain.Text2.Text = frmMain.Text2.Text & (left(mybuff, bytesread))
'        DoEvents
'        Sleep (100)
'    Loop While nRet = STILL_ACTIVE

'        DoEvents
'        Sleep (70)

    ' Wait for the shelled application to finish:
    ret = WaitForSingleObject(proc.hProcess, INFINITE)

If UseReturnPipe = 1 Then
        DoEvents
        Sleep (200)
        
    bSuccess = ReadFile(hReadPipe, mybuff, Len(mybuff), bytesread, 0&)
    If bSuccess = 1 Then
        sReturnStr = left(mybuff, bytesread)
    Else
        '===Error
        sReturnStr = "Error: ReadFile failed. " & Err.LastDllError
    End If
    ExecCmdPipe = sReturnStr
Else
    ExecCmdPipe = vbNullString
End If

    ret = CloseHandle(proc.hProcess)
    ret = CloseHandle(proc.hThread)
    ret = CloseHandle(hReadPipe)
    ret = CloseHandle(hWritePipe)
    

End Function

