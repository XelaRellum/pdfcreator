Attribute VB_Name = "modFileInUse"
Option Explicit
 
Private Const GENERIC_READ               As Long = &H80000000
Private Const OPEN_EXISTING              As Long = 3
Private Const FILE_ATTRIBUTE_NORMAL      As Long = &H80
Private Const INVALID_HANDLE_VALUE       As Long = -1
Private Const ERROR_SHARING_VIOLATION    As Long = &H20
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
 
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, ByVal lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, nSize As Long, Arguments As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
 
Public Function FileInUse(ByVal sFilename As String, Optional ByRef sErrorMsg As String = vbNullString) As Boolean
 Dim lFile As Long, lError As Long, lResult As Long, lBufLen As Long, sBuffer As String
 On Error GoTo Err_FileInUse
 lFile = CreateFile(sFilename, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
 If lFile = INVALID_HANDLE_VALUE Then
   lError = Err.LastDllError
   If lError = ERROR_SHARING_VIOLATION Then
     FileInUse = True
    Else
     sBuffer = String(256, 0)
     lResult = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, lError, 0, sBuffer, Len(sBuffer), 0)
     If lResult > 0 Then
       sErrorMsg = Left$(sBuffer, lResult - 1)
      Else
       sErrorMsg = "Unknown error!"
     End If
    End If
   Else
    CloseHandle lFile
 End If
Exit_FileInUse:
 On Error GoTo 0
 Exit Function

Err_FileInUse:
 If lFile = INVALID_HANDLE_VALUE Then
  CloseHandle lFile
 End If
 If Err.number <> 0 Then
  sErrorMsg = Err.Description
 End If
 On Error GoTo 0
End Function

