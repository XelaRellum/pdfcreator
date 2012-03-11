Attribute VB_Name = "modFileInUse"
Option Explicit

Public Function FileInUse(ByVal sFilename As String, Optional ByRef sErrorMsg As String = vbNullString) As Boolean
 Dim lFile As Long, lError As Long, lResult As Long, sBuffer As String
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
 If Err.Number <> 0 Then
  sErrorMsg = Err.Description
 End If
 On Error GoTo 0
End Function

