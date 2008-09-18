Attribute VB_Name = "modFileInUse"
Option Explicit
 
Public Function FileInUse(ByVal sFilename As String, Optional ByRef sErrorMsg As String = vbNullString) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim lFile As Long, lError As Long, lResult As Long, lBufLen As Long, sBuffer As String
50020  On Error GoTo Err_FileInUse
50030  lFile = CreateFile(sFilename, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
50040  If lFile = INVALID_HANDLE_VALUE Then
50050    lError = Err.LastDllError
50060    If lError = ERROR_SHARING_VIOLATION Then
50070      FileInUse = True
50080     Else
50090      sBuffer = String(256, 0)
50100      lResult = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, lError, 0, sBuffer, Len(sBuffer), 0)
50110      If lResult > 0 Then
50120        sErrorMsg = Left$(sBuffer, lResult - 1)
50130       Else
50140        sErrorMsg = "Unknown error!"
50150      End If
50160     End If
50170    Else
50180     CloseHandle lFile
50190  End If
Exit_FileInUse:
50210  On Error GoTo ErrPtnr_OnError
50220  Exit Function
50230
Err_FileInUse:
50250  If lFile = INVALID_HANDLE_VALUE Then
50260   CloseHandle lFile
50270  End If
50280  If Err.Number <> 0 Then
50290   sErrorMsg = Err.Description
50300  End If
50310  On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modFileInUse", "FileInUse")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

