Attribute VB_Name = "modOS"
Option Explicit

Private Const VER_PLATFORM_WIN32_WINDOWS As Long = &H1

Private Type OSVERSIONINFO
 dwOSVersionInfoSize As Long
 dwMajorVersion As Long
 dwMinorVersion As Long
 dwBuildNumber As Long
 dwPlatformId As Long
 szCSDVersion As String * 128
End Type

Private Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" _
 (lpVersionInformation As Any) As Long

Private Declare Function GetWindowsDirectory Lib "Kernel32" Alias _
 "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Function GetWinDir() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim nBuffer As String, res As Long
50020
50030  nBuffer = Space(255)
50040  res = GetWindowsDirectory(nBuffer, 255)
50050  If res > 0 Then
50060   GetWinDir = Left(nBuffer, res)
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modOS", "GetWinDir")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
