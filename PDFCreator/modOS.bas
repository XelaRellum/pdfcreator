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
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim nBuffer As String, res As Long
50050
50060  nBuffer = Space(255)
50070  res = GetWindowsDirectory(nBuffer, 255)
50080  If res > 0 Then
50090   GetWinDir = Left(nBuffer, res)
50100  End If
50110 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50120 Exit Function
ErrPtnr_OnError:
50141 Select Case ErrPtnr.OnError("modOS", "GetWinDir")
      Case 0: Resume
50160 Case 1: Resume Next
50170 Case 2: Exit Function
50180 Case 3: End
50190 End Select
50200 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
