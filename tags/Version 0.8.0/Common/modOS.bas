Attribute VB_Name = "modOS"
Option Explicit

Private Const Ver_Platform_Win32_Windows As Long = &H1

Private Type OSVersionInfo
 dwOSVersionInfoSize As Long
 dwMajorVersion As Long
 dwMinorVersion As Long
 dwBuildNumber As Long
 dwPlatformId As Long
 szCSDVersion As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
 (lpVersionInformation As Any) As Long

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias _
 "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Function IsWin9xMe() As Boolean
 Dim os As OSVersionInfo, res As Long
 os.dwOSVersionInfoSize = Len(os)
 res = GetVersionEx(os)
 If os.dwPlatformId = Ver_Platform_Win32_Windows Then
   IsWin9xMe = True
  Else
   IsWin9xMe = False
 End If
End Function

Public Function GetWinDir() As String
 Dim nBuffer As String, res As Long

 nBuffer = Space(255)
 res = GetWindowsDirectory(nBuffer, 255)
 If res > 0 Then
  GetWinDir = Left(nBuffer, res)
 End If
End Function
