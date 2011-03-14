Attribute VB_Name = "modProfile"
Private Const TOKEN_QUERY = (&H8)
Private Declare Function GetAllUsersProfileDirectory Lib "userenv.dll" Alias "GetAllUsersProfileDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function GetDefaultUserProfileDirectory Lib "userenv.dll" Alias "GetDefaultUserProfileDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function GetProfilesDirectory Lib "userenv.dll" Alias "GetProfilesDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function GetUserProfileDirectory Lib "userenv.dll" Alias "GetUserProfileDirectoryA" (ByVal hToken As Long, ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long

Public Function GetUserProfileStrings() As String
 Dim sBuffer As String, ret As Long, hToken As Long, tStr As String
 sBuffer = String(255, 0)
 GetAllUsersProfileDirectory sBuffer, 255
 tStr = "GetAllUsersProfileDirectory: " & StripTerminator(sBuffer)
 sBuffer = String(255, 0)
 GetDefaultUserProfileDirectory sBuffer, 255
 tStr = tStr & vbCrLf & "GetDefaultUserProfileDirectory: " & StripTerminator(sBuffer)
 sBuffer = String(255, 0)
 GetProfilesDirectory sBuffer, 255
 tStr = tStr & vbCrLf & "GetProfilesDirectory: " & StripTerminator(sBuffer)
 sBuffer = String(255, 0)
 OpenProcessToken GetCurrentProcess, TOKEN_QUERY, hToken
 GetUserProfileDirectory hToken, sBuffer, 255
 tStr = tStr & vbCrLf & "GetUserProfileDirectory: " & StripTerminator(sBuffer)
 GetUserProfileStrings = tStr
End Function

Public Function GetTheUserProfileStrings() As String
 Dim sBuffer As String, ret As Long, hToken As Long, tStr As String
 sBuffer = String(255, 0)
 OpenProcessToken GetCurrentProcess, TOKEN_QUERY, hToken
 GetUserProfileDirectory hToken, sBuffer, 255
 GetTheUserProfileStrings = StripTerminator(sBuffer)
End Function

Private Function StripTerminator(sInput As String) As String
 Dim ZeroPos As Long
 ZeroPos = InStr(1, sInput, Chr$(0))
 If ZeroPos > 0 Then
   StripTerminator = Left$(sInput, ZeroPos - 1)
  Else
   StripTerminator = sInput
 End If
End Function
