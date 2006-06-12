Attribute VB_Name = "modWindowsVersion"
'-----------------------------------------------------------------------------------------
' Copyright ©1996-2004 VBnet, Randy Birch. All Rights Reserved Worldwide.
'        Terms of use http://vbnet.mvps.org/terms/pages/terms.htm
'-----------------------------------------------------------------------------------------
' modified by Frank Heindörfer 2004

Public Function IsBackOfficeServer() As Boolean
 Dim osv As OSVERSIONINFOEX
 'Returns True if Microsoft BackOffice components are installed

 'OSVERSIONINFOEX supported on NT4 or
 'later only, so a test is required
 'before using
 If IsWinNT4Plus() Then
  osv.OSVSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
   IsBackOfficeServer = (osv.wSuiteMask And VER_SUITE_BACKOFFICE)
  End If
 End If
End Function

Public Function IsBladeServer() As Boolean
 Dim osv As OSVERSIONINFOEX
 'Returns True if Windows Server 2003 Web Edition is installed

 'OSVERSIONINFOEX supported on NT4 or
 'later only, so a test is required
 'before using
 If IsWin2003Server() Then
  osv.OSVSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
   IsBladeServer = (osv.wSuiteMask And VER_SUITE_BLADE)
  End If
 End If
End Function

Public Function IsDomainController() As Boolean
 Dim osv As OSVERSIONINFOEX
 'Returns True if the server is a domain
 'controller (Win 2000 or later), including
 'under active directory

 'OSVERSIONINFOEX supported on NT4 or
 'later only, so a test is required
 'before using
 If IsWin2000Server() Then
  osv.OSVSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
   IsDomainController = (osv.wProductType = VER_NT_SERVER) And _
                        (osv.wProductType = VER_NT_DOMAIN_CONTROLLER)
  End If
 End If
End Function

Public Function IsEnterpriseServer() As Boolean
 Dim osv As OSVERSIONINFOEX
 'Returns True if Windows NT 4.0 Enterprise Edition,
 'Windows 2000 Advanced Server, or Windows Server 2003
 'Enterprise Edition is installed.

 'OSVERSIONINFOEX supported on NT4 or
 'later only, so a test is required
 'before using
 If IsWinNT4Plus() Then
  osv.OSVSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
   IsEnterpriseServer = (osv.wProductType = VER_NT_SERVER) And _
                        (osv.wSuiteMask And VER_SUITE_ENTERPRISE)
  End If
 End If
End Function

Public Function IsWin2000AdvancedServer() As Boolean
 Dim osv As OSVERSIONINFOEX
 'Returns True if Windows 2000 Advanced Server

 'OSVERSIONINFOEX supported on NT4 or
 'later only, so a test is required
 'before using
 If IsWin2000Plus() Then
  osv.OSVSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
   IsWin2000AdvancedServer = ((osv.wProductType = VER_NT_SERVER) Or _
                              (osv.wProductType = VER_NT_DOMAIN_CONTROLLER)) And _
                              (osv.wSuiteMask And VER_SUITE_ENTERPRISE)
  End If
 End If
End Function

Public Function IsWin2000Server() As Boolean
 Dim osv As OSVERSIONINFOEX
 'Returns True if Windows 2000 Server

 'OSVERSIONINFOEX supported on NT4 or
 'later only, so a test is required
 'before using
 If IsWin2000() Then
  osv.OSVSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
   IsWin2000Server = (osv.wProductType = VER_NT_SERVER)
  End If
 End If
End Function

Public Function IsSmallBusinessServer() As Boolean
 Dim osv As OSVERSIONINFOEX
 'Returns True if Microsoft Small Business Server is installed

 'OSVERSIONINFOEX supported on NT4 or
 'later only, so a test is required
 'before using
 If IsWinNT4Plus() Then
   osv.OSVSize = Len(osv)
   If GetVersionEx(osv) = 1 Then
    IsSmallBusinessServer = (osv.wSuiteMask And VER_SUITE_SMALLBUSINESS)
   End If
 End If
End Function

Public Function IsSmallBusinessRestrictedServer() As Boolean
 Dim osv As OSVERSIONINFOEX
 'Returns True if Microsoft Small Business Server
 'is installed with the restrictive client license
 'in force

 'OSVERSIONINFOEX supported on NT4 or
 'later only, so a test is required
 'before using
 If IsWinNT4Plus() Then
  osv.OSVSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
   IsSmallBusinessRestrictedServer = (osv.wSuiteMask And VER_SUITE_SMALLBUSINESS_RESTRICTED)
  End If
 End If
End Function

Public Function IsTerminalServer() As Boolean
 Dim osv As OSVERSIONINFOEX
 'Returns True if Terminal Services is installed
 'OSVERSIONINFOEX supported on NT4 or
 'later only, so a test is required
 'before using
 If IsWinNT4Plus() Then
  osv.OSVSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
   IsTerminalServer = (osv.wSuiteMask And VER_SUITE_TERMINAL) Or (osv.wSuiteMask And VER_SUITE_SINGLEUSERTS)
  End If
 End If
End Function

Public Function IsWin95() As Boolean
 Dim osv As OSVERSIONINFO, BuildNumber As Long
 osv.OSVSize = Len(osv)
 If GetVersionEx(osv) = 1 Then
  If (osv.dwBuildNumber And &HFFFF&) > &H7FFF Then
    BuildNumber = (osv.dwBuildNumber And &HFFFF&) - &H10000
   Else
    BuildNumber = osv.dwBuildNumber And &HFFFF&
  End If
  IsWin95 = (osv.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And _
            (osv.dwVerMajor = 4 And ((osv.dwVerMinor = 0) Or (osv.dwVerMinor = 3))) 'And (BuildNumber = 950)
 End If
End Function

Public Function IsWin98() As Boolean
 'returns True if running Win98
 Dim osv As OSVERSIONINFO
 osv.OSVSize = Len(osv)
 If GetVersionEx(osv) = 1 Then
  IsWin98 = (osv.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And _
            (osv.dwVerMajor = 4 And osv.dwVerMinor = 10) And _
            (osv.dwBuildNumber >= 2222)
 End If
End Function

Public Function IsWinME() As Boolean
 'returns True if running Windows ME
 Dim osv As OSVERSIONINFO
 osv.OSVSize = Len(osv)
 If GetVersionEx(osv) = 1 Then
  IsWinME = (osv.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And _
            (osv.dwVerMajor = 4 And osv.dwVerMinor = 90) And _
            (osv.dwBuildNumber >= 3000)
 End If
End Function

Public Function IsWinNT4() As Boolean
 'returns True if running WinNT4
 Dim osv As OSVERSIONINFO
 osv.OSVSize = Len(osv)
 If GetVersionEx(osv) = 1 Then
  IsWinNT4 = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
             (osv.dwVerMajor = 4 And osv.dwVerMinor = 0) And _
             (osv.dwBuildNumber >= 1381)
 End If
End Function

Public Function IsWinNT4Plus() As Boolean
 'returns True if running Windows NT4 or later
 Dim osv As OSVERSIONINFO
 osv.OSVSize = Len(osv)
 If GetVersionEx(osv) = 1 Then
  IsWinNT4Plus = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                 (osv.dwVerMajor >= 4)
 End If
End Function

Public Function IsWinNT4Server() As Boolean
 'returns True if running Windows NT4 Server
 Dim osv As OSVERSIONINFOEX
 If IsWinNT4() Then
  osv.OSVSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
   IsWinNT4Server = (osv.wProductType And VER_NT_SERVER)
  End If
 End If
End Function

Public Function IsWinNT4Workstation() As Boolean
 'returns True if running Windows NT4 Workstation
 Dim osv As OSVERSIONINFOEX
 If IsWinNT4() Then
  osv.OSVSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
   IsWinNT4Workstation = (osv.wProductType And VER_NT_WORKSTATION)
  End If
 End If
End Function

Public Function IsWin2000() As Boolean
 'returns True if running Win2000 (NT5)
 Dim osv As OSVERSIONINFO
 osv.OSVSize = Len(osv)
 If GetVersionEx(osv) = 1 Then
  IsWin2000 = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
              (osv.dwVerMajor = 5 And osv.dwVerMinor = 0) And _
              (osv.dwBuildNumber >= 2195)
 End If
End Function

Public Function IsWin2000Plus() As Boolean
 'returns True if running Windows 2000 or later
 Dim osv As OSVERSIONINFO
 osv.OSVSize = Len(osv)
 If GetVersionEx(osv) = 1 Then
  IsWin2000Plus = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                  (osv.dwVerMajor = 5 And osv.dwVerMinor >= 0)
 End If
End Function

Public Function IsWin2003Server() As Boolean
 'returns True if running Windows 2003 (.NET) Server
 Dim osv As OSVERSIONINFO
 osv.OSVSize = Len(osv)
 If GetVersionEx(osv) = 1 Then
  IsWin2003Server = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                    (osv.dwVerMajor = 5 And osv.dwVerMinor = 2) And _
                    (osv.dwBuildNumber = 3790)
 End If
End Function

Public Function IsWin2000Workstation() As Boolean
 'returns True if running Windows NT4 Workstation
 Dim osv As OSVERSIONINFOEX
 If IsWin2000() Then
  osv.OSVSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
   IsWin2000Workstation = (osv.wProductType And VER_NT_WORKSTATION)
  End If
 End If
End Function

Public Function IsWinXP() As Boolean
 'returns True if running Windows XP
 Dim osv As OSVERSIONINFO
 osv.OSVSize = Len(osv)
 If GetVersionEx(osv) = 1 Then
  IsWinXP = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
            (osv.dwVerMajor = 5 And osv.dwVerMinor = 1) And _
            (osv.dwBuildNumber >= 2600)
 End If
End Function

Public Function IsWinXPSP2() As Boolean
 'returns True if running Windows XP SP2 (Service Pack 2)
 Dim osv As OSVERSIONINFOEX
 If IsWinXP() Then
  osv.OSVSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
   IsWinXPSP2 = InStr(osv.szCSDVersion, "Service Pack 2") > 0
  End If
 End If
End Function

Public Function IsWinVista() As Boolean
 'returns True if running Windows Vista
 Dim osv As OSVERSIONINFO
 osv.OSVSize = Len(osv)
 If GetVersionEx(osv) = 1 Then
  IsWinVista = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
            (osv.dwVerMajor = 6 And osv.dwVerMinor = 0)
 End If
End Function

Public Function IsWinXPPlus() As Boolean
 'returns True if running Windows XP or later
 Dim osv As OSVERSIONINFO
 osv.OSVSize = Len(osv)
 If GetVersionEx(osv) = 1 Then
  IsWinXPPlus = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                (osv.dwVerMajor >= 5 And osv.dwVerMinor >= 1)
 End If
End Function

Public Function IsWinXPHomeEdition() As Boolean
 'returns True if running Windows XP Home Edition
 Dim osv As OSVERSIONINFOEX
 If IsWinXP() Then
  osv.OSVSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
   IsWinXPHomeEdition = ((osv.wSuiteMask And VER_SUITE_PERSONAL) = VER_SUITE_PERSONAL)
  End If
 End If
End Function

Public Function IsWinXPProEdition() As Boolean
 'returns True if running WinXP Pro
 Dim osv As OSVERSIONINFOEX
 If IsWinXP() Then
  osv.OSVSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
   IsWinXPProEdition = Not ((osv.wSuiteMask And VER_SUITE_PERSONAL) = VER_SUITE_PERSONAL)
  End If
 End If
End Function

Public Function IsWin9xMe() As Boolean
 'returns True if running Win95, Win98 or WinMe
 If IsWin95 = True Or IsWin98 = True Or IsWinME = True Then
  IsWin9xMe = True
 End If
End Function

Public Function GetWinVersionStr() As String
 Dim tStr As String, win As RGB_WINVER

 If IsBackOfficeServer Then
  tStr = tStr & " BackOfficeServer"
 End If
 If IsBladeServer Then
  tStr = tStr & " BladeServer"
 End If
 If IsDomainController Then
  tStr = tStr & " DomainController"
 End If
 If IsEnterpriseServer Then
  tStr = tStr & " EnterpriseServer"
 End If
 If IsSmallBusinessRestrictedServer Then
  tStr = tStr & " SmallBusinessRestrictedServer"
 End If
 If IsSmallBusinessServer Then
  tStr = tStr & " SmallBusinessServer"
 End If
 If IsTerminalServer Then
  tStr = tStr & " TerminalServer"
 End If
 If IsWin2000 Then
  tStr = tStr & " Win2000"
 End If
 If IsWin2000AdvancedServer Then
  tStr = tStr & " Win2000AdvancedServer"
 End If
 If IsWin2000Server Then
  tStr = tStr & " Win2000Server"
 End If
 If IsWin2000Workstation Then
  tStr = tStr & " Win2000Workstation"
 End If
 If IsWin2003Server Then
  tStr = tStr & " Win2003Server"
 End If
 If IsWin95 Then
  tStr = tStr & " Win95"
 End If
 If IsWin95OSR2 Then
  tStr = tStr & " Win95OSR2"
 End If
 If IsWin98 Then
  tStr = tStr & " Win98"
 End If
 If IsWinME Then
  tStr = tStr & " WinME"
 End If
 If IsWinNT4 Then
  tStr = tStr & " WinNT4"
 End If
 If IsWinNT4Server Then
  tStr = tStr & " WinNT4Server"
 End If
 If IsWinNT4Workstation Then
  tStr = tStr & " WinNT4Workstation"
 End If
 If IsWinXP Then
  tStr = tStr & " WinXP"
 End If
 If IsWinXPHomeEdition Then
  tStr = tStr & " WinXPHomeEdition"
 End If
 If IsWinXPProEdition Then
  tStr = tStr & " WinXPProEdition"
 End If
 If IsWinXPSP2 Then
  tStr = tStr & " WinXPSP2"
 End If

 tStr = Trim$(tStr)
 If Len(tStr) > 0 Then
  tStr = " [" & tStr & "]"
 End If
 Call GetWinVersion(win)
 GetWinVersionStr = win.VersionName & " " & win.VersionNo & " Build " & _
  win.BuildNo & " (" & win.ServicePack & ")" & tStr
End Function

Private Function GetWinVersion(win As RGB_WINVER) As String
 Dim osv As OSVERSIONINFOEX, pos As Integer, sVer As String, sBuild As String
 osv.OSVSize = Len(osv)
 If GetVersionEx(osv) = 1 Then
  win.PlatformID = osv.PlatformID
  Select Case osv.PlatformID
   Case VER_PLATFORM_WIN32s:   win.VersionName = "Win32s"
   Case VER_PLATFORM_WIN32_NT: win.VersionName = "Windows NT"
   Select Case osv.dwVerMajor
    Case 4:  win.VersionName = "Windows NT"
    Case 5:
     Select Case osv.dwVerMinor
      Case 0:  win.VersionName = "Windows 2000"
      Case 1:  win.VersionName = "Windows XP"
      Case 2:  win.VersionName = "Windows 2003"
      Case Else:  win.VersionName = "VerMajor: 5 -> Unknown 'VerMinor':" & osv.dwVerMinor
     End Select
    Case 6:
     Select Case osv.dwVerMinor
      Case 0:
       If osv.wProductType = VER_NT_WORKSTATION Then
         win.VersionName = "Windows Vista"
        Else
         win.VersionName = "Windows Server Longhorn"
       End If
      Case Else:  win.VersionName = "VerMajor: 6 -> Unknown 'VerMinor':" & osv.dwVerMinor
     End Select
    Case Else: win.VersionName = "Unknown 'VerMajor':" & osv.dwVerMajor
   End Select
   Case VER_PLATFORM_WIN32_WINDOWS:
    Select Case osv.dwVerMinor
     Case 0:    win.VersionName = "Windows 95"
     Case 90:   win.VersionName = "Windows ME"
     Case Else: win.VersionName = "Windows 98"
    End Select
  End Select
  'Get the version number
  win.VersionNo = osv.dwVerMajor & "." & osv.dwVerMinor
  'Get the build
  win.BuildNo = (osv.dwBuildNumber And &HFFFF&)
  'Any additional info. In Win9x, this can be
  '"any arbitrary string" provided by the
  'manufacturer. In NT, this is the service pack.
  pos = InStr(osv.szCSDVersion, Chr$(0))
  If pos Then
   win.ServicePack = Left$(osv.szCSDVersion, pos - 1)
  End If
 End If
End Function

