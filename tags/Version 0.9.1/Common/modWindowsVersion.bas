Attribute VB_Name = "modWindowsVersion"
'-----------------------------------------------------------------------------------------
' Copyright ©1996-2004 VBnet, Randy Birch. All Rights Reserved Worldwide.
'        Terms of use http://vbnet.mvps.org/terms/pages/terms.htm
'-----------------------------------------------------------------------------------------
' modified by Frank Heindörfer 2004

Public Function IsBackOfficeServer() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim osv As OSVERSIONINFOEX
50020  'Returns True if Microsoft BackOffice components are installed
50030
50040  'OSVERSIONINFOEX supported on NT4 or
50050  'later only, so a test is required
50060  'before using
50070  If IsWinNT4Plus() Then
50080   osv.OSVSize = Len(osv)
50090   If GetVersionEx(osv) = 1 Then
50100    IsBackOfficeServer = (osv.wSuiteMask And VER_SUITE_BACKOFFICE)
50110   End If
50120  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsBackOfficeServer")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsBladeServer() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim osv As OSVERSIONINFOEX
50020  'Returns True if Windows Server 2003 Web Edition is installed
50030
50040  'OSVERSIONINFOEX supported on NT4 or
50050  'later only, so a test is required
50060  'before using
50070  If IsWin2003Server() Then
50080   osv.OSVSize = Len(osv)
50090   If GetVersionEx(osv) = 1 Then
50100    IsBladeServer = (osv.wSuiteMask And VER_SUITE_BLADE)
50110   End If
50120  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsBladeServer")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsDomainController() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim osv As OSVERSIONINFOEX
50020  'Returns True if the server is a domain
50030  'controller (Win 2000 or later), including
50040  'under active directory
50050
50060  'OSVERSIONINFOEX supported on NT4 or
50070  'later only, so a test is required
50080  'before using
50090  If IsWin2000Server() Then
50100   osv.OSVSize = Len(osv)
50110   If GetVersionEx(osv) = 1 Then
50120    IsDomainController = (osv.wProductType = VER_NT_SERVER) And _
                        (osv.wProductType = VER_NT_DOMAIN_CONTROLLER)
50140   End If
50150  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsDomainController")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsEnterpriseServer() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim osv As OSVERSIONINFOEX
50020  'Returns True if Windows NT 4.0 Enterprise Edition,
50030  'Windows 2000 Advanced Server, or Windows Server 2003
50040  'Enterprise Edition is installed.
50050
50060  'OSVERSIONINFOEX supported on NT4 or
50070  'later only, so a test is required
50080  'before using
50090  If IsWinNT4Plus() Then
50100   osv.OSVSize = Len(osv)
50110   If GetVersionEx(osv) = 1 Then
50120    IsEnterpriseServer = (osv.wProductType = VER_NT_SERVER) And _
                        (osv.wSuiteMask And VER_SUITE_ENTERPRISE)
50140   End If
50150  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsEnterpriseServer")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsWin2000AdvancedServer() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim osv As OSVERSIONINFOEX
50020  'Returns True if Windows 2000 Advanced Server
50030
50040  'OSVERSIONINFOEX supported on NT4 or
50050  'later only, so a test is required
50060  'before using
50070  If IsWin2000Plus() Then
50080   osv.OSVSize = Len(osv)
50090   If GetVersionEx(osv) = 1 Then
50100    IsWin2000AdvancedServer = ((osv.wProductType = VER_NT_SERVER) Or _
                              (osv.wProductType = VER_NT_DOMAIN_CONTROLLER)) And _
                              (osv.wSuiteMask And VER_SUITE_ENTERPRISE)
50130   End If
50140  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsWin2000AdvancedServer")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsWin2000Server() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim osv As OSVERSIONINFOEX
50020  'Returns True if Windows 2000 Server
50030
50040  'OSVERSIONINFOEX supported on NT4 or
50050  'later only, so a test is required
50060  'before using
50070  If IsWin2000() Then
50080   osv.OSVSize = Len(osv)
50090   If GetVersionEx(osv) = 1 Then
50100    IsWin2000Server = (osv.wProductType = VER_NT_SERVER)
50110   End If
50120  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsWin2000Server")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsSmallBusinessServer() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim osv As OSVERSIONINFOEX
50020  'Returns True if Microsoft Small Business Server is installed
50030
50040  'OSVERSIONINFOEX supported on NT4 or
50050  'later only, so a test is required
50060  'before using
50070  If IsWinNT4Plus() Then
50080    osv.OSVSize = Len(osv)
50090    If GetVersionEx(osv) = 1 Then
50100     IsSmallBusinessServer = (osv.wSuiteMask And VER_SUITE_SMALLBUSINESS)
50110    End If
50120  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsSmallBusinessServer")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsSmallBusinessRestrictedServer() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim osv As OSVERSIONINFOEX
50020  'Returns True if Microsoft Small Business Server
50030  'is installed with the restrictive client license
50040  'in force
50050
50060  'OSVERSIONINFOEX supported on NT4 or
50070  'later only, so a test is required
50080  'before using
50090  If IsWinNT4Plus() Then
50100   osv.OSVSize = Len(osv)
50110   If GetVersionEx(osv) = 1 Then
50120    IsSmallBusinessRestrictedServer = (osv.wSuiteMask And VER_SUITE_SMALLBUSINESS_RESTRICTED)
50130   End If
50140  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsSmallBusinessRestrictedServer")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsTerminalServer() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim osv As OSVERSIONINFOEX
50020  'Returns True if Terminal Services is installed
50030  'OSVERSIONINFOEX supported on NT4 or
50040  'later only, so a test is required
50050  'before using
50060  If IsWinNT4Plus() Then
50070   osv.OSVSize = Len(osv)
50080   If GetVersionEx(osv) = 1 Then
50090    IsTerminalServer = (osv.wSuiteMask And VER_SUITE_TERMINAL) Or (osv.wSuiteMask And VER_SUITE_SINGLEUSERTS)
50100   End If
50110  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsTerminalServer")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsWin95() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim osv As OSVERSIONINFO, BuildNumber As Long
50020  osv.OSVSize = Len(osv)
50030  If GetVersionEx(osv) = 1 Then
50040   If (osv.dwBuildNumber And &HFFFF&) > &H7FFF Then
50050     BuildNumber = (osv.dwBuildNumber And &HFFFF&) - &H10000
50060    Else
50070     BuildNumber = osv.dwBuildNumber And &HFFFF&
50080   End If
50090   IsWin95 = (osv.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And _
            (osv.dwVerMajor = 4 And osv.dwVerMinor = 0) And _
            (BuildNumber = 950)
50120  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsWin95")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsWin95OSR2() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim osv As OSVERSIONINFO, BuildNumber As Long
50020  osv.OSVSize = Len(osv)
50030  If GetVersionEx(osv) = 1 Then
50040   If (osv.dwBuildNumber And &HFFFF&) > &H7FFF Then
50050     BuildNumber = (osv.dwBuildNumber And &HFFFF&) - &H10000
50060    Else
50070     BuildNumber = osv.dwBuildNumber And &HFFFF&
50080   End If
50090   IsWin95OSR2 = (osv.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And _
                (osv.dwVerMajor = 4 And osv.dwVerMinor = 0) And _
                (BuildNumber = 1111)
50120  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsWin95OSR2")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsWin98() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  'returns True if running Win98
50020  Dim osv As OSVERSIONINFO
50030  osv.OSVSize = Len(osv)
50040  If GetVersionEx(osv) = 1 Then
50050   IsWin98 = (osv.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And _
            (osv.dwVerMajor = 4 And osv.dwVerMinor = 10) And _
            (osv.dwBuildNumber >= 2222)
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsWin98")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsWinME() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  'returns True if running Windows ME
50020  Dim osv As OSVERSIONINFO
50030  osv.OSVSize = Len(osv)
50040  If GetVersionEx(osv) = 1 Then
50050   IsWinME = (osv.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And _
            (osv.dwVerMajor = 4 And osv.dwVerMinor = 90) And _
            (osv.dwBuildNumber >= 3000)
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsWinME")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsWinNT4() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  'returns True if running WinNT4
50020  Dim osv As OSVERSIONINFO
50030  osv.OSVSize = Len(osv)
50040  If GetVersionEx(osv) = 1 Then
50050   IsWinNT4 = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
             (osv.dwVerMajor = 4 And osv.dwVerMinor = 0) And _
             (osv.dwBuildNumber >= 1381)
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsWinNT4")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsWinNT4Plus() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  'returns True if running Windows NT4 or later
50020  Dim osv As OSVERSIONINFO
50030  osv.OSVSize = Len(osv)
50040  If GetVersionEx(osv) = 1 Then
50050   IsWinNT4Plus = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                 (osv.dwVerMajor >= 4)
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsWinNT4Plus")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsWinNT4Server() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  'returns True if running Windows NT4 Server
50020  Dim osv As OSVERSIONINFOEX
50030  If IsWinNT4() Then
50040   osv.OSVSize = Len(osv)
50050   If GetVersionEx(osv) = 1 Then
50060    IsWinNT4Server = (osv.wProductType And VER_NT_SERVER)
50070   End If
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsWinNT4Server")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsWinNT4Workstation() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  'returns True if running Windows NT4 Workstation
50020  Dim osv As OSVERSIONINFOEX
50030  If IsWinNT4() Then
50040   osv.OSVSize = Len(osv)
50050   If GetVersionEx(osv) = 1 Then
50060    IsWinNT4Workstation = (osv.wProductType And VER_NT_WORKSTATION)
50070   End If
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsWinNT4Workstation")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsWin2000() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  'returns True if running Win2000 (NT5)
50020  Dim osv As OSVERSIONINFO
50030  osv.OSVSize = Len(osv)
50040  If GetVersionEx(osv) = 1 Then
50050   IsWin2000 = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
              (osv.dwVerMajor = 5 And osv.dwVerMinor = 0) And _
              (osv.dwBuildNumber >= 2195)
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsWin2000")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsWin2000Plus() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  'returns True if running Windows 2000 or later
50020  Dim osv As OSVERSIONINFO
50030  osv.OSVSize = Len(osv)
50040  If GetVersionEx(osv) = 1 Then
50050   IsWin2000Plus = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                  (osv.dwVerMajor = 5 And osv.dwVerMinor >= 0)
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsWin2000Plus")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsWin2003Server() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  'returns True if running Windows 2003 (.NET) Server
50020  Dim osv As OSVERSIONINFO
50030  osv.OSVSize = Len(osv)
50040  If GetVersionEx(osv) = 1 Then
50050   IsWin2003Server = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                    (osv.dwVerMajor = 5 And osv.dwVerMinor = 2) And _
                    (osv.dwBuildNumber = 3790)
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsWin2003Server")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsWin2000Workstation() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  'returns True if running Windows NT4 Workstation
50020  Dim osv As OSVERSIONINFOEX
50030  If IsWin2000() Then
50040   osv.OSVSize = Len(osv)
50050   If GetVersionEx(osv) = 1 Then
50060    IsWin2000Workstation = (osv.wProductType And VER_NT_WORKSTATION)
50070   End If
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsWin2000Workstation")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsWinXP() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  'returns True if running Windows XP
50020  Dim osv As OSVERSIONINFO
50030  osv.OSVSize = Len(osv)
50040  If GetVersionEx(osv) = 1 Then
50050   IsWinXP = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
            (osv.dwVerMajor = 5 And osv.dwVerMinor = 1) And _
            (osv.dwBuildNumber >= 2600)
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsWinXP")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsWinXPSP2() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  'returns True if running Windows XP SP2 (Service Pack 2)
50020  Dim osv As OSVERSIONINFOEX
50030  If IsWinXP() Then
50040   osv.OSVSize = Len(osv)
50050   If GetVersionEx(osv) = 1 Then
50060    IsWinXPSP2 = InStr(osv.szCSDVersion, "Service Pack 2") > 0
50070   End If
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsWinXPSP2")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsWinXPPlus() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  'returns True if running Windows XP or later
50020  Dim osv As OSVERSIONINFO
50030  osv.OSVSize = Len(osv)
50040  If GetVersionEx(osv) = 1 Then
50050   IsWinXPPlus = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                (osv.dwVerMajor >= 5 And osv.dwVerMinor >= 1)
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsWinXPPlus")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsWinXPHomeEdition() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  'returns True if running Windows XP Home Edition
50020  Dim osv As OSVERSIONINFOEX
50030  If IsWinXP() Then
50040   osv.OSVSize = Len(osv)
50050   If GetVersionEx(osv) = 1 Then
50060    IsWinXPHomeEdition = ((osv.wSuiteMask And VER_SUITE_PERSONAL) = VER_SUITE_PERSONAL)
50070   End If
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsWinXPHomeEdition")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsWinXPProEdition() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  'returns True if running WinXP Pro
50020  Dim osv As OSVERSIONINFOEX
50030  If IsWinXP() Then
50040   osv.OSVSize = Len(osv)
50050   If GetVersionEx(osv) = 1 Then
50060    IsWinXPProEdition = Not ((osv.wSuiteMask And VER_SUITE_PERSONAL) = VER_SUITE_PERSONAL)
50070   End If
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsWinXPProEdition")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsWin9xMe() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  'returns True if running Win95, Win98 or WinMe
50020  If IsWin95 = True Or IsWin95OSR2 = True Or IsWin98 = True Or IsWinME = True Then
50030   IsWin9xMe = True
50040  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "IsWin9xMe")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetWinVersionStr() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStr As String, win As RGB_WINVER
50020
50030  If IsBackOfficeServer Then
50040   tStr = tStr & " BackOfficeServer"
50050  End If
50060  If IsBladeServer Then
50070   tStr = tStr & " BladeServer"
50080  End If
50090  If IsDomainController Then
50100   tStr = tStr & " DomainController"
50110  End If
50120  If IsEnterpriseServer Then
50130   tStr = tStr & " EnterpriseServer"
50140  End If
50150  If IsSmallBusinessRestrictedServer Then
50160   tStr = tStr & " SmallBusinessRestrictedServer"
50170  End If
50180  If IsSmallBusinessServer Then
50190   tStr = tStr & " SmallBusinessServer"
50200  End If
50210  If IsTerminalServer Then
50220   tStr = tStr & " TerminalServer"
50230  End If
50240  If IsWin2000 Then
50250   tStr = tStr & " Win2000"
50260  End If
50270  If IsWin2000AdvancedServer Then
50280   tStr = tStr & " Win2000AdvancedServer"
50290  End If
50300  If IsWin2000Server Then
50310   tStr = tStr & " Win2000Server"
50320  End If
50330  If IsWin2000Workstation Then
50340   tStr = tStr & " Win2000Workstation"
50350  End If
50360  If IsWin2003Server Then
50370   tStr = tStr & " Win2003Server"
50380  End If
50390  If IsWin95 Then
50400   tStr = tStr & " Win95"
50410  End If
50420  If IsWin95OSR2 Then
50430   tStr = tStr & " Win95OSR2"
50440  End If
50450  If IsWin98 Then
50460   tStr = tStr & " Win98"
50470  End If
50480  If IsWinME Then
50490   tStr = tStr & " WinME"
50500  End If
50510  If IsWinNT4 Then
50520   tStr = tStr & " WinNT4"
50530  End If
50540  If IsWinNT4Server Then
50550   tStr = tStr & " WinNT4Server"
50560  End If
50570  If IsWinNT4Workstation Then
50580   tStr = tStr & " WinNT4Workstation"
50590  End If
50600  If IsWinXP Then
50610   tStr = tStr & " WinXP"
50620  End If
50630  If IsWinXPHomeEdition Then
50640   tStr = tStr & " WinXPHomeEdition"
50650  End If
50660  If IsWinXPProEdition Then
50670   tStr = tStr & " WinXPProEdition"
50680  End If
50690  If IsWinXPSP2 Then
50700   tStr = tStr & " WinXPSP2"
50710  End If
50720
50730  tStr = Trim$(tStr)
50740  If Len(tStr) > 0 Then
50750   tStr = " [" & tStr & "]"
50760  End If
50770  Call GetWinVersion(win)
50780  GetWinVersionStr = win.VersionName & " " & win.VersionNo & " Build " & _
  win.BuildNo & " (" & win.ServicePack & ")" & tStr
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "GetWinVersionStr")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function GetWinVersion(win As RGB_WINVER) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim osv As OSVERSIONINFO, pos As Integer, sVer As String, sBuild As String
50020  osv.OSVSize = Len(osv)
50030  If GetVersionEx(osv) = 1 Then
50040   win.PlatformID = osv.PlatformID
50051   Select Case osv.PlatformID
         Case VER_PLATFORM_WIN32s:   win.VersionName = "Win32s"
50070    Case VER_PLATFORM_WIN32_NT: win.VersionName = "Windows NT"
50081    Select Case osv.dwVerMajor
          Case 4:  win.VersionName = "Windows NT"
50100     Case 5:
50111      Select Case osv.dwVerMinor
            Case 0:  win.VersionName = "Windows 2000"
50130       Case 1:  win.VersionName = "Windows XP"
50140      End Select
50150    End Select
50160    Case VER_PLATFORM_WIN32_WINDOWS:
50171     Select Case osv.dwVerMinor
           Case 0:    win.VersionName = "Windows 95"
50190      Case 90:   win.VersionName = "Windows ME"
50200      Case Else: win.VersionName = "Windows 98"
50210     End Select
50220   End Select
50230   'Get the version number
50240   win.VersionNo = osv.dwVerMajor & "." & osv.dwVerMinor
50250   'Get the build
50260   win.BuildNo = (osv.dwBuildNumber And &HFFFF&)
50270   'Any additional info. In Win9x, this can be
50280   '"any arbitrary string" provided by the
50290   'manufacturer. In NT, this is the service pack.
50300   pos = InStr(osv.szCSDVersion, Chr$(0))
50310   If pos Then
50320    win.ServicePack = Left$(osv.szCSDVersion, pos - 1)
50330   End If
50340  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modWindowsVersion", "GetWinVersion")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

