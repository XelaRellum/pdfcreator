Attribute VB_Name = "modWindowsVersion"
'-----------------------------------------------------------------------------------------
' Copyright ©1996-2004 VBnet, Randy Birch. All Rights Reserved Worldwide.
'        Terms of use http://vbnet.mvps.org/terms/pages/terms.htm
'-----------------------------------------------------------------------------------------

'dwPlatformId
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

'os product type values
Private Const VER_NT_WORKSTATION = &H1
Private Const VER_NT_DOMAIN_CONTROLLER = &H2
Private Const VER_NT_SERVER = &H3

'product types
Private Const VER_SERVER_NT = &H80000000
Private Const VER_WORKSTATION_NT = &H40000000

Private Const VER_SUITE_SMALLBUSINESS = &H1
Private Const VER_SUITE_ENTERPRISE = &H2
Private Const VER_SUITE_BACKOFFICE = &H4
Private Const VER_SUITE_COMMUNICATIONS = &H8
Private Const VER_SUITE_TERMINAL = &H10
Private Const VER_SUITE_SMALLBUSINESS_RESTRICTED = &H20
Private Const VER_SUITE_EMBEDDEDNT = &H40
Private Const VER_SUITE_DATACENTER = &H80
Private Const VER_SUITE_SINGLEUSERTS = &H100
Private Const VER_SUITE_PERSONAL = &H200
Private Const VER_SUITE_BLADE = &H400

Private Const OSV_LENGTH As Long = 148
Private Const OSVEX_LENGTH As Long = 156

Private Type OSVERSIONINFO
 OSVSize         As Long         'size, in bytes, of this data structure
 dwVerMajor      As Long         'ie NT 3.51, dwVerMajor = 3; NT 4.0, dwVerMajor = 4.
 dwVerMinor      As Long         'ie NT 3.51, dwVerMinor = 51; NT 4.0, dwVerMinor= 0.
 dwBuildNumber   As Long         'NT: build number of the OS
                                 'Win9x: build number of the OS in low-order word.
                                 '       High-order word contains major & minor ver nos.
 PlatformID      As Long         'Identifies the operating system platform.
 szCSDVersion    As String * 128 'NT: string, such as "Service Pack 3"
                                 'Win9x: string providing arbitrary additional information
End Type

Private Type OSVERSIONINFOEX
 OSVSize            As Long
 dwVerMajor        As Long
 dwVerMinor         As Long
 dwBuildNumber      As Long
 PlatformID         As Long
 szCSDVersion       As String * 128
 wServicePackMajor  As Integer
 wServicePackMinor  As Integer
 wSuiteMask         As Integer
 wProductType       As Byte
 wReserved          As Byte
End Type

'defined As Any to support OSVERSIONINFO and OSVERSIONINFOEX
Private Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" _
 (lpVersionInformation As Any) As Long


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
50010  'returns True if running Win95
50020  Dim osv As OSVERSIONINFO
50030  osv.OSVSize = Len(osv)
50040  If GetVersionEx(osv) = 1 Then
50050   IsWin95 = (osv.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And _
            (osv.dwVerMajor = 4 And osv.dwVerMinor = 0) And _
            (osv.dwBuildNumber = 950)
50080  End If
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
50010  'returns True if running Win95
50020  Dim osv As OSVERSIONINFO
50030  osv.OSVSize = Len(osv)
50040  If GetVersionEx(osv) = 1 Then
50050   IsWin95OSR2 = (osv.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And _
                (osv.dwVerMajor = 4 And osv.dwVerMinor = 0) And _
                (osv.dwBuildNumber = 1111)
50080  End If
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
50020  If IsWin95 = True Or IsWin98 = True Or IsWinME = True Then
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

