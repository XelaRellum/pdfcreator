Attribute VB_Name = "modVersionInfo"
Option Explicit

' *** Den Artikel zu diesem Modul finden Sie unter http://www.aboutvb.de/khw/artikel/khwversioninfo.htm ***

Private Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Private mVersionInfoEx As OSVERSIONINFOEX

Public Enum viMajorConstants
    viMajorWindowsNT351 = 3
    viMajorWindows9x_NT4 = 4
    viMajorWindows2000_XP = 5
End Enum

Public Enum viMinorConstants
    viMinorWindows95_NT4_2000 = 0
    viMinorXP = 1
    viMinor2003 = 2
    viMinorWindows98 = 10
    viMinorWindowsNT351 = 51
    viMinorWindowsMe = 90
End Enum

Public Enum viPlatformIdConstants
    viPlatformWIN32s = 0
    viPlatformWIN32_9x = 1
    viPlatformWIN32_NT = 2
End Enum

Public Enum viProductTypeConstants
    viPTWorkstation = 1
    viPTDomainController = 2
    viPTServer = 3
End Enum

Public Enum viSuiteMaskConstants
    viSMSmallBusiness = &H1
    viSMEnterprise = &H2
    viSMBackOffice = &H4
    viSMCommunications = &H8
    viSMTerminal = &H10
    viSMSmallBusinessRestricted = &H20
    viSMEmbeddedNT = &H40
    viSMDataCenter = &H80
End Enum

Public Enum viWinVersionConstants
    viWinUnknown = 0
    viWin95 = 1
    viWin98 = 2
    viWinME = 3
    viWinNT351 = 4
    viWinNT4 = 5
    viWin2000 = 6
    viWinXP = 7
    viWin2003 = 8
End Enum

Private pCSDVersion As String

Public Property Get BuildNumber() As Long
    BuildNumber = mVersionInfoEx.dwBuildNumber
End Property

Public Property Get CSDVersion() As String
    CSDVersion = pCSDVersion
End Property

Public Property Get Major() As viMajorConstants
    Major = mVersionInfoEx.dwMajorVersion
End Property

Public Property Get Minor() As viMinorConstants
    Minor = mVersionInfoEx.dwMinorVersion
End Property

Public Property Get PlatformId() As Long
    PlatformId = mVersionInfoEx.dwPlatformId
End Property

Public Property Get ProductType() As Byte
    ProductType = mVersionInfoEx.wProductType
End Property

Public Property Get SPMajor() As Integer
    SPMajor = mVersionInfoEx.wServicePackMajor
End Property

Public Property Get SPMinor() As Integer
    SPMinor = mVersionInfoEx.wServicePackMinor
End Property

Public Property Get SuiteMask() As Integer
    SuiteMask = mVersionInfoEx.wSuiteMask
End Property

Public Property Get Win95() As Boolean
    With mVersionInfoEx
        Win95 = (.dwMajorVersion = viMajorWindows9x_NT4) And (.dwMinorVersion = viMinorWindows95_NT4_2000) And (.dwPlatformId = viPlatformWIN32_9x)
    End With
End Property

Public Property Get Win98() As Boolean
    With mVersionInfoEx
        Win98 = (.dwMajorVersion = viMajorWindows9x_NT4) And (.dwMinorVersion = viMinorWindows98) And (.dwPlatformId = viPlatformWIN32_9x)
    End With
End Property

Public Property Get WinME() As Boolean
    With mVersionInfoEx
        WinME = (.dwMajorVersion = viMajorWindows9x_NT4) And (.dwMinorVersion = viMinorWindowsMe) And (.dwPlatformId = viPlatformWIN32_9x)
    End With
End Property

Public Property Get WinNT351() As Boolean
    With mVersionInfoEx
        WinNT351 = (.dwMajorVersion = viMajorWindowsNT351) And (.dwMinorVersion = viMinorWindowsNT351) And (.dwPlatformId = viPlatformWIN32_NT)
    End With
End Property

Public Property Get WinNT4() As Boolean
    With mVersionInfoEx
        WinNT4 = (.dwMajorVersion = viMajorWindows9x_NT4) And (.dwMinorVersion = viMinorWindows95_NT4_2000) And (.dwPlatformId = viPlatformWIN32_NT)
    End With
End Property

Public Property Get Win2000() As Boolean
    With mVersionInfoEx
        Win2000 = (.dwMajorVersion = viMajorWindows2000_XP) And (.dwMinorVersion = viMinorWindows95_NT4_2000) And (.dwPlatformId = viPlatformWIN32_NT)
    End With
End Property

Public Property Get WinXP() As Boolean
    With mVersionInfoEx
        WinXP = (.dwMajorVersion = viMajorWindows2000_XP) And (.dwMinorVersion = viMinorXP) And (.dwPlatformId = viPlatformWIN32_NT)
    End With
End Property

Public Property Get Win2003() As Boolean
    With mVersionInfoEx
        Win2003 = (.dwMajorVersion = viMajorWindows2000_XP) And (.dwMinorVersion = viMinor2003) And (.dwPlatformId = viPlatformWIN32_NT)
    End With
End Property

Public Property Get WinVersion() As viWinVersionConstants
    Select Case True
        Case Win95
            WinVersion = viWin95
        Case Win98
            WinVersion = viWin98
        Case WinME
            WinVersion = viWinME
        Case WinNT351
            WinVersion = viWinNT351
        Case WinNT4
            WinVersion = viWinNT4
        Case Win2000
            WinVersion = viWin2000
        Case WinXP
            WinVersion = viWinXP
        Case Win2003
            WinVersion = viWin2003
    End Select
End Property

Public Property Get WinVersionText() As String
 Select Case True
  Case Win95
   WinVersionText = "Windows 95 (" & pCSDVersion & ")"
  Case Win98
   WinVersionText = "Windows 98 (" & pCSDVersion & ")"
  Case WinME
   WinVersionText = "Windows ME (" & pCSDVersion & ")"
  Case WinNT351
   WinVersionText = "Windows NT 3.51 (" & pCSDVersion & ")"
  Case WinNT4
   WinVersionText = "Windows NT 4.0 (" & pCSDVersion & ")"
  Case Win2000
   WinVersionText = "Windows 2000 (" & pCSDVersion & ")"
  Case WinXP
   WinVersionText = "Windows XP (" & pCSDVersion & ")"
  Case Win2003
   WinVersionText = "Windows 2003 (" & pCSDVersion & ")"
 End Select
End Property

Public Sub ReadVersionInfo()
 Dim nVersionInfo As OSVERSIONINFO, nPos As Long

 With nVersionInfo
  .dwOSVersionInfoSize = Len(nVersionInfo)
  GetVersionEx nVersionInfo
  nPos = InStr(.szCSDVersion, Chr$(0))
  If nPos Then
    pCSDVersion = Trim$(Left$(.szCSDVersion, nPos - 1))
   Else
    pCSDVersion = Trim$(.szCSDVersion)
  End If
  If (.dwPlatformId = viPlatformWIN32_NT) And (.dwMajorVersion = viMajorWindows2000_XP) Then
   Else
    If Len(pCSDVersion) >= 14 Then
     If (.dwPlatformId = viPlatformWIN32_NT) And (.dwMajorVersion = viMajorWindows9x_NT4) And (CLng(Mid$(pCSDVersion, 14)) > 5) Then
      Else
       LSet mVersionInfoEx = nVersionInfo
       Exit Sub
     End If
    End If
  End If
 End With
    mVersionInfoEx.dwOSVersionInfoSize = Len(mVersionInfoEx)
    GetVersionEx mVersionInfoEx
End Sub

