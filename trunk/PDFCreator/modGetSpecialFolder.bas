Attribute VB_Name = "modGetSpecialFolder"
'// ---------------------------------------------------------------------------
'// Modul:    modGetSpecialFolder (21.06.2001)
'//           Pfade von besonderen Systemordnern ermitteln
'//
'// Copyright ©2001 Thorsten Dörfler (doerfler.t@vb-hellfire.de)
'//           http://www.vb-hellfire.de
'// ---------------------------------------------------------------------------
Option Explicit

Private Type SHITEMID
 cb   As Long
 abID As Byte
End Type

Private Type ITEMIDLIST
 mkid As SHITEMID
End Type

Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByRef ppidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Const S_OK = 0
Private Const MAX_PATH = 260

Private Const CSIDL_FLAG_CREATE = &H8000&

Public Enum ShellSpecialFolderConstants
  ssfDESKTOP = &H0                   '// <Desktop>
  ssfPROGRAMS = &H2                  '// Startmenü\Programme
  ssfPERSONAL = &H5                  '// Eigene Dateien
  ssfFAVORITES = &H6                 '// <Benutzer>\Favoriten
  ssfSTARTUP = &H7                   '// Startmenü\Programme\Autostart
  ssfRECENT = &H8                    '// <Benutzer>\Recent
  ssfSENDTO = &H9                    '// <Benutzer>\SendTo
  ssfSTARTMENU = &HB                 '// <Benutzer>\Startmenü
  ssfDESKTOPDIRECTORY = &H10         '// <Benutzer>\Desktop
  ssfNETHOOD = &H13                  '// <Benutzer>\Netzwerkumgebung
  ssfFONTS = &H14                    '// Windows\Fonts
  ssfTEMPLATES = &H15                '// <Benutzer>\Vorlagen
  ssfCOMMONSTARTMENU = &H16          '// All Users\Startmenü
  ssfCOMMONPROGRAMS = &H17           '// All Users\Startmenü\Programme
  ssfCOMMONSTARTUP = &H18            '// All Users\Startmenü\Autostart
  ssfCOMMONDESKTOPDIRECTORY = &H19   '// All Users\Desktop
  ssfAPPDATA = &H1A                  '// <Benutzer>\Anwendungsdaten
  ssfPRINTHOOD = &H1B                '// <Benutzer>\Druckumgebung
  ssfLOCALAPPDATA = &H1C             '// <Benutzer>\Lokale Einstell.\Anwendungsdaten
  ssfALTSTARTUP = &H1D               '// Startup
  ssfCOMMONALTSTARTUP = &H1E         '// Common Startup
  ssfCOMMONFAVORITES = &H1F          '// All Users\Favoriten
  ssfINTERNET_CACHE = &H20           '// <Benutzer>\Lokale Einstell.\Temp. Internet Files
  ssfCOOKIES = &H21                  '// <Benutzer>\Cookies
  ssfHISTORY = &H22                  '// <Benutzer>\Lokale Einstell.\Verlauf
  ssfCOMMONAPPDATA = &H23            '// All Users\Anwendungsdaten
  ssfWINDOWS = &H24                  '// GetWindowsDirectory()
  ssfSYSTEM = &H25                   '// GetSystemDirectory()
  ssfPROGRAMFILES = &H26             '// C:\Programme
  ssfMYPICTURES = &H27               '// Eigene Bilder
  ssfPROFILE = &H28                  '// Benutzerprofil
  ssfPROGRAMFILESCOMMON = &H2B       '// C:\Programme\Gemeinsame Dateien
  ssfCOMMONTEMPLATES = &H2D          '// All Users\Vorlagen
  ssfCOMMONDOCUMENTS = &H2E          '// All Users\Dokumente
  ssfCOMMONADMINTOOLS = &H2F         '// All Users\Startmenü\Programme\Verwaltung
  ssfADMINTOOLS = &H30               '// <Benutzer>\Startmenü\Programme\Verwaltung
End Enum

Private Type SHFILEINFO
  hIcon         As Long
  iIcon         As Long
  dwAttributes  As Long
  szDisplayName As String * MAX_PATH
  szTypeName    As String * 80
End Type

Private Declare Function SHGetFileInfo Lib "shell32.dll" _
        Alias "SHGetFileInfoA" ( _
        ByRef pszPath As Any, _
        ByVal dwFileAttributes As Long, _
        ByRef psfi As SHFILEINFO, _
        ByVal cbFileInfo As Long, _
        ByVal uFlags As Long _
              ) As Long

Private Const SHGFI_DISPLAYNAME As Long = &H200

Public Enum ShellNamespaceName
 DESKTOP_CLSID = 0
 INTERNET_CLSID = 1
 MYCOMPUTER_CLSID = 2
 MYFILES_CLSID = 3
 NETHOOD_CLSID = 4
 PRINTERS_CLSID = 5
 RECYCLEBIN_CLSID = 6
End Enum


Public Function GetSpecialFolder(ByVal Folder As ShellSpecialFolderConstants, Optional ByVal ForceCreate As Boolean) As String
 Dim tIIDL As ITEMIDLIST, strPath As String, hMod As Long

 If (ForceCreate) Then
  Folder = Folder Or CSIDL_FLAG_CREATE
 End If

 If SHGetSpecialFolderLocation(0, Folder, tIIDL) = S_OK Then
   strPath = Space$(MAX_PATH)
   If SHGetPathFromIDList(tIIDL.mkid.cb, strPath) <> 0 Then
    GetSpecialFolder = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
   End If
  Else
   strPath = Space$(MAX_PATH)
   hMod = LoadLibrary("shfolder")
   If (hMod <> 0) Then
    If SHGetFolderPath(0, Folder, 0, 0, strPath) = S_OK Then
     GetSpecialFolder = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
    End If
    FreeLibrary hMod
   End If
 End If
End Function

Public Function GetShellNamespaceName(ByVal Namespace As ShellNamespaceName) As String
 Dim lSHFI As SHFILEINFO, lRet As Long, CLSID(6) As String

 CLSID(0) = "{00021400-0000-0000-C000-000000000046}" ' DESKTOP_CLSID
 CLSID(0) = "{5984FFE0-28D4-11CF-AE66-08002B2E1262}" ' DESKTOP_CLSID
 CLSID(1) = "{871C5380-42A0-1069-A2EA-08002B30309D}" ' INTERNET_CLSID
 CLSID(2) = "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" ' MYCOMPUTER_CLSID
 CLSID(3) = "{450D8FBA-AD25-11D0-98A8-0800361B1103}" ' MYFILES_CLSID
 CLSID(4) = "{208D2C60-3AEA-1069-A2D7-08002B30309D}" ' NETHOOD_CLSID
 CLSID(5) = "{2227A280-3AEA-1069-A2DE-08002B30309D}" ' PRINTERS_CLSID
 CLSID(6) = "{645FF040-5081-101B-9F08-00AA002F954E}" ' RECYCLEBIN_CLSID

 lRet = SHGetFileInfo(ByVal "::" & CLSID(Namespace), 0, lSHFI, Len(lSHFI), SHGFI_DISPLAYNAME)

 If CBool(lRet) Then
  GetShellNamespaceName = Left$(lSHFI.szDisplayName, InStr(1, lSHFI.szDisplayName, vbNullChar) - 1)
 End If
End Function
