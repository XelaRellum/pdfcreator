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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tIIDL As ITEMIDLIST, strPath As String, hMod As Long
50020
50030  If (ForceCreate) Then
50040   Folder = Folder Or CSIDL_FLAG_CREATE
50050  End If
50060
50070  If SHGetSpecialFolderLocation(0, Folder, tIIDL) = S_OK Then
50080    strPath = Space$(MAX_PATH)
50090    If SHGetPathFromIDList(tIIDL.mkid.cb, strPath) <> 0 Then
50100     GetSpecialFolder = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
50110    End If
50120   Else
50130    strPath = Space$(MAX_PATH)
50140    hMod = LoadLibrary("shfolder")
50150    If (hMod <> 0) Then
50160     If SHGetFolderPath(0, Folder, 0, 0, strPath) = S_OK Then
50170      GetSpecialFolder = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
50180     End If
50190     FreeLibrary hMod
50200    End If
50210  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGetSpecialFolder", "GetSpecialFolder")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetShellNamespaceName(ByVal Namespace As ShellNamespaceName) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim lSHFI As SHFILEINFO, lRet As Long, CLSID(6) As String
50020
50030  CLSID(0) = "{00021400-0000-0000-C000-000000000046}" ' DESKTOP_CLSID
50040  CLSID(0) = "{5984FFE0-28D4-11CF-AE66-08002B2E1262}" ' DESKTOP_CLSID
50050  CLSID(1) = "{871C5380-42A0-1069-A2EA-08002B30309D}" ' INTERNET_CLSID
50060  CLSID(2) = "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" ' MYCOMPUTER_CLSID
50070  CLSID(3) = "{450D8FBA-AD25-11D0-98A8-0800361B1103}" ' MYFILES_CLSID
50080  CLSID(4) = "{208D2C60-3AEA-1069-A2D7-08002B30309D}" ' NETHOOD_CLSID
50090  CLSID(5) = "{2227A280-3AEA-1069-A2DE-08002B30309D}" ' PRINTERS_CLSID
50100  CLSID(6) = "{645FF040-5081-101B-9F08-00AA002F954E}" ' RECYCLEBIN_CLSID
50110
50120  lRet = SHGetFileInfo(ByVal "::" & CLSID(Namespace), 0, lSHFI, Len(lSHFI), SHGFI_DISPLAYNAME)
50130
50140  If CBool(lRet) Then
50150   GetShellNamespaceName = Left$(lSHFI.szDisplayName, InStr(1, lSHFI.szDisplayName, vbNullChar) - 1)
50160  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGetSpecialFolder", "GetShellNamespaceName")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
