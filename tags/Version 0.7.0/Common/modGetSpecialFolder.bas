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

Public Function GetSpecialFolder(ByVal Folder As ShellSpecialFolderConstants, _
                        Optional ByVal ForceCreate As Boolean) As String
  Dim tIIDL   As ITEMIDLIST
  Dim strPath As String
  Dim hMod    As Long

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
