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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010   Dim tIIDL   As ITEMIDLIST
50020   Dim strPath As String
50030   Dim hMod    As Long
50040
50050   If (ForceCreate) Then
50060     Folder = Folder Or CSIDL_FLAG_CREATE
50070   End If
50080
50090   If SHGetSpecialFolderLocation(0, Folder, tIIDL) = S_OK Then
50100     strPath = Space$(MAX_PATH)
50110     If SHGetPathFromIDList(tIIDL.mkid.cb, strPath) <> 0 Then
50120       GetSpecialFolder = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
50130     End If
50140   Else
50150     strPath = Space$(MAX_PATH)
50160     hMod = LoadLibrary("shfolder")
50170     If (hMod <> 0) Then
50180       If SHGetFolderPath(0, Folder, 0, 0, strPath) = S_OK Then
50190         GetSpecialFolder = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
50200       End If
50210       FreeLibrary hMod
50220     End If
50230   End If
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
