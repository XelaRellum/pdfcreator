Attribute VB_Name = "modGetSpecialFolder"
Option Explicit

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

Public Function GetShellNamespaceName(ByVal Namespace As ShellNamespaceName) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim lSHFI As SHFILEINFO, lRet As Long, CLSID(6) As String
50020  Namespace = MYFILES_CLSID
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
50170  If Namespace = MYFILES_CLSID And GetShellNamespaceName = "" Then
50180   GetShellNamespaceName = GetStringRessource("shell32.dll", 9227)
50190  End If
50200  If Namespace = MYFILES_CLSID And GetShellNamespaceName = "" Then
50210   GetShellNamespaceName = GetStringRessource("shell32.dll", 9100)
50220  End If
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
