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
 Dim lSHFI As SHFILEINFO, lRet As Long, CLSID(6) As String
 Namespace = MYFILES_CLSID
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
 If Namespace = MYFILES_CLSID And GetShellNamespaceName = "" Then
  GetShellNamespaceName = GetStringRessource("shell32.dll", 9227)
 End If
 If Namespace = MYFILES_CLSID And GetShellNamespaceName = "" Then
  GetShellNamespaceName = GetStringRessource("shell32.dll", 9100)
 End If
End Function
