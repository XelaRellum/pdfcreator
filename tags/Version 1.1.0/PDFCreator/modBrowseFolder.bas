Attribute VB_Name = "modBrowseFolder"
'This module contains all the declarations to use the
'Windows 95 Shell API to use the browse for folders
'dialog box.  To use the browse for folders dialog box,
'please call the BrowseForFolders function using the
'syntax: stringFolderPath=BrowseForFolders(Hwnd,TitleOfDialog)
'
'For more demo projects, please visit out web site at
'http://www.btinternet.com/~jelsoft/
'
'To contact us, please send an email to jelsoft@btinternet.com
'
' 08-Jun-2006
' Added browse for files by Frank Heindörfer
Option Explicit

Public Function BrowseForFolderFiles(hWndOwner As Long, sPrompt As String, Optional OnlyFolders As Boolean = True) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  'declare variables to be used
50020  Dim iNull As Integer, lpIDList As Long, lResult As Long, sPath As String, _
  udtBI As BrowseInfo
50040
50050  'initialise variables
50060  With udtBI
50070   .hWndOwner = hWndOwner
50080   .lpszTitle = lstrcat(sPrompt, "")
50090   If OnlyFolders Then
50100     .ulFlags = BIF_RETURNONLYFSDIRS
50110    Else
50120     .ulFlags = BIF_BROWSEINCLUDEFILES
50130   End If
50140  End With
50150
50160  'Call the browse for folder API
50170  lpIDList = SHBrowseForFolder(udtBI)
50180
50190  'get the resulting string path
50200  If lpIDList Then
50210   sPath = String$(MAX_PATH, 0)
50220   lResult = SHGetPathFromIDList(lpIDList, sPath)
50230   Call CoTaskMemFree(lpIDList)
50240   iNull = InStr(sPath, vbNullChar)
50250   If iNull Then sPath = Left$(sPath, iNull - 1)
50260  End If
50270
50280  'If cancel was pressed, sPath = ""
50290  BrowseForFolderFiles = sPath
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modBrowseFolder", "BrowseForFolderFiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
