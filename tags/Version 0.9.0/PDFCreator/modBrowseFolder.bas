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

Option Explicit

Public Function BrowseForFolder(hWndOwner As Long, sPrompt As String) As String
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
50090   .ulFlags = BIF_RETURNONLYFSDIRS
50100  End With
50110
50120  'Call the browse for folder API
50130  lpIDList = SHBrowseForFolder(udtBI)
50140
50150  'get the resulting string path
50160  If lpIDList Then
50170   sPath = String$(MAX_PATH, 0)
50180   lResult = SHGetPathFromIDList(lpIDList, sPath)
50190   Call CoTaskMemFree(lpIDList)
50200   iNull = InStr(sPath, vbNullChar)
50210   If iNull Then sPath = Left$(sPath, iNull - 1)
50220  End If
50230
50240  'If cancel was pressed, sPath = ""
50250  BrowseForFolder = sPath
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modBrowseFolder", "BrowseForFolder")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
