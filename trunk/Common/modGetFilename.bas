Attribute VB_Name = "modGetFilename"
Option Explicit

Public Function GetFilename(Defaultfilename As String, InitPath As String, _
 Filterindex As Long, Filter As String, SaveOpenType As eSaveOpenType, _
 Cancelled As Boolean, Optional OwnerForm As Form = 0) As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020  Dim sFilter() As String, tStr As String, Files As Collection, _
  Filename As String
50040
50050  If LenB(InitPath) = 0 Then
50060   InitPath = GetMyFiles
50070  End If
50080
50090  If SaveOpenType = SaveFile Then
50100   tStr = "*.*"
50110   If InStr(Filter, "|") > 0 Then
50120    sFilter = Split(Filter, "|")
50130    If (Filterindex * 2 + 1) <= UBound(sFilter) Then
50140     tStr = sFilter(Filterindex * 2 + 1)
50150    End If
50160   End If
50170   Filterindex = SaveFileDialog(Filename, Defaultfilename, _
   Filter, tStr, InitPath, LanguageStrings.SaveOpenSaveTitle, _
    OFN_EXPLORER + OFN_PATHMUSTEXIST + OFN_LONGNAMES + OFN_HIDEREADONLY + _
    OFN_OVERWRITEPROMPT, OwnerForm.hwnd)
50210   If Filterindex < 0 Then
50220     SaveOpenCancel = True: Cancelled = True
50230    Else
50240     Set GetFilename = New Collection
50250     GetFilename.Add Filename
50260     SaveOpenCancel = False: Cancelled = False
50270   End If
50280  End If
50290  If SaveOpenType = OpenFile Then
50300   Filterindex = OpenFileDialog(Files, Defaultfilename, _
   Filter, tStr, InitPath, _
   LanguageStrings.SaveOpenOpenTitle, _
    OFN_ALLOWMULTISELECT + OFN_EXPLORER + OFN_FILEMUSTEXIST + OFN_LONGNAMES + OFN_NODEREFERENCELINKS, _
    OwnerForm.hwnd)
50350   Set GetFilename = Files
50360   If Filterindex < 0 Then
50370     SaveOpenCancel = True: Cancelled = True
50380    Else
50390     SaveOpenCancel = False: Cancelled = False
50400   End If
50410  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGetFilename", "GetFilename")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
