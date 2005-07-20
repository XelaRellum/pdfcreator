Attribute VB_Name = "modGetFilename"
Option Explicit

Public Function GetFilename(Defaultfilename As String, InitPath As String, _
 FilterIndex As Long, Filter As String, SaveOpenType As eSaveOpenType, _
 Cancelled As Boolean, Optional OwnerForm As Form = 0) As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020  Dim sFilter() As String, tstr As String, Files As Collection, _
  Filename As String, Ext As String
50040
50050  If LenB(InitPath) = 0 Then
50060   InitPath = GetMyFiles
50070  End If
50080
50090  If SaveOpenType = SaveFile Then
50100   tstr = "*.*"
50110   If InStr(Filter, "|") > 0 Then
50120    sFilter = Split(Filter, "|")
50130    If (FilterIndex * 2 + 1) <= UBound(sFilter) Then
50140     tstr = sFilter(FilterIndex * 2 + 1)
50150    End If
50160   End If
50170   FilterIndex = SaveFileDialog(Filename, Defaultfilename, _
   Filter, tstr, InitPath, LanguageStrings.SaveOpenSaveTitle, _
    OFN_EXPLORER + OFN_PATHMUSTEXIST + OFN_LONGNAMES + OFN_HIDEREADONLY + _
    OFN_OVERWRITEPROMPT, OwnerForm.hwnd, FilterIndex)
50210   If FilterIndex < 0 Then
50220     SaveOpenCancel = True: Cancelled = True
50230    Else
50240 '    tStr = "*.*"
50250 '    If InStr(Filter, "|") > 0 Then
50260 '     sFilter = Split(Filter, "|")
50270 '     If (Filterindex * 2 - 1) <= UBound(sFilter) Then
50280 '      tStr = sFilter(Filterindex * 2 - 1)
50290 '     End If
50300 '    End If
50310 '    If Len(tStr) > 0 Then
50320 '     If InStr(tStr, ".") > 0 Then
50330 '      If InStr(InStrRev(tStr, "."), "*") <= 0 Then
50340 '       SplitPath Filename, , , , , Ext
50350 '       If UCase$(Mid(tStr, InStrRev(tStr, ".") + 1)) <> UCase$(Ext) Then
50360 '        Filename = Filename & "." & LCase$(Mid(tStr, InStrRev(tStr, ".") + 1))
50370 '       End If
50380 '      End If
50390 '     End If
50400 '    End If
50410     Set GetFilename = New Collection
50420     GetFilename.Add Filename
50430     SaveOpenCancel = False: Cancelled = False
50440   End If
50450  End If
50460  If SaveOpenType = OpenFile Then
50470   FilterIndex = OpenFileDialog(Files, Defaultfilename, _
   Filter, tstr, InitPath, _
   LanguageStrings.SaveOpenOpenTitle, _
    OFN_ALLOWMULTISELECT + OFN_EXPLORER + OFN_FILEMUSTEXIST + OFN_LONGNAMES + OFN_NODEREFERENCELINKS, _
    OwnerForm.hwnd)
50520   Set GetFilename = Files
50530   If FilterIndex < 0 Then
50540     SaveOpenCancel = True: Cancelled = True
50550    Else
50560     SaveOpenCancel = False: Cancelled = False
50570   End If
50580  End If
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
