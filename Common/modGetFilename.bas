Attribute VB_Name = "modGetFilename"
Option Explicit

Public Function GetFilename( _
 Defaultfilename As String, InitPath As String, Filterindex As Long, _
 Filter As String, SaveOpenType As eSaveOpenType, Cancelled As Boolean, _
 Optional OwnerForm As Form = 0) As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim sFilter() As String, tStr As String, Files As Collection, Filename As String
50020
50030  If LenB(InitPath) = 0 Then
50040   InitPath = GetMyFiles
50050  End If
50060
50070 ' If IsWinNT4 = True Then
50080 '   With frmSaveOpen
50090 '    .Filename = Defaultfilename
50100 '    .Filterindex = Filterindex
50110 '    .SaveOpenType = SaveOpenType
50120 '    .Filter = Filter
50130 '    .Show vbModal, OwnerForm
50140 '   End With
50150 '   Set GetFilename = SaveOpenFilename
50160 '   Filterindex = SaveOpenFilterindex
50170 '  Else
50180    If SaveOpenType = saveFile Then
50190     tStr = "*.*"
50200     If InStr(Filter, "|") > 0 Then
50210      sFilter = Split(Filter, "|")
50220      If (Filterindex * 2 + 1) <= UBound(sFilter) Then
50230       tStr = sFilter(Filterindex * 2 + 1)
50240      End If
50250     End If
50260     Filterindex = SaveFileDialog(Filename, Defaultfilename, _
     Filter, tStr, InitPath, _
     LanguageStrings.SaveOpenSaveTitle, _
      OFN_EXPLORER + OFN_FILEMUSTEXIST + OFN_LONGNAMES + OFN_NODEREFERENCELINKS + OFN_HIDEREADONLY + OFN_OVERWRITEPROMPT, _
      OwnerForm.hwnd)
50310     Set GetFilename = New Collection
50320     GetFilename.Add Filename
50330     If Filterindex < 0 Then
50340       SaveOpenCancel = True: Cancelled = True
50350      Else
50360       SaveOpenCancel = False: Cancelled = False
50370     End If
50380    End If
50390    If SaveOpenType = OpenFile Then
50400     Filterindex = OpenFileDialog(Files, Defaultfilename, _
     Filter, tStr, InitPath, _
     LanguageStrings.SaveOpenOpenTitle, _
      OFN_ALLOWMULTISELECT + OFN_EXPLORER + OFN_FILEMUSTEXIST + OFN_LONGNAMES + OFN_NODEREFERENCELINKS, _
      OwnerForm.hwnd)
50450     Set GetFilename = Files
50460     If Filterindex < 0 Then
50470       SaveOpenCancel = True: Cancelled = True
50480      Else
50490       SaveOpenCancel = False: Cancelled = True
50500     End If
50510    End If
50520 ' End If
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

