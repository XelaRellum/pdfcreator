Attribute VB_Name = "modGetFilename"
Option Explicit

Public Function GetFilename( _
 Defaultfilename As String, InitPath As String, Filterindex As Long, _
 Filter As String, SaveOpenType As eSaveOpenType, _
 Cancelled As Boolean, Optional OwnerForm As Form = 0) As Collection
 Dim sFilter() As String, tStr As String, Files As Collection, Filename As String

 If LenB(InitPath) = 0 Then
  InitPath = GetMyFiles
 End If

 If WinNT4 = True Then
   With frmSaveOpen
    .Filename = Defaultfilename
    .Filterindex = Filterindex
    .SaveOpenType = SaveOpenType
    .Filter = Filter
    .Show vbModal, OwnerForm
   End With
   Set GetFilename = SaveOpenFilename
   Filterindex = SaveOpenFilterindex
   Cancelled = SaveOpenCancel
  Else
   If SaveOpenType = saveFile Then
    tStr = "*.*"
    If InStr(Filter, "|") > 0 Then
     sFilter = Split(Filter, "|")
     If (Filterindex * 2 + 1) <= UBound(sFilter) Then
      tStr = sFilter(Filterindex * 2 + 1)
     End If
    End If
    Filterindex = SaveFileDialog(Filename, Defaultfilename, _
     Filter, tStr, InitPath, _
     LanguageStrings.SaveOpenSaveTitle, _
      OFN_EXPLORER + OFN_FILEMUSTEXIST + OFN_LONGNAMES + OFN_NODEREFERENCELINKS + OFN_HIDEREADONLY, _
      OwnerForm.hwnd)
    Set GetFilename = New Collection
    GetFilename.Add Filename
    If Filterindex < 0 Then
      Cancelled = True
     Else
      Cancelled = False
    End If
   End If
   If SaveOpenType = OpenFile Then
    Filterindex = OpenFileDialog(Files, Defaultfilename, _
     Filter, tStr, InitPath, _
     LanguageStrings.SaveOpenOpenTitle, _
      OFN_ALLOWMULTISELECT + OFN_EXPLORER + OFN_FILEMUSTEXIST + OFN_LONGNAMES + OFN_NODEREFERENCELINKS, _
      OwnerForm.hwnd)
    Set GetFilename = Files
    If Filterindex < 0 Then
      Cancelled = True
     Else
      Cancelled = False
    End If
   End If
 End If
End Function



