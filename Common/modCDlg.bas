Attribute VB_Name = "modCDlg"
Option Explicit

Public Function OpenFileDialog(Files As Collection, Optional InitFilename As String = "", _
 Optional Filter As String, Optional DefaultFileExtension As String = "*.*", _
 Optional InitDir As String = "", Optional DialogTitle As String = "", _
 Optional Flags As OpenSaveFlags, Optional hwnd As Long = 0) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  'res = -1 : UserCancel
50020  'res > 0  : Count of Files
50030
50040  Dim buff As String, tFil As String, buffA() As String, i As Long, ofn As OPENFILENAME
50050  If Len(Filter) > 0 Then
50060    If InStr(Filter, "|") > 0 Then
50070      tFil = Replace(Filter, "|", vbNullChar)
50080      tFil = tFil & vbNullChar & vbNullChar
50090     Else
50100      tFil = Filter & vbNullChar & vbNullChar
50110    End If
50120   Else
50130    tFil = vbNullChar & vbNullChar
50140  End If
50150
50160  With ofn
50170   .nStructSize = Len(ofn)
50180   .hWndOwner = hwnd
50190   .sFilter = tFil
50200   .nFilterIndex = 0
50210   .sFile = InitFilename & Space$(1024) & vbNullChar & vbNullChar
50220   .nMaxFile = Len(.sFile)
50230   .sDefFileExt = DefaultFileExtension & vbNullChar & vbNullChar
50240   .sFileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
50250   .nMaxTitle = Len(ofn.sFileTitle)
50260   If InitDir = vbNullString Then
50270     .sInitialDir = App.Path & vbNullChar & vbNullChar
50280    Else
50290     .sInitialDir = InitDir & vbNullChar & vbNullChar
50300   End If
50310   .sDialogTitle = DialogTitle
50320   .Flags = Flags
50330  End With
50340
50350  Set Files = New Collection
50360  If GetOpenFileName(ofn) <> 0 Then
50370    buff = Trim$(Replace$(Left$(ofn.sFile, Len(ofn.sFile) - 2), vbNullChar & vbNullChar, ""))
50380    Do While Right$(buff, 1) = vbNullChar
50390     buff = Mid(buff, 1, Len(buff) - 1)
50400     DoEvents
50410    Loop
50420    If Len(buff) > 3 Then
50430     If InStr(buff, vbNullChar) > 0 Then
50440       buffA = Split(buff, vbNullChar)
50450       For i = LBound(buffA) + 1 To UBound(buffA)
50460        If Len(buffA(i)) > 0 Then
50470         Files.Add CompletePath(buffA(0)) & buffA(i)
50480        End If
50490       Next i
50500      Else
50510       Files.Add buff
50520     End If
50530    End If
50540    OpenFileDialog = Files.Count
50550   Else
50560    OpenFileDialog = -1
50570  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modCDlg", "OpenFileDialog")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function SaveFileDialog(Filename As String, Optional InitFilename As String = "", _
 Optional Filter As String, Optional DefaultFileExtension As String = "*.*", _
 Optional InitDir As String = "", Optional DialogTitle As String = "", _
 Optional Flags As OpenSaveFlags, Optional hwnd As Long = 0) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  'res = -1 : UserCancel
50020  'res > 0  : Ok
50030
50040  Dim buff As String, tFil As String, buffA() As String, i As Long, ofn As OPENFILENAME
50050  If Len(Filter) > 0 Then
50060    If InStr(Filter, "|") > 0 Then
50070      tFil = Replace(Filter, "|", vbNullChar)
50080      tFil = tFil & vbNullChar & vbNullChar
50090     Else
50100      tFil = Filter & vbNullChar & vbNullChar
50110    End If
50120   Else
50130    tFil = vbNullChar & vbNullChar
50140  End If
50150
50160  With ofn
50170   .nStructSize = Len(ofn)
50180   .hWndOwner = hwnd
50190   .sFilter = tFil
50200   .nFilterIndex = 0
50210   .sFile = InitFilename & Space$(1024) & vbNullChar & vbNullChar
50220   .nMaxFile = Len(.sFile)
50230   .sDefFileExt = DefaultFileExtension & vbNullChar & vbNullChar
50240   .sFileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
50250   .nMaxTitle = Len(ofn.sFileTitle)
50260   If InitDir = vbNullString Then
50270     .sInitialDir = App.Path & vbNullChar & vbNullChar
50280    Else
50290     .sInitialDir = InitDir & vbNullChar & vbNullChar
50300   End If
50310   .sDialogTitle = DialogTitle
50320   .Flags = Flags
50330  End With
50340
50350  If GetSaveFileName(ofn) <> 0 Then
50360    buff = Trim$(Replace$(Left$(ofn.sFile, Len(ofn.sFile) - 2), vbNullChar & vbNullChar, ""))
50370    Do While Right$(buff, 1) = vbNullChar
50380     buff = Mid(buff, 1, Len(buff) - 1)
50390     DoEvents
50400    Loop
50410    If Len(buff) > 3 Then
50420     Filename = buff
50430    End If
50440    SaveFileDialog = ofn.nFilterIndex
50450   Else
50460    SaveFileDialog = -1
50470  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modCDlg", "SaveFileDialog")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
