Attribute VB_Name = "modPDF"
Option Explicit

'This variable will be used in frmMain and frmPrinting
Public PDFSpoolfile As String

Public Type tPSComment
 StartByte As Long
 EndByte As Long
 Comment As String
End Type

Public Type tPSHeader
 StartComment As tPSComment
 Title As tPSComment
 Creator As tPSComment
 CreationDate As tPSComment
 CreateFor As tPSComment
 EndComment As tPSComment
End Type

Public Function GetPDFTitle(FileName As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Const Buffer = 5000
50020  Dim fn As Long, bufStr As String, pTitleStart As String, pTitleEnd As String
50030
50040  If Len(Dir(FileName)) = 0 Then
50050   GetPDFTitle = vbNullString
50060   Exit Function
50070  End If
50080
50090  fn = FreeFile
50100  Open FileName For Input As #fn
50110  If FileLen(FileName) < Buffer Then
50120    bufStr = Input(FileLen(FileName), #fn)
50130   Else
50140    bufStr = Input(Buffer, #fn)
50150  End If
50160  Close #fn
50170
50180  pTitleStart = InStr(1, bufStr, "%%Title:", vbTextCompare)
50190
50200  If pTitleStart = 0 Then
50210   GetPDFTitle = vbNullString
50220   Exit Function
50230  End If
50240  pTitleEnd = InStr(pTitleStart, bufStr, vbLf, vbTextCompare)
50250
50260  If Mid$(bufStr, pTitleEnd - 1, 1) = vbCr Then
50270   pTitleEnd = pTitleEnd - 1
50280  End If
50290
50300  GetPDFTitle = Trim$(Mid$(bufStr, pTitleStart + 8, pTitleEnd - pTitleStart - 8))
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "GetPDFTitle")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub SavePDFTitle(FileName As String, Title As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Const Buffer = 5000
50020  Dim fn As Long, bufStr As String, pTitleStart As String, pTitleEnd As String, _
  oldTitle As String
50040
50050  If Dir(FileName) = "" Then
50060   Exit Sub
50070  End If
50080
50090  fn = FreeFile
50100  Open FileName For Input As #fn
50110  If LOF(fn) > 0 Then
50120   bufStr = Input(LOF(fn) - 1, #fn)
50130  End If
50140  Close #fn
50150
50160  pTitleStart = InStr(1, bufStr, "%%Title:", vbTextCompare)
50170  pTitleEnd = InStr(pTitleStart, bufStr, vbLf, vbTextCompare)
50180
50190  If Mid$(bufStr, pTitleEnd - 1, 1) = vbCr Then
50200   pTitleEnd = pTitleEnd - 1
50210  End If
50220
50230  If pTitleStart = 0 Then
50240   Exit Sub
50250  End If
50260  oldTitle = Trim$(Mid$(bufStr, pTitleStart + 8, pTitleEnd - pTitleStart - 8))
50270  Replace$ bufStr, oldTitle, Title
50280  Open FileName For Output As #fn
50290   Print #fn, bufStr
50300  Close #fn
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "SavePDFTitle")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function GetPSHeader(FileName As String) As tPSHeader
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long, bufStr As String, PSHeader As tPSHeader, Buffer As Long
50020  DoEvents
50030  fn = FreeFile
50040  If FileLen(FileName) = 0 Then
50050   Exit Function
50060  End If
50070  Buffer = 5000
50080  If FileLen(FileName) < Buffer Then
50090   Buffer = FileLen(FileName)
50100  End If
50110
50120  Open FileName For Binary Access Read As fn
50130  bufStr = Space$(Buffer)
50140  Get #fn, 1, bufStr
50150  Close #fn
50160
50170  With PSHeader
50180   .StartComment = GetPSComment(bufStr, "%!")
50190   .CreateFor = GetPSComment(bufStr, "%%For:")
50200   .CreationDate = GetPSComment(bufStr, "%%CreationDate:")
50210   .Creator = GetPSComment(bufStr, "%%Creator:")
50220   .Title = GetPSComment(bufStr, "%%Title:")
50230   .EndComment = GetPSComment(bufStr, "%%EndComments")
50240  End With
50250  GetPSHeader = PSHeader
50260  DoEvents
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "GetPSHeader")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub PutPSHeader(FileName As String, PSHeader As tPSHeader)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long, bufStr As String
50020  fn = FreeFile
50030  Open FileName For Binary Access Read As fn
50040  bufStr = Space$(LOF(fn))
50050  Get #fn, 1, bufStr
50060  Close #fn
50070  If PSHeader.StartComment.StartByte = -1 And PSHeader.EndComment.StartByte = -1 Then
50080   bufStr = GetPSHeaderString(PSHeader) & Chr(&HA) & bufStr
50090   fn = FreeFile
50100   Open FileName For Output As #fn
50110   Print #fn, bufStr
50120   Close #fn
50130   Exit Sub
50140  End If
50150  If PSHeader.StartComment.StartByte >= 0 And PSHeader.EndComment.StartByte >= 0 Then
50160   bufStr = Mid$(bufStr, 1, PSHeader.StartComment.StartByte - 1) & _
   GetPSHeaderString(PSHeader) & Mid$(bufStr, PSHeader.EndComment.EndByte + 1)
50180   fn = FreeFile
50190   Open FileName For Output As #fn
50200   Print #fn, bufStr
50210   Close #fn
50220   Exit Sub
50230  End If
50240  If PSHeader.StartComment.StartByte >= 0 And PSHeader.EndComment.StartByte = -1 Then
50250   bufStr = GetPSHeaderString(PSHeader) & Mid$(bufStr, PSHeader.StartComment.EndByte + 1)
50260   fn = FreeFile
50270   Open FileName For Output As #fn
50280   Print #fn, bufStr
50290   Close #fn
50300   Exit Sub
50310  End If
50320  If PSHeader.StartComment.StartByte = -1 And PSHeader.EndComment.StartByte >= 0 Then
50330   bufStr = GetPSHeaderString(PSHeader) & bufStr
50340   fn = FreeFile
50350   Open FileName For Output As #fn
50360   Print #fn, bufStr
50370   Close #fn
50380   Exit Sub
50390  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "PutPSHeader")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function GetPSComment(ByRef bufStr As String, Comment As String) As tPSComment
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim PSComment As tPSComment
50020  PSComment.StartByte = -1
50030  If InStr(UCase$(bufStr), UCase$(Comment)) > 0 Then
50040   With PSComment
50050    .StartByte = InStr(bufStr, Comment)
50060    .EndByte = InStr(.StartByte, bufStr, Chr$(&HA))
50070    If .EndByte - (.StartByte + Len(Comment)) > 0 Then
50080     .Comment = Mid$(bufStr, .StartByte + Len(Comment), .EndByte - (.StartByte + Len(Comment)))
50090     .Comment = Replace(Replace(.Comment, Chr$(&HA), ""), Chr$(&HD), "")
50100    End If
50110   End With
50120  End If
50130  GetPSComment = PSComment
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "GetPSComment")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function GetPSHeaderString(PSHeader As tPSHeader) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStr As String
50020  If PSHeader.StartComment.StartByte = -1 Then
50030    tStr = "%!PS-Adobe-3.0" & Chr$(&HA)
50040   Else
50050    tStr = "%!" & PSHeader.StartComment.Comment & Chr$(&HA)
50060  End If
50070
50080  tStr = tStr & "%%For:" & PSHeader.CreateFor.Comment & Chr$(&HA)
50090  tStr = tStr & "%%CreationDate:" & PSHeader.CreationDate.Comment & Chr$(&HA)
50100  tStr = tStr & "%%Creator:" & PSHeader.Creator.Comment & Chr$(&HA)
50110  tStr = tStr & "%%Title:" & PSHeader.Title.Comment & Chr$(&HA)
50120
50130  tStr = tStr & "%%EndComments" & Chr$(&HA)
50140  GetPSHeaderString = tStr
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "GetPSHeaderString")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetAutosaveFilename(Postscriptfile As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim FileName As String, Pathname As String
50020
50030  If Options.UseAutosaveDirectory = 1 Then
50040    Pathname = Trim$(Options.AutosaveDirectory)
50050   Else
50060    Pathname = Trim$(Options.LastSaveDirectory)
50070  End If
50080  If Right$(Pathname, 1) <> "\" Then
50090   Pathname = Pathname & "\"
50100  End If
50110
50120  FileName = GetSubstFilename(Postscriptfile, Options.AutosaveFilename)
50130
50140  Select Case Options.AutosaveFormat
  Case 0: 'PDF
50160    FileName = FileName & ".pdf"
50170   Case 1: 'PNG
50180    FileName = FileName & ".png"
50190   Case 2: 'JPEG
50200    FileName = FileName & ".jpg"
50210   Case 3: 'BMP
50220    FileName = FileName & ".bmp"
50230   Case 4: 'PCX
50240    FileName = FileName & ".pcx"
50250   Case 5: 'TIFF
50260    FileName = FileName & ".tif"
50270   Case 6: 'PS
50280    FileName = FileName & ".ps"
50290   Case 7: 'EPS
50300    FileName = FileName & ".eps"
50310  End Select
50320  GetAutosaveFilename = Pathname & ReplaceForbiddenChars(FileName)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "GetAutosaveFilename")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetSubstFilename(Postscriptfile As String, TokenFilename As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim PSHeader As tPSHeader, Author As String, _
  Title As String, Username As String, Computername As String, i As Long, _
  DateTime As String, FileName As String, tStr As String, tList() As String, _
  Subst() As String
50050
50060  DateTime = Format$(Now, "yyyymmddhhMMss")
50070  PSHeader = GetPSHeader(Postscriptfile)
50080  If Options.UseStandardAuthor = 1 Then
50090    Author = Options.StandardAuthor
50100   Else
50110    Author = PSHeader.CreateFor.Comment
50120  End If
50130
50140  If Options.FilenameSubstitutionsOnlyInTitle = 1 Then
50150   tList = Split(Options.FilenameSubstitutions, "\")
50160   Title = PSHeader.Title.Comment
50170   If UBound(tList) >= 0 Then
50180    For i = 0 To UBound(tList)
50190     Subst = Split(tList(i), "|")
50200     If UBound(Subst) = 0 Then
50210       tStr = ""
50220      Else
50230       tStr = Subst(1)
50240     End If
50250     Title = Replace(Title, Subst(0), tStr, , , vbTextCompare)
50260    Next i
50270   End If
50280  End If
50290
50300  Username = GetUsername
50310  Computername = GetComputerName
50320
50330  FileName = TokenFilename
50340  FileName = Replace(FileName, "<DateTime>", DateTime, , , vbTextCompare)
50350  FileName = Replace(FileName, "<Computername>", Computername, , , vbTextCompare)
50360  FileName = Replace(FileName, "<Username>", Username, , , vbTextCompare)
50370  FileName = Replace(FileName, "<Title>", Title, , , vbTextCompare)
50380  FileName = Replace(FileName, "<Author>", Author, , , vbTextCompare)
50390
50400  If Options.FilenameSubstitutionsOnlyInTitle = 0 Then
50410   tList = Split(Options.FilenameSubstitutions, "\")
50420   If UBound(tList) >= 0 Then
50430    For i = 0 To UBound(tList)
50440     Subst = Split(tList(i), "|")
50450     If UBound(Subst) = 0 Then
50460       tStr = ""
50470      Else
50480       tStr = Subst(1)
50490     End If
50500     FileName = Replace(FileName, Subst(0), tStr, , , vbTextCompare)
50510    Next i
50520   End If
50530  End If
50540  If Options.RemoveSpaces = 1 Then
50550   FileName = Trim$(FileName)
50560  End If
50570  GetSubstFilename = FileName
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "GetSubstFilename")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function CheckIfPSFile(ByVal FileName As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim lHeader As tPSHeader
50020  CheckIfPSFile = False
50030  lHeader = GetPSHeader(FileName)
50040  If InStr(1, lHeader.StartComment.Comment, "PS", vbTextCompare) > 0 Then
50050   CheckIfPSFile = True
50060  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "CheckIfPSFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
