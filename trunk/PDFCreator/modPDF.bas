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

Public Type tPDFDocInfo
 Author As String
 CreationDate As String ' (D:20040222125120)
' CreationDate = "D:" & Format(Now, "YYYYMMDDHHNNSS")
 Creator As String
 Keywords As String
 ModifyDate As String '  (D:20040222125120)
 Subject As String
 Title As String
End Type

Public Function GetPDFTitle(Filename As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim psH As tPSHeader
50020  psH = GetPSHeader(Filename)
50030  GetPDFTitle = psH.Title.Comment
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

Public Sub SavePDFTitle(Filename As String, Title As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Const Buffer = 5000
50020  Dim fn As Long, bufStr As String, pTitleStart As String, pTitleEnd As String, _
  oldTitle As String
50040
50050  If Dir(Filename) = "" Then
50060   Exit Sub
50070  End If
50080
50090  fn = FreeFile
50100  Open Filename For Input As #fn
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
50280  Open Filename For Output As #fn
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

Public Function GetPSHeader(Filename As String) As tPSHeader
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long, bufStr As String, PSHeader As tPSHeader, Buffer As Long
50020  DoEvents
50030  fn = FreeFile
50040  If FileLen(Filename) = 0 Then
50050   Exit Function
50060  End If
50070  Buffer = 5000
50080  If FileLen(Filename) < Buffer Then
50090   Buffer = FileLen(Filename)
50100  End If
50110
50120  Open Filename For Binary Access Read As fn
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

Public Sub PutPSHeader(Filename As String, PSHeader As tPSHeader, Optional stb As StatusBar)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim TempHeaderFile As String, Tempfile1 As String, Tempfile2 As String, _
  fn As Long, Files As Collection, buf As Long
50030  If PSHeader.StartComment.StartByte = -1 And PSHeader.EndComment.StartByte = -1 Then
50040   TempHeaderFile = GetTempFile(CompletePath(GetPDFCreatorTempfolder), "~HE")
50050   Open TempHeaderFile For Output As #fn
50060   Print #fn, GetPSHeaderString(PSHeader) & Chr(&HA)
50070   Close #fn
50080   Tempfile1 = GetTempFile(CompletePath(GetPDFCreatorTempfolder), "~NF")
50090   Set Files = New Collection
50100   Files.Add TempHeaderFile
50110   Files.Add Filename
50120   CombineFiles Tempfile1, Files, stb
50130   Kill Filename: Kill TempHeaderFile
50140   Name Tempfile1 As Filename
50150   Exit Sub
50160  End If
50170  If PSHeader.StartComment.StartByte >= 0 And PSHeader.EndComment.StartByte >= 0 Then
50180   ' Complete PS-Header found
50190   '  Replace the existing the header
50200   Set Files = New Collection
50210   If PSHeader.StartComment.StartByte > 1 Then
50220    Tempfile1 = GetTempFile(CompletePath(GetPDFCreatorTempfolder), "~NF")
50230    CutFiles Filename, Tempfile1, 1, PSHeader.StartComment.StartByte - 1
50240    Files.Add Tempfile1
50250   End If
50260   TempHeaderFile = GetTempFile(CompletePath(GetPDFCreatorTempfolder), "~HE")
50270   fn = FreeFile
50280   Open TempHeaderFile For Output As #fn
50290   Print #fn, GetPSHeaderString(PSHeader) & Chr(&HA)
50300   Close #fn
50310   Files.Add TempHeaderFile
50320   Tempfile2 = GetTempFile(CompletePath(GetPDFCreatorTempfolder), "~NF")
50330   CutFiles Filename, Tempfile2, PSHeader.EndComment.EndByte + 1, FileLen(Filename)
50340   Files.Add Tempfile2
50350   Tempfile1 = GetTempFile(CompletePath(GetPDFCreatorTempfolder), "~NF")
50360   CombineFiles Tempfile1, Files, stb
50370   If Dir(Tempfile2) <> "" Then
50380    Kill Tempfile2
50390   End If
50400   Kill Filename
50410   ' Kill TempHeaderFile
50420   Name Tempfile1 As Filename
50430   Exit Sub
50440  End If
50450  If PSHeader.StartComment.StartByte >= 0 And PSHeader.EndComment.StartByte = -1 Then
50460   ' Incomplete PS-Header found (only Startcomment)
50470   '  Put the header on the beginning of the file
50480   TempHeaderFile = GetTempFile(CompletePath(GetPDFCreatorTempfolder), "~HE")
50490   Open TempHeaderFile For Output As #fn
50500   Print #fn, GetPSHeaderString(PSHeader) & Chr(&HA)
50510   Close #fn
50520   Tempfile1 = GetTempFile(CompletePath(GetPDFCreatorTempfolder), "~NF")
50530   Set Files = New Collection
50540   Files.Add TempHeaderFile
50550   Files.Add Filename
50560   CombineFiles Tempfile1, Files, stb
50570   Kill Filename: Kill TempHeaderFile
50580   Name Tempfile1 As Filename
50590   Exit Sub
50600  End If
50610  If PSHeader.StartComment.StartByte = -1 And PSHeader.EndComment.StartByte >= 0 Then
50620   ' Incomplete PS-Header found (only Endcomment)
50630   '  Put the header on the beginning of the file
50640   TempHeaderFile = GetTempFile(CompletePath(GetPDFCreatorTempfolder), "~HE")
50650   Open TempHeaderFile For Output As #fn
50660   Print #fn, GetPSHeaderString(PSHeader) & Chr(&HA)
50670   Close #fn
50680   Tempfile1 = GetTempFile(CompletePath(GetPDFCreatorTempfolder), "~NF")
50690   Set Files = New Collection
50700   Files.Add TempHeaderFile
50710   Files.Add Filename
50720   CombineFiles Tempfile1, Files, stb
50730   Kill Filename: Kill TempHeaderFile
50740   Name Tempfile1 As Filename
50750   Exit Sub
50760  End If
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

Public Sub PutPSHeaderAlt(Filename As String, PSHeader As tPSHeader)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long, bufStr As String
50020  fn = FreeFile
50030  Open Filename For Binary Access Read As fn
50040  bufStr = Space$(LOF(fn))
50050  Get #fn, 1, bufStr
50060  Close #fn
50070  If PSHeader.StartComment.StartByte = -1 And PSHeader.EndComment.StartByte = -1 Then
50080   bufStr = GetPSHeaderString(PSHeader) & Chr(&HA) & bufStr
50090   fn = FreeFile
50100   Open Filename For Output As #fn
50110   Print #fn, bufStr
50120   Close #fn
50130   Exit Sub
50140  End If
50150  If PSHeader.StartComment.StartByte >= 0 And PSHeader.EndComment.StartByte >= 0 Then
50160   bufStr = Mid$(bufStr, 1, PSHeader.StartComment.StartByte - 1) & _
   GetPSHeaderString(PSHeader) & Mid$(bufStr, PSHeader.EndComment.EndByte + 1)
50180   fn = FreeFile
50190   Open Filename For Output As #fn
50200   Print #fn, bufStr
50210   Close #fn
50220   Exit Sub
50230  End If
50240  If PSHeader.StartComment.StartByte >= 0 And PSHeader.EndComment.StartByte = -1 Then
50250   bufStr = GetPSHeaderString(PSHeader) & Mid$(bufStr, PSHeader.StartComment.EndByte + 1)
50260   fn = FreeFile
50270   Open Filename For Output As #fn
50280   Print #fn, bufStr
50290   Close #fn
50300   Exit Sub
50310  End If
50320  If PSHeader.StartComment.StartByte = -1 And PSHeader.EndComment.StartByte >= 0 Then
50330   bufStr = GetPSHeaderString(PSHeader) & bufStr
50340   fn = FreeFile
50350   Open Filename For Output As #fn
50360   Print #fn, bufStr
50370   Close #fn
50380   Exit Sub
50390  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "PutPSHeaderAlt")
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
50110    .Comment = Trim$(.Comment)
50120    If Len(.Comment) > 0 Then
50130     If Mid$(.Comment, 1, 1) = "(" Then
50140      .Comment = Mid(.Comment, 2)
50150     End If
50160     If Len(.Comment) > 0 Then
50170      If Mid$(.Comment, Len(.Comment), 1) = ")" Then
50180       .Comment = Mid(.Comment, 1, Len(.Comment) - 1)
50190      End If
50200     End If
50210    End If
50220    .Comment = ReplaceEncodingChars(.Comment)
50230   End With
50240  End If
50250  GetPSComment = PSComment
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
50010  Dim Filename As String, Pathname As String
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
50120  Filename = GetSubstFilename(Postscriptfile, Options.AutosaveFilename)
50130
50141  Select Case Options.AutosaveFormat
              Case 0: 'PDF
50160    Filename = Filename & ".pdf"
50170   Case 1: 'PNG
50180    Filename = Filename & ".png"
50190   Case 2: 'JPEG
50200    Filename = Filename & ".jpg"
50210   Case 3: 'BMP
50220    Filename = Filename & ".bmp"
50230   Case 4: 'PCX
50240    Filename = Filename & ".pcx"
50250   Case 5: 'TIFF
50260    Filename = Filename & ".tif"
50270   Case 6: 'PS
50280    Filename = Filename & ".ps"
50290   Case 7: 'EPS
50300    Filename = Filename & ".eps"
50310  End Select
50320  GetAutosaveFilename = Pathname & ReplaceForbiddenChars(Filename)
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

Public Function GetSubstFilename(Postscriptfile As String, TokenFilename As String, _
 Optional WithoutAuthor As Boolean = False, Optional Preview As Boolean = False) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim PSHeader As tPSHeader, Author As String, _
  Title As String, UserName As String, Computername As String, i As Long, _
  DateTime As String, Filename As String, tStr As String, tList() As String, _
  Subst() As String, UserProfilPath As String, MyFiles As String
50050
50060  DateTime = Format$(Now, "yyyymmddhhMMss")
50070  If Preview = False Then
50080    PSHeader = GetPSHeader(Postscriptfile)
50090    If Options.UseStandardAuthor = 1 Then
50100      Author = Options.StandardAuthor
50110     Else
50120      Author = PSHeader.CreateFor.Comment
50130    End If
50140  Else
50150    PSHeader.Title.Comment = "'Preview Title'"
50160    Author = "'Preview Author'"
50170  End If
50180
50190  If Options.FilenameSubstitutionsOnlyInTitle = 1 Then
50200   tList = Split(Options.FilenameSubstitutions, "\")
50210   Title = PSHeader.Title.Comment
50220   If UBound(tList) >= 0 Then
50230    For i = 0 To UBound(tList)
50240     Subst = Split(tList(i), "|")
50250     If UBound(Subst) = 0 Then
50260       tStr = ""
50270      Else
50280       tStr = Subst(1)
50290     End If
50300     Title = Replace(Title, Subst(0), tStr, , , vbTextCompare)
50310    Next i
50320   End If
50330  End If
50340
50350  UserName = GetUsername
50360  Computername = GetComputerName
50370  MyFiles = GetMyFiles
50380
50390  Filename = TokenFilename
50400  Filename = Replace(Filename, "<DateTime>", DateTime, , , vbTextCompare)
50410  Filename = Replace(Filename, "<Computername>", Computername, , , vbTextCompare)
50420  Filename = Replace(Filename, "<Username>", UserName, , , vbTextCompare)
50430  Filename = Replace(Filename, "<Title>", Title, , , vbTextCompare)
50440  If WithoutAuthor = False Then
50450   Filename = Replace(Filename, "<Author>", Author, , , vbTextCompare)
50460  End If
50470  Filename = Replace(Filename, "<MyFiles>", MyFiles, , , vbTextCompare)
50480
50490  If Options.FilenameSubstitutionsOnlyInTitle = 0 Then
50500   tList = Split(Options.FilenameSubstitutions, "\")
50510   If UBound(tList) >= 0 Then
50520    For i = 0 To UBound(tList)
50530     Subst = Split(tList(i), "|")
50540     If UBound(Subst) = 0 Then
50550       tStr = ""
50560      Else
50570       tStr = Subst(1)
50580     End If
50590     Filename = Replace(Filename, Subst(0), tStr, , , vbTextCompare)
50600    Next i
50610   End If
50620  End If
50630  If Options.RemoveSpaces = 1 Then
50640   Filename = Trim$(Filename)
50650  End If
50660  GetSubstFilename = Filename
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

Public Function CheckIfPSFile(ByVal Filename As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim lHeader As tPSHeader
50020  CheckIfPSFile = False
50030  lHeader = GetPSHeader(Filename)
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

Public Sub ReplacePSHeaderInFile(PSHeader As tPSHeader, Filename As String, _
 Optional BufferSize As Long = 65536)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim bsize As Long, j As Long, fnSource As Long, fnDest As Long, fpos As Long, _
  tLen As Long, sBuffer As String, offset As Long, Tempfile As String, _
  TempHeaderFile As String
50040
50050  Open Filename For Input As #fnSource
50060  TempHeaderFile = GetTempFile(CompletePath(GetPDFCreatorTempfolder), "~RP")
50070  Open Tempfile For Output As #fnDest
50080
50090  ' All bytes befor
50100  bsize = BufferSize
50110  offset = PSHeader.StartComment.StartByte - 1
50120  For j = 1 To offset \ bsize
50130   fpos = (j - 1) * bsize + 1 + offset
50140   Seek #fnSource, fpos
50150   sBuffer = Input(bsize, fnSource)
50160   Put #fnDest, , sBuffer
50170   tLen = tLen + bsize
50180   DoEvents
50190  Next j
50200  If LOF(offset) > (j - 1) * bsize Then
50210   fpos = (j - 1) * bsize + 1 + offset
50220   Seek #fnSource, fpos
50230   sBuffer = Input(offset - (j - 1) * bsize, fnSource)
50240   Put #fnDest, , sBuffer
50250   tLen = tLen + (offset - (j - 1) * bsize)
50260  End If
50270
50280  ' The PsHeader
50290  Put #fnDest, , GetPSHeaderString(PSHeader)
50300
50310  ' All bytes after
50320  offset = PSHeader.EndComment.EndByte + 1
50330  For j = 1 To (LOF(fnSource) - offset + 1) \ bsize
50340   fpos = (j - 1) * bsize + 1 + offset
50350   Seek #fnSource, fpos
50360   sBuffer = Input(bsize, fnSource)
50370   Put #fnDest, , sBuffer
50380   tLen = tLen + bsize
50390   DoEvents
50400  Next j
50410  If (LOF(fnSource) - offset + 1) > (j - 1) * bsize Then
50420   fpos = (j - 1) * bsize + 1 + offset
50430   Seek #fnSource, fpos
50440   sBuffer = Input(LOF(fnSource) - (j - 1) * bsize, fnSource)
50450   Put #fnDest, , sBuffer
50460   tLen = tLen + (LOF(fnSource) - (j - 1) * bsize)
50470  End If
50480  Close #fnSource: Close #fnDest
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "ReplacePSHeaderInFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub CutFiles(InFile As String, OutFile As String, StartByte As Long, EndByte As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fnSource As Long, fnDest As Long, buf As Long, tLen As String, i As Long, _
  sBuffer As String
50030
50040
50050  buf = 65536
50060
50070  fnSource = FreeFile
50080  Open InFile For Binary As #fnSource
50090  fnDest = FreeFile
50100  Open OutFile For Output As #fnDest
50110
50120  'Read until StartByte - 1
50130  tLen = EndByte - StartByte + 1
50140
50150
50160  For i = 1 To tLen \ buf
50170   Seek #fnSource, StartByte + (i - 1) * buf
50180   sBuffer = Input(buf, fnSource)
50190   Print #fnDest, sBuffer;
50200  Next i
50210  If tLen > (i - 1) * buf Then
50220   Seek #fnSource, StartByte + (i - 1) * buf
50230   sBuffer = Input(tLen - buf * (i - 1), fnSource)
50240   Print #fnDest, sBuffer;
50250  End If
50260  Close fnDest: Close fnSource
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "CutFiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub AppendPDFDocInfo(PSFile As String, PDFDocInfo As tPDFDocInfo)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long, DocInfoStr As String
50020  If LenB(Dir(PSFile)) > 0 Then
50030   DocInfoStr = Chr$(13) & "%!PS-Adobe-3.0 EPSF-3.0"
50040   DocInfoStr = DocInfoStr & Chr$(13) & "%%BoundingBox: 0 0 72 72"
50050   DocInfoStr = DocInfoStr & Chr$(13) & "%%EndProlog"
50060   DocInfoStr = DocInfoStr & Chr$(13) & "/pdfmark where {pop} {userdict /pdfmark /cleartomark load put} ifelse"
50070   DocInfoStr = DocInfoStr & Chr$(13) & "["
50080   With PDFDocInfo
50090    DocInfoStr = DocInfoStr & "/Author (" & .Author & ")"
50100    DocInfoStr = DocInfoStr & Chr$(13) & "/CreationDate (D:" & .CreationDate & ")"
50110    DocInfoStr = DocInfoStr & Chr$(13) & "/Creator (" & .Creator & ")"
50120    DocInfoStr = DocInfoStr & Chr$(13) & "/Keywords (" & .Keywords & ")"
50130    DocInfoStr = DocInfoStr & Chr$(13) & "/ModDate (D:" & .ModifyDate & ")"
50140    DocInfoStr = DocInfoStr & Chr$(13) & "/Subject (" & .Subject & ")"
50150    DocInfoStr = DocInfoStr & Chr$(13) & "/Title (" & .Title & ")"
50160   End With
50170   DocInfoStr = DocInfoStr & Chr$(13) & "/DOCINFO pdfmark"
50180   DocInfoStr = DocInfoStr & Chr$(13) & "%%EOF"
50190   fn = FreeFile
50200   Open PSFile For Append As fn
50210   Print #fn, DocInfoStr;
50220   Close #fn
50230  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "AppendPDFDocInfo")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function ReplaceEncodingChars(Str1 As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Integer, tStr As String
50020  tStr = ""
50030  ' First we look for oct encoding chars
50040  For i = 127 To 255
50050   Str1 = Replace$(Str1, "\" & Oct$(i), Chr$(i))
50060  Next i
50070  ReplaceEncodingChars = Str1
50080  ' Second we look for hex encoding chars
50090  If Len(Str1) >= 4 Then
50100   If Mid$(Str1, 1, 1) = "<" And Mid$(Str1, Len(Str1), 1) = ">" Then
50110    If Len(Str1) Mod 2 = 0 Then
50120     For i = 2 To Len(Str1) - 1 Step 2
50130      If IsNumeric("&H" & Mid$(Str1, i, 2)) = True Then
50140        If CByte("&H" & Mid$(Str1, i, 2)) > 255 Then
50150          Exit Function
50160         Else
50170          tStr = tStr & Chr$(CByte("&H" & Mid$(Str1, i, 2)))
50180        End If
50190       Else
50200        Exit Function
50210      End If
50220     Next i
50230    End If
50240   End If
50250  End If
50260  If Len(tStr) > 0 Then
50270   ReplaceEncodingChars = tStr
50280  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "ReplaceEncodingChars")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
