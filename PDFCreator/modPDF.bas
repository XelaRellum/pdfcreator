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
 Const Buffer = 5000
 Dim fn As Long, bufStr As String, pTitleStart As String, pTitleEnd As String

 If Len(Dir(Filename)) = 0 Then
  GetPDFTitle = vbNullString
  Exit Function
 End If

 fn = FreeFile
 Open Filename For Input As #fn
 If FileLen(Filename) < Buffer Then
   bufStr = Input(FileLen(Filename), #fn)
  Else
   bufStr = Input(Buffer, #fn)
 End If
 Close #fn

 pTitleStart = InStr(1, bufStr, "%%Title:", vbTextCompare)

 If pTitleStart = 0 Then
  GetPDFTitle = vbNullString
  Exit Function
 End If
 pTitleEnd = InStr(pTitleStart, bufStr, vbLf, vbTextCompare)

 If Mid$(bufStr, pTitleEnd - 1, 1) = vbCr Then
  pTitleEnd = pTitleEnd - 1
 End If

 GetPDFTitle = Trim$(Mid$(bufStr, pTitleStart + 8, pTitleEnd - pTitleStart - 8))
End Function

Public Sub SavePDFTitle(Filename As String, Title As String)
 Const Buffer = 5000
 Dim fn As Long, bufStr As String, pTitleStart As String, pTitleEnd As String, _
  oldTitle As String

 If Dir(Filename) = "" Then
  Exit Sub
 End If

 fn = FreeFile
 Open Filename For Input As #fn
 If LOF(fn) > 0 Then
  bufStr = Input(LOF(fn) - 1, #fn)
 End If
 Close #fn

 pTitleStart = InStr(1, bufStr, "%%Title:", vbTextCompare)
 pTitleEnd = InStr(pTitleStart, bufStr, vbLf, vbTextCompare)

 If Mid$(bufStr, pTitleEnd - 1, 1) = vbCr Then
  pTitleEnd = pTitleEnd - 1
 End If

 If pTitleStart = 0 Then
  Exit Sub
 End If
 oldTitle = Trim$(Mid$(bufStr, pTitleStart + 8, pTitleEnd - pTitleStart - 8))
 Replace$ bufStr, oldTitle, Title
 Open Filename For Output As #fn
  Print #fn, bufStr
 Close #fn
End Sub

Public Function GetPSHeader(Filename As String) As tPSHeader
 Dim fn As Long, bufStr As String, PSHeader As tPSHeader, Buffer As Long
 DoEvents
 fn = FreeFile
 If FileLen(Filename) = 0 Then
  Exit Function
 End If
 Buffer = 5000
 If FileLen(Filename) < Buffer Then
  Buffer = FileLen(Filename)
 End If

 Open Filename For Binary Access Read As fn
 bufStr = Space$(Buffer)
 Get #fn, 1, bufStr
 Close #fn

 With PSHeader
  .StartComment = GetPSComment(bufStr, "%!")
  .CreateFor = GetPSComment(bufStr, "%%For:")
  .CreationDate = GetPSComment(bufStr, "%%CreationDate:")
  .Creator = GetPSComment(bufStr, "%%Creator:")
  .Title = GetPSComment(bufStr, "%%Title:")
  .EndComment = GetPSComment(bufStr, "%%EndComments")
 End With
 GetPSHeader = PSHeader
 DoEvents
End Function

Public Sub PutPSHeader(Filename As String, PSHeader As tPSHeader, Optional stb As StatusBar)
 Dim TempHeaderFile As String, Tempfile1 As String, Tempfile2 As String, _
  fn As Long, Files As Collection, buf As Long
 If PSHeader.StartComment.StartByte = -1 And PSHeader.EndComment.StartByte = -1 Then
  TempHeaderFile = GetTempFile(CompletePath(GetPDFCreatorTempfolder), "~HE")
  Open TempHeaderFile For Output As #fn
  Print #fn, GetPSHeaderString(PSHeader) & Chr(&HA)
  Close #fn
  Tempfile1 = GetTempFile(CompletePath(GetPDFCreatorTempfolder), "~NF")
  Set Files = New Collection
  Files.Add TempHeaderFile
  Files.Add Filename
  CombineFiles Tempfile1, Files, stb
  Kill Filename: Kill TempHeaderFile
  Name Tempfile1 As Filename
  Exit Sub
 End If
 If PSHeader.StartComment.StartByte >= 0 And PSHeader.EndComment.StartByte >= 0 Then
  ' Complete PS-Header found
  '  Replace the existing the header
  Set Files = New Collection
  If PSHeader.StartComment.StartByte > 1 Then
   Tempfile1 = GetTempFile(CompletePath(GetPDFCreatorTempfolder), "~NF")
   CutFiles Filename, Tempfile1, 1, PSHeader.StartComment.StartByte - 1
   Files.Add Tempfile1
  End If
  TempHeaderFile = GetTempFile(CompletePath(GetPDFCreatorTempfolder), "~HE")
  fn = FreeFile
  Open TempHeaderFile For Output As #fn
  Print #fn, GetPSHeaderString(PSHeader) & Chr(&HA)
  Close #fn
  Files.Add TempHeaderFile
  Tempfile2 = GetTempFile(CompletePath(GetPDFCreatorTempfolder), "~NF")
  CutFiles Filename, Tempfile2, PSHeader.EndComment.EndByte + 1, FileLen(Filename)
  Files.Add Tempfile2
  Tempfile1 = GetTempFile(CompletePath(GetPDFCreatorTempfolder), "~NF")
  CombineFiles Tempfile1, Files, stb
  If Dir(Tempfile2) <> "" Then
   Kill Tempfile2
  End If
  Kill Filename
  ' Kill TempHeaderFile
  Name Tempfile1 As Filename
  Exit Sub
 End If
 If PSHeader.StartComment.StartByte >= 0 And PSHeader.EndComment.StartByte = -1 Then
  ' Incomplete PS-Header found (only Startcomment)
  '  Put the header on the beginning of the file
  TempHeaderFile = GetTempFile(CompletePath(GetPDFCreatorTempfolder), "~HE")
  Open TempHeaderFile For Output As #fn
  Print #fn, GetPSHeaderString(PSHeader) & Chr(&HA)
  Close #fn
  Tempfile1 = GetTempFile(CompletePath(GetPDFCreatorTempfolder), "~NF")
  Set Files = New Collection
  Files.Add TempHeaderFile
  Files.Add Filename
  CombineFiles Tempfile1, Files, stb
  Kill Filename: Kill TempHeaderFile
  Name Tempfile1 As Filename
  Exit Sub
 End If
 If PSHeader.StartComment.StartByte = -1 And PSHeader.EndComment.StartByte >= 0 Then
  ' Incomplete PS-Header found (only Endcomment)
  '  Put the header on the beginning of the file
  TempHeaderFile = GetTempFile(CompletePath(GetPDFCreatorTempfolder), "~HE")
  Open TempHeaderFile For Output As #fn
  Print #fn, GetPSHeaderString(PSHeader) & Chr(&HA)
  Close #fn
  Tempfile1 = GetTempFile(CompletePath(GetPDFCreatorTempfolder), "~NF")
  Set Files = New Collection
  Files.Add TempHeaderFile
  Files.Add Filename
  CombineFiles Tempfile1, Files, stb
  Kill Filename: Kill TempHeaderFile
  Name Tempfile1 As Filename
  Exit Sub
 End If
End Sub

Public Sub PutPSHeaderAlt(Filename As String, PSHeader As tPSHeader)
 Dim fn As Long, bufStr As String
 fn = FreeFile
 Open Filename For Binary Access Read As fn
 bufStr = Space$(LOF(fn))
 Get #fn, 1, bufStr
 Close #fn
 If PSHeader.StartComment.StartByte = -1 And PSHeader.EndComment.StartByte = -1 Then
  bufStr = GetPSHeaderString(PSHeader) & Chr(&HA) & bufStr
  fn = FreeFile
  Open Filename For Output As #fn
  Print #fn, bufStr
  Close #fn
  Exit Sub
 End If
 If PSHeader.StartComment.StartByte >= 0 And PSHeader.EndComment.StartByte >= 0 Then
  bufStr = Mid$(bufStr, 1, PSHeader.StartComment.StartByte - 1) & _
   GetPSHeaderString(PSHeader) & Mid$(bufStr, PSHeader.EndComment.EndByte + 1)
  fn = FreeFile
  Open Filename For Output As #fn
  Print #fn, bufStr
  Close #fn
  Exit Sub
 End If
 If PSHeader.StartComment.StartByte >= 0 And PSHeader.EndComment.StartByte = -1 Then
  bufStr = GetPSHeaderString(PSHeader) & Mid$(bufStr, PSHeader.StartComment.EndByte + 1)
  fn = FreeFile
  Open Filename For Output As #fn
  Print #fn, bufStr
  Close #fn
  Exit Sub
 End If
 If PSHeader.StartComment.StartByte = -1 And PSHeader.EndComment.StartByte >= 0 Then
  bufStr = GetPSHeaderString(PSHeader) & bufStr
  fn = FreeFile
  Open Filename For Output As #fn
  Print #fn, bufStr
  Close #fn
  Exit Sub
 End If
End Sub

Private Function GetPSComment(ByRef bufStr As String, Comment As String) As tPSComment
 Dim PSComment As tPSComment
 PSComment.StartByte = -1
 If InStr(UCase$(bufStr), UCase$(Comment)) > 0 Then
  With PSComment
   .StartByte = InStr(bufStr, Comment)
   .EndByte = InStr(.StartByte, bufStr, Chr$(&HA))
   If .EndByte - (.StartByte + Len(Comment)) > 0 Then
    .Comment = Mid$(bufStr, .StartByte + Len(Comment), .EndByte - (.StartByte + Len(Comment)))
    .Comment = Replace(Replace(.Comment, Chr$(&HA), ""), Chr$(&HD), "")
   End If
  End With
 End If
 GetPSComment = PSComment
End Function

Private Function GetPSHeaderString(PSHeader As tPSHeader) As String
 Dim tStr As String
 If PSHeader.StartComment.StartByte = -1 Then
   tStr = "%!PS-Adobe-3.0" & Chr$(&HA)
  Else
   tStr = "%!" & PSHeader.StartComment.Comment & Chr$(&HA)
 End If

 tStr = tStr & "%%For:" & PSHeader.CreateFor.Comment & Chr$(&HA)
 tStr = tStr & "%%CreationDate:" & PSHeader.CreationDate.Comment & Chr$(&HA)
 tStr = tStr & "%%Creator:" & PSHeader.Creator.Comment & Chr$(&HA)
 tStr = tStr & "%%Title:" & PSHeader.Title.Comment & Chr$(&HA)

 tStr = tStr & "%%EndComments" & Chr$(&HA)
 GetPSHeaderString = tStr
End Function

Public Function GetAutosaveFilename(Postscriptfile As String) As String
 Dim Filename As String, Pathname As String

 If Options.UseAutosaveDirectory = 1 Then
   Pathname = Trim$(Options.AutosaveDirectory)
  Else
   Pathname = Trim$(Options.LastSaveDirectory)
 End If
 If Right$(Pathname, 1) <> "\" Then
  Pathname = Pathname & "\"
 End If

 Filename = GetSubstFilename(Postscriptfile, Options.AutosaveFilename)

 Select Case Options.AutosaveFormat
        Case 0: 'PDF
   Filename = Filename & ".pdf"
  Case 1: 'PNG
   Filename = Filename & ".png"
  Case 2: 'JPEG
   Filename = Filename & ".jpg"
  Case 3: 'BMP
   Filename = Filename & ".bmp"
  Case 4: 'PCX
   Filename = Filename & ".pcx"
  Case 5: 'TIFF
   Filename = Filename & ".tif"
  Case 6: 'PS
   Filename = Filename & ".ps"
  Case 7: 'EPS
   Filename = Filename & ".eps"
 End Select
 GetAutosaveFilename = Pathname & ReplaceForbiddenChars(Filename)
End Function

Public Function GetSubstFilename(Postscriptfile As String, TokenFilename As String, _
 Optional WithoutAuthor As Boolean = False, Optional Preview As Boolean = False) As String
 Dim PSHeader As tPSHeader, Author As String, _
  Title As String, UserName As String, Computername As String, i As Long, _
  DateTime As String, Filename As String, tStr As String, tList() As String, _
  Subst() As String, UserProfilPath As String, MyFiles As String

 DateTime = Format$(Now, "yyyymmddhhMMss")
 If Preview = False Then
   PSHeader = GetPSHeader(Postscriptfile)
   If Options.UseStandardAuthor = 1 Then
     Author = Options.StandardAuthor
    Else
     Author = PSHeader.CreateFor.Comment
   End If
 Else
   PSHeader.Title.Comment = "'Preview Title'"
   Author = "'Preview Author'"
 End If

 If Options.FilenameSubstitutionsOnlyInTitle = 1 Then
  tList = Split(Options.FilenameSubstitutions, "\")
  Title = PSHeader.Title.Comment
  If UBound(tList) >= 0 Then
   For i = 0 To UBound(tList)
    Subst = Split(tList(i), "|")
    If UBound(Subst) = 0 Then
      tStr = ""
     Else
      tStr = Subst(1)
    End If
    Title = Replace(Title, Subst(0), tStr, , , vbTextCompare)
   Next i
  End If
 End If

 UserName = GetUsername
 Computername = GetComputerName
 MyFiles = GetMyFiles

 Filename = TokenFilename
 Filename = Replace(Filename, "<DateTime>", DateTime, , , vbTextCompare)
 Filename = Replace(Filename, "<Computername>", Computername, , , vbTextCompare)
 Filename = Replace(Filename, "<Username>", UserName, , , vbTextCompare)
 Filename = Replace(Filename, "<Title>", Title, , , vbTextCompare)
 If WithoutAuthor = False Then
  Filename = Replace(Filename, "<Author>", Author, , , vbTextCompare)
 End If
 Filename = Replace(Filename, "<MyFiles>", MyFiles, , , vbTextCompare)

 If Options.FilenameSubstitutionsOnlyInTitle = 0 Then
  tList = Split(Options.FilenameSubstitutions, "\")
  If UBound(tList) >= 0 Then
   For i = 0 To UBound(tList)
    Subst = Split(tList(i), "|")
    If UBound(Subst) = 0 Then
      tStr = ""
     Else
      tStr = Subst(1)
    End If
    Filename = Replace(Filename, Subst(0), tStr, , , vbTextCompare)
   Next i
  End If
 End If
 If Options.RemoveSpaces = 1 Then
  Filename = Trim$(Filename)
 End If
 GetSubstFilename = Filename
End Function

Public Function CheckIfPSFile(ByVal Filename As String) As Boolean
 Dim lHeader As tPSHeader
 CheckIfPSFile = False
 lHeader = GetPSHeader(Filename)
 If InStr(1, lHeader.StartComment.Comment, "PS", vbTextCompare) > 0 Then
  CheckIfPSFile = True
 End If
End Function

Public Sub ReplacePSHeaderInFile(PSHeader As tPSHeader, Filename As String, _
 Optional BufferSize As Long = 65536)
 Dim bsize As Long, j As Long, fnSource As Long, fnDest As Long, fpos As Long, _
  tLen As Long, sBuffer As String, offset As Long, Tempfile As String, _
  TempHeaderFile As String

 Open Filename For Input As #fnSource
 TempHeaderFile = GetTempFile(CompletePath(GetPDFCreatorTempfolder), "~RP")
 Open Tempfile For Output As #fnDest

 ' All bytes befor
 bsize = BufferSize
 offset = PSHeader.StartComment.StartByte - 1
 For j = 1 To offset \ bsize
  fpos = (j - 1) * bsize + 1 + offset
  Seek #fnSource, fpos
  sBuffer = Input(bsize, fnSource)
  Put #fnDest, , sBuffer
  tLen = tLen + bsize
  DoEvents
 Next j
 If LOF(offset) > (j - 1) * bsize Then
  fpos = (j - 1) * bsize + 1 + offset
  Seek #fnSource, fpos
  sBuffer = Input(offset - (j - 1) * bsize, fnSource)
  Put #fnDest, , sBuffer
  tLen = tLen + (offset - (j - 1) * bsize)
 End If

 ' The PsHeader
 Put #fnDest, , GetPSHeaderString(PSHeader)

 ' All bytes after
 offset = PSHeader.EndComment.EndByte + 1
 For j = 1 To (LOF(fnSource) - offset + 1) \ bsize
  fpos = (j - 1) * bsize + 1 + offset
  Seek #fnSource, fpos
  sBuffer = Input(bsize, fnSource)
  Put #fnDest, , sBuffer
  tLen = tLen + bsize
  DoEvents
 Next j
 If (LOF(fnSource) - offset + 1) > (j - 1) * bsize Then
  fpos = (j - 1) * bsize + 1 + offset
  Seek #fnSource, fpos
  sBuffer = Input(LOF(fnSource) - (j - 1) * bsize, fnSource)
  Put #fnDest, , sBuffer
  tLen = tLen + (LOF(fnSource) - (j - 1) * bsize)
 End If
 Close #fnSource: Close #fnDest
End Sub

Private Sub CutFiles(InFile As String, OutFile As String, StartByte As Long, EndByte As Long)
 Dim fnSource As Long, fnDest As Long, buf As Long, tLen As String, i As Long, _
  sBuffer As String


 buf = 65536

 fnSource = FreeFile
 Open InFile For Binary As #fnSource
 fnDest = FreeFile
 Open OutFile For Output As #fnDest

 'Read until StartByte - 1
 tLen = EndByte - StartByte + 1


 For i = 1 To tLen \ buf
  Seek #fnSource, StartByte + (i - 1) * buf
  sBuffer = Input(buf, fnSource)
  Print #fnDest, sBuffer;
 Next i
 If tLen > (i - 1) * buf Then
  Seek #fnSource, StartByte + (i - 1) * buf
  sBuffer = Input(tLen - buf * (i - 1), fnSource)
  Print #fnDest, sBuffer;
 End If
 Close fnDest: Close fnSource
End Sub

Public Sub AppendPDFDocInfo(PSFile As String, PDFDocInfo As tPDFDocInfo)
 Dim fn As Long, DocInfoStr As String
 If LenB(Dir(PSFile)) > 0 Then
  DocInfoStr = Chr$(13) & "%!PS-Adobe-3.0 EPSF-3.0"
  DocInfoStr = DocInfoStr & Chr$(13) & "%%BoundingBox: 0 0 72 72"
  DocInfoStr = DocInfoStr & Chr$(13) & "%%EndProlog"
  DocInfoStr = DocInfoStr & Chr$(13) & "/pdfmark where {pop} {userdict /pdfmark /cleartomark load put} ifelse"
  DocInfoStr = DocInfoStr & Chr$(13) & "["
  With PDFDocInfo
   DocInfoStr = DocInfoStr & "/Author (" & .Author & ")"
   DocInfoStr = DocInfoStr & Chr$(13) & "/CreationDate (D:" & .CreationDate & ")"
   DocInfoStr = DocInfoStr & Chr$(13) & "/Creator (" & .Creator & ")"
   DocInfoStr = DocInfoStr & Chr$(13) & "/Keywords (" & .Keywords & ")"
   DocInfoStr = DocInfoStr & Chr$(13) & "/ModDate (D:" & .ModifyDate & ")"
   DocInfoStr = DocInfoStr & Chr$(13) & "/Subject (" & .Subject & ")"
   DocInfoStr = DocInfoStr & Chr$(13) & "/Title (" & .Title & ")"
  End With
  DocInfoStr = DocInfoStr & Chr$(13) & "/DOCINFO pdfmark"
  DocInfoStr = DocInfoStr & Chr$(13) & "%%EOF"
  fn = FreeFile
  Open PSFile For Append As fn
  Print #fn, DocInfoStr;
  Close #fn
 End If
End Sub
