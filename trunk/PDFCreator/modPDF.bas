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
 Const Buffer = 5000
 Dim fn As Long, bufStr As String, pTitleStart As String, pTitleEnd As String

 If Len(Dir(FileName)) = 0 Then
  GetPDFTitle = vbNullString
  Exit Function
 End If

 fn = FreeFile
 Open FileName For Input As #fn
 If FileLen(FileName) < Buffer Then
   bufStr = Input(FileLen(FileName), #fn)
  Else
   bufStr = Input(Buffer, #fn)
 End If
 Close #fn

 pTitleStart = InStr(1, bufStr, "%%Title:", vbTextCompare)
 pTitleEnd = InStr(pTitleStart, bufStr, vbLf, vbTextCompare)

 If Mid$(bufStr, pTitleEnd - 1, 1) = vbCr Then
  pTitleEnd = pTitleEnd - 1
 End If

 If pTitleStart = 0 Then
  GetPDFTitle = vbNullString
  Exit Function
 End If
 GetPDFTitle = Trim$(Mid$(bufStr, pTitleStart + 8, pTitleEnd - pTitleStart - 8))
End Function

Public Sub SavePDFTitle(FileName As String, Title As String)
 Const Buffer = 5000
 Dim fn As Long, bufStr As String, pTitleStart As String, pTitleEnd As String, _
  oldTitle As String

 If Dir(FileName) = "" Then
  Exit Sub
 End If

 fn = FreeFile
 Open FileName For Input As #fn
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
 Open FileName For Output As #fn
  Print #fn, bufStr
 Close #fn
End Sub

Public Function GetPSHeader(FileName As String) As tPSHeader
 Dim fn As Long, bufStr As String, PSHeader As tPSHeader, Buffer As Long
 DoEvents
 fn = FreeFile
 If FileLen(FileName) = 0 Then
  Exit Function
 End If
 Buffer = 5000
 If FileLen(FileName) < Buffer Then
  Buffer = FileLen(FileName)
 End If

 Open FileName For Binary Access Read As fn
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

Public Sub PutPSHeader(FileName As String, PSHeader As tPSHeader)
 Dim fn As Long, bufStr As String
 fn = FreeFile
 Open FileName For Binary Access Read As fn
 bufStr = Space$(LOF(fn))
 Get #fn, 1, bufStr
 Close #fn
 If PSHeader.StartComment.StartByte = -1 And PSHeader.EndComment.StartByte = -1 Then
  bufStr = GetPSHeaderString(PSHeader) & Chr(&HA) & bufStr
  fn = FreeFile
  Open FileName For Output As #fn
  Print #fn, bufStr
  Close #fn
  Exit Sub
 End If
 If PSHeader.StartComment.StartByte >= 0 And PSHeader.EndComment.StartByte >= 0 Then
  bufStr = Mid$(bufStr, 1, PSHeader.StartComment.StartByte - 1) & _
   GetPSHeaderString(PSHeader) & Mid$(bufStr, PSHeader.EndComment.EndByte + 1)
  fn = FreeFile
  Open FileName For Output As #fn
  Print #fn, bufStr
  Close #fn
  Exit Sub
 End If
 If PSHeader.StartComment.StartByte >= 0 And PSHeader.EndComment.StartByte = -1 Then
  bufStr = GetPSHeaderString(PSHeader) & Mid$(bufStr, PSHeader.StartComment.EndByte + 1)
  fn = FreeFile
  Open FileName For Output As #fn
  Print #fn, bufStr
  Close #fn
  Exit Sub
 End If
 If PSHeader.StartComment.StartByte = -1 And PSHeader.EndComment.StartByte >= 0 Then
  bufStr = GetPSHeaderString(PSHeader) & bufStr
  fn = FreeFile
  Open FileName For Output As #fn
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
 Dim FileName As String, Pathname As String

 If Options.UseAutosaveDirectory = 1 Then
   Pathname = Trim$(Options.AutosaveDirectory)
  Else
   Pathname = Trim$(Options.LastSaveDirectory)
 End If
 If Right$(Pathname, 1) <> "\" Then
  Pathname = Pathname & "\"
 End If

 FileName = GetSubstFilename(Postscriptfile, Options.AutosaveFilename)

 Select Case Options.AutosaveFormat
  Case 0: 'PDF
   FileName = FileName & ".pdf"
  Case 1: 'PNG
   FileName = FileName & ".png"
  Case 2: 'JPEG
   FileName = FileName & ".jpg"
  Case 3: 'BMP
   FileName = FileName & ".bmp"
  Case 4: 'PCX
   FileName = FileName & ".pcx"
  Case 5: 'TIFF
   FileName = FileName & ".tif"
  Case 6: 'PS
   FileName = FileName & ".ps"
  Case 7: 'EPS
   FileName = FileName & ".eps"
 End Select
 GetAutosaveFilename = Pathname & ReplaceForbiddenChars(FileName)
End Function

Public Function GetSubstFilename(Postscriptfile As String, TokenFilename As String) As String
 Dim PSHeader As tPSHeader, Author As String, _
  Title As String, Username As String, Computername As String, i As Long, _
  DateTime As String, FileName As String, tStr As String, tList() As String, _
  Subst() As String

 DateTime = Format$(Now, "yyyymmddhhMMss")
 PSHeader = GetPSHeader(Postscriptfile)
 If Options.UseStandardAuthor = 1 Then
   Author = Options.StandardAuthor
  Else
   Author = PSHeader.CreateFor.Comment
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

 Username = GetUsername
 Computername = GetComputerName

 FileName = TokenFilename
 FileName = Replace(FileName, "<DateTime>", DateTime, , , vbTextCompare)
 FileName = Replace(FileName, "<Computername>", Computername, , , vbTextCompare)
 FileName = Replace(FileName, "<Username>", Username, , , vbTextCompare)
 FileName = Replace(FileName, "<Title>", Title, , , vbTextCompare)
 FileName = Replace(FileName, "<Author>", Author, , , vbTextCompare)

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
    FileName = Replace(FileName, Subst(0), tStr, , , vbTextCompare)
   Next i
  End If
 End If
 If Options.RemoveSpaces = 1 Then
  FileName = Trim$(FileName)
 End If
 GetSubstFilename = FileName
End Function

Public Function CheckIfPSFile(ByVal FileName As String) As Boolean
 Dim lHeader As tPSHeader
 CheckIfPSFile = False
 lHeader = GetPSHeader(FileName)
 If InStr(1, lHeader.StartComment.Comment, "PS", vbTextCompare) > 0 Then
  CheckIfPSFile = True
 End If
End Function
