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
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim psH As tPSHeader
50050  psH = GetPSHeader(Filename)
50060  GetPDFTitle = psH.Title.Comment
50070 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50080 Exit Function
ErrPtnr_OnError:
50101 Select Case ErrPtnr.OnError("modPDF", "GetPDFTitle")
      Case 0: Resume
50120 Case 1: Resume Next
50130 Case 2: Exit Function
50140 Case 3: End
50150 End Select
50160 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetPSHeader(Filename As String) As tPSHeader
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long, bufStr As String, PSHeader As tPSHeader, Buffer As Long
50020  If FileExists(Filename) And FileInUse(Filename) = False Then
50030   DoEvents
50040   fn = FreeFile
50050   If FileLen(Filename) = 0 Then
50060    Exit Function
50070   End If
50080   Buffer = 5000
50090   If FileLen(Filename) < Buffer Then
50100    Buffer = FileLen(Filename)
50110   End If
50120
50130   Open Filename For Binary Access Read As fn
50140   bufStr = Space$(Buffer)
50150   Get #fn, 1, bufStr
50160   Close #fn
50170
50180   With PSHeader
50190    .StartComment = GetPSComment(bufStr, "%!")
50200    .CreateFor = GetPSComment(bufStr, "%%For:")
50210    .CreationDate = GetPSComment(bufStr, "%%CreationDate:")
50220    .Creator = GetPSComment(bufStr, "%%Creator:")
50230    .Title = GetPSComment(bufStr, "%%Title:")
50240    .EndComment = GetPSComment(bufStr, "%%EndComments")
50250   End With
50260   GetPSHeader = PSHeader
50270   DoEvents
50280  End If
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

Private Function GetPSComment(ByRef bufStr As String, Comment As String) As tPSComment
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim PSComment As tPSComment
50050  PSComment.StartByte = -1
50060  If InStr(UCase$(bufStr), UCase$(Comment)) > 0 Then
50070   With PSComment
50080    .StartByte = InStr(bufStr, Comment)
50090    .EndByte = InStr(.StartByte, bufStr, Chr$(&HA))
50100    If .EndByte - (.StartByte + Len(Comment)) > 0 Then
50110     .Comment = Mid$(bufStr, .StartByte + Len(Comment), .EndByte - (.StartByte + Len(Comment)))
50120     .Comment = Replace(Replace(.Comment, Chr$(&HA), ""), Chr$(&HD), "")
50130    End If
50140    .Comment = Trim$(.Comment)
50150    If Len(.Comment) > 0 Then
50160     If Mid$(.Comment, 1, 1) = "(" Then
50170      .Comment = Mid(.Comment, 2)
50180     End If
50190     If Len(.Comment) > 0 Then
50200      If Mid$(.Comment, Len(.Comment), 1) = ")" Then
50210       .Comment = Mid(.Comment, 1, Len(.Comment) - 1)
50220      End If
50230     End If
50240    End If
50250    .Comment = ReplaceEncodingChars(.Comment)
50260   End With
50270  End If
50280  GetPSComment = PSComment
50290 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50300 Exit Function
ErrPtnr_OnError:
50321 Select Case ErrPtnr.OnError("modPDF", "GetPSComment")
      Case 0: Resume
50340 Case 1: Resume Next
50350 Case 2: Exit Function
50360 Case 3: End
50370 End Select
50380 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function GetPSHeaderString(PSHeader As tPSHeader) As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim tStr As String
50050  If PSHeader.StartComment.StartByte = -1 Then
50060    tStr = "%!PS-Adobe-3.0" & Chr$(&HA)
50070   Else
50080    tStr = "%!" & PSHeader.StartComment.Comment & Chr$(&HA)
50090  End If
50100
50110  tStr = tStr & "%%For:" & PSHeader.CreateFor.Comment & Chr$(&HA)
50120  tStr = tStr & "%%CreationDate:" & PSHeader.CreationDate.Comment & Chr$(&HA)
50130  tStr = tStr & "%%Creator:" & PSHeader.Creator.Comment & Chr$(&HA)
50140  tStr = tStr & "%%Title:" & PSHeader.Title.Comment & Chr$(&HA)
50150
50160  tStr = tStr & "%%EndComments" & Chr$(&HA)
50170  GetPSHeaderString = tStr
50180 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50190 Exit Function
ErrPtnr_OnError:
50211 Select Case ErrPtnr.OnError("modPDF", "GetPSHeaderString")
      Case 0: Resume
50230 Case 1: Resume Next
50240 Case 2: Exit Function
50250 Case 3: End
50260 End Select
50270 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetAutosaveFilename(Postscriptfile As String) As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim Filename As String, Pathname As String
50050
50060  If Options.UseAutosaveDirectory = 1 Then
50070    Pathname = CompletePath(Trim$(Options.AutosaveDirectory))
50080   Else
50090    Pathname = CompletePath(Trim$(Options.LastSaveDirectory))
50100  End If
50110
50120  Filename = GetSubstFilename(Postscriptfile, Options.AutosaveFilename)
50130
50140  If Len(Filename) > 0 And Options.RemoveAllKnownFileExtensions = 1 Then
50150   Filename = RemoveAllKnownFileExtensions(Filename)
50160  End If
50170
50181  Select Case Options.AutosaveFormat
              Case 0: 'PDF
50200    Filename = Filename & ".pdf"
50210   Case 1: 'PNG
50220    Filename = Filename & ".png"
50230   Case 2: 'JPEG
50240    Filename = Filename & ".jpg"
50250   Case 3: 'BMP
50260    Filename = Filename & ".bmp"
50270   Case 4: 'PCX
50280    Filename = Filename & ".pcx"
50290   Case 5: 'TIFF
50300    Filename = Filename & ".tif"
50310   Case 6: 'PS
50320    Filename = Filename & ".ps"
50330   Case 7: 'EPS
50340    Filename = Filename & ".eps"
50350  End Select
50360  GetAutosaveFilename = CompletePath(GetSubstFilename(Postscriptfile, Pathname)) & ReplaceForbiddenChars(Filename)
50370 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50380 Exit Function
ErrPtnr_OnError:
50401 Select Case ErrPtnr.OnError("modPDF", "GetAutosaveFilename")
      Case 0: Resume
50420 Case 1: Resume Next
50430 Case 2: Exit Function
50440 Case 3: End
50450 End Select
50460 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetSubstFilename(Postscriptfile As String, TokenFilename As String, _
 Optional WithoutAuthor As Boolean = False, Optional Preview As Boolean = False) As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim PSHeader As tPSHeader, Author As String, _
  Title As String, UserName As String, Computername As String, i As Long, _
  DateTime As String, Filename As String, tStr As String, tList() As String, _
  Subst() As String, UserProfilPath As String, MyFiles As String, Path As String
50080
50090  If Len(TokenFilename) = 0 Then
50100   Exit Function
50110  End If
50120
50130  DateTime = GetDocDate("", Options.StandardDateformat, CStr(Now))
50140  If Preview = False Then
50150    If FileExists(Postscriptfile) = True Then
50160     PSHeader = GetPSHeader(Postscriptfile)
50170    End If
50180    If Options.UseStandardAuthor = 1 Then
50190      Author = Options.StandardAuthor
50200     Else
50210      Author = PSHeader.CreateFor.Comment
50220    End If
50230   Else
50240    PSHeader.Title.Comment = "'Preview Title'"
50250    Author = "'Preview Author'"
50260  End If
50270
50280  If Options.FilenameSubstitutionsOnlyInTitle = 1 Then
50290   tList = Split(Options.FilenameSubstitutions, "\")
50300   Title = PSHeader.Title.Comment
50310   If UBound(tList) >= 0 Then
50320    For i = 0 To UBound(tList)
50330     Subst = Split(tList(i), "|")
50340     If UBound(Subst) = 0 Then
50350       tStr = ""
50360      Else
50370       tStr = Subst(1)
50380     End If
50390     Title = Replace(Title, Subst(0), tStr, , , vbTextCompare)
50400    Next i
50410   End If
50420  End If
50430
50440  UserName = GetDocUsername(Postscriptfile, Preview)
50450
50460  Computername = GetComputerName
50470  MyFiles = GetMyFiles
50480
50490  Filename = TokenFilename
50500  Filename = Replace(Filename, "<DateTime>", DateTime, , , vbTextCompare)
50510  Filename = Replace(Filename, "<Computername>", Computername, , , vbTextCompare)
50520  Filename = Replace(Filename, "<Username>", UserName, , , vbTextCompare)
50530  Filename = Replace(Filename, "<Title>", Title, , , vbTextCompare)
50540  If WithoutAuthor = False Then
50550   Filename = Replace(Filename, "<Author>", Author, , , vbTextCompare)
50560  End If
50570  Filename = Replace(Filename, "<MyFiles>", MyFiles, , , vbTextCompare)
50580
50590  If Options.FilenameSubstitutionsOnlyInTitle = 0 Then
50600   tList = Split(Options.FilenameSubstitutions, "\")
50610   If UBound(tList) >= 0 Then
50620    For i = 0 To UBound(tList)
50630     Subst = Split(tList(i), "|")
50640     If UBound(Subst) = 0 Then
50650       tStr = ""
50660      Else
50670       tStr = Subst(1)
50680     End If
50690     Filename = Replace(Filename, Subst(0), tStr, , , vbTextCompare)
50700    Next i
50710   End If
50720  End If
50730  If Options.RemoveSpaces = 1 Then
50740   Filename = Trim$(Filename)
50750  End If
50760  GetSubstFilename = Filename
50770 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50780 Exit Function
ErrPtnr_OnError:
50801 Select Case ErrPtnr.OnError("modPDF", "GetSubstFilename")
      Case 0: Resume
50820 Case 1: Resume Next
50830 Case 2: Exit Function
50840 Case 3: End
50850 End Select
50860 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsPostscriptFile(ByVal Filename As String) As Boolean
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim lHeader As tPSHeader
50050  IsPostscriptFile = False
50060  lHeader = GetPSHeader(Filename)
50070  If InStr(1, lHeader.StartComment.Comment, "PS", vbTextCompare) > 0 Then
50080   IsPostscriptFile = True
50090  End If
50100 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50110 Exit Function
ErrPtnr_OnError:
50131 Select Case ErrPtnr.OnError("modPDF", "IsPostscriptFile")
      Case 0: Resume
50150 Case 1: Resume Next
50160 Case 2: Exit Function
50170 Case 3: End
50180 End Select
50190 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub AppendPDFDocInfo(PSFile As String, PDFDocInfo As tPDFDocInfo)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim fn As Long, DocInfoStr As String
50050  If FileExists(PSFile) = True Then
50060   DocInfoStr = Chr$(13) & "%!PS-Adobe-3.0 EPSF-3.0"
50070   DocInfoStr = DocInfoStr & Chr$(13) & "%%BoundingBox: 0 0 72 72"
50080   DocInfoStr = DocInfoStr & Chr$(13) & "%%EndProlog"
50090   DocInfoStr = DocInfoStr & Chr$(13) & "/pdfmark where {pop} {userdict /pdfmark /cleartomark load put} ifelse"
50100   DocInfoStr = DocInfoStr & Chr$(13) & "["
50110   With PDFDocInfo
50120    DocInfoStr = DocInfoStr & "/Author (" & EncodeChars(.Author) & ")"
50130    If LenB(Trim$(.CreationDate)) = 0 Then
50140      DocInfoStr = DocInfoStr & Chr$(13) & "/CreationDate ()"
50150     Else
50160      DocInfoStr = DocInfoStr & Chr$(13) & "/CreationDate (" & EncodeChars(.CreationDate) & ")"
50170    End If
50180    DocInfoStr = DocInfoStr & Chr$(13) & "/Creator (" & EncodeChars(.Creator) & ")"
50190    DocInfoStr = DocInfoStr & Chr$(13) & "/Keywords (" & EncodeChars(.Keywords) & ")"
50200    If LenB(Trim$(.ModifyDate)) = 0 Then
50210      DocInfoStr = DocInfoStr & Chr$(13) & "/ModDate ()"
50220     Else
50230      DocInfoStr = DocInfoStr & Chr$(13) & "/ModDate (" & EncodeChars(.ModifyDate) & ")"
50240    End If
50250    DocInfoStr = DocInfoStr & Chr$(13) & "/Subject (" & EncodeChars(.Subject) & ")"
50260    DocInfoStr = DocInfoStr & Chr$(13) & "/Title (" & EncodeChars(.Title) & ")"
50270 '   DocInfoStr = DocInfoStr & Chr$(13) & "/Producer ()"
50280   End With
50290   DocInfoStr = DocInfoStr & Chr$(13) & "/DOCINFO pdfmark"
50300   DocInfoStr = DocInfoStr & Chr$(13) & "%%EOF"
50310   fn = FreeFile
50320   Open PSFile For Append As fn
50330   Print #fn, DocInfoStr;
50340   Close #fn
50350  End If
50360 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50370 Exit Sub
ErrPtnr_OnError:
50391 Select Case ErrPtnr.OnError("modPDF", "AppendPDFDocInfo")
      Case 0: Resume
50410 Case 1: Resume Next
50420 Case 2: Exit Sub
50430 Case 3: End
50440 End Select
50450 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function ReplaceEncodingChars(Str1 As String) As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Integer, tStr As String
50050  tStr = ""
50060  ' First we look for oct encoding chars
50070  For i = 127 To 255
50080   Str1 = Replace$(Str1, "\" & Oct$(i), Chr$(i))
50090  Next i
50100  ReplaceEncodingChars = Str1
50110  ' Second we look for hex encoding chars
50120  If Len(Str1) >= 4 Then
50130   If Mid$(Str1, 1, 1) = "<" And Mid$(Str1, Len(Str1), 1) = ">" Then
50140    If Len(Str1) Mod 2 = 0 Then
50150     For i = 2 To Len(Str1) - 1 Step 2
50160      If IsNumeric("&H" & Mid$(Str1, i, 2)) = True Then
50170        If CByte("&H" & Mid$(Str1, i, 2)) > 255 Then
50180          Exit Function
50190         Else
50200          tStr = tStr & Chr$(CByte("&H" & Mid$(Str1, i, 2)))
50210        End If
50220       Else
50230        Exit Function
50240      End If
50250     Next i
50260    End If
50270   End If
50280  End If
50290  If Len(tStr) > 0 Then
50300   ReplaceEncodingChars = tStr
50310  End If
50320 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50330 Exit Function
ErrPtnr_OnError:
50351 Select Case ErrPtnr.OnError("modPDF", "ReplaceEncodingChars")
      Case 0: Resume
50370 Case 1: Resume Next
50380 Case 2: Exit Function
50390 Case 3: End
50400 End Select
50410 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetDocUsername(Postscriptfile As String, NoFile As Boolean) As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim UserName As String, Path As String, i As Long, PSHeader As tPSHeader
50050  If NoFile = False Then
50060   If Len(Postscriptfile) > 0 Then
50070    If FileExists(Postscriptfile) = True Then
50080     PSHeader = GetPSHeader(Postscriptfile)
50090     If Len(PSHeader.CreateFor.Comment) > 0 Then
50100      UserName = PSHeader.CreateFor.Comment
50110     End If
50120     SplitPath Postscriptfile, , Path
50130     Path = CompletePath(Path)
50140     If Len(Path) > 1 Then
50150      Path = Left(Path, Len(Path) - 1)
50160      i = InStrRev(Path, "\", , vbTextCompare)
50170      If i > 0 Then
50180       If Len(UserName) = 0 Then
50190        UserName = Mid$(Path, i + 1)
50200       End If
50210      End If
50220     End If
50230    End If
50240   End If
50250  End If
50260  If Len(UserName) = 0 Or UCase$(UserName) = UCase$(App.EXEName) Then
50270   UserName = Environ$("Redmon_User")
50280  End If
50290 ' If Len(UserName) = 0 Then
50300 '  UserName = PSHeader.CreateFor.Comment
50310 ' End If
50320  If Len(UserName) = 0 Then
50330   UserName = GetUsername
50340  End If
50350  GetDocUsername = UserName
50360 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50370 Exit Function
ErrPtnr_OnError:
50391 Select Case ErrPtnr.OnError("modPDF", "GetDocUsername")
      Case 0: Resume
50410 Case 1: Resume Next
50420 Case 2: Exit Function
50430 Case 3: End
50440 End Select
50450 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetDocDate(Optional StandardDate As String = "", Optional StandardDateformat As String = "", Optional UseThisdate As String = "") As String
 On Error Resume Next
 Dim tStr As String, DateFormat As String, Usingdate As String

 If LenB(Trim$(StandardDate)) = 0 Then ' No standard date
   Usingdate = UseThisdate
  Else
   If LenB(RemoveLeadingAndTrailingQuotes(Trim$(StandardDate))) = 0 Then 'Empty date
     Usingdate = ""
    Else
     Usingdate = StandardDate
   End If
 End If

 If Len(StandardDateformat) > 0 Then
   DateFormat = StandardDateformat
  Else
   DateFormat = "YYYYMMDDHHNNSS"
 End If

 tStr = Format$(Usingdate, DateFormat)
 If LenB(tStr) = 0 Then
  tStr = Usingdate
 End If
 GetDocDate = tStr
End Function

Public Function FormatPrintDocumentDate(tDate As String) As String
 Dim tStr As Long, m As Long, d As Long, Y As Long
 On Error Resume Next
 FormatPrintDocumentDate = tDate
 If InStr(tDate, "/") > 0 Then
   tStr = Mid(tDate, 1, InStr(tDate, "/") - 1)
   If IsNumeric(tStr) = False Then
    Exit Function
   End If
   m = CLng(tStr)
   If m < 1 Or m > 12 Then
    Exit Function
   End If
  Else
   Exit Function
 End If
 If InStr(InStr(tDate, "/") + 1, tDate, "/") > 0 Then
   tStr = Mid(tDate, InStr(tDate, "/") + 1, InStr(InStr(tDate, "/") + 1, tDate, "/") - (InStr(tDate, "/") + 1))
   If IsNumeric(tStr) = False Then
    Exit Function
   End If
   d = CLng(tStr)
   If d < 1 Or d > 31 Then
    Exit Function
   End If
  Else
   Exit Function
 End If
 If InStr(tDate, " ") > 0 Then
   tStr = Mid(tDate, InStr(InStr(tDate, "/") + 1, tDate, "/") + 1, InStr(tDate, " ") - (InStr(InStr(tDate, "/") + 1, tDate, "/") + 1))
   If IsNumeric(tStr) = False Then
    Exit Function
   End If
   Y = CLng(tStr)
  Else
   Exit Function
 End If
 FormatPrintDocumentDate = CStr(DateSerial(Y, m, d)) + Mid(tDate, InStr(tDate, " "))
End Function

Public Function EncodeChars(ByVal Str1 As String) As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long, tStr As String
50050  Str1 = Replace(Str1, "\", "\\")
50060  Str1 = Replace(Str1, "{", "\{")
50070  Str1 = Replace(Str1, "}", "\}")
50080  Str1 = Replace(Str1, "[", "\[")
50090  Str1 = Replace(Str1, "]", "\]")
50100  Str1 = Replace(Str1, "(", "\(")
50110  Str1 = Replace(Str1, ")", "\)")
50120  For i = 1 To Len(Str1)
50130   If Asc(Mid(Str1, i, 1)) > 127 Then
50140     tStr = tStr & "\" & CStr(Oct(Asc(Mid(Str1, i, 1))))
50150    Else
50160     tStr = tStr & Mid(Str1, i, 1)
50170   End If
50180  Next i
50190  EncodeChars = tStr
50200 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50210 Exit Function
ErrPtnr_OnError:
50231 Select Case ErrPtnr.OnError("modPDF", "EncodeChars")
      Case 0: Resume
50250 Case 1: Resume Next
50260 Case 2: Exit Function
50270 Case 3: End
50280 End Select
50290 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub ConvertPostscriptFile(InputFilename As String, OutputFilename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Ext As String, Tempfile As String
50020  IFIsPS = False
50030  If LenB(InputFilename) > 0 Then
50040   If FileExists(InputFilename) = True Then
50050     If LenB(OutputFilename) > 0 Then
50060       If IsPostscriptFile(InputFilename) = True Then
50070        If GsDllLoaded = 0 Then
50080         Exit Sub
50090        End If
50100        SplitPath OutputFilename, , , , , Ext
50110        GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
50120        If GsDllLoaded = 0 Then
50130         MsgBox LanguageStrings.MessagesMsg08
50140        End If
50151        Select Case UCase$(Ext)
              Case "PDF"
50170          CallGScript InputFilename, OutputFilename, Options, PDFWriter
50180         Case "PNG"
50190          CallGScript InputFilename, OutputFilename, Options, PNGWriter
50200         Case "JPG"
50210          CallGScript InputFilename, OutputFilename, Options, JPEGWriter
50220         Case "BMP"
50230          CallGScript InputFilename, OutputFilename, Options, BMPWriter
50240         Case "PCX"
50250          CallGScript InputFilename, OutputFilename, Options, PCXWriter
50260         Case "TIF"
50270          CallGScript InputFilename, OutputFilename, Options, TIFFWriter
50280         Case "PS"
50290          CallGScript InputFilename, OutputFilename, Options, PSWriter
50300         Case "EPS"
50310          CallGScript InputFilename, OutputFilename, Options, EPSWriter
50320        End Select
50330       End If
50340       If GsDllLoaded <> 0 Then
50350        UnloadDLLComplete GsDllLoaded
50360       End If
50370       Exit Sub
50380      Else
50390       If IsPostscriptFile(InputFilename) = True Then
50400         IfLoggingWriteLogfile "Get inputfile: " & InputFilename
50410         WriteToSpecialLogfile "Get inputfile: " & InputFilename
50420         If DirExists(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory & "\" & Environ$("REDMON_USER")) = False Then
50430          MakePath CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory & "\" & Environ$("REDMON_USER")
50440         End If
50450         If IsWin9xMe = True Then
50460           Tempfile = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory & "\", "~PD")
50470          Else
50480           Tempfile = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory & "\" & Environ$("REDMON_USER"), "~PD")
50490         End If
50500         KillFile Tempfile
50510         If FileExists(Tempfile) = False Then
50520          'MsgBox ">" & CommandSwitch("IF", True) & vbCrLf & ">" & Tempfile
50530          If PDFCreatorPrinter Then
50540            IfLoggingWriteLogfile "Move the inputfile to " & Tempfile
50550            WriteToSpecialLogfile "Move the inputfile to " & Tempfile
50560            Name CommandSwitch("IF", True) As Tempfile
50570           Else
50580            IfLoggingWriteLogfile "Copy the inputfile to " & Tempfile
50590            WriteToSpecialLogfile "Copy the inputfile to " & Tempfile
50600            FileCopy InputFilename, Tempfile
50610          End If
50620         End If
50630         IFIsPS = True
50640        Else
50650         MsgBox LanguageStrings.MessagesMsg06
50660       End If
50670       DoEvents
50680     End If
50690    Else
50700     If LenB(InputFilename) > 0 Then
50710      MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & vbCrLf & _
      "InputFile -IF" & vbCrLf & ">" & InputFilename & "<", vbExclamation + vbOKOnly
50730     End If
50740   End If
50750  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "ConvertPostscriptFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

