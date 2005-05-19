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
 BoundingBox As tPSComment
 Orientation As tPSComment
 Title As tPSComment
 Creator As tPSComment
 CreationDate As tPSComment
 CreateFor As tPSComment
 DocumentNeededResources As tPSComment
 DocumentSuppliedResources As tPSComment
 DocumentData As tPSComment
 Pages As tPSComment
 PageOrder As tPSComment
 TargetDevice As tPSComment
 LanguageLevel As tPSComment
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

Public Function GetPSTitle(Filename As String) As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim tStr As String, psH As tPSHeader, pbg As PropertyBag
50050  Set pbg = ReadInfoSpoolfile(Filename)
50060  tStr = pbg.ReadProperty("REDMON_DOCNAME")
50070  If LenB(tStr) = 0 Then
50080   psH = GetPSHeader(Filename)
50090   tStr = psH.Title.Comment
50100  End If
50110  GetPSTitle = tStr
50120 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50130 Exit Function
ErrPtnr_OnError:
50151 Select Case ErrPtnr.OnError("modPDF", "GetPSTitle")
      Case 0: Resume
50170 Case 1: Resume Next
50180 Case 2: Exit Function
50190 Case 3: End
50200 End Select
50210 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetPSHeader(Filename As String) As tPSHeader
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim fn As Long, bufStr As String, PSHeader As tPSHeader, Buffer As Long
50050  If FileExists(Filename) And FileInUse(Filename) = False Then
50060   DoEvents
50070   fn = FreeFile
50080   If FileLen(Filename) = 0 Then
50090    Exit Function
50100   End If
50110   Buffer = 5000
50120   If FileLen(Filename) < Buffer Then
50130    Buffer = FileLen(Filename)
50140   End If
50150
50160   Open Filename For Binary Access Read As fn
50170   bufStr = Space$(Buffer)
50180   Get #fn, 1, bufStr
50190   Close #fn
50200
50210   With PSHeader
50220    .StartComment = GetPSComment(bufStr, "%!")
50230    .CreateFor = GetPSComment(bufStr, "%%For:")
50240    .CreationDate = GetPSComment(bufStr, "%%CreationDate:")
50250    .Creator = GetPSComment(bufStr, "%%Creator:")
50260    .Pages = GetPSComment(bufStr, "%%Pages:")
50270    .Title = GetPSComment(bufStr, "%%Title:")
50280    .EndComment = GetPSComment(bufStr, "%%EndComments")
50290   End With
50300   GetPSHeader = PSHeader
50310   DoEvents
50320  End If
50330 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50340 Exit Function
ErrPtnr_OnError:
50361 Select Case ErrPtnr.OnError("modPDF", "GetPSHeader")
      Case 0: Resume
50380 Case 1: Resume Next
50390 Case 2: Exit Function
50400 Case 3: End
50410 End Select
50420 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
50090    .EndByte = InStr(.StartByte, bufStr, vbLf)
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
50040  Dim PSHeader As tPSHeader, Author As String, ClientComputer As String, _
  Title As String, UserName As String, Computername As String, i As Long, _
  DateTime As String, Filename As String, tStr As String, tList() As String, _
  Subst() As String, UserProfilPath As String, MyFiles As String, _
  MyDesktop As String, Path As String, pbg As PropertyBag
50090
50100  Set pbg = InitPropertyBag
50110
50120  If Len(TokenFilename) = 0 Then
50130   Exit Function
50140  End If
50150
50160  DateTime = GetDocDate("", Options.StandardDateformat, CStr(Now))
50170  If Preview = False Then
50180    If FileExists(Postscriptfile) = True Then
50190     PSHeader = GetPSHeader(Postscriptfile)
50200     Set pbg = ReadInfoSpoolfile(Postscriptfile)
50210    End If
50220    If Options.UseStandardAuthor = 1 Then
50230      Author = Options.StandardAuthor
50240     Else
50250      Author = PSHeader.CreateFor.Comment
50260    End If
50270    If LenB(pbg.ReadProperty("REDMON_MACHINE")) > 0 Then
50280      ClientComputer = ReplaceForbiddenChars(pbg.ReadProperty("REDMON_MACHINE"), "")
50290     Else
50300      ClientComputer = ReplaceForbiddenChars(Environ$("REDMON_MACHINE"), "")
50310      If LenB(ClientComputer) = 0 Then
50320       ClientComputer = GetComputerName
50330      End If
50340    End If
50350   Else
50360    PSHeader.Title.Comment = "'Preview Title'"
50370    Author = "'Preview Author'"
50380    ClientComputer = "'Preview ClientComputer'"
50390  End If
50400
50410  If Options.FilenameSubstitutionsOnlyInTitle = 1 Then
50420   tList = Split(Options.FilenameSubstitutions, "\")
50430   Title = PSHeader.Title.Comment
50440   If UBound(tList) >= 0 Then
50450    For i = 0 To UBound(tList)
50460     Subst = Split(tList(i), "|")
50470     If UBound(Subst) = 0 Then
50480       tStr = ""
50490      Else
50500       tStr = Subst(1)
50510     End If
50520     Title = Replace(Title, Subst(0), tStr, , , vbTextCompare)
50530    Next i
50540   End If
50550  End If
50560
50570  UserName = GetDocUsername(Postscriptfile, Preview)
50580
50590  Computername = GetComputerName
50600  MyFiles = GetMyFiles
50610  MyDesktop = GetDesktop
50620
50630  Filename = TokenFilename
50640  Filename = Replace(Filename, "<DateTime>", DateTime, , , vbTextCompare)
50650  Filename = Replace(Filename, "<Computername>", Computername, , , vbTextCompare)
50660
50670  Filename = Replace(Filename, "<ClientComputer>", ClientComputer, , , vbTextCompare)
50680  Filename = Replace(Filename, "<Username>", UserName, , , vbTextCompare)
50690  Filename = Replace(Filename, "<Title>", Title, , , vbTextCompare)
50700  If WithoutAuthor = False Then
50710   Filename = Replace(Filename, "<Author>", Author, , , vbTextCompare)
50720  End If
50730  Filename = Replace(Filename, "<MyFiles>", MyFiles, , , vbTextCompare)
50740  Filename = Replace(Filename, "<MyDesktop>", MyDesktop, , , vbTextCompare)
50750
50760  tStr = "DOCNAME"
50770  If Preview = True Then
50780    Filename = Replace(Filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
50790   Else
50800    Filename = Replace(Filename, "<REDMON_" & tStr & ">", pbg.ReadProperty("REDMON_" & tStr), , , vbTextCompare)
50810  End If
50820  tStr = "JOB"
50830  If Preview = True Then
50840    Filename = Replace(Filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
50850   Else
50860    Filename = Replace(Filename, "<REDMON_" & tStr & ">", pbg.ReadProperty("REDMON_" & tStr), , , vbTextCompare)
50870  End If
50880  tStr = "MACHINE"
50890  If Preview = True Then
50900    Filename = Replace(Filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
50910   Else
50920    Filename = Replace(Filename, "<REDMON_" & tStr & ">", pbg.ReadProperty("REDMON_" & tStr), , , vbTextCompare)
50930  End If
50940  tStr = "PORT"
50950  If Preview = True Then
50960    Filename = Replace(Filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
50970   Else
50980    Filename = Replace(Filename, "<REDMON_" & tStr & ">", pbg.ReadProperty("REDMON_" & tStr), , , vbTextCompare)
50990  End If
51000  tStr = "PRINTER"
51010  If Preview = True Then
51020    Filename = Replace(Filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
51030   Else
51040    Filename = Replace(Filename, "<REDMON_" & tStr & ">", pbg.ReadProperty("REDMON_" & tStr), , , vbTextCompare)
51050  End If
51060  tStr = "SESSIONID"
51070  If Preview = True Then
51080    Filename = Replace(Filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
51090   Else
51100    Filename = Replace(Filename, "<REDMON_" & tStr & ">", pbg.ReadProperty("REDMON_" & tStr), , , vbTextCompare)
51110  End If
51120  tStr = "USER"
51130  If Preview = True Then
51140    Filename = Replace(Filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
51150   Else
51160    Filename = Replace(Filename, "<REDMON_" & tStr & ">", pbg.ReadProperty("REDMON_" & tStr), , , vbTextCompare)
51170  End If
51180
51190  If Options.FilenameSubstitutionsOnlyInTitle = 0 Then
51200   tList = Split(Options.FilenameSubstitutions, "\")
51210   If UBound(tList) >= 0 Then
51220    For i = 0 To UBound(tList)
51230     Subst = Split(tList(i), "|")
51240     If UBound(Subst) = 0 Then
51250       tStr = ""
51260      Else
51270       tStr = Subst(1)
51280     End If
51290     Filename = Replace(Filename, Subst(0), tStr, , , vbTextCompare)
51300    Next i
51310   End If
51320  End If
51330  If Options.RemoveSpaces = 1 Then
51340   Filename = Trim$(Filename)
51350  End If
51360  GetSubstFilename = Filename
51370 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
51380 Exit Function
ErrPtnr_OnError:
51401 Select Case ErrPtnr.OnError("modPDF", "GetSubstFilename")
      Case 0: Resume
51420 Case 1: Resume Next
51430 Case 2: Exit Function
51440 Case 3: End
51450 End Select
51460 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long, DocInfoStr As String, PDFDocInfoStr As String, tzi As clsTimeZoneInformation, tStr As String
50020  With PDFDocInfo
50030   If LenB(.Author) > 0 Then
50040    PDFDocInfoStr = PDFDocInfoStr & Chr$(13) & "/Author (" & EncodeChars(.Author) & ")"
50050   End If
50060   If LenB(.CreationDate) > 0 Or LenB(.ModifyDate) > 0 Then
50070    Set tzi = New clsTimeZoneInformation
50080    tStr = Format(TimeSerial(0, tzi.DaylightToGMT, 0), "hh'mm")
50090    If tzi.DaylightToGMT >= 0 Then
50100      tStr = "+" & tStr
50110     Else
50120      tStr = "-" & tStr
50130    End If
50140   End If
50150   If LenB(Trim$(.CreationDate)) > 0 Then
50160    PDFDocInfoStr = PDFDocInfoStr & Chr$(13) & "/CreationDate (D:" & EncodeChars(.CreationDate) & tStr & ")"
50170   End If
50180   If LenB(.Creator) > 0 Then
50190    PDFDocInfoStr = PDFDocInfoStr & Chr$(13) & "/Creator (" & EncodeChars(.Creator) & ")"
50200   End If
50210   If LenB(.Keywords) > 0 Then
50220    PDFDocInfoStr = PDFDocInfoStr & Chr$(13) & "/Keywords (" & EncodeChars(.Keywords) & ")"
50230   End If
50240   If LenB(Trim$(.ModifyDate)) > 0 Then
50250    PDFDocInfoStr = PDFDocInfoStr & Chr$(13) & "/ModDate (D:" & EncodeChars(.ModifyDate) & tStr & ")"
50260   End If
50270   If LenB(.CreationDate) > 0 Or LenB(.ModifyDate) > 0 Then
50280    Set tzi = Nothing
50290   End If
50300   If LenB(.Subject) > 0 Then
50310    PDFDocInfoStr = PDFDocInfoStr & Chr$(13) & "/Subject (" & EncodeChars(.Subject) & ")"
50320   End If
50330   If LenB(.Title) > 0 Then
50340    PDFDocInfoStr = PDFDocInfoStr & Chr$(13) & "/Title (" & EncodeChars(.Title) & ")"
50350   End If
50360  End With
50370  If FileExists(PSFile) = True And LenB(PDFDocInfoStr) > 0 Then
50380   DocInfoStr = Chr$(13) & "/pdfmark where {pop} {userdict /pdfmark /cleartomark load put} ifelse"
50390   DocInfoStr = DocInfoStr & Chr$(13) & "["
50400   DocInfoStr = DocInfoStr & Chr$(13) & PDFDocInfoStr
50410   DocInfoStr = DocInfoStr & Chr$(13) & "/DOCINFO pdfmark"
50420   DocInfoStr = DocInfoStr & Chr$(13) & "%%EOF"
50430   fn = FreeFile
50440   Open PSFile For Append As fn
50450   Print #fn, DocInfoStr;
50460   Close #fn
50470  End If
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
50040  Dim UserName As String, Path As String, i As Long, PSHeader As tPSHeader, _
  pbg As PropertyBag
50060  If NoFile = False Then
50070   If Len(Postscriptfile) > 0 Then
50080    If FileExists(Postscriptfile) = True Then
50090     Set pbg = ReadInfoSpoolfile(Postscriptfile)
50100     If LenB(pbg.ReadProperty("REDMON_USER")) > 0 Then
50110      UserName = pbg.ReadProperty("REDMON_USER")
50120     End If
50130     If LenB(UserName) = 0 Then
50140      PSHeader = GetPSHeader(Postscriptfile)
50150      If Len(PSHeader.CreateFor.Comment) > 0 Then
50160       UserName = PSHeader.CreateFor.Comment
50170      End If
50180     End If
50190    End If
50200   End If
50210  End If
50220  If LenB(UserName) = 0 Then
50230 ' If LenB(UserName) = 0 Or UCase$(UserName) = UCase$(App.EXEName) Then
50240   UserName = Environ$("Redmon_User")
50250  End If
50260  If Len(UserName) = 0 Then
50270   UserName = GetUsername
50280  End If
50290  GetDocUsername = UserName
50300 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50310 Exit Function
ErrPtnr_OnError:
50331 Select Case ErrPtnr.OnError("modPDF", "GetDocUsername")
      Case 0: Resume
50350 Case 1: Resume Next
50360 Case 2: Exit Function
50370 Case 3: End
50380 End Select
50390 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim Ext As String, Tempfile As String
50050  IFIsPS = False
50060  If LenB(InputFilename) > 0 Then
50070   If FileExists(InputFilename) = True Then
50080     If LenB(OutputFilename) > 0 Then
50090       If IsPostscriptFile(InputFilename) = True Then
50100        If GsDllLoaded = 0 Then
50110         Exit Sub
50120        End If
50130        SplitPath OutputFilename, , , , , Ext
50140        GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
50150        If GsDllLoaded = 0 Then
50160         MsgBox LanguageStrings.MessagesMsg08
50170        End If
50181        Select Case UCase$(Ext)
              Case "PDF"
50200          CallGScript InputFilename, OutputFilename, Options, PDFWriter
50210         Case "PNG"
50220          CallGScript InputFilename, OutputFilename, Options, PNGWriter
50230         Case "JPG"
50240          CallGScript InputFilename, OutputFilename, Options, JPEGWriter
50250         Case "BMP"
50260          CallGScript InputFilename, OutputFilename, Options, BMPWriter
50270         Case "PCX"
50280          CallGScript InputFilename, OutputFilename, Options, PCXWriter
50290         Case "TIF"
50300          CallGScript InputFilename, OutputFilename, Options, TIFFWriter
50310         Case "PS"
50320          CallGScript InputFilename, OutputFilename, Options, PSWriter
50330         Case "EPS"
50340          CallGScript InputFilename, OutputFilename, Options, EPSWriter
50350        End Select
50360       End If
50370       If GsDllLoaded <> 0 Then
50380        UnloadDLLComplete GsDllLoaded
50390       End If
50400       Exit Sub
50410      Else
50420       If IsPostscriptFile(InputFilename) = True Then
50430         IFIsPS = True
50440        Else
50450         MsgBox LanguageStrings.MessagesMsg06
50460       End If
50470       DoEvents
50480     End If
50490    Else
50500     If LenB(InputFilename) > 0 Then
50510      MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & vbCrLf & _
      "InputFile -IF" & vbCrLf & ">" & InputFilename & "<", vbExclamation + vbOKOnly
50530     End If
50540   End If
50550  End If
50560 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50570 Exit Sub
ErrPtnr_OnError:
50591 Select Case ErrPtnr.OnError("modPDF", "ConvertPostscriptFile")
      Case 0: Resume
50610 Case 1: Resume Next
50620 Case 2: Exit Sub
50630 Case 3: End
50640 End Select
50650 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
