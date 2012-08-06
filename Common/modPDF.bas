Attribute VB_Name = "modPDF"
Option Explicit

'This variable is used in frmMain and frmPrinting
Public PDFSpoolfile As String
Public CurrentInfoSpoolFile As String

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

Public Enum eCodePage
 CP_NoEncoding = 0
 CP_UTF8 = 65001
 CP_UTF16 = 65002
End Enum

Public Function GetPSTitleFromPSString(PSString As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim psH As tPSHeader
50020  If LenB(PSString) > 0 Then
50030   psH = GetPSHeader(PSString, True)
50040   GetPSTitleFromPSString = psH.Title.Comment
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "GetPSTitleFromPSString")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetPSTitle(PostScriptFilename As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim psH As tPSHeader
50020  psH = GetPSHeader(PostScriptFilename)
50030  GetPSTitle = psH.Title.Comment
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "GetPSTitle")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetPSHeader(filename As String, Optional FileNameIsPSString = False) As tPSHeader
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long, bufStr As String, PSHeader As tPSHeader, buffer As Long, FileSize As Currency
50020
50030   If FileNameIsPSString = False Then
50040     If FileExists(filename) Then
50050      If FileInUse(filename) = False Then
50060       DoEvents
50070       fn = FreeFile
50080       FileSize = GetFileLength(filename)
50090       If FileSize = 0 Then
50100        Exit Function
50110       End If
50120       buffer = 5000
50130       If FileSize > 0 And FileSize < buffer Then
50140        buffer = FileSize
50150       End If
50160
50170       Open filename For Binary Access Read As fn
50180       bufStr = Space$(buffer)
50190       Get #fn, 1, bufStr
50200       Close #fn
50210       DoEvents
50220      End If
50230     End If
50240    Else
50250     bufStr = filename
50260   End If
50270
50280  With PSHeader
50290   .StartComment = GetPSComment(bufStr, "%!")
50300   .CreateFor = GetPSComment(bufStr, "%%For:")
50310   .CreationDate = GetPSComment(bufStr, "%%CreationDate:")
50320   .Creator = GetPSComment(bufStr, "%%Creator:")
50330   .Pages = GetPSComment(bufStr, "%%Pages:")
50340   .Title = GetPSComment(bufStr, "%%Title:")
50350   .EndComment = GetPSComment(bufStr, "%%EndComments")
50360  End With
50370  GetPSHeader = PSHeader
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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim PSComment As tPSComment, tStr As String
50020  PSComment.StartByte = -1
50030  If InStr(UCase$(bufStr), UCase$(Comment)) > 0 Then
50040   With PSComment
50050    .StartByte = InStr(bufStr, Comment)
50060    .EndByte = InStr(.StartByte, bufStr, vbLf)
50070    If .EndByte - (.StartByte + Len(Comment)) > 0 Then
50080     .Comment = Mid$(bufStr, .StartByte + Len(Comment), .EndByte - (.StartByte + Len(Comment)))
50090     .Comment = Replace(Replace(.Comment, Chr$(&HA), ""), Chr$(&HD), "")
50100    End If
50110    .Comment = Trim$(.Comment)
50120    If LenB(.Comment) > 0 Then
50130      tStr = ReplaceEncodingChars(.Comment)
50140      If tStr <> .Comment Then
50150        .Comment = RemoveEncodedPostscriptChars(RemoveLeadingAndTrailingBrackets(tStr))
50160       Else
50170        .Comment = tStr
50180      End If
50190    End If
50200   End With
50210  End If
50220  GetPSComment = PSComment
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

Public Function GetAutosaveFilename(InfoSpoolFileName As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim filename As String, Pathname As String
50020
50030  If Options.UseAutosaveDirectory = 1 Then
50040    Pathname = GetSubstFilename2(Options.AutosaveDirectory, False, , InfoSpoolFileName, False)
50050   Else
50060    Pathname = GetSubstFilename2(Options.LastSaveDirectory, False)
50070  End If
50080
50090  filename = GetSubstFilename(InfoSpoolFileName, Options.AutosaveFilename)
50100
50110  If Len(filename) > 0 And Options.RemoveAllKnownFileExtensions = 1 Then
50120   filename = RemoveAllKnownFileExtensions(filename)
50130  End If
50140
50150 '  Case 0: 'PDF
50160 '  Case 1: 'PNG
50170 '  Case 2: 'JPEG
50180 '  Case 3: 'BMP
50190 '  Case 4: 'PCX
50200 '  Case 5: 'TIFF
50210 '  Case 6: 'PS
50220 '  Case 7: 'EPS
50230 '  Case 8: 'TXT
50240 '  Case 9: 'PDFA
50250 '  Case 10: 'PDFX
50260 '  Case 11: 'PSD
50270 '  Case 12: 'PCL
50280 '  Case 13: 'RAW
50291  Select Case Options.AutosaveFormat
        Case 0, 9, 10: 'PDF
50310    filename = filename & ".pdf"
50320   Case 1: 'PNG
50330    filename = filename & ".png"
50340   Case 2: 'JPEG
50350    filename = filename & ".jpg"
50360   Case 3: 'BMP
50370    filename = filename & ".bmp"
50380   Case 4: 'PCX
50390    filename = filename & ".pcx"
50400   Case 5: 'TIFF
50410    filename = filename & ".tif"
50420   Case 6: 'PS
50430    filename = filename & ".ps"
50440   Case 7: 'EPS
50450    filename = filename & ".eps"
50460   Case 8: 'TXT
50470    filename = filename & ".txt"
50480   Case 11: 'PSD
50490    filename = filename & ".psd"
50500   Case 12: 'PCL
50510    filename = filename & ".pcl"
50520   Case 13: 'RAW
50530    filename = filename & ".raw"
50540   Case 14: 'SVG
50550    filename = filename & ".svg"
50560  End Select
50570  GetAutosaveFilename = CompletePath(Pathname) & ReplaceForbiddenChars(filename)
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

Public Function GetSubstFilename(InfoSpoolFileName As String, TokenFilename As String, _
 Optional WithoutAuthor As Boolean = False, Optional Preview As Boolean = False, _
 Optional bReplaceForbiddenChars As Boolean = True) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020  Dim Author As String, ClientComputer As String, ClientUsername As String, _
  Title As String, UserName As String, Computername As String, i As Long, _
  DateTime As String, filename As String, tStr As String, tList() As String, _
  Subst() As String, UserProfilPath As String, MyFiles As String, _
  MyDesktop As String, Path As String, isf As clsInfoSpoolFile, FilePath As String
50070  Dim tokenParameter As String, currentDateTime As Date, oldFileName As String, counterStr As String
50080
50090  If Len(TokenFilename) = 0 Then
50100   Exit Function
50110  End If
50120
50130  Set isf = New clsInfoSpoolFile
50140  isf.ReadInfoFile InfoSpoolFileName
50150
50160  If Preview = False Then
50170    If LenB(isf.FirstClientComputer) > 0 Then
50180     tStr = ReplaceForbiddenChars(isf.FirstClientComputer, "")
50190    End If
50200    If Mid$(tStr, 1, 2) = "\\" And IsIPAddress(Mid$(tStr, 3)) Then
50210     If Options.ClientComputerResolveIPAddress = 1 Then
50220      tStr = "\\" & GetHostNameFromIP(tStr)
50230     End If
50240    End If
50250    If LenB(tStr) = 0 Then
50260      ClientComputer = GetComputerName
50270     Else
50280      ClientComputer = tStr
50290    End If
50300
50310    If LenB(isf.FirstUserName) > 0 Then
50320     tStr = ReplaceForbiddenChars(isf.FirstUserName, "")
50330    End If
50340    If LenB(tStr) = 0 Then
50350      ClientUsername = GetUsername
50360     Else
50370      ClientUsername = tStr
50380    End If
50390   Else
50400    Title = "'Preview Title'"
50410    Author = "'Preview Author'"
50420    ClientComputer = "'Preview ClientComputer'"
50430    ClientUsername = "'Preview ClientUsername'"
50440  End If
50450
50460  If Options.FilenameSubstitutionsOnlyInTitle = 1 Then
50470   tList = Split(Options.FilenameSubstitutions, "\")
50480   Title = isf.FirstDocumentTitle
50490   If UBound(tList) >= 0 Then
50500    For i = 0 To UBound(tList)
50510     Subst = Split(tList(i), "|")
50520     If UBound(Subst) = 0 Then
50530       tStr = ""
50540      Else
50550       tStr = Subst(1)
50560     End If
50570     Title = Replace(Title, Subst(0), tStr, , , vbTextCompare)
50580    Next i
50590   End If
50600  End If
50610
50620  UserName = GetDocUsernameFromInfoSpoolFile(InfoSpoolFileName)
50630
50640  Computername = GetComputerName
50650  MyFiles = GetMyFiles
50660  MyDesktop = GetDesktop
50670
50680  filename = TokenFilename
50690
50700  currentDateTime = Now
50710  tStr = "DateTime"
50720  Do
50730   tokenParameter = Trim$(ExtractTokenParameter(filename, tStr))
50740   If LenB(tokenParameter) = 0 Then
50750    tokenParameter = Options.StandardDateformat
50760   End If
50770   DateTime = GetDocDate("", tokenParameter, CStr(currentDateTime))
50780   oldFileName = filename
50790   filename = ReplaceToken(filename, tStr, DateTime)
50800  Loop Until oldFileName = filename
50810
50820  filename = Replace(filename, "<Computername>", Computername, , , vbTextCompare)
50830
50840  filename = Replace(filename, "<Username>", GetUsername, , , vbTextCompare)
50850
50860  filename = Replace(filename, "<MyFiles>", CompletePath(MyFiles), , , vbTextCompare)
50870  filename = Replace(filename, "<MyDesktop>", CompletePath(MyDesktop), , , vbTextCompare)
50880
50890  If Options.Counter = 922337203685477@ Then
50900   Options.Counter = 0
50910  End If
50920  Options.Counter = Round(Options.Counter)
50930  filename = Replace(filename, "<Counter>", Format$(Options.Counter + 1, String(15, "0")), , , vbTextCompare)
50940
50950  filename = Replace(filename, "<Username>", GetUsername, , , vbTextCompare)
50960
50970  filename = Replace(filename, "<MyFiles>", CompletePath(MyFiles), , , vbTextCompare)
50980  filename = Replace(filename, "<MyDesktop>", CompletePath(MyDesktop), , , vbTextCompare)
50990
51000  If Options.Counter = 922337203685477@ Then
51010   Options.Counter = 0
51020  End If
51030  Options.Counter = Round(Options.Counter)
51040  tStr = "Counter"
51050  Do
51060   tokenParameter = Trim$(ExtractTokenParameter(filename, tStr))
51070   If LenB(tokenParameter) = 0 Then
51080    tokenParameter = String(15, "0")
51090   End If
51100   counterStr = Format$(Options.Counter + 1, tokenParameter)
51110   oldFileName = filename
51120   filename = ReplaceToken(filename, tStr, counterStr)
51130  Loop Until oldFileName = filename
51140
51150  tStr = "Title"
51160  If Preview = True Then
51170    filename = Replace(filename, "<" & tStr & ">", "'Preview " & tStr & "'", , , vbTextCompare)
51180   Else
51190    filename = Replace(filename, "<" & tStr & ">", Title, , , vbTextCompare)
51200  End If
51210  tStr = "DocumentFilename"
51220  If Preview = True Then
51230    filename = Replace(filename, "<" & tStr & ">", "'Preview " & tStr & "'", , , vbTextCompare)
51240   Else
51250    SplitPath isf.FirstDocumentTitle, , , , FilePath
51260    filename = Replace(filename, "<" & tStr & ">", FilePath, , , vbTextCompare)
51270  End If
51280  tStr = "DocumentPath"
51290  If Preview = True Then
51300    filename = Replace(filename, "<" & tStr & ">", "'Preview " & tStr & "'", , , vbTextCompare)
51310   Else
51320    SplitPath isf.FirstDocumentTitle, , FilePath
51330    filename = Replace(filename, "<" & tStr & ">", FilePath, , , vbTextCompare)
51340  End If
51350
51360  If WithoutAuthor = False Then
51370   tStr = "Author"
51380   If Preview = True Then
51390     filename = Replace(filename, "<" & tStr & ">", "'Preview " & tStr & "'", , , vbTextCompare)
51400    Else
51410     filename = Replace(filename, "<" & tStr & ">", isf.FirstUserName, , , vbTextCompare)
51420   End If
51430  End If
51440
51450  tStr = "JobID"
51460  If Preview = True Then
51470    filename = Replace(filename, "<" & tStr & ">", "'Preview " & tStr & "'", , , vbTextCompare)
51480   Else
51490    filename = Replace(filename, "<" & tStr & ">", isf.FirstJobID, , vbTextCompare)
51500  End If
51510  tStr = "ClientComputer"
51520  If Preview = True Then
51530    filename = Replace(filename, "<" & tStr & ">", "'Preview " & tStr & "'", , , vbTextCompare)
51540   Else
51550    filename = Replace(filename, "<" & tStr & ">", isf.FirstClientComputer, , , vbTextCompare)
51560  End If
51570  tStr = "PrinterName"
51580  If Preview = True Then
51590    filename = Replace(filename, "<" & tStr & ">", "'Preview " & tStr & "'", , , vbTextCompare)
51600   Else
51610    filename = Replace(filename, "<" & tStr & ">", isf.FirstPrinterName, , , vbTextCompare)
51620  End If
51630  tStr = "SessionID"
51640  If Preview = True Then
51650    filename = Replace(filename, "<" & tStr & ">", "'Preview " & tStr & "'", , , vbTextCompare)
51660   Else
51670    filename = Replace(filename, "<" & tStr & ">", isf.FirstSessionID, , , vbTextCompare)
51680  End If
51690
51700  If Options.FilenameSubstitutionsOnlyInTitle = 0 Then
51710   tList = Split(Options.FilenameSubstitutions, "\")
51720   If UBound(tList) >= 0 Then
51730    For i = 0 To UBound(tList)
51740     Subst = Split(tList(i), "|")
51750     If UBound(Subst) = 0 Then
51760       tStr = ""
51770      Else
51780       tStr = Subst(1)
51790     End If
51800     filename = Replace(filename, Subst(0), tStr, , , vbTextCompare)
51810    Next i
51820   End If
51830  End If
51840  If bReplaceForbiddenChars Then
51850   filename = ReplaceForbiddenChars(filename)
51860  End If
51870  If Options.RemoveSpaces = 1 Then
51880   filename = Trim$(filename)
51890  End If
51900  GetSubstFilename = filename
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

Public Function IsPostscriptFile(ByVal filename As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim lHeader As tPSHeader
50020  If NoPSCheck Or Options.NoPSCheck = 1 Then
50030   IsPostscriptFile = True
50040   Exit Function
50050  End If
50060  IsPostscriptFile = False
50070  lHeader = GetPSHeader(filename)
50080  If InStr(1, lHeader.StartComment.Comment, "PS", vbTextCompare) > 0 Then
50090   IsPostscriptFile = True
50100  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "IsPostscriptFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsPDFFile(ByVal filename As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long, bufStr As String
50020  IsPDFFile = False
50030  If FileExists(filename) Then
50040   If GetFileLength(filename) > 3 Then
50050    fn = FreeFile
50060    Open filename For Binary Access Read As fn
50070    bufStr = Space$(4)
50080    Get #fn, 1, bufStr
50090    Close #fn
50100    If UCase$(bufStr) = "%PDF" Then
50110     IsPDFFile = True
50120    End If
50130   End If
50140  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "IsPDFFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function AdjustCultureCalendar(dateStr As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If GetUserLocaleInfo(LOCALE_USER_DEFAULT&, LOCALE_ICALENDARTYPE) = 7 Then ' Thai
50020   Mid(dateStr, 1, 4) = CStr(CLng(Mid$(dateStr, 1, 4)) - 543) ' Adjust thai year
50030  End If
50040  AdjustCultureCalendar = dateStr
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "AdjustCultureCalendar")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetMetadataString(PDFDocInfo As tPDFDocInfo) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim DocInfoStr As String, PDFDocInfoStr As String, tzi As clsTimeZoneInformation, _
  tStr As String, ttStr As String, CodePage As Long, dStr As String
50030  CodePage = eCodePage.CP_UTF16
50040  With PDFDocInfo
50050   PDFDocInfoStr = PDFDocInfoStr & Chr$(13)
50060   If LenB(Trim$(.Author)) > 0 Then
50070     tStr = EncodeChars(CodePage, .Author)
50080    Else
50090     tStr = "()"
50100   End If
50110   PDFDocInfoStr = PDFDocInfoStr & Chr$(13) & "/Author " & tStr
50120   If LenB(.CreationDate) > 0 Or LenB(.ModifyDate) > 0 Then
50130    Set tzi = New clsTimeZoneInformation
50140    If tzi.DayLight Then
50150      ttStr = Format(TimeSerial(0, tzi.DaylightToGMT, 0), "hh'mm'")
50160     Else
50170      ttStr = Format(TimeSerial(0, tzi.NormaltimeToGMT, 0), "hh'mm'")
50180    End If
50190    If tzi.DaylightToGMT >= 0 Then
50200      ttStr = "+" & ttStr
50210     Else
50220      ttStr = "-" & ttStr
50230    End If
50240   End If
50250   If LenB(Trim$(.CreationDate)) > 0 Then
50260     tStr = "(D:" & AdjustCultureCalendar(.CreationDate) & ttStr & ")"
50270    Else
50280     tStr = "()"
50290   End If
50300   PDFDocInfoStr = PDFDocInfoStr & Chr$(13) & "/CreationDate " & tStr
50310   If LenB(.Creator) > 0 Then
50320     tStr = EncodeChars(CodePage, .Creator)
50330    Else
50340     tStr = "()"
50350   End If
50360   PDFDocInfoStr = PDFDocInfoStr & Chr$(13) & "/Creator " & tStr
50370   If LenB(Trim$(.Keywords)) > 0 Then
50380     tStr = EncodeChars(CodePage, .Keywords)
50390    Else
50400     tStr = "()"
50410   End If
50420   PDFDocInfoStr = PDFDocInfoStr & Chr$(13) & "/Keywords " & tStr
50430   If LenB(Trim$(.ModifyDate)) > 0 Then
50440     tStr = "(D:" & AdjustCultureCalendar(.ModifyDate) & ttStr & ")"
50450    Else
50460     tStr = "()"
50470   End If
50480   PDFDocInfoStr = PDFDocInfoStr & Chr$(13) & "/ModDate " & tStr
50490   If LenB(.CreationDate) > 0 Or LenB(.ModifyDate) > 0 Then
50500    Set tzi = Nothing
50510   End If
50520   If LenB(Trim$(.Subject)) > 0 Then
50530     tStr = EncodeChars(CodePage, .Subject)
50540    Else
50550     tStr = "()"
50560   End If
50570   PDFDocInfoStr = PDFDocInfoStr & Chr$(13) & "/Subject " & tStr
50580   If LenB(Trim$(.Title)) > 0 Then
50590     tStr = EncodeChars(CodePage, .Title)
50600    Else
50610     tStr = "()"
50620   End If
50630   PDFDocInfoStr = PDFDocInfoStr & Chr$(13) & "/Title " & tStr
50640  End With
50650  GetMetadataString = PDFDocInfoStr
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "GetMetadataString")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function CreatePDFDocInfoFile(InfoSpoolFile As String, PDFDocInfo As tPDFDocInfo)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long, MetadataString As String, DocInfoStr As String, Path As String, File As String, PDFDocInfoFile As String
50020  MetadataString = GetMetadataString(PDFDocInfo)
50030  SplitPath InfoSpoolFile, , Path, , File
50040  PDFDocInfoFile = CompletePath(Path) & File & ".mtd"
50050  If FileExists(InfoSpoolFile) = True And LenB(MetadataString) > 0 Then
50060   DocInfoStr = Chr$(13) & "/pdfmark where {pop} {userdict /pdfmark /cleartomark load put} ifelse"
50070   DocInfoStr = DocInfoStr & Chr$(13) & "["
50080   DocInfoStr = DocInfoStr & Chr$(13) & MetadataString
50090   DocInfoStr = DocInfoStr & Chr$(13) & "/DOCINFO pdfmark"
50100   DocInfoStr = DocInfoStr & Chr$(13) & "%%EOF"
50110   fn = FreeFile
50120   Open PDFDocInfoFile For Output As fn
50130   Print #fn, DocInfoStr;
50140   Close #fn
50150   CreatePDFDocInfoFile = PDFDocInfoFile
50160  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "CreatePDFDocInfoFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function CreatePDFDocViewFile(InfoSpoolFile As String, PageLayout As Long, PageMode As Long, StartPage As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long, DocViewStr As String, Path As String, File As String, PDFDocViewFile As String
50020  If PageLayout > 0 Or PageMode > 0 Or StartPage <> 1 Then
50030   SplitPath InfoSpoolFile, , Path, , File
50040   PDFDocViewFile = CompletePath(Path) & File & ".dvw"
50050   If FileExists(InfoSpoolFile) = True Then
50060    DocViewStr = Chr$(13) & "/pdfmark where {pop} {userdict /pdfmark /cleartomark load put} ifelse"
50070    DocViewStr = DocViewStr & Chr$(13) & "["
50080    If PageLayout > 0 Then
50090     DocViewStr = DocViewStr & Chr$(13) & "/PageLayout"
50101     Select Case PageLayout
           Case 1:
50120       DocViewStr = DocViewStr & " /OneColumn"
50130      Case 2:
50140       DocViewStr = DocViewStr & " /TwoColumnLeft"
50150      Case 3:
50160       DocViewStr = DocViewStr & " /TwoColumnRight"
50170      Case 4:
50180       DocViewStr = DocViewStr & " /TwoPageLeft"
50190      Case 5:
50200       DocViewStr = DocViewStr & " /TwoPageRight"
50210      Case Else:
50220       DocViewStr = DocViewStr & " /SinglePage"
50230     End Select
50240    End If
50250    If PageMode > 0 Then
50260     DocViewStr = DocViewStr & Chr$(13) & "/PageMode"
50271     Select Case PageMode
           Case 1:
50290       DocViewStr = DocViewStr & " /UseOutlines"
50300      Case 2:
50310       DocViewStr = DocViewStr & " /UseThumbs"
50320      Case 3:
50330       DocViewStr = DocViewStr & " /FullScreen"
50340      Case 4:
50350       DocViewStr = DocViewStr & " /UseOC"
50360      Case 5:
50370       DocViewStr = DocViewStr & " /UseAttachments"
50380      Case Else:
50390       DocViewStr = DocViewStr & " /UseNone"
50400     End Select
50410    End If
50420    If StartPage > 1 Then
50430     DocViewStr = DocViewStr & " /Page " & CStr(StartPage)
50440    End If
50450    DocViewStr = DocViewStr & Chr$(13) & "/DOCVIEW pdfmark"
50460    DocViewStr = DocViewStr & Chr$(13) & "%%EOF"
50470    fn = FreeFile
50480    Open PDFDocViewFile For Output As fn
50490    Print #fn, DocViewStr;
50500    Close #fn
50510    CreatePDFDocViewFile = PDFDocViewFile
50520   End If
50530  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "CreatePDFDocViewFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function ReplaceEncodingChars(ByVal Str1 As String) As String
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

Public Function GetDocUsernameFromInfoSpoolFile(InfoSpoolFile As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim UserName As String, isf As clsInfoSpoolFile
50020  If LenB(InfoSpoolFile) > 0 Then
50030   If FileExists(InfoSpoolFile) Then
50040    Set isf = New clsInfoSpoolFile
50050    isf.ReadInfoFile InfoSpoolFile
50060    If LenB(isf.FirstUserName) > 0 Then
50070     UserName = isf.FirstUserName
50080    End If
50090   End If
50100  End If
50110  If Len(UserName) = 0 Then
50120   UserName = GetUsername
50130  End If
50140  GetDocUsernameFromInfoSpoolFile = UserName
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "GetDocUsernameFromInfoSpoolFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetDocUsernameFromPostScriptFile(PostscriptFile As String, NoFile As Boolean) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim UserName As String, PSHeader As tPSHeader
50020  If NoFile = False Then
50030   If Len(PostscriptFile) > 0 Then
50040    If FileExists(PostscriptFile) = True Then
50050     If IsPostscriptFile(PostscriptFile) = True Then
50060      If LenB(UserName) = 0 Then
50070       PSHeader = GetPSHeader(PostscriptFile)
50080       If Len(PSHeader.CreateFor.Comment) > 0 Then
50090        UserName = PSHeader.CreateFor.Comment
50100       End If
50110      End If
50120     End If
50130    End If
50140   End If
50150  End If
50160  If Len(UserName) = 0 Then
50170   UserName = GetUsername
50180  End If
50190  GetDocUsernameFromPostScriptFile = UserName
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "GetDocUsernameFromPostScriptFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

'Public Function GetClientMachineFromInfoSpoolFile(InfoSpoolFile As String) As String
' Dim ClientMachine As String, isf As clsInfoSpoolFile
'
' If LenB(InfoSpoolFile) > 0 Then
'  If FileExists(InfoSpoolFile) Then
'   Set isf = New clsInfoSpoolFile
'   isf.ReadInfoFile InfoSpoolFile
'   If LenB(isf.FirstUserName) > 0 Then
'    ClientMachine = isf.FirstClientComputer
'   End If
'  End If
' End If
' If Len(ClientMachine) = 0 Then
'  ClientMachine = GetComputerName
' End If
' GetClientMachineFromInfoSpoolFile = ClientMachine
'End Function

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

Public Function EncodeChars(ByVal CodePage As eCodePage, ByVal Str1 As String) As String ' UTF-16, UTF-8 conversion
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, tStr As String, Size As Long, buffer() As Byte, c As Long, tL As Long
50020
50030  If LenB(Str1) = 0 Then
50040   Exit Function
50050  End If
50061  Select Case CodePage
        Case eCodePage.CP_NoEncoding
50080    EncodeChars = "(" & Str1 & ")"
50090   Case eCodePage.CP_UTF8
50100    For i = 1 To Len(Str1)
50110     c = AscW(Mid$(Str1, i, 1))
50121     Select Case c
           Case 0 To &H7F&
50140       tStr = tStr & String(2 - Len(Hex(c)), "0") & Hex(c)
50150      Case &H80& To &H7FF&
50160       tL = &HC0& Or ((c And &H3FC0&) \ &H40&)
50170       tStr = tStr & String(2 - Len(Hex(tL)), "0") & Hex(tL)
50180       tL = &H80& Or (c And &H3F&)
50190       tStr = tStr & String(2 - Len(Hex(tL)), "0") & Hex(tL)
50200      Case &H800& To &HFFFF&
50210       tL = &HE0& Or ((c And &HF000&) \ &H1000&)
50220       tStr = tStr & String(2 - Len(Hex(tL)), "0") & Hex(tL)
50230       tL = &H80& Or ((c And &HFC0&) \ &H40&)
50240       tStr = tStr & String(2 - Len(Hex(tL)), "0") & Hex(tL)
50250       tL = &H80& Or (c And &H3F&)
50260       tStr = tStr & String(2 - Len(Hex(tL)), "0") & Hex(tL)
50270     End Select
50280    Next i
50290    EncodeChars = "<BFBBEF" & tStr & ">"
50300 '  Case eCodePage.CP_UTF8
50310 '   Size = 3
50320 '   ReDim Buffer(0 To 2)
50330 '   MoveMemoryLongToByte Buffer(0), ESignatureUTF8, 3
50340 '   For i = 1 To Len(Str1)
50350 '    c = MCh(Str1, i)
50360 '    Select Case c
50370 '     Case 0 To &H7F&
50380 '      ReDim Preserve Buffer(0 To Size)
50390 '      Buffer(Size) = c
50400 '      Size = Size + 1
50410 '     Case &H80& To &H7FF&
50420 '      ReDim Preserve Buffer(0 To Size + 1)
50430 '      Buffer(Size) = &HC0& Or ((c And &H3FC0&) \ &H40&)
50440 '      Buffer(Size + 1) = &H80& Or (c And &H3F&)
50450 '      Size = Size + 2
50460 '     Case &H800& To &HFFFF&
50470 '      ReDim Preserve Buffer(0 To Size + 2)
50480 '      Buffer(Size) = &HE0& Or ((c And &HF000&) \ &H1000&)
50490 '      Buffer(Size + 1) = &H80& Or ((c And &HFC0&) \ &H40&)
50500 '      Buffer(Size + 2) = &H80& Or (c And &H3F&)
50510 '      Size = Size + 3
50520 '     End Select
50530 '    Next i
50540   Case eCodePage.CP_UTF16
50550    For i = 1 To Len(Str1)
50560     c = AscW(Mid$(Str1, i, 1))
50570     tStr = tStr & Right("0000" & Hex(c), 4)
50580    Next i
50590    EncodeChars = "<FEFF" & tStr & ">"
50600  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "EncodeChars")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function RemoveEncodedPostscriptChars(ByVal Str1 As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If LenB(Str1) = 0 Then
50020   RemoveEncodedPostscriptChars = ""
50030  End If
50040  Str1 = Replace(Str1, "\\", "\")
50050  Str1 = Replace(Str1, "\{", "{")
50060  Str1 = Replace(Str1, "\}", "}")
50070  Str1 = Replace(Str1, "\[", "[")
50080  Str1 = Replace(Str1, "\]", "]")
50090  Str1 = Replace(Str1, "\(", "(")
50100  Str1 = Replace(Str1, "\)", ")")
50110  RemoveEncodedPostscriptChars = Str1
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "RemoveEncodedPostscriptChars")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function EncodeCharsOctal(ByVal Str1 As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, tStr As String
50020  Str1 = Replace(Str1, "\", "\\")
50030  Str1 = Replace(Str1, "{", "\{")
50040  Str1 = Replace(Str1, "}", "\}")
50050  Str1 = Replace(Str1, "[", "\[")
50060  Str1 = Replace(Str1, "]", "\]")
50070  Str1 = Replace(Str1, "(", "\(")
50080  Str1 = Replace(Str1, ")", "\)")
50090  For i = 1 To Len(Str1)
50100   If Asc(Mid(Str1, i, 1)) > 127 Then
50110     tStr = tStr & "\" & CStr(Oct(Asc(Mid(Str1, i, 1))))
50120    Else
50130     tStr = tStr & Mid(Str1, i, 1)
50140   End If
50150  Next i
50160  EncodeCharsOctal = tStr
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "EncodeCharsOctal")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
