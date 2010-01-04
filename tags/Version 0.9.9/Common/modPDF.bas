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
 title As tPSComment
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
 title As String
End Type

Public Enum eCodePage
 CP_NoEncoding = 0
 CP_UTF8 = 65001
 CP_UTF16 = 65002
End Enum

Public Function GetPSTitle(filename As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStr As String, psH As tPSHeader, isf As InfoSpoolFile
50020  isf = ReadInfoSpoolfile(filename)
50030  tStr = isf.REDMON_DOCNAME
50040  If LenB(tStr) = 0 Then
50050   psH = GetPSHeader(filename)
50060   tStr = psH.title.Comment
50070  End If
50080  GetPSTitle = tStr
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

Public Function GetPSHeader(filename As String) As tPSHeader
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long, bufStr As String, PSHeader As tPSHeader, buffer As Long
50020  If FileExists(filename) And FileInUse(filename) = False Then
50030   DoEvents
50040   fn = FreeFile
50050   If FileLen(filename) = 0 Then
50060    Exit Function
50070   End If
50080   buffer = 5000
50090   If FileLen(filename) < buffer Then
50100    buffer = FileLen(filename)
50110   End If
50120
50130   Open filename For Binary Access Read As fn
50140   bufStr = Space$(buffer)
50150   Get #fn, 1, bufStr
50160   Close #fn
50170
50180   With PSHeader
50190    .StartComment = GetPSComment(bufStr, "%!")
50200    .CreateFor = GetPSComment(bufStr, "%%For:")
50210    .CreationDate = GetPSComment(bufStr, "%%CreationDate:")
50220    .Creator = GetPSComment(bufStr, "%%Creator:")
50230    .Pages = GetPSComment(bufStr, "%%Pages:")
50240    .title = GetPSComment(bufStr, "%%Title:")
50250    .EndComment = GetPSComment(bufStr, "%%EndComments")
50260   End With
50270   GetPSHeader = PSHeader
50280   DoEvents
50290  End If
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
50010  Dim PSComment As tPSComment
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

Public Function GetAutosaveFilename(PostscriptFile As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim filename As String, Pathname As String
50020
50030  If Options.UseAutosaveDirectory = 1 Then
50040    Pathname = GetSubstFilename2(Options.AutosaveDirectory, False, , PostscriptFile)
50050   Else
50060    Pathname = GetSubstFilename2(Options.LastSaveDirectory, False)
50070  End If
50080
50090  filename = GetSubstFilename(PostscriptFile, Options.AutosaveFilename)
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
50570  GetAutosaveFilename = CompletePath(GetSubstFilename2(Pathname, False)) & ReplaceForbiddenChars(filename)
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

Public Function GetSubstFilename(PostscriptFile As String, TokenFilename As String, _
 Optional WithoutAuthor As Boolean = False, Optional Preview As Boolean = False, _
 Optional NoReplaceForbiddenChars As Boolean = False) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim PSHeader As tPSHeader, Author As String, ClientComputer As String, _
  title As String, UserName As String, Computername As String, i As Long, _
  DateTime As String, filename As String, tStr As String, tList() As String, _
  Subst() As String, UserProfilPath As String, MyFiles As String, _
  MyDesktop As String, Path As String, isf As InfoSpoolFile, FilePath As String
50060
50070  If Len(TokenFilename) = 0 Then
50080   Exit Function
50090  End If
50100
50110  DateTime = GetDocDate("", Options.StandardDateformat, CStr(Now))
50120  If Preview = False Then
50130    If FileExists(PostscriptFile) = True Then
50140     PSHeader = GetPSHeader(PostscriptFile)
50150     isf = ReadInfoSpoolfile(PostscriptFile)
50160    End If
50170    If Options.UseStandardAuthor = 1 Then
50180      Author = Options.StandardAuthor
50190     Else
50200      Author = PSHeader.CreateFor.Comment
50210    End If
50220    If LenB(isf.REDMON_MACHINE) > 0 Then
50230      tStr = ReplaceForbiddenChars(isf.REDMON_MACHINE, "")
50240     Else
50250      tStr = ReplaceForbiddenChars(Environ$("REDMON_MACHINE"), "")
50260    End If
50270    If Mid$(tStr, 1, 2) = "\\" And IsIPAddress(Mid$(tStr, 3)) Then
50280     If Options.ClientComputerResolveIPAddress = 1 Then
50290      tStr = "\\" & GetHostNameFromIP(tStr)
50300     End If
50310    End If
50320    If LenB(tStr) = 0 Then
50330      ClientComputer = GetComputerName
50340     Else
50350      ClientComputer = tStr
50360    End If
50370   Else
50380    PSHeader.title.Comment = "'Preview Title'"
50390    Author = "'Preview Author'"
50400    ClientComputer = "'Preview ClientComputer'"
50410  End If
50420
50430  If Options.FilenameSubstitutionsOnlyInTitle = 1 Then
50440   tList = Split(Options.FilenameSubstitutions, "\")
50450   title = PSHeader.title.Comment
50460   If UBound(tList) >= 0 Then
50470    For i = 0 To UBound(tList)
50480     Subst = Split(tList(i), "|")
50490     If UBound(Subst) = 0 Then
50500       tStr = ""
50510      Else
50520       tStr = Subst(1)
50530     End If
50540     title = Replace(title, Subst(0), tStr, , , vbTextCompare)
50550    Next i
50560   End If
50570  End If
50580
50590  UserName = GetDocUsername(PostscriptFile, Preview)
50600
50610  Computername = GetComputerName
50620  MyFiles = GetMyFiles
50630  MyDesktop = GetDesktop
50640
50650  filename = TokenFilename
50660  filename = Replace(filename, "<DateTime>", DateTime, , , vbTextCompare)
50670  filename = Replace(filename, "<Computername>", Computername, , , vbTextCompare)
50680
50690  filename = Replace(filename, "<ClientComputer>", ClientComputer, , , vbTextCompare)
50700  filename = Replace(filename, "<Username>", UserName, , , vbTextCompare)
50710  filename = Replace(filename, "<Title>", title, , , vbTextCompare)
50720  If WithoutAuthor = False Then
50730   filename = Replace(filename, "<Author>", Author, , , vbTextCompare)
50740  End If
50750
50760  filename = Replace(filename, "<MyFiles>", CompletePath(MyFiles), , , vbTextCompare)
50770  filename = Replace(filename, "<MyDesktop>", CompletePath(MyDesktop), , , vbTextCompare)
50780
50790  If Options.Counter = 922337203685477@ Then
50800   Options.Counter = 0
50810  End If
50820  Options.Counter = Round(Options.Counter)
50830  filename = Replace(filename, "<Counter>", Format$(Options.Counter + 1, String(15, "0")), , , vbTextCompare)
50840
50850  tStr = "DOCNAME"
50860  If Preview = True Then
50870    filename = Replace(filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
50880   Else
50890    filename = Replace(filename, "<REDMON_" & tStr & ">", isf.REDMON_DOCNAME, , , vbTextCompare)
50900  End If
50910
50920  tStr = "DOCNAME_FILE"
50930  If Preview Then
50940    filename = Replace(filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
50950   Else
50960    SplitPath isf.REDMON_DOCNAME, , , , FilePath
50970    filename = Replace(filename, "<REDMON_" & tStr & ">", FilePath, , , vbTextCompare)
50980  End If
50990  tStr = "DOCNAME_PATH"
51000  If Preview Then
51010    filename = Replace(filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
51020   Else
51030    SplitPath isf.REDMON_DOCNAME, , FilePath
51040    filename = Replace(filename, "<REDMON_" & tStr & ">", FilePath, , , vbTextCompare)
51050  End If
51060
51070  tStr = "JOB"
51080  If Preview = True Then
51090    filename = Replace(filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
51100   Else
51110    filename = Replace(filename, "<REDMON_" & tStr & ">", isf.REDMON_JOB, , , vbTextCompare)
51120  End If
51130  tStr = "MACHINE"
51140  If Preview = True Then
51150    filename = Replace(filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
51160   Else
51170    filename = Replace(filename, "<REDMON_" & tStr & ">", isf.REDMON_MACHINE, , , vbTextCompare)
51180  End If
51190  tStr = "PORT"
51200  If Preview = True Then
51210    filename = Replace(filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
51220   Else
51230    filename = Replace(filename, "<REDMON_" & tStr & ">", isf.REDMON_PORT, , , vbTextCompare)
51240  End If
51250  tStr = "PRINTER"
51260  If Preview = True Then
51270    filename = Replace(filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
51280   Else
51290    filename = Replace(filename, "<REDMON_" & tStr & ">", isf.REDMON_PRINTER, , , vbTextCompare)
51300  End If
51310  tStr = "SESSIONID"
51320  If Preview = True Then
51330    filename = Replace(filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
51340   Else
51350    filename = Replace(filename, "<REDMON_" & tStr & ">", isf.REDMON_SESSIONID, , , vbTextCompare)
51360  End If
51370  tStr = "USER"
51380  If Preview = True Then
51390    filename = Replace(filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
51400   Else
51410    filename = Replace(filename, "<REDMON_" & tStr & ">", isf.REDMON_USER, , , vbTextCompare)
51420  End If
51430
51440  If Options.FilenameSubstitutionsOnlyInTitle = 0 Then
51450   tList = Split(Options.FilenameSubstitutions, "\")
51460   If UBound(tList) >= 0 Then
51470    For i = 0 To UBound(tList)
51480     Subst = Split(tList(i), "|")
51490     If UBound(Subst) = 0 Then
51500       tStr = ""
51510      Else
51520       tStr = Subst(1)
51530     End If
51540     filename = Replace(filename, Subst(0), tStr, , , vbTextCompare)
51550    Next i
51560   End If
51570  End If
51580  If Not NoReplaceForbiddenChars Then
51590   filename = ReplaceForbiddenChars(filename)
51600  End If
51610  If Options.RemoveSpaces = 1 Then
51620   filename = Trim$(filename)
51630  End If
51640  GetSubstFilename = filename
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
50040   If FileLen(filename) > 3 Then
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

Public Function GetMetadataString(PDFDocInfo As tPDFDocInfo) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim DocInfoStr As String, PDFDocInfoStr As String, tzi As clsTimeZoneInformation, _
  tStr As String, ttStr As String, CodePage As Long
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
50260     tStr = "(D:" & .CreationDate & ttStr & ")"
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
50440     tStr = "(D:" & .ModifyDate & ttStr & ")"
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
50580   If LenB(Trim$(.title)) > 0 Then
50590     tStr = EncodeChars(CodePage, .title)
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

Public Sub AppendPDFDocInfo(PSFile As String, PDFDocInfo As tPDFDocInfo)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long, MetadataString As String, DocInfoStr As String
50020  MetadataString = GetMetadataString(PDFDocInfo)
50030  If FileExists(PSFile) = True And LenB(MetadataString) > 0 Then
50040   DocInfoStr = Chr$(13) & "/pdfmark where {pop} {userdict /pdfmark /cleartomark load put} ifelse"
50050   DocInfoStr = DocInfoStr & Chr$(13) & "["
50060   DocInfoStr = DocInfoStr & Chr$(13) & MetadataString
50070   DocInfoStr = DocInfoStr & Chr$(13) & "/DOCINFO pdfmark"
50080   DocInfoStr = DocInfoStr & Chr$(13) & "%%EOF"
50090   fn = FreeFile
50100   Open PSFile For Append As fn
50110   Print #fn, DocInfoStr;
50120   Close #fn
50130  End If
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

Public Function GetDocUsername(PostscriptFile As String, NoFile As Boolean) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim UserName As String, Path As String, i As Long, PSHeader As tPSHeader, _
  isf As InfoSpoolFile
50030  If NoFile = False Then
50040   If Len(PostscriptFile) > 0 Then
50050    If FileExists(PostscriptFile) = True Then
50060     isf = ReadInfoSpoolfile(PostscriptFile)
50070     If LenB(isf.REDMON_USER) > 0 Then
50080      UserName = isf.REDMON_USER
50090     End If
50100     If LenB(UserName) = 0 Then
50110      PSHeader = GetPSHeader(PostscriptFile)
50120      If Len(PSHeader.CreateFor.Comment) > 0 Then
50130       UserName = PSHeader.CreateFor.Comment
50140      End If
50150     End If
50160    End If
50170   End If
50180  End If
50190  If LenB(UserName) = 0 Then
50200   UserName = Environ$("Redmon_User")
50210  End If
50220  If Len(UserName) = 0 Then
50230   UserName = GetUsername
50240  End If
50250  GetDocUsername = UserName
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "GetDocUsername")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetClientMachine(PostscriptFile As String, NoFile As Boolean) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ClientMachine As String, Path As String, i As Long, PSHeader As tPSHeader, _
  isf As InfoSpoolFile
50030  If NoFile = False Then
50040   If Len(PostscriptFile) > 0 Then
50050    If FileExists(PostscriptFile) = True Then
50060     isf = ReadInfoSpoolfile(PostscriptFile)
50070     If LenB(isf.REDMON_MACHINE) > 0 Then
50080      ClientMachine = isf.REDMON_MACHINE
50090     End If
50100    End If
50110   End If
50120  End If
50130  If LenB(ClientMachine) = 0 Then
50140   ClientMachine = Environ$("Redmon_Machine")
50150  End If
50160  If Len(ClientMachine) = 0 Then
50170   ClientMachine = GetComputerName
50180  End If
50190  GetClientMachine = ClientMachine
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDF", "GetClientMachine")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
