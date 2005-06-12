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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tstr As String, psH As tPSHeader, isf As InfoSpoolFile
50020  isf = ReadInfoSpoolfile(Filename)
50030  tstr = isf.REDMON_DOCNAME
50040  If LenB(tstr) = 0 Then
50050   psH = GetPSHeader(Filename)
50060   tstr = psH.Title.Comment
50070  End If
50080  GetPSTitle = tstr
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
50230    .Pages = GetPSComment(bufStr, "%%Pages:")
50240    .Title = GetPSComment(bufStr, "%%Title:")
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

Private Function GetPSHeaderString(PSHeader As tPSHeader) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tstr As String
50020  If PSHeader.StartComment.StartByte = -1 Then
50030    tstr = "%!PS-Adobe-3.0" & Chr$(&HA)
50040   Else
50050    tstr = "%!" & PSHeader.StartComment.Comment & Chr$(&HA)
50060  End If
50070
50080  tstr = tstr & "%%For:" & PSHeader.CreateFor.Comment & Chr$(&HA)
50090  tstr = tstr & "%%CreationDate:" & PSHeader.CreationDate.Comment & Chr$(&HA)
50100  tstr = tstr & "%%Creator:" & PSHeader.Creator.Comment & Chr$(&HA)
50110  tstr = tstr & "%%Title:" & PSHeader.Title.Comment & Chr$(&HA)
50120
50130  tstr = tstr & "%%EndComments" & Chr$(&HA)
50140  GetPSHeaderString = tstr
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
50040    Pathname = CompletePath(Trim$(Options.AutosaveDirectory))
50050   Else
50060    Pathname = CompletePath(Trim$(Options.LastSaveDirectory))
50070  End If
50080
50090  Filename = GetSubstFilename(Postscriptfile, Options.AutosaveFilename)
50100
50110  If Len(Filename) > 0 And Options.RemoveAllKnownFileExtensions = 1 Then
50120   Filename = RemoveAllKnownFileExtensions(Filename)
50130  End If
50140
50151  Select Case Options.AutosaveFormat
              Case 0: 'PDF
50170    Filename = Filename & ".pdf"
50180   Case 1: 'PNG
50190    Filename = Filename & ".png"
50200   Case 2: 'JPEG
50210    Filename = Filename & ".jpg"
50220   Case 3: 'BMP
50230    Filename = Filename & ".bmp"
50240   Case 4: 'PCX
50250    Filename = Filename & ".pcx"
50260   Case 5: 'TIFF
50270    Filename = Filename & ".tif"
50280   Case 6: 'PS
50290    Filename = Filename & ".ps"
50300   Case 7: 'EPS
50310    Filename = Filename & ".eps"
50320  End Select
50330  GetAutosaveFilename = CompletePath(GetSubstFilename(Postscriptfile, Pathname)) & ReplaceForbiddenChars(Filename)
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
50010  Dim PSHeader As tPSHeader, Author As String, ClientComputer As String, _
  Title As String, UserName As String, Computername As String, i As Long, _
  DateTime As String, Filename As String, tstr As String, tList() As String, _
  Subst() As String, UserProfilPath As String, MyFiles As String, _
  MyDesktop As String, Path As String, isf As InfoSpoolFile
50060
50070  If Len(TokenFilename) = 0 Then
50080   Exit Function
50090  End If
50100
50110  DateTime = GetDocDate("", Options.StandardDateformat, CStr(Now))
50120  If Preview = False Then
50130    If FileExists(Postscriptfile) = True Then
50140     PSHeader = GetPSHeader(Postscriptfile)
50150     isf = ReadInfoSpoolfile(Postscriptfile)
50160    End If
50170    If Options.UseStandardAuthor = 1 Then
50180      Author = Options.StandardAuthor
50190     Else
50200      Author = PSHeader.CreateFor.Comment
50210    End If
50220    If LenB(isf.REDMON_MACHINE) > 0 Then
50230      ClientComputer = ReplaceForbiddenChars(isf.REDMON_MACHINE, "")
50240     Else
50250      ClientComputer = ReplaceForbiddenChars(Environ$("REDMON_MACHINE"), "")
50260      If LenB(ClientComputer) = 0 Then
50270       ClientComputer = GetComputerName
50280      End If
50290    End If
50300   Else
50310    PSHeader.Title.Comment = "'Preview Title'"
50320    Author = "'Preview Author'"
50330    ClientComputer = "'Preview ClientComputer'"
50340  End If
50350
50360  If Options.FilenameSubstitutionsOnlyInTitle = 1 Then
50370   tList = Split(Options.FilenameSubstitutions, "\")
50380   Title = PSHeader.Title.Comment
50390   If UBound(tList) >= 0 Then
50400    For i = 0 To UBound(tList)
50410     Subst = Split(tList(i), "|")
50420     If UBound(Subst) = 0 Then
50430       tstr = ""
50440      Else
50450       tstr = Subst(1)
50460     End If
50470     Title = Replace(Title, Subst(0), tstr, , , vbTextCompare)
50480    Next i
50490   End If
50500  End If
50510
50520  UserName = GetDocUsername(Postscriptfile, Preview)
50530
50540  Computername = GetComputerName
50550  MyFiles = GetMyFiles
50560  MyDesktop = GetDesktop
50570
50580  Filename = TokenFilename
50590  Filename = Replace(Filename, "<DateTime>", DateTime, , , vbTextCompare)
50600  Filename = Replace(Filename, "<Computername>", Computername, , , vbTextCompare)
50610
50620  Filename = Replace(Filename, "<ClientComputer>", ClientComputer, , , vbTextCompare)
50630  Filename = Replace(Filename, "<Username>", UserName, , , vbTextCompare)
50640  Filename = Replace(Filename, "<Title>", Title, , , vbTextCompare)
50650  If WithoutAuthor = False Then
50660   Filename = Replace(Filename, "<Author>", Author, , , vbTextCompare)
50670  End If
50680  Filename = Replace(Filename, "<MyFiles>", MyFiles, , , vbTextCompare)
50690  Filename = Replace(Filename, "<MyDesktop>", MyDesktop, , , vbTextCompare)
50700
50710  tstr = "DOCNAME"
50720  If Preview = True Then
50730    Filename = Replace(Filename, "<REDMON_" & tstr & ">", "'Preview REDMON_" & tstr & "'", , , vbTextCompare)
50740   Else
50750    Filename = Replace(Filename, "<REDMON_" & tstr & ">", isf.REDMON_DOCNAME, , , vbTextCompare)
50760  End If
50770  tstr = "JOB"
50780  If Preview = True Then
50790    Filename = Replace(Filename, "<REDMON_" & tstr & ">", "'Preview REDMON_" & tstr & "'", , , vbTextCompare)
50800   Else
50810    Filename = Replace(Filename, "<REDMON_" & tstr & ">", isf.REDMON_JOB, , , vbTextCompare)
50820  End If
50830  tstr = "MACHINE"
50840  If Preview = True Then
50850    Filename = Replace(Filename, "<REDMON_" & tstr & ">", "'Preview REDMON_" & tstr & "'", , , vbTextCompare)
50860   Else
50870    Filename = Replace(Filename, "<REDMON_" & tstr & ">", isf.REDMON_MACHINE, , , vbTextCompare)
50880  End If
50890  tstr = "PORT"
50900  If Preview = True Then
50910    Filename = Replace(Filename, "<REDMON_" & tstr & ">", "'Preview REDMON_" & tstr & "'", , , vbTextCompare)
50920   Else
50930    Filename = Replace(Filename, "<REDMON_" & tstr & ">", isf.REDMON_PORT, , , vbTextCompare)
50940  End If
50950  tstr = "PRINTER"
50960  If Preview = True Then
50970    Filename = Replace(Filename, "<REDMON_" & tstr & ">", "'Preview REDMON_" & tstr & "'", , , vbTextCompare)
50980   Else
50990    Filename = Replace(Filename, "<REDMON_" & tstr & ">", isf.REDMON_PRINTER, , , vbTextCompare)
51000  End If
51010  tstr = "SESSIONID"
51020  If Preview = True Then
51030    Filename = Replace(Filename, "<REDMON_" & tstr & ">", "'Preview REDMON_" & tstr & "'", , , vbTextCompare)
51040   Else
51050    Filename = Replace(Filename, "<REDMON_" & tstr & ">", isf.REDMON_SESSIONID, , , vbTextCompare)
51060  End If
51070  tstr = "USER"
51080  If Preview = True Then
51090    Filename = Replace(Filename, "<REDMON_" & tstr & ">", "'Preview REDMON_" & tstr & "'", , , vbTextCompare)
51100   Else
51110    Filename = Replace(Filename, "<REDMON_" & tstr & ">", isf.REDMON_USER, , , vbTextCompare)
51120  End If
51130
51140  If Options.FilenameSubstitutionsOnlyInTitle = 0 Then
51150   tList = Split(Options.FilenameSubstitutions, "\")
51160   If UBound(tList) >= 0 Then
51170    For i = 0 To UBound(tList)
51180     Subst = Split(tList(i), "|")
51190     If UBound(Subst) = 0 Then
51200       tstr = ""
51210      Else
51220       tstr = Subst(1)
51230     End If
51240     Filename = Replace(Filename, Subst(0), tstr, , , vbTextCompare)
51250    Next i
51260   End If
51270  End If
51280  If Options.RemoveSpaces = 1 Then
51290   Filename = Trim$(Filename)
51300  End If
51310  GetSubstFilename = Filename
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

Public Function IsPostscriptFile(ByVal Filename As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim lHeader As tPSHeader
50020  IsPostscriptFile = False
50030  lHeader = GetPSHeader(Filename)
50040  If InStr(1, lHeader.StartComment.Comment, "PS", vbTextCompare) > 0 Then
50050   IsPostscriptFile = True
50060  End If
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

Public Sub AppendPDFDocInfo(PSFile As String, PDFDocInfo As tPDFDocInfo)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long, DocInfoStr As String, PDFDocInfoStr As String, tzi As clsTimeZoneInformation, tstr As String
50020  With PDFDocInfo
50030   If LenB(.Author) > 0 Then
50040    PDFDocInfoStr = PDFDocInfoStr & Chr$(13) & "/Author (" & EncodeChars(.Author) & ")"
50050   End If
50060   If LenB(.CreationDate) > 0 Or LenB(.ModifyDate) > 0 Then
50070    Set tzi = New clsTimeZoneInformation
50080    tstr = Format(TimeSerial(0, tzi.DaylightToGMT, 0), "hh'mm")
50090    If tzi.DaylightToGMT >= 0 Then
50100      tstr = "+" & tstr
50110     Else
50120      tstr = "-" & tstr
50130    End If
50140   End If
50150   If LenB(Trim$(.CreationDate)) > 0 Then
50160    PDFDocInfoStr = PDFDocInfoStr & Chr$(13) & "/CreationDate (D:" & EncodeChars(.CreationDate) & tstr & ")"
50170   End If
50180   If LenB(.Creator) > 0 Then
50190    PDFDocInfoStr = PDFDocInfoStr & Chr$(13) & "/Creator (" & EncodeChars(.Creator) & ")"
50200   End If
50210   If LenB(.Keywords) > 0 Then
50220    PDFDocInfoStr = PDFDocInfoStr & Chr$(13) & "/Keywords (" & EncodeChars(.Keywords) & ")"
50230   End If
50240   If LenB(Trim$(.ModifyDate)) > 0 Then
50250    PDFDocInfoStr = PDFDocInfoStr & Chr$(13) & "/ModDate (D:" & EncodeChars(.ModifyDate) & tstr & ")"
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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Integer, tstr As String
50020  tstr = ""
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
50170          tstr = tstr & Chr$(CByte("&H" & Mid$(Str1, i, 2)))
50180        End If
50190       Else
50200        Exit Function
50210      End If
50220     Next i
50230    End If
50240   End If
50250  End If
50260  If Len(tstr) > 0 Then
50270   ReplaceEncodingChars = tstr
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

Public Function GetDocUsername(Postscriptfile As String, NoFile As Boolean) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim UserName As String, Path As String, i As Long, PSHeader As tPSHeader, _
  isf As InfoSpoolFile
50030  If NoFile = False Then
50040   If Len(Postscriptfile) > 0 Then
50050    If FileExists(Postscriptfile) = True Then
50060     isf = ReadInfoSpoolfile(Postscriptfile)
50070     If LenB(isf.REDMON_USER) > 0 Then
50080      UserName = isf.REDMON_USER
50090     End If
50100     If LenB(UserName) = 0 Then
50110      PSHeader = GetPSHeader(Postscriptfile)
50120      If Len(PSHeader.CreateFor.Comment) > 0 Then
50130       UserName = PSHeader.CreateFor.Comment
50140      End If
50150     End If
50160    End If
50170   End If
50180  End If
50190  If LenB(UserName) = 0 Then
50200 ' If LenB(UserName) = 0 Or UCase$(UserName) = UCase$(App.EXEName) Then
50210   UserName = Environ$("Redmon_User")
50220  End If
50230  If Len(UserName) = 0 Then
50240   UserName = GetUsername
50250  End If
50260  GetDocUsername = UserName
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

Public Function GetDocDate(Optional StandardDate As String = "", Optional StandardDateformat As String = "", Optional UseThisdate As String = "") As String
 On Error Resume Next
 Dim tstr As String, DateFormat As String, Usingdate As String

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

 tstr = Format$(Usingdate, DateFormat)
 If LenB(tstr) = 0 Then
  tstr = Usingdate
 End If
 GetDocDate = tstr
End Function

Public Function FormatPrintDocumentDate(tDate As String) As String
 Dim tstr As Long, m As Long, d As Long, Y As Long
 On Error Resume Next
 FormatPrintDocumentDate = tDate
 If InStr(tDate, "/") > 0 Then
   tstr = Mid(tDate, 1, InStr(tDate, "/") - 1)
   If IsNumeric(tstr) = False Then
    Exit Function
   End If
   m = CLng(tstr)
   If m < 1 Or m > 12 Then
    Exit Function
   End If
  Else
   Exit Function
 End If
 If InStr(InStr(tDate, "/") + 1, tDate, "/") > 0 Then
   tstr = Mid(tDate, InStr(tDate, "/") + 1, InStr(InStr(tDate, "/") + 1, tDate, "/") - (InStr(tDate, "/") + 1))
   If IsNumeric(tstr) = False Then
    Exit Function
   End If
   d = CLng(tstr)
   If d < 1 Or d > 31 Then
    Exit Function
   End If
  Else
   Exit Function
 End If
 If InStr(tDate, " ") > 0 Then
   tstr = Mid(tDate, InStr(InStr(tDate, "/") + 1, tDate, "/") + 1, InStr(tDate, " ") - (InStr(InStr(tDate, "/") + 1, tDate, "/") + 1))
   If IsNumeric(tstr) = False Then
    Exit Function
   End If
   Y = CLng(tstr)
  Else
   Exit Function
 End If
 FormatPrintDocumentDate = CStr(DateSerial(Y, m, d)) + Mid(tDate, InStr(tDate, " "))
End Function

Public Function EncodeChars(ByVal Str1 As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, tstr As String
50020  Str1 = Replace(Str1, "\", "\\")
50030  Str1 = Replace(Str1, "{", "\{")
50040  Str1 = Replace(Str1, "}", "\}")
50050  Str1 = Replace(Str1, "[", "\[")
50060  Str1 = Replace(Str1, "]", "\]")
50070  Str1 = Replace(Str1, "(", "\(")
50080  Str1 = Replace(Str1, ")", "\)")
50090  For i = 1 To Len(Str1)
50100   If Asc(Mid(Str1, i, 1)) > 127 Then
50110     tstr = tstr & "\" & CStr(Oct(Asc(Mid(Str1, i, 1))))
50120    Else
50130     tstr = tstr & Mid(Str1, i, 1)
50140   End If
50150  Next i
50160  EncodeChars = tstr
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
50400         IFIsPS = True
50410        Else
50420         MsgBox LanguageStrings.MessagesMsg06
50430       End If
50440       DoEvents
50450     End If
50460    Else
50470     If LenB(InputFilename) > 0 Then
50480      MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & vbCrLf & _
      "InputFile -IF" & vbCrLf & ">" & InputFilename & "<", vbExclamation + vbOKOnly
50500     End If
50510   End If
50520  End If
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
