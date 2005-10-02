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
50010  Dim tStr As String, psH As tPSHeader, isf As InfoSpoolFile
50020  isf = ReadInfoSpoolfile(Filename)
50030  tStr = isf.REDMON_DOCNAME
50040  If LenB(tStr) = 0 Then
50050   psH = GetPSHeader(Filename)
50060   tStr = psH.Title.Comment
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

Public Function GetAutosaveFilename(PostscriptFile As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Filename As String, Pathname As String
50020
50030  If Options.UseAutosaveDirectory = 1 Then
50040    Pathname = GetSubstFilename2(Options.AutosaveDirectory, False, , PostscriptFile)
50050   Else
50060    Pathname = GetSubstFilename2(Options.LastSaveDirectory, False)
50070  End If
50080
50090  Filename = GetSubstFilename(PostscriptFile, Options.AutosaveFilename)
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
50320   Case 8: 'TXT
50330    Filename = Filename & ".txt"
50340  End Select
50350  GetAutosaveFilename = CompletePath(GetSubstFilename2(Pathname, False)) & ReplaceForbiddenChars(Filename)
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
 Optional WithoutAuthor As Boolean = False, Optional Preview As Boolean = False) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim PSHeader As tPSHeader, Author As String, ClientComputer As String, _
  Title As String, UserName As String, Computername As String, i As Long, _
  DateTime As String, Filename As String, tStr As String, tList() As String, _
  Subst() As String, UserProfilPath As String, MyFiles As String, _
  MyDesktop As String, Path As String, isf As InfoSpoolFile
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
50380    PSHeader.Title.Comment = "'Preview Title'"
50390    Author = "'Preview Author'"
50400    ClientComputer = "'Preview ClientComputer'"
50410  End If
50420
50430  If Options.FilenameSubstitutionsOnlyInTitle = 1 Then
50440   tList = Split(Options.FilenameSubstitutions, "\")
50450   Title = PSHeader.Title.Comment
50460   If UBound(tList) >= 0 Then
50470    For i = 0 To UBound(tList)
50480     Subst = Split(tList(i), "|")
50490     If UBound(Subst) = 0 Then
50500       tStr = ""
50510      Else
50520       tStr = Subst(1)
50530     End If
50540     Title = Replace(Title, Subst(0), tStr, , , vbTextCompare)
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
50650  Filename = TokenFilename
50660  Filename = Replace(Filename, "<DateTime>", DateTime, , , vbTextCompare)
50670  Filename = Replace(Filename, "<Computername>", Computername, , , vbTextCompare)
50680
50690  Filename = Replace(Filename, "<ClientComputer>", ClientComputer, , , vbTextCompare)
50700  Filename = Replace(Filename, "<Username>", UserName, , , vbTextCompare)
50710  Filename = Replace(Filename, "<Title>", Title, , , vbTextCompare)
50720  If WithoutAuthor = False Then
50730   Filename = Replace(Filename, "<Author>", Author, , , vbTextCompare)
50740  End If
50750  Filename = Replace(Filename, "<MyFiles>", CompletePath(MyFiles), , , vbTextCompare)
50760  Filename = Replace(Filename, "<MyDesktop>", CompletePath(MyDesktop), , , vbTextCompare)
50770
50780  tStr = "DOCNAME"
50790  If Preview = True Then
50800    Filename = Replace(Filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
50810   Else
50820    Filename = Replace(Filename, "<REDMON_" & tStr & ">", isf.REDMON_DOCNAME, , , vbTextCompare)
50830  End If
50840  tStr = "JOB"
50850  If Preview = True Then
50860    Filename = Replace(Filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
50870   Else
50880    Filename = Replace(Filename, "<REDMON_" & tStr & ">", isf.REDMON_JOB, , , vbTextCompare)
50890  End If
50900  tStr = "MACHINE"
50910  If Preview = True Then
50920    Filename = Replace(Filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
50930   Else
50940    Filename = Replace(Filename, "<REDMON_" & tStr & ">", isf.REDMON_MACHINE, , , vbTextCompare)
50950  End If
50960  tStr = "PORT"
50970  If Preview = True Then
50980    Filename = Replace(Filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
50990   Else
51000    Filename = Replace(Filename, "<REDMON_" & tStr & ">", isf.REDMON_PORT, , , vbTextCompare)
51010  End If
51020  tStr = "PRINTER"
51030  If Preview = True Then
51040    Filename = Replace(Filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
51050   Else
51060    Filename = Replace(Filename, "<REDMON_" & tStr & ">", isf.REDMON_PRINTER, , , vbTextCompare)
51070  End If
51080  tStr = "SESSIONID"
51090  If Preview = True Then
51100    Filename = Replace(Filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
51110   Else
51120    Filename = Replace(Filename, "<REDMON_" & tStr & ">", isf.REDMON_SESSIONID, , , vbTextCompare)
51130  End If
51140  tStr = "USER"
51150  If Preview = True Then
51160    Filename = Replace(Filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
51170   Else
51180    Filename = Replace(Filename, "<REDMON_" & tStr & ">", isf.REDMON_USER, , , vbTextCompare)
51190  End If
51200
51210  If Options.FilenameSubstitutionsOnlyInTitle = 0 Then
51220   tList = Split(Options.FilenameSubstitutions, "\")
51230   If UBound(tList) >= 0 Then
51240    For i = 0 To UBound(tList)
51250     Subst = Split(tList(i), "|")
51260     If UBound(Subst) = 0 Then
51270       tStr = ""
51280      Else
51290       tStr = Subst(1)
51300     End If
51310     Filename = Replace(Filename, Subst(0), tStr, , , vbTextCompare)
51320    Next i
51330   End If
51340  End If
51350  If Options.RemoveSpaces = 1 Then
51360   Filename = Trim$(Filename)
51370  End If
51380  GetSubstFilename = Filename
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
50020  If NoPSCheck Or Options.NoPSCheck = 1 Then
50030   IsPostscriptFile = True
50040   Exit Function
50050  End If
50060  IsPostscriptFile = False
50070  lHeader = GetPSHeader(Filename)
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
50080    tStr = Format(TimeSerial(0, tzi.DaylightToGMT, 0), "hh'mm'")
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

Public Function EncodeChars(ByVal Str1 As String) As String
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
50160  EncodeChars = tStr
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
50030  If LenB(InputFilename) = 0 Then
50040   Exit Sub
50050  End If
50060  If FileExists(InputFilename) = False Then
50070   If LenB(InputFilename) > 0 Then
50080    MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & vbCrLf & _
    "InputFile -IF" & vbCrLf & ">" & InputFilename & "<", vbExclamation + vbOKOnly
50100   End If
50110   Exit Sub
50120  End If
50130
50140  If LenB(OutputFilename) > 0 Then
50150    If IsPostscriptFile(InputFilename) = True Then
50160     If GsDllLoaded = 0 Then
50170      Exit Sub
50180     End If
50190     SplitPath OutputFilename, , , , , Ext
50200     GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
50210     If GsDllLoaded = 0 Then
50220      MsgBox LanguageStrings.MessagesMsg08
50230     End If
50241     Select Case UCase$(Ext)
           Case "PDF"
50260       CallGScript InputFilename, OutputFilename, Options, PDFWriter
50270      Case "PNG"
50280       CallGScript InputFilename, OutputFilename, Options, PNGWriter
50290      Case "JPG"
50300       CallGScript InputFilename, OutputFilename, Options, JPEGWriter
50310      Case "BMP"
50320       CallGScript InputFilename, OutputFilename, Options, BMPWriter
50330      Case "PCX"
50340       CallGScript InputFilename, OutputFilename, Options, PCXWriter
50350      Case "TIF"
50360       CallGScript InputFilename, OutputFilename, Options, TIFFWriter
50370      Case "PS"
50380       CallGScript InputFilename, OutputFilename, Options, PSWriter
50390      Case "EPS"
50400       CallGScript InputFilename, OutputFilename, Options, EPSWriter
50410      Case "TXT"
50420       CallGScript InputFilename, OutputFilename, Options, TXTWriter
50430     End Select
50440    End If
50450 '   If GsDllLoaded <> 0 Then
50460 '    UnloadDLLComplete GsDllLoaded
50470 '   End If
50480    ConvertedOutputFilename = OutputFilename
50490    ReadyConverting = True
50500    Exit Sub
50510   Else
50520    If FileExists(InputFilename) = True Then
50530     If IsPostscriptFile(InputFilename) = True Then
50540       IFIsPS = True
50550      Else
50560       MsgBox LanguageStrings.MessagesMsg06 & vbCrLf & vbCrLf & InputFilename
50570     End If
50580    End If
50590  End If
50600  DoEvents
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
