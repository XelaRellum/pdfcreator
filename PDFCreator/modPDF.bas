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
50040    Pathname = CompletePath(Trim$(Options.AutosaveDirectory))
50050   Else
50060    Pathname = CompletePath(Trim$(Options.LastSaveDirectory))
50070  End If
50080
50090  Filename = GetSubstFilename(Postscriptfile, Options.AutosaveFilename)
50100
50111  Select Case Options.AutosaveFormat
              Case 0: 'PDF
50130    Filename = Filename & ".pdf"
50140   Case 1: 'PNG
50150    Filename = Filename & ".png"
50160   Case 2: 'JPEG
50170    Filename = Filename & ".jpg"
50180   Case 3: 'BMP
50190    Filename = Filename & ".bmp"
50200   Case 4: 'PCX
50210    Filename = Filename & ".pcx"
50220   Case 5: 'TIFF
50230    Filename = Filename & ".tif"
50240   Case 6: 'PS
50250    Filename = Filename & ".ps"
50260   Case 7: 'EPS
50270    Filename = Filename & ".eps"
50280  End Select
50290  GetAutosaveFilename = CompletePath(GetSubstFilename(Postscriptfile, Pathname)) & ReplaceForbiddenChars(Filename)
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
  Subst() As String, UserProfilPath As String, MyFiles As String, Path As String
50050
50060  If Len(TokenFilename) = 0 Then
50070   Exit Function
50080  End If
50090
50100  DateTime = GetDocDate("", Options.StandardDateformat, CStr(Now))
50110  If Preview = False Then
50120    If FileExists(Postscriptfile) = True Then
50130     PSHeader = GetPSHeader(Postscriptfile)
50140    End If
50150    If Options.UseStandardAuthor = 1 Then
50160      Author = Options.StandardAuthor
50170     Else
50180      Author = PSHeader.CreateFor.Comment
50190    End If
50200   Else
50210    PSHeader.Title.Comment = "'Preview Title'"
50220    Author = "'Preview Author'"
50230  End If
50240
50250  If Options.FilenameSubstitutionsOnlyInTitle = 1 Then
50260   tList = Split(Options.FilenameSubstitutions, "\")
50270   Title = PSHeader.Title.Comment
50280   If UBound(tList) >= 0 Then
50290    For i = 0 To UBound(tList)
50300     Subst = Split(tList(i), "|")
50310     If UBound(Subst) = 0 Then
50320       tStr = ""
50330      Else
50340       tStr = Subst(1)
50350     End If
50360     Title = Replace(Title, Subst(0), tStr, , , vbTextCompare)
50370    Next i
50380   End If
50390  End If
50400
50410  UserName = GetDocUsername(Postscriptfile, Preview)
50420
50430  Computername = GetComputerName
50440  MyFiles = GetMyFiles
50450
50460  Filename = TokenFilename
50470  Filename = Replace(Filename, "<DateTime>", DateTime, , , vbTextCompare)
50480  Filename = Replace(Filename, "<Computername>", Computername, , , vbTextCompare)
50490  Filename = Replace(Filename, "<Username>", UserName, , , vbTextCompare)
50500  Filename = Replace(Filename, "<Title>", Title, , , vbTextCompare)
50510  If WithoutAuthor = False Then
50520   Filename = Replace(Filename, "<Author>", Author, , , vbTextCompare)
50530  End If
50540  Filename = Replace(Filename, "<MyFiles>", MyFiles, , , vbTextCompare)
50550
50560  If Options.FilenameSubstitutionsOnlyInTitle = 0 Then
50570   tList = Split(Options.FilenameSubstitutions, "\")
50580   If UBound(tList) >= 0 Then
50590    For i = 0 To UBound(tList)
50600     Subst = Split(tList(i), "|")
50610     If UBound(Subst) = 0 Then
50620       tStr = ""
50630      Else
50640       tStr = Subst(1)
50650     End If
50660     Filename = Replace(Filename, Subst(0), tStr, , , vbTextCompare)
50670    Next i
50680   End If
50690  End If
50700  If Options.RemoveSpaces = 1 Then
50710   Filename = Trim$(Filename)
50720  End If
50730  GetSubstFilename = Filename
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
50010  Dim fn As Long, DocInfoStr As String
50020  If FileExists(PSFile) = True Then
50030   DocInfoStr = Chr$(13) & "%!PS-Adobe-3.0 EPSF-3.0"
50040   DocInfoStr = DocInfoStr & Chr$(13) & "%%BoundingBox: 0 0 72 72"
50050   DocInfoStr = DocInfoStr & Chr$(13) & "%%EndProlog"
50060   DocInfoStr = DocInfoStr & Chr$(13) & "/pdfmark where {pop} {userdict /pdfmark /cleartomark load put} ifelse"
50070   DocInfoStr = DocInfoStr & Chr$(13) & "["
50080   With PDFDocInfo
50090    DocInfoStr = DocInfoStr & "/Author (" & .Author & ")"
50100    If LenB(Trim$(.CreationDate)) = 0 Then
50110      DocInfoStr = DocInfoStr & Chr$(13) & "/CreationDate ()"
50120     Else
50130      DocInfoStr = DocInfoStr & Chr$(13) & "/CreationDate (" & .CreationDate & ")"
50140    End If
50150    DocInfoStr = DocInfoStr & Chr$(13) & "/Creator (" & .Creator & ")"
50160    DocInfoStr = DocInfoStr & Chr$(13) & "/Keywords (" & .Keywords & ")"
50170    If LenB(Trim$(.ModifyDate)) = 0 Then
50180      DocInfoStr = DocInfoStr & Chr$(13) & "/ModDate ()"
50190     Else
50200      DocInfoStr = DocInfoStr & Chr$(13) & "/ModDate (" & .ModifyDate & ")"
50210    End If
50220    DocInfoStr = DocInfoStr & Chr$(13) & "/Subject (" & .Subject & ")"
50230    DocInfoStr = DocInfoStr & Chr$(13) & "/Title (" & .Title & ")"
50240 '   DocInfoStr = DocInfoStr & Chr$(13) & "/Producer ()"
50250   End With
50260   DocInfoStr = DocInfoStr & Chr$(13) & "/DOCINFO pdfmark"
50270   DocInfoStr = DocInfoStr & Chr$(13) & "%%EOF"
50280   fn = FreeFile
50290   Open PSFile For Append As fn
50300   Print #fn, DocInfoStr;
50310   Close #fn
50320  End If
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

Public Function GetDocUsername(Postscriptfile As String, NoFile As Boolean) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim UserName As String, Path As String, i As Long, PSHeader As tPSHeader
50020  If NoFile = False Then
50030   If Len(Postscriptfile) > 0 Then
50040    If FileExists(Postscriptfile) = True Then
50050     PSHeader = GetPSHeader(Postscriptfile)
50060     SplitPath Postscriptfile, , Path
50070     Path = CompletePath(Path)
50080     If Len(Path) > 1 Then
50090      Path = Left(Path, Len(Path) - 1)
50100      i = InStrRev(Path, "\", , vbTextCompare)
50110      If i > 0 Then
50120       UserName = Mid$(Path, i + 1)
50130      End If
50140     End If
50150    End If
50160   End If
50170  End If
50180  If Len(UserName) = 0 Or UCase$(UserName) = UCase$(App.EXEName) Then
50190   UserName = Environ$("Redmon_User")
50200  End If
50210  If Len(UserName) = 0 Then
50220   UserName = PSHeader.CreateFor.Comment
50230  End If
50240  If Len(UserName) = 0 Then
50250   UserName = GetUsername
50260  End If
50270  GetDocUsername = UserName
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

