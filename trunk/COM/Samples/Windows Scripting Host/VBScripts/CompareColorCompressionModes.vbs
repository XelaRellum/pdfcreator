' CompareColorCompressionModes script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.sf.net/projects/pdfcreator
' Version: 2.0.0.0
' Date: April, 12. 2005
' Author: Frank Heindörfer
' Comments: This script shows how you decrease the size of a pdf-file
'           using diffent compression modes.

Option Explicit

Const Title = "Sample - Compression", _
      HeadStatus1 = "PDFCreator", _
      HeadStatus2 = "At first we create an image. This can take a few of minutes.", _
      HeadStatus3 = "Now we convert the files as a pdf-file: ", _
      HeadStatus4 = "Saving the image as postcript-file"

Const ColorSteps = 256

'Const wi = 100, hei = 100       ' Small and simple image, fast calculation
Const wi = 1200, hei = 1200     ' Big and exacter image, slow calculation

Const maxTime = 10    ' in seconds (PDFCreator: Max. time for calculation)
Const sleepTime = 250 ' in milliseconds (PDFCreator: Wait to complete orders

Dim im(), ColorTable(), PSFileName, oIE
Dim WshShell, fso, PDFCreator, DefaultPrinter, ReadyState, opath, _
 ScriptBaseName, AppTitle, res(), sf

InitScript

CreateColorTableRandom
CreateFractal
CreatePSFile
CheckCompressions
ShowResults
CloseScript

Public Sub ShowResults
 Dim tStr, c, i, f
 tStr = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01//EN"""
 tStr = tStr & vbcrlf & "<html>"
 tStr = tStr & vbcrlf & "<head>"
 tStr = tStr & vbcrlf & "<title>" & HeadStatus1 & "</title>"
 tStr = tStr & vbcrlf & "</head>"
 tStr = tStr & vbcrlf & "<body>"
 tStr = "<h1>" & HeadStatus1 & "</h1><strong>CompareColorCompressionModes results</strong><br><br>"
 tStr = tStr & "<table border=""1"" cellpadding=""4"" style=""border-collapse:collapse;empty-cells:show;"">"
 tStr = tStr & vbcrlf & "<tr>"
 tStr = tStr & vbcrlf & "<th>Filename</th>"
 tStr = tStr & vbcrlf & "<th>Size</th>"
 tStr = tStr & vbcrlf & "<th>Link</th>"
 tStr = tStr & vbcrlf & "</tr>"
 c = 0
 For i = 0 to Ubound(res,2)
  tStr = tStr & vbcrlf & "<tr>"
  tStr = tStr & vbcrlf & "<td>       " & ReplaceForbiddenChars(res(0, i)) & "</td>"
  tStr = tStr & vbcrlf & "<td align=""right"">       " & ReplaceForbiddenChars(res(1, i)) & " Bytes</td>"
  tStr = tStr & vbcrlf & "<td><a href=""file:///" & Replace(ReplaceForbiddenChars(res(2, i)),"\", "/") & """>" & ReplaceForbiddenChars(res(2, i)) & "</a></td>"
  tStr = tStr & vbcrlf & "</tr>"
 Next
 tStr = tStr & vbcrlf & "</table>"

 Set f = fso.OpenTextFile("results.htm", 2, True)
 f.WriteLine tStr
 f.Close
 Set WshShell = WScript.CreateObject("WScript.Shell")
 WshShell.Run "results.htm"
End Sub


Private Function ReplaceForbiddenChars(value)
 Dim tStr
 tStr=Replace(value, "&", "&amp;")
 tStr=Replace(tStr, "<", "&lt;")
 tStr=Replace(tStr, ">", "&gt;")
 tStr=Replace(tStr, """", "&quot;")
 ReplaceForbiddenChars = tStr
End Function

Public Sub CheckCompressions
 Dim fname

 fname = "Compression test - Automatic"
 CheckCompression fname, 0, 0, 0, 0, 0, 0

 PDFCreator.cOption("PDFCompressionColorResample") = 0

 fname = "Compression test - ZIP No DownSample"
 CheckCompression fname, 6, 0, 0, 0, 0, 0

 PDFCreator.cOption("PDFCompressionColorResample") = 1
 PDFCreator.cOption("PDFCompressionColorResampleChoice") = 0
 PDFCreator.cOption("PDFCompressionColorResolution") = 64
 fname = "Compression test - ZIP DownSample 64"
 CheckCompression fname, 6, 0, 0, 0, 0, 0

 PDFCreator.cOption("PDFCompressionColorResolution") = 32
 fname = "Compression test - ZIP DownSample 32"
 CheckCompression fname, 6, 0, 0, 0, 0, 0

 PDFCreator.cOption("PDFCompressionColorResolution") = 16
 fname = "Compression test - ZIP DownSample 16"
 CheckCompression fname, 6, 0, 0, 0, 0, 0

 PDFCreator.cOption("PDFCompressionColorResolution") = 8
 fname = "Compression test - ZIP DownSample 8"
 CheckCompression fname, 6, 0, 0, 0, 0, 0

 PDFCreator.cOption("PDFCompressionColorResample") = 0
 fname = "Compression test - JPEG-Minimum Standard"
 CheckCompression fname, 5, PDFCreator.cStandardOption("PDFCompressionColorCompressionJPEGMinimumFactor"), 0, 0, 0, 0

 fname = "Compression test - JPEG-Low Standard"
 CheckCompression fname, 4, 0, PDFCreator.cStandardOption("PDFCompressionColorCompressionJPEGLowFactor"), 0, 0, 0

 fname = "Compression test - JPEG-Medium Standard"
 CheckCompression fname, 3, 0, 0, PDFCreator.cStandardOption("PDFCompressionColorCompressionJPEGMediumFactor"), 0, 0

 fname = "Compression test - JPEG-High Standard"
 CheckCompression fname, 2, 0, 0, 0, PDFCreator.cStandardOption("PDFCompressionColorCompressionJPEGHighFactor"), 0

 fname = "Compression test - JPEG-Maximum Standard"
 CheckCompression fname, 1, 0, 0, 0, 0, PDFCreator.cStandardOption("PDFCompressionColorCompressionJPEGMaximumFactor")

 fname = "Compression test - JPEG-Max 5"
 CheckCompression fname, 1, 0, 0, 0, 0, 5

 fname = "Compression test - JPEG-Max 10"
 CheckCompression fname, 1, 0, 0, 0, 0, 10

 fname = "Compression test - JPEG-Max 15"
 CheckCompression fname, 1, 0, 0, 0, 0, 15

 PDFCreator.cOption("PDFCompressionColorResample") = 1
 PDFCreator.cOption("PDFCompressionColorResampleChoice") = 0

 fname = "Compression test - JPEG-Max 15 DownSample 32"
 PDFCreator.cOption("PDFCompressionColorResolution") = 32
 CheckCompression fname, 1, 0, 0, 0, 0, 15

 fname = "Compression test - JPEG-Max 15 DownSample 16"
 PDFCreator.cOption("PDFCompressionColorResolution") = 16
 CheckCompression fname, 1, 0, 0, 0, 0, 15

 fname = "Compression test - JPEG-Max 15 DownSample 4"
 PDFCreator.cOption("PDFCompressionColorResolution") = 4
 CheckCompression fname, 1, 0, 0, 0, 0, 15
End Sub

Private Sub CheckCompression(Filename, PDFCompressionColorCompressionChoice, _
 PDFCompressionColorCompressionJPEGMinimumFactor, _
 PDFCompressionColorCompressionJPEGLowFactor, _
 PDFCompressionColorCompressionJPEGMediumFactor, _
 PDFCompressionColorCompressionJPEGHighFactor, _
 PDFCompressionColorCompressionJPEGMaximumFactor)
 Dim c, f, inFname, outFname, tStr
 If IsDim(res) Then
   ReDim Preserve res(3, Ubound(res,2)+1)
  Else
   ReDim res(3, 0)
   oIE.Document.Body.InnerHTML = "<h1>" & HeadStatus1 & "</h1><strong>" & HeadStatus3 & "</strong><br><br>"
 End If

 tStr = oIE.Document.Body.InnerHTML
 oIE.Document.Body.InnerHTML = tStr & "Convert: <em>" & Filename & "</em> ...<br>"

 inFname = CompletePath(fso.GetParentFolderName(Wscript.ScriptFullname)) & PSFileName
 outFname = CompletePath(opath) & Filename & ".pdf"

 If (fso.FileExists(outFname)) Then
  On Error Resume Next
  fso.Deletefile(outFname)
  If err.Number <> 0 Then
   MsgBox """" & outFname & """ is in use.", vbCritical + vbSystemModal, AppTitle
  End If
  On Error Goto 0
 End If

 With PDFCreator
  .cOption("PDFCompressionColorCompressionChoice") = PDFCompressionColorCompressionChoice
  .cOption("PDFCompressionColorCompressionJPEGMinimumFactor") = PDFCompressionColorCompressionJPEGMinimumFactor
  .cOption("PDFCompressionColorCompressionJPEGLowFactor") = PDFCompressionColorCompressionJPEGLowFactor
  .cOption("PDFCompressionColorCompressionJPEGMediumFactor") = PDFCompressionColorCompressionJPEGMediumFactor
  .cOption("PDFCompressionColorCompressionJPEGHighFactor") = PDFCompressionColorCompressionJPEGHighFactor
  .cOption("PDFCompressionColorCompressionJPEGMaximumFactor") = PDFCompressionColorCompressionJPEGMaximumFactor
  ReadyState = 0
  .cConvertPostscriptfile inFname, outFname
 End With

 c = 0
 Do While (ReadyState = 0) and (c < (maxTime * 1000 / sleepTime))
  c = c + 1
  Wscript.Sleep sleepTime
 Loop

 If ReadyState = 0 then
  MsgBox "Max. time is up!", vbExclamation + vbSystemModal, AppTitle
  Wscript.quit
 End If

 Set f = fso.GetFile(outFname)
 res(0, Ubound(res,2)) = fso.GetBaseName(outFname)
 res(1, Ubound(res,2)) = f.Size
 res(2, Ubound(res,2)) = outFname
 Wscript.Sleep 250

 oIE.Document.Body.InnerHTML = tStr & "Convert: <em>" & Filename & "</em> - ready<br>"
end Sub

Private Sub CreatePSFile
 CreatePSHeader
 CreatePSBodyFractal
' CreatePSBodyColorTable
 CreatePSFooter
End Sub

Private Sub CreateFractal
' Const X1 = -2, Y1 = 2, X2 = 2, Y2 = -2
 Const X1 = -1.52, Y1 = 0.05, X2 = -1.29, Y2 = -0.18

 Dim StepX, StepY, X, Y, Cx, Cy, Zx, Zy, Color, TempX, TempY

 StepX = Abs(X2 - X1) / wi
 StepY = Abs(Y2 - Y1) / hei

 For X = 0 To wi-1
  For Y = 0 To hei-1
   Cx = X1 + X * StepX
   Cy = Y2 + Y * StepY
   Zx = 0.0: Zy = 0.0
   Color = 0.0
   While (Not (Zx * Zx + Zy * Zy > 4)) And Color < (ColorSteps-1)
    TempX = Zx
    TempY = Zy
    Zx = TempX * TempX - TempY * TempY + Cx
    Zy = 2.0 * TempX * TempY + Cy
    Color = Color + 1.0
   Wend
   im(X, Y, 0) = ColorTable(Color, 0)
   im(X, Y, 1) = ColorTable(Color, 1)
   im(X, Y, 2) = ColorTable(Color, 2)
  Next
  oIE.Document.Body.InnerHTML = "<h1>" & HeadStatus1 & "</h1>" & _
   "<strong>" & HeadStatus2 & "</strong><br><br>Progress: " & _
   X + 1  & " [" & wi & "] = " & FormatNumber(100 * (X + 1) / wi, 2) & "%"
 Next
End Sub

Private Sub CreateColorTableRandom
 Const Steps = 10
 Dim sw, i, j
 sw = Fix(ColorSteps/(Steps-1)) + 1
 ReDim CS(Steps - 1, 2)
 For i = 0 To (Steps - 1)
  For j = 0 To 2
   CS(i, j) = Fix(RND*256)
  Next
 Next
 CS(0,0)=196: CS(0,1)=0:   CS(0,2)=0
 CS(1,0)=196: CS(1,1)=196: CS(1,2)=0
 CS(2,0)=196: CS(2,1)=0:   CS(2,2)=196
 CS(Steps-1,0)=255: CS(Steps-1,1)=0: CS(Steps-1,2)=0

 For i=0 to ColorSteps - 1
  For j=0 to 2
   ColorTable(i,j) = CS(Fix(i/sw),j) + (CS(Fix(i/sw)+1,j) - CS(Fix(i/sw),j))*(i - (i\sw) * sw)/sw
  Next
 Next
End Sub

Private Sub CreatePSHeader
 Const ForWriting = 2
 Dim f
 Set f = fso.OpenTextFile(PSFileName, ForWriting, True)
 f.WriteLine "%!PS-Adobe-1.0"
 f.WriteLine "%%Creator: PDFCreator"
 f.WriteLine "%%For: PDFCreator"
 f.WriteLine "%%Title: Testcompression"
 f.WriteLine "%%Pages: 1"
 f.WriteLine "%%PageOrder: Ascend"
 f.WriteLine "%%DocumentMedia: A4 595 842 0 () ()"
 f.WriteLine "%%Orientation: Portrait"
 f.WriteLine "%%BoundingBox: 0 0 596 842"
 f.WriteLine "%%EndComments"
 f.WriteLine "%%EndProlog"
 f.WriteLine "%%Page: 1 1"
 f.Close
End Sub

Private Sub CreatePSBodyFractal
 Const ForAppending = 8
 Dim f, i, j, k, tStr
 Set f = fso.OpenTextFile(PSFileName, ForAppending, True)
 f.WriteLine "gsave"
 f.WriteLine Round((596-sf*wi)/2) & " "  & Round((842-sf*hei)/2) & " translate"
 f.WriteLine sf*wi & " " & sf*hei & " scale"
 f.WriteLine wi & " " & hei & " 8 [" & wi & " 0 0 -" & hei & " 0 " & hei & "]"
 f.WriteLine "{currentfile 3 " & wi & " mul string readhexstring pop} bind"
 f.WriteLine "false 3"
 f.WriteLine "colorimage"
 for j=0 to hei-1
  tStr = ""
  for i=0 to wi-1
   for k =0 to 2
    if len(hex(im(i,j,k)))=1 then
      tStr = tStr & "0" & hex(im(i,j,k))
     else
      tStr = tStr & hex(im(i,j,k))
    end if
   next
  next
  f.WriteLine tStr
  oIE.Document.Body.InnerHTML = "<h1>" & HeadStatus1 & "</h1>" & _
   "<strong>" & HeadStatus4 & "</strong><br><br>Progress: " & _
   j + 1  & " [" & hei & "] = " & FormatNumber(100 * (j + 1) / hei, 2) & "%"
 next
 f.WriteLine "grestore"
 f.Close
End Sub

Private Sub CreatePSBodyColorTable
 Const ForAppending = 8
 Dim f, i, j, k, tStr
 f.WriteLine "gsave"
 f.WriteLine "10 " & 842-50 & " translate"
 f.WriteLine 2*ColorSteps & " 40  scale"
 f.WriteLine ColorSteps & " 20 8 [" & ColorSteps & " 0 0 -20 0  20]"
 f.WriteLine "{currentfile 3 " & ColorSteps & " mul string readhexstring pop} bind"
 f.WriteLine "false 3"
 f.WriteLine "colorimage"
 for j=0 to 20-1
  tStr = ""
  for i=0 to ColorSteps-1
   for k =0 to 2
    if len(hex(ColorTable(i,k)))=1 then
      tStr = tStr & "0" & hex(ColorTable(i,k))
     else
      tStr = tStr & hex(ColorTable(i,k))
    end if
   next
  next
  f.WriteLine tStr
 next
 f.WriteLine "grestore"
 f.Close
End Sub

Private Sub CreatePSFooter
 Const ForAppending = 8
 Dim f, tStr
 Set f = fso.OpenTextFile(PSFileName, ForAppending, True)
 f.WriteLine "[ /PageMode /UseNone /Page 1 /View [/FitB] /DOCVIEW pdfmark"
 f.WriteLine "showpage"
 f.WriteLine "%%Trailer"
 f.WriteLine "%%EOF"
 f.Close
End Sub

Private Sub InitScript
 Redim im(wi-1, hei-1, 2)
 ReDim ColorTable(ColorSteps -1 , 2)

 Set fso = CreateObject("Scripting.FileSystemObject")
 ScriptBaseName = fso.GetBaseName(Wscript.ScriptFullname)
 AppTitle = "PDFCreator - " & ScriptBaseName
 If CDbl(Replace(WScript.Version,".",",")) < 5.1 then
  MsgBox "You need the ""Windows Scripting Host version 5.1"" or greater!", vbCritical + vbSystemModal, AppTitle
  Wscript.Quit
 End if

 RND -1
 Randomize 60
 PSFileName = Title & ".ps"
 sf = 400/wi

 Set oIE = WScript.CreateObject("InternetExplorer.Application", "IE_")

 With oIE
  .height = 400
  .width = 500
  .menubar = false
  .toolbar = false
  .statusbar = false
  .resizable = false
  .visible = true
  .Navigate "about:blank"
 End With
 Do Until oIE.ReadyState=4
  WScript.Sleep 50
 Loop
 oIE.Navigate ("javascript:'<TITLE>PDFCreator - CompareColorCompressionModes</TITLE>'")

 opath = fso.GetParentFolderName(Wscript.ScriptFullname)
 Set PDFCreator = Wscript.CreateObject("PDFCreator.clsPDFCreator", "PDFCreator_")
 PDFCreator.cStart "/NoProcessingAtStartup"
 With PDFCreator
  .cClearcache
  .cClearLogfile
  .cOption("Logging") = 1
  .cOption("PDFCompressionColorResample") = 0
 End With
End Sub

Private Function CompletePath(Path)
 If Right(Path, 1) <> "\" Then
   CompletePath = Path & "\"
  Else
   CompletePath = Path
 End If
End Function

Private Function IsDim(Arr)
 Dim nLBound
 On Error Resume Next
 nLBound = LBound(Arr)
 IsDim = Not CBool(Err.Number)
End Function

Private Sub CloseScript
 With PDFCreator
  .cClearcache
  WScript.Sleep 200
  .cClose
 End With
 fso.DeleteFile(PSFilename)
 With oIE
  .Visible = 0
  .Quit
 End With
End Sub

'--- PDFCreator events ---

Public Sub PDFCreator_eReady()
 ReadyState = 1
End Sub

Public Sub PDFCreator_eError()
 MsgBox "An error is occured!" & vbcrlf & vbcrlf & _
  "Error [" & PDFCreator.cErrorDetail("Number") & "]: " & PDFcreator.cErrorDetail("Description"), vbCritical + vbSystemModal, AppTitle
 Wscript.Quit
End Sub