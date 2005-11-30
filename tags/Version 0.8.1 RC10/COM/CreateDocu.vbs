Option Explicit

Const HTMLFile = "COM interface - documentation.htm"

Dim fso, tStrf, i, ma1, ma2
Dim f,  s,  p,  e, erc, opc, inc
Dim fc, sc, pc, ec, ercc, opcc, incc

Set s = CreateObject("Scripting.Dictionary") ' Subroutines
Set f = CreateObject("Scripting.Dictionary") ' Functions
Set p = CreateObject("Scripting.Dictionary") ' Properties
Set e = CreateObject("Scripting.Dictionary") ' Events

Set erc = CreateObject("Scripting.Dictionary") ' clsPDFCreatorError
Set inc = CreateObject("Scripting.Dictionary") ' clsPDFCreatorInfoSpoolFile
Set opc = CreateObject("Scripting.Dictionary") ' clsPDFCreatorOptions

tStrf = Split(ReadContent("..\PDFCreator\clsPDFCreator.cls"), vbcrlf)

fc = 0: sc = 0: pc = 0
For i=LBound(tStrf) to Ubound(tStrf)
 If Instr(ucase(Trim(tStrf(i))),"PUBLIC FUNCTION") = 1 Then
  fc = fc +1
  f.add fc, trim(tStrf(i)) & "<br>"
 End If
 If Instr(ucase(Trim(tStrf(i))),"PUBLIC SUB") = 1 Then
  sc = sc +1
  s.add sc, trim(tStrf(i)) & "<br>"
 End If
 If Instr(ucase(Trim(tStrf(i))),"PUBLIC PROPERTY") = 1 Then
  pc = pc +1
  p.add pc, trim(tStrf(i)) & "<br>"
 End If
 If Instr(ucase(Trim(tStrf(i))),"PUBLIC EVENT") = 1 Then
  ec = ec +1
  e.add ec, trim(tStrf(i)) & "<br>"
 End If
Next

tStrf = Split(ReadContent("..\PDFCreator\clsPDFCreatorError.cls"), vbcrlf)
ercc = 0
For i=LBound(tStrf) to Ubound(tStrf)
 If Instr(ucase(Trim(tStrf(i))),"PUBLIC") = 1 Then
  ercc = ercc +1
  erc.add ercc, trim(tStrf(i)) & "<br>"
 End If
Next

tStrf = Split(ReadContent("..\PDFCreator\clsPDFCreatorInfoSpoolFile.cls"), vbcrlf)
incc = 0
For i=LBound(tStrf) to Ubound(tStrf)
 If Instr(ucase(Trim(tStrf(i))),"PUBLIC") = 1 Then
  incc = incc +1
  inc.add incc, trim(tStrf(i)) & "<br>"
 End If
Next

tStrf = Split(ReadContent("..\PDFCreator\clsPDFCreatorOptions.cls"), vbcrlf)
opcc = 0
For i=LBound(tStrf) to Ubound(tStrf)
 If Instr(ucase(Trim(tStrf(i))),"PUBLIC") = 1 Then
  opcc = opcc +1
  opc.add opcc, trim(tStrf(i)) & "<br>"
 End If
Next

Call CreateHTMLFile(HTMLFile, Header & CreateDocu & Footer)

Private Function ReadContent(Filename)
 Dim fso, f
 Set fso = CreateObject("Scripting.FileSystemObject")
 If fso.FileExists(Filename) Then
   Set f = fso.OpenTextFile(Filename, 1)
   ReadContent = f.ReadAll
   f.Close
  Else
   Msgbox "'" & Filename & "' doesn't exist!", vbCritical & vbSystemModal
   Wscript.Quit
 End if
End Function

Private Sub CreateHTMLFile(Filename, Content)
 Dim tf
 Set fso = CreateObject("Scripting.FileSystemObject")
 Set tf = fso.CreateTextFile(Filename, True)
 tf.Write Content
 tf.Close
End Sub

Private Function Header
 Dim tStr, title
 title = "PDFCreator COM-interface"
 tStr = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01//EN"""
 tStr = tStr & vbcrlf & "<html>"
 tStr = tStr & vbcrlf & "<head>"
 tStr = tStr & vbcrlf & "<title>" & title & "</title>"
 tStr = tStr & vbcrlf & "</head>"
 tStr = tStr & vbcrlf & "<body>"
 tStr = tStr & vbcrlf & "<h1>" & title & "</h1>"
 Header = tStr & vbcrlf
End Function

Private Function CreateDocu
 Dim a, i, tStr

 Set e = Sort(e,1,1)
 if e.count> 0 then
  a = e.Items
  tStr = tStr & "<p>" & vbcrlf
  tStr = tStr & "<table border=""1"" cellpadding=""4"" style=""border-collapse:collapse;empty-cells:show;"">" & vbcrlf
  tStr = tStr & "<tr><th>Events</th></tr><tr><td nowrap>" & vbcrlf
  For i = 0 To e.Count -1
   tStr = tStr & a(i) & vbcrlf
  Next
  tStr = tStr & "</td></tr></table></p>" & vbcrlf
 end if

 Set p = Sort(p,Len("Public Property Let c "),Len("Public Property L"))
 if p.count> 0 then
  a = p.Items
  tStr = tStr & "<p>" & vbcrlf
  tStr = tStr & "<table border=""1"" cellpadding=""4"" style=""border-collapse:collapse;empty-cells:show;"">" & vbcrlf
  tStr = tStr & "<tr><th>Properties</th></tr><tr><td nowrap>" & vbcrlf
  For i = 0 To p.Count -1
   tStr = tStr & a(i) & vbcrlf
  Next
  tStr = tStr & "</td></tr></table></p>" & vbcrlf
 end if

 Set f = Sort(f,1,1)
 if f.count> 0 then
  a = f.Items
  tStr = tStr & "<p>" & vbcrlf
  tStr = tStr & "<table border=""1"" cellpadding=""4"" style=""border-collapse:collapse;empty-cells:show;"">" & vbcrlf
  tStr = tStr & "<tr><th>Functions</th></tr><tr><td nowrap>" & vbcrlf
  For i = 0 To f.Count -1
   tStr = tStr & a(i) & vbcrlf
  Next
  tStr = tStr & "</td></tr></table></p>" & vbcrlf
 end if

 Set s = Sort(s,1,1)
 if s.count> 0 then
  a = s.Items
  tStr = tStr & "<p>" & vbcrlf
  tStr = tStr & "<table border=""1"" cellpadding=""4"" style=""border-collapse:collapse;empty-cells:show;"">" & vbcrlf
  tStr = tStr & "<tr><th>Subroutines</th></tr><tr><td nowrap>" & vbcrlf
  For i = 0 To s.Count -1
   if instr(a(i),"'")>0 then
     tStr = tStr & Replace(a(i),"'","'<span style=""color:green"">",1,1) & "</span>" & vbcrlf
    else
     tStr = tStr & a(i) & vbcrlf
   end if
  Next
  tStr = tStr & "</td></tr></table></p>" & vbcrlf
 end if

 Set erc = Sort(erc,1,1)
 if erc.count> 0 then
  a = erc.Items
  tStr = tStr & "<p>" & vbcrlf
  tStr = tStr & "<table border=""1"" cellpadding=""4"" style=""border-collapse:collapse;empty-cells:show;"">" & vbcrlf
  tStr = tStr & "<tr><th>clsPDFCreatorError</th></tr><tr><td nowrap>" & vbcrlf
  For i = 0 To erc.Count -1
   tStr = tStr & a(i) & vbcrlf
  Next
  tStr = tStr & "</td></tr></table></p>" & vbcrlf
 end if

 Set inc = Sort(inc,1,1)
 if inc.count> 0 then
  a = inc.Items
  tStr = tStr & "<p>" & vbcrlf
  tStr = tStr & "<table border=""1"" cellpadding=""4"" style=""border-collapse:collapse;empty-cells:show;"">" & vbcrlf
  tStr = tStr & "<tr><th>clsPDFCreatorInfoSpoolFile</th></tr><tr><td nowrap>" & vbcrlf
  For i = 0 To inc.Count -1
   tStr = tStr & a(i) & vbcrlf
  Next
  tStr = tStr & "</td></tr></table></p>"
 end if

 Set opc = Sort(opc,1,1)
 if opc.count> 0 then
  a = opc.Items
  tStr = tStr & "<p>" & vbcrlf
  tStr = tStr & "<table border=""1"" cellpadding=""4"" style=""border-collapse:collapse;empty-cells:show;"">" & vbcrlf
  tStr = tStr & "<tr><th>clsPDFCreatorOptions</th></tr><tr><td nowrap>" & vbcrlf
  For i = 0 To opc.Count -1
   tStr = tStr & a(i) & vbcrlf
  Next
  tStr = tStr & "</td></tr></table></p>" & vbcrlf
 end if

 CreateDocu = tStr
End Function

Private Function Footer
 Dim tStr
 tStr = "</body>"
 tStr = tStr & vbcrlf & "</html>"
 Footer = tStr
End Function

Private Function Sort(d,fs,ss)
 Dim i, j, a, t
 If d.Count<=1 then
  Set Sort = d
  Exit Function
 End If
 a = d.items
 For i=lbound(a) to ubound(a)
  If len(a(i))>=fs and len(a(i))>=ss then
   a(i)=Mid(a(i),fs,Instr(fs,a(i),"(")-(fs-1)) & Mid(a(i),ss,1) & a(i)
'   msgbox a(i)
'   wscript.quit
  End If
 Next
 For i=lbound(a) to ubound(a)-1
  For j=i+1 to ubound(a)
   if a(i) >a(j) then
    t=a(i)
    a(i)=a(j)
    a(j)=t
   end if
  next
 next
 d.RemoveAll
 if fs=1 and ss=1 then
'  msgbox ubound(a) & vbcrlf & a(ubound(a))

 end if

 For i=lbound(a) to ubound(a)
  If instr(1,a(i),"(")>0 Then
    a(i)=mid(a(i), instr(1,a(i),"("))
    If len(a(i))>2 then
     a(i)=Mid(a(i),3)
    End If
   Else
    a(i)=mid(a(i), 2)
  End If
  d.add i, a(i)
 Next
 Set Sort = d
End Function