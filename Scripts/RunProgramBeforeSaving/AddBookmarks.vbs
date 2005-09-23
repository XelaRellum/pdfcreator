' AddBookmarks.vbs.vbs script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.sf.net/projects/pdfcreator
' Version: 1.1.0.0
' Date: September, 1. 2005
' Author: Frank Heindörfer
' Comment: This script adds simple bookmarks to a postscript file.

Option Explicit

Const AppTitle = "PDFCreator - AddBookmarks"
Const ForReading = 1, ForAppending = 8

Dim objArgs, fname, fso, f, pages, i

Set objArgs = WScript.Arguments

If objArgs.Count = 0 Then
 MsgBox "This script needs a parameter!", vbExclamation, AppTitle
 WScript.Quit
End If

fname = objArgs(0)

pages = GetCountOfPagesFromPostscriptfile(fname)

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(fname, ForAppending, True)
f.writeline "  [/Title (Page " & 1 & ") /Page " & 1 & " /View [/XYZ null null 1] /Count " & pages & " /OUT pdfmark"
For i=1 to pages
 f.writeline "[/Page " & i & " /View [/XYZ null null 1] /Title (Page " & i & ") /OUT pdfmark"
Next
f.Close

Private Function GetCountOfPagesFromPostscriptfile(PostscriptFile)
 Dim fso, f, fstr, pp
 Set fso = CreateObject("Scripting.FileSystemObject")
 Set f = fso.OpenTextFile(PostscriptFile, ForReading, True)
 fstr = f.ReadAll
 f.Close
 pp = InstrRev(fstr, "%%Pages:", -1, 1)
 If pp <= 0 Then
  GetCountOfPagesFromPostscriptfile = 1
  Exit Function
 End If
 pp = Instr(pp, fstr," ", 1)
 If pp <= 0 Then
  GetCountOfPagesFromPostscriptfile = 1
  Exit Function
 End If
 fstr = Trim(Mid(fstr,pp))
 fstr = Replace(fstr, chr(10), " ", 1, -1, 1)
 fstr = Replace(fstr, chr(13), " ", 1, -1, 1)
 pp = Instr(1, fstr," ", 1)
 If pp <= 0 Then
  GetCountOfPagesFromPostscriptfile = 1
  Exit Function
 End If
 fstr=mid(fstr,1,pp-1)
 If Not IsNumeric(fstr) Then
  fstr = 1
 End If
 GetCountOfPagesFromPostscriptfile = fstr
End Function