' AddPDFDocview.vbs.vbs script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.sf.net/projects/pdfcreator
' Version: 1.0.0.0
' Date: September, 1. 2005
' Author: Frank Heindörfer
' Comment: This script adds some infos to a postscript file, how the viewer should open the pdf file.

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

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(fname, ForAppending, True)
f.WriteLine

' Use one /PageMode setting

' Open the file in fullscreen
' f.writeline "[ /PageMode FullScreen /Docview pdfmark"

' Open the file without bookmarks and thumbnails, start with page 1 and fit the document (Fit, FitB)
f.WriteLine "[ /PageMode /UseNone /Page 1 /View [/Fit] /DOCVIEW pdfmark"

' Use one {Catalog} object

' Define the pagelayout (SinglePage, OneColumn, TwoColumnLeft, TwoColumnRight)
' f.writeline "[ {Catalog} << PageLayout /TwoColumnRight >> /PUT pdfmark"

' Define the viewer preferences (HideToolbar, HideMenubar, HideWindowUI, FitWindow, CenterWindow, NonFullScreenPageMode, DisplayDocTitle, Direction (L2R, R2L))
f.writeline "[ {Catalog} << /ViewerPreferences << /HideToolbar true /HideWindowUI true  /HideMenubar true >> >> /PUT pdfmark"

f.WriteLine "%%EOF"
f.Close