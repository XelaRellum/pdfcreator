' Logger script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.sf.net/projects/pdfcreator
' Version: 1.0.0.0
' Date: July, 18. 2005
' Author: Frank Heindörfer

Option Explicit

Const AppTitle = "PDFCreator - Logger"
Const LogFile = "PDFCreator-Logfile.csv"

Dim objArgs, sep, fso, f


Set objArgs = WScript.Arguments

If objArgs.Count = 0 Then
 MsgBox "You can't call the sctipt from commandline!", vbExclamation, AppTitle
 WScript.Quit
End If

Set fso = CreateObject("Scripting.FileSystemObject")

If Not fso.FileExists(LogFile) Then
 WriteToFile LogFile, "Time", "File", "Filsize", "User", "Machine"
ENd If

Set f = fso.GetFile(objArgs(0))
WriteToFile LogFile, Now, objArgs(0), f.Size, objArgs(1), Replace(objArgs(2),"\\","")


Private Sub WriteToFile(File, Str1, Str2, Str3, Str4, Str5)
 Const ForAppending = 8
 Dim fso, f
 Set fso = CreateObject("Scripting.FileSystemObject")
 Set f = fso.OpenTextFile(File, ForAppending, True)
 sep = GetListSeparator
 f.WriteLine ReplaceForbiddenChars(Str1, sep) & sep & _
  ReplaceForbiddenChars(Str2, sep) & sep & _
  ReplaceForbiddenChars(Str3, sep) & sep & _
  ReplaceForbiddenChars(Str4, sep) & sep & _
  ReplaceForbiddenChars(Str5, sep)
 f.Close
End Sub

Private Function ReplaceForbiddenChars(Str1, sep)
 If Instr(Str1,"""") Then
  Str1 = Replace(Str1, """", """""")
 End IF
 If Instr(Str1,sep) Then
  Str1 = """" & Str1 & """"
 End IF
 ReplaceForbiddenChars = Str1
End Function

Private Function GetListSeparator
 On Error Resume Next
 Dim WshShell, sep
 Set WshShell = WScript.CreateObject("WScript.Shell")
 sep = WshShell.RegRead("HKCU\Control Panel\International\sList")
 If LenB(sep) = 0 Then
   GetListSeparator = ";"
  else
   GetListSeparator = sep
 End If  
End Function