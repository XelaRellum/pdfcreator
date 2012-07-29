' SendMail script
' Part of PDFCreator
' License: GPLv3
' Homepage: http://www.pdfforge.org/
' Windows Scripting Host version: 5.1
' Version: 1.0.0.0
' Date: May 23, 2012
' Author: Philip Chinery
' Comments: This script sends a mail using blat
'           The script requires blat.exe (http://www.blat.net/)

Dim cmdline, subject, receipient, sender, server, user, password, bodyFile, blat, additionalParams, file

' Please configure this section to suit your needs

' Receipient's E-Mail address
receipient = "admin@localhost"

' Your E-Mail address
sender = "admin@localhost"

' Subject for the mail
subject = "A new file was converted"

' Server name or IP address
server = "localhost"

' user name - leave empty if none required
user = ""

' password  - leave empty if none required
password = ""

' a plain text file containing the mail body text
bodyFile = ""

' Path to blat.exe (including blat.exe)
blat = "C:\Blat\blat.exe"

' Add other blat params here, if required. i.e. you can use " -log C:\blat-log.txt" for logging
additionalParams = "-log C:\blat-log.txt"

' the actual scripts starts here

Set Wshshell = CreateObject("wscript.shell")

Set objArgs = WScript.Arguments

If objArgs.Count = 0 Then
 MsgBox "This script needs a parameter!", vbExclamation, AppTitle
 WScript.Quit
End If

if bodyFile = "" then
  MsgBox "The body file is not defined. Please configure this script first!"
  WScript.Quit
end if

file = objArgs(0)

cmdline = """" & blat & """ """ & bodyFile & """ -t " & receipient & " -server " & server & " -f " & sender & " -s """ & subject & """ " & additionalParams & " -attach """ & file & """"

if user <> "" then
  cmdline = cmdline & " -u " & user
end if

if password <> "" then
  cmdline = cmdline & " -pw " & password
end if

wshshell.Run cmdline, 1, True