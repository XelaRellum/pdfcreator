// Convert2TIFF.js script
// Part of PDFCreator
// License: GPL
// Homepage: http://www.pdfforge.org/products/pdfcreator
// Windows Scripting Host version: 5.1
// Version: 1.0.0.0
// Date: March, 19. 2007
// Author: Frank Heindörfer
// Comments: This script convert a printable file in a tiff-file using 
//           the com interface of PDFCreator.

var maxTime = 30    // in seconds
var sleepTime = 250 // in milliseconds

var objArgs, ifname, fso, PDFCreator, DefaultPrinter, ReadyState,
 i, c, Scriptname;

fso = new ActiveXObject("Scripting.FileSystemObject");

Scriptname = fso.GetFileName(WScript.ScriptFullname);

if (WScript.Version < 5.1)
{
 WScript.Echo("You need the \"Windows Scripting Host version 5.1\" or greater!");
 WScript.Quit();
}

if (WScript.arguments.length == 0)
{
 WScript.Echo("Syntax: \t" + Scriptname + " <Filename>\r\n\tor use \"Drag and Drop\"!");
 WScript.Quit();
}

PDFCreator = WScript.CreateObject("PDFCreator.clsPDFCreator", "PDFCreator_");
PDFCreator.cStart("/NoProcessingAtStartup");

PDFCreator.cOption("UseAutosave") = 1;
PDFCreator.cOption("UseAutosaveDirectory") = 1;
PDFCreator.cOption("AutosaveFormat") = 5;                             // 5 = TIFF
DefaultPrinter = PDFCreator.cDefaultprinter;
PDFCreator.cDefaultprinter = "PDFCreator";
PDFCreator.cClearcache();

for (i = 0; i < WScript.arguments.length; i++)
{
 ifname = WScript.arguments.item(i)

 if (!fso.FileExists(ifname))
 {
   WScript.Echo("Can't find the file: " + ifname);
   break;
 }
 if (!PDFCreator.cIsPrintable(ifname))
 {
  WScript.Echos("Converting: " + ifname + "\r\n\r\nAn error is occured: File is not printable!");
  WScript.Quit();
 }

 ReadyState = 0

 PDFCreator.cOption("AutosaveDirectory") = fso.GetParentFolderName(ifname);
 PDFCreator.cOption("AutosaveFilename") = fso.GetBaseName(ifname);
 PDFCreator.cPrintfile(ifname);
 PDFCreator.cPrinterStop = false;

 c = 0
 while ((ReadyState == 0) && (c < (maxTime * 1000 / sleepTime)))
 {
  c = c + 1;
  WScript.Sleep(sleepTime);
 }
 if (ReadyState == 0)
 {
  WScript.Echo("Converting: " + ifname + "\r\n\r\nAn error is occured: Time is up!");
  break;
 }
}

PDFCreator.cDefaultprinter = DefaultPrinter
PDFCreator.cClearcache();
WScript.Sleep(200);
PDFCreator.cClose();

//--- PDFCreator events ---

function PDFCreator_eReady()
{
 ReadyState = 1
}

function PDFCreator_eError()
{
 WScript.Echo("An error is occured!\r\n\r\n"  +
  "Error [" + PDFCreator.cErrorDetail("Number") + "]: " + PDFCreator.cErrorDetail("Description"));
 WScript.Quit();
}