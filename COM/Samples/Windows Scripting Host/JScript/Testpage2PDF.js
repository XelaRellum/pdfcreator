// Testpage2PDF.js script
// Part of PDFCreator
// License: GPL
// Homepage: http://www.sf.net/projects/pdfcreator
// Version: 1.0.0.0
// Date: March, 19 2007
// Author: Frank Heindörfer
// Comments: Save the test page as pdf-file using
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

PDFCreator = WScript.CreateObject("PDFCreator.clsPDFCreator", "PDFCreator_");
PDFCreator.cStart("/NoProcessingAtStartup");

ReadyState = 0
PDFCreator.cOption("UseAutosave") = 1;
PDFCreator.cOption("UseAutosaveDirectory") = 1;
PDFCreator.cOption("AutosaveDirectory") = fso.GetParentFolderName(WScript.ScriptFullname);
PDFCreator.cOption("AutosaveFilename") = "Testpage - PDFCreator";
PDFCreator.cOption("AutosaveFormat") = 0;                             // 0 = PDF
DefaultPrinter = PDFCreator.cDefaultprinter;
PDFCreator.cDefaultprinter = "PDFCreator";
PDFCreator.cClearcache();
PDFCreator.cPrintPDFCreatorTestpage();
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