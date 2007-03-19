// TestEvents.js script
// Part of PDFCreator
// License: GPL
// Homepage: http://www.sf.net/projects/pdfcreator
// Version: 1.0.0.0
// Date: March, 19 2007
// Author: Frank Heindörfer
// Comments: Test the events of the com interface of PDFCreator.

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

PDFCreator.cTestEvent("Ready");
PDFCreator.cTestEvent("Error");
PDFCreator.cTestEvent("Unknown");

PDFCreator.cClose();

//--- PDFCreator events ---

function PDFCreator_eReady()
{
 WScript.Echo("Ready!");
}

function PDFCreator_eError()
{
 WScript.Echo("An error is occured!\r\n\r\n"  +
  "Error [" + PDFCreator.cErrorDetail("Number") + "]: " + PDFCreator.cErrorDetail("Description"));
}