# -*- coding: ISO-8859-1 -*-
# Testpage2PDF.py script
# Part of PDFCreator
# License: GPL
# Homepage: http://www.pdfforge.org/products/pdfcreator
# python 2.5, pywin32 2.5.1
# Version: 1.0.0.0
# Date: December, 24. 2007
# Author: Frank Heindörfer
# Comments: Save the test page as pdf-file using
#           the com interface of PDFCreator.

import win32com.client as com
import os

from time import sleep
from win32print import GetDefaultPrinter, SetDefaultPrinter 
from pythoncom import PumpWaitingMessages

sleepTime = 0.25 # in seconds
maxTime = 10     # in seconds
ReadyState = 0

class clsPDFCreatorEvents(object):
	def OneReady(self):
		print "Ready."
		global ReadyState
		ReadyState = 1
	def OneError(self):
		ReadyState = 0
		global PDFCreator
		print "An error is occured!\r\n Error [" + \
		str(PDFCreator.cErrorDetail("Number")) + "]: " + \
		PDFCreator.cErrorDetail("Description")
		errNum = PDFCreator.cErrorDetail("Number")
		PDFCreator = None
		os._exit(errNum)
		
PDFCreator = com.Dispatch('PDFCreator.clsPDFCreator')
PDFCreatorEvents = com.DispatchWithEvents(PDFCreator, clsPDFCreatorEvents)
ProgramIsRunning = PDFCreator.cProgramIsRunning
PDFCreator.cStart("/NoProcessingAtStartup", 1)

options = PDFCreator.cOptions
options.UseAutosave = 1
options.UseAutosaveDirectory = 1
options.AutosaveDirectory = os.getcwd()
options.AutosaveFilename = "Testpage - PDFCreator"
options.AutosaveFormat = 0                            # 0 = PDF
PDFCreator.cOptions = options
PDFCreator.cSaveOptions()
CurrentPrinter = GetDefaultPrinter()
SetDefaultPrinter("PDFCreator")
PDFCreator.cClearCache()
PDFCreator.cPrintPDFCreatorTestpage()
PDFCreator.cPrinterStop = 0

c = 0

while (ReadyState == 0) and (c < (maxTime / sleepTime)):
	c = c + 1
	sleep(sleepTime)
	PumpWaitingMessages() 

SetDefaultPrinter(CurrentPrinter)
sleep(sleepTime)

if ProgramIsRunning == 0:
	PDFCreator.cClose

if ReadyState == 0:
	a = 0
	print "Creating test page as pdf.\r\n\r\n" + \
	"An error is occured: Time is up!"
