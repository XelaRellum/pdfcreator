# -*- coding: ISO-8859-1 -*-
# Convert2PDF.py script
# Part of PDFCreator
# License: GPL
# Homepage: http://www.pdfforge.org/products/pdfcreator
# python 2.5, pywin32 2.5.1
# Version: 1.0.0.0
# Date: December, 24. 2007
# Author: Frank Heindörfer
# Comments: This script convert a printable file in a pdf-file using 
#           the com interface of PDFCreator.

import win32com.client as com
from win32print import GetDefaultPrinter, SetDefaultPrinter 
import os
import sys

from time import sleep
from pythoncom import PumpWaitingMessages, PumpMessages

sleepTime = 1  # in seconds
maxTime = 10   # in seconds
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

if len(sys.argv) <= 1:
	print "Syntax: " + os.path.basename(sys.argv[0]) + " <Filename>\r\n\tor use ""Drag and Drop""!"
	os._exit(-255)

PDFCreator = com.Dispatch('PDFCreator.clsPDFCreator')
PDFCreatorEvents = com.DispatchWithEvents(PDFCreator, clsPDFCreatorEvents)
ProgramIsRunning = PDFCreator.cProgramIsRunning
PDFCreator.cStart("/NoProcessingAtStartup", 1)

options = PDFCreator.cOptions
oldOptions = PDFCreator.cOptions
options.UseAutosave = 1
options.UseAutosaveDirectory = 1
options.AutosaveFormat = 0                            # 0 = PDF
CurrentPrinter = GetDefaultPrinter()
SetDefaultPrinter('PDFCreator')
PDFCreator.cClearCache
PDFCreator.cPrinterStop = 0

for i in xrange(1, len(sys.argv)):
	ifname = str(sys.argv[i])
	if not os.path.exists(ifname):
		print "Can't find the file: " + ifname
		exit()
	if not PDFCreator.cIsPrintable(ifname):
		print "Converting: " + ifname + "\r\n\r\nAn error is occured: File is not printable!"
		exit()
	ReadyState = 0
	dirname = os.path.dirname(ifname)
	if dirname == "":
		dirname = os.path.dirname(sys.argv[0])
	options.AutosaveDirectory = dirname
	options.AutosaveFilename = os.path.splitext(os.path.basename(ifname))[0]
	PDFCreator.cSaveOptions(options)
	
	PDFCreator.cPrintFile(ifname)
	
	c = 0
	while (ReadyState == 0) and (c < (maxTime / sleepTime)):
		c = c + 1
		sleep(sleepTime)
		PumpWaitingMessages()
	if ReadyState == 0:
		print "Converting: " + ifname + "\r\n\r\nAn error is occured: Time is up!"
		exit()
		
PDFCreator.cOptions = options
PDFCreator.cSaveOptions(oldOptions)
SetDefaultPrinter(CurrentPrinter)
PDFCreator.cClose
PDFCreator = None
print "Done."