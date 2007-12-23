# -*- coding: ISO-8859-1 -*-
# TestEvents.py script
# Part of PDFCreator
# License: GPL
# Homepage: http://www.pdfforge.org/products/pdfcreator
# python 2.5, pywin32 2.5.1
# Version: 1.0.0.0
# Date: December, 24. 2007
# Author: Frank Heindörfer
# Comments: Test the events of the com interface of PDFCreator.

import win32com.client as com

class PDFCreatorEvents(object):
	def OneReady(self):
		print "Ready."
	def OneError(self):
		print "An error is occured!\r\n Error [" + \
		str(PDFCreator.cErrorDetail("Number")) + "]: " + \
		PDFCreator.cErrorDetail("Description")
		
PDFCreator = com.DispatchWithEvents('PDFCreator.clsPDFCreator', PDFCreatorEvents)
PDFCreator.cTestEvent("Ready")
PDFCreator.cTestEvent("Error")
PDFCreator.cTestEvent("Unknown")
PDFCreator.cClose
PDFCreator = None
