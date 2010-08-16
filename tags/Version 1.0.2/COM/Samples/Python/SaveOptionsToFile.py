# -*- coding: ISO-8859-1 -*-
# SaveOptionsToFile.py script
# Part of PDFCreator
# License: GPL
# Homepage: http://www.pdfforge.org/products/pdfcreator
# python 2.5, pywin32 2.5.1
# Version: 1.0.0.0
# Date: December, 24. 2007
# Author: Frank Heindörfer
# Comments: Save the pdfcreator options as ini-file.

import win32com.client as com
import os

PDFCreator = com.Dispatch('PDFCreator.clsPDFCreator')
ProgramIsRunning = PDFCreator.cProgramIsRunning
PDFCreator.cStart("/NoProcessingAtStartup", 1)

PDFCreator.cSaveOptionsToFile(os.getcwd() + '\\PDFCreator.ini')

if ProgramIsRunning:
	PDFCreator.cClose