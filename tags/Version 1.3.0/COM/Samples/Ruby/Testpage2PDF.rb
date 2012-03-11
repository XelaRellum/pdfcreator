# Testpage2PDF.rb script
# Part of PDFCreator
# License: GPL
# Homepage: http://www.pdfforge.org/products/pdfcreator
# Ruby version: 1.8.6.0
# Version: 1.0.0.1
# Date: December, 4. 2007
# Author: Frank Heindörfer
# Comments: Save the test page as pdf-file using
#           the com interface of PDFCreator.

require 'win32ole'
require 'win32/process'

SleepTime = 10
$readyState = 0

pdfcreator = WIN32OLE.new('PDFCreator.clsPDFCreator')
event = WIN32OLE_EVENT.new(pdfcreator)
event.on_event('eReady') do
 $readyState = 1
end
event.on_event('eError') do
 print "An error is occured!\nError [", 
  pdfcreator.cErrorDetail('Number'), "]: ",
  pdfcreator.cErrorDetail('Description')
 Process.exit
end

pdfcreator.cStart('/NoProcessingAtStartup')
pdfcreator.setproperty('cOption', 'UseAutosave', 1)
pdfcreator.setproperty('cOption', 'UseAutosaveDirectory', 1)
pdfcreator.setproperty('cOption', 'AutosaveDirectory', Dir.getwd.gsub(/\//,'\\'))
pdfcreator.setproperty('cOption', 'AutosaveFilename', 'Testpage - PDFCreator')
pdfcreator.setproperty('cOption', 'AutosaveFormat', 0) # 0 = PDF
DefaultPrinter = pdfcreator.cDefaultprinter
pdfcreator.setproperty('cDefaultprinter', 'PDFCreator')
pdfcreator.cClearCache()
pdfcreator.cPrintPDFCreatorTestpage()
sleep 2
pdfcreator.setproperty('cPrinterStop', false)

1.upto SleepTime do
 pdfcreator.cOption('UseAutosave') # dummy command, otherwise the com server can't send an event to the script
 break if $readyState != 0
 sleep 1
end

pdfcreator.setproperty('cDefaultprinter', DefaultPrinter)
pdfcreator.cClearCache()
pdfcreator.cClose()

if $readyState == 0 then
 print "Creating test page as pdf.\nAn error is occured: Time is up!"
else
 print "Ready"
end
pdfcreator = nil
Process.exit
