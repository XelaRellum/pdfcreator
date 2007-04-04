# Convert2PDF.rb script
# Part of PDFCreator
# License: GPL
# Homepage: http://www.sf.net/projects/pdfcreator
# Version: 1.0.0.0
# Date: April, 4. 2007
# Author: Frank Heindörfer
# Comments: This script convert a printable file in a pdf-file using 
#           the com interface of PDFCreator.

require 'win32ole'

SleepTime = 30
$readyState = 0

pdfcreator = WIN32OLE.new('PDFCreator.clsPDFCreator')
event = WIN32OLE_EVENT.new(pdfcreator)
event.on_event('eReady') do
 puts 'Ready'
 $readyState = 1
end
event.on_event('eError') do
 puts 'An error is occured!\nError [', 
  pdfcreator.cErrorDetail('Number'), ']: ',
  pdfcreator.cErrorDetail('Description')
 pdfcreator.cClose()
 pdfcreator = nil
 Process.exit
end

if ARGV.length == 0 then
 puts 'Syntax:\tConvert2PDF.rb <Filename>!'
 pdfcreator.cClose()
 pdfcreator = nil
 Process.exit
end

pdfcreator.cStart('/NoProcessingAtStartup')
pdfcreator.setproperty('cOption', 'UseAutosave', 1)
pdfcreator.setproperty('cOption', 'UseAutosaveDirectory', 1)
pdfcreator.setproperty('cOption', 'AutosaveFormat', 0) # 0 = PDF
DefaultPrinter = pdfcreator.cDefaultprinter
pdfcreator.setproperty('cDefaultprinter', 'PDFCreator')
pdfcreator.cClearCache()
pdfcreator.setproperty('cPrinterStop', false)


0.upto(ARGV.length - 1) do |i|
 ifname =  ARGV[i]
 if File.dirname(ifname) == '.'
  ifname = Dir.getwd.gsub(/\//,'\\') + '\\' + ifname
 end
 if !FileTest.exist?(ARGV[i]) then
  print 'Can''t find the file: ', ifname
  break
 end
 if !pdfcreator.cIsPrintable(ifname) then
  print 'Converting: ', ifname, '\n',
   ' An error is occured: File is not printable!'
  break
 end
 $readyState = 0

 pdfcreator.setproperty('cOption', 'AutosaveDirectory', File.dirname(ifname))
 pdfcreator.setproperty('cOption', 'AutosaveFilename', File.basename(ifname, File.extname(ifname)))
 pdfcreator.cPrintfile(ifname)

 1.upto SleepTime do
  pdfcreator.cOption('UseAutosave') # dummy command, otherwise the com server can't send an event to the script
  break if $readyState != 0
  sleep 1
 end
 if $readyState == 0 then
  print 'Converting: ', ifname, '\n',
   ' An error is occured: Time is up!'
  break 
 end
end

pdfcreator.setproperty('cDefaultprinter', DefaultPrinter)
pdfcreator.cClearCache()
pdfcreator.cClose()
pdfcreator = nil
Process.exit 
