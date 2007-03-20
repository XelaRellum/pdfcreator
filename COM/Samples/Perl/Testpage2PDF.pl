# Testpage2PDF.pl script
# Part of PDFCreator
# License: GPL
# Homepage: http://www.sf.net/projects/pdfcreator
# Version: 1.0.0.0
# Date: March, 20. 2007
# Author: Frank Heindörfer
#Perl version: ActivePerl-5.6.1.638-MSWin32-x86
# Comments: Save the test page as pdf-file using
#           the com interface of PDFCreator.

use strict;
use Win32::OLE;
use Cwd;

my $PDFCreator = Win32::OLE->new("PDFCreator.clsPDFCreator", "cClose") || die "Could not start PDFCreator!";

$PDFCreator->cStart("/NoProcessingAtStartup");

my $PDFCreatorOptions = Win32::OLE->new("PDFCreator.clsPDFCreator") || die "Could not get a PDFCreator options object!";
my $PDFCreatorOptionsSave = Win32::OLE->new("PDFCreator.clsPDFCreator") || die "Could not get a PDFCreator options object!";

$PDFCreatorOptions = $PDFCreator->{cOptions};
$PDFCreatorOptionsSave = $PDFCreatorOptions;

my $cdir = getcwd();
$cdir =~ s/\//\\/g;

$PDFCreatorOptions->{AutosaveDirectory} =$cdir;
$PDFCreatorOptions->{UseAutosave} = 1;
$PDFCreatorOptions->{UseAutosaveDirectory} = 1;
$PDFCreatorOptions->{AutosaveFilename} = "Testpage - PDFCreator";
$PDFCreatorOptions->{AutosaveFormat} = 0;
$PDFCreator->{cOptions} = $PDFCreatorOptions;

my $DefaultPrinter = $PDFCreator->cDefaultPrinter();
$PDFCreator->cDefaultPrinter("PDFCreator");
$PDFCreator->cClearCache();
$PDFCreator->cPrintPDFCreatorTestpage();

my $counter = 0;
until (($PDFCreator->{cCountOfPrintjobs} == 0) || ($counter > 30))
{
 if ($counter == 0)
 {
  $PDFCreator->{cPrinterStop} = 0;
 }
 $counter++;
 sleep(1);
}

$PDFCreator->{cOptions} = $PDFCreatorOptionsSave;
$PDFCreator->cDefaultPrinter($DefaultPrinter);
sleep(1);
$PDFCreator->cClose();
sleep(1);

print "Ready";
