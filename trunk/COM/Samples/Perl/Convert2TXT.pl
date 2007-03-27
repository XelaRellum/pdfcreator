# Convert2TXT.pl script
# Part of PDFCreator
# License: GPL
# Homepage: http://www.sf.net/projects/pdfcreator
# Version: 1.0.0.0
# Date: March, 20. 2007
# Author: Frank Heindörfer
#Perl version: ActivePerl-5.6.1.638-MSWin32-x86
# Comments: This script convert a printable file in a txt-file using 
#           the com interface of PDFCreator.
# This script doesn't use com events. (alpha level) -> http://search.cpan.org/~jdb/libwin32-0.27/OLE/lib/Win32/OLE.pm#Events

use strict;
use Win32::OLE;
use File::Basename;
use Cwd;

if (!@ARGV)
{
 print "Syntax: perl $0 <Filename>";
 exit
}

my $PDFCreator = Win32::OLE->new("PDFCreator.clsPDFCreator", "cClose") || die "Could not start PDFCreator!";
 
$PDFCreator->cStart("/NoProcessingAtStartup");

my $PDFCreatorOptions = Win32::OLE->new("PDFCreator.clsPDFCreator") || die "Could not get a PDFCreator options object!";

$PDFCreatorOptions = $PDFCreator->{cOptions};

my $cdir = getcwd();
$cdir =~ s/\//\\/g;

$PDFCreatorOptions->{UseAutosave} = 1;
$PDFCreatorOptions->{UseAutosaveDirectory} = 1;
$PDFCreatorOptions->{AutosaveFormat} = 8;                             # 8 = Ascii
my $DefaultPrinter = $PDFCreator->cDefaultPrinter();
$PDFCreator->cDefaultPrinter("PDFCreator");
$PDFCreator->cClearCache();

foreach my $ifname (@ARGV)
{
 my ($file,$dir,$ext) = fileparse($ifname, qr/\.[^.]*/);
 if ($dir eq ".\\") { $dir = $cdir; $ifname = $dir . "\\" . $ifname}

 $PDFCreatorOptions->{AutosaveDirectory} = $dir;
 $PDFCreatorOptions->{AutosaveFilename} = basename($file);
 $PDFCreator->{cOptions} = $PDFCreatorOptions;

 if (!$PDFCreator->cIsPrintable($ifname))
 {
  print "Converting: $ifname\n\nAn error is occured: File '$ifname' is not printable!";
  exit;
 } 

 $PDFCreator->cPrintfile($ifname);

 until (($PDFCreator->{cCountOfPrintjobs} != 0) )
 {
  sleep(1);  # PDFCreator needs time for printing.
 }

 my $counter = 0;
 until (($PDFCreator->{cCountOfPrintjobs} == 0) || ($counter > 300))
 {
  if ($counter == 0)
  {
   $PDFCreator->{cPrinterStop} = 0;
  }
  $counter++;
  sleep(1);
 }
}

$PDFCreator->cDefaultPrinter($DefaultPrinter);
sleep(1);
$PDFCreator->cClose();
sleep(1);

print "Ready";
