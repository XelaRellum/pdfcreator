Name:     pdfforge.org
Author:   Frank Heindörfer
Email:    software@heindoerfer.com
Homepage: http://www.pdfforge.org
License:  FairPlay License Version 1.0 (FairPlay License.txt)
Remark:   Using itextsharp 4.0.7.0.
Date:     24-Dec-2007


Classes, methods of the pdfforge.dll.

pdfforge.DLLInfo
 string Author { get;}
 string Company { get;}
 string Copyright { get;}
 string License { get;}
 string Name { get;}
 string Version { get;}

pdfforge.pdf
 pdfforge.pdf.PDF
 {
  int AddTextToPDFFile(string sourceFilename, string destinationFilename, int fromPage, int toPage, ref PDFText textObject);
  int CopyPDFFile(string sourceFilename, string destinationFilename, int fromPage, int toPage);
  int CreatePDFTestDocument(string destinationFilename, int countOfPages, string additonalText);
  void EncryptPDFFile(string sourceFilename, string destinationFilename, ref PDFEncryptor enc);
  int FileLength(string filename);
  string GetMetadata(string sourceFilename, string key);
  int Images2PDF(ref string[] sourceFilenames, string destinationFilename, bool fitImage);
  void MergePDFFiles(ref string[] sourceFilenames, string destinationFilename);
  int NumberOfPages(string filename);
  int NUp(string sourceFilename, string destinationFilename, int pagesPerPage);
  string PDFVersion(string filename);
  int RemoveEmptyPagesFromPDFFile(string sourceFilename, string destinationFilename);
  int RemovePageFromPDFFile(string sourceFilename, string destinationFilename, int pageNumber);
  int SetBackgroundColor(string sourceFilename, string destinationFilename, int fromPage, int toPage, byte Red, byte Green, byte Blue);
  void SetMetadata(string sourceFilename, string destinationFilename, string author, string creator, string keywords, string subject, string title);
  bool SetMetadataKey(string sourceFilename, string destinationFilename, string key, string value);
  void SignPDFFile(string sourceFilename, string destinationFilename, string certficateFilename, string certifcatePassword, string signatureReason,
	string signatureContact, string signatureLocation, bool signatureVisible, int signaturePositionLowerLeftX, int signaturePositionLowerLeftY, int signaturePositionUpperRightX,
	int signaturePositionUpperRightY, bool multiSignatures, ref PDFEncryptor enc);
  int SplitPDFFile(string sourceFilename, string destinationFilename);
  int StampPDFFileWithImage(string sourceFilename, string destinationFilename, string imageFilename, int fromPage, int toPage, bool overUnder, float fillOpacity, int blendMode);
  int StampPDFFileWithPDFFile(string sourceFilename, string destinationFilename, string pdfFilename, int fromPage, int toPage, bool overUnder, float fillOpacity, int blendMode);
  void UpdateXMPMetadata(string sourceFilename, string destinationFilename);
 }
 pdfforge.pdf.PDFText
 {
  string Text { get; set;}
  string FontPath { get; set;}
  string FontName { get; set;}
  float FontSize { get; set;}
  byte FontColorRed { get; set;}
  byte FontColorGreen { get; set;}
  byte FontColorBlue { get; set;}
  float Rotation { get; set;}
  float XPosition { get; set;}
  float YPosition { get; set;}
  float FillOpacity { get; set;}
  }
 pdfforge.pdf.PDFEncryptor
 {
  bool Strength { get; set;}
  string UserPassword { get; set;}
  string OwnerPassword { get; set;}
  bool AllowAssembly { get; set;}
  bool AllowCopy { get; set;}
  bool AllowDegradedPrinting { get; set;}
  bool AllowFillIn { get; set;}
  bool AllowModifyAnnotations { get; set;}
  bool AllowModifyContents { get; set;}
  bool AllowPrinting { get; set;}
  bool AllowScreenreaders { get; set;}
 }

pdfforge.Tools
 bool CreateTestImage(string destinationFilename, int Red, int Green, int Blue);

