Name:     pdfforge.dll
Version:  1.5.0.0
Authors:  Frank Heindörfer, Hannes Smurawsky
Email:    frank@pdfforge.org
Homepage: http://www.pdfforge.org
License:  FairPlay License Version 1.0 (FairPlay License.txt)
Remark:   Using itextsharp 4.1.6.0 (compiled from source)
Date:     September 8, 2009


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
  int AddCropMarksToPDFFile(string sourceFilename, string destinationFilename, int fromPage, int toPage, float borderTop, float borderBottom, float borderLeft, float borderRight, ref PDFLine lineObject);
  int AddLineToPDFFile(string sourceFilename, string destinationFilename, int fromPage, int toPage, ref PDFLine lineObject);
  int AddPageNumberToPDFFile(string sourceFilename, string destinationFilename, int fromPage, int toPage, int startPageNumber, int NumberOfPages, int pageNumberPosition, float borderXMillimeter, float borderYMillimeter, ref PDFText textObject);
  int AddTextToPDFFile(string sourceFilename, string destinationFilename, int fromPage, int toPage, ref PDFText textObject);
  int Brochure(string sourceFilename, string destinationFilename);
  int CopyPDFFile(string sourceFilename, string destinationFilename, int fromPage, int toPage);
  int CreatePDFTestDocument(string destinationFilename, int countOfPages, string additionalText);
  void EncryptPDFFile(string sourceFilename, string destinationFilename, ref PDFEncryptor enc);
  int FileLength(string filename);
  string GetMetadata(string sourceFilename, string key);
  int Images2PDF(ref string[] sourceFilenames, string destinationFilename, int scaleMode);
  int Images2PDF(ref object[] sourceFilenames, string destinationFilename, int scaleMode);
  void MergePDFFiles(ref string[] sourceFilenames, string destinationFilename, bool FilenamesAsBookmarks);
  void MergePDFFiles(ref object[] sourceFilenames, string destinationFilename, bool FilenamesAsBookmarks);
  int NumberOfPages(string filename);
  int NUp(string sourceFilename, string destinationFilename, int pagesPerPage);
  string PDFVersion(string filename);
  int RemoveEmptyPagesFromPDFFile(string sourceFilename, string destinationFilename);
  int RemovePageFromPDFFile(string sourceFilename, string destinationFilename, int pageNumber);
  int ReplacePagesFromPDFFile(string sourceFilename1, string sourceFilename2, string destinationFilename, int source1FromPage, int source1ToPage, int source2FromPage, int source2ToPage);
  int SetBackgroundColor(string sourceFilename, string destinationFilename, int fromPage, int toPage, byte Red, byte Green, byte Blue);
  void SetMetadata(string sourceFilename, string destinationFilename, string author, string creator, string keywords, string subject, string title);
  bool SetMetadataKey(string sourceFilename, string destinationFilename, string key, string value);
  void SignPDFFile(string sourceFilename, string destinationFilename, string certficateFilename, string certifcatePassword,
      string signatureReason, string signatureContact, string signatureLocation,
      bool signatureVisible,
      int signaturePositionLowerLeftX, int signaturePositionLowerLeftY, int signaturePositionUpperRightX, int signaturePositionUpperRightY,
      bool multiSignatures, ref PDFEncryptor enc);
  int SplitPDFFile(string sourceFilename, string destinationFilename);
  int StampPDFFileWithImage(string sourceFilename, string destinationFilename, string imageFilename, int fromPage, int toPage, bool overUnder, float fillOpacity, int blendMode);
  int StampPDFFileWithPDFFile(string sourceFilename, string destinationFilename, string pdfFilename, int fromPage, int toPage, bool overUnder, float fillOpacity, int blendMode);
  void UpdateXMPMetadata(string sourceFilename, string destinationFilename);
 }
 pdfforge.pdf.PDFLine
 {
  float LineThickness { get; set;}
  float FromX { get; set;}
  float FromY { get; set;}
  float ToX { get; set;}
  float ToY { get; set;}
  float UnitsOn { get; set;}
  float UnitsOff { get; set;}
  float Phase { get; set;}
  byte LineColorRed { get; set;}
  byte LineColorGreen { get; set;}
  byte LineColorBlue { get; set;}
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
