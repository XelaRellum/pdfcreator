VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPDFCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event eReady()
Public Event eError()

' --- Class subroutines - begin ---
Private Sub Class_Initialize()
End Sub
' --- Class subroutines - end ---

' --- Public properties - begin ---
Public Property Get cError() As clsPDFCreatorError
End Property

Public Property Get cErrorDetail(ByVal PropertyName As String) As Variant
End Property

Public Sub cErrorClear()
End Sub

Public Property Get cGhostscriptVersion() As String
End Property

Public Property Get cOutputFilename() As String
End Property

Public Property Get cPDFCreatorApplicationPath() As String
End Property

Public Property Get cIsLogfileDialogDisplayed() As Boolean
End Property

Public Property Get cIsOptionsDialogDisplayed() As Boolean
End Property

Public Property Get cProgramRelease(Optional ByVal WithBeta As Boolean = True) As String
End Property

Public Property Get cProgramIsRunning() As Boolean
End Property

Public Property Get cWindowsVersion() As String
End Property

Public Property Get cVisible() As Boolean
End Property

Public Property Let cVisible(ByVal value As Boolean)
End Property

Public Property Get cInstalledAsServer() As Boolean
End Property

Public Property Get cPrinterStop() As Boolean
End Property

Public Property Let cPrinterStop(ByVal value As Boolean)
End Property

Public Property Get cOptionsNames() As Collection
End Property

Public Property Get cOption(ByVal PropertyName As String) As Variant
End Property

Public Property Let cOption(ByVal PropertyName As String, ByVal value As Variant)
End Property

Public Property Get cOptions() As clsPDFCreatorOptions
End Property

Public Property Set cOptions(ByVal Options1 As clsPDFCreatorOptions)
End Property

Public Property Get cStandardOption(ByVal PropertyName As String) As Variant
End Property

Public Property Get cStandardOptions() As clsPDFCreatorOptions
End Property

Public Property Get cPostscriptInfo(ByVal PostscriptFilename As String, ByVal PropertyName As String) As String
End Property

Public Property Get cPrintjobInfos(ByVal PrintjobFilename As String) As clsPDFCreatorInfoSpoolFile
End Property

Public Property Get cPrintjobInfo(ByVal PrintjobFilename As String, ByVal PropertyName As String) As String
End Property

Public Property Get cCountOfPrintjobs() As Long
End Property

Public Property Get cPrintjobFilename(ByVal JobNumber As Long) As String
End Property

Public Property Get cDefaultPrinter() As String
End Property

Public Property Let cDefaultPrinter(ByVal Printername As String)
End Property

Public Property Get cStopURLPrinting() As Boolean
End Property

Public Property Let cStopURLPrinting(ByVal value As Boolean)
End Property


Public Property Get cWindowState() As Long
End Property

Public Property Let cWindowState(ByVal value As Long) ' Value must be '0 = normal' or '1 = minimized' or '2 = maximized'
End Property

Public Property Get cIsConverted() As Boolean
End Property

Public Property Let cIsConverted(ByVal value As Boolean)
End Property

Public Property Get cInstanceCounter() As Long
End Property
' --- Public properties - end ---

' --- Public functions/Subroutines - begin ---
Public Sub cAddPrintjob(ByVal Filename As String) ' Supported file types: Postscript, JPEG, BMP
End Sub

Public Sub cDeletePrintjob(ByVal JobNumber As Long)
End Sub

Public Sub cMovePrintjobBottom(ByVal JobNumber As Long)
End Sub

Public Sub cMovePrintjobTop(ByVal JobNumber As Long)
End Sub

Public Sub cMovePrintjobUp(ByVal JobNumber As Long)
End Sub

Public Sub cMovePrintjobDown(ByVal JobNumber As Long)
End Sub

Public Function cClose() As Boolean
End Function

Public Function cCloseRunningSession() As Boolean ' This function is deprecated. Use "cClose()"!
End Function
 
Public Function cStart(Optional ByVal Params As String, Optional ByVal ForceInitialize As Boolean = False) As Boolean
End Function

Public Sub cClearCache()
End Sub

Public Sub cClearLogfile()
End Sub

Public Sub cConvertPostscriptfile(ByVal InputFilename As String, ByVal OutputFilename As String) ' This function is deprecated! Use "cConvertFile()" instead.
End Sub

Public Sub cConvertFile(ByVal InputFilename As String, ByVal OutputFilename As String, Optional ByVal SubFormat As String = "") ' Allowed values for 'SubFormat': "PDF/A-1b" for "PDF/A-1b"files; "PDF/X" for "PDF/X"-files
End Sub

Public Sub cTestEvent(ByVal EventName As String)
End Sub

Public Sub cShowLogfileDialog(ByVal value As Boolean)
End Sub

Public Sub cShowOptionsDialog(ByVal value As Boolean)
End Sub

Public Sub cSendMail(ByVal OutputFilename As String, ByVal Recipients As String)
End Sub

Public Function cIsPrintable(ByVal Filename As String) As Boolean
End Function

Public Sub cCombineAll()
End Sub

Public Function cGetPDFCreatorPrinters() As Collection
End Function

Public Function cGetLogfile() As String
End Function

Public Sub cWriteToLogfile(ByVal LogStr As String)
End Sub

Public Sub cPrintFile(ByVal Filename As String)
End Sub

Public Sub cPrintURL(ByVal URL As String, Optional ByVal TimeBetweenLoadAndPrint As Long = 0)
End Sub

Public Sub cPrintPDFCreatorTestpage()
End Sub

Public Sub cPrintPrinterTestpage(Optional ByVal Printername As String)
End Sub

Public Function cReadOptions() As clsPDFCreatorOptions
End Function

Public Sub cSaveOptions(Optional ByVal Options1 As Variant) ' Options1 must be a clsPDFCreatorOptions type
End Sub

Public Function cReadOptionsFromFile(ByVal INIFilename As String) As clsPDFCreatorOptions
End Function

Public Sub cSaveOptionsToFile(ByVal INIFilename As String, Optional ByVal Options1 As Variant) ' Options1 must be a clsPDFCreatorOptions type
End Sub

Public Function GetPDFCreatorPrinters() As Collection
End Function
' --- Public functions/Subroutines - end ---

Public Function cGhostscriptRun(Arguments() As Variant) As Boolean
End Function

' --- Private functions/Subroutines - end ---
Private Sub eActionTimer_Timer()
End Sub

Private Function CheckForRunningProgramErrors() As Boolean
End Function

Private Function cGetPrintjobInfoFromPrintjobInfos(PropertyName As String, sInfo As InfoSpoolFile, FunctionName As String) As Variant
End Function

Private Function cGetOptionFromOptions(PropertyName As String, Options As tOptions) As Variant
End Function

Private Function cGetInfoSpoolfile(sInfos As InfoSpoolFile) As clsPDFCreatorInfoSpoolFile
End Function

Private Function cGetOptions(sOptions As tOptions) As clsPDFCreatorOptions
End Function

Private Sub cLetOption(PropertyName As String, value As Variant)
End Sub

Private Function cLetOptions(sOptions As Variant) As tOptions
End Function
' --- Private functions/Subroutines - end ---
