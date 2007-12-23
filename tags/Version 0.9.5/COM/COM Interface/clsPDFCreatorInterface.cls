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

Public Property Let cVisible(ByVal Value As Boolean)
End Property

Public Property Get cInstalledAsServer() As Boolean
End Property

Public Property Get cPrinterStop() As Boolean
End Property

Public Property Let cPrinterStop(ByVal Value As Boolean)
End Property

Public Property Get cOptionsNames() As Collection
End Property

Public Property Get cOption(ByVal PropertyName As String) As Variant
End Property

Public Property Let cOption(ByVal PropertyName As String, ByVal Value As Variant)
End Property

Public Property Get cOptions() As clsPDFCreatorOptions
End Property

Public Property Set cOptions(ByVal Options1 As clsPDFCreatorOptions)
End Property

Public Property Get cStandardOption(ByVal PropertyName As String) As Variant
End Property

Public Property Get cStandardOptions() As clsPDFCreatorOptions
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

Public Property Get cWindowState() As Long
End Property

Public Property Let cWindowState(ByVal ProgramWindowState As Long)
End Property

Public Property Get cIsConverted() As Boolean
End Property

Public Property Let cIsConverted(ByVal Value As Boolean)
End Property
' --- Public properties - end ---

' --- Public functions/Subroutines - begin ---
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

Public Function cCloseRunningSession() As Boolean
End Function

Public Function cStart(Optional ByVal Params As String, Optional ByVal ForceInitialize As Boolean = False) As Boolean
End Function

Public Sub cClearCache()
End Sub

Public Sub cClearLogfile()
End Sub

Public Sub cConvertPostscriptfile(ByVal InputFilename As String, ByVal OutputFilename As String)
End Sub

Public Sub cTestEvent(ByVal EventName As String)
End Sub

Public Sub cShowLogfileDialog(ByVal Value As Boolean)
End Sub

Public Sub cShowOptionsDialog(ByVal Value As Boolean)
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
' --- Public functions/Subroutines - end ---

