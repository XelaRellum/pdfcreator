Attribute VB_Name = "modPrinterSet"
Option Explicit

Private Const VER_PLATFORM_WIN32s As Long = &H0
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = &H1
Private Const VER_PLATFORM_WIN32_NT As Long = &H2

Private Const ERROR_INSUFFICIENT_BUFFER As Long = 122

Private Const PRINTER_ATTRIBUTE_DEFAULT As Long = &H4

Private Const SMTO_NORMAL = &H0
Private Const SMTO_BLOCK = &H1
Private Const SMTO_ABORTIFHUNG = &H2
Private Const HWND_BROADCAST = &HFFFF&
Private Const WM_SETTINGCHANGE = &H1A


Private Type OSVERSIONINFO
 dwOSVersionInfoSize As Long
 dwMajorVersion As Long
 dwMinorVersion As Long
 dwBuildNumber As Long
 dwPlatformId As Long
 szCSDVersion As String * 128
End Type

Private Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrn As Long, pDefault As Any) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrn As Long) As Long
Private Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal Command As Long) As Long
Private Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetDefaultprinter Lib "winspool.drv" Alias "SetDefaultPrinterA" (ByVal pszPrinter As String) As Long
Private Declare Function GetProfileString Lib "Kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function WriteProfileString Lib "Kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long

Public Sub SetDefaultprinterInProg(Printername As String)
 Dim os As OSVERSIONINFO

 os.dwOSVersionInfoSize = Len(os)
 Call GetVersionEx(os)
 Select Case os.dwPlatformId
  Case VER_PLATFORM_WIN32_WINDOWS
   Call DefaultPrinterSet9x(Printername)
  Case VER_PLATFORM_WIN32_NT
   Call DefaultPrinterSetNT(os.dwMajorVersion, Printername)
 End Select
End Sub

Private Sub DefaultPrinterSet9x(Printername As String)
 Dim hPrn As Long, BytesNeeded As Long, Buffer() As Byte, BytesUsed As Long, _
  Attributes As Long

 Call OpenPrinter(Printername, hPrn, ByVal 0&)
 If hPrn <> 0 Then
  Call GetPrinter(hPrn, 2, ByVal 0&, 0, BytesNeeded)
  If Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
   ReDim Buffer(0 To BytesNeeded - 1) As Byte
   If GetPrinter(hPrn, 2, Buffer(0), BytesNeeded, BytesUsed) Then
    Const AttribOffset As Long = 13 * 4&
    Call CopyMemory(Attributes, Buffer(AttribOffset), 4&)
    Attributes = Attributes Or PRINTER_ATTRIBUTE_DEFAULT
    Call CopyMemory(Buffer(AttribOffset), Attributes, 4&)
    If SetPrinter(hPrn, 2, Buffer(0), 0) Then
     Call SettingChangeAlert(500)
    End If
   End If
  End If
  Call ClosePrinter(hPrn)
 End If
End Sub

Private Sub DefaultPrinterSetNT(ByVal MajorVersion As Long, Printername As String)
 Dim os As OSVERSIONINFO, BufSize As Long, pPrinterName As Long, Result As String, _
  comma As Long
 If MajorVersion >= 5 Then
   Call SetDefaultprinter(Printername)
  Else
   BufSize = 1024
   Result = Space$(BufSize)
   If GetProfileString("PrinterPorts", ByVal Printername, "", Result, BufSize) Then
    comma = InStr(Result, ",")
    comma = InStr(comma + 1, Result, ",")
    If comma <> 0 Then
     Result = Left$(Result, comma - 1)
     Result = Printername & "," & Result
     Call WriteProfileString("Windows", "device", Result)
     Call SettingChangeAlert(500)
    End If
   End If
 End If
End Sub

Private Sub SettingChangeAlert(Optional ByVal Delay As Long = 500)
 Call SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0, SMTO_NORMAL, Delay, ByVal 0&)
End Sub



