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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim os As OSVERSIONINFO
50020
50030  os.dwOSVersionInfoSize = Len(os)
50040  Call GetVersionEx(os)
50051  Select Case os.dwPlatformId
        Case VER_PLATFORM_WIN32_WINDOWS
50070    Call DefaultPrinterSet9x(Printername)
50080   Case VER_PLATFORM_WIN32_NT
50090    Call DefaultPrinterSetNT(os.dwMajorVersion, Printername)
50100  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinterSet", "SetDefaultprinterInProg")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub DefaultPrinterSet9x(Printername As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hPrn As Long, BytesNeeded As Long, Buffer() As Byte, BytesUsed As Long, _
  Attributes As Long
50030
50040  Call OpenPrinter(Printername, hPrn, ByVal 0&)
50050  If hPrn <> 0 Then
50060   Call GetPrinter(hPrn, 2, ByVal 0&, 0, BytesNeeded)
50070   If Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
50080    ReDim Buffer(0 To BytesNeeded - 1) As Byte
50090    If GetPrinter(hPrn, 2, Buffer(0), BytesNeeded, BytesUsed) Then
50100     Const AttribOffset As Long = 13 * 4&
50110     Call CopyMemory(Attributes, Buffer(AttribOffset), 4&)
50120     Attributes = Attributes Or PRINTER_ATTRIBUTE_DEFAULT
50130     Call CopyMemory(Buffer(AttribOffset), Attributes, 4&)
50140     If SetPrinter(hPrn, 2, Buffer(0), 0) Then
50150      Call SettingChangeAlert(500)
50160     End If
50170    End If
50180   End If
50190   Call ClosePrinter(hPrn)
50200  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinterSet", "DefaultPrinterSet9x")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub DefaultPrinterSetNT(ByVal MajorVersion As Long, Printername As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim os As OSVERSIONINFO, BufSize As Long, pPrinterName As Long, Result As String, _
  comma As Long
50030  If MajorVersion >= 5 Then
50040    Call SetDefaultprinter(Printername)
50050   Else
50060    BufSize = 1024
50070    Result = Space$(BufSize)
50080    If GetProfileString("PrinterPorts", ByVal Printername, "", Result, BufSize) Then
50090     comma = InStr(Result, ",")
50100     comma = InStr(comma + 1, Result, ",")
50110     If comma <> 0 Then
50120      Result = Left$(Result, comma - 1)
50130      Result = Printername & "," & Result
50140      Call WriteProfileString("Windows", "device", Result)
50150      Call SettingChangeAlert(500)
50160     End If
50170    End If
50180  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinterSet", "DefaultPrinterSetNT")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SettingChangeAlert(Optional ByVal Delay As Long = 500)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Call SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0, SMTO_NORMAL, Delay, ByVal 0&)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinterSet", "SettingChangeAlert")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub



