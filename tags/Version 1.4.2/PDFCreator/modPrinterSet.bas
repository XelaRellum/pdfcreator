Attribute VB_Name = "modPrinterSet"
Option Explicit

Public Sub SetDefaultprinterInProg(PrinterName As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim os As OSVERSIONINFO
50020
50030  os.OSVSize = Len(os)
50040  Call GetVersionEx(os)
50051  Select Case os.PlatformID
        Case VER_PLATFORM_WIN32_WINDOWS
50070    Call DefaultPrinterSet9x(PrinterName)
50080   Case VER_PLATFORM_WIN32_NT
50090    Call DefaultPrinterSetNT(os.dwVerMajor, PrinterName)
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

Private Sub DefaultPrinterSet9x(PrinterName As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hPrn As Long, BytesNeeded As Long, buffer() As Byte, BytesUsed As Long, _
  Attributes As Long, pd As PRINTER_DEFAULTS
50030
50040  Call OpenPrinter(PrinterName, hPrn, pd)
50050  If hPrn <> 0 Then
50060   Call GetPrinter(hPrn, 2, ByVal 0&, 0, BytesNeeded)
50070   If Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
50080    ReDim buffer(0 To BytesNeeded - 1) As Byte
50090    If GetPrinter(hPrn, 2, buffer(0), BytesNeeded, BytesUsed) Then
50100     Const AttribOffset As Long = 13 * 4&
50110     Call MoveMemory(Attributes, buffer(AttribOffset), 4&)
50120     Attributes = Attributes Or PRINTER_ATTRIBUTE_DEFAULT
50130     Call MoveMemory(buffer(AttribOffset), Attributes, 4&)
50140     If SetPrinter(hPrn, 2, buffer(0), 0) Then
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

Private Sub DefaultPrinterSetNT(ByVal MajorVersion As Long, PrinterName As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim os As OSVERSIONINFO, BufSize As Long, pPrinterName As Long, Result As String, _
  comma As Long
50030  If MajorVersion >= 5 Then
50040    Call SetDefaultprinter(PrinterName)
50050   Else
50060    BufSize = 1024
50070    Result = Space$(BufSize)
50080    If GetProfileString("PrinterPorts", ByVal PrinterName, "", Result, BufSize) Then
50090     comma = InStr(Result, ",")
50100     comma = InStr(comma + 1, Result, ",")
50110     If comma <> 0 Then
50120      Result = Left$(Result, comma - 1)
50130      Result = PrinterName & "," & Result
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



