Attribute VB_Name = "modShell"
Option Explicit

Public Function ShellAndWait(ByVal Operation As String, _
                             ByVal FilePath As String, _
                             Optional Parameter As String, _
                             Optional WorkingFolder As String, _
                             Optional WindowSize As ShowConstants = 1, _
                             Optional WaitFor As WaitConstants = 0, _
                             Optional Milliseconds As Long = -1, _
                             Optional CloseProcess As Boolean = False) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020  Dim res As Long, ShExInfo As SHELLEXECUTEINFO
50030  If WorkingFolder = vbNullString Then
50040   WorkingFolder = FilePath
50050  End If
50060  With ShExInfo
50070   .cbSize = Len(ShExInfo)
50080   .fMask = SEE_MASK_FLAG_NO_UI Or SEE_MASK_NOCLOSEPROCESS
50090   .hwnd = frmMain.hwnd
50100   .lpVerb = Operation
50110   .lpFile = FilePath
50120   .lpParameters = Parameter
50130   .lpDirectory = WorkingFolder
50140   .nShow = WindowSize
50150  End With
50160
50170  res = ShellExecuteEx(ShExInfo)
50180
50190  If res = 0 Then
50200 '  ShellAndWait = ShellExecError(ShExInfo.hInstApp)
50210   Exit Function
50220  End If
50230
50240  If WaitFor <> WCNone Then
50250   If WaitFor = WCInitialisiert Then
50260     res = WaitForInputIdle(ShExInfo.hProcess, Milliseconds)
50270    Else
50280     Do
50290      DoEvents
50300     Loop Until WaitForSingleObject(ShExInfo.hProcess, 0) <> WAIT_TIMEOUT
50310   End If
50320   If res = WAIT_FAILED Then
50330    ShellAndWait = "Wait on process failed."
50340   End If
50350  End If
50360
50370  If CloseProcess = True Then
50380   res = TerminateProcess(ShExInfo.hProcess, 1)
50390   If res <> 0 Then
50400    ShellAndWait = "Can't close application."
50410   End If
50420  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modShell", "ShellAndWait")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

'Private Function ShellExecError(ByVal ErrorCode As Integer) As String
'      Select Case ErrorCode
'          Case 2:  ShellExecError = Dateipfad & " wurde nicht gefunden."
'          Case 3:  ShellExecError = Dateipfad & " wurde nicht gefunden."
'          Case 5:  ShellExecError = "Zugriff verweigert."
'          Case 8:  ShellExecError = "Nicht genügend Speicher verfügbar."
'          Case 26: ShellExecError = Dateipfad & " konnte nicht göffnet werden da sie bereits verwendet wird."
'          Case 27: ShellExecError = "Dateityp ist nicht ausreichend Assizoiert."
'          Case 28: ShellExecError = "DDE Zeitlimit wurde ereicht."
'          Case 29: ShellExecError = "DDE ist gescheitert."
'          Case 30: ShellExecError = "DDE konnte nicht gestartet werden."
'          Case 32: ShellExecError = "Eine benötigte Dll wurde nicht gefunden."
'          Case 31: ShellExecError = Dateipfad & " ist mit keiner Anwendung verknüpft."
'       End Select
'End Function
'
'
