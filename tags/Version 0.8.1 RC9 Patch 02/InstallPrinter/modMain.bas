Attribute VB_Name = "modMain"
Option Explicit

Private Enum eInstallAction
 Nil = 0
 Install = 1
 UnInstall = 2
End Enum

Private InstallAction As eInstallAction, InstallSpecialPrinter As Boolean, LogPath As String, _
 AppDir As String

Private Portname As String, Monitorname As String, Drivername As String, Printername As String

Public Sub Main()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim LogFile As String
50020  InstallSpecialPrinter = False
50030  Monitorname = "PDFCreator"
50040  Portname = "PDFCreator:"
50050  Drivername = "PDFCreator"
50060  Printername = "PDFCreator"
50070  LogPath = CompletePath(App.Path)
50080  AppDir = CompletePath(AppDir)
50090  AnalyzeCommandlineParameters
50100  If InstallAction = Install Then
50110   LogFile = AppDir & "SetupLog.txt"
50120   If InstallSpecialPrinter = False Then
50130     InstallWindowsPrinter Monitorname, Portname, Drivername, Printername, LogFile, AppDir
50140     WriteInstalldate2Registry
50150    Else
50160     InstallAdditionalWindowsPrinter Printername, LogFile, AppDir
50170   End If
50180  End If
50190
50200  If InstallAction = UnInstall Then
50210   LogFile = CompletePath(GetTempPathApi) & "PDFCreatorUninstall.txt"
50220   UnInstallWindowsPrinter Monitorname, Portname, Drivername, Printername, LogFile
50230  End If
50240  If InstallAction = Install Or InstallAction = UnInstall Then
50250   WriteToLog "--------------------------------------------------", LogFile
50260   WriteToLog "MSI-Installer: Installer2Go", LogFile
50270   WriteToLog "Computername: " & GetComputerName, LogFile
50280   WriteToLog "Username: " & GetUsername, LogFile
50290   WriteToLog "WinDir: " & GetWindowsDirectory, LogFile
50300   WriteToLog "SysDir: " & GetSystemDirectory, LogFile
50310   WriteToLog "TempDir: " & GetTempPathApi, LogFile
50320   WriteToLog "CurrentDir: " & CurDir, LogFile
50330   WriteToLog "Path: " & Environ$("Path"), LogFile
50340   WriteToLog "Internet Explorer version: " & GetIExplorerVersion, LogFile
50350  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "Main")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub AnalyzeCommandlineParameters()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim csInstall As String, csUninstall As String, csPrintername As String, csLogPath As String, _
  csAppDir As String
50030  If Len(VBA.Command$) > 0 Then
50040   csInstall = CommandSwitch("IN", False)
50050   csUninstall = CommandSwitch("UNIN", False)
50060   csPrintername = CommandSwitch("PRINTERNAME", True)
50070   csLogPath = CommandSwitch("LOGPATH", True)
50080   csAppDir = CommandSwitch("APPDIR", True)
50090   If UCase$(csInstall) = "STALL" And UCase$(csUninstall) = "STALL" Then
50100    MsgBox "Don't use INSTALL and UNINSTALL at the same time!" & vbCrLf & _
    "Program canceled!"
50120   End If
50130   InstallAction = Nil
50140   If UCase$(csInstall) = "STALL" Then
50150    InstallAction = Install
50160   End If
50170   If UCase$(csUninstall) = "STALL" Then
50180    InstallAction = UnInstall
50190   End If
50200   If LenB(csPrintername) > 0 Then
50210    InstallSpecialPrinter = True
50220    Printername = csPrintername
50230   End If
50240   If LenB(csLogPath) > 0 Then
50250    LogPath = CompletePath(csLogPath)
50260   End If
50270   If LenB(csAppDir) > 0 Then
50280    AppDir = CompletePath(csAppDir)
50290   End If
50300  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "AnalyzeCommandlineParameters")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub WriteInstalldate2Registry()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry, d As Date
50020  d = Now
50030  Set reg = New clsRegistry
50040  With reg
50050   .hkey = HKEY_LOCAL_MACHINE
50060   .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50070   .SetRegistryValue "InstallDate", Year(d) & Format$(Month(d), "00") & Format(Day(d), "00"), REG_SZ
50080  End With
50090  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "WriteInstalldate2Registry")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

