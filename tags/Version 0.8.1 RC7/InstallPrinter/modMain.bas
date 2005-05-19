Attribute VB_Name = "modMain"
Option Explicit

Private Enum eInstallAction
 Nil = 0
 Install = 1
 UnInstall = 2
End Enum

Private InstallAction As eInstallAction, InstallSpecialPrinter As Boolean

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
50070  AnalyzeCommandlineParameters
50080  If InstallAction = Install Then
50090   LogFile = CompletePath(App.Path) & "SetupLog.txt"
50100   If InstallSpecialPrinter = False Then
50110     InstallWindowsPrinter Monitorname, Portname, Drivername, Printername, LogFile
50120     WriteInstalldate2Registry
50130    Else
50140     InstallAdditionalWindowsPrinter Printername, LogFile
50150   End If
50160  End If
50170
50180  If InstallAction = UnInstall Then
50190   LogFile = CompletePath(GetTempPathApi) & "PDFCreatorUninstall.txt"
50200   UnInstallWindowsPrinter Monitorname, Portname, Drivername, Printername, LogFile
50210  End If
50220  If InstallAction = Install Or InstallAction = UnInstall Then
50230   WriteToLog "--------------------------------------------------", LogFile
50240   WriteToLog "MSI-Installer: Installer2Go", LogFile
50250   WriteToLog "Computername: " & GetComputerName, LogFile
50260   WriteToLog "Username: " & GetUsername, LogFile
50270   WriteToLog "WinDir: " & GetWindowsDirectory, LogFile
50280   WriteToLog "SysDir: " & GetSystemDirectory, LogFile
50290   WriteToLog "TempDir: " & GetTempPathApi, LogFile
50300   WriteToLog "CurrentDir: " & CurDir, LogFile
50310   WriteToLog "Path: " & Environ$("Path"), LogFile
50320   WriteToLog "Internet Explorer version: " & GetIExplorerVersion, LogFile
50330  End If
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
50010  If Len(VBA.Command$) > 0 Then
50020   If UCase$(CommandSwitch("IN", False)) = "STALL" And _
     UCase$(CommandSwitch("UNIN", False)) = "STALL" Then
50040    MsgBox "Don't use INSTALL and UNINSTALL at the same time!" & vbCrLf & _
    "Program canceled!"
50060   End If
50070   InstallAction = Nil
50080   If UCase$(CommandSwitch("IN", False)) = "STALL" Then
50090    InstallAction = Install
50100   End If
50110   If UCase$(CommandSwitch("UNIN", False)) = "STALL" Then
50120    InstallAction = UnInstall
50130   End If
50140   If LenB(CommandSwitch("PRINTERNAME", True)) > 0 Then
50150    InstallSpecialPrinter = True
50160    Printername = CommandSwitch("PRINTERNAME", True)
50170   End If
50180  End If
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

