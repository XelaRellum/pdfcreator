Attribute VB_Name = "modShell"
Option Explicit

Private Const WAIT_FAILED = &HFFFFFFFF


Public Declare Function WaitForInputIdle Lib "user32" ( _
   ByVal hProcess As Long, _
   ByVal dwMilliseconds As Long _
) As Long

Public Declare Function WaitForSingleObject Lib "Kernel32" ( _
   ByVal hHandle As Long, _
   ByVal dwMilliseconds As Long _
) As Long

Private Declare Function TerminateProcess Lib "Kernel32" ( _
   ByVal hProcess As Long, _
   ByVal uExitCode As Long _
) As Long

Private Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" ( _
   lpExecInfo As SHELLEXECUTEINFO _
) As Long

Private Type SHELLEXECUTEINFO
   cbSize As Long
   fMask As Long
   hwnd As Long
   lpVerb As String
   lpFile As String
   lpParameters As String
   lpDirectory As String
   nShow As Long
   hInstApp As Long
   lpIDList As Long
   lpClass As String
   hkeyClass As Long
   dwHotKey As Long
   hIcon As Long
   hProcess As Long
End Type

'SHELLEXECUTEINFO fMask-Konstanten
Private Const SEE_MASK_NOCLOSEPROCESS = &H40 'Füllt die Struktur Option hProcess mit dem Process Handle der gestarteten Anwendung
Private Const WAIT_OBJECT_0 = &H0 'Das

'SHELLEXECUTEINFO nShow-Konstanten
Public Enum ShowConstants
   WVersteckt = 0    'Versteckt das Fenster
   WNormal = 1       'Zeigt es ganz normal an
   WMaximiert = 3    'Maximiert das Fenster
   WMinimiert = 6    'Minimiert das Fenster
End Enum

Public Enum WaitConstants
    wcNone = 0
    wcInitialisiert = 1
    WCTermination = 2
End Enum

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
50020    Dim RetVal As Long
50030    Dim ShExInfo As SHELLEXECUTEINFO
50040
50050    '//////////////////////////////////////////////////////////////////////////////
50060    ' Initialisierung
50070    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
50080
50090    If WorkingFolder = "" Then WorkingFolder = FilePath
50100
50110    With ShExInfo
50120       .cbSize = Len(ShExInfo)
50130       .fMask = SEE_MASK_NOCLOSEPROCESS
50140       .hwnd = 0
50150       .lpVerb = Operation
50160       .lpFile = FilePath
50170       .lpParameters = Parameter
50180       .lpDirectory = WorkingFolder
50190       .nShow = WindowSize
50200    End With
50210
50220    '/////////////////////////////////////////////////////////////////////////////
50230    ' Anwendung ausführen
50240    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
50250
50260    RetVal = ShellExecuteEx(ShExInfo)
50270
50280    If RetVal = 0 Then
50290        'Ein Fehler ist aufgetreten
50300 '       ShellAndWait = ShellExecError(ShExInfo.hInstApp)
50310        Exit Function
50320    End If
50330
50340    '/////////////////////////////////////////////////////////////////////////////
50350    ' Warten auf Prozess
50360    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
50370
50380    If WaitFor <> wcNone Then
50390
50400       If WaitFor = wcInitialisiert Then
50410          ' Warten bis die Anwendung initialisiert ist
50420          RetVal = WaitForInputIdle(ShExInfo.hProcess, Milliseconds)
50430
50440       Else
50450          ' Warten bis die Anwendung beendet
50460          RetVal = WaitForSingleObject(ShExInfo.hProcess, Milliseconds)
50470
50480       End If
50490
50500       If RetVal = WAIT_FAILED Then ShellAndWait = "Warten auf Prozess fehlgeschlagen."
50510
50520    End If
50530
50540    '/////////////////////////////////////////////////////////////////////////////
50550    ' SCHLIEßEN DES PROZESSES
50560    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
50570    If CloseProcess = True Then
50580       RetVal = TerminateProcess(ShExInfo.hProcess, 1)
50590       If RetVal <> 0 Then ShellAndWait = "Schließen der Anwendung fehlgeschlagen."
50600    End If
50610
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
