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

   Dim RetVal As Long
   Dim ShExInfo As SHELLEXECUTEINFO

   '//////////////////////////////////////////////////////////////////////////////
   ' Initialisierung
   '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

   If WorkingFolder = "" Then WorkingFolder = FilePath

   With ShExInfo
      .cbSize = Len(ShExInfo)
      .fMask = SEE_MASK_NOCLOSEPROCESS
      .hwnd = 0
      .lpVerb = Operation
      .lpFile = FilePath
      .lpParameters = Parameter
      .lpDirectory = WorkingFolder
      .nShow = WindowSize
   End With

   '/////////////////////////////////////////////////////////////////////////////
   ' Anwendung ausführen
   '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

   RetVal = ShellExecuteEx(ShExInfo)

   If RetVal = 0 Then
       'Ein Fehler ist aufgetreten
'       ShellAndWait = ShellExecError(ShExInfo.hInstApp)
       Exit Function
   End If

   '/////////////////////////////////////////////////////////////////////////////
   ' Warten auf Prozess
   '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

   If WaitFor <> wcNone Then

      If WaitFor = wcInitialisiert Then
         ' Warten bis die Anwendung initialisiert ist
         RetVal = WaitForInputIdle(ShExInfo.hProcess, Milliseconds)

      Else
         ' Warten bis die Anwendung beendet
         RetVal = WaitForSingleObject(ShExInfo.hProcess, Milliseconds)

      End If

      If RetVal = WAIT_FAILED Then ShellAndWait = "Warten auf Prozess fehlgeschlagen."

   End If

   '/////////////////////////////////////////////////////////////////////////////
   ' SCHLIEßEN DES PROZESSES
   '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
   If CloseProcess = True Then
      RetVal = TerminateProcess(ShExInfo.hProcess, 1)
      If RetVal <> 0 Then ShellAndWait = "Schließen der Anwendung fehlgeschlagen."
   End If

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
