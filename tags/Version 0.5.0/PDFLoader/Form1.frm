VERSION 5.00
Begin VB.Form frmLoader 
   Caption         =   "PDFLoader"
   ClientHeight    =   855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   855
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows-Standard
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "frmLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTempPath Lib "KERNEL32" Alias _
  "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer _
  As String) As Long
Private Declare Function GetTempFileName Lib "KERNEL32" Alias _
  "GetTempFileNameA" (ByVal lpszPath As String, ByVal _
  lpPrefixString As String, ByVal wUnique As Long, ByVal _
  lpTempFileName As String) As Long

Private Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)

Private Const MAX_PATH = 260
Dim strTempDir As String

Private AppPath As String

Private Sub Form_Load()
AppPath = App.Path
   
  Dim lRet As Long

  strTempDir = Space(MAX_PATH)
  lRet = GetTempPath(MAX_PATH, strTempDir)
  strTempDir = Left(strTempDir, lRet)
End Sub

Private Sub Timer1_Timer()
On Error GoTo eTrap:
Dim Random
If Dir$(AppPath & "\spool.ps") <> "" Then
Timer1.Interval = 20

  Dim lRet As Long
  Dim strTempFile As String

  strTempFile = Space(MAX_PATH)
  lRet = GetTempFileName(strTempDir, "Est", 0, strTempFile)
  If lRet <> 0 Then
    strTempFile = Left(strTempFile, InStr(strTempFile, Chr(0)) - 1) & ".ps"
  End If
  
Name AppPath & "\spool.ps" As strTempFile
Shell AppPath & "\PDFCreator.exe -f" & strTempFile, vbNormalFocus
Timer1.Interval = 1000
End If
Exit Sub

eTrap:
If Dir$(strTempFile) <> "" Then Kill strTempFile
Exit Sub

End Sub
