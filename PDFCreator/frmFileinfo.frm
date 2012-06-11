VERSION 5.00
Begin VB.Form frmFileinfo 
   BorderStyle     =   4  'Festes Werkzeugfenster
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3810
   Icon            =   "frmFileinfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox txt 
      Appearance      =   0  '2D
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   210
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   105
      Width           =   3375
   End
End
Attribute VB_Name = "frmFileinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public InfoSpoolFileName As String

Private Sub Form_Activate()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim isf As clsInfoSpoolFile, isfi As clsInfoSpoolFileInfo, i As Long, File As String, strInfo As String, _
  sumFileSizes As Long
50030
50040  SplitPath InfoSpoolFileName, , , , File
50050  Me.Caption = File
50060
50070  Set isf = New clsInfoSpoolFile
50080  isf.ReadInfoFile InfoSpoolFileName
50090  If isf.InfoFiles.Count > 1 Then
50100   sumFileSizes = 0
50110   For i = 1 To isf.InfoFiles.Count
50120    Set isfi = isf.InfoFiles(i)
50130    If FileExists(isfi.SpoolFileName) Then
50140     sumFileSizes = sumFileSizes + GetFileLength(isfi.SpoolFileName)
50150    End If
50160   Next i
50170   strInfo = "Count of Jobs: " & CStr(isf.InfoFiles.Count)
50180   strInfo = strInfo & vbCrLf & "Sum file sizes: " & Format$(sumFileSizes, "#,##0") & " bytes"
50190   If LenB(txt.Text) = 0 Then
50200     txt.Text = strInfo
50210    Else
50220     txt.Text = txt.Text & vbCrLf & vbCrLf & strInfo
50230   End If
50240  End If
50250  For i = 1 To isf.InfoFiles.Count
50260   Set isfi = isf.InfoFiles(i)
50270   strInfo = "Computer: " & isfi.ClientComputer
50280   strInfo = strInfo & vbCrLf & "Document title: " & isfi.DocumentTitle
50290   strInfo = strInfo & vbCrLf & "JobID: " & isfi.JobID
50300   strInfo = strInfo & vbCrLf & "Printer name: " & isfi.PrinterName
50310   strInfo = strInfo & vbCrLf & "SessionID: " & isfi.SessionID
50320   strInfo = strInfo & vbCrLf & "Spool-filename: " & isfi.SpoolFileName
50330   If FileExists(isfi.SpoolFileName) Then
50340     strInfo = strInfo & vbCrLf & "Spool-file size: " & Format$(GetFileLength(isfi.SpoolFileName), "#,##0") & " bytes"
50350    Else
50360     strInfo = strInfo & vbCrLf & "Spool-file size: File doesn't exist anymore!"
50370   End If
50380   strInfo = strInfo & vbCrLf & "Username: " & isfi.UserName
50390   strInfo = strInfo & vbCrLf & "WinStation: " & isfi.WinStation
50400   If LenB(txt.Text) = 0 Then
50410     txt.Text = strInfo
50420    Else
50430     txt.Text = txt.Text & vbCrLf & vbCrLf & strInfo
50440   End If
50450  Next i
50460  Me.Width = Me.TextWidth(txt.Text) + 100
50470  Me.Height = Me.Height - Me.ScaleHeight + Me.TextHeight(txt.Text) + 100
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmFileinfo", "Form_Activate")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Load()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txt.Left = 0
50020  txt.Top = 0
50030  Set Me.Font = txt.Font
50040  With Options
50050   SetFontControls Me.Controls, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50060  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmFileinfo", "Form_Load")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Resize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txt.Width = Me.ScaleWidth
50020  txt.Height = Me.ScaleHeight
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmFileinfo", "Form_Resize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub ChangeLanguage()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  'Dummy
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmFileinfo", "ChangeLanguage")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
