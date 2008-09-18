VERSION 5.00
Begin VB.Form frmFileinfo 
   BorderStyle     =   4  'Festes Werkzeugfenster
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3810
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

Public SpoolFilename As String

Private Sub Form_Activate()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim isf As InfoSpoolFile, File As String
50020  SplitPath SpoolFilename, , , , File
50030  Me.Caption = File
50040  isf = ReadInfoSpoolfile(SpoolFilename)
50050  txt.Text = "Computer:         " & isf.Computer
50060  txt.Text = txt.Text & vbCrLf & "Created:          " & isf.Created
50070 ' txt.Text = txt.Text & vbCrLf & "SpoolFilename:    " & isf.SpoolFilename
50080  txt.Text = txt.Text & vbCrLf & "SpoolerAccount:   " & isf.SpoolerAccount
50090  txt.Text = txt.Text & vbCrLf
50100  txt.Text = txt.Text & vbCrLf & "REDMON_DOCNAME:   " & isf.REDMON_DOCNAME
50110  txt.Text = txt.Text & vbCrLf & "REDMON_JOB:       " & isf.REDMON_JOB
50120  txt.Text = txt.Text & vbCrLf & "REDMON_MACHINE:   " & isf.REDMON_MACHINE
50130  txt.Text = txt.Text & vbCrLf & "REDMON_PORT:      " & isf.REDMON_PORT
50140  txt.Text = txt.Text & vbCrLf & "REDMON_PRINTER:   " & isf.REDMON_PRINTER
50150  txt.Text = txt.Text & vbCrLf & "REDMON_SESSIONID: " & isf.REDMON_SESSIONID
50160  txt.Text = txt.Text & vbCrLf & "REDMON_USER:      " & isf.REDMON_USER
50170  Me.Width = Me.TextWidth(txt.Text) + 100
50180  Me.Height = Me.Height - Me.ScaleHeight + Me.TextHeight(txt.Text) + 100
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
