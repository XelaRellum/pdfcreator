VERSION 5.00
Begin VB.Form frmPrintfiles 
   Caption         =   "PDFCreator"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   Icon            =   "frmPrintfiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows-Standard
   Visible         =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   1890
      TabIndex        =   2
      Top             =   1470
      Width           =   1905
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   105
      Top             =   1365
   End
   Begin PDFCreator.XP_ProgressBar xpPgb 
      Height          =   330
      Left            =   210
      TabIndex        =   0
      Top             =   960
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   6956042
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "1 (1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   2880
      TabIndex        =   4
      Top             =   720
      Width           =   405
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Size:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   210
      TabIndex        =   3
      Top             =   525
      Width           =   435
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "File:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   210
      TabIndex        =   1
      Top             =   210
      Width           =   375
   End
End
Attribute VB_Name = "frmPrintfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  CancelPrintfiles = True
50050  cmd.Enabled = False
50060 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50070 Exit Sub
ErrPtnr_OnError:
50091 Select Case ErrPtnr.OnError("frmPrintfiles", "cmd_Click")
      Case 0: Resume
50110 Case 1: Resume Next
50120 Case 2: Exit Sub
50130 Case 3: End
50140 End Select
50150 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Load()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim Files As Collection
50050  CancelPrintfiles = False
50060  Caption = App.Title & " " & GetProgramReleaseStr
50070  RemoveX Me
50080  Set Files = GetFiles(PrintFilename, "")
50090  If Files.Count > 1 Then
50100   Visible = True
50110  End If
50120  lbl(0).Caption = LanguageStrings.ListFilename
50130  lbl(1).Caption = LanguageStrings.ListSize
50140  With xpPgb
50150   .Min = 1
50160   .Max = Files.Count
50170   .Color = vbGreen
50180   .Font.Bold = True
50190   .ShowText = True
50200  End With
50210  Timer1.Enabled = True
50220 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50230 Exit Sub
ErrPtnr_OnError:
50251 Select Case ErrPtnr.OnError("frmPrintfiles", "Form_Load")
      Case 0: Resume
50270 Case 1: Resume Next
50280 Case 2: Exit Sub
50290 Case 3: End
50300 End Select
50310 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Resize()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If WindowState = vbMinimized Then
50050   Exit Sub
50060  End If
50070  xpPgb.Left = (Width - xpPgb.Width) / 2
50080  Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
50090  cmd.Move (Width - cmd.Width) / 2, cmd.Top
50100 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50110 Exit Sub
ErrPtnr_OnError:
50131 Select Case ErrPtnr.OnError("frmPrintfiles", "Form_Resize")
      Case 0: Resume
50150 Case 1: Resume Next
50160 Case 2: Exit Sub
50170 Case 3: End
50180 End Select
50190 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Timer1_Timer()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Timer1.Enabled = False
50050  PrintFile Me, PrintFilename, xpPgb, lbl(0), lbl(1), lbl(2)
50060  Unload Me
50070 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50080 Exit Sub
ErrPtnr_OnError:
50101 Select Case ErrPtnr.OnError("frmPrintfiles", "Timer1_Timer")
      Case 0: Resume
50120 Case 1: Resume Next
50130 Case 2: Exit Sub
50140 Case 3: End
50150 End Select
50160 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
