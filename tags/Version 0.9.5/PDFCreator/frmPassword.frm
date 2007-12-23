VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Enter Passwords"
   ClientHeight    =   3675
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4530
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3240
      Width           =   1215
   End
   Begin PDFCreator.dmFrame dmFraOwnerPass 
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2143
      Caption         =   "Owner password"
      BarColorFrom    =   16744576
      BarColorTo      =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtOwnerPass 
         Appearance      =   0  '2D
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtOwnerPassRepeat 
         Appearance      =   0  '2D
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   780
         Width           =   3015
      End
      Begin VB.Label lblOwnerPass 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   420
         Width           =   735
      End
      Begin VB.Label lblOwnerPassRepeat 
         AutoSize        =   -1  'True
         Caption         =   "Repeat:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   780
         Width           =   570
      End
   End
   Begin PDFCreator.dmFrame dmFraUserPass 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2143
      Caption         =   "User password"
      BarColorFrom    =   16744576
      BarColorTo      =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtUserPass 
         Appearance      =   0  '2D
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtUserPassRepeat 
         Appearance      =   0  '2D
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   780
         Width           =   3015
      End
      Begin VB.Label lblUserPass 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblUserPassRepeat 
         AutoSize        =   -1  'True
         Caption         =   "Repeat:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   570
      End
   End
   Begin VB.CheckBox chkSavePasswords 
      Appearance      =   0  '2D
      Caption         =   "Save passwords for this session."
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   4215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   3240
      Width           =   1215
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bSuccess As Boolean, bFinished As Boolean, iPasswords As Integer

Private oldMousePointer As Long

Private Sub CancelButton_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  bSuccess = False
50020  bFinished = True: Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPassword", "CancelButton_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If KeyCode = vbKeyF1 Then
50020   KeyCode = 0
50030   Call HTMLHelp_ShowTopic("html\pdfsecurity.htm")
50040  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPassword", "Form_KeyDown")
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
50010  Me.Icon = LoadResPicture(2120, vbResIcon)
50020  oldMousePointer = Screen.MousePointer
50030  Me.KeyPreview = True
50040  Screen.MousePointer = vbNormal
50050  ChangeLanguage
50060  With Options
50070   dmFraUserPass.Enabled = .PDFUserPass
50080   dmFraOwnerPass.Enabled = .PDFOwnerPass
50090  End With
50100  bSuccess = False
50110  bFinished = False
50120  ShowAcceleratorsInForm Me, True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPassword", "Form_Load")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Unload(Cancel As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Screen.MousePointer = oldMousePointer
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPassword", "Form_Unload")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub OKButton_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim aw As Long
50020  If txtUserPass.Text <> txtUserPassRepeat.Text Then
50030   MsgBox LanguageStrings.MessagesMsg21, vbCritical
50040   Exit Sub
50050  End If
50060
50070  If txtOwnerPass.Text <> txtOwnerPassRepeat.Text Then
50080   MsgBox LanguageStrings.MessagesMsg22, vbCritical
50090   Exit Sub
50100  End If
50110
50120  If Len(txtUserPass.Text) = 0 And dmFraUserPass.Enabled = True Then
50130   aw = MsgBox(LanguageStrings.MessagesMsg24, vbQuestion + vbYesNo)
50140   If aw = vbNo Then
50150    Exit Sub
50160   End If
50170  End If
50180
50190  If Len(txtOwnerPass.Text) = 0 And dmFraOwnerPass.Enabled = True Then
50200   aw = MsgBox(LanguageStrings.MessagesMsg25, vbQuestion + vbYesNo)
50210   If aw = vbNo Then
50220    Exit Sub
50230   End If
50240  End If
50250
50260  If chkSavePasswords.Value = 1 Then
50270    SavePasswordsForThisSession = True
50280   Else
50290    SavePasswordsForThisSession = False
50300  End If
50310
50320  If Len(txtOwnerPass.Text) = 0 And Len(txtUserPass.Text) = 0 Then
50330   SavePasswordsForThisSession = False
50340  End If
50350  OwnerPassword = txtOwnerPass.Text: UserPassword = txtUserPass.Text
50360  bSuccess = True
50370  bFinished = True: Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPassword", "OKButton_Click")
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
50010  With LanguageStrings
50020   Caption = .OptionsPDFEnterPasswords
50030   dmFraUserPass.Caption = .OptionsUserPass
50040   dmFraOwnerPass.Caption = .OptionsOwnerPass
50050   lblUserPass.Caption = .OptionsPDFSetPassword
50060   lblUserPassRepeat.Caption = .OptionsPDFRepeatPassword
50070   lblOwnerPass.Caption = .OptionsPDFSetPassword
50080   lblOwnerPassRepeat.Caption = .OptionsPDFRepeatPassword
50090   OKButton.Caption = .OptionsPassOK
50100   CancelButton.Caption = .OptionsPassCancel
50110   chkSavePasswords.Caption = .OptionsSavePasswords
50120  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPassword", "ChangeLanguage")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
