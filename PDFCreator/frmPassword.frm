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
   Begin VB.Frame fraOwnerPass 
      Caption         =   "Owner Password"
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   4335
      Begin VB.TextBox txtOwnerPassRepeat 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtOwnerPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   300
         Width           =   3015
      End
      Begin VB.Label lblOwnerPassRepeat 
         AutoSize        =   -1  'True
         Caption         =   "Repeat:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   570
      End
      Begin VB.Label lblOwnerPass 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame fraUserPass 
      Caption         =   "User Password"
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtUserPassRepeat 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   660
         Width           =   3015
      End
      Begin VB.TextBox txtUserPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblUserPassRepeat 
         AutoSize        =   -1  'True
         Caption         =   "Repeat:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   570
      End
      Begin VB.Label lblUserPass 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CheckBox chkSavePasswords 
      Caption         =   "Save passwords for this session."
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   4215
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bSuccess As Boolean, bFinished As Boolean, iPasswords As Integer

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
50010  Me.KeyPreview = True
50020  Me.Icon = frmMain.Icon
50030  With LanguageStrings
50040   Caption = .OptionsPDFEnterPasswords
50050   fraUserPass.Caption = .OptionsUserPass
50060   fraOwnerPass.Caption = .OptionsOwnerPass
50070   lblUserPass.Caption = .OptionsPDFSetPassword
50080   lblUserPassRepeat.Caption = .OptionsPDFRepeatPassword
50090   lblOwnerPass.Caption = .OptionsPDFSetPassword
50100   lblOwnerPassRepeat.Caption = .OptionsPDFRepeatPassword
50110   OKButton.Caption = .OptionsPassOK
50120   CancelButton.Caption = .OptionsPassCancel
50130   chkSavePasswords.Caption = .OptionsSavePasswords
50140  End With
50150  With Options
50160   fraUserPass.Enabled = .PDFUserPass
50170   fraOwnerPass.Enabled = .PDFOwnerPass
50180  End With
50190  bSuccess = False
50200  bFinished = False
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
50120  If Len(txtUserPass.Text) = 0 And fraUserPass.Enabled = True Then
50130   aw = MsgBox(LanguageStrings.MessagesMsg24, vbQuestion + vbYesNo)
50140   If aw = vbNo Then
50150    Exit Sub
50160   End If
50170  End If
50180
50190  If Len(txtOwnerPass.Text) = 0 And fraOwnerPass.Enabled = True Then
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

