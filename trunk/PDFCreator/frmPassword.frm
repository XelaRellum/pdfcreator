VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Enter Passwords"
   ClientHeight    =   3495
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4530
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
   End
   Begin PDFCreator.dmFrame dmFraOwnerPass 
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   4335
      _extentx        =   7646
      _extenty        =   1931
      caption         =   "Owner password"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmPassword.frx":000C
      Begin VB.CheckBox chkShowOwnerPasswordChars 
         Caption         =   "Show password"
         Height          =   195
         Left            =   1200
         TabIndex        =   10
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtOwnerPass 
         Appearance      =   0  '2D
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label lblOwnerPass 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Width           =   735
      End
   End
   Begin PDFCreator.dmFrame dmFraUserPass 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _extentx        =   7646
      _extenty        =   1931
      caption         =   "User password"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmPassword.frx":0038
      Begin VB.CheckBox chkShowUserPasswordChars 
         Caption         =   "Show password"
         Height          =   195
         Left            =   1200
         TabIndex        =   9
         Top             =   720
         Width           =   3015
      End
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
      Begin VB.Label lblUserPass 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.CheckBox chkSavePasswords 
      Appearance      =   0  '2D
      Caption         =   "Save passwords for this session."
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   4215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   3000
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

Private Sub chkShowOwnerPasswordChars_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkShowOwnerPasswordChars.value = 1 Then
50020    txtOwnerPass.PasswordChar = ""
50030   Else
50040    txtOwnerPass.PasswordChar = "*"
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPassword", "chkShowOwnerPasswordChars_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkShowUserPasswordChars_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkShowUserPasswordChars.value = 1 Then
50020    txtUserPass.PasswordChar = ""
50030   Else
50040    txtUserPass.PasswordChar = "*"
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPassword", "chkShowUserPasswordChars_Click")
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
50030   Call HTMLHelp_ShowTopic("html\security.html")
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
50120
50130  With Options
50140   SetFontControls Me.Controls, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50150  End With
50160
50170  ShowAcceleratorsInForm Me, True
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
50020
50030  If Len(txtUserPass.Text) = 0 And dmFraUserPass.Enabled = True Then
50040   aw = MsgBox(LanguageStrings.MessagesMsg24, vbQuestion + vbYesNo)
50050   If aw = vbNo Then
50060    Exit Sub
50070   End If
50080  End If
50090
50100  If Len(txtOwnerPass.Text) = 0 And dmFraOwnerPass.Enabled = True Then
50110   aw = MsgBox(LanguageStrings.MessagesMsg25, vbQuestion + vbYesNo)
50120   If aw = vbNo Then
50130    Exit Sub
50140   End If
50150  End If
50160
50170  If chkSavePasswords.value = 1 Then
50180    SavePasswordsForThisSession = True
50190   Else
50200    SavePasswordsForThisSession = False
50210  End If
50220
50230  If Len(txtOwnerPass.Text) = 0 And Len(txtUserPass.Text) = 0 Then
50240   SavePasswordsForThisSession = False
50250  End If
50260  OwnerPassword = txtOwnerPass.Text: UserPassword = txtUserPass.Text
50270  bSuccess = True
50280  bFinished = True: Unload Me
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
50060   chkShowUserPasswordChars.Caption = .OptionsPDFUserPasswordShowChars
50070   chkShowOwnerPasswordChars.Caption = .OptionsPDFOwnerPasswordShowChars
50080   lblOwnerPass.Caption = .OptionsPDFSetPassword
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
