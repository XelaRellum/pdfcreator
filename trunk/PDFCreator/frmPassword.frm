VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Enter Passwords"
   ClientHeight    =   2790
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4695
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraOwnerPass 
      Caption         =   "Owner Password"
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   2775
      Begin VB.TextBox txtOwnerPassRepeat 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   680
         Width           =   1695
      End
      Begin VB.TextBox txtOwnerPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label lblOwnerPassRepeat 
         Caption         =   "Repeat:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblOwnerPass 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame fraUserPass 
      Caption         =   "User Password"
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2775
      Begin VB.TextBox txtUserPassRepeat 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   660
         Width           =   1695
      End
      Begin VB.TextBox txtUserPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label lblUserPassRepeat 
         Caption         =   "Repeat:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblUserPass 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bSuccess As Boolean
Public bFinished As Boolean
Public iPasswords As Integer

Private Sub CancelButton_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 bSuccess = False
50020 bFinished = True
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

Private Sub Form_Load()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 Me.Caption = LanguageStrings.OptionsPDFEnterPasswords
50020 fraUserPass.Caption = LanguageStrings.OptionsUserPass
50030 fraOwnerPass.Caption = LanguageStrings.OptionsOwnerPass
50040 lblUserPass.Caption = LanguageStrings.OptionsPDFSetPassword
50050 lblUserPassRepeat.Caption = LanguageStrings.OptionsPDFRepeatPassword
50060 lblOwnerPass.Caption = LanguageStrings.OptionsPDFSetPassword
50070 lblOwnerPassRepeat.Caption = LanguageStrings.OptionsPDFRepeatPassword
50080 OKButton.Caption = LanguageStrings.OptionsPassOK
50090 CancelButton.Caption = LanguageStrings.OptionsPassCancel
50100 fraUserPass.Enabled = Options.PDFUserPass
50110 fraOwnerPass.Enabled = Options.PDFOwnerPass
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
50010 If txtUserPass.Text <> txtUserPassRepeat.Text Then
50020   MsgBox "User Passes do not match", vbCritical
50030   Exit Sub
50040 End If
50050
50060 If txtOwnerPass.Text <> txtOwnerPassRepeat.Text Then
50070   MsgBox "Owner Passes do not match", vbCritical
50080   Exit Sub
50090 End If
50100
50110 bSuccess = True
50120 bFinished = True
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


