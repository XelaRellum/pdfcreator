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
         Caption         =   "Repeat:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblOwnerPass 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1095
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
         Caption         =   "Repeat:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblUserPass 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1095
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
 bSuccess = False
 bFinished = True: Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF1 Then
  KeyCode = 0
  Call HTMLHelp_ShowTopic("html\pdfsecurity.htm")
 End If
End Sub

Private Sub Form_Load()
 Me.KeyPreview = True
 With LanguageStrings
  Caption = .OptionsPDFEnterPasswords
  fraUserPass.Caption = .OptionsUserPass
  fraOwnerPass.Caption = .OptionsOwnerPass
  lblUserPass.Caption = .OptionsPDFSetPassword
  lblUserPassRepeat.Caption = .OptionsPDFRepeatPassword
  lblOwnerPass.Caption = .OptionsPDFSetPassword
  lblOwnerPassRepeat.Caption = .OptionsPDFRepeatPassword
  OKButton.Caption = .OptionsPassOK
  CancelButton.Caption = .OptionsPassCancel
  chkSavePasswords.Caption = .OptionsSavePasswords
 End With
 With Options
  fraUserPass.Enabled = .PDFUserPass
  fraOwnerPass.Enabled = .PDFOwnerPass
 End With
 bSuccess = False
 bFinished = False
End Sub

Private Sub OKButton_Click()
 Dim aw As Long
 If txtUserPass.Text <> txtUserPassRepeat.Text Then
  MsgBox LanguageStrings.MessagesMsg21, vbCritical
  Exit Sub
 End If

 If txtOwnerPass.Text <> txtOwnerPassRepeat.Text Then
  MsgBox LanguageStrings.MessagesMsg22, vbCritical
  Exit Sub
 End If

 If Len(txtUserPass.Text) = 0 And fraUserPass.Enabled = True Then
  aw = MsgBox(LanguageStrings.MessagesMsg24, vbQuestion + vbYesNo)
  If aw = vbNo Then
   Exit Sub
  End If
 End If

 If Len(txtOwnerPass.Text) = 0 And fraOwnerPass.Enabled = True Then
  aw = MsgBox(LanguageStrings.MessagesMsg25, vbQuestion + vbYesNo)
  If aw = vbNo Then
   Exit Sub
  End If
 End If

 If chkSavePasswords.Value = 1 Then
   SavePasswordsForThisSession = True
  Else
   SavePasswordsForThisSession = False
 End If

 If Len(txtOwnerPass.Text) = 0 And Len(txtUserPass.Text) = 0 Then
  SavePasswordsForThisSession = False
 End If
 OwnerPassword = txtOwnerPass.Text: UserPassword = txtUserPass.Text
 bSuccess = True
 bFinished = True: Unload Me
End Sub

