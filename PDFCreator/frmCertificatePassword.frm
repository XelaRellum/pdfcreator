VERSION 5.00
Begin VB.Form frmCertificatePassword 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Certificate password"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5055
   Icon            =   "frmCertificatePassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin PDFCreator.dmFrame dmFraCerticatePasword 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3836
      Caption         =   "Certificate password"
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
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   375
         Left            =   3360
         TabIndex        =   4
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CheckBox chkShowCertificatePassword 
         Appearance      =   0  '2D
         Caption         =   "Show password"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox txtCertificatePassword 
         Appearance      =   0  '2D
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   480
         Width           =   3975
      End
      Begin VB.Image imgValidCertificatePassword 
         Height          =   240
         Left            =   4440
         Picture         =   "frmCertificatePassword.frx":014A
         Top             =   720
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgInvalidCertificatePassword 
         Height          =   240
         Left            =   4320
         Picture         =   "frmCertificatePassword.frx":06D4
         Top             =   505
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmCertificatePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ScreenMousePointer As Long
Private x509 As Object
Public certFilename As String

Public Sub ChangeLanguage()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With LanguageStrings
50020   Me.Caption = .OptionsPDFSigningEnterCerticatePassword
50030   dmFraCerticatePasword.Caption = .OptionsPDFSigningCerticatePassword
50040   chkShowCertificatePassword.Caption = .OptionsPDFSigningCerticatePasswordShowPassword
50050   cmdOk.Caption = .OptionsPDFSigningCerticatePasswordOk
50060   cmdCancel.Caption = .OptionsPDFSigningCerticatePasswordCancel
50070  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmCertificatePassword", "ChangeLanguage")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkShowCertificatePassword_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkShowCertificatePassword.value <> 1 Then
50020    txtCertificatePassword.PasswordChar = "*"
50030   Else
50040    txtCertificatePassword.PasswordChar = ""
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmCertificatePassword", "chkShowCertificatePassword_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdCancel_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmCertificatePassword", "cmdCancel_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdOK_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If LenB(txtCertificatePassword.Text) > 0 Then
50020   PFXPassword = txtCertificatePassword.Text
50030  End If
50040  Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmCertificatePassword", "cmdOK_Click")
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
50010  ChangeLanguage
50020  ScreenMousePointer = Screen.MousePointer
50030  Screen.MousePointer = vbDefault
50040  Set x509 = CreateObject("pdfforge.X509.X509")
50050  imgValidCertificatePassword.Left = imgInvalidCertificatePassword.Left
50060  imgValidCertificatePassword.Top = imgInvalidCertificatePassword.Top
50070
50080  With Options
50090   SetFontControls Me.Controls, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50100  End With
50110
50120  ShowAcceleratorsInForm Me, True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmCertificatePassword", "Form_Load")
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
50010  Set x509 = Nothing
50020  Screen.MousePointer = ScreenMousePointer
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmCertificatePassword", "Form_Unload")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetIcon(value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If value = True Then
50020    If imgValidCertificatePassword.Visible = False Then
50030     imgValidCertificatePassword.Visible = True
50040     imgInvalidCertificatePassword.Visible = False
50050    End If
50060   Else
50070    If imgInvalidCertificatePassword.Visible = False Then
50080     imgInvalidCertificatePassword.Visible = True
50090     imgValidCertificatePassword.Visible = False
50100    End If
50110  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmCertificatePassword", "SetIcon")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtCertificatePassword_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If LenB(txtCertificatePassword.Text) = 0 Then
50020    cmdOk.Enabled = False
50030    SetIcon False
50040   Else
50050    If x509.IsValidCertificatePassword(certFilename, txtCertificatePassword.Text) Then
50060      cmdOk.Enabled = True
50070      SetIcon True
50080     Else
50090      cmdOk.Enabled = False
50100      SetIcon False
50110    End If
50120  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmCertificatePassword", "txtCertificatePassword_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
