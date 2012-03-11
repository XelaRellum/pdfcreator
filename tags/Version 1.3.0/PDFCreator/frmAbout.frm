VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Info"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8370
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdClose 
      Caption         =   "OK"
      Height          =   495
      Left            =   6720
      TabIndex        =   12
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please donate to support us"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   13
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Label lblTranslator 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Your Translation is done by: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   1920
      Width           =   5535
   End
   Begin VB.Image imgDonate 
      Height          =   465
      Left            =   5520
      MousePointer    =   2  'Kreuz
      Picture         =   "frmAbout.frx":628A
      Top             =   4100
      Width           =   930
   End
   Begin VB.Label lblLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "www.activevb.de"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   10
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label lblLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "www.vb-hellfire.de"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   9
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label lblLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "vbnet.mvps.org"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   8
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Label lblLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "www.planet-source-code.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   2640
      TabIndex        =   7
      Top             =   3720
      Width           =   3015
   End
   Begin VB.Line Line3 
      X1              =   2640
      X2              =   5760
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line2 
      X1              =   2640
      X2              =   5040
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      X1              =   2640
      X2              =   3720
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "www.aboutvb.de"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   6
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "For Tutorials and Help"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "All translators, who help you to understand PDFCreator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   1680
      Width           =   5535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "The testers, who help us to find and solve bugs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   1440
      Width           =   5535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Authors"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Special Thanks to"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label lblHomepage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Philip Chinery && Frank Heindörfer - www.pdfforge.org"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   0
      Top             =   600
      Width           =   5535
   End
   Begin VB.Image Image1 
      Height          =   4710
      Left            =   0
      Picture         =   "frmAbout.frx":65EF
      Top             =   0
      Width           =   2460
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAbout", "cmdClose_Click")
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
50010  Dim i As Long
50020
50030  ChangeLanguage
50040
50050  If LanguageStrings.CommonLanguagename = "Deutsch" Or LanguageStrings.CommonLanguagename = "English" Then
50060    lblTranslator.Caption = ""
50070   Else
50080    lblTranslator.Caption = """" + LanguageStrings.CommonLanguagename + """ by " + LanguageStrings.CommonAuthor
50090  End If
50100
50110  imgDonate.MousePointer = 99
50120  imgDonate.MouseIcon = LoadResPicture(1000, vbResCursor)
50130  lblHomepage.MousePointer = 99
50140  lblHomepage.MouseIcon = LoadResPicture(1000, vbResCursor)
50150  For i = 0 To lblLink.Count - 1
50160   lblLink(i).MousePointer = 99
50170   lblLink(i).MouseIcon = LoadResPicture(1000, vbResCursor)
50180  Next i
50190
50200  With Options
50210   SetFontControls Me.Controls, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50220  End With
50230
50240  ShowAcceleratorsInForm Me, True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAbout", "Form_Load")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub imgDonate_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  OpenDocument PaypalPDFCreator
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAbout", "imgDonate_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub lblHomepage_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  OpenDocument "http://www.pdfforge.org"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAbout", "lblHomepage_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub lblLink_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  OpenDocument "http://" + lblLink(Index).Caption
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAbout", "lblLink_Click")
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
50020   Me.Caption = .DialogInfoTitle & " PDFCreator " & App.Major & "." & App.Minor & "." & App.Revision
50030  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAbout", "ChangeLanguage")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

