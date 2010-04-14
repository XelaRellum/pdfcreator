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
         Name            =   "Arial Rounded MT Bold"
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
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   1920
      Width           =   5295
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
         Name            =   "Arial Rounded MT Bold"
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
         Name            =   "Arial Rounded MT Bold"
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
         Name            =   "Arial Rounded MT Bold"
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
         Name            =   "Arial Rounded MT Bold"
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
      X2              =   4800
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
         Name            =   "Arial Rounded MT Bold"
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
         Name            =   "Arial Rounded MT Bold"
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
         Name            =   "Arial Rounded MT Bold"
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
         Name            =   "Arial Rounded MT Bold"
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
         Name            =   "Arial Rounded MT Bold"
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
      Caption         =   "Special Thanks"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
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
         Name            =   "Arial Rounded MT Bold"
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
Unload Me
End Sub

Private Sub Form_Load()
lblTranslator.Caption = LanguageStrings.CommonLanguagename + " by " + LanguageStrings.CommonAuthor
imgDonate.MousePointer = 99
imgDonate.MouseIcon = LoadResPicture(1000, vbResCursor)
lblHomepage.MousePointer = 99
lblHomepage.MouseIcon = LoadResPicture(1000, vbResCursor)
Dim i As Integer
For i = 0 To lblLink.Count - 1
    lblLink(i).MousePointer = 99
    lblLink(i).MouseIcon = LoadResPicture(1000, vbResCursor)
Next
End Sub

Private Sub imgDonate_Click()
OpenDocument PaypalPDFCreator
End Sub

Private Sub lblHomepage_Click()
OpenDocument "http://www.pdfforge.org"
End Sub

Private Sub lblLink_Click(Index As Integer)
OpenDocument "http://" + lblLink(Index).Caption
End Sub
