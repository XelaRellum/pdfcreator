VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "About"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Pfeil
   ScaleHeight     =   4950
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Label lbl 
      Alignment       =   2  'Zentriert
      Caption         =   "http://pdfcreator.sourceforge.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   6
      Left            =   2880
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   6
      Top             =   3720
      Width           =   3495
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Zentriert
      Caption         =   "Homepage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Index           =   5
      Left            =   2880
      MousePointer    =   1  'Pfeil
      TabIndex        =   5
      Top             =   3240
      Width           =   3495
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Zentriert
      Caption         =   "Frank Heindörfer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   4
      Left            =   2880
      MousePointer    =   1  'Pfeil
      TabIndex        =   4
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Zentriert
      Caption         =   "Philip Chinery"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   3
      Left            =   2880
      MousePointer    =   1  'Pfeil
      TabIndex        =   3
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Zentriert
      Caption         =   "Authors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Index           =   2
      Left            =   2880
      MousePointer    =   1  'Pfeil
      TabIndex        =   2
      Top             =   1920
      Width           =   3495
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Zentriert
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Index           =   1
      Left            =   2880
      MousePointer    =   1  'Pfeil
      TabIndex        =   1
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Zentriert
      Caption         =   "PDFCreator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Index           =   0
      Left            =   2880
      MousePointer    =   1  'Pfeil
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Image img 
      Height          =   4710
      Left            =   120
      MousePointer    =   1  'Pfeil
      Picture         =   "frmAbout.frx":08CA
      Top             =   120
      Width           =   2460
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
 With lbl(6)
  .MouseIcon = LoadResPicture(1000, vbResCursor)
  .MousePointer = vbCustom
 End With
End Sub

Private Sub lbl_DblClick(index As Integer)
 OpenDocument lbl(6).Caption
End Sub
