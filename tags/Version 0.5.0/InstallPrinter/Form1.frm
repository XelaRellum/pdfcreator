VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Drucker Installation"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows-Standard
   Visible         =   0   'False
   Begin VB.CommandButton cmdAll 
      Caption         =   "Alles Installieren"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmdDriver 
      Caption         =   "Treiber Installieren"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmdPort 
      Caption         =   "Port Installieren"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Text            =   "PDFCreator"
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Text            =   "PDFCreator"
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdPrinter 
      Caption         =   "Drucker Installieren"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Treiber:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Drucker:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAll_Click()
Dim a As Boolean, b As Boolean, c As Boolean
a = AddMyPort
b = AddDriver
c = InstallPrinter(Text2.Text, Text3.Text, App.Path & "\spool.ps", , "PDFCreator")
'If a = True And b = True And c = True Then MsgBox "Druckerinstallation erfolgreich abgeschlossen", vbOKOnly, "Druckerinstallation"
End
End Sub

Private Sub cmdPort_Click()
AddMyPort
End Sub

Private Sub cmdDriver_Click()
AddDriver
End Sub

Private Sub cmdPrinter_Click()
Label1.Caption = App.Path & "\spool.ps"
InstallPrinter Text2.Text, Text3.Text, App.Path & "\spool.ps", , "PDFCreator"
End Sub

Private Sub Form_Load()
cmdAll_Click
End Sub
