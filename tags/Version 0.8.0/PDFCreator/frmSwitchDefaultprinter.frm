VERSION 5.00
Begin VB.Form frmSwitchDefaultprinter 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Form1"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "frmSwitchDefaultprinter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   495
      Index           =   1
      Left            =   3240
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CheckBox chkAskAgain 
      Caption         =   "Don't ask me again."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label lblSwitchDefaultprinter 
      Caption         =   "It is necessary to temporarily set PDFCreator as defaultprinter."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmSwitchDefaultprinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click(Index As Integer)
 Select Case Index
  Case 0:
   ChangeDefaultprinter = True
   Options.NoConfirmMessageSwitchingDefaultprinter = chkAskAgain.Value
   SaveOptions Options
  Case 1:
 End Select
 Unload Me
End Sub

Private Sub Form_Load()
 Caption = App.EXEName
 ChangeDefaultprinter = False
 With LanguageStrings
  lblSwitchDefaultprinter.Caption = .MessagesMsg35
  chkAskAgain.Caption = .MessagesMsg36
 End With
 chkAskAgain.Value = Options.NoConfirmMessageSwitchingDefaultprinter
End Sub
