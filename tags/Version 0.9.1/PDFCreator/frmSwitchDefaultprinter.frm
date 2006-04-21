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
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CheckBox chkAskAgain 
      Appearance      =   0  '2D
      Caption         =   "Don't ask me again."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4455
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Index           =   1
      Left            =   3180
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblSwitchDefaultprinter 
      AutoSize        =   -1  'True
      Caption         =   "It is necessary to temporarily set PDFCreator as defaultprinter."
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4545
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSwitchDefaultprinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case Index
        Case 0:
50030    ChangeDefaultprinter = True
50040    Options.NoConfirmMessageSwitchingDefaultprinter = chkAskAgain.Value
50050    SaveOptions Options
50060   Case 1:
50070  End Select
50080  Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSwitchDefaultprinter", "cmd_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Initialize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim f As Form
50020  For Each f In Forms
50030   If UCase$(f.Name) = "FRMMAIN" Then
50040    f.ShowFrmMain
50050    Exit Sub
50060   End If
50070  Next
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSwitchDefaultprinter", "Form_Initialize")
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
50020  Caption = App.EXEName
50030  ChangeDefaultprinter = False
50040  With LanguageStrings
50050   lblSwitchDefaultprinter.Caption = .MessagesMsg35
50060   chkAskAgain.Caption = .MessagesMsg36
50070  End With
50080  chkAskAgain.Value = Options.NoConfirmMessageSwitchingDefaultprinter
50090  ShowAcceleratorsInForm Me, True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSwitchDefaultprinter", "Form_Load")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
