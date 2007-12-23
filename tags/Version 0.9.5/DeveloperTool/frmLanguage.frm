VERSION 5.00
Begin VB.Form frmLanguage 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Language"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   Icon            =   "frmLanguage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdEngClp 
      Height          =   285
      Index           =   1
      Left            =   6645
      Picture         =   "frmLanguage.frx":0E42
      Style           =   1  'Grafisch
      TabIndex        =   15
      ToolTipText     =   "Copy from clipboard"
      Top             =   1890
      Width           =   330
   End
   Begin VB.CommandButton cmdEngClp 
      Height          =   285
      Index           =   0
      Left            =   6195
      Picture         =   "frmLanguage.frx":11CC
      Style           =   1  'Grafisch
      TabIndex        =   14
      ToolTipText     =   "Copy to clipboard"
      Top             =   1890
      Width           =   330
   End
   Begin VB.CommandButton cmdGerClp 
      Height          =   285
      Index           =   1
      Left            =   6645
      Picture         =   "frmLanguage.frx":1556
      Style           =   1  'Grafisch
      TabIndex        =   13
      ToolTipText     =   "Copy from clipboard"
      Top             =   2625
      Width           =   330
   End
   Begin VB.CommandButton cmdGerClp 
      Height          =   285
      Index           =   0
      Left            =   6195
      Picture         =   "frmLanguage.frx":18E0
      Style           =   1  'Grafisch
      TabIndex        =   12
      ToolTipText     =   "Copy to clipboard"
      Top             =   2625
      Width           =   330
   End
   Begin VB.CommandButton cmdKeyClp 
      Height          =   285
      Index           =   1
      Left            =   6645
      Picture         =   "frmLanguage.frx":1C6A
      Style           =   1  'Grafisch
      TabIndex        =   11
      ToolTipText     =   "Copy from clipboard"
      Top             =   1200
      Width           =   330
   End
   Begin VB.CommandButton cmdKeyClp 
      Height          =   285
      Index           =   0
      Left            =   6195
      Picture         =   "frmLanguage.frx":1FF4
      Style           =   1  'Grafisch
      TabIndex        =   10
      ToolTipText     =   "Copy to clipboard"
      Top             =   1200
      Width           =   330
   End
   Begin VB.ComboBox cmbSection 
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   6855
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Cancel"
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Save"
      Height          =   495
      Index           =   0
      Left            =   5760
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   5940
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   5940
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   5940
   End
   Begin VB.Label lbl 
      Caption         =   "German"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "English"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "Key"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "Section"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click(Index As Integer)
 Dim item As ListItem, Str1 As String, Str2 As String, Str3 As String, _
  Str4 As String, aw As Long
 Select Case Index
  Case 0:
   cmbSection.Text = Trim$(cmbSection.Text)
   If cmbSection.Text = "" Then
    MsgBox "Empty Section is not allowed!", vbExclamation
    cmbSection.SetFocus
    Exit Sub
   End If
   txt(0).Text = Trim$(txt(0).Text)
   If txt(0).Text = "" Then
    MsgBox "Empty key is not allowed!", vbExclamation
    txt(0).SetFocus
    Exit Sub
   End If
   txt(1).Text = Trim$(txt(1).Text)
   If txt(1).Text = "" Then
    aw = MsgBox("The english text is empty. Is this correct?", vbQuestion Or vbYesNo)
    If aw = vbNo Then
     txt(1).SetFocus
     Exit Sub
    End If
   End If
   txt(2).Text = Trim$(txt(2).Text)
   If txt(2).Text = "" Then
    aw = MsgBox("The german text is empty. Is this correct?", vbQuestion Or vbYesNo)
    If aw = vbNo Then
     txt(2).SetFocus
     Exit Sub
    End If
   End If

   Str1 = cmbSection.Text
   Str2 = txt(0).Text

   If Len(txt(1).Text) = 0 Then
     Str3 = " "
    Else
     Str3 = txt(1).Text
   End If

   txt(2).Text = Trim$(txt(2).Text)
   If Len(txt(2).Text) = 0 Then
     Str4 = " "
    Else
     Str4 = txt(2).Text
   End If

   If frmMain.AddLanguagesItem(Str1, Str2, Str3, Str4) = True Then
    Unload Me
   End If
  Case 1:
   Unload Me
 End Select
End Sub

Private Sub cmdEngClp_Click(Index As Integer)
 With txt(1)
  Select Case Index
   Case 0:
    If Len(.Text) > 0 Then
     Clipboard.Clear
     Clipboard.SetText .Text, vbCFText
    End If
   Case 1:
    If Len(Clipboard.GetText) > 0 Then
     .Text = Clipboard.GetText
    End If
  End Select
 End With
End Sub

Private Sub cmdGerClp_Click(Index As Integer)
 With txt(2)
  Select Case Index
   Case 0:
    If Len(.Text) > 0 Then
     Clipboard.Clear
     Clipboard.SetText .Text, vbCFText
    End If
   Case 1:
    If Len(Clipboard.GetText) > 0 Then
     .Text = Clipboard.GetText
    End If
  End Select
 End With
End Sub

Private Sub cmdKeyClp_Click(Index As Integer)
 With txt(0)
  Select Case Index
   Case 0:
    If Len(.Text) > 0 Then
     Clipboard.Clear
     Clipboard.SetText .Text, vbCFText
    End If
   Case 1:
    If Len(Clipboard.GetText) > 0 Then
     .Text = Clipboard.GetText
    End If
  End Select
 End With
End Sub

Private Sub txt_GotFocus(Index As Integer)
 With txt(Index)
  If LenB(.Text) > 0 Then
   If .Tag = "0" Then
    .SelStart = 0
    .SelLength = LenB(.Text)
    .Tag = "1"
    Else
    .Tag = "0"
   End If
  End If
 End With
End Sub
