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
      Width           =   6855
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   6855
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   6855
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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim item As ListItem, Str1 As String, Str2 As String, Str3 As String, _
  Str4 As String, aw As Long
50030  Select Case Index
  Case 0:
50050    cmbSection.Text = Trim$(cmbSection.Text)
50060    If cmbSection.Text = "" Then
50070     MsgBox "Empty Section is not allowed!", vbExclamation
50080     cmbSection.SetFocus
50090     Exit Sub
50100    End If
50110    txt(0).Text = Trim$(txt(0).Text)
50120    If txt(0).Text = "" Then
50130     MsgBox "Empty key is not allowed!", vbExclamation
50140     txt(0).SetFocus
50150     Exit Sub
50160    End If
50170    txt(1).Text = Trim$(txt(1).Text)
50180    If txt(1).Text = "" Then
50190     aw = MsgBox("The english text is empty. Is this correct?", vbQuestion Or vbYesNo)
50200     If aw = vbNo Then
50210      txt(1).SetFocus
50220      Exit Sub
50230     End If
50240    End If
50250    txt(2).Text = Trim$(txt(2).Text)
50260    If txt(2).Text = "" Then
50270     aw = MsgBox("The german text is empty. Is this correct?", vbQuestion Or vbYesNo)
50280     If aw = vbNo Then
50290      txt(2).SetFocus
50300      Exit Sub
50310     End If
50320    End If
50330
50340    Str1 = cmbSection.Text
50350    Str2 = txt(0).Text
50360
50370    If Len(txt(1).Text) = 0 Then
50380      Str3 = " "
50390     Else
50400      Str3 = txt(1).Text
50410    End If
50420
50430    txt(2).Text = Trim$(txt(2).Text)
50440    If Len(txt(2).Text) = 0 Then
50450      Str4 = " "
50460     Else
50470      Str4 = txt(2).Text
50480    End If
50490
50500    If frmMain.AddLanguagesItem(Str1, Str2, Str3, Str4) = True Then
50510     Unload Me
50520    End If
50530   Case 1:
50540    Unload Me
50550  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmLanguage", "cmd_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub


