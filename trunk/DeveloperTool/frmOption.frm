VERSION 5.00
Begin VB.Form frmOption 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Option"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   Icon            =   "frmOption.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox cmbComment 
      Height          =   315
      Left            =   120
      TabIndex        =   16
      Top             =   360
      Width           =   4095
   End
   Begin VB.Frame fra 
      Caption         =   "Value"
      Height          =   2415
      Left            =   4440
      TabIndex        =   8
      Top             =   840
      Width           =   2775
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lbl 
         Caption         =   "Right Limit (use < for no positiv limit)"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label lbl 
         Caption         =   "Left Limit (use > for no negativ limit)"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label lbl 
         Caption         =   "Standard"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown-Liste
      TabIndex        =   6
      Top             =   2760
      Width           =   4095
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   4095
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4095
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Cancel"
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Save"
      Height          =   495
      Index           =   0
      Left            =   6000
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Comment"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "Type"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "Objectname"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click(Index As Integer)
 Dim item As ListItem, aw As Long, _
  Str1 As String, Str2 As String, Str3 As String, Str4 As String, Str5 As String, Str6 As String
 Select Case Index
  Case 0:
   txt(0).Text = Trim$(txt(0).Text)
   If txt(0).Text = "" Then
    MsgBox "Empty name is not allowed!", vbExclamation
    Exit Sub
   End If
   txt(1).Text = Trim$(txt(1).Text)
   If txt(1).Text = "" Then
    aw = MsgBox("If you use an empty objectname, you must change the source code in 'CreateModOptions (ShowOptions, GetOptions)' in 'frmMain'!", vbExclamation Or vbOKCancel)
    If aw = vbCancel Then
     Exit Sub
    End If
   End If
   Str1 = txt(0).Text
   Str2 = txt(1).Text
   Str3 = cmbType.List(cmbType.ListIndex)
   txt(2).Text = Trim$(txt(2).Text)
   If Len(txt(2).Text) = 0 Then
     Str4 = " "
    Else
     Str4 = txt(2).Text
   End If
   txt(3).Text = Trim$(txt(3).Text)
   If Len(txt(3).Text) = 0 Then
     Str5 = " "
    Else
     Str5 = txt(3).Text
   End If
   txt(4).Text = Trim$(txt(4).Text)
   If Len(txt(4).Text) = 0 Then
     Str6 = " "
    Else
     Str6 = txt(4).Text
   End If
   If frmMain.AddOptionsItem(Str1, Str2, Str3, Str4, Str5, Str6, cmbComment.Text) = True Then
    Unload Me
   End If
  Case 1:
   Unload Me
 End Select
End Sub

Private Sub Form_Load()
 With cmbType
  .AddItem "Boolean"
  .AddItem "Byte"
  .AddItem "Long"
  .AddItem "String"
  .AddItem "Double"
  .ListIndex = 3
 End With
End Sub
