VERSION 5.00
Begin VB.Form frmProfile 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Profile"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6510
   Icon            =   "frmProfile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows-Standard
   Begin PDFCreator.dmFrame dmFrProfile 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2143
      Caption         =   "Profile"
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
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4920
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtProfile 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   6015
      End
   End
End
Attribute VB_Name = "frmProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Profiles As Collection, ProfileAction As eProfileAction, CurrentProfile As String

Private Sub AddRenameProfile()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case ProfileAction
        Case eProfileAction.AddProfileAction
50030    frmOptions.AddProfile Trim$(txtProfile.Text)
50040   Case eProfileAction.RenameProfileAction
50050    frmOptions.RenameProfile Trim$(txtProfile.Text)
50060  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmProfile", "AddRenameProfile")
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
Select Case ErrPtnr.OnError("frmProfile", "cmdCancel_Click")
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
50010  AddRenameProfile
50020  Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmProfile", "cmdOK_Click")
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
50020   cmdOK.Caption = .OptionsProfileOk
50030   cmdCancel.Caption = .OptionsProfileCancel
50040  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmProfile", "ChangeLanguage")
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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmProfile", "Form_Load")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtProfile_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case ProfileAction
        Case eProfileAction.AddProfileAction
50030    If ProfileExists(Trim$(txtProfile.Text)) = True Then
50040      cmdOK.Enabled = False
50050     Else
50060      cmdOK.Enabled = True
50070    End If
50080   Case eProfileAction.RenameProfileAction
50090    If CurrentProfile = Trim$(txtProfile.Text) Then
50100      cmdOK.Enabled = False
50110     ElseIf LCase$(CurrentProfile) = LCase(Trim$(txtProfile.Text)) Then
50120      cmdOK.Enabled = True
50130     Else
50140      If ProfileExists(Trim$(txtProfile.Text)) = True Then
50150        cmdOK.Enabled = False
50160       Else
50170        cmdOK.Enabled = True
50180      End If
50190    End If
50200  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmProfile", "txtProfile_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtProfile_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case KeyAscii
        Case Asc("0") To Asc("9"), Asc("a") To Asc("z"), Asc("A") To Asc("Z"), Asc(" "), 8, 13, 27
50030    If KeyAscii = 27 Then
50040     Unload Me
50050    End If
50060    If KeyAscii = 13 Then
50070     AddRenameProfile
50080     Unload Me
50090    End If
50100   Case Else
50110    KeyAscii = 0
50120  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmProfile", "txtProfile_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function ProfileExists(Profilename) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  For i = 1 To Profiles.Count
50030   If StrComp(Profiles(i), Profilename, vbTextCompare) = 0 Then
50040    ProfileExists = True
50050    Exit Function
50060   End If
50070  Next i
50080  ProfileExists = False
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmProfile", "ProfileExists")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
