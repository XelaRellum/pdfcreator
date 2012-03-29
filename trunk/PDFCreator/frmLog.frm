VERSION 5.00
Begin VB.Form frmLog 
   Caption         =   "Form1"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Visible         =   0   'False
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   495
      Left            =   105
      TabIndex        =   2
      Top             =   3045
      Width           =   1335
   End
   Begin VB.CheckBox chkLogging 
      Appearance      =   0  '2D
      Caption         =   "Logging"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   360
      Top             =   675
   End
   Begin VB.TextBox txtLog 
      Appearance      =   0  '2D
      BackColor       =   &H00C0FFFF&
      Height          =   2415
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Top             =   435
      Width           =   4455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   3045
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   3255
      TabIndex        =   4
      Top             =   3045
      Width           =   1335
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ClearLogfile
50020  With txtLog
50030   .Text = ReadLogfile
50040   .SelStart = 0
50050  End With
50060  cmdSave.Enabled = False
50070  cmdClear.Enabled = False
50080  cmdClose.SetFocus
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmLog", "cmdClear_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdClose_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If ShowOnlyLogfile Then
50020   If (chkLogging.value = 1 And Logging = False) Or (chkLogging.value = 0 And Logging = True) Then
50030    Options.Logging = chkLogging.value
50040    SaveOption Options, "Logging"
50050    If chkLogging.value = 1 Then
50060      Logging = True
50070     Else
50080      Logging = False
50090    End If
50100   End If
50110  End If
50120  Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmLog", "cmdClose_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdSave_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim cFiles As Collection, Cancel As Boolean
50020  Set cFiles = GetFilename("PDFCreator", GetMyFiles, 0, _
  LanguageStrings.DialogPrinterLogfiles & " (*.log)|*.log", SaveFile, Cancel, Me)
50040  If Cancel = True Then
50050   Exit Sub
50060  End If
50070  If cFiles.Count > 0 And FileExists(CompletePath(PDFCreatorLogfilePath) & PDFCreatorLogfile) = True Then
50080   FileCopy CompletePath(PDFCreatorLogfilePath) & PDFCreatorLogfile, cFiles.Item(1)
50090  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmLog", "cmdSave_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If KeyCode = vbKeyF1 Then
50020   KeyCode = 0
50030   Call HTMLHelp_ShowTopic("html\pdfcreator-user-manual.html")
50040  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmLog", "Form_KeyDown")
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
50010  KeyPreview = True
50020  With Options
50030   SetFontControls Me.Controls, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50040  End With
50050  ChangeLanguage
50060  If Logging Then
50070    chkLogging.value = 1
50080   Else
50090    chkLogging.value = 0
50100  End If
50110  chkLogging.Visible = ShowOnlyLogfile
50120  Call SendMessage(txtLog.hwnd, WM_SETTEXT, 0&, ByVal CStr(ReadLogfile))
50130  If Len(txtLog.Text) = 0 Then
50140    cmdSave.Enabled = False
50150    cmdClear.Enabled = False
50160   Else
50170    If InStr(txtLog.Text, vbCrLf) > 0 Then
50180     txtLog.SelStart = Len(txtLog.Text) - InStrRev(txtLog.Text, vbCrLf)
50190    End If
50200  End If
50210  If ShowOnlyLogfile = True Then
50220   FormInTaskbar Me, True, True
50230  End If
50240  With Screen
50250   Height = 0.75 * .Height
50260   Width = 0.75 * .Width
50270   Move (.Width - Width) / 2, (.Height - Height) / 2
50280  End With
50290
50300  With Options
50310   SetFontControls Me.Controls, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50320  End With
50330
50340  ShowAcceleratorsInForm Me, True
50350  Timer1.Enabled = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmLog", "Form_Load")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Resize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim minHeight As Long, minWidth As Long
50020  minWidth = 320 * Screen.TwipsPerPixelX
50030  minHeight = 200 * Screen.TwipsPerPixelY
50040  If Me.Height < minHeight Or Me.Width < minWidth Then
50050   Me.Height = minHeight: Me.Width = minWidth
50060   Exit Sub
50070  End If
50080  With chkLogging
50090   .Left = 0
50100   .Top = 100
50110   .Width = Me.Width - 100
50120  End With
50130  With txtLog
50140   .Top = Abs(CLng(ShowOnlyLogfile)) * (chkLogging.Top + chkLogging.Height + 100)
50150   .Left = 0
50160   .Width = ScaleWidth
50170   .Height = ScaleHeight - cmdClose.Height - .Top - 200
50180  End With
50190  cmdClear.Top = txtLog.Top + txtLog.Height + 100
50200  cmdClear.Left = txtLog.Left + 50
50210  cmdClose.Top = cmdClear.Top
50220  cmdClose.Left = txtLog.Left + txtLog.Width - cmdClose.Width - 50
50230  cmdSave.Top = cmdClear.Top
50240  cmdSave.Left = txtLog.Left + (txtLog.Width - cmdClose.Width) / 2 - 100
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmLog", "Form_Resize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetCursorOnTheBeginningOfTheLastLine(txtBox As Control)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStrf() As String
50020  If Len(txtBox.Text) > 0 Then
50030   If InStr(1, txtBox.Text, vbCrLf) > 0 Then
50040     txtBox.SelStart = InStrRev(txtBox.Text, vbCrLf) + Len(vbCrLf) - 1
50050    Else
50060     txtBox.SelStart = 0
50070   End If
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmLog", "SetCursorOnTheBeginningOfTheLastLine")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Timer1_Timer()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Timer1.Enabled = False
50020  SetCursorOnTheBeginningOfTheLastLine txtLog
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmLog", "Timer1_Timer")
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
50020   Caption = .LoggingLogfile
50030   cmdClose.Caption = .LoggingClose
50040   cmdClear.Caption = .LoggingClear
50050   cmdSave.Caption = .PrintingSave
50060  End With
50070  chkLogging.Caption = LanguageStrings.DialogPrinterLogging
50080  If ShowOnlyLogfile = True Then
50090   Caption = "PDFCreator - " & Caption
50100  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmLog", "ChangeLanguage")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
