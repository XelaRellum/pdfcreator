VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "PDFCreator"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4950
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   3240
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdEMail 
      Caption         =   "eMail"
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   8
      Top             =   2400
      Width           =   4695
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Optionen"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CheckBox chkDebug 
      Caption         =   "Debug Mode"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox txtInputFile 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   4215
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   1
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdStart 
      Appearance      =   0  '2D
      BackColor       =   &H00D9E5E9&
      Caption         =   "Speichern"
      Height          =   495
      Left            =   3480
      MaskColor       =   &H00D9E5E9&
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblTitle 
      Caption         =   "Titel des Dokuments:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label lblPSFile 
      Caption         =   "Zu konvertierende PS-Datei:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkDebug_Click()
If chkDebug.Value = 1 Then
frmMain.Height = 4770
Else
frmMain.Height = 2460
End If
End Sub

Private Sub cmdEMail_Click()
Dim R As Long
Dim lRet As Long
Dim outFile As String
Dim outFileName As String
Dim objOutlook As New Outlook.Application
Dim objMail As MailItem

Set objMail = objOutlook.CreateItem(olMailItem)

If Dir$(txtInputFile.Text) = "" Then MsgBox "Cannot find Source File", vbCritical, "Fehler": Exit Sub
frmProcess.Visible = True
frmProcess.anmProcess.Play
frmProcess.Refresh
frmMain.Visible = False
SaveTitle txtInputFile.Text, txtTitle.Text

outFileName = Space(260)
lRet = GetTempPath(260, outFileName)
outFileName = left(outFileName, lRet)
outFile = Space(260)
If Right(outFileName, 1) <> "\" Then outFileName = outFileName + "\"

outFile = outFileName + EMAIL_NAME & ".pdf"

CallGScript txtInputFile.Text, outFile
If frmMain.chkDebug.Value = 0 Then
  objMail.Attachments.Add outFile
  objMail.Display
  Kill (txtInputFile.Text)
  Kill (outFile)
  End
End If

frmProcess.Visible = False
frmMain.Visible = True
End Sub

Private Sub cmdOpen_Click()
Dim sOpen As SelectedFile
Dim Count As Integer
Dim FileList As String

    'On Error GoTo e_Trap
    
    FileDialog.sFilter = DOKUMENT_PS & Chr$(0) & "*.ps" & Chr$(0) & DOKUMENT_ALL & Chr$(0) & "*.*"
    
    FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
    FileDialog.sDlgTitle = SELECT_FILE
    FileDialog.sInitDir = GetSetting("PDFCreator", "Settings", "InitDir", App.Path & "\")
    sOpen = ShowOpen(Me.hWnd)
    If Err.number <> 32755 And sOpen.bCanceled = False Then
        txtInputFile = sOpen.sLastDirectory & sOpen.sFiles(1)
        txtTitle.Text = modGhostScript.GetTitle(txtInputFile)
    End If
    Exit Sub

e_Trap:
    Exit Sub
    Resume Next
End Sub

Private Sub cmdOptions_Click()
frmOptions.Show vbModal
End Sub

Private Sub cmdStart_Click()
Dim sSave As SelectedFile
Dim Count As Integer
Dim FileList As String
Dim R As Long
Dim outFile As String

If Dir$(txtInputFile.Text) = "" Then MsgBox "Cannot find Source File", vbCritical, "Fehler": GoTo e_Trap:

    'On Error GoTo e_Trap
    
    FileDialog.sFilter = DOKUMENT_PDF & Chr$(0) & "*.pdf"
    
    FileDialog.flags = OFN_HIDEREADONLY
    FileDialog.sDlgTitle = SAVE_FILE
    FileDialog.sInitDir = GetSetting("PDFCreator", "Settings", "InitDir", App.Path & "\")
    FileDialog.sFile = txtTitle.Text & ".pdf"
    sSave = ShowSave(Me.hWnd)
    If Err.number <> 32755 And sSave.bCanceled = False Then
        outFile = sSave.sLastDirectory & sSave.sFiles(1)
        SaveSetting "PDFCreator", "Settings", "InitDir", sSave.sLastDirectory
        If Right(LCase(outFile), 4) <> ".pdf" Then outFile = outFile & ".pdf"
        frmProcess.Visible = True
        frmProcess.anmProcess.Play
        frmProcess.Refresh
        frmMain.Visible = False
        SaveTitle txtInputFile.Text, txtTitle.Text
        CallGScript txtInputFile.Text, outFile
        If frmMain.chkDebug.Value = 0 Then
            Kill (txtInputFile.Text)
            R = ShellExecute(Me.hWnd, "Open", outFile, vbNullString, App.Path, vbMaximizedFocus)
        End
        End If
        frmProcess.Visible = False
        frmProcess.anmProcess.Play
        frmMain.Visible = True
    End If
    Exit Sub
    
e_Trap:
    Exit Sub
    Resume
End Sub

Private Sub Command1_Click()
OptimizePDF "D:\Eigene Dateien\Visual Basic\GhostScript\Readme.pdf", "D:\Eigene Dateien\Visual Basic\GhostScript\Readme2.pdf"
End Sub

Private Sub Form_Load()
frmOptions.Visible = False
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE

If Command$ <> vbNullString Then
    Select Case left$(Command$, 2)
      Case "-f"
        txtInputFile.Text = Right$(Command$, Len(Command$) - 2)
        txtTitle.Text = modGhostScript.GetTitle(txtInputFile)
      Case "%%"
      
      Case Else
        txtInputFile.Text = Command$
        txtTitle.Text = modGhostScript.GetTitle(txtInputFile)
    End Select
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
