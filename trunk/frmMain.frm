VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "PDFCreator"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4950
   StartUpPosition =   2  'Bildschirmmitte
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
      Width           =   2175
   End
   Begin VB.CheckBox chkDebug 
      Caption         =   "Debug Mode"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2895
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
      Caption         =   "Starten"
      Height          =   495
      Left            =   2640
      MaskColor       =   &H00D9E5E9&
      TabIndex        =   4
      Top             =   1440
      Width           =   2175
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
frmMain.Height = 2700
End If
End Sub

Private Sub cmdOpen_Click()
Dim sOpen As SelectedFile
Dim Count As Integer
Dim FileList As String

    'On Error GoTo e_Trap
    
    FileDialog.sFilter = "PostScript-Dateien (*.ps)" & Chr$(0) & "*.ps" & Chr$(0) & "Alle Dateien (*.*)" & Chr$(0) & "*.*"
    
    FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
    FileDialog.sDlgTitle = "Datei wählen"
    FileDialog.sInitDir = GetSetting("PDFCreator", "Settings", "InitDir", App.Path & "\")
    sOpen = ShowOpen(Me.hWnd)
    If Err.number <> 32755 And sOpen.bCanceled = False Then
        txtInputFile = sOpen.sLastDirectory & sOpen.sFiles(1)
        txtTitle.Text = modGhostScript.GetTitle(txtInputFile)
    End If
    Exit Sub

e_Trap:
    Exit Sub
    Resume
End Sub

Private Sub cmdOptions_Click()
frmOptions.Show vbModal
End Sub

Private Sub cmdStart_Click()
Dim sSave As SelectedFile
Dim Count As Integer
Dim FileList As String
Dim R As Long
Dim OutFile As String

If Dir$(txtInputFile.Text) = "" Then MsgBox "Quelldatei nicht gefunden", vbCritical, "Fehler": GoTo e_Trap:

    'On Error GoTo e_Trap
    
    FileDialog.sFilter = "PDF Dokumente (*.pdf)" & Chr$(0) & "*.pdf"
    
    ' See Standard CommonDialog Flags for all options
    FileDialog.flags = OFN_HIDEREADONLY
    FileDialog.sDlgTitle = "Speichern unter"
    FileDialog.sInitDir = GetSetting("PDFCreator", "Settings", "InitDir", App.Path & "\")
    sSave = ShowSave(Me.hWnd)
    If Err.number <> 32755 And sSave.bCanceled = False Then
        OutFile = sSave.sLastDirectory & sSave.sFiles(1)
        SaveSetting "PDFCreator", "Settings", "InitDir", sSave.sLastDirectory
        If Right(LCase(OutFile), 4) <> ".pdf" Then OutFile = OutFile & ".pdf"
        frmProcess.Visible = True
        frmProcess.anmProcess.Play
        frmProcess.Refresh
        frmMain.Visible = False
        SaveTitle txtInputFile.Text, txtTitle.Text
        CallGScript txtInputFile.Text, OutFile
        If frmMain.chkDebug.Value = 0 Then
            Kill (txtInputFile.Text)
            R = ShellExecute(Me.hWnd, "Open", OutFile, vbNullString, App.Path, vbMaximizedFocus)
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
