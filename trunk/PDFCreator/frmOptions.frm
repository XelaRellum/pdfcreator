VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "PDF Optionen"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame fraCompress 
      Caption         =   "Komprimierung"
      Height          =   3855
      Left            =   360
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Frame fraGrey 
         Caption         =   "Graustufenbilder"
         Height          =   975
         Left            =   120
         TabIndex        =   26
         Top             =   1680
         Width           =   5535
         Begin VB.TextBox txtGreyRes 
            Height          =   285
            Left            =   4440
            TabIndex        =   31
            Top             =   540
            Width           =   735
         End
         Begin VB.ComboBox cmbGreyResample 
            Height          =   315
            ItemData        =   "frmOptions.frx":08CA
            Left            =   2280
            List            =   "frmOptions.frx":08D7
            Style           =   2  'Dropdown-Liste
            TabIndex        =   30
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   540
            Width           =   1935
         End
         Begin VB.CheckBox chkGreyResample 
            Caption         =   "Resample"
            Height          =   255
            Left            =   2280
            TabIndex        =   29
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox cmbGreyComp 
            Height          =   315
            ItemData        =   "frmOptions.frx":0916
            Left            =   120
            List            =   "frmOptions.frx":092F
            Style           =   2  'Dropdown-Liste
            TabIndex        =   28
            Top             =   540
            Width           =   1935
         End
         Begin VB.CheckBox chkGreyComp 
            Caption         =   "Komprimieren"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblGreyRes 
            Caption         =   "Auflösung"
            Height          =   255
            Left            =   4440
            TabIndex        =   32
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame fraMono 
         Caption         =   "Monochrombilder"
         Height          =   975
         Left            =   120
         TabIndex        =   19
         Top             =   2760
         Width           =   5535
         Begin VB.CheckBox chkMonoComp 
            Caption         =   "Komprimieren"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox cmbMonoComp 
            Height          =   315
            ItemData        =   "frmOptions.frx":0987
            Left            =   120
            List            =   "frmOptions.frx":0991
            Style           =   2  'Dropdown-Liste
            TabIndex        =   23
            Top             =   540
            Width           =   1935
         End
         Begin VB.CheckBox chkMonoResample 
            Caption         =   "Resample"
            Height          =   255
            Left            =   2280
            TabIndex        =   22
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox cmbMonoResample 
            Height          =   315
            ItemData        =   "frmOptions.frx":09A1
            Left            =   2280
            List            =   "frmOptions.frx":09AE
            Style           =   2  'Dropdown-Liste
            TabIndex        =   21
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   540
            Width           =   1935
         End
         Begin VB.TextBox txtMonoRes 
            Height          =   285
            Left            =   4440
            TabIndex        =   20
            Top             =   540
            Width           =   735
         End
         Begin VB.Label lblMonoRes 
            Caption         =   "Auflösung"
            Height          =   255
            Left            =   4440
            TabIndex        =   25
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.CheckBox chkTextComp 
         Caption         =   "Texte komprimieren"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   4335
      End
      Begin VB.Frame fraColor 
         Caption         =   "Farbbilder"
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   5535
         Begin VB.TextBox txtColorRes 
            Height          =   285
            Left            =   4440
            TabIndex        =   18
            Top             =   540
            Width           =   735
         End
         Begin VB.ComboBox cmbColorResample 
            Height          =   315
            ItemData        =   "frmOptions.frx":09ED
            Left            =   2280
            List            =   "frmOptions.frx":09FA
            Style           =   2  'Dropdown-Liste
            TabIndex        =   16
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   540
            Width           =   1935
         End
         Begin VB.CheckBox chkColorResample 
            Caption         =   "Resample"
            Height          =   255
            Left            =   2280
            TabIndex        =   15
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox cmbColorComp 
            Height          =   315
            ItemData        =   "frmOptions.frx":0A39
            Left            =   120
            List            =   "frmOptions.frx":0A52
            Style           =   2  'Dropdown-Liste
            TabIndex        =   14
            Top             =   540
            Width           =   1935
         End
         Begin VB.CheckBox chkColorComp 
            Caption         =   "Komprimieren"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblColorRes 
            Caption         =   "Auflösung"
            Height          =   255
            Left            =   4440
            TabIndex        =   17
            Top             =   240
            Width           =   975
         End
      End
   End
   Begin VB.Frame fraColors 
      Caption         =   "Farbverwaltung"
      Height          =   3495
      Left            =   360
      TabIndex        =   41
      Top             =   600
      Visible         =   0   'False
      Width           =   5775
      Begin VB.ComboBox cmbColorModel 
         Height          =   315
         ItemData        =   "frmOptions.frx":0AAA
         Left            =   120
         List            =   "frmOptions.frx":0AB7
         Style           =   2  'Dropdown-Liste
         TabIndex        =   48
         Tag             =   "RGB|CMYK|GRAY"
         Top             =   360
         Width           =   2775
      End
      Begin VB.Frame fraColorOptions 
         Caption         =   "Optionen"
         Height          =   1455
         Left            =   120
         TabIndex        =   44
         Top             =   1920
         Width           =   5535
         Begin VB.CheckBox chkPreserverHalftone 
            Caption         =   "Halbton Informationen beibehalten"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   1080
            Width           =   3375
         End
         Begin VB.CheckBox chkPreserveTransfer 
            Caption         =   "Transferfunktionen beibehalten"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Tag             =   "Remove|Preserve"
            Top             =   720
            Width           =   3615
         End
         Begin VB.CheckBox chkPreserveOverprint 
            Caption         =   "Überdruckeinstellungen beibehalten"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   4335
         End
      End
      Begin VB.CheckBox chkCMYKtoRGB 
         Caption         =   "CMYK Bilder in RGB konvertieren"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.Frame fraFonts 
      Caption         =   "Schriftoptionen"
      Height          =   2895
      Left            =   360
      TabIndex        =   36
      Top             =   600
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox txtSubSetPerc 
         Height          =   285
         Left            =   120
         TabIndex        =   39
         Top             =   1140
         Width           =   495
      End
      Begin VB.CheckBox chkSubSetFonts 
         Caption         =   "Teilschriften einbetten, wenn Zahl der benutzten Zeichen geringer als:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   780
         Width           =   5295
      End
      Begin VB.CheckBox chkEmbedAll 
         Caption         =   "Alle Schriften einbetten"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label lblPerc 
         Caption         =   "%"
         Height          =   255
         Left            =   720
         TabIndex        =   40
         Top             =   1200
         Width           =   615
      End
   End
   Begin VB.Frame fraGeneral 
      Caption         =   "Allgemeine Optionen"
      Height          =   2655
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   5775
      Begin VB.CheckBox chkASCII85 
         Caption         =   "Binärdaten in ASCII85 umwandeln"
         Height          =   255
         Left            =   2280
         TabIndex        =   35
         Top             =   2280
         Width           =   2895
      End
      Begin VB.ComboBox cmbOverprint 
         Height          =   315
         ItemData        =   "frmOptions.frx":0B14
         Left            =   2280
         List            =   "frmOptions.frx":0B1E
         Style           =   2  'Dropdown-Liste
         TabIndex        =   34
         Top             =   1740
         Width           =   2655
      End
      Begin VB.TextBox txtRes 
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox cmbCompat 
         Height          =   315
         ItemData        =   "frmOptions.frx":0B50
         Left            =   2280
         List            =   "frmOptions.frx":0B5D
         Style           =   2  'Dropdown-Liste
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox cmbRotate 
         Height          =   315
         ItemData        =   "frmOptions.frx":0BB8
         Left            =   2280
         List            =   "frmOptions.frx":0BC5
         Style           =   2  'Dropdown-Liste
         TabIndex        =   3
         Tag             =   "None|All|PageByPage"
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label lblOverprint 
         Alignment       =   1  'Rechts
         Caption         =   "Überdrucken:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label lblResolution 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Auflösung:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1260
         Width           =   1935
      End
      Begin VB.Label lblCompat 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Kompatibilität:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   1935
      End
      Begin VB.Label lblAutoRotate 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Seiten automatisch drehen:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   780
         Width           =   1935
      End
      Begin VB.Label lblDPI 
         BackStyle       =   0  'Transparent
         Caption         =   "dpi"
         Height          =   255
         Left            =   3000
         TabIndex        =   6
         Top             =   1260
         Width           =   735
      End
   End
   Begin VB.Frame fraLanguage 
      Caption         =   "Sprache"
      Height          =   3255
      Left            =   360
      TabIndex        =   49
      Top             =   600
      Visible         =   0   'False
      Width           =   5775
      Begin VB.ListBox lstLanguage 
         Height          =   2595
         ItemData        =   "frmOptions.frx":0BE7
         Left            =   240
         List            =   "frmOptions.frx":0BE9
         TabIndex        =   51
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton cmdApplyLang 
         Caption         =   "Sprache Einstellen"
         Height          =   375
         Left            =   2520
         TabIndex        =   50
         Top             =   2700
         Width           =   1695
      End
      Begin VB.Label lblDynLanguage 
         Height          =   255
         Left            =   2520
         TabIndex        =   55
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label lblDynAuthor 
         Height          =   255
         Left            =   2520
         TabIndex        =   54
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label lblLanguage 
         Caption         =   "Sprache:"
         Height          =   255
         Left            =   2520
         TabIndex        =   53
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lblAutor 
         Caption         =   "Autor:"
         Height          =   255
         Left            =   2520
         TabIndex        =   52
         Top             =   1080
         Width           =   2415
      End
   End
   Begin VB.Frame fraSecurity 
      Caption         =   "Not implented yet"
      Height          =   3255
      Left            =   360
      TabIndex        =   42
      Top             =   600
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Übernehmen"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   4680
      Width           =   1575
   End
   Begin MSComctlLib.TabStrip tabOptions 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7858
      TabMinWidth     =   884
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Allgemein"
            Key             =   "keyGeneral"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Komprimierung"
            Key             =   "keyCompress"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Schriften"
            Key             =   "keyFonts"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Farben"
            Key             =   "keyColors"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sicherheit"
            Key             =   "keySecurity"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sprache"
            Key             =   "keyLanguage"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public colLangFiles As New Collection

Private Sub cmdApply_Click()
SaveAllSettings
Me.Hide
End Sub

Private Sub cmdApplyLang_Click()
UpdateLang
SaveSetting "PDFCreator", "Settings", "Language", lstLanguage.Text
GetAllSettings
End Sub

Public Sub UpdateLang()
Dim colValues As Collection
Dim x As Long
Dim strLang As String
Dim RetSplit As Variant
Dim myStr As String
Dim myStr2 As String
Dim lString As String

For i = 1 To colLangFiles.Count
If left$(colLangFiles(i), Len(lstLanguage.Text)) = lstLanguage.Text Then strLang = Right$(colLangFiles(i), Len(colLangFiles(i)) - (Len(lstLanguage.Text) + 1))
Next i

'Doing frmMain
GetIniAppAllKeys App.Path & "\languages\" & strLang, "MAIN", colValues
For i = 0 To frmMain.Controls.Count - 1
  x = FindKey(frmMain.Controls(i).Name, colValues)
  If x <> 0 Then frmMain.Controls(i).Caption = GetIni(App.Path & "\languages\" & strLang, "MAIN", colValues(x), frmMain.Controls(i).Caption)
Next i

'Doing frmOptions
GetIniAppAllKeys App.Path & "\languages\" & strLang, "OPTIONS", colValues
For i = 0 To frmOptions.Controls.Count - 1
  x = FindKey(frmOptions.Controls(i).Name, colValues)
  If x <> 0 Then 'And GetIni(App.Path & "\languages\" & strLang, "OPTIONS", colValues(x), "") <> "" Then
  lString = GetIni(App.Path & "\languages\" & strLang, "OPTIONS", colValues(x), "")
    If left$(frmOptions.Controls(i).Name, 3) = "cmb" Then
      myStr = lString
      myStr2 = myStr
      sCount = InCount(myStr2)
      RetSplit = Split(myStr, "|")
      frmOptions.Controls(i).Clear
      For m = 0 To sCount
        frmOptions.Controls(i).AddItem RetSplit(m)
      Next m
    Else
      frmOptions.Controls(i).Caption = lString
    End If
  End If
Next i

'Doing Tabs
lString = GetIni(App.Path & "\languages\" & strLang, "OPTIONS", "OptionTab", "")
myStr = lString
myStr2 = myStr
sCount = InCount(myStr2)
RetSplit = Split(myStr, "|")
For i = 1 To tabOptions.Tabs.Count
tabOptions.Tabs(i) = RetSplit(i - 1)
Next i

frmOptions.Caption = GetIni(App.Path & "\languages\" & strLang, "OPTIONS", "frmOptions", frmOptions.Caption)
frmProcess.Caption = GetIni(App.Path & "\languages\" & strLang, "PROCESSING", "frmProcess", frmProcess.Caption)
frmProcess.lblStatus.Caption = GetIni(App.Path & "\languages\" & strLang, "PROCESSING", "lblStatus", frmProcess.lblStatus.Caption)

DOKUMENT_PS = GetIni(App.Path & "\languages\" & strLang, "FILES", "DocumentPS", "PostScript-File (*.ps)")
DOKUMENT_PDF = GetIni(App.Path & "\languages\" & strLang, "FILES", "DocumentPDF", "PDF-Document (*.pdf)")
DOKUMENT_ALL = GetIni(App.Path & "\languages\" & strLang, "FILES", "DocumentAll", "All Files (*.*)")
SELECT_FILE = GetIni(App.Path & "\languages\" & strLang, "FILES", "Open", "Select File")
SAVE_FILE = GetIni(App.Path & "\languages\" & strLang, "FILES", "SaveAs", "Save as")

EMAIL_NAME = GetIni(App.Path & "\languages\" & strLang, "EMAIL", "DocumentName", "Document")
End Sub

Private Function InCount(str As String)
Dim aCount As Integer
Dim Pos As Integer
Pos = 1
Pos = InStr(Pos, str, "|")

Do While Pos <> 0
aCount = aCount + 1
str = Right(str, Len(str) - Pos)
Pos = InStr(1, str, "|")
Loop
InCount = aCount
End Function

Private Sub Command2_Click()
Dim colValues As Collection
Dim x As Long

GetIniAppAllKeys App.Path & "\languages\deutsch.ini", "GENERAL", colValues
x = FindKey("DisplayName", colValues)
MsgBox colValues(x) & " = " & GetIni(App.Path & "\languages\deutsch.ini", "GENERAL", colValues(x), "")
End Sub

Private Sub Form_Load()
GetLangFiles
lstLanguage.Text = GetSetting("PDFCreator", "Settings", "Language", "English")
UpdateLang
GetAllSettings
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
    Cancel = True
    SaveAllSettings
    Me.Hide
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveAllSettings
End Sub

Private Sub lstLanguage_Click()
UpdateInfos
End Sub

Private Sub UpdateInfos()
Dim colValues As Collection
Dim x As Long
Dim strLang As String

For i = 1 To colLangFiles.Count
If left$(colLangFiles(i), Len(lstLanguage.Text)) = lstLanguage.Text Then strLang = Right$(colLangFiles(i), Len(colLangFiles(i)) - (Len(lstLanguage.Text) + 1))
Next i

GetIniAppAllKeys App.Path & "\languages\" & strLang, "GENERAL", colValues
x = FindKey("DisplayName", colValues)
lblDynLanguage = GetIni(App.Path & "\languages\" & strLang, "GENERAL", colValues(x), "")
x = FindKey("Author", colValues)
lblDynAuthor = GetIni(App.Path & "\languages\" & strLang, "GENERAL", colValues(x), "")

End Sub

Private Sub tabOptions_Click()
fraGeneral.Visible = False
fraCompress.Visible = False
fraFonts.Visible = False
fraColors.Visible = False
fraSecurity.Visible = False
fraLanguage.Visible = False

Select Case tabOptions.SelectedItem
  
  Case tabOptions.Tabs.Item(1).Caption
    fraGeneral.Visible = True
  
  Case tabOptions.Tabs.Item(2).Caption
    fraCompress.Visible = True
    
  Case tabOptions.Tabs.Item(3).Caption
    fraFonts.Visible = True

  Case tabOptions.Tabs.Item(4).Caption
    fraColors.Visible = True
    
  Case tabOptions.Tabs.Item(5).Caption
    fraSecurity.Visible = True

  Case tabOptions.Tabs.Item(6).Caption
    fraLanguage.Visible = True
    
End Select
End Sub

Public Sub GetLangFiles()
Dim lFiles As New Collection
Dim colValues As New Collection
Dim x As Long

Name1 = Dir(App.Path & "\languages\", vbDirectory)   ' Ersten Eintrag abrufen.
Do While Name1 <> ""   ' Schleife beginnen.
   ' Aktuelles und übergeordnetes Verzeichnis ignorieren.
   If Name1 <> "." And Name1 <> ".." Then
         lFiles.Add Name1
   End If
   Name1 = Dir   ' Nächsten Eintrag abrufen.
Loop

For i = 1 To lFiles.Count
GetIniAppAllKeys App.Path & "\languages\" & lFiles(i), "GENERAL", colValues
x = FindKey("DisplayName", colValues)
lstLanguage.AddItem GetIni(App.Path & "\languages\" & lFiles(i), "GENERAL", colValues(x), "")
colLangFiles.Add GetIni(App.Path & "\languages\" & lFiles(i), "GENERAL", colValues(x), "") & "=" & lFiles(i)
Next i
End Sub

Public Sub GetAllSettings()
lstLanguage.Text = GetSetting("PDFCreator", "Settings", "Language", "English")

'General
cmbCompat.ListIndex = GetSetting("PDFCreator", "Settings", "Compatibility", 1)
cmbRotate.ListIndex = GetSetting("PDFCreator", "Settings", "AutoRotatePage", 0)
txtRes.Text = GetSetting("PDFCreator", "Settings", "Resolution", 600)
cmbOverprint.ListIndex = GetSetting("PDFCreator", "Settings", "Overprint", 0)
chkASCII85.Value = GetSetting("PDFCreator", "Settings", "ASCII85", 0)

'Compression
chkTextComp.Value = GetSetting("PDFCreator", "Settings", "TextCompression", 1)

chkColorComp.Value = GetSetting("PDFCreator", "Settings", "chkColorComp", 0)
chkColorResample.Value = GetSetting("PDFCreator", "Settings", "chkColorResample", 0)
cmbColorComp.ListIndex = GetSetting("PDFCreator", "Settings", "cmbColorComp", 0)
cmbColorResample.ListIndex = GetSetting("PDFCreator", "Settings", "cmbColorResample", 0)
txtColorRes.Text = GetSetting("PDFCreator", "Settings", "ColorRes", 300)

chkGreyComp.Value = GetSetting("PDFCreator", "Settings", "chkGreyComp", 1)
chkGreyResample.Value = GetSetting("PDFCreator", "Settings", "chkGreyResample", 0)
cmbGreyComp.ListIndex = GetSetting("PDFCreator", "Settings", "cmbGreyComp", 0)
cmbGreyResample.ListIndex = GetSetting("PDFCreator", "Settings", "cmbGreyResample", 0)
txtGreyRes.Text = GetSetting("PDFCreator", "Settings", "GreyRes", 300)

chkMonoComp.Value = GetSetting("PDFCreator", "Settings", "chkMonoComp", 1)
chkMonoResample.Value = GetSetting("PDFCreator", "Settings", "chkMonoResample", 0)
cmbMonoComp.ListIndex = GetSetting("PDFCreator", "Settings", "cmbMonoComp", 0)
cmbMonoResample.ListIndex = GetSetting("PDFCreator", "Settings", "cmbMonoResample", 0)
txtMonoRes.Text = GetSetting("PDFCreator", "Settings", "MonoRes", 1200)

'Fonts
chkEmbedAll.Value = GetSetting("PDFCreator", "Settings", "EmbedAll", 1)
chkSubSetFonts.Value = GetSetting("PDFCreator", "Settings", "SubSetFonts", 1)
txtSubSetPerc.Text = GetSetting("PDFCreator", "Settings", "SubSetPerc", 100)

'Colors
cmbColorModel.ListIndex = GetSetting("PDFCreator", "Settings", "ColorModel", 1)
chkCMYKtoRGB.Value = GetSetting("PDFCreator", "Settings", "CMYKtoRGB", 1)
chkPreserveOverprint.Value = GetSetting("PDFCreator", "Settings", "PreserveOverprint", 1)
chkPreserveTransfer.Value = GetSetting("PDFCreator", "Settings", "PreserveTransfer", 1)
chkPreserverHalftone.Value = GetSetting("PDFCreator", "Settings", "PreserveHalftone", 0)
End Sub

Public Sub SaveAllSettings()
SaveSetting "PDFCreator", "Settings", "Language", lstLanguage.Text

'General
SaveSetting "PDFCreator", "Settings", "Compatibility", cmbCompat.ListIndex
SaveSetting "PDFCreator", "Settings", "AutoRotatePage", cmbRotate.ListIndex
SaveSetting "PDFCreator", "Settings", "Resolution", txtRes.Text
SaveSetting "PDFCreator", "Settings", "Overprint", cmbOverprint.ListIndex
SaveSetting "PDFCreator", "Settings", "ASCII85", chkASCII85.Value

'Compression
SaveSetting "PDFCreator", "Settings", "TextCompression", chkTextComp.Value

SaveSetting "PDFCreator", "Settings", "chkColorComp", chkColorComp.Value
SaveSetting "PDFCreator", "Settings", "chkColorResample", chkColorResample.Value
SaveSetting "PDFCreator", "Settings", "cmbColorComp", cmbColorComp.ListIndex
SaveSetting "PDFCreator", "Settings", "cmbColorResample", cmbColorResample.ListIndex
SaveSetting "PDFCreator", "Settings", "ColorRes", txtColorRes.Text

SaveSetting "PDFCreator", "Settings", "chkGreyComp", chkGreyComp.Value
SaveSetting "PDFCreator", "Settings", "chkGreyResample", chkGreyResample.Value
SaveSetting "PDFCreator", "Settings", "cmbGreyComp", cmbGreyComp.ListIndex
SaveSetting "PDFCreator", "Settings", "cmbGreyResample", cmbGreyResample.ListIndex
SaveSetting "PDFCreator", "Settings", "GreyRes", txtGreyRes.Text

SaveSetting "PDFCreator", "Settings", "chkMonoComp", chkMonoComp.Value
SaveSetting "PDFCreator", "Settings", "chkMonoResample", chkMonoResample.Value
SaveSetting "PDFCreator", "Settings", "cmbMonoComp", cmbMonoComp.ListIndex
SaveSetting "PDFCreator", "Settings", "cmbMonoResample", cmbMonoResample.ListIndex
SaveSetting "PDFCreator", "Settings", "MonoRes", txtMonoRes.Text

'Fonts
SaveSetting "PDFCreator", "Settings", "EmbedAll", chkEmbedAll.Value
SaveSetting "PDFCreator", "Settings", "SubSetFonts", chkSubSetFonts.Value
SaveSetting "PDFCreator", "Settings", "SubSetPerc", txtSubSetPerc.Text

'Colors
SaveSetting "PDFCreator", "Settings", "ColorModel", cmbColorModel.ListIndex
SaveSetting "PDFCreator", "Settings", "CMYKtoRGB", chkCMYKtoRGB.Value
SaveSetting "PDFCreator", "Settings", "PreserveOverprint", chkPreserveOverprint.Value
SaveSetting "PDFCreator", "Settings", "PreserveTransfer", chkPreserveTransfer.Value
SaveSetting "PDFCreator", "Settings", "PreserveHalftone", chkPreserverHalftone.Value
End Sub


