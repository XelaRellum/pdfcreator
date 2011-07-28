VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlOptDocument 
   ClientHeight    =   5895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12840
   ScaleHeight     =   5895
   ScaleWidth      =   12840
   ToolboxBitmap   =   "ctlOptDocument.ctx":0000
   Begin PDFCreator.dmFrame dmFraProgDocument2 
      Height          =   2610
      Left            =   6480
      TabIndex        =   8
      Top             =   480
      Width           =   6195
      _extentx        =   10927
      _extenty        =   4604
      caption         =   "Document 2"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "ctlOptDocument.ctx":0312
      Begin VB.CheckBox chkUseFixPaperSize 
         Appearance      =   0  '2D
         Caption         =   "Use a fix papersize"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   6000
      End
      Begin VB.ComboBox cmbDocumentPapersizes 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown-Liste
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox chkUseCustomPapersize 
         Appearance      =   0  '2D
         Caption         =   "Use a custom papersize"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   1200
         Width           =   5760
      End
      Begin VB.TextBox txtCustomPapersizeWidth 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   640
         TabIndex        =   14
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtCustomPapersizeHeight 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   1920
         TabIndex        =   15
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblCustomPapersizeWidth 
         AutoSize        =   -1  'True
         Caption         =   "Width"
         Height          =   195
         Left            =   640
         TabIndex        =   12
         Top             =   1560
         Width           =   420
      End
      Begin VB.Label lblCustomPapersizeHeight 
         AutoSize        =   -1  'True
         Caption         =   "Height"
         Height          =   195
         Left            =   1920
         TabIndex        =   13
         Top             =   1560
         Width           =   465
      End
      Begin VB.Label lblCustomPapersizeInfo 
         AutoSize        =   -1  'True
         Caption         =   "Units of 1/72 of an inch."
         Height          =   195
         Left            =   640
         TabIndex        =   16
         Top             =   2280
         Width           =   1725
      End
   End
   Begin PDFCreator.dmFrame dmFraProgStamp 
      Height          =   2610
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   6195
      _extentx        =   10927
      _extenty        =   4604
      caption         =   "Stamp"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "ctlOptDocument.ctx":033E
      Begin VB.TextBox txtStampString 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   3495
      End
      Begin VB.PictureBox picStampFontColor 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4800
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   22
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox chkStampUseOutlineFont 
         Appearance      =   0  '2D
         Caption         =   "Use outline font"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1560
         Width           =   5895
      End
      Begin VB.CommandButton cmdStampFont 
         Caption         =   "..."
         Height          =   315
         Left            =   3720
         TabIndex        =   21
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtOutlineFontThickness 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         Height          =   255
         Left            =   2040
         TabIndex        =   26
         Text            =   "0"
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label lblStampString 
         AutoSize        =   -1  'True
         Caption         =   "Stampstring"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   825
      End
      Begin VB.Label lblStampFontcolor 
         AutoSize        =   -1  'True
         Caption         =   "Font-color"
         Height          =   195
         Left            =   4800
         TabIndex        =   19
         Top             =   480
         Width           =   705
      End
      Begin VB.Label lblOutlineFontThickness 
         AutoSize        =   -1  'True
         Caption         =   "Outline font thickness"
         Height          =   195
         Left            =   390
         TabIndex        =   25
         Top             =   2040
         Width           =   1530
      End
      Begin VB.Label lblFontNameSize 
         AutoSize        =   -1  'True
         Caption         =   "Arial, 12"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   570
      End
   End
   Begin PDFCreator.dmFrame dmFraProgDocument1 
      Height          =   2250
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6195
      _extentx        =   10927
      _extenty        =   3969
      caption         =   "Document 1"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "ctlOptDocument.ctx":036A
      Begin VB.CheckBox chkUseCreationDateNow 
         Appearance      =   0  '2D
         Caption         =   "Use the current Date/Time for 'Creation Date'"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   5985
      End
      Begin VB.TextBox txtStandardAuthor 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   3495
      End
      Begin VB.CheckBox chkUseStandardAuthor 
         Appearance      =   0  '2D
         Caption         =   "Use Standardauthor"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   5985
      End
      Begin VB.ComboBox cmbAuthorTokens 
         Appearance      =   0  '2D
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "ctlOptDocument.ctx":0396
         Left            =   3720
         List            =   "ctlOptDocument.ctx":0398
         Style           =   2  'Dropdown-Liste
         TabIndex        =   5
         Top             =   840
         Width           =   2415
      End
      Begin VB.CheckBox chkOnePagePerFile 
         Appearance      =   0  '2D
         Caption         =   "One page per file"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1890
         Width           =   5985
      End
      Begin VB.Label lblAuthorTokens 
         AutoSize        =   -1  'True
         Caption         =   "Add a Author-Token"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3720
         TabIndex        =   3
         Top             =   600
         Width           =   1440
      End
   End
   Begin MSComctlLib.TabStrip tbstrProgDocument 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   10186
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ctlOptDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private StampFont As tFont
Private ControlsEnabled As Boolean

Public Sub SetControlsEnabled(value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ControlsEnabled = value
50020
50030  chkUseStandardAuthor.Enabled = value
50040  If chkUseStandardAuthor.Enabled = True And chkUseStandardAuthor.value = 1 Then
50050    txtStandardAuthor.Enabled = True
50060    txtStandardAuthor.BackColor = &H80000005
50070    lblAuthorTokens.Enabled = True
50080    cmbAuthorTokens.Enabled = True
50090    cmbAuthorTokens.BackColor = &H80000005
50100   Else
50110    txtStandardAuthor.Enabled = False
50120    txtStandardAuthor.BackColor = &H8000000F
50130    lblAuthorTokens.Enabled = False
50140    cmbAuthorTokens.Enabled = False
50150    cmbAuthorTokens.BackColor = &H8000000F
50160  End If
50170
50180  chkUseCreationDateNow.Enabled = value
50190  chkOnePagePerFile.Enabled = value
50200  dmFraProgDocument1.Enabled = value
50210
50220  lblStampString.Enabled = value
50230  txtStampString.Enabled = value
50240  lblStampFontcolor.Enabled = value
50250  picStampFontColor.Enabled = value
50260  lblFontNameSize.Enabled = value
50270  chkStampUseOutlineFont.Enabled = value
50280  If chkStampUseOutlineFont.Enabled = True And chkStampUseOutlineFont.value = 1 Then
50290    lblOutlineFontThickness.Enabled = True
50300    txtOutlineFontThickness.Enabled = True
50310    txtOutlineFontThickness.BackColor = &H80000005
50320   Else
50330    lblOutlineFontThickness.Enabled = False
50340    txtOutlineFontThickness.Enabled = False
50350    txtOutlineFontThickness.BackColor = &H8000000F
50360  End If
50370
50380  dmFraProgStamp.Enabled = value
50390
50400  dmFraProgDocument2.Enabled = value
50410  chkUseFixPaperSize.Enabled = value
50420  If chkUseFixPaperSize.Enabled = True And chkUseFixPaperSize.value = 1 Then
50430    chkUseCustomPapersize.Enabled = True
50440    If chkUseCustomPapersize.value = 1 Then
50450      lblCustomPapersizeWidth.Enabled = True
50460      lblCustomPapersizeHeight.Enabled = True
50470      txtCustomPapersizeWidth.Enabled = True
50480      txtCustomPapersizeWidth.BackColor = &H80000005
50490      txtCustomPapersizeHeight.Enabled = True
50500      txtCustomPapersizeHeight.BackColor = &H80000005
50510      lblCustomPapersizeInfo.Enabled = True
50520      cmbDocumentPapersizes.Enabled = False
50530      cmbDocumentPapersizes.BackColor = &H8000000F
50540     Else
50550      cmbDocumentPapersizes.Enabled = True
50560      cmbDocumentPapersizes.BackColor = &H80000005
50570      lblCustomPapersizeWidth.Enabled = False
50580      lblCustomPapersizeHeight.Enabled = False
50590      txtCustomPapersizeWidth.Enabled = False
50600      txtCustomPapersizeWidth.BackColor = &H8000000F
50610      txtCustomPapersizeHeight.Enabled = False
50620      txtCustomPapersizeHeight.BackColor = &H8000000F
50630      lblCustomPapersizeInfo.Enabled = False
50640    End If
50650   Else
50660    cmbDocumentPapersizes.Enabled = False
50670    chkUseCustomPapersize.Enabled = False
50680    lblCustomPapersizeWidth.Enabled = False
50690    lblCustomPapersizeHeight.Enabled = False
50700    txtCustomPapersizeWidth.Enabled = False
50710    txtCustomPapersizeWidth.BackColor = &H8000000F
50720    txtCustomPapersizeHeight.Enabled = False
50730    txtCustomPapersizeHeight.BackColor = &H8000000F
50740    lblCustomPapersizeInfo.Enabled = False
50750  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptDocument", "SetControlsEnabled")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UserControl_Initialize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ctl As Control
50020  With tbstrProgDocument
50030   .Top = 0
50040   .Left = 0
50050   .Height = dmFraProgDocument1.Height + 50 + dmFraProgStamp.Height + 600
50060   .Visible = True
50070  End With
50080  UserControl.Width = tbstrProgDocument.Width
50090  UserControl.Height = tbstrProgDocument.Height
50100  With dmFraProgDocument1
50110   .Top = tbstrProgDocument.ClientTop + 100
50120   .Left = tbstrProgDocument.Left + (tbstrProgDocument.Width - .Width) / 2
50130  End With
50140  With dmFraProgStamp
50150   .Top = dmFraProgDocument1.Top + dmFraProgDocument1.Height + 50
50160   .Left = tbstrProgDocument.Left + (tbstrProgDocument.Width - .Width) / 2
50170  End With
50180  With dmFraProgDocument2
50190   .Top = dmFraProgDocument1.Top
50200   .Left = dmFraProgDocument1.Left
50210  End With
50220  With tbstrProgDocument.Tabs
50230   .Clear
50240   .Add
50250   .Add
50260  End With
50270  With cmbAuthorTokens
50280   .AddItem "<Computername>"
50290   .AddItem "<ClientComputer>"
50300   .AddItem "<DateTime>"
50310   .AddItem "<Title>"
50320   .AddItem "<Username>"
50330   .AddItem "<Counter>"
50340   .AddItem "<REDMON_DOCNAME>"
50350   .AddItem "<REDMON_DOCNAME_FILE>"
50360   .AddItem "<REDMON_DOCNAME_PATH>"
50370   .AddItem "<REDMON_JOB>"
50380   .AddItem "<REDMON_MACHINE>"
50390   .AddItem "<REDMON_PORT>"
50400   .AddItem "<REDMON_PRINTER>"
50410   .AddItem "<REDMON_SESSIONID>"
50420   .AddItem "<REDMON_USER>"
50430   .ListIndex = 0
50440  End With
50450  With cmbDocumentPapersizes
50460   .AddItem "11x17"
50470   .AddItem "ledger"
50480   .AddItem "legal"
50490   .AddItem "letter"
50500   .AddItem "lettersmall"
50510   .AddItem "archE"
50520   .AddItem "archD"
50530   .AddItem "archC"
50540   .AddItem "archB"
50550   .AddItem "archA"
50560   .AddItem "a0"
50570   .AddItem "a1"
50580   .AddItem "a2"
50590   .AddItem "a3"
50600   .AddItem "a4"
50610   .AddItem "a4small"
50620   .AddItem "a5"
50630   .AddItem "a6"
50640   .AddItem "a7"
50650   .AddItem "a8"
50660   .AddItem "a9"
50670   .AddItem "a10"
50680   .AddItem "isob0"
50690   .AddItem "isob1"
50700   .AddItem "isob2"
50710   .AddItem "isob3"
50720   .AddItem "isob4"
50730   .AddItem "isob5"
50740   .AddItem "isob6"
50750   .AddItem "c0"
50760   .AddItem "c1"
50770   .AddItem "c2"
50780   .AddItem "c3"
50790   .AddItem "c4"
50800   .AddItem "c5"
50810   .AddItem "c6"
50820   .AddItem "jisb0"
50830   .AddItem "jisb1"
50840   .AddItem "jisb2"
50850   .AddItem "jisb3"
50860   .AddItem "jisb4"
50870   .AddItem "jisb5"
50880   .AddItem "jisb6"
50890   .AddItem "b0"
50900   .AddItem "b1"
50910   .AddItem "b2"
50920   .AddItem "b3"
50930   .AddItem "b4"
50940   .AddItem "b5"
50950   .AddItem "flsa"
50960   .AddItem "flse"
50970   .AddItem "halfletter"
50980   .ListIndex = 0
50990  End With
51000  tbstrProgDocument.ZOrder 1
51010  tbstrProgDocument.Tabs(1).Selected = True
51020
51030  SetFrames Options.OptionsDesign
51040
51050  SetFont
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptDocument", "UserControl_Initialize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetFont()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With Options
50020   SetFontControls UserControl.Controls, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50030  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptDocument", "SetFont")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetFrames(OptionsDesign As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ctl As Control
50020  For Each ctl In UserControl.Controls
50030   If TypeOf ctl Is dmFrame Then
50040    SetFrame ctl, OptionsDesign
50050   End If
50060  Next ctl
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptDocument", "SetFrames")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetLanguageStrings()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With LanguageStrings
50020   tbstrProgDocument.Tabs(1).Caption = .OptionsProgramDocumentDescription1
50030   tbstrProgDocument.Tabs(2).Caption = .OptionsProgramDocumentDescription2
50040   dmFraProgDocument1.Caption = .OptionsProgramDocumentDescription1
50050   dmFraProgDocument2.Caption = .OptionsProgramDocumentDescription2
50060   dmFraProgStamp.Caption = .OptionsStamp
50070   chkUseStandardAuthor.Caption = .OptionsUseStandardauthor
50080   lblAuthorTokens.Caption = .OptionsStandardAuthorToken
50090   chkUseCreationDateNow.Caption = .OptionsUseCreationDateNow
50100   chkOnePagePerFile.Caption = .OptionsOnePagePerFile
50110   lblStampString.Caption = .OptionsStampString
50120   lblStampFontcolor.Caption = .OptionsStampFontColor
50130   chkStampUseOutlineFont.Caption = .OptionsStampUseOutlineFont
50140   lblOutlineFontThickness.Caption = .OptionsStampOutlineFontThickness
50150   chkUseFixPaperSize.Caption = .OptionsUseFixPapersize
50160   chkUseCustomPapersize.Caption = .OptionsUseCustomPapersize
50170   lblCustomPapersizeWidth.Caption = .OptionsCustomPapersizeWidth
50180   lblCustomPapersizeHeight.Caption = .OptionsCustomPapersizeHeight
50190   lblCustomPapersizeInfo.Caption = .OptionsCustomPapersizeInfo
50200  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptDocument", "SetLanguageStrings")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetOptions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  With Options1
50030   chkUseStandardAuthor.value = .UseStandardAuthor
50040   txtStandardAuthor.Text = .StandardAuthor
50050   chkUseCreationDateNow.value = .UseCreationDateNow
50060   chkOnePagePerFile.value = .OnePagePerFile
50070   picStampFontColor.BackColor = HTMLColorToOleColor(.StampFontColor)
50080   txtOutlineFontThickness.Text = .StampOutlineFontthickness
50090   txtStampString.Text = .StampString
50100   chkStampUseOutlineFont.value = .StampUseOutlineFont
50110   chkUseCustomPapersize.value = .UseCustomPaperSize
50120   chkUseFixPaperSize.value = .UseFixPapersize
50130   txtCustomPapersizeHeight.Text = .DeviceHeightPoints
50140   txtCustomPapersizeWidth.Text = .DeviceWidthPoints
50150   For i = 0 To cmbDocumentPapersizes.ListCount - 1
50160    If UCase$(cmbDocumentPapersizes.List(i)) = UCase$(.Papersize) Then
50170     cmbDocumentPapersizes.ListIndex = i
50180     Exit For
50190    End If
50200   Next i
50210
50220   StampFont.Name = .StampFontname
50230   StampFont.Size = .StampFontsize
50240
50250   lblFontNameSize.Caption = .StampFontname & ", " & .StampFontsize
50260  End With
50270
50280  If chkUseStandardAuthor.value = 1 And ControlsEnabled Then
50290    txtStandardAuthor.Enabled = True
50300    txtStandardAuthor.BackColor = &H80000005
50310   Else
50320    txtStandardAuthor.Enabled = False
50330    txtStandardAuthor.BackColor = &H8000000F
50340  End If
50350
50360  If lblOutlineFontThickness.Left + lblOutlineFontThickness.Width + 50 + txtOutlineFontThickness.Width > dmFraProgStamp.Width Then
50370    txtOutlineFontThickness.Left = dmFraProgStamp.Width - txtOutlineFontThickness.Width - 10
50380   Else
50390    txtOutlineFontThickness.Left = lblOutlineFontThickness.Left + lblOutlineFontThickness.Width + 50
50400  End If
50410  txtOutlineFontThickness.Top = lblOutlineFontThickness.Top + (lblOutlineFontThickness.Height - txtOutlineFontThickness.Height) / 2
50420  If chkStampUseOutlineFont.value = 1 And ControlsEnabled Then
50430    lblOutlineFontThickness.Enabled = True
50440    txtOutlineFontThickness.Enabled = True
50450    txtOutlineFontThickness.BackColor = &H80000005
50460   Else
50470    lblOutlineFontThickness.Enabled = False
50480    txtOutlineFontThickness.Enabled = False
50490    txtOutlineFontThickness.BackColor = &H8000000F
50500  End If
50510  If chkUseFixPaperSize.value = 1 And ControlsEnabled Then
50520    cmbDocumentPapersizes.Enabled = True
50530    chkUseCustomPapersize.Enabled = True
50540    If chkUseCustomPapersize.value = 1 Then
50550      lblCustomPapersizeWidth.Enabled = True
50560      lblCustomPapersizeHeight.Enabled = True
50570      txtCustomPapersizeWidth.Enabled = True
50580      txtCustomPapersizeWidth.BackColor = &H80000005
50590      txtCustomPapersizeHeight.Enabled = True
50600      txtCustomPapersizeHeight.BackColor = &H80000005
50610      lblCustomPapersizeInfo.Enabled = True
50620      cmbDocumentPapersizes.Enabled = True
50630      lblCustomPapersizeInfo.Enabled = True
50640     Else
50650      cmbDocumentPapersizes.Enabled = True
50660      lblCustomPapersizeWidth.Enabled = False
50670      lblCustomPapersizeHeight.Enabled = False
50680      txtCustomPapersizeWidth.Enabled = False
50690      txtCustomPapersizeWidth.BackColor = &H8000000F
50700      txtCustomPapersizeHeight.Enabled = False
50710      txtCustomPapersizeHeight.BackColor = &H8000000F
50720      lblCustomPapersizeInfo.Enabled = False
50730      lblCustomPapersizeInfo.Enabled = False
50740    End If
50750   Else
50760    cmbDocumentPapersizes.Enabled = False
50770    chkUseCustomPapersize.Enabled = False
50780    lblCustomPapersizeWidth.Enabled = False
50790    lblCustomPapersizeHeight.Enabled = False
50800    txtCustomPapersizeWidth.Enabled = False
50810    txtCustomPapersizeWidth.BackColor = &H8000000F
50820    txtCustomPapersizeHeight.Enabled = False
50830    txtCustomPapersizeHeight.BackColor = &H8000000F
50840    lblCustomPapersizeInfo.Enabled = False
50850  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptDocument", "SetOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub GetOptions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With Options1
50020   .OnePagePerFile = Abs(chkOnePagePerFile.value)
50030   .UseCreationDateNow = Abs(chkUseCreationDateNow.value)
50040   .StandardAuthor = txtStandardAuthor.Text
50050   .UseStandardAuthor = Abs(chkUseStandardAuthor.value)
50060   .StampString = txtStampString.Text
50070   .StampUseOutlineFont = Abs(chkStampUseOutlineFont.value)
50080   .UseCustomPaperSize = Abs(chkUseCustomPapersize.value)
50090   .UseFixPapersize = Abs(chkUseFixPaperSize.value)
50100   .StampFontColor = OleColorToHTMLColor(picStampFontColor.BackColor)
50110   .StampFontname = StampFont.Name
50120   .StampFontsize = StampFont.Size
50130   If LenB(txtOutlineFontThickness.Text) > 0 Then
50140    .StampOutlineFontthickness = txtOutlineFontThickness.Text
50150   End If
50160   If cmbDocumentPapersizes.ListCount > 0 Then
50170    If cmbDocumentPapersizes.ListIndex > 0 Then
50180     .Papersize = cmbDocumentPapersizes.List(cmbDocumentPapersizes.ListIndex)
50190    End If
50200   End If
50210   If LenB(txtCustomPapersizeHeight.Text) > 0 Then
50220    .DeviceHeightPoints = txtCustomPapersizeHeight.Text
50230   End If
50240   If LenB(txtCustomPapersizeWidth.Text) > 0 Then
50250    .DeviceWidthPoints = txtCustomPapersizeWidth.Text
50260   End If
50270  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptDocument", "GetOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub tbstrProgDocument_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case tbstrProgDocument.SelectedItem.Index
        Case 1
50030    dmFraProgDocument2.Enabled = False
50040    dmFraProgDocument2.Visible = False
50050    dmFraProgDocument1.Visible = True
50060    dmFraProgStamp.Visible = True
50070    If ControlsEnabled Then
50080     dmFraProgDocument1.Enabled = True
50090     dmFraProgStamp.Enabled = True
50100    End If
50110   Case 2
50120    dmFraProgDocument1.Enabled = False
50130    dmFraProgDocument1.Visible = False
50140    dmFraProgStamp.Enabled = False
50150    dmFraProgStamp.Visible = False
50160    dmFraProgDocument2.Visible = True
50170    If ControlsEnabled Then
50180     dmFraProgDocument2.Enabled = True
50190    End If
50200  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptDocument", "tbstrProgDocument_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtStampString_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ViewIt As Boolean
50020  If LenB(txtStampString.Text) > 0 Then
50030    ViewIt = True
50040   Else
50050    ViewIt = False
50060  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptDocument", "txtStampString_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkStampUseOutlineFont_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkStampUseOutlineFont.value = 1 Then
50020    lblOutlineFontThickness.Enabled = True
50030    txtOutlineFontThickness.Enabled = True
50040    txtOutlineFontThickness.BackColor = &H80000005
50050   Else
50060    lblOutlineFontThickness.Enabled = False
50070    txtOutlineFontThickness.Enabled = False
50080    txtOutlineFontThickness.BackColor = &H8000000F
50090  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptDocument", "chkStampUseOutlineFont_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkUseCustomPapersize_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkUseCustomPapersize.value = 1 Then
50020    lblCustomPapersizeWidth.Enabled = True
50030    lblCustomPapersizeHeight.Enabled = True
50040    txtCustomPapersizeWidth.Enabled = True
50050    txtCustomPapersizeWidth.BackColor = &H80000005
50060    txtCustomPapersizeHeight.Enabled = True
50070    txtCustomPapersizeHeight.BackColor = &H80000005
50080    lblCustomPapersizeInfo.Enabled = True
50090    cmbDocumentPapersizes.Enabled = False
50100    cmbDocumentPapersizes.BackColor = &H8000000F
50110   Else
50120    cmbDocumentPapersizes.Enabled = True
50130    cmbDocumentPapersizes.BackColor = &H80000005
50140    lblCustomPapersizeWidth.Enabled = False
50150    lblCustomPapersizeHeight.Enabled = False
50160    txtCustomPapersizeWidth.Enabled = False
50170    txtCustomPapersizeWidth.BackColor = &H8000000F
50180    txtCustomPapersizeHeight.Enabled = False
50190    txtCustomPapersizeHeight.BackColor = &H8000000F
50200    lblCustomPapersizeInfo.Enabled = False
50210  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptDocument", "chkUseCustomPapersize_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkUseFixPaperSize_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkUseFixPaperSize.value = 1 Then
50020    cmbDocumentPapersizes.Enabled = True
50030    chkUseCustomPapersize.Enabled = True
50040    If chkUseCustomPapersize.value = 1 Then
50050      lblCustomPapersizeWidth.Enabled = True
50060      lblCustomPapersizeHeight.Enabled = True
50070      txtCustomPapersizeWidth.Enabled = True
50080      txtCustomPapersizeWidth.BackColor = &H80000005
50090      txtCustomPapersizeHeight.Enabled = True
50100      txtCustomPapersizeHeight.BackColor = &H80000005
50110      lblCustomPapersizeInfo.Enabled = True
50120      cmbDocumentPapersizes.Enabled = False
50130     Else
50140      cmbDocumentPapersizes.Enabled = True
50150      lblCustomPapersizeWidth.Enabled = False
50160      lblCustomPapersizeHeight.Enabled = False
50170      txtCustomPapersizeWidth.Enabled = False
50180      txtCustomPapersizeWidth.BackColor = &H8000000F
50190      txtCustomPapersizeHeight.Enabled = False
50200      txtCustomPapersizeHeight.BackColor = &H8000000F
50210      lblCustomPapersizeInfo.Enabled = False
50220    End If
50230   Else
50240    cmbDocumentPapersizes.Enabled = False
50250    chkUseCustomPapersize.Enabled = False
50260    lblCustomPapersizeWidth.Enabled = False
50270    lblCustomPapersizeHeight.Enabled = False
50280    txtCustomPapersizeWidth.Enabled = False
50290    txtCustomPapersizeWidth.BackColor = &H8000000F
50300    txtCustomPapersizeHeight.Enabled = False
50310    txtCustomPapersizeHeight.BackColor = &H8000000F
50320    lblCustomPapersizeInfo.Enabled = False
50330  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptDocument", "chkUseFixPaperSize_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkUseStandardAuthor_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkUseStandardAuthor.value = 1 Then
50020    txtStandardAuthor.Enabled = True
50030    txtStandardAuthor.BackColor = &H80000005
50040    cmbAuthorTokens.Enabled = True
50050    cmbAuthorTokens.BackColor = &H80000005
50060    lblAuthorTokens.Enabled = True
50070   Else
50080    txtStandardAuthor.Enabled = False
50090    txtStandardAuthor.BackColor = &H8000000F
50100    cmbAuthorTokens.Enabled = False
50110    cmbAuthorTokens.BackColor = &H8000000F
50120    lblAuthorTokens.Enabled = False
50130  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptDocument", "chkUseStandardAuthor_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbAuthorTokens_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtStandardAuthor.Text = txtStandardAuthor.Text & cmbAuthorTokens.Text
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptDocument", "cmbAuthorTokens_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdStampFont_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long, Font As tFont
50020  Font.Name = Options.StampFontname
50030  Font.Size = Options.StampFontsize
50040  If OpenFontDialog(Font, UserControl.Parent.hwnd) > 0 Then
50050   StampFont.Name = Font.Name
50060   StampFont.Size = Font.Size
50070   lblFontNameSize.Caption = Font.Name & ", " & Font.Size
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptDocument", "cmdStampFont_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub picStampFontColor_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim c As OLE_COLOR
50020  If OpenColorDialog(c, UserControl.Parent.hwnd) = 1 Then
50030   picStampFontColor.BackColor = c
50040   Options.StampFontColor = OleColorToHTMLColor(c)
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptDocument", "picStampFontColor_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtCustomPapersizeHeight_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  KeyAscii = AllowedKeypressChars(KeyAscii)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptDocument", "txtCustomPapersizeHeight_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtCustomPapersizeWidth_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  KeyAscii = AllowedKeypressChars(KeyAscii)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptDocument", "txtCustomPapersizeWidth_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtOutlineFontThickness_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  KeyAscii = AllowedKeypressChars(KeyAscii)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptDocument", "txtOutlineFontThickness_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtStandardAuthor_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtStandardAuthor.ToolTipText = txtStandardAuthor.Text
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptDocument", "txtStandardAuthor_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

