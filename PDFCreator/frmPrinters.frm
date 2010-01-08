VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPrinters 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Form1"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6780
   Icon            =   "frmPrinters.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows-Standard
   Begin PDFCreator.dmFrame dmFraPrinters 
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6495
      _extentx        =   11456
      _extenty        =   8281
      caption         =   "Printers"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmPrinters.frx":628A
      Begin VB.TextBox txtNewPrinter 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   3720
         Width           =   2895
      End
      Begin MSComctlLib.ImageList imlPrinters 
         Left            =   2400
         Top             =   600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   20
         ImageHeight     =   20
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrinters.frx":62B6
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cmbProfile 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown-Liste
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddPrinter 
         Caption         =   "Add printer"
         Enabled         =   0   'False
         Height          =   555
         Left            =   120
         TabIndex        =   5
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CommandButton cmdDelPrinter 
         Caption         =   "Del printer"
         Height          =   555
         Left            =   4920
         TabIndex        =   4
         Top             =   4080
         Width           =   1455
      End
      Begin MSComctlLib.ListView lsvPrinters 
         Height          =   2655
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imlPrinters"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblAdminNotice 
         Caption         =   "You must be an administrator to install or delete a printer!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   3120
         Visible         =   0   'False
         Width           =   6210
      End
      Begin VB.Label lblNewPrinterName 
         AutoSize        =   -1  'True
         Caption         =   "New printer name"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   3480
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   1575
   End
End
Attribute VB_Name = "frmPrinters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PrinterProfiles As Collection, PPrinters As Collection

Public Sub ChangeLanguage()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With LanguageStrings
50020   dmFraPrinters.Caption = .PrintersPrinters
50030   cmdDelPrinter.Caption = .PrintersPrinterDel
50040   cmdAddPrinter.Caption = .PrintersPrinterAdd
50050   cmdClose.Caption = .PrintersClose
50060   cmdSave.Caption = .PrintersSave
50070   lsvPrinters.ColumnHeaders(1).Text = .PrintersPrinter
50080   lsvPrinters.ColumnHeaders(2).Text = .PrintersProfile
50090   lblAdminNotice.Caption = .PrintersAdminNotice
50100   lblNewPrinterName.Caption = .PrintersNewPrinterName
50110   cmbProfile.List(0) = .OptionsProfileDefaultName
50120  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinters", "ChangeLanguage")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbProfile_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If lsvPrinters.ListItems.count > 0 Then
50020   lsvPrinters.ListItems(lsvPrinters.SelectedItem.Index).ListSubItems(1).Text = cmbProfile.Text
50030  End If
50040  cmbProfile.Visible = False
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinters", "cmbProfile_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub AddPrinter()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Printername As String, c As Long, lItem As ListItem
50020
50030  txtNewPrinter.Text = Trim$(txtNewPrinter.Text)
50040  Printername = txtNewPrinter.Text
50050
50060  If Printername <> vbNullString Then
50070   If LenB(Printername) > 0 Then
50080    If PrinterIsInstalled(Printername) Then
50090      MsgBox LanguageStrings.MessagesMsg40
50100      Exit Sub
50110     Else
50120      c = Printers.count
50130      Call InstallWindowsPrinter("PDFCreator", "PDFCreator:", "PDFCreator", Printername, "", App.Path)
50140      If (Printers.count > c) Then
50150       Set lItem = lsvPrinters.ListItems.Add(, "K" & Printername, Printername, , 1)
50160       lItem.SubItems(1) = LanguageStrings.OptionsProfileDefaultName
50170       lItem.Selected = True
50180       If lsvPrinters.ListItems.count > 1 And cmdDelPrinter.Enabled = False Then
50190        cmdDelPrinter.Enabled = True
50200       End If
50210      End If
50220    End If
50230   End If
50240  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinters", "AddPrinter")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdAddPrinter_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim count As Long
50020  count = lsvPrinters.ListItems.count
50030  AddPrinter
50040  If count <> lsvPrinters.ListItems.count And cmdClose.Enabled = True Then
50050   RemoveX Me
50060   cmdClose.Enabled = False
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinters", "cmdAddPrinter_Click")
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
50010  Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinters", "cmdClose_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdDelPrinter_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim c As Long, count As Long
50020  count = lsvPrinters.ListItems.count
50030  If lsvPrinters.ListItems.count > 0 Then
50040   c = Printers.count
50050   UnInstallWindowsPrinter "PDFCreator", "PDFCreator:", "PDFCreator", lsvPrinters.SelectedItem.Text, "", True
50060   If c > Printers.count Then
50070    cmbProfile.Visible = False
50080    lsvPrinters.ListItems.Remove lsvPrinters.SelectedItem.Index
50090    If lsvPrinters.ListItems.count > 0 Then
50100     lsvPrinters.ListItems(lsvPrinters.SelectedItem.Index).Selected = True
50110    End If
50120    If lsvPrinters.ListItems.count <= 1 Then
50130     cmdDelPrinter.Enabled = False
50140    End If
50150   End If
50160  End If
50170  If count <> lsvPrinters.ListItems.count And cmdClose.Enabled = True Then
50180   RemoveX Me
50190   cmdClose.Enabled = False
50200  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinters", "cmdDelPrinter_Click")
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
50010  Dim PrinterProfiles  As Collection, i As Long, sa(1) As String
50020  Set PrinterProfiles = New Collection
50030
50040  For i = 1 To lsvPrinters.ListItems.count
50050   If LCase$(LanguageStrings.OptionsProfileDefaultName) = LCase$(lsvPrinters.ListItems(i).SubItems(1)) Then
50060     sa(1) = ""
50070    Else
50080     sa(1) = lsvPrinters.ListItems(i).SubItems(1)
50090   End If
50100   sa(0) = lsvPrinters.ListItems(i).Text
50110   PrinterProfiles.Add sa
50120  Next i
50130
50140  SavePrinterProfiles PrinterProfiles
50150  Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinters", "cmdSave_Click")
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
50010  Dim tStr As String
50020  lsvPrinters.ColumnHeaders.Clear
50030  tStr = "Printer"
50040  lsvPrinters.ColumnHeaders.Add , tStr, tStr, lsvPrinters.Width / 2 - 100
50050  tStr = "Profile"
50060  lsvPrinters.ColumnHeaders.Add , tStr, tStr, lsvPrinters.Width / 2 - 100
50070  Init
50080  ChangeLanguage
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinters", "Form_Load")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetPrinterProfiles()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Profiles As Collection, profile As Variant, p As Variant, i As Long, j As Long
50020
50030  lsvPrinters.ListItems.Clear
50040
50050  Set PPrinters = New Collection
50060  Set PrinterProfiles = GetPrinterProfiles
50070
50080  For Each p In GetPDFCreatorPrinters
50090   lsvPrinters.ListItems.Add , "K" & p, p, , 1
50100  Next p
50110
50120  For i = 1 To lsvPrinters.ListItems.count
50130   For j = 1 To PrinterProfiles.count
50140    If UCase$(lsvPrinters.ListItems(i).Text) = UCase$(PrinterProfiles(j)(0)) Then
50150     If ProfileExists(PrinterProfiles(j)(1)) Then
50160      lsvPrinters.ListItems(i).SubItems(1) = PrinterProfiles(j)(1)
50170     End If
50180     Exit For
50190    End If
50200   Next j
50210  Next i
50220
50230  For i = 1 To lsvPrinters.ListItems.count
50240   If LenB(lsvPrinters.ListItems(i).SubItems(1)) = 0 Then
50250    lsvPrinters.ListItems(i).SubItems(1) = LanguageStrings.OptionsProfileDefaultName
50260   End If
50270  Next i
50280
50290  cmbProfile.Clear
50300  cmbProfile.AddItem LanguageStrings.OptionsProfileDefaultName
50310
50320  Set Profiles = GetProfiles
50330  For Each profile In Profiles
50340   cmbProfile.AddItem profile
50350  Next profile
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinters", "SetPrinterProfiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Init()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Me.Caption = App.ProductName
50020
50030  If Not IsAdmin Then
50040   lblNewPrinterName.Enabled = False
50050   txtNewPrinter.Enabled = False
50060   cmdDelPrinter.Enabled = False
50070   cmdAddPrinter.Enabled = False
50080   lblAdminNotice.Visible = True
50090  End If
50100  SetPrinterProfiles
50110  If lsvPrinters.ListItems.count > 1 Then
50120    If lsvPrinters.SelectedItem.Index <= 0 Then
50130     cmdDelPrinter.Enabled = False
50140    End If
50150   Else
50160    cmdDelPrinter.Enabled = False
50170  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinters", "Init")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub lsvPrinters_ItemClick(ByVal Item As MSComctlLib.ListItem)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020
50030  With cmbProfile
50040   .Width = lsvPrinters.ColumnHeaders(2).Width
50050   .Left = lsvPrinters.ColumnHeaders(2).Left + 190
50060   .Top = lsvPrinters.Top + Item.Top + 40
50070   For i = 1 To .ListCount
50080    If UCase$(.List(i - 1)) = UCase$(Item.ListSubItems(1).Text) Then
50090     .ListIndex = i - 1
50100    End If
50110   Next i
50120 '  .Text = Item.ListSubItems(1).Text
50130   .Visible = True
50140   .SetFocus
50150  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinters", "lsvPrinters_ItemClick")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function PrinterExists(Printername As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim p As Printer
50020  For Each p In Printers
50030   If StrComp(p.DeviceName, Printername, vbTextCompare) = 0 Then
50040    PrinterExists = True
50050    Exit Function
50060   End If
50070  Next
50080  PrinterExists = False
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinters", "PrinterExists")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub txtNewPrinter_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If LenB(txtNewPrinter.Text) > 0 Then
50020    If PrinterExists(Trim$(txtNewPrinter.Text)) = True Then
50030      cmdAddPrinter.Enabled = False
50040     Else
50050      cmdAddPrinter.Enabled = True
50060    End If
50070   Else
50080    cmdAddPrinter.Enabled = False
50090  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinters", "txtNewPrinter_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtNewPrinter_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If Len(txtNewPrinter.Text) = 128 Then
50020   KeyAscii = 0
50030   Exit Sub
50040  End If
50051  Select Case KeyAscii
        Case Asc("!"), Asc("\"), Asc(",")
50070    KeyAscii = 0
50080   Case Else
50090    If KeyAscii = 13 Then
50100     If cmdAddPrinter.Enabled Then
50110      AddPrinter
50120     End If
50130    End If
50140  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmPrinters", "txtNewPrinter_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

