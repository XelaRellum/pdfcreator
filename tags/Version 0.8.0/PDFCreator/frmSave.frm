VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSave 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Form1"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8880
   Icon            =   "frmSave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin MSComctlLib.ImageList imlFilenameIcons 
      Left            =   1800
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":058A
            Key             =   "PDF"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":0E64
            Key             =   "PNG"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":117E
            Key             =   "JPG"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":1498
            Key             =   "BMP"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":17B2
            Key             =   "PCX"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":1ACC
            Key             =   "TIFF"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":1DE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":2102
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtFilename 
      Height          =   285
      Left            =   3000
      TabIndex        =   0
      Top             =   4320
      Width           =   5775
   End
   Begin VB.ComboBox cmbFiletypes 
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown-Liste
      TabIndex        =   1
      Top             =   4680
      Width           =   5775
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Save"
      Height          =   495
      Index           =   1
      Left            =   7200
      TabIndex        =   2
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   1575
   End
   Begin MSComctlLib.TreeView tvwFolder 
      Height          =   4875
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   8599
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   0
      LabelEdit       =   1
      LineStyle       =   1
      PathSeparator   =   "|"
      Style           =   7
      SingleSel       =   -1  'True
      ImageList       =   "imlTvwIcons"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imlTvwIcons 
      Left            =   2400
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":241E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":2530
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":2642
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":2754
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":2866
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":2978
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":2A8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":2B9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":2CAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":2DC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":2ED2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":3224
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsvFilenames 
      Height          =   4095
      Left            =   3000
      TabIndex        =   5
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7223
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LastNode As MSComctlLib.Node, cmdCancel As Boolean, cDDF As clsDrivesDirsFiles, _
 tSaveFilename As String, tSaveFilepath As String

Private Sub cmbFiletypes_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Path As String, Filename As String, Ext As String
50020  SplitPath cDDF.CurrentDirectory & "\" & txtFilename.Text, , Path, , Filename, Ext
50031  Select Case cmbFiletypes.ListIndex
        Case 0:
50050    cDDF.FilePattern = "*.pdf"
50060    If UCase$(Ext) <> "PDF" Then
50070     txtFilename.Text = Filename & ".pdf"
50080    End If
50090   Case 1:
50100    cDDF.FilePattern = "*.png"
50110    If UCase$(Ext) <> "PNG" Then
50120     txtFilename.Text = Filename & ".png"
50130    End If
50140   Case 2:
50150    cDDF.FilePattern = "*.jpg"
50160    If UCase$(Ext) <> "JPG" Then
50170     txtFilename.Text = Filename & ".jpg"
50180    End If
50190   Case 3:
50200    cDDF.FilePattern = "*.bmp"
50210    If UCase$(Ext) <> "BMP" Then
50220     txtFilename.Text = Filename & ".bmp"
50230    End If
50240   Case 4:
50250    cDDF.FilePattern = "*.pcx"
50260    If UCase$(Ext) <> "PCX" Then
50270     txtFilename.Text = Filename & ".pcx"
50280    End If
50290   Case 5:
50300    cDDF.FilePattern = "*.tif"
50310    If UCase$(Ext) <> "TIF" Then
50320     txtFilename.Text = Filename & ".tif"
50330    End If
50340   Case 6:
50350    cDDF.FilePattern = "*.ps"
50360    If UCase$(Ext) <> "PS" Then
50370     txtFilename.Text = Filename & ".ps"
50380    End If
50390   Case 7:
50400    cDDF.FilePattern = "*.eps"
50410    If UCase$(Ext) <> "EPS" Then
50420     txtFilename.Text = Filename & ".eps"
50430    End If
50440  End Select
50450  RefreshFilelist
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSave", "cmbFiletypes_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmd_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case Index
        Case 1:
50030    If CanSave = False Then
50040     Exit Sub
50050    End If
50060    cmdCancel = False
50070    frmPrinting.SaveFilename = tSaveFilename
50080    frmPrinting.SaveFilterIndex = cmbFiletypes.ListIndex
50090    Options.LastSaveDirectory = tSaveFilepath
50100    SaveOptions Options
50110  End Select
50120  Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSave", "cmd_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub RefreshFilelist()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, item As ListItem, cFiles As Collection
50020  lsvFilenames.ListItems.Clear
50030  Set cFiles = cDDF.GetFiles
50040  For i = 1 To cFiles.Count
50050   Set item = lsvFilenames.ListItems.Add(, , cFiles.item(i))
50060   item.SubItems(1) = Format$(FileLen(CompletePath(cDDF.CurrentDirectory) & cFiles.item(i)) / 1024, "#,##0") & " KB"
50070   item.SubItems(2) = GetFileAttributesStr(CompletePath(cDDF.CurrentDirectory) & cFiles.item(i))
50080   item.SmallIcon = cmbFiletypes.ListIndex + 1
50090  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSave", "RefreshFilelist")
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
50030   Call HTMLHelp_ShowTopic("html\welcome.htm")
50040  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSave", "Form_KeyDown")
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
50010  Dim MyFilesName As String, MyDesktopName As String, MyDesktop As String, _
  MyFiles As String, reg As clsRegistry, res As Long, Icon As IPictureDisp, _
  cDrives As Collection, i As Long, drvS() As String, cDir As Collection
50040  Me.KeyPreview = True
50050  Screen.MousePointer = vbHourglass
50060  cmdCancel = True
50070  Set cDDF = New clsDrivesDirsFiles
50080  Set cDrives = cDDF.GetDrives(True)
50090  With LanguageStrings
50100   Me.Caption = .SaveTitle
50110   cmd(0).Caption = .SaveCancel
50120   cmd(1).Caption = .SaveSave
50130  End With
50140
50150  With lsvFilenames
50160   .View = lvwReport
50170  End With
50180  With lsvFilenames.ColumnHeaders
50190   .Clear
50200   .Add , "Filename", LanguageStrings.SaveFilename, 3200
50210   .Add , "Size", LanguageStrings.SaveSize, 1500, lvwColumnRight
50220   .Add , "Attributes", LanguageStrings.SaveAttributes, 1000, lvwColumnLeft
50230  End With
50240  With LanguageStrings
50250   cmbFiletypes.AddItem .ListPDFFiles & " (*.pdf)"
50260   cmbFiletypes.AddItem .PrintingPNGFiles & " (*.png)"
50270   cmbFiletypes.AddItem .PrintingJPEGFiles & " (*.jpg)"
50280   cmbFiletypes.AddItem .PrintingBMPFiles & " (*.bmp)"
50290   cmbFiletypes.AddItem .PrintingPCXFiles & " (*.pcx)"
50300   cmbFiletypes.AddItem .PrintingTIFFFiles & " (*.tif)"
50310   cmbFiletypes.AddItem .PrintingPSFiles & " (*.ps)"
50320   cmbFiletypes.AddItem .PrintingEPSFiles & " (*.eps)"
50330  End With
50340  Set lsvFilenames.SmallIcons = Nothing
50350  With imlFilenameIcons.ListImages
50360   res = LoadIcon(Small, "*.pdf", Icon)
50370   If res = 0 Then
50380    .Remove 1
50390    .Add 1, , Icon
50400   End If
50410   res = LoadIcon(Small, "*.png", Icon)
50420   If res = 0 Then
50430    .Remove 2
50440    .Add 2, , Icon
50450   End If
50460   res = LoadIcon(Small, "*.jpg", Icon)
50470   If res = 0 Then
50480    .Remove 3
50490    .Add 3, , Icon
50500   End If
50510   res = LoadIcon(Small, "*.bmp", Icon)
50520   If res = 0 Then
50530    .Remove 4
50540    .Add 4, , Icon
50550   End If
50560   res = LoadIcon(Small, "*.pcx", Icon)
50570   If res = 0 Then
50580    .Remove 5
50590    .Add 5, , Icon
50600   End If
50610   res = LoadIcon(Small, "*.tif", Icon)
50620   If res = 0 Then
50630    .Remove 6
50640    .Add 6, , Icon
50650   End If
50660   res = LoadIcon(Small, "*.ps", Icon)
50670   If res = 0 Then
50680    .Remove 7
50690    .Add 7, , Icon
50700   End If
50710   res = LoadIcon(Small, "*.eps", Icon)
50720   If res = 0 Then
50730    .Remove 7
50740    .Add 7, , Icon
50750   End If
50760  End With
50770  Set lsvFilenames.SmallIcons = imlFilenameIcons
50780  DoEvents
50790  cDDF.FilePattern = "*.pdf"
50800
50810  MyFilesName = GetShellNamespaceName(MYFILES_CLSID): MyFiles = GetMyFiles
50820  If CheckPath(MyFiles) = True Then
50830   tvwFolder.Nodes.Add , , "S" & MyFiles, MyFilesName, 8
50840   If IsRemovableDrive(Mid$(MyFiles, 1, 2)) = False Then
50850     If Dir(MyFiles, vbDirectory + vbHidden) <> "" Then
50860      cDDF.CurrentDirectory = MyFiles
50870      Set cDir = cDDF.GetDirectories
50880      If cDir.Count > 0 Then
50890       tvwFolder.Nodes.Add "S" & MyFiles, tvwChild
50900      End If
50910     End If
50920    Else
50930     tvwFolder.Nodes.Add "S" & MyFiles, tvwChild
50940   End If
50950  End If
50960  Set reg = New clsRegistry
50970  reg.hkey = HKEY_CLASSES_ROOT
50980  reg.KeyRoot = "CLSID"
50990  reg.Subkey = "{00021400-0000-0000-C000-000000000046}"
51000  MyDesktopName = reg.GetRegistryValue("")
51010  Set reg = Nothing
51020  tvwFolder.Nodes.Add , , "S" & GetDesktop, MyDesktopName, 2
51030  MyDesktop = GetDesktop
51040  If CheckPath(MyFiles) = True Then
51050   If IsRemovableDrive(Mid$(MyDesktop, 1, 2)) = False Then
51060     If Dir(MyDesktop, vbDirectory + vbHidden) <> "" Then
51070      cDDF.CurrentDirectory = MyDesktop
51080      Set cDir = cDDF.GetDirectories
51090      If cDir.Count > 0 Then
51100       tvwFolder.Nodes.Add "S" & MyDesktop, tvwChild
51110      End If
51120     End If
51130    Else
51140     tvwFolder.Nodes.Add "S" & MyDesktop, tvwChild
51150   End If
51160  End If
51170
51180  For i = 1 To cDrives.Count
51190   drvS = Split(cDrives.item(i), Chr$(0))
51200   tvwFolder.Nodes.Add , , "N" & UCase$(drvS(0)), UCase$(Left$(drvS(0), 1)), GetDriveIcon(drvS(0))
51210   If IsRemovableDrive(Mid$(drvS(0), 1, 2) & "\") = True Then
51220     tvwFolder.Nodes.Add "N" & UCase$(drvS(0)), tvwChild
51230    Else
51240     If IsGoodDrive(Mid$(drvS(0), 1, 2)) = True Then
51250       If Dir(Mid$(drvS(0), 1, 2) & "\", vbDirectory + vbHidden) <> "" Then
51260        cDDF.CurrentDirectory = drvS(0)
51270        Set cDir = cDDF.GetDirectories
51280        If cDir.Count > 0 Then
51290         tvwFolder.Nodes.Add "N" & UCase$(drvS(0)), tvwChild
51300        End If
51310       End If
51320      Else
51330       tvwFolder.Nodes.Add "N" & UCase$(drvS(0)), tvwChild
51340     End If
51350   End If
51360  Next i
51370  tvwFolder.Nodes("S" & MyFiles).Selected = True
51380  With txtFilename
51390   .Text = frmPrinting.SaveFilename
51400   .SelStart = 0
51410   .SelLength = Len(.Text)
51420  End With
51430  cmbFiletypes.ListIndex = 0
51440  DoEvents
51450  SetLastSaveDirectory
51460  Screen.MousePointer = vbNormal
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSave", "Form_Load")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function GetDriveIcon(Drive As String) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim DriveType As Long
50020  GetDriveIcon = 6
50030  If Len(Trim$(Drive)) > 0 Then
50040   DriveType = GetDriveType(Mid$(Drive, 1, 2))
50051   Select Case DriveType
         Case 2: 'Removable
50070     GetDriveIcon = 3
50080    Case 3: 'Fixed
50090     GetDriveIcon = 6
50100    Case 4: 'Remote
50110     GetDriveIcon = 9
50120    Case 5: 'CDRom
50130     GetDriveIcon = 1
50140    Case 6: 'Ramdisk
50150     GetDriveIcon = 7
50160   End Select
50170  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSave", "GetDriveIcon")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function IsRemovableDrive(Drive As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim DriveType As Long
50020  IsRemovableDrive = False
50030  DriveType = GetDriveType(Mid$(Drive, 1, 2))
50040  If DriveType = 2 Or DriveType = 4 Or DriveType = 5 Then
50050   IsRemovableDrive = True
50060  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSave", "IsRemovableDrive")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub Form_Unload(Cancel As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If cmdCancel = True Then
50020   frmPrinting.SaveCancel = True
50030  End If
50040  Set LastNode = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSave", "Form_Unload")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub lsvFilenames_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If lsvFilenames.ListItems.Count = 0 Then
50020   Exit Sub
50030  End If
50040  With txtFilename
50050   .Text = lsvFilenames.SelectedItem.Text
50060   .SetFocus
50070   .SelStart = 0
50080   .SelLength = Len(.Text)
50090  End With
50100  DoEvents
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSave", "lsvFilenames_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub tvwFolder_Expand(ByVal Node As MSComctlLib.Node)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  cmd(1).Enabled = False
50020  SetNode Node
50030  Call RefreshFilelist
50040  cmd(1).Enabled = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSave", "tvwFolder_Expand")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetNode(ByVal Node As MSComctlLib.Node)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 ' On Error Resume Next
50020  Dim strRelative As String, i As Long, intFolderPos As Long, strFolderName As String, _
  strNewPath As String, dirS() As String, cDir As Collection, cDir2 As Collection, _
  icInd As Long, icIndEx As Long
50050  MousePointer = vbHourglass
50060  If Node.Child.Text = "" Then
50070   tvwFolder.Nodes.Remove Node.Child.Index
50080   strRelative = Node.Key
50090   If CheckPath(Mid$(strRelative, 2)) = True Then
50100     cDDF.CurrentDirectory = Mid$(strRelative, 2)
50110     intFolderPos = Len(Mid$(strRelative, 2)) + 1
50120     Set cDir = cDDF.GetDirectories(, True)
50130     For i = 1 To cDir.Count
50140      dirS = Split(cDir.item(i), Chr$(0))
50150      strFolderName = dirS(0)
50160      strNewPath = CompletePath(CompletePath(Mid$(strRelative, 2)) & strFolderName)
50170      If (dirS(1) And vbHidden) > 0 Then
50180        icInd = 11: icIndEx = 12
50190       Else
50200        icInd = 4: icIndEx = 5
50210      End If
50220      tvwFolder.Nodes.Add strRelative, tvwChild, Mid$(strRelative, 1, 1) & strNewPath, strFolderName, icInd
50230      If CheckPath(strNewPath) = True Then
50240       cDDF.CurrentDirectory = strNewPath
50250       Set cDir2 = cDDF.GetDirectories(, True, True)
50260       If cDir2.Count > 0 Then
50270        tvwFolder.Nodes.Add Mid$(strRelative, 1, 1) & strNewPath, tvwChild, , ""
50280        tvwFolder.Nodes(Mid$(strRelative, 1, 1) & strNewPath).Image = icInd
50290        tvwFolder.Nodes(Mid$(strRelative, 1, 1) & strNewPath).ExpandedImage = icIndEx
50300       End If
50310       cDDF.CurrentDirectory = Mid$(strRelative, 2)
50320      End If
50330      DoEvents
50340     Next i
50350    Else
50360   End If
50370  End If
50380  MousePointer = vbDefault
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSave", "SetNode")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub tvwFolder_NodeClick(ByVal Node As MSComctlLib.Node)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If LastNode Is Nothing = False Then
50020   If LastNode.Image = 5 Then
50030    LastNode.Image = 4
50040   End If
50050   If LastNode.ExpandedImage = 5 Then
50060    LastNode.ExpandedImage = 4
50070   End If
50080  End If
50090  Set LastNode = Node
50100  If CheckPath(Mid$(tvwFolder.SelectedItem.Key, 2)) = False Then
50110    cmd(1).Enabled = False
50120    With lsvFilenames
50130     .Font.Italic = True
50140     .ListItems.Clear
50150     .ColumnHeaders(1).Width = 5700
50160     .ListItems.Add , , LanguageStrings.MessagesMsg07
50170    End With
50180    Exit Sub
50190   Else
50200    cmd(1).Enabled = True
50210    With lsvFilenames
50220     .Font.Italic = False
50230     .ColumnHeaders(1).Width = 3200
50240    End With
50250  End If
50260  If Node.ExpandedImage = 4 Then
50270   Node.ExpandedImage = 5
50280  End If
50290  If Node.Image = 4 Then
50300   Node.Image = 5
50310  End If
50320  cDDF.CurrentDirectory = Mid$(tvwFolder.SelectedItem.Key, 2)
50330  RefreshFilelist
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSave", "tvwFolder_NodeClick")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtFilename_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If KeyAscii = 13 Then
50020   If CanSave = False Then
50030    Exit Sub
50040   End If
50050   cmdCancel = False
50060   frmPrinting.SaveFilename = tSaveFilename
50070   frmPrinting.SaveFilterIndex = cmbFiletypes.ListIndex
50080   Options.LastSaveDirectory = tSaveFilepath
50090   SaveOptions Options
50100   Unload Me
50110  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSave", "txtFilename_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function CanSave() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Path As String, Ext As String, SaveFilename As String, res As Long, _
  tStr As String
50030  Me.MousePointer = vbHourglass
50040  CanSave = True
50050  If InStr(txtFilename.Text, ":") > 0 Then
50060    tStr = txtFilename.Text
50070   Else
50080    tStr = CompletePath(cDDF.CurrentDirectory) & txtFilename.Text
50090  End If
50100  If IsValidPath(tStr, False) = False Then
50110    MsgBox LanguageStrings.MessagesMsg10, vbOKOnly
50120    CanSave = False
50130    Me.MousePointer = vbNormal
50140    Exit Function
50150   Else
50160    If IsGoodDrive(tStr) = False Then
50170     MsgBox LanguageStrings.MessagesMsg07, vbOKOnly
50180     CanSave = False
50190     Me.MousePointer = vbNormal
50200     Exit Function
50210    End If
50220  End If
50230  SplitPath txtFilename.Text, , Path, SaveFilename, , Ext
50240  If Len(Path) > 0 Then
50250    If DirExists(Path) = False Then
50260     If Left(Path, 2) = "\\" Then
50270      MsgBox LanguageStrings.MessagesMsg10, vbOKOnly
50280      CanSave = False
50290      Me.MousePointer = vbNormal
50300      Exit Function
50310     End If
50320     res = MsgBox(LanguageStrings.MessagesMsg09, vbYesNo)
50330     If res = vbNo Then
50340      CanSave = False
50350      Me.MousePointer = vbNormal
50360      Exit Function
50370     End If
50380     If MakePath(Path) = False Then
50390      MsgBox LanguageStrings.MessagesMsg07, vbOKOnly
50400      CanSave = False
50410      Me.MousePointer = vbNormal
50420      Exit Function
50430     End If
50440    End If
50450    SaveFilename = CompletePath(Path) & SaveFilename
50460    tSaveFilepath = CompletePath(Path)
50470   Else
50480    SaveFilename = CompletePath(cDDF.CurrentDirectory) & txtFilename.Text
50490    tSaveFilepath = CompletePath(cDDF.CurrentDirectory)
50500  End If
50511  Select Case cmbFiletypes.ListIndex
        Case 0:
50530    If UCase$(Ext) <> "PDF" Then
50540     SaveFilename = SaveFilename & ".pdf"
50550     txtFilename.Text = txtFilename.Text & ".pdf"
50560    End If
50570   Case 1:
50580    If UCase$(Ext) <> "PNG" Then
50590     SaveFilename = SaveFilename & ".png"
50600     txtFilename.Text = txtFilename.Text & ".png"
50610    End If
50620   Case 2:
50630    If UCase$(Ext) <> "JPG" Then
50640     SaveFilename = SaveFilename & ".jpg"
50650     txtFilename.Text = txtFilename.Text & ".jpg"
50660    End If
50670   Case 3:
50680    If UCase$(Ext) <> "BMP" Then
50690     SaveFilename = SaveFilename & ".bmp"
50700     txtFilename.Text = txtFilename.Text & ".bmp"
50710    End If
50720   Case 4:
50730    If UCase$(Ext) <> "PCX" Then
50740     SaveFilename = SaveFilename & ".pcx"
50750     txtFilename.Text = txtFilename.Text & ".pcx"
50760    End If
50770   Case 5:
50780    If UCase$(Ext) <> "TIF" Then
50790     SaveFilename = SaveFilename & ".tif"
50800     txtFilename.Text = txtFilename.Text & ".tif"
50810    End If
50820   Case 6:
50830    If UCase$(Ext) <> "PS" Then
50840     SaveFilename = SaveFilename & ".ps"
50850     txtFilename.Text = txtFilename.Text & ".ps"
50860    End If
50870   Case 7:
50880    If UCase$(Ext) <> "EPS" Then
50890     SaveFilename = SaveFilename & ".eps"
50900     txtFilename.Text = txtFilename.Text & ".eps"
50910    End If
50920  End Select
50930  If Dir(SaveFilename) <> "" Then
50940   res = MsgBox(LanguageStrings.MessagesMsg05, vbYesNo)
50950   If res = vbNo Then
50960    CanSave = False
50970    Me.MousePointer = vbNormal
50980    Exit Function
50990   End If
51000  End If
51010  If Dir(SaveFilename) <> "" Then
51020   If GetAttr(SaveFilename) And vbReadOnly = vbReadOnly Then
51030    If GetAttr(SaveFilename) Or vbArchive = vbArchive Then
51040     res = vbArchive
51050    End If
51060    SetAttr SaveFilename, res
51070   End If
51080  End If
51090  res = vbYes
51100  If FileInUse(SaveFilename) Then
51110   MsgBox LanguageStrings.MessagesMsg34, vbExclamation
51120   CanSave = False
51130   Me.MousePointer = vbNormal
51140   Exit Function
51150  End If
51160  tSaveFilename = SaveFilename
51170  Me.MousePointer = vbNormal
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSave", "CanSave")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub SetLastSaveDirectory()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim LSDir As String, i As Long, nChild As Node, nNode As Node, tStr As String, _
  found As Boolean, Folder() As String, nNodeKey As String
50030  LSDir = UCase$(Options.LastSaveDirectory)
50040  If LSDir = "" Then
50050   LSDir = GetMyFiles
50060  End If
50070  If Right$(LSDir, 1) = "\" Then
50080   LSDir = Mid$(LSDir, 1, Len(LSDir) - 1)
50090  End If
50100  Set nNode = tvwFolder.Nodes(1)
50110  Do While Not (nNode.Parent Is Nothing)
50120   Set nNode = nNode.Parent: DoEvents
50130  Loop
50140  Set nNode = nNode.FirstSibling
50150  Do While Not (nNode Is Nothing)
50160   If Right$(nNode.Key, 1) = "\" Then
50170     nNodeKey = Mid$(nNode.Key, 2, Len(nNode.Key) - 2)
50180    Else
50190     nNodeKey = Mid$(nNode.Key, 2)
50200   End If
50210   If DirExists(LSDir) = True Then
50220    If InStr(LSDir, UCase$(nNodeKey)) > 0 Then
50230     nNode.Selected = True: DoEvents
50240     tStr = Mid$(LSDir, Len(nNodeKey) + 2)
50250     If Len(tStr) = 0 Then
50260      Exit Sub
50270     End If
50280     If InStr(tStr, "\") > 0 Then
50290       Folder = Split(tStr, "\")
50300       For i = 0 To UBound(Folder)
50310        Set nChild = tvwFolder.SelectedItem.Child
50320        Do While Not (nChild Is Nothing)
50330         If UCase$(nChild.Text) = UCase$(Folder(i)) Then
50340          tvwFolder.Nodes(nChild.Key).Selected = True
50350          If tvwFolder.Nodes(nChild.Key).Children = 0 Then
50360           tvwFolder_NodeClick tvwFolder.Nodes(nChild.Key)
50370          End If
50380          DoEvents
50390          found = True: Exit Do
50400         End If
50410         Set nChild = nChild.Next
50420        Loop
50430        If found = False Then
50440         Exit Sub
50450        End If
50460       Next i
50470       Exit Sub
50480      Else
50490       Set nChild = nNode.Child
50500       found = False
50510       Do While Not (nChild Is Nothing)
50520        If UCase$(nChild.Text) = UCase$(tStr) Then
50530         tvwFolder.Nodes(nChild.Key).Selected = True: DoEvents
50540         found = True: Exit Sub
50550        End If
50560        Set nChild = nChild.Next
50570       Loop
50580       If found = False Then
50590        Exit Sub
50600       End If
50610     End If
50620    End If
50630   End If
50640   Set nNode = nNode.Next: DoEvents
50650  Loop
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSave", "SetLastSaveDirectory")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
