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
            Picture         =   "frmSave.frx":08CA
            Key             =   "PDF"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":11A4
            Key             =   "PNG"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":14BE
            Key             =   "JPG"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":17D8
            Key             =   "BMP"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":1AF2
            Key             =   "PCX"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":1E0C
            Key             =   "TIFF"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":2126
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":2442
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
   Begin VB.DirListBox Dir1 
      Height          =   540
      Left            =   3960
      TabIndex        =   7
      Top             =   5160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   5520
      TabIndex        =   6
      Top             =   5160
      Visible         =   0   'False
      Width           =   1455
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
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":275E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":2870
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":2982
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":2A94
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":2BA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":2CB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":2DCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":2EDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":2FEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":3100
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

Private LastNode As MSComctlLib.Node, cmdCancel As Boolean, tSaveFilename As String, tSaveFilepath As String

Private Sub cmbFiletypes_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Path As String, FileName As String, Ext As String
50020  SplitPath Dir1.Path & "\" & txtFilename.Text, , Path, , FileName, Ext
50030  Select Case cmbFiletypes.ListIndex
  Case 0:
50050    File1.Pattern = "*.pdf"
50060    If UCase$(Ext) <> "PDF" Then
50070     txtFilename.Text = FileName & ".pdf"
50080    End If
50090   Case 1:
50100    File1.Pattern = "*.png"
50110    If UCase$(Ext) <> "PNG" Then
50120     txtFilename.Text = FileName & ".png"
50130    End If
50140   Case 2:
50150    File1.Pattern = "*.jpg"
50160    If UCase$(Ext) <> "JPG" Then
50170     txtFilename.Text = FileName & ".jpg"
50180    End If
50190   Case 3:
50200    File1.Pattern = "*.bmp"
50210    If UCase$(Ext) <> "BMP" Then
50220     txtFilename.Text = FileName & ".bmp"
50230    End If
50240   Case 4:
50250    File1.Pattern = "*.pcx"
50260    If UCase$(Ext) <> "PCX" Then
50270     txtFilename.Text = FileName & ".pcx"
50280    End If
50290   Case 5:
50300    File1.Pattern = "*.tif"
50310    If UCase$(Ext) <> "TIF" Then
50320     txtFilename.Text = FileName & ".tif"
50330    End If
50340   Case 6:
50350    File1.Pattern = "*.ps"
50360    If UCase$(Ext) <> "PS" Then
50370     txtFilename.Text = FileName & ".ps"
50380    End If
50390   Case 7:
50400    File1.Pattern = "*.eps"
50410    If UCase$(Ext) <> "EPS" Then
50420     txtFilename.Text = FileName & ".eps"
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
50010  Select Case Index
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
50020  Dim i As Long, item As ListItem
50030  lsvFilenames.ListItems.Clear
50040  For i = 0 To File1.ListCount - 1
50050   Set item = lsvFilenames.ListItems.Add(, , File1.List(i))
50060   item.SubItems(1) = Format$(FileLen(CompletePath(File1.Path) & File1.List(i)) / 1024, "#,##0") & " KB"
50070   item.SubItems(2) = GetFileAttributesStr(CompletePath(File1.Path) & File1.List(i))
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

Private Sub Dir1_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  File1.Path = Dir1.Path
50020  Call RefreshFilelist
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSave", "Dir1_Change")
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
50010  Dim i As Long
50020  Dim MyFilesName As String, MyDesktopName As String, MyDesktop As String, _
  MyFiles As String, reg As clsRegistry, res As Long, Icon As IPictureDisp, _
  cDrives As Collection
50050  Screen.MousePointer = vbHourglass
50060  cmdCancel = True
50070  Set cDrives = GetDrives
50080  With LanguageStrings
50090   Me.Caption = .SaveTitle
50100   cmd(0).Caption = .SaveCancel
50110   cmd(1).Caption = .SaveSave
50120  End With
50130
50140  With lsvFilenames
50150   .View = lvwReport
50160  End With
50170  With lsvFilenames.ColumnHeaders
50180   .Clear
50190   .Add , "Filename", LanguageStrings.SaveFilename, 3200
50200   .Add , "Size", LanguageStrings.SaveSize, 1500, lvwColumnRight
50210   .Add , "Attributes", LanguageStrings.SaveAttributes, 1000, lvwColumnLeft
50220  End With
50230  With LanguageStrings
50240   cmbFiletypes.AddItem .ListPDFFiles & " (*.pdf)"
50250   cmbFiletypes.AddItem .PrintingPNGFiles & " (*.png)"
50260   cmbFiletypes.AddItem .PrintingJPEGFiles & " (*.jpg)"
50270   cmbFiletypes.AddItem .PrintingBMPFiles & " (*.bmp)"
50280   cmbFiletypes.AddItem .PrintingPCXFiles & " (*.pcx)"
50290   cmbFiletypes.AddItem .PrintingTIFFFiles & " (*.tif)"
50300   cmbFiletypes.AddItem .PrintingPSFiles & " (*.ps)"
50310   cmbFiletypes.AddItem .PrintingEPSFiles & " (*.eps)"
50320  End With
50330  Set lsvFilenames.SmallIcons = Nothing
50340  With imlFilenameIcons.ListImages
50350   res = LoadIcon(Small, "*.pdf", Icon)
50360   If res = 0 Then
50370    .Remove 1
50380    .Add 1, , Icon
50390   End If
50400   res = LoadIcon(Small, "*.png", Icon)
50410   If res = 0 Then
50420    .Remove 2
50430    .Add 2, , Icon
50440   End If
50450   res = LoadIcon(Small, "*.jpg", Icon)
50460   If res = 0 Then
50470    .Remove 3
50480    .Add 3, , Icon
50490   End If
50500   res = LoadIcon(Small, "*.bmp", Icon)
50510   If res = 0 Then
50520    .Remove 4
50530    .Add 4, , Icon
50540   End If
50550   res = LoadIcon(Small, "*.pcx", Icon)
50560   If res = 0 Then
50570    .Remove 5
50580    .Add 5, , Icon
50590   End If
50600   res = LoadIcon(Small, "*.tif", Icon)
50610   If res = 0 Then
50620    .Remove 6
50630    .Add 6, , Icon
50640   End If
50650   res = LoadIcon(Small, "*.ps", Icon)
50660   If res = 0 Then
50670    .Remove 7
50680    .Add 7, , Icon
50690   End If
50700   res = LoadIcon(Small, "*.eps", Icon)
50710   If res = 0 Then
50720    .Remove 7
50730    .Add 7, , Icon
50740   End If
50750  End With
50760  Set lsvFilenames.SmallIcons = imlFilenameIcons
50770  DoEvents
50780  File1.Pattern = "*.pdf"
50790
50800  MyFilesName = GetShellNamespaceName(MYFILES_CLSID): MyFiles = GetMyFiles
50810  If CheckPath(MyFiles) = True Then
50820   tvwFolder.Nodes.Add , , "S" & MyFiles, MyFilesName, 8
50830   If IsRemovableDrive(Mid$(MyFiles, 1, 2)) = False Then
50840     If Dir(MyFiles, vbDirectory) <> "" Then
50850      Dir1.Path = MyFiles
50860      If Dir1.ListCount > 0 Then
50870       tvwFolder.Nodes.Add "S" & MyFiles, tvwChild
50880      End If
50890     End If
50900    Else
50910     tvwFolder.Nodes.Add "S" & MyFiles, tvwChild
50920   End If
50930  End If
50940  Set reg = New clsRegistry
50950  reg.hkey = HKEY_CLASSES_ROOT
50960  reg.KeyRoot = "CLSID"
50970  reg.Subkey = "{00021400-0000-0000-C000-000000000046}"
50980  MyDesktopName = reg.GetRegistryValue("")
50990  Set reg = Nothing
51000  tvwFolder.Nodes.Add , , "S" & GetDesktop, MyDesktopName, 2
51010  MyDesktop = GetDesktop
51020  If CheckPath(MyFiles) = True Then
51030   If IsRemovableDrive(Mid$(MyDesktop, 1, 2)) = False Then
51040     If Dir(MyDesktop, vbDirectory) <> "" Then
51050      Dir1.Path = MyDesktop
51060      If Dir1.ListCount > 0 Then
51070       tvwFolder.Nodes.Add "S" & MyDesktop, tvwChild
51080      End If
51090     End If
51100    Else
51110     tvwFolder.Nodes.Add "S" & MyDesktop, tvwChild
51120   End If
51130  End If
51140
51150  For i = 1 To cDrives.Count
51160   tvwFolder.Nodes.Add , , "N" & UCase$(cDrives.item(i)), UCase$(Left$(cDrives.item(i), 1)), GetDriveIcon(cDrives.item(i))
51170 '51170   If IsRemovableDrive(Mid$(cDrives.item(i), 1, 2)) = False Then
51180   If IsRemovableDrive(Mid$(cDrives.item(i), 1, 2) & "\") = True Then
51190     tvwFolder.Nodes.Add "N" & UCase$(cDrives.item(i)), tvwChild
51200    Else
51210     If IsGoodDrive(Mid$(cDrives.item(i), 1, 2)) = True Then
51220       If Dir(Mid$(cDrives.item(i), 1, 2) & "\", vbDirectory) <> "" Then
51230        Dir1.Path = cDrives.item(i)
51240        If Dir1.ListCount > 0 Then
51250         tvwFolder.Nodes.Add "N" & UCase$(cDrives.item(i)), tvwChild
51260        End If
51270       End If
51280      Else
51290       tvwFolder.Nodes.Add "N" & UCase$(cDrives.item(i)), tvwChild
51300     End If
51310   End If
51320  Next i
51330  tvwFolder.Nodes("S" & MyFiles).Selected = True
51340  With txtFilename
51350   .Text = frmPrinting.SaveFilename
51360   .SelStart = 0
51370   .SelLength = Len(.Text)
51380  End With
51390  cmbFiletypes.ListIndex = 0
51400  DoEvents
51410  SetLastSaveDirectory
51420  Screen.MousePointer = vbNormal
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
50050   Select Case DriveType
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
50010  With txtFilename
50020   .Text = lsvFilenames.SelectedItem.Text
50030   .SetFocus
50040   .SelStart = 0
50050   .SelLength = Len(.Text)
50060  End With
50070  DoEvents
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
50010  SetNode Node
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
 On Error Resume Next
 Dim strRelative As String, i As Long, intFolderPos As Long, strFolderName As String, _
  strNewPath As String
 MousePointer = vbHourglass
 If Node.Child.Text = "" Then
  tvwFolder.Nodes.Remove Node.Child.Index
  strRelative = Node.Key
  If CheckPath(Mid$(strRelative, 2)) = True Then
    Dir1.Path = Mid$(strRelative, 2)
    intFolderPos = Len(Mid$(strRelative, 2)) + 1
    For i = 0 To Dir1.ListCount - 1
     strFolderName = Mid$(Dir1.List(i), intFolderPos)
     strNewPath = Mid$(strRelative, 2) & strFolderName & "\"
     tvwFolder.Nodes.Add strRelative, tvwChild, Mid$(strRelative, 1, 1) & strNewPath, strFolderName, 4
     If CheckPath(strNewPath) = True Then
      Dir1.Path = strNewPath
      If Dir1.ListCount > 0 Then
       tvwFolder.Nodes.Add Mid$(strRelative, 1, 1) & strNewPath, tvwChild, , ""
       tvwFolder.Nodes(Mid$(strRelative, 1, 1) & strNewPath).Image = 4
       tvwFolder.Nodes(Mid$(strRelative, 1, 1) & strNewPath).ExpandedImage = 5
      End If
      Dir1.Path = Mid$(strRelative, 2)
     End If
    Next i
   Else
  End If
 End If
 MousePointer = vbDefault
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
50320  Dir1.Path = Mid$(tvwFolder.SelectedItem.Key, 2)
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
50080    tStr = CompletePath(Dir1.Path) & txtFilename.Text
50090  End If
50100  If IsValidPath(tStr, False) = False Then
50110    MsgBox LanguageStrings.MessagesMsg10, vbOKOnly
50120    CanSave = False
50130    Exit Function
50140   Else
50150    If IsGoodDrive(tStr) = False Then
50160     MsgBox LanguageStrings.MessagesMsg07, vbOKOnly
50170     CanSave = False
50180     Exit Function
50190    End If
50200  End If
50210  SplitPath txtFilename.Text, , Path, SaveFilename, , Ext
50220  If Len(Path) > 0 Then
50230    If Len(Dir(Path, vbDirectory)) = 0 Then
50240     res = MsgBox(LanguageStrings.MessagesMsg09, vbYesNo)
50250     If res = vbNo Then
50260      CanSave = False
50270      Exit Function
50280     End If
50290     MakePath Path
50300    End If
50310    SaveFilename = CompletePath(Path) & SaveFilename
50320    tSaveFilepath = CompletePath(Path)
50330   Else
50340    SaveFilename = CompletePath(Dir1.Path) & txtFilename.Text
50350    tSaveFilepath = CompletePath(Dir1.Path)
50360  End If
50370  Select Case cmbFiletypes.ListIndex
  Case 0:
50390    If UCase$(Ext) <> "PDF" Then
50400     SaveFilename = SaveFilename & ".pdf"
50410     txtFilename.Text = txtFilename.Text & ".pdf"
50420    End If
50430   Case 1:
50440    If UCase$(Ext) <> "PNG" Then
50450     SaveFilename = SaveFilename & ".png"
50460     txtFilename.Text = txtFilename.Text & ".png"
50470    End If
50480   Case 2:
50490    If UCase$(Ext) <> "JPG" Then
50500     SaveFilename = SaveFilename & ".jpg"
50510     txtFilename.Text = txtFilename.Text & ".jpg"
50520    End If
50530   Case 3:
50540    If UCase$(Ext) <> "BMP" Then
50550     SaveFilename = SaveFilename & ".bmp"
50560     txtFilename.Text = txtFilename.Text & ".bmp"
50570    End If
50580   Case 4:
50590    If UCase$(Ext) <> "PCX" Then
50600     SaveFilename = SaveFilename & ".pcx"
50610     txtFilename.Text = txtFilename.Text & ".pcx"
50620    End If
50630   Case 5:
50640    If UCase$(Ext) <> "TIF" Then
50650     SaveFilename = SaveFilename & ".tif"
50660     txtFilename.Text = txtFilename.Text & ".tif"
50670    End If
50680   Case 6:
50690    If UCase$(Ext) <> "PS" Then
50700     SaveFilename = SaveFilename & ".ps"
50710     txtFilename.Text = txtFilename.Text & ".ps"
50720    End If
50730   Case 7:
50740    If UCase$(Ext) <> "EPS" Then
50750     SaveFilename = SaveFilename & ".eps"
50760     txtFilename.Text = txtFilename.Text & ".eps"
50770    End If
50780  End Select
50790  If Dir(SaveFilename) <> "" Then
50800   res = MsgBox(LanguageStrings.MessagesMsg05, vbYesNo)
50810   If res = vbNo Then
50820    CanSave = False
50830    Exit Function
50840   End If
50850  End If
50860  If Dir(SaveFilename) <> "" Then
50870   If GetAttr(SaveFilename) And vbReadOnly = vbReadOnly Then
50880    If GetAttr(SaveFilename) Or vbArchive = vbArchive Then
50890     res = vbArchive
50900    End If
50910    SetAttr SaveFilename, res
50920   End If
50930  End If
50940  tSaveFilename = SaveFilename
50950  Me.MousePointer = vbNormal
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
50210   If Dir(LSDir, vbDirectory) <> "" Then
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
