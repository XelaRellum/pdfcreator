VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSaveOpen 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "PDFCreator"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8880
   Icon            =   "frmSaveOpen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3120
      Top             =   5160
   End
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
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSaveOpen.frx":058A
            Key             =   "PDF"
            Object.Tag             =   "*.PDF"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSaveOpen.frx":0E64
            Key             =   "PNG"
            Object.Tag             =   "*.PNG"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSaveOpen.frx":117E
            Key             =   "JPG"
            Object.Tag             =   "*.JPG"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSaveOpen.frx":1498
            Key             =   "BMP"
            Object.Tag             =   "*.BMP"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSaveOpen.frx":17B2
            Key             =   "PCX"
            Object.Tag             =   "*.PCX"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSaveOpen.frx":1ACC
            Key             =   "TIFF"
            Object.Tag             =   "*.TIF"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSaveOpen.frx":1DE6
            Key             =   "PS"
            Object.Tag             =   "*.PS"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSaveOpen.frx":2102
            Key             =   "EPS"
            Object.Tag             =   "*.EPS"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSaveOpen.frx":241E
            Key             =   "All"
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
      Caption         =   "&Save/Open"
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
            Picture         =   "frmSaveOpen.frx":30F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSaveOpen.frx":320A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSaveOpen.frx":331C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSaveOpen.frx":342E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSaveOpen.frx":3540
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSaveOpen.frx":3652
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSaveOpen.frx":3764
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSaveOpen.frx":3876
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSaveOpen.frx":3988
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSaveOpen.frx":3A9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSaveOpen.frx":3BAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSaveOpen.frx":3EFE
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
Attribute VB_Name = "frmSaveOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Filename As String, Filter As String, Filterindex As Long, SaveOpenType As eSaveOpenType

Private LastNode As MSComctlLib.Node, cmdCancel As Boolean, cDDF As clsDrivesDirsFiles, _
 tSaveFilename As String, tSaveFilepath As String, sFilter() As String

Private Sub cmbFiletypes_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Path As String, tFilename As String, Ext As String, Ext2 As String
50020  If InStr(Filter, "|") > 0 Then
50030   SplitPath cDDF.CurrentDirectory & "\" & sFilter(cmbFiletypes.ListIndex * 2 + 1), , Path, , tFilename, Ext2
50040   SplitPath cDDF.CurrentDirectory & "\" & txtFilename.Text, , Path, , tFilename, Ext
50050   cDDF.FilePattern = sFilter(cmbFiletypes.ListIndex * 2 + 1)
50060   If UCase$(Ext) <> UCase$(Ext2) Then
50070    If SaveOpenType = saveFile Then
50080     txtFilename.Text = tFilename & "." & Ext2
50090    End If
50100   End If
50110   RefreshFilelist
50120  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSaveOpen", "cmbFiletypes_Click")
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
50070    SaveOpenCancel = False
50080    Set SaveOpenFilename = New Collection
50090    SaveOpenFilename.Add tSaveFilename
50100    SaveOpenFilterindex = cmbFiletypes.ListIndex + 1
50110    Options.LastSaveDirectory = tSaveFilepath
50120    SaveOptions Options
50130  End Select
50140  Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSaveOpen", "cmd_Click")
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
50010  Dim i As Long, j As Long, item As ListItem, cFiles As Collection, _
  Ext As String
50030
50040  lsvFilenames.ListItems.Clear
50050  Set cFiles = cDDF.GetFiles
50060  For i = 1 To cFiles.Count
50070   Set item = lsvFilenames.ListItems.Add(, , cFiles.item(i))
50080   SplitPath CutNull(cFiles.item(i)), , , , , Ext
50090   item.SubItems(1) = Format$(FileLen(CompletePath(cDDF.CurrentDirectory) & CutNull(cFiles.item(i))) / 1024, "#,##0") & " KB"
50100   item.SubItems(2) = GetFileAttributesStr(CompletePath(cDDF.CurrentDirectory) & cFiles.item(i))
50110   item.SmallIcon = 9
50120   For j = 1 To imlFilenameIcons.ListImages.Count
50130    If "*." & UCase$(Ext) = UCase$(imlFilenameIcons.ListImages(j).Tag) Then
50140     item.SmallIcon = j
50150     Exit For
50160    End If
50170   Next j
50180 '  item.SmallIcon = cmbFiletypes.ListIndex + 1
50190  Next i
50200  If lsvFilenames.ListItems.Count > 0 And SaveOpenType = OpenFile Then
50210   txtFilename.Text = lsvFilenames.ListItems(1).Text
50220  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSaveOpen", "RefreshFilelist")
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
Select Case ErrPtnr.OnError("frmSaveOpen", "Form_KeyDown")
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
  cDrives As Collection, i As Long, j As Long, drvS() As String, _
  cDir As Collection, tStr As String
50050  Me.KeyPreview = True
50060  Screen.MousePointer = vbHourglass
50070  cmdCancel = True
50080  SaveOpenCancel = True
50090  Set cDDF = New clsDrivesDirsFiles
50100  Set cDrives = cDDF.GetDrives(True)
50110  With LanguageStrings
50120   cmd(0).Caption = .SaveOpenCancel
50130   If SaveOpenType = 0 Then
50140    Me.Caption = .SaveOpenSaveTitle
50150    cmd(1).Caption = .SaveOpenSave
50160   End If
50170   If SaveOpenType = 1 Then
50180    Me.Caption = .SaveOpenOpenTitle
50190    cmd(1).Caption = .SaveOpenOpen
50200   End If
50210  End With
50220
50230  With lsvFilenames
50240   .View = lvwReport
50250  End With
50260  With lsvFilenames.ColumnHeaders
50270   .Clear
50280   .Add , "Filename", LanguageStrings.SaveOpenFilename, 3200
50290   .Add , "Size", LanguageStrings.SaveOpenSize, 1500, lvwColumnRight
50300   .Add , "Attributes", LanguageStrings.SaveOpenAttributes, 1000, lvwColumnLeft
50310  End With
50320  If InStr(Filter, "|") > 0 Then
50330   sFilter = Split(Filter, "|")
50340   Set lsvFilenames.SmallIcons = Nothing
50350   For i = LBound(sFilter) To UBound(sFilter) Step 2
50360    cmbFiletypes.AddItem sFilter(i)
50370    For j = 1 To imlFilenameIcons.ListImages.Count
50380     res = LoadIcon(Small, imlFilenameIcons.ListImages(j).Tag, Icon)
50390     If res = 0 Then
50400      tStr = imlFilenameIcons.ListImages(j).Tag
50410      imlFilenameIcons.ListImages.Remove j
50420      imlFilenameIcons.ListImages.Add j, , Icon
50430      imlFilenameIcons.ListImages(j).Tag = UCase$(tStr)
50440     End If
50450    Next j
50460   Next i
50470   Set lsvFilenames.SmallIcons = imlFilenameIcons
50480  End If
50490  DoEvents
50500  cDDF.FilePattern = sFilter(2 * Filterindex + 1)
50510
50520  MyFilesName = GetShellNamespaceName(MYFILES_CLSID): MyFiles = GetMyFiles
50530  If CheckPath(MyFiles) = True Then
50540   tvwFolder.Nodes.Add , , "S" & MyFiles, MyFilesName, 8
50550   If IsRemovableDrive(Mid$(MyFiles, 1, 2)) = False Then
50560     If Dir(MyFiles, vbDirectory + vbHidden) <> "" Then
50570      cDDF.CurrentDirectory = MyFiles
50580      Set cDir = cDDF.GetDirectories
50590      If cDir.Count > 0 Then
50600       tvwFolder.Nodes.Add "S" & MyFiles, tvwChild
50610      End If
50620     End If
50630    Else
50640     tvwFolder.Nodes.Add "S" & MyFiles, tvwChild
50650   End If
50660  End If
50670  Set reg = New clsRegistry
50680  reg.hkey = HKEY_CLASSES_ROOT
50690  reg.KeyRoot = "CLSID"
50700  reg.Subkey = "{00021400-0000-0000-C000-000000000046}"
50710  MyDesktopName = reg.GetRegistryValue("")
50720  Set reg = Nothing
50730  tvwFolder.Nodes.Add , , "S" & GetDesktop, MyDesktopName, 2
50740  MyDesktop = GetDesktop
50750  If CheckPath(MyFiles) = True Then
50760   If IsRemovableDrive(Mid$(MyDesktop, 1, 2)) = False Then
50770     If Dir(MyDesktop, vbDirectory + vbHidden) <> "" Then
50780      cDDF.CurrentDirectory = MyDesktop
50790      Set cDir = cDDF.GetDirectories
50800      If cDir.Count > 0 Then
50810       tvwFolder.Nodes.Add "S" & MyDesktop, tvwChild
50820      End If
50830     End If
50840    Else
50850     tvwFolder.Nodes.Add "S" & MyDesktop, tvwChild
50860   End If
50870  End If
50880
50890  For i = 1 To cDrives.Count
50900   drvS = Split(cDrives.item(i), Chr$(0))
50910   tvwFolder.Nodes.Add , , "N" & UCase$(drvS(0)), UCase$(Left$(drvS(0), 1)), GetDriveIcon(drvS(0))
50920   If IsRemovableDrive(Mid$(drvS(0), 1, 2) & "\") = True Then
50930     tvwFolder.Nodes.Add "N" & UCase$(drvS(0)), tvwChild
50940    Else
50950     If IsGoodDrive(Mid$(drvS(0), 1, 2)) = True Then
50960       If Dir(Mid$(drvS(0), 1, 2) & "\", vbDirectory + vbHidden) <> "" Then
50970        cDDF.CurrentDirectory = drvS(0)
50980        Set cDir = cDDF.GetDirectories
50990        If cDir.Count > 0 Then
51000         tvwFolder.Nodes.Add "N" & UCase$(drvS(0)), tvwChild
51010        End If
51020       End If
51030      Else
51040       tvwFolder.Nodes.Add "N" & UCase$(drvS(0)), tvwChild
51050     End If
51060   End If
51070  Next i
51080  tvwFolder.Nodes("S" & MyFiles).Selected = True
51090  With txtFilename
51100   .Text = Filename
51110   .SelStart = 0
51120   .SelLength = Len(.Text)
51130  End With
51140  cmbFiletypes.ListIndex = 0
51150  DoEvents
51160  SetLastSaveDirectory
51170  If SaveOpenType = OpenFile Then
51180   txtFilename.Locked = True
51190   Timer1.Enabled = True
51200  End If
51210  Screen.MousePointer = vbNormal
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSaveOpen", "Form_Load")
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
Select Case ErrPtnr.OnError("frmSaveOpen", "GetDriveIcon")
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
Select Case ErrPtnr.OnError("frmSaveOpen", "IsRemovableDrive")
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
50020   SaveOpenCancel = True
50030  End If
50040  Set LastNode = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSaveOpen", "Form_Unload")
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
Select Case ErrPtnr.OnError("frmSaveOpen", "lsvFilenames_Click")
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
50010  Dim tStr As String
50020  If InStr(txtFilename.Text, ":") > 0 Then
50030    tStr = txtFilename.Text
50040   Else
50050    tStr = CompletePath(cDDF.CurrentDirectory) & txtFilename.Text
50060  End If
50070  If FileExists(tStr) = True Then
50080    cmd(1).Enabled = True
50090   Else
50100    cmd(1).Enabled = False
50110    RefreshFilelist
50120  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSaveOpen", "Timer1_Timer")
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
Select Case ErrPtnr.OnError("frmSaveOpen", "tvwFolder_Expand")
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
Select Case ErrPtnr.OnError("frmSaveOpen", "SetNode")
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
Select Case ErrPtnr.OnError("frmSaveOpen", "tvwFolder_NodeClick")
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
50060   Set SaveOpenFilename = New Collection
50070   SaveOpenFilename.Add tSaveFilename
50080   SaveOpenFilterindex = cmbFiletypes.ListIndex
50090   Options.LastSaveDirectory = tSaveFilepath
50100   SaveOptions Options
50110   Unload Me
50120  End If
50130  If Len(txtFilename.Text) > 0 Then
50140    cmd(1).Enabled = True
50150   Else
50160    cmd(1).Enabled = False
50170  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSaveOpen", "txtFilename_KeyPress")
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
  tStr As String, Ext2 As String
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
50230  SplitPath sFilter(cmbFiletypes.ListIndex * 2 + 1), , Path, , SaveFilename, Ext2
50240  SplitPath txtFilename.Text, , Path, SaveFilename, , Ext
50250  If Len(Path) > 0 Then
50260    If DirExists(Path) = False Then
50270     If Left(Path, 2) = "\\" Then
50280      MsgBox LanguageStrings.MessagesMsg10, vbOKOnly
50290      CanSave = False
50300      Me.MousePointer = vbNormal
50310      Exit Function
50320     End If
50330     res = MsgBox(LanguageStrings.MessagesMsg09, vbYesNo)
50340     If res = vbNo Then
50350      CanSave = False
50360      Me.MousePointer = vbNormal
50370      Exit Function
50380     End If
50390     If MakePath(Path) = False Then
50400      MsgBox LanguageStrings.MessagesMsg07, vbOKOnly
50410      CanSave = False
50420      Me.MousePointer = vbNormal
50430      Exit Function
50440     End If
50450    End If
50460    SaveFilename = CompletePath(Path) & SaveFilename
50470    tSaveFilepath = CompletePath(Path)
50480   Else
50490    SaveFilename = CompletePath(cDDF.CurrentDirectory) & txtFilename.Text
50500    tSaveFilepath = CompletePath(cDDF.CurrentDirectory)
50510  End If
50520
50530  If UCase$(Ext) <> UCase$(Ext2) Then
50540   SaveFilename = SaveFilename & "." & Ext2
50550   txtFilename.Text = SaveFilename
50560  End If
50570
50580  If SaveOpenType = saveFile Then
50590   If Dir(SaveFilename) <> "" Then
50600    res = MsgBox(LanguageStrings.MessagesMsg05, vbYesNo)
50610    If res = vbNo Then
50620     CanSave = False
50630     Me.MousePointer = vbNormal
50640     Exit Function
50650    End If
50660   End If
50670   If Dir(SaveFilename) <> "" Then
50680    If GetAttr(SaveFilename) And vbReadOnly = vbReadOnly Then
50690     If GetAttr(SaveFilename) Or vbArchive = vbArchive Then
50700      res = vbArchive
50710     End If
50720     SetAttr SaveFilename, res
50730    End If
50740   End If
50750   res = vbYes
50760   If FileInUse(SaveFilename) Then
50770    MsgBox LanguageStrings.MessagesMsg34, vbExclamation
50780    CanSave = False
50790    Me.MousePointer = vbNormal
50800    Exit Function
50810   End If
50820  End If
50830  tSaveFilename = SaveFilename
50840  Me.MousePointer = vbNormal
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSaveOpen", "CanSave")
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
50530         tvwFolder_NodeClick tvwFolder.Nodes(nChild.Key)
50540         tvwFolder.Nodes(nChild.Key).Selected = True: DoEvents
50550         cDDF.CurrentDirectory = Mid(nChild.Key, 2)
50560         RefreshFilelist
50570         found = True: Exit Sub
50580        End If
50590        Set nChild = nChild.Next
50600       Loop
50610       If found = False Then
50620        Exit Sub
50630       End If
50640     End If
50650    End If
50660   End If
50670   Set nNode = nNode.Next: DoEvents
50680  Loop
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSaveOpen", "SetLastSaveDirectory")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function CutNull(tStr As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If InStr(tStr, Chr$(0)) > 0 Then
50020    CutNull = Mid(tStr, 1, InStr(tStr, Chr$(0)) - 1)
50030   Else
50040    CutNull = tStr
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSaveOpen", "CutNull")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
