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
 Dim Path As String, tFilename As String, Ext As String, Ext2 As String
 If InStr(Filter, "|") > 0 Then
  SplitPath cDDF.CurrentDirectory & "\" & sFilter(cmbFiletypes.ListIndex * 2 + 1), , Path, , tFilename, Ext2
  SplitPath cDDF.CurrentDirectory & "\" & txtFilename.Text, , Path, , tFilename, Ext
  cDDF.FilePattern = sFilter(cmbFiletypes.ListIndex * 2 + 1)
  If UCase$(Ext) <> UCase$(Ext2) Then
   If SaveOpenType = saveFile Then
    txtFilename.Text = tFilename & "." & Ext2
   End If
  End If
  RefreshFilelist
 End If
End Sub

Private Sub cmd_Click(Index As Integer)
 Select Case Index
  Case 1:
   If CanSave = False Then
    Exit Sub
   End If
   cmdCancel = False
   SaveOpenCancel = False
   Set SaveOpenFilename = New Collection
   SaveOpenFilename.Add tSaveFilename
   SaveOpenFilterindex = cmbFiletypes.ListIndex + 1
   Options.LastSaveDirectory = tSaveFilepath
   SaveOptions Options
 End Select
 Unload Me
End Sub

Private Sub RefreshFilelist()
 Dim i As Long, j As Long, item As ListItem, cFiles As Collection, _
  Ext As String

 lsvFilenames.ListItems.Clear
 Set cFiles = cDDF.GetFiles
 For i = 1 To cFiles.Count
  Set item = lsvFilenames.ListItems.Add(, , cFiles.item(i))
  SplitPath CutNull(cFiles.item(i)), , , , , Ext
  item.SubItems(1) = Format$(FileLen(CompletePath(cDDF.CurrentDirectory) & CutNull(cFiles.item(i))) / 1024, "#,##0") & " KB"
  item.SubItems(2) = GetFileAttributesStr(CompletePath(cDDF.CurrentDirectory) & cFiles.item(i))
  item.SmallIcon = 9
  For j = 1 To imlFilenameIcons.ListImages.Count
   If "*." & UCase$(Ext) = UCase$(imlFilenameIcons.ListImages(j).Tag) Then
    item.SmallIcon = j
    Exit For
   End If
  Next j
'  item.SmallIcon = cmbFiletypes.ListIndex + 1
 Next i
 If lsvFilenames.ListItems.Count > 0 And SaveOpenType = OpenFile Then
  txtFilename.Text = lsvFilenames.ListItems(1).Text
 End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF1 Then
  KeyCode = 0
  Call HTMLHelp_ShowTopic("html\welcome.htm")
 End If
End Sub

Private Sub Form_Load()
 Dim MyFilesName As String, MyDesktopName As String, MyDesktop As String, _
  MyFiles As String, reg As clsRegistry, res As Long, Icon As IPictureDisp, _
  cDrives As Collection, i As Long, j As Long, drvS() As String, _
  cDir As Collection, tStr As String
 Me.KeyPreview = True
 Screen.MousePointer = vbHourglass
 cmdCancel = True
 SaveOpenCancel = True
 Set cDDF = New clsDrivesDirsFiles
 Set cDrives = cDDF.GetDrives(True)
 With LanguageStrings
  cmd(0).Caption = .SaveOpenCancel
  If SaveOpenType = 0 Then
   Me.Caption = .SaveOpenSaveTitle
   cmd(1).Caption = .SaveOpenSave
  End If
  If SaveOpenType = 1 Then
   Me.Caption = .SaveOpenOpenTitle
   cmd(1).Caption = .SaveOpenOpen
  End If
 End With

 With lsvFilenames
  .View = lvwReport
 End With
 With lsvFilenames.ColumnHeaders
  .Clear
  .Add , "Filename", LanguageStrings.SaveOpenFilename, 3200
  .Add , "Size", LanguageStrings.SaveOpenSize, 1500, lvwColumnRight
  .Add , "Attributes", LanguageStrings.SaveOpenAttributes, 1000, lvwColumnLeft
 End With
 If InStr(Filter, "|") > 0 Then
  sFilter = Split(Filter, "|")
  Set lsvFilenames.SmallIcons = Nothing
  For i = LBound(sFilter) To UBound(sFilter) Step 2
   cmbFiletypes.AddItem sFilter(i)
   For j = 1 To imlFilenameIcons.ListImages.Count
    res = LoadIcon(Small, imlFilenameIcons.ListImages(j).Tag, Icon)
    If res = 0 Then
     tStr = imlFilenameIcons.ListImages(j).Tag
     imlFilenameIcons.ListImages.Remove j
     imlFilenameIcons.ListImages.Add j, , Icon
     imlFilenameIcons.ListImages(j).Tag = UCase$(tStr)
    End If
   Next j
  Next i
  Set lsvFilenames.SmallIcons = imlFilenameIcons
 End If
 DoEvents
 cDDF.FilePattern = sFilter(2 * Filterindex + 1)

 MyFilesName = GetShellNamespaceName(MYFILES_CLSID): MyFiles = GetMyFiles
 If CheckPath(MyFiles) = True Then
  tvwFolder.Nodes.Add , , "S" & MyFiles, MyFilesName, 8
  If IsRemovableDrive(Mid$(MyFiles, 1, 2)) = False Then
    If Dir(MyFiles, vbDirectory + vbHidden) <> "" Then
     cDDF.CurrentDirectory = MyFiles
     Set cDir = cDDF.GetDirectories
     If cDir.Count > 0 Then
      tvwFolder.Nodes.Add "S" & MyFiles, tvwChild
     End If
    End If
   Else
    tvwFolder.Nodes.Add "S" & MyFiles, tvwChild
  End If
 End If
 Set reg = New clsRegistry
 reg.hkey = HKEY_CLASSES_ROOT
 reg.KeyRoot = "CLSID"
 reg.Subkey = "{00021400-0000-0000-C000-000000000046}"
 MyDesktopName = reg.GetRegistryValue("")
 Set reg = Nothing
 tvwFolder.Nodes.Add , , "S" & GetDesktop, MyDesktopName, 2
 MyDesktop = GetDesktop
 If CheckPath(MyFiles) = True Then
  If IsRemovableDrive(Mid$(MyDesktop, 1, 2)) = False Then
    If Dir(MyDesktop, vbDirectory + vbHidden) <> "" Then
     cDDF.CurrentDirectory = MyDesktop
     Set cDir = cDDF.GetDirectories
     If cDir.Count > 0 Then
      tvwFolder.Nodes.Add "S" & MyDesktop, tvwChild
     End If
    End If
   Else
    tvwFolder.Nodes.Add "S" & MyDesktop, tvwChild
  End If
 End If

 For i = 1 To cDrives.Count
  drvS = Split(cDrives.item(i), Chr$(0))
  tvwFolder.Nodes.Add , , "N" & UCase$(drvS(0)), UCase$(Left$(drvS(0), 1)), GetDriveIcon(drvS(0))
  If IsRemovableDrive(Mid$(drvS(0), 1, 2) & "\") = True Then
    tvwFolder.Nodes.Add "N" & UCase$(drvS(0)), tvwChild
   Else
    If IsGoodDrive(Mid$(drvS(0), 1, 2)) = True Then
      If Dir(Mid$(drvS(0), 1, 2) & "\", vbDirectory + vbHidden) <> "" Then
       cDDF.CurrentDirectory = drvS(0)
       Set cDir = cDDF.GetDirectories
       If cDir.Count > 0 Then
        tvwFolder.Nodes.Add "N" & UCase$(drvS(0)), tvwChild
       End If
      End If
     Else
      tvwFolder.Nodes.Add "N" & UCase$(drvS(0)), tvwChild
    End If
  End If
 Next i
 tvwFolder.Nodes("S" & MyFiles).Selected = True
 With txtFilename
  .Text = Filename
  .SelStart = 0
  .SelLength = Len(.Text)
 End With
 cmbFiletypes.ListIndex = 0
 DoEvents
 SetLastSaveDirectory
 If SaveOpenType = OpenFile Then
  txtFilename.Locked = True
  Timer1.Enabled = True
 End If
 Screen.MousePointer = vbNormal
End Sub

Private Function GetDriveIcon(Drive As String) As Long
 Dim DriveType As Long
 GetDriveIcon = 6
 If Len(Trim$(Drive)) > 0 Then
  DriveType = GetDriveType(Mid$(Drive, 1, 2))
  Select Case DriveType
   Case 2: 'Removable
    GetDriveIcon = 3
   Case 3: 'Fixed
    GetDriveIcon = 6
   Case 4: 'Remote
    GetDriveIcon = 9
   Case 5: 'CDRom
    GetDriveIcon = 1
   Case 6: 'Ramdisk
    GetDriveIcon = 7
  End Select
 End If
End Function

Private Function IsRemovableDrive(Drive As String) As Boolean
 Dim DriveType As Long
 IsRemovableDrive = False
 DriveType = GetDriveType(Mid$(Drive, 1, 2))
 If DriveType = 2 Or DriveType = 4 Or DriveType = 5 Then
  IsRemovableDrive = True
 End If
End Function

Private Sub Form_Unload(Cancel As Integer)
 If cmdCancel = True Then
  SaveOpenCancel = True
 End If
 Set LastNode = Nothing
End Sub

Private Sub lsvFilenames_Click()
 If lsvFilenames.ListItems.Count = 0 Then
  Exit Sub
 End If
 With txtFilename
  .Text = lsvFilenames.SelectedItem.Text
  .SetFocus
  .SelStart = 0
  .SelLength = Len(.Text)
 End With
 DoEvents
End Sub

Private Sub Timer1_Timer()
 Dim tStr As String
 If InStr(txtFilename.Text, ":") > 0 Then
   tStr = txtFilename.Text
  Else
   tStr = CompletePath(cDDF.CurrentDirectory) & txtFilename.Text
 End If
 If FileExists(tStr) = True Then
   cmd(1).Enabled = True
  Else
   cmd(1).Enabled = False
   RefreshFilelist
 End If
End Sub

Private Sub tvwFolder_Expand(ByVal Node As MSComctlLib.Node)
 cmd(1).Enabled = False
 SetNode Node
 Call RefreshFilelist
 cmd(1).Enabled = True
End Sub

Private Sub SetNode(ByVal Node As MSComctlLib.Node)
' On Error Resume Next
 Dim strRelative As String, i As Long, intFolderPos As Long, strFolderName As String, _
  strNewPath As String, dirS() As String, cDir As Collection, cDir2 As Collection, _
  icInd As Long, icIndEx As Long
 MousePointer = vbHourglass
 If Node.Child.Text = "" Then
  tvwFolder.Nodes.Remove Node.Child.Index
  strRelative = Node.Key
  If CheckPath(Mid$(strRelative, 2)) = True Then
    cDDF.CurrentDirectory = Mid$(strRelative, 2)
    intFolderPos = Len(Mid$(strRelative, 2)) + 1
    Set cDir = cDDF.GetDirectories(, True)
    For i = 1 To cDir.Count
     dirS = Split(cDir.item(i), Chr$(0))
     strFolderName = dirS(0)
     strNewPath = CompletePath(CompletePath(Mid$(strRelative, 2)) & strFolderName)
     If (dirS(1) And vbHidden) > 0 Then
       icInd = 11: icIndEx = 12
      Else
       icInd = 4: icIndEx = 5
     End If
     tvwFolder.Nodes.Add strRelative, tvwChild, Mid$(strRelative, 1, 1) & strNewPath, strFolderName, icInd
     If CheckPath(strNewPath) = True Then
      cDDF.CurrentDirectory = strNewPath
      Set cDir2 = cDDF.GetDirectories(, True, True)
      If cDir2.Count > 0 Then
       tvwFolder.Nodes.Add Mid$(strRelative, 1, 1) & strNewPath, tvwChild, , ""
       tvwFolder.Nodes(Mid$(strRelative, 1, 1) & strNewPath).Image = icInd
       tvwFolder.Nodes(Mid$(strRelative, 1, 1) & strNewPath).ExpandedImage = icIndEx
      End If
      cDDF.CurrentDirectory = Mid$(strRelative, 2)
     End If
     DoEvents
    Next i
   Else
  End If
 End If
 MousePointer = vbDefault
End Sub

Private Sub tvwFolder_NodeClick(ByVal Node As MSComctlLib.Node)
 If LastNode Is Nothing = False Then
  If LastNode.Image = 5 Then
   LastNode.Image = 4
  End If
  If LastNode.ExpandedImage = 5 Then
   LastNode.ExpandedImage = 4
  End If
 End If
 Set LastNode = Node
 If CheckPath(Mid$(tvwFolder.SelectedItem.Key, 2)) = False Then
   cmd(1).Enabled = False
   With lsvFilenames
    .Font.Italic = True
    .ListItems.Clear
    .ColumnHeaders(1).Width = 5700
    .ListItems.Add , , LanguageStrings.MessagesMsg07
   End With
   Exit Sub
  Else
   cmd(1).Enabled = True
   With lsvFilenames
    .Font.Italic = False
    .ColumnHeaders(1).Width = 3200
   End With
 End If
 If Node.ExpandedImage = 4 Then
  Node.ExpandedImage = 5
 End If
 If Node.Image = 4 Then
  Node.Image = 5
 End If
 cDDF.CurrentDirectory = Mid$(tvwFolder.SelectedItem.Key, 2)
 RefreshFilelist
End Sub

Private Sub txtFilename_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  If CanSave = False Then
   Exit Sub
  End If
  cmdCancel = False
  Set SaveOpenFilename = New Collection
  SaveOpenFilename.Add tSaveFilename
  SaveOpenFilterindex = cmbFiletypes.ListIndex
  Options.LastSaveDirectory = tSaveFilepath
  SaveOptions Options
  Unload Me
 End If
 If Len(txtFilename.Text) > 0 Then
   cmd(1).Enabled = True
  Else
   cmd(1).Enabled = False
 End If
End Sub

Private Function CanSave() As Boolean
 Dim Path As String, Ext As String, SaveFilename As String, res As Long, _
  tStr As String, Ext2 As String
 Me.MousePointer = vbHourglass
 CanSave = True
 If InStr(txtFilename.Text, ":") > 0 Then
   tStr = txtFilename.Text
  Else
   tStr = CompletePath(cDDF.CurrentDirectory) & txtFilename.Text
 End If
 If IsValidPath(tStr, False) = False Then
   MsgBox LanguageStrings.MessagesMsg10, vbOKOnly
   CanSave = False
   Me.MousePointer = vbNormal
   Exit Function
  Else
   If IsGoodDrive(tStr) = False Then
    MsgBox LanguageStrings.MessagesMsg07, vbOKOnly
    CanSave = False
    Me.MousePointer = vbNormal
    Exit Function
   End If
 End If
 SplitPath sFilter(cmbFiletypes.ListIndex * 2 + 1), , Path, , SaveFilename, Ext2
 SplitPath txtFilename.Text, , Path, SaveFilename, , Ext
 If Len(Path) > 0 Then
   If DirExists(Path) = False Then
    If Left(Path, 2) = "\\" Then
     MsgBox LanguageStrings.MessagesMsg10, vbOKOnly
     CanSave = False
     Me.MousePointer = vbNormal
     Exit Function
    End If
    res = MsgBox(LanguageStrings.MessagesMsg09, vbYesNo)
    If res = vbNo Then
     CanSave = False
     Me.MousePointer = vbNormal
     Exit Function
    End If
    If MakePath(Path) = False Then
     MsgBox LanguageStrings.MessagesMsg07, vbOKOnly
     CanSave = False
     Me.MousePointer = vbNormal
     Exit Function
    End If
   End If
   SaveFilename = CompletePath(Path) & SaveFilename
   tSaveFilepath = CompletePath(Path)
  Else
   SaveFilename = CompletePath(cDDF.CurrentDirectory) & txtFilename.Text
   tSaveFilepath = CompletePath(cDDF.CurrentDirectory)
 End If

 If UCase$(Ext) <> UCase$(Ext2) Then
  SaveFilename = SaveFilename & "." & Ext2
  txtFilename.Text = SaveFilename
 End If

 If SaveOpenType = saveFile Then
  If Dir(SaveFilename) <> "" Then
   res = MsgBox(LanguageStrings.MessagesMsg05, vbYesNo)
   If res = vbNo Then
    CanSave = False
    Me.MousePointer = vbNormal
    Exit Function
   End If
  End If
  If Dir(SaveFilename) <> "" Then
   If GetAttr(SaveFilename) And vbReadOnly = vbReadOnly Then
    If GetAttr(SaveFilename) Or vbArchive = vbArchive Then
     res = vbArchive
    End If
    SetAttr SaveFilename, res
   End If
  End If
  res = vbYes
  If FileInUse(SaveFilename) Then
   MsgBox LanguageStrings.MessagesMsg34, vbExclamation
   CanSave = False
   Me.MousePointer = vbNormal
   Exit Function
  End If
 End If
 tSaveFilename = SaveFilename
 Me.MousePointer = vbNormal
End Function

Private Sub SetLastSaveDirectory()
 Dim LSDir As String, i As Long, nChild As Node, nNode As Node, tStr As String, _
  found As Boolean, Folder() As String, nNodeKey As String
 LSDir = UCase$(Options.LastSaveDirectory)
 If LSDir = "" Then
  LSDir = GetMyFiles
 End If
 If Right$(LSDir, 1) = "\" Then
  LSDir = Mid$(LSDir, 1, Len(LSDir) - 1)
 End If
 Set nNode = tvwFolder.Nodes(1)
 Do While Not (nNode.Parent Is Nothing)
  Set nNode = nNode.Parent: DoEvents
 Loop
 Set nNode = nNode.FirstSibling
 Do While Not (nNode Is Nothing)
  If Right$(nNode.Key, 1) = "\" Then
    nNodeKey = Mid$(nNode.Key, 2, Len(nNode.Key) - 2)
   Else
    nNodeKey = Mid$(nNode.Key, 2)
  End If
  If DirExists(LSDir) = True Then
   If InStr(LSDir, UCase$(nNodeKey)) > 0 Then
    nNode.Selected = True: DoEvents
    tStr = Mid$(LSDir, Len(nNodeKey) + 2)
    If Len(tStr) = 0 Then
     Exit Sub
    End If
    If InStr(tStr, "\") > 0 Then
      Folder = Split(tStr, "\")
      For i = 0 To UBound(Folder)
       Set nChild = tvwFolder.SelectedItem.Child
       Do While Not (nChild Is Nothing)
        If UCase$(nChild.Text) = UCase$(Folder(i)) Then
         tvwFolder.Nodes(nChild.Key).Selected = True
         If tvwFolder.Nodes(nChild.Key).Children = 0 Then
          tvwFolder_NodeClick tvwFolder.Nodes(nChild.Key)
         End If
         DoEvents
         found = True: Exit Do
        End If
        Set nChild = nChild.Next
       Loop
       If found = False Then
        Exit Sub
       End If
      Next i
      Exit Sub
     Else
      Set nChild = nNode.Child
      found = False
      Do While Not (nChild Is Nothing)
       If UCase$(nChild.Text) = UCase$(tStr) Then
        tvwFolder_NodeClick tvwFolder.Nodes(nChild.Key)
        tvwFolder.Nodes(nChild.Key).Selected = True: DoEvents
        cDDF.CurrentDirectory = Mid(nChild.Key, 2)
        RefreshFilelist
        found = True: Exit Sub
       End If
       Set nChild = nChild.Next
      Loop
      If found = False Then
       Exit Sub
      End If
    End If
   End If
  End If
  Set nNode = nNode.Next: DoEvents
 Loop
End Sub

Private Function CutNull(tStr As String) As String
 If InStr(tStr, Chr$(0)) > 0 Then
   CutNull = Mid(tStr, 1, InStr(tStr, Chr$(0)) - 1)
  Else
   CutNull = tStr
 End If
End Function
