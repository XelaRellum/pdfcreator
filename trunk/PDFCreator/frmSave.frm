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
 Dim Path As String, FileName As String, Ext As String
 SplitPath Dir1.Path & "\" & txtFilename.Text, , Path, , FileName, Ext
 Select Case cmbFiletypes.ListIndex
  Case 0:
   File1.Pattern = "*.pdf"
   If UCase$(Ext) <> "PDF" Then
    txtFilename.Text = FileName & ".pdf"
   End If
  Case 1:
   File1.Pattern = "*.png"
   If UCase$(Ext) <> "PNG" Then
    txtFilename.Text = FileName & ".png"
   End If
  Case 2:
   File1.Pattern = "*.jpg"
   If UCase$(Ext) <> "JPG" Then
    txtFilename.Text = FileName & ".jpg"
   End If
  Case 3:
   File1.Pattern = "*.bmp"
   If UCase$(Ext) <> "BMP" Then
    txtFilename.Text = FileName & ".bmp"
   End If
  Case 4:
   File1.Pattern = "*.pcx"
   If UCase$(Ext) <> "PCX" Then
    txtFilename.Text = FileName & ".pcx"
   End If
  Case 5:
   File1.Pattern = "*.tif"
   If UCase$(Ext) <> "TIF" Then
    txtFilename.Text = FileName & ".tif"
   End If
  Case 6:
   File1.Pattern = "*.ps"
   If UCase$(Ext) <> "PS" Then
    txtFilename.Text = FileName & ".ps"
   End If
  Case 7:
   File1.Pattern = "*.eps"
   If UCase$(Ext) <> "EPS" Then
    txtFilename.Text = FileName & ".eps"
   End If
 End Select
 RefreshFilelist
End Sub

Private Sub cmd_Click(Index As Integer)
 Select Case Index
  Case 1:
   If CanSave = False Then
    Exit Sub
   End If
   cmdCancel = False
   frmPrinting.SaveFilename = tSaveFilename
   frmPrinting.SaveFilterIndex = cmbFiletypes.ListIndex
   Options.LastSaveDirectory = tSaveFilepath
   SaveOptions Options
 End Select
 Unload Me
End Sub

Private Sub RefreshFilelist()
 On Local Error Resume Next
 Dim i As Long, item As ListItem
 lsvFilenames.ListItems.Clear
 For i = 0 To File1.ListCount - 1
  Set item = lsvFilenames.ListItems.Add(, , File1.List(i))
  item.SubItems(1) = Format$(FileLen(File1.Path & "\" & File1.List(i)) / 1024, "#,##0") & " KB"
  item.SubItems(2) = GetFileAttributesStr(File1.Path & "\" & File1.List(i))
  item.SmallIcon = cmbFiletypes.ListIndex + 1
 Next i
End Sub

Private Sub Dir1_Change()
 File1.Path = Dir1.Path
 Call RefreshFilelist
End Sub

Private Sub Form_Load()
 Dim i As Long
 Dim MyFilesName As String, MyDesktopName As String, MyDesktop As String, _
  MyFiles As String, reg As clsRegistry, res As Long, Icon As IPictureDisp, _
  cDrives As Collection
 Screen.MousePointer = vbHourglass
 cmdCancel = True
 Set cDrives = GetDrives
 With LanguageStrings
  Me.Caption = .SaveTitle
  cmd(0).Caption = .SaveCancel
  cmd(1).Caption = .SaveSave
 End With

 With lsvFilenames
  .View = lvwReport
 End With
 With lsvFilenames.ColumnHeaders
  .Clear
  .Add , "Filename", LanguageStrings.SaveFilename, 3200
  .Add , "Size", LanguageStrings.SaveSize, 1500, lvwColumnRight
  .Add , "Attributes", LanguageStrings.SaveAttributes, 1000, lvwColumnLeft
 End With
 With LanguageStrings
  cmbFiletypes.AddItem .ListPDFFiles & " (*.pdf)"
  cmbFiletypes.AddItem .PrintingPNGFiles & " (*.png)"
  cmbFiletypes.AddItem .PrintingJPEGFiles & " (*.jpg)"
  cmbFiletypes.AddItem .PrintingBMPFiles & " (*.bmp)"
  cmbFiletypes.AddItem .PrintingPCXFiles & " (*.pcx)"
  cmbFiletypes.AddItem .PrintingTIFFFiles & " (*.tif)"
  cmbFiletypes.AddItem .PrintingPSFiles & " (*.ps)"
  cmbFiletypes.AddItem .PrintingEPSFiles & " (*.eps)"
 End With
 Set lsvFilenames.SmallIcons = Nothing
 With imlFilenameIcons.ListImages
  res = LoadIcon(Small, "*.pdf", Icon)
  If res = 0 Then
   .Remove 1
   .Add 1, , Icon
  End If
  res = LoadIcon(Small, "*.png", Icon)
  If res = 0 Then
   .Remove 2
   .Add 2, , Icon
  End If
  res = LoadIcon(Small, "*.jpg", Icon)
  If res = 0 Then
   .Remove 3
   .Add 3, , Icon
  End If
  res = LoadIcon(Small, "*.bmp", Icon)
  If res = 0 Then
   .Remove 4
   .Add 4, , Icon
  End If
  res = LoadIcon(Small, "*.pcx", Icon)
  If res = 0 Then
   .Remove 5
   .Add 5, , Icon
  End If
  res = LoadIcon(Small, "*.tif", Icon)
  If res = 0 Then
   .Remove 6
   .Add 6, , Icon
  End If
  res = LoadIcon(Small, "*.ps", Icon)
  If res = 0 Then
   .Remove 7
   .Add 7, , Icon
  End If
  res = LoadIcon(Small, "*.eps", Icon)
  If res = 0 Then
   .Remove 7
   .Add 7, , Icon
  End If
 End With
 Set lsvFilenames.SmallIcons = imlFilenameIcons
 DoEvents
 File1.Pattern = "*.pdf"

 MyFilesName = GetShellNamespaceName(MYFILES_CLSID): MyFiles = GetMyFiles
 If CheckPath(MyFiles) = True Then
  tvwFolder.Nodes.Add , , "S" & MyFiles, MyFilesName, 8
  If IsRemovableDrive(Mid$(MyFiles, 1, 2)) = False Then
    If Dir(MyFiles, vbDirectory) <> "" Then
     Dir1.Path = MyFiles
     If Dir1.ListCount > 0 Then
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
    If Dir(MyDesktop, vbDirectory) <> "" Then
     Dir1.Path = MyDesktop
     If Dir1.ListCount > 0 Then
      tvwFolder.Nodes.Add "S" & MyDesktop, tvwChild
     End If
    End If
   Else
    tvwFolder.Nodes.Add "S" & MyDesktop, tvwChild
  End If
 End If

 For i = 1 To cDrives.Count
  tvwFolder.Nodes.Add , , "N" & UCase$(cDrives.item(i)), UCase$(Left$(cDrives.item(i), 1)), GetDriveIcon(cDrives.item(i))
'51170   If IsRemovableDrive(Mid$(cDrives.item(i), 1, 2)) = False Then
  If IsRemovableDrive(Mid$(cDrives.item(i), 1, 2) & "\") = True Then
    tvwFolder.Nodes.Add "N" & UCase$(cDrives.item(i)), tvwChild
   Else
    If IsGoodDrive(Mid$(cDrives.item(i), 1, 2)) = True Then
      If Dir(Mid$(cDrives.item(i), 1, 2) & "\", vbDirectory) <> "" Then
       Dir1.Path = cDrives.item(i)
       If Dir1.ListCount > 0 Then
        tvwFolder.Nodes.Add "N" & UCase$(cDrives.item(i)), tvwChild
       End If
      End If
     Else
      tvwFolder.Nodes.Add "N" & UCase$(cDrives.item(i)), tvwChild
    End If
  End If
 Next i
 tvwFolder.Nodes("S" & MyFiles).Selected = True
 With txtFilename
  .Text = frmPrinting.SaveFilename
  .SelStart = 0
  .SelLength = Len(.Text)
 End With
 cmbFiletypes.ListIndex = 0
 DoEvents
 SetLastSaveDirectory
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
  frmPrinting.SaveCancel = True
 End If
 Set LastNode = Nothing
End Sub

Private Sub lsvFilenames_Click()
 With txtFilename
  .Text = lsvFilenames.SelectedItem.Text
  .SetFocus
  .SelStart = 0
  .SelLength = Len(.Text)
 End With
 DoEvents
End Sub

Private Sub tvwFolder_Expand(ByVal Node As MSComctlLib.Node)
 SetNode Node
End Sub

Private Sub SetNode(ByVal Node As MSComctlLib.Node)
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
 Dir1.Path = Mid$(tvwFolder.SelectedItem.Key, 2)
End Sub

Private Sub txtFilename_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  If CanSave = False Then
   Exit Sub
  End If
  cmdCancel = False
  frmPrinting.SaveFilename = tSaveFilename
  frmPrinting.SaveFilterIndex = cmbFiletypes.ListIndex
  Options.LastSaveDirectory = tSaveFilepath
  SaveOptions Options
  Unload Me
 End If
End Sub

Private Function CanSave() As Boolean
 Dim Path As String, Ext As String, SaveFilename As String, res As Long, _
  tStr As String
 Me.MousePointer = vbHourglass
 CanSave = True
 If InStr(txtFilename.Text, ":") > 0 Then
   tStr = txtFilename.Text
  Else
   tStr = CompletePath(Dir1.Path) & txtFilename.Text
 End If
 If IsValidPath(tStr, False) = False Then
   MsgBox LanguageStrings.MessagesMsg10, vbOKOnly
   CanSave = False
   Exit Function
  Else
   If IsGoodDrive(tStr) = False Then
    MsgBox LanguageStrings.MessagesMsg07, vbOKOnly
    CanSave = False
    Exit Function
   End If
 End If
 SplitPath txtFilename.Text, , Path, SaveFilename, , Ext
 If Len(Path) > 0 Then
   If Len(Dir(Path, vbDirectory)) = 0 Then
    res = MsgBox(LanguageStrings.MessagesMsg09, vbYesNo)
    If res = vbNo Then
     CanSave = False
     Exit Function
    End If
    MakePath Path
   End If
   SaveFilename = CompletePath(Path) & SaveFilename
   tSaveFilepath = CompletePath(Path)
  Else
   SaveFilename = CompletePath(Dir1.Path) & txtFilename.Text
   tSaveFilepath = CompletePath(Dir1.Path)
 End If
 Select Case cmbFiletypes.ListIndex
  Case 0:
   If UCase$(Ext) <> "PDF" Then
    SaveFilename = SaveFilename & ".pdf"
    txtFilename.Text = txtFilename.Text & ".pdf"
   End If
  Case 1:
   If UCase$(Ext) <> "PNG" Then
    SaveFilename = SaveFilename & ".png"
    txtFilename.Text = txtFilename.Text & ".png"
   End If
  Case 2:
   If UCase$(Ext) <> "JPG" Then
    SaveFilename = SaveFilename & ".jpg"
    txtFilename.Text = txtFilename.Text & ".jpg"
   End If
  Case 3:
   If UCase$(Ext) <> "BMP" Then
    SaveFilename = SaveFilename & ".bmp"
    txtFilename.Text = txtFilename.Text & ".bmp"
   End If
  Case 4:
   If UCase$(Ext) <> "PCX" Then
    SaveFilename = SaveFilename & ".pcx"
    txtFilename.Text = txtFilename.Text & ".pcx"
   End If
  Case 5:
   If UCase$(Ext) <> "TIF" Then
    SaveFilename = SaveFilename & ".tif"
    txtFilename.Text = txtFilename.Text & ".tif"
   End If
  Case 6:
   If UCase$(Ext) <> "PS" Then
    SaveFilename = SaveFilename & ".ps"
    txtFilename.Text = txtFilename.Text & ".ps"
   End If
  Case 7:
   If UCase$(Ext) <> "EPS" Then
    SaveFilename = SaveFilename & ".eps"
    txtFilename.Text = txtFilename.Text & ".eps"
   End If
 End Select
 If Dir(SaveFilename) <> "" Then
  res = MsgBox(LanguageStrings.MessagesMsg05, vbYesNo)
  If res = vbNo Then
   CanSave = False
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
  If Dir(LSDir, vbDirectory) <> "" Then
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
        tvwFolder.Nodes(nChild.Key).Selected = True: DoEvents
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
