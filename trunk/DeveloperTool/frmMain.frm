VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "PDFCreator Developer Tools"
   ClientHeight    =   6390
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12930
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   12930
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame fra 
      Caption         =   "Inc file"
      Height          =   1755
      Index           =   4
      Left            =   3570
      TabIndex        =   36
      Top             =   2415
      Width           =   3315
      Begin VB.CommandButton cmdIncFile 
         Caption         =   "Convert direct"
         Height          =   495
         Index           =   2
         Left            =   1680
         TabIndex        =   40
         Top             =   735
         Width           =   1575
      End
      Begin VB.CommandButton cmdIncFile 
         Caption         =   "Save inc file"
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   105
         TabIndex        =   38
         Top             =   1050
         Width           =   1575
      End
      Begin VB.CommandButton cmdIncFile 
         Caption         =   "Load language file"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   37
         Top             =   735
         Width           =   1575
      End
      Begin VB.TextBox txtIncFile 
         BackColor       =   &H00C0FFFF&
         Height          =   495
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Beides
         TabIndex        =   39
         Top             =   240
         Width           =   3105
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Stamppage"
      Height          =   1860
      Index           =   3
      Left            =   210
      TabIndex        =   30
      Top             =   2415
      Width           =   3315
      Begin VB.CommandButton cmdStamppage 
         Caption         =   "Show with GSView"
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   105
         TabIndex        =   32
         Top             =   1365
         Width           =   1575
      End
      Begin VB.CommandButton cmdStamppage 
         Caption         =   "Save Stamppage"
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   105
         TabIndex        =   33
         Top             =   1050
         Width           =   1575
      End
      Begin VB.TextBox txtStamppage 
         BackColor       =   &H00C0FFFF&
         Height          =   495
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Beides
         TabIndex        =   35
         Top             =   240
         Width           =   3105
      End
      Begin VB.CommandButton cmdStamppage 
         Caption         =   "Load Stamppage"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   735
         Width           =   1575
      End
      Begin VB.CommandButton cmdStamppage 
         Caption         =   "Copy Stamppage to Clipboard"
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   1680
         TabIndex        =   31
         Top             =   735
         Width           =   1575
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Printer Registry Data Settings"
      Height          =   1755
      Index           =   5
      Left            =   6930
      TabIndex        =   24
      Top             =   2415
      Width           =   3360
      Begin VB.CommandButton cmdPrintRegData 
         Caption         =   "Save WinNT Setup-Includefile"
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   1680
         TabIndex        =   28
         Top             =   1155
         Width           =   1575
      End
      Begin VB.CommandButton cmdPrintRegData 
         Caption         =   "Load WinNT Regfile and convert"
         Height          =   495
         Index           =   2
         Left            =   1680
         TabIndex        =   29
         Top             =   735
         Width           =   1575
      End
      Begin VB.CommandButton cmdPrintRegData 
         Caption         =   "Save Win9x Setup-Includefile"
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   1155
         Width           =   1575
      End
      Begin VB.CommandButton cmdPrintRegData 
         Caption         =   "Load Win9x Regfile and convert"
         Height          =   495
         Index           =   0
         Left            =   105
         TabIndex        =   26
         Top             =   735
         Width           =   1575
      End
      Begin VB.TextBox txtPrintRegData 
         BackColor       =   &H00C0FFFF&
         Height          =   495
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Beides
         TabIndex        =   25
         Top             =   240
         Width           =   3105
      End
   End
   Begin MSComctlLib.StatusBar stb 
      Align           =   2  'Unten ausrichten
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   6135
      Width           =   12930
      _ExtentX        =   22807
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   210
      Top             =   5775
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fra 
      Caption         =   "Testpage"
      Height          =   1965
      Index           =   2
      Left            =   6930
      TabIndex        =   7
      Top             =   420
      Width           =   3315
      Begin VB.CommandButton cmdTestpage 
         Caption         =   "Create modTestpage.bas"
         Enabled         =   0   'False
         Height          =   495
         Index           =   4
         Left            =   1680
         TabIndex        =   41
         Top             =   1050
         Width           =   1575
      End
      Begin VB.CommandButton cmdTestpage 
         Caption         =   "Copy Testpage to Clipboard"
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   1680
         TabIndex        =   14
         Top             =   630
         Width           =   1575
      End
      Begin VB.CommandButton cmdTestpage 
         Caption         =   "Show with GSView"
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   105
         TabIndex        =   13
         Top             =   1470
         Width           =   1575
      End
      Begin VB.CommandButton cmdTestpage 
         Caption         =   "Save Testpage"
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   1050
         Width           =   1575
      End
      Begin VB.CommandButton cmdTestpage 
         Caption         =   "Load Testpage"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   615
         Width           =   1575
      End
      Begin VB.TextBox txtTestpage 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Beides
         TabIndex        =   10
         Top             =   240
         Width           =   3105
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Languages"
      Height          =   1965
      Index           =   1
      Left            =   3570
      TabIndex        =   6
      Top             =   420
      Width           =   3285
      Begin VB.CommandButton cmdLanguages 
         Caption         =   "Load"
         Height          =   495
         Index           =   4
         Left            =   1695
         TabIndex        =   20
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton cmdLanguages 
         Caption         =   "Create modLanguages.bas"
         Enabled         =   0   'False
         Height          =   495
         Index           =   5
         Left            =   1695
         TabIndex        =   22
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdLanguages 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton cmdLanguages 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   1695
         TabIndex        =   18
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton cmdLanguages 
         Caption         =   "Edit"
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdLanguages 
         Caption         =   "Add"
         Height          =   495
         Index           =   0
         Left            =   105
         TabIndex        =   16
         Top             =   735
         Width           =   1575
      End
      Begin MSComctlLib.ListView lsvLanguages 
         Height          =   450
         Left            =   105
         TabIndex        =   15
         Top             =   240
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   794
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Options"
      Height          =   1950
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   420
      Width           =   3285
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Load"
         Height          =   495
         Index           =   4
         Left            =   1635
         TabIndex        =   9
         Top             =   1410
         Width           =   1575
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Create modOptions.bas"
         Enabled         =   0   'False
         Height          =   495
         Index           =   5
         Left            =   1635
         TabIndex        =   21
         Top             =   1050
         Width           =   1575
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   60
         TabIndex        =   8
         Top             =   1410
         Width           =   1575
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   1635
         TabIndex        =   5
         Top             =   690
         Width           =   1575
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Edit"
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   60
         TabIndex        =   4
         Top             =   1050
         Width           =   1575
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Add"
         Height          =   495
         Index           =   0
         Left            =   60
         TabIndex        =   3
         Top             =   690
         Width           =   1575
      End
      Begin MSComctlLib.ListView lsvOptions 
         Height          =   465
         Left            =   30
         TabIndex        =   2
         Top             =   225
         Width           =   3150
         _ExtentX        =   5556
         _ExtentY        =   820
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin MSComctlLib.TabStrip tbstr 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   9975
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Options"
            Key             =   "Options"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Languages"
            Key             =   "Languages"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Testpage"
            Key             =   "Testpage"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Stamppage"
            Key             =   "Stamppage"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Create inc files for setup"
            Key             =   "inc"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Printer Registry Data Settings"
            Key             =   "WinPrintRegData"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnFile 
      Caption         =   "&File"
      Begin VB.Menu mnExit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum vType
 ShortString = 0
 LongString = 1
End Enum

Private Enum eOSTyp
 Win9x = 0
 WinNt = 1
End Enum

Private EditItem As Boolean, TempFile1 As String, _
 ChangeOptions As Boolean, ChangeLanguages As Boolean, ChangeTestpage As Boolean, _
 ChangeStamppage As Boolean, LastIncFile As String

Public Function AddLanguagesItem(Str1 As String, Str2 As String, Str3 As String, Str4 As String) As Boolean
 Dim Item As ListItem, i As Long
 AddLanguagesItem = True

 If EditItem = False Then
   For i = 1 To lsvLanguages.ListItems.Count
    If UCase$(Str2) = UCase$(lsvLanguages.ListItems(i).SubItems(1)) And UCase$(Str1) = UCase$(lsvLanguages.ListItems(i).Text) Then
     MsgBox "The key '" & Str2 & "' already exists in the section '" & Str1 & "'!", vbExclamation
     AddLanguagesItem = False
     Exit Function
    End If
   Next i
   For i = lsvLanguages.ListItems.Count To 1 Step -1
    If UCase$(Str1) = UCase$(lsvLanguages.ListItems(i).Text) Then
     Set Item = lsvLanguages.ListItems.Add(i + 1, , Str1)
     Exit For
    End If
   Next i
   If i = 0 Then
    Set Item = lsvLanguages.ListItems.Add(, , Str1)
   End If
  Else
   Set Item = lsvLanguages.SelectedItem
   Item.Text = Str1
 End If
 Item.SubItems(1) = Str2
 Item.SubItems(2) = Str3
 Item.SubItems(3) = Str4
 ChangeLanguages = True
End Function

Public Function AddOptionsItem(Str1 As String, Str2 As String, Str3 As String, Str4 As String, Str5 As String, Str6 As String, Comment As String) As Boolean
 Dim Item As ListItem, i As Long
 AddOptionsItem = True
 If EditItem = False Then
   For i = 1 To lsvOptions.ListItems.Count
    If UCase$(Str1) = UCase$(lsvOptions.ListItems(i).SubItems(1)) Then
     MsgBox "The name of the option '" & Str1 & "' already exists!", vbExclamation
     AddOptionsItem = False
     Exit Function
    End If
   Next i
   For i = 1 To lsvOptions.ListItems.Count
    If Len(Str2) > 0 And UCase$(Str2) = UCase$(lsvOptions.ListItems(i).SubItems(2)) Then
     MsgBox "The objectname of the option '" & Str2 & "' already exists!", vbExclamation
     AddOptionsItem = False
     Exit Function
    End If
   Next i
   For i = lsvOptions.ListItems.Count To 1 Step -1
    If UCase$(Comment) = UCase$(lsvOptions.ListItems(i).Text) Then
     Set Item = lsvOptions.ListItems.Add(i + 1, , Comment)
     Exit For
    End If
   Next i
   If i = 0 Then
    Set Item = lsvOptions.ListItems.Add(, , Comment)
   End If
  Else
   Set Item = lsvOptions.SelectedItem
   Item.Text = Comment
 End If
 Item.SubItems(1) = Str1
 Item.SubItems(2) = Str2
 Item.SubItems(3) = Str3
 Item.SubItems(4) = Str4
 Item.SubItems(5) = Str5
 Item.SubItems(6) = Str6
 ChangeOptions = True
End Function

Private Function CheckType(TypeStr As String) As Boolean
 CheckType = False
 If UCase$(TypeStr) = "BOOLEAN" Then
  CheckType = True
  Exit Function
 End If
 If UCase$(TypeStr) = "BYTE" Then
  CheckType = True
  Exit Function
 End If
 If UCase$(TypeStr) = "LONG" Then
  CheckType = True
  Exit Function
 End If
 If UCase$(TypeStr) = "STRING" Then
  CheckType = True
  Exit Function
 End If
 If UCase$(TypeStr) = "DOUBLE" Then
  CheckType = True
  Exit Function
 End If
End Function

Private Sub cmdIncFile_Click(Index As Integer)
 Dim fn As Long, GSViewPath As String
 
 On Error GoTo ErrorHandler
 Select Case Index
  Case 0 ' Load
   With cdlg
    .CancelError = True
    .Filename = ""
    .Filter = "Languages files (*.ini)|*.ini"
    .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist
    .InitDir = App.Path & "\..\PDFCreator\Languages"
    .ShowOpen
    txtIncFile.Text = GetKeysAndValuesFromInifile("Setup", .Filename)
    If Len(txtIncFile.Text) > 0 Then
     cmdIncFile(1).Enabled = True
    End If
   End With
  Case 1 ' Save
   With cdlg
    .Filename = LastIncFile & ".inc"
    .Filter = "Setup inc files (*.inc)|*.inc"
    .Flags = cdlOFNPathMustExist & cdlOFNOverwritePrompt Or cdlOFNLongNames
    .InitDir = App.Path & "\..\Setup"
    .ShowSave
    SaveFile .Filename, txtIncFile.Text
   End With
  Case 2 ' Convert direct
   With cdlg
    .CancelError = True
    .Filename = ""
    .Filter = "Languages files (*.ini)|*.ini"
    .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist
    .InitDir = App.Path & "\..\PDFCreator\Languages"
    .ShowOpen
    txtIncFile.Text = GetKeysAndValuesFromInifile("Setup", .Filename)
    If Len(txtIncFile.Text) > 0 Then
     cmdIncFile(1).Enabled = True
    End If
   End With
   SaveFile App.Path & "\..\Setup\" & LastIncFile & ".inc", txtIncFile.Text
 End Select
 Exit Sub
ErrorHandler:
 If Err.Number = 32755 Then
  Exit Sub
 End If
 MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub cmdLanguages_Click(Index As Integer)
 Dim aw As Long, tStr As String, i As Long
 Select Case Index
  Case 0: 'Add
   EditItem = False
   With frmLanguage
    If lsvLanguages.SelectedItem Is Nothing Then
      .cmbSection.Text = ""
     Else
      .cmbSection.Text = lsvLanguages.SelectedItem.Text
    End If
    If lsvLanguages.ListItems.Count > 0 Then
     tStr = UCase$(lsvLanguages.ListItems(1).Text)
     .cmbSection.AddItem lsvLanguages.ListItems(1).Text
     For i = 2 To lsvLanguages.ListItems.Count
      If UCase$(lsvLanguages.ListItems(i).Text) <> tStr Then
       tStr = UCase$(lsvLanguages.ListItems(i).Text)
       .cmbSection.AddItem lsvLanguages.ListItems(i).Text
      End If
     Next i
    End If
    .Show vbModal, Me
   End With
  Case 1: 'Edit
   EditItem = True
   ShowLanguage
  Case 2: 'Delete
   aw = MsgBox("Delete this entry?", vbQuestion Or vbYesNo)
   If aw = vbYes Then
    lsvLanguages.ListItems.Remove lsvLanguages.SelectedItem.Index
   End If
  Case 3: 'Save
   cmdLanguages(3).Enabled = False
   Screen.MousePointer = vbHourglass
   SaveTemplate
   ChangeLanguages = False
   Screen.MousePointer = vbNormal
   cmdLanguages(3).Enabled = True
  Case 4: 'Load
   cmdLanguages(4).Enabled = False
   Screen.MousePointer = vbHourglass
   lsvLanguages.Enabled = False
   ReadTemplate
   lsvLanguages.Enabled = True
   Screen.MousePointer = vbNormal
   cmdLanguages(4).Enabled = True
  Case 5: 'Create
   If ChangeLanguages = True Then
    MsgBox "You have change the languages settings! Please save the languages settings first?" & vbCrLf & "modLanguages.bas is not createt!", vbInformation
    Exit Sub
   End If
   CreateModLanguages
 End Select
 If lsvLanguages.ListItems.Count = 0 Then
   cmdLanguages(1).Enabled = False
   cmdLanguages(2).Enabled = False
   cmdLanguages(3).Enabled = False
   cmdLanguages(5).Enabled = False
  Else
   cmdLanguages(1).Enabled = True
   cmdLanguages(2).Enabled = True
   cmdLanguages(3).Enabled = True
   cmdLanguages(5).Enabled = True
 End If
 Screen.MousePointer = vbHourglass
 stb.Panels("Count").Text = lsvLanguages.ListItems.Count & " Entries"
 If stb.Panels.Count >= 3 Then
  stb.Panels("EngCount").Text = EngCount & " english entries"
  stb.Panels("GerCount").Text = GerCount & " german entries"
 End If
 Screen.MousePointer = vbNormal
End Sub

Private Sub cmdOptions_Click(Index As Integer)
 Dim aw As Long, tStr As String, i As Long
 
 On Error GoTo ErrorHandler
 
 Select Case Index
  Case 0: ' Add
   EditItem = False
   With frmOption
    If lsvOptions.SelectedItem Is Nothing Then
      .cmbComment.Text = ""
     Else
      .cmbComment.Text = lsvOptions.SelectedItem.Text
    End If
    If lsvOptions.ListItems.Count > 0 Then
     tStr = UCase$(lsvOptions.ListItems(1).Text)
     .cmbComment.AddItem lsvOptions.ListItems(1).Text
     For i = 2 To lsvOptions.ListItems.Count
      If UCase$(lsvOptions.ListItems(i).Text) <> tStr Then
       tStr = UCase$(lsvOptions.ListItems(i).Text)
       .cmbComment.AddItem lsvOptions.ListItems(i).Text
      End If
     Next i
    End If
    .Show vbModal, Me
   End With
   If stb.Panels.Count >= 3 Then
    stb.Panels("EngCount").Text = EngCount & " english entries"
    stb.Panels("GerCount").Text = GerCount & " german entries"
   End If
  Case 1: 'Edit
   EditItem = True
   ShowOption
   If stb.Panels.Count >= 3 Then
    stb.Panels("EngCount").Text = EngCount & " english entries"
    stb.Panels("GerCount").Text = GerCount & " german entries"
   End If
  Case 2: 'Delete
   aw = MsgBox("Delete this option?", vbQuestion Or vbYesNo)
   If aw = vbYes Then
    lsvOptions.ListItems.Remove lsvOptions.SelectedItem.Index
    ChangeOptions = True
   End If
   If stb.Panels.Count >= 3 Then
    stb.Panels("EngCount").Text = EngCount & " english entries"
    stb.Panels("GerCount").Text = GerCount & " german entries"
   End If
  Case 3: 'Save
   With cdlg
    .Filename = "Options.txt"
    .Filter = "(*.txt)|*.txt"
    .Flags = cdlOFNPathMustExist & cdlOFNOverwritePrompt Or cdlOFNLongNames
    .InitDir = App.Path & "\Options"
    .ShowSave
    SaveOptions .Filename
    ChangeOptions = False
   End With
  Case 4: 'Load
   lsvOptions.Enabled = False
   With cdlg
    .CancelError = True
    .Filename = "Options.txt"
    .Filter = "(*.txt)|*.txt"
    .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist
    .InitDir = App.Path & "\Options"
    .ShowOpen
    LoadOptions .Filename
   End With
   lsvOptions.Enabled = True
  Case 5: 'Create
   lsvOptions.Enabled = False
   CreateModOptions
   lsvOptions.Enabled = True
 End Select
 If lsvOptions.ListItems.Count = 0 Then
   cmdOptions(1).Enabled = False
   cmdOptions(2).Enabled = False
   cmdOptions(3).Enabled = False
   cmdOptions(5).Enabled = False
  Else
   cmdOptions(1).Enabled = True
   cmdOptions(2).Enabled = True
   cmdOptions(3).Enabled = True
   cmdOptions(5).Enabled = True
 End If
 stb.Panels("Count").Text = lsvOptions.ListItems.Count & " Options"
 
 With lsvOptions
  .SortOrder = lvwAscending
  .SortKey = 2
  .Sorted = True
  .SortKey = 1
  .Sorted = True
  .Sorted = False
 End With
 Exit Sub
ErrorHandler:
 lsvOptions.Enabled = True
 If Err.Number <> 32755 Then
  MsgBox Err.Number & ": " & Err.Description
 End If
End Sub

Private Sub cmdStamppage_Click(Index As Integer)
 Dim fn As Long, GSViewPath As String
 
 On Error GoTo ErrorHandler
 Select Case Index
  Case 0 ' Load
   With cdlg
    .CancelError = True
    .Filename = "Stamppage.ps"
    .Filter = "Postscript Files (*.ps)|*.ps|(*.txt)|*.txt"
    .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist
    .InitDir = App.Path & "\Stamppage"
    .ShowOpen
    LoadFileInTextbox .Filename, txtStamppage
    cmdStamppage(1).Enabled = True
   End With
  Case 1 ' Save
   With cdlg
    .Filename = "Stamppage.ps"
    .Filter = "Postscript Files (*.ps)|*.ps|(*.txt)|*.txt"
    .Flags = cdlOFNPathMustExist & cdlOFNOverwritePrompt Or cdlOFNLongNames
    .InitDir = App.Path & "\Stamppage"
    .ShowSave
    SaveFile .Filename, txtStamppage.Text
   End With
  Case 2 ' Show
   fn = FreeFile
   Open TempFile1 For Output As #fn
   Print #fn, txtStamppage.Text
   Close #fn
   GSViewPath = "c:\programme\ghostgum\gsview\gsview32.exe"
   If Dir(GSViewPath) <> "" Then
     Shell GSViewPath & " """ & TempFile1 & """"
    Else
     MsgBox "'gsview32.exe' cannot found. Please change this source in module 'cmdTestpage' in frmMain!", vbExclamation
   End If
  Case 3 ' Clipboard
   'Replace 0A0D with 0D and copy testpage to clipboard
   Clipboard.SetText Replace$(txtStamppage.Text, Chr$(&HD), ""), vbCFText
 End Select
 Exit Sub
ErrorHandler:
 If Err.Number = 32755 Then
  Exit Sub
 End If
 MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub cmdTestpage_Click(Index As Integer)
 Dim fn As Long, GSViewPath As String
 
 On Error GoTo ErrorHandler
 Select Case Index
  Case 0 ' Load
   With cdlg
    .CancelError = True
    .Filename = "PDFCreator.ps"
    .Filter = "Postscript Files (*.ps)|*.ps|(*.txt)|*.txt"
    .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist
    .InitDir = App.Path & "\Testpage"
    .ShowOpen
    LoadFileInTextbox .Filename, txtTestpage
    cmdTestpage(1).Enabled = True
    If Len(txtTestpage.Text) > 0 Then
     cmdTestpage(4).Enabled = True
'     cmdTestpage(5).Enabled = True
    End If
   End With
  Case 1 ' Save
   With cdlg
    .Filename = "PDFCreator.ps"
    .Filter = "Postscript Files (*.ps)|*.ps|(*.txt)|*.txt"
    .Flags = cdlOFNPathMustExist & cdlOFNOverwritePrompt Or cdlOFNLongNames
    .InitDir = App.Path & "\Testpage"
    .ShowSave
    SaveFile .Filename, txtTestpage.Text
   End With
  Case 2 ' Show
   fn = FreeFile
   Open TempFile1 For Output As #fn
   Print #fn, txtTestpage.Text
   Close #fn
   GSViewPath = "c:\programme\ghostgum\gsview\gsview32.exe"
   If Dir(GSViewPath) <> "" Then
     Shell GSViewPath & " """ & TempFile1 & """"
    Else
     MsgBox "'gsview32.exe' cannot found. Please change this source in module 'cmdTestpage' in frmMain!", vbExclamation
   End If
  Case 3 ' Clipboard
   'Replace 0A0D with 0D and copy testpage to clipboard
   Clipboard.SetText Replace$(txtTestpage.Text, Chr$(&HD), ""), vbCFText
  Case 4 ' Create modTestpage.bas
   With cdlg
    .CancelError = True
    .Filename = "modTestpage.bas"
    .Filter = "Postscript Files (*.bas)|*.bas"
    .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist
    .InitDir = App.Path & "\Testpage"
    .ShowSave
    CreateModTestpage .Filename, txtTestpage.Text
   End With
'  Case 5 ' Save compressed
'   With cdlg
'    .Filename = "PDFCreator-compressed.ps"
'    .Filter = "Postscript Files (*.ps)|*.ps|(*.txt)|*.txt"
'    .Flags = cdlOFNPathMustExist & cdlOFNOverwritePrompt Or cdlOFNLongNames
'    .InitDir = App.Path & "\Testpage"
'    .ShowSave
'    Dim b() As Byte
'    b = StrConv(txtTestpage.Text, vbFromUnicode)
'    Compress_Huffman_Dynamic b
'    SaveCompressedFile .Filename, b
'   End With
 End Select
 Exit Sub
ErrorHandler:
 If Err.Number = 32755 Then
  Exit Sub
 End If
 MsgBox Err.Number & " " & Err.Description
End Sub

Private Function ConvertType(TypeStr As String, convTypeTo As vType)
 TypeStr = Trim$(TypeStr)
 If Len(TypeStr) > 0 Then
  If convTypeTo = ShortString Then
   If UCase$(Mid$(TypeStr, 1, 1)) = "B" Then
    ConvertType = "B"
   End If
   If UCase$(Mid$(TypeStr, 1, 1)) = "L" Then
    ConvertType = "L"
   End If
   If UCase$(Mid$(TypeStr, 1, 1)) = "S" Then
    ConvertType = "S"
   End If
   If UCase$(Mid$(TypeStr, 1, 1)) = "D" Then
    ConvertType = "D"
   End If
  End If
  If convTypeTo = LongString Then
   If UCase$(Mid$(TypeStr, 1, 1)) = "B" Then
    ConvertType = "Boolean"
   End If
   If UCase$(Mid$(TypeStr, 1, 1)) = "L" Then
    ConvertType = "Long"
   End If
   If UCase$(Mid$(TypeStr, 1, 1)) = "S" Then
    ConvertType = "String"
   End If
   If UCase$(Mid$(TypeStr, 1, 1)) = "D" Then
    ConvertType = "Double"
   End If
  End If
 End If
End Function

Private Sub CreateModLanguages()
 Dim fn As Long, ini As New clsINI, Secs As Collection, keys As Collection, _
  i As Long, j As Long, tStr As String, Filename As String

 fn = FreeFile

 Filename = App.Path & "\..\Common\modLanguages.bas"

 Open Filename For Output As #fn
 Print #fn, "Attribute VB_Name = ""modLanguage"""
 Print #fn, "Option Explicit"
 Print #fn, ""
 Print #fn, "' Module automatically generated with LanguagesTool from Frank Heindörfer"
 Print #fn, "' 2004"
 Print #fn, "' Email: thesmilyface@users.sourceforge.net"
 Print #fn, ""
 Print #fn, "Public Type tLanguageStrings"

 ini.Filename = App.Path & "\..\PDFCreator\Languages\english.ini"
 Set Secs = ini.GetAllSectionsFromInifile(, True)
 For i = 1 To Secs.Count
  ini.Section = Secs.Item(i)
  If UCase$(Secs(i)) <> "SETUP" Then
   Set keys = ini.GetAllKeysFromSection(, , , True)
   For j = 1 To keys.Count
    Print #1, " " & Secs.Item(i) & keys.Item(j)(0) & " As String"
   Next j
   If i < Secs.Count Then
    Print #fn, ""
   End If
  End If
 Next i
 Print #fn, "End Type"
 Print #fn, ""

 Print #fn, "Public LanguageStrings As tLanguageStrings"
 Print #fn, ""
 Print #fn, "Public Sub LoadLanguage(ByVal Languagefile As String)"
 Print #fn, " InitLanguagesStrings"
 For i = 1 To Secs.Count
  If UCase$(Secs(i)) <> "SETUP" Then
   Print #fn, " Load" & Secs.Item(i) & "Strings Languagefile"
  End If
 Next i
 Print #fn, "End Sub"
 Print #fn, ""
 For i = 1 To Secs.Count
  If UCase$(Secs(i)) <> "SETUP" Then
   Print #fn, "Private Sub Load" & Secs.Item(i) & "Strings(ByVal Languagefile As String)"
   Print #fn, " Dim hLang As New clsHash"

   Print #fn, " ReadINISection Languagefile, """ & Secs.Item(i) & """, hLang"
   Print #fn, " With LanguageStrings"

   ini.Section = Secs.Item(i)
   Set keys = ini.GetAllKeysFromSection(, , , True)
   For j = 1 To keys.Count
    Print #fn, "  ." & Secs.Item(i) & keys.Item(j)(0) & " = Replace$(hLang.Retrieve(""" & keys.Item(j)(0) & """, ." & Secs.Item(i) & keys.Item(j)(0) & "),""/n"",vbCrLf)"
   Next j
   Print #fn, " End With"

   Print #fn, " Set hLang = Nothing"
   Print #fn, "End Sub"
   Print #fn, ""
  End If
 Next i

 ' InitLanguagesStrings
 Print #fn, "Public Sub InitLanguagesStrings()"
 Print #fn, " With LanguageStrings"
 For i = 1 To Secs.Count
  If UCase$(Secs(i)) <> "SETUP" Then
   ini.Section = Secs.Item(i)
   Set keys = ini.GetAllKeysFromSection(, , , True)
   For j = 1 To keys.Count
    Print #fn, "  ." & Secs.Item(i) & keys.Item(j)(0) & " = """ & keys.Item(j)(1) & """"
   Next j
   If i < Secs.Count Then
    Print #fn, ""
   End If
  End If
 Next i
 Print #fn, " End With"
 Print #fn, "End Sub"
 Print #fn, ""
 Close #fn
 With frmText
  .Filename = Filename
  .Show vbModal, Me
 End With
End Sub

Private Sub CreateModOptions()
 Dim fn As Long, Filename As String, i As Long, ma As Boolean, Filename2 As String

 With lsvOptions
  .SortOrder = lvwAscending
  .SortKey = 1
  .Sorted = True
 End With

 fn = FreeFile
 Filename = App.Path & "\..\Common\modOptions.bas"
 Open Filename For Output As #fn
 Print #fn, "Attribute VB_Name = ""modOptions"""
 Print #fn, "Option Explicit"
 Print #fn, ""
 Print #fn, "' Module automatically generated with LanguagesTool from Frank Heindörfer"
 Print #fn, "' 2003"
 Print #fn, "' Email: thesmilyface@users.sourceforge.net"
 Print #fn, ""
 Print #fn, "Public Type tOptions"
 For i = 1 To lsvOptions.ListItems.Count
  Select Case UCase$(lsvOptions.ListItems(i).SubItems(3))
   Case "BOOLEAN":
    Print #fn, " " & lsvOptions.ListItems(i).SubItems(1) & " As Long"
   Case Else:
    Print #fn, " " & lsvOptions.ListItems(i).SubItems(1) & " As " & lsvOptions.ListItems(i).SubItems(3)
  End Select
 Next i
 Print #fn, "End Type"
 Print #fn, ""
 Print #fn, "Public Options As tOptions"
 Print #fn, ""
 Print #fn, "Public Function StandardOptions() As tOptions"
 Print #fn, " Dim myOptions As tOptions, reg as clsRegistry"
 Print #fn, " With myOptions"
 For i = 1 To lsvOptions.ListItems.Count
  Select Case UCase$(lsvOptions.ListItems(i).SubItems(1))
   Case UCase$("DirectoryGhostscriptBinaries")
    Print #fn, "  Set reg = New clsRegistry"
    Print #fn, "  reg.hkey = HKEY_LOCAL_MACHINE"
    Print #fn, "  reg.KeyRoot = ""SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"" & Uninstall_GUID"
    Print #fn, "  ." & lsvOptions.ListItems(i).SubItems(1) & " = CompletePath(reg.GetRegistryValue(""GhostscriptDirectoryBinaries""))"
    Print #fn, "  Set reg = Nothing"
   Case UCase$("DirectoryGhostscriptLibraries")
    Print #fn, "  Set reg = New clsRegistry"
    Print #fn, "  reg.hkey = HKEY_LOCAL_MACHINE"
    Print #fn, "  reg.KeyRoot = ""SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"" & Uninstall_GUID"
    Print #fn, "  ." & lsvOptions.ListItems(i).SubItems(1) & " = CompletePath(reg.GetRegistryValue(""GhostscriptDirectoryLibraries""))"
    Print #fn, "  Set reg = Nothing"
   Case UCase$("DirectoryGhostscriptFonts")
    Print #fn, "  Set reg = New clsRegistry"
    Print #fn, "  reg.hkey = HKEY_LOCAL_MACHINE"
    Print #fn, "  reg.KeyRoot = ""SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"" & Uninstall_GUID"
    Print #fn, "  ." & lsvOptions.ListItems(i).SubItems(1) & " = CompletePath(reg.GetRegistryValue(""GhostscriptDirectoryFonts""))"
    Print #fn, "  Set reg = Nothing"
   Case UCase$("DirectoryGhostscriptResource")
    Print #fn, "  Set reg = New clsRegistry"
    Print #fn, "  reg.hkey = HKEY_LOCAL_MACHINE"
    Print #fn, "  reg.KeyRoot = ""SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"" & Uninstall_GUID"
    Print #fn, "  ." & lsvOptions.ListItems(i).SubItems(1) & " = CompletePath(reg.GetRegistryValue(""GhostscriptDirectoryResource""))"
    Print #fn, "  Set reg = Nothing"
   Case UCase$("DirectoryJava")
    Print #fn, "  ." & lsvOptions.ListItems(i).SubItems(1) & " = CompletePath(GetSpecialFolder(ssfSYSTEM))"
   Case UCase$("Printertemppath")
    Print #fn, "  If InstalledAsServer Then"
    Print #fn, "    ." & lsvOptions.ListItems(i).SubItems(1) & " = completepath(GetPDFCreatorApplicationPath) & ""Temp\"""
    Print #fn, "   Else"
    Print #fn, "    ." & lsvOptions.ListItems(i).SubItems(1) & " = ""<Temp>PDFCreator\"""
    Print #fn, "  End If"
   Case UCase$("AutoSaveDirectory"), UCase$("LastSaveDirectory")
    Print #fn, "  If InstalledAsServer Then"
    Print #fn, "    ." & lsvOptions.ListItems(i).SubItems(1) & " = ""C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"""
    Print #fn, "   Else"
    Print #fn, "    ." & lsvOptions.ListItems(i).SubItems(1) & " = ""<MyFiles>"""
    Print #fn, "  End If"
   Case Else
    If Len(lsvOptions.ListItems(i).SubItems(4)) = 0 Then
      Print #fn, "  ." & lsvOptions.ListItems(i).SubItems(1) & " = vbNullString"
     Else
      If UCase$(lsvOptions.ListItems(i).SubItems(3)) = "DOUBLE" Then
        Print #fn, "  ." & lsvOptions.ListItems(i).SubItems(1) & " = Replace$(""" & lsvOptions.ListItems(i).SubItems(4) & """, ""."", GetDecimalChar)"
       Else
        Print #fn, "  ." & lsvOptions.ListItems(i).SubItems(1) & " = """ & Replace$(lsvOptions.ListItems(i).SubItems(4), """", """""") & """"
      End If
    End If
  End Select
 Next i
 Print #fn, " End With"
 Print #fn, " If UseINI Then"
 Print #fn, "   If Not IsWin9xMe Then"
 Print #fn, "    myOptions = ReadOptionsINI(myOptions, CompletePath(GetDefaultAppData) & ""PDFCreator.ini"", False, False)"
 Print #fn, "   End If"
 Print #fn, "  Else"
 Print #fn, "   If Not IsWin9xMe Then"
 Print #fn, "    myOptions = ReadOptionsReg(myOptions, "".DEFAULT\Software\PDFCreator"", HKEY_USERS, False, False)"
 Print #fn, "   End If"
 Print #fn, " End If"
 Print #fn, " StandardOptions = myOptions"
 Print #fn, "End Function"
 Print #fn, ""

 Print #fn, "Public Function ReadOptions(Optional NoMsg As Boolean = False, Optional hProfile As hkey = HKEY_CURRENT_USER) As tOptions"
 Print #fn, " Dim myOptions As tOptions"
 Print #fn, " If InstalledAsServer Then"
 Print #fn, "   If UseINI Then"
 Print #fn, "     WriteToSpecialLogfile ""INI-Read options: CommonAppData"""
 Print #fn, "     myOptions = ReadOptionsINI(myOptions, Completepath(GetCommonAppData) & ""PDFCreator.ini"", HKEY_LOCAL_MACHINE, NoMsg)"
 Print #fn, "    Else"
 Print #fn, "     WriteToSpecialLogfile ""Reg-Read options: HKEY_LOCAL_MACHINE"""
 Print #fn, "     myOptions = ReadOptionsReg(myOptions, ""Software\PDFCreator"", HKEY_LOCAL_MACHINE, HKEY_LOCAL_MACHINE, NoMsg)"
 Print #fn, "   End If"
 Print #fn, "  Else"
 Print #fn, "   If UseINI Then"
 Print #fn, "     If Not IsWin9xMe Then"
 Print #fn, "       WriteToSpecialLogfile ""INI-Read options: DefaultAppData"""
 Print #fn, "       myOptions = ReadOptionsINI(myOptions, Completepath(GetDefaultAppData) & ""PDFCreator.ini"", HKEY_USERS, NoMsg)"
 Print #fn, "       WriteToSpecialLogfile ""INI-Read options: User"""
 Print #fn, "       myOptions = ReadOptionsINI(myOptions, PDFCreatorINIFile, hProfile, NoMsg, False)"
 Print #fn, "      Else"
 Print #fn, "       WriteToSpecialLogfile ""INI-Read options: User"""
 Print #fn, "       myOptions = ReadOptionsINI(myOptions, PDFCreatorINIFile, hProfile, NoMsg)"
 Print #fn, "     End If"
 Print #fn, "     WriteToSpecialLogfile ""INI-Read options: CommonAppData"""
 Print #fn, "     myOptions = ReadOptionsINI(myOptions, Completepath(GetCommonAppData) & ""PDFCreator.ini"", HKEY_LOCAL_MACHINE, NoMsg, False)"
 Print #fn, "    Else"
 Print #fn, "     If Not IsWin9xMe Then"
 Print #fn, "       WriteToSpecialLogfile ""Reg-Read options: HKEY_USERS"""
 Print #fn, "       myOptions = ReadOptionsReg(myOptions, "".DEFAULT\Software\PDFCreator"", HKEY_USERS, NoMsg)"
 Print #fn, "       WriteToSpecialLogfile ""Reg-Read options: HKEY_CURRENT_USER ["" & hProfile & ""]"""
 Print #fn, "       myOptions = ReadOptionsReg(myOptions, ""Software\PDFCreator"", hProfile, NoMsg, False)"
 Print #fn, "      Else"
 Print #fn, "       WriteToSpecialLogfile ""Reg-Read options: HKEY_CURRENT_USER ["" & hProfile & ""]"""
 Print #fn, "       myOptions = ReadOptionsReg(myOptions, ""Software\PDFCreator"", hProfile, NoMsg)"
 Print #fn, "     End If"
 Print #fn, "     WriteToSpecialLogfile ""Reg-Read options: HKEY_LOCAL_MACHINE"""
 Print #fn, "     myOptions = ReadOptionsReg(myOptions, ""Software\PDFCreator"", HKEY_LOCAL_MACHINE, NoMsg, False)"
 Print #fn, "   End If"
 Print #fn, " End If"
 Print #fn, " ReadOptions = myOptions"
 Print #fn, "End Function"
 Print #fn, ""

 Print #fn, "Public Function ReadOptionsINI(myOptions As tOptions, PDFCreatorINIFile As String, Optional hkey1 As hkey = HKEY_CURRENT_USER, Optional NoMsg as Boolean = False, Optional UseStandard as Boolean = True) As tOptions"
 Print #fn, " Dim ini As clsINI, tStr as String, hOpt As New clsHash"
 Print #fn, " ReadOptionsINI = myOptions"
 Print #fn, " Set ini = New clsINI"
 Print #fn, " ini.Filename = PDFCreatorINIFile"
 Print #fn, " ini.Section = ""Options"""
 Print #fn, " If ini.Checkinifile = False Then"
 Print #fn, "  If UseStandard Then"
 Print #fn, "   ReadOptionsINI = StandardOptions"
 Print #fn, "  End If"
 Print #fn, "  Exit Function"
 Print #fn, " End If"
 Print #fn, " ReadINISection PDFCreatorINIFile, ""Options"", hOpt"
 Print #fn, " With myOptions"
 For i = 1 To lsvOptions.ListItems.Count
  ma = False
  If UCase$(lsvOptions.ListItems(i).SubItems(3)) = "BOOLEAN" Then
   Print #fn, "  tStr = hOpt.Retrieve(""" & lsvOptions.ListItems(i).SubItems(1) & """)"
   Print #fn, "  If IsNumeric(tStr) Then"
   Print #fn, "    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then"
   Print #fn, "      ." & lsvOptions.ListItems(i).SubItems(1) & " = CLng(tStr)"
   Print #fn, "     Else"
   Print #fn, "      If UseStandard Then"
   Print #fn, "       ." & lsvOptions.ListItems(i).SubItems(1) & " = " & lsvOptions.ListItems(i).SubItems(4)
   Print #fn, "      End If"
   Print #fn, "    End If"
   Print #fn, "   Else"
   Print #fn, "    If UseStandard Then"
   Print #fn, "     ." & lsvOptions.ListItems(i).SubItems(1) & " = " & lsvOptions.ListItems(i).SubItems(4)
   Print #fn, "    End If"
   Print #fn, "  End If"
   ma = True
  End If
  If UCase$(lsvOptions.ListItems(i).SubItems(3)) = "LONG" Then
   Print #fn, "  tStr = hOpt.Retrieve(""" & lsvOptions.ListItems(i).SubItems(1) & """)"
   If lsvOptions.ListItems(i).SubItems(5) = ">" And lsvOptions.ListItems(i).SubItems(6) = "<" Then
    Print #fn, "  If IsNumeric(tStr) Then"
    Print #fn, "    ." & lsvOptions.ListItems(i).SubItems(1) & " = CLng(tStr)"
    Print #fn, "   Else"
    Print #fn, "    If UseStandard Then"
    Print #fn, "     ." & lsvOptions.ListItems(i).SubItems(1) & " = " & lsvOptions.ListItems(i).SubItems(4)
    Print #fn, "    End If"
    Print #fn, "  End If"
   End If
   If lsvOptions.ListItems(i).SubItems(5) <> ">" And lsvOptions.ListItems(i).SubItems(6) = "<" Then
    Print #fn, "  If IsNumeric(tStr) Then"
    Print #fn, "    If CLng(tStr) >= " & CLng(lsvOptions.ListItems(i).SubItems(5)) & " Then"
    Print #fn, "      ." & lsvOptions.ListItems(i).SubItems(1) & " = CLng(tStr)"
    Print #fn, "     Else"
    Print #fn, "      If UseStandard Then"
    Print #fn, "       ." & lsvOptions.ListItems(i).SubItems(1) & " = " & lsvOptions.ListItems(i).SubItems(4)
    Print #fn, "      End If"
    Print #fn, "    End If"
    Print #fn, "   Else"
    Print #fn, "    If UseStandard Then"
    Print #fn, "     ." & lsvOptions.ListItems(i).SubItems(1) & " = " & lsvOptions.ListItems(i).SubItems(4)
    Print #fn, "    End If"
    Print #fn, "  End If"
    ma = True
   End If
   If lsvOptions.ListItems(i).SubItems(5) <> ">" And lsvOptions.ListItems(i).SubItems(6) <> "<" Then
    Print #fn, "  If IsNumeric(tStr) Then"
    Print #fn, "    If CLng(tStr) >= " & CLng(lsvOptions.ListItems(i).SubItems(5)) & " And CLng(tStr) <= " & CLng(lsvOptions.ListItems(i).SubItems(6)) & " Then"
    Print #fn, "      ." & lsvOptions.ListItems(i).SubItems(1) & " = CLng(tStr)"
    Print #fn, "     Else"
    Print #fn, "      If UseStandard Then"
    Print #fn, "       ." & lsvOptions.ListItems(i).SubItems(1) & " = " & lsvOptions.ListItems(i).SubItems(4)
    Print #fn, "      End If"
    Print #fn, "    End If"
    Print #fn, "   Else"
    Print #fn, "    If UseStandard Then"
    Print #fn, "     ." & lsvOptions.ListItems(i).SubItems(1) & " = " & lsvOptions.ListItems(i).SubItems(4)
    Print #fn, "    End If"
    Print #fn, "  End If"
    ma = True
   End If
  End If
  If UCase$(lsvOptions.ListItems(i).SubItems(3)) = "STRING" Then
   Print #fn, "  tStr = hOpt.Retrieve(""" & lsvOptions.ListItems(i).SubItems(1) & """)"
   Select Case UCase$(lsvOptions.ListItems(i).SubItems(1))
    Case UCase$("AutoSaveDirectory"), UCase$("LastSaveDirectory")
     Print #fn, "  If LenB(Trim$(tStr)) > 0 Then"
     Print #fn, "    ." & lsvOptions.ListItems(i).SubItems(1) & " = CompletePath(tStr)"
     Print #fn, "   Else"
     Print #fn, "    If UseStandard Then"
     Print #fn, "     If InstalledAsServer Then"
     Print #fn, "       ." & lsvOptions.ListItems(i).SubItems(1) & " = ""C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"""
     Print #fn, "      Else"
     Print #fn, "       ." & lsvOptions.ListItems(i).SubItems(1) & " = ""<MyFiles>"""
     Print #fn, "     End If"
     Print #fn, "    End If"
     Print #fn, "  End If"
    Case UCase$("DirectoryGhostscriptBinaries")
     Print #fn, "  If LenB(Trim$(tStr)) > 0 Then"
     Print #fn, "    ." & lsvOptions.ListItems(i).SubItems(1) & " = CompletePath(tStr)"
     Print #fn, "   Else"
     Print #fn, "    If UseStandard Then"
     Print #fn, "     tStr = GetPDFCreatorApplicationPath"
     Print #fn, "     ." & lsvOptions.ListItems(i).SubItems(1) & " = CompletePath(tStr)"
     Print #fn, "    End If"
     Print #fn, "  End If"
    Case UCase$("DirectoryGhostscriptLibraries")
     Print #fn, "  If LenB(Trim$(tStr)) > 0 Then"
     Print #fn, "    ." & lsvOptions.ListItems(i).SubItems(1) & " = CompletePath(tStr)"
     Print #fn, "   Else"
     Print #fn, "    If UseStandard Then"
     Print #fn, "     tStr = GetPDFCreatorApplicationPath & ""lib"""
     Print #fn, "     ." & lsvOptions.ListItems(i).SubItems(1) & " = CompletePath(tStr)"
     Print #fn, "    End If"
     Print #fn, "  End If"
    Case UCase$("DirectoryGhostscriptFonts")
     Print #fn, "  If LenB(Trim$(tStr)) > 0 Then"
     Print #fn, "    ." & lsvOptions.ListItems(i).SubItems(1) & " = CompletePath(tStr)"
     Print #fn, "   Else"
     Print #fn, "    If UseStandard Then"
     Print #fn, "     tStr = GetPDFCreatorApplicationPath & ""fonts"""
     Print #fn, "     ." & lsvOptions.ListItems(i).SubItems(1) & " = CompletePath(tStr)"
     Print #fn, "    End If"
     Print #fn, "  End If"
    Case UCase$("Printertemppath")
     Print #fn, "  WriteToSpecialLogfile ""hOpt.Retrieve(""""PrinterTemppath"""")="" & tStr"
     Print #fn, "  WriteToSpecialLogfile ""Options.PrinterTemppath1="" & .PrinterTemppath"
     Print #fn, "  If hkey1 = HKEY_USERS Then"
     Print #fn, "    If LenB(tStr) > 0 And LenB(.PrinterTemppath) = 0 Then"
     Print #fn, "      .PrinterTemppath = tStr"
     Print #fn, "     Else"
     Print #fn, "      If UseStandard Then"
     Print #fn, "        .PrinterTemppath = GetTempPath"
     Print #fn, "       Else"
     Print #fn, "        .PrinterTemppath = tStr"
     Print #fn, "      End If"
     Print #fn, "    End If"
     Print #fn, "   Else"
     Print #fn, "    If LenB(Trim$(tStr)) > 0 Then"
     Print #fn, "     If DirExists(GetSubstFilename2(tStr, False, , , hkey1)) = True Then"
     Print #fn, "       .PrinterTemppath = tStr"
     Print #fn, "      Else"
     Print #fn, "       MakePath ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))"
     Print #fn, "       If DirExists(ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))) = False Then"
     Print #fn, "         If UseStandard Then"
     Print #fn, "           .PrinterTemppath = GetTempPath"
     Print #fn, "          Else"
     Print #fn, "           .PrinterTemppath = """""
     Print #fn, "           If NoMsg = False Then"
     Print #fn, "            MsgBox ""PrinterTemppath: '"" & tStr & ""' = '"" & ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1)) & ""'"" & _"
     Print #fn, "             vbCrLf & vbCrLf & LanguageStrings.MessagesMsg07"
     Print #fn, "           End If"
     Print #fn, "         End If"
     Print #fn, "        Else"
     Print #fn, "         .PrinterTemppath = tStr"
     Print #fn, "       End If"
     Print #fn, "     End If"
     Print #fn, "    End If"
     Print #fn, "  End If"
     Print #fn, "  WriteToSpecialLogfile ""Options.PrinterTemppath2="" & .PrinterTemppath"
    Case Else
     Print #fn, "  If LenB(tStr) = 0 And LenB(""" & Trim$(Replace$(lsvOptions.ListItems(i).SubItems(4), """", """""")) & """) > 0 And UseStandard Then"
     Print #fn, "    ." & lsvOptions.ListItems(i).SubItems(1) & " = """ & Replace$(lsvOptions.ListItems(i).SubItems(4), """", """""") & """"
     Print #fn, "   Else"
     Print #fn, "    If LenB(tStr) > 0 Then"
     Print #fn, "     ." & lsvOptions.ListItems(i).SubItems(1) & " = tStr"
     Print #fn, "    End If"
     Print #fn, "  End If"
   End Select
   ma = True
  End If
  If UCase$(lsvOptions.ListItems(i).SubItems(3)) = "DOUBLE" Then
   Print #fn, "  tStr = hOpt.Retrieve(""" & lsvOptions.ListItems(i).SubItems(1) & """)"
   If lsvOptions.ListItems(i).SubItems(5) = ">" And lsvOptions.ListItems(i).SubItems(6) = "<" Then
    Print #fn, "  If IsNumeric(Replace$(tStr, ""."", GetDecimalChar)) Then"
    Print #fn, "    ." & lsvOptions.ListItems(i).SubItems(1) & " = CDbl(Replace$(tStr, ""."", GetDecimalChar))"
    Print #fn, "   Else"
    Print #fn, "    If UseStandard Then"
    Print #fn, "     ." & lsvOptions.ListItems(i).SubItems(1) & " = Replace$(""" & lsvOptions.ListItems(i).SubItems(4) & """, ""."", GetDecimalChar)"
    Print #fn, "    End If"
    Print #fn, "  End If"
   End If
   If lsvOptions.ListItems(i).SubItems(5) <> ">" And lsvOptions.ListItems(i).SubItems(6) = "<" Then
    Print #fn, "  If IsNumeric(Replace$(tStr, ""."", GetDecimalChar)) Then"
    Print #fn, "    If CDbl(Replace$(tStr, ""."", GetDecimalChar)) >= " & CDbl(lsvOptions.ListItems(i).SubItems(5)) & " Then"
    Print #fn, "      ." & lsvOptions.ListItems(i).SubItems(1) & " = CDbl(Replace$(tStr, ""."", GetDecimalChar))"
    Print #fn, "     Else"
    Print #fn, "      If UseStandard Then"
    Print #fn, "       ." & lsvOptions.ListItems(i).SubItems(1) & " = Replace$(""" & lsvOptions.ListItems(i).SubItems(4) & """, ""."", GetDecimalChar)"
    Print #fn, "      End If"
    Print #fn, "    End If"
    Print #fn, "   Else"
    Print #fn, "    If UseStandard Then"
    Print #fn, "     ." & lsvOptions.ListItems(i).SubItems(1) & " = Replace$(""" & lsvOptions.ListItems(i).SubItems(4) & """, ""."", GetDecimalChar)"
    Print #fn, "    End If"
    Print #fn, "  End If"
    ma = True
   End If
   If lsvOptions.ListItems(i).SubItems(5) <> ">" And lsvOptions.ListItems(i).SubItems(6) <> "<" Then
    Print #fn, "  If IsNumeric(Replace$(tStr, ""."", GetDecimalChar)) Then"
    Print #fn, "    If CDbl(Replace$(tStr, ""."", GetDecimalChar)) >= " & CDbl(lsvOptions.ListItems(i).SubItems(5)) & " And CLng(tStr) <= " & CDbl(lsvOptions.ListItems(i).SubItems(6)) & " Then"
    Print #fn, "      ." & lsvOptions.ListItems(i).SubItems(1) & " = CDbl(Replace$(tStr, ""."", GetDecimalChar))"
    Print #fn, "     Else"
    Print #fn, "      If UseStandard Then"
    Print #fn, "       ." & lsvOptions.ListItems(i).SubItems(1) & " = Replace$(""" & lsvOptions.ListItems(i).SubItems(4) & """, ""."", GetDecimalChar)"
    Print #fn, "      End If"
    Print #fn, "    End If"
    Print #fn, "   Else"
    Print #fn, "    If UseStandard Then"
    Print #fn, "     ." & lsvOptions.ListItems(i).SubItems(1) & " = Replace$(""" & lsvOptions.ListItems(i).SubItems(4) & """, ""."", GetDecimalChar)"
    Print #fn, "    End If"
    Print #fn, "  End If"
    ma = True
   End If
  End If
  If ma = False Then
   MsgBox "Typ not Set"
   Stop
  End If
 Next i
 Print #fn, " End With"
 Print #fn, " Set ini = Nothing"
 Print #fn, " ReadOptionsINI = myOptions"
 Print #fn, "End Function"
 Print #fn, ""
 Print #fn, "Public Sub CorrectOptions()"
 Print #fn, " Options.AutosaveDirectory = Trim$(Options.AutosaveDirectory)"
 Print #fn, " Options.PrinterTemppath = Trim$(Options.PrinterTemppath)"
 Print #fn, " If LenB(Options.AutosaveDirectory) = 0 Then"
 Print #fn, "  Options.AutosaveDirectory = ""<MyFiles>\"""
 Print #fn, " End If"
 Print #fn, " If LenB(Options.PrinterTemppath) = 0 Then"
 Print #fn, "  Options.PrinterTemppath = ""<Temp>PDFCreator\"""
 Print #fn, " End If"
 Print #fn, "End Sub"
 Print #fn, ""
 Print #fn, "Public Sub SaveOptions(sOptions as tOptions)"
 Print #fn, " CorrectOptions"
 Print #fn, " If InstalledAsServer Then"
 Print #fn, "   If UseINI Then"
 Print #fn, "     SaveOptionsINI sOptions, Completepath(GetCommonAppData) & ""PDFCreator.ini"""
 Print #fn, "    Else"
 Print #fn, "     SaveOptionsReg sOptions, HKEY_LOCAL_MACHINE"
 Print #fn, "   End If"
 Print #fn, "  Else"
 Print #fn, "   If UseINI Then"
 Print #fn, "     SaveOptionsINI sOptions, PDFCreatorINIFile"
 Print #fn, "    Else"
 Print #fn, "     SaveOptionsReg sOptions"
 Print #fn, "   End If"
 Print #fn, " End If"
 Print #fn, "End Sub"
 Print #fn, ""
 Print #fn, "Public Sub SaveOption(sOptions As tOptions, OptionName As String)"
 Print #fn, " If InstalledAsServer Then"
 Print #fn, "   If UseINI Then"
 Print #fn, "     SaveOptionINI sOptions, OptionName, Completepath(GetCommonAppData) & ""PDFCreator.ini"""
 Print #fn, "    Else"
 Print #fn, "     SaveOptionReg sOptions, OptionName, HKEY_LOCAL_MACHINE"
 Print #fn, "   End If"
 Print #fn, "  Else"
 Print #fn, "   If UseINI Then"
 Print #fn, "     SaveOptionINI sOptions, OptionName, PDFCreatorINIFile"
 Print #fn, "    Else"
 Print #fn, "     SaveOptionReg sOptions, OptionName"
 Print #fn, "   End If"
 Print #fn, " End If"
 Print #fn, "End Sub"
 Print #fn, ""
 Print #fn, "Public Sub SaveOptionINI(sOptions As tOptions, OptionName As String, PDFCreatorINIFile As String)"
 Print #fn, " Dim ini As clsINI"
 Print #fn, " Set ini = New clsINI"
 Print #fn, " ini.Filename = PDFCreatorINIFile"
 Print #fn, " ini.Section = ""Options"""
 Print #fn, " If ini.CheckIniFile = False Then"
 Print #fn, "  ini.CreateIniFile"
 Print #fn, " End If"
 Print #fn, " With sOptions"
 Print #fn, "  Select Case UCase$(OptionName)"
 For i = 1 To lsvOptions.ListItems.Count
  Select Case UCase$(lsvOptions.ListItems(i).SubItems(3))
   Case "BOOLEAN": Print #fn, "  Case """ & UCase$(lsvOptions.ListItems(i).SubItems(1)) & """:ini.SaveKey CStr(Abs(." & lsvOptions.ListItems(i).SubItems(1) & ")), """ & lsvOptions.ListItems(i).SubItems(1) & """"
   Case "DOUBLE": Print #fn, "  Case """ & UCase$(lsvOptions.ListItems(i).SubItems(1)) & """:ini.SaveKey Replace$(CStr(." & lsvOptions.ListItems(i).SubItems(1) & "), GetDecimalChar, "".""), """ & lsvOptions.ListItems(i).SubItems(1) & """"
   Case Else: Print #fn, "  Case """ & UCase$(lsvOptions.ListItems(i).SubItems(1)) & """:ini.SaveKey CStr(." & lsvOptions.ListItems(i).SubItems(1) & "), """ & lsvOptions.ListItems(i).SubItems(1) & """"
  End Select
 Next i
 Print #fn, "  End Select"
 Print #fn, " End With"
 Print #fn, " Set ini = Nothing"
 Print #fn, "End Sub"
 Print #fn, ""
 Print #fn, "Public Sub SaveOptionsINI(sOptions as tOptions, PDFCreatorINIFile As String)"
 Print #fn, " Dim ini As clsINI"
 Print #fn, " Set ini = New clsINI"
 Print #fn, " ini.Filename = PDFCreatorINIFile"
 Print #fn, " ini.Section = ""Options"""
 Print #fn, " If ini.CheckInifile = False Then"
 Print #fn, "  ini.CreateInifile"
 Print #fn, " End If"
 Print #fn, " With sOptions"
 For i = 1 To lsvOptions.ListItems.Count
  Select Case UCase$(lsvOptions.ListItems(i).SubItems(3))
   Case "BOOLEAN": Print #fn, "  ini.SaveKey CStr(Abs(." & lsvOptions.ListItems(i).SubItems(1) & ")), """ & lsvOptions.ListItems(i).SubItems(1) & """"
   Case "DOUBLE": Print #fn, "  ini.SaveKey Replace$(CStr(." & lsvOptions.ListItems(i).SubItems(1) & "), GetDecimalChar, "".""), """ & lsvOptions.ListItems(i).SubItems(1) & """"
   Case Else: Print #fn, "  ini.SaveKey CStr(." & lsvOptions.ListItems(i).SubItems(1) & "), """ & lsvOptions.ListItems(i).SubItems(1) & """"
  End Select
 Next i
 Print #fn, " End With"
 Print #fn, " Set ini = Nothing"
 Print #fn, "End Sub"
 Print #fn, ""
  
 With lsvOptions
  .SortOrder = lvwAscending
  .SortKey = 1
  .Sorted = True
  .SortKey = 0
  .Sorted = True
 End With

 Print #fn, "Public Function ReadOptionsReg(myOptions As tOptions, KeyRoot as String, Optional hkey1 as hkey = HKEY_CURRENT_USER, Optional NoMsg as Boolean = False, Optional UseStandard as Boolean = True) As tOptions"
 Print #fn, " Dim reg As clsRegistry, tStr as String"
 Print #fn, " Set reg = New clsRegistry"
 Print #fn, " reg.hkey = hkey1"
 Print #fn, " reg.KeyRoot = KeyRoot"
 Print #fn, " With myOptions"
 For i = 1 To lsvOptions.ListItems.Count
  ma = False
  If i = 1 Then
    Print #fn, "  reg.Subkey = """ & Replace(Trim$(lsvOptions.ListItems(i).Text), " ", "\") & """"
   Else
    If UCase$(Replace(Trim$(lsvOptions.ListItems(i - 1).Text), " ", "\")) <> UCase$(Replace(Trim$(lsvOptions.ListItems(i).Text), " ", "\")) Then
     Print #fn, "  reg.Subkey = """ & Replace(Trim$(lsvOptions.ListItems(i).Text), " ", "\") & """"
    End If
  End If
  If UCase$(lsvOptions.ListItems(i).SubItems(3)) = "BOOLEAN" Then
   Print #fn, "  tStr = reg.GetRegistryValue(""" & lsvOptions.ListItems(i).SubItems(1) & """)"
   Print #fn, "  If IsNumeric(tStr) Then"
   Print #fn, "    If CLng(tStr) = 0 Or CLng(tStr) = 1 Then"
   Print #fn, "      ." & lsvOptions.ListItems(i).SubItems(1) & " = CLng(tStr)"
   Print #fn, "     Else"
   Print #fn, "      If UseStandard Then"
   Print #fn, "       ." & lsvOptions.ListItems(i).SubItems(1) & " = " & lsvOptions.ListItems(i).SubItems(4)
   Print #fn, "      End If"
   Print #fn, "    End If"
   Print #fn, "   Else"
   Print #fn, "    If UseStandard Then"
   Print #fn, "     ." & lsvOptions.ListItems(i).SubItems(1) & " = " & lsvOptions.ListItems(i).SubItems(4)
   Print #fn, "    End If"
   Print #fn, "  End If"
   ma = True
  End If
  If UCase$(lsvOptions.ListItems(i).SubItems(3)) = "LONG" Then
   Print #fn, "  tStr = reg.GetRegistryValue(""" & lsvOptions.ListItems(i).SubItems(1) & """)"
   If lsvOptions.ListItems(i).SubItems(5) = ">" And lsvOptions.ListItems(i).SubItems(6) = "<" Then
    Print #fn, "  If Isnumeric(tStr) Then"
    Print #fn, "    ." & lsvOptions.ListItems(i).SubItems(1) & " = CLng(tStr)"
    Print #fn, "   Else"
    Print #fn, "    If UseStandard Then"
    Print #fn, "     ." & lsvOptions.ListItems(i).SubItems(1) & " = " & lsvOptions.ListItems(i).SubItems(4)
    Print #fn, "    End If"
    Print #fn, "  End If"
   End If
   If lsvOptions.ListItems(i).SubItems(5) <> ">" And lsvOptions.ListItems(i).SubItems(6) = "<" Then
    Print #fn, "  If Isnumeric(tStr) Then"
    Print #fn, "    If CLng(tStr) >= " & CLng(lsvOptions.ListItems(i).SubItems(5)) & " Then"
    Print #fn, "      ." & lsvOptions.ListItems(i).SubItems(1) & " = CLng(tStr)"
    Print #fn, "     Else"
    Print #fn, "      If UseStandard Then"
    Print #fn, "       ." & lsvOptions.ListItems(i).SubItems(1) & " = " & lsvOptions.ListItems(i).SubItems(4)
    Print #fn, "      End If"
    Print #fn, "    End If"
    Print #fn, "   Else"
    Print #fn, "    If UseStandard Then"
    Print #fn, "     ." & lsvOptions.ListItems(i).SubItems(1) & " = " & lsvOptions.ListItems(i).SubItems(4)
    Print #fn, "    End If"
    Print #fn, "  End If"
    ma = True
   End If
   If lsvOptions.ListItems(i).SubItems(5) <> ">" And lsvOptions.ListItems(i).SubItems(6) <> "<" Then
    Print #fn, "  If Isnumeric(tStr) Then"
    Print #fn, "    If CLng(tStr) >= " & CLng(lsvOptions.ListItems(i).SubItems(5)) & " And CLng(tStr) <= " & CLng(lsvOptions.ListItems(i).SubItems(6)) & " Then"
    Print #fn, "      ." & lsvOptions.ListItems(i).SubItems(1) & " = CLng(tStr)"
    Print #fn, "     Else"
    Print #fn, "      If UseStandard Then"
    Print #fn, "       ." & lsvOptions.ListItems(i).SubItems(1) & " = " & lsvOptions.ListItems(i).SubItems(4)
    Print #fn, "      End If"
    Print #fn, "    End If"
    Print #fn, "   Else"
    Print #fn, "    If UseStandard Then"
    Print #fn, "     ." & lsvOptions.ListItems(i).SubItems(1) & " = " & lsvOptions.ListItems(i).SubItems(4)
    Print #fn, "    End If"
    Print #fn, "  End If"
    ma = True
   End If
  End If
  If UCase$(lsvOptions.ListItems(i).SubItems(3)) = "STRING" Then
   Print #fn, "  tStr = reg.GetRegistryValue(""" & lsvOptions.ListItems(i).SubItems(1) & """)"
   Select Case UCase$(lsvOptions.ListItems(i).SubItems(1))
    Case UCase$("AutoSaveDirectory"), UCase$("LastSaveDirectory")
     Print #fn, "  If LenB(Trim$(tStr)) > 0 Then"
     Print #fn, "    ." & lsvOptions.ListItems(i).SubItems(1) & " = CompletePath(tStr)"
     Print #fn, "   Else"
     Print #fn, "    If UseStandard Then"
     Print #fn, "     If InstalledAsServer Then"
     Print #fn, "       ." & lsvOptions.ListItems(i).SubItems(1) & " = ""C:\PDFs\<REDMON_MACHINE>\<REDMON_USER>"""
     Print #fn, "      Else"
     Print #fn, "       ." & lsvOptions.ListItems(i).SubItems(1) & " = ""<MyFiles>"""
     Print #fn, "     End If"
     Print #fn, "    End If"
     Print #fn, "  End If"
    Case UCase$("DirectoryGhostscriptBinaries")
     Print #fn, "  If LenB(Trim$(tStr)) > 0 Then"
     Print #fn, "    ." & lsvOptions.ListItems(i).SubItems(1) & " = CompletePath(tStr)"
     Print #fn, "   Else"
     Print #fn, "    If UseStandard Then"
     Print #fn, "     tStr = GetPDFCreatorApplicationPath"
     Print #fn, "     ." & lsvOptions.ListItems(i).SubItems(1) & " = CompletePath(tStr)"
     Print #fn, "    End If"
     Print #fn, "  End If"
    Case UCase$("DirectoryGhostscriptLibraries")
     Print #fn, "  If LenB(Trim$(tStr)) > 0 Then"
     Print #fn, "    ." & lsvOptions.ListItems(i).SubItems(1) & " = CompletePath(tStr)"
     Print #fn, "   Else"
     Print #fn, "    If UseStandard Then"
     Print #fn, "     tStr = GetPDFCreatorApplicationPath & ""lib"""
     Print #fn, "     ." & lsvOptions.ListItems(i).SubItems(1) & " = CompletePath(tStr)"
     Print #fn, "    End If"
     Print #fn, "  End If"
    Case UCase$("DirectoryGhostscriptFonts")
     Print #fn, "  If LenB(Trim$(tStr)) > 0 Then"
     Print #fn, "    ." & lsvOptions.ListItems(i).SubItems(1) & " = CompletePath(tStr)"
     Print #fn, "   Else"
     Print #fn, "    If UseStandard Then"
     Print #fn, "     tStr = GetPDFCreatorApplicationPath & ""fonts"""
     Print #fn, "     ." & lsvOptions.ListItems(i).SubItems(1) & " = CompletePath(tStr)"
     Print #fn, "    End If"
     Print #fn, "  End If"
    Case UCase$("Printertemppath")
     Print #fn, "  WriteToSpecialLogfile ""reg.GetRegistryValue(""""PrinterTemppath"""")="" & tStr"
     Print #fn, "  WriteToSpecialLogfile ""Options.PrinterTemppath1="" & .PrinterTemppath"
     Print #fn, "  If hkey1 = HKEY_USERS Then"
     Print #fn, "    If LenB(tStr) > 0 And LenB(.PrinterTemppath) = 0 Then"
     Print #fn, "      .PrinterTemppath = tStr"
     Print #fn, "     Else"
     Print #fn, "      If UseStandard Then"
     Print #fn, "        .PrinterTemppath = GetTempPath"
     Print #fn, "       Else"
     Print #fn, "        .PrinterTemppath = tStr"
     Print #fn, "      End If"
     Print #fn, "    End If"
     Print #fn, "   Else"
     Print #fn, "    If LenB(Trim$(tStr)) > 0 Then"
     Print #fn, "     If DirExists(GetSubstFilename2(tStr, False, , , hkey1)) = True Then"
     Print #fn, "       .PrinterTemppath = tStr"
     Print #fn, "      Else"
     Print #fn, "       MakePath ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))"
     Print #fn, "       If DirExists(ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1))) = False Then"
     Print #fn, "         If UseStandard Then"
     Print #fn, "           .PrinterTemppath = GetTempPath"
     Print #fn, "          Else"
     Print #fn, "           .PrinterTemppath = """""
     Print #fn, "           If NoMsg = False Then"
     Print #fn, "            MsgBox ""PrinterTemppath: '"" & tStr & ""' = '"" & ResolveEnvironment(GetSubstFilename2(tStr, False, , , hkey1)) & ""'"" & _"
     Print #fn, "             vbCrLf & vbCrLf & LanguageStrings.MessagesMsg07"
     Print #fn, "           End If"
     Print #fn, "         End If"
     Print #fn, "        Else"
     Print #fn, "         .PrinterTemppath = tStr"
     Print #fn, "       End If"
     Print #fn, "     End If"
     Print #fn, "    End If"
     Print #fn, "  End If"
     Print #fn, "  WriteToSpecialLogfile ""Options.PrinterTemppath2="" & .PrinterTemppath"
    Case Else
     Print #fn, "  If LenB(tStr) = 0 And LenB(""" & Trim$(Replace$(lsvOptions.ListItems(i).SubItems(4), """", """""")) & """) > 0 And UseStandard Then"
     Print #fn, "    ." & lsvOptions.ListItems(i).SubItems(1) & " = """ & Replace$(lsvOptions.ListItems(i).SubItems(4), """", """""") & """"
     Print #fn, "   Else"
     Print #fn, "    If LenB(tStr) > 0 Then"
     Print #fn, "     ." & lsvOptions.ListItems(i).SubItems(1) & " = tStr"
     Print #fn, "    End If"
     Print #fn, "  End If"
   End Select
   ma = True
  End If
  If UCase$(lsvOptions.ListItems(i).SubItems(3)) = "DOUBLE" Then
   Print #fn, "  tStr = reg.GetRegistryValue(""" & lsvOptions.ListItems(i).SubItems(1) & """)"
   If lsvOptions.ListItems(i).SubItems(5) = ">" And lsvOptions.ListItems(i).SubItems(6) = "<" Then
    Print #fn, "  If IsNumeric(tStr) Then"
    Print #fn, "    ." & lsvOptions.ListItems(i).SubItems(1) & " = CDbl(Replace$(tStr, ""."", GetDecimalChar))"
    Print #fn, "   Else"
    Print #fn, "    If UseStandard Then"
    Print #fn, "     ." & lsvOptions.ListItems(i).SubItems(1) & " = Replace$(""" & lsvOptions.ListItems(i).SubItems(4) & """, ""."", GetDecimalChar)"
    Print #fn, "    End If"
    Print #fn, "  End If"
   End If
   If lsvOptions.ListItems(i).SubItems(5) <> ">" And lsvOptions.ListItems(i).SubItems(6) = "<" Then
    Print #fn, "  If IsNumeric(Replace$(tStr, ""."", GetDecimalChar)) Then"
    Print #fn, "    If CDbl(Replace$(tStr, ""."", GetDecimalChar)) >= " & CDbl(lsvOptions.ListItems(i).SubItems(5)) & " Then"
    Print #fn, "      ." & lsvOptions.ListItems(i).SubItems(1) & " = CDbl(Replace$(tStr, ""."", GetDecimalChar))"
    Print #fn, "     Else"
    Print #fn, "      If UseStandard Then"
    Print #fn, "       ." & lsvOptions.ListItems(i).SubItems(1) & " = Replace$(""" & lsvOptions.ListItems(i).SubItems(4) & """, ""."", GetDecimalChar)"
    Print #fn, "      End If"
    Print #fn, "    End If"
    Print #fn, "   Else"
    Print #fn, "    If UseStandard Then"
    Print #fn, "     ." & lsvOptions.ListItems(i).SubItems(1) & " = Replace$(""" & lsvOptions.ListItems(i).SubItems(4) & """, ""."", GetDecimalChar)"
    Print #fn, "    End If"
    Print #fn, "  End If"
    ma = True
   End If
   If lsvOptions.ListItems(i).SubItems(5) <> ">" And lsvOptions.ListItems(i).SubItems(6) <> "<" Then
    Print #fn, "  If IsNumeric(Replace$(tStr, ""."", GetDecimalChar)) Then"
    Print #fn, "    If CDbl(Replace$(tStr, ""."", GetDecimalChar)) >= " & CDbl(lsvOptions.ListItems(i).SubItems(5)) & " And CLng(tStr) <= " & CDbl(lsvOptions.ListItems(i).SubItems(6)) & " Then"
    Print #fn, "      ." & lsvOptions.ListItems(i).SubItems(1) & " = CDbl(Replace$(tStr, ""."", GetDecimalChar))"
    Print #fn, "     Else"
    Print #fn, "      If UseStandard Then"
    Print #fn, "       ." & lsvOptions.ListItems(i).SubItems(1) & " = Replace$(""" & lsvOptions.ListItems(i).SubItems(4) & """, ""."", GetDecimalChar)"
    Print #fn, "      End If"
    Print #fn, "    End If"
    Print #fn, "   Else"
    Print #fn, "    If UseStandard Then"
    Print #fn, "     ." & lsvOptions.ListItems(i).SubItems(1) & " = Replace$(""" & lsvOptions.ListItems(i).SubItems(4) & """, ""."", GetDecimalChar)"
    Print #fn, "    End If"
    Print #fn, "  End If"
    ma = True
   End If
  End If
  If ma = False Then
   MsgBox "Typ not Set"
   Stop
  End If
 Next i
 Print #fn, " End With"
 Print #fn, " Set reg = Nothing"
 Print #fn, " ReadOptionsReg = MyOptions"
 Print #fn, "End Function"
 Print #fn, ""
 Print #fn, "Public Sub SaveOptionREG(sOptions as tOptions, OptionName as String, Optional hkey1 as hkey = HKEY_CURRENT_USER)"
 Print #fn, " Dim reg As clsRegistry"
 Print #fn, " Set reg = New clsRegistry"
 Print #fn, " reg.hkey = hkey1"
 Print #fn, " reg.KeyRoot = ""Software\PDFCreator"""
 Print #fn, " With sOptions"
 For i = 1 To lsvOptions.ListItems.Count
  If i = 1 Then
    Print #fn, "  reg.Subkey = """ & Replace(Trim$(lsvOptions.ListItems(i).Text), " ", "\") & """"
   Else
    If UCase$(Replace(Trim$(lsvOptions.ListItems(i - 1).Text), " ", "\")) <> UCase$(Replace(Trim$(lsvOptions.ListItems(i).Text), " ", "\")) Then
     Print #fn, "  reg.Subkey = """ & Replace(Trim$(lsvOptions.ListItems(i).Text), " ", "\") & """"
    End If
  End If
  Print #fn, "  If UCase$(OptionName) = """ & UCase$(lsvOptions.ListItems(i).SubItems(1)) & """ Then"
  Print #fn, "   If Not reg.KeyExists Then"
  Print #fn, "    reg.CreateKey"
  Print #fn, "   End If"
  Select Case UCase$(lsvOptions.ListItems(i).SubItems(3))
   Case "BOOLEAN": Print #fn, "   reg.SetRegistryValue """ & lsvOptions.ListItems(i).SubItems(1) & """,CStr(Abs(." & lsvOptions.ListItems(i).SubItems(1) & ")), REG_SZ"
   Case "DOUBLE": Print #fn, "  reg.SetRegistryValue """ & lsvOptions.ListItems(i).SubItems(1) & """, Replace$(CStr(." & lsvOptions.ListItems(i).SubItems(1) & "), GetDecimalChar, "".""), REG_SZ"
   Case Else: Print #fn, "   reg.SetRegistryValue """ & lsvOptions.ListItems(i).SubItems(1) & """,CStr(." & lsvOptions.ListItems(i).SubItems(1) & "), REG_SZ"
  End Select
 Print #fn, "   Set reg = Nothing"
 Print #fn, "   Exit Sub"
 Print #fn, "  End If"
 Next i
 Print #fn, " End With"
 Print #fn, " Set reg = Nothing"
 Print #fn, "End Sub"
 Print #fn, ""
 Print #fn, "Public Sub SaveOptionsREG(sOptions as tOptions, Optional hkey1 as hkey = HKEY_CURRENT_USER)"
 Print #fn, " Dim reg As clsRegistry"
 Print #fn, " Set reg = New clsRegistry"
 Print #fn, " reg.hkey = hkey1"
 Print #fn, " reg.KeyRoot = ""Software\PDFCreator"""
 Print #fn, " If Not reg.KeyExists Then"
 Print #fn, "  reg.CreateKey"
 Print #fn, " End If"
 Print #fn, " With sOptions"
 For i = 1 To lsvOptions.ListItems.Count
  If i = 1 Then
    Print #fn, "  reg.Subkey = """ & Replace(Trim$(lsvOptions.ListItems(i).Text), " ", "\") & """"
    Print #fn, "  If Not reg.KeyExists Then"
    Print #fn, "   reg.CreateKey"
    Print #fn, "  End If"
   Else
    If UCase$(Replace(Trim$(lsvOptions.ListItems(i - 1).Text), " ", "\")) <> UCase$(Replace(Trim$(lsvOptions.ListItems(i).Text), " ", "\")) Then
     Print #fn, "  reg.Subkey = """ & Replace(Trim$(lsvOptions.ListItems(i).Text), " ", "\") & """"
     Print #fn, "  If Not reg.KeyExists Then"
     Print #fn, "   reg.CreateKey"
     Print #fn, "  End If"
    End If
  End If
  Select Case UCase$(lsvOptions.ListItems(i).SubItems(3))
   Case "BOOLEAN": Print #fn, "  reg.SetRegistryValue """ & lsvOptions.ListItems(i).SubItems(1) & """, CStr(Abs(." & lsvOptions.ListItems(i).SubItems(1) & ")), REG_SZ"
   Case "DOUBLE": Print #fn, "  reg.SetRegistryValue """ & lsvOptions.ListItems(i).SubItems(1) & """, Replace$(CStr(." & lsvOptions.ListItems(i).SubItems(1) & "), GetDecimalChar, "".""), REG_SZ"
   Case Else: Print #fn, "  reg.SetRegistryValue """ & lsvOptions.ListItems(i).SubItems(1) & """,CStr(." & lsvOptions.ListItems(i).SubItems(1) & "), REG_SZ"
  End Select
 Next i
 Print #fn, " End With"
 Print #fn, " Set reg = Nothing"
 Print #fn, "End Sub"
 Print #fn, ""
 With lsvOptions
  .SortOrder = lvwAscending
  .SortKey = 1
  .Sorted = True
 End With
 Print #fn, "Public Sub ShowOptions(Frm as Form, sOptions as tOptions)"
 Print #fn, " On Error Resume Next"
 Print #fn, " Dim i as Long, tList() as String, tStrA() As String, lsv As ListView"
 Print #fn, " With sOptions"
 For i = 1 To lsvOptions.ListItems.Count
  If IsSpecialString(lsvOptions.ListItems(i).SubItems(1)) = False Then
   If UCase$(lsvOptions.ListItems(i).SubItems(1)) = UCase("Programfont") Then
     Print #fn, "  For i=0 to frm.cmbFonts.Listcount - 1"
     Print #fn, "    If Ucase$(frm.cmbFonts.List(i)) = Ucase$(." & lsvOptions.ListItems(i).SubItems(1) & ") Then"
     Print #fn, "     frm.cmbFonts.Listindex = i"
     Print #fn, "     Exit For"
     Print #fn, "    End If"
     Print #fn, "  Next i"
    Else
     If UCase$(lsvOptions.ListItems(i).SubItems(1)) = UCase$("FilenameSubstitutions") Then
       Print #fn, "  Set lsv = Frm.lsvFilenameSubst"
       Print #fn, "  tList = Split(.FilenameSubstitutions, ""\"")"
       Print #fn, "  For i = 0 To UBound(tList)"
       Print #fn, "   If InStr(tList(i), ""|"") <= 0 Then"
       Print #fn, "    tList(i) = tList(i) & ""|"""
       Print #fn, "   End If"
       Print #fn, "   If UBound(Split(tList(i), ""|"")) = 1 Then"
       Print #fn, "    tStrA = Split(tList(i), ""|"")"
       Print #fn, "    lsv.ListItems.Add , , tStrA(0)"
       Print #fn, "    lsv.ListItems(lsv.ListItems.Count).SubItems(1) = tStrA(1)"
       Print #fn, "   End If"
       Print #fn, "  Next i"
       Print #fn, "  If lsv.ListItems.Count > 0 Then"
       Print #fn, "   lsv.ListItems(1).Selected = True"
       Print #fn, "   Frm.txtFilenameSubst(0).Text = lsv.ListItems(1).Text"
       Print #fn, "   Frm.txtFilenameSubst(0).ToolTipText = Frm.txtFilenameSubst(0).Text"
       Print #fn, "   Frm.txtFilenameSubst(1).Text = lsv.ListItems(1).SubItems(1)"
       Print #fn, "   Frm.txtFilenameSubst(1).ToolTipText = Frm.txtFilenameSubst(1).Text"
       Print #fn, "  End If"
     Else
      If UCase$(lsvOptions.ListItems(i).SubItems(1)) <> UCase$("OptionsEnabled") And UCase$(lsvOptions.ListItems(i).SubItems(1)) <> UCase$("OptionsVisible") Then
       Print #fn, "  frm." & lsvOptions.ListItems(i).SubItems(2) & " = ." & lsvOptions.ListItems(i).SubItems(1)
      End If
    End If
   End If
  End If
 Next i
 Print #fn, " End With"
 Print #fn, "End Sub"
 Print #fn, ""
 Print #fn, "Public Sub GetOptions(Frm as Form, sOptions as tOptions)"
 Print #fn, " Dim i as Long, tStr as String, lsv As ListView"
 Print #fn, " With sOptions"
 For i = 1 To lsvOptions.ListItems.Count
  If IsSpecialString(lsvOptions.ListItems(i).SubItems(1)) = False Then
    If UCase$(lsvOptions.ListItems(i).SubItems(1)) = UCase$("FilenameSubstitutions") Then
      Print #fn, " tStr="""""
      Print #fn, " Set lsv = Frm.lsvFilenameSubst"
      Print #fn, " For i = 1 To lsv.ListItems.Count"
      Print #fn, "  If i < lsv.ListItems.Count Then"
      Print #fn, "    tStr = tStr & lsv.ListItems(i).Text & ""|"" & lsv.ListItems(i).SubItems(1) & ""\"""
      Print #fn, "   Else"
      Print #fn, "    tStr = tStr & lsv.ListItems(i).Text & ""|"" & lsv.ListItems(i).SubItems(1)"
      Print #fn, "  End If"
      Print #fn, " Next i"
      Print #fn, " ." & lsvOptions.ListItems(i).SubItems(1) & " = tStr"
     Else
      If UCase$(lsvOptions.ListItems(i).SubItems(1)) = UCase$("PDFEncryptor") And UCase$(lsvOptions.ListItems(i).SubItems(1)) <> UCase$("OptionsVisible") Then
        Print #fn, " If Frm.cmbPDFEncryptor.ListIndex < 0 Then"
        Print #fn, "   ." & lsvOptions.ListItems(i).SubItems(1) & " = 0"
        Print #fn, "  Else"
        Print #fn, "   ." & lsvOptions.ListItems(i).SubItems(1) & " =  frm." & lsvOptions.ListItems(i).SubItems(2)
        Print #fn, " End If"
       Else
        If UCase$(lsvOptions.ListItems(i).SubItems(1)) <> UCase$("OptionsEnabled") And UCase$(lsvOptions.ListItems(i).SubItems(1)) <> UCase$("OptionsVisible") Then
         If UCase$(lsvOptions.ListItems(i).SubItems(3)) = "BOOLEAN" Then
           Print #fn, " ." & lsvOptions.ListItems(i).SubItems(1) & " =  Abs(frm." & lsvOptions.ListItems(i).SubItems(2) & ")"
          Else
           Print #fn, " ." & lsvOptions.ListItems(i).SubItems(1) & " =  frm." & lsvOptions.ListItems(i).SubItems(2)
         End If
        End If
      End If
    End If
  End If
 Next i
 Print #fn, " End With"
 Print #fn, "End Sub"
 Print #fn, ""
 Print #fn, "Public Sub SetPrinterStop(StopPrinter as Boolean)"
 Print #fn, " If StopPrinter = True Then"
 Print #fn, "   Options.PrinterStop = 1"
 Print #fn, "   PrinterStop = True"
 Print #fn, "   PrintSelectedJobs = False"
 Print #fn, "  Else"
 Print #fn, "   Options.PrinterStop = 0"
 Print #fn, "   PrinterStop = False"
 Print #fn, " End If"
 Print #fn, " SaveOption Options, ""Printerstop"""
 Print #fn, "End Sub"
 Print #fn, ""
 Print #fn, "Public Sub SetLogging(Logging as Boolean)"
 Print #fn, " If Logging = True Then"
 Print #fn, "   Options.Logging = 1"
 Print #fn, "  Else"
 Print #fn, "   Options.Logging = 0"
 Print #fn, " End If"
 Print #fn, " SaveOption Options, ""Logging"""
 Print #fn, "End Sub"
 Print #fn, ""
 Print #fn, "Public Sub SetLanguage(Language as String)"
 Print #fn, " Options.Language = Language"
 Print #fn, " SaveOptions Options"
 Print #fn, "End Sub"
 Print #fn, ""
 Print #fn, "Public Sub ReadLanguageFromOptions(Optional hProfile As hkey = HKEY_CURRENT_USER)"
 Print #fn, " Dim sLanguage As String"
 Print #fn, " If InstalledAsServer Then"
 Print #fn, "   If UseINI Then"
 Print #fn, "     sLanguage = ReadLanguageFromOptionsINI(sLanguage, Completepath(GetCommonAppData) & ""PDFCreator.ini"")"
 Print #fn, "    Else"
 Print #fn, "     sLanguage = ReadLanguageFromOptionsReg(sLanguage, ""Software\PDFCreator"", HKEY_LOCAL_MACHINE)"
 Print #fn, "   End If"
 Print #fn, "  Else"
 Print #fn, "   If UseINI Then"
 Print #fn, "     If Not IsWin9xMe Then"
 Print #fn, "       sLanguage = ReadLanguageFromOptionsINI(sLanguage, Completepath(GetDefaultAppData) & ""PDFCreator.ini"")"
 Print #fn, "       sLanguage = ReadLanguageFromOptionsINI(sLanguage, PDFCreatorINIFile, False)"
 Print #fn, "      Else"
 Print #fn, "       sLanguage = ReadLanguageFromOptionsINI(sLanguage, PDFCreatorINIFile)"
 Print #fn, "     End If"
 Print #fn, "     sLanguage = ReadLanguageFromOptionsINI(sLanguage, Completepath(GetCommonAppData) & ""PDFCreator.ini"", False)"
 Print #fn, "    Else"
 Print #fn, "     If Not IsWin9xMe Then"
 Print #fn, "       sLanguage = ReadLanguageFromOptionsReg(sLanguage, "".DEFAULT\Software\PDFCreator"", HKEY_USERS)"
 Print #fn, "       sLanguage = ReadLanguageFromOptionsReg(sLanguage, ""Software\PDFCreator"", hProfile, False)"
 Print #fn, "      Else"
 Print #fn, "       sLanguage = ReadLanguageFromOptionsReg(sLanguage, ""Software\PDFCreator"", hProfile)"
 Print #fn, "     End If"
 Print #fn, "     sLanguage = ReadLanguageFromOptionsReg(sLanguage, ""Software\PDFCreator"", HKEY_LOCAL_MACHINE, False)"
 Print #fn, "   End If"
 Print #fn, " End If"
 Print #fn, " Options.Language = sLanguage"
 Print #fn, "End Sub"
 Print #fn, ""
 Print #fn, "Public Function ReadLanguageFromOptionsINI(Language As String, PDFCreatorINIFile As String, Optional UseStandard as Boolean = True) As String"
 Print #fn, " Dim hOpt As clsHash, tStr as String, opt as tOptions"
 Print #fn, " ReadLanguageFromOptionsINI = Language"
 Print #fn, " If FileExists(PDFCreatorINIFile) = False Then"
 Print #fn, "  If UseStandard Then"
 Print #fn, "   opt = StandardOptions"
 Print #fn, "   ReadLanguageFromOptionsINI = opt.Language"
 Print #fn, "  End If"
 Print #fn, "  Exit Function"
 Print #fn, " End If"
 Print #fn, " Set hOpt = New clsHash"
 Print #fn, " ReadINISection PDFCreatorINIFile, ""Options"", hOpt"
 Print #fn, " tStr = Trim$(hOpt.Retrieve(""Language""))"
 Print #fn, " If LenB(tStr) > 0 Then"
 Print #fn, "   ReadLanguageFromOptionsINI = tStr"
 Print #fn, "  Else"
 Print #fn, "   If UseStandard Then"
 Print #fn, "     ReadLanguageFromOptionsINI = ""english"""
 Print #fn, "    Else"
 Print #fn, "     ReadLanguageFromOptionsINI = Language"
 Print #fn, "   End If"
 Print #fn, " End If"
 Print #fn, " Set hOpt = Nothing"
 Print #fn, "End Function"
 Print #fn, ""
 Print #fn, "Public Function ReadLanguageFromOptionsReg(Language As String, KeyRoot as String, Optional hProfile as hkey = HKEY_CURRENT_USER, Optional UseStandard as Boolean = True) As String"
 Print #fn, " Dim reg As clsRegistry, tStr as String"
 Print #fn, " Set reg = New clsRegistry"
 Print #fn, " With reg"
 Print #fn, "  .KeyRoot = KeyRoot"
 Print #fn, "  .Subkey = ""Program"""
 Print #fn, "  .hkey = hProfile"
 Print #fn, "  tStr = Trim$(reg.GetRegistryValue(""Language""))"
 Print #fn, " End With"
 Print #fn, " If LenB(tStr) > 0 Then"
 Print #fn, "   ReadLanguageFromOptionsReg = tStr"
 Print #fn, "  Else"
 Print #fn, "   If UseStandard Then"
 Print #fn, "     ReadLanguageFromOptionsReg = ""english"""
 Print #fn, "    Else"
 Print #fn, "     ReadLanguageFromOptionsReg = Language"
 Print #fn, "   End If"
 Print #fn, " End If"
 Print #fn, " Set reg = Nothing"
 Print #fn, "End Function"
 Print #fn, ""
 Print #fn, "Public Function UseINI() As Boolean"
 Print #fn, " Dim reg As clsRegistry, tStr as String"
 Print #fn, " Set reg = New clsRegistry"
 Print #fn, " UseINI = False"
 Print #fn, " With reg"
 Print #fn, "  .hkey = HKEY_LOCAL_MACHINE"
 Print #fn, "  .KeyRoot = ""SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"" & Uninstall_GUID"
 Print #fn, "  tStr = Trim$(.GetRegistryValue(""UseINI""))"
 Print #fn, "  If tStr = ""1"" Then"
 Print #fn, "   UseINI = True"
 Print #fn, "  End If"
 Print #fn, " End With"
 Print #fn, " Set reg = Nothing"
 Print #fn, "End Function"
 Print #fn, ""
 Close #fn

 Dim inStrF As String, outStrF As String, replStr As String, _
  tStr1 As String, tStr2 As String
 Filename2 = App.Path & "\..\PDFCreator\clsPDFCreator.cls"
 If FileExists(Filename2) = False Then
  MsgBox "File: 'clsPDFCreator' doesn't exist!", vbCritical
  Exit Sub
 End If
 
 fn = FreeFile
 Open Filename2 For Binary As #fn
 inStrF = Space$(LOF(fn))
 Get #fn, , inStrF
 Close #fn
 tStr1 = "Private Function cGetOptions(sOptions As tOptions) As clsPDFCreatorOptions"
  
 replStr = "VERSION 1.0 CLASS"
 replStr = replStr & vbCrLf & "BEGIN"
 replStr = replStr & vbCrLf & "  MultiUse = -1  'True"
 replStr = replStr & vbCrLf & "  Persistable = 0  'NotPersistable"
 replStr = replStr & vbCrLf & "  DataBindingBehavior = 0  'vbNone"
 replStr = replStr & vbCrLf & "  DataSourceBehavior = 0   'vbNone"
 replStr = replStr & vbCrLf & "  MTSTransactionMode = 0   'NotAnMTSObject"
 replStr = replStr & vbCrLf & "End"
 replStr = replStr & vbCrLf & "Attribute VB_Name = ""clsPDFCreatorOptions"""
 replStr = replStr & vbCrLf & "Attribute VB_GlobalNameSpace = False"
 replStr = replStr & vbCrLf & "Attribute VB_Creatable = True"
 replStr = replStr & vbCrLf & "Attribute VB_PredeclaredId = False"
 replStr = replStr & vbCrLf & "Attribute VB_Exposed = True"
 replStr = replStr & vbCrLf & "Option Explicit"
 
 For i = 1 To lsvOptions.ListItems.Count
  Select Case UCase$(lsvOptions.ListItems(i).SubItems(3))
   Case "BOOLEAN":
    replStr = replStr & vbCrLf & "Public " & lsvOptions.ListItems(i).SubItems(1) & " As Long"
   Case Else:
    replStr = replStr & vbCrLf & "Public " & lsvOptions.ListItems(i).SubItems(1) & " As " & lsvOptions.ListItems(i).SubItems(3)
  End Select
 Next i
 fn = FreeFile
 Open App.Path & "\..\PDFCreator\clsPDFCreatorOptions.cls" For Output As #fn
 Print #fn, replStr
 Close #fn

 tStr1 = "Private Function cLetOptions(sOptions As Variant) As tOptions"
 tStr2 = "End Function"
 replStr = tStr1
 replStr = replStr & vbCrLf & " With cLetOptions"
 For i = 1 To lsvOptions.ListItems.Count
  replStr = replStr & vbCrLf & "  ." & lsvOptions.ListItems(i).SubItems(1) & " = sOptions." & lsvOptions.ListItems(i).SubItems(1)
 Next i
 replStr = replStr & vbCrLf & " End With"
' inStrF = outStrF
 If InStr(1, inStrF, tStr1, vbTextCompare) > 0 And InStr(1, inStrF, tStr2, vbTextCompare) > 0 Then
   outStrF = Mid(inStrF, 1, InStr(1, inStrF, tStr1, vbTextCompare) - 1) & _
    replStr & vbCrLf & Mid(inStrF, InStr(InStr(1, inStrF, tStr1, vbTextCompare), inStrF, _
    tStr2, vbTextCompare))
  Else
   MsgBox "Error 2 !!!", vbCritical
   Exit Sub
 End If

 tStr1 = "Private Function cGetOptions(sOptions As tOptions) As clsPDFCreatorOptions"
 tStr2 = "End Function"
 replStr = tStr1
 replStr = replStr & vbCrLf & " Set cGetOptions = New clsPDFCreatorOptions"
 replStr = replStr & vbCrLf & " With cGetOptions"
 For i = 1 To lsvOptions.ListItems.Count
  replStr = replStr & vbCrLf & "  ." & lsvOptions.ListItems(i).SubItems(1) & " = sOptions." & lsvOptions.ListItems(i).SubItems(1)
 Next i
 replStr = replStr & vbCrLf & " End With"
 inStrF = outStrF
 If InStr(1, inStrF, tStr1, vbTextCompare) > 0 And InStr(1, inStrF, tStr2, vbTextCompare) Then
   outStrF = Mid(inStrF, 1, InStr(1, inStrF, tStr1, vbTextCompare) - 1) & _
    replStr & vbCrLf & Mid(inStrF, InStr(InStr(1, inStrF, tStr1, vbTextCompare), inStrF, _
    tStr2, vbTextCompare))
  Else
   MsgBox "Error 3 !!!", vbCritical
   Exit Sub
 End If

 tStr1 = "Private Function cGetOptionFromOptions(PropertyName As String, Options As tOptions) As Variant"
 tStr2 = "End Function"
 replStr = tStr1
 replStr = replStr & vbCrLf & " Select Case UCase$(PropertyName)"
 For i = 1 To lsvOptions.ListItems.Count
  With lsvOptions.ListItems(i)
   replStr = replStr & vbCrLf & "  Case """ & UCase$(.SubItems(1)) & """: cGetOptionFromOptions = Options." & .SubItems(1)
  End With
 Next i
 replStr = replStr & vbCrLf & " End Select"
 inStrF = outStrF
 If InStr(1, inStrF, tStr1, vbTextCompare) > 0 And InStr(1, inStrF, tStr2, vbTextCompare) Then
   outStrF = Mid(inStrF, 1, InStr(1, inStrF, tStr1, vbTextCompare) - 1) & _
    replStr & vbCrLf & Mid(inStrF, InStr(InStr(1, inStrF, tStr1, vbTextCompare), inStrF, _
    tStr2, vbTextCompare))
  Else
   MsgBox "Error 4 !!!", vbCritical
   Exit Sub
 End If
 
 tStr1 = "Private Sub cLetOption(PropertyName As String, Value As Variant)"
 tStr2 = "End Sub"
 replStr = tStr1
 replStr = replStr & vbCrLf & " Select Case UCase$(PropertyName)"
 For i = 1 To lsvOptions.ListItems.Count
  With lsvOptions.ListItems(i)
   replStr = replStr & vbCrLf & "  Case """ & UCase$(.SubItems(1)) & """: Options." & .SubItems(1) & " = Value"
  End With
 Next i
 replStr = replStr & vbCrLf & "  Case Else:"
 replStr = replStr & vbCrLf & "   mError.Number = 3"
 replStr = replStr & vbCrLf & "   mError.Description = Replace$(Replace$(ErrDescr3, ""%1"", PropertyName), ""%2"", ""in cLetOption"")"
 replStr = replStr & vbCrLf & "   RaiseEvent eError"
 replStr = replStr & vbCrLf & " End Select"
 inStrF = outStrF
 If InStr(1, inStrF, tStr1, vbTextCompare) > 0 And InStr(1, inStrF, tStr2, vbTextCompare) Then
   outStrF = Mid(inStrF, 1, InStr(1, inStrF, tStr1, vbTextCompare) - 1) & _
    replStr & vbCrLf & Mid(inStrF, InStr(InStr(1, inStrF, tStr1, vbTextCompare), inStrF, _
    tStr2, vbTextCompare))
  Else
   MsgBox "Error 5 !!!", vbCritical
   Exit Sub
 End If
 
 tStr1 = "Public Property Get cOptionsNames() As Collection"
 tStr2 = "End Property"
 replStr = tStr1
 replStr = replStr & vbCrLf & " Set cOptionsNames = New Collection"
 replStr = replStr & vbCrLf & " With cOptionsNames"
 For i = 1 To lsvOptions.ListItems.Count
  With lsvOptions.ListItems(i)
   replStr = replStr & vbCrLf & "  .Add """ & .SubItems(1) & """"
  End With
 Next i
 replStr = replStr & vbCrLf & " End With"
 inStrF = outStrF
 If InStr(1, inStrF, tStr1, vbTextCompare) > 0 And InStr(1, inStrF, tStr2, vbTextCompare) Then
   outStrF = Mid(inStrF, 1, InStr(1, inStrF, tStr1, vbTextCompare) - 1) & _
    replStr & vbCrLf & Mid(inStrF, InStr(InStr(1, inStrF, tStr1, vbTextCompare), inStrF, _
    tStr2, vbTextCompare))
  Else
   MsgBox "Error 6 !!!", vbCritical
   Exit Sub
 End If
 
 fn = FreeFile
 Open Filename2 For Output As #fn
 Print #fn, outStrF
 Close #fn
 
 With frmText
  .Filename = Filename
  .Show vbModal, Me
 End With
End Sub

Private Function EngCount() As Long
 Dim i As Long, c As Long
 c = 0
 For i = 1 To lsvLanguages.ListItems.Count
  If Trim$(lsvLanguages.ListItems(i).SubItems(2)) <> "" Then
   c = c + 1
  End If
 Next i
 EngCount = c
End Function

Private Sub cmdPrintRegData_Click(Index As Integer)
 Dim RegFilename As String, IncludeFilename As String, fn As Long
 Select Case Index
  Case 0:
   RegFilename = CompletePath(App.Path) & "WinPrinterRegData\Win9xPrinterRegData.reg"
   If LenB(Dir(RegFilename)) > 0 Then
     If LoadAndConvert(RegFilename, Win9x) = False Then
      Exit Sub
     End If
    Else
     MsgBox "Regfile found!", vbCritical
     Exit Sub
   End If
   If LenB(txtPrintRegData.Text) > 0 Then
    cmdPrintRegData(1).Enabled = True
   End If
  Case 1:
   IncludeFilename = CompletePath(App.Path) & "..\Setup\Win9xPrinterRegData.inc"
   fn = FreeFile
   Open IncludeFilename For Output As fn
   Print #fn, txtPrintRegData.Text;
   Close #fn
  Case 2:
   RegFilename = CompletePath(App.Path) & "WinPrinterRegData\WinNTPrinterRegData.reg"
   If LenB(Dir(RegFilename)) > 0 Then
     If LoadAndConvert(RegFilename, WinNt) = False Then
      Exit Sub
     End If
    Else
     MsgBox "Regfile found!", vbCritical
     Exit Sub
   End If
   If LenB(txtPrintRegData.Text) > 0 Then
    cmdPrintRegData(3).Enabled = True
   End If
  Case 3:
   IncludeFilename = CompletePath(App.Path) & "..\Setup\WinNTPrinterRegData.inc"
   fn = FreeFile
   Open IncludeFilename For Output As fn
   Print #fn, txtPrintRegData.Text;
   Close #fn
 End Select
End Sub

Private Sub Form_Load()
 With lsvOptions.ColumnHeaders
  .Clear
  .Add , "Comment", "Comment", 1000
  .Add , "Name", "Name", 1000
  .Add , "ObjectName", "Objectname", 1000
  .Add , "Type", "Type", 1000
  .Add , "Standard", "Standard", 1000, lvwColumnRight
  .Add , "Left Limit", "Left Limit", 1000, lvwColumnRight
  .Add , "Right Limit", "Right Limit", 1000, lvwColumnRight
 End With
 ShowFrame
 With lsvLanguages.ColumnHeaders
  .Clear
  .Add , "Section", "Section", 1000
  .Add , "Key", "Key", 2000
  .Add , "English", "English", 1000
  .Add , "German", "German", 1000
 End With
 Me.Caption = Me.Caption & " [Version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
 TempFile1 = App.Path & "\Temp1.ps"
 ChangeOptions = False
 ChangeLanguages = False
 ChangeTestpage = False
 ChangeStamppage = False
End Sub

Private Sub Form_Resize()
 If WindowState <> 1 Then
  With tbstr
   .Top = Me.ScaleTop + 50
   .Left = Me.ScaleLeft + 50
   .Width = Me.ScaleWidth - 100
   .Height = Me.ScaleHeight - 100 - stb.Height
  End With
  ShowFrame
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Dim aw As Long
 If ChangeOptions = True Then
  aw = MsgBox("You have change the options! Cancel the program without saving?", vbQuestion Or vbYesNo)
  If aw = vbNo Then
   Cancel = True
   Exit Sub
  End If
 End If
 If ChangeLanguages = True Then
  aw = MsgBox("You have change the languages settings! Cancel the program without saving?", vbQuestion Or vbYesNo)
  If aw = vbNo Then
   Cancel = True
   Exit Sub
  End If
 End If
 If ChangeTestpage = True Then
  aw = MsgBox("You have change the testpage! Cancel the program without saving?", vbQuestion Or vbYesNo)
  If aw = vbNo Then
   Cancel = True
   Exit Sub
  End If
 End If
 If ChangeStamppage = True Then
  aw = MsgBox("You have change the stamppage! Cancel the program without saving?", vbQuestion Or vbYesNo)
  If aw = vbNo Then
   Cancel = True
   Exit Sub
  End If
 End If

 If Dir(TempFile1) <> "" Then
  Kill TempFile1
 End If
End Sub

Private Function GerCount() As Long
 Dim i As Long, c As Long
 c = 0
 For i = 1 To lsvLanguages.ListItems.Count
  If Trim$(lsvLanguages.ListItems(i).SubItems(3)) <> "" Then
   c = c + 1
  End If
 Next i
 GerCount = c
End Function

Private Sub LoadOptions(Filename As String)
 Dim fn As Long, tStr As String, tStrf() As String, i As Long, _
  Item As ListItem, j As Long, c As Long, aw As Long, Comment As String, flag As Long
 c = 0: flag = 0
 fn = FreeFile
 lsvOptions.ListItems.Clear
 Open Filename For Input As #fn
 Do While Not EOF(fn)
  aw = vbOK
  c = c + 1
  Line Input #fn, tStr
  tStr = Trim$(tStr)
  If Len(tStr) > 0 Then
   If Mid(tStr, 1, 1) <> "'" Then
     If InStr(tStr, "|") > 0 Then
       tStrf = Split(tStr, "|")
       Set Item = lsvOptions.ListItems.Add(, , Comment)
       For i = LBound(tStrf) To UBound(tStrf)
        If i = 2 Then
          If CheckType(tStrf(i)) = True Then
            Item.SubItems(i + 1) = tStrf(i)
           Else
            If flag < 1 Then
             aw = MsgBox("Is this a old 'options.txt'?", vbQuestion Or vbYesNo)
             If aw = vbYes Then
               flag = 1
              Else
               flag = 2
             End If
            End If
            If flag = 1 Then
             Select Case UCase$(tStrf(i))
              Case "B":
               Item.SubItems(i + 1) = "Boolean"
              Case "D":
               Item.SubItems(i + 1) = "Double"
              Case "L":
               Item.SubItems(i + 1) = "Long"
              Case "S":
               Item.SubItems(i + 1) = "String"
              Case Else:
               aw = MsgBox("Typeerror (unknown type '" & tStrf(i) & "') in line " & c, vbExclamation Or vbOKCancel)
             End Select
            End If
            If flag = 2 Then
             aw = MsgBox("Typeerror (unknown type '" & tStrf(i) & "') in line " & c, vbExclamation Or vbOKCancel)
            End If
          End If
'          item.SubItems(i) = ConvertType(tStrf(i), LongString)
         Else
          tStrf(i) = Trim$(tStrf(i))
          If Len(tStrf(i)) = 0 Then
           tStrf(i) = vbNullString
          End If
          Item.SubItems(i + 1) = tStrf(i)
        End If
       Next i
      Else
       Set Item = lsvOptions.ListItems.Add(, , Comment)
       Item.SubItems(1) = tStr
     End If
    Else
     i = InStr(tStr, "<"): j = InStr(tStr, ">")
     If i > 0 And j > 0 And j > i Then
      Comment = Mid(tStr, i + 1, j - i - 1)
     End If
   End If
  End If
  If aw = vbCancel Then
   Exit Sub
  End If
  DoEvents
 Loop
 Close #fn
 With lsvOptions
  .SortOrder = lvwAscending
  .SortKey = 1
  .Sorted = True
  .SortKey = 0
  .Sorted = True
  .Sorted = False
 End With
End Sub

Private Sub LoadFileInTextbox(Filename As String, txt As TextBox)
 Dim fn As Long
 fn = FreeFile
 Open Filename For Input As #fn
 Call SendMessage(txt.hwnd, WM_SETTEXT, 0&, ByVal CStr(Input(LOF(fn), #fn)))
 Close #fn
End Sub

Private Sub lsvLanguages_DblClick()
 EditItem = True
 ShowLanguage
End Sub

Private Sub lsvLanguages_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
  EditItem = True
  ShowLanguage
 End If
End Sub

Private Sub lsvOptions_DblClick()
 EditItem = True
 ShowOption
End Sub

Private Sub lsvOptions_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
 With lsvOptions
  If .FullRowSelect Then
   .Refresh
  End If
 End With
End Sub

Private Sub mnExit_Click()
 Unload Me
End Sub

Private Sub ReadTemplate()
 Dim ini As clsINI, Secs As Collection, keys As Collection, _
  i As Long, j As Long, k As Long, c As Long, Item As ListItem
 Set ini = New clsINI
 ini.Filename = App.Path & "\..\PDFCreator\Languages\english.ini"
 If ini.CheckIniFile = False Then
  MsgBox "File 'english.ini' not found!"
  Exit Sub
 End If
 Set Secs = ini.GetAllSectionsFromInifile(, True)
 For i = 1 To Secs.Count - 1
  c = 1
  For j = i + 1 To Secs.Count
   If Secs.Item(i) = Secs.Item(j) Then
    c = c + 1
   End If
  Next j
  If c > 1 Then
   MsgBox "Error: There are " & c & " Sections '" & Secs.Item(i) & "'!", vbExclamation
  End If
 Next i

 lsvLanguages.ListItems.Clear
 For i = 1 To Secs.Count
  Set keys = ini.GetAllKeysFromSection(Secs.Item(i), , , True)
  For j = 1 To keys.Count - 1
   c = 1
   For k = j + 1 To keys.Count
    If CStr(keys.Item(j)(0)) = CStr(keys.Item(k)(0)) Then
     c = c + 1
    End If
   Next k
   If c > 1 Then
    MsgBox "Error: There are " & c & " Keys '" & keys.Item(j)(0) & "' in section '" & Secs.Item(i) & "'!", vbExclamation
   End If
  Next j
  For j = 1 To keys.Count
   Set Item = lsvLanguages.ListItems.Add(, , Secs.Item(i))
   Item.SubItems(1) = keys.Item(j)(0)
   Item.SubItems(2) = keys.Item(j)(1)
  Next j
 Next i
 ini.Filename = App.Path & "\..\PDFCreator\Languages\german.ini"
 For i = 1 To lsvLanguages.ListItems.Count
  ini.Section = lsvLanguages.ListItems(i).Text
  ini.key = lsvLanguages.ListItems(i).SubItems(1)
  lsvLanguages.ListItems(i).SubItems(3) = ini.GetKeyFromSection
  DoEvents
 Next i
 stb.Panels("EngCount").Text = EngCount & " english entries"
 stb.Panels("GerCount").Text = GerCount & " german entries"
End Sub

Private Sub RefreshFrame()
 Dim i As Long
 Select Case tbstr.SelectedItem.Index - 1
  Case 0:
   With lsvOptions
    .Top = 200
    .Left = 100
    .Width = fra(0).Width - 200
    .Height = fra(0).Height - 800
    .ColumnHeaders(4).Width = 1000
    .ColumnHeaders(5).Width = 2000
    .ColumnHeaders(6).Width = 1000
    .ColumnHeaders(7).Width = 1000
    For i = 2 To 3
     .ColumnHeaders(i).Width = (.Width - 350 - (.ColumnHeaders(1).Width + .ColumnHeaders(4).Width + .ColumnHeaders(5).Width + .ColumnHeaders(6).Width + .ColumnHeaders(7).Width)) / 2
    Next i
    For i = 0 To 2
     cmdOptions(i).Top = .Top + .Height + 50
     cmdOptions(i).Left = .Left + i * (cmdOptions(i).Width + 100)
    Next i
    For i = 3 To 5
     cmdOptions(i).Top = .Top + .Height + 50
     cmdOptions(i).Left = .Left + .Width - (i - 2) * (cmdOptions(i).Width + 100) + 100
    Next i
   End With
   With stb
    .Panels.Clear
    .Panels.Add , "Count", lsvOptions.ListItems.Count & " Options"
   End With
  Case 1:
   With lsvLanguages
    .Top = 200
    .Left = 100
    .Width = fra(1).Width - 200
    .Height = fra(1).Height - 800
    .ColumnHeaders(1).Width = 1000
    .ColumnHeaders(2).Width = 2500
    For i = 3 To 4
     .ColumnHeaders(i).Width = (.Width - 350 - (.ColumnHeaders(1).Width + .ColumnHeaders(2).Width)) / 2
    Next i
    For i = 0 To 2
     cmdLanguages(i).Top = .Top + .Height + 50
     cmdLanguages(i).Left = .Left + i * (cmdLanguages(i).Width + 100)
    Next i
    For i = 3 To 5
     cmdLanguages(i).Top = .Top + .Height + 50
     cmdLanguages(i).Left = .Left + .Width - (i - 2) * (cmdLanguages(i).Width + 100) + 100
    Next i
   End With
   With stb
    .Panels.Clear
    .Panels.Add , "Count", lsvOptions.ListItems.Count & " Entries"
    .Panels.Add , "EngCount", lsvOptions.ListItems.Count & " Entries"
    .Panels("EngCount").Width = 2000
    .Panels.Add , "GerCount", lsvOptions.ListItems.Count & " Entries"
    .Panels("GerCount").Width = 2000
   End With
  Case 2:
   With txtTestpage
    .Top = 200
    .Left = 100
    .Width = fra(tbstr.SelectedItem.Index - 1).Width - 200
    .Height = fra(tbstr.SelectedItem.Index - 1).Height - 900
   End With
   With cmdTestpage
    .Item(0).Top = txtTestpage.Top + txtTestpage.Height + 50
    .Item(0).Left = txtTestpage.Left
    .Item(1).Top = .Item(0).Top
    .Item(1).Left = .Item(0).Left + .Item(0).Width + 150
    .Item(2).Top = .Item(0).Top
    .Item(2).Left = .Item(1).Left + .Item(1).Width + 150
    .Item(3).Top = .Item(0).Top
    .Item(3).Left = .Item(2).Left + .Item(2).Width + 150
    .Item(4).Top = .Item(0).Top
    .Item(4).Left = .Item(3).Left + .Item(3).Width + 150
'    .Item(5).Top = .Item(0).Top
'    .Item(5).Left = .Item(4).Left + .Item(4).Width + 150
   End With
   With stb
    .Panels.Clear
   End With
  Case 3:
   With txtStamppage
    .Top = 200
    .Left = 100
    .Width = fra(tbstr.SelectedItem.Index - 1).Width - 200
    .Height = fra(tbstr.SelectedItem.Index - 1).Height - 900
   End With
   With cmdStamppage
    .Item(0).Top = txtStamppage.Top + txtStamppage.Height + 50
    .Item(0).Left = txtStamppage.Left
    .Item(1).Top = .Item(0).Top
    .Item(1).Left = .Item(0).Left + .Item(0).Width + 150
    .Item(2).Top = .Item(0).Top
    .Item(2).Left = .Item(1).Left + .Item(1).Width + 150
    .Item(3).Top = .Item(0).Top
    .Item(3).Left = .Item(2).Left + .Item(2).Width + 150
   End With
   With stb
    .Panels.Clear
   End With
  Case 4:
   With txtIncFile
    .Top = 200
    .Left = 100
    .Width = fra(tbstr.SelectedItem.Index - 1).Width - 200
    .Height = fra(tbstr.SelectedItem.Index - 1).Height - 900
   End With
   With cmdIncFile
    .Item(0).Top = txtIncFile.Top + txtIncFile.Height + 50
    .Item(0).Left = txtIncFile.Left
    .Item(1).Top = .Item(0).Top
    .Item(1).Left = .Item(0).Left + .Item(0).Width + 150
    .Item(2).Top = .Item(1).Top
    .Item(2).Left = .Item(1).Left + .Item(1).Width + 150
   End With
   With stb
    .Panels.Clear
   End With
  Case 5:
   With stb
    .Panels.Clear
   End With
   With txtPrintRegData
    .Top = 200
    .Left = 100
    .Width = fra(tbstr.SelectedItem.Index - 1).Width - 200
    .Height = fra(tbstr.SelectedItem.Index - 1).Height - 900
   End With
   cmdPrintRegData(0).Top = txtPrintRegData.Top + txtPrintRegData.Height + 50
   cmdPrintRegData(0).Left = txtPrintRegData.Left
   cmdPrintRegData(1).Top = cmdPrintRegData(0).Top
   cmdPrintRegData(1).Left = cmdPrintRegData(0).Left + cmdPrintRegData(0).Width + 150
   cmdPrintRegData(2).Top = txtPrintRegData.Top + txtPrintRegData.Height + 50
   cmdPrintRegData(2).Left = cmdPrintRegData(1).Left + cmdPrintRegData(0).Width + 550
   cmdPrintRegData(3).Top = cmdPrintRegData(0).Top
   cmdPrintRegData(3).Left = cmdPrintRegData(2).Left + cmdPrintRegData(0).Width + 150
 End Select
End Sub

Private Sub ShowFrame()
 Dim i As Long
 For i = 1 To tbstr.Tabs.Count
  fra(i - 1).Visible = False
 Next i
 i = tbstr.SelectedItem.Index
 With fra(i - 1)
  .Left = tbstr.ClientLeft
  .Top = tbstr.ClientTop
  .Width = tbstr.ClientWidth
  .Height = tbstr.ClientHeight
  .Visible = True
  .ZOrder
 End With
 RefreshFrame
End Sub

Private Sub tbstr_Click()
 tbstr.Enabled = False
 ShowFrame
 tbstr.Enabled = True
End Sub

Private Sub SaveOptions(Filename As String)
 Dim i As Long, j As Long, fn As Long, tStr As String, tStrC As String
 fn = FreeFile
 Open Filename For Output As #fn
 tStrC = lsvOptions.ListItems(1).Text
 Print #fn, "'Optionname|OptionControlvalue|Type|Standard|LeftLimit|RightLimit"
 Print #fn, ""
 Print #fn, "' <" & tStrC & ">"
 For i = 1 To lsvOptions.ListItems.Count
  If UCase$(tStrC) <> UCase$(lsvOptions.ListItems(i).Text) Then
   tStrC = Trim$(lsvOptions.ListItems(i).Text)
   Print #fn, vbCrLf & "' <" & tStrC & ">"
  End If
  tStr = lsvOptions.ListItems(i).SubItems(1)
  For j = 2 To lsvOptions.ColumnHeaders.Count - 1
   tStr = tStr & "|" & Trim$(lsvOptions.ListItems(i).SubItems(j))
  Next j
  Print #fn, tStr
 Next i
 Close fn
End Sub

Private Sub SaveFile(Filename As String, txtStr As String)
 Dim fn As Long
 fn = FreeFile
 Open Filename For Output As #fn
 Print #fn, txtStr
 Close #fn
End Sub

Private Sub SaveCompressedFile(Filename As String, B() As Byte)
 Dim fn As Long
 fn = FreeFile
 Open Filename For Binary As #fn
 Put #fn, , B
 Close #fn
End Sub

Private Sub SaveTemplate()
 Dim ini As clsINI, i As Long

 Set ini = New clsINI
 If DirExists(App.Path & "\..\PDFCreator\Languages\") = False Then
  MkDir App.Path & "\..\PDFCreator\Languages\"
 End If
 ini.Filename = App.Path & "\..\PDFCreator\Languages\english.ini"
 ini.CreateIniFile
 With lsvLanguages
  For i = 1 To .ListItems.Count
   ini.SaveKey .ListItems(i).SubItems(2), .ListItems(i).SubItems(1), .ListItems(i).Text
   DoEvents
  Next i
 End With

 ini.Filename = App.Path & "\..\PDFCreator\Languages\german.ini"
 ini.CreateIniFile
 With lsvLanguages
  For i = 1 To .ListItems.Count
   ini.SaveKey .ListItems(i).SubItems(3), .ListItems(i).SubItems(1), .ListItems(i).Text
   DoEvents
  Next i
 End With
 Set ini = Nothing
 FileCopy App.Path & "\..\PDFCreator\Languages\english.ini", App.Path & "\..\TransTool\english.ini"
 FileCopy App.Path & "\..\PDFCreator\Languages\german.ini", App.Path & "\..\TransTool\german.ini"
End Sub

Private Sub ShowOption()
 Dim Item As ListItem, tStr As String, i As Long
 With frmOption
  .cmbComment.Text = lsvOptions.SelectedItem.Text
  If lsvOptions.ListItems.Count > 0 Then
   tStr = UCase$(lsvOptions.ListItems(1).Text)
   .cmbComment.AddItem lsvOptions.ListItems(1).Text
   For i = 2 To lsvOptions.ListItems.Count
    If UCase$(lsvOptions.ListItems(i).Text) <> tStr Then
     tStr = UCase$(lsvOptions.ListItems(i).Text)
     .cmbComment.AddItem lsvOptions.ListItems(i).Text
    End If
   Next i
  End If
  .txt(0).Text = lsvOptions.SelectedItem.SubItems(1)
  .txt(1).Text = lsvOptions.SelectedItem.SubItems(2)

  If UCase$(lsvOptions.SelectedItem.SubItems(3)) = "BOOLEAN" Then
   .cmbType.ListIndex = 0
  End If
  If UCase$(lsvOptions.SelectedItem.SubItems(3)) = "BYTE" Then
   .cmbType.ListIndex = 1
  End If
  If UCase$(lsvOptions.SelectedItem.SubItems(3)) = "LONG" Then
   .cmbType.ListIndex = 2
  End If
  If UCase$(lsvOptions.SelectedItem.SubItems(3)) = "STRING" Then
   .cmbType.ListIndex = 3
  End If
  If UCase$(lsvOptions.SelectedItem.SubItems(3)) = "DOUBLE" Then
   .cmbType.ListIndex = 4
  End If

  .txt(2).Text = lsvOptions.SelectedItem.SubItems(4)
  .txt(3).Text = lsvOptions.SelectedItem.SubItems(5)
  .txt(4).Text = lsvOptions.SelectedItem.SubItems(6)
  .Show vbModal, Me
 End With
End Sub

Private Sub ShowLanguage()
 Dim Item As ListItem, tStr As String, i As Long
 With frmLanguage
  .cmbSection.Text = lsvLanguages.SelectedItem.Text

  If lsvLanguages.ListItems.Count > 0 Then
   tStr = UCase$(lsvLanguages.ListItems(1).Text)
   .cmbSection.AddItem lsvLanguages.ListItems(1).Text
   For i = 2 To lsvLanguages.ListItems.Count
    If UCase$(lsvLanguages.ListItems(i).Text) <> tStr Then
     tStr = UCase$(lsvLanguages.ListItems(i).Text)
     .cmbSection.AddItem lsvLanguages.ListItems(i).Text
    End If
   Next i
  End If

  .txt(0).Text = lsvLanguages.SelectedItem.SubItems(1)
  .txt(1).Text = lsvLanguages.SelectedItem.SubItems(2)
  .txt(2).Text = lsvLanguages.SelectedItem.SubItems(3)
  .Show vbModal, Me
 End With
End Sub

Private Sub txtIncFile_Change()
 If Len(txtIncFile.Text) = 0 Then
   cmdIncFile(1).Enabled = False
  Else
   cmdIncFile(1).Enabled = True
 End If
End Sub

Private Sub txtStamppage_Change()
 If Len(txtStamppage.Text) = 0 Then
   cmdStamppage(1).Enabled = False
   cmdStamppage(2).Enabled = False
   cmdStamppage(3).Enabled = False
  Else
   cmdStamppage(1).Enabled = True
   cmdStamppage(2).Enabled = True
   cmdStamppage(3).Enabled = True
 End If
End Sub

Private Sub txtTestpage_Change()
 If Len(txtTestpage.Text) = 0 Then
   cmdTestpage(1).Enabled = False
   cmdTestpage(2).Enabled = False
   cmdTestpage(3).Enabled = False
  Else
   cmdTestpage(1).Enabled = True
   cmdTestpage(2).Enabled = True
   cmdTestpage(3).Enabled = True
 End If
End Sub

Private Function LoadAndConvert(RegFilename As String, os As eOSTyp) As Boolean
 Dim fn As Long, found As Boolean, resStr As String, tStr As String, sa() As String, _
  i As Long, j As Long, regStr As String

 LoadAndConvert = True
 Select Case os
  Case eOSTyp.Win9x
   regStr = """DEFAULT DEVMODE""=HEX:"
  Case eOSTyp.WinNt
   regStr = """PRINTERDATA""=HEX:"
 End Select
 found = False: fn = FreeFile
 Open RegFilename For Input As #fn
 Do While Not EOF(fn)
  Line Input #fn, tStr
  If InStr(UCase$(tStr), regStr) > 0 Then
   found = True
   resStr = Mid(tStr, Len(regStr) + 1)
   Do While Not EOF(fn)
    Line Input #fn, tStr
    If InStr(tStr, """") Or LenB(Trim$(tStr)) = 0 Or Mid$(Trim$(tStr), 1, 1) = "[" Then
     Exit Do
    End If
    resStr = resStr & tStr
    DoEvents
   Loop
  End If
  If found = True Then
   Exit Do
  End If
  DoEvents
 Loop
 Close #fn
 If found = False Then
   MsgBox "Cannot found the related printer regdata!"
   LoadAndConvert = False
   Exit Function
  Else
   resStr = Replace$(resStr, "\", "")
   resStr = Replace$(resStr, " ", "")
   resStr = Replace$(resStr, vbCr, "")
   resStr = Replace$(resStr, vbLf, "")
   sa = Split(resStr, ",")
   tStr = "": resStr = ""
   j = 0
   For i = 0 To UBound(sa)
    j = j + 1
    If j = 1 Then
     tStr = tStr + "''"
    End If
    tStr = tStr + GetHStr(sa(i))
    If j = 32 Then
     tStr = tStr + vbCrLf
     j = 0
    End If
   Next i
 End If
 If LenB(tStr) >= 3 Then
  If Mid(tStr, Len(tStr) - 1) = vbCrLf Then
   tStr = Mid(tStr, 1, Len(tStr) - 1)
  End If
 End If
 txtPrintRegData.Text = tStr
End Function

Private Function GetHStr(NumberStr As String) As String
 Dim tStr As String, i As Long
 tStr = CStr(CLng("&h" + NumberStr))
 For i = 1 To 3 - Len(tStr)
  tStr = "0" & tStr
 Next i
 GetHStr = "#" + tStr
End Function

Private Function IsSpecialString(specialString As String) As Boolean
 Dim ss As Collection, i As Long
 Set ss = New Collection
 With ss
  .Add "PrinterStop"
  .Add "LastSaveDirectory"
  .Add "Language"
  .Add "Logging"
  .Add "LogLines"
  .Add "GetOptions"
  .Add "OnePagePerFile"
  .Add "RemoveAllKnownFileExtensions"
  .Add "PDFOwnerPasswordString"
  .Add "PDFUserPasswordString"
  .Add "SendMailMethod"
  .Add "StandardKeywords"
  .Add "StandardDateformat"
  .Add "StandardCreationdate"
  .Add "StandardModifydate"
  .Add "StandardSubject"
  .Add "StandardTitle"
  .Add "StampFontColor"
  .Add "StampFontname"
  .Add "StampFontsize"
  .Add "StampOutlineFontthickness"
  .Add "StampString"
  .Add "StampUseOutlineFont"
  .Add "StartStandardProgram"
  .Add "StandardSaveformat"
  .Add "PDFOptimize"
  .Add "StandardMailDomain"
  .Add "Toolbars"
  .Add "DontUseDocumentSettings"
  .Add "Papersize"
  .Add "DeviceHeightPoints"
  .Add "DeviceWidthPoints"
  .Add "ClientComputerResolveIPAddress"
  .Add "DisableEmail"
  .Add "PrintAfterSaving"
  .Add "PrintAfterSavingPrinter"
  .Add "PrintAfterSavingNoCancel"
  .Add "PrintAfterSavingQueryUser"
  .Add "PrintAfterSavingDuplex"
  .Add "PrintAfterSavingTumble"
  .Add "NoPSCheck"
  .Add "PDFCompressionColorCompressionJPEGMaximumFactor"
  .Add "PDFCompressionColorCompressionJPEGHighFactor"
  .Add "PDFCompressionColorCompressionJPEGMediumFactor"
  .Add "PDFCompressionColorCompressionJPEGLowFactor"
  .Add "PDFCompressionColorCompressionJPEGMinimumFactor"
  .Add "PDFCompressionGreyCompressionJPEGMaximumFactor"
  .Add "PDFCompressionGreyCompressionJPEGHighFactor"
  .Add "PDFCompressionGreyCompressionJPEGMediumFactor"
  .Add "PDFCompressionGreyCompressionJPEGLowFactor"
  .Add "PDFCompressionGreyCompressionJPEGMinimumFactor"
  .Add "SendEmailAfterAutoSaving"
 End With
 IsSpecialString = False
 For i = 1 To ss.Count
  If UCase$(ss(i)) = UCase$(specialString) Then
   IsSpecialString = True
   Exit For
  End If
 Next i
 Set ss = Nothing
End Function

Private Function GetKeysAndValuesFromInifile(Section As String, Filename As String) As String
 Dim ini As New clsINI, keys As Collection, _
  i As Long, tStr As String, File As String
 ini.Filename = Filename
 SplitPath Filename, , , , File
 Set keys = ini.GetAllKeysFromSection(Section)
 For i = 1 To keys.Count
  If Len(tStr) = 0 Then
    tStr = LCase$(File) & "." & keys(i)(0) & "=" & keys(i)(1)
   Else
    tStr = tStr & vbCrLf & LCase$(File) & "." & keys(i)(0) & "=" & keys(i)(1)
  End If
 Next i
 Set ini = Nothing
 LastIncFile = File
 If Len(tStr) > 0 Then
  GetKeysAndValuesFromInifile = GetSortedText(tStr)
 End If
End Function

Private Function GetSortedText(txt As String) As String
 Dim tStrf() As String, coll As Collection, i As Long, j As Long, tStr As String
 GetSortedText = txt
 If InStr(1, txt, vbCrLf, vbTextCompare) Then
  tStrf = Split(txt, vbCrLf)
  Set coll = New Collection
  coll.Add tStrf(0)
  For i = 1 To UBound(tStrf)
   For j = 1 To coll.Count
    If tStrf(i) < coll(j) Then
     coll.Add tStrf(i), , j
     Exit For
    End If
   Next j
   If j > coll.Count Then
    coll.Add tStrf(i)
   End If
  Next i
  tStr = coll(1)
  For j = 2 To coll.Count
   tStr = tStr & vbCrLf & coll(j)
  Next j
  Set coll = Nothing
  GetSortedText = tStr
 End If
End Function

Private Sub CreateModTestpage(Filename As String, Str1 As String)
 Dim fn As Long, tStrf() As String, i As Long
 tStrf = Split(Str1, vbCrLf)
 fn = FreeFile
 
 Open Filename For Output As #fn
 
 Print #fn, "Attribute VB_Name = ""modTestpage"""
 Print #fn, "Option Explicit"
 Print #fn, ""
 Print #fn, "Public Function GetTestpage() As String"
 Print #fn, " Dim tStr As String"
 Print #fn, " tStr = """""
 For i = LBound(tStrf) To UBound(tStrf)
 Print #fn, " tStr = tStr & """ & tStrf(i) & """ & vbCr"
 Next i
 Print #fn, " GetTestpage = tStr"
 Print #fn, "End Function"
 
 Close #fn
End Sub

