VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "TransTool"
   ClientHeight    =   7980
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6300
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame fraAbout 
      Height          =   1275
      Left            =   0
      TabIndex        =   5
      Top             =   6090
      Width           =   6255
      Begin VB.Label lbl 
         Caption         =   "ProgramVersion"
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   6
         Top             =   210
         Width           =   5055
      End
      Begin VB.Label lbl 
         Caption         =   "License"
         Height          =   255
         Index           =   1
         Left            =   105
         TabIndex        =   7
         Top             =   450
         Width           =   5055
      End
      Begin VB.Label lbl 
         Caption         =   "Author"
         Height          =   255
         Index           =   2
         Left            =   105
         TabIndex        =   8
         Top             =   690
         Width           =   4935
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "HomepageSourceforge"
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   4
         Left            =   4425
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   10
         Top             =   930
         Width           =   1650
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Homepage"
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   3
         Left            =   105
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   9
         Top             =   930
         Width           =   780
      End
      Begin VB.Image ImgPaypal 
         Height          =   465
         Left            =   5265
         Picture         =   "frmAbout.frx":548A
         Top             =   210
         Width           =   930
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Close"
      Height          =   450
      Left            =   4920
      TabIndex        =   0
      Top             =   7455
      Width           =   1260
   End
   Begin VB.PictureBox picBackground 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1080
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   1
      Top             =   7410
      Width           =   735
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   2
      Top             =   7410
      Width           =   735
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   2625
      Top             =   7455
   End
   Begin VB.PictureBox picAbout 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   120
      Picture         =   "frmAbout.frx":57EF
      ScaleHeight     =   4500
      ScaleWidth      =   6000
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2100
      Top             =   7455
   End
   Begin VB.Frame fraDescription 
      Caption         =   "Description"
      Height          =   1215
      Left            =   0
      TabIndex        =   4
      Top             =   4800
      Width           =   6255
      Begin VB.PictureBox picDescription 
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   105
         ScaleHeight     =   960
         ScaleWidth      =   6105
         TabIndex        =   11
         Top             =   210
         Width           =   6105
         Begin VB.TextBox txtDescription 
            Alignment       =   2  'Zentriert
            Appearance      =   0  '2D
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'Kein
            Height          =   735
            Left            =   0
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   0
            Width           =   5775
         End
      End
   End
   Begin VB.Shape shpRec 
      BorderColor     =   &H80000010&
      Height          =   4695
      Left            =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const border = 50, picBorder = 5

Private sCol1 As OLE_COLOR, sCol2 As OLE_COLOR, AboutText As Collection, _
 yOffs As Long, fontColl As Collection

Private Sub cmd_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAbout", "cmd_Click")
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
50010  lbl(0).Caption = App.ProductName & " " & App.Major & "." & App.Minor & "." & App.Revision
50020  lbl(1).Caption = "License: GNU GENERAL PUBLIC LICENSE"
50030  lbl(2).Caption = "Author: Frank Heindörfer, Philip Chinery (c) 2004"
50040  lbl(3).Caption = "http://www.pdfcreator.de.vu"
50050  lbl(3).DragIcon = LoadResPicture(2001, vbResCursor)
50060  lbl(4).Caption = "http://www.sf.net/projects/pdfcreator"
50070  lbl(4).DragIcon = LoadResPicture(2001, vbResCursor)
50080  txtDescription.Text = "TransTool is a part from PDFCreator. With TransTool you can create and edit language files for PDFCreator. Language files are normal ini files."
50090  sCol1 = lbl(3).ForeColor
50100  sCol2 = lbl(4).ForeColor
50110
50120  With picAbout
50130   .AutoRedraw = True
50140   .ScaleMode = vbPixels
50150
50160   .Visible = True
50170  End With
50180  SetAboutText
50190  With picBuffer
50200   .Width = picAbout.Width
50210   .Height = picAbout.Height
50220   .ScaleMode = vbPixels
50230   .AutoRedraw = True
50240   .Visible = False
50250   yOffs = .Height \ Screen.TwipsPerPixelY
50260  End With
50270  With picBackground
50280   .Width = picAbout.Width
50290   .Height = picAbout.Height
50300   .ScaleMode = vbPixels
50310   .AutoRedraw = True
50320   .Visible = False
50330   Set .Picture = picAbout.Picture
50340  End With
50350
50360  Timer1.Enabled = True
50370  Timer2.Interval = 50
50380  Timer2.Enabled = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAbout", "Form_Load")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Resize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With shpRec
50020   .Top = border
50030   .Left = border
50040   .Width = picAbout.Width + 2 * border + 50
50050   .Height = picAbout.Height + 2 * border + 20
50060  End With
50070  picAbout.Top = shpRec.Top + border
50080  Me.Width = shpRec.Width + 2 * border + 100
50090  With fraDescription
50100   .Top = shpRec.Top + shpRec.Height + border
50110   .Left = shpRec.Left
50120   .Width = shpRec.Width
50130   .Height = 800
50140  End With
50150  With picDescription
50160   .Top = 250
50170   .Left = border
50180   .Width = fraDescription.Width - 100
50190   .Height = fraDescription.Height - 300
50200  End With
50210  With txtDescription
50220   .Top = 0
50230   .Left = 0
50240   .Width = picDescription.Width
50250   .Height = picDescription.Height
50260  End With
50270  With fraAbout
50280   .Top = fraDescription.Top + fraDescription.Height + border
50290   .Left = fraDescription.Left
50300   .Width = fraDescription.Width
50310  End With
50320
50330  lbl(0).Top = 200
50340  lbl(0).Left = 70
50350  lbl(1).Top = lbl(0).Top + lbl(0).Height
50360  lbl(1).Left = lbl(0).Left
50370  lbl(2).Top = lbl(1).Top + lbl(1).Height
50380  lbl(2).Left = lbl(1).Left
50390  lbl(3).Top = lbl(2).Top + lbl(2).Height
50400  lbl(3).Left = lbl(2).Left
50410  lbl(4).Top = lbl(3).Top
50420  lbl(4).Left = fraAbout.Width - lbl(4).Width - lbl(0).Left
50430  cmd.Top = fraAbout.Top + fraAbout.Height + 100
50440  Me.Height = cmd.Top + cmd.Height + (Me.Height - Me.ScaleHeight) + 100
50450  imgPaypal.Top = lbl(0).Top
50460  imgPaypal.Left = fraAbout.Width - imgPaypal.Width - lbl(0).Left
50470  Timer2.Interval = 40
50480  Timer2.Enabled = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAbout", "Form_Resize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Unload(Cancel As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim dc As Long, rct As Rect
50020  dc = GetWindowDC(Me.hwnd)
50030  GetWindowRect Me.hwnd, rct
50040  frmMain.picAbout.Width = (rct.Right - rct.Left) * Screen.TwipsPerPixelX
50050  frmMain.picAbout.Height = (rct.Bottom - rct.Top) * Screen.TwipsPerPixelY
50060
50070  With frmMain.picAbout
50080   .Cls
50090   BitBlt .hdc, 0, 0, .Width, .Height, dc, 0, 0, vbSrcCopy
50100  End With
50110
50120  ReleaseDC Me.hwnd, dc
50130  frmMain.LastAboutLeft = Me.Left
50140  frmMain.LastAboutTop = Me.Top
50150  frmAboutPicLT.Show
50160  frmAboutPicRT.Show
50170  frmAboutPicLB.Show
50180  frmAboutPicRB.Show
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAbout", "Form_Unload")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ImgPaypal_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  OpenDocument Paypal
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAbout", "ImgPaypal_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ImgPaypal_DblClick()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  OpenDocument Paypal
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAbout", "ImgPaypal_DblClick")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub lbl_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case Index
        Case 3, 4
50030    Call ShellExecute(Me.hwnd, "Open", lbl(Index).Caption, "", App.Path, 1)
50040  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAbout", "lbl_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub lbl_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case Index
        Case 3, 4
50030    If Source Is lbl(Index) Then
50040     With lbl(Index)
50050      Call ShellExecute(Me.hwnd, "Open", lbl(Index).Caption, "", App.Path, 1)
50060 '     .Font.Underline = False
50070      .ForeColor = &H808000
50080      .Tag = 1
50090     End With
50100    End If
50110  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAbout", "lbl_DragDrop")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub lbl_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case Index
        Case 3, 4
50030    If State = vbLeave Then
50040     With lbl(Index)
50050      .Drag vbEndDrag
50060      If .Tag = "1" Then
50070        .Font.Underline = True
50080        .ForeColor = &H808000
50090       Else
50100        .Font.Underline = False
50110        .ForeColor = &H8000&
50120      End If
50130     End With
50140    End If
50150  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAbout", "lbl_DragOver")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case Index
        Case 3, 4
50030    With lbl(Index)
50040      If .Tag = "1" Then
50050        .Font.Underline = True
50060        .ForeColor = &H808000
50070       Else
50080        .ForeColor = vbBlue
50090        .Font.Underline = True
50100      End If
50110     .Drag vbBeginDrag
50120    End With
50130  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAbout", "lbl_MouseMove")
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
50010  Timer1.Enabled = False
50020  MoveMouseToCommandButton cmd
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAbout", "Timer1_Timer")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Timer2_Timer()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, l As Long, tL As Long, tLine As clsTextLine
50020
50030 ' On Error Resume Next
50040  Timer2.Enabled = False
50050  With picBuffer
50060
50070 '  Set .Picture = LoadResPicture(1001, vbResBitmap)
50080   Call BitBlt(.hdc, 0, .ScaleTop, .ScaleWidth, _
   .ScaleHeight, picBackground.hdc, 0, 0, vbSrcCopy)
50100   .Refresh
50110
50120   ' Shaddow
50130   tL = 0
50140   For i = 1 To AboutText.Count
50150    Set tLine = AboutText(i)
50160    Set .Font = tLine.Font
50170    .ForeColor = RGB(32, 32, 32)
50180    .CurrentY = yOffs + tL + 4
50190    .CurrentX = (.ScaleWidth / 2) - (.TextWidth(tLine.Text) / 2) + 4
50200    picBuffer.Print tLine.Text
50210    tL = tL + tLine.Font.Size + 12
50220   Next i
50230   ' Text
50240   tL = 0
50250   For i = 1 To AboutText.Count
50260    Set tLine = AboutText(i)
50270    Set .Font = tLine.Font
50280    .ForeColor = tLine.ForeColor
50290    .CurrentY = yOffs + tL
50300    .CurrentX = (.ScaleWidth / 2) - (.TextWidth(tLine.Text) / 2)
50310    picBuffer.Print tLine.Text
50320    tL = tL + tLine.Font.Size + 12
50330   Next i
50340   Call BitBlt(picAbout.hdc, picBorder, picAbout.ScaleTop + picBorder, picAbout.ScaleWidth - 2 * picBorder, _
   picAbout.ScaleHeight - 2 * picBorder, .hdc, picBorder, picBorder, vbSrcCopy)
50360   picAbout.Refresh
50370
50380   yOffs = yOffs - 1
50390   If .CurrentY < 0 Then
50400    yOffs = .Height \ Screen.TwipsPerPixelY
50410   End If
50420  End With
50430  Timer2.Enabled = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAbout", "Timer2_Timer")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ChangeFont(Index As Long, f As StdFont)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  fontColl.Remove Index
50020  fontColl.Add f, , Index
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAbout", "ChangeFont")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function GetFont(ctrl As Control) As StdFont
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Set GetFont = New StdFont
50020  With ctrl.Font
50030   GetFont.Bold = .Bold
50040   GetFont.Charset = .Charset
50050   GetFont.Italic = .Italic
50060   GetFont.Name = .Name
50070   GetFont.Size = .Size
50080   GetFont.Strikethrough = .Strikethrough
50090   GetFont.Underline = .Underline
50100   GetFont.Weight = .Weight
50110  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAbout", "GetFont")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

 
Private Sub SetAboutText()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim f As StdFont, emptyLineHeight As Long, paragraphHeight As Long, _
  tLine As clsTextLine
50030
50040  Set AboutText = New Collection
50050  Set tLine = New clsTextLine
50060  With tLine
50070   .ForeColor = vbYellow: .Font.Size = 24: .Font.Underline = True: .Font.Charset = 0
50080   .Text = "TransTool"
50090  End With
50100  AboutText.Add tLine
50110  Set tLine = New clsTextLine
50120  With tLine
50130   .ForeColor = vbYellow: .Font.Size = 8: .Font.Charset = 0
50140   .Text = "(A part of PDFCreator)"
50150  End With
50160  AboutText.Add tLine
50170  Set tLine = New clsTextLine
50180  With tLine
50190   .Font.Size = 12
50200   .Text = " "
50210  End With
50220  AboutText.Add tLine
50230  Set tLine = New clsTextLine
50240  With tLine
50250   .ForeColor = vbYellow: .Font.Size = 16: .Font.Underline = True: .Font.Charset = 0
50260   .Text = "Authors"
50270  End With
50280  AboutText.Add tLine
50290  Set tLine = New clsTextLine
50300  With tLine
50310   .ForeColor = vbYellow: .Font.Size = 12: .Font.Bold = True: .Font.Italic = True: .Font.Charset = 0
50320   .Text = "Frank Heindörfer"
50330  End With
50340  AboutText.Add tLine
50350  Set tLine = New clsTextLine
50360  With tLine
50370   .ForeColor = vbYellow: .Font.Size = 12: .Font.Bold = True: .Font.Italic = True: .Font.Charset = 0
50380   .Text = "Philip Chinery"
50390  End With
50400  AboutText.Add tLine
50410  Set tLine = New clsTextLine
50420  With tLine
50430   .Font.Size = 12
50440   .Text = " "
50450  End With
50460  AboutText.Add tLine
50470  Set tLine = New clsTextLine
50480  With tLine
50490   .ForeColor = vbYellow: .Font.Size = 16: .Font.Underline = True: .Font.Charset = 0
50500   .Text = "Our special thanks to:"
50510  End With
50520  AboutText.Add tLine
50530  Set tLine = New clsTextLine
50540  With tLine
50550   .ForeColor = vbYellow: .Font.Size = 12: .Font.Bold = True: .Font.Italic = True: .Font.Charset = 0
50560   .Text = "All sponsors"
50570  End With
50580  AboutText.Add tLine
50590  Set tLine = New clsTextLine
50600  With tLine
50610   .ForeColor = vbYellow: .Font.Size = 12: .Font.Bold = True: .Font.Italic = True: .Font.Charset = 0
50620   .Text = "All beta testers"
50630  End With
50640  AboutText.Add tLine
50650  Set tLine = New clsTextLine
50660  With tLine
50670   .ForeColor = vbYellow: .Font.Size = 12: .Font.Bold = True: .Font.Italic = True: .Font.Charset = 0
50680   .Text = "All authors of the languages config files"
50690  End With
50700  AboutText.Add tLine
50710  Set tLine = New clsTextLine
50720  With tLine
50730   .Font.Size = 6
50740   .Text = " "
50750  End With
50760  AboutText.Add tLine
50770  Set tLine = New clsTextLine
50780  With tLine
50790   .ForeColor = vbYellow: .Font.Size = 16: .Font.Underline = True: .Font.Charset = 0
50800   .Text = "For hints and tips"
50810  End With
50820  AboutText.Add tLine
50830  Set tLine = New clsTextLine
50840  With tLine
50850   .Font.Size = 6
50860   .Text = " "
50870  End With
50880  AboutText.Add tLine
50890  Set tLine = New clsTextLine
50900  With tLine
50910   .ForeColor = vbYellow: .Font.Size = 12: .Font.Underline = True: .Font.Bold = True: .Font.Charset = 0
50920   .Text = "Visual Basic"
50930  End With
50940  AboutText.Add tLine
50950  Set tLine = New clsTextLine
50960  With tLine
50970   .ForeColor = vbYellow: .Font.Size = 12: .Font.Bold = True: .Font.Italic = True: .Font.Charset = 0
50980   .Text = "www.aboutvb.de"
50990  End With
51000  AboutText.Add tLine
51010  Set tLine = New clsTextLine
51020  With tLine
51030   .ForeColor = vbYellow: .Font.Size = 12: .Font.Bold = True: .Font.Italic = True: .Font.Charset = 0
51040   .Text = "www.activevb.de"
51050  End With
51060  AboutText.Add tLine
51070  Set tLine = New clsTextLine
51080  With tLine
51090   .ForeColor = vbYellow: .Font.Size = 12: .Font.Bold = True: .Font.Italic = True: .Font.Charset = 0
51100   .Text = "www.vb-hellfire.de"
51110  End With
51120  AboutText.Add tLine
51130  Set tLine = New clsTextLine
51140  With tLine
51150   .ForeColor = vbYellow: .Font.Size = 12: .Font.Bold = True: .Font.Italic = True: .Font.Charset = 0
51160   .Text = "www.mvps.org/vbnet"
51170  End With
51180  AboutText.Add tLine
51190  Set tLine = New clsTextLine
51200  With tLine
51210   .ForeColor = vbYellow: .Font.Size = 12: .Font.Bold = True: .Font.Italic = True: .Font.Charset = 0
51220   .Text = "www.planet-source-code.com"
51230  End With
51240  AboutText.Add tLine
51250  Set tLine = New clsTextLine
51260  With tLine
51270   .ForeColor = vbYellow: .Font.Size = 12: .Font.Underline = True: .Font.Bold = True: .Font.Charset = 0
51280   .Text = "Ghostscript"
51290  End With
51300  AboutText.Add tLine
51310  Set tLine = New clsTextLine
51320  With tLine
51330   .ForeColor = vbYellow: .Font.Size = 12: .Font.Bold = True: .Font.Italic = True: .Font.Charset = 0
51340   .Text = "www.ghostscript.com"
51350  End With
51360  AboutText.Add tLine
51370  Set tLine = New clsTextLine
51380  With tLine
51390   .ForeColor = vbYellow: .Font.Size = 12: .Font.Underline = True: .Font.Bold = True: .Font.Charset = 0
51400   .Text = "Innosetup"
51410  End With
51420  AboutText.Add tLine
51430  Set tLine = New clsTextLine
51440  With tLine
51450   .ForeColor = vbYellow: .Font.Size = 12: .Font.Bold = True: .Font.Italic = True: .Font.Charset = 0
51460   .Text = "www.innosetup.org"
51470  End With
51480  AboutText.Add tLine
51490  Set tLine = New clsTextLine
51500  With tLine
51510   .Font.Size = 16
51520   .Text = " "
51530  End With
51540  AboutText.Add tLine
51550  Set tLine = New clsTextLine
51560  With tLine
51570   .ForeColor = vbYellow: .Font.Size = 16: .Font.Underline = True: .Font.Bold = True: .Font.Charset = 0
51580   .Text = "And all other users they are help us."
51590  End With
51600  AboutText.Add tLine
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAbout", "SetAboutText")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetStandardFont()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim f As StdFont, i As Long
50020  Set f = GetFont(picAbout)
50030  Set fontColl = New Collection
50040  For i = 1 To AboutText.Count
50050   fontColl.Add f
50060  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAbout", "SetStandardFont")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
