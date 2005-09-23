VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Form1"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows-Standard
   WindowState     =   2  'Maximiert
   Begin MSComctlLib.ListView ListView1 
      Height          =   3900
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6879
      View            =   3
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const fcw = 500

Private WithEvents PDFCreator1 As clsPDFCreator
Attribute PDFCreator1.VB_VarHelpID = -1
Private pErr As clsPDFCreatorError

Private Sub Form_Load()
 Dim c As Collection, i As Long, l As ListItem
 Set pErr = New clsPDFCreatorError
 With ListView1
  .ColumnHeaders.Clear
  .ColumnHeaders.Add , "Dummy"
  .ColumnHeaders.Add , "Nr", "", , lvwColumnRight
  .ColumnHeaders.Add , "ON", "Option"
  .ColumnHeaders.Add , "O", "Value"
  .ColumnHeaders.Add , "SO", "Standard value"
 End With
 
 Set PDFCreator1 = New clsPDFCreator
 With PDFCreator1
  .cVisible = False
  If .cStart("/NoProcessingAtStartup") = True Then
    Set c = .cOptionsNames
    For i = 1 To c.Count
    Set l = ListView1.ListItems.Add()
     l.SubItems(1) = i
     l.SubItems(2) = c(i)
     l.SubItems(3) = PDFCreator1.cOption(c(i))
     l.SubItems(4) = PDFCreator1.cStandardOption(c(i))
    Next i
    .cClose
  End If
 End With
End Sub

Private Sub Form_Resize()
 Dim cw As Double
 With ListView1
  .Top = 0
  .Left = 0
  .Width = Me.ScaleWidth
  .Height = Me.ScaleHeight
  .ColumnHeaders(1).Width = 0
  .ColumnHeaders(2).Width = fcw
  cw = (.Width - fcw - 400) / 3
  .ColumnHeaders(3).Width = cw
  .ColumnHeaders(4).Width = cw
  .ColumnHeaders(5).Width = cw
 End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set PDFCreator1 = Nothing
End Sub

Private Sub PDFCreator1_eError()
 Set pErr = PDFCreator1.cError
 MsgBox pErr.Description & vbCrLf & vbCrLf & _
  "Program closes now."
 Unload Me
End Sub
