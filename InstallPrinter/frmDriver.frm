VERSION 5.00
Begin VB.Form frmDriver 
   Caption         =   "Form1"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Löschen"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2880
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Driver Path"
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton frmDriver 
      Caption         =   "Installieren"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   1575
   End
End
Attribute VB_Name = "frmDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim lngDriverDirectoryLevel    As Long
    Dim lngDriverDirectoryNeeded   As Long
    Dim bytDriverDirectoryBuffer() As Byte
    Dim strDriverDirectory         As String * 512
    Dim lngWin32apiResultCode      As Long

    lngDriverDirectoryLevel = 1
    lngWin32apiResultCode = GetPrinterDriverDirectory(vbNullString, vbNullString, lngDriverDirectoryLevel, ByVal vbNullString, 0, lngDriverDirectoryNeeded)
    ReDim bytDriverDirectoryBuffer(lngDriverDirectoryNeeded - 1)
    lngWin32apiResultCode = GetPrinterDriverDirectory(vbNullString, vbNullString, lngDriverDirectoryLevel, bytDriverDirectoryBuffer(0), lngDriverDirectoryNeeded, lngDriverDirectoryNeeded)
    lngWin32apiResultCode = lstrcpy(ByVal strDriverDirectory, bytDriverDirectoryBuffer(0))

    Text1.Text = Left(strDriverDirectory, InStr(strDriverDirectory, vbNullChar) - 1)
End Sub

Private Sub Command2_Click()
Dim success
success = DeletePrinterDriver(vbNullString, vbNullString, Text2.Text)

If success = 0 Then
  RaiseAPIError
Else
  MsgBox Text2.Text & " wurde erfolgreich gelöscht"
End If
End Sub

Private Sub frmDriver_Click()
AddDriver
End Sub
