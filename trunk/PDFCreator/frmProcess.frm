VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProcess 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Erstelle PDF"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
   Icon            =   "frmProcess.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4290
   StartUpPosition =   2  'Bildschirmmitte
   Tag             =   "2160"
   Begin MSComCtl2.Animation anmProcess 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2143
      _Version        =   393216
      AutoPlay        =   -1  'True
      FullWidth       =   273
      FullHeight      =   81
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Zentriert
      Caption         =   "Erstelle PDF..."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   4095
   End
End
Attribute VB_Name = "frmProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias _
    "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
    lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
    hwndCallback As Long) As Long

Private Sub Form_Load()
anmProcess.Open App.Path & "\filemove.avi"
End Sub
