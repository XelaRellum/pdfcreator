VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Form1"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   5400
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   3
      Top             =   1890
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create PDF"
      Height          =   435
      Left            =   105
      TabIndex        =   1
      Top             =   1890
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   105
      Width           =   6300
   End
   Begin VB.Label lbl 
      Caption         =   "Status: Program was started."
      ForeColor       =   &H000000C0&
      Height          =   555
      Left            =   105
      TabIndex        =   2
      Top             =   2415
      Width           =   6345
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const recDepth = 3

Private Const EM_FMTLINES = &HC8

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
 (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private WithEvents PDFCreator1 As PDFCreator.clsPDFCreator
Attribute PDFCreator1.VB_VarHelpID = -1
Private pErr As clsPDFCreatorError, noStart As Boolean, fac As Double, _
 opt As clsPDFCreatorOptions

Private Sub Form_Load()
 noStart = True
 fac = Sqr(2#) / 2#
 Set PDFCreator1 = New clsPDFCreator
 Set pErr = New clsPDFCreatorError
 With PDFCreator1
  .cVisible = True
  If .cStart("/NoProcessingAtStartup") = False Then
    If .cStart("/NoProcessingAtStartup", True) = False Then
     Command1.Enabled = False
     Exit Sub
    End If
    lbl.Caption = "Status: Use an existing running instance!"
    .cVisible = True
  End If
  ' Get the options
  Set opt = .cOptions
  .cClearCache
  Picture1.Picture = LoadResPicture(101, vbResBitmap)
  noStart = False
 End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
 If noStart = False Then
  DoEvents
  PDFCreator1.cClose
 End If
 DoEvents
 Set PDFCreator1 = Nothing
 Set pErr = Nothing
 Set opt = Nothing
End Sub

Private Sub Command1_Click()
 Dim pic As IPictureDisp, sw As Long, sh As Long, r As Long
 
 Command1.Enabled = False
 lbl.Caption = "Status: Start creating pdf ..."
 With opt
  .AutosaveDirectory = App.Path
  .AutosaveFilename = App.EXEName
  .UseAutosave = 1
  .UseAutosaveDirectory = 1
  .AutosaveFormat = 0 ' PDF
 End With
 Set PDFCreator1.cOptions = opt
 
 Set Printer = Printers(PrinterIndex("PDFCreator"))
 With Printer
  .ScaleMode = vbPixels
  sw = .ScaleWidth
  sh = .ScaleHeight
  r = (0.8 * sw) / 2#
  .PrintQuality = 150
  .PaintPicture Picture1.Picture, .ScaleWidth - Picture1.ScaleWidth * 6.3 - 100, 100
  .Font.Size = 12
  DrawCircles sw / 2#, 1.2 * sh / 2#, r, recDepth
  .ForeColor = vbBlack
  PrintTextOnPrinter Text1, 400, 1000
  .EndDoc
 End With
 PDFCreator1.cPrinterStop = False
 Command1.Enabled = True
 ' You can't restore the options here, because the printjob isn't ready!
End Sub

Private Function PrinterIndex(Printername As String) As Long
 Dim i As Long
' Show all printers
' Debug.Print "Printers [" & Printers.Count & "]:"
' For i = 0 To Printers.Count - 1
'  Debug.Print Printers(i).DeviceName
' Next i
 For i = 0 To Printers.Count - 1
  If UCase(Printers(i).DeviceName) = UCase$(Printername) Then
   PrinterIndex = i
   Exit For
  End If
 Next i
End Function

Private Sub PrintTextOnPrinter(txt As TextBox, _
 Optional xPos As Long = 0, Optional yPos As Long = 0)
 Dim tStr As String, tStrf() As String, i As Long
 tStr = TranslateSoftbreaksInHardbreaks(Text1)
 If LenB(tStr) = 0 Then
  Exit Sub
 End If
 Printer.CurrentX = xPos
 Printer.CurrentY = yPos
 If InStr(1, tStr, vbCrLf, vbTextCompare) > 0 Then
   tStrf = Split(tStr, vbCrLf)
   For i = LBound(tStrf) To UBound(tStrf)
    Printer.Print tStrf(i)
    Printer.CurrentX = xPos
   Next i
  Else
   Printer.Print tStr
 End If
End Sub

Private Function TranslateSoftbreaksInHardbreaks(txt As TextBox) As String
 Call SendMessage(txt.hwnd, EM_FMTLINES, -1, 0&)
 TranslateSoftbreaksInHardbreaks = Replace$(txt.Text, Chr$(13) & Chr$(13) & Chr$(10), vbCrLf)
 Call SendMessage(txt.hwnd, EM_FMTLINES, 0, 0&)
End Function

Private Sub DrawCircles(xm As Long, ym As Long, r As Long, Optional rec As Long = 1)
 Dim col As Long
 col = vbRed
 Printer.Circle (xm, ym), r, col
 Printer.Circle (xm - r / 2, ym), r / 2, col
 Printer.Circle (xm + r / 2, ym), r / 2, col
 Printer.Circle (xm, ym - r / 2), r / 2, col
 Printer.Circle (xm, ym + r / 2), r / 2, col
 Printer.Circle (xm, ym), r / 2, col
 Printer.Circle (xm - fac * r / 2, ym - fac * r / 2), r / 2, col
 Printer.Circle (xm + fac * r / 2, ym - fac * r / 2), r / 2, col
 Printer.Circle (xm - fac * r / 2, ym + fac * r / 2), r / 2, col
 Printer.Circle (xm + fac * r / 2, ym + fac * r / 2), r / 2, col
 If rec = 1 Then
   Exit Sub
  Else
   DrawCircles xm - r / 2, ym, r / 2, rec - 1
   DrawCircles xm + r / 2, ym, r / 2, rec - 1
   DrawCircles xm, ym - r / 2, r / 2, rec - 1
   DrawCircles xm, ym + r / 2, r / 2, rec - 1
   DrawCircles xm, ym, r / 2, rec - 1
   DrawCircles xm - fac * r / 2, ym - fac * r / 2, r / 2, rec - 1
   DrawCircles xm + fac * r / 2, ym - fac * r / 2, r / 2, rec - 1
   DrawCircles xm - fac * r / 2, ym + fac * r / 2, r / 2, rec - 1
   DrawCircles xm + fac * r / 2, ym + fac * r / 2, r / 2, rec - 1
 End If
End Sub

Private Sub PDFCreator1_eReady()
 lbl.Caption = "Status: """ & PDFCreator1.cOutputFilename & """ was created!"
 PDFCreator1.cPrinterStop = True
End Sub

Private Sub PDFCreator1_eError()
 Set pErr = PDFCreator1.cError
 lbl.Caption = "Status: ""Error[" & pErr.Number & "]: " & pErr.Description
End Sub
