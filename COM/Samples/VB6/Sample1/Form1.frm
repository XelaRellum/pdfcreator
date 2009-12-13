VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Form1"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox cmbCountOfPages 
      Height          =   315
      Left            =   1260
      Style           =   2  'Dropdown-Liste
      TabIndex        =   6
      Top             =   2625
      Width           =   750
   End
   Begin VB.TextBox txtFilename 
      Height          =   330
      Left            =   105
      TabIndex        =   4
      Top             =   2100
      Width           =   6315
   End
   Begin VB.TextBox txtStatus 
      BackColor       =   &H00C0FFFF&
      Height          =   1485
      Left            =   105
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   3
      Top             =   3465
      Width           =   6315
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   5250
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   2
      Top             =   1155
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create PDF"
      Height          =   435
      Left            =   5145
      TabIndex        =   1
      Top             =   2625
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
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      Caption         =   "Count of pages:"
      Height          =   195
      Left            =   105
      TabIndex        =   7
      Top             =   2625
      Width           =   1125
   End
   Begin VB.Line Line1 
      X1              =   105
      X2              =   6405
      Y1              =   3255
      Y2              =   3255
   End
   Begin VB.Label lblFilename 
      AutoSize        =   -1  'True
      Caption         =   "Filename"
      Height          =   195
      Left            =   105
      TabIndex        =   5
      Top             =   1890
      Width           =   630
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

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private WithEvents PDFCreator1 As PDFCreator.clsPDFCreator
Attribute PDFCreator1.VB_VarHelpID = -1
Private pErr As clsPDFCreatorError, opt As clsPDFCreatorOptions
Private noStart As Boolean, fac As Double, StartTime As Date

Private Sub Form_Load()
 Dim i As Long
 noStart = True
 fac = Sqr(2#) / 2#
 With cmbCountOfPages
  For i = 1 To 100
   .AddItem CStr(i)
  Next i
  .ListIndex = 0
  .Left = lblCount.Left + lblCount.Width + 100
  .Top = lblCount.Top + (lblCount.Height - .Height) / 2
 End With
 txtFilename.Text = CompletePath(App.Path) & App.EXEName & ".pdf"
 Set PDFCreator1 = New clsPDFCreator
 Set pErr = New clsPDFCreatorError
 With PDFCreator1
  .cVisible = True
  If .cStart("/NoProcessingAtStartup") = False Then
    If .cStart("/NoProcessingAtStartup", True) = False Then
     Command1.Enabled = False
     Exit Sub
    End If
    AddStatus "Use an existing running instance!"
    .cVisible = True
  End If
  ' Get the options
  Set opt = .cOptions
  .cClearCache
  Picture1.Picture = LoadResPicture(101, vbResBitmap)
  noStart = False
 End With
 AddStatus "Program started"
End Sub

Private Sub Form_Unload(Cancel As Integer)
 If noStart = False Then
  PDFCreator1.cClose
  While PDFCreator1.cProgramIsRunning
   DoEvents
   Sleep 100
  Wend
 End If
 Set PDFCreator1 = Nothing
 Set pErr = Nothing
 Set opt = Nothing
End Sub

Private Sub Command1_Click()
 Dim pic As IPictureDisp, sw As Long, sh As Long, r As Long, _
  Path As String, Filename As String, i As Long
 SplitPath txtFilename.Text, , Path, Filename
 Command1.Enabled = False
 AddStatus "Start creating pdf ..."
 With opt
  .AutosaveDirectory = Path
  .AutosaveFilename = Filename
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
  .Font.Size = 12
  .ForeColor = vbBlack
  For i = 1 To cmbCountOfPages.ListIndex
   .PaintPicture Picture1.Picture, .ScaleWidth - Picture1.ScaleWidth * 6.3 - 100, 100
   DrawCircles sw / 2#, 1.2 * sh / 2#, r, recDepth
   PrintTextOnPrinter Text1, 400, 1000
   .NewPage
  Next i
  .PaintPicture Picture1.Picture, .ScaleWidth - Picture1.ScaleWidth * 6.3 - 100, 100
  DrawCircles sw / 2#, 1.2 * sh / 2#, r, recDepth
  PrintTextOnPrinter Text1, 400, 1000
  .EndDoc
 End With
 PDFCreator1.cPrinterStop = False
 StartTime = Now
 Command1.Enabled = True
 Screen.MousePointer = vbHourglass
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

Private Sub AddStatus(Str1 As String)
 With txtStatus
  If LenB(.Text) = 0 Then
    .Text = Time & ": " & Str1
   Else
    .Text = .Text & vbCrLf & Time & ": " & Str1
  End If
  .SelStart = Len(.Text)
 End With
End Sub

Public Function CompletePath(Path As String) As String
 If Len(Path) = 0 Then
  Exit Function
 End If
 Path = Trim$(Path)
 If Right$(Path, 1) = "\" Then
   CompletePath = LTrim$(Path)
  Else
   CompletePath = LTrim$(Path) & "\"
 End If
End Function

Public Sub SplitPath(FullPath As String, Optional Drive As String, Optional Path As String, Optional Filename As String, Optional File As String, Optional Extension As String)
 Dim nPos As Integer
 nPos = InStrRev(FullPath, "\")
 If nPos > 0 Then
   If Left$(FullPath, 2) = "\\" Then
    If nPos = 2 Then
     Drive = FullPath: Path = vbNullString: Filename = vbNullString: File = vbNullString
     Extension = vbNullString
     Exit Sub
    End If
   End If
   Path = Left$(FullPath, nPos - 1)
   Filename = Mid$(FullPath, nPos + 1)
   nPos = InStrRev(Filename, ".")
   If nPos > 0 Then
     File = Left$(Filename, nPos - 1)
     Extension = Mid$(Filename, nPos + 1)
    Else
     File = Filename
     Extension = vbNullString
   End If
  Else
   nPos = InStrRev(FullPath, ":")
   If nPos > 0 Then
     Path = Mid(FullPath, 1, nPos - 1): Filename = Mid(FullPath, nPos + 1)
     nPos = InStrRev(Filename, ".")
     If nPos > 0 Then
       File = Left$(Filename, nPos - 1)
       Extension = Mid$(Filename, nPos + 1)
      Else
       File = Filename
       Extension = vbNullString
     End If
    Else
     Path = vbNullString: Filename = FullPath
     nPos = InStrRev(Filename, ".")
     If nPos > 0 Then
       File = Left$(Filename, nPos - 1)
       Extension = Mid$(Filename, nPos + 1)
      Else
       File = Filename
       Extension = vbNullString
     End If
   End If
 End If
 If Left$(Path, 2) = "\\" Then
   nPos = InStr(3, Path, "\")
   If nPos Then
     Drive = Left$(Path, nPos - 1)
    Else
     Drive = Path
   End If
  Else
   If Len(Path) = 2 Then
    If Right$(Path, 1) = ":" Then
     Path = Path & "\"
    End If
   End If
   If Mid$(Path, 2, 2) = ":\" Then
    Drive = Left$(Path, 2)
   End If
 End If
End Sub

Private Sub PDFCreator1_eReady()
 AddStatus """" & PDFCreator1.cOutputFilename & """ was created! (" & _
  DateDiff("s", StartTime, Now) & " seconds)"
 PDFCreator1.cPrinterStop = True
 Screen.MousePointer = vbNormal
End Sub

Private Sub PDFCreator1_eError()
 Set pErr = PDFCreator1.cError
 AddStatus "Error[" & pErr.Number & "]: " & pErr.Description
 Screen.MousePointer = vbNormal
End Sub
