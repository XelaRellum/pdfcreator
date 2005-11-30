VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Language Installer"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7995
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   3360
      TabIndex        =   7
      Text            =   "http://pdfcreator.wurzel6.de/translations/"
      Top             =   360
      Width           =   4335
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   3360
      TabIndex        =   6
      Top             =   2640
      Width           =   4335
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "Install"
      Height          =   735
      Left            =   3360
      TabIndex        =   4
      Top             =   3120
      Width           =   1935
   End
   Begin VB.ListBox lstFiles 
      Height          =   5715
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   3360
      TabIndex        =   1
      Text            =   "language_list.php"
      Top             =   960
      Width           =   4335
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "Get Languages"
      Height          =   735
      Left            =   3360
      TabIndex        =   0
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblOnlineFiles 
      Caption         =   "Online File Source:"
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lblPath 
      Caption         =   "PDFCreator Path:"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   2400
      Width           =   3615
   End
   Begin VB.Label lblSource 
      Caption         =   "Downloadlist Filename:"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source!

Option Explicit

Private Declare Function InternetOpen Lib "wininet" Alias _
        "InternetOpenA" (ByVal sAgent As String, ByVal _
        lAccessType As Long, ByVal sProxyName As String, ByVal _
        sProxyBypass As String, ByVal lFlags As Long) As Long
        
Private Declare Function InternetCloseHandle Lib "wininet" _
        (ByVal hInet As Long) As Integer
        
Private Declare Function InternetReadFile Lib "wininet" _
        (ByVal hFile As Long, ByVal sBuffer As String, ByVal _
        lNumBytesToRead As Long, lNumberOfBytesRead As Long) _
        As Integer
        
Private Declare Function InternetOpenUrl Lib "wininet" Alias _
        "InternetOpenUrlA" (ByVal hInternetSession As Long, _
        ByVal lpszUrl As String, ByVal lpszHeaders As String, _
        ByVal dwHeadersLength As Long, ByVal dwFlags As Long, _
        ByVal dwContext As Long) As Long


Const INTERNET_OPEN_TYPE_DIRECT = 1
Const INTERNET_OPEN_TYPE_PROXY = 3
Const INTERNET_FLAG_RELOAD = &H80000000

Const UserAgent = "PDFCreator"

Function GetFile(strFile As String) As String
    Dim l&, Buffer$, hOpen&, hFile&, Result&
    l = 50000
    Buffer = Space(l)

    hOpen = InternetOpen(UserAgent, INTERNET_OPEN_TYPE_DIRECT, _
                         vbNullString, vbNullString, 0)
  
    hFile = InternetOpenUrl(hOpen, strFile, vbNullString, _
                            ByVal 0&, INTERNET_FLAG_RELOAD, _
                            ByVal 0&)
                            
    Call InternetReadFile(hFile, Buffer, l, Result&)
    Call InternetCloseHandle(hFile)
    Call InternetCloseHandle(hOpen)
    
    GetFile = Left$(Buffer, Result)
End Function

Private Sub cmdGet_Click()
  Dim strLanguages() As String
  Dim i As Integer

    MousePointer = vbHourglass
    
    txtURL.Text = CompletePath(txtURL.Text, "/")
    strLanguages = Split(GetFile(txtURL.Text & txtFile), vbLf)
    lstFiles.Clear
    
    For i = LBound(strLanguages) To UBound(strLanguages)
        If ((strLanguages(i) <> vbNullString) And (InStr(1, strLanguages(i), ":"))) Then lstFiles.AddItem (strLanguages(i))
    Next i
    
    MousePointer = vbDefault
End Sub

Private Sub cmdInstall_Click()

  Dim strLangFile As String, strPath As String, strFile As String
  Dim strLang() As String
   
  If lstFiles.Text = vbNullString Then Exit Sub

  txtURL.Text = CompletePath(txtURL.Text, "/")
  strLang = Split(lstFiles.Text, ":")

  MousePointer = vbHourglass
  strLangFile = GetFile(txtURL.Text & strLang(1) & "/" & strLang(0))
  If InStr(1, strLangFile, "[Common]", vbTextCompare) = 0 Then
    MsgBox "Error: An Invalid File was downloaded", vbCritical
    MousePointer = vbDefault
    Exit Sub
  End If
  
  strPath = CompletePath(txtPath.Text, "\")
  
  If strLangFile = vbNullString Then
    MsgBox "Could not retrieve Translation file!", vbCritical
    MousePointer = vbDefault
    Exit Sub
  End If
  
  If Not DirExists(strPath) Then
    MsgBox "Destination Path does not exist!", vbCritical
    MousePointer = vbDefault
    Exit Sub
  End If


  strFile = strPath & "Languages\" & strLang(0)

  If FileExists(strFile) Then
    If MsgBox("Destination File already exists. Overwrite?", vbYesNo) = vbNo Then
      MousePointer = vbDefault
      Exit Sub
    End If
  End If
  
  Open strFile For Output As #1
    Print #1, strLangFile
  Close #1
  
  MsgBox "File succesfully installed"

  MousePointer = vbDefault
End Sub

Private Function DirExists(DirStr As String) As Boolean
 On Error GoTo ErrorHandler
 DirExists = GetAttr(DirStr) Or vbDirectory
 Exit Function
ErrorHandler:
 DirExists = False
End Function

Private Function FileExists(FileStr As String) As Boolean
 On Error GoTo ErrorHandler
 FileExists = GetAttr(FileStr)
 Exit Function
ErrorHandler:
 FileExists = False
End Function

Private Function CompletePath(strPath As String, Trailer As String) As String
  If (Right$(strPath, 1) <> Trailer) Then strPath = strPath & Trailer
  CompletePath = strPath
End Function

Private Sub Form_Load()
  txtPath = App.Path
End Sub
