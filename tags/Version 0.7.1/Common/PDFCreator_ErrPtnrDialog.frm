VERSION 5.00
Begin VB.Form ErrPtnr 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Attention ! An error has been occured."
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   ControlBox      =   0   'False
   Icon            =   "PDFCreator_ErrPtnrDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   Begin VB.Frame framFehler 
      Caption         =   "Error description:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1695
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   7215
      Begin VB.Label lblErrTitel 
         Caption         =   "Error-Nr:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblErrTitel 
         Caption         =   "Modul:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblErrTitel 
         Caption         =   "Procedure:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   8
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblErrTitel 
         BackStyle       =   0  'Transparent
         Caption         =   "Line:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   10
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblError 
         Caption         =   "lblError(0)"
         Height          =   675
         Index           =   0
         Left            =   1920
         TabIndex        =   5
         Top             =   240
         Width           =   5205
      End
      Begin VB.Label lblError 
         AutoSize        =   -1  'True
         Caption         =   "lblError(1)"
         Height          =   195
         Index           =   1
         Left            =   1920
         TabIndex        =   7
         Top             =   960
         Width           =   660
      End
      Begin VB.Label lblError 
         AutoSize        =   -1  'True
         Caption         =   "lblError(2)"
         Height          =   195
         Index           =   2
         Left            =   1920
         TabIndex        =   9
         Top             =   1200
         Width           =   660
      End
      Begin VB.Label lblError 
         AutoSize        =   -1  'True
         Caption         =   "lblError(3)"
         Height          =   195
         Index           =   3
         Left            =   1920
         TabIndex        =   11
         Top             =   1440
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fest Einfach
         Caption         =   " ! "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   900
         Left            =   90
         TabIndex        =   3
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame framProgInfo 
      Caption         =   "Program-Information:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   2640
      Width           =   7215
      Begin VB.Label lblProgInfo 
         Alignment       =   2  'Zentriert
         Height          =   195
         Left            =   30
         TabIndex        =   13
         Top             =   240
         Width           =   7110
      End
   End
   Begin VB.Frame framProtocol 
      Caption         =   "Error-Protocol ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   0
      TabIndex        =   14
      Top             =   3240
      Width           =   7215
      Begin VB.CommandButton cmdProtocol 
         Caption         =   "&WWW"
         Height          =   495
         Index           =   3
         Left            =   5640
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdProtocol 
         Caption         =   "&Save"
         Height          =   495
         Index           =   1
         Left            =   2040
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdProtocol 
         Caption         =   "&EMail"
         Height          =   495
         Index           =   2
         Left            =   3840
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdProtocol 
         Caption         =   "&Print"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame framCallStack 
      Caption         =   "CallStack:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   0
      TabIndex        =   19
      Top             =   4080
      Width           =   7215
      Begin VB.ListBox lstCallStack 
         Height          =   675
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   6975
      End
   End
   Begin VB.Frame framContinue 
      Caption         =   "Next Programstep..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   0
      TabIndex        =   21
      Top             =   5040
      Width           =   7215
      Begin VB.CommandButton cmdContinue 
         Caption         =   "&Repeat Command"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdContinue 
         Caption         =   "&Next Command"
         Height          =   495
         Index           =   1
         Left            =   2040
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdContinue 
         Caption         =   "Exit &Procedure"
         Height          =   495
         Index           =   2
         Left            =   3840
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdContinue 
         Caption         =   "E&xit Program"
         Height          =   495
         Index           =   3
         Left            =   5640
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame framComment 
      Caption         =   "Comment:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.Label lblComment 
         Caption         =   "lblComment"
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "ErrPtnr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The following code block includes the dialog settings
' which are modified ONLY by the ErrorPartner wizard.
'
'-- START Settings --- Do not modify THIS line!
Const ErrPtnr_LogFile% = 1
Const ErrPtnr_EMail$ = "thesmilyface@users.sourceforge.net"
Const ErrPtnr_WWW$ = "www.sourceforge.net/projects/pdfcreator"
Const ErrPtnr_Comment$ = "An error has occured. Look for technical support in the forums (WWW). " + vbCrLf + "Further advices and updates you can find also in the internet (WWW)."
Const ErrPtnr_cmd_End% = 1
Const ErrPtnr_cmd_Exit% = 1
Const ErrPtnr_cmd_Next% = 1
Const ErrPtnr_cmd_Resume% = 1
Const ErrPtnr_fram_CallStack% = 1
Const ErrPtnr_cmd_WWW% = 1
Const ErrPtnr_cmd_FileSave% = 1
Const ErrPtnr_cmd_EMail% = 0
Const ErrPtnr_cmd_Print% = 1
Const ErrPtnr_fram_Protocol% = 1
Const ErrPtnr_fram_ErrInfo% = 1
Const ErrPtnr_fram_ProgInfo% = 1
Const ErrPtnr_fram_Comment% = 1
'-- END Settings --- Do not modify THIS line!
'

' We need this declaration/function to open a URL from
' within the error dialog.
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' We need this declaration/function to identify the current
' Windows version.
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Declare Function GetVersionEx& Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO)
Private Const VER_PLATFORM_WIN32_NT& = 2
Private Const VER_PLATFORM_WIN32_WINDOWS& = 1
Private Const VER_PLATFORM_WIN32s& = 0

Dim Continuation%
Private Sub cmdContinue_Click(Index As Integer)
'.------------------------------------------------------------------------------
'.  Function :  The user just decided how to continue after the error
'.              situation. The error dialog will be closed.
'.------------------------------------------------------------------------------
Continuation% = Index
Hide
End Sub
Private Sub BuildDialog()
'.------------------------------------------------------------------------------
'.  Function :  Display the error dialog corresponding to the settings.
'.------------------------------------------------------------------------------

Dim y&  ' y-position of the next dialog frame
y& = 0
' comments
With framComment
    .Visible = ErrPtnr_fram_Comment
    If ErrPtnr_fram_Comment Then
        .Move 0, y&
        y& = y& + .Height
        lblComment = ErrPtnr_Comment
    End If
End With
'--------------
' error description
With framFehler
    .Visible = ErrPtnr_fram_ErrInfo
    If ErrPtnr_fram_ErrInfo Then
        .Move 0, y&
        y& = y& + .Height
    End If
End With
'--------------
' program information
With framProgInfo
    .Visible = ErrPtnr_fram_ProgInfo
    If ErrPtnr_fram_ProgInfo Then
        .Move 0, y&
        y& = y& + .Height
    End If
End With
'--------------
' protocol buttons
With framProtocol
    .Visible = ErrPtnr_fram_Protocol
    If ErrPtnr_fram_Protocol Then
        .Move 0, y&
        y& = y& + .Height
        cmdProtocol(0).Visible = ErrPtnr_cmd_Print
        cmdProtocol(1).Visible = ErrPtnr_cmd_FileSave
        cmdProtocol(2).Visible = ErrPtnr_cmd_EMail
        cmdProtocol(3).Visible = ErrPtnr_cmd_WWW
    End If
End With
'--------------
' call stack
With framCallStack
    .Visible = ErrPtnr_fram_CallStack
    If ErrPtnr_fram_CallStack Then
        .Move 0, y&
        y& = y& + .Height
    End If
End With
'--------------
' continuation buttons
With framContinue
    .Move 0, y&
    y& = y& + .Height
    cmdContinue(0).Visible = ErrPtnr_cmd_Resume
    cmdContinue(1).Visible = ErrPtnr_cmd_Next
    cmdContinue(2).Visible = ErrPtnr_cmd_Exit
    cmdContinue(3).Visible = ErrPtnr_cmd_End
End With

'----------------------------------------------------------
' adjust the form's height and width
'--------------
Width = (Width - ScaleWidth) + framFehler.Width
Height = (Height - ScaleHeight) + y&
Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2

End Sub
Public Function OnError%(AktModul$, AktProc$)
'.------------------------------------------------------------------------------
'.  Function :  The automatic error handler routines inserted in the project by
'.              ErrorPartner call this function.
'.              1. identify the error conditions
'.              2. display the error dialog (modal)
'.              3. Return the contionuation code to the calling function
'.------------------------------------------------------------------------------

' fill the dialog's label captions
lblError(0) = Err.Number & " (" & Err.Description & ")"
lblError(1) = AktModul$
lblError(2) = AktProc$
lblError(3) = Erl

' create programm information
If Len(lblProgInfo) = 0 Then lblProgInfo = App.EXEName + "  V" & App.Major & "." & App.Minor & "." & App.Revision
Screen.MousePointer = vbDefault

' build the error dialog
BuildDialog

' write the log file
If ErrPtnr_LogFile% Then WriteLogfile

' display the dialog
Beep
Show vbModal

' return the continuation code to the calling procedure
OnError% = Continuation%

End Function
Public Sub CallStack(a$)
'.------------------------------------------------------------------------------
'.  Function :  push a procedure to the call stack
'.------------------------------------------------------------------------------
a$ = Time$ + " " + a$
With lstCallStack
    .AddItem a$, 0
    If .ListCount > 20 Then .RemoveItem .ListCount - 1
End With
'--------------------------
' If you comment out the following line the call stack will be displayed
' in the VB IDE's debug window.
'Debug.Print a$
'--------------------------
End Sub
Private Sub BuildProtocol(a$)
'.------------------------------------------------------------------------------
'.  Function :  concatenate and append the error protocol
'.------------------------------------------------------------------------------

Dim t$, r$, f%, e&, b$
Dim v As OSVERSIONINFO

t$ = String$(70, "-") + vbCrLf
r$ = Space$(10)

' header
a$ = t$
a$ = a$ + "Errorprotocol" + vbCrLf
' program info
a$ = a$ + t$
a$ = a$ + UCase$(framProgInfo.Caption) + vbCrLf
a$ = a$ + r$ + lblProgInfo + vbCrLf
' error description
a$ = a$ + t$
a$ = a$ + UCase$(framFehler.Caption) + vbCrLf
For f% = 0 To 3
    a$ = a$ + r$ + lblErrTitel(f%) + vbTab + lblError(f%) + vbCrLf
Next
a$ = a$ + r$ + "Date/Time:" + vbTab + Date$ + " / " + Time$ + vbCrLf
' call stack
a$ = a$ + t$
a$ = a$ + UCase$(framCallStack.Caption) + vbCrLf
For f% = 0 To lstCallStack.ListCount - 1
    a$ = a$ + r$ + lstCallStack.List(f%) + vbCrLf
Next
' system information
a$ = a$ + t$
a$ = a$ + "SYSTEMINFO:" + vbCrLf
a$ = a$ + r$ + "OS:" + vbTab
v.dwOSVersionInfoSize = 148
e& = GetVersionEx&(v)
Select Case v.dwPlatformId
Case VER_PLATFORM_WIN32_NT: a$ = a$ + "Windows NT" + vbCrLf
Case VER_PLATFORM_WIN32_WINDOWS: a$ = a$ + "Windows 95" + vbCrLf
Case VER_PLATFORM_WIN32s: a$ = a$ + "Win32s" + vbCrLf
End Select
a$ = a$ + r$ + "Version:" + vbTab & v.dwMajorVersion & "." & v.dwMinorVersion & vbCrLf
a$ = a$ + r$ + "Build:" + vbTab & (v.dwBuildNumber And &HFFFF&) & "   " & LPSTRToVBString$(v.szCSDVersion) & vbCrLf
' Ende
a$ = a$ + t$

End Sub
Private Sub cmdProtocol_Click(Index As Integer)
'.------------------------------------------------------------------------------
'.  Function :  How does the user want to preceed with
'.              the error protocol?
'.------------------------------------------------------------------------------

Dim Prot$, n$, h%, p%, w$, zl%, z$

' concatenate the error protocoll
BuildProtocol Prot$

Select Case Index
'--------------------------
Case 0  ' print the protocol
'--------------------------
    Prot$ = Prot$ + vbCrLf
    Do
        ' fetch next line
        p% = InStr(Prot$, vbCrLf)
        If p% = 0 Then Exit Do
        n$ = Left$(Prot$, p% - 1) + " "
        Prot$ = Mid$(Prot$, p% + 2)
        ' print the line word by word with a word wrap in column 70
        zl% = 0
        z$ = vbNullString
        Do
            p% = InStr(n$, " ")
            If p% = 0 Then Exit Do
            w$ = Left$(n$, p%)
            n$ = Mid$(n$, p% + 1)
            zl% = zl% + p%
            If zl% > 73 Then
                If Len(z$) Then Printer.Print z$: z$ = Space$(20)
                zl% = p%
            End If
            z$ = z$ + w$
        Loop
        If Len(z$) Then Printer.Print z$
    Loop
    Printer.EndDoc
'--------------------------
Case 1  ' save the protocol
'--------------------------
    n$ = App.Path
    If Right$(n$, 1) <> "\" Then n$ = n$ + "\"
    n$ = n$ + App.EXEName + ".ERR"
    If Len(Dir$(n$)) Then
        If MsgBox("Overwrite existing file '" + n$ + "' ?", vbYesNo) = vbNo Then Exit Sub
    End If
    h% = FreeFile
    Open n$ For Output As #h%
    Print #h%, Prot$
    Close h%
    MsgBox "The errorprotocol was written into the file '" + n$ + "' !"
'--------------------------
Case 2  ' send the protocoll by email
'--------------------------
    z$ = "Now your standard email-client will be opened. The errordescription and the errorprotocol"
    z$ = z$ + vbCrLf + "had been copied into the clipboard for your comfortability."
    z$ = z$ + vbCrLf + ""
    z$ = z$ + vbCrLf + "Please copy the content of the clipboard into your email by pressing [Ctrl] and [Ins] simultaneous."
    MsgBox z$
    Clipboard.Clear
    Clipboard.SetText Prot$, 1
    Call ShellExecute(hWnd, "Open", "mailto:" + Trim$(ErrPtnr_EMail$), "", "", 1)
'--------------------------
Case 3  ' open a url
'--------------------------
    Clipboard.Clear
    Clipboard.SetText Prot$, 1
    Call ShellExecute(hWnd, "Open", ErrPtnr_WWW$, "", "", 1)
'--------------------------
End Select

End Sub
Public Sub SetProgInfo(Text$)
'.------------------------------------------------------------------------------
'  Define the contents of the program information text to be diplayed in the
'  error dialog. For example: "Program: YourApp v.1.99 / Serial no: 12345"
'  Syntax:
'       ErrPtnr.SetProgInfo "Program: YourApp v.1.99 / Serial no: 12345"
'.------------------------------------------------------------------------------
lblProgInfo = Text$
End Sub
Private Sub WriteLogfile()
'.------------------------------------------------------------------------------
'   append the error information to an existing log file or create a new one
'.------------------------------------------------------------------------------

Dim Prot$, n$, h%

' concatenate the error protocoll text
BuildProtocol Prot$

' append it to the log file
n$ = App.Path
If Right$(n$, 1) <> "\" Then n$ = n$ + "\"
n$ = n$ + App.EXEName + ".EPT"
h% = FreeFile
Open n$ For Append As #h%
Print #h%, Prot$
Close h%

End Sub
' extracts a VB string from a buffer containing a null terminated string
Public Function LPSTRToVBString$(ByVal s$)
Dim nullpos&
nullpos& = InStr(s$, Chr$(0))
If nullpos > 0 Then
    LPSTRToVBString = Left$(s$, nullpos - 1)
Else
    LPSTRToVBString = vbNullString
End If
End Function
