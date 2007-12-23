VERSION 5.00
Begin VB.Form ErrPtnr 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "ProgInfo - Fehlermeldung"
   ClientHeight    =   5775
   ClientLeft      =   1395
   ClientTop       =   2220
   ClientWidth     =   7230
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7230
   Begin VB.Frame framComment 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.Label lblComment 
         BackStyle       =   0  'Transparent
         Caption         =   "lblComment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6975
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame framFehler 
      Caption         =   "Error description:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   7215
      Begin VB.Label lblErrTitel 
         Caption         =   "Error-Nr:"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblErrTitel 
         Caption         =   "Modul:"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblErrTitel 
         Caption         =   "Procedure:"
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblErrTitel 
         BackStyle       =   0  'Transparent
         Caption         =   "Line:     "
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   10
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblError 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "lblError(0)"
         ForeColor       =   &H80000008&
         Height          =   465
         Index           =   0
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   5085
      End
      Begin VB.Label lblError 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "lblError(1)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   7
         Top             =   720
         Width           =   840
      End
      Begin VB.Label lblError 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "lblError(2)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   9
         Top             =   960
         Width           =   840
      End
      Begin VB.Label lblError 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "lblError(3)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   11
         Top             =   1200
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Height          =   870
         Left            =   105
         TabIndex        =   3
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame framProtocol 
      Caption         =   "Error-Protocol ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   0
      TabIndex        =   12
      Top             =   2520
      Width           =   7215
      Begin VB.CommandButton cmdProtocol 
         Caption         =   "&WWW"
         Height          =   375
         Index           =   3
         Left            =   5640
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdProtocol 
         Caption         =   "&Save"
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdProtocol 
         Caption         =   "&EMail"
         Height          =   375
         Index           =   2
         Left            =   3840
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdProtocol 
         Caption         =   "&Print"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   13
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   0
      TabIndex        =   17
      Top             =   3240
      Width           =   7215
      Begin VB.ListBox lstCallStack 
         Appearance      =   0  '2D
         Height          =   1395
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   18
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   0
      TabIndex        =   19
      Top             =   4920
      Width           =   7215
      Begin VB.CommandButton cmdContinue 
         Caption         =   "&Repeat Command"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdContinue 
         Caption         =   "&Next Command"
         Height          =   495
         Index           =   1
         Left            =   2040
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdContinue 
         Caption         =   "Ex&it Procedure"
         Height          =   495
         Index           =   2
         Left            =   3840
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdContinue 
         Caption         =   "E&xit Program"
         Height          =   495
         Index           =   3
         Left            =   5640
         TabIndex        =   23
         Top             =   240
         Width           =   1335
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
Const ErrPtnr_LogFile% = 0
Const ErrPtnr_EMail$ = ""
Const ErrPtnr_WWW$ = "www.pdfforge.org"
Const ErrPtnr_Comment$ = "An error has occured. Look for technical support in the forums (WWW). " + vbCrLf + "Further advices and updates you can find also in the internet (WWW)."
Const ErrPtnr_cmd_End% = 1
Const ErrPtnr_cmd_Exit% = 1
Const ErrPtnr_cmd_Next% = 1
Const ErrPtnr_cmd_Resume% = 1
Const ErrPtnr_fram_CallStack% = 0
Const ErrPtnr_cmd_WWW% = 1
Const ErrPtnr_cmd_FileSave% = 1
Const ErrPtnr_cmd_EMail% = 0
Const ErrPtnr_cmd_Print% = 1
Const ErrPtnr_fram_Protocol% = 1
Const ErrPtnr_fram_ErrInfo% = 1
Const ErrPtnr_fram_Comment% = 1
'-- END Settings --- Do not modify THIS line!
'

' We need this declaration/function to open a URL from
' within the error dialog.
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

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

' count of rows for the CallStack-Listbox/Collection
Private Const CallStackLength% = 40
' The CallStack is during Runtime in a collection
Private CallStackCollection As New Collection
' DON'T fill the ListBox direct during Runtime,
' becuse the formcode is then loaded and must
' manually unload to end the main-programm !!!
' Thias was the bug in this form bevor V6.0.130)

Dim Continuation%, ProgInfo$
Private Sub cmdContinue_Click(Index As Integer)
'.------------------------------------------------------------------------------
'.  Function :  The user just decided how to continue after the error
'.              situation. The error dialog will be closed.
'.------------------------------------------------------------------------------
Continuation% = Index
Hide
End Sub
Private Sub BuildDialog(ContBits%)
'.------------------------------------------------------------------------------
'.  Function :  Display the error dialog corresponding to the settings.
'.------------------------------------------------------------------------------

Dim Y&  ' y-position of the next dialog frame
Y& = 0
' comments
With framComment
    .Visible = ErrPtnr_fram_Comment
    If ErrPtnr_fram_Comment Then
        .Move 0, Y&
        Y& = Y& + .Height
        lblComment = ErrPtnr_Comment
    End If
End With
'--------------
' error description
With framFehler
    .Visible = ErrPtnr_fram_ErrInfo
    If ErrPtnr_fram_ErrInfo Then
        .Move 0, Y&
        Y& = Y& + .Height
    End If
End With
'--------------
' protocol buttons
With framProtocol
    .Visible = ErrPtnr_fram_Protocol
    If ErrPtnr_fram_Protocol Then
        .Move 0, Y&
        Y& = Y& + .Height
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
        .Move 0, Y&
        Y& = Y& + .Height
    End If
End With
'--------------
' continuation buttons
With framContinue
    .Move 0, Y&
    Y& = Y& + .Height
    cmdContinue(0).Visible = (ErrPtnr_cmd_Resume And (ContBits% And 1) = 1)
    cmdContinue(1).Visible = (ErrPtnr_cmd_Next And (ContBits% And 2) = 2)
    cmdContinue(2).Visible = (ErrPtnr_cmd_Exit And (ContBits% And 4) = 4)
    cmdContinue(3).Visible = (ErrPtnr_cmd_End And (ContBits% And 8) = 8)
End With

'----------------------------------------------------------
' adjust the form's height and width
'--------------
Width = (Width - ScaleWidth) + framFehler.Width
Height = (Height - ScaleHeight) + Y&
Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2

End Sub
Public Function OnError%(AktModul$, AktProc$, Optional ContBits% = 255)
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
If Len(ProgInfo$) = 0 Then ProgInfo$ = App.EXEName + "  V" & App.Major & "." & App.Minor & "." & App.Revision
Caption = ProgInfo$ + " - Error message"
Screen.MousePointer = vbDefault

' fill the CallStack-listbox
CallStack2Listbox

' build the error dialog
BuildDialog ContBits%

' write the log file
If ErrPtnr_LogFile% Then WriteLogfile

' display the dialog
Beep
Show vbModal

' return the continuation code to the calling procedure
OnError% = Continuation%

End Function
Private Sub BuildProtocol(a$)
'.------------------------------------------------------------------------------
'.  Function :  concatenate and append the error protocol
'.------------------------------------------------------------------------------

Dim t$, R$, f%, e&, B$
Dim v As OSVERSIONINFO

t$ = String$(70, "-") + vbCrLf
R$ = Space$(10)

' header
a$ = t$
a$ = a$ + Chr$(80) & Chr$(68) & Chr$(70) & Chr$(67) & Chr$(114) & Chr$(101) & Chr$(97) & Chr$(116) & Chr$(111) & Chr$(114) & Chr$(32) & Chr$(45) & Chr$(32) & Chr$(119) & Chr$(119) & Chr$(119) & Chr$(46) & Chr$(112) & Chr$(100) & Chr$(102) & Chr$(102) & Chr$(111) & Chr$(114) & Chr$(103) & Chr$(101) & Chr$(46) & Chr$(111) & Chr$(114) & Chr$(103) & Chr$(13) & Chr$(10)
a$ = a$ + t$
a$ = a$ + "Errorprotocol" + vbCrLf
' program info
a$ = a$ + t$
a$ = a$ + Caption + vbCrLf
' error description
a$ = a$ + t$
a$ = a$ + UCase$(framFehler.Caption) + vbCrLf
For f% = 0 To 3
    a$ = a$ + R$ + lblErrTitel(f%) + vbTab + lblError(f%) + vbCrLf
Next
a$ = a$ + R$ + "Date/Time:" + vbTab + Date$ + " / " + time$ + vbCrLf
' call stack
a$ = a$ + t$
a$ = a$ + UCase$(framCallStack.Caption) + vbCrLf
For f% = 0 To lstCallStack.ListCount - 1
    a$ = a$ + R$ + lstCallStack.List(f%) + vbCrLf
Next
' system information
a$ = a$ + t$
a$ = a$ + "SYSTEMINFO:" + vbCrLf
a$ = a$ + " " + _
 Replace$(GetWinVersionStr, " [", vbCrLf + "  [", 1, 1, vbTextCompare) + vbCrLf
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
        z$ = ""
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
    Call ShellExecute(hwnd, "Open", "mailto:" + Trim$(ErrPtnr_EMail$), "", "", 1)
'--------------------------
Case 3  ' open a url
'--------------------------
    Clipboard.Clear
    Clipboard.SetText Prot$, 1
    Call ShellExecute(hwnd, "Open", ErrPtnr_WWW$, "", "", 1)
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
ProgInfo$ = Text$
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
    LPSTRToVBString = ""
End If
End Function
Public Sub CallStack(a$)
'.------------------------------------------------------------------------------
'.  Function :  push a procedure to the call stack
'.------------------------------------------------------------------------------
a$ = time$ + " " + a$
With CallStackCollection
    If .Count Then .Add a$, , 1 Else .Add a$
    If .Count > CallStackLength% Then .Remove .Count
End With
'--------------------------
' If you comment out the following line the call stack will be displayed
' in the VB IDE's debug window.
'Debug.Print a$
'--------------------------
End Sub
Public Sub CallStackParam(ParamName$, v As Variant)
'.------------------------------------------------------------------------------
'.  Function :  push a procedure's parameter to the call stack
'.------------------------------------------------------------------------------
Dim a$
Do
    If IsArray(v) Then a$ = "[ARRAY]": Exit Do
    If IsNull(v) Then a$ = "[NULL]": Exit Do
    If IsEmpty(v) Then a$ = "[EMPTY]": Exit Do
    If IsObject(v) Then
        If v Is Nothing Then
            a$ = "[NOTHING]"
        Else
            a$ = "[OBJECT:" + TypeName(v) + "]"
        End If
        Exit Do
    End If
    a$ = CStr(v): Exit Do
Loop
a$ = vbTab + ParamName$ + ":" + vbTab + a$
With CallStackCollection
    .Add a$, , , 1
    If .Count > CallStackLength% Then .Remove .Count
End With
'--------------------------
' If you comment out the following line the call stack will be displayed
' in the VB IDE's debug window.
'Debug.Print a$
'--------------------------
End Sub
Private Sub CallStack2Listbox()
'.------------------------------------------------------------------------------
'.  Sub :  fill the listbox from the CallStack-Array
'.------------------------------------------------------------------------------
Dim a As Variant
lstCallStack.Clear
For Each a In CallStackCollection
    lstCallStack.AddItem CStr(a)
Next
End Sub

