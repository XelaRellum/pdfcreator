VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStdIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetStdHandle Lib "Kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function ReadFile Lib "Kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "Kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long

Private Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_ERROR_HANDLE = -12&

Public Function StdIn(TempDir As String, SpooltimeSeconds As Double) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim sBuffer As String, lBytesRead As Long, lStdInHandle As Long
50020  Dim fn As Long, Tempfile As String, tTime1 As Date
50030  Const iBufLen As Long = 1048576
50040 ' Const iBufLen As Long = 2048
50050
50060  lStdInHandle = GetStdHandle(STD_INPUT_HANDLE)
50070  StdIn = vbNullString
50080
50090  Tempfile = GetTempFile(TempDir, "~PS")
50100  fn = FreeFile
50110  Open Tempfile For Output As #fn
50120  tTime1 = Now
50130  sBuffer = String(iBufLen, 0)
50140  Do
50150   If ReadFile(lStdInHandle, ByVal sBuffer, iBufLen, lBytesRead, ByVal 0&) <> 0 Then
50160 '    MsgBox DateDiff("s", tTime1, Now): Exit Do
50170     DoEvents
50180     If Len(Left$(sBuffer, lBytesRead)) > 0 Then
50190      Print #fn, Left$(sBuffer, lBytesRead);
50200     End If
50210    Else
50220     Exit Do
50230   End If
50240 '  DoEvents
50250  Loop
50260  SpooltimeSeconds = DateDiff("s", tTime1, Now)
50270  Close #fn: DoEvents
50280  StdIn = Tempfile
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsStdIO", "StdIn")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub StdOut(sData As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Call WriteFile(GetStdHandle(STD_OUTPUT_HANDLE), sData & vbCrLf, Len(sData & vbCrLf), 0, ByVal 0&)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsStdIO", "StdOut")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub StdErr(sData As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Call WriteFile(GetStdHandle(STD_ERROR_HANDLE), sData & vbCrLf, Len(sData & vbCrLf), 0, ByVal 0&)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsStdIO", "StdErr")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

