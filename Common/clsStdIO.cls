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

Public Function StdIn(TempDir As String, SpooltimeSeconds As Double) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim sBuffer As String, lBytesRead As Long, lStdInHandle As Long, _
  fn As Long, Tempfile As String, tTime1 As Date, Path As String
50030  Const iBufLen As Long = 1048576
50040 ' Const iBufLen As Long = 2048
50050
50060  lStdInHandle = GetStdHandle(STD_INPUT_HANDLE)
50070  StdIn = vbNullString
50080  Path = CompletePath(TempDir) & PDFCreatorSpoolDirectory & "\" & Environ$("REDMON_USER") & "\"
50090
50100  If DirExists(Path) = False Then
50110   MakePath Path
50120  End If
50130  Tempfile = GetTempFile(Path, "~PS")
50140  fn = FreeFile
50150  Open Tempfile For Output As #fn
50160  tTime1 = Now
50170  sBuffer = String(iBufLen, 0)
50180  Do
50190   If ReadFile(lStdInHandle, ByVal sBuffer, iBufLen, lBytesRead, ByVal 0&) <> 0 Then
50200     DoEvents
50210     If Len(Left$(sBuffer, lBytesRead)) > 0 Then
50220      Print #fn, Left$(sBuffer, lBytesRead);
50230     End If
50240    Else
50250     Exit Do
50260   End If
50270  Loop
50280  SpooltimeSeconds = DateDiff("s", tTime1, Now)
50290  Close #fn: DoEvents
50300  StdIn = Tempfile
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

