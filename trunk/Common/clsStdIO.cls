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
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim sBuffer As String, lBytesRead As Long, lStdInHandle As Long, _
  fn As Long, Tempfile As String, tTime1 As Date, Path As String
50060  Const iBufLen As Long = 1048576
50070 ' Const iBufLen As Long = 2048
50080
50090  lStdInHandle = GetStdHandle(STD_INPUT_HANDLE)
50100  StdIn = vbNullString
50110  If IsWin9xMe = True Then
50120    Path = CompletePath(TempDir) & PDFCreatorSpoolDirectory & "\"
50130   Else
50140    Path = CompletePath(TempDir) & PDFCreatorSpoolDirectory & "\" & Environ$("REDMON_USER") & "\"
50150  End If
50160  If DirExists(Path) = False Then
50170   MakePath Path
50180  End If
50190  Tempfile = GetTempFile(Path, "~PS")
50200  fn = FreeFile
50210  Open Tempfile For Output As #fn
50220  tTime1 = Now
50230  sBuffer = String(iBufLen, 0)
50240  Do
50250   If ReadFile(lStdInHandle, ByVal sBuffer, iBufLen, lBytesRead, ByVal 0&) <> 0 Then
50260     DoEvents
50270     If Len(Left$(sBuffer, lBytesRead)) > 0 Then
50280      Print #fn, Left$(sBuffer, lBytesRead);
50290     End If
50300    Else
50310     Exit Do
50320   End If
50330  Loop
50340  SpooltimeSeconds = DateDiff("s", tTime1, Now)
50350  Close #fn: DoEvents
50360  StdIn = Tempfile
50370 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50380 Exit Function
ErrPtnr_OnError:
50401 Select Case ErrPtnr.OnError("clsStdIO", "StdIn")
      Case 0: Resume
50420 Case 1: Resume Next
50430 Case 2: Exit Function
50440 Case 3: End
50450 End Select
50460 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub StdOut(sData As String)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Call WriteFile(GetStdHandle(STD_OUTPUT_HANDLE), sData & vbCrLf, Len(sData & vbCrLf), 0, ByVal 0&)
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("clsStdIO", "StdOut")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub StdErr(sData As String)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Call WriteFile(GetStdHandle(STD_ERROR_HANDLE), sData & vbCrLf, Len(sData & vbCrLf), 0, ByVal 0&)
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("clsStdIO", "StdErr")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

