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

Public Function StdIn(SpooltimeSeconds As Double) As String
 Dim sBuffer As String, lBytesRead As Long, lStdInHandle As Long
 Dim fn As Long, Tempfile As String, tTime1 As Date
 Const iBufLen As Long = 1048576
' Const iBufLen As Long = 2048

 lStdInHandle = GetStdHandle(STD_INPUT_HANDLE)
 StdIn = vbNullString

 Tempfile = GetTempFile(GetTempPath, "~PS")
 fn = FreeFile
 Open Tempfile For Output As #fn
 tTime1 = Now
 sBuffer = String(iBufLen, 0)
 Do
  If ReadFile(lStdInHandle, ByVal sBuffer, iBufLen, lBytesRead, ByVal 0&) <> 0 Then
'    MsgBox DateDiff("s", tTime1, Now): Exit Do
    DoEvents
    If Len(Left$(sBuffer, lBytesRead)) > 0 Then
     Print #fn, Left$(sBuffer, lBytesRead);
    End If
   Else
    Exit Do
  End If
'  DoEvents
 Loop
 SpooltimeSeconds = DateDiff("s", tTime1, Now)
 Close #fn: DoEvents
 StdIn = Tempfile
End Function

Public Sub StdOut(sData As String)
 Call WriteFile(GetStdHandle(STD_OUTPUT_HANDLE), sData & vbCrLf, Len(sData & vbCrLf), 0, ByVal 0&)
End Sub

Public Sub StdErr(sData As String)
 Call WriteFile(GetStdHandle(STD_ERROR_HANDLE), sData & vbCrLf, Len(sData & vbCrLf), 0, ByVal 0&)
End Sub
