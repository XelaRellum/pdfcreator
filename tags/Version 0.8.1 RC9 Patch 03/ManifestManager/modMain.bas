Attribute VB_Name = "modMain"
Option Explicit

Private Declare Function GetFullPathName Lib "kernel32.dll" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function BeginUpdateResource Lib "kernel32.dll" Alias "BeginUpdateResourceA" (ByVal pFileName As String, ByVal bDeleteExistingResources As Long) As Long
Private Declare Function UpdateResource Lib "kernel32.dll" Alias "UpdateResourceA" (ByVal hUpdate As Long, ByVal lpType As Any, ByVal lpName As Any, ByVal wLanguage As Long, ByRef lpData As Any, ByVal cbData As Long) As Long
Private Declare Function EndUpdateResource Lib "kernel32.dll" Alias "EndUpdateResourceA" (ByVal hUpdate As Long, ByVal fDiscard As Long) As Long

Private Const manifest As String = _
  "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf & _
  "<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" " & _
  "manifestVersion=""1.0"">" & vbCrLf & _
  "<assemblyIdentity  type=""win32"" processorArchitecture=""*"" " & _
  "version=""6.0.0.0"" name=""Company"" />" & vbCrLf & _
  "<dependency>" & vbCrLf & _
  "<dependentAssembly>" & vbCrLf & _
  "<assemblyIdentity type=""win32"" name=""Microsoft.Windows.Common-Controls"" " & _
  "version=""6.0.0.0"" " & _
  "language=""*"" processorArchitecture=""*"" publicKeyToken=""6595b64144ccf1df""/>" & vbCrLf & _
  "</dependentAssembly>" & vbCrLf & _
  "</dependency>" & vbCrLf & _
  "</assembly>"

Private AddFile As String, DelFile As String

Private Sub ShowHelp()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With App
50020   MsgBox .EXEName & " " & .Major & "." & .Minor & "." & .Revision & vbCrLf & vbCrLf & _
   "Parameters:" & vbCrLf & _
   " /Add to add a manifest file to a exe-file" & vbCrLf & _
   " /Del to delete a manifest file from a exe-file" & vbCrLf & vbCrLf & _
   "Examples:" & vbCrLf & _
   " " & .EXEName & " /Add=""C:\Test.exe""" & vbCrLf & _
   " " & .EXEName & " /Del=""C:\Test.exe"""
50090  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "ShowHelp")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub Main()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 AnalyzeCommandlineParameters
50020  If LenB(AddFile) = 0 And LenB(DelFile) = 0 Then
50030   ShowHelp
50040  End If
50050  If LenB(AddFile) > 0 Then
50060   If AddXPStyle(AddFile) = False Then
50070    MsgBox "Error add manifest to: " & AddFile
50080   End If
50090  End If
50100  If LenB(DelFile) > 0 Then
50110   If DelXPStyle(DelFile) = False Then
50120    MsgBox "Error del manifest from: " & DelFile
50130   End If
50140  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "Main")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub AnalyzeCommandlineParameters()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ' Commandswitches
50020  ' -ADD<Filename>
50030  ' -DEL<Filename>
50040  Dim cSwitch As String, Path As String
50050  If Len(VBA.Command$) > 0 Then
50060   cSwitch = CommandSwitch("ADD", True)
50070   SplitPath cSwitch, , Path
50080   If LenB(cSwitch) > 0 Then
50090    If FileExists(cSwitch) = True Then
50100     AddFile = cSwitch
50110    End If
50120   End If
50130   cSwitch = CommandSwitch("DEL", True)
50140   If LenB(cSwitch) > 0 Then
50150    If FileExists(cSwitch) = True Then
50160     AddFile = cSwitch
50170    End If
50180   End If
50190  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "AnalyzeCommandlineParameters")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function FileExists(FileStr As String) As Boolean
 On Error GoTo ErrorHandler
 FileExists = GetAttr(FileStr)
 Exit Function
ErrorHandler:
 FileExists = False
End Function

Private Function AddXPStyle(ByVal Filename As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010   Dim handle As Long
50020
50030   handle = BeginUpdateResource(Filename, False)
50040   If handle = 0 Then Exit Function
50050   UpdateResource handle, 24&, 1&, 0&, ByVal manifest, Len(manifest)
50060   EndUpdateResource handle, False
50070
50080   AddXPStyle = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "AddXPStyle")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function DelXPStyle(ByVal Filename As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010   Dim handle As Long
50020
50030   handle = BeginUpdateResource(Filename, False)
50040   If handle = 0 Then Exit Function
50050   UpdateResource handle, 24&, 1&, 0&, ByVal 0&, 0&
50060   EndUpdateResource handle, False
50070
50080   DelXPStyle = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "DelXPStyle")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function InCollection(colTest As Collection, sKey As String) As Boolean
50010  On Error GoTo ErrorHandler
50020  If VarType(colTest.Item(sKey)) = vbObject Then
50030  End If
50040  InCollection = True
50050 Exit Function
ErrorHandler:
50070  InCollection = False
End Function


Public Sub SplitPath(FullPath As String, Optional Drive As String, Optional Path As String, Optional Filename As String, Optional File As String, Optional Extension As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim nPos As Integer
50020  nPos = InStrRev(FullPath, "\")
50030  If nPos > 0 Then
50040    If Left$(FullPath, 2) = "\\" Then
50050     If nPos = 2 Then
50060      Drive = FullPath: Path = vbNullString: Filename = vbNullString: File = vbNullString
50070      Extension = vbNullString
50080      Exit Sub
50090     End If
50100    End If
50110    Path = Left$(FullPath, nPos - 1)
50120    Filename = Mid$(FullPath, nPos + 1)
50130    nPos = InStrRev(Filename, ".")
50140    If nPos > 0 Then
50150      File = Left$(Filename, nPos - 1)
50160      Extension = Mid$(Filename, nPos + 1)
50170     Else
50180      File = Filename
50190      Extension = vbNullString
50200    End If
50210   Else
50220    nPos = InStrRev(FullPath, ":")
50230    If nPos > 0 Then
50240      Path = Mid(FullPath, 1, nPos - 1): Filename = Mid(FullPath, nPos + 1)
50250      nPos = InStrRev(Filename, ".")
50260      If nPos > 0 Then
50270        File = Left$(Filename, nPos - 1)
50280        Extension = Mid$(Filename, nPos + 1)
50290       Else
50300        File = Filename
50310        Extension = vbNullString
50320      End If
50330     Else
50340      Path = vbNullString: Filename = FullPath
50350      nPos = InStrRev(Filename, ".")
50360      If nPos > 0 Then
50370        File = Left$(Filename, nPos - 1)
50380        Extension = Mid$(Filename, nPos + 1)
50390       Else
50400        File = Filename
50410        Extension = vbNullString
50420      End If
50430    End If
50440  End If
50450  If Left$(Path, 2) = "\\" Then
50460    nPos = InStr(3, Path, "\")
50470    If nPos Then
50480      Drive = Left$(Path, nPos - 1)
50490     Else
50500      Drive = Path
50510    End If
50520   Else
50530    If Len(Path) = 2 Then
50540     If Right$(Path, 1) = ":" Then
50550      Path = Path & "\"
50560     End If
50570    End If
50580    If Mid$(Path, 2, 2) = ":\" Then
50590     Drive = Left$(Path, 2)
50600    End If
50610  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "SplitPath")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function FullPath(Path As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010   Dim nBuffer As String
50020   Dim nFilePart As String
50030   Dim nLen As Long
50040
50050   nBuffer = Space$(255)
50060   nLen = GetFullPathName(Path, Len(nBuffer), nBuffer, nFilePart)
50070   If nLen Then
50080     FullPath = Left$(nBuffer, nLen)
50090   End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "FullPath")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

