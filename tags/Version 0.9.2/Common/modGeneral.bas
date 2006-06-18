Attribute VB_Name = "modGeneral"
Option Explicit

Public Enum eSaveOpenType
 SaveFile = 0
 OpenFile = 1
End Enum

Public Enum eSortModeFiles
 notSorted = 0
 SortedByDate = 1
 SortedByName = 2
End Enum

Public Type InfoSpoolFile
 REDMON_PORT As String
 REDMON_JOB As String
 REDMON_PRINTER As String
 REDMON_MACHINE As String
 REDMON_USER As String
 REDMON_DOCNAME As String
 REDMON_FILENAME As String
 REDMON_SESSIONID As String
 SpoolFilename As String
 SpoolerAccount As String
 Computer As String
 Created As String
End Type

Public Enum RelativePathErrors
 rpErrTooManySteps = 50001
 rpErrDifferentRoot = 50011
End Enum

Public Function ANSItoASCII(ByVal AnsiString As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  CharToOem AnsiString, AnsiString
50020  ANSItoASCII = AnsiString
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "ANSItoASCII")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function ASCIItoANSI(ByVal AsciiString As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  OemToChar AsciiString, AsciiString
50020  ASCIItoANSI = AsciiString
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "ASCIItoANSI")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub CombineFiles(ByVal Filename As String, Files As Collection, _
 Optional BufferSize As Long = 65536, Optional stb As StatusBar)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, j As Long, fnSource As Long, fnDest As Long, sBuffer As String, _
  aLen As Double, tLen As Double, bsize As Long, fpos As Long
50030
50040  bsize = BufferSize
50050  Filename = Trim$(Filename)
50060  If Filename = vbNullString Or Files.Count = 0 Or Right$(Filename, 1) = "\" Then
50070   Exit Sub
50080  End If
50090  If Files.Count = 1 Then
50100   Exit Sub
50110  End If
50120  fnDest = FreeFile
50130  aLen = 0: tLen = 0
50140  For i = 1 To Files.Count
50150   aLen = aLen + FileLen(Files.Item(i))
50160  Next i
50170  Open Filename For Binary As #fnDest
50180  For i = 1 To Files.Count
50190   DoEvents
50200   If FileExists(Files.Item(i)) = False Then
50210    MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & vbCrLf & Files.Item(i)
50220   End If
50230   If FileLen(Files.Item(i)) > 0 Then
50240    fnSource = FreeFile
50250    Open Files.Item(i) For Binary Access Read As #fnSource
50260    If bsize > LOF(fnSource) Then
50270     bsize = LOF(fnSource)
50280    End If
50290    fpos = 1
50300    For j = 1 To LOF(fnSource) \ bsize
50310     fpos = (j - 1) * bsize + 1
50320     Seek #fnSource, fpos
50330     sBuffer = Input(bsize, fnSource)
50340     Put #fnDest, , sBuffer
50350     tLen = tLen + bsize
50360     If Not stb Is Nothing Then
50370      stb.Panels("Percent").Text = Format(CDbl(tLen) / CDbl(aLen), "0.0%")
50380     End If
50390     DoEvents
50400    Next j
50410    If LOF(fnSource) > (j - 1) * bsize Then
50420     fpos = (j - 1) * bsize + 1
50430     Seek #fnSource, fpos
50440     sBuffer = Input(LOF(fnSource) - (j - 1) * bsize, fnSource)
50450     Put #fnDest, , sBuffer
50460     tLen = tLen + (LOF(fnSource) - (j - 1) * bsize)
50470    End If
50480    Close #fnSource
50490   End If
50500   DoEvents
50510  Next i
50520  For i = 1 To Files.Count
50530   KillFile Files.Item(i)
50540   KillInfoSpoolfile Files.Item(i)
50550   DoEvents
50560  Next i
50570  Close #fnDest
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "CombineFiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub CombineFilesOld(ByVal Filename As String, Files As Collection, stb As StatusBar)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, fnSource As Long, fnDest As Long, sBuffer As String, _
  aLen As Double, tLen As Double
50030
50040  Filename = Trim$(Filename)
50050  If Filename = vbNullString Or Files.Count = 0 Or Right$(Filename, 1) = "\" Then
50060   Exit Sub
50070  End If
50080  If FileExists(Filename) = True Then
50090   Exit Sub
50100  End If
50110  If Files.Count = 1 Then
50120   Exit Sub
50130  End If
50140  fnDest = FreeFile
50150  aLen = 0
50160  For i = 1 To Files.Count
50170   aLen = aLen + FileLen(Files.Item(i))
50180  Next i
50190  Open Filename For Binary As #fnDest
50200  For i = 1 To Files.Count
50210   DoEvents
50220   If FileLen(Files.Item(i)) > 0 Then
50230    fnSource = FreeFile
50240    Open Files.Item(i) For Binary Access Read As #fnSource
50250    sBuffer = String(LOF(fnSource), Chr$(0))
50260    Get #fnSource, , sBuffer
50270    Put #fnDest, , sBuffer
50280    Close #fnSource
50290   End If
50300   tLen = tLen + FileLen(Files.Item(i))
50310   stb.Panels("Percent").Text = Format$(tLen / aLen, "0.0%")
50320   KillFile Files.Item(i)
50330   DoEvents
50340  Next i
50350  Close #fnDest
50360  stb.Panels("Percent").Text = vbNullString
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "CombineFilesOld")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub ComboSetListWidth(oCombo As Object, Optional ByVal nFixWidth As Variant, Optional ByVal nScaleMode As Variant)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With oCombo
50020   If IsMissing(nScaleMode) Or IsMissing(nFixWidth) Then
50030    nScaleMode = .Parent.ScaleMode
50040   End If
50050   If IsMissing(nFixWidth) Then
50060    Dim i As Long, nWidth As Long
50070    nFixWidth = 0
50080    For i = 0 To .ListCount - 1
50090     nWidth = .Parent.TextWidth(.List(i))
50100     If nWidth > nFixWidth Then
50110      nFixWidth = nWidth
50120     End If
50130    Next i
50140    nFixWidth = nFixWidth + .Parent.ScaleX(10, vbPixels, nScaleMode)
50150    If .ListCount > 8 Then
50160     nFixWidth = nFixWidth + .Parent.ScaleX(15, vbPixels, nScaleMode)
50170    End If
50180   End If
50190   SendMessage .hwnd, CB_SETDROPPEDWIDTH, .Parent.ScaleX(nFixWidth, nScaleMode, vbPixels), 0&
50200  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "ComboSetListWidth")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function CompletePath(Path As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If Len(Path) = 0 Then
50020   Exit Function
50030  End If
50040  Path = Trim$(Path)
50050  If Right$(Path, 1) = "\" Then
50060    CompletePath = LTrim$(Path)
50070   Else
50080    CompletePath = LTrim$(Path) & "\"
50090  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "CompletePath")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function CreateDir(Path As String) As Boolean
 On Error GoTo ErrorHandler
 MkDir Path
 CreateDir = True
Exit Function
ErrorHandler:
 CreateDir = False
End Function

Private Function UnQualifyPath(ByVal sFolder As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  sFolder = LTrim$(sFolder)
50020  UnQualifyPath = sFolder
50030  If LenB(sFolder) > 0 Then
50040   If Right$(sFolder, 1) = "\" Then
50050    UnQualifyPath = Left$(sFolder, Len(sFolder) - 1)
50060   End If
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "UnQualifyPath")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function DirExists(ByVal DirStr As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim WFD As WIN32_FIND_DATA, hFile As Long
50020  DirStr = UnQualifyPath(DirStr)
50030  hFile = FindFirstFileA(DirStr, WFD)
50040  DirExists = (hFile <> INVALID_HANDLE_VALUE) And _
                (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY)
50060  Call FindClose(hFile)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "DirExists")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function FileExists(ByVal FileStr As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim WFD As WIN32_FIND_DATA, hFile As Long
50020  hFile = FindFirstFileA(FileStr, WFD)
50030  FileExists = hFile <> INVALID_HANDLE_VALUE
50040  Call FindClose(hFile)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "FileExists")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsFilePrintable(Filename As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Ext As String, reg As clsRegistry
50020  IsFilePrintable = False
50030  If Len(Filename) = 0 Then
50040   Exit Function
50050  End If
50060  SplitPath Filename, , , , , Ext
50070  If Len(Ext) = 0 Then
50080   Exit Function
50090  End If
50100  Set reg = New clsRegistry
50110  reg.hkey = HKEY_LOCAL_MACHINE
50120  reg.KeyRoot = "Software\CLASSES\." & Ext
50130  If reg.KeyExists = False Then
50140   Set reg = Nothing
50150   Exit Function
50160  End If
50170  reg.KeyRoot = "Software\CLASSES\" & reg.GetRegistryValue("") & "\shell\print"
50180  If reg.KeyExists = False Then
50190   Set reg = Nothing
50200   Exit Function
50210  End If
50220  IsFilePrintable = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "IsFilePrintable")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetComputerName() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Const MAX_COMPUTERNAME_LENGTH As Long = 31
50020  Dim tStr As String
50030  tStr = String(MAX_COMPUTERNAME_LENGTH + 1, Chr$(0))
50040  GetComputerNameA tStr, MAX_COMPUTERNAME_LENGTH + 1
50050  GetComputerName = Left$(tStr, InStr(tStr, Chr$(0)) - 1)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetComputerName")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetDesktop() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim TempDesktop As String, reg As clsRegistry
50020  Set reg = New clsRegistry
50030  reg.hkey = HKEY_CURRENT_USER
50040  reg.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
50050  If IsWin9xMe = True Then
50060    TempDesktop = reg.GetRegistryValue("Common Desktop")
50070   Else
50080    TempDesktop = reg.GetRegistryValue("Desktop")
50090  End If
50100  Set reg = Nothing
50110  If Trim$(TempDesktop) = vbNullString Then
50120   TempDesktop = "C:\"
50130  End If
50140  GetDesktop = TempDesktop
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetDesktop")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetDrives() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim l As Long, Buffer As String, res As Long, drives As String, drv() As String, _
  i As Long
50030  l = 64: Buffer = Space(l)
50040  res = GetLogicalDriveStrings(l, Buffer)
50050  drives = Left$(Buffer, res)
50060  Set GetDrives = New Collection
50070  If Len(drives) > 0 Then
50080   If InStr(drives, Chr$(0)) > 0 Then
50090     drv = Split(drives, Chr$(0))
50100     For i = LBound(drv) To UBound(drv)
50110      If Trim$(drv(i)) <> vbNullString Then
50120       GetDrives.Add drv(i)
50130      End If
50140     Next i
50150    Else
50160     GetDrives.Add drives
50170   End If
50180  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetDrives")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetFileAttributesStr(Filename As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hFind As Long, WFD As WIN32_FIND_DATA, attr As Long, AA As String
50020  hFind = FindFirstFileA(Filename, WFD)
50030  attr = WFD.dwFileAttributes
50040  If attr And FILE_ATTRIBUTE_ARCHIVE Then AA = AA & "A"
50050  If attr And FILE_ATTRIBUTE_COMPRESSED Then AA = AA & "C"
50060  If attr And FILE_ATTRIBUTE_DIRECTORY Then AA = AA & "D"
50070  If attr And FILE_ATTRIBUTE_HIDDEN Then AA = AA & "H"
50080  If attr And FILE_ATTRIBUTE_NORMAL Then AA = AA & "N"
50090  If attr And FILE_ATTRIBUTE_READONLY Then AA = AA & "R"
50100  If attr And FILE_ATTRIBUTE_SYSTEM Then AA = AA & "S"
50110  GetFileAttributesStr = AA
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetFileAttributesStr")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetFiles(ByVal Path As String, Optional Searchmask As String = "*.*", _
 Optional Sorted As eSortModeFiles = notSorted) As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tColl1 As Collection, tColl2 As Collection, tFilename As String, _
  i As Long, tStrf() As String
50030  Set GetFiles = New Collection
50040  Path = Trim$(Path)
50050  If Len(Searchmask) > 0 Then
50060    If Not DirExists(Path) Then
50070     Exit Function
50080    End If
50090    Path = CompletePath(Path)
50100    tFilename = Dir(Path & Searchmask)
50110   Else
50120    If Not FileExists(Path) Then
50130     Exit Function
50140    End If
50150    tFilename = Dir(Path)
50160    SplitPath Path, , Path
50170    Path = CompletePath(Path)
50180  End If
50190  Set tColl1 = New Collection
50200  Do While tFilename <> ""
50211   Select Case Sorted
         Case eSortModeFiles.SortedByDate
50230     AddSortedStr tColl1, Format$(FileDateTime(Path & tFilename), "yyyymmddhhnnss") & "|" & Path & tFilename
50240    Case eSortModeFiles.SortedByName
50250     AddSortedStr tColl1, "|" & Path & tFilename
50260    Case Else
50270     tColl1.Add "|" & Path & tFilename
50280   End Select
50290   tFilename = Dir()
50300   DoEvents
50310  Loop
50320  Set tColl2 = New Collection
50330  For i = 1 To tColl1.Count
50340   tStrf = Split(tColl1(i), "|")
50350   SplitPath tStrf(1), , Path, tFilename
50360   Path = CompletePath(Path)
50370   tColl2.Add Path & "|" & Path & tFilename & "|" & FileLen(Path & tFilename) & "|" & FileDateTime(Path & tFilename)
50380  Next i
50390  Set GetFiles = tColl2
50400 ' Set tColl = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetFiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetDefaultAppData() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long, TempDefaultAppData As String
50020  TempDefaultAppData = Space$(MAX_PATH)
50030  If IsWin95 Or IsWinNT4 Then
50040    If SHGetFolderPathB(0, ssfAPPDATA Or CSIDL_FLAG_CREATE, -1, SHGFP_TYPE_DEFAULT, TempDefaultAppData) = 0 Then
50050     GetDefaultAppData = Left$(TempDefaultAppData, InStr(1, TempDefaultAppData, vbNullChar) - 1)
50060    End If
50070   Else
50080    If SHGetFolderPath(0, ssfAPPDATA Or CSIDL_FLAG_CREATE, -1, SHGFP_TYPE_DEFAULT, TempDefaultAppData) = 0 Then
50090     GetDefaultAppData = Left$(TempDefaultAppData, InStr(1, TempDefaultAppData, vbNullChar) - 1)
50100    End If
50110  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetDefaultAppData")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetCommonAppData() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim TempCommonAppData As String, reg As clsRegistry
50020  Set reg = New clsRegistry
50030  reg.hkey = HKEY_LOCAL_MACHINE
50040  reg.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
50050  TempCommonAppData = reg.GetRegistryValue("Common AppData")
50060  Set reg = Nothing
50070  If Trim$(TempCommonAppData) = vbNullString Then
50080   TempCommonAppData = "C:\"
50090  End If
50100  GetCommonAppData = TempCommonAppData
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetCommonAppData")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetMyAppData() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim TempMyAppData As String, reg As clsRegistry
50020  Set reg = New clsRegistry
50030  reg.hkey = HKEY_CURRENT_USER
50040  reg.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
50050  TempMyAppData = reg.GetRegistryValue("AppData")
50060  Set reg = Nothing
50070  If Trim$(TempMyAppData) = vbNullString Then
50080   TempMyAppData = "C:\"
50090  End If
50100  GetMyAppData = TempMyAppData
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetMyAppData")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetMyFiles() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim TempMyFiles, reg As clsRegistry
50020  Set reg = New clsRegistry
50030  reg.hkey = HKEY_CURRENT_USER
50040  reg.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
50050  TempMyFiles = reg.GetRegistryValue("Personal")
50060  Set reg = Nothing
50070  If Trim$(TempMyFiles) = vbNullString Then
50080   TempMyFiles = "C:\"
50090  End If
50100  GetMyFiles = TempMyFiles
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetMyFiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetFontsDirectory() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim TempMyFiles, reg As clsRegistry, tStr As String
50020  Set reg = New clsRegistry
50030  With reg
50040   .hkey = HKEY_CURRENT_USER
50050   .KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer"
50060   .Subkey = "Shell Folders"
50070   tStr = .GetRegistryValue("Fonts")
50080  End With
50090  Set reg = Nothing
50100  If LenB(tStr) = 0 Then
50110   tStr = CompletePath(GetWindowsDirectory) & "Fonts"
50120  End If
50130  GetFontsDirectory = tStr
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetFontsDirectory")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetMyLocalAppData() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim TempMyLocalAppData As String, reg As clsRegistry
50020  Set reg = New clsRegistry
50030  reg.hkey = HKEY_CURRENT_USER
50040  reg.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
50050  TempMyLocalAppData = reg.GetRegistryValue("Local AppData")
50060  Set reg = Nothing
50070  If Trim$(TempMyLocalAppData) = vbNullString Then
50080   TempMyLocalAppData = "C:\"
50090  End If
50100  GetMyLocalAppData = TempMyLocalAppData
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetMyLocalAppData")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetShortName(ByVal sLongFileName As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim lRetVal As Long, sShortPathName As String, iLen As Integer
50020  sShortPathName = Space(255)
50030  iLen = Len(sShortPathName)
50040  lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
50050  GetShortName = Left(sShortPathName, lRetVal)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetShortName")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetStringRessource(DLLExeFilename As String, ResID As Long) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hLib As Long, sBuf As String, lLen As Long
50020  hLib = LoadLibrary(DLLExeFilename)
50030  If hLib Then
50040   sBuf = Space$(255)
50050   lLen = LoadString(hLib, ResID, sBuf, Len(sBuf))
50060   If lLen Then
50070     sBuf = Left$(sBuf, lLen)
50080    Else
50090     sBuf = vbNullString
50100   End If
50110   FreeLibrary hLib
50120  End If
50130  GetStringRessource = sBuf
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetStringRessource")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetSystemDirectory() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long, Sysdir As String
50020  Sysdir = Space$(MAX_PATH)
50030  res = GetSystemDirectoryA(Sysdir, MAX_PATH)
50040  If res > 0 Then
50050    GetSystemDirectory = Left$(Sysdir, res)
50060   Else
50070    GetSystemDirectory = "C:\Windows"
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetSystemDirectory")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetTempFile(Optional ByVal Path As String, Optional Prefix As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long, Tempfile As String, tPath As String
50020  Tempfile = Space$(MAX_PATH)
50030  tPath = LTrim$(Path)
50040  If DirExists(tPath) = False Then
50050   MakePath tPath
50060  End If
50070
50080  res = GetTempFileNameA(tPath, Prefix, 0, Tempfile)
50090  If res <> 0 Then
50100    GetTempFile = Left$(Tempfile, InStr(Tempfile, Chr$(0)) - 1)
50110   Else
50120    GetTempFile = "~.tmp"
50130  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetTempFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetTempPath() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long, TempDir As String
50020
50030  If IsWin9xMe = True Then
50040    ' In Win9xMe you cannot use the username because the username can
50050    ' have forbidden chars like \/:*?"<>|
50060    TempDir = CompletePath(GetTempPathApi) & "PDFCreator\"
50070   Else
50080    If IsWinNT4 = True Then
50090      TempDir = GetMyAppData
50100      If LenB(Environ$("Redmon_User")) > 0 Then
50110        TempDir = CompletePath(TempDir) & "PDFCreator\" & Environ$("Redmon_User")
50120       Else
50130        TempDir = CompletePath(TempDir) & "PDFCreator\" & GetUsername
50140      End If
50150     Else
50160      TempDir = GetMyLocalAppData
50170      TempDir = Left(TempDir, InStrRev(TempDir, "\")) & "Temp\PDFCreator\"
50180    End If
50190  End If
50200  If Trim$(TempDir) = vbNullString Then
50210   TempDir = GetPDFCreatorApplicationPath & "Temp\"
50220  End If
50230
50240  GetTempPath = CompletePath(TempDir)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetTempPath")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetTempPathReg(hProfile As hkey) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry, tStr As String
50020  Set reg = New clsRegistry
50030  With reg
50040   .hkey = hProfile
50050   .KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
50060   tStr = CompletePath(.GetRegistryValue("Local Settings"))
50070  End With
50080  Set reg = Nothing
50090  If LenB(tStr) = 0 Then
50100    If IsWin9xMe = True Then
50110      tStr = CompletePath(GetTempPathApi)
50120     Else
50130      If IsWinNT4 = True Then
50140       tStr = CompletePath(GetTempPathApi)
50150       If LenB(Environ$("Redmon_User")) > 0 Then
50160         tStr = tStr & Environ$("Redmon_User")
50170        Else
50180         tStr = tStr & GetUsername
50190       End If
50200      End If
50210    End If
50220   Else
50230    tStr = CompletePath(tStr) & "Temp\"
50240  End If
50250  GetTempPathReg = tStr
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetTempPathReg")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetTempPathApi() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long, TempDir As String
50020  res = GetTempPathA(0, "")
50030  TempDir = Space$(res)
50040  res = GetTempPathA(res, TempDir)
50050  If res > 0 Then
50060    GetTempPathApi = Left$(TempDir, res)
50070   Else
50080    GetTempPathApi = "C:\"
50090  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetTempPathApi")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetUsername() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Const MAX_USERNAME_LENGTH = 128
50020  Dim tStr As String
50030  tStr = String(MAX_USERNAME_LENGTH, Chr$(0))
50040  GetUserNameA tStr, MAX_USERNAME_LENGTH
50050  GetUsername = Left$(tStr, InStr(tStr, Chr$(0)) - 1)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetUsername")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetWindowsDirectory() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long, WinDir As String
50020  WinDir = Space$(MAX_PATH)
50030  res = GetWindowsDirectoryA(WinDir, MAX_PATH)
50040  If res > 0 Then
50050    GetWindowsDirectory = Left$(WinDir, res)
50060   Else
50070    GetWindowsDirectory = "C:\Windows"
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetWindowsDirectory")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub HTMLHelp_ShowTopic(Optional sTopicFile As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hwndHelp As Long
50020  If sTopicFile = vbNullString Then
50030    hwndHelp = HtmlHelp(0, HelpFile, HH_DISPLAY_TOPIC, ByVal 0&)
50040   Else
50050    hwndHelp = HtmlHelpTopic(0, HelpFile, HH_DISPLAY_TOPIC, sTopicFile)
50060  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "HTMLHelp_ShowTopic")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function IsForbiddenChars(chkStr As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, Forbiddenchars As String
50020  IsForbiddenChars = False
50030  Forbiddenchars = "\/:*?<>|%"""
50040  For i = 1 To Len(Forbiddenchars)
50050   If InStr(chkStr, Mid$(Forbiddenchars, i, 1)) > 0 Then
50060    IsForbiddenChars = True
50070    Exit Function
50080   End If
50090  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "IsForbiddenChars")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsFormLoaded(Form As Form) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim nForm As Form
50020  For Each nForm In Forms
50030   If nForm Is Form Then
50040    IsFormLoaded = True
50050    Exit For
50060   End If
50070  Next
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "IsFormLoaded")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function KillFile(Filename As String) As Boolean
 On Error Resume Next
 KillFile = False
 If FileExists(Filename) = True Then
  If FileInUse(Filename) = False Then
   Kill Filename
   If FileExists(Filename) = False Then
    KillFile = True
   End If
  End If
 End If
End Function

Public Function LoadDLL(DLLPath As String) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  LoadDLL = LoadLibrary(DLLPath)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "LoadDLL")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function LoadIcon(ByVal lngSize As IconSize, ByVal strPath As String, Icon As IPictureDisp) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim iuUnkown As IUnknown, itIcon As IconType, CLSID As CLSIdType, _
  sfiShellInfo As ShellFileInfoType
50030  Call SHGetFileInfo(strPath, 0&, sfiShellInfo, Len(sfiShellInfo), SHGFI_USEFILEATTRIBUTES Or lngSize)
50040  With itIcon
50050   .cbSize = Len(itIcon)
50060   .picType = vbPicTypeIcon
50070   .hIcon = sfiShellInfo.hIcon
50080  End With
50090  CLSID.ID(8) = &HC0
50100  CLSID.ID(15) = &H46
50110  Call OleCreatePictureIndirect(itIcon, CLSID, 1&, iuUnkown)
50120  Set Icon = iuUnkown
50130  If Icon Is Nothing Then
50140    LoadIcon = 1
50150   Else
50160    LoadIcon = 0
50170  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "LoadIcon")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function MakePath(Path As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim dirs() As String, tStr As String, i As Long
50020  MakePath = False
50030  Path = LTrim$(Path)
50040  If Len(Path) <= 3 Then
50050   Exit Function
50060  End If
50070  Path = Replace$(Path, "/", "\", , , vbTextCompare)
50080  If InStr(1, Path, "\", vbTextCompare) = 0 Then
50090   Exit Function
50100  End If
50110  dirs = Split(Path, "\")
50120  If Mid$(Path, 2, 1) = ":" Then ' local drive
50130   If UCase$(Mid$(Path, 1, 1)) < "A" Or UCase$(Mid$(Path, 1, 1)) > "Z" Then
50140    Exit Function
50150   End If
50160   tStr = dirs(0)
50170   For i = 1 To UBound(dirs)
50180    tStr = tStr & "\" & dirs(i)
50190    If DirExists(tStr) = False Then
50200     Call CreateDir(tStr)
50210     If DirExists(tStr) = False Then
50220      Exit Function
50230     End If
50240    End If
50250   Next i
50260   MakePath = True
50270  End If
50280  If Mid(Path, 1, 2) = "\\" Then ' network share
50290   If UBound(dirs) < 4 Then
50300    Exit Function
50310   End If
50320   tStr = "\\" & dirs(2) & "\" & dirs(3)
50330   For i = 4 To UBound(dirs)
50340    tStr = tStr & "\" & dirs(i)
50350    If DirExists(tStr) = False Then
50360     Call CreateDir(tStr)
50370     If DirExists(tStr) = False Then
50380      Exit Function
50390     End If
50400    End If
50410   Next i
50420   MakePath = True
50430  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "MakePath")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub OpenDocument(sFilename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim sDirectory As String, res As Long, handle As Long
50020
50030  handle = GetDesktopWindow()
50040  res = ShellExecute(handle, "open", sFilename, _
    vbNullString, vbNullString, vbNormalFocus)
50060
50070  If res = SE_ERR_NOTFOUND Then
50080   ElseIf res = SE_ERR_NOASSOC Then
50090     sDirectory = Space$(260)
50100     sDirectory = CompletePath(GetSystemDirectory)
50110     Call ShellExecute(handle, vbNullString, _
      "RUNDLL32.EXE", "shell32.dll,OpenAs_RunDLL " & _
      sFilename, sDirectory, vbNormalFocus)
50140   End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "OpenDocument")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub EditDocument(sFilename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long, handle As Long
50020  handle = GetDesktopWindow()
50030  res = ShellExecute(handle, "edit", sFilename, vbNullString, vbNullString, vbNormalFocus)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "EditDocument")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub RemoveX(Frm As Form)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hMenu As Long, nPosition As Long
50020
50030  hMenu = GetSystemMenu(Frm.hwnd, 0)
50040  If hMenu <> 0 Then
50050   nPosition = GetMenuItemCount(hMenu)
50060   Call RemoveMenu(hMenu, nPosition - 1, MF_REMOVE Or MF_BYPOSITION)
50070   Call RemoveMenu(hMenu, nPosition - 2, MF_REMOVE Or MF_BYPOSITION)
50080   Call DrawMenuBar(Frm.hwnd)
50090  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "RemoveX")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function ReplaceForbiddenChars(chkStr As String, Optional ReplaceChar As String = "_", Optional AdditionalForbiddenChars = "") As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStr As String, i As Long, Forbiddenchars As String
50020  Forbiddenchars = "\/:*?<>|""" & AdditionalForbiddenChars
50030  tStr = chkStr
50040  For i = 1 To Len(Forbiddenchars)
50050   tStr = Replace$(tStr, Mid$(Forbiddenchars, i, 1), ReplaceChar)
50060  Next i
50070  ReplaceForbiddenChars = tStr
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "ReplaceForbiddenChars")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function RunProgramWait(strCmdLine As String, Optional ShowWindow As Boolean = True) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim lngRetCode As Long, proc As PROCESS_INFORMATION, Start As STARTUPINFO, _
  secAttr As SECURITY_ATTRIBUTES, lngRet As Long, lngExit As Long
50030
50040  Start.cb = Len(Start)
50050  If ShowWindow = False Then
50060   Start.dwFlags = STARTF_USESHOWWINDOW
50070   Start.wShowWindow = SW_HIDE
50080  End If
50090
50100  lngRet = CreateProcessA(0&, strCmdLine, _
  0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, Start, proc)
50120
50130  GetExitCodeProcess proc.hProcess, lngExit
50140
50150  Do While lngExit = STILL_ACTIVE
50160   DoEvents
50170   Sleep 100
50180   GetExitCodeProcess proc.hProcess, lngExit
50190  Loop
50200  lngRet = CloseHandle(proc.hProcess)
50210
50220  RunProgramWait = lngExit
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "RunProgramWait")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
Select Case ErrPtnr.OnError("modGeneral", "SplitPath")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub MoveMouseToCommandButton(cmdButton As Object)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim pt As POINTAPI, cRect As Rect
50020  GetCursorPos pt
50030  GetWindowRect cmdButton.hwnd, cRect
50040  With cRect
50050   pt.x = (.Left + .Right) / 2
50060   pt.Y = (.Top + .Bottom) / 2
50070  End With
50080  ScreenToAbsolute pt
50090  mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_MOVE, pt.x, pt.Y, 0, GetMessageExtraInfo()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "MoveMouseToCommandButton")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function IsValidPath(Path As String, Optional ByVal TestUNCPaths As Boolean = True) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, c As Long, nbytes() As Byte, nChar As String, nOK As Boolean, _
  nParentPath As String, nPos As Long, nRoot As String, nPath As String, _
  nPathParts() As String, nStart As Long
50040
50050  Const kCharAsterisk = 42, kCharBackSlash = 92, kCharColon = 58, _
  kCharGreaterThan = 62, kCharLowerThan = 60, kCharPipe = 124, _
  kCharQuestion = 63, kCharQuote = 34, kCharSlash = 47
50080
50090  nPath = Path
50100  PathUnquoteSpaces nPath
50110  nPath = Left$(nPath, InStrNullChar(nPath))
50120  If Not CBool(PathIsUNC(nPath) And Not TestUNCPaths) Then
50130   If CBool(PathFileExists(nPath)) Then
50140    IsValidPath = True
50150    Exit Function
50160   End If
50170  End If
50180  nRoot = nPath
50190  If PathStripToRoot(nRoot) Then
50200   nRoot = Left$(nRoot, InStrNullChar(nRoot))
50210   If PathIsUNC(nRoot) Then
50220     nPos = InStrRev(nRoot, "\")
50230     nRoot = Left$(nRoot, nPos - 1)
50240     If PathIsUNCServer(nRoot) Then
50250      nPath = Mid$(nPath, Len(nRoot) + 1)
50260     End If
50270    Else
50280     nPath = Mid$(nPath, Len(nRoot) + 1)
50290   End If
50300   If Len(nPath) Then
50310    nPathParts = Split(nPath, "\")
50320    If Len(nPathParts(0)) Then
50330      nStart = 0
50340     Else
50350      nStart = 1
50360    End If
50370    For i = nStart To UBound(nPathParts)
50380     nbytes = StrConv(nPathParts(i), vbFromUnicode)
50390     For c = 0 To UBound(nbytes)
50401      Select Case nbytes(c)
            Case Is < 32, Is > 255, kCharAsterisk, kCharBackSlash, kCharColon, kCharGreaterThan, kCharLowerThan, kCharPipe, kCharQuestion, kCharQuote, kCharSlash
50420        nOK = False
50430        Exit For
50440       Case Else
50450        nOK = True
50460      End Select
50470     Next c
50480     If Not nOK Then
50490      Exit For
50500     End If
50510    Next i
50520    IsValidPath = nOK
50530   End If
50540  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "IsValidPath")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub FormInTaskbar(Form As Form, ByVal ShowInTaskBar As Boolean, Optional ByVal NoFlicker As Boolean, Optional SetVisible As Boolean = True)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim nStyle As Long, nVisible As Boolean
50020  Const GWL_EXSTYLE = (-20), WS_EX_APPWINDOW = &H40000
50030  With Form
50040   nVisible = .Visible
50050   If NoFlicker And nVisible Then
50060    LockWindowUpdate GetDesktopWindow()
50070   End If
50080   .Visible = False
50090   nStyle = GetWindowLong(.hwnd, GWL_EXSTYLE)
50101   Select Case ShowInTaskBar
         Case True
50120     nStyle = nStyle Or WS_EX_APPWINDOW
50130    Case False
50140     nStyle = nStyle And Not WS_EX_APPWINDOW
50150   End Select
50160   SetWindowLong .hwnd, GWL_EXSTYLE, nStyle
50170   .Refresh
50180   If SetVisible Then
50190    .Visible = nVisible
50200   End If
50210   If NoFlicker And nVisible Then
50220    LockWindowUpdate 0&
50230   End If
50240  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "FormInTaskbar")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function RemoveCompletePath(Path As String, Optional ShowDialog As Boolean) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ret As Long, SHFILEOP As SHFILEOPSTRUCT, ShowD As Long, FOBuf() As Byte, _
  LenSHFileOp As Long
50030
50040  LenSHFileOp = LenB(SHFILEOP)
50050  ReDim FOBuf(1 To LenSHFileOp)
50060  If ShowDialog Then
50070    ShowD = 0
50080   Else
50090    ShowD = FOF_SILENT
50100  End If
50110
50120  With SHFILEOP
50130   .wFunc = FO_DELETE
50140   .pFrom = Path
50150   .pTo = vbNullString
50160   .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR Or ShowD
50170  End With
50180
50190  ret = SHFileOperation(SHFILEOP)
50200  DoEvents
50210
50220  Call MoveMemory(FOBuf(1), SHFILEOP, LenSHFileOp)
50230  Call MoveMemory(FOBuf(21), FOBuf(19), 12)
50240  Call MoveMemory(SHFILEOP, FOBuf(1), LenSHFileOp)
50250
50260  If ret <> 0 Then
50270   If SHFILEOP.fAnyOperationAborted <> 0 Then
50280    ret = -1
50290   End If
50300  End If
50310  RemoveCompletePath = ret
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "RemoveCompletePath")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function RemoveLeadingAndTrailingQuotes(tStr As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If Len(tStr) > 0 Then
50020   If Mid$(tStr, 1, 1) = """" Then
50030    tStr = Mid$(tStr, 2)
50040   End If
50050   If Len(tStr) > 0 Then
50060    If Mid$(tStr, Len(tStr), 1) = """" Then
50070     tStr = Mid$(tStr, 1, Len(tStr) - 1)
50080    End If
50090   End If
50100  End If
50110  RemoveLeadingAndTrailingQuotes = tStr
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "RemoveLeadingAndTrailingQuotes")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub ScreenToAbsolute(lpPoint As POINTAPI)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  lpPoint.x = lpPoint.x * (&HFFFF& / GetSystemMetrics(SM_CXSCREEN))
50020  lpPoint.Y = lpPoint.Y * (&HFFFF& / GetSystemMetrics(SM_CYSCREEN))
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "ScreenToAbsolute")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetOptimalComboboxHeigth(cmb As ComboBox, ParentForm As Form)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim pt As POINTAPI, rc As Rect, cWidth As Long, OldScaleMode As Long
50020
50030  OldScaleMode = ParentForm.ScaleMode
50040  ParentForm.ScaleMode = vbPixels
50050
50060  cWidth = cmb.Width
50070  Call GetWindowRect(cmb.hwnd, rc)
50080  pt.x = rc.Left
50090  pt.Y = rc.Top
50100  Call ScreenToClient(ParentForm.hwnd, pt)
50110  Call MoveWindow(cmb.hwnd, pt.x, pt.Y, cmb.Width, _
  SendMessage(cmb.hwnd, CB_GETITEMHEIGHT, 0, ByVal 0) * (cmb.ListCount + 2), _
  True)
50140  ParentForm.ScaleMode = OldScaleMode
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "SetOptimalComboboxHeigth")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function UnLoadDLL(DLLHandle As Long) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long
50020  If DLLHandle <> 0 Then
50030   res = FreeLibrary(DLLHandle)
50040   If res <> 0 Then
50050     UnLoadDLL = True
50060    Else
50070     UnLoadDLL = False
50080   End If
50090  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "UnLoadDLL")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function UnloadDLLComplete(ByRef DLLHandle As Long) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim c As Long
50020  If DLLHandle = 0 Then
50030   Exit Function
50040  End If
50050  c = 0
50060  Do While UnLoadDLL(DLLHandle)
50070   c = c + 1
50080   DoEvents
50090  Loop
50100  UnloadDLLComplete = c: DLLHandle = 0
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "UnloadDLLComplete")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetStrFromPtrA(lpszA As Long) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  GetStrFromPtrA = String$(lstrlen(ByVal lpszA), 0)
50020  Call lstrcpy(ByVal GetStrFromPtrA, ByVal lpszA)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetStrFromPtrA")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function RaiseAPIError(Optional ErrorNumber As Long = 0) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ErrorMsg As String, ErrNum As Long
50020  If ErrorNumber <> 0 Then
50030    ErrNum = ErrorNumber
50040   Else
50050    ErrNum = Err.LastDllError
50060  End If
50070  ErrorMsg = String(256, 0)
50080  ErrorMsg = Left$(ErrorMsg, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, ErrNum, 0&, ErrorMsg, Len(ErrorMsg), ByVal 0))
50090  If Mid(ErrorMsg, Len(ErrorMsg) - 1) = vbCrLf Then
50100   ErrorMsg = Mid(ErrorMsg, 1, Len(ErrorMsg) - 2)
50110  End If
50120  RaiseAPIError = ErrNum & ": " & ErrorMsg
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "RaiseAPIError")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function StringInCollection(coll As Collection, Str1 As String, Optional CaseSensitive As Boolean = False) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, tStr As String
50020  StringInCollection = False: tStr = UCase$(Str1)
50030  If Not coll Is Nothing And Len(Str1) > 0 Then
50040   For i = 1 To coll.Count
50050    If CaseSensitive = True Then
50060      If coll(i) = Str1 Then
50070       StringInCollection = True
50080       Exit Function
50090      End If
50100     Else
50110      If UCase$(coll(i)) = UCase$(Str1) Then
50120       StringInCollection = True
50130       Exit Function
50140      End If
50150    End If
50160   Next i
50170  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "StringInCollection")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function RemoveAllKnownFileExtensions(Filename As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tColl As Collection, Ext As String, File As String, OldFilename As String, reg As clsRegistry
50020  RemoveAllKnownFileExtensions = Filename
50030  SplitPath Filename, , , , File, Ext
50040  If LenB(Ext) > 0 Then
50050   Set reg = New clsRegistry
50060   reg.hkey = HKEY_CLASSES_ROOT
50070   reg.KeyRoot = "." & Ext
50080   Do Until reg.KeyExists = False
50090    SplitPath Filename, , , , File, Ext
50100    Filename = File
50110    reg.KeyRoot = "." & Ext
50120    DoEvents
50130   Loop
50140   If LenB(Ext) > 0 Then
50150     RemoveAllKnownFileExtensions = File & "." & Ext
50160    Else
50170     RemoveAllKnownFileExtensions = File
50180   End If
50190   Set reg = Nothing
50200  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "RemoveAllKnownFileExtensions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetIExplorerVersion() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry
50020  Set reg = New clsRegistry
50030  With reg
50040   .hkey = HKEY_LOCAL_MACHINE
50050   .KeyRoot = "SOFTWARE\Microsoft\Internet Explorer"
50060   GetIExplorerVersion = .GetRegistryValue("Version")
50070  End With
50080  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetIExplorerVersion")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function FormISLoaded(FormName As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  FormISLoaded = False
50030  For i = 1 To Forms.Count
50040   If UCase$(FormName) = UCase$(Forms(i - 1).Name) Then
50050    FormISLoaded = True
50060   End If
50070  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "FormISLoaded")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function InCollection(colTest As Collection, sKey As String) As Boolean
'
' Check to see if item [sKey] is in collection [colTest].
' Return True if it is, false if not
'
 On Error GoTo ErrorHandler
 If VarType(colTest.Item(sKey)) = vbObject Then
  '
  ' This test will indicate if the item actually exists in the
  ' collection. No further checking is needed.
  '
 End If
 InCollection = True
Exit Function
ErrorHandler:
   InCollection = False
End Function

Public Sub DrawBorder3D(ByVal obj As Object, Index As Integer, BorderWidth As Long)
 Dim R As Rect, DrawObj As Object, tRedraw As Boolean, tDrawwidth As Integer, _
  tScalemode As Long, RectType As Long, RectStyle As Long

 On Error Resume Next
 If BorderWidth <= 0 Then
  Exit Sub
 End If
 RectStyle = BF_RECT
 RectType = Choose(Index, EDGE_RAISED, EDGE_SUNKEN, EDGE_ETCHED, EDGE_BUMP)
 If TypeOf obj Is Form Or TypeOf obj Is PictureBox Then
   With obj
    tScalemode = .ScaleMode
    .ScaleMode = vbPixels
    R.Left = .ScaleLeft
    R.Top = .ScaleTop
    R.Right = .ScaleWidth
    R.Bottom = .ScaleHeight
   End With
   Set DrawObj = obj
  Else
   With obj
    R.Left = .Left - BorderWidth
    R.Top = .Top - BorderWidth
    R.Right = R.Left + .Width + BorderWidth + 1
    R.Bottom = R.Top + .Height + BorderWidth + 1
   End With
   Set DrawObj = obj.Container
 End If
 With DrawObj
  tRedraw = .AutoRedraw
  tDrawwidth = .DrawWidth
  .DrawWidth = BorderWidth
  .AutoRedraw = True
  DrawEdge .hdc, R, RectType, RectStyle
  .AutoRedraw = tRedraw
  .DrawWidth = tDrawwidth
  .Refresh
  .ScaleMode = tScalemode
 End With
End Sub

Public Sub WriteInfoSpoolfile(SpoolFilename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim isf As InfoSpoolFile, fn As Long, Path As String, File As String, _
  infFile As String
50030
50040  With isf
50050   .REDMON_PORT = Environ$("REDMON_PORT")
50060   .REDMON_JOB = Environ$("REDMON_JOB")
50070   .REDMON_PRINTER = Environ$("REDMON_PRINTER")
50080   .REDMON_MACHINE = Environ$("REDMON_MACHINE")
50090   .REDMON_USER = Environ$("REDMON_USER")
50100   .REDMON_DOCNAME = Environ$("REDMON_DOCNAME")
50110   .REDMON_FILENAME = Environ$("REDMON_FILENAME")
50120   .REDMON_SESSIONID = Environ$("REDMON_SESSIONID")
50130   .SpoolFilename = SpoolFilename
50140   .SpoolerAccount = GetUsername
50150   .Computer = GetComputerName
50160   .Created = Now
50170  End With
50180  fn = FreeFile
50190  SplitPath SpoolFilename, , Path, , File
50200  infFile = CompletePath(Path) & File & ".inf"
50210  Open infFile For Binary As #fn
50220  Put #fn, , isf
50230  Close #fn
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "WriteInfoSpoolfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function ReadInfoSpoolfile(SpoolFilename As String) As InfoSpoolFile
 On Error Resume Next
 Dim tVar As Variant, fn As Long, infFile As String, _
  Path As String, File As String
 SplitPath SpoolFilename, , Path, , File
 infFile = CompletePath(Path) & File & ".inf"
 If FileExists(infFile) And Not FileInUse(infFile) Then
  fn = FreeFile
  Open infFile For Binary As #fn
  Get #fn, , ReadInfoSpoolfile
  Close #fn
 End If
End Function

Public Function IsInIDE() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  IsInIDE = App.LogMode <> 1
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "IsInIDE")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub KillInfoSpoolfile(SpoolFilename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim infFile As String, Path As String, File As String
50020  SplitPath SpoolFilename, , Path, , File
50030  infFile = CompletePath(Path) & File & ".inf"
50040  KillFile infFile
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "KillInfoSpoolfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function ScreenResolution() As Integer
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ScreenResolution = GetDeviceCaps(GetDC(0), BITSPIXEL)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "ScreenResolution")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function ResolveEnvironmentApi(ByVal Str1 As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim expStr As String, Length As Long
50020  expStr = ""
50030  Length = ExpandEnvironmentStrings(Str1, expStr, 0)
50040  expStr = String$(Length - 1, 0)
50050  Length = ExpandEnvironmentStrings(Str1, expStr, Length)
50060  ResolveEnvironmentApi = expStr
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "ResolveEnvironmentApi")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function ResolveEnvironment(ByVal Str1 As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, envStr As String, tStr As String
50020  i = 1
50030  Do
50040   envStr = Environ$(i)
50050   If InStr(1, envStr, "=") > 1 Then
50060    tStr = UCase$(Mid(envStr, 1, InStr(1, envStr, "=") - 1))
50070    Str1 = Replace$(Str1, "%" & tStr & "%", Mid(envStr, InStr(1, envStr, "=") + 1), , , vbTextCompare)
50080   End If
50090   DoEvents
50100   i = i + 1
50110  Loop Until envStr = ""
50120  ResolveEnvironment = Str1
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "ResolveEnvironment")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function FormIsModal(f As Form) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  FormIsModal = (GetWindowLong(f.hwnd, GWL_EXSTYLE) And WS_EX_DLGMODALFRAME) = WS_EX_DLGMODALFRAME
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "FormIsModal")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function ExistsAnModalForm() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim f As Form
50020  For Each f In Forms
50030   If FormIsModal(f) Then
50040    ExistsAnModalForm = True
50050    Exit Function
50060   End If
50070  Next f
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "ExistsAnModalForm")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function ResolveRelativePath(RelativePath As String, Optional BasePath As String, Optional PathSeparator As String = "\") As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim nBasePath As String, nBaseParts As Variant, nPathParts As Variant, _
  i As Integer, p As Integer, nPath As String, nResolvedPath As String, _
  nServerRoot As String
50040
50050  If Len(BasePath) Then
50060    nBasePath = BasePath
50070   Else
50080    nBasePath = CurDir
50090  End If
50100  If Right$(nBasePath, 1) = PathSeparator Then
50110   nBasePath = Left$(nBasePath, Len(nBasePath) - 1)
50120  End If
50130  If Left$(nBasePath, 2) = "\\" Then
50140   nBasePath = Mid$(nBasePath, 3)
50150   nServerRoot = "\\"
50160  End If
50170  If LenB(RelativePath) > 0 Then
50180    nPathParts = Split(RelativePath, PathSeparator)
50190    If nPathParts(0) = ".." Then
50200      nBaseParts = Split(nBasePath, PathSeparator)
50210      For i = 0 To UBound(nPathParts)
50220       If nPathParts(i) = ".." Then
50230         p = p + 1
50240        Else
50250         If p Then
50260          nPath = nPath & PathSeparator & nPathParts(i)
50270         End If
50280       End If
50290      Next i
50300      If p > UBound(nBaseParts) Then
50310        Err.Raise rpErrTooManySteps, "modRelativePaths.ResolvePath", "rpErrTooManySteps"
50320       Else
50330        For i = 0 To UBound(nBaseParts) - p
50340         nResolvedPath = nResolvedPath & nBaseParts(i) & PathSeparator
50350        Next i
50360        ResolveRelativePath = nServerRoot & nResolvedPath & Mid$(nPath, 2)
50370      End If
50380     Else
50390      ResolveRelativePath = nServerRoot & nBasePath & PathSeparator & RelativePath
50400    End If
50410   Else
50420    ResolveRelativePath = nBasePath & PathSeparator & RelativePath
50430  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "ResolveRelativePath")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function MakeRelativePath(Path As String, Optional BasePath As String, Optional PathSeparator As String = "\") As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim nPath As String, nPathParts As Variant, nBasePath As String, nBaseParts As Variant, i As Integer, _
  nRelative As Boolean, nRelativePath As String, p As Integer
50030
50040  If Left$(Path, 2) = "\\" Then
50050    nPath = Mid$(Path, 3)
50060   Else
50070    nPath = Path
50080  End If
50090  nPathParts = Split(nPath, PathSeparator)
50100  If Len(BasePath) Then
50110    nBasePath = BasePath
50120   Else
50130    nBasePath = CurDir
50140  End If
50150  If Right$(nBasePath, 1) = PathSeparator Then
50160   nBasePath = Left$(nBasePath, Len(nBasePath) - 1)
50170  End If
50180  If Left$(nBasePath, 2) = "\\" Then
50190   nBasePath = Mid$(nBasePath, 3)
50200  End If
50210  nBaseParts = Split(nBasePath, PathSeparator)
50220  If LCase$(nBaseParts(0)) <> LCase(nPathParts(0)) Then
50230   Err.Raise rpErrDifferentRoot, "modRelativePaths.MakeRelativePath", "rpErrDifferentRoot"
50240  End If
50250  For i = 1 To UBound(nBaseParts)
50260   If nRelative Then
50270     nRelativePath = "..\" & nRelativePath
50280    Else
50290     If LCase$(nBaseParts(i)) <> LCase(nPathParts(i)) Then
50300      nRelative = True
50310      nRelativePath = ".."
50320      For p = i To UBound(nPathParts)
50330       nRelativePath = nRelativePath & PathSeparator & nPathParts(p)
50340      Next p
50350     End If
50360   End If
50370  Next i
50380  If Len(nRelativePath) Then
50390    MakeRelativePath = nRelativePath
50400   Else
50410    MakeRelativePath = nPathParts(UBound(nPathParts))
50420  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "MakeRelativePath")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetDecimalChar() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  GetDecimalChar = Mid$(CStr(1.5), 2, 1)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetDecimalChar")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub CreateTextFile(Filename As String, Str1 As String, Optional OpenmodeAppend As Boolean = False)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long
50020  fn = FreeFile
50030  If OpenmodeAppend Then
50040    Open Filename For Append As #fn
50050   Else
50060    Open Filename For Output As #fn
50070  End If
50080  Print #fn, Str1
50090  Close #fn
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "CreateTextFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function MakeLong(ByVal LOWORD As Integer, ByVal HIWORD As Integer) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010   MakeLong = (HIWORD * &H10000) Or (LOWORD And &HFFFF&)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "MakeLong")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub ShowAcceleratorsInForm(ByRef f As Form, ByVal newValue As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Call SendMessageByNum(f.hwnd, WM_CHANGEUISTATE, _
  MakeLong(IIf(newValue, UIS_CLEAR, UIS_SET), UISF_HIDEACCEL), 0&)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "ShowAcceleratorsInForm")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function IsFileEditable(Filename As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Ext As String, reg As clsRegistry
50020  IsFileEditable = False
50030  If Len(Filename) = 0 Then
50040   Exit Function
50050  End If
50060  SplitPath Filename, , , , , Ext
50070  If Len(Ext) = 0 Then
50080   Exit Function
50090  End If
50100  Set reg = New clsRegistry
50110  reg.hkey = HKEY_LOCAL_MACHINE
50120  reg.KeyRoot = "Software\CLASSES\." & Ext
50130  If reg.KeyExists = False Then
50140   Set reg = Nothing
50150    Exit Function
50160   End If
50170  reg.KeyRoot = "Software\CLASSES\" & reg.GetRegistryValue("") & "\shell\print"
50180  If reg.KeyExists = False Then
50190   Set reg = Nothing
50200   Exit Function
50210  End If
50220  IsFileEditable = True
50230  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "IsFileEditable")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
