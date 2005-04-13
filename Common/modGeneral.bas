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

Public Function ANSItoASCII(ByVal AnsiString As String) As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  CharToOem AnsiString, AnsiString
50050  ANSItoASCII = AnsiString
50060 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50070 Exit Function
ErrPtnr_OnError:
50091 Select Case ErrPtnr.OnError("modGeneral", "ANSItoASCII")
      Case 0: Resume
50110 Case 1: Resume Next
50120 Case 2: Exit Function
50130 Case 3: End
50140 End Select
50150 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function ASCIItoANSI(ByVal AsciiString As String) As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  OemToChar AsciiString, AsciiString
50050  ASCIItoANSI = AsciiString
50060 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50070 Exit Function
ErrPtnr_OnError:
50091 Select Case ErrPtnr.OnError("modGeneral", "ASCIItoANSI")
      Case 0: Resume
50110 Case 1: Resume Next
50120 Case 2: Exit Function
50130 Case 3: End
50140 End Select
50150 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub CombineFiles(ByVal Filename As String, Files As Collection, _
 Optional BufferSize As Long = 65536, Optional stb As StatusBar)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long, j As Long, fnSource As Long, fnDest As Long, sBuffer As String, _
  aLen As Double, tLen As Double, bsize As Long, fpos As Long
50060
50070  bsize = BufferSize
50080  Filename = Trim$(Filename)
50090  If Filename = vbNullString Or Files.Count = 0 Or Right$(Filename, 1) = "\" Then
50100   Exit Sub
50110  End If
50120  If Files.Count = 1 Then
50130   Exit Sub
50140  End If
50150  fnDest = FreeFile
50160  aLen = 0: tLen = 0
50170  For i = 1 To Files.Count
50180   aLen = aLen + FileLen(Files.Item(i))
50190  Next i
50200  Open Filename For Binary As #fnDest
50210  For i = 1 To Files.Count
50220   DoEvents
50230   If FileExists(Files.Item(i)) = False Then
50240    MsgBox "Not found: " & Files.Item(i)
50250   End If
50260   If FileLen(Files.Item(i)) > 0 Then
50270    fnSource = FreeFile
50280    Open Files.Item(i) For Binary Access Read As #fnSource
50290    If bsize > LOF(fnSource) Then
50300     bsize = LOF(fnSource)
50310    End If
50320    fpos = 1
50330    For j = 1 To LOF(fnSource) \ bsize
50340     fpos = (j - 1) * bsize + 1
50350     Seek #fnSource, fpos
50360     sBuffer = Input(bsize, fnSource)
50370     Put #fnDest, , sBuffer
50380     tLen = tLen + bsize
50390     If IsObject(stb) = True Then
50400      stb.Panels("Percent").Text = Format(CDbl(tLen) / CDbl(aLen), "0.0%")
50410     End If
50420     DoEvents
50430    Next j
50440    If LOF(fnSource) > (j - 1) * bsize Then
50450     fpos = (j - 1) * bsize + 1
50460     Seek #fnSource, fpos
50470     sBuffer = Input(LOF(fnSource) - (j - 1) * bsize, fnSource)
50480     Put #fnDest, , sBuffer
50490     tLen = tLen + (LOF(fnSource) - (j - 1) * bsize)
50500    End If
50510    Close #fnSource
50520   End If
50530   DoEvents
50540   KillFile Files.Item(i)
50550   KillInfoSpoolfile Files.Item(i)
50560   DoEvents
50570  Next i
50580  Close #fnDest
50590 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50600 Exit Sub
ErrPtnr_OnError:
50621 Select Case ErrPtnr.OnError("modGeneral", "CombineFiles")
      Case 0: Resume
50640 Case 1: Resume Next
50650 Case 2: Exit Sub
50660 Case 3: End
50670 End Select
50680 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub CombineFilesOld(ByVal Filename As String, Files As Collection, stb As StatusBar)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long, fnSource As Long, fnDest As Long, sBuffer As String, _
  aLen As Double, tLen As Double
50060
50070  Filename = Trim$(Filename)
50080  If Filename = vbNullString Or Files.Count = 0 Or Right$(Filename, 1) = "\" Then
50090   Exit Sub
50100  End If
50110  If FileExists(Filename) = True Then
50120   Exit Sub
50130  End If
50140  If Files.Count = 1 Then
50150   Exit Sub
50160  End If
50170  fnDest = FreeFile
50180  aLen = 0
50190  For i = 1 To Files.Count
50200   aLen = aLen + FileLen(Files.Item(i))
50210  Next i
50220  Open Filename For Binary As #fnDest
50230  For i = 1 To Files.Count
50240   DoEvents
50250   If FileLen(Files.Item(i)) > 0 Then
50260    fnSource = FreeFile
50270    Open Files.Item(i) For Binary Access Read As #fnSource
50280    sBuffer = String(LOF(fnSource), Chr$(0))
50290    Get #fnSource, , sBuffer
50300    Put #fnDest, , sBuffer
50310    Close #fnSource
50320   End If
50330   tLen = tLen + FileLen(Files.Item(i))
50340   stb.Panels("Percent").Text = Format$(tLen / aLen, "0.0%")
50350   KillFile Files.Item(i)
50360   DoEvents
50370  Next i
50380  Close #fnDest
50390  stb.Panels("Percent").Text = vbNullString
50400 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50410 Exit Sub
ErrPtnr_OnError:
50431 Select Case ErrPtnr.OnError("modGeneral", "CombineFilesOld")
      Case 0: Resume
50450 Case 1: Resume Next
50460 Case 2: Exit Sub
50470 Case 3: End
50480 End Select
50490 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub ComboSetListWidth(oCombo As Object, Optional ByVal nFixWidth As Variant, Optional ByVal nScaleMode As Variant)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  With oCombo
50050   If IsMissing(nScaleMode) Or IsMissing(nFixWidth) Then
50060    nScaleMode = .Parent.ScaleMode
50070   End If
50080   If IsMissing(nFixWidth) Then
50090    Dim i As Long, nWidth As Long
50100    nFixWidth = 0
50110    For i = 0 To .ListCount - 1
50120     nWidth = .Parent.TextWidth(.List(i))
50130     If nWidth > nFixWidth Then
50140      nFixWidth = nWidth
50150     End If
50160    Next i
50170    nFixWidth = nFixWidth + .Parent.ScaleX(10, vbPixels, nScaleMode)
50180    If .ListCount > 8 Then
50190     nFixWidth = nFixWidth + .Parent.ScaleX(15, vbPixels, nScaleMode)
50200    End If
50210   End If
50220   SendMessage .hwnd, CB_SETDROPPEDWIDTH, .Parent.ScaleX(nFixWidth, nScaleMode, vbPixels), 0&
50230  End With
50240 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50250 Exit Sub
ErrPtnr_OnError:
50271 Select Case ErrPtnr.OnError("modGeneral", "ComboSetListWidth")
      Case 0: Resume
50290 Case 1: Resume Next
50300 Case 2: Exit Sub
50310 Case 3: End
50320 End Select
50330 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function CompletePath(Path As String) As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If Len(Path) = 0 Then
50050   Exit Function
50060  End If
50070  Path = Trim$(Path)
50080  If Right$(Path, 1) = "\" Then
50090    CompletePath = LTrim$(Path)
50100   Else
50110    CompletePath = LTrim$(Path) & "\"
50120  End If
50130 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50140 Exit Function
ErrPtnr_OnError:
50161 Select Case ErrPtnr.OnError("modGeneral", "CompletePath")
      Case 0: Resume
50180 Case 1: Resume Next
50190 Case 2: Exit Function
50200 Case 3: End
50210 End Select
50220 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function CreateDir(Path As String) As Boolean
 On Error GoTo ErrorHandler
 MkDir Path
 CreateDir = True
Exit Function
ErrorHandler:
 CreateDir = False
End Function

Public Function DirExists(DirStr As String) As Boolean
 On Error GoTo ErrorHandler
 DirExists = GetAttr(DirStr) Or vbDirectory
 Exit Function
ErrorHandler:
 DirExists = False
End Function

Public Function FileExists(FileStr As String) As Boolean
 On Error GoTo ErrorHandler
 FileExists = GetAttr(FileStr)
 Exit Function
ErrorHandler:
 FileExists = False
End Function

Public Function IsFilePrintable(Filename As String) As Boolean
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim Ext As String, reg As clsRegistry
50050  IsFilePrintable = False
50060  If Len(Filename) = 0 Then
50070   Exit Function
50080  End If
50090  SplitPath Filename, , , , , Ext
50100  If Len(Ext) = 0 Then
50110   Exit Function
50120  End If
50130  Set reg = New clsRegistry
50140  reg.hkey = HKEY_LOCAL_MACHINE
50150  reg.KeyRoot = "Software\CLASSES\." & Ext
50160  If reg.KeyExists = False Then
50170   Set reg = Nothing
50180   Exit Function
50190  End If
50200  reg.KeyRoot = "Software\CLASSES\" & reg.GetRegistryValue("") & "\shell\print"
50210  If reg.KeyExists = False Then
50220   Set reg = Nothing
50230   Exit Function
50240  End If
50250  IsFilePrintable = True
50260 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50270 Exit Function
ErrPtnr_OnError:
50291 Select Case ErrPtnr.OnError("modGeneral", "IsFilePrintable")
      Case 0: Resume
50310 Case 1: Resume Next
50320 Case 2: Exit Function
50330 Case 3: End
50340 End Select
50350 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetComputerName() As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Const MAX_COMPUTERNAME_LENGTH As Long = 31
50050  Dim tstr As String
50060  tstr = String(MAX_COMPUTERNAME_LENGTH + 1, Chr$(0))
50070  GetComputerNameA tstr, MAX_COMPUTERNAME_LENGTH + 1
50080  GetComputerName = Left$(tstr, InStr(tstr, Chr$(0)) - 1)
50090 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50100 Exit Function
ErrPtnr_OnError:
50121 Select Case ErrPtnr.OnError("modGeneral", "GetComputerName")
      Case 0: Resume
50140 Case 1: Resume Next
50150 Case 2: Exit Function
50160 Case 3: End
50170 End Select
50180 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetDesktop() As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim TempDesktop As String, reg As clsRegistry
50050  Set reg = New clsRegistry
50060  reg.hkey = HKEY_CURRENT_USER
50070  reg.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
50080  If IsWin9xMe = True Then
50090    TempDesktop = reg.GetRegistryValue("Common Desktop")
50100   Else
50110    TempDesktop = reg.GetRegistryValue("Desktop")
50120  End If
50130  Set reg = Nothing
50140  If Trim$(TempDesktop) = vbNullString Then
50150   TempDesktop = "C:\"
50160  End If
50170  GetDesktop = TempDesktop
50180 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50190 Exit Function
ErrPtnr_OnError:
50211 Select Case ErrPtnr.OnError("modGeneral", "GetDesktop")
      Case 0: Resume
50230 Case 1: Resume Next
50240 Case 2: Exit Function
50250 Case 3: End
50260 End Select
50270 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetDrives() As Collection
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim l As Long, Buffer As String, res As Long, drives As String, drv() As String, _
  i As Long
50060  l = 64: Buffer = Space(l)
50070  res = GetLogicalDriveStrings(l, Buffer)
50080  drives = Left$(Buffer, res)
50090  Set GetDrives = New Collection
50100  If Len(drives) > 0 Then
50110   If InStr(drives, Chr$(0)) > 0 Then
50120     drv = Split(drives, Chr$(0))
50130     For i = LBound(drv) To UBound(drv)
50140      If Trim$(drv(i)) <> vbNullString Then
50150       GetDrives.Add drv(i)
50160      End If
50170     Next i
50180    Else
50190     GetDrives.Add drives
50200   End If
50210  End If
50220 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50230 Exit Function
ErrPtnr_OnError:
50251 Select Case ErrPtnr.OnError("modGeneral", "GetDrives")
      Case 0: Resume
50270 Case 1: Resume Next
50280 Case 2: Exit Function
50290 Case 3: End
50300 End Select
50310 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetFileAttributesStr(Filename As String) As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim hFind As Long, WFD As WIN32_FIND_DATA, attr As Long, AA As String
50050  hFind = FindFirstFileA(Filename, WFD)
50060  attr = WFD.dwFileAttributes
50070  If attr And FILE_ATTRIBUTE_ARCHIVE Then AA = AA & "A"
50080  If attr And FILE_ATTRIBUTE_COMPRESSED Then AA = AA & "C"
50090  If attr And FILE_ATTRIBUTE_DIRECTORY Then AA = AA & "D"
50100  If attr And FILE_ATTRIBUTE_HIDDEN Then AA = AA & "H"
50110  If attr And FILE_ATTRIBUTE_NORMAL Then AA = AA & "N"
50120  If attr And FILE_ATTRIBUTE_READONLY Then AA = AA & "R"
50130  If attr And FILE_ATTRIBUTE_SYSTEM Then AA = AA & "S"
50140  GetFileAttributesStr = AA
50150 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50160 Exit Function
ErrPtnr_OnError:
50181 Select Case ErrPtnr.OnError("modGeneral", "GetFileAttributesStr")
      Case 0: Resume
50200 Case 1: Resume Next
50210 Case 2: Exit Function
50220 Case 3: End
50230 End Select
50240 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetFiles(ByVal Path As String, Optional Searchmask As String = "*.*", _
 Optional Sorted As eSortModeFiles = notSorted) As Collection
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim tColl1 As Collection, tColl2 As Collection, tFilename As String, _
  i As Long, tStrf() As String
50060  If Len(Searchmask) > 0 Then
50070    tFilename = Dir(CompletePath(Trim$(Path)) & Searchmask)
50080   Else
50090    tFilename = Dir(Path)
50100    SplitPath Path, , Path
50110    Path = CompletePath(Path)
50120  End If
50130  Set tColl1 = New Collection
50140  Do While tFilename <> ""
50151   Select Case Sorted
         Case eSortModeFiles.SortedByDate
50170     AddSortedStr tColl1, Format$(FileDateTime(Path & tFilename), "yyyymmddhhnnss") & "|" & Path & tFilename
50180    Case eSortModeFiles.SortedByName
50190     AddSortedStr tColl1, "|" & Path & tFilename
50200    Case Else
50210     tColl1.Add "|" & Path & tFilename
50220   End Select
50230   tFilename = Dir()
50240   DoEvents
50250  Loop
50260  Set tColl2 = New Collection
50270  For i = 1 To tColl1.Count
50280   tStrf = Split(tColl1(i), "|")
50290   SplitPath tStrf(1), , Path, tFilename
50300   Path = CompletePath(Path)
50310   tColl2.Add Path & "|" & Path & tFilename & "|" & FileLen(Path & tFilename) & "|" & FileDateTime(Path & tFilename)
50320  Next i
50330  Set GetFiles = tColl2
50340 ' Set tColl = Nothing
50350 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50360 Exit Function
ErrPtnr_OnError:
50381 Select Case ErrPtnr.OnError("modGeneral", "GetFiles")
      Case 0: Resume
50400 Case 1: Resume Next
50410 Case 2: Exit Function
50420 Case 3: End
50430 End Select
50440 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim TempMyFiles, reg As clsRegistry, tstr As String
50050  Set reg = New clsRegistry
50060  With reg
50070   .hkey = HKEY_CURRENT_USER
50080   .KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer"
50090   .Subkey = "Shell Folders"
50100   tstr = .GetRegistryValue("Fonts")
50110  End With
50120  Set reg = Nothing
50130  If LenB(tstr) = 0 Then
50140   tstr = CompletePath(GetWindowsDirectory) & "Fonts"
50150  End If
50160  GetFontsDirectory = tstr
50170 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50180 Exit Function
ErrPtnr_OnError:
50201 Select Case ErrPtnr.OnError("modGeneral", "GetFontsDirectory")
      Case 0: Resume
50220 Case 1: Resume Next
50230 Case 2: Exit Function
50240 Case 3: End
50250 End Select
50260 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim lRetVal As Long, sShortPathName As String, iLen As Integer
50050  sShortPathName = Space(255)
50060  iLen = Len(sShortPathName)
50070  lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
50080  GetShortName = Left(sShortPathName, lRetVal)
50090 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50100 Exit Function
ErrPtnr_OnError:
50121 Select Case ErrPtnr.OnError("modGeneral", "GetShortName")
      Case 0: Resume
50140 Case 1: Resume Next
50150 Case 2: Exit Function
50160 Case 3: End
50170 End Select
50180 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetStringRessource(DLLExeFilename As String, ResID As Long) As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim hLib As Long, sBuf As String, lLen As Long
50050  hLib = LoadLibrary(DLLExeFilename)
50060  If hLib Then
50070   sBuf = Space$(255)
50080   lLen = LoadString(hLib, ResID, sBuf, Len(sBuf))
50090   If lLen Then
50100     sBuf = Left$(sBuf, lLen)
50110    Else
50120     sBuf = vbNullString
50130   End If
50140   FreeLibrary hLib
50150  End If
50160  GetStringRessource = sBuf
50170 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50180 Exit Function
ErrPtnr_OnError:
50201 Select Case ErrPtnr.OnError("modGeneral", "GetStringRessource")
      Case 0: Resume
50220 Case 1: Resume Next
50230 Case 2: Exit Function
50240 Case 3: End
50250 End Select
50260 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetSystemDirectory() As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim res As Long, Sysdir As String
50050  Sysdir = Space$(MAX_PATH)
50060  res = GetSystemDirectoryA(Sysdir, MAX_PATH)
50070  If res > 0 Then
50080    GetSystemDirectory = Left$(Sysdir, res)
50090   Else
50100    GetSystemDirectory = "C:\Windows"
50110  End If
50120 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50130 Exit Function
ErrPtnr_OnError:
50151 Select Case ErrPtnr.OnError("modGeneral", "GetSystemDirectory")
      Case 0: Resume
50170 Case 1: Resume Next
50180 Case 2: Exit Function
50190 Case 3: End
50200 End Select
50210 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetTempFile(Optional ByVal Path As String, Optional Prefix As String) As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim res As Long, Tempfile As String, tPath As String
50050  Tempfile = Space$(MAX_PATH)
50060  tPath = LTrim$(Path)
50070  If DirExists(tPath) = False Then
50080   MakePath tPath
50090  End If
50100
50110  res = GetTempFileNameA(tPath, Prefix, 0, Tempfile)
50120  If res <> 0 Then
50130    GetTempFile = Left$(Tempfile, InStr(Tempfile, Chr$(0)) - 1)
50140   Else
50150    GetTempFile = "~.tmp"
50160  End If
50170 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50180 Exit Function
ErrPtnr_OnError:
50201 Select Case ErrPtnr.OnError("modGeneral", "GetTempFile")
      Case 0: Resume
50220 Case 1: Resume Next
50230 Case 2: Exit Function
50240 Case 3: End
50250 End Select
50260 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetTempPath() As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim res As Long, TempDir As String
50050
50060  If IsWin9xMe = True Then
50070    ' In Win9xMe you cannot use the username because the username can
50080    ' have forbidden chars like \/:*?"<>|
50090    TempDir = CompletePath(GetTempPathApi) & "PDFCreator\"
50100   Else
50110    If IsWinNT4 = True Then
50120      TempDir = GetMyAppData
50130      If LenB(Environ$("Redmon_User")) > 0 Then
50140        TempDir = CompletePath(TempDir) & "PDFCreator\" & Environ$("Redmon_User")
50150       Else
50160        TempDir = CompletePath(TempDir) & "PDFCreator\" & GetUsername
50170      End If
50180     Else
50190      TempDir = GetMyLocalAppData
50200      TempDir = Left(TempDir, InStrRev(TempDir, "\")) & "Temp\PDFCreator\"
50210    End If
50220  End If
50230  If Trim$(TempDir) = vbNullString Then
50240   TempDir = CompletePath(App.Path) & "Temp\"
50250  End If
50260
50270  GetTempPath = CompletePath(TempDir)
50280 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50290 Exit Function
ErrPtnr_OnError:
50311 Select Case ErrPtnr.OnError("modGeneral", "GetTempPath")
      Case 0: Resume
50330 Case 1: Resume Next
50340 Case 2: Exit Function
50350 Case 3: End
50360 End Select
50370 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetTempPathApi() As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim res As Long, TempDir As String
50050
50060  TempDir = Space$(MAX_PATH)
50070  res = GetTempPathA(MAX_PATH, TempDir)
50080  If res > 0 Then
50090    GetTempPathApi = Left$(TempDir, res)
50100   Else
50110    GetTempPathApi = "C:\"
50120  End If
50130 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50140 Exit Function
ErrPtnr_OnError:
50161 Select Case ErrPtnr.OnError("modGeneral", "GetTempPathApi")
      Case 0: Resume
50180 Case 1: Resume Next
50190 Case 2: Exit Function
50200 Case 3: End
50210 End Select
50220 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetUsername() As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Const MAX_USERNAME_LENGTH = 128
50050  Dim tstr As String
50060  tstr = String(MAX_USERNAME_LENGTH, Chr$(0))
50070  GetUserNameA tstr, MAX_USERNAME_LENGTH
50080  GetUsername = Left$(tstr, InStr(tstr, Chr$(0)) - 1)
50090 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50100 Exit Function
ErrPtnr_OnError:
50121 Select Case ErrPtnr.OnError("modGeneral", "GetUsername")
      Case 0: Resume
50140 Case 1: Resume Next
50150 Case 2: Exit Function
50160 Case 3: End
50170 End Select
50180 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetWindowsDirectory() As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim res As Long, WinDir As String
50050  WinDir = Space$(MAX_PATH)
50060  res = GetWindowsDirectoryA(WinDir, MAX_PATH)
50070  If res > 0 Then
50080    GetWindowsDirectory = Left$(WinDir, res)
50090   Else
50100    GetWindowsDirectory = "C:\Windows"
50110  End If
50120 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50130 Exit Function
ErrPtnr_OnError:
50151 Select Case ErrPtnr.OnError("modGeneral", "GetWindowsDirectory")
      Case 0: Resume
50170 Case 1: Resume Next
50180 Case 2: Exit Function
50190 Case 3: End
50200 End Select
50210 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long, Forbiddenchars As String
50050  IsForbiddenChars = False
50060  Forbiddenchars = "\/:*?<>|%"""
50070  For i = 1 To Len(Forbiddenchars)
50080   If InStr(chkStr, Mid$(Forbiddenchars, i, 1)) > 0 Then
50090    IsForbiddenChars = True
50100    Exit Function
50110   End If
50120  Next i
50130 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50140 Exit Function
ErrPtnr_OnError:
50161 Select Case ErrPtnr.OnError("modGeneral", "IsForbiddenChars")
      Case 0: Resume
50180 Case 1: Resume Next
50190 Case 2: Exit Function
50200 Case 3: End
50210 End Select
50220 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsFormLoaded(Form As Form) As Boolean
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim nForm As Form
50050  For Each nForm In Forms
50060   If nForm Is Form Then
50070    IsFormLoaded = True
50080    Exit For
50090   End If
50100  Next
50110 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50120 Exit Function
ErrPtnr_OnError:
50141 Select Case ErrPtnr.OnError("modGeneral", "IsFormLoaded")
      Case 0: Resume
50160 Case 1: Resume Next
50170 Case 2: Exit Function
50180 Case 3: End
50190 End Select
50200 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  LoadDLL = LoadLibrary(DLLPath)
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Function
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("modGeneral", "LoadDLL")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Function
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function LoadIcon(ByVal lngSize As IconSize, ByVal strPath As String, Icon As IPictureDisp) As Long
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim iuUnkown As IUnknown, itIcon As IconType, CLSID As CLSIdType, _
  sfiShellInfo As ShellFileInfoType
50060  Call SHGetFileInfo(strPath, 0&, sfiShellInfo, Len(sfiShellInfo), SHGFI_USEFILEATTRIBUTES Or lngSize)
50070  With itIcon
50080   .cbSize = Len(itIcon)
50090   .picType = vbPicTypeIcon
50100   .hIcon = sfiShellInfo.hIcon
50110  End With
50120  CLSID.ID(8) = &HC0
50130  CLSID.ID(15) = &H46
50140  Call OleCreatePictureIndirect(itIcon, CLSID, 1&, iuUnkown)
50150  Set Icon = iuUnkown
50160  If Icon Is Nothing Then
50170    LoadIcon = 1
50180   Else
50190    LoadIcon = 0
50200  End If
50210 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50220 Exit Function
ErrPtnr_OnError:
50241 Select Case ErrPtnr.OnError("modGeneral", "LoadIcon")
      Case 0: Resume
50260 Case 1: Resume Next
50270 Case 2: Exit Function
50280 Case 3: End
50290 End Select
50300 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function MakePath(Path As String) As Boolean
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim dirs() As String, tstr As String, i As Long
50050  MakePath = False
50060  Path = LTrim$(Path)
50070  If Len(Path) <= 3 Then
50080   Exit Function
50090  End If
50100  Path = Replace$(Path, "/", "\", , , vbTextCompare)
50110  If InStr(1, Path, "\", vbTextCompare) = 0 Then
50120   Exit Function
50130  End If
50140  dirs = Split(Path, "\")
50150  If Mid$(Path, 2, 1) = ":" Then ' local drive
50160   If UCase$(Mid$(Path, 1, 1)) < "A" Or UCase$(Mid$(Path, 1, 1)) > "Z" Then
50170    Exit Function
50180   End If
50190   tstr = dirs(0)
50200   For i = 1 To UBound(dirs)
50210    tstr = tstr & "\" & dirs(i)
50220    If DirExists(tstr) = False Then
50230     Call CreateDir(tstr)
50240     If DirExists(tstr) = False Then
50250      Exit Function
50260     End If
50270    End If
50280   Next i
50290   MakePath = True
50300  End If
50310  If Mid(Path, 1, 2) = "\\" Then ' network share
50320   If UBound(dirs) < 4 Then
50330    Exit Function
50340   End If
50350   tstr = "\\" & dirs(2) & "\" & dirs(3)
50360   For i = 4 To UBound(dirs)
50370    tstr = tstr & "\" & dirs(i)
50380    If DirExists(tstr) = False Then
50390     Call CreateDir(tstr)
50400     If DirExists(tstr) = False Then
50410      Exit Function
50420     End If
50430    End If
50440   Next i
50450   MakePath = True
50460  End If
50470 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50480 Exit Function
ErrPtnr_OnError:
50501 Select Case ErrPtnr.OnError("modGeneral", "MakePath")
      Case 0: Resume
50520 Case 1: Resume Next
50530 Case 2: Exit Function
50540 Case 3: End
50550 End Select
50560 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub OpenDocument(sFilename As String)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim sDirectory As String, res As Long, handle As Long
50050
50060  handle = GetDesktopWindow()
50070  res = ShellExecute(handle, "open", sFilename, _
    vbNullString, vbNullString, vbNormalFocus)
50090
50100  If res = SE_ERR_NOTFOUND Then
50110   ElseIf res = SE_ERR_NOASSOC Then
50120     sDirectory = Space$(260)
50130     sDirectory = CompletePath(GetSystemDirectory)
50140     Call ShellExecute(handle, vbNullString, _
      "RUNDLL32.EXE", "shell32.dll,OpenAs_RunDLL " & _
      sFilename, sDirectory, vbNormalFocus)
50170   End If
50180 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50190 Exit Sub
ErrPtnr_OnError:
50211 Select Case ErrPtnr.OnError("modGeneral", "OpenDocument")
      Case 0: Resume
50230 Case 1: Resume Next
50240 Case 2: Exit Sub
50250 Case 3: End
50260 End Select
50270 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub RemoveX(Frm As Form)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim hMenu As Long, nPosition As Long
50050
50060  hMenu = GetSystemMenu(Frm.hwnd, 0)
50070  If hMenu <> 0 Then
50080   nPosition = GetMenuItemCount(hMenu)
50090   Call RemoveMenu(hMenu, nPosition - 1, MF_REMOVE Or MF_BYPOSITION)
50100   Call RemoveMenu(hMenu, nPosition - 2, MF_REMOVE Or MF_BYPOSITION)
50110   Call DrawMenuBar(Frm.hwnd)
50120  End If
50130 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50140 Exit Sub
ErrPtnr_OnError:
50161 Select Case ErrPtnr.OnError("modGeneral", "RemoveX")
      Case 0: Resume
50180 Case 1: Resume Next
50190 Case 2: Exit Sub
50200 Case 3: End
50210 End Select
50220 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function ReplaceForbiddenChars(chkStr As String, Optional ReplaceChar As String = "_", Optional AdditionalForbiddenChars = "") As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim tstr As String, i As Long, Forbiddenchars As String
50050  Forbiddenchars = "\/:*?<>|""" & AdditionalForbiddenChars
50060  tstr = chkStr
50070  For i = 1 To Len(Forbiddenchars)
50080   tstr = Replace$(tstr, Mid$(Forbiddenchars, i, 1), ReplaceChar)
50090  Next i
50100  ReplaceForbiddenChars = tstr
50110 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50120 Exit Function
ErrPtnr_OnError:
50141 Select Case ErrPtnr.OnError("modGeneral", "ReplaceForbiddenChars")
      Case 0: Resume
50160 Case 1: Resume Next
50170 Case 2: Exit Function
50180 Case 3: End
50190 End Select
50200 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function RunProgramWait(strCmdLine As String, Optional ShowWindow As Boolean = True) As Long
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim lngRetCode As Long, proc As PROCESS_INFORMATION, Start As STARTUPINFO, _
  secAttr As SECURITY_ATTRIBUTES, lngRet As Long, lngExit As Long
50060
50070  Start.cb = Len(Start)
50080  If ShowWindow = False Then
50090   Start.dwFlags = STARTF_USESHOWWINDOW
50100   Start.wShowWindow = SW_HIDE
50110  End If
50120
50130  lngRet = CreateProcessA(0&, strCmdLine, _
  0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, Start, proc)
50150
50160  GetExitCodeProcess proc.hProcess, lngExit
50170
50180  Do While lngExit = STILL_ACTIVE
50190   DoEvents
50200   Sleep 100
50210   GetExitCodeProcess proc.hProcess, lngExit
50220  Loop
50230  lngRet = CloseHandle(proc.hProcess)
50240
50250  RunProgramWait = lngExit
50260 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50270 Exit Function
ErrPtnr_OnError:
50291 Select Case ErrPtnr.OnError("modGeneral", "RunProgramWait")
      Case 0: Resume
50310 Case 1: Resume Next
50320 Case 2: Exit Function
50330 Case 3: End
50340 End Select
50350 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub SetFont(Frm As Form, ByVal Fontname As String, ByVal Charset As Long, ByVal Fontsize As Long)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim ctl As Control
50050
50060  For Each ctl In Frm.Controls
50070   If TypeOf ctl Is Label Or _
     TypeOf ctl Is Form Or _
     TypeOf ctl Is ComboBox Or _
     TypeOf ctl Is CheckBox Or _
     TypeOf ctl Is CommandButton Or _
     TypeOf ctl Is ListView Or _
     TypeOf ctl Is StatusBar Or _
     TypeOf ctl Is TextBox Or _
     TypeOf ctl Is Frame Then
50160    With ctl
50170     .Font = Fontname
50180     If Not (TypeOf ctl Is StatusBar) And Not (TypeOf ctl Is ListView) Then
50190      .Fontsize = Fontsize
50200     End If
50210     .Font.Charset = Charset
50220    End With
50230   End If
50240  Next ctl
50250 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50260 Exit Sub
ErrPtnr_OnError:
50281 Select Case ErrPtnr.OnError("modGeneral", "SetFont")
      Case 0: Resume
50300 Case 1: Resume Next
50310 Case 2: Exit Sub
50320 Case 3: End
50330 End Select
50340 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SplitPath(FullPath As String, Optional Drive As String, Optional Path As String, Optional Filename As String, Optional File As String, Optional Extension As String)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim nPos As Integer
50050  nPos = InStrRev(FullPath, "\")
50060  If nPos > 0 Then
50070    If Left$(FullPath, 2) = "\\" Then
50080     If nPos = 2 Then
50090      Drive = FullPath: Path = vbNullString: Filename = vbNullString: File = vbNullString
50100      Extension = vbNullString
50110      Exit Sub
50120     End If
50130    End If
50140    Path = Left$(FullPath, nPos - 1)
50150    Filename = Mid$(FullPath, nPos + 1)
50160    nPos = InStrRev(Filename, ".")
50170    If nPos > 0 Then
50180      File = Left$(Filename, nPos - 1)
50190      Extension = Mid$(Filename, nPos + 1)
50200     Else
50210      File = Filename
50220      Extension = vbNullString
50230    End If
50240   Else
50250    nPos = InStrRev(FullPath, ":")
50260    If nPos > 0 Then
50270      Path = Mid(FullPath, 1, nPos - 1): Filename = Mid(FullPath, nPos + 1)
50280      nPos = InStrRev(Filename, ".")
50290      If nPos > 0 Then
50300        File = Left$(Filename, nPos - 1)
50310        Extension = Mid$(Filename, nPos + 1)
50320       Else
50330        File = Filename
50340        Extension = vbNullString
50350      End If
50360     Else
50370      Path = vbNullString: Filename = FullPath
50380      nPos = InStrRev(Filename, ".")
50390      If nPos > 0 Then
50400        File = Left$(Filename, nPos - 1)
50410        Extension = Mid$(Filename, nPos + 1)
50420       Else
50430        File = Filename
50440        Extension = vbNullString
50450      End If
50460    End If
50470  End If
50480  If Left$(Path, 2) = "\\" Then
50490    nPos = InStr(3, Path, "\")
50500    If nPos Then
50510      Drive = Left$(Path, nPos - 1)
50520     Else
50530      Drive = Path
50540    End If
50550   Else
50560    If Len(Path) = 2 Then
50570     If Right$(Path, 1) = ":" Then
50580      Path = Path & "\"
50590     End If
50600    End If
50610    If Mid$(Path, 2, 2) = ":\" Then
50620     Drive = Left$(Path, 2)
50630    End If
50640  End If
50650 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50660 Exit Sub
ErrPtnr_OnError:
50681 Select Case ErrPtnr.OnError("modGeneral", "SplitPath")
      Case 0: Resume
50700 Case 1: Resume Next
50710 Case 2: Exit Sub
50720 Case 3: End
50730 End Select
50740 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub MoveMouseToCommandButton(cmdButton As Object)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim pt As POINTAPI, cRect As Rect
50050  GetCursorPos pt
50060  GetWindowRect cmdButton.hwnd, cRect
50070  With cRect
50080   pt.x = (.Left + .Right) / 2
50090   pt.Y = (.Top + .Bottom) / 2
50100  End With
50110  ScreenToAbsolute pt
50120  mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_MOVE, pt.x, pt.Y, 0, GetMessageExtraInfo()
50130 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50140 Exit Sub
ErrPtnr_OnError:
50161 Select Case ErrPtnr.OnError("modGeneral", "MoveMouseToCommandButton")
      Case 0: Resume
50180 Case 1: Resume Next
50190 Case 2: Exit Sub
50200 Case 3: End
50210 End Select
50220 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function IsValidPath(Path As String, Optional ByVal TestUNCPaths As Boolean = True) As Boolean
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long, c As Long, nBytes() As Byte, nChar As String, nOK As Boolean, _
  nParentPath As String, nPos As Long, nRoot As String, nPath As String, _
  nPathParts() As String, nStart As Long
50070
50080  Const kCharAsterisk = 42, kCharBackSlash = 92, kCharColon = 58, _
  kCharGreaterThan = 62, kCharLowerThan = 60, kCharPipe = 124, _
  kCharQuestion = 63, kCharQuote = 34, kCharSlash = 47
50110
50120  nPath = Path
50130  PathUnquoteSpaces nPath
50140  nPath = Left$(nPath, InStrNullChar(nPath))
50150  If Not CBool(PathIsUNC(nPath) And Not TestUNCPaths) Then
50160   If CBool(PathFileExists(nPath)) Then
50170    IsValidPath = True
50180    Exit Function
50190   End If
50200  End If
50210  nRoot = nPath
50220  If PathStripToRoot(nRoot) Then
50230   nRoot = Left$(nRoot, InStrNullChar(nRoot))
50240   If PathIsUNC(nRoot) Then
50250     nPos = InStrRev(nRoot, "\")
50260     nRoot = Left$(nRoot, nPos - 1)
50270     If PathIsUNCServer(nRoot) Then
50280      nPath = Mid$(nPath, Len(nRoot) + 1)
50290     End If
50300    Else
50310     nPath = Mid$(nPath, Len(nRoot) + 1)
50320   End If
50330   If Len(nPath) Then
50340    nPathParts = Split(nPath, "\")
50350    If Len(nPathParts(0)) Then
50360      nStart = 0
50370     Else
50380      nStart = 1
50390    End If
50400    For i = nStart To UBound(nPathParts)
50410     nBytes = StrConv(nPathParts(i), vbFromUnicode)
50420     For c = 0 To UBound(nBytes)
50431      Select Case nBytes(c)
            Case Is < 32, Is > 255, kCharAsterisk, kCharBackSlash, kCharColon, kCharGreaterThan, kCharLowerThan, kCharPipe, kCharQuestion, kCharQuote, kCharSlash
50450        nOK = False
50460        Exit For
50470       Case Else
50480        nOK = True
50490      End Select
50500     Next c
50510     If Not nOK Then
50520      Exit For
50530     End If
50540    Next i
50550    IsValidPath = nOK
50560   End If
50570  End If
50580 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50590 Exit Function
ErrPtnr_OnError:
50611 Select Case ErrPtnr.OnError("modGeneral", "IsValidPath")
      Case 0: Resume
50630 Case 1: Resume Next
50640 Case 2: Exit Function
50650 Case 3: End
50660 End Select
50670 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub FormInTaskbar(Form As Form, ByVal ShowInTaskBar As Boolean, Optional ByVal NoFlicker As Boolean)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim nStyle As Long, nVisible As Boolean
50050  Const GWL_EXSTYLE = (-20), WS_EX_APPWINDOW = &H40000
50060  With Form
50070   nVisible = .Visible
50080   If NoFlicker And nVisible Then
50090    LockWindowUpdate GetDesktopWindow()
50100   End If
50110   .Visible = False
50120   nStyle = GetWindowLong(.hwnd, GWL_EXSTYLE)
50131   Select Case ShowInTaskBar
         Case True
50150     nStyle = nStyle Or WS_EX_APPWINDOW
50160    Case False
50170     nStyle = nStyle And Not WS_EX_APPWINDOW
50180   End Select
50190   SetWindowLong .hwnd, GWL_EXSTYLE, nStyle
50200   .Refresh
50210   .Visible = nVisible
50220   If NoFlicker And nVisible Then
50230    LockWindowUpdate 0&
50240   End If
50250  End With
50260 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50270 Exit Sub
ErrPtnr_OnError:
50291 Select Case ErrPtnr.OnError("modGeneral", "FormInTaskbar")
      Case 0: Resume
50310 Case 1: Resume Next
50320 Case 2: Exit Sub
50330 Case 3: End
50340 End Select
50350 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function RemoveCompletePath(Path As String, Optional ShowDialog As Boolean) As Long
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim ret As Long, SHFILEOP As SHFILEOPSTRUCT, ShowD As Long, FOBuf() As Byte, _
  LenSHFileOp As Long
50060
50070  LenSHFileOp = LenB(SHFILEOP)
50080  ReDim FOBuf(1 To LenSHFileOp)
50090  If ShowDialog Then
50100    ShowD = 0
50110   Else
50120    ShowD = FOF_SILENT
50130  End If
50140
50150  With SHFILEOP
50160   .wFunc = FO_DELETE
50170   .pFrom = Path
50180   .pTo = vbNullString
50190   .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR Or ShowD
50200  End With
50210
50220  ret = SHFileOperation(SHFILEOP)
50230  DoEvents
50240
50250  Call MoveMemory(FOBuf(1), SHFILEOP, LenSHFileOp)
50260  Call MoveMemory(FOBuf(21), FOBuf(19), 12)
50270  Call MoveMemory(SHFILEOP, FOBuf(1), LenSHFileOp)
50280
50290  If ret <> 0 Then
50300   If SHFILEOP.fAnyOperationAborted <> 0 Then
50310    ret = -1
50320   End If
50330  End If
50340  RemoveCompletePath = ret
50350 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50360 Exit Function
ErrPtnr_OnError:
50381 Select Case ErrPtnr.OnError("modGeneral", "RemoveCompletePath")
      Case 0: Resume
50400 Case 1: Resume Next
50410 Case 2: Exit Function
50420 Case 3: End
50430 End Select
50440 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function RemoveLeadingAndTrailingQuotes(tstr As String) As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If Len(tstr) > 0 Then
50050   If Mid$(tstr, 1, 1) = """" Then
50060    tstr = Mid$(tstr, 2)
50070   End If
50080   If Len(tstr) > 0 Then
50090    If Mid$(tstr, Len(tstr), 1) = """" Then
50100     tstr = Mid$(tstr, 1, Len(tstr) - 1)
50110    End If
50120   End If
50130  End If
50140  RemoveLeadingAndTrailingQuotes = tstr
50150 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50160 Exit Function
ErrPtnr_OnError:
50181 Select Case ErrPtnr.OnError("modGeneral", "RemoveLeadingAndTrailingQuotes")
      Case 0: Resume
50200 Case 1: Resume Next
50210 Case 2: Exit Function
50220 Case 3: End
50230 End Select
50240 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub ScreenToAbsolute(lpPoint As POINTAPI)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  lpPoint.x = lpPoint.x * (&HFFFF& / GetSystemMetrics(SM_CXSCREEN))
50050  lpPoint.Y = lpPoint.Y * (&HFFFF& / GetSystemMetrics(SM_CYSCREEN))
50060 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50070 Exit Sub
ErrPtnr_OnError:
50091 Select Case ErrPtnr.OnError("modGeneral", "ScreenToAbsolute")
      Case 0: Resume
50110 Case 1: Resume Next
50120 Case 2: Exit Sub
50130 Case 3: End
50140 End Select
50150 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetOptimalComboboxHeigth(cmb As ComboBox, ParentForm As Form)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim pt As POINTAPI, rc As Rect, cWidth As Long, OldScaleMode As Long
50050
50060  OldScaleMode = ParentForm.ScaleMode
50070  ParentForm.ScaleMode = vbPixels
50080
50090  cWidth = cmb.Width
50100  Call GetWindowRect(cmb.hwnd, rc)
50110  pt.x = rc.Left
50120  pt.Y = rc.Top
50130  Call ScreenToClient(ParentForm.hwnd, pt)
50140  Call MoveWindow(cmb.hwnd, pt.x, pt.Y, cmb.Width, _
  SendMessage(cmb.hwnd, CB_GETITEMHEIGHT, 0, ByVal 0) * (cmb.ListCount + 2), _
  True)
50170  ParentForm.ScaleMode = OldScaleMode
50180 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50190 Exit Sub
ErrPtnr_OnError:
50211 Select Case ErrPtnr.OnError("modGeneral", "SetOptimalComboboxHeigth")
      Case 0: Resume
50230 Case 1: Resume Next
50240 Case 2: Exit Sub
50250 Case 3: End
50260 End Select
50270 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function UnLoadDLL(DLLHandle As Long) As Boolean
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim res As Long
50050  If DLLHandle <> 0 Then
50060   res = FreeLibrary(DLLHandle)
50070   If res <> 0 Then
50080     UnLoadDLL = True
50090    Else
50100     UnLoadDLL = False
50110   End If
50120  End If
50130 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50140 Exit Function
ErrPtnr_OnError:
50161 Select Case ErrPtnr.OnError("modGeneral", "UnLoadDLL")
      Case 0: Resume
50180 Case 1: Resume Next
50190 Case 2: Exit Function
50200 Case 3: End
50210 End Select
50220 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function UnloadDLLComplete(ByRef DLLHandle As Long) As Long
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim c As Long
50050  If DLLHandle = 0 Then
50060   Exit Function
50070  End If
50080  c = 0
50090  Do While UnLoadDLL(DLLHandle)
50100   c = c + 1
50110   DoEvents
50120  Loop
50130  UnloadDLLComplete = c: DLLHandle = 0
50140 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50150 Exit Function
ErrPtnr_OnError:
50171 Select Case ErrPtnr.OnError("modGeneral", "UnloadDLLComplete")
      Case 0: Resume
50190 Case 1: Resume Next
50200 Case 2: Exit Function
50210 Case 3: End
50220 End Select
50230 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetStrFromPtrA(lpszA As Long) As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  GetStrFromPtrA = String$(lstrlen(ByVal lpszA), 0)
50050  Call lstrcpy(ByVal GetStrFromPtrA, ByVal lpszA)
50060 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50070 Exit Function
ErrPtnr_OnError:
50091 Select Case ErrPtnr.OnError("modGeneral", "GetStrFromPtrA")
      Case 0: Resume
50110 Case 1: Resume Next
50120 Case 2: Exit Function
50130 Case 3: End
50140 End Select
50150 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function RaiseAPIError() As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim ErrorMsg As String, ErrNum As Long
50050  ErrNum = Err.LastDllError
50060  ErrorMsg = String(256, 0)
50070  ErrorMsg = Left$(ErrorMsg, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, ErrNum, 0&, ErrorMsg, Len(ErrorMsg), ByVal 0))
50080  If Mid(ErrorMsg, Len(ErrorMsg) - 1) = vbCrLf Then
50090   ErrorMsg = Mid(ErrorMsg, 1, Len(ErrorMsg) - 2)
50100  End If
50110  RaiseAPIError = ErrNum & ": " & ErrorMsg
50120 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50130 Exit Function
ErrPtnr_OnError:
50151 Select Case ErrPtnr.OnError("modGeneral", "RaiseAPIError")
      Case 0: Resume
50170 Case 1: Resume Next
50180 Case 2: Exit Function
50190 Case 3: End
50200 End Select
50210 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetAllFileExtensions() As Collection
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim reg As clsRegistry, tColl As Collection, i As Long
50050  Set reg = New clsRegistry
50060  Set tColl = reg.EnumRegistryKeys(HKEY_CLASSES_ROOT, "")
50070  Set GetAllFileExtensions = New Collection
50080  For i = 1 To tColl.Count
50090   If Len(tColl(i)) > 0 Then
50100    If Mid(tColl(i), 1, 1) = "." Then
50110     GetAllFileExtensions.Add Mid(tColl(i), 2)
50120    End If
50130   End If
50140  Next i
50150  Set reg = Nothing
50160 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50170 Exit Function
ErrPtnr_OnError:
50191 Select Case ErrPtnr.OnError("modGeneral", "GetAllFileExtensions")
      Case 0: Resume
50210 Case 1: Resume Next
50220 Case 2: Exit Function
50230 Case 3: End
50240 End Select
50250 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function StringInCollection(coll As Collection, Str1 As String, Optional CaseSensitive As Boolean = False) As Boolean
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long, tstr As String
50050  StringInCollection = False: tstr = UCase$(Str1)
50060  If Not coll Is Nothing And Len(Str1) > 0 Then
50070   For i = 1 To coll.Count
50080    If CaseSensitive = True Then
50090      If coll(i) = Str1 Then
50100       StringInCollection = True
50110       Exit Function
50120      End If
50130     Else
50140      If UCase$(coll(i)) = UCase$(Str1) Then
50150       StringInCollection = True
50160       Exit Function
50170      End If
50180    End If
50190   Next i
50200  End If
50210 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50220 Exit Function
ErrPtnr_OnError:
50241 Select Case ErrPtnr.OnError("modGeneral", "StringInCollection")
      Case 0: Resume
50260 Case 1: Resume Next
50270 Case 2: Exit Function
50280 Case 3: End
50290 End Select
50300 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function RemoveAllKnownFileExtensions(Filename As String) As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim tColl As Collection, Ext As String, File As String, OldFilename As String
50050  RemoveAllKnownFileExtensions = Filename
50060  SplitPath Filename, , , , File, Ext
50070  If LenB(Ext) > 0 Then
50080   Set tColl = GetAllFileExtensions
50090   Do
50100    SplitPath Filename, , , , File, Ext
50110    Filename = File
50120    DoEvents
50130   Loop Until StringInCollection(tColl, Ext) = False
50140   If LenB(Ext) > 0 Then
50150     RemoveAllKnownFileExtensions = File & "." & Ext
50160    Else
50170     RemoveAllKnownFileExtensions = File
50180   End If
50190  End If
50200 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50210 Exit Function
ErrPtnr_OnError:
50231 Select Case ErrPtnr.OnError("modGeneral", "RemoveAllKnownFileExtensions")
      Case 0: Resume
50250 Case 1: Resume Next
50260 Case 2: Exit Function
50270 Case 3: End
50280 End Select
50290 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetIExplorerVersion() As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim reg As clsRegistry
50050  Set reg = New clsRegistry
50060  With reg
50070   .hkey = HKEY_LOCAL_MACHINE
50080   .KeyRoot = "SOFTWARE\Microsoft\Internet Explorer"
50090   GetIExplorerVersion = .GetRegistryValue("Version")
50100  End With
50110  Set reg = Nothing
50120 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50130 Exit Function
ErrPtnr_OnError:
50151 Select Case ErrPtnr.OnError("modGeneral", "GetIExplorerVersion")
      Case 0: Resume
50170 Case 1: Resume Next
50180 Case 2: Exit Function
50190 Case 3: End
50200 End Select
50210 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function FormISLoaded(FormName As String) As Boolean
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long
50050  FormISLoaded = False
50060  For i = 1 To Forms.Count
50070   If UCase$(FormName) = UCase$(Forms(i - 1).Name) Then
50080    FormISLoaded = True
50090   End If
50100  Next i
50110 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50120 Exit Function
ErrPtnr_OnError:
50141 Select Case ErrPtnr.OnError("modGeneral", "FormISLoaded")
      Case 0: Resume
50160 Case 1: Resume Next
50170 Case 2: Exit Function
50180 Case 3: End
50190 End Select
50200 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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

Public Sub WriteInfoSpoolfile(Spoolfilename As String)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim pbg As PropertyBag, fn As Long, Path As String, File As String, _
  infFile As String
50060  Set pbg = New PropertyBag
50070  With pbg
50080   .WriteProperty "REDMON_PORT", Environ$("REDMON_PORT")
50090   .WriteProperty "REDMON_JOB", Environ$("REDMON_JOB")
50100   .WriteProperty "REDMON_PRINTER", Environ$("REDMON_PRINTER")
50110   .WriteProperty "REDMON_MACHINE", Environ$("REDMON_MACHINE")
50120   .WriteProperty "REDMON_USER", Environ$("REDMON_USER")
50130   .WriteProperty "REDMON_DOCNAME", Environ$("REDMON_DOCNAME")
50140   .WriteProperty "REDMON_FILENAME", Environ$("REDMON_FILENAME")
50150   .WriteProperty "REDMON_SESSIONID", Environ$("REDMON_SESSIONID")
50160   .WriteProperty "SpoolFilename", Spoolfilename
50170   .WriteProperty "LoggedOnUser", GetUsername
50180   .WriteProperty "Computer", GetComputerName
50190   .WriteProperty "Created", Now
50200  End With
50210  fn = FreeFile
50220  SplitPath Spoolfilename, , Path, , File
50230  infFile = CompletePath(Path) & File & ".inf"
50240  Open infFile For Binary As #fn
50250  Put #fn, , pbg.Contents
50260  Close #fn
50270 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50280 Exit Sub
ErrPtnr_OnError:
50301 Select Case ErrPtnr.OnError("modGeneral", "WriteInfoSpoolfile")
      Case 0: Resume
50320 Case 1: Resume Next
50330 Case 2: Exit Sub
50340 Case 3: End
50350 End Select
50360 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
Public Function InitPropertyBag(Optional Spoolfilename As String = "") As PropertyBag
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Set InitPropertyBag = New PropertyBag
50050  With InitPropertyBag
50060   .WriteProperty "REDMON_PORT", Environ$("REDMON_PORT")
50070   .WriteProperty "REDMON_JOB", Environ$("REDMON_JOB")
50080   .WriteProperty "REDMON_PRINTER", Environ$("REDMON_PRINTER")
50090   .WriteProperty "REDMON_MACHINE", Environ$("REDMON_MACHINE")
50100   .WriteProperty "REDMON_USER", Environ$("REDMON_USER")
50110   .WriteProperty "REDMON_DOCNAME", Environ$("REDMON_DOCNAME")
50120   .WriteProperty "REDMON_FILENAME", Environ$("REDMON_FILENAME")
50130   .WriteProperty "REDMON_SESSIONID", Environ$("REDMON_SESSIONID")
50140   .WriteProperty "SpoolFilename", Spoolfilename
50150   .WriteProperty "LoggedOnUser", GetUsername
50160   .WriteProperty "Computer", GetComputerName
50170   .WriteProperty "Created", Now
50180  End With
50190 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50200 Exit Function
ErrPtnr_OnError:
50221 Select Case ErrPtnr.OnError("modGeneral", "InitPropertyBag")
      Case 0: Resume
50240 Case 1: Resume Next
50250 Case 2: Exit Function
50260 Case 3: End
50270 End Select
50280 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function ReadInfoSpoolfile(Spoolfilename As String) As PropertyBag
 On Error Resume Next
 Dim tVar As Variant, fn As Long, infFile As String, _
  Path As String, File As String
 Set ReadInfoSpoolfile = InitPropertyBag(Spoolfilename)
 SplitPath Spoolfilename, , Path, , File
 infFile = CompletePath(Path) & File & ".inf"
 If FileExists(infFile) And Not FileInUse(infFile) Then
  fn = FreeFile
  Open infFile For Binary As #fn
  Get #fn, , tVar
  Close #fn
  ReadInfoSpoolfile.Contents = tVar
 End If
End Function

Public Function IsInIDE() As Boolean
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  IsInIDE = App.LogMode <> 1
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Function
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("modGeneral", "IsInIDE")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Function
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function


Public Sub KillInfoSpoolfile(Spoolfilename As String)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim infFile As String, Path As String, File As String
50050  SplitPath Spoolfilename, , Path, , File
50060  infFile = CompletePath(Path) & File & ".inf"
50070  KillFile infFile
50080 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50090 Exit Sub
ErrPtnr_OnError:
50111 Select Case ErrPtnr.OnError("modGeneral", "KillInfoSpoolfile")
      Case 0: Resume
50130 Case 1: Resume Next
50140 Case 2: Exit Sub
50150 Case 3: End
50160 End Select
50170 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function ScreenResolution() As Integer
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  ScreenResolution = GetDeviceCaps(GetDC(0), BITSPIXEL)
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Function
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("modGeneral", "ScreenResolution")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Function
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
