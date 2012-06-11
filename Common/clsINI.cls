VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private iFilename As String, iSection As String, iKey As String, _
 iErrCode As Long, iErrVar(1 To 2) As String

Public Property Let key(ByVal vData As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  iKey = vData
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsINI", "key [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get key() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  key = iKey
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsINI", "key [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let Section(ByVal vData As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  iSection = vData
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsINI", "Section [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get Section() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Section = iSection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsINI", "Section [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let filename(ByVal vData As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  iFilename = Trim$(vData)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsINI", "filename [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get filename() As String
Attribute filename.VB_UserMemId = 0
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  filename = iFilename
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsINI", "filename [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Function CheckIniFile(Optional filename As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tFilename As String
50020  CheckIniFile = False
50030  tFilename = GetFilename(filename)
50040  If Len(tFilename) = 0 Then
50050   Exit Function
50060  End If
50070  If Not FileExists(tFilename) Then
50080   Exit Function
50090  End If
50100  CheckIniFile = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsINI", "CheckIniFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetAllSectionsFromInifile(Optional ByVal filename As String, Optional Sorted As Boolean = False) As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tFilename As String, i As Long, j As Long, nBuffer As String, _
  res As Long, tStr As String, nItems() As String
50030
50040  Const kBufferSize = 32767
50050
50060  Set GetAllSectionsFromInifile = New Collection
50070
50080  tFilename = GetFilename(filename)
50090
50100  If CheckIniFile(tFilename) = True Then
50110   nBuffer = Space$(kBufferSize)
50120   res = GetPrivateProfileSectionNames(nBuffer, kBufferSize, tFilename & vbNullString)
50130   If res = 0 Then
50140    Exit Function
50150   End If
50160   nItems = Split(Left$(nBuffer, res), Chr$(0))
50170   If Sorted = True Then
50180    For i = LBound(nItems) To UBound(nItems) - 1
50190     For j = i To UBound(nItems)
50200      If StrComp(nItems(i), nItems(j), vbTextCompare) = 1 Then
50210       tStr = nItems(i): nItems(i) = nItems(j): nItems(j) = tStr
50220      End If
50230     Next j
50240    Next i
50250   End If
50260   For i = 0 To UBound(nItems)
50270    If Trim$(nItems(i)) <> vbNullString Then
50280     GetAllSectionsFromInifile.Add nItems(i)
50290    End If
50300   Next i
50310  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsINI", "GetAllSectionsFromInifile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function CheckSection(Optional Section As String, Optional filename As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tFilename As String, tSection As String, i As Long, tColl As Collection
50020  CheckSection = False
50030  tFilename = GetFilename(filename)
50040  tSection = GetSectionname(Section)
50050  Set tColl = GetAllSectionsFromInifile(tFilename)
50060  For i = 1 To tColl.Count
50070   If UCase$(tSection) = UCase$(tColl.Item(i)) Then
50080    CheckSection = True
50090    Set tColl = Nothing
50100    Exit Function
50110   End If
50120  Next i
50130  Set tColl = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsINI", "CheckSection")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetAllKeysFromSection(Optional Section As String, Optional filename As String, Optional WithComments As Boolean = False, Optional Sorted As Boolean = False) As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Const kBufferSize = 32767
50020  Dim tFilename As String, tSection As String, nBuffer As String, res As Long, _
  nItems() As String, nItem() As String, i As Long, tColl As Collection, _
  ts(1) As String, j As Long, tStr As String
50050
50060  tFilename = GetFilename(filename)
50070  tSection = GetSectionname(Section)
50080
50090  Set tColl = New Collection
50100  Set GetAllKeysFromSection = tColl
50110  If CheckSection(tSection, tFilename) = True Then
50120   nBuffer = Space$(kBufferSize)
50130   res = GetPrivateProfileSection(tSection, nBuffer, kBufferSize, tFilename)
50140   If res = 0 Then
50150    Exit Function
50160   End If
50170   nItems = Split(Left$(nBuffer, res), Chr$(0))
50180   If Sorted = True Then
50190    For i = LBound(nItems) To UBound(nItems) - 1
50200     For j = i To UBound(nItems)
50210      If StrComp(nItems(i), nItems(j), vbTextCompare) = 1 Then
50220       tStr = nItems(i): nItems(i) = nItems(j): nItems(j) = tStr
50230      End If
50240     Next j
50250    Next i
50260   End If
50270   For i = LBound(nItems) To UBound(nItems)
50280    If Len(Trim$(nItems(i))) > 0 Then
50290     If WithComments = True Then
50300       If InStr(nItems(i), "=") > 0 Then
50310         nItem = Split(nItems(i), "=")
50320         ts(0) = nItem(0)
50330         ts(1) = nItem(1)
50340         For j = 2 To UBound(nItem)
50350          ts(1) = ts(1) & "=" & nItem(j)
50360         Next j
50370         tColl.Add ts
50380        Else
50390         ts(0) = nItems(i): ts(1) = vbNullString
50400         tColl.Add ts
50410       End If
50420      Else
50430       If InStr(Trim$(nItems(i)), ";") <> 1 Then
50440        If InStr(nItems(i), "=") > 0 Then
50450          nItem = Split(nItems(i), "=")
50460          ts(0) = nItem(0)
50470          ts(1) = nItem(1)
50480          For j = 2 To UBound(nItem)
50490           ts(1) = ts(1) & "=" & nItem(j)
50500          Next j
50510          tColl.Add ts
50520         Else
50530          ts(0) = nItems(i): ts(1) = vbNullString
50540          tColl.Add ts
50550        End If
50560       End If
50570     End If
50580    End If
50590   Next i
50600   Set GetAllKeysFromSection = tColl
50610  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsINI", "GetAllKeysFromSection")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function CheckKey(Optional key As String, Optional Section As String, Optional filename As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tFilename As String, tSection As String, tKey As String, _
  i As Long, tColl As Collection
50030  CheckKey = False
50040  tFilename = GetFilename(filename)
50050  tSection = GetSectionname(Section)
50060  tKey = GetKeyname(key)
50070
50080  Set tColl = GetAllKeysFromSection(tSection, tFilename)
50090  For i = 1 To tColl.Count
50100   If UCase$(tKey) = UCase$(tColl.Item(i)(0)) Then
50110    CheckKey = True
50120    Set tColl = Nothing
50130    Exit Function
50140   End If
50150  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsINI", "CheckKey")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetKeyFromSection(Optional key As String, Optional Section As String, Optional filename As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Const kBufferSize = 1024
50020  Dim nBuffer As String, tKey As String, tSection As String, tFilename As String, _
  tColl As Collection, i As Long, res As Long
50040  tFilename = GetFilename(filename)
50050  tSection = GetSectionname(Section)
50060  tKey = GetKeyname(key)
50070
50080  GetKeyFromSection = vbNullString
50090  If CheckSection(tSection, tFilename) = True Then
50100   Set tColl = GetAllKeysFromSection(tSection, tFilename)
50110   For i = 1 To tColl.Count
50120    If UCase$(tKey) = UCase$(tColl.Item(i)(0)) Then
50130     GetKeyFromSection = tColl.Item(i)(1)
50140     Set tColl = Nothing
50150     Exit Function
50160    End If
50170   Next i
50180  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsINI", "GetKeyFromSection")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub DeleteIniFile(Optional filename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tFilename As String
50020  tFilename = GetFilename(filename)
50030  If Dir(tFilename) <> vbNullString Then
50040   Kill tFilename
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsINI", "DeleteIniFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub CreateIniFile(Optional filename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tFilename As String, fn As Long
50020  tFilename = GetFilename(filename)
50030
50040  fn = FreeFile
50050  Open tFilename For Output As #fn
50060  Close #fn
50070  FlushInifile
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsINI", "CreateIniFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub CreateSection(Optional Section As String, Optional filename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tFilename As String, tSection As String, res As Long, fn As Long
50020
50030  tFilename = GetFilename(filename)
50040  tSection = GetSectionname(Section)
50050
50060  If CheckSection(tSection, tFilename) = False Then
50070   fn = FreeFile
50080   Open tFilename For Append As #fn
50090   Print #fn, "[" & tSection & "]"
50100   Close #fn
50110  End If
50120  FlushInifile
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsINI", "CreateSection")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub DeleteAllSectionFromInifile(Optional filename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tKey As String, tSection As String, tFilename As String, res As Long, _
  i As Long, tColl As Collection
50030
50040  tFilename = GetFilename(filename)
50050
50060  If CheckIniFile(tFilename) = True Then
50070   res = WritePrivateProfileString(vbNullString, vbNullString, vbNullString, tFilename)
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsINI", "DeleteAllSectionFromInifile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub DeleteSectionFromInifile(Optional Section As String, Optional filename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tKey As String, tSection As String, tFilename As String, res As Long, _
  i As Long, tColl As Collection
50030
50040  tFilename = GetFilename(filename)
50050  tSection = GetSectionname(Section)
50060
50070  If CheckSection(tSection, tFilename) = True Then
50080   res = WritePrivateProfileString(tSection, vbNullString, vbNullString, tFilename)
50090  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsINI", "DeleteSectionFromInifile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub DeleteKeyFromSection(Optional key As String, Optional Section As String, Optional filename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tKey As String, tSection As String, tFilename As String, res As Long
50020
50030  tFilename = GetFilename(filename)
50040  tSection = GetSectionname(Section)
50050  tKey = GetKeyname(key)
50060
50070  If CheckKey(tKey, tSection, tFilename) = True Then
50080   res = WritePrivateProfileString(tSection, tKey, vbNullString, tFilename)
50090  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsINI", "DeleteKeyFromSection")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function SaveKey(value As String, Optional key As String, Optional Section As String, Optional filename As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tKey As String, tSection As String, tFilename As String, res As Long
50020
50030  tFilename = GetFilename(filename)
50040  tSection = GetSectionname(Section)
50050  tKey = GetKeyname(key)
50060
50070  res = WritePrivateProfileString(tSection & vbNullString, tKey & vbNullString, _
  value & vbNullString, tFilename & vbNullString)
50090  If res <> 0 Then
50100   SaveKey = True
50110  End If
50120  FlushInifile
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsINI", "SaveKey")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub FlushInifile(Optional filename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long, tFilename As String
50020  tFilename = GetFilename(filename)
50030  res = WritePrivateProfileString(vbNullString, vbNullString, vbNullString, tFilename & vbNullString)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsINI", "FlushInifile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function GetFilename(ByVal filename As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  filename = Trim$(filename)
50020  If Len(filename) = 0 Then
50030    GetFilename = iFilename
50040   Else
50050    GetFilename = filename
50060  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsINI", "GetFilename")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function GetSectionname(ByVal Sectionname As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Sectionname = Trim$(Sectionname)
50020  If Len(Sectionname) = 0 Then
50030    GetSectionname = iSection
50040   Else
50050    GetSectionname = Sectionname
50060  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsINI", "GetSectionname")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function GetKeyname(ByVal Keyname As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Keyname = Trim$(Keyname)
50020  If Len(Keyname) = 0 Then
50030    GetKeyname = iKey
50040   Else
50050    GetKeyname = Keyname
50060  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsINI", "GetKeyname")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
