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

Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias _
 "GetPrivateProfileSectionA" (ByVal Section As String, ByVal Buffer As String, ByVal Size As Long, ByVal Filename As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias _
 "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias _
 "GetPrivateProfileStringA" (ByVal Section As String, ByVal Key As String, ByVal Default As String, ByVal Buffer As String, ByVal Size As Long, ByVal Filename As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias _
 "WritePrivateProfileSectionA" (ByVal Section As String, ByVal Key As String, ByVal Filename As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias _
 "WritePrivateProfileStringA" (ByVal Section As String, ByVal Key As String, ByVal Setting As String, ByVal Filename As String) As Long

Private iFilename As String, iSection As String, iKey As String, _
 iErrCode As Long, iErrVar(1 To 2) As String

Public Property Let Key(ByVal vData As String)
 iKey = vData
End Property

Public Property Get Key() As String
 Key = iKey
End Property

Public Property Let Section(ByVal vData As String)
 iSection = vData
End Property

Public Property Get Section() As String
 Section = iSection
End Property

Public Property Let Filename(ByVal vData As String)
 iFilename = Trim$(vData)
End Property

Public Property Get Filename() As String
Attribute Filename.VB_UserMemId = 0
 Filename = iFilename
End Property

Public Function CheckIniFile(Optional Filename As String) As Boolean
 Dim tFilename As String
 CheckIniFile = False
 tFilename = GetFilename(Filename)
 If Len(tFilename) = 0 Then
  SetError -1
  Exit Function
 End If
 If Dir(tFilename) = "" Then
  SetError -2: SetErrorVariable 1, tFilename
  Exit Function
 End If
 CheckIniFile = True
End Function

Public Function GetAllSectionsFromInifile(Optional ByVal Filename As String, Optional Sorted As Boolean = False) As Collection
 On Local Error GoTo ErrorHandler
 Dim tFilename As String, i As Long, j As Long, nBuffer As String, _
  res As Long, tStr As String, nItems() As String

 Const kBufferSize = 32767

 tFilename = GetFilename(Filename)

 If CheckIniFile(tFilename) = True Then
  Set GetAllSectionsFromInifile = New Collection
  nBuffer = Space$(kBufferSize)
  res = GetPrivateProfileSectionNames(nBuffer, kBufferSize, tFilename & vbNullString)
  If res = 0 Then
   SetError -3: SetErrorVariable 1, tFilename
   Exit Function
  End If
  nItems = Split(Left$(nBuffer, res), Chr$(0))
  If Sorted = True Then
   For i = LBound(nItems) To UBound(nItems) - 1
    For j = i To UBound(nItems)
     If StrComp(nItems(i), nItems(j), vbTextCompare) = 1 Then
      tStr = nItems(i): nItems(i) = nItems(j): nItems(j) = tStr
     End If
    Next j
   Next i
  End If
  For i = 0 To UBound(nItems)
   If Trim$(nItems(i)) <> "" Then
    GetAllSectionsFromInifile.Add nItems(i)
   End If
  Next i
 End If
 Exit Function

ErrorHandler:
 SetError -99: SetErrorVariable 1, "GetAllSectionsFromInifile": SetErrorVariable 2, _
  "[" & Err.Number & ": " & Err.Description & "]"
 Err.Clear
End Function

Public Function CheckSection(Optional Section As String, Optional Filename As String) As Boolean
 On Local Error GoTo ErrorHandler
 Dim tFilename As String, tSection As String, i As Long, tcoll As Collection
 CheckSection = False
 tFilename = GetFilename(Filename)
 tSection = GetSectionname(Section)
 Set tcoll = GetAllSectionsFromInifile(tFilename)
 For i = 1 To tcoll.Count
  If UCase$(tSection) = UCase$(tcoll.Item(i)) Then
   CheckSection = True
   Set tcoll = Nothing
   Exit Function
  End If
 Next i
 Set tcoll = Nothing
 Exit Function

ErrorHandler:
 SetError -99: SetErrorVariable 1, "checkSection": SetErrorVariable 2, _
  "[" & Err.Number & ": " & Err.Description & "]"
 Err.Clear
End Function

Public Function GetAllKeysFromSection(Optional Section As String, Optional Filename As String, Optional WithComments As Boolean = False, Optional Sorted As Boolean = False) As Collection
 On Local Error GoTo ErrorHandler
 Const kBufferSize = 32767
 Dim tFilename As String, tSection As String, nBuffer As String, res As Long, _
  nItems() As String, nItem() As String, i As Long, tcoll As Collection, _
  ts(1) As String, j As Long, tStr As String

 tFilename = GetFilename(Filename)
 tSection = GetSectionname(Section)

 Set tcoll = New Collection
 Set GetAllKeysFromSection = tcoll
 If CheckSection(tSection, tFilename) = True Then
  nBuffer = Space$(kBufferSize)
  res = GetPrivateProfileSection(tSection, nBuffer, kBufferSize, tFilename)
  If res = 0 Then
   SetError -4: SetErrorVariable 1, tFilename
   Exit Function
  End If
  nItems = Split(Left$(nBuffer, res), Chr$(0))
  If Sorted = True Then
   For i = LBound(nItems) To UBound(nItems) - 1
    For j = i To UBound(nItems)
     If StrComp(nItems(i), nItems(j), vbTextCompare) = 1 Then
      tStr = nItems(i): nItems(i) = nItems(j): nItems(j) = tStr
     End If
    Next j
   Next i
  End If
  For i = LBound(nItems) To UBound(nItems)
   If Trim$(nItems(i)) <> "" Then
    If WithComments = True Then
      If InStr(nItems(i), "=") > 0 Then
        nItem = Split(nItems(i), "=")
        ts(0) = nItem(0): ts(1) = nItem(1)
        tcoll.Add ts
       Else
        ts(0) = nItems(i): ts(1) = ""
        tcoll.Add ts
      End If
     Else
      If InStr(Trim$(nItems(i)), ";") <> 1 Then
       If InStr(nItems(i), "=") > 0 Then
         nItem = Split(nItems(i), "=")
         ts(0) = nItem(0): ts(1) = nItem(1)
         tcoll.Add ts
        Else
         ts(0) = nItems(i): ts(1) = ""
         tcoll.Add ts
       End If
      End If
    End If
     
   End If
  Next i
  Set GetAllKeysFromSection = tcoll
 End If
 Exit Function

ErrorHandler:
 SetError -99: SetErrorVariable 1, "GetAllKeysFromSection": SetErrorVariable 2, _
  "[" & Err.Number & ": " & Err.Description & "]"
 Err.Clear
End Function

Public Function CheckKey(Optional Key As String, Optional Section As String, Optional Filename As String) As Boolean
 On Local Error GoTo ErrorHandler
 Dim tFilename As String, tSection As String, tKey As String, _
  i As Long, tcoll As Collection
 CheckKey = False
 tFilename = GetFilename(Filename)
 tSection = GetSectionname(Section)
 tKey = GetKeyname(Key)
 
 Set tcoll = GetAllKeysFromSection(tSection, tFilename)
 For i = 1 To tcoll.Count
  If UCase$(tKey) = UCase$(tcoll.Item(i)(0)) Then
   CheckKey = True
   Set tcoll = Nothing
   Exit Function
  End If
 Next i
 Exit Function

ErrorHandler:
 SetError -99: SetErrorVariable 1, "checkKey": SetErrorVariable 2, _
  "[" & Err.Number & ": " & Err.Description & "]"
 Err.Clear
End Function

Public Function GetKeyFromSection(Optional Key As String, Optional Section As String, Optional Filename As String) As String
 On Local Error GoTo ErrorHandler
 Const kBufferSize = 1024
 Dim nBuffer As String, tKey As String, tSection As String, tFilename As String, _
  tcoll As Collection, i As Long, res As Long
 tFilename = GetFilename(Filename)
 tSection = GetSectionname(Section)
 tKey = GetKeyname(Key)
 
 GetKeyFromSection = ""
 If CheckSection(tSection, tFilename) = True Then
   Set tcoll = GetAllKeysFromSection(tSection, tFilename)
   For i = 1 To tcoll.Count
    If UCase$(tKey) = UCase$(tcoll.Item(i)(0)) Then
     GetKeyFromSection = tcoll.Item(i)(1)
     Set tcoll = Nothing
     Exit Function
    End If
   Next i
  Else
   SetError -5: SetErrorVariable 1, tKey: SetErrorVariable 2, tSection
 End If
 Exit Function

ErrorHandler:
 SetError -99: SetErrorVariable 1, "GetKeyFromSection": SetErrorVariable 2, _
  "[" & Err.Number & ": " & Err.Description & "]"
 Err.Clear
End Function

Public Sub DeleteIniFile(Optional Filename As String)
 On Local Error GoTo ErrorHandler
 Dim tFilename As String
 tFilename = GetFilename(Filename)
 If Dir(tFilename) <> "" Then
  Kill tFilename
 End If
 Exit Sub

ErrorHandler:
 SetError -99: SetErrorVariable 1, "DeleteIniFile": SetErrorVariable 2, _
  "[" & Err.Number & ": " & Err.Description & "]"
 Err.Clear
End Sub

Public Sub CreateIniFile(Optional Filename As String)
 On Local Error GoTo ErrorHandler
 Dim tFilename As String, fn As Long
 tFilename = GetFilename(Filename)
 
 If tFilename = "" Then
  SetError -9
 End If
 If Dir(tFilename) <> "" Then
  SetError -10: SetErrorVariable 1, tFilename
 End If
 
 fn = FreeFile
 Open tFilename For Output As #fn
 Close #fn
 FlushInifile
 Exit Sub

ErrorHandler:
 SetError -99: SetErrorVariable 1, "CreateIniFile": SetErrorVariable 2, _
  "[" & Err.Number & ": " & Err.Description & "]"
 Err.Clear
End Sub

Public Sub CreateSection(Optional Section As String, Optional Filename As String)
 On Local Error GoTo ErrorHandler
 Dim tFilename As String, tSection As String, res As Long, fn As Long
 
 tFilename = GetFilename(Filename)
 tSection = GetSectionname(Section)
 
 If CheckSection(tSection, tFilename) = False Then
   fn = FreeFile
   Open tFilename For Append As #fn
   Print #fn, "[" & tSection & "]"
   Close #fn
  Else
   SetError -11: SetErrorVariable 1, tSection: SetErrorVariable 2, tFilename
 End If
 FlushInifile
 Exit Sub

ErrorHandler:
 SetError -99: SetErrorVariable 1, "CreateIniFile": SetErrorVariable 2, _
  "[" & Err.Number & ": " & Err.Description & "]"
 Err.Clear
End Sub


Public Sub DeleteAllSectionFromInifile(Optional Filename As String)
 On Local Error GoTo ErrorHandler
 Dim tKey As String, tSection As String, tFilename As String, res As Long, _
  i As Long, tcoll As Collection

 tFilename = GetFilename(Filename)
 
 If CheckIniFile(tFilename) = True Then
   res = WritePrivateProfileString(vbNullString, vbNullString, vbNullString, tFilename)
   If res = 0 Then
    SetError -99: SetErrorVariable 1, "DeleteAllSectionFromInifile": SetErrorVariable 2, "Unknown Error"
   End If
  Else
   SetError -2: SetErrorVariable 1, tFilename
 End If
 Exit Sub

ErrorHandler:
 SetError -99: SetErrorVariable 1, "DeleteAllSectionFromInifile": SetErrorVariable 2, _
  "[" & Err.Number & ": " & Err.Description & "]"
 Err.Clear
End Sub

Public Sub DeleteSectionFromInifile(Optional Section As String, Optional Filename As String)
 On Local Error GoTo ErrorHandler
 Dim tKey As String, tSection As String, tFilename As String, res As Long, _
  i As Long, tcoll As Collection

 tFilename = GetFilename(Filename)
 tSection = GetSectionname(Section)
 
 If CheckSection(tSection, tFilename) = True Then
   res = WritePrivateProfileString(tSection, vbNullString, vbNullString, tFilename)
   If res = 0 Then
    SetError -99: SetErrorVariable 1, "DeleteSectionFromInifile": SetErrorVariable 2, "Unknown Error"
   End If
  Else
   SetError -6: SetErrorVariable 1, tKey: SetErrorVariable 2, tSection
 End If
 Exit Sub

ErrorHandler:
 SetError -99: SetErrorVariable 1, "DeleteSectionFromInifile": SetErrorVariable 2, _
  "[" & Err.Number & ": " & Err.Description & "]"
 Err.Clear
End Sub

Public Sub DeleteKeyFromSection(Optional Key As String, Optional Section As String, Optional Filename As String)
 On Local Error GoTo ErrorHandler
 Dim tKey As String, tSection As String, tFilename As String, res As Long

 tFilename = GetFilename(Filename)
 tSection = GetSectionname(Section)
 tKey = GetKeyname(Key)
 If tKey = "" Then
  SetError -8: SetErrorVariable 1, tSection
 End If
 
 If CheckKey(tKey, tSection, tFilename) = True Then
   res = WritePrivateProfileString(tSection, tKey, vbNullString, tFilename)
   If res = 0 Then
    SetError -99: SetErrorVariable 1, "DeleteKeyFromSection": SetErrorVariable 2, "Unknown Error"
   End If
  Else
   SetError -7: SetErrorVariable 1, tKey: SetErrorVariable 2, tSection
 End If

 Exit Sub

ErrorHandler:
 SetError -99: SetErrorVariable 1, "DeleteKeyFromSection": SetErrorVariable 2, _
  "[" & Err.Number & ": " & Err.Description & "]"
 Err.Clear
End Sub

Public Function SaveKey(Value As String, Optional Key As String, Optional Section As String, Optional Filename As String) As Boolean
 Dim tKey As String, tSection As String, tFilename As String, res As Long

 tFilename = GetFilename(Filename)
 tSection = GetSectionname(Section)
 tKey = GetKeyname(Key)

 res = WritePrivateProfileString(tSection & vbNullString, tKey & vbNullString, _
  Value & vbNullString, tFilename & vbNullString)
 FlushInifile
 If res = 0 Then
  SetError -12: SetErrorVariable 1, tKey: SetErrorVariable 2, tSection
 End If
End Function

Public Sub FlushInifile(Optional Filename As String)
 Dim res As Long, tFilename As String
 tFilename = GetFilename(Filename)
 res = WritePrivateProfileString(vbNullString, vbNullString, vbNullString, tFilename & vbNullString)
End Sub


Private Sub SetError(ErrCode As Long)
 iErrCode = ErrCode
End Sub

Private Sub SetErrorVariable(Index As Long, ErrVar As String)
 iErrVar(Index) = Trim$(ErrVar)
End Sub

Public Property Get ErrCode() As Long
 ErrCode = iErrCode
End Property

Public Property Get ErrDescription() As String
 Const Prefix = "INI-Class: "
 Select Case iErrCode
  Case 0:
   ErrDescription = ""
  Case -1:
   ErrDescription = Prefix & "Empty Filename."
  Case -2:
   ErrDescription = Prefix & "Filename <" & iErrVar(1) & "> not found."
  Case -3:
   ErrDescription = Prefix & "GetSections: No Sections found in <" & iErrVar(1) & ">."
  Case -4:
   ErrDescription = Prefix & "GetKeys: No Settings found in <" & iErrVar(1) & ">."
  Case -5:
   ErrDescription = Prefix & "GetKey: Key <" & iErrVar(1) & "> not found in Section <" & iErrVar(2) & ">."
  Case -6:
   ErrDescription = Prefix & "DeleteSection: Section <" & iErrVar(1) & "> not found in Inifile <" & iErrVar(2) & ">."
  Case -7:
   ErrDescription = Prefix & "DeleteKey: Key <" & iErrVar(1) & "> not found in Section <" & iErrVar(2) & ">."
  Case -8:
   ErrDescription = Prefix & "DeleteKey: Cannot delete empty key in Section <" & iErrVar(1) & ">."
  Case -9:
   ErrDescription = Prefix & "CreateInifile: Cannot create Inifile: Empty filename."
  Case -10:
   ErrDescription = Prefix & "CreateInifile: Filename <" & iErrVar(1) & "> already exists."
  Case -11:
   ErrDescription = Prefix & "CreateSection: Section <" & iErrVar(1) & "> not found in Inifile <" & iErrVar(2) & ">."
  Case -12:
   ErrDescription = Prefix & "CreateKey: Cannot create Key <" & iErrVar(1) & "> in Section <" & iErrVar(2) & ">."
  Case Else
   ErrDescription = Prefix & "Unknown Error in <" & iErrVar(1) & "> Error-Code [" & iErrCode & "]  Error-Description [" & iErrVar(2) & "]."
 End Select
End Property

Public Sub ErrClear()
 iErrCode = 0: iErrVar(1) = "": iErrVar(2) = ""
End Sub

Private Function GetFilename(ByVal Filename As String) As String
 Filename = Trim$(Filename)
 If Filename = "" Then
   GetFilename = iFilename
  Else
   GetFilename = Filename
 End If
End Function

Private Function GetSectionname(ByVal Sectionname As String) As String
 Sectionname = Trim$(Sectionname)
 If Sectionname = "" Then
   GetSectionname = iSection
  Else
   GetSectionname = Sectionname
 End If
End Function

Private Function GetKeyname(ByVal Keyname As String) As String
 Keyname = Trim$(Keyname)
 If Keyname = "" Then
   GetKeyname = iKey
  Else
   GetKeyname = Keyname
 End If
End Function

Private Sub Class_Initialize()
 ErrClear
End Sub
