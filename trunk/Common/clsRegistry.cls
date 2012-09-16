VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type SECURITY_ATTRIBUTES
 nLength As Long
 lpSecurityDescriptor As Long
 bInheritHandle As Long
End Type

Private Type FILETIME
 dwLowDateTime As Long
 dwHighDateTime As Long
End Type

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal lpcbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hkey As Long, phkResult As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Source As Any, ByVal numBytes As Long)

Const REG_OPTION_VOLATILE = 1
Const REG_OPTION_NON_VOLATILE = 0
Const SYNCHRONIZE = &H100000
Const READ_CONTROL = &H20000
Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Const STANDARD_RIGHTS_ALL = &H1F0000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Const KEY_EXECUTE = (KEY_READ)
Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Const ERROR_MORE_DATA = 234
Const ERROR_NO_MORE_ITEMS = &H103
Const ERROR_KEY_NOT_FOUND = &H2
Const REG_CREATED_NEW_KEY = &H1


Dim mvarhKeySet As Long
Dim mvarKeyRoot As String
Dim mvarSubKey As String
Dim mvarErrorMsg As String
Dim Security As SECURITY_ATTRIBUTES

Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Property Get ErrorMsg() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ErrorMsg = mvarErrorMsg
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsRegistry", "ErrorMsg [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get hkey() As hkey
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  hkey = mvarhKeySet
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsRegistry", "hkey [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let hkey(ByVal vData As hkey)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  mvarhKeySet = vData
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsRegistry", "hkey [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get KeyRoot() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  KeyRoot = mvarKeyRoot
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsRegistry", "KeyRoot [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let KeyRoot(ByVal vData As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  mvarKeyRoot = vData
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsRegistry", "KeyRoot [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get SubKey() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  SubKey = mvarSubKey
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsRegistry", "SubKey [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let SubKey(ByVal vData As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  mvarSubKey = vData
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsRegistry", "SubKey [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Function KeyExists() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim handle As Long, ret As Long
50020  If RegOpenKeyEx(mvarhKeySet, GetCompleteKey(mvarKeyRoot & "\UNSINN" & mvarSubKey), 0, KEY_READ, handle) Then
50030   If handle <> 0 Then RegCloseKey (handle)
50040   KeyExists = False
50050   Exit Function
50060  End If
50070  If handle <> 0 Then RegCloseKey (handle)
50080  KeyExists = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsRegistry", "KeyExists")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function CreateKey(Optional Keyname As Variant) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim handle As Long, disp As Long, RetVal As Long, tStr As String
50020  Keyname = IIf(IsMissing(Keyname), mvarSubKey, CStr(Keyname))
50030  If mvarKeyRoot = vbNullString Then
50040    tStr = Keyname
50050   Else
50060    tStr = mvarKeyRoot & "\" & Keyname
50070  End If
50080  RetVal = RegCreateKeyEx(mvarhKeySet, tStr, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, Security, handle, disp)
50090  tStr = String(64, Chr$(0))
50100  mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
50110  If handle <> 0 Then RegCloseKey (handle)
50120  If RetVal Then
50130   Exit Function
50140  End If
50150  CreateKey = RetVal
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsRegistry", "CreateKey")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub DeleteSubkeys(ByVal Keyname As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Const Length As Long = 256
50020  Dim SubKeys As Collection, subKeyName As String, i As Long, tStr As String, handle As Long
50030
50040  If RegOpenKeyEx(mvarhKeySet, GetCompleteKey(mvarKeyRoot & "\" & Keyname), 0&, KEY_ALL_ACCESS, handle) <> ERROR_SUCCESS Then
50050   If handle <> 0 Then RegCloseKey (handle)
50060   Exit Sub
50070  End If
50080
50090  tStr = Space$(Length)
50100  Set SubKeys = New Collection
50110  i = 0
50120
50130  Do
50140   subKeyName = tStr
50150   If RegEnumKey(handle, i, subKeyName, Length) <> ERROR_SUCCESS Then
50160    Exit Do
50170   End If
50180   i = i + 1
50190   subKeyName = Left$(subKeyName, InStr(subKeyName, Chr$(0)) - 1)
50200   SubKeys.Add subKeyName
50210  Loop
50220
50230  For i = 1 To SubKeys.Count
50240   DeleteKeyWithSubkeys Keyname & "\" & SubKeys(i)
50250   RegDeleteKey mvarhKeySet, GetCompleteKey(mvarKeyRoot & "\" & Keyname) & "\" & SubKeys(i)
50260  Next i
50270
50280  If handle <> 0 Then RegCloseKey (handle)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsRegistry", "DeleteSubkeys")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub DeleteKeyWithSubkeys(ByVal Keyname As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  DeleteSubkeys Keyname
50020  DeleteKey Keyname
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsRegistry", "DeleteKeyWithSubkeys")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function DeleteKey(Keyname As String) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim RetVal As Long, handle As Long, tStr As String
50020  RetVal = RegDeleteKey(mvarhKeySet, mvarKeyRoot & "\" & Keyname)
50030  tStr = String(64, Chr$(0))
50040  mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
50050  If handle <> 0 Then RegCloseKey (handle)
50060  If RetVal Then
50070   Exit Function
50080  End If
50090  DeleteKey = RetVal
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsRegistry", "DeleteKey")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function DeleteValue(ValueName As String) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim RetVal As Long, handle As Long, tStr As String
50020  RetVal = RegOpenKeyEx(mvarhKeySet, GetCompleteKey(mvarKeyRoot & "\" & mvarSubKey), REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, handle)
50030  tStr = String(64, Chr$(0))
50040  mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
50050  If RetVal <> 0 Then
50060   DeleteValue = RetVal
50070   If handle <> 0 Then RegCloseKey (handle)
50080   Exit Function
50090  End If
50100  DeleteValue = RegDeleteValue(handle, ValueName)
50110  If handle <> 0 Then RegCloseKey (handle)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsRegistry", "DeleteValue")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function EnumRegistryKeys(ByVal hkey As hkey, ByVal Keyname As String) As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim handle As Long, Length As Long, Index As Long, subKeyName As String, _
  fFiletime As FILETIME
50030  Set EnumRegistryKeys = New Collection
50040
50050  If Len(Keyname) Then
50060   If RegOpenKeyEx(hkey, Keyname, 0, KEY_READ, handle) Then Exit Function
50070   hkey = handle
50080  End If
50090
50100  Do
50110   Length = 260
50120   subKeyName = Space$(Length)
50130   If RegEnumKeyEx(hkey, Index, subKeyName, Length, 0, "", vbNull, fFiletime) = ERROR_NO_MORE_ITEMS Then Exit Do
50140   subKeyName = Left$(subKeyName, InStr(subKeyName, vbNullChar) - 1)
50150   EnumRegistryKeys.Add subKeyName, subKeyName
50160   Index = Index + 1
50170  Loop
50180
50190  If handle <> 0 Then RegCloseKey (handle)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsRegistry", "EnumRegistryKeys")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Function EnumRegistryValues(ByVal hkey As hkey, ByVal Keyname As String) As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim handle As Long, Index As Long, valueType As Long, Name As String, nameLen As Long, _
  resLong As Long, resString As String, Length As Long, valueInfo(0 To 1) As Variant, _
  RetVal As Long, i As Integer, vTemp As Variant
50040
50050  Set EnumRegistryValues = New Collection
50060
50070  If Len(Keyname) Then
50080   If RegOpenKeyEx(hkey, Keyname, 0, KEY_READ, handle) Then
50090    If handle <> 0 Then RegCloseKey (handle)
50100    Exit Function
50110   End If
50120   hkey = handle
50130  End If
50140
50150  Do
50160   nameLen = 260
50170   Name = Space$(nameLen)
50180   Length = 4096
50190   ReDim resBinary(0 To Length - 1) As Byte
50200
50210   RetVal = RegEnumValue(hkey, Index, Name, nameLen, ByVal 0&, valueType, resBinary(0), Length)
50220
50230   If RetVal = ERROR_MORE_DATA Then
50240    ReDim resBinary(0 To Length - 1) As Byte
50250    RetVal = RegEnumValue(hkey, Index, Name, nameLen, ByVal 0&, valueType, resBinary(0), Length)
50260   End If
50270   If RetVal Then Exit Do
50280
50290   valueInfo(0) = Left$(Name, nameLen)
50300
50311   Select Case valueType
         Case REG_DWORD
50330     CopyMemory resLong, resBinary(0), 4
50340     valueInfo(1) = resLong
50350    Case REG_SZ
50360     If Length <> 0 Then
50370       resString = Space$(Length - 1)
50380       CopyMemory ByVal resString, resBinary(0), Length - 1
50390       valueInfo(1) = resString
50400      Else
50410       valueInfo(1) = vbNullString
50420     End If
50430    Case REG_EXPAND_SZ
50440     If Length <> 0 Then
50450       resString = Space$(Length - 1)
50460       CopyMemory ByVal resString, resBinary(0), Length - 1
50470       Length = ExpandEnvironmentStrings(resString, resString, Len(resString) + 2)
50480       valueInfo(1) = TrimNull(resString)
50490      Else
50500       valueInfo(1) = vbNullString
50510     End If
50520    Case REG_BINARY
50530     If Length < UBound(resBinary) + 1 Then
50540      ReDim Preserve resBinary(0 To Length - 1) As Byte
50550     End If
50560     resString = ""
50570     For i = 0 To UBound(resBinary)
50580      If Len(Trim(Hex(resBinary(i)))) < 2 Then
50590        resString = resString & " 0" & Trim(Hex(resBinary(i)))
50600       Else
50610        resString = resString & " " & Trim(Hex(resBinary(i)))
50620      End If
50630     Next i
50640     valueInfo(1) = LTrim(resString) 'Get rid of leading space
50650    Case REG_MULTI_SZ
50660     resString = Space$(Length - 2)
50670     CopyMemory ByVal resString, resBinary(0), Length - 2
50680     resString = Replace(resString, vbNullChar, ",", , , vbBinaryCompare)
50690     valueInfo(1) = resString
50700    Case Else
50710   End Select
50720   EnumRegistryValues.Add valueInfo, valueInfo(0)
50730   Index = Index + 1
50740  Loop
50750  If handle <> 0 Then RegCloseKey (handle)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsRegistry", "EnumRegistryValues")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetRegistryValue(ByVal ValueName As String, Optional DefaultValue As Variant) As Variant
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim handle As Long, resLong As Long, resString As String, TestString As String, _
  resBinary() As Byte, Length As Long, RetVal As Long, valueType As Long, expStr As String
50030
50040  GetRegistryValue = IIf(IsMissing(DefaultValue), Empty, DefaultValue)
50050  If RegOpenKeyEx(mvarhKeySet, GetCompleteKey(mvarKeyRoot & "\" & mvarSubKey), REG_OPTION_NON_VOLATILE, KEY_READ, handle) Then
50060   If handle <> 0 Then RegCloseKey (handle)
50070   Exit Function
50080  End If
50090
50100  Length = 1024
50110  ReDim resBinary(0 To Length - 1) As Byte
50120
50130  RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), Length)
50140  If RetVal = ERROR_MORE_DATA Then
50150   ReDim resBinary(0 To Length - 1) As Byte
50160   RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), Length)
50170  End If
50180
50190  If RetVal = ERROR_KEY_NOT_FOUND Then
50200   If handle <> 0 Then RegCloseKey (handle)
50210   Exit Function
50220  End If
50230
50241  Select Case valueType
        Case REG_DWORD
50260    CopyMemory resLong, resBinary(0), 4
50270    GetRegistryValue = resLong
50280   Case REG_SZ
50290    If Length <> 0 Then
50300     resString = Space$(Length - 1)
50310     CopyMemory ByVal resString, resBinary(0), Length - 1
50320     GetRegistryValue = resString
50330    End If
50340   Case REG_EXPAND_SZ
50350    If Length <> 0 Then
50360     resString = Space$(Length - 1)
50370     CopyMemory ByVal resString, resBinary(0), Length - 1
50380     expStr = ""
50390     Length = ExpandEnvironmentStrings(resString, expStr, 0)
50400     expStr = Space$(Length + Len(vbNullChar))
50410     Length = ExpandEnvironmentStrings(resString, expStr, Len(expStr))
50420     Length = InStr(1&, expStr, vbNullChar)
50430     If Length <> 0& Then
50440      expStr = Left$(expStr, Length - 1&)
50450     End If
50460     GetRegistryValue = expStr
50470    End If
50480   Case REG_BINARY
50490    If Length <> UBound(resBinary) + 1 Then
50500     ReDim Preserve resBinary(0 To Length - 1) As Byte
50510    End If
50520    GetRegistryValue = resBinary()
50530   Case REG_MULTI_SZ
50540    resString = Space$(Length - 2)
50550    CopyMemory ByVal resString, resBinary(0), Length - 2
50560    TestString = resString
50570    If Len(TrimNull(TestString)) > 0 Then GetRegistryValue = resString
50580   Case Else
50590  End Select
50600  If handle <> 0 Then RegCloseKey (handle)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsRegistry", "GetRegistryValue")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function SetRegistryValue(ByVal ValueName As String, value As Variant, DType As DataType) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim handle As Long, lngValue As Long, strValue As String, binValue() As Byte, Length As Long, _
  RetVal As Long, tStr As String
50030
50040  If RegOpenKeyEx(hkey, GetCompleteKey(mvarKeyRoot & "\" & mvarSubKey), REG_OPTION_NON_VOLATILE, KEY_WRITE, handle) Then
50050   SetRegistryValue = False 'CVar("Error!")
50060   If handle <> 0 Then RegCloseKey (handle)
50070   Exit Function
50080  End If
50090
50101  Select Case DType
        Case REG_DWORD
50120    lngValue = value
50130    RetVal = RegSetValueEx(handle, ValueName, 0, REG_DWORD, lngValue, 4)
50140    tStr = String(64, Chr$(0))
50150    mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
50160   Case REG_SZ
50170    strValue = value
50180    If Len(strValue) = 0 Then
50190      RetVal = RegSetValueEx(handle, ValueName, 0, REG_SZ, ByVal strValue, 0)
50200     Else
50210      RetVal = RegSetValueEx(handle, ValueName, 0, REG_SZ, ByVal strValue, Len(strValue) + 1)
50220    End If
50230    tStr = String(64, Chr$(0))
50240    mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
50250   Case REG_BINARY
50260    binValue = value
50270    Length = UBound(binValue) - LBound(binValue) + 1
50280    RetVal = RegSetValueEx(handle, ValueName, 0, REG_BINARY, binValue(LBound(binValue)), Length)
50290    tStr = String(64, Chr$(0))
50300    mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
50310   Case REG_EXPAND_SZ
50320    strValue = value
50330    If Len(strValue) = 0 Then
50340      RetVal = RegSetValueEx(handle, ValueName, 0, REG_EXPAND_SZ, ByVal strValue, 0)
50350     Else
50360      RetVal = RegSetValueEx(handle, ValueName, 0, REG_EXPAND_SZ, ByVal strValue, Len(strValue) + 1)
50370    End If
50380    tStr = String(64, Chr$(0))
50390    mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
50400   Case REG_MULTI_SZ
50410    strValue = value
50420    If Len(strValue) = 0 Then
50430      RetVal = RegSetValueEx(handle, ValueName, 0, REG_MULTI_SZ, ByVal strValue, 0)
50440     Else
50450      RetVal = RegSetValueEx(handle, ValueName, 0, REG_MULTI_SZ, ByVal strValue, Len(strValue) + 1)
50460    End If
50470    tStr = String(64, Chr$(0))
50480    mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
50490   Case Else
50500  End Select
50510  If handle <> 0 Then RegCloseKey (handle)
50520  SetRegistryValue = (RetVal = 0)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsRegistry", "SetRegistryValue")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function ReadRemoteRegistryValue(ByVal sRemoteComputer As String, ByVal hkey As hkey, ByVal ValueName As String, Optional KeyPath As String) As Variant
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim handle As Long, lReturnCode As Long, lHive As Long, lhRemoteRegistry As Long, valueType As Long, _
  resLong As Long, resString As String, TestString As String, resBinary() As Byte, Length As Long, _
  RetVal As Long, RegPath As String, tStr As String, expStr As String
50040
50050  RegPath = IIf(IsMissing(KeyPath), GetCompleteKey(mvarKeyRoot & "\" & mvarSubKey), KeyPath)
50060
50070  If RegConnectRegistry(sRemoteComputer, hkey, lhRemoteRegistry) Then
50080   ReadRemoteRegistryValue = CVar("Error!")
50090   Exit Function
50100  End If
50110  lReturnCode = RegOpenKeyEx(lhRemoteRegistry, RegPath, 0, KEY_ALL_ACCESS, handle)
50120
50130  Length = 1024
50140  ReDim resBinary(0 To Length - 1) As Byte
50150  RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), Length)
50160
50170  If RetVal = ERROR_MORE_DATA Then
50180   ReDim resBinary(0 To Length - 1) As Byte
50190   RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), Length)
50200   tStr = String(64, Chr$(0))
50210   mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
50220  End If
50230
50240  If RetVal = ERROR_KEY_NOT_FOUND Then
50250   If handle <> 0 Then RegCloseKey (handle)
50260   Exit Function
50270  End If
50280
50291  Select Case valueType
        Case REG_DWORD
50310    CopyMemory resLong, resBinary(0), 4
50320    ReadRemoteRegistryValue = resLong
50330   Case REG_SZ
50340    If Length <> 0 Then
50350     resString = Space$(Length - 1)
50360     CopyMemory ByVal resString, resBinary(0), Length - 1
50370     ReadRemoteRegistryValue = resString
50380    End If
50390   Case REG_EXPAND_SZ
50400    If Length <> 0 Then
50410     resString = Space$(Length - 1)
50420     CopyMemory ByVal resString, resBinary(0), Length - 1
50430     expStr = ""
50440     Length = ExpandEnvironmentStrings(resString, expStr, 0)
50450     expStr = String$(Length - 1, 0)
50460     Length = ExpandEnvironmentStrings(resString, expStr, Length)
50470     ReadRemoteRegistryValue = expStr
50480    End If
50490   Case REG_BINARY
50500    If Length <> UBound(resBinary) + 1 Then
50510     ReDim Preserve resBinary(0 To Length - 1) As Byte
50520    End If
50530    ReadRemoteRegistryValue = resBinary()
50540   Case REG_MULTI_SZ
50550    resString = Space$(Length - 2)
50560    CopyMemory ByVal resString, resBinary(0), Length - 2
50570    TestString = resString
50580    If Len(TrimNull(TestString)) > 0 Then ReadRemoteRegistryValue = resString
50590   Case Else
50600  End Select
50610  If handle <> 0 Then RegCloseKey (handle)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsRegistry", "ReadRemoteRegistryValue")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function WriteRemoteRegistryValue(ByVal sRemoteComputer As String, ByVal hkey As hkey, ByVal ValueName As String, value As Variant, DType As DataType, Optional KeyPath As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim handle As Long, lngValue As Long, strValue As String, binValue() As Byte, Length As Long, _
  RetVal As Long, RegPath As String, lhRemoteRegistry As Long, lReturnCode As Long, tStr As String
50030
50040  RegPath = IIf(IsMissing(KeyPath), GetCompleteKey(mvarKeyRoot & "\" & mvarSubKey), KeyPath)
50050
50060  If RegConnectRegistry(sRemoteComputer, hkey, lhRemoteRegistry) Then
50070   WriteRemoteRegistryValue = CVar("Error!")
50080   Exit Function
50090  End If
50100  lReturnCode = RegOpenKeyEx(lhRemoteRegistry, RegPath, 0, KEY_ALL_ACCESS, handle)
50110
50121  Select Case DType
        Case REG_DWORD
50140    lngValue = value
50150    RetVal = RegSetValueEx(handle, ValueName, 0, REG_DWORD, lngValue, 4)
50160    tStr = String(64, Chr$(0))
50170    mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
50180   Case REG_SZ
50190    strValue = value
50200    RetVal = RegSetValueEx(handle, ValueName, 0, REG_SZ, ByVal strValue, Len(strValue))
50210    tStr = String(64, Chr$(0))
50220    mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
50230   Case REG_BINARY
50240    binValue = value
50250    Length = UBound(binValue) - LBound(binValue) + 1
50260    RetVal = RegSetValueEx(handle, ValueName, 0, REG_BINARY, binValue(LBound(binValue)), Length)
50270    tStr = String(64, Chr$(0))
50280    mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
50290   Case REG_EXPAND_SZ
50300    strValue = value
50310    RetVal = RegSetValueEx(handle, ValueName, 0, REG_EXPAND_SZ, ByVal strValue, Len(strValue))
50320    tStr = String(64, Chr$(0))
50330    mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
50340   Case REG_MULTI_SZ
50350    strValue = value
50360    RetVal = RegSetValueEx(handle, ValueName, 0, REG_MULTI_SZ, ByVal strValue, Len(strValue))
50370    tStr = String(64, Chr$(0))
50380    mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
50390   Case Else
50400  End Select
50410  If handle <> 0 Then RegCloseKey (handle)
50420  WriteRemoteRegistryValue = (RetVal = 0)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsRegistry", "WriteRemoteRegistryValue")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function TrimNull(Item As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim pos As Integer
50020  pos = InStr(Item, Chr$(0))
50030  If pos Then Item = Left$(Item, pos - 1)
50040  TrimNull = Item
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsRegistry", "TrimNull")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function GetCompleteKey(KeyString As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStr As String
50020  GetCompleteKey = KeyString
50030  If LenB(KeyString) > 0 Then
50040   tStr = KeyString
50050   Do While Right$(tStr, 1) = "\"
50060    tStr = Mid$(tStr, 1, Len(tStr) - 1)
50070    DoEvents
50080   Loop
50090   GetCompleteKey = tStr
50100  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsRegistry", "GetCompleteKey")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
