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
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hkey As Long, phkResult As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal numBytes As Long)

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

Enum DataType
 REG_SZ = &H1
 REG_EXPAND_SZ = &H2
 REG_BINARY = &H3
 REG_DWORD = &H4
 REG_MULTI_SZ = &H7
End Enum

Enum hkey
 HKEY_CLASSES_ROOT = &H80000000
 HKEY_CURRENT_USER = &H80000001
 HKEY_LOCAL_MACHINE = &H80000002
 HKEY_USERS = &H80000003
 HKEY_PERFORMANCE_DATA = &H80000004
 HKEY_CURRENT_CONFIG = &H80000005
 HKEY_DYN_DATA = &H80000006
End Enum

Dim mvarhKeySet As Long
Dim mvarKeyRoot As String
Dim mvarSubKey As String
Dim mvarErrorMsg As String
Dim Security As SECURITY_ATTRIBUTES

Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Property Get ErrorMsg() As String
 ErrorMsg = mvarErrorMsg
End Property

Public Property Get hkey() As hkey
 hkey = mvarhKeySet
End Property

Public Property Let hkey(ByVal vData As hkey)
 mvarhKeySet = vData
End Property

Public Property Get KeyRoot() As String
 KeyRoot = mvarKeyRoot
End Property

Public Property Let KeyRoot(ByVal vData As String)
 mvarKeyRoot = vData
End Property

Public Property Get Subkey() As String
 Subkey = mvarSubKey
End Property

Public Property Let Subkey(ByVal vData As String)
 mvarSubKey = vData
End Property

Public Function KeyExists() As Boolean
 Dim handle As Long, ret As Long
 If RegOpenKeyEx(mvarhKeySet, mvarKeyRoot & "\" & mvarSubKey, 0, KEY_READ, handle) Then
  KeyExists = False
  Exit Function
 End If
 KeyExists = True
End Function

Public Function CreateKey(Optional Keyname As Variant) As String
 Dim handle As Long, disp As Long, RetVal As Long, tStr As String
 Keyname = IIf(IsMissing(Keyname), mvarSubKey, CStr(Keyname))
 If mvarKeyRoot = "" Then
   tStr = Keyname
  Else
   tStr = mvarKeyRoot & "\" & Keyname
 End If
 RetVal = RegCreateKeyEx(mvarhKeySet, tStr, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, Security, handle, disp)
 tStr = String(64, Chr$(0))
 mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
 If RetVal Then Exit Function
 RegCloseKey (handle)
 CreateKey = RetVal
End Function

Public Function DeleteKey(Keyname As String) As Long
 Dim RetVal As Long, handle As Long, tStr As String
 RetVal = RegDeleteKey(mvarhKeySet, mvarKeyRoot & "\" & Keyname)
 tStr = String(64, Chr$(0))
 mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
 If RetVal Then Exit Function
 RegCloseKey (handle)
 DeleteKey = RetVal
End Function

Public Function DeleteValue(ValueName As String) As Long
 Dim RetVal As Long, handle As Long, tStr As String
 RetVal = RegOpenKeyEx(mvarhKeySet, mvarKeyRoot & "\" & mvarSubKey, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, handle)
 tStr = String(64, Chr$(0))
 mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
 If RetVal <> 0 Then
  DeleteValue = RetVal
  Exit Function
 End If
 DeleteValue = RegDeleteValue(handle, ValueName)
 RegCloseKey (handle)
End Function

Public Function EnumRegistryKeys(ByVal hkey As hkey, ByVal Keyname As String) As _
                Collection
 Dim handle As Long, Length As Long, Index As Long, subkeyName As String, _
  fFiletime As FILETIME
 Set EnumRegistryKeys = New Collection

 If Len(Keyname) Then
  If RegOpenKeyEx(hkey, Keyname, 0, KEY_READ, handle) Then Exit Function
  hkey = handle
 End If

 Do
  Length = 260
  subkeyName = Space$(Length)
  If RegEnumKeyEx(hkey, Index, subkeyName, Length, 0, "", vbNull, fFiletime) = ERROR_NO_MORE_ITEMS Then Exit Do
  subkeyName = Left$(subkeyName, InStr(subkeyName, vbNullChar) - 1)
  EnumRegistryKeys.Add subkeyName, subkeyName
  Index = Index + 1
 Loop

 If handle Then RegCloseKey handle
End Function

Function EnumRegistryValues(ByVal hkey As hkey, ByVal Keyname As String) As _
    Collection
    Dim handle As Long
    Dim Index As Long
    Dim valueType As Long
    Dim Name As String
    Dim nameLen As Long
    Dim resLong As Long
    Dim resString As String
    Dim Length As Long
    Dim valueInfo(0 To 1) As Variant
    Dim RetVal As Long
    Dim i As Integer
    Dim vTemp As Variant

    Set EnumRegistryValues = New Collection

    If Len(Keyname) Then
        If RegOpenKeyEx(hkey, Keyname, 0, KEY_READ, handle) Then Exit Function
        hkey = handle
    End If

    Do
        nameLen = 260
        Name = Space$(nameLen)
        Length = 4096
        ReDim resBinary(0 To Length - 1) As Byte

        RetVal = RegEnumValue(hkey, Index, Name, nameLen, ByVal 0&, valueType, _
            resBinary(0), Length)

        If RetVal = ERROR_MORE_DATA Then
            ReDim resBinary(0 To Length - 1) As Byte
            RetVal = RegEnumValue(hkey, Index, Name, nameLen, ByVal 0&, _
                valueType, resBinary(0), Length)
        End If
        If RetVal Then Exit Do

        valueInfo(0) = Left$(Name, nameLen)

        Select Case valueType
            
            Case REG_DWORD
                CopyMemory resLong, resBinary(0), 4
                valueInfo(1) = resLong

            Case REG_SZ
                If Length <> 0 Then
                    resString = Space$(Length - 1)
                    CopyMemory ByVal resString, resBinary(0), Length - 1
                    valueInfo(1) = resString
                Else
                    valueInfo(1) = vbNullString
                End If

            Case REG_EXPAND_SZ
                If Length <> 0 Then
                    resString = Space$(Length - 1)
                    CopyMemory ByVal resString, resBinary(0), Length - 1
                    Length = ExpandEnvironmentStrings(resString, resString, Len(resString))
                    valueInfo(1) = TrimNull(resString)
                Else
                    valueInfo(1) = vbNullString
                End If

            Case REG_BINARY
                If Length < UBound(resBinary) + 1 Then
                    ReDim Preserve resBinary(0 To Length - 1) As Byte
                End If
                    For i = 0 To UBound(resBinary)
                         resString = resString & " " & Format(Trim(Hex(resBinary(i))), "0#")
                    Next i
                    valueInfo(1) = LTrim(resString) 'Get rid of leading space

            Case REG_MULTI_SZ
                resString = Space$(Length - 2)
                CopyMemory ByVal resString, resBinary(0), Length - 2

                resString = Replace(resString, vbNullChar, ",", , , vbBinaryCompare)
                valueInfo(1) = resString

            Case Else
        End Select

        EnumRegistryValues.Add valueInfo, valueInfo(0)

        Index = Index + 1
    Loop

    If handle Then RegCloseKey handle

End Function

Public Function GetRegistryValue(ByVal ValueName As String, Optional DefaultValue As Variant) As Variant
    Dim handle As Long
    Dim resLong As Long
    Dim resString As String
    Dim TestString As String
    Dim resBinary() As Byte
    Dim Length As Long
    Dim RetVal As Long
    Dim valueType As Long

        GetRegistryValue = IIf(IsMissing(DefaultValue), Empty, DefaultValue)

        If RegOpenKeyEx(mvarhKeySet, mvarKeyRoot & "\" & mvarSubKey, REG_OPTION_NON_VOLATILE, KEY_READ, handle) Then
           Exit Function
        End If

        Length = 1024
        ReDim resBinary(0 To Length - 1) As Byte

        RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), _
            Length)
        If RetVal = ERROR_MORE_DATA Then
            ReDim resBinary(0 To Length - 1) As Byte
            RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), _
                Length)
        End If

        If RetVal = ERROR_KEY_NOT_FOUND Then
                 RegCloseKey (handle)
                 Exit Function
        End If

        Select Case valueType
            Case REG_DWORD
                CopyMemory resLong, resBinary(0), 4
                GetRegistryValue = resLong

            Case REG_SZ
                If Length <> 0 Then
                    resString = Space$(Length - 1)
                    CopyMemory ByVal resString, resBinary(0), Length - 1
                    GetRegistryValue = resString
                End If

            Case REG_EXPAND_SZ
                If Length <> 0 Then
                    resString = Space$(Length - 1)
                    CopyMemory ByVal resString, resBinary(0), Length - 1

                    Length = ExpandEnvironmentStrings(resString, resString, Len(resString))
                    GetRegistryValue = Left$(resString, Length)

                End If

            Case REG_BINARY
                If Length <> UBound(resBinary) + 1 Then
                    ReDim Preserve resBinary(0 To Length - 1) As Byte
                End If
                GetRegistryValue = resBinary()

            Case REG_MULTI_SZ
                resString = Space$(Length - 2)
                CopyMemory ByVal resString, resBinary(0), Length - 2
                TestString = resString
                If Len(TrimNull(TestString)) > 0 Then GetRegistryValue = resString

            Case Else
        End Select

     RegCloseKey (handle)

End Function

Public Function SetRegistryValue(ByVal ValueName As String, Value As Variant, DType As DataType) As Boolean
    Dim handle As Long
    Dim lngValue As Long
    Dim strValue As String
    Dim binValue() As Byte
    Dim Length As Long
    Dim RetVal As Long, tStr As String

    If RegOpenKeyEx(hkey, mvarKeyRoot & "\" & mvarSubKey, REG_OPTION_NON_VOLATILE, KEY_WRITE, handle) Then
       SetRegistryValue = False 'CVar("Error!")
       Exit Function
    End If

    Select Case DType
        Case REG_DWORD
            lngValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_DWORD, lngValue, 4)
            tStr = String(64, Chr$(0))
            mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
        Case REG_SZ
            strValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_SZ, ByVal strValue, _
                Len(strValue))
            tStr = String(64, Chr$(0))
            mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
        Case REG_BINARY
            binValue = Value
            Length = UBound(binValue) - LBound(binValue) + 1
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_BINARY, _
                                   binValue(LBound(binValue)), Length)
            tStr = String(64, Chr$(0))
            mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
        Case REG_EXPAND_SZ
            strValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_EXPAND_SZ, ByVal strValue, _
                Len(strValue))
            tStr = String(64, Chr$(0))
            mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
        Case REG_MULTI_SZ
            strValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_MULTI_SZ, ByVal strValue, _
                Len(strValue))
            tStr = String(64, Chr$(0))
            mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
        Case Else
    End Select

     RegCloseKey (handle)
    SetRegistryValue = (RetVal = 0)

End Function

Public Function ReadRemoteRegistryValue(ByVal sRemoteComputer As String, ByVal hkey As hkey, ByVal ValueName As String, Optional KeyPath As String) As Variant

    Dim handle As Long
    Dim lReturnCode, lHive, lhRemoteRegistry As Long
    Dim valueType As Long
    Dim resLong As Long
    Dim resString As String
    Dim TestString As String
    Dim resBinary() As Byte
    Dim Length As Long
    Dim RetVal As Long
    Dim RegPath As String, tStr As String


        RegPath = IIf(IsMissing(KeyPath), mvarKeyRoot & "\" & mvarSubKey, KeyPath)

        If RegConnectRegistry(sRemoteComputer, hkey, lhRemoteRegistry) Then
            ReadRemoteRegistryValue = CVar("Error!")
            Exit Function
        End If
        lReturnCode = RegOpenKeyEx(lhRemoteRegistry, RegPath, 0, KEY_ALL_ACCESS, handle)

        Length = 1024
        ReDim resBinary(0 To Length - 1) As Byte

        RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), _
            Length)

        If RetVal = ERROR_MORE_DATA Then
            ReDim resBinary(0 To Length - 1) As Byte
            RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), _
                Length)
            tStr = String(64, Chr$(0))
            mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
        End If

        If RetVal = ERROR_KEY_NOT_FOUND Then
            RegCloseKey (handle)
            Exit Function
        End If

        Select Case valueType
            Case REG_DWORD
                CopyMemory resLong, resBinary(0), 4
                ReadRemoteRegistryValue = resLong

            Case REG_SZ
                If Length <> 0 Then
                    resString = Space$(Length - 1)
                    CopyMemory ByVal resString, resBinary(0), Length - 1
                    ReadRemoteRegistryValue = resString
                End If

            Case REG_EXPAND_SZ
                If Length <> 0 Then
                    resString = Space$(Length - 1)
                    CopyMemory ByVal resString, resBinary(0), Length - 1

                    Length = ExpandEnvironmentStrings(resString, resString, Len(resString))
                    ReadRemoteRegistryValue = Left$(resString, Length)

                End If

            Case REG_BINARY
                If Length <> UBound(resBinary) + 1 Then
                    ReDim Preserve resBinary(0 To Length - 1) As Byte
                End If
                ReadRemoteRegistryValue = resBinary()

            Case REG_MULTI_SZ
                resString = Space$(Length - 2)
                CopyMemory ByVal resString, resBinary(0), Length - 2
                TestString = resString
                If Len(TrimNull(TestString)) > 0 Then ReadRemoteRegistryValue = resString

            Case Else
        End Select

        RegCloseKey (handle)

End Function

Public Function WriteRemoteRegistryValue(ByVal sRemoteComputer As String, ByVal hkey As hkey, ByVal ValueName As String, Value As Variant, DType As DataType, Optional KeyPath As String) As Boolean
    Dim handle As Long
    Dim lngValue As Long
    Dim strValue As String
    Dim binValue() As Byte
    Dim Length As Long
    Dim RetVal As Long
    Dim RegPath As String
    Dim lhRemoteRegistry As Long
    Dim lReturnCode     As Long, tStr As String


        RegPath = IIf(IsMissing(KeyPath), mvarKeyRoot & "\" & mvarSubKey, KeyPath)

        If RegConnectRegistry(sRemoteComputer, hkey, lhRemoteRegistry) Then
            WriteRemoteRegistryValue = CVar("Error!")
            Exit Function
        End If
        lReturnCode = RegOpenKeyEx(lhRemoteRegistry, RegPath, 0, KEY_ALL_ACCESS, handle)

    Select Case DType
        Case REG_DWORD
            lngValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_DWORD, lngValue, 4)
            tStr = String(64, Chr$(0))
            mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
        Case REG_SZ
            strValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_SZ, ByVal strValue, _
                Len(strValue))
            tStr = String(64, Chr$(0))
            mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
        Case REG_BINARY
            binValue = Value
            Length = UBound(binValue) - LBound(binValue) + 1
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_BINARY, _
                                   binValue(LBound(binValue)), Length)
            tStr = String(64, Chr$(0))
            mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
        Case REG_EXPAND_SZ
            strValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_EXPAND_SZ, ByVal strValue, _
                Len(strValue))
            tStr = String(64, Chr$(0))
            mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
        Case REG_MULTI_SZ
            strValue = Value
            RetVal = RegSetValueEx(handle, ValueName, 0, REG_MULTI_SZ, ByVal strValue, _
                Len(strValue))
            tStr = String(64, Chr$(0))
            mvarErrorMsg = Left$(tStr, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, RetVal, 0&, tStr, Len(tStr), ByVal 0))
        Case Else
    End Select

     RegCloseKey (handle)
    WriteRemoteRegistryValue = (RetVal = 0)
End Function

Public Function TrimNull(item As String) As String
 Dim pos As Integer
 pos = InStr(item, Chr$(0))
 If pos Then item = Left$(item, pos - 1)
 TrimNull = item
End Function
