Attribute VB_Name = "modLanguage"
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public strLangdata As String
'Public colLangFiles As Collection

Public Function GetIni(ByVal FileName As String, ByVal AppName As String, ByVal KeyName As String, ByVal Default As String) As String
   Dim prv              As Long, retVal As String
   retVal = String$(255, 0)
   prv = GetPrivateProfileString(AppName, KeyName, "", retVal, Len(retVal), FileName)
   If prv = 0 Then
      GetIni = Default
   Else
      GetIni = left(retVal, InStr(retVal, Chr(0)) - 1)
   End If
End Function

Public Function SetIni(ByVal FileName As String, ByVal AppName As String, ByVal KeyName As String, ByVal KeyValue As String) As Long
   SetIni = WritePrivateProfileString(AppName, KeyName, KeyValue, FileName)
End Function

Public Function DelIniApp(ByVal FileName As String, ByVal AppName As String) As Long
   DelIniApp = WritePrivateProfileString(AppName, vbNullString, "", FileName)
End Function

Public Function DelIniKey(ByVal FileName As String, ByVal AppName As String, ByVal KeyName As String) As Long
   DelIniKey = WritePrivateProfileString(AppName, KeyName, vbNullString, FileName)
End Function

Public Function GetIniAllApps(ByVal FileName As String, AppsCollection As Collection)
   Dim retVal           As String
   Dim tmp              As String, tmp2 As String
   Set AppsCollection = New Collection
   retVal = String$(1024, 0)
   prv = GetPrivateProfileString(vbNullString, vbNullString, "", retVal, Len(retVal), FileName)
   retVal = left(retVal, prv)
   For i = 1 To Len(retVal)
      tmp2 = Mid(retVal, i, 1)
      If tmp2 <> Chr(0) Then
         tmp = tmp & tmp2
      ElseIf tmp2 = Chr(0) And tmp <> "" Then
         AppsCollection.Add tmp
         tmp = ""
      End If
   Next
   GetIniAllApps = AppsCollection.Count
End Function

Public Function GetIniAppAllKeys(ByVal FileName As String, ByVal AppName As String, KeysCollection As Collection)
   Dim retVal           As String
   Dim tmp              As String, tmp2 As String
   Set KeysCollection = New Collection
   retVal = String$(1024, 0)
   prv = GetPrivateProfileString(AppName, vbNullString, "", retVal, Len(retVal), FileName)
   retVal = left(retVal, prv)
   For i = 1 To Len(retVal)
      tmp2 = Mid(retVal, i, 1)
      If tmp2 <> Chr(0) Then
         tmp = tmp & tmp2
      ElseIf tmp2 = Chr(0) And tmp <> "" Then
         KeysCollection.Add tmp
         tmp = ""
      End If
   Next
   GetIniAppAllKeys = KeysCollection.Count
End Function

Public Function FindKey(strKey As String, colValues As Collection)
FindKey = 0
If colValues.Count < 0 Then Exit Function
For i = 1 To colValues.Count
If colValues(i) = strKey Then FindKey = i
Next i
End Function
