VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'---[ Associative Array for VB, v1.0 ]---------------------'
'                                                          '
' Associative Array for VB is Copyright (C) 2000 Ashley    '
' Ross, aka Eldren.                                        '
' It is under the GNU GPL, version 2 or greater. This      '
' means that you may use and modify it freely, within the  '
' terms of the GPL. See http://www.fsf.org/ for more info. '
'                                                          '
' If you make any changes to this, please tell me about    '
' them, possibly including the newer copy in your email.   '
' Thanks.                                                  '
'                                                          '
'---------[ Ashley Ross aka Eldren ~ eldrenr@yahoo.com ]---'

Option Explicit
Option Compare Text

Private aKey() As String, aVal() As Variant, Upper As Long, Last As Long

Public Sub Add(Key As String, Value As Variant)
 Dim i As Long, thisKey As String, found As Boolean

 ReDim Preserve aKey(Upper)
 ReDim Preserve aVal(Upper)

 For i = 0 To Upper
  thisKey = aKey(i)
  If StrComp(thisKey, Key) = 0 Then
   found = True
   Exit For
  End If
 Next i

 If found Then
   aKey(i) = Key
   aVal(i) = Value
   Last = i
  Else
   aKey(Upper) = Key
   aVal(Upper) = Value
   Last = Upper
   Upper = Upper + 1
 End If
End Sub

Public Function Retrieve(Key As String, Optional Default As Variant = 0) As Variant
 Dim i As Long, thisKey As String, found As Boolean

 If Upper > 0 Then
  For i = 0 To Upper - 1
   thisKey = aKey(i)
   If StrComp(thisKey, Key) = 0 Then
    found = True
    Exit For
   End If
  Next i

  If found Then
   Retrieve = aVal(i)
   Last = i
   Exit Function
  End If
 End If

 If Default = 0 Then
   Retrieve = -1
  Else
   Retrieve = Default
 End If
End Function

Public Sub Delete(Key As String)
 Dim i As Long, i2 As Long, thisKey As String, found As Boolean

 If Upper = 0 Then Exit Sub
 For i = 0 To Upper - 1
  thisKey = aKey(i)
  If StrComp(thisKey, Key) = 0 Then
   found = True
   Exit For
  End If
 Next i

 If found = False Then Exit Sub

 Last = i

 i = i + 1
 For i2 = i To Upper - 1
  aKey(i2 - 1) = aKey(i2)
  aVal(i2 - 1) = aVal(i2)
 Next i2

 Upper = Upper - 1
 ReDim Preserve aKey(Upper)
 ReDim Preserve aVal(Upper)
End Sub

Public Function Keys() As String()
 Dim i As Long, Key() As String
 ReDim Key(0 To Upper - 1)

 For i = 0 To Upper - 1
  Key(i) = aKey(i)
 Next i

 Keys = Key
End Function

Public Function Values() As Variant()
 Dim i As Long, Val() As Variant
 ReDim Val(0 To Upper - 1)

 For i = 0 To Upper - 1
  Val(i) = aVal(i)
 Next i

 Values = Val
End Function

Public Function Every() As Variant()
 Dim KeyVal() As Variant
 ReDim KeyVal(0 To 1)

 Last = Last + 1
 If Last > Upper - 1 Then Last = 0
 KeyVal(0) = aKey(Last)
 KeyVal(1) = aVal(Last)
 Every = KeyVal
End Function
