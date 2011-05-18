Attribute VB_Name = "modCollAddSorted"
Option Explicit

' *** Den Artikel zu diesem Modul finden Sie unter http://www.aboutvb.de/khw/artikel/khwcolladdsorted.htm ***

Public Function AddSortedStr(Collection As Collection, Item As String, Optional key As Variant, Optional ByVal CompareMethod As VbCompareMethod = vbBinaryCompare) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim nCount As Long
50020     Dim nHigh As Long
50030     Dim nLow As Long
50040     Dim nTest As Long
50050
50060     With Collection
50070         nCount = .Count
50080         If nCount Then
50090             If StrComp(Item, .Item(1), CompareMethod) < 0 Then
50100                 .Add Item, key, 1
50110                 AddSortedStr = 1
50120             ElseIf StrComp(Item, .Item(nCount), CompareMethod) > 0 Then
50130                 .Add Item, key
50140                 AddSortedStr = nCount + 1
50150             Else
50160                 nLow = 1
50170                 nHigh = nCount
50180                 Do
50190                     nTest = (nLow + nHigh) \ 2
50200                     If nTest = nLow Then
50210                         Exit Do
50220                     End If
50231                     Select Case StrComp(Item, .Item(nTest), CompareMethod)
                              Case Is < 0
50250                             nHigh = nTest
50260                         Case 0
50270                             Exit Do
50280                         Case Is > 0
50290                             nLow = nTest
50300                     End Select
50310                 Loop
50320                 If nTest < nCount Then
50330                     Do While StrComp(.Item(nTest + 1), Item, CompareMethod) = 0
50340                         nTest = nTest + 1
50350                         If nTest = nCount Then
50360                             Exit Do
50370                         End If
50380                     Loop
50390                 End If
50400                 .Add Item, key, , nTest
50410                 AddSortedStr = nTest + 1
50420             End If
50430         Else
50440             .Add Item, key
50450             AddSortedStr = 1
50460         End If
50470     End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modCollAddSorted", "AddSortedStr")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function AddSortedInt(Collection As Collection, Item As Integer, Optional key As Variant) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim nCount As Long
50020     Dim nHigh As Long
50030     Dim nLow As Long
50040     Dim nTest As Long
50050     Dim nTestItem As Integer
50060
50070     With Collection
50080         nCount = .Count
50090         If nCount Then
50100             If Item < .Item(1) Then
50110                 .Add Item, key, 1
50120                 AddSortedInt = 1
50130             ElseIf Item > .Item(nCount) Then
50140                 .Add Item, key
50150                 AddSortedInt = nCount + 1
50160             Else
50170                 nLow = 1
50180                 nHigh = nCount
50190                 Do
50200                     nTest = (nLow + nHigh) \ 2
50210                     If nTest = nLow Then
50220                         Exit Do
50230                     End If
50240                     nTestItem = .Item(nTest)
50251                     Select Case Item
                              Case Is < nTestItem
50270                             nHigh = nTest
50280                         Case nTestItem
50290                             Exit Do
50300                         Case Is > nTestItem
50310                             nLow = nTest
50320                     End Select
50330                 Loop
50340                 If nTest < nCount Then
50350                     Do While .Item(nTest + 1) = Item
50360                         nTest = nTest + 1
50370                         If nTest = nCount Then
50380                             Exit Do
50390                         End If
50400                     Loop
50410                 End If
50420                 .Add Item, key, , nTest
50430                 AddSortedInt = nTest + 1
50440             End If
50450         Else
50460             .Add Item, key
50470             AddSortedInt = 1
50480         End If
50490     End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modCollAddSorted", "AddSortedInt")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function AddSortedLng(Collection As Collection, Item As Long, Optional key As Variant) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim nCount As Long
50020     Dim nHigh As Long
50030     Dim nLow As Long
50040     Dim nTest As Long
50050     Dim nTestItem As Long
50060
50070     With Collection
50080         nCount = .Count
50090         If nCount Then
50100             If Item < .Item(1) Then
50110                 .Add Item, key, 1
50120                 AddSortedLng = 1
50130             ElseIf Item > .Item(nCount) Then
50140                 .Add Item, key
50150                 AddSortedLng = nCount + 1
50160             Else
50170                 nLow = 1
50180                 nHigh = nCount
50190                 Do
50200                     nTest = (nLow + nHigh) \ 2
50210                     If nTest = nLow Then
50220                         Exit Do
50230                     End If
50240                     nTestItem = .Item(nTest)
50251                     Select Case Item
                              Case Is < nTestItem
50270                             nHigh = nTest
50280                         Case nTestItem
50290                             Exit Do
50300                         Case Is > nTestItem
50310                             nLow = nTest
50320                     End Select
50330                 Loop
50340                 If nTest < nCount Then
50350                     Do While .Item(nTest + 1) = Item
50360                         nTest = nTest + 1
50370                         If nTest = nCount Then
50380                             Exit Do
50390                         End If
50400                     Loop
50410                 End If
50420                 .Add Item, key, , nTest
50430                 AddSortedLng = nTest + 1
50440             End If
50450         Else
50460             .Add Item, key
50470             AddSortedLng = 1
50480         End If
50490     End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modCollAddSorted", "AddSortedLng")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function AddSortedSng(Collection As Collection, Item As Single, Optional key As Variant) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim nCount As Long
50020     Dim nHigh As Long
50030     Dim nLow As Long
50040     Dim nTest As Long
50050     Dim nTestItem As Single
50060
50070     With Collection
50080         nCount = .Count
50090         If nCount Then
50100             If Item < .Item(1) Then
50110                 .Add Item, key, 1
50120                 AddSortedSng = 1
50130             ElseIf Item > .Item(nCount) Then
50140                 .Add Item, key
50150                 AddSortedSng = nCount + 1
50160             Else
50170                 nLow = 1
50180                 nHigh = nCount
50190                 Do
50200                     nTest = (nLow + nHigh) \ 2
50210                     If nTest = nLow Then
50220                         Exit Do
50230                     End If
50240                     nTestItem = .Item(nTest)
50251                     Select Case Item
                              Case Is < nTestItem
50270                             nHigh = nTest
50280                         Case nTestItem
50290                             Exit Do
50300                         Case Is > nTestItem
50310                             nLow = nTest
50320                     End Select
50330                 Loop
50340                 If nTest < nCount Then
50350                     Do While .Item(nTest + 1) = Item
50360                         nTest = nTest + 1
50370                         If nTest = nCount Then
50380                             Exit Do
50390                         End If
50400                     Loop
50410                 End If
50420                 .Add Item, key, , nTest
50430                 AddSortedSng = nTest + 1
50440             End If
50450         Else
50460             .Add Item, key
50470             AddSortedSng = 1
50480         End If
50490     End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modCollAddSorted", "AddSortedSng")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function AddSortedDbl(Collection As Collection, Item As Double, Optional key As Variant) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim nCount As Long
50020     Dim nHigh As Long
50030     Dim nLow As Long
50040     Dim nTest As Long
50050     Dim nTestItem As Double
50060
50070     With Collection
50080         nCount = .Count
50090         If nCount Then
50100             If Item < .Item(1) Then
50110                 .Add Item, key, 1
50120                 AddSortedDbl = 1
50130             ElseIf Item > .Item(nCount) Then
50140                 .Add Item, key
50150                 AddSortedDbl = nCount + 1
50160             Else
50170                 nLow = 1
50180                 nHigh = nCount
50190                 Do
50200                     nTest = (nLow + nHigh) \ 2
50210                     If nTest = nLow Then
50220                         Exit Do
50230                     End If
50240                     nTestItem = .Item(nTest)
50251                     Select Case Item
                              Case Is < nTestItem
50270                             nHigh = nTest
50280                         Case nTestItem
50290                             Exit Do
50300                         Case Is > nTestItem
50310                             nLow = nTest
50320                     End Select
50330                 Loop
50340                 If nTest < nCount Then
50350                     Do While .Item(nTest + 1) = Item
50360                         nTest = nTest + 1
50370                         If nTest = nCount Then
50380                             Exit Do
50390                         End If
50400                     Loop
50410                 End If
50420                 .Add Item, key, , nTest
50430                 AddSortedDbl = nTest + 1
50440             End If
50450         Else
50460             .Add Item, key
50470             AddSortedDbl = 1
50480         End If
50490     End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modCollAddSorted", "AddSortedDbl")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function AddSortedCur(Collection As Collection, Item As Currency, Optional key As Variant) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim nCount As Long
50020     Dim nHigh As Long
50030     Dim nLow As Long
50040     Dim nTest As Long
50050     Dim nTestItem As Currency
50060
50070     With Collection
50080         nCount = .Count
50090         If nCount Then
50100             If Item < .Item(1) Then
50110                 .Add Item, key, 1
50120                 AddSortedCur = 1
50130             ElseIf Item > .Item(nCount) Then
50140                 .Add Item, key
50150                 AddSortedCur = nCount + 1
50160             Else
50170                 nLow = 1
50180                 nHigh = nCount
50190                 Do
50200                     nTest = (nLow + nHigh) \ 2
50210                     If nTest = nLow Then
50220                         Exit Do
50230                     End If
50240                     nTestItem = .Item(nTest)
50251                     Select Case Item
                              Case Is < nTestItem
50270                             nHigh = nTest
50280                         Case nTestItem
50290                             Exit Do
50300                         Case Is > nTestItem
50310                             nLow = nTest
50320                     End Select
50330                 Loop
50340                 If nTest < nCount Then
50350                     Do While .Item(nTest + 1) = Item
50360                         nTest = nTest + 1
50370                         If nTest = nCount Then
50380                             Exit Do
50390                         End If
50400                     Loop
50410                 End If
50420                 .Add Item, key, , nTest
50430                 AddSortedCur = nTest + 1
50440             End If
50450         Else
50460             .Add Item, key
50470             AddSortedCur = 1
50480         End If
50490     End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modCollAddSorted", "AddSortedCur")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function


