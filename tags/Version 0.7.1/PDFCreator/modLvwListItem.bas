Attribute VB_Name = "modLvwListItem"
Option Explicit

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Type ListSubItemStore
    Bold As String
    ForeColor As Long
    Key As String
    ReportIcon As Variant
    Tag As Variant
    Text As String
    ToolTipText As String
End Type

Public Type ListItemStore
    Bold As Boolean
    Checked As Boolean
    ForeColor As Long
    Ghosted As Boolean
    Icon As Variant
    Key As String
    ListSubItems() As ListSubItemStore
    Selected As Boolean
    SmallIcon As Variant
    Tag As Variant
    Text As String
    ToolTipText As String
End Type

Public Function LvwGetCountSelectedItems(ListView As ListView, Optional ByVal LockUpdate As Boolean) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, c As Long
50020  c = 0
50030
50040  With ListView
50050   If LockUpdate = True Then
50060    LockWindowUpdate .hWnd
50070   End If
50080   For i = 1 To .ListItems.Count
50090    If .ListItems(i).Selected = True Then
50100     c = c + 1
50110    End If
50120   Next i
50130  End With
50140  LvwGetCountSelectedItems = c
50150  If LockUpdate = True Then
50160   LockWindowUpdate 0&
50170  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modLvwListItem", "LvwGetCountSelectedItems")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function LvwRemoveSelectedItems(ListView As ListView, Optional ByVal LockUpdate As Boolean) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  With ListView
50030   If LockUpdate = True Then
50040    LockWindowUpdate .hWnd
50050   End If
50060   For i = .ListItems.Count To 1 Step -1
50070    If .ListItems(i).Selected = True Then
50080     .ListItems.Remove i
50090    End If
50100   Next i
50110  End With
50120  If LockUpdate = True Then
50130   LockWindowUpdate 0&
50140  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modLvwListItem", "LvwRemoveSelectedItems")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub LvwListItemToTop(ListView As ListView, Optional KeyIndexListItem As Variant, Optional ByVal LockUpdate As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim nListItemStore As ListItemStore, nListItem As ListItem
50020
50030  With ListView
50040   If LockUpdate Then
50050             LockWindowUpdate .hWnd
50060         End If
50070         With .ListItems
50080             Set nListItem = zGetListItem(ListView, KeyIndexListItem)
50090             If Not (nListItem Is Nothing) Then
50100                 If nListItem.Index > 1 Then
50110                     LvwGetListItemStore nListItemStore, ListView, nListItem, True
50120                     LvwInsertListItemStore ListView, nListItemStore, , 1
50130                 End If
50140             End If
50150         End With
50160     End With
50170     If LockUpdate Then
50180         LockWindowUpdate 0&
50190     End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modLvwListItem", "LvwListItemToTop")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub LvwListItemDown(ListView As ListView, Optional KeyIndexListItem As Variant, Optional Steps As Long = 1, Optional ByVal LockUpdate As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim nListItemStore As ListItemStore
50020     Dim nListItem As ListItem
50030     Dim nIndex As Long
50040
50050     With ListView
50060         If LockUpdate Then
50070             LockWindowUpdate .hWnd
50080         End If
50090         With .ListItems
50100             Set nListItem = zGetListItem(ListView, KeyIndexListItem)
50110             If Not (nListItem Is Nothing) Then
50120                 nIndex = nListItem.Index
50130                 If nIndex > 1 Then
50140                     nIndex = nIndex - Steps
50150                     If nIndex < 1 Then
50160                         nIndex = 1
50170                     End If
50180                     LvwGetListItemStore nListItemStore, ListView, nListItem, True
50190                     LvwInsertListItemStore ListView, nListItemStore, , nIndex
50200                 End If
50210             End If
50220         End With
50230     End With
50240     If LockUpdate Then
50250         LockWindowUpdate 0&
50260     End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modLvwListItem", "LvwListItemDown")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub LvwListItemUp(ListView As ListView, Optional KeyIndexListItem As Variant, Optional Steps As Long = 1, Optional ByVal LockUpdate As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim nListItemStore As ListItemStore, nListItem As ListItem, nIndex As Long
50020
50030  With ListView
50040   If LockUpdate Then
50050    LockWindowUpdate .hWnd
50060   End If
50070   With .ListItems
50080    Set nListItem = zGetListItem(ListView, KeyIndexListItem)
50090    If Not (nListItem Is Nothing) Then
50100     nIndex = nListItem.Index
50110     If nIndex < .Count Then
50120      nIndex = nIndex + Steps
50130      LvwGetListItemStore nListItemStore, ListView, nListItem, True
50140      If nIndex > .Count Then
50150        LvwInsertListItemStore ListView, nListItemStore
50160       Else
50170        LvwInsertListItemStore ListView, nListItemStore, , nIndex
50180      End If
50190     End If
50200    End If
50210   End With
50220  End With
50230  If LockUpdate Then
50240   LockWindowUpdate 0&
50250  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modLvwListItem", "LvwListItemUp")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub LvwListItemToBottom(ListView As ListView, Optional KeyIndexListItem As Variant, Optional ByVal LockUpdate As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim nListItemStore As ListItemStore, nListItem As ListItem
50020
50030  With ListView
50040   If LockUpdate Then
50050    LockWindowUpdate .hWnd
50060   End If
50070   With .ListItems
50080    Set nListItem = zGetListItem(ListView, KeyIndexListItem)
50090    If Not (nListItem Is Nothing) Then
50100     If nListItem.Index < .Count Then
50110      LvwGetListItemStore nListItemStore, ListView, nListItem, True
50120      LvwInsertListItemStore ListView, nListItemStore
50130     End If
50140    End If
50150   End With
50160  End With
50170  If LockUpdate Then
50180   LockWindowUpdate 0&
50190  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modLvwListItem", "LvwListItemToBottom")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub LvwGetListItemStore(ListItemStore As ListItemStore, ListView As ListView, Optional KeyIndexListItem As Variant, Optional ByVal Remove As Boolean, Optional ByVal LockUpdate As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim nListItem As ListItem
50020     Dim nKeyIsListItem As Boolean
50030     Dim nListSubItem As ListSubItem
50040     Dim i As Integer
50050
50060     With ListView
50070         If LockUpdate Then
50080             LockWindowUpdate .hWnd
50090         End If
50100         With .ListItems
50110             Set nListItem = zGetListItem(ListView, KeyIndexListItem)
50120             If Not (nListItem Is Nothing) Then
50130                 With nListItem
50140                     ListItemStore.Bold = .Bold
50150                     ListItemStore.Checked = .Checked
50160                     ListItemStore.ForeColor = .ForeColor
50170                     ListItemStore.Ghosted = .Ghosted
50180                     ListItemStore.Icon = .Icon
50190                     ListItemStore.Key = .Key
50200                     ListItemStore.Selected = .Selected
50210                     ListItemStore.SmallIcon = .SmallIcon
50220                     If IsObject(.Tag) Then
50230                         Set ListItemStore.Tag = .Tag
50240                     Else
50250                         ListItemStore.Tag = .Tag
50260                     End If
50270                     ListItemStore.Text = .Text
50280                     ListItemStore.ToolTipText = .ToolTipText
50290                     If .ListSubItems.Count Then
50300                         ReDim ListItemStore.ListSubItems(1 To .ListSubItems.Count)
50310                         For i = 1 To .ListSubItems.Count
50320                             With .ListSubItems(i)
50330                                 ListItemStore.ListSubItems(i).Bold = .Bold
50340                                 ListItemStore.ListSubItems(i).ForeColor = .ForeColor
50350                                 ListItemStore.ListSubItems(i).Key = .Key
50360                                 ListItemStore.ListSubItems(i).ReportIcon = .ReportIcon
50370                                 If IsObject(.Tag) Then
50380                                     Set ListItemStore.ListSubItems(i).Tag = .Tag
50390                                 Else
50400                                     ListItemStore.ListSubItems(i).Tag = .Tag
50410                                 End If
50420                                 ListItemStore.ListSubItems(i).Text = .Text
50430                                 ListItemStore.ListSubItems(i).ToolTipText = .ToolTipText
50440                             End With
50450                         Next 'i
50460                     Else
50470                         ReDim ListItemStore.ListSubItems(0)
50480                     End If
50490                 End With
50500                 If Remove Then
50510                     .Remove nListItem.Index
50520                 End If
50530             End If
50540         End With
50550     End With
50560     If LockUpdate Then
50570         LockWindowUpdate 0&
50580     End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modLvwListItem", "LvwGetListItemStore")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function zGetListItem(ListView As ListView, Optional KeyIndexListItem As Variant) As ListItem
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With ListView
50020   If IsMissing(KeyIndexListItem) Then
50030     Set zGetListItem = .SelectedItem
50040    ElseIf IsObject(KeyIndexListItem) Then
50050      If TypeOf KeyIndexListItem Is ListItem Then
50060       Set zGetListItem = KeyIndexListItem
50070      End If
50080     Else
50090      Set zGetListItem = .ListItems.item(KeyIndexListItem)
50100   End If
50110  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modLvwListItem", "zGetListItem")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub LvwInsertListItemStore(ListView As ListView, ListItemStore As ListItemStore, Optional NewKey As String, Optional ByVal Before As Variant, Optional ByVal After As Variant, Optional ByVal BeforeSelectedItem As Variant, Optional ByVal AfterSelectedItem As Variant, Optional ByVal IgnoreSelection As Boolean)
    Dim nIndex As Variant
    Dim nListItem As ListItem
    Dim nListSubItem As ListSubItem
    Dim i As Integer

    With ListView
        If Not IsMissing(BeforeSelectedItem) Then
            If Not (.SelectedItem Is Nothing) Then
                If CBool(BeforeSelectedItem) Then
                    nIndex = .SelectedItem.Index
                End If
            End If
        ElseIf Not IsMissing(AfterSelectedItem) Then
            If Not (.SelectedItem Is Nothing) Then
                If CBool(AfterSelectedItem) Then
                    nIndex = .SelectedItem.Index + 1
                End If
            End If
        ElseIf Not IsMissing(Before) Then
            nIndex = Before
        ElseIf Not IsMissing(After) Then
            nIndex = After + 1
        End If
    End With
    With ListItemStore
        On Error Resume Next
        If IsEmpty(nIndex) Then
            Set nListItem = ListView.ListItems.Add(, .Key, .Text, .Icon, .SmallIcon)
        ElseIf nIndex = 0 Then
            Exit Sub
        Else
            If nIndex > ListView.ListItems.Count Then
                Set nListItem = ListView.ListItems.Add(, .Key, .Text, .Icon, .SmallIcon)
            Else
                Set nListItem = ListView.ListItems.Add(nIndex, .Key, .Text, .Icon, .SmallIcon)
            End If
        End If
        nListItem.Bold = .Bold
        nListItem.Checked = .Checked
        nListItem.ForeColor = .ForeColor
        nListItem.Ghosted = .Ghosted
        If Len(NewKey) Then
            nListItem.Key = NewKey
        Else
            nListItem.Key = .Key
        End If
        If IsObject(.Tag) Then
            Set nListItem.Tag = .Tag
        Else
            nListItem.Tag = .Tag
        End If
        nListItem.ToolTipText = .ToolTipText
        If UBound(.ListSubItems) Then
            For i = 1 To UBound(.ListSubItems)
                Set nListSubItem = nListItem.ListSubItems.Add(, .ListSubItems(i).Key, .ListSubItems(i).Text, .ListSubItems(i).ReportIcon, .ListSubItems(i).ToolTipText)
                nListSubItem.Bold = .ListSubItems(i).Bold
                nListSubItem.ForeColor = .ListSubItems(i).ForeColor
                If IsObject(.ListSubItems(i).Tag) Then
                    Set nListSubItem.Tag = .ListSubItems(i).Tag
                Else
                    nListSubItem.Tag = .ListSubItems(i).Tag
                End If
            Next 'i
        End If
        If Not IgnoreSelection Then
            nListItem.Selected = .Selected
        End If
    End With
End Sub

