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
 Dim i As Long, c As Long
 c = 0
 
 With ListView
  If LockUpdate = True Then
   LockWindowUpdate .hWnd
  End If
  For i = 1 To .ListItems.Count
   If .ListItems(i).Selected = True Then
    c = c + 1
   End If
  Next i
 End With
 LvwGetCountSelectedItems = c
 If LockUpdate = True Then
  LockWindowUpdate 0&
 End If
End Function

Public Function LvwRemoveSelectedItems(ListView As ListView, Optional ByVal LockUpdate As Boolean) As Long
 Dim i As Long
 With ListView
  If LockUpdate = True Then
   LockWindowUpdate .hWnd
  End If
  For i = .ListItems.Count To 1 Step -1
   If .ListItems(i).Selected = True Then
    .ListItems.Remove i
   End If
  Next i
 End With
 If LockUpdate = True Then
  LockWindowUpdate 0&
 End If
End Function

Public Sub LvwListItemToTop(ListView As ListView, Optional KeyIndexListItem As Variant, Optional ByVal LockUpdate As Boolean)
 Dim nListItemStore As ListItemStore, nListItem As ListItem
    
 With ListView
  If LockUpdate Then
            LockWindowUpdate .hWnd
        End If
        With .ListItems
            Set nListItem = zGetListItem(ListView, KeyIndexListItem)
            If Not (nListItem Is Nothing) Then
                If nListItem.index > 1 Then
                    LvwGetListItemStore nListItemStore, ListView, nListItem, True
                    LvwInsertListItemStore ListView, nListItemStore, , 1
                End If
            End If
        End With
    End With
    If LockUpdate Then
        LockWindowUpdate 0&
    End If
End Sub

Public Sub LvwListItemDown(ListView As ListView, Optional KeyIndexListItem As Variant, Optional Steps As Long = 1, Optional ByVal LockUpdate As Boolean)
    Dim nListItemStore As ListItemStore
    Dim nListItem As ListItem
    Dim nIndex As Long
    
    With ListView
        If LockUpdate Then
            LockWindowUpdate .hWnd
        End If
        With .ListItems
            Set nListItem = zGetListItem(ListView, KeyIndexListItem)
            If Not (nListItem Is Nothing) Then
                nIndex = nListItem.index
                If nIndex > 1 Then
                    nIndex = nIndex - Steps
                    If nIndex < 1 Then
                        nIndex = 1
                    End If
                    LvwGetListItemStore nListItemStore, ListView, nListItem, True
                    LvwInsertListItemStore ListView, nListItemStore, , nIndex
                End If
            End If
        End With
    End With
    If LockUpdate Then
        LockWindowUpdate 0&
    End If
End Sub

Public Sub LvwListItemUp(ListView As ListView, Optional KeyIndexListItem As Variant, Optional Steps As Long = 1, Optional ByVal LockUpdate As Boolean)
 Dim nListItemStore As ListItemStore, nListItem As ListItem, nIndex As Long
    
 With ListView
  If LockUpdate Then
   LockWindowUpdate .hWnd
  End If
  With .ListItems
   Set nListItem = zGetListItem(ListView, KeyIndexListItem)
   If Not (nListItem Is Nothing) Then
    nIndex = nListItem.index
    If nIndex < .Count Then
     nIndex = nIndex + Steps
     LvwGetListItemStore nListItemStore, ListView, nListItem, True
     If nIndex > .Count Then
       LvwInsertListItemStore ListView, nListItemStore
      Else
       LvwInsertListItemStore ListView, nListItemStore, , nIndex
     End If
    End If
   End If
  End With
 End With
 If LockUpdate Then
  LockWindowUpdate 0&
 End If
End Sub

Public Sub LvwListItemToBottom(ListView As ListView, Optional KeyIndexListItem As Variant, Optional ByVal LockUpdate As Boolean)
 Dim nListItemStore As ListItemStore, nListItem As ListItem
    
 With ListView
  If LockUpdate Then
   LockWindowUpdate .hWnd
  End If
  With .ListItems
   Set nListItem = zGetListItem(ListView, KeyIndexListItem)
   If Not (nListItem Is Nothing) Then
    If nListItem.index < .Count Then
     LvwGetListItemStore nListItemStore, ListView, nListItem, True
     LvwInsertListItemStore ListView, nListItemStore
    End If
   End If
  End With
 End With
 If LockUpdate Then
  LockWindowUpdate 0&
 End If
End Sub

Public Sub LvwGetListItemStore(ListItemStore As ListItemStore, ListView As ListView, Optional KeyIndexListItem As Variant, Optional ByVal Remove As Boolean, Optional ByVal LockUpdate As Boolean)
    Dim nListItem As ListItem
    Dim nKeyIsListItem As Boolean
    Dim nListSubItem As ListSubItem
    Dim i As Integer
    
    With ListView
        If LockUpdate Then
            LockWindowUpdate .hWnd
        End If
        With .ListItems
            Set nListItem = zGetListItem(ListView, KeyIndexListItem)
            If Not (nListItem Is Nothing) Then
                With nListItem
                    ListItemStore.Bold = .Bold
                    ListItemStore.Checked = .Checked
                    ListItemStore.ForeColor = .ForeColor
                    ListItemStore.Ghosted = .Ghosted
                    ListItemStore.Icon = .Icon
                    ListItemStore.Key = .Key
                    ListItemStore.Selected = .Selected
                    ListItemStore.SmallIcon = .SmallIcon
                    If IsObject(.Tag) Then
                        Set ListItemStore.Tag = .Tag
                    Else
                        ListItemStore.Tag = .Tag
                    End If
                    ListItemStore.Text = .Text
                    ListItemStore.ToolTipText = .ToolTipText
                    If .ListSubItems.Count Then
                        ReDim ListItemStore.ListSubItems(1 To .ListSubItems.Count)
                        For i = 1 To .ListSubItems.Count
                            With .ListSubItems(i)
                                ListItemStore.ListSubItems(i).Bold = .Bold
                                ListItemStore.ListSubItems(i).ForeColor = .ForeColor
                                ListItemStore.ListSubItems(i).Key = .Key
                                ListItemStore.ListSubItems(i).ReportIcon = .ReportIcon
                                If IsObject(.Tag) Then
                                    Set ListItemStore.ListSubItems(i).Tag = .Tag
                                Else
                                    ListItemStore.ListSubItems(i).Tag = .Tag
                                End If
                                ListItemStore.ListSubItems(i).Text = .Text
                                ListItemStore.ListSubItems(i).ToolTipText = .ToolTipText
                            End With
                        Next 'i
                    Else
                        ReDim ListItemStore.ListSubItems(0)
                    End If
                End With
                If Remove Then
                    .Remove nListItem.index
                End If
            End If
        End With
    End With
    If LockUpdate Then
        LockWindowUpdate 0&
    End If
End Sub

Private Function zGetListItem(ListView As ListView, Optional KeyIndexListItem As Variant) As ListItem
 With ListView
  If IsMissing(KeyIndexListItem) Then
    Set zGetListItem = .SelectedItem
   ElseIf IsObject(KeyIndexListItem) Then
     If TypeOf KeyIndexListItem Is ListItem Then
      Set zGetListItem = KeyIndexListItem
     End If
    Else
     On Error Resume Next
     Set zGetListItem = .ListItems.item(KeyIndexListItem)
     On Error GoTo 0
  End If
 End With
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
                    nIndex = .SelectedItem.index
                End If
            End If
        ElseIf Not IsMissing(AfterSelectedItem) Then
            If Not (.SelectedItem Is Nothing) Then
                If CBool(AfterSelectedItem) Then
                    nIndex = .SelectedItem.index + 1
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

