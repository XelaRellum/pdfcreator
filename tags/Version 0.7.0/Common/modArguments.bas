Attribute VB_Name = "modArguments"
Option Explicit

'This code is from www.aboutvb.de

Private mArguments As Collection
Private pCommandLine As String

Public Property Get CommandLine() As String
 CommandLine = pCommandLine
End Property

Public Property Let CommandLine(New_CommandLine As String)
 Dim nCommandLine As String
 nCommandLine = Trim$(New_CommandLine)
 If pCommandLine <> nCommandLine Then
  pCommandLine = nCommandLine
  zGetArguments
 End If
End Property

Public Property Get CommandArgumentsCount() As Long
 zInitArguments
 CommandArgumentsCount = mArguments.Count
End Property

Public Property Get CommandArgument(ByVal Index As Long, Optional ByVal ReducedQuotes As Boolean) As String
 zInitArguments
 If ReducedQuotes Then
   CommandArgument = ReduceQuotes(mArguments(Index))
  Else
   CommandArgument = mArguments(Index)
 End If
End Property

Public Property Get CommandSwitch(Switch As String, Optional ByVal ReducedQuotes As Boolean) As Variant
 Dim i As Integer, nArgument As String, nCommandSwitch As String

 zInitArguments
 For i = 1 To mArguments.Count
  nArgument = mArguments(i)
  Select Case Left$(nArgument, 1)
   Case "-", "/"
    If Mid$(UCase$(nArgument), 2, Len(Switch)) = UCase$(Switch) Then
     If ReducedQuotes Then
       nCommandSwitch = ReduceQuotes(Mid$(nArgument, Len(Switch) + 2))
      Else
       nCommandSwitch = Mid$(nArgument, Len(Switch) + 2)
     End If
     If Left$(nCommandSwitch, 1) = "=" Then
       CommandSwitch = Trim$(Mid$(nCommandSwitch, 2))
      Else
       CommandSwitch = Trim$(nCommandSwitch)
     End If
     Exit Property
    End If
  End Select
 Next i
End Property

Public Function ReduceQuotes(Arg As String) As String
 Dim nArg As String
 ReduceQuotes = Arg
 nArg = Arg
 If Left$(nArg, 1) = Chr$(34) Then
  If Right$(nArg, 1) = Chr$(34) Then
   nArg = Replace$(Arg, Chr$(34) & Chr$(34), Chr$(34))
   ReduceQuotes = Mid$(nArg, 2, Len(nArg) - 2)
  End If
 End If
End Function

Private Sub zGetArguments()
 Dim nCommandLine As String, nParts() As String, i As Integer
 If Len(pCommandLine) = 0 Then
  pCommandLine = Trim$(VBA.Command$)
 End If
 If Len(pCommandLine) = 0 Then
  Set mArguments = New Collection
  Exit Sub
 End If
 nCommandLine = " " & Replace(pCommandLine, Chr$(34) & Chr$(34), Chr$(1)) & " "
 nParts = Split(nCommandLine, Chr$(34))
 For i = 0 To UBound(nParts)
  If i And 1 Then
   nParts(i) = Replace$(nParts(i), " ", Chr$(2))
   nParts(i) = Replace$(nParts(i), "/", Chr$(3))
   nParts(i) = Replace$(nParts(i), "-", Chr$(4))
   nParts(i) = Chr$(34) & nParts(i) & Chr$(34)
  End If
 Next i
 nCommandLine = Trim$(Join(nParts, ""))
 nCommandLine = Replace$(nCommandLine, "/", " /")
 nCommandLine = Replace$(nCommandLine, "-", " -")
 nParts = Split(nCommandLine, " ")
 Set mArguments = New Collection
 For i = 0 To UBound(nParts)
  If Len(nParts(i)) Then
   nParts(i) = Replace$(nParts(i), Chr$(1), Chr$(34) & Chr$(34))
   nParts(i) = Replace$(nParts(i), Chr$(2), " ")
   nParts(i) = Replace$(nParts(i), Chr$(3), "/")
   nParts(i) = Replace$(nParts(i), Chr$(4), "-")
   mArguments.Add nParts(i), nParts(i)
  End If
 Next i
End Sub

Private Sub zInitArguments()
 If mArguments Is Nothing Then
  zGetArguments
 End If
End Sub
