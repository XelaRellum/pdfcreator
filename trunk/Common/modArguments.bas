Attribute VB_Name = "modArguments"
Option Explicit

'This code is from www.aboutvb.de

Private mArguments As Collection
Private pCommandLine As String

Public Property Get CommandLine() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  CommandLine = pCommandLine
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modArguments", "CommandLine [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let CommandLine(New_CommandLine As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim nCommandLine As String
50020  nCommandLine = Trim$(New_CommandLine)
50030  If pCommandLine <> nCommandLine Then
50040   pCommandLine = nCommandLine
50050   zGetArguments
50060  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modArguments", "CommandLine [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get CommandArgumentsCount() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  zInitArguments
50020  CommandArgumentsCount = mArguments.Count
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modArguments", "CommandArgumentsCount [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get CommandArgument(ByVal Index As Long, Optional ByVal ReducedQuotes As Boolean) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  zInitArguments
50020  If ReducedQuotes Then
50030    CommandArgument = ReduceQuotes(mArguments(Index))
50040   Else
50050    CommandArgument = mArguments(Index)
50060  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modArguments", "CommandArgument [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get CommandSwitch(Switch As String, Optional ByVal ReducedQuotes As Boolean) As Variant
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Integer, nArgument As String, nCommandSwitch As String
50020
50030  zInitArguments
50040  For i = 1 To mArguments.Count
50050   nArgument = mArguments(i)
50061   Select Case Left$(nArgument, 1)
         Case "-", "/"
50080     If Mid$(UCase$(nArgument), 2, Len(Switch)) = UCase$(Switch) Then
50090      If ReducedQuotes Then
50100        nCommandSwitch = ReduceQuotes(Mid$(nArgument, Len(Switch) + 2))
50110       Else
50120        nCommandSwitch = Mid$(nArgument, Len(Switch) + 2)
50130      End If
50140      If Left$(nCommandSwitch, 1) = "=" Then
50150        CommandSwitch = Trim$(Mid$(nCommandSwitch, 2))
50160       Else
50170        CommandSwitch = Trim$(nCommandSwitch)
50180      End If
50190      Exit Property
50200     End If
50210   End Select
50220  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modArguments", "CommandSwitch [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Function ReduceQuotes(Arg As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim nArg As String
50020  ReduceQuotes = Arg
50030  nArg = Arg
50040  If Left$(nArg, 1) = Chr$(34) Then
50050   If Right$(nArg, 1) = Chr$(34) Then
50060    nArg = Replace$(Arg, Chr$(34) & Chr$(34), Chr$(34))
50070    ReduceQuotes = Mid$(nArg, 2, Len(nArg) - 2)
50080   End If
50090  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modArguments", "ReduceQuotes")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub zGetArguments()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim nCommandLine As String, nParts() As String, i As Integer
50020  If Len(pCommandLine) = 0 Then
50030   pCommandLine = Trim$(VBA.Command$)
50040  End If
50050  If Len(pCommandLine) = 0 Then
50060   Set mArguments = New Collection
50070   Exit Sub
50080  End If
50090  nCommandLine = " " & Replace(pCommandLine, Chr$(34) & Chr$(34), Chr$(1)) & " "
50100  nParts = Split(nCommandLine, Chr$(34))
50110  For i = 0 To UBound(nParts)
50120   If i And 1 Then
50130    nParts(i) = Replace$(nParts(i), " ", Chr$(2))
50140    nParts(i) = Replace$(nParts(i), "/", Chr$(3))
50150    nParts(i) = Replace$(nParts(i), "-", Chr$(4))
50160    nParts(i) = Chr$(34) & nParts(i) & Chr$(34)
50170   End If
50180  Next i
50190  nCommandLine = Trim$(Join(nParts, ""))
50200  nCommandLine = Replace$(nCommandLine, "/", " /")
50210  nCommandLine = Replace$(nCommandLine, "-", " -")
50220  nParts = Split(nCommandLine, " ")
50230  Set mArguments = New Collection
50240  For i = 0 To UBound(nParts)
50250   If Len(nParts(i)) Then
50260    nParts(i) = Replace$(nParts(i), Chr$(1), Chr$(34) & Chr$(34))
50270    nParts(i) = Replace$(nParts(i), Chr$(2), " ")
50280    nParts(i) = Replace$(nParts(i), Chr$(3), "/")
50290    nParts(i) = Replace$(nParts(i), Chr$(4), "-")
50300    mArguments.Add nParts(i), nParts(i)
50310   End If
50320  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modArguments", "zGetArguments")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub zInitArguments()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If mArguments Is Nothing Then
50020   zGetArguments
50030  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modArguments", "zInitArguments")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
