Attribute VB_Name = "modGsApi"
' Copyright (c) 2002 Dan Mount and Ghostgum Software Pty Ltd
'
' Permission is hereby granted, free of charge, to any person obtaining
' a copy of this software and associated documentation files (the
' "Software"), to deal in the Software without restriction, including
' without limitation the rights to use, copy, modify, merge, publish,
' distribute, sublicense, and/or sell copies of the Software, and to
' permit persons to whom the Software is furnished to do so, subject to
' the following conditions:
'
' The above copyright notice and this permission notice shall be
' included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
' EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
' MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
' NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS
' BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
' ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
' CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.


' This is an example of how to call the Ghostscript DLL from
' Visual Basic 6.  This example converts colorcir.ps to PDF.
' The display device is not supported.
'
' This code is not compatible with VB.NET.  There is another
' example which does work with VB.NET.  Differences include:
' 1. VB.NET uses GCHandle to get pointer
'    VB6 uses StrPtr/VarPtr
' 2. VB.NET Integer is 32bits, Long is 64bits
'    VB6 Integer is 16bits, Long is 32bits
' 3. VB.NET uses IntPtr for pointers
'    VB6 uses Long for pointers
' 4. VB.NET strings are always Unicode
'    VB6 can create an ANSI string
' See the following URL for some VB6 / VB.NET details
'  http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dnvb600/html/vb6tovbdotnet.asp

Option Explicit

'------------------------------------------------
'API Calls Start
'------------------------------------------------
'Win32 API
'GhostScript API
Public Const GsDll = "gsdll32.dll"

Private Declare Function gsapi_revision Lib "gsdll32.dll" (ByVal pGSRevisionInfo As Long, ByVal intLen As Long) As Long
Private Declare Function gsapi_new_instance Lib "gsdll32.dll" (ByRef lngGSInstance As Long, ByVal lngCallerHandle As Long) As Long
Private Declare Function gsapi_set_stdio Lib "gsdll32.dll" (ByVal lngGSInstance As Long, ByVal gsdll_stdin As Long, ByVal gsdll_stdout As Long, ByVal gsdll_stderr As Long) As Long
Private Declare Sub gsapi_delete_instance Lib "gsdll32.dll" (ByVal lngGSInstance As Long)
Private Declare Function gsapi_init_with_args Lib "gsdll32.dll" (ByVal lngGSInstance As Long, ByVal lngArgumentCount As Long, ByVal lngArguments As Long) As Long
Private Declare Function gsapi_run_file Lib "gsdll32.dll" (ByVal lngGSInstance As Long, ByVal strFileName As String, ByVal intErrors As Long, ByVal intExitCode As Long) As Long
Private Declare Function gsapi_exit Lib "gsdll32.dll" (ByVal lngGSInstance As Long) As Long
'------------------------------------------------
'API Calls End
'------------------------------------------------


'------------------------------------------------
'UDTs Start
'------------------------------------------------
Private Type GS_Revision
    strProduct As Long
    strCopyright As Long
    intRevision As Long
    intRevisionDate As Long
End Type

Public Type tGhostscriptRevision
 strProduct As String
 strCopyright As String
 intRevision As Long
 intRevisionDate As Long
End Type
'------------------------------------------------
'UDTs End
'------------------------------------------------

Public GSRevision As tGhostscriptRevision

'------------------------------------------------
'Callback Functions Start
'------------------------------------------------
'These are only required if you use gsapi_set_stdio

Public Function gsdll_stdin(ByVal intGSInstanceHandle As Long, ByVal strz As Long, ByVal intBytes As Long) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     ' We don't have a console, so just return EOF
50020     gsdll_stdin = 0
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGsApi", "gsdll_stdin")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function gsdll_stdout(ByVal intGSInstanceHandle As Long, ByVal strz As Long, ByVal intBytes As Long) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     ' If you can think of a more efficient method, please tell me!
50020     ' We need to convert from a byte buffer to a string
50030     ' First we create a byte array of the appropriate size
50040       Dim aByte() As Byte
50050       ReDim aByte(intBytes)
50060       ' Then we get the address of the byte array
50070       Dim ptrByte As Long
50080       ptrByte = VarPtr(aByte(0))
50090       ' Then we copy the buffer to the byte array
50100       MoveMemoryLong ptrByte, strz, intBytes
50110       ' Then we copy the byte array to a string, character by character
50120       Dim tStr As String
50130       Dim i As Long
50140       For i = 0 To intBytes - 1
50150           tStr = tStr + Chr(aByte(i))
50160       Next
50170       ' Finally we output the message
50180       ReturnValue tStr
50190       gsdll_stdout = intBytes
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGsApi", "gsdll_stdout")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function gsdll_stderr(ByVal intGSInstanceHandle As Long, ByVal strz As Long, ByVal intBytes As Long) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     gsdll_stderr = gsdll_stdout(intGSInstanceHandle, strz, intBytes)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGsApi", "gsdll_stderr")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
'------------------------------------------------
'Callback Functions End
'------------------------------------------------


'------------------------------------------------
'User Defined Functions Start
'------------------------------------------------
Public Function AnsiZtoString(ByVal strz As Long) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Rem We need to convert from a byte buffer to a string
50020     Dim byteCh(1) As Byte
50030     Dim bOK As Boolean
50040     bOK = True
50050     Dim ptrByte As Long
50060     ptrByte = VarPtr(byteCh(0))
50070     Dim j As Long
50080     j = 0
50090     Dim str As String
50100     While bOK
50110         ' This is how to do pointer arithmetic!
50120         MoveMemoryLong ptrByte, strz + j, 1
50130         If byteCh(0) = 0 Then
50140             bOK = False
50150         Else
50160             str = str + Chr(byteCh(0))
50170         End If
50180         j = j + 1
50190     Wend
50200     AnsiZtoString = str
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGsApi", "AnsiZtoString")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function CheckRevision(ByVal intRevision As Long) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     ' Check revision number of Ghostscript
50020     Dim intReturn As Long
50030     Dim udtGSRevInfo As GS_Revision
50040     intReturn = gsapi_revision(VarPtr(udtGSRevInfo), 16)
50050     Dim str As String
50060     str = "Revision=" & udtGSRevInfo.intRevision
50070     str = str & "  RevisionDate=" & udtGSRevInfo.intRevisionDate
50080     str = str & "  Product=" & AnsiZtoString(udtGSRevInfo.strProduct)
50090     str = str & "  Copyright = " & AnsiZtoString(udtGSRevInfo.strCopyright)
50100     ReturnValue str
50110     'MsgBox (str)
50120
50130     If udtGSRevInfo.intRevision = intRevision Then
50140         CheckRevision = True
50150     Else
50160         CheckRevision = False
50170     End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGsApi", "CheckRevision")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function CallGS(ByRef astrGSArgs() As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim intReturn As Long
50020     Dim intGSInstanceHandle As Long
50030     Dim aAnsiArgs() As String
50040     Dim aPtrArgs() As Long
50050     Dim intCounter As Long
50060     Dim intElementCount As Long
50070     Dim iTemp As Long
50080     Dim callerHandle As Long
50090     Dim ptrArgs As Long
50100
50110     ' Print out the revision details.
50120     ' If we want to insist on a particular version of Ghostscript
50130     ' we should check the return value of CheckRevision().
50140     'CheckRevision (705)
50150
50160     ' Load Ghostscript and get the instance handle
50170     intReturn = gsapi_new_instance(intGSInstanceHandle, callerHandle)
50180     If (intReturn < 0) Then
50190         CallGS = False
50200         Return
50210     End If
50220
50230     ' Capture stdio
50240     intReturn = gsapi_set_stdio(intGSInstanceHandle, AddressOf gsdll_stdin, AddressOf gsdll_stdout, AddressOf gsdll_stderr)
50250
50260     If (intReturn >= 0) Then
50270         ' Convert the Unicode strings to null terminated ANSI byte arrays
50280         ' then get pointers to the byte arrays.
50290         intElementCount = UBound(astrGSArgs)
50300         ReDim aAnsiArgs(intElementCount)
50310         ReDim aPtrArgs(intElementCount)
50320
50330         For intCounter = 0 To intElementCount
50340             aAnsiArgs(intCounter) = StrConv(astrGSArgs(intCounter), vbFromUnicode)
50350             aPtrArgs(intCounter) = StrPtr(aAnsiArgs(intCounter))
50360         Next
50370         ptrArgs = VarPtr(aPtrArgs(0))
50380
50390         intReturn = gsapi_init_with_args(intGSInstanceHandle, intElementCount + 1, ptrArgs)
50400
50410         ' Stop the Ghostscript interpreter
50420         gsapi_exit (intGSInstanceHandle)
50430     End If
50440
50450     ' release the Ghostscript instance handle
50460     gsapi_delete_instance (intGSInstanceHandle)
50470
50480     If (intReturn >= 0) Then
50490         CallGS = True
50500     Else
50510         CallGS = False
50520     End If
50530
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGsApi", "CallGS")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetGhostscriptRevision() As tGhostscriptRevision
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim intReturn As Long, udtGSRevInfo As GS_Revision
50020  intReturn = gsapi_revision(VarPtr(udtGSRevInfo), 16)
50030  With GetGhostscriptRevision
50040   .intRevision = udtGSRevInfo.intRevision
50050   .intRevisionDate = udtGSRevInfo.intRevisionDate
50060   .strCopyright = AnsiZtoString(udtGSRevInfo.strCopyright)
50070   .strProduct = AnsiZtoString(udtGSRevInfo.strProduct)
50080  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGsApi", "GetGhostscriptRevision")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

