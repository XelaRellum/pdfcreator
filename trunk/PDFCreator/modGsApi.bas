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
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As Long, ByVal Source As Long, ByVal bytes As Long)

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
'------------------------------------------------
'UDTs End
'------------------------------------------------



'------------------------------------------------
'Callback Functions Start
'------------------------------------------------
'These are only required if you use gsapi_set_stdio

Public Function gsdll_stdin(ByVal intGSInstanceHandle As Long, ByVal strz As Long, ByVal intBytes As Long) As Long
    ' We don't have a console, so just return EOF
    gsdll_stdin = 0
End Function

Public Function gsdll_stdout(ByVal intGSInstanceHandle As Long, ByVal strz As Long, ByVal intBytes As Long) As Long
    ' If you can think of a more efficient method, please tell me!
    ' We need to convert from a byte buffer to a string
    ' First we create a byte array of the appropriate size

'    If frmMain.chkDebug.Value = True Then
      Dim aByte() As Byte
      ReDim aByte(intBytes)
      ' Then we get the address of the byte array
      Dim ptrByte As Long
      ptrByte = VarPtr(aByte(0))
      ' Then we copy the buffer to the byte array
      CopyMemory ptrByte, strz, intBytes
      ' Then we copy the byte array to a string, character by character
      Dim str As String
      Dim i As Long
      For i = 0 To intBytes - 1
          str = str + Chr(aByte(i))
      Next
      ' Finally we output the message
      'Debug.Print (str)
      ReturnValue str
      gsdll_stdout = intBytes
'    End If
End Function

Public Function gsdll_stderr(ByVal intGSInstanceHandle As Long, ByVal strz As Long, ByVal intBytes As Long) As Long
    gsdll_stderr = gsdll_stdout(intGSInstanceHandle, strz, intBytes)
End Function
'------------------------------------------------
'Callback Functions End
'------------------------------------------------


'------------------------------------------------
'User Defined Functions Start
'------------------------------------------------
Public Function AnsiZtoString(ByVal strz As Long) As String
    Rem We need to convert from a byte buffer to a string
    Dim byteCh(1) As Byte
    Dim bOK As Boolean
    bOK = True
    Dim ptrByte As Long
    ptrByte = VarPtr(byteCh(0))
    Dim j As Long
    j = 0
    Dim str As String
    While bOK
        ' This is how to do pointer arithmetic!
        CopyMemory ptrByte, strz + j, 1
        If byteCh(0) = 0 Then
            bOK = False
        Else
            str = str + Chr(byteCh(0))
        End If
        j = j + 1
    Wend
    AnsiZtoString = str
End Function

Public Function CheckRevision(ByVal intRevision As Long) As Boolean
    ' Check revision number of Ghostscript
    Dim intReturn As Long
    Dim udtGSRevInfo As GS_Revision
    intReturn = gsapi_revision(VarPtr(udtGSRevInfo), 16)
    Dim str As String
    str = "Revision=" & udtGSRevInfo.intRevision
    str = str & "  RevisionDate=" & udtGSRevInfo.intRevisionDate
    str = str & "  Product=" & AnsiZtoString(udtGSRevInfo.strProduct)
    str = str & "  Copyright = " & AnsiZtoString(udtGSRevInfo.strCopyright)
    ReturnValue str
    'MsgBox (str)

    If udtGSRevInfo.intRevision = intRevision Then
        CheckRevision = True
    Else
        CheckRevision = False
    End If
End Function

Public Function CallGS(ByRef astrGSArgs() As String) As Boolean
    Dim intReturn As Long
    Dim intGSInstanceHandle As Long
    Dim aAnsiArgs() As String
    Dim aPtrArgs() As Long
    Dim intCounter As Long
    Dim intElementCount As Long
    Dim iTemp As Long
    Dim callerHandle As Long
    Dim ptrArgs As Long

    ' Print out the revision details.
    ' If we want to insist on a particular version of Ghostscript
    ' we should check the return value of CheckRevision().
    'CheckRevision (705)

    ' Load Ghostscript and get the instance handle
    intReturn = gsapi_new_instance(intGSInstanceHandle, callerHandle)
    If (intReturn < 0) Then
        CallGS = False
        Return
    End If

    ' Capture stdio
    intReturn = gsapi_set_stdio(intGSInstanceHandle, AddressOf gsdll_stdin, AddressOf gsdll_stdout, AddressOf gsdll_stderr)

    If (intReturn >= 0) Then
        ' Convert the Unicode strings to null terminated ANSI byte arrays
        ' then get pointers to the byte arrays.
        intElementCount = UBound(astrGSArgs)
        ReDim aAnsiArgs(intElementCount)
        ReDim aPtrArgs(intElementCount)

        For intCounter = 0 To intElementCount
            aAnsiArgs(intCounter) = StrConv(astrGSArgs(intCounter), vbFromUnicode)
            aPtrArgs(intCounter) = StrPtr(aAnsiArgs(intCounter))
        Next
        ptrArgs = VarPtr(aPtrArgs(0))

        intReturn = gsapi_init_with_args(intGSInstanceHandle, intElementCount + 1, ptrArgs)

        ' Stop the Ghostscript interpreter
        gsapi_exit (intGSInstanceHandle)
    End If

    ' release the Ghostscript instance handle
    gsapi_delete_instance (intGSInstanceHandle)

    If (intReturn >= 0) Then
        CallGS = True
    Else
        CallGS = False
    End If

End Function
