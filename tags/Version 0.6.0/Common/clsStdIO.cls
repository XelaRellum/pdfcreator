VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStdIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'#####################################################################################
'#  Standard IO (StdIn/StdOut/StdErr) Wrapper Class (clsStdIO.cls)
'#      By: Nick Campbeln
'#      Aug 4, 2002
'#
'#      Revision History:
'#          1.0.2 (Aug 26, 2002):
'#              Fixed a (very) stupid coding error in StdErr() - Never renamed copied sub name to "StdErr" (guess I made the addition too quickly =)
'#          1.0.1 (Aug 4, 2002):
'#              (Quickly) Added ability to interface with StdErr
'#          1.0 (Aug 4, 2002):
'#              Initial Release
'#
'#      Copyright © 2002 Nick Campbeln (opensource@nick.campbeln.com)
'#          This source code is provided 'as-is', without any express or implied warranty. In no event will the author(s) be held liable for any damages arising from the use of this source code. Permission is granted to anyone to use this source code for any purpose, including commercial applications, and to alter it and redistribute it freely, subject to the following restrictions:
'#          1. The origin of this source code must not be misrepresented; you must not claim that you wrote the original source code. If you use this source code in a product, an acknowledgment in the product documentation would be appreciated but is not required.
'#          2. Altered source versions must be plainly marked as such, and must not be misrepresented as being the original source code.
'#          3. This notice may not be removed or altered from any source distribution.
'#              (NOTE: This license is borrowed from zLib.)
'#
'#  Please remember to vote on PSC.com if you like this code!
'#  Code URL: http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=37607&lngWId=1
'#####################################################################################

    '#### Declare the required APIs for StdIn/Out
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long

    '#### Declare the required consts for StdIn/StdOut/StdErr
Private Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_ERROR_HANDLE = -12&



'#####################################################################################
'# Public class subs/functions
'#####################################################################################
'#########################################################
'# Reads the data from StdIn, returning it to the caller
'#########################################################
Public Function StdIn() As String
    Dim sBuffer As String
    Dim lBytesRead As Long
    Dim lStdInHandle As Long
    Const iBufLen As Long = 1048576
        '#### Setup lStdInHandle and default the return value
    lStdInHandle = GetStdHandle(STD_INPUT_HANDLE)
    StdIn = ""

        '#### Loop while reading from StdIn
    Do
            '#### (Re)Set the sBuffer to a length of 1024
        sBuffer = String(iBufLen, 0)

            '#### If we can successfully read from StdIn
        If (ReadFile(lStdInHandle, ByVal sBuffer, iBufLen, lBytesRead, ByVal 0&)) Then
                '#### Append the data collected above onto the return value
            DoEvents
            StdIn = StdIn & Left$(sBuffer, lBytesRead)

            '#### Else StdIn was not able to to read, so fall out of the loop
        Else
            Exit Do
        End If
    Loop
End Function


'#########################################################
'# Writes the passed sData to StdOut
'#########################################################
Public Sub StdOut(sData As String)
        '#### Call the API, passing in STD_OUTPUT_HANDLE as the file handle to write to
    Call WriteFile(GetStdHandle(STD_OUTPUT_HANDLE), sData & vbCrLf, Len(sData & vbCrLf), 0, ByVal 0&)
End Sub


'#########################################################
'# Writes the passed sData to StdErr
'#########################################################
Public Sub StdErr(sData As String)
        '#### Call the API, passing in STD_ERROR_HANDLE as the file handle to write to
    Call WriteFile(GetStdHandle(STD_ERROR_HANDLE), sData & vbCrLf, Len(sData & vbCrLf), 0, ByVal 0&)
End Sub

