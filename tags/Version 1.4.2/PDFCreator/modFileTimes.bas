Attribute VB_Name = "modFileTimes"
Option Explicit
' Based on http://www.vb-helper.com/howto_file_times.html

' Return False if there is an error.
Public Function GetFileTimes(ByVal filename As String, ByRef dateCreated As Date, ByRef dateAccessed As Date, _
 ByRef dateWritten As Date, ByVal localTime As Boolean) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020  Dim fileHandle As Long
50030  Dim creationTime As FILETIME, accessTime As FILETIME, writeTime As FILETIME, tmpFileTime As FILETIME
50040
50050  GetFileTimes = True
50060
50070 ' Open the file.
50080  fileHandle = CreateFile(filename, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
50090  If fileHandle = 0 Then
50100   GetFileTimes = False
50110   Exit Function
50120  End If
50130
50140 ' Get the times.
50150  If GetFileTime(fileHandle, creationTime, accessTime, writeTime) = 0 Then
50160   GetFileTimes = False
50170   Exit Function
50180  End If
50190
50200 ' Close the file.
50210  If CloseHandle(fileHandle) = 0 Then
50220   GetFileTimes = False
50230   Exit Function
50240  End If
50250
50260 ' See if we should convert to the local file system time.
50270  If localTime Then
50280 ' Convert to local file system time.
50290   FileTimeToLocalFileTime creationTime, tmpFileTime
50300   creationTime = tmpFileTime
50310   FileTimeToLocalFileTime accessTime, tmpFileTime
50320   accessTime = tmpFileTime
50330   FileTimeToLocalFileTime writeTime, tmpFileTime
50340   writeTime = tmpFileTime
50350  End If
50360
50370  ' Convert into dates.
50380  dateCreated = FileTimeToDate(creationTime)
50390  dateAccessed = FileTimeToDate(accessTime)
50400  dateWritten = FileTimeToDate(writeTime)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modFileTimes", "GetFileTimes")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

' Convert the FILETIME structure into a Date.
Private Function FileTimeToDate(ft As FILETIME) As Date
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 ' FILETIME units are 100s of nanoseconds.
50020  Const TICKS_PER_SECOND = 10000000
50030
50040  Dim lo_time As Double, hi_time As Double, seconds As Double, hours As Double
50050  Dim theDate As Date
50060
50070  ' Get the low order data.
50080  If ft.dwLowDateTime < 0 Then
50090    lo_time = 2 ^ 31 + (ft.dwLowDateTime And &H7FFFFFFF)
50100   Else
50110    lo_time = ft.dwLowDateTime
50120  End If
50130
50140  ' Get the high order data.
50150  If ft.dwHighDateTime < 0 Then
50160    hi_time = 2 ^ 31 + (ft.dwHighDateTime And &H7FFFFFFF)
50170   Else
50180    hi_time = ft.dwHighDateTime
50190  End If
50200
50210  ' Combine them and turn the result into hours.
50220  seconds = (lo_time + 2 ^ 32 * hi_time) / TICKS_PER_SECOND
50230  hours = CLng(seconds / 3600)
50240  seconds = seconds - hours * 3600
50250
50260  ' Make the date.
50270  theDate = DateAdd("h", hours, DateSerial(1601, 1, 1))
50280  theDate = DateAdd("s", seconds, theDate)
50290  FileTimeToDate = theDate
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modFileTimes", "FileTimeToDate")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

