Attribute VB_Name = "modResAVI"
'// ---------------------------------------------------------------------------
'// Modul:    modResAVI
'//           AVI Animationen aus Ressourcen abspielen
'//
'// Copyright ©2001 Thorsten Dörfler (doerfler.t@vb-hellfire.de)
'//           http://www.vb-hellfire.de
'// ---------------------------------------------------------------------------
Option Explicit

Private Const WM_USER = &H400
Private Const ACM_OPEN = (WM_USER + 100)
Private Const ACM_PLAY = (WM_USER + 101)
Private Const ACM_STOP = (WM_USER + 102)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Enum ResAnimateConstants
  ranOpen = 1
  ranPlay = 2
  ranSeek = 3
  ranStop = 4
  ranClose = 5
End Enum

'// ---------------------------------------------------------------------------
'// Funktion:     ResAnimate
'//               AVI Animationen aus Ressourcen abspielen
'// Parameter:
'//   Animation   => Referenz auf ein Animation Control, das die Animation
'//                  wiedergeben soll.
'//   Cmd         => ResAnimate Konstante, die angibt welcher Befehl ausgeführt
'//                  werden soll:
'//                  ranOpen -> Öffnet die unter [ID] angegebene AVI Ressource
'//                  ranPlay -> Spielt die zuvor geöffnete Animation von
'//                             [StartFrame] bis [EndFrame] ab. [Repeat] legt
'//                             die Anzahl der Wiederholungen fest.
'//                             (Standard: alle Frames endlos abspielen).
'//                  ranSeek -> Zeigt ein bestimmtes Frame, [StartFrame] an.
'//                  ranStop -> Beendet die Wiedergabe
'//                  ranClose-> Schließt die Animation, löscht die Anzeige.
'// ---------------------------------------------------------------------------
Public Function ResAnimate(ByRef Animation As Animation, _
                           ByVal cmd As ResAnimateConstants, _
                  Optional ByVal ID As Long, _
                  Optional ByVal StartFrame As Integer = 0, _
                  Optional ByVal EndFrame As Integer = -1, _
                  Optional ByVal Repeat As Long = -1) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010   Dim lngRet As Long
50020
50031   Select Case cmd
          Case ranOpen
50050       '// Animation laden:
50060       lngRet = SendMessage(Animation.hwnd, ACM_OPEN, App.hInstance, ByVal ID)
50070     Case ranPlay
50080       '// Animation abspielen:
50090       lngRet = SendMessage(Animation.hwnd, ACM_PLAY, Repeat, _
                           ByVal CLng(EndFrame * &H10000 + StartFrame))
50110     Case ranSeek
50120       '// Bestimmtes Frame (StartFrame) abspielen:
50130       lngRet = SendMessage(Animation.hwnd, ACM_PLAY, 1, _
                           ByVal CLng(StartFrame * &H10000 + StartFrame))
50150     Case ranStop
50160       '// Animation anhalten:
50170       lngRet = SendMessage(Animation.hwnd, ACM_STOP, 0&, ByVal 0&)
50180     Case ranClose
50190       '// Animation schließen:
50200       lngRet = SendMessage(Animation.hwnd, ACM_OPEN, App.hInstance, _
                           ByVal vbNullString)
50220       '// Anzeige durch aus- und wiedereinblenden löschen:
50230       Animation.Visible = False
50240       Animation.Visible = True
50250   End Select
50260
50270   ResAnimate = (lngRet <> 0)
50280
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modResAVI", "ResAnimate")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
