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
  Dim lngRet As Long

  Select Case cmd
    Case ranOpen
      '// Animation laden:
      lngRet = SendMessage(Animation.hwnd, ACM_OPEN, App.hInstance, ByVal ID)
    Case ranPlay
      '// Animation abspielen:
      lngRet = SendMessage(Animation.hwnd, ACM_PLAY, Repeat, _
                           ByVal CLng(EndFrame * &H10000 + StartFrame))
    Case ranSeek
      '// Bestimmtes Frame (StartFrame) abspielen:
      lngRet = SendMessage(Animation.hwnd, ACM_PLAY, 1, _
                           ByVal CLng(StartFrame * &H10000 + StartFrame))
    Case ranStop
      '// Animation anhalten:
      lngRet = SendMessage(Animation.hwnd, ACM_STOP, 0&, ByVal 0&)
    Case ranClose
      '// Animation schließen:
      lngRet = SendMessage(Animation.hwnd, ACM_OPEN, App.hInstance, _
                           ByVal vbNullString)
      '// Anzeige durch aus- und wiedereinblenden löschen:
      Animation.Visible = False
      Animation.Visible = True
  End Select

  ResAnimate = (lngRet <> 0)

End Function
