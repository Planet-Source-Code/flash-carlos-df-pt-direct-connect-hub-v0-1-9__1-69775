Attribute VB_Name = "mSizing"
Option Explicit

'----------------------------------------------'
'  Copyright (C) fLaSh - Carlos Ferreira       '
'  E-Mail: Carlosferreiracarlos@hotmail.com    '
'----------------------------------------------'

'Sizing restriction support module. Sizable forms are hooked to this module

' Screen Points in Pixels
Private Type PointAPI
  X As Long
  Y As Long
End Type

' structure for window sizing
Private Type MINMAXINFO
  ptReserved As PointAPI
  ptMaxSize As PointAPI
  ptMaxPosition As PointAPI
  ptMinTrackSize As PointAPI
  ptMaxTrackSize As PointAPI
End Type

' used to assign new WndProc
Private Declare Function SetWindowLong Lib "user32" _
       Alias "SetWindowLongA" _
      (ByVal hWnd As Long, _
       ByVal nIndex As Long, _
       ByVal dwNewLong As Long) As Long

' used to invoke old WndProc
Private Declare Function CallWindowProc Lib "user32" _
       Alias "CallWindowProcA" _
      (ByVal lpPrevWndFunc As Long, _
       ByVal hWnd As Long, _
       ByVal uMsg As Long, _
       ByVal wParam As Long, _
       ByVal lParam As Long) As Long

' used to copy MINMAXINFO structure
Private Declare Sub CopyMemory Lib "kernel32" _
      Alias "RtlMoveMemory" _
     (hpvDest As Any, _
      hpvSource As Any, _
      ByVal cbCopy As Long)

'Save / Load frmHub position
Private Type FormPosition
    Left    As Long
    Top     As Long
    Width   As Long
    Height  As Long
End Type

'Stuff
Private Const GWL_WNDPROC As Long = (-4)
Private Const WM_GETMINMAXINFO As Long = &H24

'Publics
Public Const g_WinMinW As Long = 9405         'minimum main window width
Public Const g_WinMinH As Long = 5445         'minimum main window height

Public Const PrMinW As Long = 5000          'minimum main window width
Public Const PrMinH As Long = 2350          'minimum main window height

' Local Storage
Public G_HbWnd As Long   'subclass hook for Main win
Public G_PrWnd As Long   'subclass hook for Main win

' HookWin(): Subclass hwnd
Public Sub HookWin(ByVal hWnd As Long, PrvhWnd As Long)
1:  PrvhWnd = SetWindowLong( _
     hWnd, _
     GWL_WNDPROC, _
     AddressOf PWndProc)
End Sub

' UnhookWin(): remove subclass hook
Public Sub UnhookWin(ByVal hWnd As Long, PrvhWnd As Long)
1:  Call SetWindowLong( _
     hWnd, _
     GWL_WNDPROC, _
     PrvhWnd)
5   PrvhWnd = 0
End Sub

'  Subclassing Form. Copy the following, and add handling for additional
'                    messages of interest.
Private Function PWndProc( _
        ByVal hWnd As Long, _
        ByVal uMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long

2:  On Error GoTo Err

4:  Static Calcs As Boolean '1-time flag
5:  Static SX As Long       'screenx
6:  Static SY As Long       'screeny
7:  Static STX As Long      'TwipsPerPixelX
8:  Static STY As Long      'TwipsPerPixelY
  
10:  Dim Result As Long
11:  Dim MnMxInfo As MINMAXINFO

  ' do this one time
14:  If Not Calcs Then
15:    STX = Screen.TwipsPerPixelX
16:    STY = Screen.TwipsPerPixelY
17:    SX = Screen.Width \ STX
18:    SY = Screen.Height \ STY
19:    Calcs = True
20:  End If

  ' handle messages
23:  Select Case hWnd            'check which form
       Case frmHub.hWnd          'is main VC form
         Select Case uMsg        'check message
           Case WM_GETMINMAXINFO 'sizing
27:          CopyMemory MnMxInfo, lParam, LenB(MnMxInfo)
28:          With MnMxInfo
29:            With .ptMinTrackSize  'set min size
30:              .X = g_WinMinW \ STX
31:              .Y = g_WinMinH \ STY
32:            End With
33:            With .ptMaxPosition
34:              .X = 0
35:              .Y = 0
36:            End With
37:            With .ptMaxTrackSize
38:              .X = SX
39:              .Y = SY
40:            End With
41:            With .ptMaxSize
42:              .X = SX
43:              .Y = SY
44:            End With
45:          End With
46:          CopyMemory ByVal lParam, MnMxInfo, LenB(MnMxInfo)
47:          PWndProc = 0
           Case Else
49:          PWndProc = CallWindowProc( _
            G_HbWnd, _
            hWnd, _
            uMsg, _
            wParam, _
            lParam)
        End Select
     Case Else 'frmProperties
      Select Case uMsg        'check message
        Case WM_GETMINMAXINFO 'sizing
59:          CopyMemory MnMxInfo, lParam, LenB(MnMxInfo)
60:          With MnMxInfo
61:            With .ptMinTrackSize  'set min size
62:              .X = PrMinW \ STX
63:              .Y = PrMinH \ STY
64:            End With
65:            With .ptMaxPosition
66:              .X = 0
67:              .Y = 0
68:            End With
69:            With .ptMaxTrackSize
70:              .X = SX
71:              .Y = SY
72:            End With
73:            With .ptMaxSize
74:              .X = SX
75:              .Y = SY
76:            End With
77:          End With
78:          CopyMemory ByVal lParam, MnMxInfo, LenB(MnMxInfo)
79:          PWndProc = 0
        Case Else
81:          PWndProc = CallWindowProc( _
            G_PrWnd, _
            hWnd, _
            uMsg, _
            wParam, _
            lParam)
      End Select
    End Select
  
90: Exit Function
91:
Err:
93:    HandleError Err.Number, Err.Description, Erl & "|" & "mSizing.SaveFormSize()"
End Function

Public Sub RestoreFormSize()

2:  On Error GoTo Err

4:  Dim sData       As String
5:  Dim saSizes()   As String
6:  Dim uPosition   As FormPosition
    
8:    With uPosition
          'Retrieve the form's saved positions
10:        sData = g_objSettings.frmHubPosition
       
11:        If Len(sData) = 0 Then
12:            .Left = frmHub.Left
13:            .Top = frmHub.Top
14:        Else
15:            saSizes() = Split(sData, ",")
16:            If UBound(saSizes) < 4 Then
17:                ReDim Preserve saSizes(4)
18:            End If
19:            .Left = Val(Trim$(saSizes(0)))
20:            .Top = Val(Trim$(saSizes(1)))
21:        End If

22:        If .Left < 0 Then
23:            .Left = frmHub.Left
24:        End If
25:        If .Left > Screen.Width - .Width Then
26:            .Left = Screen.Width - .Width
27:        End If

29:        If .Top < 0 Then
30:            .Top = frmHub.Top
31:        End If
32:        If .Top > Screen.Height - .Height Then
33:            .Top = Screen.Height - .Height
34:        End If

          'Position the form. Moving the form here will establish
          'its normal restored positions.
38:        frmHub.Move .Left, .Top
39:    End With
    
41:  Exit Sub
42:
Err:
44:  HandleError Err.Number, Err.Description, Erl & "|" & "mSizing.RestoreFormSize()"
End Sub

Public Sub SaveFormSize()
1:    On Error GoTo Err
2:    Dim saSizes()   As String

4:    ReDim saSizes(1)
5:    If frmHub.WindowState = vbNormal Then
          'These values would be wrong if the
          'form was minimized or maximized.
8:        saSizes(0) = CStr(frmHub.Left)
9:        saSizes(1) = CStr(frmHub.Top)
10:    End If
        
12:   g_objSettings.frmHubPosition = Join(saSizes, ",")

13:  Exit Sub
14:
Err:
16:    HandleError Err.Number, Err.Description, Erl & "|" & "mSizing.SaveFormSize()"
End Sub
