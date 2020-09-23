Attribute VB_Name = "mMsgBox"
Option Explicit

'----------------------------------------------'
'  Copyright (C) fLaSh - Carlos Ferreira       '
'  E-Mail: Carlosferreiracarlos@hotmail.com    '
'----------------------------------------------'

'MsgBeep
'Play a wav file instead of the system default beep (if supported)
' Types ands API calls
Public Enum BeepType
  beepSystemDefault = &HFFFFFFFF  'same as using the VB Beep command
  beepSystemAsterisk = &H40&
  beepSystemExclamation = &H30&
  beepSystemHand = &H10&
  beepSystemQuestion = &H20&
'  beepSystemDefault = &H0&
End Enum

Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Private Declare Function waveOutGetNumDevs Lib "winmm" () As Long

'CenterMsgBoxOnForm
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Const GWL_HINSTANCE = (-6)
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10
Private Const HCBT_ACTIVATE = 5
Private Const WH_CBT = 5

Private hHook As Long
Private FrmhWnd As Long

' Beep function
Public Function MsgBeep(MsgType As BeepType) As Boolean
    On Error GoTo Err
    
3:  If waveOutGetNumDevs() Then
4:    Call MessageBeep(MsgType)
5:    MsgBeep = True              'we could sound off
6:  Else
7:    Beep
8:    MsgBeep = False             'sounded off with default beep
9:  End If

11: Exit Function

Err:
14: HandleError Err.Number, Err.Description, Erl & "|mMsgBox.MsgBeep()"
End Function

Public Function MsgBoxCenter(ParentForm As Form, Msg As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String = "PTDCH") As VbMsgBoxResult
    On Error GoTo Err
2:  Dim hInst As Long
3:  Dim Thread As Long
  
5:   If ParentForm.WindowState = vbMinimized Or ParentForm.Visible = False Then
6:      MsgBox Msg, Buttons, Title
7:   Else
       'Set up the CBT hook
9:      FrmhWnd = ParentForm.hWnd
10:     hInst = GetWindowLong(ParentForm.hWnd, GWL_HINSTANCE)
11:     Thread = GetCurrentThreadId()
12:     hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProc1, hInst, Thread)
        'Display the message box
14:     MsgBoxCenter = MsgBox(Msg, Buttons, Title)
15:  End If
  
17:  Exit Function

Err:
20:   HandleError Err.Number, Err.Description, Erl & "|mMsgBox.MsgBoxCenter()"
End Function

Private Function WinProc1(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
1:  Dim rectForm As RECT, rectMsg As RECT
2:  Dim X As Long, Y As Long
    On Error GoTo Err
    'On HCBT_ACTIVATE, show the MsgBox centered over Form1
5:  If lMsg = HCBT_ACTIVATE Then
       'Get the coordinates of the form and the message box so that
       'you can determine where the center of the form is located
8:     GetWindowRect FrmhWnd, rectForm
9:     GetWindowRect wParam, rectMsg
10:    X = (rectForm.Left + (rectForm.Right - rectForm.Left) / 2) - ((rectMsg.Right - rectMsg.Left) / 2)
11:    Y = (rectForm.Top + (rectForm.Bottom - rectForm.Top) / 2) - ((rectMsg.Bottom - rectMsg.Top) / 2)
       'Position the msgbox
13:    SetWindowPos wParam, 0, X, Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
       'Release the CBT hook
15:    UnhookWindowsHookEx hHook
16:  End If
17:  WinProc1 = False

19:   Exit Function

Err:
22:   HandleError Err.Number, Err.Description, Erl & "|mMsgBox.WinProc1()"
End Function
