Attribute VB_Name = "mTextBoxs"
Option Explicit

'----------------------------------------------'
'  Copyright (C) fLaSh - Carlos Ferreira       '
'  E-Mail: Carlosferreiracarlos@hotmail.com    '
'----------------------------------------------'

' Delete popup menu in textboxs

Private Declare Function SetWindowLong Lib "user32" Alias _
   "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, _
   ByVal dwNewLong As Long) As Long
 
Private Declare Function CallWindowProc Lib "user32" _
   Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
   ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, _
   ByVal lParam As Long) As Long

Private Const GWL_WNDPROC = (-4)

Private Const WM_CONTEXTMENU = &H7B

Global lpPrevWndProc As Long
Global gHW As Long

Public Sub DelPopUpMenu()

2:    Dim i As Integer
  
4:    On Error GoTo Err
  
6:    For i = 0 To frmHub.txtVSl.count - 1
7:        gHW = frmHub.txtVSl(i).hWnd
8:        Hook
9:    Next i
  
11:   Exit Sub
12:
Err:
14:   HandleError Err.Number, Err.Description, Erl & "|" & "mTextBoxs.DelPopUpMenu()"
End Sub

Private Sub Hook()
1:   On Error GoTo Err
2:   lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, _
                                    AddressOf gWindowProc)
4:   Exit Sub
5:
Err:
7:   HandleError Err.Number, Err.Description, Erl & "|" & "mTextBoxs.Hook()"
End Sub

Private Sub Unhook()
1:   Dim temp As Long
2:   temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub

Private Function gWindowProc(ByVal hWnd As Long, ByVal Msg As Long, _
                 ByVal wParam As Long, ByVal lParam As Long) As Long
1: On Error GoTo Err

3:     If Msg = WM_CONTEXTMENU Then
4:        'Debug.Print "Intercepted WM_CONTEXTMENU at " & Now
5:        gWindowProc = True
6:     Else ' Send all other messages to the default message handler
7:        gWindowProc = CallWindowProc(lpPrevWndProc, hWnd, Msg, wParam, lParam)
8:     End If
     
10:   Exit Function
11:
Err:
13:   HandleError Err.Number, Err.Description, Erl & "|" & "mTextBoxs.gWindowProc()"
End Function
