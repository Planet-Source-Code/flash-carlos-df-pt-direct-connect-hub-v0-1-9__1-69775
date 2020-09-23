Attribute VB_Name = "mListView"
Option Explicit

'----------------------------------------------'
'  Copyright (C) fLaSh - Carlos Ferreira       '
'  E-Mail: Carlosferreiracarlos@hotmail.com    '
'----------------------------------------------'

Private Const LVS_EX_FULLROWSELECT = &H20
Private Const LVS_EX_GRIDLINES = &H1
Private Const LVM_FIRST = &H1000
Private Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)
Private Const LVIS_STATEIMAGEMASK As Long = &HF000
Private Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)
Private Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55

Private Type LVITEM
   mask         As Long
   iItem        As Long
   iSubItem     As Long
   State        As Long
   stateMask    As Long
   pszText      As String
   cchTextMax   As Long
   iImage       As Long
   lParam       As Long
   iIndent      As Long
End Type

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
   (ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" _
   (ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Public Sub LVFullRow()
1:    Dim rStyle As Long
2:    Dim R As Long
   
4:    On Error GoTo Err

6:     With frmHub
          'get the current ListView style
8:        rStyle = SendMessageLong(.lvwCommands.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
9:        rStyle = SendMessageLong(.lvwPermIPBan.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
10:       rStyle = SendMessageLong(.lvwTempIPBan.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
12:       rStyle = SendMessageLong(.lvwUsers.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
13:       rStyle = SendMessageLong(.lvwPlugins.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
14:       rStyle = SendMessageLong(.lvwRegistered.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
15:       rStyle = SendMessageLong(.lvwBans.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)

          'set the extended bit
18:       rStyle = rStyle Or LVS_EX_FULLROWSELECT
   
          'set the new ListView style
21:       R = SendMessageLong(.lvwCommands.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
22:       R = SendMessageLong(.lvwPermIPBan.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
23:       R = SendMessageLong(.lvwTempIPBan.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
24:       R = SendMessageLong(.lvwUsers.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
25:       R = SendMessageLong(.lvwPlugins.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
26:       R = SendMessageLong(.lvwRegistered.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
27:       R = SendMessageLong(.lvwBans.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)

29:    End With
       
31: Exit Sub
32:
Err:
33: HandleError Err.Number, Err.Description, Erl & "|" & "mListView.LVFullRow()"
End Sub

Public Function LVOptFRow(frmLV As ListView)
1:    Dim rStyle As Long
2:    Dim R As Long
   
4:    On Error GoTo Err

          'get the current ListView style
8:        rStyle = SendMessageLong(frmLV.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)

          'set the extended bit
12:       rStyle = rStyle Or LVS_EX_FULLROWSELECT
   
          'set the new ListView style
15:       R = SendMessageLong(frmLV.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
       
17: Exit Function
18:
Err:
20: HandleError Err.Number, Err.Description, Erl & "|" & "mListView.LVOptFRow()"
End Function
