Attribute VB_Name = "mSysTray"
Option Explicit

'----------------------------------------------'
'  Copyright (C) fLaSh - Carlos Ferreira       '
'  E-Mail: Carlosferreiracarlos@hotmail.com    '
'----------------------------------------------'

'BEGIN POPUP PROPER CODE /////////////////////////////////////////////////////////////
'This is required so that when a popup menu is called the menu
'dismisses correctly if no item is chosen
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'BEGIN EXPLORER.EXE CRASH DETECTION CODE /////////////////////////////////////////////
'This is used to find registered window messages, for explorer crash detection
'as well as to find the begining of program definable user messages ie: WM_APP
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

'This will be used to help generate a GUID for use with creating our
'WM_TRAYHOOK
Private Declare Function CoCreateGuid Lib "ole32.dll" (pGuid As GUID) As Long

Public Type NOTIFYICONDATA
    cbSize              As Long             'Size of NotifyIconData struct
    hWnd                As Long             'Window handle for the window handling the icon events
    uID                 As Long             'Icon ID (to allow multiple icons per application)
    uFlags              As Long             'NIF Flags
    uCallbackMessage    As Long             'The message received for the system tray icon
    hIcon               As Long             'The memory location of our icon if NIF_ICON is specifed
    szTip               As String * 128     'Tooltip if NIF_TIP is specified (64 characters max)
    dwState             As Long
    dwStateMask         As Long
    szInfo              As String * 256
    uTimeout            As Long
    szInfoTitle         As String * 64
    dwInfoFlags         As Long
End Type

Public Enum BalloonIcon
    ICON_NONE = 0
    ICON_INFO = 1
    ICON_WARNING = 2
    ICON_ERROR = 3
    ICON_USER = 4
End Enum

'BEGIN GUID HELPER
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

'Public mWndProcNext As Long
'Private bIsHooked As Boolean

Private TrayIcon As NOTIFYICONDATA

'used to indentify different tray icons if used
'Public WM_APP As Long  'For user defined window messages
Public WM_TRAYHOOK As Long 'The tray icon window message

'BEGIN EXPLORER.EXE CRASH DETECTION CODE
Public mTaskbarCreated As Long

'CONSTANT DECLARES
'--------------------------------
'BEGIN HOOKING CODE
Private Const GWL_WNDPROC = (-4)
Private Const GWL_USERDATA = (-21)

'Window messages relating to balloon tips and the like branch from here
Private Const WM_USER As Long = &H400

'Here are some mouse "events" to play with
'we are only going to use two in our example,
'however you feel free to use whatever you'd like (:

'Left Button
Public Const WM_LBUTTONDOWN As Long = &H201
Public Const WM_LBUTTONUP As Long = &H202
Public Const WM_LBUTTONDBLCLK As Long = &H203
' Middle Button
Public Const WM_MBUTTONDOWN As Long = &H207
Public Const WM_MBUTTONUP As Long = &H208
Public Const WM_MBUTTONDBLCLK As Long = &H209
' Right Button
Public Const WM_RBUTTONDOWN As Long = &H204
Public Const WM_RBUTTONUP As Long = &H205
Public Const WM_RBUTTONDBLCLK As Long = &H206

' Shell_NotifyIconA() messages
Private Const NIM_ADD As Long = &H0     'Add icon to the System Tray
Private Const NIM_MODIFY As Long = &H1  'Modify System Tray icon
Private Const NIM_DELETE As Long = &H2  'Delete icon from System Tray

'NotifyIconData Flags
Private Const NIF_MESSAGE As Long = &H1 'uCallbackMessage in NOTIFYICONDATA is valid
Private Const NIF_ICON As Long = &H2 ' hIcon in NOTIFYICONDATA is valid
Private Const NIF_TIP As Long = &H4 'szTip in NOTIFYICONDATA is valid
Private Const NIF_INFO As Long = &H10 'for use with balloons

'Balloon tip icon constants
Private Const NIIF_NONE As Long = &H0
Private Const NIIF_WARNING As Long = &H2
Private Const NIIF_ERROR As Long = &H3
Private Const NIIF_USER As Long = &H4
Private Const NIIF_INFO As Long = &H1

'Balloon tip sound constants
Private Const NIIF_NOSOUND As Long = &H10

'Balloon tip notification messages
Public Const NIN_BALLOONSHOW As Long = WM_USER + &H2 'when the balloon is drawn
Public Const NIN_BALLOONHIDE As Long = WM_USER + &H3 'when the balloon disappears—for example, when the icon is deleted. This message is not sent if the balloon is dismissed because of a timeout or a mouse click.
Public Const NIN_BALLOONTIMEOUT As Long = WM_USER + &H4 'when the balloon is dismissed because of a timeout
Public Const NIN_BALLOONUSERCLICK As Long = WM_USER + &H5 'when the balloon is dismissed because of a mouse click.

Public Function CreateTrayIcon(ByRef Owner As Form, _
                               ByVal luID As Long, _
                      Optional ByRef ToolTip As String = "", _
                      Optional ByRef tIcon As StdPicture) As Long
                    
2:    On Error GoTo Err

      With TrayIcon
4:        .cbSize = Len(TrayIcon) 'This size is always the len(NOTIFYICONDATA)
5:        .hWnd = Owner.hWnd  'Which form is this icon for
6:        .uID = luID
7:        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE    'set valid data inputs
        
        'You see this is where most VB tray icon codes are bad, no offense
        'they use a hack that uses the message WM_MOUSEMOVE for notification
        'this way VB can handle the message using its built in
        'Form_MouseMove event. This is not the way to do it in my opinion
        'because you can't have multiple icons for one form, my method allows
        'for this, not to mention the inability to detect explorer.exe crashes
        'but using the hack method you don't need to use message hooking either...
        'What window message should be sent during an event
17:        .uCallbackMessage = WM_TRAYHOOK
18:        .szTip = Trim(ToolTip$) & vbNullChar   'set the tooltip
19:        If tIcon Is Nothing Then
20:            .hIcon = Owner.Icon
21:        Else
22:            .hIcon = tIcon
23:        End If

      End With
    'Create the tray icon with an API call
27:    CreateTrayIcon = Shell_NotifyIcon(NIM_ADD, TrayIcon)
28:  Exit Function

30:
Err:
32:    HandleError Err.Number, Err.Description, Erl & "|" & "mSysTray.CreateTrayIcon()"
End Function

Public Function ModifyTrayIcon(ByRef Owner As Form, _
                               ByVal luID As Long, _
                      Optional ByRef ToolTip As String = "", _
                      Optional ByRef tIcon As StdPicture) As Long
    
2:    On Error GoTo Err

4:    With TrayIcon
5:        .cbSize = Len(TrayIcon)
6:        .hWnd = Owner.hWnd
7:        .uID = luID
8:        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
9:        .uCallbackMessage = WM_TRAYHOOK
10:        If Not tIcon Is Nothing Then
11:            .hIcon = tIcon
12:        End If
13:        If ToolTip <> "" Then .szTip = Trim(ToolTip$) & vbNullChar
14:    End With
    'Update the tray icon with an API call
16:    ModifyTrayIcon = Shell_NotifyIcon(NIM_MODIFY, TrayIcon)
17:  Exit Function
18:
Err:
20:  HandleError Err.Number, Err.Description, Erl & "|" & "mSysTray.ModifyTrayIcon()"
End Function

Public Function DeleteTrayIcon(ByVal luID As Long) As Long

2:    On Error GoTo Err

3:    With TrayIcon
4:        .cbSize = Len(TrayIcon)
5:        .uID = luID
6:        .uFlags = NIM_DELETE
7:        .uCallbackMessage = WM_TRAYHOOK
8:    End With
    
     'Remove the tray icon with an API call
11:   DeleteTrayIcon = Shell_NotifyIcon(NIM_DELETE, TrayIcon)
12:  Exit Function
13:
Err:
15:    HandleError Err.Number, Err.Description, Erl & "|" & "mSysTray.DeleteTrayIcon()"
End Function

Public Function InsertHook(ByRef Owner As Form) As Long
    
2:    On Error GoTo Err

4:    Dim lResult As Long
    
    'Remove preexisting hook
    'Call RemoveHook(Owner)
    
9:    InsertHook = SetWindowLong(Owner.hWnd, GWL_WNDPROC, AddressOf GlobalMessageCatcher)
10:    If InsertHook Then
11:        lResult = SetWindowLong(Owner.hWnd, GWL_USERDATA, ObjPtr(Owner))
        'bIsHooked = True
13:    End If
14:  Exit Function
15:
Err:
17:    HandleError Err.Number, Err.Description, Erl & "|" & "mSysTray.InsertHook()"
End Function

Public Sub RemoveHook(ByRef Owner As Form, _
                      ByVal lHookID As Long)

2:    On Error GoTo Err
     'Remove the hook and revert control back to VB
    
4:    If lHookID Then    'Make sure we really are hooked
5:        SetWindowLong Owner.hWnd, GWL_WNDPROC, lHookID
        'bIsHooked = False
7:    End If

9:   Exit Sub
10:
Err:
12:  HandleError Err.Number, Err.Description, Erl & "|" & "mSysTray.RemoveHook()"
End Sub

Public Function GlobalMessageCatcher(ByVal shWnd As Long, _
                                     ByVal uMsg As Long, _
                                     ByVal wParam As Long, _
                                     ByVal lParam As Long) As Long
                                     
2:    On Error GoTo Err
    
4:    If shWnd = frmHub.hWnd Then
5:        GlobalMessageCatcher = frmHub.WindowProcSysTray(shWnd, uMsg, wParam, lParam)
6:    End If
    
8:    Exit Function

10:
Err:
12:   HandleError Err.Number, Err.Description, Erl & "|" & "mSysTray.GlobalMessageCatcher()"
End Function

Public Function PopupBalloon(ByRef Owner As Form, _
                             ByVal luID As Long, _
                             ByRef Title As String, _
                             ByRef Message As String, _
                    Optional ByVal IconType As BalloonIcon = ICON_INFO, _
                    Optional ByVal Sound As Boolean = True, _
                    Optional ByRef tIcon As StdPicture) As Long

2:    On Error GoTo Err

    'This line is optional, if you include it new balloon tips erase old ones
    'if you omit it a balloon tip queue so to speak is created, and as they timeout
    'new ones appear
7:    Call RemoveBalloon(Owner, luID)
8:    With TrayIcon
9:         .cbSize = Len(TrayIcon)
10:        .hWnd = Owner.hWnd
11:        .uID = luID
12:        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIM_MODIFY
13:        .uCallbackMessage = WM_TRAYHOOK
14:        If tIcon Is Nothing Then
15:            .hIcon = Owner.Icon
16:        Else
17:            .hIcon = tIcon
18:        End If
19:        .dwState = 0
20:        .dwStateMask = 0
21:        .szInfo = Message & Chr(0)
22:        .szInfoTitle = Title & Chr(0)
23:        Select Case IconType
              Case ICON_NONE
25:                .dwInfoFlags = NIIF_NONE
              Case ICON_INFO
27:                .dwInfoFlags = NIIF_INFO
              Case ICON_WARNING
28:                .dwInfoFlags = NIIF_WARNING
              Case ICON_ERROR
31:                .dwInfoFlags = NIIF_ERROR
              Case ICON_USER
33:                .dwInfoFlags = NIIF_USER
           End Select
35:        If Not Sound Then .dwInfoFlags = .dwInfoFlags Or NIIF_NOSOUND
36:    End With
    
38:    PopupBalloon = Shell_NotifyIcon(NIM_MODIFY, TrayIcon)

40:  Exit Function
41:
Err:
43:    HandleError Err.Number, Err.Description, Erl & "|" & "mSysTray.PopupBalloon()"
End Function

Public Function RemoveBalloon(ByRef Owner As Form, _
                              ByVal luID As Long) As Long

2:    On Error GoTo Err
3:    With TrayIcon
4:        .cbSize = Len(TrayIcon)
5:        .hWnd = Owner.hWnd
6:        .uID = luID
7:        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIM_MODIFY
8:        .uCallbackMessage = WM_TRAYHOOK
9:        .hIcon = Owner.Icon
10:        .dwState = 0
11:        .dwStateMask = 0
12:        .szInfo = Chr(0)
13:        .szInfoTitle = Chr(0)
14:        .dwInfoFlags = NIIF_NONE
15:    End With
16:    RemoveBalloon = Shell_NotifyIcon(NIM_MODIFY, TrayIcon)

18:  Exit Function
19:
Err:
21:    HandleError Err.Number, Err.Description, Erl & "|" & "mSysTray.RemoveBalloon()"
End Function

Public Function GetGUID() As String

2:   On Error GoTo Err
 
3:   Dim udtGUID As GUID

4:   If (CoCreateGuid(udtGUID) = 0) Then
5:      GetGUID = _
        String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
        String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
        String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
        IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
        IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
        IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
        IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
        IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
        IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
        IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
        IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
17:  End If
    
19:  Exit Function

21:
Err:
23:    HandleError Err.Number, Err.Description, Erl & "|" & "mSysTray.GetGUID()"
End Function
