VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMagneticWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'========================================================================================
' Subclasser declarations
'========================================================================================

Private Enum eMsgWhen
    [MSG_AFTER] = 1                                  'Message calls back after the original (previous) WndProc
    [MSG_BEFORE] = 2                                 'Message calls back before the original (previous) WndProc
    [MSG_BEFORE_AND_AFTER] = MSG_AFTER Or MSG_BEFORE 'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES     As Long = -1          'All messages added or deleted
Private Const CODE_LEN         As Long = 197         'Length of the machine code in bytes
Private Const GWL_WNDPROC      As Long = -4          'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04         As Long = 88          'Table B (before) address patch offset
Private Const PATCH_05         As Long = 93          'Table B (before) entry count patch offset
Private Const PATCH_08         As Long = 132         'Table A (after) address patch offset
Private Const PATCH_09         As Long = 137         'Table A (after) entry count patch offset

Private Type tSubData                                'Subclass data type
    hWnd                       As Long               'Handle of the window being subclassed
    nAddrSub                   As Long               'The address of our new WndProc (allocated memory).
    nAddrOrig                  As Long               'The address of the pre-existing WndProc
    nMsgCntA                   As Long               'Msg after table entry count
    nMsgCntB                   As Long               'Msg before table entry count
    aMsgTblA()                 As Long               'Msg after table array
    aMsgTblB()                 As Long               'Msg Before table array
End Type

Private sc_aSubData()          As tSubData           'Subclass data array
Private sc_aBuf(1 To CODE_LEN) As Byte               'Code buffer byte array
Private sc_pCWP                As Long               'Address of the CallWindowsProc
Private sc_pEbMode             As Long               'Address of the EbMode IDE break/stop/running function
Private sc_pSWL                As Long               'Address of the SetWindowsLong function
  
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

'========================================================================================
' cMagneticWnd
'========================================================================================
'-- API

Private Type PointAPI
    X1 As Long
    Y1 As Long
End Type

Private Type RECT2
    X1 As Long
    Y1 As Long
    X2 As Long
    Y2 As Long
End Type

Private Const SPI_GETWORKAREA  As Long = 48

Private Const WM_SIZING        As Long = &H214
Private Const WM_MOVING        As Long = &H216
Private Const WM_ENTERSIZEMOVE As Long = &H231
Private Const WM_EXITSIZEMOVE  As Long = &H232
Private Const WM_SYSCOMMAND    As Long = &H112
Private Const WM_COMMAND       As Long = &H111

Private Const WMSZ_LEFT        As Long = 1
Private Const WMSZ_RIGHT       As Long = 2
Private Const WMSZ_TOP         As Long = 3
Private Const WMSZ_TOPLEFT     As Long = 4
Private Const WMSZ_TOPRIGHT    As Long = 5
Private Const WMSZ_BOTTOM      As Long = 6
Private Const WMSZ_BOTTOMLEFT  As Long = 7
Private Const WMSZ_BOTTOMRIGHT As Long = 8

Private Const SC_MINIMIZE      As Long = &HF020&
Private Const SC_RESTORE       As Long = &HF120&

Private Const SWP_NOSIZE       As Long = &H1
Private Const SWP_NOZORDER     As Long = &H4
Private Const SWP_NOACTIVATE   As Long = &H10

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function BeginDeferWindowPos Lib "user32" (ByVal nNumWindows As Long) As Long
Private Declare Function DeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long, ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function EndDeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT2) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT2, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function UnionRect Lib "user32" (lpDestRect As RECT2, lpSrc1Rect As RECT2, lpSrc2Rect As RECT2) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

'-- Private types:

Private Type WND_INFO
    hWnd       As Long
    hWndParent As Long
    Glue       As Boolean
End Type

'-- Private constants:

Private Const LB_RECT As Long = 16

'-- Private variables:

Private m_uWndInfo()  As WND_INFO
Private m_lWndCount   As Long
Private m_rcWnd()     As RECT2
Private m_ptAnchor    As PointAPI
Private m_ptOffset    As PointAPI
Private m_ptCurr      As PointAPI
Private m_ptLast      As PointAPI

'-- Property variables:

Private m_lSnapWidth As Long

Private Sub Class_Initialize()
    
    '-- Default snap width
    m_lSnapWidth = 10
    
    '-- Initialize array (handled windows info)
    ReDim m_uWndInfo(0) As WND_INFO
    m_lWndCount = 0
End Sub

Private Sub Class_Terminate()
    
    '-- Stop subclassing
    If (m_lWndCount) Then
        Call Subclass_StopAll
    End If
End Sub

'========================================================================================
' Subclass handler: MUST be the first Public routine in this file.
'                   That includes public properties also.
'========================================================================================

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
'
'Parameters:
'   bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
'   bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
'   lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
'   lng_hWnd - The window handle
'   uMsg     - The message number
'   wParam   - Message related data
'   lParam   - Message related data
'
'Notes:
'   If you really know what you're doing, it's possible to change the values of the
'   hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
'   values get passed to the default handler.. and optionaly, the 'after' callback

16:    On Error GoTo Err

18:    Dim rcWnd As RECT2
19:    Dim LC    As Long
  
21:    Select Case uMsg
        
        '-- Size/Move starting
           Case WM_ENTERSIZEMOVE
            
            '-- Get Desktop area (as first rectangle)
27:            Call SystemParametersInfo(SPI_GETWORKAREA, 0, m_rcWnd(0), 0)
            
            '-- Get rectangles of all handled windows
30:            For LC = 1 To m_lWndCount
                
                '-- Window maximized ?
33:                If (IsZoomed(m_uWndInfo(LC).hWnd)) Then
                    '-- Take work are rectangle
35:                    Call CopyMemory(m_rcWnd(LC), m_rcWnd(0), LB_RECT)
36:                Else
                    '-- Get window rectangle
38:                    Call GetWindowRect(m_uWndInfo(LC).hWnd, m_rcWnd(LC))
39:                End If
                
                '-- Is it our current window ?
42:                If (m_uWndInfo(LC).hWnd = lng_hWnd) Then
                    '-- Get anchor-offset
44:                    Call GetCursorPos(m_ptAnchor)
45:                    Call GetCursorPos(m_ptLast)
46:                    m_ptOffset.X1 = m_rcWnd(LC).X1 - m_ptLast.X1
47:                    m_ptOffset.Y1 = m_rcWnd(LC).Y1 - m_ptLast.Y1
48:                End If
49:            Next LC
        
        '-- Sizing
52:        Case WM_SIZING
            
54:            Call CopyMemory(rcWnd, ByVal lParam, LB_RECT)
55:            Call pvSizeRect(lng_hWnd, rcWnd, wParam)
56:            Call CopyMemory(ByVal lParam, rcWnd, LB_RECT)
            
58:            bHandled = True
59:            lReturn = 1
        
        '-- Moving
62:        Case WM_MOVING
            
64:            Call CopyMemory(rcWnd, ByVal lParam, LB_RECT)
65:            Call pvMoveRect(lng_hWnd, rcWnd)
66:            Call CopyMemory(ByVal lParam, rcWnd, LB_RECT)
            
68:            bHandled = True
69:            lReturn = 1
        
        '-- Size/Move finishing
72:        Case WM_EXITSIZEMOVE
            
74:            Call pvCheckGlueing
            
        '-- Special case: *menu* call
77:        Case WM_SYSCOMMAND
            
79:            If (wParam = SC_MINIMIZE Or wParam = SC_RESTORE) Then
80:                Call pvCheckGlueing
81:            End If
        
        '-- Special case: *control* call
84:        Case WM_COMMAND
            
86:            Call pvCheckGlueing
87:    End Select
88:  Exit Sub
89:
Err:
90:  HandleError Err.Number, Err.Description, Erl & "|" & "clsMagneticWnd.zSubclass_Proc()"
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Function AddWindow(ByVal hWnd As Long, Optional ByVal hWndParent As Long = 0) As Boolean
1:    On Error GoTo Err
2:    Dim LC As Long
    
    '-- Already in collection ?
4:    For LC = 1 To m_lWndCount
5:        If (hWnd = m_uWndInfo(LC).hWnd) Then Exit Function
6:    Next LC
    
    '-- Validate windows
9:    If (IsWindow(hWnd) And (IsWindow(hWndParent) Or hWndParent = 0)) Then
        
        '-- Increase count
12:        m_lWndCount = m_lWndCount + 1
        '-- Resize arrays
14:        ReDim Preserve m_uWndInfo(0 To m_lWndCount)
15:        ReDim Preserve m_rcWnd(0 To m_lWndCount)
        
        '-- Add info
18:        With m_uWndInfo(m_lWndCount)
19:            .hWnd = hWnd
20:            .hWndParent = hWndParent
21:        End With
        
        '-- Check glueing for first time
24:        Call pvCheckGlueing
        
        '-- Start subclassing
27:        Call Subclass_Start(hWnd)
28:        Call Subclass_AddMsg(hWnd, WM_ENTERSIZEMOVE)
29:        Call Subclass_AddMsg(hWnd, WM_SIZING, [MSG_BEFORE])
30:        Call Subclass_AddMsg(hWnd, WM_MOVING, [MSG_BEFORE])
31:        Call Subclass_AddMsg(hWnd, WM_EXITSIZEMOVE)
32:        Call Subclass_AddMsg(hWnd, WM_SYSCOMMAND)
33:        Call Subclass_AddMsg(hWnd, WM_COMMAND)
        
        '-- Success
36:        AddWindow = True
37:    End If
38:  Exit Function
39:
Err:
41:    HandleError Err.Number, Err.Description, Erl & "|" & "clsMagneticWnd.AddWindow()"
End Function

Public Function RemoveWindow(ByVal hWnd As Long) As Boolean
1:    On Error GoTo Err
2:    Dim lc1 As Long
3:    Dim lc2 As Long

5:    For lc1 = 1 To m_lWndCount
        
7:        If (hWnd = m_uWndInfo(lc1).hWnd) Then
            
            '-- Move down
10:            For lc2 = lc1 To m_lWndCount - 1
11:                m_uWndInfo(lc2) = m_uWndInfo(lc2 + 1)
12:            Next lc2
            
            '-- Resize arrays
15:            m_lWndCount = m_lWndCount - 1
16:            ReDim Preserve m_uWndInfo(m_lWndCount)
17:            ReDim Preserve m_rcWnd(m_lWndCount)
            
            '-- Remove parent relationships
20:            For lc2 = 1 To m_lWndCount
21:                If (m_uWndInfo(lc2).hWndParent = hWnd) Then
22:                    m_uWndInfo(lc2).hWndParent = 0
23:                End If
24:            Next lc2
            
            '-- Stop subclassing / verify connections
27:            Call Subclass_Stop(hWnd)
28:            Call pvCheckGlueing
            
            '-- Success
31:            RemoveWindow = True
32:            Exit For
33:        End If
34:    Next lc1
35:  Exit Function
36:
Err:
38:    HandleError Err.Number, Err.Description, Erl & "|" & "clsMagneticWnd.RemoveWindow()"
End Function

Public Sub CheckGlueing()
    '-- Check ALL windows for possible new *connections*.
2:    Call pvCheckGlueing
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Get SnapWidth() As Long
1:    SnapWidth = m_lSnapWidth
End Property

Public Property Let SnapWidth(ByVal New_SnapWidth As Long)
1:    m_lSnapWidth = New_SnapWidth
End Property

'========================================================================================
' Private
'========================================================================================

Private Sub pvSizeRect(ByVal hWnd As Long, rcWnd As RECT2, ByVal lfEdge As Long)
1:    On Error GoTo Err
2:    Dim rcTmp As RECT2
3:    Dim LC    As Long
    
    '-- Get a copy
6:    Call CopyMemory(rcTmp, rcWnd, LB_RECT)
    
    '-- Check all windows
9:    For LC = 0 To m_lWndCount
        
11:        With m_rcWnd(LC)
            
            '-- Avoid current window
14:            If (m_uWndInfo(LC).hWnd <> hWnd) Then
                
                '-- X magnetism
17:                If (rcWnd.Y1 < .Y2 + m_lSnapWidth And rcWnd.Y2 > .Y1 - m_lSnapWidth) Then
                    
                      Select Case lfEdge
                        
                         Case WMSZ_LEFT, WMSZ_TOPLEFT, WMSZ_BOTTOMLEFT
                    
22:                        Select Case True
                             Case Abs(rcTmp.X1 - .X1) < m_lSnapWidth: rcWnd.X1 = .X1
                             Case Abs(rcTmp.X1 - .X2) < m_lSnapWidth: rcWnd.X1 = .X2
                           End Select
                
                         Case WMSZ_RIGHT, WMSZ_TOPRIGHT, WMSZ_BOTTOMRIGHT
                        
29:                        Select Case True
                             Case Abs(rcTmp.X2 - .X1) < m_lSnapWidth: rcWnd.X2 = .X1
                             Case Abs(rcTmp.X2 - .X2) < m_lSnapWidth: rcWnd.X2 = .X2
                          End Select
                       End Select
34:                End If
                
                '-- Y magnetism
37:                If (rcWnd.X1 < .X2 + m_lSnapWidth And rcWnd.X2 > .X1 - m_lSnapWidth) Then
                    
                       Select Case lfEdge
                        
                         Case WMSZ_TOP, WMSZ_TOPLEFT, WMSZ_TOPRIGHT
                        
43:                        Select Case True
                             Case Abs(rcTmp.Y1 - .Y1) < m_lSnapWidth: rcWnd.Y1 = .Y1
                             Case Abs(rcTmp.Y1 - .Y2) < m_lSnapWidth: rcWnd.Y1 = .Y2
                           End Select
                    
                         Case WMSZ_BOTTOM, WMSZ_BOTTOMLEFT, WMSZ_BOTTOMRIGHT
                        
50:                        Select Case True
                             Case Abs(rcTmp.Y2 - .Y1) < m_lSnapWidth: rcWnd.Y2 = .Y1
                             Case Abs(rcTmp.Y2 - .Y2) < m_lSnapWidth: rcWnd.Y2 = .Y2
                          End Select
54:                    End Select
55:                End If
56:            End If
57:        End With
58:    Next LC
59:  Exit Sub
60:
Err:
62:  HandleError Err.Number, Err.Description, Erl & "|" & "clsMagneticWnd.pvSizeRect()"
End Sub

Private Sub pvMoveRect(ByVal hWnd As Long, rcWnd As RECT2)
1:   On Error GoTo Err
2:     Dim lc1   As Long
3:     Dim lc2   As Long
4:     Dim lWId  As Long
5:     Dim rcTmp As RECT2
6:     Dim lOffx As Long
7:     Dim lOffy As Long
8:     Dim hDWP  As Long
    
    '== Get current cursor position
    
12:    Call GetCursorPos(m_ptCurr)
    
    '== Check magnetism for current window
    
    '-- 'Move' current window
17:    Call OffsetRect(rcWnd, (m_ptCurr.X1 - rcWnd.X1) + m_ptOffset.X1, 0)
18:    Call OffsetRect(rcWnd, 0, (m_ptCurr.Y1 - rcWnd.Y1) + m_ptOffset.Y1)
    
    '-- Check all windows
21:    For lc1 = 0 To m_lWndCount
        
        '-- Avoid current window
24:        If (m_uWndInfo(lc1).hWnd <> hWnd) Then
                
            '-- Avoid child windows
27:            If (m_uWndInfo(lc1).Glue = False Or m_uWndInfo(lc1).hWndParent <> hWnd) Then
                    
29:                With m_rcWnd(lc1)
                
                    '-- X magnetism
32:                    If (rcWnd.Y1 < .Y2 + m_lSnapWidth And rcWnd.Y2 > .Y1 - m_lSnapWidth) Then
                    
34:                        Select Case True
                             Case Abs(rcWnd.X1 - .X1) < m_lSnapWidth: lOffx = .X1 - rcWnd.X1
                             Case Abs(rcWnd.X1 - .X2) < m_lSnapWidth: lOffx = .X2 - rcWnd.X1
                             Case Abs(rcWnd.X2 - .X1) < m_lSnapWidth: lOffx = .X1 - rcWnd.X2
                             Case Abs(rcWnd.X2 - .X2) < m_lSnapWidth: lOffx = .X2 - rcWnd.X2
39:                        End Select
40:                    End If
                    
                    '-- Y magnetism
43:                    If (rcWnd.X1 < .X2 + m_lSnapWidth And rcWnd.X2 > .X1 - m_lSnapWidth) Then
                    
45:                        Select Case True
                             Case Abs(rcWnd.Y1 - .Y1) < m_lSnapWidth: lOffy = .Y1 - rcWnd.Y1
                             Case Abs(rcWnd.Y1 - .Y2) < m_lSnapWidth: lOffy = .Y2 - rcWnd.Y1
                             Case Abs(rcWnd.Y2 - .Y1) < m_lSnapWidth: lOffy = .Y1 - rcWnd.Y2
                             Case Abs(rcWnd.Y2 - .Y2) < m_lSnapWidth: lOffy = .Y2 - rcWnd.Y2
50:                        End Select
51:                    End If
52:                End With
53:            End If
54:        End If
55:    Next lc1
    
    '== Check magnetism for child windows
    
59:    For lc1 = 1 To m_lWndCount
        
        '-- Child and connected window ?
62:        If (m_uWndInfo(lc1).Glue And m_uWndInfo(lc1).hWndParent = hWnd) Then
            
            '-- 'Move' child window
65:            Call CopyMemory(rcTmp, m_rcWnd(lc1), LB_RECT)
66:            Call OffsetRect(rcTmp, m_ptCurr.X1 - m_ptAnchor.X1, 0)
67:            Call OffsetRect(rcTmp, 0, m_ptCurr.Y1 - m_ptAnchor.Y1)
            
69:            For lc2 = 0 To m_lWndCount
                                        
71:                If (lc1 <> lc2) Then
                    
                    '-- Avoid child windows
74:                    If (m_uWndInfo(lc2).Glue = False And m_uWndInfo(lc2).hWnd <> hWnd) Then
                    
76:                        With m_rcWnd(lc2)
                    
                            '-- X magnetism
79:                            If (rcTmp.Y1 < .Y2 + m_lSnapWidth And rcTmp.Y2 > .Y1 - m_lSnapWidth) Then
                                
81:                               Select Case True
                                     Case Abs(rcTmp.X1 - .X1) < m_lSnapWidth: lOffx = .X1 - rcTmp.X1
                                     Case Abs(rcTmp.X1 - .X2) < m_lSnapWidth: lOffx = .X2 - rcTmp.X1
                                     Case Abs(rcTmp.X2 - .X1) < m_lSnapWidth: lOffx = .X1 - rcTmp.X2
                                     Case Abs(rcTmp.X2 - .X2) < m_lSnapWidth: lOffx = .X2 - rcTmp.X2
86:                                End Select
87:                            End If
                            
                            '-- Y magnetism
90:                            If (rcTmp.X1 < .X2 + m_lSnapWidth And rcTmp.X2 > .X1 - m_lSnapWidth) Then
                            
92:                                Select Case True
                                     Case Abs(rcTmp.Y1 - .Y1) < m_lSnapWidth: lOffy = .Y1 - rcTmp.Y1
                                     Case Abs(rcTmp.Y1 - .Y2) < m_lSnapWidth: lOffy = .Y2 - rcTmp.Y1
                                     Case Abs(rcTmp.Y2 - .Y1) < m_lSnapWidth: lOffy = .Y1 - rcTmp.Y2
                                     Case Abs(rcTmp.Y2 - .Y2) < m_lSnapWidth: lOffy = .Y2 - rcTmp.Y2
97:                                End Select
98:                            End If
99:                        End With
100:                   End If
101:               End If
102:           Next lc2
103:       End If
104:   Next lc1
    
    '== Apply offsets
    
108:    Call OffsetRect(rcWnd, lOffx, lOffy)
    
    '== Glueing (move child windows, if any)
    
112:    hDWP = BeginDeferWindowPos(1)
    
114:    For lc1 = 1 To m_lWndCount
115:        With m_uWndInfo(lc1)
            '-- Is parent our current window ?
116:            If (.hWndParent = hWnd And .Glue) Then
                '-- Move 'child' window
118:                lWId = pvWndGetInfoIndex(hWnd)
119:                With m_rcWnd(lc1)
120:                    Call DeferWindowPos(hDWP, m_uWndInfo(lc1).hWnd, 0, .X1 - (m_rcWnd(lWId).X1 - rcWnd.X1), .Y1 - (m_rcWnd(lWId).Y1 - rcWnd.Y1), 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOZORDER)
121:                End With
122:            End If
123:        End With
124:    Next lc1
    
126:    Call EndDeferWindowPos(hDWP)
    
    '== Store last cursor position
    
130:    m_ptLast = m_ptCurr
131:  Exit Sub
132:
Err:
134:  HandleError Err.Number, Err.Description, Erl & "|" & "clsMagneticWnd.pvMoveRect()"
End Sub

Private Sub pvCheckGlueing()
1:    On Error GoTo Err
2:     Dim lcMain As Long
3:     Dim lc1    As Long
4:     Dim lc2    As Long
5:     Dim lWId   As Long
    
    '-- Get all windows rectangles / Reset glueing
8:     For lc1 = 1 To m_lWndCount
9:        Call GetWindowRect(m_uWndInfo(lc1).hWnd, m_rcWnd(lc1))
10:        m_uWndInfo(lc1).Glue = False
11:    Next lc1
    
    '-- Check direct connection
13:    For lc1 = 1 To m_lWndCount
14:        If (m_uWndInfo(lc1).hWndParent) Then
            '-- Get parent window info index
16:            lWId = pvWndParentGetInfoIndex(m_uWndInfo(lc1).hWndParent)
            '-- Connected ?
18:            m_uWndInfo(lc1).Glue = pvWndsConnected(m_rcWnd(lWId), m_rcWnd(lc1))
19:        End If
20:    Next lc1
    
    '-- Check indirect connection
22:    For lcMain = 1 To m_lWndCount
        
24:        For lc1 = 1 To m_lWndCount
            
26:            If (m_uWndInfo(lc1).Glue) Then
                
28:                For lc2 = 1 To m_lWndCount
                
30:                    If (lc1 <> lc2) Then
                    
32:                        If (m_uWndInfo(lc1).hWndParent = m_uWndInfo(lc2).hWndParent) Then
                            '-- Connected ?
34:                            If (m_uWndInfo(lc2).Glue = False) Then
35:                                m_uWndInfo(lc2).Glue = pvWndsConnected(m_rcWnd(lc1), m_rcWnd(lc2))
36:                            End If
37:                        End If
38:                    End If
39:                Next lc2
40:            End If
41:        Next lc1
42:    Next lcMain
43:  Exit Sub
44:
Err:
46:  HandleError Err.Number, Err.Description, Erl & "|" & "clsMagneticWnd.pvCheckGlueing()"
End Sub

Private Function pvWndsConnected(rcWnd1 As RECT2, rcWnd2 As RECT2) As Boolean
1:  On Error GoTo Err
2:  Dim rcUnion As RECT2
  
    '-- Calc. union rectangle of windows
5:    Call UnionRect(rcUnion, rcWnd1, rcWnd2)
    
    '-- Bounding glue-rectangle
8:    If ((rcUnion.X2 - rcUnion.X1) <= (rcWnd1.X2 - rcWnd1.X1) + (rcWnd2.X2 - rcWnd2.X1) And _
        (rcUnion.Y2 - rcUnion.Y1) <= (rcWnd1.Y2 - rcWnd1.Y1) + (rcWnd2.Y2 - rcWnd2.Y1) _
         ) Then
        
        '-- Edge coincidences ?
11:        If (rcWnd1.X1 = rcWnd2.X1 Or rcWnd1.X1 = rcWnd2.X2 Or _
            rcWnd1.X2 = rcWnd2.X1 Or rcWnd1.X2 = rcWnd2.X2 Or _
            rcWnd1.Y1 = rcWnd2.Y1 Or rcWnd1.Y1 = rcWnd2.Y2 Or _
            rcWnd1.Y2 = rcWnd2.Y1 Or rcWnd1.Y2 = rcWnd2.Y2 _
            ) Then
            
12:            pvWndsConnected = True
13:        End If
14:    End If
15:  Exit Function
16:
Err:
18:    HandleError Err.Number, Err.Description, Erl & "|" & "clsMagneticWnd.pvWndsConnected()"
End Function

Private Function pvWndGetInfoIndex(ByVal hWnd As Long) As Long
1:  Dim LC As Long
    
3:    For LC = 1 To m_lWndCount
4:        If (m_uWndInfo(LC).hWnd = hWnd) Then
5:            pvWndGetInfoIndex = LC
6:            Exit For
7:        End If
8:    Next LC
End Function

Private Function pvWndParentGetInfoIndex(ByVal hWndParent As Long) As Long
1:  Dim LC As Long
    
3:    For LC = 1 To m_lWndCount
4:        If (m_uWndInfo(LC).hWnd = hWndParent) Then
5:            pvWndParentGetInfoIndex = LC
6:            Exit For
7:        End If
8:    Next LC
End Function

'========================================================================================
' Subclass code - The programmer may call any of the following Subclass_??? routines
'========================================================================================

Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
'Parameters:
'   lng_hWnd - The handle of the window for which the uMsg is to be added to the callback table
'   uMsg     - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
'   When     - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
  
7:     With sc_aSubData(zIdx(lng_hWnd))
8:        If (When And eMsgWhen.MSG_BEFORE) Then
9:            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
10:        End If
11:        If (When And eMsgWhen.MSG_AFTER) Then
12:            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
13:        End If
14:    End With
End Sub

Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Delete a message from the table of those that will invoke a callback.
'Parameters:
'   lng_hWnd - The handle of the window for which the uMsg is to be removed from the callback table
'   uMsg     - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
'   When     - Whether the msg is to be removed from the before, after or both callback tables
  
7:     With sc_aSubData(zIdx(lng_hWnd))
8:         If (When And eMsgWhen.MSG_BEFORE) Then
9:            Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
10:        End If
11:        If (When And eMsgWhen.MSG_AFTER) Then
12:            Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
13:        End If
14:    End With
End Sub

Private Function Subclass_InIDE() As Boolean
'Return whether we're running in the IDE.
2:    Debug.Assert zSetTrue(Subclass_InIDE)
End Function

Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
'Start subclassing the passed window handle
'Parameters:
'   lng_hWnd - The handle of the window to be subclassed
'Returns;
'   The sc_aSubData() index

7:   On Error GoTo Err

9:   Dim i                        As Long                       'Loop index
10:  Dim j                        As Long                       'Loop index
11:  Dim nSubIdx                  As Long                       'Subclass data index
12:  Dim sSubCode                 As String                     'Subclass code string
  
14:  Const GMEM_FIXED             As Long = 0                   'Fixed memory GlobalAlloc flag
15:  Const PAGE_EXECUTE_READWRITE As Long = &H40&               'Allow memory to execute without violating XP SP2 Data Execution Prevention
16:  Const PATCH_01               As Long = 18                  'Code buffer offset to the location of the relative address to EbMode
17:  Const PATCH_02               As Long = 68                  'Address of the previous WndProc
18:  Const PATCH_03               As Long = 78                  'Relative address of SetWindowsLong
19:  Const PATCH_06               As Long = 116                 'Address of the previous WndProc
20:  Const PATCH_07               As Long = 121                 'Relative address of CallWindowProc
21:  Const PATCH_0A               As Long = 186                 'Address of the owner object
22:  Const FUNC_CWP               As String = "CallWindowProcA" 'We use CallWindowProc to call the original WndProc
23:  Const FUNC_EBM               As String = "EbMode"          'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
24:  Const FUNC_SWL               As String = "SetWindowLongA"  'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
25:  Const MOD_USER               As String = "user32"          'Location of the SetWindowLongA & CallWindowProc functions
26:  Const MOD_VBA5               As String = "vba5"            'Location of the EbMode function if running VB5
27:  Const MOD_VBA6               As String = "vba6"            'Location of the EbMode function if running VB6

    'If it's the first time through here..
30:    If (sc_aBuf(1) = 0) Then

        'Build the hex pair subclass string
33:        sSubCode = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D0000005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D000000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E32F09C978078B450CF2AF75248D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF501CC3"
    
        'Convert the string from hex pairs to bytes and store in the machine code buffer
36:        i = 1
37:        Do While j < CODE_LEN
38:            j = j + 1
39:            sc_aBuf(j) = CByte("&H" & Mid$(sSubCode, i, 2))                       'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
40:            i = i + 2
41:        Loop                                                                      'Next pair of hex characters
    
        'Get API function addresses
44:        If (Subclass_InIDE) Then                                                  'If we're running in the VB IDE
45:            sc_aBuf(16) = &H90                                                    'Patch the code buffer to enable the IDE state code
46:            sc_aBuf(17) = &H90                                                    'Patch the code buffer to enable the IDE state code
47:            sc_pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                            'Get the address of EbMode in vba6.dll
48:            If (sc_pEbMode = 0) Then                                              'Found?
49:                sc_pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                        'VB5 perhaps
50:            End If
51:        End If
    
53:        Call zPatchVal(VarPtr(sc_aBuf(1)), PATCH_0A, ObjPtr(Me))                  'Patch the address of this object instance into the static machine code buffer
    
55:        sc_pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                   'Get the address of the CallWindowsProc function
56:        sc_pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                   'Get the address of the SetWindowLongA function
57:        ReDim sc_aSubData(0 To 0) As tSubData                                     'Create the first sc_aSubData element
    
59:      Else
60:        nSubIdx = zIdx(lng_hWnd, True)
61:        If (nSubIdx = -1) Then                                                    'If an sc_aSubData element isn't being re-cycled
62:            nSubIdx = UBound(sc_aSubData()) + 1                                   'Calculate the next element
63:            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                  'Create a new sc_aSubData element
64:        End If
    
66:        Subclass_Start = nSubIdx
67:    End If

69:    With sc_aSubData(nSubIdx)
        
71:        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                             'Allocate memory for the machine code WndProc
72:        Call VirtualProtect(ByVal .nAddrSub, CODE_LEN, PAGE_EXECUTE_READWRITE, i) 'Mark memory as executable
73:        Call RtlMoveMemory(ByVal .nAddrSub, sc_aBuf(1), CODE_LEN)                 'Copy the machine code from the static byte array to the code array in sc_aSubData
    
75:        .hWnd = lng_hWnd                                                          'Store the hWnd
76:        .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)                'Set our WndProc in place
    
78:        Call zPatchRel(.nAddrSub, PATCH_01, sc_pEbMode)                           'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
79:        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                           'Original WndProc address for CallWindowProc, call the original WndProc
80:        Call zPatchRel(.nAddrSub, PATCH_03, sc_pSWL)                              'Patch the relative address of the SetWindowLongA api function
81:        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                           'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
82:        Call zPatchRel(.nAddrSub, PATCH_07, sc_pCWP)                              'Patch the relative address of the CallWindowProc api function
83:    End With
84:  Exit Function
85:
Err:
87:  HandleError Err.Number, Err.Description, Erl & "|" & "clsMagneticWnd.Subclass_Start()"
End Function

Private Sub Subclass_StopAll()
'Stop all subclassing
  
3:  Dim i As Long
  
4:    i = UBound(sc_aSubData())                                                     'Get the upper bound of the subclass data array
5:    Do While i >= 0                                                               'Iterate through each element
6:        With sc_aSubData(i)
7:            If (.hWnd <> 0) Then                                                  'If not previously Subclass_Stop'd
8:                Call Subclass_Stop(.hWnd)                                         'Subclass_Stop
9:            End If
10:       End With
    
12:      i = i - 1                                                                 'Next element
13:   Loop
End Sub

Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
'Stop subclassing the passed window handle
'Parameters:
'   lng_hWnd - The handle of the window to stop being subclassed
  
4:     With sc_aSubData(zIdx(lng_hWnd))
5:         Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)                       'Restore the original WndProc
6:         Call zPatchVal(.nAddrSub, PATCH_05, 0)                                    'Patch the Table B entry count to ensure no further 'before' callbacks
7:         Call zPatchVal(.nAddrSub, PATCH_09, 0)                                    'Patch the Table A entry count to ensure no further 'after' callbacks
8:         Call GlobalFree(.nAddrSub)                                                'Release the machine code memory
9:         .hWnd = 0                                                                 'Mark the sc_aSubData element as available for re-use
10:        .nMsgCntB = 0                                                             'Clear the before table
11:        .nMsgCntA = 0                                                             'Clear the after table
12:        Erase .aMsgTblB                                                           'Erase the before table
13:        Erase .aMsgTblA                                                           'Erase the after table
14:    End With
End Sub

'----------------------------------------------------------------------------------------
'These z??? routines are exclusively called by the Subclass_??? routines.
'----------------------------------------------------------------------------------------

Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
'Worker sub for Subclass_AddMsg
2:    On Error GoTo Err
3:  Dim nEntry  As Long                                                             'Message table entry index
4:  Dim nOff1   As Long                                                             'Machine code buffer offset 1
5:  Dim nOff2   As Long                                                             'Machine code buffer offset 2
  
7:    If (uMsg = ALL_MESSAGES) Then                                                 'If all messages
8:        nMsgCnt = ALL_MESSAGES                                                    'Indicates that all messages will callback
9:      Else                                                                        'Else a specific message number
10:        Do While nEntry < nMsgCnt                                                 'For each existing entry. NB will skip if nMsgCnt = 0
11:            nEntry = nEntry + 1
        
13:            If (aMsgTbl(nEntry) = 0) Then                                         'This msg table slot is a deleted entry
14:                aMsgTbl(nEntry) = uMsg                                            'Re-use this entry
15:                Exit Sub                                                          'Bail
16:            ElseIf (aMsgTbl(nEntry) = uMsg) Then                                  'The msg is already in the table!
17:                Exit Sub                                                          'Bail
18:            End If
19:        Loop                                                                      'Next entry

21:        nMsgCnt = nMsgCnt + 1                                                     'New slot required, bump the table entry count
22:        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                              'Bump the size of the table.
23:        aMsgTbl(nMsgCnt) = uMsg                                                   'Store the message number in the table
24:    End If

26:    If (When = eMsgWhen.MSG_BEFORE) Then                                          'If before
27:        nOff1 = PATCH_04                                                          'Offset to the Before table
28:        nOff2 = PATCH_05                                                          'Offset to the Before table entry count
29:      Else                                                                        'Else after
30:        nOff1 = PATCH_08                                                          'Offset to the After table
31:        nOff2 = PATCH_09                                                          'Offset to the After table entry count
32:    End If

34:    If (uMsg <> ALL_MESSAGES) Then
35:        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                          'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
36:    End If
37:    Call zPatchVal(nAddr, nOff2, nMsgCnt)                                         'Patch the appropriate table entry count
38:  Exit Sub
39:
Err:
41:  HandleError Err.Number, Err.Description, Erl & "|" & "clsMagneticWnd.zAddMsg()"
End Sub

Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
'Return the memory address of the passed function in the passed dll
2:    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
3:    Debug.Assert zAddrFunc                                                        'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
'Worker sub for Subclass_DelMsg
  
3:     Dim nEntry As Long
  
5:     If (uMsg = ALL_MESSAGES) Then                                                 'If deleting all messages
6:         nMsgCnt = 0                                                               'Message count is now zero
7:         If When = eMsgWhen.MSG_BEFORE Then                                        'If before
8:             nEntry = PATCH_05                                                     'Patch the before table message count location
9:         Else                                                                    'Else after
10:            nEntry = PATCH_09                                                     'Patch the after table message count location
11:        End If
12:        Call zPatchVal(nAddr, nEntry, 0)                                          'Patch the table message count to zero
13:     Else                                                                        'Else deleteting a specific message
14:        Do While nEntry < nMsgCnt                                                 'For each table entry
15:            nEntry = nEntry + 1
16:            If (aMsgTbl(nEntry) = uMsg) Then                                      'If this entry is the message we wish to delete
17:                aMsgTbl(nEntry) = 0                                               'Mark the table slot as available
18:                Exit Do                                                           'Bail
19:            End If
20:        Loop                                                                      'Next entry
21:    End If
End Sub

Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
'Get the sc_aSubData() array index of the passed hWnd
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
  
4:    zIdx = UBound(sc_aSubData)
5:    Do While zIdx >= 0                                                            'Iterate through the existing sc_aSubData() elements
6:        With sc_aSubData(zIdx)
7:            If (.hWnd = lng_hWnd) Then                                            'If the hWnd of this element is the one we're looking for
8:                If (Not bAdd) Then                                                'If we're searching not adding
9:                    Exit Function                                                 'Found
10:                End If
11:            ElseIf (.hWnd = 0) Then                                               'If this an element marked for reuse.
12:                If (bAdd) Then                                                    'If we're adding
13:                    Exit Function                                                 'Re-use it
14:                End If
15:            End If
16:        End With
17:        zIdx = zIdx - 1                                                           'Decrement the index
18:    Loop
  
20:    If (Not bAdd) Then
21:        Debug.Assert False                                                        'hWnd not found, programmer error
22:    End If

'If we exit here, we're returning -1, no freed elements were found
End Function

Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
'Patch the machine code buffer at the indicated offset with the relative address to the target address.
2:    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
'Patch the machine code buffer at the indicated offset with the passed value
2:    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
'Worker function for Subclass_InIDE
2:    zSetTrue = True
3:    bValue = True
End Function

