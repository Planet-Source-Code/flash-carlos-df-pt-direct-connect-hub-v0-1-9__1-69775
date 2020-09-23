Attribute VB_Name = "mYSubclass"
Option Explicit

'************************************************************************
' SSubTmr object
' Copyright © 1998-1999 Steve McMahon for vbAccelerator
' Mod by fLaSh for PT DC Hub
'************************************************************************

' The implementation of the Subclassing part of the SSubTmr object.
' Use this module + clsYISubclass.Cls to replace dependency on the DLL.

' declares:
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Const GWL_WNDPROC = (-4)
Private Const WM_DESTROY = &H2

' SubTimer is independent of VBCore, so it hard codes error handling

Public Enum EErrorWindowProc
    eeBaseWindowProc = 13080 ' WindowProc
    eeCantSubclass           ' Can't subclass window
    eeAlreadyAttached        ' Message already handled by another class
    eeInvalidWindow          ' Invalid window
    eeNoExternalWindow       ' Can't modify external window
End Enum

Private m_iCurrentMessage As Long
Private m_iProcOld As Long
Private m_f As Long

Public Property Get CurrentMessage() As Long
1:  CurrentMessage = m_iCurrentMessage
End Property

Private Sub ErrRaise(e As Long)
   On Error GoTo Err
   
2:   Dim sText As String, sSource As String
   
4:   If e > 1000 Then
5:        sSource = App.EXEName & ".WindowProc"
          Select Case e
            Case eeCantSubclass
8:               sText = "Can't subclass window"
            Case eeAlreadyAttached
10:               sText = "Message already handled by another class"
            Case eeInvalidWindow
12:               sText = "Invalid window"
            Case eeNoExternalWindow
14:               sText = "Can't modify external window"
          End Select
16:       Err.Raise e Or vbObjectError, sSource, sText
17:   Else
      ' Raise standard Visual Basic error
19:   Err.Raise e, sSource
20:   End If
   
22:   Exit Sub
23:
Err:
24:   HandleError Err.Number, Err.Description, Erl & "|mYSubclass.ErrRaise()"
End Sub

Private Property Get MessageCount(ByVal hWnd As Long) As Long
    On Error GoTo Err
    
    Dim sName As String
4:    sName = "C" & hWnd
5:    MessageCount = GetProp(hWnd, sName)
    
7:  Exit Property

Err:
10: HandleError Err.Number, Err.Description, Erl & "|mYSubclass.MessageCount()"
End Property

Private Property Let MessageCount(ByVal hWnd As Long, ByVal count As Long)
    On Error GoTo Err
    
    Dim sName As String
4:    m_f = 1
5:    sName = "C" & hWnd
6:    m_f = SetProp(hWnd, sName, count)
7:    If (count = 0) Then
8:      RemoveProp hWnd, sName
9:    End If
'   logMessage "Changed message count for " & Hex(hwnd) & " to " & count

12:  Exit Property

Err:
14:  HandleError Err.Number, Err.Description, Erl & "|mYSubclass.MessageCount()"
End Property

Private Property Get OldWindowProc(ByVal hWnd As Long) As Long
    On Error GoTo Err
    
    Dim sName As String
4:    sName = hWnd
5:    OldWindowProc = GetProp(hWnd, sName)

7:  Exit Property

Err:
10:  HandleError Err.Number, Err.Description, Erl & "|mYSubclass.OldWindowProc()"
End Property
    
Private Property Let OldWindowProc(ByVal hWnd As Long, ByVal lPtr As Long)
   On Error GoTo Err
   
   Dim sName As String
4:   m_f = 1
5:   sName = hWnd
6:   m_f = SetProp(hWnd, sName, lPtr)
7:   If (lPtr = 0) Then
8:      RemoveProp hWnd, sName
9:   End If
'   logMessage "Changed Window Proc for " & Hex(hwnd) & " to " & Hex(lPtr)

12: Exit Property

Err:
15:  HandleError Err.Number, Err.Description, Erl & "|mYSubclass.OldWindowProc()"
End Property

Private Property Get MessageClassCount(ByVal hWnd As Long, ByVal iMsg As Long) As Long
    On Error GoTo Err
    
    Dim sName As String
4:    sName = hWnd & "#" & iMsg & "C"
5:    MessageClassCount = GetProp(hWnd, sName)
    
7:    Exit Property

Err:
10:   HandleError Err.Number, Err.Description, Erl & "|mYSubclass.MessageClassCount()"
End Property

Private Property Let MessageClassCount(ByVal hWnd As Long, ByVal iMsg As Long, ByVal count As Long)
    On Error GoTo Err
    
    Dim sName As String
4:    sName = hWnd & "#" & iMsg & "C"
5:    m_f = SetProp(hWnd, sName, count)
6:    If (count = 0) Then
7:       RemoveProp hWnd, sName
8:    End If
'   logMessage "Changed message count for " & Hex(hwnd) & " Message " & iMsg & " to " & count
    
11:    Exit Property

Err:
13:    HandleError Err.Number, Err.Description, Erl & "|mYSubclass.MessageClassCount()"
End Property

Private Property Get MessageClass(ByVal hWnd As Long, ByVal iMsg As Long, ByVal Index As Long) As Long
   On Error GoTo Err
    
3:   Dim sName As String
4:   sName = hWnd & "#" & iMsg & "#" & Index
5:   MessageClass = GetProp(hWnd, sName)
   
7:   Exit Property

Err:
10:  HandleError Err.Number, Err.Description, Erl & "|mYSubclass.MessageClass()"
End Property
    
Private Property Let MessageClass(ByVal hWnd As Long, ByVal iMsg As Long, ByVal Index As Long, ByVal classPtr As Long)
   On Error GoTo Err
   
3:   Dim sName As String
4:   sName = hWnd & "#" & iMsg & "#" & Index
5:   m_f = SetProp(hWnd, sName, classPtr)
6:   If (classPtr = 0) Then
7:      RemoveProp hWnd, sName
8:   End If
'   logMessage "Changed message class for " & Hex(hwnd) & " Message " & iMsg & " Index " & index & " to " & Hex(classPtr)
   
11:  Exit Property

Err:
14:  HandleError Err.Number, Err.Description, Erl & "|mYSubclass.MessageClass()"
End Property

Sub AttachMessage( _
      iwp As clsYISubclass, _
      ByVal hWnd As Long, _
      ByVal iMsg As Long _
   )

    Dim procOld As Long
    Dim msgCount As Long
    Dim msgClassCount As Long
    Dim msgClass As Long
    
    On Error GoTo Err

   ' --------------------------------------------------------------------
   ' 1) Validate window
   ' --------------------------------------------------------------------
12:   If IsWindow(hWnd) = False Then
13:      ErrRaise eeInvalidWindow
14:      Exit Sub
15:   End If
   
17:   If IsWindowLocal(hWnd) = False Then
18:      ErrRaise eeNoExternalWindow
19:      Exit Sub
20:   End If

   ' --------------------------------------------------------------------
   ' 2) Check if this class is already attached for this message:
   ' --------------------------------------------------------------------
25:   msgClassCount = MessageClassCount(hWnd, iMsg)
26:   If (msgClassCount > 0) Then
27:      For msgClass = 1 To msgClassCount
28:         If (MessageClass(hWnd, iMsg, msgClass) = ObjPtr(iwp)) Then
29:            ErrRaise eeAlreadyAttached
30:            Exit Sub
31:         End If
32:      Next msgClass
33:   End If

   ' --------------------------------------------------------------------
   ' 3) Associate this class with this message for this window:
   ' --------------------------------------------------------------------
38:   MessageClassCount(hWnd, iMsg) = MessageClassCount(hWnd, iMsg) + 1
39:   If (m_f = 0) Then
      ' Failed, out of memory:
41:      ErrRaise 5
42:      Exit Sub
43:   End If
   
   ' --------------------------------------------------------------------
   ' 4) Associate the class pointer:
   ' --------------------------------------------------------------------
48:   MessageClass(hWnd, iMsg, MessageClassCount(hWnd, iMsg)) = ObjPtr(iwp)
49:   If (m_f = 0) Then
      ' Failed, out of memory:
51:      MessageClassCount(hWnd, iMsg) = MessageClassCount(hWnd, iMsg) - 1
52:      ErrRaise 5
53:      Exit Sub
54:   End If

   ' --------------------------------------------------------------------
   ' 5) Get the message count
   ' --------------------------------------------------------------------
59:   msgCount = MessageCount(hWnd)
60:   If msgCount = 0 Then
      
      ' Subclass window by installing window procedure
63:      procOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
64:      If procOld = 0 Then
         ' remove class:
66:         MessageClass(hWnd, iMsg, MessageClassCount(hWnd, iMsg)) = 0
         ' remove class count:
68:         MessageClassCount(hWnd, iMsg) = MessageClassCount(hWnd, iMsg) - 1
         
70:         ErrRaise eeCantSubclass
71:         Exit Sub
72:      End If
      
      ' Associate old procedure with handle
75:      OldWindowProc(hWnd) = procOld
76:      If m_f = 0 Then
         ' SPM: Failed to VBSetProp, windows properties database problem.
         ' Has to be out of memory.
         
         ' Put the old window proc back again:
81:         SetWindowLong hWnd, GWL_WNDPROC, procOld
         ' remove class:
83:         MessageClass(hWnd, iMsg, MessageClassCount(hWnd, iMsg)) = 0
         ' remove class count:
85:         MessageClassCount(hWnd, iMsg) = MessageClassCount(hWnd, iMsg) - 1
         
         ' Raise an error:
88:         ErrRaise 5
89:         Exit Sub
90:      End If
91:   End If
   
      
   ' Count this message
95:   MessageCount(hWnd) = MessageCount(hWnd) + 1
96:   If m_f = 0 Then
      ' SPM: Failed to set prop, windows properties database problem.
      ' Has to be out of memory
      
      ' remove class:
101:      MessageClass(hWnd, iMsg, MessageClassCount(hWnd, iMsg)) = 0
      ' remove class count contribution:
103:      MessageClassCount(hWnd, iMsg) = MessageClassCount(hWnd, iMsg) - 1
      
      ' If we haven't any messages on this window then remove the subclass:
106:      If (MessageCount(hWnd) = 0) Then
         ' put old window proc back again:
108:         procOld = OldWindowProc(hWnd)
109:         If Not (procOld = 0) Then
110:            SetWindowLong hWnd, GWL_WNDPROC, procOld
111:            OldWindowProc(hWnd) = 0
112:         End If
113:      End If
      
      ' Raise the error:
116:      ErrRaise 5
117:      Exit Sub
118:   End If
       
120:   Exit Sub
121:
Err:
122:   HandleError Err.Number, Err.Description, Erl & "|mYSubclass.AttachMessage()"
End Sub

Sub DetachMessage( _
      iwp As clsYISubclass, _
      ByVal hWnd As Long, _
      ByVal iMsg As Long _
   )
    Dim msgClassCount As Long
    Dim msgClass As Long
    Dim msgClassIndex As Long
    Dim msgCount As Long
    Dim procOld As Long
    
    On Error GoTo Err
    
   ' --------------------------------------------------------------------
   ' 1) Validate window
   ' --------------------------------------------------------------------
12:   If IsWindow(hWnd) = False Then
      ' for compatibility with the old version, we don't
      ' raise a message:
      ' ErrRaise eeInvalidWindow
16:      Exit Sub
17:   End If
   
19:   If IsWindowLocal(hWnd) = False Then
      ' for compatibility with the old version, we don't
      ' raise a message:
      ' ErrRaise eeNoExternalWindow
23:      Exit Sub
24:   End If
    
   ' --------------------------------------------------------------------
   ' 2) Check if this message is attached for this class:
   ' --------------------------------------------------------------------
29:   msgClassCount = MessageClassCount(hWnd, iMsg)
30:   If (msgClassCount > 0) Then
31:      msgClassIndex = 0
32:      For msgClass = 1 To msgClassCount
33:         If (MessageClass(hWnd, iMsg, msgClass) = ObjPtr(iwp)) Then
34:            msgClassIndex = msgClass
35:            Exit For
36:         End If
37:      Next msgClass
      
39:      If (msgClassIndex = 0) Then
         ' fail silently
41:         Exit Sub
42:      Else
         ' remove this message class:
         
         ' a) Anything above this index has to be shifted up:
46:         For msgClass = msgClassIndex To msgClassCount - 1
47:            MessageClass(hWnd, iMsg, msgClass) = MessageClass(hWnd, iMsg, msgClass + 1)
48:         Next msgClass
         
         ' b) The message class at the end can be removed:
50:         MessageClass(hWnd, iMsg, msgClassCount) = 0
         
         ' c) Reduce the message class count:
53:         MessageClassCount(hWnd, iMsg) = MessageClassCount(hWnd, iMsg) - 1
         
55:      End If
      
57:   Else
       ' fail silently
59:      Exit Sub
60:   End If
   
   ' ---------------------------------------------------------------------
   ' 3) Reduce the message count:
   ' ---------------------------------------------------------------------
65:   msgCount = MessageCount(hWnd)
66:   If (msgCount = 1) Then
      ' remove the subclass:
68:      procOld = OldWindowProc(hWnd)
69:      If Not (procOld = 0) Then
         ' Unsubclass by reassigning old window procedure
71:         Call SetWindowLong(hWnd, GWL_WNDPROC, procOld)
72:      End If
      ' remove the old window proc:
74:      OldWindowProc(hWnd) = 0
75:   End If
76:   MessageCount(hWnd) = MessageCount(hWnd) - 1
   
78:   Exit Sub
79:
Err:
80:   HandleError Err.Number, Err.Description, Erl & "|mYSubclass.DetachMessage()"
End Sub

Private Function WindowProc( _
      ByVal hWnd As Long, _
      ByVal iMsg As Long, _
      ByVal wParam As Long, _
      ByVal lParam As Long _
   ) As Long
    
    Dim procOld As Long
    Dim msgClassCount As Long
    Dim bCalled As Boolean
    Dim pSubClass As Long
    Dim iwp As clsYISubclass
    Dim iwpT As clsYISubclass
    Dim iIndex As Long
    Dim bDestroy As Boolean
    
    On Error GoTo Err
    
   ' Get the old procedure from the window
11:  procOld = OldWindowProc(hWnd)
15:   Debug.Assert procOld <> 0
    
17:   If (procOld = 0) Then
      ' we can't work, we're not subclassed properly.
19:      Exit Function
20:   End If
    
    ' SPM - in this version I am allowing more than one class to
    ' make a subclass to the same hWnd and Msg.  Why am I doing
    ' this?  Well say the class in question is a control, and it
    ' wants to subclass its container.  In this case, we want
    ' all instances of the control on the form to receive the
    ' form notification message.
     
    ' Get the number of instances for this msg/hwnd:
30:   bCalled = False
   
32:   If (MessageClassCount(hWnd, iMsg) > 0) Then
33:      iIndex = MessageClassCount(hWnd, iMsg)
      
35:      Do While (iIndex >= 1)
36:         pSubClass = MessageClass(hWnd, iMsg, iIndex)
         
38:         If (pSubClass = 0) Then
               ' Not handled by this instance
40:         Else
               ' Turn pointer into a reference:
42:            CopyMemory iwpT, pSubClass, 4
43:            Set iwp = iwpT
44:            CopyMemory iwpT, 0&, 4
            
               ' Store the current message, so the client can check it:
47:            m_iCurrentMessage = iMsg
            
49:            With iwp
                  ' Preprocess (only checked first time around):
51:               If (iIndex = 1) Then
52:                  If (.MsgResponse = emrPreprocess) Then
53:                     If Not (bCalled) Then
54:                        WindowProc = CallWindowProc(procOld, hWnd, iMsg, _
                                                  wParam, ByVal lParam)
56:                        bCalled = True
57:                     End If
58:                  End If
59:               End If
                 ' Consume (this message is always passed to all control
                 ' instances regardless of whether any single one of them
                 ' requests to consume it):
63:               WindowProc = .WindowProc(hWnd, iMsg, wParam, ByVal lParam)
64:            End With
65:         End If
         
67:         iIndex = iIndex - 1
68:      Loop
      
         ' PostProcess (only check this the last time around):
71:      If Not (iwp Is Nothing) And Not (procOld = 0) Then
72:          If iwp.MsgResponse = emrPostProcess Then
73:             If Not (bCalled) Then
74:                WindowProc = CallWindowProc(procOld, hWnd, iMsg, _
                                          wParam, ByVal lParam)
76:                bCalled = True
77:             End If
78:          End If
79:      End If
            
81:   Else
         ' Not handled:
83:      If (iMsg = WM_DESTROY) Then
           ' If WM_DESTROY isn't handled already, we should
           ' clear up any subclass
86:         pClearUp hWnd
87:         WindowProc = CallWindowProc(procOld, hWnd, iMsg, _
                                    wParam, ByVal lParam)
         
90:      Else
91:         WindowProc = CallWindowProc(procOld, hWnd, iMsg, _
                                    wParam, ByVal lParam)
93:      End If
94:   End If
    
96:   Exit Function
97:
Err:
99:    HandleError Err.Number, Err.Description, Erl & "|mYSubclass.WindowProc()"
End Function
 
Public Function CallOldWindowProc( _
      ByVal hWnd As Long, _
      ByVal iMsg As Long, _
      ByVal wParam As Long, _
      ByVal lParam As Long _
   ) As Long
   
    On Error GoTo Err
    
4:    Dim iProcOld As Long
5:    iProcOld = OldWindowProc(hWnd)
    
7:    If Not (iProcOld = 0) Then
8:      CallOldWindowProc = CallWindowProc(iProcOld, hWnd, iMsg, wParam, lParam)
9:    End If
    
11:    Exit Function

Err:
14:    HandleError Err.Number, Err.Description, Erl & "|mYSubclass.CallOldWindowProc()"
End Function

Function IsWindowLocal(ByVal hWnd As Long) As Boolean
    On Error GoTo Err
    
3:    Dim idWnd As Long
4:    Call GetWindowThreadProcessId(hWnd, idWnd)
5:    IsWindowLocal = (idWnd = GetCurrentProcessId())
    
7:    Exit Function
Err:
9:    HandleError Err.Number, Err.Description, Erl & "|mYSubclass.IsWindowLocal()"
End Function

Private Sub logMessage(ByVal sMsg As String)
1:   'Debug.Print sMsg
End Sub

Private Sub pClearUp(ByVal hWnd As Long)
1:    Dim msgCount As Long
2:    Dim procOld As Long
    
4:    On Error GoTo Err
     
     ' this is only called if you haven't explicitly cleared up
     ' your subclass from the caller.  You will get a minor
     ' resource leak as it does not clear up any message
     ' specific properties.
10:   msgCount = MessageCount(hWnd)
11:   If (msgCount > 0) Then
         ' remove the subclass:
13:      procOld = OldWindowProc(hWnd)
14:      If Not (procOld = 0) Then
            ' Unsubclass by reassigning old window procedure
16:         Call SetWindowLong(hWnd, GWL_WNDPROC, procOld)
17:      End If
         ' remove the old window proc:
19:      OldWindowProc(hWnd) = 0
20:      MessageCount(hWnd) = 0
21:   End If
   
23:   Exit Sub

Err:
26:   HandleError Err.Number, Err.Description, Erl & "|mYSubclass.pClearUp()"
End Sub
