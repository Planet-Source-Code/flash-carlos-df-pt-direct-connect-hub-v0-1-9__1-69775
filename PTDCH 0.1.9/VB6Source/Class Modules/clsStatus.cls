VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

#If Status Then
    'Set the list box's horizontal scroll bars
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Const LB_SETHORIZONTALEXTENT = &H194

    #Const LIMIT_MSG = True
    
    #If LIMIT_MSG Then
        Private m_lngMsgCountMChat      As Long
        Private m_lngMsgCountPMLog      As Long
        Private m_lngMsgCountMiscLog    As Long
        Private Const c_iMLinesChat     As Integer = 350
        Private Const c_iMLinesPMLog    As Integer = 250
        Private Const c_iMLinesMiscLog  As Integer = 150
    #End If
    
    Private m_lvwUsers                  As ListItems
    Private m_lvwItem                   As ListItem
    Private m_lstMessages(2)            As ListBox
    Private m_frmHub                    As Form
    
    Private Sub Class_Initialize()
1:        Set m_lvwUsers = frmHub.lvwUsers.ListItems
2:        Set m_lstMessages(0) = frmHub.lstStatus(0)
3:        Set m_lstMessages(1) = frmHub.lstStatus(1)
4:        Set m_lstMessages(2) = frmHub.lstStatus(2)
5:        Set m_frmHub = frmHub
    End Sub
    
    Private Sub Class_Terminate()
1:        Set m_lvwUsers = Nothing
2:        Set m_lstMessages(0) = Nothing
3:        Set m_lstMessages(1) = Nothing
4:        Set m_lstMessages(2) = Nothing
5:        Set m_frmHub = Nothing
    End Sub
    
    'Adds a user to the listview
    Public Sub UAdd(ByRef objUser As clsUser)
1:        Dim strTemp       As String
          
3:        On Error GoTo Err
        
5:        strTemp = CStr(objUser.iWinsockIndex)
          
          'Add listitem with unknown icon
8:        Set m_lvwItem = m_lvwUsers.Add(, strTemp & "i", "(No name)")
        
          'Add extra info
11:        m_lvwItem.SubItems(1) = CStr(objUser.IP)
12:        m_lvwItem.SubItems(2) = strTemp
13:        m_lvwItem.SubItems(3) = CDate(objUser.ConnectedSince)
      
18:        Set m_lvwItem = Nothing
        
20:        Exit Sub
        
22:
Err:
        'HandleError Err.Number, Err.Description, Erl & "|" & "clsStatus.UAdd()"
    End Sub
    
    'Updates a users info
    Public Sub UUpdate(ByRef intIndex As Integer, ByRef strName As String)
1:        On Error GoTo Err
    
        'Must remember to add more stuff
        'Suggestions appreciated (ie add icons)
    
        'Update name
7:        m_lvwUsers(intIndex & "i").Text = strName
        
9:        Exit Sub
        
11:
Err:
        'HandleError Err.Number, Err.Description, Erl & "|" & "clsStatus.UUpdate()"
    End Sub
    
    'Removes a user from the listview
    Public Sub URemove(ByRef intIndex As Integer)
1:        On Error Resume Next
        
3:        m_lvwUsers.Remove intIndex & "i"
    End Sub
    
    'Adds a message to the message listbox
    Public Sub MAdd(ByRef strMessage As String) ', ByRef strIP As String, ByRef lngColour As Long)

       On Error Resume Next

       'PM Log
5:      If Left(strMessage, 4) = "$To:" Then
            #If LIMIT_MSG Then
7:             If m_lngMsgCountPMLog > c_iMLinesPMLog Then
8:                m_lstMessages(1).RemoveItem 0 'remove first line
9:                m_lngMsgCountPMLog = m_lstMessages(1).ListCount - 1 'set m_lngMsgCountPMLog count
10:             Else
11:                m_lngMsgCountPMLog = m_lngMsgCountPMLog + 1
12:             End If
            #End If
14:          m_lstMessages(1).AddItem ("[" & Now & "]" & strMessage)
             'set selected last line
16:          m_lstMessages(1).Selected(m_lstMessages(1).ListCount - 1) = True
17:          SetListboxHSbar 1
        
        'Main Chat Log
20:     ElseIf Left(strMessage, 1) <> "$" Then
            #If LIMIT_MSG Then
22:              If m_lngMsgCountMChat > c_iMLinesMiscLog Then
23:                 m_lstMessages(0).RemoveItem 0 'remove first line
24:                 m_lngMsgCountMChat = m_lstMessages(0).ListCount - 1 'set m_lngMsgCountMChat count
25:              Else
26:                 m_lngMsgCountMChat = m_lngMsgCountMChat + 1
27:              End If
            #End If
29:          m_lstMessages(0).AddItem (strMessage)
             'set selected last line
31:          m_lstMessages(0).Selected(m_lstMessages(0).ListCount - 1) = True
32:          SetListboxHSbar 0
       
        'Misc protocol log
35:     Else

37:          If EnabledMiscLog Then
                #If LIMIT_MSG Then
39:                  If m_lngMsgCountMiscLog > c_iMLinesMiscLog Then
40:                     m_lstMessages(2).RemoveItem 0 'remove first line
41:                     m_lngMsgCountMiscLog = m_lstMessages(2).ListCount - 1 'set m_lngMsgCountMiscLog count
42:                  Else
43:                     m_lngMsgCountMiscLog = m_lngMsgCountMiscLog + 1
44:                  End If
                #End If
46:             m_lstMessages(2).AddItem (strMessage)
                'set selected last line
48:             m_lstMessages(2).Selected(m_lstMessages(2).ListCount - 1) = True
49:             SetListboxHSbar 2
50:          End If
51:     End If
        
    End Sub
    
    'Clears all messages
    Public Sub MClear(iIndex As Integer)
1:    On Error GoTo Err
2:      m_lstMessages(iIndex).Clear
        #If LIMIT_MSG Then
            Select Case iIndex
               Case 0: m_lngMsgCountMChat = 0
               Case 1: m_lngMsgCountPMLog = 0
               Case 2: m_lngMsgCountMiscLog = 0
            End Select
        #End If
10:     SetListboxHSbar iIndex
11:     Exit Sub
12:
Err:
14:     HandleError Err.Number, Err.Description, Erl & "|" & "clsStatus.MClear(" & iIndex & ")"
    End Sub
    
    'Clears m_lvwUsers
    Public Sub UClear()
1:        On Error GoTo Err
        
3:        m_lvwUsers.Clear
4:        Exit Sub
5:
Err:
6:        HandleError Err.Number, Err.Description, Erl & "|" & "clsStatus.UClear()"
    End Sub
    
    Private Sub SetListboxHSbar(iIndex As Integer)
    ' Set the list box's horizontal extent so it
    ' can display its longest entry. This routine
    ' assumes the form is using the same font as
    ' the list box.
5:      On Error GoTo Err

7:      Dim i As Integer
8:      Dim new_len As Long
9:      Dim max_len As Long

11:        For i = 0 To m_lstMessages(iIndex).ListCount - 1
12:            new_len = 10 + m_frmHub.ScaleX(m_frmHub.TextWidth(m_lstMessages(iIndex).List(i)), m_frmHub.ScaleMode, vbPixels)
13:            If max_len < new_len Then max_len = new_len
14:        Next i

16:        SendMessage m_lstMessages(iIndex).hWnd, _
                    LB_SETHORIZONTALEXTENT, _
                         max_len, 0
        
20:        Exit Sub
    
21:
Err:
22:        HandleError Err.Number, Err.Description, Erl & "|" & "clsStatus.SetListboxHSbar(" & iIndex & ")"
    End Sub
    
    'This function will be for little memory..
    'It is not necessary to generate the log, if the listbox is not visible
    'this because this log use a lot of memory..
    Private Function EnabledMiscLog() As Boolean
1:        If m_frmHub.WindowState = vbMinimized Or _
             m_lstMessages(2).Visible = False Then _
               EnabledMiscLog = False _
          Else EnabledMiscLog = True
    End Function
#End If
