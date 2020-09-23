VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Collection" ,"clsUser"
Attribute VB_Ext_KEY = "Member0" ,"clsUser"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'Status window - Setting this value to true turns on the Status / Admin panel.
'                Must be set in the Properties dialog (just included here for clarity)
'#Const STATUS = True

Private m_lngOpCount        As Long
Private m_lngPeakUsers      As Long
Private m_lngPeakOps        As Long
Private m_dblTotalBytes     As Double
Private m_dblPeakBytes      As Double
Private m_strNickList       As String
Private m_strOpList         As String
Private m_colUsers          As Collection
Private m_colLoggingIn      As Collection
Private m_colNames          As Collection
Private m_colNLoggingIn     As Collection
Private m_objUser           As clsUser

#If Status Then
    Private m_lblTS         As Label
    Private m_lblCU         As Label
    Private m_lblCO         As Label
    Private m_lblPS         As Label
    Private m_lblPU         As Label
    Private m_lblPO         As Label
#End If

Private Sub Class_Initialize()
1:    Set m_colUsers = New Collection
2:    Set m_colLoggingIn = New Collection
3:    Set m_colNames = New Collection
4:    Set m_colNLoggingIn = New Collection

    #If Status Then
7:        Set m_lblCU = frmHub.lblHolder(55)
8:        Set m_lblCO = frmHub.lblHolder(56)
9:        Set m_lblTS = frmHub.lblHolder(57)
10:        Set m_lblPU = frmHub.lblHolder(58)
11:        Set m_lblPO = frmHub.lblHolder(59)
12:        Set m_lblPS = frmHub.lblHolder(60)
    #End If
End Sub

Private Sub Class_Terminate()
1:    Set m_colUsers = Nothing
2:    Set m_colLoggingIn = Nothing
3:    Set m_colNames = Nothing
4:    Set m_colNLoggingIn = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
1:    On Error Resume Next
    'This is needed to use the For Each...Next loop
3:    Set NewEnum = m_colUsers.[_NewEnum]
End Property

'Clears out collection (DOES NOT DISCONNECT USERS)
Friend Sub Clear()
1:    On Error GoTo Err

3:    Set m_colUsers = New Collection
4:    Set m_colLoggingIn = New Collection
5:    Set m_colNames = New Collection
    
7:    m_strNickList = vbNullString
8:    m_strOpList = vbNullString
9:    m_dblTotalBytes = 0
10:    m_lngOpCount = 0
    
    #If Status Then
13:        m_lblCU.Caption = "0"
14:        m_lblCO.Caption = "0"
15:        m_lblTS.Caption = "0 Bytes"
    #End If
    
18:    Exit Sub
    
20:
Err:
21:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHub.Clear()"
End Sub

'Add a user to the user collection
Friend Function Add(ByRef intIndex As Integer) As clsUser
1:    On Error GoTo Err
    
    'Create new class
4:    Set Add = New clsUser
5:    Add.iWinsockIndex = intIndex
    
    'Make sure this index isn't being used
8:    If Exists(intIndex) Then Remove intIndex
    
    'Add to collection
11:    m_colLoggingIn.Add Add, CStr(intIndex)
    
13:    Exit Function
    
15:
Err:
16:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHub.Add(" & intIndex & ")"
End Function

Friend Sub Remove(ByRef intIndex As Integer)
1:    Dim strTemp     As String
2:    Dim lngLen      As Long

4:    On Error GoTo Err

    'Refer to object
7:    Set m_objUser = ItemByWinsockIndex(intIndex)

9:    strTemp = m_objUser.sName
10:    lngLen = LenB(strTemp)

12:    If m_objUser.State = Logged_In Then
        'Remove from collection / update share size
14:        m_colUsers.Remove CStr(intIndex)

16:        If m_objUser.Visible Then
17:            m_dblTotalBytes = m_dblTotalBytes - m_objUser.iBytesShared

            #If Status Then
20:                m_lblTS.Caption = g_objFunctions.ShareSize(m_dblTotalBytes)
            #End If
22:        End If

        'Remove from nicklist / collection if needed
25:        If lngLen Then
26:            m_colNames.Remove strTemp

            'Update nicklist
29:            RemoveNL strTemp, m_objUser.bOperator

            #If Status Then
32:                m_lblCU.Caption = m_colUsers.count
            #End If

            'If operator remove from both lists
36:            If m_objUser.bOperator Then
                'Update op count
38:                m_lngOpCount = m_lngOpCount - 1

                #If Status Then
41:                     m_lblCO.Caption = m_lngOpCount
                #End If
43:            End If
44:        End If
45:    Else
        'Remove from collection / update share size
47:        m_colLoggingIn.Remove CStr(intIndex)
48:        If m_objUser.Visible Then m_dblTotalBytes = m_dblTotalBytes - m_objUser.iBytesShared

50:        On Error Resume Next
        'Remove from collection if needed
52:        If lngLen Then m_colNLoggingIn.Remove strTemp
53:    End If

    'Clear reference so object can be released
56:    Set m_objUser = Nothing

58:    Exit Sub

60:
Err:
61:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHub.Remove(" & intIndex & ")"
End Sub

'Check if user is online by winsock index
Public Function Exists(ByRef intIndex As Integer) As Long
1:    On Error Resume Next
    
3:    Exists = ObjPtr(m_colUsers(CStr(intIndex)))
    
5:    If Exists Then
6:        Exists = -1
7:    Else
8:        Exists = ObjPtr(m_colLoggingIn(CStr(intIndex)))
        
10:        If Exists Then
11:            Exists = 1
12:        Else
13:            Exists = 0
14:        End If
15:    End If
End Function

'Check if user is online by name
Public Function Online(ByRef strName As String) As Long
1:    On Error Resume Next
    
3:    Online = ObjPtr(m_colNames(strName))
    
5:    If Online Then
6:        Online = -1
7:    Else
8:        Online = ObjPtr(m_colNLoggingIn(strName))
9:        If Online Then Online = 1
10:    End If
End Function

'Find a user by winsock index
Public Function ItemByWinsockIndex(ByRef intIndex As Integer) As clsUser
1:    On Error Resume Next
    
3:    Set ItemByWinsockIndex = m_colUsers(CStr(intIndex))
4:    If ObjPtr(ItemByWinsockIndex) = 0 Then Set ItemByWinsockIndex = m_colLoggingIn(CStr(intIndex))
End Function

'Find a user by name
Public Function ItemByName(ByRef strName As String) As clsUser
1:    On Error Resume Next
    
3:    Set ItemByName = m_colNames(strName)
4:    If ObjPtr(ItemByName) = 0 Then Set ItemByName = m_colNLoggingIn(strName)
End Function

'Find a user by IP
Public Function ItemByIP(ByRef strIP As String) As clsUser
1:    On Error GoTo Err
    
3:    For Each ItemByIP In m_colUsers
4:        If ItemByIP.IP = strIP Then Exit Function
5:    Next
    
7:    For Each ItemByIP In m_colLoggingIn
8:        If ItemByIP.IP = strIP Then Exit Function
9:    Next
    
11:    Set ItemByIP = Nothing
    
13:    Exit Function
    
15:
Err:
16:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHub.ItemByIP(""" & strIP & """)"
End Function

'Get a random user object
Public Function ItemByRandom() As clsUser
1:    Dim lngCount As Long

3:    On Error GoTo Err
    
    'Set new seed number
6:    Randomize GetTickCount

    'Get number of connected users
9:    lngCount = m_colUsers.count

    'Make sure there are users in collection
12:    If lngCount Then _
        Set ItemByRandom = m_colUsers(CInt(lngCount * Rnd + 1))
    
15:    Exit Function
    
17:
Err:
18:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHub.ItemByRandom()"
End Function

'Check if a user is in the nicklist
Public Function CheckList(ByRef strName As String) As Boolean
1:    On Error Resume Next
    
3:    CheckList = (LeftB$(m_strNickList, LenB(strName) + 4) = strName & "$$")
4:    If Not CheckList Then CheckList = InStrB(1, m_strNickList, "$$" & strName & "$$")
End Function

'Appends name to nicklist
Public Sub AppendNL(ByRef strName As String, ByRef blnOperator As Boolean)
1:    On Error GoTo Err

3:    m_strNickList = m_strNickList & strName & "$$"
4:    If blnOperator Then m_strOpList = m_strOpList & strName & "$$"
    
6:    Exit Sub
    
8:
Err:
9:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHub.AppendNL(""" & strName & """, " & blnOperator & ")"
End Sub

'Removes name from nicklist
Public Sub RemoveNL(ByRef strName As String, ByRef blnOperator As Boolean)
1:    Dim lngLen As Long
    
3:    On Error GoTo Err
    
5:    lngLen = LenB(strName)

    'If found in nicklist, remove!
8:    If InStrB(1, m_strNickList, "$$" & strName & "$$") Then
9:        m_strNickList = Replace(m_strNickList, "$$" & strName & "$$", "$$")
10:    ElseIf LeftB$(m_strNickList, lngLen + 2) = strName & "$" Then
11:        m_strNickList = MidB$(m_strNickList, lngLen + 5)
12:    End If

    'Check oplist if necessary
15:    If blnOperator Then
16:        If InStrB(1, m_strOpList, "$$" & strName & "$$") Then
17:            m_strOpList = Replace(m_strOpList, "$$" & strName & "$$", "$$")
18:        ElseIf LeftB$(m_strOpList, lngLen + 2) = strName & "$" Then
19:            m_strOpList = MidB$(m_strOpList, lngLen + 5)
20:        End If
21:    End If
    
23:    Exit Sub
    
25:
Err:
26:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHub.RemoveNL(""" & strName & """, " & blnOperator & ")"
End Sub


'User count
Public Property Get count() As Long
1:    On Error Resume Next
2:    count = m_colUsers.count
End Property

'Peak user count
Public Property Get PeakUsers() As Long
1:    PeakUsers = m_lngPeakUsers
End Property

'Op count
Public Property Get OpCount() As Long
1:    OpCount = m_lngOpCount
End Property

'Peak op count
Public Property Get PeakOps() As Long
1:    PeakOps = m_lngPeakOps
End Property

'Total bytes
Public Property Let iTotalBytesShared(ByRef dblData As Double)
1:    On Error GoTo Err

3:    m_dblTotalBytes = dblData
    
    #If Status Then
6:        m_lblTS.Caption = g_objFunctions.ShareSize(dblData)
    #End If
    
    'Update peak shared if necessary
10:    If dblData > m_dblPeakBytes Then
11:        m_dblPeakBytes = dblData
        
        #If Status Then
14:            m_lblPS.Caption = g_objFunctions.ShareSize(dblData)
        #End If
16:    End If
    
18:    Exit Property
    
20:
Err:
21:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHub.iTotalBytesShared(" & dblData & ") (Let)"
End Property

Public Property Get iTotalBytesShared() As Double
1:    If m_dblTotalBytes < 0 Then ResetStats

3:    iTotalBytesShared = m_dblTotalBytes
End Property

'Peak bytes shared
Public Property Get iPeakBytesShared() As Double
1:    iPeakBytesShared = m_dblPeakBytes
End Property

'Nick list
Public Property Get NickList() As String
1:    NickList = m_strNickList
End Property

'Op nick list
Public Property Get OpList() As String
1:    OpList = m_strOpList
End Property

'Logging in collection
Public Property Get colLoggingIn() As Collection
1:    Set colLoggingIn = m_colLoggingIn
End Property

'Send data to all users (remember the |)
Public Sub SendToAll(ByRef strData As String)
1:    On Error GoTo Err
    'On Error Resume Next

4:    For Each m_objUser In m_colUsers
5:        m_objUser.SendData strData
6:    Next
    
8:    Set m_objUser = Nothing
    
10:    Exit Sub
    
12:
Err:
13:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHub.SendToAll(""" & strData & """)"
End Sub

'Send main chat message to all users
Public Sub SendChatToAll(ByRef strFrom As String, ByRef strMessage As String)
1:    Dim strTemp As String
    
3:    On Error GoTo Err
    
5:    strTemp = "<" & strFrom & "> " & strMessage & "|"
    
7:    For Each m_objUser In m_colUsers
8:        m_objUser.SendData strTemp
9:    Next
    
11:    Set m_objUser = Nothing
    
13:    Exit Sub
    
15:
Err:
16:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHub.SendChatToAll(""" & strFrom & """, """ & strMessage & """)"
End Sub

'Send a private message to all users
Public Sub SendPrivateToAll(ByRef strFrom As String, ByRef strMessage As String)
1:    Dim strTemp As String
    
3:    On Error GoTo Err
    
5:    strTemp = " From: " & strFrom & " $<" & strFrom & "> " & strMessage & "|"
    
7:    For Each m_objUser In m_colUsers
8:        m_objUser.SendData "$To: " & m_objUser.sName & strTemp
9:    Next
    
11:    Set m_objUser = Nothing
    
13:    Exit Sub
    
15:
Err:
16:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHub.SendPrivateToAll(""" & strFrom & """, """ & strMessage & """)"
End Sub


'Send data to ops (remember the |)
Public Sub SendToOps(ByRef strData As String)
1:    On Error GoTo Err
    
3:    For Each m_objUser In m_colUsers
4:        If m_objUser.bOperator Then m_objUser.SendData strData
5:    Next
    
7:    Set m_objUser = Nothing
    
9:    Exit Sub
    
11:
Err:
12:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHub.SendToOps(""" & strData & """)"
End Sub

'Send data to UnRegistereds (remember the |)
Public Sub SendToUnReg(ByRef strData As String)
1:    On Error GoTo Err
    
3:    For Each m_objUser In m_colUsers
4:        If m_objUser.Class < Registered Then m_objUser.SendData strData
5:    Next
    
7:    Set m_objUser = Nothing
    
9:    Exit Sub
    
11:
Err:
12:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHub.SendToOps(""" & strData & """)"
End Sub
'Send main chat message to all ops
Public Sub SendChatToOps(ByRef strFrom As String, ByRef strMessage As String)
1:    Dim strTemp As String
    
3:    On Error GoTo Err
    
5:    strTemp = "<" & strFrom & "> " & strMessage & "|"
    
7:    For Each m_objUser In m_colUsers
8:        If m_objUser.bOperator Then m_objUser.SendData strTemp
9:    Next
    
11:    Set m_objUser = Nothing
    
13:    Exit Sub
    
15:
Err:
16:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHub.SendChatToOps(""" & strFrom & """, """ & strMessage & """)"
End Sub

'Send main chat message to all ops
Public Sub SendChatToUnReg(ByRef strFrom As String, ByRef strMessage As String)
1:    Dim strTemp As String
    
3:    On Error GoTo Err
    
5:    strTemp = "<" & strFrom & "> " & strMessage & "|"
    
7:    For Each m_objUser In m_colUsers
8:        If m_objUser.Class < Registered Then m_objUser.SendData strTemp
9:    Next
    
11:    Set m_objUser = Nothing
    
13:    Exit Sub
    
15:
Err:
16:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHub.SendChatToOps(""" & strFrom & """, """ & strMessage & """)"
End Sub

'Send a private message to all ops
Public Sub SendPrivateToOps(ByRef strFrom As String, ByRef strMessage As String)
1:    Dim strTemp As String
    
3:    On Error GoTo Err
    
5:    strTemp = " From: " & strFrom & " $<" & strFrom & "> " & strMessage & "|"
    
7:    For Each m_objUser In m_colUsers
8:        If m_objUser.bOperator Then m_objUser.SendData "$To: " & m_objUser.sName & strTemp
9:    Next
    
11:    Set m_objUser = Nothing
    
13:    Exit Sub
    
15:
Err:
16:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHub.SendPrivateToOps(""" & strFrom & """, """ & strMessage & """)"
End Sub

'Send a private message to all UnRegistered
Public Sub SendPrivateToUnReg(ByRef strFrom As String, ByRef strMessage As String)
1:    Dim strTemp As String
    
3:    On Error GoTo Err
    
5:    strTemp = " From: " & strFrom & " $<" & strFrom & "> " & strMessage & "|"
    
7:    For Each m_objUser In m_colUsers
8:        If m_objUser.Class < Registered Then m_objUser.SendData "$To: " & m_objUser.sName & strTemp
9:    Next
    
11:    Set m_objUser = Nothing
    
13:    Exit Sub
    
15:
Err:
16:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHub.SendPrivateToUnReg(""" & strFrom & """, """ & strMessage & """)"
End Sub
'Send to non quicklist clients
Public Sub SendToNQ(ByRef strData As String)
1:    On Error GoTo Err
    
3:    For Each m_objUser In m_colUsers
4:        If Not m_objUser.QuickList Then m_objUser.SendData strData
5:    Next
    
7:    Set m_objUser = Nothing
    
9:    Exit Sub
    
11:
Err:
12:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHub.SendToNQ(""" & strData & """)"
End Sub

'Send to users who are not in away mode
Public Sub SendToNA(ByRef strData As String)
1:    On Error GoTo Err
    
3:    For Each m_objUser In m_colUsers
4:        If Not m_objUser.isAFK Then m_objUser.SendData strData
5:    Next
    
7:    Exit Sub
    
9:
Err:
10:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHub.SendToNA(""" & strData & """)"
End Sub

'Redirect all users to an address (default is one in settings)
Public Sub RedirectAll(Optional ByRef strAddress As String)
1:    Dim strTemp As String

3:    On Error GoTo Err

    'Choose right address
6:    If LenB(strAddress) Then
7:        strTemp = "$ForceMove " & strAddress & "|"
8:    Else
9:        strTemp = "$ForceMove " & g_objSettings.RedirectIP & "|"
10:    End If

12:    On Error Resume Next
    'Redirect all logged in users
14:    For Each m_objUser In m_colUsers
15:        m_objUser.SendData strTemp
16:        DoEvents
17:        m_objUser.Disconnect
18:    Next

    'Redirect all logging in users
21:    For Each m_objUser In m_colLoggingIn
22:        m_objUser.SendData strTemp
23:        DoEvents
24:        m_objUser.Disconnect
25:    Next

27:    On Error GoTo Err

29:    Set m_objUser = Nothing
    
31:    Exit Sub

33:
Err:
34:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHub.RedirectAll(""" & strAddress & """)"
End Sub

'Redirect all non ops to an address (default is one in settings)
Public Sub RedirectNonOps(Optional ByRef strAddress As String)
1:    Dim strTemp As String

3:    On Error GoTo Err

    'Choose right address
6:    If LenB(strAddress) Then
7:        strTemp = "$ForceMove " & strAddress & "|"
8:    Else
9:        strTemp = "$ForceMove " & g_objSettings.RedirectIP & "|"
10:    End If

12:    On Error Resume Next
    'Redirect all logged in non ops
14:    For Each m_objUser In m_colUsers
15:        If Not m_objUser.bOperator Then
16:            m_objUser.SendData strTemp
17:            DoEvents
18:            m_objUser.Disconnect
19:        End If
20:    Next

    'Redirect all logging in users (none can be ops yet)
23:    For Each m_objUser In m_colLoggingIn
24:        m_objUser.SendData strTemp
25:        DoEvents
26:        m_objUser.Disconnect
27:    Next

29:    On Error GoTo Err

31:    Set m_objUser = Nothing

33:    Exit Sub

35:
Err:
36:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHub.RedirectNonOps(""" & strAddress & """)"
End Sub

'Resets the statistics (basically recalculates shared bytes / op count
'and makes the current values the new peaks)
Public Sub ResetStats()
1:    On Error GoTo Err
    
    'Reset them to zero
4:    m_lngOpCount = 0
5:    m_dblTotalBytes = 0
    
    'Rebuild current by going through the logged in collection
8:    For Each m_objUser In m_colUsers
9:        If m_objUser.bOperator Then m_lngOpCount = m_lngOpCount + 1
10:        m_dblTotalBytes = m_dblTotalBytes + m_objUser.iBytesShared
11:    Next
    
    'Set peaks to current
14:    m_lngPeakUsers = m_colUsers.count
15:    m_lngPeakOps = m_lngOpCount
16:    m_dblPeakBytes = m_dblTotalBytes
    
    'If using status version, update interface
    #If Status Then
20:        m_lblCU.Caption = m_lngPeakUsers
21:        m_lblCO.Caption = m_lngOpCount
22:        m_lblTS.Caption = g_objFunctions.ShareSize(m_dblTotalBytes)
23:        m_lblPU.Caption = m_lngPeakUsers
24:        m_lblPO.Caption = m_lngOpCount
25:        m_lblPS.Caption = m_lblTS.Caption
    #End If
    
28:    Exit Sub
    
30:
Err:
31:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHub.RebuildStats()"
End Sub

'Hidden to scripts - checks for users who have been logging in for an extended
'                    period (longer than 5 minutes)
Friend Sub CheckExtendedLogIn()
1:    On Error GoTo Err

3:    For Each m_objUser In m_colLoggingIn
4:        If DateDiff("n", m_objUser.ConnectedSince, Now) > 4 Then frmHub.CloseSocket m_objUser.iWinsockIndex
5:    Next
    
7:    Set m_objUser = Nothing
    
9:    Exit Sub
    
11:
Err:
12:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHub.CheckExtendedLogIn()"
End Sub

'Hidden to scripts - adds user to name-indexed collection
Friend Sub UpdateName(ByRef objUser As clsUser)
1:    Dim strName As String

3:    On Error Resume Next
    
5:    strName = objUser.sName
    
    'Make sure name is removed if in use
8:    m_colNLoggingIn.Remove strName
    
    'Add to collection
11:    m_colNLoggingIn.Add objUser, strName
    
    #If Status Then
        'Update user list
15:        g_objStatus.UUpdate objUser.iWinsockIndex, strName
    #End If
End Sub

'Hidden to scripts - moves user object to logged in collection from logging
'                    in collection, updates nick/oplist and total bytes
Friend Sub UpdateLogIn(ByRef objUser As clsUser)
1:    Dim strIndex        As String
2:    Dim strParts        As String
3:    Dim lngCount        As Long
4:    Dim strTMyinfosStr  As String

6:    On Error GoTo Err

    'Make sure the winsock index actually exists in the
    'logging in collection
10:    If Not Exists(objUser.iWinsockIndex) = 1 Then Exit Sub
    
12:    strIndex = CStr(objUser.iWinsockIndex)

    'Set state to logged in
15:    objUser.State = Logged_In

    'Switch collections
18:    m_colLoggingIn.Remove strIndex
19:    m_colUsers.Add objUser, strIndex
    
21:    lngCount = m_colUsers.count
    
    #If Status Then
24:        m_lblCU.Caption = lngCount
    #End If
    
    'Remove from connection attempts / hammering collection
    'If removed we won't know if they are hammering...
    'frmHub.UpdateConnectAttempt objUser.Winsock, True
    
    'Update peak users
32:    If lngCount > m_lngPeakUsers Then
33:        m_lngPeakUsers = lngCount
        
        #If Status Then
36:            m_lblPU.Caption = lngCount
        #End If
38:    End If
    
40:    strIndex = objUser.sName
    
    'Swap objects again
43:    On Error Resume Next
44:    m_colNLoggingIn.Remove strIndex
45:    On Error GoTo Err
46:    m_colNames.Add objUser, strIndex
    
    'Update nick list
49:    If objUser.Visible Then
50:        m_strNickList = m_strNickList & strIndex & "$$"
        
        'Prepare data to send to users
53:        If objUser.bOperator Then
            'Update op count / list
55:            m_lngOpCount = m_lngOpCount + 1
56:            m_strOpList = m_strOpList & strIndex & "$$"
            
            #If Status Then
59:                m_lblCO.Caption = m_lngOpCount
            #End If
            
            'Update peak op count
63:            If m_lngOpCount > m_lngPeakOps Then
64:                m_lngPeakOps = m_lngOpCount
                
                #If Status Then
67:                    m_lblPO.Caption = m_lngOpCount
                #End If
69:            End If
        
71:            strParts = "$Hello " & strIndex & "|" & objUser.sMyInfoString & "|$OpList " & strIndex & "$$|"
72:            strIndex = objUser.sMyInfoString & "|$OpList " & strIndex & "$$|"
73:        Else
74:            strParts = "$Hello " & strIndex & "|" & objUser.sMyInfoString & "|"
75:            strIndex = objUser.sMyInfoString & "|"
76:        End If
77:    Else
78:        If objUser.bOperator Then
            'Update op count / list
80:            m_lngOpCount = m_lngOpCount + 1
            
            #If Status Then
83:                m_lblCO.Caption = m_lngOpCount
            #End If
            
            'Update peak op count
87:            If m_lngOpCount > m_lngPeakOps Then
88:                m_lngPeakOps = m_lngOpCount
                
                #If Status Then
91:                    m_lblPO.Caption = m_lngOpCount
                #End If
93:            End If
94:        End If
95:    End If
                
    'If they have a queued nicklist, send it
98:    If objUser.QNL Then
        'Loop through user collection and send message if visible
        'If not, then only send a MyINFO stream
101:        If objUser.Visible Then
102:            If Not objUser.NoHello Then objUser.SendData "$NickList " & m_strNickList & "|"
        
104:            lngCount = ObjPtr(objUser)
        
106:            For Each m_objUser In m_colUsers
                'ghost fix
108:                If objUser.State = Disconnected Then Exit Sub
109:                If Not lngCount = ObjPtr(m_objUser) Then
110:                    If g_objSettings.HideMyinfos Then
111:                        If m_objUser.Class < g_objSettings.MinMyinfoFakeCls Then
112:                            If m_objUser.NoHello Then
113:                                If objUser.bOperator Then
114:                                    m_objUser.SendData objUser.sMyInfoFakeString & "|" & "$OpList " & objUser.sName & "$$|"
115:                                Else
116:                                    m_objUser.SendData objUser.sMyInfoFakeString & "|"
117:                                End If
118:                            Else
119:                                If objUser.bOperator Then
120:                                    m_objUser.SendData "$Hello " & objUser.sName & "|" & objUser.sMyInfoFakeString & "|" & "$OpList " & objUser.sName & "$$|"
121:                                Else
122:                                    m_objUser.SendData "$Hello " & objUser.sName & "|" & objUser.sMyInfoFakeString & "|"
123:                                End If
124:                            End If
125:                        Else
126:                            If m_objUser.NoHello Then
127:                                m_objUser.SendData strIndex
128:                            Else
129:                                m_objUser.SendData strParts
130:                            End If
131:                        End If
132:                    Else
133:                        If m_objUser.NoHello Then
134:                                m_objUser.SendData strIndex
135:                            Else
136:                                m_objUser.SendData strParts
137:                            End If
138:                    End If
139:                End If

                'Build MyINFO stream
142:                If m_objUser.Visible Then
143:                    If g_objSettings.HideMyinfos Then
144:                        If objUser.Class < g_objSettings.MinMyinfoFakeCls Then
145:                            strTMyinfosStr = strTMyinfosStr & m_objUser.sMyInfoFakeString & "|"
146:                        Else
147:                            strTMyinfosStr = strTMyinfosStr & m_objUser.sMyInfoString & "|"
148:                        End If
149:                    Else
150:                        strTMyinfosStr = strTMyinfosStr & m_objUser.sMyInfoString & "|"
151:                    End If
152:                End If
153:            Next
            'Send MyINFO stream and oplist
155:            objUser.SendData strTMyinfosStr & "$OpList " & m_strOpList & "|"

157:        Else
158:            If Not objUser.NoHello Then objUser.SendData "$NickList " & m_strNickList & strIndex & "$$|"

            'Send own MyINFO string
161:            If g_objSettings.HideMyinfos Then
162:                If objUser.Class < g_objSettings.MinMyinfoFakeCls Then
163:                    objUser.SendData objUser.sMyInfoFakeString & "|"
164:                Else
165:                    objUser.SendData objUser.sMyInfoString & "|"
166:                End If
167:            Else
168:                objUser.SendData objUser.sMyInfoString & "|"
169:            End If

171:            lngCount = ObjPtr(objUser)

            'Build MyINFO stream
174:            For Each m_objUser In m_colUsers
175:                If Not lngCount = ObjPtr(m_objUser) Then
176:                    If m_objUser.Visible Then
177:                        If g_objSettings.HideMyinfos Then
178:                            If objUser.Class < g_objSettings.MinMyinfoFakeCls Then
179:                                strTMyinfosStr = strTMyinfosStr & m_objUser.sMyInfoFakeString & "|"
180:                            Else
181:                                strTMyinfosStr = strTMyinfosStr & m_objUser.sMyInfoString & "|"
182:                            End If
183:                        Else
184:                            strTMyinfosStr = strTMyinfosStr & m_objUser.sMyInfoString & "|"
185:                        End If
186:                    End If
187:                End If
188:            Next

190:            If objUser.bOperator Then
191:                strTMyinfosStr = strTMyinfosStr & "$OpList " & m_strOpList & strIndex & "$$|"
192:            Else
193:                strTMyinfosStr = strTMyinfosStr & "$OpList " & m_strOpList & "|"
194:            End If

            'Send MyINFO stream and oplist to user
197:            objUser.SendData strTMyinfosStr

199:        End If

        'Send Bot MyINFOs
202:        frmHub.UpdateBots objUser
203:    Else
204:        If objUser.Visible Then
205:            lngCount = ObjPtr(objUser)
        
            'Loop through user collection and send message
208:            For Each m_objUser In m_colUsers
                'ghost fix
210:                If objUser.State = Disconnected Then Exit Sub
211:                If Not lngCount = ObjPtr(m_objUser) Then
212:                    If g_objSettings.HideMyinfos Then
213:                        If m_objUser.Class < g_objSettings.MinMyinfoFakeCls Then
214:                            If objUser.bOperator Then
215:                                If m_objUser.NoHello Then
216:                                    m_objUser.SendData objUser.sMyInfoFakeString & "|" & "$OpList " & objUser.sName & "$$|"
217:                                Else
218:                                    m_objUser.SendData "$Hello " & objUser.sName & "|" & objUser.sMyInfoFakeString & "|" & "$OpList " & objUser.sName & "$$|"
219:                                End If
220:                            Else
221:                                If m_objUser.NoHello Then
222:                                    m_objUser.SendData objUser.sMyInfoFakeString & "|"
223:                                Else
224:                                    m_objUser.SendData "$Hello " & objUser.sName & "|" & objUser.sMyInfoFakeString & "|"
225:                                End If
226:                            End If
227:                        Else
228:                            If m_objUser.NoHello Then
229:                                m_objUser.SendData strIndex
230:                            Else
231:                                m_objUser.SendData strParts
232:                            End If
233:                        End If
234:                    Else
235:                        If m_objUser.NoHello Then
236:                            m_objUser.SendData strIndex
237:                        Else
238:                            m_objUser.SendData strParts
239:                        End If
240:                    End If
241:                End If
242:            Next
243:        End If
244:    End If
    
246:    Set m_objUser = Nothing
    
    'Send on join message if needed
    Select Case g_objSettings.SendJoinMsg
        Case 1
            'If g_objSettings.ParseMotd Then
            '    objUser.SendPrivate g_objSettings.BotName, g_objFunctions.MotdEscape(objUser, g_objSettings.JoinMsg)
            'Else
252:                objUser.SendPrivate g_objSettings.BotName, g_objSettings.JoinMsg
            'End If
        Case 2
            'If g_objSettings.ParseMotd Then
            '    objUser.SendChat g_objSettings.BotName, g_objFunctions.MotdEscape(objUser, g_objSettings.JoinMsg)
            'Else
257:                objUser.SendChat g_objSettings.BotName, g_objSettings.JoinMsg
            'End If
259:    End Select
    
261:    Exit Sub
    
263:
Err:
264:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHub.UpdateLogIn(" & DebugUser(objUser) & ")"
End Sub
