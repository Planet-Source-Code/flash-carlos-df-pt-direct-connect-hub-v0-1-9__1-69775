VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public DataFragment                 As String

Private m_strName                   As String
Private m_strIP                     As String
Private m_strMyINFO                 As String
Private m_strFakeMyINFO             As String
Private m_strLanguageID             As String
Private m_strSupports               As String

Private m_datConnectedSince         As Date

Private m_dblBytesShared            As Double
Private m_dblVersion                As Double

Private m_intIndex                  As Integer

Private m_blnQNL                    As Boolean
Private m_blnQuickList              As Boolean
Private m_blnOperator               As Boolean
Private m_blnVisible                As Boolean
Private m_blnAFK                    As Boolean
Private m_blnMute                   As Boolean
Private m_blnPassive                As Boolean
Private m_blnUserCommand            As Boolean
Private m_blnNoHello                As Boolean
Private m_blnNetInfo                As Boolean
Private m_blnZLine                  As Boolean
Private m_blnZPipe                 As Boolean

Private m_intClass                  As enuClass
Private m_intState                  As enuState
Private m_wskUser                   As Winsock

#If FLASHCHAT Then
    Private m_blnNullCharSeparator  As Boolean
#End If

Private m_blnChatOnly           As Boolean
    
Private m_bytMICount            As Byte
Private m_bytNLCount            As Byte
Private m_bytPSCount            As Byte
Private m_bytASCount            As Byte
Private m_lngMITick             As Long
Private m_lngNLTick             As Long
Private m_lngPSTick             As Long
Private m_lngASTick             As Long

Private m_bytCCount            As Byte
'Private m_bytGCount            As Byte
Private m_lngCTick             As Long
'svn 216
Private m_lngDropTick          As Long
'Private m_lngGTick             As Long

Private Sub Class_Initialize()
1:    m_datConnectedSince = Now
End Sub

Public Property Let sName(ByRef strName As String)
1:    m_strName = strName
End Property

Public Property Get sName() As String
1:    sName = m_strName
End Property

Public Property Get IP() As String
1:    IP = m_strIP
End Property

Public Property Let sMyInfoString(ByRef strData As String)

2:    m_strMyINFO = strData

4:    If g_objSettings.HideMyinfos Then
5:        m_strFakeMyINFO = "$MyINFO $All " & g_objRegExps.CaptureSubStr(strData, GETNICK) & " $ $" & g_objRegExps.CaptureSubStr(strData, GETCONTYPE) & g_objRegExps.CaptureSubStr(strData, GETSTATUS) & "$$" & g_objRegExps.CaptureDbl(strData, GETSHARESIZE) & "$"
6:    End If

End Property

Public Property Get sMyInfoString() As String
1:    sMyInfoString = m_strMyINFO
End Property

Public Property Let Supports(ByRef strData As String)
1:    m_strSupports = strData
End Property

Public Property Get Supports() As String
1:    Supports = m_strSupports
End Property

Public Property Get ConnectedSince() As Date
1:    ConnectedSince = m_datConnectedSince
End Property

Public Property Let iBytesShared(ByRef dblData As Double)
1:    m_dblBytesShared = dblData
End Property

Public Property Get iBytesShared() As Double
1:    iBytesShared = m_dblBytesShared
End Property

Public Property Let iVersion(ByRef dblData As Double)
1:    m_dblVersion = dblData
End Property

Public Property Get iVersion() As Double
1:    iVersion = m_dblVersion
End Property

Public Property Let iWinsockIndex(ByRef intData As Integer)
1:    m_intIndex = intData
End Property

Public Property Get iWinsockIndex() As Integer
1:    iWinsockIndex = m_intIndex
End Property

Public Property Let QNL(ByRef blnData As Boolean)
1:    m_blnQNL = blnData
End Property

Public Property Get QNL() As Boolean
1:    QNL = m_blnQNL
End Property

Public Property Let QuickList(ByRef blnData As Boolean)
1:    m_blnQuickList = blnData
End Property

Public Property Get QuickList() As Boolean
1:    QuickList = m_blnQuickList
End Property

Public Property Let NoHello(ByRef blnData As Boolean)
1:    m_blnNoHello = blnData
End Property

Public Property Get NoHello() As Boolean
1:    NoHello = m_blnNoHello
End Property

Public Property Let ZLine(ByRef blnData As Boolean)
1:    m_blnZLine = blnData
End Property

Public Property Get ZLine() As Boolean
1:    ZLine = m_blnZLine
End Property

Public Property Let ZPipe(ByRef blnData As Boolean)
1:    m_blnZPipe = blnData
End Property

Public Property Get ZPipe() As Boolean
1:    ZPipe = m_blnZPipe
End Property

Public Property Get bOperator() As Boolean
1:    bOperator = m_blnOperator
End Property

Public Property Get Visible() As Boolean
1:    Visible = m_blnVisible
End Property

Public Property Let isAFK(ByRef blnData As Boolean)
1:    m_blnAFK = blnData
End Property

Public Property Get isAFK() As Boolean
1:    isAFK = m_blnAFK
End Property

Public Property Let Mute(ByRef blnData As Boolean)
1:    m_blnMute = blnData
End Property

Public Property Get Mute() As Boolean
1:    Mute = m_blnMute
End Property

Public Property Let Passive(ByRef blnData As Boolean)
1:    m_blnPassive = blnData
End Property

Public Property Get Passive() As Boolean
1:    Passive = m_blnPassive
End Property

Public Property Let UserCommand(ByRef blnData As Boolean)
1:    m_blnUserCommand = blnData
End Property

Public Property Get UserCommand() As Boolean
1:    UserCommand = m_blnUserCommand
End Property

#If FLASHCHAT Then
    Public Property Let NullCharSeparator(ByRef blnData As Boolean)
1:        m_blnNullCharSeparator = blnData
    End Property
    
    Public Property Get NullCharSeparator() As Boolean
1:        NullCharSeparator = m_blnNullCharSeparator
    End Property
#End If

Public Property Let ChatOnly(ByRef blnData As Boolean)
1:    m_blnChatOnly = blnData
End Property
    
Public Property Get ChatOnly() As Boolean
1:    ChatOnly = m_blnChatOnly
End Property
    
Public Property Let NetInfo(ByRef blnData As Boolean)
1:    m_blnNetInfo = blnData
End Property

Public Property Get NetInfo() As Boolean
1:    NetInfo = m_blnNetInfo
End Property

Public Property Let Class(ByRef intData As enuClass)
1:    m_intClass = intData
2:    m_blnOperator = (intData >= Op)
    
    'Set visible status
    Select Case intData
        Case Invisible, InvisibleOp, InvisibleSuperOp, InvisibleAdmin
5:            m_blnVisible = False
        Case Else
6:            m_blnVisible = True
7:    End Select
End Property

Public Property Get Class() As enuClass
1:    Class = m_intClass
End Property

Public Property Let State(ByRef intData As enuState)
1:    m_intState = intData
End Property

Public Property Get State() As enuState
1:    State = m_intState
End Property

Public Property Set Winsock(ByRef wskData As Winsock)
1:    Set m_wskUser = wskData
2:    If ObjPtr(wskData) Then m_strIP = m_wskUser.RemoteHostIP
End Property

Public Property Get Winsock() As Winsock
1:    Set Winsock = m_wskUser
End Property

Public Property Let sMyInfoFakeString(ByRef strData As String)
1:    m_strFakeMyINFO = strData
End Property

Public Property Get sMyInfoFakeString() As String
1:    sMyInfoFakeString = m_strFakeMyINFO
End Property

Public Property Let sLanguageID(ByRef strData As String)
1:    m_strLanguageID = strData
End Property

Public Property Get sLanguageID() As String
1:    sLanguageID = m_strLanguageID
End Property

Friend Function FloodCheck(ByRef bytID As Byte) As Boolean
1:    Dim lngTick As Long
    
3:    lngTick = GetTickCount
    
    Select Case bytID
        Case 0 'Active search
5:            If Abs((lngTick - m_lngASTick)) > g_objSettings.FWInterval Then
6:                m_bytASCount = 1
7:                m_lngASTick = lngTick
8:            Else
9:                m_bytASCount = m_bytASCount + 1

11:                If m_bytASCount > g_objSettings.FWActiveSearch Then
12:                    Kick g_objSettings.FWBanLength
13:                    FloodCheck = True
14:                End If
15:            End If
        Case 1 'Passive search
16:            If Abs((lngTick - m_lngPSTick)) > g_objSettings.FWInterval Then
17:                m_bytPSCount = 1
18:                m_lngPSTick = lngTick
19:            Else
20:                m_bytPSCount = m_bytPSCount + 1
                
22:                If m_bytPSCount > g_objSettings.FWPassiveSearch Then
23:                    Kick g_objSettings.FWBanLength
24:                    FloodCheck = True
25:                End If
26:            End If
        Case 2 'MyINFO
27:            If Abs((lngTick - m_lngMITick)) > g_objSettings.FWInterval Then
28:                m_bytMICount = 1
29:                m_lngMITick = lngTick
30:            Else
31:                m_bytMICount = m_bytMICount + 1
                
33:                If m_bytMICount > g_objSettings.FWMyINFO Then
34:                    Kick g_objSettings.FWBanLength
35:                    FloodCheck = True
36:                End If
37:            End If
        Case 3 'GetNickList
38:            If Abs((lngTick - m_lngNLTick)) > g_objSettings.FWInterval Then
39:                m_bytNLCount = 1
40:                m_lngNLTick = lngTick
41:            Else
42:                m_bytNLCount = m_bytNLCount + 1
                
44:                If m_bytNLCount > g_objSettings.FWGetNickList Then
45:                    Kick g_objSettings.FWBanLength
46:                    FloodCheck = True
47:                End If
48:            End If
        Case 4 'Main Chat, PM
49:            If Abs((lngTick - m_lngCTick)) > g_objSettings.FWInterval Then
50:                m_bytCCount = 1
51:                m_lngCTick = lngTick
52:            Else
53:                m_bytCCount = m_bytCCount + 1
                
                'svn 216
56:                If Abs((lngTick - m_lngDropTick)) > g_objSettings.FWDropMsgInterval Then
57:                    m_lngDropTick = lngTick
58:                Else
59:                    SendChat g_objSettings.BotName, "your message was not sent, please send your messages a bit slower."
60:                    FloodCheck = True
61:                End If
                
                'g_objSettings.FWMainChat is setable only in xml and script...
64:                If m_bytCCount > g_objSettings.FWMainChat Then
65:                    Kick g_objSettings.FWBanLength
66:                    FloodCheck = True
67:                End If
68:            End If

70:    End Select
End Function

'Send any message to user (don't forget the |)
Public Sub SendData(ByRef strData As String)

2:    Dim objZLib         As clsZLib
    
4:    On Error Resume Next

    #If SVN Then
7:        If LenB(strData) Then
8:            g_objFileAccess.AppendFile G_LOGPATH, Now & " --> " & "TO: " & g_colUsers.ItemByIP(m_wskUser.RemoteHostIP).sName & " - " & m_wskUser.RemoteHostIP & " - " & strData
            'If m_blnZLine Then g_objFileAccess.AppendFile G_LOGPATH, Now & " --> " & "TO: " & g_colUsers.ItemByIP(m_wskUser.RemoteHostIP).sName & " - " & m_wskUser.RemoteHostIP & " - " & ZLineEsc(objZLib.CompressString(strData))
10:        End If
    #End If

13:    On Error GoTo Err
    
    #If FLASHCHAT Then
16:        If m_blnNullCharSeparator Then
17:            m_wskUser.SendData Replace(strData, "|", vbNullChar)
18:        Else
            Select Case True
                Case m_blnZLine
19:                    If LenB(strData) > g_objSettings.ZLINELENGHT Then
                        'seem like "$Z " is already obsolete
21:                        Set objZLib = New clsZLib
                        
23:                        m_wskUser.SendData "$Z " & ZLineEsc(objZLib.CompressString(strData)) & "|"
24:                    Else
25:                        m_wskUser.SendData strData
26:                    End If
                Case m_blnZPipe
                    'DC++ "arne" accepted Zline "ZPipe", need no escape or trailling pipe.
28:                    If LenB(strData) > g_objSettings.ZLINELENGHT Then
29:                        Set objZLib = New clsZLib
30:                        m_wskUser.SendData "$ZOn|" & objZLib.CompressString(strData)
31:                    Else
32:                        m_wskUser.SendData strData
33:                    End If
                Case Else
34:                    m_wskUser.SendData strData
35:            End Select
36:        End If
    #Else
        Select Case True
            Case m_blnZLine
38:                If LenB(strData) > g_objSettings.ZLINELENGHT Then
39:                    Set objZLib = New clsZLib
40:                    m_wskUser.SendData "$Z " & ZLineEsc(objZLib.CompressString(strData)) & "|"
41:                Else
42:                    m_wskUser.SendData strData
43:                End If
            Case m_blnZPipe
                'DC++ "arne" accepted Zline "ZPipe", need no escape or trailling pipe.
45:                If LenB(strData) > g_objSettings.ZLINELENGHT Then
46:                    Set objZLib = New clsZLib
47:                    m_wskUser.SendData "$ZOn|" & objZLib.CompressString(strData)
48:                Else
49:                    m_wskUser.SendData strData
50:                End If
            Case Else
51:                m_wskUser.SendData strData
52:        End Select
    #End If

55:    Set objZLib = Nothing

57:    Exit Sub

59:
Err:
'#If SVN Then
'    HandleError Err.Number, Err.Description, Erl & "|" & "clsUser_SendData() " & strData & "|" & m_strName
'#End If
63:    If ObjPtr(m_wskUser) Then frmHub.CloseSocket m_intIndex
End Sub

'Send a main chat message
Public Sub SendChat(ByRef strFrom As String, ByRef strMessage As String)
1:    On Error GoTo Err
    
    #If FLASHCHAT Then
4:        SendData "<" & strFrom & "> " & strMessage & "|"
    #Else
6:        m_wskUser.SendData "<" & strFrom & "> " & strMessage & "|"
    #End If
   
9:    Exit Sub
    
11:
Err:
12:    If ObjPtr(m_wskUser) Then frmHub.CloseSocket m_intIndex
End Sub

'Send a private message
Public Sub SendPrivate(ByRef strFrom As String, ByRef strMessage As String, Optional ByRef strOtherFrom As String)
1:    On Error GoTo Err
    
3:    If LenB(strOtherFrom) Then
        #If FLASHCHAT Then
5:            SendData "$To: " & m_strName & " From: " & strFrom & " $<" & strOtherFrom & "> " & strMessage & "|"
        #Else
7:            m_wskUser.SendData "$To: " & m_strName & " From: " & strFrom & " $<" & strOtherFrom & "> " & strMessage & "|"
        #End If
9:    Else
        #If FLASHCHAT Then
11:            SendData "$To: " & m_strName & " From: " & strFrom & " $<" & strFrom & "> " & strMessage & "|"
        #Else
13:            m_wskUser.SendData "$To: " & m_strName & " From: " & strFrom & " $<" & strFrom & "> " & strMessage & "|"
        #End If
15:    End If
    
17:    Exit Sub
    
19:
Err:
20:    If ObjPtr(m_wskUser) Then frmHub.CloseSocket m_intIndex
End Sub

'Redirect to strAddress - default is the default redirect address
Public Sub Redirect(Optional ByRef strAddress As String)
1:    On Error GoTo Err
    
3:    If LenB(strAddress) Then
4:        m_wskUser.SendData "$ForceMove " & strAddress & "|"
5:    Else
6:        frmHub.NextRedirect
7:        m_wskUser.SendData "$ForceMove " & g_objSettings.RedirectIP & "|"
8:    End If
 
10:   DoEvents

      'Show pupop notification ..
13:   If g_objSettings.PopUpUserRedirected Then _
              g_objFunctions.ShowBallon "PTDCH - " & g_objSettings.HubName, "Connection Redirected" & vbNewLine & "IP: " & m_wskUser.LocalIP, 0, True

16:   AddLog vbNewLine & "Connection Redirected" & vbNewLine & _
                         "IP: " & m_wskUser.LocalIP & vbNewLine & _
                         "Address: " & strAddress, 4, True
19:
Err:
    'HandleError Err.Number, Err.Description, Erl & "|" & "clsUser_Redirect() " & strAddress & "|" & g_objSettings.RedirectIP
22:    frmHub.CloseSocket m_intIndex
End Sub

'Disconnect them from the hub
Public Sub Disconnect()
1:    If ObjPtr(m_wskUser) Then frmHub.CloseSocket m_intIndex
End Sub

'Disconnect them and ban them by IP for lngMinutes
Public Sub Kick(Optional ByRef lngMinutes As Long)
1:    On Error Resume Next

2:    If lngMinutes Then
3:       g_objIPBans.Add m_strIP, lngMinutes
         'Show Ballon notification
5:       If g_objSettings.PopUpUserKick Then _
               g_objFunctions.ShowBallon "PTDCH - " & g_objSettings.HubName, _
                                         "IP Kicked Temporary" & vbNewLine & _
                                         "IP: " & m_strIP & vbNewLine & _
                                         "Time:" & lngMinutes, 1, True
10:      AddLog vbNewLine & "IP Kicked Temporary" & vbNewLine & _
                            "IP: " & m_strIP & vbNewLine & _
                            "Time:" & lngMinutes, 4, True
13:    Else
14:      g_objIPBans.Add m_strIP, g_objSettings.DefaultBanTime
         'Show Ballon notification
16:      If g_objSettings.PopUpUserKick Then _
                g_objFunctions.ShowBallon "PTDCH - " & g_objSettings.HubName, _
                                          "IP Kicked Permanent" & vbNewLine & _
                                          "IP: " & m_strIP, 1, True
20:      AddLog "IP Kicked Permanent" & vbNewLine & _
                "IP: " & m_strIP, 9, True
22:   End If
      
24:    frmHub.CloseSocket m_intIndex
End Sub
'Disconnect them and ban their IP permanently
Public Sub Ban()
1:    On Error Resume Next
2:    g_objIPBans.Add m_strIP
3:    frmHub.CloseSocket m_intIndex
End Sub

Public Function GetCoreMsgStr(ByRef strStringID As String) As String
'------------------------------------------------------------------
'Purpose:   Get proper core/reason messages
'           base on string key name and user language preference.
'
'Params:    strStringID :   Key of the string in the collection (a.e. "LoggedIn")
'           objUser :   object user object
'
'Return:    Return the proper string for each registered users, default to english strings if not.
'
'------------------------------------------------------------------
  
12:    Dim colLang As Collection
  
14:    On Error Resume Next
  
16:    If IsEmpty(m_strLanguageID) Then
        'this one is sure since it is(will be) hardcoded.
18:        Set colLang = g_colLanguages("en")
19:    Else
20:        Set colLang = g_colLanguages(m_strLanguageID)
  
22:        If IsEmpty(colLang) Then
            'this one is sure since it is(will be) hardcoded.
24:            Set colLang = g_colLanguages("en")
25:        End If
26:    End If
  
28:    GetCoreMsgStr = colLang.Item(strStringID)
  
30:    Set colLang = Nothing
  
End Function

Private Function ZLineEsc(ByRef strData As String) As String
'------------------------------------------------------------------
'Purpose:   Escaping DC protocol commands ending pipe "|" when sending data to a ZLine enabled clients
'
'Params:    strString :   string that need escaping due to ZLine
'
'Return:    Return the characters escaped string
'
'------------------------------------------------------------------
9:    Dim strTmp          As String

11:    On Error GoTo Err

    'escape \, warning! a \ in a pattern must be escaped with a \
14:    strTmp = g_objRegExps.REReplace(strData, "\\", "\\")
    'escape |, warning a | in a pattern must be escaped with a \
16:    strTmp = g_objRegExps.REReplace(strTmp, "\|", "\P")

18:    ZLineEsc = strTmp

20:    Exit Function

22:
Err:
23:    HandleError Err.Number, Err.Description, Erl & "|" & "clsUser_ZLineEsc(" & strData & ")"
End Function
