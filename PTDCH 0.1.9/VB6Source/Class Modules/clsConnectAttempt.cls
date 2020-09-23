VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConnectAttempt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_strIP             As String
Private m_datLastAttempt    As Date
Private m_bytCount          As Byte

Private Sub Class_Initialize()
1:    m_datLastAttempt = Now
End Sub

Friend Property Let IP(ByRef strData As String)
1:    m_strIP = strData
End Property

Friend Property Get IP() As String
1:    IP = m_strIP
End Property

Friend Property Get LastAttempt() As Date
1:    LastAttempt = m_datLastAttempt
End Property

Friend Function Check(ByRef wskUser As Winsock) As Boolean

2:    On Error GoTo Err
    'Check if 10 seconds has gone by since last attempt
    'If so, then reset connection attempts back to 1
5:    If DateDiff("s", m_datLastAttempt, Now) > 10 Then
6:        m_bytCount = 1
7:        m_datLastAttempt = Now
8:    Else
        'Increment attempt count
10:        m_bytCount = m_bytCount + 1

        'If the user tries to connect more than 2 times (in 10 seconds), then try redirect
        'more than 6 times, disconnect / ban them for 120 minutes.
        '100 is the shortess we can use to be safe...(password guessing/hammering can be keept up to ~80 mins)
        Select Case m_bytCount
            Case 1, 2, 3, 4, 5
                'Tolerance treshole
            Case 6
                'Redirect elsewhere attempt
17:                wskUser.SendData "<" & g_objSettings.BotName & "> You are being redirected (by PTDCH Anti Hammering).|<" & g_objSettings.BotName & "> You are being redirected to DEVNULL.|$ForceMove " & g_objSettings.HammeringRd & "|"

19:                DoEvents
20:                wskUser.Close

22:                Check = True
            Case Else
23:                g_objIPBans.Add m_strIP, 120
24:                wskUser.SendData "<" & g_objSettings.BotName & "> You are being banned for 30 minutes (by PTDCH Anti Hammering).|"
25:                DoEvents
26:                wskUser.Close

28:                Check = True
29:        End Select
        
31:        m_datLastAttempt = Now
32:    End If

34:    Exit Function

36:
Err:
37:    HandleError Err.Number, Err.Description, Erl & "|" & "clsConnectAttempt.Check(" & m_bytCount & "|" & m_strIP & ")"
End Function
