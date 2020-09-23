VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFailedReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_bytCount          As Byte
Private m_datLastAttempt    As Date

Private Sub Class_Initialize()
1:    m_datLastAttempt = Now
End Sub

Friend Function Check(ByRef curUser As clsUser) As Boolean
    'If an hour has passed, then ignore previous attempts (except current)
2:    If DateDiff("h", m_datLastAttempt, Now) Then
3:        m_bytCount = 1
4:    Else
        'Increment attempts / last attempt
6:        m_bytCount = m_bytCount + 1
7:    End If
    
    'Check if they have surpassed max number of attempts
10:    If m_bytCount >= g_objSettings.MaxPassAttempts Then
11:        curUser.Kick 60
12:        Check = True
13:    Else
14:        m_datLastAttempt = Now
15:    End If
End Function
