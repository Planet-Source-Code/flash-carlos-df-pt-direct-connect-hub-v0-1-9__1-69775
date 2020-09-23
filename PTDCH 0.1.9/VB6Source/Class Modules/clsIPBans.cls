VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIPBans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Private m_lvwTemp   As ListItems
'Private m_lvwPerm   As ListItems
'Private m_lvwItem   As ListItem
Private m_colTemp   As Collection
Private m_colPerm   As Collection
Private m_objTB     As clsTempBan

Private Sub Class_Initialize()
    ''Link class to GUI
    'Set m_lvwTemp = frmHub.lvwTempIPBan.ListItems
    'Set m_lvwPerm = frmHub.lvwPermIPBan.ListItems
4:    Set m_colTemp = New Collection
5:    Set m_colPerm = New Collection
End Sub

Private Sub Class_Terminate()
    'Set m_lvwTemp = Nothing
    'Set m_lvwPerm = Nothing
3:    Set m_colTemp = Nothing
4:    Set m_colPerm = Nothing
End Sub

'Ban a new IP
'lngMinutes = -1 = Perm ban
'           > 0  = Temp ban
Public Sub Add(ByRef strIP As String, Optional ByRef lngMinutes As Long = -1)
1:    On Error GoTo Err
    
    'Select right ban type
4:    If lngMinutes = -1 Then
5:        If Not CheckPerm(strIP) Then
            'm_lvwPerm.Add , strIP, strIP
7:            m_colPerm.Add strIP, strIP
            
            'Run script event
10:            frmHub.SEvent_AddedPermBan strIP
11:        End If
12:    Else
        'If already banned, then replace old ban time with new
14:        If CheckTemp(strIP) Then
15:            m_colTemp(strIP).ExpDate = DateAdd("n", lngMinutes, CheckDate(m_colTemp(strIP).ExpDate))
16:        Else
17:            Set m_objTB = New clsTempBan
18:            m_objTB.IP = strIP
19:            m_objTB.ExpDate = DateAdd("n", lngMinutes, Now)
20:            m_colTemp.Add m_objTB, strIP
21:            Set m_objTB = Nothing
22:        End If
        
        'If CheckTemp(strIP) Then _
        '    m_lvwTemp(strIP).Tag = DateAdd("n", lngMinutes, Now) _
        'Else _
        '    m_lvwTemp.Add(, strIP, strIP).Tag = DateAdd("n", lngMinutes, Now)
28:    End If
    
30:    Exit Sub
    
32:
Err:
33:    HandleError Err.Number, Err.Description, Erl & "|" & "clsIPBans.Add(""" & strIP & """, " & lngMinutes & ")"
End Sub

'Remove an IP
'intType = 0 = Both
'        = 1 = Temp
'        = 2 = Perm
Public Sub Remove(ByRef strIP As String, Optional ByRef intType As Integer)
1:    On Error GoTo Err
    
    'Select which ban lists we should check/remove
    Select Case intType
        Case 0
4:            If CheckTemp(strIP) Then m_colTemp.Remove strIP 'm_lvwTemp.Remove strIP
5:            If CheckPerm(strIP) Then m_colPerm.Remove strIP 'm_lvwPerm.Remove strIP
        Case 1
6:            If CheckTemp(strIP) Then m_colTemp.Remove strIP 'm_lvwTemp.Remove strIP
        Case 2
7:            If CheckPerm(strIP) Then m_colPerm.Remove strIP 'm_lvwPerm.Remove strIP
8:    End Select
    
10:    Exit Sub
    
12:
Err:
13:    HandleError Err.Number, Err.Description, Erl & "|" & "clsIPBans.Remove(""" & strIP & """)"
End Sub

'Check if the ips are banned
'   0    = Not banned
'   -1   = Perm banned
'   Else = Temp banned (length in minutes)
Public Function Check(ByRef strIP As String) As Long
1:    On Error GoTo Err
    
3:    If CheckPerm(strIP) Then
4:        Check = -1
5:    ElseIf CheckTemp(strIP) Then
        'Check if the date has expired
7:        Check = DateDiff("n", Now, m_colTemp(strIP).ExpDate)
        'Check = DateDiff("n", Now, m_lvwTemp(strIP).Tag)
        
        'If the date is earlier than the current date, the temp ip ban
        'has expired
12:        If Check < 1 Then
13:            m_colTemp.Remove strIP
14:            Check = 0
15:        End If
16:    End If
    
18:    Exit Function
    
20:
Err:
21:    HandleError Err.Number, Err.Description, Erl & "|" & "clsIPBans.Check(""" & strIP & """)"
End Function

Private Function CheckPerm(ByRef strIP As String) As Boolean
1:    On Error GoTo Err
    
3:    m_colPerm.Item strIP
4:    CheckPerm = True
    'CheckPerm = ObjPtr(m_lvwPerm(strIP))
    
7:    Exit Function
8:
Err:
End Function

Private Function CheckTemp(ByRef strIP As String) As Boolean
1:    On Error GoTo Err
    
3:    m_colTemp.Item strIP
4:    CheckTemp = True
    'CheckTemp = ObjPtr(m_lvwTemp(strIP))
    
7:    Exit Function
8:
Err:
End Function

'Clear temp ban list
Public Sub ClearTemp()
1:    Set m_colTemp = New Collection
    'm_lvwTemp.Clear
End Sub

'Clear perm ban list
Public Sub ClearPerm()
1:    Set m_colPerm = New Collection
    'm_lvwPerm.Clear
End Sub

'Loop through listitems to get a list of temp banned ips (with expiration date)
Public Function TempList() As String
    'Dim lvwItem As ListItem
    
3:    On Error GoTo Err
    
    'Don't loop if there is nothing in the collection
6:    If m_colTemp.count Then
7:        For Each m_objTB In m_colTemp
8:            TempList = TempList & m_objTB.IP & " " & m_objTB.ExpDate & "|"
9:        Next
        
11:        Set m_objTB = Nothing
12:    End If
    
    'If m_lvwTemp.Count = 0 Then Exit Function
    
    'For Each lvwItem In m_lvwTemp
    '    TempList = TempList & lvwItem.Text & " " & lvwItem.Tag & "|"
    'Next
    
20:    Exit Function
    
22:
Err:
23:    HandleError Err.Number, Err.Description, Erl & "|" & "clsIPBans.TempList()"
End Function

'Loop through listitems to get a list of perm banned ips
Public Function PermList() As String
1:    Dim varLoop     As Variant
    'Dim lvwItem As ListItem
    
4:    On Error GoTo Err
    
    'Don't loop if there is nothing in the collection
7:    If m_colPerm.count Then
8:        For Each varLoop In m_colPerm
9:            PermList = PermList & varLoop & "|"
10:        Next
11:    End If
    
    'If m_lvwPerm.Count = 0 Then Exit Function
    
    'For Each lvwItem In m_lvwPerm
    '    PermList = PermList & lvwItem.Text & "|"
    'Next
    
19:    Exit Function
    
21:
Err:
22:    HandleError Err.Number, Err.Description, Erl & "|" & "clsIPBans.PermList()"
End Function

Public Property Get TempItems() As Collection
1:    Set TempItems = m_colTemp
End Property

Public Property Get PermItems() As Collection
1:    Set PermItems = m_colPerm
End Property

'Public Property Get TempItems() As ListItems
'    Set TempItems = m_lvwTemp
'End Property

'Public Property Get PermItems() As ListItem
'    Set PermItems = m_lvwPerm
'End Property

Private Function CheckDate(ByRef datCheck As Date) As Date
1:    If DateDiff("n", datCheck, Now) Then
2:        CheckDate = Now
3:    Else
4:        CheckDate = datCheck
5:    End If
End Function
