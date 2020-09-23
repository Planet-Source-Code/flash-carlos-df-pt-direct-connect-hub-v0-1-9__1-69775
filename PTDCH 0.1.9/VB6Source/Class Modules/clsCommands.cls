VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Collection" ,"clsUser"
Attribute VB_Ext_KEY = "Member0" ,"clsUser"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private m_lvwItems      As ListItems
Private m_colCommands   As Collection

Private Sub Class_Initialize()
1:    Set m_lvwItems = frmHub.lvwCommands.ListItems
2:    Set m_colCommands = New Collection
End Sub

Private Sub Class_Terminate()
1:    Set m_lvwItems = Nothing
2:    Set m_colCommands = Nothing
End Sub

'Add a command
Public Function Add(ByRef intID As Integer, ByRef strName As String, strDescription As String, ByRef intClass As enuClass, ByRef blnEnabled As Boolean) As Integer
'Scheduler comment: edit to set listview defaults. Ex: TrigEna default = False
' add optionnal parrams if want to be able to set a trigger when adding a command ; ))

4:    Dim objCommand  As clsCommand
5:    Dim objCmd      As clsCommand
6:    Dim lvwItem     As ListItem
7:    Dim intRandID   As Double
    
9:    On Error GoTo Err

11:    If intID = 0 Then
12:        Set objCmd = New clsCommand
13:
TryAgain:
        'Get random ID
15:      Randomize GetTickCount
16:      intRandID = Int(32767 * Rnd)
        'Reserve ID < 51 for compatibility with older versions
18:        If intRandID < 51 Then GoTo TryAgain
        
         'check if ID already exist.
21:      For Each objCmd In m_colCommands
22:         If intRandID = objCmd.ID Then GoTo TryAgain
23:      Next
            
25:      Set objCmd = Nothing
26:      intID = intRandID
27:    End If
    
    'Don't add if it already exists
30:    If Exists(strName) Then
        'Scheduler, need edit here.
        'Check if Optionnal params are available first.
33:        Edit strName, vbNullString, strDescription, intClass, blnEnabled
34:    Else
        'Scheduler, need edit here.
        'Check if Optionnal params are available first.

        'Create objects
39:        Set lvwItem = m_lvwItems.Add(, strName, strName)
40:        Set objCommand = New clsCommand
        
        'Update listview
43:        lvwItem.SubItems(1) = g_objFunctions.ClassName(intClass)
44:        lvwItem.SubItems(2) = CStr(blnEnabled)
        
        'Update collection
47:        objCommand.Name = strName
48:        objCommand.Description = strDescription
49:        objCommand.ID = intID
50:        objCommand.Enabled = blnEnabled
51:        objCommand.Class = intClass
        
53:        m_colCommands.Add objCommand, strName
54:    End If

56:    Add = intID
    
58:    Exit Function
    
60:
Err:
61:    HandleError Err.Number, Err.Description, Erl & "|" & "clsCommands.Add(" & intID & ", " & strName & ", " & strDescription & ", " & intClass & ", " & blnEnabled & ")"
End Function

'Edit a command
Public Sub Edit(ByRef strOldName As String, ByRef strNewName As String, ByRef strDescription As String, ByRef intClass As enuClass, ByRef blnEnabled As Boolean)
'Scheduler, need edit in here.
'Check if scheduler params are available or use default.

4:    Dim objCommand  As clsCommand
5:    Dim lvwItem     As ListItem
    
7:    On Error GoTo Err
    
    'Only edit if it exists
10:    If Exists(strOldName) Then
        'If the old name doesn't equal the new name, we must update the keys
        Select Case strNewName
            Case vbNullString, strOldName
                'Reference objects
13:                Set objCommand = m_colCommands(strOldName)
14:                Set lvwItem = m_lvwItems(strOldName)
            Case Else
                'Update collections/keys and reference objects
16:                Set objCommand = m_colCommands(strOldName)
17:                Remove strOldName
18:                m_colCommands.Add objCommand, strNewName
19:                Set lvwItem = m_lvwItems.Add(, strNewName, strNewName)
20:                objCommand.Name = strNewName
21:        End Select

        'Update class
24:        objCommand.Enabled = blnEnabled
25:        objCommand.Class = intClass
26:        objCommand.Description = strDescription
    
        'Update listview
29:        lvwItem.SubItems(1) = g_objFunctions.ClassName(intClass)
30:        lvwItem.SubItems(2) = CStr(blnEnabled)
31:    End If

33:    Exit Sub
    
35:
Err:
36:    HandleError Err.Number, Err.Description, Erl & "|" & "clsCommands.Edit(" & strOldName & ", " & strNewName & ", " & strDescription & ", " & intClass & ", " & blnEnabled & ")"
End Sub

'Remove a command
Public Sub Remove(ByRef strKey As String)
1:    On Error Resume Next
    
3:    m_lvwItems.Remove strKey
4:    m_colCommands.Remove strKey
End Sub

'Clear out commands
Public Sub Clear()
1:    m_lvwItems.Clear
2:    Set m_colCommands = New Collection
End Sub

'Retrieve command object
Public Function Item(ByRef strKey As String) As clsCommand
Attribute Item.VB_UserMemId = 0
1:    On Error Resume Next
    
3:    Set Item = m_colCommands(strKey)
End Function

'Find out if a command is in the collection
Public Function Exists(ByRef strKey As String) As Boolean
1:    On Error GoTo Err
    
3:    Exists = ObjPtr(m_colCommands(strKey))
    
5:    Exit Function
    
7:
Err:
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
1:    On Error Resume Next
    
    'This is needed to use the For Each...Next loop
4:    Set NewEnum = m_colCommands.[_NewEnum]
End Property

Public Sub Execute(ByRef objUser As clsUser, ByRef strTrigger As String, ByRef blnMainChat As Boolean)
1:    frmHub.ProcessTrigger objUser, strTrigger, blnMainChat
End Sub

