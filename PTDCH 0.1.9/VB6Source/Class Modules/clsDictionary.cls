VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'General Library
'-------------------------------
'
'  Copyright (C) Andrew Osmond
'  E-Mail: tasburrfoot@users.sourceforge.net

Option Explicit

Private m_blnObject     As Boolean
Private m_intCompare    As VbCompareMethod
Private m_colItems      As Collection
Private m_colKeys       As Collection

Private Sub Class_Initialize()
1:    Set m_colItems = New Collection
2:    Set m_colKeys = New Collection
End Sub

Private Sub Class_Terminate()
1:    Set m_colItems = Nothing
2:    Set m_colKeys = Nothing
End Sub

Public Sub Add(ByRef varKey As Variant, ByRef varItem As Variant)
1:    On Error GoTo Err

3:    Dim strKey  As String

    'If the count is 0, then determine what kinds of items we are adding
6:    If m_colItems.count = 0 Then _
        m_blnObject = IsObject(varItem)
    
    'Select compare mode to determine key
10:    If m_intCompare = vbTextCompare Then _
        strKey = LCase$(CStr(varKey)) _
    Else _
        strKey = CStr(varKey)
        
    'Add to key/item collections
16:    m_colItems.Add varItem, strKey
17:    m_colKeys.Add varKey, strKey
    
19:    Exit Sub
    
21:
Err:
End Sub

Public Function Exists(ByRef varKey As Variant) As Boolean
1:    On Error GoTo Err
    
    'Select key based on compare mode
    'If the object doesn't exist, an error will be raised, leaving Exists = False
5:    If m_intCompare = vbTextCompare Then _
        m_colKeys.Item LCase$(CStr(varKey)) _
    Else _
        m_colKeys.Item CStr(varKey)
    
10:    Exists = True
    
12:    Exit Function
    
14:
Err:
End Function

Public Sub Remove(ByRef varKey As Variant)
1:    On Error GoTo Err
    
3:    Dim strKey      As String
    
    'Select key based on compare mode
6:    If m_intCompare = vbTextCompare Then _
        strKey = LCase$(CStr(varKey)) _
    Else _
        strKey = CStr(varKey)
    
    'Remove item/key
12:    m_colItems.Remove strKey
13:    m_colKeys.Remove strKey
    
15:    Exit Sub
    
17:
Err:
End Sub

Public Sub RemoveAll()
    'Create new collections to clear all items out
2:    Set m_colItems = New Collection
3:    Set m_colKeys = New Collection
End Sub

Public Property Get Items() As Collection
1:    Set Items = m_colItems
End Property

Public Property Get Keys() As Collection
1:    Set Keys = m_colKeys
End Property

Public Property Set Items(ByRef colData As Collection)
1:    Set m_colItems = colData
End Property

Public Property Set Keys(ByRef colData As Collection)
1:    Set m_colKeys = colData
End Property

Public Property Get count() As Long
1:    count = m_colItems.count
End Property

Public Property Get CompareMode() As VbCompareMethod
1:    CompareMode = m_intCompare
End Property

Public Property Let CompareMode(ByVal intData As VbCompareMethod)
1:    m_intCompare = intData
End Property

Public Property Get Item(ByRef varKey As Variant) As Variant
Attribute Item.VB_UserMemId = 0
1:    On Error GoTo Err

    'If it is a object collection, then we must use Set
4:    If m_blnObject Then
        'Key depends on compare mode
6:        If m_intCompare = vbTextCompare Then _
            Set Item = m_colItems(LCase$(CStr(varKey))) _
        Else _
            Set Item = m_colItems(CStr(varKey))
10:    Else
        'Key depends on compare mode
12:        If m_intCompare = vbTextCompare Then _
            Item = m_colItems(LCase$(CStr(varKey))) _
        Else _
            Item = m_colItems(CStr(varKey))
16:    End If
        
18:    Exit Property
    
20:
Err:
End Property

Public Property Let Item(ByRef varKey As Variant, ByRef varNewItem As Variant)
1:    On Error GoTo Err
    
3:    Dim strKey  As String
    
5:    If Exists(varKey) Then
        'Modify key based on compare type
7:        If m_intCompare = vbTextCompare Then _
            strKey = LCase$(CStr(varKey)) _
        Else _
            strKey = CStr(varKey)
    
        'Remove old item from collection and add new one
13:        m_colItems.Remove strKey
14:        m_colItems.Add varNewItem, strKey
15:    Else
16:        Add varKey, varNewItem
17:    End If
    
19:    Exit Property
    
21:
Err:
End Property

Public Property Set Item(ByRef varKey As Variant, ByRef varNewItem As Variant)
1:    On Error GoTo Err
    
3:    Dim strKey  As String
    
5:    If Exists(varKey) Then
        'Modify key based on compare type
7:        If m_intCompare = vbTextCompare Then _
            strKey = LCase$(CStr(varKey)) _
        Else _
            strKey = CStr(varKey)
    
        'Remove old item from collection and add new one
13:        m_colItems.Remove strKey
14:        m_colItems.Add varNewItem, strKey
15:    Else
16:        Add varKey, varNewItem
17:    End If
    
19:    Exit Property
    
21:
Err:
End Property

Public Property Let Key(ByRef varKey As Variant, ByRef varNewKey As Variant)
1:    On Error GoTo Err
    
3:    Dim varItem     As Variant

    'Get item
6:    If m_blnObject Then _
        Set varItem = Item(varKey) _
    Else _
        varItem = Item(varKey)
    
    'Remove old key and add item/key based on new one
12:    Remove varKey
13:    Add varNewKey, varItem
    
15:    Exit Property
    
17:
Err:
End Property

