VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXMLParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'General Library
'-------------------------------
'
'  Copyright (C) Andrew Osmond
'  E-Mail: tasburrfoot@users.sourceforge.net

Private m_strData       As String
Private m_strValue      As String
Private m_colNodes      As Collection

Private Sub Class_Initialize()
1:    Set m_colNodes = New Collection
End Sub

Private Sub Class_Terminate()
1:    Set m_colNodes = Nothing
End Sub

Public Property Let Data(ByRef strData As String)
1:    m_strData = strData
End Property

Public Property Get Data() As String
1:    Data = m_strData
End Property

Public Property Get Value() As String
1:    Value = m_strValue
End Property

Public Property Get Nodes() As Collection
1:    Set Nodes = m_colNodes
End Property

Public Sub Clear()
1:    m_strData = vbNullString
2:    m_strValue = vbNullString
3:    Set m_colNodes = New Collection
End Sub

Public Function Exists(ByRef strName As String) As Boolean
1:    On Error GoTo DNE
    
3:    m_colNodes.Item strName
4:    Exists = True
    
6:
DNE:
End Function

Public Sub Parse()
1:    Dim objNode     As clsXMLNode
2:    Dim strData     As String
3:    Dim strTag      As String
4:    Dim lngPos      As Long
5:    Dim lngPos2     As Long
    
7:    On Error GoTo Err
    
    'Clear out old information if needed
10:    If m_colNodes.count Then Set m_colNodes = New Collection
11:    If LenB(m_strValue) Then m_strValue = vbNullString

13:    strData = m_strData
    
    'Remove comments
16:    lngPos = InStrB(1, strData, "<!--")
    
18:    Do While lngPos
        'Find ending
20:        lngPos2 = InStrB(lngPos, strData, "-->")
        
        'If it exists, replace comment with null chars
        'else replace the rest of the string with them
24:        If lngPos2 Then _
            ZeroMemory ByVal StrPtr(strData) + lngPos - 1, lngPos2 - lngPos + 6 _
        Else _
            ZeroMemory ByVal StrPtr(strData) + lngPos - 1, LenB(strData) - lngPos + 1: Exit Do

29:        lngPos = InStrB(1, strData, "<!--")
30:    Loop
    
    'Remove other tags
33:    lngPos = InStrB(1, strData, "<?")
    
35:    Do While lngPos
        'Find ending
37:        lngPos2 = InStrB(lngPos, strData, "?>")
    
        'If it exists, replace comment with null chars
        'else replace the rest of the string with them
41:        If lngPos2 Then _
            ZeroMemory ByVal StrPtr(strData) + lngPos - 1, lngPos2 - lngPos + 4 _
        Else _
            ZeroMemory ByVal StrPtr(strData) + lngPos - 1, LenB(strData) - lngPos + 1: Exit Do

46:        lngPos = InStrB(1, strData, "<?")
47:    Loop
    
    'Replace null chars with nothing
50:    strData = Replace(strData, vbNullChar, vbNullString)
    
    'Split up the nodes
53:    lngPos = InStrB(1, strData, "<")
    
55:    Do While lngPos
56:        Set objNode = New clsXMLNode
        
58:        lngPos = lngPos + 2
        
        'An error will occur if there is an ">" in the attributes
        '(It should have the value &lt; / &gt;)
        'If there is no ">" then it will exit the loop
63:        lngPos2 = InStrB(lngPos, strData, ">")
64:        If lngPos2 = 0 Then
        '#If SVN Then
66:            Print #G_ERRORFILE, "clsXMLParser.Parse() Missing an >, xml file contain un-escaped character(s)" & strData
        '#End If
68:            Exit Do
69:        End If
        
71:        strTag = MidB$(strData, lngPos, lngPos2 - lngPos)
72:        strData = MidB$(strData, lngPos2 + 2)
        
74:        objNode.Name = strTag
        
        'If the tag name ends with a "/", then there is no value
77:        If Not AscW(RightB$(strTag, 2)) = 47 Then
78:            strTag = objNode.Name
            
80:            lngPos = InStrB(1, strData, "</" & strTag & ">")
            
            'If no end tag is found, exit loop
83:            If lngPos = 0 Then
            '#If SVN Then
85:                Print #G_ERRORFILE, "clsXMLParser.Parse() Missing an </>, xml file contain un-escaped character(s)" & strData
            '#End If
87:                Exit Do
88:            End If
            
90:            objNode.Value = LeftB$(strData, lngPos - 1)
91:            strData = MidB$(strData, lngPos + 6)
92:        End If
        
94:        On Error Resume Next
        
        'Add to collection
97:        m_colNodes.Add objNode, strTag
        
        'If there was an error, add to collection unindexed
100:        If Err.Number Then
        #If SVN Then
102:            Print #G_ERRORFILE, "clsXMLParser.Parse() m_colNodes.Add, error..., added unindexed." & objNode.Value
        #End If
104:            m_colNodes.Add objNode
105:            Err.Clear
106:        End If
        
108:        On Error GoTo Err
        
        'Find next tag
111:        lngPos = InStrB(1, strData, "<")
112:    Loop
    
    'If there is any data left, place it in m_strValue
115:    If LenB(strData) Then
116:            m_strValue = strData
        #If SVN Then
118:            Print #G_ERRORFILE, "clsXMLParser.Parse() data left.., stored in m_strValue" & m_strValue
        #End If
120:        End If
121:    Exit Sub

123:
Err:
124:    HandleError Err.Number, Err.Description, Erl & "|" & "clsXMLParser.Parse()"
End Sub

Public Sub Create()
1:    Dim objNode     As clsXMLNode
2:    Dim intFF       As Integer
3:    Dim i           As Long
4:    Dim strPath     As String
    
6:    On Error GoTo Err
    
8:    If m_colNodes.count Then
        'Turn on random number generating with a seed number of the ms since OS start
10:        Randomize GetTickCount
        
        'Create path which is semi random
13:        strPath = G_APPPATH & "\XMLParser-CreateTemp-" & CLng(Rnd * 100) & "-" & CLng(Rnd * 1000) & ".xml"
        
        'Now attempt different paths until the file does not exist
16:        Do While g_objFileAccess.FileExists(strPath)
17:            Randomize GetTickCount
18:            intFF = intFF + 101
19:            strPath = G_APPPATH & "\XMLParser-CreateTemp-" & CLng(Rnd * intFF) & "-" & CLng(Rnd * intFF * 10) & ".xml"
20:        Loop
        
        'Open path for appending
23:        intFF = FreeFile
        
25:        Open strPath For Append As intFF
        
27:        For Each objNode In m_colNodes
28:            IndexNode objNode, intFF
29:        Next
        
31:        Close intFF
        
        'Read data and then delete file
34:        m_strData = g_objFileAccess.ReadFile(strPath)
35:        g_objFileAccess.DeleteFile strPath
36:    End If
    
38:    Exit Sub
    
40:
Err:
41:    HandleError Err.Number, Err.Description, Erl & "|" & "clsXMLNode.Create()"

    'If there is an error, make sure temp file was deleted
44:    If LenB(strPath) Then _
        g_objFileAccess.DeleteFile strPath
End Sub

Private Sub IndexNode(ByRef curNode As clsXMLNode, ByRef intFF As Integer)
1:    Dim strName     As String
2:    Dim objAttr     As clsXMLAttribute
3:    Dim objNode     As clsXMLNode
4:    Dim colTemp     As Collection

6:    On Error GoTo Err
    
8:    strName = curNode.Name
    
10:    Set colTemp = curNode.Attributes
    
    'Check if we need to add attributes to the tag
13:    If colTemp.count Then
14:        Print #intFF, "<" & strName;
        
        'Loop through them and add in standard format
17:        For Each objAttr In colTemp
18:            Print #intFF, " " & objAttr.Name & "=""" & XMLEscape(objAttr.Value) & """";
19:        Next
        
        'Add value (if any) to tag
22:        Print #intFF, ">" & XMLEscape(curNode.Value);
23:    Else
24:        Print #intFF, "<" & strName & ">" & XMLEscape(curNode.Value);
25:    End If
    
27:    Set colTemp = curNode.Nodes
    
    'If there are any subnodes in this node, index them
30:    If colTemp.count Then
31:        For Each objNode In colTemp
32:            IndexNode objNode, intFF
33:        Next
34:    End If
    
    'Close the tag
37:    Print #intFF, "</" & strName & ">"
    
39:    Exit Sub
    
41:
Err:
42:    HandleError Err.Number, Err.Description, Erl & "|" & "clsXMLParser.IndexNode()"
End Sub


