VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXMLNode"
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


Private m_strName           As String
Private m_strValue          As String
Private m_colAttributes     As Collection
Private m_colNodes          As Collection

Private Sub Class_Initialize()
1:    Set m_colNodes = New Collection
2:    Set m_colAttributes = New Collection
End Sub

Private Sub Class_Terminate()
1:    Set m_colNodes = Nothing
2:    Set m_colAttributes = Nothing
End Sub

Public Property Get Name() As String
1:    Name = m_strName
End Property

Public Property Let Name(ByRef strData As String)
1:    Dim objAttr     As clsXMLAttribute
2:    Dim lngPos      As Long
3:    Dim strTemp     As String
     
5:    On Error GoTo Err
     
    'Clear out collection if needed
8:    If m_colAttributes.count Then Set m_colAttributes = New Collection
     
    'Check if there are any attributes
11:    lngPos = InStrB(1, strData, " ")
    
13:    If lngPos Then
        'If so extract name, then start looping through attributes
15:        m_strName = LeftB$(strData, lngPos - 1)
16:        strTemp = MidB$(strData, lngPos + 2)
        
        'If it ends in a slash, remove it and any rightend spaces
19:        If AscW(RightB$(strTemp, 2)) = 47 Then strTemp = RTrim$(LeftB$(strTemp, LenB(strTemp) - 2))
         
21:        lngPos = InStrB(1, strTemp, "=")
        
23:        Do While lngPos
24:            Set objAttr = New clsXMLAttribute

26:            objAttr.Name = LTrim$(LeftB$(strTemp, lngPos - 1))
27:            strTemp = MidB$(strTemp, lngPos + 2)
             
            'If there are quote marks (there should be, but just in case),
            'we find the next quote the mark the end, otherwise a space marks
            'the end
32:            If AscW(strTemp) = 34 Then
33:                lngPos = InStrB(3, strTemp, """")
34:                objAttr.Value = XMLUnescape(MidB$(strTemp, 3, lngPos - 3))
35:                strTemp = MidB$(strTemp, lngPos + 4)
36:            Else
37:                lngPos = InStrB(1, strTemp, " ")
                'If there is no space, then it is the end of the attributes
39:                If lngPos Then
40:                    objAttr.Value = XMLUnescape(LeftB$(strTemp, lngPos - 1))
41:                    strTemp = MidB$(strTemp, lngPos + 2)
42:                Else
43:                    objAttr.Value = XMLUnescape(strTemp)
44:                    Exit Do
45:                End If
46:            End If
         
            'Add attribute to collection
49:            m_colAttributes.Add objAttr, objAttr.Name
         
            'Find next attribute
52:            lngPos = InStrB(1, strTemp, "=")
53:        Loop
54:    Else
55:        m_strName = strData
        
        'If it ends in a slash, remove it
58:        If AscW(RightB$(m_strName, 2)) = 47 Then m_strName = LeftB$(m_strName, LenB(m_strName) - 2)
59:    End If
     
61:    Exit Property
     
63:
Err:
64:    HandleError Err.Number, Err.Description, Erl & "|" & "Let clsXMLNode.Name(""" & strData & """)"
End Property

Public Property Get Value() As String
1:    Value = m_strValue
End Property

Public Property Let Value(ByRef strInput As String)
1:    Dim objNode     As clsXMLNode
2:    Dim lngPos      As Long
3:    Dim lngPos2     As Long
4:    Dim strData     As String
5:    Dim strTag      As String
    
7:    On Error GoTo Err
    
    'Clear out nodes if needed
10:    If m_colNodes.count Then Set m_colNodes = New Collection
    
12:    strData = strInput
13:    lngPos = InStrB(1, strData, "<")
    
15:    Do While lngPos
16:        Set objNode = New clsXMLNode

        'what is in between "<!\[CDATA\[" and "]]>" must not be processed
        'it must directly be assigne as the value.
        'XMLESCSTART = "<!\[CDATA\["
        'XMLESCEND = "]]>"
'       If lngPos = regexp.GetPos(strData, XMLESC) Then
'               objNode.Value = MidB$(strData, regexp.GetPos(strData, XMLESCEND))

25:        lngPos = lngPos + 2
        
        'An error will occur if there is an ">" in the attributes
        '(It should have the value &lt; / &gt;)
        'If there is no ">" then it will exit the loop
30:        lngPos2 = InStrB(lngPos, strData, ">")
31:        If lngPos2 = 0 Then
32:            Print #G_ERRORFILE, "clsXMLNode.Value Missing an >, xml file contain un-escaped character(s)" & strData
33:            Exit Do
34:        End If
        
36:        strTag = MidB$(strData, lngPos, lngPos2 - lngPos)
37:        strData = MidB$(strData, lngPos2 + 2)
        
39:        objNode.Name = strTag
        
        'If the tag name ends with a "/", then there is no value
42:        If Not AscW(RightB$(strTag, 2)) = 47 Then
43:            strTag = objNode.Name
            
45:            lngPos = InStrB(1, strData, "</" & strTag & ">")
            
            'If no end tag is found, exit loop
48:            If lngPos = 0 Then
49:                If Not strData = "" Then _
                        Print #G_ERRORFILE, "clsXMLNode.Value Missing an </>, xml file contain un-escaped character(s)" & strData
50:                Exit Do
51:            End If
            
53:            objNode.Value = LeftB$(strData, lngPos - 1)
54:            strData = MidB$(strData, lngPos + LenB(strTag) + 6)
55:        Else
56:            strTag = objNode.Name
57:        End If
        
59:        On Error Resume Next

        'Add to collection
62:        m_colNodes.Add objNode, strTag
        
        'If there was an error, then just add the node unindexed
65:        If Err.Number Then
66:            m_colNodes.Add objNode
            #If SVN Then
68:            Print #G_ERRORFILE, "clsXMLNode.Value m_colNodes.Add, error..., added unindexed." & objNode.Value
            #End If
70:            Err.Clear
71:        End If
        
73:        On Error GoTo Err
        
        'Find next tag
76:        lngPos = InStrB(1, strData, "<")
77:    Loop

    'If there is anything left, then assign it to value
80:    If LenB(strData) Then
    #If SVN Then
82:        Print #G_ERRORFILE, "clsXMLNode.Value data left.., stored in m_strValue" & strData
    #End If
84:        m_strValue = XMLUnescape(strData)
85:    End If

87:    Exit Property

89:
Err:
90:    HandleError Err.Number, Err.Description, Erl & "|" & "Let clsXMLNode.Value(""" & strData & """)"
End Property

Public Property Get Nodes() As Collection
1:    Set Nodes = m_colNodes
End Property

Public Property Get Attributes() As Collection
1:    Set Attributes = m_colAttributes
End Property

Public Function Exists(ByRef strName As String, Optional ByVal bytType As Byte) As Boolean
1:    On Error GoTo DNE
    
3:    If bytType Then _
        m_colAttributes.Item strName _
    Else _
        m_colNodes.Item strName

8:    Exists = True
    
10:
DNE:
End Function

