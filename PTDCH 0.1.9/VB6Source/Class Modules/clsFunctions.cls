VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'These are functions which the script can access directly without need for
'the Functions. part. (ex calling LoadObj, would be done as "LoadObj wskScript(1)"
'not "Functions.LoadObj wskScript(1)")

Option Explicit

Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Allows access to the load object sub (for wskScript - perhaps other uses)
Public Sub LoadObj(ByRef Object As Object)
1:    On Error Resume Next 'If object is already loaded, screw the error
2:    Load Object
End Sub

'Allows access to the unload object sub (for wskScript - perhaps other uses)
Public Sub UnloadObj(ByRef Object As Object)
1:    On Error Resume Next 'If object is not loaded, screw the error
2:    Unload Object
End Sub

'Create a new collection object and gives access it
Public Function NewCollection() As Collection
1:    Set NewCollection = New Collection
End Function

'Create a huffman class
Public Function NewHuffman() As clsHuffman
1:    Set NewHuffman = New clsHuffman
End Function

'Create a BZip2 class
Public Function NewBZip2() As clsBZip2
1:    Set NewBZip2 = New clsBZip2
End Function

'Create a ZLib class
Public Function NewZLib() As clsZLib
1:    Set NewZLib = New clsZLib
End Function

'Create an XML parser class
Public Function NewXMLParser() As clsXMLParser
1:    Set NewXMLParser = New clsXMLParser
End Function

'Create an XML node
Public Function NewXMLNode() As clsXMLNode
1:    Set NewXMLNode = New clsXMLNode
End Function

'Create an XML attribute
Public Function NewXMLAttribute() As clsXMLAttribute
1:    Set NewXMLAttribute = New clsXMLAttribute
End Function

'Create new Jet engine
Public Function NewJetEngine() As JetEngine
1:    Set NewJetEngine = New JetEngine
End Function

'Create new Database connection
Public Function NewConnection() As Connection
1:    Set NewConnection = New Connection
End Function

'Create new custom Dictionary
Public Function NewDictionary() As clsDictionary
1:    Set NewDictionary = New clsDictionary
End Function

''Create new file
'Public Function NewFile() As clsFile
'    Set NewFile = New clsFile
'End Function

'String parsing function
'Return string after the first occurance of strFind
Public Function AfterFirst(ByVal strString As String, ByVal strFind As String) As String
1:    Dim i As Long
    
3:    On Error GoTo Err
    
5:    i = InStrB(1, strString, strFind)
6:    If i Then AfterFirst = MidB$(strString, i + LenB(strFind))
    
8:    Exit Function
    
10:
Err:
11:    HandleError Err.Number, Err.Description, Erl & "|" & "clsFunctions.AfterFirst(""" & strString & """, """ & strFind & """)"
End Function

'String parsing function
'Return string before the first occurance of strFind
Public Function BeforeFirst(ByVal strString As String, ByVal strFind As String) As String
1:    Dim i As Long
    
3:    On Error GoTo Err
    
5:    i = InStrB(1, strString, strFind)
6:    If i Then BeforeFirst = LeftB$(strString, i - 1) Else BeforeFirst = strString
    
8:    Exit Function
    
10:
Err:
11:    HandleError Err.Number, Err.Description, Erl & "|" & "clsFunctions.BeforeFirst(""" & strString & """, """ & strFind & """)"
End Function

'String parsing function
'Return string after the last occurance of strFind
Public Function AfterLast(ByVal strString As String, ByVal strFind As String) As String
1:    Dim i As Long
  
3:    On Error GoTo Err
  
5:    i = InStrRev(strString, strFind)
6:    If i Then AfterLast = Mid$(strString, i + Len(strFind))
    
8:    Exit Function
    
10:
Err:
11:    HandleError Err.Number, Err.Description, Erl & "|" & "clsFunctions.AfterLast(""" & strString & """, """ & strFind & """)"
End Function

'String parsing function
'Return string before the last occurance of strFind
Public Function BeforeLast(ByVal strString As String, ByVal strFind As String) As String
1:    Dim i As Long
    
3:    On Error GoTo Err
    
5:    i = InStrRev(strString, strFind)
6:    If i Then BeforeLast = Left$(strString, i - 1) Else BeforeLast = strString
    
8:    Exit Function
    
10:
Err:
11:    HandleError Err.Number, Err.Description, Erl & "|" & "clsFunctions.BeforeLast(""" & strString & """, """ & strFind & """)"
End Function

'String parsing function
'Return string between the first occurance of strFirst and strSecond
Public Function BetweenFirst(ByVal strString As String, ByVal strFirst As String, ByVal strSecond As String) As String
1:    Dim i As Long, q As Long
  
3:    On Error GoTo Err
  
5:    i = InStrB(1, strString, strFirst)
6:    If i Then
7:        i = i + LenB(strFirst)
8:        q = InStrB(i, strString, strSecond)
9:        If q Then BetweenFirst = MidB$(strString, i, q - i)
10:    End If
    
12:    Exit Function
    
14:
Err:
15:    HandleError Err.Number, Err.Description, Erl & "|" & "clsFunctions.BetweenFirst(""" & strString & """, """ & strFirst & """, """ & strSecond & """)"
End Function

'String parsing function
'Return string between the last occurance of strFirst and strSecond
Public Function BetweenLast(ByVal strString As String, ByVal strFirst As String, ByVal strSecond As String) As String
1:    Dim i As Long, q As Long
  
3:    On Error GoTo Err
  
5:    i = InStrRev(strString, strFirst)
6:    If i Then
7:        i = i + Len(strFirst)
8:        q = InStrRev(strString, strSecond)
9:        If q > i Then BetweenLast = Mid$(strString, i, q - i)
10:    End If
    
12:    Exit Function
    
14:
Err:
15:    HandleError Err.Number, Err.Description, Erl & "|" & "clsFunctions.BetweenLast(""" & strString & """, """ & strFirst & """, """ & strSecond & """)"
End Function

'Shell
Public Function ShellExec(ByVal strFile As String, Optional ByVal strParameters As String, Optional ByVal strDirectory As String, Optional ByVal lngShowCmd As VbAppWinStyle = vbNormalFocus) As Long
1:    ShellExec = ShellExecute(frmHub.hWnd, vbNullString, strFile, strParameters, strDirectory, lngShowCmd)
End Function

'Determines the closest size in the approriate *B
Public Function ShareSize(ByVal dblBytes As Double) As String
1:    On Error GoTo Err
    
    Select Case dblBytes
        Case Is > 1099511627775#
3:            ShareSize = FormatNumber(dblBytes / 1099511627776#, 2) & " TiB"
        Case Is > 1073741823
4:            ShareSize = FormatNumber(dblBytes / 1073741824, 2) & " GiB"
        Case Is > 1048575
5:            ShareSize = FormatNumber(dblBytes / 1048576, 2) & " MiB"
        Case Is > 1023
6:            ShareSize = FormatNumber(dblBytes / 1024, 2) & " KiB"
        Case Else
7:            ShareSize = dblBytes & " Bytes"
8:    End Select
    
10:    Exit Function
    
12:
Err:
13:    HandleError Err.Number, Err.Description, Erl & "|" & "clsFunctions.ShareSize(" & dblBytes & ")"
End Function

'Converts a date to one used in database
Public Function DBDate(ByVal datDate As Date) As String
1:    On Error GoTo Err

3:    DBDate = Format$(datDate, "yyyy-mm-dd hh:mm:ss")
    
5:    Exit Function
    
7:
Err:
8:    HandleError Err.Number, Err.Description, Erl & "|" & "clsFunctions.DBDate(" & datDate & ")"
End Function

'Converts a class long value to a string value
Public Function ClassName(ByVal intClass As enuClass) As String
1:    On Error GoTo Err

    Select Case intClass
        Case -1: ClassName = "Locked"
        Case 0: ClassName = "Unknown"
        Case 1: ClassName = "Normal"
        Case 2: ClassName = "Mentored"
        Case 3: ClassName = "Registered"
        Case 4: ClassName = "Invisible"
        Case 5: ClassName = "VIP"
        Case 6: ClassName = "Op"
        Case 7: ClassName = "Invisible Op"
        Case 8: ClassName = "Super Op"
        Case 9: ClassName = "Invisible Super Op"
        Case 10: ClassName = "Admin"
        Case 11: ClassName = "Invisible Admin"
        Case Else: ClassName = "Invalid"
3:    End Select
    
5:    Exit Function
    
7:
Err:
8:    HandleError Err.Number, Err.Description, Erl & "|" & "clsFunctions.ClassName(" & intClass & ")"
End Function

'Returns milliseconds since system start
Public Function TickCount() As Long
1:    TickCount = GetTickCount
End Function

'Message box wrapper for JScript
Public Function Alert(ByRef strMessage As String, Optional ByVal intButtons As VbMsgBoxStyle = vbOKOnly, Optional ByRef strTitle As String = "JScript") As VbMsgBoxResult
1:    Alert = MsgBox(strMessage, intButtons, strTitle)
End Function

'Input box wrapper for JScript
Public Function Prompt(ByRef strMessage As String, Optional ByRef strTitle As String = "JScript", Optional ByRef strDefault As String) As String
1:    Prompt = InputBox(strMessage, strTitle, strDefault)
End Function

Public Sub ShowBallon(ByRef sTitle As String, _
                      ByRef sMsg As String, _
                   Optional lIconType As Integer = 4, _
                   Optional bSound As Boolean = True)
1:    On Error GoTo Err
2:        Select Case lIconType
              Case 0
4:             lIconType = ICON_INFO
              Case 1
6:             lIconType = ICON_WARNING
              Case 2
8:             lIconType = ICON_ERROR
              Case 3
10:            lIconType = ICON_USER
              Case Else
11:            lIconType = ICON_NONE
           End Select
           
13:    PopupBalloon frmHub, 111&, sTitle, sMsg, lIconType, bSound
         
15:  Exit Sub
16:
Err:
18:  HandleError Err.Number, Err.Description, Erl & "|" & "clsFunctions.ShowBallon()"
End Sub

'Get a string of random characters of specified length
Public Function RandomChars(ByRef lngLen As Long, ByRef chrStart As String, ByRef chrEnd As String) As String
1:    Dim i       As Long
2:    Dim k       As Integer
3:    Dim j       As Integer
    
5:    On Error GoTo Err
    
    'Get char codes
8:    k = Asc(chrStart)
9:    j = Asc(chrEnd)
    
    'Create buffer
12:    RandomChars = Space$(lngLen)
    
    'Set new seed number to get random ones
15:    Randomize GetTickCount
    
    'Start pasting in characters
18:    For i = 1 To lngLen
19:        Mid$(RandomChars, i, 1) = Chr$(Int((j - k + 1) * Rnd + k))
20:    Next
    
22:    Exit Function
    
24:
Err:
25:    HandleError Err.Number, Err.Description, Erl & "|" & "clsFunctions.RandomChars(" & lngLen & ", """ & chrStart & """, """ & chrEnd & """)", Err.LastDllError
End Function

'Converts a Variant user object to a clsUser reference
Public Function CUser(ByRef Expression As Variant) As clsUser
1:    On Error GoTo Err
    
3:    Set CUser = Expression
    
5:
Err:
End Function

'Converts a Variant object reference an Object reference
Public Function CObj(ByRef Expression As Variant) As Object
1:    On Error GoTo Err
    
3:    Set CObj = Expression
    
5:
Err:
End Function

Public Function GetENLangStr(ByVal strString As String) As String
'------------------------------------------------------------------
'Purpose:   Get core/reason messages
'
'Params:    strString :      string key to be translated (a.e. "LoggedIn")
'
'Return:    Return proper string
'
'Added:     svn 136
'------------------------------------------------------------------
10:    On Error GoTo Err
    
12:    Dim objLang As Collection

    'this one is sure since it is(will be) hardcoded.
15:    Set objLang = g_colLanguages("En")

17:    GetENLangStr = objLang.Item(strString)

19:    Set objLang = Nothing
20:    Exit Function

22:
Err:
23:    HandleError Err.Number, Err.Description, Erl & "|" & "clsFunctions.GetENLangStr()|String missing in En language.|" & strString
End Function

Public Function ValidLang(ByVal strLangID As String) As Boolean
'------------------------------------------------------------------
'Purpose:   To know if a language is supported
'
'Params:    strLangID : Language Id
'
'Return:    Return True if supported
'
'Added:     svn 139
'------------------------------------------------------------------
10:    Dim sItem     As Variant

12:    For Each sItem In g_colLanguages("Supported")
13:        If CStr(sItem) = strLangID Then
14:            ValidLang = True
15:            Exit Function
16:        End If
17:    Next

End Function

Public Function ListLangs() As String
'------------------------------------------------------------------
'Purpose:   List supported languages(users lauguages)
'
'Return:    Return  formated listing of supported languages
'           languages Id / language international name / localised lauguage name
'
'Added:     svn 139
'------------------------------------------------------------------
9:    Dim colLang     As Collection
10:    Dim sItem       As Variant
11:    Dim sTmp        As String
    
13: On Error GoTo Err
14:    sTmp = "Valid Language IDs are:" & vbNewLine
    
16:    For Each sItem In g_colLanguages("Supported")
17:        Set colLang = g_colLanguages.Item(sItem)
        
19:        sTmp = sTmp & sItem & vbTab & colLang.Item("InternationalName") & vbTab & colLang.Item("NationalName") & vbNewLine
20:    Next

22:    ListLangs = sTmp
23:    Exit Function
24:
Err:
25:    HandleError Err.Number, Err.Description, Erl & "|" & "clsFunctions.ListLangs()"
End Function

Public Sub AddScriptMessages(strPath As String)
1:    Dim objXML          As clsXMLParser
2:    Dim objNode         As clsXMLNode
3:    Dim objSubNode      As clsXMLNode
4:    Dim colNodes        As Collection
5:    Dim colSubNodes     As Collection
6:    Dim colAttributes   As Collection
7:    Dim objLang         As Collection

9:    On Error GoTo Err

11:    Set objXML = New clsXMLParser

13:    objXML.Data = g_objFileAccess.ReadFile(strPath)
14:    objXML.Parse

16:    Set colNodes = objXML.Nodes(1).Nodes

    'Just in case...
    On Error Resume Next

21:    For Each objNode In colNodes
22:        Set colSubNodes = objNode.Nodes

24:        Set objLang = g_colLanguages(objNode.Name)

        #If SVN Then
     '   Print #G_ERRORFILE, "Set to g_colLanguages(objNode.Name): " & objNode.Name
        #End If

30:        For Each objSubNode In colSubNodes
31:            objLang.Remove (objSubNode.Name)
32:            objLang.Add objSubNode.Value, objSubNode.Name

        #If SVN Then
     '   Print #G_ERRORFILE, "objSubNode.Value: " & objSubNode.Value & "    objSubNode.Name: " & objSubNode.Name
        #End If

            'If .Add causes an error then SubNode.Value already exists
            'so we have to edit objLang in case of possible changes
40:            If Err.Number Then
41:                objLang.Item(objSubNode.Name) = objSubNode.Value
42:                Err.Clear
43:            End If
44:        Next

46:        Set objLang = Nothing
47:    Next

49:    On Error GoTo Err

51:    objXML.Clear

53:    Set objSubNode = Nothing
54:    Set objNode = Nothing
55:    Set colSubNodes = Nothing
56:    Set colNodes = Nothing

58:    Exit Sub

60:
Err:
61:    HandleError Err.Number, Err.Description, Erl & "|" & "clsFunctions.AddScriptMessages(" & strPath & ")"
End Sub

'zerschneidet den String sData an den Bytes mit sSegmentChar und liefert das angegebene Segment zurück.
Public Function GetSegment(SegmentID, sData, sSegmentChar)
1:  Dim n, i, oi

3:  i = 0
4:  oi = i
5:  i = InStr(i + 1, sData, sSegmentChar)
6:  Do Until i = 0
7:    n = n + 1
8:    oi = i
9:    i = InStr(i + 1, sData, sSegmentChar)

11:    If n = SegmentID And i > 0 Then
12:      GetSegment = Mid(sData, oi + 1, i - oi - 1)
13:      Exit Do
14:    End If
15:  Loop
End Function

Public Function SplitParameter(sParameters, sDivChars)
1:  Dim n

3:  n = InStr(sParameters, sDivChars)
4:  If n = 0 Then
5:    SplitParameter = sParameters
6:    sParameters = ""
7:  Else
8:    SplitParameter = Left(sParameters, n - 1)
9:    sParameters = Trim(Mid(sParameters, n + Len(sDivChars)))
10:  End If
End Function

Public Function DictItemDel(sKey, dict)
1:    If CStr(sKey) = "" Then
2:        DictItemDel = "Missing parameter <item>."
3:    Else
4:        If dict.Exists(CStr(sKey)) Then
5:            dict.Remove CStr(sKey)
6:            DictItemDel = " deleted."
7:        Else
8:            DictItemDel = " not found."
9:        End If
10:        DictItemDel = "Item """ & sKey & """ " & DictItemDel
11:    End If
End Function

Public Function ForceMinLen(sData, nLen, sFillChar)
1:    Dim n, m, t

3:    n = nLen - Len(sData)
4:    If n < 0 Then n = 0

6:    t = ""
7:    For m = 1 To n
8:        t = t & sFillChar
9:    Next
10:    ForceMinLen = t & sData
End Function

Public Function ReplaceHubVars(ByRef curUser As clsUser, ByVal strMessage As String) As String
'    *%[maxhubs]              Max hubs
'    *%[minslots]             Min slots
'        %[maxslots]             Max slots
'        %[minversion]           DC++ min version, NMDC min version
'    *%[hsratio]              Hub/slot ratio
'        %[bsratio]              Bandwidth/slot ratio
'        %[maxshare]             Max share (uses ShareSize function formatting)
'    *%[minshare]             Min share

10:    On Error Resume Next
    
12:    If InStrB(1, strMessage, "%[minslots]") Then strMessage = Replace(strMessage, "%[minslots]", g_objSettings.MinSlots)
13:    If InStrB(1, strMessage, "%[minshare]") Then strMessage = Replace(strMessage, "%[minshare]", g_objSettings.MinShare)
14:    If InStrB(1, strMessage, "%[hsratio]") Then strMessage = Replace(strMessage, "%[hsratio]", g_objSettings.DCSlotsPerHub)
15:    If InStrB(1, strMessage, "%[maxhubs]") Then strMessage = Replace(strMessage, "%[maxhubs]", g_objSettings.DCMaxHubs)

17:    ReplaceHubVars = strMessage

End Function

Public Function ReplaceUserVars(ByRef curUser As clsUser, ByVal strMessage As String) As String
'    *%[nick]                 CurUser user name
'    *%[ip]                   CurUser user IP

4:    On Error Resume Next
    
6:    If InStrB(1, strMessage, "%[nick]") Then strMessage = Replace(strMessage, "%[nick]", curUser.sName)
7:    If InStrB(1, strMessage, "%[ip]") Then strMessage = Replace(strMessage, "%[ip]", curUser.IP)

9:    ReplaceUserVars = strMessage

End Function

