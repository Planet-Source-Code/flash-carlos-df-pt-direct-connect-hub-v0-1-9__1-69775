VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegExps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Regular Expression Functions
Option Explicit

Private m_objRegExp             As RegExp
Private m_objMatch              As Match
Private m_colMatches            As MatchCollection

Public Function TestStr(ByVal strString As String, ByVal strPattern As String) As Boolean
'-----------------------------------------------------------------
'Purpose:  test if a regular expression pattern is found in a string.(general)
'
'Params:  strString :  the string to test
'         strPattern : regular expression pattern
'
'Returns:  true if found
'------------------------------------------------------------------
9:    On Error GoTo Err

11:    Set m_objRegExp = New RegExp
    'Set the pattern
13:    m_objRegExp.Pattern = strPattern

    ' Set Case Insensitivity.
16:    m_objRegExp.IgnoreCase = True

    'Set global applicability. find all matches...
19:    m_objRegExp.Global = True

    'Test the String
22:    TestStr = m_objRegExp.Test(strString)

24:    Set m_objRegExp = Nothing

26: Exit Function

28:
Err:
29:    HandleError Err.Number, Err.Description, Erl & "|" & "clsRegExps.TestStr() " & strPattern
30:    Set m_objRegExp = Nothing
End Function

Public Function CaptureSubStr(ByVal strString As String, ByVal strPattern As String) As String
'-----------------------------------------------------------------
'Purpose:  capture a substring from a string using a regular expression pattern.
'
'Params:  strString :  the string to parse
'         strPattern : a regular expression pattern
'
'Returns:  first captured value. nothing if no capture or more then 1
'------------------------------------------------------------------
9:    On Error GoTo Err

11:    Set m_objRegExp = New RegExp

    'Set the pattern
14:    m_objRegExp.Pattern = strPattern

    ' Set Case Insensitivity.
17:    m_objRegExp.IgnoreCase = False

    'Set global applicability. find all matches...
20:    m_objRegExp.Global = True

    'Test whether the String can be compared.
23:    If (m_objRegExp.Test(strString) = True) Then
        'Get the matches. Execute search.
25:        Set m_colMatches = m_objRegExp.Execute(strString)
        'Set to first matches collection
27:        Set m_objMatch = m_colMatches(0)
        'set return to first match's value
29:        CaptureSubStr = m_objMatch.SubMatches(0)

31:    Else
32:        CaptureSubStr = Empty
33:    End If

35:    Set m_objMatch = Nothing
36:    Set m_colMatches = Nothing
37:    Set m_objRegExp = Nothing

39:
Err:
40:    HandleError Err.Number, Err.Description, Erl & "|" & "clsRegExps.CaptureSubStr() " & strPattern & "|" & strString
End Function

Public Function CaptureDbl(ByVal strString As String, ByVal strPattern As String) As Double
 '-----------------------------------------------------------------
'Purpose:  capture a substring within a string. (Specific to numbers)
'
'Params:  strString :  the string to parse
'         strPattern : a regular expression pattern (must capture numeric value(s) only
'
'Returns:  captured value or result of the addition of the captured values
'------------------------------------------------------------------
9:    Dim intTmp      As Double

11:    intTmp = 0

13:    Set m_objRegExp = New RegExp

15:    On Error GoTo Err
    'Set the pattern
17:    m_objRegExp.Pattern = strPattern

    ' Set Case Insensitivity.
20:    m_objRegExp.IgnoreCase = False

    'Set global applicability. find all matches...
23:    m_objRegExp.Global = True

    'Test whether the String can be compared.
26:    If (m_objRegExp.Test(strString) = True) Then
        'Get the matches. Execute search.
28:        Set m_colMatches = m_objRegExp.Execute(strString)

        Select Case m_colMatches.count
            '1 capture only, V:, S:, share size
            Case 1
31:                Set m_objMatch = m_colMatches(0)

                Select Case m_objMatch.SubMatches.count
                    Case 1
33:                        intTmp = m_objMatch.SubMatches(0)

                    'Hub count H:0 or H:0/0/0
                    Case 4
36:                        If IsEmpty(m_objMatch.SubMatches(3)) Then
37:                            intTmp = CDbl(m_objMatch.SubMatches(0))
38:                            intTmp = intTmp + CDbl(m_objMatch.SubMatches(1))

40:                            If g_objSettings.DCIncludeOPed Then
41:                                intTmp = intTmp + CDbl(m_objMatch.SubMatches(2))
42:                            End If
43:                        Else
44:                            intTmp = m_objMatch.SubMatches(3)
45:                        End If

                    Case Else
47:                        g_colUsers.SendChatToOps "DDCH_Degug", "CaptureDbl::m_objMatch.SubMatches.Count: " & m_objMatch.SubMatches.count & vbNewLine & "From : " & strString & vbNewLine & "Pattern : " & strPattern
48:                End Select

            Case Else
50:                g_colUsers.SendChatToOps "DDCH_Degug", "CaptureDbl::m_colMatches.Count: " & m_colMatches.count & vbNewLine & "From : " & strString & vbNewLine & "Pattern : " & strPattern
51:        End Select
52:    End If

54:    CaptureDbl = intTmp

56:    Set m_objMatch = Nothing
57:    Set m_colMatches = Nothing
58:    Set m_objRegExp = Nothing

60:
Err:
61:    HandleError Err.Number, Err.Description, Erl & "|" & "clsRexExps.CaptureDbl() " & strPattern & "|" & strString
End Function

Public Function REReplace(ByVal strString As String, ByVal strPattern As String, ByVal strReplace As String) As String
'-----------------------------------------------------------------
'Purpose:  Replace all occurence of a substring within a string
'
'Params:   strString :  the string to parse
'          strPattern : regular expression pattern
'
'Returns:  modified string
'------------------------------------------------------------------
9:    On Error GoTo Err

11:    Set m_objRegExp = New RegExp

    'Set the pattern
14:    m_objRegExp.Pattern = strPattern

    ' Set Case Insensitivity.
17:    m_objRegExp.IgnoreCase = False

    'Set global applicability. Replace all matches...
20:    m_objRegExp.Global = True

    'Test whether the String can be compared.
23:    If (m_objRegExp.Test(strString) = True) Then
24:        REReplace = m_objRegExp.Replace(strString, strReplace)
25:    Else
26:        REReplace = strString
27:    End If

29:    Set m_objRegExp = Nothing

31:    Exit Function

33:
Err:
34:    HandleError Err.Number, Err.Description, Erl & "|" & "clsRegExps.REReplace() " & strPattern & "|" & strString & "|" & strReplace
End Function

Public Function REMatchesCol(ByVal strString As String, ByVal strPattern As String) As MatchCollection
'-----------------------------------------------------------------
'Purpose:   Create a match(es) collection
'
'Params:    strString :  the string to parse
'           strPattern : regular expression pattern
'
'Returns:   Collection containing one, or more, Match(es), or empty Matchcollection
'
'Comment:   This only create the Matchcollection based on the pattern.
'           The number of Matchcollection and captured matches is base on the pattern given to this function.
'           Each match can be a captured match or an empty captured match, base on if a capturing pattern can match something in strString.
'           It is up to the coders to deal with the data in Matchcollection in the proper way.
'------------------------------------------------------------------
14:    On Error GoTo Err

16:    Set m_objRegExp = New RegExp

    'Set the pattern
19:    m_objRegExp.Pattern = strPattern

    ' Set Case Insensitivity.
22:    m_objRegExp.IgnoreCase = False

    'Set global applicability. find all matches...
25:    m_objRegExp.Global = True

    'Test whether the String can be compared.
28:    If m_objRegExp.Test(strString) Then
        'Get the matches. Execute search.
30:        Set REMatchesCol = m_objRegExp.Execute(strString)
31:    End If

33:    Set m_objRegExp = Nothing

35:
Err:
36:    HandleError Err.Number, Err.Description, Erl & "|" & "clsRegExps.REMatchesCol() " & strPattern & "|" & strString
End Function

Public Function AdvertTest(ByVal strString As String, ByVal strDeny As String, ByVal strAllow As String) As Boolean
'-----------------------------------------------------------------
'Purpose:  test if a regular expression pattern is found in a string
'           (Advertising specific)
'
'Params:    strString :     the string to test
'           arrDeny :       regular expression pattern(expected to be an array)
'           arrAllow:       regular expression pattern(expected to be an array)
'
'Returns:  true if match a denyed expression without matching an allowed one
'------------------------------------------------------------------

12:    Dim blnResultDeny   As Boolean
13:    Dim blnResultAllow  As Boolean
    'Dim X               As Long
    'Dim Y               As Long

17:    On Error GoTo Err

19:    Set m_objRegExp = New RegExp

        'set the pattern
22:        m_objRegExp.Pattern = strDeny
        'do not find all matches..., on first match should be enough.
24:        m_objRegExp.Global = False
        'set case insensitive
26:        m_objRegExp.IgnoreCase = True

28:        blnResultDeny = m_objRegExp.Test(strString)

30:        If blnResultDeny Then
                'set the pattern
32:                m_objRegExp.Pattern = strAllow
                'find all matches...,in case
34:                m_objRegExp.Global = True
                'set case insensitive
36:                m_objRegExp.IgnoreCase = True

38:                blnResultAllow = m_objRegExp.Test(strString)
                'exit on first match
40:                If blnResultAllow Then
41:                    AdvertTest = False
42:                    Set m_objRegExp = Nothing
43:                    Exit Function
44:                End If
            'no allow expression found, there are no reason to continu
            'g_colUsers.SendChatToAll g_objsettings.BotName, g_objSettings.AntiAdvertising
47:            AdvertTest = True
48:            Set m_objRegExp = Nothing
49:            Exit Function
50:        End If
        
    'no deny expression found
53:    AdvertTest = False
54:    Set m_objRegExp = Nothing

56: Exit Function

58:
Err:
59:    AdvertTest = False
60:    HandleError Err.Number, Err.Description, Erl & "|" & "clsRegExp.AdvertTest()"
End Function
