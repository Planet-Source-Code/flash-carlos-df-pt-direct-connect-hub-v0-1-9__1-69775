VERSION 5.00
Begin VB.Form frmScript 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   525
   LinkTopic       =   "Form1"
   ScaleHeight     =   360
   ScaleWidth      =   525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
'API calls
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

'Constants
Private Const PPC_JSCRIPT       As Integer = 1
Private Const PPC_VBSCRIPT      As Integer = 2
'Private Const PPC_PERLSCRIPT    As Integer = 3

Private Const PPC_LIBRARY       As Integer = 4 Or PPC_VBSCRIPT Or PPC_JSCRIPT 'Or PPC_PERL
Private Const PPC_INCLUDE       As Integer = 8 Or PPC_VBSCRIPT Or PPC_JSCRIPT 'Or PPC_PERL
Private Const PPC_ENDIF         As Integer = 16 Or PPC_VBSCRIPT
Private Const PPC_ELSE          As Integer = 32 Or PPC_VBSCRIPT
Private Const PPC_IF            As Integer = 64 Or PPC_VBSCRIPT
Private Const PPC_ELSEIF        As Integer = 128 Or PPC_VBSCRIPT
Private Const PPC_CONST         As Integer = 256 Or PPC_VBSCRIPT

Private Const CHR_SHARP         As Integer = 35
Private Const CHR_AT            As Integer = 64

'Types
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type

'Private variables
Private WithEvents m_objUpdate   As clsHTTPDownload
Attribute m_objUpdate.VB_VarHelpID = -1

Private Sub LvwAddItem(intIndex As Integer, _
                       strName As String, _
                       strLanguage As String)
1:    On Error GoTo Err
2:    Dim lvwItem As Variant
3:    Set lvwItem = frmHub.lvwScripts.ListItems.Add(intIndex, intIndex & "s", strName)

5:    lvwItem.SubItems(1) = "Inactive"
6:    lvwItem.SubItems(2) = strLanguage
7:    lvwItem.SubItems(3) = "False"

9:    Set lvwItem = Nothing
    
10:   Exit Sub
Err:
12:   HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.LvwAddItem(" & intIndex & ", " & strName & ", " & strLanguage & ")"
End Sub

Public Sub SLoadDir()
1:     Static blnLoaded    As Boolean

3:     Dim objSC           As ScriptControl
4:     Dim frmWS           As frmSocks
5:     Dim objSV           As clsDictionary
6:     Dim WFD             As WIN32_FIND_DATA
7:     Dim frmLoop         As Form

9:     Dim lngOne          As Long
10:    Dim lngTwo          As Long
11:    Dim intIndex        As Integer
       Dim i               As Integer
12:    Dim strTemp         As String
13:    Dim strLanguage     As String

15:    On Error GoTo Err
   
       'Check to see if it's been loaded before
       'If True, then unload forms/listitems
22:    If blnLoaded Then

           'Delete controls
25:        lngTwo = frmHub.ScriptControl.UBound
26:        If lngTwo Then
27:            For lngOne = 1 To lngTwo
28:                Unload frmHub.ScriptControl(lngOne)
29:                Unload frmHub.tmrScriptTimer(lngOne)
30:            Next
31:        End If

           'Clear forms
34:        For Each frmLoop In Forms
35:            Select Case frmLoop.Name
                   Case "frmProperties", "frmSocks"
37:                     Unload frmLoop
               End Select
39:        Next

41:        Set frmLoop = Nothing

           ' Erase array
46:        Erase sciMain()

           ' Unload objects
49:        For i = 1 To frmHub.picSciMain.UBound
50:             Unload frmHub.picSciMain(i)
51:        Next i

           Set g_colSWinsocks = Nothing
           Set g_colSVariables = Nothing
           
53:    Else

54:        blnLoaded = True

55:    End If
    
       Set g_colSWinsocks = New Collection
       Set g_colSVariables = New Collection
           
       'Clear listview and tab strip
58:    frmHub.lvwScripts.ListItems.Clear
59:    frmHub.tbsScripts.Tabs.Clear

       'Resize/clear out event array
62:    frmHub.SResizeArrEvent 1, False
    
       'If not found scripts..
65:    If Dir(G_APPPATH & "\Scripts\") = "" Then
66:        Call CreateDefautScript
67:        Exit Sub
68:    End If

       'Get first file handle
71:    lngOne = FindFirstFile(G_APPPATH & "\Scripts\*.*", WFD)
    
       'If it doesn't equal -1, then there are files
74:    If Not lngOne = -1 Then

76:        Do Until lngTwo = 18&

               'Can't be a directory
79:            If Not (WFD.dwFileAttributes And &H10) = vbDirectory Then

                    'Extract file name
82:                 lngTwo = InStrB(1, WFD.cFileName, vbNullChar)
                
84:                 If lngTwo Then _
                         strTemp = LeftB$(WFD.cFileName, lngTwo) _
                    Else strTemp = WFD.cFileName
                    
88:                 lngTwo = InStrRev(strTemp, ".")
                
                    'Check extension and determine language
91:                 Select Case Mid$(strTemp, lngTwo + 1)
                        Case "vbs", "script": strLanguage = "VBScript"
                        Case "js": strLanguage = "JScript"
'                       Case "pl": strLanguage = "PerlScript"
                        Case Else: GoTo NextLoop
                    End Select
                    
                    'Increment count
99:                 intIndex = intIndex + 1
                    
                    'Load new code editor.. and add new item to listview
102:                Call LvwAddItem(intIndex, strTemp, strLanguage)
103:                Call AddNewCodeEditor(intIndex, strTemp, strLanguage)
                    
                   'Load objects
106:                Load frmHub.ScriptControl(intIndex)
107:                Load frmHub.tmrScriptTimer(intIndex)

                    'Get scriptcontrol
110:                Set objSC = frmHub.ScriptControl(intIndex)

                   'Load winsock collection
113:                Set frmWS = New frmSocks
114:                Set frmWS.Script = objSC

115:                frmWS.Tag = CStr(intIndex)
116:                g_colSWinsocks.Add frmWS, CStr(intIndex)
                    
                   'Load static var dictionary
119:                Set objSV = New clsDictionary
120:                g_colSVariables.Add objSV, CStr(intIndex)
    
                    'Set settings
123:                objSC.Language = strLanguage
124:                objSC.Timeout = g_objSettings.ScriptTimeout
125:                objSC.UseSafeSubset = g_objSettings.ScriptSafeMode

127:           End If
                
129:
NextLoop:
               'Get next file
132:           lngTwo = FindNextFile(lngOne, WFD)
        
               'Exit if it's zero
135:           If lngTwo = 0 Then Exit Do

137:        Loop
        
            'Redim array if needed
140:        Select Case intIndex
                Case 0, 1
                Case Else
143:                    frmHub.SResizeArrEvent intIndex, False
            End Select

146:   End If
      
148:   Exit Sub
    
150:
Err:
152:   HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.SLoadDir()"
End Sub

Public Sub SLoadScript(strName As String)
1:     Dim objSC           As ScriptControl
2:     Dim frmWS           As frmSocks
3:     Dim objSV           As clsDictionary
4:     Dim frmLoop         As Form

6:     Dim intIndex        As Integer
7:     Dim strLanguage     As String
8:     Dim lngTwo As Long

10:    On Error GoTo Err

       If frmHub.ScriptControl.UBound <> 0 Then
12:         intIndex = (frmHub.ScriptControl.UBound + 1)
       Else
14:         intIndex = 1
       End If
       
       'Check extension and determine language
15:    If Right(strName, 3) = "vbs" Or Right(strName, 6) = "script" Then
16:         strLanguage = "VBScript"
17:    ElseIf Right(strName, 2) = "js" Then
18:         strLanguage = "JScript"
'19:    ElseIf Right(strName, 2) = "pl" Then
'20:         strLanguage = "PerlScript"
21:    Else
22:         Exit Sub
23:    End If

25:    Call LvwAddItem(intIndex, strName, strLanguage)
26:    Call AddNewCodeEditor(intIndex, strName, strLanguage)
                    
       'Load objects
29:    Load frmHub.ScriptControl(intIndex)
30:    Load frmHub.tmrScriptTimer(intIndex)
                    
       'Get scriptcontrol
33:    Set objSC = frmHub.ScriptControl(intIndex)

       'Load winsock collection
36:    Set frmWS = New frmSocks
37:    Set frmWS.Script = objSC

39:    frmWS.Tag = CStr(intIndex)
40:    g_colSWinsocks.Add frmWS, CStr(intIndex)
                    
       'Load static var dictionary
43:    Set objSV = New clsDictionary
44:    g_colSVariables.Add objSV, CStr(intIndex)
    
       'Set settings
47:    objSC.Language = strLanguage
48:    objSC.Timeout = g_objSettings.ScriptTimeout
49:    objSC.UseSafeSubset = g_objSettings.ScriptSafeMode

51:    If frmHub.ScriptControl.UBound <> 1 Or frmHub.ScriptControl.UBound <> 0 Then
52:         frmHub.SResizeArrEvent intIndex, False
53:    Else
54:         frmHub.SResizeArrEvent intIndex, True
55:    End If
       
57:    Exit Sub
    
59:
Err:
61:   HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.SLoadScript(" & strName & ")"
End Sub

Private Sub CreateDefautScript()
1:     On Error GoTo Err
2:     Dim strTemp As String
        
4:     Const strC_VBScript = "Option Explicit" & vbNewLine & vbNewLine & _
                              "Sub Main()" & vbNewLine & vbNewLine & _
                              vbTab & "MsgBox ""Hello World!"", , ""VBScript""" & vbNewLine & vbNewLine & _
                              "End Sub" & vbNewLine

9:     strTemp = g_colMessages.Item("msgNewScript") & ".vbs"

11:    g_objFileAccess.WriteFile G_APPPATH & "\Scripts\" & strTemp, strC_VBScript

13:    LvwAddItem 1, strTemp, "VBScripts"
14:    AddNewCodeEditor 1, strTemp, "VBScripts"

15:    Exit Sub
16:
Err:
18:    HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.CreatDefautScript()"
End Sub

Public Sub SSave(Optional intIndex As Integer = 0)
1:    On Error GoTo Err
2:    Dim i As Integer
3:    Dim strTemp As String
      
5:    With frmHub

7:        If intIndex = 0 Then
              'Save selected
9:            For i = 1 To .lvwScripts.ListItems.count
10:               If .lvwScripts.ListItems(i).Selected Then
11:                   strTemp = sciMain(i).Text
                      '
13:                   .tbsScripts.Tabs(i).Tag = strTemp
14:                   g_objFileAccess.WriteFile G_APPPATH & "\Scripts\" & .tbsScripts.Tabs(i).Key, strTemp
                      '
16:                   sciMain(i).ClearUndoBuffer
17:                   Exit Sub
18:               End If
19:           Next
20:       Else
              'Save by Index
22:           strTemp = sciMain(intIndex).Text
              '
24:           .tbsScripts.Tabs(intIndex).Tag = strTemp
25:           g_objFileAccess.WriteFile G_APPPATH & "\Scripts\" & .tbsScripts.Tabs(intIndex).Key, strTemp
              '
27:           sciMain(intIndex).ClearUndoBuffer
28:       End If
          
30:   End With

32:   Exit Sub
Err:
34:   HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.SSave(" & intIndex & ")"
End Sub

Public Sub SResetByName(strName As String, _
         Optional ByVal blnUpDateCode As Boolean = True, _
         Optional ByVal blnFirst As Boolean)
         
2:    Dim intIndex    As Integer
3:    Dim strTemp   As String
    
5:    On Error GoTo Err

7:    With frmHub
    
9:         For intIndex = 1 To .lvwScripts.ListItems.count
10:            If .lvwScripts.ListItems(intIndex).Text = strName Then
11:                If blnUpDateCode Then .tbsScripts.Tabs(intIndex).Tag = sciMain(intIndex).Text
12:                If SetSReset(intIndex, .ScriptControl(intIndex), blnFirst) Then
13:                     If blnUpDateCode Then Call SSave(intIndex)
14:                End If
15:                Exit Sub
16:            End If
17:        Next
    
19:   End With
    
21:   Exit Sub
    
23:
Err:
25:   HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.SResetByName(" & strName & ")"
End Sub

'Reset script.. only update the scripts to file, if no errors
Public Sub SReset(Optional ByVal lngSel As Long, _
                  Optional ByVal blnUpDateCode As Boolean = True, _
                  Optional ByVal blnFirst As Boolean)

2:    Dim intIndex    As Integer
3:    Dim strTemp   As String
    
5:    On Error GoTo Err
    
7:    With frmHub
                        
9:       Select Case lngSel
        
            '*********************************************
            Case -2 'All checked scripts
            '*********************************************
            
15:                 For intIndex = 1 To .lvwScripts.ListItems.count
16:                     If .lvwScripts.ListItems(intIndex).Checked Then
17:                         If blnUpDateCode Then .tbsScripts.Tabs(intIndex).Tag = sciMain(intIndex).Text
18:                         If SetSReset(intIndex, .ScriptControl(intIndex), blnFirst) Then
19:                                If blnUpDateCode Then Call SSave(intIndex)
20:                         End If
21:                     End If
22:                 Next

            '*********************************************
            Case -1 'All scripts
            '*********************************************

28:                 For intIndex = 1 To .lvwScripts.ListItems.count
29:                     If blnUpDateCode Then .tbsScripts.Tabs(intIndex).Tag = sciMain(intIndex).Text
30:                     If SetSReset(intIndex, .ScriptControl(intIndex), blnFirst) Then
31:                            If blnUpDateCode Then Call SSave(intIndex)
32:                     End If
33:                 Next
            
            '*********************************************
            Case 0 'Single script
            '*********************************************

39:                For intIndex = 1 To .lvwScripts.ListItems.count
40:                     If .lvwScripts.ListItems(intIndex).Selected Then
41:                        If blnUpDateCode Then .tbsScripts.Tabs(intIndex).Tag = sciMain(intIndex).Text
42:                        If SetSReset(intIndex, .ScriptControl(intIndex), blnFirst) Then
43:                               If blnUpDateCode Then Call SSave(intIndex)
44:                               Exit Sub
45:                        End If
46:                     End If
47:                Next

            '*********************************************
            Case Is > 0 ' by Index
            '*********************************************

53:               If blnUpDateCode Then .tbsScripts.Tabs(lngSel).Tag = sciMain(lngSel).Text
54:               If SetSReset(CInt(lngSel), .ScriptControl(CInt(lngSel)), blnFirst) Then
55:                        If blnUpDateCode Then Call SSave(CInt(lngSel))
56:               End If
                  
        End Select
        
60:   End With
    
62:   Exit Sub
    
64:
Err:
66:   HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.SReset(" & lngSel & ")"
End Sub

Private Function SetSReset(ByVal intIndex As Integer, _
                         objSC As ScriptControl, _
                         Optional ByVal blnFirst As Boolean) As Boolean
1:    Dim intChar     As Integer
2:    Dim strCode     As String
3:    Dim strPath     As String
    
5:     If Not blnFirst Then
6:       On Error Resume Next

         'Raise UnloadMain() event
9:       objSC.Run "UnloadMain"
10:    End If

11:    On Error GoTo Err
    
      'Reset script code/objects then readd objects
14:    objSC.Reset

      'Forms
17:    objSC.AddObject "frmHub", frmHub
18:    objSC.AddObject "frmScript", frmScript
    
      'Default VB objects
21:    objSC.AddObject "App", App
22:    objSC.AddObject "Forms", Forms

      'Default DC objects
25:    objSC.AddObject "tmrScriptTimer", frmHub.tmrScriptTimer(intIndex)
26:    objSC.AddObject "colUsers", g_colUsers

      'Extended PTDCH objects
29:    objSC.AddObject "wskScript", g_colSWinsocks(CStr(intIndex)).wskScript
30:    objSC.AddObject "colStatic", g_colSVariables(CStr(intIndex))
31:    objSC.AddObject "ScriptCtrl", objSC
32:    objSC.AddObject "Settings", g_objSettings
33:    objSC.AddObject "Functions", g_objFunctions, True
34:    objSC.AddObject "colRegistered", g_objRegistered
35:    objSC.AddObject "colIPBans", g_objIPBans
36:    objSC.AddObject "FileAccess", g_objFileAccess
37:    objSC.AddObject "colCommands", g_colCommands
39:    objSC.AddObject "RegExps", g_objRegExps
40:    objSC.AddObject "colLanguages", g_colLanguages
       'objSC.AddObject "colSheduler", g_colSheduler
    
       'Get first char to identify language
44:    intChar = AscW(objSC.Language)
    
46:    If intChar = 80 Then
47:        objSC.AddCode frmHub.tbsScripts.Tabs(intIndex).Tag
48:    Else
          'Prepare code buffer
50:        strCode = GenTempFile()
51:        g_objFileAccess.WriteFile strCode, frmHub.tbsScripts.Tabs(intIndex).Tag
        
           'Do preparsing actions if JScript/VBScript
54:        strPath = SSPrereset(objSC, strCode, vbNullString, intChar = 86)
        
           'Read code to control
57:        On Error Resume Next
58:        objSC.AddCode g_objFileAccess.ReadFile(strPath)
59:        On Error GoTo Err
        
61:        g_objFileAccess.DeleteFile strPath
62:        g_objFileAccess.DeleteFile strCode
63:    End If

       'Clear error text..
66:    frmHub.txtScriptError.Text = ""

       'If there was an error, then tell the user, and cancel reset
69:    If objSC.Error.Number Then

         'Report error / Add Log
72:       MsgBeep beepSystemDefault 'alert sound
73:       frmHub.txtScriptError.Text = "[" & Now & "] " & "Error resetting " & intIndex & " - Error " & objSC.Error.Number & " (" & objSC.Error.Description & ") on line " & CInt(objSC.Error.Line)
75:       AddLog "Error Resetting Script: " & intIndex & " - Error " & objSC.Error.Number & " (" & objSC.Error.Description & ") on line " & CInt(objSC.Error.Line), 4
          
          'Remove code/objects again
78:       objSC.Reset
79:       frmHub.SClearEvents intIndex

          'Make sure listitem is unchecked
82:       With frmHub.lvwScripts
83:            .ListItems(intIndex).Checked = False
84:            .ListItems(intIndex).SubItems(1) = "Inactive"
85:            .ListItems(intIndex).SubItems(3) = CStr(sciMain(intIndex).Modified)
86:            .ListItems(intIndex).SubItems(4) = Now
87:       End With

          ' return true
90:       SetSReset = False

92:    Else
          
          'Set events
95:       frmHub.SFindEvents intIndex
        
          'Make sure listitem is checked
98:       With frmHub.lvwScripts
99:            If .ListItems(intIndex).Checked = False Then _
                     .ListItems(intIndex).Checked = True
101:           .ListItems(intIndex).SubItems(1) = "Active"
102:           .ListItems(intIndex).SubItems(3) = CStr(sciMain(intIndex).Modified)
103:           .ListItems(intIndex).SubItems(4) = Now
104:      End With
            
          'Run Main
107:      On Error Resume Next
108:      objSC.Run "Main"

          ' return false
111:      SetSReset = True

113:   End If
   
115:   Exit Function
    
117:
Err:
119:   HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.SetSReset(" & intIndex & ")"
End Function

Public Sub SStopByName(strName As String)
         
2:    Dim intIndex    As Integer
3:    Dim strTemp   As String
    
5:    On Error GoTo Err

7:    With frmHub
    
9:         For intIndex = 1 To .lvwScripts.ListItems.count
10:            If .lvwScripts.ListItems(intIndex).Text = strName Then
11:                 SetSStop intIndex, .ScriptControl(intIndex)
12:                 Exit Sub
13:            End If
14:        Next
    
16:   End With
    
18:   Exit Sub
    
20:
Err:
22:   HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.SStopByName(" & strName & ")"
End Sub

Public Sub SStop(Optional ByVal lngSel As Long)

2:    Dim intIndex As Integer
    
4:    On Error GoTo Err
    
6:    With frmHub

8:       Select Case lngSel
        
            '*********************************************
            Case -2 'All checked scripts
            '*********************************************

14:                 For intIndex = 1 To .lvwScripts.ListItems.count
15:                     If .lvwScripts.ListItems(intIndex).Checked Then
                            'Stop script..
17:                         SetSStop intIndex, .ScriptControl(intIndex)
18:                     End If
19:                 Next

            '*********************************************
            Case -1 'All scripts
            '*********************************************
            
25:                 For intIndex = 1 To .lvwScripts.ListItems.count
                        'Stop script..
27:                     SetSStop intIndex, .ScriptControl(intIndex)
28:                 Next
                
            '*********************************************
            Case 0 'Single script
            '*********************************************

34:                 For intIndex = 1 To .lvwScripts.ListItems.count
35:                     If .lvwScripts.ListItems(intIndex).Selected Then
                            'Stop script..
37:                         SetSStop intIndex, .ScriptControl(intIndex)
38:                     End If
39:                 Next

            '*********************************************
            Case Is > 0 ' by Index
            '*********************************************
                   
                   'Stop script..
46:                SetSStop lngSel, .ScriptControl(CInt(lngSel))

        End Select

50:   End With

52:   Exit Sub
    
54:
Err:
56:   HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.SStop(" & lngSel & ")"
End Sub

Private Sub SetSStop(ByVal intIndex As Integer, objSC As ScriptControl)
1:    On Error Resume Next
    
      'Raise UnloadMain() event
4:    objSC.Run "UnloadMain"
    
6:    On Error GoTo Err
        
      'Reset all code/objects
9:    objSC.Reset
    
      'Set script event enabled status' to false
12:   frmHub.SClearEvents intIndex

      'Uncheck listitem
15:   With frmHub.lvwScripts
16:        If .ListItems(intIndex).Checked Then _
                 .ListItems(intIndex).Checked = False
18:        .ListItems(intIndex).SubItems(1) = "Inactive"
19:        .ListItems(intIndex).SubItems(2) = CStr(sciMain(intIndex).Modified)
20:        .ListItems(intIndex).SubItems(4) = Now
21:   End With
      
23:   Exit Sub

25:
Err:
27:    HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.SetSStop(" & intIndex & ")"
End Sub

Private Sub AddNewCodeEditor(intIndex As Integer, _
                             strName As String, _
                             strLanguage As String)
1:     On Error GoTo Err

3:     Dim strTemp, strTemp2 As String

7:     ReDim Preserve sciMain(intIndex)
            
9:     Set sciMain(intIndex) = New clsYScintilla

10:    strTemp = g_objFileAccess.ReadFile(G_APPPATH & "\Scripts\" & strName)
        
12:    Load frmHub.picSciMain(intIndex)

14:    Select Case intIndex
            Case 1: frmHub.picSciMain(intIndex).Visible = True
            Case Else: frmHub.picSciMain(intIndex).Visible = False
       End Select

25:    If Len(strName) > 18 Then _
            strTemp2 = Left(strName, 16) & ".." _
       Else strTemp2 = strName

29:    frmHub.tbsScripts.Tabs.Add (intIndex), strName, strTemp2
30:    frmHub.tbsScripts.Tabs(intIndex).Tag = strTemp
         
32:    sciMain(intIndex).CreateScintilla frmHub.picSciMain(intIndex)

34:    sciMain(intIndex).SetFixedFont "Courier New", 10
        
       'Give the scrollbar a nice long width to
       'handle a long line which may occur.
38:    sciMain(intIndex).ScrollWidth = 10000
       'This is absolutly an imperative line
40:    sciMain(intIndex).Attach frmHub.picSciMain(intIndex)
41:    sciMain(intIndex).Folding = True
42:    sciMain(intIndex).LineNumbers = True
43:    sciMain(intIndex).AutoIndent = True
44:    sciMain(intIndex).SetMarginWidth MarginLineNumbers, 50
45:    sciMain(intIndex).ContextMenu = True
46:    sciMain(intIndex).LineBreak = SC_EOL_CRLF


49:    Call Highlighter.SetHighlighterBasedOnExt(sciMain(intIndex), strName)

51:    sciMain(intIndex).Text = strTemp
            
53:    frmHub.tbsScripts.ZOrder vbSendToBack
       
55:    frmHub.Form_Resize

57:    sciMain(intIndex).ClearUndoBuffer
       
59:    Exit Sub
60:
Err:
62:    HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.AddNewCodeEditor(" & intIndex & ")"
End Sub

Public Sub SProperties(ByRef strIndex As String, _
                       ByRef strName As String, _
                       ByRef lngType As Long)

2:    Dim frmProp As Form
3:    Dim Modal As Byte

5:    Dim sFile As String
6:    Dim sXML As String

8:    On Error GoTo Err
    
      'If *.xml file properties not found, create on new..
11:     sFile = G_APPPATH & "\Scripts\" & (LeftB(strName, InStrB(1, strName, ".") - 1) & ".xml")
12:     If Not g_objFileAccess.FileExists(sFile) Then
13:         sXML = _
            "<Properties>" & vbNewLine & _
            vbTab & "<Author></Author>" & vbNewLine & _
            vbTab & "<Copyright></Copyright>" & vbNewLine & _
            vbTab & "<Version></Version>" & vbNewLine & _
            vbTab & "<Website></Website>" & vbNewLine & _
            vbTab & "<Description></Description>" & vbNewLine & _
            vbTab & "<Comments></Comments>" & vbNewLine & _
            "</Properties>"
22:         g_objFileAccess.WriteFile (sFile), sXML
23:     Else 'if found
           'Loop through to find if the form exists
25:         For Each frmProp In Forms
               'Check to see if it's the right kind of form
27:            If frmProp.Name = "frmProperties" Then
28:               If frmProp.Tag = strIndex Then
                     'Set focus
30:                  frmProp.SetFocus
31:                  Set frmProp = Nothing
32:                  Exit Sub
33:               End If
34:            End If
35:         Next
36:     End If
        'We haven't found a form and must create one
38:     Set frmProp = New frmProperties

40:     frmProp.Tag = strIndex
41:     frmProp.PType = lngType
42:     frmProp.file = strName

        'Set Full Selected Rum
45:     LVOptFRow frmProp.lvwProperties
46:     frmProp.stBar.Panels(1).Text = strName
        
        ' hook window for sizing control
        ' Disable the following line if you will be debugging form.
50:     Call HookWin(frmProp.hWnd, G_PrWnd)

52:     frmProp.Show Modal, frmHub
    
54:  Exit Sub
    
56:
Err:
58:    HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.PSLoad()"
End Sub

Public Sub XmlBooleanLoad()
1:   Dim objXML          As clsXMLParser
2:   Dim objNode         As clsXMLNode
3:   Dim colNodes        As Collection
4:   Dim colSubNodes     As Collection

6:   Dim strTemp         As String
7:   Dim i               As Integer

10:   On Error GoTo Err

12:    Set objXML = New clsXMLParser
      
17:    strTemp = G_APPPATH & "\Settings\Scripts.xml"

19:    If g_objFileAccess.FileExists(strTemp) Then
         
21:       objXML.Data = g_objFileAccess.ReadFile(strTemp)
22:       objXML.Parse

24:       Set colNodes = objXML.Nodes(1).Nodes

26:       On Error Resume Next

28:       For Each objNode In colNodes
29:            Set colSubNodes = objNode.Attributes
30:            With frmHub.lvwScripts
31:                For i = 1 To .ListItems.count
32:                   If .ListItems(i).Text = CStr(colSubNodes("Name").Value) Then
33:                         .ListItems(i).Checked = colSubNodes("Value").Value
34:                   End If
35:                Next
36:            End With
37:       Next

39:       On Error GoTo Err
    
41:       objXML.Clear
    
43:       Set objNode = Nothing
44:       Set colSubNodes = Nothing
45:       Set colNodes = Nothing

47:   End If

49:   Exit Sub
    
51:
Err:
53:   HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.XmlBooleanLoad()"
End Sub

Public Sub XmlBooleanSave()
1:    On Error GoTo Err
2:    Dim intFF       As Integer
3:    Dim strTemp     As String
4:    Dim i           As Integer

      'Save Scripts Value (Checked or UnChecked)
    
8:     strTemp = G_APPPATH & "\Settings\Scripts.xml"

10:    If g_objFileAccess.FileExists(strTemp) Then g_objFileAccess.DeleteFile strTemp
 
12:    intFF = FreeFile

14:    Open strTemp For Append As intFF
15:      Print #intFF, "<Scripts>"
16:        With frmHub.lvwScripts
17:            For i = 1 To .ListItems.count
18:                 Print #intFF, vbTab & "<Script Name=""" & .ListItems(i).Text & """" & " Value=""" & .ListItems(i).Checked & """ />"
19:            Next
20:        End With
21:      Print #intFF, "</Scripts>";
22:    Close intFF
    
24:   Exit Sub
    
26:
Err:
28:   HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.XmlBooleanSave()"
End Sub

Private Function SSPrereset(ByRef objSC As ScriptControl, _
                            ByRef strRead As String, _
                            ByRef strWrite As String, _
                            ByRef blnVBScript As Boolean) As String
    '------------------------------------------------------------------
    'Purpose:   To perform preprocessing commands given by the script,
    '           which are denoted by the symbol # or @
    '
    'Params:
    '           objSC:          Reference to script's control object
    '           strRead:        Path to script input file
    '           strWrite:       Path to script output file (generated if
    '                           not given) where code with the preprocessor
    '                           instructions are interpreted
    '           blnVBScript:    Toggles if language is VBScript or JScript
    '
    'Returns:
    '           strWrite (if it was given, it returns the same as the
    '           given path, otherwise it returns the path to the
    '           temporary file generated)
    '------------------------------------------------------------------
    
19:    Dim intRead         As Integer
20:    Dim intWrite        As Integer
21:    Dim intFlag         As Integer
    
23:    On Error GoTo Err
    
    'If VBScript, we search for #
    'If JScript, we search for @
27:    If blnVBScript Then
28:        intFlag = CHR_SHARP
29:    Else
30:        intFlag = CHR_AT
31:    End If
    
    'Open script for reading
34:    intRead = FreeFile
35:    Open strRead For Binary Access Read Lock Read Write As intRead
    
    'Create temporary file for appending to
38:    intWrite = FreeFile
    
40:    If StrPtr(strWrite) Then
41:        If LenB(Dir(strWrite)) Then
42:            Kill strWrite
43:        End If
44:    Else
45:        strWrite = GenTempFile()
46:    End If
        
48:    Open strWrite For Append Lock Read Write As intWrite
    
    'Begin preprocessing
51:    Preproc intRead, intWrite, intFlag, objSC
    
    'Close file handles
54:    Close intRead
55:    Close intWrite
    
    'Return path to code
58:    SSPrereset = strWrite
    
60:    Exit Function
    
62:
Err:
63:    HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.SSPrereset(, """ & strRead & """, """ & strWrite & """, " & blnVBScript & ")"
End Function

Private Sub ParseIf(ByVal intRead As Integer, _
                    ByVal intWrite As Integer, _
                    ByVal intChar As Integer, _
                    ByRef strExp As String, _
                    ByRef objSC As ScriptControl)
    '------------------------------------------------------------------
    'Purpose:   To interpret the preprocessor command #If, #ElseIf
    '           #Else and #End If.
    '
    '           #If/etc works just like their counterparts without
    '           the # except for one difference; the boolean expression
    '           is only ever evaluated once before starting the script.
    '           The code included in the script is for whichever
    '           statement for the #Ifs/ElseIfs evaluates to true first
    '           or the code for #Else if all are false and it is
    '           included.
    '
    '           Example:
    '               #If <expression> Then
    '                   'Include code for this exp
    '               #ElseIf <expression> Then
    '                   'If the first wasn't true, try this
    '               #Else
    '                   'Alright neither was true, include this code
    '               #End If
    '
    'Params:
    '           intRead:        File handle on input for the script code
    '           intWrite:       File handle on file to output parsed
    '                           code to
    '           intChar:        Character which denotes a preprocessor
    '                           command (#/@)
    '           strExp:         Expression which is to be evaluated
    '                           to determine if the code is to be
    '                           included or skipped (includes #If
    '                           and trailing Then)
    '           objSC:          Script's control object
    '------------------------------------------------------------------
    
35:    Dim strLine     As String
36:    Dim intCount    As Integer
37:    Dim intRet      As Integer
    
39:    On Error GoTo Err
    
    'Set count of totatal #End If to 1
42:    intCount = 1
    
    'Trim out expression to evaluate
45:    strExp = MidB$(strExp, 9, LenB(strExp) - 19)
    
    'Get a boolean out of it
48:    intRet = CBool(objSC.Eval(strExp))
    
    'Keep looping until we get a flag to exit
    'In other words, keep looping until we find an #If or #ElseIf which
    'evaluates to true, an #Else which is always true, or an #End If
53:    Do Until intRet Or EOF(intRead)
        'Make sure we don't pass the file boundaries
55:        Do Until EOF(intRead)
            'Take in a line and trim
57:            Line Input #intRead, strLine
58:            strLine = TrueTrim(strLine)
            
            'Skip null lines
61:            If LenB(strLine) Then
                'If it is a preproc command, everything is fine and dandy
63:                If AscW(strLine) = intChar Then
                    'Ignore #Include, #If, #Library, etc because this part of the
                    'code block is being ignored anyways
                    Select Case PreProcCmd(strLine, intChar)
                        Case PPC_IF
66:                            intCount = intCount + 1
                        Case PPC_ELSEIF
                            'Another expression to parse out and check
68:                            intRet = CBool(objSC.Eval(MidB$(strLine, 17, LenB(strLine) - 27)))
69:                            Exit Do
                        Case PPC_ELSE
                            '#Else means everything else failed and we must use this block
71:                            intRet = -1
72:                            Exit Do
                        Case PPC_ENDIF
73:                            intCount = intCount - 1
                            
75:                            If intCount = 0 Then
                                'Nothing left to the #If; we're done
77:                                intRet = 1
78:                                Exit Do
79:                            End If
80:                    End Select
81:                End If
82:            End If
83:        Loop
84:    Loop
    
    'Continue beyond this point only if there is a code block to include
87:    If intRet = -1 Then
        'Make sure we don't pass the file boundaries
89:        Do Until EOF(intRead)
            'Read line and trim whitespace
91:            Line Input #intRead, strLine
92:            strLine = TrueTrim(strLine)
            
            'Skip null lines
95:            If LenB(strLine) Then
                'Preproc command?
97:                If AscW(strLine) = intChar Then
                    'If so, check the type; now we must parse all of them
                    'because this is code we want to include in the script
    
                    'Another note is that we shouldn't parse any preproc commands
                    'if they are inside other blocks - hence the reason for the
                    'intRet checks
                    Select Case PreProcCmd(strLine, intChar)
                        Case PPC_LIBRARY
104:                            If intRet Then
105:                                ParseLibrary strLine, objSC
106:                            End If
                        Case PPC_INCLUDE
107:                            If intRet Then
108:                                ParseInclude intRead, intWrite, intChar, strLine, objSC
109:                            End If
                        Case PPC_CONST
110:                            If intRet Then
111:                                ParseConst intWrite, strLine, objSC
112:                            End If
                        Case PPC_IF
113:                            If intRet Then
114:                                ParseIf intRead, intWrite, intChar, strLine, objSC
115:                            End If
                        Case PPC_ELSEIF, PPC_ELSE
                            'If we've found an #ElseIf or #Else, that means we have
                            'to trim the rest of the #If/#End If block out before
                            'finishing up
119:                            intRet = 0
                        Case PPC_ENDIF
                            'Found the end of the block; exit out of the loop
121:                            Exit Do
122:                    End Select
123:                Else
                    'Only add code to script if we are in the block we want to keep
125:                    If intRet = -1 Then
126:                        Print #intWrite, strLine
127:                    End If
128:                End If
129:            Else
130:                If intRet = -1 Then
131:                    Print #intWrite, ""
132:                End If
133:            End If
134:        Loop
135:    End If

137:    Exit Sub
    
139:
Err:
140:    HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.ParseIf(" & intRead & ", " & intWrite & ", " & intChar & ", """ & strExp & """, )"
End Sub

Private Sub Preproc(ByVal intRead As Integer, _
                    ByVal intWrite As Integer, _
                    ByVal intChar As Integer, _
                    ByRef objSC As ScriptControl)
    '------------------------------------------------------------------
    'Purpose:   Searches for preprocessing commands in code and
    '           then calls the appropriate function to process
    '           any commands it finds
    '
    'Params:
    '           intRead:        File handle on input for the script code
    '           intWrite:       File handle on file to output parsed
    '                           code to
    '           intChar:        Character which denotes a preprocessor
    '                           command (#/@)
    '           objSC:          Script's control object
    '
    'Returns:
    '           Copy of strString without trailing/leading whitespace
    '------------------------------------------------------------------
    
18:    Dim strLine         As String
    
20:    On Error GoTo Err
    
    'Do until the end of the file is reached
23:    Do Until EOF(intRead)
        'Read line and trim whitespace
25:        Line Input #intRead, strLine
26:        strLine = TrueTrim(strLine)
        
        'Skip empty lines
29:        If LenB(strLine) Then
            'Is it a preproc command?
31:            If AscW(strLine) = intChar Then
                'Find out!
                Select Case PreProcCmd(strLine, intChar)
                    Case PPC_INCLUDE
33:                        ParseInclude intRead, intWrite, intChar, strLine, objSC
                    Case PPC_LIBRARY
34:                        ParseLibrary strLine, objSC
                    Case PPC_CONST
35:                        ParseConst intWrite, strLine, objSC
                    Case PPC_IF
36:                        ParseIf intRead, intWrite, intChar, strLine, objSC
                    Case PPC_ELSE, PPC_ELSEIF, PPC_ENDIF
                        'Orphaned statements it appears; just ignore 'em
38:                End Select
39:            Else
                'Just a regular line of code then; write to document
41:                Print #intWrite, strLine
42:            End If
43:        Else
44:            Print #intWrite, ""
45:        End If
46:    Loop
    
48:    Exit Sub
    
50:
Err:
    'Just to keep Error.txt clean because of
    '<Line Input> bug if last line of file is empty
    '"Input past end of file"
'    If (LCase(strLine) = "end sub") Or (LCase(strLine) = "end function") Or (strLine = "") Then _
'        Exit Sub
56:    If Err.Number = 62 Then Exit Sub
    
58:    HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.Preproc(" & intRead & ", " & intWrite & ", " & intChar & ", """ & strLine & """, )"
End Sub

Private Function PreProcCmd(ByRef strLine As String, _
                            ByVal intChar As Integer) As Integer
    '------------------------------------------------------------------
    'Purpose:   To check if the line starting with intChar is a valid
    '           preproc command, and if it is, to confirm that it has
    '           the proper form
    '
    'Params:
    '           strLine:        Line of code to evaluate to check if
    '                           it is a valid preprocessor command
    '           intChar:        Character which denotes a preproc command
    '
    'Returns:
    '           0 if it is not a valid command, otherwise it returns
    '           a unique numerical code denoting the type of command
    '           that we have
    '------------------------------------------------------------------
    
17:    Dim strTemp     As String
18:    Dim lngPos      As Long
    
20:    On Error GoTo Err
    
    'Set to all lower case
23:    strTemp = MidB$(LCase$(strLine), 3)
    
    'Pretty straight forward...#Else and #End If
    'are the only 2 commands thus far without any parameters
    'so check them first
28:    If strTemp = "end if" Then
29:        PreProcCmd = PPC_ENDIF
30:    ElseIf strTemp = "else" Then
31:        PreProcCmd = PPC_ELSE
32:    Else
        'OK, not either of those, extract first word
34:        lngPos = InStrB(1, strTemp, " ")
        
        'Make sure there is a word to extract
37:        If lngPos Then
            Select Case LeftB$(strTemp, lngPos - 1)
                Case "const"
38:                    PreProcCmd = PPC_CONST
                Case "if"
                    '#If has to end in " Then" to be valid
40:                    If RightB$(strTemp, 10) = " then" Then
41:                        PreProcCmd = PPC_IF
42:                    End If
                Case "elseif"
                    '#ElseIf has to end in " Then" to be valid
44:                    If RightB$(strTemp, 10) = " then" Then
45:                        PreProcCmd = PPC_ELSEIF
46:                    End If
                Case "include"
                    '#Include must have "" surrounding the path
48:                    If AscW(MidB$(strTemp, lngPos + 2)) = CHR_DQUOTE Then
49:                        If AscW(RightB$(strTemp, 2)) = CHR_DQUOTE Then
50:                            PreProcCmd = PPC_INCLUDE
51:                        End If
52:                    End If
                Case "library"
                    '#Library must have "" surrounding the name
54:                    If AscW(MidB$(strTemp, lngPos + 2)) = CHR_DQUOTE Then
55:                        If AscW(RightB$(strTemp, 2)) = CHR_DQUOTE Then
56:                            PreProcCmd = PPC_INCLUDE
57:                        End If
58:                    End If
59:            End Select
60:        End If
61:    End If
    
    'Confirm that the command we found
64:    If intChar = CHR_SHARP Then
65:        If (PreProcCmd And PPC_VBSCRIPT) = 0 Then
66:            PreProcCmd = 0
67:        End If
68:    Else
69:        If (PreProcCmd And PPC_JSCRIPT) = 0 Then
70:            PreProcCmd = 0
71:        End If
72:    End If
    
74:    Exit Function
    
76:
Err:
77:    HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.PreProcCmd(""" & strLine & """, " & intChar & ")"
End Function

Private Sub ParseConst(ByVal intWrite As Integer, _
                       ByRef strLine As String, _
                       ByRef objSC As ScriptControl)
    '------------------------------------------------------------------
    'Purpose:   To parse a preproc constant and add it to the script's
    '           control object to allow expressions to use it (#If,If,
    '           etc)
    '
    '           Format:
    '               #Const <name> = <value>
    '               #Const MYCONST = "this is a constant!!!"
    '               #Const MYNUM = 453
    '
    'Params:
    '           intWrite:       File handle on file to output parsed
    '                           code to
    '           strLine:        Code to extract constant from
    '           objSC:          Script's control object
    '------------------------------------------------------------------
    
18:    On Error GoTo Err
    
    'Add blank line to script to make line numbers more accurate
21:    Print #intWrite, ""
    
    'Create constant
24:    objSC.ExecuteStatement MidB$(strLine, 3)
    
26:    Exit Sub
    
28:
Err:
29:    HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.ParseConst(" & intWrite & ", """ & strLine & """, )"
End Sub

Private Sub ParseInclude(ByVal intRead As Integer, _
                         ByVal intWrite, _
                         ByVal intChar As Integer, _
                         ByRef strLine As String, _
                         ByRef objSC As ScriptControl)
    '------------------------------------------------------------------
    'Purpose:   To parse an #Include statement, which includes code
    '           from another file and inserts it into the script.
    '           Since the included code might contain preproc
    '           commands as well, we must start the process again
    '           for this as well (call to Preproc)
    '
    '           Base directory is the DDCH installation folder
    '
    '           Format:
    '               #Include "<path_to_file>"
    '               #Include "\Scripts\Includes\header.vbs"
    '
    'Params:
    '           intRead:        File handle on input for the script code
    '           intWrite:       File handle on file to output parsed
    '                           code to
    '           intChar:        Character which denotes a preproc command
    '           strLine:        Code to extract constant from
    '           objSC:          Script's control object
    '------------------------------------------------------------------
    
23:    Dim intIR       As Integer
    
25:    On Error GoTo Err
    
    'Open the file for reading
28:    intIR = FreeFile
29:    Open G_APPPATH & "\" & MidB$(strLine, 21, LenB(strLine) - 23) For Binary Access Read Lock Read Write As intIR
    
    'Begin preproc on the external code if any
32:    Preproc intIR, intWrite, intChar, objSC
    
    'Close file
35:    Close intIR
    
37:    Exit Sub
    
39:
Err:
40:    HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.ParseInclude(" & intRead & ", " & intWrite & ", " & intChar & ", """ & strLine & """, )"
End Sub

Private Sub ParseLibrary(ByRef strLib As String, _
                         ByRef objSC As ScriptControl)
'
End Sub
