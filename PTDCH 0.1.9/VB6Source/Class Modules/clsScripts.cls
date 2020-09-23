VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScripts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'API calls
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

'Constants
Private Const PPC_JSCRIPT       As Integer = 1
Private Const PPC_VBSCRIPT      As Integer = 2
Private Const PPC_PERLSCRIPT    As Integer = 3

Private Const PPC_LIBRARY       As Integer = 4 Or PPC_VBSCRIPT Or PPC_JSCRIPT Or PPC_PERL
Private Const PPC_INCLUDE       As Integer = 8 Or PPC_VBSCRIPT Or PPC_JSCRIPT Or PPC_PERL
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

Private Sub Class_Initialize()
'    Set m_lstScript = frmHub.lstScripts
End Sub

Private Sub Class_Terminate()
'    Set m_ListBox = Nothing
End Sub

Public Sub SLoad(Optional ByVal bytRLMode As Byte)

1:    Static blnLoaded    As Boolean

6:    Dim objSC           As ScriptControl
7:    Dim frmWS           As frmSocks
8:    Dim objSV           As clsDictionary
9:    Dim WFD             As WIN32_FIND_DATA

11:    Dim lngOne          As Long
12:    Dim lngTwo          As Long
13:    Dim intIndex        As Integer
14:    Dim strTemp         As String
15:    Dim strLanguage     As String

17:    On Error GoTo Err

19:    Set lvwItems = frmHub.lvwScripts.ListItems
20:    Set g_colSWinsocks = New Collection
21:    Set g_colSVariables = New Collection
    
       'Check to see if it's been loaded before
       'If True, then unload forms/listitems
25:    If blnLoaded Then
           'Delete controls
37:        lngTwo = frmHub.ScriptControl.UBound
39:        If lngTwo Then
40:            For lngOne = 1 To lngTwo
41:                Unload frmHub.ScriptControl(lngOne)
42:                Unload frmHub.tmrScriptTimer(lngOne)
43:            Next
44:        End If
           'Clear listview
53:        frmHub.lstScripts.Clear
56:    Else
57:        blnLoaded = True
58:    End If
    
       'Resize/clear out event array
61:    frmHub.SResizeArrEvent 1, False
    
       'Get first file handle
64:    lngOne = FindFirstFile(G_APPPATH & "\Scripts\*.*", WFD)
    
       'If it doesn't equal -1, then there are files
67:    If Not lngOne = -1 Then
68:        Do Until lngTwo = 18&
               'Can't be a directory
70:            If Not (WFD.dwFileAttributes And &H10) = vbDirectory Then
                   'Extract file name
72:                lngTwo = InStrB(1, WFD.cFileName, vbNullChar)
                
74:                If lngTwo Then _
                        strTemp = LeftB$(WFD.cFileName, lngTwo) _
                   Else strTemp = WFD.cFileName
                    
79:                lngTwo = InStrRev(strTemp, ".")
                
                    'Check extension and determine language
                    Select Case Mid$(strTemp, lngTwo + 1)
                        Case "vbs", "script": strLanguage = "VBScript"
                        Case "js": strLanguage = "JScript"
                        Case "pl": strLanguage = "PerlScript"
                        Case Else: GoTo NextLoop
82:                 End Select
                    
                    frmHub.lstScripts.AddItem strTemp
                    
                    'Increment count
85:                 intIndex = intIndex + 1
                    
                    'Load objects
                    Load frmHub.ScriptControl(intIndex)
                    Load frmHub.tmrScriptTimer(intIndex)
                    
                    'Get scriptcontrol
                    Set objSC = frmHub.ScriptControl(intIndex)

                   'Load winsock collection
104:                Set frmWS = New frmSocks
105:                Set frmWS.Script = objSC

106:                frmWS.Tag = CStr(intIndex)
107:                g_colSWinsocks.Add frmWS, CStr(intIndex)
                    
                   'Load static var dictionary
110:                Set objSV = New clsDictionary
111:                g_colSVariables.Add objSV, CStr(intIndex)
    
                    'Set settings
114:                objSC.Language = strLanguage
115:                objSC.Timeout = g_objSettings.ScriptTimeout
116:                objSC.UseSafeSubset = g_objSettings.ScriptSafeMode
117:            End If

119:
NextLoop:
                'Get next file
121:            lngTwo = FindNextFile(lngOne, WFD)
        
                'Exit if it's zero
124:            If lngTwo = 0 Then Exit Do
125:        Loop
        
            'Redim array if needed
            Select Case intIndex
                Case 0, 1
                Case Else
128:                    frmHub.SResizeArrEvent intIndex, False
129:        End Select
        End If

140:    Exit Sub
    
142:
Err:
143:    HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.SLoad()"
End Sub
