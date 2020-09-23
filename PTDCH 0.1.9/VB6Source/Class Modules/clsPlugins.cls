VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlugins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'----------------------------------------------'
'  Copyright (C) fLaSh - Carlos Ferreira       '
'  E-Mail: Carlosferreiracarlos@hotmail.com    '
'----------------------------------------------'

Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)

Public Enum RegOp
    Register = 1
    UnRegister = 0
End Enum

Private lvwItems        As ListItems
Private lvwItem         As ListItem

Private mPlg_Path       As String
Private mE_pos          As Integer

Private m_sPlugins()    As String
Private m_iPlgCount     As Integer
Private m_ErrLoading    As Boolean

Public Function RegisterActiveX(lzAxDll As String, mRegOption As RegOp) As Boolean
1:   On Error GoTo Err

3:   Dim mLib As Long, DllProcAddress As Long
4:   Dim mThread
5:   Dim sWait As Long
6:   Dim mExitCode As Long
7:   Dim lpThreadID As Long

9:    mLib = LoadLibrary(lzAxDll)
    
10:    If mLib = 0 Then RegisterActiveX = False: Exit Function

12:    If mRegOption = Register Then
13:        DllProcAddress = GetProcAddress(mLib, "DllRegisterServer")
14:    Else
15:        DllProcAddress = GetProcAddress(mLib, "DllUnregisterServer")
16:    End If
    
18:    If DllProcAddress = 0 Then
19:        RegisterActiveX = False
20:        Exit Function
21:    Else
22:        mThread = CreateThread(ByVal 0, 0, ByVal DllProcAddress, ByVal 0, 0, lpThreadID)
23:        If mThread = 0 Then
24:            FreeLibrary mLib
25:            RegisterActiveX = False
26:            Exit Function
27:        Else
28:            sWait = WaitForSingleObject(mThread, 10000)
29:            If sWait <> 0 Then
30:                FreeLibrary mLib
31:                mExitCode = GetExitCodeThread(mThread, mExitCode)
32:                ExitThread mExitCode
33:                Exit Function
34:            Else
35:                FreeLibrary mLib
36:                CloseHandle mThread
37:            End If
38:        End If
39:    End If
    
41:    RegisterActiveX = True
    
43:   Exit Function
44:
Err:
46:    HandleError Err.Number, Err.Description, Erl & "|" & "clsPlugins.RegisterActiveX()"
End Function

Private Function GetFileExt(lzFile As String) As String
1:    mE_pos = InStrRev(lzFile, ".", Len(lzFile), vbBinaryCompare)
2:    If mE_pos = 0 Then
3:        GetFileExt = lzFile
4:    Else
5:        GetFileExt = UCase(Mid(lzFile, mE_pos + 1, Len(lzFile)))
6:    End If
End Function

Private Function GetFileTitle(lzFile As String) As String
1: Dim X As Integer, mE_pos As Integer
2:    For X = 1 To Len(lzFile)
3:        If Mid(lzFile, X, 1) = "." Then mE_pos = X
4:    Next X
    
6:    If mE_pos = 0 Then
7:        GetFileTitle = lzFile
8:    Else
9:        GetFileTitle = Mid(lzFile, 1, mE_pos)
10:   End If
End Function

Public Sub InstallPlugins()
1:    On Error GoTo Err

3:    Dim strFileName As String, strFile As String, strPlgServ As String
4:    Dim intIndex As Integer
5:    Dim objPlugin As Object
   
7:    Dim PlgObjRun As Object
     
13:   lvwItems.Clear
                      
15:   mPlg_Path = (G_APPPATH & "\Plugins\")

17:   If Dir(mPlg_Path) = "" Then Exit Sub
   
19:   With frmHub

20:      strFileName = Dir(mPlg_Path)
21:      .imlAddIns.ListImages.Clear

22:      intIndex = 0

24:      Do While strFileName <> ""

26:         If GetFileExt(strFileName) = "DLL" Then

28:            strFile = mPlg_Path & strFileName
29:            strPlgServ = GetFileTitle(strFileName) & "main"

               'Already installed ?
32:            If Not AlreadyInstalled(strFileName) Then
33:                If Not RegisterActiveX(strFile, Register) Then
34:                     AddLog "There was an error loading the Plugin:" & strFileName, 4
35:                     m_ErrLoading = True
36:                Else
37:                     AddLog "Plugin " & strFileName & " Installed with success.", 6
38:                     GoTo LoadPlugin
39:                End If
40:            Else
LoadPlugin:
42:                Set objPlugin = CreateObject(strPlgServ)
43:                If objPlugin.loadplug <> 1 Then
44:                     AddLog "There was an error loading the Plugin: " & strFileName, 4
45:                Else

47:                   intIndex = intIndex + 1
                     
                     'add image
50:                   On Error Resume Next
51:                  .imlAddIns.ListImages.Add intIndex, , objPlugin.ipic
52:                   On Error GoTo Err
                      
                      'Add listitem
55:                   If objPlugin.PlgSetup Then
                         'add key in list view for availabed setup button
57:                      Set lvwItem = lvwItems.Add(, strPlgServ, , , intIndex)
58:                   Else
59:                      Set lvwItem = lvwItems.Add(, , , , intIndex)
60:                   End If
                      
                      'Add sub listitems
63:                   lvwItem.SubItems(1) = objPlugin.Name
64:                   lvwItem.SubItems(2) = objPlugin.Version
65:                   lvwItem.SubItems(3) = objPlugin.Author
66:                   lvwItem.SubItems(4) = objPlugin.Description
67:                   lvwItem.SubItems(5) = objPlugin.Release
68:                   lvwItem.SubItems(6) = objPlugin.Comments
69:
                     'Run plugin SubMain .. if then ..
71:                   Set PlgObjRun = CreateObject(strPlgServ)

73:                   PlgObjRun.SubMain frmHub
                     
75:                   Set objPlugin = Nothing
76:                   Set PlgObjRun = Nothing

78:                   strPlgServ = ""
79:                   strFile = ""

81:                   AddLog "Plugin " & strFileName & " loaded with success.", 6
                      
83:               End If
84:            End If

86:         End If

88:         strFileName = Dir
89:         DoEvents

91:      Loop
         
93:   End With
   
95:   Exit Sub

99:   If g_objFileAccess.FileExists(G_APPPATH & "\Plugins\Plugins.xml") Then _
            g_objFileAccess.DeleteFile G_APPPATH & "\Plugins\Plugins.xml"
 
102:
Err:
104:   HandleError Err.Number, Err.Description, Erl & "|" & "clsPlugins.InstallPlugins()"
105:   Resume Next
End Sub

Private Function AlreadyInstalled(strName As String) As Boolean
1:   Dim objXML          As clsXMLParser
2:   Dim objNode         As clsXMLNode
3:   Dim colNodes        As Collection
4:   Dim colSubNodes     As Collection

6:   Dim strTemp         As String
7:   Dim i               As Integer

10:   On Error GoTo Err

12:    Set objXML = New clsXMLParser
      
17:    strTemp = G_APPPATH & "\Plugins\Plugins.xml"

19:    AlreadyInstalled = False

21:    If g_objFileAccess.FileExists(strTemp) Then
         
23:       objXML.Data = g_objFileAccess.ReadFile(strTemp)
24:       objXML.Parse

26:       Set colNodes = objXML.Nodes(1).Nodes

28:       On Error Resume Next

30:       For Each objNode In colNodes
31:            Set colSubNodes = objNode.Attributes
               'Check if plugin name is found
33:            If strName = CStr(colSubNodes("Name").Value) Then
                   'The dir is verified was modified ?
35:                If CStr(XMLUnescape(colSubNodes("Value").Value)) = CStr(G_APPPATH & "\Scripts\") Then
36:                     AlreadyInstalled = True
37:                     Exit For
38:                End If
39:            End If
40:       Next

42:       On Error GoTo Err
    
44:       objXML.Clear
    
46:       Set objNode = Nothing
47:       Set colSubNodes = Nothing
48:       Set colNodes = Nothing
          
50:   End If

      'Add plugin name to array
54:   ReDim Preserve m_sPlugins((m_iPlgCount))
55:   m_sPlugins(m_iPlgCount) = strName

57:   m_iPlgCount = m_iPlgCount + 1
          
59:   Exit Function
    
61:
Err:
63:   HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.AlreadyInstalled(" & strName & ")"
End Function

Private Sub XmlSave()
1:    On Error GoTo Err
2:    Dim intFF       As Integer
3:    Dim strTemp     As String
4:    Dim i           As Integer
    
6:     strTemp = G_APPPATH & "\Plugins\Plugins.xml"

8:     If g_objFileAccess.FileExists(strTemp) Then g_objFileAccess.DeleteFile strTemp
 
10:    If m_ErrLoading Then Exit Sub
 
12:    intFF = FreeFile

14:    Open strTemp For Append As intFF

16:      Print #intFF, "<Plugin>"

18:         For i = 0 To m_iPlgCount - 1
19:              Print #intFF, vbTab & "<Plugin Name=""" & m_sPlugins(i) & """" & " Value=""" & CStr(XMLUnescape(G_APPPATH & "\Scripts\")) & """ />"
20:         Next

22:      Print #intFF, "</Plugin>";

24:    Close intFF
    
26:   Exit Sub
    
28:
Err:
30:   HandleError Err.Number, Err.Description, Erl & "|" & "frmScript.XmlBooleanSave()"
End Sub

Private Sub Class_Initialize()
1:   Set lvwItems = frmHub.lvwPlugins.ListItems
End Sub

Private Sub Class_Terminate()
1:   Call XmlSave
2:   Set lvwItems = Nothing
3:   Set lvwItem = Nothing
4:   Erase m_sPlugins()
End Sub
