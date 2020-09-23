VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHTTPDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private WithEvents m_wskLoop    As Winsock
Attribute m_wskLoop.VB_VarHelpID = -1
Private m_frmParent             As frmSock
Private m_intIndex              As Integer
Private m_intFF                 As Integer
Private m_strData               As String
Private m_strHeader             As String
Private m_strFile               As String

Public Event OnError(ByVal lngNumber As Long, ByRef strDescription As String)
Public Event OnDownloaded(ByRef strHeader As String, ByRef strData As String)

Private Sub Class_Initialize()
    'Create our dummy form to get it's winsock
2:    Set m_frmParent = New frmSock
3:    Set m_wskLoop = m_frmParent.wskSock
End Sub

Private Sub Class_Terminate()
1:    On Error GoTo Err

    'Make sure it is closed
4:    If m_wskLoop.State Then
5:        m_wskLoop.Close
6:        DoEvents
7:    End If
    
    'Clear out winsock / form
10:    Set m_wskLoop = Nothing
11:    Set m_frmParent = Nothing
    
13:    Exit Sub
    
15:
Err:
16:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHTTPDownload_Terminate()"
End Sub

Public Property Let Host(ByRef strData As String)
1:    m_wskLoop.RemoteHost = strData
End Property

Public Property Let Port(ByRef lngData As Long)
1:    m_wskLoop.RemotePort = lngData
End Property

Public Property Let file(ByRef strData As String)
1:    m_wskLoop.Tag = strData
End Property

Public Property Get Data() As String
1:    Data = m_strData
End Property

Public Property Get Header() As String
1:    Header = m_strHeader
End Property

Public Property Get InUse() As Boolean
1:    InUse = m_wskLoop.State
End Property

Friend Sub Connect()
1:    m_wskLoop.Connect
End Sub

Private Sub m_wskLoop_Close()
1:    Dim lngPos      As Long
2:    Dim strLine     As String
    
4:    On Error GoTo Err

    'Make sure winsock is closed
    'If not, we get an infinite loop
8:    If m_wskLoop.State Then m_wskLoop.Close
    
    'Close file
11:    Close m_intFF
    
    'Read data
14:    m_strData = g_objFileAccess.ReadFile(m_strFile)
    
    'If not data, raise error and exit
17:    If LenB(m_strData) = 0 Then
18:        RaiseEvent OnError(-202, "No data downloaded")
19:        Exit Sub
20:    End If
    
22:    lngPos = InStrB(1, m_strData, vbTwoLine)
    
24:    If lngPos = 0 Then
        'Raise error because there is no header/data
26:        RaiseEvent OnError(-201, "No header seperator present")
27:    Else
28:        m_strHeader = LeftB$(m_strData, lngPos - 1)
29:        m_strData = MidB$(m_strData, lngPos + 4)
        
        'Check header for errors (404, etc)
32:        lngPos = InStrB(1, m_strHeader, " ") + 2
33:        strLine = MidB$(m_strHeader, lngPos, InStrB(lngPos, m_strHeader, " ") - lngPos)
        
        'If it's not numeric, then raise an error!
36:        If IsNumeric(strLine) Then
37:            lngPos = CLng(strLine)
            
39:            If lngPos = 200 Then
                'If no errors raise the downloaded event
41:                RaiseEvent OnDownloaded(m_strHeader, m_strData)
42:            Else
                'Raise event telling an error occured
44:                RaiseEvent OnError(lngPos, "Server error from header")
45:            End If
46:        Else
47:            RaiseEvent OnError(-200, "Error in processing header")
48:        End If
49:    End If
    
51:    Exit Sub
    
53:
Err:
    'RaiseEvent OnError(Err.Number, Err.Description)
55:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHTTPDownload::m_wskLoop_Close()"
End Sub

Private Sub m_wskLoop_Connect()
1:    Dim strTag As String
    
3:    On Error GoTo Err

5:    strTag = m_wskLoop.Tag

    'HTTP protocol, requesting file
8:    m_wskLoop.SendData "GET /" & strTag & " HTTP/1.1" & vbNewLine & _
                     "Host: " & m_wskLoop.RemoteHost & ":" & m_wskLoop.RemotePort & vbNewLine & _
                     "Connection: close" & vbTwoLine
                                        
12:    If LenB(strTag) Then _
        m_strFile = G_APPPATH & "\" & strTag _
    Else _
        m_strFile = G_APPPATH & "\" & m_wskLoop.RemoteHost & ".txt"
    
    'Delete file if needed
18:    If g_objFileAccess.FileExists(m_strFile) Then Kill m_strFile
    
20:    m_intFF = FreeFile
    
    'Open file for append
23:    Open m_strFile For Append As m_intFF
    
25:    Exit Sub
    
27:
Err:
    'RaiseEvent OnError(Err.Number, Err.Description)
29:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHTTPDownload::m_wskLoop_Connect()"
End Sub

Private Sub m_wskLoop_DataArrival(ByVal bytesTotal As Long)
1:    Dim strData As String
    
3:    On Error GoTo Err

    'Get data and append to file
6:    m_wskLoop.GetData strData, vbString
7:    Print #m_intFF, strData;
    
9:    Exit Sub
    
11:
Err:
    'RaiseEvent OnError(Err.Number, Err.Description)
13:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHTTPDownload::m_wskLoop_DataArrival()"
End Sub

Private Sub m_wskLoop_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
1:    m_wskLoop.Close
    'RaiseEvent OnError(Number, Description)
3:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHTTPDownload::m_wskLoop_Error()"
End Sub
