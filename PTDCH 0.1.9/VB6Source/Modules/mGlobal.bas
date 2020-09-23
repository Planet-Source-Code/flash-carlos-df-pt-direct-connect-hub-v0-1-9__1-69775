Attribute VB_Name = "mGlobal"
Option Explicit

'Compiler conditions

'Debug mode - It means that it will print messages to the VB IDE debug window
'             Still is executed, even if compiled, so it is CPU friendly if
'             it is turned off when compiling
#Const DEBUG_MODE = False

'API calls
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long

'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long


Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
   (ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
    
'API Move Form ///////////////////////////////////////////////////////////////////////
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
   
'Public API for Sintax Coloriong /////////////////////////////////////////////////////
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Const EM_LINEINDEX = &HBB
Public lLineTracker             As Long
Public bDirty                   As Boolean
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'Get Memory Status
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

'API Stuff
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

'Constants
Public Const vbSVNVersion       As String = " - Beta 3d"
Public Const vbVersion          As String = "0.1.9"

#If SVN Then
    Private Const vbBeta            As String = "Debug"
#Else
    Private Const vbBeta            As String = vbNullString
#End If

Public Const vbLock             As String = "EXTENDEDPROTOCOLDEFDEFDEFDEFDEFDEFDEF Pk=DDCH" & vbVersion & vbSVNVersion & "DEFDEF"
'Public Const vbKey              As String = "4ÑÀ° A Ñ±±ÀÀ0€0 0 0 0 0 0 0"
Public Const vbWelcome          As String = "This hub is running V. " & vbVersion & " of the PTDCH produced by fLaSh (UpTime: %[UpTime])|"
Public Const vbChar5            As String = ""
Public Const vbChar160          As String = " "
Public Const vbTwoLine          As String = vbNewLine & vbNewLine
Public Const vbPartialClassList As String = "2 = Mentored" & vbNewLine & "3 = Registered" & vbNewLine & "4 = Invisible" & vbNewLine & "5 = VIP" & vbNewLine & "6 = Operator" & vbNewLine & "7 = Invisible Operator" & vbNewLine & "8 = Super Operator" & vbNewLine & "9 = Invisible Super Operator" & vbNewLine & "10 = Admin" & vbNewLine & "11 = Invisible Admin"
Public Const vbScriptConst      As String = "Const vbVersion = " & vbVersion & ":Const vbBeta = """ & vbBeta & """"
Public Const vbSFC              As Long = 26
Public Const vbReleaseDate      As Date = #12/16/2007 7:31:00 PM#

Public Const CHR_CR            As Integer = 13
Public Const CHR_LF            As Integer = 10
Public Const CHR_TAB           As Integer = 9
Public Const CHR_SPACE         As Integer = 32
Public Const CHR_DQUOTE        As Integer = 34

'-- Script function constant script event boolean array identifiers
Public Const vbSMain                            As Long = 0
Public Const vbSDataArrival                     As Long = 1
Public Const vbSAttemptedConnection             As Long = 2
Public Const vbSUserConnected                   As Long = 3
Public Const vbSRegConnected                    As Long = 4
Public Const vbSOpConnected                     As Long = 5
Public Const vbSUserQuit                        As Long = 6
Public Const vbSStartedServing                  As Long = 7
Public Const vbSSysTrayDoubleClick              As Long = 8
Public Const vbSAddedRegisteredUser             As Long = 9
Public Const vbSwskScript_Close                 As Long = 10
Public Const vbSwskScript_Connect               As Long = 11
Public Const vbSwskScript_ConnectionRequest     As Long = 12
Public Const vbSwskScript_DataArrival           As Long = 13
Public Const vbSwskScript_Error                 As Long = 14
Public Const vbStmrScriptTimer_Timer            As Long = 15
Public Const vbSAddedPermBan                    As Long = 16
Public Const vbSStartedRedirecting              As Long = 17
Public Const vbSStoppedServing                  As Long = 18
Public Const vbSMouseOverSysTray                As Long = 19
Public Const vbSMassMessage                     As Long = 20
Public Const vbSUnloadMain                      As Long = 21
Public Const vbSError                           As Long = 22
Public Const vbSTimeout                         As Long = 23
Public Const vbSRemovedRegisteredUser           As Long = 24
Public Const vbSCustComArrival                  As Long = 25
Public Const vbSPreDataArrival                  As Long = 26
Public Const vbSFailedConf                      As Long = 27

'Flood protection constants
Public Const vbFWMyINFO         As Long = 5
Public Const vbFWGetNickList    As Long = 35
Public Const vbFWActiveSearch   As Long = 35
Public Const vbFWPassiveSearch  As Long = 3
Public Const vbFWMilliseconds   As Long = 10000

'Public variables
Public G_APPPATH                As String
Public G_ERRORFILE              As Integer

Public G_CUNLOAD                As Boolean

#If SVN Then
    Public G_LOGPATH            As String
#End If

'Public objects
Public g_objFileAccess          As clsFileAccess
Public g_objFunctions           As clsFunctions
Public g_colUsers               As clsHub
Public g_objIPBans              As clsIPBans
Public g_objRegistered          As clsRegistered
Public g_objSettings            As clsSettings
Public g_objStatus              As clsStatus
'Public g_objSysTray             As clsSysTray
Public g_colCommands            As clsCommands
Public g_colSWinsocks           As Collection
Public g_colSVariables          As Collection
'Public g_colStatements          As Collection
Public g_objRegExps             As clsRegExps
Public g_colLanguages           As Collection
'***PLAN***
Public g_colScheduler           As clsPlan
'Public g_colScheduler            As clsCommands
'***PLAN END***

Public g_colMessages            As clsDictionary

Public Highlighter              As clsYHighlighter

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(32) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(32) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

'Enums
Public Enum enuState
    Disconnected = -1
    Wait_Key = 0
    Wait_Validate = 1
    Wait_Pass = 2
    Wait_PassPM = 3
    Wait_Info = 4
    Logged_In = 5
End Enum

Public Enum enuClass
    Locked = -1
    Unknown = 0
    Normal = 1
    Mentored = 2
    Registered = 3
    Invisible = 4
    vip = 5
    Op = 6
    InvisibleOp = 7
    SuperOp = 8
    InvisibleSuperOp = 9
    Admin = 10
    InvisibleAdmin = 11
End Enum

'Scheduler
'Public Enum enuAddType
'    Year = 6
'    Month = 5
'    week = 4
'    Day = 3
'    Hour = 2
'    Minute = 1
'    once = 0    <-might not be needed...
'End Enum
'Scheduler

Public Enum enuAlert
    MinShare = 0
    FakeTag = 1
    MinSlots = 2
    HSRatio = 3
    BSRatio = 4
    MaxHubs = 5
    DCppversion = 6
    NMDCVersion = 7
    NoTag = 8
    FakeShare = 9
    MaxShare = 10
    MaxSlots = 11
    Socks5 = 12
    PassiveMode = 13
    NoCOClients = 14
End Enum

Public Enum enuOpenFileMode
    vbRandom = 0
    vbInput = 1
    vbOutput = 2
    vbAppend = 3
    vbBinary = 4
End Enum

Public Type Highlighter
  StyleBold(127) As Long
  StyleItalic(127) As Long
  StyleUnderline(127) As Long
  StyleVisible(127) As Long
  StyleEOLFilled(127) As Long
  StyleFore(127) As Long
  StyleBack(127) As Long
  StyleSize(127) As Long
  StyleFont(127) As String
  StyleName(127) As String
  Keywords(7) As String
  strFilter As String
  strComment As String
  strName As String
  iLang As Long
End Type
Public G_Highlighters() As Highlighter
Public sciMain()         As clsYScintilla

Private m_lngTaskbarMsg         As Long
Private m_lngPrevProc           As Long

'Log messages to the user
Public Sub AddLog(strMsg As String, _
                  Optional intColor As Integer, _
                  Optional bTime As Boolean = True, _
                  Optional bBold As Boolean = False, _
                  Optional bUnderline As Boolean = False)
1: On Error GoTo Err

3:   With frmHub
4:        If bTime Then
5:           .rtbLog.SelStart = Len(.rtbLog)
6:           .rtbLog.SelColor = QBColor(5)
7:           .rtbLog.SelText = "[" & Time & "] "
8:        End If
9:        If bBold Then .rtbLog.SelBold = True Else .rtbLog.SelBold = False
10:       If bUnderline Then .rtbLog.SelUnderline = True Else .rtbLog.SelUnderline = False
11:       .rtbLog.SelStart = Len(.rtbLog)
12:       .rtbLog.SelColor = QBColor(intColor)
13:       .rtbLog.SelText = strMsg & vbCrLf
14:  End With

16: Exit Sub
17:
Err:
19: HandleError Err.Number, Err.Description, Erl & "|" & "mGlobal.AddLog()"
End Sub

Public Sub HandleError(ByRef lngNumber As Long, ByRef strDescription As String, ByRef strMethod As String, Optional ByRef lngDLLError As Long)
1:    Dim strError As String
2:    Dim i As Long
3:    On Error GoTo Err
      
      'Error log format :
      'Date-Time|Method|Number|DLLError|Description|Version|Beta|

      'Prevent error number '0' from being logged
9:    If lngNumber Then
         'Add beta version if it is a beta
11:      strError = Now & "|" & UTCDate & "|" & strMethod & "|" & lngNumber & "|" & lngDLLError & "|" & strDescription & "|" & vbVersion & "|" & vbBeta & "|" & vbSVNVersion & "|"
12:      AddLog "Error: " & UTCDate & "|" & strMethod & "|" & lngNumber & "|" & lngDLLError & "|" & strDescription, 4, False, True, False
         'Print to Debug window if in debug mode
         #If DEBUG_MODE Then
15:            Debug.Print strError
         #End If
        
          'Print to error log
19:       Print #G_ERRORFILE, strError
20:    End If

22:    Err.Clear
    
24:    Exit Sub
25:
Err:
    #If DEBUG_MODE Then
28:        Debug.Print Now & "|basGlobal.HandleError()|" & Err.Number & "|" & Err.Description & "|" & Err.LastDllError
    #End If
    
31: Err.Clear
33: Resume Next
End Sub

Public Function UTCDate(Optional ByVal strRef As String) As Date
1:    Dim k As TIME_ZONE_INFORMATION
      On Error GoTo Err
    'Get time zone difference
4:    GetTimeZoneInformation k
    
    'If a date is specified, then use that one, else use current
7:    If LenB(strRef) Then _
        UTCDate = DateAdd("n", k.Bias, CDate(strRef)) _
    Else _
        UTCDate = DateAdd("n", k.Bias, Now)
    Exit Function
Err:
12:  HandleError Err.Number, Err.Description, Erl & "|mGlobal.UTCDate()"
End Function

Public Function IIfLng(ByVal Expression As Boolean, ByRef TruePart As Long, ByRef FalsePart As Long) As Long
1:    If Expression Then IIfLng = TruePart Else IIfLng = FalsePart
End Function

Public Function XMLUnescape(ByRef strData As String) As String
1:    On Error GoTo Err
    
3:    Dim lngPos As Long
    
5:    XMLUnescape = strData
    
7:    If LenB(XMLUnescape) Then
8:        lngPos = InStrB(1, XMLUnescape, "&")
        
        'If there is a & in the string, that is where we should start searching
11:        If lngPos Then
            'Make sure there is a semi colon, telling us there may be escape sequences
13:            If InStrB(lngPos, XMLUnescape, ";") Then
                'Escape various illegal characters
15:                If InStrB(lngPos, XMLUnescape, "&lt;") Then XMLUnescape = Replace(XMLUnescape, "&lt;", "<")
16:                If InStrB(lngPos, XMLUnescape, "&gt;") Then XMLUnescape = Replace(XMLUnescape, "&gt;", ">")
17:                If InStrB(lngPos, XMLUnescape, "&quot;") Then XMLUnescape = Replace(XMLUnescape, "&quot;", """")
18:                If InStrB(lngPos, XMLUnescape, "&apos;") Then XMLUnescape = Replace(XMLUnescape, "&apos;", "'")
19:                If InStrB(lngPos, XMLUnescape, "&amp;") Then XMLUnescape = Replace(XMLUnescape, "&amp;", "&")
20:            End If
21:        End If
22:    End If
    
24:    Exit Function
    
26:
Err:
27:    HandleError Err.Number, Err.Description, Erl & "|" & "basGlobal.XMLUnescape(" & strData & ")"
End Function

Public Function XMLEscape(ByRef strData As String) As String
1:    On Error GoTo Err
    
3:    XMLEscape = strData
    
    'Check for the illegal characters
6:    If InStrB(1, XMLEscape, "&") Then XMLEscape = Replace(XMLEscape, "&", "&amp;")
7:    If InStrB(1, XMLEscape, "<") Then XMLEscape = Replace(XMLEscape, "<", "&lt;")
8:    If InStrB(1, XMLEscape, ">") Then XMLEscape = Replace(XMLEscape, ">", "&gt;")
9:    If InStrB(1, XMLEscape, """") Then XMLEscape = Replace(XMLEscape, """", "&quot;")
10:   If InStrB(1, XMLEscape, "'") Then XMLEscape = Replace(XMLEscape, "'", "&apos;")

12:    Exit Function
    
14:
Err:
15:    HandleError Err.Number, Err.Description, Erl & "|" & "basGlobal.XMLEscape(" & strData & ")"
End Function

Public Function GetByte(ByVal lngData As Long) As Byte
1:    GetByte = CByte(lngData And 255)
End Function

Public Function DebugUser(ByRef objUser As clsUser) As String
1:    If ObjPtr(objUser) Then DebugUser = "[""" & objUser.sName & """," & objUser.bOperator & "," & objUser.iWinsockIndex & ",""" & objUser.Supports & """,""" & objUser.sMyInfoString & """]"
End Function

Public Sub SetTaskbarMsg(ByVal lngProc As Long, ByVal lngMsg As Long)
1:    m_lngTaskbarMsg = lngMsg
2:    m_lngPrevProc = lngProc
End Sub

Public Function GenTempFile() As String
1:    Do
2:        Randomize GetTickCount
3:        GenTempFile = G_APPPATH & "\T" & GetTickCount & Rnd & ".tmp"
4:    Loop While g_objFileAccess.FileExists(GenTempFile)
End Function

Public Function TrueTrim(ByRef strString As String) As String
    '------------------------------------------------------------------
    'Purpose:   To trim any kind of whitespace from the beginning and
    '           and then end of a string. Whitespace includes spaces,
    '           tabs, carriage returns and line feeds
    '
    'Params:
    '           strString:      String to remove leading and trailing
    '                           whitespace from
    '
    'Returns:
    '           Copy of strString without trailing/leading whitespace
    '------------------------------------------------------------------

14:    Dim arr_intChars()      As Integer
15:    Dim i                   As Long
16:    Dim lngStart            As Long
17:    Dim lngEnd              As Long
       On Error GoTo Err
    'Get length of string
20:    lngEnd = Len(strString) - 1

    'Make sure there is something to trim
23:    If lngEnd >= 0 Then
        'Set start to first character
25:        lngStart = 1

        'Open character array on string
28:        OpenChrArr arr_intChars, strString

        'Find position of first non-whitespace character
31:        For i = 0 To lngEnd
            Select Case arr_intChars(i)
                Case CHR_SPACE, CHR_TAB, CHR_LF, CHR_CR
32:                    lngStart = lngStart + 1
                Case Else
33:                    Exit For
34:            End Select
35:        Next

        'Find position of last non-whitespace character
38:        For i = lngEnd To lngStart Step -1
            Select Case arr_intChars(i)
                Case CHR_SPACE, CHR_TAB, CHR_LF, CHR_CR
39:                    lngEnd = lngEnd - 1
                Case Else
40:                    Exit For
41:            End Select
42:        Next

        'Close character array
45:        CloseChrArr arr_intChars

        'Extract trimmed string
48:        TrueTrim = Mid$(strString, lngStart, lngEnd - lngStart + 2)
49:    End If

51:    Exit Function

Err:
54:    HandleError Err.Number, Err.Description, Erl & "|mGlobal.TrueTrim()"
End Function

Public Function ValidIP(ByVal strIPAddress As String) As Boolean
'Function to check if IP is valid --> if it's not higher than 255.255.255.255
2:  Dim sArray As Variant
  
4:  On Error GoTo Err
    
6:    sArray = Split(strIPAddress, ".")
7:      If sArray(0) > 255 Or sArray(1) > 255 Or sArray(2) > 255 Or sArray(3) > 255 Then
8:        ValidIP = False
      Else
10:        ValidIP = True
      End If

Exit Function
Err:
11:  ValidIP = False
End Function

Public Function ProcessorInfo() As String
    On Error GoTo Err
2:  Dim X As Object
3:  Dim Y As String
4:  Dim ReadKey As String

6:  On Error Resume Next

8:    ReadKey = ("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\0\Processornamestring")

10:   Set X = CreateObject("wscript.shell")
11:   Y = X.regread(ReadKey)
      
13:   Y = Replace(Y, "  ", "")
14:   ProcessorInfo = Y

16:   Exit Function

Err:
19:   HandleError Err.Number, Err.Description, Erl & "|mGlobal.ProcessorInfo()"
End Function

Public Function CharCount(str As String, Char As String) As Long
'Get character count in a string
2:    CharCount = UBound(Split(LCase(str), LCase(Char)))
End Function

'Formate digits ex: 12:1:7 for --> 12:01:07
Public Function strZero(ByVal strValor As String, ByVal bytComprimento As Byte)
1:   If Len(strValor) <= bytComprimento Then
2:      strZero = String(bytComprimento - Len(strValor), "0") & strValor
3:   Else
4:      strZero = strValor
5:   End If
End Function

'Pause the app without freezing it ('Sleep' freezes the app)
Public Function Pause(HowLong As Long)
1:  Dim Start&
2:  Start = GetTickCount()
3:  Do
4:    DoEvents
5:  Loop Until Start + HowLong < GetTickCount
End Function

Public Function frmMove(frm As Form)
1:  Call ReleaseCapture
2:  Call SendMessage(frm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Function

Public Function GetAppVersion() As String
1:  GetAppVersion = CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
End Function

Public Function HubUpTime() As String
1:  On Error GoTo Err

3:   Dim iMonths, iWeeks As Integer, iDays As Integer, iHours As Integer, lMinutes As Long, lSeconds As Long
4:   Dim lCurrTime As Long
5:   Dim strTmp As String
   
7:    lCurrTime = DateDiff("s", CDate(frmHub.ServingDate), DateTime.Now)
8:    lSeconds = lCurrTime Mod 60
9:    lMinutes = (lCurrTime \ 60) Mod 60
10:   iHours = (lCurrTime \ 3600) Mod 24
11:   iDays = (lCurrTime \ 86400) Mod 7
12:   iWeeks = (lCurrTime \ 604800) Mod 4
13:   iMonths = (lCurrTime \ 604800)
      
15:   If iMonths > 0 Then strTmp = strTmp & iMonths & IIf(iMonths = 1, " month, ", " months, ")
16:   If iWeeks > 0 Then strTmp = strTmp & iWeeks & IIf(iWeeks = 1, " week, ", " weeks, ")
17:   If iDays > 0 Then strTmp = strTmp & iDays & IIf(iDays = 1, " day, ", " days, ")
      
19:   If iHours > 0 Then strTmp = strTmp & iHours & IIf(iHours = 1, " hour, ", " hours, ")
20:   If lMinutes > 0 Then strTmp = strTmp & lMinutes & IIf(lMinutes = 1, " minute and ", " minutes and ")
   
22:   strTmp = strTmp & lSeconds & IIf(lSeconds = 1, " second", " seconds")
   
24:   HubUpTime = strTmp

26:   Exit Function
    
28:
Err:
30:    HandleError Err.Number, Err.Description, Erl & "|" & "mGlobal.HubUpTime()"
End Function
