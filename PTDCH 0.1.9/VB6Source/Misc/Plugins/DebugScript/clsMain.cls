VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------'
'  Copyright (C) fLaSh - Carlos Ferreira       '
'  E-Mail: Carlosferreiracarlos@hotmail.com    '
'----------------------------------------------'

'Option Explicit

'Set plugin properties
Private Const mName          As String = "DebugScript"
Private Const mVersion       As String = "0.1"
Private Const mAuthor        As String = "fLaSh"
Private Const mDescription   As String = "Active Debugger is intended to debug script by using it as run time debugger by print ing valuable information onto debugger's screen such as variable contents, current state of any object etc.."
Private Const mRelease       As Date = #11/18/2007 6:11:40 PM#
Private Const mComments      As String = "Enjoy!"
Private Const mPlgSetup      As Boolean = False 'Set True if availabed plugin settings

Dim objFrm As frmMain
Dim mDefautForeColor As Integer

Function LoadPlug() As Integer
    'This just allows the IDE to know if the plug-ins was loaded
    LoadPlug = 1
End Function

Function iPic() As IPictureDisp
    'This is the picture function. that sends back an IPictureDisp picture
    ' This picture is then used for the add-ins button.
    ' Change your picture in the resource file.
    ' Note mask color for the bitmap is RGB(255,0,255)
    Set iPic = LoadResPicture(101, vbResBitmap)
End Function

'Add the plug-in mame, version, autor, realse date and comments..
Public Property Get Name() As String
    Name = mName
End Property

Public Property Get Version() As String
    Version = mVersion
End Property

Public Property Get Author() As String
    Author = mAuthor
End Property

Public Property Get Description() As String
    Description = mDescription
End Property

Public Property Get Release() As String
    Release = mRelease
End Property

Public Property Get Comments() As String
    Comments = mComments
End Property

Public Property Get PlgSetup() As Boolean
    PlgSetup = mPlgSetup
End Property

Public Sub SubMain(mObject As Object)
    '
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical, mName & " - SubMain()"
    Err.Clear
End Sub

Public Sub LoadPlugin(Optional mObject As Object)
    On Error GoTo Err

    objFrm.Show
    'Init
    DebugPrint "Debugger Started..", , , 3
    
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical, mName & " - LoadPlugin()"
    Err.Clear
End Sub

Public Sub UnloadPlugin()
    On Error GoTo Err
    
    Call Unload(objFrm)

    Exit Sub
Err:
    MsgBox Err.Description, vbCritical, mName & " - UnloadPlugin()"
    Err.Clear
End Sub

Public Function Init(Optional sTitle As String = "Active Scripting Debugger Window", _
                     Optional iForeColor As Integer = 0, _
                     Optional iBackColor As Integer = 15, _
                     Optional sFontName As String = "Tahoma", _
                     Optional iFontSize As Integer = 9)
                     
    On Error GoTo Err
    
    If iForeColor > 15 Then iForeColor = 0
    If iBackColor > 15 Then iBackColor = 15
    
    objFrm.Caption = sTitle
    mDefautForeColor = iForeColor
    
    With objFrm.rtbDebug
        .BackColor = QBColor(iBackColor)
        .Font.Name = sFontName
        .Font.Size = iFontSize
    End With
    
  Exit Function
Err:
  MsgBox Err.Description, vbCritical, mName & " - Init()"
  Err.Clear
End Function

Public Function DebugPrint(sMsg As String, _
                  Optional sProced As String = "", _
                  Optional sComment As String = "", _
                  Optional iColor As Long = -1, _
                  Optional bTime As Boolean = True, _
                  Optional bBold As Boolean = False, _
                  Optional bUnderline As Boolean = False)
                  
    On Error GoTo Err
    
    Dim strTemp  As String
    
    If iColor > 15 Or iColor < 0 Then iColor = mDefautForeColor
    
    strTemp = sMsg
    If sProced <> "" Then strTemp = strTemp & "|" & sProced
    If sComment <> "" Then strTemp = strTemp & "|" & sComment & "|"
    
    With objFrm.rtbDebug
       If bTime Then
           .SelStart = Len(objFrm.rtbDebug.Text)
           .SelColor = QBColor(5)
           .SelText = "[" & Time & "] "
       End If
       If bBold Then .SelBold = True Else .SelBold = False
       If bUnderline Then .SelUnderline = True Else .SelUnderline = False
       .SelStart = Len(objFrm.rtbDebug.Text)
       .SelColor = QBColor(iColor)
       .SelText = strTemp & vbCrLf
    End With

    Exit Function
Err:
    MsgBox Err.Description, vbCritical, mName & " - DebugPrint()"
    Err.Clear
End Function

Public Function Clear()
    On Error GoTo Err
    
    objFrm.rtbDebug.Text = ""
    DebugPrint "Debugger Cleared..", , , 3
  
    Exit Function
Err:
    MsgBox Err.Description, vbCritical, mName & " - Clear()"
    Err.Clear
End Function

Private Sub Class_Initialize()
    On Error GoTo Err

    Set objFrm = New frmMain
    mDefautForeColor = 0
    
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical, mName & " - Class_Initialize()"
    Err.Clear
End Sub

Private Sub Class_Terminate()
    On Error GoTo Err

    Set objFrm = Nothing

    Exit Sub
Err:
    MsgBox Err.Description, vbCritical, mName & " - Class_Terminate()"
    Err.Clear
End Sub
