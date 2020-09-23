VERSION 5.00
Begin VB.Form frmUpDate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Update Check"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdButton 
      Caption         =   "Check"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Download"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Close"
      Height          =   375
      Index           =   0
      Left            =   5280
      TabIndex        =   1
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox txtVersion 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   1
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox txtVersion 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   0
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox txtInfo 
      Height          =   2055
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   960
      Width           =   6255
   End
   Begin VB.Line Line 
      BorderColor     =   &H00C0C0C0&
      Index           =   7
      X1              =   6480
      X2              =   6480
      Y1              =   600
      Y2              =   120
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last Version"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   6
      Top             =   0
      Width           =   2295
   End
   Begin VB.Line Line 
      BorderColor     =   &H00C0C0C0&
      Index           =   6
      X1              =   6120
      X2              =   6480
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line 
      BorderColor     =   &H00C0C0C0&
      Index           =   5
      X1              =   120
      X2              =   480
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line 
      BorderColor     =   &H00C0C0C0&
      Index           =   4
      X1              =   120
      X2              =   120
      Y1              =   600
      Y2              =   120
   End
   Begin VB.Line Line 
      BorderColor     =   &H00C0C0C0&
      Index           =   3
      X1              =   120
      X2              =   6480
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Version"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   5
      Top             =   0
      Width           =   2295
   End
   Begin VB.Line Line 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      X1              =   120
      X2              =   480
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   120
      X2              =   120
      Y1              =   3120
      Y2              =   840
   End
   Begin VB.Line Line 
      BorderColor     =   &H00C0C0C0&
      Index           =   2
      X1              =   120
      X2              =   6600
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "History"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   720
      Width           =   3015
   End
End
Attribute VB_Name = "frmUpDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_objUpdate      As clsHTTPDownload
Attribute m_objUpdate.VB_VarHelpID = -1
Private strURL                      As String
Private m_Notific                   As Boolean

Public Sub Notific(bNotific As Boolean)
    On Error GoTo Err
    
3:     m_Notific = bNotific
     
5:     Set m_objUpdate = New clsHTTPDownload
     
       'Prepare update class
8:     m_objUpdate.Host = "127.0.0.1" '/www.ptdch.com"
9:     m_objUpdate.Port = 80
10:    m_objUpdate.file = "ptdch_v.xml"
     
12:    m_objUpdate.Connect
     
14:  Exit Sub

16:
Err:
17:  HandleError Err.Number, Err.Description, Erl & "|" & "frmUpDate.Notific()"
End Sub

Private Sub cmdButton_Click(Index As Integer)
1:   On Error GoTo Err
     Select Case Index
        Case 0
4:           Unload Me
        Case 1 'Download
6:            g_objFunctions.ShellExec strURL
        Case 2 'Check for updates
8:          m_Notific = False
9:          If m_objUpdate.InUse Then
10:             MsgBoxCenter Me, g_colMessages.Item("msgUpdating"), vbOKOnly, g_colMessages.Item("msgUpdate")
11:          Else
12:             m_objUpdate.Connect
13:          End If
      End Select
15:  Exit Sub

16:
Err:
18:  HandleError Err.Number, Err.Description, Erl & "|" & "frmUpDate.cmdButton_Click(" & Index & ")"
End Sub

Private Sub Form_Load()

2:   On Error GoTo Err

4:     txtVersion(0).Text = vbVersion
5:     Set m_objUpdate = New clsHTTPDownload
     'Prepare update class
7:     m_objUpdate.Host = "127.0.0.1" '/www.ptdch.com"
8:     m_objUpdate.Port = 80
9:     m_objUpdate.file = "ptdch_version.xml"
    
11:  Exit Sub

13:
Err:
14:  HandleError Err.Number, Err.Description, Erl & "|" & "frmUpDate.Form_Load()"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub

Private Sub Form_Paint()
1:  If g_objSettings.blSkin Then _
         PaintTileFormBackground Me, iResPic(g_objSettings.lngSkin)
End Sub

Private Sub Form_Unload(Cancel As Integer)
1:    Set m_objUpdate = Nothing
2:    Set frmUpDate = Nothing
End Sub

'------------------------------------------------------------------------------
'Update events
'------------------------------------------------------------------------------

Private Sub m_objUpdate_OnDownloaded(strHeader As String, strData As String)
1:    Dim objXML          As clsXMLParser
2:    Dim objNode         As clsXMLNode
3:    Dim strVersion      As String
5:    Dim strCaption      As String
6:    Dim strMessage      As String
7:    Dim datRelease      As Date
    
9:    On Error GoTo Err

11:    Set objXML = New clsXMLParser
    
       'Set data and parse
14:    objXML.Data = strData
15:    objXML.Parse
       
17:    strURL = ""
       
       'Loop through to find properties
20:    For Each objNode In objXML.Nodes(1).Nodes
        'Find out which node it is
        Select Case objNode.Name
            Case "Version": strVersion = objNode.Value
            Case "ReleaseDate": datRelease = CDate(objNode.Value)
            Case "URL": strURL = objNode.Value
            Case "Caption": strCaption = objNode.Value
            Case "Message": strMessage = objNode.Value
28:        End Select
29:    Next
    
31:    txtVersion(1).Text = strVersion
32:    txtInfo.Text = strMessage
       
       'Compare release dates
35:    If DateDiff("s", vbReleaseDate, datRelease) > 0 Then
36:       If Not m_Notific Then
37:           cmdButton(1).Enabled = True
38:       Else
             'Cut message to correct length
40:          If LenB(strMessage) > 1600 Then _
                  strMessage = LeftB$(strMessage, 1600) & "...(continued)" & vbTwoLine _
                      Else _
                         strMessage = strMessage & vbTwoLine
             'Shell the url open if they want the update
45:          If MsgBox(strMessage & g_colMessages.Item("msgDownload"), vbYesNo Or vbQuestion, strCaption) = vbYes Then
46:             g_objFunctions.ShellExec strURL
47:          End If
48:          Call Unload(Me)
49:       End If
50:    Else
          'No new updates
52:       If Not m_Notific Then
53:          cmdButton(1).Enabled = False
54:       Else
55:          MsgBox g_colMessages.Item("msgMostRecent"), vbOKOnly Or vbInformation, g_colMessages.Item("msgNoUpdate")
56:          Call Unload(Me)
57:       End If
58:    End If
    
60:    Exit Sub
    
62:
Err:
64:    HandleError Err.Number, Err.Description, Erl & "|" & "frmUpDate.m_objUpdate_OnDownloaded(""" & strHeader & """, """ & strData & """)"
End Sub

Private Sub m_objUpdate_OnError(ByVal lngNumber As Long, strDescription As String)
1:    MsgBox Replace(Replace(g_colMessages.Item("msgDownloadError"), "%[number]", lngNumber), "%[description]", strDescription), vbOKOnly Or vbCritical, g_colMessages.Item("msgUpdateError")
End Sub
