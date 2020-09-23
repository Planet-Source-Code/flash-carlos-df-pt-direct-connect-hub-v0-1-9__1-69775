VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCAccounts 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Convert Accounts for PTDCH database"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdInfo 
      Caption         =   "?"
      Height          =   300
      Left            =   4680
      TabIndex        =   15
      Top             =   900
      Width           =   300
   End
   Begin ComctlLib.ListView lvwNoErrors 
      Height          =   1935
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Account Nº"
         Object.Width           =   1412
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Password"
         Object.Width           =   2789
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Profile"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   12
      Tag             =   "1"
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Tag             =   "1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtCount 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtDir 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   900
      Width           =   3135
   End
   Begin VB.OptionButton optXMLtype 
      Caption         =   "PtokaX"
      Height          =   195
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Value           =   -1  'True
      Width           =   195
   End
   Begin VB.OptionButton optXMLtype 
      Caption         =   "YnHub"
      Height          =   195
      Index           =   0
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   195
   End
   Begin ComctlLib.ListView lvwWithErrors 
      Height          =   1335
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2355
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Account Nº"
         Object.Width           =   1412
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Error Description"
         Object.Width           =   2363
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Password"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Profile"
         Object.Width           =   1060
      EndProperty
   End
   Begin VB.Label lblCheck 
      BackStyle       =   0  'Transparent
      Caption         =   "PtokaX"
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   11
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblOption 
      BackStyle       =   0  'Transparent
      Caption         =   "YnHub"
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   10
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Select the type of database:"
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
      Left            =   1920
      TabIndex        =   9
      Top             =   0
      Width           =   3015
   End
   Begin VB.Line Line 
      BorderColor     =   &H00C0C0C0&
      Index           =   2
      X1              =   1440
      X2              =   4920
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   1440
      X2              =   1440
      Y1              =   720
      Y2              =   120
   End
   Begin VB.Line Line 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      X1              =   1440
      X2              =   1800
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "With Errors"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "No Errors"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Shape Shape 
      BorderColor     =   &H00C0C0C0&
      Height          =   1095
      Index           =   20
      Left            =   120
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Accounts Count:"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmCAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Convert accounts from the file *.xml database of the outhers Hub Software (YnHub or PtokaX) for PTDCH database ;-)
'Add accounts to PTDCH database very fast..
'and show details in List View ;-)
Private Sub cmdConvert_Click()
1:     Dim objXML          As clsXMLParser
2:     Dim objNode         As clsXMLNode
3:     Dim objSubNode      As clsXMLNode
4:     Dim colNodes        As Collection
5:     Dim colSubNodes     As Collection
    
7:     Dim strNick         As String
8:     Dim strPass         As String
9:     Dim lngProf         As Long
10:    Dim strRsult        As String
11:    Dim i               As Integer
       
13:    Dim lvwItems        As ListItems
14:    Dim lvwItem         As ListItem

16:    Dim sFileName       As String
17:    Dim sProfileNode    As String
18:    Dim sProfileType(2) As String
19:    Dim sPassType       As String
              
21:    On Error GoTo Err

23:    If optXMLtype(1).Value Then 'PtokaX
24:       sFileName = "\RegisteredUsers.xml"
25:       sProfileNode = "RegisteredUser"
26:       sPassType = "Password"
27:       sProfileType(0) = "0" 'Master
28:       sProfileType(1) = "1" 'Op
29:       sProfileType(2) = "2" 'Vip
30:    Else 'YnHub
31:       sFileName = "\accounts.xml"
32:       sProfileNode = "Account"
33:       sPassType = "Pass"
34:       sProfileType(0) = "Owner"
35:       sProfileType(1) = "OP"
36:       sProfileType(2) = "VIP"
37:    End If

39:    Set g_objFileAccess = New clsFileAccess
40:    Set objXML = New clsXMLParser
    
42:    If g_objFileAccess.FileExists(txtDir.Text & sFileName) Then
 
44:        objXML.Data = g_objFileAccess.ReadFile(txtDir.Text & sFileName)
45:        objXML.Parse
    
47:        Set colNodes = objXML.Nodes(1).Nodes
    
49:        txtCount.Text = ""
50:        lvwNoErrors.ListItems.Clear
51:        lvwWithErrors.ListItems.Clear
52:        i = 1
           
54:        For Each objNode In colNodes
             'Just in case...
56:           On Error Resume Next
              
58:           Set colSubNodes = objNode.Nodes
              Select Case objNode.Name
                Case sProfileNode
61:                    For Each objSubNode In colSubNodes
                        Select Case objSubNode.Name
                           Case "Nick"
64:                                strNick = objSubNode.Value
                           Case sPassType
66:                                strPass = objSubNode.Value
                           Case "Profile"
68:                             Select Case objSubNode.Value
                                   Case sProfileType(0)
70:                                     lngProf = 9 'Super Op
                                   Case sProfileType(1)
72:                                     lngProf = 6  'Op
                                   Case sProfileType(2)
74:                                     lngProf = 5  'Vip
                                   Case Else 'Reg or other
76:                                     lngProf = 3
                                End Select
                    
79:                             Select Case g_objRegistered.Add(strNick, strPass, lngProf, "PtokaX/Database")
                                   Case 0
81:                                     strRsult = "No error"
                                   Case 1
83:                                     strRsult = "Registered already"
                                   Case 2
85:                                     strRsult = "Name longer than 40 chars"
                                   Case 3
87:                                     strRsult = "Password longer than 20 chars"
                                End Select
                                
90:                             On Error GoTo Err
                                
93:                             If strRsult = "No error" Then
94:                                Set lvwItems = lvwNoErrors.ListItems
95:                                'Add listitem
96:                                Set lvwItem = lvwItems.Add(, , Val(i))
97:                                lvwItem.SubItems(1) = strNick
98:                                lvwItem.SubItems(2) = strPass
99:                                lvwItem.SubItems(3) = lngProf
100:                            Else
101:                               Set lvwItems = lvwWithErrors.ListItems
102:                               'Add listitem
103:                               Set lvwItem = lvwItems.Add(, , Val(i))
104:                               lvwItem.SubItems(1) = strNick
105:                               lvwItem.SubItems(2) = strRsult
106:                               lvwItem.SubItems(3) = strPass
107:                               lvwItem.SubItems(4) = lngProf
108:                            End If

110:                            i = i + 1

112:                         End Select
113:                    Next
114:                'Case Else
115:                 DoEvents
116:            End Select
117:        Next

119:        txtCount.Text = Val(i - 1)
       
121:        On Error GoTo Err
           
122:        objXML.Clear
        
124:       Set objSubNode = Nothing
125:       Set objNode = Nothing
126:       Set colSubNodes = Nothing
127:       Set colNodes = Nothing

129:       optXMLtype(0).Enabled = True
130:       optXMLtype(1).Enabled = True
131:       cmdConvert.Enabled = False
132:    End If
    
134:  Exit Sub

136:
Err:
138:  HandleError Err.Number, Err.Description, Erl & "|" & "frmCAccounts.cmdConvert_Click()"
End Sub

Private Sub cmdBrowse_Click()
       Dim cD          As New clsCommonDialog
       Dim strNovaPath As String
       Dim strMsg      As String
       Dim strFile     As String
       
6:     On Error GoTo Err

8:       If optXMLtype(1).Value Then
9:          strMsg = g_colMessages.Item("msgConvRegsDir") & "PtokaX (../PtokaX/cfg/)"
10:          strFile = "RegisteredUsers.xml"
11:       Else
12:          strMsg = g_colMessages.Item("msgConvRegsDir") & "YnHub (../YnHub/Settings/) "
13:          strFile = "accounts.xml"
14:       End If
   
16:       strNovaPath = cD.VBBrowseFolder(Me.hWnd, strMsg)
       
18:       If Trim(strNovaPath) = "" Then Exit Sub
19:       txtDir = strNovaPath & "\"
       
21:       If Not g_objFileAccess.FileExists(txtDir & strFile) Then
22:          cmdConvert.Enabled = False
23:          optXMLtype(0).Enabled = True
24:          optXMLtype(1).Enabled = True
25:          MsgBoxCenter Me, g_colMessages.Item("msgConvRegsNoXML") & "(" & strFile & ")", vbInformation
26:       Else
27:          cmdConvert.Enabled = True
28:          optXMLtype(0).Enabled = False
29:          optXMLtype(1).Enabled = False
30:       End If

32:  Exit Sub

34:
Err:
36:  HandleError Err.Number, Err.Description, Erl & "|" & "frmCAccounts.cmdBrowse_Click()"
End Sub

Private Sub cmdClose_Click()
1:    Unload Me
End Sub

Private Sub cmdInfo_Click()

2: MsgBoxCenter Me, "The conversion process was conceived for the following criteria:" & vbNewLine & _
         vbNewLine & _
         vbTab & "Type of the database in PtokaX (RegisteredUsers.xml)" & vbNewLine & _
         vbTab & "  Node Value: <RegisteredUser>" & vbNewLine & _
         vbTab & "    SubNode Value:" & vbNewLine & _
         vbTab & "      -Nick: " & "<Nick>Respective nick</Nick>" & vbNewLine & _
         vbTab & "      -Password: " & "<Password>Respective password</Password>" & vbNewLine & _
         vbTab & "      -Profile: " & vbNewLine & _
         vbTab & "          SuperOp: " & "<Profile>0</Profile>" & vbNewLine & _
         vbTab & "          Op: " & "<Profile>1</Profile>" & vbNewLine & _
         vbTab & "          Vip: " & "<Profile>2</Profile>" & vbNewLine & _
         vbTab & "          Reg: " & "Other Value" & vbNewLine & _
         vbNewLine & _
         vbTab & "Type of the database in YnHub (accounts.xml)" & vbNewLine & _
         vbTab & "  Node Value: <Account>" & vbNewLine & _
         vbTab & "    SubNode Value:" & vbNewLine & _
         vbTab & "      -Nick: " & "<Nick>Respective nick</Nick>" & vbNewLine & _
         vbTab & "      -Password: " & "<Pass>Respective password</Pass>" & vbNewLine & _
         vbTab & "      -Profile: " & vbNewLine & _
         vbTab & "          SuperOp: " & "<Profile>Owner</Profile>" & vbNewLine & _
         vbTab & "          Op: " & "<Profile>OP</Profile>" & vbNewLine & _
         vbTab & "          Vip: " & "<Profile>VIP</Profile>" & vbNewLine & _
         vbTab & "          Reg: " & "Other Value" & vbNewLine & vbNewLine & _
         "*Note: These criteria are used by defaut in the HubSoft.", _
         vbOKOnly Or vbInformation, "PTDCH - Info - Convert Accounts for PTDCH database"

End Sub

Private Sub Form_Load()
    On Error GoTo Err
      'Set Full Selected Rum
3:     LVOptFRow frmCAccounts.lvwNoErrors
4:     LVOptFRow frmCAccounts.lvwWithErrors
      
6:     Me.Caption = g_colMessages.Item("msgConvertRegs")
7:     Labels(0).Caption = g_colMessages.Item("msgConvRegsCount")
8:     Labels(1).Caption = g_colMessages.Item("msgConvRegsNoErr")
9:     Labels(2).Caption = g_colMessages.Item("msgConvRegsWithErr")
10:    lblTitle(0).Caption = g_colMessages.Item("msgConvRegsDBType")
11:    cmdBrowse.Caption = g_colMessages.Item("msgConvRegsBrowse")
12:    cmdConvert.Caption = g_colMessages.Item("msgConvRegsConv")
13:    cmdClose.Caption = g_colMessages.Item("msgClose")
14:    With lvwNoErrors
15:         .ColumnHeaders(1).Text = g_colMessages.Item("msgConvAccountN")
16:         .ColumnHeaders(2).Text = g_colMessages.Item("msgConvName")
17:         .ColumnHeaders(3).Text = g_colMessages.Item("msgConvPassword")
18:         .ColumnHeaders(4).Text = g_colMessages.Item("msgConvProfile")
19:    End With
20:    With lvwWithErrors
21:         .ColumnHeaders(1).Text = g_colMessages.Item("msgConvAccountN")
22:         .ColumnHeaders(2).Text = g_colMessages.Item("msgConvName")
23:         .ColumnHeaders(3).Text = g_colMessages.Item("msgConvErr")
24:         .ColumnHeaders(4).Text = g_colMessages.Item("msgConvPassword")
25:         .ColumnHeaders(5).Text = g_colMessages.Item("msgConvProfile")
26:    End With
27:  Exit Sub

29:
Err:
31:  HandleError Err.Number, Err.Description, Erl & "|" & "frmCAccounts.Form_Load()"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub

Private Sub Form_Paint()
1: If g_objSettings.blSkin Then _
     PaintTileFormBackground Me, iResPic(g_objSettings.lngSkin)
End Sub

Private Sub Form_Unload(Cancel As Integer)
1:   Set frmCAccounts = Nothing
End Sub

Private Sub Labels_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub

Private Sub lblTitle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub
