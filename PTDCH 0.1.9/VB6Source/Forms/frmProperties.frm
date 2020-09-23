VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmProperties 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "(Name) Code (Script)"
   ClientHeight    =   2655
   ClientLeft      =   165
   ClientTop       =   135
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin ComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2400
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   9313
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   3855
      Left            =   5640
      ScaleHeight     =   1678.632
      ScaleMode       =   0  'User
      ScaleWidth      =   312
      TabIndex        =   3
      Top             =   -120
      Visible         =   0   'False
      Width           =   30
   End
   Begin ComctlLib.ListView lvwProperties 
      Height          =   2295
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Property"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtValue 
      Height          =   2295
      Left            =   1900
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   30
      Width           =   3615
   End
   Begin VB.Image imgSplitter 
      Height          =   2280
      Left            =   1815
      MouseIcon       =   "frmProperties.frx":0000
      MousePointer    =   99  'Custom
      Top             =   30
      Width           =   100
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuProp 
         Caption         =   "Add"
         Index           =   0
      End
      Begin VB.Menu mnuProp 
         Caption         =   "Remover"
         Index           =   1
      End
      Begin VB.Menu mnuProp 
         Caption         =   "Save"
         Index           =   2
      End
      Begin VB.Menu mnuProp 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuProp 
         Caption         =   "Exit"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Cool FX Magnetic Windows
Private Magnetic As New clsMagneticWnd

Private m_strFile   As String
Private m_lngType   As Long

Private mIsMoving As Boolean ' Is splitter in motion?
Private Const mSplitLimit = 500 ' Minimum width for ListView and TextBox

Private Sub SizeControls(X As Single)
    ' Set sizes and locations for the movable controls
2:  On Error Resume Next

    'Set the ListView width
5:    If X < 500 Then X = 500
6:    If X > (Me.Width - 500) Then X = Me.Width - 500
    
8:    lvwProperties.Width = X
9:    imgSplitter.Left = X

    ' Set up the TextBox
12:    With txtValue
13:        .Left = X + 75
14:        .Width = Me.Width - (lvwProperties.Width + 220)
15:    End With

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub

Private Sub Form_Paint()
1: If g_objSettings.blSkin Then _
     PaintTileFormBackground Me, iResPic(g_objSettings.lngSkin)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
1:    PSave
End Sub

Public Property Let file(ByRef strData As String)
1:    Dim objXML      As clsXMLParser
2:    Dim objNode     As clsXMLNode
3:    Dim objNodes    As Collection
4:    Dim lvwItems    As ListItems
5:    Dim strName     As String
    
7:    On Error GoTo Err
    
    'Extract XML property file path
10:    m_strFile = strData
    
    Select Case m_lngType
        Case 0 'Script
12:            m_strFile = G_APPPATH & "\Scripts\" & LeftB$(strData, InStrB(1, strData, ".") - 1) & ".xml"
13:            Caption = strData & " Properties (Script)"
        Case 1 'Template
14:            m_strFile = G_APPPATH & "\Templates\" & LeftB$(strData, InStrB(1, strData, ".") - 1) & ".xml"
15:            Caption = strData & " Properties (Template)"
        Case 2 'Method
16:            m_strFile = G_APPPATH & "\Methods\" & LeftB$(strData, InStrB(1, strData, ".") - 1) & ".xml"
17:            Caption = strData & " Properties (Method)"
18:    End Select
    
20:    Set lvwItems = lvwProperties.ListItems
21:    Set objXML = New clsXMLParser
    
    'If the file exists, then load properties, else load defaults
24:    If g_objFileAccess.FileExists(m_strFile) Then _
        objXML.Data = g_objFileAccess.ReadFile(m_strFile) _
    Else _
        objXML.Data = g_objFileAccess.ReadFile(G_APPPATH & "Settings\DefaultProps.xml")
    
29:    objXML.Parse
    
    'Get nodes
32:    Set objNodes = objXML.Nodes(1).Nodes
    
    'Loop through them
35:    For Each objNode In objNodes
36:        lvwItems.Add(, , objNode.Name).Tag = objNode.Value
37:    Next
    
39:    lvwProperties_ItemClick lvwItems(1)
    
41:    Exit Property
    
43:
Err:
44:    HandleError Err.Number, Err.Description, Erl & "|" & "frmProperties.File(" & strData & ")"
End Property

Public Property Get file() As String
1:    file = m_strFile
End Property

Public Property Let PType(ByVal lngData As Long)
1:    m_lngType = lngData
End Property
Public Property Get PType() As Long
1:    PType = m_lngType
End Property

Public Sub PSave()
1:    Dim strTemp     As String
2:    Dim intFF       As Integer
3:    Dim lvwItem     As ListItem
4:    Dim lvwItems    As ListItems
    
6:    On Error GoTo Err
    
    'Delete file if it exists
9:    If g_objFileAccess.FileExists(m_strFile) Then g_objFileAccess.DeleteFile m_strFile
    
    'Save last node clicked
12:    Set lvwItems = lvwProperties.ListItems
    
    'If there are any properties, then save the last change
15:    If lvwItems.count Then
16:        Set lvwItem = lvwItems(1)
17:        lvwProperties_ItemClick lvwItem
18:        lvwItem.Selected = True
19:    End If
    
21:    intFF = FreeFile
    
    'Open XML file for append
24:    Open m_strFile For Append As intFF
    
26:    Print #intFF, "<Properties>"
    
    'Append all node values
29:    For Each lvwItem In lvwItems
30:        strTemp = lvwItem.Text
31:        Print #intFF, "<" & strTemp & ">" & XMLEscape(lvwItem.Tag) & "</" & strTemp & ">"
32:    Next
    
34:    Print #intFF, "</Properties>";
    
36:    Close intFF
    
38:    Exit Sub
    
40:
Err:
41:    HandleError Err.Number, Err.Description, Erl & "|" & "frmProperties.PSave()"
End Sub

Private Sub Form_Resize()

2:    On Error GoTo Err
      
4:    Dim formWidth As Long
    
6:    formWidth = Me.ScaleWidth

8:    lvwProperties.Height = Me.ScaleHeight - lvwProperties.Top - stBar.Height - 30

10:    With txtValue
          '.Width = formWidth - lvwProperties.Width - 120
12:        .Height = lvwProperties.Height
13:    End With

15:    With imgSplitter
16:        .Top = lvwProperties.Top + 30
17:        .Height = lvwProperties.Height - 30
18:        SizeControls .Left
19:    End With
    
21:   Exit Sub
    
23:
Err:
25:    HandleError Err.Number, Err.Description, Erl & "|" & "frmProperties_Resize()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
3:    frmHub.SetFocus
End Sub

' Use the PictureBox as a marker for the new split location
Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:  On Error Resume Next
2:  With imgSplitter
3:        picSplitter.Move .Left, .Top, .Width \ 3, .Height - 20
4:    End With
5:    picSplitter.Visible = True
6:    mIsMoving = True
End Sub


Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:  On Error Resume Next
2:    Dim sglPos As Single
3:    If mIsMoving Then
4:        sglPos = X + imgSplitter.Left
5:        If sglPos < mSplitLimit Then
6:            picSplitter.Left = mSplitLimit
7:        ElseIf sglPos > Me.Width - mSplitLimit Then
8:            picSplitter.Left = Me.Width - mSplitLimit
9:        Else
10:           picSplitter.Left = sglPos
11:       End If
12:   End If
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    SizeControls picSplitter.Left
2:    picSplitter.Visible = False
3:    mIsMoving = False
End Sub

Private Sub lvwProperties_ItemClick(ByVal Item As ComctlLib.ListItem)
1:    Static intLastIndex As Integer
    
3:    On Error Resume Next
    
    'If previous item was selected, them set data into it
6:    If intLastIndex Then _
        lvwProperties.ListItems(intLastIndex).Tag = txtValue.Text
        
    'Put value of new property into it
10:    txtValue.Text = Item.Tag
11:    intLastIndex = Item.Index
End Sub

Private Sub lvwProperties_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    If Button = 2 Then PopupMenu mnuEdit
End Sub

Private Sub mnuProp_Click(Index As Integer)
1:      Dim strName     As String
2:      Dim lvwItem     As ListItem
3:      Dim lvwItems    As ListItems
    
5:      On Error GoTo Err
    
        Select Case Index
           Case 0 'Add
               With frmMulti
10:                 .Caption = "Add Property"
11:                 .Label1.Caption = "Enter the name of the new property"
12:                 .Show vbModal, Me
13:                 strName = .txtStr.Text
               End With
15:            Set frmMulti = Nothing
               
17:            If LenB(strName) Then
18:               If InStrB(1, strName, " ") Then strName = Replace(strName, " ", "_")
19:               lvwProperties.ListItems.Add , , strName
20:            End If

           Case 1 'Remove
23:            Set lvwItems = lvwProperties.ListItems
            
25:            For Each lvwItem In lvwItems
26:                If lvwItem.Selected Then lvwItems.Remove lvwItem.Index
27:            Next
           Case 2 'Save
29:            PSave
           Case 4
31:            Unload Me
32:     End Select
    
34:  Exit Sub
    
36:
Err:
38:    HandleError Err.Number, Err.Description, Erl & "|" & "frmProperties.mnuPropList_Click(" & Index & ")"
End Sub

Private Sub stBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub
