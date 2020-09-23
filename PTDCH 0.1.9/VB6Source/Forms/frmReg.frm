VERSION 5.00
Begin VB.Form frmReg 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Register user"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Add"
   Begin VB.CommandButton cmdClipboard 
      Caption         =   "Clipboard"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdReg 
      Caption         =   "Add"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtInfo 
      Height          =   1575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   1920
      Width           =   3855
   End
   Begin VB.ComboBox cmbClass 
      Height          =   315
      ItemData        =   "frmReg.frx":0000
      Left            =   720
      List            =   "frmReg.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox txtPass 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      MaxLength       =   20
      TabIndex        =   3
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      MaxLength       =   40
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Class"
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   4
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the password"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the name tou want to register"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub InicializeReg()
1:  On Error GoTo Err
2:   If Me.Tag = CStr("Add") Then
3:       Me.Caption = g_colMessages.Item("msgRegUser")
4:       Labels(0).Caption = g_colMessages.Item("msgEnterRegName")
5:       Labels(1).Caption = g_colMessages.Item("msgEnterPass")
6:       Labels(2).Caption = g_colMessages.Item("msgEnterClass")
7:       cmdReg.Caption = g_colMessages.Item("msgAdd")
8:   ElseIf Me.Tag = CStr("Edit") Then
9:       Me.Caption = g_colMessages.Item("msgEditRegged")
10:      Labels(0).Caption = g_colMessages.Item("msgRegUser")
11:      Labels(1).Caption = g_colMessages.Item("msgEnterNewPass")
12:      Labels(2).Caption = g_colMessages.Item("msgEnterClass")
13:      cmdReg.Caption = g_colMessages.Item("msgEdit")
         txtName.Enabled = False
14:  ElseIf Me.Tag = CStr("Rename") Then
16:      Me.Caption = g_colMessages.Item("msgRenameUser")
17:      Labels(0).Caption = g_colMessages.Item("msgEnterNewName")
18:      Labels(1).Caption = g_colMessages.Item("msgEnterPass")
19:      Labels(2).Caption = g_colMessages.Item("msgEnterClass")
20:      cmdReg.Caption = g_colMessages.Item("msgRemame")
21:      cmbClass.Enabled = False
22:      txtPass.Enabled = False
23:  Else: GoTo Err
24:  End If

26:  cmdClose.Caption = g_colMessages.Item("msgClose")
27:  cmdClipBoard.Caption = g_colMessages.Item("msgClipboard")

29:  Exit Sub
    
31:
Err:
33:    HandleError Err.Number, Err.Description, Erl & "|" & "frmReg.InicializeReg()"
End Sub

Private Sub cmdClipBoard_Click()
1:  Clipboard.Clear
2:  Clipboard.SetText txtInfo.Text
End Sub

Private Sub cmdClose_Click()
1:  Unload Me
End Sub

Private Sub cmdReg_Click()
1:  On Error GoTo Err
   
2:    Dim lClass As Long
3:    Dim sTxt As String
4:    Dim sType As String
               
6:      If txtName.Text = "" Then
7:         MsgBoxCenter Me, Labels(0).Caption, vbInformation
8:         Exit Sub
9:      End If
    
11:     If txtPass.Text = "" Then
12:        MsgBoxCenter Me, Labels(1).Caption, vbInformation
13:        Exit Sub
14:     End If
    
16:     Select Case cmbClass.Text
           Case "2 = Mentored": lClass = 2
           Case "3 = Registered": lClass = 3
           Case "4 = Invisible": lClass = 4
           Case "5 = VIP": lClass = 5
           Case "6 = Operator": lClass = 6
           Case "7 = Invisible Operator": lClass = 7
           Case "8 = Super Operator": lClass = 8
           Case "9 = Invisible Super Operator": lClass = 9
           Case "10 = Admin": lClass = 10
           Case "11 = Invisible Admin": lClass = 11
        End Select
   
29      If Me.Tag = "Add" Then
            Select Case g_objRegistered.Add(txtName.Text, txtPass.Text, lClass)
               Case 0 'No error
               Case 1 'Registered already
33:              MsgBoxCenter Me, txtName & g_colMessages.Item("msgAlreadyRegged"), vbInformation
34:              Exit Sub
               Case 2 'Name longer than 40 chars
                '  MsgBox g_colMessages.Item("msgInvalidRegName")
                '  Exit Sub
               Case 3 'Password longer than 20 chars
                ' MsgBox g_colMessages.Item("msgInvalidPass")
                ' Exit Sub
            End Select
42:         sType = g_colMessages.Item("msgRegAdd")  '"Register add at: "
43:     ElseIf Me.Tag = "Edit" Then
44:         g_objRegistered.Edit txtName.Text, txtPass.Text, lClass
45:         sType = g_colMessages.Item("msgRegUpdate") '"Register updated at: "
46:     ElseIf Me.Tag = "Rename" Then
47:         g_objRegistered.Rename txtName.Tag, txtName.Text
48:         sType = g_colMessages.Item("msgRegUpdate") '"Register updated at: "
49:     End If
         
51:      sTxt = "PT DC Hub " & vbVersion & vbNewLine & vbNewLine & _
                sType & Now & vbNewLine & _
                "Nick: " & txtName.Text & vbNewLine & _
                "Password: " & txtPass.Text & vbNewLine & _
                "Status: " & cmbClass.Text
56:      txtInfo.Text = sTxt
    
58:      cmdClipBoard.Enabled = True
    
60:    Exit Sub
    
61:
Err:
63:    HandleError Err.Number, Err.Description, Erl & "|" & "frmReg.cmdReg_Click()"
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
1:   Set frmReg = Nothing
End Sub

Private Sub Labels_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub
