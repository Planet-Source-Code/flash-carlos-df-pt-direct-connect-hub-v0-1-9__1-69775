VERSION 5.00
Begin VB.Form frmCommand 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Command"
   ClientHeight    =   3075
   ClientLeft      =   315
   ClientTop       =   525
   ClientWidth     =   4425
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCommand.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdButton 
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CheckBox chkEnabled 
      Appearance      =   0  'Flat
      Caption         =   "Enabled"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   195
   End
   Begin VB.TextBox txtTrigger 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox cmbClass 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmCommand.frx":169A
      Left            =   1440
      List            =   "frmCommand.frx":16B4
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtDescription 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label lbbCheck 
      BackStyle       =   0  'Transparent
      Caption         =   "Enabled"
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblHolder 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Trigger"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblHolder 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum class"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  On Error GoTo Err

3:    Me.Caption = g_colMessages.Item("msgCommand")
    
5:    lblHolder(0).Caption = g_colMessages.Item("msgCmdTrigger")
6:    lblHolder(1).Caption = g_colMessages.Item("msgCmdMinClas")
7:    lbbCheck(0).Caption = g_colMessages.Item("msgCmdEnabled")
    
9:    cmdButton(0).Caption = g_colMessages.Item("msgOK")
10:   cmdButton(1).Caption = g_colMessages.Item("msgCancel")
    
12:    Exit Sub
    
14:
Err:
16:    HandleError Err.Number, Err.Description, Erl & "|" & "frmCommand.Form_Load()"
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
1:   Set frmCommand = Nothing
End Sub

Private Sub lblHolder_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub

Private Sub cmdButton_Click(Index As Integer)
1:    Dim strReason As String
    
3:    On Error GoTo Err
    
    'Save the command changes if needed
6:    If Index = 0 Then g_colCommands.Edit Tag, txtTrigger.Text, txtDescription.Text, cmbClass.ItemData(cmbClass.ListIndex), chkEnabled.Value
    
8:    Unload Me
    
10:    Exit Sub
    
12:
Err:
13:    HandleError Err.Number, Err.Description, Erl & "|" & "frmCommand.lblMenu_Click(" & Index & ")"
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
1:   If KeyAscii = 13 Then _
          Call cmdButton_Click(0)
End Sub

Private Sub txtTrigger_KeyPress(KeyAscii As Integer)
1:   If KeyAscii = 13 Then _
          Call cmdButton_Click(0)
End Sub
