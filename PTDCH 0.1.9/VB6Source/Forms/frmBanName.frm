VERSION 5.00
Begin VB.Form frmBanName 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ban Name"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAddRem 
      Caption         =   "Add"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtReason 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      ToolTipText     =   "Enter the reason why you're banning the name (optional)"
      Top             =   1200
      Width           =   3615
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
      Width           =   3615
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the reason why you're banning the name (optional)"
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the name to ban."
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmBanName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub InicializeBan()
 
  On Error GoTo Err
  
4:   If Me.Tag = "Add" Then
5:      Caption = g_colMessages.Item("msgBanName")
6:      Labels(0).Caption = g_colMessages.Item("msgEnterBanName")
7:      Labels(1).Caption = g_colMessages.Item("msgEnterBanReason")
8:      cmdAddRem.Caption = g_colMessages.Item("msgAdd")
9:   Else
10:     Caption = g_colMessages.Item("msgRenameBan")
11:     Labels(0).Caption = g_colMessages.Item("msgEnterReplace")
12:     Labels(1).Caption = g_colMessages.Item("msgEnterBanName")
13:     txtReason.Text = frmHub.lblHolder(50).Caption
14:     txtReason.Enabled = False
15:     cmdAddRem.Caption = g_colMessages.Item("msgRemame")
16:     txtReason.Locked = True
17:  End If
18:  cmdClose.Caption = g_colMessages.Item("msgClose")
  
20:  Exit Sub

22:
Err:
24:  HandleError Err.Number, Err.Description, Erl & "|" & "frmBanName.Form_Load()"
End Sub
Private Sub cmdAddRem_Click()

  On Error GoTo Err

4:    If Me.Tag = "Add" Then
         Select Case g_objRegistered.Add(txtName.Text, txtReason.Text, Locked)
           Case 1
7:           MsgBoxCenter Me, txtName.Text & g_colMessages.Item("msgAlreadyRegged"), vbInformation
           Case 2 'Name longer than 40 chars
9:           MsgBox g_colMessages.Item("msgInvalidBanName"), , g_colMessages.Item("msgBanName"), vbInformation
         'Case 3 'This error should not be returned b/c reasons for banning names don't have limits
         End Select
12:   ElseIf Me.Tag = "Rename" Then
13:       g_objRegistered.Rename txtName.Tag, txtName.Text
14:   End If
    
16:   Call cmdClose_Click
    
18:  Exit Sub

20:
Err:
21:  HandleError Err.Number, Err.Description, Erl & "|" & "frmBanName.cmdAddRem_Click()"
End Sub

Private Sub cmdClose_Click()
1:   Unload Me
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
1:   Set frmBanName = Nothing
End Sub

Private Sub Labels_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub

Private Sub txtReason_KeyPress(KeyAscii As Integer)
1:   If KeyAscii = 13 Then _
          Call cmdAddRem_Click
End Sub
