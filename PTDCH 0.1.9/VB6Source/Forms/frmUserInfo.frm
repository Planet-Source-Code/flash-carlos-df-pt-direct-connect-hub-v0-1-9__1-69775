VERSION 5.00
Begin VB.Form frmUserInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "User Info"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdClipBoard 
      Caption         =   "Clipboard"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtInfo 
      Height          =   2295
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   0
      Top             =   60
      Width           =   4575
   End
End
Attribute VB_Name = "frmUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClipBoard_Click()
1:   Clipboard.Clear
2:   Clipboard.SetText txtInfo.Text
End Sub

Private Sub cmdOK_Click()
1:   Unload Me
End Sub

Private Sub Form_Load()
1:   cmdOK.Caption = g_colMessages.Item("msgOK")
2:   cmdClipBoard.Caption = g_colMessages.Item("msgClipboard")
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
1:   Set frmUserInfo = Nothing
End Sub
