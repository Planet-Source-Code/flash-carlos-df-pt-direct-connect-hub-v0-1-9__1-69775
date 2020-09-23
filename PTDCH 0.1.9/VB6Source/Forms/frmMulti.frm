VERSION 5.00
Begin VB.Form frmMulti 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Multi Use"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4590
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtStr 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3260
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtStrMultiLine 
      Height          =   1005
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
1:   txtStrMultiLine.Text = ""
2:   txtStr.Text = ""
3:   Unload Me
End Sub

Private Sub cmdOK_Click()
1:   Me.Hide
End Sub

Private Sub Form_Load()
1:   DoEvents
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub

Private Sub Form_Paint()
1: If g_objSettings.blSkin Then _
     PaintTileFormBackground Me, iResPic(g_objSettings.lngSkin)
End Sub
