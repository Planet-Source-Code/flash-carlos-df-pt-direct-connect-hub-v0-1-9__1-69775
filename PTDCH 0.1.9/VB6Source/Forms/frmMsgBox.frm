VERSION 5.00
Begin VB.Form frmMsgBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PT DC Hub x.x.x"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2625
   Icon            =   "frmMsgBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   2625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblHolder 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "There is already one instance of this particular PTDCH exe running."
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Set caption to proper format
    Me.Caption = "PT Direct Connect Hub " & vbVersion
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
