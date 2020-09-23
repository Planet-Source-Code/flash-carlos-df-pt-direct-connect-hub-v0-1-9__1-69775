VERSION 5.00
Begin VB.Form frmHelpFind 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Help - Find"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdButton 
      Caption         =   "Close"
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Find Previous"
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Find Next"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmHelpFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cScintilla As clsYScintilla

Private Sub cmdButton_Click(Index As Integer)
    Select Case Index
        Case 0: cScintilla.FindText txtFind.Text, False, False, False, False, False, False, False
        Case 1: cScintilla.FindPrev
        Case 2: Unload Me
    End Select
End Sub

Private Sub Form_Load()
1:  On Error GoTo Err
2:  If g_objSettings.MagneticWin Then _
        Call Magnetic.AddWindow(frmHelpFind.hWnd)  'Cool FX Windows

5:   cmdButton(0).Caption = g_colMessages.Item("msgSCIFindNext")
6:   cmdButton(1).Caption = g_colMessages.Item("msgSCIFindPrev")
7:   cmdButton(2).Caption = g_colMessages.Item("msgClose")

9:   Me.Caption = g_colMessages.Item("msgSCIFind")

11:  Exit Sub
Err:
13:  HandleError Err.Number, Err.Description, Erl & "|frmHelpFind.Form_Load()"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub

Private Sub Form_Paint()
1:   If g_objSettings.blSkin Then _
       PaintTileFormBackground Me, iResPic(g_objSettings.lngSkin)
End Sub

Private Sub Form_Unload(Cancel As Integer)
1:   Call Unload(frmHelpFind)
2:   Set cScintilla = Nothing
3:   Set frmHelpFind = Nothing
End Sub
