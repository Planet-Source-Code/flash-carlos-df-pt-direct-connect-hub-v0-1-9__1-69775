VERSION 5.00
Begin VB.Form frmBanPerm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ban permanent IP"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtIP 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the IP to permanet ban."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmBanPerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
1:   Unload Me
End Sub

Private Sub cmdCancel_GotFocus()
1:   txtIP.SelStart = 0
2:   txtIP.SelLength = Len(txtIP)
End Sub

Private Sub cmdOK_Click()
   On Error GoTo Err

3:   txtIP.Text = Replace(txtIP.Text, " ", "")

5:   If ValidIP(txtIP) Then
6:      g_objIPBans.Add txtIP.Text
7:      Unload Me
8:   Else
9:      MsgBoxCenter Me, """" & txtIP.Text & """" & g_colMessages.Item("msgIPNotValide"), vbInformation
10:  End If
   
12:  Exit Sub

14:
Err:
16:  HandleError Err.Number, Err.Description, Erl & "|" & "frmBanPerm.cmdOK_Click()"
End Sub

Private Sub Form_Load()
   On Error GoTo Err
2:    Me.Caption = g_colMessages.Item("msgBanPermIP")
3:    Label.Caption = g_colMessages.Item("msgEnterPermIP")
4:    cmdCancel.Caption = g_colMessages.Item("msgCancel")
5:    cmdOK.Caption = g_colMessages.Item("msgOK")
6:  Exit Sub

8:
Err:
10:  HandleError Err.Number, Err.Description, Erl & "|" & "frmBanPerm.Form_Load()"
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
1:   Set frmBanPerm = Nothing
End Sub

Private Sub Label_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub

Private Sub txtIP_KeyPress(KeyAscii As Integer)
1:   If KeyAscii = 13 Then _
          Call cmdOK_Click
End Sub
