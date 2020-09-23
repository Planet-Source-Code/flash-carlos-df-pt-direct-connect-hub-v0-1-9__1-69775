VERSION 5.00
Begin VB.Form frmNewScript 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "New Script"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3840
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox cmbScript 
      Height          =   315
      ItemData        =   "frmNewScript.frx":0000
      Left            =   720
      List            =   "frmNewScript.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "New Script"
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select script type:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the name of the script:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmNewScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
1:     On Error GoTo Err

3:     Dim i             As Integer
4:     Dim blnIsVB       As Boolean
5:     Dim strOne        As String
6:     Const strC_VBScript = "Option Explicit" & vbNewLine & vbNewLine & _
                              "Sub Main()" & vbNewLine & vbNewLine & _
                              vbTab & "MsgBox ""Hello World!"", , ""VBScript""" & vbNewLine & vbNewLine & _
                              "End Sub" & vbNewLine

11:    Const strC_JScript = vbNewLine & vbNewLine & _
                            "function Main()" & vbNewLine & "{" & vbNewLine & _
                            vbTab & "alert(""Hello World"");" & vbNewLine & vbNewLine & _
                            "}" & vbNewLine

16:    Select Case cmbScript.Text
           Case "VBScript (*.script)": strOne = txtName.Text & ".script": blnIsVB = True
           Case "VBScript (*.vbs)": strOne = txtName.Text & ".vbs": blnIsVB = True
           Case "JScript (*.js)": strOne = txtName.Text & ".js": blnIsVB = False
           'Case "PerlScript - *.pl": strOne = txtName.Text & ".pl": blnIsVB = False
       End Select

23:    If txtName.Text = "" Then txtName.Text = g_colMessages.Item("msgNewScript")

       'Make sure there isn't another script with the same name
26:    For i = 1 To frmHub.lvwScripts.ListItems.count
27:         If frmHub.lvwScripts.ListItems(i).Text = strOne Then
28:             MsgBoxCenter Me, strOne & g_colMessages.Item("msgScriptAlready"), vbInformation, g_colMessages.Item("msgNewScript")
29:             Exit Sub
30:         End If
31:    Next

       'Create new file
34:    If blnIsVB Then _
            g_objFileAccess.WriteFile G_APPPATH & "\Scripts\" & strOne, strC_VBScript _
       Else g_objFileAccess.WriteFile G_APPPATH & "\Scripts\" & strOne, strC_JScript

       'Reolad all scripts..
39:    frmScript.SLoadScript strOne

41:    Unload Me

43:  Exit Sub
44:
Err:
46:  HandleError Err.Number, Err.Description, Erl & "|" & "frmNewScript.cmdAdd()"
End Sub

Private Sub cmdCancel_Click()
1:  Unload Me
End Sub

Private Sub Form_Load()

2:  On Error GoTo Err
 
4:   Me.Caption = g_colMessages.Item("msgNewScript")
 
6:   cmbScript.AddItem "VBScript (*.script)"
7:   cmbScript.AddItem "VBScript (*.vbs)"
8:   cmbScript.AddItem "JScript (*.js)"
'9:   cmbScript.AddItem "PerlScript (*.pl)"
10:   cmbScript.Text = "VBScript (*.vbs)"

12:   txtName.Text = g_colMessages.Item("msgNewScript")

14:  Label(0).Caption = g_colMessages.Item("msgNewScriptName")
15:  Label(1).Caption = g_colMessages.Item("msgNewScriptType")
     
17:  cmdCancel.Caption = g_colMessages.Item("msgCancel")
18:  cmdAdd.Caption = g_colMessages.Item("msgAdd")

20:  Exit Sub
21:
Err:
23:  HandleError Err.Number, Err.Description, Erl & "|" & "frmNewScript.cmdCancel_Click()"
End Sub

Private Sub Form_Paint()
1: If g_objSettings.blSkin Then _
     PaintTileFormBackground Me, iResPic(g_objSettings.lngSkin)
End Sub

Private Sub Form_Unload(Cancel As Integer)
1:  Set frmNewScript = Nothing
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub

Private Sub Label_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub

Private Sub txtName_Change()
    'replace invalide chars
2:    txtName.Text = Replace(txtName.Text, "*", "_")
3:    txtName.Text = Replace(txtName.Text, ":", "_")
4:    txtName.Text = Replace(txtName.Text, "<", "_")
5:    txtName.Text = Replace(txtName.Text, ">", "_")
6:    txtName.Text = Replace(txtName.Text, "\", "_")
7:    txtName.Text = Replace(txtName.Text, "?", "_")
8:    txtName.Text = Replace(txtName.Text, "|", "_")
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
1:   If KeyAscii = 13 Then _
          Call cmdAdd_Click
End Sub
