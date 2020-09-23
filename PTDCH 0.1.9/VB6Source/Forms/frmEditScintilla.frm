VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmEditScintilla 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBordTab 
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   4440
      ScaleHeight     =   260
      ScaleMode       =   0  'User
      ScaleWidth      =   2460
      TabIndex        =   31
      Top             =   60
      Width           =   2460
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   5160
      TabIndex        =   30
      Top             =   2640
      Width           =   1455
   End
   Begin VB.ComboBox cmbReplace 
      Height          =   315
      Left            =   1560
      TabIndex        =   6
      Top             =   1080
      Width           =   3255
   End
   Begin VB.ComboBox cmbFind 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
   Begin VB.PictureBox picLbls 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   0
      Left            =   240
      ScaleHeight     =   615
      ScaleWidth      =   1215
      TabIndex        =   2
      Top             =   600
      Width           =   1215
      Begin VB.Label lblHolder 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Find what:"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   110
         Width           =   1215
      End
   End
   Begin VB.PictureBox picCmds 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   0
      Left            =   4920
      ScaleHeight     =   1455
      ScaleWidth      =   1575
      TabIndex        =   8
      Top             =   600
      Width           =   1575
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find Next"
         Enabled         =   0   'False
         Height          =   375
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmdFindPrev 
         Caption         =   "Find Previous"
         Enabled         =   0   'False
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.PictureBox picLbls 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   1
      Left            =   240
      ScaleHeight     =   855
      ScaleWidth      =   1215
      TabIndex        =   4
      Top             =   600
      Width           =   1215
      Begin VB.Label lblHolder 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Replace with:"
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   7
         Top             =   580
         Width           =   1215
      End
      Begin VB.Label lblHolder 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Find what:"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   5
         Top             =   110
         Width           =   1215
      End
   End
   Begin VB.PictureBox pictOptions 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   240
      ScaleHeight     =   975
      ScaleWidth      =   4215
      TabIndex        =   24
      Top             =   1560
      Width           =   4215
      Begin VB.CheckBox chkRegExp 
         BackColor       =   &H80000005&
         Caption         =   "Regular expression"
         Height          =   195
         Left            =   0
         TabIndex        =   28
         Top             =   720
         Width           =   2595
      End
      Begin VB.CheckBox chkCase 
         BackColor       =   &H80000005&
         Caption         =   "Match case"
         Height          =   195
         Left            =   0
         TabIndex        =   27
         Top             =   480
         Width           =   2595
      End
      Begin VB.CheckBox chkWhole 
         BackColor       =   &H80000005&
         Caption         =   "Match whole word only"
         Height          =   195
         Left            =   0
         TabIndex        =   26
         Top             =   240
         Width           =   2595
      End
      Begin VB.CheckBox chkWrap 
         BackColor       =   &H80000005&
         Caption         =   "Wrap around"
         Height          =   195
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Value           =   1  'Checked
         Width           =   2595
      End
   End
   Begin VB.PictureBox picGoTo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   360
      ScaleHeight     =   1455
      ScaleWidth      =   6015
      TabIndex        =   15
      Top             =   720
      Width           =   6015
      Begin VB.CommandButton cmdGo 
         Caption         =   "Go"
         Height          =   375
         Left            =   4200
         TabIndex        =   18
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtColumn 
         Height          =   285
         Left            =   4200
         TabIndex        =   17
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox txtLine 
         Height          =   285
         Left            =   1800
         TabIndex        =   16
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lblGoTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Column: "
         Height          =   195
         Index           =   3
         Left            =   3480
         TabIndex        =   23
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblGoTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Line:"
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   22
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label lblGoTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Line: "
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   21
         Top             =   600
         Width           =   945
      End
      Begin VB.Label lblGoTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Column:"
         Height          =   195
         Index           =   1
         Left            =   3480
         TabIndex        =   20
         Top             =   165
         Width           =   570
      End
      Begin VB.Label lblGoTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destination Line:"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   19
         Top             =   165
         Width           =   1185
      End
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2130
      Left            =   120
      ScaleHeight     =   2722.483
      ScaleMode       =   0  'User
      ScaleWidth      =   6495
      TabIndex        =   32
      Top             =   390
      Width           =   6495
   End
   Begin ComctlLib.TabStrip tbsMenu 
      Height          =   2535
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4471
      TabWidthStyle   =   2
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Find"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Replace"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "GoTo"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picCmds 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   1
      Left            =   4920
      ScaleHeight     =   1455
      ScaleWidth      =   1575
      TabIndex        =   11
      Top             =   600
      Width           =   1575
      Begin VB.CommandButton cmdReplaceAll 
         Caption         =   "Replace All"
         Enabled         =   0   'False
         Height          =   375
         Left            =   0
         TabIndex        =   14
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmdReplace 
         Caption         =   "Replace"
         Enabled         =   0   'False
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdFindNextR 
         Caption         =   "Find Next"
         Enabled         =   0   'False
         Height          =   375
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Label lblReplace 
      BackStyle       =   0  'Transparent
      Caption         =   "Replaced xxx times"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   2700
      Width           =   3375
   End
End
Attribute VB_Name = "frmEditScintilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cScintilla As clsYScintilla

'Cool FX Magnetic Windows
Private Magnetic As New clsMagneticWnd

Public Sub SetObjects(intTblSelected As Integer, _
                Optional blnUpDateOnlyGoToLine As Boolean = False)
    On Error GoTo Err
    
3: If blnUpDateOnlyGoToLine Then GoTo Hell
                
    Select Case intTblSelected
        Case 1
7:            pictOptions.Visible = True
8:            cmbFind.Visible = True
9:            picCmds(0).Visible = True
10:            picLbls(0).Visible = True
            '
12:            cmbReplace.Visible = False
13:            picCmds(1).Visible = False
14:            picLbls(1).Visible = False
15:            picGoTo.Visible = False
16:            lblReplace.Visible = False
        Case 2
18:            pictOptions.Visible = True
19:            cmbFind.Visible = True
20:            cmbReplace.Visible = True
21:            picCmds(1).Visible = True
22:            picLbls(1).Visible = True
23:            lblReplace.Visible = True
            '
25:            picCmds(0).Visible = False
26:            picLbls(0).Visible = False
27:            picGoTo.Visible = False
        Case 3
29:            picGoTo.Visible = True
            '
31:            pictOptions.Visible = False
32:            picCmds(0).Visible = False
33:            picCmds(1).Visible = False
34:            picLbls(0).Visible = False
35:            picLbls(1).Visible = False
36:            cmbFind.Visible = False
37:            cmbReplace.Visible = False
38:            lblReplace.Visible = False
Hell:
40:            lblGoTo(2).Caption = lblGoTo(2).Tag & cScintilla.GetCurLine
41:            lblGoTo(4).Caption = lblGoTo(4).Tag & cScintilla.GetLastLine
42:            lblGoTo(3).Caption = lblGoTo(3).Tag & cScintilla.GetColumn

43:            If txtLine.Text = "" Then txtLine.Text = 1
44:            If txtColumn.Text = "" Then txtColumn.Text = 1

    End Select
    
48:    lblReplace.Caption = ""
      
50:    Exit Sub
    
52:
Err:
54:    HandleError Err.Number, Err.Description, Erl & "|" & "frmEditScintilla.SetObjects()"
End Sub

Private Sub cmbFind_Change()
   Dim i As Integer
   Dim X As Boolean
   
   On Error GoTo Err
  
6:    If Not cmbFind.Text <> "" Then
7:        cmdFind.Enabled = False
8:        cmdFindPrev.Enabled = False
       '
10:       cmdFindNextR.Enabled = False
11:       cmdReplace.Enabled = False
12:       cmdReplaceAll.Enabled = False
13:    Else
14:       cmdFind.Enabled = True
15:       cmdFindPrev.Enabled = True
       '
17:       cmdFindNextR.Enabled = True
18:       cmdReplace.Enabled = True
19:       cmdReplaceAll.Enabled = True
20:    End If
   
22:   X = True

24:   If cmbFind.Text = "" Or Replace(cmbFind.Text, " ", "") = "" Then Exit Sub

26:   For i = 0 To cmbFind.ListCount - 1
27:       If cmbFind.List(i) = cmbFind.Text Then X = False
28:   Next i

30:   If X Then cmbFind.AddItem cmbFind.Text    'Set text in cmbFind

32:    Exit Sub
    
34:
Err:
36:    HandleError Err.Number, Err.Description, Erl & "|" & "frmEditScintilla.cmbFind_Change()"
End Sub

Private Sub cmbFind_Click()
1:   Call cmbFind_Change
End Sub

Private Sub cmbReplace_Change()
   Dim i As Integer
   Dim X As Boolean
   
   On Error GoTo Err
       
6:   X = True

8:   If cmbReplace.Text = "" Or Replace(cmbReplace.Text, " ", "") = "" Then Exit Sub

10:   For i = 0 To cmbReplace.ListCount - 1
11:       If cmbReplace.List(i) = cmbReplace.Text Then X = False
12:   Next i

14:   If X Then cmbReplace.AddItem cmbReplace.Text    'Set text in cmbReplace

16:    Exit Sub
    
18:
Err:
20:    HandleError Err.Number, Err.Description, Erl & "|" & "frmEditScintilla.cmbReplace_Change()"
End Sub

Private Sub cmdFindNextR_Click()
1:  cScintilla.FindText cmbFind.Text, False, False, chkWrap.Value, chkCase.Value, False, chkWhole.Value, chkRegExp.Value
2:  lblReplace.Caption = ""
End Sub

Private Sub cmdGo_Click()
    Dim iLine As Long, iCol As Long
    
    On Error GoTo Err
   
5:    If txtLine.Text = "" Or _
        Not IsNumeric(txtLine.Text) Then txtLine.Text = 1
        
8:    If txtColumn.Text = "" Or _
        Not IsNumeric(txtColumn.Text) Then txtColumn.Text = 1
    
11:    iLine = txtLine.Text
12:    iCol = txtColumn.Text
13:    cScintilla.GotoLineColumn iLine, iCol
    
15:    SetObjects 3, True
    
17:    Exit Sub
    
19:
Err:
21:    HandleError Err.Number, Err.Description, Erl & "|" & "frmEditScintilla.cmdGo_Click()"
End Sub

Private Sub cmdReplace_Click()
   Dim iGetPos As Long
  
   On Error GoTo Err
    
5:  cScintilla.FindText cmbFind.Text, False, False, chkWrap.Value, chkCase.Value, False, chkWhole.Value, chkRegExp.Value

8:  If cmbReplace.Text <> "" Then
9:    cScintilla.ReplaceSel cmbReplace.Text
10:    cScintilla.SetSel cScintilla.GetCurPos - 1, cScintilla.GetCurPos - 1 + Len(cmbReplace.Text)
11:  Else
12:    iGetPos = cScintilla.GetCurPos
13:    cScintilla.ReplaceSel ""
14:    cScintilla.SetSel iGetPos - LenB(cmbFind.Text), iGetPos - LenB(cmbFind.Text)
15:    cScintilla.FindText cmbFind.Text, False, False, chkWrap.Value, chkCase.Value, False, chkWhole.Value, chkRegExp.Value
16:  End If
  
18:  lblReplace.Caption = ""
    
20:    Exit Sub
    
22:
Err:
24:    HandleError Err.Number, Err.Description, Erl & "|" & "frmEditScintilla.cmdReplace_Click()"
End Sub

Private Sub cmdReplaceAll_Click()
    Dim iRep As Long
    
    On Error GoTo Err

5:    iRep = cScintilla.ReplaceAll(cmbFind.Text, cmbReplace.Text, chkCase.Value, chkRegExp.Value, chkWhole.Value, False)
    
7:    If iRep > 0 Then
8:       lblReplace.Caption = Replace(lblReplace.Tag, "%[times]", iRep) '"Replaced " & iRep & " times"
9:    Else
10:      MsgBox "No instances of " & """" & cmbFind.Text & """" & " were found in document"
11:   End If

13:    Exit Sub
    
15:
Err:
17:    HandleError Err.Number, Err.Description, Erl & "|" & "frmEditScintilla.cmdReplaceAll_Click()"
End Sub

Private Sub Form_Activate()
1:    SetObjects 3, True
End Sub

Private Sub Form_Load()
    On Error GoTo Err
2:  If g_objSettings.MagneticWin Then _
        Call Magnetic.AddWindow(frmEditScintilla.hWnd)  'Cool FX Windows
       'Set languages..
5:     tbsMenu.Tabs(1).Caption = g_colMessages.Item("msgSCIFind")
6:     tbsMenu.Tabs(2).Caption = g_colMessages.Item("msgSCIReplace")
7:     tbsMenu.Tabs(3).Caption = g_colMessages.Item("msgSCIGoTo")
       '
9:     cmdFindNextR.Caption = g_colMessages.Item("msgSCIFindNext")
10:    cmdReplace.Caption = g_colMessages.Item("msgSCIReplace")
11:    cmdReplaceAll.Caption = g_colMessages.Item("msgSCIReplaceAll")
       '
13:    cmdFind.Caption = g_colMessages.Item("msgSCIFindNext")
14:    cmdFindPrev.Caption = g_colMessages.Item("msgSCIFindPrev")
       '
16:    cmdGo.Caption = g_colMessages.Item("msgSCIGo")
       '
18:    cmdClose.Caption = g_colMessages.Item("msgClose")
       '
19:    chkWrap.Caption = g_colMessages.Item("msgSCIWrap")
20:    chkWhole.Caption = g_colMessages.Item("msgSCIWhole")
21:    chkCase.Caption = g_colMessages.Item("msgSCICase")
22:    chkRegExp.Caption = g_colMessages.Item("msgSCIRegExp")
       '
24:    lblHolder(0).Caption = g_colMessages.Item("msgSCIFindWhat")
25:    lblHolder(1).Caption = g_colMessages.Item("msgSCIFindWhat")
26:    lblHolder(2).Caption = g_colMessages.Item("msgSCIReplWith")
27:    lblGoTo(0).Caption = g_colMessages.Item("msgSCIDestLine")
28:    lblGoTo(1).Caption = g_colMessages.Item("msgSCIColumn")
       '
30:    lblGoTo(2).Caption = g_colMessages.Item("msgSCICurrLine")
31:    lblGoTo(2).Tag = g_colMessages.Item("msgSCICurrLine")
       '
33:    lblGoTo(3).Caption = g_colMessages.Item("msgSCIColumn")
34:    lblGoTo(3).Tag = g_colMessages.Item("msgSCIColumn")
       '
36:    lblGoTo(4).Caption = g_colMessages.Item("msgSCILastLine")
37:    lblGoTo(4).Tag = g_colMessages.Item("msgSCILastLine")
       '
39:    lblReplace.Tag = g_colMessages.Item("msgSCIReplTimes")

41:   Exit Sub

Err:
44:   HandleError Err.Number, Err.Description, Erl & "|frmEditScintilla.Form_Load()"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub

Private Sub Form_Paint()
1:   If g_objSettings.blSkin Then _
       PaintTileFormBackground Me, iResPic(g_objSettings.lngSkin)
End Sub

Private Sub lblHolder_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub

Private Sub lblReplace_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub

Private Sub picBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub

Private Sub picBordTab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub

Private Sub picBordTab_Paint(Index As Integer)
1:   If g_objSettings.blSkin Then _
         PaintTilePicBackground Me.picBordTab(Index), iResPic(g_objSettings.lngSkin)
End Sub

Private Sub picCmds_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub

Private Sub picGoTo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub

Private Sub picLbls_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub

Private Sub pictOptions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub

Private Sub tbsMenu_Click()
1:    Call SetObjects(Val(tbsMenu.SelectedItem.Index))
2:    Me.Caption = tbsMenu.SelectedItem.Caption
End Sub

Private Sub cmdFind_Click()
1:    cScintilla.FindText cmbFind.Text, False, False, chkWrap.Value, chkCase.Value, False, chkWhole.Value, chkRegExp.Value
End Sub

Private Sub cmdFindPrev_Click()
1:    cScintilla.FindPrev
End Sub

Private Sub cmdClose_Click()
1:    Me.Hide
End Sub

Private Sub txtColumn_KeyPress(KeyAscii As Integer)
1:   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtLine_KeyPress(KeyAscii As Integer)
1:   If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0
End Sub
