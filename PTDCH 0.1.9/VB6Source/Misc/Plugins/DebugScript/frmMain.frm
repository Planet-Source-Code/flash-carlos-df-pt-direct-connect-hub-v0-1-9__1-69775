VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Active Scripting Debugger Window"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtbDebug 
      DataSource      =   "(None)"
      Height          =   2775
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4895
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      RightMargin     =   1e7
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMain.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuEdits 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuEdit 
         Caption         =   "Copy"
         Index           =   0
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Clear"
         Index           =   2
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Top Most"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------'
'  Copyright (C) fLaSh - Carlos Ferreira       '
'  E-Mail: Carlosferreiracarlos@hotmail.com    '
'----------------------------------------------'

'Option Explicit

Private Declare Function SetWindowPos Lib "user32" _
      (ByVal hWnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long
      
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
   (ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Private Function SetTopMostWindow(hWnd As Long, TopMost As Boolean) As Long
    On Error GoTo Err
    
    If TopMost = True Then
       SetTopMostWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, _
          0, FLAGS)
    Else
       SetTopMostWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, _
          0, 0, FLAGS)
       SetTopMostWindow = False
    End If
    
    Exit Function
Err:
    MsgBox Err.Description, vbCritical, mName & " - SetTopMostWindow()"
    Err.Clear
End Function

Private Sub Form_Load()
    'Set top most by defaut
    mnuEdit(4).Checked = Not mnuEdit(4).Checked
    SetTopMostWindow Me.hWnd, CBool(mnuEdit(4).Checked)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    rtbDebug.Move 30, 30, ScaleWidth - 60, ScaleHeight - 60
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Err
    
    Set frmMain = Nothing
    
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical, mName & " - Form_Unload()"
    Err.Clear
End Sub

Private Sub mnuEdit_Click(Index As Integer)
    On Error GoTo Err
    
    Select Case Index
        Case 0: SendMessageLong rtbDebug.hWnd, &H301, 0, 0
        Case 2: rtbDebug.Text = ""
        Case 4
             mnuEdit(4).Checked = Not mnuEdit(4).Checked
             SetTopMostWindow Me.hWnd, CBool(mnuEdit(4).Checked)
    End Select
    
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical, mName & " - mnuEdit_Click()"
    Err.Clear
End Sub

Private Sub rtbDebug_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err
    
     If Button = 2 Then
        If rtbDebug.SelText = "" Then _
             mnuEdit(0).Enabled = False _
        Else mnuEdit(0).Enabled = True
        If rtbDebug.Text = "" Then _
             mnuEdit(2).Enabled = False _
        Else mnuEdit(2).Enabled = True
        PopupMenu mnuEdits
    End If
    
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical, mName & " - rtbDebug_MouseDown()"
    Err.Clear
End Sub
