VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmHelp 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "ScriptHelp.txt"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSciMain 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   0
      ScaleHeight     =   4335
      ScaleWidth      =   6615
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
   Begin ComctlLib.StatusBar stb 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4500
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   ""
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   ""
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   ""
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Cool FX Magnetic Windows
Private Magnetic As New clsMagneticWnd

Private WithEvents m_sciMain As clsYScintilla
Attribute m_sciMain.VB_VarHelpID = -1
Dim m_bIsIniSCI As Boolean

Private Sub Form_Load()
1:  On Error GoTo Err

3:  If g_objFileAccess.FileExists(G_APPPATH & "\Settings\ScriptHelp.vbs") Then
4:      Set m_sciMain = New clsYScintilla
  
6:      If g_objSettings.MagneticWin Then _
            Call Magnetic.AddWindow(frmHelp.hWnd)  'Cool FX Windows
        
9:      m_sciMain.CreateScintilla picSciMain
10:     m_sciMain.SetFixedFont "Courier New", 10

        ' Give the scrollbar a nice long width to handle a long line which may
        ' occur.
14:     m_sciMain.ScrollWidth = 10000

        'This is absolutly an imperative line
17:     m_sciMain.Attach picSciMain
  
19:     m_sciMain.LoadFile G_APPPATH & "\Settings\ScriptHelp.vbs"
  
21:     m_sciMain.Folding = False
22:     m_sciMain.ShowCallTips = True
23:     m_sciMain.LineNumbers = True
24:     m_sciMain.AutoIndent = True
25:     m_sciMain.ReadOnly = True

27:     m_sciMain.SetMarginWidth MarginLineNumbers, 50
  
29:     Call Highlighter.SetHighlighterBasedOnExt(m_sciMain, ".vbs")
        
31:     m_bIsIniSCI = True
32:  Else
33:     m_bIsIniSCI = False
34:     MsgBox "File not found: " & vbNewLine & _
              G_APPPATH & "\Settings\ScriptHelp.vbs", vbCritical
36:  End If
     
37:  Exit Sub
    
39:
Err:
41:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHelp.Form_Load()"
End Sub

Private Sub Form_Resize()
1:  On Error Resume Next
2:  If m_bIsIniSCI Then
3:      picSciMain.Width = Me.ScaleWidth
4:      picSciMain.Height = Me.ScaleHeight - (stb.Height)
5:      m_sciMain.SizeScintilla 0, 0, Me.ScaleWidth / Screen.TwipsPerPixelX, (Me.ScaleHeight / Screen.TwipsPerPixelY) - (stb.Height / Screen.TwipsPerPixelY)
6:  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
1:  Magnetic.RemoveWindow frmHelp.hWnd
2:  If m_bIsIniSCI Then
3:       'This is absolutly an imperative line
4:      m_sciMain.Detach picSciMain
5:      Set m_sciMain = Nothing
6:  End If
End Sub

Private Sub m_sciMain_UpdateUI()
1:  If m_bIsIniSCI Then
2:      stb.Panels(1).Text = "CurrentLine: " & m_sciMain.GetCurLine
3:      stb.Panels(2).Text = "Column: " & m_sciMain.GetColumn
4:      stb.Panels(3).Text = "Lines: " & m_sciMain.GetLineCount
5:  End If
End Sub
