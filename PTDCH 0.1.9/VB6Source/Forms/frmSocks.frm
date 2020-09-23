VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSocks 
   BorderStyle     =   0  'None
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   645
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin MSWinsockLib.Winsock wskScript 
      Index           =   0
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_objScript             As ScriptControl
Private m_blnConnect            As Boolean
Private m_blnClose              As Boolean
Private m_blnConnectionRequest  As Boolean
Private m_blnDataArrival        As Boolean
Private m_blnError              As Boolean

'NOTES :

'- frmSocks is not used by the interface; it is used by scripts so that each
'  one can have it's own winsock collection / events

Public Sub SetBools(ByRef blnConnect As Boolean, ByRef blnClose As Boolean, ByRef blnConnectionRequest As Boolean, ByRef blnDataArrival As Boolean, ByRef blnError As Boolean)
1:    m_blnConnect = blnConnect
2:    m_blnClose = blnClose
3:    m_blnConnectionRequest = blnConnectionRequest
4:    m_blnDataArrival = blnDataArrival
5:    m_blnError = blnError
End Sub

Public Property Set Script(ByRef objData As ScriptControl)
1:    Set m_objScript = objData
End Property

Private Sub wskScript_Close(Index As Integer)
1:    On Error Resume Next
    
3:    If m_blnClose Then m_objScript.Run "wskScript_Close", Index
End Sub

Private Sub wskScript_Connect(Index As Integer)
1:    On Error Resume Next
    
3:    If m_blnConnect Then m_objScript.Run "wskScript_Connect", Index
End Sub

Private Sub wskScript_ConnectionRequest(Index As Integer, ByVal requestID As Long)
1:    On Error Resume Next
    
3:    If m_blnConnectionRequest Then m_objScript.Run "wskScript_ConnectionRequest", Index, requestID
End Sub

Private Sub wskScript_DataArrival(Index As Integer, ByVal bytesTotal As Long)
1:    On Error Resume Next
    
3:    If m_blnDataArrival Then m_objScript.Run "wskScript_DataArrival", Index, bytesTotal
End Sub

Private Sub wskScript_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
1:    On Error Resume Next
    
3:    If m_blnError Then m_objScript.Run "wskScript_Error", Index, Number, Description
End Sub
