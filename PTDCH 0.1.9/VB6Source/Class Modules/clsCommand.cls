VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Class        As enuClass
Public ID           As Integer
Public Name         As String
Public Enabled      As Boolean
Public Description  As String

'------------------------------------------------
'Scheduler
'Public TrigStart As Date
'Public TrigEnd As Date

'Public TrigNext As Date
 
'Public TrigParams As String
 
'Public TrigMinutes As Integer
'Public TrigEnum As Integer

'TrigEna is to force disable even if TrigEnd is not reached
'Public TrigEna As Boolean
'Scheduler

'-------------------------------------------------
'Private m_intClass As enuClass
'Private m_intID As Integer
'Private m_strName As String
'Private m_blnEnabled As Boolean
'Private m_strDescription As String

'Public Property Let Class(ByRef intData As enuClass)
'    m_intClass = intData
'End Property

'Public Property Get Class() As enuClass
'    Class = m_intClass
'End Property

'Public Property Let ID(ByRef intData As Integer)
'    m_intID = intData
'End Property

'Public Property Get ID() As Integer
'    ID = m_intID
'End Property

'Public Property Let Name(ByRef strData As String)
'    m_strName = strData
'End Property

'Public Property Get Name() As String
'    Name = m_strName
'End Property

'Public Property Let Enabled(ByRef blnData As Boolean)
'    m_blnEnabled = blnData
'End Property

'Public Property Get Enabled() As Boolean
'    Enabled = m_blnEnabled
'End Property

'Public Property Let Description(ByRef strData As String)
'    m_strDescription = strData
'End Property

'Public Property Get Description() As String
'    Description = m_strDescription
'End Property

