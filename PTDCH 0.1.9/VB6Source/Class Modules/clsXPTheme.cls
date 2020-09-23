VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXPTheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'----------------------------------------------'
'  Copyright (C) fLaSh - Carlos Ferreira       '-
'  E-Mail: Carlosferreiracarlos@hotmail.com    ' '
'  2007/02                                     ' '
'----------------------------------------------' '
'------------------------------------------------'---------------------'
'This module, when emplemented as described below, will cause the      '-
'compiled EXE to use Xp-Style buttons, fields, etc. if the system is   ' '
'and XP-type system.                                                   ' '
'----------------------------------------------------------------------'-'----------------'
'  Note: -Buttons with their Style set to Graphical will not be changed. This is because  '-
'        internally, they are no longer button classes, but picture boxes that VB treats  ' '
'        like buttons. Cool cheat, if you ask me.                                         ' '
'        -Warning "OptionButons" or "Butons" in the Frames, it causes problems!           ' '
'-----------------------------------------------------------------------------------------' '
  '-----------------------------------------------------------------------------------------'
'API stuff
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

' Incorporate OS stuff so that we will not need this module
Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion      As Long
  dwMinorVersion      As Long
  dwBuildNumber       As Long
  dwPlatformId        As Long
  szCSDVersion        As String * 128 'Maintenance string for PSS usage
End Type

Private Const VER_PLATFORM_WIN32_NT = 2
Private Const ICC_USEREX_CLASSES = &H200

Public Sub InitializeXP()

2:  Dim OSV As OSVERSIONINFO
3:  Dim iccex As tagInitCommonControlsEx
4:  Dim b As Boolean
  
6:   On Error Resume Next
  
8:   OSV.dwOSVersionInfoSize = Len(OSV)                   'set size of info block
  
10:  If CBool(GetVersionEx(OSV)) Then                     'get OS info
11:     If OSV.dwPlatformId = VER_PLATFORM_WIN32_NT Then  'WinNT type ?
12:        If OSV.dwMajorVersion >= 5 Then                'XP series (or later)?
              ' now tell the system to use XP stuff
14:           With iccex 'manifest string built-in from the resource "PTDCH.RES"
15:               .lngSize = LenB(iccex)                  'structure size
16:               .lngICC = ICC_USEREX_CLASSES            'Tell system "I WANT XP Buttons"
17:           End With
18:           On Error GoTo 0
19:           b = InitCommonControlsEx(iccex)             'process info
20:        End If
21:     End If
22:  End If

End Sub
