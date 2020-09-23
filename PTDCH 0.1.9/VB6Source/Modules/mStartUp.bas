Attribute VB_Name = "mStartUp"
Option Explicit

'----------------------------------------------'
'  Copyright (C) fLaSh - Carlos Ferreira       '
'  E-Mail: Carlosferreiracarlos@hotmail.com    '
'----------------------------------------------'

'Start PTDCH at windows starting
Public Sub AddRegRun()
1:   On Error GoTo Err
2:   'Dim Reg As Object
3:   'Set Reg = CreateObject("Wscript.shell")
4:   'Reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\" & App.EXEName, App.Path & "\" & App.EXEName & ".exe"
5:   Exit Sub
Err:
7:   HandleError Err.Number, Err.Description, Erl & "|mStartUp.AddRegRun()"
End Sub

Public Sub RemRegRun()
1:   On Error GoTo Err
2:   'Dim Reg As Object
3:   'Set Reg = CreateObject("Wscript.Shell")
4:   'On Error Resume Next
5:   'Reg.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\" & App.EXEName
6:   Exit Sub
Err:
8:   HandleError Err.Number, Err.Description, Erl & "|mStartUp.RemRegRun()"
End Sub
