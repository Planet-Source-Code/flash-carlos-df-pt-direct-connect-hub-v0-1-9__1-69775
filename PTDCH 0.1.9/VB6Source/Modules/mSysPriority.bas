Attribute VB_Name = "mSysPriority"
Option Explicit
'---------------------------------------------------------------------------------------
' Module: Set Process(Application) Priority
' Date: 2007
' Author: fLaSh - Carlos D.F.
' Purpose: Used to Set and Retrieve the Priority of Your Processes
' Returns: If the function succeeds, the return value is nonzero.
'              If the function fails, the return value is zero.
' Requirements: Windows NT/2000/XP: Included in Windows NT 3.1 and later.
'                       Windows 95/98/Me: Included in Windows 95 and later.
'Notes: No current support for Above_Normal and Below_Normal
'        For Explanation of the Levels Look Below all Procedures
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
' RealTimePriority: Specify this class for a process that has the highest possible priority.
'                 The threads of the process preempt the threads of all other processes,
'                 including operating system processes performing important tasks.
'                 For example, a real-time process that executes for more than a very brief interval
'                 can cause disk caches not to flush or cause the mouse to be unresponsive.
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' HighPriority: Specify this class for a process that performs time-critical tasks that must be executed immediately.
'                 The threads of the process preempt the threads of normal or idle priority class processes.
'                 An example is the Task List, which must respond quickly when called by the user,
'                 regardless of the load on the operating system.
'                 Use extreme care when using the high-priority class,
'                 because a high-priority class application can use nearly all available CPU time.
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' NormalPriority: Specify this class for a process with no special scheduling needs.
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' IdlePriority: Specify this class for a process whose threads run only when the system is idle.
'                 The threads of the process are preempted by the threads of any process running in a higher priority class.
'                 An example is a screen saver.
'---------------------------------------------------------------------------------------
'
'Some API Declarations
Private Declare Function GetCurrentProcess Lib "kernel32" _
      () As Long
Private Declare Function SetPriorityClass Lib "kernel32" _
      (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function GetPriorityClass Lib "kernel32" _
      (ByVal hProcess As Long) As Long

'---------------------------------------------------------------------------------------
' Procedure : SetPriority
' Purpose   : Sets the Priority Level of the Current Program
'---------------------------------------------------------------------------------------
Private Function SetPriority(PriorityClass As String) As Long
1:    On Error GoTo Err

3:    SetPriority = SetPriorityClass(GetCurrentProcess, PriorityClass)
4:  Exit Function
5:
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "basPriority.SetPriority()"
End Function

Public Function SetPriorityLivel(Priority As Integer)
    On Error GoTo Err
2:  Dim X As Long

4:   Select Case Priority
        Case 0
6:         X = SetPriority(&H40)    ' Idle
7:         AddLog "Process(Application) Priority: Idle", 2
        Case 1
9:         X = SetPriority(&H20)    ' Normal
10:         AddLog "Process(Application) Priority: Normal", 2
        Case 2
12:        X = SetPriority(&H80)    ' High
13:        AddLog "Process(Application) Priority: High", 2
        Case 3
15:        X = SetPriority(&H100)   ' Realtime
16:        AddLog "Process(Application) Priority: Realtime", 2
     End Select

19:  Exit Function

Err:
22:  HandleError Err.Number, Err.Description, Erl & "|mSysPriority.SetPriorityLivel()"
End Function
    
'---------------------------------------------------------------------------------------
' Procedure : GetPriority
' Purpose   : Used to Retrieve the Current Priority Class
' Returns : String
'---------------------------------------------------------------------------------------
Private Function GetPriority() As Long
    On Error GoTo Err
2:    GetPriority = (GetPriorityClass(GetCurrentProcess))
    Exit Function
Err:
3:  HandleError Err.Number, Err.Description, Erl & "|mSysPriority.GetPriority()"
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetPriorityName
' Purpose   : Returns the Priority Level Name in Place of a Value as above
'---------------------------------------------------------------------------------------
Private Function GetPriorityName() As String
1:  On Error GoTo Err
2:   Dim lngPriority As Long
3:   lngPriority = GetPriority
   
5:   Select Case lngPriority
       Case 256
7:         GetPriorityName = "Realtime"
       Case 128
9:         GetPriorityName = "High"
       Case 32
11:        GetPriorityName = "Normal"
       Case 64
13:         GetPriorityName = "Idle"
     End Select
15: Exit Function
16:
Err:
18:    HandleError Err.Number, Err.Description, Erl & "|" & "basPriority.GetPriorityName()"
End Function

