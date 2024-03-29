Option Explicit

'About plugin DebugScript.dll
'Active Debugger is intended to debug script by using it as run time debugger 
'by print ing valuable information onto debugger's screen such as variable contents, 
'current state of any object etc..

'The colors of the window debug use the format QBColor of VB
'QB Color - Info:
	'0 = Black 
	'1 = Blue 
	'2 = Green 
	'3 = Cyan 
	'4 = Red 
	'5 = Magenta 
	'6 = Yellow 
	'7 = White 
	'8 = Gray 
	'9 = Light Blue 
	'10 = Light Green 
	'11 = Light Cyan 
	'12 = Light Red 
	'13 = Light Magenta 
	'14 = Light Yellow 
	'15 = Bright White

'Plugin object
Dim objPlg

Sub Main()

	'Create Object
	Set objPlg = CreateObject("DebugScript.Main")

	'This just allows the IDE to know if the plug-ins was loaded
	If objPlg.loadplug <> 1 Then
		MsgBoxCenter Me, "There was an error while loading the plugin.", vbCritical
		Exit Sub
	Else
		'Public Sub LoadPlugin(Optional mObject As Object)
		objPlg.LoadPlugin 'frmHub

		'Public Function Init(Optional sTitle As String = "Active Scripting Debugger Window", _
		'                     Optional iForeColor As Integer = 0, _
		'                     Optional iBackColor As Integer = 15, _
		'                     Optional sFontName As String = "Tahoma", _
		'                     Optional iFontSize As Integer = 9)
		objPlg.Init "Script Debug.vbs - Active Scripting Debugger Window", 0, 15, "Courier New", 10

		'Public Function DebugPrint(sMsg As String, _
		'                  Optional sProced As String = "", _
		'                  Optional sComment As String = "", _
		'                  Optional iColor As Long = -1, _
		'                  Optional bTime As Boolean = True, _
		'                  Optional bBold As Boolean = False, _
		'                  Optional bUnderline As Boolean = False)
		objPlg.DebugPrint "This is a real time debugger windows!!", "Main()", "No comments", -1, True, True, False
	End If

End Sub

Sub UnloadMain()
	'Public Sub UnloadPlugin()
	objPlg.UnloadPlugin
	Set objPlg = Nothing
End Sub

Sub Error(Line)

	'Public Function Clear()
	objPlg.Clear 'Clear debug windows

	Set objPlg = Nothing

	MsgBox Now & vbNewline & _
		"- Line: " & line & vbNewline & _
		"- Number: " & Err.Number & vbNewline & _
		"- "& Err.Description & vbNewline, "DebugScript"

End Sub  