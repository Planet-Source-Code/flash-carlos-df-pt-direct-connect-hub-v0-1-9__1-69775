'-------------------------------------------------------------------'
' 	Script for PTDCH and DDCH created by fLaSh - Carlos.DF	    '
' '-----------------------------------------------------------------'
' '		   						    '
' '			'----------------------'		    '
' '			'   Up Time Control    '-		    '
' '			'      Version 1.8     ' '		    '
' '			'        bY fLaSh      ' '		    '
' '			'----------------------' '		    '
' '			   '---------------------'		    '
' '								    '
' '-----------------------------------------------------------------'
' '		   E-mail: carlosferreiracarlos@hotmail.com	    '
' '			 Braga S.Victor - Portugal 		    '
' '			  DC Hub: ptdch.no-ip.org	    	    '
' '			    Release: 2007-11-22			    '
' '-----------------------------------------------------------------'
'
'History:
'	-v1.0: *initial release
'	-v1.2: *fixed smmal bug
'	-v1.4: *add new commands for Ops
'	-v1.6: *add new commands for Ops and Vips
'	       *add cconditions of compilations for Debug Mode
'	       *show Tops UpTime
'	-v1.8: *fixed smmal bug in tops uptime 
'	       *add formate for tops
'	       *add media for tops uptime
'	       *new table in db for tops uptime
'	       *add optimizations in all code
'

Option Explicit

'Script Settings '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Const M_iMinClass = 3		'Minimum class to use uptime control
Const M_bShowTopInfo = True	'Display info in the tops UpTime
Const M_iDisplayMsgOps = True	'Display info about commands used by users to Ops
Const M_bShowTopMedia = True	'Display uptime media in tops
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

'DEBUG MODE ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'Trun OFF (False) this compilation condition, 
'This is only used to develop the script..
#Const DEBUG_MODE = False
'WARNING: This Const Use High memory and cause SPAM!!
Const M_bDebugToLog = False	'
Const M_bDebugToUser = False	'
Const M_sDebugUser = "fLaSh"	'User who recieves the debug messages
Const M_sScriptVerison = "v1.8"  'Script version
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

'SQL connection object
Dim objJetCon

'Commands Prefix
Dim M_CmdPrefix

Sub Main()
 
#If DEBUG_MODE Then
	Call DebugPrint("UpTime Control Script Started ..")
#End If

	M_CmdPrefix = Chr(Settings.CPrefix)

	'Add language suport..
	Call AddScriptMessages(".\Scripts\UpTimeControl-Lib\UpTimeControl.xml")

	DoEvents

	'Open connection
	Call OpenJetCon(objJetCon)

	DoEvents

	'Add Hub commands
	Call AddHubCommands

End Sub

Sub UnloadMain()

	'Remove all commands ..
	With colCommands
	   'Normal
	   .Remove("UpTimeHub")
	   .Remove("UpTime")
	   .Remove("UpTimeTotal")
	   .Remove("UpTimeLast")
	   .Remove("UpTimeConn")
	   .Remove("UpTimeRecord")
	   .Remove("UpTimeTop")
	   'VIP
	   .Remove("UpTimeTop5")
	   .Remove("UpTimeTop10")
	   .Remove("UpTimeTop20")
	   'Op
	   .Remove("UpTimeUser")
	   .Remove("UpTimeTotalUser")
	   .Remove("UpTimeLastUser")
	   .Remove("UpTimeConnUser")
	   .Remove("UpTimeRecordUser")
	   .Remove("UpTimeTop5Hide")
	   .Remove("UpTimeTop10Hide")
	   .Remove("UpTimeTop20Hide")
	   .Remove("UpTimeTopAllHide")
	   .Remove("UpTimeTopAll")
	   'Super Op
	   .Remove("UpTimeClsName")
	   .Remove("UpTimeCls60")
	   'Admin
	   .Remove("UpTimeCls")
	End With

	DoEvents

	'----------------------------------------------------------------------------
	'NOTE:  not to use this code here, because this duplicates the uptime of session, 
	'	if the script just goes only restarted..
	'	It is more effective to use this code in Sub StoppedServing()
	'Dim strUser, objUser
	'Just case..
	'On Error Resume Next
	'UpDate DB now .. 
	'For Each strUser In colUsers
	'	'Up Time Control Only for users regs..
	'	Set objUser = colUsers.ItemByName(CStr(strUser.sName))
	'	If curUser.Class >= M_iMinClass Then _
	'		Call UserQuit(objUser)
	'Next
	'----------------------------------------------------------------------------

	objJetCon.Close
	Set objJetCon = Nothing

#If DEBUG_MODE Then
	Call DebugPrint("UpTime Control Script Stoped ..")
#End If

End Sub

Sub StoppedServing()
	Dim strUser, objUser
	'Just case..
	On Error Resume Next
	'UpDate DB now .. 
	For Each strUser In colUsers
		'Up Time Control Only for users regs..
		Set objUser = colUsers.ItemByName(CStr(strUser.sName))
		If curUser.Class >= M_iMinClass Then _
			Call UserQuit(objUser): DoEvents
	Next
End Sub

Sub OpConnected(curUser)
	Call UpTimeControl(curUser)
End Sub

Sub RegConnected(curUser)
	Call UpTimeControl(curUser)
End Sub

Sub UserConnected(curUser)
	Call UpTimeControl(curUser)
End Sub

Sub UserQuit(curUser)

	If curUser.Class < M_iMinClass Then Exit Sub
	
	Dim strQuery, strResult
	Dim intUpTime, intUpTimeSession, intUpTimeTotal
	Dim intUpTimeRecord
	Dim dtLastUpTime

	'Check if user found in the database..
	'This envent doesn't filter the strings entirely.
	'Therefore that am to add this code To avoid mistake.
	'***********************************************************************************
	strQuery = "SELECT UserName " & _
		   "FROM UpTime " & _
		   "WHERE UserName = '" & curUser.sName & "'"
	strResult = SelectSQL(strQuery, "UserName")
	'***********************************************************************************

	'If user name not found in db record..
	If Not strResult <> "" Then
	#If DEBUG_MODE Then
		Call DebugPrint("User name not found at UserQuit(" & curUser.sName & ")")
	#End If
		Exit Sub
	End If

	'Get UpTime Total
	'***********************************************************************************
	strQuery = "SELECT UpTimeTotal " & _
		   "FROM UpTime " & _
		   "WHERE UserName = '" & curUser.sName & "'"
	intUpTime = SelectSQL(strQuery, "UpTimeTotal")
	'***********************************************************************************
	'Just case..
	If IsNull(intUpTime) Then intUpTime = 1

	'Calculat UpTime in secunds from the loged in Hub
	intUpTimeSession = CalcDateToSec(curUser.ConnectedSince, Now)

	'Calculat total UpTime
	intUpTimeTotal = intUpTime + intUpTimeSession

	'Update new values..
	'***********************************************************************************
	strQuery = "UPDATE UpTime " & _
		   "SET UpTimeTotal = '" & intUpTimeTotal & "' " & _
	 	   "WHERE UserName = '" & curUser.sName & "' "
	UpDateSQL strQuery
	'***********************************************************************************

	'Get UpTime record in database
	'***********************************************************************************
	strQuery = "SELECT UpTimeRecord "& _
		   "FROM UpTime " & _
		   "WHERE UserName = '" & curUser.sName & "'"
	intUpTimeRecord = SelectSQL(strQuery, "UpTimeRecord")
	'***********************************************************************************
	'Just case..
	If IsNull(intUpTimeRecord) Then intUpTimeRecord = 1

	'Update new record if then..
	If intUpTimeRecord < intUpTimeSession Then
		'***********************************************************************************
		strQuery = "UPDATE UpTime " & _
			   "SET UpTimeRecord = " & intUpTimeSession & "," & _
			   "    UpTimeRecordDate = '" & curUser.ConnectedSince & " @ " & Now & "' " & _
			   "WHERE UserName = '" & curUser.sName & "'"
		UpDateSQL strQuery
		'***********************************************************************************
	End If

	dtLastUpTime = CalcDateToSec(curUser.ConnectedSince, Now)

	'Update last access and last UpTime
	'***********************************************************************************
	strQuery = "UPDATE UpTime " & _
		   "SET LastUpTimeDate = '" & curUser.ConnectedSince & " @ " & Now & "'" & _
		   ", LastUpTime = '" & dtLastUpTime & "'" & _
		   ", LastAccessDate = '" & Now & "' " & _
		   "WHERE UserName = '" & curUser.sName & "'"
	UpDateSQL strQuery
	'***********************************************************************************

#If DEBUG_MODE Then
	Call DebugPrint("Record updated at exit.. User name: " & curUser.sName)
#End If

End Sub

Sub CustComArrival(curUser, objCommand, sData, blnMC)

	Dim strQuery
	Dim strUpTimeHub, strUpTime, intUpTimeSession, strUpTimeTotal
	Dim strLastUpTime, strLastUpTimeDate
	Dim intUpTimeConn
	Dim strUpTimeRecord, strUpTimeRecordDate
	Dim strUserRecord, strUserName
	Dim objRS
	Dim strSplit
	Dim objUser
	Dim strTemp
	Dim i
	
	'Up Time Control Only for users regs..
	If curUser.Class < M_iMinClass Then Exit Sub

	Select Case Cstr(objCommand.Name)
		Case "UpTimeHub", "UpTime", "UpTimeUser", "UpTimeTotal", "UpTimeTotalUser", _
		     "UpTimeLast", "UpTimeLastUser", "UpTimeConn", "UpTimeConnUser", _
		     "UpTimeRecord", "UpTimeRecordUser", "UpTimeClsName", "UpTimeCls60", _
		     "UpTimeCls", "UpTimeTop", "UpTimeTop5", "UpTimeTop10", "UpTimeTop20", _
		     "UpTimeTop5Hide", "UpTimeTop10Hide", "UpTimeTop20Hide", "UpTimeTopAllHide", _
		     "UpTimeTopAll"

			'Inform the Operators..
			If M_iDisplayMsgOps Then 
				colUsers.SendChatToOps Settings.BotName & " (OPs)", _
					curUser.GetCoreMsgStr("UPSendCmd") & curUser.sName & _
					" (" & M_CmdPrefix & objCommand.Name & ")"
			End If

			#If DEBUG_MODE Then
				Call DebugPrint("Hub Command (" & M_CmdPrefix & objCommand.Name & ") sending for user name: " & curUser.sName)
			#End If

	End Select


	'Start Hub commands
	Select Case Cstr(objCommand.Name)
	
		'====================================================================
		Case "UpTimeHub" 
		'====================================================================
		
			strUpTimeHub = UpTime(frmHub.ServingDate, curUser.sName, 0)
			Reply curUser, curUser.GetCoreMsgStr("UPTimeHub") & strUpTimeHub, blnMC

		'====================================================================
		Case "UpTime", "UpTimeUser"
		'====================================================================
		
			If CStr(objCommand.Name) = "UpTime" Then
				strUpTime = UpTime(curUser.ConnectedSince, curUser.sName, 0)
				Reply curUser, curUser.GetCoreMsgStr("UPTime") & strUpTime, blnMC
			ElseIf CStr(objCommand.Name) = "UpTimeUser" Then
				strSplit = AfterFirst(sData, " ")
				If TestStringOk(strSplit) Then
					If colUsers.OnLine(CStr(strSplit)) Then
						Set objUser = colUsers.ItemByName(CStr(strSplit))
						strUpTime = UpTime(objUser.ConnectedSince, objUser.sName, 0)
						Reply curUser, Replace(curUser.GetCoreMsgStr("UPTimeUser"), "%[user]", objUser.sName) & strUpTime, blnMC
					Else
						Reply curUser, Replace(curUser.GetCoreMsgStr("UPOffLine"), "%[user]", strSplit), blnMC
						Exit Sub
					End If
				Else
					Reply curUser, curUser.GetCoreMsgStr("UPSintax") & M_CmdPrefix & objCommand.Name & " <user>", blnMC
					Exit Sub
				End If
			End If
			
		'====================================================================	
		Case "UpTimeTotal", "UpTimeTotalUser"
		'====================================================================
		
			If CStr(objCommand.Name) = "UpTimeTotal" Then
				'Get total time loged and convert for secunds..
				intUpTimeSession = CalcDateToSec(curUser.ConnectedSince, Now)
				strUserName = curUser.sName
			ElseIf CStr(objCommand.Name) = "UpTimeTotalUser" Then
				strSplit = AfterFirst(sData, " ")
				If TestStringOk(strSplit) Then
					If colUsers.OnLine(CStr(strSplit)) Then
						Set objUser = colUsers.ItemByName(CStr(strSplit))
						intUpTimeSession = CalcDateToSec(objUser.ConnectedSince, Now)
						strUserName = objUser.sName
					Else
						Reply curUser, Replace(curUser.GetCoreMsgStr("UPOffLine"), "%[user]", strSplit), blnMC
						Exit Sub
					End If
				Else
					Reply curUser, curUser.GetCoreMsgStr("UPSintax") & M_CmdPrefix & objCommand.Name & " <user>", blnMC
					Exit Sub
				End If
			End If
			'Get UpTime Total
			'***********************************************************************************
			strQuery = "SELECT UpTimeTotal " & _
				   " FROM UpTime " & _
				   " WHERE UserName = '" & strUserName & "'"
			strUpTime = SelectSQL(strQuery, "UpTimeTotal") 
			'***********************************************************************************
			'Just case..
			If IsNull(strUpTime) Then strUpTime = 1
			strUpTimeTotal = strUpTime + intUpTimeSession
			'Convert for secunds and send msg to main chat..
			If CStr(objCommand.Name) = "UpTimeTotal" Then  
				strUpTimeTotal = UpTime(strUpTimeTotal, curUser.sName, 0) 
				Reply curUser, curUser.GetCoreMsgStr("UPTotal") & strUpTimeTotal, blnMC
			Else
				strUpTimeTotal = UpTime(strUpTimeTotal, objUser.sName, 0)
				Reply curUser, Replace(curUser.GetCoreMsgStr("UPTotalUser"), "%[user]", objUser.sName) & strUpTimeTotal, blnMC
			End If
			
		'====================================================================	
		Case "UpTimeLast", "UpTimeLastUser"
		'====================================================================
		
			If CStr(objCommand.Name) = "UpTimeLast" Then
				strUserName = curUser.sName
			ElseIf CStr(objCommand.Name) = "UpTimeLastUser" Then
				strSplit = AfterFirst(sData, " ")
				If TestStringOk(strSplit) Then
					If colUsers.OnLine(CStr(strSplit)) Then
						Set objUser = colUsers.ItemByName(CStr(strSplit))
						strUserName = objUser.sName
					Else
						Reply curUser, Replace(curUser.GetCoreMsgStr("UPOffLine"), "%[user]", strSplit), blnMC
						Exit Sub
					End If
				Else
					Reply curUser, curUser.GetCoreMsgStr("UPSintax") & M_CmdPrefix & objCommand.Name & " <user>", blnMC
					Exit Sub
				End If
			End If
			'Get Last upTime
			'***********************************************************************************
			strQuery = "SELECT LastUpTimeDate, LastUpTime " & _
				   " FROM UpTime " & _
				   " WHERE UserName = '" & strUserName & "'"
			strLastUpTime = SelectSQL(strQuery, "LastUpTime") 
			strLastUpTimeDate = SelectSQL(strQuery, "LastUpTimeDate")
			'***********************************************************************************
			If strLastUpTime <> "" Or strLastUpTimeDate <> "" Then
				strLastUpTime = UpTime(strLastUpTime, strUserName, 0)
				If CStr(objCommand.Name) = "UpTimeLast" Then
					strLastUpTimeDate = Replace(curUser.GetCoreMsgStr("UPLast"), "%[date]", strLastUpTimeDate)
					strTemp = strLastUpTimeDate & strLastUpTime
				Else
					strTemp = Replace(curUser.GetCoreMsgStr("UPLastUser"), "%[date]", strLastUpTimeDate)
					strTemp = Replace(strTemp, "%[user]", strUserName)
					strTemp = strTemp & strLastUpTime
				End If
				Reply curUser, strTemp, blnMC
			Else
				Reply curUser, curUser.GetCoreMsgStr("UPNoData"), blnMC
			End If
			
		'====================================================================	
		Case "UpTimeConn", "UpTimeConnUser"
		'====================================================================
		
			If CStr(objCommand.Name) = "UpTimeConn" Then
				strUserName = curUser.sName
			ElseIf CStr(objCommand.Name) = "UpTimeConnUser" Then
				strSplit = AfterFirst(sData, " ")
				If TestStringOk(strSplit) Then
					If colUsers.OnLine(CStr(strSplit)) Then
						Set objUser = colUsers.ItemByName(CStr(strSplit))
						strUserName = objUser.sName
					Else
						Reply curUser, Replace(curUser.GetCoreMsgStr("UPOffLine"), "%[user]", strSplit), blnMC
						Exit Sub
					End If
				Else
					Reply curUser, curUser.GetCoreMsgStr("UPSintax") & M_CmdPrefix & objCommand.Name & " <user>", blnMC
					Exit Sub
				End If
			End If
			'Get upTime connections
			'***********************************************************************************
			strQuery = "SELECT * " & _
				   " FROM UpTime " & _
				   " WHERE UserName = '" & strUserName & "'"
			intUpTimeConn = SelectSQL(strQuery, "TotalConnections")
			If Not intUpTimeConn <> "" Then
				Reply curUser, curUser.GetCoreMsgStr("UPNoData"), blnMC
				Exit Sub
			End If
			If Not IsNull(intUpTimeConn) Then
				If CStr(objCommand.Name) = "UpTimeConn" Then _
				     Reply curUser, curUser.GetCoreMsgStr("UPConnect") & intUpTimeConn, blnMC _
				Else Reply curUser, Replace(curUser.GetCoreMsgStr("UPConnectUser"), "%[user]", strUserName) & intUpTimeConn, blnMC
			Else
				Reply curUser, curUser.GetCoreMsgStr("UPNoData"), blnMC
			End If
			
		'====================================================================	
		Case "UpTimeRecord", "UpTimeRecordUser"
		'====================================================================
		
			If CStr(objCommand.Name) = "UpTimeRecord" Then
				strUserName = curUser.sName
			ElseIf CStr(objCommand.Name) = "UpTimeRecordUser" Then
				strSplit = AfterFirst(sData, " ")
				If TestStringOk(strSplit) Then
					If colUsers.OnLine(CStr(strSplit)) Then
						Set objUser = colUsers.ItemByName(CStr(strSplit))
						strUserName = objUser.sName
					Else
						Reply curUser, Replace(curUser.GetCoreMsgStr("UPOffLine"), "%[user]", strSplit), blnMC
						Exit Sub
					End If
				Else
					Reply curUser, curUser.GetCoreMsgStr("UPSintax") & M_CmdPrefix & objCommand.Name & " <user>", blnMC
					Exit Sub
				End If
			End If
			'Get Record upTime 
			'***********************************************************************************
			strQuery = "SELECT UpTimeRecord, UpTimeRecordDate " & _
				   "FROM UpTime " & _
				   "WHERE UserName = '" & strUserName & "'"
			strUpTimeRecord = SelectSQL(strQuery, "UpTimeRecord") 
			strUpTimeRecordDate = SelectSQL(strQuery, "UpTimeRecordDate")
			'***********************************************************************************
			If Not IsNull(strUpTimeRecord) And Not IsNull(strUpTimeRecordDate) Then
				If Not strUpTimeRecord <> "" Then
					Reply curUser, curUser.GetCoreMsgStr("UPNoData"), blnMC
					Exit Sub
				End If
				strUpTimeRecord = UpTime(strUpTimeRecord, strUserName, 0)
				If CStr(objCommand.Name) = "UpTimeRecord" Then
					strUpTimeRecordDate = Replace(curUser.GetCoreMsgStr("UPTop"), "%[date]", strUpTimeRecordDate)
				Else
					strTemp = Replace(curUser.GetCoreMsgStr("UPTopUser"), "%[date]", strUpTimeRecordDate)
					strTemp = Replace(strTemp, "%[user]", strUserName)
					strUpTimeRecordDate = strTemp
				End If
				Reply curUser, strUpTimeRecordDate & strUpTimeRecord, blnMC
			Else
				Reply curUser, curUser.GetCoreMsgStr("UPNoData"), blnMC
			End If
			
		'====================================================================	
		Case "UpTimeClsName"
		'====================================================================
		
			strSplit = AfterFirst(sData, " ")
			If TestStringOk(strSplit) Then
				'Check if user found in the database..
				strQuery = "SELECT UserName " & _
					   "FROM UpTime " & _
					   "WHERE UserName = '" & strSplit & "'"
				strTemp = SelectSQL(strQuery, "UserName")
				'If found recourd..
				If LenB(strTemp) Then
					strQuery = "DELETE FROM UpTime " & _
						   "WHERE UserName = '" & strTemp & "'"
					DeleteSQL(strQuery)
					'
					Reply curUser, Replace(curUser.GetCoreMsgStr("UPClsUser"), "%[user]", strTemp), blnMC
				Else
					Reply curUser, curUser.GetCoreMsgStr("UPNoData"), blnMC
				End If
			Else
				Reply curUser, curUser.GetCoreMsgStr("UPSintax") & M_CmdPrefix & objCommand.Name & " <user>", blnMC
			End If
			
		'====================================================================	
		Case "UpTimeCls60"
		'====================================================================
		
			' Accounts older than this amount 60 days will be deleted
			
			'Set temp date..
			strTemp = DateAdd("d", -60, Now)
			
			strQuery = "SELECT * FROM UpTime " & _
				   "WHERE LastAccessDate > '" & strTemp & "'"
					   
			'Get Records afecteds..		   
			i = GetAfectedRowsSQL(strQuery)
			
			'Delete Recourds ..
			strQuery = "DELETE FROM UpTime " & _
				   "WHERE LastAccessDate > '" & strTemp & "'"
			DeleteSQL strQuery
			
			If i <> 0 Then
				Reply curUser, curUser.GetCoreMsgStr("UPCls60") & _
					" " & Replace(curUser.GetCoreMsgStr("UPClsInfo"), "%[i]", i), blnMC
			Else
				Reply curUser, curUser.GetCoreMsgStr("UPNoData"), blnMC	
			End If

		'====================================================================	
		Case "UpTimeCls"
		'====================================================================
		
			'Clear BD (Reset UpTime Control)
			
			'Get Records afecteds..
			i = GetAfectedRowsSQL("SELECT * FROM UpTime")

			strQuery = "DELETE FROM UpTime "
			DeleteSQL strQuery
			
			If i <> 0 Then
				'Send Chat To All users
				colUsers.SendChatToAll Settings.BotName, curUser.GetCoreMsgStr("UPClsAll")
				'Reply total registers deleted..
				Reply curUser, Replace(curUser.GetCoreMsgStr("UPClsInfo"), "%[i]", i), blnMC
			Else
				Reply curUser, curUser.GetCoreMsgStr("UPNoData"), blnMC	
			End If
			
		'====================================================================
		Case "UpTimeTop", "UpTimeTop5", "UpTimeTop10", "UpTimeTop20"
		'====================================================================
		
			If CStr(objCommand.Name) = "UpTimeTop" Then
				'*************************************************************
				'Get Top UpTime 
				strQuery = "SELECT UpTimeTotal, UserName " & _ 
					   "FROM UpTime " & _
					   "WHERE UpTimeTotal In(SELECT Max(UpTimeTotal) FROM UpTime)" 
				
				Set objRS = objJetCon.Execute(strQuery)

				If Not objRS.EOF Then
					strUpTime = objRS.Fields("UpTimeTotal")
					strUserName = objRS.Fields("UserName")
					
					If strUpTime <> "" And strUserName <> "" Then
						strUpTime = UpTime(strUpTime, curUser.sName, 0)
						strUserName = Replace(curUser.GetCoreMsgStr("UPTopRecord"), "%[user]", strUserName)
						Reply curUser, strUserName & strUpTime, blnMC
					Else
						Reply curUser, curUser.GetCoreMsgStr("UPNoData"), blnMC
					End If
				Else
					Reply curUser, curUser.GetCoreMsgStr("UPNoData"), blnMC
				End If

				Set objRS = Nothing
				'*************************************************************

			ElseIf CStr(objCommand.Name) = "UpTimeTop5" Then
				Call PreTopsUpTime(5, curUser, blnMC, False)
			ElseIf CStr(objCommand.Name) = "UpTimeTop10" Then
				Call PreTopsUpTime(10, curUser, blnMC, False)
			ElseIf CStr(objCommand.Name) = "UpTimeTop20" Then
				Call PreTopsUpTime(20, curUser, blnMC, False)
			End If

		'====================================================================
		Case "UpTimeTop5Hide", "UpTimeTop10Hide", "UpTimeTop20Hide", "UpTimeTopAllHide", "UpTimeTopAll"
		'====================================================================

			If CStr(objCommand.Name) = "UpTimeTop5Hide" Then
				Call PreTopsUpTime(5, curUser, blnMC, True)
			ElseIf CStr(objCommand.Name) = "UpTimeTop10Hide" Then
				Call PreTopsUpTime(10, curUser, blnMC, True)
			ElseIf CStr(objCommand.Name) = "UpTimeTop20Hide" Then
				Call PreTopsUpTime(20, curUser, blnMC, True)
			ElseIf CStr(objCommand.Name) = "UpTimeTopAllHide" Then
				Call PreTopsUpTime("All", curUser, blnMC, True)
			ElseIf CStr(objCommand.Name) = "UpTimeTopAll" Then
				Call PreTopsUpTime("All", curUser, blnMC, False)
			End If

	End Select

End Sub

Sub AddHubCommands()

	With colCommands
	   'Reg
	   If Not .Exists("UpTimeHub") Then .Add 6667, "UpTimeHub", "UPDesc1", 3, True
	   If Not .Exists("UpTime") Then .Add 6668, "UpTime", "UPDesc2", 3, True
	   If Not .Exists("UpTimeTotal") Then .Add 6669, "UpTimeTotal", "UPDesc3", 3, True
	   If Not .Exists("UpTimeLast") Then .Add 6670, "UpTimeLast", "UPDesc4", 3, True
	   If Not .Exists("UpTimeConn") Then .Add 6671, "UpTimeConn", "UPDesc5", 3, True
	   If Not .Exists("UpTimeRecord") Then .Add 6672, "UpTimeRecord", "UPDesc6", 3, True
	   If Not .Exists("UpTimeTop") Then .Add 6673, "UpTimeTop", "UPDesc7", 3, True
	   'Vip
	   If Not .Exists("UpTimeTop5") Then .Add 6674, "UpTimeTop5", "UPDesc8", 5, True
	   If Not .Exists("UpTimeTop10") Then .Add 6675, "UpTimeTop10", "UPDesc9", 5, True
	   If Not .Exists("UpTimeTop20") Then .Add 6676, "UpTimeTop20", "UPDesc10", 5, True
	   'OP
	   If Not .Exists("UpTimeUser") Then .Add 6677, "UpTimeUser", "UPDesc11", 6, True
	   If Not .Exists("UpTimeTotalUser") Then .Add 6678, "UpTimeTotalUser", "UPDesc12", 6, True
	   If Not .Exists("UpTimeLastUser") Then .Add 6679, "UpTimeLastUser", "UPDesc13", 6, True
	   If Not .Exists("UpTimeConnUser") Then .Add 6680, "UpTimeConnUser", "UPDesc14", 6, True
	   If Not .Exists("UpTimeRecordUser") Then .Add 6681, "UpTimeRecordUser", "UPDesc15",6 , True
	   If Not .Exists("UpTimeTop5Hide") Then .Add 6682, "UpTimeTop5Hide", "UPDesc16", 6, True
	   If Not .Exists("UpTimeTop10Hide") Then .Add 6683, "UpTimeTop10Hide", "UPDesc17", 6, True
	   If Not .Exists("UpTimeTop20Hide") Then .Add 6684, "UpTimeTop20Hide", "UPDesc18", 6, True
	   If Not .Exists("UpTimeTopAllHide") Then .Add 6685, "UpTimeTopAllHide", "UPDesc19",6 , True
	   If Not .Exists("UpTimeTopAll") Then .Add 6686, "UpTimeTopAll", "UPDesc20",6 , True
	   'Super Op
	   If Not .Exists("UpTimeClsName") Then .Add 6687, "UpTimeClsName", "UPDesc21", 8, True
	   If Not .Exists("UpTimeCls60") Then .Add 6688, "UpTimeCls60", "UPDesc22", 8, True
	   'Admin
	   If Not .Exists("UpTimeCls") Then .Add 6689, "UpTimeCls", "UPDesc23", 10, True
	End With

End Sub

Sub AddRigthClickMenu(curUser)

	'Add rigth click menu
	With curUser
	   If .Class >= 3 Then 'Registereds or eight..
		.SendData "$UserCommand 1 3 UpTime Control\UpTime Hub$<%[mynick]> " & M_CmdPrefix & "UpTimeHub&#124;|"
		.SendData "$UserCommand 1 3 UpTime Control\Your UpTime$<%[mynick]> " & M_CmdPrefix & "UpTime&#124;|"
		.SendData "$UserCommand 1 3 UpTime Control\Your UpTime Total$<%[mynick]> " & M_CmdPrefix & "UpTimeTotal&#124;|"
		.SendData "$UserCommand 1 3 UpTime Control\Your Last UpTime$<%[mynick]> " & M_CmdPrefix & "UpTimeLast&#124;|"
		.SendData "$UserCommand 1 3 UpTime Control\Your UpTime Connections$<%[mynick]> " & M_CmdPrefix & "UpTimeConn&#124;|"
		.SendData "$UserCommand 1 3 UpTime Control\Your UpTime Record$<%[mynick]> " & M_CmdPrefix & "UpTimeRecord&#124;|"
		.SendData "$UserCommand 1 3 UpTime Control\UpTime Top$<%[mynick]> " & M_CmdPrefix & "UpTimeTop&#124;|"
	   End If
	   If .Class >= 5 Then 'Vips or eight..
		.SendData "$UserCommand 1 3 UpTime Control\VIP\UpTime Top5$<%[mynick]> " & M_CmdPrefix & "UpTimeTop5&#124;|"
		.SendData "$UserCommand 1 3 UpTime Control\VIP\UpTime Top10$<%[mynick]> " & M_CmdPrefix & "UpTimeTop10&#124;|"
		.SendData "$UserCommand 1 3 UpTime Control\VIP\UpTime Top20$<%[mynick]> " & M_CmdPrefix & "UpTimeTop20&#124;|"		
	   End If
	   If .Class >= 6 Then 'Operatores or eight..
		.SendData "$UserCommand 1 3 UpTime Control\Operator\Get User UpTime$<%[mynick]> " & M_CmdPrefix & "UpTimeUser %[nick]&#124;|"
		.SendData "$UserCommand 1 3 UpTime Control\Operator\Get User UpTime Total$<%[mynick]> " & M_CmdPrefix & "UpTimeTotalUser %[nick]&#124;|"
		.SendData "$UserCommand 1 3 UpTime Control\Operator\Get User UpTime Last$<%[mynick]> " & M_CmdPrefix & "UpTimeLastUser %[nick]&#124;|"
		.SendData "$UserCommand 1 3 UpTime Control\Operator\Get User UpTime Connections$<%[mynick]> " & M_CmdPrefix & "UpTimeConnUser %[nick]&#124;|"
		.SendData "$UserCommand 1 3 UpTime Control\Operator\Get User UpTime Record $<%[mynick]> " & M_CmdPrefix & "UpTimeRecordUser %[nick]&#124;|"
		.SendData "$UserCommand 1 3 UpTime Control\Operator\UpTime Top5 (Hide)$<%[mynick]> " & M_CmdPrefix & "UpTimeTop5Hide&#124;|"
		.SendData "$UserCommand 1 3 UpTime Control\Operator\UpTime Top10 (Hide)$<%[mynick]> " & M_CmdPrefix & "UpTimeTop10Hide&#124;|"
		.SendData "$UserCommand 1 3 UpTime Control\Operator\UpTime Top20 (Hide)$<%[mynick]> " & M_CmdPrefix & "UpTimeTop20Hide&#124;|"	
		.SendData "$UserCommand 1 3 UpTime Control\Operator\UpTime Top All (Hide)$<%[mynick]> " & M_CmdPrefix & "UpTimeTopAllHide&#124;|"
		.SendData "$UserCommand 1 3 UpTime Control\Operator\UpTime Top All$<%[mynick]> " & M_CmdPrefix & "UpTimeTopAll&#124;|"
	   End If
	   If .Class >= 7 Then 'Super Op or eight.. 
		.SendData "$UserCommand 1 3 UpTime Control\Super Operator\Clear on register by name$<%[mynick]> " & M_CmdPrefix & "UpTimeClsName %[nick]&#124;|"
		.SendData "$UserCommand 1 3 UpTime Control\Super Operator\Clear register under 60 days$<%[mynick]> " & M_CmdPrefix & "UpTimeCls60&#124;|"
	   End If
	   If .Class >= 10 Then 'Admin
		.SendData "$UserCommand 1 3 UpTime Control\Admin\Clear DB $<%[mynick]> " & M_CmdPrefix & "UpTimeCls&#124;|"
	   End If
	End With

End Sub

Function UpTimeControl(curUser) 

	Dim strQuery
	Dim strUpTimeHub
	Dim strResult
	Dim objRS
	Dim intTimeTotal, intTotalConn

	If curUser.Class < M_iMinClass Then Exit Function

	'Add rigth click menu
	Call AddRigthClickMenu(curUser)

	'Check if user found in the database..
	'***********************************************************************************
	strQuery = "SELECT UserName " & _
		   "FROM UpTime " & _
	           "WHERE UserName = '" & curUser.sName & "'"
	strResult=SelectSQL(strQuery, "UserName")
	'***********************************************************************************

	'If not found recourd..
	If Not strResult <> "" Then 
		'***********************************************************************************
		strQuery = "INSERT INTO UpTime(UserName, UpTimeTotal, LastUpTime, TotalConnections, UpTimeRecord, InitialAccess) " & _
			              "VALUES('" & curUser.sName & "',0 , 0, 1 , 0,'" & Now & "')"
		InsertSQL strQuery
		'***********************************************************************************

	#If DEBUG_MODE Then
		Call DebugPrint("New record add to data base. User name: " & curUser.sName)
	#End If

	Else 'If found recourd..
		'Get Connections Total
		'***********************************************************************************
		strQuery = "SELECT TotalConnections " & _
				"FROM UpTime " & _
				"WHERE UserName = '" & curUser.sName & "'"
		intTotalConn = SelectSQL(strQuery, "TotalConnections")
		'***********************************************************************************
		If IsNull(intTotalConn) Then _
			intTotalConn = 1 _
		Else	intTotalConn = CInt(intTotalConn + 1)

		'Update new values..
		'***********************************************************************************
		strQuery = "UPDATE UpTime " & _
			   "SET TotalConnections = " & intTotalConn & " " & _
			   "WHERE UserName = '" & curUser.sName & "'"
		UpDateSQL strQuery
		'***********************************************************************************

	#If DEBUG_MODE Then
		Call DebugPrint("Record updated at login.. User name: " & curUser.sName)
	#End If
	
	End If

	DoEvents

End Function

Function UpTime(varData, strUser, intFormat)

	Dim lngYears, lngMonths, lngWeeks, lngDays, lngHours, lngMinutes, lngSeconds
	Dim lngTime
	Dim strTemp
	Dim objUser

	'Set user object..
	Set objUser = colUsers.ItemByName(CStr(strUser))

	If IsNumeric(varData) Then
		'if varData is secunds..
		lngTime = varData
	Else 'ElseIf IsDate(varData) Then
		'calc.. time .. basead in secunds..
		lngTime = DateDiff("s", CDate(varData), Now)
	End If

	'Convert values..
	lngSeconds = (lngTime) Mod 60
	lngMinutes = (lngTime \ 60) Mod 60
	lngHours = (lngTime \ 3600) Mod 24
	lngDays = (lngTime \ 86400) Mod 7
	lngWeeks = (lngTime \ 604800) Mod 4
	lngMonths = (lngTime \ 2419200) Mod 12
	lngYears = (lngTime \ 29030400)
	
	'Generate string *******************************************************************************
	'
	'String format ex: 1 year(s), 2 week(s), 3 day(s), 4 hour(s), 5 minute(s) and seconds
	If intFormat = 0 Then 
		If lngYears > 0 Then _
			strTemp = lngYears & " " & objUser.GetCoreMsgStr("UPYears") & ", "
		If lngMonths > 0 Then _
			strTemp = strTemp & lngMonths & " " & objUser.GetCoreMsgStr("UPMonths") & ", "
		If lngWeeks > 0 Then _
			strTemp = strTemp & lngWeeks & " " & objUser.GetCoreMsgStr("UPWeeks") & ", "
		If lngDays > 0 Then _
			strTemp = strTemp & lngDays & " " & objUser.GetCoreMsgStr("UPDays") & ", "
		If lngHours > 0 Then _
			strTemp = strTemp & lngHours & " " & objUser.GetCoreMsgStr("UPHours") & ", "
		If lngMinutes > 0 Then _
			strTemp = strTemp & lngMinutes & " " & objUser.GetCoreMsgStr("UPMinutes") & _
				  " " & objUser.GetCoreMsgStr("UPand") & " " & _
				  lngSeconds & " " & objUser.GetCoreMsgStr("UPSecunds")
	'
	'String format ex: [Y:year][W:week][D:day] HH:MM:SS
	ElseIf intFormat = 1 Then
		'Years
		If lngYears > 0 Then _
			strTemp = "[Y:" & lngYears & "] "
		'Months
		strTemp = strTemp & "[M:" & lngMonths & "] "
		'Weeks
		strTemp = strTemp & "[W:" & lngWeeks & "] "
		'Days 
		strTemp = strTemp & "[D:" & lngDays & "] "
		'Hours
		strTemp = strTemp & " " & RZero(lngHours, 2)
		'Minutes
		strTemp = strTemp & ":" & RZero(lngMinutes, 2)
		'Secunds
		strTemp = strTemp & ":" & RZero(lngSeconds, 2)
	End If
	'
	'***********************************************************************************************

	DoEvents

	'Return UpTime to String..
	UpTime = strTemp

End Function

Sub PreTopsUpTime(varNTop, curUser, blnMC, blnHide)

	'I created this extra Sub, because it was not to be very orderly by query, 
	'it is like this everything orderly and actualizado until the command to 
	'be run by the user.
	'Now a temporary table will only be used to generate all the necessary information.
	'In the end, this same data will be removed.
	'	Temporary table -->UpTimeTmp

	Dim objRS
	Dim strQuery
	Dim objUser
	Dim lngTotalUpTime
	Dim strTotalUpTime
	Dim blnUpTimeRecord
	Dim strUpTimeRecord
	Dim strTemp
	Dim strOnLine
	Dim i
	
	'Delete temp table data
	'*****************************************************
	Set objRS = objJetCon.Execute("DELETE FROM UpTimeTmp")
	Set objRS = Nothing
#If DEBUG_MODE Then
	Call DebugPrint("DELETE FROM UpTimeTmp (PreTopsUpTime)")
#End If
	'*****************************************************

	If IsNumeric(varNTop) Then _
		strTemp = "SELECT TOP " & varNTop & " * " _
	Else	strTemp = "SELECT * "

	strQuery = strTemp & _
		   "FROM UpTime " & _
		   "ORDER BY UpTimeTotal DESC"

	Set objRS = objJetCon.Execute(strQuery)

	'Note: DB Params
	'	- objRS.Collect(0) = UserName
	'	- objRS.Collect(1) = UpTimeTotal
	'	- objRS.Collect(2) = LastUpTime
	'	- objRS.Collect(3) = LastUpTimeDate
	'	- objRS.Collect(4) = UpTimeRecord
	'	- objRS.Collect(5) = UpTimeRecordDate
	'	- objRS.Collect(6) = TotalConnections
	'	- objRS.Collect(7) = LastAccessDate
	'	- objRS.Collect(8) = InitialAccess

	If Not objRS.EOF Then 
		Do Until objRS.EOF
			'If the user online.. 
			If colUsers.OnLine(Cstr(objRS.Collect(0))) Then
	
				'up date user uptime 
				Set objUser = colUsers.ItemByName(CStr(objRS.Collect(0)))
				lngTotalUpTime = objRS.Collect(1) + CalcDateToSec(objUser.ConnectedSince, Now)
				
				'if UpTime record this session is height..
				If objRS.Collect(4) < CalcDateToSec(objUser.ConnectedSince, Now) Then _
					blnUpTimeRecord = True _
				Else	blnUpTimeRecord = False

				strOnLine = "True"

			Else
				lngTotalUpTime = objRS.Collect(1)

				strOnLine = "False"
			End If
				
			If blnUpTimeRecord Then _
				strUpTimeRecord = "'" & CalcDateToSec(objUser.ConnectedSince, Now) & "', '" & Now & " @ ???? " & "', " _
			Else	strUpTimeRecord = "'" & objRS.Collect(4) & "', '" & objRS.Collect(5) & "', "

			'***********************************************************************************
			strQuery = "INSERT INTO UpTimeTmp(UserName, UpTimeTotal, " & _
							 "LastUpTime, LastUpTimeDate, " & _
							 "UpTimeRecord, UpTimeRecordDate, " & _
							 "TotalConnections, LastAccessDate, " & _
							 "InitialAccess, OnLine) " & _
				   "VALUES('" & objRS.Collect(0) & "', '" & lngTotalUpTime & "', " & _
					  "'" & objRS.Collect(2) & "', '" & objRS.Collect(3) & "', " & _
						strUpTimeRecord & _
					  "'" & objRS.Collect(6) & "', '" & objRS.Collect(7) & "', " & _
					  "'" & objRS.Collect(8) & "', '" & strOnLine & "')"

			InsertSQL strQuery
			'***********************************************************************************

			'Move next record..
			objRS.MoveNext

			'Refresh process..
			DoEvents
		Loop
	Else

		Reply curUser, curUser.GetCoreMsgStr("UPNoData"), blnMC
		Set objRS = Nothing
		Set objUser = Nothing
		Exit Sub

	End If

	Set objRS = Nothing
	Set objUser = Nothing
	
	DoEvents

	Call TopsUpTime(varNTop, curUser, blnMC, blnHide)

End Sub

Sub TopsUpTime(varNTop, curUser, blnMC, blnHide)

	Dim objRS
	Dim strQuery
	Dim strTemp
	Dim srtMedia(2)
	Dim objUser
	Dim strTotalUpTime, strUpTimeRecord
	Dim lngRanking
	Dim strOnLine
	Dim strLine, strTempLine
	Dim Tab1, Tab2, Tab3, Tab4, Tab5
	Dim i

	Tab1 = vbTab 
	Tab2 = vbTab + vbTab
	Tab3 = vbTab + vbTab + vbTab
	Tab4 = vbTab + vbTab + vbTab + vbTab
	Tab5 = vbTab + vbTab + vbTab + vbTab + vbTab
	
	strLine = String(210, "-")

	strQuery = "SELECT *" & _
		   "FROM UpTimeTmp " & _
		   "WHERE UpTimeTotal > 60 " & _
		   "ORDER BY UpTimeTotal DESC"

	'Note:  "WHERE UpTimeTotal >= 60" ->If larger than one minute..

	Set objRS = objJetCon.Execute(strQuery)

	'Note: DB Params
	'	- objRS.Collect(0) = UserName
	'	- objRS.Collect(1) = UpTimeTotal
	'	- objRS.Collect(2) = LastUpTime
	'	- objRS.Collect(3) = LastUpTimeDate
	'	- objRS.Collect(4) = UpTimeRecord
	'	- objRS.Collect(5) = UpTimeRecordDate
	'	- objRS.Collect(6) = TotalConnections
	'	- objRS.Collect(7) = LastAccessDate
	'	- objRS.Collect(8) = InitialAccess
	'	- objRS.Collect(9) = OnLine

	If Not objRS.EOF Then

		strTemp = vbNewLine & strLine & vbNewLine & _
			"Top " & varNTop & " - UpTime Control" & vbNewLine & _
			strLine & vbNewLine & _
			strTemp & curUser.GetCoreMsgStr("UPRanking") & Tab1 & _
				  curUser.GetCoreMsgStr("UPTopTotal") & Tab2 & _
				  curUser.GetCoreMsgStr("UPTopTRecord") & Tab2 & _
				  curUser.GetCoreMsgStr("UPTopConnections") & Tab1 & _
				  "OnLine" & Tab1 & _
				  curUser.GetCoreMsgStr("UPUser") & vbNewLine & _
			strLine & vbNewLine
				  
		Do Until objRS.EOF

			lngRanking = i + 1

			strTotalUpTime = UpTime(objRS.Collect(1), curUser.sName, 1)

			'Convert UpTime Record..
			strUpTimeRecord = UpTime(objRS.Collect(4), curUser.sName, 1)

			'User online?
			If Cbool(objRS.Collect(9)) Then _
				strOnLine = curUser.GetCoreMsgStr("UPOnLineYes") _
			Else	strOnLine = curUser.GetCoreMsgStr("UPOnLineNo")

			strTempLine = Cstr(lngRanking & Tab1 & strTotalUpTime & Tab1 & _
					   strUpTimeRecord & Tab1 & objRS.Collect(6) & Tab2 & _
					   strOnLine & Tab1 & objRS.Collect(0))
			
			'If don't exist number of the Top request..
			If IsNumeric(varNTop) Then _
				If varNTop = i Then  _
					Exit Do 
			'Generate line
			strTemp = strTemp & strTempLine & vbNewLine

			'Set ranking number..
			i = i + 1 

			'Move next record..
			objRS.MoveNext

			'Refresh process..
			DoEvents
		Loop

	Else
		MsgBox "Erro db temp is NULL records - GoTo line 1042"
		'Reply curUser, curUser.GetCoreMsgStr("UPNoData"), blnMC
		Set objRS = Nothing
		Exit Sub

	End If

	If IsNumeric(varNTop) Then
		If varNTop > i Then _
			strTemp = strTemp & strLine & vbNewLine & _
				Replace(curUser.GetCoreMsgStr("UPTopNotComplete"), "%[records]", (i)) & _
					vbNewLine
	End If

	Set objRS = Nothing

	If M_bShowTopMedia Then
		strQuery = "SELECT Avg(UpTimeTotal), Avg(UpTimeRecord), Avg(TotalConnections)" & _
			   "FROM UpTimeTmp " & _
			   "WHERE UpTimeTotal > 60 "

		Set objRS = objJetCon.Execute(strQuery)	
	
		srtMedia(0) = UpTime(objRS.Collect(0), curUser.sName, 1)
		srtMedia(1) = UpTime(objRS.Collect(1), curUser.sName, 1)
		srtMedia(2) = FormatNumber(objRS.Collect(2), 0)

		strTemp = vbNewLine & strTemp & strLine & vbNewLine & _
			  "Media - UpTime Control" & vbNewLine & strLine & vbNewLine & Tab1 & _
			  curUser.GetCoreMsgStr("UPTopTotal") & Tab3 & _
			  curUser.GetCoreMsgStr("UPTopTRecord") & Tab3 & _
			  curUser.GetCoreMsgStr("UPTopConnections") & Tab3 & vbNewLine & _
			  Tab1 & srtMedia(0) & Tab2 & srtMedia(1) & Tab2 & srtMedia(2) & vbNewLine

		Set objRS = Nothing
		DoEvents
	End If

	If M_bShowTopInfo Then
		strTemp = strTemp & strLine & vbNewLine & _
			"*Info: " & _
			" Y: " & curUser.GetCoreMsgStr("UPYears") & _
			" - M: " & curUser.GetCoreMsgStr("UPMonths") & _
			" - W: " & curUser.GetCoreMsgStr("UPWeeks") & _
			" - D: " & curUser.GetCoreMsgStr("UPDays") & vbNewLine & _
			"*Up Time Control " & M_sScriptVerison & " created by fLaSh" & vbNewLine

		DoEvents
	End If

	If blnHide Then
		'Send in private
		Reply curUser, strTemp & strLine & vbNewLine, blnMC
	Else
		'Send Chat To All users
		colUsers.SendChatToAll Settings.BotName, strTemp & strLine
	End If

	Set objRS = Nothing

	DoEvents

	'Delete temp table data
	'*****************************************************
	Set objRS = objJetCon.Execute("DELETE FROM UpTimeTmp")
	Set objRS = Nothing
#If DEBUG_MODE Then
	Call DebugPrint("DELETE FROM UpTimeTmp (TopsUpTime)")
#End If
	'*****************************************************

End Sub

Function RZero(strValue, bytSize)
	'Format string, ex: 1 to 01
	If Len(strValue) <= bytSize Then _
		RZero = String(bytSize - Len(strValue), "0") & strValue _
	Else    RZero = strValue
End Function

Function CalcDateToSec(strInDate, strOutDate)
	'Calculat Dates To Secunds
	CalcDateToSec = DateDiff("s", CDate(strInDate), CDate(strOutDate))
End Function

Function SelectSQL(strQuery, strField)
	Dim objRS
	Set objRS = objJetCon.Execute(strQuery)
	If Not objRS.EOF Then
		On Error Resume Next
		SelectSQL = objRS.Fields(strField)
	End If
	Set objRS = Nothing
	DoEvents
#If DEBUG_MODE Then
	Call DebugPrint("SelectSQL(" & strQuery & ", " & strField & ")")
#End If
End Function

Sub InsertSQL(strQuery) 
	Dim objRS
	Set objRS = objJetCon.Execute(strQuery)
	Set objRS = Nothing
	DoEvents
#If DEBUG_MODE Then
	Call DebugPrint("InsertSQL(" & strQuery & ")")
#End If
End Sub

Sub UpDateSQL(strQuery) 
	Dim objRS
	Set objRS = objJetCon.Execute(strQuery)
	Set objRS = Nothing
	DoEvents
#If DEBUG_MODE Then
	Call DebugPrint("UpDateSQL(" & strQuery & ")")
#End If
End Sub

Sub DeleteSQL(strQuery) 
	Dim objRS
	Set objRS = objJetCon.Execute(strQuery)
	Set objRS = Nothing
	DoEvents
#If DEBUG_MODE Then
	Call DebugPrint("DeleteSQL(" & strQuery & ")")
#End If
End Sub

Function GetAfectedRowsSQL(strQuery)
	Dim objRS
	Dim i
	Set objRS = objJetCon.Execute(strQuery)
	If Not objRS.EOF Then
		Do Until objRS.EOF
			i = i + 1
			DoEvents
			objRS.MoveNext
		Loop
	End If
	GetAfectedRowsSQL = i
	Set objRS = Nothing
#If DEBUG_MODE Then
	Call DebugPrint("GetAfectedRowsSQL(" & strQuery & ") (Afected Rows: " & i  & ")")
#End If
End Function

Sub OpenJetCon(objConnection)
	'Connection for database
	Set objConnection = NewConnection()
	objConnection.ConnectionTimeout = 10
	objConnection.ConnectionString = _
		"Provider=Microsoft.Jet.OLEDB.4.0;" & _
		"Data Source=.\Scripts\UpTimeControl-Lib\UpTimeControl.mdb"
	objConnection.Open
End Sub

Sub Reply(curUser, sMsg, blnMC)
	'Send chat message 
	Select Case blnMC
		Case True: curUser.SendChat Settings.BotName, CStr(sMsg)
		Case False: curUser.SendPrivate Settings.BotName, CStr(sMsg)
	End Select
	DoEvents
#If DEBUG_MODE Then
	Call DebugPrint("Message Replyed for User name: " & curUser.sName)
	DoEvents
#End If
End Sub

Function TestStringOk(strData)
	If Replace(strData, " ", "") <> "" Then _
		TestStringOk = True _
	Else	TestStringOk = False
End Function

Sub DoEvents()
	frmHub.DoEventsForMe
End Sub

#If DEBUG_MODE Then
Sub DebugPrint(sMsg)
	If M_bDebugToUser Then
		'Sends a main chat debug messages
		If colUsers.Online(CStr(M_sDebugUser)) Then 
			colUsers.ItemByName(CStr(M_sDebugUser)).SendPrivate Settings.HubName, _
				" UpTime Control @ Debug Mode" & vbNewLine & sMsg & vbNewLine
			DoEvents
		End If
	End If
	If M_bDebugToLog Then
		FileAccess.AppendFile FileAccess.AppPath & _
			"\Scripts\UpTimeControl-Lib\DebugMode.log", "[" & Now & "] " & sMsg
		DoEvents
	End If
End Sub
#End If

'Errors debug messages 
Sub Error(Line)

	Dim strErr
	strErr = "Private Message: Script Error" & vbNewLine & _
		 "Script: UpTime Control" & vbNewLine & _
		 "When: " & Now & vbNewLine & _
		 "Line: " & Line & vbNewLine & _
		 "Error Number: " & Err.Number & vbNewLine & _
		 "Error Description: " & Err.Description & vbNewLine
		 
#If DEBUG_MODE Then
	Call DebugPrint(strErr)
#End If

	If TestStringOk(M_sDebugUser) Then 
		MsgBox strErr, , Settings.HubName & " - UpTime Control"
	Else
		If colUsers.Online(CStr(M_sDebugUser)) Then 
			colUsers.ItemByName(CStr(M_sDebugUser)).SendPrivate Settings.HubName, strErr
			DoEvents
		End If
	End If

	DoEvents

	FileAccess.AppendFile FileAccess.AppPath & "\Scripts\UpTimeControl-Lib\UpTimeControl.log", _
		Now & "|" & Err.Number & "|" & Err.Description & "|" & Line & "|"

End Sub  