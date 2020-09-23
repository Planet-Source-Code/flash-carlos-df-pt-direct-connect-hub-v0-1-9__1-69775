'----------------------------------------------------------------------------------
'                                         _.oo.   
'                 _.u[[/;:,.         .odMMMMMM' 
'              .o888UU[[[/;:-.  .o@P^    MMM^   
'             oN88888UU[[[/;::-.        dP^   
'            dNMMNN888UU[[[/;:--.   .o@P^
'           ,MMMMMMN888UU[[/;::-. o@^    		   Zatarc Bot
'           NNMMMNN888UU[[[/~.o@P^          		By   
'           888888888UU[[[/o@^-..    		GhOstFaCE @ Asgard  
'          oI8888UU[[[/o@P^:--..   
'       .@^  YUU[[[/o@^;::---..                      Modified
'     oMP     ^/o@P^;:::---..                           By
'  .dMMM    .o@^ ^;::---...                         ShadowKidX aka AchaicLight
' dMMMMMMM@^`       `^^^^                               and
'YMMMUP^                                            TheLastOne
' ^^
'ArcLight
'	Designed for the ArcLight System
'0.975
'	Collective update from previous ArcLight 0.965
'	New commands brought over from ArcLight 0.965:
'		ipinfo (mainly useful if not register users, or are on a LAN with Wireless)
'		ipinfo included in userinfo
'		flood (Warning: sends 1000 PMs and crashes the recipients client
'		listtempban
'		listpermban
'		sayme (impersonate another user with /me already included)
'		setjoinmsg (change the motd without using the gui)
'	Readded the removed commands from 0.962
'	Added boolean for minChatClass False=disabled
'	Added precompile for Prefix usable in registration
'	Removed Customizaton features (bot pm forwarding, customcom help organization)
'	
'0.981
'	Version included in DDCH 0.3.39
'0.971
'	Added Unload script commands(TUFF)
'	Bugs fixes/typos/languages(TUFF)
'	Begin adding languages support(JDommi/TheNOP/TUFF)
'	Added #Include example.(TheNOP)
'	Added precompile "#Const/#If/#End If" example.(TheNOP)
'	 (disable, prevent the code for the "air" command from being loaded in memory.)
'	 (warning, know "bug" with precompile conditions, http://www.shadowdc.com/forums/index.php?topic=765.0)

'0.965 ArcLight
'	Mute messages re-enabled.
'	MuteList is now cleared every ~2 Hours
'	Locker Script interferes with main PreData, will merge in future
'
'0.964 ArcLight
'	Fixed setjoinmsg so that it properly updates the GUI.
'	Fixed default kick display.
'	Mystery Code was in fact the pattern matching for kicking.
'
'0.962 ArcLight
'	 Mostly cleaning. Removed several obsolete functions and commands. Reorganized right click menu. Renamed command 'bc' to 'mass'.
'	 Mute, flood, userinfo, and ban listing commands fixed.  Added ArcLightLocker support.  Added iplist command.   
'	 Removed commands: ops, botsay, botsayme, setmaxshare, setmaxhubs, lockchat
'	 Removed Functions: CheckIP, CountIP
'	 Discovered that tabbing is all screwed up because of the differences between the DDCH internal script editor and notepad.
'	 Monitoring mystery code
'
'ArcLight							  
'        Designed for the ArcLight System
'        Includes new commands such as ipinfo, the addition of ipinfo to userinfo, flood, listtempban, listpermban,
'        sayme, and setjoinmsg
'
'0.958
'	Added PreData screening of PMs to allow hubowner to specify minimum class that can PM other users. OPs can always be PMed
'	Added simple boolean to turn rightclickcommands on/off. Off state would save a little bw
'	Divided "ops" and "users" into two different commands to avoid Ops looking for the ops list to also get the list of users. If you have a LOT of registered users concider setting the command "users" to be SOP or Admin only
'	Fixed the Report command
'0.955
'	Added ops/users command. Regular users will get a list of operators. Ops a list of regs/vips/ops and Admins a complete list of users plus their passwords
'	Due to popular request i added a !motd command to show users the login message again without reconnecting
'	Fixed a typo in the unload commands. UnbanUser should now be totally gone. The scripted Unban command will work with unban user aswell (as always)
'	Added command to clean the userdatabase for accounts older than 90 days. Remember to inform your Ops and Regs about this
'	Added offline userinfo. Each successfull login will get his/her own txt file which the Userinfo command will bring up if the user isnt online (thx Loki)
'
'ToDo: Complet the languages support
'      Revise CustComArrival
'      Revise Commands adding code. See: http://www.shadowdc.com/forums/index.php?topic=747.0
'	(An #Include can be added after Option Explicit to outsource the variables dimming to an other file)
'      Revise IsBan
'
'----------------------------------------------------------------------------------
Option Explicit

Private MuteList
'-----------------------------------------
' Script Settings
'-----------------------------------------
Const sAdmin = "fLaSh"			'User who recieves the debug messages
Const sPrefix = ""				'Prefix required for users to self reg
Const sOwner = "fLaSh"			'HubOwner shown in Hubinfo
Const sForum = "" 	'Forum adress shown in hubinfo
Const bMinChatEnabled=False
Const iMinclass = 1				'Classes above or equal this can send PMs to non OPs (set to 1 for all)
Const sHomePage = "http://HublistChecker.pt.vu/"	'Hubs homepage adress shown in hubinfo
Const sScriptv = "0.975 Standard ArcLight"
Const bOfflineInfo = False			'Change to false if you dont want to keep track of userinfo after users leave the hub
Const bAddRightclick = True			'Change to false if you dont want the righclick commands to be sent to users and ops (saves a little bw)
#Const AIRENA = False				'Enable the air command
#Const PREENA = False				'Enable the Prefix for self-registration command
'-----------------------------------------
'Do not change anything below unless you know what your doing
'-----------------------------------------

Dim CmdPrefix, sTopic, bFreezeMC, bMute, oIPs, iTime, iCounter, aAdmins


	aAdmins = Array("fLaSh") '<<<<<<<<<Change these to your admin user names, add asmany as you want


Sub Main()

	sTopic="Topic"
	'Load language(s) messages strings
	AddScriptMessages ".\Scripts\Zatarc-Lib\Zatarc.lng"

	frmHub.RegisterBotName Settings.OpChatName, True, 0, "OPs Only Chat<OP Chat>", "BOT", ,8
	frmHub.RegisterBotName Settings.BotName, True, 0, "Security<Security>", "BOT", ,8

	bMute=False
	bFreezeMC = False
	CmdPrefix = Chr(Settings.CPrefix)

	Set MuteList=NewDictionary()
	
	Set oIPs  = CreateObject("Scripting.Dictionary")
	
	iTime = 1
	iCounter = 0
	tmrScriptTimer.interval=60000 
	tmrScriptTimer.enabled=True

'Commands ID below 51 are reserved and can't be use.
'-----------------------------------------
' Load the Commands
'-----------------------------------------
	If Not colCommands.Exists("about") Then colCommands.Add 101, "about", "ZatAboutDesc", 1, True
	If Not colCommands.Exists("me") Then colCommands.Add 102, "me", "ZatMeDesc", 3, True
	If Not colCommands.Exists("help") Then colCommands.Add 103, "help", "ZatHelpDesc", 1, True
	If Not colCommands.Exists("report") Then colCommands.Add 104, "report", "ZatReportDesc", 1, True
#If AIRENA Then
	If Not colCommands.Exists("air") Then colCommands.Add 105, "air", "ZatAirDesc", 1, True
#End If
	If Not colCommands.Exists("ops") Then colCommands.Add 106, "ops", "ZatOpsDesc", 1, True
	If Not colCommands.Exists("motd") Then colCommands.Add 107, "motd", "ZatMotdDesc", 1, True
	If Not colCommands.Exists("network") Then colCommands.Add 201, "network", "ZatNetDesc", 1, True
	If Not colCommands.Exists("rules") Then colCommands.Add 202, "rules", "ZatRulesDesc", 1, True
	If Not colCommands.Exists("regme") Then colCommands.Add 301, "regme", "ZatRegmeDesc", 1, True
	If Not colCommands.Exists("userinfo") Then colCommands.Add 404, "userinfo", "ZatUIDesc", 6, True
	If Not colCommands.Exists("myinfo") Then colCommands.Add 403, "myinfo", "ZatMyInfoDesc", 1, True
	If Not colCommands.Exists("hubinfo") Then colCommands.Add 402, "hubinfo", "ZatHubInfoDesc", 1, True
	If Not colCommands.Exists("myip") Then colCommands.Add 401, "myip", "ZatMyIpDesc", 1, True

	If Not colCommands.Exists("addreg") Then colCommands.Add 503, "addreg", "ZatAddRegDesc", 6, True
	If Not colCommands.Exists("delreg") Then colCommands.Add 502, "delreg", "ZatDelRegDesc", 6, True
	If Not colCommands.Exists("setlanguage") Then colCommands.Add 504, "setlanguage", "ZatLanguageDesc", 3, True

	If Not colCommands.Exists("drop") Then colCommands.Add 607, "drop", "ZatDropDesc", 6, True
	If Not colCommands.Exists("kick") Then colCommands.Add 606, "kick", "ZatKickDesc", 6, True
	If Not colCommands.Exists("ban") Then colCommands.Add 605, "ban", "ZatBanDesc", 8, True
	If Not colCommands.Exists("bannick") Then colCommands.Add 604, "bannick", "ZatBanNickDesc", 8, True
	If Not colCommands.Exists("banip") Then colCommands.Add 603, "banip", "ZatBanIPDesc", 8, True
	If Not colCommands.Exists("tban") Then colCommands.Add 602, "tban", "ZatTBanDesc", 6, True
	If Not colCommands.Exists("tbanip") Then colCommands.Add 601, "tbanip", "ZatTBanIPDesc", 6, True
	If Not colCommands.Exists("unban") Then colCommands.Add 600, "unban", "ZatUnbanDesc", 8, True

	If Not colCommands.Exists("cleartemp") Then colCommands.Add 703, "cleartemp", "ZatClearTempDesc", 10, True
	If Not colCommands.Exists("clearipbans") Then colCommands.Add 702, "clearipbans", "ZatClearIPDesc", 10, True

	If Not colCommands.Exists("listtempban") Then colCommands.Add 705, "listtempban", "Lists the Temp Bans.", 8, True
	If Not colCommands.Exists("listpermban") Then colCommands.Add 704, "listpermban", "Lists the Perm Bans.", 10, True
	If Not colCommands.Exists("flood") Then colCommands.Add 706, "flood", "Floods a user with 1000 PMs from randomly generated names.", 10, True
	If Not colCommands.Exists("ipinfo") Then colCommands.Add 707, "ipinfo", "IP info", 6, True
	If Not colCommands.Exists("iplist") Then colCommands.Add 709, "iplist", "List of all the users and IPs connected to the hub.", 6, True 

	If Not colCommands.Exists("mute") Then colCommands.Add 803, "mute", "ZatMuteDesc", 6, True
	If Not colCommands.Exists("unmute") Then colCommands.Add 802, "unmute", "ZatUnMuteDesc", 6, True
	
	If Not colCommands.Exists("topic") Then colCommands.Add 901, "topic", "ZatTopicDesc", 8, True
	If Not colCommands.Exists("say") Then colCommands.Add 902, "say", "ZatSayDesc", 8, True
	If Not colCommands.Exists("sayme") Then colCommands.Add 903, "sayme", "Impersonates the bot and uses the /me command", 8, True
	If Not colCommands.Exists("mass") Then colCommands.Add 904, "mass", "ZatBCDesc", 8, True
	If Not colCommands.Exists("userlimit") Then colCommands.Add 905, "userlimit", "ZatUserlimitDesc", 10, True
	If Not colCommands.Exists("setminshare") Then colCommands.Add 906, "setminshare", "ZatSetMinShareDesc", 10, True
	If Not colCommands.Exists("setredirect") Then colCommands.Add 908, "setredirect", "ZatSetRedirectDesc", 10, True
	If Not colCommands.Exists("setminslots") Then colCommands.Add 910, "setminslots", "ZatSetMinSlotsDesc", 10, True
	If Not colCommands.Exists("setmaxslots") Then colCommands.Add 911, "setmaxslots", "ZatSetMaxSlotsDesc", 10, True
	If Not colCommands.Exists("setslotratio") Then colCommands.Add 912, "setslotratio", "ZatSetSlotRatioDesc", 10, True
	If Not colCommands.Exists("setjoinmsg") Then colCommands.Add 710, "setjoinmsg", "Sets the join message to the given text", 10, True
	If Not colCommands.Exists("resetports") Then colCommands.Add 913, "resetports", "ZatResetPortsDesc", 10, True
	If Not colCommands.Exists("lockchat") Then colCommands.Add 914, "lockchat", "ZatLockChatDesc", 6, True
	If Not colCommands.Exists("users") Then colCommands.Add 915, "users", "ZatUsersDesc", 6, True
	If Not colCommands.Exists("cleandb") Then colCommands.Add 916, "cleandb", "ZatCleanDBDesc", 10, True
	'If Not colCommands.Exists("shedule") Then colCommands.Add 917, "shedule", "ZatSheduleDesc", 10, True
	If Not colCommands.Exists("shedule") Then colCommands.Add 917, "shedule", "DNS Update Schedule", 10, True

End Sub

Function PreDataArrival(curUser, sData)
	
	Dim sReason
	Dim sTarget
	Dim objUser
	Dim i
	Dim aData

	aData = Split(MidB(sData, 3), " ", 3)

	Select Case AscW(sData)
		Case 36
			Select Case CStr(aData(0))
				Case "To:"
				
					If bMute Then
						'only check users class if mute is enabled. save one test most of the time...
						If curUser.Class < 5 Then
							'PMs are disabled for them
							PreDataArrival = Empty
							Exit Function
						End If
					End If
				
					If bMinChatEnabled=True Then
						If curUser.Class >= iMinclass Then
							 ' User can send his PM
							PreDataArrival = sData
							Exit Function
						Else
							If colUsers.Online(CStr(aData(1))) Then
								If colUsers.ItemByName(CStr(aData(1))).Class > 5 Then
									 ' User can send PM to OPs only
									PreDataArrival = sData
									Exit Function
								Else
									 ' User cant send PM
									PreDataArrival = Empty
									Exit Function
								End If
							End If
						End If
					End If
				Case "Kick"
					'disable protocol $Kick so it don't try to double kick
					'PreDataArrival=Empty
					PreDataArrival=sData
					Exit Function
				Case Else
					'Relay to other script(s) and possibly to hub
					'PreDataArrival=sData
			End Select

		Case 60
			If bMute Then
				'only check users class if Main Chat freeze is enable. save one test most of the time...
				If curUser.Class < 5 Then
					'Main Chat is disabled for them
					PreDataArrival = Empty
					Exit Function
				End If
			End If
			If bFreezeMC Then
				'only check users class if Main Chat freeze is enable. save one test most of the time...
				If curUser.Class < 5 Then
					'Main Chat is disabled for them
					PreDataArrival = Empty
					Exit Function
				End If
			End If

			If curUser.bOperator Then
				'only if "is kicking" pattern is found. use hub's function  ;D
				If RegExps.TestStr(sData,".*is kicking[ ]\S*[ ]because:[ ]") Then
					'RegExp matches capturing could be use to check if "is kicking" pattern is found and get sTarget at the same time.
					'me think this should do-> If LenB(sTarget = RegExps.CaptureSubStr(sData,".*is kicking[ ](\S*)[ ]because:[ ]")) Then
					sTarget = BetweenFirst(sData, "is kicking ", " ")

					If colUsers.Online(CStr(sTarget)) Then 	
						Set objUser = colUsers.ItemByName(CStr(sTarget))
						sReason = AfterFirst(sData, "because: ")
						i = IsBan(sReason)

						If objUser.Class < curUser.Class Then
							Select Case CLng(i)
								Case -1
									'frmHub.DoEventsForMe
									colIPBans.Add objUser.IP
									'curUser.SendChat Settings.BotName, "The user, "& sTarget  & ", was kicked. IP:"& objUser.IP &" is banned permanently"
									'make all Ops aware of the kick., "WasKickedPerm" added to Zatarc lang file.
									colUsers.SendChatToAll Settings.BotName, TagReplace(curUser.GetCoreMsgStr("ZatWasKickedPerm"), curUser, objUser)
									objUser.Disconnect
								Case 0
									'frmHub.DoEventsForMe
									colIPBans.Add objUser.IP,Settings.DefaultBanTime
									'curUser.SendChat Settings.BotName, "The user, "& sTarget & ", was kicked. IP:"& objUser.IP
									'make all Ops aware of the kick., "KickedBy" is default hub msg.
									colUsers.SendChatToAll Settings.BotName, TagReplace(curUser.GetCoreMsgStr("KickedBy"), curUser, objUser)
									SendToAdmin CStr(CStr(objUser.sName)&" was kicked by "&curUser.sName&" because: "&CStr(sReason))
									objUser.Disconnect
								Case Else
									'frmHub.DoEventsForMe
									colIPBans.Add objUser.IP,CLng(i)
									'curUser.SendChat Settings.BotName, "The user, "& sTarget & ", was kicked. IP:"& objUser.IP &" is banned for "& CLng(i) &" minutes"
									'make all Ops aware of the kick.., "WasKickedTmp" added to Zatarc lang file.
									'replace %[tban] separatly, not use often...
									colUsers.SendChatToAll Settings.BotName, Replace(TagReplace(curUser.GetCoreMsgStr("ZatWasKickedTmp"), curUser, objUser), "%[tban]", i)
									objUser.Disconnect
							End Select
						End If
					End If
				End If
			End If
		Case Else
			'Relay to other script(s) and possibly to hub
			'in case script(s) use an other protocol or new protocol is added to hub
			'PreDataArrival = sData
	End Select
	'If we have not exit it mean that all seem to be ok, relay sData to the hub and/or other script(s)
	PreDataArrival = sData
End Function

Function TagReplace(strString, curUser, objUser)
'------------------------------------------------------------------
'Purpose:   Replace %[var] with proper data
'
'Params:    curUser		Current user object
'	objUser		A user object in colUsers
'
'Return	Modified string or same input string if no var(s) were replaced
'------------------------------------------------------------------

	If InStrB(strString, "%[user]") Then strString = Replace(strString, "%[user]", objUser.sName)
	If InStrB(strString, "%[op]") Then strString = Replace(strString, "%[op]", curUser.sName)
	If InStrB(strString, "%[ip]") Then strString = Replace(strString, "%[ip]", objUser.IP)
	TagReplace = strString

End Function

'------------------------------------------------------------------
Function sFDate(sDate)
'------------------------------------------------------------------

	sFDate=WeekDayName(Weekday(sDate), True)+" "+CStr(Day(sDate))+"."+CStr(Month(sDate))+"."+CStr(Year(sDate))+" "+FormatDateTime(sDate,4)+":"
	If Second(sDate)<10 Then sFDate=sFDate+"0"+CStr(Second(sDate)) Else sFDate=sFDate+CStr(Second(sDate))

End Function

'------------------------------------------------------------------
Sub SendToAdmin(sData)
'------------------------------------------------------------------
	
	Dim ii

	For ii=0 To UBound(aAdmins)
		If colUsers.Online(CStr(aAdmins(ii))) Then
			colUsers.ItemByName(CStr(aAdmins(ii))).SendPrivate Settings.BotName,CStr(sData)
		End If
	Next

End Sub

'-----------------------------------------Start of Commands-----------------------------------------
Sub CustComArrival(curUser, objCommand, sData, blnMC)
'suggestion:
' all commands that need dimmed variables should be moved to their own subs.
'reason:
'to speedup things a little...,more commands might be added later
' so the list of variables to dim at execution of any command will grow...

	Dim sSplit
	Dim targetClass
	Dim sSplot
	Dim sMsg
	Dim r
	Dim user
	Dim objUser
	Dim s
	Dim objCom
	Dim sUnit
	Dim sTempIPList,sPermIPList
	Dim sReplyInfo

	'help cmd
	Dim sNormal
	Dim sReg
	Dim sVip
	Dim sOp
	Dim sSop
	Dim sCAdmin
	Dim sTmpDesc

#If AIRENA Then
	'air cmd
	Dim GetURL
	Dim http
	Dim wwwTxt
#End If

	Select Case objCommand.Name
		'-----------------------------------------
		Case "help"
		'----------------------------------------- 

			sNormal = vbNewLine&"User Commands: "
			sReg = vbNewLine&"Registered Users: "
			sVip = vbNewLine&"VIP Commands: "
			sOp = vbNewLine&"Operator Commands: "
			sSop = vbNewLine&"Super Operator Commands: "
			sCAdmin = vbNewLine&"Admins Commands: "

			For Each objCom In colCommands
				If objCom.Class =< curUser.Class Then
					sTmpDesc = curUser.GetCoreMsgStr(objCom.Description) 

					If sTmpDesc = "" Then  sTmpDesc = objCom.Description 


						If objCom.ID=504 Then
                                                sReg = sReg & vbNewLine & VbTab & CmdPrefix & objCom.Name & VbTab&VbTab & sTmpDesc
						ElseIf objCom.ID=704 Or objCom.ID=705 Then
                                                sCAdmin = sCAdmin & vbNewLine & VbTab & CmdPrefix & objCom.Name &VbTab& VbTab & sTmpDesc
						ElseIf objCom.ID=908 Then
                                                sSop = sSop & vbNewLine & VbTab & CmdPrefix & objCom.Name & VbTab&VbTab & sTmpDesc
						ElseIf objCom.ID=709 Then
                                                sSop = sSop & vbNewLine & VbTab & CmdPrefix & objCom.Name & VbTab&VbTab&VbTab & sTmpDesc
						ElseIf objCom.ID=702 Or objCom.ID=703 Or (objCom.ID=>906 And objCom.ID<=913) Or objCom.ID=710 Then
                                                sCAdmin = sCAdmin & vbNewLine & VbTab & CmdPrefix & objCom.Name & VbTab&VbTab & sTmpDesc
						ElseIf objCom.ID=301 Or objCom.ID=201 Then

						Else

					Select Case objCom.Class
						Case 1,2
							'sNormal = sNormal&vbNewLine&vbTab&(CmdPrefix)&objCom.Name&vbTab&vbTab&objCom.Description
							sNormal = sNormal & vbNewLine & VbTab & CmdPrefix & objCom.Name & VbTab & VbTab & VbTab & sTmpDesc
						Case 3,4
							sReg = sReg & vbNewLine & VbTab & CmdPrefix & objCom.Name & VbTab & VbTab & VbTab & sTmpDesc
						Case 5
							sVip = sVip & vbNewLine & VbTab & CmdPrefix & objCom.Name & VbTab & VbTab & VbTab & sTmpDesc
						Case 6,7
							sOp = sOp & vbNewLine & VbTab & CmdPrefix & objCom.Name & VbTab & VbTab & VbTab & sTmpDesc
						Case 8,9
							sSop = sSop & vbNewLine & VbTab & CmdPrefix & objCom.Name & VbTab & VbTab & VbTab & sTmpDesc
						Case 10,11
							sCAdmin = sCAdmin & vbNewLine & VbTab & CmdPrefix & objCom.Name & VbTab & VbTab & VbTab & sTmpDesc
					End Select
					End If
				End If
			Next
				Select Case curUser.Class
					Case 1,2
						Reply curUser, sNormal,blnMC
					Case 3,4
						Reply curUser, sNormal&sReg,blnMC
					Case 5
						Reply curUser, sNormal&sReg&sVIP,blnMC
					Case 6,7
						Reply curUser, sNormal&sReg&sVIP&sOp,blnMC
					Case 8,9
						Reply curUser, sNormal&sReg&sVIP&sOp&sSop,blnMC
					Case 10,11
						Reply curUser, sNormal&sReg&sVIP&sOp&sSop&sCAdmin,blnMC
				End Select
		'-----------------------------------------
		Case "motd"
		'-----------------------------------------
			Reply curUser, Settings.JoinMsg, blnMC
		'-----------------------------------------
		Case "network"
		'-----------------------------------------
			Reply curUser, FileAccess.ReadFile(".\Scripts\Zatarc-Lib\network.txt"),blnMC

		'-----------------------------------------
		Case "about"
		'-----------------------------------------
			Reply curUser," "&vbNewLine&vbNewLine _
				&VbTab&"This hub is running"&vbNewLine _
				&VbTab&"Script:"&VbTab&VbTab&"ArcLight-Modified Zatarc-MultiBot"&vbNewLine _
				&VbTab&"By:"&VbTab&VbTab&"GhOstFaCE"&vbNewLine _
				&VbTab&"Modified By:"&VbTab&"ShadowKidX & TheLastOne"&vbNewLine _
				&VbTab&"Current Version By:"&VbTab&"ArchaicLight (aka ShadowKidX)"&vbNewLine _
				&VbTab&"Version:"&VbTab&VbTab&(sScriptv)&vbNewLine _
				&VbTab&"http://thought.queensdc.ca/ddch/"&vbNewLine,blnMC
			
		'-----------------------------------------
		Case "rules"
		'-----------------------------------------
			Reply curUser, FileAccess.ReadFile(".\Scripts\Zatarc-Lib\rules.txt"),blnMC
		'-----------------------------------------
		Case "regme"
		'-----------------------------------------
			sSplot = AfterFirst(sData," ")
	#If PREENA Then
			If LCase(Left(CStr(curUser.sName), Len(sPrefix))) = LCase(sPrefix) Then
	#End If
				If LenB(sSplot) Then
					Select Case colRegistered.Add(curUser.sName, CStr(sSplot), 3, "Self-registration")
						Case 0: Reply curUser, "You have been successfully registered with the following password : " & sSplot,blnMC
						Case 1: Reply curUser, "You are already registered.",blnMC
						'is default hub msg.
						Case 3: Reply curUser, curUser.GetCoreMsgStr("PassLength"),blnMC
					End Select
				Else
					Reply curUser, "Proper syntax is "&(CmdPrefix)&"regme <password>",blnMC
				End If
	#If PREENA Then
			Else
				Reply curUser, "Registered users are required to use the prefix [SGF]. Add [SGF] infront of your username and try again", blnMC
			End If
	#End If
	#If AIRENA Then
		'-----------------------------------------
		Case "air"
		'-----------------------------------------
			GetURL = "http://curzed.asgards.org/air.txt"
			Set Http = CreateObject("Microsoft.XMLhttp")
			Http.Open "GET",GetURL,False
			Http.Send
			wwwTxt = Http.ResponseText
			Set Http = Nothing
			wwwTxt = Replace(wwwTxt , Chr(10), VbCrLf)
			Reply curUser, CStr(wwwTxt),blnMC
	#End If

		'-----------------------------------------
		Case "myip" 
		'----------------------------------------- 
			Reply curUser, "Your IP: "&curUser.IP,blnMC

		'-----------------------------------------
		Case "me"
		'----------------------------------------- 
			sSplot = AfterFirst(sData," ")
			If LenB(sSplot) Then
				colUsers.SendToAll "* "&curUser.sName&" "&CStr(sSplot)&"|"
			Else
				Reply curUser, "Proper syntax is "&(CmdPrefix)&"me <chatdata>",blnMC
			End If
		'----------------------------------------- 
		Case "mute"
		'----------------------------------------- 
			sSplot = AfterFirst(sData," ")
			If LenB(sSplot) Then
				If CStr(sSplot) = "*" Then
					For Each User In colUsers
						If user.Class<5 Then User.mute=True
					Next
					bMute=True
					colUsers.SendChatToAll Settings.BotName,CStr("All users muted by "&curUser.sName&".")
					colUsers.SendPrivateToOps Settings.BotName,CStr("All users muted by "&curUser.sName&".")
					frmHub.DoEventsForMe	
				ElseIf colUsers.Online(CStr(sSplot)) Then
					If colUsers.ItemByName(CStr(sSplot)).Class >= curUser.Class Then
						Reply curUser, "You can only mute users lower than your class",blnMC :Exit Sub
						frmHub.DoEventsForMe
					End If

					If Not MuteList.Exists(CStr(sSplot)) Then
						Reply curUser, "User Muted",blnMC
						colUsers.SendChatToAll Settings.BotName,CStr(CStr(sSplot)&" was muted by "&curUser.sName&".")
						colUsers.SendPrivateToOps Settings.BotName,CStr(CStr(sSplot)&" was muted by "&curUser.sName&".")
						colUsers.ItemByName(CStr(sSplot)).Mute = True
						MuteList.Add sSplot,curUser.sName
					Else
						Reply curUser, "User already muted",blnMC
					End If
				End If
			Else
				Reply curUser, "Proper syntax is "&(CmdPrefix)&"mute <user> or * if muting the whole hub",blnMC
			End If

		'----------------------------------------- 
		Case "unmute"
		'----------------------------------------- 
			sSplot = AfterFirst(sData," ")
			If LenB(sSplot) Then
				If CStr(sSplot) = "*" Then
					For Each User In colUsers
						'If user.Class<5 Then User.mute=False
						User.mute=False
					Next

					bMute = False
					colUsers.SendChatToAll Settings.BotName,CStr("All users unmuted by "&curUser.sName&".")
					colUsers.SendPrivateToOps Settings.BotName,CStr("All users unmuted by "&curUser.sName&".")
					frmHub.DoEventsForMe
				ElseIf colUsers.Online(CStr(sSplot)) Then
					If Not MuteList.Exists(CStr(sSplot)) Then
						Reply curUser, "User is not muted",blnMC
					Else
						Call MuteList.Remove(CStr(sSplot))
						colUsers.ItemByName(CStr(sSplot)).Mute = False
						Reply curUser, "User unmuted",blnMC
						colUsers.SendChatToAll Settings.BotName,CStr(CStr(sSplot)&" was unmuted by "&curUser.sName&".")
						colUsers.SendPrivateToOps Settings.BotName,CStr(CStr(sSplot)&" was unmuted by "&curUser.sName&".")
					End If
				End If
			Else
				Reply curUser, "Proper syntax is "&(CmdPrefix)&"unmute <user> or * if unmuting the whole hub",blnMC
			End If

		'----------------------------------------- 
		Case "drop"
		'-----------------------------------------

			sSplit=Split(sData," ",3)
			If UBound(sSplit) < 1 Then Reply curUser, "Proper syntax is "&(CmdPrefix)&"drop <user> <reason>",blnMC :Exit Sub
			If Not colUsers.Online(CStr(sSplit(1))) Then Reply curUser, "User is not online in this hub", blnMC :Exit Sub

			If UBound(sSplit) = 1 Then
				r = "(No Reason Given)"
			Else
				r = sSplit(2)
			End If

			targetClass = colRegistered.Registered(CStr(sSplit(1)))

			If targetClass >= curUser.Class Then 
				Reply curUser, "you cant drop a higher class than yourself", blnMC 
			Else
				Reply curUser, "The user "&CStr(sSplit(1))&" has been disconnected because: "&CStr(r), blnMC
				SendToAdmin CStr(CStr(sSplit(1))&" was disconnected by "&curUser.sName&" because: "&CStr(r))
				colUsers.ItemByName(CStr(sSplit(1))).SendPrivate Settings.BotName,"You have been disconnected because: "&CStr(r)
				frmHub.DoEventsForMe
				'Set objUser = colUsers.ItemByName(CStr(sSplit(1)))
				'objUser.Disconnect()
				colUsers.ItemByName(CStr(sSplit(1))).Disconnect()
				'frmHub.DoEventsForMe
			End If

		'----------------------------------------- 
		Case "kick"
		'----------------------------------------- 
			sSplit=Split(sData," ",3)
			If UBound(sSplit) < 1 Then Reply curUser, "Proper syntax is "&(CmdPrefix)&"kick <user> <reason>",blnMC :Exit Sub
			If Not colUsers.Online(CStr(sSplit(1))) Then Reply curUser, "User is not online in this hub", blnMC :Exit Sub

			If UBound(sSplit) = 1 Then
				r = "(No Reason Given)"
			Else
				r = sSplit(2)
			End If

			targetClass = colRegistered.Registered(CStr(sSplit(1)))

			If targetClass >= curUser.Class Then
				Reply curUser, "you cant kick a higher class than yourself", blnMC
			Else
				Reply curUser, "The user "&CStr(sSplit(1))&" has been kicked because: "&CStr(r), blnMC
				colUsers.ItemByName(CStr(sSplit(1))).SendPrivate Settings.BotName,"You have been kicked because: "&CStr(r)
				frmHub.DoEventsForMe
				colUsers.SendChatToAll curUser.sName,CStr(curUser.sName&" is kicking "&CStr(sSplit(1))&" because: "&CStr(r))
				colUsers.SendChatToAll Settings.BotName,"The user, "&CStr(sSplit(1))&", was kicked by "&curUser.sName&". IP: "&CStr(colUsers.ItemByName(CStr(sSplit(1))).IP)
				SendToAdmin CStr(CStr(sSplit(1))&" was kicked by "&curUser.sName&" because: "&CStr(r))
				frmHub.DoEventsForMe   
				colUsers.ItemByName(CStr(sSplit(1))).Kick
			End If

		'----------------------------------------- 
		Case "ban"
		'----------------------------------------- 
			sSplit=Split(sData," ",3)
			If UBound(sSplit) < 1 Then Reply curUser, "Proper syntax is "&(CmdPrefix)&"ban <user> <reason>",blnMC :Exit Sub
			If Not colUsers.Online(CStr(sSplit(1))) Then Reply curUser, "User is not online in this hub", blnMC :Exit Sub

			If UBound(sSplit) = 1 Then
				r = "(No Reason Given)"
			Else
				r = sSplit(2)
			End If

			targetClass = colRegistered.Registered(CStr(sSplit(1)))

			If targetClass >= curUser.Class Then
				Reply curUser, "you cant ban a higher class than yourself", blnMC
			Else
				Reply curUser, "The user "&CStr(sSplit(1))&" has been banned because: "&CStr(r), blnMC
				colUsers.SendPrivateToOps Settings.BotName,CStr("The user "&CStr(sSplit(1))&" has been ip and nick banned by "&curUser.sName&" because: "&CStr(r))
				colUsers.ItemByName(CStr(sSplit(1))).SendPrivate Settings.BotName,"You have been banned because: "&CStr(r)
				frmHub.DoEventsForMe
				colIPBans.Add CStr(colUsers.ItemByName(CStr(sSplit(1))).IP)
				Call colRegistered.Add(CStr(sSplit(1)), CStr( r), -1, CStr(curUser.sName))
				colUsers.ItemByName(CStr(sSplit(1))).Disconnect
			End If

		'----------------------------------------- 
		Case "bannick"
		'----------------------------------------- 
			sSplit=Split(sData," ",3)
			If UBound(sSplit) < 1 Then Reply curUser, "Proper syntax is "&(CmdPrefix)&"bannick <user> <reason>",blnMC :Exit Sub

			If UBound(sSplit) = 1 Then
				r = "(No Reason Given)"
			Else
				r = sSplit(2)
			End If

			Reply curUser, "The user "&CStr(sSplit(1))&" has been banned because: "&CStr(r), blnMC
			colUsers.SendPrivateToOps Settings.BotName,CStr("The user "&CStr(sSplit(1))&" has been nick banned by "&curUser.sName&" because: "&CStr(r))
			frmHub.DoEventsForMe
			Call colRegistered.Add(CStr(sSplit(1)), CStr( r), -1, CStr(curUser.sName))

		'----------------------------------------- 
		Case "banip"
		'----------------------------------------- 
			sSplot = AfterFirst(sData," ")
			If LenB(sSplot) Then
				frmHub.DoEventsForMe
				Call colIPBans.Add (CStr(sSplot))
				Reply curUser, "The IP "&CStr(sSplot)&" Has been banned", blnMC
				colUsers.SendPrivateToOps Settings.BotName,CStr("The user "&CStr(sSplot)&" has been ip banned by "&curUser.sName&" because: "&CStr(r))
			Else
				Reply curUser, "Proper syntax is "&(CmdPrefix)&"banip <IP>",blnMC
			End If

		'----------------------------------------- 
		Case "unban"
		'----------------------------------------- 
			sSplot = AfterFirst(sData," ")
			If LenB(sSplot) Then
				Select Case True
					Case CBool(colIPBans.Check(CStr(sSplot)))
						Call colIPBans.Remove(CStr(sSplot)) 
						Reply curUser, "The IP "&CStr(sSplot)&" has been unbanned",blnMC
						colUsers.SendPrivateToOps Settings.BotName,CStr("The IP "&CStr(sSplot)&" has been unbanned by "&curUser.sName&".")
					Case colRegistered.Registered(CStr(sSplot)) = -1
						colRegistered.Remove(CStr(sSplot))
						Reply curUser, "The username: "& sSplot &" has been unbanned", blnMC
						colUsers.SendPrivateToOps Settings.BotName,CStr("The username "&sSplot&" has been unbanned by "&curUser.sName&".")
					Case Else
						Reply curUser, "Sorry but "& sSplot &" is not banned",blnMC
				End Select
			Else
				Reply curUser, "Proper syntax is "& CmdPrefix &"unban <Nick/IP>",blnMC
			End If

		'----------------------------------------- 
		Case "tban"
		'----------------------------------------- 
			sSplit=Split(sData," ",4)
			If UBound(sSplit) < 2 Then Reply curUser, "Proper syntax is "&(CmdPrefix)&"tban <Nick> <Hours> <Reason>",blnMC :Exit Sub
			If Not colUsers.Online(CStr(sSplit(1))) Then Reply curUser, "User is not online in this hub", blnMC :Exit Sub

			If UBound(sSplit) = 2 Then
				r = "(No Reason Given)"
			Else
				r = sSplit(3)
			End If

			targetClass = colRegistered.Registered(CStr(sSplit(1)))
			If targetClass >= curUser.Class Then
				Reply curUser, "you cant kickban a higher class than yourself", blnMC
			Else
				Reply curUser, "The user "&CStr(sSplit(1))&" has been banned for "&CStr(sSplit(2))&" hours because: "&CStr(r), blnMC
				colUsers.SendPrivateToOps Settings.BotName,CStr("The user "&CStr(sSplit(1))&" has been banned for "&CStr(sSplit(2))&" hours by "&curUser.sName&" because: "&CStr(r))
				colUsers.ItemByName(CStr(sSplit(1))).SendPrivate Settings.BotName,"You have been banned for "&CStr(sSplit(2))&" hours because: "&CStr(r)
				frmHub.DoEventsForMe
				colIPBans.Add colUsers.ItemByName(CStr(sSplit(1))).IP,CStr(sSplit(2))*60
				colUsers.ItemByName(CStr(sSplit(1))).Disconnect
			End If
		'-----------------------------------------
		Case "tbanip"
		'-----------------------------------------
			sSplit=Split(sData," ",3)
			If UBound(sSplit) < 2 Then Reply curUser, "Proper syntax is "&(CmdPrefix)&"tbanip <IP> <Hours>",blnMC :Exit Sub
			Call colIPBans.Add(CStr(sSplit(1)),CStr(sSplit(2))*60)
			Reply curUser, "The IP "&CStr(sSplit(1))&" Has been banned for "&CStr(sSplit(2))&" hours", blnMC
			colUsers.SendPrivateToOps Settings.BotName,CStr("The IP "&CStr(sSplit(1))&" has been banned for "&CStr(sSplit(2))&" hours by "&curUser.sName&".")
			frmHub.DoEventsForMe
		'-----------------------------------------
		Case "cleartemp"
		'-----------------------------------------
			colIPBans.ClearTemp
			Reply curUser, "The Tempban list has been purged", blnMC
			colUsers.SendPrivateToOps Settings.BotName,CStr("The Tempban list has been purged by "&curUser.sName)
			frmHub.RefreshGUI
			frmHub.DoEventsForMe
		'-----------------------------------------
		Case "clearipbans"
		'-----------------------------------------
			colIPBans.ClearPerm
			Reply curUser, "The IP-Permban list has been purged",blnMC
			colUsers.SendPrivateToOps Settings.BotName,CStr("The IP-Permban list has been purged by "&curUser.sName)
			frmHub.RefreshGUI
			frmHub.DoEventsForMe
		'-----------------------------------------
		Case "listtempban"
		'----------------------------------------- 
			sTempIPList=Replace(CStr(colIPBans.TempList), "|", vbNewLine)
			curUser.SendPrivate Settings.BotName, "The Following IPs are temporarily banned:"&vbNewLine&sTempIPList
			frmHub.RefreshGUI
			frmHub.DoEventsForMe
		'-----------------------------------------
		Case "listpermban"
		'-----------------------------------------
			sPermIPList=Replace(CStr(colIPBans.PermList), "|", vbNewLine)
			curUser.SendPrivate Settings.BotName, "The Following IPs are permanently banned:"&vbNewLine&sPermIPList
			frmHub.RefreshGUI
			frmHub.DoEventsForMe

		'-----------------------------------------
		Case "flood"
		'-----------------------------------------
			Dim x
			sSplit=Split(sData," ",3)
			If UBound(sSplit) < 1 Then Reply curUser, "Proper syntax is "&(CmdPrefix)&"flood <user> <reason>",blnMC :Exit Sub
			If Not colUsers.Online(CStr(sSplit(1))) Then Reply curUser, "User is not online in this hub", blnMC :Exit Sub
			
			If UBound(sSplit) = 1 Then
				r = "(You've overstayed your welcome. You need to reboot before it's too late.)"
			Else
				r = sSplit(2)
			End If
			
			targetClass = colRegistered.Registered(CStr(sSplit(1)))
			
			If targetClass >= curUser.Class Then
				Reply curUser, "You can't flood a higher class than yourself", blnMC
			Else
				Reply curUser, "Flooding user "&CStr(sSplit(1))&" because: "&CStr(r), blnMC
				For i=1 To 1000
					x = Chr(Int(Rnd * 26 + 65)) + Chr(Int(Rnd * 26 + 65)) + Chr(Int(Rnd * 26 + 65)) + Chr(Int(Rnd * 26 + 65)) + Chr(Int(Rnd * 26 + 65)) + Chr(Int(Rnd * 26 + 65))
					colUsers.ItemByName(CStr(sSplit(1))).SendData "$Hello " + CStr(x) + "|"
					colUsers.ItemByName(CStr(sSplit(1))).SendPrivate CStr(x), CStr(r)
				Next
				colUsers.SendChatToAll Settings.BotName,CStr("Flooding user "&CStr(sSplit(1))&" because: "&CStr(r))
				frmHub.DoEventsForMe
			End If

		'-----------------------------------------
		Case "ipinfo"
		'-----------------------------------------
			sSplot = AfterFirst(sData," ")
			If LenB(sSplot) Then
				frmHub.DoEventsForMe
				If oIPs.Exists(CStr(sSplot)) Then
					sReplyInfo=CStr(""+VbTab&"Logon information from ip "&CStr(sSplot)&"\n"&oIPs(CStr(sSplot)))
				Else
					sReplyInfo="No data available"
				End If
				Reply curUser, CStr(sReplyInfo), blnMC
			Else
				Reply curUser, "Proper syntax is "&(CmdPrefix)&"ipinfo <IP>",blnMC
			End If
			frmHub.DoEventsForMe
			
		'-----------------------------------------
		Case "addreg"
		'-----------------------------------------
			sSplit = Split(sData," ",4)
			If UBound(sSplit) < 3 Then Reply curUser, "Proper syntax is "&(CmdPrefix)&"addreg <Nick> <Password> <Class (3(reg) 5(vip) 6(op) 8(sop) 10(admin)>",blnMC :Exit Sub
			If CByte(sSplit(3)) > CByte(curUser.Class) And CByte(curUser.Class) < 10 Then Reply curUser, "Can't register user of higher class.",blnMC :Exit Sub

			Select Case colRegistered.Add(CStr(sSplit(1)), CStr(sSplit(2)), CByte(sSplit(3)), curUser.sName)
				Case 0
					Reply curUser, "The "&Classname(sSplit(3))&" user: "&sSplit(1)&" made successfully with the password: "&sSplit(2),blnMC
					If colUsers.Online(CStr(sSplit(1))) Then
						colUsers.ItemByName(CStr(sSplit(1))).SendPrivate Settings.BotName,"You have been registered with the password: '"& sSplit(2) &"' Reconnect for it to take effect"
					End If
				Case 1
					Reply curUser, "Account already exist",blnMC
				Case 2
					Reply curUser, "Username is too long, use a shorter",blnMC
				Case 3
					Reply curUser, "Password longer than 20 characters",blnMC
			End Select
			frmHub.RefreshGUI()
			frmHub.DoEventsForMe

		'----------------------------------------- 
		Case "delreg"
		'----------------------------------------- 
			sSplot = AfterFirst(sData," ")
			If LenB(sSplot) Then
				targetClass = colRegistered.Registered(CStr(sSplot))		
				If targetClass = 0 Then Reply curUser, CStr(sSplot)&" isnt registered", blnMC :Exit Sub
				If targetClass > curUser.Class Then Reply curUser, "you cant unregister a higher class than yourself", blnMC :Exit Sub
				Reply curUser,"User: "& (CStr(sSplot)) & " Status Removed", blnMC
				Call colRegistered.Remove(CStr(sSplot))
			Else
				Reply curUser, "Proper syntax is "&(CmdPrefix)&"delreg <Nick>", blnMC
			End If	


		'-----------------------------------------  
		Case "setlanguage"
		'-----------------------------------------
			sSplit = Split(sData," ",3)

			If UBound(sSplit) < 2 Then Reply curUser, "Proper syntax is "&(CmdPrefix)&"setlanguage <Nick> <LanguageID>" & vbNewLine & ListLangs,blnMC :Exit Sub
			If curUser.Class < 5 Then sSplit(1)=curUser.sName
			If LCase(sSplit(1))="me" Then sSplit(1)=curUser.sName

			Select Case colRegistered.SetLanguage(CStr(sSplit(1)), CStr(LCase(sSplit(2))))
				Case 0
				  Reply curUser, "The language for user "&sSplit(1)&" was successfully set to '"&sSplit(2)&"'",blnMC
					If colUsers.Online(CStr(sSplit(1))) Then
						colUsers.ItemByName(CStr(sSplit(1))).SendPrivate Settings.BotName,"Your prefered language was set to: '"& sSplit(2) &"' Reconnect for it to take effect"
					End If
				Case 1
					Reply curUser, "Account doesn't exist",blnMC
				Case 2
					Reply curUser, "Not a valid language" & vbNewLine & ListLangs,blnMC
			End Select

		'-----------------------------------------  
		Case "say"
		'-----------------------------------------
			sSplit = Split(sData, " ",3)
			If UBound(sSplit)< 2 Then Reply curUser, "Proper syntax is "&(CmdPrefix)&"say <Nick> <ChatData>",blnMC :Exit Sub
			colUsers.SendToAll "<"&sSplit(1)&"> "&sSplit(2)&"|"
			SendToAdmin CStr(curUser.sName&" says: <"&sSplit(1)&"> "&sSplit(2)&"|")
			'colUsers.SendPrivateToOps Settings.BotName,Cstr(curUser.sName&" says: <"&sSplit(1)&"> "&sSplit(2)&"|")


		'-----------------------------------------
		Case "sayme"
		'-----------------------------------------
			sSplit = Split(sData, " ",3)
			If UBound(sSplit)< 2 Then Reply curUser, "Proper syntax is "&(CmdPrefix)&"sayme <Nick> <ChatData>",blnMC :Exit Sub
			colUsers.SendToAll "*"&sSplit(1)&" "&sSplit(2)&"|"
			SendToAdmin CStr(curUser.sName&" says: *"&sSplit(1)&" "&sSplit(2)&"|")
			'colUsers.SendPrivateToOps Settings.BotName,Cstr(curUser.sName&" says: <"&sSplit(1)&"> "&sSplit(2)&"|")

		'-----------------------------------------  
		Case "report"
		'-----------------------------------------
			sSplit = Split(sData, " ",3)
			If UBound(sSplit)< 2 Then Reply curUser, "Proper syntax is "&(CmdPrefix)&"report <User> <Reason>",blnMC :Exit Sub		
			colUsers.SendPrivateToOps Settings.OpChatname, "REPORT User: '" & sSplit(1) &"' Reason: '" & sSplit(2) & "' Reporter: '"& curUser.sName &"' IP: '"& curUser.IP &"'"
			Reply curUser, "Your report has been sent to the ops.",blnMC

		'-----------------------------------------
		Case "mass"
		'-----------------------------------------
			sSplot = AfterFirst(sData," ")
			If LenB(sSplot) Then
				colUsers.SendPrivateToAll CStr(Settings.BotName), CStr(sSplot)
				Reply curUser, "Sent to "&colUsers.Count&" users",blnMC
			Else
				Reply curUser, "Proper syntax is "&(CmdPrefix)&"pm <Message>",blnMC
			End If

		'-----------------------------------------
		Case "topic"
		'-----------------------------------------
			sTopic = AfterFirst(sData," ")
			If LenB(sSplot) Then
				SendToAdmin CStr(curUser.sName) & " changed topic to " & CStr(sTopic)
				colUsers.SendToAll "$HubName " & Settings.HubName & " - " & CStr(sTopic) & "|"
			Else
				sSplot=vbNullString
				SendToAdmin CStr(curUser.sName) & " changed topic to " & CStr(sTopic)
				colUsers.SendToAll "$HubName " & Settings.HubName &" - "& CStr(sTopic) & "|"
			End If
		'-----------------------------------------
		Case "userlimit"
		'-----------------------------------------
			sSplot = AfterFirst(sData," ")
			If LenB(sSplot) Then
				Settings.MaxUsers = Int(sSplot)
				Reply curUser, "Max users updated to: "&Settings.MaxUsers,blnMC
				frmHub.RefreshGUI()
			Else
				Reply curUser, "Proper syntax is "&(CmdPrefix)&"userlimit <amount>",blnMC
			End If
		'-----------------------------------------
		Case "setredirect"
		'-----------------------------------------
			sSplot = AfterFirst(sData," ")
			If LenB(sSplot) Then
				Settings.RedirectAddress = CStr(sSplot)
				Reply curUser, "Redirect Address updated to: "&Settings.RedirectAddress,blnMC
				frmHub.RefreshGUI()
			Else
				Reply curUser, "Proper syntax is "&(CmdPrefix)&"setredirect <address>",blnMC
			End If

		'-----------------------------------------
		Case "setminslots"
		'-----------------------------------------
			sSplot = AfterFirst(sData," ")
			If LenB(sSplot) Then
				Settings.MinSlots = CByte(sSplot)
				Reply curUser, "mininimum slots updated to: "&Settings.MinSlots,blnMC
				frmHub.RefreshGUI()
			Else
				Reply curUser, "Proper syntax is "&(CmdPrefix)&"setminslots <amount>",blnMC
			End If
		'-----------------------------------------
		Case "setmaxslots"
		'-----------------------------------------
			sSplot = AfterFirst(sData," ")
			If LenB(sSplot) Then
				Settings.MaxSlots = CByte(sSplot)
				Reply curUser, "Maximum slots updated to: "&Settings.MaxSlots,blnMC
				frmHub.RefreshGUI()
			Else
				Reply curUser, "Proper syntax is "&(CmdPrefix)&"setmaxslots <amount>",blnMC
			End If
		'-----------------------------------------
		Case "setslotratio"
		'-----------------------------------------
			sSplot = AfterFirst(sData," ")
			If LenB(sSplot) Then
				Settings.DCSlotsPerHub = CDbl(sSplot)
				Reply curUser, "slots pr hub updated to: "&Settings.DCSlotsPerHub,blnMC
				frmHub.RefreshGUI()
			Else
				Reply curUser, "Proper syntax is "&(CmdPrefix)&"setslotratio <amount>",blnMC
			End If

		'-----------------------------------------
		Case "resetports"
		'-----------------------------------------
			Dim iPort

			On Error Resume Next

			For iPort = 0 To frmHub.wskListen.UBound
				frmHub.wskListen(i).Close
				frmHub.wskListen(i).Listen
				Reply curUser, "Hubs listening ports reset",blnMC
			Next

			On Error GoTo 0

		'-----------------------------------------
		Case "lockchat"
		'-----------------------------------------
		If bFreezeMC = False Then
			bFreezeMC = True
			Reply curUser, "Mainchat is locked",blnMC
		Else
			bFreezeMC = False
			Reply curUser, "Mainchat is unlocked",blnMC
		End If

		'-----------------------------------------
		' MinShareSize  is the multiplyer
		' IMinShare is interface
		' Minshare is the actual share
		Case "setminshare"
		'-----------------------------------------
			sSplit = Split(sData," ",3)
			If UBound(sSplit) < 1 Then Reply curUser, "Proper syntax is "&(CmdPrefix)&"setminshare <amount> <unit> (Units are: B KB MB and GB)",blnMC :Exit Sub
				Select Case LCase(sSplit(2))
					Case "b"
						sUnit="0"
					Case "kb"
						sUnit="1"
					Case "mb"
						sUnit="2"
					Case "gb"
						sUnit="3"
					Case "tb"
						sUnit="4"
					Case Else
						Reply curUser, "Proper syntax is "&(CmdPrefix)&"setminshare <amount> <unit> (Units are: B KB MB GB or TB)",blnMC :Exit Sub
				End Select
				Settings.IMinShare = CDbl(sSplit(1))
				Settings.MinShare = CDbl(sSplit(1))
				Settings.MinShareSize = CByte(sUnit)
				Reply curUser, "Minimum share updated to: " & Settings.MinShare & " " &CStr(sSplit(2)), blnMC
				frmHub.RefreshGUI()


		'-----------------------------------------
		Case "setjoinmsg"
		'-----------------------------------------
			sSplot = AfterFirst(sData," ")
			If LenB(sSplot) Then
				Settings.JoinMsg=CStr(sSplot)
				SendToAdmin CStr("The join message has been changed by "&curUser.sName&vbNewLine&"The join message has been set to: "&vbNewLine&vbNewLine&sSplot&"|")
				frmHub.RefreshGUI()
				frmHub.DoEventsForMe
			Else
				Reply curUser, "Proper syntax is "&(CmdPrefix)&"setjoinmsg <Data>",blnMC
			End If

		'-----------------------------------------
		Case "hubinfo"
		'-----------------------------------------
			Reply curUser, ""& vbNewLine _
				&VbTab&"Hub Current Settings & Info " &vbNewLine _
				&VbTab&"Hubname: " & Settings.HubName&vbNewLine _
				&VbTab&"RedirectAddress: " & Settings.RedirectAddress&vbNewLine _
				&VbTab&"HubAddress: " & Settings.HubIP&vbNewLine _
				&VbTab&"MinShare: " & Settings.iMinShare&vbNewLine _
				&VbTab&"MaxShare: " & Settings.iMaxShare&vbNewLine _
				&VbTab&"MinSlots: " & Settings.MinSlots&vbNewLine _
				&VbTab&"Desc: " & Settings.HubDesc&vbNewLine _
				&VbTab&"Ports: " & Settings.Ports&vbNewLine _
				&VbTab&"MaxHubs: " & Settings.DCMaxHubs&vbNewLine _
				&VbTab&"MaxUsers: " & Settings.MaxUsers&vbNewLine _
				&VbTab&"SlotPerHub: " & Settings.DCSlotsPerHub&vbNewLine _
				&VbTab&"Default BanTime: " & Settings.DefaultBanTime&vbNewLine _
				&VbTab&"HubOwner: " & CStr(sOwner)&vbNewLine _
				&VbTab&"Forum: " & CStr(sForum)&vbNewLine _
				&VbTab&"HomePage: " & CStr(sHomePage)&vbNewLine,blnMC

		'-----------------------------------------
		Case "myinfo"
		'-----------------------------------------
			Reply curUser," "&vbNewLine _
				&VbTab&"Name: "&(curUser.sName)&vbNewLine _
				&VbTab&"My IP: "&(curUser.IP)&vbNewLine _
				&VbTab&"Operator:"&(curUser.bOperator)&vbNewLine _
				&VbTab&"Class: " &curUser.Class &vbNewLine _
				&VbTab&"Language:"&curUser.sLanguageID &vbNewLine _
				&VbTab&"Connected: "&(curUser.ConnectedSince)&vbNewLine _
				&VbTab&"My Share: "&ShareSize(curUser.iBytesShared)&vbNewLine _
				&VbTab&"My Share in bytes: "&(curUser.iBytesShared)& " (bytes)"&vbNewLine _
				&VbTab&"Time Elapsed in Hub: " &frmHub.MinToDate(DateDiff("n",curUser.ConnectedSince,Now))&vbNewLine,blnMC

		'-----------------------------------------
		Case "userinfo"
		'-----------------------------------------
			Dim sConn
			Dim sHubs
			Dim sSlots
			Dim sVersion
			Dim sOfflineinfo
			Dim sExists
			Dim userIP
			
			sSplot = AfterFirst(sData, " ")
			sExists = FileAccess.FileExists(".\Scripts\Zatarc-Lib\Userinfo\" & sSplot & ".txt")
			If LenB(sSplot) Then
				If colUsers.OnLine(CStr(sSplot)) Then
					Set objUser = colUsers.ItemByName(CStr(sSplot))
					sVersion=BetweenFirst(objUser.sMyInfoString, "V:",",")
					sHubs=BetweenFirst(objUser.sMyInfoString, "H:",",")
					sSlots=BetweenFirst(objUser.sMyInfoString, "S:",">")
					sConn=BetweenFirst(objUser.sMyInfoString, ">$ $","$")
					sReplyInfo=""
					If oIPs.Exists(CStr(colUsers.ItemByName(CStr(sSplot)).IP)) Then
						sReplyInfo=CStr("Logon information from ip "&CStr(colUsers.ItemByName(CStr(sSplot)).IP)&":"&vbNewLine&VbTab&VbTab & oIPs(CStr(colUsers.ItemByName(CStr(sSplot)).IP)))
					Else
						sReplyInfo=CStr("Logon information from ip "&CStr(colUsers.ItemByName(CStr(sSplot)).IP)&":"&vbNewLine&VbTab&VbTab&"No data available")
					End If
					Reply curUser,"User is Online: "&vbNewLine&vbNewLine _
						&VbTab&"Name: "&objUser.sName&vbNewLine _
						&VbTab&"IP : "&(objUser.IP)&vbNewLine _
						&VbTab&"Operator: "&(objUser.bOperator)&vbNewLine _
						&VbTab&"Language: "&(objUser.sLanguageID)&vbNewLine _
						&VbTab&"Client Version: "&CStr(sVersion) &vbNewLine _
						&VbTab&"Connected Since : "&(objUser.ConnectedSince)&vbNewLine _
						&VbTab&"Slots: "&CStr(sSlots) &vbNewLine _
						&VbTab&"Hubs: "&CStr(sHubs) &vbNewLine _
						&VbTab&"Line: "&CStr(sConn) &vbNewLine _
						&VbTab&"Share: "&ShareSize(objUser.iBytesShared)&vbNewLine _
						&VbTab&"Share in bytes: "&(objUser.iBytesShared)& " (bytes)"&vbNewLine _
						&VbTab&"Spent Time in Hub: " &frmHub.MinToDate(DateDiff("n",objUser.ConnectedSince,Now)) &vbNewLine _
						&VbTab&"Muted: "&objUser.Mute&vbNewLine _
						&VbTab&"Class: "&objUser.Class&vbNewLine&VbTab&sReplyInfo&vbNewLine,blnMC
				ElseIf sExists< 0 Then
					sOfflineinfo = "User is Offline: "&vbNewLine & FileAccess.ReadFile(".\Scripts\Zatarc-Lib\Userinfo\" & sSplot & ".txt")
					userIP = CStr(BeforeFirst(AfterFirst(sOfflineinfo, "IP: "), vbNewLine))
					If oIPs.Exists(CStr(userIP)) Then
						sOfflineinfo= sOfflineinfo&VbTab&CStr("Logon information from ip "&CStr(userIP)&":"&vbNewLine&VbTab&VbTab & oIPs(CStr(userIP)))
					End If
					Reply curUser, sOfflineinfo, blnMC
				Else
					Reply curUser, "User isnt online or not in hub",blnMC
				End If	
			Else
				Reply curUser, "Proper syntax is "&(CmdPrefix)&"userinfo <username>",blnMC
			End If

		'-----------------------------------------
		Case "ops"
		'-----------------------------------------
			Dim sOut,sTxt,sList,sUser,sPass

			sTxt = colRegistered.GetList(3, 0)
			sList  = Split(LeftB(sTxt, LenB(sTxt)-2), vbNewLine)
			For Each sUser In sList
			    Call FileAccess.WriteFile(CStr("LastUserCheck.txt"),CStr(sUser))
				sUser  = Split(sUser, "|")
				sPass = colRegistered.GetInfo(CStr(sUser(0)), "Password")
				Select Case curUser.Class
					Case 10,11
						If CInt(sUser(1)) >= 6 Then
							sOut  = sOut & VbTab & ClassName(CInt(sUser(1))) & VbTab & VbTab & _
							sUser(0) & VbTab & VbTab & sPass & vbNewLine
						End If
					Case Else
						If CInt(sUser(1)) >= 6 Then
							sOut  = sOut & VbTab & ClassName(CInt(sUser(1))) & VbTab & VbTab & _
							sUser(0) & vbNewLine
						End If
				End Select
			Next
			sOut = vbNewLine & sOut
			Reply curUser, sOut,blnMC	

		'-----------------------------------------
		Case "users"
		'-----------------------------------------   
			sTxt = colRegistered.GetList(3, 0)
			sList  = Split(LeftB(sTxt, LenB(sTxt)-2), vbNewLine)
			For Each sUser In sList
				Call FileAccess.WriteFile(CStr("LastUserCheck.txt"),CStr(sUser))
				sUser  = Split(sUser, "|")
				sPass = colRegistered.GetInfo(CStr(sUser(0)), "Password")
				Select Case curUser.Class
					Case 10,11
						If CInt(sUser(1)) >= 3 Then
							sOut  = sOut & VbTab & ClassName(CInt(sUser(1))) & VbTab & VbTab & _
							sUser(0) & VbTab & VbTab & sPass & vbNewLine
						End If
					Case 6,7,8,9
						If CInt(sUser(1)) >= 3 Then
							sOut  = sOut & VbTab & ClassName(CInt(sUser(1))) & VbTab & VbTab & _
							sUser(0) & vbNewLine
						End If
				End Select
			Next
			sOut = vbNewLine & sOut
			Reply curUser, sOut,blnMC
			
		'-----------------------------------------
		Case "iplist"
		'-----------------------------------------
			Dim Output
			Output = "IP Listing: "
			For Each sUser In colUsers
				Output = Output &vbNewLine &CStr(sUser.IP) &VbTab &VbTab &CStr(sUser.sName)				
			Next
			curUser.SendPrivate Settings.BotName,CStr(Output)
			
		'-----------------------------------------
		Case "cleandb"
		'-----------------------------------------
			Dim iDays
			Dim i
			iDays = 180  ' Accounts older than this amount (in days) will be deleted
			i = 0
			frmHub.oPermaCon.Execute "DELETE UsrClass.* FROM UsrClass, UsrDynamic WHERE DateDiff('d', UsrDynamic.LastLogin, Now()) > " & iDays & " AND UsrClass.UserName=UsrDynamic.UserName", i, 129
			Reply curUser, "All accounts older than " & CStr(iDays) & " have been deleted. This clean has deleted  " & i & " accounts",blnMC
			frmHub.DoEventsForMe
		'-----------------------------------------
		Case "shedule"
		'-----------------------------------------
			Dim pErg
			pErg = colSheduler.Plan(AfterFirst(CStr(sData)," "))
		  	Reply curUser, CStr(pErg), blnMC

'-----------------------------------------End Of commands-----------------------------------------
	End Select
End Sub

Sub UserConnected(curUser)
	If MuteList.Exists(CStr(curUser.sName)) Then
		curUser.Mute = True
	End If
	If bOfflineInfo = True Then

		'the Dim here will occure no matter if bOfflineInfo = True or not
		'it is better to call a sub to log those info only if it is enabled
		'that will prevent Dimming when not needed

		Dim sVersion
		Dim sHubs
		Dim sSlots
		Dim sConn
		Dim objUser
	
		Set objUser = colUsers.ItemByName(CStr(curUser.sName))
		sVersion=RegExps.CaptureSubStr(objUser.sMyInfoString, "V:(\d\.\d{1,4})")
		sHubs=RegExps.CaptureSubStr(objUser.sMyInfoString, "(H:(\d{1,3})/(\d{1,3})/(\d{1,3}))")
		sSlots=RegExps.CaptureSubStr(objUser.sMyInfoString, "S:(\d{1,3})")
		sConn=RegExps.CaptureSubStr(objUser.sMyInfoString, "\$[ ]\$(.{3,16})[\x01\x02\x03\x04\x05\x06\x07\x08\x09\x0A\x0B\x0C\x0D\x0E\x0F]")
		FileAccess.DeleteFile ".\Scripts\Zatarc-Lib\UserInfo\"&curUser.sName&".txt"
		FileAccess.AppendFile ".\Scripts\Zatarc-Lib\UserInfo\"&curUser.sName&".txt", "" &vbNewLine _ 
			&VbTab&"Name: "&(objUser.sName)&vbNewLine _
			&VbTab&"IP : "&(objUser.IP)&vbNewLine _
			&VbTab&"Operator: "&(objUser.bOperator)&vbNewLine _
			&VbTab&"Language: "&(objUser.sLanguageID)&vbNewLine _
			&VbTab&"Client Version: "&CStr(sVersion) &vbNewLine _
			&VbTab&"Connected Since : "&(objUser.ConnectedSince)&vbNewLine _
			&VbTab&"Slots: "&CStr(sSlots) &vbNewLine _
			&VbTab&"Hubs: "&CStr(sHubs) &vbNewLine _
			&VbTab&"Line: "&CStr(sConn) &vbNewLine _
			&VbTab&"Share: "&ShareSize(objUser.iBytesShared)&vbNewLine _
			&VbTab&"Share in bytes: "&(objUser.iBytesShared)& " (bytes)"&vbNewLine _
			&VbTab&"Spent Time in Hub: " &frmHub.MinToDate(DateDiff("n",objUser.ConnectedSince,Now)) &vbNewLine _
			&VbTab&"Muted: "&objUser.Mute&vbNewLine _
			&VbTab&"Class: "&objUser.Class&vbNewLine
	End If
	If bAddRightclick = True Then
		'curUser.SendData "<"&Settings.BotName&"> Enhanced Right Click Support for Zatarc "&sScriptv&" is Enabled |"
		Call iUserCommands(curUser)
	End If
	curUser.SendData "$HubName " & Settings.HubName & " - " & CStr(sTopic) & "|"
	If oIPs.Exists(CStr(curUser.IP)) Then
		oIPs(CStr(curUser.IP))=CStr(curUser.sName)+VbTab+sFdate(Now())+VbCrLf+oIPs(CStr(curUser.IP))
	Else
		oIPs(CStr(curUser.IP))=CStr(curUser.sName)+VbTab+sFdate(Now())
	End If
End Sub

Sub RegConnected(curUser)
	If MuteList.Exists(CStr(curUser.sName)) Then
		curUser.Mute = True
	End If
	If bAddRightclick = True Then
		Call UserConnected(curUser)
		If int(curUser.Class) > 3 Then Call iRegCommands(curUser)
	End If
	If oIPs.Exists(CStr(curUser.IP)) Then
		oIPs(CStr(curUser.IP))=CStr(curUser.sName)+VbTab+sFdate(Now())+VbCrLf+oIPs(CStr(curUser.IP))
	Else
		oIPs(CStr(curUser.IP))=CStr(curUser.sName)+VbTab+sFdate(Now())
	End If
End Sub

Sub OpConnected(curUser)
	If MuteList.Exists(CStr(curUser.sName)) Then
		curUser.Mute = True
	End If
	If oIPs.Exists(CStr(curUser.IP)) Then
		oIPs(CStr(curUser.IP))=CStr(curUser.sName)+VbTab+sFdate(Now())+VbCrLf+oIPs(CStr(curUser.IP))
	Else
		oIPs(CStr(curUser.IP))=CStr(curUser.sName)+VbTab+sFdate(Now())
	End If
	If bAddRightclick = True Then
		'curUser.SendData "<"&Settings.BotName&"> Enhanced Right Click Support for Zatarc "&sScriptv&" is Enabled |"
		Call iUserCommands(curUser)
		Call iRegCommands(curUser)
		Call iOPCommands(curUser)
		If int(curUser.Class) > 7 Then Call iSOPCommands(curUser)
		If int(curUser.Class) > 9 Then Call iAdminCommands(curUser)
	End If
	curUser.SendData "$HubName " & Settings.HubName & " - " & CStr(sTopic) & "|"
End Sub


'------------------------------------------------------------------
Sub tmrScriptTimer_Timer()
'------------------------------------------------------------------
	iCounter=iCounter+1 
	If iCounter=> iTime *120 Then
		iCounter=0
		MuteList.RemoveAll
		colUsers.SendPrivateToOps Settings.BotName ,"MuteList Cleared"
	End If 
End Sub
'------------------------------------------------------------------
Function IsBan(sData)
'suggestion:
'could use hub function to capture value(s)
'Set colMatches = RegExps.REMatchesCol(sData, "_ban_\d*(h|m|d)?")
'or 'Set colMatches = RegExps.REMatchesCol(sData, "(?:_ban_(\d{0,7})(h|m|d|s)?)"
'------------------------------------------------------------------

	Dim regEx
	Dim colMatches
	Dim sTime
	Dim sMultip
	Dim i

	IsBan = 0

	Set regEx = New RegExp

	regEx.Pattern = "_ban_\d*(h|m|d)?"
	regEx.IgnoreCase = True

	Set colMatches = regEx.Execute(sData)

	If colMatches.Count Then
		'get time
		sTime = AfterLast(colMatches(0).Value, "_")

		If LenB(sTime) Then
			'we have time given
			'check the last char
			i = AscW(RightB(sTime, 2))
			If i > 57 Or i < 48 Then
				'it's not number... RegExp makes sure that it's valid multiplier
				sMultip = RightB(sTime, 2)
				sTime = LeftB(sTime, LenB(sTime) - 2)
			End If

			Select Case sMultip
				Case "h"
					IsBan = sTime*60
				Case "m"
					IsBan = sTime
				Case "d"
					IsBan = sTime*1440
				Case Else
					IsBan = sTime '* your default multiplier
			End Select
		Else
			IsBan = -1 'no time given = perm
		End If
	End If
End Function

'------------------------------------------------------------------
Sub Reply(curUser, sMsg, blnMC)
'------------------------------------------------------------------
'Purpose: reply in main chat if message was received in main chat else reply in PM
'
'Params:	curUser :	object of the user on colUsers
'			sMsg :		message to be send
'			blnMC :		Msg is from main(true) chat or PM(false)
'
'------------------------------------------------------------------
'	curUser.SendPrivate Settings.BotName, cstr(sMsg)
	Select Case blnMC
		Case True: curUser.SendChat Settings.BotName, CStr(sMsg)
		Case False: curUser.SendPrivate Settings.BotName, CStr(sMsg)
	End Select
End Sub

'------------------------------------------------------------------
Sub Error(Line)
'------------------------------------------------------------------
'Purpose: Errors Logging
'
'Params:	Line :	line at witch the error occure
'
'------------------------------------------------------------------
	FileAccess.AppendFile FileAccess.AppPath & "\Scripts\Zatarc-Lib\Zatarc_Script_Error.log", Now & "|" & Err.Number & "|" & Err.Description & "|" & Line & "|"
	If colUsers.Online(CStr(sAdmin)) Then
		colUsers.ItemByName(CStr(sAdmin)).SendPrivate Settings.HubName & " Zatarc_Script_Error ", Now & "-" & "Error occured at line " & Line & " (Number: " & Err.Number & "; " & Err.Description & ")"
	End If
End Sub

#Include "Scripts\Zatarc-Lib\basZatarc_UCmds.bas"
#Include "Scripts\Zatarc-Lib\basZatarc_UnLoad.bas" 