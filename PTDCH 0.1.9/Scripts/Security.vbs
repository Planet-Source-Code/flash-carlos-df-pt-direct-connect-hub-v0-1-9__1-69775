'------------------------------------------------------------------
'DDCH security enhancement Ver. 1.03 by TheNOP
'
'WARNING...This script is for version 0.3.38 svn 198 and higher only.
'		It WON'T work with earlyer DDCH versions.
'
'Features are :
'scan UNregistred users for advertising, while trying to minimise false possitive.
'Please add/change the authorised URL/mail addy in the allow array, to let unregistered users help other user in your hub.
'notice are sent to Ops so they can see if it can help or be harmfull to users and for admin to add the link in the allow list.
'
'notice of possible advertising are sent to sAdmin regardless of the sender class. sAdmin will see even those sent by admins in PMs.
'unregistering users for advertising is left at the discretion of sAdmin.
'
'Bloc ([main chat/PM]:optionnal)/Search/CTM/RCTM messages for 1 minute, after loggin in to the hub.(unregistered users only)
'prevent hit and run from spambots/searchbots(aka hubs runners) while not disturbing registred users.
'
'Bloc Search/CTM/RCTM messages of a user if he does not share enough.See LOWSHAREBLOC and MINSHARE settings
'
'Bloc advertising in nick/description/e-mail fields.(unfinish)
'
'Even if this script is fully fonctionnal, in my opignon it is not complete.
'It is up to the owners to add patterns to allow or deny Const as specific cases arise.
'
'Here some links where you can have help with regular expressions.
'http://www.regexlib.com/CheatSheet.htm
'http://www.regular-expressions.info/tutorial.html
'http://www.shadowdc.com/forums/index.php
'http://www.thescriptvault.net/forum//index.php

Option Explicit

Private m_dctDeny
Private m_dctUnMute

'/////////////// configurable ///////////////////
Const sAdmin = "fLaSh"

#Const NotifysAdmin = True	'True to enable URL spying on registred users messages.(messages containing "un-authorised to unreg stuff" only, not every PMs.)
#Const CHATDELAY = False	'True to enable chat delay at logging time. (affect unregistred users only.)
#Const ZLINEONLY = False	'True to deny non ZLine supporting clients in the hub
#Const LOWSHAREBLOC = False	'True to deny Search/CTM/RCTM, only bloc users with lower share size then minimum allowed
#Const ADVINMYINFO = False	'Deny advertising in nick/description or e-mail field
#Const FAKESHARE = False	'Deny Share size and/or digits sequences  in nick. ex: mynick123456, mynick99999 )
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


#If LOWSHAREBLOC Then
	Private m_objRequests
'/////////////// configurable ///////////////////
	Const UMINSHARE = 10	'Minimun share size require to be able to Search/CTM/RCTM. (in Bytes 10 = 10 Bytes, 1024 = 1 KB)
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	Dim iCounter
#End If

#If ADVINMYINFO Then
'/////////////// configurable ///////////////////
	Const TestURL = "(?:[^$,@\.)]{1,25}\.)?[^$,@\.)]{1,25}\.[^$,@\.)]{1,25}\.(?:[^$,@\.)]{1,25}\.)?[^$,@\.)]{2,5}"
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
#End If


#If FAKESHARE Then
'/////////////// configurable ///////////////////
	Const TestFake = ""
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
#End If

'/////////////// configurable ///////////////////
Const Deny = "(?:(?:\d{1,3}\.){3}\d{1,3})|(?:\S{1,25}\.\S{1,25}\.[^ \)]{2,5})|(?:w\s{0,15}?w\s{0,15}?w\s{0,15}?\.)|(?:h\s{0,15}?t\s{0,15}?t\s{0,15}?p\s{0,15}?:\s{0,15}?/\s{0,15}?/)|(?:d\s{0,15}?c\s{0,15}?h\s{0,15}?u\s{0,15}?b\s{0,15}?:\s{0,15}?/\s{0,15}?/)|(?:\.\s{0,15}?n\s{0,15}?o\s{0,15}?-?\s{0,15}?i\s{0,15}?p)|(?:\.\s{0,15}?c\s{0,15}?o\s{0,15}?m)"
Const Allow = "(?:\S{1,20}@\S{1,20}\.(?:\S{1,20}\.)?\S{1,5})|(?:baine-hell-realm\.no-ip\.info)|(?:\d\d:\d\d:\d\d\]|\[\d\d:\d\d\])|(?:google\.)"
Const PREVENTMCCLEANUP = "[\n\r\t]{7,}"
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'------------------------------------------------------------------
'HELPER FOR
'
'Default Const allow

'(?:\S{1,20}@\S{1,20}\.(?:\S{1,20}\.)?\S{1,5})		e-mail addy  ex: firstname.lastname@gmail.com, do not remove unless emails addy are not allowed
'(?:baine-hell-realm\.no-ip\.info)		baine-hell-realm.no-ip.info  <--for example only, can be removed
'(?:\[\d\d:\d\d:\d\d\]|\[\d\d:\d\d\])		dc time stamp format <-DO NOT remove this one if you keep ":port" denying pattern
'(?:google\.)					google.

'Default Const deny

'(?:(?:\d{1,3}\.){3}\d{1,3})								IPs format 0.0.0.0 to 999.999.999.999
'(?:\S{2,25}\.\S{2,25}\.[^ \)]{2,5})							any valid URL format including some e-mail addy, see allow list
'(?:w\s{0,15}?w\s{0,15}?w\s{0,15}?\.)							www or w  w  w  .
'(?:h\s{0,15}?t\s{0,15}?t\s{0,15}?p\s{0,15}?:\s{0,15}?/\s{0,15}?/)			http:// or h  t  t  p  :  /    /
'(?:d\s{0,15}?c\s{0,15}?h\s{0,15}?u\s{0,15}?b\s{0,15}?:\s{0,15}?/\s{0,15}?/)	dchub:// or d  c  h  u    b  :  /  /
'(?:\.\s{0,15}?n\s{0,15}?o\s{0,15}?-?\s{0,15}?i\s{0,15}?p)				.no-ip or .  n  o    -  i  p
'(?:\.\s{0,15}?c\s{0,15}?o\s{0,15}?m)							.com or .  c    o  m

'------------------------------------------------------------------
Sub Main()

	Set m_dctDeny = NewDictionary
	Set m_dctUnMute = NewDictionary

  #If LOWSHAREBLOC Then
	Set m_objRequests=NewDictionary
  #End If

	AddScriptMessages ".\Scripts\Security_Lib\Security.lng"

	tmrScriptTimer.Interval = 10000
	tmrScriptTimer.Enabled = True

End Sub

'------------------------------------------------------------------
Sub tmrScriptTimer_Timer()

	Dim sUser
	Dim strUser

	If m_dctDeny.Count Then
		For Each sUser In m_dctDeny.Keys
			If DateDiff("n", Now, m_dctDeny.Item(CStr(sUser))) < -1 Then m_dctDeny.Remove(CStr(sUser))
		Next
	End If

	If m_dctUnMute.Count Then
		For Each strUser In m_dctUnMute.Keys
			If DateDiff("n", Now, m_dctUnMute.Item(CStr(strUser))) < -2 Then
				If colUsers.Online(CStr(strUser)) Then
					colUsers.ItemByName(CStr(strUser)).Mute = False
					colUsers.ItemByName(CStr(strUser)).SendChat Settings.BotName, CStr(colUsers.ItemByName(CStr(strUser)).GetCoreMsgStr("SecuUserUnMute"))
				End If

				m_dctUnMute.Remove(CStr(strUser))
			End If
		Next
	End If

  #If LOWSHAREBLOC Then
	'Make sure we don't keep old object in collection
	iCounter=iCounter+1

	If iCounter > 60 Then
		iCounter=0
		m_objRequests.RemoveAll
	End If
  #End If

End Sub

'------------------------------------------------------------------
Function PreDataArrival(curUser, sData)

	Dim aData

	aData = Split(MidB(sData, 3), " ", 3)

	Select Case AscW(sData)
		Case 36
			#If NotifysAdmin Then
				If LeftB(sData,8) = "$To:" Then
					If colRegistered.Registered(curUser.sName) Then
						If colUsers.Online(CStr(sAdmin)) Then
							'spy on links sent by registered users.
							If RegExps.AdvertTest(CStr(sData), CStr(Deny), CStr(Allow)) Then
								If colUsers.Online(CStr(sAdmin)) Then colUsers.ItemByName(CStr(sAdmin)).SendPrivate Settings.HubName & "Security", Now & "-" & TagReplace(curUser.GetCoreMsgStr("SecuAdmWarn"), curUser) & sData
							End If
						End If
					End If
				End If
			#End If

			If Not CBool(colRegistered.Registered(curUser.sName)) Then
				Select Case CStr(aData(0))
					Case "Search"
						If m_dctDeny.Exists(curUser.sName) Then PreDataArrival = Empty :Exit Function
					Case "ConnectToMe"
						If m_dctDeny.Exists(curUser.sName) Then PreDataArrival = Empty :Exit Function
					Case "RevConnectToMe"
						If m_dctDeny.Exists(curUser.sName) Then PreDataArrival = Empty :Exit Function
					Case "To:"
					  #If CHATDELAY Then
							If m_dctDeny.Exists(curUser.sName) Then PreDataArrival = Empty :Exit Function
					  #End If

						If Not curUser.Mute Then
							'scan for advertising
							If RegExps.AdvertTest(CStr(sData), CStr(Deny), CStr(Allow)) Then
								'warn Ops
								colUsers.SendPrivateToOps Settings.BotName, Now & "-" & Replace(TagReplace(curUser.GetCoreMsgStr("SecuOpsWarn"), curUser), "%[chat]", "PM") & sData
								curUser.Mute = True
								'warn user
								curUser.SendPrivate Settings.BotName, CStr(curUser.GetCoreMsgStr("SecuUserWarn"))
								m_dctUnMute.Add curUser.sName, Now
								PreDataArrival = Empty
								Exit Function
							End If
						End If
					Case Else
						'Relay to other script(s) and possibly to hub
						'PreDataArrival=sData
				End Select
			End If

		Select Case CStr(aData(0))

		#If LOWSHAREBLOC Then
			Case "Search"
				If curUser.State = 5 Then
					If curUser.iBytesShared < CDbl(UMINSHARE) Then
						PreDataArrival=Empty
						Exit Function
					End If
				End If
			Case "RevConnectToMe"
				If curUser.State = 5 Then
					If curUser.iBytesShared > CDbl(UMINSHARE) Then
						If Not colUsers.Online(CStr(aData(2))) Then PreDataArrival=Empty :Exit Function

						If colUsers.ItemByName(CStr(aData(2))).iBytesShared < CDbl(UMINSHARE) Then
							If m_objRequests.Exists(CStr(aData(2))) Then
								m_objRequests(CStr(aData(2))) = CStr(aData(1))
							Else
								m_objRequests.Add CStr(aData(2)), CStr(aData(1))
							End If
						End If
					Else
						PreDataArrival=Empty
						Exit Function
					End If
				End If
			Case "ConnectToMe"
				If curUser.State = 5 Then
					If curUser.iBytesShared > CDbl(UMINSHARE) Then
						'PreDataArrival=sData
					ElseIf m_objRequests.Exists(CStr(curUser.sName)) Then
						If m_objRequests(CStr(curUser.sName))=CStr(aData(1)) Then
							m_objRequests.Remove(CStr(curUser.sName))
						Else
							PreDataArrival=Empty
							Exit Function
						End If
					Else
						PreDataArrival=Empty
						Exit Function
					End If
				End If
		#End If

		#If ZLINEONLY Or ADVINMYINFO Or FAKESHARE Then
			Case "MyINFO"
				'Make sure hub pingers are not affected
				Select Case RegExps.CaptureSubStr(sData, "\$ALL\s([\S]{1,40})")
					Case "{HubListPinger}","[rug.nl]pinger","[1stleg]Pinger","pInger"
						'add a send main chat/PM message here to monitor pingers if you want.
						PreDataArrival = sData
						Exit Function
				End Select

				'Nick changed in MyINFO without user disconnecting.
'				If Not curUser.State = 5 Then
'					If curUser.QuickList Then
'						PreDataArrival = sData
'						Exit Function
'					Else
'						If Not MidB(arrData(2),1,LenB(curUser.sName))=curUser.sName Then
'							colUsers.SendChatToOps "MyINFO_Spoof_detection", "MyINFO Spoofing detected..."&vbNewLine&"From : "&curUser.sName&" IP : "&curUser.IP& vbNewLine& sData
'							curUser.Kick(60)
'							PreDataArrival = Empty
'							Exit Function
'						End If
'					End If
'				Else
'					If Not MidB(arrData(2),1,LenB(curUser.sName))=curUser.sName Then
'						colUsers.SendChatToOps "MyINFO_Spoof_detection", "MyINFO Spoofing detected..."&vbNewLine&"From : "&curUser.sName&" IP : "&curUser.IP& vbNewLine& sData
'						curUser.Kick(60)
'						PreDataArrival = Empty
'						Exit Function
'					End If
'				End If
				
			#If FAKESHARE Then
					If RegExps.TestStr(sData, CStr(TestFake)) Then
					curUser.SendPrivate Settings.BotName, "Bad Nick and/or Share size"
					frmHub.DoEventsForMe
					curUser.Kick
				End If
			#End If

			#If ADVINMYINFO Then
				'will probably need 2 check, description and email. check if @ in email field first...
				'If RegExps.TestStr(sData, CStr(TestURL)) Then
				'	If colUsers.Online(cstr(sAdmin)) Then colUsers.ItemByName(cstr(sAdmin)).SendPrivate Settings.HubName & " Debug", Now & "-" & sData
				'End If
			#End If

			#If ZLINEONLY Then
				'Only if not registred
				If Not CBool(colRegistered.Registered(RegExps.CaptureSubStr(sData, "\$ALL\s([\S]{1,40})"))) Then
					If Not (curUser.ZLine Or curUser.ZPipe) Then

						curUser.SendPrivate Settings.BotName, "Your client does not support ZLine, " _
							& vbNewLine & "Links to clients that support it:" _
							& vbNewLine & "http://virus27.free.fr/zion++/zion++_vert.php" _
							& vbNewLine & "http://www.imperialnet.org/forum/index.php?topic=1883.0" _
							& " bye bye..."

						frmHub.DoEventsForMe
						curUser.Kick
					End If
				End If
			#End If
		#End If

			Case Else
		End Select


		Case 60
			'only for Unregistered users
			If Not CBool(colRegistered.Registered(curUser.sName)) Then
				#If CHATDELAY Then
					If m_dctDeny.Exists(curUser.sName) Then PreDataArrival = Empty :Exit Function
				#End If

				'scan for advertising
				If Not curUser.Mute Then
					If RegExps.AdvertTest(sData, CStr(Deny), CStr(Allow)) Then
						'warn Ops
						colUsers.SendPrivateToOps Settings.BotName, Now & "-" & Replace(TagReplace(curUser.GetCoreMsgStr("SecuOpsWarn"), curUser), "%[chat]", "main chat") & sData
						curUser.Mute = True
						'warn user
						curUser.SendChat Settings.BotName, CStr(curUser.GetCoreMsgStr("SecuUserWarn"))
						m_dctUnMute.Add curUser.sName, Now
						PreDataArrival = Empty
						Exit Function
					End If
				End If
			End If

			'Prevent main chat cleanup from clients, if user is not an OP
			If Not curUser.bOperator Then
				If RegExps.TestStr(sData, PREVENTMCCLEANUP) Then PreDataArrival = Empty :Exit Function
			End If

		Case Else
			'Relay to other script(s) and possibly to hub
			'in case script(s) use an other protocol or new protocol is added to hub
			'PreDataArrival = sData
	End Select
	'If we have not exit it mean that all seem to be ok
	'relay sData to the hub and/or other script(s)
	PreDataArrival = sData

End Function

'------------------------------------------------------------------
Sub UserConnected(curUser)

	'add a chat/Search/CTM/RCTM delay for the user
	If Not m_dctDeny.Exists(curUser.sName) Then m_dctDeny.Add curUser.sName, Now

	're-mute user if his mute time is not finish
	If m_dctUnMute.Exists(curUser.sName) Then curUser.Mute = True

End Sub

'------------------------------------------------------------------
Sub UserQuit(curUser)

	'make sure the user won't get re-mute on reconnect if an OP un-mute the user before his time end up
	If m_dctUnMute.Exists(curUser.sName) Then
		If Not curUser.Mute Then m_dctUnMute.Remove(curUser.sName)
	End If

  #If LOWSHAREBLOC Then
	If m_objRequests.Exists(curUser.sName) Then
		m_objRequests.Remove(curUser.sName)
	End If
  #End If

End Sub

'------------------------------------------------------------------
Function TagReplace(strString, curUser)
'------------------------------------------------------------------
'Purpose:	Replace %[var] with proper data
'
'Params:	curUser
'			Current user object
'
'Return:	Modified string or same input string if no var(s) were replaced
'------------------------------------------------------------------

	If InStrB(strString, "%[user]") Then strString = Replace(strString, "%[user]", curUser.sName)
	If InStrB(strString, "%[ip]") Then strString = Replace(strString, "%[ip]", curUser.IP)
	If InStrB(strString, "%[nl]") Then strString = Replace(strString, "%[nl]", vbNewLine)

	TagReplace = strString

End Function

'------------------------------------------------------------------
Sub Error(Line)

	FileAccess.AppendFile FileAccess.AppPath & "\Scripts\Security_Lib\Security_Script_Error.log", Now & "|" & Err.Number & "|" & Err.Description & "|" & Line & "|"

	If colUsers.Online(CStr(sAdmin)) Then
		colUsers.ItemByName(CStr(sAdmin)).SendPrivate Settings.HubName & "Security_Script_Error ", Now & "-" & "Error occured at line " & Line & " (Number: " & Err.Number & "; " & Err.Description & ")"
	End If
End Sub 