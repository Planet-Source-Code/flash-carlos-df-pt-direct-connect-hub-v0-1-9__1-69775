'----------------------------------------- 
' UserCommands sending
' GhOstFaCE @ Asgard
'----------------------------------------- 
Sub iUserCommands(CurUser)
	CurUser.SendData "$UserCommand 0 3 |" 'separator
	CurUser.SendData "$UserCommand 1 2 "&Settings.HubName&"\Hub Rules$<%[mynick]> "&ChrW(Settings.CPrefix)&"rules&#124;|"
	CurUser.SendData "$UserCommand 1 3 "&Settings.HubName&"\Show Hub Command List$<%[mynick]> "&ChrW(Settings.CPrefix)&"help&#124;|"
	CurUser.SendData "$UserCommand 1 3 "&Settings.HubName&"\Show Your Info$<%[mynick]> "&ChrW(Settings.CPrefix)&"myinfo&#124;|"
	CurUser.SendData "$UserCommand 1 3 "&Settings.HubName&"\Show Your IP$<%[mynick]> "&ChrW(Settings.CPrefix)&"myip&#124;|"
	CurUser.SendData "$UserCommand 1 2 "&Settings.HubName&"\Report Misconduct$<%[mynick]> "&ChrW(Settings.CPrefix)&"report %[line:user] %[line:reason]&#124;|"
	CurUser.SendData "$UserCommand 1 3 "&Settings.HubName&"\Show login message$<%[mynick]> "&ChrW(Settings.CPrefix)&"motd&#124;|"
	CurUser.SendData "$UserCommand 1 3 "&Settings.HubName&"\Show Hub Info$<%[mynick]> "&ChrW(Settings.CPrefix)&"hubinfo&#124;|"
	CurUser.SendData "$UserCommand 1 3 "&Settings.HubName&"\Script Info$<%[mynick]> "&ChrW(Settings.CPrefix)&"about&#124;|"
	CurUser.SendData "$UserCommand 1 3 "&Settings.HubName&"\Network Info$<%[mynick]> "&ChrW(Settings.CPrefix)&"network&#124;|"
	CurUser.SendData "$UserCommand 1 3 "&Settings.HubName&"\Show Operators$<%[mynick]> "&ChrW(Settings.CPrefix)&"ops&#124;|"
	CurUser.SendData "$UserCommand 1 2 "&Settings.HubName&"\Self Register$<%[mynick]> "&ChrW(Settings.CPrefix)&"regme %[line:Password]&#124;|"
End Sub

Sub iRegCommands(CurUser)
	CurUser.SendData "$UserCommand 2 3 "&Settings.HubName&"\Speak in the third person$<%[mynick]> "&ChrW(Settings.CPrefix)&"me %[line:text]&#124;|"
	CurUser.SendData "$UserCommand 2 3 "&Settings.HubName&"\Set User's Language$<%[mynick]> "&ChrW(Settings.CPrefix)&"setlanguage %[nick] %[line:LanguageID ('DE', 'EN', 'FR']&#124;|"
End Sub

Sub iOPCommands(CurUser)
	CurUser.SendData "$UserCommand 0 3 |" 'separator
	CurUser.SendData "$UserCommand 1 3 Operator\Get Userinfo$<%[mynick]> "&ChrW(Settings.CPrefix)&"userinfo %[nick]&#124;|"
	CurUser.SendData "$UserCommand 1 3 Operator\Get IP Listing$<%[mynick]> "&ChrW(Settings.CPrefix)&"iplist&#124;|"
	CurUser.SendData "$UserCommand 2 6 Operator\Disconnect a user$<%[mynick]> "&ChrW(Settings.CPrefix)&"drop %[nick] %[line:Reason?]&#124;|"
	CurUser.SendData "$UserCommand 1 3 Operator\Kick User$<%[mynick]> "&ChrW(Settings.CPrefix)&"kick %[nick] %[line:reason?]&#124;|"
	CurUser.SendData "$UserCommand 2 6 Operator\Bans Nick For X Hours$<%[mynick]> "&ChrW(Settings.CPrefix)&"tban %[nick] %[line:hours?] %[line:Reason?]&#124;|"
	CurUser.SendData "$UserCommand 1 3 Operator\Bans IP For X Hours$<%[mynick]> "&ChrW(Settings.CPrefix)&"tbanip %[line:ip] %[line:hours]&#124;|"
	CurUser.SendData "$UserCommand 1 3 Operator\Mute User/s$<%[mynick]> "&ChrW(Settings.CPrefix)&"mute %[nick]&#124;|"
	CurUser.SendData "$UserCommand 1 3 Operator\UnMute User/s$<%[mynick]> "&ChrW(Settings.CPrefix)&"unmute %[nick]&#124;|"
End Sub

Sub iSOPCommands(CurUser)
	CurUser.SendData "$UserCommand 1 3 Super Operator\Ban A Nick$<%[mynick]> "&ChrW(Settings.CPrefix)&"bannick %[nick] %[line:Reason?]&#124;|"
	CurUser.SendData "$UserCommand 2 6 Super Operator\Ban an IP$<%[mynick]> "&ChrW(Settings.CPrefix)&"banip %[line:ip]&#124;|"
	CurUser.SendData "$UserCommand 1 3 Super Operator\Kick And Ban User By IP And Nick$<%[mynick]> "&ChrW(Settings.CPrefix)&"ban %[nick] %[line:Reason?]&#124;|"
	CurUser.SendData "$UserCommand 1 3 Super Operator\List Temp IP Bans$<%[mynick]> "&ChrW(Settings.CPrefix)&"listtempban&#124;|"
	CurUser.SendData "$UserCommand 2 6 Super Operator\Unbans IP Or Nick$<%[mynick]> "&ChrW(Settings.CPrefix)&"unban %[line:Nick or IP]&#124;|"
	CurUser.SendData "$UserCommand 2 6 Super Operator\Change The Hub Topic$<%[mynick]> "&ChrW(Settings.CPrefix)&"topic %[line:Topic]&#124;|"
	CurUser.SendData "$UserCommand 1 3 Super Operator\Impersonate User$<%[mynick]> "&ChrW(Settings.CPrefix)&"say %[nick] %[line:Data]&#124;|"
	CurUser.SendData "$UserCommand 1 3 Super Operator\Impersonate User /+me$<%[mynick]> "&ChrW(Settings.CPrefix)&"sayme %[nick] %[line:Data]&#124;|"
	CurUser.SendData "$UserCommand 1 3 Super Operator\Mass Message$<%[mynick]> "&ChrW(Settings.CPrefix)&"mass %[line:Message]&#124;|"
	CurUser.SendData "$UserCommand 1 3 Super Operator\Change The Hubs Max Users$<%[mynick]> "&ChrW(Settings.CPrefix)&"userlimit %[line:New Limit]&#124;|"
	CurUser.SendData "$UserCommand 2 6 Super Operator\Change The Hubs Redirect Address$<%[mynick]> "&ChrW(Settings.CPrefix)&"setredirect %[line:New addy]&#124;|"
End Sub

Sub iAdminCommands(CurUser)
	CurUser.SendData "$UserCommand 2 6 Admin\Add A New Account$<%[mynick]> "&ChrW(Settings.CPrefix)&"addreg %[nick] %[line:Password?] %[line:Class (3(reg) 5(vip) 6(op) 8(sop) 10(admin)]&#124;|"
	CurUser.SendData "$UserCommand 2 6 Admin\Delete an account$<%[mynick]> "&ChrW(Settings.CPrefix)&"delreg %[line:Nick]&#124;|"
	CurUser.SendData "$UserCommand 1 3 Admin\List Perm IP Bans$<%[mynick]> "&ChrW(Settings.CPrefix)&"listpermban&#124;|"
	CurUser.SendData "$UserCommand 1 2 Admin\Purge Temp Bans$<%[mynick]> "&ChrW(Settings.CPrefix)&"cleartemp&#124;|"
	CurUser.SendData "$UserCommand 1 3 Admin\Purge Perm IP Bans$<%[mynick]> "&ChrW(Settings.CPrefix)&"clearipbans&#124;|"
	CurUser.SendData "$UserCommand 1 3 Admin\Flood$<%[mynick]> "&ChrW(Settings.CPrefix)&"flood %[nick] %[line:Reason]&#124;|"
	CurUser.SendData "$UserCommand 2 6 Admin\Set Join Message$<%[mynick]> "&ChrW(Settings.CPrefix)&"setjoinmsg %[line:Message]&#124;|"
	CurUser.SendData "$UserCommand 2 6 Admin\Change The Hubs Minimum Share$<%[mynick]> "&ChrW(Settings.CPrefix)&"setminshare %[line:amount] %[line:units (B KB MB or GB)]&#124;|"
	CurUser.SendData "$UserCommand 2 6 Admin\Change The Hubs Maximum Hub Limit$<%[mynick]> "&ChrW(Settings.CPrefix)&"setmaxhubs %[line:new max hubs]&#124;|"
	CurUser.SendData "$UserCommand 2 6 Admin\Change The Hubs Minimum Slots$<%[mynick]> "&ChrW(Settings.CPrefix)&"setminslots %[line:new min slots]&#124;|"
	CurUser.SendData "$UserCommand 2 6 Admin\Change The Hubs Maximum Slots$<%[mynick]> "&ChrW(Settings.CPrefix)&"setmaxslots %[line:new max slots]&#124;|"
	CurUser.SendData "$UserCommand 2 6 Admin\Change The Hubs Slot Ratio$<%[mynick]> "&ChrW(Settings.CPrefix)&"setslotratio %[line:new slot ratio]&#124;|"
	CurUser.SendData "$UserCommand 2 6 Admin\Reset Ports$<%[mynick]> "&ChrW(Settings.CPrefix)&"resetports&#124;|"
	CurUser.SendData "$UserCommand 2 6 Admin\Clean Database$<%[mynick]> "&ChrW(Settings.CPrefix)&"cleandb&#124;|"
End Sub