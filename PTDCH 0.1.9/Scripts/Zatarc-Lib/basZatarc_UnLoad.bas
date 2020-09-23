'----------------------------------------- 
' Unload the Zatarc Commands
' TUFF
'----------------------------------------- 
Sub UnloadMain()

	colCommands.Remove("about")
	colCommands.Remove("me")
	colCommands.Remove("help")
	colCommands.Remove("report")

	colCommands.Remove("air")

	colCommands.Remove("ops")
	colCommands.Remove("motd")
	colCommands.Remove("network")
	colCommands.Remove("rules")
	colCommands.Remove("regme")
	colCommands.Remove("userinfo")
	colCommands.Remove("myinfo")
	colCommands.Remove("hubinfo")
	colCommands.Remove("myip")

	colCommands.Remove("addreg")
	colCommands.Remove("delreg")
	colCommands.Remove("setlanguage")

	colCommands.Remove("drop")
	colCommands.Remove("kick")
	colCommands.Remove("ban")
	colCommands.Remove("bannick")
	colCommands.Remove("banip")
	colCommands.Remove("tban")
	colCommands.Remove("tbanip")
	colCommands.Remove("unban")

	colCommands.Remove("cleartemp")
	colCommands.Remove("clearipbans")

	colCommands.Remove("listtempban")
	colCommands.Remove("listpermban")
	colCommands.Remove("flood")
	colCommands.Remove("ipinfo")
	colCommands.Remove("iplist")

	colCommands.Remove("mute")
	colCommands.Remove("unmute")

	colCommands.Remove("topic")
	colCommands.Remove("say")
	colCommands.Remove("sayme")
	colCommands.Remove("mass")
	colCommands.Remove("userlimit")
	colCommands.Remove("setminshare")
	colCommands.Remove("setredirect")
	colCommands.Remove("setminslots")
	colCommands.Remove("setmaxslots")
	colCommands.Remove("setslotratio")
	colCommands.Remove("setjoinmsg")
	colCommands.Remove("resetports")
	colCommands.Remove("lockchat")
	colCommands.Remove("users")
	colCommands.Remove("cleandb")
	colCommands.Remove("shedule")

End Sub