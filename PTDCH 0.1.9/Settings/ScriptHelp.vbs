'----------------------------------------------------------------------------------------------------------------------------------------------
'=========================================================== Scripting - Interface ============================================================
'----------------------------------------------------------------------------------------------------------------------------------------------

'Object list
	ScriptCtrl		'Current script control object
	tmrScriptTimer	'The script's timer
	wskScript		'Control array of winsocks
	frmHub			'The main form
	mdiScriptEditor	'Script editor form
	colUsers		'Collection class for the users
	colRegistered		'Easy access to the registered user list (NOT a real collection)
	colIPBans		'Easy access to the IP ban list (NOT a real collection)
	colCommands		'Commands collection
	colStatic		'Dictionary to store variables which are not lost upon reset
	Settings		'Access various settings here
	Functions		'Various functions are found here for scripts which can be accessed with the "Functions." part
	FileAccess		'Various file related functions
	App			'VB's app object
	Forms			'VB's collection-like forms object
	colLanguages		'Collection of languages strings	

'----------------------------------------------------------------------------------------------------------------------------------------------
'Functions (clsFunctions) (can be accessed WITHOUT the "Functions.")
	'Creates a new collection object (ex Set mCol = NewCollection)
	Function NewCollection() As Collection
	
	'Gain access to a huffman (de)compression class
	Function NewHuffman() As clsHuffman	
	
	'Gain access to a BZip2 (de)compression class
	Function NewBZip2() As clsBZ2
	
	'Gain access to a ZLib (de)compression class
	Function NewZLib() As clsZLib	
	
	'Creates a new clsXMLParser class
	Function NewXMLParser() As clsXMLParser	
	
	'Creates a new clsXMLNode class
	Function NewXMLNode() As clsXMLNode		
	
	'Creates a new clsXMLAttribute class
	Function NewXMLAttribute() As clsXMLAttribute	
	
	'Creates a new Connection object
	Function NewConnection() As Connection
	
	'Creates a new JetEngine object
	Function NewJetEngine() As JetEngine	
	
	'Creates a new clsDictionary object (equivalent of Scripting.Dictionary object except for Keys/Items return Collection object rather than arrays)
	Function NewDictionary() As clsDictionary	
	
	'Loads the new object (for use with wskScript(x) mainly)
	Sub LoadObj(Object As Object)			
	
	'Unloads the object
	Sub UnloadObj(Object As Object)			
	
	'Use the Shell command on sFile
	Function ShellExec(sFile As String, Optional sParameters As String, _				
			Optional sDirectory As String, Optional lShowCmd As Long) As Long
			
	'Determines the closest size in the approriate *B (ex 1024 returns 1.00 KB)
	Function ShareSize(lBytes As Double) As String				
	
	'Returns string after the first occurance of sFind *
	Function AfterFirst(sString As String, sFind As String) As String	
	
	'Returns string before the first occurance of sFind *
	Function BeforeFirst(sString As String, sFind As String) As String	
	
	'Returns string after the last occurance of sFind *
	Function AfterLast(sString As String, sFind As String) As String	
	
	'Returns string before the last occurance of sFind *
	Function BeforeLast(sString As String, sFind As String) As String	
	
	'Returns string between the first occurance of sFirst and sSecond *
	Function BetweenFirst(sString As String, sFirst As String, sSecond As String) As String		
	
	'Return string between the last occurance of sFirst and sSecond *
	Function BetweenLast(sString As String, sFirst As String, sSecond As String) As String		
	
	'Converts a date to that which is used in the user database
	Function DBDate(dDate As Date) As String	
	
	'Converts a numeric class value to it's string representation
	Function ClassName(intClass As enuClass) As String	
	
	'Returns the milliseconds since the CPU started
	Function TickCount() As Long	
	
	'Returns a string of random characters of length lngLen (range used is min = chrStart to max = chrEnd)
	Function RandomChars(lngLen As Long, chrStart As String, chrEnd As String) As String	
	
	'Converts a user class variant to a user class clsUser object
	Function CUser(varUser As Variant) As clsUser							

											'(Just for JScripts, but VBScripts can use it)
	'Wrapper for the input box function
	Function Prompt(sMessage As String, Optional sTitle As String, Optional sDefault _		
			As String) As String
			
	'Wrapper for the message box function
	Function Alert(sMessage As String, Optional sTitle As String, _
			Optional sDefault As String) As VbMsgBoxResult
			
	'Return a string from UsersMessages.xml(From default strings collection "en")
	Function GetENLangStr(strStringID As String) As String		
	
	'Use when sure the user is not registered and want to send a default hub's reason
	'Return True if strLangID is a supported languages.(users languages support)
	Function ValidLang(strLangID As String) As Boolean		
	
	'Return formated listing of supported languages(users lauguages)
	'languages Id / language international name / localised lauguage name
	Function ListLangs() As String		
	
	'Return strMessage with replacement for the following variables: 
		'Replace %[nick] with curUser.name.
		'Replace %[ip] with curUser.IP.
	Function ReplaceUserVars(curUser As clsUser, strMessage As String) As String	
	
	'Return strMessage with replacement for the following variables:
		'Replace %[maxhubs] with hub's Max hubs current setting
		'Replace %[minslots] with hub's Min slots current setting
		'Replace %[hsratio] with hub's Hub/slot ratio current setting
		'Replace %[minshare] with hub's Min share current setting
	Function ReplaceUserVars(curUser As clsUser, strMessage As String) As String			

	'* Does not require that you use CStr() to convert variant variables to strings for the values of sString / sFirst / sSecond / sFind
'----------------------------------------------------------------------------------------------------------------------------------------------
'FileAccess (clsFileAccess)
	'Read the text for the file located at sPath and returns it
	Function ReadFile(sPath As String) As String		
	
	'Writes the text, sData, to the file at sPath
	Sub WriteFile(sPath As String, sData As String)			
	
	'Appends the text to the file
	Sub AppendFile(sPath As String, sAppend As String, Optional bCarriageReturn = True)	
	
	'Deletes a file
	Sub DeleteFile(sPath As String)					
	
	'Renames a file (also can be used to move a file as sOld and sNew as you need the full path)
	Sub RenameFile(sOld As String, sNew As String)			
	
	'Copies sPath to sCopy
	Sub CopyFile(sPath As String, sCopy As String)			
	
	'Creates a directory
	Sub CreateDir(sPath As String)					
	
	'Gets the attributes of a file
	Function FileAttributes(sPath As String) As VbFileAttribute		
	
	'Check if a file or directory exists
	Function FileExists(sPath As String) As Boolean			
	
	'Returns the base directory where PTDCH is installed/running
	Property AppPath() As String						
	
	'Calls VBs Dir function
	Function VDir(Optional PathName As String, Optional Attributes As VbFileAttribute) As String	
	
	'Returns a string value from INI file (sDefault returned if not found)			
	Function GetSStr(Section As String, Key As String, Default As String, File As String) As String	
	
	'Returns a long value from INI file (lDefault returned if not found)
	Function GetSLng(Section As String, Key As String, Default As Long, File As String) As Long		
	
	'Returns a boolean value from INI file (Default returned if not found)
	Function GetSBool(Section As String, Key As String, Default As Boolean, File As String) As Boolean	
	
	'Returns a double value from INI file (Default returned if not found)
	Function GetSDbl(Section As String, Key As String, Default As Double, File As String) As Double		
	
	'Writes the value to an INI file
	Sub WriteSVar(Section As String, Key As String, Value, File As String)	
	
'----------------------------------------------------------------------------------------------------------------------------------------------
'frmHub (Form)
	'Calls the VB DoEvents function (suspends execution)
	Function DoEventsForMe() As Long		
	
	'Lists the bot name in the nicklist. If bOperator is true, then it is also added to the op list.
	'The other values are obvious; they are the values which show up in the userlist window.
	Sub RegisterBotName(sName As String, Optional bOperator As Boolean = True, _
				Optional dShare As Double, Optional sDescription As String, _
				Optional sConnection As String, Optional sEmail As String, Optional lIcon As Long)		
				
	'Removes a bot name from the nicklist (and oplist if necessary); can remove normal users also (use invisible status instead)
	Sub UnregisterBotName(sName As String)			
	
	'Closes a winsock and removes a user from the collection
	Sub CloseSocket(Index As Integer)		
	
	'Reloads settings (however scripts can only be loaded once)
	Sub LoadSettings()					
	
	'Sets the default settings
	Sub LoadDefaultSettings()				
	
	'Saves registered list, perm/temp ip bans, scripts, etc to various files
	Sub SaveSettings()			
	
	'Refreshs interface
	Sub RefreshGUI()			
	
	'Gotos next redirect IP (if there are multiple addresses)
	Sub NextRedirect()		
	
	'Stops/starts serving (depends on what it is doing at the moment)
	Sub SwitchServing()		
	
	'Show PoUp Ballon notification
		'Icon Type:
			'0 = ICON_INFO
			'1 = ICON_WARNING
			'2 = ICON_ERROR
			'3 = ICON_USER
			'4 = ICON_NONE
	Sub ShowBallon(sMsg As String, _
                       sTitle As String, _
                       Optional lIconType As Integer = 4, _
                       Optional bSound As Boolean = True)
					   
	'Lock to key (n = 5 for client-client and hub-client interactions)
	Function LockToKey(sLock As String, n As Long) As String	
	
	'Converts minutes to a string representation (ie x years, x weeks, x days, x hours and x minutes)
	Function MinToDate(ByVal lngMinutes As Long) As String			
	
	'Returns the date (same format as Now) which the hub started/stopped serving
	Property ServingDate() As Date						
	
	'Returns True if serving, False if not
	Property IsServing() As Boolean			
	
	'Returns connection to registered user database
	Property oPermaCon() As Connection			
	
'----------------------------------------------------------------------------------------------------------------------------------------------
'colUsers (clsHub)
	'Returns the number of logged in operators
	Property OpCount() As Long			
	
	'Returns the number of users
	Property Count() As Long			
	
	'Peak number of connected ops
	Property PeakOps() As Long	
	
	'Peak number of connected users
	Property PeakUsers() As Long	
	
	'The nicklist (read only)
	Property NickList() As String	
	
	'The op list (read only)
	Property OpList() As String			
	
	'Total bytes shared in the hub
	Property iTotalBytesShared() As Double		
	
	'Peak bytes shared in the hub
	Property iPeakBytesShared() As Double	
	
	'clsUser collection of users still logging in
	Property colLoggingIn() As Collection		
	
	'Checks if a user is online
	Function Online(sName As String) As Boolean			
	
	'Check if an item is in the collection by winsock index
	Function Exists(iIndex As Integer) As Boolean		
	
	'Check if a name is in the nicklist
	Function CheckList(strName As String) As Boolean	
	
	'Sends raw data to all users
	Sub SendToAll(sData As String)				
	
	'Sends a main chat message to all users
	Sub SendChatToAll(sName As String, sMessage As String)		
	
	'Mass message to all users
	Sub SendPrivateToAll(sName As String, sMessage As String)	
	
	'Sends raw data to all operators
	Sub SendToOps(sData As String)					
	
	'Sends a main chat message to all operators
	Sub SendChatToOps(sName As String, sMessage As String)		
	
	'Mass message to all ops
	Sub SendPrivateToOps(sName As String, sMessage As String)		
	
	'Sends raw data to all non-QuickList clients
	Sub SendToNQ(sData As String)				
	
	'Sends raw data to all non-away users
	Sub SendToNA(sData As String)			
	
	'Redirects all users
	Sub RedirectAll(Optional sAddress As String = DefaultRedir)		
	
	'Redirects all non-ops
	Sub RedirectNonOps(Optional sAddress As String = DefaultRedir)		
	
	'Find user's object based on IP
	Function ItemByIP(sIP As String) As clsUser			
	
	'Find user's object based on name
	Function ItemByName(sName As String) As clsUser		
	
	'Find user's object based on winsock index
	Function ItemByWinsockItem(iIndex As Integer) As clsUser	
	
'----------------------------------------------------------------------------------------------------------------------------------------------
'colRegistered (clsRegistered)
	'Adds a user to the registered user list (sAdmName = "Admin / GUI" by default - is optional)
	'(returns 1, 2 or 3 in case of error - see below)
	Function Add(sName As String, sPassword As String, iClass As enuClass, sAdmName As String) As Integer  
    
	'Removes a registered user from the list
	Sub Remove(sName As String)			
	
	'Edits are currently registered user (returns 1, 2 or 3 in case of error - see below)
	Function Edit(sName As String, sPassword As String, iClass As enuClass)		
	
	'Change the name of a registered user
	Sub Rename(sOldName As String, sNewName As String)	
	
	'Checks if the password is correct - if it is, then thier user class value is returned
	Function Check(sName As String, sPassword As String) As enuClass	
	
	'Returns 0 if a user is not registered, otherwise returns their user class value
	Function Registered(sName As String) As enuClass	
	
	'Retrieves information from the database on a user
				'--Values for sField--
		'All = UserName|Password|Class|ClassName|RegedBy|RegDate|LastLogin|LastIP|
		'ClassName = name of their class
		'Password = their password (or ban reason for banned names)
		'RegDate = date they were registered
		'RegedBy = who registered them
		'LastLogin = last time/date they logged in
		'LastIP = last IP they successfully logged in with
		'Perm = boolean of a ban is permanent (banned names only)
		'Language = language specified in account
	Function GetInfo(sName As String, Optional sField As String = "All") As String	
	
	'Retrieves a list of all users and their class (format : "<name>|<class>*" where * = new line)
		'If iClass = 0 then it gets all users, otherwise ones only of the specified class					
		'iSort values :
  			'0 = Do not sort
			'1 = Class
			'2 = Username
			'3 = Class, then Username
	'Add/Edit function returned value legend
		'0 = No error
		'1 = Already registered / Not registered
		'2 = Name too long
		'3 = Password too long
	Function GetList(ByRef iSort As Integer, Optional iClass As enuClass = 0) As String		
	
'----------------------------------------------------------------------------------------------------------------------------------------------
'colIPBans (clsIPBans)
	'Bans an ip (if lTime = -1, it is a perm ban, otherwise it is a temp ban; in minutes)
	Sub Add(sIP As String, Optional lTime As Long = -1)	
	
	'Unbans an ip, if banned
	Sub Remove(sIP As String)		
	
	'Clears the temp IP ban
	Sub ClearTemp()	
	
	'Clears the perm IP ban
	Sub ClearPerm()		
	
	'Check if an ip is banned; returns -1 is perm banned, 0 is not, otherwise is the time left in minutes for the temp ban
	Function Check(sIP As String) As Long			
	
	'List of temp ip bans (| seperator)
	Function TempList() As String	
	
	'List of perm ip bans (| seperator)
	Function PermList() As String		
	
	'Collection of temp ip bans (objects [clsTempBan])
	Property TempItems() As Collection		
	
	'Collection of perm ip bans (variants containing string ip)
	Property PermItems() As Collection		
	
'----------------------------------------------------------------------------------------------------------------------------------------------
'colCommands (clsCommands)
	'Adds a command to the collection	
	Sub Add(intIndex As Integer, strName As String, strDescription As String, _		
			intClass As enuClass, blnEnabled As Boolean)
		
	'Edits a command
	Sub Edit(strOldName As String, strNewName As String, strDescription As String, _
			intClass As enuClass, blnEnabled As Boolean)	
		
	'Removes command object from collection / listview
	Sub Remove(strKey As String)		
	
	'Clears out command collection / listview
	Sub Clear()							
	
	'Returns command object of trigger
	Function Item(strKey As String) As clsCommand		
	
	'Returns true if a command exists
	Function Exists(strKey As String) As Boolean	
	
	'Executes a command
	Sub Execute(curUser As clsUser, strTrigger As String, blnMainChat As Boolean)		
	
'----------------------------------------------------------------------------------------------------------------------------------------------
'Command object (clsCommand)
	'Min class allowed to use the command
	Property Class() enuClass	
	
	'Command ID (> 50 means it is a custom command)
	Property ID() As Integer		
	
	'Name / trigger
	Property Name() As String		
	
	'True if enabled, false if disabled
	Property Enabled() As Boolean	
	
	'Description of command
	Property Description() As String		
	
'----------------------------------------------------------------------------------------------------------------------------------------------
'Temp ban object (clsTempBan)
	'IP banned
	Property IP() As String			
	
	'The date the IP ban expires
	Property ExpDate() As Date		
	
'----------------------------------------------------------------------------------------------------------------------------------------------
'XML Parser (clsXMLParser)
	'The XML data to be parsed / that was created (based on nodes)
	Property Data() As String					
	
	'Any data left over after parsing the XML file (not true XML)
	Property Value() As String				
	
	'First level collection of nodes (clsXMLNode objects) (indexed by name where possible)
	Property Nodes() As Collection		
	
	'Parses the XML data in Data, and fills up Nodes collection
	Sub Parse()								
	
	'Creates XML data and puts in Data, based on Nodes collection
	Sub Create()		
	
'----------------------------------------------------------------------------------------------------------------------------------------------
'XML Node (clsXMLNode)
	'Name of the XML tag
	Property Name() As String		
	
	'Non-XML value between beginning and ending XML tags
	Property Value() As String		
	
	'Collection of xml nodes within this tag (clsXMLNode objects) (indexed by name where possible)
	Property Nodes() As Collection		
	
	'Collection of xml attributes for this tag (clsXMLAttribute objects (indexed by name)
	Property Attributes() As Collection		
	
'----------------------------------------------------------------------------------------------------------------------------------------------
'XML Attribute (clsXMLAttribute)
	'Name of the attribute
	Property Name() As String	
	
	'Value of the attribute
	Property Value() As String	
	
'----------------------------------------------------------------------------------------------------------------------------------------------
'Huffman compression (clsHuffman)
	'Decompresses the source file and writes it to the dest file
	Sub DecodeFile(SourceFile As String, DestFile As String)
	
	'Decompresses the text, and returns the uncompressed file text
	Function DecodeString(Text As String) As String	
	
	'Compresses the source file and writes it to the dest file
	Sub EncodeFile(SourceFile As String, DestFile As String)	
	
	'Compresses the text, and returns the compressed file text
	Function EncodeString(Text As String) As String			
	
'----------------------------------------------------------------------------------------------------------------------------------------------
'BZip2 compression (clsBZ2)
	'Decompresses the source file and writes it to the dest file
	Sub DecompressFile(sInput As String, sOutput As String)		
	
	'Decompresses the text, and returns the uncompressed file text
	Function DecompressString(sInput As String) As String		
	
	'Compresses the source file and writes it to the dest file
	Sub CompressFile(sInput As String, sOutput As String)		
	
	'Compresses the text, and returns the compressed file text
	Function CompressString(sInput As String) As String			
	
'----------------------------------------------------------------------------------------------------------------------------------------------
'ZLib compression (clsZLib)
	'Decompresses the source file and writes it to the dest file
	Sub DecompressFile(sInput As String, sOutput As String)		
	
	'Decompresses the text, and returns the uncompressed file text
	Function DecompressString(sInput As String) As String		
	
	'Compresses the source file and writes it to the dest file
	Sub CompressFile(sInput As String, sOutput As String)		
	
	'Compresses the text, and returns the compressed file tex
	Function CompressString(sInput As String) As String		
	
'----------------------------------------------------------------------------------------------------------------------------------------------
'User object (clsUser)
	'Disconnect and permanently ban a user via IP
	Sub Ban()				
	
	'Disconnect a user
	Sub Disconnect()	
	
	'Disconnects a user, and bans them via IP for x minutes
	Sub Kick(Optional iMinutes As Long)		
	
	'Redirect a user - default address is the Redirect IP
	Sub Redirect(Optional sAddress As String)	
	
	'Sends a main chat message to a user
	Sub SendChat(sName As String, sMessage As String)	
	
	'Sends raw data to a user (don't forget trailing |)
	Sub SendData(sData As String)			
	
	'Sends a private message to a user from sName
	'(if sOtherName is given, message uses sOtherName, while the PM is from sName)
	Sub SendPrivate(sName As String, sMessage As String, _			
			Optional sOtherName As String)		
			
	'Return a message from UsersMessages.xml(from a collection.User language preference specific, default to english if none.)
	'Can be use to get a hub's core/reason or script message base on it string key.Default to english if user is not registered.
	Function GetCoreMsgStr(strStringID As String) As String		
	
	'Returns their username
	Property sName() As String					
	
	'Returns their language
	Property sLanguageID() As String		
	
	'Winsock/collection index
	Property iWinsockIndex() As Integer		
	
	'Their winsock control
	Property Winsock() As Winsock		
	
	'Version sent by client (if client sent a non-numeric value for version, the value will be -1)
	Property iVersion() As Single	
	
	'Their logged in status
	Property State() As enuState	
	
	'Their user class (5 and greater are operators)
	Property Class() As enuClass	
	
	'QuickList support status
	Property QuickList() As Boolean		
	
	'NoHello support status
	Property NoHello() As Boolean		
	
	'Their MyINFO string
	Property sMyInfoString() As String		
	
	'User's IP
	Property IP() As String			
	
	'Whether they are an op or not
	Property bOperator As Boolean		
	
	'Whether the user is in away more or not
	Property isAfk() As Boolean			
	
	'Shared bytes
	Property iBytesShared() As Double		
	
	'Supports string from a client
	Property Supports() As String			
	
	'Date/time connected to hub
	Property ConnectedSince() As Date		
	
	'Whether or not a nicklist was ever requested
	Property QNL() As Boolean				
	
	'Equals False if invisible
	Property Visible() As Boolean		
	
	'Returns/sets mute status (if equals True, then cannot chat in the main chat or in pm)
	Property Mute() As Boolean			
	
	'UserCommand support status
	Property UserCommand() As Boolean		
	
	'Returns true if a NetInfo command was recieved
	Property NetInfo() As Boolean	
	
'Default available messages keys/message

   'Warning!!! Keys are case sensitive.

   'More messages can be added to UsersMessages.xml for scripts usage but is not the recommended way.
	'Add the new strings to the colLanguages collection instead, an <en> section MUST be provided.
	'The english strings are not optionnal as they are the default if user didn't chose any prefered language yet.

	InternationalName = "English" 
	NationalName = "English"
	LoggedIn = "Logged in."
	ChatMode = "Can't connect because user %USER% is in chat only mode."
	MaxHubs = "You are connected to too many hubs. %[maxhubs] hubs max. Disconnect from some and reconnect."
	MinSlots = "You do not have enough slots open. %[minslots] slot(s) min."
	MaxSlots = "You have too many slots open. %[maxslots] slots max."
	HSRatio = "You have not met the hub per slot ratio. %[hsratio] slot per hub min."
	BSRatio = "You have not met the bandwidth (in KB/s) per slot ratio (as measured by the limiter you are using) %[bsratio]KB/s per slot."
	MaxShare = "You are sharing more than maximum allowed amount. %[maxshare] max." 
	MinShare = "You have not met the minimum share. %[minshare] minimum."
	DCppMinVersion = "You are using an outdated DC++ client. Please goto http://dcplusplus.sourceforge.net/ and update it."
	NMDCMinVersion = "You are using an outdated NMDC client. Please goto http://www.neo-modus.com/ and update it. If you are using another client, please change the version setting."
	DenyNoTag = "You do not have an identification tag for your client (ie <++, <DC, etc). Please enable your tag, if possible."
	Faker = "You are suspected of trying to cheat. Goodbye."
	Socks5 = "Socks5 mode not allowed."
	PassiveMode = "Passive mode not allowed."
	PassLength = "Passwords cannot be longer than 20 characters."
	NickLength = "Your nickname cannot be longer than 40 characters."
	NickTaken = "Your nickname has already been taken."
	ChrInNick = "Your nickname has an invalid character " ' / or a (space)."
	WrongPassRedir = "The password was incorrect. You are being redirected to "
	WrongPass = "The password was incorrect."
	PassMode = "This hub is running in password mode. Please supply the global password."
	RegPass = "Your nickname is registered. Please supply the password."
	RedirectedBecause = "You are being redirected because: "
	RedirectedTo = "You are being redirected to: "
	FullRedirTo = "This hub is currently full. You are being redirected to: "
	Full = "This hub is currently full."
	RegOnlyRedirTo = "This hub is for registered users only. You are being redirected to: "
	RegOnly = "This hub is for registered users only."
	BannedBecause = "You are being banned because: " 
	IPPermBan = "Your IP is permanently banned."
	IPBanned = "Your IP is banned!"
	IPTempBan = "Your IP is temporarily banned for "
	KickedBecause = "You are being kicked because: "
	KickedBy = "The user, %USER%, was kicked by %OP%. IP: %IP%"
	IsKicking = "%OP% is kicking %USER% because: %REASON%"
'----------------------------------------------------------------------------------------------------------------------------------------------
'Settings (clsSettings) (Use this to access settings in the GUI - ex Settings.HubName)
'New Things are marked for *
	
	HubName                  As String
	HubDesc                  As String
	HubIP                    As String
	HubPassword              As String
	BotName                  As String
	OpChatName               As String
	JoinMsg                  As String
	RedirectIP               As String
	RedirectAddress          As String
'--------NEW REDIRECT ADDRESSES-------------------
	ForMinShareRedirectAddress       As String
	ForMaxShareRedirectAddress       As String
	ForMinSlotsRedirectAddress       As String
	ForMaxSlotsRedirectAddress       As String
	ForMaxHubsRedirectAddress        As String
	ForSlotPerHubRedirectAddress     As String
	ForNoTagRedirectAddress          As String
	ForTooOldDcppRedirectAddress     As String
	ForTooOldNMDCRedirectAddress     As String
	ForBWPerSlotRedirectAddress      As String
	ForFakeShareRedirectAddress      As String
	ForFakeTagRedirectAddress        As String
	ForPasModeRedirectAddress        As String
'---------------STOP HERE--------------------------
	RegisterIP               As String
	Ports                    As String
	CSeperator               As String
	MaxHubsMsg               As String
	MinSlotsMsg              As String
	MaxSlotsMsg              As String
	DCppMinVersionMsg        As String
	HSRatioMsg               As String
	BSRatioMsg               As String
	MinShareMsg              As String
	NMDCMinVersionMsg        As String
	DenyNoTagMsg             As String
	MaxShareMsg              As String
	FakeShareMsg             As String
	FakeTagMsg               As String
	MassMessage              As String
	OpMassMessage            As String
	UnRegMassMessage         As String
	Interface                As String
	Socks5Msg                As String
	PassiveModeMsg           As String
	NoCOClientsMsg           As String
	HammeringRd              As String
	NoIPDNS1                 As String
	NoIPDNS2                 As String
	NoIPDNS3                 As String
	NoIPDNS4                 As String
	NoIPUser                 As String
	NoIPPass                 As String
	DynDNS1                  As String
	DynDNS2                  As String
	DynDNS3                  As String
	DynDNS4                  As String
	DynDNSUser               As String
	DynDNSPass               As String

	DefaultBanTime           As Long
	ScriptTimeout            As Long
	FWBanLength              As Long
	Port                     As Long
	MaxMessageLen            As Long
	DataFragmentLen          As Long
'svn 216
	ConDropInterval          As Long
	FWDropMsgInterval        As Long

	DCMaxHubs                As Byte
	MaxSlots                 As Byte
	DCOSlots                 As Byte
	MinSlots                 As Byte
	MinShareSize             As Byte
	MaxShareSize             As Byte
	CPrefix                  As Byte
	DCOSpeed                 As Byte
	SendJoinMsg              As Byte
	MaxPassAttempts          As Byte
	FWGetNickList            As Byte
	FWActiveSearch           As Byte
	FWPassiveSearch          As Byte
	FWMyINFO                 As Byte
	FWMainChat               As Byte
	MinMyinfoFakeCls         As Byte

	MinPassiveSearchLen      As Integer
	FWInterval               As Integer
	MaxUsers                 As Integer
	MinSearchCls             As Integer
	MinConnectCls            As Integer
'MTU packet size is generaly 1492, TCP overhead is about 40 Bytes
'minus some bytes in case.
	ZLINELENGHT              As Integer


	IMinShare                As Double
	IMaxShare                As Double
	MinShare                 As Double
	MaxShare                 As Double
	DCSlotsPerHub            As Double
	DCBandPerSlot            As Double
	DCMinVersion             As Double
	NMDCMinVersion           As Double

	MinClsSearchSend         As Boolean
	MinClsConnectSend        As Boolean
	AutoCheckUpdate          As Boolean
	AutoKickMLDC             As Boolean
	DenySocks5               As Boolean
	DenyPassive              As Boolean

	AutoRegister             As Boolean
	AutoRedirect             As Boolean
	AutoRedirectFull         As Boolean
	AutoRedirectNonReg       As Boolean
	AutoRedirectFullNonReg   As Boolean
	AutoRedirectFullNonOps   As Boolean
	AutoStart                As Boolean
	CompactDBOnExit          As Boolean
	ConfirmExit              As Boolean
	DCValidateTags           As Boolean
	DCIncludeOPed            As Boolean
	OPBypass                 As Boolean
	PreloadWinsocks          As Boolean
	SendMessageAFK           As Boolean
	RegOnly                  As Boolean
	MentoringSystem          As Boolean
	PreventSearchBots        As Boolean
	DescriptiveBanMsg        As Boolean
	UseOpChat                As Boolean
	UseBotName               As Boolean
	Passive                  As Boolean
	RedirectFMS              As Boolean
	RedirectFGP              As Boolean
	FilterCPrefix            As Boolean
	EnabledCommands          As Boolean
	ScriptSafeMode           As Boolean
	StartMinimized           As Boolean
	SendMsgAsPrivate         As Boolean
	PasswordMode             As Boolean
	WordWrap                 As Boolean
	DenyNoTag                As Boolean
	HideFadeImg              As Boolean
	CheckFakeShare           As Boolean
	EnableFloodWall          As Boolean
	PreventGuessPass         As Boolean
	OpsCanRedirect           As Boolean
	ChatOnly                 As Boolean
	VIPUseOpChat             As Boolean
	MinimizeTray             As Boolean
'---------REDIRECT CHECK BOXES--------------------
	RedirectFTooOldDCpp      As Boolean
	RedirectFTooOldNMDC      As Boolean
	RedirectFNoTag           As Boolean
	RedirectFMinShare        As Boolean
	RedirectFMaxShare        As Boolean
	RedirectFMaxSlots        As Boolean
	RedirectFMinSlots        As Boolean
	RedirectFMaxHubs         As Boolean
	RedirectFSlotPerHub      As Boolean
	RedirectFBWPerSlot       As Boolean
	RedirectFFakeShare       As Boolean
	RedirectFFakeTag         As Boolean
	RedirectFPasMode         As Boolean
'--------------STOP IN HERE-----------------------
	HideMyinfos              As Boolean
	ACOClients               As Boolean
	DynUpdate                As Boolean
	DynDNSUpdateEna          As Boolean
	NoIPUpdateEna            As Boolean
	EnabledScheduler         As Boolean
	NoIPUpdateStartUp        As Boolean
'-------------------------------------------------

'-------------- NOTIFICATIONS --------------------
	PopUpNewReg              As Boolean
	PopUpOpConected          As Boolean
	PopUpOpDisconected       As Boolean
	PopUpUserKick            As Boolean
	PopUpUserBaned           As Boolean
	PopUpUserRedirected      As Boolean
	PopUpStartedServing      As Boolean
	PopUpStopedServing       As Boolean

' --------------------- Otheres ---------------------
	StartWin                 As Boolean
	MoveForm                 As Boolean
	PriorityBl               As Boolean
	PriorityVal              As Integer
	frmHubPosition           As String
	lngSkin                  As Long
	blSkin                   As Boolean
	RndSkin                  As Boolean
	Plugins                  As Boolean
	MagneticWin              As Boolean
'----------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------------------------
'Script Events
	'Called when script is loaded
	Sub Main()			
	
	'Called when the hub is unloading
	Sub UnloadMain()	
	
	'When data is sent by a user
	Sub DataArrival(curUser, sData)		
	
	'When the hub starts serving
	Sub StartedServing()	
	
	'When the hub stops serving
	Sub StoppedServing()	
	
	'Triggered when a user is registered (in GUI or otherwise...) using the colRegistered class
	Sub AddedRegisteredUser(sName)	
	
	'Triggered when a user is unregistered
	Sub RemovedRegisteredUser(sName)	
	
	'When an IP is perm banned (with the GUI)
	Sub AddedPermBan(sIP)			
	
	'Call when a mass message (to all or to ops) is sent by the owner (in the GUI)
	Sub MassMessage(sMessage)	
	
	'Triggered JUST BEFORE the redirects start (with a DoEvents afterwards)
	Sub StartedRedirecting()	
	
	'User left the hub
	Sub UserQuit(curUser)		
	
	'User connected to the hub (MyINFO accessable)
	Sub UserConnected(curUser)			
	
	'Op connected to the hub (MyINFO accessable)
	Sub OpConnected(curUser)		
	
	'Registered user connected to the hub (MyINFO accessable)
	Sub RegConnected(curUser)		
	
	'Connection request from sIP
	Sub AttemptedConnection(sIP)	
	
	'Timer event, when it goes off
	Sub tmrScriptTimer_Timer()		
	
	'Script winsock close event
	Sub wskScript_Close(Index)		
	
	'Script winsock connect event
	Sub wskScript_Connect(Index)	
	
	'Script winsock connectionrequest event
	Sub wskScript_ConnectionRequest(Index, requestID)	
	
	'Script winsock dataarrival event
	Sub wskScript_DataArrival(Index, bytesTotal)	
	
	'Script winsock error event
	Sub wskScript_Error(Index, Number, Description)		
	
	'When an error occurs, and "On Error Resume Next" is not used, this sub is called (use the Err object)
	Sub Error(Line)		
	
	'Raised when your script timeouts
	Sub Timeout()		
	
	'Raised when a command in the colCommand collection (and listview) which is not supported by PTDCH (and has an ID greater than 50)
	Sub CustComArrival(curUser, objCommand, strMessage, blnMainChat)	
	
	'Raised BEFORE the hub processes the data; return value is the data you want the hub to process (if empty, the hub skips processing)
	'(The last reset script gets priority on processing, then other script(s), if not empty, then to the hub.)
	Function PreDataArrival(curUser, strData)		
			
	'Notes
		'- NMDCH has two documented events which PTDCH does not have - AddedMultiHub(sMultiHubRemoteHost) and MultiHubDataChunkIn (sCurData)
		'- when UserQuit is called, the user's object has already been removed from the collection - once the sub is called, the object is destroyed
		'- OpConnected for NMDCH was for any registered user; this has been split into OpConnected and RegConnected
		'- the wskScript events are for the custom winsock, wskScript, which you can pretty much do anything with
		'- The User/Op/RegConnected events are called AFTER the hub recieves their MyINFO string (they are then considered to be logged in)
		'- NMDCH misspelled DataArival; that has been changed to DataArrival (see script converter)
		'- All raw data is recieved by DataArrival ($Key, $ValidateNick, etc)
'Other
'----------------------------------------------------------------------------------------------------------------------------------------------
'enuClass values
	Locked = -1
	Unknown = 0
	Normal = 1
	Mentored = 2
	Registered = 3
	Invisible = 4
	VIP = 5
	Op = 6
	InvisibleOp = 7
	SuperOp = 8
	InvisibleSuperOp = 9
	Admin = 10
	InvisibleAdmin = 11

'enuState values
	Wait_Key = 0
	Wait_Validate = 1
	Wait_Pass = 2
	Wait_PassPM = 3
	Wait_Info = 4
	Logged_In = 5

'enuOpenFileMode values
    vbRandom = 0
    vbInput = 1
    vbOutput = 2
	vbAppend = 3
    vbBinary = 4

'----------------------------------------------------------------------------------------------------------------------------------------------
'VB6 File access statement wrapper (clsFile)
	'Opens a file (returns Error code If any)
	Function FOpen(strPath As String, intMode As enuOpenFileMode) As Long	
	
	'Closes file (returns Error code)
	Function FClose() As Long		
	
	'Wrapper For "Print" statement (blnCR toggles the carriage Return)
	Function FPrint(strText As String, blnCR As Boolean) As Long	
	
	'Wrapper For "Write" statement (blnCR toggles the carriage Return)
	Function FWrite(strText As String, blnCR As Boolean) As Long	
	
	'Wrapper For "Put" statement
	Function FPut(varData As Variant, Optional lngNumber As Long) As Long		
	
	'Wrapper For Input Function
	Function FInput(lngLen As Long) As String		
	
	'Wrapper For "Line Input" statement
	Function FLineInput() As String		
	
	'Wrapper For "Get" statement
	Function FGet(varData As Variant, Optional lngNumber As Long) As Long	
	
	'Wrapper For "Width" statement
	Function FWidth(lngWidth As Long) As Long		
	
	'Wrapper For Lof Function
	Function FLOF() As Long				
	
	'Wrapper For FileAttr Function
	Function Attributes(intType As Integer) As Long		
	
	'Wrapper For Eof Function
	Function FEOF() As Boolean		
	
	'Wrapper For Loc Function
	Function FLOC() As Long			
	
	'Wrapper For Seek Function
	Function FSeek() As Long	
	
	'Path To file (read only)
	Property Path() As String		
	
	'Returns True If a file Is opened (read only)
	Property Opened() As Boolean		
	
	'File number which the file Is opened With (read only)
	Property FileNumber() As Integer		
	
	'Mode the file was opened With (read only)
	Property Mode() As enuOpenFileMode		
	
'----------------------------------------------------------------------------------------------------------------------------------------------
'RegExps (clsRegExps)
	'Pattern(s) must be digit(s) capturing pattern(s).
		'Return 0 If no match Or If more Then one collections
		'For Single number capture patterns.
		'Or For 3 numbers captures patterns.
		'(EX: hub count),Return = cap(1)+cap(2)[+cap(3) :optional, see DCIncludeOPed]									 
	Function CaptureDbl(strString As String, strPattern As String) As Double		

	'Return the String/substring the pattern matched.
		'Return Empty If no match Is found.
		'Note: Pattern should be a Single capturing pattern.
		'Only the first captured match of the first collection Is returned.
	Function CaptureSubStr(strString As String, strPattern As String) As String		
	
	'Return True If a match(es) Is/are found.
	'Pattern should be a none capturing pattern.
	Function TestStr(strString As String, strPattern As String) As Boolean 			
	
	'Replace ALL occurence(s) of sPattern To sReplace
	'Return the modified String
	Function REReplace(sString As String, sPattern As String, sReplace As String) As String 
	
	'This create a Matches collection based On the pattern.
		'The number of Matches collections And captured matches Is Base On the pattern given To this Function.
		'Each match can be a captured match Or an empty captured match, Base On If a capturing pattern can match something In strString.
		'It Is up To the coders To deal With the data In Matchcollection In the proper way.
	Function REMatchesCol(strString As String, strPattern As String) As MatchCollection	
	
	'Returns:  True If match a denyed expression without matching an allowed one
	Function AdvertTest(strString As String, strDeny As String, strAllow As String) As Boolean	

'----------------------------------------------------------------------------------------------------------------------------------------------
'frmScript ScriptCtrl 

	'Reolad all scripts from dir
	Sub SLoadDir()

	'intIndex parameters:
	'	0 =Save selected
	'	1 =Save by Index
	Sub SSave(Optional intIndex As Integer = 0)
	
	'Reset script by script name (file name)
	Sub SResetByName(strName As String, _
         Optional ByVal blnUpDateCode As Boolean = True, _
         Optional ByVal blnFirst As Boolean)
		 
	'Reset script.. only update the scripts to file, if no errors
	'intIndex parameters:
	'	-2 =All checked scripts
	'	-1 =All scripts
	'	-0 =Single script (selected in listview)
	'	>1 = by script Index
	Sub SReset(Optional ByVal lngSel As Long, _
               Optional ByVal blnUpDateCode As Boolean = True, _
               Optional ByVal blnFirst As Boolean)
	
	'Stopt script by name (file name)
	Sub SStopByName(strName As String)
	
	'Stop script.. 
	'intIndex parameters:
	'	-2 =All checked scripts
	'	-1 =All scripts
	'	-0 =Single script (selected in listview)
	'	>1 = by script Index	
	Sub SStop(Optional ByVal lngSel As Long)
					
'----------------------------------------------------------------------------------------------------------------------------------------------
'=========================================================== Scripting - Interface ============================================================