VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public HubName                  As String
Public HubDesc                  As String
Public HubIP                    As String
Public HubPassword              As String
Public BotName                  As String
Public OpChatName               As String
Public JoinMsg                  As String
Public RedirectIP               As String
Public RedirectAddress          As String
'--------NEW REDIRECT ADDRESSES-------------------
Public ForMinShareRedirectAddress       As String
Public ForMaxShareRedirectAddress       As String
Public ForMinSlotsRedirectAddress       As String
Public ForMaxSlotsRedirectAddress       As String
Public ForMaxHubsRedirectAddress        As String
Public ForSlotPerHubRedirectAddress     As String
Public ForNoTagRedirectAddress          As String
Public ForTooOldDcppRedirectAddress     As String
Public ForTooOldNMDCRedirectAddress     As String
Public ForBWPerSlotRedirectAddress      As String
Public ForFakeShareRedirectAddress      As String
Public ForFakeTagRedirectAddress        As String
Public ForPasModeRedirectAddress        As String
'---------------STOP HERE--------------------------
Public RegisterIP               As String
Public Ports                    As String
Public CSeperator               As String
Public MaxHubsMsg               As String
Public MinSlotsMsg              As String
Public MaxSlotsMsg              As String
Public DCppMinVersionMsg        As String
Public HSRatioMsg               As String
Public BSRatioMsg               As String
Public MinShareMsg              As String
Public NMDCMinVersionMsg        As String
Public DenyNoTagMsg             As String
Public MaxShareMsg              As String
Public FakeShareMsg             As String
Public FakeTagMsg               As String
Public MassMessage              As String
Public OpMassMessage            As String
Public UnRegMassMessage         As String
Public Interface                As String
Public Socks5Msg                As String
Public PassiveModeMsg           As String
Public NoCOClientsMsg           As String
Public HammeringRd              As String
Public NoIPDNS1                 As String
Public NoIPDNS2                 As String
Public NoIPDNS3                 As String
Public NoIPDNS4                 As String
Public NoIPUser                 As String
Public NoIPPass                 As String
Public DynDNS1                  As String
Public DynDNS2                  As String
Public DynDNS3                  As String
Public DynDNS4                  As String
Public DynDNSUser               As String
Public DynDNSPass               As String

Public DefaultBanTime           As Long
Public ScriptTimeout            As Long
Public FWBanLength              As Long
Public Port                     As Long
Public MaxMessageLen            As Long
Public DataFragmentLen          As Long
'svn 216
Public ConDropInterval          As Long
Public FWDropMsgInterval        As Long

Public DCMaxHubs                As Byte
Public MaxSlots                 As Byte
Public DCOSlots                 As Byte
Public MinSlots                 As Byte
Public MinShareSize             As Byte
Public MaxShareSize             As Byte
Public CPrefix                  As Byte
Public DCOSpeed                 As Byte
Public SendJoinMsg              As Byte
Public MaxPassAttempts          As Byte
Public FWGetNickList            As Byte
Public FWActiveSearch           As Byte
Public FWPassiveSearch          As Byte
Public FWMyINFO                 As Byte
Public FWMainChat               As Byte
Public MinMyinfoFakeCls         As Byte

Public MinPassiveSearchLen      As Integer
Public FWInterval               As Integer
Public MaxUsers                 As Integer
Public MinSearchCls             As Integer
Public MinConnectCls            As Integer
'MTU packet size is generaly 1492, TCP overhead is about 40 Bytes
'minus some bytes in case.
Public ZLINELENGHT              As Integer


Public IMinShare                As Double
Public IMaxShare                As Double
Public MinShare                 As Double
Public MaxShare                 As Double
Public DCSlotsPerHub            As Double
Public DCBandPerSlot            As Double
Public DCMinVersion             As Double
Public NMDCMinVersion           As Double

Public MinClsSearchSend         As Boolean
Public MinClsConnectSend        As Boolean
Public AutoCheckUpdate          As Boolean
Public AutoKickMLDC             As Boolean
Public DenySocks5               As Boolean
Public DenyPassive              As Boolean

Public AutoRegister             As Boolean
Public AutoRedirect             As Boolean
Public AutoRedirectFull         As Boolean
Public AutoRedirectNonReg       As Boolean
Public AutoRedirectFullNonReg   As Boolean
Public AutoRedirectFullNonOps   As Boolean
Public AutoStart                As Boolean
Public CompactDBOnExit          As Boolean
Public ConfirmExit              As Boolean
Public DCValidateTags           As Boolean
Public DCIncludeOPed            As Boolean
Public OPBypass                 As Boolean
Public PreloadWinsocks          As Boolean
Public SendMessageAFK           As Boolean
Public RegOnly                  As Boolean
Public MentoringSystem          As Boolean
Public PreventSearchBots        As Boolean
Public DescriptiveBanMsg        As Boolean
Public UseOpChat                As Boolean
Public UseBotName               As Boolean
Public Passive                  As Boolean
Public RedirectFMS              As Boolean
Public RedirectFGP              As Boolean
Public FilterCPrefix            As Boolean
Public EnabledCommands          As Boolean
Public ScriptSafeMode           As Boolean
Public StartMinimized           As Boolean
Public SendMsgAsPrivate         As Boolean
Public PasswordMode             As Boolean
Public WordWrap                 As Boolean
Public DenyNoTag                As Boolean
Public HideFadeImg              As Boolean
Public CheckFakeShare           As Boolean
Public EnableFloodWall          As Boolean
Public PreventGuessPass         As Boolean
Public OpsCanRedirect           As Boolean
Public ChatOnly                 As Boolean
Public VIPUseOpChat             As Boolean
Public MinimizeTray             As Boolean
'---------REDIRECT CHECK BOXES--------------------
Public RedirectFTooOldDCpp      As Boolean
Public RedirectFTooOldNMDC      As Boolean
Public RedirectFNoTag           As Boolean
Public RedirectFMinShare        As Boolean
Public RedirectFMaxShare        As Boolean
Public RedirectFMaxSlots        As Boolean
Public RedirectFMinSlots        As Boolean
Public RedirectFMaxHubs         As Boolean
Public RedirectFSlotPerHub      As Boolean
Public RedirectFBWPerSlot       As Boolean
Public RedirectFFakeShare       As Boolean
Public RedirectFFakeTag         As Boolean
Public RedirectFPasMode         As Boolean
'--------------STOP IN HERE-----------------------
Public HideMyinfos              As Boolean
Public ACOClients               As Boolean
Public DynUpdate                As Boolean
Public DynDNSUpdateEna          As Boolean
Public NoIPUpdateEna            As Boolean
Public EnabledScheduler         As Boolean
Public NoIPUpdateStartUp        As Boolean
'-------------------------------------------------

'#################################################

'-------------- NOTIFICATIONS --------------------
Public PopUpNewReg              As Boolean
Public PopUpOpConected          As Boolean
Public PopUpOpDisconected       As Boolean
Public PopUpUserKick            As Boolean
Public PopUpUserBaned           As Boolean
Public PopUpUserRedirected      As Boolean
Public PopUpStartedServing      As Boolean
Public PopUpStopedServing       As Boolean

' --------------------- Otheres ---------------------
Public StartWin                 As Boolean
Public MoveForm                 As Boolean
Public PriorityBl               As Boolean
Public PriorityVal              As Integer
Public frmHubPosition           As String
Public lngSkin                  As Long
Public blSkin                   As Boolean
Public RndSkin                  As Boolean
Public Plugins                  As Boolean
Public MagneticWin              As Boolean
'----------------------------------------------------
