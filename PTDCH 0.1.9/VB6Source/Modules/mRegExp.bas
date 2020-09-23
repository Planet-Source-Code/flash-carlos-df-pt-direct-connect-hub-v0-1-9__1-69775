Attribute VB_Name = "mRegExp"
'Regular Expressions Constants

Option Explicit

'MyINFOs sample
'$MyINFO $ALL JIM <++ V:0.674,M:5,H:255/0/1,S:2>$ $DSL$$57270215163$
'$MyINFO $ALL Femme2 FT<++ V:0.304,M:P,H:4,S:1>$ $DSL$ft$11359870483$
'$MyINFO $ALL test <DCDM 0.0485 svn338><++ V:0.401,M:A,H:1/0/1,S:3,R:1>$ $28.8Kbps$$0$
'$MyINFO $ALL GhOst [Asgardian][Everybody Dances with the Grim Reaper]<AsG++ V:0.062,M:A,H:1/4/32,S:3>$ $Pirate$$16932719403$


'Regular expressions constants
'Function CaptureSubStr, The regular expressions MUST capture only ONE SubMatches
Public Const GETNICK            As String = "\$A[lL][lL]\s([\S]{1,40})"
Public Const GETDCMODE          As String = ",M:([AP5]),"
Public Const GETCONTYPE         As String = "\$[ ]\$(.{3,16})[\x01\x02\x03\x04\x05\x06\x07\x08\x09\x0A\x0B\x0C\x0D\x0E\x0F]"
Public Const GETSTATUS          As String = "([\x01\x02\x03\x04\x05\x06\x07\x08\x09\x0A\x0B\x0C\x0D\x0E\x0F])\$"

Public Const GETFROMNICKINPM    As String = "From:[ ]([^ ]{1,40})"
Public Const GETNICKINPMMSG     As String = "[ ]\$<([^>]{1,40})"
'Public Const MYINFOCAPTURE     As String = "\$MyINFO[ ]\$ALL[ ]([^ ]{1,40})[ ](?:.{1,40})?<([^ ]{2,10})[ ]V:(\d\.\d{1,4})[^,]*,M:([AP5]),H:(\d{1,3}/\d{1,3}/\d{1,3}),S:(\d{1,3})(?:,(?:O:|R:|B:|L:|U:|B)(\d{1,3}))?>\$[ ]\$([^\$]{3,12})([\x01\x02\x03\x04\x05\x06\x07\x08\x09\x0A\x0B\x0C\x0D\x0E\x0F])\$([^\$]{0,40})\$(\d{1,16})\$$"

'Function CaptureDbl, MUST be Numerical SubMatches capture only.
'The 3 numerical SubMatches are added together.
Public Const GETHUBCOUNT        As String = "H:(\d{1,3})/(\d{1,3})/(\d{1,3})|H:(\d{1,3}),"
Public Const GETSHARESIZE       As String = "\$(\d{1,15})\$$"
Public Const GETVERSION         As String = "V:(\d\.\d{1,4})"
'GETSLOTS does not accept S:* it will return 0
'if slot check is disabled S:0 won't be consider as cheating.
'if enable it will return 0 in both case S:0 and S:*
Public Const GETSLOTS           As String = "S:(\d{1,3})"

'Function TestStr, Do not use capture parenthesis if it can be avoid
'Or use (?:pattern) to avoid capturing when doing atomic grouping.
'Public Const VALIDATEDCCMDS     As String = "\$ConnectToMe\s|\$RevConnnectToMe\s|\$Search\s|\$\MyINFO\s|\$GetINFO\s|\$To\s|\$<[\S]{1,40}>\s|\$GetNickList\s|\$ValidateNick\s|\$MyPass\s|\$Version\s|\$Key\s|\$SR\s"
Public Const CHRSTODENYINNICK   As String = """|'|/|\s|\$"

'almost impossible share sizes, sizes often use by dumb fakers or bots.
Public Const DENYSHARESIZE    As String = "0{6}|1{6}|2{6}|3{6}|4{6}|5{6}|6{6}|7{6}|8{6}|9{6}|098765|987654|876543|765432|654321|543210|123456|234567|345678|456789|567890"
