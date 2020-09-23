Attribute VB_Name = "mSubMain"
Option Explicit

'API calls
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long

Private mXP     As clsXPTheme
Private mPlgins As clsPlugins

Sub Main()
     'Turn off nasty error messages which might lead to crashing (b/c of API calls)
2:   SetErrorMode &H1 Or &H2

4:   If App.PrevInstance Then frmLoading.Show: _
        frmLoading.lblIsRunning.Visible = True: _
        MsgBeep beepSystemExclamation: Pause 3000: End
        
8:   Dim i As Integer
9:   Set g_objFileAccess = New clsFileAccess
10:  Set mXP = New clsXPTheme
11:  Set mPlgins = New clsPlugins

13:  On Error GoTo Err

15:  Call CheckFiles
16:  Call CheckDirs
17:  Call CheckDLLs
     
     'Set app for XP-style, if available.. use the XML manifest in resource ;-)
20:  Call mXP.InitializeXP
      
24:  frmLoading.Show

     'Inicialize Hub conf...
27:  Call Load(frmHub)

     'Set List View styles.. I use this code, because that controls usaded in this project MS Commom Controls SP 5 ;-)
     'this code extend funcionality in List View

32:  Call LVFullRow

34:  Call SetFlatBorder 'Set flat style in the pictureboxs

37:    With frmHub

            'Set caption to proper format
40:         .Caption = "PT Direct Connect Hub " & vbVersion & vbSVNVersion
              
            'skin ///////////////////////////////////////////////////////////////
            'add themes to combobox
45:         .cmbSkin.AddItem "01-Defaut"
46:         .cmbSkin.AddItem "02-Cyan Blue"
47:         .cmbSkin.AddItem "03-Cyan Green"
48:         .cmbSkin.AddItem "04-Metallic"
49:         .cmbSkin.AddItem "05-Metallic Blue"
50:         .cmbSkin.AddItem "06-Metallic Green"
51:         .cmbSkin.AddItem "07-Metallic Navy Blue"
52:         .cmbSkin.AddItem "08-Metallic Oliver"
53:         .cmbSkin.AddItem "09-Texture Grain"
54:         .cmbSkin.AddItem "10-Texture Spater"
55:         .cmbSkin.AddItem "11-Texture Tiles"
56:         .cmbSkin.AddItem "12-Texture Toxedo"
57:         .cmbSkin.AddItem "13-Blue Berry"
58:         .cmbSkin.AddItem "14-Glace Table"
59:         .cmbSkin.AddItem "15-Pink"
60:         .cmbSkin.AddItem "16-Gun Blue"
61:         .cmbSkin.AddItem "17-Gun Metal"
        
            ' if checkbox is checked Randomize skin
64:         If g_objSettings.blSkin Then
65:            If g_objSettings.RndSkin And g_objSettings.blSkin Then
66:               Randomize
67:               g_objSettings.lngSkin = CInt((16) * Rnd + 1)
68:            End If
69:         End If
            ' set combobox text
71:         Select Case g_objSettings.lngSkin
               Case 1: .cmbSkin.Text = "01-Defaut"
               Case 2: .cmbSkin.Text = "02-Cyan Blue"
               Case 3: .cmbSkin.Text = "03-Cyan Green"
               Case 4: .cmbSkin.Text = "04-Metallic"
               Case 5: .cmbSkin.Text = "05-Metallic Blue"
               Case 6: .cmbSkin.Text = "06-Metallic Green"
               Case 7: .cmbSkin.Text = "07-Metallic Navy Blue"
               Case 8: .cmbSkin.Text = "08-Metallic Oliver"
               Case 9: .cmbSkin.Text = "09-Texture Grain"
               Case 10: .cmbSkin.Text = "10-Texture Spater"
               Case 11: .cmbSkin.Text = "11-Texture Tiles"
               Case 12: .cmbSkin.Text = "12-Texture Toxedo"
               Case 13: .cmbSkin.Text = "13-Blue Berry"
               Case 14: .cmbSkin.Text = "14-Glace Table"
               Case 15: .cmbSkin.Text = "15-Pink"
               Case 16: .cmbSkin.Text = "16-Gun Blue"
               Case 17: .cmbSkin.Text = "17-Gun Metal"
            End Select
            'END Skin \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
             
            'Combo Boxs
76:         .cmbRegistered.AddItem "All Classes"
77:         .cmbRegistered.AddItem "Non-OPs only"
78:         .cmbRegistered.AddItem "OPs and above"
79:         .cmbRegistered.AddItem "Admins and above"
80:         .cmbRegistered.Text = "All Classes"

            'Load notepad text from the file
83:        If g_objFileAccess.FileExists(G_APPPATH & "\Settings\notepad.txt") Then _
                 .txtNotePad.Text = g_objFileAccess.ReadFile(G_APPPATH & "\Settings\notepad.txt")
            
            
            'Load Plugins if then
88:        If g_objSettings.Plugins Then
89:           Call mPlgins.InstallPlugins
90:        Else
91:           .lvwPlugins.Enabled = False
92:        End If

            'Load Scripts
95:        With frmScript
96:           Call Load(frmScript)
97:           .SLoadDir
98:           .XmlBooleanLoad
99:           .SReset -2, False, False
100:       End With
            
            'set text bot name in the tab Status
103:       .txtStForm.Text = g_objSettings.BotName
            
            'set Processor Info in the tab Misc
106:       .txtSystem(6).Text = ProcessorInfo
            
            'Load pictures from the resource file
109:        .picLog(0).Picture = iResPic(102)
110:        .picLog(1).Picture = iResPic(103)
111:        .picLog(2).Picture = iResPic(104)
            
113:        Unload frmLoading
           
            'Set dimension
116:       .Width = g_WinMinW
117:       .Height = g_WinMinH

            'restore form in the last windows position ..
120:        RestoreFormSize
            
122:        If g_objSettings.StartMinimized Then
123:           .Show
124:           .WindowState = vbMinimized
125:        Else
126:           .Show
127:        End If

129:       .RefreshGUI True

131:    End With

133:    Call AboutTxt
134:    Call DelPopUpMenu
        
        'check for updates if then
137:    If g_objSettings.AutoCheckUpdate Then Call frmUpDate.Notific(True)

139:    Set mXP = Nothing
140:    Set mPlgins = Nothing
            
142:   Exit Sub
143:
Err:
145:   HandleError Err.Number, Err.Description, Erl & "|" & "mSubMain.Main()"
146:   Resume Next
End Sub

Private Sub CheckFiles()

2:   On Error GoTo Err
     Dim strTemp(2) As String
     Dim i As Integer
     
4:    If Not g_objFileAccess.FileExists(App.Path & "\DBs\userdb.mdb") Then
5:       MsgBox "No database (userdb.mdb) found; closing hub", vbOKOnly Or vbCritical, "PTDCH"
6:       End 'Hard end
7:    End If

9:    strTemp(0) = App.Path & "\Settings\VB.bin"
10:   strTemp(1) = App.Path & "\Settings\JScripts.bin"
11:   strTemp(2) = App.Path & "\Settings\sqL.bin"
      
13:   For i = 0 To 2
14:      If Not (g_objFileAccess.FileExists(strTemp(i))) Then _
                MsgBox "HighLihter '" & strTemp(i) & "' file not found!", vbCritical
16:   Next
      
18:    If Not g_objFileAccess.FileExists(App.Path & "\Settings\UsersMessages.xml") Then
          'create defaut users massages
20:       CreateUsersMessagesXML
21:   End If

23:   Exit Sub
24:
Err:
26:    HandleError Err.Number, Err.Description, Erl & "|" & "mSubMain.CheckDirs()"
End Sub

Private Sub CheckDirs()

     'Make sure we don't loose owners previous hub dirs
3:   On Error GoTo Err

5:      Dim i As Integer
6:      Dim sPath(6) As String
      
8:      sPath(0) = App.Path & "\Logs"
9:      sPath(2) = App.Path & "\DBs"
10:      sPath(3) = App.Path & "\Settings"
11:      sPath(4) = App.Path & "\Scripts"
12:      sPath(5) = App.Path & "\Plugins"
13:      sPath(6) = App.Path & "\Languages"
       
15:      For i = 0 To 6
16:        If Not (g_objFileAccess.FileExists(sPath(i))) Then _
                g_objFileAccess.CreateDir sPath(i)
18:      Next i

20:   Exit Sub
21:
Err:
22:    HandleError Err.Number, Err.Description, Erl & "|" & "mSubMain.CheckDirs()"
End Sub

Private Sub CheckDLLs()
1:   On Error GoTo Err

3:   Dim strDll(3) As String
4:   Dim i As Integer
     
6:   strDll(0) = G_APPPATH & "\libbz2.dll"
7:   strDll(1) = G_APPPATH & "\MyIPTools.DLL"
8:   strDll(2) = G_APPPATH & "\zlib.dll"
9:   strDll(3) = G_APPPATH & "\SciLexer.dll"
     
11:  For i = 0 To 3
12:        If Not (g_objFileAccess.FileExists(strDll(i))) Then _
                MsgBox "Failed to initialize the '" & strDll(i) & "' interface." & vbNewLine & vbNewLine & _
                       "Please verify that '" & strDll(i) & "' is in the program" & vbNewLine & _
                       "directory or the system32 directory.", vbCritical
16:  Next
     
18:  Exit Sub
19:
Err:
21:  HandleError Err.Number, Err.Description, Erl & "|" & "mSubMain.CheckDLLs()"
End Sub

Public Sub AddTxtAbout(strMsg As String, Optional intColor As Integer, Optional bBold As Boolean = False, Optional bUnderline As Boolean = False)
1:   With frmHub
2:        .rtbAbout.SelStart = Len(.rtbAbout)
3:        .rtbAbout.SelColor = QBColor(intColor) 'In color!
4:        If bBold Then .rtbAbout.SelBold = True Else .rtbAbout.SelBold = False
5:        If bUnderline Then .rtbAbout.SelUnderline = True Else .rtbAbout.SelUnderline = False
6:        .rtbAbout.SelAlignment = vbCenter
7:        .rtbAbout.SelText = strMsg
8:  End With
End Sub

Private Sub AboutTxt()

2:     On Error GoTo Err

4:     Dim t As String

6:     Const sTxtLine As String = "--------------------------------------------------------------------------------------------"
    
8:     AddTxtAbout sTxtLine & vbNewLine, 8, True
9:     AddTxtAbout "PTDCH V:" & vbVersion & " is:" & vbNewLine, 3, True, True
10:    AddTxtAbout "-Created by fLaSh" & vbNewLine
11:    AddTxtAbout "-Programmed in MS Visual Basic 6" & vbNewLine
12:    AddTxtAbout "-Free and 100% open source" & vbNewLine
13:    AddTxtAbout "-Licenced under GPL" & vbNewLine
14:    AddTxtAbout "-Based in DDCH" & vbNewLine
15:    AddTxtAbout sTxtLine & vbNewLine, 8, True
16:    AddTxtAbout "Author Info" & vbNewLine, 3, True, True
17:    AddTxtAbout "My name is Carlos Ferreira" & vbNewLine
18:    AddTxtAbout "My Contact:" & vbNewLine
19:    AddTxtAbout "-E-mail: Carlosferreiracarlos@hotmail.com" & vbNewLine
20:    AddTxtAbout "-Phone: 966 506 396" & vbNewLine
21:    AddTxtAbout "-Home: Braga, S. Victor - Portugal" & vbNewLine
22:    AddTxtAbout sTxtLine & vbNewLine, 8, True
23:    AddTxtAbout "Release:" & vbNewLine
24:    AddTxtAbout vbReleaseDate & vbNewLine
25:    AddTxtAbout "HomePage:" & vbNewLine
26:    AddTxtAbout "http://HublistChecker.pt.vu/" & vbNewLine
27:    AddTxtAbout sTxtLine & vbNewLine, 8, True
28:    AddTxtAbout "Comment" & vbNewLine, 3, True, True
29:    AddTxtAbout "PTDCH is a server-software for the Direct Connect P2P Network." & vbNewLine
30:    AddTxtAbout "My goal when creating PTDCH was beside making a cool server software, actually for my part mostly about learning. The core idea was making it an easy task to setup a hub with the most common features. You might say that it was supposed to be a hub-in-a-box. This turned out to be a lot of work and frankly I was neither prepared nor experienced enough to fully complete the task." & vbNewLine
31:    AddTxtAbout "PTDCH uses highly optimized code but still not to the extent that makes it hard to read, wich is what crippled parts of the PTDCH code." & vbNewLine
32:    AddTxtAbout sTxtLine & vbNewLine, 8, True
33:    AddTxtAbout "Thanks and my Respect to creatores of DDCH:" & vbNewLine, 3, True, True
34:    AddTxtAbout "The Left Hand, ButterflySoul, HaArD and Selyb." & vbNewLine
35:    AddTxtAbout sTxtLine & vbNewLine, 8, True
36:    AddTxtAbout "If you find any bugs, or have any comments about this Hub Software, please mail me." & vbNewLine
37:    AddTxtAbout sTxtLine & vbNewLine, 8, True
38:    AddTxtAbout "Thank you for using this HubSoft :)" & vbNewLine
39:    AddTxtAbout sTxtLine & vbNewLine, 8, True
40:    AddTxtAbout "Oficial DC Hub Address" & vbNewLine
41:    AddTxtAbout "ptdch.no-ip.org" & vbNewLine
42:    AddTxtAbout sTxtLine & vbNewLine, 8, True
43:    AddTxtAbout "Regards," & vbNewLine
44:    AddTxtAbout "fLaSh - Carlos D.F." & vbNewLine
45:    AddTxtAbout "Braga (S.Victor) - Portugal"
    
46:    frmHub.rtbAbout.SelStart = 0
    
48:    Exit Sub
49:
Err:
50:    HandleError Err.Number, Err.Description, Erl & "|" & "mSubMain.AboutTxt()"
End Sub

Public Sub LoadDfsMessages()

2:     On Error GoTo Err
   
       'pre-define Strings LANGUAGE
5:     g_colMessages.Add "msgMissingStr", "This string is still missing in language file English.xml. Please contact DDCH-Team on http://www.shadowdc.com/forums"
6:     g_colMessages.Add "msgYourIP", "Your IP is %[ip] . Copy to clipboard?"
7:     g_colMessages.Add "msgExitDDCH", "Press Yes to confirm exiting DDCH."
8:     g_colMessages.Add "msgUpdating", "Already in progress of downloading an update."
9:     g_colMessages.Add "msgRedirAll", "Redirect all users including operators?"
10:    g_colMessages.Add "msgGettingIP", "Already attempting to determine IP."
11:    g_colMessages.Add "msgMostRecent", "You have the most recent version of DDCH."
12:    g_colMessages.Add "msgDownload", "Do you wish to download it?"
13:    g_colMessages.Add "msgInvalidBanName", "You cannot ban names longer than 40 characters."
14:    g_colMessages.Add "msgClearPermIPs", "Press 'Yes' to confirm clearing the permanent IP ban list."
15:    g_colMessages.Add "msgAlreadyAdded", " has already been added to the list."
16:    g_colMessages.Add "msgInvalidBanLength", "The ban length must be numeric"
17:    g_colMessages.Add "msgClearTempIPs", "Press Yes to confirm clearing the temporary IP ban list."
18:    g_colMessages.Add "msgAlreadyRegged", " is already registered."
19:    g_colMessages.Add "msgInvalidRegName", "Registered names cannot be longer than 40 characters."
20:    g_colMessages.Add "msgInvalidPass", "Passwords cannot be longer than 20 characters."
21:    g_colMessages.Add "msgInvalidClass", "Invalid class."
22:    g_colMessages.Add "msgNotRegged", " is not registered."
23:    g_colMessages.Add "msgPortInUse", "Port %[port] is already in use."
24:    g_colMessages.Add "msgEnterRedirUsersAddress", "Enter the address to redirect the users to"
25:    g_colMessages.Add "msgEnterPM", "Enter private message to send to all users"
26:    g_colMessages.Add "msgEnterOpPM", "Enter private message to send to all operators"
27:    g_colMessages.Add "msgEnterBanName", "Enter the name to ban"
28:    g_colMessages.Add "msgEnterBanReason", "Enter the reason why you're banning the name (optional)"
29:    g_colMessages.Add "msgEnterReplace", "Enter the name to replace "
30:    g_colMessages.Add "msgEnterPermIP", "Enter the IP to permanently ban."
31:    g_colMessages.Add "msgEnterRemIP", "Enter the IP to remove"
32:    g_colMessages.Add "msgEnterDataToSel", "Enter the data to send to the selected users"
33:    g_colMessages.Add "msgEnterDataToAll", "Enter the data to send to all users"
34:    g_colMessages.Add "msgEnterLength", "Enter the length of the ban (in minutes)"
35:    g_colMessages.Add "msgKickReason", "Reason for kick"
36:    g_colMessages.Add "msgEnterRedirAddress", "Enter the address to redirect to"
37:    g_colMessages.Add "msgRedirReason", "Reason for redirect"
38:    g_colMessages.Add "msgBanReason", "Reason for ban"
39:    g_colMessages.Add "msgEnterTag", "Enter the tag to add"
40:    g_colMessages.Add "msgEnterTempIP", "Enter the IP to temporarily ban."
41:    g_colMessages.Add "msgEnterBanLength", "Enter the length in minutes to ban the IP."
42:    g_colMessages.Add "msgRenameBan", "Rename Ban"
43:    g_colMessages.Add "msgEnterRegName", "Enter the name you want to register"
44:    g_colMessages.Add "msgEnterPass", "Enter the password for "
45:    g_colMessages.Add "msgEnterClass", "Enter the class for "
46:    g_colMessages.Add "msgEnterNewPass", "Enter the new password for "
47:    g_colMessages.Add "msgEnterNewClass", "Enter the new class for "
48:    g_colMessages.Add "msgEnterNewName", "Enter the new name for "
49:    g_colMessages.Add "msgConfirmExit", "Confirm Exit"
50:    g_colMessages.Add "msgUpdate", "Update in progress"
51:    g_colMessages.Add "msgRedirUsers", "Redirect users"
52:    g_colMessages.Add "msgDetectIP", "Detect IP"
53:    g_colMessages.Add "msgNoUpdate", "No update avaliable"
54:    g_colMessages.Add "msgBanName", "Ban Name"
55:    g_colMessages.Add "msgConfirmClear", "Confirm clear"
56:    g_colMessages.Add "msgBanTempIP", "Ban Temporary IP"
57:    g_colMessages.Add "msgRegUser", "Register user"
58:    g_colMessages.Add "msgEditRegged", "Edit registered user"
59:    g_colMessages.Add "msgStartServing", "Start serving"
60:    g_colMessages.Add "msgMassMsg", "Mass Message"
61:    g_colMessages.Add "msgMassMsgOp", "Op Mass Message"
62:    g_colMessages.Add "msgMassMsgUnReg", "UnReg Mass Message"
63:    g_colMessages.Add "msgBanPermIP", "Ban Permanent IP"
64:    g_colMessages.Add "msgRemoveIP", "Remove IP"
65:    g_colMessages.Add "msgSendToSel", "Send data (selected)"
66:    g_colMessages.Add "msgSendToAll", "Send data (all)"
67:    g_colMessages.Add "msgKickSel", "Kick (selected)"
68:    g_colMessages.Add "msgRedirSel", "Redirect (selected)"
69:    g_colMessages.Add "msgBan", "Ban"
70:    g_colMessages.Add "msgAddTag", "Add tag"
71:    g_colMessages.Add "msgRenameUser", "Rename user"
72:    g_colMessages.Add "msgKick", "Kick"
73:    g_colMessages.Add "msgRedir", "Redirect"
74:    g_colMessages.Add "msgIPError", "An error occured while trying to retrieve your IP from www.whatismyip.org. Your IP, as can be determined locally, is %[ip]. Copy to clipboard?"
75:    g_colMessages.Add "msgDownloadError", "An error occured while downloading the update (%[number]: %[description])."
76:    g_colMessages.Add "msgUpdateError", "Update Error"
77:    g_colMessages.Add "msgIPNotValide", " IP addresse is not valide."
78:    g_colMessages.Add "msgDays", "Day(s):"
79:    g_colMessages.Add "msgHours", "Hour(s):"
80:    g_colMessages.Add "msgMinutes", "Minute(s):"
       'Defaut forms buttons -----------------------------------------------------
81:    g_colMessages.Add "msgClose", "Close"
82:    g_colMessages.Add "msgCancel", "Cancel"
83:    g_colMessages.Add "msgOK", "OK"
84:    g_colMessages.Add "msgAdd", "Add"
85:    g_colMessages.Add "msgRemame", "Rename"
86:    g_colMessages.Add "msgEdit", "Edit"
87:    g_colMessages.Add "msgClipboard", "Clipboard"
       ' Strings for frmCAccounts
89:    g_colMessages.Add "msgConvertRegs", "Convert Accounts for PTDCH database"
90:    g_colMessages.Add "msgConvRegsDBType", "Select the type of database:"
91:    g_colMessages.Add "msgConvRegsNoErr", "No Errors"
92:    g_colMessages.Add "msgConvRegsWithErr", "With Errors"
93:    g_colMessages.Add "msgConvRegsCount", "Accounts Count:"
94:    g_colMessages.Add "msgConvRegsDir", "Select directory of the "
95:    g_colMessages.Add "msgConvRegsNoXML", "XML file not found! "
96:    g_colMessages.Add "msgConvRegsBrowse", "Browse"
97:    g_colMessages.Add "msgConvRegsConv", "Convert"
98:    g_colMessages.Add "msgConvAccountN", "Account Nº"
99:    g_colMessages.Add "msgConvName", "Name"
100:   g_colMessages.Add "msgConvPassword", "Password"
101:   g_colMessages.Add "msgConvProfile", "Profile"
102:   g_colMessages.Add "msgConvErr", "Error Description"
       ' Strings for frmCommand
104:   g_colMessages.Add "msgCommand", "Edit Command"
105:   g_colMessages.Add "msgCmdEnabled", "Enabled"
106:   g_colMessages.Add "msgCmdTrigger", "Trigger"
107:   g_colMessages.Add "msgCmdMinClas", "Minimum class"
       'Strings for frmNewScript
109:   g_colMessages.Add "msgNewScript", "New Script"
110:   g_colMessages.Add "msgNewScriptName", "Enter the name of the script:"
111:   g_colMessages.Add "msgNewScriptType", "Select script type:"
112:   g_colMessages.Add "msgScriptAlready", " is a name already in use by another script."
       '
114:   g_colMessages.Add "msgRegAdd", "Register add at: "
115:   g_colMessages.Add "msgRegUpdate", "Register updated at: "
       'Strings for frmEditScintilla
117:   g_colMessages.Add "msgSCIFind", "Find"
118:   g_colMessages.Add "msgSCIReplace", "Replace"
119:   g_colMessages.Add "msgSCIGoTo", "GoTo"
120:   g_colMessages.Add "msgSCIFindNext", "Find Next"
121:   g_colMessages.Add "msgSCIReplace", "Replace"
122:   g_colMessages.Add "msgSCIReplaceAll", "Replace All"
123:   g_colMessages.Add "msgSCIFindPrev", "Find Previous"
124:   g_colMessages.Add "msgSCIGo", "Go"
125:   g_colMessages.Add "msgSCIWrap", "Wrap around"
126:   g_colMessages.Add "msgSCIWhole", "Match whole word only"
127:   g_colMessages.Add "msgSCICase", "Match case"
128:   g_colMessages.Add "msgSCIRegExp", "Regular expression"
129:   g_colMessages.Add "msgSCIFindWhat", "Find what:"
130:   g_colMessages.Add "msgSCIReplWith:", "Replace with:"
131:   g_colMessages.Add "msgSCIDestLine", "Destination Line:"
132:   g_colMessages.Add "msgSCICurrLine", "Current Line: "
133:   g_colMessages.Add "msgSCILastLine", "Last Line:"
134:   g_colMessages.Add "msgSCIColumn", "Column:"
135:   g_colMessages.Add "msgSCIReplTimes", "Replaced %[times] times"

137:  Exit Sub

139:
Err:
141:    HandleError Err.Number, Err.Description, Erl & "|" & "mSubMain.LoadDfsMessages()"
End Sub

Public Sub LoadDfsSettings()

2:    On Error GoTo Err

4:    g_objSettings.HubName = "PT DC Hub Demo"
5:    g_objSettings.HubDesc = "[PTDCH " & vbVersion & "]"
6:    g_objSettings.HubIP = "127.0.0.1"
7:    g_objSettings.BotName = "Security"
8:    g_objSettings.OpChatName = "OpChat"
    'g_objSettings.JoinMsg = vbNullstring
    'g_objSettings.RedirectIP = vbNullString
    'g_objSettings.RedirectAddress = vbNullString
    ' NEW REDIRECT ADDRESSES
    'g_objSettings.ForMinShareRedirectAddress = vbNullString
    'g_objSettings.ForMaxShareRedirectAddress = vbNullString
    'g_objSettings.ForMinSlotsRedirectAddress = vbNullString
    'g_objSettings.ForMaxSlotsRedirectAddress = vbNullString
    'g_objSettings.ForMaxHubsRedirectAddress = vbNullString
    'g_objSettings.ForSlotPerHubRedirectAddress = vbNullString
    'g_objSettings.ForNoTagHubRedirectAddress = vbNullString
    'g_objSettings.ForTooOldDcppRedirectAddress = vbNullString
    'g_objSettings.ForTooOldNMDCRedirectAddress = vbNullString
    'g_objSettings.ForBWPerSlotRedirectAddress = vbNullString
    'g_objSettings.ForFakeShareRedirectAddress = vbNullString
    'g_objSettings.ForFakeTagRedirectAddress = vbNullString
    'g_objSettings.ForPasModeRedirectAddress = vbNullString
       '
27:    g_objSettings.RegisterIP = "dcreg.mine.nu;reg.hublist.org;dcinfo.dynu.com;hubreg.1stleg.com"
28:    g_objSettings.Ports = "1411;411"
29:    g_objSettings.CSeperator = " "
30:    g_objSettings.MinShareMsg = "You have not met the minimum share."
31:    g_objSettings.DCppMinVersionMsg = "You are using an outdated DC++ client. Please goto http://dcplusplus.sourceforge.net/ and update it."
32:    g_objSettings.MinSlotsMsg = "You do not have enough slots open."
33:    g_objSettings.MaxSlotsMsg = "You have too many slots open."
34:    g_objSettings.HSRatioMsg = "You have not met the hub per slot ratio."
35:    g_objSettings.BSRatioMsg = "You have not met the bandwidth (in KB/s) per slot ratio (as measured by the limiter you are using)."
37:    g_objSettings.MaxHubsMsg = "You are connected to too many hubs. Disconnect from some and reconnect."
38:    g_objSettings.NMDCMinVersionMsg = "You are using an outdated NMDC client. Please goto http://www.neo-modus.com/ and update it. If you are using another client, please change the version setting."
39:    g_objSettings.DenyNoTagMsg = "You do not have an identification tag for your client (ie <++, <DC, etc). Please enable your tag, if possible."
40:    g_objSettings.MaxShareMsg = "You are sharing more than maximum allowed amount."
41:    g_objSettings.FakeShareMsg = "You are suspected of trying to cheat. Goodbye."
42:    g_objSettings.FakeTagMsg = "You are suspected of trying to cheat. Goodbye."
43:    g_objSettings.Socks5Msg = "Socks5 mode not allowed."
44:    g_objSettings.PassiveModeMsg = "Passive mode not allowed."
45:    g_objSettings.NoCOClientsMsg = "Chat only clients are not allowed in here."
46:    g_objSettings.HammeringRd = "a.b.c"

47:    g_objSettings.MaxUsers = 150
48:    g_objSettings.DefaultBanTime = 5
    'g_objSettings.IMinShare = 0
50:    g_objSettings.ScriptTimeout = 15000
51:    g_objSettings.DCMaxHubs = 50
52:    g_objSettings.DCOSlots = 1
53:    g_objSettings.MinSlots = 1
54:    g_objSettings.MaxSlots = 30
55:    g_objSettings.MinShareSize = 3
56:    g_objSettings.MaxShareSize = 3
57:    g_objSettings.CPrefix = 43
58:    g_objSettings.DCOSpeed = 10
59:    g_objSettings.FWInterval = 10000
60:    g_objSettings.FWBanLength = 120
61:    g_objSettings.FWMyINFO = 5
62:    g_objSettings.FWGetNickList = 5
63:    g_objSettings.FWActiveSearch = 15
64:    g_objSettings.FWPassiveSearch = 3
65:    g_objSettings.MaxPassAttempts = 3
66:    g_objSettings.DataFragmentLen = 2048
    'g_objSettings.SendJoinMsg = 0
    'svn 216
69:    g_objSettings.ConDropInterval = 250
70:    g_objSettings.FWDropMsgInterval = 300
    
72:    g_objSettings.DCSlotsPerHub = 1
73:    g_objSettings.DCBandPerSlot = 1
74:    g_objSettings.DCMinVersion = 0.181
75:    g_objSettings.NMDCMinVersion = 0

77:    g_objSettings.MinConnectCls = 1
    
79:    g_objSettings.MinClsConnectSend = True
80:    g_objSettings.MinClsSearchSend = True
    'g_objSettings.AutoCheckUpdate = False
82:    g_objSettings.AutoKickMLDC = True
    
'-----SOCKS5 CHECK--------------------------
    'g_objSettings.DenySocks5 = False
'-----SOCKS5 CHECK END----------------------
    'g_objSettings.DenyPassive = False

89:    g_objSettings.AutoRegister = True
    'g_objSettings.AutoRedirect = False
    'g_objSettings.AutoRedirectFull = False
    'g_objSettings.AutoRedirectNonReg = False
    'g_objSettings.AutoRedirectFullNonReg = False
    'g_objSettings.AutoRedirectFullNonOps = False
    'g_objSettings.AutoStart = False
96:    g_objSettings.CompactDBOnExit = True
    'g_objSettings.ConfirmExit = False
99:    g_objSettings.DCValidateTags = True
    'g_objSettings.DCIncludeOPed
101:    g_objSettings.OPBypass = True
102:    g_objSettings.PreloadWinsocks = True
103:    g_objSettings.SendMessageAFK = True
    'g_objSettings.RegOnly = False
    'g_objSettings.MentoringSystem = False
    'g_objSettings.PreventSearchBots = False
107:    g_objSettings.DescriptiveBanMsg = True
108:    g_objSettings.UseOpChat = True
109:    g_objSettings.UseBotName = True
   'g_objSettings.DisablePassiveSeach = False
    
'-------------Notifications--------------------
112:    g_objSettings.PopUpNewReg = True
    'g_objSettings.PopUpOpConected = False
    'g_objSettings.PopUpOpDisconected = False
115:    g_objSettings.PopUpUserKick = True
116:    g_objSettings.PopUpUserBaned = True
    'g_objSettings.PopUpUserRedirected = False
    'g_objSettings.PopUpStartedServing = False
    'g_objSettings.PopUpStopedServing = False
'    g_objSettings.RedirectFBWPerSlot = False
'    g_objSettings.RedirectFFakeShare = False
'    g_objSettings.RedirectFFakeTag = False
'    g_objSettings.RedirectFPasMode = False
    
'------------End Here----------------------------
127:    g_objSettings.FilterCPrefix = True
128:    g_objSettings.EnabledCommands = True
    'g_objSettings.ScriptSafeMode = False
    'g_objSettings.StartMinimized = False
    'g_objSettings.SendMsgAsPrivate = False
    'g_objSettings.DenyNoTag = False
    'g_objSettings.HideFadeImg = False
134:    g_objSettings.CheckFakeShare = True
135:    g_objSettings.PreventGuessPass = True
136:    g_objSettings.EnableFloodWall = True
137:    g_objSettings.OpsCanRedirect = True
138:    g_objSettings.MinimizeTray = True
139:    g_objSettings.HideMyinfos = True
140:    g_objSettings.ACOClients = True
141:    g_objSettings.MinMyinfoFakeCls = 5
    
       'svn 159 , setable only in xml and scripts...
144:    g_objSettings.FWMainChat = 20
       'g_objSettings.FWGlobal = 60
146:    g_objSettings.ZLINELENGHT = 1400

        'Defaut language
226:    g_objSettings.Interface = "English"
        
        'System Priority defaut value
229:    g_objSettings.PriorityVal = 1
        'g_objSettings.PriorityBl = False
231:    frmHub.sldPriority.Enabled = False
         
        'set defaut skin -------------------------
234:    g_objSettings.blSkin = True
235:    g_objSettings.lngSkin = 1 '01-Defaut

237:    g_objSettings.Plugins = True
        
239:    Call LoadDfsMessages
        
241:    DoEvents

243:    AddLog "Hub settings loaded.", 2

245:  Exit Sub

247:
Err:
249:    HandleError Err.Number, Err.Description, Erl & "|" & "mSubMain.LoadDfsSettings()"
End Sub
