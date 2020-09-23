VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Const sFileName = "XML\Plan.xml"
Const sDict = "Scripting.Dictionary"

Const sKomma = ", "
Const vbSpace = " "
Const sNoRight = "Not enough rights."
Const sNoCMD = "Unknown command."

Const sKeyAdd = "add"
Const sKeyDel = "del"
Const sKeySet = "set"
Const sKeyLst = "list"
Const sKeyShw = "show"

Const sTKeyDaily = "daily:" '20:30
Const sTKeyHourly = "hourly:" '45

Dim dictPlan

Private Sub Class_Initialize()
1:    Set dictPlan = CreateObject(sDict)
2:    Call LoadDict(dictPlan, sFileName)
End Sub

Private Sub Class_Terminate()
1:    Set dictPlan = Nothing
End Sub

Public Function Plan(sParameter)
1:    Dim n, sCMD, sIKey, sParam, sPTime, sPLastDone, sPMSG
2:    Dim t, z, vKeys, sKey, sItem, m, tmp

4:    sParam = Trim(sParameter)
5:    sCMD = LCase(g_objFunctions.SplitParameter(sParam, " "))
6:    sIKey = g_objFunctions.SplitParameter(sParam, " ")
7:    If sCMD = "" Then sCMD = sKeyLst

    'Commands
    Select Case LCase(sCMD)
        Case sKeyAdd, sKeySet
10:            If CStr(sIKey) = "" Then
11:                Plan = "Missing or wrong plan name."
12:            Else
13:                If dictPlan.Exists(CStr(sIKey)) Then
14:                    sItem = "$" & dictPlan.Item(CStr(sIKey)) & "$"
15:                    sPTime = g_objFunctions.GetSegment(1, sItem, "$")
16:                    sPLastDone = g_objFunctions.GetSegment(2, sItem, "$")
17:                    sPMSG = g_objFunctions.GetSegment(3, sItem, "$")
18:                Else
19:                    sPTime = ""
20:                    sPLastDone = 0
21:                    sPMSG = ""
22:                End If

24:                If CStr(sPTime) = "" Then sPTime = "(no time)"
25:                If CStr(sPMSG) = "" Then sPMSG = "(no action)"

27:                If CStr(sParam) <> "" Then
28:                    t = g_objFunctions.SplitParameter(sParam, vbSpace)
29:                    If LCase(Left(t, Len(sTKeyDaily))) = sTKeyDaily Or LCase(Left(t, Len(sTKeyHourly))) = sTKeyHourly Then
30:                        sPTime = t
31:                        sPLastDone = 0
32:                        If sParam <> "" Then
33:                            sPMSG = sParam
34:                        End If
35:                    Else
36:                        sPMSG = Trim(t & " " & sParam)
37:                    End If
38:                End If

40:                dictPlan.Item(CStr(sIKey)) = sPTime & "$" & sPLastDone & "$" & sPMSG
41:                Plan = "Action " & sIKey & " in time plan actualisized. " & Plan(sKeyShw & " " & sIKey)
42:                tmp = sIKey & vbTab
43:                If InStr(sPTime, "hourly:") Then _
                    tmp = tmp + vbTab
45:                tmp = tmp & sPTime
46:                tmp = tmp & vbTab & sPMSG
47:                frmHub.lstPlan.AddItem tmp
48:            End If
49:            Call SaveDict(dictPlan, sFileName)

        Case sKeyDel
51:            For X = 0 To dictPlan.count - 1
52:                If sIKey = g_objFunctions.BeforeFirst(frmHub.lstPlan.List(X), vbTab) Then _
                    frmHub.lstPlan.RemoveItem X
54:            Next
55:            Plan = g_objFunctions.DictItemDel(sIKey, dictPlan)
56:            Call SaveDict(dictPlan, sFileName)

        Case sKeyLst
58:            Plan = ""

60:            vKeys = dictPlan.Keys
61:            For m = 1 To dictPlan.count
62:                sKey = CStr(vKeys(m - 1))
63:                sItem = "$" & dictPlan.Item(CStr(sKey)) & "$"

65:                sPTime = g_objFunctions.GetSegment(1, sItem, "$")
66:                sPLastDone = g_objFunctions.GetSegment(2, sItem, "$")
67:                sPMSG = g_objFunctions.GetSegment(3, sItem, "$")

69:                z = "   " & sKey
70:                z = z & vbTab & "  " & sPTime
71:                z = z & vbTab & "  last: " & sPLastDone

73:                Plan = Plan & z & vbCrLf
74:            Next

76:            If Plan = "" Then
77:                Plan = "No items to do."
78:            Else
79:                Plan = "Items in time plan:" & vbCrLf & Plan
80:            End If

        Case sKeyShw
82:            If CStr(sIKey) = "" Then
83:                Plan = "Missing or wrong item name."
84:            Else
85:                If dictPlan.Exists(CStr(sIKey)) Then
86:                    sItem = "$" & dictPlan.Item(CStr(sIKey)) & "$"
87:                    sPTime = g_objFunctions.GetSegment(1, sItem, "$")
88:                    sPLastDone = g_objFunctions.GetSegment(2, sItem, "$")
89:                    sPMSG = g_objFunctions.GetSegment(3, sItem, "$")

91:                    Plan = "Details of the item:" & vbCrLf
92:                    Plan = Plan & vbTab & "ID:" & vbTab & sIKey & vbCrLf
93:                    Plan = Plan & vbTab & "planned:" & vbTab & sPTime & vbCrLf
94:                    Plan = Plan & vbTab & "last:" & vbTab & sPLastDone & vbCrLf
95:                    Plan = Plan & vbTab & "Action:" & vbTab & sPMSG & vbCrLf
96:                Else
97:                    Plan = "Item """ & sIKey & """ not found."
98:                End If
99:            End If

        Case Else
101:            Plan = sNoCMD

103:    End Select
End Function

Public Function PlanDo()
1:    Dim vKeys, sKey, sItem, m
2:    Dim sPTime, sPLastDone, sPMSG
3:    Dim sLastHaveDone

5:    vKeys = dictPlan.Keys
6:    For m = 1 To dictPlan.count
7:        sKey = CStr(vKeys(m - 1))
8:        sItem = "$" & dictPlan.Item(CStr(sKey)) & "$"

10:        sPTime = g_objFunctions.GetSegment(1, sItem, "$")
11:        sPLastDone = g_objFunctions.GetSegment(2, sItem, "$")
12:        sPMSG = g_objFunctions.GetSegment(3, sItem, "$")

14:        sLastHaveDone = LastHaveDone(sPTime)
15:        If sPLastDone <> sLastHaveDone Then
16:            If CStr(sPLastDone) <> "0" Then
17:                PlanDo = PlanDo & sPMSG & "$"
18:            End If
19:            dictPlan.Item(CStr(sKey)) = sPTime & "$" & sLastHaveDone & "$" & sPMSG
20:        End If
21:    Next
End Function

Private Function LastHaveDone(sPTime)
1:    Dim n, sKey, sData
2:    Dim nSecSinceMN

4:    n = InStr(sPTime, ":")
5:    If n > 0 Then
6:        sKey = Left(sPTime, n)
7:        sData = Mid(sPTime, n + 1)

        Select Case LCase(sKey)
            Case sTKeyDaily
9:                nSecSinceMN = DateDiff("s", 0, CDate(Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)))
10:                If CLng(nSecSinceMN) >= DateDiff("s", 0, CDate(sData)) Then
                    'already done today
12:                    LastHaveDone = g_objFunctions.ForceMinLen(Day(Now), 2, "0") & "." & g_objFunctions.ForceMinLen(Month(Now), 2, "0") & "." & Year(Now)
13:                Else
                    'has to be done
15:                    LastHaveDone = g_objFunctions.ForceMinLen(Day(DateAdd("d", -1, Now)), 2, "0") & "." & g_objFunctions.ForceMinLen(Month(DateAdd("d", -1, Now)), 2, "0") & "." & Year(DateAdd("d", -1, Now))
16:                End If
17:                LastHaveDone = LastHaveDone & " " & g_objFunctions.ForceMinLen(Hour(CDate(sData)), 2, "0") & ":" & g_objFunctions.ForceMinLen(Minute(CDate(sData)), 2, "0") & ":" & g_objFunctions.ForceMinLen(Second(CDate(sData)), 2, "0")

            Case sTKeyHourly
'                sData = Minute(sData)   '?????????????
20:                If CLng(Minute(Now)) >= CLng(sData) Then
                    'already done
22:                    LastHaveDone = g_objFunctions.ForceMinLen(Day(Now), 2, "0") & "." & g_objFunctions.ForceMinLen(Month(Now), 2, "0") & "." & Year(Now) & " " & g_objFunctions.ForceMinLen(Hour(Now), 2, "0")
23:                Else
                    'to do
25:                    LastHaveDone = g_objFunctions.ForceMinLen(Day(DateAdd("h", -1, Now)), 2, "0") & "." & g_objFunctions.ForceMinLen(Month(DateAdd("h", -1, Now)), 2, "0") & "." & Year(DateAdd("h", -1, Now)) & " " & g_objFunctions.ForceMinLen(Hour(DateAdd("h", -1, Now)), 2, "0")
26:                End If
27:                LastHaveDone = LastHaveDone & ":" & g_objFunctions.ForceMinLen(CLng(sData), 2, "0")

29:        End Select
30:    End If
End Function

Private Function SaveDict(dictDict, sFileName)
1:    Dim Keys, X, FSO, f

3:    Set FSO = CreateObject("Scripting.FileSystemObject")
4:    Set f = FSO.CreateTextFile(sFileName, True)
5:    Keys = dictDict.Keys
6:    f.WriteLine "<Plan>"
7:    f.WriteLine CStr(dictDict.count)
8:    For X = 0 To (dictDict.count - 1)
9:        f.WriteLine CStr(Keys(X))
10:        f.WriteLine CStr(Replace(dictDict(CStr(Keys(X))), vbCrLf, "|"))
11:    Next
12:    f.WriteLine "</Plan>"
13:    f.Close
14:    SaveDict = dictDict.count
End Function

Private Function LoadDict(dictDict, sFileName)
1:    Const ForReading = 1
2:    Dim i, X, FSO, f, tmpU, tmpT, tmp

4:    Set FSO = CreateObject("Scripting.FileSystemObject")
5:    If (FSO.FileExists(sFileName)) Then
6:        Set f = FSO.OpenTextFile(sFileName, ForReading)
7:        tmp = f.ReadLine
8:        i = CInt(f.ReadLine)
9:        If i > 0 Then
10:            For X = 1 To i
11:                tmpU = CStr(f.ReadLine)
12:                tmpT = CStr(f.ReadLine)
13:                dictDict(tmpU) = Replace(tmpT, "|", vbCrLf)
14:                tmp = g_objFunctions.BeforeFirst(tmpT, "$") & vbTab & g_objFunctions.AfterLast(tmpT, "$")
15:                If InStr(tmp, "hourly:") Then tmp = Replace(tmp, vbTab, vbTab & vbTab)
16:                frmHub.lstPlan.AddItem tmpU & vbTab & tmp
17:            Next
18:        End If
19:        f.Close
20:    End If
21:    LoadDict = i
End Function



