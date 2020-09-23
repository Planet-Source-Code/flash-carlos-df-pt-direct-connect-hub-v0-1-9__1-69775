VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsZLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function Compress Lib "zlib.dll" Alias "compress" (ByVal compr As String, ByRef comprLen As Any, ByVal buf As String, ByVal buflen As Long) As Long
Private Declare Function Decompress Lib "zlib.dll" Alias "uncompress" (ByVal uncompr As String, ByRef uncomprLen As Any, ByVal compr As String, ByVal lcompr As Long) As Long

Public Function CompressString(ByRef strData As String) As String
1:    Dim i As Long
2:    Dim q As Long

4:    On Error GoTo Err

    'Prepare buffer
7:    i = Len(strData)
8:    q = (i * 1.01) + 13
9:    CompressString = Space$(q)

    'If <> 0, then an error occured
12:    If Compress(CompressString, q, strData, i) Then
13:        CompressString = vbNullString
14:    Else
        'Put the uncompressed length identifier in front of compressed data
        'easyer on ressources to uncompressed data if known.
        'CompressString = i & ":" & Left$(CompressString, q)
        
        'In some cases might not be available when receiving, external to the hub, data...
20:        CompressString = Left$(CompressString, q)
21:    End If

23:    Exit Function

25:
Err:
26:    HandleError Err.Number, Err.Description, Erl & "|" & "clsZLib.CompressString()"
End Function

Public Function DecompressString(ByRef strData As String) As String
1:    Dim i           As Long
2:    Dim q           As Long
3:    Dim strBuffer   As String

5:    On Error GoTo Err

    'Guess length of uncompressed data
    'up to a 49/1 ratio can be achive...
9:    i = LenB(strData) * 25

    'q = InStrB(1, strData, ":")
    'If q Then
    '    i = CLng(LeftB$(strData, q - 1))
    '    strBuffer = MidB$(strData, q + 2)

        'Prepare buffer
17:        DecompressString = Space$(i)
    'Else
        'If not found, exit
    '    Exit Function
    'End If

    'If it returns a non-zero value, an error occured (so clean out DecompressString value)
    'If Not Decompress(DecompressString, i, strBuffer, Len(strBuffer)) = 0 Then DecompressString = vbNullString
25:    If Not Decompress(DecompressString, i, strData, Len(strBuffer)) = 0 Then DecompressString = vbNullString

27:    Exit Function

29:
Err:
30:    HandleError Err.Number, Err.Description, Erl & "|" & "clsZLib.DecompressString()"
End Function

Public Sub CompressFile(ByRef strInput As String, ByRef strOutput As String)
1:    On Error GoTo Err
    
    'I could make it slightly faster, and might do so later
    'However this is quite acceptable right now
5:    g_objFileAccess.WriteFile strOutput, CompressString(g_objFileAccess.ReadFile(strInput))
    
7:    Exit Sub
    
9:
Err:
10:    HandleError Err.Number, Err.Description, Erl & "|" & "clsZLib.CompressFile()"
End Sub

Public Sub DecompressFile(ByRef strInput As String, ByRef strOutput As String)
1:    On Error GoTo Err
    
    'I could make it slightly faster, and might do so later
    'However this is quite acceptable right now
5:    g_objFileAccess.WriteFile strOutput, DecompressString(g_objFileAccess.ReadFile(strInput))
    
7:    Exit Sub
    
9:
Err:
10:    HandleError Err.Number, Err.Description, Erl & "|" & "clsZLib.DecompressFile()"
End Sub
