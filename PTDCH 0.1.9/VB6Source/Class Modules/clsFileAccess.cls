VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetPPS Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpstrKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePPS Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpstrKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPPI Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpstrKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

'Type
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type

'Write to a file
Public Function WriteFile(ByRef strPath As String, ByRef strData As String) As Long
1:    Dim intFF As Integer
    
3:    On Error GoTo Err
    
    'Make sure file does not exist
6:    If FileExists(strPath) Then Kill strPath

8:    intFF = FreeFile
    
    'Write in binary mode (faster)
11:    Open strPath For Binary Access Write As intFF
12:    Put intFF, , strData
13:    Close intFF
    
15:    Exit Function
    
17:
Err:
18:    WriteFile = Err.Number
19:    HandleError WriteFile, Err.Description, Erl & "|" & "clsFileAccess.WriteFile(""" & strPath & """, """ & strData & """)"
End Function

'Append to a file
Public Function AppendFile(ByRef strPath As String, ByRef strData As String, Optional ByRef blnCarriageReturn As Boolean = True) As Long
1:    Dim intFF As Integer
    
3:    On Error GoTo Err

5:    intFF = FreeFile
    
    'Open in Append mode
8:    Open strPath For Append As intFF
    
    'Remove carriage return if necessary
11:    If blnCarriageReturn Then
12:        Print #intFF, strData
13:    Else
14:        Print #intFF, strData;
15:    End If
    
17:    Close intFF
    
19:    Exit Function
    
21:
Err:
22:    AppendFile = Err.Number
23:    HandleError AppendFile, Err.Description, Erl & "|" & "clsFileAccess.AppendFile(""" & strPath & """, """ & strData & """, " & blnCarriageReturn & ")"
End Function

'Delete a file
Public Function DeleteFile(ByRef strPath As String) As Long
1:    On Error GoTo Err
    
3:    Kill strPath

5:    Exit Function
    
7:
Err:
8:    DeleteFile = Err.Number
End Function

'Read from a file
Public Function ReadFile(ByRef strPath As String) As String
1:    Dim intFF   As Integer
2:    Dim i       As Long
    
4:    On Error GoTo Err

    'Read only if the file exists
7:    If FileExists(strPath) Then
    
9:        intFF = FreeFile
    
        'Open in binary mode for speed
12:        Open strPath For Binary Access Read As intFF
    
        'If length is zero, we don't need to read from the file
15:        i = LOF(intFF)
16:        If i Then
17:            ReadFile = Space$(i)
18:            Get intFF, , ReadFile
19:        End If
    
21:        Close intFF
22:    End If
    
24:    Exit Function
    
26:
Err:
27:    HandleError Err.Number, Err.Description, Erl & "|" & "clsFileAccess.ReadFile(""" & strPath & """)"
End Function

'Create a directory
Public Function CreateDir(ByRef strPath As String) As Long
1:    On Error GoTo Err
    
3:    If FileExists(strPath) = 0 Then MkDir strPath
    
5:    Exit Function
    
7:
Err:
8:    CreateDir = Err.Number
End Function

'Copy a file
Public Function CopyFile(ByRef strPath As String, ByRef strCopy As String) As Long
1:    On Error GoTo Err
    
3:    FileCopy strPath, strCopy

5:    Exit Function
    
7:
Err:
8:    CopyFile = Err.Number
End Function

'Rename a file
Public Function RenameFile(ByRef strOld As String, ByRef strNew As String) As Long
1:    On Error GoTo Err
    
3:    Name strOld As strNew

5:    Exit Function
    
7:
Err:
8:    RenameFile = Err.Number
End Function

'Get a file's attributes
Public Function FileAttributes(ByRef strPath As String) As VbFileAttribute
1:    FileAttributes = GetFileAttributes(strPath)
End Function

'Check if a file exists
Public Function FileExists(ByRef strPath As String) As Boolean
1:    FileExists = Not (GetFileAttributes(strPath) = -1)
End Function

'Access the dir command
Public Function VDir(Optional ByRef PathName As Variant, Optional ByVal Attributes As VbFileAttribute = vbNormal) As String
1:    VDir = Dir(PathName, Attributes)
End Function

'Get the app path
Public Property Get AppPath() As String
1:    AppPath = G_APPPATH
End Property

'Get a String from an INI file
Public Function GetSStr(ByVal strSection As String, ByVal strKey As String, ByVal strDefault As String, ByVal strFile As String, Optional ByVal lngBuffer As Long = 255) As String
1:    On Error GoTo Err
    
    'Create the buffer
4:    GetSStr = String$(lngBuffer, vbNullChar)

6:    If GetPPS(strSection, strKey, strDefault, GetSStr, Len(GetSStr), strFile) Then
        'a vbNullChar doesn't need a "- 1" at the end
8:        GetSStr = LeftB$(GetSStr, InStrB(1, GetSStr, vbNullChar))
9:    Else
10:        WritePPS strSection, strKey, strDefault, strFile
11:        GetSStr = strDefault
12:    End If
    
14:    Exit Function
    
16:
Err:
17:    HandleError Err.Number, Err.Description, Erl & "|" & "clsFileAccess.GetSStr()"
End Function

'Get a Double from an INI file
Public Function GetSDbl(ByVal strSection As String, ByVal strKey As String, ByVal dblDefault As Double, ByVal strFile As String, Optional ByVal lngBuffer As Long = 255) As Double
1:    Dim strBuffer As String
    
3:    On Error GoTo Err
    
    'Create the buffer
6:    strBuffer = String$(lngBuffer, vbNullChar)

8:    If GetPPS(strSection, strKey, dblDefault, strBuffer, Len(strBuffer), strFile) Then
        'a vbNullChar doesn't need a "- 1" at the end
10:        GetSDbl = CDbl(LeftB$(strBuffer, InStrB(1, strBuffer, vbNullChar)))
11:    Else
12:        WritePPS strSection, strKey, CStr(dblDefault), strFile
13:        GetSDbl = dblDefault
14:    End If
    
16:    Exit Function
    
18:
Err:
19:    HandleError Err.Number, Err.Description, Erl & "|" & "clsFileAccess.GetSDbl()"
End Function

'Get a Boolean from an INI file
Public Function GetSBool(ByVal strSection As String, ByVal strKey As String, ByVal blnDefault As Boolean, ByVal strFile As String, Optional ByVal lngBuffer As Long = 40) As Boolean
1:    Dim strBuffer As String
    
3:    On Error GoTo Err
    
    'Create the buffer
6:    strBuffer = String$(lngBuffer, vbNullChar)

8:    If GetPPS(strSection, strKey, blnDefault, strBuffer, Len(strBuffer), strFile) Then
        'a vbNullChar doesn't need a "- 1" at the end
10:        GetSBool = CBool(LeftB$(strBuffer, InStrB(1, strBuffer, vbNullChar)))
11:    Else
12:        WritePPS strSection, strKey, CStr(blnDefault), strFile
13:        GetSBool = blnDefault
14:    End If
    
16:    Exit Function
    
18:
Err:
19:    HandleError Err.Number, Err.Description, Erl & "|" & "clsFileAccess.GetSBool()"
End Function

'Get a Long from an INI file
Public Function GetSLng(ByVal strSection As String, ByVal strKey As String, ByVal lngDefault As Long, ByVal strFile As String) As Long
1:    On Error GoTo Err
    
3:    GetSLng = GetPPI(strSection, strKey, lngDefault, strFile)
    
5:    If GetSLng = 0 Then
6:        WritePPS strSection, strKey, CStr(lngDefault), strFile
7:        GetSLng = lngDefault
8:    End If
    
10:    Exit Function
    
12:
Err:
13:    HandleError Err.Number, Err.Description, Erl & "|" & "clsFileAccess.GetSLng()"
End Function

'Write a string to an INI file
Public Sub WriteSVar(ByVal strSection As String, ByVal strKey As String, ByVal Value As String, ByVal file As String)
1:    WritePPS strSection, strKey, Value, file
End Sub

' ------------------------ NEW INTERFACE LANGUAGE ------------------------
Public Function ListFiles(ByVal strPath As String) As String()
'------------------------------------------------------------------
'Purpose:   read folder content (file names) into a string array
'
'Params:    strPath         = path to folder AND file pattern
'                             a.e. "C:\*.bat"
'
'Returns:   array of files
'------------------------------------------------------------------
9:    Dim WFD             As WIN32_FIND_DATA
    
11:    Dim lngOne          As Long
12:    Dim lngTwo          As Long
13:    Dim intIndex        As Integer
14:    Dim strTemp         As String
15:    Dim ReadFolder      As String

17:    On Error GoTo Err
    
    'Get first file handle
20:    lngOne = FindFirstFile(strPath, WFD)
    
    'If it doesn't equal -1, then there are files
23:    If Not lngOne = -1 Then
24:        Do Until lngTwo = 18&
            'Can't be a directory
26:            If Not (WFD.dwFileAttributes And &H10) = vbDirectory Then
                'Extract file name
28:                lngTwo = InStrB(1, WFD.cFileName, vbNullChar)
                
30:                If lngTwo Then _
                    strTemp = LeftB$(WFD.cFileName, lngTwo) _
                Else _
                    strTemp = WFD.cFileName
                    
                'Increment count
36:                intIndex = intIndex + 1
                
                'Add item
39:                ReadFolder = ReadFolder & strTemp & "\"
            
41:            End If
                
            'Get next file
44:            lngTwo = FindNextFile(lngOne, WFD)
        
            'Exit if it's zero
47:            If lngTwo = 0 Then Exit Do
48:        Loop
        
50:    End If
    
52:    If InStr(ReadFolder, "\") Then
53:        ReadFolder = Mid$(ReadFolder, 1, Len(ReadFolder) - 1)
54:        ListFiles = Split(ReadFolder, "\")
55:    End If

    
58:    Exit Function
    
60:
Err:
61:    HandleError Err.Number, Err.Description, Erl & "|" & "ReadFolder()"
End Function
' ---------------------- NEW INTERFACE LANGUAGE END ----------------------

