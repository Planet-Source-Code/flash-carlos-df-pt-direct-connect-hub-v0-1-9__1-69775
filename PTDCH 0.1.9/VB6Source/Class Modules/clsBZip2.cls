VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBZip2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function Compress Lib "libbz2.dll" Alias "BZ2_bzBuffToBuffCompress" (dest As Any, destLen As Long, Source As Any, ByVal sourceLen As Long, ByVal blockSize100k As Long, ByVal Verbosity As Long, ByVal workFactor As Long) As Long
Private Declare Function Decompress Lib "libbz2.dll" Alias "BZ2_bzBuffToBuffDecompress" (dest As Any, destLen As Long, Source As Any, ByVal sourceLen As Long, ByVal Small As Long, ByVal Verbosity As Long) As Long

Public Function CompressString(ByRef strData As String) As String
1:    Dim i               As Long
2:    Dim q               As Long
3:    Dim arrData()       As Byte
4:    Dim arrCompressed() As Byte
    
6:    On Error GoTo Err
    
    'Prepare lengths of buffers
9:    i = Len(strData)
10:    q = (i * 1.01) + 600
    
    'Resize arrays and copy strData into the byte array
13:    ReDim arrCompressed(q) As Byte
14:    ReDim arrData(i - 1) As Byte
15:    CopyMemory arrData(0), ByVal strData, i
    
    'Call the Compress DLL function
18:    Compress arrCompressed(0), q, arrData(0), i, 9, 0, 0
    
    'If q != 0 Then we should convert it back to a string
21:    If q Then
22:        CompressString = Space$(q)
23:        CopyMemory ByVal CompressString, arrCompressed(0), q
24:    End If
        
26:    Exit Function
    
28:
Err:
29:    HandleError Err.Number, Err.Description, Erl & "|" & "clsBZip2.CompressString()"
End Function

Public Function DecompressString(ByRef strData As String, Optional ByRef lngFactor As Long = 500) As String
1:    Dim i                   As Long
2:    Dim q                   As Long
3:    Dim arrData()           As Byte
4:    Dim arrDecompressed()   As Byte
    
6:    On Error GoTo Err
    
    'Prepare the lengths of the buffers
9:    i = Len(strData)
10:    q = i * lngFactor
    
    'Resize arrays and copy strData into the byte array
13:    ReDim arrDecompressed(q) As Byte
14:    ReDim arrData(i - 1) As Byte
15:    CopyMemory arrData(0), ByVal strData, i
    
    'Call the Decompress DLL function
18:    Decompress arrDecompressed(0), q, arrData(0), i, 0, 0
    
    'If q != 0 then we should convert it back to a string
21:    If q Then
22:        DecompressString = Space$(q)
23:        CopyMemory ByVal DecompressString, arrDecompressed(0), q
24:    End If
    
26:    Exit Function

28:
Err:
29:    HandleError Err.Number, Err.Description, Erl & "|" & "clsBZip2.DecompressString()"
End Function

Public Sub CompressFile(ByRef strInput As String, ByRef strOutput As String)
1:    Dim i                   As Long
2:    Dim q                   As Long
3:    Dim intFF               As Integer
4:    Dim strCompressed       As String
5:    Dim arrData()           As Byte
6:    Dim arrCompressed()     As Byte
    
8:    On Error GoTo Err
    
    'Make sure file exists
11:    If Not g_objFileAccess.FileExists(strInput) Then Exit Sub
    
13:    intFF = FreeFile
    
    'Open file
16:    Open strInput For Binary Access Read As intFF
        
18:        i = LOF(intFF)
19:        If i Then
            'Read into byte array
21:            ReDim arrData(i - 1) As Byte
22:            Get intFF, , arrData
23:        Else
            'If i = 0, then we don't need to compress it
25:            Exit Sub
26:        End If
        
28:    Close intFF
    
    'Prepare buffer
31:    q = (i * 1.01) + 600
32:    ReDim arrCompressed(q) As Byte
    
    'Compress it
35:    Compress arrCompressed(0), q, arrData(0), i, 9, 0, 0
    
    'Input it into a string
38:    strCompressed = Space$(q)
39:    CopyMemory ByVal strCompressed, arrCompressed(0), q
    
    'Make sure output file does not exist
42:    If g_objFileAccess.FileExists(strOutput) Then Kill strOutput
    
44:    intFF = FreeFile
    
    'Write to file
47:    Open strOutput For Binary Access Write As intFF
48:        Put intFF, , strCompressed
49:    Close intFF
    
51:    Exit Sub
    
53:
Err:
54:    HandleError Err.Number, Err.Description, Erl & "|" & "clsBZip2.CompressFile()"
End Sub

Public Sub DecompressFile(ByRef strInput As String, ByRef strOutput As String, Optional ByRef lngFactor As Long = 500)
1:    Dim i                   As Long
2:    Dim q                   As Long
3:    Dim intFF               As Integer
4:    Dim strDecompressed     As String
5:    Dim arrData()           As Byte
6:    Dim arrDecompressed()   As Byte
    
8:    On Error GoTo Err
    
    'Make sure file exists
11:    If Not g_objFileAccess.FileExists(strInput) Then Exit Sub
    
13:    intFF = FreeFile
    
    'Open file
16:    Open strInput For Binary Access Read As intFF
        
18:        i = LOF(intFF)
19:        If i Then
            'Read into byte array
21:            ReDim arrData(i - 1) As Byte
22:            Get intFF, , arrData
23:        Else
            'If i = 0, then we don't need to decompress it
25:            Exit Sub
26:        End If
        
28:    Close intFF
    
    'Prepare buffer
31:    q = i * lngFactor
32:    ReDim arrDecompressed(q) As Byte
    
    'Compress it
35:    Decompress arrDecompressed(0), q, arrData(0), i, 0, 0
    
37:    strDecompressed = Space$(q)
38:    CopyMemory ByVal strDecompressed, arrDecompressed(0), q
    
    'Make sure output file does not exist
41:    If g_objFileAccess.FileExists(strOutput) Then Kill strOutput
    
43:    intFF = FreeFile
    
    'Write to file
46:    Open strOutput For Binary Access Write As intFF
47:        Put intFF, , strDecompressed
48:    Close intFF
    
50:    Exit Sub
    
52:
Err:
53:    HandleError Err.Number, Err.Description, Erl & "|" & "clsBZip2.DecompressFile()"
End Sub
