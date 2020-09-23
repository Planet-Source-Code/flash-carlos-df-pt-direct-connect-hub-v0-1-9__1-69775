VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHuffman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Huffman Encoding/Decoding Class
'-------------------------------
'
'(c) 2000, Fredrik Qvarfort
'

'Some modifications (Former hub Team) - prevent error when loading files
'                               - removed the progress event (as this is only used by scripts)
'                               - added my default error handling

Option Explicit

''Progress Values for the encoding routine
'Private Const PROGRESS_CALCFREQUENCY As Long = 7
'Private Const PROGRESS_CALCCRC As Long = 5
'Private Const PROGRESS_ENCODING As Long = 88

''Progress Values for the decoding routine
'Private Const PROGRESS_DECODING As Long = 89
'Private Const PROGRESS_CHECKCRC As Long = 11

''Events
'Event Progress(Procent As Integer)

Private Type HUFFMANTREE
  ParentNode As Integer
  RightNode As Integer
  LeftNode As Integer
  Value As Integer
  Weight As Long
End Type

Private Type ByteArray
  count As Byte
  Data() As Byte
End Type

Public Sub EncodeFile(SourceFile As String, DestFile As String)
1:  Dim ByteArray() As Byte
2:  Dim Filenr As Integer
3:  Dim i As Integer
  
5:  On Error GoTo Err
  
  'Make sure the source file exists
8:  If (Not FileExist(SourceFile)) Then
9:    Err.Raise vbObjectError, "clsHuffman.EncodeFile()", "Source file does not exist"
10:  End If
  
  'Read the data from the sourcefile
13:  Filenr = FreeFile
14:  Open SourceFile For Binary As #Filenr
  
16:  i = LOF(Filenr)
17:  If i Then
18:    ReDim ByteArray(0 To i - 1)
19:    Get #Filenr, , ByteArray()
20:  End If
    
22:  Close #Filenr
  
  'Compress the data
25:  Call EncodeByte(ByteArray(), UBound(ByteArray) + 1)
  
  'If the destination file exist we need to
  'destroy it because opening it as binary
  'will not clear the old data
30:  If (FileExist(DestFile)) Then Kill DestFile
  
  'Save the destination string
33:  Open DestFile For Binary As #Filenr
34:  Put #Filenr, , ByteArray()
35:  Close #Filenr

37:  Exit Sub

39:
Err:
40:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHuffman.EncodeFile()"
End Sub

Public Sub DecodeFile(SourceFile As String, DestFile As String)
1:  Dim ByteArray() As Byte
2:  Dim Filenr As Integer
3:  Dim i As Integer
  
5:  On Error GoTo Err

  'Make sure the source file exists
8:  If (Not FileExist(SourceFile)) Then
9:    Err.Raise vbObjectError, "clsHuffman.DecodeFile()", "Source file does not exist"
10:  End If
  
  'Read the data from the sourcefile
13:  Filenr = FreeFile
14:  Open SourceFile For Binary As #Filenr
  
16:  i = LOF(Filenr)
17:  If i Then
18:    ReDim ByteArray(0 To i - 1)
19:    Get #Filenr, , ByteArray()
20:  End If
    
22:  Close #Filenr
  
  'Uncompress the data
25:  Call DecodeByte(ByteArray(), UBound(ByteArray) + 1)
  
  'If the destination file exist we need to
  'destroy it because opening it as binary
  'will not clear the old data
30:  If (FileExist(DestFile)) Then Kill DestFile
  
  'Save the destination string
33:  Open DestFile For Binary As #Filenr
34:  Put #Filenr, , ByteArray()
35:  Close #Filenr

37:  Exit Sub

39:
Err:
40:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHuffman.DecodeFile()"
End Sub

Private Sub CreateTree(Nodes() As HUFFMANTREE, NodesCount As Long, Char As Long, bytes As ByteArray)
1:  Dim a As Integer
2:  Dim NodeIndex As Long
  
4:  On Error GoTo Err

6:  NodeIndex = 0
7:  For a = 0 To (bytes.count - 1)
8:    If (bytes.Data(a) = 0) Then
      'Left node
10:      If (Nodes(NodeIndex).LeftNode = -1) Then
11:        Nodes(NodeIndex).LeftNode = NodesCount
12:        Nodes(NodesCount).ParentNode = NodeIndex
13:        Nodes(NodesCount).LeftNode = -1
14:        Nodes(NodesCount).RightNode = -1
15:        Nodes(NodesCount).Value = -1
16:        NodesCount = NodesCount + 1
17:      End If
18:      NodeIndex = Nodes(NodeIndex).LeftNode
19:    ElseIf (bytes.Data(a) = 1) Then
      'Right node
21:      If (Nodes(NodeIndex).RightNode = -1) Then
22:        Nodes(NodeIndex).RightNode = NodesCount
23:        Nodes(NodesCount).ParentNode = NodeIndex
24:        Nodes(NodesCount).LeftNode = -1
25:        Nodes(NodesCount).RightNode = -1
26:        Nodes(NodesCount).Value = -1
27:        NodesCount = NodesCount + 1
28:      End If
29:      NodeIndex = Nodes(NodeIndex).RightNode
30:    Else
31:      Stop
32:    End If
33:  Next
  
35:  Nodes(NodeIndex).Value = Char
  
37:  Exit Sub
  
39:
Err:
40:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHuffman.CreateTree()"
End Sub

Private Sub EncodeByte(ByteArray() As Byte, ByteLen As Long)
1:  Dim i As Long
2:  Dim j As Long
3:  Dim Char As Byte
4:  Dim BitPos As Byte
5:  Dim lNode1 As Long
6:  Dim lNode2 As Long
7:  Dim lNodes As Long
8:  Dim lLength As Long
9:  Dim count As Integer
10:  Dim lWeight1 As Long
11:  Dim lWeight2 As Long
12:  Dim Result() As Byte
13:  Dim ByteValue As Byte
14:  Dim ResultLen As Long
15:  Dim bytes As ByteArray
16:  Dim NodesCount As Integer
  'Dim NewProgress As Integer
  'Dim CurrProgress As Integer
19:  Dim BitValue(0 To 7) As Byte
20:  Dim CharCount(0 To 255) As Long
21:  Dim Nodes(0 To 511) As HUFFMANTREE
22:  Dim CharValue(0 To 255) As ByteArray
  
24:  On Error GoTo Err
  
  'If the source string is empty or contains
  'only one character we return it uncompressed
  'with the prefix string "HEO" & vbCr
29:  If (ByteLen = 0) Then
30:    ReDim Preserve ByteArray(0 To ByteLen + 3)
31:    If (ByteLen > 0) Then
32:      Call CopyMemory(ByteArray(4), ByteArray(0), ByteLen)
33:    End If
34:    ByteArray(0) = 72 '"H"
35:    ByteArray(1) = 69 '"E"
36:    ByteArray(2) = 48 '"0"
37:    ByteArray(3) = 13 'vbCr
38:    Exit Sub
39:  End If
  
  'Create the temporary result array and make
  'space for identifier, checksum, textlen and
  'the ASCII values inside the Huffman Tree
44:  ReDim Result(0 To 522)
  
  'Prefix the destination string with the
  '"HE3" & vbCr identification string
48:  Result(0) = 72
49:  Result(1) = 69
50:  Result(2) = 51
51:  Result(3) = 13
52:  ResultLen = 4
  
  'Count the frequency of each ASCII code
55:  For i = 0 To (ByteLen - 1)
56:    CharCount(ByteArray(i)) = CharCount(ByteArray(i)) + 1
    'If (i Mod 1000 = 0) Then
    '  NewProgress = i / ByteLen * PROGRESS_CALCFREQUENCY
    '  If (NewProgress <> CurrProgress) Then
    '    CurrProgress = NewProgress
    '    RaiseEvent Progress(CurrProgress)
    '  End If
    'End If
64:  Next
  
  'Create a leaf for each character
67:  For i = 0 To 255
68:    If (CharCount(i) > 0) Then
69:      With Nodes(NodesCount)
70:        .Weight = CharCount(i)
71:        .Value = i
72:        .LeftNode = -1
73:        .RightNode = -1
74:        .ParentNode = -1
75:      End With
76:      NodesCount = NodesCount + 1
77:    End If
78:  Next
  
  'Create the Huffman Tree
81:  For lNodes = NodesCount To 2 Step -1
    'Get the two leafs with the smallest weights
83:    lNode1 = -1: lNode2 = -1
84:    For i = 0 To (NodesCount - 1)
85:      If (Nodes(i).ParentNode = -1) Then
86:        If (lNode1 = -1) Then
87:          lWeight1 = Nodes(i).Weight
88:          lNode1 = i
89:        ElseIf (lNode2 = -1) Then
90:          lWeight2 = Nodes(i).Weight
91:          lNode2 = i
92:        ElseIf (Nodes(i).Weight < lWeight1) Then
93:          If (Nodes(i).Weight < lWeight2) Then
94:            If (lWeight1 < lWeight2) Then
95:              lWeight2 = Nodes(i).Weight
96:              lNode2 = i
97:            Else
98:              lWeight1 = Nodes(i).Weight
99:              lNode1 = i
100:            End If
101:          Else
102:            lWeight1 = Nodes(i).Weight
103:            lNode1 = i
104:          End If
105:        ElseIf (Nodes(i).Weight < lWeight2) Then
106:          lWeight2 = Nodes(i).Weight
107:          lNode2 = i
108:        End If
109:      End If
110:    Next
    
    'Create a new leaf
113:    With Nodes(NodesCount)
114:      .Weight = lWeight1 + lWeight2
115:      .LeftNode = lNode1
116:      .RightNode = lNode2
117:      .ParentNode = -1
118:      .Value = -1
119:    End With
    
    'Set the parentnodes of the two leafs
122:    Nodes(lNode1).ParentNode = NodesCount
123:    Nodes(lNode2).ParentNode = NodesCount
    
    'Increase the node counter
126:    NodesCount = NodesCount + 1
127:  Next

  'Traverse the tree to get the bit sequence
  'for each character, make temporary room in
  'the data array to hold max theoretical size
132:  ReDim bytes.Data(0 To 255)
133:  Call CreateBitSequences(Nodes(), NodesCount - 1, bytes, CharValue)
  
  'Calculate the length of the destination
  'string after encoding
137:  For i = 0 To 255
138:    If (CharCount(i) > 0) Then
139:      lLength = lLength + CharValue(i).count * CharCount(i)
140:    End If
141:  Next
142:  lLength = IIfLng(lLength Mod 8 = 0, lLength \ 8, lLength \ 8 + 1)
  
  'If the destination is larger than the source
  'string we leave it uncompressed and prefix
  'it with a 4 byte header ("HE0" & vbCr)
147:  If ((lLength = 0) Or (lLength > ByteLen)) Then
148:    ReDim Preserve ByteArray(0 To ByteLen + 3)
149:    Call CopyMemory(ByteArray(4), ByteArray(0), ByteLen)
150:    ByteArray(0) = 72
151:    ByteArray(1) = 69
152:    ByteArray(2) = 48
153:    ByteArray(3) = 13
154:    Exit Sub
155:  End If
  
  'Add a simple checksum value to the result
  'header for corruption identification
159:  Char = 0
160:  For i = 0 To (ByteLen - 1)
161:    Char = Char Xor ByteArray(i)
    'If (i Mod 10000 = 0) Then
    '  NewProgress = i / ByteLen * PROGRESS_CALCCRC + PROGRESS_CALCFREQUENCY
    '  If (NewProgress <> CurrProgress) Then
    '    CurrProgress = NewProgress
    '    RaiseEvent Progress(CurrProgress)
    '  End If
    'End If
169:  Next
170:  Result(ResultLen) = Char
171:  ResultLen = ResultLen + 1
  
  'Add the length of the source string to the
  'header for corruption identification
175:  Call CopyMemory(Result(ResultLen), ByteLen, 4)
176:  ResultLen = ResultLen + 4
  
  'Create a small array to hold the bit values,
  'this is faster than calculating on-fly
180:  For i = 0 To 7
181:    BitValue(i) = 2 ^ i
182:  Next
  
  'Store the number of characters used
185:  count = 0
186:  For i = 0 To 255
187:    If (CharValue(i).count > 0) Then
188:      count = count + 1
189:    End If
190:  Next
191:  Call CopyMemory(Result(ResultLen), count, 2)
192:  ResultLen = ResultLen + 2
  
  'Store the used characters and the length
  'of their respective bit sequences
196:  count = 0
197:  For i = 0 To 255
198:    If (CharValue(i).count > 0) Then
199:      Result(ResultLen) = i
200:      ResultLen = ResultLen + 1
201:      Result(ResultLen) = CharValue(i).count
202:      ResultLen = ResultLen + 1
203:      count = count + 16 + CharValue(i).count
204:    End If
205:  Next
  
  'Make room for the Huffman Tree in the
  'destination byte array
209:  ReDim Preserve Result(0 To ResultLen + count \ 8)
  
  'Store the Huffman Tree into the result
  'converting the bit sequences into bytes
213:  BitPos = 0
214:  ByteValue = 0
215:  For i = 0 To 255
216:    With CharValue(i)
217:      If (.count > 0) Then
218:        For j = 0 To (.count - 1)
219:          If (.Data(j)) Then ByteValue = ByteValue + BitValue(BitPos)
220:          BitPos = BitPos + 1
221:          If (BitPos = 8) Then
222:            Result(ResultLen) = ByteValue
223:            ResultLen = ResultLen + 1
224:            ByteValue = 0
225:            BitPos = 0
226:          End If
227:        Next
228:      End If
229:    End With
230:  Next
231:  If (BitPos > 0) Then
232:    Result(ResultLen) = ByteValue
233:    ResultLen = ResultLen + 1
234:  End If
  
  'Resize the destination string to be able to
  'contain the encoded string
238:  ReDim Preserve Result(0 To ResultLen - 1 + lLength)
  
  'Now we can encode the data by exchanging each
  'ASCII byte for its appropriate bit string.
242:  Char = 0
243:  BitPos = 0
244:  For i = 0 To (ByteLen - 1)
245:    With CharValue(ByteArray(i))
246:      For j = 0 To (.count - 1)
247:        If (.Data(j) = 1) Then Char = Char + BitValue(BitPos)
248:        BitPos = BitPos + 1
249:        If (BitPos = 8) Then
250:          Result(ResultLen) = Char
251:          ResultLen = ResultLen + 1
252:          BitPos = 0
253:          Char = 0
254:        End If
255:      Next
256:    End With
    'If (i Mod 10000 = 0) Then
    '  NewProgress = i / ByteLen * PROGRESS_ENCODING + PROGRESS_CALCCRC + PROGRESS_CALCFREQUENCY
    '  If (NewProgress <> CurrProgress) Then
    '    CurrProgress = NewProgress
    '    RaiseEvent Progress(CurrProgress)
    '  End If
    'End If
264:  Next

  'Add the last byte
267:  If (BitPos > 0) Then
268:    Result(ResultLen) = Char
269:    ResultLen = ResultLen + 1
270:  End If
  
  'Return the destination in string format
273:  ReDim ByteArray(0 To ResultLen - 1)
274:  Call CopyMemory(ByteArray(0), Result(0), ResultLen)

  'Make sure we get a "100%" progress message
  'If (CurrProgress <> 100) Then
  '  RaiseEvent Progress(100)
  'End If

281:  Exit Sub

283:
Err:
284:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHuffman.EncodeByte()"
End Sub

Public Function DecodeString(Text As String) As String
1:  Dim ByteArray() As Byte
  
3:  On Error GoTo Err
  
  'Convert the string to a byte array
6:  ByteArray() = StrConv(Text, vbFromUnicode)
  
  'Compress the byte array
9:  Call DecodeByte(ByteArray, Len(Text))
  
  'Convert the compressed byte array to a string
12:  DecodeString = StrConv(ByteArray(), vbUnicode)

14:  Exit Function
  
16:
Err:
17:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHuffman.DecodeString()"
End Function

Public Function EncodeString(Text As String) As String
1:  Dim ByteArray() As Byte
  
3:  On Error GoTo Err
  
  'Convert the string to a byte array
6:  ByteArray() = StrConv(Text, vbFromUnicode)
  
  'Compress the byte array
9:  Call EncodeByte(ByteArray, Len(Text))
  
  'Convert the compressed byte array to a string
12:  EncodeString = StrConv(ByteArray(), vbUnicode)

14:  Exit Function

16:
Err:
17:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHuffman.EncodeString()"
End Function

Private Sub DecodeByte(ByteArray() As Byte, ByteLen As Long)
1:  Dim i As Long
2:  Dim j As Long
  'Dim Pos As Long
4:  Dim Char As Byte
5:  Dim CurrPos As Long
6:  Dim count As Integer
7:  Dim CheckSum As Byte
8:  Dim Result() As Byte
9:  Dim BitPos As Integer
10:  Dim NodeIndex As Long
11:  Dim ByteValue As Byte
12:  Dim ResultLen As Long
13:  Dim NodesCount As Long
14:  Dim lResultLen As Long
  'Dim NewProgress As Integer
  'Dim CurrProgress As Integer
17:  Dim BitValue(0 To 7) As Byte
18:  Dim Nodes(0 To 511) As HUFFMANTREE
19:  Dim CharValue(0 To 255) As ByteArray
  
21:  On Error GoTo Err
  
23:  If (ByteArray(0) <> 72) Or (ByteArray(1) <> 69) Or (ByteArray(3) <> 13) Then
    'The source did not contain the identification
    'string "HE?" & vbCr where ? is undefined at
    'the moment (does not matter)
27:  ElseIf (ByteArray(2) = 48) Then
    'The text is uncompressed, return the substring
    'Decode = MidB$(Text, 9)
30:    Call CopyMemory(ByteArray(0), ByteArray(4), ByteLen - 4)
31:    ReDim Preserve ByteArray(0 To ByteLen - 5)
32:    Exit Sub
33:  ElseIf (ByteArray(2) <> 51) Then
    'This is not a Huffman encoded string
35:    Err.Raise vbObjectError, "HuffmanDecode()", "The data either was not compressed with HE3 or is corrupt (identification string not found)"
36:    Exit Sub
37:  End If
  
39:  CurrPos = 5
    
  'Extract the checksum
42:  CheckSum = ByteArray(CurrPos - 1)
43:  CurrPos = CurrPos + 1
  
  'Extract the length of the original string
46:  Call CopyMemory(ResultLen, ByteArray(CurrPos - 1), 4)
47:  CurrPos = CurrPos + 4
48:  lResultLen = ResultLen
  
  'If the compressed string is empty we can
  'skip the function right here
52:  If (ResultLen = 0) Then Exit Sub
  
  'Create the result array
55:  ReDim Result(0 To ResultLen - 1)
  
  'Get the number of characters used
58:  Call CopyMemory(count, ByteArray(CurrPos - 1), 2)
59:  CurrPos = CurrPos + 2
  
  'Get the used characters and their
  'respective bit sequence lengths
63:  For i = 1 To count
64:    With CharValue(ByteArray(CurrPos - 1))
65:      CurrPos = CurrPos + 1
66:      .count = ByteArray(CurrPos - 1)
67:      CurrPos = CurrPos + 1
68:      ReDim .Data(0 To .count - 1)
69:    End With
70:  Next
  
  'Create a small array to hold the bit values,
  'this is (still) faster than calculating on-fly
74:  For i = 0 To 7
75:    BitValue(i) = 2 ^ i
76:  Next
  
  'Extract the Huffman Tree, converting the
  'byte sequence to bit sequences
80:  ByteValue = ByteArray(CurrPos - 1)
81:  CurrPos = CurrPos + 1
82:  BitPos = 0
83:  For i = 0 To 255
84:    With CharValue(i)
85:      If (.count > 0) Then
86:        For j = 0 To (.count - 1)
87:          If (ByteValue And BitValue(BitPos)) Then .Data(j) = 1
88:          BitPos = BitPos + 1
89:          If (BitPos = 8) Then
90:            ByteValue = ByteArray(CurrPos - 1)
91:            CurrPos = CurrPos + 1
92:            BitPos = 0
93:          End If
94:        Next
95:      End If
96:    End With
97:  Next
98:  If (BitPos = 0) Then CurrPos = CurrPos - 1
  
  'Create the Huffman Tree
101:  NodesCount = 1
102:  Nodes(0).LeftNode = -1
103:  Nodes(0).RightNode = -1
104:  Nodes(0).ParentNode = -1
105:  Nodes(0).Value = -1
106:  For i = 0 To 255
107:    Call CreateTree(Nodes(), NodesCount, i, CharValue(i))
108:  Next
  
  'Decode the actual data
111:  ResultLen = 0
112:  For CurrPos = CurrPos To ByteLen
113:    ByteValue = ByteArray(CurrPos - 1)
114:    For BitPos = 0 To 7
115:      If (ByteValue And BitValue(BitPos)) Then
116:        NodeIndex = Nodes(NodeIndex).RightNode
117:      Else
118:        NodeIndex = Nodes(NodeIndex).LeftNode
119:      End If
120:      If (Nodes(NodeIndex).Value > -1) Then
121:        Result(ResultLen) = Nodes(NodeIndex).Value
122:        ResultLen = ResultLen + 1
123:        If (ResultLen = lResultLen) Then GoTo DecodeFinished
124:        NodeIndex = 0
125:      End If
126:    Next
    'If (CurrPos Mod 10000 = 0) Then
    '  NewProgress = CurrPos / ByteLen * PROGRESS_DECODING
    '  If (NewProgress <> CurrProgress) Then
    '    CurrProgress = NewProgress
    '    RaiseEvent Progress(CurrProgress)
    '  End If
    'End If
134:  Next
135:
DecodeFinished:

  'Verify data to check for corruption.
138:  Char = 0
139:  For i = 0 To (ResultLen - 1)
140:    Char = Char Xor Result(i)
    'If (i Mod 10000 = 0) Then
    '  NewProgress = i / ResultLen * PROGRESS_CHECKCRC + PROGRESS_DECODING
    '  If (NewProgress <> CurrProgress) Then
    '    CurrProgress = NewProgress
    '    RaiseEvent Progress(CurrProgress)
    '  End If
    'End If
148:  Next
149:  If (Char <> CheckSum) Then
150:    Err.Raise vbObjectError, "clsHuffman.Decode()", "The data might be corrupted (checksum did not match expected value)"
151:  End If

  'Return the uncompressed string
154:  ReDim ByteArray(0 To ResultLen - 1)
155:  Call CopyMemory(ByteArray(0), Result(0), ResultLen)
  
  'Make sure we get a "100%" progress message
  'If (CurrProgress <> 100) Then
  '  RaiseEvent Progress(100)
  'End If

162:  Exit Sub

164:
Err:
165:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHuffman.DecodeByte()"
End Sub

Private Sub CreateBitSequences(Nodes() As HUFFMANTREE, ByVal NodeIndex As Integer, bytes As ByteArray, CharValue() As ByteArray)
1:  Dim NewBytes As ByteArray
  
3:  On Error GoTo Err

  'If this is a leaf we set the characters bit
  'sequence in the CharValue array
7:  If (Nodes(NodeIndex).Value > -1) Then
8:    CharValue(Nodes(NodeIndex).Value) = bytes
9:    Exit Sub
10:  End If
  
  'Traverse the left child
13:  If (Nodes(NodeIndex).LeftNode > -1) Then
14:    NewBytes = bytes
15:    NewBytes.Data(NewBytes.count) = 0
16:    NewBytes.count = NewBytes.count + 1
17:    Call CreateBitSequences(Nodes(), Nodes(NodeIndex).LeftNode, NewBytes, CharValue)
18:  End If
  
  'Traverse the right child
21:  If (Nodes(NodeIndex).RightNode > -1) Then
22:    NewBytes = bytes
23:    NewBytes.Data(NewBytes.count) = 1
24:    NewBytes.count = NewBytes.count + 1
25:    Call CreateBitSequences(Nodes(), Nodes(NodeIndex).RightNode, NewBytes, CharValue)
26:  End If
  
28:  Exit Sub
  
30:
Err:
31:    HandleError Err.Number, Err.Description, Erl & "|" & "clsHuffman.CreateBitSequences()"
End Sub

Private Function FileExist(filename As String) As Boolean
1:  On Error GoTo FileDoesNotExist
  
3:  Call FileLen(filename)
4:  FileExist = True
5:  Exit Function
  
7:
FileDoesNotExist:
8:  FileExist = False
End Function
