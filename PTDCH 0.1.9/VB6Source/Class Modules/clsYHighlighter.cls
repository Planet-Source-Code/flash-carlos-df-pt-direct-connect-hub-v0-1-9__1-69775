VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsYHighlighter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private hlCount As Long
Private HCount As Integer
Private CurHigh As Integer
Private Const ciIncriment As Integer = 15000
Private lOffset As Long

Private sBuffer As String

Private Sub ReInit()
1: sBuffer = ""
2: lOffset = 0
End Sub

Private Function GetString() As String
1: GetString = Left$(sBuffer, lOffset)
2: sBuffer = "" 'reset
End Function

'This function lets you assign a string to the concating buffer.
Private Sub SetString(ByRef Source As String)
1: sBuffer = Source & String$(ciIncriment, 0)
End Sub

Private Function FindHighlighter(strLangName As String) As Integer
  On Error GoTo Err
  
3:  Dim i As Integer
  
5:   For i = 0 To UBound(G_Highlighters) - 1
6:    If UCase(G_Highlighters(i).strName) = UCase(strLangName) Then
7:      FindHighlighter = i
8:      Exit Function
9:    End If
10:  Next i
  
12:  Exit Function

Err:
15:  HandleError Err.Number, Err.Description, Erl & "|clsYHighlighter.FindHighlighter()"
End Function

Private Function GetstrExtension(sFileName As String) As String
    On Error GoTo Err
    
3:  Dim lPos As Long
4:  lPos = InStrRev(sFileName, ".")
    
6:  If lPos = 0 Then
7:      GetstrExtension = " "
8:  Else
9:      GetstrExtension = LCase$(Right$(sFileName, Len(sFileName) - lPos))
10: End If
    
12: Exit Function
    
Err:
    HandleError Err.Number, Err.Description, Erl & "|clsYHighlighter.GetstrExtension()"
End Function

Public Function SetHighlighter(cScintilla As clsYScintilla, _
                               strHighlighter As String)
  On Error GoTo Err
  
3:   Dim i As Long, X As Integer
  
4:  X = FindHighlighter(strHighlighter)
5:  cScintilla.StyleClearALL
6:  cScintilla.StartStyle
  
8:  For i = 0 To 127
9:    cScintilla.SetStyleBold i, G_Highlighters(X).StyleBold(i)
10:    cScintilla.SetStyleItalic i, G_Highlighters(X).StyleItalic(i)
11:    cScintilla.SetStyleUnderline i, G_Highlighters(X).StyleUnderline(i)
12:    cScintilla.SetStyleVisible i, G_Highlighters(X).StyleVisible(i)
13:    If G_Highlighters(X).StyleFont(i) <> "" Then cScintilla.SetStyleFont i, G_Highlighters(X).StyleFont(i)
14:    cScintilla.SetStyleFore i, G_Highlighters(X).StyleFore(i)
15:    cScintilla.SetStyleBack i, G_Highlighters(X).StyleBack(i)
16:    cScintilla.SetStyleSize i, G_Highlighters(X).StyleSize(i)
17:    cScintilla.SetStyleEOLFilled i, G_Highlighters(X).StyleEOLFilled(i)
18:  Next i
  
20:  For i = 0 To 7
21:    If G_Highlighters(X).Keywords(i) <> "" Then cScintilla.SetKeywords i, G_Highlighters(X).Keywords(i)
22:  Next i
  
24:  cScintilla.SetLexer G_Highlighters(X).iLang
25:  cScintilla.Colourise
26:  CurHigh = X
  
28:  Exit Function

Err:
30:  HandleError Err.Number, Err.Description, Erl & "|clsYHighlighter.SetHighlighter()"
End Function

Private Function LoadHighlighter(strFile As String)
  On Error GoTo Err
  
3:   Dim fFile As Integer
4:   fFile = FreeFile
  
6:   ReDim Preserve G_Highlighters(0 To HCount + 1)
  
7:   Open strFile For Binary Access Read As #fFile
8:      Get #fFile, , G_Highlighters(HCount)
9:      G_Highlighters(HCount).strName = strFile
10:  Close #fFile
  
11:  FreeFile fFile
12:  HCount = HCount + 1
  
14:  Exit Function

Err:
17:    HandleError Err.Number, Err.Description, Erl & "|clsYHighlighter.LoadHighlighter()"
End Function

Public Sub LoadDirectory(strDir As String)
  On Error GoTo Err
  
3:  Dim str As String, i As Long
4:  hlCount = 0
  
5:  If Right(strDir, 1) <> "\" Then strDir = strDir & "\"
6:  str = Dir(strDir & "\*bin")
7:  Erase G_Highlighters
8:  HCount = 0
  
10:  Do Until str = ""
11:    hlCount = hlCount + 1
12:    LoadHighlighter strDir & "\" & str
13:    str = Dir
14:  Loop
  
16:  Exit Sub

Err:
19:  HandleError Err.Number, Err.Description, Erl & "|clsYHighlighter.LoadDirectory()"
End Sub

Public Sub SetHighlighterBasedOnExt(cScintilla As clsYScintilla, _
                                    strFileName As String)
1:   On Error GoTo Err
2:   Dim strExtension As String, strClrExt As String, X As Long
  
4:   strExtension = LCase$(Mid$(strFileName, InStrRev(strFileName, ".") + 1, Len(strFileName) - InStrRev(strFileName, ".")))
  
6:   For X = 0 To UBound(G_Highlighters)
7:      If InStr(1, G_Highlighters(X).strFilter, strExtension) Then
8:          On Error Resume Next
9:          SetHighlighter cScintilla, G_Highlighters(X).strName
10:     End If
11:  Next X

13:  Exit Sub
14:
Err:
16:  HandleError Err.Number, Err.Description, Erl & "|clsYHighlighter.SetHighlighterBasedOnExt(" & strFileName & ")"
End Sub
