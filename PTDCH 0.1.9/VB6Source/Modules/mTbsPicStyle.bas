Attribute VB_Name = "mTbsPicStyle"
Option Explicit

'----------------------------------------------'
'  Copyright (C) fLaSh - Carlos Ferreira       '
'  E-Mail: Carlosferreiracarlos@hotmail.com    '
'----------------------------------------------'

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

Private Const cFlatB = True

Private Sub FlatBorder(ByVal hWnd As Long)

    If cFlatB Then
2:       On Error GoTo Err
3:       Dim TFlat As Long

5:       TFlat = GetWindowLong(hWnd, GWL_EXSTYLE)
6:       TFlat = TFlat And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
7:       SetWindowLong hWnd, GWL_EXSTYLE, TFlat
8:       SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
      End If
10: Exit Sub
11:
Err:
12:    HandleError Err.Number, Err.Description, Erl & "|" & "mTbsPicStyle.FlatBorder()"
End Sub

Public Sub SetFlatBorder()

2:    On Error GoTo Err
3:    Dim i As Integer
   
5:    With frmHub
6:      For i = 0 To .picTab.count - 1
7:       FlatBorder .picTab(i).hWnd
         PicBkgToTabStrip .picTab(i), .tbsMenu
8:      Next i
        
10:     For i = 0 To .picSTab.count - 1
11:       FlatBorder .picSTab(i).hWnd
          PicBkgToTabStrip .picSTab(i), .tbsSecurity
12:     Next i
     
14:     For i = 0 To .picITab.count - 1
15:       FlatBorder .picITab(i).hWnd
           PicBkgToTabStrip .picITab(i), .tbsInteractions
16:     Next i
     
18:     For i = 0 To .picTabAdv.count - 1
19:       FlatBorder .picTabAdv(i).hWnd
          PicBkgToTabStrip .picTabAdv(i), .tabAdv
20:     Next i
     
22:     For i = 0 To .picHelp.count - 1
24:       FlatBorder .picHelp(i).hWnd
          PicBkgToTabStrip .picHelp(i), .tbsHelp
25:     Next i
     
27:     For i = 0 To .picInfo.count - 1
28:       FlatBorder .picInfo(i).hWnd
29:       PicBkgToTabStrip .picInfo(i), .tbsInfo
30:     Next i

32:     For i = 0 To .picStatus.count - 1
33:       FlatBorder .picStatus(i).hWnd
34:       PicBkgToTabStrip .picStatus(i), .tbsStatus
35:     Next i

37:     FlatBorder .sldPriority.hWnd
38:     FlatBorder .rtbAbout.hWnd
39:     FlatBorder .rtbLog.hWnd
        
41:     FlatBorder .picLog(0).hWnd
42:     FlatBorder .picLog(1).hWnd
        
44:     FlatBorder .picStInfo.hWnd
        
46:     FlatBorder .sldStatus.hWnd
47:     FlatBorder .tlbScript.hWnd
48:     FlatBorder .lvwScripts.hWnd
49:     FlatBorder .dtgSql.hWnd

51:   End With

53:  Exit Sub
54:
Err:
56:  HandleError Err.Number, Err.Description, Erl & "|" & "mTbsPicStyle.SetFlatBorder()"
End Sub

' Shape a picturebox Background to a 5.0 Tabstrip. This is
' useful when you are placing a picturebox control container
' on a Tabstrip, and want to be sure that the picturebox will
' fill the tabstrip body.
Private Sub PicBkgToTabStrip(pBackground As PictureBox, TbStrip As Object)
1:  pBackground.Left = TbStrip.Left + 80 '15         'right of left border
2:  pBackground.Width = TbStrip.Width - 170 '60      'keep inside right border
3:  pBackground.Top = TbStrip.Top + 360  '330        'below top border
4:  pBackground.Height = TbStrip.Height - 455 '375   'above bottom border
End Sub

'Repaint tiling
Public Sub PaintTileFormBackground(MyForm As Form, MyPicture As IPictureDisp)
1:  Dim i As Long, j As Long
2:    For i = 0 To MyForm.ScaleWidth Step 1770      'Used original image size .. draw across top
3:      For j = 0 To MyForm.ScaleHeight Step 2070   'Used original image size .. draw across height
4:        MyForm.PaintPicture MyPicture, i, j       'draw a frame
5:      Next j
6:    Next i
End Sub

'Repaint tiling
Public Sub PaintTilePicBackground(Mypic As PictureBox, MyPicture As IPictureDisp)
1:  Dim i As Long, j As Long
2:    For i = 0 To Mypic.ScaleWidth Step 1770     'Used original image size .. draw across top
3:      For j = 0 To Mypic.ScaleHeight Step 2070  'Used original image size .. draw across height
4:        Mypic.PaintPicture MyPicture, i, j       'draw a frame
5:      Next j
6:    Next i
End Sub

Public Function iResPic(iID As Integer) As IPictureDisp
1: On Error GoTo Err
'2:  If iID > 17 Or iID = 0 Then iID = 1
    'This is the picture function. that sends back an IPictureDisp picture
    ' This picture is then used for themes load from the PTDCH.RES
5:  Set iResPic = LoadResPicture(iID, vbResBitmap)
6: Exit Function
7:
Err:
8: HandleError Err.Number, Err.Description, Erl & "|" & "mTbsPicStyle.iResPic(" & iID & ")"
End Function
