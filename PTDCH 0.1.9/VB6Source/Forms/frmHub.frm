VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmHub 
   Appearance      =   0  'Flat
   Caption         =   "PT DC Hub x.x.x"
   ClientHeight    =   4935
   ClientLeft      =   2550
   ClientTop       =   2760
   ClientWidth     =   9285
   Icon            =   "frmHub.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   9285
   Begin ComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   152
      Top             =   4680
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   9
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1588
            Text            =   "00:00:00"
            TextSave        =   "00:00:00"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "[M:Months](W:Weeks)(D:Days) Hours:Minutes:Seconds"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            TextSave        =   "22:12"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Time"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "19-12-2007"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Date"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "0"
            TextSave        =   "0"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Users Online"
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Text            =   "0 Bytes"
            TextSave        =   "0 Bytes"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Shared total"
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "0"
            TextSave        =   "0"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Op Online"
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            Text            =   "-- : -- %"
            TextSave        =   "-- : -- %"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Memory Used"
         EndProperty
         BeginProperty Panel8 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Hub IP"
         EndProperty
         BeginProperty Panel9 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "DSN Status"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBordTab 
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   9170
      ScaleHeight     =   260
      ScaleMode       =   0  'User
      ScaleWidth      =   300
      TabIndex        =   154
      Top             =   60
      Width           =   300
   End
   Begin VB.PictureBox picHideObj 
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   3315
      TabIndex        =   435
      Top             =   4680
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Timer tmrScriptTimer 
         Enabled         =   0   'False
         Index           =   0
         Left            =   960
         Top             =   0
      End
      Begin VB.Timer tmrBackground 
         Enabled         =   0   'False
         Left            =   0
         Top             =   0
      End
      Begin VB.Timer tmrSysInfo 
         Interval        =   1000
         Left            =   480
         Top             =   0
      End
      Begin MSScriptControlCtl.ScriptControl ScriptControl 
         Index           =   0
         Left            =   1440
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
      End
      Begin MSWinsockLib.Winsock wskRegister 
         Index           =   0
         Left            =   480
         Top             =   480
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wskLoop 
         Index           =   0
         Left            =   0
         Top             =   480
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wskListen 
         Index           =   0
         Left            =   960
         Top             =   480
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin ComctlLib.ImageList imlScripts 
         Left            =   2640
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16711935
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   17
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":15162
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":154B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":15806
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":15B58
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":15EAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":161FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":1654E
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":168A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":16BF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":16F44
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":17296
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":175E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":1793A
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":17C8C
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":17FDE
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":18330
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmHub.frx":18682
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin ComctlLib.ImageList imlAddIns 
         Left            =   2040
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   16711935
         _Version        =   327682
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   4115
      Index           =   0
      Left            =   120
      ScaleHeight     =   4110
      ScaleWidth      =   9015
      TabIndex        =   235
      Top             =   420
      Width           =   9015
      Begin VB.PictureBox picLog 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00C0C0C0&
         ForeColor       =   &H00C0C0C0&
         Height          =   2565
         Index           =   0
         Left            =   4560
         ScaleHeight     =   2565
         ScaleWidth      =   4095
         TabIndex        =   430
         Top             =   480
         Width           =   4095
         Begin VB.Label LblShadowed 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Produced bY fLaSh"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   210
            Index           =   2
            Left            =   1080
            TabIndex        =   432
            Top             =   600
            Width           =   1785
         End
         Begin VB.Label LblShadowed 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PT Direct Connect Hub"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   210
            Index           =   0
            Left            =   840
            TabIndex        =   431
            Top             =   360
            Width           =   2205
         End
         Begin VB.Label LblShadowed 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PT Direct Connect Hub"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   1
            Left            =   855
            TabIndex        =   433
            Top             =   375
            Width           =   2205
         End
         Begin VB.Label LblShadowed 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Produced bY fLaSh"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   3
            Left            =   1095
            TabIndex        =   434
            Top             =   615
            Width           =   1785
         End
      End
      Begin VB.TextBox txtUpTime 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   236
         Text            =   "00:00:00"
         Top             =   3525
         Width           =   1815
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   36
         Left            =   1920
         TabIndex        =   7
         Tag             =   "RedirectAddress"
         ToolTipText     =   "Seperate addresses with a semicolon (;)"
         Top             =   3360
         Width           =   2175
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   4
         Left            =   1920
         TabIndex        =   6
         Tag             =   "RegisterIP"
         ToolTipText     =   "Seperate addresses with a semicolon (;)"
         Top             =   3000
         Width           =   2175
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   3
         Left            =   1920
         TabIndex        =   5
         Tag             =   "Ports"
         ToolTipText     =   "Seperate ports with a semicolon (;) (First port in the list is the one used for registration purposes)"
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   2
         Left            =   1920
         TabIndex        =   4
         Tag             =   "HubIP"
         ToolTipText     =   "The address for your hub"
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   1
         Left            =   1920
         MaxLength       =   140
         TabIndex        =   3
         Tag             =   "HubDesc"
         ToolTipText     =   "A short description of your hub (140 characters max)"
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   0
         Left            =   1920
         MaxLength       =   70
         TabIndex        =   2
         Tag             =   "HubName"
         ToolTipText     =   "Name of your hub (70 characters max)"
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Start Server"
         Height          =   375
         Index           =   1
         Left            =   6840
         TabIndex        =   0
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   5
         X1              =   4440
         X2              =   4800
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   4
         X1              =   4440
         X2              =   8760
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   3
         X1              =   4440
         X2              =   4440
         Y1              =   3120
         Y2              =   360
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to the world of"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   383
         Top             =   240
         Width           =   3735
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   11
         X1              =   6720
         X2              =   7080
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   10
         X1              =   6720
         X2              =   8760
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   9
         X1              =   6720
         X2              =   6720
         Y1              =   3960
         Y2              =   3360
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Hub Control"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   3
         Left            =   7200
         TabIndex        =   245
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   8
         X1              =   4440
         X2              =   4800
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   7
         X1              =   4440
         X2              =   6480
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   6
         X1              =   4440
         X2              =   4440
         Y1              =   3960
         Y2              =   3360
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "UpTime"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   244
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   2
         X1              =   120
         X2              =   480
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   1
         X1              =   120
         X2              =   120
         Y1              =   3960
         Y2              =   240
      End
      Begin VB.Line Line 
         BorderColor     =   &H00C0C0C0&
         Index           =   0
         X1              =   120
         X2              =   4200
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label lblHolder 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Redirect Address"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   99
         Left            =   240
         TabIndex        =   243
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label lblHolder 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Register Address"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   242
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label lblHolder 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Listening Ports"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   241
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblHolder 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   240
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblHolder 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   239
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblHolder 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   238
         Top             =   600
         Width           =   1575
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H00C0C0C0&
         Height          =   1095
         Index           =   1
         Left            =   240
         Top             =   1560
         Width           =   3975
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H00C0C0C0&
         Height          =   1095
         Index           =   2
         Left            =   240
         Top             =   2760
         Width           =   3975
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H00C0C0C0&
         Height          =   1095
         Index           =   0
         Left            =   240
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Hub Settings"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   237
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   4115
      Index           =   8
      Left            =   120
      ScaleHeight     =   4110
      ScaleWidth      =   9015
      TabIndex        =   230
      Top             =   420
      Visible         =   0   'False
      Width           =   9015
      Begin VB.PictureBox picBordTab 
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   1
         Left            =   3120
         ScaleHeight     =   260
         ScaleMode       =   0  'User
         ScaleWidth      =   5760
         TabIndex        =   234
         Top             =   60
         Width           =   5760
      End
      Begin VB.PictureBox picHelp 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   0
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8715
         TabIndex        =   231
         Top             =   420
         Width           =   8715
         Begin VB.PictureBox picLog 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            FillColor       =   &H00C0C0C0&
            ForeColor       =   &H00C0C0C0&
            Height          =   465
            Index           =   2
            Left            =   240
            MouseIcon       =   "frmHub.frx":189D4
            MousePointer    =   99  'Custom
            ScaleHeight     =   465
            ScaleWidth      =   930
            TabIndex        =   421
            Top             =   2850
            Width           =   930
         End
         Begin VB.PictureBox picLog 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            FillColor       =   &H00C0C0C0&
            ForeColor       =   &H00C0C0C0&
            Height          =   2565
            Index           =   1
            Left            =   120
            ScaleHeight     =   2565
            ScaleWidth      =   2475
            TabIndex        =   380
            Top             =   120
            Width           =   2475
         End
         Begin RichTextLib.RichTextBox rtbAbout 
            Height          =   2895
            Left            =   2640
            TabIndex        =   450
            Top             =   120
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   5106
            _Version        =   393217
            BackColor       =   16777215
            BorderStyle     =   0
            Enabled         =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmHub.frx":18B26
         End
         Begin VB.Label LabelsURL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   0
            Left            =   1200
            MouseIcon       =   "frmHub.frx":18BA8
            MousePointer    =   99  'Custom
            TabIndex        =   382
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Label LabelsURL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Home Page"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   1
            Left            =   1200
            MouseIcon       =   "frmHub.frx":18CFA
            MousePointer    =   99  'Custom
            TabIndex        =   381
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   645
            Index           =   4
            Left            =   120
            Top             =   2760
            Width           =   2465
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   285
            Index           =   23
            Left            =   2640
            Top             =   3120
            Width           =   5895
         End
         Begin VB.Label LblShadowed 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Direct Connect P2P Network"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Index           =   4
            Left            =   3960
            TabIndex        =   232
            Top             =   3120
            Width           =   3120
         End
         Begin VB.Label LblShadowed 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Direct Connect P2P Network"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   5
            Left            =   3975
            TabIndex        =   233
            Top             =   3135
            Width           =   3120
         End
      End
      Begin VB.PictureBox picHelp 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   1
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8715
         TabIndex        =   346
         Top             =   420
         Visible         =   0   'False
         Width           =   8715
         Begin VB.TextBox txtNotePad 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   3285
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   153
            Top             =   120
            Width           =   8535
         End
      End
      Begin ComctlLib.TabStrip tbsHelp 
         Height          =   3945
         Left            =   60
         TabIndex        =   151
         Top             =   60
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   6959
         TabWidthStyle   =   2
         TabFixedWidth   =   2646
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   2
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "About PTDCH"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "NotePad"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   4115
      Index           =   6
      Left            =   120
      ScaleHeight     =   4110
      ScaleWidth      =   9015
      TabIndex        =   229
      Top             =   420
      Visible         =   0   'False
      Width           =   9015
      Begin VB.PictureBox picBordTab 
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   6
         Left            =   7620
         ScaleHeight     =   260
         ScaleMode       =   0  'User
         ScaleWidth      =   1380
         TabIndex        =   378
         Top             =   60
         Width           =   1380
      End
      Begin VB.PictureBox picStatus 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   0
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8655
         TabIndex        =   384
         Top             =   430
         Width           =   8655
         Begin VB.TextBox txtStForm 
            Height          =   285
            Left            =   60
            MaxLength       =   40
            TabIndex        =   396
            Top             =   3070
            Width           =   1030
         End
         Begin VB.PictureBox picStInfo 
            BorderStyle     =   0  'None
            Height          =   1935
            Left            =   5810
            ScaleHeight     =   1935
            ScaleWidth      =   2775
            TabIndex        =   390
            Top             =   900
            Visible         =   0   'False
            Width           =   2775
            Begin VB.Label lblStatus 
               BackStyle       =   0  'Transparent
               Caption         =   "6 = Send PM To UnRegistered"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   498
               Top             =   1560
               Width           =   2535
            End
            Begin VB.Label lblStatus 
               BackStyle       =   0  'Transparent
               Caption         =   "4 = Send PM To All"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   497
               Top             =   1080
               Width           =   2535
            End
            Begin VB.Label lblStatus 
               BackStyle       =   0  'Transparent
               Caption         =   "5 = Send PM To Op"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   6
               Left            =   120
               TabIndex        =   395
               Top             =   1320
               Width           =   2535
            End
            Begin VB.Label lblStatus 
               BackStyle       =   0  'Transparent
               Caption         =   "3 = Send Chat To UnRegistered"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   394
               Top             =   840
               Width           =   2535
            End
            Begin VB.Label lblStatus 
               BackStyle       =   0  'Transparent
               Caption         =   "2 = Send Chat To Op"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   393
               Top             =   600
               Width           =   2535
            End
            Begin VB.Label lblStatus 
               BackStyle       =   0  'Transparent
               Caption         =   "1 = Send Chat To All"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   392
               Top             =   360
               Width           =   2535
            End
            Begin VB.Label lblStatus 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "------------- Send Chat -------------"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   391
               Top             =   120
               Width           =   2535
            End
         End
         Begin VB.CommandButton cmdStSend 
            Caption         =   "Send"
            Height          =   495
            Left            =   5160
            TabIndex        =   400
            Top             =   2880
            Width           =   855
         End
         Begin VB.OptionButton optStSend 
            Caption         =   "Data"
            Height          =   195
            Index           =   1
            Left            =   6120
            TabIndex        =   402
            Top             =   3150
            Width           =   195
         End
         Begin VB.OptionButton optStSend 
            Caption         =   "Chat"
            Height          =   195
            Index           =   0
            Left            =   6120
            TabIndex        =   401
            Top             =   2920
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.TextBox txtStSend 
            Height          =   285
            Left            =   1140
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   398
            Top             =   3070
            Width           =   3975
         End
         Begin VB.ListBox lstStatus 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2790
            Index           =   0
            Left            =   60
            TabIndex        =   385
            Top             =   60
            Width           =   8535
         End
         Begin ComctlLib.Slider sldStatus 
            Height          =   315
            Left            =   6990
            TabIndex        =   429
            Tag             =   "PriorityVal"
            Top             =   2880
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Min             =   1
            Max             =   6
            SelectRange     =   -1  'True
            SelStart        =   1
            Value           =   1
         End
         Begin VB.Label lblOptStSend 
            BackStyle       =   0  'Transparent
            Caption         =   "Data"
            Height          =   255
            Index           =   1
            Left            =   6360
            TabIndex        =   423
            Top             =   3150
            Width           =   615
         End
         Begin VB.Label lblOptStSend 
            BackStyle       =   0  'Transparent
            Caption         =   "Chat"
            Height          =   255
            Index           =   0
            Left            =   6360
            TabIndex        =   422
            Top             =   2925
            Width           =   615
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "Message or Data"
            Height          =   255
            Index           =   26
            Left            =   1140
            TabIndex        =   399
            Top             =   2880
            Width           =   2655
         End
         Begin VB.Label Labels 
            BackStyle       =   0  'Transparent
            Caption         =   "Form"
            Height          =   255
            Index           =   9
            Left            =   60
            TabIndex        =   397
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label lblStatus 
            BackStyle       =   0  'Transparent
            Caption         =   "1    2    3    4    5    6"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   0
            Left            =   7050
            TabIndex        =   389
            Top             =   3240
            Width           =   1575
         End
      End
      Begin VB.PictureBox picStatus 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   4
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8655
         TabIndex        =   404
         Top             =   430
         Visible         =   0   'False
         Width           =   8655
         Begin ComctlLib.ListView lvwUsers 
            Height          =   3255
            Left            =   3840
            TabIndex        =   405
            Top             =   120
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   5741
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "IP"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   2
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Winsock Index"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   2
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Connected Since"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   300
            Index           =   24
            Left            =   120
            Top             =   120
            Width           =   3615
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Statistics"
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   418
            Top             =   180
            Width           =   3135
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   1095
            Index           =   22
            Left            =   120
            Top             =   1920
            Width           =   3615
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   1095
            Index           =   21
            Left            =   120
            Top             =   600
            Width           =   3615
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Connected users :"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   45
            Left            =   240
            TabIndex        =   417
            Top             =   810
            Width           =   1815
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Connected operators :"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   47
            Left            =   240
            TabIndex        =   416
            Top             =   1050
            Width           =   1815
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Total shared :"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   51
            Left            =   240
            TabIndex        =   415
            Top             =   1290
            Width           =   1815
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Peak users :"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   52
            Left            =   240
            TabIndex        =   414
            Top             =   2130
            Width           =   1815
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Peak operators :"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   53
            Left            =   240
            TabIndex        =   413
            Top             =   2370
            Width           =   1815
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Peak shared :"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   54
            Left            =   240
            TabIndex        =   412
            Top             =   2610
            Width           =   1815
         End
         Begin VB.Label lblHolder 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   55
            Left            =   2160
            TabIndex        =   411
            Top             =   810
            Width           =   1575
         End
         Begin VB.Label lblHolder 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   56
            Left            =   2160
            TabIndex        =   410
            Top             =   1050
            Width           =   1575
         End
         Begin VB.Label lblHolder 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "0 Bytes"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   57
            Left            =   2160
            TabIndex        =   409
            Top             =   1290
            Width           =   1575
         End
         Begin VB.Label lblHolder 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   58
            Left            =   2160
            TabIndex        =   408
            Top             =   2130
            Width           =   1575
         End
         Begin VB.Label lblHolder 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   59
            Left            =   2160
            TabIndex        =   407
            Top             =   2370
            Width           =   1575
         End
         Begin VB.Label lblHolder 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "0 Bytes"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   60
            Left            =   2160
            TabIndex        =   406
            Top             =   2610
            Width           =   1575
         End
      End
      Begin VB.PictureBox picStatus 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   1
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8655
         TabIndex        =   388
         Top             =   430
         Visible         =   0   'False
         Width           =   8655
         Begin VB.ListBox lstStatus 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3375
            Index           =   1
            ItemData        =   "frmHub.frx":18E4C
            Left            =   60
            List            =   "frmHub.frx":18E4E
            TabIndex        =   403
            Top             =   60
            Width           =   8535
         End
      End
      Begin VB.PictureBox picStatus 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   3
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8655
         TabIndex        =   386
         Top             =   430
         Visible         =   0   'False
         Width           =   8655
         Begin RichTextLib.RichTextBox rtbLog 
            DataSource      =   "(None)"
            Height          =   3375
            Left            =   60
            TabIndex        =   387
            Top             =   60
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   5953
            _Version        =   393217
            BorderStyle     =   0
            Enabled         =   -1  'True
            ScrollBars      =   3
            Appearance      =   0
            RightMargin     =   1e7
            OLEDragMode     =   0
            OLEDropMode     =   0
            TextRTF         =   $"frmHub.frx":18E50
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.PictureBox picStatus 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   2
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8655
         TabIndex        =   419
         Top             =   430
         Visible         =   0   'False
         Width           =   8655
         Begin VB.ListBox lstStatus 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3375
            Index           =   2
            ItemData        =   "frmHub.frx":18ECB
            Left            =   60
            List            =   "frmHub.frx":18ECD
            TabIndex        =   420
            Top             =   60
            Width           =   8535
         End
      End
      Begin ComctlLib.TabStrip tbsStatus 
         Height          =   3945
         Left            =   60
         TabIndex        =   379
         Top             =   60
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   6959
         TabWidthStyle   =   2
         TabFixedWidth   =   2646
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   5
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Main Chat Log"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "PM Log"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Protocol Misc Log"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "System Log"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Status"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   4115
      Index           =   7
      Left            =   120
      ScaleHeight     =   4110
      ScaleWidth      =   9015
      TabIndex        =   228
      Top             =   420
      Visible         =   0   'False
      Width           =   9015
      Begin VB.PictureBox picBordTab 
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   5
         Left            =   4440
         ScaleHeight     =   260
         ScaleMode       =   0  'User
         ScaleWidth      =   4605
         TabIndex        =   371
         Top             =   60
         Width           =   4600
      End
      Begin VB.PictureBox picInfo 
         BorderStyle     =   0  'None
         Height          =   3400
         Index           =   0
         Left            =   180
         ScaleHeight     =   3405
         ScaleWidth      =   8655
         TabIndex        =   348
         Top             =   480
         Width           =   8655
         Begin VB.CommandButton cmdButton 
            Caption         =   "Check updates"
            Height          =   525
            Index           =   0
            Left            =   5640
            TabIndex        =   491
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CommandButton cmdButton 
            Caption         =   "Detect IP"
            Height          =   525
            Index           =   5
            Left            =   7080
            TabIndex        =   490
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CommandButton cmdButton 
            Caption         =   "Reload Settings"
            Height          =   525
            Index           =   4
            Left            =   7080
            TabIndex        =   488
            Top             =   2520
            Width           =   1335
         End
         Begin VB.CommandButton cmdButton 
            Caption         =   "Save Settings"
            Height          =   525
            Index           =   3
            Left            =   5640
            TabIndex        =   487
            Top             =   2520
            Width           =   1335
         End
         Begin VB.ComboBox cmbInterface 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5640
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   485
            Tag             =   "Interface"
            ToolTipText     =   "Set Interface Language for DDCH"
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox txtSystem 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   355
            Text            =   "512.000 Kb"
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtSystem 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   354
            Text            =   "512.000 Kb"
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtSystem 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   3660
            Locked          =   -1  'True
            TabIndex        =   353
            Text            =   "1500.000 Kb"
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtSystem 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   3660
            Locked          =   -1  'True
            TabIndex        =   352
            Text            =   "1500.000 Kb"
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtSystem 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   5
            Left            =   3660
            Locked          =   -1  'True
            TabIndex        =   351
            Text            =   "1500.000 Kbs"
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox txtSystem 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   4
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   350
            Text            =   "1500.000 Kbs"
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox txtSystem 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   349
            Top             =   2040
            Width           =   4695
         End
         Begin ComctlLib.ProgressBar pgrMemory 
            Height          =   300
            Left            =   360
            TabIndex        =   366
            ToolTipText     =   "Psysical Memory"
            Top             =   2760
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   529
            _Version        =   327682
            Appearance      =   0
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   18
            X1              =   5520
            X2              =   5880
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   19
            X1              =   5520
            X2              =   5520
            Y1              =   2040
            Y2              =   1320
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   20
            X1              =   5520
            X2              =   8400
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "UpDate"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   6
            Left            =   6000
            TabIndex        =   492
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   15
            X1              =   5520
            X2              =   5880
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   16
            X1              =   5520
            X2              =   5520
            Y1              =   3120
            Y2              =   2400
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   17
            X1              =   5520
            X2              =   8400
            Y1              =   3120
            Y2              =   3120
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Settings"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   5
            Left            =   6000
            TabIndex        =   489
            Top             =   2280
            Width           =   2295
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   12
            X1              =   5520
            X2              =   5880
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   13
            X1              =   5520
            X2              =   5520
            Y1              =   840
            Y2              =   240
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   14
            X1              =   5520
            X2              =   8400
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Interface"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   4
            Left            =   6000
            TabIndex        =   486
            Top             =   120
            Width           =   2295
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   47
            X1              =   240
            X2              =   5160
            Y1              =   3120
            Y2              =   3120
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   46
            X1              =   240
            X2              =   600
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   45
            X1              =   240
            X2              =   240
            Y1              =   3120
            Y2              =   2640
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Memory Usage: ---"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   13
            Left            =   720
            TabIndex        =   367
            Top             =   2520
            Width           =   2175
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   27
            X1              =   240
            X2              =   600
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   28
            X1              =   240
            X2              =   240
            Y1              =   1080
            Y2              =   240
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   29
            X1              =   240
            X2              =   2520
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Psysical Memory (K)"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   9
            Left            =   720
            TabIndex        =   365
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label Labels 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   2
            Left            =   360
            TabIndex        =   364
            Top             =   360
            Width           =   675
         End
         Begin VB.Label Labels 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Avaible:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   3
            Left            =   360
            TabIndex        =   363
            Top             =   720
            Width           =   630
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   30
            X1              =   2880
            X2              =   2880
            Y1              =   1080
            Y2              =   240
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   31
            X1              =   2880
            X2              =   5160
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Virtual Memory (K)"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   10
            Left            =   3360
            TabIndex        =   362
            Top             =   120
            Width           =   1695
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   32
            X1              =   2880
            X2              =   3240
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Label Labels 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   7
            Left            =   3000
            TabIndex        =   361
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Labels 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Avaible:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   8
            Left            =   3000
            TabIndex        =   360
            Top             =   720
            Width           =   630
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   33
            X1              =   240
            X2              =   240
            Y1              =   1680
            Y2              =   1200
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   34
            X1              =   240
            X2              =   5160
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Page File Usage (K)"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   11
            Left            =   720
            TabIndex        =   359
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   35
            X1              =   240
            X2              =   600
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label Labels 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Avaible:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   1
            Left            =   2640
            TabIndex        =   358
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Labels 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   0
            Left            =   360
            TabIndex        =   357
            Top             =   1320
            Width           =   615
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   36
            X1              =   240
            X2              =   600
            Y1              =   1920
            Y2              =   1920
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   37
            X1              =   240
            X2              =   240
            Y1              =   2400
            Y2              =   1920
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   38
            X1              =   240
            X2              =   5160
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Processor"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   12
            Left            =   720
            TabIndex        =   356
            Top             =   1800
            Width           =   1695
         End
      End
      Begin VB.PictureBox picInfo 
         BorderStyle     =   0  'None
         Height          =   3400
         Index           =   2
         Left            =   180
         ScaleHeight     =   3405
         ScaleWidth      =   8655
         TabIndex        =   372
         Top             =   480
         Visible         =   0   'False
         Width           =   8655
         Begin VB.CommandButton cmdButton 
            Caption         =   "UnRegistereds"
            Enabled         =   0   'False
            Height          =   375
            Index           =   9
            Left            =   240
            TabIndex        =   496
            Top             =   2520
            Width           =   1935
         End
         Begin VB.CommandButton cmdButton 
            Caption         =   "All"
            Enabled         =   0   'False
            Height          =   375
            Index           =   7
            Left            =   240
            TabIndex        =   494
            Top             =   1560
            Width           =   1935
         End
         Begin VB.CommandButton cmdButton 
            Caption         =   "Operatores"
            Enabled         =   0   'False
            Height          =   375
            Index           =   8
            Left            =   240
            TabIndex        =   493
            Top             =   2040
            Width           =   1935
         End
         Begin VB.CommandButton cmdConvDatabase 
            Caption         =   "Convert database"
            Height          =   375
            Left            =   240
            TabIndex        =   373
            Top             =   480
            Width           =   1935
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   21
            X1              =   120
            X2              =   480
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   22
            X1              =   120
            X2              =   120
            Y1              =   3120
            Y2              =   1440
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   23
            X1              =   120
            X2              =   5160
            Y1              =   3120
            Y2              =   3120
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Mas Messages"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   7
            Left            =   600
            TabIndex        =   495
            Top             =   1320
            Width           =   3615
         End
         Begin VB.Label Labels 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Database (*.XML) of the YnHub or PtokaX."
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   375
            Top             =   930
            Width           =   5055
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Convert Accounts for PTDCH database"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   8
            Left            =   600
            TabIndex        =   374
            Top             =   240
            Width           =   4815
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   26
            X1              =   120
            X2              =   5160
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   25
            X1              =   120
            X2              =   120
            Y1              =   1200
            Y2              =   360
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   24
            X1              =   120
            X2              =   480
            Y1              =   360
            Y2              =   360
         End
      End
      Begin VB.PictureBox picInfo 
         BorderStyle     =   0  'None
         Height          =   3400
         Index           =   1
         Left            =   180
         ScaleHeight     =   3405
         ScaleWidth      =   8655
         TabIndex        =   368
         Top             =   480
         Visible         =   0   'False
         Width           =   8655
         Begin VB.CheckBox chkData 
            Caption         =   "Enabled Plugins"
            Height          =   195
            Index           =   68
            Left            =   8280
            TabIndex        =   376
            Tag             =   "Plugins"
            ToolTipText     =   "This option requests restart the application.."
            Top             =   3050
            Width           =   195
         End
         Begin VB.CommandButton cmdPlugins 
            Caption         =   "Setup"
            Enabled         =   0   'False
            Height          =   300
            Left            =   120
            TabIndex        =   370
            Top             =   3000
            Width           =   1215
         End
         Begin ComctlLib.ListView lvwPlugins 
            Height          =   2880
            Left            =   60
            TabIndex        =   369
            Top             =   60
            Width           =   8445
            _ExtentX        =   14896
            _ExtentY        =   5080
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   327682
            Icons           =   "imlAddIns"
            SmallIcons      =   "imlAddIns"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   7
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   ""
               Object.Width           =   176
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Name"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   2
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Version"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   2
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Author"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   4
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Description"
               Object.Width           =   7585
            EndProperty
            BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   2
               SubItemIndex    =   5
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Release"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   6
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Comments"
               Object.Width           =   3528
            EndProperty
         End
         Begin VB.Label lblCheck 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Enabled Plugins"
            Height          =   255
            Index           =   68
            Left            =   6600
            TabIndex        =   377
            ToolTipText     =   "This option requests restart the application.."
            Top             =   3045
            Width           =   1575
         End
      End
      Begin ComctlLib.TabStrip tbsInfo 
         Height          =   3975
         Left            =   60
         TabIndex        =   347
         Top             =   60
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7011
         TabWidthStyle   =   2
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   3
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Memory Info"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Plugins Info"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Others"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   4115
      Index           =   5
      Left            =   120
      ScaleHeight     =   4110
      ScaleWidth      =   9015
      TabIndex        =   227
      TabStop         =   0   'False
      Top             =   420
      Visible         =   0   'False
      Width           =   9015
      Begin ComctlLib.Toolbar tlbScript 
         Height          =   390
         Left            =   60
         TabIndex        =   451
         Top             =   60
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         ImageList       =   "imlScripts"
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   21
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Undo"
               Object.ToolTipText     =   "Undo"
               Object.Tag             =   ""
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Redo"
               Object.ToolTipText     =   "Redo"
               Object.Tag             =   ""
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Find"
               Object.ToolTipText     =   "Find"
               Object.Tag             =   ""
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Replace"
               Object.ToolTipText     =   "Replace"
               Object.Tag             =   ""
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "GoToLine"
               Object.ToolTipText     =   "Go To Line"
               Object.Tag             =   ""
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Save Only"
               Object.ToolTipText     =   "Save Only"
               Object.Tag             =   ""
               ImageIndex      =   6
            EndProperty
            BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Save and Reset Script"
               Object.ToolTipText     =   "Save and Reset Script"
               Object.Tag             =   "9"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Clear"
               Object.ToolTipText     =   "Clear"
               Object.Tag             =   ""
               ImageIndex      =   8
            EndProperty
            BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Hide Scripts"
               Object.ToolTipText     =   "Hide Scripts"
               Object.Tag             =   ""
               ImageIndex      =   9
               Style           =   1
            EndProperty
            BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Hide Tabs"
               Object.ToolTipText     =   "Hide Tabs"
               Object.Tag             =   ""
               ImageIndex      =   10
               Style           =   1
            EndProperty
            BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Enabled Tabs"
               Object.ToolTipText     =   "Enabled Tabs"
               Object.Tag             =   ""
               ImageIndex      =   11
               Style           =   1
            EndProperty
            BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "New"
               Object.ToolTipText     =   "New"
               Object.Tag             =   ""
               ImageIndex      =   16
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   3
               MixedState      =   -1  'True
            EndProperty
            BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "Menu"
               Object.ToolTipText     =   "Menu"
               Object.Tag             =   ""
               ImageIndex      =   17
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvwScripts 
         Height          =   3495
         Left            =   7440
         TabIndex        =   500
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   6165
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "State"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Script Type"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Modified"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.TextBox txtScriptError 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   60
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   449
         Top             =   3720
         Width           =   8895
      End
      Begin VB.PictureBox picSciMain 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2655
         Index           =   0
         Left            =   120
         ScaleHeight     =   2655
         ScaleWidth      =   7155
         TabIndex        =   448
         Top             =   840
         Visible         =   0   'False
         Width           =   7155
      End
      Begin ComctlLib.TabStrip tbsScripts 
         Height          =   3135
         Left            =   60
         TabIndex        =   499
         Top             =   480
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   5530
         TabWidthStyle   =   2
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   1
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "New Script.vbs"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   4115
      Index           =   4
      Left            =   120
      ScaleHeight     =   4110
      ScaleWidth      =   9015
      TabIndex        =   205
      Top             =   420
      Visible         =   0   'False
      Width           =   9015
      Begin VB.CommandButton cmdButton 
         Caption         =   "Redirect users"
         Enabled         =   0   'False
         Height          =   465
         Index           =   2
         Left            =   5520
         TabIndex        =   150
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "Allow ops to redirect (admins unaffected)"
         BeginProperty DataFormat 
            Type            =   4
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   8
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   30
         Left            =   3720
         TabIndex        =   149
         Tag             =   "OpsCanRedirect"
         ToolTipText     =   "Check to allow operators to use the redirect ability (admins can always redirect)"
         Top             =   3000
         Width           =   195
      End
      Begin VB.OptionButton optRedirect 
         Appearance      =   0  'Flat
         Caption         =   "All users"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   3840
         TabIndex        =   143
         ToolTipText     =   "Redirects all users to your redirect address"
         Top             =   1200
         Width           =   195
      End
      Begin VB.OptionButton optRedirect 
         Appearance      =   0  'Flat
         Caption         =   "Do not redirect"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   3840
         TabIndex        =   148
         ToolTipText     =   "Do not redirect anyone (disconnects if the hub is full)"
         Top             =   2400
         Width           =   195
      End
      Begin VB.OptionButton optRedirect 
         Appearance      =   0  'Flat
         Caption         =   "Only if full and the user is not registered"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   3840
         TabIndex        =   147
         ToolTipText     =   "Only redirect if the user is not registered and the hub is full"
         Top             =   2160
         Width           =   195
      End
      Begin VB.OptionButton optRedirect 
         Appearance      =   0  'Flat
         Caption         =   "Only if full and the user is not an op"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   3840
         TabIndex        =   146
         ToolTipText     =   "Only redirect if the user is an operator (or of a higher class) and the hub is full"
         Top             =   1920
         Width           =   195
      End
      Begin VB.OptionButton optRedirect 
         Appearance      =   0  'Flat
         Caption         =   "Only if full"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   3840
         TabIndex        =   145
         ToolTipText     =   "Only redirect users if the hub is full"
         Top             =   1680
         Width           =   195
      End
      Begin VB.OptionButton optRedirect 
         Appearance      =   0  'Flat
         Caption         =   "Unregistered users"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   3840
         TabIndex        =   144
         ToolTipText     =   "Only unregistered users are redirected (must have ""Allow only registered users"" in Security/Advanced checked)"
         Top             =   1440
         Width           =   195
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   32
         Left            =   1560
         TabIndex        =   124
         Tag             =   "ForTooOldDcppRedirectAddress"
         Top             =   2655
         Width           =   2000
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "Too Old DC++"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   44
         Left            =   120
         TabIndex        =   137
         Tag             =   "RedirectFTooOldDCpp"
         Top             =   2640
         Width           =   195
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   28
         Left            =   1560
         TabIndex        =   125
         Tag             =   "ForTooOldNMDCRedirectAddress"
         Top             =   3015
         Width           =   2000
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "Too Old NMDC"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   40
         Left            =   120
         TabIndex        =   138
         Tag             =   "RedirectFTooOldNMDC"
         Top             =   3000
         Width           =   195
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   31
         Left            =   1560
         TabIndex        =   123
         Tag             =   "ForSlotPerHubRedirectAddress"
         Top             =   2280
         Width           =   2000
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "For Slot / Hub"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   43
         Left            =   120
         TabIndex        =   136
         Tag             =   "RedirectFSlotPerHub"
         Top             =   2260
         Width           =   195
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "For KB / Slot"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   45
         Left            =   120
         TabIndex        =   139
         Tag             =   "RedirectFBWPerSlot"
         Top             =   3360
         Width           =   195
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   33
         Left            =   1560
         TabIndex        =   126
         Tag             =   "ForBWPerSlotRedirectAddress"
         Top             =   3375
         Width           =   2000
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   35
         Left            =   1560
         TabIndex        =   127
         Tag             =   "ForFakeShareRedirectAddress"
         Top             =   3720
         Width           =   2000
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   34
         Left            =   5160
         TabIndex        =   128
         Tag             =   "ForFakeTagRedirectAddress"
         Top             =   120
         Width           =   2000
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "For Fake Tag"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   46
         Left            =   3720
         TabIndex        =   141
         Tag             =   "RedirectFFakeTag"
         Top             =   105
         Width           =   195
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "For Fake Share"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   47
         Left            =   120
         TabIndex        =   140
         Tag             =   "RedirectFFakeShare"
         Top             =   3705
         Width           =   195
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "Passive Mode"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   49
         Left            =   3720
         TabIndex        =   142
         Tag             =   "RedirectFPasMode"
         Top             =   465
         Width           =   195
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   24
         Left            =   5160
         TabIndex        =   129
         Tag             =   "ForPasModeRedirectAddress"
         Top             =   480
         Width           =   2000
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   25
         Left            =   1560
         TabIndex        =   118
         Tag             =   "ForMaxShareRedirectAddress"
         Top             =   480
         Width           =   2000
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   26
         Left            =   1560
         TabIndex        =   120
         Tag             =   "ForMaxSlotsRedirectAddress"
         Top             =   1200
         Width           =   2000
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   29
         Left            =   1560
         TabIndex        =   121
         Tag             =   "ForMaxHubsRedirectAddress"
         Top             =   1560
         Width           =   2000
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   30
         Left            =   1560
         TabIndex        =   122
         Tag             =   "ForNoTagRedirectAddress"
         Top             =   1920
         Width           =   2000
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   27
         Left            =   1560
         TabIndex        =   119
         Tag             =   "ForMinSlotsRedirectAddress"
         Top             =   840
         Width           =   2000
      End
      Begin VB.TextBox txtData 
         Height          =   250
         Index           =   5
         Left            =   1560
         TabIndex        =   117
         Tag             =   "ForMinShareRedirectAddress"
         Top             =   120
         Width           =   2000
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "For Max Share"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   37
         Left            =   120
         TabIndex        =   131
         Tag             =   "RedirectFMaxShare"
         Top             =   460
         Width           =   195
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "For Min Slots"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   39
         Left            =   120
         TabIndex        =   132
         Tag             =   "RedirectFMinSlots"
         Top             =   820
         Width           =   195
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "For Max Hubs"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   41
         Left            =   120
         TabIndex        =   134
         Tag             =   "RedirectFMaxHubs"
         Top             =   1540
         Width           =   195
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "For No Tag"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   42
         Left            =   120
         TabIndex        =   135
         Tag             =   "RedirectFNoTag"
         Top             =   1900
         Width           =   195
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "For Max Slots"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   38
         Left            =   120
         TabIndex        =   133
         Tag             =   "RedirectFMaxSlots"
         Top             =   1180
         Width           =   195
      End
      Begin VB.CheckBox chkData 
         Appearance      =   0  'Flat
         Caption         =   "For Min Share"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   22
         Left            =   120
         TabIndex        =   130
         Tag             =   "RedirectFMinShare"
         Top             =   100
         Width           =   195
      End
      Begin VB.Label lblOptRedirect 
         BackStyle       =   0  'Transparent
         Caption         =   "Do not redirect"
         Height          =   255
         Index           =   5
         Left            =   4080
         TabIndex        =   226
         ToolTipText     =   "Do not redirect anyone (disconnects if the hub is full)"
         Top             =   2400
         Width           =   4695
      End
      Begin VB.Label lblOptRedirect 
         BackStyle       =   0  'Transparent
         Caption         =   "Only if full and the user is not registered"
         Height          =   255
         Index           =   3
         Left            =   4080
         TabIndex        =   225
         ToolTipText     =   "Only redirect if the user is not registered and the hub is full"
         Top             =   2160
         Width           =   4695
      End
      Begin VB.Label lblOptRedirect 
         BackStyle       =   0  'Transparent
         Caption         =   "Only if full and the user is not an op"
         Height          =   255
         Index           =   4
         Left            =   4080
         TabIndex        =   224
         ToolTipText     =   "Only redirect if the user is an operator (or of a higher class) and the hub is full"
         Top             =   1920
         Width           =   4695
      End
      Begin VB.Label lblOptRedirect 
         BackStyle       =   0  'Transparent
         Caption         =   "Only if full"
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   223
         ToolTipText     =   "Only redirect users if the hub is full"
         Top             =   1680
         Width           =   3615
      End
      Begin VB.Label lblOptRedirect 
         BackStyle       =   0  'Transparent
         Caption         =   "Unregistered users"
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   222
         ToolTipText     =   "Only unregistered users are redirected (must have ""Allow only registered users"" in Security/Advanced checked)"
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Label lblOptRedirect 
         BackStyle       =   0  'Transparent
         Caption         =   "All users"
         Height          =   255
         Index           =   0
         Left            =   4080
         TabIndex        =   221
         ToolTipText     =   "Redirects all users to your redirect address"
         Top             =   1200
         Width           =   3615
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "Allow ops to redirect (admins unaffected)"
         Height          =   375
         Index           =   30
         Left            =   3960
         TabIndex        =   220
         ToolTipText     =   "Check to allow operators to use the redirect ability (admins can always redirect)"
         Top             =   3000
         Width           =   4815
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "Passive Mode"
         Height          =   375
         Index           =   49
         Left            =   3960
         TabIndex        =   219
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "For Fake Tag"
         Height          =   375
         Index           =   46
         Left            =   3960
         TabIndex        =   218
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "For Fake Share"
         Height          =   375
         Index           =   47
         Left            =   360
         TabIndex        =   217
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "For KB / Slot"
         Height          =   375
         Index           =   45
         Left            =   360
         TabIndex        =   216
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "Too Old NMDC"
         Height          =   375
         Index           =   40
         Left            =   360
         TabIndex        =   215
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "Too Old DC++"
         Height          =   375
         Index           =   44
         Left            =   360
         TabIndex        =   214
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "For Slot / Hub"
         Height          =   375
         Index           =   43
         Left            =   360
         TabIndex        =   213
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "For No Tag"
         Height          =   375
         Index           =   42
         Left            =   360
         TabIndex        =   212
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "For Max Hubs"
         Height          =   375
         Index           =   41
         Left            =   360
         TabIndex        =   211
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "For Max Slots"
         Height          =   375
         Index           =   38
         Left            =   360
         TabIndex        =   210
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "For Min Slots"
         Height          =   375
         Index           =   39
         Left            =   360
         TabIndex        =   209
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "For Max Share"
         Height          =   375
         Index           =   37
         Left            =   360
         TabIndex        =   208
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         BackStyle       =   0  'Transparent
         Caption         =   "For Min Share"
         Height          =   375
         Index           =   22
         Left            =   360
         TabIndex        =   207
         Top             =   120
         Width           =   1215
      End
      Begin VB.Shape Shape 
         BorderColor     =   &H00C0C0C0&
         Height          =   1935
         Index           =   20
         Left            =   3720
         Top             =   840
         Width           =   5175
      End
      Begin VB.Label lblHolder 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Redirect Options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   24
         Left            =   3840
         TabIndex        =   206
         Top             =   885
         Width           =   4815
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   4115
      Index           =   3
      Left            =   120
      ScaleHeight     =   4110
      ScaleWidth      =   9015
      TabIndex        =   155
      Top             =   420
      Visible         =   0   'False
      Width           =   9015
      Begin VB.PictureBox picBordTab 
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   3
         Left            =   7650
         ScaleHeight     =   260
         ScaleMode       =   0  'User
         ScaleWidth      =   1395
         TabIndex        =   156
         Top             =   60
         Width           =   1400
      End
      Begin VB.PictureBox picTabAdv 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   0
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8655
         TabIndex        =   188
         Top             =   430
         Width           =   8655
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Allow only registered users"
            BeginProperty DataFormat 
               Type            =   4
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   8
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   85
            Tag             =   "RegOnly"
            ToolTipText     =   "Only registered users may connect to the hub"
            Top             =   3120
            Width           =   195
         End
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   13
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   193
            Text            =   "0"
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   12
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   192
            Text            =   "0"
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   11
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   191
            Text            =   "0"
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   10
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   190
            Text            =   "0"
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Send to any user, including users below min class"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   34
            Left            =   240
            MaskColor       =   &H00000000&
            TabIndex        =   81
            Tag             =   "MinClsConnectSend"
            ToolTipText     =   "Strip all MyINFO before sending to unregistered users"
            Top             =   2160
            Width           =   195
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   12
            LargeChange     =   5
            Left            =   855
            Max             =   0
            Min             =   32000
            SmallChange     =   5
            TabIndex        =   77
            Tag             =   "MaxMessageLen"
            Top             =   840
            Width           =   255
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   11
            Left            =   855
            Max             =   0
            Min             =   11
            TabIndex        =   80
            Tag             =   "MinConnectCls"
            Top             =   1800
            Width           =   255
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   10
            Left            =   855
            Max             =   0
            Min             =   11
            TabIndex        =   78
            Tag             =   "MinSearchCls"
            Top             =   1200
            Width           =   255
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Run in chat only mode"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   32
            Left            =   240
            TabIndex        =   82
            Tag             =   "ChatOnly"
            ToolTipText     =   "Disables searching/connecting for all users; chatting in private messages and the main chat permitted"
            Top             =   2400
            Width           =   195
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   2
            Left            =   855
            Max             =   -1
            Min             =   99
            TabIndex        =   76
            Tag             =   "MinPassiveSearchLen"
            Top             =   480
            Width           =   255
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Send main chat messages to users in away mode"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   83
            Tag             =   "SendMessageAFK"
            ToolTipText     =   "All (NMDC) users who are in away mode will recieve main chat messages"
            Top             =   2640
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Hide MyINFOs to unregistered users"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   36
            Left            =   240
            TabIndex        =   84
            Tag             =   "HideMyinfos"
            Top             =   2880
            Width           =   195
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   1
            LargeChange     =   10
            Left            =   855
            Max             =   1
            Min             =   32767
            TabIndex        =   75
            Tag             =   "MaxUsers"
            Top             =   120
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Send to all users, including users below min class"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   240
            MaskColor       =   &H00000000&
            TabIndex        =   79
            Tag             =   "MinClsSearchSend"
            ToolTipText     =   "Check to allow users above/equal to the min class to search users below the min class"
            Top             =   1560
            Width           =   195
         End
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   9
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   189
            Text            =   "0"
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Allow only registered users"
            Height          =   255
            Index           =   10
            Left            =   480
            TabIndex        =   453
            ToolTipText     =   "Only registered users may connect to the hub"
            Top             =   3120
            Width           =   3855
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Show MyINFOs only to OPs"
            Height          =   255
            Index           =   36
            Left            =   480
            TabIndex        =   203
            Top             =   2880
            Width           =   6855
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Send main chat messages to users in away mode"
            Height          =   255
            Index           =   7
            Left            =   480
            TabIndex        =   202
            ToolTipText     =   "All (NMDC) users who are in away mode will recieve main chat messages"
            Top             =   2640
            Width           =   7095
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Run in chat only mode"
            Height          =   255
            Index           =   32
            Left            =   480
            TabIndex        =   201
            ToolTipText     =   "Disables searching/connecting for all users; chatting in private messages and the main chat permitted"
            Top             =   2400
            Width           =   7095
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Send to any user, including users below min class"
            Height          =   255
            Index           =   34
            Left            =   480
            TabIndex        =   200
            ToolTipText     =   "Strip all MyINFO before sending to unregistered users"
            Top             =   2160
            Width           =   7095
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Send to all users, including users below min class"
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   199
            ToolTipText     =   "Check to allow users above/equal to the min class to search users below the min class"
            Top             =   1560
            Width           =   7095
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "Maximum main chat message length"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   88
            Left            =   1200
            TabIndex        =   198
            Top             =   840
            Width           =   6315
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum class required for downloading"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   85
            Left            =   1200
            TabIndex        =   197
            Top             =   1800
            Width           =   6195
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum class required for searching"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   8
            Left            =   1200
            TabIndex        =   196
            Top             =   1200
            Width           =   6315
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum passive search request length"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   31
            Left            =   1200
            TabIndex        =   195
            Top             =   480
            Width           =   6315
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "Max Users"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   6
            Left            =   1200
            TabIndex        =   194
            Top             =   120
            Width           =   2895
         End
      End
      Begin VB.PictureBox picTabAdv 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   1
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8655
         TabIndex        =   178
         Top             =   430
         Visible         =   0   'False
         Width           =   8655
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Register hub with public hub list"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   87
            Tag             =   "AutoRegister"
            ToolTipText     =   "Register the hub with the selected servers"
            Top             =   480
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Compact user database on exit"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   14
            Left            =   240
            TabIndex        =   88
            Tag             =   "CompactDBOnExit"
            ToolTipText     =   "Will compress the database to a smaller size on exit"
            Top             =   720
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Preload winsocks on start serving"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   86
            Tag             =   "PreloadWinsocks"
            ToolTipText     =   "Will preload sufficent (typically) connections assuming your entire hub was full (faster but uses more memory)"
            Top             =   240
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Update DynDNS / No-IP service"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   52
            Left            =   240
            TabIndex        =   93
            Tag             =   "DynUpdate"
            Top             =   1920
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Check for updates on start up"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   28
            Left            =   240
            TabIndex        =   89
            Tag             =   "AutoCheckUpdate"
            ToolTipText     =   "Will check for a new version of PTDCH on start up"
            Top             =   960
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Start serving on program start"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   90
            Tag             =   "AutoStart"
            ToolTipText     =   "Automatically start serving when PTDCH is opened"
            Top             =   1200
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Start minimized to the system tray"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   26
            Left            =   240
            TabIndex        =   91
            Tag             =   "StartMinimized"
            ToolTipText     =   "When the hub is started, it will remain hidden in the system tray"
            Top             =   1440
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Start PTDCH at windows starting"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   5
            Left            =   240
            MaskColor       =   &H00908675&
            TabIndex        =   92
            Tag             =   "StartWin"
            Top             =   1680
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   1095
            Index           =   3
            Left            =   120
            Top             =   2280
            Visible         =   0   'False
            Width           =   8415
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Update DynDNS / No-IP service"
            Height          =   255
            Index           =   52
            Left            =   480
            TabIndex        =   187
            Top             =   1920
            Width           =   3495
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Start minimized to the system tray"
            Height          =   255
            Index           =   26
            Left            =   480
            TabIndex        =   186
            ToolTipText     =   "When the hub is started, it will remain hidden in the system tray"
            Top             =   1440
            Width           =   7215
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Start serving on program start"
            Height          =   255
            Index           =   9
            Left            =   480
            TabIndex        =   185
            ToolTipText     =   "Automatically start serving when PTDCH is opened"
            Top             =   1200
            Width           =   7215
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Check for updates on start up"
            Height          =   255
            Index           =   28
            Left            =   480
            TabIndex        =   184
            ToolTipText     =   "Will check for a new version of PTDCH on start up"
            Top             =   960
            Width           =   7215
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Register hub with public hub list"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   183
            ToolTipText     =   "Register the hub with the selected servers"
            Top             =   480
            Width           =   7215
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Compact user database on exit"
            Height          =   255
            Index           =   14
            Left            =   480
            TabIndex        =   182
            ToolTipText     =   "Will compress the database to a smaller size on exit"
            Top             =   720
            Width           =   7215
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Start PTDCH at windows starting"
            Height          =   255
            Index           =   5
            Left            =   480
            TabIndex        =   181
            Top             =   1680
            Visible         =   0   'False
            Width           =   6975
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Preload winsocks on start serving"
            Height          =   255
            Index           =   4
            Left            =   480
            TabIndex        =   180
            ToolTipText     =   "Will preload sufficent (typically) connections assuming your entire hub was full (faster but uses more memory)"
            Top             =   240
            Width           =   7215
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00808080&
            Height          =   915
            Index           =   92
            Left            =   240
            TabIndex        =   179
            Top             =   2400
            Visible         =   0   'False
            Width           =   8085
         End
      End
      Begin VB.PictureBox picTabAdv 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   3
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8655
         TabIndex        =   452
         Top             =   420
         Visible         =   0   'False
         Width           =   8655
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "UpDate No-IP DNS at Starting PTDCH"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   70
            Left            =   480
            TabIndex        =   483
            Tag             =   "NoIPUpdateStartUp"
            ToolTipText     =   "No-IP DNS Service"
            Top             =   120
            Width           =   195
         End
         Begin VB.CommandButton cmdButton 
            Caption         =   "Force  Update "
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3255
            Index           =   6
            Left            =   120
            TabIndex        =   482
            Top             =   120
            Width           =   255
         End
         Begin VB.TextBox txtData 
            Height          =   250
            Index           =   38
            Left            =   1680
            TabIndex        =   467
            Tag             =   "NoIPUser"
            ToolTipText     =   "No-IP DNS Service Account Name"
            Top             =   480
            Width           =   2484
         End
         Begin VB.TextBox txtData 
            Height          =   250
            IMEMode         =   3  'DISABLE
            Index           =   42
            Left            =   1680
            PasswordChar    =   "*"
            TabIndex        =   466
            Tag             =   "NoIPPass"
            ToolTipText     =   "No-IP DNS Service Password"
            Top             =   840
            Width           =   2484
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Enable No-IP DNS Update(s)"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   71
            Left            =   8040
            TabIndex        =   465
            Tag             =   "NoIPUpdateEna"
            ToolTipText     =   "No-IP DNS Service"
            Top             =   120
            Width           =   195
         End
         Begin VB.TextBox txtData 
            Height          =   250
            Index           =   41
            Left            =   4920
            TabIndex        =   464
            Tag             =   "NoIPDNS3"
            ToolTipText     =   "No-IP DNS Service"
            Top             =   960
            Width           =   2970
         End
         Begin VB.TextBox txtData 
            Height          =   250
            Index           =   40
            Left            =   4920
            TabIndex        =   463
            Tag             =   "NoIPDNS2"
            ToolTipText     =   "No-IP DNS Service"
            Top             =   600
            Width           =   2970
         End
         Begin VB.TextBox txtData 
            Height          =   250
            Index           =   39
            Left            =   1200
            TabIndex        =   462
            Tag             =   "NoIPDNS1"
            ToolTipText     =   "No-IP DNS Service"
            Top             =   1320
            Width           =   2970
         End
         Begin VB.TextBox txtData 
            Height          =   250
            Index           =   37
            Left            =   4920
            TabIndex        =   461
            Tag             =   "NoIPDNS4"
            ToolTipText     =   "No-IP DNS Service"
            Top             =   1320
            Width           =   2970
         End
         Begin VB.TextBox txtData 
            Height          =   250
            Index           =   43
            Left            =   4920
            TabIndex        =   460
            Tag             =   "DynDNS4"
            ToolTipText     =   "No-IP DNS Service"
            Top             =   3000
            Width           =   2970
         End
         Begin VB.TextBox txtData 
            Height          =   250
            Index           =   44
            Left            =   1200
            TabIndex        =   459
            Tag             =   "DynDNS1"
            ToolTipText     =   "No-IP DNS Service"
            Top             =   3000
            Width           =   2970
         End
         Begin VB.TextBox txtData 
            Height          =   250
            Index           =   45
            Left            =   4920
            TabIndex        =   458
            Tag             =   "DynDNS2"
            ToolTipText     =   "No-IP DNS Service"
            Top             =   2280
            Width           =   2970
         End
         Begin VB.TextBox txtData 
            Height          =   250
            Index           =   46
            Left            =   4920
            TabIndex        =   457
            Tag             =   "DynDNS3"
            ToolTipText     =   "No-IP DNS Service"
            Top             =   2640
            Width           =   2970
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Enable Dyn DNS Update(s)"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   69
            Left            =   8040
            TabIndex        =   456
            Tag             =   "DynDNSUpdateEna"
            ToolTipText     =   "No-IP DNS Service"
            Top             =   1800
            Width           =   195
         End
         Begin VB.TextBox txtData 
            Height          =   250
            IMEMode         =   3  'DISABLE
            Index           =   47
            Left            =   1680
            PasswordChar    =   "*"
            TabIndex        =   455
            Tag             =   "DynDNSPass"
            ToolTipText     =   "No-IP DNS Service Password"
            Top             =   2520
            Width           =   2484
         End
         Begin VB.TextBox txtData 
            Height          =   250
            Index           =   48
            Left            =   1680
            TabIndex        =   454
            Tag             =   "DynDNSUser"
            ToolTipText     =   "No-IP DNS Service Account Name"
            Top             =   2160
            Width           =   2484
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "UpDate No-IP DNS at Starting PTDCH"
            Height          =   255
            Index           =   70
            Left            =   720
            TabIndex        =   484
            ToolTipText     =   "No-IP DNS Service"
            Top             =   120
            Width           =   3375
         End
         Begin VB.Label lblCheck 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Enable No-IP DNS Update(s)"
            Height          =   255
            Index           =   71
            Left            =   4200
            TabIndex        =   481
            Top             =   120
            Width           =   3735
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DSN 1"
            Height          =   255
            Index           =   30
            Left            =   360
            TabIndex        =   480
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DSN 2"
            Height          =   255
            Index           =   34
            Left            =   4200
            TabIndex        =   479
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DSN 3"
            Height          =   255
            Index           =   35
            Left            =   4200
            TabIndex        =   478
            Top             =   960
            Width           =   615
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DSN 4"
            Height          =   255
            Index           =   36
            Left            =   4200
            TabIndex        =   477
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   255
            Index           =   29
            Left            =   720
            TabIndex        =   476
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            Height          =   495
            Index           =   37
            Left            =   480
            TabIndex        =   475
            Top             =   840
            Width           =   1095
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   39
            X1              =   8280
            X2              =   8280
            Y1              =   1680
            Y2              =   240
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   40
            X1              =   1200
            X2              =   8280
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Label lblCheck 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Enable Dyn DNS Update(s)"
            Height          =   255
            Index           =   69
            Left            =   4080
            TabIndex        =   474
            Top             =   1800
            Width           =   3855
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DSN 1"
            Height          =   255
            Index           =   38
            Left            =   360
            TabIndex        =   473
            Top             =   3000
            Width           =   735
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DSN 2"
            Height          =   255
            Index           =   39
            Left            =   4200
            TabIndex        =   472
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DSN 3"
            Height          =   255
            Index           =   41
            Left            =   4200
            TabIndex        =   471
            Top             =   2640
            Width           =   615
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DSN 4"
            Height          =   255
            Index           =   42
            Left            =   4200
            TabIndex        =   470
            Top             =   3000
            Width           =   615
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   255
            Index           =   43
            Left            =   720
            TabIndex        =   469
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            Height          =   495
            Index           =   46
            Left            =   720
            TabIndex        =   468
            Top             =   2520
            Width           =   855
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   48
            X1              =   8280
            X2              =   8280
            Y1              =   3360
            Y2              =   1920
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   49
            X1              =   1200
            X2              =   8280
            Y1              =   3360
            Y2              =   3360
         End
      End
      Begin VB.PictureBox picTabAdv 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   2
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8655
         TabIndex        =   168
         Top             =   430
         Visible         =   0   'False
         Width           =   8655
         Begin VB.CommandButton cmdPopup 
            Caption         =   "None"
            Height          =   375
            Index           =   4
            Left            =   5640
            TabIndex        =   106
            Top             =   2760
            Width           =   1215
         End
         Begin VB.CommandButton cmdPopup 
            Caption         =   "App"
            Height          =   375
            Index           =   3
            Left            =   4320
            TabIndex        =   105
            Top             =   2760
            Width           =   1215
         End
         Begin VB.CommandButton cmdPopup 
            Caption         =   "Error"
            Height          =   375
            Index           =   2
            Left            =   3000
            TabIndex        =   104
            Top             =   2760
            Width           =   1215
         End
         Begin VB.CommandButton cmdPopup 
            Caption         =   "Warning"
            Height          =   375
            Index           =   1
            Left            =   1680
            TabIndex        =   103
            Top             =   2760
            Width           =   1215
         End
         Begin VB.CommandButton cmdPopup 
            Caption         =   "Info"
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   102
            Top             =   2760
            Width           =   1215
         End
         Begin VB.CheckBox chkData 
            Caption         =   "Show popup on new user registed"
            Height          =   195
            Index           =   55
            Left            =   240
            TabIndex        =   94
            Tag             =   "PopUpNewReg"
            Top             =   240
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Caption         =   "Show popup on Op conected"
            Height          =   195
            Index           =   56
            Left            =   240
            TabIndex        =   95
            Tag             =   "PopUpOpConected"
            Top             =   480
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Caption         =   "Show popup on Op disconected"
            Height          =   195
            Index           =   57
            Left            =   240
            TabIndex        =   96
            Tag             =   "PopUpOpDisconected"
            Top             =   720
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Caption         =   "Show popup on user kicked"
            Height          =   195
            Index           =   58
            Left            =   240
            TabIndex        =   97
            Tag             =   "PopUpUserKick"
            Top             =   960
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Caption         =   "Show popup on user baned"
            Height          =   195
            Index           =   59
            Left            =   240
            TabIndex        =   98
            Tag             =   "PopUpUserBaned"
            Top             =   1200
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Caption         =   "Show popup on user redirected"
            Height          =   195
            Index           =   60
            Left            =   240
            TabIndex        =   99
            Tag             =   "PopUpUserRedirected"
            Top             =   1440
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Caption         =   "Show popup on started serving"
            Height          =   195
            Index           =   62
            Left            =   240
            TabIndex        =   100
            Tag             =   "PopUpStartedServing"
            Top             =   1680
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Caption         =   "Show popup on stoped serving"
            Height          =   195
            Index           =   63
            Left            =   240
            TabIndex        =   101
            Tag             =   "PopUpStopedServing"
            Top             =   1920
            Width           =   195
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Poup test"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   14
            Left            =   720
            TabIndex        =   177
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   44
            X1              =   240
            X2              =   6960
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   43
            X1              =   240
            X2              =   240
            Y1              =   3240
            Y2              =   2640
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   42
            X1              =   240
            X2              =   600
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Show popup on stoped serving"
            Height          =   255
            Index           =   63
            Left            =   480
            TabIndex        =   176
            Top             =   1920
            Width           =   5535
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Show popup on started serving"
            Height          =   255
            Index           =   62
            Left            =   480
            TabIndex        =   175
            Top             =   1680
            Width           =   5535
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Show popup on user redirected"
            Height          =   255
            Index           =   60
            Left            =   480
            TabIndex        =   174
            Top             =   1440
            Width           =   5535
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Show popup on user baned"
            Height          =   255
            Index           =   59
            Left            =   480
            TabIndex        =   173
            Top             =   1200
            Width           =   5535
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Show popup on user kicked"
            Height          =   255
            Index           =   58
            Left            =   480
            TabIndex        =   172
            Top             =   960
            Width           =   5535
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Show popup on Op disconected"
            Height          =   255
            Index           =   57
            Left            =   480
            TabIndex        =   171
            Top             =   720
            Width           =   5535
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Show popup on Op conected"
            Height          =   255
            Index           =   56
            Left            =   480
            TabIndex        =   170
            Top             =   480
            Width           =   5535
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Show popup on new user registed"
            Height          =   255
            Index           =   55
            Left            =   480
            TabIndex        =   169
            Top             =   240
            Width           =   5535
         End
      End
      Begin VB.PictureBox picTabAdv 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   4
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8655
         TabIndex        =   157
         Top             =   430
         Visible         =   0   'False
         Width           =   8655
         Begin VB.CheckBox chkData 
            Caption         =   "Enabled Skin"
            Height          =   195
            Index           =   66
            Left            =   4920
            TabIndex        =   111
            Tag             =   "blSkin"
            Top             =   480
            Width           =   195
         End
         Begin VB.ComboBox cmbSkin 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4920
            Style           =   2  'Dropdown List
            TabIndex        =   112
            Top             =   840
            Width           =   2535
         End
         Begin VB.CheckBox chkData 
            Caption         =   "Allow usage"
            Height          =   195
            Index           =   65
            Left            =   480
            TabIndex        =   116
            Tag             =   "PriorityBl"
            Top             =   2040
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Move Form when clicking in any part"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   64
            Left            =   240
            TabIndex        =   110
            Tag             =   "MoveForm"
            Top             =   960
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Enabled Magnetic Windows"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   54
            Left            =   240
            TabIndex        =   109
            Tag             =   "MagneticWin"
            Top             =   720
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Confirm exit"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   107
            Tag             =   "ConfirmExit"
            ToolTipText     =   "Ask if you want to exit the hub before unloading application"
            Top             =   240
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Minimize to system tray"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   35
            Left            =   240
            TabIndex        =   108
            Tag             =   "MinimizeTray"
            ToolTipText     =   "Check to have PTDCH minimize to the system tray"
            Top             =   480
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Caption         =   "Roundom skin at startup PTDCH"
            Enabled         =   0   'False
            Height          =   195
            Index           =   67
            Left            =   4920
            TabIndex        =   115
            Tag             =   "RndSkin"
            Top             =   1680
            Width           =   195
         End
         Begin VB.CommandButton cmdSkin 
            Caption         =   "<"
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   4920
            TabIndex        =   113
            Top             =   1200
            Width           =   375
         End
         Begin VB.CommandButton cmdSkin 
            Caption         =   ">"
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   5280
            TabIndex        =   114
            Top             =   1200
            Width           =   375
         End
         Begin ComctlLib.Slider sldPriority 
            Height          =   315
            Left            =   360
            TabIndex        =   424
            Tag             =   "PriorityVal"
            Top             =   2280
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   556
            _Version        =   327682
            LargeChange     =   1
            Max             =   3
            SelectRange     =   -1  'True
            SelStart        =   1
            Value           =   1
         End
         Begin VB.Label lblPriority 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Real Time"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   3
            Left            =   3480
            TabIndex        =   428
            Top             =   2580
            Width           =   810
         End
         Begin VB.Label lblPriority 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "High"
            Enabled         =   0   'False
            ForeColor       =   &H008080FF&
            Height          =   195
            Index           =   2
            Left            =   2520
            TabIndex        =   427
            Top             =   2580
            Width           =   420
         End
         Begin VB.Label lblPriority 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Normal"
            Enabled         =   0   'False
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   426
            Top             =   2580
            Width           =   495
         End
         Begin VB.Label lblPriority 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Idle"
            Enabled         =   0   'False
            ForeColor       =   &H00C0C000&
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   425
            Top             =   2580
            Width           =   270
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Enabled Skin"
            Height          =   255
            Index           =   66
            Left            =   5160
            TabIndex        =   167
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Skin"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   19
            Left            =   5160
            TabIndex        =   166
            Top             =   120
            Width           =   1455
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   59
            X1              =   4680
            X2              =   8520
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   58
            X1              =   4680
            X2              =   4680
            Y1              =   2040
            Y2              =   240
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   57
            X1              =   4680
            X2              =   5040
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Allow usage"
            Height          =   255
            Index           =   65
            Left            =   720
            TabIndex        =   165
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label lblPriority 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "*Note: High and Real Time not recommended."
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   164
            Top             =   2955
            Width           =   3975
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "System Priority"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   18
            Left            =   720
            TabIndex        =   163
            Top             =   1680
            Width           =   2895
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   56
            X1              =   240
            X2              =   4320
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   55
            X1              =   240
            X2              =   240
            Y1              =   3240
            Y2              =   1800
         End
         Begin VB.Line Line 
            BorderColor     =   &H00C0C0C0&
            Index           =   54
            X1              =   240
            X2              =   600
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Move Form when clicking in any part"
            Height          =   255
            Index           =   64
            Left            =   480
            TabIndex        =   162
            Top             =   960
            Width           =   3855
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Enabled Magnetic Windows"
            Height          =   255
            Index           =   54
            Left            =   480
            TabIndex        =   161
            Top             =   720
            Width           =   3855
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Minimize to system tray"
            Height          =   255
            Index           =   35
            Left            =   480
            TabIndex        =   160
            ToolTipText     =   "Check to have PTDCH minimize to the system tray"
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Confirm exit"
            Height          =   255
            Index           =   6
            Left            =   480
            TabIndex        =   159
            ToolTipText     =   "Ask if you want to exit the hub before unloading application"
            Top             =   240
            Width           =   3855
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Roundom skin at startup PTDCH"
            Height          =   255
            Index           =   67
            Left            =   5160
            TabIndex        =   158
            Top             =   1680
            Width           =   3375
         End
      End
      Begin ComctlLib.TabStrip tabAdv 
         Height          =   3945
         Left            =   60
         TabIndex        =   204
         Top             =   60
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   6959
         TabWidthStyle   =   2
         TabFixedWidth   =   2646
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   5
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Hub/Users"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Application"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Notifications"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "DSN UpDate"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Miscllaneous"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   4115
      Index           =   2
      Left            =   120
      ScaleHeight     =   4110
      ScaleWidth      =   9015
      TabIndex        =   290
      Top             =   420
      Visible         =   0   'False
      Width           =   9015
      Begin VB.PictureBox picBordTab 
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   4
         Left            =   6120
         ScaleHeight     =   260
         ScaleMode       =   0  'User
         ScaleWidth      =   2820
         TabIndex        =   291
         Top             =   60
         Width           =   2820
      End
      Begin VB.PictureBox picITab 
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   3495
         Index           =   0
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8715
         TabIndex        =   302
         Top             =   430
         Width           =   8715
         Begin VB.TextBox txtData 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Index           =   6
            Left            =   2640
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   42
            Tag             =   "JoinMsg"
            ToolTipText     =   "Message that is sent when a user connects"
            Top             =   360
            Width           =   5895
         End
         Begin VB.TextBox txtData 
            Height          =   250
            Index           =   7
            Left            =   480
            TabIndex        =   34
            Tag             =   "BotName"
            ToolTipText     =   "Name of the bot (used for core messages)"
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtData 
            Height          =   285
            Index           =   8
            Left            =   480
            TabIndex        =   36
            Tag             =   "OpChatName"
            ToolTipText     =   "Name of bot where ops can chat privately"
            Top             =   1080
            Width           =   1575
         End
         Begin VB.OptionButton optJM 
            Appearance      =   0  'Flat
            Caption         =   "Do not send"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   40
            ToolTipText     =   "Do not send an on join message"
            Top             =   2760
            Width           =   195
         End
         Begin VB.OptionButton optJM 
            Appearance      =   0  'Flat
            Caption         =   "Send as main chat message"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   39
            ToolTipText     =   "Sends to the user's main chat window (from bot name)"
            Top             =   2520
            Width           =   195
         End
         Begin VB.OptionButton optJM 
            Appearance      =   0  'Flat
            Caption         =   "Send as private message"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   38
            ToolTipText     =   "Sends in a private message window (from bot name)"
            Top             =   2280
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Send messages in private messages (otherwise they are sent in main chat messages)"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   27
            Left            =   120
            TabIndex        =   41
            Tag             =   "SendMsgAsPrivate"
            ToolTipText     =   "Check to send reasons for failing various rules via a pm (from the default bot name)"
            Top             =   3120
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   33
            Left            =   240
            TabIndex        =   37
            Tag             =   "VIPUseOpChat"
            ToolTipText     =   "Check to allow VIPs to use the op chat"
            Top             =   1440
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Op chat name"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   18
            Left            =   240
            TabIndex        =   35
            Tag             =   "UseOpChat"
            ToolTipText     =   "Check to have the op chat listed in the user list (therefore enabled)"
            Top             =   825
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Bot name"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   21
            Left            =   240
            TabIndex        =   33
            Tag             =   "UseBotName"
            ToolTipText     =   "Check to have the bot listed in the user list"
            Top             =   240
            Width           =   195
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Send messages in private messages (otherwise they are sent in main chat messages)"
            Height          =   255
            Index           =   27
            Left            =   360
            TabIndex        =   311
            ToolTipText     =   "Check to send reasons for failing various rules via a pm (from the default bot name)"
            Top             =   3120
            Width           =   8175
         End
         Begin VB.Label lblOptJM 
            BackStyle       =   0  'Transparent
            Caption         =   "Do not send"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   310
            ToolTipText     =   "Do not send an on join message"
            Top             =   2760
            Width           =   2175
         End
         Begin VB.Label lblOptJM 
            BackStyle       =   0  'Transparent
            Caption         =   "Send as main chat message"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   309
            ToolTipText     =   "Sends to the user's main chat window (from bot name)"
            Top             =   2520
            Width           =   2175
         End
         Begin VB.Label lblOptJM 
            BackStyle       =   0  'Transparent
            Caption         =   "Send as private message"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   308
            ToolTipText     =   "Sends in a private message window (from bot name)"
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Op chat name"
            Height          =   255
            Index           =   18
            Left            =   480
            TabIndex        =   307
            ToolTipText     =   "Check to have the op chat listed in the user list (therefore enabled)"
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Bot name"
            Height          =   255
            Index           =   21
            Left            =   480
            TabIndex        =   306
            ToolTipText     =   "Check to have the bot listed in the user list"
            Top             =   240
            Width           =   1815
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   1815
            Index           =   16
            Left            =   120
            Top             =   120
            Width           =   2415
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "MOTD"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   305
            Top             =   2040
            Width           =   2415
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Entre Welcome Message in here.(MOTD)"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   7
            Left            =   2640
            TabIndex        =   304
            Top             =   120
            Width           =   5895
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Permit VIPs to use the op chat"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   33
            Left            =   480
            TabIndex        =   303
            ToolTipText     =   "Check to allow VIPs to use the op chat"
            Top             =   1440
            Width           =   1935
         End
      End
      Begin VB.PictureBox picITab 
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   3495
         Index           =   3
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8715
         TabIndex        =   292
         Top             =   430
         Visible         =   0   'False
         Width           =   8715
         Begin VB.TextBox txtTagRules 
            BackColor       =   &H8000000F&
            ForeColor       =   &H00000000&
            Height          =   1035
            Left            =   3360
            MultiLine       =   -1  'True
            TabIndex        =   293
            Top             =   480
            Width           =   5175
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            Height          =   250
            Index           =   9
            Left            =   6420
            TabIndex        =   70
            Tag             =   "NMDCMinVersion"
            ToolTipText     =   "Minimum client version for NMDC clients"
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            Height          =   250
            Index           =   17
            Left            =   6420
            TabIndex        =   71
            Tag             =   "DCMinVersion"
            ToolTipText     =   "0 = Disabled"
            Top             =   2040
            Width           =   1095
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Deny all clients without a recognized tag"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   3360
            TabIndex        =   72
            Tag             =   "DenyNoTag"
            ToolTipText     =   "Disconnects all users without a <++ (or another supported) tag"
            Top             =   2400
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Prevent (certain) search tools from searching"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   15
            Left            =   3360
            TabIndex        =   73
            Tag             =   "PreventSearchBots"
            ToolTipText     =   "Prevents search tools such as MoGLO and DCSearch from searching (not connecting)"
            Top             =   2640
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Automatically kick MLDonkey clients"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   3360
            TabIndex        =   74
            Tag             =   "AutoKickMLDC"
            ToolTipText     =   "Automatically kick MLDonkey clients (recommended)"
            Top             =   2880
            Width           =   195
         End
         Begin VB.ListBox lstTagsEx 
            Height          =   2595
            Left            =   1800
            TabIndex        =   69
            Top             =   480
            Width           =   1455
         End
         Begin VB.ListBox lstTagsDef 
            BackColor       =   &H8000000F&
            Height          =   2595
            Left            =   120
            TabIndex        =   68
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Automatically kick MLDonkey clients"
            Height          =   255
            Index           =   8
            Left            =   3600
            TabIndex        =   301
            ToolTipText     =   "Automatically kick MLDonkey clients (recommended)"
            Top             =   2880
            Width           =   4935
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Prevent (certain) search tools from searching"
            Height          =   255
            Index           =   15
            Left            =   3600
            TabIndex        =   300
            ToolTipText     =   "Prevents search tools such as MoGLO and DCSearch from searching (not connecting)"
            Top             =   2640
            Width           =   4935
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Deny all clients without a recognized tag"
            Height          =   255
            Index           =   2
            Left            =   3600
            TabIndex        =   299
            ToolTipText     =   "Disconnects all users without a <++ (or another supported) tag"
            Top             =   2400
            Width           =   4935
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum NMDC version"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   11
            Left            =   3360
            TabIndex        =   298
            Top             =   1680
            Width           =   3015
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum DC++ version"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   12
            Left            =   3480
            TabIndex        =   297
            Top             =   2040
            Width           =   2895
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Special handling that is built in"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   10
            Left            =   3360
            TabIndex        =   296
            Top             =   240
            Width           =   5235
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "Accepted client tags"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   32
            Left            =   1800
            TabIndex        =   295
            Top             =   60
            Width           =   1575
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "Default client tags"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   33
            Left            =   120
            TabIndex        =   294
            Top             =   60
            Width           =   1455
         End
      End
      Begin VB.PictureBox picITab 
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   3495
         Index           =   4
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8715
         TabIndex        =   338
         Top             =   430
         Visible         =   0   'False
         Width           =   8715
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   8
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   342
            Top             =   1440
            Width           =   855
         End
         Begin VB.VScrollBar vslData 
            Enabled         =   0   'False
            Height          =   255
            Index           =   13
            Left            =   3855
            Max             =   1
            Min             =   32555
            TabIndex        =   341
            Tag             =   "MinStrZBloc"
            Top             =   1440
            Value           =   1
            Width           =   255
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Log incoming"
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   50
            Left            =   240
            TabIndex        =   340
            Tag             =   "LogIn"
            ToolTipText     =   "Check to send reasons for failing various rules via a pm (from the default bot name)"
            Top             =   2760
            Width           =   1572
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Log outgoing"
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   51
            Left            =   240
            TabIndex        =   339
            Tag             =   "LogOut"
            ToolTipText     =   "Check to send reasons for failing various rules via a pm (from the default bot name)"
            Top             =   3000
            Width           =   1452
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Caption         =   "Length of string *2 (zbloc)"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   90
            Left            =   2520
            TabIndex        =   344
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Reserve for SVN Build please no need to translate"
            Enabled         =   0   'False
            ForeColor       =   &H00808080&
            Height          =   495
            Left            =   5040
            TabIndex        =   343
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "SVN Debug Options"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   345
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.PictureBox picITab 
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   3495
         Index           =   2
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8715
         TabIndex        =   333
         Top             =   430
         Visible         =   0   'False
         Width           =   8715
         Begin VB.ListBox lstPlan 
            Height          =   645
            Left            =   120
            TabIndex        =   67
            Top             =   2640
            Visible         =   0   'False
            Width           =   8415
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Scheduler"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   53
            Left            =   120
            TabIndex        =   66
            Tag             =   "EnabledScheduler"
            ToolTipText     =   "Check to enable sheduler"
            Top             =   2380
            Width           =   195
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            Height          =   250
            Index           =   18
            Left            =   5880
            MaxLength       =   1
            TabIndex        =   64
            Tag             =   "CPrefix"
            ToolTipText     =   "Enter the prefix for which your built-in commands respond to (single character)"
            Top             =   2040
            Width           =   495
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Built-in command prefix (check to filter from main chat)"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   23
            Left            =   120
            TabIndex        =   63
            Tag             =   "FilterCPrefix"
            ToolTipText     =   "Filters messages starting with the prefix character from the main chat"
            Top             =   2160
            Width           =   195
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            Height          =   250
            Index           =   20
            Left            =   8040
            TabIndex        =   65
            Tag             =   "CSeperator"
            ToolTipText     =   "The seperator in which command params are seperated by"
            Top             =   2040
            Width           =   492
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Enabled"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   29
            Left            =   120
            TabIndex        =   62
            Tag             =   "EnabledCommands"
            ToolTipText     =   "Check to enable commands"
            Top             =   1920
            Width           =   195
         End
         Begin ComctlLib.ListView lvwCommands 
            Height          =   1695
            Left            =   120
            TabIndex        =   61
            Top             =   120
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   2990
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Command trigger"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Minimum class"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Enabled"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Scheduler"
            Height          =   255
            Index           =   53
            Left            =   360
            TabIndex        =   337
            ToolTipText     =   "Check to enable sheduler"
            Top             =   2400
            Width           =   2895
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Built-in command prefix (check to filter from main chat)"
            Height          =   255
            Index           =   23
            Left            =   360
            TabIndex        =   336
            ToolTipText     =   "Filters messages starting with the prefix character from the main chat"
            Top             =   2160
            Width           =   5415
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Enabled"
            Height          =   255
            Index           =   29
            Left            =   360
            TabIndex        =   335
            ToolTipText     =   "Check to enable commands"
            Top             =   1920
            Width           =   2775
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Seperator"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   48
            Left            =   6360
            TabIndex        =   334
            Top             =   2040
            Width           =   1695
         End
      End
      Begin VB.PictureBox picITab 
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   3495
         Index           =   1
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8715
         TabIndex        =   312
         Top             =   430
         Visible         =   0   'False
         Width           =   8715
         Begin VB.ComboBox cmbData 
            Height          =   315
            Index           =   1
            ItemData        =   "frmHub.frx":18ECF
            Left            =   2520
            List            =   "frmHub.frx":18EDF
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Tag             =   "MinShareSize"
            Top             =   480
            Width           =   735
         End
         Begin VB.ComboBox cmbData 
            Height          =   315
            Index           =   2
            ItemData        =   "frmHub.frx":18EF9
            Left            =   2520
            List            =   "frmHub.frx":18F09
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Tag             =   "MaxShareSize"
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   12
            Left            =   7920
            TabIndex        =   53
            Tag             =   "DCMaxHubs"
            ToolTipText     =   "0 = Disabled"
            Top             =   840
            Width           =   372
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   11
            Left            =   7920
            TabIndex        =   52
            Tag             =   "DCSlotsPerHub"
            ToolTipText     =   "0 = Disabled / Decimal values accepted (ex: 0.5 slots per hub)"
            Top             =   480
            Width           =   372
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            Height          =   250
            Index           =   14
            Left            =   7440
            TabIndex        =   55
            Tag             =   "DCOSpeed"
            ToolTipText     =   "Grants the extra slot(s) when O: tag (if present) is equal or greater to this value"
            Top             =   1560
            Width           =   375
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   13
            Left            =   4800
            TabIndex        =   54
            Tag             =   "DCOSlots"
            ToolTipText     =   "0 = Disabled"
            Top             =   1200
            Width           =   372
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            Height          =   250
            Index           =   16
            Left            =   5160
            TabIndex        =   56
            Tag             =   "DCBandPerSlot"
            ToolTipText     =   "0 = Disabled / Decimal values accepted (ex: 4.5 kb/s per slot)"
            Top             =   1920
            Width           =   375
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Include hubs where user is an operator"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   12
            Left            =   4200
            TabIndex        =   57
            Tag             =   "DCIncludeOPed"
            ToolTipText     =   "In DC++ > 0.24 (among), include OPed hubs in hub count ?"
            Top             =   2280
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Deny Socks5 Connection"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   25
            Left            =   4200
            TabIndex        =   59
            Tag             =   "Denysocks5"
            ToolTipText     =   "Deny connection with socks5 (Validate tags must be enable)"
            Top             =   2760
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Deny Passive mode Connections"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   48
            Left            =   4200
            TabIndex        =   60
            Tag             =   "DenyPassive"
            ToolTipText     =   "Deny connection from client in Passive Mode (Validate tags must be enable)"
            Top             =   3000
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Validate tags (helps prevent fake tags)"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   4200
            TabIndex        =   58
            Tag             =   "DCValidateTags"
            ToolTipText     =   "Kick client if anomaly in tags ? (such as H: or S: missing, wrong order, etc)"
            Top             =   2520
            Width           =   195
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   250
            Index           =   15
            Left            =   3360
            TabIndex        =   50
            Tag             =   "MinSlots"
            ToolTipText     =   "Minimum value for total slots of the client / 0 = Disabled"
            Top             =   2640
            Width           =   375
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0.00000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   250
            Index           =   10
            Left            =   960
            LinkTimeout     =   0
            TabIndex        =   43
            Tag             =   "IMinShare"
            ToolTipText     =   "0 = Disabled / Minimum share size"
            Top             =   480
            Width           =   1455
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Use mentoring system"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   19
            Left            =   240
            TabIndex        =   47
            Tag             =   "MentoringSystem"
            ToolTipText     =   "See Mentoring.txt for details - superior min share system"
            Top             =   1800
            Width           =   195
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Perform minor anti share faking checks"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   16
            Left            =   240
            TabIndex        =   49
            Tag             =   "CheckFakeShare"
            ToolTipText     =   "Checks for traditional faking patterns (only the inexperienced are usually caught with this)"
            Top             =   2280
            Width           =   195
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0.00000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   250
            Index           =   22
            Left            =   960
            LinkTimeout     =   0
            TabIndex        =   45
            Tag             =   "IMaxShare"
            ToolTipText     =   "0 = Disabled / Maximum share size"
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   250
            Index           =   23
            Left            =   3360
            TabIndex        =   51
            Tag             =   "MaxSlots"
            ToolTipText     =   "Maximum value for total slots of the client / 0 = Disabled"
            Top             =   3000
            Width           =   375
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Ops/VIPs bypass all share and slot rules"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   13
            Left            =   240
            TabIndex        =   48
            Tag             =   "OPBypass"
            ToolTipText     =   "Check to have Ops/VIPs bypass share and slot rules"
            Top             =   2040
            Width           =   195
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Deny Passive mode Connections"
            Height          =   255
            Index           =   48
            Left            =   4440
            TabIndex        =   332
            ToolTipText     =   "Deny connection from client in Passive Mode (Validate tags must be enable)"
            Top             =   3000
            Width           =   3855
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Deny Socks5 Connection"
            Height          =   255
            Index           =   25
            Left            =   4440
            TabIndex        =   331
            ToolTipText     =   "Deny connection with socks5 (Validate tags must be enable)"
            Top             =   2760
            Width           =   3855
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Validate tags (helps prevent fake tags)"
            Height          =   255
            Index           =   11
            Left            =   4440
            TabIndex        =   330
            ToolTipText     =   "Kick client if anomaly in tags ? (such as H: or S: missing, wrong order, etc)"
            Top             =   2520
            Width           =   3855
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Include hubs where user is an operator"
            Height          =   255
            Index           =   12
            Left            =   4440
            TabIndex        =   329
            ToolTipText     =   "In DC++ > 0.24 (among), include OPed hubs in hub count ?"
            Top             =   2280
            Width           =   3855
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Perform minor anti share faking checks"
            Height          =   255
            Index           =   16
            Left            =   480
            TabIndex        =   328
            ToolTipText     =   "Checks for traditional faking patterns (only the inexperienced are usually caught with this)"
            Top             =   2280
            Width           =   3375
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Ops/VIPs bypass all share and slot rules"
            Height          =   255
            Index           =   13
            Left            =   480
            TabIndex        =   327
            ToolTipText     =   "Check to have Ops/VIPs bypass share and slot rules"
            Top             =   2040
            Width           =   3375
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Use mentoring system"
            Height          =   255
            Index           =   19
            Left            =   480
            TabIndex        =   326
            ToolTipText     =   "See Mentoring.txt for details - superior min share system"
            Top             =   1800
            Width           =   3375
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   1695
            Index           =   19
            Left            =   120
            Top             =   1680
            Width           =   3855
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   1455
            Index           =   18
            Left            =   120
            Top             =   120
            Width           =   3855
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Tag options"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   27
            Left            =   4200
            TabIndex        =   325
            Top             =   195
            Width           =   4215
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Maximum Hubs"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   19
            Left            =   5400
            TabIndex        =   324
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Slot/Hub"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   17
            Left            =   5400
            TabIndex        =   323
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "when upload speed <"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   22
            Left            =   4320
            TabIndex        =   322
            Top             =   1560
            Width           =   3015
         End
         Begin VB.Label lblHolder 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "KB/s"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   23
            Left            =   7920
            TabIndex        =   321
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label lblHolder 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "extra slot(s) if automated slot opens"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   21
            Left            =   5280
            TabIndex        =   320
            Top             =   1200
            Width           =   3015
         End
         Begin VB.Label lblHolder 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Grant"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   20
            Left            =   4320
            TabIndex        =   319
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Require"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   16
            Left            =   4200
            TabIndex        =   318
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label lblHolder 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "KB/s per slot (limiting upload)"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   44
            Left            =   5640
            TabIndex        =   317
            Top             =   1920
            Width           =   2715
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum Slots"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   28
            Left            =   720
            TabIndex        =   316
            Top             =   2640
            Width           =   2535
         End
         Begin VB.Label lblHolder 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum share size"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   25
            Left            =   960
            TabIndex        =   315
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label lblHolder 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Maximum share size"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   40
            Left            =   960
            TabIndex        =   314
            Top             =   840
            Width           =   2295
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Maximum Slots"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   89
            Left            =   720
            TabIndex        =   313
            Top             =   3000
            Width           =   2535
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   3255
            Index           =   17
            Left            =   4080
            Top             =   120
            Width           =   4455
         End
      End
      Begin ComctlLib.TabStrip tbsInteractions 
         Height          =   3945
         Left            =   60
         TabIndex        =   32
         Top             =   60
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   6959
         TabWidthStyle   =   2
         TabFixedWidth   =   2646
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   4
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "General"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "User Controls"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Commands"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Clients Controls"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   4115
      Index           =   1
      Left            =   120
      ScaleHeight     =   4110
      ScaleWidth      =   9015
      TabIndex        =   246
      Top             =   420
      Visible         =   0   'False
      Width           =   9015
      Begin VB.PictureBox picBordTab 
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   2
         Left            =   7640
         ScaleHeight     =   260
         ScaleMode       =   0  'User
         ScaleWidth      =   1380
         TabIndex        =   247
         Top             =   60
         Width           =   1380
      End
      Begin VB.PictureBox picSTab 
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   3495
         Index           =   0
         Left            =   120
         ScaleHeight     =   3495
         ScaleMode       =   0  'User
         ScaleWidth      =   8728.685
         TabIndex        =   289
         Top             =   430
         Width           =   8655
         Begin VB.TextBox txtDBRegCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   447
            Text            =   "0"
            Top             =   140
            Width           =   495
         End
         Begin VB.CommandButton cmdDB 
            Caption         =   "Rename"
            Height          =   300
            Index           =   8
            Left            =   3480
            TabIndex        =   446
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdDB 
            Caption         =   "Edit"
            Height          =   300
            Index           =   6
            Left            =   2640
            TabIndex        =   444
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdDB 
            Caption         =   "Rem"
            Height          =   300
            Index           =   4
            Left            =   1800
            TabIndex        =   442
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdDB 
            Caption         =   "Add"
            Height          =   300
            Index           =   2
            Left            =   960
            TabIndex        =   440
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdDB 
            Caption         =   "Refresh"
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   438
            Top             =   120
            Width           =   855
         End
         Begin MSAdodcLib.Adodc adoUsers 
            Height          =   285
            Left            =   6360
            Top             =   120
            Visible         =   0   'False
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            ConnectMode     =   1
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   1
            CommandType     =   1
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DBs\userdb.mdb;Mode=ReadWrite;Persist Security Info=False"
            OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DBs\userdb.mdb;Mode=ReadWrite;Persist Security Info=False"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   $"frmHub.frx":18F23
            Caption         =   "adoUsers"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.ComboBox cmbRegistered 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmHub.frx":190D9
            Left            =   6360
            List            =   "frmHub.frx":190DB
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   120
            Width           =   2175
         End
         Begin ComctlLib.ListView lvwRegistered 
            Height          =   2895
            Left            =   120
            TabIndex        =   436
            Top             =   480
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   5106
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   8
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "User Name"
               Object.Width           =   2624
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Password"
               Object.Width           =   2624
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Class"
               Object.Width           =   1050
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Class Name"
               Object.Width           =   2519
            EndProperty
            BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   2
               SubItemIndex    =   4
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Reged By"
               Object.Width           =   2519
            EndProperty
            BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               SubItemIndex    =   5
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Reg Date"
               Object.Width           =   2519
            EndProperty
            BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               SubItemIndex    =   6
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Last Login"
               Object.Width           =   2519
            EndProperty
            BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   2
               SubItemIndex    =   7
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Last IP"
               Object.Width           =   2519
            EndProperty
         End
      End
      Begin VB.PictureBox picSTab 
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   3495
         Index           =   2
         Left            =   120
         ScaleHeight     =   3495
         ScaleMode       =   0  'User
         ScaleWidth      =   8668.174
         TabIndex        =   279
         Top             =   430
         Visible         =   0   'False
         Width           =   8595
         Begin VB.TextBox txtBanFilter 
            Height          =   250
            Left            =   6720
            TabIndex        =   17
            Top             =   2880
            Width           =   1575
         End
         Begin VB.OptionButton optBanFilter 
            Appearance      =   0  'Flat
            Caption         =   "Do not filter"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   6840
            TabIndex        =   16
            Top             =   2520
            Value           =   -1  'True
            Width           =   193
         End
         Begin VB.OptionButton optBanFilter 
            Appearance      =   0  'Flat
            Caption         =   "Begin with"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   6840
            TabIndex        =   13
            Top             =   1800
            Width           =   193
         End
         Begin VB.OptionButton optBanFilter 
            Appearance      =   0  'Flat
            Caption         =   "End in"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   6840
            TabIndex        =   15
            Top             =   2280
            Width           =   193
         End
         Begin VB.OptionButton optBanFilter 
            Appearance      =   0  'Flat
            Caption         =   "Contain"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   6840
            TabIndex        =   14
            Top             =   2040
            Width           =   193
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   0
            LargeChange     =   5
            Left            =   7680
            Max             =   -1
            Min             =   32767
            TabIndex        =   12
            Tag             =   "DefaultBanTime"
            Top             =   800
            Width           =   253
         End
         Begin ComctlLib.ListView lvwPermIPBan 
            Height          =   2775
            Left            =   240
            TabIndex        =   10
            Top             =   480
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   4895
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "IP"
               Object.Width           =   2519
            EndProperty
         End
         Begin ComctlLib.ListView lvwTempIPBan 
            Height          =   2775
            Left            =   2640
            TabIndex        =   11
            Top             =   480
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   4895
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "IP"
               Object.Width           =   2519
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Expire"
               Object.Width           =   2519
            EndProperty
         End
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   7181
            Locked          =   -1  'True
            TabIndex        =   280
            Top             =   800
            Width           =   495
         End
         Begin VB.Label lblOptBanFilter 
            BackStyle       =   0  'Transparent
            Caption         =   "Do not filter"
            Height          =   255
            Index           =   0
            Left            =   7080
            TabIndex        =   288
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label lblOptBanFilter 
            BackStyle       =   0  'Transparent
            Caption         =   "End in"
            Height          =   255
            Index           =   1
            Left            =   7080
            TabIndex        =   287
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label lblOptBanFilter 
            BackStyle       =   0  'Transparent
            Caption         =   "Contain"
            Height          =   255
            Index           =   2
            Left            =   7080
            TabIndex        =   286
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblOptBanFilter 
            BackStyle       =   0  'Transparent
            Caption         =   "Begin with"
            Height          =   255
            Index           =   3
            Left            =   7080
            TabIndex        =   285
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   2055
            Index           =   8
            Left            =   6480
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "IP Ban Filter"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   83
            Left            =   6600
            TabIndex        =   284
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   1095
            Index           =   7
            Left            =   6480
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Default kick temp ban length (minutes)"
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   15
            Left            =   6600
            TabIndex        =   283
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "Temporary IP Bans"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   14
            Left            =   2640
            TabIndex        =   282
            Top             =   240
            Width           =   3255
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   3255
            Index           =   6
            Left            =   2520
            Top             =   120
            Width           =   3855
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "Permanent IP Bans"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   281
            Top             =   240
            Width           =   2175
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   3255
            Index           =   5
            Left            =   120
            Top             =   120
            Width           =   2295
         End
      End
      Begin VB.PictureBox picSTab 
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   3495
         Index           =   3
         Left            =   120
         ScaleHeight     =   3495
         ScaleMode       =   0  'User
         ScaleWidth      =   8728.685
         TabIndex        =   251
         Top             =   430
         Visible         =   0   'False
         Width           =   8655
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   7
            Left            =   7440
            Locked          =   -1  'True
            TabIndex        =   258
            Top             =   2880
            Width           =   615
         End
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   257
            Top             =   2880
            Width           =   615
         End
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   7440
            Locked          =   -1  'True
            TabIndex        =   256
            Top             =   2520
            Width           =   615
         End
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   255
            Top             =   2520
            Width           =   615
         End
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   6600
            Locked          =   -1  'True
            TabIndex        =   254
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   6960
            Locked          =   -1  'True
            TabIndex        =   253
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtVSl 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   252
            Top             =   1060
            Width           =   495
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Enable flood wall"
            BeginProperty DataFormat 
               Type            =   4
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   8
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   20
            Left            =   7920
            TabIndex        =   24
            Tag             =   "EnableFloodWall"
            ToolTipText     =   "Aids in preventing flooding via traditional means (ie MyINFO, nicklist, search, etc)"
            Top             =   600
            Width           =   193
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   3
            LargeChange     =   100
            Left            =   7575
            Max             =   1000
            Min             =   32000
            SmallChange     =   100
            TabIndex        =   25
            Tag             =   "FWInterval"
            Top             =   960
            Value           =   1000
            Width           =   255
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   4
            LargeChange     =   10
            Left            =   7215
            Max             =   -1
            Min             =   32767
            TabIndex        =   26
            Tag             =   "FWBanLength"
            Top             =   1320
            Value           =   1
            Width           =   255
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   5
            LargeChange     =   100
            Left            =   8055
            Max             =   1
            Min             =   254
            TabIndex        =   29
            Tag             =   "FWMyINFO"
            Top             =   2520
            Value           =   254
            Width           =   255
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   6
            LargeChange     =   100
            Left            =   8055
            Max             =   1
            Min             =   254
            TabIndex        =   31
            Tag             =   "FWGetNickList"
            Top             =   2880
            Value           =   254
            Width           =   255
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   7
            LargeChange     =   100
            Left            =   6255
            Max             =   1
            Min             =   254
            TabIndex        =   28
            Tag             =   "FWActiveSearch"
            Top             =   2520
            Value           =   254
            Width           =   255
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   8
            LargeChange     =   100
            Left            =   6255
            Max             =   1
            Min             =   254
            TabIndex        =   30
            Tag             =   "FWPassiveSearch"
            Top             =   2880
            Value           =   254
            Width           =   255
         End
         Begin VB.TextBox txtData 
            Alignment       =   2  'Center
            Height          =   250
            Index           =   19
            Left            =   7320
            MaxLength       =   10
            TabIndex        =   27
            Tag             =   "DataFragmentLen"
            Text            =   "2048"
            ToolTipText     =   "Limit messages and protocol commands length. Be careful not to set too low. Default is 2048"
            Top             =   1800
            Width           =   855
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Redirect users who give wrong password"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   31
            Left            =   240
            TabIndex        =   22
            Tag             =   "RedirectFGP"
            ToolTipText     =   "Redirects users who do not know the password to the redirect hub set in General settings"
            Top             =   2280
            Width           =   193
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Run in password mode"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   21
            Tag             =   "PasswordMode"
            ToolTipText     =   "Require all unregistered users to send the password specified below before logging in"
            Top             =   2040
            Width           =   193
         End
         Begin VB.TextBox txtData 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   250
            Index           =   21
            Left            =   2160
            TabIndex        =   23
            Tag             =   "HubPassword"
            ToolTipText     =   "Global password used in password mode"
            Top             =   2760
            Width           =   1335
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Use descriptive ban messages"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   17
            Left            =   240
            TabIndex        =   18
            Tag             =   "DescriptiveBanMsg"
            ToolTipText     =   "ex. ""Your IP is permanently banned."" versus ""Your IP is banned!"""
            Top             =   240
            Width           =   194
         End
         Begin VB.CheckBox chkData 
            Appearance      =   0  'Flat
            Caption         =   "Prevent brute force password guessing"
            BeginProperty DataFormat 
               Type            =   4
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   8
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   24
            Left            =   240
            TabIndex        =   19
            Tag             =   "PreventGuessPass"
            ToolTipText     =   "Limits the number of times you may attempt to log in before being temporarily banned"
            Top             =   600
            Width           =   194
         End
         Begin VB.VScrollBar vslData 
            Height          =   255
            Index           =   9
            Left            =   1800
            Max             =   1
            Min             =   10
            TabIndex        =   20
            Tag             =   "MaxPassAttempts"
            Top             =   1080
            Value           =   10
            Width           =   255
         End
         Begin VB.Label lblCheck 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Enable flood wall"
            Height          =   255
            Index           =   20
            Left            =   4440
            TabIndex        =   278
            ToolTipText     =   "Aids in preventing flooding via traditional means (ie MyINFO, nicklist, search, etc)"
            Top             =   600
            Width           =   3375
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Redirect users who give wrong password"
            Height          =   495
            Index           =   31
            Left            =   480
            TabIndex        =   277
            ToolTipText     =   "Redirects users who do not know the password to the redirect hub set in General settings"
            Top             =   2280
            Width           =   3375
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Run in password mode"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   276
            ToolTipText     =   "Require all unregistered users to send the password specified below before logging in"
            Top             =   2040
            Width           =   3375
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Prevent brute force password guessing"
            Height          =   375
            Index           =   24
            Left            =   480
            TabIndex        =   275
            ToolTipText     =   "Limits the number of times you may attempt to log in before being temporarily banned"
            Top             =   600
            Width           =   3375
         End
         Begin VB.Label lblCheck 
            BackStyle       =   0  'Transparent
            Caption         =   "Use descriptive ban messages"
            Height          =   375
            Index           =   17
            Left            =   480
            TabIndex        =   274
            ToolTipText     =   "ex. ""Your IP is permanently banned."" versus ""Your IP is banned!"""
            Top             =   240
            Width           =   3375
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   3255
            Index           =   11
            Left            =   4080
            Top             =   120
            Width           =   4455
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Flood wall options"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   61
            Left            =   4320
            TabIndex        =   273
            Top             =   240
            Width           =   4035
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Flooding interval checks last"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   62
            Left            =   4200
            TabIndex        =   272
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "ms"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   63
            Left            =   7920
            TabIndex        =   271
            Top             =   960
            Width           =   255
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "MyINFO"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   64
            Left            =   6600
            TabIndex        =   270
            Top             =   2520
            Width           =   735
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nicklist"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   65
            Left            =   6480
            TabIndex        =   269
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Active search"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   66
            Left            =   4080
            TabIndex        =   268
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Passive search"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   67
            Left            =   4080
            TabIndex        =   267
            Top             =   2880
            Width           =   1455
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ban user if flooding for"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   68
            Left            =   4200
            TabIndex        =   266
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "minutes"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   69
            Left            =   7560
            TabIndex        =   265
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Number of permitted sendings during interval "
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   76
            Left            =   4200
            TabIndex        =   264
            Top             =   2160
            Width           =   4215
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Max messages and protocol length"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   9
            Left            =   4200
            TabIndex        =   263
            ToolTipText     =   "Limit messages and protocol commands length. Be careful not to set too low. Default is 2048"
            Top             =   1800
            Width           =   3015
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   1335
            Index           =   10
            Left            =   120
            Top             =   120
            Width           =   3855
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Global password mode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   78
            Left            =   480
            TabIndex        =   262
            Top             =   1680
            Width           =   2895
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   77
            Left            =   480
            TabIndex        =   261
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   1815
            Index           =   9
            Left            =   120
            Top             =   1560
            Width           =   3855
         End
         Begin VB.Label lblHolder 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Permit"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   80
            Left            =   240
            TabIndex        =   260
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblHolder 
            BackStyle       =   0  'Transparent
            Caption         =   "attempts"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   81
            Left            =   2160
            TabIndex        =   259
            Top             =   1080
            Width           =   1455
         End
      End
      Begin VB.PictureBox picSTab 
         BorderStyle     =   0  'None
         Height          =   3495
         Index           =   4
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   8655
         TabIndex        =   501
         Top             =   430
         Visible         =   0   'False
         Width           =   8655
         Begin VB.PictureBox picBordTab 
            BorderStyle     =   0  'None
            Height          =   300
            Index           =   7
            Left            =   3000
            ScaleHeight     =   260
            ScaleMode       =   0  'User
            ScaleWidth      =   5580
            TabIndex        =   508
            Top             =   60
            Width           =   5580
         End
         Begin VB.CommandButton cmdSql 
            Caption         =   "Cls"
            Height          =   285
            Index           =   1
            Left            =   1200
            TabIndex        =   507
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtSqlErr 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   504
            Top             =   480
            Width           =   6735
         End
         Begin VB.CommandButton cmdSql 
            Caption         =   "Run"
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   503
            Top             =   480
            Width           =   975
         End
         Begin VB.PictureBox picSqlSCI 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2415
            Left            =   120
            ScaleHeight     =   2415
            ScaleWidth      =   8295
            TabIndex        =   502
            Top             =   840
            Width           =   8295
         End
         Begin MSDataGridLib.DataGrid dtgSql 
            Height          =   2895
            Left            =   120
            TabIndex        =   505
            Top             =   480
            Visible         =   0   'False
            Width           =   8355
            _ExtentX        =   14737
            _ExtentY        =   5106
            _Version        =   393216
            BorderStyle     =   0
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin ComctlLib.TabStrip tbsDbManager 
            Height          =   3375
            Left            =   60
            TabIndex        =   506
            Top             =   60
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   5953
            TabWidthStyle   =   2
            _Version        =   327682
            BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
               NumTabs         =   2
               BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "Query String"
                  Key             =   ""
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
                  Caption         =   "Data Contents"
                  Key             =   ""
                  Object.Tag             =   ""
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
      End
      Begin VB.PictureBox picSTab 
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   3495
         Index           =   1
         Left            =   120
         ScaleHeight     =   3495
         ScaleMode       =   0  'User
         ScaleWidth      =   8728.685
         TabIndex        =   248
         Top             =   430
         Visible         =   0   'False
         Width           =   8655
         Begin VB.CommandButton cmdDB 
            Caption         =   "Edit"
            Height          =   300
            Index           =   7
            Left            =   2640
            TabIndex        =   445
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdDB 
            Caption         =   "Rem"
            Height          =   300
            Index           =   5
            Left            =   1800
            TabIndex        =   443
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdDB 
            Caption         =   "Add"
            Height          =   300
            Index           =   3
            Left            =   960
            TabIndex        =   441
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdDB 
            Caption         =   "Refresh"
            Height          =   300
            Index           =   1
            Left            =   120
            TabIndex        =   439
            Top             =   120
            Width           =   855
         End
         Begin MSAdodcLib.Adodc adoBans 
            Height          =   270
            Left            =   7200
            Top             =   3000
            Visible         =   0   'False
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   476
            ConnectMode     =   1
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   1
            CommandType     =   1
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DBs\userdb.mdb;Persist Security Info=False"
            OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\DBs\userdb.mdb;Persist Security Info=False"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   $"frmHub.frx":190DD
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin ComctlLib.ListView lvwBans 
            Height          =   1815
            Left            =   120
            TabIndex        =   437
            Top             =   480
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   3201
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "User Name"
               Object.Width           =   3498
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   2
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Perm"
               Object.Width           =   1050
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   2
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Banned by"
               Object.Width           =   3498
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   2
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Reference Date"
               Object.Width           =   3848
            EndProperty
         End
         Begin VB.Shape Shape 
            BorderColor     =   &H00C0C0C0&
            Height          =   975
            Index           =   12
            Left            =   120
            Top             =   2400
            Width           =   8415
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Reason :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   49
            Left            =   960
            TabIndex        =   250
            Top             =   2520
            Width           =   6615
         End
         Begin VB.Label lblHolder 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            DataField       =   "Reason"
            DataSource      =   "adoBans"
            Height          =   495
            Index           =   50
            Left            =   240
            TabIndex        =   249
            Top             =   2760
            Width           =   8175
         End
      End
      Begin ComctlLib.TabStrip tbsSecurity 
         Height          =   3945
         Left            =   60
         TabIndex        =   8
         Top             =   60
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   6959
         TabWidthStyle   =   2
         TabFixedWidth   =   2646
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   5
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Registed Users"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Name Bans"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "IP Bans"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Advanced"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "BD Manager"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin ComctlLib.TabStrip tbsMenu 
      Height          =   4575
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   8070
      TabWidthStyle   =   2
      TabFixedWidth   =   1785
      TabFixedHeight  =   489
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   9
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "General"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Security"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Interactions"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Advanced"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Redirections"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Scripts"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Status"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Misc"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Full Help"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Tray Menu"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuTray 
         Caption         =   "Show"
         Index           =   0
      End
      Begin VB.Menu mnuTray 
         Caption         =   "Hide"
         Index           =   1
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Registered user list"
      Index           =   1
      Visible         =   0   'False
      Begin VB.Menu mnuRegistered 
         Caption         =   "Copy User Name"
         Index           =   0
      End
      Begin VB.Menu mnuRegistered 
         Caption         =   "Copy Password"
         Index           =   1
      End
      Begin VB.Menu mnuRegistered 
         Caption         =   "Copy Last IP"
         Index           =   2
      End
      Begin VB.Menu mnuRegistered 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuRegistered 
         Caption         =   "Copy All"
         Index           =   4
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Temp IP ban list"
      Index           =   2
      Visible         =   0   'False
      Begin VB.Menu mnuTempIPBan 
         Caption         =   "Add"
         Index           =   0
      End
      Begin VB.Menu mnuTempIPBan 
         Caption         =   "Remove"
         Index           =   1
      End
      Begin VB.Menu mnuTempIPBan 
         Caption         =   "Clear"
         Index           =   2
      End
      Begin VB.Menu mnuTempIPBan 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuTempIPBan 
         Caption         =   "Refresh list extract"
         Index           =   4
      End
      Begin VB.Menu mnuTempIPBan 
         Caption         =   "Clear list extract"
         Index           =   5
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Perm IP ban list"
      Index           =   3
      Visible         =   0   'False
      Begin VB.Menu mnuPermIPBan 
         Caption         =   "Add"
         Index           =   0
      End
      Begin VB.Menu mnuPermIPBan 
         Caption         =   "Remove"
         Index           =   1
      End
      Begin VB.Menu mnuPermIPBan 
         Caption         =   "Clear"
         Index           =   2
      End
      Begin VB.Menu mnuPermIPBan 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuPermIPBan 
         Caption         =   "Refresh list extract"
         Index           =   4
      End
      Begin VB.Menu mnuPermIPBan 
         Caption         =   "Clear list extract"
         Index           =   5
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Locked names"
      Index           =   4
      Visible         =   0   'False
      Begin VB.Menu mnuLocked 
         Caption         =   "Copy User Name"
         Index           =   0
      End
      Begin VB.Menu mnuLocked 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuLocked 
         Caption         =   "Copy All"
         Index           =   2
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Tags"
      Index           =   5
      Visible         =   0   'False
      Begin VB.Menu mnuTags 
         Caption         =   "Add"
         Index           =   0
      End
      Begin VB.Menu mnuTags 
         Caption         =   "Remove"
         Index           =   1
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Status"
      Index           =   6
      Visible         =   0   'False
      Begin VB.Menu mnuStatus 
         Caption         =   "Copy"
         Index           =   0
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "Clear"
         Index           =   1
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Users"
      Index           =   7
      Visible         =   0   'False
      Begin VB.Menu mnuUsers 
         Caption         =   "Send data (selected)"
         Index           =   0
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "Send data (all)"
         Index           =   1
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "Disconnect"
         Index           =   2
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "Kick"
         Index           =   3
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "Redirect"
         Index           =   4
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "Ban"
         Index           =   5
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "(De)mute"
         Index           =   6
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "Properties (selected)"
         Index           =   7
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Plan"
      Index           =   8
      Visible         =   0   'False
      Begin VB.Menu mnuPlan 
         Caption         =   "Add"
         Index           =   0
      End
      Begin VB.Menu mnuPlan 
         Caption         =   "Remove"
         Index           =   1
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Scripts"
      Index           =   9
      Visible         =   0   'False
      Begin VB.Menu mnuScripts 
         Caption         =   "Reset / Save"
         Index           =   0
      End
      Begin VB.Menu mnuScripts 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuScripts 
         Caption         =   "Stop"
         Index           =   2
      End
      Begin VB.Menu mnuScripts 
         Caption         =   "Stop All"
         Index           =   3
      End
      Begin VB.Menu mnuScripts 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuScripts 
         Caption         =   "Reolad (Checkeds)"
         Index           =   5
      End
      Begin VB.Menu mnuScripts 
         Caption         =   "Reolad Dir"
         Index           =   6
      End
      Begin VB.Menu mnuScripts 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuScripts 
         Caption         =   "Properties"
         Index           =   8
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "SCI"
      Index           =   10
      Visible         =   0   'False
      Begin VB.Menu mnuCodeRTB 
         Caption         =   "View WhiteSpace"
         Index           =   0
      End
      Begin VB.Menu mnuCodeRTB 
         Caption         =   "Line Number"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuCodeRTB 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuCodeRTB 
         Caption         =   "Misc"
         Index           =   3
         Begin VB.Menu mnuCodeRTB1 
            Caption         =   "Insert Date/Time"
            Index           =   0
         End
         Begin VB.Menu mnuCodeRTB1 
            Caption         =   "Script Info.."
            Index           =   1
         End
         Begin VB.Menu mnuCodeRTB1 
            Caption         =   "Scripts Help"
            Index           =   2
            Begin VB.Menu mnuCodeRTB3 
               Caption         =   "VBScript Documentation"
               Index           =   0
            End
            Begin VB.Menu mnuCodeRTB3 
               Caption         =   "JScript Documentation"
               Index           =   1
            End
         End
         Begin VB.Menu mnuCodeRTB1 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuCodeRTB1 
            Caption         =   "Save as Script.."
            Index           =   4
         End
         Begin VB.Menu mnuCodeRTB1 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnuCodeRTB1 
            Caption         =   "Clear Undo Buffer"
            Index           =   6
         End
      End
      Begin VB.Menu mnuCodeRTB 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuCodeRTB 
         Caption         =   "Help"
         Index           =   5
      End
      Begin VB.Menu mnuCodeRTB 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuCodeRTB 
         Caption         =   "Word Wrap"
         Index           =   7
      End
      Begin VB.Menu mnuCodeRTB 
         Caption         =   "ReadOnly"
         Index           =   8
      End
      Begin VB.Menu mnuCodeRTB 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuCodeRTB 
         Caption         =   "Plug-ins"
         Index           =   10
         Begin VB.Menu mnuPlugIn 
            Caption         =   "No Plug-Ins Found"
            Enabled         =   0   'False
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "frmHub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Compiler conditions

'PredataArrival - Setting this value to true, turns on the optional "predataarrival"
'                 option; this adds a new sub to the mix called PreDataArrival
#Const PreDataArrival = True

'DataArrival - Setting this value to true, turns on the default event DataArrival.
'              It is a CPU intensive event, and if you are not using it in your scripts
'              I suggest you set this value to false
#Const DataArrival = True

'ObjectNotSet - Makes a check in wskLoop_DataArrival to make sure user object exists
#Const OBJECTNOTSET = True

'ColFreeSocks - Uses a collection to find free winsocks (otherwise it loops)
#Const COLFREESOCKS = True

'Status window - Setting this value to true turns on the Status / Admin panel.
'                Must be set in the Properties dialog (just included here for clarity)
#Const Status = True

'API calls
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long

'Send E-Mail
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long
Private Const SW_SHOWMAXIMIZED = 3

'API calls MyIPTools.dll /////////////////////////////////////////////////////////////
Private Declare Function DetectIP Lib "MyIPTools.DLL" () As Variant
Private Declare Function ResolveHost Lib "MyIPTools.DLL" (sIP As Variant) As Variant
Private Declare Function UpdateDynDNS Lib "MyIPTools.DLL" (User As Variant, Pass As Variant, Host As Variant, auto As Boolean, sIP As Variant, mail As Boolean) As Variant
Private Declare Function UpdateNoIP Lib "MyIPTools.DLL" (User As Variant, Pass As Variant, Host As Variant, sIP As Variant) As Variant
Private Declare Function IPinRange Lib "MyIPTools.DLL" (sIP As Variant, eIP As Variant, IP As Variant) As Boolean
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'Types
Private Type typBot
    Name As String
    MyINFO As String
    Operator As Boolean
End Type

'Private objects
Private WithEvents m_objDetectIP    As clsHTTPDownload
Attribute m_objDetectIP.VB_VarHelpID = -1
Private m_objPermaCon               As Connection
Private m_objLoopUser               As clsUser
Private m_wskLoopItem               As Winsock
Private m_colTags                   As Collection
Private m_colFailedReg              As Collection
Private m_colConnectAttempts        As Collection
Private m_colRevConnects            As Collection

#If COLFREESOCKS Then
    Private m_colFreeSocks          As Collection
#End If

'#If PREDATAARRIVAL Then
'    Private m_intPDIndex        As Integer
'#End If

'Cool FX Magnetic Windows
Private Magnetic                    As New clsMagneticWnd

Private m_sciSql                    As clsYScintilla
Attribute m_sciSql.VB_VarHelpID = -1

'Private vars
Private m_arrScriptEvents()         As Boolean
Private m_blnServing                As Boolean
Private m_blnCommaDecimal           As Boolean
Private m_lngScriptEventsUB         As Long
Private m_lngRedirectUB             As Long
Private m_lngBotsUB                 As Long
Private m_lngBanFilter              As Long
Private m_arrRedirectIPs()          As String
Private m_arrBots()                 As typBot
Private m_datServingDate            As Date
Private m_datForceDNSUpdate         As Date

' NEW INTERFACE LANGUAGE /////////////////////////////////////////////////////////////
Private m_arrDynaCap(2)             As String
'limit and initialise for 51 clients tags
Private m_arrTagRules(50)           As String
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Const IPCheckInterval = 10  ' => every X minutes

Private lIntervalMin                As Long

Private Service(9)                  As String
Private Host(9)                     As String
Private User(9)                     As String
Private Pass(9)                     As String

Private mPlgObj                     As Object

'Private vars for sys tray
Private m_lHookID                   As Long
Private m_bSound                    As Boolean

'allows the use of 'tab' with in the SCI
Dim m_TabsStop()                    As Boolean

'Objects for data base explorer
Private m_objConn                   As Object
Private m_objRS                     As New ADODB.Recordset

'------------------------------------------------------------------------------
'Form events
'------------------------------------------------------------------------------
Private Sub Form_Activate()
1:  If picTab(5).Visible Then SCI_Focus
End Sub
Private Sub Form_Initialize()
    'Turn off nasty error messages which might lead to crashing (b/c of API calls)
1:  SetErrorMode &H1 Or &H2
End Sub
Private Sub Form_Load()
      On Error GoTo Err

3:      G_APPPATH = App.Path
4:      G_ERRORFILE = FreeFile

5:      AddLog "Aplication started.", 2
        
8:      Set g_objFileAccess = New clsFileAccess
  
    #If SVN Then
11:     G_LOGPATH = G_APPPATH & "\Logs\MsgDebugLog.log"
    #End If
    
        'Open error handling file
15:     Open G_APPPATH & "\Logs\Error.log" For Append As G_ERRORFILE
        
        'Set comma status
18:     m_blnCommaDecimal = InStrB(1, CStr(0.1), ",")

        'Open database
21:     Set m_objPermaCon = New Connection

23:     m_objPermaCon.ConnectionTimeout = 10
24:     m_objPermaCon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & G_APPPATH & "\DBs\userdb.mdb"
25:     m_objPermaCon.Mode = adModeReadWrite

27:     On Error GoTo NoDB
28:     m_objPermaCon.Open
29:     On Error GoTo Err
    
        'Wait until it is open before continuing
32:     Do Until m_objPermaCon.State = adStateOpen
33:        DoEvents
34:     Loop

        'Create our core global objects
37:     Set g_objFunctions = New clsFunctions
38:     Set g_colUsers = New clsHub
39:     Set g_objIPBans = New clsIPBans
40:     Set g_objRegistered = New clsRegistered
41:     Set g_objSettings = New clsSettings
42:     Set g_colCommands = New clsCommands
43:     Set g_objRegExps = New clsRegExps
        'PLAN
45:     Set g_colScheduler = New clsPlan
        'Initialise PTDCH messages, language
47:     Set g_colMessages = New clsDictionary

        'USER LANGUAGE
50:     Set g_colLanguages = New Collection

        ' hook window for sizing control
        ' Disable the following line if you will be debugging form.
54:     Call HookWin(Me.hWnd, G_HbWnd)
        
    #If Status Then
57:     Set g_objStatus = New clsStatus
    #End If

        'Create local objects
62:     Set m_objDetectIP = New clsHTTPDownload
63:     Set m_colFailedReg = New Collection
64:     Set m_colConnectAttempts = New Collection
65:     Set m_colRevConnects = New Collection

        'Set local vars
73:     m_lngBotsUB = -1

        'Add system tray icon
77:     Call SysTrayAdd
 
        'Load settings
81:     LoadDefaultSettings
82:     LoadSettings
  
85:     If g_objSettings.MagneticWin Then _
             Call Magnetic.AddWindow(frmHub.hWnd)  'Cool FX Windows
        
88:     tmrBackground.Interval = 60000 'Set background timer interval to every 20 mins

        'Prepare detect ip class
96:     m_objDetectIP.Host = "www.whatismyip.org"
97:     m_objDetectIP.Port = 80
    
       'Do extra actions
100:    If g_objSettings.AutoStart Then cmdButton_Click 1
    
        'Load DynamicIPServices for automatic IP updating
104:    lIntervalMin = IPCheckInterval 'to update services directly after start if neccessary
        
        'tmrUpdateIPs.Interval = 60 * 1000 'check every minute
107:    If g_objSettings.DynUpdate = True Then lblHolder(92).Visible = True
108:    If g_objSettings.DynUpdate = True Then LoadDynIPs

110:    If g_objSettings.EnabledScheduler = True Then
111:       lstPlan.Visible = True
102:    Else
113:       lstPlan.Visible = False
114:    End If

        'UpDate No-IP DNS at Starting PTDCH
117:    If g_objSettings.NoIPUpdateStartUp Then
118:        m_datForceDNSUpdate = Now
119:        UpdateDNSs
120:    End If

122:    Set Highlighter = New clsYHighlighter
124:    Highlighter.LoadDirectory G_APPPATH & "\Settings"
        
126:    Call IniDbExplorer

        'Hide TabControl
129:    tlbScript.Buttons.Item(15).Value = tbrPressed
130:    tbsScripts.Visible = False
131:
132: Exit Sub

133:
Err:
135:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub_Load()"
136:    Resume Next
    
138:
NoDB:
140:    MsgBoxCenter Me, "No database found; closing hub", vbOKOnly Or vbCritical, "PTDCH"
141:    Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

      On Error GoTo Err

      Dim lngPtr  As Long

      'Confirm exit if needed
7:    If g_objSettings.ConfirmExit And (UnloadMode <> vbAppWindows) Then
8:        If MsgBoxCenter(Me, g_colMessages.Item("msgExitDDCH"), vbYesNo Or vbQuestion Or vbDefaultButton2, g_colMessages.Item("msgConfirmExit")) = vbNo Then
9:           Cancel = 1
10:          Exit Sub
11:       End If
12:    End If
       
       'Confirm the unloading
15:    G_CUNLOAD = True

'17:    lvwScripts_DblClick 'Check if there was mod in scripts..
18:    Me.Hide
       'Close hub if it's still serving
20:    If m_blnServing Then SwitchServing

22:    tmrSysInfo.Enabled = False
       
       'Call unload event
25:    SEvent_UnloadMain

      'Save settings
28:    SaveSettings
       
      'Save Scripts Value in XML file
31:    frmScript.XmlBooleanSave
       
       'Remove all stray bot names from the database
34:    If Not m_lngBotsUB = -1 Then
35:        For lngPtr = 0 To m_lngBotsUB
36:            g_objRegistered.Remove m_arrBots(lngPtr).Name
37:        Next
38:        Erase m_arrBots()
39:    End If

41:    SysTrayRem

43: Exit Sub
    
45:
Err:
47:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub_QueryUnload()"
46:    Resume Next
End Sub
Private Sub Form_Terminate()
1:     Dim i As Integer
2:     On Error GoTo Err

        'Compress database if needed
5:      If g_objSettings.CompactDBOnExit Then
6:          Dim objEngine As JetEngine
      
8:          Set objEngine = New JetEngine
9:          objEngine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=.\DBs\userdb.mdb", "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=.\DBs\tempdb.mdb"
10:         Set objEngine = Nothing
12:         Kill ".\DBs\userdb.mdb"
            'Refresh .. and sleep semall time..
13:         Call Pause(100)
14:         Name ".\DBs\tempdb.mdb" As "DBs\userdb.mdb"
15:     End If

17:     On Error GoTo Err2

       'Clear process of the Forms
20:     Set frmBanName = Nothing
21:     Set frmBanPerm = Nothing
22:     Set frmBanTemp = Nothing
23:     Set frmCAccounts = Nothing
24:     Set frmCommand = Nothing
25:     Set frmEditScintilla = Nothing
26:     Set frmHelp = Nothing
27:     Set frmLoading = Nothing
28:     Set frmMulti = Nothing
29:     Set frmNewScript = Nothing
30:     Set frmProperties = Nothing
31:     Set frmReg = Nothing
32:     Set frmScript = Nothing
33:     Set frmSock = Nothing
34:     Set frmSocks = Nothing
35:     Set frmUpDate = Nothing
36:     Set frmUserInfo = Nothing

        'Clear all var process..
39:     Set g_objFunctions = Nothing
40:     Set g_colUsers = Nothing
41:     Set g_objIPBans = Nothing
42:     Set g_objRegistered = Nothing
43:     Set g_objSettings = Nothing
44:     Set g_colCommands = Nothing
45:     Set g_objRegExps = Nothing
46:     Set g_colScheduler = Nothing
47:     Set g_colMessages = Nothing
48:     Set g_colLanguages = Nothing
49:     Set m_objDetectIP = Nothing
50:     Set m_colFailedReg = Nothing
51:     Set m_colConnectAttempts = Nothing
52:     Set m_colRevConnects = Nothing
53:     Set g_objFileAccess = Nothing
54:     Set g_objStatus = Nothing
55:     Set g_colSWinsocks = Nothing
56:     Set g_colSVariables = Nothing
57:     Set Highlighter = Nothing
58:     Set m_sciSql = Nothing
60:     Set m_objConn = Nothing
61:     Set m_objRS = Nothing
        '
63:     Erase G_Highlighters()
64:     Erase sciMain()
        '
66:     Set frmHub = Nothing

68:     End ' hard end
    
70:   Exit Sub
71:
Err:
73:   MsgBox "An error occured while attempting to compress your database (" & Err.Number & " - " & Err.Description & ")", vbCritical, "PTDCH"
74:   End
75:
Err2:
77:   MsgBox "Error terminate hub (" & Err.Number & " - " & Err.Description & ")", vbCritical, "PTDCH"
78:   End
End Sub
Private Sub Form_Unload(Cancel As Integer)
1:    Dim frmLoop As Form
2:    Dim lngPtr  As Long
3:    Dim i As Integer
      
5:    On Error GoTo Err

      'Unload any other forms left over
8:    lngPtr = ObjPtr(Me)
  
10:   For Each frmLoop In Forms
11:        If Not ObjPtr(frmLoop) = lngPtr Then _
               Call Unload(frmLoop): Set frmLoop = Nothing
13:   Next
       
      'Close the error file
16:   Close G_ERRORFILE
    
      'Close connections to database
19:   m_objPermaCon.Close
20:   m_objConn.Close
      
      'This is absolutly an imperative line
23:   For i = 1 To UBound(sciMain)
24:        sciMain(i).Detach picSciMain(i)
25:   Next
26:   m_sciSql.Detach picSqlSCI

28:   Exit Sub
29:
Err:
31:   MsgBox "Error unloading hub - " & Err.Number & " (" & Err.Description & ")", vbCritical, "PTDCH"
32:   Resume Next
End Sub
Public Sub Form_Resize()
1:   On Error GoTo Err
3:   Dim i As Integer

     Select Case Me.WindowState
       
           '***************************
           Case vbMinimized
           '***************************
          
11:           mnuTray(0).Enabled = True
12:           mnuTray(1).Enabled = False

             'Hide the form if selected
15:           If g_objSettings.MinimizeTray And Me.WindowState = vbMinimized Then _
                   Me.Hide: Exit Sub
           
           '***************************
           Case Else 'vbNormal Or vbMaximized
           '***************************
           
22:           mnuTray(0).Enabled = False
23:           mnuTray(1).Enabled = True
           
25:           With tbsMenu
26:              If Not Me.Width < 9390 Then
27:                      .Width = Me.Width - 240
28:              End If
29:              If Not Me.Height < 5445 Then
30:                      .Height = Me.Height - 850
31:              End If
32:           End With
           
34:           For i = 0 To picTab.count - 1
35:              With picTab(i)
36:                   .Left = tbsMenu.Left + 80
37:                   .Width = tbsMenu.Width - 170
38:                   .Top = tbsMenu.Top + 360
39:                   .Height = tbsMenu.Height - 455
40:              End With
41:           Next

43:           picBordTab(0).Width = Me.Width
           
45:           With tbsSecurity
46:               If Not Me.Height < 5445 Then .Width = (Me.Width - 600)
47:               If Not Me.Height < 5445 Then .Height = (Me.Height - 1500)
48:           End With
           
50:           For i = 0 To picSTab.count - 1
51:              With picSTab(i)
52:                   .Left = tbsSecurity.Left + 80
53:                   .Width = tbsSecurity.Width - 170
54:                   .Top = tbsSecurity.Top + 360
55:                   .Height = tbsSecurity.Height - 455
56:              End With
57:           Next

59:           picBordTab(2).Width = picTab(1).Width
           
61:           With tbsInteractions
62:               If Not Me.Height < 5445 Then .Width = (Me.Width - 600)
63:               If Not Me.Height < 5445 Then .Height = (Me.Height - 1500)
64:           End With
           
66:           For i = 0 To picITab.count - 1
67:              With picITab(i)
68:                   .Left = tbsInteractions.Left + 80
69:                   .Width = tbsInteractions.Width - 170
70:                   .Top = tbsInteractions.Top + 360
71:                   .Height = tbsInteractions.Height - 455
72:              End With
73:           Next

75:           picBordTab(4).Width = picTab(2).Width
           
77:           With tabAdv
78:               If Not Me.Height < 5445 Then .Width = (Me.Width - 600)
79:               If Not Me.Height < 5445 Then .Height = (Me.Height - 1500)
80:           End With
           
82:           For i = 0 To picTabAdv.count - 1
83:              With picTabAdv(i)
84:                   .Left = tabAdv.Left + 80
85:                   .Width = tabAdv.Width - 170
86:                   .Top = tabAdv.Top + 360
87:                   .Height = tabAdv.Height - 455
88:              End With
89:           Next

91:           picBordTab(3).Width = picTab(3).Width
              
92:           With tbsHelp
93:               If Not Me.Height < 5445 Then .Width = (Me.Width - 600)
94:               If Not Me.Height < 5445 Then .Height = (Me.Height - 1500)
95:           End With
              
97:           For i = 0 To picHelp.count - 1
98:              With picHelp(i)
99:                   .Left = tbsHelp.Left + 80
100:                  .Width = tbsHelp.Width - 170
101:                  .Top = tbsHelp.Top + 360
102:                  .Height = tbsHelp.Height - 455
103:             End With
104:          Next

106:          picBordTab(1).Width = picTab(8).Width
              
108:          If lvwScripts.Visible Then
109:                 With tbsScripts
110:                      .Left = 60
111:                      .Width = (picTab(5).Width - lvwScripts.Width - 200)
112:                      .Top = 80 + tlbScript.Height
113:                      .Height = (picTab(5).Height - txtScriptError.Height - 230) - tlbScript.Height
114:                 End With
                  
116:                 With lvwScripts
117:                      .Left = (tbsScripts.Width + 120)
118:                      .Height = (picTab(5).Height - txtScriptError.Height - 230)
119:                      .Top = 60
120:                 End With
121:           Else
122:                 With tbsScripts
123:                      .Left = 60
124:                      .Width = (picTab(5).Width - 150)
125:                      .Top = 80 + tlbScript.Height
126:                      .Height = (picTab(5).Height - txtScriptError.Height - 230) - tlbScript.Height
127:                 End With
128:           End If

               If tbsScripts.Visible Then
130:               For i = 1 To picSciMain.count - 1
131:                     With picSciMain(i)
132:                          .Left = tbsScripts.Left + 15
133:                          .Width = tbsScripts.Width - 60
134:                          .Top = tbsScripts.Top + 330
135:                          .Height = tbsScripts.Height - 375
136:                     End With
137:               Next
138:           Else
139:               For i = 1 To picSciMain.count - 1
140:                     With picSciMain(i)
141:                          .Left = tbsScripts.Left
142:                          .Width = tbsScripts.Width
143:                          .Top = tbsScripts.Top
144:                          .Height = tbsScripts.Height
145:                     End With
146:               Next
147:           End If
               
149:           For i = 1 To UBound(sciMain)
150:                sciMain(i).SizeScintilla 0, 0, picSciMain(i).ScaleWidth / Screen.TwipsPerPixelX, (picSciMain(i).ScaleHeight / Screen.TwipsPerPixelY)
151:           Next
                
153:           With txtScriptError
154:                .Left = 60
155:                .Width = (picTab(5).Width - 150)
156:                .Top = (tbsScripts.Height) + 140 + tlbScript.Height
157:           End With

159:           With tlbScript
160:                .Left = 60
161:                .Width = tbsScripts.Width
162:           End With

164:           With txtNotePad
165:                .Left = 60
166:                .Top = 60
167:                .Height = (picHelp(1).Height - 140)
168:           End With
                 
170:           With lvwRegistered
171:                .Width = (picSTab(0).Width - 180)
172:                .Height = (picSTab(0).Height - 600)
173:           End With
                
175:           With tbsDbManager
176:                .Top = 60
177:                .Left = 60
178:                .Width = (picSTab(0).Width - 140)
179:                .Height = (picSTab(0).Height - 140)
180:           End With
           
182:           With picSqlSCI
183:                .Left = tbsDbManager.Left + 60
184:                .Width = tbsDbManager.Width - 150
185:                .Top = 840
186:                .Height = tbsDbManager.Height - 840
187:           End With
                
189:           m_sciSql.SizeScintilla 0, 0, picSqlSCI.ScaleWidth / Screen.TwipsPerPixelX, (picSqlSCI.ScaleHeight / Screen.TwipsPerPixelY)
     
191:           With dtgSql
192:                .Left = tbsDbManager.Left + 60
193:                .Width = tbsDbManager.Width - 150
194:                .Top = 440
195:                .Height = tbsDbManager.Height - 480
196:           End With
           
198:           With txtSqlErr
199:                .Left = 1680
200:                .Width = tbsDbManager.Width - 1750
201:           End With
           
203:           picBordTab(7).Width = picSTab(4).Width
                
205:           cmbRegistered.Left = (lvwRegistered.Width - cmbRegistered.Width + 100)
206:           txtDBRegCount.Left = (lvwRegistered.Width - cmbRegistered.Width - txtDBRegCount.Width)
     
      End Select
   
210:  Exit Sub
211:
Err:
212:    HandleError Err.Number, Err.Description, Erl & "|frmHub.Form_Resize()"
213:    Resume Next
End Sub
'------------------------------------------------------------------------------
'End Form events
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'SQL Explorer events
'------------------------------------------------------------------------------
Private Function RemoveSqlComments(strSqlCmd) As String
1:    On Error GoTo Err
      'Var.. for all string
3:    Dim a, b As Integer
      'Var.. for temp line
5:    Dim c, d As Integer
      '
7:    Dim strTemp As String
8:    Dim strLine As String
    
10:    b = 1
    
       'Remove sql comments from string
13:    For a = 1 To Len(strSqlCmd)
14:        If Mid(strSqlCmd, a, 1) = Chr(10) Then
               'Check if is sql comment in this line
16:            strLine = Mid(strSqlCmd, b, (a - b))
17:            If strLine <> Chr(10) Then
18:                For c = 1 To Len(strLine)
19:                    If CStr(Mid(strLine, c, 2)) = "--" Then
20:                        Exit For
21:                    Else
22:                        If CStr(Mid(strLine, c, 1)) <> " " Then
23:                            strTemp = strTemp & Mid(strSqlCmd, b, (a - b))
24:                            Exit For
25:                        End If
26:                    End If
27:                Next
28:            End If
29:            b = a + 1
30:        End If
31:    Next a

33:    RemoveSqlComments = Replace(strTemp, Chr(10), " ")

35:    Exit Function
36:
Err:
38:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.RemoveSqlComments(" & strSqlCmd & ")"
End Function
Private Sub cmdSql_Click(Index As Integer)
1:    On Error GoTo Err
2:    Dim i As Integer
3:    Dim strTemp As String
4:    Dim strSqlCmd As String

6:    Select Case Index

        'Run sql connection string
        '*******************************************************
        Case 0
        '*******************************************************
        
            'Check if slected text from SCI..
14:          strTemp = CStr(m_sciSql.GetSelText)
15:          If Not Len(strTemp) <> 1 Then _
                    strTemp = m_sciSql.Text

18:          strSqlCmd = RemoveSqlComments(strTemp)

20:          txtSqlErr.Text = ""

22:          On Error GoTo ErrSql
             
24:          If m_objRS.State <> 0 Then _
                    m_objRS.Close
                    
27:          m_objRS.Open strSqlCmd, _
                          m_objConn, _
                          adOpenKeyset, _
                          adLockOptimistic
          
32:          Set dtgSql.DataSource = m_objRS

34:          dtgSql.Refresh

36:          txtSqlErr.Text = "[" & Now & "] No syntax errors in SQL command."
             
        'Clear and add defaut sql string
        '*******************************************************
        Case 1
        '*******************************************************
        
43:          On Error GoTo Err
            
45:          Dim objRS As New ADODB.Recordset
46:          Set objRS = m_objConn.OpenSchema(adSchemaTables)

48:          strTemp = "-- Database Tables (userdb.mdb)" & vbNewLine
                  
50:          Do While Not objRS.EOF
51:              i = objRS.Fields.count

53:              If UCase(Left(objRS.Fields("TABLE_NAME"), 4)) <> "MSYS" Then _
                       strTemp = strTemp & "--" & vbTab & objRS.Fields("TABLE_NAME") & vbNewLine

56:              objRS.MoveNext
57:          Loop
               
59:          strTemp = strTemp & vbNewLine & "-- Demo (Show all Registered Users) " & vbNewLine & _
                      "SELECT UsrClass.UserName, UsrStatic.Pass, UsrClass.Class, ClassTypes.Name, UsrStatic.RegedBy, UsrStatic.RegDate, UsrDynamic.LastLogin, UsrDynamic.LastIP " & vbNewLine & _
                      "FROM UsrStatic INNER JOIN (ClassTypes INNER JOIN (UsrClass INNER JOIN UsrDynamic ON UsrClass.UserName = UsrDynamic.UserName) ON ClassTypes.ID = UsrClass.Class) ON (UsrStatic.UserName = UsrDynamic.UserName) AND (UsrStatic.UserName = UsrClass.UserName) " & vbNewLine & _
                      "ORDER BY UsrClass.UserName" & vbNewLine
                      
64:          m_sciSql.Text = strTemp
65:          m_sciSql.ClearUndoBuffer

      End Select
     
69:   m_sciSql.SetFocus
     
71:   Exit Sub
72:
ErrSql:
74:
75:   txtSqlErr.Text = "[" & Now & "] Error: " & Err.Description
76:   Err.Clear
77:   m_sciSql.SetFocus
78:   Exit Sub
79:
Err:
81:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.cmdSql_Click(" & Index & ")"
82:   m_sciSql.SetFocus
End Sub
Public Sub IniDbExplorer()

2:    On Error GoTo Err

      'Create code editor *****************************************************
5:    Set m_sciSql = New clsYScintilla
    
7:    m_sciSql.CreateScintilla picSqlSCI
8:    m_sciSql.SetFixedFont "Courier New", 10

      ' Give the scrollbar a nice long width to handle a long line which may
      ' occur.
12:   m_sciSql.ScrollWidth = 10000

      'This is absolutly an imperative line
15:   m_sciSql.Attach picSqlSCI

17:   m_sciSql.LineNumbers = True
18:   m_sciSql.AutoIndent = True

20:   m_sciSql.SetMarginWidth MarginLineNumbers, 50
   
22:   Call Highlighter.SetHighlighterBasedOnExt(m_sciSql, ".sql")
      '************************************************************************
  
      'Create new connection *****************************************************
28:   Dim strTemp As String
  
30:   Set m_objConn = CreateObject("Adodb.Connection")
  
32:   m_objConn.CursorLocation = adUseClient
33:   m_objConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & G_APPPATH & "\DBs\userdb.mdb" & ";Persist Security Info=False"
34:   m_objConn.Open
        
36:   If g_objFileAccess.FileExists(G_APPPATH & "\Settings\bdManager.sql") Then
37:       strTemp = g_objFileAccess.ReadFile(G_APPPATH & "\Settings\bdManager.sql")
38:   End If
     
40:   If strTemp <> "" Then
41:       m_sciSql.Text = strTemp
42:       m_sciSql.ClearUndoBuffer
43:   Else
44:       Call cmdSql_Click(1)
45:   End If

     '************************************************************************
  
49:  Exit Sub
50:
Err:
52:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.IniDbExplorer()"
End Sub
'------------------------------------------------------------------------------
'End SQL Explorer events
'------------------------------------------------------------------------------

Private Sub SCI_Focus()
1:  Dim i As Integer
2:  For i = 1 To picSciMain.count - 1
3:        If picSciMain(i).Visible Then Exit For
4:  Next
5:  sciMain(i).SetFocus
End Sub

Private Sub lvwScripts_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    'Check if we should reset or stop
2:    If Item.Checked Then _
           frmScript.SReset Item.Index, False, True _
      Else frmScript.SStop Item.Index
End Sub

Private Sub lvwScripts_DblClick()
1:    On Error GoTo Err
2:    Dim i As Integer
3:    For i = 1 To lvwScripts.ListItems.count
4:        If lvwScripts.ListItems(i).Selected Then
5:            tbsScripts.Tabs(i).Selected = True
6:            Exit Sub
7:        End If
8:    Next
9:    Exit Sub
10:
Err:
11:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lvwScripts_DblClick()"
End Sub

Private Sub lvwScripts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    On Error GoTo Err
2:    Dim i As Integer
3:    Dim blnIsSelected As Boolean

      For i = 1 To lvwScripts.ListItems.count
4:        If lvwScripts.ListItems(i).Selected Then
5:            blnIsSelected = True
6:            Exit For
7:        End If
8:    Next
  
10:   If blnIsSelected Then
11:       mnuScripts(0).Enabled = True
12:       mnuScripts(2).Enabled = True
13:       mnuScripts(8).Enabled = True
14:   Else
15:       mnuScripts(0).Enabled = False
16:       mnuScripts(2).Enabled = False
17:       mnuScripts(8).Enabled = False
18:   End If

      'Popup menu if left button is pressed
21:   If Button = 2 Then PopupMenu mnuPopUp(9)

23:   Exit Sub
24:
Err:
26:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lvwScripts_MouseDown()"
End Sub

Private Sub mnuCodeRTB1_Click(Index As Integer)
1:    On Error GoTo Err
2:    Dim i As Integer
   
4:    For i = 1 To picSciMain.count - 1
5:        If picSciMain(i).Visible Then Exit For
6:    Next
    
8:    Select Case Index
        Case 0 'Insert Date/Time
10:            sciMain(i).SelText = Now
        Case 1 'Script Info
12:              Dim strInfo, strSize, strName, strString As String

14:              strString = sciMain(i).Text
15:              strName = frmHub.tbsScripts.Tabs(i).Key
                 
17:              strInfo = "Script strInfo - " & strName & vbCrLf & vbCrLf & _
                           "Characters: " & Len(strString) & vbCrLf & _
                           "Lines: " & CharCount(strString, vbCrLf) + 1 & vbCrLf & _
                           "Words: " & CharCount(strString, " ") + 1

22:              strSize = Len(strString)

                  'Calculate Size
25:              If strSize > 1000 Then
26:                 strSize = FormatNumber(strSize / 1024, 2)

28:                 MsgBoxCenter Me, strInfo & vbCrLf & vbCrLf & _
                            "The size of the file is:" & vbCrLf & _
                            FormatNumber(strSize, 2) & " Kb.", vbOKOnly + vbInformation
31:              Else
32:                 MsgBoxCenter Me, strInfo & vbCrLf & vbCrLf & _
                            "The size of the file is:" & vbCrLf & _
                            FormatNumber(strSize, 2) & " bytes.", vbOKOnly + vbInformation
35:              End If

        Case 4 'Save as..
38:         Dim cD As New clsCommonDialog
39:         Dim sFile As String

40:         sFile = frmHub.tbsScripts.Tabs(i).Key

42:         If (cD.VBGetSaveFileName(sFile, _
                  Filter:="VBScript (*.script)|.script|VBScript (*.vbs)|.script|All Files (*.*)|*.*", _
                      DefaultExt:="htm", _
                         Owner:=Me.hWnd)) Then
                '
47:             g_objFileAccess.WriteFile sFile, sciMain(i).Text
48:         End If
        Case 6 'Clear Undo Buffer
50:         sciMain(i).ClearUndoBuffer
51:         lvwScripts.ListItems(i).SubItems(3) = CStr(sciMain(i).Modified)
      End Select

54:   SCI_Focus

56:   Exit Sub
57:
Err:
59:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuCodeRTB1_Click(" & Index & ")"
End Sub

Private Sub mnuCodeRTB3_Click(Index As Integer)
1:    On Error GoTo Err
      Select Case Index
          Case 0 'VBScript documentation
4:              g_objFunctions.ShellExec "http://msdn2.microsoft.com/en-us/library/t0aew7h6.aspx"
          Case 1 'JScript documentation
6:              g_objFunctions.ShellExec "http://msdn2.microsoft.com/en-us/library/hbxc2t98(vs.71).aspx"
      End Select
    
9:    Exit Sub
10:
Err:
11:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuCodeRTB3_Click(" & Index & ")"
End Sub

Private Sub tbsDbManager_Click()

2:   On Error GoTo Err
 
4:   Select Case tbsDbManager.SelectedItem.Index
        Case 1
6:          dtgSql.Visible = False
7:          cmdSql(0).Visible = True
8:          cmdSql(1).Visible = True
9:          txtSqlErr.Visible = True
10:         picSqlSCI.Visible = True
11:         m_sciSql.SetFocus
        Case 2
13:         dtgSql.Visible = True
14:         cmdSql(0).Visible = False
15:         cmdSql(1).Visible = False
16:         txtSqlErr.Visible = False
17:         picSqlSCI.Visible = False
     End Select
    
20: Exit Sub
21:
Err:
23:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tbsDbManager_Click()"
24:    Resume Next
End Sub

Private Sub tbsScripts_Click()

2:   On Error GoTo Err
 
4:   Dim i, i2 As Integer
   
6:   i2 = Val(tbsScripts.SelectedItem.Index)
7:   If picSciMain(i2).Visible = True Then Exit Sub
    
9:   For i = 1 To picSciMain.count - 1
10:     picSciMain(i).Visible = False
11:  Next i
   
13:  i = Val(tbsScripts.SelectedItem.Index)
14:  picSciMain(i).Visible = True

16:  If frmEditScintilla.Visible Then frmEditScintilla.Visible = False

18:  Exit Sub
19:
Err:
21:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tbsScripts_Click()"
     Resume Next
End Sub

Private Sub tlbScript_ButtonClick(ByVal Button As ComctlLib.Button)
1:      Dim i As Integer
2:      Dim intIndex As Integer

4:      On Error GoTo Err

6:      For i = 1 To picSciMain.count - 1
7:            If picSciMain(i).Visible Then intIndex = i: Exit For
8:      Next
        
10:     Select Case CStr(Button.Key)
            Case "Undo"
12:              sciMain(intIndex).Undo
            Case "Redo"
14:              sciMain(intIndex).Redo
            Case "Find"
16:              sciMain(intIndex).EditSciText 1, Me
            Case "Replace"
18:             sciMain(intIndex).EditSciText 2, Me
            Case "GoToLine"
20:             sciMain(intIndex).EditSciText 3, Me
            Case "Save Only"
22:             Call frmScript.SSave(i)
            Case "Save and Reset Script"
24:             Call frmScript.SReset(i, True, True)
            Case "Clear"
26:             sciMain(intIndex).Text = ""
            Case "Hide Scripts"
28:             If Button.Value Then _
                     lvwScripts.Visible = False _
                Else lvwScripts.Visible = True
31:             Call Form_Resize
            Case "Hide Tabs"
33:             If Button.Value Then _
                     tbsScripts.Visible = False _
                Else tbsScripts.Visible = True
36:             Call Form_Resize
            Case "Enabled Tabs"
38:             If Button.Value Then
                    'Remove all tabs..
40:                  ReDim m_TabsStop(0 To Controls.count - 1) As Boolean
41:                  For i = 0 To Controls.count - 1
42:                     On Error Resume Next
43:                     m_TabsStop(i) = Controls(i).TabStop
44:                     Controls(i).TabStop = False
45:                  Next
46:                  On Error GoTo Err
47:             Else
                     'Add All Tabs..
49:                  For i = 0 To Controls.count - 1
50:                     On Error Resume Next
51:                     Controls(i).TabStop = m_TabsStop(i)
52:                  Next
53:                  On Error GoTo Err
54:             End If
            Case "New"
56:             frmNewScript.Show vbModal, Me
            Case "Menu"
58:             PopupMenu frmHub.mnuPopUp(10), 0, 4800, 800
        End Select

61:     Select Case Button.Key
            Case "Find", "Replace", "GoToLine", "New"
            Case Else: SCI_Focus
        End Select

66:   Exit Sub
67:
Err:
69:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tlbScript_ButtonClick(" & Button.Key & ")"
End Sub

Private Sub cmbSkin_Click()
   
      On Error GoTo Err
   
4:    cmdSkin(0).Enabled = True
5:    cmdSkin(1).Enabled = True
   
      Select Case cmbSkin.Text
         Case "01-Defaut"
9:          g_objSettings.lngSkin = 1
10:          cmdSkin(0).Enabled = False
         Case "02-Cyan Blue"
12:          g_objSettings.lngSkin = 2
         Case "03-Cyan Green"
14:          g_objSettings.lngSkin = 3
         Case "04-Metallic"
16:          g_objSettings.lngSkin = 4
         Case "05-Metallic Blue"
18:          g_objSettings.lngSkin = 5
         Case "06-Metallic Green"
20:          g_objSettings.lngSkin = 6
         Case "07-Metallic Navy Blue"
22:          g_objSettings.lngSkin = 7
         Case "08-Metallic Oliver"
24:          g_objSettings.lngSkin = 8
         Case "09-Texture Grain"
26:          g_objSettings.lngSkin = 9
         Case "10-Texture Spater"
28:          g_objSettings.lngSkin = 10
         Case "11-Texture Tiles"
30:          g_objSettings.lngSkin = 11
         Case "12-Texture Toxedo"
32:          g_objSettings.lngSkin = 12
         Case "13-Blue Berry"
34:          g_objSettings.lngSkin = 13
         Case "14-Glace Table"
36:          g_objSettings.lngSkin = 14
         Case "15-Pink"
38:          g_objSettings.lngSkin = 15
         Case "16-Gun Blue"
40:          g_objSettings.lngSkin = 16
         Case "17-Gun Metal"
42:          g_objSettings.lngSkin = 17
43:          cmdSkin(1).Enabled = False
      End Select
   
46:   Dim i As Integer
47:   On Error Resume Next
      'Refresh all picture box .. very fast
49:   For i = 0 To picTab.count - 1: picTab(i).Refresh: Next i
50:   For i = 0 To picSTab.count - 1: picSTab(i).Refresh: Next i
51:   For i = 0 To picITab.count - 1: picITab(i).Refresh: Next i
52:   For i = 0 To picTabAdv.count - 1: picTabAdv(i).Refresh: Next i
53:   For i = 0 To picHelp.count - 1: picHelp(i).Refresh: Next i
54:   For i = 0 To picBordTab.count - 1: picBordTab(i).Refresh: Next i
55:   For i = 0 To picInfo.count - 1: picInfo(i).Refresh: Next i
56:   For i = 0 To picStatus.count - 1: picStatus(i).Refresh: Next i

58:   Call Form_Paint
59:   Me.Refresh
   
61:   Exit Sub

63:
Err:
65:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.cmbSkin_Click()"
End Sub

Private Sub cmdButton_Click(Index As Integer)
1:    Dim strTemp(3) As String
2:    Dim frm As New frmMulti

4:    On Error GoTo Err

      Select Case Index
        Case 0 'Check for updates
8:             frmUpDate.Show vbModal, Me
        Case 1 'Start / stop serving
10:            SwitchServing
        Case 2 'Redirect users
12:            NextRedirect
13:                With frm
14:                    .Label1.Caption = g_colMessages.Item("msgEnterRedirUsersAddress")
15:                    .Caption = g_colMessages.Item("msgRedirUsers")
16:                    .txtStr.Text = g_objSettings.RedirectIP
17:                    .Show vbModal, Me
18:                    strTemp(0) = .txtStr.Text
19:                End With
20:                Set frm = Nothing
               'Check and see if they pressed cancel
21:            If LenB(strTemp(0)) Then
                Select Case MsgBoxCenter(Me, g_colMessages.Item("msgRedirAll"), vbYesNo Or vbQuestion, g_colMessages.Item("msgRedirUsers"))
                    Case vbYes
24:                        g_colUsers.RedirectAll strTemp(0)
                    Case vbNo
26:                        g_colUsers.RedirectNonOps strTemp(0)
                    Case Else
28:                        Exit Sub
29:                End Select
                'Raise script event
31:                SEvent_StartedRedirecting
32:            End If
        Case 3 'Save settings
34:            SaveSettings
        Case 4 'Reload settings
36:            LoadDefaultSettings
37:            LoadSettings
        Case 5 'Detect IP
39:            strTemp(0) = DetectHubIP
40:            With frm
41:                .Label1.Caption = g_colMessages.Item("msgDetectIP")
42:                .Caption = "IP"
43:                .cmdCancel.Visible = False
44:                If Not strTemp(0) = vbNullString Then
45:                     .txtStr.Text = DetectHubIP
46:                Else 'change message to "try again" ?
47:                     .txtStr.Text = g_colMessages.Item("msgGettingIP")
48:                End If
49:                .Show vbModal, Me
50:            End With
51:            Set frm = Nothing
        Case 6 'Force UpDate
52:            m_datForceDNSUpdate = Now
53:            UpdateDNSs
        Case 7, 8, 9 ' Mass Messages..
        
56:            strTemp(0) = g_colMessages.Item("msgEnterPM")

58:            If Index = 7 Then 'Mass Messages to All
59:                 strTemp(1) = g_colMessages.Item("msgMassMsg")
60:                 strTemp(2) = g_objSettings.MassMessage
61:                 AddLog "Mass Messages To All :" & strTemp(2), 6
62:            ElseIf Index = 8 Then 'Mass Messages to Ops
63:                 strTemp(1) = g_colMessages.Item("msgMassMsgOp")
64:                 strTemp(2) = g_objSettings.OpMassMessage
65:                 AddLog "Mass Messages To Ops :" & strTemp(2), 6
66:            ElseIf Index = 9 Then 'Mass Messages to UnReg
67:                 strTemp(1) = g_colMessages.Item("msgMassMsgUnReg")
68:                 strTemp(2) = g_objSettings.UnRegMassMessage
69:                 AddLog "Mass Messages To UnReg :" & strTemp(2), 6
70:            End If
71:            With frm
72:                 .Label1.Caption = strTemp(0)
73:                 .Caption = strTemp(1)
74:                 .Height = 2280
75:                 .txtStr.Visible = False
76:                 .txtStrMultiLine.Visible = True
77:                 .cmdCancel.Top = 1440
78:                 .cmdOK.Top = 1440
79:                 .txtStrMultiLine = strTemp(2)
80:                 .Show vbModal, Me
81:                 strTemp(3) = .txtStrMultiLine.Text
82:            End With
83:            Set frm = Nothing

85:            If Not strTemp(3) <> "" Then Exit Sub

87:            If Index = 7 Then 'Mass Messages to All
88:                 g_objSettings.MassMessage = strTemp(3)
89:                 g_colUsers.SendPrivateToAll g_objSettings.BotName, strTemp(3)
90:            ElseIf Index = 8 Then 'Mass Messages to Ops
91:                 g_objSettings.OpMassMessage = strTemp(3)
92:                 g_colUsers.SendPrivateToOps g_objSettings.BotName, strTemp(3)
93:            ElseIf Index = 9 Then 'Mass Messages to UnReg
94:                 g_objSettings.UnRegMassMessage = strTemp(3)
95:                 g_colUsers.SendPrivateToUnReg g_objSettings.BotName, strTemp(3)
96:            End If

98:            SEvent_MassMessage strTemp(3)

100:    End Select

102:    Exit Sub
    
104:
Err:
106:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.cmdButton_Click(" & Index & ")"
End Sub

Private Sub cmdConvDatabase_Click()
1:   frmCAccounts.Show vbModal, Me
End Sub

Private Sub cmdDB_Click(Index As Integer)
1:  Dim strClass    As String
2:  Dim lvwItem     As ListItem
3:  Dim lvwItems    As ListItems
4:  On Error GoTo Err
    
6:  If Index = 4 Or Index = 6 Or Index = 8 Then If lvwRegistered.ListItems.count = 0 Then Exit Sub
7:  If Index = 5 Or Index = 7 Then If lvwBans.ListItems.count = 0 Then Exit Sub
    
9:    If Index = 4 Or Index = 6 Or Index = 8 Then '---------------------------------
10:       Set lvwItem = lvwRegistered.SelectedItem
11:       Set lvwItems = lvwRegistered.ListItems
       'Check if selected
13:       If lvwItem.Selected Then
14:            Select Case CLng(lvwItems(lvwItem.Index).SubItems(2))
                    'Case -1: strClass = "-1 = Locked"
                    'Case 0: strClass = "0 = Unknown"
                    'Case 1: strClass = "1 = Regular"
                    'Case 2: strClass = "2 = Mentored"
                    Case 3: strClass = "3 = Registered"
                    Case 4: strClass = "4 = Invisible"
                    Case 5: strClass = "5 = VIP"
                    Case 6: strClass = "6 = Operator"
                    Case 7: strClass = "7 = Invisible Operator"
                    Case 8: strClass = "8 = Super Operator"
                    Case 9: strClass = "9 = Invisible Super Operator"
                    Case 10: strClass = "10 = Admin"
                    Case 11: strClass = "11 = Invisible Admin"
                    Case Else: strClass = "2 = Mentored"
                End Select
30:       End If
31:    End If '---------------------------------------------------------
    
33:    Select Case Index '----------------------------------------------
            Case 0: Call DBGetRegRecord 'Refresh Reg
            Case 1: Call DBGetBanRecord 'Refresh Ban
            Case 2 'Add Reg ---------------------------------------------
37:             With frmReg
38:                  Load frmReg
39:                 .Tag = "Add"
40:                 .cmbClass = "3 = Registered"
41:                 .InicializeReg 'Perpare Form
42:                 .Show vbModal, Me
43:                 Pause (500): DBGetRegRecord 'Refresh Reg
44:             End With
            Case 3 'Add Ban ---------------------------------------------
45:             With frmBanName
46:                 .Tag = "Add"
47:                 .InicializeBan
48:                 .Show vbModal, Me
49:                 Pause (1000): DBGetBanRecord 'Refresh Ban
50:             End With
            Case 4 'Rem Reg ---------------------------------------------
52:             Set lvwItem = lvwRegistered.SelectedItem
            'Check if selected
54:             If lvwItem.Selected Then _
                    g_objRegistered.Remove lvwItem.Text: _
                      lvwRegistered.ListItems.Remove CInt(lvwItem.Index)
            Case 5 'Rem Ban ---------------------------------------------
58:             Set lvwItem = lvwBans.SelectedItem
            'Check if selected
60:             If lvwItem.Selected Then _
                   g_objRegistered.Remove lvwItem.Text: _
                     lvwBans.ListItems.Remove CInt(lvwItem.Index)
            Case 6 ' Edit Reg -------------------------------------------
            'Check if selected
66:             If lvwItem.Selected Then
67:                With frmReg
68:                    Load frmReg
69:                    .Tag = "Edit"
70:                    .txtPass.Text = lvwItems(lvwItem.Index).SubItems(1)  'Pass
71:                   .txtName.Text = lvwItem.Text
72:                   .cmbClass = strClass
73:                   .InicializeReg
74:                   .Show vbModal, Me
75:                End With
76:             End If
            Case 7 'Rename Ban ------------------------------------------
78:             Set lvwItem = lvwBans.SelectedItem
            'Check if selected
80:             If lvwItem.Selected Then
81:                With frmBanName
82:                    .Tag = "Rename"
83:                    .txtName.Text = lvwItem.Text
84:                    .txtName.Tag = lvwItem.Text
85:                    .txtReason.Text = lblHolder(50).Caption 'Reason
86:                    .InicializeBan
87:                    .Show vbModal, Me
88:                End With
89:             End If
            Case 8 'Rename Reg ------------------------------------------
            'Check if selected
92:               If lvwItem.Selected Then
93:                  With frmReg
94:                       Load frmReg
95:                      .Tag = "Rename"
96:                      .txtPass.Text = lvwItems(lvwItem.Index).SubItems(1)  'Pass
97:                      .txtName.Text = lvwItem.Text
98:                      .txtName.Tag = lvwItem.Text 'Used in rename.. firts name
99:                      .cmbClass = strClass
100:                      .InicializeReg
101:                      .Show vbModal, Me
102:                End With
103:             End If
        End Select '-----------------------------------------------------

106: Exit Sub
    
108:
Err:
110:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.cmdDB_Click(" & Index & ")"
End Sub

Private Sub cmdPlugins_Click()
1:   On Error GoTo Err

3:   If lvwPlugins.SelectedItem.Selected = True Then
 
5:      Set mPlgObj = CreateObject(lvwPlugins.SelectedItem.Key)

7:      If mPlgObj.loadplug <> 1 Then
8:         MsgBoxCenter Me, "There was an error while loading the plugin.", vbCritical
9:         Exit Sub
10:     Else
11:        mPlgObj.LoadPlugin frmHub
12:     End If
13:  Else
15:      cmdPlugins.Enabled = False
16:  End If
   
18:   Exit Sub
19:
Err:
21:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.cmdPlugins()"
End Sub

Private Sub cmdPopup_Click(Index As Integer)
1:   g_objFunctions.ShowBallon "PT Direct Connect Hub " & vbVersion, "Created by fLaSh", Index, True
End Sub

Private Sub cmdSkin_Click(Index As Integer)
1:   On Error Resume Next
     Select Case Index
        Case 0
4:         cmbSkin.Text = cmbSkin.List(cmbSkin.ListIndex - 1)
        Case 1
6:         cmbSkin.Text = cmbSkin.List(cmbSkin.ListIndex + 1)
     End Select
End Sub

Sub CreateDynIPsXML()
    Dim strTemp As String
    Dim intFF As Integer
    
    On Error GoTo Err
    
6:    strTemp = G_APPPATH & "\Settings\DynIPs.xml"

8:    intFF = FreeFile

       'Append to file
11:    Open strTemp For Output As intFF

12:    Print #intFF, "<DynIPs>"
13:    Print #intFF, vbTab & "<!-- Service,Host,User,Pass -->"
14:    Print #intFF, vbTab & "<!-- if file does not exist, updating is disabled -->"
15:    Print #intFF, vbTab & "<0></0>"
16:    Print #intFF, vbTab & "<1></1>"
17:    Print #intFF, vbTab & "<2></2>"
18:    Print #intFF, vbTab & "<3></3>"
19:    Print #intFF, vbTab & "<4></4>"
20:    Print #intFF, vbTab & "<5></5>"
21:    Print #intFF, vbTab & "<6></6>"
22:    Print #intFF, vbTab & "<7></7>"
23:    Print #intFF, vbTab & "<8></8>"
24:    Print #intFF, vbTab & "<9></9>"
25:    Print #intFF, vbTab & "<!-- More than 10 services will be ignored -->"
26:    Print #intFF, "</DynIPs>";

28:    Close intFF

30:    Exit Sub
31:
Err:
33:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.CreateDynIPsXML()"
End Sub

Private Sub LoadDynIPs()
' should go in loadsettings sub
    
3:    Dim objXML          As clsXMLParser
4:    Dim objNode         As clsXMLNode
5:    Dim colNodes        As Collection
6:    Dim colAttributes   As Collection
    
8:    Dim strTemp     As String
9:    Dim strValues() As String

11:    On Error Resume Next
    
13:    strTemp = G_APPPATH & "\Settings\DynIPs.xml"
        
15:    If Not (g_objFileAccess.FileExists(strTemp)) Then
16:            CreateDynIPsXML
17:            Exit Sub
18:        End If

20:    Set objXML = New clsXMLParser

22:    objXML.Data = g_objFileAccess.ReadFile(strTemp)
23:    objXML.Parse

25:    Set colNodes = objXML.Nodes(1).Nodes

'    'Just in case...
28:    On Error Resume Next

30:    For Each objNode In colNodes
31:        strTemp = objNode.Value
32:        If (strTemp <> "") And (Val(objNode.Name) < 10) Then
33:            strValues = Split(strTemp, ",")
34:            If UBound(strValues) = 3 Then
35:                Service(objNode.Name) = strValues(0)
36:                Host(objNode.Name) = strValues(1)
37:                User(objNode.Name) = strValues(2)
38:                Pass(objNode.Name) = strValues(3)
'                tmrUpdateIPs.Enabled = True
40:            End If
41:        End If
42:    Next

44:    strTemp = CStr(UBound(Service))
45:    objXML.Clear
46:    Set objNode = Nothing
47:    Set colNodes = Nothing
End Sub

Private Sub LabelsURL_Click(Index As Integer)
1:   On Error Resume Next
     Select Case Index
        Case 0 'Send e-mail
2:         On Error Resume Next
3:         ShellExecute Me.hWnd, "open", "mailto:carlosferreiracarlos@hotmail.com?subject=About the PT DC Hub V." & vbVersion & "...&body=I have tested the software and...", 0&, 0&, vbNormal
        Case 1 'GoTo HomePage
4:         On Error Resume Next
5:         ShellExecute Me.hWnd, "open", "http://HublistChecker.pt.vu/", "", "", 3   'SW_SHOWMAXIMIZED
   End Select
End Sub

Private Sub LabelsURL_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'1: LabelsURL(Index).ForeColor = &HFFC0C0
End Sub

Private Sub lblHolder_Change(Index As Integer)
1:    On Error GoTo Err
2:    Dim strTmp(1) As String

      'Connected Users
4:   If Index = 55 Then
5:        If Len(g_objSettings.HubName) > 22 Then _
               strTmp(0) = Left(g_objSettings.HubName, 20) & ".." _
          Else strTmp(0) = g_objSettings.HubName
          
9:        If m_blnServing Then
10:             strTmp(1) = "PT DC Hub " & vbVersion & vbNewLine & vbNewLine & _
                              strTmp(0) & vbNewLine & _
                              lblHolder(45).Caption & lblHolder(55).Caption
13:       Else
14:             strTmp(1) = "PT DC Hub " & vbVersion & vbNewLine & vbNewLine & _
                              strTmp(0)
16:       End If
          
18:       stbMain.Panels(4).Text = lblHolder(55).Caption
19:       SysTrayUpDate strTmp(1)
20:   End If
    
      'Connected Op
23:   If Index = 56 Then _
          stbMain.Panels(6).Text = lblHolder(56).Caption
    
      'Shared Total
27:   If Index = 57 Then _
          stbMain.Panels(5).Text = lblHolder(57).Caption
29:
       
31:   Exit Sub
32:
Err:
33:   HandleError Err.Number, Err.Description, Erl & "|frmHub.lblHolder_Change(" & Index & ")"
End Sub

Private Sub lvwBans_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
1:  On Error GoTo Err
    
    'Sort listview by column clicked
4:   lvwBans.SortKey = ColumnHeader.Index - 1
5:   lvwBans.SortOrder = IIfLng(lvwBans.SortOrder, lvwAscending, lvwDescending)
6:   lvwBans.Sorted = True
    
8:   Exit Sub
10:
Err:
11:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lvwBans_ColumnClick()"
End Sub

Private Sub lvwBans_ItemClick(ByVal Item As ComctlLib.ListItem)
1:   lblHolder(50).Caption = CStr(Item.Tag)
End Sub

Private Sub lvwBans_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    If Button = 2 Then PopupMenu mnuPopUp(4)
End Sub

'------------------------------------------------------------------------------
'GUI related events
'------------------------------------------------------------------------------

Private Sub lvwCommands_DblClick()
1:    Dim lvwItem As ListItem
2:    Dim strKey  As String
    
4:    On Error GoTo Err
    
6:    Set lvwItem = lvwCommands.SelectedItem
    
    'Make sure an item is selected
9:    If ObjPtr(lvwItem) Then
10:        Load frmCommand
        
        'Update GUI
13:        strKey = lvwItem.Text
14:        frmCommand.txtTrigger.Text = strKey
15:        frmCommand.Tag = lvwItem.Text
16:        frmCommand.cmbClass.Text = lvwItem.SubItems(1)
17:        frmCommand.txtDescription.Text = g_colCommands(strKey).Description
18:        frmCommand.chkEnabled.Value = Abs(CBool(lvwItem.SubItems(2)))
    
20:        frmCommand.Show vbModal, Me
21:    End If



25:    Exit Sub
    
27:
Err:
29:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lvwCommands_DblClick()"
End Sub

Private Sub lstPlan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    If Button = 2 Then PopupMenu mnuPopUp(8)
End Sub

Private Sub lvwPlugins_ItemClick(ByVal Item As ComctlLib.ListItem)
1:    If Not Item.Key = "" Then
2:       cmdPlugins.Enabled = True
3:    Else
4:       cmdPlugins.Enabled = False
5:    End If
End Sub

Private Sub lvwPlugins_LostFocus()
1:   If lvwPlugins.SelectedItem.Selected = False Then _
           cmdPlugins.Enabled = False
End Sub

Private Sub lvwRegistered_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
1:  On Error GoTo Err
    
    'Sort listview by column clicked
4:   lvwRegistered.SortKey = ColumnHeader.Index - 1
5:   lvwRegistered.SortOrder = IIfLng(lvwRegistered.SortOrder, lvwAscending, lvwDescending)
6:   lvwRegistered.Sorted = True
    
8:   Exit Sub
10:
Err:
11:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lvwRegistered_ColumnClick()"
End Sub

Private Sub lvwRegistered_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    If Button = 2 Then PopupMenu mnuPopUp(1)
End Sub

#If Status Then
    Private Sub lvwUsers_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
1:      On Error GoTo Err
    
        'Sort listview by column clicked
4:      lvwUsers.SortKey = ColumnHeader.Index - 1
5:      lvwUsers.SortOrder = IIfLng(lvwUsers.SortOrder, lvwAscending, lvwDescending)
6:      lvwUsers.Sorted = True
    
8:      Exit Sub
10:
Err:
11:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lvwTempIPBan_ColumnClick()"
    End Sub
    Private Sub lvwUsers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:        If Button = 2 Then PopupMenu mnuPopUp(7)
    End Sub
#End If

Private Sub lstTagsEx_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    If Button = 2 Then PopupMenu mnuPopUp(5)
End Sub

Private Sub lvwPermIPBan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    If Button = 2 Then PopupMenu mnuPopUp(3)
End Sub

Private Sub lvwTempIPBan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:    If Button = 2 Then PopupMenu mnuPopUp(2)
End Sub

Private Sub lstTagsDef_Click()
' ------------------------ NEW MOD INTERFACE LANGUAGE ------------------------
2:    If lstTagsDef.ListIndex = -1 Then
3:        txtTagRules.Text = Replace(m_arrTagRules(10), "%[LF]", vbNewLine)
4:    Else
5:        txtTagRules.Text = Replace(m_arrTagRules(lstTagsDef.ListIndex), "%[LF]", vbNewLine)
6:    End If

'    Select Case lstTagsDef.ListIndex
        ' ++
'        Case 0: txtTagRules.Text = "'NoHello' even if it's not in $Supports statement." & vbNewLine & vbNewLine & "Tests for minimum DC++ version V:______"
        ' DC
'        Case 1: txtTagRules.Text = "Skips standard DC++ O:# tests." & vbNewLine & vbNewLine & "O:# is used for free/open slots."
        ' DCGUI
'        Case 2: txtTagRules.Text = "* in it's slot param (S:*) means unlimited slots." & vbNewLine & vbNewLine & "* in it's limiter param (L:*) means unlimited bandwidth." & vbNewLine & vbNewLine & "Reports bandwidth limit on a per slot basis, not total."
        ' DC:Pro
'        Case 5: txtTagRules.Text = "Uses F:#Down/#Up to report bandwidth limiting."
        ' SdDC++
'        Case 8: txtTagRules.Text = "slot param has the format S:#/#"
        ' Chat (Gadgets Flash Add-on)
'        Case 9: txtTagRules.Text = "If you are using this option then you can figure it out for yourself."
'        Case Else: txtTagRules.Text = "None"
'    End Select
' ----------------------- NEW MOD INTERFACE LANGUAGE END ----------------------
End Sub

Private Sub lstTagsDef_LostFocus()
1:    lstTagsDef.ListIndex = -1
' ------------------------ NEW MOD INTERFACE LANGUAGE ------------------------
3:    txtTagRules.Text = Replace(m_arrTagRules(10), "%[LF]", vbNewLine)
'    txtTagRules.Text = "Select a Default Tag to see if it has any special processing rules."
' ----------------------- NEW MOD INTERFACE LANGUAGE END ----------------------
End Sub

'------------------------------------------------------------------------------
'Detect IP events
'------------------------------------------------------------------------------

Private Sub m_objDetectIP_OnDownloaded(strHeader As String, strData As String)
1:    Dim strTemp As String
    
3:    On Error GoTo Err
    
    'Skip vbNewLine in front
6:    strTemp = MidB$(strData, 5)
    
8:    If MsgBox(Replace(g_colMessages.Item("msgYourIP"), "%[IP]", strTemp), vbYesNo Or vbQuestion, g_colMessages.Item("msgDetectIP")) = vbYes Then
9:        Clipboard.Clear
10:        Clipboard.SetText strTemp
11:    End If
        
13:    Exit Sub
    
15:
Err:
16:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.m_objDetectIP_OnDownloaded(""" & strHeader & """, """ & strData & """)"
End Sub

Private Sub m_objDetectIP_OnError(ByVal lngNumber As Long, strDescription As String)
1:    If MsgBox(Replace(g_colMessages.Item("msgIPError"), "%[IP]", wskListen(0).LocalIP), vbYesNo, g_colMessages.Item("msgDetectIP")) = vbYes Then
2:        Clipboard.Clear
3:        Clipboard.SetText wskListen(0).LocalIP
4:    End If
End Sub

Private Sub mnuCodeRTB_Click(Index As Integer)
1:    On Error GoTo Err
2:    Dim i As Integer
3:    Dim Modal As Byte
     
5:    If Index = 5 Then
6:        frmHelp.Show Modal, Me
7:        Exit Sub
8:    End If
   
10:   For i = 1 To UBound(sciMain)
11:      Select Case Index
            Case 0 'View WhiteSpace
13:               If i = 1 Then mnuCodeRTB(0).Checked = Not mnuCodeRTB(0).Checked
14:               sciMain(i).ViewWhiteSpace = mnuCodeRTB(0).Checked
            Case 1 'Line Number
16:               If i = 1 Then mnuCodeRTB(1).Checked = Not mnuCodeRTB(1).Checked
17:               sciMain(i).LineNumbers = mnuCodeRTB(1).Checked
            Case 7 'Word Wrap
19:               If i = 1 Then mnuCodeRTB(7).Checked = Not mnuCodeRTB(7).Checked
20:               sciMain(i).WordWrap = mnuCodeRTB(7).Checked
            Case 8 'ReadOnly
22:               If i = 1 Then mnuCodeRTB(8).Checked = Not mnuCodeRTB(8).Checked
23:               sciMain(i).ReadOnly = mnuCodeRTB(8).Checked
         End Select
25:   Next

27:   If Index <> 5 Then SCI_Focus
    
29:   Exit Sub
30:
Err:
32:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuCodeRTB_Click(" & Index & ")"
End Sub

'------------------------------------------------------------------------------
'Menu events
'------------------------------------------------------------------------------

Private Sub mnuPlan_Click(Index As Integer)
1:    Dim Name As String
2:    Dim When As String
3:    Dim when1 As String
4:    Dim when2 As String
5:    Dim CMD As String
6:    Dim tmp As String
7:    Dim X As Integer
   
9:   On Error Resume Next
   
    Select Case Index
        Case 0 'Add
11:            Name = InputBox("Enter a new item name", "Plan")
12:            If LenB(Name) Then
13:                When = InputBox("Enter when action should start" _
                & vbNewLine & "daily:[hh]:[mm] - hourly:[mm]", "Plan")
15:                If LenB(When) Then
16:                    CMD = InputBox("Enter command to execute", "Plan")
17:                    If LenB(CMD) Then
                               ' check against wrong timing syntax
19:                            If InStr(When, "hourly:") Then
20:                                When = g_objFunctions.AfterLast(When, ":")
21:                                If Val(When) < 1 Or Val(When) > 60 Then When = "00"
22:                                When = "hourly:" & When
23:                            End If
24:                            If InStr(When, "daily:") Then
25:                                when2 = g_objFunctions.AfterLast(When, ":")
26:                                when1 = g_objFunctions.BeforeLast(When, ":")
27:                                If Val(when2) < 1 Or Val(when2) > 60 Then when2 = "00"
28:                                If Val(when1) < 1 Or Val(when1) > 23 Then when1 = "00"
29:                                When = "daily:" & when1 & ":" & when2
30:                            End If

32:                        tmp = g_colScheduler.Plan("add " & Name & " " & When & " " & CMD)
36:                    End If
37:                End If
38:            End If
        Case 1 'Remove
39:            For X = 0 To lstPlan.ListCount - 1
40:                If lstPlan.Selected(X) Then
41:                    tmp = g_objFunctions.BeforeFirst(lstPlan.List(X), vbTab)
42:                    tmp = g_colScheduler.Plan("del " & tmp)
43:                End If
44:            Next
48:    End Select
End Sub

Private Sub mnuLocked_Click(Index As Integer)
  Dim lvwItem     As ListItem
  Dim lvwItems    As ListItems
  Dim strTxt      As String
  Set lvwItem = lvwBans.SelectedItem
  Set lvwItems = lvwBans.ListItems
  
7:  On Error GoTo Err

9:  If lvwItems.count = 0 Then Exit Sub
    'Check if selected
11:    If lvwItem.Selected Then
12:        Clipboard.Clear
           Select Case Index
               Case 0 'Copy User Name
15:                Clipboard.SetText (lvwItem.Text)
               Case 2 'Copy All
17:                strTxt = "PT DC Hub " & vbVersion & " - Ban Name" & vbNewLine & vbNewLine & _
                         "User Name: " & lvwItem.Text & vbNewLine & _
                         "Perm: " & lvwItem.SubItems(1) & vbNewLine & _
                         "Banned By: " & lvwItem.SubItems(2) & vbNewLine & _
                         "Reference Date: " & lvwItem.SubItems(3) & vbNewLine & _
                         "Reason: " & lblHolder(50).Caption
22:                Clipboard.SetText (strTxt)
          End Select
24:    End If
    
26:  Exit Sub
27:
Err:
29:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuLocked_Click(" & Index & ")"
End Sub

Private Sub mnuPermIPBan_Click(Index As Integer)
1:    Dim strIP       As String
2:    Dim lngMinutes  As Long
3:    Dim varLoop     As Variant
4:    Dim colBans     As Collection
5:    Dim lvwItems    As ListItems

7:    On Error GoTo Err
    
      Select Case Index
        Case 0 'Add
               frmBanPerm.Show vbModal, Me
16:            GoTo RefreshN
        Case 1 'Remove
18:            If ObjPtr(lvwPermIPBan.SelectedItem) Then
19:                strIP = lvwPermIPBan.SelectedItem.Key
20:                lvwPermIPBan.ListItems.Remove strIP
21:                g_objIPBans.Remove strIP
'22:            Else
'23:                strIP = InputBox(g_colMessages.Item("msgEnterRemIP"), g_colMessages.Item("msgRemoveIP"))
'24:                If LenB(strIP) Then g_objIPBans.Remove strIP
25:            End If
        Case 2 'Clear
27:            If MsgBoxCenter(Me, g_colMessages.Item("msgClearPermIPs"), vbYesNo Or vbExclamation, g_colMessages.Item("msgConfirmClear")) = vbYes Then _
                g_objIPBans.ClearPerm
        Case 4 'Refresh list extract
RefreshN:
31:            Set colBans = g_objIPBans.PermItems
32:            Set lvwItems = lvwPermIPBan.ListItems
        
            'Clear out items first
35:            lvwItems.Clear
            
            Select Case m_lngBanFilter
                Case 0 'No filter
39:                    For Each varLoop In colBans
40:                        lvwItems.Add , varLoop, varLoop
41:                    Next
                Case 1 'End in
43:                    strIP = txtBanFilter.Text
44:                    lngMinutes = LenB(strIP)
                
46:                    For Each varLoop In colBans
47:                        If RightB$(varLoop, lngMinutes) = strIP Then
48:                              lvwItems.Add , varLoop, varLoop
49:                           End If
50:                    Next
                Case 2 'Contain
52:                    strIP = txtBanFilter.Text
53:                    For Each varLoop In colBans
54:                        If InStrB(1, varLoop, strIP) Then
55:                              lvwItems.Add , varLoop, varLoop
56:                           End If
57:                    Next
                Case 3 'Begin with
59:                    strIP = txtBanFilter.Text
60:                    lngMinutes = LenB(strIP)
                
62:                    For Each varLoop In colBans
63:                        If LeftB$(varLoop, lngMinutes) = strIP Then
64:                              lvwItems.Add , varLoop, varLoop
65:                           End If
66:                    Next
67:            End Select
        Case 5 'Clear list extract
69:            lvwPermIPBan.ListItems.Clear
70:    End Select
    
72:    Exit Sub
    
74:
Err:
76:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuPermIPBan_Click(" & Index & ")"
End Sub

Private Sub mnuPlugIn_Click(Index As Integer)
1:   On Error GoTo Err
   
3:   Set mPlgObj = CreateObject(mnuPlugIn(Index).Tag)
    
5:   If mPlgObj.loadplug <> 1 Then
6:        MsgBoxCenter Me, "There was an error while loading the plugin.", vbCritical
7:        Exit Sub
8:   Else
9:        mPlgObj.LoadPlugin frmHub
10:  End If
    
12:  Exit Sub
13:
Err:
14:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuPlugIn_Click(" & Index & ") - " & mnuPlugIn(Index).Tag
End Sub

Public Sub mnuScripts_Click(Index As Integer)
1:  On Error GoTo Err
3:    Dim i As Integer

4:    For i = 1 To picSciMain.count - 1
5:            If picSciMain(i).Visible Then Exit For
6:    Next

8:    With frmScript
9:        Select Case Index
                Case 0 'Save/Reset
11:                 Call .SReset(i, True, True)
                Case 2 'Stop
13:                 .SStop i
                Case 3 'Stop All
15:                 .SStop -2
                Case 5 'Reolad Checkeds
17:                 Call .SReset(-2, True, True)
                Case 6 'Reolad Dir
18:                 .XmlBooleanSave
19:                 .SLoadDir
20:                 .XmlBooleanLoad
21:                 .SReset -2, False, False
                Case 8 'Properties
23:                 .SProperties CStr(i & "s"), lvwScripts.ListItems(i).Text, 0
          End Select
25:  End With

27:  If Index <> 8 Then SCI_Focus

29:  Exit Sub
30:
Err:
31:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuScripts_Click(" & Index & ")"
End Sub

#If Status Then
    Private Sub lstStatus_DblClick(Index As Integer)
1:       On Error GoTo Err
2:       With frmMulti
3:           .cmdCancel.Visible = False
4:           .Label1.Visible = False
5:           .Caption = "Copy to Clipboard Text"
6:           .txtStr.Top = 120
7:           .cmdOK.Top = 520
8:           .Height = 1350
9:           .txtStr.Text = CStr(lstStatus(Index))
10:          .Show vbModal, Me
11:       End With
12:      Exit Sub
13:
Err:
15:      HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.lstStatus_DblClick(" & Index & ")"
    End Sub
    Private Sub lstStatus_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:        If Button = 2 Then PopupMenu mnuPopUp(6)
    End Sub
    Private Sub mnuStatus_Click(Index As Integer)
1:        Dim lngLoop     As Long
2:        Dim lngUB       As Long
3:        Dim strCopy     As String
4:        Dim i           As Integer
    
6:        On Error GoTo Err
          
8:           If picStatus(0).Visible Then
9:             i = 0
10:          ElseIf picStatus(1).Visible Then
11:             i = 1
12:          ElseIf picStatus(2).Visible Then
13:             i = 2
14:          End If
          
16:          Select Case Index
                Case 0 'Copy
17:                lngUB = lstStatus(i).ListCount - 1
            
                  'Clear the clipboard before we start
20:                Clipboard.Clear
                
                  'Loop through and find all selected items
23:                For lngLoop = 0 To lngUB
24:                    If lstStatus(i).Selected(lngLoop) Then strCopy = strCopy & lstStatus(i).List(lngLoop) & vbNewLine
25:                Next
                
                  'Set the clipboard to all selected text
28:                Clipboard.SetText strCopy
            
               Case 1 'Clear
31:                g_objStatus.MClear i
32:        End Select
    
34:        Exit Sub
    
36:
Err:
38:        HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuStatus_Click(" & Index & ")"
    End Sub
    Private Sub mnuUsers_Click(Index As Integer)
1:        Dim lvwItem     As ListItem
2:        Dim lvwItems    As ListItems
3:        Dim strOne      As String
4:        Dim strTwo      As String
5:        Dim intOne      As Integer
6:        Dim lngOne      As Long

8:        On Error GoTo Err
        
        'Get selected item
11:        Set lvwItem = lvwUsers.SelectedItem

        'Get listitem collection
14:        Set lvwItems = lvwUsers.ListItems

        Select Case Index
            'Send data (selected)
            Case 0
                'Get message
18:                strOne = InputBox(g_colMessages.Item("msgEnterDataToSel"), g_colMessages.Item("msgSendToSel"))

20:                If LenB(strOne) Then
21:                    For Each lvwItem In lvwItems
                        'Send if selected
23:                        If lvwItem.Selected Then _
                            g_colUsers.ItemByName(lvwItem.Text).SendData strOne
25:                    Next
26:                End If
            'Send data (all)
            Case 1
                'Get message
29:                strOne = InputBox(g_colMessages.Item("msgEnterDataToAll"), g_colMessages.Item("msgSendToAll"))

31:                If LenB(strOne) Then g_colUsers.SendToAll strOne
            'Disconnect
            Case 2
33:                If ObjPtr(lvwItem) = 0 Then Exit Sub

35:                wskLoop_Close CInt(lvwItem.SubItems(2))
            'Kick
            Case 3
37:                If ObjPtr(lvwItem) = 0 Then Exit Sub

                'Get ban length
40:                strTwo = InputBox(g_colMessages.Item("msgEnterLength"), g_colMessages.Item("msgKickSel"), "0")
41:                If LenB(strTwo) Then lngOne = CLng(Val(strTwo)) Else Exit Sub

                'Get reason
44:                strOne = InputBox(g_colMessages.Item("msgKickReason"), g_colMessages.Item("msgKick"), g_colUsers.ItemByWinsockIndex(CInt(lvwItem.SubItems(2))).GetCoreMsgStr("KickedBecause"))

46:                If LenB(strOne) Then
47:                    Set m_objLoopUser = g_colUsers.ItemByWinsockIndex(CInt(lvwItem.SubItems(2)))
48:                    m_objLoopUser.SendChat g_objSettings.BotName, strOne
49:                    DoEvents
50:                    m_objLoopUser.Kick lngOne

52:                    Set m_objLoopUser = Nothing
53:                End If
            'Redirect
            Case 4
55:                If ObjPtr(lvwItem) = 0 Then Exit Sub

                'Get address
58:                strTwo = InputBox(g_colMessages.Item("msgEnterRedirAddress"), g_colMessages.Item("msgRedirSel"), g_objSettings.RedirectIP)
59:                If LenB(strTwo) = 0 Then Exit Sub

                'Get reason
62:                strOne = InputBox(g_colMessages.Item("msgRedirReason"), g_colMessages.Item("msgRedir"), g_colUsers.ItemByWinsockIndex(CInt(lvwItem.SubItems(2))).GetCoreMsgStr("RedirectedBecause"))

64:                If LenB(strOne) Then
65:                    Set m_objLoopUser = g_colUsers.ItemByWinsockIndex(CInt(lvwItem.SubItems(2)))

67:                    m_objLoopUser.SendChat g_objSettings.BotName, strOne
68:                    m_objLoopUser.Redirect strTwo

70:                    Set m_objLoopUser = Nothing
71:                End If
            'Ban
            Case 5
73:                If ObjPtr(lvwItem) = 0 Then Exit Sub

                'Get message (if we're sending one)
76:                strOne = InputBox(g_colMessages.Item("msgEnterBanReason"), g_colMessages.Item("msgBan"), g_colUsers.ItemByWinsockIndex(CInt(lvwItem.SubItems(2))).GetCoreMsgStr("BannedBecause"))

78:                If LenB(strOne) Then
79:                    Set m_objLoopUser = g_colUsers.ItemByWinsockIndex(CInt(lvwItem.SubItems(2)))

81:                    m_objLoopUser.SendChat g_objSettings.BotName, strOne
82:                    DoEvents
83:                    m_objLoopUser.Ban

85:                    Set m_objLoopUser = Nothing
86:                End If
            '(De)mute
            Case 6
88:                If ObjPtr(lvwItem) = 0 Then Exit Sub

                'Swap mute status
91:                Set m_objLoopUser = g_colUsers.ItemByWinsockIndex(CInt(lvwItem.SubItems(2)))
92:                m_objLoopUser.Mute = Not m_objLoopUser.Mute
93:                Set m_objLoopUser = Nothing
            'Properties (selected)
            Case 7
95:                For Each lvwItem In lvwItems
                    'Check if selected
97:                    If lvwItem.Selected Then
                        'Get object / settings
99:                        Set m_objLoopUser = g_colUsers.ItemByWinsockIndex(CInt(lvwItem.SubItems(2)))

                        'Create property string
102:                        strOne = "Name : " & m_objLoopUser.sName & vbNewLine & _
                                 "Winsock Index : " & m_objLoopUser.iWinsockIndex & vbNewLine & _
                                 "IP : " & m_objLoopUser.IP & vbNewLine & _
                                 "Connected Since : " & m_objLoopUser.ConnectedSince & vbNewLine & _
                                 "Class : " & m_objLoopUser.Class & vbNewLine & _
                                 "Language : " & m_objLoopUser.sLanguageID & vbNewLine & _
                                 "Version : " & m_objLoopUser.iVersion & vbNewLine & _
                                 "Share : " & g_objFunctions.ShareSize(m_objLoopUser.iBytesShared) & vbNewLine & _
                                 "MyINFO : " & m_objLoopUser.sMyInfoString & vbNewLine & _
                                 "Supports : " & m_objLoopUser.Supports '& vbTwoLine

                        'Append to collection
114:                        strTwo = strTwo & strOne
115:                    End If
116:                Next

118:                Set m_objLoopUser = Nothing

                    frmUserInfo.txtInfo = strTwo
                    frmUserInfo.Show vbModal, Me
                    
122:        End Select

124:        RefreshGUI

126:        Exit Sub

128:
Err:
129:        HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuUsers_Click(" & Index & ")"
    End Sub
    Private Sub txtStForm_Change()
1:       If txtStForm.Text = "" Then txtStForm.Text = g_objSettings.BotName
    End Sub
    Private Sub txtStSend_KeyPress(KeyAscii As Integer)
1:      If KeyAscii = 13 Then _
           KeyAscii = 0: Call cmdStSend_Click
    End Sub
    Private Sub sldStatus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:       picStInfo.Visible = True
    End Sub
    Private Sub sldStatus_Change()
1:      If optStSend(1).Value And sldStatus.Value = 6 Then _
             cmdStSend.Enabled = False _
        Else cmdStSend.Enabled = True
4:      Dim i As Integer
5:      For i = 2 To 7
6:            lblStatus(i).FontUnderline = False
7:      Next
8:      lblStatus(sldStatus.Value + 1).FontUnderline = True
    End Sub
    Private Sub cmdStSend_Click()
        Dim strMsg   As String
        Dim srtForm  As String
     
        On Error GoTo Err
     
6:         strMsg = txtStSend.Text
7:         srtForm = txtStForm.Text
8:         txtStSend.Text = ""

10:        If optStSend(0).Value Then 'Send Chat
              Select Case sldStatus.Value
                Case 1 'Send Chat To All
13:                    g_colUsers.SendChatToAll srtForm, strMsg
14:                    AddLog "Send Chat To All :" & "<" & srtForm & "> " & strMsg, 6
                Case 2 'Send Chat To Op
16:                    g_colUsers.SendChatToOps srtForm, strMsg
17:                    AddLog "Send Chat To Op: " & "<" & srtForm & "> " & strMsg, 6
                Case 3 'Send Chat To UnRegistered
18:                    g_colUsers.SendChatToUnReg srtForm, strMsg
19:                    AddLog "Send Chat To UnRegistered: " & "<" & srtForm & "> " & strMsg, 6
                Case 4 'Send PM To All
21:                    g_colUsers.SendPrivateToAll srtForm, strMsg
22:                    AddLog "Send PM To All: " & "< " & srtForm & " > " & strMsg, 6
                Case 5 'Send PM To Op
24:                    g_colUsers.SendPrivateToOps srtForm, strMsg
25:                    AddLog "Send PM To Op: " & "<" & srtForm & ">" & strMsg, 6
                Case 6 'Send PM To UnRegistered
27:                    g_colUsers.SendPrivateToUnReg srtForm, strMsg
28:                    AddLog "Send PM To UnRegistered: " & "<" & srtForm & ">" & strMsg, 6
              End Select
30:           SEvent_MassMessage strMsg
              #If Status Then
32:                g_objStatus.MAdd "<" & srtForm & "> " & strMsg
              #End If
34:        ElseIf optStSend(1) Then 'Send Data
              Select Case sldStatus.Value
                Case 1 'Send Data To All
37:                    g_colUsers.SendToAll strMsg
38:                    AddLog "Send Data To All: " & strMsg, 6
                Case 2 'Send Data To Op
40:                    g_colUsers.SendToOps strMsg
41:                    AddLog "Send Data To Op: " & strMsg, 6
                Case 3 'Send Data To UnRegistered
43:                    g_colUsers.SendToUnReg strMsg
44:                    AddLog "Send Data To UnRegistered: " & strMsg, 6
                Case 4 'Send Data No Away Mode
46:                    g_colUsers.SendToNA strMsg
47:                    AddLog "Send Data No Away Mode: " & strMsg, 6
                Case 5 'Send Data Non Quick List Clients
49:                    g_colUsers.SendToNQ strMsg
50:                    AddLog "Send Data Non Quick List Clients: " & strMsg, 6
                Case 6
                        '
                        '
              End Select
              #If Status Then
56:                g_objStatus.MAdd strMsg
              #End If
58:        End If
59:     Exit Sub
60:
Err:
62:        HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.cmdStSend_Click()"
    End Sub
    Private Sub cmdStSend_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:       picStInfo.Visible = False
    End Sub
    Private Sub lblStatus_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:       picStInfo.Visible = False
    End Sub
    Private Sub optStSend_Click(Index As Integer)
1:       On Error GoTo Err
         Select Case Index
             Case 0
4:               lblStatus(1).Caption = "------------- Send Chat -------------"
5:               lblStatus(2).Caption = "1 = Send Chat To All"
6:               lblStatus(3).Caption = "2 = Send Chat To Op"
7:               lblStatus(4).Caption = "3 = Send Chat To UnRegistered"
8:               lblStatus(5).Caption = "4 = Send PM To All"
9:               lblStatus(6).Caption = "5 = Send PM To Op"
10:              lblStatus(7).Caption = "6 = Send PM To UnRegistered"
11:              txtStForm.Enabled = True
12:              txtStForm.BackColor = &H80000005
             Case 1
14:              lblStatus(1).Caption = "------------- Send Data -------------"
15:              lblStatus(2).Caption = "1 = Send Data To All"
16:              lblStatus(3).Caption = "2 = Send Data To Op"
17:              lblStatus(4).Caption = "3 = Send Data To UnRegistered"
18:              lblStatus(5).Caption = "4 = Send Data No Away Mode"
19:              lblStatus(6).Caption = "5 = Send Data Non QuickListClients"
20:              lblStatus(7).Caption = "6 = ----------------"
21:              txtStForm.Enabled = False
22:              txtStForm.BackColor = &H8000000F
         End Select
24:      sldStatus.Value = 1
25:      Exit Sub
26:
Err:
28:      HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.optStSend_Click(" & Index & ")"
    End Sub
    Private Sub optStSend_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:       picStInfo.Visible = False
    End Sub
#End If

Private Sub picLog_Click(Index As Integer)
1:   'If Index = 2 Then
2:   '   On Error Resume Next
3:   '   ShellExecute Me.hWnd, "open", "https://www.paypal.com/xclick/business=PTDCH.hubsoftware%40gmail.com&item_name=PTDCH+Hub+Software&currency_code=EUR", "", "", 3   'SW_SHOWMAXIMIZED
4:   'End If
End Sub

Private Sub mnuTags_Click(Index As Integer)
1:    Dim lngLoop     As Long
2:    Dim lngUB       As Long
3:    Dim strTag      As String
4:    Dim objTag      As clsTag
    
6:    On Error GoTo Err
    
    Select Case Index
        Case 0 'Add
8:            strTag = InputBox(g_colMessages.Item("msgEnterTag"), g_colMessages.Item("msgAddTag"))
            
10:            If LenB(strTag) Then
                'Make sure the tag isn't already in the list
12:                On Error Resume Next
13:                m_colTags.Item strTag
                
15:                If Err.Number Then
16:                    On Error GoTo Err
                
                    'Add to list
19:                    lstTagsEx.AddItem strTag
                    
                    'Add to collection
22:                    Set objTag = New clsTag
                    
24:                    objTag.Name = strTag
                    'If a user re-adds a default Tag give it the right default ID
                    Select Case objTag.Name
                        Case "++": objTag.ID = 1
                        Case "DC": objTag.ID = 2
                        Case "DCGUI": objTag.ID = 3
                        Case "oDC": objTag.ID = 4
                        Case "QuickDC": objTag.ID = 5
                        Case "DC:Pro": objTag.ID = 6
                        Case "SDC": objTag.ID = 7
                        Case "StrgDC++": objTag.ID = 10
                        Case "SdDC++": objTag.ID = 8
                        Case "Z++": objTag.ID = 11
                        Case "Chat": objTag.ID = 9
                        Case Else: objTag.ID = -1
26:                    End Select
                    
28:                    m_colTags.Add objTag, strTag
                    
30:                    Set objTag = Nothing
31:                Else
32:                    MsgBoxCenter Me, strTag & g_colMessages.Item("msgAlreadyAdded"), vbInformation, "PTDCH"
33:                End If
34:            End If
        Case 1 'Remove
            'Make sure some tags are selected before looping through
36:            If lstTagsEx.SelCount Then
37:                lngUB = lstTagsEx.ListCount - 1
            
39:                For lngLoop = 0 To lngUB
40:                    If lstTagsEx.Selected(lngLoop) Then
                        'Remove from collection/list
42:                        m_colTags.Remove lstTagsEx.List(lngLoop)
43:                        lstTagsEx.RemoveItem lngLoop
                        
45:                        Exit For
46:                    End If
47:                Next
48:            End If
49:    End Select
    
51:    Exit Sub
    
53:
Err:
54:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuTags_Click(" & Index & ")"
End Sub

Private Sub mnuTempIPBan_Click(Index As Integer)
1:    Dim strIP       As String
2:    Dim lngMinutes  As Long
3:    Dim lvwItems    As Variant
4:    Dim X           As Variant
5:    Dim colBans     As Collection
6:    Dim objTB       As clsTempBan
    
8:    On Error GoTo Err
    
10:    Select Case Index
        Case 0 'Add
12:            With frmBanTemp
13:               .Caption = g_colMessages.Item("msgBanTempIP")
14:               .Labels(0).Caption = g_colMessages.Item("msgEnterBanLength")
15:               .Labels(1).Caption = g_colMessages.Item("msgEnterBanLength")
16:               .Show vbModal, Me
17:            End With
               GoTo RefreshN
        Case 1 'Remove
20:            If ObjPtr(lvwTempIPBan.SelectedItem) Then
21:                strIP = lvwTempIPBan.SelectedItem.Text
22:                lvwTempIPBan.ListItems.Remove CStr(strIP & "s")
24:                g_objIPBans.Remove strIP
25:            Else
'26:                strIP = InputBox(g_colMessages.Item("msgEnterRemIP"), g_colMessages.Item("msgRemoveIP"))
'27:                If LenB(strIP) Then g_objIPBans.Remove strIP
28:            End If
        Case 2 'Clear
30:            If MsgBoxCenter(Me, g_colMessages.Item("msgClearTempIPs"), vbYesNo Or vbExclamation, g_colMessages.Item("msgConfirmClear")) = vbYes Then _
                g_objIPBans.ClearTemp
        Case 4 'Refresh list extract
RefreshN:
33:            Set colBans = g_objIPBans.TempItems
34:            Set lvwItems = lvwTempIPBan.ListItems
            
            'Clear out items first
37:            lvwItems.Clear
            'Note: this char "." is not valide key.. use replace for ":", because cause error in item key..
            Select Case m_lngBanFilter
                Case 0 'No filter
41:                    For Each objTB In colBans
42:                        Set X = lvwItems.Add(, objTB.IP & "s", objTB.IP)
                           X.SubItems(1) = objTB.ExpDate
44:                    Next
                Case 1 'End in
46:                    strIP = txtBanFilter.Text
47:                    lngMinutes = LenB(strIP)
48:                    For Each objTB In colBans
49:                        If RightB$(objTB.IP, lngMinutes) = strIP Then
50:                           Set X = lvwItems.Add(, objTB.IP & "s", objTB.IP)
51:                           X.SubItems(1) = objTB.ExpDate
52:                        End If
53:                    Next
                Case 2 'Contain
55:                    strIP = txtBanFilter.Text
56:                    For Each objTB In colBans
57:                        If InStrB(1, objTB.IP, strIP) Then
58:                           Set X = lvwItems.Add(, objTB.IP & "s", objTB.IP)
59:                           X.SubItems(1) = objTB.ExpDate
60:                        End If
61:                    Next
                Case 3 'Begin with
63:                    strIP = txtBanFilter.Text
64:                    lngMinutes = LenB(strIP)
65:                    For Each objTB In colBans
66:                        If LeftB$(objTB.IP, lngMinutes) = strIP Then
67:                           Set X = lvwItems.Add(, objTB.IP & "s", objTB.IP)
68:                           X.SubItems(1) = objTB.ExpDate
69:                        End If
70:                    Next
71:            End Select
        Case 5 'Clear list extract
73:            lvwTempIPBan.ListItems.Clear
74:    End Select
    
76:    Exit Sub
    
78:
Err:
80:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuTempIPBan_Click(" & Index & ")"
End Sub

Private Sub mnuRegistered_Click(Index As Integer)
  Dim lvwItem     As ListItem
  Dim lvwItems    As ListItems
  Dim strTxt      As String
  
  Set lvwItem = lvwRegistered.SelectedItem
  Set lvwItems = lvwRegistered.ListItems
       
8:  On Error GoTo Err
9:  If lvwItems.count = 0 Then Exit Sub
    'Check if selected
11:    If lvwItem.Selected Then
12:        Clipboard.Clear
           Select Case Index
               Case 0 'Copy User Name
15:                Clipboard.SetText (lvwItem.Text)
               Case 1 'Copy Password
17:                Clipboard.SetText (lvwItem.SubItems(1))
               Case 2 'Copy Last IP
19:                Clipboard.SetText (lvwItem.SubItems(7))
               Case 4 'Copy All
21:                strTxt = "PT DC Hub " & vbVersion & " - Reg Name" & vbNewLine & vbNewLine & _
                            "User Name: " & lvwItem.Text & vbNewLine & _
                            "Password: " & lvwItem.SubItems(1) & vbNewLine & _
                            "Class: " & lvwItem.SubItems(2) & "=" & lvwItem.SubItems(3) & vbNewLine & _
                            "Reged By: " & lvwItem.SubItems(4) & vbNewLine & _
                            "Reg Date: " & lvwItem.SubItems(5) & vbNewLine & _
                            "Last Login: " & lvwItem.SubItems(6) & vbNewLine & _
                            "Last IP: " & lvwItem.SubItems(7)
29:                Clipboard.SetText (strTxt)
           End Select
31:    End If
    
33:   Exit Sub
34:
Err:
36:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuRegistered_Click(" & Index & ")"
End Sub

Private Sub mnuTray_Click(Index As Integer)
1:    On Error GoTo Err
    
      Select Case Index
        Case 0 'Show
3:            WindowState = vbNormal
4:            Show
        Case 1 'Hide
5:            WindowState = vbMinimized
      End Select
      
8:    Exit Sub
9:
Err:
11:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.mnuTray_Click(" & Index & ")"
End Sub

'------------------------------------------------------------------------------
'Changed settings events
'------------------------------------------------------------------------------
Private Sub cmbData_Click(Index As Integer)
1:    On Error GoTo Err

    Select Case Index
        Case 1, 2 'Share sizes
3:            CallByName g_objSettings, cmbData(Index).Tag, VbLet, CByte(cmbData(Index).ListIndex)
        Case Else
4:            CallByName g_objSettings, cmbData(Index).Tag, VbLet, cmbData(Index).Text
5:    End Select
    
7:    Exit Sub
    
9:
Err:
10:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.cmbData_Click(" & Index & ")"
End Sub

' ------------------------ NEW INTERFACE LANGUAGE ------------------------
Private Sub cmbInterface_Click()
    
2:     If g_objSettings.Interface = cmbInterface.Text Then Exit Sub
3:     g_objSettings.Interface = cmbInterface.Text
4:     cmdButton_Click 3    'Save Settings
    
6:     Dim objXML          As clsXMLParser
7:     Dim objNode         As clsXMLNode
8:     Dim objSubNode      As clsXMLNode
9:     Dim colNodes        As Collection
10:    Dim colSubNodes     As Collection
11:    Dim colAttributes   As Collection
    
13:    Dim strTemp         As String
14:    Dim X               As Integer
15:    On Error GoTo Err

17:    Set objXML = New clsXMLParser
       
19:    Call ClearTranslations
       
       'Set new Interface Language
22:    strTemp = G_APPPATH & "\Languages\" & cmbInterface.Text & ".xml"
    
24:    If g_objFileAccess.FileExists(strTemp) Then
        
26:        g_objSettings.Interface = cmbInterface.Text
    
28:        objXML.Data = g_objFileAccess.ReadFile(strTemp)
29:        objXML.Parse
    
31:        Set colNodes = objXML.Nodes(1).Nodes

33:        On Error Resume Next 'Just in case...

35:        For Each objNode In colNodes
36:            Set colSubNodes = objNode.Nodes
    
            Select Case objNode.Name
                Case "DynamicCaptions"
40:                    For Each objSubNode In colSubNodes
41:                        m_arrDynaCap(X) = objSubNode.Value
42:                        X = X + 1
43:                    Next
                Case "Captions"
45:                    For Each objSubNode In colSubNodes
46:                        TranslateCtrlCaption objSubNode.Name, objSubNode.Value
47:                    Next
                Case "Texts"
49:                    For Each objSubNode In colSubNodes
50:                        TranslateTexts objSubNode.Name, objSubNode.Value
51:                    Next
                Case "TabSCaption"
53:                    For Each objSubNode In colSubNodes
54:                        TranslateTabSCaption objSubNode.Name, objSubNode.Value
55:                    Next
                Case "ToolTips"
57:                    For Each objSubNode In colSubNodes
58:                        TranslateCtrlToolTip objSubNode.Name, objSubNode.Value
59:                    Next
                Case "ListView"
61:                    For Each objSubNode In colSubNodes
62:                        TranslateListViewCaption objSubNode.Name, objSubNode.Value
63:                    Next
                Case "Captions"
65:                    For Each objSubNode In colSubNodes
66:                        TranslateCtrlCaption objSubNode.Name, objSubNode.Value
67:                    Next
                Case "TagsHelp"
69:                    For Each objSubNode In colSubNodes
80:                        m_arrTagRules(objSubNode.Name) = objSubNode.Value
81:                    Next
                Case "ToolTips"
83:                    For Each objSubNode In colSubNodes
84:                        TranslateCtrlToolTip objSubNode.Name, objSubNode.Value
85:                    Next
                Case "HubStringDef"
87:                    For Each objSubNode In colSubNodes
88:                        g_colMessages.Item(objSubNode.Name) = objSubNode.Value
89:                    Next
                Case "ToolBar"
91:                    For Each objSubNode In colSubNodes
92:                        TranslateToolBar objSubNode.Name, objSubNode.Value
93:                    Next
94:            End Select
95:        Next
        
97:        On Error GoTo Err
        
99:       objXML.Clear
        
100:       Set objSubNode = Nothing
101:       Set objNode = Nothing
102:       Set colSubNodes = Nothing
103:       Set colNodes = Nothing

105:   End If

107:   txtTagRules.Text = Replace(m_arrTagRules(10), "%[LF]", vbNewLine)
    
109:   If m_blnServing = False Then
110:      cmdButton(1).Caption = m_arrDynaCap(0)
111:   Else
112:     cmdButton(1).Caption = m_arrDynaCap(1)
113:   End If
    
114: Exit Sub
    
116:
Err:
118:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.cmbInterface_Click()"
End Sub
' ----------------------- NEW INTERFACE LANGUAGE END ----------------------

Private Sub cmbRegistered_Click()

2:      On Error GoTo Err

        Select Case cmbRegistered.ListIndex
            Case 0 'All classes
6:            adoUsers.RecordSource = "SELECT UsrClass.UserName, UsrStatic.Pass, UsrClass.Class, ClassTypes.Name, UsrStatic.RegedBy, UsrStatic.RegDate, UsrDynamic.LastLogin, UsrDynamic.LastIP FROM UsrStatic INNER JOIN (ClassTypes INNER JOIN (UsrClass INNER JOIN UsrDynamic ON UsrClass.UserName = UsrDynamic.UserName) ON ClassTypes.ID = UsrClass.Class) ON (UsrStatic.UserName = UsrDynamic.UserName) AND (UsrStatic.UserName = UsrClass.UserName) ORDER BY UsrClass.UserName;"
            Case 1 'Non ops
8:            adoUsers.RecordSource = "SELECT UsrClass.UserName, UsrStatic.Pass, UsrClass.Class, ClassTypes.Name, UsrStatic.RegedBy, UsrStatic.RegDate, UsrDynamic.LastLogin, UsrDynamic.LastIP FROM UsrStatic INNER JOIN (ClassTypes INNER JOIN (UsrClass INNER JOIN UsrDynamic ON UsrClass.UserName = UsrDynamic.UserName) ON ClassTypes.ID = UsrClass.Class) ON (UsrStatic.UserName = UsrDynamic.UserName) AND (UsrStatic.UserName = UsrClass.UserName) Where ((([Class] > 1) And ([Class] < 6))) ORDER BY UsrClass.UserName;"
            Case 2 'Ops and above
10:           adoUsers.RecordSource = "SELECT UsrClass.UserName, UsrStatic.Pass, UsrClass.Class, ClassTypes.Name, UsrStatic.RegedBy, UsrStatic.RegDate, UsrDynamic.LastLogin, UsrDynamic.LastIP FROM UsrStatic INNER JOIN (ClassTypes INNER JOIN (UsrClass INNER JOIN UsrDynamic ON UsrClass.UserName = UsrDynamic.UserName) ON ClassTypes.ID = UsrClass.Class) ON (UsrStatic.UserName = UsrDynamic.UserName) AND (UsrStatic.UserName = UsrClass.UserName) Where (((UsrClass.Class) > 5)) ORDER BY UsrClass.UserName;"
            Case 3 'Admins and above
12:           adoUsers.RecordSource = "SELECT UsrClass.UserName, UsrStatic.Pass, UsrClass.Class, ClassTypes.Name, UsrStatic.RegedBy, UsrStatic.RegDate, UsrDynamic.LastLogin, UsrDynamic.LastIP FROM UsrStatic INNER JOIN (ClassTypes INNER JOIN (UsrClass INNER JOIN UsrDynamic ON UsrClass.UserName = UsrDynamic.UserName) ON ClassTypes.ID = UsrClass.Class) ON (UsrStatic.UserName = UsrDynamic.UserName) AND (UsrStatic.UserName = UsrClass.UserName) Where (((UsrClass.Class) > 9)) ORDER BY UsrClass.UserName;"
13:     End Select
   
14:     Call DBGetRegRecord
    
16:     Exit Sub
    
18:
Err:
20:     HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.cmbRegistered_Click()"
End Sub

Private Sub chkData_Click(Index As Integer)

2:    Dim objUser     As Object
3:    Dim strData     As String
    
5:    On Error GoTo Err

7:    CallByName g_objSettings, chkData(Index).Tag, VbLet, CBool(chkData(Index).Value)
    
9:    If Index = 52 Then
10:        If chkData(52).Value Then
11:            lblHolder(92).Visible = True
               Shape(3).Visible = True
12:            g_objSettings.DynUpdate = True
13:        Else
14:            lblHolder(92).Visible = False
               Shape(3).Visible = False
15:            g_objSettings.DynUpdate = False
16:        End If
17:    End If

19:        If Index = 53 Then
20:            If chkData(53).Value Then
21:                lstPlan.Visible = True
22:                g_objSettings.EnabledScheduler = True
23:            Else
24:                lstPlan.Visible = False
25:                g_objSettings.EnabledScheduler = False
26:            End If
27:        End If

31:    If Index = 41 Then
32:        If chkData(41).Value Then
33:            For Each objUser In g_colUsers
34:                strData = objUser.sMyInfoString
35:                objUser.sMyInfoFakeString = "$MyINFO $ALL " & g_objRegExps.CaptureSubStr(strData, GETNICK) & " $ $$$" & g_objRegExps.CaptureDbl(strData, GETSHARESIZE) & "$"
36:            Next
37:        End If
38:    End If
       
40:    If Index = 54 Then
41:       If chkData(54).Value Then
42:          Magnetic.AddWindow frmHub.hWnd
43:       Else
44:          Magnetic.RemoveWindow frmHub.hWnd
45:       End If
46:    End If
        
48:    If Index = 65 Then
49:       sldPriority.Value = g_objSettings.PriorityVal
50:       If chkData(65).Value = False Then
51:          sldPriority.Enabled = False
52:          lblPriority(0).Enabled = False
53:          lblPriority(1).Enabled = False
54:          lblPriority(2).Enabled = False
55:          lblPriority(3).Enabled = False
56:          SetPriorityLivel 1
57:       Else
58:          sldPriority.Enabled = True
59:          lblPriority(0).Enabled = True
60:          lblPriority(1).Enabled = True
61:          lblPriority(2).Enabled = True
62:          lblPriority(3).Enabled = True
63:          SetPriorityLivel (g_objSettings.PriorityVal)
64:       End If
65:    End If
      
67:    If Index = 66 Then
68:          If chkData(66).Value Then
                g_objSettings.blSkin = True
69:             Call cmbSkin_Click
70:             cmbSkin.Enabled = True
71:             chkData(67).Enabled = True
72:             cmdSkin(0).Enabled = True
73:             cmdSkin(1).Enabled = True
74:          Else
75:             If Not g_objSettings.lngSkin = 0 Then
76:                Dim i As Integer
77:                g_objSettings.blSkin = False
78:                On Error Resume Next
                  'Refresh all picture box .. very fast
80:                For i = 0 To picTab.count - 1: picTab(i).Cls: Next i
81:                For i = 0 To picSTab.count - 1: picSTab(i).Cls: Next i
82:                For i = 0 To picITab.count - 1: picITab(i).Cls: Next i
83:                For i = 0 To picTabAdv.count - 1: picTabAdv(i).Cls: Next i
84:                For i = 0 To picHelp.count - 1: picHelp(i).Cls: Next i
85:                For i = 0 To picBordTab.count - 1: picBordTab(i).Cls: Next i
86:                For i = 0 To picInfo.count - 1: picInfo(i).Cls: Next i

88:                Call Form_Paint
89:                Me.Refresh
90:             End If
91:             On Error GoTo Err
92:             cmbSkin.Enabled = False
93:             chkData(67).Enabled = False
94:             cmdSkin(0).Enabled = False
95:             cmdSkin(1).Enabled = False
96:          End If
97:    End If


100:  Exit Sub
101:
Err:
102:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.chkData_Click(" & Index & ")"
End Sub

Private Sub optBanFilter_Click(Index As Integer)
1:    m_lngBanFilter = Index
End Sub

Private Sub optJM_Click(Index As Integer)
1:    g_objSettings.SendJoinMsg = Index
End Sub

Private Sub optRedirect_Click(Index As Integer)
1:    On Error GoTo Err

    'Set previous option to false
    Select Case True
        Case g_objSettings.AutoRedirect: g_objSettings.AutoRedirect = False
        Case g_objSettings.AutoRedirectFull: g_objSettings.AutoRedirectFull = False
        Case g_objSettings.AutoRedirectFullNonReg: g_objSettings.AutoRedirectFullNonReg = False
        Case g_objSettings.AutoRedirectFullNonOps: g_objSettings.AutoRedirectFullNonOps = False
        Case g_objSettings.AutoRedirectNonReg: g_objSettings.AutoRedirectNonReg = False
4:    End Select
    
    'Set correct option to true
    Select Case Index
        Case 0: g_objSettings.AutoRedirect = True
        Case 1: g_objSettings.AutoRedirectNonReg = True
        Case 2: g_objSettings.AutoRedirectFull = True
        Case 3: g_objSettings.AutoRedirectFullNonReg = True
        Case 4: g_objSettings.AutoRedirectFullNonOps = True
7:    End Select
    
9:    Exit Sub
    
11:
Err:
12:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.optRedirect_Click(" & Index & ")"
End Sub

Private Sub sldPriority_Scroll()
'set Process Priority Class
2:    SetPriorityLivel (sldPriority.Value)
3:    g_objSettings.PriorityVal = Val(sldPriority.Value)
End Sub

Private Sub tabAdv_Click()

 On Error GoTo Err
 
   Dim i, i2 As Integer

   i2 = Val(tabAdv.SelectedItem.Index - 1)
   If picTabAdv(i2).Visible = True Then Exit Sub
      
   For i = 0 To picTabAdv.count - 1
     picTabAdv(i).Visible = False
   Next i
   
   i = Val(tabAdv.SelectedItem.Index - 1)
   picTabAdv(i).Refresh
   picTabAdv(i).Visible = True
   
Exit Sub
Err:
    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tbsMenu_Click()"
    Resume Next
End Sub

Private Sub tbsHelp_Click()
 On Error GoTo Err
 
   Dim i, i2 As Integer
   
   i2 = Val(tbsHelp.SelectedItem.Index - 1)
   If picHelp(i2).Visible = True Then Exit Sub
   
   For i = 0 To picHelp.count - 1
     picHelp(i).Visible = False
   Next i
   
   i = Val(tbsHelp.SelectedItem.Index - 1)
   
   picHelp(i).Visible = True

Exit Sub
Err:
    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tbsHelp_Click()"
    Resume Next
End Sub

Private Sub tbsInfo_Click()
 On Error GoTo Err
 
   Dim i, i2 As Integer
   
   i2 = Val(tbsInfo.SelectedItem.Index - 1)
   If picInfo(i2).Visible = True Then Exit Sub
   
   For i = 0 To picInfo.count - 1
     picInfo(i).Visible = False
   Next i
   
   i = Val(tbsInfo.SelectedItem.Index - 1)
   
   picInfo(i).Visible = True

Exit Sub
Err:
    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tbsInfo_Click()"
    Resume Next
End Sub

Private Sub tbsInteractions_Click()
 On Error GoTo Err
 
   Dim i, i2 As Integer
   
5:    i2 = Val(tbsInteractions.SelectedItem.Index - 1)
6:    If picITab(i2).Visible = True Then Exit Sub
   
8:    For i = 0 To picITab.count - 1
9:      picITab(i).Visible = False
10:   Next i
   
12:    i = Val(tbsInteractions.SelectedItem.Index - 1)
13:    picITab(i).Visible = True

    Exit Sub
16:
Err:
17:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tbsMenu_Click()"
    Resume Next
End Sub

Private Sub tbsMenu_Click()
 
2: On Error GoTo Err

4:    Dim i, i2, i3 As Integer
      
6:    i2 = Val(tbsMenu.SelectedItem.Index - 1)
8:    If picTab(i2).Visible = True Then Exit Sub
      
10:   For i = 0 To picTab.count - 1
11:      picTab(i).Visible = False
12:   Next i

14:   i = Val(tbsMenu.SelectedItem.Index - 1)
15:   picTab(i).Refresh
16:   picTab(i).Visible = True

18:   If i = 5 Then Form_Resize: SCI_Focus

#If Not Status Then
21:   If i = 6 Then ' if not status then
22:       If tbsStatus.Enabled Then _
                lstStatus(0).AddItem "Status desabled..": _
                tbsStatus.Enabled = False: _
                picTab(6).Enabled = False
26:   End If
#End If

29:   If frmEditScintilla.Visible Then frmEditScintilla.Visible = False

31:   Exit Sub
32:
Err:
34:   HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tbsMenu_Click()"
35:   Resume Next
End Sub

Private Sub tbsSecurity_Click()

2: On Error GoTo Err
 
4:   Dim i, i2 As Integer
   
6:   i2 = Val(tbsSecurity.SelectedItem.Index - 1)
7:   If picSTab(i2).Visible = True Then Exit Sub
    
9:   For i = 0 To picSTab.count - 1
10:     picSTab(i).Visible = False
11:   Next i
   
13:   i = Val(tbsSecurity.SelectedItem.Index - 1)
14:   picSTab(i).Visible = True
   
16: Exit Sub
17:
Err:
19:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tbsSecurity_Click()"
       Resume Next
End Sub

Private Sub tbsStatus_Click()
 On Error GoTo Err
 
   Dim i, i2 As Integer
   
4:    i2 = Val(tbsStatus.SelectedItem.Index - 1)
5:    If picStatus(i2).Visible = True Then Exit Sub
   
7:    For i = 0 To picStatus.count - 1
8:      picStatus(i).Visible = False
9:    Next i
   
11:  i = Val(tbsStatus.SelectedItem.Index - 1)
12:  picStatus(i).Visible = True

     Exit Sub
15:
Err:
17:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tbsMenu_Click()"
     Resume Next
End Sub

Private Sub tmrBackground_Timer()
1:    Static bytCount As Byte

3:    On Error GoTo Err

5:   If g_objSettings.DynUpdate Then UpdateIPs_Timer
' Comment: should'nt we set it at a fixed 10 minutes ? UpdateIPs_Timer can be call in sub load at hub startup...
   
'***PLAN***
9:   If g_objSettings.EnabledScheduler Then Plan_Timer
'   If g_objSettings.EnabledScheduler Then TriggerCmds
' Comment: isn't a 5 minutes accuracy be enough or it must really stay 1 minute ?
'***PLAN END***

14:    If DateDiff("n", Now, m_datForceDNSUpdate) > 0 Then
15:        m_datForceDNSUpdate = Empty
16:    End If

18:    bytCount = bytCount + 1

'   If (bytCount Mod 15) = 0 Then (call the sub to If see if update is needed)
'   if protection is added against possible propagation delay, it could go in (mod 10)

    'Check if we should do anything
24:    If (bytCount Mod 10) = 0 Then
        'Remove users logging in for more than 5 minute (even if we make the check
        '                                                every 10 minutes)
27:        g_colUsers.CheckExtendedLogIn
28:        UpdateDNSs

30:        If (bytCount Mod 20) = 0 Then
            'Register the hub if needed
31:            If g_objSettings.AutoRegister Then
32:                For Each m_wskLoopItem In wskRegister
33:                    If m_wskLoopItem.State Then m_wskLoopItem.Close: DoEvents
34:                    m_wskLoopItem.Connect
35:                Next
    #If SVN Then
37:        g_objFileAccess.AppendFile G_LOGPATH, "m_wskLoopItem.Connect: " & m_wskLoopItem.RemoteHost
    #End If
39:                Set m_wskLoopItem = Nothing
40:            End If

            'If greater then 59 (ie an hour), remove all outdated
            'hammer/password guess records
44:            If bytCount > 59 Then
45:                CheckOutdatedRecords
46:                bytCount = 0
47:            End If
48:        End If

50:        Call RefreshGUI(True)

52:    End If

54:    Exit Sub
    
56:
Err:
58:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tmrBackground_Timer(bytCount = " & bytCount & ")"
End Sub

'***PLAN***
Private Sub Plan_Timer()
1:    Dim sMsg, n, sCMD, sParameter, oUser, t
2:    Dim curUser As clsUser

4:    t = g_colScheduler.PlanDo

6:    Do Until t = vbNullString
7:        sMsg = Trim(g_objFunctions.SplitParameter(t, "$"))
8:        If sMsg <> vbNullString Then
9:            sParameter = sMsg
10:            sCMD = g_objFunctions.SplitParameter(sParameter, " ")
11:            For Each oUser In g_colUsers
12:                If g_colUsers.ItemByName(CStr(oUser.sName)).bOperator Then
13:                    Set curUser = g_colUsers.ItemByName(CStr(oUser.sName))
14:                    sCMD = g_objFunctions.AfterFirst(sCMD, Chr(g_objSettings.CPrefix))
15:                    ProcessTrigger curUser, CStr(sCMD & " " & sParameter), True
16:                    Exit For
17:                End If
18:            Next
19:        End If
20:    Loop
End Sub

Private Sub tmrSysInfo_Timer()
1:    On Error GoTo Err

       'Close sub for not to use memory ..
4:     If Me.Visible = False Or Me.WindowState = vbMinimized Then _
           Exit Sub
    
       'Get global memory status
7:     Dim tMS As MEMORYSTATUS
8:     Dim iMonths As Integer, iWeeks As Integer, iDays As Integer, iHours As Integer, iMinutes As Integer, iSeconds As Integer
9:     Dim currTime As Long
10:    Dim t As String
11:    Dim strSysTrayTip As String
     
       'Memory System ///////////////////////////////////////////////////////
       'Length of structure
14:    tMS.dwLength = Len(tMS)
       'Get global memory status
16:    GlobalMemoryStatus tMS
    
       'Print memory status
19:    txtSystem(0).Text = Format$(tMS.dwTotalPhys / 1024, "###,###,###") & " Kb"
20:    txtSystem(1).Text = Format$(tMS.dwAvailPhys / 1024, "###,###,###") & " Kb"
21:    txtSystem(2).Text = Format$(tMS.dwTotalVirtual / 1024, "###,###,###") & " Kb"
22:    txtSystem(3).Text = Format$(tMS.dwAvailVirtual / 1024, "###,###,###") & " Kb"
23:    txtSystem(4).Text = Format$(tMS.dwTotalPageFile / 1024, "###,###,###") & " Kb"
24:    txtSystem(5).Text = Format$(tMS.dwAvailPageFile / 1024, "###,###,###") & " Kb"
    
26:    pgrMemory.Value = tMS.dwMemoryLoad
27:    lblTitle(13).Caption = "Memory Usage: " & tMS.dwMemoryLoad & "%"
28:    stbMain.Panels(7).Text = tMS.dwMemoryLoad & "%"
       '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
       
       'Hub UpTime //////////////////////////////////////////////////////////
       If m_blnServing Then
          'Get date of the server started
33:       currTime = DateDiff("s", ServingDate, DateTime.Now)
          'Calc.. date iMinutes/iSeconds/iHours/iDays/iWeeks and iMonths
35:       iSeconds = currTime Mod 60
36:       iMinutes = (currTime \ 60) Mod 60
37:       iHours = (currTime \ 3600) Mod 24
38:       iDays = (currTime \ 86400) Mod 7
39:       iWeeks = currTime \ 604800 Mod 4
40:       iMonths = (currTime \ 2419200)

42:       If iMonths > 0 Then t = "[M:" & iMonths & "["
43:       If iWeeks > 0 Then t = "[W:" & iWeeks & "["
44:       If iDays > 0 Then t = t & "[D:" & iDays & "] "

46:       t = t & strZero(iHours, 2) & ":"
47:       t = t & strZero(iMinutes, 2) & ":"
48:       t = t & strZero(iSeconds, 2)

50:       txtUpTime.Text = t
51:       stbMain.Panels(1).Text = t
52:    Else
53:       t = "00:00:00"
54:       txtUpTime.Text = t
55:       stbMain.Panels(1).Text = "00:00:00"
56:    End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
       
59: Exit Sub
60:
Err:
62:
63:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.tmrSysInfo_Timer"
64:  Resume Next
End Sub

Private Sub txtData_Change(Index As Integer)
'------------------------------------------------------------------
'Purpose:   Update settings variables from text boxs
'
'Params:        Index
'               Index of the text box in the text box collection
'
'Added: Former Dev
'
'Changed:       RTD svn ?
'               TheNOP svn 26
'
'Comment:       New Cases should not be added if an actual Case can be use.
'               Just add the index number to the proper Case.
'------------------------------------------------------------------
15:    On Error GoTo Err

    Select Case Index
        Case 18 'Prefix
17:            If LenB(txtData(Index).Text) Then _
                g_objSettings.CPrefix = AscW(txtData(18).Text) _
            Else _
                txtData(Index).Text = ChrW$(g_objSettings.CPrefix)
        Case 19 'Long values
21:            CallByName g_objSettings, txtData(Index).Tag, VbLet, CLng(txtData(Index).Text)
        Case 15, 12, 13, 14 'Byte values
22:            CallByName g_objSettings, txtData(Index).Tag, VbLet, CByte(txtData(Index).Text)
'--ROLL NEW REDIRECT TXTDATA BOXES INDEXES----------------------
Case 5
24:        g_objSettings.ForMinShareRedirectAddress = txtData(5).Text
Case 25
25:        g_objSettings.ForMaxShareRedirectAddress = txtData(25).Text
Case 26
26:        g_objSettings.ForMaxSlotsRedirectAddress = txtData(26).Text
Case 27
27:        g_objSettings.ForMinSlotsRedirectAddress = txtData(27).Text
Case 28
28:        g_objSettings.ForTooOldNMDCRedirectAddress = txtData(28).Text
Case 29
29:        g_objSettings.ForMaxHubsRedirectAddress = txtData(29).Text
Case 30
30:        g_objSettings.ForNoTagRedirectAddress = txtData(30).Text
Case 31
31:        g_objSettings.ForSlotPerHubRedirectAddress = txtData(31).Text
Case 32
32:        g_objSettings.ForTooOldDcppRedirectAddress = txtData(32).Text
Case 33
33:        g_objSettings.ForBWPerSlotRedirectAddress = txtData(33).Text
Case 34
34:        g_objSettings.ForFakeTagRedirectAddress = txtData(34).Text
Case 35
35:        g_objSettings.ForFakeShareRedirectAddress = txtData(35).Text
Case 24
36:        g_objSettings.ForPasModeRedirectAddress = txtData(24).Text
            
'------------------AND END HERE-----------------------------------------
            'm_arrRedirectIPs = Split(g_objSettings.RedirectAddress, ";")
            'm_lngRedirectUB = UBound(m_arrRedirectIPs)
       Case 36
41:            g_objSettings.RedirectAddress = txtData(36).Text
42:            m_arrRedirectIPs = Split(g_objSettings.RedirectAddress, ";")
43:            m_lngRedirectUB = UBound(m_arrRedirectIPs)
            'If UBound = -1, then set the UBound to 0 to prevent crashes
            'otherwise set the RedirectIP to the first one
46:            If m_lngRedirectUB = -1 Then _
                m_lngRedirectUB = 0 _
            Else _
                g_objSettings.RedirectIP = m_arrRedirectIPs(0)
        Case 10 'Min share
50:            g_objSettings.IMinShare = CDbl(txtData(Index).Text)
51:            g_objSettings.MinShare = g_objSettings.IMinShare * (1024 ^ g_objSettings.MinShareSize)
        Case 22 'Max share
52:            g_objSettings.IMaxShare = CDbl(txtData(Index).Text)
53:            g_objSettings.MaxShare = g_objSettings.IMaxShare * (1024 ^ g_objSettings.MaxShareSize)
        Case 9, 11, 16, 17 'Double values
54:            CallByName g_objSettings, txtData(Index).Tag, VbLet, CDbl(txtData(Index).Text)

        Case Else 'Regular strings
56:            CallByName g_objSettings, txtData(Index).Tag, VbLet, txtData(Index).Text
57:    End Select
    
59:    Exit Sub
    
61:
Err:
62:    txtData(Index).Text = CallByName(g_objSettings, txtData(Index).Tag, VbGet)
End Sub

Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
1:    On Error GoTo Err
    
    'For numeric settings in textboxes, only allow numbers and backspace
    '(as well as decimals where required)

    Select Case Index
        Case 12, 13, 14, 15, 19 'Longs / Integers / Bytes
            Select Case KeyAscii
                Case 48 To 57, 8
                Case Else
6:                    KeyAscii = 0
7:            End Select
        Case 9, 10, 11, 16, 17, 22 'Doubles
            Select Case KeyAscii
                Case 48 To 57, 8, 46, 44
                Case Else
8:                    KeyAscii = 0
9:            End Select
10:    End Select
    
12:    Exit Sub
    
14:
Err:
15:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.txtData_KeyPress(" & Index & ")"
End Sub

Private Sub vslData_Change(Index As Integer)
1:    On Error GoTo Err

3:    CallByName g_objSettings, vslData(Index).Tag, VbLet, vslData(Index).Value

    'Update linked label caption
    Select Case Index
        Case 0: txtVSl(0).Text = g_objSettings.DefaultBanTime
        Case 1: txtVSl(9).Text = g_objSettings.MaxUsers
        Case 2: txtVSl(10).Text = g_objSettings.MinPassiveSearchLen
        Case 3: txtVSl(2).Text = g_objSettings.FWInterval
        Case 4: txtVSl(3).Text = g_objSettings.FWBanLength
        Case 5: txtVSl(5).Text = g_objSettings.FWMyINFO
        Case 6: txtVSl(7).Text = g_objSettings.FWGetNickList
        Case 7: txtVSl(4).Text = g_objSettings.FWActiveSearch
        Case 8: txtVSl(6).Text = g_objSettings.FWPassiveSearch
        Case 9: txtVSl(1).Text = g_objSettings.MaxPassAttempts
        Case 10: txtVSl(12).Text = g_objSettings.MinSearchCls
        Case 11: txtVSl(13).Text = g_objSettings.MinConnectCls
        Case 12: txtVSl(11).Text = g_objSettings.MaxMessageLen
6:    End Select
    
8:    Exit Sub
    
10:
Err:
11:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.vslData_Change(" & Index & ")"
End Sub

'------------------------------------------------------------------------------
'Winsock events
'------------------------------------------------------------------------------
Private Sub wskListen_Close(Index As Integer)
1:    On Error GoTo Err
    
    'Ignore error, make then listen again
4:    wskListen(Index).Close
5:    DoEvents
6:    wskListen(Index).Listen
    
8:    Exit Sub
    
10:
Err:
End Sub
Private Sub wskListen_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
1:    On Error GoTo Err

    'Ignore error, make then listen again
4:    wskListen(Index).Close
5:    DoEvents
6:    wskListen(Index).Listen
    
8:    Exit Sub
    
10:
Err:
End Sub
Private Sub wskLoop_Close(Index As Integer)
1:    Dim curUser As clsUser

3:    On Error GoTo Err

    'Make sure winsock is closed, otherwise we'll get an endless loop of Close events
6:    wskLoop(Index).Close

8:    Set curUser = g_colUsers.ItemByWinsockIndex(Index)

    'Remove them from the collection as needed
11:    If ObjPtr(curUser) Then
12:        If ObjPtr(curUser.Winsock) Then
13:            g_colUsers.Remove Index

            #If Status Then
16:                g_objStatus.URemove Index
            #End If

            'Send out quit message
20:            If curUser.State = Logged_In Then
21:                If curUser.Visible Then g_colUsers.SendToAll "$Quit " & curUser.sName & "|"
22:            End If

              'Call the sub UserQuit()
25:            SEvent_UserQuit curUser

              'Show pupop notification ..
26:           If g_objSettings.PopUpOpDisconected And curUser.Class >= 6 Then _
                    g_objFunctions.ShowBallon "PTDCH  - " & g_objSettings.HubName, "Op Disconected" & vbNewLine & "Nick: " & curUser.sName, 0, True
                    
27:            Set curUser.Winsock = Nothing
28:        End If
29:    End If

    #If COLFREESOCKS Then
32:        On Error Resume Next
33:        m_colFreeSocks.Add wskLoop(Index), CStr(Index)
    #End If
        
36:    Exit Sub

38:
Err:
40:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.wskLoop_Close()"
41:    Resume Next
End Sub
Private Sub wskListen_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    
2:    Static lngDTick As Long
    'Static strIP    As String
    
5:    Dim lngTick    As Long
    
7:    Dim intIndex    As Integer
8:    Dim blnFull     As Boolean
9:    Dim lng         As Long
10:    Dim blnLoaded   As Boolean
11:    Dim wskUser     As Winsock

13:    On Error GoTo Err
   
    'Check if the hub is full
16:    blnFull = (g_colUsers.count >= g_objSettings.MaxUsers)

    #If COLFREESOCKS Then
        'Check for free socket in collection
20:        If m_colFreeSocks.count Then
21:            Set wskUser = m_colFreeSocks(1)
22:            intIndex = wskUser.Index
23:            m_colFreeSocks.Remove CStr(intIndex)
24:        Else
            'Get an unused winsock
26:            intIndex = wskLoop.UBound + 1
        
28:            Load wskLoop(intIndex)
29:            Set wskUser = wskLoop(intIndex)
30:        End If

    #Else
33:        intIndex = wskLoop.UBound

        'If it's full, we're more likely to find a free winsock at the end
36:        If blnFull Then
37:            For lng = intIndex To 0
38:                If wskLoop(lng).State = 0 Then intIndex = lng: blnLoaded = True: Exit For
39:            Next
40:        Else
41:            For lng = 0 To intIndex
42:                If wskLoop(lng).State = 0 Then intIndex = lng: blnLoaded = True: Exit For
43:            Next
44:        End If

        'Load new winsock object if it never found one
47:        If Not blnLoaded Then
48:            intIndex = intIndex + 1
49:            Load wskLoop(intIndex)
50:        End If

52:        Set wskUser = wskLoop(intIndex)

    #End If
      
    'Accept the request
57:    wskUser.Accept requestID
    
    'Check if their IP is banned
60:    requestID = g_objIPBans.Check(wskUser.RemoteHostIP)
    
    Select Case requestID
        Case 0 'Not banned
            'Check for hammering
63:            If Not UpdateConnectAttempt(wskUser, False) Then

65:                lngTick = GetTickCount
                '1 second = 1 000 milliseconds, expected default setting (< 250 ?) ~ 500 milliseconds
67:                If Abs((lngTick - lngDTick)) > g_objSettings.ConDropInterval Then
68:                    lngDTick = lngTick
69:                Else 'moved here to fix a DoS possibility...
70:                    wskUser.SendData "<" & g_objSettings.BotName & "> " & Replace(vbWelcome, "%[UpTime]", HubUpTime()) & "|"
71:                    wskUser.SendData "<" & g_objSettings.BotName & "> If you have had problems to enter here: try again in ± 10 secs or port " & Replace(g_objSettings.Ports, ";", " or ") & "|"
72:                    DoEvents
73:                    wskUser.Close
                    
                    #If COLFREESOCKS Then
76:                        m_colFreeSocks.Add wskUser, CStr(intIndex)
                    #End If
                    
79:                    Set wskUser = Nothing
                    
81:                    Exit Sub
82:                End If
 
                'Redirect as needed
85:                If g_objSettings.AutoRedirect Then
86:                    NextRedirect

88:                    wskUser.SendData "<" & g_objSettings.BotName & "> " & g_objFunctions.GetENLangStr("RedirectedTo") & g_objSettings.RedirectIP & "|$ForceMove " & g_objSettings.RedirectIP & "|"
                    
90:                    DoEvents
91:                    wskUser.Close
                    
                    #If COLFREESOCKS Then
94:                        m_colFreeSocks.Add wskUser, CStr(intIndex)
                    #End If
                    
97:                    Set wskUser = Nothing
                    
99:                    Exit Sub
100:                Else
                    'If it's full, check if we need to redirect
102:                    If blnFull Then
                        'Certain redirect types must wait till the user sends their nick
104:                        If Not g_objSettings.AutoRedirectFullNonReg Then
105:                            If Not g_objSettings.AutoRedirectFullNonOps Then
                                'Redirect as needed
107:                                If g_objSettings.AutoRedirectFull Then
108:                                    NextRedirect
109:                                    wskUser.SendData "<" & g_objSettings.BotName & "> " & g_objFunctions.GetENLangStr("FullRedirTo") & g_objSettings.RedirectIP & "|" & "$ForceMove " & g_objSettings.RedirectIP & "|"
110:                                Else
111:                                    wskUser.SendData "<" & g_objSettings.BotName & "> " & g_objFunctions.GetENLangStr("Full") & "|"
112:                                End If
                    
114:                                DoEvents
115:                                wskUser.Close
                    
                                #If COLFREESOCKS Then
118:                                    m_colFreeSocks.Add wskUser, CStr(intIndex)
                                #End If
                    
121:                                Set wskUser = Nothing
                                
123:                                Exit Sub
124:                            End If
125:                        End If
126:                    End If
                    
                    'If we get this far, the user is not connected to the hub
129:                    Set m_objLoopUser = g_colUsers.Add(intIndex)
130:                    Set m_objLoopUser.Winsock = wskUser
                    
                    #If Status Then
133:                        g_objStatus.UAdd m_objLoopUser
                    #End If
                    
                    'Send lock
137:                    wskUser.SendData "$Lock " & vbLock & "|"
    


                #If SVN Then
142:                On Error Resume Next
143:                    g_objFileAccess.AppendFile G_LOGPATH, Now & " --> " & wskUser.RemoteHostIP & " - " & "$Lock " & vbLock & "|"
144:                On Error GoTo Err
                #End If
                


                    'Call the sub AttemptedConnection(sIP)
149:                    SEvent_AttemptedConnection m_objLoopUser.IP

151:                End If
152:            End If

        Case -1 'Perm banned
            'Use descriptive ban message if needed (gives length of ban)
154:            If g_objSettings.DescriptiveBanMsg Then
155:                wskUser.SendData "<" & g_objSettings.BotName & "> " & g_objFunctions.GetENLangStr("IPPermBan") & "|"
156:            Else
157:                wskUser.SendData "<" & g_objSettings.BotName & "> " & g_objFunctions.GetENLangStr("IPBanned") & "|"
158:            End If
            
160:            DoEvents
161:            wskUser.Close
            
            #If COLFREESOCKS Then
164:                m_colFreeSocks.Add wskUser, CStr(intIndex)
            #End If
        Case Else 'Temp banned
            'Use descriptive ban message if needed (gives length of ban)
167:            If g_objSettings.DescriptiveBanMsg Then
168:                wskUser.SendData "<" & g_objSettings.BotName & "> " & g_objFunctions.GetENLangStr("IPTempBan") & MinToDate(requestID) & ".|"
169:            Else
170:                wskUser.SendData "<" & g_objSettings.BotName & "> " & g_objFunctions.GetENLangStr("IPBanned") & "|"
171:            End If

173:            DoEvents
174:            wskUser.Close
            
            #If COLFREESOCKS Then
177:                m_colFreeSocks.Add wskUser, CStr(intIndex)
            #End If
179:    End Select
    
181:    Set m_objLoopUser = Nothing

183:    Exit Sub
    
185:
Err:
186:    On Error Resume Next

188:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.wskListen_ConnectionRequest()", Err.LastDllError
189:    Set m_objLoopUser = Nothing
    
    'Make sure that if the connection was accepted it is closed
192:    If ObjPtr(wskUser) Then
193:        If wskUser.State = 7 Then
194:            wskUser.Close
195:            If g_colUsers.Exists(intIndex) Then g_colUsers.Remove intIndex
196:        End If
        
198:        Set wskUser = Nothing
199:    End If
End Sub
Private Sub wskLoop_DataArrival(Index As Integer, ByVal bytesTotal As Long)
1:    Dim lngPos      As Long
2:    Dim strIP       As String
3:    Dim strData     As String
4:    Dim strKey      As String
5:    Dim strParts    As String
6:    Dim strCommand  As String
7:    Dim curUser     As clsUser
8:    Dim strObjIP    As String
9:    Dim strTMyinfosStr  As String

11:    On Error GoTo Err
    
    'Prepare object / data
14:    Set curUser = g_colUsers.ItemByWinsockIndex(Index)
    
    #If OBJECTNOTSET Then
17:        If ObjPtr(curUser) = 0 Then
18:            wskLoop_Close Index
19:            Exit Sub
20:        End If
    #End If
    
23:    wskLoop(Index).GetData strData, vbString

    'Concat fragmented data if any
26:    If LenB(curUser.DataFragment) > 0 Then
27:        strData = curUser.DataFragment & strData
28:        curUser.DataFragment = vbNullString
29:    End If

    #If SVN Then
32:        If LenB(strData) Then
33:            g_objFileAccess.AppendFile G_LOGPATH, Now & " <-- " & curUser.IP & " - " & curUser.sName & " - " & strData
34:        End If
    #End If

    'Using numbers (especially longs since they are 32 bit) is much faster
    'than strings. This is an optimized approach to the DC protocol and is
    'unique to DDCH (as far as I can tell from other open source hubs)
      
    'Rather than examining the protocol as a whole, it checks the first
    'character and the length, where possible, otherwise the last character
    '(LenB is faster than AscW(RightB$(strKey, 2))
    
    #If FLASHCHAT Then
46:        If curUser.NullCharSeparator Then _
            strData = Replace(strData, vbNullChar, "|")
    #End If
    
50:    lngPos = InStrB(1, strData, "|")
    
52:    On Error GoTo LoopErr
    
    'Do while there is a | in strData
55:    Do While lngPos
        
57:        strCommand = LeftB$(strData, lngPos - 1)
58:        strData = MidB$(strData, lngPos + 2)
        
        #If PreDataArrival Then
61:            If LenB(strCommand) Then
62:                On Error Resume Next

64:                For lngPos = 1 To m_lngScriptEventsUB
65:                    If m_arrScriptEvents(lngPos, vbSPreDataArrival) Then _
                        strCommand = ScriptControl(lngPos).Run("PreDataArrival", curUser, strCommand)
67:                    If LenB(strCommand) = 0 Then Exit For
68:                Next
69:                On Error GoTo Err
70:            End If
        #End If
        
        'Don't process command if it's empty (ignore if it's a single char)
74:        If LenB(strCommand) > 2 Then
            'Find out type of command it is
            '   -- $ = Protocol command
            '   -- < = Main chat message
            
            #If Status Then
                'Add to listbox
81:                g_objStatus.MAdd strCommand
            #End If
            
            '#If PREDATAARRIVAL Then
            '    On Error GoTo AfterPD
            '
            '    'This runs the PreDataArrival event
            '    '
            '    '  -- Parameters : curUser (the current user's clsUser object)
            '    '                : strData (data that was sent)
            '    '  -- Format     : Function PreDataArrival(curUser, strData)
            '    '
            '    '  -- Called when a user sends data to the hub, but before the hub parses
            '    '     it
            '    '  -- It should return the string it should parse
            '
            '    If m_intPDIndex Then strCommand = ScriptControl(m_intPDIndex).Run("PreDataArrival", curUser, strCommand)
            '    If Not LenB(strCommand) > 2 Then GoTo NextLoop
            '
100:
AfterPD:    '
            '    On Error GoTo Err
            '#End If
            
            Select Case AscW(strCommand)
                Case 36 '$
                    'Check if there is a " "; if there is remove the key from
                    'data and seperate it's params
106:                    lngPos = InStrB(1, strCommand, " ")
                    
108:                    If lngPos Then
109:                        strKey = MidB$(strCommand, 3, lngPos - 3)
110:                        strParts = MidB$(strCommand, lngPos + 2)
111:                    Else
112:                        strKey = MidB$(strCommand, 3)
113:                    End If
                    
                    'Start parsing!
                    
                    'Notes -- This is structured in such a way so that the most
                    '         common messages are at the top of the Select Case;
                    '         this makes it even more efficent =)
                    '
                    '      -- Also due to this format, DC protocol commands ARE
                    '         CASE SENSITIVE; I will NEVER add support for bots/
                    '         etc which do not follow the protocol properly
                    '
                    '      -- In case some of you are wondering, this is a quite
                    '         inaccurate way of parsing messages...if DDCH does
                    '         not support a message, it might think it's another
                    '         unrelated message. For that reason, I support
                    '         all documented protocol extensions (at least to the
                    '         point of making sure it doesn't get confused)
                    
                    Select Case AscW(strKey)
                        Case 83 'S
                            'Possible messages :
                            '   -- Search
                            '   -- SR (passive search result; active is sent via UDP directly to the client)
                            '   -- Supports
                            
                            Select Case LenB(strKey)
                                Case 12
                                    'Search
                                    '
                                    '   -- Format   : $Search [<ip:port>//Hub:<name>] <T/F>?<T/F>?<size>?<type>?<search>|
                                    '   -- Response : N/A (clients will respond with $SR)
                                    '
                                    '   -- Standard protocol message
                                    '   -- It has either an ip and port or a name (active versus passive).
                                    '      First T/F toggles whether or not the size of files is restricted.
                                    '      Second T/F toggles whether limit is upper (T) or lower (F)
                                    '      <size> is the size in bytes of the file
                                    '      <type> is the type of file it can be (document, audio, etc)
                                    '      <search> is the string to search for (spaces are converted to $)
                                    
                                    'If the user has not sent their MyINFO string, check if bots are allowed
                                    'The second check is 95-99% effective, as I've yet to see a search tool
                                    'which sends GetNickList (you must request it to search)
                                
154:                                    If g_objSettings.PreventSearchBots Then
155:                                        If Not curUser.State = Logged_In Then wskLoop_Close Index: Exit Sub
156:                                        If Not curUser.QNL Then wskLoop_Close Index: Exit Sub
157:                                    End If
                                    

160:                                    If Not g_objSettings.ChatOnly Then
161:                                        If curUser.ChatOnly Then curUser.Kick g_objSettings.DefaultBanTime
                                    
163:                                        If curUser.Class >= g_objSettings.MinSearchCls Then
164:                                            If AscW(strParts) = 72 Then
                                                'Make sure the user isn't flooding
166:                                                If g_objSettings.EnableFloodWall Then _
                                                    If curUser.FloodCheck(1) Then _
                                                        Exit Sub
                                                    
                                                'Set to passive
171:                                                curUser.Passive = True
                                            
                                                'Allow search is passive searches are not disabled
                                                Select Case g_objSettings.MinPassiveSearchLen
                                                    Case -1 'Disabled
                                                    Case 0, Is <= Len(Trim$(Mid$(strParts, InStrRev(strParts, "?") + 1)))
                                                        'Make sure they aren't faking their name
175:                                                        If curUser.sName = MidB$(strParts, 9, InStrB(1, strParts, " ") - 9) Then
176:                                                            lngPos = ObjPtr(curUser)
                                                        
178:                                                            If g_objSettings.MinClsSearchSend Then
                                                                'Don't send search to person who sent it
180:                                                                For Each m_objLoopUser In g_colUsers
181:                                                                    If Not m_objLoopUser.Passive Then _
                                                                        If m_objLoopUser.Visible Then _
                                                                            If Not ObjPtr(m_objLoopUser) = lngPos Then _
                                                                                m_objLoopUser.SendData strCommand & "|"
185:                                                                Next
186:                                                            Else
                                                                'Don't send search to person who sent it
188:                                                                For Each m_objLoopUser In g_colUsers
189:                                                                    If Not m_objLoopUser.Passive Then _
                                                                        If m_objLoopUser.Class >= g_objSettings.MinSearchCls Then _
                                                                            If m_objLoopUser.Visible Then _
                                                                                If Not ObjPtr(m_objLoopUser) = lngPos Then _
                                                                                    m_objLoopUser.SendData strCommand & "|"
194:                                                                Next
195:                                                            End If
                                                        
197:                                                            Set m_objLoopUser = Nothing
198:                                                        Else
199:                                                            wskLoop_Close Index
200:                                                            Exit Sub
201:                                                        End If
202:                                                End Select
203:                                            Else
                                                'Make sure they aren't flooding
205:                                                If g_objSettings.EnableFloodWall Then _
                                                    If curUser.FloodCheck(0) Then _
                                                        Exit Sub
                                                        
                                                'Set to active
210:                                                curUser.Passive = False
                                                
212:                                                strKey = curUser.IP
                                                
                                                'Find out if the IP is a local range (skip IP match check if it is)
                                                Select Case CByte(LeftB$(strKey, InStrB(1, strKey, ".") - 1))
                                                    Case 192, 127, 10
                                                    Case Else
215:                                                        strParts = LeftB$(strParts, InStrB(1, strParts, ":") - 1)
                                                        
                                                        'If IP doesn't match, then fix it
218:                                                        If Not strKey = strParts Then _
                                                            strCommand = Replace(strCommand, strParts, strKey, 1, 1)
220:                                                End Select
                                                
222:                                                lngPos = ObjPtr(curUser)
                                                        
224:                                                If g_objSettings.MinClsSearchSend Then
                                                    'Don't send search to person who sent it
226:                                                    For Each m_objLoopUser In g_colUsers
227:                                                        If m_objLoopUser.Visible Then _
                                                            If Not ObjPtr(m_objLoopUser) = lngPos Then _
                                                                m_objLoopUser.SendData strCommand & "|"
230:                                                    Next
231:                                                Else
                                                    'Don't send search to person who sent it
233:                                                    For Each m_objLoopUser In g_colUsers
234:                                                        If m_objLoopUser.Class >= g_objSettings.MinSearchCls Then _
                                                            If m_objLoopUser.Visible Then _
                                                                If Not ObjPtr(m_objLoopUser) = lngPos Then _
                                                                    m_objLoopUser.SendData strCommand & "|"
238:                                                    Next
239:                                                End If
                                                    
241:                                                Set m_objLoopUser = Nothing
242:                                            End If
243:                                        End If
244:                                    End If
                                Case 4
                                    'SR
                                    '
                                    '   -- Format   : $SR <from> <fpath><char5><fsize> <fslots>/<tslots><char5><hubname> (<hubip>[:<hubport>])<char5><to>|
                                    '               : $SR <from> <directory> <fslots>/<tslots><char5><hubname> (<hubip>[:<hubport>])<char5><to>|
                                    '   -- Response : N/A
                                    '
                                    '   -- Standard protocol message
                                    '   -- Result for passive searches
                                    '   -- Two forms; first is for files, second is for directories
                                    
                                    'Make sure we even need to bother with the check
                                    Select Case False
                                        Case g_objSettings.OPBypass, (curUser.Class >= vip)
256:                                            lngPos = InStrB(1, strParts, " ")
                                            
                                            'Make sure their nickname matches the search result name
259:                                            If LeftB$(strParts, lngPos - 1) = curUser.sName Then
260:                                                strParts = MidB$(strParts, lngPos + 2)
                                            
262:                                                lngPos = InStrB(1, strParts, "/") + 2
263:                                                strParts = MidB$(strParts, lngPos, InStrB(lngPos, strParts, vbChar5) - lngPos)
                                            
                                                'If the total slots value isn't numerical, kick
266:                                                If IsNumeric(strParts) Then
267:                                                    lngPos = GetByte(CLng(strParts))
                                                    
                                                    'Check min slots
270:                                                    If g_objSettings.MinSlots Then
271:                                                        If lngPos < g_objSettings.MinSlots Then
272:                                                            FailedConf curUser, MinSlots
273:                                                            Exit Sub
274:                                                        End If
275:                                                    End If
                                                   
                                                    
                                                    'Check max slots
279:                                                    If g_objSettings.MaxSlots Then
280:                                                        If lngPos > g_objSettings.MaxSlots Then
281:                                                            FailedConf curUser, MaxSlots
282:                                                            Exit Sub
283:                                                        End If
284:                                                    End If
285:                                                Else
286:                                                    wskLoop_Close Index
287:                                                    Exit Sub
288:                                                End If
289:                                            Else
290:                                                curUser.Kick 60
291:                                                Exit Sub
292:                                            End If
                                            
                                            'checking slot here is risky.(MyINFO are delayed a bit client side in order not to spam hubs.)
                                            'If it finds two vbChar5, then it is a directory, else a file
                                            'If (LenB(strParts) - LenB(Replace(strParts, vbChar5, vbNullString))) = 4 Then
                                            '    strKey = LeftB$(strParts, InStrB(1, strParts, vbChar5) - 1)
                                            '    lngPos = CLng(MidB$(strKey, LenB(strKey) - InStrB(1, StrReverse(strKey), "/") + 2))
                                            'Else
                                            '    strKey = MidB$(strParts, InStrB(InStrB(1, strParts, vbChar5), strParts, " "))
                                            '    lngPos = InStrB(1, strKey, "/") + 2
                                            '    lngPos = CLng(MidB$(strKey, lngPos, InStrB(lngPos, strKey, vbChar5) - lngPos))
                                            'End If
                                           '
                                           ' 'Check for fake slots
                                           ' If curUser.Slots Then _
                                           '     If Not curUser.Slots = lngPos Then _
                                           '         FailedConf curUser, FakeTag: Exit Sub

310:                                    End Select
                                    
                                    'Find out who the result should be sent to
313:                                    lngPos = InStrRev(strCommand, vbChar5)
314:                                    strParts = Mid$(strCommand, lngPos + 1)
                                        
                                    'If online, send result to client
317:                                    If g_colUsers.Online(strParts) Then _
                                        g_colUsers.ItemByName(strParts).SendData Left$(strCommand, lngPos - 1) & "|"
                                Case 16
                                    'Supports
                                    '
                                    '   -- Format   : $Supports <ext> <ext_etc>|
                                    '   -- Response : $Supports <etc> <ext_etc>|
                                    '
                                    '   -- Protocol extension (in response to EXTENDEDPROTOCOL in Lock string)
                                    '   -- Allows client to extend abilities
                                    '   -- Only extensions which both the client and hub support
                                    '      should be sent back to the client
                                    
329:                                    strParts = strParts & " "
330:                                    lngPos = InStrB(1, strParts, " ")
                                    
332:                                    curUser.ZLine = False
333:                                    curUser.ZPipe = False
334:                                    curUser.QuickList = False
335:                                    curUser.NoHello = False
336:                                    curUser.UserCommand = False
337:                                    curUser.ChatOnly = False

                                    #If FLASHCHAT Then
340:                                        curUser.NullCharSeparator = False
                                    #End If
                                    
                                    'Find out which extensions both support
344:                                    Do While lngPos
345:                                        strKey = LeftB$(strParts, lngPos - 1)
346:                                        strParts = MidB$(strParts, lngPos + 2)

                                        Select Case strKey
                                            Case "QuickList"
348:                                                curUser.Supports = curUser.Supports & " QuickList"
349:                                                curUser.QuickList = True
350:                                                curUser.NoHello = True
                                            Case "UserCommand"
351:                                                curUser.Supports = curUser.Supports & " UserCommand"
352:                                                curUser.UserCommand = True
                                            Case "NoHello"
353:                                                curUser.Supports = curUser.Supports & " NoHello"
354:                                                curUser.NoHello = True
                                            Case "NoGetINFO", "TTHSearch", "UserIP2", "UserIP", "xKick", "BotINFO"
355:                                                curUser.Supports = curUser.Supports & " " & strKey
                                            Case "ZPipe"
356:                                                curUser.Supports = curUser.Supports & " ZPipe"
357:                                                curUser.ZPipe = True
                                            Case "ZLine"
358:                                                curUser.Supports = curUser.Supports & " ZLine"
359:                                                curUser.ZLine = True
                                            Case "ChatOnly"
360:                                                curUser.Supports = curUser.Supports & " ChatOnly"
361:                                                curUser.ChatOnly = True

                                        #If FLASHCHAT Then
                                            Case "NullCharSeparator"
364:                                                curUser.Supports = curUser.Supports & " NullCharSeparator"
365:                                                curUser.NullCharSeparator = True
                                        #End If

368:                                        End Select

370:                                        lngPos = InStrB(1, strParts, " ")
371:                                    Loop

                                    'Remote leading space
374:                                    strKey = LTrim$(curUser.Supports)
375:                                    curUser.Supports = strKey
                                    
                                    'If it supports UserCommand, then send the clear command
378:                                    If curUser.UserCommand Then _
                                        curUser.SendData "$Supports " & strKey & "|$UserCommand 255 7 |" _
                                    Else _
                                        curUser.SendData "$Supports " & strKey & "|"
382:                            End Select
                        Case 77 'M
                            'Possible messages :
                            '   -- MyINFO
                            '   -- MyPass
                            '   -- MultiConnectToMe (ignored)
                            '   -- MultiSearch (ignored)
                            '   -- MyIP (ignored)
                            
                            Select Case AscW(RightB$(strKey, 2))
                                Case 79
                                    'MyINFO
                                    '
                                    '   -- Format   : $MyINFO $ALL <name> <description>$ $<connection><char>$<email>$<share>$|
                                    '   -- Response : N/A
                                    '
                                    '   -- Standard protocol message
                                    '   -- Contains all the info of a user
                                    '   -- Once this is sent, and assuming it passes the rules,
                                    '      the client has logged in
                                    
                                    'See if we can limit MyINFOs sending when Hide MyINFOs is enabled.
                                    'only send to registered users if it has changed.
                                    
                                    'Check if the user is flooding
404:                                    If g_objSettings.EnableFloodWall Then _
                                        If curUser.FloodCheck(2) Then _
                                            Exit Sub

408:                                    If curUser.State = Logged_In Then
                                        'If the MyINFO string has changed, then continue
410:                                        If Not curUser.sMyInfoString = strCommand Then
                                            'If the name matches, process it
412:                                            If MidB$(strParts, 11, InStrB(13, strParts, " ") - 11) = curUser.sName Then
                                                'If it passes the rules, then send it out to all users
414:                                                If ProcessMyINFO(curUser, strParts) Then
                                                    'But only if the user is Visible
416:                                                    If curUser.Visible Then
417:                                                        For Each m_objLoopUser In g_colUsers
418:                                                            If curUser.State = Disconnected Then Exit Sub
                                                            '#If FLASHCHAT Then
420:                                                                If Not m_objLoopUser.ChatOnly Then
421:                                                                    If g_objSettings.HideMyinfos Then
422:                                                                        If m_objLoopUser.Class < g_objSettings.MinMyinfoFakeCls Then
423:                                                                            m_objLoopUser.SendData curUser.sMyInfoFakeString & "|"
424:                                                                        Else
425:                                                                            m_objLoopUser.SendData strCommand & "|"
426:                                                                        End If
427:                                                                    Else
428:                                                                        m_objLoopUser.SendData strCommand & "|"
429:                                                                    End If
430:                                                                End If
                                                            '#Else
                                                            '    If g_objSettings.HideMyinfos Then
                                                            '        If m_objLoopUser.Class < g_objSettings.MinMyinfoFakeCls Then
                                                            '            m_objLoopUser.SendData curUser.sMyInfoFakeString & "|"
                                                            '        Else
                                                            '            m_objLoopUser.SendData strCommand & "|"
                                                            '        End If
                                                            '    Else
                                                            '        m_objLoopUser.SendData strCommand & "|"
                                                            '    End If
                                                            '#End If
442:                                                        Next
                                                        
444:                                                        Set m_objLoopUser = Nothing
445:                                                    Else
446:                                                        If g_objSettings.HideMyinfos Then
447:                                                            If curUser.Class < g_objSettings.MinMyinfoFakeCls Then
448:                                                                curUser.SendData curUser.sMyInfoFakeString & "|"
449:                                                            Else
450:                                                                curUser.SendData strCommand & "|"
451:                                                            End If
452:                                                        Else
453:                                                            curUser.SendData strCommand & "|"
454:                                                        End If
455:                                                    End If
456:                                                Else
                        'fail hub's rule(s)
458:                                                    Exit Sub
459:                                                End If
460:                                            Else
                        'not same nick in his myinfo
462:                                                wskLoop_Close Index
463:                                                Exit Sub
464:                                            End If
465:                                        End If
466:                                    Else
                                        'If they are not logged in, and support QuickList, they
                                        'must be logging in; if not, then the handshake is almost done
469:                                        If curUser.QuickList Then
                                            'Discontinue processing if they fail nick validation
471:                                            If Not ValidateNick(curUser, MidB$(strParts, 11, InStrB(13, strParts, " ") - 11), strParts) Then _
                                                Exit Sub
473:                                        Else
                                            'Make sure they aren't faking the name
475:                                            If MidB$(strParts, 11, InStrB(13, strParts, " ") - 11) = curUser.sName Then
                                                'Should be waiting for the MyINFO string
477:                                                If curUser.State = Wait_Info Then
                                                    'Check to see if it passes the rules
479:                                                    If ProcessMyINFO(curUser, strParts) Then
480:                                                        g_colUsers.UpdateLogIn curUser
                                                        
                                                        'Call the right sub
                                                        Select Case curUser.Class
                                                            Case Normal: SEvent_UserConnected curUser
                                                            Case Mentored, Registered, Invisible, vip: SEvent_RegConnected curUser
                                                            Case Else
481:                                                            SEvent_OpConnected curUser
                                                                'Show pupop notification ..
482:                                                            If g_objSettings.PopUpOpConected Then _
                                                                    g_objFunctions.ShowBallon "PTDCH  - " & g_objSettings.HubName, "Op Connected" & vbNewLine & "Nick: " & curUser.sName, 0, True
483:                                                        End Select
484:                                                    Else
                        'fail hub's rule(s)
486:                                                        Exit Sub
487:                                                    End If
488:                                                Else
                        'wrong handshake
490:                                                    wskLoop_Close Index
491:                                                    Exit Sub
492:                                                End If
493:                                            Else
                        'fail nick validation
495:                                                wskLoop_Close Index
496:                                                Exit Sub
497:                                            End If
498:                                        End If
499:                                    End If
                                Case 115
                                    'MyPass
                                    '
                                    '   -- Format   : $MyPass <password>|
                                    '   -- Response : $BadPass|  //  $LogedIn <name>|
                                    '
                                    '   -- Standard protocol message (registered users only)
                                    '   -- If the user is registered they send the password for their
                                    '      account. If it's wrong, send $BadPass, otherwise send
                                    '      $LogedIn
                                    
                                    'Make sure the user is supposed to send a password
                                    Select Case curUser.State
                                        Case Wait_Pass
                                            'User is registered
                                            
513:                                            strKey = curUser.sName
                                    
                                            'Check if password is correct
516:                                            lngPos = g_objRegistered.Check(strKey, strParts)
                                            
                                            'If it's a nonzero value, the password is correct
519:                                            If lngPos Then
520:                                                If g_objSettings.PreventGuessPass Then UpdateFailedReg curUser, True

522:                                                curUser.Class = lngPos
                                        
                                                'If there is a logged user already with the same name
                                                'disconnect them
526:                                                If g_colUsers.Online(strKey) = -1 Then _
                                                    wskLoop_Close g_colUsers.ItemByName(strKey).iWinsockIndex
                                        
                                                'We need their MyINFO string before logging them in,
                                                'so unless they are using QuickList, don't log them in
531:                                                If curUser.QuickList Then
532:                                                    If lngPos > vip Then
533:                                                        curUser.SendData "$LogedIn " & strKey & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("LoggedIn") & ".|"
534:                                                    Else
535:                                                        curUser.SendData "<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("LoggedIn") & ".|"
536:                                                    End If
                                                    
                                                    'Check MyINFO
539:                                                    If ProcessMyINFO(curUser, curUser.sMyInfoString) Then
                                                        'Set this to true, because QuickList clients don't send $GetNickList
541:                                                        curUser.QNL = True
                                                
543:                                                        g_colUsers.UpdateLogIn curUser
                                            
                                                        'Raise script event
545:                                                        If lngPos > vip Then
546:                                                            SEvent_OpConnected curUser
                                                                'Show pupop notification ..
547:                                                            If g_objSettings.PopUpOpConected Then _
                                                                    g_objFunctions.ShowBallon "PTDCH  - " & g_objSettings.HubName, "Op Connected" & vbNewLine & "Nick: " & curUser.sName, 0, True
548:                                                        Else
549:                                                            SEvent_RegConnected curUser
550:                                                        End If
551:                                                    Else
552:                                                        Exit Sub
553:                                                    End If
554:                                                Else
555:                                                    If lngPos > vip Then
556:                                                        curUser.SendData "$Hello " & strKey & "|$LogedIn " & strKey & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("LoggedIn") & ".|"
557:                                                    Else
558:                                                        curUser.SendData "$Hello " & strKey & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("LoggedIn") & ".|"
559:                                                    End If
560:                                                    curUser.State = Wait_Info
561:                                                End If
                                        
                                                'Update log in status in database
564:                                                m_objPermaCon.Execute "UPDATE UsrDynamic Set LastLogin=#" & Format$(Now, "yyyy-mm-dd hh:mm:ss") & "#, LastIP=""" & curUser.IP & """ WHERE UserName=""" & curUser.sName & """", , 129
565:                                            Else
566:                                                If g_objSettings.PreventGuessPass Then UpdateFailedReg curUser, False

568:                                                curUser.SendData "<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("WrongPass") & "|$BadPass|"

570:                                                DoEvents
571:                                                wskLoop_Close Index
                                                
573:                                                Exit Sub
574:                                            End If
                                        Case Wait_PassPM
                                            'Not registered, but the hub is running in PM mode
                                            
                                            'Make sure the password is correct
578:                                            If strParts = g_objSettings.HubPassword Then
579:                                                curUser.Class = Normal

581:                                                curUser.SendData "$Hello " & curUser.sName & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("LoggedIn") & "|"
                                                
                                                'We need their MyINFO string before logging them in,
                                                'so unless they are using QuickList, don't log them in
585:                                                If curUser.QuickList Then
586:                                                    g_colUsers.UpdateLogIn curUser
                                            
                                                    'Raise script event
589:                                                    SEvent_UserConnected curUser
590:                                                Else
591:                                                    curUser.State = Wait_Info
592:                                                End If
593:                                            Else
                                                'Send redirect request or just tell them they got it wrong
595:                                                If g_objSettings.RedirectFGP Then
596:                                                    NextRedirect
597:                                                    curUser.SendData "<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("WrongPassRedir") & g_objSettings.RedirectIP & "|$BadPass|$ForceMove " & g_objSettings.RedirectIP & "|"
598:                                                Else
599:                                                    curUser.SendData "<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("WrongPass") & "|$BadPass|"
600:                                                End If
                                                
602:                                                DoEvents
603:                                                wskLoop_Close Index
                                                
605:                                                Exit Sub
606:                                            End If
607:                                    End Select
                                'Case 101
                                '    'MultiConnectToMe
                                'Case 104
                                '    'MultiSearch
                                'Case 80
                                '    'MyIP
                                '
                                '    curUser.SendData "$YourIP " & curUser.IP & "|"
616:                            End Select
                        Case 67 'C
                            'Possible messages :
                            '   -- ConnectToMe
                            '   -- ClientID (ignored)
                            
621:                            If LenB(strKey) = 22 Then
                                'ConnectToMe
                                '
                                '   -- Format   : $ConnectToMe <name> <ip>:<port>|
                                '   -- Response : N/A
                                '
                                '   -- Standard protocol message
                                '   -- Active users send this for <name> to connect to their IP
                                '      on the specified port to intiate a file transfer connection
                                
631:                                If Not g_objSettings.ChatOnly Then
                      
633:                                    If curUser.State = Logged_In Then
                                               'They know why ; )
635:                                               If curUser.ChatOnly Then curUser.Kick g_objSettings.DefaultBanTime: Exit Sub

637:                                        curUser.Passive = False
                                    
                                        'If using the mentoring system, then we may have to disconnect the user
640:                                        If g_objSettings.MentoringSystem Then
                                            'Do not check users who are being mentored
                                            'or are an VIP/Op
643:                                            lngPos = curUser.Class
                                            
                                            Select Case curUser.Class
                                                Case 2, Is < vip
                                                    'Make the min share check
646:                                                    If curUser.iBytesShared < g_objSettings.MinShare Then _
                                                        FailedConf curUser, MinShare: Exit Sub
648:                                            End Select
649:                                        End If
                                        
651:                                        lngPos = InStrB(1, strParts, " ")
652:                                        strKey = LeftB$(strParts, lngPos - 1)
                                        
                                        Select Case True
                                            Case QueuedConnect(strKey & "|" & curUser.sName), curUser.Class >= g_objSettings.MinConnectCls
                                                'Make the MLDC check if needed
655:                                                If g_objSettings.AutoKickMLDC Then
656:                                                    strParts = MidB$(strParts, lngPos + 2)
657:                                                    lngPos = InStrB(1, strParts, ":")
658:                                                    strIP = MidB$(strParts, lngPos + 2)
                                                    
                                                    'Make sure the port is numeric
661:                                                    If IsNumeric(strIP) Then
                                                        'If the client is listening on port 4444 for connections
                                                        'they are a MLDC client
664:                                                        If CLng(strIP) = 4444 Then
665:                                                            curUser.Kick 60
666:                                                            Exit Sub
667:                                                        Else
668:                                                            strParts = LeftB$(strParts, lngPos - 1)
669:                                                        End If
670:                                                    Else
671:                                                        curUser.Kick 60
672:                                                        Exit Sub
673:                                                    End If
674:                                                Else
675:                                                    lngPos = lngPos + 2
676:                                                    strParts = MidB$(strParts, lngPos, InStrB(lngPos, strParts, ":") - lngPos)
677:                                                End If
                                                
679:                                                strIP = curUser.IP
                                                
                                                'Find out if the IP is a local range (replace IP if needed)
                                                ' Range 1: Class A - 10.0.0.0 through 10.255.255.255
                                                ' Range 2: Class B - 172.16.0.0 through 172.31.255.255
                                                ' Range 3: Class C - 192.168.0.0 through 192.168.255.255
                                                Select Case CByte(LeftB$(strIP, InStrB(1, strIP, ".") - 1))
                                                    Case 192, 10, 172
                                                        '$ConnectToMe TheNOP_log 64.228.81.77:3340
                                                        '$ConnectToMe <strKey> <strParts>:<Port>
687:                                                        If g_colUsers.Online(strKey) Then
688:                                                            strObjIP = g_colUsers.ItemByName(strKey).IP
                                                            'Find out if the IP that he want to connect to is also a LAN IP
                                                            Select Case CByte(LeftB$(strObjIP, InStrB(1, strObjIP, ".") - 1))
                                                                Case 192, 10, 172
                                                                    'is also LAN, If IP is not the same, then fix it
691:                                                                    If Not strIP = strParts Then
692:                                                                        strCommand = Replace(strCommand, strParts, strIP, 1, 1)
693:                                                                    End If
694:                                                            End Select
                                                        'Else
                                                            'attempt to get ride of possible ghosts
                                                            'curUser.SendData "$Quit " & strKey & "|"
698:                                                        End If
                                                    Case 127

                                                    Case Else
                                                        'If IP is not the same, then fix it
701:                                                        If Not strParts = strIP Then _
                                                            strCommand = Replace(strCommand, strParts, strIP, 1, 1)
703:                                                End Select
                            
705:                                                If g_colUsers.Online(strKey) Then
                            
                                                '#If FLASHCHAT Then
708:                                                    If g_colUsers.ItemByName(strKey).ChatOnly Then
709:                                                        curUser.SendChat g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("ChatMode"), "%[user]", strKey)
710:                                                    Else
                                                '#End If
                                                
713:                                                        If g_objSettings.MinClsConnectSend Then
714:                                                            g_colUsers.ItemByName(strKey).SendData strCommand & "|"
715:                                                        Else
716:                                                            Set m_objLoopUser = g_colUsers.ItemByName(strKey)
                                                    
718:                                                            If m_objLoopUser.Class >= g_objSettings.MinConnectCls Then
719:                                                                g_colUsers.ItemByName(strKey).SendData strCommand & "|"
720:                                                            End If

722:                                                            Set m_objLoopUser = Nothing
723:                                                        End If
                                                
                                                '#If FLASHCHAT Then
726:                                                    End If
                                                '#End If
                                                'Else
                                                    'attempt to get ride of a possible ghost
                                                    'curUser.SendData "$Quit " & strKey & "|"
731:                                                End If
732:                                        End Select
733:                                    End If
734:                                End If
735:                            End If
                        Case 82 'R
                            'Possible messages :
                            '   -- RevConnectToMe
                            
                            'RevConnectToMe
                            '
                            '   -- Format   : $RevConnectToMe <name> <othername>|
                            '   -- Response : $ConnectToMe <name> <otherip>| (from client)
                            '
                            '   -- Standard protocol message (passive clients)
                            '   -- For passive mode connecting; name is the passive client
                            '      and othername is the client it wants to connect to.
                            '      Assuming the other client is active, it will respond with
                            '      a ConnectToMe message so that the passive user can connect to
                            '      them

751:                            If Not g_objSettings.ChatOnly Then
                                'The user must be logged in to connect to other uses
753:                                If curUser.State = Logged_In Then
                                              'They know why ; )
755:                      If curUser.ChatOnly Then curUser.Kick g_objSettings.DefaultBanTime: Exit Sub

757:                                    If curUser.Class >= g_objSettings.MinConnectCls Then
                                    
                                        'add a Myinfo check here, to see if tag is really showing passive...
760:                                        curUser.Passive = True
                                    
762:                                        lngPos = InStrB(1, strParts, " ")
                                        
                                        'Make sure the user isn't faking their name
765:                                        If LeftB$(strParts, lngPos - 1) = curUser.sName Then
766:                                            strKey = MidB$(strParts, lngPos + 2)
                                            
                                            'If user is online, forward the message to them
769:                                            If g_colUsers.Online(strKey) Then
770:                                                If g_colUsers.ItemByName(strKey).ChatOnly Then
771:                                                        curUser.SendChat g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("ChatMode"), "%[user]", strKey)
772:                                                Else
773:                                                    If g_objSettings.MinClsConnectSend Then
774:                                                        If g_objSettings.MinConnectCls Then
775:                                                            g_colUsers.ItemByName(strKey).SendData strCommand & "|"
                                                            
                                                            'On Error GoTo NextLoop
778:                        On Error Resume Next
                                                            
780:                                                            m_colRevConnects.Add Now, curUser.sName & "|" & strKey

782:                                                        If Err.Number = 457 Then
783:                            Err.Clear
784:                            On Error GoTo Err
785:                            GoTo NextLoop
786:                        End If

788:                                                        End If
789:                                                    Else
790:                                                        Set m_objLoopUser = g_colUsers.ItemByName(strKey)
                                                    
792:                                                        If m_objLoopUser.Class >= g_objSettings.MinConnectCls Then
793:                                                            If g_objSettings.MinConnectCls Then
794:                                                                m_objLoopUser.SendData strData & "|"
                                                                
                                                                'On Error GoTo NextLoop
797:                                  On Error Resume Next

799:                                                                m_colRevConnects.Add Now, curUser.sName & "|" & strKey

801:                                                        If Err.Number = 457 Then
802:                            Err.Clear
803:                            On Error GoTo Err
804:                            GoTo NextLoop
805:                        End If

807:                                                            End If
808:                                                        End If
                                                    
810:                                                        Set m_objLoopUser = Nothing
811:                                                    End If
812:                                                End If
                                            'Else
                                                'attempt to get ride of possible ghosts
                                                'curUser.SendData "$Quit " & strKey & "|"
816:                                            End If
817:                                        End If
818:                                    End If
819:                                End If
820:                            End If
                        Case 71 'G
                            'Possible messages :
                            '   -- GetNickList
                            '   -- GetInfo (ignored)
                            
825:                            If LenB(strKey) = 22 Then
                                'GetNickList
                                '
                                '   -- Format   : $GetNickList|
                                '   -- Response : $NickList <name>$$[<name_etc>$$]|$OpList <name>$$[<name_etc>$$]|
                                '
                                '   -- Standard protocl message
                                '   -- Retrieves the list of users / ops connected to the hub
                                '   -- Traditionally, a GetINFO should be sent to get the MyINFOs
                                '      of the users, but DDCH sends all the MyINFOs with the nicklist
                                '      and ignores GetINFO requests
                                
                                'Delaying the nicklist is is a major bandwidth saving feature
                                'GetINFO, NickList and OpList are not sent until MyINFO is validated (and passes)
                                'Also MyINFO is not sent if it fails the checks (and neither is Hello, therefore not
                                'requiring a Quit message)
                                
                                'Check for flooding
843:                                If g_objSettings.EnableFloodWall Then _
                                    If curUser.FloodCheck(3) Then _
                                        Exit Sub

                                'If user isn't logged in, then queue the nicklist
848:                                If curUser.State = Logged_In Then
                                    'If the user is not visible, then we must add their nickname to the lists
850:                                    If curUser.Visible Then
                                        'Only send the oplist if the user is using QuickList;
                                        'otherwise we send both the nicklist and the oplist
853:                                        If Not curUser.NoHello Then curUser.SendData "$NickList " & g_colUsers.NickList & "|"
854:                                        curUser.SendData "$OpList " & g_colUsers.OpList & "|"
855:                                    Else
856:                                        strKey = curUser.sName
                                        
                                        'Add user's name to nicklist if we're sending
859:                                        If Not curUser.NoHello Then curUser.SendData "$NickList " & g_colUsers.NickList & strKey & "$$|"
                                        
                                        'Add user's name to oplist if they are an operator
862:                                        If curUser.bOperator Then _
                                            curUser.SendData "$OpList " & g_colUsers.OpList & strKey & "$$|" _
                                        Else _
                                            curUser.SendData "$OpList " & g_colUsers.OpList & "|"
                                            
                                        'Send their MyINFO string to themselves (since they are invisible)
868:                                        If g_objSettings.HideMyinfos Then
869:                                            If curUser.Class < g_objSettings.MinMyinfoFakeCls Then
870:                                                curUser.SendData curUser.sMyInfoFakeString & "|"
871:                                            Else
872:                                                curUser.SendData curUser.sMyInfoString & "|"
873:                                            End If
874:                                        Else
875:                                            curUser.SendData curUser.sMyInfoString & "|"
876:                                        End If
877:                                    End If

                                    #If FLASHCHAT Then
                                        'ChatOnly client, should not need to refresh Nicklist
881:                                     If Not curUser.NullCharSeparator Then
                                    #End If

                                    'Build MyINFO stream
885:                                    For Each m_objLoopUser In g_colUsers
886:                                        If m_objLoopUser.Visible Then
887:                                            If g_objSettings.HideMyinfos Then
888:                                                If curUser.Class < g_objSettings.MinMyinfoFakeCls Then
889:                                                    strTMyinfosStr = strTMyinfosStr & m_objLoopUser.sMyInfoFakeString & "|"
890:                                                Else
891:                                                    strTMyinfosStr = strTMyinfosStr & m_objLoopUser.sMyInfoString & "|"
892:                                                End If
893:                                            Else
894:                                                strTMyinfosStr = strTMyinfosStr & m_objLoopUser.sMyInfoString & "|"
895:                                            End If
896:                                        End If
897:                                    Next
                                    'Send MyINFO stream
899:                                    curUser.SendData strTMyinfosStr
                                    'TheNOP End

                                    'Send Bot MyINFOs
903:                                    UpdateBots curUser

                                    #If FLASHCHAT Then
906:                                        End If
                                    #End If
908:                                Else
909:                                    curUser.QNL = True
910:                                End If
911:                            End If
                        Case 84 'T
                            'Possible messages :
                            '   -- To: (Private message)

                            'To:
                            '
                            '   -- Format   : $To: <name> From: <from> $<<from>> <message>|
                            '   -- Response : N/A
                            '
                            '   -- Standard protocol message
                            '   -- Sends a "private" message to another user (the hub owner can
                            '      actually read it so it isn't really private; but hub owners which
                            '      do that are quite lame and I wish I could add idiotic protection to
                            '      DDCH to prevent them from using it *ahem*)

                            'Check if the user is muted
927:                            If Not curUser.Mute Then

                            'Make sure the user isn't flooding
930:                            If g_objSettings.EnableFloodWall Then
                                'don't check >= vips
932:                                If curUser.Class < vip Then
                                    'If = 0 then disable main chat flood check
934:                                    If g_objSettings.FWMainChat Then
935:                                        If curUser.FloodCheck(4) Then Exit Sub
936:                                    End If
937:                                End If
938:                            End If

940:                                strKey = LeftB$(strParts, InStrB(3, strParts, " ") - 1)

                                'If the name is either the bot name or op chat name, take special actions
                                Select Case strKey
                                    Case g_objSettings.BotName
                                        'PMs to the bot normally mean it's a command

945:                                        strKey = MidB$(strParts, InStrB(InStrB(1, strParts, "$"), strParts, " ") + 2)
946:                                        If LenB(strKey) Then _
                                            If AscW(strKey) = g_objSettings.CPrefix Then _
                                                If g_objSettings.EnabledCommands Then ProcessTrigger curUser, MidB$(strKey, 3), False

                                    Case g_objSettings.OpChatName
                                        'Op chat

                                        'Check if only ops or if vips can use the op chat
953:                                        If g_objSettings.VIPUseOpChat Then
954:                                            If curUser.Class < vip Then
955:                                                GoTo NextLoop
956:                                            End If
957:                                        Else
958:                                            If Not curUser.bOperator Then
959:                                                GoTo NextLoop
960:                                            End If
961:                                        End If

'Process commands from OpChat
'warning        strKey    is re-used...., this need changes
'                                        strKey = MidB$(strParts, InStrB(InStrB(1, strParts, "$"), strParts, " ") + 2)

'                                        If LenB(strKey) Then
'                                            If AscW(strKey) = g_objSettings.CPrefix Then
'                                                If g_objSettings.EnabledCommands Then
'                        ProcessTrigger curUser, MidB$(strKey, 3), False
'                        If g_objSettings.FilterCPrefix Then
'                               GoTo NextLoop
'                       End If
'                       End If
'                   End If

977:                                        strKey = " From: " & strKey & " " & MidB$(strParts, InStrB(1, strParts, "$")) & "|"

                                        'Get pointer to curUser object, so that we can make sure
                                        'we don't send this message to the person who sent it
981:                                        lngPos = ObjPtr(curUser)
                                        'TheNOP svn 159
983:                                        If g_objSettings.VIPUseOpChat Then
984:                                            For Each m_objLoopUser In g_colUsers
985:                                                If m_objLoopUser.Class >= vip Then
986:                                                    If Not lngPos = ObjPtr(m_objLoopUser) Then
987:                                                        m_objLoopUser.SendData "$To: " & m_objLoopUser.sName & strKey
988:                                                    End If
989:                                                End If
990:                                            Next
991:                                        Else
992:                                            For Each m_objLoopUser In g_colUsers
993:                                                If m_objLoopUser.Class >= Op Then
994:                                                    If Not lngPos = ObjPtr(m_objLoopUser) Then
995:                                                        m_objLoopUser.SendData "$To: " & m_objLoopUser.sName & strKey
996:                                                    End If
997:                                                End If
998:                                            Next
999:                                        End If

                                    Case Else
                                        'Normal user; check if they are online
                                        'svn 216
                                        'check for nick spoofing
1004:                                        If curUser.sName = g_objRegExps.CaptureSubStr(strCommand, GETFROMNICKINPM) Then
1005:                                            If curUser.sName = g_objRegExps.CaptureSubStr(strCommand, GETNICKINPMMSG) Then
1006:                                                If g_colUsers.Online(strKey) Then g_colUsers.ItemByName(strKey).SendData strCommand & "|"
1007:                                            Else
1008:                                                curUser.Kick 60
1009:                                                Exit Sub
1010:                                            End If
1011:                                        Else
1012:                                            curUser.Kick 60
1013:                                            Exit Sub
1014:                                        End If

1016:                                End Select
1017:                            End If
                        Case 75 'K
                            'Possible messages :
                            '   -- Key
                            '   -- Kick
                            
                            Select Case LenB(strKey)
                                Case 6
                                    'Key
                                    '
                                    '   -- Format   : $Key <string>|
                                    '   -- Response : N/A
                                    '
                                    '   -- Standard protocol message (ignored)
                                    '   -- Sent in response to $Lock
                                    '   -- Originally it was a security check
                                    '      to prevent unauthorized DC clients from
                                    '      connecting to the hub; now it's just useless
                                    
                                    'If Not strParts = vbKey Then wskLoop_Close Index: Exit Sub
1034:                                    curUser.State = Wait_Validate
                                Case 8
                                    'Kick
                                    '
                                    '   -- Format   : $Kick <name>|
                                    '   -- Response : N/A
                                    '
                                    '   -- Standard protocol message (for ops only)
                                    '   -- Disconnects and temp bans the user using <name>
                                    
1043:                                    If curUser.bOperator Then
1044:                                        If g_colUsers.Online(strParts) Then
1045:                                            Set m_objLoopUser = g_colUsers.ItemByName(strParts)
                                            
                                            'Cannot kick user above or equal to own class (unless the user is an admin)
                                            Select Case curUser.Class
                                                Case Admin, InvisibleAdmin, Is > m_objLoopUser.Class
                                                    'Get IP because we need it for message
1049:                                                    strKey = m_objLoopUser.IP
1050:                                                    DoEvents
1051:                                                    m_objLoopUser.Kick
1052:                                                    g_colUsers.SendChatToOps g_objSettings.BotName, Replace(Replace(Replace(g_objFunctions.GetENLangStr("KickedBy"), "%[user]", strParts), "%[op]", curUser.sName), "%[ip]", strKey)
1053:                                            End Select
                                            
1055:                                            Set m_objLoopUser = Nothing
1056:                                        End If
1057:                                    End If
1058:                            End Select
                        Case 86 'V
                            'Possible messages :
                            '   -- ValidateNick
                            '   -- Version
                            
                            Select Case LenB(strKey)
                                Case 24
                                    'ValidateNick
                                    '
                                    '   -- Format   : $ValidateNick <name>|
                                    '   -- Response : $ValidateDenide <name>|  //  $GetPass|  //  $Hello <name>|
                                    '
                                    '   -- Standard protocol message
                                    '   -- Sent in response to intial message $Lock
                                    '   -- Used to see if <name> is used; if it is, then send
                                    '      $ValidateDenide; if it's not taken the send $GetPass
                                    '      if it's registered, or send $Hello if not
                                    
                                    'Rout to another sub (if it returns false, not check any more data)
1075:                                    If Not ValidateNick(curUser, strParts) Then _
                                        Exit Sub
                                Case 14
                                    'Version
                                    '
                                    '   -- Format   : $Version <version>|
                                    '   -- Response : N/A
                                    '
                                    '   -- Standard protocol message (optional)
                                    '   -- <version> contains (in NMDC) the clients version
                                    '   -- Can be overridden in all other clients (virtually all)
                                    
                                    'Convert to proper decimal
1087:                                    If m_blnCommaDecimal Then
1088:                                        curUser.iVersion = StrToDbl(Replace(strParts, ".", ","))
1089:                                    Else
1090:                                        curUser.iVersion = Val(strParts)
1091:                                    End If
                                    
                                    Select Case False
                                        Case g_objSettings.OPBypass, curUser.Class >= vip
                                            'NMDC min version check
1094:                                            If g_objSettings.NMDCMinVersion Then _
                                                If g_objSettings.NMDCMinVersion > curUser.iVersion Then _
                                                    FailedConf curUser, NMDCVersion: Exit Sub
1097:                                    End Select
1098:                            End Select
                        Case 79 'O
                            'Possible messages :
                            '   -- OpForceMove (Redirect)
                            
                            'OpForceMove
                            '
                            '   -- Format   : $OpForceMove $Who:<name>$Where:<address>$Msg:<message>|
                            '   -- Response : N/A
                            '
                            '   -- Standard protocol message (for ops only)
                            '   -- Redirects a user to <address> (private messages them <message>)
                            '   -- A point worthy of mention is that the client can choose
                            '      to ignore the message; so make sure they are disconnected
                            
1112:                            If curUser.bOperator Then
                                'Check if the user can redirect
1114:                                If Not g_objSettings.OpsCanRedirect Then _
                                    If curUser.Class < Admin Then _
                                        GoTo NextLoop

1118:                                strParts = MidB$(strParts, 11)
1119:                                lngPos = InStrB(1, strParts, "$")
1120:                                strKey = LeftB$(strParts, lngPos - 1)
                                
                                'Check if user is online
1123:                                If g_colUsers.Online(strKey) Then
1124:                                    Set m_objLoopUser = g_colUsers.ItemByName(strKey)
                                    
                                    'Cannot redirect user above or equal to own class (unless the user is an admin)
                                    Select Case curUser.Class
                                        Case Admin, InvisibleAdmin, Is > m_objLoopUser.Class
1127:                                            strParts = MidB$(strParts, lngPos + 14)
1128:                                            lngPos = InStrB(1, strParts, "$")
                                    
                                            'Private message user reason, then redirect
1131:                                            m_objLoopUser.SendPrivate curUser.sName, MidB$(strParts, lngPos + 10)
1132:                                            m_objLoopUser.Redirect LeftB$(strParts, lngPos - 1)
1133:                                    End Select
                                    
1135:                                    Set m_objLoopUser = Nothing
1136:                                End If
1137:                            End If
                        Case 78 'N
                            'Possible messages :
                            '   -- NetINFO
                            
                            'NetINFO
                            '
                            '   -- Format   : $NetINFO <slots>$<hubs>$<mode>|
                            '                 $NetINFO <slots>$<hubs>$<mode>$<bandwidth>|
                            '      Response : N/A
                            '
                            '   -- Protocol Extension
                            '   -- Supported by NMDC only (as a substitute for tags)
                            '   -- This is sent in response to a $GetNetInfo|
                            '   -- Two different forms; version 2.02 of DC includes the upload
                            '      bandwidth limit value
                            
1153:                            curUser.NetInfo = True
                            
                            'Skip ops if necessary
                            Select Case False
                                Case g_objSettings.OPBypass, curUser.Class >= vip
                                    'Extract slots, hubs and bandwidth if provided
1157:                                    lngPos = InStrB(1, strParts, "$")
1158:                                    bytesTotal = CLng(LeftB$(strParts, lngPos - 1))
1159:                                    strParts = MidB$(strParts, lngPos + 2)
                                    
1161:                                    lngPos = InStrB(1, strParts, "$")
1162:                                    strKey = LeftB$(strParts, lngPos - 1)
1163:                                    strParts = MidB$(strParts, lngPos + 2)
1164:                                    lngPos = CLng(strKey)
                                    
                                    'Find out if we need to get upload bandwidth
1167:                                    If Not LenB(strParts) = 2 Then _
                                        strParts = MidB$(strParts, 5) _
                                    Else _
                                        strParts = "0"
                                
                                    'Max hubs
1173:                                    If g_objSettings.DCMaxHubs Then _
                                        If lngPos > g_objSettings.DCMaxHubs Then _
                                            FailedConf curUser, MaxHubs: Exit Sub
                                    
                                    'Min slots
1178:                                    If g_objSettings.MinSlots Then _
                                        If bytesTotal < g_objSettings.MinSlots Then _
                                            FailedConf curUser, MinSlots: Exit Sub
                                            
                                    'Max slots
1183:                                    If g_objSettings.MaxSlots Then _
                                        If bytesTotal > g_objSettings.MaxSlots Then _
                                            FailedConf curUser, MaxSlots: Exit Sub
                                    
                                    'Hub/Slot ratio
1188:                                    If g_objSettings.DCSlotsPerHub Then _
                                        If (bytesTotal / lngPos) < g_objSettings.DCSlotsPerHub Then _
                                            FailedConf curUser, HSRatio: Exit Sub
                                    
                                    'Bandwidth/Slot ratio
1193:                                    If g_objSettings.DCBandPerSlot Then _
                                        If Not strParts = "0" Then _
                                            If (CLng(strParts) / bytesTotal) < g_objSettings.DCBandPerSlot Then _
                                                FailedConf curUser, BSRatio: Exit Sub
1197:                            End Select
                        Case 85 'U
                            'Possible messages :
                            '   -- UserIP
                            
                            'UserIP
                            '   -- Format   : $UserIP <name>[$$<name_etc>]|
                            '      Response : $UserIP <name> <ip>[$$<name_etc> <ip_etc>]|
                            '
                            '   -- Protocol Extension
                            '   -- Ops can get anyone's IP, while users can
                            '      only get their own
                                
1209:                            If curUser.bOperator Then
1210:                                strKey = "$UserIP "
1211:                                strParts = strParts & "$$"
1212:                                lngPos = InStrB(1, strParts, "$$")

                                'Loop to find all ip requests
1215:                                Do While lngPos
1216:                                    strIP = LeftB$(strParts, lngPos - 1)
1217:                                    strParts = MidB$(strParts, lngPos + 4)
                                        
                                    'If online add their ip to the list, otherwise add blank
1220:                                    If g_colUsers.Online(strIP) Then _
                                        strKey = strKey & strIP & " " & g_colUsers.ItemByName(strIP).IP & "$$" _
                                    Else _
                                        strKey = strKey & strIP & "  $$"
                                        
1225:                                    lngPos = InStrB(1, strParts, "$$")
1226:                                Loop
                                
                                'Remove last $$
1229:                                curUser.SendData LeftB$(strKey, LenB(strKey) - 4) & "|"
1230:                            Else
1231:                                curUser.SendData "$UserIP " & curUser.sName & " " & curUser.IP & "|"
1232:                            End If
                        Case 120 'x
                            'Possible messages :
                            '   -- xKick
                            
                            'xKick
                            '   -- Format   : $xKick <name>$<length>$<show>$<msg>|
                            '      Response : N/A
                            '
                            '   -- Protocol Extension
                            '   -- For operators only
                            '   -- Allows op to choose how long the user will be
                            '      banned for as well as whether or not the kick
                            '      will be seen the in the main chat
                            '   -- xKick also contains the kick reason
                            
1247:                            If curUser.bOperator Then
1248:                                lngPos = InStrB(1, strParts, "$")
1249:                                strKey = LeftB$(strParts, lngPos - 1)
                                
                                'Make sure user to kick is online
1252:                                If g_colUsers.Online(strKey) Then
1253:                                    Set m_objLoopUser = g_colUsers.ItemByName(strKey)
                                    
                                    'Cannot kick user above or equal to own class (unless the user is an admin)
                                    Select Case curUser.Class
                                        Case Admin, InvisibleAdmin, Is > m_objLoopUser.Class
1256:                                            strParts = MidB$(strParts, lngPos + 2)
1257:                                            lngPos = InStrB(1, strParts, "$")
                                    
                                            'Ban the IP for the correct length
1260:                                            g_objIPBans.Add m_objLoopUser.IP, CLng(LeftB$(strParts, lngPos - 1))
                                    
1262:                                            strParts = MidB$(strParts, lngPos + 2)
                                    
                                            'Check to see if it should be sent to the main chat
1265:                                            If AscW(strParts) = 84 Then
1266:                                                strParts = MidB$(strParts, InStrB(1, strParts, "$") + 2)
                                    
1268:                                                g_colUsers.SendChatToAll curUser.sName, Replace(Replace(Replace(g_objFunctions.GetENLangStr("IsKicking"), "%[op]", curUser.sName), "%[user]", strKey), "%[reason]", strParts)
                                        
                                                'Send private message and disconnect user
1271:                                                m_objLoopUser.SendPrivate curUser.sName, curUser.GetCoreMsgStr("KickedBecause") & strParts
1272:                                                DoEvents
1273:                                                wskLoop_Close Index
                                        
                                                'Notify that the user has been disconnected
1276:                                                g_colUsers.SendChatToAll g_objSettings.BotName, Replace(Replace(Replace(g_objFunctions.GetENLangStr("KickedBy"), "%[user]", strKey), "%[op]", curUser.sName), "%[ip]", m_objLoopUser.IP)
1277:                                            Else
1278:                                                strParts = MidB$(strParts, InStrB(1, strParts, "$") + 2)
1279:                                                g_colUsers.SendChatToOps curUser.sName, Replace(Replace(Replace(g_objFunctions.GetENLangStr("IsKicking"), "%[op]", curUser.sName), "%[user]", strKey), "%[reason]", strParts)
                                        
                                                'Send private message and disconnect user
1282:                                                m_objLoopUser.SendPrivate curUser.sName, curUser.GetCoreMsgStr("KickedBecause") & strParts
1283:                                                DoEvents
1284:                                                wskLoop_Close Index
                                        
                                                'Notify that the user has been disconnected
1287:                                                g_colUsers.SendChatToOps g_objSettings.BotName, Replace(Replace(Replace(g_objFunctions.GetENLangStr("KickedBy"), "%[user]", strKey), "%[op]", curUser.sName), "%[ip]", m_objLoopUser.IP)
1288:                                            End If
1289:                                    End Select
                                    
1291:                                    Set m_objLoopUser = Nothing
1292:                                End If
1293:                            End If
                        Case 66 'B
                            'Possible messages :
                            '   -- BotINFO
                            '   -- BlackDC (ignored)
                            
1298:                            If AscW(RightB$(strKey, 2)) = 79 Then 'O
                                'BotINFO
                                '   -- Format   : $BotINFO|
                                '      Response : $HubINFO <name>$<ip>$<port>$<description>$<maxusers>$<minshare>$<minslots>$<maxhubs>$<extra>$|
                                '
                                '   -- Protocol Extension
                                '   -- Used by Gadget's hub pinger for www.hublist.org
                                '   -- Gives various extra information on hub such as
                                '      the min share/slots, max hubs, etc
                                
1308:                                curUser.SendData "$HubINFO " & g_objSettings.HubName & "$" & g_objSettings.HubIP & ":" & g_objSettings.Port & _
                                                 "$" & g_objSettings.HubDesc & "$" & g_objSettings.MaxUsers & "$" & g_objSettings.MinShare & _
                                                 "$" & g_objSettings.MinSlots & "$" & g_objSettings.DCMaxHubs & "$DDCH " & vbVersion & " Built-In$|"
1311:                            End If
                        Case 122 'z
                            'Possible messages :
                            '  -- zSearch
                            
                            
                        'Case Else
                            'Perhaps one day I'll do something with unknown messages
1318:                    End Select
                Case 60 '<
                    'Main chat message
                    '
                    '   -- Format   : <<name>> <message>|
                    '   -- Response : N/A
                    '
                    '   -- Standard protocol message
                    '   -- Sends a main chat message to all users

                    'Make sure the user is logged in
1328:                    If curUser.State = Logged_In Then
                        'Check if the user is muted
1330:                        If Not curUser.Mute Then
                                                            
                            ' TheNOP svn 159
                            'Make sure the user isn't flooding
1334:                            If g_objSettings.EnableFloodWall Then
                                'If = 0 then disable PM flood checking
1336:                                If g_objSettings.FWMainChat Then
                                    'don't check >= vips, kick raw can occure fast ;)
1338:                                    If curUser.Class < vip Then
1339:                                        If curUser.FloodCheck(4) Then Exit Sub
1340:                                    End If
1341:                                End If
1342:                            End If

                            'Truncate message if necessary
1345:                            If g_objSettings.MaxMessageLen Then
1346:                                If curUser.Class < vip Then
1347:                                    lngPos = Len(strCommand)
                                    'svn 216  getlang...
1349:                                    If lngPos > g_objSettings.MaxMessageLen Then
1350:                                            curUser.SendChat g_objSettings.BotName, "your message was to big to be sent to other users"
                                    'If lngPos > g_objSettings.MaxMessageLen Then strCommand = Left$(strCommand, g_objSettings.MaxMessageLen)
1352:                                            Exit Sub
1353:                                        End If
1354:                                End If
1355:                            End If
                        
1357:                            lngPos = InStrB(1, strCommand, "> ")
                        
                            'Make sure the client isn't trying to fake it's username
                            'If so, then replace the fake name with the real one
1361:                            If Not MidB$(strCommand, 3, lngPos - 3) = curUser.sName Then
                                'Replace
                                'strCommand = "<" & curUser.sName & "> " & MidB$(strCommand, lngPos + 4)
                                'lngPos = InStrB(1, strCommand, " ")
                                'svn 216
                                'Kick without message, they know why they are kicked...
1367:                                curUser.Kick 60
1368:                                Exit Sub
1369:                            End If
                        
                            'Get the message
1372:                            strKey = MidB$(strCommand, lngPos + 4)
                        
                            'Don't send if there is no string
1375:                            If LenB(strKey) Then
                                'If first character is the command prefix, then process it
1377:                                If AscW(strKey) = g_objSettings.CPrefix Then
1378:                                    If g_objSettings.EnabledCommands Then ProcessTrigger curUser, MidB$(strKey, 3), True
1379:                                    If g_objSettings.FilterCPrefix Then lngPos = 0 Else lngPos = -1
1380:                                End If
                            
                                'Send out main chat message
1383:                                If lngPos Then _
                                    If g_objSettings.SendMessageAFK Then _
                                        g_colUsers.SendToAll strCommand & "|" _
                                    Else _
                                        g_colUsers.SendToNA strCommand & "|"
1388:                            End If
1389:                        End If
1390:                    End If
1391:            End Select
            
            'Call dataarrival event if necessary
            #If DataArrival Then
                'SEvent_DataArrival curUser, strCommand
                
                'This runs the DataArrival event
                '
                '  -- Parameters : curUser (the current user's clsUser object)
                '                : strData (data that was sent)
                '  -- Format     : Sub DataArrival(curUser, strData)
                '
                '  -- Called when a user sends data to the hub
                '  -- Difference with NMDCH is DDCH sends ALL data to the script, while
                '     NMDCH does it selectively (nothing before and including ValidateNick)
    
1407:                On Error Resume Next
    
1409:                For lngPos = 1 To m_lngScriptEventsUB
1410:                    If m_arrScriptEvents(lngPos, vbSDataArrival) Then _
                        ScriptControl(lngPos).Run "DataArrival", curUser, strCommand
1412:                Next
                
1414:                On Error GoTo Err
            #End If
            
1417:        End If
    
1419:
NextLoop:
        'Find next pipe to parse the next message
1421:        lngPos = InStrB(1, strData, "|")
1422:    Loop

    'If there is any data left over, put the fragment into user's data var
1425:    If LenB(strData) Then
1426:        If curUser.bOperator Then
1427:            curUser.DataFragment = strData
1428:        Else
1429:            If LenB(curUser.DataFragment) > (g_objSettings.DataFragmentLen * 2) Then
1430:                curUser.DataFragment = vbNullString
1431:            Else
1432:                curUser.DataFragment = strData
1433:            End If
1434:        End If
1435:    End If

1437:    Exit Sub

1439:
LoopErr:
    'Error occured when trying to parse message
1441:    HandleError Err.Number, Err.Description, Erl & "|" & "wskLoop_DataArrival() (Loop - strCommand = """ & strCommand & """; strData = """ & strData & """)"

1443:    Exit Sub
1444:
Err:
    'Error occured before parsing occured
1446:    HandleError Err.Number, Err.Description, Erl & "|" & "wskLoop_DataArrival() (Preloop - " & strData & ")"
End Sub
Private Sub wskLoop_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
1:    wskLoop_Close Index
End Sub
Private Sub wskRegister_Close(Index As Integer)
    'Just close it; don't really care about the errors here
2:    wskRegister(Index).Close
End Sub
Private Sub wskRegister_DataArrival(Index As Integer, ByVal bytesTotal As Long)
1:    Dim strData     As String
2:    Dim strCommand  As String
3:    Dim strLock     As String
4:    Dim lngPos      As Long
    
6:    On Error Resume Next
    'Get data
8:    wskRegister(Index).GetData strData, vbString

    #If SVN Then
11:        g_objFileAccess.AppendFile G_LOGPATH, Now & " <-- " & wskRegister(Index).RemoteHostIP & " - " & strData
    #End If
13:    On Error GoTo Err
    
15:    lngPos = InStrB(1, strData, "|")
16:    Do While lngPos
17:        strCommand = LeftB$(strData, lngPos - 1)
18:        strData = MidB$(strData, lngPos + 2)
    
        'Possible messages :
        '   -- Lock
    
23:        If LeftB$(strCommand, 10) = "$Lock" Then
            'Lock
            '
            '   -- Format   : $Lock <string> pk=<astring>|
            '   -- Response : $Key <string>|<name>|<ip>[:<port>]|<description>|<users>|<bytes>|
            '
            '   -- Standard protocol message
            '   -- The <string> from Lock is decoded into the <string> for key;
            '      The information which follows is various details about the hub
            '      like the hub name, address, description, users, and total shared
            '      bytes
        
35:            lngPos = wskRegister(Index).LocalPort
        
37:            strLock = "$Key " & LockToKey(MidB$(strCommand, 13, LenB(strCommand) - 14), _
                                            ((lngPos \ 256) + (lngPos And 255)) And 255) _
                                 & "|"
                             
            'Add char160 to the end of the hub name to prevent MoGLO, MoSearch and GLOSearch
42:            If g_objSettings.PreventSearchBots Then _
                strLock = strLock & g_objSettings.HubName & vbChar160 & "|" _
            Else _
                strLock = strLock & g_objSettings.HubName & "|"
                
            'If the port is 411, then don't add it to the address
48:            If g_objSettings.Port = 411 Then _
                strLock = strLock & g_objSettings.HubIP & "|" _
            Else _
                strLock = strLock & g_objSettings.HubIP & ":" & g_objSettings.Port & "|"
                
            'Add char160 to the end of the hub description to prevent MoGLO, MoSearch and GLOSearch
54:            If g_objSettings.PreventSearchBots Then _
                strLock = strLock & g_objSettings.HubDesc & vbChar160 & "|" _
            Else _
                strLock = strLock & g_objSettings.HubDesc & "|"
                
            'Add user count and total bytes
60:            strLock = strLock & g_colUsers.count & "|" & g_colUsers.iTotalBytesShared & "|"
            
62:            On Error Resume Next
            
            #If SVN Then
65:                g_objFileAccess.AppendFile G_LOGPATH, Now & " --> " & wskRegister(Index).RemoteHostIP & " - " & strLock
            #End If
            
68:            On Error GoTo Err
            'Submit registration
70:            wskRegister(Index).SendData strLock
71:        End If
        
73:        lngPos = InStrB(1, strData, "|")
74:    Loop
    
    'The auto disconnect has been removed due to behaviour with NMDCH
    'After researching it with my own registration server, I found it doesn't
    'disconnect, but rather it is the server which ends the connection
  
    'This could have had an impact with registering with vandel405.dynip.com
    'but now I believe however it perfectly emulates NMDCH behaviour in this respect
  
    'DoEvents
    'wskRegister(Index).Close
    
86:    Exit Sub
    
88:
Err:
89:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.wskRegister_DataArrival()"
End Sub
Private Sub wskRegister_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Just close it; don't really care about the errors here
2:    On Error Resume Next
    
    #If SVN Then
5:        g_objFileAccess.AppendFile G_LOGPATH, "wskRegister_Error: " & Description & " | Scode: " & Scode & " | Index: " & Index
    #End If
7:    wskRegister(Index).Close
End Sub
'------------------------------------------------------------------------------
'Winsock events
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' Core related private/public methods
'------------------------------------------------------------------------------
Public Sub SwitchServing()
1:    Dim lngLoop         As Long
2:    Dim lngUB           As Long
3:    Dim lngPos          As Long
4:    Dim arrTemp()       As String
    
6:    On Error GoTo Err
    
    'Find out which state we're in
9:    If m_blnServing Then
        'Stop serving
11:        m_blnServing = False
        
        'GUI related
' ------------------------ NEW INTERFACE LANGUAGE ------------------------
15:        cmdButton(1).Caption = m_arrDynaCap(0) '"Start Serving"
' ----------------------- NEW INTERFACE LANGUAGE END ----------------------
17:        txtData(3).Enabled = True
18:        txtData(4).Enabled = True
19:        txtData(7).Enabled = True
20:        txtData(8).Enabled = True
21:        chkData(21).Enabled = True
22:        chkData(18).Enabled = True

23:        cmdButton(7).Enabled = False
24:        cmdButton(8).Enabled = False
25:        cmdButton(9).Enabled = False
26:        cmdButton(2).Enabled = False
27:        tmrBackground.Enabled = False
        
        'Clear out listening winsocks
28:        wskListen(0).Close
29:        lngUB = wskListen.UBound
        
30:        If lngUB Then
31:            For lngLoop = 1 To lngUB
32:                wskListen(lngLoop).Close
33:                Unload wskListen(lngLoop)
34:            Next
35:        End If
        
        'Clear out user winsocks
38:        If wskLoop(0).State Then wskLoop(0).Close
39:        lngUB = wskLoop.UBound
        
41:        If lngUB Then
42:            For lngLoop = 1 To lngUB
43:                If wskLoop(lngLoop).State Then wskLoop(lngLoop).Close
44:                Unload wskLoop(lngLoop)
45:            Next
46:        End If
        
        'Clear out registration winsocks
49:        If wskRegister(0).State Then wskRegister(0).Close
50:        lngUB = wskRegister.UBound
        
52:        If lngUB Then
53:            For lngLoop = 1 To lngUB
54:                If wskRegister(lngLoop).State Then wskRegister(lngLoop).Close
55:                Unload wskRegister(lngLoop)
56:            Next
57:        End If
        
        'Clear out  collections
        #If COLFREESOCKS Then
61:            Set m_colFreeSocks = Nothing
        #End If
    
64:        g_colUsers.Clear
        
        'Remove bot names, if used
67:        If g_objSettings.UseBotName Then UnregisterBotName g_objSettings.BotName
68:        If g_objSettings.UseOpChat Then UnregisterBotName g_objSettings.OpChatName
               
        'Set serving date
71:        m_datServingDate = Now

        'Raise event
74:        SEvent_StoppedServing
          
        'Show Ballon notification
77:        If g_objSettings.PopUpStopedServing Then g_objFunctions.ShowBallon "PTDCH - " & g_objSettings.HubName, "Server Stoped", 0, True
78:        AddLog "Server stoped.", 1
          
79:    Else

        'Show Ballon notification
82:        If g_objSettings.PopUpStartedServing Then g_objFunctions.ShowBallon "PTDCH - " & g_objSettings.HubName, "Server Started", 0, True
83:        AddLog "Server started.", 1
84:        AddLog "Listening ports: " & g_objSettings.Ports, 1
        'Start serving
86:        m_blnServing = True
        'GUI related
' ------------------------ NEW INTERFACE LANGUAGE ------------------------
89:        cmdButton(1).Caption = m_arrDynaCap(1)  '"Stop Serving"
' ----------------------- NEW INTERFACE LANGUAGE END ----------------------
91:        txtData(3).Enabled = False
92:        txtData(4).Enabled = False
93:        txtData(7).Enabled = False
94:        txtData(8).Enabled = False
95:        chkData(18).Enabled = False
96:        chkData(21).Enabled = False

98:        cmdButton(7).Enabled = True
99:        cmdButton(8).Enabled = True
100:       cmdButton(9).Enabled = True
101:       cmdButton(2).Enabled = True

        'Create objects
        #If COLFREESOCKS Then
104            Set m_colFreeSocks = New Collection
        #End If
        
107:        tmrBackground.Enabled = True
        
        'Get listening ports
110:        arrTemp = Split(g_objSettings.Ports, ";")
111:        lngUB = UBound(arrTemp)

113:        For lngLoop = 0 To lngUB
114:            If IsNumeric(arrTemp(lngLoop)) Then
                'Load winsock as necessary
116:                If lngLoop Then _
                    Load wskListen(lngLoop) _
                Else _
                    g_objSettings.Port = CInt(arrTemp(0))
                    
                'Set port and listen
122:                wskListen(lngLoop).LocalPort = CLng(arrTemp(lngLoop))
                
124:                On Error Resume Next
125:                wskListen(lngLoop).Listen
126:                On Error GoTo Err
                
                'Check if there was an error
129:                If Err.Number = 10048 Then
130:                    MsgBoxCenter Me, Replace(g_colMessages.Item("msgPortInUse"), "%[port]", arrTemp(lngLoop)), vbCritical, g_colMessages.Item("msgStartServing")
131:                    Err.Clear
132:                End If
133:            End If
134:        Next
        
        'Get registration servers
137:        If LenB(g_objSettings.RegisterIP) Then
138:            arrTemp = Split(g_objSettings.RegisterIP, ";")
139:            lngUB = UBound(arrTemp)
        
141:            For lngLoop = 0 To lngUB
                'Load winsock
143:                If lngLoop Then Load wskRegister(lngLoop)
            
                'Get port or set to default 2501
146:                lngPos = InStrB(1, arrTemp(lngLoop), ":")
147:                If lngPos Then
148:                    wskRegister(lngLoop).RemoteHost = LeftB$(arrTemp(lngLoop), lngPos - 1)
149:                    arrTemp(lngLoop) = MidB$(arrTemp(lngLoop), lngPos + 2)
                    
                    'If not numeric, set to default
153:                    If IsNumeric(arrTemp(lngLoop)) Then _
                        wskRegister(lngLoop).RemotePort = CLng(arrTemp(lngLoop)) _
                    Else
154:                        wskRegister(lngLoop).RemotePort = 2501
155:                Else
156:                    wskRegister(lngLoop).RemoteHost = arrTemp(lngLoop)
157:                    wskRegister(lngLoop).RemotePort = 2501
158:                End If
159:            Next
160:        End If
        
        'Preload winsocks if necessary
163:        If g_objSettings.PreloadWinsocks Then
164:            lngUB = g_objSettings.MaxUsers + (((g_objSettings.MaxUsers \ 100) + 1) * 5)
            
166:            For lngLoop = 1 To lngUB
167:                Load wskLoop(lngLoop)
                
                #If COLFREESOCKS Then
170:                    m_colFreeSocks.Add wskLoop(lngLoop), CStr(lngLoop)
                #End If
172:            Next
173:        End If
        
        'Add bot names which were registered before serving was started
176:        If Not m_lngBotsUB = -1 Then
177:            For lngLoop = 0 To m_lngBotsUB
178:                g_colUsers.AppendNL m_arrBots(lngLoop).Name, m_arrBots(lngLoop).Operator
179:            Next
180:        End If
        
        'Register bot names
183:        If g_objSettings.UseBotName Then RegisterBotName g_objSettings.BotName
184:        If g_objSettings.UseOpChat Then RegisterBotName g_objSettings.OpChatName
        
        
        'Set serving date
188:        m_datServingDate = Now
        
        'Raise event
191:        SEvent_StartedServing
192:    End If

        'Clear user listview on start serving
        #If Status Then
196:            g_objStatus.UClear
        #End If
    
199:    Exit Sub
    
201:
Err:
203:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.SwitchServing()"
204:    Resume Next
End Sub

Private Sub FailedConf(ByRef curUser As clsUser, ByRef intType As enuAlert)
1:    Dim strMessage As String

3:    On Error GoTo Err

5:   If SEvent_FailedConf(curUser, intType) Then Exit Sub

    'Find out which message to send
    Select Case intType
'----------ROLL---USERS---REDIRECT--TO--RIGHT--ADDRESS-----------------
'-------------------------------
        Case MaxHubs
'-------------------------------
            'Redirect if necessary
12:            If g_objSettings.RedirectFMaxHubs Then
13:                NextRedirect
            
                'Send message as either private or in the main chat
16:                If g_objSettings.SendMsgAsPrivate Then
17:                    curUser.SendPrivate g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("MaxHubs"), "%[maxhubs]", g_objSettings.DCMaxHubs, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForMaxHubsRedirectAddress & "|$ForceMove " & g_objSettings.ForMaxHubsRedirectAddress
18:                Else
19:                    curUser.SendChat g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("MaxHubs"), "%[maxhubs]", g_objSettings.DCMaxHubs, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForMaxHubsRedirectAddress & "|$ForceMove " & g_objSettings.ForMaxHubsRedirectAddress
20:                End If
21:            Else
22:                strMessage = Replace(curUser.GetCoreMsgStr("MaxHubs"), "%[maxhubs]", g_objSettings.DCMaxHubs, 1, 1)
23:            End If
'-------------------------------
        Case MinSlots
'-------------------------------
            'Redirect if necessary
27:            If g_objSettings.RedirectFMinSlots Then
28:                NextRedirect
                'Send message as either private or in the main chat
30:                If g_objSettings.SendMsgAsPrivate Then
31:                    curUser.SendPrivate g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("MinSlots"), "%[minslots]", g_objSettings.MinSlots, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForMinSlotsRedirectAddress & "|$ForceMove " & g_objSettings.ForMinSlotsRedirectAddress
32:                Else
33:                    curUser.SendChat g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("MinSlots"), "%[minslots]", g_objSettings.MinSlots, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForMinSlotsRedirectAddress & "|$ForceMove " & g_objSettings.ForMinSlotsRedirectAddress
34:                End If
35:            Else
36:                strMessage = Replace(curUser.GetCoreMsgStr("MinSlots"), "%[minslots]", g_objSettings.MinSlots, 1, 1)
37:            End If
'-------------------------------
        Case MaxSlots
'-------------------------------
            'Redirect if necessary
41:            If g_objSettings.RedirectFMaxSlots Then
42:                NextRedirect
                'Send message as either private or in the main chat
44:                If g_objSettings.SendMsgAsPrivate Then
45:                    curUser.SendPrivate g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("MaxSlots"), "%[maxslots]", g_objSettings.MaxSlots, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForMaxSlotsRedirectAddress & "|$ForceMove " & g_objSettings.ForMaxSlotsRedirectAddress
46:                Else
47:                    curUser.SendChat g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("MaxSlots"), "%[maxslots]", g_objSettings.MaxSlots, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForMaxSlotsRedirectAddress & "|$ForceMove " & g_objSettings.ForMaxSlotsRedirectAddress
48:                End If
49:            Else
50:                strMessage = Replace(curUser.GetCoreMsgStr("MaxSlots"), "%[maxslots]", g_objSettings.MaxSlots, 1, 1)
51:            End If
'-------------------------------
        Case NMDCVersion
'-------------------------------
            'Redirect if necessary
55:            If g_objSettings.RedirectFTooOldNMDC Then
56:                NextRedirect
                'Send message as either private or in the main chat
58:                If g_objSettings.SendMsgAsPrivate Then
59:                    curUser.SendPrivate g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("NMDCMinVersion"), "%[minversion]", g_objSettings.NMDCMinVersion, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForTooOldNMDCRedirectAddress & "|$ForceMove " & g_objSettings.ForTooOldNMDCRedirectAddress
60:                Else
61:                    curUser.SendChat g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("NMDCMinVersion"), "%[minversion]", g_objSettings.NMDCMinVersion, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForTooOldNMDCRedirectAddress & "|$ForceMove " & g_objSettings.ForTooOldNMDCRedirectAddress
62:                End If
63:            Else
64:                strMessage = Replace(curUser.GetCoreMsgStr("NMDCMinVersion"), "%[minversion]", g_objSettings.NMDCMinVersion, 1, 1)
65:            End If
'-------------------------------
        Case DCppversion
'-------------------------------
            'Redirect if necessary
69:            If g_objSettings.RedirectFTooOldDCpp Then
70:                NextRedirect
                'Send message as either private or in the main chat
72:                If g_objSettings.SendMsgAsPrivate Then
73:                    curUser.SendPrivate g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("DCppMinVersion"), "%[minversion]", g_objSettings.DCMinVersion, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForTooOldDcppRedirectAddress & "|$ForceMove " & g_objSettings.ForTooOldDcppRedirectAddress
74:                Else
75:                    curUser.SendChat g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("DCppMinVersion"), "%[minversion]", g_objSettings.DCMinVersion, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForTooOldDcppRedirectAddress & "|$ForceMove " & g_objSettings.ForTooOldDcppRedirectAddress
76:                End If
77:            Else
78:                strMessage = Replace(curUser.GetCoreMsgStr("DCppMinVersion"), "%[minversion]", g_objSettings.DCMinVersion, 1, 1)
79:            End If
'-------------------------------
        Case HSRatio
'-------------------------------
            'Redirect if necessary
83:            If g_objSettings.RedirectFSlotPerHub Then
84:                NextRedirect
                'Send message as either private or in the main chat
86:                If g_objSettings.SendMsgAsPrivate Then
87:                    curUser.SendPrivate g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("HSRatio"), "%[hsratio]", g_objSettings.DCSlotsPerHub, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForSlotPerHubRedirectAddress & "|$ForceMove " & g_objSettings.ForSlotPerHubRedirectAddress
88:                Else
89:                    curUser.SendChat g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("HSRatio"), "%[hsratio]", g_objSettings.DCSlotsPerHub, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForSlotPerHubRedirectAddress & "|$ForceMove " & g_objSettings.ForSlotPerHubRedirectAddress
90:                End If
91:            Else
92:                strMessage = Replace(curUser.GetCoreMsgStr("HSRatio"), "%[hsratio]", g_objSettings.DCSlotsPerHub, 1, 1)
93:            End If
'-------------------------------
        Case BSRatio
'-------------------------------
        'Redirect if necessary
97:            If g_objSettings.RedirectFBWPerSlot Then
98:                NextRedirect
                'Send message as either private or in the main chat
100:                If g_objSettings.SendMsgAsPrivate Then
101:                    curUser.SendPrivate g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("BSRatio"), "%[bsratio]", g_objSettings.DCBandPerSlot, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForBWPerSlotRedirectAddress & "|$ForceMove " & g_objSettings.ForBWPerSlotRedirectAddress
102:                Else
103:                    curUser.SendChat g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("BSRatio"), "%[bsratio]", g_objSettings.DCBandPerSlot, 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForBWPerSlotRedirectAddress & "|$ForceMove " & g_objSettings.ForBWPerSlotRedirectAddress
104:                End If
105:            Else
106:                strMessage = Replace(curUser.GetCoreMsgStr("BSRatio"), "%[bsratio]", g_objSettings.DCBandPerSlot, 1, 1)
107:            End If
'-------------------------------
        Case NoTag
'-------------------------------
            'Redirect if necessary
111:            If g_objSettings.RedirectFNoTag Then
112:                NextRedirect
                'Send message as either private or in the main chat
114:                If g_objSettings.SendMsgAsPrivate Then
115:                    curUser.SendPrivate g_objSettings.BotName, curUser.GetCoreMsgStr("DenyNoTag") & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForNoTagRedirectAddress & "|$ForceMove " & g_objSettings.ForNoTagRedirectAddress
116:                Else
117:                    curUser.SendChat g_objSettings.BotName, curUser.GetCoreMsgStr("DenyNoTag") & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForNoTagRedirectAddress & "|$ForceMove " & g_objSettings.ForNoTagRedirectAddress
118:                End If
119:            Else
120:                strMessage = curUser.GetCoreMsgStr("DenyNoTag")
121:           End If
'-------------------------------
        Case MaxShare
'-------------------------------
            'Redirect if necessary
125:            If g_objSettings.RedirectFMaxShare Then
126:                NextRedirect
                'Send message as either private or in the main chat
128:                If g_objSettings.SendMsgAsPrivate Then
129:                    curUser.SendPrivate g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("MaxShare"), "%[maxshare]", g_objFunctions.ShareSize(g_objSettings.MaxShare), 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForMaxShareRedirectAddress & "|$ForceMove " & g_objSettings.ForMaxShareRedirectAddress
130:                Else
131:                    curUser.SendChat g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("MaxShare"), "%[maxshare]", g_objFunctions.ShareSize(g_objSettings.MaxShare), 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForMaxShareRedirectAddress & "|$ForceMove " & g_objSettings.ForMaxShareRedirectAddress
132:                End If
133:            Else
134:                strMessage = Replace(curUser.GetCoreMsgStr("MaxShare"), "%[maxshare]", g_objFunctions.ShareSize(g_objSettings.MaxShare), 1, 1)
135:            End If
'-------------------------------
        Case FakeShare
'-------------------------------
            'Redirect if necessary
139:            If g_objSettings.RedirectFFakeShare Then
140:                NextRedirect
                'Send message as either private or in the main chat
142:                If g_objSettings.SendMsgAsPrivate Then
143:                    curUser.SendPrivate g_objSettings.BotName, curUser.GetCoreMsgStr("Faker") & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForFakeShareRedirectAddress & "|$ForceMove " & g_objSettings.ForFakeShareRedirectAddress
144:                Else
145:                    curUser.SendChat g_objSettings.BotName, curUser.GetCoreMsgStr("Faker") & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForFakeShareRedirectAddress & "|$ForceMove " & g_objSettings.ForFakeShareRedirectAddress
146:                End If
147:            Else
148:                If g_objSettings.SendMsgAsPrivate Then
149:                    curUser.SendPrivate g_objSettings.BotName, curUser.GetCoreMsgStr("Faker")
150:                Else
151:                    curUser.SendChat g_objSettings.BotName, curUser.GetCoreMsgStr("Faker")
152:                End If
                
154:                DoEvents
155:                g_objIPBans.Add curUser.IP, 180
156:           End If
'-------------------------------
        Case FakeTag
'-------------------------------
            'Redirect if necessary
160:            If g_objSettings.RedirectFFakeTag Then
161:                NextRedirect
                'Send message as either private or in the main chat
163:                If g_objSettings.SendMsgAsPrivate Then
164:                    curUser.SendPrivate g_objSettings.BotName, curUser.GetCoreMsgStr("Faker") & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForFakeTagRedirectAddress & "|$ForceMove " & g_objSettings.ForFakeTagRedirectAddress
165:                Else
166:                    curUser.SendChat g_objSettings.BotName, curUser.GetCoreMsgStr("Faker") & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForFakeTagRedirectAddress & "|$ForceMove " & g_objSettings.ForFakeTagRedirectAddress
167:                End If
168:            Else
169:                If g_objSettings.SendMsgAsPrivate Then
170:                    curUser.SendPrivate g_objSettings.BotName, curUser.GetCoreMsgStr("Faker")
171:                Else
172:                    curUser.SendChat g_objSettings.BotName, curUser.GetCoreMsgStr("Faker")
173:                End If
                
175:                DoEvents
176:                g_objIPBans.Add curUser.IP, 180
177:           End If
'-------------------------------
        Case MinShare
'-------------------------------
            'Redirect if necessary
181:            If g_objSettings.RedirectFMinShare Then
182:                NextRedirect
            
                'Send message as either private or in the main chat
185:                If g_objSettings.SendMsgAsPrivate Then
186:                    curUser.SendPrivate g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("MinShare"), "%[minshare]", g_objFunctions.ShareSize(g_objSettings.MinShare), 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForMinShareRedirectAddress & "|$ForceMove " & g_objSettings.ForMinShareRedirectAddress
187:                Else
188:                    curUser.SendChat g_objSettings.BotName, Replace(curUser.GetCoreMsgStr("MinShare"), "%[minshare]", g_objFunctions.ShareSize(g_objSettings.MinShare), 1, 1) & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForMinShareRedirectAddress & "|$ForceMove " & g_objSettings.ForMinShareRedirectAddress
189:                End If
190:            Else
191:                strMessage = Replace(curUser.GetCoreMsgStr("MinShare"), "%[minshare]", g_objFunctions.ShareSize(g_objSettings.MinShare), 1, 1)
192:            End If

'-------------------------------
        Case Socks5
'-------------------------------
            'g_objSettings.Socks5Msg
            'I think socks5 dont follow redirecting?(rarely)
198:            strMessage = curUser.GetCoreMsgStr("Socks5")

'-------------------------------
        Case PassiveMode
'-------------------------------
202:            If g_objSettings.RedirectFPasMode Then
203:                NextRedirect
                'Send message as either private or in the main chat
205:                If g_objSettings.SendMsgAsPrivate Then
206:                    curUser.SendPrivate g_objSettings.BotName, curUser.GetCoreMsgStr("PassiveMode") & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForPasModeRedirectAddress & "|$ForceMove " & g_objSettings.ForPasModeRedirectAddress
207:                Else
208:                    curUser.SendChat g_objSettings.BotName, curUser.GetCoreMsgStr("PassiveMode") & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RedirectedTo") & g_objSettings.ForPasModeRedirectAddress & "|$ForceMove " & g_objSettings.ForPasModeRedirectAddress
209:                End If
210:            Else
211:                 strMessage = curUser.GetCoreMsgStr("PassiveMode")
212:           End If
           
        Case NoCOClients
214:            strMessage = curUser.GetCoreMsgStr("NoCOClients")

216:    End Select
'---ROLL-----------END---------------REDIRECT---------PART

    'If there is no message, don't send anything
220:    If LenB(strMessage) Then
        'Send message as either private or in the main chat
222:        If g_objSettings.SendMsgAsPrivate Then
223:            curUser.SendPrivate g_objSettings.BotName, strMessage
224:        Else
225:            curUser.SendChat g_objSettings.BotName, strMessage
226:        End If
227:    End If

    'Close winsock
230:    DoEvents
231:    wskLoop_Close curUser.iWinsockIndex

233:    Exit Sub

235:
Err:
236:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.FailedConf(, " & intType & ")"
End Sub

Private Function ProcessMyINFO(ByRef curUser As clsUser, ByRef strMyInfo As String) As Boolean
1:    Dim arrSplit()  As String
2:    Dim arrTag()    As String
3:    Dim lngLoop     As Long
4:    Dim lngUB       As Long
5:    Dim dblShare    As Double
6:    Dim dblVersion  As Double
7:    Dim lngSlots    As Long
8:    Dim lngHubs     As Long
9:    Dim lngO        As Long
10:    Dim intID       As Integer
11:    Dim strStatus      As String

13:    On Error GoTo Err
    'It is parsed upto the $ALL part of the string
    'Format : $ALL <name> <description>$ $<connection><chr_flag>$<email>$<share>$

    'Check if client is ChatOnly and/or in away mode
18:    strStatus = g_objRegExps.CaptureSubStr(strMyInfo, GETSTATUS)

20:    If LenB(strStatus) Then
        Select Case AscW(strStatus)
            Case 1, 4, 5, 8, 9
21:             curUser.isAFK = False
            Case 2, 3, 6, 7, 10, 11
22:             curUser.isAFK = True
            Case 12, 13
                'ChatOnly client
24:             If g_objSettings.ACOClients Then
25:                 curUser.ChatOnly = True
26:                 curUser.isAFK = False
27:             Else
28:                 If g_objSettings.OPBypass Then
29:                     If curUser.Class < vip Then
30:                         FailedConf curUser, NoCOClients
31:                         Exit Function
32:                    Else
33:                        curUser.ChatOnly = True
34:                        curUser.isAFK = False
35:                     End If
36:                Else
37:                    FailedConf curUser, NoCOClients
38:                    Exit Function
39:                 End If
40:             End If
            Case 14, 15
                'ChatOnly client
42:             If g_objSettings.ACOClients Then
43:                 curUser.ChatOnly = True
44:                 curUser.isAFK = True
45:             Else
46:                 If g_objSettings.OPBypass Then
47:                     If curUser.Class < vip Then
48:                         FailedConf curUser, NoCOClients
49:                         Exit Function
50:                    Else
51:                        curUser.ChatOnly = True
52:                        curUser.isAFK = True
53:                     End If
54:                Else
55:                    FailedConf curUser, NoCOClients
56:                    Exit Function
57:                 End If
58:             End If
            Case Else
                'unknown Status--> buggy client or fake tag
60:                 If g_objSettings.DCValidateTags Then
61:                        If g_objSettings.OPBypass Then
62:                            If curUser.Class < vip Then
                                'Fake tag...
64:                             FailedConf curUser, FakeTag
65:                             Exit Function
66:                            Else
67:                                curUser.isAFK = False
68:                            End If
69:                        Else
                            'Fake tag...
71:                            FailedConf curUser, FakeTag
72:                            Exit Function
73:                        End If
74:                 Else
75:                    curUser.isAFK = False
76:                 End If
77:     End Select
78:    Else
            'Missing status flag
80:            If g_objSettings.DCValidateTags Then
81:                If g_objSettings.OPBypass Then
82:                    If curUser.Class < vip Then
                        'Fake tag...
84:                        FailedConf curUser, FakeTag
85:                        Exit Function
86:                    Else
87:                        curUser.isAFK = False
88:                    End If
89:                Else
                    'Fake tag...
91:                    FailedConf curUser, FakeTag
92:                    Exit Function
93:                End If
94:            Else
95:                curUser.isAFK = False
96:            End If
97:    End If

99:    intID = 0

101:    arrSplit = Split(MidB$(strMyInfo, 11), "$")
    
    'Make sure we have the right number of params
104:    If UBound(arrSplit) = 5 Then
        'Get their share
106:        dblShare = CDbl(arrSplit(4))
    
        'Make sure the rules apply to this user
        '#If FLASHCHAT Then
110:            If Not curUser.ChatOnly Then
        '#End If
            
113:        If g_objSettings.OPBypass Then
114:            If curUser.Class < vip Then intID = 1
115:        Else
116:            intID = 1
117:        End If
            
        '#If FLASHCHAT Then
120:            End If
        '#End If

        'If the ID is nonzero, then check rules
124:        If intID Then
            'Set back to zero
126:            intID = 0

            'Check to see if this is an MLDonkey client
            'They usually have a space in front of their share size
130:            If g_objSettings.AutoKickMLDC Then _
                If AscW(arrSplit(4)) = 32 Then _
                    curUser.Kick 60: Exit Function
            
            'Check if they are fake sharing
135:        If g_objSettings.CheckFakeShare Then
136:            If dblShare Then
                Select Case True
                    Case Round(Round(dblShare / 1073741824, 6) * 1073741824, 0) = dblShare
137:                        FailedConf curUser, FakeShare
138:                        Exit Function
                    Case g_objRegExps.TestStr(CStr(dblShare), DENYSHARESIZE)
139:                        FailedConf curUser, FakeShare
140:                        Exit Function
141:                End Select
142:            End If
143:        End If

            'Min share check
146:            If g_objSettings.MentoringSystem Then
                'Good will policy if using mentoring system (must share something)
148:                If dblShare <= 0 Then FailedConf curUser, MinShare: Exit Function
149:            Else
150:                If g_objSettings.MinShare > dblShare Or dblShare < 0 Then FailedConf curUser, MinShare: Exit Function
151:            End If
    
            'Max share check
154:            If g_objSettings.MaxShare Then _
                If g_objSettings.MaxShare < dblShare Then _
                    FailedConf curUser, MaxShare: Exit Function
        
            'Tag checks (get last "<")
159:            lngLoop = InStrB(1, StrReverse(arrSplit(0)), "<")

161:            If lngLoop Then
                'If found, check tag name
163:                lngLoop = LenB(arrSplit(0)) - lngLoop + 2
164:                lngUB = InStrB(lngLoop, arrSplit(0), ">")
                
                'If there is a ">" then check for supported tag names
167:                If lngUB Then
168:                    arrSplit(0) = MidB$(arrSplit(0), lngLoop, lngUB - lngLoop)
169:                    lngUB = InStrB(1, arrSplit(0), " ")
                    
171:                    If lngUB Then
                        'Get the ID
173:                        On Error Resume Next
174:                        intID = m_colTags(LeftB$(arrSplit(0), lngUB - 1)).ID
175:                        On Error GoTo Err
                        
                        'If the ID is nonzero, then it is a real tag
178:                        If intID Then
179:                            arrTag = Split(MidB$(arrSplit(0), lngUB + 2), ",")
            
181:                            lngUB = UBound(arrTag)
            
                            'If lngUB is less than 3, they must be faking their tag
184:                            If lngUB < 3 Then
185:                                If g_objSettings.DCValidateTags Then FailedConf curUser, FakeTag: Exit Function
186:                            Else
                                'Perform the rest of the validation checks if needed
188:                                If g_objSettings.DCValidateTags Then
                                    'First element must be V
190:                                    If AscW(arrTag(0)) = 86 Then arrTag(0) = MidB$(arrTag(0), 5) Else FailedConf curUser, FakeTag: Exit Function
                
                                    'Second element must be M
193:                                    If AscW(arrTag(1)) = 77 Then arrTag(1) = MidB$(arrTag(1), 5) Else FailedConf curUser, FakeTag: Exit Function
    
                                    Select Case g_objRegExps.CaptureSubStr(strMyInfo, GETDCMODE)
                                        Case "A"
195:                                            curUser.Passive = False
                                            
                                        Case "P"
197:                                            If g_objSettings.DenyPassive Then
198:                                                FailedConf curUser, PassiveMode: Exit Function
199:                                            End If
                                            
201:                                            curUser.Passive = True
                                            
                                        Case "5"
203:                                            If g_objSettings.DenySocks5 Then
204:                                                FailedConf curUser, Socks5: Exit Function
205:                                            End If

207:                                            curUser.Passive = True
                                            
                                        Case Else
                                            'more then 1 capture collection or no valid mode
210:                                            FailedConf curUser, FakeTag: Exit Function

212:                                    End Select

                                    'Third element must be H
215:                                    If AscW(arrTag(2)) = 72 Then arrTag(2) = MidB$(arrTag(2), 5) Else FailedConf curUser, FakeTag: Exit Function
                
                                    'Fourth element must be S
218:                                    If AscW(arrTag(3)) = 83 Then arrTag(3) = MidB$(arrTag(3), 5) Else FailedConf curUser, FakeTag: Exit Function
219:                                Else
                                    'Skip beginning "C:" (where C = character) in the beginning of each array element
221:                                    arrTag(0) = MidB$(arrTag(0), 5)
222:                                    arrTag(1) = MidB$(arrTag(1), 5)
223:                                    arrTag(2) = MidB$(arrTag(2), 5)
224:                                    arrTag(3) = MidB$(arrTag(3), 5)
225:                                End If
            
                                'Check the min DC++ version if it is a ++ tag
228:                                If intID = 1 Then
                                    'DC++ does not require $Hello to be sent before $MyINFO, so therefore
                                    'we can skip sending it, saving bandwidth while there are no needed
                                    'protocol changes which DC++ must recognize)
232:                                    curUser.NoHello = True
                                    
                                    'Extract version
235:                                    If m_blnCommaDecimal Then _
                                        dblVersion = StrToDbl(Replace(arrTag(0), ".", ",")) _
                                    Else _
                                        dblVersion = Val(arrTag(0))
                            
240:                                    If g_objSettings.DCMinVersion Then _
                                        If g_objSettings.DCMinVersion > dblVersion Then FailedConf curUser, DCppversion: Exit Function
242:                                End If
                    
                                'Check for exceptions in S:
                                Select Case intID
                                    Case 8
                                        'SdDC++ has the format S:#/# and not S:#
246:                                        lngSlots = GetByte(MidB$(arrTag(3), InStrB(3, arrTag(3), "/") + 2))
                                    Case 3
                                        'DCGUI sometimes has * in it's S: param to denote unlimited slots
                                        'At the moment, I'm content to let them bypass the min slot requirement
249:                                        If AscW(arrTag(3)) = 42 Then _
                                            lngSlots = g_objSettings.MinSlots + 1 _
                                        Else _
                                            lngSlots = CLng(arrTag(3))
                                    Case Else
253:                                        lngSlots = CLng(arrTag(3))
254:                                End Select
                                
                                'If the number of slots is 0, then kick if
                                'validating tags, else ignore and set to 1
                                '(for division purposes)
259:                                If lngSlots = 0 Then
260:                                    If g_objSettings.DCValidateTags Then
261:                                        FailedConf curUser, FakeTag: Exit Function
262:                                    Else
263:                                        lngSlots = 1
264:                                    End If
265:                                End If
                
                                'Check for tag extensions
268:                                If lngUB > 3 Then
269:                                    For lngLoop = 3 To lngUB
                                        'Find out if we support it
                                        Select Case AscW(arrTag(lngLoop))
                                            Case 79 'O
                                                'Format - O:#
                                                
                                                'DC tags (NMDC 2.0) use O: for free/open slots
274:                                                If Not intID = 2 Then
275:                                                    lngO = CLng(MidB$(arrTag(lngLoop), 5))
276:                                                    If lngO > g_objSettings.DCOSpeed Then lngSlots = lngSlots + g_objSettings.DCOSlots
277:                                                End If
                                            Case 76, 66, 85 'L, B, U
                                                'Format - L:#, B:#, U:#
                                        
                                                'Perform bandwidth/slot ratio check
281:                                                If g_objSettings.DCBandPerSlot Then
282:                                                    If intID = 3 Then
                                                        'DCGUI may have * in it's limiter param meaning
                                                        'it is not limiting
285:                                                        arrTag(lngLoop) = MidB$(arrTag(lngLoop), 5)
                                                        
287:                                                        If Not AscW(CStr(arrTag(lngLoop))) = 42 Then
288:                                                                If CLng(arrTag(lngLoop)) < g_objSettings.DCBandPerSlot Then
289:                                                                    FailedConf curUser, BSRatio
290:                                                                    Exit Function
291:                                                                End If
292:                                                            End If
293:                                                    Else
294:                                                        If (CLng(MidB$(arrTag(lngLoop), 5)) / lngSlots) < g_objSettings.DCBandPerSlot Then _
                                                            FailedConf curUser, BSRatio: Exit Function
296:                                                    End If
297:                                                End If
                                            Case 70 'F
                                                'Format - F:#/#
                                        
                                                'Perform bandwidth/slot ratio check after
                                                'extracting upload limit
302:                                                If g_objSettings.DCBandPerSlot Then
303:                                                    If (CLng(MidB$(arrTag(lngLoop), InStrB(1, arrTag(lngLoop), "/") + 2)) / lngSlots) < g_objSettings.DCBandPerSlot Then _
                                                        FailedConf curUser, BSRatio: Exit Function
305:                                                End If
306:                                        End Select
307:                                    Next
308:                                End If
                
                                'Min slot check
311:                                If g_objSettings.MinSlots > lngSlots Then FailedConf curUser, MinSlots: Exit Function
                                
                                
                                'Max slot check
315:                                If g_objSettings.MaxSlots Then
316:                                    If g_objSettings.MaxSlots < lngSlots Then FailedConf curUser, MaxSlots: Exit Function
317:                                End If
                                
                                'Split up the max hubs if using DC++ 0.24 / other clients
320:                                If InStrB(1, arrTag(2), "/") Then
                                    'If using DC++ and if their version is pre 0.24, then they are faking
322:                                    If intID = 1 Then _
                                        If g_objSettings.DCValidateTags Then _
                                            If dblVersion < 0.24 Then _
                                                FailedConf curUser, FakeTag: Exit Function
                                        
327:                                    arrTag = Split(arrTag(2), "/")
                
                                    'H: property MUST have 3 array elements; otherwise they are faking
330:                                    If UBound(arrTag) = 2 Then
                                        'Make sure all items are numerical
                                        Select Case False
                                            Case IsNumeric(arrTag(0)), IsNumeric(arrTag(1)), IsNumeric(arrTag(2)): FailedConf curUser, FakeTag: Exit Function
332:                                        End Select
                                    
                                        'Count up total - included hubs where opped if necessary
335:                                        If g_objSettings.DCIncludeOPed Then _
                                            lngHubs = CLng(arrTag(0)) + CLng(arrTag(1)) + CLng(arrTag(2)) _
                                        Else _
                                            lngHubs = CLng(arrTag(0)) + CLng(arrTag(1))
339:                                    Else
340:                                        If g_objSettings.DCValidateTags Then
341:                                            FailedConf curUser, FakeTag
342:                                            Exit Function
343:                                        End If
344:                                    End If
345:                                Else
                                    'If using DC++ and if their version is post 0.24, they are faking
347:                                    If intID = 1 Then _
                                        If g_objSettings.DCValidateTags Then _
                                            If dblVersion >= 0.24 Then _
                                                FailedConf curUser, FakeTag: Exit Function
                                            
352:                                    lngHubs = CLng(arrTag(2))
353:                                End If
                
                                'If the user is not registered and their hub count is zero, they are faking
                                'If they are registered, this prevents division by 0
357:                                If lngHubs = 0 Then
                                    Select Case True
                                        Case curUser.Class > Normal, g_objSettings.PasswordMode
358:                                            lngHubs = 1
                                        Case Else
                                            'If the user is using QuickList, and they are not registered,
                                            'their hub count has yet to increment, assuming they are logging
                                            'in
362:                                            If curUser.QuickList Then
363:                                                If curUser.State = Logged_In Then _
                                                    If g_objSettings.DCValidateTags Then FailedConf curUser, FakeTag: Exit Function
365:                                            Else
366:                                                If g_objSettings.DCValidateTags Then FailedConf curUser, FakeTag: Exit Function
367:                                            End If
368:                                    End Select
369:                                End If
                
                                'Max hub check
372:                                If g_objSettings.DCMaxHubs Then _
                                    If g_objSettings.DCMaxHubs < lngHubs Then FailedConf curUser, MaxHubs: Exit Function
                
                                'Slots per hub check
376:                                If g_objSettings.DCSlotsPerHub Then _
                                    If (lngSlots / lngHubs) < g_objSettings.DCSlotsPerHub Then _
                                        FailedConf curUser, HSRatio: Exit Function
                                
                                'TheNOP svn 40
                                'Set passive status (if active, then add slots for fake slot check)
                                'If AscW(arrTag(1)) = 65 Then
                                '    curUser.Passive = False
                                'Else
                                '    curUser.Passive = True
                                'End If
387:                            End If
            
389:                            Erase arrTag
390:                        Else
391:                            GoTo NoTag
392:                        End If
393:                    Else
394:                        GoTo NoTag
395:                    End If
396:                Else
397:                    GoTo NoTag
398:                End If
399:            Else
400:
NoTag:
                'If needed, disconnect the user since they have no tag
402:                If g_objSettings.DenyNoTag Then
403:                    FailedConf curUser, NoTag
404:                    Exit Function
405:                Else
                    'Make another MLDC check
407:                    If g_objSettings.AutoKickMLDC Then _
                        If RightB$(arrSplit(0), 26) = "donkey client" Or RightB$(arrSplit(0), 22) = "mldc client" Then _
                            curUser.Kick 60: Exit Function
                            
411:                    If curUser.State = Logged_In Then
412:                        If curUser.NetInfo Then _
                            curUser.SendData "$GetNetInfo|"
414:                    Else
                        'Send GetNetInfo if the user is using NMDC2 (first attempt)
416:                        dblVersion = curUser.iVersion
                        
418:                        If dblVersion = 1.0091 Then
419:                            curUser.SendData "$GetNetInfo|"
420:                        Else
421:                            If dblVersion >= 2 Then _
                                If dblVersion <= 3 Then _
                                    curUser.SendData "$GetNetInfo|"
424:                        End If
425:                    End If
426:                End If
427:            End If
428:        End If
    
        'Check if they are in away mode
        'Select Case AscW(RightB$(arrSplit(2), 2))
        '    Case 2, 3, 6, 7, 10, 11
        '        curUser.isAFK = True
        '    Case Else
        '        curUser.isAFK = False
        'End Select
    
        'Update settings, do not count ChatOnly...
439:        curUser.sMyInfoString = "$MyINFO " & strMyInfo
            
441:        If Not curUser.ChatOnly Then
                'Don't add invisible user's share
443:                If curUser.Visible Then
444:                g_colUsers.iTotalBytesShared = g_colUsers.iTotalBytesShared - curUser.iBytesShared + dblShare
445:                curUser.iBytesShared = dblShare
446:                End If
447:        End If
    
        'Passed all checks
450:        ProcessMyINFO = True
451:    Else
        'Not DC complient
453:        wskLoop_Close curUser.iWinsockIndex
454:    End If
    
456:    Exit Function

458:
Err:
459:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.ProcessMyINFO(, """ & strMyInfo & """)"

    'Something is wrong with their MyINFO string so we disconnect them
462:    wskLoop_Close curUser.iWinsockIndex
End Function

Private Function ValidateNick(ByRef curUser As clsUser, ByRef strName As String, Optional ByRef strMyInfo As String) As Boolean
1:    Dim i           As Integer
2:    Dim objTmp      As Object
3:    Dim objUser     As clsUser

5:    On Error GoTo Err

    'Cannot be longer than 40 chars
8:    If LenB(strName) > 80 Then
9:        curUser.SendData "<" & g_objSettings.BotName & "> " & g_objFunctions.GetENLangStr("NickLength") & "|"

11:        DoEvents
12:        wskLoop_Close curUser.iWinsockIndex

14:        Exit Function
15:    End If

    'Disallow certain characters, """|'|/|\s"
18:    If g_objRegExps.TestStr(strName, CHRSTODENYINNICK) Then
19:        curUser.SendData "<" & g_objSettings.BotName & "> " & g_objFunctions.GetENLangStr("ChrInNick") & "|"

21:        DoEvents
22:        wskLoop_Close curUser.iWinsockIndex

24:        Exit Function
25:    End If

27:    On Error Resume Next

    'Copy to sLanguageID property(user language preference)
30:    Set objTmp = m_objPermaCon.Execute("Select UsrStatic.i18n From UsrStatic Where UsrStatic.UserName=""" & strName & """;", , 1)
31:    If LenB(objTmp.Collect(0)) Then curUser.sLanguageID = objTmp.Collect(0)
32:    If curUser.sLanguageID = vbNullString Then curUser.sLanguageID = "En"

34:    On Error GoTo Err

    'Find out their registered status
    Select Case g_objRegistered.Registered(strName)
        Case Locked 'Nickname is banned
            'Determine message to send
38:            If g_objSettings.DescriptiveBanMsg Then
39:                curUser.SendData "<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("IPPermBan") & "|"
40:            Else
41:                curUser.SendData "<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("IPBanned") & "|"
42:            End If

44:            DoEvents
45:            wskLoop_Close curUser.iWinsockIndex
        Case Unknown 'Not registered
46:            If g_objSettings.PreventSearchBots Then
                'If not registered, then it could be a search tool
48:                If InStrB(1, strName, "search") Then
49:                    wskLoop_Close curUser.iWinsockIndex
50:                    Exit Function
51:                End If
52:            End If

            'If only registered users, get rid of them
55:            If g_objSettings.RegOnly Then
                'Redirect if necessary, otherwise disconnect
57:                If g_objSettings.AutoRedirectNonReg Then
58:                    NextRedirect
59:                    curUser.SendData "<" & g_objSettings.BotName & "> " & Replace(vbWelcome, "%[UpTime]", HubUpTime()) & "<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RegOnlyRedirTo") & g_objSettings.RedirectIP & "|$ForceMove " & g_objSettings.RedirectIP & "|"

61:                    DoEvents
62:                    wskLoop_Close curUser.iWinsockIndex
63:                Else
64:                    curUser.SendData "<" & g_objSettings.BotName & "> " & Replace(vbWelcome, "%[UpTime]", HubUpTime()) & "<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RegOnly") & "|"

66:                    DoEvents
67:                    wskLoop_Close curUser.iWinsockIndex
68:                End If

70:                Exit Function
71:            Else
                'If redirecting only non-registered or non-opped users if full, make the check
73:                If g_objSettings.AutoRedirectFullNonOps Or g_objSettings.AutoRedirectFullNonReg Then
74:                        If g_colUsers.count >= g_objSettings.MaxUsers Then
75:                            NextRedirect

77:                            curUser.SendData "<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("FullRedirTo") & g_objSettings.RedirectIP & "|$ForceMove " & g_objSettings.RedirectIP & "|"

79:                            DoEvents
80:                            wskLoop_Close curUser.iWinsockIndex

82:                            Exit Function
83:                        End If
84:                End If

                'Check if their nickname is in use
87:                If g_colUsers.Online(strName) Then
88:                    i = g_colUsers.ItemByName(strName).iWinsockIndex

                    'If it is, then check if the user's winsock is closed
91:                    If wskLoop(i).State = 0 Then
92:                        wskLoop_Close i
93:                    Else
                        'If it is still open, then compare their IPs
                        'If they are the same, disconnect the ghost
96:                        If wskLoop(i).RemoteHostIP = curUser.IP Then
97:                            wskLoop_Close i
98:                        Else
99:                            curUser.SendData "<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("NickTaken") & "|$ValidateDenide|"

101:                            DoEvents
102:                            wskLoop_Close curUser.iWinsockIndex

104:                            Exit Function
105:                        End If
106:                    End If
107:                End If

109:                If LenB(strMyInfo) Then
110:                    If ProcessMyINFO(curUser, strMyInfo) Then
                        'Since we have their MyINFO string, they must be QuickList
                        'They are not registered so that means they are now fully logged in
                        'unless the hub is running in password mode
114:                        curUser.sName = strName
115:                        g_colUsers.UpdateName curUser

                        'If the hub is using password mode, then ask for it
118:                        If g_objSettings.PasswordMode Then
119:                            curUser.SendData "<" & g_objSettings.BotName & "> " & Replace(vbWelcome, "%[UpTime]", HubUpTime()) & "$HubName " & g_objSettings.HubName & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("PassMode") & "|$GetPass|"

121:                            curUser.State = Wait_PassPM
122:                        Else
123:                            curUser.Class = Normal

                            'Send hub message and hub name
126:                            curUser.SendData "<" & g_objSettings.BotName & "> " & Replace(vbWelcome, "%[UpTime]", HubUpTime()) & "$HubName " & g_objSettings.HubName & "|"

128:                            curUser.QNL = True
129:                            g_colUsers.UpdateLogIn curUser
130:                            SEvent_UserConnected curUser
131:                        End If
132:                    Else
133:                        Exit Function
134:                    End If
135:                Else
                    'Add to user name collection
137:                    curUser.sName = strName
138:                    g_colUsers.UpdateName curUser

                    'If the hub is using password mode, ask for the password, else
                    'wait for their MyINFO string
142:                    If g_objSettings.PasswordMode Then
143:                        curUser.SendData "<" & g_objSettings.BotName & "> " & Replace(vbWelcome, "%[UpTime]", HubUpTime()) & "$HubName " & g_objSettings.HubName & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("PassMode") & "|$GetPass|"

145:                        curUser.State = Wait_PassPM
146:                    Else
147:                        curUser.State = Wait_Info
148:                        curUser.Class = Normal
                            'Send welcome message / hub name / $Hello
150:                        curUser.SendData "<" & g_objSettings.BotName & "> " & Replace(vbWelcome, "%[UpTime]", HubUpTime()) & "$HubName " & g_objSettings.HubName & "|$Hello " & strName & "|"

152:                    End If
153:                End If
154:            End If
        Case Mentored, Invisible, Registered, vip 'Registered - Non op
            'Redirect if necessary
156:            If g_objSettings.AutoRedirectFullNonOps Then
157:                If g_colUsers.count >= g_objSettings.MaxUsers Then
158:                    NextRedirect
159:                    curUser.SendData "<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("FullRedirTo") & g_objSettings.RedirectIP & "|$ForceMove " & g_objSettings.RedirectIP & "|"

161:                    DoEvents
162:                    wskLoop_Close curUser.iWinsockIndex

164:                    Exit Function
165:                End If
166:            End If

            Select Case g_colUsers.Online(strName)
                Case 0
                Case -1
                'm_colNames
                    'i = g_colUsers.ItemByName(strName).iWinsockIndex
                    'If Not i = curUser.iWinsockIndex Then
                    '    If wskLoop(i).State = 0 Then wskLoop_Close i
                    'End If
                    
174:                    For Each objUser In g_colUsers
175:                        If objUser.sName = strName Then
176:                            i = objUser.iWinsockIndex
177:                            If Not i = curUser.iWinsockIndex Then
178:                                If wskLoop(i).State = 0 Then wskLoop_Close i
179:                            End If
180:                        End If
181:                    Next
                Case 1
                'm_colNLoggingIn
183:                    For Each objUser In g_colUsers
184:                        If objUser.sName = strName Then
185:                            If Not objUser.iWinsockIndex = curUser.iWinsockIndex Then
186:                                wskLoop_Close objUser.iWinsockIndex
187:                            End If
188:                        End If
189:                    Next
190:            End Select

192:            curUser.sName = strName
193:            g_colUsers.UpdateName curUser

            'Set MyINFO string if they are QuickList to check after log in
196:            If LenB(strMyInfo) Then curUser.sMyInfoString = strMyInfo

            'Send welcome / password request
199:            curUser.SendData "<" & g_objSettings.BotName & "> " & Replace(vbWelcome, "%[UpTime]", HubUpTime()) & "$HubName " & g_objSettings.HubName & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RegPass") & "|$GetPass|"

201:            curUser.State = Wait_Pass
        Case Else 'Registered - Op
            'If their nickname is taken, either disconnect currently logging in
            'user, or wait until they do a password check (if it isn't a ghost)
            Select Case g_colUsers.Online(strName)
                Case 0
                Case -1
                    'i = g_colUsers.ItemByName(strName).iWinsockIndex
                    'If Not i = curUser.iWinsockIndex Then
                    '    If wskLoop(i).State = 0 Then wskLoop_Close i
                    'End If
208:                    For Each objUser In g_colUsers
209:                        If objUser.sName = strName Then
210:                            i = objUser.iWinsockIndex
211:                            If Not i = curUser.iWinsockIndex Then
212:                                If wskLoop(i).State = 0 Then wskLoop_Close i
213:                            End If
214:                        End If
215:                    Next

                Case 1
217:                    For Each objUser In g_colUsers
218:                        If objUser.sName = strName Then
219:                            If Not objUser.iWinsockIndex = curUser.iWinsockIndex Then
220:                                wskLoop_Close objUser.iWinsockIndex
221:                            End If
222:                        End If
223:                    Next

225:            End Select

227:            curUser.sName = strName
228:            g_colUsers.UpdateName curUser

            'Set MyINFO string if they are QuickList to check after log in
231:            If LenB(strMyInfo) Then curUser.sMyInfoString = strMyInfo

            'Send welcome / password request
234:            curUser.SendData "<" & g_objSettings.BotName & "> " & Replace(vbWelcome, "%[UpTime]", HubUpTime()) & "$HubName " & g_objSettings.HubName & "|<" & g_objSettings.BotName & "> " & curUser.GetCoreMsgStr("RegPass") & "|$GetPass|"

236:            curUser.State = Wait_Pass
237:    End Select

239:    ValidateNick = True

241:    Exit Function

243:
Err:
244:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.ValidateNick(, """ & strName & """, """ & strMyInfo & """)"

    'Something is wrong with their nickname/etc so we disconnect them
247:    wskLoop_Close curUser.iWinsockIndex
End Function

Friend Sub ProcessTrigger(ByRef objUser As clsUser, ByRef strTrigger As String, ByRef blnMainChat As Boolean)
1:    Dim arrCommand()        As String
2:    Dim strIP               As String
3:    Dim strMsg              As String
4:    Dim arrTmp()            As String
5:    Dim strTmp              As String
6:    Dim lngTmp              As Long
7:    Dim intTmp              As Integer
8:    Dim objTmp              As Object
9:    Dim varTmp              As Variant
10:    Dim objCommand          As clsCommand

12:    On Error GoTo Err

14:    arrCommand = Split(strTrigger, g_objSettings.CSeperator, 3)

16:    If UBound(arrCommand) = -1 Then Exit Sub

18:    On Error Resume Next
19:    Set objCommand = g_colCommands(arrCommand(0))
20:    On Error GoTo Err

22:    If ObjPtr(objCommand) Then
        'Commands and their corresponding ID

        'reg = 1
        'admin = 2
        'ban = 3
        'banip = 4
        'banuser = 5
        'close = 6
        'info = 7
        'iplist = 8
        'listbanip = 9
        'ipscan = 10
        'listbanuser = 11
        'unbanip = 12
        'unbanuser = 13
        'help = 14
      
        'Make sure command is enabled
41:        If Not objCommand.Enabled Then Exit Sub
      
        'Make sure the user has permission to use the command
44:        If objUser.Class < objCommand.Class Then Exit Sub
      
        Select Case objCommand.ID
            'Case 1
            'Case 2
            'Case 3
            'Case 4
            'Case 5
            'Case 6
            'Case 7
            'Case 8
            'Case 9
            'Case 10
            'Case 11
            'Case 12
            'Case 13
            'Case 14
                
            Case Else
61:                If objCommand.ID > 50 Then SEvent_CustComArrival objUser, objCommand, strTrigger, blnMainChat
62:        End Select
  
    
65:    End If

67:    Exit Sub
  
69:
Err:
70:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.ProcessTrigger()"
End Sub

Private Sub Notify(ByRef strMessage As String)
1:    On Error GoTo Err

    'This will be changed later

5:    For Each m_objLoopUser In g_colUsers
6:        If m_objLoopUser.Class > InvisibleSuperOp Then m_objLoopUser.SendPrivate g_objSettings.BotName, strMessage
7:    Next
  
9:    Set m_objLoopUser = Nothing
  
11:    Exit Sub

13:
Err:
14:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.Notify()"
15:    Set m_objLoopUser = Nothing
End Sub

Private Function StrToDbl(ByVal strConvert As String) As Double
1:    Dim arrString()     As Byte
2:    Dim lngLoop         As Long
3:    Dim lngUB           As Long
4:    Dim blnDecimal      As Boolean
5:    Dim bytDecimal      As Byte
    
7:    On Error GoTo Err
    
    'Converts a string to a double value
    'This only works for comma decimal based systems (swap 46 and 44 for a
    'period decimal system; however you should use Val() for those systems)
    
    'Replace useless , or . (whichever is not the decimal)
14:    If m_blnCommaDecimal Then
15:        strConvert = Replace(strConvert, ".", vbNullString)
16:        bytDecimal = 44
17:    Else
18:        strConvert = Replace(strConvert, ",", vbNullString)
19:        bytDecimal = 46
20:    End If
    
22:    lngUB = LenB(strConvert) - 1
    
    'Make sure it isn't a zero length string
25:    If Not lngUB = -1 Then
        'Copy into array
27:        ReDim arrString(0 To lngUB) As Byte
28:        CopyMemory arrString(0), ByVal StrPtr(strConvert), lngUB
        
        'Loop through and find first non numeric char
31:        For lngLoop = 0 To lngUB Step 2
            Select Case arrString(lngLoop)
                Case 48 To 57
                Case bytDecimal: If blnDecimal Then lngLoop = lngLoop - 2: Exit For Else blnDecimal = True
                Case Else: Exit For
32:            End Select
33:        Next
34:    End If
    
    'If it wasn't the first character, then convert numerical characters to string
37:    If lngLoop Then StrToDbl = CDbl(LeftB$(strConvert, lngLoop))
    
39:    Exit Function
    
41:
Err:
42:    HandleError Err.Number, Err.Description, Erl & "|" & "frmMain.StrToDbl(" & strConvert & ")", Err.LastDllError
End Function

'------------------------------------------------------------------------------
' Setting related methods
'------------------------------------------------------------------------------
Public Sub LoadDefaultSettings()

2:      Dim lngLoop     As Long
        On Error GoTo Err

    #If FLASHCHAT Then
6:      Dim objTag(10)   As New clsTag
    #Else
8:      Dim objTag(9)   As New clsTag
    #End If
    
    
    'pre-defined TagsHelp
13:    m_arrTagRules(0) = "'NoHello' even if it's not in $Supports statement.%[LF]%[LF] Tests for minimum DC++ version V:______"
14:    m_arrTagRules(1) = "Skips standard DC++ O:# tests.%[LF]%[LF] O:# is used for free/open slots."
15:    m_arrTagRules(2) = "* in it's slot param (S:*) means unlimited slots.%[LF]%[LF] * in it's limiter param (L:*) means unlimited bandwidth.%[LF]%[LF] Reports bandwidth limit on a per slot basis, not total."
16:    m_arrTagRules(3) = "None"
17:    m_arrTagRules(4) = "None"
18:    m_arrTagRules(5) = "Uses F:#Down/#Up to report bandwidth limiting."
19:    m_arrTagRules(6) = "None"
20:    m_arrTagRules(7) = "None"
21:    m_arrTagRules(8) = "Slot param has the format S:#/#"
22:    m_arrTagRules(9) = "If you are using this option then you can figure it out for yourself."
23:    m_arrTagRules(10) = "Select a Default Tag to see if it has any special processing rules."
24:    m_arrTagRules(11) = "None"
    
26:    Set m_colTags = New Collection
27:    lstTagsDef.Clear
    
29:    objTag(0).Name = "++"
30:    objTag(0).ID = 1
31:    objTag(1).Name = "DC"
32:    objTag(1).ID = 2
33:    objTag(2).Name = "DCGUI"
34:    objTag(2).ID = 3
35:    objTag(3).Name = "oDC"
36:    objTag(3).ID = 4
37:    objTag(4).Name = "QuickDC"
38:    objTag(4).ID = 5
39:    objTag(5).Name = "DC:Pro"
40:    objTag(5).ID = 6
41:    objTag(6).Name = "SDC"
42:    objTag(6).ID = 7
43:    objTag(7).Name = "StrgDC++"
44:    objTag(7).ID = 10
45:    objTag(8).Name = "SdDC++"
46:    objTag(8).ID = 8
47:    objTag(9).Name = "Z++"
48:    objTag(9).ID = 11

    #If FLASHCHAT Then
51:    objTag(10).Name = "Chat"
52:    objTag(10).ID = 9
    #End If
    
    #If FLASHCHAT Then
56:    For lngLoop = 0 To 10
    #Else
57:    For lngLoop = 0 To 9
    #End If
59:        m_colTags.Add objTag(lngLoop), objTag(lngLoop).Name
60:        lstTagsDef.AddItem objTag(lngLoop).Name
61:    Next

63:    Call LoadDfsSettings
       
65:    DoEvents

67:  Exit Sub

69:
Err:
71:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.LoadDefaultSettings()"
End Sub
Public Sub LoadSettings()
1:     Dim objXML          As clsXMLParser
2:     Dim objNode         As clsXMLNode
3:     Dim objSubNode      As clsXMLNode
4:     Dim colNodes        As Collection
5:     Dim colSubNodes     As Collection
6:     Dim colAttributes   As Collection
7:     Dim colSupported    As Collection
8:     Dim m_colLangString As Collection
9:     Dim lvwItem         As ListItem
10:    Dim lvwItems        As ListItems
11:    Dim objTag          As clsTag
12:    Dim lngLoop         As Long
13:    Dim strTemp         As String
14:    Dim strATemp        As String
15:    Dim strSettVer      As String
16:    Dim arrTmp()        As String
17:    Dim X               As Integer
18:    Dim objCmd          As clsCommand
    
20:    On Error GoTo Err
    
22:    If g_objFileAccess.FileExists(G_APPPATH & "\PTDCH.xml") Then
23:        g_objFileAccess.CopyFile G_APPPATH & "\XML.xml", G_APPPATH & "\Settings\PTDCH.xml"
24:        g_objFileAccess.CopyFile G_APPPATH & "\Commands.xml", G_APPPATH & "\Settings\Commands.xml"
25:        g_objFileAccess.CopyFile G_APPPATH & "\PermIPBans.xml", G_APPPATH & "\Settings\PermIPBans.xml"
26:        g_objFileAccess.CopyFile G_APPPATH & "\TempIPBans.xml", G_APPPATH & "\Settings\TempIPBans.xml"
27:        g_objFileAccess.CopyFile G_APPPATH & "\DefaultProps.xml", G_APPPATH & "\Settings\DefaultProps.xml"
        'These are not parts of any previously released hubsoft. but just in case...
29:        g_objFileAccess.DeleteFile G_APPPATH & "\*.xml"
30:    End If

32:    Set objXML = New clsXMLParser

'---------------------------------------------------------------------------------
    'Load regular settings
36:    strTemp = G_APPPATH & "\Settings\PTDCH.xml"
    
39:    If g_objFileAccess.FileExists(strTemp) Then
40:        objXML.Data = g_objFileAccess.ReadFile(strTemp)
41:        objXML.Parse
    
43:        Set colNodes = objXML.Nodes(1).Nodes
    
        'Just in case...
        'On Error Resume Next
        'Set the Version from the Settings File.
48:        strSettVer = objXML.Nodes(1).Attributes("Version").Value
    
50:        For Each objNode In colNodes
51:            Set colSubNodes = objNode.Nodes
            'Using the CallByName sub may be a bit slower, but it is VERY convient
            'The settings parser no longer needs mothering! Woot!
            Select Case objNode.Name
                Case "Long"
56:                    For Each objSubNode In colSubNodes
57:                        CallByName g_objSettings, objSubNode.Name, VbLet, CLng(objSubNode.Value)
58:                    Next
                Case "Integer"
60:                    For Each objSubNode In colSubNodes
61:                        CallByName g_objSettings, objSubNode.Name, VbLet, CInt(objSubNode.Value)
62:                    Next
                Case "Boolean"
64:                    For Each objSubNode In colSubNodes
65:                        CallByName g_objSettings, objSubNode.Name, VbLet, CBool(objSubNode.Value)
66:                    Next
                Case "Double"
68:                    For Each objSubNode In colSubNodes
69:                        CallByName g_objSettings, objSubNode.Name, VbLet, CDbl(objSubNode.Value)
70:                    Next
                Case "String"
72:                    For Each objSubNode In colSubNodes
73:                        CallByName g_objSettings, objSubNode.Name, VbLet, objSubNode.Value
84:                    Next
                Case "Byte"
86:                    For Each objSubNode In colSubNodes
87:                        CallByName g_objSettings, objSubNode.Name, VbLet, CByte(objSubNode.Value)
88:                    Next
                Case "Tags"
                    ' If we have a Settings Version then all accepted Tags are saved
                    ' so clear Collection created in LoadDefaultSettings.
92:                    If strSettVer >= "0.1.1" Then Set m_colTags = New Collection
93:                    For Each objSubNode In colSubNodes
94:                        Set objTag = New clsTag
95:                        Set colAttributes = objSubNode.Attributes
    
97:                       objTag.Name = colAttributes("Name").Value
                        ' If the loaded Tag is one of the defaults give it the right default ID
                           Select Case objTag.Name
                            Case "++": objTag.ID = 1
101:                            Case "DC": objTag.ID = 2
102:                            Case "DCGUI": objTag.ID = 3
103:                            Case "oDC": objTag.ID = 4
104:                            Case "QuickDC": objTag.ID = 5
105:                            Case "DC:Pro": objTag.ID = 6
106:                            Case "SDC": objTag.ID = 7
107:                           Case "StrgDC++": objTag.ID = 10
108:                           Case "SdDC++": objTag.ID = 8
109:                           Case "Z++": objTag.ID = 11
110:                           Case "Chat": objTag.ID = 9
111:                           Case Else: objTag.ID = -1
112:                      End Select
    
114:                      m_colTags.Add objTag, objTag.Name
115:                  Next
116:            End Select
117:        Next
    
119:        On Error GoTo Err
    
        'Set min share value
122:        g_objSettings.MinShare = g_objSettings.IMinShare * (1024 ^ g_objSettings.MinShareSize)
    
124:        objXML.Clear
    
126:        Set objSubNode = Nothing
127:        Set objNode = Nothing
128:        Set colSubNodes = Nothing
129:        Set colNodes = Nothing
130:    End If

'---------------------------------------------------------------------------------
    'Load perm IP bans
134:    strTemp = G_APPPATH & "\Settings\PermIPBans.xml"

136:    If g_objFileAccess.FileExists(strTemp) Then
138:        objXML.Data = g_objFileAccess.ReadFile(strTemp)
139:        objXML.Parse
    
140:        Set colNodes = objXML.Nodes(1).Nodes
        'Set lvwItems = lvwPermIPBan.ListItems
    
        'Just in case...
        'On Error Resume Next
    
        'Make sure ban list is cleared if we are reloading settings
            g_objIPBans.ClearPerm
        'lvwItems.Clear
    
141:        For Each objNode In colNodes
142:            g_objIPBans.Add objNode.Value
            'strTemp = objNode.Value
            'lvwItems.Add , strTemp, strTemp
145:        Next
    
147:        On Error GoTo Err
    
149:        objXML.Clear
    
151:        Set objNode = Nothing
152:        Set colNodes = Nothing
153:    End If

'---------------------------------------------------------------------------------
    'Load temp IP bans
157:    strTemp = G_APPPATH & "\Settings\TempIPBans.xml"

159:    If g_objFileAccess.FileExists(strTemp) Then
160:        objXML.Data = g_objFileAccess.ReadFile(strTemp)
161:        objXML.Parse
    
163:        Set colNodes = objXML.Nodes(1).Nodes
        'Set lvwItems = lvwTempIPBan.ListItems
    
        'Just in case...
        'On Error Resume Next
    
        'Make sure ban list is cleared if we are reloading settings
170:        g_objIPBans.ClearTemp
        'lvwItems.Clear
    
171:        For Each objNode In colNodes
172:            lngLoop = DateDiff("n", Now, objNode.Attributes("Date").Value)
    
            'Make sure the date hasn't expired
175:            If lngLoop > 0 Then _
                g_objIPBans.Add objNode.Value, lngLoop
177:        Next
    
179:        On Error GoTo Err
    
181:        objXML.Clear
    
183:        Set objNode = Nothing
184:        Set colNodes = Nothing
185:    End If

'---------------------------------------------------------------------------------
    'Commands
189:    strTemp = G_APPPATH & "\Settings\Commands.xml"

190:    If g_objFileAccess.FileExists(strTemp) Then
        'Clear old commands / add defaults
192:        g_colCommands.Clear
        'shouldn't this be remove ??? except language.
        'g_colCommands.Add 1, "reg", "The register command panel", Admin, True
        'g_colCommands.Add 2, "admin", "The admin command panel", Admin, True
        'g_colCommands.Add 3, "ban", "Bans (aka locks) a username", SuperOp, True
        'g_colCommands.Add 4, "banip", "Bans an IP", SuperOp, True
        'g_colCommands.Add 5, "banuser", "Disconnects and perm bans a user (by IP)", SuperOp, True
        'g_colCommands.Add 6, "close", "Disconnects a user", Op, True
        'g_colCommands.Add 7, "info", "Retrieves information on about user", Op, True
        'g_colCommands.Add 8, "iplist", "Lists the IPs / Names of connected users", Op, True
        'g_colCommands.Add 9, "listbanip", "Lists the IPs currently banned", SuperOp, True
        'g_colCommands.Add 10, "ipscan", "Checks for users who are connected more than once (on the same IP)", SuperOp, True
        'g_colCommands.Add 11, "listbanuser", "Lists the banned (aka locked) usernames", SuperOp, True
        'g_colCommands.Add 12, "unbanip", "Unban an IP", SuperOp, True
        'g_colCommands.Add 13, "unbanuser", "Unban (aka unlock) a username", SuperOp, True
        'g_colCommands.Add 14, "help", "Description is created by all of the other commands.", Op, True
        'g_colCommands.Add 15, "language", "Changes language preference for scripts which have multi-language support.", Mentored, True
 
210:        objXML.Data = g_objFileAccess.ReadFile(strTemp)
211:        objXML.Parse
    
        'On Error Resume Next
    
215:        Set colNodes = objXML.Nodes(1).Nodes
    
217:        For Each objNode In colNodes
218:            Set colSubNodes = objNode.Attributes
219:            g_colCommands.Add CInt(colSubNodes("ID").Value), colSubNodes("Trigger").Value, colSubNodes("Description").Value, CInt(colSubNodes("Class").Value), CBool(colSubNodes("Enabled").Value)
220:        Next
    
222:        On Error GoTo Err
    
224:        objXML.Clear
    
226:        Set objNode = Nothing
227:        Set colSubNodes = Nothing
228:        Set colNodes = Nothing

        '-----------------------------------------
        ' Unload the default Commands by ID in case some names are re-used.
        ' removing commands by name would render changes made in gui useless
        '-----------------------------------------
235:        For Each objCmd In g_colCommands
236:            If objCmd.ID < 51 Then g_colCommands.Remove (objCmd.Name)
237:        Next
        
239:        Set objCmd = Nothing
        'If g_colCommands.Exists("ban") Then g_colCommands.Remove ("ban")
        'If g_colCommands.Exists("banip") Then g_colCommands.Remove ("banip")
        'If g_colCommands.Exists("banuser") Then g_colCommands.Remove ("banuser")
        'If g_colCommands.Exists("close") Then g_colCommands.Remove ("close")
        'If g_colCommands.Exists("info") Then g_colCommands.Remove ("info")
        'If g_colCommands.Exists("iplist") Then g_colCommands.Remove ("iplist")
        'If g_colCommands.Exists("listbanip") Then g_colCommands.Remove ("listbanip")
        'If g_colCommands.Exists("ipscan") Then g_colCommands.Remove ("ipscan")
        'If g_colCommands.Exists("listbanuser") Then g_colCommands.Remove ("listbanuser")
        'If g_colCommands.Exists("unbanip") Then g_colCommands.Remove ("unbanip")
        'If g_colCommands.Exists("unbanuser") Then g_colCommands.Remove ("unbanuser")
        'If g_colCommands.Exists("help") Then g_colCommands.Remove ("help")
        'If g_colCommands.Exists("language") Then g_colCommands.Remove ("language")
  
254:    End If

'---------------------------------------------------------------------------------

258:    m_arrDynaCap(0) = "Start Serving"
259:    m_arrDynaCap(1) = "Stop Serving"

'---------------------------------------------------------------------------------

'---------------------------------------------------------------------------------
    'Load Core/Reason Messages, a default must exist restart the hub if it don't
265:    strTemp = G_APPPATH & "\Settings\UsersMessages.xml"

267:    objXML.Data = g_objFileAccess.ReadFile(strTemp)
268:    objXML.Parse

270:    Set colNodes = objXML.Nodes(1).Nodes

    'Just in case...
     On Error Resume Next

275:    Set colSupported = New Collection

277:    For Each objNode In colNodes
278:        Set colSubNodes = objNode.Nodes

280:        Set m_colLangString = New Collection

282:        For Each objSubNode In colSubNodes
283:            m_colLangString.Add objSubNode.Value, objSubNode.Name
284:        Next

286:        colSupported.Add objNode.Name, objNode.Name
287:        g_colLanguages.Add m_colLangString, objNode.Name

289:        Set m_colLangString = Nothing
290:    Next

292:    g_colLanguages.Add colSupported, "Supported"

294:    On Error GoTo Err

296:    objXML.Clear
297:    Call ClearTranslations
298:    Set colSupported = Nothing
299:    Set objSubNode = Nothing
300:    Set objNode = Nothing
301:    Set colSubNodes = Nothing
302:    Set colNodes = Nothing

'---------------------------------------------------------------------------------
    'Load Interface Language
307:    On Error GoTo ErrLng

309:    If Not LenB(g_objSettings.Interface) Then
310:          g_objSettings.Interface = "English"
311:    End If

313:    strTemp = G_APPPATH & "\Languages\" & g_objSettings.Interface & ".xml"

315:    If Not g_objFileAccess.FileExists(strTemp) Then
316:        g_objSettings.Interface = "English"
317:        If Not g_objFileAccess.FileExists(G_APPPATH & "\Languages\English.xml") Then
                'Create defaut language file .. if is not found
319:            CreateEGLanguageXML
320:        End If
321:    End If

322:    objXML.Data = g_objFileAccess.ReadFile(strTemp)
323:    objXML.Parse

325:    Set colNodes = objXML.Nodes(1).Nodes

        'Just in case...
328:    On Error Resume Next
329:    X = 0
            
331:        For Each objNode In colNodes
332:            Set colSubNodes = objNode.Nodes
333:
334:            Select Case objNode.Name
                    Case "DynamicCaptions"
335:                    For Each objSubNode In colSubNodes
336:                        m_arrDynaCap(X) = objSubNode.Value
337:                        X = X + 1
338:                    Next
                Case "Texts"
340:                    For Each objSubNode In colSubNodes
341:                        TranslateTexts objSubNode.Name, objSubNode.Value
342:                    Next
                Case "TabSCaption"
343:                    For Each objSubNode In colSubNodes
344:                        TranslateTabSCaption objSubNode.Name, objSubNode.Value
345:                    Next
                Case "ListView"
347:                    For Each objSubNode In colSubNodes
348:                        TranslateListViewCaption objSubNode.Name, objSubNode.Value
349:                    Next
                Case "Captions"
350:                    For Each objSubNode In colSubNodes
351:                        TranslateCtrlCaption objSubNode.Name, objSubNode.Value
352:                    Next
                Case "ToolTips"
354:                    For Each objSubNode In colSubNodes
355:                        TranslateCtrlToolTip objSubNode.Name, objSubNode.Value
356:                    Next
                Case "TagsHelp"
358:                    For Each objSubNode In colSubNodes
359:                        m_arrTagRules(objSubNode.Name) = objSubNode.Value
360:                    Next
                Case "HubStringDef"
362:                    For Each objSubNode In colSubNodes
                            'g_objFileAccess.AppendFile G_APPPATH & "\debug.txt", "Name: " & objSubNode.Name & " Value: " & objSubNode.Value
364:                        g_colMessages(CStr(objSubNode.Name)) = CStr(objSubNode.Value)
365:                    Next
                Case "ToolTips"
367:                    For Each objSubNode In colSubNodes
368:                        TranslateTexts objSubNode.Name, objSubNode.Value
369:                    Next
370:            End Select
371:       Next
    
373:    On Error GoTo ErrLng
    
375:    objXML.Clear
    
377:    Set objSubNode = Nothing
378:    Set objNode = Nothing
379:    Set colSubNodes = Nothing
380:    Set colNodes = Nothing

382:    txtTagRules.Text = Replace(m_arrTagRules(10), "%[LF]", vbNewLine)

384:    If m_blnServing = False Then
385:        cmdButton(1).Caption = m_arrDynaCap(0)
386:    Else
387:        cmdButton(1).Caption = m_arrDynaCap(1)
388:    End If

        'Fill Language ComboBox
391:    cmbInterface.Clear


394:    arrTmp = g_objFileAccess.ListFiles(G_APPPATH & "\Languages\*.xml")

395:    For X = 0 To UBound(arrTmp)
396:        If (g_objFunctions.AfterLast(arrTmp(X), ".") = "xml") And (InStr(arrTmp(X), "_User") = 0) Then
397:            cmbInterface.AddItem (g_objFunctions.BeforeLast(arrTmp(X), ".xml"))
398:        End If
399:    Next
        
401:    cmbInterface.Text = g_objSettings.Interface


404: Exit Sub
405:
Err:
407:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.LoadSettings()|" & strTemp & "|"
408:    Resume Next
409: Exit Sub
410:
ErrLng: ' if error not because language file
412:
413:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.LoadSettings[Interface Language]()|" & strTemp & "|"
414:    CreateEGLanguageXML
415:    g_objSettings.Interface = "English"
416:    Resume Next
End Sub
Public Sub SaveSettings()
1:    Dim intFF       As Integer
2:    Dim strTemp     As String
3:    Dim varLoop     As Variant
4:    Dim objCommand  As clsCommand
5:    Dim objTag      As clsTag
6:    Dim objTB       As clsTempBan
7:    Dim i           As Integer
8:    Dim lvwItems    As ListItems
     'Now before I get any emails about all the file appending, read this
     'Using string concation (&) is several times slower than using this append method
     '& reallocates the string in the memory each time it's used, while this method does not (well on a much smaller scale)
12:    On Error GoTo Err

14:    strTemp = G_APPPATH & "\Settings\PTDCH.xml"

    'If the settings file exists, delete it
17:    If g_objFileAccess.FileExists(strTemp) Then g_objFileAccess.DeleteFile strTemp

19:    Call SaveFormSize ' save form position now

21:    intFF = FreeFile
    'Append to PTDCH.xml
23:    Open strTemp For Append As intFF

25:    Print #intFF, "<Settings Version=""" & vbVersion & """>"
27:    Print #intFF, vbTab & "<String>"
28:        Print #intFF, vbTab & vbTab & "<frmHubPosition>" & g_objSettings.frmHubPosition & "</frmHubPosition>"
29:        Print #intFF, vbTab & vbTab & "<HubName>" & XMLEscape(g_objSettings.HubName) & "</HubName>"
30:        Print #intFF, vbTab & vbTab & "<HubDesc>" & XMLEscape(g_objSettings.HubDesc) & "</HubDesc>"
31:        Print #intFF, vbTab & vbTab & "<HubIP>" & g_objSettings.HubIP & "</HubIP>"
32:        Print #intFF, vbTab & vbTab & "<Ports>" & g_objSettings.Ports & "</Ports>"
33:        Print #intFF, vbTab & vbTab & "<RegisterIP>" & g_objSettings.RegisterIP & "</RegisterIP>"
        '-------------NEW REDIRECT ADDRESSES-----------------------------------------------------------------------------------------------
35:        Print #intFF, vbTab & vbTab & "<ForMinShareRedirectAddress>" & g_objSettings.ForMinShareRedirectAddress & "</ForMinShareRedirectAddress>"
36:        Print #intFF, vbTab & vbTab & "<ForMaxShareRedirectAddress>" & g_objSettings.ForMaxShareRedirectAddress & "</ForMaxShareRedirectAddress>"
37:        Print #intFF, vbTab & vbTab & "<ForMinSlotsRedirectAddress>" & g_objSettings.ForMinSlotsRedirectAddress & "</ForMinSlotsRedirectAddress>"
38:        Print #intFF, vbTab & vbTab & "<ForMaxSlotsRedirectAddress>" & g_objSettings.ForMaxSlotsRedirectAddress & "</ForMaxSlotsRedirectAddress>"
39:        Print #intFF, vbTab & vbTab & "<ForMaxHubsRedirectAddress>" & g_objSettings.ForMaxHubsRedirectAddress & "</ForMaxHubsRedirectAddress>"
40:        Print #intFF, vbTab & vbTab & "<ForNoTagRedirectAddress>" & g_objSettings.ForNoTagRedirectAddress & "</ForNoTagRedirectAddress>"
41:        Print #intFF, vbTab & vbTab & "<ForTooOldDcppRedirectAddress>" & g_objSettings.ForTooOldDcppRedirectAddress & "</ForTooOldDcppRedirectAddress>"
42:        Print #intFF, vbTab & vbTab & "<ForTooOldNMDCRedirectAddress>" & g_objSettings.ForTooOldNMDCRedirectAddress & "</ForTooOldNMDCRedirectAddress>"
43:        Print #intFF, vbTab & vbTab & "<ForSlotPerHubRedirectAddress>" & g_objSettings.ForSlotPerHubRedirectAddress & "</ForSlotPerHubRedirectAddress>"
44:        Print #intFF, vbTab & vbTab & "<ForBWPerSlotRedirectAddress>" & g_objSettings.ForBWPerSlotRedirectAddress & "</ForBWPerSlotRedirectAddress>"
45:        Print #intFF, vbTab & vbTab & "<ForFakeShareRedirectAddress>" & g_objSettings.ForFakeShareRedirectAddress & "</ForFakeShareRedirectAddress>"
46:        Print #intFF, vbTab & vbTab & "<ForFakeTagRedirectAddress>" & g_objSettings.ForFakeTagRedirectAddress & "</ForFakeTagRedirectAddress>"
47:        Print #intFF, vbTab & vbTab & "<ForPasModeRedirectAddress>" & g_objSettings.ForPasModeRedirectAddress & "</ForPasModeRedirectAddress>"
        '----------------STOP HERE----------------------------------------------------------------------------------------------------------
49:        Print #intFF, vbTab & vbTab & "<RedirectAddress>" & g_objSettings.RedirectAddress & "</RedirectAddress>"
50:        Print #intFF, vbTab & vbTab & "<BotName>" & g_objSettings.BotName & "</BotName>"
51:        Print #intFF, vbTab & vbTab & "<OpChatName>" & g_objSettings.OpChatName & "</OpChatName>"
52:        Print #intFF, vbTab & vbTab & "<JoinMsg>" & XMLEscape(g_objSettings.JoinMsg) & "</JoinMsg>"
53:        Print #intFF, vbTab & vbTab & "<CSeperator>" & g_objSettings.CSeperator & "</CSeperator>"
54:        Print #intFF, vbTab & vbTab & "<HubPassword>" & XMLEscape(g_objSettings.HubPassword) & "</HubPassword>"
55:        Print #intFF, vbTab & vbTab & "<MassMessage>" & XMLEscape(g_objSettings.MassMessage) & "</MassMessage>"
56:        Print #intFF, vbTab & vbTab & "<OpMassMessage>" & XMLEscape(g_objSettings.OpMassMessage) & "</OpMassMessage>"
57:        Print #intFF, vbTab & vbTab & "<UnRegMassMessage>" & XMLEscape(g_objSettings.UnRegMassMessage) & "</UnRegMassMessage>"
        ' ------------------------ NEW INTERFACE LANGUAGE ------------------------
58:        Print #intFF, vbTab & vbTab & "<Interface>" & g_objSettings.Interface & "</Interface>"
        ' ----------------------- NEW INTERFACE LANGUAGE END ----------------------
60:        Print #intFF, vbTab & vbTab & "<HammeringRd>" & g_objSettings.HammeringRd & "</HammeringRd>"
61:        Print #intFF, vbTab & "<NoIPDNS1>" & g_objSettings.NoIPDNS1 & "</NoIPDNS1>"
62:        Print #intFF, vbTab & "<NoIPDNS2>" & g_objSettings.NoIPDNS2 & "</NoIPDNS2>"
63:        Print #intFF, vbTab & "<NoIPDNS3>" & g_objSettings.NoIPDNS3 & "</NoIPDNS3>"
64:        Print #intFF, vbTab & "<NoIPDNS4>" & g_objSettings.NoIPDNS4 & "</NoIPDNS4>"
65:        Print #intFF, vbTab & "<NoIPUser>" & XMLEscape(g_objSettings.NoIPUser) & "</NoIPUser>"
66:        Print #intFF, vbTab & "<NoIPPass>" & XMLEscape(g_objSettings.NoIPPass) & "</NoIPPass>"
67:        Print #intFF, vbTab & "<DynDNS1>" & g_objSettings.DynDNS1 & "</DynDNS1>"
68:        Print #intFF, vbTab & "<DynDNS2>" & g_objSettings.DynDNS2 & "</DynDNS2>"
69:        Print #intFF, vbTab & "<DynDNS3>" & g_objSettings.DynDNS3 & "</DynDNS3>"
70:        Print #intFF, vbTab & "<DynDNS4>" & g_objSettings.DynDNS4 & "</DynDNS4>"
71:        Print #intFF, vbTab & "<DynDNSUser>" & XMLEscape(g_objSettings.DynDNSUser) & "</DynDNSUser>"
72:        Print #intFF, vbTab & "<DynDNSPass>" & XMLEscape(g_objSettings.DynDNSPass) & "</DynDNSPass>"
73:    Print #intFF, vbTab & "</String>"

77:    Print #intFF, vbTab & "<Boolean>"
78:        Print #intFF, vbTab & vbTab & "<DenySocks5>" & g_objSettings.DenySocks5 & "</DenySocks5>"
79:        Print #intFF, vbTab & vbTab & "<DenyPassive>" & g_objSettings.DenyPassive & "</DenyPassive>"
80:        Print #intFF, vbTab & vbTab & "<AutoCheckUpdate>" & g_objSettings.AutoCheckUpdate & "</AutoCheckUpdate>"
81:        Print #intFF, vbTab & vbTab & "<AutoKickMLDC>" & g_objSettings.AutoKickMLDC & "</AutoKickMLDC>"
82:        Print #intFF, vbTab & vbTab & "<AutoRegister>" & g_objSettings.AutoRegister & "</AutoRegister>"
83:        Print #intFF, vbTab & vbTab & "<AutoRedirect>" & g_objSettings.AutoRedirect & "</AutoRedirect>"
84:        Print #intFF, vbTab & vbTab & "<AutoRedirectFull>" & g_objSettings.AutoRedirectFull & "</AutoRedirectFull>"
85:        Print #intFF, vbTab & vbTab & "<AutoRedirectNonReg>" & g_objSettings.AutoRedirectNonReg & "</AutoRedirectNonReg>"
86:        Print #intFF, vbTab & vbTab & "<AutoRedirectFullNonReg>" & g_objSettings.AutoRedirectFullNonReg & "</AutoRedirectFullNonReg>"
87:        Print #intFF, vbTab & vbTab & "<AutoRedirectFullNonOps>" & g_objSettings.AutoRedirectFullNonOps & "</AutoRedirectFullNonOps>"
88:        Print #intFF, vbTab & vbTab & "<AutoStart>" & g_objSettings.AutoStart & "</AutoStart>"
89:        Print #intFF, vbTab & vbTab & "<CompactDBOnExit>" & g_objSettings.CompactDBOnExit & "</CompactDBOnExit>"
90:        Print #intFF, vbTab & vbTab & "<ConfirmExit>" & g_objSettings.ConfirmExit & "</ConfirmExit>"
91:        Print #intFF, vbTab & vbTab & "<DCValidateTags>" & g_objSettings.DCValidateTags & "</DCValidateTags>"
92:        Print #intFF, vbTab & vbTab & "<DCIncludeOPed>" & g_objSettings.DCIncludeOPed & "</DCIncludeOPed>"
93:        Print #intFF, vbTab & vbTab & "<OPBypass>" & g_objSettings.OPBypass & "</OPBypass>"
94:        Print #intFF, vbTab & vbTab & "<PreloadWinsocks>" & g_objSettings.PreloadWinsocks & "</PreloadWinsocks>"
95:        Print #intFF, vbTab & vbTab & "<SendMessageAFK>" & g_objSettings.SendMessageAFK & "</SendMessageAFK>"
96:        Print #intFF, vbTab & vbTab & "<RegOnly>" & g_objSettings.RegOnly & "</RegOnly>"
97:        Print #intFF, vbTab & vbTab & "<MentoringSystem>" & g_objSettings.MentoringSystem & "</MentoringSystem>"
98:        Print #intFF, vbTab & vbTab & "<PreventSearchBots>" & g_objSettings.PreventSearchBots & "</PreventSearchBots>"
99:        Print #intFF, vbTab & vbTab & "<DescriptiveBanMsg>" & g_objSettings.DescriptiveBanMsg & "</DescriptiveBanMsg>"
100:        Print #intFF, vbTab & vbTab & "<UseOpChat>" & g_objSettings.UseOpChat & "</UseOpChat>"
101:        Print #intFF, vbTab & vbTab & "<UseBotName>" & g_objSettings.UseBotName & "</UseBotName>"
102:        Print #intFF, vbTab & vbTab & "<Passive>" & g_objSettings.Passive & "</Passive>"
103:        Print #intFF, vbTab & vbTab & "<DynUpdate>" & g_objSettings.DynUpdate & "</DynUpdate>"
104:        Print #intFF, vbTab & vbTab & "<DynDNSUpdateEna>" & g_objSettings.DynDNSUpdateEna & "</DynDNSUpdateEna>"
105:        Print #intFF, vbTab & vbTab & "<NoIPUpdateEna>" & g_objSettings.NoIPUpdateEna & "</NoIPUpdateEna>"
106:        Print #intFF, vbTab & vbTab & "<NoIPUpdateStartUp>" & g_objSettings.NoIPUpdateStartUp & "</NoIPUpdateStartUp>"
               
        '-------------NEW REDIRECT CHECK BOXES-----------------------------------------------------------------------------------------------
109:        Print #intFF, vbTab & vbTab & "<RedirectFMinShare>" & g_objSettings.RedirectFMinShare & "</RedirectFMinShare>"
110:        Print #intFF, vbTab & vbTab & "<RedirectFMaxShare>" & g_objSettings.RedirectFMaxShare & "</RedirectFMaxShare>"
111:        Print #intFF, vbTab & vbTab & "<RedirectFMinSlots>" & g_objSettings.RedirectFMinSlots & "</RedirectFMinSlots>"
112:        Print #intFF, vbTab & vbTab & "<RedirectFMaxSlots>" & g_objSettings.RedirectFMaxSlots & "</RedirectFMaxSlots>"
113:        Print #intFF, vbTab & vbTab & "<RedirectFMaxHubs>" & g_objSettings.RedirectFMaxHubs & "</RedirectFMaxHubs>"
114:        Print #intFF, vbTab & vbTab & "<RedirectFSlotPerHub>" & g_objSettings.RedirectFSlotPerHub & "</RedirectFSlotPerHub>"
115:        Print #intFF, vbTab & vbTab & "<RedirectFNoTag>" & g_objSettings.RedirectFNoTag & "</RedirectFNoTag>"
116:        Print #intFF, vbTab & vbTab & "<RedirectFTooOldDCpp>" & g_objSettings.RedirectFTooOldDCpp & "</RedirectFTooOldDCpp>"
117:        Print #intFF, vbTab & vbTab & "<RedirectFTooOldNMDC>" & g_objSettings.RedirectFTooOldNMDC & "</RedirectFTooOldNMDC>"
118:        Print #intFF, vbTab & vbTab & "<RedirectFBWPerSlot>" & g_objSettings.RedirectFBWPerSlot & "</RedirectFBWPerSlot>"
119:        Print #intFF, vbTab & vbTab & "<RedirectFFakeShare>" & g_objSettings.RedirectFFakeShare & "</RedirectFFakeShare>"
120:        Print #intFF, vbTab & vbTab & "<RedirectFFakeTag>" & g_objSettings.RedirectFFakeTag & "</RedirectFFakeTag>"
121:        Print #intFF, vbTab & vbTab & "<RedirectFPasMode>" & g_objSettings.RedirectFPasMode & "</RedirectFPasMode>"
        '----------------STOP HERE----------------------------------------------------------------------------------------------------------
122:        Print #intFF, vbTab & vbTab & "<FilterCPrefix>" & g_objSettings.FilterCPrefix & "</FilterCPrefix>"
123:        Print #intFF, vbTab & vbTab & "<EnabledCommands>" & g_objSettings.EnabledCommands & "</EnabledCommands>"
124:        Print #intFF, vbTab & vbTab & "<ScriptSafeMode>" & g_objSettings.ScriptSafeMode & "</ScriptSafeMode>"
125:        Print #intFF, vbTab & vbTab & "<StartMinimized>" & g_objSettings.StartMinimized & "</StartMinimized>"
126:        Print #intFF, vbTab & vbTab & "<SendMsgAsPrivate>" & g_objSettings.SendMsgAsPrivate & "</SendMsgAsPrivate>"
127:        Print #intFF, vbTab & vbTab & "<PasswordMode>" & g_objSettings.PasswordMode & "</PasswordMode>"
128:        Print #intFF, vbTab & vbTab & "<WordWrap>" & g_objSettings.WordWrap & "</WordWrap>"
129:        Print #intFF, vbTab & vbTab & "<DenyNoTag>" & g_objSettings.DenyNoTag & "</DenyNoTag>"
130:        Print #intFF, vbTab & vbTab & "<HideFadeImg>" & g_objSettings.HideFadeImg & "</HideFadeImg>"
131:        Print #intFF, vbTab & vbTab & "<CheckFakeShare>" & g_objSettings.CheckFakeShare & "</CheckFakeShare>"
132:        Print #intFF, vbTab & vbTab & "<PreventGuessPass>" & g_objSettings.PreventGuessPass & "</PreventGuessPass>"
133:        Print #intFF, vbTab & vbTab & "<EnableFloodWall>" & g_objSettings.EnableFloodWall & "</EnableFloodWall>"
134:        Print #intFF, vbTab & vbTab & "<RedirectFGP>" & g_objSettings.RedirectFGP & "</RedirectFGP>"
135:        Print #intFF, vbTab & vbTab & "<OpsCanRedirect>" & g_objSettings.OpsCanRedirect & "</OpsCanRedirect>"
136:        Print #intFF, vbTab & vbTab & "<ChatOnly>" & g_objSettings.ChatOnly & "</ChatOnly>"
137:        Print #intFF, vbTab & vbTab & "<VIPUseOpChat>" & g_objSettings.VIPUseOpChat & "</VIPUseOpChat>"
138:        Print #intFF, vbTab & vbTab & "<MinClsSearchSend>" & g_objSettings.MinClsSearchSend & "</MinClsSearchSend>"
139:        Print #intFF, vbTab & vbTab & "<MinClsConnectSend>" & g_objSettings.MinClsConnectSend & "</MinClsConnectSend>"
140:        Print #intFF, vbTab & vbTab & "<MinimizeTray>" & g_objSettings.MinimizeTray & "</MinimizeTray>"
141:        Print #intFF, vbTab & vbTab & "<HideMyinfos>" & g_objSettings.HideMyinfos & "</HideMyinfos>"
142:        Print #intFF, vbTab & vbTab & "<ACOClients>" & g_objSettings.ACOClients & "</ACOClients>"
143:        Print #intFF, vbTab & vbTab & "<EnabledScheduler>" & g_objSettings.EnabledScheduler & "</EnabledScheduler>"
144:        Print #intFF, vbTab & vbTab & "<PriorityBl>" & g_objSettings.PriorityBl & "</PriorityBl>"
145:        Print #intFF, vbTab & vbTab & "<PopUpNewReg>" & g_objSettings.PopUpNewReg & "</PopUpNewReg>"
146:        Print #intFF, vbTab & vbTab & "<PopUpOpConected>" & g_objSettings.PopUpOpConected & "</PopUpOpConected>"
147:        Print #intFF, vbTab & vbTab & "<PopUpOpDisconected>" & g_objSettings.PopUpOpDisconected & "</PopUpOpDisconected>"
148:        Print #intFF, vbTab & vbTab & "<PopUpUserKick>" & g_objSettings.PopUpUserKick & "</PopUpUserKick>"
149:        Print #intFF, vbTab & vbTab & "<PopUpUserBaned>" & g_objSettings.PopUpUserBaned & "</PopUpUserBaned>"
150:        Print #intFF, vbTab & vbTab & "<PopUpUserRedirected>" & g_objSettings.PopUpUserRedirected & "</PopUpUserRedirected>"
151:        Print #intFF, vbTab & vbTab & "<PopUpStartedServing>" & g_objSettings.PopUpStartedServing & "</PopUpStartedServing>"
152:        Print #intFF, vbTab & vbTab & "<PopUpUserBaned>" & g_objSettings.PopUpUserBaned & "</PopUpUserBaned>"
153:        Print #intFF, vbTab & vbTab & "<PopUpUserRedirected>" & g_objSettings.PopUpUserRedirected & "</PopUpUserRedirected>"
154:        Print #intFF, vbTab & vbTab & "<PopUpStopedServing>" & g_objSettings.PopUpStopedServing & "</PopUpStopedServing>"
155:        Print #intFF, vbTab & vbTab & "<MoveForm>" & g_objSettings.MoveForm & "</MoveForm>"
156:        Print #intFF, vbTab & vbTab & "<MagneticWin>" & g_objSettings.MagneticWin & "</MagneticWin>"
157:        Print #intFF, vbTab & vbTab & "<StartWin>" & g_objSettings.StartWin & "</StartWin>"
158:        Print #intFF, vbTab & vbTab & "<blSkin>" & g_objSettings.blSkin & "</blSkin>"
159:        Print #intFF, vbTab & vbTab & "<RndSkin>" & g_objSettings.RndSkin & "</RndSkin>"
160:        Print #intFF, vbTab & vbTab & "<Plugins>" & g_objSettings.Plugins & "</Plugins>"
161:        Print #intFF, vbTab & "</Boolean>"

167:        Print #intFF, vbTab & "<Byte>"
168:        Print #intFF, vbTab & vbTab & "<DCMaxHubs>" & g_objSettings.DCMaxHubs & "</DCMaxHubs>"
169:        Print #intFF, vbTab & vbTab & "<DCOSlots>" & g_objSettings.DCOSlots & "</DCOSlots>"
170:        Print #intFF, vbTab & vbTab & "<MinSlots>" & g_objSettings.MinSlots & "</MinSlots>"
171:        Print #intFF, vbTab & vbTab & "<MaxSlots>" & g_objSettings.MaxSlots & "</MaxSlots>"
172:        Print #intFF, vbTab & vbTab & "<MinShareSize>" & g_objSettings.MinShareSize & "</MinShareSize>"
173:        Print #intFF, vbTab & vbTab & "<MaxShareSize>" & g_objSettings.MaxShareSize & "</MaxShareSize>"
174:        Print #intFF, vbTab & vbTab & "<CPrefix>" & g_objSettings.CPrefix & "</CPrefix>"
175:        Print #intFF, vbTab & vbTab & "<DCOSpeed>" & g_objSettings.DCOSpeed & "</DCOSpeed>"
176:        Print #intFF, vbTab & vbTab & "<SendJoinMsg>" & g_objSettings.SendJoinMsg & "</SendJoinMsg>"
177:        Print #intFF, vbTab & vbTab & "<MaxPassAttempts>" & g_objSettings.MaxPassAttempts & "</MaxPassAttempts>"
178:        Print #intFF, vbTab & vbTab & "<FWMyINFO>" & g_objSettings.FWMyINFO & "</FWMyINFO>"
179:        Print #intFF, vbTab & vbTab & "<FWGetNickList>" & g_objSettings.FWGetNickList & "</FWGetNickList>"
180:        Print #intFF, vbTab & vbTab & "<FWActiveSearch>" & g_objSettings.FWActiveSearch & "</FWActiveSearch>"
181:        Print #intFF, vbTab & vbTab & "<FWPassiveSearch>" & g_objSettings.FWPassiveSearch & "</FWPassiveSearch>"
182:        Print #intFF, vbTab & vbTab & "<MinMyinfoFakeCls>" & g_objSettings.MinMyinfoFakeCls & "</MinMyinfoFakeCls>"
        'TheNOP svn 159 , hidden, no interface setting.
184:        Print #intFF, vbTab & vbTab & "<FWMainchat>" & g_objSettings.FWMainChat & "</FWMainchat>"
        'Print #intFF, vbTab & vbTab & "<FWGlobal>" & g_objSettings.FWGlobal & "</FWGlobal>"
186:    Print #intFF, vbTab & "</Byte>"

188:    Print #intFF, vbTab & "<Integer>"
189:        Print #intFF, vbTab & vbTab & "<MinPassiveSearchLen>" & g_objSettings.MinPassiveSearchLen & "</MinPassiveSearchLen>"
190:        Print #intFF, vbTab & vbTab & "<FWInterval>" & g_objSettings.FWInterval & "</FWInterval>"
191:        Print #intFF, vbTab & vbTab & "<FWBanLength>" & g_objSettings.FWBanLength & "</FWBanLength>"
192:        Print #intFF, vbTab & vbTab & "<MinConnectCls>" & g_objSettings.MinConnectCls & "</MinConnectCls>"
193:        Print #intFF, vbTab & vbTab & "<MinSearchCls>" & g_objSettings.MinSearchCls & "</MinSearchCls>"
194:        Print #intFF, vbTab & vbTab & "<ZLINELENGHT>" & g_objSettings.ZLINELENGHT & "</ZLINELENGHT>"
195:        Print #intFF, vbTab & vbTab & "<PriorityVal>" & g_objSettings.PriorityVal & "</PriorityVal>"

198:    Print #intFF, vbTab & "</Integer>"

200:    Print #intFF, vbTab & "<Long>"
201:        Print #intFF, vbTab & vbTab & "<MaxUsers>" & g_objSettings.MaxUsers & "</MaxUsers>"
202:        Print #intFF, vbTab & vbTab & "<DefaultBanTime>" & g_objSettings.DefaultBanTime & "</DefaultBanTime>"
203:        Print #intFF, vbTab & vbTab & "<ScriptTimeout>" & g_objSettings.ScriptTimeout & "</ScriptTimeout>"
204:        Print #intFF, vbTab & vbTab & "<MaxMessageLen>" & g_objSettings.MaxMessageLen & "</MaxMessageLen>"
205:        Print #intFF, vbTab & vbTab & "<DataFragmentLen>" & g_objSettings.DataFragmentLen & "</DataFragmentLen>"
206:        Print #intFF, vbTab & vbTab & "<ConDropInterval>" & g_objSettings.ConDropInterval & "</ConDropInterval>"
207:        Print #intFF, vbTab & vbTab & "<FWDropMsgInterval>" & g_objSettings.FWDropMsgInterval & "</FWDropMsgInterval>"
208:        Print #intFF, vbTab & vbTab & "<lngSkin>" & g_objSettings.lngSkin & "</lngSkin>"

210:    Print #intFF, vbTab & "</Long>"

212:    Print #intFF, vbTab & "<Double>"
213:        Print #intFF, vbTab & vbTab & "<IMinShare>" & g_objSettings.IMinShare & "</IMinShare>"
214:        Print #intFF, vbTab & vbTab & "<IMaxShare>" & g_objSettings.IMaxShare & "</IMaxShare>"
215:        Print #intFF, vbTab & vbTab & "<DCSlotsPerHub>" & g_objSettings.DCSlotsPerHub & "</DCSlotsPerHub>"
216:        Print #intFF, vbTab & vbTab & "<DCBandPerSlot>" & g_objSettings.DCBandPerSlot & "</DCBandPerSlot>"
217:        Print #intFF, vbTab & vbTab & "<DCMinVersion>" & g_objSettings.DCMinVersion & "</DCMinVersion>"
218:        Print #intFF, vbTab & vbTab & "<NMDCMinVersion>" & g_objSettings.NMDCMinVersion & "</NMDCMinVersion>"
219:    Print #intFF, vbTab & "</Double>"

221:    Print #intFF, vbTab & "<Tags>"
222:        For Each objTag In m_colTags
223:            Print #intFF, vbTab & vbTab & "<Tag ";
224:             Print #intFF, "Name=""" & objTag.Name & """ ";
225:             Print #intFF, "/>"
226:        Next
227:    Print #intFF, vbTab & "</Tags>"

229:    Print #intFF, "</Settings>";

231:    Close intFF

'---------------------------------------------------------------------------------
    'Perm IP bans
235:    strTemp = G_APPPATH & "\Settings\PermIPBans.xml"

237:    If g_objFileAccess.FileExists(strTemp) Then g_objFileAccess.DeleteFile strTemp

239:    intFF = FreeFile

    'Append to file
242:    Open strTemp For Append As intFF

244:    Print #intFF, "<PermIPBans>"

246:    For Each varLoop In g_objIPBans.PermItems
247:        Print #intFF, vbTab & "<IP>" & varLoop & "</IP>"
248:    Next

250:    Print #intFF, "</PermIPBans>";

252:    Close intFF

'---------------------------------------------------------------------------------
    'Temp IP Bans
256:    strTemp = G_APPPATH & "\Settings\TempIPBans.xml"

258:    If g_objFileAccess.FileExists(strTemp) Then g_objFileAccess.DeleteFile strTemp

260:    intFF = FreeFile

    'Append to file
263:    Open strTemp For Append As intFF

265:    Print #intFF, "<TempIPBans>"

267:    For Each objTB In g_objIPBans.TempItems
268:        If DateDiff("n", Now, objTB.ExpDate) > 0 Then _
            Print #intFF, vbTab & "<IP Date=""" & objTB.ExpDate & """>" & objTB.IP & "</IP>"
270:    Next

272:    Print #intFF, "</TempIPBans>";

274:    Close intFF

'---------------------------------------------------------------------------------
    'Commands
278:    strTemp = G_APPPATH & "\Settings\Commands.xml"

280:    If g_objFileAccess.FileExists(strTemp) Then g_objFileAccess.DeleteFile strTemp

282:    intFF = FreeFile

284:    Open strTemp For Append As intFF

286:    Print #intFF, "<Commands>"

    'Loop through command collection
289:    For Each objCommand In g_colCommands
290:        Print #intFF, vbTab & "<Command ID=""" & objCommand.ID & """ Trigger=""" & objCommand.Name & """ Class=""" & objCommand.Class & """ Enabled=""" & objCommand.Enabled & """ Description=""" & XMLEscape(objCommand.Description) & """ />"
291:    Next

293:    Print #intFF, "</Commands>";

295:    Close intFF

'---------------------------------------------------------------------------------
298:    strTemp = G_APPPATH & "\Settings\UsersMessages.xml"
    
    'Must exist but it must not be replace
301:    If Not g_objFileAccess.FileExists(strTemp) Then

303:        intFF = FreeFile
    
305:        Open strTemp For Append As intFF
    
307:        Print #intFF, "<Languages>"
308:        Print #intFF, vbTab & "<En>"
309:        Print #intFF, vbTab & vbTab & "<InternationalName>English</InternationalName>"
310:        Print #intFF, vbTab & vbTab & "<NationalName>English</NationalName>"
311:        Print #intFF, vbTab & vbTab & "<LoggedIn>Logged in.</LoggedIn>"
312:        Print #intFF, vbTab & vbTab & "<ChatMode>Can&apos;t connect because user %[user] is in chat only mode.</ChatMode>"
313:        Print #intFF, vbTab & vbTab & "<MaxHubs>You are connected to too many hubs. %[maxhubs] hubs max. Disconnect from some and reconnect.</MaxHubs>"
314:        Print #intFF, vbTab & vbTab & "<MinSlots>You do not have enough slots open. %[minslots] slot(s) min.</MinSlots>"
315:        Print #intFF, vbTab & vbTab & "<MaxSlots>You have too many slots open. %[maxslots] slots max.</MaxSlots>"
316:        Print #intFF, vbTab & vbTab & "<HSRatio>You have not met the hub per slot ratio. %[hsratio] slot per hub min.</HSRatio>"
317:        Print #intFF, vbTab & vbTab & "<BSRatio>You have not met the bandwidth (in KB/s) per slot ratio (as measured by the limiter you are using) %[bsratio]KB/s per slot.</BSRatio>"
318:        Print #intFF, vbTab & vbTab & "<MaxShare>You are sharing more than maximum allowed amount. %[maxshare] max.</MaxShare>"
319:        Print #intFF, vbTab & vbTab & "<MinShare>You have not met the minimum share. %[minshare] minimum.</MinShare>"
320:        Print #intFF, vbTab & vbTab & "<DCppMinVersion>You are using an outdated DC++ client. Please goto http://dcplusplus.sourceforge.net/ and update it.</DCppMinVersion>"
321:        Print #intFF, vbTab & vbTab & "<NMDCMinVersion>You are using an outdated NMDC client. Please goto http://www.neo-modus.com/ and update it. If you are using another client, please change the version setting.</NMDCMinVersion>"
322:        Print #intFF, vbTab & vbTab & "<DenyNoTag>You do not have an identification tag for your client (ie &lt;++, &lt;DC, etc). Please enable your tag, if possible.</DenyNoTag>"
323:        Print #intFF, vbTab & vbTab & "<Faker>You are suspected of trying to cheat. Goodbye.</Faker>"
324:        Print #intFF, vbTab & vbTab & "<Socks5>Socks5 mode not allowed.</Socks5>"
325:        Print #intFF, vbTab & vbTab & "<PassiveMode>Passive mode not allowed.</PassiveMode>"
326:        Print #intFF, vbTab & vbTab & "<PassLength>Passwords cannot be longer than 20 characters.</PassLength>"
327:        Print #intFF, vbTab & vbTab & "<NickLength>Your nickname cannot be longer than 40 characters.</NickLength>"
328:        Print #intFF, vbTab & vbTab & "<NickTaken>Your nickname has already been taken.</NickTaken>"
329:        Print #intFF, vbTab & vbTab & "<ChrInNick>Your nickname has an invalid character &quot; &apos; / or a (space).</ChrInNick>"
330:        Print #intFF, vbTab & vbTab & "<WrongPassRedir>The password was incorrect. You are being redirected to </WrongPassRedir>"
331:        Print #intFF, vbTab & vbTab & "<WrongPass>The password was incorrect.</WrongPass>"
332:        Print #intFF, vbTab & vbTab & "<PassMode>This hub is running in password mode. Please supply the global password.</PassMode>"
333:        Print #intFF, vbTab & vbTab & "<RegPass>Your nickname is registered. Please supply the password.</RegPass>"
334:        Print #intFF, vbTab & vbTab & "<RedirectedBecause>You are being redirected because: </RedirectedBecause>"
335:        Print #intFF, vbTab & vbTab & "<RedirectedTo>You are being redirected to: </RedirectedTo>"
336:        Print #intFF, vbTab & vbTab & "<FullRedirTo>This hub is currently full. You are being redirected to: </FullRedirTo>"
337:        Print #intFF, vbTab & vbTab & "<Full>This hub is currently full.</Full>"
338:        Print #intFF, vbTab & vbTab & "<RegOnlyRedirTo>This hub is for registered users only. You are being redirected to: </RegOnlyRedirTo>"
339:        Print #intFF, vbTab & vbTab & "<RegOnly>This hub is for registered users only.</RegOnly>"
340:        Print #intFF, vbTab & vbTab & "<BannedBecause>You are being banned because: </BannedBecause>"
341:        Print #intFF, vbTab & vbTab & "<IPPermBan>Your IP is permanently banned.</IPPermBan>"
342:        Print #intFF, vbTab & vbTab & "<IPBanned>Your IP is banned!</IPBanned>"
343:        Print #intFF, vbTab & vbTab & "<IPTempBan>Your IP is temporarily banned for </IPTempBan>"
344:        Print #intFF, vbTab & vbTab & "<KickedBecause>You are being kicked because: </KickedBecause>"
345:        Print #intFF, vbTab & vbTab & "<KickedBy>The user, %[user], was kicked by %[op]. IP: %[ip]</KickedBy>"
346:        Print #intFF, vbTab & vbTab & "<IsKicking>%[op] is kicking %[user] because: %[reason]</IsKicking>"
347:        Print #intFF, vbTab & vbTab & "<NoCOClients>ChatOnly clients are not allowed in here.</NoCOClients>"
348:        Print #intFF, vbTab & "</En>"
349:        Print #intFF, "</Languages>"
    
351:        Close intFF
352:    End If

'---------------------------------------------------------------------------------
        'Add\Rem auto start up at windows
       
357:     If g_objSettings.StartWin Then _
              AddRegRun _
         Else RemRegRun
'---------------------------------------------------------------------------------
        'Save txtNotepad to text file
362:     g_objFileAccess.WriteFile G_APPPATH & "\Settings\notepad.txt", txtNotePad.Text

        'Save sql commands to text file
365:     g_objFileAccess.WriteFile G_APPPATH & "\Settings\bdManager.sql", m_sciSql.Text

367:     AddLog "Save settings..", 2
368:  Exit Sub
    
370:
Err:
372:  HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.SaveSettings()"
End Sub
'------------------------------------------------------------------------------
' End Setting related methods
'------------------------------------------------------------------------------

Public Sub RefreshGUI(Optional bRefreshRegBanIPs = False)

1:    Dim lng         As Long
2:    Dim objTag      As clsTag
3:    Dim lvwItems    As ListItems

      On Error Resume Next
    
      'Set text boxes
8:     For lng = 0 To txtData.UBound
9:        If lng = 18 Then _
            txtData(18).Text = ChrW$(CallByName(g_objSettings, txtData(lng).Tag, VbGet)) _
        Else _
            txtData(lng).Text = CallByName(g_objSettings, txtData(lng).Tag, VbGet)
13:    Next
    
       'Set check boxes
15:    For lng = 0 To chkData.UBound
16:        chkData(lng).Value = Abs(CallByName(g_objSettings, chkData(lng).Tag, VbGet))
17:    Next
    
       'Set scroll bars
20:    For lng = 0 To vslData.UBound
21:        vslData(lng).Value = CallByName(g_objSettings, vslData(lng).Tag, VbGet)
22:    Next
     
       'Set combo boxes
25:    cmbData(0).Text = CallByName(g_objSettings, cmbData(0).Tag, VbGet)
26:    cmbData(1).ListIndex = CallByName(g_objSettings, cmbData(1).Tag, VbGet)
27:    cmbData(2).ListIndex = CallByName(g_objSettings, cmbData(2).Tag, VbGet)

29:    lvwItems("minshare").SubItems(1) = g_objSettings.MinShareMsg
30:    lvwItems("dcppminversion").SubItems(1) = g_objSettings.DCppMinVersionMsg
31:    lvwItems("minslots").SubItems(1) = g_objSettings.MinSlotsMsg
32:    lvwItems("maxslots").SubItems(1) = g_objSettings.MaxSlotsMsg
33:    lvwItems("hsratio").SubItems(1) = g_objSettings.HSRatioMsg
34:    lvwItems("bsratio").SubItems(1) = g_objSettings.BSRatioMsg
35:    lvwItems("maxhubs").SubItems(1) = g_objSettings.MaxHubsMsg
36:    lvwItems("nmdcminversion").SubItems(1) = g_objSettings.NMDCMinVersionMsg
37:    lvwItems("denynotag").SubItems(1) = g_objSettings.DenyNoTagMsg
38:    lvwItems("maxshare").SubItems(1) = g_objSettings.MaxShareMsg
39:    lvwItems("fakeshare").SubItems(1) = g_objSettings.FakeShareMsg
40:    lvwItems("faketag").SubItems(1) = g_objSettings.FakeTagMsg
41:    lvwItems("socks5").SubItems(1) = g_objSettings.Socks5Msg
42:    lvwItems("passivemode").SubItems(1) = g_objSettings.PassiveModeMsg
43:    lvwItems("NoCOClients").SubItems(1) = g_objSettings.NoCOClientsMsg

45:    Set lvwItems = Nothing
        
       'Add tags to ListBox
48:    lstTagsEx.Clear
    
50:    For Each objTag In m_colTags
51:        lstTagsEx.AddItem objTag.Name
52:    Next
    
        'Set redirect option
55:    Select Case True
            Case g_objSettings.AutoRedirect: optRedirect(0).Value = True
            Case g_objSettings.AutoRedirectNonReg: optRedirect(1).Value = True
            Case g_objSettings.AutoRedirectFull: optRedirect(2).Value = True
            Case g_objSettings.AutoRedirectFullNonReg: optRedirect(3).Value = True
            Case g_objSettings.AutoRedirectFullNonOps: optRedirect(4).Value = True
            Case Else: optRedirect(5).Value = True
       End Select
    
        'Set join message option
65:     optJM(g_objSettings.SendJoinMsg).Value = True

67:     If bRefreshRegBanIPs Then
68:         Call DBGetRegRecord 'Refresh registered user list
69:         Call DBGetBanRecord 'Refresh banned nicknames
            'refresh IP Ban Perm/Temp
71:         Call mnuTempIPBan_Click(4)
72:         Call mnuPermIPBan_Click(4)
73:         Me.Refresh
74:     End If
        
        'Refresh all controls..this process refreshes the memory used ram, minimeze ptdch and see task manager ;-)
77:     If Me.WindowState = vbMinimized Then
78:         Dim CTL As Control
79:         Dim i As Integer
81:         For Each CTL In frmHub.Controls
82:             On Local Error Resume Next
83:             CTL.Refresh
84:             DoEvents
85:         Next
86:         Me.Visible = True
87:         DoEvents: Me.Refresh
88:         Me.Visible = False
89:    End If
90:
End Sub

'------------------------------------------------------------------------------
'Assorted support methods
'------------------------------------------------------------------------------
Private Sub UpdateFailedReg(ByRef curUser As clsUser, ByRef blnLoggedIn As Boolean)
1:    Dim objFR   As clsFailedReg
2:    Dim strKey  As String

4:    On Error GoTo Err
    'Create key
6:    strKey = LCase$(curUser.sName) & "|" & curUser.IP
    
8:    On Error Resume Next
    
    'If they logged in sucessfully, then remove failed attempts from collection
11:    If blnLoggedIn Then
12:        m_colFailedReg.Remove strKey
13:    Else
        'Get object
15:        Set objFR = m_colFailedReg(strKey)
        
17:        On Error GoTo Err
        
        'If they have never tried to log in, create object
20:        If ObjPtr(objFR) = 0 Then
21:            Set objFR = New clsFailedReg
22:            m_colFailedReg.Add objFR, strKey
23:        End If
        
        'Update check
26:        If objFR.Check(curUser) Then _
            m_colFailedReg.Remove strKey
28:    End If
    
30:    Exit Sub
    
32:
Err:
33:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.UpdateFailedReg(, " & blnLoggedIn & ")"
End Sub
Friend Function UpdateConnectAttempt(ByRef wskUser As Winsock, ByRef blnLoggedIn As Boolean) As Boolean
1:    Dim objCA   As clsConnectAttempt
2:    Dim strIP   As String
    
4:    On Error GoTo Err
    
6:    strIP = wskUser.RemoteHostIP
    
    'If logged in, then we can just
    'svn 223
10:    If Not blnLoggedIn Then
    'If blnLoggedIn Then
        'If removed we won't know if they are hammering...
        'm_colConnectAttempts.Remove strIP
    'Else
        'Attempt to retrieve object
16:        On Error Resume Next
17:        Set objCA = m_colConnectAttempts(strIP)
18:        On Error GoTo Err
        
        'If it doesn't exist, create a new one
21:        If ObjPtr(objCA) = 0 Then
22:            Set objCA = New clsConnectAttempt
23:            objCA.IP = strIP
            
25:            m_colConnectAttempts.Add objCA, strIP
26:        End If
        
        'Check if they are hammering; if so, delete the record (as they are banned
        'for 30 minutes and won't be needed)
30:        If objCA.Check(wskUser) Then
31:            UpdateConnectAttempt = True
32:            m_colConnectAttempts.Remove strIP
33:        End If
34:    End If
    
36:    Exit Function
    
38:
Err:
39:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.UpdateConnectAttempt(, " & blnLoggedIn & ")"
End Function
Private Sub CheckOutdatedRecords()
1:    On Error GoTo Err
    
3:    Dim objCA   As clsConnectAttempt
    
    'Loop through collection
6:    For Each objCA In m_colConnectAttempts
7:        If DateDiff("n", objCA.LastAttempt, Now) > 10 Then _
            m_colConnectAttempts.Remove objCA.IP
9:    Next
        
11:    Exit Sub
    
13:
Err:
14:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.CheckOutdatedRecords()"
End Sub
Public Function LockToKey(ByRef strLock As String, Optional ByVal n As Long = 5) As String
1:    Dim arrChar() As Integer
2:    Dim arrRet() As Integer
3:    Dim i As Long
4:    Dim ub As Long
  
6:    On Error GoTo Err

    'n = 5 for hub and client locks
  
    'The lock only continues to the first space (Pk= comes after)
11:    i = InStrB(1, strLock, " Pk=")
12:    If i Then strLock = LeftB$(strLock, i - 1)

    'Make sure the lock is longer than 2 characters
15:    ub = Len(strLock)
16:    If ub < 3 Then LockToKey = "Invalid lock": Exit Function
    
    'Create buffers to hold vars
19:    ReDim arrChar(1 To ub) As Integer
20:    ReDim arrRet(1 To ub) As Integer
21:    LockToKey = String$(ub * 10, vbNullChar)
    
    'Set first character of string
24:    arrChar(1) = AscW(strLock)
    
    'Set all others and Xor the current and the previous together
27:    For i = 2 To ub
28:        arrChar(i) = AscW(Mid$(strLock, i))
29:        arrRet(i) = arrChar(i) Xor arrChar(i - 1)
30:    Next
    
    'Create first character based on first, last, second last and n from lock
33:    arrRet(1) = arrChar(1) Xor arrChar(ub) Xor arrChar(ub - 1) Xor n
    
    'Delete lock array since it is no longer needed
36:    Erase arrChar
    
    'Set i to 1 so that it starts on the first character
39:    i = 1
    
    'Now loop through and fix all the characters
42:    For n = 1 To ub
43:        arrRet(n) = ((CLng(arrRet(n)) * 16) And 240) Or ((arrRet(n) \ 16) And 15)
        
        'Escape if needed (increment position by 10 if escape is used)
        Select Case arrRet(n)
            Case 0: Mid$(LockToKey, i, 10) = "/%DCN000%/": i = i + 10
            Case 5: Mid$(LockToKey, i, 10) = "/%DCN005%/": i = i + 10
            Case 36: Mid$(LockToKey, i, 10) = "/%DCN036%/": i = i + 10
            Case 96: Mid$(LockToKey, i, 10) = "/%DCN096%/": i = i + 10
            Case 124: Mid$(LockToKey, i, 10) = "/%DCN124%/": i = i + 10
            Case 126: Mid$(LockToKey, i, 10) = "/%DCN126%/": i = i + 10
            Case Else: Mid$(LockToKey, i, 1) = Chr$(arrRet(n)): i = i + 1
46:        End Select
47:    Next
    
    'Erase array containing Xor-ed values
50:    Erase arrRet
    
    'Trim off extra space in the buffer
53:    LockToKey = Left$(LockToKey, i - 1)

55:    Exit Function

57:
Err:
58:    HandleError Err.Number, Err.Description, Erl & "|" & "frmMain.LockToKey(" & strLock & ", " & n & ")"
End Function

'Closes the user's winsock
Public Sub CloseSocket(ByRef intIndex As Integer)
1:    wskLoop_Close intIndex
End Sub

'Calls the VB DoEvents function from VBS
Public Function DoEventsForMe() As Long
1:    DoEventsForMe = DoEvents
End Function

'Registers a bot name in the lists
Public Sub RegisterBotName(ByRef strName As String, Optional ByRef blnOperator As Boolean = True, Optional ByVal dblShare As Double, Optional ByRef strDescription As String, Optional ByRef strConnection As String, Optional ByRef strEmail As String, Optional ByVal lngIcon As Long = 1, Optional ByRef blnOverwrite As Boolean = True)
1:    Dim lngLoop     As Long
2:    Dim blnHold     As Boolean
3:    Dim strOne      As String
4:    Dim strTwo      As String
      
6:    On Error GoTo Err
    
    'Check if it has already been registered
9:    lngLoop = IsRegisteredBotName(strName)
    
    'Make sure we lock the bot name so nobody can log in with it
12:    g_objRegistered.Add strName, "Auto bot name locking system", Locked, "PTDCH / Core"
    
    'If -1, then it isn't, else it is
15:    If lngLoop = -1 Then
        'Fix description if needed
17:        If LenB(strDescription) Then
18:            If InStrB(1, strDescription, "$") Then strDescription = Replace(strDescription, "$", "_")
19:            If InStrB(1, strDescription, "|") Then strDescription = Replace(strDescription, "|", "_")
20:        End If
    
        'Resize array
23:        m_lngBotsUB = m_lngBotsUB + 1
24:        ReDim Preserve m_arrBots(0 To m_lngBotsUB) As typBot
        
        'Update new array element
27:        m_arrBots(m_lngBotsUB).Name = strName
28:        m_arrBots(m_lngBotsUB).MyINFO = "$MyINFO $ALL " & strName & " " & strDescription & "$ $" & strConnection & ChrW$(lngIcon) & "$" & strEmail & "$" & dblShare & "$|"
29:        m_arrBots(m_lngBotsUB).Operator = blnOperator
        
31:        If m_blnServing Then
            'Add to nicklist
33:            g_colUsers.AppendNL strName, blnOperator
        
            'Prepare buffers
36:            If blnOperator Then _
                strOne = m_arrBots(m_lngBotsUB).MyINFO & "$OpList " & strName & "$$|" _
            Else _
                strOne = m_arrBots(m_lngBotsUB).MyINFO
                
41:            strTwo = "$Hello " & strName & "|" & strOne
            
            'Send to all users
44:            For Each m_objLoopUser In g_colUsers
45:                If m_objLoopUser.NoHello Then _
                    m_objLoopUser.SendData strOne _
                Else _
                    m_objLoopUser.SendData strTwo
49:            Next
            
51:            Set m_objLoopUser = Nothing
52:        End If
53:    Else
        'Check if we should overwrite
55:        If blnOverwrite Then
            'Fix description if needed
57:            If LenB(strDescription) Then
58:                If InStrB(1, strDescription, "$") Then strDescription = Replace(strDescription, "$", "_")
59:                If InStrB(1, strDescription, "|") Then strDescription = Replace(strDescription, "|", "_")
60:            End If
        
            'Update array
63:            m_arrBots(lngLoop).MyINFO = "$MyINFO $ALL " & strName & " " & strDescription & "$ $" & strConnection & ChrW$(lngIcon) & "$" & strEmail & "$" & dblShare & "$|"
            
65:            blnHold = m_arrBots(lngLoop).Operator
            
            'Update oplist if necessary or else just myinfo string
68:            If blnHold = blnOperator Then
69:                If m_blnServing Then _
                    g_colUsers.SendToAll m_arrBots(lngLoop).MyINFO
71:            Else
72:                m_arrBots(lngLoop).Operator = blnOperator
                
                'Update only if serving
75:                If m_blnServing Then
76:                    g_colUsers.RemoveNL strName, blnHold
77:                    g_colUsers.AppendNL strName, blnOperator
78:                    g_colUsers.SendToAll "$OpList " & g_colUsers.OpList & "|" & m_arrBots(lngLoop).MyINFO
79:                End If
80:            End If
81:        End If
82:    End If
    
84:    Exit Sub
    
86:
Err:
87:    If ObjPtr(m_objLoopUser) Then Set m_objLoopUser = Nothing
88:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.RegisterBotName()"
End Sub

'Unregisters a bot name in the lists
Public Sub UnregisterBotName(ByRef strName As String)
1:    Dim lngLoop         As Long
    
3:    On Error GoTo Err

    'See if it is registered
6:    lngLoop = IsRegisteredBotName(strName)
    
    'Make sure we unregister the bot name
9:    g_objRegistered.Remove strName
    
    'If it is registered, then get to work!
12:    If Not lngLoop = -1 Then
13:        m_lngBotsUB = m_lngBotsUB - 1
        
        'Update nicklist (local and remote) if serving
16:        If m_blnServing Then
17:            g_colUsers.RemoveNL strName, m_arrBots(lngLoop).Operator
18:            g_colUsers.SendToAll "$Quit " & strName & "|"
19:        End If
        
        'If there are no bots left in the array, destroy it, otherwise resize it
22:        If m_lngBotsUB = -1 Then
23:            Erase m_arrBots
24:        Else
25:            m_arrBots(lngLoop) = m_arrBots(m_lngBotsUB + 1)
26:            ReDim Preserve m_arrBots(0 To m_lngBotsUB) As typBot
27:        End If
28:    End If
    
30:    Exit Sub
    
32:
Err:
33:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.UnregisterBotName(""" & strName & """)"
End Sub

'Returns -1 is not registered, index in array otherwise
Public Function IsRegisteredBotName(ByRef strName As String) As Long
1:    Dim lngLoop As Long
    
3:    On Error GoTo Err
    
    'Make any needed character replacements in the nickname
6:    If InStrB(1, strName, " ") Then strName = Replace(strName, " ", "_")
7:    If InStrB(1, strName, "$") Then strName = Replace(strName, "$", "_")
8:    If InStrB(1, strName, "|") Then strName = Replace(strName, "|", "_")
    
    'Set to -1, meaning it hasn't found the bot name
11:    IsRegisteredBotName = -1
    
    'Make sure they are bots in the array first
14:    If Not m_lngBotsUB = -1 Then
        'Loop through and see if the name matches any; if it does, return array index
16:        For lngLoop = 0 To m_lngBotsUB
17:            If m_arrBots(lngLoop).Name = strName Then IsRegisteredBotName = lngLoop: Exit For
18:        Next
19:    End If
    
21:    Exit Function
    
23:
Err:
24:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.IsRegisteredBotName(""" & strName & """)"
End Function

'Switches to the next redirect address
Public Sub NextRedirect()
1:    Static lngindex As Long
    
3:    On Error GoTo Err
    
    'If the UBound is zero, then there is only one IP
6:    If m_lngRedirectUB Then
        'Increment index
8:        lngindex = lngindex + 1
    
        'If index has surpassed the max index, then set it back to the beginning
11:        If lngindex > m_lngRedirectUB Then lngindex = 0
        
        'Set redirect IP in settings to the new IP
14:        g_objSettings.RedirectIP = m_arrRedirectIPs(lngindex)
15:    End If
    
17:    Exit Sub
    
19:
Err:
20:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.NextRedirect()"
End Sub

'Sends bot MyINFOs to a user
Friend Sub UpdateBots(ByRef curUser As clsUser)
1:    Dim lngLoop As Long
    
3:    On Error GoTo Err
    
    'Make sure there are bot names to send
6:    If Not m_lngBotsUB = -1 Then
        'Loop through and send MyINFO strings
8:        For lngLoop = 0 To m_lngBotsUB
9:            curUser.SendData m_arrBots(lngLoop).MyINFO
10:        Next
11:    End If
    
13:    Exit Sub

15:
Err:
16:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.UpdateBots()"
End Sub

'Converts minutes to it's equivalent length in years, months, weeks, days, hours and minutes
Public Function MinToDate(ByVal lngMinutes As Long) As String
1:    Dim lngYears As Long
2:    Dim lngMonths As Long
3:    Dim lngWeeks As Long
4:    Dim lngDays As Long
5:    Dim lngHours As Long
    
7:    On Error GoTo Err
    
    'If there are more than 59 minutes, there is at least 1 hour
10:    If lngMinutes > 59 Then
11:        lngHours = lngMinutes \ 60
12:        lngMinutes = lngMinutes Mod 60
        'If there are more than 23 hours, there is at least 1 day
14:        If lngHours > 23 Then
15:            lngDays = lngHours \ 24
16:            lngHours = lngHours Mod 24
            
            'If there are more than 29 days, there is at least 1 month
19:            If lngDays > 29 Then
20:                lngMonths = lngDays \ 30
21:                lngDays = lngDays Mod 30
            'If there are more than 7 days, there is at least 1 week
23:            Else
24:                If lngDays > 6 Then
25:                    lngWeeks = lngDays \ 7
26:                    lngDays = lngDays Mod 7
27:                End If
28:            End If
            'If there are more than 11 months, there is at least 1 year
30:            If lngMonths > 11 Then
31:                lngYears = lngMonths \ 12
32:                lngMonths = lngMonths Mod 12
33:            End If
34:        End If
35:    End If

    'Construct length in words
38:    If lngYears > 1 Then MinToDate = lngYears & " years, " _
    Else: If lngYears Then MinToDate = lngYears & " year, "
40:    If lngMonths > 1 Then MinToDate = MinToDate & lngMonths & " months, " _
    Else: If lngMonths Then MinToDate = MinToDate & lngMonths & " month, "
42:    If lngWeeks > 1 Then MinToDate = MinToDate & lngWeeks & " weeks, " _
    Else: If lngWeeks Then MinToDate = MinToDate & lngWeeks & " week, "
44:    If lngDays > 1 Then MinToDate = MinToDate & lngDays & " days, " _
    Else: If lngDays Then MinToDate = MinToDate & lngDays & " day, "
46:    If lngHours > 1 Then MinToDate = MinToDate & lngHours & " hours, " _
    Else: If lngHours Then MinToDate = MinToDate & lngHours & " hour, "
    
    'If there are no minutes to add, then we need to remove the extra space/comma
50:    If lngMinutes > 1 Then
51:        MinToDate = MinToDate & lngMinutes & " minutes"
52:        Exit Function
53:    Else
54:        If lngMinutes Then
55:            MinToDate = MinToDate & lngMinutes & " minute"
56:            Exit Function
57:        End If
58:    End If
59:    If LenB(MinToDate) Then
60:            MinToDate = LeftB$(MinToDate, LenB(MinToDate) - 4)
61:        Else
62:            MinToDate = "0 minutes"
63:    End If
    
65:    Exit Function
    
67:
Err:
68:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.MinToDate()"
End Function

Public Property Get oPermaCon() As Connection
1:    Set oPermaCon = m_objPermaCon
End Property

Public Property Get ServingDate() As Date
1:    ServingDate = m_datServingDate
End Property

Public Property Get IsServing() As Boolean
1:    IsServing = m_blnServing
End Property

'------------------------------------------------------------------------------
'Script related functions / events
'------------------------------------------------------------------------------
Private Sub tmrScriptTimer_Timer(Index As Integer)
1:    On Error GoTo Err
    
    'This runs the tmrScriptTimer_Timer event
    '
    '  -- Parameters : None
    '  -- Format     : Sub tmrScriptTimer_Timer()
    '
    '  -- Called whenever the alloted interval for the timer has gone
    '     off
    
11:    If m_arrScriptEvents(Index, vbStmrScriptTimer_Timer) Then _
        ScriptControl(Index).Run "tmrScriptTimer_Timer" _
    Else _
        tmrScriptTimer(Index).Enabled = False

16:
Err:
End Sub
Private Sub ScriptControl_Error(Index As Integer)
1:    On Error GoTo Err
    
    'This runs the Error event
    '
    '  -- Parameters : lngLine (line the error occured on)
    '  -- Format     : Sub Error(lngLine)
    '
    '  -- Called when an error occurs and On Error Resume Next is not used
    
10:    If m_arrScriptEvents(Index, vbSError) Then _
        ScriptControl(Index).Run "Error", ScriptControl(Index).Error.Line

13:
Err:
End Sub

'Private Sub ScriptControl_Timeout(Index As Integer)
'    On Error Resume Next
'
'    'This runs the Timeout event
'    '
'    '  -- Parameters : None
'    '  -- Format     : Sub Timeout()
'    '
'    '  -- Run when the script code timeouts
'
'    If m_arrScriptEvents(Index, vbSTimeout) Then _
'        ScriptControl(Index).Run "Timeout"
'End Sub
Private Sub SEvent_AttemptedConnection(ByRef strIP As String)
1:    Dim lng As Long

    'This runs the AttemptedConnection event
    '
    '  -- Parameters : strIP (IP of user who is trying to connect)
    '  -- Format     : Sub AttemptedConnection(strIP)
    '
    '  -- Called when a new user tries to connect to the hub (before any
    '     messages are exchanged)
    
11:    On Error Resume Next
    
13:    For lng = 1 To m_lngScriptEventsUB
14:        If m_arrScriptEvents(lng, vbSAttemptedConnection) Then _
                ScriptControl(lng).Run "AttemptedConnection", strIP
16:    Next
End Sub
Private Sub SEvent_UserConnected(ByRef curUser As clsUser)
1:    Dim lng As Long

    'This runs the UserConnected event
    '
    '  -- Parameters : curUser (user's clsUser object)
    '  -- Format     : Sub UserConnected(curUser)
    '
    '  -- Called when a new user logs in (sends MyINFO basically)
    '  -- There is a slight difference with this and NMDCH; DDCH calls this
    '     event after it recieves MyINFO and the like, while NMDCH calls this
    '     right after the ValidateNick message
    
13:    On Error Resume Next
    
15:    For lng = 1 To m_lngScriptEventsUB
16:        If m_arrScriptEvents(lng, vbSUserConnected) Then _
                ScriptControl(lng).Run "UserConnected", curUser
18:    Next
End Sub
Private Sub SEvent_RegConnected(ByRef curUser As clsUser)
1:    Dim lng As Long

    'This runs the RegConnected event
    '
    '  -- Parameters : curUser (user's clsUser object)
    '  -- Format     : Sub RegConnected(curUser)
    '
    '  -- Called when a new user logs in (sends MyINFO basically)
    '  -- There is a slight difference with this and NMDCH; NMDCH combines this
    '     event with OpConnected, which is wrong...they are not an op so people
    '     can get confused
    
13:    On Error Resume Next
    
15:    For lng = 1 To m_lngScriptEventsUB
16:        If m_arrScriptEvents(lng, vbSRegConnected) Then _
            ScriptControl(lng).Run "RegConnected", curUser
18:    Next
End Sub
Private Sub SEvent_OpConnected(ByRef curUser As clsUser)
1:    Dim lng As Long

    'This runs the OpConnected event
    '
    '  -- Parameters : curUser (user's clsUser object)
    '  -- Format     : Sub OpConnected(curUser)
    '
    '  -- Called when a new user logs in (sends MyINFO basically)
    '  -- There are two slight differences with this and NMDCH; DDCH calls this
    '     event after it recieves MyINFO and the like, while NMDCH calls this
    '     right after the ValidateNick message. Second DDCH only calls this event
    '     with OPERATORS and calls non-ops-but-registered users with RegConnected
    
14:    On Error Resume Next
    
16:    For lng = 1 To m_lngScriptEventsUB
17:        If m_arrScriptEvents(lng, vbSOpConnected) Then _
               ScriptControl(lng).Run "OpConnected", curUser
19:    Next
End Sub
Private Sub SEvent_UserQuit(ByRef curUser As clsUser)
1:    Dim lng As Long
    
    'This runs the UserQuit event
    '
    '  -- Parameters : curUser (user's clsUser object)
    '  -- Format     : Sub UserQuit(curUser)
    '
    '  -- Called when a user leaves the hub. After this sub is called
    '     the user's clsHub object is destroyed (however you must remove it
    '     from any collections that might contain it in the hub)
    
12:    On Error Resume Next
    
14:    For lng = 1 To m_lngScriptEventsUB
15:        If m_arrScriptEvents(lng, vbSUserQuit) Then _
                ScriptControl(lng).Run "UserQuit", curUser
17:    Next
End Sub
Private Sub SEvent_StartedServing()
1:    Dim lng As Long

    'This runs the StartedServing event
    '
    '  -- Parameters : None
    '  -- Format     : Sub StartServing()
    '
    '  -- Run when the hub starts serving
    
10:    On Error Resume Next
    
12:    For lng = 1 To m_lngScriptEventsUB
13:        If m_arrScriptEvents(lng, vbSStartedServing) Then _
                ScriptControl(lng).Run "StartedServing"
15:    Next
End Sub
Friend Sub SEvent_AddedRegisteredUser(ByRef strName As String)
1:    Dim lng As Long

    'This runs the AddedRegisteredUser event
    '
    '  -- Parameters : strName (name of the user who was registered)
    '  -- Format     : Sub AddedRegisteredUser(strName)
    '
    '  -- Run when a new user is registered (via the clsRegistered.Add function)
    
10:    On Error Resume Next
    
12:    For lng = 1 To m_lngScriptEventsUB
13:        If m_arrScriptEvents(lng, vbSAddedRegisteredUser) Then _
               ScriptControl(lng).Run "AddedRegisteredUser", strName
15:    Next
End Sub
Friend Sub SEvent_AddedPermBan(ByRef strIP As String)
1:    Dim lng As Long
    
    'This runs the AddedPermBan event
    '
    '  -- Parameters : strIP (IP that was banned)
    '  -- Format     : Sub AddedPermBan(strIP)
    '
    '  -- Called when a user perm bans an IP via the clsIPBans.Add method
    
10:    On Error Resume Next
    
12:    For lng = 1 To m_lngScriptEventsUB
13:        If m_arrScriptEvents(lng, vbSAddedPermBan) Then _
                ScriptControl(lng).Run "AddedPermBan", strIP
14:    Next
End Sub
Private Sub SEvent_StartedRedirecting()
1:    Dim lng As Long
    
    'This runs the StartedRedirecting event
    '
    '  -- Parameters : None
    '  -- Format     : Sub StartedRedirecting()
    '
    '  -- Called when the hub owner presses the Redirect All button (just before
    '     redirects are done)
    
11:    On Error Resume Next
    
13:    For lng = 1 To m_lngScriptEventsUB
14:        If m_arrScriptEvents(lng, vbSStartedRedirecting) Then _
                ScriptControl(lng).Run "StartedRedirecting"
16:    Next
End Sub
Private Sub SEvent_StoppedServing()
1:    Dim lng As Long

    'This runs the StoppedServing event
    '
    '  -- Parameters : None
    '  -- Format     : Sub StoppedServing
    '
    '  -- Called when the hub owner stops serving
    
10:    On Error Resume Next
    
12:    For lng = 1 To m_lngScriptEventsUB
13:        If m_arrScriptEvents(lng, vbSStoppedServing) Then _
                ScriptControl(lng).Run "StoppedServing"
15:    Next
End Sub
Private Sub SEvent_MassMessage(ByRef strMessage As String)
1:    Dim lng As Long

    'This runs the MassMessage event
    '
    '  -- Parameters : strMessage (message that was sent to all users)
    '  -- Format     : Sub MassMessage(strMessage)
    '
    '  -- Run when the hub owner presses the mass message button
    
10:    On Error Resume Next
    
12:    For lng = 1 To m_lngScriptEventsUB
13:        If m_arrScriptEvents(lng, vbSMassMessage) Then _
               ScriptControl(lng).Run "MassMessage", strMessage
15:    Next
End Sub
Private Sub SEvent_UnloadMain()
1:    Dim lng As Long

    'This runs the UnloadMain event
    '
    '  -- Parameters : None
    '  -- Format     : Sub UnloadMain()
    '
    '  -- Called when the hub is closing up
    
10:    On Error Resume Next
    
12:    For lng = 1 To m_lngScriptEventsUB
13:        If m_arrScriptEvents(lng, vbSUnloadMain) Then _
                ScriptControl(lng).Run "UnloadMain"
15:    Next
End Sub
Friend Sub SEvent_RemovedRegisteredUser(ByRef strName As String)
1:    Dim lng As Long

    'This runs the RemovedRegisteredUser event
    '
    '  -- Parameters : strName (name of the user who was unregistered)
    '  -- Format     : Sub RemovedRegisteredUser(strName)
    '
    '  -- Run when a user is unregistered (via the clsRegistered.Remove function)
    
10:    On Error Resume Next
    
12:    For lng = 1 To m_lngScriptEventsUB
13:        If m_arrScriptEvents(lng, vbSRemovedRegisteredUser) Then _
                ScriptControl(lng).Run "RemovedRegisteredUser", strName
15:    Next
End Sub
Private Sub SEvent_CustComArrival(ByRef curUser As clsUser, ByRef objCommand As clsCommand, ByRef strMessage As String, ByRef blnMainChat As Boolean)
1:    Dim lng As Long
        
    'This runs the CustComArrival event
    '
    '   -- Parameters : curUser (current user's object)
    '                 : objCommand (current command object from collection)
    '                 : strMessage (command text sent by user)
    '   -- Format     : Sub CustComArrival(curUser, objCommand, strMessage)
    '
    '   -- Fired when a user sends a command which is in the command collection
    '      but not supported by the hub
    
13:    On Error Resume Next
    
15:    For lng = 1 To m_lngScriptEventsUB
16:        If m_arrScriptEvents(lng, vbSCustComArrival) Then _
                ScriptControl(lng).Run "CustComArrival", curUser, objCommand, strMessage, blnMainChat
18:    Next
End Sub
Private Function SEvent_FailedConf(ByRef curUser As clsUser, ByRef intType As enuAlert) As Boolean

2:    Dim lng As Long
        
    'This runs the FailedConf event
    '
    '   -- Parameters : curUser (current user's object)
    '                 : intType As enuAlert enumerator
    '
    '   -- Return     : boolean
    '
    '   -- Fired when a user get rejected by the hub.(user fail hub's rules.)
    
13:    On Error Resume Next
    
15:    For lng = 1 To m_lngScriptEventsUB
16:        If m_arrScriptEvents(lng, vbSFailedConf) Then _
            SEvent_FailedConf = ScriptControl(lng).Run("FailedConf", curUser, intType)
18:            If SEvent_FailedConf Then Exit For
19:    Next
End Function
Friend Sub SFindEvents(ByRef intIndex As Integer)
1:    Dim lngProc     As Long
2:    Dim lngProcUB   As Long
3:    Dim objProc     As Procedure
4:    Dim objProcs    As Procedures
    
6:    On Error GoTo Err
    
    'Prepare vars
9:    Set objProcs = ScriptControl(intIndex).Procedures
10:    lngProcUB = objProcs.count
    
    'Clear out array
13:    For lngProc = 0 To vbSFC
14:        m_arrScriptEvents(intIndex, lngProc) = False
15:    Next
    
    '#If PREDATAARRIVAL Then
    '    'Make sure that PreDataArrival is disabled if it was this script that was using it
    '    If m_intPDIndex = intIndex Then m_intPDIndex = 0
    '#End If

    'Loop through procedures
23:    For Each objProc In objProcs
        'Find out which procedure it is, and set it to True in the boolean array
        Select Case LCase$(objProc.Name)
            Case "main"
25:                If objProc.NumArgs = 0 Then _
                    m_arrScriptEvents(intIndex, vbSMain) = True
         
         #If DataArrival Then
            Case "dataarrival"
29:                If objProc.NumArgs = 2 Then _
                    m_arrScriptEvents(intIndex, vbSDataArrival) = True
         #End If
            
            Case "attemptedconnection"
33:                If objProc.NumArgs = 1 Then _
                    m_arrScriptEvents(intIndex, vbSAttemptedConnection) = True
            Case "userconnected"
35:                If objProc.NumArgs = 1 Then _
                    m_arrScriptEvents(intIndex, vbSUserConnected) = True
            Case "regconnected"
37:                If objProc.NumArgs = 1 Then _
                    m_arrScriptEvents(intIndex, vbSRegConnected) = True
            Case "opconnected"
39:                If objProc.NumArgs = 1 Then _
                    m_arrScriptEvents(intIndex, vbSOpConnected) = True
            Case "userquit":
41:                If objProc.NumArgs = 1 Then _
                    m_arrScriptEvents(intIndex, vbSUserQuit) = True
            Case "startedserving"
43:                If objProc.NumArgs = 0 Then _
                    m_arrScriptEvents(intIndex, vbSStartedServing) = True
            Case "systraydoubleclick"
45:                If objProc.NumArgs = 0 Then _
                    m_arrScriptEvents(intIndex, vbSSysTrayDoubleClick) = True
            Case "addedregistereduser"
47:                If objProc.NumArgs = 1 Then _
                    m_arrScriptEvents(intIndex, vbSAddedRegisteredUser) = True
            Case "wskscript_close"
49:                If objProc.NumArgs = 1 Then _
                    m_arrScriptEvents(intIndex, vbSwskScript_Close) = True
            Case "wskscript_connect"
51:                If objProc.NumArgs = 1 Then _
                    m_arrScriptEvents(intIndex, vbSwskScript_Connect) = True
            Case "wskscript_connectionrequest"
53:                If objProc.NumArgs = 2 Then _
                    m_arrScriptEvents(intIndex, vbSwskScript_ConnectionRequest) = True
            Case "wskscript_dataarrival"
55:                If objProc.NumArgs = 2 Then _
                    m_arrScriptEvents(intIndex, vbSwskScript_DataArrival) = True
            Case "wskscript_error"
57:                If objProc.NumArgs = 3 Then _
                    m_arrScriptEvents(intIndex, vbSwskScript_Error) = True
            Case "tmrscripttimer_timer"
59:                If objProc.NumArgs = 0 Then _
                    m_arrScriptEvents(intIndex, vbStmrScriptTimer_Timer) = True
            Case "addedpermban"
61:                If objProc.NumArgs = 1 Then _
                    m_arrScriptEvents(intIndex, vbSAddedPermBan) = True
            Case "startedredirecting"
63:                If objProc.NumArgs = 0 Then _
                    m_arrScriptEvents(intIndex, vbSStartedServing) = True
            Case "stoppedserving"
65:                If objProc.NumArgs = 0 Then _
                    m_arrScriptEvents(intIndex, vbSStoppedServing) = True
            Case "mouseoversystray"
67:                If objProc.NumArgs = 0 Then _
                    m_arrScriptEvents(intIndex, vbSMouseOverSysTray) = True
            Case "massmessage"
69:                If objProc.NumArgs = 1 Then _
                    m_arrScriptEvents(intIndex, vbSMassMessage) = True
            Case "unloadmain"
71:                If objProc.NumArgs = 0 Then _
                    m_arrScriptEvents(intIndex, vbSUnloadMain) = True
            Case "error"
73:                If objProc.NumArgs = 1 Then _
                    m_arrScriptEvents(intIndex, vbSError) = True
            Case "timeout"
75:                If objProc.NumArgs = 0 Then _
                    m_arrScriptEvents(intIndex, vbSTimeout) = True
            Case "removedregistereduser"
77:                If objProc.NumArgs = 1 Then _
                    m_arrScriptEvents(intIndex, vbSRemovedRegisteredUser) = True
            Case "custcomarrival"
79:                If objProc.NumArgs = 4 Then _
                    m_arrScriptEvents(intIndex, vbSCustComArrival) = True
            Case "failedconf"
81:                If objProc.NumArgs = 2 Then _
                    If objProc.HasReturnValue Then _
                        m_arrScriptEvents(intIndex, vbSFailedConf) = True
            
            'Evaluate PreDataArrival if needed
            #If PreDataArrival Then
                Case "predataarrival"
87:                    If objProc.NumArgs = 2 Then _
                        If objProc.HasReturnValue Then _
                            m_arrScriptEvents(intIndex, vbSPreDataArrival) = True
            #End If
91:        End Select
92:    Next

    'Set winsock collection booleans
95:    g_colSWinsocks(CStr(intIndex)).SetBools m_arrScriptEvents(intIndex, vbSwskScript_Connect), _
                                               m_arrScriptEvents(intIndex, vbSwskScript_Close), _
                                               m_arrScriptEvents(intIndex, vbSwskScript_ConnectionRequest), _
                                               m_arrScriptEvents(intIndex, vbSwskScript_DataArrival), _
                                               m_arrScriptEvents(intIndex, vbSwskScript_Error)
    
101:    Exit Sub
    
103:
Err:
104:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.SFindEvents(" & intIndex & ")"
End Sub
Friend Sub SClearEvents(ByRef intIndex As Integer)
1:    Dim lng As Long
    
3:    On Error GoTo Err
    
    'Loop through array and set all procedure enabled settings to false
6:    For lng = 0 To vbSFC
7:        m_arrScriptEvents(intIndex, lng) = False
8:    Next
    
    'Set all winsock vars to false
11:    g_colSWinsocks(CStr(intIndex)).SetBools False, False, False, False, False
    
    'Set PreDataArrival index to zero if it in use by this script
    '#If PREDATAARRIVAL Then
    '    If m_intPDIndex = intIndex Then m_intPDIndex = 0
    '#End If
    
18:    Exit Sub
    
20:
Err:
21:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.SClearEvents(" & intIndex & ")"
End Sub
Friend Sub SResizeArrEvent(ByRef intSize As Integer, ByRef blnPreserve As Boolean)
1:    On Error GoTo Err
    
    'Preserve elements if needed
4:    If blnPreserve Then
5:        ReDim Preserve m_arrScriptEvents(1 To intSize, 0 To vbSFC) As Boolean
6:    Else
7:        ReDim m_arrScriptEvents(1 To intSize, 0 To vbSFC) As Boolean
8:    End If
        
       'Set new UBound
11:    m_lngScriptEventsUB = intSize
    
13:    Exit Sub
    
15:
Err:
16:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.SResizeArrEvent(" & intSize & ", " & blnPreserve & ")"
End Sub
'------------------------------------------------------------------------------
'End Script related functions / events
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'End Update DSN/IP Subs
'------------------------------------------------------------------------------
Private Sub UpdateIPs_Timer()
1:    On Error GoTo Err
    
    'do check every IPCheckInterval
4:    If lIntervalMin < IPCheckInterval Then
5:        lIntervalMin = lIntervalMin + 1
6:        Exit Sub
7:    End If
8:    lIntervalMin = 0
    
10:    Dim successor
11:    Dim HubIP
12:    Dim HostIP
13:    Dim X

15:    HubIP = frmHub.DetectHubIP
16:    If Not HubIP <> "" Then Exit Sub
17:    DoEvents
    
    'check against local IPs
20:    If IPinRange("10.0.0.0", "10.255.255.255", HubIP) Then Exit Sub
21:    If IPinRange("127.0.0.0", "127.255.255.255", HubIP) Then Exit Sub
22:    If IPinRange("172.16.0.0", "172.31.255.255", HubIP) Then Exit Sub
23:    If IPinRange("192.168.0.0", "192.168.255.255", HubIP) Then Exit Sub

25:    For X = 0 To 9 'a maximum of 10 services should be enough
26:        If Service(X) = "" Then Exit Sub
        
28:        DoEvents
29:        HostIP = frmHub.ResolveHostName(CStr(Host(X)))
  
31:        If HubIP <> HostIP Then
32:            DoEvents
33:            successor = frmHub.UpdateIP(CStr(Service(X)), CStr(User(X)), CStr(Pass(X)), CStr(Host(X)))
34:            g_colUsers.SendPrivateToOps g_objSettings.BotName, "IP UPDATE " & CStr(Host(X)) & ": " & CStr(successor)
35:        End If
'            g_colUsers.SendPrivateToOps g_objSettings.BotName, "IP UPDATE " & CStr(Host(X)) & ": " & CStr(successor)
37:    Next

39:    Exit Sub
40:
Err:
41:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.UpdateIPs_Timer"
End Sub
Private Function QueuedConnect(ByRef strName As String) As Boolean
1:    Dim datTemp     As Date
    
3:    On Error GoTo DNE
    

6:    datTemp = m_colRevConnects(strName)
7:    m_colRevConnects.Remove strName
8:    QueuedConnect = (DateDiff("s", datTemp, Now) < 60)
    
10:    Exit Function
    
12:
DNE:
13:    QueuedConnect = False
End Function
Public Function UpdateIP(Service, UserName, Password, HostName)
1:    On Error Resume Next
2:    If LCase(Service) = "dyndns" Then _
        UpdateIP = UpdateDynDNS(UserName, Password, HostName, True, "", False)
4:    If LCase(Service) = "noip" Then _
        UpdateIP = UpdateNoIP(UserName, Password, HostName, "")
End Function
Private Sub UpdateDNSs()
1:    Dim strTemp As String
2:    Dim successor
3:    Dim HubIP
4:    Dim HostIP
    
6:    On Error GoTo Err
    
8:    HubIP = DetectHubIP
9:    DoEvents
    
11:    If HubIP = vbNullString Then Exit Sub
    
13:    If g_objSettings.NoIPUpdateEna Then
14:        If g_objSettings.NoIPDNS1 <> vbNullString Then
15:            HostIP = ResolveHostName(g_objSettings.NoIPDNS1)
16:            DoEvents
               'Add Satatus Log text and Set stbMain.Panels(8).Text
18:            If Not HubIP = HostIP Then
19:                successor = UpdNOIP(g_objSettings.NoIPUser, g_objSettings.NoIPPass, g_objSettings.NoIPDNS1)
20:                DoEvents
21:                strTemp = "IP UpDate: " & g_objSettings.NoIPDNS1 & " - " & successor: AddLog strTemp, 2: stbMain.Panels(9).Text = strTemp
22:            End If
23:        End If
        
25:        If g_objSettings.NoIPDNS2 <> vbNullString Then
26:            HostIP = ResolveHostName(g_objSettings.NoIPDNS2)
27:            DoEvents
               'Add Satatus Log text and Set stbMain.Panels(8).Text
29:            If Not HubIP = HostIP Then
30:                successor = UpdNOIP(g_objSettings.NoIPUser, g_objSettings.NoIPPass, g_objSettings.NoIPDNS2)
31:                DoEvents
32:                strTemp = "IP UpDate : " & g_objSettings.NoIPDNS2 & " - " & successor: AddLog strTemp, 2: stbMain.Panels(9).Text = strTemp
33:            End If
34:        End If
        
36:        If g_objSettings.NoIPDNS3 <> vbNullString Then
37:            HostIP = ResolveHostName(g_objSettings.NoIPDNS3)
38:            DoEvents
               'Add Satatus Log text and Set stbMain.Panels(8).Text
40:            If Not HubIP = HostIP Then
41:                successor = UpdNOIP(g_objSettings.NoIPUser, g_objSettings.NoIPPass, g_objSettings.NoIPDNS3)
42:                DoEvents
43:                strTemp = "IP UpDate: " & g_objSettings.NoIPDNS3 & " - " & successor: AddLog strTemp, 2: stbMain.Panels(9).Text = strTemp
44:            End If
45:        End If
        
47:        If g_objSettings.NoIPDNS4 <> vbNullString Then
48:            HostIP = ResolveHostName(g_objSettings.NoIPDNS4)
49:            DoEvents
               'Add Satatus Log text and Set stbMain.Panels(8).Text
51:            If Not HubIP = HostIP Then
52:                successor = UpdNOIP(g_objSettings.NoIPUser, g_objSettings.NoIPPass, g_objSettings.NoIPDNS4)
53:                DoEvents
54:                strTemp = "IP UpDate : " & g_objSettings.NoIPDNS4 & " - " & successor: AddLog strTemp, 2: stbMain.Panels(9).Text = strTemp
55:            End If
56:        End If
57:    End If

59:    If g_objSettings.DynDNSUpdateEna Then
60:        If g_objSettings.DynDNS1 <> vbNullString Then
61:            HostIP = ResolveHostName(g_objSettings.DynDNS1)
62:            DoEvents
               'Add Satatus Log text and Set stbMain.Panels(8).Text
64:            If Not HubIP = HostIP Then
65:                successor = UpdDynDNS(g_objSettings.DynDNSUser, g_objSettings.DynDNSPass, g_objSettings.DynDNS1)
66:                DoEvents
67:                strTemp = "IP UpDate: " & g_objSettings.DynDNS1 & " - " & successor: AddLog strTemp, 2: stbMain.Panels(9).Text = strTemp
68:            End If
69:        End If
        
71:        If g_objSettings.DynDNS2 <> vbNullString Then
72:            HostIP = ResolveHostName(g_objSettings.DynDNS2)
73:            DoEvents
               'Add Satatus Log text and Set stbMain.Panels(8).Text
75:            If Not HubIP = HostIP <> vbNullString Then
76:                successor = UpdDynDNS(g_objSettings.DynDNSUser, g_objSettings.DynDNSPass, g_objSettings.DynDNS2)
77:                DoEvents
78:                strTemp = "IP UpDate: " & g_objSettings.DynDNS2 & " - " & successor: AddLog strTemp, 2: stbMain.Panels(9).Text = strTemp
79:            End If
80:        End If
        
82:        If g_objSettings.DynDNS3 <> vbNullString Then
83:            HostIP = ResolveHostName(g_objSettings.DynDNS3)
84:            DoEvents
               'Add Satatus Log text and Set stbMain.Panels(8).Text
86:            If Not HubIP = HostIP Then
87:                successor = UpdDynDNS(g_objSettings.DynDNSUser, g_objSettings.DynDNSPass, g_objSettings.DynDNS3)
88:                DoEvents
89:                strTemp = "IP UpDate: " & g_objSettings.DynDNS3 & " - " & successor: AddLog strTemp, 2: stbMain.Panels(9).Text = strTemp
90:            End If
91:        End If
        
93:        If g_objSettings.DynDNS4 <> vbNullString Then
94:            HostIP = ResolveHostName(g_objSettings.DynDNS4)
95:            DoEvents
               'Add Satatus Log text and Set stbMain.Panels(8).Text
97:            If Not HubIP = HostIP Then
98:                successor = UpdDynDNS(g_objSettings.DynDNSUser, g_objSettings.DynDNSPass, g_objSettings.DynDNS4)
99:                DoEvents
100:               strTemp = "IP UpDate: " & g_objSettings.DynDNS4 & " - " & successor: AddLog strTemp, 2: stbMain.Panels(9).Text = strTemp
101:            End If
102:        End If
103:    End If

105:    Exit Sub
106:
Err:
107:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.UpdateDNSs"
End Sub
Public Function UpdNOIP(UserName, Password, HostName)
1:    On Error Resume Next
2:    UpdNOIP = UpdateNoIP(UserName, Password, HostName, "")
End Function
Public Function UpdDynDNS(UserName, Password, HostName)
1:    On Error Resume Next
2:    UpdDynDNS = UpdateDynDNS(UserName, Password, HostName, True, "", False)
End Function
Public Function ResolveHostName(HostName)
1:    On Error Resume Next
2:    ResolveHostName = ResolveHost(CStr(HostName))
End Function
Public Function DetectHubIP()
1:    On Error Resume Next
2:    Dim X: X = DetectIP()
3:    DetectHubIP = X
4:    If X = vbNullString Then X = "Connection refused by target machine"
5:    AddLog "Detect Hub IP: " & X, 2
6:    stbMain.Panels(8).Text = X
End Function
'------------------------------------------------------------------------------
'End Update DSN/IP Subs
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'MS Access Subs
'------------------------------------------------------------------------------
Private Sub DBGetBanRecord()
      Dim lvwItem   As ListItem
      Dim lvwItems  As ListItems
      Dim i         As Long
4:    On Error GoTo Err

     '<frmHub.adoBans SQL Command> -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
     '
     'SELECT UsrClass.UserName, BanNames.Perm, BanNames.BannedBy, BanNames.RefDate, BanNames.Reason
     'FROM UsrClass INNER JOIN BanNames ON UsrClass.UserName = BanNames.UserName
     'Where (((UsrClass.Class) = -1))
     'ORDER BY BanNames.RefDate;
     '
     '<frmHub.adoBans SQL Command/> ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

15:    lvwBans.ListItems.Clear

17:    With adoBans
18:        .Refresh
19:        Do While Not .Recordset.EOF
20:            Set lvwItems = frmHub.lvwBans.ListItems
               'Add listitem
22:            Set lvwItem = lvwItems.Add(, , CStr(.Recordset(0).Value)) 'UsrClass.UserName
23:            lvwItem.SubItems(1) = CBool(.Recordset(1).Value) 'BanNames.Perm
24:            lvwItem.SubItems(2) = CStr(.Recordset(2).Value) 'BanNames.BannedBy
25:            lvwItem.SubItems(3) = CDate(.Recordset(3).Value) 'BanNames.RefDate
26:            On Error Resume Next
27:            lvwItem.Tag = CStr(.Recordset(4).Value) 'BanNames.Reason
28:            On Error GoTo Err
29:            .Recordset.MoveNext
30:        Loop
31:    End With
 
33:    lblHolder(50).Caption = ""

35:   Exit Sub
36:
Err:
38:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.DBGetBanRecord()"
End Sub
Private Sub DBGetRegRecord()
    Dim lvwItem   As ListItem
    Dim lvwItems  As ListItems
    Dim i         As Long
4:  On Error GoTo Err

    '<frmHub.adoUsers SQL Command> -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '
    'SELECT UsrClass.UserName, UsrStatic.Pass, UsrClass.Class, ClassTypes.Name, UsrStatic.RegedBy, UsrStatic.RegDate, UsrDynamic.LastLogin, UsrDynamic.LastIP
    'FROM UsrStatic INNER JOIN (ClassTypes INNER JOIN (UsrClass INNER JOIN UsrDynamic ON UsrClass.UserName = UsrDynamic.UserName) ON ClassTypes.ID = UsrClass.Class) ON (UsrStatic.UserName = UsrDynamic.UserName) AND (UsrStatic.UserName = UsrClass.UserName)
    'ORDER BY UsrClass.UserName;
    '
    '<frmHub.adoUsers SQL Command/> ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

14:    lvwRegistered.ListItems.Clear

16:    With adoUsers
17:        .Refresh
18:        Do While Not .Recordset.EOF
19:            Set lvwItems = frmHub.lvwRegistered.ListItems
               'Add listitem
21:            Set lvwItem = lvwItems.Add(, , CStr(.Recordset(0).Value)) 'UsrClass.UserName
22:            lvwItem.SubItems(1) = CStr(.Recordset(1).Value) 'UsrStatic.Pass
23:            lvwItem.SubItems(2) = CInt(.Recordset(2).Value) 'UsrClass.Class
24:            lvwItem.SubItems(3) = CStr(.Recordset(3).Value) 'ClassTypes.Name
25:            lvwItem.SubItems(4) = CStr(.Recordset(4).Value) 'UsrStatic.RegedBy
26:            lvwItem.SubItems(5) = CDate(.Recordset(5).Value) 'UsrStatic.RegDate
27:            On Error Resume Next
28:            lvwItem.SubItems(6) = CDate(.Recordset(6).Value) 'UsrDynamic.LastLogin
29:            lvwItem.SubItems(7) = CStr(.Recordset(7).Value) 'UsrDynamic.LastIP
30:            On Error GoTo Err
31:            .Recordset.MoveNext
32:        Loop
33:    End With

35:    txtDBRegCount.Text = CInt(lvwRegistered.ListItems.count)

37:    Exit Sub
38:
Err:
40:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.DBGetRegRecord()"
End Sub
'------------------------------------------------------------------------------
'End MS Access Subs
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'System Tray Subs
'------------------------------------------------------------------------------
Public Sub SysTrayUpDate(strToolTip As String)
1:    On Error GoTo Err
2:    ModifyTrayIcon Me, 111&, strToolTip
3:    Exit Sub
4:
Err:
6:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.SysTrayUpDate(" & strToolTip & ")"
End Sub
Private Sub SysTrayAdd()
1:    On Error GoTo Err
2:    Dim strTmp As String
    
     'Set sounds to be enabled by default
5:    m_bSound = True
              
      'This gets us a globaly unique ID so that we can be sure the message
      'we use for getting our programs messages is unique
9:    WM_TRAYHOOK = RegisterWindowMessage(GetGUID())
    
      'This retrieves the window message for when the taskbar is created
      'since usually the application is run after the taskbar is created
      'it is safe to assume that if your program receives this message
      'any icon in the tray that was there is now gone and needs to be
      'recreated with a call to Shell_NotifyIcon(NIM_ADD, x)
16:    mTaskbarCreated = RegisterWindowMessage("PTDCH")
    
18:    If Len(g_objSettings.HubName) > 22 Then _
            strTmp = Left(g_objSettings.HubName, 20) & ".." _
       Else strTmp = g_objSettings.HubName
       
       'Create the tray icon
23:    CreateTrayIcon frmHub, 111&, "PT DC Hub " & vbVersion & vbNewLine & vbNewLine & strTmp
       
       'Start the message hook
26:    m_lHookID = InsertHook(frmHub)

28:    Exit Sub

30:
Err:
32:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.SysTrayAdd()"
End Sub
Private Sub SysTrayRem()
1:    On Error GoTo Err

      'Remove system tray icon
4:    DeleteTrayIcon 111&
      'Remove the message hook  <=!!!IMPORTANT!!!
6:    RemoveHook frmHub, m_lHookID
      
8:  Exit Sub
    
10:
Err:
12:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.SysTrayRem()"
End Sub
Friend Function WindowProcSysTray( _
    ByVal shWnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

2:    On Error GoTo Err

    'Friend Function for sys tray
    'This is our message handler
    
7:     If shWnd = Me.hWnd Then 'First we check to see if the message is for this window
           Select Case uMsg    'Then we look at the message
              Case mTaskbarCreated    'This message is for when the taskbar is created
                  'if the taskbar was created, chances are explorer.exe had crashed
11:                CreateTrayIcon Me, 111&, "PT Direct Connect Hub " & vbVersion, Me.Icon  'recreate the tray icon
        
               Case WM_TRAYHOOK 'Our user defined window message
                'if we get this we know that lParam carries the "event"
                'that occured on the tray icon
            
17:                Select Case lParam
                      Case WM_LBUTTONDBLCLK   'Left button dbl clicked
                    
20:                        If Me.WindowState = vbMinimized Then

22:                            SetForegroundWindow Me.hWnd
23:                            Me.WindowState = vbNormal
24:                            Me.Show
25:                        End If
                        
                      Case WM_RBUTTONUP   'Right button released
                    
28:                        SetForegroundWindow Me.hWnd
29:                        RemoveBalloon Me, 111&
30:                        PopupMenu Me.mnuPopUp(0)
                    
32:                    Case NIN_BALLOONUSERCLICK
                          'User clicked the balloon.
                          '
35:                    Case NIN_BALLOONTIMEOUT
                          'Balloon disapeared floated away, or was dismissed.
37:              End Select
    
39:        End Select
40:    End If

    'also pass them to VB
43:    WindowProcSysTray = CallWindowProc(m_lHookID, shWnd, uMsg, wParam, lParam)
    
45:  Exit Function

47:
Err:
49:    HandleError Err.Number, Err.Description, Erl & "|" & "frmHub.WindowProcSysTray(""" & shWnd & """, """ & uMsg & """, """ & wParam & """, """ & lParam & """)"
End Function
'------------------------------------------------------------------------------
'End System Tray Subs
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'Paint events
'------------------------------------------------------------------------------
Private Sub Form_Paint()
1: If g_objSettings.blSkin Then _
     PaintTileFormBackground Me, iResPic(g_objSettings.lngSkin)
End Sub
Private Sub picBordTab_Paint(Index As Integer)
1: If g_objSettings.blSkin Then _
   PaintTilePicBackground Me.picBordTab(Index), iResPic(g_objSettings.lngSkin)
End Sub
Private Sub picHelp_Paint(Index As Integer)
1: If g_objSettings.blSkin Then _
    PaintTilePicBackground Me.picHelp(Index), iResPic(g_objSettings.lngSkin)
End Sub
Private Sub picInfo_Paint(Index As Integer)
1: If g_objSettings.blSkin Then _
    PaintTilePicBackground Me.picInfo(Index), iResPic(g_objSettings.lngSkin)
End Sub
Private Sub picITab_Paint(Index As Integer)
1: If g_objSettings.blSkin Then _
    PaintTilePicBackground Me.picITab(Index), iResPic(g_objSettings.lngSkin)
End Sub
Private Sub picSTab_Paint(Index As Integer)
1: If g_objSettings.blSkin Then _
    PaintTilePicBackground Me.picSTab(Index), iResPic(g_objSettings.lngSkin)
End Sub
Private Sub picStatus_Paint(Index As Integer)
1: If g_objSettings.blSkin Then _
    PaintTilePicBackground Me.picStatus(Index), iResPic(g_objSettings.lngSkin)
End Sub
Private Sub picStInfo_Paint()
1: If g_objSettings.blSkin Then _
    PaintTilePicBackground Me.picStInfo, iResPic(g_objSettings.lngSkin)
End Sub
Private Sub picTab_Paint(Index As Integer)
1: If g_objSettings.blSkin Then _
    PaintTilePicBackground Me.picTab(Index), iResPic(g_objSettings.lngSkin)
End Sub
Private Sub picTabAdv_Paint(Index As Integer)
1: If g_objSettings.blSkin Then _
        PaintTilePicBackground Me.picTabAdv(Index), iResPic(g_objSettings.lngSkin)
End Sub
'------------------------------------------------------------------------------
'End Paint events
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'End Move Form events
'------------------------------------------------------------------------------
Private Sub pgrMemory_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub
Private Sub picBordTab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub
Private Sub picHelp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub
Private Sub picInfo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub
Private Sub picITab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub
Private Sub picLog_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If Not Index = 2 Then _
         If g_objSettings.MoveForm Then _
              Call frmMove(Me)
End Sub
Private Sub picSTab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub
Private Sub picStatus_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub
Private Sub picStatus_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   picStInfo.Visible = False
End Sub
Private Sub picTab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub
Private Sub picTabAdv_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub
Private Sub stbMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub
Private Sub lblOptBanFilter_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub
Private Sub lblOptJM_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub
Private Sub lblOptRedirect_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub
Private Sub lblOptStSend_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub
Private Sub Labels_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub
Private Sub lblCheck_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub
Private Sub lblHolder_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub
Private Sub lblShadowed_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub
Private Sub lblTitle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
1:   If g_objSettings.MoveForm Then _
          Call frmMove(Me)
End Sub
'------------------------------------------------------------------------------
'End Move Form events
'------------------------------------------------------------------------------
