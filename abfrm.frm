VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About SRISOFTwrite 2017"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10575
   Icon            =   "abfrm.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   10575
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   3000
      Left            =   3240
      TabIndex        =   0
      Top             =   600
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   5292
      _Version        =   393216
      Tabs            =   10
      TabsPerRow      =   10
      TabHeight       =   882
      TabCaption(0)   =   "V1.1"
      TabPicture(0)   =   "abfrm.frx":0ECA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "v1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "V1.2"
      TabPicture(1)   =   "abfrm.frx":0EE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "v2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "V1.3"
      TabPicture(2)   =   "abfrm.frx":0F02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "v3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "V1.4"
      TabPicture(3)   =   "abfrm.frx":0F1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "v4"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "V1.5"
      TabPicture(4)   =   "abfrm.frx":0F3A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "v5"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "V1.6"
      TabPicture(5)   =   "abfrm.frx":0F56
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "v6"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "V1.7"
      TabPicture(6)   =   "abfrm.frx":0F72
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "v7"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "V1.8"
      TabPicture(7)   =   "abfrm.frx":0F8E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "v8"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "V2016"
      TabPicture(8)   =   "abfrm.frx":0FAA
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "v2016"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "V.2017"
      TabPicture(9)   =   "abfrm.frx":0FC6
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "v2017"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).ControlCount=   1
      Begin RichTextLib.RichTextBox v1 
         Height          =   2055
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3625
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"abfrm.frx":0FE2
      End
      Begin RichTextLib.RichTextBox v2 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   2
         Top             =   720
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3625
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"abfrm.frx":1064
      End
      Begin RichTextLib.RichTextBox v3 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   3
         Top             =   720
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3625
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"abfrm.frx":10E6
      End
      Begin RichTextLib.RichTextBox v4 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   4
         Top             =   720
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3625
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"abfrm.frx":1168
      End
      Begin RichTextLib.RichTextBox v5 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   5
         Top             =   720
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3625
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"abfrm.frx":11EA
      End
      Begin RichTextLib.RichTextBox v6 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   6
         Top             =   720
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3625
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"abfrm.frx":126C
      End
      Begin RichTextLib.RichTextBox v7 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   7
         Top             =   720
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3625
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"abfrm.frx":12EE
      End
      Begin RichTextLib.RichTextBox v8 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   8
         Top             =   720
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3625
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"abfrm.frx":1370
      End
      Begin RichTextLib.RichTextBox v2016 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   9
         Top             =   720
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3625
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"abfrm.frx":13F2
      End
      Begin RichTextLib.RichTextBox v2017 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   10
         Top             =   720
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3625
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"abfrm.frx":1474
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SRISOFTwrite 2017"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   1680
      Picture         =   "abfrm.frx":14F6
      Top             =   480
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "Version updates"
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   3240
      TabIndex        =   12
      Top             =   360
      Width           =   1260
   End
   Begin VB.Label Tips 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " You can know tips from me. "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   10575
   End
   Begin VB.Image Image1 
      Height          =   9015
      Left            =   0
      Picture         =   "abfrm.frx":23C0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10575
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
v1.LoadFile (App.Path & "\Files\Documents\SRI SOFT Write 1.1.RTF")
v2.LoadFile (App.Path & "\Files\Documents\SRI SOFT Write 1.2.RTF")
v3.LoadFile (App.Path & "\Files\Documents\SRI SOFT Write 1.3.RTF")
v4.LoadFile (App.Path & "\Files\Documents\SRI SOFT Write 1.4.RTF")
v5.LoadFile (App.Path & "\Files\Documents\SRI SOFT Write 1.5.RTF")
v6.LoadFile (App.Path & "\Files\Documents\SRI SOFT Write 1.6.RTF")
v7.LoadFile (App.Path & "\Files\Documents\SRI SOFT Write 1.7.RTF")
v8.LoadFile (App.Path & "\Files\Documents\SRI SOFT Write 1.8.RTF")
v2016.LoadFile (App.Path & "\Files\Documents\SRI SOFT Write 2016.RTF")
v2017.LoadFile (App.Path & "\Files\Documents\SRI SOFT Write 2017.RTF")
End Sub


Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
SSTab1.Visible = False
Tips.Caption = " To see version updates, carry mouse to the 'Version Updates' box. "
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
SSTab1.Visible = True
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can see old versions' updates and new version's updates from here using tabs. "
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

End Sub

Private Sub Tips_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can know tips from me. "
End Sub
