VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form7 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About SRI Calendar"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10395
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   1095
      Left            =   6720
      TabIndex        =   1
      Top             =   480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1931
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Version 1.0"
      TabPicture(0)   =   "celaboutfrm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Text1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Version 2.0"
      TabPicture(1)   =   "celaboutfrm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text2"
      Tab(1).ControlCount=   1
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "celaboutfrm.frx":0038
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   555
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "celaboutfrm.frx":005E
         Top             =   360
         Width           =   2775
      End
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
      TabIndex        =   4
      Top             =   0
      Width           =   10455
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   1920
      Picture         =   "celaboutfrm.frx":00A3
      Top             =   5640
      Width           =   720
   End
   Begin VB.Label Label7 
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
      Left            =   360
      TabIndex        =   0
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   6495
      Left            =   0
      Picture         =   "celaboutfrm.frx":0F6D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10455
   End
   Begin VB.Image Image3 
      Height          =   6495
      Left            =   0
      Picture         =   "celaboutfrm.frx":28ABF5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image1.Visible = True
Tips.Caption = " You can know tips from me. "
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image1.Visible = True
Tips.Caption = " You can know tips from me. "
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image1.Visible = False
Tips.Caption = " You can see picture of SRI Calendar ver.1.0 now "
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image1.Visible = True
Tips.Caption = " You can see picture of SRI Calendar ver.2.0 now "
End Sub

