VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form10 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About Song Player"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10560
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   735
      Left            =   6720
      TabIndex        =   0
      Top             =   480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1296
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Version 1.0"
      TabPicture(0)   =   "sonaboutfrm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Text1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "sonaboutfrm.frx":001C
         Top             =   360
         Width           =   2775
      End
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
      Left            =   1200
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   2760
      Picture         =   "sonaboutfrm.frx":0043
      Top             =   360
      Width           =   720
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
      TabIndex        =   2
      Top             =   0
      Width           =   10575
   End
   Begin VB.Image Image1 
      Height          =   6615
      Left            =   0
      Picture         =   "sonaboutfrm.frx":0F0D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10575
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can see picture of SRI Song Player ver.1.0 now "
End Sub

Private Sub Tips_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can know tips from me. "
End Sub
