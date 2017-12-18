VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form9 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About SRI Calculator"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10560
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1931
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Version 1.0"
      TabPicture(0)   =   "calaboutfrm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Text1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Version 2.0"
      TabPicture(1)   =   "calaboutfrm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text2"
      Tab(1).ControlCount=   1
      Begin VB.TextBox Text2 
         Height          =   555
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "calaboutfrm.frx":0038
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "calaboutfrm.frx":0074
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
      Width           =   10575
   End
   Begin VB.Image Image3 
      Height          =   720
      Left            =   9720
      Picture         =   "calaboutfrm.frx":009A
      Top             =   360
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
      Height          =   255
      Left            =   8160
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   6855
      Left            =   0
      Picture         =   "calaboutfrm.frx":0F64
      Stretch         =   -1  'True
      Top             =   240
      Width           =   10575
   End
   Begin VB.Image Image2 
      Height          =   6855
      Left            =   0
      Picture         =   "calaboutfrm.frx":110ABE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10575
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image1.Visible = True
Tips.Caption = " You can know tips from me. "
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image1.Visible = True
Tips.Caption = " You can know tips from me. "
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image1.Visible = False
Tips.Caption = " You can see picture of SRI Calculator ver.1.0 now "
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Image1.Visible = True
Tips.Caption = " You can see picture of SRI Calculator ver.2.0 now "
End Sub
