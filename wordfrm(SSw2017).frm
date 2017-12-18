VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " SRISOFTwrite 2017 "
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   10320
   DrawWidth       =   3
   Icon            =   "wordfrm(SSw2017).frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   10320
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox second 
      Height          =   285
      Left            =   5520
      TabIndex        =   135
      Text            =   "Seconds"
      Top             =   8400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5655
      Left            =   120
      TabIndex        =   120
      Top             =   120
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9975
      _Version        =   393217
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"wordfrm(SSw2017).frx":0ECA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   15
      Text            =   "192"
      Top             =   8400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   840
      MaxLength       =   3
      TabIndex        =   14
      Text            =   "192"
      Top             =   8400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   240
      MaxLength       =   3
      TabIndex        =   13
      Text            =   "255"
      Top             =   8400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6840
      MaxLength       =   2
      TabIndex        =   12
      Text            =   "10"
      Top             =   8400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox scres 
      Height          =   285
      Left            =   5400
      TabIndex        =   10
      Text            =   "None"
      Top             =   5160
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox screensaver 
      Height          =   285
      Left            =   5400
      TabIndex        =   9
      Text            =   "0"
      Top             =   5520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox hrs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   960
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   8
      Text            =   "0"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox min 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   3240
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   6
      Text            =   "0"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox sec 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   5520
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   4
      Text            =   "0"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox dem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   7800
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "00"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   8850
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   397
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "4:25 PM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "7/24/2017"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   6588
            Text            =   "SRISOFTwrite 2017"
            TextSave        =   "SRISOFTwrite 2017"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2415
      Left            =   120
      TabIndex        =   16
      Top             =   5880
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   4260
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      MouseIcon       =   "wordfrm(SSw2017).frx":0F41
      TabCaption(0)   =   "Menu"
      TabPicture(0)   =   "wordfrm(SSw2017).frx":259B
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Command1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command11"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Edit"
      TabPicture(1)   =   "wordfrm(SSw2017).frx":25B7
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cc"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command8"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command9"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Command19"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command10"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command79"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Font"
      TabPicture(2)   =   "wordfrm(SSw2017).frx":25D3
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Command73"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Tools"
      TabPicture(3)   =   "wordfrm(SSw2017).frx":25EF
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command71"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Command72"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Command74"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Settings"
      TabPicture(4)   =   "wordfrm(SSw2017).frx":260B
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "SSTab3"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.CommandButton Command74 
         BackColor       =   &H00EFEFEF&
         Height          =   855
         Left            =   -72720
         Picture         =   "wordfrm(SSw2017).frx":2627
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Song Player"
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton Command72 
         BackColor       =   &H00EFEFEF&
         Height          =   855
         Left            =   -73680
         Picture         =   "wordfrm(SSw2017).frx":34F1
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Calendar"
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton Command79 
         BackColor       =   &H00EFEFEF&
         Height          =   615
         Left            =   -71640
         Picture         =   "wordfrm(SSw2017).frx":43BB
         Style           =   1  'Graphical
         TabIndex        =   129
         ToolTipText     =   "Delete all"
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00EFEFEF&
         Height          =   615
         Left            =   -72240
         Picture         =   "wordfrm(SSw2017).frx":49E5
         Style           =   1  'Graphical
         TabIndex        =   130
         ToolTipText     =   "Delete"
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command19 
         BackColor       =   &H00EFEFEF&
         Height          =   615
         Left            =   -72840
         Picture         =   "wordfrm(SSw2017).frx":500F
         Style           =   1  'Graphical
         TabIndex        =   131
         ToolTipText     =   "Duplicator"
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00EFEFEF&
         Height          =   615
         Left            =   -73440
         Picture         =   "wordfrm(SSw2017).frx":5591
         Style           =   1  'Graphical
         TabIndex        =   132
         ToolTipText     =   "Paste"
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00EFEFEF&
         Height          =   615
         Left            =   -74040
         Picture         =   "wordfrm(SSw2017).frx":5B13
         Style           =   1  'Graphical
         TabIndex        =   133
         ToolTipText     =   "Copy"
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FAFAFA&
         Height          =   735
         Left            =   4560
         Picture         =   "wordfrm(SSw2017).frx":6095
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Close"
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FAFAFA&
         Height          =   735
         Left            =   3720
         Picture         =   "wordfrm(SSw2017).frx":6F77
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Print"
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FAFAFA&
         Height          =   735
         Left            =   2880
         Picture         =   "wordfrm(SSw2017).frx":7E59
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Save As"
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FAFAFA&
         Height          =   735
         Left            =   2040
         Picture         =   "wordfrm(SSw2017).frx":8D3B
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Save"
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FAFAFA&
         Height          =   735
         Left            =   1200
         Picture         =   "wordfrm(SSw2017).frx":9C1D
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Open"
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00EFEFEF&
         Height          =   615
         Left            =   -74640
         Picture         =   "wordfrm(SSw2017).frx":AAFF
         Style           =   1  'Graphical
         TabIndex        =   134
         ToolTipText     =   "Cut"
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H008080FF&
         Caption         =   "Emergency Exit"
         Height          =   375
         Left            =   8400
         MaskColor       =   &H00400040&
         Style           =   1  'Graphical
         TabIndex        =   112
         ToolTipText     =   "Emergency Exit"
         Top             =   1560
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.Timer cc 
         Interval        =   1
         Left            =   -65520
         Top             =   1320
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FAFAFA&
         Height          =   735
         Left            =   360
         Picture         =   "wordfrm(SSw2017).frx":B081
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "New"
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command71 
         BackColor       =   &H00EFEFEF&
         Height          =   855
         Left            =   -74640
         Picture         =   "wordfrm(SSw2017).frx":BF63
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Calculator"
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton Command73 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Font"
         Height          =   1815
         Left            =   -66480
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Font properties"
         Top             =   120
         Width           =   1455
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   17
         Top             =   120
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   3413
         _Version        =   393216
         TabOrientation  =   1
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         MouseIcon       =   "wordfrm(SSw2017).frx":CE2D
         TabCaption(0)   =   "Screen saver"
         TabPicture(0)   =   "wordfrm(SSw2017).frx":E487
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label6"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "SSTab4"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "List1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "tt"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Command75"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Combo8"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Text2"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "Back color"
         TabPicture(1)   =   "wordfrm(SSw2017).frx":E4A3
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame10"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Cursors"
         TabPicture(2)   =   "wordfrm(SSw2017).frx":E4BF
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Drive1"
         Tab(2).Control(1)=   "Dir1"
         Tab(2).Control(2)=   "File1"
         Tab(2).Control(3)=   "Frame11"
         Tab(2).Control(4)=   "Command84"
         Tab(2).Control(5)=   "Command83"
         Tab(2).ControlCount=   6
         TabCaption(3)   =   "About"
         TabPicture(3)   =   "wordfrm(SSw2017).frx":E4DB
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Command78"
         Tab(3).Control(1)=   "Command80"
         Tab(3).Control(2)=   "Command81"
         Tab(3).Control(3)=   "Command82"
         Tab(3).Control(4)=   "Command86"
         Tab(3).ControlCount=   5
         Begin VB.CommandButton Command86 
            Caption         =   "About maker"
            Height          =   1335
            Left            =   -71520
            TabIndex        =   176
            ToolTipText     =   "Get abot the maker"
            Top             =   120
            Width           =   1575
         End
         Begin VB.CommandButton Command83 
            Caption         =   "Set"
            Height          =   855
            Left            =   -74760
            TabIndex        =   128
            ToolTipText     =   "set cursor at select path"
            Top             =   600
            Width           =   2055
         End
         Begin VB.CommandButton Command84 
            Caption         =   "No icons"
            Height          =   1455
            Left            =   -67680
            TabIndex        =   126
            ToolTipText     =   "Remove text icon"
            Top             =   120
            Width           =   615
         End
         Begin VB.Frame Frame11 
            Caption         =   "SSw cursors"
            Height          =   1455
            Left            =   -66960
            TabIndex        =   121
            ToolTipText     =   "SSw cursors"
            Top             =   120
            Width           =   1695
            Begin VB.CommandButton Command85 
               Caption         =   "Set"
               Height          =   195
               Left            =   120
               TabIndex        =   127
               ToolTipText     =   "set SSw cursor"
               Top             =   1200
               Width           =   1455
            End
            Begin VB.OptionButton Option4 
               Caption         =   "pencil icon"
               Height          =   195
               Left            =   120
               TabIndex        =   125
               Top             =   960
               Width           =   1095
            End
            Begin VB.OptionButton Option3 
               Caption         =   "pen icon"
               Height          =   195
               Left            =   120
               TabIndex        =   124
               Top             =   720
               Width           =   975
            End
            Begin VB.OptionButton Option2 
               Caption         =   "eraser icon"
               Height          =   195
               Left            =   120
               TabIndex        =   123
               Top             =   480
               Width           =   1095
            End
            Begin VB.OptionButton Option1 
               Caption         =   "correction icon"
               Height          =   255
               Left            =   120
               TabIndex        =   122
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.FileListBox File1 
            Height          =   1455
            Left            =   -70080
            Pattern         =   "*.cur"
            TabIndex        =   119
            ToolTipText     =   "cursor files"
            Top             =   120
            Width           =   2295
         End
         Begin VB.DirListBox Dir1 
            Height          =   1440
            Left            =   -72600
            TabIndex        =   118
            ToolTipText     =   "cursor path"
            Top             =   120
            Width           =   2415
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   -74760
            TabIndex        =   117
            ToolTipText     =   "driver of path"
            Top             =   120
            Width           =   2055
         End
         Begin VB.CommandButton Command82 
            Caption         =   "About Sri Song Player"
            Height          =   375
            Left            =   -73320
            TabIndex        =   116
            ToolTipText     =   "Get about SRISon. 1.0 "
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CommandButton Command81 
            Caption         =   "About Sri Calculator"
            Height          =   375
            Left            =   -73320
            TabIndex        =   115
            ToolTipText     =   "Get about SRICal. 2.0 & past versions"
            Top             =   600
            Width           =   1695
         End
         Begin VB.CommandButton Command80 
            Caption         =   "About Sri Calendar"
            Height          =   375
            Left            =   -73320
            TabIndex        =   114
            ToolTipText     =   "Get about SRICel. 2.0 & past versions"
            Top             =   120
            Width           =   1695
         End
         Begin VB.CommandButton Command78 
            Caption         =   "About SRISOFTwrite 2017"
            Height          =   1335
            Left            =   -74760
            TabIndex        =   113
            ToolTipText     =   "Get about SSw 2017 & past versions"
            Top             =   120
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   8880
            MaxLength       =   2
            TabIndex        =   29
            Text            =   "10"
            ToolTipText     =   "Wait time"
            Top             =   360
            Width           =   375
         End
         Begin VB.ComboBox Combo8 
            Height          =   315
            ItemData        =   "wordfrm(SSw2017).frx":E4F7
            Left            =   8880
            List            =   "wordfrm(SSw2017).frx":E501
            TabIndex        =   28
            Text            =   "Seconds"
            ToolTipText     =   "Type of time"
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton Command75 
            Caption         =   "Set"
            Height          =   375
            Left            =   7320
            TabIndex        =   27
            ToolTipText     =   "Set Screen saver"
            Top             =   1200
            Width           =   2415
         End
         Begin VB.Timer tt 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   9360
            Top             =   480
         End
         Begin VB.ListBox List1 
            Height          =   1035
            ItemData        =   "wordfrm(SSw2017).frx":E517
            Left            =   7320
            List            =   "wordfrm(SSw2017).frx":E530
            TabIndex        =   26
            ToolTipText     =   "Screen saver"
            Top             =   120
            Width           =   1455
         End
         Begin VB.Frame Frame10 
            Caption         =   "Background color and Text box back color"
            Height          =   1335
            Left            =   -74760
            TabIndex        =   18
            ToolTipText     =   "Background color"
            Top             =   120
            Width           =   4935
            Begin VB.OptionButton Option6 
               Caption         =   "Text box "
               Height          =   255
               Left            =   3600
               TabIndex        =   178
               Top             =   480
               Width           =   1215
            End
            Begin VB.OptionButton Option5 
               Caption         =   "Background"
               Height          =   255
               Left            =   3600
               TabIndex        =   177
               Top             =   240
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.CommandButton Command76 
               Caption         =   "S e t"
               Height          =   1000
               Left            =   3240
               TabIndex        =   25
               Top             =   240
               Width           =   255
            End
            Begin VB.TextBox bdt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1920
               MaxLength       =   3
               TabIndex        =   24
               Text            =   "0"
               Top             =   960
               Width           =   495
            End
            Begin VB.TextBox gdt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1920
               MaxLength       =   3
               TabIndex        =   23
               Text            =   "0"
               Top             =   600
               Width           =   495
            End
            Begin VB.TextBox rdt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1920
               MaxLength       =   3
               TabIndex        =   22
               Text            =   "0"
               Top             =   240
               Width           =   495
            End
            Begin VB.HScrollBar bd 
               Height          =   255
               LargeChange     =   20
               Left            =   120
               Max             =   255
               TabIndex        =   21
               Top             =   960
               Value           =   1
               Width           =   1695
            End
            Begin VB.HScrollBar gd 
               Height          =   255
               LargeChange     =   20
               Left            =   120
               Max             =   255
               TabIndex        =   20
               Top             =   600
               Value           =   1
               Width           =   1695
            End
            Begin VB.HScrollBar rd 
               Height          =   255
               LargeChange     =   20
               Left            =   120
               Max             =   255
               TabIndex        =   19
               Top             =   240
               Value           =   1
               Width           =   1695
            End
            Begin VB.Shape Shape2 
               BackColor       =   &H00000000&
               BackStyle       =   1  'Opaque
               Height          =   1000
               Left            =   2520
               Top             =   240
               Width           =   615
            End
         End
         Begin TabDlg.SSTab SSTab4 
            Height          =   1455
            Left            =   120
            TabIndex        =   30
            ToolTipText     =   "Screen Saver"
            Top             =   120
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   2566
            _Version        =   393216
            TabOrientation  =   1
            Style           =   1
            Tabs            =   5
            TabsPerRow      =   5
            TabHeight       =   520
            TabCaption(0)   =   "Color magic"
            TabPicture(0)   =   "wordfrm(SSw2017).frx":E587
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label1"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label2"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label3"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Label4"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "Label5"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "cmr"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "cm"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "Combo3"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "Combo4"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "Combo5"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "Combo6"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "Text1"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "Combo7"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "Check10"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).ControlCount=   14
            TabCaption(1)   =   "Ping Pong"
            TabPicture(1)   =   "wordfrm(SSw2017).frx":E5A3
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Check2"
            Tab(1).Control(1)=   "Text3"
            Tab(1).Control(2)=   "Combo9"
            Tab(1).Control(3)=   "pp"
            Tab(1).Control(4)=   "rndc"
            Tab(1).Control(5)=   "Label10"
            Tab(1).Control(6)=   "Label8"
            Tab(1).ControlCount=   7
            TabCaption(2)   =   "Objects"
            TabPicture(2)   =   "wordfrm(SSw2017).frx":E5BF
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Check9"
            Tab(2).Control(1)=   "Text4"
            Tab(2).Control(2)=   "Check6"
            Tab(2).Control(3)=   "Check5"
            Tab(2).Control(4)=   "Check3"
            Tab(2).Control(5)=   "sq"
            Tab(2).Control(6)=   "bx"
            Tab(2).Control(7)=   "ru"
            Tab(2).Control(8)=   "bg"
            Tab(2).Control(9)=   "Label11"
            Tab(2).ControlCount=   10
            TabCaption(3)   =   "SRISOFTwrite logo"
            TabPicture(3)   =   "wordfrm(SSw2017).frx":E5DB
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "Text5"
            Tab(3).Control(1)=   "Check7"
            Tab(3).Control(2)=   "sl"
            Tab(3).Control(3)=   "Combo10"
            Tab(3).Control(4)=   "Check4"
            Tab(3).Control(5)=   "Label9"
            Tab(3).ControlCount=   6
            TabCaption(4)   =   "Slide Show"
            TabPicture(4)   =   "wordfrm(SSw2017).frx":E5F7
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "sw"
            Tab(4).Control(1)=   "Check12"
            Tab(4).Control(2)=   "Text7"
            Tab(4).Control(3)=   "ss"
            Tab(4).Control(4)=   "Combo11"
            Tab(4).Control(5)=   "ssr"
            Tab(4).Control(6)=   "Label16"
            Tab(4).ControlCount=   7
            Begin VB.CheckBox Check10 
               Caption         =   "Randomize color"
               Height          =   375
               Left            =   2400
               TabIndex        =   52
               ToolTipText     =   "Color magic color randomize"
               Top             =   720
               Width           =   1575
            End
            Begin VB.CheckBox Check9 
               Caption         =   "Background"
               Height          =   195
               Left            =   -72480
               TabIndex        =   51
               ToolTipText     =   "Objects' background color randomize"
               Top             =   240
               Value           =   1  'Checked
               Width           =   1215
            End
            Begin VB.TextBox Text5 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   -74280
               MaxLength       =   2
               TabIndex        =   50
               Text            =   "10"
               ToolTipText     =   "SSw logo timing"
               Top             =   440
               Width           =   375
            End
            Begin VB.CheckBox Check7 
               Caption         =   "Moveble"
               Height          =   195
               Left            =   -74880
               TabIndex        =   49
               ToolTipText     =   "SSw logo moving"
               Top             =   120
               Width           =   1095
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Randomize color"
               Height          =   375
               Left            =   -72840
               TabIndex        =   48
               ToolTipText     =   "Ping pong color randomize"
               Top             =   120
               Width           =   1575
            End
            Begin VB.TextBox Text3 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   -74280
               MaxLength       =   2
               TabIndex        =   47
               Text            =   "10"
               ToolTipText     =   "Ping pong timinig"
               Top             =   560
               Width           =   375
            End
            Begin VB.ComboBox Combo9 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "wordfrm(SSw2017).frx":E613
               Left            =   -74280
               List            =   "wordfrm(SSw2017).frx":E629
               Sorted          =   -1  'True
               TabIndex        =   46
               Text            =   "Red"
               ToolTipText     =   "Ping pong color"
               Top             =   120
               Width           =   1335
            End
            Begin VB.TextBox Text4 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   -74280
               MaxLength       =   2
               TabIndex        =   45
               Text            =   "10"
               ToolTipText     =   "Objects' timing"
               Top             =   600
               Width           =   375
            End
            Begin VB.CheckBox Check6 
               Caption         =   "Round"
               Height          =   195
               Left            =   -73320
               TabIndex        =   44
               ToolTipText     =   "Objects' rounds"
               Top             =   240
               Value           =   1  'Checked
               Width           =   855
            End
            Begin VB.CheckBox Check5 
               Caption         =   "Box"
               Height          =   195
               Left            =   -74040
               TabIndex        =   43
               ToolTipText     =   "Objects' filled square"
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox Check3 
               Caption         =   "Squre"
               Height          =   195
               Left            =   -74880
               TabIndex        =   42
               ToolTipText     =   "Objects' squre"
               Top             =   240
               Width           =   855
            End
            Begin VB.ComboBox Combo7 
               Height          =   315
               ItemData        =   "wordfrm(SSw2017).frx":E655
               Left            =   1200
               List            =   "wordfrm(SSw2017).frx":E65F
               TabIndex        =   41
               Text            =   "Seconds"
               ToolTipText     =   "Color magic timing type"
               Top             =   720
               Width           =   1095
            End
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   720
               MaxLength       =   2
               TabIndex        =   40
               Text            =   "10"
               ToolTipText     =   "Color magic timinig"
               Top             =   740
               Width           =   375
            End
            Begin VB.ComboBox Combo6 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "wordfrm(SSw2017).frx":E675
               Left            =   4080
               List            =   "wordfrm(SSw2017).frx":E68B
               Sorted          =   -1  'True
               TabIndex        =   39
               Text            =   "Red"
               ToolTipText     =   "Color magic fourth color"
               Top             =   360
               Width           =   1335
            End
            Begin VB.ComboBox Combo5 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "wordfrm(SSw2017).frx":E6B7
               Left            =   2760
               List            =   "wordfrm(SSw2017).frx":E6CD
               Sorted          =   -1  'True
               TabIndex        =   38
               Text            =   "Yellow"
               ToolTipText     =   "Color magic thrid color"
               Top             =   360
               Width           =   1335
            End
            Begin VB.ComboBox Combo4 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "wordfrm(SSw2017).frx":E6F9
               Left            =   1440
               List            =   "wordfrm(SSw2017).frx":E70F
               Sorted          =   -1  'True
               TabIndex        =   37
               Text            =   "Green"
               ToolTipText     =   "Color magic second color"
               Top             =   360
               Width           =   1335
            End
            Begin VB.ComboBox Combo3 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "wordfrm(SSw2017).frx":E73B
               Left            =   120
               List            =   "wordfrm(SSw2017).frx":E751
               Sorted          =   -1  'True
               TabIndex        =   36
               Text            =   "Blue"
               ToolTipText     =   "Color magic first color"
               Top             =   360
               Width           =   1335
            End
            Begin VB.Timer cm 
               Enabled         =   0   'False
               Left            =   6600
               Top             =   600
            End
            Begin VB.Timer cmr 
               Enabled         =   0   'False
               Interval        =   1000
               Left            =   6600
               Top             =   480
            End
            Begin VB.Timer pp 
               Enabled         =   0   'False
               Left            =   -68400
               Top             =   600
            End
            Begin VB.Timer rndc 
               Enabled         =   0   'False
               Interval        =   1000
               Left            =   -68400
               Top             =   480
            End
            Begin VB.Timer sq 
               Enabled         =   0   'False
               Left            =   -71040
               Top             =   600
            End
            Begin VB.Timer bx 
               Enabled         =   0   'False
               Left            =   -70440
               Top             =   600
            End
            Begin VB.Timer ru 
               Enabled         =   0   'False
               Left            =   -69840
               Top             =   600
            End
            Begin VB.Timer bg 
               Enabled         =   0   'False
               Interval        =   1000
               Left            =   -69240
               Top             =   600
            End
            Begin VB.Timer sl 
               Enabled         =   0   'False
               Left            =   -69120
               Top             =   600
            End
            Begin VB.ComboBox Combo10 
               Height          =   315
               ItemData        =   "wordfrm(SSw2017).frx":E77D
               Left            =   -73800
               List            =   "wordfrm(SSw2017).frx":E787
               TabIndex        =   35
               Text            =   "Seconds"
               ToolTipText     =   "SSw logo timing type"
               Top             =   420
               Width           =   1095
            End
            Begin VB.CheckBox Check4 
               Caption         =   "Logo color background"
               Height          =   195
               Left            =   -73800
               TabIndex        =   34
               ToolTipText     =   "SSw logo color background"
               Top             =   120
               Width           =   2295
            End
            Begin VB.Timer sw 
               Enabled         =   0   'False
               Interval        =   1
               Left            =   -70800
               Top             =   840
            End
            Begin VB.CheckBox Check12 
               Caption         =   "Suffle pictures"
               Height          =   375
               Left            =   -74880
               TabIndex        =   33
               ToolTipText     =   "Slide show picture suffle"
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox Text7 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   -74280
               MaxLength       =   2
               TabIndex        =   32
               Text            =   "10"
               ToolTipText     =   "Slide show timing"
               Top             =   120
               Width           =   375
            End
            Begin VB.Timer ss 
               Enabled         =   0   'False
               Left            =   -68640
               Top             =   480
            End
            Begin VB.ComboBox Combo11 
               Height          =   315
               ItemData        =   "wordfrm(SSw2017).frx":E79D
               Left            =   -73800
               List            =   "wordfrm(SSw2017).frx":E7A7
               TabIndex        =   31
               Text            =   "Seconds"
               ToolTipText     =   "Slide show timing type"
               Top             =   110
               Width           =   1095
            End
            Begin VB.Timer ssr 
               Left            =   -68640
               Top             =   240
            End
            Begin VB.Label Label9 
               Caption         =   "Speed"
               Height          =   255
               Left            =   -74880
               TabIndex        =   63
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Label11 
               Caption         =   "Timinig"
               Height          =   255
               Left            =   -74880
               TabIndex        =   62
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Label10 
               Caption         =   "Color"
               Height          =   255
               Left            =   -74880
               TabIndex        =   61
               Top             =   120
               Width           =   615
            End
            Begin VB.Label Label8 
               Caption         =   "Timinig"
               Height          =   255
               Left            =   -74880
               TabIndex        =   60
               Top             =   600
               Width           =   615
            End
            Begin VB.Label Label7 
               Caption         =   "Color"
               Height          =   255
               Left            =   120
               TabIndex        =   59
               Top             =   -480
               Width           =   1575
            End
            Begin VB.Label Label5 
               Caption         =   "Timinig"
               Height          =   255
               Left            =   120
               TabIndex        =   58
               Top             =   780
               Width           =   615
            End
            Begin VB.Label Label4 
               Caption         =   "Fourth color"
               Height          =   255
               Left            =   4080
               TabIndex        =   57
               Top             =   120
               Width           =   1575
            End
            Begin VB.Label Label3 
               Caption         =   "Third color"
               Height          =   255
               Left            =   2760
               TabIndex        =   56
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label Label2 
               Caption         =   "Second color"
               Height          =   255
               Left            =   1440
               TabIndex        =   55
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   "First color"
               Height          =   255
               Left            =   120
               TabIndex        =   54
               Top             =   120
               Width           =   855
            End
            Begin VB.Label Label16 
               Caption         =   "Speed"
               Height          =   255
               Left            =   -74880
               TabIndex        =   53
               Top             =   165
               Width           =   735
            End
         End
         Begin VB.Label Label6 
            Caption         =   "Wait"
            Height          =   255
            Left            =   8880
            TabIndex        =   64
            Top             =   120
            Width           =   495
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   75
         Top             =   120
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   3413
         _Version        =   393216
         TabOrientation  =   1
         Style           =   1
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         MouseIcon       =   "wordfrm(SSw2017).frx":E7BD
         TabCaption(0)   =   "Style"
         TabPicture(0)   =   "wordfrm(SSw2017).frx":FE17
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame7"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame5"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Effects"
         TabPicture(1)   =   "wordfrm(SSw2017).frx":FE33
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Command15"
         Tab(1).Control(1)=   "Command14"
         Tab(1).Control(2)=   "Command13"
         Tab(1).Control(3)=   "Command12"
         Tab(1).ControlCount=   4
         TabCaption(2)   =   "Alignment"
         TabPicture(2)   =   "wordfrm(SSw2017).frx":FE4F
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Command18"
         Tab(2).Control(1)=   "Command17"
         Tab(2).Control(2)=   "align"
         Tab(2).Control(3)=   "Command16"
         Tab(2).ControlCount=   4
         TabCaption(3)   =   "Case and Space"
         TabPicture(3)   =   "wordfrm(SSw2017).frx":FE6B
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame1"
         Tab(3).Control(1)=   "Frame2"
         Tab(3).ControlCount=   2
         TabCaption(4)   =   "Color"
         TabPicture(4)   =   "wordfrm(SSw2017).frx":FE87
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame8"
         Tab(4).Control(1)=   "Frame9"
         Tab(4).ControlCount=   2
         Begin VB.CommandButton Command15 
            BackColor       =   &H00EFEFEF&
            Height          =   615
            Left            =   -72960
            Picture         =   "wordfrm(SSw2017).frx":FEA3
            Style           =   1  'Graphical
            TabIndex        =   105
            ToolTipText     =   "Strike throgh"
            Top             =   600
            Width           =   615
         End
         Begin VB.CommandButton Command14 
            BackColor       =   &H00EFEFEF&
            Height          =   615
            Left            =   -73560
            Picture         =   "wordfrm(SSw2017).frx":104CD
            Style           =   1  'Graphical
            TabIndex        =   106
            ToolTipText     =   "Underline"
            Top             =   600
            Width           =   615
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H00EFEFEF&
            Height          =   615
            Left            =   -74160
            Picture         =   "wordfrm(SSw2017).frx":10AF7
            Style           =   1  'Graphical
            TabIndex        =   107
            ToolTipText     =   "Italic"
            Top             =   600
            Width           =   615
         End
         Begin VB.CommandButton Command18 
            BackColor       =   &H00EFEFEF&
            Height          =   615
            Left            =   -73560
            Picture         =   "wordfrm(SSw2017).frx":11121
            Style           =   1  'Graphical
            TabIndex        =   109
            ToolTipText     =   "Right align"
            Top             =   600
            Width           =   615
         End
         Begin VB.CommandButton Command17 
            BackColor       =   &H00EFEFEF&
            Height          =   615
            Left            =   -74160
            Picture         =   "wordfrm(SSw2017).frx":1174B
            Style           =   1  'Graphical
            TabIndex        =   110
            ToolTipText     =   "Center align"
            Top             =   600
            Width           =   615
         End
         Begin VB.Timer align 
            Interval        =   1
            Left            =   -71880
            Top             =   480
         End
         Begin VB.CommandButton Command16 
            BackColor       =   &H00EFEFEF&
            Height          =   615
            Left            =   -74760
            Picture         =   "wordfrm(SSw2017).frx":11D75
            Style           =   1  'Graphical
            TabIndex        =   111
            ToolTipText     =   "Left align"
            Top             =   600
            Width           =   615
         End
         Begin VB.CommandButton Command12 
            BackColor       =   &H00EFEFEF&
            Height          =   615
            Left            =   -74760
            Picture         =   "wordfrm(SSw2017).frx":1239F
            Style           =   1  'Graphical
            TabIndex        =   108
            ToolTipText     =   "Bold"
            Top             =   600
            Width           =   615
         End
         Begin VB.Frame Frame1 
            Caption         =   "Case"
            Height          =   1335
            Left            =   -74880
            TabIndex        =   100
            Top             =   120
            Width           =   2655
            Begin VB.CommandButton Command23 
               BackColor       =   &H00FFFFFF&
               Height          =   615
               Left            =   1920
               Picture         =   "wordfrm(SSw2017).frx":129C9
               Style           =   1  'Graphical
               TabIndex        =   101
               ToolTipText     =   "First letter simp"
               Top             =   480
               Width           =   615
            End
            Begin VB.CommandButton Command22 
               BackColor       =   &H00FFFFFF&
               Height          =   615
               Left            =   1320
               Picture         =   "wordfrm(SSw2017).frx":12FF3
               Style           =   1  'Graphical
               TabIndex        =   102
               ToolTipText     =   "First letter caps"
               Top             =   480
               Width           =   615
            End
            Begin VB.CommandButton Command21 
               BackColor       =   &H00FFFFFF&
               Height          =   615
               Left            =   720
               Picture         =   "wordfrm(SSw2017).frx":1361D
               Style           =   1  'Graphical
               TabIndex        =   103
               ToolTipText     =   "Simpal (Lower case)"
               Top             =   480
               Width           =   615
            End
            Begin VB.CommandButton Command20 
               BackColor       =   &H00FFFFFF&
               Height          =   615
               Left            =   120
               Picture         =   "wordfrm(SSw2017).frx":13C47
               Style           =   1  'Graphical
               TabIndex        =   104
               ToolTipText     =   "Capital (Upper case)"
               Top             =   480
               Width           =   615
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Space"
            Height          =   1335
            Left            =   -72120
            TabIndex        =   94
            Top             =   120
            Width           =   2775
            Begin VB.Frame Frame3 
               Height          =   975
               Left            =   120
               TabIndex        =   97
               ToolTipText     =   "Add spaces"
               Top             =   240
               Width           =   1575
               Begin VB.TextBox NoSp 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   390
                  Left            =   840
                  MaxLength       =   1
                  TabIndex        =   99
                  Text            =   "1"
                  Top             =   360
                  Width           =   495
               End
               Begin VB.CommandButton Command24 
                  BackColor       =   &H00FFFFFF&
                  Height          =   615
                  Left            =   120
                  Picture         =   "wordfrm(SSw2017).frx":14271
                  Style           =   1  'Graphical
                  TabIndex        =   98
                  ToolTipText     =   "Add"
                  Top             =   240
                  Width           =   615
               End
            End
            Begin VB.Frame Frame4 
               Height          =   975
               Left            =   1800
               TabIndex        =   95
               Top             =   240
               Width           =   855
               Begin VB.CommandButton Command25 
                  BackColor       =   &H00FFFFFF&
                  Height          =   615
                  Left            =   120
                  Picture         =   "wordfrm(SSw2017).frx":1489B
                  Style           =   1  'Graphical
                  TabIndex        =   96
                  ToolTipText     =   "Remove spaces"
                  Top             =   240
                  Width           =   615
               End
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Size"
            Height          =   1335
            Left            =   120
            TabIndex        =   87
            Top             =   120
            Width           =   3615
            Begin VB.CommandButton Command27 
               BackColor       =   &H00EFEFEF&
               Height          =   615
               Left            =   720
               Picture         =   "wordfrm(SSw2017).frx":14EC5
               Style           =   1  'Graphical
               TabIndex        =   92
               ToolTipText     =   "Font size decrease"
               Top             =   600
               Width           =   615
            End
            Begin VB.CommandButton Command26 
               BackColor       =   &H00EFEFEF&
               Height          =   615
               Left            =   120
               Picture         =   "wordfrm(SSw2017).frx":154EF
               Style           =   1  'Graphical
               TabIndex        =   93
               ToolTipText     =   "Font size increase"
               Top             =   600
               Width           =   615
            End
            Begin VB.HScrollBar HScroll1 
               Height          =   255
               LargeChange     =   10
               Left            =   120
               Max             =   100
               Min             =   8
               TabIndex        =   91
               Top             =   240
               Value           =   8
               Width           =   1215
            End
            Begin VB.Frame Frame6 
               Height          =   1095
               Left            =   1440
               TabIndex        =   88
               Top             =   120
               Width           =   2055
               Begin VB.ComboBox Combo1 
                  Height          =   765
                  ItemData        =   "wordfrm(SSw2017).frx":15B19
                  Left            =   120
                  List            =   "wordfrm(SSw2017).frx":15B4D
                  Style           =   1  'Simple Combo
                  TabIndex        =   90
                  Text            =   "8"
                  ToolTipText     =   "Font size"
                  Top             =   200
                  Width           =   855
               End
               Begin VB.CommandButton Command28 
                  Caption         =   "Change"
                  Height          =   735
                  Left            =   1080
                  TabIndex        =   89
                  ToolTipText     =   "Change select font size"
                  Top             =   200
                  Width           =   855
               End
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Face"
            Height          =   1335
            Left            =   3840
            TabIndex        =   85
            Top             =   120
            Width           =   4335
            Begin VB.ComboBox Combo2 
               Height          =   960
               ItemData        =   "wordfrm(SSw2017).frx":15B8F
               Left            =   120
               List            =   "wordfrm(SSw2017).frx":15B91
               Sorted          =   -1  'True
               Style           =   1  'Simple Combo
               TabIndex        =   86
               ToolTipText     =   "Font face"
               Top             =   240
               Width           =   4095
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Color Pack"
            Height          =   1335
            Left            =   -74880
            TabIndex        =   84
            Top             =   120
            Width           =   2655
            Begin VB.CommandButton Command48 
               BackColor       =   &H005B00CC&
               Height          =   250
               Left            =   2280
               Style           =   1  'Graphical
               TabIndex        =   175
               ToolTipText     =   "Dark Maroon"
               Top             =   960
               Width           =   250
            End
            Begin VB.CommandButton Command40 
               BackColor       =   &H007D00FB&
               Height          =   250
               Left            =   2280
               Style           =   1  'Graphical
               TabIndex        =   174
               ToolTipText     =   "Maroon"
               Top             =   720
               Width           =   250
            End
            Begin VB.CommandButton Command38 
               BackColor       =   &H00A459FF&
               Height          =   250
               Left            =   2280
               Style           =   1  'Graphical
               TabIndex        =   173
               ToolTipText     =   "Light Maroon"
               Top             =   480
               Width           =   250
            End
            Begin VB.CommandButton Command67 
               BackColor       =   &H00C000C0&
               Height          =   250
               Left            =   2040
               Style           =   1  'Graphical
               TabIndex        =   172
               ToolTipText     =   "Dark Rose"
               Top             =   960
               Width           =   250
            End
            Begin VB.CommandButton Command57 
               BackColor       =   &H00FF00FF&
               Height          =   250
               Left            =   2040
               Style           =   1  'Graphical
               TabIndex        =   171
               ToolTipText     =   "Rose"
               Top             =   720
               Width           =   250
            End
            Begin VB.CommandButton Command47 
               BackColor       =   &H00FF80FF&
               Height          =   250
               Left            =   2040
               Style           =   1  'Graphical
               TabIndex        =   170
               ToolTipText     =   "Light Rose"
               Top             =   480
               Width           =   250
            End
            Begin VB.CommandButton Command68 
               BackColor       =   &H0098014D&
               Height          =   250
               Left            =   1800
               Style           =   1  'Graphical
               TabIndex        =   169
               ToolTipText     =   "Dark Purple"
               Top             =   960
               Width           =   250
            End
            Begin VB.CommandButton Command60 
               BackColor       =   &H00F10179&
               Height          =   250
               Left            =   1800
               Style           =   1  'Graphical
               TabIndex        =   168
               ToolTipText     =   "Purple"
               Top             =   720
               Width           =   250
            End
            Begin VB.CommandButton Command58 
               BackColor       =   &H00FEA5D1&
               Height          =   250
               Left            =   1800
               Style           =   1  'Graphical
               TabIndex        =   167
               ToolTipText     =   "Light Purple"
               Top             =   480
               Width           =   250
            End
            Begin VB.CommandButton Command30 
               BackColor       =   &H00C99DFF&
               Height          =   250
               Left            =   2280
               Style           =   1  'Graphical
               TabIndex        =   166
               ToolTipText     =   "Sun Light Maroon"
               Top             =   240
               Width           =   250
            End
            Begin VB.CommandButton Command37 
               BackColor       =   &H00FFC0FF&
               Height          =   250
               Left            =   2040
               Style           =   1  'Graphical
               TabIndex        =   165
               ToolTipText     =   "Sun Light Rose"
               Top             =   240
               Width           =   250
            End
            Begin VB.CommandButton Command50 
               BackColor       =   &H00FFD0E8&
               Height          =   250
               Left            =   1800
               Style           =   1  'Graphical
               TabIndex        =   164
               ToolTipText     =   "Sun Light Purple"
               Top             =   240
               Width           =   250
            End
            Begin VB.CommandButton Command66 
               BackColor       =   &H00C00000&
               Height          =   250
               Left            =   1560
               Style           =   1  'Graphical
               TabIndex        =   163
               ToolTipText     =   "Dark Big Blue"
               Top             =   960
               Width           =   250
            End
            Begin VB.CommandButton Command65 
               BackColor       =   &H00C0C000&
               Height          =   250
               Left            =   1320
               Style           =   1  'Graphical
               TabIndex        =   162
               ToolTipText     =   "Dark Small Blue"
               Top             =   960
               Width           =   250
            End
            Begin VB.CommandButton Command64 
               BackColor       =   &H0000C000&
               Height          =   250
               Left            =   1080
               Style           =   1  'Graphical
               TabIndex        =   161
               ToolTipText     =   "Dark Green"
               Top             =   960
               Width           =   250
            End
            Begin VB.CommandButton Command63 
               BackColor       =   &H0000C0C0&
               Height          =   250
               Left            =   840
               Style           =   1  'Graphical
               TabIndex        =   160
               ToolTipText     =   "Dark Yellow"
               Top             =   960
               Width           =   250
            End
            Begin VB.CommandButton Command62 
               BackColor       =   &H000040C0&
               Height          =   250
               Left            =   600
               Style           =   1  'Graphical
               TabIndex        =   159
               ToolTipText     =   "Dark Orange"
               Top             =   960
               Width           =   250
            End
            Begin VB.CommandButton Command56 
               BackColor       =   &H00FF0000&
               Height          =   250
               Left            =   1560
               Style           =   1  'Graphical
               TabIndex        =   158
               ToolTipText     =   "Big Blue"
               Top             =   720
               Width           =   250
            End
            Begin VB.CommandButton Command55 
               BackColor       =   &H00FFFF00&
               Height          =   250
               Left            =   1320
               Style           =   1  'Graphical
               TabIndex        =   157
               ToolTipText     =   "Small Blue"
               Top             =   720
               Width           =   250
            End
            Begin VB.CommandButton Command54 
               BackColor       =   &H0000FF00&
               Height          =   250
               Left            =   1080
               Style           =   1  'Graphical
               TabIndex        =   156
               ToolTipText     =   "Green"
               Top             =   720
               Width           =   250
            End
            Begin VB.CommandButton Command53 
               BackColor       =   &H0000FFFF&
               Height          =   250
               Left            =   840
               Style           =   1  'Graphical
               TabIndex        =   155
               ToolTipText     =   "Yellow"
               Top             =   720
               Width           =   250
            End
            Begin VB.CommandButton Command52 
               BackColor       =   &H000080FF&
               Height          =   250
               Left            =   600
               Style           =   1  'Graphical
               TabIndex        =   154
               ToolTipText     =   "Orange"
               Top             =   720
               Width           =   250
            End
            Begin VB.CommandButton Command46 
               BackColor       =   &H00FF8080&
               Height          =   250
               Left            =   1560
               Style           =   1  'Graphical
               TabIndex        =   153
               ToolTipText     =   "Light Big Blue"
               Top             =   480
               Width           =   250
            End
            Begin VB.CommandButton Command45 
               BackColor       =   &H00FFFF80&
               Height          =   250
               Left            =   1320
               Style           =   1  'Graphical
               TabIndex        =   152
               ToolTipText     =   "Light Small Blue"
               Top             =   480
               Width           =   250
            End
            Begin VB.CommandButton Command44 
               BackColor       =   &H0080FF80&
               Height          =   250
               Left            =   1080
               Style           =   1  'Graphical
               TabIndex        =   151
               ToolTipText     =   "Light Green"
               Top             =   480
               Width           =   250
            End
            Begin VB.CommandButton Command43 
               BackColor       =   &H0080FFFF&
               Height          =   250
               Left            =   840
               Style           =   1  'Graphical
               TabIndex        =   150
               ToolTipText     =   "Light Yellow"
               Top             =   480
               Width           =   250
            End
            Begin VB.CommandButton Command42 
               BackColor       =   &H0080C0FF&
               Height          =   250
               Left            =   600
               Style           =   1  'Graphical
               TabIndex        =   149
               ToolTipText     =   "Light Orange"
               Top             =   480
               Width           =   250
            End
            Begin VB.CommandButton Command36 
               BackColor       =   &H00FFC0C0&
               Height          =   250
               Left            =   1560
               Style           =   1  'Graphical
               TabIndex        =   148
               ToolTipText     =   "Sun Light Big Blue"
               Top             =   240
               Width           =   250
            End
            Begin VB.CommandButton Command35 
               BackColor       =   &H00FFFFC0&
               Height          =   250
               Left            =   1320
               Style           =   1  'Graphical
               TabIndex        =   147
               ToolTipText     =   "Sun Light Small Blue"
               Top             =   240
               Width           =   250
            End
            Begin VB.CommandButton Command34 
               BackColor       =   &H00C0FFC0&
               Height          =   250
               Left            =   1080
               Style           =   1  'Graphical
               TabIndex        =   146
               ToolTipText     =   "Sun Light Green"
               Top             =   240
               Width           =   250
            End
            Begin VB.CommandButton Command33 
               BackColor       =   &H00C0FFFF&
               Height          =   250
               Left            =   840
               Style           =   1  'Graphical
               TabIndex        =   145
               ToolTipText     =   "Sun Light Yellow"
               Top             =   240
               Width           =   250
            End
            Begin VB.CommandButton Command32 
               BackColor       =   &H00C0E0FF&
               Height          =   250
               Left            =   600
               Style           =   1  'Graphical
               TabIndex        =   144
               ToolTipText     =   "Sun Light Orange"
               Top             =   240
               Width           =   250
            End
            Begin VB.CommandButton Command61 
               BackColor       =   &H000000C0&
               Height          =   250
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   143
               ToolTipText     =   "Dark Red"
               Top             =   960
               Width           =   250
            End
            Begin VB.CommandButton Command51 
               BackColor       =   &H000000FF&
               Height          =   250
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   142
               ToolTipText     =   "Red"
               Top             =   720
               Width           =   250
            End
            Begin VB.CommandButton Command41 
               BackColor       =   &H008080FF&
               Height          =   250
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   141
               ToolTipText     =   "Light Red"
               Top             =   480
               Width           =   250
            End
            Begin VB.CommandButton Command31 
               BackColor       =   &H00C0C0FF&
               Height          =   250
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   140
               ToolTipText     =   "Sun Light Red"
               Top             =   240
               Width           =   250
            End
            Begin VB.CommandButton Command59 
               BackColor       =   &H00000000&
               Height          =   250
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   139
               ToolTipText     =   "Black"
               Top             =   960
               Width           =   250
            End
            Begin VB.CommandButton Command49 
               BackColor       =   &H00808080&
               Height          =   255
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   138
               ToolTipText     =   "Ash"
               Top             =   720
               Width           =   250
            End
            Begin VB.CommandButton Command39 
               BackColor       =   &H00C0C0C0&
               Height          =   250
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   137
               ToolTipText     =   "Light Ash"
               Top             =   480
               Width           =   250
            End
            Begin VB.CommandButton Command29 
               BackColor       =   &H00FFFFFF&
               Height          =   250
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   136
               ToolTipText     =   "White"
               Top             =   240
               Width           =   250
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "RGB Colors"
            Height          =   1335
            Left            =   -72120
            TabIndex        =   76
            Top             =   120
            Width           =   3975
            Begin VB.HScrollBar r 
               Height          =   255
               LargeChange     =   20
               Left            =   120
               Max             =   255
               TabIndex        =   83
               Top             =   240
               Value           =   1
               Width           =   2055
            End
            Begin VB.HScrollBar g 
               Height          =   255
               LargeChange     =   20
               Left            =   120
               Max             =   255
               TabIndex        =   82
               Top             =   600
               Value           =   1
               Width           =   2055
            End
            Begin VB.HScrollBar b 
               Height          =   255
               LargeChange     =   20
               Left            =   120
               Max             =   255
               TabIndex        =   81
               Top             =   960
               Value           =   1
               Width           =   2055
            End
            Begin VB.TextBox rt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2280
               MaxLength       =   3
               TabIndex        =   80
               Text            =   "0"
               ToolTipText     =   "Red value"
               Top             =   240
               Width           =   495
            End
            Begin VB.TextBox gt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2280
               MaxLength       =   3
               TabIndex        =   79
               Text            =   "0"
               ToolTipText     =   "Green value"
               Top             =   600
               Width           =   495
            End
            Begin VB.TextBox bt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2280
               MaxLength       =   3
               TabIndex        =   78
               Text            =   "0"
               ToolTipText     =   "Blue value"
               Top             =   960
               Width           =   495
            End
            Begin VB.CommandButton Command69 
               Caption         =   "S e t"
               Height          =   1000
               Left            =   3600
               TabIndex        =   77
               ToolTipText     =   "Set color to select word"
               Top             =   240
               Width           =   255
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00000000&
               BackStyle       =   1  'Opaque
               Height          =   1005
               Left            =   2880
               Top             =   240
               Width           =   615
            End
         End
      End
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Tenth version"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   179
      Top             =   8640
      Width           =   10260
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
      Top             =   8400
      Width           =   10335
   End
   Begin VB.Image Image1 
      Height          =   5655
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   10095
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   2760
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   5040
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   7320
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   3720
      Picture         =   "wordfrm(SSw2017).frx":15B93
      Top             =   3240
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label17 
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
      Left            =   4200
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Shape Ball 
      BackColor       =   &H80000007&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu m 
      Caption         =   "Menu"
      Begin VB.Menu new 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu open 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu save 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu saveas 
         Caption         =   "Save As"
      End
      Begin VB.Menu print 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
         Shortcut        =   ^{F12}
      End
      Begin VB.Menu emergencyexit 
         Caption         =   "Emergency Exit"
         Shortcut        =   +^{F12}
      End
   End
   Begin VB.Menu e 
      Caption         =   "Edit"
      Begin VB.Menu cut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu copy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu paste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu duplicateall 
         Caption         =   "Duplicate All"
         Shortcut        =   ^D
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu del 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu delete 
         Caption         =   "Delete All"
         Shortcut        =   +{DEL}
      End
   End
   Begin VB.Menu f 
      Caption         =   "Font"
      Begin VB.Menu style 
         Caption         =   "Style"
         Begin VB.Menu size 
            Caption         =   "Size"
            Shortcut        =   ^M
         End
         Begin VB.Menu face 
            Caption         =   "Face"
            Shortcut        =   ^Q
         End
      End
      Begin VB.Menu effects 
         Caption         =   "Effects"
         Begin VB.Menu bold 
            Caption         =   "Bold"
            Shortcut        =   ^B
         End
         Begin VB.Menu italic 
            Caption         =   "Italic"
            Shortcut        =   ^I
         End
         Begin VB.Menu underline 
            Caption         =   "Underline"
            Shortcut        =   ^U
         End
         Begin VB.Menu strikethrought 
            Caption         =   "Strike Throught"
            Shortcut        =   ^W
         End
      End
      Begin VB.Menu alignment 
         Caption         =   "Alignment"
         Begin VB.Menu aleft 
            Caption         =   "Left"
            Shortcut        =   ^J
         End
         Begin VB.Menu acenter 
            Caption         =   "Center"
            Shortcut        =   ^E
         End
         Begin VB.Menu aright 
            Caption         =   "Right"
            Shortcut        =   ^R
         End
      End
      Begin VB.Menu caseandspace 
         Caption         =   "Case and Space"
         Begin VB.Menu case 
            Caption         =   "Case"
            Begin VB.Menu caseu 
               Caption         =   "Upper Case"
               Shortcut        =   ^A
            End
            Begin VB.Menu casel 
               Caption         =   "Lower Case"
               Shortcut        =   ^F
            End
            Begin VB.Menu caps1 
               Caption         =   "First letter capital"
               Shortcut        =   ^G
            End
            Begin VB.Menu simp1 
               Caption         =   "First letter simpal"
               Shortcut        =   ^H
            End
         End
         Begin VB.Menu sp 
            Caption         =   "Space"
            Begin VB.Menu addspace 
               Caption         =   "Add one/more spaces"
               Shortcut        =   ^L
            End
            Begin VB.Menu delspace 
               Caption         =   "Delete space/s"
               Shortcut        =   ^K
            End
         End
      End
   End
   Begin VB.Menu tools 
      Caption         =   "Tools"
      Begin VB.Menu calculator 
         Caption         =   "Calculator"
         Shortcut        =   {F5}
      End
      Begin VB.Menu calendar 
         Caption         =   "Calendar"
         Shortcut        =   {F6}
      End
      Begin VB.Menu musicplayer 
         Caption         =   "Music Player"
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu about 
      Caption         =   "About"
      Begin VB.Menu helpto 
         Caption         =   "About SRISOFTwrite 2017"
         Shortcut        =   {F1}
      End
      Begin VB.Menu abcel 
         Caption         =   "About SRI Calendar"
         Shortcut        =   {F2}
      End
      Begin VB.Menu abcal 
         Caption         =   "About SRI Caculator"
         Shortcut        =   {F3}
      End
      Begin VB.Menu absp 
         Caption         =   "About SRI Song Player"
         Shortcut        =   {F4}
      End
      Begin VB.Menu abm 
         Caption         =   "About maker"
         Shortcut        =   {F9}
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XInc, YInc, XPos, YPos


Private Sub abcal_Click()
Form9.Show (vbModal)
End Sub

Private Sub abcel_Click()
Form7.Show (vbModal)
End Sub

Private Sub abm_Click()
make.Show (vbModal)
End Sub

Private Sub absp_Click()
Form10.Show (vbModal)
End Sub

Private Sub acenter_Click()
'mnu center
Command17_Click
End Sub

Private Sub addspace_Click()
'mnu add spaces
Dim Ch, CIndex, NewText
Dim s As Integer
s = Val(InputBox("Enter the number of spaces", "Add space/s"))
NewText = ""
For CIndex = 1 To Len(RichTextBox1.SelText)
Ch = Mid(RichTextBox1.SelText, CIndex, 1)
NewText = NewText + Ch
    If (Asc(Ch) = 32) Then
    NewText = NewText + Space(s)
    End If
Next
RichTextBox1.SelText = NewText
End Sub

Private Sub aleft_Click()
'mnu left
Command16_Click
End Sub

Private Sub align_Timer()
'align mnu checked
If RichTextBox1.SelAlignment = 0 Then
aleft.Checked = True
acenter.Checked = False
aright.Checked = False
ElseIf RichTextBox1.SelAlignment = 2 Then
aleft.Checked = False
acenter.Checked = True
aright.Checked = False
ElseIf RichTextBox1.SelAlignment = 1 Then
aleft.Checked = False
acenter.Checked = False
aright.Checked = True
End If
End Sub

Private Sub aright_Click()
'mnu right
Command18_Click
End Sub

Private Sub b_Change()
Shape1.BackColor = RGB(r.Value, g.Value, b.Value)
bt.Text = b.Value
End Sub

Private Sub bdt_Change()
bdt.Text = Val(bdt.Text)
If bdt.Text = "" Then
bdt.Text = "0"
bd.Value = bdt.Text
ElseIf bdt.Text > 255 Then
bdt.Text = "255"
bd.Value = bdt.Text
ElseIf bdt.Text < 0 Then
bdt.Text = "0"
bd.Value = bdt.Text
ElseIf bdt.Text >= 0 And 255 Then
bd.Value = bdt.Text
Else
bdt.Text = "0"
End If
End Sub

Private Sub bg_Timer()
Form2.BackColor = QBColor(Rnd * 15)
End Sub

Private Sub bd_Change()
Shape2.BackColor = RGB(rd.Value, gd.Value, bd.Value)
bdt.Text = bd.Value
End Sub

Private Sub bld_Change()

End Sub

Private Sub blt_Change()

End Sub

Private Sub bground_Click()
On Error GoTo err:
CommonDialog1.ShowColor
Form2.BackColor = CommonDialog1.Color
err:
Exit Sub
End Sub

Private Sub bold_Click()
'mnu bold
Command12_Click
End Sub

Private Sub bt_Change()
bt.Text = Val(bt.Text)
If bt.Text = "" Then
bt.Text = "0"
b.Value = bt.Text
ElseIf bt.Text > 255 Then
bt.Text = "255"
b.Value = bt.Text
ElseIf bt.Text < 0 Then
bt.Text = "0"
b.Value = bt.Text
ElseIf bt.Text > 0 And 255 Then
b.Value = bt.Text
Else
bt.Text = "0"
End If
End Sub

Private Sub bx_Timer()
 Line (GetRndV(10000), GetRndV(10000))- _
 Step(GetRndV(1000), GetRndV(1000)), _
 RGB(GetRndV(256), GetRndV(256), _
 GetRndV(256)), BF
End Sub

Private Sub calculator_Click()
'mnu cal
Form4.Show
End Sub

Private Sub calendar_Click()
'mnu cel
Form5.Show
End Sub

Private Sub caps1_Click()
'mnu 1caps
Command22_Click
End Sub

Private Sub cc_Timer()
'cut mnu enabled
If RichTextBox1.SelText = "" Then
cut.Enabled = False
Copy.Enabled = False
Else
cut.Enabled = True
Copy.Enabled = True
End If
End Sub

Private Sub Check10_Click()
If Check10.Value = 1 Then
Combo3.Enabled = False
Combo4.Enabled = False
Combo5.Enabled = False
Combo6.Enabled = False
Else
Combo3.Enabled = True
Combo4.Enabled = True
Combo5.Enabled = True
Combo6.Enabled = True
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Combo9.Enabled = False
Else
Combo9.Enabled = True
End If
End Sub

Private Sub cm_Timer()
Static s As Integer
        If Combo7.Text = "Seconds" Then cm.Interval = Text1.Text * 1000
        If Combo7.Text = "Minutes" Then cm.Interval = Text1.Text * 1000 * 60
s = s + 1
If s = 1 Then
    If Combo3.Text = "Blue" Then Me.BackColor = vbBlue
    If Combo3.Text = "Black" Then Me.BackColor = vbBlack
    If Combo3.Text = "Green" Then Me.BackColor = vbGreen
    If Combo3.Text = "Red" Then Me.BackColor = vbRed
    If Combo3.Text = "White" Then Me.BackColor = vbWhite
    If Combo3.Text = "Yellow" Then Me.BackColor = vbYellow
ElseIf s = 2 Then
    If Combo4.Text = "Blue" Then Me.BackColor = vbBlue
    If Combo4.Text = "Black" Then Me.BackColor = vbBlack
    If Combo4.Text = "Green" Then Me.BackColor = vbGreen
    If Combo4.Text = "Red" Then Me.BackColor = vbRed
    If Combo4.Text = "White" Then Me.BackColor = vbWhite
    If Combo4.Text = "Yellow" Then Me.BackColor = vbYellow
ElseIf s = 3 Then
    If Combo5.Text = "Blue" Then Me.BackColor = vbBlue
    If Combo5.Text = "Black" Then Me.BackColor = vbBlack
    If Combo5.Text = "Green" Then Me.BackColor = vbGreen
    If Combo5.Text = "Red" Then Me.BackColor = vbRed
    If Combo5.Text = "White" Then Me.BackColor = vbWhite
    If Combo5.Text = "Yellow" Then Me.BackColor = vbYellow
ElseIf s = 4 Then
    If Combo6.Text = "Blue" Then Me.BackColor = vbBlue
    If Combo6.Text = "Black" Then Me.BackColor = vbBlack
    If Combo6.Text = "Green" Then Me.BackColor = vbGreen
    If Combo6.Text = "Red" Then Me.BackColor = vbRed
    If Combo6.Text = "White" Then Me.BackColor = vbWhite
    If Combo6.Text = "Yellow" Then Me.BackColor = vbYellow
If s = 4 Then s = 0
End If
End Sub

Private Sub cmr_Timer()
Form2.BackColor = QBColor(Rnd * 15)
End Sub

Private Sub colorpal_Click()
On Error GoTo err:
'mnu font color
CommonDialog1.ShowColor
RichTextBox1.SelColor = CommonDialog1.Color
err:
Exit Sub
End Sub

Private Sub Combo2_Click()
'font
RichTextBox1.SelFontName = Combo2.Text
End Sub

Private Sub Command1_Click()
'new
Dim strmsg As String
strmsg = MsgBox("Do you want to save this document?", vbQuestion + vbYesNoCancel, "New")
If strmsg = vbYes Then
    Call newsave
Else
If strmsg = vbNo Then
    RichTextBox1.Text = ""
    RichTextBox1.BackColor = vbWhite
    RichTextBox1.SelStrikeThru = False
    RichTextBox1.SelItalic = False
    RichTextBox1.SelBold = False
    RichTextBox1.SelUnderline = False
    RichTextBox1.SelColor = vbBlack
    RichTextBox1.SelFontSize = 8
Form2.Caption = " SRISOFTwrite 2017"
Else
End If
End If
End Sub

Public Sub newsave()
'new save
Dim sFile As String
If Left$(Form2.Caption, 18) = " SRISOFTwrite 2017" Then
    If Form2 Is Nothing Then Exit Sub
    CommonDialog1.DialogTitle = "Save As"
    CommonDialog1.CancelError = False
    CommonDialog1.Filter = "All Files (*.*)|*.*"
    CommonDialog1.ShowSave
    If Len(CommonDialog1.FileName) = 0 Then
    Exit Sub
    End If
sFile = CommonDialog1.FileName
Form2.RichTextBox1.SaveFile sFile
Else
sFile = Form2.Caption
Form2.RichTextBox1.SaveFile sFile
End If
'''''''''''''''''''''''''''''''''''''''''''''''
RichTextBox1.Text = ""
RichTextBox1.BackColor = vbWhite
RichTextBox1.SelStrikeThru = False
RichTextBox1.SelItalic = False
RichTextBox1.SelBold = False
RichTextBox1.SelUnderline = False
RichTextBox1.SelColor = vbBlack
RichTextBox1.SelFontSize = 8
Form2.Caption = " SRISOFTwrite 2017"
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can get new document here. "
End Sub

Private Sub Command10_Click()
'delete
RichTextBox1.SelText = ""
End Sub

Private Sub Command10_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can delete select part here. "
End Sub

Private Sub Command11_Click()
'emergency exit
emergencyexit_Click
End Sub

Private Sub Command11_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can emergency exit your files here. "
End Sub

Private Sub Command12_Click()
'bold
If RichTextBox1.SelBold = False Then
RichTextBox1.SelBold = True
Bold.Checked = True
Else
RichTextBox1.SelBold = False
Bold.Checked = False
End If
End Sub

Private Sub Command12_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can bold select part here. "
End Sub

Private Sub Command13_Click()
'italic
If RichTextBox1.SelItalic = False Then
RichTextBox1.SelItalic = True
Italic.Checked = True
Else
RichTextBox1.SelItalic = False
Italic.Checked = False
End If
End Sub

Private Sub Command13_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can italic select part here. "
End Sub

Private Sub Command14_Click()
'underline
If RichTextBox1.SelUnderline = False Then
RichTextBox1.SelUnderline = True
Underline.Checked = True
Else
RichTextBox1.SelUnderline = False
Underline.Checked = False
End If
End Sub

Private Sub Command14_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can underline select part here. "
End Sub

Private Sub Command15_Click()
'strike through
If RichTextBox1.SelStrikeThru = False Then
RichTextBox1.SelStrikeThru = True
strikethrought.Checked = True
Else
RichTextBox1.SelStrikeThru = False
strikethrought.Checked = False
End If
End Sub

Private Sub Command15_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can strike through select part here. "
End Sub

Private Sub Command16_Click()
'l align
RichTextBox1.SelAlignment = 0
End Sub

Private Sub Command16_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can change to left align select part here. "
End Sub

Private Sub Command17_Click()
'c align
RichTextBox1.SelAlignment = 2
End Sub

Private Sub Command17_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can change to center align select part here. "
End Sub

Private Sub Command18_Click()
'r Align
RichTextBox1.SelAlignment = 1
End Sub

Private Sub Command18_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can change to right align select part here. "
End Sub

Private Sub Command19_Click()
'duplicator
RichTextBox1.SelText = RichTextBox1.Text
End Sub

Private Sub Command19_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can duplicate your document here. "
End Sub

Private Sub Command2_Click()
'open
Dim sFile As String
CommonDialog1.DialogTitle = "Open"
CommonDialog1.CancelError = False
CommonDialog1.Filter = "All Files (*.*)|*.*"
CommonDialog1.ShowOpen
    If Len(CommonDialog1.FileName) = 0 Then
    Exit Sub
    End If
sFile = CommonDialog1.FileName
RichTextBox1.LoadFile sFile
Form2.Caption = sFile
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can open document here. "
End Sub

Private Sub Command20_Click()
'caps
RichTextBox1.SelText = UCase(RichTextBox1.SelText)
End Sub

Private Sub Command20_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can change to Upper Case select part here. "
End Sub

Private Sub Command21_Click()
'simp
RichTextBox1.SelText = LCase(RichTextBox1.SelText)
End Sub

Private Sub Command21_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can change to Lower Case select part here. "
End Sub

Private Sub Command22_Click()
'1 st caps
Dim SFound As Boolean
Dim Ch, NText, CIndex
NText = ""
SFound = True
For CIndex = 1 To Len(RichTextBox1.SelText)
Ch = Mid(RichTextBox1.SelText, CIndex, 1)
   If SFound And (Asc(Ch) >= 97) And (Asc(Ch) <= 122) Then
   Ch = Chr(Asc(Ch) - 32)
    End If
SFound = (Asc(Ch) = 32)
NText = NText + Ch
Next
RichTextBox1.SelText = NText
End Sub

Private Sub Command22_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can change words' first letter to Upper case select part here. "
End Sub

Private Sub Command23_Click()
'1st simp
Dim SFound As Boolean
Dim Ch, NText, CIndex
NText = ""
SFound = True
For CIndex = 1 To Len(RichTextBox1.SelText)
Ch = Mid(RichTextBox1.SelText, CIndex, 1)
   If SFound And (Asc(Ch) >= 65) And (Asc(Ch) <= 90) Then
   Ch = Chr(Asc(Ch) + 32)
    End If
SFound = (Asc(Ch) = 32)
NText = NText + Ch
Next
RichTextBox1.SelText = NText
End Sub

Private Sub Command23_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can change words' first letter to Lower case select part here. "
End Sub

Private Sub Command24_Click()
'add spaces
Dim Ch, CIndex, NewText
    If NoSp.Text = "" Then
    NoSp.Text = "1"
    NewText = ""
    For CIndex = 1 To Len(RichTextBox1.SelText)
    Ch = Mid(RichTextBox1.SelText, CIndex, 1)
    NewText = NewText + Ch
        If (Asc(Ch) = 32) Then
        NewText = NewText + Space(1)
        End If
    Next
    RichTextBox1.SelText = NewText
    Else
    NewText = ""
    For CIndex = 1 To Len(RichTextBox1.SelText)
    Ch = Mid(RichTextBox1.SelText, CIndex, 1)
    NewText = NewText + Ch
        If (Asc(Ch) = 32) Then
        NewText = NewText + Space(NoSp)
        End If
    Next
    RichTextBox1.SelText = NewText
    End If
End Sub

Private Sub Command25_Click()
'delete spaces
Dim Ch, CIndex, NewText
NewText = ""
For CIndex = 1 To Len(RichTextBox1.SelText)
Ch = Mid(RichTextBox1.SelText, CIndex, 1)
    If (Asc(Ch) <> 32) Then
    NewText = NewText + Ch
    End If
Next
RichTextBox1.SelText = NewText
End Sub

Private Sub Command26_Click()
On Error GoTo err
'increase font size
RichTextBox1.SelFontSize = RichTextBox1.SelFontSize + 1
err:
End Sub

Private Sub Command27_Click()
On Error GoTo err
'decrease font size
RichTextBox1.SelFontSize = RichTextBox1.SelFontSize - 1
err:
End Sub

Private Sub Command28_Click()
'size change
If Combo1.Text < 8 Then
Combo1.Text = 8
RichTextBox1.SelFontSize = Combo1.Text
ElseIf Combo1.Text > 100 Then
Combo1.Text = 100
RichTextBox1.SelFontSize = Combo1.Text
Else
RichTextBox1.SelFontSize = Combo1.Text
End If
End Sub

Private Sub Command29_Click()
RichTextBox1.SelColor = vbWhite
End Sub

Private Sub Command3_Click()
'save
Dim sFile As String
If Left$(Form2.Caption, 18) = " SRISOFTwrite 2017" Then
    If Form2 Is Nothing Then Exit Sub
    CommonDialog1.DialogTitle = "Save As"
    CommonDialog1.CancelError = False
    CommonDialog1.Filter = "All Files (*.*)|*.*"
    CommonDialog1.ShowSave
    If Len(CommonDialog1.FileName) = 0 Then
    Exit Sub
    End If
    sFile = CommonDialog1.FileName
    Form2.Caption = sFile
    Form2.RichTextBox1.SaveFile sFile
    Else
    sFile = Form2.Caption
    Form2.RichTextBox1.SaveFile sFile
    End If
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can save current document here. "
End Sub

Private Sub Command30_Click()
RichTextBox1.SelColor = RGB(255, 157, 201)
End Sub

Private Sub Command31_Click()
RichTextBox1.SelColor = RGB(255, 192, 192)
End Sub

Private Sub Command32_Click()
RichTextBox1.SelColor = RGB(255, 224, 192)
End Sub

Private Sub Command33_Click()
RichTextBox1.SelColor = RGB(255, 255, 192)
End Sub

Private Sub Command34_Click()
RichTextBox1.SelColor = RGB(192, 255, 192)
End Sub

Private Sub Command35_Click()
RichTextBox1.SelColor = RGB(192, 255, 255)
End Sub

Private Sub Command36_Click()
RichTextBox1.SelColor = RGB(192, 192, 255)
End Sub

Private Sub Command37_Click()
RichTextBox1.SelColor = RGB(255, 192, 255)
End Sub

Private Sub Command38_Click()
RichTextBox1.SelColor = RGB(255, 89, 164)
End Sub

Private Sub Command39_Click()
RichTextBox1.SelColor = RGB(192, 192, 192)
End Sub

Private Sub Command4_Click()
'save as
Dim sFile As String
If Form2 Is Nothing Then Exit Sub
    CommonDialog1.DialogTitle = "Save As"
    CommonDialog1.CancelError = False
    CommonDialog1.Filter = "All Files (*.*)|*.*"
    CommonDialog1.ShowSave
    If Len(CommonDialog1.FileName) = 0 Then
    Exit Sub
    End If
    sFile = CommonDialog1.FileName
    Form2.Caption = sFile
    Form2.RichTextBox1.SaveFile sFile
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can save as current document here. "
End Sub

Private Sub Command40_Click()
RichTextBox1.SelColor = RGB(251, 0, 125)
End Sub

Private Sub Command41_Click()
RichTextBox1.SelColor = RGB(255, 128, 128)
End Sub

Private Sub Command42_Click()
RichTextBox1.SelColor = RGB(255, 192, 128)
End Sub

Private Sub Command43_Click()
RichTextBox1.SelColor = RGB(255, 255, 128)
End Sub

Private Sub Command44_Click()
RichTextBox1.SelColor = RGB(128, 255, 128)
End Sub

Private Sub Command45_Click()
RichTextBox1.SelColor = RGB(128, 255, 255)
End Sub

Private Sub Command46_Click()
RichTextBox1.SelColor = RGB(128, 128, 255)
End Sub

Private Sub Command47_Click()
RichTextBox1.SelColor = RGB(255, 128, 255)
End Sub

Private Sub Command48_Click()
RichTextBox1.SelColor = RGB(204, 0, 91)
End Sub

Private Sub Command49_Click()
RichTextBox1.SelColor = RGB(128, 128, 128)
End Sub

Private Sub Command5_Click()
'exit
Dim strmsg As String
Dim sFile As String
    If RichTextBox1.Text = "" Then
    End
    Else
    strmsg = MsgBox("Do you want to save this document?", vbQuestion + vbYesNoCancel, "Exit")
    If strmsg = vbYes Then
    Call saveasnew
    Else
    If strmsg = vbNo Then
    End
    Else
    End If
    End If
    End If
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can close from current document here. "
End Sub

Private Sub Command50_Click()
RichTextBox1.SelColor = RGB(232, 208, 255)
End Sub

Private Sub Command51_Click()
RichTextBox1.SelColor = RGB(255, 0, 0)
End Sub

Private Sub Command52_Click()
RichTextBox1.SelColor = RGB(255, 128, 0)
End Sub

Private Sub Command53_Click()
RichTextBox1.SelColor = RGB(255, 255, 0)
End Sub

Private Sub Command54_Click()
RichTextBox1.SelColor = RGB(0, 255, 0)
End Sub

Private Sub Command55_Click()
RichTextBox1.SelColor = RGB(0, 255, 255)
End Sub

Private Sub Command56_Click()
RichTextBox1.SelColor = RGB(0, 0, 255)
End Sub

Private Sub Command57_Click()
RichTextBox1.SelColor = RGB(255, 0, 255)
End Sub

Private Sub Command58_Click()
RichTextBox1.SelColor = RGB(209, 165, 254)
End Sub

Private Sub Command59_Click()
RichTextBox1.SelColor = vbBlack
End Sub

Private Sub Command6_Click()
'print
On Error Resume Next
    If Form2 Is Nothing Then Exit Sub
    CommonDialog1.DialogTitle = "Print"
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlPDReturnDC + cdlPDNoPageNums
    If Form2CommonDialog1.RichTextBox1.SelLength = 0 Then
    CommonDialog1.Flags = CommonDialog1.Flags + cdlPDAllPages
    Else
    CommonDialog1.Flags = CommonDialog1.Flags + cdlPDSelection
    End If
    CommonDialog1.ShowPrinter
    If err <> MSComDlgCommonDialog1.cdlCancel Then
    Form2CommonDialog1.RichTextBox1.SelPrint CommonDialog1.hDC
    End If
End Sub

Private Sub saveasnew()
'exit save
Dim sFile As String
    If Left$(Form2.Caption, 20) = " SRISOFTwrite 2017" Then
    If Form2 Is Nothing Then Exit Sub
    CommonDialog1.DialogTitle = "Save As"
    CommonDialog1.CancelError = False
    CommonDialog1.Filter = "All Files (*.*)|*.*"
    CommonDialog1.ShowSave
    If Len(CommonDialog1.FileName) = 0 Then
    Exit Sub
    End If
    sFile = CommonDialog1.FileName
    Form2.Caption = sFile
    Form2.RichTextBox1.SaveFile sFile
    End
    Else
    sFile = Form2.Caption
    Form2.RichTextBox1.SaveFile sFile
    End
    End If
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can print current document here. "
End Sub

Private Sub Command60_Click()
RichTextBox1.SelColor = RGB(121, 1, 241)
End Sub

Private Sub Command61_Click()
RichTextBox1.SelColor = RGB(192, 0, 0)
End Sub

Private Sub Command62_Click()
RichTextBox1.SelColor = RGB(192, 64, 0)
End Sub

Private Sub Command63_Click()
RichTextBox1.SelColor = RGB(192, 192, 0)
End Sub

Private Sub Command64_Click()
RichTextBox1.SelColor = RGB(0, 192, 0)
End Sub

Private Sub Command65_Click()
RichTextBox1.SelColor = RGB(0, 192, 192)
End Sub

Private Sub Command66_Click()
RichTextBox1.SelColor = RGB(0, 0, 192)
End Sub

Private Sub Command67_Click()
RichTextBox1.SelColor = RGB(192, 0, 192)
End Sub

Private Sub Command68_Click()
RichTextBox1.SelColor = RGB(77, 1, 152)
End Sub

Private Sub Command69_Click()
RichTextBox1.SelColor = Shape1.BackColor
End Sub

Private Sub Command7_Click()
'cut
Clipboard.SetText (RichTextBox1.SelRTF)
Form2.RichTextBox1.SelText = vbNullString
End Sub

Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can cut select part to clip board here. "
End Sub

Private Sub Command70_Click()
On Error GoTo err:
'font color
CommonDialog1.ShowColor
RichTextBox1.SelColor = CommonDialog1.Color
err:
Exit Sub
End Sub

Private Sub Command70_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can get color palatte and change font color of select part here. "
End Sub

Private Sub Command71_Click()
'show calculator
Form4.Show
End Sub

Private Sub Command71_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can get calculator here. "
End Sub

Private Sub Command72_Click()
'calendar show
Form5.Show
End Sub

Private Sub Command72_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can get calendar here. "
End Sub

Private Sub Command73_Click()
'show fontbox
CommonDialog1.Flags = cdlCFEffects Or cdlCFBoth
CommonDialog1.ShowFont
RichTextBox1.SelFontSize = CommonDialog1.FontSize
RichTextBox1.SelBold = CommonDialog1.FontBold
RichTextBox1.SelItalic = CommonDialog1.FontItalic
RichTextBox1.SelUnderline = CommonDialog1.FontUnderline
RichTextBox1.SelStrikeThru = CommonDialog1.FontStrikethru
RichTextBox1.SelFontName = CommonDialog1.FontName
RichTextBox1.SelColor = CommonDialog1.Color
End Sub

Private Sub Command73_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can get font properties here. "
End Sub

Private Sub Command74_Click()
'show song player
Form6.Show
End Sub

Private Sub Command74_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can get song player here. "
End Sub

Private Sub Command75_Click()
Dim strmsg As String
If List1.Text = "None" Then
tt.Enabled = False
Else
    If Text2.Text < 10 Then
    strmsg = MsgBox("Wait time must between 10 sec. to 99 minutes.", vbCritical + vbOKOnly, "Error")
    tt.Enabled = False
    Else
    tt.Enabled = True
    screensaver.Text = "0"
    Text6.Text = Text2.Text
    scres.Text = List1.Text
    Second.Text = Combo8.Text
    End If
End If
End Sub

Private Sub Command75_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can set screen saver here. "
End Sub

Private Sub Command76_Click()
If Option5.Value = True Then
Me.BackColor = Shape2.BackColor
Text8.Text = rdt.Text
Text9.Text = gdt.Text
Text10.Text = bdt.Text
ElseIf Option6.Value = True Then
RichTextBox1.BackColor = Shape2.BackColor
End If
End Sub

Private Sub Command77_Click()
On Error GoTo err:
CommonDialog1.ShowColor
RichTextBox1.BackColor = CommonDialog1.Color
err:
Exit Sub
End Sub

Private Sub Command77_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can change back color of text box here. "
End Sub

Private Sub Command78_Click()
Form8.Show (vbModal)
End Sub

Private Sub Command78_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can get about of SRISOFTwrite 2017 here. "
End Sub

Private Sub Command79_Click()
'delete
Dim strmsg As String
strmsg = MsgBox("Are sure about all text delete?", vbQuestion + vbYesNo, "Delete all")
If strmsg = vbYes Then
RichTextBox1.Text = ""
Else
End If
End Sub

Private Sub Command79_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can delete all here. "
End Sub

Private Sub Command8_Click()
'copy
Clipboard.SetText (RichTextBox1.SelRTF)
End Sub

Private Sub Command8_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can copy select part to clip board here. "
End Sub

Private Sub Command80_Click()
Form7.Show (vbModal)
End Sub

Private Sub Command80_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can get about of SRI Calendar(ver.2.0) here. "
End Sub

Private Sub Command81_Click()
Form9.Show (vbModal)
End Sub

Private Sub Command81_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can get about of SRI Calculator(ver.2.0) here. "
End Sub

Private Sub Command82_Click()
Form10.Show (vbModal)
End Sub

Private Sub Command82_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can get about of SRI Song Player(ver.1.0) here. "
End Sub

Private Sub Command83_Click()
RichTextBox1.MousePointer = rtfCustom
RichTextBox1.MouseIcon = LoadPicture(Dir1.Path + "\" + File1.FileName)
End Sub

Private Sub Command83_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can set icon of text box here. "
End Sub

Private Sub Command84_Click()
RichTextBox1.MousePointer = rtfDefault
End Sub

Private Sub Command84_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can remove icon of text box here. "
End Sub

Private Sub Command85_Click()
RichTextBox1.MousePointer = rtfCustom
If Option3.Value = True Then RichTextBox1.MouseIcon = LoadPicture(App.Path & "\Files\Icons\SS\pen icon.ICO")
If Option4.Value = True Then RichTextBox1.MouseIcon = LoadPicture(App.Path & "\Files\Icons\SS\pencil icon.ICO")
If Option2.Value = True Then RichTextBox1.MouseIcon = LoadPicture(App.Path & "\Files\Icons\SS\eraser icon.ICO")
If Option1.Value = True Then RichTextBox1.MouseIcon = LoadPicture(App.Path & "\Files\Icons\SS\correction icon.ICO")
End Sub

Private Sub Command86_Click()
make.Show (vbModal)
End Sub

Private Sub Command9_Click()
'paste
Form2.RichTextBox1.SelRTF = Clipboard.GetText
End Sub

Private Sub Command9_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can paste select part from clip board here. "
End Sub

Private Sub copy_Click()
'mnu copy
Command8_Click
End Sub

Private Sub cut_Click()
'mnu cut
Command7_Click
End Sub

Private Sub del_Click()
Command10_Click
End Sub

Private Sub delete_Click()
Command79_Click
End Sub

Private Sub delspace_Click()
'mnu delete space
Command25_Click
End Sub

Private Sub Dir1_Change()
On Error GoTo err:
File1.FileName = Dir1.Path
err:
Exit Sub
End Sub

Private Sub Drive1_Change()
On Error GoTo err:
Dir1.Path = Drive1.Drive
err:
Exit Sub
End Sub

Private Sub duplicateall_Click()
Command19_Click
End Sub

Private Sub emergencyexit_Click()
'mnu emergency exit
Dim strmsg As String
strmsg = MsgBox("Are you sure?", vbQuestion + vbYesNo, "Emergency Exit")
    If strmsg = vbYes Then
    End
    Else
    End If
End Sub

Private Sub exit_Click()
'mnu exit
Command5_Click
End Sub

Private Sub face_Click()
'mnu font
Form3.Show (vbModal)
End Sub

Private Sub firstcap_Click()

End Sub

Private Sub Form_Click()
        Cls
        SSTab1.Visible = True
        screensaver = 0
        List1.Text = scres.Text
        RichTextBox1.Visible = True
        Form2.BackColor = RGB(Text8.Text, Text9.Text, Text10.Text)
        Label12.Visible = False
        Label13.Visible = False
        Label14.Visible = False
        Image1.Visible = False
        ss.Enabled = False
        ssr.Enabled = False
        dem.Visible = False
        sec.Visible = False
        Min.Visible = False
        hrs.Visible = False
        sl.Enabled = False
        Ball.Visible = False
        pp.Enabled = False
        cm.Enabled = False
        cmr.Enabled = False
        sq.Enabled = False
        rndc.Enabled = False
        bx.Enabled = False
        ru.Enabled = False
        bg.Enabled = False
        Image2.Visible = False
        Label17.Visible = False
        sw.Enabled = False
        StatusBar1.Visible = True
        m.Visible = True
        e.Visible = True
        f.Visible = True
        tools.Visible = True
        about.Visible = True
End Sub

Private Sub Form_Load()
 'font load
Dim i As Integer
For i = 1 To Screen.FontCount
Combo2.AddItem Screen.Fonts(i)
Next i
List1.Text = "None"
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
screensaver = 0
     List1.Text = scres.Text
     Text2.Text = Text6.Text
     Combo8.Text = Second.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
Command5_Click
End Sub

Private Sub Frame10_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can change background color here. "
End Sub

Private Sub Frame11_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can change and set defalt icons of text box here. "
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can increase spaces select part here. "
End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can decrease spaces select part here. "
End Sub

Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can get size options here. "
End Sub

Private Sub Frame7_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You change font face here. "
End Sub

Private Sub Frame8_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can change font color of select part here. "
End Sub

Private Sub Frame9_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can change font color(RGB) of select part here. "
End Sub

Private Sub g_Change()
Shape1.BackColor = RGB(r.Value, g.Value, b.Value)
gt.Text = g.Value
End Sub

Private Sub gd_Change()
Shape2.BackColor = RGB(rd.Value, gd.Value, bd.Value)
gdt.Text = gd.Value
End Sub



Private Sub gdt_Change()
gdt.Text = Val(gdt.Text)
If gdt.Text = "" Then
gdt.Text = "0"
gd.Value = gdt.Text
ElseIf gdt.Text > 255 Then
gdt.Text = "255"
gd.Value = gdt.Text
ElseIf gdt.Text < 0 Then
gdt.Text = "0"
gd.Value = gdt.Text
ElseIf gdt.Text >= 0 And 255 Then
gd.Value = gdt.Text
Else
gdt.Text = "0"
End If
End Sub

Private Sub gt_Change()
gt.Text = Val(gt.Text)
If gt.Text = "" Then
gt.Text = "0"
g.Value = gt.Text
ElseIf gt.Text > 255 Then
gt.Text = "255"
g.Value = gt.Text
ElseIf gt.Text < 0 Then
gt.Text = "0"
g.Value = gt.Text
ElseIf gt.Text >= 0 And 255 Then
g.Value = gt.Text
Else
gt.Text = "0"
End If
End Sub

Private Sub helpto_Click()
Form8.Show (vbModal)
End Sub

Private Sub HScroll1_Change()
'font size
RichTextBox1.SelFontSize = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
'font size
RichTextBox1.SelFontSize = HScroll1.Value
End Sub

Private Sub ic_Timer()

End Sub

Private Sub icon_Timer()

End Sub

Private Sub Image1_Click()
        SSTab1.Visible = True
        Cls
        screensaver = 0
        List1.Text = scres.Text
        RichTextBox1.Visible = True
        Form2.BackColor = RGB(255, 192, 192)
        Label12.Visible = False
        Label13.Visible = False
        Label14.Visible = False
        Image1.Visible = False
        ss.Enabled = False
        ssr.Enabled = False
        dem.Visible = False
        sec.Visible = False
        Min.Visible = False
        hrs.Visible = False
        sl.Enabled = False
        Ball.Visible = False
        pp.Enabled = False
        cm.Enabled = False
        cmr.Enabled = False
        sq.Enabled = False
        rndc.Enabled = False
        bx.Enabled = False
        ru.Enabled = False
        bg.Enabled = False
        Image2.Visible = False
        Label17.Visible = False
        sw.Enabled = False
End Sub

Private Sub italic_Click()
'mnu italic
Command13_Click
End Sub

Private Sub casel_Click()
'mnu lcase
Command21_Click
End Sub

Private Sub lt_Change()

End Sub

Private Sub musicplayer_Click()
'mnu mus
Form6.Show
End Sub

Private Sub new_Click()
'mnu new
Command1_Click
End Sub

Private Sub pd_Timer()

End Sub

Private Sub rd_Change()
Shape2.BackColor = RGB(rd.Value, gd.Value, bd.Value)
rdt.Text = rd.Value
End Sub

Private Sub rdt_Change()
rdt.Text = Val(rdt.Text)
If rdt.Text = "" Then
rdt.Text = "0"
rd.Value = rdt.Text
ElseIf rdt.Text > 255 Then
rdt.Text = "255"
rd.Value = rdt.Text
ElseIf rdt.Text < 0 Then
rdt.Text = "0"
rd.Value = rdt.Text
ElseIf rdt.Text >= 0 And 255 Then
rd.Value = rdt.Text
Else
rdt.Text = "0"
End If
End Sub

Private Sub RichTextBox1_KeyPress(KeyAscii As Integer)
screensaver = 0
End Sub

Private Sub ru_Timer()
 Circle (GetRndV(10000), GetRndV(10000)), _
 GetRndV(1000), RGB(GetRndV(256), _
 GetRndV(256), GetRndV(256))
End Sub

Private Sub sl_Timer()
Static s As Integer
If Check4.Value = 0 Then
Image2.Visible = True
Label17.Visible = True
 XPos = Rnd * 10000
 YPos = Rnd * 6000
 Image2.Top = YPos
 Image2.Left = XPos
 Label17.Top = Image2.Top + 240
 Label17.Left = Image2.Left + 480
 If XPos > 11000 Or XPos < 6000 Then
 XInc = -XInc
 End If
 If YPos > 10000 Or YPos < 5000 Then
 YInc = -YInc
 End If
ElseIf Check4.Value = 1 Then
s = s + 1
Image2.Visible = True
Label17.Visible = True
 XPos = Rnd * 10000
 YPos = Rnd * 6000
 Image2.Top = YPos
 Image2.Left = XPos
 Label17.Top = Image2.Top + 240
 Label17.Left = Image2.Left + 480
 If XPos > 11000 Or XPos < 6000 Then
 XInc = -XInc
 End If
 If YPos > 10000 Or YPos < 5000 Then
 YInc = -YInc
 End If
If s = 1 Then Me.BackColor = vbBlue
If s = 2 Then Me.BackColor = vbGreen
If s = 3 Then Me.BackColor = vbYellow
If s = 4 Then Me.BackColor = vbRed
If s = 4 Then s = 0
End If
End Sub

Private Sub sq_Timer()
 Line (GetRndV(10000), GetRndV(10000))- _
 Step(GetRndV(1000), GetRndV(1000)), _
 RGB(GetRndV(256), GetRndV(256), GetRndV(256)), B
End Sub

Private Sub open_Click()
'mnu open
Command2_Click
End Sub

Private Sub paste_Click()
'mnu paste
Command9_Click
End Sub

Private Sub pp_Timer()
 Ball.Visible = True
 XPos = XPos + XInc
 YPos = YPos + YInc
 Ball.Top = YPos
 Ball.Left = XPos
 If XPos > 11000 Or XPos < -1000 Then
 XInc = -XInc
 End If
 If YPos > 10000 Or YPos < -1000 Then
 YInc = -YInc
 End If
End Sub

Private Sub ppr_Timer()
 Ball.Visible = True
 XPos = XPos + XInc
 YPos = YPos + YInc
 Ball.Top = YPos
 Ball.Left = XPos
 If XPos > 11000 Or XPos < -1000 Then
 XInc = -XInc
 End If
 If YPos > 10000 Or YPos < -1000 Then
 YInc = -YInc
 End If
End Sub

Private Sub print_Click()
'mnu print
Command6_Click
End Sub

Private Sub r_Change()
Shape1.BackColor = RGB(r.Value, g.Value, b.Value)
rt.Text = r.Value
End Sub

Private Sub RichTextBox1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
screensaver = 0
List1.Text = scres.Text
Text2.Text = Text6.Text
Combo8.Text = Second.Text
On Error GoTo err:
HScroll1.Value = RichTextBox1.SelFontSize
Combo1.Text = RichTextBox1.SelFontSize
Combo2.Text = RichTextBox1.SelFontName
err:
Exit Sub
End Sub

Private Sub rndc_Timer()
Ball.FillColor = QBColor(Rnd * 15)
End Sub

Private Sub rt_Change()
rt.Text = Val(rt.Text)
If rt.Text = "" Then
rt.Text = "0"
r.Value = rt.Text
ElseIf rt.Text > 255 Then
rt.Text = "255"
r.Value = rt.Text
ElseIf rt.Text < 0 Then
rt.Text = "0"
r.Value = rt.Text
ElseIf rt.Text > 0 And 255 Then
r.Value = rt.Text
Else
rt.Text = "0"
End If
End Sub

Private Sub save_Click()
'mnu save
Command3_Click
End Sub

Private Sub saveas_Click()
'mnu saveas
Command4_Click
End Sub

Private Sub simp1_Click()
'mnu 1simp
Command23_Click
End Sub

Private Sub size_Click()
'mnu size
On Error GoTo err:
Dim fosiz As Integer
fosiz = Val(InputBox("Enter between 8 and 100 number.", "Font Size"))
    Select Case fosiz
    Case Is < 8
    MsgBox "The number must beween 8 and 100 number", vbCritical + vbOKOnly, "Font Size (error)"
    Case Is > 100
    MsgBox "The number must beween 8 and 100 number", vbCritical + vbOKOnly, "Font Size (error)"
    Case 8 To 100
    RichTextBox1.SelFontSize = fosiz
    Case Else
    MsgBox "The number is invalid", vbCritical + vbOKOnly, "Font Size (error)"
    End Select
err:
Exit Sub
End Sub

Private Sub ss_Timer()
Static X As Byte
X = X + 1
If X = 1 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.1 (1).JPG")
If X = 2 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.1 (2).JPG")
If X = 3 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.1 (3).JPG")
If X = 4 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.2 (1).JPG")
If X = 5 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.2 (2).JPG")
If X = 6 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.3 (1).JPG")
If X = 7 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.3 (2).JPG")
If X = 8 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.4.JPG")
If X = 9 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.5 (1).JPG")
If X = 10 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.5 (2).JPG")
If X = 11 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.6 (1).JPG")
If X = 12 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.6 (2).JPG")
If X = 13 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.6 (3).JPG")
If X = 14 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.6 (4).JPG")
If X = 15 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.7 (1).JPG")
If X = 16 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.7 (2).JPG")
If X = 17 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.8.JPG")
If X = 18 Then Image1.Picture = LoadPicture("Other\Old versions\SS 2016 v.1 (1).JPG")
If X = 19 Then Image1.Picture = LoadPicture("Other\Old versions\SS 2016 v.1 (2).JPG")
If X = 20 Then Image1.Picture = LoadPicture("Other\Old versions\SS 2016 v.1 (3).JPG")
If X = 21 Then Image1.Picture = LoadPicture("Other\Old versions\SS 2016 v.1 (4).JPG")
If X = 21 Then X = 0
End Sub

Private Sub ssaver_Click()

End Sub

Private Sub ssr_Timer()
Static X As Byte
X = (Rnd * 21)
Randomize
If X = 1 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.1 (1).JPG")
If X = 2 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.1 (2).JPG")
If X = 3 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.1 (3).JPG")
If X = 4 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.2 (1).JPG")
If X = 5 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.2 (2).JPG")
If X = 6 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.3 (1).JPG")
If X = 7 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.3 (2).JPG")
If X = 8 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.4.JPG")
If X = 9 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.5 (1).JPG")
If X = 10 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.5 (2).JPG")
If X = 11 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.6 (1).JPG")
If X = 12 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.6 (2).JPG")
If X = 13 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.6 (3).JPG")
If X = 14 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.6 (4).JPG")
If X = 15 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.7 (1).JPG")
If X = 16 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.7 (2).JPG")
If X = 17 Then Image1.Picture = LoadPicture("Other\Old versions\SS 1.8.JPG")
If X = 18 Then Image1.Picture = LoadPicture("Other\Old versions\SS 2016 v.1 (1).JPG")
If X = 19 Then Image1.Picture = LoadPicture("Other\Old versions\SS 2016 v.1 (2).JPG")
If X = 20 Then Image1.Picture = LoadPicture("Other\Old versions\SS 2016 v.1 (3).JPG")
If X = 21 Then Image1.Picture = LoadPicture("Other\Old versions\SS 2016 v.1 (4).JPG")
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
screensaver = 0
Tips.Caption = " This is the Tool bar "
     List1.Text = scres.Text
     Text2.Text = Text6.Text
     Combo8.Text = Second.Text
End Sub

Private Sub SSTab2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
screensaver = 0
Tips.Caption = " You can get font properties here. "
     List1.Text = scres.Text
     Text2.Text = Text6.Text
     Combo8.Text = Second.Text
End Sub

Private Sub SSTab3_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
screensaver = 0
End Sub

Private Sub SSTab4_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
screensaver = 0
Tips.Caption = " You can change settings of screen saver here. "
     List1.Text = scres.Text
     Text2.Text = Text6.Text
     Combo8.Text = Second.Text
End Sub

Private Sub StatusBar1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " This is the Status bar which caps lock, num lock, scroll lock, instert, date, time, version shown to you. "
End Sub

Private Sub strikethrought_Click()
'mnu strike throught
Command15_Click
End Sub

Private Sub caseu_Click()
'mnu ucase
Command20_Click
End Sub

Private Sub sw_Timer()
dem.Text = dem.Text + 1
If dem.Text = "99" Then
sec.Text = sec.Text + 1
dem.Text = "00"
Else
If sec.Text = "59" Then
Min.Text = Min + 1
sec.Text = "00"
Else
If Min.Text = "59" Then
hrs.Text = hrs + 1
Min.Text = "00"
End If
End If
End If
End Sub

Private Sub tbox_Click()
On Error GoTo err:
CommonDialog1.ShowColor
RichTextBox1.BackColor = CommonDialog1.Color
err:
Exit Sub
End Sub

Private Sub Tips_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can know tips from me. "
End Sub

Private Sub tt_Timer()
screensaver.Text = screensaver.Text + 1
If Second.Text = "Seconds" Then
    If screensaver.Text = Val(Text6.Text) Then
    RichTextBox1.Visible = False
    SSTab1.Visible = False
    Call ssw
    Tips.Caption = " You can see screen saver now. "
    StatusBar1.Visible = False
    m.Visible = False
    e.Visible = False
    f.Visible = False
    tools.Visible = False
about.Visible = False
    End If
ElseIf Second.Text = "Minutes" Then
    RichTextBox1.Visible = False
    SSTab1.Visible = False
    If screensaver.Text = Val(Text6.Text) * 60 Then
    Call ssw
    Tips.Caption = " You can see screen saver now. "
    End If
End If
End Sub

Private Sub underline_Click()
'mnu underline
Command14_Click
End Sub

Function GetR(Max) As Integer
GetR = Int(Rnd * Max)
End Function

Private Function GetRndV(Max) As Integer
GetRndV = Int(Rnd * Max)
End Function


Private Sub ssw()
If scres.Text = "Color magic" Then
    If Check10.Value = 0 Then
        cm.Enabled = True
        If Combo7.Text = "Seconds" Then cm.Interval = Text1.Text * 1000
        If Combo7.Text = "Minutes" Then cm.Interval = Text1.Text * 1000 * 60
    Else
        cmr.Enabled = True
          If Combo7.Text = "Seconds" Then cmr.Interval = Text1.Text * 1000
          If Combo7.Text = "Minutes" Then cmr.Interval = Text1.Text * 1000 * 60
    End If
    
ElseIf scres.Text = "Ping Pong" Then
    If Check2.Value = 0 Then
        pp.Enabled = True
        pp.Interval = Text3.Text
        If Combo9.Text = "Blue" Then Ball.FillColor = vbBlue
        If Combo9.Text = "Black" Then Ball.FillColor = vbBlack
        If Combo9.Text = "Green" Then Ball.FillColor = vbGreen
        If Combo9.Text = "Red" Then Ball.FillColor = vbRed
        If Combo9.Text = "White" Then Ball.FillColor = vbWhite
        If Combo9.Text = "Yellow" Then Ball.FillColor = vbYellow
        XInc = 25 + GetR(50)
        YInc = 25 + GetR(50)
        If GetR(3) Mod 2 = 1 Then
        XInc = -XInc
        End If
        If GetR(3) Mod 2 = 1 Then
        YInc = -YInc
        End If
        XPos = GetR(10000) + 500
        YPos = GetR(10000) + 500
    Else
        pp.Enabled = True
        rndc.Enabled = True
        pp.Interval = Text3.Text
        Ball.FillColor = QBColor(Rnd * 15)
         XInc = 25 + GetR(50)
        YInc = 25 + GetR(50)
        If GetR(3) Mod 2 = 1 Then
        XInc = -XInc
        End If
        If GetR(3) Mod 2 = 1 Then
        YInc = -YInc
        End If
        XPos = GetR(10000) + 500
        YPos = GetR(10000) + 500
    End If
    
ElseIf scres.Text = "Objects" Then
    
    If Check3.Value = 1 Then
        sq.Enabled = True
        sq.Interval = Text4.Text
        Else
        sq.Enabled = False
    End If
    If Check5.Value = 1 Then
        bx.Enabled = True
        bx.Interval = Text4.Text
        Else
        bx.Enabled = False
   End If
    If Check6.Value = 1 Then
        ru.Enabled = True
        ru.Interval = Text4.Text
        Else
        ru.Enabled = False
   End If
    If Check9.Value = 1 Then
        bg.Enabled = True
        Else
        bg.Enabled = False
    End If

    
   
    
ElseIf scres.Text = "SriSoft Write logo" Then
    If Check7.Value = 0 Then
    Image2.Visible = True
    Label17.Visible = True
    Else
    Image2.Visible = True
    Label17.Visible = True
        sl.Enabled = True
        If Combo10.Text = "Seconds" Then sl.Interval = Text5.Text * 1000
        If Combo10.Text = "Minutes" Then sl.Interval = Text5.Text * 1000 * 60
        XInc = 25 + GetR(50)
        YInc = 25 + GetR(50)
        If GetR(3) Mod 2 = 1 Then
        XInc = -XInc
        End If
        If GetR(3) Mod 2 = 1 Then
        YInc = -YInc
        End If
        XPos = GetR(10000) + 500
        YPos = GetR(6000) + 500
    End If
ElseIf scres.Text = "Stop Watch" Then
    sw.Enabled = True
    dem.Visible = True
    sec.Visible = True
    Min.Visible = True
    hrs.Visible = True
    Label12.Visible = True
    Label13.Visible = True
    Label14.Visible = True
ElseIf scres.Text = "Slide Show" Then
   If Check12.Value = 0 Then
    ss.Enabled = True
    Image1.Visible = True
    If Combo11.Text = "Seconds" Then ss.Interval = Text7.Text * 1000
    If Combo11.Text = "Minutes" Then ss.Interval = Text7.Text * 1000 * 60
    Else
    ssr.Enabled = True
    Image1.Visible = True
    If Combo11.Text = "Seconds" Then ssr.Interval = Text7.Text * 1000
    If Combo11.Text = "Minutes" Then ssr.Interval = Text7.Text * 1000 * 60
    End If
Else
        tt.Enabled = False
        Cls
        Form2.BackColor = RGB(Text8.Text, Text9.Text, Text10.Text)
        Label12.Visible = False
        Label13.Visible = False
        Label14.Visible = False
        Image1.Visible = False
        ss.Enabled = False
        ssr.Enabled = False
        dem.Visible = False
        sec.Visible = False
        Min.Visible = False
        hrs.Visible = False
        sl.Enabled = False
        Ball.Visible = False
        pp.Enabled = True
        cm.Enabled = False
        cmr.Enabled = False
        sq.Enabled = False
        rndc.Enabled = False
        bx.Enabled = False
        ru.Enabled = False
        bg.Enabled = False
        Image2.Visible = False
        Label17.Visible = False
        sw.Enabled = False
End If
End Sub
