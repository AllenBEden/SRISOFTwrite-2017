VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sri Song Player (v.1.0)"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8145
   Icon            =   "sonfrm.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   0
      TabIndex        =   6
      Top             =   1800
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "Open"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Volume"
      Height          =   375
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0000FF00&
      Caption         =   "Mute"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7320
      MaskColor       =   &H00C0FFFF&
      TabIndex        =   2
      Top             =   1440
      Width           =   735
   End
   Begin MCI.MMControl MMControl1 
      Height          =   450
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   794
      _Version        =   393216
      Frames          =   2
      RecordMode      =   0
      UpdateInterval  =   1
      PlayEnabled     =   -1  'True
      PauseEnabled    =   -1  'True
      AutoEnable      =   0   'False
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7080
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      TabIndex        =   7
      Top             =   7080
      Width           =   8175
   End
   Begin VB.Image Image2 
      Height          =   5295
      Left            =   -480
      Picture         =   "sonfrm.frx":0ECA
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   8655
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   240
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File name"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   675
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   0
      Picture         =   "sonfrm.frx":D0A2F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
MMControl1.Silent = True
Else
MMControl1.Silent = False
End If
End Sub

Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can mute movie or song here. "
End Sub

Private Sub Command1_Click()
Shell "sndvol32.exe", vbNormalFocus
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can change volume controls by click here. "
End Sub

Private Sub Command2_Click()
CommonDialog1.DialogTitle = "Open"
CommonDialog1.Filter = "Video File Format(*.AVI)/*.avi"
CommonDialog1.ShowOpen
MMControl1.FileName = CommonDialog1.FileName
Label2.Caption = CommonDialog1.FileName
MMControl1.Command = "Open"
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can open any avi file by click here. "
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can know tips from me. "
End Sub

Private Sub Image2_DblClick()
Form10.Show (vbModal)
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can see about 'SRI Song Player' by double click here. "
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can see loaded avi file's path here. "
End Sub

Private Sub MMControl1_PlayClick(Cancel As Integer)
MMControl1.Command = "prev"
End Sub

Private Sub MMControl1_StatusUpdate()
ProgressBar1.Max = MMControl1.Length + 1
ProgressBar1.Value = MMControl1.Position
End Sub


Private Sub Tips_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can know tips from me. "
End Sub
