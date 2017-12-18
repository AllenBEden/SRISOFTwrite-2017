VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   Caption         =   "SRISOFTwrite 2017"
   ClientHeight    =   5775
   ClientLeft      =   2100
   ClientTop       =   2235
   ClientWidth     =   9990
   FillColor       =   &H00C0C0FF&
   FillStyle       =   0  'Solid
   Icon            =   "loadfrm(SSw2017).frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "loadfrm(SSw2017).frx":0ECA
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   9990
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8160
      Top             =   360
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   150
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Visible         =   0   'False
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   265
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      MaxLength       =   23
      MouseIcon       =   "loadfrm(SSw2017).frx":1794
      MousePointer    =   99  'Custom
      PasswordChar    =   "X"
      TabIndex        =   7
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Show"
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   3600
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "loadfrm(SSw2017).frx":205E
      TabIndex        =   6
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Run"
      Height          =   255
      Left            =   1080
      MaskColor       =   &H00C0C0FF&
      MouseIcon       =   "loadfrm(SSw2017).frx":2928
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8400
      Top             =   360
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8640
      Top             =   360
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"loadfrm(SSw2017).frx":31F2
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   8880
      Top             =   360
   End
   Begin VB.CommandButton Command1 
      DownPicture     =   "loadfrm(SSw2017).frx":3274
      Height          =   375
      Left            =   9360
      Picture         =   "loadfrm(SSw2017).frx":37F6
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   360
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      BorderStyle     =   4  'Dash-Dot
      FillColor       =   &H000000FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   1095
      Left            =   960
      Top             =   4200
      Width           =   3495
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " You can know tips from me. "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2115
   End
   Begin VB.Label Label1 
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
      Left            =   2400
      TabIndex        =   0
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   1920
      Picture         =   "loadfrm(SSw2017).frx":3D78
      Top             =   1560
      Width           =   720
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   5550
      Left            =   120
      Picture         =   "loadfrm(SSw2017).frx":4C42
      Stretch         =   -1  'True
      Top             =   120
      Width           =   9720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If Text1.PasswordChar = "X" Then
Text1.PasswordChar = ""
Else
Text1.PasswordChar = "X"
End If
End Sub

Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Label2.Caption = " You can change visible of password from click here. "
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Label2.Caption = " You can close the application from click here. "
End Sub

Private Sub Command2_Click()
Form2.Show
Unload Me
End Sub

Private Sub Command3_Click()
If Text1.Text = RichTextBox1.Text Then
Text1.Enabled = False
Check1.Enabled = False
Command3.Enabled = False
Timer2.Enabled = True
ProgressBar1.Visible = True
Command1.Visible = False
Else
Text2.Text = Val(Text2.Text) + "1"
Check1.Value = 1
Text1.Text = " Password is incorrect. "
End If
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Label2.Caption = " You can enter your files from click here. "
End Sub

Private Sub Form_Load()
RichTextBox1.LoadFile (App.Path & "\Files\PW\PW.RTF")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Label2.Caption = " You can know tips from me. "
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Label2.Caption = " You can know tips from me. "
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Label2.Caption = " You can type your password here. "
End Sub

Private Sub Text2_Change()
If Text2.Text = "3" Then
Timer3.Enabled = True
Text1.Enabled = False
Check1.Enabled = False
Command3.Enabled = False
Command1.Visible = False
End If
End Sub

Private Sub Timer1_Timer()
Static s As Integer
s = s + 1
If s = 1 Then Form1.MouseIcon = LoadPicture(App.Path & "\Files\Icons\SS\arrow blue.ICO")
If s = 2 Then Form1.MouseIcon = LoadPicture(App.Path & "\Files\Icons\SS\arrow green.ICO")
If s = 3 Then Form1.MouseIcon = LoadPicture(App.Path & "\Files\Icons\SS\arrow yellow.ICO")
If s = 4 Then Form1.MouseIcon = LoadPicture(App.Path & "\Files\Icons\SS\arrow red.ICO")
If s = 4 Then s = 0
End Sub

Private Sub Timer2_Timer()
If ProgressBar1.Value = 100 Then
Timer2.Enabled = False
Timer4.Enabled = True
Else
ProgressBar1.Value = ProgressBar1.Value + 1
Label2.Caption = " Loading... " & "(" & ProgressBar1.Value & "%) "
End If
End Sub

Private Sub Timer3_Timer()
Label2.Caption = " Your files are searching by thief "
Static s As Integer
s = s + 1
If s = 5 Then End
End Sub


Private Sub Timer4_Timer()
s = s + 1
If s = 1 Then
Form2.Show
Unload Me
End If
End Sub
