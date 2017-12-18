VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sri Calculator (v.2.0)"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12450
   Icon            =   "calfrm.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   12450
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   630
      Left            =   120
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H0000C000&
      Caption         =   "Paste text"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H0000C000&
      Caption         =   "Copy Text"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H0000C000&
      Caption         =   "Delete text"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H0000C000&
      Caption         =   "Reset"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4680
      Width           =   4815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "With demical places."
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   5760
      Value           =   1  'Checked
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   630
      Left            =   120
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   840
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H008080FF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H008080FF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command0 
      BackColor       =   &H008080FF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H008080FF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H008080FF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command00 
      BackColor       =   &H008080FF&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H008080FF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H008080FF&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H0000FFFF&
      Caption         =   "Total"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   4815
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H008080FF&
      Caption         =   "+ (Addition)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Narkisim"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H008080FF&
      Caption         =   "- (Subtraction)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Narkisim"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H008080FF&
      Caption         =   "x (Multiplication)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Narkisim"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   1695
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H008080FF&
      Caption         =   "/ (Divide)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Narkisim"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   720
      Left            =   11640
      Picture         =   "calfrm.frx":0ECA
      Top             =   5400
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
      Left            =   10080
      TabIndex        =   26
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   6015
      Left            =   5040
      Picture         =   "calfrm.frx":1D94
      Stretch         =   -1  'True
      Top             =   120
      Width           =   7335
   End
   Begin VB.Label tips 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " You can know tips from me. "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   6240
      Width           =   12435
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   6720
      Picture         =   "calfrm.frx":1118EE
      Top             =   600
      Width           =   720
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
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7200
      TabIndex        =   23
      Top             =   840
      Width           =   2055
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   4920
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   4920
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   4920
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line2 
      X1              =   3120
      X2              =   3120
      Y1              =   1680
      Y2              =   4080
   End
   Begin VB.Line Line1 
      X1              =   4920
      X2              =   120
      Y1              =   1560
      Y2              =   1560
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can change total with demical places or without demical places by tick here. "
End Sub

Private Sub Command0_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can type shown number by click here. "
End Sub

Private Sub Command00_Click()
Text1.Text = Text1.Text + "00"
End Sub

Private Sub Command0_Click()
Text1.Text = Text1.Text + "0"
End Sub

Private Sub Command00_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can type shown number by click here. "
End Sub

Private Sub Command1_Click()
Text1.Text = Text1.Text + "1"
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can type shown number by click here. "
End Sub

Private Sub Command10_Click()
If Option1.Value = True Then
    If Check1.Value = 1 Then
    Text1.Text = Val(Text2.Text) + Val(Text1.Text)
    Else
    Text1.Text = Int(Val(Text2.Text)) + Int(Val(Text1.Text))
    End If
ElseIf Option2.Value = True Then
    If Check1.Value = 1 Then
    Text1.Text = Val(Text2.Text) - Val(Text1.Text)
    Else
    Text1.Text = Int(Val(Text2.Text)) - Int(Val(Text1.Text))
    End If
ElseIf Option3.Value = True Then
    If Check1.Value = 1 Then
    Text1.Text = Val(Text2.Text) * Val(Text1.Text)
    Else
    Text1.Text = Int(Val(Text2.Text)) * Int(Val(Text1.Text))
    End If
ElseIf Option4.Value = True Then
    If Check1.Value = 1 Then
    Text1.Text = Val(Text2.Text) / Val(Text1.Text)
    Else
    Text1.Text = Int(Val(Text2.Text)) / Int(Val(Text1.Text))
    End If
End If
End Sub

Private Sub Command10_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can get total by click here. "
End Sub

Private Sub Command11_Click()
Text1.Text = Text1.Text + "."
End Sub

Private Sub Command11_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can type demical point from click here. "
End Sub

Private Sub Command13_Click()
Text1.Text = ""
End Sub

Private Sub Command13_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can delete active number by click here. "
End Sub

Private Sub Command14_Click()
Clipboard.SetText (Text1.Text)
End Sub

Private Sub Command14_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can copy active number by click here. "
End Sub

Private Sub Command15_Click()
Text1.Text = Clipboard.GetText
End Sub

Private Sub Command15_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can paste to active number by click here. "
End Sub

Private Sub Command16_Click()
'reset
Text1.Text = ""
Text2.Text = ""
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Command10.Enabled = False
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
End Sub

Private Sub Command16_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can reset calculator by click here. "
End Sub

Private Sub Command2_Click()
Text1.Text = Text1.Text + "2"
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can type shown number by click here. "
End Sub

Private Sub Command3_Click()
Text1.Text = Text1.Text + "3"
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can type shown number by click here. "
End Sub

Private Sub Command4_Click()
Text1.Text = Text1.Text + "4"
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can type shown number by click here. "
End Sub

Private Sub Command5_Click()
Text1.Text = Text1.Text + "5"
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can type shown number by click here. "
End Sub

Private Sub Command6_Click()
Text1.Text = Text1.Text + "6"
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can type shown number by click here. "
End Sub

Private Sub Command7_Click()
Text1.Text = Text1.Text + "7"
End Sub

Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can type shown number by click here. "
End Sub

Private Sub Command8_Click()
Text1.Text = Text1.Text + "8"
End Sub

Private Sub Command8_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can type shown number by click here. "
End Sub

Private Sub Command9_Click()
Text1.Text = Text1.Text + "9"
End Sub

Private Sub Command9_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can type shown number by click here. "
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can know tips from me. "
End Sub

Private Sub Image1_DblClick()
Form9.Show (vbModal)
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can see about 'SRI Calculator' by double click here. "
End Sub

Private Sub Option1_Click()
Text2.Text = Text1.Text
Text1.Text = ""
Command10.Enabled = True
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
End Sub

Private Sub Option1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can type shown math mark by click here. "
End Sub

Private Sub Option2_Click()
Text2.Text = Text1.Text
Text1.Text = ""
Command10.Enabled = True
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
End Sub

Private Sub Option2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can type shown math mark by click here. "
End Sub

Private Sub Option3_Click()
Text2.Text = Text1.Text
Text1.Text = ""
Command10.Enabled = True
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
End Sub

Private Sub Option3_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can type shown math mark by click here. "
End Sub

Private Sub Option4_Click()
Text2.Text = Text1.Text
Text1.Text = ""
Command10.Enabled = True
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
End Sub

Private Sub Option4_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can type shown math mark by click here. "
End Sub

Private Sub Text1_Change()
If Text1.Text = "" Then
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Else
    If Text2.Text = "" Then
    Option1.Enabled = True
    Option2.Enabled = True
    Option3.Enabled = True
    Option4.Enabled = True
    Else
    Option1.Enabled = False
    Option2.Enabled = False
    Option3.Enabled = False
    Option4.Enabled = False
    End If
End If
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can see active number and total from here. "
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can see inactive(last) number from here. "
End Sub
