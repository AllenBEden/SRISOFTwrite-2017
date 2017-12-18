VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sri Calendar (v.2.0)"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9945
   Icon            =   "celfrm.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox y 
      Height          =   315
      Left            =   1560
      TabIndex        =   11
      Top             =   3120
      Width           =   945
   End
   Begin VB.ComboBox m 
      Height          =   1155
      ItemData        =   "celfrm.frx":0ECA
      Left            =   840
      List            =   "celfrm.frx":0EF2
      Style           =   1  'Simple Combo
      TabIndex        =   10
      Text            =   "1"
      Top             =   3120
      Width           =   615
   End
   Begin VB.ComboBox d 
      Height          =   1155
      ItemData        =   "celfrm.frx":0F1D
      Left            =   120
      List            =   "celfrm.frx":0F7E
      Style           =   1  'Simple Combo
      TabIndex        =   9
      Text            =   "1"
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Go"
      Height          =   735
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3480
      Width           =   945
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   5400
      Width           =   3060
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2640
      Top             =   3720
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   4680
      Width           =   3060
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2310
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   -2147483635
      BackColor       =   16711680
      BorderStyle     =   1
      Appearance      =   0
      MaxSelCount     =   10
      MonthBackColor  =   12648447
      MultiSelect     =   -1  'True
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   193069058
      TitleBackColor  =   49152
      TitleForeColor  =   192
      TrailingForeColor=   12632319
      CurrentDate     =   36526
      MaxDate         =   402133
      MinDate         =   -45653
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
      TabIndex        =   14
      Top             =   5760
      Width           =   9975
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
      Left            =   3840
      TabIndex        =   13
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   3360
      Picture         =   "celfrm.frx":0FF5
      Top             =   4800
      Width           =   720
   End
   Begin VB.Label Label6 
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Day"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   345
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3120
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Go to the date -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   1605
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Time :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   540
   End
   Begin VB.Image Image1 
      Height          =   5535
      Left            =   3240
      Picture         =   "celfrm.frx":1EBF
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim past As String
On Error GoTo errormes:
MonthView1.Value = d.Text + "/" + m.Text + "/" + y.Text
GoTo proend
errormes:
past = MsgBox("Check your date again. ", _
vbCritical + vbOKOnly, "Error")
proend:
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can jump to the date. "
End Sub

Private Sub Form_Load()
MonthView1.Value = Date
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can know tips from me. "
End Sub

Private Sub Image1_DblClick()
Form7.Show (vbModal)
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can see about 'SRI Calendar' by double click here. "
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can see time of now. "
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Tips.Caption = " You can see date of today. "
End Sub

Private Sub Timer1_Timer()
Text1.Text = Date
Text2.Text = Time
End Sub

