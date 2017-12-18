VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SRISOFTwrite 2017 Password changer"
   ClientHeight    =   2175
   ClientLeft      =   5940
   ClientTop       =   7140
   ClientWidth     =   6255
   FillColor       =   &H00FF0000&
   Icon            =   "password changer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "Show"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   10
      Top             =   480
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin RichTextLib.RichTextBox rt 
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Top             =   840
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   503
      _Version        =   393217
      MultiLine       =   0   'False
      MaxLength       =   23
      TextRTF         =   $"password changer.frx":0ECA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"password changer.frx":0F45
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   6015
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2400
      MaxLength       =   23
      PasswordChar    =   "X"
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2400
      MaxLength       =   23
      PasswordChar    =   "X"
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SRISOFTwrite 2017"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1920
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1560
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1200
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   840
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   480
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Re-Enter your new password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Enter your new password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Enter your current password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As String
Dim d As String

Private Sub Check1_Click()
If Check1.Value = 1 Then
Text1.PasswordChar = ""
Else
Text1.PasswordChar = "X"
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Text2.PasswordChar = ""
Else
Text2.PasswordChar = "X"
End If
End Sub

Private Sub Command1_Click()
If Text1.Text = RichTextBox1.Text Then
    If rt.Text = Text2.Text Then
    rt.SaveFile (App.Path & "\Files\PW\PW.RTF")
    End
    Else
    s = MsgBox("Re-Enter password is wrong.", vbCritical + vbOKOnly, "error(Retype)")
    End If
ElseIf Text1.Text <> RichTextBox1 Then
d = MsgBox("Your current password is wrong.", vbCritical + vbOKOnly, "error(Retype)")
End If
End Sub

Private Sub Form_Load()
RichTextBox1.LoadFile (App.Path & "\Files\PW\PW.RTF")
End Sub
