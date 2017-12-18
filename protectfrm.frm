VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Protect"
   ClientHeight    =   2955
   ClientLeft      =   2835
   ClientTop       =   3360
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1745.911
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Ok"
      Height          =   255
      Left            =   240
      MaskColor       =   &H00C0C0FF&
      MouseIcon       =   "protectfrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   3255
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Show"
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   2760
      MaskColor       =   &H00FFFFFF&
      MouseIcon       =   "protectfrm.frx":08CA
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text3 
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
      Left            =   240
      MaxLength       =   23
      MouseIcon       =   "protectfrm.frx":1194
      MousePointer    =   99  'Custom
      PasswordChar    =   "X"
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "0"
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin RichTextLib.RichTextBox pw 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   450
      _Version        =   393217
      TextRTF         =   $"protectfrm.frx":1A5E
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Someone tries to enter your files."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Password"
      ForeColor       =   &H000000C0&
      Height          =   270
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   720
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      BorderStyle     =   4  'Dash-Dot
      FillColor       =   &H000000FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   1095
      Left            =   120
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sFile As String

Private Sub Check1_Click()
If Text3.PasswordChar = "X" Then
Text3.PasswordChar = ""
Else
Text3.PasswordChar = "X"
End If
End Sub

Private Sub Command3_Click()
If Text3.Text <> pw.Text Then
Timer1.Enabled = True
Else
If Text3.Text = pw.Text Then
With Form2
.SSTab1.Visible = True
.Cls
.screensaver = 0
.List1.Text = .scres.Text
.RichTextBox1.Visible = True
 Form2.BackColor = RGB(255, 192, 192)
.Label12.Visible = False
.Label13.Visible = False
.Label14.Visible = False
.Image1.Visible = False
.ss.Enabled = False
.ssr.Enabled = False
.dem.Visible = False
.sec.Visible = False
.min.Visible = False
.hrs.Visible = False
.sl.Enabled = False
.Ball.Visible = False
.pp.Enabled = False
.cm.Enabled = False
.cmr.Enabled = False
.sq.Enabled = False
.rndc.Enabled = False
.bx.Enabled = False
.ru.Enabled = False
.bg.Enabled = False
.Image2.Visible = False
.Label17.Visible = False
.sw.Enabled = False
End With
Unload Me


End If
End If

End Sub

Private Sub Form_Load()
pw.LoadFile (App.Path & "\Files\PW\PW.RTF")
End Sub

Private Sub Timer1_Timer()
Dim s As Integer
s = s + 1
If s = 1 Then MsgBox "Retype your password within 10 seconds.", vbCritical + vbOKOnly, "Wrong password"
If s = 1 Then Text3.Text = ""
If s = 10 Then Me.Height = 2175
If s = 10 Then Text3.Enabled = False
If s = 10 Then Command3.Enabled = False
If s = 10 Then Check1.Enabled = False
If s = 11 Then Label1.ForeColor = vbBlack
If s = 11 Then Label1.BackColor = vbRed
If s = 11 Then Beep
If s = 12 Then Label1.ForeColor = vbRed
If s = 12 Then Label1.BackColor = vbBlack
If s = 12 Then Beep
If s = 13 Then Label1.ForeColor = vbBlack
If s = 13 Then Label1.BackColor = vbRed
If s = 13 Then Beep
If s = 14 Then Label1.ForeColor = vbRed
If s = 14 Then Label1.BackColor = vbBlack
If s = 14 Then Beep
If s = 15 Then
'''''''''''''''''''''''''''''''''''
Unload Me
With Form2
If Left$(Form2.Caption, 18) = " SRISOFTwrite 2017" Then
    If Form2 Is Nothing Then Exit Sub
    .CommonDialog1.DialogTitle = "Save As"
    .CommonDialog1.CancelError = False
    .CommonDialog1.Filter = "All Files (*.*)|*.*"
    .CommonDialog1.ShowSave
        If Len(.CommonDialog1.FileName) = 0 Then
        Exit Sub
        End If
    sFile = .CommonDialog1.FileName
    Form2.Caption = sFile
    Form2.RichTextBox1.SaveFile sFile
        End
    Else
    sFile = Form2.Caption
    Form2.RichTextBox1.SaveFile sFile
        End
    End If
End With
    End If
    

    
End Sub
