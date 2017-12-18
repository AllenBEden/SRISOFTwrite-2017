VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Font Size"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4470
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Apply"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "OK"
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      Height          =   4275
      ItemData        =   "font form.frx":0000
      Left            =   120
      List            =   "font form.frx":0002
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Form2.RichTextBox1.SelFontName = Combo2.Text
Unload Me
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form2.screensaver = 0
End Sub

Private Sub Command2_Click()
Form2.RichTextBox1.SelFontName = Combo2.Text
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form2.screensaver = 0
End Sub

Private Sub Form_Load()
'font load to font form
Dim i As Integer
For i = 1 To Screen.FontCount
Combo2.AddItem Screen.Fonts(i)
Next i
Combo2.Text = Form2.RichTextBox1.SelFontName
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form2.screensaver = 0
End Sub
