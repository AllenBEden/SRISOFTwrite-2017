VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form help 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Help about tools"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9630
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   825
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1455
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "ImageCombo1"
      ImageList       =   "ImageList1"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8040
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   57
      ImageHeight     =   49
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "helptoolfrm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "helptoolfrm.frx":05B6
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
ImageCombo1.ComboItems.Add 1, "K1", "Nero Software", 1
ImageCombo1.ComboItems.Add 2, "K2", "Microsoft Word", 2



End Sub
