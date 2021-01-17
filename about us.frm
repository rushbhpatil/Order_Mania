VERSION 5.00
Begin VB.Form Form12 
   Caption         =   "About Us"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9825
   LinkTopic       =   "Form12"
   ScaleHeight     =   9360
   ScaleWidth      =   9825
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Back"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   11640
      Left            =   0
      Picture         =   "about us.frx":0000
      Stretch         =   -1  'True
      Top             =   -1320
      Width           =   9855
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
Unload Me
End Sub

