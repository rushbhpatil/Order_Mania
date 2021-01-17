VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   Caption         =   "Order Mania"
   ClientHeight    =   8895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   ScaleHeight     =   17075.93
   ScaleMode       =   0  'User
   ScaleWidth      =   3580.119
   WindowState     =   2  'Maximized
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Feedback"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   18825
      TabIndex        =   0
      Top             =   8790
      Width           =   1215
   End
   Begin VB.Image Image7 
      Height          =   1695
      Left            =   16575
      Picture         =   "Front.frx":0000
      Stretch         =   -1  'True
      Top             =   9315
      Width           =   1695
   End
   Begin VB.Image Image6 
      Height          =   1200
      Left            =   18915
      Picture         =   "Front.frx":10692
      Stretch         =   -1  'True
      Top             =   7605
      Width           =   1095
   End
   Begin VB.Image Image5 
      Height          =   1335
      Left            =   18795
      Picture         =   "Front.frx":19405
      Stretch         =   -1  'True
      Top             =   9480
      Width           =   1335
   End
   Begin VB.Image Image4 
      Height          =   2655
      Left            =   5010
      Picture         =   "Front.frx":25364
      Stretch         =   -1  'True
      Top             =   855
      Width           =   10455
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   19440
      Picture         =   "Front.frx":2B1B4
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   4320
      Left            =   11265
      Picture         =   "Front.frx":2C506
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   3690
   End
   Begin VB.Image Image1 
      Height          =   4320
      Left            =   6105
      Picture         =   "Front.frx":5BB5C
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   3525
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
End
End Sub




Private Sub Image1_Click()
Form2.Show
Unload Me
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 1
End Sub

Private Sub Image2_Click()
Form4.Show
Unload Me
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.BorderStyle = 1
End Sub

Private Sub Image3_Click()
Dim a As String

a = MsgBox("Are you sure you want to Exit?", vbYesNo, "Alert!!")

If a = vbYes Then
End
Else
Exit Sub
End If
End Sub

Private Sub Image5_Click()
Form12.Show
Unload Me
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.BorderStyle = 1
End Sub

Private Sub Image6_Click()
Form13.Show
Unload Me
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.BorderStyle = 1
End Sub


Private Sub Image7_Click()
Form15.Show
Unload Me
End Sub

Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.BorderStyle = 1
End Sub

Private Sub Label1_Click()
Image6.BorderStyle = 1
Form13.Show
Unload Me
End Sub

Private Sub Label18_Click()
Form13.Show
Unload Me
End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.BorderStyle = 1
End Sub
