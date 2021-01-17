VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00404080&
   Caption         =   "General"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17325
   LinkTopic       =   "Form8"
   ScaleHeight     =   9135
   ScaleWidth      =   17325
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdlogout 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   18960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Feedback"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8610
      TabIndex        =   3
      Top             =   10395
      Width           =   2655
   End
   Begin VB.Image feedback 
      Height          =   1785
      Left            =   9000
      Picture         =   "general.frx":0000
      Stretch         =   -1  'True
      Top             =   8565
      Width           =   1830
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   1095
      Left            =   7845
      Shape           =   4  'Rounded Rectangle
      Top             =   1665
      Width           =   3855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00404080&
      Caption         =   "Order Mania"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   5640
      TabIndex        =   1
      Top             =   1920
      Width           =   8535
   End
   Begin VB.Image total_sale 
      Height          =   4695
      Left            =   13215
      Picture         =   "general.frx":C5DC
      Stretch         =   -1  'True
      Top             =   3705
      Width           =   3855
   End
   Begin VB.Image recent_order 
      Height          =   4455
      Left            =   7935
      Picture         =   "general.frx":111CF
      Stretch         =   -1  'True
      Top             =   3705
      Width           =   3735
   End
   Begin VB.Image invoice 
      Height          =   4815
      Left            =   2775
      Picture         =   "general.frx":1A787
      Stretch         =   -1  'True
      Top             =   3465
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   885
      TabIndex        =   0
      Top             =   1635
      Width           =   3015
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdlogout_Click()
Dim a As String

a = MsgBox("Are you sure you want to Logout?", vbYesNo, "Alert!!")

If a = vbYes Then
Form1.Show
Unload Me
Else
Exit Sub
End If

End Sub

Private Sub feedback_Click()
Form14.Show
Unload Me
End Sub

Private Sub feedback_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
feedback.BorderStyle = 1
End Sub

Private Sub invoice_Click()
Form6.Show
Unload Me
End Sub

Private Sub invoice_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
invoice.BorderStyle = 1
End Sub

Private Sub Label2_Click()
feedback.BorderStyle = 1
Form14.Show
Unload Me
End Sub

Private Sub recent_order_Click()
Form9.Show
Unload Me
End Sub

Private Sub recent_order_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
recent_order.BorderStyle = 1
End Sub

Private Sub total_sale_Click()
Form10.Show
Unload Me
End Sub

Private Sub total_sale_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
total_sale.BorderStyle = 1
End Sub
