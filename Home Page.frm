VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   Caption         =   "Welcome"
   ClientHeight    =   8895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleMode       =   0  'User
   ScaleWidth      =   1494.845
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "bill"
      Height          =   1695
      Left            =   1560
      TabIndex        =   4
      Top             =   3480
      Width           =   2415
   End
   Begin VB.CommandButton cmdCust 
      Height          =   1575
      Left            =   10320
      Picture         =   "Front.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7680
      Width           =   2655
   End
   Begin VB.CommandButton cmdAdmin 
      Height          =   1575
      Left            =   5280
      Picture         =   "Front.frx":1B99
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7680
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C000C0&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H000000C0&
      Height          =   1455
      Left            =   4890
      TabIndex        =   3
      Top             =   1965
      Width           =   8535
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   1095
      Left            =   7155
      Shape           =   4  'Rounded Rectangle
      Top             =   1725
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   2880
      TabIndex        =   2
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Image Image3 
      Height          =   750
      Left            =   16320
      Picture         =   "Front.frx":3C5D
      Top             =   1320
      Width           =   750
   End
   Begin VB.Image Image2 
      Height          =   2250
      Left            =   10560
      Picture         =   "Front.frx":4FAF
      Top             =   5280
      Width           =   2250
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   5280
      Picture         =   "Front.frx":C5A4
      Top             =   5160
      Width           =   2250
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdAdmin_Click()
Form2.Show
Form1.Hide
End Sub

Private Sub cmdCust_Click()
Form4.Show
End Sub

Private Sub Command1_Click()
Form6.Show
End Sub

Private Sub Command2_Click()
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("D:\OrderMania\ordermania.mdb")
Set rs = db.OpenRecordset("select * from order1")

While Not rs.EOF
rs.Edit
rs.Fields(8).Value = "Paid"

rs.Update
rs.MoveNext

Wend

End Sub

Private Sub Image3_Click()
End
End Sub

