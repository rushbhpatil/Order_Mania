VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000080&
   Caption         =   "Guest Login"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12060
   LinkTopic       =   "Form3"
   ScaleHeight     =   7845
   ScaleWidth      =   12060
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Height          =   975
      Left            =   4200
      Picture         =   "guestlogin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DataField       =   "tbno"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "orderado"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   4080
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      Caption         =   "Enter Your Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   2
      Top             =   3600
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   4560
      Picture         =   "guestlogin.frx":1B0F
      Top             =   960
      Width           =   2250
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset


Private Sub Command1_Click()


Form4.Show
Unload Me

End Sub

Private Sub Form_Load()
Text1.Text = ""
End Sub

Private Sub Image2_Click()

End Sub
