VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00C000C0&
   Caption         =   "Admin Login"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   11055
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtuser 
      BorderStyle     =   0  'None
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Top             =   3840
      Width           =   2895
   End
   Begin VB.TextBox txtpass 
      BorderStyle     =   0  'None
      DataSource      =   "Adodc1"
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   6240
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   4920
      Width           =   2895
   End
   Begin VB.CommandButton cmdexit 
      Height          =   975
      Left            =   9000
      Picture         =   "Form2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5880
      Width           =   2175
   End
   Begin VB.CommandButton cmdlogin 
      Height          =   975
      Left            =   6240
      Picture         =   "Form2.frx":173A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc loginado 
      Height          =   855
      Left            =   2400
      Top             =   7320
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\OrderMania\Database1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\OrderMania\Database1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Login"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   6480
      Picture         =   "Form2.frx":3249
      Top             =   960
      Width           =   2250
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C000C0&
      Caption         =   "Username"
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C000C0&
      Caption         =   "Password"
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   4560
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdlogin_Click()
loginado.RecordSource = "select * from Login where User_name='" + txtuser.Text + "' and Password='" + txtpass.Text + "'"
loginado.Refresh

If loginado.Recordset.EOF Then

MsgBox "Login failed,Try Again..!!!", vbCritical, "Please Enter correct Username and Password"
txtuser.Text = ""
txtpass.Text = ""
txtuser.SetFocus
Else
Form1.Show

MsgBox "Login Successful.", vbInformation, "Successful Attempt"


End If

End Sub

Private Sub Label2_Click()

End Sub
