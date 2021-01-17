VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00C000C0&
   Caption         =   "Admin Login"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13050
   LinkTopic       =   "Form2"
   ScaleHeight     =   8820
   ScaleWidth      =   13050
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   15000
      TabIndex        =   9
      Top             =   6840
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   1680
      Top             =   7680
   End
   Begin VB.CommandButton cmdcreatenewuser 
      Appearance      =   0  'Flat
      Caption         =   "Command1"
      Height          =   735
      Left            =   8805
      Picture         =   "AdminLogin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9180
      Width           =   2415
   End
   Begin VB.TextBox txtuser 
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   3
      Top             =   5760
      Width           =   2895
   End
   Begin VB.TextBox txtpass 
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   8400
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   6840
      Width           =   2895
   End
   Begin VB.CommandButton cmdexit 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   10320
      MaskColor       =   &H00C000C0&
      Picture         =   "AdminLogin.frx":0E1F
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7800
      Width           =   2295
   End
   Begin VB.CommandButton cmdlogin 
      BackColor       =   &H00C000C0&
      Height          =   975
      Left            =   7335
      MaskColor       =   &H00C000C0&
      Picture         =   "AdminLogin.frx":CF70
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7800
      Width           =   2295
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   1095
      Left            =   7920
      Shape           =   4  'Rounded Rectangle
      Top             =   885
      Width           =   3855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C000C0&
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
      Left            =   5655
      TabIndex        =   10
      Top             =   1125
      Width           =   8535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   19410
      TabIndex        =   8
      Top             =   6855
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000007&
      Height          =   135
      Left            =   14655
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000007&
      Height          =   135
      Left            =   14415
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000007&
      Height          =   135
      Left            =   14175
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   12735
      TabIndex        =   7
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   8640
      Picture         =   "AdminLogin.frx":EA7F
      Top             =   2880
      Width           =   2250
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C000C0&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   5
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C000C0&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   4
      Top             =   6480
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset

Private Sub cmdcreatenewuser_Click()
Form5.Show
End Sub

Private Sub cmdExit_Click()
Form1.Show
Unload Me
End Sub

Private Sub cmdlogin_Click()
Dim u, t As String
Set db = OpenDatabase("D:\OrderMania\ordermania.mdb")
Set rs = db.OpenRecordset("select * from login")

If txtuser.Text = "" And txtpass.Text = "" Then
       MsgBox "Please Enter User Name and Password", vbCritical, "Alert!!"
       Exit Sub
ElseIf Len(txtpass.Text) < 8 Then
        MsgBox "Password Must Contain atleast 8 Characters", vbCritical, "Alert"
        txtpass.Text = ""
        txtpass.SetFocus
        Exit Sub
End If
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       rs.FindFirst "username='" & txtuser.Text & "'"
      
       If rs.EOF Then
             MsgBox "Incorrect username", vbOKOnly, "Alert!!"
             txtuser.Text = ""
             txtpass.Text = ""
             txtuser.SetFocus
             Exit Sub
       Else
          If rs.Fields(1).Value <> txtpass.Text Then
             MsgBox "Incorrect password", vbOKOnly, "Alert!!"
             txtuser.Text = ""
             txtpass.Text = ""
             txtuser.SetFocus
            Exit Sub
          End If
       End If
   End If

    Timer1.Enabled = True
    ProgressBar1.Visible = True
    Shape1.Visible = True
    Shape2.Visible = True
    Shape3.Visible = True
    Label3.Visible = True
    Label4.Visible = True
End Sub



Private Sub Timer1_Timer()
      If Shape1.Visible Then
         Shape2.Visible = True
         Shape1.Visible = False
         Shape3.Visible = False
      ElseIf Shape2.Visible Then
         Shape3.Visible = True
         Shape2.Visible = False
         Shape1.Visible = False
      ElseIf Shape3.Visible Then
         Shape1.Visible = True
         Shape2.Visible = False
         Shape3.Visible = False
      End If
    ProgressBar1.Value = ProgressBar1.Value + 5
    Label3.Caption = "Loading"
    Label4.Caption = ProgressBar1.Value & "%"
      If (ProgressBar1.Value = ProgressBar1.Max) Then
         Timer1.Enabled = False
         Form8.Show
         Form8.Label1.Caption = "User = " & StrConv(txtuser.Text, vbProperCase) & ""
         Unload Me
      End If
End Sub





Private Sub Form_Load()
txtuser.Text = ""
txtpass.Text = ""
txtuser.TabIndex = 0
txtpass.TabIndex = 1
cmdlogin.TabIndex = 2

End Sub

