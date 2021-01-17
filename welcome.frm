VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form11 
   Caption         =   "Welcome"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13740
   LinkTopic       =   "Form11"
   ScaleHeight     =   7845
   ScaleWidth      =   13740
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   4680
      TabIndex        =   0
      Top             =   10560
      Visible         =   0   'False
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   375
      Left            =   15840
      TabIndex        =   1
      Top             =   10440
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000007&
      Height          =   135
      Left            =   9270
      Shape           =   4  'Rounded Rectangle
      Top             =   10170
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000007&
      Height          =   135
      Left            =   9915
      Shape           =   4  'Rounded Rectangle
      Top             =   10170
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000007&
      Height          =   135
      Left            =   10560
      Shape           =   4  'Rounded Rectangle
      Top             =   10170
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   11040
      Left            =   0
      Picture         =   "welcome.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20400
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Timer1.Enabled = True
    ProgressBar1.Visible = True
     Shape1.Visible = True
     Shape2.Visible = True
    Shape3.Visible = True
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
    Label1.Caption = ProgressBar1.Value & "%"
      If (ProgressBar1.Value = ProgressBar1.Max) Then
         Timer1.Enabled = False
         Form1.Show
         Unload Me
      End If
     
End Sub
