VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00004080&
   Caption         =   "Master Password"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4875
   LinkTopic       =   "Form5"
   ScaleHeight     =   4845
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00404080&
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      MaskColor       =   &H00404080&
      TabIndex        =   1
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox txtmpass 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   4335
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   330
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      Caption         =   "Enter Master Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   720
      Width           =   3375
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Set db = OpenDatabase("D:\OrderMania\ordermania.mdb")
Set rs = db.OpenRecordset("select * from login")

If txtmpass.Text = "" Then
       MsgBox "Please Enter Master Password"
       txtmpass.SetFocus
       Exit Sub
End If
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       rs.FindFirst "username='master'"
          If rs.Fields(0).Value = "master" And rs.Fields(1).Value <> txtmpass.Text Then
             MsgBox "Incorrect password", vbOKOnly
             txtmpass.Text = ""
             txtmpass.SetFocus
            Exit Sub
          Else
             MsgBox "Master Password Is Correct", vbOKOnly
             Form3.Show
             Unload Me
             
          End If
    End If
End Sub

Private Sub Command2_Click()
Form2.Show
Unload Me
End Sub

