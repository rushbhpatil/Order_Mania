VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C000C0&
   Caption         =   "Welcome"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdmin 
      Height          =   1575
      Left            =   2760
      Picture         =   "pr1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   2295
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

Private Sub cmdlogin_Click()
loginado.RecordSource = "select * from Login where User_name='" + txtuser.Text + "' and Password='" + txtpass.Text + "'"
loginado.Refresh

If loginado.Recordset.EOF Then

MsgBox "Login failed,Try Again..!!!", vbCritical, "Please Enter correct Username and Password"
txtuser.Text = ""
txtpass.Text = ""
txtuser.SetFocus
Else
MsgBox "Login Successful.", vbInformation, "Successful Attempt"
End If

End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdAdmin_Click()
Form2.Show
Form1.Hide
End Sub
