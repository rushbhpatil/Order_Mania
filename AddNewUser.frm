VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00008080&
   Caption         =   "Create New User "
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12630
   LinkTopic       =   "Form3"
   ScaleHeight     =   11055
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00008080&
      Height          =   975
      Left            =   11145
      MaskColor       =   &H00008080&
      Picture         =   "AddNewUser.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6915
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.TextBox txtcpass 
      BorderStyle     =   0  'None
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
      Left            =   9825
      TabIndex        =   3
      Top             =   5595
      Width           =   3255
   End
   Begin VB.TextBox txtpass 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   9825
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   4635
      Width           =   3255
   End
   Begin VB.TextBox txtuser 
      BorderStyle     =   0  'None
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
      Left            =   9825
      TabIndex        =   1
      Top             =   3675
      Width           =   3255
   End
   Begin VB.CommandButton cmdsave 
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      Height          =   975
      Left            =   8505
      Picture         =   "AddNewUser.frx":0A07
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6915
      Width           =   2175
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
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   6285
      TabIndex        =   8
      Top             =   1680
      Width           =   8535
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   1095
      Left            =   8610
      Shape           =   4  'Rounded Rectangle
      Top             =   1455
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   5535
      Left            =   6915
      Shape           =   4  'Rounded Rectangle
      Top             =   2955
      Width           =   7335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00008080&
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7305
      TabIndex        =   6
      Top             =   5685
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00008080&
      Caption         =   "Enter Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7605
      TabIndex        =   5
      Top             =   4725
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008080&
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8175
      TabIndex        =   4
      Top             =   3765
      Width           =   1455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Form2.Show 'Admin Login
Unload Me
End Sub

Private Sub cmdsave_Click()

Dim db1 As Database
Dim rs1 As Recordset

Set db1 = OpenDatabase("D:\OrderMania\ordermania.mdb")
Set rs1 = db1.OpenRecordset("select * from login")

Set rs1 = db1.OpenRecordset("select * from login where username='" + txtuser.Text + "'")
If (Not rs1.EOF) Then
    MsgBox "Sorry!! User already exists. Try another username", vbCritical
    txtpass.Text = ""
    txtcpass.Text = ""
    txtuser.Text = ""
    txtuser.SetFocus
    Exit Sub
End If

If txtuser.Text = "" Then
MsgBox "Please Enter User Name", vbExclamation
txtuser.SetFocus
Exit Sub

ElseIf txtpass.Text = "" Then
MsgBox "Please Enter Password", vbExclamation
txtpass.SetFocus
Exit Sub

ElseIf txtcpass.Text = "" Then
MsgBox "Please Enter Confirm Password", vbExclamation
txtcpass.SetFocus
Exit Sub

ElseIf Len(txtcpass.Text) < 8 Then
        MsgBox "Confirm Password Must Contain atleast 8 Characters", vbCritical, "Alert"
        txtcpass.SetFocus
        Exit Sub
End If


If txtpass.Text = txtcpass.Text Then
 rs1.AddNew
 rs1.Fields(0).Value = txtuser.Text
 rs1.Fields(1).Value = txtpass.Text
 rs1.Update
 MsgBox "New User Created Successfully", vbOKOnly, "Success"
 Form2.Show
 Unload Me
Else
 MsgBox "Password Does not match", vbRetryCancel, "Alert!!"
 txtpass.Text = ""
 txtcpass.Text = ""
 txtpass.SetFocus
End If


End Sub

Private Sub txtpass_LostFocus()
If Len(txtpass.Text) < 8 Then
        MsgBox "Password Must Contain atleast 8 Characters", vbCritical, "Alert"
        txtpass.SetFocus
        Exit Sub
End If
End Sub
