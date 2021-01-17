VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form13 
   BackColor       =   &H00C000C0&
   Caption         =   "Restaurant Survey Form"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   19140
   LinkTopic       =   "Form13"
   ScaleHeight     =   9135
   ScaleWidth      =   19140
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00800080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   75
      TabIndex        =   2
      Top             =   3150
      Width           =   20250
      Begin VB.Frame Frame2 
         BackColor       =   &H00800080&
         Height          =   735
         Left            =   16920
         TabIndex        =   32
         Top             =   1875
         Width           =   2175
         Begin VB.OptionButton Option2 
            BackColor       =   &H00800080&
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1320
            TabIndex        =   34
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00800080&
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   375
            Left            =   240
            TabIndex        =   33
            Top             =   240
            Width           =   855
         End
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   1560
         Left            =   11760
         TabIndex        =   26
         Top             =   5145
         Width           =   7560
         _ExtentX        =   13335
         _ExtentY        =   2752
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"feedback.frx":0000
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H000080FF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   14325
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   6945
         Width           =   1575
      End
      Begin VB.CommandButton cmdSubmit 
         BackColor       =   &H000080FF&
         Caption         =   "Submit"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11925
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   6960
         Width           =   1575
      End
      Begin VB.TextBox txtMobileNo 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   15060
         TabIndex        =   20
         Top             =   495
         Width           =   2295
      End
      Begin VB.TextBox txtName 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8040
         TabIndex        =   18
         Top             =   480
         Width           =   3375
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   6165
         TabIndex        =   14
         Top             =   6960
         Width           =   2295
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   6165
         TabIndex        =   12
         Top             =   6030
         Width           =   2295
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   6150
         TabIndex        =   10
         Top             =   5085
         Width           =   2295
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   6150
         TabIndex        =   8
         Top             =   4020
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   6135
         TabIndex        =   6
         Top             =   2925
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   6165
         TabIndex        =   4
         Top             =   1905
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   121503745
         CurrentDate     =   43738
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Excellent"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   16950
         TabIndex        =   31
         Top             =   4560
         Width           =   1080
      End
      Begin VB.Image Excellent 
         Height          =   825
         Left            =   17040
         Picture         =   "feedback.frx":00BD
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   870
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Good"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   15750
         TabIndex        =   30
         Top             =   4560
         Width           =   1080
      End
      Begin VB.Image Good 
         Height          =   825
         Left            =   15825
         Picture         =   "feedback.frx":BD39
         Stretch         =   -1  'True
         Top             =   3585
         Width           =   870
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Average"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   14520
         TabIndex        =   29
         Top             =   4560
         Width           =   1080
      End
      Begin VB.Image Average 
         Height          =   825
         Left            =   14610
         Picture         =   "feedback.frx":19DD6
         Stretch         =   -1  'True
         Top             =   3585
         Width           =   870
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Poor"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   255
         Left            =   13695
         TabIndex        =   28
         Top             =   4560
         Width           =   615
      End
      Begin VB.Image Poor 
         Height          =   825
         Left            =   13530
         Picture         =   "feedback.frx":224A4
         Stretch         =   -1  'True
         Top             =   3585
         Width           =   870
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Dissatisfied"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   12195
         TabIndex        =   27
         Top             =   4575
         Width           =   1215
      End
      Begin VB.Image Dissatisfied 
         Height          =   825
         Left            =   12390
         Picture         =   "feedback.frx":30555
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   870
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00FF00FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile Number"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   12525
         TabIndex        =   19
         Top             =   525
         Width           =   2415
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00FF00FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   5475
         TabIndex        =   17
         Top             =   495
         Width           =   2415
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00FF00FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Overall Experience"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   375
         Left            =   13860
         TabIndex        =   16
         Top             =   2940
         Width           =   2775
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00FF00FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Washrooms properly clean and maintained?"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   12045
         TabIndex        =   15
         Top             =   1995
         Width           =   4215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FF00FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Speed of service"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3495
         TabIndex        =   13
         Top             =   6960
         Width           =   2055
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FF00FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Cleanliness"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3945
         TabIndex        =   11
         Top             =   6000
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FF00FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Food Quality"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3885
         TabIndex        =   9
         Top             =   5085
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FF00FF&
         BackStyle       =   0  'Transparent
         Caption         =   "How is the dinning room(s), decor, lightning, music and overall ambience?"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   1260
         TabIndex        =   7
         Top             =   3825
         Width           =   4575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FF00FF&
         BackStyle       =   0  'Transparent
         Caption         =   "What brought you in today?"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2145
         TabIndex        =   5
         Top             =   2880
         Width           =   3615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FF00FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Visited Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4035
         TabIndex        =   3
         Top             =   1890
         Width           =   1695
      End
   End
   Begin VB.Label Label15 
      Caption         =   "Label15"
      Height          =   495
      Left            =   9600
      TabIndex        =   25
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      Height          =   495
      Left            =   9600
      TabIndex        =   24
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "We'd Love to Hear From You"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   435
      Left            =   8190
      TabIndex        =   23
      Top             =   2565
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Restaurant Survey Form"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   7800
      TabIndex        =   1
      Top             =   1755
      Width           =   4815
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   1095
      Left            =   8295
      Shape           =   4  'Rounded Rectangle
      Top             =   375
      Width           =   3855
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
      ForeColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   5970
      TabIndex        =   0
      Top             =   585
      Width           =   8535
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag, flag1 As Integer

Private Sub Average_Click()
flag1 = 3
Average.BorderStyle = 1
End Sub

Private Sub cmdCancel_Click()
Form1.Show
Unload Me
End Sub

Private Sub cmdSubmit_Click()
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase("D:\OrderMania\ordermania.mdb")
Set rs = db.OpenRecordset("select * from feedback")


 rs.AddNew
 rs.Fields(0).Value = StrConv(txtName.Text, vbProperCase)
 rs.Fields(1).Value = txtMobileNo.Text
 rs.Fields(2).Value = DTPicker1.Value
 rs.Fields(3).Value = Combo1.Text
 rs.Fields(4).Value = Combo2.Text
 rs.Fields(5).Value = Combo3.Text
 rs.Fields(6).Value = Combo4.Text
 rs.Fields(7).Value = Combo5.Text
 If RichTextBox1.Text = " Is there any scope for improvements? Please Share your points" Then
 rs.Fields(10).Value = "Nothing"
 Else
 rs.Fields(10).Value = RichTextBox1.Text
 End If
 If flag = 0 Then
 rs.Fields(8).Value = Option1.Caption
 Else
 rs.Fields(8).Value = Option2.Caption
 End If
 
 If flag1 = 1 Then
 rs.Fields(9).Value = "Dissatisfied"
 ElseIf flag1 = 2 Then
 rs.Fields(9).Value = "Poor"
 ElseIf flag1 = 3 Then
 rs.Fields(9).Value = "Average"
 ElseIf flag1 = 4 Then
 rs.Fields(9).Value = "Good"
 ElseIf flag1 = 5 Then
 rs.Fields(9).Value = "Excellent"
 End If
 
 If txtName.Text = "" Then
  MsgBox "Please Enter Name", vbCritical, "oops!!"
  txtName.SetFocus
  Exit Sub
 ElseIf txtMobileNo.Text = "" Then
  MsgBox "Please Enter Mobile Number", vbCritical, "oops!!"
  txtMobileNo.SetFocus
  Exit Sub
 ElseIf Len(txtMobileNo.Text) < 10 Then
  MsgBox "Please Enter Valid Mobile Number", vbCritical, "oops!!"
  txtMobileNo.SetFocus
  Exit Sub
 ElseIf IsNull(DTPicker1.Value) Then
  MsgBox "Please Select a Date", vbCritical, "oops!!"
  Exit Sub
 ElseIf Combo1.Text = "Select" Or Combo2.Text = "Select" Or Combo3.Text = "Select" Or Combo4.Text = "Select" Or Combo5.Text = "Select" Then
   MsgBox "Please Select a Option", vbCritical, "oops!!"
   Exit Sub
 ElseIf Option1.Value = False And Option2.Value = False Then
   MsgBox "Please Select a Option(Yes/No)", vbCritical, "oops!!"
   Exit Sub
 ElseIf Excellent.BorderStyle = 0 And Poor.BorderStyle = 0 And Good.BorderStyle = 0 And Average.BorderStyle = 0 And Dissatisfied.BorderStyle = 0 Then
   MsgBox "Please Select a Overall Experience", vbCritical, "oops!!"
   Exit Sub
 End If

 rs.Update
MsgBox "Your feedback is important to us. We value and appreciate receiving your compliments, suggestions or complaints,It will help us to Improve our Services", vbInformation, "Thank you"
Call allclear
Form1.Show
Unload Me
End Sub

Private Sub Combo1_Change()
Combo1.AddItem "Location"
Combo1.AddItem "Social Media"
Combo1.AddItem "Advertisment"
Combo1.AddItem "Recommendation"
Combo1.AddItem "Repeat Customer"
Combo1.AddItem "Other"
End Sub

Private Sub Combo2_Change()
Combo2.AddItem "Excellent"
Combo2.AddItem "Good"
Combo2.AddItem "Average"
Combo2.AddItem "Poor"
Combo2.AddItem "Dissatisfied"
End Sub

Private Sub Combo3_Change()
Combo3.AddItem "Excellent"
Combo3.AddItem "Good"
Combo3.AddItem "Average"
Combo3.AddItem "Poor"
Combo3.AddItem "Dissatisfied"
End Sub

Private Sub Combo4_Change()
Combo4.AddItem "Excellent"
Combo4.AddItem "Good"
Combo4.AddItem "Average"
Combo4.AddItem "Poor"
Combo4.AddItem "Dissatisfied"
End Sub

Private Sub Combo5_Change()
Combo5.AddItem "Excellent"
Combo5.AddItem "Good"
Combo5.AddItem "Average"
Combo5.AddItem "Poor"
Combo5.AddItem "Dissatisfied"
End Sub

Private Sub Dissatisfied_Click()
flag1 = 1
Dissatisfied.BorderStyle = 1
End Sub

Private Sub DTPicker1_lostfocus()
If Format(DTPicker1.Value, "mm-dd-yyyy") > Format(Now, "mm-dd-yyyy") Then
MsgBox ("The Visit Should not be Larger than Todays Date"), vbExclamation, "Please Select Appropriate Date"
DTPicker1.Value = ""
Exit Sub

End If
End Sub

Private Sub Excellent_Click()
flag1 = 5
Excellent.BorderStyle = 1
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
Dim ch As String

    ch = Chr$(KeyAscii)
    If Not ((ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyRight Or kayascii = vbKeyLeft) Then
        KeyAscii = 0
    End If
End Sub
Private Sub Form_Load()
Call allclear
End Sub

Private Sub Good_Click()
flag1 = 4
Good.BorderStyle = 1
End Sub

Private Sub Option1_Click()
flag = 0
End Sub

Private Sub Option2_Click()
flag = 1
End Sub

Private Sub Poor_Click()
flag1 = 2
Poor.BorderStyle = 1
End Sub

Private Sub RichTextBox1_Change()
RichTextBox1.MaxLength = 255
End Sub

Private Sub RichTextBox1_GotFocus()
RichTextBox1.Text = ""
End Sub


Private Sub txtMobileNo_GotFocus()
txtMobileNo.MaxLength = 10
End Sub

Private Sub txtMobileNo_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyRight Or kayascii = vbKeyLeft Then
Else
KeyAscii = 0
End If

End Sub

Private Sub allclear()
txtName.Text = ""
txtMobileNo.Text = ""
DTPicker1.Value = ""
RichTextBox1.Text = " Is there any scope for improvements? Please Share your points"
Option1.Value = False
Option2.Value = False
Combo1.Text = "Select"
Combo2.Text = "Select"
Combo3.Text = "Select"
Combo4.Text = "Select"
Combo5.Text = "Select"
Excellent.BorderStyle = 0
Poor.BorderStyle = 0
Good.BorderStyle = 0
Average.BorderStyle = 0
Dissatisfied.BorderStyle = 0
End Sub
