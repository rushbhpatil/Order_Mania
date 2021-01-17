VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00800080&
   Caption         =   "Recent Orders"
   ClientHeight    =   10170
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12675
   LinkTopic       =   "Form9"
   ScaleHeight     =   11055
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12630
      TabIndex        =   12
      Top             =   8070
      Width           =   2295
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12630
      TabIndex        =   11
      Top             =   6990
      Width           =   2295
   End
   Begin VB.CommandButton cmdready 
      Caption         =   "Ready"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12630
      TabIndex        =   10
      Top             =   5910
      Width           =   2295
   End
   Begin VB.ListBox ListItems 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Columns         =   2
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   4305
      Left            =   5280
      TabIndex        =   7
      Top             =   5355
      Width           =   3135
   End
   Begin VB.ListBox ListQty 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   4305
      Left            =   9195
      TabIndex        =   6
      Top             =   5355
      Width           =   975
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      BackColor       =   &H00C000C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   13020
      TabIndex        =   5
      Top             =   3165
      Width           =   2895
   End
   Begin VB.TextBox txtTableNo 
      Alignment       =   2  'Center
      BackColor       =   &H00C000C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   8880
      TabIndex        =   3
      Top             =   3150
      Width           =   2055
   End
   Begin VB.TextBox txtOrdNo 
      Alignment       =   2  'Center
      BackColor       =   &H00C000C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   5070
      TabIndex        =   2
      Top             =   3180
      Width           =   2055
   End
   Begin VB.Line Line6 
      X1              =   3795
      X2              =   16185
      Y1              =   2625
      Y2              =   2625
   End
   Begin VB.Line Line5 
      X1              =   10740
      X2              =   10740
      Y1              =   4275
      Y2              =   10155
   End
   Begin VB.Line Line4 
      X1              =   3795
      X2              =   16170
      Y1              =   5115
      Y2              =   5115
   End
   Begin VB.Line Line2 
      X1              =   8715
      X2              =   8715
      Y1              =   4275
      Y2              =   10155
   End
   Begin VB.Line Line1 
      X1              =   3795
      X2              =   16170
      Y1              =   4290
      Y2              =   4275
   End
   Begin VB.Shape Shape1 
      Height          =   9855
      Left            =   3795
      Top             =   315
      Width           =   12375
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   1095
      Left            =   8145
      Shape           =   4  'Rounded Rectangle
      Top             =   915
      Width           =   3855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
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
      ForeColor       =   &H00FFC0FF&
      Height          =   1455
      Left            =   5805
      TabIndex        =   13
      Top             =   1110
      Width           =   8535
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "Items"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6030
      TabIndex        =   9
      Top             =   4650
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "Qty"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9435
      TabIndex        =   8
      Top             =   4635
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11145
      TabIndex        =   4
      Top             =   3270
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "Order No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3630
      TabIndex        =   1
      Top             =   3315
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "Table Number"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7275
      TabIndex        =   0
      Top             =   3240
      Width           =   1455
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset

Private Sub cmdExit_Click()
Form8.Show
Unload Me
End Sub

Private Sub cmdNext_Click()
Dim ord As Integer
Set db = OpenDatabase("D:\OrderMania\ordermania.mdb")
Set rs = db.OpenRecordset("select * from order1")

While Not rs.EOF
If rs.Fields(8).Value = "Pending" And rs.Fields(9).Value = "Waiting" Then
ord = rs.Fields(0).Value
GoTo 1
End If
rs.MoveNext
Wend

1 While Not rs.EOF
If rs.Fields(8).Value = "Pending" And rs.Fields(0).Value = ord Then
      txtOrdNo.Text = rs.Fields(0).Value
      txtTableNo.Text = rs.Fields(1).Value
      txtName.Text = rs.Fields(7).Value
      ListItems.AddItem (rs.Fields(2).Value)
      ListQty.AddItem (rs.Fields(3).Value)
      
 End If
 rs.MoveNext
      
Wend

If txtOrdNo.Text = "" And txtName.Text = "" Then
MsgBox "Wait for The Next Order To be Placed", vbOKOnly, "No New Order Found"
Exit Sub
End If

End Sub

Private Sub cmdready_Click()
Set db = OpenDatabase("D:\OrderMania\ordermania.mdb")
Set rs = db.OpenRecordset("select * from order1")
While Not rs.EOF
If rs.Fields(7).Value = txtName.Text And rs.Fields(0).Value = txtOrdNo.Text Then
rs.Edit
rs.Fields(9).Value = "Ready"
rs.Update
End If
rs.MoveNext
Wend

txtOrdNo.Text = ""
txtTableNo.Text = ""
txtName.Text = ""
ListItems.Clear
ListQty.Clear

End Sub

Private Sub Form_Load()
txtOrdNo.Text = ""
txtTableNo.Text = ""
txtName.Text = ""
ListItems.Clear
ListQty.Clear
End Sub

Private Sub txtOrdNo_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtTableNo_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

