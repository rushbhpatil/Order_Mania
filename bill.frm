VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Invoice"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17070
   ForeColor       =   &H00008080&
   LinkTopic       =   "Form6"
   ScaleHeight     =   9120
   ScaleWidth      =   17070
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """?"" #,##0.00;(""?"" #,##0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14220
      TabIndex        =   30
      Top             =   6180
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """?"" #,##0.00;(""?"" #,##0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14220
      TabIndex        =   28
      Top             =   5565
      Width           =   1215
   End
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
      Left            =   12855
      TabIndex        =   27
      Top             =   8805
      Width           =   2295
   End
   Begin VB.CommandButton cmdpaid 
      Caption         =   "Paid"
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
      Left            =   8910
      TabIndex        =   25
      Top             =   8805
      Width           =   2295
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
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
      Left            =   12990
      TabIndex        =   24
      Top             =   7245
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12975
      TabIndex        =   23
      Top             =   2145
      Width           =   2520
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """?"" #,##0.00;(""?"" #,##0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14190
      TabIndex        =   21
      Top             =   3780
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """?"" #,##0.00;(""?"" #,##0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14205
      TabIndex        =   19
      Top             =   4365
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """?"" #,##0.00;(""?"" #,##0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14220
      TabIndex        =   17
      Top             =   4905
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
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
      Height          =   510
      Left            =   4215
      TabIndex        =   15
      Top             =   2160
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Bill"
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
      Left            =   4695
      TabIndex        =   9
      Top             =   8805
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
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
      Height          =   495
      Left            =   10650
      TabIndex        =   7
      Top             =   2145
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Height          =   495
      Left            =   7785
      TabIndex        =   4
      Top             =   2145
      Width           =   2295
   End
   Begin VB.ListBox List4 
      Appearance      =   0  'Flat
      Height          =   4515
      Left            =   11235
      TabIndex        =   3
      Top             =   3555
      Width           =   855
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      Height          =   4515
      Left            =   9945
      TabIndex        =   2
      Top             =   3555
      Width           =   735
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   4515
      Left            =   8625
      TabIndex        =   1
      Top             =   3555
      Width           =   855
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   4515
      Left            =   4785
      TabIndex        =   0
      Top             =   3555
      Width           =   3135
   End
   Begin VB.Line Line8 
      X1              =   12300
      X2              =   15735
      Y1              =   5445
      Y2              =   5445
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12915
      TabIndex        =   31
      Top             =   6225
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Amount Paid"
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
      Left            =   12540
      TabIndex        =   29
      Top             =   5595
      Width           =   1455
   End
   Begin VB.Label lblstatus 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13095
      TabIndex        =   26
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Line Line7 
      X1              =   12315
      X2              =   12315
      Y1              =   2820
      Y2              =   8220
   End
   Begin VB.Line Line6 
      X1              =   3975
      X2              =   15735
      Y1              =   3405
      Y2              =   3405
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   735
      Left            =   8655
      Shape           =   4  'Rounded Rectangle
      Top             =   705
      Width           =   3135
   End
   Begin VB.Line Line5 
      X1              =   10920
      X2              =   10920
      Y1              =   2820
      Y2              =   8220
   End
   Begin VB.Line Line4 
      X1              =   9720
      X2              =   9720
      Y1              =   8220
      Y2              =   2820
   End
   Begin VB.Line Line3 
      X1              =   8280
      X2              =   8280
      Y1              =   8220
      Y2              =   2820
   End
   Begin VB.Line Line2 
      X1              =   3960
      X2              =   15720
      Y1              =   8220
      Y2              =   8220
   End
   Begin VB.Line Line1 
      X1              =   3960
      X2              =   15720
      Y1              =   2835
      Y2              =   2835
   End
   Begin VB.Shape Shape1 
      Height          =   9975
      Left            =   3975
      Top             =   285
      Width           =   11775
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Sub Total"
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
      Left            =   12660
      TabIndex        =   22
      Top             =   3855
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "GST (5%)"
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
      Left            =   12675
      TabIndex        =   20
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Grand Total"
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
      Left            =   12630
      TabIndex        =   18
      Top             =   4980
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Customer Name"
      Height          =   255
      Left            =   5040
      TabIndex        =   16
      Top             =   1740
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Amount"
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
      Left            =   11070
      TabIndex        =   14
      Top             =   2970
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
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
      Left            =   10035
      TabIndex        =   13
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Price"
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
      Left            =   8175
      TabIndex        =   12
      Top             =   2955
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
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
      Left            =   5430
      TabIndex        =   11
      Top             =   2970
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Time  And  Date"
      Height          =   255
      Left            =   13650
      TabIndex        =   10
      Top             =   1785
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
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
      ForeColor       =   &H000040C0&
      Height          =   855
      Left            =   6720
      TabIndex        =   8
      Top             =   780
      Width           =   6975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Table Number"
      Height          =   255
      Left            =   11010
      TabIndex        =   6
      Top             =   1785
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Order Number"
      Height          =   255
      Left            =   8145
      TabIndex        =   5
      Top             =   1785
      Width           =   1215
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset

Private Sub cmdExit_Click()
Form1.Show
Unload Me
End Sub

Private Sub cmdpaid_Click()
Dim db1 As Database
Dim rs1 As Recordset
Set db1 = OpenDatabase("D:\OrderMania\ordermania.mdb")
Set rs1 = db1.OpenRecordset("select * from order1")
While Not rs1.EOF
  If StrConv(Text3.Text, vbProperCase) = rs1.Fields(7).Value And Text1.Text = rs1.Fields(0).Value And rs1.Fields(8).Value = "Paid" Then
      MsgBox "The Bill is Already Paid", vbOKOnly
      Exit Sub
  End If
  If StrConv(Text3.Text, vbProperCase) = rs1.Fields(7).Value And Text1.Text = rs1.Fields(0).Value And rs1.Fields(8).Value = "Pending" Then
     rs1.Edit
     rs1.Fields(8).Value = "Paid"
     rs1.Update
     
     lblstatus.ForeColor = vbGreen
     lblstatus.Caption = "Paid"
  End If
  
  rs1.MoveNext
Wend

MsgBox "Payment Received", vbOKOnly
End Sub

Private Sub Command1_Click()
Dim tot As Integer
Dim gst, gtot As Double
tot = 0
Set db = OpenDatabase("D:\OrderMania\ordermania.mdb")
Set rs = db.OpenRecordset("select * from order1")


While Not rs.EOF
  If StrConv(Text3.Text, vbProperCase) = rs.Fields(7).Value And Text1.Text = rs.Fields(0).Value Then

      List1.AddItem (rs.Fields(2).Value)
      List2.AddItem (rs.Fields(4).Value)
      List3.AddItem (rs.Fields(3).Value)
      List4.AddItem (rs.Fields(5).Value)
      Text2.Text = rs.Fields(1).Value
      Text7.Text = rs.Fields(6).Value
      tot = tot + rs.Fields(5).Value
      lblstatus.Caption = rs.Fields(8).Value
    End If

If lblstatus.Caption = "Paid" Then
            lblstatus.ForeColor = vbGreen
            
Else
            lblstatus.ForeColor = vbRed
End If
rs.MoveNext
Wend
Text6.Text = tot
gst = tot * 0.05
Text5.Text = gst
gtot = tot + gst
Text4.Text = gtot


  
End Sub

Private Sub cmdclear_Click()
List1.Clear
List2.Clear
List3.Clear
List4.Clear
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
lblstatus.Caption = ""
End Sub

Private Sub Form_Load()
Text3.TabIndex = 0
Text1.TabIndex = 1
Command1.TabIndex = 2
List1.Clear
List2.Clear
List3.Clear
List4.Clear
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
lblstatus.Caption = ""
End Sub


Private Sub Text8_Change()
Dim bl As Double
bl = CDbl(Text8.Text) - CDbl(Text4.Text)
Text9.Text = bl
End Sub

