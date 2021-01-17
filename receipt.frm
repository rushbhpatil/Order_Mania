VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Order Receipt"
   ClientHeight    =   9540
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15450
   LinkTopic       =   "Form7"
   ScaleHeight     =   9540
   ScaleWidth      =   15450
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H80000000&
      Caption         =   "Print "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4680
      Width           =   2535
   End
   Begin VB.ListBox ListItems 
      Appearance      =   0  'Flat
      Height          =   4515
      Left            =   930
      TabIndex        =   11
      Top             =   4605
      Width           =   3135
   End
   Begin VB.ListBox ListPrice 
      Appearance      =   0  'Flat
      Height          =   4515
      Left            =   4770
      TabIndex        =   10
      Top             =   4605
      Width           =   855
   End
   Begin VB.ListBox ListQty 
      Appearance      =   0  'Flat
      Height          =   4515
      Left            =   6090
      TabIndex        =   9
      Top             =   4605
      Width           =   735
   End
   Begin VB.ListBox ListAmt 
      Appearance      =   0  'Flat
      Height          =   4515
      Left            =   7380
      TabIndex        =   8
      Top             =   4605
      Width           =   855
   End
   Begin VB.TextBox txtOrdNo 
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
      Left            =   3960
      TabIndex        =   7
      Top             =   3195
      Width           =   2295
   End
   Begin VB.TextBox txtTableNo 
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
      Left            =   6570
      TabIndex        =   6
      Top             =   3180
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000000&
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox txtName 
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
      Left            =   300
      TabIndex        =   4
      Top             =   3195
      Width           =   3375
   End
   Begin VB.TextBox txtGtot 
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
      Left            =   10200
      TabIndex        =   3
      Top             =   6990
      Width           =   1200
   End
   Begin VB.TextBox txtGST 
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
      Left            =   10215
      TabIndex        =   2
      Top             =   6375
      Width           =   1215
   End
   Begin VB.TextBox txtSub 
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
      Left            =   10200
      TabIndex        =   1
      Top             =   5775
      Width           =   1215
   End
   Begin VB.TextBox txtDate 
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
      Height          =   450
      Left            =   9000
      TabIndex        =   0
      Top             =   3195
      Width           =   1680
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "Before Leaving This Page"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   16335
      TabIndex        =   31
      Top             =   5115
      Width           =   2655
   End
   Begin VB.Label Label18 
      Caption         =   "Please Take a print out of the receipt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   16380
      TabIndex        =   30
      Top             =   4830
      Width           =   2655
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "Contact: 9096339183 / 8446318988"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4410
      TabIndex        =   29
      Top             =   2085
      Width           =   3855
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "Viman Nagar,Pune 411014"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4710
      TabIndex        =   28
      Top             =   1710
      Width           =   3255
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "221B Baker Street,"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4710
      TabIndex        =   27
      Top             =   1350
      Width           =   3255
   End
   Begin VB.Label Label14 
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10965
      TabIndex        =   26
      Top             =   2820
      Width           =   495
   End
   Begin VB.Label Label5 
      Height          =   300
      Left            =   10785
      TabIndex        =   25
      Top             =   3255
      Width           =   855
   End
   Begin VB.Line Line8 
      X1              =   8460
      X2              =   8460
      Y1              =   3885
      Y2              =   9480
   End
   Begin VB.Line Line7 
      X1              =   105
      X2              =   11865
      Y1              =   4425
      Y2              =   4425
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   735
      Left            =   4740
      Shape           =   4  'Rounded Rectangle
      Top             =   435
      Width           =   3135
   End
   Begin VB.Line Line6 
      X1              =   105
      X2              =   11865
      Y1              =   2625
      Y2              =   2625
   End
   Begin VB.Line Line1 
      X1              =   105
      X2              =   11865
      Y1              =   3870
      Y2              =   3870
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Order Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4395
      TabIndex        =   23
      Top             =   2820
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Table Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7020
      TabIndex        =   22
      Top             =   2820
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
      Height          =   615
      Left            =   4800
      TabIndex        =   21
      Top             =   495
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9675
      TabIndex        =   20
      Top             =   2835
      Width           =   495
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
      Left            =   1575
      TabIndex        =   19
      Top             =   4020
      Width           =   1695
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
      Left            =   4635
      TabIndex        =   18
      Top             =   4005
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
      Left            =   6180
      TabIndex        =   17
      Top             =   4035
      Width           =   495
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
      Left            =   7155
      TabIndex        =   16
      Top             =   4020
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1185
      TabIndex        =   15
      Top             =   2790
      Width           =   1575
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
      Left            =   8655
      TabIndex        =   14
      Top             =   7020
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
      Left            =   8685
      TabIndex        =   13
      Top             =   6435
      Width           =   1215
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
      Left            =   8670
      TabIndex        =   12
      Top             =   5805
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   9375
      Left            =   120
      Top             =   120
      Width           =   11775
   End
   Begin VB.Line Line3 
      X1              =   4425
      X2              =   4425
      Y1              =   9480
      Y2              =   3870
   End
   Begin VB.Line Line4 
      X1              =   5865
      X2              =   5865
      Y1              =   9480
      Y2              =   3870
   End
   Begin VB.Line Line5 
      X1              =   7065
      X2              =   7065
      Y1              =   3870
      Y2              =   9480
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdPrint_Click()
Form7.PrintForm
Form1.Show
Unload Me
End Sub

Private Sub Command1_Click()
Dim a As String
MsgBox "It's Our Pleasure to Serve You, Your order will be on your table shortly", vbOKOnly, "Thank You"
Form1.Show
Unload Me
End Sub


Private Sub Form_Load()
ListItems.Clear
ListPrice.Clear
ListQty.Clear
ListAmt.Clear
txtOrdNo.Text = ""
txtTableNo.Text = ""
txtName.Text = ""
txtGtot.Text = ""
Label5.Caption = Time
txtGST.Text = Format(txtGST.Text, "#.##")
txtGtot.Text = Format(txtGtot.Text, "#.##")

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
Private Sub txtGtot_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub txtSub_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub txtGST_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub txtDate_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
