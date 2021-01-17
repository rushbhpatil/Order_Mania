VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Invoice"
   ClientHeight    =   9780
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17070
   ForeColor       =   &H00008080&
   LinkTopic       =   "Form6"
   ScaleHeight     =   9780
   ScaleWidth      =   17070
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   6855
      Left            =   13320
      TabIndex        =   35
      Top             =   1920
      Width           =   6855
      Begin VB.CommandButton cmdApply 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Apply"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5235
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1215
         Width           =   1215
      End
      Begin VB.CommandButton CmdGenerate 
         BackColor       =   &H00E0E0E0&
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
         Left            =   915
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   4095
         Width           =   2295
      End
      Begin VB.CommandButton cmdclear 
         BackColor       =   &H00E0E0E0&
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
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2820
         Width           =   2295
      End
      Begin VB.CommandButton cmdpaid 
         BackColor       =   &H00E0E0E0&
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
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   4110
         Width           =   2295
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00E0E0E0&
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
         Height          =   900
         Left            =   3810
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   5370
         Width           =   2295
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Print"
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
         Left            =   930
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   5400
         Width           =   2295
      End
      Begin VB.ComboBox ComboDiscount 
         Height          =   315
         Left            =   2025
         TabIndex        =   36
         Text            =   "Combo1"
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Discount Type"
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
         Left            =   225
         TabIndex        =   43
         Top             =   1320
         Width           =   1695
      End
   End
   Begin VB.TextBox txtDiscount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   10380
      TabIndex        =   31
      Top             =   5100
      Width           =   1215
   End
   Begin VB.ListBox ListItems 
      Appearance      =   0  'Flat
      Height          =   4710
      Left            =   855
      TabIndex        =   27
      Top             =   5085
      Width           =   3015
   End
   Begin VB.TextBox txtChange 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   10410
      TabIndex        =   25
      Top             =   7605
      Width           =   1215
   End
   Begin VB.TextBox txtAmtPaid 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   10395
      TabIndex        =   23
      Top             =   6975
      Width           =   1215
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   9165
      TabIndex        =   21
      Top             =   3090
      Width           =   2520
   End
   Begin VB.TextBox txtSubTot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   10365
      TabIndex        =   19
      Top             =   4575
      Width           =   1215
   End
   Begin VB.TextBox txtGST 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   10395
      TabIndex        =   17
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox txtGtot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   10395
      TabIndex        =   15
      Top             =   6225
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Left            =   270
      TabIndex        =   13
      Top             =   3120
      Width           =   3375
   End
   Begin VB.TextBox txtTableNo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Left            =   6615
      TabIndex        =   6
      Top             =   3105
      Width           =   2175
   End
   Begin VB.TextBox txtOrdNo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Left            =   3990
      TabIndex        =   3
      Top             =   3120
      Width           =   2295
   End
   Begin VB.ListBox ListAmt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   4710
      Left            =   7410
      TabIndex        =   2
      Top             =   5085
      Width           =   855
   End
   Begin VB.ListBox ListQty 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   4710
      Left            =   6135
      TabIndex        =   1
      Top             =   5100
      Width           =   735
   End
   Begin VB.ListBox ListPrice 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   4710
      Left            =   4710
      TabIndex        =   0
      Top             =   5100
      Width           =   855
   End
   Begin VB.Label Label21 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10110
      TabIndex        =   34
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Caption         =   "Time:"
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
      Left            =   9315
      TabIndex        =   33
      Top             =   3975
      Width           =   615
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Caption         =   "Discount"
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
      Left            =   8895
      TabIndex        =   32
      Top             =   5130
      Width           =   1215
   End
   Begin VB.Image ImageStamp 
      Enabled         =   0   'False
      Height          =   1965
      Left            =   9210
      Picture         =   "invoice.frx":0000
      Stretch         =   -1  'True
      Top             =   8505
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Line Line4 
      X1              =   5925
      X2              =   5925
      Y1              =   3780
      Y2              =   10515
   End
   Begin VB.Line Line3 
      X1              =   4320
      X2              =   4320
      Y1              =   3795
      Y2              =   10515
   End
   Begin VB.Line Line2 
      X1              =   165
      X2              =   11925
      Y1              =   2535
      Y2              =   2535
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
      Left            =   4515
      TabIndex        =   30
      Top             =   1275
      Width           =   3255
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
      Left            =   4515
      TabIndex        =   29
      Top             =   1635
      Width           =   3255
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
      Left            =   4215
      TabIndex        =   28
      Top             =   2010
      Width           =   3855
   End
   Begin VB.Line Line8 
      X1              =   8505
      X2              =   11940
      Y1              =   6735
      Y2              =   6735
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
      Left            =   9105
      TabIndex        =   26
      Top             =   7650
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
      Left            =   8730
      TabIndex        =   24
      Top             =   7020
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
      Left            =   9285
      TabIndex        =   22
      Top             =   8115
      Width           =   1935
   End
   Begin VB.Line Line7 
      X1              =   8505
      X2              =   8505
      Y1              =   3765
      Y2              =   10515
   End
   Begin VB.Line Line6 
      X1              =   165
      X2              =   11925
      Y1              =   4350
      Y2              =   4350
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   855
      Left            =   4470
      Shape           =   4  'Rounded Rectangle
      Top             =   300
      Width           =   3255
   End
   Begin VB.Line Line5 
      X1              =   7110
      X2              =   7110
      Y1              =   3795
      Y2              =   10560
   End
   Begin VB.Line Line1 
      X1              =   150
      X2              =   11910
      Y1              =   3780
      Y2              =   3780
   End
   Begin VB.Shape Shape1 
      Height          =   10485
      Left            =   150
      Top             =   45
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
      Left            =   8835
      TabIndex        =   20
      Top             =   4650
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
      Left            =   8850
      TabIndex        =   18
      Top             =   5700
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
      Left            =   8790
      TabIndex        =   16
      Top             =   6255
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
      Left            =   1230
      TabIndex        =   14
      Top             =   2685
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
      Left            =   7260
      TabIndex        =   12
      Top             =   3915
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
      Left            =   6225
      TabIndex        =   11
      Top             =   3945
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
      Left            =   4320
      TabIndex        =   10
      Top             =   3975
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
      Left            =   1590
      TabIndex        =   9
      Top             =   3945
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
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
      Left            =   9795
      TabIndex        =   8
      Top             =   2745
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
      Left            =   4560
      TabIndex        =   7
      Top             =   405
      Width           =   3135
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
      Height          =   255
      Left            =   6975
      TabIndex        =   5
      Top             =   2745
      Width           =   1335
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
      Left            =   4575
      TabIndex        =   4
      Top             =   2745
      Width           =   1215
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tot, Tot2, gst, gtot, dis, b1 As Double

Private Sub cmdApply_Click()
If ComboDiscount.Text = "Birthday" And Val(txtSubTot.Text) < 200 Then
    MsgBox "The Order Value Should be Atleast Rs.200/- to Avail This Offer", vbInformation, "Not Applicable"
    ComboDiscount.Text = "N/A"
    Exit Sub
ElseIf ComboDiscount.Text = "Wedding Anniversary" And Val(txtSubTot.Text) < 300 Then
    MsgBox "The Order Value Should be Atleast Rs.300/- to Avail This Offer", vbInformation, "Not Applicable"
    ComboDiscount.Text = "N/A"
    Exit Sub

ElseIf ComboDiscount.Text = "N/A" Then
txtDiscount.Text = 0

ElseIf ComboDiscount.Text = "Birthday" Then
txtDiscount.Text = Tot * 0.2

ElseIf ComboDiscount.Text = "Wedding Anniversary" Then
txtDiscount.Text = Tot * 0.25
End If

Tot2 = Tot - CDbl(txtDiscount.Text)
gst = Tot2 * 0.05
txtGST.Text = gst
gtot = Tot2 + gst
txtGtot.Text = gtot
End Sub

Private Sub cmdExit_Click()
Form8.Show
Unload Me
End Sub

Private Sub cmdpaid_Click()
Dim db1 As Database
Dim rs1 As Recordset
Dim rsc As Recordset
Dim flag As Integer
flag = 0
Set db1 = OpenDatabase("D:\OrderMania\ordermania.mdb")
Set rs1 = db1.OpenRecordset("select * from order1")
Set rsc = db1.OpenRecordset("select * from common")
While Not rs1.EOF
   
  If StrConv(txtName.Text, vbProperCase) = rs1.Fields(7).Value And txtOrdNo.Text = rs1.Fields(0).Value And rs1.Fields(8).Value = "Paid" Then
      MsgBox "The Bill is Already Paid", vbOKOnly
      Call allclear
      Exit Sub
  End If
  If StrConv(txtName.Text, vbProperCase) = rs1.Fields(7).Value And txtOrdNo.Text = rs1.Fields(0).Value And rs1.Fields(8).Value = "Pending" Then
     rs1.Edit
     rs1.Fields(8).Value = "Paid"
     rs1.Update
     flag = 1
     lblstatus.ForeColor = vbGreen
     lblstatus.Caption = "Paid"
     
  End If
  
  rs1.MoveNext
Wend

While Not rsc.EOF
      If rsc.Fields(1).Value = StrConv(txtName.Text, vbProperCase) And rsc.Fields(0).Value = txtOrdNo.Text And flag = 1 Then
        rsc.Edit
        rsc.Fields(2).Value = txtDiscount.Text
        rsc.Fields(3).Value = txtGtot.Text
        rsc.Update
       End If
     rsc.MoveNext
  Wend

ImageStamp.Visible = True
MsgBox "Payment Received", vbOKOnly
End Sub

Private Sub cmdPrint_Click()
 Form6.PrintForm

End Sub

Private Sub CmdGenerate_Click()
Dim dbi As Database
Dim rsi As Recordset
Tot = 0
Set dbi = OpenDatabase("D:\OrderMania\ordermania.mdb")
Set rsi = dbi.OpenRecordset("select * from order1")

  While Not rsi.EOF
   If StrConv(txtName.Text, vbProperCase) = rsi.Fields(7).Value And txtOrdNo.Text = rsi.Fields(0).Value And rsi.Fields(9).Value = "Waiting" And rsi.Fields(8).Value = "Pending" Then
      MsgBox "Customer Yet to be Served", vbCritical, "Order Waiting"
      Exit Sub
   ElseIf StrConv(txtName.Text, vbProperCase) = rsi.Fields(7).Value And txtOrdNo.Text = rsi.Fields(0).Value Then

      ListItems.AddItem (rsi.Fields(2).Value)
      ListPrice.AddItem (rsi.Fields(4).Value)
      ListQty.AddItem (rsi.Fields(3).Value)
      ListAmt.AddItem (rsi.Fields(5).Value)
      txtTableNo.Text = rsi.Fields(1).Value
      txtDate.Text = rsi.Fields(6).Value
      Tot = Tot + rsi.Fields(5).Value
      lblstatus.Caption = rsi.Fields(8).Value
    End If

If lblstatus.Caption = "Paid" Then
            lblstatus.ForeColor = vbGreen
            
Else
            lblstatus.ForeColor = vbRed
End If
rsi.MoveNext
Wend

If ComboDiscount.Text = "N/A" Then
dis = 0
End If
txtDiscount.Text = dis

txtSubTot.Text = Tot
Tot2 = Tot - dis
gst = Tot2 * 0.05
txtGST.Text = gst
gtot = Tot + gst
txtGtot.Text = gtot

ImageStamp.Enabled = True

End Sub

Private Sub cmdclear_Click()
Call allclear
txtName.SetFocus
End Sub


Private Sub Form_Load()
b1 = 0
txtName.TabIndex = 0
txtOrdNo.TabIndex = 1
CmdGenerate.TabIndex = 2
Label21.Caption = Time
Call allclear
ComboDiscount.AddItem "N/A"
ComboDiscount.AddItem "Birthday"
ComboDiscount.AddItem "Wedding Anniversary"
txtGST.Text = Format(txtGST.Text, "#.##")
txtGtot.Text = Format(txtGtot.Text, "#.##")
txtDiscount.Text = Format(txtDiscount.Text, "#.##")
txtChange.Text = Format(txtChange.Text, "#.##")

End Sub



Private Sub txtOrdNo_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyRight Or kayascii = vbKeyLeft Then
Else
KeyAscii = 0
End If
End Sub
Private Sub txtName_KeyPress(KeyAscii As Integer)
Dim ch As String

    ch = Chr$(KeyAscii)
    If Not ((ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyRight Or kayascii = vbKeyLeft) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtAmtPaid_Change()
txtChange.Text = CDbl(txtAmtPaid.Text) - CDbl(txtGtot.Text)
End Sub

Private Sub allclear()
ListItems.Clear
ListPrice.Clear
ListQty.Clear
ListAmt.Clear
txtOrdNo.Text = ""
txtTableNo.Text = ""
txtName.Text = ""
txtGtot.Text = ""
txtGST.Text = ""
txtSubTot.Text = ""
txtDate.Text = ""
txtAmtPaid.Text = ""
txtChange.Text = ""
txtDiscount.Text = ""
lblstatus.Caption = ""
ComboDiscount.Text = "N/A"
Tot = 0
Tot2 = 0
gst = 0
gtot = 0
dis = 0
End Sub

Private Sub txtTableNo_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtGtot_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub txtGST_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub txtSubTot_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
