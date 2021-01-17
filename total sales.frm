VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form10 
   BackColor       =   &H00404000&
   Caption         =   "Total Sales"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16395
   LinkTopic       =   "Form10"
   ScaleHeight     =   9165
   ScaleWidth      =   16395
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtDiscount 
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
      Left            =   18135
      TabIndex        =   15
      Top             =   8580
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00808080&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10665
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3510
      Width           =   1575
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
      Left            =   18120
      TabIndex        =   11
      Top             =   7950
      Width           =   1215
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
      Left            =   18135
      TabIndex        =   8
      Top             =   9870
      Width           =   1215
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
      Left            =   18135
      TabIndex        =   7
      Top             =   9240
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   1440
      Top             =   3960
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   100
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\OrderMania\ordermania.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\OrderMania\ordermania.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from order1"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "total sales.frx":0000
      Height          =   5415
      Left            =   2520
      TabIndex        =   3
      Top             =   5520
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   9551
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00808080&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3525
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   11520
      TabIndex        =   1
      Top             =   2610
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   32896
      CalendarTitleBackColor=   192
      CalendarTitleForeColor=   8421631
      CheckBox        =   -1  'True
      CustomFormat    =   "dd-MM-yyyy"
      DateIsNull      =   -1  'True
      Format          =   123207681
      CurrentDate     =   43737
      MaxDate         =   46022
      MinDate         =   43101
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   6525
      TabIndex        =   0
      Top             =   2625
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   49344
      CalendarTitleBackColor=   192
      CalendarTitleForeColor=   8421631
      CheckBox        =   -1  'True
      CustomFormat    =   "dd-MM-yyyy"
      DateIsNull      =   -1  'True
      Format          =   123207681
      CurrentDate     =   43737
      MaxDate         =   46022
      MinDate         =   43466
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      Caption         =   "Total Discount"
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
      Height          =   375
      Left            =   16230
      TabIndex        =   16
      Top             =   8595
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      X1              =   15825
      X2              =   20025
      Y1              =   7605
      Y2              =   7605
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Height          =   3975
      Left            =   15825
      Top             =   6645
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Height          =   2895
      Left            =   4815
      Top             =   1980
      Width           =   10680
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      Caption         =   "Total Sale Between Dates"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   16215
      TabIndex        =   13
      Top             =   6975
      Width           =   3495
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
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
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   16605
      TabIndex        =   12
      Top             =   7980
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
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
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   16545
      TabIndex        =   10
      Top             =   9945
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
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
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   16590
      TabIndex        =   9
      Top             =   9255
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   10815
      TabIndex        =   6
      Top             =   2655
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   5385
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   735
      Left            =   8820
      Shape           =   4  'Rounded Rectangle
      Top             =   765
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
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
      Left            =   6885
      TabIndex        =   4
      Top             =   840
      Width           =   6975
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmdExit_Click()
Form8.Show
Unload Me
End Sub

Private Sub cmdSearch_Click()
Dim date1 As Date
Dim date2 As Date

If IsNull(DTPicker1.Value And DTPicker2.Value) Then

    MsgBox "You must select date", vbCritical, "Warning"
    Exit Sub
End If

date1 = Format(DTPicker1.Value, "mm-dd-yyyy")
date2 = Format(DTPicker2.Value, "mm-dd-yyyy")
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\OrderMania\ordermania.mdb;Persist Security Info=False"
rs.CursorLocation = adUseClient

If DTPicker2.Value < DTPicker1.Value Then
MsgBox "End Date Cannot Be Lesser Then Start Date", vbCritical, "Wrong Input"
Exit Sub
Else
Adodc1.RecordSource = "select * from order1 where order_date between #" & date1 & "# and #" & date2 & "#"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Please Enter Another Date", vbCritical, "No Record Found"
Else
Adodc1.Caption = Adodc1.RecordSource
End If
End If

con.Close

Call sale

End Sub


Public Sub sale()
Dim date3 As Date
Dim date4 As Date
Dim db As Database
Dim rsc As Recordset
Dim i As Integer
Dim Tot, gst, gtot, tds As Double
tds = 0
date3 = Format(DTPicker1.Value, "mm-dd-yyyy")
date4 = Format(DTPicker2.Value, "mm-dd-yyyy")
Set db = OpenDatabase("D:\OrderMania\ordermania.mdb")
Set rsc = db.OpenRecordset("select * from common where Order_Date between #" & date3 & "# and #" & date4 & "#")

For i = 0 To Adodc1.Recordset.RecordCount - 1
    Tot = Tot + CDbl(DataGrid1.Columns(5).Text)
    Adodc1.Recordset.MoveNext
Next i
    
While Not rsc.EOF
tds = tds + rsc.Fields(2).Value
rsc.MoveNext
Wend

txtSub.Text = Tot
txtDiscount.Text = tds
Tot = Tot - tds
gst = Tot * 0.05
txtGST.Text = gst
gtot = Tot + gst
txtGtot.Text = gtot

End Sub

Private Sub txtSub_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub txtGST_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub txtGtot_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

