VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form14 
   Caption         =   "Customer FeedBack"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18480
   LinkTopic       =   "Form14"
   ScaleHeight     =   9165
   ScaleWidth      =   18480
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "customer feedback.frx":0000
      Height          =   3375
      Left            =   450
      TabIndex        =   5
      Top             =   7335
      Width           =   19350
      _ExtentX        =   34131
      _ExtentY        =   5953
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   15600
      Top             =   5400
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
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
      RecordSource    =   "select * from feedback"
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
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000011&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000011&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000011&
      Caption         =   "Show Graph"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   6375
      Left            =   390
      OleObjectBlob   =   "customer feedback.frx":0015
      TabIndex        =   0
      Top             =   450
      Visible         =   0   'False
      Width           =   12375
   End
   Begin VB.Label Label1 
      Caption         =   "Graph"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4800
      TabIndex        =   4
      Top             =   2280
      Width           =   6255
   End
   Begin VB.Shape Shape1 
      Height          =   6615
      Left            =   270
      Top             =   330
      Width           =   12615
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim c As Double
Private Sub star(i As Integer)
Dim a, b As Integer
Dim d As Double
Set db = OpenDatabase("D:\OrderMania\ordermania.mdb")
Set rs = db.OpenRecordset("select * from feedback")

b = (rs.RecordCount + 1) * 5
a = 0
While Not rs.EOF
If rs.Fields(i).Value = "Excellent" Then
a = a + 5
ElseIf rs.Fields(i).Value = "Good" Then
a = a + 4
ElseIf rs.Fields(i).Value = "Average" Then
a = a + 3
ElseIf rs.Fields(i).Value = "Poor" Then
a = a + 2
ElseIf rs.Fields(i).Value = "Dissatisfied" Then
a = a + 1
End If
rs.MoveNext
Wend

c = (a / b) * 10
End Sub

Private Sub Command1_Click()
MSChart1.Visible = True
Label1.Visible = False
Dim X(1 To 5) As Variant
Dim p, q, r, s, t As Double
Call star(4)
p = c
Call star(5)
q = c
Call star(6)
r = c
Call star(7)
s = c
Call star(9)
t = c


X(1) = p
X(2) = q
X(3) = r
X(4) = s
X(5) = t

MSChart1.ChartData = X
MSChart1.Column = 1
MSChart1.ColumnLabel = "Overall Ambience"
MSChart1.Column = 2
MSChart1.ColumnLabel = "Food Quality"
MSChart1.Column = 3
MSChart1.ColumnLabel = "Cleanliness"
MSChart1.Column = 4
MSChart1.ColumnLabel = "Speed of Service"
MSChart1.Column = 5
MSChart1.ColumnLabel = "Overall Experience"


End Sub

Private Sub Command2_Click()
Form8.Show
Unload Me
End Sub

Private Sub Command3_Click()
MSChart1.EditCopy

Printer.Print " "
Printer.PaintPicture Clipboard.GetData(), 20, 700  'Position of the Chart on print
Printer.EndDoc
End Sub

