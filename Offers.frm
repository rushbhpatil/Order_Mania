VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form15 
   BackColor       =   &H00404040&
   Caption         =   "Special Offers"
   ClientHeight    =   9735
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15015
   LinkTopic       =   "Form15"
   ScaleHeight     =   9735
   ScaleWidth      =   15015
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   16560
      Top             =   6720
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Back"
      Height          =   375
      Left            =   465
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   495
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1320
      Top             =   8040
   End
   Begin VB.Image ImageBankSlide 
      Height          =   2895
      Left            =   3615
      Stretch         =   -1  'True
      Top             =   7980
      Width           =   13095
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   15480
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   1300
      ImageHeight     =   480
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Offers.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Offers.frx":1C90D2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   735
      Left            =   2835
      Shape           =   4  'Rounded Rectangle
      Top             =   570
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
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
      Left            =   2970
      TabIndex        =   3
      Top             =   630
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000007&
      Height          =   135
      Index           =   2
      Left            =   10500
      Top             =   135
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000007&
      Height          =   135
      Index           =   1
      Left            =   9855
      Top             =   135
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblFrwd 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   99.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   2655
      Left            =   13320
      TabIndex        =   1
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lblBack 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   99.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   2655
      Left            =   6105
      TabIndex        =   0
      Top             =   2250
      Width           =   855
   End
   Begin VB.Image ImageOfferSlide 
      Height          =   7500
      Left            =   7380
      Stretch         =   -1  'True
      Top             =   345
      Width           =   5520
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Offers.frx":2D3524
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Offers.frx":D06782
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b

Private Sub cmdBack_Click()
Form1.Show
Unload Me
End Sub

Private Sub Form_Load()
a = 0
b = 0
ImageBankSlide.Picture = ImageList2.ListImages(2).Picture
ImageOfferSlide.Picture = ImageList1.ListImages(2).Picture
Timer1.Enabled = True
Timer2.Enabled = True
Shape1(1).Visible = True
Shape1(2).Visible = True
Shape1(2).BackColor = vbWhite
Shape1(1).BackColor = RGB(124, 124, 124)

End Sub

Private Sub ImageOfferSlide_Click()
Timer1.Enabled = True
End Sub

Private Sub lblFrwd_Click()
Timer1.Enabled = False
a = a + 1
If a <= ImageList1.ListImages.Count Then
ImageOfferSlide.Picture = ImageList1.ListImages(a).Picture
Else
a = 0
End If
If ImageOfferSlide.Picture = ImageList1.ListImages(2).Picture Then
Shape1(2).BackColor = vbWhite
Shape1(1).BackColor = RGB(124, 124, 124)
ElseIf ImageOfferSlide.Picture = ImageList1.ListImages(1).Picture Then
Shape1(2).BackColor = RGB(124, 124, 124)
Shape1(1).BackColor = vbWhite
End If
End Sub

Private Sub lblBack_Click()
Timer1.Enabled = False
a = a + 1
If a <= ImageList1.ListImages.Count Then
ImageOfferSlide.Picture = ImageList1.ListImages(a).Picture
Else
a = 0
End If
If ImageOfferSlide.Picture = ImageList1.ListImages(2).Picture Then
Shape1(2).BackColor = vbWhite
Shape1(1).BackColor = RGB(124, 124, 124)
ElseIf ImageOfferSlide.Picture = ImageList1.ListImages(1).Picture Then
Shape1(2).BackColor = RGB(124, 124, 124)
Shape1(1).BackColor = vbWhite
End If

End Sub

Private Sub lblFrwd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblFrwd.BackStyle = 1
End Sub

Private Sub lblFrwd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblFrwd.BackStyle = 0
End Sub

Private Sub lblBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblBack.BackStyle = 1
End Sub
Private Sub lblBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblBack.BackStyle = 0
End Sub

Private Sub Timer1_Timer()
a = a + 1
If a <= ImageList1.ListImages.Count Then
ImageOfferSlide.Picture = ImageList1.ListImages(a).Picture
Else
a = 0
End If
If ImageOfferSlide.Picture = ImageList1.ListImages(2).Picture Then
Shape1(2).BackColor = vbWhite
Shape1(1).BackColor = RGB(124, 124, 124)
ElseIf ImageOfferSlide.Picture = ImageList1.ListImages(1).Picture Then
Shape1(2).BackColor = RGB(124, 124, 124)
Shape1(1).BackColor = vbWhite
End If
End Sub

Private Sub Timer2_Timer()
b = b + 1
If b <= ImageList2.ListImages.Count Then
ImageBankSlide.Picture = ImageList2.ListImages(b).Picture
Else
b = 0
End If
End Sub
