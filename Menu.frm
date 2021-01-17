VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Menu"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15645
   LinkTopic       =   "Form4"
   ScaleHeight     =   11055
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox TextOrNO 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   8100
      TabIndex        =   145
      Top             =   1035
      Width           =   2415
   End
   Begin VB.CommandButton cmdBill 
      Caption         =   "View Receipt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10800
      TabIndex        =   101
      Top             =   9960
      Width           =   1695
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3165
      TabIndex        =   47
      Top             =   9885
      Width           =   1815
   End
   Begin VB.TextBox TxtName 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   3300
      TabIndex        =   46
      Top             =   1035
      Width           =   3615
   End
   Begin VB.TextBox TxtDate 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   15420
      TabIndex        =   10
      Text            =   " "
      Top             =   1035
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Beverages"
      Height          =   5415
      Left            =   5400
      TabIndex        =   9
      Top             =   1770
      Width           =   4455
      Begin VB.CheckBox Check18 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Mojito Mocktail"
         Height          =   375
         Left            =   270
         TabIndex        =   37
         Top             =   4860
         Width           =   1455
      End
      Begin VB.CheckBox Check17 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sprite"
         Height          =   375
         Left            =   255
         TabIndex        =   36
         Top             =   4230
         Width           =   1095
      End
      Begin VB.CheckBox Check16 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Coke "
         Height          =   495
         Left            =   240
         TabIndex        =   35
         Top             =   3525
         Width           =   1095
      End
      Begin VB.CheckBox Check15 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cold Coffee"
         Height          =   375
         Left            =   225
         TabIndex        =   34
         Top             =   2940
         Width           =   1215
      End
      Begin VB.CheckBox Check14 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Bottled Water"
         Height          =   375
         Left            =   225
         TabIndex        =   33
         Top             =   2340
         Width           =   1335
      End
      Begin VB.CheckBox Check13 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Coffee"
         Height          =   495
         Left            =   225
         TabIndex        =   32
         Top             =   1620
         Width           =   975
      End
      Begin VB.CheckBox Check12 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Tea"
         Height          =   375
         Left            =   225
         TabIndex        =   31
         Top             =   1140
         Width           =   1215
      End
      Begin Project1.UserControl1 UserControl12 
         Height          =   375
         Left            =   1860
         TabIndex        =   113
         Top             =   1170
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl15 
         Height          =   375
         Left            =   1875
         TabIndex        =   114
         Top             =   2910
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl16 
         Height          =   375
         Left            =   1875
         TabIndex        =   115
         Top             =   3600
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl17 
         Height          =   375
         Left            =   1905
         TabIndex        =   116
         Top             =   4260
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl18 
         Height          =   375
         Left            =   1920
         TabIndex        =   117
         Top             =   4845
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl13 
         Height          =   375
         Left            =   1845
         TabIndex        =   118
         Top             =   1695
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl14 
         Height          =   375
         Left            =   1875
         TabIndex        =   119
         Top             =   2340
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin VB.Image Image5 
         Height          =   885
         Left            =   1590
         Picture         =   "Menu.frx":0000
         Stretch         =   -1  'True
         Top             =   165
         Width           =   960
      End
      Begin VB.Label Label33 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.200/-"
         Height          =   255
         Left            =   3360
         TabIndex        =   44
         Top             =   4935
         Width           =   735
      End
      Begin VB.Label Label32 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.25/-"
         Height          =   375
         Left            =   3375
         TabIndex        =   43
         Top             =   4350
         Width           =   855
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.25/-"
         Height          =   255
         Left            =   3330
         TabIndex        =   42
         Top             =   3690
         Width           =   735
      End
      Begin VB.Label Label30 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.35/-"
         Height          =   375
         Left            =   3330
         TabIndex        =   41
         Top             =   2985
         Width           =   735
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.20/-"
         Height          =   375
         Left            =   3300
         TabIndex        =   40
         Top             =   2460
         Width           =   735
      End
      Begin VB.Label Label28 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.20/-"
         Height          =   375
         Left            =   3330
         TabIndex        =   39
         Top             =   1815
         Width           =   855
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.12/-"
         Height          =   375
         Left            =   3330
         TabIndex        =   38
         Top             =   1215
         Width           =   855
      End
   End
   Begin VB.TextBox TxtTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13965
      TabIndex        =   8
      Text            =   " "
      Top             =   10080
      Width           =   2415
   End
   Begin VB.ComboBox ComboTable 
      BackColor       =   &H00FFC0C0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   11805
      TabIndex        =   7
      Top             =   1035
      Width           =   2415
   End
   Begin VB.CommandButton CmdDone 
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   765
      TabIndex        =   6
      Top             =   9885
      Width           =   2055
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Extras"
      Height          =   3615
      Left            =   5415
      TabIndex        =   90
      Top             =   7290
      Width           =   4455
      Begin VB.CheckBox Check43 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Lachha Paratha"
         Height          =   375
         Left            =   240
         TabIndex        =   95
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CheckBox Check42 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Butter Naan"
         Height          =   375
         Left            =   240
         TabIndex        =   94
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CheckBox Check41 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Naan"
         Height          =   255
         Left            =   240
         TabIndex        =   93
         Top             =   2040
         Width           =   975
      End
      Begin VB.CheckBox Check40 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Butter Roti"
         Height          =   375
         Left            =   240
         TabIndex        =   92
         Top             =   1425
         Width           =   1215
      End
      Begin VB.CheckBox Check39 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Roti"
         Height          =   255
         Left            =   225
         TabIndex        =   91
         Top             =   840
         Width           =   1575
      End
      Begin Project1.UserControl1 UserControl41 
         Height          =   375
         Left            =   1875
         TabIndex        =   120
         Top             =   1425
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl42 
         Height          =   375
         Left            =   1875
         TabIndex        =   121
         Top             =   2010
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl43 
         Height          =   375
         Left            =   1890
         TabIndex        =   122
         Top             =   2550
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl44 
         Height          =   375
         Left            =   1890
         TabIndex        =   123
         Top             =   3135
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl40 
         Height          =   375
         Left            =   1860
         TabIndex        =   124
         Top             =   825
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.35/-"
         Height          =   210
         Left            =   3285
         TabIndex        =   100
         Top             =   3225
         Width           =   615
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.35/-"
         Height          =   255
         Left            =   3270
         TabIndex        =   99
         Top             =   2610
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.30/-"
         Height          =   255
         Left            =   3240
         TabIndex        =   98
         Top             =   2070
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.20/-"
         Height          =   255
         Left            =   3240
         TabIndex        =   97
         Top             =   1530
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.15/-"
         Height          =   255
         Left            =   3240
         TabIndex        =   96
         Top             =   840
         Width           =   735
      End
      Begin VB.Image Image6 
         Height          =   735
         Left            =   840
         Picture         =   "Menu.frx":280B
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "BreakFast"
      Height          =   7815
      Left            =   405
      TabIndex        =   2
      Top             =   1770
      Width           =   4815
      Begin Project1.UserControl1 UserControl1 
         Height          =   375
         Left            =   2445
         TabIndex        =   102
         Top             =   1605
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin VB.CheckBox Check11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cheese Garlic Bread"
         Height          =   495
         Left            =   405
         TabIndex        =   20
         Top             =   7185
         Width           =   1800
      End
      Begin VB.CheckBox Check10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cheese Sandwich"
         Height          =   375
         Left            =   405
         TabIndex        =   19
         Top             =   6705
         Width           =   1695
      End
      Begin VB.CheckBox Check9 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Club Sandwich"
         Height          =   495
         Left            =   405
         TabIndex        =   18
         Top             =   6105
         Width           =   1455
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Masala Dosa"
         Height          =   495
         Left            =   405
         TabIndex        =   17
         Top             =   5505
         Width           =   1575
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Plain Dosa"
         Height          =   855
         Left            =   405
         TabIndex        =   16
         Top             =   4665
         Width           =   1575
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Dahi Wada"
         Height          =   855
         Left            =   405
         TabIndex        =   15
         Top             =   4065
         Width           =   1335
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Upma"
         Height          =   615
         Left            =   405
         TabIndex        =   14
         Top             =   3585
         Width           =   1095
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Steamed Idli"
         Height          =   495
         Left            =   405
         TabIndex        =   13
         Top             =   3105
         Width           =   1215
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "VadaPav"
         Height          =   615
         Left            =   405
         TabIndex        =   5
         Top             =   2505
         Width           =   975
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Samosa"
         Height          =   735
         Left            =   405
         TabIndex        =   4
         Top             =   1905
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Poha"
         Height          =   615
         Left            =   405
         TabIndex        =   3
         Top             =   1425
         Width           =   975
      End
      Begin Project1.UserControl1 UserControl4 
         Height          =   375
         Left            =   2445
         TabIndex        =   103
         Top             =   3195
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl5 
         Height          =   375
         Left            =   2445
         TabIndex        =   104
         Top             =   3780
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl6 
         Height          =   375
         Left            =   2445
         TabIndex        =   105
         Top             =   4365
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl7 
         Height          =   375
         Left            =   2430
         TabIndex        =   106
         Top             =   4980
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl8 
         Height          =   375
         Left            =   2415
         TabIndex        =   107
         Top             =   5595
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl9 
         Height          =   375
         Left            =   2430
         TabIndex        =   108
         Top             =   6150
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl10 
         Height          =   375
         Left            =   2430
         TabIndex        =   109
         Top             =   6735
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl11 
         Height          =   375
         Left            =   2445
         TabIndex        =   110
         Top             =   7290
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl2 
         Height          =   375
         Left            =   2445
         TabIndex        =   111
         Top             =   2145
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl3 
         Height          =   375
         Left            =   2460
         TabIndex        =   112
         Top             =   2640
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin VB.Image Image4 
         Height          =   1215
         Left            =   1440
         Picture         =   "Menu.frx":9ADF
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.20/-"
         Height          =   255
         Left            =   3630
         TabIndex        =   89
         Top             =   1650
         Width           =   735
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.70/-"
         Height          =   255
         Left            =   3630
         TabIndex        =   30
         Top             =   6825
         Width           =   615
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.120/-"
         Height          =   255
         Left            =   3630
         TabIndex        =   29
         Top             =   7320
         Width           =   735
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.60/-"
         Height          =   255
         Left            =   3630
         TabIndex        =   28
         Top             =   6210
         Width           =   735
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.35/-"
         Height          =   255
         Left            =   3630
         TabIndex        =   27
         Top             =   5010
         Width           =   615
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.30/-"
         Height          =   255
         Left            =   3630
         TabIndex        =   26
         Top             =   4410
         Width           =   735
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.20/-"
         Height          =   255
         Left            =   3630
         TabIndex        =   25
         Top             =   3810
         Width           =   615
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.20/-"
         Height          =   255
         Left            =   3630
         TabIndex        =   24
         Top             =   3210
         Width           =   615
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.12/-"
         Height          =   255
         Left            =   3630
         TabIndex        =   23
         Top             =   2610
         Width           =   735
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.15/-"
         Height          =   255
         Left            =   3630
         TabIndex        =   22
         Top             =   2130
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.45/-"
         Height          =   255
         Index           =   0
         Left            =   3630
         TabIndex        =   21
         Top             =   5610
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Main Course"
      Height          =   7575
      Left            =   10125
      TabIndex        =   48
      Top             =   1770
      Width           =   9855
      Begin VB.CheckBox Check38 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Hyderabadi Chicken Biryani"
         Height          =   375
         Index           =   1
         Left            =   5190
         TabIndex        =   78
         Top             =   6720
         Width           =   2295
      End
      Begin VB.CheckBox Check37 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Prawn Biryani"
         Height          =   255
         Index           =   1
         Left            =   5190
         TabIndex        =   77
         Top             =   6225
         Width           =   1335
      End
      Begin VB.CheckBox Check36 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Egg Biryani"
         Height          =   375
         Index           =   1
         Left            =   5190
         TabIndex        =   76
         Top             =   5655
         Width           =   1335
      End
      Begin VB.CheckBox Check35 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Kashmiri Style Shab Deg"
         Height          =   375
         Index           =   1
         Left            =   5190
         TabIndex        =   75
         Top             =   5025
         Width           =   2055
      End
      Begin VB.CheckBox Check34 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Mutton Korma"
         Height          =   375
         Index           =   1
         Left            =   5190
         TabIndex        =   74
         Top             =   4440
         Width           =   2175
      End
      Begin VB.CheckBox Check33 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Afghani Chicken Korma"
         Height          =   495
         Index           =   1
         Left            =   5190
         TabIndex        =   73
         Top             =   3705
         Width           =   2055
      End
      Begin VB.CheckBox Check32 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Parsi Salli Murgh"
         Height          =   255
         Index           =   1
         Left            =   5190
         TabIndex        =   72
         Top             =   3255
         Width           =   1575
      End
      Begin VB.CheckBox Check31 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Masala Pomfret Fish"
         Height          =   495
         Index           =   1
         Left            =   5190
         TabIndex        =   71
         Top             =   2535
         Width           =   1935
      End
      Begin VB.CheckBox Check29 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Egg Curry"
         Height          =   375
         Index           =   1
         Left            =   5160
         TabIndex        =   70
         Top             =   1455
         Width           =   1335
      End
      Begin VB.CheckBox Check30 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Tandoori Prawns"
         Height          =   495
         Index           =   1
         Left            =   5190
         TabIndex        =   69
         Top             =   1935
         Width           =   1575
      End
      Begin VB.CheckBox Check28 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Kashmiri Biryani"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   58
         Top             =   6840
         Width           =   1695
      End
      Begin VB.CheckBox Check27 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Veg. Biryani"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   57
         Top             =   6360
         Width           =   1335
      End
      Begin VB.CheckBox Check26 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Malai Kofta"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   56
         Top             =   5760
         Width           =   1335
      End
      Begin VB.CheckBox Check25 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Red Lentil Curry"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   55
         Top             =   5160
         Width           =   1815
      End
      Begin VB.CheckBox Check24 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Indian Stuffed Eggplant"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   54
         Top             =   4560
         Width           =   2175
      End
      Begin VB.CheckBox Check23 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Tofu Keema"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   53
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CheckBox Check22 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Navratan Korma"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   52
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CheckBox Check21 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Shahi Paneer"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   51
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CheckBox Check20 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Chickpea Curry"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   50
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CheckBox Check19 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Shev Bhaji"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   49
         Top             =   1560
         Width           =   1335
      End
      Begin Project1.UserControl1 UserControl20 
         Height          =   375
         Left            =   2400
         TabIndex        =   125
         Top             =   1545
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl23 
         Height          =   375
         Left            =   2400
         TabIndex        =   126
         Top             =   3330
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl24 
         Height          =   375
         Left            =   2400
         TabIndex        =   127
         Top             =   3930
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl25 
         Height          =   375
         Left            =   2400
         TabIndex        =   128
         Top             =   4605
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl26 
         Height          =   375
         Left            =   2400
         TabIndex        =   129
         Top             =   5205
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl27 
         Height          =   375
         Left            =   2385
         TabIndex        =   130
         Top             =   5760
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl28 
         Height          =   375
         Left            =   2400
         TabIndex        =   131
         Top             =   6315
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl29 
         Height          =   375
         Left            =   2400
         TabIndex        =   132
         Top             =   6885
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl21 
         Height          =   375
         Left            =   2370
         TabIndex        =   133
         Top             =   2145
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl22 
         Height          =   375
         Left            =   2385
         TabIndex        =   134
         Top             =   2745
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl30 
         Height          =   375
         Left            =   7620
         TabIndex        =   135
         Top             =   1440
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl33 
         Height          =   375
         Left            =   7620
         TabIndex        =   136
         Top             =   3225
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl34 
         Height          =   375
         Left            =   7605
         TabIndex        =   137
         Top             =   3780
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl35 
         Height          =   375
         Left            =   7605
         TabIndex        =   138
         Top             =   4455
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl36 
         Height          =   375
         Left            =   7605
         TabIndex        =   139
         Top             =   5070
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl37 
         Height          =   375
         Left            =   7620
         TabIndex        =   140
         Top             =   5655
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl38 
         Height          =   375
         Left            =   7605
         TabIndex        =   141
         Top             =   6165
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl39 
         Height          =   375
         Left            =   7605
         TabIndex        =   142
         Top             =   6735
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl31 
         Height          =   375
         Left            =   7620
         TabIndex        =   143
         Top             =   2040
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
      End
      Begin Project1.UserControl1 UserControl32 
         Height          =   375
         Left            =   7590
         TabIndex        =   144
         Top             =   2655
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
      End
      Begin VB.Image Image3 
         Height          =   855
         Left            =   360
         Picture         =   "Menu.frx":11C5C
         Stretch         =   -1  'True
         Top             =   480
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   960
         Left            =   5040
         Picture         =   "Menu.frx":1C216
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   4800
         X2              =   4800
         Y1              =   7320
         Y2              =   360
      End
      Begin VB.Label Label42 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.290/-"
         Height          =   375
         Index           =   1
         Left            =   8805
         TabIndex        =   88
         Top             =   5760
         Width           =   735
      End
      Begin VB.Label Label44 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.310/-"
         Height          =   375
         Index           =   1
         Left            =   8775
         TabIndex        =   87
         Top             =   6825
         Width           =   975
      End
      Begin VB.Label Label43 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.330/-"
         Height          =   375
         Index           =   1
         Left            =   8760
         TabIndex        =   86
         Top             =   6330
         Width           =   615
      End
      Begin VB.Label Label41 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.370/-"
         Height          =   255
         Index           =   1
         Left            =   8775
         TabIndex        =   85
         Top             =   5130
         Width           =   735
      End
      Begin VB.Label Label40 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.320/-"
         Height          =   375
         Index           =   1
         Left            =   8760
         TabIndex        =   84
         Top             =   4530
         Width           =   735
      End
      Begin VB.Label Label39 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.350/-"
         Height          =   375
         Index           =   1
         Left            =   8760
         TabIndex        =   83
         Top             =   3930
         Width           =   855
      End
      Begin VB.Label Label38 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.370/-"
         Height          =   255
         Index           =   1
         Left            =   8760
         TabIndex        =   82
         Top             =   3330
         Width           =   975
      End
      Begin VB.Label Label37 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.360/-"
         Height          =   255
         Index           =   1
         Left            =   8715
         TabIndex        =   81
         Top             =   2700
         Width           =   855
      End
      Begin VB.Label Label36 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.350/-"
         Height          =   255
         Index           =   1
         Left            =   8715
         TabIndex        =   80
         Top             =   2070
         Width           =   735
      End
      Begin VB.Label Label35 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.250/-"
         Height          =   255
         Index           =   1
         Left            =   8700
         TabIndex        =   79
         Top             =   1530
         Width           =   735
      End
      Begin VB.Label Label44 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.250/-"
         Height          =   375
         Index           =   0
         Left            =   3795
         TabIndex        =   68
         Top             =   6990
         Width           =   975
      End
      Begin VB.Label Label43 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.230/-"
         Height          =   375
         Index           =   0
         Left            =   3795
         TabIndex        =   67
         Top             =   6390
         Width           =   615
      End
      Begin VB.Label Label42 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.230/-"
         Height          =   375
         Index           =   0
         Left            =   3795
         TabIndex        =   66
         Top             =   5790
         Width           =   735
      End
      Begin VB.Label Label41 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.210/-"
         Height          =   255
         Index           =   0
         Left            =   3795
         TabIndex        =   65
         Top             =   5190
         Width           =   735
      End
      Begin VB.Label Label40 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.190/-"
         Height          =   375
         Index           =   0
         Left            =   3795
         TabIndex        =   64
         Top             =   4590
         Width           =   735
      End
      Begin VB.Label Label39 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.200/-"
         Height          =   375
         Index           =   0
         Left            =   3795
         TabIndex        =   63
         Top             =   3990
         Width           =   855
      End
      Begin VB.Label Label38 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.220/-"
         Height          =   255
         Index           =   0
         Left            =   3795
         TabIndex        =   62
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Rs.190/-"
         Height          =   255
         Index           =   0
         Left            =   3780
         TabIndex        =   61
         Top             =   2820
         Width           =   855
      End
      Begin VB.Label Label36 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.150/-"
         Height          =   255
         Index           =   0
         Left            =   3795
         TabIndex        =   60
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label35 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rs.130/-"
         Height          =   255
         Index           =   0
         Left            =   3795
         TabIndex        =   59
         Top             =   1560
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   1200
         Picture         =   "Menu.frx":2BB5E
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1275
      End
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   18600
      TabIndex        =   147
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   735
      Left            =   9045
      Shape           =   4  'Rounded Rectangle
      Top             =   165
      Width           =   3135
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
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
      ForeColor       =   &H000040C0&
      Height          =   855
      Left            =   7455
      TabIndex        =   146
      Top             =   210
      Width           =   6255
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      Height          =   735
      Left            =   1800
      TabIndex        =   45
      Top             =   1095
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13125
      TabIndex        =   12
      Top             =   10245
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   375
      Left            =   14640
      TabIndex        =   11
      Top             =   1095
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Table no"
      Height          =   495
      Left            =   10800
      TabIndex        =   1
      Top             =   1095
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Order no"
      Height          =   375
      Left            =   7320
      TabIndex        =   0
      Top             =   1095
      Width           =   615
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim rsc As Recordset
Dim Tot, a As Integer



Private Sub cmdCancel_Click()
Unload Me
Form1.Show
End Sub

Private Sub ComboTable_Change()
ComboTable.AddItem "TakeAway Counter"
ComboTable.AddItem "1"
ComboTable.AddItem "2"
ComboTable.AddItem "3"
ComboTable.AddItem "4"
ComboTable.AddItem "5"
ComboTable.AddItem "6"
ComboTable.AddItem "7"
ComboTable.AddItem "8"
ComboTable.AddItem "9"
ComboTable.AddItem "10"
End Sub

Private Sub CmdDone_Click()
Dim flag, re As Integer
flag = 0
Tot = 0
Set db = OpenDatabase("D:\OrderMania\ordermania.mdb")
Set rs = db.OpenRecordset("select * from order1")
Set rsc = db.OpenRecordset("select * from common")

If StrConv(txtName.Text, vbProperCase) = "" Then
    MsgBox ("Please Enter Your Name"), vbExclamation
    txtName.SetFocus
    flag = 1
End If

For re = 1 To rs.RecordCount
rs.MoveLast
If ComboTable.Text = "Select" Then
       MsgBox ("Please Select The Table Number"), vbExclamation
       ComboTable.SetFocus
       flag = 1
 Else
        If Check1.Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(3).Value = UserControl1.Text
        rs.Fields(2).Value = Check1.Caption
        rs.Fields(4).Value = "20"
        rs.Fields(5).Value = UserControl1.Text * 20
        Tot = Tot + (UserControl1.Text * 20)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check2.Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(3).Value = UserControl2.Text
        rs.Fields(2).Value = Check2.Caption
        rs.Fields(4).Value = "15"
        rs.Fields(5).Value = UserControl2.Text * 15
        Tot = Tot + (UserControl2.Text * 15)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        
        End If
        
        If Check3.Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl3.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check3.Caption
        rs.Fields(4).Value = "12"
        rs.Fields(5).Value = UserControl3.Text * 12
        Tot = Tot + (UserControl3.Text * 12)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"

        rs.Update
        
        End If
        
        If Check4.Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl4.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check4.Caption
        rs.Fields(4).Value = "20"
        rs.Fields(5).Value = UserControl4.Text * 20
        Tot = Tot + (UserControl4.Text * 20)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        
        End If
        
        If Check5.Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl5.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check5.Caption
        rs.Fields(4).Value = "20"
        rs.Fields(5).Value = UserControl5.Text * 20
        Tot = Tot + (UserControl5.Text * 20)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        
        End If
        
        If Check6.Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl6.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check6.Caption
        rs.Fields(4).Value = "30"
        rs.Fields(5).Value = UserControl6.Text * 30
        Tot = Tot + (UserControl6.Text * 30)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        
        End If
        
        If Check7.Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl7.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check7.Caption
        rs.Fields(4).Value = "35"
        rs.Fields(5).Value = UserControl7.Text * 35
        Tot = Tot + (UserControl7.Text * 35)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check8.Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl8.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check8.Caption
        rs.Fields(4).Value = "45"
        rs.Fields(5).Value = UserControl8.Text * 45
        Tot = Tot + (UserControl8.Text * 45)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check9.Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl9.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check9.Caption
        rs.Fields(4).Value = "60"
        rs.Fields(5).Value = UserControl9.Text * 60
        Tot = Tot + (UserControl9.Text * 60)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check10.Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl10.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check10.Caption
        rs.Fields(4).Value = "70"
        rs.Fields(5).Value = UserControl10.Text * 70
        Tot = Tot + (UserControl10.Text * 70)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check11.Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl11.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check11.Caption
        rs.Fields(4).Value = "120"
        rs.Fields(5).Value = UserControl11.Text * 120
        Tot = Tot + (UserControl11.Text * 120)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check12.Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl12.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check12.Caption
        rs.Fields(4).Value = "12"
        rs.Fields(5).Value = UserControl12.Text * 12
        Tot = Tot + (UserControl12.Text * 12)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check13.Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl13.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check13.Caption
        rs.Fields(4).Value = "20"
        rs.Fields(5).Value = UserControl13.Text * 20
        Tot = Tot + (UserControl13.Text * 20)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check14.Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl14.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check14.Caption
        rs.Fields(4).Value = "20"
        rs.Fields(5).Value = UserControl14.Text * 20
        Tot = Tot + (UserControl14.Text * 20)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check15.Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl15.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check15.Caption
        rs.Fields(4).Value = "35"
        rs.Fields(5).Value = UserControl15.Text * 35
        Tot = Tot + (UserControl15.Text * 35)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check16.Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl16.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check16.Caption
        rs.Fields(4).Value = "25"
        rs.Fields(5).Value = UserControl16.Text * 25
        Tot = Tot + (UserControl16.Text * 25)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check17.Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl17.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check17.Caption
        rs.Fields(4).Value = "25"
        rs.Fields(5).Value = UserControl17.Text * 25
        Tot = Tot + (UserControl17.Text * 25)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check18.Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl18.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check18.Caption
        rs.Fields(4).Value = "200"
        rs.Fields(5).Value = UserControl18.Text * 200
        Tot = Tot + (UserControl18.Text * 200)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check19(0).Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl20.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check19(0).Caption
        rs.Fields(4).Value = "130"
        rs.Fields(5).Value = UserControl20.Text * 130
        Tot = Tot + (UserControl20.Text * 130)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check20(0).Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl21.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check20(0).Caption
        rs.Fields(4).Value = "150"
        rs.Fields(5).Value = UserControl21.Text * 150
        Tot = Tot + (UserControl21.Text * 150)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check21(0).Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl22.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check21(0).Caption
        rs.Fields(4).Value = "190"
        rs.Fields(5).Value = UserControl22.Text * 190
        Tot = Tot + (UserControl22.Text * 190)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check22(0).Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl23.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check22(0).Caption
        rs.Fields(4).Value = "220"
        rs.Fields(5).Value = UserControl23.Text * 220
        Tot = Tot + (UserControl23.Text * 220)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check23(0).Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl24.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check23(0).Caption
        rs.Fields(4).Value = "200"
        rs.Fields(5).Value = UserControl24.Text * 200
        Tot = Tot + (UserControl24.Text * 200)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check24(0).Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl25.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check24(0).Caption
        rs.Fields(4).Value = "190"
        rs.Fields(5).Value = UserControl25.Text * 190
        Tot = Tot + (UserControl25.Text * 190)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check25(0).Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl26.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check25(0).Caption
        rs.Fields(4).Value = "210"
        rs.Fields(5).Value = UserControl26.Text * 210
        Tot = Tot + (UserControl26.Text * 210)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check26(0).Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl27.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check26(0).Caption
        rs.Fields(4).Value = "230"
        rs.Fields(5).Value = UserControl27.Text * 230
        Tot = Tot + (UserControl27.Text * 230)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check27(0).Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl28.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check27(0).Caption
        rs.Fields(4).Value = "230"
        rs.Fields(5).Value = UserControl28.Text * 230
        Tot = Tot + (UserControl28.Text * 230)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check28(0).Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl29.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check28(0).Caption
        rs.Fields(4).Value = "250"
        rs.Fields(5).Value = UserControl29.Text * 250
        Tot = Tot + (UserControl29.Text * 250)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check29(1).Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl30.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check29(1).Caption
        rs.Fields(4).Value = "250"
        rs.Fields(5).Value = UserControl30.Text * 250
        Tot = Tot + (UserControl30.Text * 250)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check30(1).Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl31.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check30(1).Caption
        rs.Fields(4).Value = "350"
        rs.Fields(5).Value = UserControl31.Text * 350
        Tot = Tot + (UserControl31.Text * 350)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check31(1).Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl32.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check31(1).Caption
        rs.Fields(4).Value = "360"
        rs.Fields(5).Value = UserControl32.Text * 360
        Tot = Tot + (UserControl32.Text * 360)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check32(1).Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl33.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check32(1).Caption
        rs.Fields(4).Value = "370"
        rs.Fields(5).Value = UserControl33.Text * 370
        Tot = Tot + (UserControl33.Text * 370)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check33(1).Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl34.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check33(1).Caption
        rs.Fields(4).Value = "350"
        rs.Fields(5).Value = UserControl34.Text * 350
        Tot = Tot + (UserControl34.Text * 350)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check34(1).Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl35.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check34(1).Caption
        rs.Fields(4).Value = "320"
        rs.Fields(5).Value = UserControl35.Text * 320
        Tot = Tot + (UserControl35.Text * 320)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check35(1).Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl36.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check35(1).Caption
        rs.Fields(4).Value = "370"
        rs.Fields(5).Value = UserControl36.Text * 370
        Tot = Tot + (UserControl36.Text * 370)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check36(1).Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl37.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check36(1).Caption
        rs.Fields(4).Value = "290"
        rs.Fields(5).Value = UserControl37.Text * 290
        Tot = Tot + (UserControl37.Text * 290)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check37(1).Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl38.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check37(1).Caption
        rs.Fields(4).Value = "330"
        rs.Fields(5).Value = UserControl38.Text * 330
        Tot = Tot + (UserControl38.Text * 330)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check38(1).Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl39.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check38(1).Caption
        rs.Fields(4).Value = "310"
        rs.Fields(5).Value = UserControl39.Text * 310
        Tot = Tot + (UserControl39.Text * 310)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check39.Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl40.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check39.Caption
        rs.Fields(4).Value = "15"
        rs.Fields(5).Value = UserControl40.Text * 15
        Tot = Tot + (UserControl40.Text * 15)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check40.Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl41.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check40.Caption
        rs.Fields(4).Value = "20"
        rs.Fields(5).Value = UserControl41.Text * 20
        Tot = Tot + (UserControl41.Text * 20)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check41.Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl42.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check41.Caption
        rs.Fields(4).Value = "30"
        rs.Fields(5).Value = UserControl42.Text * 30
        Tot = Tot + (UserControl42.Text * 30)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check42.Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl43.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check42.Caption
        rs.Fields(4).Value = "35"
        rs.Fields(5).Value = UserControl43.Text * 35
        Tot = Tot + (UserControl43.Text * 35)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
        
        If Check43.Value = 1 Then
        rs.AddNew
        rs.Fields(7).Value = StrConv(txtName.Text, vbProperCase)
        rs.Fields(0).Value = TextOrNO.Text
        rs.Fields(1).Value = ComboTable.Text
        rs.Fields(3).Value = UserControl44.Text
        rs.Fields(6).Value = txtDate.Text
        rs.Fields(2).Value = Check43.Caption
        rs.Fields(4).Value = "35"
        rs.Fields(5).Value = UserControl44.Text * 35
        Tot = Tot + (UserControl44.Text * 35)
        rs.Fields(8).Value = "Pending"
        rs.Fields(9).Value = "Waiting"
        rs.Update
        End If
    
End If
rs.MoveNext
Next re
TxtTotal.Text = Tot
Tot = CInt(TxtTotal.Text)


rsc.MoveLast
rsc.AddNew
rsc.Fields(1).Value = StrConv(txtName.Text, vbProperCase)
rsc.Fields(0).Value = TextOrNO.Text
rsc.Fields(2).Value = 0
rsc.Fields(3).Value = 0
rsc.Fields(4).Value = txtDate.Text
rsc.Update

If flag = 0 Then
MsgBox "Your Order has been Placed", vbOKOnly, "Done"
End If
End Sub

Private Sub cmdBill_Click()
Dim gst, gtot, Tot1 As Double
Tot1 = 0
Set db = OpenDatabase("D:\OrderMania\ordermania.mdb")
Set rs = db.OpenRecordset("select * from order1")


While Not rs.EOF
  If StrConv(txtName.Text, vbProperCase) = rs.Fields(7).Value And TextOrNO.Text = rs.Fields(0).Value Then

      Form7.ListItems.AddItem (rs.Fields(2).Value)
      Form7.ListPrice.AddItem (rs.Fields(4).Value)
      Form7.ListQty.AddItem (rs.Fields(3).Value)
      Form7.ListAmt.AddItem (rs.Fields(5).Value)
      Form7.txtTableNo.Text = rs.Fields(1).Value
      Form7.txtDate.Text = rs.Fields(6).Value
      Tot1 = Tot1 + rs.Fields(5).Value
    End If
rs.MoveNext
Wend
Form7.txtSub.Text = Tot1
gst = Tot1 * 0.05
Form7.txtGST.Text = gst
gtot = Tot1 + gst
Form7.txtGtot.Text = gtot
Form7.txtName.Text = StrConv(txtName.Text, vbProperCase)
Form7.txtOrdNo.Text = TextOrNO.Text
Form7.Show
Unload Me

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
Dim ch As String

    ch = Chr$(KeyAscii)
    If Not ((ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyRight Or kayascii = vbKeyLeft) Then
        KeyAscii = 0
    End If
End Sub


Private Sub Form_Activate()
txtName.SetFocus
End Sub
Private Sub generateOID()
Dim db1 As Database
Dim rsc As Recordset
Set db1 = OpenDatabase("D:\OrderMania\ordermania.mdb")
Set rsc = db1.OpenRecordset("select * from common")
If rsc.EOF = True Then
TextOrNO.Text = 1
Exit Sub
Else
rsc.MoveLast
TextOrNO.Text = rsc.Fields(0).Value + 1
End If

End Sub
Private Sub Form_Load()

Call generateOID
txtName.Text = ""
ComboTable.Text = "Select"
txtDate.Text = Date
lblTime.Caption = Time
TextOrNO.Enabled = False
txtDate.Enabled = False
txtName.TabIndex = 0
ComboTable.TabIndex = 1
End Sub


