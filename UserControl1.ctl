VERSION 5.00
Begin VB.UserControl UserControl1 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   " "
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   255
      Left            =   615
      TabIndex        =   1
      Top             =   30
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      Height          =   255
      Left            =   15
      TabIndex        =   0
      Top             =   30
      Width           =   255
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Property Get Text() As Integer
Text = Text1.Text
End Property

Public Property Let Text(ByVal NewValue As Integer)
Text1.Text = NewValue
PropertyChanged Text
End Property


Private Sub Command1_Click()
Text1.Text = Text1.Text + 1
End Sub

Private Sub Command1_LostFocus()
If Text1.Text = 0 Then
 MsgBox "The Quantity should not be Less Than ONE", vbInformation, "Alert!!"
 Exit Sub
End If
End Sub

Private Sub Command2_Click()
Text1.Text = Text1.Text - 1
End Sub

Private Sub Command2_LostFocus()
If Text1.Text = 0 Then
 MsgBox "The Quantity should not be Less Than ONE", vbInformation, "Alert!!"
 Exit Sub
End If
End Sub

Private Sub UserControl_Initialize()
Text1.Text = 1
Text1.Enabled = False
End Sub
