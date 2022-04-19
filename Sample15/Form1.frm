VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "滚动条控件示例"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "背景色"
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "前景色"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   2160
      Width           =   855
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   480
      ScaleHeight     =   915
      ScaleWidth      =   1395
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Visual Basic 6.0"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "调色板"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Label2.ForeColor = Picture1.BackColor
End Sub

Private Sub Command2_Click()
    Label2.BackColor = Picture1.BackColor
End Sub

Private Sub Form_Load()
    HScroll1.Max = 15
    HScroll1.LargeChange = 2
    HScroll1.SmallChange = 1
End Sub


Private Sub HScroll1_Change()
    Picture1.BackColor = QBColor(HScroll1.Value)
End Sub


Private Sub HScroll1_Scroll()
    Picture1.BackColor = QBColor(HScroll1.Value)
End Sub
