VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "我的Visual Basic程序"
   ClientHeight    =   2685
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   4095
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "您输入的是"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "请输入姓名"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Text2.Text = Text1.Text
   
End Sub

Private Sub Command2_Click()
    Text2.Text = ""
    Text1.Text = ""
    
End Sub

Private Sub Command3_Click()
    End
End Sub
