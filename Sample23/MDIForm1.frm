VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   2880
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   6708
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   492
      Left            =   0
      ScaleHeight     =   444
      ScaleWidth      =   6660
      TabIndex        =   0
      Top             =   0
      Width           =   6708
      Begin VB.CommandButton Command2 
         Caption         =   "退出"
         Height          =   372
         Left            =   1680
         TabIndex        =   2
         Top             =   0
         Width           =   1092
      End
      Begin VB.CommandButton Command1 
         Caption         =   "排列子窗体"
         Height          =   372
         Left            =   240
         TabIndex        =   1
         Top             =   0
         Width           =   1092
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim choose As Integer
    choose = InputBox("排列方式选择,请输入一个数值:(0-3)")
   MDIForm1.Arrange choose
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub MDIForm_Load()
    Form1.Show
    Form2.Show
    Form3.Show

End Sub
