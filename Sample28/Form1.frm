VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "自定义坐标系统"
   ClientHeight    =   6336
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   11172
   LinkTopic       =   "Form1"
   ScaleHeight     =   6336
   ScaleWidth      =   11172
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "画坐标系"
      Height          =   492
      Left            =   9240
      TabIndex        =   1
      Top             =   5160
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "画圆"
      Height          =   372
      Left            =   9240
      TabIndex        =   0
      Top             =   5760
      Width           =   1212
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Circle (0, 0), 1000
End Sub

Private Sub Command2_Click()

    Form1.Scale (-2000, 2000)-(2000, -2000) '设置自定义坐标系统
    Form1.DrawStyle = 0 '设置画线样式为实线
    Form1.DrawWidth = 2 '设置画线宽度为2个像素
    Form1.Line (-2000, 0)-(2000, 0) '画X轴，亮点的Y坐标值为0
    Form1.Line (2000 - 100, 30)-(2000, 0) '画x轴正向箭头
    Form1.Line (2000 - 100, -30)-(2000, 0) '
    Form1.CurrentX = 2000 - 150
    Form1.CurrentY = -100
    Form1.Print "X"
    
    Form1.Line (0, -2000)-(0, 2000) '画Y轴，两点x坐标值为0
    Form1.Line (-30, 2000 - 100)-(0, 2000) '画Y轴的箭头
    Form1.Line (30, 2000 - 100)-(0, 2000)
    Form1.CurrentX = 100
    Form1.CurrentY = 2000 - 100
    Form1.Print "Y"
    Form1.CurrentX = 100
    Form1.CurrentY = -100
    Form1.Print "O(0,0)"
    
End Sub
