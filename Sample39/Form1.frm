VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5160
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   8340
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "旋转"
      Height          =   492
      Left            =   6720
      TabIndex        =   0
      Top             =   4200
      Width           =   972
   End
   Begin VB.Line Line1 
      X1              =   2760
      X2              =   6000
      Y1              =   2040
      Y2              =   2040
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Const PAI As Single = 3.1415926
Dim mAngle As Integer '角度
Dim mColor As Integer '颜色
Dim mR As Single '半径
Private Sub Command1_Click()
    Dim t As Single
    mColor = mColor + 1
    If (mColor >= 16) Then mColor = 0   '颜色在0-15
    mAngle = mAngle + 10
    If mAngle >= 360 Then mAngle = 0 '角度在0~360
    t = mAngle * PAI / 180#     '将度数转换弧度
    Line1.X2 = Line1.X1 + mR * Cos(t)
    Line1.Y2 = Line1.Y1 + mR * Sin(t)
    Line1.BorderColor = QBColor(mColor)
End Sub

Private Sub Form_Load()
    Dim tx As Single, ty As Single
    tx = Line1.X2 - Line1.X1
    ty = Line1.Y2 - Line1.Y1
    mR = Sqr(tx * tx + ty + ty)
    Line1.X1 = ScaleWidth / 2#
    Line1.Y1 = Form1.ScaleHeight / 2#
    Line1.X2 = Line1.X1 + mR
    Line1.Y2 = Line1.Y1
    mColor = 0
    mAngle = 0
End Sub


