VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Circle方法效果图"
   ClientHeight    =   4848
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   4848
   ScaleWidth      =   8700
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PAI  As Single = 3.141592654
Private Sub Form_Click()
    Dim aifa!, t!, r!, x!, y!, x0!, y0!
    Cls
    r = ScaleHeight / 4# * 0.9  '求圆的半径
    x0 = ScaleWidth / 2# '客户中心坐标
    y0 = ScaleHeight / 2#
    For aifa = 0# To 36# Step 18#
        t = aifa * PAI / 180#
        x = r * Cos(t) + x0
        y = r * Sin(t) + y0
        Circle (x, y), r
    Next aifa
End Sub

