VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3144
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   4944
   LinkTopic       =   "Form1"
   ScaleHeight     =   3144
   ScaleWidth      =   4944
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "颜色渐变"
      Height          =   372
      Left            =   3840
      TabIndex        =   0
      Top             =   2640
      Width           =   852
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim i As Integer, sc As Single
    Dim x As Integer, y As Integer
    Dim r%, g%, b%
    x = Form1.ScaleWidth
    y = Form1.ScaleHeight
    
    sc = 255# / x       '设置需要改变颜色的增量
    For i = 0 To x
        r = (x - i) * sc
        g = (x - i) * sc
        b = (x - i) * sc
        Form1.Line (i, 0)-(i, y), RGB(r, g, b)
    Next i
End Sub
