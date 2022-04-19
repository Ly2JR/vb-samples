VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "pSet方法"
   ClientHeight    =   5652
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   9552
   LinkTopic       =   "Form1"
   ScaleHeight     =   5652
   ScaleWidth      =   9552
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "PSet方法"
      Height          =   372
      Left            =   8160
      TabIndex        =   1
      Top             =   5160
      Width           =   852
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   4812
      Left            =   0
      ScaleHeight     =   4764
      ScaleWidth      =   9504
      TabIndex        =   0
      Top             =   0
      Width           =   9552
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const PAI As Single = 3.141592654

Private Sub Command1_Click()
    Dim x As Single, y As Single
    Dim xt As Single, yt As Single
    Dim t As Single, aifa As Integer
    Picture1.ScaleMode = 6
    x = Picture1.ScaleWidth / 2
    y = Picture1.ScaleHeight / 2
    For aifa = 0 To 14400 Step 1
        t = PAI * aifa / 180#
        xt = Cos(t) + t * Sin(t)
        yt = -(Sin(t) - t * Cos(t))
        Picture1.PSet (xt + x, yt + y), vbBlack
    Next aifa
End Sub
