VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Move方法"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "相对移动"
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "绝对移动"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Form1.Move 1000, 1000
End Sub


Private Sub Command2_Click()
    Form1.Move Left + 200, Top + 200
    
End Sub
