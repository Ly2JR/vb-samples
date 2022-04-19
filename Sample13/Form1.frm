VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4875
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox Combo3 
      Height          =   300
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   1800
      Left            =   1800
      Style           =   1  'Simple Combo
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
Dim data(7) As String
Dim i As Integer
    data(0) = "计算机基础"
    data(1) = "数据结构"
    data(2) = "数据库原理"
    data(3) = "VB程序设计"
    data(4) = "C语言"
    data(5) = "微机原理"
    data(6) = "多媒体技术"
    
    For i = 0 To UBound(data) - 1
        Combo1.AddItem data(i)
        Combo2.AddItem data(i)
        Combo3.AddItem data(i)
    Next
End Sub
