VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "列表框控件示例"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "<<"
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">>"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.ListBox List2 
      Height          =   2040
      Left            =   2760
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   2040
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "已选课程"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "可选课程"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If List1.ListIndex < 0 Then
        MsgBox "没有选择项"
        Exit Sub
    End If
    List2.AddItem List1.Text
    List1.RemoveItem List1.ListIndex
End Sub

Private Sub Command2_Click()
 If List2.ListIndex < 0 Then
        MsgBox "没有选择项"
        Exit Sub
    End If
    List1.AddItem List2.Text
    List2.RemoveItem List2.ListIndex
End Sub

Private Sub Command3_Click()
    Do While List1.ListCount
        List2.AddItem List1.List(0)
        List1.RemoveItem 0
    Loop
End Sub

Private Sub Command4_Click()
  Do While List2.ListCount
        List1.AddItem List2.List(0)
        List2.RemoveItem 0
    Loop
End Sub

Private Sub Form_Load()
    List1.AddItem "计算机文件基础"
    List1.AddItem "数据结构"
    List1.AddItem "软件工程"
    List1.AddItem "数据库原理"
    List1.AddItem "VB程序设计"
    List1.AddItem "计算机文件基础"
End Sub
