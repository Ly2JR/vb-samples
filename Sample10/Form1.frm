VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "加法计算器"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "="
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Text1_LostFocus()
    If Not IsNumeric(Text1.Text) Then
        MsgBox "您输入了非法字符"
        Text1.Text = ""
        Text1.SetFocus
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not IsNumeric(Text2.Text) Then
            MsgBox "您输入了非数字字符"
            Text2.Text = ""
        End If
    End If
End Sub

Private Sub Text3_GotFocus()
    Text3.Text = Val(Text1.Text) + Val(Text2.Text)
End Sub
