VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   3825
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "密码"
      Height          =   180
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "账号"
      Height          =   180
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If Text2.Text <> "Hello" Then
        If MsgBox("密码错误", vbRetryCancel + vbExclamation, "输入密码") = vbRetry Then
            Text2.Text = ""
            Text2.SetFocus
        Else
            End
        End If
    End If
End Sub

Private Sub Command2_Click()
    End
End Sub


Private Sub Text1_LostFocus()
    If Not IsNumeric(Text1.Text) Then
        MsgBox "账号有非数字字符错误"
        Text1.Text = ""
        Text1.SetFocus
    End If
End Sub
