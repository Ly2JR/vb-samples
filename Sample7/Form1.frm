VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "命令按钮演示"
   ClientHeight    =   1635
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4320
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1635
   ScaleWidth      =   4320
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "退出(&X)"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "关闭(&C)"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打开窗体(&O)"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Load Form2
    Form2.Show
End Sub


Private Sub Command2_Click()
    Unload Form2
End Sub

Private Sub Command3_Click()
    End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Text1.Text = Chr(KeyAscii) & " - " & KeyAscii
End Sub

