VERSION 5.00
Begin VB.Form frmMove 
   Caption         =   "文字移动"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdExit 
      Caption         =   "结束"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "向下移动"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdRiggt 
      Caption         =   "向右移动"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblMove 
      Caption         =   "欢迎使用Visual Basic"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "frmMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDown_Click()
    If lblMove.Top > frmMove.Height Then
        lblMove.Top = 0
    Else
        lblMove.Top = lblMove.Top + 50
    End If
    
End Sub

Private Sub cmdRiggt_Click()
    If lblMove.Left > frmMove.Width Then
        lblMove.Left = -frmMove.Width
        
    Else
        lblMove.Left = lblMove.Left + 50
    End If
End Sub
