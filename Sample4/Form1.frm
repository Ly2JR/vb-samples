VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MouseMove事件演示"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Y坐标"
      Height          =   180
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "X坐标"
      Height          =   180
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   450
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1.Text = X
    Text2.Text = Y
End Sub
