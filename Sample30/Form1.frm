VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "线宽演示"
   ClientHeight    =   2940
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   5808
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   5808
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "线宽演示"
      Height          =   420
      Left            =   4800
      TabIndex        =   0
      Top             =   2400
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
    Dim i As Integer
    Form1.CurrentX = 0
    Form1.CurrentY = ScaleHeight / 2
    Form1.FillColor = QBColor(0)
    For i = 1 To 10
        Form1.DrawWidth = i * 3
        Form1.Line -Step(ScaleWidth / 15, 0)
    Next i
End Sub
