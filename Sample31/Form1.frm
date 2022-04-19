VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "演示所画线的线型"
   ClientHeight    =   2784
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   8136
   LinkTopic       =   "Form1"
   ScaleHeight     =   2784
   ScaleWidth      =   8136
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "线型演示"
      Height          =   372
      Left            =   6840
      TabIndex        =   0
      Top             =   2280
      Width           =   972
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
    Print "DrawStle     0       1       2       3       4       5       6"
    Print
    Print "线型        实线   长虚线   点线   点画线  点点画线  透明线  内实线"
    Print
    Print "图示   ";
    CurrentY = CurrentY + 200
    DrawWidth = 1
    For i = 0 To 6
        DrawStyle = i
        CurrentX = CurrentX + 150
        Line -Step(700, 0)
        
    Next i
    
End Sub
