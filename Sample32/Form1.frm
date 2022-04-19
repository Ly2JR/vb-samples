VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "演示填充样式"
   ClientHeight    =   2916
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   7704
   LinkTopic       =   "Form1"
   ScaleHeight     =   2916
   ScaleWidth      =   7704
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "填充演示"
      Height          =   372
      Left            =   6720
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
    Dim x As Integer, y As Integer
    Print "FillStyle    0       1       2       3       4      5    6       7"
    Print
    Print "图示 ";
    x = CurrentX
    y = CurrentY
    DrawWidth = 1
    For i = 0 To 7
        FillColor = QBColor(1)
        FillStyle = i
        Line (x, y)-(x + 500, y + 500), , B
        x = x + 600
    Next i
End Sub
