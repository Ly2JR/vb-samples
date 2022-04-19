VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3636
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   5988
   LinkTopic       =   "Form1"
   ScaleHeight     =   3636
   ScaleWidth      =   5988
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    Dim i As Integer
    Dim j As Integer
    Cls '清除窗体
    Scale (-120, 120)-(120, -120)   '远点在窗体中心，X轴向右、Y轴向上
    DrawWidth = 2
    Line (-115, 0)-(115, 0) 'X轴
    Line (110, 4)-(115, 0)
    Line (110, -4)-(115, 0)
    CurrentX = 110: CurrentY = 20: Print "X" '显示X轴坐标
    Line (0, -115)-(0, 115) '画Y轴
    Line (-2, 105)-(0, 115)
    Line (2, 105)-(0, 115)
    CurrentX = 5: CurrentY = 118: Print "Y" '显示Y轴
    For i = -100 To 100 Step 20 'X轴坐标刻度
        If i <> 0 Then
            Line (i, 5)-(i, 0)
            CurrentX = i - 7: CurrentY = -5: Print i
        Else
            CurrentX = 3: CurrentY = -1: Print "0"
        End If
    Next i
    For j = -100 To 100 Step 20
        If j <> 0 Then
            Line (0, j)-(2, j)
            CurrentX = 5: CurrentY = j + 8: Print j
        End If
    Next j
End Sub

