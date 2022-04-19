VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5808
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   11232
   LinkTopic       =   "Form1"
   ScaleHeight     =   5808
   ScaleWidth      =   11232
   StartUpPosition =   3  '¥∞ø⁄»± °
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    Scale (-320, 240)-(320, -240)
    Circle (-100, 120), 30
    CurrentX = -180
    CurrentY = 0
    Print "ª≠‘≤Circle(-100,120),30"
    
    Scale (-500, 180)-(500, -180)
    Circle (80, 80), 40, , -0.0001, -1.57
    CurrentX = 0
    CurrentY = 0
    Print "ª≠…»–ŒCircle(80,80),30,,-0.0001,-1.57"
    
    
    Scale (-320, 380)-(320, -380)
    Circle (-100, -180), 30, , , , 0.5
    CurrentX = -180
    CurrentY = -240
    Print "ª≠Õ÷‘≤Circle(-100,-180),30,,,,0.5"
    
    Scale (-200, -180)-(200, 180)
    Circle (80, 80), 30, , -2.1, 0.7
    CurrentX = 0
    CurrentY = 140
    Print "ª≠‘≤ª°Circle(80,80),30,,-2.1,0.7"
End Sub

