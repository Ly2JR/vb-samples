VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5556
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   ScaleHeight     =   5556
   ScaleWidth      =   9420
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        CurrentX = X
        CurrentY = Y
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim drawColor As Long
    If Button = 1 Then
        Select Case Shift
            Case 1
                    drawColor = vbRed
            Case 2
                drawColor = vbGreen
            Case 3
                drawColor = vbBlue
            Case Else
                drawColor = vbBlack
        End Select
         Line -(X, Y), drawColor
    End If
   
End Sub
