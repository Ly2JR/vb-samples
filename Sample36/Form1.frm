VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6192
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   11796
   LinkTopic       =   "Form1"
   ScaleHeight     =   6192
   ScaleWidth      =   11796
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
 Dim i As Integer
    Dim x As Single
    Dim y As Single
    Dim colorCode As Integer
    Scale (-320, 240)-(320, -240)
    Line (-300, 220)-(300, -220), , B
    Randomize
    For i = 1 To 100
        x = 300 * Rnd
        If (Rnd < 0.5) Then x = -x
        y = 220 * Rnd
        If Rnd < 0.5 Then y = -y
        colorCode = 15 * Rnd
        Line (0, 0)-(x, y), QBColor(colorCode)
    Next i
End Sub

