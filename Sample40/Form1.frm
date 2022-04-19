VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4992
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   11676
   LinkTopic       =   "Form1"
   ScaleHeight     =   4992
   ScaleWidth      =   11676
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Shape Shape1 
      Height          =   252
      Index           =   0
      Left            =   360
      Top             =   480
      Width           =   732
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    Dim i As Integer
    Dim x As Integer, y As Integer
    y = Shape1(0).Top
    Shape1(0).Shape = 0 '0¾ØÐÎ
    Shape1(0).BorderWidth = 2
    For i = 1 To 5
        Load Shape1(i)
        Shape1(i).Shape = i
        x = Shape1(i - 1).Left + Shape1(i - 1).Width + 200
        Shape1(i).Left = x
        Shape1(i).Top = y
        Shape1(i).Visible = True
    Next i
End Sub
