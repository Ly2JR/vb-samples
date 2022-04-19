VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Í¼Ïñ¼ÓÔØ"
   ClientHeight    =   4275
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   3660
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
  Form1.Picture = LoadPicture(App.Path + "\img_3.jpg")
    Form1.Width = Form1.Picture.Width
    Form1.Height = Form1.Picture.Height
End Sub

Private Sub Form_DblClick()
Form1.Picture = Nothing

End Sub


