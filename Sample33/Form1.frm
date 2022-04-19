VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6984
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   11604
   LinkTopic       =   "Form1"
   ScaleHeight     =   6984
   ScaleWidth      =   11604
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "当前坐标"
      Height          =   372
      Left            =   9960
      TabIndex        =   1
      Top             =   6360
      Width           =   1332
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   6132
      Left            =   0
      ScaleHeight     =   6084
      ScaleWidth      =   11556
      TabIndex        =   0
      Top             =   0
      Width           =   11604
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
    Randomize
    For i = 1 To 100
        Picture1.ForeColor = QBColor(Int(Rnd * 16))
        Picture1.CurrentX = Picture1.ScaleWidth * Rnd
        Picture1.CurrentY = Picture1.ScaleHeight * Rnd
        Picture1.Print "*"
    Next i
End Sub
