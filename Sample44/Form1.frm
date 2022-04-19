VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5136
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   8628
   LinkTopic       =   "Form1"
   ScaleHeight     =   5136
   ScaleWidth      =   8628
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   612
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1812
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    DataReport1.Show
End Sub
