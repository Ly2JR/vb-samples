VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   4200
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command6 
      Caption         =   "ÍË³ö"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Çå³ý"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "É¾³ý"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Õ³Ìù"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "¼ôÇÐ"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "¸´ÖÆ"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   4335
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Clipboard.Clear
    Clipboard.SetText Text1.SelText
End Sub


Private Sub Command2_Click()
    Clipboard.Clear
    Clipboard.SetText Text1.SelText
    Text1.SelText = ""
End Sub

Private Sub Command3_Click()
    Text1.SelText = Clipboard.GetText
End Sub

Private Sub Command4_Click()
    Text1.SelText = ""
End Sub

Private Sub Command5_Click()
    Text1.Text = ""
End Sub

Private Sub Command6_Click()
    End
End Sub
