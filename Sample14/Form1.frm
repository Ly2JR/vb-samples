VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Image控件示例"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   3465
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "拉伸"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "压缩"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   4560
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Stretch"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   4335
      Left            =   0
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim H, W As Integer

Private Sub Check1_Click()
    Image1.Stretch = Check1.Value
End Sub

Private Sub Command1_Click()
    If Image1.Height < 500 Then
        Command1.Enabled = False
    Else
        Image1.Height = Image1.Height - 100
        Command2.Enabled = True
    End If
End Sub

Private Sub Command2_Click()
     If Image1.Height > 4500 Then
        Command2.Enabled = False
    Else
        Image1.Height = Image1.Height + 100
    End If
End Sub

Private Sub Form_Load()
    Image1.Picture = LoadPicture(App.Path + "\img_3.jpg")
    
    H = Image1.Height
    W = Image1.Width
    Check1.Value = Image1.Stretch
End Sub


