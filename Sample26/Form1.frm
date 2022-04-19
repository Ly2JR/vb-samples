VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3084
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   5232
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3084
   ScaleWidth      =   5232
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   492
      Left            =   3960
      TabIndex        =   2
      Top             =   2520
      Width           =   1092
   End
   Begin VB.PictureBox Picture1 
      Height          =   1332
      Left            =   2520
      ScaleHeight     =   1284
      ScaleWidth      =   2484
      TabIndex        =   1
      Top             =   120
      Width           =   2532
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   840
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   360
      Top             =   360
   End
   Begin VB.Label Label1 
      Caption         =   "a"
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   2640
      Width           =   1092
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim n As Integer
Dim m As Integer
Private Sub Command1_Click()
    Picture1.Cls
    Timer1.Enabled = True
    Timer2.Enabled = True
    Command1.Enabled = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    m = m + 1
    If Chr(KeyAscii) = Label1.Caption Then
        Label1.Caption = ""
        n = n + 1
    End If
End Sub

Private Sub Timer1_Timer()
    Randomize
    If Label1.Caption = "" Then
        Label1.Top = Form1.Height - Label1.Height
        Label1.Caption = Chr(CInt(Rnd * 26 + 97))
    Else
        Label1.Top = Label1.Top - 10
    End If
    If Label1.Top <= 0 Then
        Label1.Top = Form1.Height - Label1.Height
    End If
End Sub

Private Sub Timer2_Timer()
    Timer1.Enabled = False
    Timer2.Enabled = False
    Picture1.Cls
    Picture1.Print "击键次数：" & m & "次"
    Picture1.Print "正确次数：" & n & "次"
    If m > 0 Then
        Picture1.Print "正确率为:" & n / m * 100&; "%"
        n = 0
        m = 0
        Command1.Enabled = True
    End If
End Sub
