VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3336
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   6828
   LinkTopic       =   "Form3"
   ScaleHeight     =   3336
   ScaleWidth      =   6828
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      Height          =   1812
      Left            =   240
      ScaleHeight     =   1764
      ScaleWidth      =   6324
      TabIndex        =   3
      Top             =   720
      Width           =   6372
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   372
      Left            =   5280
      TabIndex        =   2
      Top             =   2760
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "统计"
      Height          =   372
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6372
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim aph(1 To 26) As Integer
Dim i%, length%, j%, c As String * 1
length = Len(Text1.Text)
For i = 1 To length
    c = UCase(Mid(Text1.Text, i, 1))
    If c >= "A" And c <= "Z" Then   '大写字符A的ASC码为65
        j = Asc(c) - 64
        aph(j) = aph(j) + 1
    End If
Next i
Picture1.Cls
j = 0
For i = 1 To 26
    If aph(i) <> 0 Then
        Picture1.Print Spc(2); Chr(i + 64); "="; Format(aph(i), "@@");
        j = j + 1
        If j Mod 8 = 0 Then Picture1.Print      '每行输出8 个
    End If
Next
End Sub

Private Sub Command2_Click()
    End
End Sub

