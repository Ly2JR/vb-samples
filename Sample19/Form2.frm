VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3288
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   5652
   LinkTopic       =   "Form2"
   ScaleHeight     =   3288
   ScaleWidth      =   5652
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "最高"
      Height          =   372
      Left            =   4680
      TabIndex        =   11
      Top             =   2520
      Width           =   732
   End
   Begin VB.CommandButton Command3 
      Caption         =   "后一个"
      Height          =   372
      Left            =   4680
      TabIndex        =   8
      Top             =   1800
      Width           =   732
   End
   Begin VB.TextBox Text3 
      Height          =   372
      Left            =   960
      TabIndex        =   7
      Top             =   1800
      Width           =   3612
   End
   Begin VB.CommandButton Command2 
      Caption         =   "前一个"
      Height          =   372
      Left            =   4680
      TabIndex        =   5
      Top             =   1200
      Width           =   732
   End
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   960
      TabIndex        =   4
      Top             =   1200
      Width           =   3612
   End
   Begin VB.CommandButton Command1 
      Caption         =   "新增"
      Height          =   372
      Left            =   4680
      TabIndex        =   2
      Top             =   480
      Width           =   732
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   3612
   End
   Begin VB.Label Label5 
      Caption         =   "0/0"
      Height          =   252
      Left            =   1080
      TabIndex        =   10
      Top             =   2520
      Width           =   3492
   End
   Begin VB.Label Label4 
      Caption         =   "位置"
      Height          =   252
      Left            =   360
      TabIndex        =   9
      Top             =   2520
      Width           =   612
   End
   Begin VB.Label Label3 
      Caption         =   "总分"
      Height          =   252
      Left            =   360
      TabIndex        =   6
      Top             =   1800
      Width           =   492
   End
   Begin VB.Label Label2 
      Caption         =   "专业"
      Height          =   252
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   492
   End
   Begin VB.Label Label1 
      Caption         =   "姓名"
      Height          =   252
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   492
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim n%, i%      '总个数、当前位置
Dim stud(100) As StudType

Private Sub Command1_Click()
   
    If n < 100 Then
        If Text1.Text <> "" Then n = n + 1
        
    Else
        MsgBox prompt:="人数已经达到100了"
        Exit Sub
    End If
    If n = 0 Or n = i Then Exit Sub
    i = n
    With stud(n)
        .Name = Text1.Text
        .Special = Text2.Text
        .Total = Val(Text3.Text)
    End With
    Text1.Text = "": Text2.Text = "": Text3.Text = ""
    Label5.Caption = CStr(i) & "/" & CStr(n)
End Sub

Private Sub Command2_Click()
    If i = 0 Then Exit Sub
    If i > 1 Then i = i - 1
    With stud(i)
        Text1.Text = .Name
        
        Text2.Text = .Special
        
        Text3.Text = .Total

    End With
    Label5.Caption = CStr(i) & "/" & CStr(n)
End Sub

Private Sub Command3_Click()
    If i = 0 Then Exit Sub
    If i < n Then i = i + 1
    With stud(i)
        Text1.Text = .Name
        Text2.Text = .Special
        Text3.Text = .Total
    End With
     Label5.Caption = CStr(i) & "/" & CStr(n)
End Sub

Private Sub Command4_Click()
 Dim max%, maxi%, j%
    If n = 0 Then Exit Sub
    max = stud(1).Total
    maxi = 1
    For j = 2 To n
        If stud(j).Total > max Then
            max = stud(j).Total
            maxi = j
        End If
    Next j
    With stud(maxi)
        Text1.Text = .Name
        Text2.Text = .Special
        
        Text3.Text = .Total
    End With
    i = maxi
 Label5.Caption = CStr(i) & "/" & CStr(n)
End Sub

