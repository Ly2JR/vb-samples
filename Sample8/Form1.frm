VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "������֤"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command3 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��������"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "����"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "�û���"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tryCount As Integer
Private Sub Command1_Click()
    tryCount = tryCount + 1
    If (tryCount > 3) Then
        MsgBox "�Ƿ��û�"
        End
    End If
    
    If Trim(Text1.Text) = "" Or Trim(Text2.Text) = "" Then
        MsgBox "�û����������������"
        Exit Sub
    End If
    If Trim(Text1.Text) = "tg" And Trim(Text2.Text) = "123456" Then
        MsgBox "������ȷ����ӭʹ��"
    Else
        If Trim(Text1.Text) <> "tg" Then
            MsgBox "�û�������ȷ,����������"
            Text1.Text = ""
            Text1.SetFocus
        Else
            MsgBox "���벻��ȷ,����������"
            Text2.Text = ""
            Text2.SetFocus
        End If
    End If
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
End Sub

Private Sub Command3_Click()
    End
End Sub
