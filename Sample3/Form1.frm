VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "������ʾ"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command4 
      Caption         =   "ɾ����"
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��б"
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�ֺ�"
      Height          =   615
      Left            =   1320
      TabIndex        =   2
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����������ʾ"
      Height          =   180
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Label1.FontName = "����"
End Sub

Private Sub Command2_Click()
    If (Label1.FontSize > 25) Then
        Label1.FontSize = Form1.FontSize
    Else
        Label1.FontSize = Label1.FontSize + 2
    End If
End Sub

Private Sub Command3_Click()
  Label1.FontItalic = True
    
End Sub

Private Sub Command4_Click()
      Label1.FontStrikethru = True
End Sub
