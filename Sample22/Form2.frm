VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "ѡ����ȤС��"
   ClientHeight    =   2340
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   4056
   LinkTopic       =   "Form2"
   ScaleHeight     =   2340
   ScaleWidth      =   4056
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "ע��"
      Height          =   372
      Left            =   2640
      TabIndex        =   5
      Top             =   1440
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������һ��"
      Height          =   372
      Left            =   2640
      TabIndex        =   4
      Top             =   480
      Width           =   1212
   End
   Begin VB.Frame Frame1 
      Caption         =   "��ѡ��Ҫ�������ȤС��"
      Height          =   1932
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2412
      Begin VB.OptionButton Option3 
         Caption         =   "��Ϸ����"
         Height          =   300
         Left            =   480
         TabIndex        =   3
         Top             =   1320
         Width           =   1572
      End
      Begin VB.OptionButton Option2 
         Caption         =   "��վ���"
         Height          =   300
         Left            =   480
         TabIndex        =   2
         Top             =   840
         Width           =   1572
      End
      Begin VB.OptionButton Option1 
         Caption         =   "����ý��"
         Height          =   300
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   1572
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Me.Hide
    Form1.Show
    
End Sub

Private Sub Command2_Click()
    MsgBox "��ӭ" + Form1.Text1.Text + "����", , "ע��ɹ�"
    End
End Sub
