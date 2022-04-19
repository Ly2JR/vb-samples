VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "字体演示"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "删除线"
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "倾斜"
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "字号"
      Height          =   615
      Left            =   1320
      TabIndex        =   2
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "黑体"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "字体设置演示"
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
    Label1.FontName = "黑体"
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
