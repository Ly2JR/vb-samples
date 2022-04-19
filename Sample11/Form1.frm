VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "单选按钮、复选按钮和框架演示"
   ClientHeight    =   3420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   4545
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox Check2 
      Caption         =   "倾斜"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   2160
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Caption         =   "字号"
      Height          =   1935
      Left            =   3120
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
      Begin VB.OptionButton Option5 
         Caption         =   "30"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   615
      End
      Begin VB.OptionButton Option6 
         Caption         =   "40"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   615
      End
      Begin VB.OptionButton Option4 
         Caption         =   "20"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "字型"
      Height          =   1935
      Left            =   1800
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
      Begin VB.CheckBox Check3 
         Caption         =   "下划线"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "加粗"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "字体"
      Height          =   1935
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
      Begin VB.OptionButton Option2 
         Caption         =   "黑体"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "楷体"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "宋体"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Text            =   "Visual Basic 6.0"
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    Text1.FontBold = Check1.Value
End Sub

Private Sub Check2_Click()
    Text1.FontItalic = Check2.Value
End Sub

Private Sub Check3_Click()
    Text1.FontUnderline = Check3.Value
End Sub

Private Sub Option1_Click()
    Text1.FontName = "宋体"
End Sub

Private Sub Option2_Click()
    Text1.FontName = "黑体"
End Sub

Private Sub Option3_Click()
    Text1.FontName = "楷体"
End Sub

Private Sub Option4_Click()
    Text1.FontSize = 20
End Sub

Private Sub Option5_Click()
  Text1.FontSize = 30
End Sub

Private Sub Option6_Click()
  Text1.FontSize = 40
End Sub
