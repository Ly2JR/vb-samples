VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2196
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   4524
   LinkTopic       =   "Form2"
   ScaleHeight     =   2196
   ScaleWidth      =   4524
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Visual Basic 6.0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   2496
   End
   Begin VB.Menu PopLabel 
      Caption         =   "标签"
      Visible         =   0   'False
      Begin VB.Menu mnuColor 
         Caption         =   "红色"
      End
      Begin VB.Menu mnuFont 
         Caption         =   "黑体"
      End
      Begin VB.Menu mnuSize 
         Caption         =   "20号"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu PopLabel
    End If
End Sub

Private Sub mnuColor_Click()
    Label1.ForeColor = QBColor(12)
End Sub

Private Sub mnuFont_Click()
    Label1.FontName = "黑体"
End Sub

Private Sub mnuSize_Click()
    Label1.FontSize = 20
End Sub
