VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "模拟注册"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   5064
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3036
   ScaleWidth      =   5064
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   372
      Left            =   3240
      TabIndex        =   9
      Top             =   2040
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Caption         =   "下一步"
      Height          =   372
      Left            =   3240
      TabIndex        =   8
      Top             =   1080
      Width           =   972
   End
   Begin VB.Frame Frame1 
      Caption         =   "请输入您的资料"
      Height          =   1812
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   1932
      Begin VB.TextBox Text2 
         Height          =   264
         Left            =   600
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   1320
         Width           =   1092
      End
      Begin VB.ComboBox Combo1 
         Height          =   276
         ItemData        =   "Form1.frx":0000
         Left            =   600
         List            =   "Form1.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   840
         Width           =   1092
      End
      Begin VB.TextBox Text1 
         Height          =   264
         Left            =   600
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   360
         Width           =   1092
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "专业"
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "欢迎加入科技俱乐部"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Me.Hide
    Form2.Show
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Load()
    Combo1.AddItem "男"
    Combo1.AddItem "女"
End Sub
