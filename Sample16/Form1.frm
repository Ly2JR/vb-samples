VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Timer控件示例"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "时间"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   2280
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3720
      Top             =   2400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "倒计时:"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "定时时间(分)"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim t As Integer

Private Sub Command1_Click()
    If Command1.Caption = "停止" Then
        Timer1.Enabled = False
        
        Command1.Caption = "开始"
    Else
        Timer1.Enabled = True
        Command1.Caption = "停止"
        t = 60 * Val(Text1.Text)
        Timer1.Enabled = True
    End If
End Sub

Private Sub Command2_Click()
'    Dim a As Date
'    a = #9/10/2000#
'    Print a

'Dim b As Boolean
'b = 6 > 8
'Print b

'Dim s As String * 5
's = "23gfrewrq"
'Print s

'Dim a As Double
'a = 12.2345
'Dim b As Integer
'b = 12
'Print "a="; Format(a, "0.00"); "b="; Format(b, "0.00")
'Print "a=" & Format(a, "#.##") & "b=" & Format(b, "#.##")


Print
Print Tab(15); "变量必须先声明后使用,这是为声明?"
Print Tab(15); "素数是" & "5", "7", "11"
Print Tab(15); "素数是" & "5"; "7"; "11"
End Sub

Private Sub Timer1_Timer()
    Dim m, s As Integer
    t = t - 1
    m = Int(t / 60)
    s = t Mod 60
    Label3.Caption = m & "分" & s & "秒"
    If t = 0 Then
        Timer1.Enabled = False
        Beep
        MsgBox "时间到"
        Command1.Caption = "开始"
    End If
End Sub
