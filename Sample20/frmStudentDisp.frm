VERSION 5.00
Begin VB.Form frmStudentDisp 
   Caption         =   "显示与排序"
   ClientHeight    =   3564
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   6108
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3564
   ScaleWidth      =   6108
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出"
      Height          =   372
      Left            =   5160
      TabIndex        =   5
      Top             =   3000
      Width           =   852
   End
   Begin VB.CommandButton cmdSortByScore 
      Caption         =   "按成绩排序"
      Height          =   372
      Left            =   2880
      TabIndex        =   4
      Top             =   3000
      Width           =   1092
   End
   Begin VB.CommandButton cmdSortByName 
      Caption         =   "按姓名排序"
      Height          =   372
      Left            =   1440
      TabIndex        =   3
      Top             =   3000
      Width           =   1092
   End
   Begin VB.CommandButton cmdDispArray 
      Caption         =   "原始顺序"
      Height          =   372
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   852
   End
   Begin VB.ListBox lstStudent 
      Height          =   2208
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5892
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "按姓名排序"
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1500
   End
End
Attribute VB_Name = "frmStudentDisp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDispArray_Click()
    Dim i As Integer, k As Integer
    Dim str As String
    lstStudent.Clear
    k = UBound(stuInf)
    If (k < 0) Then
        MsgBox "目前尚未录入数据", vbOKOnly, "提示框"
        Exit Sub
    End If
    For i = 0 To k
        str = stuInf(i).name & " " & stuInf(i).sex & " " & stuInf(i).score
        lstStudent.AddItem str
    Next i
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSortByName_Click()
    Call sortByNameAndScore(1)
End Sub


Private Sub sortByNameAndScore(ByVal sele As Integer)
    Dim i As Integer, k As Integer
    Dim d() As Student
    lstStudent.Clear
    k = UBound(stuInf)
    If (k < 0) Then
        MsgBox "目前尚未录入数据!", vbOKOnly, "提示框"
        Exit Sub
    End If
    Call copy_stuinf_to_d(d)
    If (sele = 1) Then
        Call sort_d_by_name(d)
    ElseIf (sele = 2) Then
        Call sort_d_by_score(d)
    End If
    For i = 0 To k
        lstStudent.AddItem d(i).name & " " & d(i).sex & " " & d(i).score
    Next i
End Sub

Private Sub cmdSortByScore_Click()
  Call sortByNameAndScore(2)
End Sub
