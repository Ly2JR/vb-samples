VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   5730
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List1 
      Height          =   780
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   5055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "删除"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "修改"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "新增"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()

Dim conn As New ADODB.Connection
Dim connString As String
Dim sSql As String
Dim affected As Integer
connString = "Provider=SQLOLEDB.1;Data Source=DESKTOP-CETRBK3\SQLEXPRESS;Initial Catalog=master;Uid=sa;Password=sa123"
conn.Open connString

sSql = "Insert into Test(content) values('" & Text1.Text & "')"
conn.Execute sSql, affected

conn.Close
Set conn = Nothing

MsgBox IIf(affected = 1, "新增成功", "新增失败")

End Sub

Private Sub Command2_Click()

Dim conn As New ADODB.Connection
Dim connString As String
Dim sSql As String
Dim affected As Integer
connString = "Provider=SQLOLEDB.1;Data Source=DESKTOP-CETRBK3\SQLEXPRESS;Initial Catalog=master;Uid=sa;Password=sa123"
conn.Open affected

sSql = "update Test set content='" & Text1.Text & "' where id=1 "
conn.Execute sSql, affected

conn.Close
Set conn = Nothing

MsgBox IIf(affected = 1, "修改成功", "修改失败")

End Sub

Private Sub Command3_Click()

Dim conn As New ADODB.Connection
Dim connString As String
Dim sSql As String
Dim affected As Integer
connString = "Provider=SQLOLEDB.1;Data Source=DESKTOP-CETRBK3\SQLEXPRESS;Initial Catalog=master;Uid=sa;Password=sa123"
conn.Open connString

sSql = "delete Test where id=1 "
conn.Execute sSql, affected

conn.Close
Set conn = Nothing

MsgBox IIf(affected = 1, "删除成功", "删除失败")

End Sub

