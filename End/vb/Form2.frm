VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   LinkTopic       =   "Form2"
   ScaleHeight     =   3225
   ScaleWidth      =   5760
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "新增"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   240
      Width           =   5055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "修改"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "删除"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "查询"
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   780
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   5055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim affected As Integer
Dim sSql As String
sSql = "Insert into Test(content) values('" & Text1.Text & "')"
affected = CRUD(sSql)
MsgBox IIf(affected = 1, "新增成功", "新增失败")
End Sub

Private Sub Command2_Click()
Dim affected As Integer
Dim sSql As String
sSql = "update Test set content='" & Text1.Text & "' where id=1 "
affected = CRUD(sSql)
MsgBox IIf(affected = 1, "修改成功", "修改失败")
End Sub

Private Sub Command3_Click()
Dim affected As Integer
Dim sSql As String
sSql = "delete Test where id=1 "
affected = CRUD(sSql)
MsgBox IIf(affected = 1, "删除成功", "删除失败")
End Sub


Public Function CRUD(ByVal sSql As String) As Integer
Dim conn As New ADODB.Connection
Dim connString As String
Dim affected As Integer
CRUD = 0
connString = "Provider=SQLOLEDB.1;Data Source=DESKTOP-CETRBK3\SQLEXPRESS;Initial Catalog=master;Uid=sa;Password=sa123"
conn.Open connString

conn.Execute sSql, CRUD

conn.Close
Set conn = Nothing
End Function

Private Sub Command4_Click()

Dim conn As New ADODB.Connection
Dim connString As String
Dim sSql As String
Dim rest As New ADODB.Recordset
Dim test As New testCls
connString = "Provider=SQLOLEDB.1;Data Source=DESKTOP-CETRBK3\SQLEXPRESS;Initial Catalog=master;Uid=sa;Password=sa123"
conn.Open connString

sSql = "select * from test"

If rest.State = 1 Then rest.Close
rest.Open sSql, conn, adOpenStatic, adLockReadOnly

Set test.TestRecordSet = rest

List1.Clear

Do While Not test.TestRecordSet.EOF
    
    List1.AddItem test.TestDescription

    test.TestRecordSet.MoveNext
Loop

Set test = Nothing

conn.Close
Set conn = Nothing

End Sub
