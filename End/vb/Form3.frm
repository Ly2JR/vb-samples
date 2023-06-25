VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5745
   LinkTopic       =   "Form3"
   ScaleHeight     =   3015
   ScaleWidth      =   5745
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List1 
      Height          =   780
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   5055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "查询"
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "删除"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "修改"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "新增"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   975
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim affected As Integer
Dim sSql As String
sSql = "Insert into Test(content) values('" & Text1.Text & "')"
affected = G_CRUD(sSql)
MsgBox IIf(affected = 1, "新增成功", "新增失败")
End Sub

Private Sub Command2_Click()
Dim affected As Integer
Dim sSql As String
sSql = "update Test set content='" & Text1.Text & "' where id=1 "
affected = G_CRUD(sSql)
MsgBox IIf(affected = 1, "修改成功", "修改失败")
End Sub

Private Sub Command3_Click()
Dim affected As Integer
Dim sSql As String
sSql = "delete Test where id=1 "
affected = G_CRUD(sSql)
MsgBox IIf(affected = 1, "删除成功", "删除失败")
End Sub


Private Sub Command4_Click()
Dim rest As ADODB.Recordset
Set rest = G_Query("select * from test")

List1.Clear

Do While Not rest.EOF
    
    List1.AddItem rest!content

    rest.MoveNext
Loop

rest.Close
Set rest = Nothing

End Sub

Private Sub Form_Terminate()
CloseDb
End Sub
