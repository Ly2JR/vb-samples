VERSION 5.00
Begin VB.Form frmStudent 
   Caption         =   "Student"
   ClientHeight    =   3276
   ClientLeft      =   1164
   ClientTop       =   432
   ClientWidth     =   5592
   LinkTopic       =   "Form2"
   ScaleHeight     =   3276
   ScaleWidth      =   5592
   Begin VB.CommandButton Command1 
      Caption         =   "报表(&P)"
      Height          =   300
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&C)"
      Height          =   300
      Left            =   4440
      TabIndex        =   16
      Top             =   1980
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "更新(&U)"
      Height          =   300
      Left            =   3360
      TabIndex        =   15
      Top             =   1980
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "刷新(&R)"
      Height          =   300
      Left            =   2280
      TabIndex        =   14
      Top             =   1980
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除(&D)"
      Height          =   300
      Left            =   1200
      TabIndex        =   13
      Top             =   1980
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "添加(&A)"
      Height          =   300
      Left            =   120
      TabIndex        =   12
      Top             =   1980
      Width           =   975
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\98247\Desktop\VBer\Sample44\Sample44.mdb"
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   324
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Student"
      Top             =   2952
      Width           =   5592
   End
   Begin VB.TextBox txtFields 
      DataField       =   "奖学金"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   11
      Top             =   1640
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "所在系"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   4
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   9
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "出生日期"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1000
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "性别"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   2
      Left            =   2040
      MaxLength       =   2
      TabIndex        =   5
      Top             =   680
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "姓名"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   3
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "学号"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   2040
      MaxLength       =   5
      TabIndex        =   1
      Top             =   40
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Caption         =   "奖学金:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "所在系:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "出生日期:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "性别:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "姓名:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "学号:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
  Data1.Recordset.AddNew
End Sub

Private Sub cmdDelete_Click()
  '如果删除记录集的最后一条记录
  '记录或记录集中唯一的记录
  Data1.Recordset.Delete
  Data1.Recordset.MoveNext
End Sub

Private Sub cmdRefresh_Click()
  '这仅对多用户应用程序才是需要的
  Data1.Refresh
End Sub

Private Sub cmdUpdate_Click()
  Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Command1_Click()
    DataReport1.Show
End Sub

Private Sub Data1_Error(DataErr As Integer, Response As Integer)
  '这就是放置错误处理代码的地方
  '如果想忽略错误，注释掉下一行代码
  '如果想捕捉错误，在这里添加错误处理代码
  MsgBox "数据错误事件命中错误：" & Error$(DataErr)
  Response = 0  '忽略错误
End Sub

Private Sub Data1_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  '这将显示当前记录位置
  '为动态集和快照
  Data1.Caption = "记录：" & (Data1.Recordset.AbsolutePosition + 1)
  '对于 Table 对象，当记录集创建后并使用下面的行时，
  '必须设置 Index 属性
  'Data1.Caption = "记录：" & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
  '这是放置验证代码的地方
  '当下面的动作发生时，调用这个事件
  Select Case Action
    Case vbDataActionMoveFirst
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
  End Select
  Screen.MousePointer = vbHourglass
End Sub

