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
      Caption         =   "����(&P)"
      Height          =   300
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ر�(&C)"
      Height          =   300
      Left            =   4440
      TabIndex        =   16
      Top             =   1980
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "����(&U)"
      Height          =   300
      Left            =   3360
      TabIndex        =   15
      Top             =   1980
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "ˢ��(&R)"
      Height          =   300
      Left            =   2280
      TabIndex        =   14
      Top             =   1980
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ɾ��(&D)"
      Height          =   300
      Left            =   1200
      TabIndex        =   13
      Top             =   1980
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "���(&A)"
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
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
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
      DataField       =   "��ѧ��"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   11
      Top             =   1640
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "����ϵ"
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
      DataField       =   "��������"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1000
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "�Ա�"
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
      DataField       =   "����"
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
      DataField       =   "ѧ��"
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
      Caption         =   "��ѧ��:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "����ϵ:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "��������:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "�Ա�:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "����:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ѧ��:"
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
  '���ɾ����¼�������һ����¼
  '��¼���¼����Ψһ�ļ�¼
  Data1.Recordset.Delete
  Data1.Recordset.MoveNext
End Sub

Private Sub cmdRefresh_Click()
  '����Զ��û�Ӧ�ó��������Ҫ��
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
  '����Ƿ��ô��������ĵط�
  '�������Դ���ע�͵���һ�д���
  '����벶׽������������Ӵ��������
  MsgBox "���ݴ����¼����д���" & Error$(DataErr)
  Response = 0  '���Դ���
End Sub

Private Sub Data1_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  '�⽫��ʾ��ǰ��¼λ��
  'Ϊ��̬���Ϳ���
  Data1.Caption = "��¼��" & (Data1.Recordset.AbsolutePosition + 1)
  '���� Table ���󣬵���¼��������ʹ���������ʱ��
  '�������� Index ����
  'Data1.Caption = "��¼��" & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
  '���Ƿ�����֤����ĵط�
  '������Ķ�������ʱ����������¼�
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

