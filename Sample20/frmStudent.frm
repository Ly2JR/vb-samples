VERSION 5.00
Begin VB.Form frmStudent 
   Caption         =   "ѧ����Ϣ����"
   ClientHeight    =   3084
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   7044
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3084
   ScaleWidth      =   7044
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command4 
      Caption         =   "�˳�"
      Height          =   372
      Left            =   5880
      TabIndex        =   12
      Top             =   2640
      Width           =   1092
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "��ѯ"
      Height          =   372
      Left            =   3120
      TabIndex        =   11
      Top             =   2640
      Width           =   1092
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ɾ��"
      Height          =   372
      Left            =   1680
      TabIndex        =   10
      Top             =   2640
      Width           =   1092
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����"
      Height          =   372
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   972
   End
   Begin VB.ListBox lstStudent 
      Height          =   1848
      Left            =   2640
      TabIndex        =   8
      Top             =   600
      Width           =   4332
   End
   Begin VB.TextBox txtScore 
      Height          =   372
      Left            =   960
      TabIndex        =   7
      Top             =   2040
      Width           =   1332
   End
   Begin VB.Frame Frame1 
      Caption         =   "�Ա�"
      Height          =   612
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   1932
      Begin VB.OptionButton Option2 
         Caption         =   "Ů"
         Height          =   300
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   492
      End
      Begin VB.OptionButton optMale 
         Caption         =   "��"
         Height          =   300
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   492
      End
   End
   Begin VB.TextBox txtName 
      Height          =   372
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   1332
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "�ɼ���"
      Height          =   180
      Left            =   360
      TabIndex        =   6
      Top             =   2160
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "����:"
      Height          =   180
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   456
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ѧ����Ϣ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   1800
   End
End
Attribute VB_Name = "frmStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDelete_Click()
    Dim k As Integer, mname As String
    k = lstStudent.ListIndex 'ȡ���б���ѡ��������
    If (k < 0) Then
        MsgBox "��û��ѡ��Ҫɾ���Ķ���!", vbOKOnly, "��ʾ��"
    Else
        mname = RTrim(LTrim(stuInf(k).name))
        lstStudent.RemoveItem k
        Call delete_from_stuInf(k)
        MsgBox mname & "�ѱ�ɾ��!", vbOKOnly, "��ʾ��"
    End If
    cmdDelete.Enabled = False
End Sub

Private Sub cmdSave_Click()
    Dim mname  As String, msex As String
    Dim mscore As Single
    mname = LTrim(RTrim(txtName.Text))
    txtName.Text = mname
    If (Len(mname) = 0) Then
        MsgBox "��������Ϊ��!����������", vbOKOnly, "��ʾ"
        txtName.SetFocus
        Exit Sub
    End If
    mname = name_d_str(mname)
    msex = IIf(optMale.Value, "��", "Ů")
    mscore = Val(txtScore.Text)
    Call save_to_stuInf(mname, msex, mscore)
    lstStudent.AddItem mname & " " & msex & " " & " " & mscore
    txtName.Text = ""
    optMale.Value = True
    txtScore.Text = ""
    txtName.SetFocus
    cmdDelete.Enabled = False
End Sub

Private Sub cmdSort_Click()
    cmdDelete.Enabled = False
    frmStudentDisp.Show 1, frmStudent '1-ģ̬������ʽ,frmStudent-������
End Sub

Private Sub Command4_Click()
    End
End Sub

Private Sub Form_Load()
    cmdDelete.Enabled = False
    Dim d(10) As Student
    Dim i As Integer
    d(0).name = "����"
    d(0).sex = "Ů"
    d(0).score = 78
    
    d(1).name = "����"
    d(1).sex = "��"
    d(1).score = 74
    
    d(2).name = "����"
    d(2).sex = "��"
    d(2).score = 82
    
    d(3).name = "���Ĺ�"
    d(3).sex = "��"
    d(3).score = 88
    
    d(4).name = "����"
    d(4).sex = "Ů"
    d(4).score = 79
    
    d(5).name = "�Ż�"
    d(5).sex = "��"
    d(5).score = 81
    
    d(6).name = "��ѧ��"
    d(6).sex = "��"
    d(6).score = 54
    
    d(7).name = "������"
    d(7).sex = "Ů"
    d(7).score = 68
    
    d(8).name = "��Ϊ��"
    d(8).sex = "��"
    d(8).score = 73
    
    d(9).name = "���в�"
    d(9).sex = "��"
    d(9).score = 92
    For i = 0 To UBound(d) - 1
         Call save_to_stuInf(name_d_str(d(i).name), d(i).sex, d(i).score)
         lstStudent.AddItem name_d_str(d(i).name) & " " & d(i).sex & " " & " " & d(i).score
    Next i
End Sub

Private Sub lstStudent_Click()
    cmdDelete.Enabled = True
End Sub


