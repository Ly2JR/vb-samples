VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form4 
   Caption         =   "ͨ�öԻ���ʾ��"
   ClientHeight    =   3336
   ClientLeft      =   192
   ClientTop       =   816
   ClientWidth     =   5064
   LinkTopic       =   "Form4"
   ScaleHeight     =   3336
   ScaleWidth      =   5064
   StartUpPosition =   3  '����ȱʡ
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5520
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".txt"
      Filter          =   "Text Files(*.txt)|*.txt|All Filtes(*.*)|*.*"
   End
   Begin VB.TextBox Text1 
      Height          =   3012
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form4.frx":0000
      Top             =   240
      Width           =   4692
   End
   Begin VB.Menu mnuOpen 
      Caption         =   "��"
   End
   Begin VB.Menu mnuSaveAs 
      Caption         =   "���Ϊ"
   End
   Begin VB.Menu mnuColor 
      Caption         =   "��ɫ"
   End
   Begin VB.Menu mnuFont 
      Caption         =   "����"
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "��ӡ"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "�Ƴ�"
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Text1.Text = ""
    
End Sub

Private Sub mnuColor_Click()
    CommonDialog1.Action = 3
    Text1.ForeColor = CommonDialog1.Color
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuFont_Click()
    CommonDialog1.Flags = cdlCFBoth Or cdlCFEffects
    CommonDialog1.Action = 4
    Text1.FontName = CommonDialog1.FontName
    Text1.FontSize = CommonDialog1.FontSize
    Text1.FontBold = CommonDialog1.FontBold
    Text1.FontItalic = CommonDialog1.FontItalic
    Text1.FontStrikethru = CommonDialog1.FontStrikethru
    Text1.FontUnderline = CommonDialog1.FontUnderline
End Sub

Private Sub mnuOpen_Click()
    Dim inputdata As Variant
    CommonDialog1.InitDir = "c:\"
    CommonDialog1.Action = 1
    Open CommonDialog1.FileName For Input As #1     '���ļ���ȡ�ļ�����
    Do While Not EOF(1)
        Line Input #1, inputdata '��ȡ�ļ�����
        Text1.Text = Text1.Text + inputdata + vbCrLf
    Loop
    Close #1
End Sub

Private Sub mnuPrint_Click()
Dim i As Integer
    CommonDialog1.Action = 5
    For i = 1 To CommonDialog1.Copies
        Printer.Print Text1.Text            '��ӡ�ı�������
    Next
    Printer.EndDoc      '�����ĵ���ӡ
End Sub

Private Sub mnuSaveAs_Click()
    CommonDialog1.FileName = "Default.txt"
    CommonDialog1.DefaultExt = "txt"
    CommonDialog1.Action = 2
    Open CommonDialog1.FileName For Output As #1
    Print #1, Text1.Text
    Close #1
End Sub
