VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   2520
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   3624
   LinkTopic       =   "Form3"
   ScaleHeight     =   2520
   ScaleWidth      =   3624
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command4 
      Caption         =   "�����Ի���"
      Height          =   372
      Left            =   2280
      TabIndex        =   4
      Top             =   2040
      Width           =   1212
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��ӡ�Ի���"
      Height          =   372
      Left            =   2280
      TabIndex        =   3
      Top             =   1440
      Width           =   1212
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����Ի���"
      Height          =   372
      Left            =   2280
      TabIndex        =   2
      Top             =   840
      Width           =   1212
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ɫ�Ի���"
      Height          =   372
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   1212
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��������"
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
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1200
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    CommonDialog1.Action = 3
    Label1.ForeColor = CommonDialog1.Color
End Sub

Private Sub Command2_Click()
    'cdlCFScreenFont ����ʾ��Ļ����
    'cdlCFPrinterFonts:��ʾ��ӡ������
    'cdCFBoth:��ʾ��ӡ���������Ļ����
    'cdCFEffects��������Ի�����ʾɾ���ߺ��»��߸�ѡ����ɫ�б��
    CommonDialog1.Flags = cdlCFBoth Or cdlCFEffects
    CommonDialog1.Action = 4
    Label1.FontName = CommonDialog1.FontName
    Label1.FontSize = CommonDialog1.FontSize
    Label1.FontBold = CommonDialog1.FontBold
    Label1.FontItalic = CommonDialog1.FontItalic
    Label1.FontStrikethru = CommonDialog1.FontStrikethru
    Label1.FontUnderline = CommonDialog1.FontUnderline
End Sub

Private Sub Command3_Click()
    CommonDialog1.Action = 5
End Sub

Private Sub Command4_Click()
    CommonDialog1.HelpFile = "winhelp32.hlp"
    CommonDialog1.HelpCommand = cdlHelpContents
    CommonDialog1.ShowHelp
End Sub
