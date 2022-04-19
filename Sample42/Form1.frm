VERSION 5.00
Begin VB.Form frmMoving 
   Caption         =   "�򵥶�����ʾ"
   ClientHeight    =   3780
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   6144
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   6144
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer tmrMov 
      Enabled         =   0   'False
      Left            =   2400
      Top             =   600
   End
   Begin VB.CommandButton cmdSS 
      Caption         =   "��ʼ"
      Height          =   372
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   852
   End
End
Attribute VB_Name = "frmMoving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const PAI As Double = 3.1415927
Dim m_a As Double
Dim m_b As Double
Dim m_r As Double


Private Sub cmdSS_Click()
    If cmdSS.Caption = "��ʼ" Then
        tmrMov.Enabled = True
        tmrMov.Interval = 20
        cmdSS.Caption = "ֹͣ"
    Else
        tmrMov.Enabled = False
        cmdSS.Caption = "��ʼ"
    End If
End Sub

Private Sub Form_Activate()
    Dim myes As Boolean
    myes = set_coorsSys(frmMoving)   '���ô���frmMoving������ϵͳ
    If Not myes Then Exit Sub       '����ʧ�����˳�
    Call set_drawEnv(frmMoving)     '���û�ͼ����
    Call tmrMov_Timer               '���˶���Բ
End Sub


'���û�ͼ���Թ���
Private Sub set_drawEnv(obj As Object)
m_r = IIf(m_a >= m_b, m_b, m_a) / 10# '���㻭Բ�İ뾶
obj.DrawMode = 13               '���û�ͼģʽΪ13(CopyPen)
obj.FillStyle = 0               'Ϊ����Բ�������ģʽ
obj.FillColor = vbRed           '�����ɫ
obj.Circle (0, 0), m_r + m_r, vbRed '��Բ
obj.FillStyle = 1               '͸��
If m_a >= m_b Then
    obj.Circle (0, 0), m_a, vbBlue, , , m_b / m_a '����Բ���
Else
    obj.Circle (0, 0), m_b, vbBlue, , , m_b / m_a '����Բ���
End If
obj.DrawMode = 7            '����ͼ��ģʽΪ7
obj.FillStyle = 0           'ΪСԲ�������ģʽ
obj.FillColor = vbRed       'ΪСԲ���������ɫ
End Sub




'��������ϵͳ
Public Function set_coorsSys(obj As Object) As Boolean
    Dim xMin As Double, xMax As Double
    Dim yMin As Double, yMax As Double
    Dim rate    As Double
    If obj.ScaleHeight <= 0# Or obj.ScaleWidth <= 0# Then
        MsgBox "����(�細��)����߸����ò�����!"
        set_coorsSys = False
        Exit Function
    End If
    
    rate = obj.ScaleHeight / obj.ScaleWidth     '�������߿�֮��
    xMin = -1000#                                   '�趨x�᷽��2000
    xMax = 1000#
    m_a = 800#          '��ԲX��뾶800
    yMax = 1000# * rate         '����������y���򳤶�
    yMin = -yMax
    m_b = m_a * rate    '��Բy�᷽��뾶�ĳ���
    obj.Cls
    obj.Scale (xMin, yMax)-(xMax, yMin)         '��������ϵ
    set_coorsSys = True
End Function

Private Sub tmrMov_Timer()
    Static angle As Integer, ag As Boolean
    Dim x As Double, y As Double, t As Double
    ag = Not ag
    If ag Then angle = angle + 2
    If angle > 360 Then angle = 0
    t = angle * PAI / 180#
    x = m_a * Cos(t)
    y = m_b * Sin(t)
    frmMoving.Circle (x, y), m_r, vbGreen
    
End Sub
