VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�������ߺ���������"
   ClientHeight    =   3264
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   3264
   ScaleWidth      =   9060
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picCos 
      Height          =   2652
      Left            =   4560
      ScaleHeight     =   2604
      ScaleWidth      =   4404
      TabIndex        =   3
      Top             =   0
      Width           =   4452
   End
   Begin VB.CommandButton cmdDrawCos 
      Caption         =   "����������"
      Height          =   372
      Left            =   5520
      TabIndex        =   2
      Top             =   2760
      Width           =   2052
   End
   Begin VB.CommandButton cmdDrawSin 
      Caption         =   "����������"
      Height          =   372
      Left            =   1560
      TabIndex        =   1
      Top             =   2760
      Width           =   2052
   End
   Begin VB.PictureBox picSin 
      Height          =   2652
      Left            =   0
      ScaleHeight     =   2604
      ScaleWidth      =   4524
      TabIndex        =   0
      Top             =   0
      Width           =   4572
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const PAI As Single = 3.1415297


'����������
Private Sub cmdDrawCos_Click()
Dim angle As Integer
Dim x1 As Single, y1 As Single
Dim x2 As Single, y2 As Single
If cmdDrawCos.Caption = "����������" Then
    Call draw_axis(picCos)      '���ò���ͼƬ��picCos��������
    x1 = 0# * PAI / 180#        '��ת��Ϊ����
    y1 = Cos(x1)                '����cos(x)
    For angle = 0 To 360 Step 1 '��0�ȵ�360�Ȼ���������
        x2 = angle * PAI / 180# '��ת��Ϊ����
        y2 = Cos(x2)            '����cos(x)
        picCos.Line (x1, y1)-(x2, y2)   '��ǰ��ɫ���߶�
        x1 = x2                     'ǰһ���յ��Ϊ��һ�����
        y1 = y2
    Next angle
    cmdDrawCos.Caption = "�����������"
Else
    picCos.Cls
    cmdDrawCos.Caption = "����������"
End If
End Sub

'����������
Private Sub cmdDrawSin_Click()
Dim angle As Integer
Dim x As Single, y As Single
If cmdDrawSin.Caption = "����������" Then
    Call draw_axis(picSin)  '���ò���ͼƬ��picSin��������
    For angle = 0 To 360 Step 1
        x = angle * PAI / 180#      '��0�ȵ�360�Ȼ���������
        y = Sin(x)              '����Sin(x)
        picSin.PSet (x, y)      '��ǰ��ɫ��һ��
    Next angle
    cmdDrawSin.Caption = "�����������"
Else
    picSin.Cls
    cmdDrawSin.Caption = "����������"
End If
End Sub

'��������ϵ���������ἰ��̶�
Public Sub draw_axis(obj As Object)
Dim angle As Integer
Dim x As Single, y As Single, t As Single
obj.Scale (-1#, 2#)-(8#, -2#)   '��������ϵͳ
obj.DrawWidth = 2               '�����ߵĿ��
obj.Line (-2#, 0#)-(7.5, 0#) '��X�ᣬy����ֵΪ0
For angle = -90 To 360 Step 90      '��X���ϱ�ǿ̶ȣ�����90
    x = angle * PAI / 180#          '��ת��Ϊ����
    obj.Line (x, 0#)-(x, 0.1)       '���̶���
    obj.CurrentX = x - 0.4          'ȷ����ʾ�̶�ֵλ��
    obj.CurrentY = -0.1             '
    obj.Print angle                 '��ʾ�̶�ֵ
Next angle
obj.Line (7.2, 0.1)-(7.5, 0#)       '��X�᷽���ͷ
obj.Line (7.2, -0.1)-(7.5, 0#)      '
obj.CurrentX = 7.2                  '��ʾX���־λ��
obj.CurrentY = 0.5
obj.Print " X "                     '��ʾ�ַ�X
obj.Line (0#, -1.9)-(0#, 1.9)       '��ʾY��
For t = -1.5 To 1.5 Step 0.5        'Y���ϱ�ǿ̶ȣ�����0.5
    If (Abs(t) > 0.1) Then          '�̶�0������ʾ�̶�ֵ
        obj.Line (0#, t)-(0.1, t)
        obj.CurrentX = 0.2
        obj.CurrentY = t + 0.1
        obj.Print t
    End If
Next t
obj.Line (-0.1, 1.7)-(0#, 1.9)      '��Y�᷽���ͷ
obj.Line (0.1, 1.7)-(0#, 1.9)
obj.CurrentX = 0.3
obj.CurrentY = 1.9
obj.Print " Y "
End Sub
