Attribute VB_Name = "mdlPublic"
Option Explicit

'�Զ�����������
Public Type StuType
    No As Integer       'ѧ��
    Name As String * 20 '����
    Sex As String * 1   '�Ա�
    Mark(1 To 4) As Single '4�ſγ̳ɼ�
    Total As Single  '�ܷ�
End Type

Public Type MANType
    No As Integer 'ѧ��
    Name As String        '����
    Sex As String * 1   '�Ա�
    Birthdate As Date   '��������
    Speciality As String    '�س�
End Type

Public Type StudType
    Name As String * 10     '����
    Special As String * 10  'רҵ
    Total   As Single       '�ܷ�
    
End Type
