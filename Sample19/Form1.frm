VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6072
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   14796
   LinkTopic       =   "Form1"
   ScaleHeight     =   6072
   ScaleWidth      =   14796
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Sub Form_Click()
'���ַ�
Dim bExit As Boolean
bExit = False
Dim top%, bottom%, middle%
Dim i%, x%
Dim search(1 To 20) As Integer
For i = 1 To 20
    search(i) = (i - 1) * 3 + 1
    Print search(i);
Next i

x = 43
top = LBound(search)
bottom = UBound(search)
middle = (top + bottom) \ 2
Do While top <= bottom
    middle = (top + bottom) \ 2
    If x = search(middle) Then
        bExit = True
        Exit Do
    ElseIf x > search(middle) Then
        top = middle + 1
    Else
        bottom = middle - 1
    End If
Loop

Print
If bExit Then
    Print search(middle)
Else
    Print "û���ҵ�"
End If


'��������
'Dim a(0 To 10) As Integer
'Dim i%, j%, iMax%, tmp%
'For i = 0 To 10
'    a(i) = Int(Rnd * 101)
'    Print a(i);
'Next i
''�������� ѡ��
''For i = LBound(a) To UBound(a) - 1
''    iMax = i
''    For j = i + 1 To UBound(a)
''        If a(j) > a(iMax) Then iMax = j
''    Next j
''    tmp = a(i)
''    a(i) = a(iMax)
''    a(iMax) = tmp
''Next i
'
'
''�������� ð�ݷ�
'For i = UBound(a) - 1 To LBound(a) Step -1
'    For j = LBound(a) To i
'        If a(j) > a(j + 1) Then
'            tmp = a(j)
'            a(j) = a(j + 1)
'            a(j + 1) = tmp
'        End If
'    Next j
'Next i
'
'Print
'For i = 0 To 10
'    Print a(i);
'Next i



''��������
'Dim i%, t%
'Dim a(1 To 13) As Integer
'For i = 1 To 13
'    a(i) = i - 1
'Next
'
'For i = 1 To 13
'    Print a(i);
'Next i
'
'For i = 1 To 13 \ 2
'    t = a(i)
'    a(i) = a(13 - i + 1)
'    a(13 - i + 1) = t
'Next i
'Print
'For i = 1 To 13
'    Print a(i);
'Next i

''����ƽ���ֺ͸���ƽ���ֵ�����
'Dim aver!, sum!, i%, n%
'Dim mark(1 To 100) As Integer
'aver = 0
'sum = 0
'n = 0
'For i = 1 To 100
'    mark(i) = Int(Rnd * 101)
'    sum = sum + mark(i)
'Next i
'aver = sum / 100 'ƽ����
'For i = 1 To 100
'    If mark(i) > aver Then
'        n = n + 1
'    End If
'Next i
'Print n, aver;
'Print
'For i = 1 To 100
'    Print mark(i)
'Next i

''�Զ�����������
'Dim man As MANType
'With man
'    .No = 25000
'    .Name = "��ѩ÷"
'    .Sex = "Ů"
'    .Speciality = "�����黭"
'    .Birthdate = #8/13/1800#
'    Print .No; ""; .Name; ""; .Sex; ""; Format(.Birthdate, "yyyy��mm��dd��");
'End With


''����Ԫ�ص�ɾ��
'Dim a%(1 To 10), i%, k%
'For i = 1 To 10
'    a(i) = (i - 1) * 3 + 1
'Next i
'
'For i = 1 To 10
'    Print a(i);
'Next i
'
'For k = 1 To 10
'    If a(k) = 13 Then Exit For
'Next k
'
'For i = k To 9
'    a(i) = a(i + 1)
'Next i
'Print
'For i = 1 To 10
'    Print a(i);
'Next i


''��������
'Dim a%(1 To 10)
'Dim i%, k%, key%, key1 As Variant
'For i = 1 To 9
'    a(i) = (i - 1) * 3 + 1
'Next i
'For i = 1 To 9
'    Print a(i); Spc(1);
'Next i
'
'For k = 1 To 9
'    If 14 < a(k) Then Exit For
'Next
'
'For i = 9 To k Step -1
'    a(i + 1) = a(i)
'Next
'a(k) = 14
'Print
'For i = 1 To 10
'    Print a(i); Spc(1);
'Next



''����������ֵ
'Dim Max%, i%, iMax%, iA%(1 To 10)
'For i = LBound(iA) To UBound(iA)
'    iA(i) = Int(100 * Rnd) + 1
'    Print "iA(" & CStr(i) & ")=" & CStr(iA(i));
'    Print
'Next
'
'
'Max = iA(1): iMax = 1
'For i = 2 To 10
'    If iA(i) > Max Then
'        Max = iA(i)
'        iMax = i
'    End If
'Next i
'
'Print "���ֵ" & CStr(Max) + ",�±�" & CStr(iMax)
End Sub

