VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5115
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   9360
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
'����
'��һ�����������������ӣ���ȥ�����⣬�������֮�͵��ڱ���
'���磺6������1,2,3,6
Dim i%, j%, s%
For j = 1 To 100
    s = 0
    For i = 1 To j - 1
        If j Mod i = 0 Then '������
            s = s + i
        End If
    Next i
    If s = j Then
        Print j & "�������������"
    End If
Next j

'ˮ�ɻ�
'��һ��3λ����ÿλ��������֮�͵��ڱ���
'���� 153=1*1*1+5*5*5+3*3*3
''�㷨1
'Dim i%, a%, b%, c%
'For i = 100 To 999
'    a = i Mod 10 '��λ
'    b = (i \ 10) Mod 10 'ʮλ
'    c = i \ 100 '��λ
'    If i = a * a * a + b * b * b + c * c * c Then
'        Print i
'    End If
'Next i
''�㷨2
'Dim i%, a%, b%, c%
'For i = 100 To 999
'    a = i Mod 10 '��λ
'    b = (i Mod 100) \ 10 'ʮλ
'    c = i \ 100 '��λ
'    If i = a * a * a + b * b * b + c * c * c Then
'        Print i
'    End If
'Next i
''�㷨3
'Dim i%, j%, k%, a$
'For i = 1 To 9
'    For j = 0 To 9
'        For k = 0 To 9
'            If i * 100 + j * 10 + k = i ^ 3 + j ^ 3 + k ^ 3 Then
'                a = a & i & j & k & Space(2)
'
'            End If
'        Next k
'    Next j
'Next i
'Print a;

''����
''����2�Ҳ��ܱ�1�ͱ����������������������
'Dim flag As Boolean
'Dim number%, i%
'number = Val(InputBox("������һ������"))
'flag = True
'For i = 2 To number - 1 Step 1
'    If number Mod i = 0 Then
'        flag = False
'        Exit For
'    End If
'Next i
'If flag Then
'    Print number & "������"
'Else
'    Print number & "��������,�ܱ�����" & i & "����"
'End If


'�����ֵ
'�������10��100~200����
'Dim i%, x%, max%
'max = 100
'For i = 1 To 10
'    x = Int(Rnd * 101 + 100)
'    Print x;
'    If x > max Then max = x
'Next i
'
'Print
'Print "���ֵ="; max
''����Ȼ����e�Ľ���ֵ
'Dim i%, n&, t!, e!
'e = 0: n = 1
'i = 0: t = 1
'Do While t > 0.00001
'    e = e + t: i = i + 1
'    n = n * i: t = 1 / n
'Loop
'
'Print "������"; i; "��ĺ��� "; e

''�ݹ�
''쳲�����
'Dim i%
'Dim f1&, f2&, f&
'f1 = 1: f2 = 1
'For i = 3 To 20 Step 1
'    f = f2 + f1
'    Print f&; Space(2)
'    f1 = f2
'    f2 = f
'Next i


'1Ԫ����һ���1�֡�2�֡�5�ֹ�50ö�Ļ���
'i j k �ֱ��ʾ1�֡�2�֡�5��
'i+j+k=50
'i+2j+5k=100

''�㷨1
'Dim i%, j%, k%
'For i = 0 To 50
'    For j = 0 To 50
'        For k = 0 To 50
'            If i + j + k = 50 And i + 2 * j + 5 * k = 100 Then
'                Print i, j, k
'            End If
'        Next k
'    Next j
'Next i
''�㷨2
'Dim i%, j%, k%
'For i = 0 To 50
'    For j = 0 To 50
'        k = 50 - i - j
'        If i + 2 * j + 5 * k = 100 Then
'                Print i, j, k
'        End If
'    Next j
'Next i
''�㷨3
'Dim i%, j%, k%
'For k = 0 To 20
'    For j = 0 To 50
'        i = 50 - k - j
'        If 5 * k + 2 * j + i = 100 Then
'            Print i, j, k
'        End If
'    Next j
'Next k
'�㷨4
'i=50-k-j
'50-k-j+2j+5k=100
'j+4k=50
'j=50-4k
'Dim i%, j%, k%
'For k = 0 To 12
'    j = 50 - 4 * k
'    i = 50 - j - -k
'    Print i, j, k
'Next k


'���ʥ����
''�㷨1
'Dim i%, j%
'For i = 1 To 4
'    For j = 1 To 4 - i + 1
'        Print " ";
'    Next
'    For j = 1 To 2 * i - 1
'        Print "*";
'    Next
'    Print
'Next i
''�㷨2
'Dim i%, j%, max%
'max = 10
'For i = 1 To max
'    Print String(max - i + 1, " ");
'    Print String(2 * i - 1, "*")
'Next i

'���� �ײ�� 1��+2��+3��.
'�㷨1
'Dim i%, j%
'Dim sum&, s&
'sum = 0
'For i = 1 To 10 Step 1
'    s = 1
'    For j = 1 To i Step 1
'        s = j * s
'    Next j
'    sum = sum + s
'Next i
'Print sum

'�㷨2
'Dim x%, y&, sum&
'y = 1: sum = 0
'For x = 1 To 3
'    y = y * x
'    sum = sum + y
'Next x
'Print sum

''�žų˷���
'Dim i%, j%, s$
'For i = 1 To 9
'    For j = 1 To i
'       s = i & "x" & j & "=" & i + j
'       Print Tab((j - 1) * 9 + 1); s;
'    Next j
'    Print
'Next i


''���N��*
'Dim m%, i%
'm = Val(InputBox("������N�� "))
'i = 1
'Do While i <= m
'
'    Print "**********"
'    i = i + 1
'Loop


'
''�������
'Dim number&, i&
'number = Val(InputBox("������������"))
'Print "������Ϊ:" & number
'Do
'    i = number Mod 10
'    Print i;
'    number = number \ 10
'10
'Loop While number <> 0

''��������
'Dim x, y, t As Integer
'x = Val(InputBox("������һ��Xֵ"))
'y = Val(InputBox("������һ��Yֵ"))
'Print "����ǰ:", x, y
'
'If x > y Then
'    t = x
'    x = y
'    y = t
'End If
'Print "������:", x, y

''����������
'Dim x, y, z, t As Integer
'x = Val(InputBox("������һ��Xֵ"))
'y = Val(InputBox("������һ��Yֵ"))
'z = Val(InputBox("������һ��Zֵ"))
'Print "����ǰ:", x, y, z
'
'If x > y Then
'    t = x
'    x = y
'    y = t
'End If
'If x > z Then
'    t = x
'    x = z
'    z = t
'End If
'If y > z Then
'    t = y
'    y = z
'    z = t
'End If
'Print "������:", x, y, z

'�����������
'Dim x, y, z As Integer
'Dim s, area As Single
'x = Val(InputBox("������Xֵ"))
'y = Val(InputBox("������Yֵ"))
'z = Val(InputBox("������Zֵ"))
'If x < y + z And y < x + z And z < x + y Then
'    s = (x + y + z) / 2
'    area = Sqr(s * (s - x) * (s - y) * (s - z))
'    Print area
'Else
'    Print "��������߲����������ζ����Ҫ��"
'End If

'�ж����������λ��
'Dim x As Integer
'x = Val(InputBox("������һ����"))
'If x < 0 Then
'    Print "������Ǹ���"
'ElseIf x < 10 Then
'    Print "�������һλ��"
'ElseIf x < 100 Then
'    Print "���������λ��"
'Else
'    Print "���������λ���ϵ���"
'End If

'������ɫ
'Dim strColor As String
'strColor = InputBox("��������ɫ������(red,blue,green)")
'strColor = LCase(strColor)
'Select Case strColor
'    Case "red"
'        Form1.BackColor = RGB(255, 0, 0)
'    Case "blue"
'        Form1.BackColor = RGB(0, 0, 255)
'    Case "green"
'        Form1.BackColor = RGB(0, 255, 0)
'    Case Else
'        Print "�޷�ʶ�����ɫ����"
'End Select

''�ɼ��жϷ�ʽ1
'Dim mark As Single
'mark = Val(InputBox("������һ���ٷ���"))
'If mark >= 90 And mark <= 100 Then
'    Print "����"
'End If
'
'If mark >= 80 And mark < 90 Then
'    Print "����"
'End If
'
'If mark >= 70 And mark < 80 Then
'    Print "�е�"
'End If
'
'If mark >= 60 And mark < 70 Then
'    Print "����"
'End If
'
'If mark >= 0 And mark < 60 Then
'    Print "������"
'End If

''�ɼ��жϷ�ʽ2
'Dim mark As Single
'mark = Val(InputBox("������һ���ٷ���"))
'If mark >= 90 Then
'    Print "����"
'ElseIf mark >= 80 Then
'    Print "����"
'ElseIf mark >= 70 Then
'    Print "�е�"
'ElseIf mark >= 60 Then
'    Print "����"
'Else
'    Print "������"
'End If

''�ɼ��жϷ�ʽ3
'Dim mark As Single
'mark = Val(InputBox("������һ���ٷ���"))
'Select Case mark
'    Case 90 To 100
'        Print "����"
'    Case 80 To 90
'         Print "����"
'    Case 70 To 80
'         Print "�е�"
'    Case 60 To 70
'         Print "����"
'    Case Else
'           Print "������"
'End Select

''�ɼ��жϷ�ʽ4
'Dim mark As Single
'Dim grade As Integer
'mark = Val(InputBox("������һ���ٷ���"))
'grade = mark \ 10
'Select Case grade
'    Case 9, 10
'        Print "����"
'    Case 8
'         Print "����"
'    Case 7
'         Print "�е�"
'    Case 6
'         Print "����"
'    Case Else
'           Print "������"
'End Select

''��֧
'Dim x, y As Single
'x = Val(InputBox("��������ֵ"))
'Select Case x
'    Case Is > 3
'        y = x + 3
'    Case 1 To 3
'        y = x * x
'    Case Is > 0, Is < 1
'        y = x
'    Case Else
'        y = 0
'End Select
'Print y

''1+2+3...+100
'�㷨1
'Dim i%, sum%
'sum = 0
'For i = 1 To 100 Step 1
'    sum = sum + i
'Next i
'Print sum
''�㷨2
'Dim n%, sum%
'n = 1: sum = 0
'Do While n <= 100
'    sum = sum + n
'    n = n + 1
'Loop
'Print sum
''�㷨3
'Dim n%, sum%
'n = 1: sum = 0
'Do Until n > 100
'    sum = sum + n
'    n = n + 1
'Loop
'Print sum
'�㷨4
'Dim n%, sum%
'n = 1: sum = 0
'Do
'    sum = sum + n
'    n = n + 1
'Loop While n <= 100
'Print sum
''�㷨5
'Dim n%, sum%
'n = 1: sum = 0
'Do
'    sum = sum + n
'    n = n + 1
'Loop Until n > 100
'Print sum


'1+3+5...+99
'�㷨1
'Dim i%, sum%
'sum = 0
'For i = 1 To 100 Step 2
'    sum = sum + i
'Next i
'Print sum
'�㷨2
'Dim i%, sum%
'sum = 0
'For i = 1 To 100 Step 1
'    If i Mod 2 <> 0 Then
'        sum = sum + i
'    End If
'Next i
'Print sum

'��ײ�
'Dim i%, sum%
'sum = 1
'For i = 1 To 5
'    sum = sum * i
'Next
'Print sum
End Sub

