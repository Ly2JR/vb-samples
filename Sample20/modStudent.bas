Attribute VB_Name = "modStudent"
Option Explicit
'���ʾ6�����ֻ�12��Ӣ���ַ�
Public Const NAME_STR_LEN As Integer = 12
'�Զ�����������
Public Type Student
    name As String
    sex As String * 1
    score As Single
End Type
'����ȫ�ֶ�̬����
Public stuInf() As Student

Public Sub Main()
    '��ʼ��ȫ������,�½���Ͻ춼Ϊ-1
    ReDim stuInf(-1 To -1)
    frmStudent.Show
End Sub

Public Sub save_to_stuInf(name As String, sex As String, score As Single)
    Dim k As Integer '������̼�����
    k = UBound(stuInf) '��ȫ��������Ͻ�
    If (k = -1) Then        '��ʼ��ʱ�Ͻ�Ϊ-1
        k = 0
        ReDim stuInf(0 To k)    '���¶���stuinf�Ĵ�С,�½�Ϊ0
    Else
        k = k + 1
        ReDim Preserve stuInf(0 To k)   '���¶��������С,����ԭ������
    End If
    stuInf(k).name = name
    stuInf(k).score = score
    stuInf(k).sex = sex
End Sub

Public Sub delete_from_stuInf(ByVal index As Integer)
    Dim i As Integer, k As Integer
    k = UBound(stuInf)
    If (k = 0) Then                 'Ԫ����ֻ��һ��Ԫ��
        ReDim stuInf(-1 To -1)
    Else
        k = k - 1
        For i = index To k
            stuInf(i) = stuInf(i + 1)
        Next i
        ReDim Preserve stuInf(0 To k)
    End If
End Sub

Public Sub copy_stuinf_to_d(d() As Student)
    Dim i As Integer, k As Integer
    k = UBound(stuInf)
    ReDim d(0 To k)
    For i = 0 To k
        d(i) = stuInf(i)
    Next i
End Sub

'������С��������
Public Sub sort_d_by_name(d() As Student)
    Dim i As Integer, j As Integer
    Dim k As Integer
    Dim t As Student
    k = UBound(stuInf)
    For i = 0 To k - 1
        For j = 0 To k - i - 1
            If (d(j).name > d(j + 1).name) Then
                t = d(j)
                d(j) = d(j + 1)
                d(j + 1) = t
            End If
        Next j
    Next i
End Sub

'������С��������
Public Sub sort_d_by_score(d() As Student)
    Dim i As Integer, j As Integer
    Dim k As Integer
    Dim t As Student
    k = UBound(stuInf)
    For i = 0 To k - 1
        For j = 0 To k - i - 1
            If (d(j).score > d(j + 1).score) Then
                t = d(j)
                d(j) = d(j + 1)
                d(j + 1) = t
            End If
        Next j
    Next i
End Sub

'�Կ���GB2312��˵,����Ӣ���ַ����ú�һ�����ֵ���ʾ�����ͬ
'ͨ�����ַ���nameSStr�ұ߲������ɿո�ʹ���Ϊ�ƶ�����,�����ַ�����ʾ
Public Function name_d_str(ByVal nameSStr As String) As String
    Dim i As Integer, mlen As Integer
    Dim k As Integer
    Dim mnum1 As Integer    'Ӣ���ַ�����
    Dim mnum2 As Integer    '�����ַ�����
    Dim mnum  As Integer 'Ҫ��ʾ���ַ����ָ���
    mnum = 0
    mnum1 = 0
    mnum2 = 0
    mlen = Len(nameSStr)        '��Դ�ַ����ĳ���
    For i = 1 To mlen
        k = Asc(Mid(nameSStr, i, 1)) '����ָ���ַ���ASCII��
        If (k > 0 And k < 255) Then     'Ӣ���ַ�
            mnum = mnum1 + 1
            mnum = mnum + 1     'Ӣ��ÿ���ַ���1
        Else
            mnum2 = mnum2 + 1
            mnum = mnum + 2     '����ÿ���ַ���2
        End If
        If (mnum >= NAME_STR_LEN) Then Exit For  '��������ʾ���ַ�����ʱ�Ƴ�
    Next i
    If (mnum < NAME_STR_LEN) Then
        nameSStr = nameSStr + Space(NAME_STR_LEN)   '�紮"aabb �� c"
        mlen = NAME_STR_LEN - mnum2 '����ȡ���ַ�����
        name_d_str = Left(nameSStr, mlen)
    ElseIf (mnum - NAME_STR_LEN) Then           '
        mlen = i
        name_d_str = Left(nameSStr, mlen)
    Else
        mlen = i - 1
        name_d_str = Left(nameSStr, mlen) + Space(i)
    End If
End Function

