Attribute VB_Name = "modStudent"
Option Explicit
'最长显示6个汉字或12个英文字符
Public Const NAME_STR_LEN As Integer = 12
'自定义数据类型
Public Type Student
    name As String
    sex As String * 1
    score As Single
End Type
'定义全局动态数组
Public stuInf() As Student

Public Sub Main()
    '初始化全局数组,下届和上届都为-1
    ReDim stuInf(-1 To -1)
    frmStudent.Show
End Sub

Public Sub save_to_stuInf(name As String, sex As String, score As Single)
    Dim k As Integer '定义过程级变量
    k = UBound(stuInf) '求全局数组的上界
    If (k = -1) Then        '初始化时上界为-1
        k = 0
        ReDim stuInf(0 To k)    '重新定义stuinf的大小,下届为0
    Else
        k = k + 1
        ReDim Preserve stuInf(0 To k)   '重新定义数组大小,保留原有数据
    End If
    stuInf(k).name = name
    stuInf(k).score = score
    stuInf(k).sex = sex
End Sub

Public Sub delete_from_stuInf(ByVal index As Integer)
    Dim i As Integer, k As Integer
    k = UBound(stuInf)
    If (k = 0) Then                 '元素中只有一个元素
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

'姓名从小到大排序
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

'分数从小到大排序
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

'对楷体GB2312来说,两个英文字符正好和一个汉字的显示宽度相同
'通过给字符串nameSStr右边插入若干空格使其变为制定长度,便于字符串显示
Public Function name_d_str(ByVal nameSStr As String) As String
    Dim i As Integer, mlen As Integer
    Dim k As Integer
    Dim mnum1 As Integer    '英文字符个数
    Dim mnum2 As Integer    '汉字字符个数
    Dim mnum  As Integer '要显示的字符节字个数
    mnum = 0
    mnum1 = 0
    mnum2 = 0
    mlen = Len(nameSStr)        '求源字符串的长度
    For i = 1 To mlen
        k = Asc(Mid(nameSStr, i, 1)) '计算指定字符的ASCII码
        If (k > 0 And k < 255) Then     '英文字符
            mnum = mnum1 + 1
            mnum = mnum + 1     '英文每个字符加1
        Else
            mnum2 = mnum2 + 1
            mnum = mnum + 2     '汉字每个字符加2
        End If
        If (mnum >= NAME_STR_LEN) Then Exit For  '超过能显示的字符个数时推出
    Next i
    If (mnum < NAME_STR_LEN) Then
        nameSStr = nameSStr + Space(NAME_STR_LEN)   '如串"aabb 王 c"
        mlen = NAME_STR_LEN - mnum2 '计算取出字符个数
        name_d_str = Left(nameSStr, mlen)
    ElseIf (mnum - NAME_STR_LEN) Then           '
        mlen = i
        name_d_str = Left(nameSStr, mlen)
    Else
        mlen = i - 1
        name_d_str = Left(nameSStr, mlen) + Space(i)
    End If
End Function

