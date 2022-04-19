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
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
'完数
'是一个整数，其所有因子，除去本身外，因子相加之和等于本身
'例如：6，因子1,2,3,6
Dim i%, j%, s%
For j = 1 To 100
    s = 0
    For i = 1 To j - 1
        If j Mod i = 0 Then '求因子
            s = s + i
        End If
    Next i
    If s = j Then
        Print j & "输入的数是完数"
    End If
Next j

'水仙花
'是一个3位数，每位数的立法之和等于本身
'例如 153=1*1*1+5*5*5+3*3*3
''算法1
'Dim i%, a%, b%, c%
'For i = 100 To 999
'    a = i Mod 10 '个位
'    b = (i \ 10) Mod 10 '十位
'    c = i \ 100 '百位
'    If i = a * a * a + b * b * b + c * c * c Then
'        Print i
'    End If
'Next i
''算法2
'Dim i%, a%, b%, c%
'For i = 100 To 999
'    a = i Mod 10 '个位
'    b = (i Mod 100) \ 10 '十位
'    c = i \ 100 '百位
'    If i = a * a * a + b * b * b + c * c * c Then
'        Print i
'    End If
'Next i
''算法3
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

''素数
''大于2且不能被1和本身以外的整数整除的整数
'Dim flag As Boolean
'Dim number%, i%
'number = Val(InputBox("请输入一个整数"))
'flag = True
'For i = 2 To number - 1 Step 1
'    If number Mod i = 0 Then
'        flag = False
'        Exit For
'    End If
'Next i
'If flag Then
'    Print number & "是素数"
'Else
'    Print number & "不是素数,能被整数" & i & "整除"
'End If


'求最大值
'随机产生10个100~200的数
'Dim i%, x%, max%
'max = 100
'For i = 1 To 10
'    x = Int(Rnd * 101 + 100)
'    Print x;
'    If x > max Then max = x
'Next i
'
'Print
'Print "最大值="; max
''求自然对数e的近似值
'Dim i%, n&, t!, e!
'e = 0: n = 1
'i = 0: t = 1
'Do While t > 0.00001
'    e = e + t: i = i + 1
'    n = n * i: t = 1 / n
'Loop
'
'Print "计算了"; i; "项的和是 "; e

''递归
''斐波那契
'Dim i%
'Dim f1&, f2&, f&
'f1 = 1: f2 = 1
'For i = 3 To 20 Step 1
'    f = f2 + f1
'    Print f&; Space(2)
'    f1 = f2
'    f2 = f
'Next i


'1元人民币换成1分、2分、5分供50枚的换法
'i j k 分别表示1分、2分、5分
'i+j+k=50
'i+2j+5k=100

''算法1
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
''算法2
'Dim i%, j%, k%
'For i = 0 To 50
'    For j = 0 To 50
'        k = 50 - i - j
'        If i + 2 * j + 5 * k = 100 Then
'                Print i, j, k
'        End If
'    Next j
'Next i
''算法3
'Dim i%, j%, k%
'For k = 0 To 20
'    For j = 0 To 50
'        i = 50 - k - j
'        If 5 * k + 2 * j + i = 100 Then
'            Print i, j, k
'        End If
'    Next j
'Next k
'算法4
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


'输出圣诞树
''算法1
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
''算法2
'Dim i%, j%, max%
'max = 10
'For i = 1 To max
'    Print String(max - i + 1, " ");
'    Print String(2 * i - 1, "*")
'Next i

'计算 阶层加 1！+2！+3！.
'算法1
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

'算法2
'Dim x%, y&, sum&
'y = 1: sum = 0
'For x = 1 To 3
'    y = y * x
'    sum = sum + y
'Next x
'Print sum

''九九乘法表
'Dim i%, j%, s$
'For i = 1 To 9
'    For j = 1 To i
'       s = i & "x" & j & "=" & i + j
'       Print Tab((j - 1) * 9 + 1); s;
'    Next j
'    Print
'Next i


''输出N行*
'Dim m%, i%
'm = Val(InputBox("请输入N行 "))
'i = 1
'Do While i <= m
'
'    Print "**********"
'    i = i + 1
'Loop


'
''逆序输出
'Dim number&, i&
'number = Val(InputBox("请输入正整数"))
'Print "输入数为:" & number
'Do
'    i = number Mod 10
'    Print i;
'    number = number \ 10
'10
'Loop While number <> 0

''变量交换
'Dim x, y, t As Integer
'x = Val(InputBox("请输入一个X值"))
'y = Val(InputBox("请输入一个Y值"))
'Print "交换前:", x, y
'
'If x > y Then
'    t = x
'    x = y
'    y = t
'End If
'Print "交换后:", x, y

''三个数排序
'Dim x, y, z, t As Integer
'x = Val(InputBox("请输入一个X值"))
'y = Val(InputBox("请输入一个Y值"))
'z = Val(InputBox("请输入一个Z值"))
'Print "交换前:", x, y, z
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
'Print "交换后:", x, y, z

'算三角形面积
'Dim x, y, z As Integer
'Dim s, area As Single
'x = Val(InputBox("请输入X值"))
'y = Val(InputBox("请输入Y值"))
'z = Val(InputBox("请输入Z值"))
'If x < y + z And y < x + z And z < x + y Then
'    s = (x + y + z) / 2
'    area = Sqr(s * (s - x) * (s - y) * (s - z))
'    Print area
'Else
'    Print "输入的三边不符合三角形定义的要求"
'End If

'判断输入的数的位数
'Dim x As Integer
'x = Val(InputBox("请输入一个数"))
'If x < 0 Then
'    Print "输入的是负数"
'ElseIf x < 10 Then
'    Print "输入的是一位数"
'ElseIf x < 100 Then
'    Print "输入的是两位数"
'Else
'    Print "输入的是两位以上的数"
'End If

'窗体颜色
'Dim strColor As String
'strColor = InputBox("请输入颜色的名称(red,blue,green)")
'strColor = LCase(strColor)
'Select Case strColor
'    Case "red"
'        Form1.BackColor = RGB(255, 0, 0)
'    Case "blue"
'        Form1.BackColor = RGB(0, 0, 255)
'    Case "green"
'        Form1.BackColor = RGB(0, 255, 0)
'    Case Else
'        Print "无法识别的颜色名称"
'End Select

''成绩判断方式1
'Dim mark As Single
'mark = Val(InputBox("请输入一个百分数"))
'If mark >= 90 And mark <= 100 Then
'    Print "优秀"
'End If
'
'If mark >= 80 And mark < 90 Then
'    Print "良好"
'End If
'
'If mark >= 70 And mark < 80 Then
'    Print "中等"
'End If
'
'If mark >= 60 And mark < 70 Then
'    Print "及格"
'End If
'
'If mark >= 0 And mark < 60 Then
'    Print "不及格"
'End If

''成绩判断方式2
'Dim mark As Single
'mark = Val(InputBox("请输入一个百分数"))
'If mark >= 90 Then
'    Print "优秀"
'ElseIf mark >= 80 Then
'    Print "良好"
'ElseIf mark >= 70 Then
'    Print "中等"
'ElseIf mark >= 60 Then
'    Print "及格"
'Else
'    Print "不及格"
'End If

''成绩判断方式3
'Dim mark As Single
'mark = Val(InputBox("请输入一个百分数"))
'Select Case mark
'    Case 90 To 100
'        Print "优秀"
'    Case 80 To 90
'         Print "良好"
'    Case 70 To 80
'         Print "中等"
'    Case 60 To 70
'         Print "及格"
'    Case Else
'           Print "不及格"
'End Select

''成绩判断方式4
'Dim mark As Single
'Dim grade As Integer
'mark = Val(InputBox("请输入一个百分数"))
'grade = mark \ 10
'Select Case grade
'    Case 9, 10
'        Print "优秀"
'    Case 8
'         Print "良好"
'    Case 7
'         Print "中等"
'    Case 6
'         Print "及格"
'    Case Else
'           Print "不及格"
'End Select

''分支
'Dim x, y As Single
'x = Val(InputBox("请输入数值"))
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
'算法1
'Dim i%, sum%
'sum = 0
'For i = 1 To 100 Step 1
'    sum = sum + i
'Next i
'Print sum
''算法2
'Dim n%, sum%
'n = 1: sum = 0
'Do While n <= 100
'    sum = sum + n
'    n = n + 1
'Loop
'Print sum
''算法3
'Dim n%, sum%
'n = 1: sum = 0
'Do Until n > 100
'    sum = sum + n
'    n = n + 1
'Loop
'Print sum
'算法4
'Dim n%, sum%
'n = 1: sum = 0
'Do
'    sum = sum + n
'    n = n + 1
'Loop While n <= 100
'Print sum
''算法5
'Dim n%, sum%
'n = 1: sum = 0
'Do
'    sum = sum + n
'    n = n + 1
'Loop Until n > 100
'Print sum


'1+3+5...+99
'算法1
'Dim i%, sum%
'sum = 0
'For i = 1 To 100 Step 2
'    sum = sum + i
'Next i
'Print sum
'算法2
'Dim i%, sum%
'sum = 0
'For i = 1 To 100 Step 1
'    If i Mod 2 <> 0 Then
'        sum = sum + i
'    End If
'Next i
'Print sum

'求阶层
'Dim i%, sum%
'sum = 1
'For i = 1 To 5
'    sum = sum * i
'Next
'Print sum
End Sub

