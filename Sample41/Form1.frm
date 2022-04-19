VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "正弦曲线和余弦曲线"
   ClientHeight    =   3264
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   3264
   ScaleWidth      =   9060
   StartUpPosition =   3  '窗口缺省
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
      Caption         =   "画余弦曲线"
      Height          =   372
      Left            =   5520
      TabIndex        =   2
      Top             =   2760
      Width           =   2052
   End
   Begin VB.CommandButton cmdDrawSin 
      Caption         =   "画正弦曲线"
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


'画余弦曲线
Private Sub cmdDrawCos_Click()
Dim angle As Integer
Dim x1 As Single, y1 As Single
Dim x2 As Single, y2 As Single
If cmdDrawCos.Caption = "画余弦曲线" Then
    Call draw_axis(picCos)      '设置并画图片框picCos的坐标轴
    x1 = 0# * PAI / 180#        '度转换为弧度
    y1 = Cos(x1)                '计算cos(x)
    For angle = 0 To 360 Step 1 '从0度到360度画余弦曲线
        x2 = angle * PAI / 180# '度转换为弧度
        y2 = Cos(x2)            '计算cos(x)
        picCos.Line (x1, y1)-(x2, y2)   '以前景色画线段
        x1 = x2                     '前一段终点变为后一段起点
        y1 = y2
    Next angle
    cmdDrawCos.Caption = "清除余弦曲线"
Else
    picCos.Cls
    cmdDrawCos.Caption = "画余弦曲线"
End If
End Sub

'画正弦曲线
Private Sub cmdDrawSin_Click()
Dim angle As Integer
Dim x As Single, y As Single
If cmdDrawSin.Caption = "画正弦曲线" Then
    Call draw_axis(picSin)  '设置并画图片框picSin的坐标轴
    For angle = 0 To 360 Step 1
        x = angle * PAI / 180#      '从0度到360度画正弦曲线
        y = Sin(x)              '计算Sin(x)
        picSin.PSet (x, y)      '以前景色画一点
    Next angle
    cmdDrawSin.Caption = "清除正铉曲线"
Else
    picSin.Cls
    cmdDrawSin.Caption = "画正弦曲线"
End If
End Sub

'设置坐标系并画坐标轴及其刻度
Public Sub draw_axis(obj As Object)
Dim angle As Integer
Dim x As Single, y As Single, t As Single
obj.Scale (-1#, 2#)-(8#, -2#)   '定义坐标系统
obj.DrawWidth = 2               '设置线的宽度
obj.Line (-2#, 0#)-(7.5, 0#) '画X轴，y坐标值为0
For angle = -90 To 360 Step 90      '在X轴上标记刻度，步长90
    x = angle * PAI / 180#          '度转换为弧度
    obj.Line (x, 0#)-(x, 0.1)       '画刻度线
    obj.CurrentX = x - 0.4          '确定显示刻度值位置
    obj.CurrentY = -0.1             '
    obj.Print angle                 '显示刻度值
Next angle
obj.Line (7.2, 0.1)-(7.5, 0#)       '画X轴方向箭头
obj.Line (7.2, -0.1)-(7.5, 0#)      '
obj.CurrentX = 7.2                  '显示X轴标志位置
obj.CurrentY = 0.5
obj.Print " X "                     '显示字符X
obj.Line (0#, -1.9)-(0#, 1.9)       '显示Y轴
For t = -1.5 To 1.5 Step 0.5        'Y轴上标记刻度，步长0.5
    If (Abs(t) > 0.1) Then          '刻度0处不显示刻度值
        obj.Line (0#, t)-(0.1, t)
        obj.CurrentX = 0.2
        obj.CurrentY = t + 0.1
        obj.Print t
    End If
Next t
obj.Line (-0.1, 1.7)-(0#, 1.9)      '画Y轴方向箭头
obj.Line (0.1, 1.7)-(0#, 1.9)
obj.CurrentX = 0.3
obj.CurrentY = 1.9
obj.Print " Y "
End Sub
