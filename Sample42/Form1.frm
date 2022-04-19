VERSION 5.00
Begin VB.Form frmMoving 
   Caption         =   "简单动画演示"
   ClientHeight    =   3780
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   6144
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   6144
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer tmrMov 
      Enabled         =   0   'False
      Left            =   2400
      Top             =   600
   End
   Begin VB.CommandButton cmdSS 
      Caption         =   "开始"
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
    If cmdSS.Caption = "开始" Then
        tmrMov.Enabled = True
        tmrMov.Interval = 20
        cmdSS.Caption = "停止"
    Else
        tmrMov.Enabled = False
        cmdSS.Caption = "开始"
    End If
End Sub

Private Sub Form_Activate()
    Dim myes As Boolean
    myes = set_coorsSys(frmMoving)   '设置窗体frmMoving的坐标系统
    If Not myes Then Exit Sub       '设置失败则退出
    Call set_drawEnv(frmMoving)     '设置画图属性
    Call tmrMov_Timer               '画运动的圆
End Sub


'设置画图属性过程
Private Sub set_drawEnv(obj As Object)
m_r = IIf(m_a >= m_b, m_b, m_a) / 10# '计算画圆的半径
obj.DrawMode = 13               '设置画图模式为13(CopyPen)
obj.FillStyle = 0               '为画大圆设置填充模式
obj.FillColor = vbRed           '填充颜色
obj.Circle (0, 0), m_r + m_r, vbRed '画圆
obj.FillStyle = 1               '透明
If m_a >= m_b Then
    obj.Circle (0, 0), m_a, vbBlue, , , m_b / m_a '画椭圆轨道
Else
    obj.Circle (0, 0), m_b, vbBlue, , , m_b / m_a '画椭圆轨道
End If
obj.DrawMode = 7            '设置图画模式为7
obj.FillStyle = 0           '为小圆设置填充模式
obj.FillColor = vbRed       '为小圆设置填充颜色
End Sub




'设置坐标系统
Public Function set_coorsSys(obj As Object) As Boolean
    Dim xMin As Double, xMax As Double
    Dim yMin As Double, yMax As Double
    Dim rate    As Double
    If obj.ScaleHeight <= 0# Or obj.ScaleWidth <= 0# Then
        MsgBox "对象(如窗体)宽或者高设置不合适!"
        set_coorsSys = False
        Exit Function
    End If
    
    rate = obj.ScaleHeight / obj.ScaleWidth     '计算对象高宽之比
    xMin = -1000#                                   '设定x轴方向2000
    xMax = 1000#
    m_a = 800#          '椭圆X轴半径800
    yMax = 1000# * rate         '按比例计算y方向长度
    yMin = -yMax
    m_b = m_a * rate    '椭圆y轴方向半径的长度
    obj.Cls
    obj.Scale (xMin, yMax)-(xMax, yMin)         '定义坐标系
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
