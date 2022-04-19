VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ΩªÕ®µ∆ƒ£ƒ‚"
   ClientHeight    =   1716
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   7488
   LinkTopic       =   "Form1"
   ScaleHeight     =   1716
   ScaleWidth      =   7488
   StartUpPosition =   3  '¥∞ø⁄»± °
   Begin VB.Timer tmrRGY 
      Left            =   3600
      Top             =   120
   End
   Begin VB.Frame Frame1 
      Caption         =   "∫Ï¬Ãµ∆"
      Height          =   1212
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6972
      Begin VB.Label txtRGY 
         BackColor       =   &H80000008&
         Caption         =   "Label1"
         ForeColor       =   &H8000000B&
         Height          =   732
         Left            =   3960
         TabIndex        =   1
         Top             =   360
         Width           =   2532
      End
      Begin VB.Shape shpYELLOW 
         Height          =   732
         Left            =   2280
         Shape           =   3  'Circle
         Top             =   360
         Width           =   972
      End
      Begin VB.Shape shpGreen 
         Height          =   732
         Left            =   1320
         Shape           =   3  'Circle
         Top             =   360
         Width           =   972
      End
      Begin VB.Shape shpRed 
         Height          =   732
         Left            =   360
         Shape           =   3  'Circle
         Top             =   360
         Width           =   972
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const RED_ON As Integer = 1 '∫Ïµ∆¡¡
Const GREEN_ON As Integer = 2 '¬Ãµ∆¡¡
Const YELLOW_ON As Integer = 3 'ª∆µ∆¡¡
Const RGY_ALL_OFF As Integer = 10 'À˘”–µ∆œ®√
Const RED_LONG As Integer = 10      '∫Ïµ∆—” ±60√Î
Const GREEN_LONG As Integer = 20  '¬Ãµ∆—” ±70√Î
Const YELLOW_LONG As Integer = 5    'ª∆µ∆—” ±5√Î
Dim m_RGY As Integer                '±Ì æµ±«∞µ∆¡¡
Dim m_RGY_Old As Integer            'º«¬ºª∆µ∆¡¡ ±«∞√Êµƒ¡¡µ∆
Dim m_long As Integer   'µ∆¡¡ ±≥§√Î


Private Sub Form_Load()
    Call set_RGYLight(RGY_ALL_OFF) 'œ®√À˘”–µ∆
    txtRGY.BackColor = vbBlack      '…Ë÷√±≥æ∞∫⁄…´
    m_RGY = RED_ON      '≥ı º∫Ïµ∆¡¡
    m_long = RED_LONG   '≥ı º∫Ïµ∆¡¡ ±≥§
    Call set_RGYLight(m_RGY)    '…Ë÷√∫Ïµ∆¡¡∆‰À˚µ∆œ®√
    tmrRGY.Enabled = True
    tmrRGY.Interval = 100
End Sub

Private Sub tmrRGY_Timer()
    Static mtime As String
    If m_RGY = YELLOW_ON Then
        Call deal_yellowlight
    End If
    If mtime = Time() Then Exit Sub
    mtime = Time()
    m_long = m_long - 1
    txtRGY.Caption = m_long
    If m_long < -0 Then
        Call change_RGYLight
        mtime = Time()
    End If
End Sub

Private Sub deal_yellowlight()
 Static ag As Boolean
 If ag Then
    Call set_RGYLight(RGY_ALL_OFF)
    txtRGY.Caption = ""
 Else
    Call set_RGYLight(m_RGY)
    txtRGY.Caption = m_long
 End If
 ag = Not ag
End Sub

Private Sub change_RGYLight()
    If m_RGY <> YELLOW_ON Then
        m_RGY_Old = m_RGY
        m_RGY = YELLOW_ON
        m_long = YELLOW_LONG
    ElseIf m_RGY_Old = RED_ON Then
        m_RGY = GREEN_ON
        m_long = GREEN_LONG
    ElseIf m_RGY_Old = GREEN_ON Then
        m_RGY = RED_ON
        m_long = RED_LONG
    End If
    Call set_RGYLight(m_RGY)
    txtRGY.Caption = m_long
End Sub

'…Ë÷√µ∆¡¡∫Õœ®√
Private Sub set_RGYLight(ByVal rgy As Integer)
 shpRed.FillStyle = 0
 If rgy = RED_ON Then        '∫Ïµ∆¡¡
    shpRed.FillColor = vbRed
    txtRGY.ForeColor = vbRed
 Else
    shpRed.FillColor = vbBlack
 End If

shpGreen.FillStyle = 0
 If rgy = GREEN_ON Then        '∫Ïµ∆¡¡
    shpGreen.FillColor = vbGreen
    txtRGY.ForeColor = vbGreen
 Else
    shpGreen.FillColor = vbBlack
 End If
 
 shpYELLOW.FillStyle = 0
 If rgy = YELLOW_ON Then        '∫Ïµ∆¡¡
    shpYELLOW.FillColor = vbYellow
    txtRGY.ForeColor = vbYellow
 Else
    shpYELLOW.FillColor = vbBlack
 End If
End Sub
