VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "��ʾ�����ߵ�����"
   ClientHeight    =   2784
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   8136
   LinkTopic       =   "Form1"
   ScaleHeight     =   2784
   ScaleWidth      =   8136
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "������ʾ"
      Height          =   372
      Left            =   6840
      TabIndex        =   0
      Top             =   2280
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim i As Integer
    Print "DrawStle     0       1       2       3       4       5       6"
    Print
    Print "����        ʵ��   ������   ����   �㻭��  ��㻭��  ͸����  ��ʵ��"
    Print
    Print "ͼʾ   ";
    CurrentY = CurrentY + 200
    DrawWidth = 1
    For i = 0 To 6
        DrawStyle = i
        CurrentX = CurrentX + 150
        Line -Step(700, 0)
        
    Next i
    
End Sub
