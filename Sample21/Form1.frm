VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2340
   ClientLeft      =   108
   ClientTop       =   732
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   3624
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuNew 
         Caption         =   "�½�(&N)"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "��(&O)"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "����(&S)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
   End
   Begin VB.Menu muFormat 
      Caption         =   "��ʽ"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

