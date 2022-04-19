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
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuNew 
         Caption         =   "新建(&N)"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "打开(&O)"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "保存(&S)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
   End
   Begin VB.Menu muFormat 
      Caption         =   "格式"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

