VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4932
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   4932
   ScaleWidth      =   6540
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "创建目录"
      Height          =   372
      Left            =   5520
      TabIndex        =   3
      Top             =   1920
      Width           =   852
   End
   Begin VB.CommandButton Command2 
      Caption         =   "读"
      Height          =   372
      Left            =   5520
      TabIndex        =   2
      Top             =   1080
      Width           =   852
   End
   Begin VB.CommandButton Command1 
      Caption         =   "写"
      Height          =   372
      Left            =   5520
      TabIndex        =   1
      Top             =   240
      Width           =   852
   End
   Begin VB.TextBox Text1 
      Height          =   4692
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   5172
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intFileNumber As Integer
Private Sub Command1_Click()
     intFileNumber = FreeFile
     Open App.Path + "\Sample27.txt" For Append As #intFileNumber
     Write #intFileNumber, Text1.Text,
     Close #intFileNumber
     MsgBox "ok"
End Sub

Private Sub Command2_Click()
'逐行读
    'Dim InputData As Variant
    'Text1.Text = ""
    '  intFileNumber = FreeFile
    '      Open App.Path + "\Sample27.txt" For Input As #intFileNumber
    '      Do While Not EOF(intFileNumber)
    '        Line Input #intFileNumber, InputData
    '        Text1.Text = Text1.Text + InputData + vbCrLf
    '      Loop
    '      Close #intFileNumber
    
    '一次性读
    Text1.Text = ""
  intFileNumber = FreeFile
      Open App.Path + "\Sample27.txt" For Input As #intFileNumber
      Text1.Text = Input(LOF(1), 1)
      Close #intFileNumber
      MsgBox "ok"
End Sub

Private Sub Command3_Click()
    MkDir (App.Path + "\log")
    MsgBox "ok"
End Sub
