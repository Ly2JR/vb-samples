VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3300
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   6780
   StartUpPosition =   3  '窗口缺省
   Begin VB.FileListBox File1 
      Height          =   2952
      Left            =   1920
      TabIndex        =   4
      Top             =   240
      Width           =   1692
   End
   Begin VB.ComboBox Combo1 
      Height          =   276
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2880
      Width           =   1572
   End
   Begin VB.DirListBox Dir1 
      Height          =   1908
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1572
   End
   Begin VB.DriveListBox Drive1 
      Height          =   276
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   3012
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3012
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "文件类型"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Change()
    File1.Pattern = Combo1.Text
End Sub

Private Sub Combo1_Click()
File1.Pattern = Combo1.Text
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
     File1.Refresh
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    Image1.Picture = LoadPicture(File1.Path + "\" + File1.FileName)
End Sub

Private Sub Form_Load()
    Combo1.AddItem "*.bmp"
    Combo1.AddItem "*.tif"
    Combo1.AddItem "*.*"
    Combo1.Text = Combo1.List(0)
    File1.Pattern = "*.jpg"
End Sub
