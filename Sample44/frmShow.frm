VERSION 5.00
Begin VB.Form frmShow 
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   8448
   LinkTopic       =   "Form1"
   ScaleHeight     =   12420
   ScaleWidth      =   22824
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txt奖学金 
      DataField       =   "奖学金"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2592
      TabIndex        =   11
      Top             =   2484
      Width           =   660
   End
   Begin VB.TextBox txt所在系 
      DataField       =   "所在系"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2592
      TabIndex        =   9
      Top             =   2112
      Width           =   3300
   End
   Begin VB.TextBox txt出生日期 
      DataField       =   "出生日期"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2592
      TabIndex        =   7
      Top             =   1728
      Width           =   1320
   End
   Begin VB.TextBox txt性别 
      DataField       =   "性别"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2592
      TabIndex        =   5
      Top             =   1344
      Width           =   330
   End
   Begin VB.TextBox txt姓名 
      DataField       =   "姓名"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2592
      TabIndex        =   3
      Top             =   972
      Width           =   1650
   End
   Begin VB.TextBox txt学号 
      DataField       =   "学号"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2592
      TabIndex        =   1
      Top             =   588
      Width           =   825
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "奖学金:"
      Height          =   252
      Index           =   5
      Left            =   744
      TabIndex        =   10
      Top             =   2532
      Width           =   1812
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "所在系:"
      Height          =   252
      Index           =   4
      Left            =   744
      TabIndex        =   8
      Top             =   2148
      Width           =   1812
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "出生日期:"
      Height          =   252
      Index           =   3
      Left            =   744
      TabIndex        =   6
      Top             =   1776
      Width           =   1812
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "性别:"
      Height          =   252
      Index           =   2
      Left            =   744
      TabIndex        =   4
      Top             =   1392
      Width           =   1812
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "姓名:"
      Height          =   252
      Index           =   1
      Left            =   744
      TabIndex        =   2
      Top             =   1008
      Width           =   1812
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "学号:"
      Height          =   252
      Index           =   0
      Left            =   744
      TabIndex        =   0
      Top             =   636
      Width           =   1812
   End
End
Attribute VB_Name = "frmShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

