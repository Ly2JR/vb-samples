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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txt��ѧ�� 
      DataField       =   "��ѧ��"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2592
      TabIndex        =   11
      Top             =   2484
      Width           =   660
   End
   Begin VB.TextBox txt����ϵ 
      DataField       =   "����ϵ"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2592
      TabIndex        =   9
      Top             =   2112
      Width           =   3300
   End
   Begin VB.TextBox txt�������� 
      DataField       =   "��������"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2592
      TabIndex        =   7
      Top             =   1728
      Width           =   1320
   End
   Begin VB.TextBox txt�Ա� 
      DataField       =   "�Ա�"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2592
      TabIndex        =   5
      Top             =   1344
      Width           =   330
   End
   Begin VB.TextBox txt���� 
      DataField       =   "����"
      DataMember      =   "Command1"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2592
      TabIndex        =   3
      Top             =   972
      Width           =   1650
   End
   Begin VB.TextBox txtѧ�� 
      DataField       =   "ѧ��"
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
      Caption         =   "��ѧ��:"
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
      Caption         =   "����ϵ:"
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
      Caption         =   "��������:"
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
      Caption         =   "�Ա�:"
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
      Caption         =   "����:"
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
      Caption         =   "ѧ��:"
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

