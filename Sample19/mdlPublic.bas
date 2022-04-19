Attribute VB_Name = "mdlPublic"
Option Explicit

'自定义数据类型
Public Type StuType
    No As Integer       '学号
    Name As String * 20 '姓名
    Sex As String * 1   '性别
    Mark(1 To 4) As Single '4门课程成绩
    Total As Single  '总分
End Type

Public Type MANType
    No As Integer '学号
    Name As String        '姓名
    Sex As String * 1   '性别
    Birthdate As Date   '出生年月
    Speciality As String    '特长
End Type

Public Type StudType
    Name As String * 10     '姓名
    Special As String * 10  '专业
    Total   As Single       '总分
    
End Type
