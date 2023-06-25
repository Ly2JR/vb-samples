VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "testCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private id As Integer

Private content As String

Private rest As ADODB.Recordset

Public Property Get TestContent() As String
    TestContent = content
End Property


Public Property Let TestContent(vData As String)
    content = vData
End Property

Public Property Get TestDescription() As String
TestDescription = ""
    If Not rest Is Nothing Then
        If Not rest.EOF Then
            TestDescription = rest!id & "  " & rest!content
        End If
    End If
End Property

Public Property Get TestRecordSet() As ADODB.Recordset
    Set TestRecordSet = rest
End Property

Public Property Set TestRecordSet(vData As ADODB.Recordset)
    Set rest = vData
End Property


Private Sub Class_Initialize()
    If rest Is Nothing Then
        Set rest = New ADODB.Recordset
    End If
End Sub

Private Sub Class_Terminate()
    If Not rest Is Nothing Then
        If rest.State = 1 Then rest.Close
        Set rest = Nothing
    End If
End Sub
