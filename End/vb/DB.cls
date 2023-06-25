VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const g_connString As String = "Provider=SQLOLEDB.1;Data Source=DESKTOP-CETRBK3\SQLEXPRESS;Initial Catalog=master;Uid=sa;Password=sa123"

Private content As String
Private g_conn As ADODB.Connection
Private g_res As ADODB.Recordset

Public Property Get TestContent() As String
    TestContent = content
End Property

Public Property Let TestContent(vData As String)
    content = vData
End Property

Public Function CRUD(ByVal sSql As String, ByRef errMsg As String) As Integer
On Error GoTo ErrHandler:
CRUD = 0
g_conn.Execute sSql, CRUD
finally:
    Exit Function
ErrHandler:
    errMsg = VBA.Err.Description
    GoTo finally
End Function

Public Function Add(errMsg As String) As Integer
On Error GoTo ErrHandler:
Dim sSql As String
Add = 0
sSql = "Insert into Test(content) values ('" & TestContent & "')"
g_conn.Execute sSql, Add
finally:
    Exit Function
ErrHandler:
   errMsg = VBA.Err.Description
    GoTo finally
End Function

Public Function Query(ByVal sSql As String, errMsg As String) As ADODB.Recordset
On Error GoTo ErrHandler:
If g_res.State = 1 Then g_res.Close
g_res.Open sSql, g_conn, adOpenStatic, adLockReadOnly
Set Query = g_res
finally:
    Exit Function
ErrHandler:
     errMsg = VBA.Err.Description
    GoTo finally
End Function


Private Sub Class_Initialize()
If g_conn Is Nothing Then
    Set g_conn = New ADODB.Connection
End If
If g_conn.State <> 1 Then
    g_conn.Open g_connString
End If
If g_res Is Nothing Then
    Set g_res = New ADODB.Recordset
End If
End Sub

Private Sub Class_Terminate()
If Not g_res Is Nothing Then
    If g_res.State = 1 Then g_res.Close
    Set g_res = Nothing
End If
If Not g_conn Is Nothing Then
    If g_conn.State = 1 Then g_conn.Close
    Set g_conn = Nothing
End If
End Sub

