Attribute VB_Name = "mdlGloabl"
Option Explicit

Private Const g_connString As String = "Provider=SQLOLEDB.1;Data Source=DESKTOP-CETRBK3\SQLEXPRESS;Initial Catalog=master;Uid=sa;Password=sa123"

Private g_conn As ADODB.Connection

Public Sub ConnectDb()
If g_conn Is Nothing Then
    Set g_conn = New ADODB.Connection
End If
If g_conn.State <> 1 Then
    g_conn.Open g_connString
End If
End Sub

Public Sub CloseDb()
If g_conn.State = 1 Then g_conn.Close
 Set g_conn = Nothing
End Sub


Public Function G_CRUD(ByVal sSql As String) As Integer
On Error GoTo ErrHandler:
Dim affected As Integer
G_CRUD = 0
ConnectDb
g_conn.Execute sSql, G_CRUD
finally:
    Exit Function
ErrHandler:
    MsgBox VBA.Err.Description
    GoTo finally
End Function

Public Function G_Query(ByVal sSql As String) As ADODB.Recordset
On Error GoTo ErrHandler:
Dim rest As New ADODB.Recordset
If rest.State = 1 Then rest.Close
ConnectDb
rest.Open sSql, g_conn, adOpenStatic, adLockReadOnly
Set G_Query = rest
finally:
    Exit Function
ErrHandler:
    MsgBox VBA.Err.Description
    GoTo finally
End Function




