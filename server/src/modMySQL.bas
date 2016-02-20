Attribute VB_Name = "modMySQL"
Option Explicit
Public conexaoBanco As ADODB.Connection
Public resultado As ADODB.Recordset

Public Sub conectarBanco(Host As String, Port As Integer, Database As String, User As String, Pass As String)
    Set conexaoBanco = New ADODB.Connection
    
    conexaoBanco.ConnectionString = "DRIVER={MySql ODBC 3.51 Driver};SERVER=" & Host & ";Port=" & Port & ";DATABASE=" & Database & ";UID=" & User & ";PWD=" & Pass & "; OPTION=3"
    conexaoBanco.Open
End Sub

Function query(ByVal SQL As String) As ADODB.Recordset
    Set resultado = New ADODB.Recordset
 
    resultado.CursorLocation = adUseServer
    resultado.CursorType = adOpenDynamic
    resultado.LockType = adLockReadOnly
    resultado.Open SQL, conexao
    
    If resultado.EOF Then
        resultado.Close
        Exit Function
    End If
    
    query = resultado
    
End Function

Public Sub encerrarConexao()
    Tabela.Close
End Sub

Public Sub Teste()
    query ("Select nome from teste")
    MsgBox resultado!Nome
End Sub


