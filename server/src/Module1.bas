Attribute VB_Name = "modMySQL"
Public MySql As ADODB.Connection
Public resultadoSql As ADODB.Recordset
Public Query As String
Public ConnectedToDatabase As Boolean

Public Sub ConnectToDatabase()
    Set MySql = New ADODB.Connection
    MySql.ConnectionString = "Driver={Mysql ODBC 3.51 Driver}; Server=localhost;port=3306; database=ringex; user=ringex; password=ringex; option=3;"
    MySql.Open
    ConnectedToDatabase = (MySql.State = adStateOpen)
End Sub

Public Sub RunQuery()

    If Not ConnectedToDatabase Then Exit Sub
    
    Set resultadoSql = MySql.Execute(Query)

End Sub


'While Not resultadoSql.EOF
       ' MsgBox resultadoSql("gameName") & " - " & resultadoSql("motd")
    '    resultadoSql.MoveNext
    'Wend
