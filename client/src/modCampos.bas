Attribute VB_Name = "modValidation"
Public Sub limparCamposCadastrar()
    txtRUsuario = vbNullString
    txtRSenha = vbNullString
    txtRSenha2 = vbNullString
    txtREmail = vbNullString
    txtRCaptcha = vbNullString
End Sub

Public Sub limparCamposLogin()
    txtLUsuario = vbNullString
    txtLSenha = vbNullString
End Sub

Public Sub carregarCamposLogin()
    txtLUsuario = Trim$(Servers(ServerIndex).Username)
    txtLSenha = Trim$(Servers(ServerIndex).Password)
End Sub

