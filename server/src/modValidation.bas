Attribute VB_Name = "modValidation"
Public MENU_ERRO_MSG As String

Public Function validarNovoChar(ByVal Nome As String, ByVal Sexo As Long, ByVal Classe As Long) As Boolean

    If isShuttingDown Then
        MENU_ERRO_MSG = "N�o � poss�vel adicionar um personagem agora, o servidor est� sendo desligado ou reiniciado!"
        validarNovoChar = False
        Exit Function
    End If
    
    If Len(Nome) < 3 Then
        MENU_ERRO_MSG = "Nome deve ter no m�nimo 3 caracteres!"
        validarNovoChar = False
        Exit Function
    End If
    
    If Len(Nome) > 12 Then
        MENU_ERRO_MSG = "Nome deve ter no m�ximo 12 caracteres!"
        validarNovoChar = False
        Exit Function
    End If
    
    If Not validarCaracteresNome(Nome) Then
        MENU_ERRO_MSG = "Nome inv�lido. Use somente letras, n�meros, espa�os e _"
        validarNovoChar = False
        Exit Function
    End If
    
    If FindChar(Nome) Then
        MENU_ERRO_MSG = "Este nome j� est� em uso!"
        validarNovoChar = False
        Exit Function
    End If
    
    If (Sexo < SEX_MALE) Or (Sexo > SEX_FEMALE) Then
        MENU_ERRO_MSG = "Sexo inv�lido!"
        validarNovoChar = False
        Exit Function
    End If

    If Classe < 1 Or Classe > Max_Classes Then
        MENU_ERRO_MSG = "Ocorreu um erro, tente novamente!"
        validarNovoChar = False
        Exit Function
    End If

    validarNovoChar = True
End Function

Public Function validarNovaConta(ByVal Usuario As String, ByVal Email As String, ByVal Senha As String, ByVal Senha2 As String, ByVal CaptchaGerado As String, ByVal CaptchaDigitado As String) As Boolean
    
    If isShuttingDown Then
        MENU_ERRO_MSG = "N�o � poss�vel cadastrar-se agora, o servidor est� sendo desligado ou reiniciado!"
        validarNovaConta = False
        Exit Function
    End If
    
    If Len(Usuario) < 5 Then
        MENU_ERRO_MSG = "Usu�rio deve ter no m�nimo 5 caracteres!"
        validarNovaConta = False
        Exit Function
    End If
    
    If Len(Usuario) > 12 Then
        MENU_ERRO_MSG = "Usu�rio deve ter no m�ximo 12 caracteres!"
        validarNovaConta = False
        Exit Function
    End If
    
    If Not validarCaracteresNome(Usuario) Then
        MENU_ERRO_MSG = "Usu�rio inv�lido. Use somente letras, n�meros, espa�os e _"
        validarNovaConta = False
        Exit Function
    End If
    
    If AccountExist(Usuario) Then
        MENU_ERRO_MSG = "Usu�rio j� cadastrado!"
        validarNovaConta = False
        Exit Function
    End If
    
    If InStr(Email, "@") = 0 Or InStr(Email, ".") = 0 Or Len(Email) < 5 Then
        MENU_ERRO_MSG = "Email inv�lido!"
        validarNovaConta = False
        Exit Function
    End If
    
    If Len(Email) > 50 Then
        MENU_ERRO_MSG = "Email deve ter no m�ximo 50 caracteres!"
        validarNovaConta = False
        Exit Function
    End If
    
    If Len(Senha) < 5 Then
        MENU_ERRO_MSG = "Senha deve ter no m�nimo 5 caracteres!"
        validarNovaConta = False
        Exit Function
    End If
    
    If Len(Senha) > 12 Then
        MENU_ERRO_MSG = "Senha deve ter no m�ximo 12 caracteres!"
        validarNovaConta = False
        Exit Function
    End If
                
    If Senha <> Senha2 Then
        MENU_ERRO_MSG = "Senhas n�o conferem!"
        validarNovaConta = False
        Exit Function
    End If
   
    If Len(CaptchaDigitado) < 4 Then
        MENU_ERRO_MSG = "Captcha incorreto!"
        validarNovaConta = False
        Exit Function
    End If
     
    If CaptchaGerado <> CaptchaDigitado Then
        MENU_ERRO_MSG = "Captcha incorreto!"
        validarNovaConta = False
        Exit Function
    End If
    
    validarNovaConta = True
End Function


Public Function validarLogin(ByVal Usuario As String, ByVal Senha As String) As Boolean

    If isShuttingDown Then
        MENU_ERRO_MSG = "N�o � poss�vel conectar agora, o servidor est� sendo desligado ou reiniciado!"
        validarLogin = False
        Exit Function
    End If
    
    If Len(Usuario) < 5 Then
        MENU_ERRO_MSG = "Usu�rio deve ter no m�nimo 5 caracteres!"
        validarLogin = False
        Exit Function
    End If
    
    If Len(Usuario) > 12 Then
        MENU_ERRO_MSG = "Usu�rio deve ter no m�ximo 12 caracteres!"
        validarLogin = False
        Exit Function
    End If
    
    If Not validarCaracteresNome(Usuario) Then
        MENU_ERRO_MSG = "Usu�rio inv�lido. Use somente letras, n�meros, espa�os e _"
        validarLogin = False
        Exit Function
    End If
    
    If Not AccountExist(Usuario) Then
        MENU_ERRO_MSG = "Este usu�rio n�o existe!"
        validarLogin = False
        Exit Function
    End If
    
    If Len(Senha) < 5 Then
        MENU_ERRO_MSG = "Senha deve ter no m�nimo 5 caracteres!"
        validarLogin = False
        Exit Function
    End If
    
    If Len(Senha) > 12 Then
        MENU_ERRO_MSG = "Senha deve ter no m�ximo 12 caracteres!"
        validarLogin = False
        Exit Function
    End If
    
    If Not PasswordOK(Usuario, Senha) Then
        MENU_ERRO_MSG = "Senha incorreta!"
        validarLogin = False
        Exit Function
    End If
    
    If IsBanned(Usuario, True) Then
        MENU_ERRO_MSG = "Voc� est� banido por tempo indeterminado e n�o pode jogar!"
        validarLogin = False
        Exit Function
    End If
    
    If IsMultiAccounts(Usuario) Then
        MENU_ERRO_MSG = "Este usu�rio j� est� conectado!"
        validarLogin = False
        Exit Function
    End If
    
    validarLogin = True
End Function

Public Function validarCaracteresNome(ByVal palavra As String) As Boolean
Dim n As Long
Dim i As Long

    For i = 1 To Len(palavra)
        n = AscW(Mid$(palavra, i, 1))

        If Not ((n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57)) Then
            validarCaracteresNome = False
            Exit Function
        End If
    Next
    
    validarCaracteresNome = True
            
End Function
