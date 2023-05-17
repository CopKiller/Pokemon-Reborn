Attribute VB_Name = "modLanguage"
Public Sub Language()

    Select Case tmpCurLanguage
        
        ' Português
        Case 0
        
            ' Janlea de login
            TextUILoginUsername = "Usuário"
            TextUILoginPassword = "Senha"
            TextUILoginServerList = "Servidor"
            TextUILoginCheckBox = "Lembrar-me da senha?"
            TextUILoginEntryButton = "Entrar no Pokenew"
            TextUILoginInvalidUsername = "Usuário inválido!"
            TextUILoginInvalidPassword = "Senha inválida!"
            
            ' Janela de registro
            TextUIRegisterUsername = "Usuário"
            TextUIRegisterPassword = "Senha"
            TextUIRegisterEmail = "Email"
            TextUIRegisterConfirm = "Finalizar cadastro"
            TextUIRegisterCheckBox = "Mostrar a senha?"
            TextUIRegisterUsernameLenght = "Seu nome de usuário deve estar entre 3 e " & (NAME_LENGTH - 1) & " caracteres e somente letras, números e _ são permitidos."
            TextUIRegisterPasswordLenght = "Sua senha deve estar entre " & ((NAME_LENGTH - 1) / 4) & " and " & (NAME_LENGTH - 1) & "  caracteres e somente letras, números e _ são permitidos."
            TextUIRegisterPasswordMatch = "A senhas não são iguais."
            TExtUIRegisterInvalidEmail = "Email inválido"
            
            ' Janela de criação de personagem
            TextUICreateCharacterCreateButton = "Finalizar personagem"
            TextUICreateCharacterUsername = "Nome"
            TextUICreateCharacterUsernameLenght = "Seu nome de personagem deve estar entre 3 e " & (NAME_LENGTH - 1) & " caracteres e somente letras, números e _ são permitidos"
            
            ' Mensagem universal
            TextUIWait = "Espere alguns segundos antes de tentar novamente."
            
            ' Footer
            TextUIFooterCreateAccount = "Criar uma nova conta"
            If CreditVisible Then
                TextUIFooterCredits = "Fechar"
            Else
                TextUIFooterCredits = "Créditos"
            End If
            TextUIFooterDeveloper = "© Matheus R de Oliveira - 2022 ~ 2023. Todos os direitos reservados."
            TextUIFooterChangePassword = "Mudar a senha"
            
            ' Menu global
            TextUIGlobalMenuReturn = "Retornar"
            TextUIGlobalMenuOptions = "Opções"
            TextUIGlobalMenuReturnMenu = "Menu Principal"
            TextUIGlobalMenuExit = "Sair"
            
            ' Menu de opções
            TextUIOptionVideoButton = "Vídeo"
            TextUIOptionSoundButton = "Sons"
            TextUIOptionGameButton = "Jogo"
            TextUIOptionControlButton = "Controle"
            TextUIOptionFullscreen = "Tela Cheia: "
            TextUIOptionMusic = "Música"
            TextUIOptionSound = "Sons"
            TextUIOptionPath = "Interface:"
            TextUIOptionsFps = "Mostra o fps"
            TextUIOptionsPing = "Mostra o ping"
            TextUIOptionsFast = "Início Rápido"
            TextUIOptionName = "Mostrar Nome"
            TextUIOptionLanguage = "Tradução: "
            TextUIOptionUp = "Subir"
            TextUIOptionDown = "Abaixo"
            TextUIOptionLeft = "Esquerda"
            TextUIOptionRight = "Direita"
            TextUIOptionCheckMove = "Movimentos"
            TextUIOptionMoveSlot1 = "Movimento 01"
            TextUIOptionMoveSlot2 = "Movimento 02"
            TextUIOptionMoveSlot3 = "Movimento 03"
            TextUIOptionMoveSlot4 = "Movimento 04"
            TextUIOptionAttack = "Atacar"
            TextUIOptionPokeSlot1 = "Pokémon 01"
            TextUIOptionPokeSlot2 = "Pokémon 02"
            TextUIOptionPokeSlot3 = "Pokémon 03"
            TextUIOptionPokeSlot4 = "Pokémon 04"
            TextUIOptionPokeSlot5 = "Pokémon 05"
            TextUIOptionPokeSlot6 = "Pokémon 06"
            TextUIOptionHotbarSlot1 = "Hotbar 01"
            TextUIOptionHotbarSlot2 = "Hotbar 02"
            TextUIOptionHotbarSlot3 = "Hotbar 03"
            TextUIOptionHotbarSlot4 = "Hotbar 04"
            TextUIOptionHotbarSlot5 = "Hotbar 05"
            TextUIOptionInventory = "Inventário"
            TextUIOptionPokedex = "Pokédex"
            TextUIOptionInteract = "Interagir"
            TextUIOptionConvoChoice1 = "Con. Escolha 1"
            TextUIOptionConvoChoice2 = "Con. Escolha 2"
            TextUIOptionConvoChoice3 = "Con. Escolha 3"
            TextUIOptionConvoChoice4 = "Con. Escolha 4"
            
            ' Janela de Seleção de personagem
            TextUICharactersNew = "Novo Personagem"
            TextUICharactersNone = "Vazio"
            TextUICharactersUse = "Usar"
            TextUICharactersDelete = "Del"
            
            ' Chat
            TextEnterToChat = "Aperte ENTER para digitar no chat"
            
        ' Inglês
        Case 1
        
            ' Login Window
            TextUILoginUsername = "User"
            TextUILoginPassword = "Password"
            TextUILoginServerList = "Server List"
            TextUILoginCheckBox = "Remember me password?"
            TextUILoginEntryButton = "Log in to Pokenew"
            TextUILoginInvalidUsername = "Invalid username!"
            TextUILoginInvalidPassword = "Invalid password!"
            
            ' Register Window
            TextUIRegisterUsername = "Usuário"
            TextUIRegisterPassword = "Senha"
            TextUIRegisterEmail = "Email"
            TextUIRegisterConfirm = "Finalizar cadastro"
            TextUIRegisterCheckBox = "Mostrar a senha?"
            TextUIRegisterUsernameLenght = "Your username must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers and _ allowed"
            TextUIRegisterPasswordLenght = "Your password must be between " & ((NAME_LENGTH - 1) / 4) & " and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers _ allowed"
            TextUIRegisterPasswordMatch = "Password did not match"
            TExtUIRegisterInvalidEmail = "Invalid email"
            
            ' Create Character Window
            TextUICreateCharacterCreateButton = "Create Character"
            TextUICreateCharacterUsername = "Name"
            TextUICreateCharacterUsernameLenght = "Your character name must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers and _ allowed"
                        
            'Text Universal
            TextUIWait = "Wait a few seconds before trying again."
            
            ' Footer
            TextUIFooterCreateAccount = "Create an account"
            If CreditVisible Then
                TextUIFooterCredits = "Close"
            Else
                TextUIFooterCredits = "Credits"
            End If
            TextUIFooterDeveloper = "© Matheus R de Oliveira - 2022 ~ 2023. All rights reserved."
            TextUIFooterChangePassword = "Change password"
            
        ' Espanhol
        Case 2
            
            ' Login Window
            TextUILoginUsername = "Usuario"
            TextUILoginPassword = "Contraseña"
            TextUILoginServerList = "Servidor"
            TextUILoginCheckBox = "Olvidé mi contraseña?"
            TextUILoginEntryButton = "Entrar a Pokenew"
            TextUILoginInvalidUsername = "Usuario incorrecto!"
            TextUILoginInvalidPassword = "Contraseña incorrecta!"
            
            ' Register Window
            TextUIRegisterUsername = "Usuario"
            TextUIRegisterPassword = "Contraseña"
            TextUIRegisterEmail = "Email"
            TextUIRegisterConfirm = "Finalizar el registro"
            TextUIRegisterCheckBox = "Mostrar contraseña?"
            TextUIRegisterUsernameLenght = "Tu nombre de usuario debe tener entre 3 y " & (NAME_LENGTH - 1) & " caracteres de largo, solo se permiten letras y números."
            TextUIRegisterPasswordLenght = "Tu contraseña debe tener mínimo " & ((NAME_LENGTH - 1) / 4) & " and " & (NAME_LENGTH - 1) & " caracteres de largo, solo se permiten letras y números."
            TextUIRegisterPasswordMatch = "Las contraseñas no coinciden"
            TExtUIRegisterInvalidEmail = "Email no válido"
            
            ' Create Character Window
            TextUICreateCharacterCreateButton = "Finalizar el carácter"
            TextUICreateCharacterUsername = "Nombre"
            TextUICreateCharacterUsernameLenght = "El nombre de tu personaje debe estar entre 3 y " & (NAME_LENGTH - 1) & " caracteres de largo, solo se permiten letras y números."
            
            
            'Text Universal
            TextUIWait = "Espera unos segundos antes de intentar de nuevo."
            
            ' Footer
            TextUIFooterCreateAccount = "Crear una cuenta"
            If CreditVisible Then
                TextUIFooterCredits = "Cerrar"
            Else
                TextUIFooterCredits = "Créditos"
            End If
            TextUIFooterDeveloper = "© Matheus R de Oliveira - 2022 ~ 2023. Todos los derechos reservados."
            TextUIFooterChangePassword = "Cambiar contraseña"
    End Select

End Sub
