Attribute VB_Name = "modLanguage"
Public Sub Language()

    Select Case tmpCurLanguage

        ' Portugu�s
    Case 0

        ' Janlea de login
        TextUILoginUsername = "Usu�rio"
        TextUILoginPassword = "Senha"
        TextUILoginServerList = "Servidor"
        TextUILoginCheckBox = "Lembrar-me da senha?"
        TextUILoginEntryButton = "Entrar no PokeReborn"
        TextUILoginInvalidUsername = "Usu�rio inv�lido!"
        TextUILoginInvalidPassword = "Senha inv�lida!"

        ' Janela de registro
        TextUIRegisterUsername = "Usu�rio"
        TextUIRegisterPassword = "Senha"
        TextUIRegisterEmail = "Email"
        TextUIRegisterConfirm = "Finalizar cadastro"
        TextUIRegisterCheckBox = "Mostrar a senha?"
        TextUIRegisterUsernameLenght = "Seu nome de usu�rio deve estar entre 3 e " & (NAME_LENGTH - 1) & " caracteres e somente letras, n�meros e _ s�o permitidos."
        TextUIRegisterPasswordLenght = "Sua senha deve estar entre " & ((NAME_LENGTH - 1) / 4) & " and " & (NAME_LENGTH - 1) & "  caracteres e somente letras, n�meros e _ s�o permitidos."
        TextUIRegisterPasswordMatch = "A senhas n�o s�o iguais."
        TExtUIRegisterInvalidEmail = "Email inv�lido"

        ' Janela de cria��o de personagem
        TextUICreateCharacterCreateButton = "Finalizar personagem"
        TextUICreateCharacterUsername = "Nome"
        TextUICreateCharacterUsernameLenght = "Seu nome de personagem deve estar entre 3 e " & (NAME_LENGTH - 1) & " caracteres e somente letras, n�meros e _ s�o permitidos"

        ' Mensagem universal
        TextUIWait = "Espere alguns segundos antes de tentar novamente."

        ' Footer
        TextUIFooterCreateAccount = "Criar uma nova conta"
        'If CreditVisible Then
        '    TextUIFooterCredits = "Fechar"
        'Else
        TextUIFooterCredits = "Vers�o 1.0.0"
        'End If
        TextUIFooterDeveloper = "� PokeReborn - Todos os direitos reservados 2023."
        TextUIFooterChangePassword = "Mudar a senha"

        ' Menu global
        TextUIGlobalMenuReturn = "Retornar"
        TextUIGlobalMenuOptions = "Op��es"
        TextUIGlobalMenuReturnMenu = "Menu Principal"
        TextUIGlobalMenuExit = "Sair"

        ' Menu de op��es
        TextUIOptionVideoButton = "V�deo"
        TextUIOptionSoundButton = "Sons"
        TextUIOptionGameButton = "Jogo"
        TextUIOptionControlButton = "Controle"
        TextUIOptionFullscreen = "Tela Cheia: "
        TextUIOptionMusic = "M�sica"
        TextUIOptionSound = "Sons"
        TextUIOptionPath = "Interface:"
        TextUIOptionsFps = "Mostra o fps"
        TextUIOptionsPing = "Mostra o ping"
        TextUIOptionsFast = "In�cio R�pido"
        TextUIOptionName = "Mostrar Nome"
        TextUIOptionPP = "Mostrar PP Bar ao atacar"
        TextUIOptionLanguage = "Tradu��o: "
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
        TextUIOptionPokeSlot1 = "Pok�mon 01"
        TextUIOptionPokeSlot2 = "Pok�mon 02"
        TextUIOptionPokeSlot3 = "Pok�mon 03"
        TextUIOptionPokeSlot4 = "Pok�mon 04"
        TextUIOptionPokeSlot5 = "Pok�mon 05"
        TextUIOptionPokeSlot6 = "Pok�mon 06"
        TextUIOptionHotbarSlot1 = "Hotbar 01"
        TextUIOptionHotbarSlot2 = "Hotbar 02"
        TextUIOptionHotbarSlot3 = "Hotbar 03"
        TextUIOptionHotbarSlot4 = "Hotbar 04"
        TextUIOptionHotbarSlot5 = "Hotbar 05"
        TextUIOptionInventory = "Invent�rio"
        TextUIOptionPokedex = "Pok�dex"
        TextUIOptionTrainer = "Trainer"
        TextUIOptionMap = "Mapa"
        TextUIOptionRank = "Rank"
        TextUIOptionShop = "Shop"
        TextUIOptionInteract = "Interagir"
        TextUIOptionConvoChoice1 = "Con. Escolha 1"
        TextUIOptionConvoChoice2 = "Con. Escolha 2"
        TextUIOptionConvoChoice3 = "Con. Escolha 3"
        TextUIOptionConvoChoice4 = "Con. Escolha 4"

        ' Janela de Sele��o de personagem
        TextUICharactersNew = "Novo Personagem"
        TextUICharactersNone = "Vazio"
        TextUICharactersUse = "Usar"
        TextUICharactersDelete = "Del"

        ' Chat
        TextEnterToChat = "Aperte ENTER para digitar no chat"

        ' Ingl�s
    Case 1

        ' Login Window
        TextUILoginUsername = "User"
        TextUILoginPassword = "Password"
        TextUILoginServerList = "Server List"
        TextUILoginCheckBox = "Remember me password?"
        TextUILoginEntryButton = "Log in to PokeReborn"
        TextUILoginInvalidUsername = "Invalid username!"
        TextUILoginInvalidPassword = "Invalid password!"

        ' Register Window
        TextUIRegisterUsername = "Usu�rio"
        TextUIRegisterPassword = "Senha"
        TextUIRegisterEmail = "Email"
        TextUIRegisterConfirm = "Finalizar cadastro"
        TextUIRegisterCheckBox = "Mostrar a senha?"
        TextUIRegisterUsernameLenght = "Your username must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers and _ allowed"
        TextUIRegisterPasswordLenght = "Your password must be between " & ((NAME_LENGTH - 1) \ 4) & " and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers _ allowed"
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
        'If CreditVisible Then
        '    TextUIFooterCredits = "Close"
        'Else
        TextUIFooterCredits = "Version 1.0.0"
        'End If
        TextUIFooterDeveloper = "� PokeReborn - Todos os direitos reservados 2023."
        TextUIFooterChangePassword = "Change password"

        ' Menu de op��es
        TextUIOptionVideoButton = "V�deo"
        TextUIOptionSoundButton = "Sons"
        TextUIOptionGameButton = "Jogo"
        TextUIOptionControlButton = "Controle"
        TextUIOptionFullscreen = "Tela Cheia: "
        TextUIOptionMusic = "M�sica"
        TextUIOptionSound = "Sons"
        TextUIOptionPath = "Interface:"
        TextUIOptionsFps = "Mostra o fps"
        TextUIOptionsPing = "Mostra o ping"
        TextUIOptionsFast = "In�cio R�pido"
        TextUIOptionName = "Mostrar Nome"
        TextUIOptionPP = "Mostrar PP Bar ao atacar"
        TextUIOptionLanguage = "Tradu��o: "
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
        TextUIOptionPokeSlot1 = "Pok�mon 01"
        TextUIOptionPokeSlot2 = "Pok�mon 02"
        TextUIOptionPokeSlot3 = "Pok�mon 03"
        TextUIOptionPokeSlot4 = "Pok�mon 04"
        TextUIOptionPokeSlot5 = "Pok�mon 05"
        TextUIOptionPokeSlot6 = "Pok�mon 06"
        TextUIOptionHotbarSlot1 = "Hotbar 01"
        TextUIOptionHotbarSlot2 = "Hotbar 02"
        TextUIOptionHotbarSlot3 = "Hotbar 03"
        TextUIOptionHotbarSlot4 = "Hotbar 04"
        TextUIOptionHotbarSlot5 = "Hotbar 05"
        TextUIOptionInventory = "Invent�rio"
        TextUIOptionPokedex = "Pok�dex"
        TextUIOptionTrainer = "Trainer"
        TextUIOptionMap = "Mapa"
        TextUIOptionRank = "Rank"
        TextUIOptionShop = "Shop"
        TextUIOptionInteract = "Interagir"
        TextUIOptionConvoChoice1 = "Con. Escolha 1"
        TextUIOptionConvoChoice2 = "Con. Escolha 2"
        TextUIOptionConvoChoice3 = "Con. Escolha 3"
        TextUIOptionConvoChoice4 = "Con. Escolha 4"

        ' Espanhol
    Case 2

        ' Login Window
        TextUILoginUsername = "Usuario"
        TextUILoginPassword = "Contrase�a"
        TextUILoginServerList = "Servidor"
        TextUILoginCheckBox = "Olvid� mi contrase�a?"
        TextUILoginEntryButton = "Entrar a PokeReborn"
        TextUILoginInvalidUsername = "Usuario incorrecto!"
        TextUILoginInvalidPassword = "Contrase�a incorrecta!"

        ' Register Window
        TextUIRegisterUsername = "Usuario"
        TextUIRegisterPassword = "Contrase�a"
        TextUIRegisterEmail = "Email"
        TextUIRegisterConfirm = "Finalizar el registro"
        TextUIRegisterCheckBox = "Mostrar contrase�a?"
        TextUIRegisterUsernameLenght = "Tu nombre de usuario debe tener entre 3 y " & (NAME_LENGTH - 1) & " caracteres de largo, solo se permiten letras y n�meros."
        TextUIRegisterPasswordLenght = "Tu contrase�a debe tener m�nimo " & ((NAME_LENGTH - 1) / 4) & " and " & (NAME_LENGTH - 1) & " caracteres de largo, solo se permiten letras y n�meros."
        TextUIRegisterPasswordMatch = "Las contrase�as no coinciden"
        TExtUIRegisterInvalidEmail = "Email no v�lido"

        ' Create Character Window
        TextUICreateCharacterCreateButton = "Finalizar el car�cter"
        TextUICreateCharacterUsername = "Nombre"
        TextUICreateCharacterUsernameLenght = "El nombre de tu personaje debe estar entre 3 y " & (NAME_LENGTH - 1) & " caracteres de largo, solo se permiten letras y n�meros."


        'Text Universal
        TextUIWait = "Espera unos segundos antes de intentar de nuevo."

        ' Footer
        TextUIFooterCreateAccount = "Crear una cuenta"
        'If CreditVisible Then
        '    TextUIFooterCredits = "Cerrar"
        'Else
        TextUIFooterCredits = "Version 1.0.0"
        'End If
        TextUIFooterDeveloper = "� PokeReborn - Todos os direitos reservados 2023."
        TextUIFooterChangePassword = "Cambiar contrase�a"

        ' Menu de op��es
        TextUIOptionVideoButton = "V�deo"
        TextUIOptionSoundButton = "Sons"
        TextUIOptionGameButton = "Jogo"
        TextUIOptionControlButton = "Controle"
        TextUIOptionFullscreen = "Tela Cheia: "
        TextUIOptionMusic = "M�sica"
        TextUIOptionSound = "Sons"
        TextUIOptionPath = "Interface:"
        TextUIOptionsFps = "Mostra o fps"
        TextUIOptionsPing = "Mostra o ping"
        TextUIOptionsFast = "In�cio R�pido"
        TextUIOptionName = "Mostrar Nome"
        TextUIOptionPP = "Mostrar PP Bar ao atacar"
        TextUIOptionLanguage = "Tradu��o: "
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
        TextUIOptionPokeSlot1 = "Pok�mon 01"
        TextUIOptionPokeSlot2 = "Pok�mon 02"
        TextUIOptionPokeSlot3 = "Pok�mon 03"
        TextUIOptionPokeSlot4 = "Pok�mon 04"
        TextUIOptionPokeSlot5 = "Pok�mon 05"
        TextUIOptionPokeSlot6 = "Pok�mon 06"
        TextUIOptionHotbarSlot1 = "Hotbar 01"
        TextUIOptionHotbarSlot2 = "Hotbar 02"
        TextUIOptionHotbarSlot3 = "Hotbar 03"
        TextUIOptionHotbarSlot4 = "Hotbar 04"
        TextUIOptionHotbarSlot5 = "Hotbar 05"
        TextUIOptionInventory = "Invent�rio"
        TextUIOptionPokedex = "Pok�dex"
        TextUIOptionTrainer = "Trainer"
        TextUIOptionMap = "Mapa"
        TextUIOptionRank = "Rank"
        TextUIOptionShop = "Shop"
        TextUIOptionInteract = "Interagir"
        TextUIOptionConvoChoice1 = "Con. Escolha 1"
        TextUIOptionConvoChoice2 = "Con. Escolha 2"
        TextUIOptionConvoChoice3 = "Con. Escolha 3"
        TextUIOptionConvoChoice4 = "Con. Escolha 4"
    End Select

End Sub
