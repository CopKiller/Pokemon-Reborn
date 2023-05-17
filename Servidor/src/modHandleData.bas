Attribute VB_Name = "modHandleData"
Option Explicit

Public Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(CCheckPing) = GetAddress(AddressOf HandleCheckPing)
    HandleDataSub(CNewAccount) = GetAddress(AddressOf HandleNewAccount)
    HandleDataSub(CLoginInfo) = GetAddress(AddressOf HandleLoginInfo)
    HandleDataSub(CNewCharacter) = GetAddress(AddressOf HandleNewCharacter)
    HandleDataSub(CUseCharacter) = GetAddress(AddressOf HandleUseCharacter)
    HandleDataSub(CDelCharacter) = GetAddress(AddressOf HandleDelCharacter)
    HandleDataSub(CNeedMap) = GetAddress(AddressOf HandleNeedMap)
    HandleDataSub(CPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(CPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(CMapMsg) = GetAddress(AddressOf HandleMapMsg)
    HandleDataSub(CGlobalMsg) = GetAddress(AddressOf HandleGlobalMsg)
    HandleDataSub(CPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(CWarpTo) = GetAddress(AddressOf HandleWarpTo)
    HandleDataSub(CAdminWarp) = GetAddress(AddressOf HandleAdminWarp)
    HandleDataSub(CWarpToMe) = GetAddress(AddressOf HandleWarpToMe)
    HandleDataSub(CWarpMeTo) = GetAddress(AddressOf HandleWarpMeTo)
    HandleDataSub(CSetAccess) = GetAddress(AddressOf HandleSetAccess)
    HandleDataSub(CPlayerPokemonMove) = GetAddress(AddressOf HandlePlayerPokemonMove)
    HandleDataSub(CPlayerPokemonDir) = GetAddress(AddressOf HandlePlayerPokemonDir)
    HandleDataSub(CGetItem) = GetAddress(AddressOf HandleGetItem)
    HandleDataSub(CPlayerPokemonState) = GetAddress(AddressOf HandlePlayerPokemonState)
    HandleDataSub(CAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(CChangePassword) = GetAddress(AddressOf HandleChangePassword)
    HandleDataSub(CReplaceNewMove) = GetAddress(AddressOf HandleReplaceNewMove)
    HandleDataSub(CEvolvePoke) = GetAddress(AddressOf HandleEvolvePoke)
    HandleDataSub(CUseItem) = GetAddress(AddressOf HandleUseItem)
    HandleDataSub(CSwitchInvSlot) = GetAddress(AddressOf HandleSwitchInvSlot)
    HandleDataSub(CGotData) = GetAddress(AddressOf HandleGotData)
    HandleDataSub(COpenStorage) = GetAddress(AddressOf HandleOpenStorage)
    HandleDataSub(CDepositItemTo) = GetAddress(AddressOf HandleDepositItemTo)
    HandleDataSub(CSwitchStorageSlot) = GetAddress(AddressOf HandleSwitchStorageSlot)
    HandleDataSub(CWithdrawItemTo) = GetAddress(AddressOf HandleWithdrawItemTo)
    HandleDataSub(CConvo) = GetAddress(AddressOf HandleConvo)
    HandleDataSub(CProcessConvo) = GetAddress(AddressOf HandleProcessConvo)
    HandleDataSub(CDepositPokemon) = GetAddress(AddressOf HandleDepositPokemon)
    HandleDataSub(CWithdrawPokemon) = GetAddress(AddressOf HandleWithdrawPokemon)
    HandleDataSub(CSwitchStoragePokeSlot) = GetAddress(AddressOf HandleSwitchStoragePokeSlot)
    HandleDataSub(CCloseShop) = GetAddress(AddressOf HandleCloseShop)
    HandleDataSub(CBuyItem) = GetAddress(AddressOf HandleBuyItem)
    HandleDataSub(CSellItem) = GetAddress(AddressOf HandleSellItem)
    HandleDataSub(CRequest) = GetAddress(AddressOf HandleRequest)
    HandleDataSub(CRequestState) = GetAddress(AddressOf HandleRequestState)
    HandleDataSub(CAddTrade) = GetAddress(AddressOf HandleAddTrade)
    HandleDataSub(CRemoveTrade) = GetAddress(AddressOf HandleRemoveTrade)
    HandleDataSub(CTradeUpdateMoney) = GetAddress(AddressOf HandleTradeUpdateMoney)
    HandleDataSub(CSetTradeState) = GetAddress(AddressOf HandleSetTradeState)
    HandleDataSub(CTradeState) = GetAddress(AddressOf HandleTradeState)
    HandleDataSub(CScanPokedex) = GetAddress(AddressOf HandleScanPokedex)
    HandleDataSub(CMOTD) = GetAddress(AddressOf HandleMOTD)
    HandleDataSub(CCopyMap) = GetAddress(AddressOf HandleCopyMap)
    HandleDataSub(CReleasePokemon) = GetAddress(AddressOf HandleReleasePokemon)
    HandleDataSub(CGiveItemTo) = GetAddress(AddressOf HandleGiveItemTo)
    HandleDataSub(CGivePokemonTo) = GetAddress(AddressOf HandleGivePokemonTo)
    HandleDataSub(CSpawnPokemon) = GetAddress(AddressOf HandleSpawnPokemon)
    HandleDataSub(CSetLanguage) = GetAddress(AddressOf HandleSetLanguage)
    HandleDataSub(CBuyStorageSlot) = GetAddress(AddressOf HandleBuyStorageSlot)
    HandleDataSub(CSellPokeStorageSlot) = GetAddress(AddressOf HandleSellPokeStorageSlot)
    HandleDataSub(CChangeShinyRate) = GetAddress(AddressOf HandleChangeShinyRate)
    HandleDataSub(CRelearnMove) = GetAddress(AddressOf HandleRelearnMove)
    HandleDataSub(CUseRevive) = GetAddress(AddressOf HandleUseRevive)
    HandleDataSub(CAddHeld) = GetAddress(AddressOf HandleAddHeld)
    HandleDataSub(CRemoveHeld) = GetAddress(AddressOf HandleRemoveHeld)
    HandleDataSub(CStealthMode) = GetAddress(AddressOf HandleStealthMode)
    HandleDataSub(CWhosOnline) = GetAddress(AddressOf HandleWhosOnline)
    HandleDataSub(CRequestRank) = GetAddress(AddressOf HandleRequestRank)
    HandleDataSub(CHotbarUpdate) = GetAddress(AddressOf HandleHotbarUpdate)
    HandleDataSub(CUseHotbar) = GetAddress(AddressOf HandleUseHotbar)
    HandleDataSub(CCreateParty) = GetAddress(AddressOf HandleCreateParty)
    HandleDataSub(CLeaveParty) = GetAddress(AddressOf HandleLeaveParty)
    '//Editors
    HandleDataSub(CRequestEditMap) = GetAddress(AddressOf HandleRequestEditMap)
    HandleDataSub(CMap) = GetAddress(AddressOf HandleMap)
    HandleDataSub(CRequestEditNpc) = GetAddress(AddressOf HandleRequestEditNpc)
    HandleDataSub(CRequestNpc) = GetAddress(AddressOf HandleRequestNpc)
    HandleDataSub(CSaveNpc) = GetAddress(AddressOf HandleSaveNpc)
    HandleDataSub(CRequestEditPokemon) = GetAddress(AddressOf HandleRequestEditPokemon)
    HandleDataSub(CRequestPokemon) = GetAddress(AddressOf HandleRequestPokemon)
    HandleDataSub(CSavePokemon) = GetAddress(AddressOf HandleSavePokemon)
    HandleDataSub(CRequestEditItem) = GetAddress(AddressOf HandleRequestEditItem)
    HandleDataSub(CRequestItem) = GetAddress(AddressOf HandleRequestItem)
    HandleDataSub(CSaveItem) = GetAddress(AddressOf HandleSaveItem)
    HandleDataSub(CRequestEditPokemonMove) = GetAddress(AddressOf HandleRequestEditPokemonMove)
    HandleDataSub(CRequestPokemonMove) = GetAddress(AddressOf HandleRequestPokemonMove)
    HandleDataSub(CSavePokemonMove) = GetAddress(AddressOf HandleSavePokemonMove)
    HandleDataSub(CRequestEditAnimation) = GetAddress(AddressOf HandleRequestEditAnimation)
    HandleDataSub(CRequestAnimation) = GetAddress(AddressOf HandleRequestAnimation)
    HandleDataSub(CSaveAnimation) = GetAddress(AddressOf HandleSaveAnimation)
    HandleDataSub(CRequestEditSpawn) = GetAddress(AddressOf HandleRequestEditSpawn)
    HandleDataSub(CRequestSpawn) = GetAddress(AddressOf HandleRequestSpawn)
    HandleDataSub(CSaveSpawn) = GetAddress(AddressOf HandleSaveSpawn)
    HandleDataSub(CRequestEditConversation) = GetAddress(AddressOf HandleRequestEditConversation)
    HandleDataSub(CRequestConversation) = GetAddress(AddressOf HandleRequestConversation)
    HandleDataSub(CSaveConversation) = GetAddress(AddressOf HandleSaveConversation)
    HandleDataSub(CRequestEditShop) = GetAddress(AddressOf HandleRequestEditShop)
    HandleDataSub(CRequestShop) = GetAddress(AddressOf HandleRequestShop)
    HandleDataSub(CSaveShop) = GetAddress(AddressOf HandleSaveShop)
    HandleDataSub(CRequestEditQuest) = GetAddress(AddressOf HandleRequestEditQuest)
    HandleDataSub(CRequestQuest) = GetAddress(AddressOf HandleRequestQuest)
    HandleDataSub(CSaveQuest) = GetAddress(AddressOf HandleSaveQuest)
    HandleDataSub(CKickPlayer) = GetAddress(AddressOf HandleKickPlayer)
    HandleDataSub(CBanPlayer) = GetAddress(AddressOf HandleBanPlayer)
    HandleDataSub(CMutePlayer) = GetAddress(AddressOf HandleMutePlayer)
    HandleDataSub(CUnmutePlayer) = GetAddress(AddressOf HandleUnmutePlayer)
    HandleDataSub(CFlyToBadge) = GetAddress(AddressOf HandleFlyToBadge)
End Sub

Public Sub HandleData(ByVal index As Long, ByRef Data() As Byte)
Dim buffer As clsBuffer
Dim MsgType As Long

    ' Prevent from receiving a empty data
    If ((Not Data) = -1) Or ((Not Data) = 0) Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MsgType = buffer.ReadLong

    If MsgType < 1 Or MsgType >= CMSG_Count Then
        buffer.Flush
        Set buffer = Nothing
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), index, buffer.ReadBytes(buffer.Length), 0, 0
    
    buffer.Flush
    Set buffer = Nothing
End Sub

Public Sub IncomingData(ByVal index As Long, ByVal DataLength As Long)
Dim buffer() As Byte
Dim pLength As Long
Dim RemoveLimit As Byte

    RemoveLimit = NO
    If IsPlaying(index) And TempPlayer(index).UseChar > 0 Then
        If Player(index, TempPlayer(index).UseChar).Access > ACCESS_NONE Then
            RemoveLimit = YES
        End If
    End If
    
    If RemoveLimit = NO Then
        ' Check for data flooding
        If TempPlayer(index).DataBytes > 1000 Then
            If GetTickCount < TempPlayer(index).DataTimer Then
                Exit Sub
            End If
        End If
        
        ' Check for packet flooding
        If TempPlayer(index).DataPackets > 25 Then
            If GetTickCount < TempPlayer(index).DataTimer Then
                Exit Sub
            End If
        End If
    End If

    ' Check if elapsed time has passed
    TempPlayer(index).DataBytes = TempPlayer(index).DataBytes + DataLength
    If GetTickCount >= TempPlayer(index).DataTimer Then
        TempPlayer(index).DataTimer = GetTickCount + 1000
        TempPlayer(index).DataBytes = 0
        TempPlayer(index).DataPackets = 0
    End If
    
    ' Get the data from the socket now
    frmServer.Socket(index).GetData buffer(), vbUnicode, DataLength
    
    ' Prevent from receiving a empty data
    If ((Not buffer) = -1) Or ((Not buffer) = 0) Then Exit Sub
    
    TempPlayer(index).buffer.WriteBytes buffer()
    
    If TempPlayer(index).buffer.Length >= 4 Then
        pLength = TempPlayer(index).buffer.ReadLong(False)
        If pLength < 0 Then Exit Sub
    End If
    Do While pLength > 0 And pLength <= TempPlayer(index).buffer.Length - 4
        If pLength <= TempPlayer(index).buffer.Length - 4 Then
            TempPlayer(index).DataPackets = TempPlayer(index).DataPackets + 1
            TempPlayer(index).buffer.ReadLong
            HandleData index, TempPlayer(index).buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If TempPlayer(index).buffer.Length >= 4 Then
            pLength = TempPlayer(index).buffer.ReadLong(False)
            If pLength < 0 Then Exit Sub
        End If
    Loop

    TempPlayer(index).buffer.Trim
End Sub

Private Sub HandleCheckPing(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' Prevent from receiving a empty dat

    Set buffer = New clsBuffer
    buffer.WriteLong SSendPing
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Private Sub HandleNewAccount(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Username As String, Password As String, Email As String
Dim Major As Long, Minor As Long, Revision As Long
Dim CurLanguage As Byte
    
    ' Prevent from receiving a empty data
    

    If IsConnected(index) And Not IsPlaying(index) Then
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        Username = Trim$(buffer.ReadString)
        Password = Trim$(buffer.ReadString)
        Email = Trim$(buffer.ReadString)
        '//Version
        CurLanguage = buffer.ReadByte
        Major = buffer.ReadLong
        Minor = buffer.ReadLong
        Revision = buffer.ReadLong
        TempPlayer(index).ProcessorID = Trim$(buffer.ReadString)
        
        '//ToDo: Checked Banned ProcessorID
        
        '//Version
        If Not Major = App.Major Or Not Minor = App.Minor Or Not Revision = App.Revision Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert index, "Outdated version of client, Please update your client or download the latest version on site", White
                Case LANG_EN: AddAlert index, "Outdated version of client, Please update your client or download the latest version on site", White
                Case LANG_ES: AddAlert index, "Outdated version of client, Please update your client or download the latest version on site", White
            End Select
            Exit Sub
        End If
        
        If Not CheckNameInput(Username, False, (NAME_LENGTH - 1)) Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert index, "Your username must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                Case LANG_EN: AddAlert index, "Your username must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                Case LANG_ES: AddAlert index, "Your username must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
            End Select
            Exit Sub
        End If
        
        If Not CheckNameInput(Password, False, (NAME_LENGTH - 1)) Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert index, "Your password must be between " & ((NAME_LENGTH - 1) / 4) & " and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                Case LANG_EN: AddAlert index, "Your password must be between " & ((NAME_LENGTH - 1) / 4) & " and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                Case LANG_ES: AddAlert index, "Your password must be between " & ((NAME_LENGTH - 1) / 4) & " and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
            End Select
            Exit Sub
        End If
        
        If Not CheckNameInput(Email, False, (TEXT_LENGTH - 1), True) Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert index, "Invalid email", White
                Case LANG_EN: AddAlert index, "Invalid email", White
                Case LANG_ES: AddAlert index, "Invalid email", White
            End Select
            Exit Sub
        End If
        
        If AccountExist(Username) Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert index, "Username is already in used", White
                Case LANG_EN: AddAlert index, "Username is already in used", White
                Case LANG_ES: AddAlert index, "Username is already in used", White
            End Select
            Exit Sub
        End If
        
        AddAccount Username, Password, Email
        Select Case CurLanguage
            Case LANG_PT: AddAlert index, "Account created", White
            Case LANG_EN: AddAlert index, "Account created", White
            Case LANG_ES: AddAlert index, "Account created", White
        End Select
        TextAdd frmServer.txtLog, "Account '" & Username & "' has been created..."
        AddLog "Account '" & Username & "' has been created"
        
        If Not LoadAccount(index, Username) Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert index, "Failed to load account data, Please contact the developer", White
                Case LANG_EN: AddAlert index, "Failed to load account data, Please contact the developer", White
                Case LANG_ES: AddAlert index, "Failed to load account data, Please contact the developer", White
            End Select
            Exit Sub
        End If
        
        TextAdd frmServer.txtLog, "Account '" & Username & "' has logged in..."
        AddIPLog "Account '" & Username & "' has logged in on IP " & GetPlayerIP(index)
        
        '//Update Characters
        'LoadPlayerDatas Index
        SendDataLimit index
        SendCharacters index
        '//Send Connect
        SendLoginOk index
        Set buffer = Nothing
    End If
End Sub

Private Sub HandleLoginInfo(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Username As String, Password As String
Dim Major As Long, Minor As Long, Revision As Long
Dim CurLanguage As Byte
Dim i As Long

    ' Prevent from receiving a empty data
    

    If IsConnected(index) And Not IsPlaying(index) Then
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        Username = Trim$(buffer.ReadString)
        Password = Trim$(buffer.ReadString)
        '//Version
        CurLanguage = buffer.ReadByte
        Major = buffer.ReadLong
        Minor = buffer.ReadLong
        Revision = buffer.ReadLong
        TempPlayer(index).ProcessorID = Trim$(buffer.ReadString)
        
        '//Version
        If Not Major = App.Major Or Not Minor = App.Minor Or Not Revision = App.Revision Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert index, "Outdated version of client, Please update your client or download the latest version on site", White
                Case LANG_EN: AddAlert index, "Outdated version of client, Please update your client or download the latest version on site", White
                Case LANG_ES: AddAlert index, "Outdated version of client, Please update your client or download the latest version on site", White
            End Select
            Exit Sub
        End If
        
        If Not CheckNameInput(Username, False, (NAME_LENGTH - 1)) Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert index, "Your username must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                Case LANG_EN: AddAlert index, "Your username must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                Case LANG_ES: AddAlert index, "Your username must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
            End Select
            Exit Sub
        End If
        
        If Not CheckNameInput(Password, False, (NAME_LENGTH - 1)) Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert index, "Your password must be between " & ((NAME_LENGTH - 1) / 4) & " and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                Case LANG_EN: AddAlert index, "Your password must be between " & ((NAME_LENGTH - 1) / 4) & " and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                Case LANG_ES: AddAlert index, "Your password must be between " & ((NAME_LENGTH - 1) / 4) & " and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
            End Select
            Exit Sub
        End If
        
        If Not AccountExist(Username) Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert index, "Account does not exist", White
                Case LANG_EN: AddAlert index, "Account does not exist", White
                Case LANG_ES: AddAlert index, "Account does not exist", White
            End Select
            Exit Sub
        End If
        
        If Not isPasswordOK(Username, Password) Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert index, "Invalid password", White
                Case LANG_EN: AddAlert index, "Invalid password", White
                Case LANG_ES: AddAlert index, "Invalid password", White
            End Select
            Exit Sub
        End If
        
        i = FindAccount(Username)
        If i > 0 Then
            '//Check User IP
            If Len(Trim$(GetPlayerIP(i))) > 0 Then
                If Trim$(GetPlayerIP(i)) = GetPlayerIP(index) Then
                    '//Disconnect
                    CloseSocket i
                Else
                    Select Case CurLanguage
                        Case LANG_PT: AddAlert index, "Account is currently connected", White, YES
                        Case LANG_EN: AddAlert index, "Account is currently connected", White, YES
                        Case LANG_ES: AddAlert index, "Account is currently connected", White, YES
                    End Select
                    Exit Sub
                End If
            Else
                '//Disconnect
                CloseSocket i
            End If
        End If
        
        If Not LoadAccount(index, Username) Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert index, "Failed to load account data, Please contact the developer", White, YES
                Case LANG_EN: AddAlert index, "Failed to load account data, Please contact the developer", White, YES
                Case LANG_ES: AddAlert index, "Failed to load account data, Please contact the developer", White, YES
            End Select
            Exit Sub
        End If
        
        TextAdd frmServer.txtLog, "Account '" & Username & "' has logged in..."
        AddIPLog "Account '" & Username & "' has logged in on IP " & GetPlayerIP(index)
        
        SendDataLimit index
        '//Update Characters
        SendCharacters index
        '//Send connect
        SendLoginOk index
        Set buffer = Nothing
    End If
End Sub

Private Sub HandleNewCharacter(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim CharName As String, Gender As Byte, CharSelected As Byte
Dim Sprite As Long
Dim Major As Long, Minor As Long, Revision As Long
Dim CurLanguage As Byte

    ' Prevent from receiving a empty data
    

    If IsConnected(index) And Not IsPlaying(index) Then
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        CharName = Trim$(buffer.ReadString)
        Gender = buffer.ReadByte
        CharSelected = buffer.ReadByte
        '//Version
        CurLanguage = buffer.ReadByte
        Major = buffer.ReadLong
        Minor = buffer.ReadLong
        Revision = buffer.ReadLong
        TempPlayer(index).ProcessorID = Trim$(buffer.ReadString)
        
        '//Version
        If Not Major = App.Major Or Not Minor = App.Minor Or Not Revision = App.Revision Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert index, "Outdated version of client, Please update your client or download the latest version on site", White
                Case LANG_EN: AddAlert index, "Outdated version of client, Please update your client or download the latest version on site", White
                Case LANG_ES: AddAlert index, "Outdated version of client, Please update your client or download the latest version on site", White
            End Select
            Exit Sub
        End If
        
        If Not CheckNameInput(CharName, False, (NAME_LENGTH - 1)) Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert index, "Your character name must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                Case LANG_EN: AddAlert index, "Your character name must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                Case LANG_ES: AddAlert index, "Your character name must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
            End Select
            Exit Sub
        End If
        
        '//Check if CharName exist
        If CheckSameName(CharName) Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert index, "Character name already exist", White
                Case LANG_EN: AddAlert index, "Character name already exist", White
                Case LANG_ES: AddAlert index, "Character name already exist", White
            End Select
            Exit Sub
        End If
        
        '//Set specific gender sprite
        If Gender = GENDER_MALE Then
            Sprite = 1
        ElseIf Gender = GENDER_FEMALE Then
            Sprite = 2
        End If
        
        '//Add Character
        AddPlayerData index, CharSelected, CharName, Sprite
        '//Load data
        LoadPlayerData index, CharSelected
        LoadPlayerInv index, CharSelected
        LoadPlayerPokemons index, CharSelected
        LoadPlayerInvStorage index, CharSelected
        LoadPlayerPokemonStorage index, CharSelected
        LoadPlayerPokedex index, CharSelected
        
        TextAdd frmServer.txtLog, "Character '" & CharName & "' has been created..."
        AddLog "Character '" & CharName & "' has been created from Account '" & Trim$(Account(index).Username)
        
        Select Case CurLanguage
            Case LANG_PT: AddAlert index, "New Character Created", White
            Case LANG_EN: AddAlert index, "New Character Created", White
            Case LANG_ES: AddAlert index, "New Character Created", White
        End Select
        
        '//Update Characters
        SendCharacters index
        '//Send connection again
        SendLoginOk index, YES
        Set buffer = Nothing
    End If
End Sub

Private Sub HandleUseCharacter(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim CharSelected As Byte
Dim Major As Long, Minor As Long, Revision As Long
Dim CurLanguage As Byte

    ' Prevent from receiving a empty data
    

    If IsConnected(index) And Not IsPlaying(index) Then
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        CharSelected = buffer.ReadByte
        '//Version
        CurLanguage = buffer.ReadByte
        Major = buffer.ReadLong
        Minor = buffer.ReadLong
        Revision = buffer.ReadLong
        TempPlayer(index).ProcessorID = Trim$(buffer.ReadString)
        
        '//Version
        If Not Major = App.Major Or Not Minor = App.Minor Or Not Revision = App.Revision Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert index, "Outdated version of client, Please update your client or download the latest version on site", White
                Case LANG_EN: AddAlert index, "Outdated version of client, Please update your client or download the latest version on site", White
                Case LANG_ES: AddAlert index, "Outdated version of client, Please update your client or download the latest version on site", White
            End Select
            Exit Sub
        End If
        
        '//Make sure player have character on the selected slot
        If CharSelected > 0 And CharSelected <= MAX_PLAYERCHAR Then
            '//Check if in used
            If Len(Player(index, CharSelected).Name) <= 0 Then
                Select Case CurLanguage
                    Case LANG_PT: AddAlert index, "Character is empty! based on the server, please try to login again. If it didn't work, try to contact the developers", White
                    Case LANG_EN: AddAlert index, "Character is empty! based on the server, please try to login again. If it didn't work, try to contact the developers", White
                    Case LANG_ES: AddAlert index, "Character is empty! based on the server, please try to login again. If it didn't work, try to contact the developers", White
                End Select
                Exit Sub
            End If
            
            '//Make sure it's not already playing
            If TempPlayer(index).UseChar > 0 Then
                LeftGame index
                Select Case CurLanguage
                    Case LANG_PT: AddAlert index, "Your character was disconnected, Please try again", White, YES
                    Case LANG_EN: AddAlert index, "Your character was disconnected, Please try again", White, YES
                    Case LANG_ES: AddAlert index, "Your character was disconnected, Please try again", White, YES
                End Select
                Exit Sub
            End If
            
            '//Make sure it's not already playing
            If FindPlayer(Trim$(Player(index, CharSelected).Name)) > 0 Then
                Select Case CurLanguage
                    Case LANG_PT: AddAlert index, "Character is currently connnected", White, YES
                    Case LANG_EN: AddAlert index, "Character is currently connnected", White, YES
                    Case LANG_ES: AddAlert index, "Character is currently connnected", White, YES
                End Select
                Exit Sub
            End If
            
            If IsCharacterBanned(Trim$(Player(index, CharSelected).Name)) Then
                Select Case CurLanguage
                    Case LANG_PT: AddAlert index, "This character is banned", White, YES
                    Case LANG_EN: AddAlert index, "This character is banned", White, YES
                    Case LANG_ES: AddAlert index, "This character is banned", White, YES
                End Select
                Exit Sub
            End If
            
            
            '//Set use char
            TempPlayer(index).UseChar = CharSelected
            JoinGame index, CurLanguage
        Else
            Select Case CurLanguage
                Case LANG_PT: AddAlert index, "Invalid Character Slot", White
                Case LANG_EN: AddAlert index, "Invalid Character Slot", White
                Case LANG_ES: AddAlert index, "Invalid Character Slot", White
            End Select
            Exit Sub
        End If
        Set buffer = Nothing
    End If
End Sub

Private Sub HandleDelCharacter(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim CharSelected As Byte
Dim Major As Long, Minor As Long, Revision As Long
Dim CurLanguage As Byte

    ' Prevent from receiving a empty data
    

    If IsConnected(index) And Not IsPlaying(index) Then
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        CharSelected = buffer.ReadByte
        '//Version
        CurLanguage = buffer.ReadByte
        Major = buffer.ReadLong
        Minor = buffer.ReadLong
        Revision = buffer.ReadLong
        TempPlayer(index).ProcessorID = Trim$(buffer.ReadString)
        
        '//Version
        If Not Major = App.Major Or Not Minor = App.Minor Or Not Revision = App.Revision Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert index, "Outdated version of client, Please update your client or download the latest version on site", White
                Case LANG_EN: AddAlert index, "Outdated version of client, Please update your client or download the latest version on site", White
                Case LANG_ES: AddAlert index, "Outdated version of client, Please update your client or download the latest version on site", White
            End Select
            Exit Sub
        End If
        
        '//Make sure player have character on the selected slot
        If CharSelected > 0 And CharSelected <= MAX_PLAYERCHAR Then
            '//Check if in used
            If Len(Player(index, CharSelected).Name) <= 0 Then
                Select Case CurLanguage
                    Case LANG_PT: AddAlert index, "Character is empty! based on the server, please try to login again. If it didn't work, try to contact the developers", White
                    Case LANG_EN: AddAlert index, "Character is empty! based on the server, please try to login again. If it didn't work, try to contact the developers", White
                    Case LANG_ES: AddAlert index, "Character is empty! based on the server, please try to login again. If it didn't work, try to contact the developers", White
                End Select
                Exit Sub
            End If
            
            '//Delete char
            DeletePlayerData index, CharSelected
            SendCharacters index
            Select Case CurLanguage
                Case LANG_PT: AddAlert index, "Character data deleted!", White
                Case LANG_EN: AddAlert index, "Character data deleted!", White
                Case LANG_ES: AddAlert index, "Character data deleted!", White
            End Select
        Else
            Select Case CurLanguage
                Case LANG_PT: AddAlert index, "Invalid Character Slot", White
                Case LANG_EN: AddAlert index, "Invalid Character Slot", White
                Case LANG_ES: AddAlert index, "Invalid Character Slot", White
            End Select
            Exit Sub
        End If
        Set buffer = Nothing
    End If
End Sub

Private Sub HandleNeedMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim NeedMap As Byte
Dim MapNum As Long
Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    NeedMap = buffer.ReadByte
    Set buffer = Nothing
    
    If NeedMap = YES Then
        SendMap index, Player(index, TempPlayer(index).UseChar).Map
    End If
    
    '//Send Map Data
    SendMapNpcData Player(index, TempPlayer(index).UseChar).Map, index
    For i = 1 To MAX_MAP_NPC
        SendNpcPokemonData Player(index, TempPlayer(index).UseChar).Map, i, NO, 0, 0, 0, index
    Next
    
    '//Done loading
    SendJoinMap index
    
    TempPlayer(index).GettingMap = False
    
    MapNum = Player(index, TempPlayer(index).UseChar).Map
    If MapNum > 0 Then
        ChangeTempSprite index, Map(MapNum).SpriteType
    End If
    
    '//Send Weather
    SendWeatherTo index, MapNum
    
    '//Done Loading
    SendMapDone index
End Sub

Private Sub HandlePlayerMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Dir As Byte
Dim tmpX As Long, tmpY As Long
Dim RndNum As Long

    ' Prevent from receiving a empty data
    

    If Not IsPlaying(index) Then Exit Sub

    '//Check if can move
    If TempPlayer(index).GettingMap Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Dir = buffer.ReadByte
    tmpX = buffer.ReadLong
    tmpY = buffer.ReadLong
    Set buffer = Nothing

    '//Prevent hacking
    If Dir < 0 Or Dir > DIR_RIGHT Then Exit Sub

    '//ToFast
    If Player(index, TempPlayer(index).UseChar).MoveTmr > GetTickCount Then
        '//ToDo: Create something to prevent hacking
        'SendPlayerXY index, True
        'Exit Sub
    End If
    
    If TempPlayer(index).StorageType > 0 Then
        SendPlayerXY index, True
        Exit Sub
    End If
    If TempPlayer(index).InShop > 0 Then
        SendPlayerXY index, True
        Exit Sub
    End If
    If TempPlayer(index).PlayerRequest > 0 Then
        SendPlayerXY index, True
        Exit Sub
    End If
    If TempPlayer(index).InDuel > 0 Then
        SendPlayerXY index, True
        Exit Sub
    End If
    If TempPlayer(index).InTrade > 0 Then
        SendPlayerXY index, True
        Exit Sub
    End If
    If TempPlayer(index).CurConvoNum > 0 Then
        SendPlayerXY index, True
        Exit Sub
    End If
    If TempPlayer(index).InNpcDuel > 0 Then
        SendPlayerXY index, True
        Exit Sub
    End If
    If Player(index, TempPlayer(index).UseChar).Action > 0 Then
        SendPlayerXY index, True
        Exit Sub
    End If
    
    '//Desynced
    If Not Player(index, TempPlayer(index).UseChar).x = tmpX Then
        SendPlayerXY index, True
        Exit Sub
    End If
    If Not Player(index, TempPlayer(index).UseChar).y = tmpY Then
        SendPlayerXY index, True
        Exit Sub
    End If
    
    If Player(index, TempPlayer(index).UseChar).IsConfuse = YES Then
        'Dir = Random(0, 3)
        'If Dir < 0 Then Dir = 0
        'If Dir > DIR_RIGHT Then Dir = DIR_RIGHT
        RndNum = Random(1, 10)
        If RndNum = 1 Then
            Player(index, TempPlayer(index).UseChar).IsConfuse = 0
            Select Case TempPlayer(index).CurLanguage
                Case LANG_PT: AddAlert index, "You snapped out of confusion", White
                Case LANG_EN: AddAlert index, "You snapped out of confusion", White
                Case LANG_ES: AddAlert index, "You snapped out of confusion", White
            End Select
            SendPlayerStatus index
        End If
    End If

    Player(index, TempPlayer(index).UseChar).MoveTmr = GetTickCount + 200
    Call PlayerMove(index, Dir)
End Sub

Private Sub HandlePlayerDir(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Dir As Byte
    
    ' Prevent from receiving a empty data
    

    If Not IsPlaying(index) Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Dir = buffer.ReadByte
    Set buffer = Nothing
    
    '//Prevent hacking
    If Dir < 0 Or Dir > DIR_RIGHT Then Exit Sub
    
    Player(index, TempPlayer(index).UseChar).Dir = Dir
    
    SendPlayerDir index, True
End Sub

Private Sub HandleMapMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MapNum As Long, i As Long
Dim Msg As String, MsgColor As Long
Dim UpdateMsg As String
    
    ' Prevent from receiving a empty data
    

    If Not IsPlaying(index) Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    Set buffer = Nothing
    
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If Player(index, TempPlayer(index).UseChar).Muted = YES Then Exit Sub
    
    With Player(index, TempPlayer(index).UseChar)
        MapNum = .Map
        
        Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: UpdateMsg = "[Mapa]" & Trim$(.Name) & ": " & Msg
            Case LANG_EN: UpdateMsg = "[Map]" & Trim$(.Name) & ": " & Msg
            Case LANG_ES: UpdateMsg = "[Mapa]" & Trim$(.Name) & ": " & Msg
        End Select
                
        MsgColor = White
    End With
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If TempPlayer(i).UseChar > 0 Then
                If Player(i, TempPlayer(i).UseChar).Map = MapNum Then
                    '//Send Msg
                    SendChatbubble MapNum, index, TARGET_TYPE_PLAYER, Msg, White
                    SendPlayerMsg i, UpdateMsg, MsgColor
                End If
            End If
        End If
    Next
    
    AddLog UpdateMsg
End Sub

Private Sub HandleGlobalMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long
Dim Msg As String, MsgColor As Long
Dim UpdateMsg As String

    ' Prevent from receiving a empty data
    

    If Not IsPlaying(index) Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    Set buffer = Nothing
    
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If Player(index, TempPlayer(index).UseChar).Muted = YES Then Exit Sub
    
    With Player(index, TempPlayer(index).UseChar)
        Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: UpdateMsg = "[Todos] " & Trim$(.Name) & ": " & Msg
            Case LANG_EN: UpdateMsg = "[All] " & Trim$(.Name) & ": " & Msg
            Case LANG_ES: UpdateMsg = "[Todos] " & Trim$(.Name) & ": " & Msg
        End Select
        MsgColor = Yellow
    End With
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If TempPlayer(i).UseChar > 0 Then
                If TempPlayer(i).CurLanguage = TempPlayer(index).CurLanguage Then
                    '//Send Msg
                    SendPlayerMsg i, UpdateMsg, MsgColor
                End If
            End If
        End If
    Next
    
    AddLog UpdateMsg
End Sub

Private Sub HandlePlayerMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim InputName As String, i As Long
Dim Msg As String, MsgColor As Long
Dim UpdateMsg As String

    If Not IsPlaying(index) Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    InputName = buffer.ReadString
    Msg = buffer.ReadString
    Set buffer = Nothing
    
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If Player(index, TempPlayer(index).UseChar).Muted = YES Then Exit Sub
    
    i = FindPlayer(InputName)
    
    If i > 0 Then
        If Not i = index Then
            If IsPlaying(i) Then
                If TempPlayer(i).UseChar > 0 Then
                    With Player(index, TempPlayer(index).UseChar)
                        UpdateMsg = "[" & Trim$(Player(i, TempPlayer(i).UseChar).Name) & "] " & Trim$(.Name) & ": " & Msg
                        MsgColor = BrightCyan
                    End With
    
                    '//Send Msg
                    SendPlayerMsg i, UpdateMsg, MsgColor
                    SendPlayerMsg index, UpdateMsg, MsgColor
                    
                    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " whisper " & Trim$(Player(i, TempPlayer(i).UseChar).Name) & " [" & Msg & "]"
                Else
                    SendPlayerMsg index, "Player is not online!", BrightRed
                End If
            Else
                SendPlayerMsg index, "Player is not online!", BrightRed
            End If
        Else
            SendPlayerMsg index, "You cannot message yourself!", BrightRed
        End If
    Else
        SendPlayerMsg index, "Player is not online!", BrightRed
    End If
End Sub

Private Sub HandleWarpTo(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MapNum As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    With Player(index, TempPlayer(index).UseChar)
        If .Access < ACCESS_MODERATOR Then Exit Sub
    
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        MapNum = buffer.ReadLong
        Set buffer = Nothing
        
        If MapNum <= 0 Or MapNum > MAX_MAP Then Exit Sub
        
        TextAdd frmServer.txtLog, Trim$(.Name) & " warped to map#" & MapNum
        AddLog Trim$(.Name) & " warped to map#" & MapNum
        
        PlayerWarp index, MapNum, .x, .y, .Dir
    End With
End Sub

Private Sub HandleAdminWarp(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim x As Long, y As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    With Player(index, TempPlayer(index).UseChar)
        If .Access < ACCESS_MODERATOR Then Exit Sub
    
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        x = buffer.ReadLong
        y = buffer.ReadByte
        Set buffer = Nothing
        
        If x < 0 Then x = 0
        If y < 0 Then y = 0
        If x > Map(.Map).MaxX Then x = Map(.Map).MaxX
        If y > Map(.Map).MaxY Then y = Map(.Map).MaxY
        
        '//Set
        .x = x
        .y = y
        
        SendPlayerXY index, True
    End With
End Sub

Private Sub HandleWarpToMe(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim InputName As String, i As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    With Player(index, TempPlayer(index).UseChar)
        If .Access < ACCESS_MODERATOR Then Exit Sub
        
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        InputName = buffer.ReadString
        Set buffer = Nothing
    
        i = FindPlayer(InputName)
    
        If i > 0 Then
            If Not i = index Then
                If IsPlaying(i) Then
                    If TempPlayer(i).UseChar > 0 Then
                        PlayerWarp i, .Map, .x, .y, .Dir
                        
                        AddLog Trim$(Player(i, TempPlayer(i).UseChar).Name) & " warped to [" & .Map & " | " & .x & " | " & .y & "]"
                    Else
                        SendPlayerMsg index, "Player is not online!", BrightRed
                    End If
                Else
                    SendPlayerMsg index, "Player is not online!", BrightRed
                End If
            Else
                SendPlayerMsg index, "You cannot warp yourself to yourself!", BrightRed
            End If
        Else
            SendPlayerMsg index, "Player is not online!", BrightRed
        End If
    End With
End Sub

Private Sub HandleWarpMeTo(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim InputName As String, i As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    InputName = buffer.ReadString
    Set buffer = Nothing
    
    i = FindPlayer(InputName)
        
    If i > 0 Then
        If Not i = index Then
            If IsPlaying(i) Then
                If TempPlayer(i).UseChar > 0 Then
                    With Player(i, TempPlayer(i).UseChar)
                        PlayerWarp index, .Map, .x, .y, .Dir
                        
                        AddLog Trim$(.Name) & " warped to [" & .Map & " | " & .x & " | " & .y & "]"
                    End With
                Else
                    SendPlayerMsg index, "Player is not online!", BrightRed
                End If
            Else
                SendPlayerMsg index, "Player is not online!", BrightRed
            End If
        Else
            SendPlayerMsg index, "You cannot warp yourself to yourself!", BrightRed
        End If
    Else
        SendPlayerMsg index, "Player is not online!", BrightRed
    End If
End Sub

Private Sub HandleSetAccess(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim InputName As String, i As Long
Dim Access As Byte, OldAccess As Byte

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_CREATOR Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    InputName = buffer.ReadString
    Access = buffer.ReadByte
    Set buffer = Nothing
    
    i = FindPlayer(InputName)
        
    If i > 0 Then
        If Not i = index Then
            If IsPlaying(i) Then
                If TempPlayer(i).UseChar > 0 Then
                    With Player(i, TempPlayer(i).UseChar)
                        If .Access >= Player(index, TempPlayer(index).UseChar).Access Then
                            SendPlayerMsg index, "Player's access is greater than your own!", BrightRed
                        Else
                            OldAccess = .Access
                            .Access = Access
                            SendPlayerData i
                            
                            AddLog Trim$(.Name) & " got his access changed by " & Trim$(Player(index, TempPlayer(index).UseChar).Name) & " from " & OldAccess & " to " & Access
                        End If
                    End With
                Else
                    SendPlayerMsg index, "Player is not online!", BrightRed
                End If
            Else
                SendPlayerMsg index, "Player is not online!", BrightRed
            End If
        Else
            SendPlayerMsg index, "You cannot change your own access!", BrightRed
        End If
    Else
        SendPlayerMsg index, "Player is not online!", BrightRed
    End If
End Sub

Private Sub HandlePlayerPokemonMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Dir As Byte
Dim tmpX As Long, tmpY As Long
Dim DuelIndex As Long
Dim RndNum As Byte

    If Not IsPlaying(index) Then Exit Sub

    '//Check if can move
    If TempPlayer(index).GettingMap Then Exit Sub
    If PlayerPokemon(index).Num <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Dir = buffer.ReadByte
    tmpX = buffer.ReadLong
    tmpY = buffer.ReadLong
    Set buffer = Nothing

    '//Prevent hacking
    If Dir < 0 Or Dir > DIR_RIGHT Then Exit Sub

    '//ToFast
    If PlayerPokemon(index).MoveTmr > GetTickCount Then
        '//ToDo: Create something to prevent hacking
        'SendPlayerPokemonXY index, True
        'Exit Sub
    End If
    'If PlayerPokemon(Index).QueueMove > 0 Then
    '    SendPlayerPokemonXY Index, True
    '    Exit Sub
    'End If
    If PlayerPokemons(index).Data(PlayerPokemon(index).slot).Status = StatusEnum.Sleep Then
        SendPlayerPokemonXY index, True
        Exit Sub
    End If
    If PlayerPokemons(index).Data(PlayerPokemon(index).slot).Status = StatusEnum.Frozen Then
        SendPlayerPokemonXY index, True
        Exit Sub
    End If
    
    '//Desynced
    If Not PlayerPokemon(index).x = tmpX Then
        SendPlayerPokemonXY index, True
        Exit Sub
    End If
    If Not PlayerPokemon(index).y = tmpY Then
        SendPlayerPokemonXY index, True
        Exit Sub
    End If
    
    '//Status
    If PlayerPokemons(index).Data(PlayerPokemon(index).slot).Status = StatusEnum.Poison Then
        If PlayerPokemon(index).StatusMove >= 4 Then
            If PlayerPokemon(index).StatusDamage > 0 Then
                If PlayerPokemon(index).StatusDamage >= PlayerPokemons(index).Data(PlayerPokemon(index).slot).CurHP Then
                    '//Dead
                    PlayerPokemons(index).Data(PlayerPokemon(index).slot).CurHP = 0
                    SendActionMsg Player(index, TempPlayer(index).UseChar).Map, "-" & PlayerPokemon(index).StatusDamage, PlayerPokemon(index).x * 32, PlayerPokemon(index).y * 32, Magenta
                    SendPlayerPokemonVital index
                    SendPlayerPokemonFaint index
                    Exit Sub
                Else
                    '//Reduce
                    PlayerPokemons(index).Data(PlayerPokemon(index).slot).CurHP = PlayerPokemons(index).Data(PlayerPokemon(index).slot).CurHP - PlayerPokemon(index).StatusDamage
                    SendActionMsg Player(index, TempPlayer(index).UseChar).Map, "-" & PlayerPokemon(index).StatusDamage, PlayerPokemon(index).x * 32, PlayerPokemon(index).y * 32, Magenta
                    '//Update
                    SendPlayerPokemonVital index
                End If
                '//Reset
                PlayerPokemon(index).StatusMove = 0
            Else
                PlayerPokemon(index).StatusDamage = (PlayerPokemons(index).Data(PlayerPokemon(index).slot).MaxHP / 16)
            End If
        Else
            PlayerPokemon(index).StatusMove = PlayerPokemon(index).StatusMove + 1
        End If
        '//ToDo: Check if Badly Poisoned
    End If
    
    If PlayerPokemon(index).IsConfuse = YES Then
        'Dir = Random(0, 3)
        'If Dir < 0 Then Dir = 0
        'If Dir > DIR_RIGHT Then Dir = DIR_RIGHT
        RndNum = Random(1, 10)
        If RndNum = 1 Then
            PlayerPokemon(index).IsConfuse = 0
            If PlayerPokemon(index).Num > 0 Then
                Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, Trim$(Pokemon(PlayerPokemon(index).Num).Name) & " snapped out of confusion", White
                    Case LANG_EN: AddAlert index, Trim$(Pokemon(PlayerPokemon(index).Num).Name) & " snapped out of confusion", White
                    Case LANG_ES: AddAlert index, Trim$(Pokemon(PlayerPokemon(index).Num).Name) & " snapped out of confusion", White
                End Select
                SendPlayerPokemonStatus index
            End If
        End If
    End If
    
    PlayerPokemon(index).MoveTmr = GetTickCount + 200
    Call PlayerPokemonMove(index, Dir)
End Sub

Private Sub HandlePlayerPokemonDir(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Dir As Byte

    If Not IsPlaying(index) Then Exit Sub
    If PlayerPokemon(index).Num <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Dir = buffer.ReadByte
    Set buffer = Nothing
    
    '//Prevent hacking
    If Dir < 0 Or Dir > DIR_RIGHT Then Exit Sub
    
    PlayerPokemon(index).Dir = Dir
    
    SendPlayerDir index, True
End Sub

Private Sub HandleGetItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim ItemNum As Long, ItemVal As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ItemNum = buffer.ReadLong
    ItemVal = buffer.ReadLong
    Set buffer = Nothing
    
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " , Admin Rights: Get Item#" & ItemNum & " x" & ItemVal
    TryGivePlayerItem index, ItemNum, ItemVal
End Sub

Private Sub HandlePlayerPokemonState(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim State As Byte
Dim PokeSlot As Byte

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    State = buffer.ReadByte
    PokeSlot = buffer.ReadByte
    Set buffer = Nothing
    
    If Player(index, TempPlayer(index).UseChar).Action > 0 Then Exit Sub
    
    '//Check if exist
    If PokeSlot > 0 And PokeSlot <= MAX_PLAYER_POKEMON Then
        If State = 1 Then
            If PlayerPokemon(index).Num > 0 Then Exit Sub
            
            If PlayerPokemons(index).Data(PokeSlot).Num > 0 Then
                '//Make sure it's still alive
                If PlayerPokemons(index).Data(PokeSlot).CurHP > 0 Then
                    If TempPlayer(index).DuelReset = YES Then
                        '//Check if in duel, reset timer
                        If TempPlayer(index).InDuel > 0 Then
                            If IsPlaying(TempPlayer(index).InDuel) Then
                                If TempPlayer(TempPlayer(index).InDuel).UseChar > 0 Then
                                    TempPlayer(index).DuelTime = 3
                                    TempPlayer(index).DuelTimeTmr = GetTickCount + 1000
                                    TempPlayer(TempPlayer(index).InDuel).DuelTime = 3
                                    TempPlayer(TempPlayer(index).InDuel).DuelTimeTmr = GetTickCount + 1000
                                End If
                            End If
                        End If
                        If TempPlayer(index).InNpcDuel > 0 Then
                            TempPlayer(index).DuelTime = 3
                            TempPlayer(index).DuelTimeTmr = GetTickCount + 1000
                        End If
                        TempPlayer(index).DuelReset = NO
                    End If
                    If PlayerPokemons(index).Data(PokeSlot).Level <= (Player(index, TempPlayer(index).UseChar).Level + 10) Then
                        SpawnPlayerPokemon index, PokeSlot
                    Else
                        AddAlert index, "Not enough level", White
                    End If
                End If
            End If
        Else
            If PlayerPokemon(index).Num = 0 Then Exit Sub
            
            ClearPlayerPokemon index
        End If
    End If
End Sub

Private Sub HandleAttack(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MoveSlot As Byte
Dim movesetNum As Byte

    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If Not IsPlaying(index) Then Exit Sub
    
    If PlayerPokemon(index).Num > 0 Then
        If PlayerPokemon(index).slot <= 0 Then Exit Sub
        
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        MoveSlot = buffer.ReadByte
        Set buffer = Nothing
        
        '//ToFast
        If PlayerPokemon(index).AtkTmr > GetTickCount Then
            Exit Sub
        End If
            
        '//Exit out if there's a Queue Move
        If PlayerPokemon(index).QueueMove > 0 Then
            Exit Sub
        End If
        
        '//Exit out if in duel timer
        If TempPlayer(index).DuelTime > 0 Then
            Exit Sub
        End If
        
        '//Select Moveslot
        If MoveSlot > 0 Then
            '//Check if moveslot is not empty
            If PlayerPokemons(index).Data(PlayerPokemon(index).slot).Moveset(MoveSlot).Num > 0 Then
                PlayerCastMove index, PlayerPokemons(index).Data(PlayerPokemon(index).slot).Moveset(MoveSlot).Num, MoveSlot, True
            Else
                '//Use Struggle
                PlayerCastMove index, 1, 0, False
            End If
        Else
            '//Use Struggle
            PlayerCastMove index, 1, 0, False
        End If
        
        PlayerPokemon(index).AtkTmr = GetTickCount + 500
        SendAttack index, Player(index, TempPlayer(index).UseChar).Map
    End If
End Sub

Private Sub HandleChangePassword(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim newPass As String, oldPass As String
Dim filename As String

    If IsConnected(index) And Not IsPlaying(index) Then
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        newPass = Trim$(buffer.ReadString)
        oldPass = Trim$(buffer.ReadString)
        
        '//Make sure it's connected
        If Len(Account(index).Username) > 0 Then
            '//Check if oldpassword match
            If Not Trim$(Account(index).Password) = oldPass Then
                Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "Incorrect password", White
                    Case LANG_EN: AddAlert index, "Incorrect password", White
                    Case LANG_ES: AddAlert index, "Incorrect password", White
                End Select
                Exit Sub
            End If
            
            '//Save
            filename = App.Path & "\data\accounts\" & Trim$(Account(index).Username) & "\account.ini"
            
            '//Check if file exist
            If Not FileExist(filename) Then
                Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "Error loading account", White
                    Case LANG_EN: AddAlert index, "Error loading account", White
                    Case LANG_ES: AddAlert index, "Error loading account", White
                End Select
                Exit Sub
            End If
            
            Call PutVar(filename, "Account", "Password", newPass)
            Select Case TempPlayer(index).CurLanguage
                Case LANG_PT: AddAlert index, "Password successfully changed", White
                Case LANG_EN: AddAlert index, "Password successfully changed", White
                Case LANG_ES: AddAlert index, "Password successfully changed", White
            End Select
            
            TextAdd frmServer.txtLog, "Account: " & Trim$(Account(index).Username) & " changed it's password"
            AddLog "Account: " & Trim$(Account(index).Username) & " changed it's password"
        End If
        Set buffer = Nothing
    End If
End Sub

Private Sub HandleReplaceNewMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MoveSlot As Byte
Dim MoveNum As Long
Dim OldMove As Long

    '//Check Error
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If TempPlayer(index).MoveLearnPokeSlot <= 0 Or TempPlayer(index).MoveLearnPokeSlot > MAX_PLAYER_POKEMON Then Exit Sub
    If PlayerPokemons(index).Data(TempPlayer(index).MoveLearnPokeSlot).Num <= 0 Then Exit Sub
    If TempPlayer(index).MoveLearnNum <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MoveSlot = buffer.ReadByte
    Set buffer = Nothing
    
    '//Check
    If MoveSlot > 0 And MoveSlot <= MAX_MOVESET Then
        MoveNum = TempPlayer(index).MoveLearnNum
        '//Replace
        With PlayerPokemons(index).Data(TempPlayer(index).MoveLearnPokeSlot)
            OldMove = .Moveset(MoveSlot).Num
            .Moveset(MoveSlot).Num = MoveNum
            .Moveset(MoveSlot).TotalPP = PokemonMove(MoveNum).PP
            .Moveset(MoveSlot).CurPP = .Moveset(MoveSlot).TotalPP
            .Moveset(MoveSlot).CD = 0
            SendPlayerPokemonSlot index, TempPlayer(index).MoveLearnPokeSlot
            '//Send Msg
            AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & "'s pokemon, " & Trim$(Pokemon(.Num).Name) & ", learned the move " & Trim$(PokemonMove(MoveNum).Name) & " replacing " & Trim$(PokemonMove(OldMove).Name)
            SendPlayerMsg index, Trim$(Pokemon(.Num).Name) & " learned the move " & Trim$(PokemonMove(MoveNum).Name), White
        End With
        
        TempPlayer(index).MoveLearnPokeSlot = 0
        TempPlayer(index).MoveLearnNum = 0
        TempPlayer(index).MoveLearnIndex = 0
        
        '//Continue Checking New Move
        CheckNewMove index, TempPlayer(index).MoveLearnPokeSlot, TempPlayer(index).MoveLearnIndex
    Else
        TempPlayer(index).MoveLearnPokeSlot = 0
        TempPlayer(index).MoveLearnNum = 0
        TempPlayer(index).MoveLearnIndex = 0
    End If
End Sub

Private Sub HandleEvolvePoke(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim EvolveSlot As Byte
Dim PokeNum As Long, EvolveNum As Long
Dim statNu As Byte
Dim itemSlot As Byte

    '//Check Error
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If PlayerPokemon(index).Num <= 0 Then Exit Sub
    If PlayerPokemon(index).slot <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    EvolveSlot = buffer.ReadByte
    Set buffer = Nothing
    
    '//Check
    If EvolveSlot > 0 And EvolveSlot <= MAX_EVOLVE Then
        PokeNum = PlayerPokemon(index).Num
        EvolveNum = Pokemon(PokeNum).EvolveNum(EvolveSlot)
        If EvolveNum > 0 Then
            '//Check Condition
            With PlayerPokemons(index).Data(PlayerPokemon(index).slot)
                If .Level < Pokemon(.Num).EvolveLevel(EvolveSlot) Then
                    Select Case TempPlayer(index).CurLanguage
                        Case LANG_PT: AddAlert index, "Your pokemon is not yet ready to evolve", White
                        Case LANG_EN: AddAlert index, "Your pokemon is not yet ready to evolve", White
                        Case LANG_ES: AddAlert index, "Your pokemon is not yet ready to evolve", White
                    End Select
                    Exit Sub
                End If
                
                Select Case Pokemon(.Num).EvolveCondition(EvolveSlot)
                    Case EVOLVE_CONDT_ITEM
                        If Pokemon(.Num).EvolveConditionData(EvolveSlot) > 0 Then
                            itemSlot = checkItem(index, Pokemon(.Num).EvolveConditionData(EvolveSlot))
                            If itemSlot > 0 Then
                                '//Take Item
                                PlayerInv(index).Data(itemSlot).Value = PlayerInv(index).Data(itemSlot).Value - 1
                                If PlayerInv(index).Data(itemSlot).Value <= 0 Then
                                    '//Clear Item
                                    PlayerInv(index).Data(itemSlot).Num = 0
                                    PlayerInv(index).Data(itemSlot).Value = 0
                                End If
                                SendPlayerInvSlot index, itemSlot
                            Else
                                Select Case TempPlayer(index).CurLanguage
                                    Case LANG_PT: AddAlert index, "Your pokemon is not yet ready to evolve", White
                                    Case LANG_EN: AddAlert index, "Your pokemon is not yet ready to evolve", White
                                    Case LANG_ES: AddAlert index, "Your pokemon is not yet ready to evolve", White
                                End Select
                                Exit Sub
                            End If
                        End If
                End Select
                
                '//Change
                .Num = EvolveNum
                '//Calculate new stat
                For statNu = 1 To StatEnum.Stat_Count - 1
                    .Stat(statNu).Value = CalculatePokemonStat(statNu, .Num, .Level, .Stat(statNu).EV, .Stat(statNu).IV, .Nature)
                Next
                
                .MaxHP = .Stat(StatEnum.HP).Value
                
                '//Send Animation
                SendPlayAnimation Player(index, TempPlayer(index).UseChar).Map, 76, PlayerPokemon(index).x, PlayerPokemon(index).y ' ToDo: Change to 76
                
                '//Update Map Poke
                PlayerPokemon(index).Num = .Num
                
                '//Send Update
                SendPlayerPokemonSlot index, PlayerPokemon(index).slot
                SendPlayerPokemonData index, Player(index, TempPlayer(index).UseChar).Map
                
                '//Check New Move
                CheckNewMove index, PlayerPokemon(index).slot
                
                AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " evolved " & Trim$(Pokemon(PokeNum).Name) & " to " & Trim$(Pokemon(EvolveNum).Name)
                Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "Congratulations! Your " & Trim$(Pokemon(PokeNum).Name) & " has evolved into " & Trim$(Pokemon(EvolveNum).Name), White
                    Case LANG_EN: AddAlert index, "Congratulations! Your " & Trim$(Pokemon(PokeNum).Name) & " has evolved into " & Trim$(Pokemon(EvolveNum).Name), White
                    Case LANG_ES: AddAlert index, "Congratulations! Your " & Trim$(Pokemon(PokeNum).Name) & " has evolved into " & Trim$(Pokemon(EvolveNum).Name), White
                End Select

                '//Add pokedex
                AddPlayerPokedex index, .Num, YES, YES
                SendPlaySound "evolve.wav", Player(index, TempPlayer(index).UseChar).Map
            End With
        End If
    End If
End Sub

Private Sub HandleUseItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim itemSlot As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    itemSlot = buffer.ReadByte
    Set buffer = Nothing
    
    PlayerUseItem index, itemSlot
    TempPlayer(index).MapSwitchTmr = NO
End Sub

Private Sub HandleSwitchInvSlot(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim OldSlot As Byte, NewSlot As Byte
Dim OldInvData As PlayerInvDataRec, NewInvData As PlayerInvDataRec

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    OldSlot = buffer.ReadByte
    NewSlot = buffer.ReadByte
    Set buffer = Nothing
    
    If OldSlot <= 0 Then Exit Sub
    If OldSlot > MAX_PLAYER_INV Then Exit Sub
    If NewSlot <= 0 Then Exit Sub
    If NewSlot > MAX_PLAYER_INV Then Exit Sub
    
    '//Store Data
    OldInvData = PlayerInv(index).Data(OldSlot)
    NewInvData = PlayerInv(index).Data(NewSlot)
    
    '//Replace Data
    PlayerInv(index).Data(OldSlot) = NewInvData
    PlayerInv(index).Data(NewSlot) = OldInvData
    
    '//Update
    SendPlayerInvSlot index, OldSlot
    SendPlayerInvSlot index, NewSlot
End Sub

Private Sub HandleGotData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim InUsed As Byte
Dim Data1 As Long, Data2 As Long, Data3 As Long
Dim CatchRate As Single
Dim CatchValue As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    InUsed = buffer.ReadByte
    Data1 = buffer.ReadLong
    Data2 = buffer.ReadLong
    Data3 = buffer.ReadLong
    Set buffer = Nothing
    
    If InUsed = YES Then
        '//Make Sure Usage of item is available
        If Not IsPlaying(index) Then GoTo Continue
        If TempPlayer(index).UseChar <= 0 Then GoTo Continue
        If TempPlayer(index).TmpUseInvSlot <= 0 Then GoTo Continue
        If PlayerInv(index).Data(TempPlayer(index).TmpUseInvSlot).Num <= 0 Then GoTo Continue
        
        '//Check Type
        Select Case Item(PlayerInv(index).Data(TempPlayer(index).TmpUseInvSlot).Num).Type
            Case ItemTypeEnum.Pokeball
                If TempPlayer(index).TmpCatchPokeNum > 0 Then
                    Select Case TempPlayer(index).CurLanguage
                        Case LANG_PT: AddAlert index, "You are currently capturing a pokemon", White
                        Case LANG_EN: AddAlert index, "You are currently capturing a pokemon", White
                        Case LANG_ES: AddAlert index, "You are currently capturing a pokemon", White
                    End Select
                    GoTo Continue
                End If
                
                '//Make sure we still have slot
                If CountFreePokemonSlot(index) <= 0 Then
                    Select Case TempPlayer(index).CurLanguage
                        Case LANG_PT: AddAlert index, "You don't have any free slot left to capture a pokemon", White
                        Case LANG_EN: AddAlert index, "You don't have any free slot left to capture a pokemon", White
                        Case LANG_ES: AddAlert index, "You don't have any free slot left to capture a pokemon", White
                    End Select
                    GoTo Continue
                End If
                
                '//Make sure pokemon is not empty
                If Data1 <= 0 Or Data1 > MAX_GAME_POKEMON Then GoTo Continue
                '//Check if exist on map
                If MapPokemon(Data1).Num <= 0 Then GoTo Continue
                If Not MapPokemon(Data1).Map = Player(index, TempPlayer(index).UseChar).Map Then GoTo Continue
                
                '//Check if inrange
                If MapPokemon(Data1).x < Player(index, TempPlayer(index).UseChar).x - 4 Or MapPokemon(Data1).x > Player(index, TempPlayer(index).UseChar).x + 4 Or MapPokemon(Data1).y < Player(index, TempPlayer(index).UseChar).y - 4 Or MapPokemon(Data1).y > Player(index, TempPlayer(index).UseChar).y + 4 Then
                    Select Case TempPlayer(index).CurLanguage
                        Case LANG_PT: AddAlert index, "Out of range", White
                        Case LANG_EN: AddAlert index, "Out of range", White
                        Case LANG_ES: AddAlert index, "Out of range", White
                    End Select
                    GoTo Continue
                End If
                
                '//Check if catchable
                If Player(index, TempPlayer(index).UseChar).Access < ACCESS_CREATOR Then
                    If Spawn(Data1).CanCatch = YES Then
                        Select Case TempPlayer(index).CurLanguage
                            Case LANG_PT: AddAlert index, "You cannot catch this Pokemon", White
                            Case LANG_EN: AddAlert index, "You cannot catch this Pokemon", White
                            Case LANG_ES: AddAlert index, "You cannot catch this Pokemon", White
                        End Select
                        GoTo Continue
                    End If
                End If
                
                '//Make sure no one is trying to catch this pokemon
                If MapPokemon(Data1).InCatch = YES Then GoTo Continue

                If Item(PlayerInv(index).Data(TempPlayer(index).TmpUseInvSlot).Num).Data3 = YES Then
                    '// Success
                    '//Give Pokemon
                    If CountFreePokemonSlot(index) < 5 Then
                        Select Case TempPlayer(index).CurLanguage
                            Case LANG_PT: AddAlert index, "Warning: You only have few slot left for pokemon", White
                            Case LANG_EN: AddAlert index, "Warning: You only have few slot left for pokemon", White
                            Case LANG_ES: AddAlert index, "Warning: You only have few slot left for pokemon", White
                        End Select
                    End If
                        
                    '//Give Player Pokemon
                    If CatchMapPokemonData(index, Data1, TempPlayer(index).TmpCatchUseBall) Then
                        '//Success
                        '//Clear Map Pokemon
                        TempPlayer(index).TmpCatchUseBall = Item(PlayerInv(index).Data(TempPlayer(index).TmpUseInvSlot).Num).Data2
                        SendMapPokemonCatchState MapPokemon(Data1).Map, Data1, MapPokemon(Data1).x, MapPokemon(Data1).y, 2, TempPlayer(index).TmpCatchUseBall '// 0 = Init, 1 = Shake, 2 = Success, 3 = Fail
                        ClearMapPokemon Data1
                            
                        TempPlayer(index).TmpCatchPokeNum = 0
                        TempPlayer(index).TmpCatchTimer = 0
                        TempPlayer(index).TmpCatchTries = 0
                        TempPlayer(index).TmpCatchValue = 0
                        TempPlayer(index).TmpCatchUseBall = 0
                    Else
                        '//Broke
                        MapPokemon(Data1).InCatch = NO
                        MapPokemon(Data1).targetType = TARGET_TYPE_PLAYER
                        MapPokemon(Data1).TargetIndex = index
                        SendMapPokemonCatchState MapPokemon(Data1).Map, Data1, MapPokemon(Data1).x, MapPokemon(Data1).y, 3, TempPlayer(index).TmpCatchUseBall '// 0 = Init, 1 = Shake, 2 = Success, 3 = Fail
                        TempPlayer(index).TmpCatchPokeNum = 0
                        TempPlayer(index).TmpCatchTimer = 0
                        TempPlayer(index).TmpCatchTries = 0
                        TempPlayer(index).TmpCatchValue = 0
                        TempPlayer(index).TmpCatchUseBall = 0
                        Select Case TempPlayer(index).CurLanguage
                            Case LANG_PT: AddAlert index, "Your Pokeball broke", White
                            Case LANG_EN: AddAlert index, "Your Pokeball broke", White
                            Case LANG_ES: AddAlert index, "Your Pokeball broke", White
                        End Select
                    End If
                Else
                    '//Do Catch
                    CatchRate = Item(PlayerInv(index).Data(TempPlayer(index).TmpUseInvSlot).Num).Data1
                    CatchRate = CatchRate / 10
                    
                    If MapPokemon(Data1).CurHP > 0 Then
                        '//ToDo: 1 = Status Modifier
                        CatchValue = (((3 * MapPokemon(Data1).MaxHP - 2 * MapPokemon(Data1).CurHP) * Pokemon(MapPokemon(Data1).Num).CatchRate * CatchRate) / (3 * MapPokemon(Data1).MaxHP)) * 1
                        
                        If CatchValue >= 255 Then
                            '// Success
                            '//Give Pokemon
                            If CountFreePokemonSlot(index) < 5 Then
                                Select Case TempPlayer(index).CurLanguage
                                    Case LANG_PT: AddAlert index, "Warning: You only have few slot left for pokemon", White
                                    Case LANG_EN: AddAlert index, "Warning: You only have few slot left for pokemon", White
                                    Case LANG_ES: AddAlert index, "Warning: You only have few slot left for pokemon", White
                                End Select
                            End If
                            
                            '//Give Player Pokemon
                            If CatchMapPokemonData(index, Data1, TempPlayer(index).TmpCatchUseBall) Then
                                '//Success
                                '//Clear Map Pokemon
                                TempPlayer(index).TmpCatchUseBall = Item(PlayerInv(index).Data(TempPlayer(index).TmpUseInvSlot).Num).Data2
                                SendMapPokemonCatchState MapPokemon(Data1).Map, Data1, MapPokemon(Data1).x, MapPokemon(Data1).y, 2, TempPlayer(index).TmpCatchUseBall '// 0 = Init, 1 = Shake, 2 = Success, 3 = Fail
                                ClearMapPokemon Data1
                                
                                TempPlayer(index).TmpCatchPokeNum = 0
                                TempPlayer(index).TmpCatchTimer = 0
                                TempPlayer(index).TmpCatchTries = 0
                                TempPlayer(index).TmpCatchValue = 0
                                TempPlayer(index).TmpCatchUseBall = 0
                            Else
                                '//Broke
                                MapPokemon(Data1).InCatch = NO
                                MapPokemon(Data1).targetType = TARGET_TYPE_PLAYER
                                MapPokemon(Data1).TargetIndex = index
                                SendMapPokemonCatchState MapPokemon(Data1).Map, Data1, MapPokemon(Data1).x, MapPokemon(Data1).y, 3, TempPlayer(index).TmpCatchUseBall '// 0 = Init, 1 = Shake, 2 = Success, 3 = Fail
                                TempPlayer(index).TmpCatchPokeNum = 0
                                TempPlayer(index).TmpCatchTimer = 0
                                TempPlayer(index).TmpCatchTries = 0
                                TempPlayer(index).TmpCatchValue = 0
                                TempPlayer(index).TmpCatchUseBall = 0
                                Select Case TempPlayer(index).CurLanguage
                                    Case LANG_PT: AddAlert index, "Your Pokeball broke", White
                                    Case LANG_EN: AddAlert index, "Your Pokeball broke", White
                                    Case LANG_ES: AddAlert index, "Your Pokeball broke", White
                                End Select
                            End If
                        Else
                            TempPlayer(index).TmpCatchPokeNum = Data1
                            TempPlayer(index).TmpCatchTimer = GetTickCount + 250
                            TempPlayer(index).TmpCatchTries = 0
                            TempPlayer(index).TmpCatchValue = CatchValue
                            TempPlayer(index).TmpCatchUseBall = Item(PlayerInv(index).Data(TempPlayer(index).TmpUseInvSlot).Num).Data2 '//ToDo: Pokeball
                            MapPokemon(Data1).InCatch = YES
                            SendMapPokemonCatchState MapPokemon(Data1).Map, Data1, MapPokemon(Data1).x, MapPokemon(Data1).y, 0, TempPlayer(index).TmpCatchUseBall '// 0 = Init, 1 = Shake, 2 = Success, 3 = Fail
                        End If
                    End If
                End If
            Case ItemTypeEnum.Medicine
                '//Revive
                If Item(PlayerInv(index).Data(TempPlayer(index).TmpUseInvSlot).Num).Data1 = 4 Then
                    If Data1 <= 0 Or Data1 > MAX_PLAYER_POKEMON Then Exit Sub
                    
                    PlayerPokemons(index).Data(Data1).CurHP = PlayerPokemons(index).Data(Data1).MaxHP * (Item(PlayerInv(index).Data(TempPlayer(index).TmpUseInvSlot).Num).Data2 / 100)
                    SendPlayerPokemonSlot index, Data1
                    
                    Select Case TempPlayer(index).CurLanguage
                        Case LANG_PT: AddAlert index, "Pokemon was revived", White
                        Case LANG_EN: AddAlert index, "Pokemon was revived", White
                        Case LANG_ES: AddAlert index, "Pokemon was revived", White
                    End Select
                End If
        End Select
        
        '//Take Item
        PlayerInv(index).Data(TempPlayer(index).TmpUseInvSlot).Value = PlayerInv(index).Data(TempPlayer(index).TmpUseInvSlot).Value - 1
        If PlayerInv(index).Data(TempPlayer(index).TmpUseInvSlot).Value <= 0 Then
            '//Clear Item
            PlayerInv(index).Data(TempPlayer(index).TmpUseInvSlot).Num = 0
            PlayerInv(index).Data(TempPlayer(index).TmpUseInvSlot).Value = 0
        End If
        SendPlayerInvSlot index, TempPlayer(index).TmpUseInvSlot
        TempPlayer(index).TmpUseInvSlot = 0
    End If
    
Continue:
    'AddAlert Index, "Invalid Target", White
    '//Clear
    TempPlayer(index).TmpUseInvSlot = 0
End Sub

Private Sub HandleOpenStorage(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim StorageType As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    StorageType = buffer.ReadByte
    Set buffer = Nothing
    
    If StorageType > 0 Then
        If TempPlayer(index).StorageType > 0 Then Exit Sub
        TempPlayer(index).StorageType = StorageType
    Else
        TempPlayer(index).StorageType = StorageType
    End If
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " enter storage"
    SendStorage index
End Sub

Private Sub HandleDepositItemTo(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim storageSlot As Byte, StorageData As Byte, invSlot As Byte
Dim checkSameSlot As Byte
Dim gameValue As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    storageSlot = buffer.ReadByte
    StorageData = buffer.ReadByte
    invSlot = buffer.ReadByte
    gameValue = buffer.ReadLong
    Set buffer = Nothing
    
    If invSlot <= 0 Or invSlot > MAX_PLAYER_INV Then Exit Sub
    
    If gameValue <= 0 Then Exit Sub
    
    If StorageData > 0 Then
        checkSameSlot = FindSameInvStorageSlot(index, storageSlot, PlayerInv(index).Data(invSlot).Num)
        
        If checkSameSlot > 0 Then
            '//Place item to that part
            If DepositItem(index, storageSlot, checkSameSlot, invSlot, gameValue) Then
                '//Update
                SendPlayerInvSlot index, invSlot
                SendPlayerInvStorageSlot index, storageSlot, checkSameSlot
            End If
        Else
            '//Store Manually
            If PlayerInvStorage(index).slot(storageSlot).Unlocked = YES Then
                '//Check if slot is already taken
                If PlayerInvStorage(index).slot(storageSlot).Data(StorageData).Num > 0 Then
                    '//Deposit Normally
                    If GiveStorageItem(index, storageSlot, PlayerInv(index).Data(invSlot).Num, gameValue) Then
                        PlayerInv(index).Data(invSlot).Value = PlayerInv(index).Data(invSlot).Value - gameValue
                        If PlayerInv(index).Data(invSlot).Value <= 0 Then
                            PlayerInv(index).Data(invSlot).Num = 0
                            PlayerInv(index).Data(invSlot).Value = 0
                        End If
                        '//Update
                        SendPlayerInvSlot index, invSlot
                    Else
                        Select Case TempPlayer(index).CurLanguage
                            Case LANG_PT: AddAlert index, "Inventory is full", White
                            Case LANG_EN: AddAlert index, "Inventory is full", White
                            Case LANG_ES: AddAlert index, "Inventory is full", White
                        End Select
                    End If
                Else
                    '//Place item to that part
                    If DepositItem(index, storageSlot, StorageData, invSlot, gameValue) Then
                        '//Update
                        SendPlayerInvSlot index, invSlot
                        SendPlayerInvStorageSlot index, storageSlot, StorageData
                    End If
                End If
            End If
        End If
    Else
        '//Deposit Normally
        If GiveStorageItem(index, storageSlot, PlayerInv(index).Data(invSlot).Num, gameValue) Then
            PlayerInv(index).Data(invSlot).Value = PlayerInv(index).Data(invSlot).Value - gameValue
            If PlayerInv(index).Data(invSlot).Value <= 0 Then
                PlayerInv(index).Data(invSlot).Num = 0
                PlayerInv(index).Data(invSlot).Value = 0
            End If
            '//Update
            SendPlayerInvSlot index, invSlot
        Else
            Select Case TempPlayer(index).CurLanguage
                Case LANG_PT: AddAlert index, "Inventory is full", White
                Case LANG_EN: AddAlert index, "Inventory is full", White
                Case LANG_ES: AddAlert index, "Inventory is full", White
            End Select
        End If
    End If
End Sub

Private Sub HandleSwitchStorageSlot(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim storageSlot As Byte, OldSlot As Byte, NewSlot As Byte
Dim OldStorageData As PlayerInvStorageDataRec, NewStorageData As PlayerInvStorageDataRec

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    storageSlot = buffer.ReadByte
    OldSlot = buffer.ReadByte
    NewSlot = buffer.ReadByte
    Set buffer = Nothing
    
    If storageSlot <= 0 Or storageSlot > MAX_STORAGE_SLOT Then Exit Sub
    If OldSlot <= 0 Or OldSlot > MAX_STORAGE Then Exit Sub
    If NewSlot <= 0 Or NewSlot > MAX_STORAGE Then Exit Sub

    '//Store Data
    OldStorageData = PlayerInvStorage(index).slot(storageSlot).Data(OldSlot)
    NewStorageData = PlayerInvStorage(index).slot(storageSlot).Data(NewSlot)
    
    '//Replace Data
    PlayerInvStorage(index).slot(storageSlot).Data(OldSlot) = NewStorageData
    PlayerInvStorage(index).slot(storageSlot).Data(NewSlot) = OldStorageData
    
    '//Update
    SendPlayerInvStorageSlot index, storageSlot, OldSlot
    SendPlayerInvStorageSlot index, storageSlot, NewSlot
End Sub

Private Sub HandleWithdrawItemTo(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim storageSlot As Byte, StorageData As Byte, invSlot As Byte
Dim checkSameSlot As Byte
Dim gameValue As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    storageSlot = buffer.ReadByte
    StorageData = buffer.ReadByte
    invSlot = buffer.ReadByte
    gameValue = buffer.ReadLong
    Set buffer = Nothing
    
    If storageSlot <= 0 Or storageSlot > MAX_STORAGE_SLOT Then Exit Sub
    If gameValue <= 0 Then Exit Sub
    
    If invSlot > 0 Then
        checkSameSlot = FindSameItemSlot(index, PlayerInvStorage(index).slot(storageSlot).Data(StorageData).Num)
        
        If checkSameSlot > 0 Then
            '//Place item to that part
            If WithdrawItem(index, storageSlot, StorageData, checkSameSlot, gameValue) Then
                '//Update
                SendPlayerInvSlot index, checkSameSlot
                SendPlayerInvStorageSlot index, storageSlot, StorageData
            End If
        Else
            '//Check if slot is already taken
            If PlayerInv(index).Data(invSlot).Num > 0 Then
                '//Deposit Normally
                If GiveItem(index, PlayerInvStorage(index).slot(storageSlot).Data(StorageData).Num, gameValue) Then
                    PlayerInvStorage(index).slot(storageSlot).Data(StorageData).Value = PlayerInvStorage(index).slot(storageSlot).Data(StorageData).Value - gameValue
                    If PlayerInvStorage(index).slot(storageSlot).Data(StorageData).Value <= 0 Then
                        PlayerInvStorage(index).slot(storageSlot).Data(StorageData).Num = 0
                        PlayerInvStorage(index).slot(storageSlot).Data(StorageData).Value = 0
                    End If
                    '//Update
                    SendPlayerInvStorageSlot index, storageSlot, StorageData
                Else
                    Select Case TempPlayer(index).CurLanguage
                        Case LANG_PT: AddAlert index, "Inventory is full", White
                        Case LANG_EN: AddAlert index, "Inventory is full", White
                        Case LANG_ES: AddAlert index, "Inventory is full", White
                    End Select
                End If
            Else
                '//Place item to that part
                If WithdrawItem(index, storageSlot, StorageData, invSlot, gameValue) Then
                    '//Update
                    SendPlayerInvSlot index, invSlot
                    SendPlayerInvStorageSlot index, storageSlot, StorageData
                End If
            End If
        End If
    Else
        '//Deposit Normally
        If GiveItem(index, PlayerInvStorage(index).slot(storageSlot).Data(StorageData).Num, gameValue) Then
            PlayerInvStorage(index).slot(storageSlot).Data(StorageData).Value = PlayerInvStorage(index).slot(storageSlot).Data(StorageData).Value - gameValue
            If PlayerInvStorage(index).slot(storageSlot).Data(StorageData).Value <= 0 Then
                PlayerInvStorage(index).slot(storageSlot).Data(StorageData).Num = 0
                PlayerInvStorage(index).slot(storageSlot).Data(StorageData).Value = 0
            End If
            '//Update
            SendPlayerInvStorageSlot index, storageSlot, StorageData
        Else
            Select Case TempPlayer(index).CurLanguage
                Case LANG_PT: AddAlert index, "Inventory is full", White
                Case LANG_EN: AddAlert index, "Inventory is full", White
                Case LANG_ES: AddAlert index, "Inventory is full", White
            End Select
        End If
    End If
End Sub

Private Sub HandleConvo(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim cType As Byte
Dim Data1 As Long
Dim NpcNum As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    cType = buffer.ReadByte
    Data1 = buffer.ReadLong
    Set buffer = Nothing
    
    If cType = 1 Then
        '//Check if npc have convo
        If Data1 > 0 And Data1 <= MAX_MAP_NPC Then
            '//Make sure NPC is not busy
            If MapNpc(Player(index, TempPlayer(index).UseChar).Map, Data1).InBattle <= 0 Then
                NpcNum = MapNpc(Player(index, TempPlayer(index).UseChar).Map, Data1).Num
                If NpcNum > 0 Then
                    If Npc(NpcNum).Convo > 0 Then
                        TempPlayer(index).CurConvoNum = Npc(NpcNum).Convo
                        TempPlayer(index).CurConvoData = 0 '//Always start at 0
                        TempPlayer(index).CurConvoNpc = NpcNum
                        TempPlayer(index).CurConvoMapNpc = Data1
                        ProcessConversation index, TempPlayer(index).CurConvoNum, TempPlayer(index).CurConvoData, TempPlayer(index).CurConvoNpc
                    End If
                End If
            Else
                Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "You cannot talk with this NPC at the moment", White
                    Case LANG_EN: AddAlert index, "You cannot talk with this NPC at the moment", White
                    Case LANG_ES: AddAlert index, "You cannot talk with this NPC at the moment", White
                End Select
            End If
        End If
    ElseIf cType = 2 Then
        If Data1 > 0 And Data1 <= MAX_CONVERSATION Then
            TempPlayer(index).CurConvoNum = Data1
            TempPlayer(index).CurConvoData = 0 '//Always start at 0
            TempPlayer(index).CurConvoNpc = 0
            TempPlayer(index).CurConvoMapNpc = 0
            ProcessConversation index, TempPlayer(index).CurConvoNum, TempPlayer(index).CurConvoData, TempPlayer(index).CurConvoNpc
        End If
    End If
End Sub

Private Sub HandleProcessConvo(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim tReply As Byte

    If TempPlayer(index).CurConvoNum <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    tReply = buffer.ReadByte
    Set buffer = Nothing
    
    ProcessConversation index, TempPlayer(index).CurConvoNum, TempPlayer(index).CurConvoData, TempPlayer(index).CurConvoNpc, tReply
End Sub

Private Sub HandleDepositPokemon(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim PokeSlot As Byte
Dim storageSlot As Byte
Dim StorageData As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    storageSlot = buffer.ReadByte
    PokeSlot = buffer.ReadByte
    Set buffer = Nothing
    
    '//Make sure they don't deposit their last pokemon
    If CountPlayerPokemon(index) <= 1 Then
        Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: AddAlert index, "You cannot deposit your last pokemon", White
            Case LANG_EN: AddAlert index, "You cannot deposit your last pokemon", White
            Case LANG_ES: AddAlert index, "You cannot deposit your last pokemon", White
        End Select
        Exit Sub
    End If
    If PokeSlot <= 0 Or PokeSlot > MAX_PLAYER_POKEMON Then Exit Sub
    
    StorageData = FindFreePokeStorageSlot(index, storageSlot)
    '//Check if there's available slot
    If StorageData > 0 Then
        'CopyMemory ByVal VarPtr(Npc(xIndex)), ByVal VarPtr(dData(0)), dSize
        Call CopyMemory(ByVal VarPtr(PlayerPokemonStorage(index).slot(storageSlot).Data(StorageData)), ByVal VarPtr(PlayerPokemons(index).Data(PokeSlot)), LenB(PlayerPokemons(index).Data(PokeSlot)))
        Call ZeroMemory(ByVal VarPtr(PlayerPokemons(index).Data(PokeSlot)), LenB(PlayerPokemons(index).Data(PokeSlot)))
        '//reupdate order
        UpdatePlayerPokemonOrder index
        '//update
        SendPlayerPokemons index
        SendPlayerPokemonStorageSlot index, storageSlot, StorageData
    Else
        Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: AddAlert index, "There's no available space on this slot", White
            Case LANG_EN: AddAlert index, "There's no available space on this slot", White
            Case LANG_ES: AddAlert index, "There's no available space on this slot", White
        End Select
        Exit Sub
    End If
End Sub

Private Sub HandleWithdrawPokemon(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim PokeSlot As Byte
Dim storageSlot As Byte
Dim StorageData As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    storageSlot = buffer.ReadByte
    StorageData = buffer.ReadByte
    Set buffer = Nothing
    
    '//Make sure they don't deposit their last pokemon
    If CountPlayerPokemon(index) >= MAX_PLAYER_POKEMON Then
        Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: AddAlert index, "There's no slot available", White
            Case LANG_EN: AddAlert index, "There's no slot available", White
            Case LANG_ES: AddAlert index, "There's no slot available", White
        End Select
        Exit Sub
    End If
    If StorageData <= 0 Or StorageData > MAX_STORAGE Then Exit Sub
    
    PokeSlot = FindOpenPokeSlot(index)
    '//Check if there's available slot
    If PokeSlot > 0 Then
        'CopyMemory ByVal VarPtr(Npc(xIndex)), ByVal VarPtr(dData(0)), dSize
        Call CopyMemory(ByVal VarPtr(PlayerPokemons(index).Data(PokeSlot)), ByVal VarPtr(PlayerPokemonStorage(index).slot(storageSlot).Data(StorageData)), LenB(PlayerPokemonStorage(index).slot(storageSlot).Data(StorageData)))
        Call ZeroMemory(ByVal VarPtr(PlayerPokemonStorage(index).slot(storageSlot).Data(StorageData)), LenB(PlayerPokemonStorage(index).slot(storageSlot).Data(StorageData)))
        '//reupdate order
        UpdatePlayerPokemonOrder index
        '//update
        SendPlayerPokemons index
        SendPlayerPokemonStorageSlot index, storageSlot, StorageData
    Else
        Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: AddAlert index, "There's no slot available", White
            Case LANG_EN: AddAlert index, "There's no slot available", White
            Case LANG_ES: AddAlert index, "There's no slot available", White
        End Select
        Exit Sub
    End If
End Sub

Private Sub HandleReleasePokemon(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim storageSlot As Byte
Dim StorageData As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    storageSlot = buffer.ReadByte
    StorageData = buffer.ReadByte
    Set buffer = Nothing
    
    If storageSlot <= 0 Or storageSlot > MAX_STORAGE_SLOT Then Exit Sub
    If StorageData <= 0 Or StorageData > MAX_STORAGE Then Exit Sub
    
    If PlayerPokemonStorage(index).slot(storageSlot).Data(StorageData).Num > 0 Then
        Call ZeroMemory(ByVal VarPtr(PlayerPokemonStorage(index).slot(storageSlot).Data(StorageData)), LenB(PlayerPokemonStorage(index).slot(storageSlot).Data(StorageData)))
        SendPlayerPokemonStorageSlot index, storageSlot, StorageData
    End If
End Sub

Private Sub HandleSwitchStoragePokeSlot(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim storageSlot As Byte, OldSlot As Byte, NewSlot As Byte
Dim OldStorageData As PlayerPokemonStorageDataRec, NewStorageData As PlayerPokemonStorageDataRec

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    storageSlot = buffer.ReadByte
    OldSlot = buffer.ReadByte
    NewSlot = buffer.ReadByte
    Set buffer = Nothing
    
    If storageSlot <= 0 Or storageSlot > MAX_STORAGE_SLOT Then Exit Sub
    If OldSlot <= 0 Or OldSlot > MAX_STORAGE Then Exit Sub
    If NewSlot <= 0 Or NewSlot > MAX_STORAGE Then Exit Sub

    '//Store Data
    OldStorageData = PlayerPokemonStorage(index).slot(storageSlot).Data(OldSlot)
    NewStorageData = PlayerPokemonStorage(index).slot(storageSlot).Data(NewSlot)
    
    '//Replace Data
    PlayerPokemonStorage(index).slot(storageSlot).Data(OldSlot) = NewStorageData
    PlayerPokemonStorage(index).slot(storageSlot).Data(NewSlot) = OldStorageData
    
    '//Update
    SendPlayerPokemonStorageSlot index, storageSlot, OldSlot
    SendPlayerPokemonStorageSlot index, storageSlot, NewSlot
End Sub

Private Sub HandleCloseShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    TempPlayer(index).InShop = 0
    SendOpenShop index
End Sub

Private Sub HandleBuyItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim ShopSlot As Byte, ShopVal As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If TempPlayer(index).InShop <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ShopSlot = buffer.ReadByte
    ShopVal = buffer.ReadLong
    Set buffer = Nothing
    
    If ShopSlot <= 0 Or ShopSlot > MAX_SHOP_ITEM Then Exit Sub
    If ShopVal <= 0 Then Exit Sub
    
    '//Give Item
    With Player(index, TempPlayer(index).UseChar)
        If .Money >= (Shop(TempPlayer(index).InShop).ShopItem(ShopSlot).Price * ShopVal) Then
            If GiveItem(index, Shop(TempPlayer(index).InShop).ShopItem(ShopSlot).Num, ShopVal) Then
                '//Take money
                .Money = .Money - (Shop(TempPlayer(index).InShop).ShopItem(ShopSlot).Price * ShopVal)
                SendPlayerData index
                Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "You have successfully bought x" & ShopVal & " " & Trim$(Item(Shop(TempPlayer(index).InShop).ShopItem(ShopSlot).Num).Name), White
                    Case LANG_EN: AddAlert index, "You have successfully bought x" & ShopVal & " " & Trim$(Item(Shop(TempPlayer(index).InShop).ShopItem(ShopSlot).Num).Name), White
                    Case LANG_ES: AddAlert index, "You have successfully bought x" & ShopVal & " " & Trim$(Item(Shop(TempPlayer(index).InShop).ShopItem(ShopSlot).Num).Name), White
                End Select
            Else
                '//No slot left
                Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "Inventory is full", White
                    Case LANG_EN: AddAlert index, "Inventory is full", White
                    Case LANG_ES: AddAlert index, "Inventory is full", White
                End Select
            End If
        End If
    End With
End Sub

Private Sub HandleSellItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim invSlot As Byte, InvVal As Long
Dim aPrice As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If TempPlayer(index).InShop <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    invSlot = buffer.ReadByte
    InvVal = buffer.ReadLong
    Set buffer = Nothing
    
    If invSlot <= 0 Or invSlot > MAX_PLAYER_INV Then Exit Sub
    If InvVal < 0 Then Exit Sub
    
    '//Give Item
    With Player(index, TempPlayer(index).UseChar)
        If PlayerInv(index).Data(invSlot).Value < InvVal Then
            Select Case TempPlayer(index).CurLanguage
                Case LANG_PT: AddAlert index, "Invalid amount", White
                Case LANG_EN: AddAlert index, "Invalid amount", White
                Case LANG_ES: AddAlert index, "Invalid amount", White
            End Select
        Else
            If PlayerInv(index).Data(invSlot).Num > 0 Then
                aPrice = Item(PlayerInv(index).Data(invSlot).Num).Price * InvVal
                Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "You have successfully sold x" & InvVal & " " & Trim$(Item(PlayerInv(index).Data(invSlot).Num).Name) & " for $" & aPrice, White
                    Case LANG_EN: AddAlert index, "You have successfully sold x" & InvVal & " " & Trim$(Item(PlayerInv(index).Data(invSlot).Num).Name) & " for $" & aPrice, White
                    Case LANG_ES: AddAlert index, "You have successfully sold x" & InvVal & " " & Trim$(Item(PlayerInv(index).Data(invSlot).Num).Name) & " for $" & aPrice, White
                End Select
                PlayerInv(index).Data(invSlot).Value = PlayerInv(index).Data(invSlot).Value - InvVal
                If PlayerInv(index).Data(invSlot).Value <= 0 Then
                    PlayerInv(index).Data(invSlot).Num = 0
                    PlayerInv(index).Data(invSlot).Value = 0
                End If
                SendPlayerInvSlot index, invSlot
                .Money = .Money + aPrice
                If .Money >= MAX_MONEY Then
                    .Money = MAX_MONEY
                End If
                SendPlayerData index
            End If
        End If
    End With
End Sub

Private Sub HandleRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim RequestType As Byte
Dim requestIndex As Long
Dim stillGotTarget As Boolean

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    
    stillGotTarget = False
    If TempPlayer(index).PlayerRequest > 0 Then
        If IsPlaying(TempPlayer(index).PlayerRequest) Then
            If TempPlayer(TempPlayer(index).PlayerRequest).UseChar > 0 Then
                If Player(TempPlayer(index).PlayerRequest, TempPlayer(TempPlayer(index).PlayerRequest).UseChar).Map = Player(index, TempPlayer(index).UseChar).Map Then
                    '//Check if inrange
                    stillGotTarget = True
                End If
            End If
        End If
    End If
    
    If stillGotTarget Then
        Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: AddAlert index, "Cancel your last player request", White
            Case LANG_EN: AddAlert index, "Cancel your last player request", White
            Case LANG_ES: AddAlert index, "Cancel your last player request", White
        End Select
    Else
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        RequestType = buffer.ReadByte
        requestIndex = buffer.ReadLong
        Set buffer = Nothing
        
        If requestIndex <= 0 Or requestIndex > MAX_PLAYER Then Exit Sub
        If Not IsPlaying(requestIndex) Then Exit Sub
        If TempPlayer(requestIndex).UseChar <= 0 Then Exit Sub
        
        If TempPlayer(requestIndex).PlayerRequest > 0 Or TempPlayer(requestIndex).InDuel > 0 Then
            '//Can't duel
            'If RequestType = 1 Then '//Duel
                Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "Player is busy", White
                    Case LANG_EN: AddAlert index, "Player is busy", White
                    Case LANG_ES: AddAlert index, "Player is busy", White
                End Select
            'End If
            TempPlayer(index).RequestType = 0
            TempPlayer(index).PlayerRequest = 0
            SendRequest index
            Exit Sub
        End If
        
        If Map(Player(requestIndex, TempPlayer(requestIndex).UseChar).Map).Moral = 1 Then
            '//Can't duel
            If RequestType = 1 Then '//Duel
                Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "You cannot duel here", White
                    Case LANG_EN: AddAlert index, "You cannot duel here", White
                    Case LANG_ES: AddAlert index, "You cannot duel here", White
                End Select
                TempPlayer(index).RequestType = 0
                TempPlayer(index).PlayerRequest = 0
                SendRequest index
                Exit Sub
            End If
        End If
        
        If Map(Player(index, TempPlayer(index).UseChar).Map).Moral = 1 Then
            '//Can't duel
            If RequestType = 1 Then '//Duel
                Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "You cannot duel here", White
                    Case LANG_EN: AddAlert index, "You cannot duel here", White
                    Case LANG_ES: AddAlert index, "You cannot duel here", White
                End Select
                TempPlayer(index).RequestType = 0
                TempPlayer(index).PlayerRequest = 0
                SendRequest index
                Exit Sub
            End If
        End If
        
        If TempPlayer(requestIndex).StorageType > 0 Then
            '//Can't duel
            'If RequestType = 1 Then '//Duel
                Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "Player is busy", White
                    Case LANG_EN: AddAlert index, "Player is busy", White
                    Case LANG_ES: AddAlert index, "Player is busy", White
                End Select
            'End If
            TempPlayer(index).RequestType = 0
            TempPlayer(index).PlayerRequest = 0
            SendRequest index
            Exit Sub
        End If
        
        If TempPlayer(index).StorageType > 0 Then
            '//Can't duel
            'If RequestType = 1 Then '//Duel
                Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "You can't do that right now", White
                    Case LANG_EN: AddAlert index, "You can't do that right now", White
                    Case LANG_ES: AddAlert index, "You can't do that right now", White
                End Select
            'End If
            TempPlayer(index).RequestType = 0
            TempPlayer(index).PlayerRequest = 0
            SendRequest index
            Exit Sub
        End If
        
        If RequestType = 3 Then
            If TempPlayer(requestIndex).InParty > 0 Then
                Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "Player already have a party", White
                    Case LANG_EN: AddAlert index, "Player already have a party", White
                    Case LANG_ES: AddAlert index, "Player already have a party", White
                End Select
                TempPlayer(index).RequestType = 0
                TempPlayer(index).PlayerRequest = 0
                SendRequest index
                Exit Sub
            End If
            
            If PartyCount(index) >= MAX_PARTY Then
                Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "Party is full", White
                    Case LANG_EN: AddAlert index, "Party is full", White
                    Case LANG_ES: AddAlert index, "Party is full", White
                End Select
                TempPlayer(index).RequestType = 0
                TempPlayer(index).PlayerRequest = 0
                SendRequest index
                Exit Sub
            End If
            
            If TempPlayer(index).InParty <= 0 Then
                Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "You are not in a party", White
                    Case LANG_EN: AddAlert index, "You are not in a party", White
                    Case LANG_ES: AddAlert index, "You are not in a party", White
                End Select
                TempPlayer(index).RequestType = 0
                TempPlayer(index).PlayerRequest = 0
                SendRequest index
                Exit Sub
            End If
        End If
        
        TempPlayer(requestIndex).PlayerRequest = index
        TempPlayer(requestIndex).RequestType = RequestType
        TempPlayer(index).PlayerRequest = requestIndex
        TempPlayer(index).RequestType = RequestType
        '//Send Request to client
        SendRequest requestIndex
    End If
    
    Exit Sub
exitdata:
    SendRequest index
End Sub

Private Sub HandleRequestState(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim requestState As Byte
Dim RequestType As Byte
Dim requestIndex As Long
Dim i As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    requestState = buffer.ReadByte
    Set buffer = Nothing
    
    requestIndex = TempPlayer(index).PlayerRequest
    RequestType = TempPlayer(index).RequestType
    
    Select Case requestState
        Case 0 '//Cancel
            If requestIndex > 0 Then
                '//Cancel Request to index
                If IsPlaying(requestIndex) Then
                    If TempPlayer(requestIndex).UseChar > 0 Then
                        If TempPlayer(requestIndex).PlayerRequest = index Then
                            TempPlayer(requestIndex).PlayerRequest = 0
                            TempPlayer(requestIndex).RequestType = 0
                            SendRequest requestIndex
                            If RequestType = 1 Then '//Duel
                                Select Case TempPlayer(requestIndex).CurLanguage
                                    Case LANG_PT: AddAlert requestIndex, "Duel request has been cancelled", White
                                    Case LANG_EN: AddAlert requestIndex, "Duel request has been cancelled", White
                                    Case LANG_ES: AddAlert requestIndex, "Duel request has been cancelled", White
                                End Select
                            ElseIf RequestType = 2 Then '//Trade
                                Select Case TempPlayer(requestIndex).CurLanguage
                                    Case LANG_PT: AddAlert requestIndex, "Trade request has been cancelled", White
                                    Case LANG_EN: AddAlert requestIndex, "Trade request has been cancelled", White
                                    Case LANG_ES: AddAlert requestIndex, "Trade request has been cancelled", White
                                End Select
                            End If
                        End If
                    End If
                End If
            End If
            TempPlayer(index).PlayerRequest = 0
            TempPlayer(index).RequestType = 0
            SendRequest index
            Select Case TempPlayer(index).CurLanguage
                Case LANG_PT: AddAlert index, "Request has been cancelled", White
                Case LANG_EN: AddAlert index, "Request has been cancelled", White
                Case LANG_ES: AddAlert index, "Request has been cancelled", White
            End Select
        Case 1 '//Accept
            If requestIndex > 0 Then
                '//Cancel Request to index
                If IsPlaying(requestIndex) Then
                    If TempPlayer(requestIndex).UseChar > 0 Then
                        If TempPlayer(requestIndex).PlayerRequest = index Then
                            If RequestType = 1 Then '//Duel
                                '//Initiate duel
                                If CountPlayerPokemonAlive(index) > 0 Then
                                    TempPlayer(index).InDuel = requestIndex
                                    TempPlayer(index).DuelTime = 11
                                    TempPlayer(index).DuelTimeTmr = GetTickCount + 1000
                                    TempPlayer(requestIndex).InDuel = index
                                    TempPlayer(requestIndex).DuelTime = 11
                                    TempPlayer(requestIndex).DuelTimeTmr = GetTickCount + 1000
                                    Select Case TempPlayer(index).CurLanguage
                                        Case LANG_PT: AddAlert index, "Duel invitation accepted", White
                                        Case LANG_EN: AddAlert index, "Duel invitation accepted", White
                                        Case LANG_ES: AddAlert index, "Duel invitation accepted", White
                                    End Select
                                    Select Case TempPlayer(requestIndex).CurLanguage
                                        Case LANG_PT: AddAlert requestIndex, "Duel invitation accepted", White
                                        Case LANG_EN: AddAlert requestIndex, "Duel invitation accepted", White
                                        Case LANG_ES: AddAlert requestIndex, "Duel invitation accepted", White
                                    End Select
                                Else
                                    Select Case TempPlayer(index).CurLanguage
                                        Case LANG_PT: AddAlert index, "You don't have any active pokemon to use", White
                                        Case LANG_EN: AddAlert index, "You don't have any active pokemon to use", White
                                        Case LANG_ES: AddAlert index, "You don't have any active pokemon to use", White
                                    End Select
                                    Select Case TempPlayer(requestIndex).CurLanguage
                                        Case LANG_PT: AddAlert requestIndex, "Player doesn't have any active pokemon to use", White
                                        Case LANG_EN: AddAlert requestIndex, "Player doesn't have any active pokemon to use", White
                                        Case LANG_ES: AddAlert requestIndex, "Player doesn't have any active pokemon to use", White
                                    End Select
                                End If
                            ElseIf RequestType = 2 Then '//Trade
                                TempPlayer(index).InTrade = requestIndex
                                TempPlayer(requestIndex).InTrade = index
                                SendOpenTrade index
                                SendOpenTrade requestIndex
                            ElseIf RequestType = 3 Then '//Party
                                '//Join Party
                                JoinParty requestIndex, index
                                TempPlayer(index).RequestType = 0
                                TempPlayer(index).PlayerRequest = 0
                                SendRequest index
                                TempPlayer(requestIndex).RequestType = 0
                                TempPlayer(requestIndex).PlayerRequest = 0
                                SendRequest requestIndex
                            End If
                        End If
                    End If
                End If
            End If
        Case Else '//Decline
            If requestIndex > 0 Then
                '//Cancel Request to index
                If IsPlaying(requestIndex) Then
                    If TempPlayer(requestIndex).UseChar > 0 Then
                        If TempPlayer(requestIndex).PlayerRequest = index Then
                            TempPlayer(requestIndex).PlayerRequest = 0
                            TempPlayer(requestIndex).RequestType = 0
                            SendRequest requestIndex
                            Select Case TempPlayer(requestIndex).CurLanguage
                                Case LANG_PT: AddAlert requestIndex, "Request has been declined", White
                                Case LANG_EN: AddAlert requestIndex, "Request has been declined", White
                                Case LANG_ES: AddAlert requestIndex, "Request has been declined", White
                            End Select
                        End If
                    End If
                End If
            End If
            TempPlayer(index).PlayerRequest = 0
            TempPlayer(index).RequestType = 0
            SendRequest index
            Select Case TempPlayer(index).CurLanguage
                Case LANG_PT: AddAlert index, "Request has been declined", White
                Case LANG_EN: AddAlert index, "Request has been declined", White
                Case LANG_ES: AddAlert index, "Request has been declined", White
            End Select
    End Select
End Sub

Private Sub HandleAddTrade(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim TradeType As Byte, TradeSlot As Long, TradeData As Long
Dim slot As Long
Dim i As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If TempPlayer(index).InTrade <= 0 Then Exit Sub
    If TempPlayer(index).TradeSet = YES Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    TradeType = buffer.ReadByte
    TradeSlot = buffer.ReadLong
    TradeData = buffer.ReadLong
    Set buffer = Nothing
    
    '//Make sure you can't use the same slot
    For i = 1 To MAX_TRADE
        If TempPlayer(index).TradeItem(i).TradeSlot = TradeSlot Then
            Select Case TempPlayer(index).CurLanguage
                Case LANG_PT: AddAlert index, "Trade slot already in used", White
                Case LANG_EN: AddAlert index, "Trade slot already in used", White
                Case LANG_ES: AddAlert index, "Trade slot already in used", White
            End Select
            Exit Sub
        End If
    Next
    
    '//Check available trade slot
    slot = FindOpenTradeSlot(index)
    If slot > 0 Then
        With TempPlayer(index).TradeItem(slot)
            Select Case TradeType
                Case 1 '//Item
                    If TradeSlot <= 0 Or TradeSlot > MAX_PLAYER_INV Then Exit Sub
                    If TradeData <= 0 Then Exit Sub
                    
                    .Num = PlayerInv(index).Data(TradeSlot).Num
                    If TradeData > PlayerInv(index).Data(TradeSlot).Value Then
                        .Value = PlayerInv(index).Data(TradeSlot).Value
                    Else
                        .Value = TradeData
                    End If
                    
                    .Level = 0
                    For i = 1 To StatEnum.Stat_Count - 1
                        .Stat(i) = 0
                        .StatIV(i) = 0
                        .StatEV(i) = 0
                    Next
                    .CurHP = 0
                    .MaxHP = 0
                    .Nature = 0
                    .isShiny = 0
                    .Happiness = 0
                    .Gender = 0
                    .Status = 0
                    .CurExp = 0
                    .nextExp = 0
                    For i = 1 To MAX_MOVESET
                        .Moveset(i).Num = 0
                        .Moveset(i).CurPP = 0
                        .Moveset(i).TotalPP = 0
                    Next
                    .BallUsed = 0
                    .HeldItem = 0
                Case 2 '//Pokemon
                    If TradeSlot <= 0 Or TradeSlot > MAX_PLAYER_POKEMON Then Exit Sub
                    
                    .Num = PlayerPokemons(index).Data(TradeSlot).Num
                    .Value = 0
                    
                    .Level = PlayerPokemons(index).Data(TradeSlot).Level
                    For i = 1 To StatEnum.Stat_Count - 1
                        .Stat(i) = PlayerPokemons(index).Data(TradeSlot).Stat(i).Value
                        .StatIV(i) = PlayerPokemons(index).Data(TradeSlot).Stat(i).IV
                        .StatEV(i) = PlayerPokemons(index).Data(TradeSlot).Stat(i).EV
                    Next
                    .CurHP = PlayerPokemons(index).Data(TradeSlot).CurHP
                    .MaxHP = PlayerPokemons(index).Data(TradeSlot).MaxHP
                    .Nature = PlayerPokemons(index).Data(TradeSlot).Nature
                    .isShiny = PlayerPokemons(index).Data(TradeSlot).isShiny
                    .Happiness = PlayerPokemons(index).Data(TradeSlot).Happiness
                    .Gender = PlayerPokemons(index).Data(TradeSlot).Gender
                    .Status = PlayerPokemons(index).Data(TradeSlot).Status
                    .CurExp = PlayerPokemons(index).Data(TradeSlot).CurExp
                    If .Num > 0 Then
                        .nextExp = GetPokemonNextExp(.Level, Pokemon(.Num).GrowthRate)
                    Else
                        .nextExp = 0
                    End If
                    For i = 1 To MAX_MOVESET
                        .Moveset(i).Num = PlayerPokemons(index).Data(TradeSlot).Moveset(i).Num
                        .Moveset(i).CurPP = PlayerPokemons(index).Data(TradeSlot).Moveset(i).CurPP
                        .Moveset(i).TotalPP = PlayerPokemons(index).Data(TradeSlot).Moveset(i).TotalPP
                    Next
                    .BallUsed = PlayerPokemons(index).Data(TradeSlot).BallUsed
                    .HeldItem = PlayerPokemons(index).Data(TradeSlot).HeldItem
                Case Else
                    '//Error
                    Exit Sub
            End Select
            
            .TradeSlot = TradeSlot
            .Type = TradeType
        End With
        
        '//Update
        SendUpdateTradeItem index, index, slot
        SendUpdateTradeItem TempPlayer(index).InTrade, index, slot
    End If
End Sub

Private Sub HandleRemoveTrade(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim TradeSlot As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If TempPlayer(index).InTrade <= 0 Then Exit Sub
    If TempPlayer(index).TradeSet = YES Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    TradeSlot = buffer.ReadLong
    Set buffer = Nothing
    
    If TradeSlot <= 0 Or TradeSlot > MAX_TRADE Then Exit Sub
    
    Call ZeroMemory(ByVal VarPtr(TempPlayer(index).TradeItem(TradeSlot)), LenB(TempPlayer(index).TradeItem(TradeSlot)))
    '//Update
    SendUpdateTradeItem index, index, TradeSlot
    SendUpdateTradeItem TempPlayer(index).InTrade, index, TradeSlot
End Sub

Private Sub HandleTradeUpdateMoney(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim valMoney As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If TempPlayer(index).InTrade <= 0 Then Exit Sub
    If TempPlayer(index).TradeSet = YES Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    valMoney = buffer.ReadLong
    Set buffer = Nothing
    
    If valMoney > Player(index, TempPlayer(index).UseChar).Money Then
        TempPlayer(index).TradeMoney = Player(index, TempPlayer(index).UseChar).Money
    Else
        TempPlayer(index).TradeMoney = valMoney
    End If
    
    '//Update
    SendTradeUpdateMoney index, index
    SendTradeUpdateMoney TempPlayer(index).InTrade, index
End Sub

Private Sub HandleSetTradeState(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim tSet As Byte

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If TempPlayer(index).InTrade <= 0 Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    tSet = buffer.ReadByte
    Set buffer = Nothing
    
    TempPlayer(index).TradeSet = tSet
    If tSet = YES Then
        '//Set
        Select Case TempPlayer(TempPlayer(index).InTrade).CurLanguage
            Case LANG_PT: AddAlert TempPlayer(index).InTrade, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " set his trade", White
            Case LANG_EN: AddAlert TempPlayer(index).InTrade, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " set his trade", White
            Case LANG_ES: AddAlert TempPlayer(index).InTrade, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " set his trade", White
        End Select
    Else
        '//Cancel
        AddAlert TempPlayer(index).InTrade, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " cancelled his trade", White
    End If
    SendSetTradeState index, index
    SendSetTradeState TempPlayer(index).InTrade, index
End Sub

Private Sub HandleTradeState(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim tState As Byte
Dim i As Byte
Dim tradeIndex As Long
Dim pokemonCount(1 To 2) As Long, itemCount(1 To 2) As Long
Dim countPoke As Long, countItem As Long
Dim PokeSlot As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If TempPlayer(index).InTrade <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    tState = buffer.ReadByte
    Set buffer = Nothing
    
    tradeIndex = TempPlayer(index).InTrade
    If tState = 1 Then '//Accept
        TempPlayer(index).TradeAccept = YES
        Select Case TempPlayer(tradeIndex).CurLanguage
            Case LANG_PT: AddAlert tradeIndex, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " accepted the trade", White
            Case LANG_EN: AddAlert tradeIndex, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " accepted the trade", White
            Case LANG_ES: AddAlert tradeIndex, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " accepted the trade", White
        End Select
        Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: AddAlert index, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " accepted the trade", White
            Case LANG_EN: AddAlert index, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " accepted the trade", White
            Case LANG_ES: AddAlert index, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " accepted the trade", White
        End Select
        '//Check if continue
        If TempPlayer(tradeIndex).TradeAccept = YES Then
            '//Continue trading their items
            pokemonCount(1) = 0
            itemCount(1) = 0
            pokemonCount(2) = 0
            itemCount(2) = 0
            For i = 1 To MAX_TRADE
                If TempPlayer(index).TradeItem(i).Num > 0 Then
                    If TempPlayer(index).TradeItem(i).Type = 1 Then '//Item
                        itemCount(1) = itemCount(1) + 1
                    ElseIf TempPlayer(index).TradeItem(i).Type = 2 Then '//Pokemon
                        pokemonCount(1) = pokemonCount(1) + 1
                    End If
                End If
                If TempPlayer(tradeIndex).TradeItem(i).Num > 0 Then
                    If TempPlayer(tradeIndex).TradeItem(i).Type = 1 Then '//Item
                        itemCount(2) = itemCount(2) + 1
                    ElseIf TempPlayer(tradeIndex).TradeItem(i).Type = 2 Then '//Pokemon
                        pokemonCount(2) = pokemonCount(2) + 1
                    End If
                End If
            Next
            
            '//Check if player can receive item/pokemon
            countPoke = MAX_PLAYER_POKEMON - CountPlayerPokemon(tradeIndex)
            countItem = CountFreeInvSlot(tradeIndex)
            If countItem < itemCount(1) Then
                '//Error
                Select Case TempPlayer(tradeIndex).CurLanguage
                    Case LANG_PT: AddAlert tradeIndex, "You don't have enough inventory slot", White
                    Case LANG_EN: AddAlert tradeIndex, "You don't have enough inventory slot", White
                    Case LANG_ES: AddAlert tradeIndex, "You don't have enough inventory slot", White
                End Select
                Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "Trade Error", White
                    Case LANG_EN: AddAlert index, "Trade Error", White
                    Case LANG_ES: AddAlert index, "Trade Error", White
                End Select
                TempPlayer(index).TradeAccept = NO
                TempPlayer(tradeIndex).TradeAccept = NO
                TempPlayer(index).TradeSet = 0
                TempPlayer(tradeIndex).TradeSet = 0
                SendSetTradeState index, index
                SendSetTradeState tradeIndex, index
                Exit Sub
            End If
            If countPoke < pokemonCount(1) Then
                '//Error
                Select Case TempPlayer(tradeIndex).CurLanguage
                    Case LANG_PT: AddAlert tradeIndex, "You don't have enough pokemon slot", White
                    Case LANG_EN: AddAlert tradeIndex, "You don't have enough pokemon slot", White
                    Case LANG_ES: AddAlert tradeIndex, "You don't have enough pokemon slot", White
                End Select
                Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "Trade Error", White
                    Case LANG_EN: AddAlert index, "Trade Error", White
                    Case LANG_ES: AddAlert index, "Trade Error", White
                End Select
                TempPlayer(index).TradeAccept = NO
                TempPlayer(tradeIndex).TradeAccept = NO
                TempPlayer(index).TradeSet = 0
                TempPlayer(tradeIndex).TradeSet = 0
                SendSetTradeState index, index
                SendSetTradeState tradeIndex, index
                Exit Sub
            End If
            If CountPlayerPokemon(tradeIndex) <= pokemonCount(2) Then
                '//Error
                Select Case TempPlayer(tradeIndex).CurLanguage
                    Case LANG_PT: AddAlert tradeIndex, "You can't trade all your available pokemon", White
                    Case LANG_EN: AddAlert tradeIndex, "You can't trade all your available pokemon", White
                    Case LANG_ES: AddAlert tradeIndex, "You can't trade all your available pokemon", White
                End Select
                Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "Trade Error", White
                    Case LANG_EN: AddAlert index, "Trade Error", White
                    Case LANG_ES: AddAlert index, "Trade Error", White
                End Select
                TempPlayer(index).TradeAccept = NO
                TempPlayer(tradeIndex).TradeAccept = NO
                TempPlayer(index).TradeSet = 0
                TempPlayer(tradeIndex).TradeSet = 0
                SendSetTradeState index, index
                SendSetTradeState tradeIndex, index
                Exit Sub
            End If
            countPoke = MAX_PLAYER_POKEMON - CountPlayerPokemon(index)
            countItem = CountFreeInvSlot(index)
            If countItem < itemCount(2) Then
                '//Error
                Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "You don't have enough inventory slot", White
                    Case LANG_EN: AddAlert index, "You don't have enough inventory slot", White
                    Case LANG_ES: AddAlert index, "You don't have enough inventory slot", White
                End Select
                Select Case TempPlayer(tradeIndex).CurLanguage
                    Case LANG_PT: AddAlert tradeIndex, "Trade Error", White
                    Case LANG_EN: AddAlert tradeIndex, "Trade Error", White
                    Case LANG_ES: AddAlert tradeIndex, "Trade Error", White
                End Select
                TempPlayer(index).TradeAccept = NO
                TempPlayer(tradeIndex).TradeAccept = NO
                TempPlayer(index).TradeSet = 0
                TempPlayer(tradeIndex).TradeSet = 0
                SendSetTradeState index, index
                SendSetTradeState tradeIndex, index
                Exit Sub
            End If
            If countPoke < pokemonCount(2) Then
                '//Error
                Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "You don't have enough pokemon slot", White
                    Case LANG_EN: AddAlert index, "You don't have enough pokemon slot", White
                    Case LANG_ES: AddAlert index, "You don't have enough pokemon slot", White
                End Select
                Select Case TempPlayer(tradeIndex).CurLanguage
                    Case LANG_PT: AddAlert tradeIndex, "Trade Error", White
                    Case LANG_EN: AddAlert tradeIndex, "Trade Error", White
                    Case LANG_ES: AddAlert tradeIndex, "Trade Error", White
                End Select
                TempPlayer(index).TradeAccept = NO
                TempPlayer(tradeIndex).TradeAccept = NO
                TempPlayer(index).TradeSet = 0
                TempPlayer(tradeIndex).TradeSet = 0
                SendSetTradeState index, index
                SendSetTradeState tradeIndex, index
                Exit Sub
            End If
            If CountPlayerPokemon(index) <= pokemonCount(1) Then
                '//Error
                Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "You can't trade all your available pokemon", White
                    Case LANG_EN: AddAlert index, "You can't trade all your available pokemon", White
                    Case LANG_ES: AddAlert index, "You can't trade all your available pokemon", White
                End Select
                Select Case TempPlayer(tradeIndex).CurLanguage
                    Case LANG_PT: AddAlert tradeIndex, "Trade Error", White
                    Case LANG_EN: AddAlert tradeIndex, "Trade Error", White
                    Case LANG_ES: AddAlert tradeIndex, "Trade Error", White
                End Select
                TempPlayer(index).TradeAccept = NO
                TempPlayer(tradeIndex).TradeAccept = NO
                TempPlayer(index).TradeSet = 0
                TempPlayer(tradeIndex).TradeSet = 0
                SendSetTradeState index, index
                SendSetTradeState tradeIndex, index
                Exit Sub
            End If
            
            '//Continue Trading
            For i = 1 To MAX_TRADE
                '//Give Items
                With TempPlayer(index).TradeItem(i)
                    If .Type = 1 Then '//Item
                        If GiveItem(tradeIndex, PlayerInv(index).Data(.TradeSlot).Num, .Value) Then
                            
                        End If
                    ElseIf .Type = 2 Then '//Pokemon
                        PokeSlot = FindOpenPokeSlot(tradeIndex)
                        If PokeSlot > 0 Then
                            Call CopyMemory(ByVal VarPtr(PlayerPokemons(tradeIndex).Data(PokeSlot)), ByVal VarPtr(PlayerPokemons(index).Data(.TradeSlot)), LenB(PlayerPokemons(index).Data(.TradeSlot)))
                            '//Add Pokedex
                            AddPlayerPokedex tradeIndex, PlayerPokemons(tradeIndex).Data(PokeSlot).Num, YES, YES
                            '//reupdate order
                            UpdatePlayerPokemonOrder tradeIndex
                            '//update
                            SendPlayerPokemons tradeIndex
                        End If
                    End If
                End With
                With TempPlayer(tradeIndex).TradeItem(i)
                    If .Type = 1 Then '//Item
                        If GiveItem(index, PlayerInv(tradeIndex).Data(.TradeSlot).Num, .Value) Then
                            '//Take item from index
                            'PlayerInv(index).Data(.TradeSlot).Value = PlayerInv(index).Data(.TradeSlot).Value - .Value
                            'If PlayerInv(index).Data(.TradeSlot).Value <= 0 Then
                            '    PlayerInv(index).Data(.TradeSlot).Num = 0
                            '    PlayerInv(index).Data(.TradeSlot).Value = 0
                            'End If
                            '//Update
                            'SendPlayerInvSlot index, .TradeSlot
                        End If
                    ElseIf .Type = 2 Then '//Pokemon
                        PokeSlot = FindOpenPokeSlot(index)
                        If PokeSlot > 0 Then
                            Call CopyMemory(ByVal VarPtr(PlayerPokemons(index).Data(PokeSlot)), ByVal VarPtr(PlayerPokemons(tradeIndex).Data(.TradeSlot)), LenB(PlayerPokemons(tradeIndex).Data(.TradeSlot)))
                            '//Add Pokedex
                            AddPlayerPokedex index, PlayerPokemons(index).Data(PokeSlot).Num, YES, YES
                            '//reupdate order
                            UpdatePlayerPokemonOrder index
                            '//update
                            SendPlayerPokemons index
                        End If
                    End If
                End With
                
                '//Take Items
                With TempPlayer(index).TradeItem(i)
                    If .Type = 1 Then '//Item
                        PlayerInv(index).Data(.TradeSlot).Value = PlayerInv(index).Data(.TradeSlot).Value - .Value
                        If PlayerInv(index).Data(.TradeSlot).Value <= 0 Then
                            PlayerInv(index).Data(.TradeSlot).Num = 0
                            PlayerInv(index).Data(.TradeSlot).Value = 0
                        End If
                        '//Update
                        SendPlayerInvSlot index, .TradeSlot
                    ElseIf .Type = 2 Then '//Pokemon
                        Call ZeroMemory(ByVal VarPtr(PlayerPokemons(index).Data(.TradeSlot)), LenB(PlayerPokemons(index).Data(.TradeSlot)))
                        '//reupdate order
                        UpdatePlayerPokemonOrder index
                        '//update
                        SendPlayerPokemons index
                    End If
                End With
                With TempPlayer(tradeIndex).TradeItem(i)
                    If .Type = 1 Then '//Item
                        '//Take item from index
                        PlayerInv(tradeIndex).Data(.TradeSlot).Value = PlayerInv(tradeIndex).Data(.TradeSlot).Value - .Value
                        If PlayerInv(tradeIndex).Data(.TradeSlot).Value <= 0 Then
                            PlayerInv(tradeIndex).Data(.TradeSlot).Num = 0
                            PlayerInv(tradeIndex).Data(.TradeSlot).Value = 0
                        End If
                        '//Update
                        SendPlayerInvSlot tradeIndex, .TradeSlot
                    ElseIf .Type = 2 Then '//Pokemon
                        Call ZeroMemory(ByVal VarPtr(PlayerPokemons(tradeIndex).Data(.TradeSlot)), LenB(PlayerPokemons(tradeIndex).Data(.TradeSlot)))
                        '//reupdate order
                        UpdatePlayerPokemonOrder tradeIndex
                        '//update
                        SendPlayerPokemons tradeIndex
                    End If
                End With
            Next
            
            If TempPlayer(index).TradeMoney > 0 And TempPlayer(index).TradeMoney <= Player(index, TempPlayer(index).UseChar).Money Then
                Player(tradeIndex, TempPlayer(tradeIndex).UseChar).Money = Player(tradeIndex, TempPlayer(tradeIndex).UseChar).Money + TempPlayer(index).TradeMoney
                If Player(tradeIndex, TempPlayer(tradeIndex).UseChar).Money >= MAX_MONEY Then
                    Player(tradeIndex, TempPlayer(tradeIndex).UseChar).Money = MAX_MONEY
                End If
                Player(index, TempPlayer(index).UseChar).Money = Player(index, TempPlayer(index).UseChar).Money - TempPlayer(index).TradeMoney
                If Player(index, TempPlayer(index).UseChar).Money <= 0 Then
                    Player(index, TempPlayer(index).UseChar).Money = 0
                End If
                SendPlayerData index
                SendPlayerData tradeIndex
            End If
            If TempPlayer(tradeIndex).TradeMoney > 0 And TempPlayer(tradeIndex).TradeMoney <= Player(tradeIndex, TempPlayer(tradeIndex).UseChar).Money Then
                Player(index, TempPlayer(index).UseChar).Money = Player(index, TempPlayer(index).UseChar).Money + TempPlayer(tradeIndex).TradeMoney
                If Player(index, TempPlayer(index).UseChar).Money >= MAX_MONEY Then
                    Player(index, TempPlayer(index).UseChar).Money = MAX_MONEY
                End If
                Player(tradeIndex, TempPlayer(tradeIndex).UseChar).Money = Player(tradeIndex, TempPlayer(tradeIndex).UseChar).Money - TempPlayer(tradeIndex).TradeMoney
                If Player(tradeIndex, TempPlayer(tradeIndex).UseChar).Money <= 0 Then
                    Player(tradeIndex, TempPlayer(tradeIndex).UseChar).Money = 0
                End If
                SendPlayerData index
                SendPlayerData tradeIndex
            End If
            
            '//Exit Trade
            TempPlayer(index).InTrade = 0
            TempPlayer(tradeIndex).InTrade = 0
            For i = 1 To MAX_TRADE
                Call ZeroMemory(ByVal VarPtr(TempPlayer(index).TradeItem(i)), LenB(TempPlayer(index).TradeItem(i)))
                Call ZeroMemory(ByVal VarPtr(TempPlayer(tradeIndex).TradeItem(i)), LenB(TempPlayer(tradeIndex).TradeItem(i)))
            Next
            TempPlayer(index).TradeMoney = 0
            TempPlayer(tradeIndex).TradeMoney = 0
            TempPlayer(index).TradeSet = 0
            TempPlayer(tradeIndex).TradeSet = 0
            TempPlayer(index).TradeAccept = 0
            TempPlayer(tradeIndex).TradeAccept = 0
            TempPlayer(index).PlayerRequest = 0
            TempPlayer(tradeIndex).PlayerRequest = 0
            TempPlayer(index).RequestType = 0
            TempPlayer(tradeIndex).RequestType = 0
            Select Case TempPlayer(index).CurLanguage
                Case LANG_PT: AddAlert index, "Trade success", White
                Case LANG_EN: AddAlert index, "Trade success", White
                Case LANG_ES: AddAlert index, "Trade success", White
            End Select
            Select Case TempPlayer(tradeIndex).CurLanguage
                Case LANG_PT: AddAlert tradeIndex, "Trade success", White
                Case LANG_EN: AddAlert tradeIndex, "Trade success", White
                Case LANG_ES: AddAlert tradeIndex, "Trade success", White
            End Select
            SendCloseTrade tradeIndex
            SendCloseTrade index
            SendRequest tradeIndex
            SendRequest index
        End If
    Else '//Decline
        TempPlayer(index).InTrade = 0
        TempPlayer(tradeIndex).InTrade = 0
        For i = 1 To MAX_TRADE
            Call ZeroMemory(ByVal VarPtr(TempPlayer(index).TradeItem(i)), LenB(TempPlayer(index).TradeItem(i)))
            Call ZeroMemory(ByVal VarPtr(TempPlayer(tradeIndex).TradeItem(i)), LenB(TempPlayer(tradeIndex).TradeItem(i)))
        Next
        TempPlayer(index).TradeMoney = 0
        TempPlayer(tradeIndex).TradeMoney = 0
        TempPlayer(index).TradeSet = 0
        TempPlayer(tradeIndex).TradeSet = 0
        TempPlayer(index).TradeAccept = 0
        TempPlayer(tradeIndex).TradeAccept = 0
        TempPlayer(index).PlayerRequest = 0
        TempPlayer(tradeIndex).PlayerRequest = 0
        TempPlayer(index).RequestType = 0
        TempPlayer(tradeIndex).RequestType = 0
        Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: AddAlert index, "You have decline the trade", White
            Case LANG_EN: AddAlert index, "You have decline the trade", White
            Case LANG_ES: AddAlert index, "You have decline the trade", White
        End Select
        Select Case TempPlayer(tradeIndex).CurLanguage
            Case LANG_PT: AddAlert tradeIndex, "The trade was declined", White
            Case LANG_EN: AddAlert tradeIndex, "The trade was declined", White
            Case LANG_ES: AddAlert tradeIndex, "The trade was declined", White
        End Select
        SendCloseTrade tradeIndex
        SendCloseTrade index
        SendRequest tradeIndex
        SendRequest index
    End If
End Sub

Private Sub HandleScanPokedex(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim ScanType As Byte, ScanIndex As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
  
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ScanType = buffer.ReadByte
    ScanIndex = buffer.ReadLong
    Set buffer = Nothing
    
    If ScanType = 1 Then '//Map Npc
        If ScanIndex > 0 And ScanIndex <= MAX_GAME_POKEMON Then
            If MapPokemon(ScanIndex).Num > 0 Then
                If MapPokemon(ScanIndex).Map = Player(index, TempPlayer(index).UseChar).Map Then
                    '//Add Pokedex
                    AddPlayerPokedex index, MapPokemon(ScanIndex).Num, , YES
                End If
            End If
        End If
    Else '// Player
        If ScanIndex > 0 And ScanIndex <= MAX_PLAYER Then
            If PlayerPokemon(ScanIndex).Num > 0 Then
                If Player(ScanIndex, TempPlayer(ScanIndex).UseChar).Map = Player(index, TempPlayer(index).UseChar).Map Then
                    '//Add Pokedex
                    AddPlayerPokedex index, PlayerPokemon(ScanIndex).Num, , YES
                End If
            End If
        End If
    End If
End Sub

Private Sub HandleMOTD(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Text As String

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_CREATOR Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Text = Trim$(buffer.ReadString)
    Set buffer = Nothing
    
    Options.MOTD = Text
    SaveOption
    
    SendPlayerMsg index, "MOTD was changed to: " & Text, White
End Sub

Private Sub HandleCopyMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim destinationMap As Long, sourceMap As Long
Dim i As Long
Dim CurRevision As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_CREATOR Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    destinationMap = buffer.ReadLong
    sourceMap = buffer.ReadLong
    Set buffer = Nothing
    If destinationMap <= 0 Or destinationMap > MAX_MAP Then Exit Sub
    If sourceMap <= 0 Or sourceMap > MAX_MAP Then Exit Sub
    
    CurRevision = Map(destinationMap).Revision
    Call CopyMemory(ByVal VarPtr(Map(destinationMap)), ByVal VarPtr(Map(sourceMap)), LenB(Map(sourceMap)))
    Map(destinationMap).Revision = CurRevision + 1
    
    '//Save the map
    Call SaveMap(destinationMap)
    Call Create_MapCache(destinationMap)
    
    '//Send the clear data first
    Call SendMapNpcData(destinationMap)
    For i = 1 To MAX_MAP_NPC
        SendNpcPokemonData Player(index, TempPlayer(index).UseChar).Map, i, NO, 0, 0, 0, index
    Next
    '//Map Npc
    Call SpawnMapNpcs(destinationMap)
    
    '//Refresh map for everyone online
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If TempPlayer(i).UseChar > 0 Then
                With Player(i, TempPlayer(i).UseChar)
                    If .Map = destinationMap Then
                        Call PlayerWarp(i, destinationMap, .x, .y, .Dir)
                    End If
                End With
            End If
        End If
    Next
    
    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " updated map#" & Player(index, TempPlayer(index).UseChar).Map
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " updated map#" & Player(index, TempPlayer(index).UseChar).Map
    Select Case TempPlayer(index).CurLanguage
        Case LANG_PT: AddAlert index, "Copying map complete", White
        Case LANG_EN: AddAlert index, "Copying map complete", White
        Case LANG_ES: AddAlert index, "Copying map complete", White
    End Select
End Sub

Private Sub HandleGiveItemTo(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim ItemNum As Long, ItemVal As Long
Dim playerName As String
Dim i As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_CREATOR Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    playerName = Trim$(buffer.ReadString)
    ItemNum = buffer.ReadLong
    ItemVal = buffer.ReadLong
    Set buffer = Nothing
    i = FindPlayer(playerName)
    If Not IsPlaying(i) Then Exit Sub
    If TempPlayer(i).UseChar = 0 Then Exit Sub
    If ItemNum <= 0 Or ItemNum > MAX_ITEM Then Exit Sub
    If ItemVal <= 0 Then ItemVal = 1
    
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " , Admin Rights: Give Item To " & Trim$(Player(i, TempPlayer(i).UseChar).Name) & ", Item#" & ItemNum & " x" & ItemVal
    
    If Not GiveItem(i, ItemNum, ItemVal) Then
        '//Error msg
        Select Case TempPlayer(i).CurLanguage
            Case LANG_PT: AddAlert i, "Inventory is full", White
            Case LANG_EN: AddAlert i, "Inventory is full", White
            Case LANG_ES: AddAlert i, "Inventory is full", White
        End Select
        Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: AddAlert index, "Player Inventory is full", White
            Case LANG_EN: AddAlert index, "Player Inventory is full", White
            Case LANG_ES: AddAlert index, "Player Inventory is full", White
        End Select
    Else
        Select Case TempPlayer(i).CurLanguage
            Case LANG_PT: AddAlert i, "You receive a item gift from " & Trim$(Player(index, TempPlayer(index).UseChar).Name), White
            Case LANG_EN: AddAlert i, "You receive a item gift from " & Trim$(Player(index, TempPlayer(index).UseChar).Name), White
            Case LANG_ES: AddAlert i, "You receive a item gift from " & Trim$(Player(index, TempPlayer(index).UseChar).Name), White
        End Select
        Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: AddAlert index, "Player receive the item", White
            Case LANG_EN: AddAlert index, "Player receive the item", White
            Case LANG_ES: AddAlert index, "Player receive the item", White
        End Select
        '//Check if there's still free slot
        If CountFreeInvSlot(i) <= 5 Then
            Select Case TempPlayer(i).CurLanguage
                Case LANG_PT: AddAlert i, "Warning: Your inventory is almost full", White
                Case LANG_EN: AddAlert i, "Warning: Your inventory is almost full", White
                Case LANG_ES: AddAlert i, "Warning: Your inventory is almost full", White
            End Select
        End If
    End If
End Sub

Private Sub HandleGivePokemonTo(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim PokeNum As Long, Level As Long
Dim playerName As String
Dim i As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_CREATOR Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    playerName = Trim$(buffer.ReadString)
    PokeNum = buffer.ReadLong
    Level = buffer.ReadLong
    Set buffer = Nothing
    i = FindPlayer(playerName)
    If Not IsPlaying(i) Then Exit Sub
    If TempPlayer(i).UseChar = 0 Then Exit Sub
    If PokeNum <= 0 Or PokeNum > MAX_POKEMON Then Exit Sub
    If Level <= 0 Or Level > MAX_LEVEL Then Exit Sub
    
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " , Admin Rights: Give Pokemon To " & Trim$(Player(i, TempPlayer(i).UseChar).Name) & ", Pokemon#" & PokeNum & " Level" & Level
    GivePlayerPokemon i, PokeNum, Level, BallEnum.b_Pokeball
End Sub

Private Sub HandleSpawnPokemon(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MapPokeSlot As Long, isShiny As Byte

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_CREATOR Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MapPokeSlot = buffer.ReadLong
    isShiny = buffer.ReadByte
    Set buffer = Nothing
    If MapPokeSlot <= 0 Or MapPokeSlot > MAX_GAME_POKEMON Then Exit Sub
    
    ClearMapPokemon MapPokeSlot
    If isShiny = YES Then
        SpawnMapPokemon MapPokeSlot, True, YES
    Else
        SpawnMapPokemon MapPokeSlot, True
    End If
    TempPlayer(index).MapSwitchTmr = NO
End Sub

Private Sub HandleSetLanguage(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    TempPlayer(index).CurLanguage = buffer.ReadByte
    Set buffer = Nothing
End Sub

Private Sub HandleBuyStorageSlot(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim StorageType As Byte, storageSlot As Byte
Dim Amount As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
  
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    StorageType = buffer.ReadByte
    storageSlot = buffer.ReadByte
    Set buffer = Nothing
    
    If storageSlot < 0 Or storageSlot > MAX_STORAGE_SLOT Then Exit Sub
    Amount = 100000 * (storageSlot - 2)
    If Player(index, TempPlayer(index).UseChar).Money >= Amount Then
        Select Case StorageType
            Case 1 '// Item
                If PlayerInvStorage(index).slot(storageSlot).Unlocked = NO Then
                    PlayerInvStorage(index).slot(storageSlot).Unlocked = YES
                    Player(index, TempPlayer(index).UseChar).Money = Player(index, TempPlayer(index).UseChar).Money - Amount
                    SendPlayerInvStorage index
                    SendPlayerData index
                    AddAlert index, "New Item Storage slot has been unlocked", White
                End If
            Case 2 '// Pokemon
                If PlayerPokemonStorage(index).slot(storageSlot).Unlocked = NO Then
                    PlayerPokemonStorage(index).slot(storageSlot).Unlocked = YES
                    Player(index, TempPlayer(index).UseChar).Money = Player(index, TempPlayer(index).UseChar).Money - Amount
                    SendPlayerPokemonStorage index
                    SendPlayerData index
                    AddAlert index, "New Pokemon Storage slot has been unlocked", White
                End If
        End Select
    Else
        AddAlert index, "Not enough money", White
    End If
End Sub

Private Sub HandleSellPokeStorageSlot(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim storageSlot As Byte, StorageData As Byte

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
  
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    storageSlot = buffer.ReadByte
    StorageData = buffer.ReadByte
    Set buffer = Nothing
End Sub

Private Sub HandleChangeShinyRate(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Rate As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Rate = buffer.ReadLong
    Set buffer = Nothing
    
    SendPlayerMsg index, "Shiny Rate was changed to: " & Rate, White
End Sub

Private Sub HandleRelearnMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MoveSlot As Byte, PokeSlot As Byte, PokeNum As Long
Dim MoveNum As Long, oSlot As Byte
Dim x As Byte
Dim invSlot As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If PlayerPokemon(index).Num <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MoveSlot = buffer.ReadByte
    PokeSlot = buffer.ReadByte
    PokeNum = buffer.ReadLong
    Set buffer = Nothing
    
    If Not PlayerPokemon(index).Num = PokeNum Then Exit Sub
    If Not PlayerPokemon(index).slot = PokeSlot Then Exit Sub
    If MoveSlot <= 0 Or MoveSlot > MAX_POKEMON_MOVESET Then Exit Sub
    
    If PokeNum > 0 Then
        MoveNum = Pokemon(PokeNum).Moveset(MoveSlot).MoveNum
        If MoveNum > 0 Then
            For x = 1 To MAX_MOVESET
                If MoveNum = PlayerPokemons(index).Data(PokeSlot).Moveset(x).Num Then
                    Exit Sub
                End If
            Next
            If PlayerPokemons(index).Data(PokeSlot).Level < Pokemon(PokeNum).Moveset(MoveSlot).MoveLevel Then
                Exit Sub
            End If
            invSlot = FindInvItemSlot(index, 72)
            '//Check if have the required item
            If invSlot > 0 Then
                '//Take Item
                PlayerInv(index).Data(invSlot).Value = PlayerInv(index).Data(invSlot).Value - 1
                If PlayerInv(index).Data(invSlot).Value <= 0 Then
                    '//Clear Item
                    PlayerInv(index).Data(invSlot).Num = 0
                    PlayerInv(index).Data(invSlot).Value = 0
                End If
                SendPlayerInvSlot index, invSlot
            
                '//Continue
                oSlot = FindFreeMoveSlot(index, PokeSlot)
                If oSlot > 0 Then
                    PlayerPokemons(index).Data(PokeSlot).Moveset(oSlot).Num = MoveNum
                    PlayerPokemons(index).Data(PokeSlot).Moveset(oSlot).TotalPP = PokemonMove(MoveNum).PP
                    PlayerPokemons(index).Data(PokeSlot).Moveset(oSlot).CurPP = PlayerPokemons(index).Data(PokeSlot).Moveset(MoveSlot).TotalPP
                    SendPlayerPokemonSlot index, PokeSlot
                    '//Send Msg
                    SendPlayerMsg index, Trim$(Pokemon(PokeNum).Name) & " learned the move " & Trim$(PokemonMove(MoveNum).Name), White
                Else
                    '//Proceed to ask
                    TempPlayer(index).MoveLearnPokeSlot = PokeSlot
                    TempPlayer(index).MoveLearnNum = MoveNum
                    TempPlayer(index).MoveLearnIndex = 0
                    SendNewMove index
                End If
            Else
                Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "Sorry, you don't have a " & Trim$(Item(72).Name), White
                    Case LANG_EN: AddAlert index, "Sorry, you don't have a " & Trim$(Item(72).Name), White
                    Case LANG_ES: AddAlert index, "Sorry, you don't have a " & Trim$(Item(72).Name), White
                End Select
            End If
        End If
    End If
End Sub

Private Sub HandleUseRevive(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim IsMaxRev As Byte, ReviveItemNum As Long
Dim invSlot As Long, PokeSlot As Byte

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PokeSlot = buffer.ReadByte
    IsMaxRev = buffer.ReadByte
    Set buffer = Nothing
    
    If PokeSlot <= 0 Or PokeSlot > MAX_PLAYER_POKEMON Then Exit Sub
    If PlayerPokemons(index).Data(PokeSlot).Num <= 0 Then Exit Sub
    If PlayerPokemons(index).Data(PokeSlot).CurHP > 0 Then Exit Sub
    If TempPlayer(index).InDuel > 0 Then Exit Sub
    If TempPlayer(index).InNpcDuel > 0 Then Exit Sub
    
    If IsMaxRev = YES Then
        ReviveItemNum = 48
        invSlot = FindInvItemSlot(index, ReviveItemNum)
        If invSlot > 0 Then
            '//Take Item
            PlayerInv(index).Data(invSlot).Value = PlayerInv(index).Data(invSlot).Value - 1
            If PlayerInv(index).Data(invSlot).Value <= 0 Then
                '//Clear Item
                PlayerInv(index).Data(invSlot).Num = 0
                PlayerInv(index).Data(invSlot).Value = 0
            End If
            SendPlayerInvSlot index, invSlot
            
            PlayerPokemons(index).Data(PokeSlot).CurHP = PlayerPokemons(index).Data(PokeSlot).MaxHP
            SendPlayerPokemonSlot index, PokeSlot
            
            Select Case TempPlayer(index).CurLanguage
                Case LANG_PT: AddAlert index, "Pokemon was revived", White
                Case LANG_EN: AddAlert index, "Pokemon was revived", White
                Case LANG_ES: AddAlert index, "Pokemon was revived", White
            End Select
        Else
            Select Case TempPlayer(index).CurLanguage
                Case LANG_PT: AddAlert index, "Sorry, you don't have a " & Trim$(Item(ReviveItemNum).Name), White
                Case LANG_EN: AddAlert index, "Sorry, you don't have a " & Trim$(Item(ReviveItemNum).Name), White
                Case LANG_ES: AddAlert index, "Sorry, you don't have a " & Trim$(Item(ReviveItemNum).Name), White
            End Select
        End If
    Else
        ReviveItemNum = 62
        invSlot = FindInvItemSlot(index, ReviveItemNum)
        If invSlot > 0 Then
            '//Take Item
            PlayerInv(index).Data(invSlot).Value = PlayerInv(index).Data(invSlot).Value - 1
            If PlayerInv(index).Data(invSlot).Value <= 0 Then
                '//Clear Item
                PlayerInv(index).Data(invSlot).Num = 0
                PlayerInv(index).Data(invSlot).Value = 0
            End If
            SendPlayerInvSlot index, invSlot
            
            PlayerPokemons(index).Data(PokeSlot).CurHP = PlayerPokemons(index).Data(PokeSlot).MaxHP / 2
            SendPlayerPokemonSlot index, PokeSlot
            
            Select Case TempPlayer(index).CurLanguage
                Case LANG_PT: AddAlert index, "Pokemon was revived", White
                Case LANG_EN: AddAlert index, "Pokemon was revived", White
                Case LANG_ES: AddAlert index, "Pokemon was revived", White
            End Select
        Else
            Select Case TempPlayer(index).CurLanguage
                Case LANG_PT: AddAlert index, "Sorry, you don't have a " & Trim$(Item(ReviveItemNum).Name), White
                Case LANG_EN: AddAlert index, "Sorry, you don't have a " & Trim$(Item(ReviveItemNum).Name), White
                Case LANG_ES: AddAlert index, "Sorry, you don't have a " & Trim$(Item(ReviveItemNum).Name), White
            End Select
        End If
    End If
End Sub

Private Sub HandleAddHeld(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim invSlot As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    invSlot = buffer.ReadByte
    Set buffer = Nothing
    
    If invSlot <= 0 Or invSlot > MAX_PLAYER_INV Then Exit Sub
    If PlayerInv(index).Data(invSlot).Num <= 0 Then Exit Sub
    If PlayerInv(index).Data(invSlot).Value < 1 Then Exit Sub
    
    If PlayerPokemon(index).Num <= 0 Then
        Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: AddAlert index, "Spawn your pokemon first", White
            Case LANG_EN: AddAlert index, "Spawn your pokemon first", White
            Case LANG_ES: AddAlert index, "Spawn your pokemon first", White
        End Select
        Exit Sub
    End If
    If PlayerPokemon(index).slot <= 0 Then
        Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: AddAlert index, "Spawn your pokemon first", White
            Case LANG_EN: AddAlert index, "Spawn your pokemon first", White
            Case LANG_ES: AddAlert index, "Spawn your pokemon first", White
        End Select
        Exit Sub
    End If
    
    ' check previous held item
    If PlayerPokemons(index).Data(PlayerPokemon(index).slot).HeldItem > 0 Then
        Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: AddAlert index, "Your pokemon is currently holding an item", White
            Case LANG_EN: AddAlert index, "Your pokemon is currently holding an item", White
            Case LANG_ES: AddAlert index, "Your pokemon is currently holding an item", White
        End Select
        Exit Sub
    End If
    
    ' Complete Process
    ' Give item
    PlayerPokemons(index).Data(PlayerPokemon(index).slot).HeldItem = PlayerInv(index).Data(invSlot).Num
    SendPlayerPokemonSlot index, PlayerPokemon(index).slot
    
    '//Take Item
    PlayerInv(index).Data(invSlot).Value = PlayerInv(index).Data(invSlot).Value - 1
    If PlayerInv(index).Data(invSlot).Value <= 0 Then
        '//Clear Item
        PlayerInv(index).Data(invSlot).Num = 0
        PlayerInv(index).Data(invSlot).Value = 0
    End If
    SendPlayerInvSlot index, invSlot
End Sub

Private Sub HandleRemoveHeld(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim PokeSlot As Long
Dim ItemNum As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PokeSlot = buffer.ReadByte
    Set buffer = Nothing
    
    If PokeSlot <= 0 Or PokeSlot > MAX_PLAYER_POKEMON Then Exit Sub
    If PlayerPokemons(index).Data(PokeSlot).Num <= 0 Then Exit Sub
    If PlayerPokemons(index).Data(PokeSlot).HeldItem <= 0 Then Exit Sub
    
    ItemNum = PlayerPokemons(index).Data(PokeSlot).HeldItem
    
    If Not GiveItem(index, ItemNum, 1) Then
        '//Error msg
        Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: AddAlert index, "Inventory is full", White
            Case LANG_EN: AddAlert index, "Inventory is full", White
            Case LANG_ES: AddAlert index, "Inventory is full", White
        End Select
    Else
        PlayerPokemons(index).Data(PokeSlot).HeldItem = 0
        SendPlayerPokemonSlot index, PokeSlot
    End If
End Sub

Private Sub HandleStealthMode(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    
    If Player(index, TempPlayer(index).UseChar).StealthMode = YES Then
        Player(index, TempPlayer(index).UseChar).StealthMode = NO
    Else
        Player(index, TempPlayer(index).UseChar).StealthMode = YES
    End If
    SendPlayerData index
End Sub

Private Sub HandleWhosOnline(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendWhosOnline index
End Sub

Private Sub HandleRequestRank(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendRankTo index
    
    UpdateRank Trim$(Player(index, TempPlayer(index).UseChar).Name), Player(index, TempPlayer(index).UseChar).Level, Player(index, TempPlayer(index).UseChar).CurExp
End Sub

Private Sub HandleHotbarUpdate(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim HotbarSlot As Byte, invSlot As Byte

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    HotbarSlot = buffer.ReadByte
    invSlot = buffer.ReadByte
    Set buffer = Nothing
    
    If HotbarSlot <= 0 Or HotbarSlot > MAX_HOTBAR Then Exit Sub
    
    With Player(index, TempPlayer(index).UseChar)
        If invSlot > 0 And invSlot <= MAX_PLAYER_INV Then
            If PlayerInv(index).Data(invSlot).Num > 0 Then
                .Hotbar(HotbarSlot) = PlayerInv(index).Data(invSlot).Num
            Else
                .Hotbar(HotbarSlot) = 0
            End If
        Else
            .Hotbar(HotbarSlot) = 0
        End If
    End With
    SendPlayerData index
End Sub

Private Sub HandleUseHotbar(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim HotbarSlot As Byte, invSlot As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    HotbarSlot = buffer.ReadByte
    Set buffer = Nothing
    
    If HotbarSlot <= 0 Or HotbarSlot > MAX_HOTBAR Then Exit Sub
 
    With Player(index, TempPlayer(index).UseChar)
        If .Hotbar(HotbarSlot) > 0 Then
            invSlot = checkItem(index, .Hotbar(HotbarSlot))
            
            If invSlot > 0 Then
                '//Use Item
                PlayerUseItem index, invSlot
            End If
        End If
    End With
End Sub

Private Sub HandleCreateParty(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    CreateParty index
End Sub

Private Sub HandleLeaveParty(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    LeaveParty index
End Sub

'//Editors
Private Sub HandleRequestEditMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_MAPPER Then Exit Sub
    
    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested to edit map#" & Player(index, TempPlayer(index).UseChar).Map
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested to edit map#" & Player(index, TempPlayer(index).UseChar).Map

    Set buffer = New clsBuffer
    buffer.WriteLong SInitMap
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Private Sub HandleMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MapNum As Long
Dim x As Long, y As Long
Dim i As Long, a As Byte

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_MAPPER Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MapNum = Player(index, TempPlayer(index).UseChar).Map
    Call ClearMap(MapNum)
    
    With Map(MapNum)
        '//General
        .Revision = buffer.ReadLong
        .Name = Trim$(buffer.ReadString)
        .Moral = buffer.ReadByte
        
        '//Size
        .MaxX = buffer.ReadLong
        .MaxY = buffer.ReadLong
        If .MaxX < MAX_MAPX Then .MaxX = MAX_MAPX
        If .MaxY < MAX_MAPY Then .MaxY = MAX_MAPY
        
        '//Redim size
        ReDim Map(MapNum).Tile(0 To .MaxX, 0 To .MaxY)
    End With
    
    '//Tiles
    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY
            With Map(MapNum).Tile(x, y)
                '//Layer
                For i = MapLayer.Ground To MapLayer.MapLayer_Count - 1
                    For a = MapLayerType.Normal To MapLayerType.Animated
                        .Layer(i, a).Tile = buffer.ReadLong
                        .Layer(i, a).TileX = buffer.ReadLong
                        .Layer(i, a).TileY = buffer.ReadLong
                        '//Map Anim
                        .Layer(i, a).MapAnim = buffer.ReadLong
                    Next
                Next
                '//Tile Data
                .Attribute = buffer.ReadByte
                .Data1 = buffer.ReadLong
                .Data2 = buffer.ReadLong
                .Data3 = buffer.ReadLong
                .Data4 = buffer.ReadLong
            End With
        Next
    Next
    
    With Map(MapNum)
        '//Map Link
        .LinkUp = buffer.ReadLong
        .LinkDown = buffer.ReadLong
        .LinkLeft = buffer.ReadLong
        .LinkRight = buffer.ReadLong
        
        '//Map Data
        .Music = Trim$(buffer.ReadString)
        
        '//Npc
        For i = 1 To MAX_MAP_NPC
            .Npc(i) = buffer.ReadLong
            ClearMapNpc MapNum, i
        Next
        
        '//Moral
        .Sheltered = buffer.ReadByte
        .IsCave = buffer.ReadByte
        .CaveLight = buffer.ReadByte
        .SpriteType = buffer.ReadByte
        .StartWeather = buffer.ReadByte
    End With
    Set buffer = Nothing
    
    '//Save the map
    Call SaveMap(MapNum)
    Call Create_MapCache(MapNum)
    
    '//Send the clear data first
    Call SendMapNpcData(MapNum)
    For i = 1 To MAX_MAP_NPC
        SendNpcPokemonData Player(index, TempPlayer(index).UseChar).Map, i, NO, 0, 0, 0, index
    Next
    '//Map Npc
    Call SpawnMapNpcs(MapNum)
    
    '//Refresh map for everyone online
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If TempPlayer(i).UseChar > 0 Then
                With Player(i, TempPlayer(i).UseChar)
                    If .Map = MapNum Then
                        Call PlayerWarp(i, MapNum, .x, .y, .Dir)
                    End If
                End With
            End If
        End If
    Next
    
    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " updated map#" & Player(index, TempPlayer(index).UseChar).Map
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " updated map#" & Player(index, TempPlayer(index).UseChar).Map
End Sub

Private Sub HandleRequestEditNpc(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested to edit npc"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested to edit npc"

    Set buffer = New clsBuffer
    buffer.WriteLong SInitNpc
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Private Sub HandleRequestNpc(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested npc data"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested npc data"
    
    SendNpcs index
End Sub

Private Sub HandleSaveNpc(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim xIndex As Long
Dim dSize As Long
Dim dData() As Byte

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    xIndex = buffer.ReadLong
    If xIndex < 0 Or xIndex > MAX_NPC Then Exit Sub
    dSize = LenB(Npc(xIndex))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(Npc(xIndex)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
    
    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " save npc#" & xIndex & " data"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " save npc#" & xIndex & " data"
    
    Call SendUpdateNpcToAll(xIndex)
    Call SaveNpc(xIndex)
End Sub

Private Sub HandleRequestEditPokemon(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested to edit Pokemon"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested to edit Pokemon"

    Set buffer = New clsBuffer
    buffer.WriteLong SInitPokemon
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Private Sub HandleRequestPokemon(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested Pokemon data"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested Pokemon data"
    
    SendPokemons index
End Sub

Private Sub HandleSavePokemon(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim xIndex As Long
Dim dSize As Long
Dim dData() As Byte

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    xIndex = buffer.ReadLong
    If xIndex < 0 Or xIndex > MAX_POKEMON Then Exit Sub
    dSize = LenB(Pokemon(xIndex))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(Pokemon(xIndex)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
    
    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " save pokemon#" & xIndex & " data"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " save pokemon#" & xIndex & " data"
    
    Call SendUpdatePokemonToAll(xIndex)
    Call SavePokemon(xIndex)
End Sub

Private Sub HandleRequestEditItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested to edit Item"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested to edit Item"

    Set buffer = New clsBuffer
    buffer.WriteLong SInitItem
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Private Sub HandleRequestItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested Item data"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested Item data"
    
    SendItems index
End Sub

Private Sub HandleSaveItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim xIndex As Long
Dim dSize As Long
Dim dData() As Byte

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    xIndex = buffer.ReadLong
    If xIndex < 0 Or xIndex > MAX_ITEM Then Exit Sub
    dSize = LenB(Item(xIndex))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(Item(xIndex)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
    
    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " save item#" & xIndex & " data"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " save item#" & xIndex & " data"
    
    Call SendUpdateItemToAll(xIndex)
    Call SaveItem(xIndex)
End Sub

Private Sub HandleRequestEditPokemonMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested to edit Pokemon Move"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested to edit Pokemon Move"

    Set buffer = New clsBuffer
    buffer.WriteLong SInitPokemonMove
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Private Sub HandleRequestPokemonMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested Pokemon Move data"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested Pokemon Move data"
    
    SendPokemonMoves index
End Sub

Private Sub HandleSavePokemonMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim xIndex As Long
Dim dSize As Long
Dim dData() As Byte

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    xIndex = buffer.ReadLong
    If xIndex < 0 Or xIndex > MAX_POKEMON_MOVE Then Exit Sub
    dSize = LenB(PokemonMove(xIndex))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(PokemonMove(xIndex)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
    
    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " save pokemonmove#" & xIndex & " data"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " save pokemonmove#" & xIndex & " data"
    
    Call SendUpdatePokemonMoveToAll(xIndex)
    Call SavePokemonMove(xIndex)
End Sub

Private Sub HandleRequestEditAnimation(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested to edit Animation"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested to edit Animation"

    Set buffer = New clsBuffer
    buffer.WriteLong SInitAnimation
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Private Sub HandleRequestAnimation(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested Animation data"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested Animation data"
    
    SendAnimations index
End Sub

Private Sub HandleSaveAnimation(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim xIndex As Long
Dim dSize As Long
Dim dData() As Byte

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    xIndex = buffer.ReadLong
    If xIndex < 0 Or xIndex > MAX_ANIMATION Then Exit Sub
    dSize = LenB(Animation(xIndex))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(Animation(xIndex)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
    
    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " save animation#" & xIndex & " data"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " save animation#" & xIndex & " data"
    
    Call SendUpdateAnimationToAll(xIndex)
    Call SaveAnimation(xIndex)
End Sub

Private Sub HandleRequestEditSpawn(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested to edit Spawn"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested to edit Spawn"

    Set buffer = New clsBuffer
    buffer.WriteLong SInitSpawn
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Private Sub HandleRequestSpawn(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested Spawn data"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested Spawn data"
    
    SendSpawns index
End Sub

Private Sub HandleSaveSpawn(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim xIndex As Long
Dim dSize As Long
Dim dData() As Byte

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    xIndex = buffer.ReadLong
    If xIndex < 0 Or xIndex > MAX_GAME_POKEMON Then Exit Sub
    dSize = LenB(Spawn(xIndex))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(Spawn(xIndex)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
    
    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " save Spawn#" & xIndex & " data"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " save Spawn#" & xIndex & " data"
    
    Call SendUpdateSpawnToAll(xIndex)
    Call SaveSpawn(xIndex)
    
    '//Update Data
    ClearMapPokemon xIndex, True
    MapPokemon(xIndex).PokemonIndex = Spawn(xIndex).PokeNum
    MapPokemon(xIndex).Respawn = GetTickCount
    '//Update HighIndex
    If Spawn(xIndex).PokeNum > 0 Then
        If xIndex > Pokemon_HighIndex Then
            Pokemon_HighIndex = xIndex
            SendPokemonHighIndex
        End If
    End If
End Sub

Private Sub HandleRequestEditConversation(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested to edit Conversation"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested to edit Conversation"

    Set buffer = New clsBuffer
    buffer.WriteLong SInitConversation
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Private Sub HandleRequestConversation(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested Conversation data"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested Conversation data"
    
    SendConversations index
End Sub

Private Sub HandleSaveConversation(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim xIndex As Long
Dim dSize As Long
Dim dData() As Byte

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    xIndex = buffer.ReadLong
    If xIndex < 0 Or xIndex > MAX_GAME_POKEMON Then Exit Sub
    dSize = LenB(Conversation(xIndex))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(Conversation(xIndex)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
    
    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " save Conversation#" & xIndex & " data"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " save Conversation#" & xIndex & " data"
    
    Call SendUpdateConversationToAll(xIndex)
    Call SaveConversation(xIndex)
End Sub

Private Sub HandleRequestEditShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested to edit Shop"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested to edit Shop"

    Set buffer = New clsBuffer
    buffer.WriteLong SInitShop
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Private Sub HandleRequestShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested Shop data"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested Shop data"
    
    SendShops index
End Sub

Private Sub HandleSaveShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim xIndex As Long
Dim dSize As Long
Dim dData() As Byte

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    xIndex = buffer.ReadLong
    If xIndex < 0 Or xIndex > MAX_GAME_POKEMON Then Exit Sub
    dSize = LenB(Shop(xIndex))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(Shop(xIndex)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
    
    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " save Shop#" & xIndex & " data"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " save Shop#" & xIndex & " data"
    
    Call SendUpdateShopToAll(xIndex)
    Call SaveShop(xIndex)
End Sub

Private Sub HandleRequestEditQuest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested to edit Quest"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested to edit Quest"

    Set buffer = New clsBuffer
    buffer.WriteLong SInitQuest
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Private Sub HandleRequestQuest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested Quest data"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " requested Quest data"
    
    SendQuests index
End Sub

Private Sub HandleSaveQuest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim xIndex As Long
Dim dSize As Long
Dim dData() As Byte

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    xIndex = buffer.ReadLong
    If xIndex < 0 Or xIndex > MAX_GAME_POKEMON Then Exit Sub
    dSize = LenB(Quest(xIndex))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(Quest(xIndex)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
    
    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " save Quest#" & xIndex & " data"
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " save Quest#" & xIndex & " data"
    
    Call SendUpdateQuestToAll(xIndex)
    Call SaveQuest(xIndex)
End Sub

Private Sub HandleKickPlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim TargetIndex As Long
Dim InputName As String

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    InputName = buffer.ReadString
    Set buffer = Nothing
    
    TargetIndex = FindPlayer(InputName)
    
    If TargetIndex <= 0 Or TargetIndex > MAX_PLAYER Then Exit Sub
    If Not IsPlaying(TargetIndex) Then Exit Sub
    If TempPlayer(TargetIndex).UseChar <= 0 Then Exit Sub
    If Player(index, TempPlayer(index).UseChar).Access <= Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Access Then Exit Sub
    
    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " kick " & Trim$(Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Name)
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " kick " & Trim$(Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Name)
    
    CloseSocket TargetIndex
End Sub

Private Sub HandleBanPlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim TargetIndex As Long
Dim InputName As String

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    InputName = buffer.ReadString
    Set buffer = Nothing
    
    TargetIndex = FindPlayer(InputName)
    
    If TargetIndex <= 0 Or TargetIndex > MAX_PLAYER Then Exit Sub
    If Not IsPlaying(TargetIndex) Then Exit Sub
    If TempPlayer(TargetIndex).UseChar <= 0 Then Exit Sub
    If Player(index, TempPlayer(index).UseChar).Access <= Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Access Then Exit Sub
    
    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " ban " & Trim$(Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Name)
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " ban " & Trim$(Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Name)
    
    'BanIP GetPlayerIP(TargetIndex)
    BanCharacter Trim$(Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Name)
    CloseSocket TargetIndex
End Sub

Private Sub HandleMutePlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim TargetIndex As Long
Dim InputName As String

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    InputName = buffer.ReadString
    Set buffer = Nothing
    
    TargetIndex = FindPlayer(InputName)
    
    If TargetIndex <= 0 Or TargetIndex > MAX_PLAYER Then Exit Sub
    If Not IsPlaying(TargetIndex) Then Exit Sub
    If TempPlayer(TargetIndex).UseChar <= 0 Then Exit Sub
    If Player(index, TempPlayer(index).UseChar).Access <= Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Access Then Exit Sub
    
    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " ban " & Trim$(Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Name)
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " ban " & Trim$(Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Name)
    
    Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Muted = YES
End Sub

Private Sub HandleUnmutePlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim TargetIndex As Long
Dim InputName As String

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    InputName = buffer.ReadString
    Set buffer = Nothing
    
    TargetIndex = FindPlayer(InputName)
    
    If TargetIndex <= 0 Or TargetIndex > MAX_PLAYER Then Exit Sub
    If Not IsPlaying(TargetIndex) Then Exit Sub
    If TempPlayer(TargetIndex).UseChar <= 0 Then Exit Sub
    If Player(index, TempPlayer(index).UseChar).Access <= Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Access Then Exit Sub
    
    TextAdd frmServer.txtLog, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " ban " & Trim$(Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Name)
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " ban " & Trim$(Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Name)
    
    Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Muted = NO
End Sub

Private Sub HandleFlyToBadge(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim TargetIndex As Byte

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    TargetIndex = buffer.ReadByte
    Set buffer = Nothing
End Sub
