Attribute VB_Name = "modPlayer"
Option Explicit

' **********************
' ** Player Functions **
' **********************
Public Function GetPlayerIP(ByVal index As Long) As String
    If index <= 0 Or index > MAX_PLAYER Then Exit Function
    GetPlayerIP = frmServer.Socket(index).RemoteHostIP
End Function

Private Function TotalPlayerOnMap(ByVal MapNum As Long) As Long
    Dim i As Long
    Dim count As Long

    count = 0
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If TempPlayer(i).UseChar > 0 Then
                If Player(i, TempPlayer(i).UseChar).Map = MapNum Then
                    count = count + 1
                End If
            End If
        End If
    Next
    TotalPlayerOnMap = count
End Function

Public Function TotalPlayerOnline()
    Dim i As Long
    Dim count As Long

    count = 0
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If TempPlayer(i).UseChar > 0 Then
                count = count + 1
            End If
        End If
    Next
    TotalPlayerOnline = count
End Function

Public Sub PlayerWarp(ByVal index As Long, ByVal MapNum As Long, ByVal x As Long, ByVal Y As Long, ByVal Dir As Byte)
    Dim OldMap As Long

    '//Exit out when error
    If index <= 0 Or index > MAX_PLAYER Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Or TempPlayer(index).UseChar > MAX_PLAYERCHAR Then Exit Sub
    If MapNum <= 0 Or MapNum > MAX_MAP Then Exit Sub

    '//Correct error position
    If x <= 0 Then x = 0
    If x > Map(MapNum).MaxX Then x = Map(MapNum).MaxX
    If Y <= 0 Then Y = 0
    If Y > Map(MapNum).MaxY Then Y = Map(MapNum).MaxY

    OldMap = Player(index, TempPlayer(index).UseChar).Map

    '//Update position
    With Player(index, TempPlayer(index).UseChar)
        .Map = MapNum
        .x = x
        .Y = Y
        .Dir = Dir
    End With

    '//If map did not match
    If Not OldMap = MapNum Then
        '//Clear player data on old map
        SendLeaveMap index, OldMap

        '//Clear Target
        ClearMyTarget index, OldMap

        '//Check if there's still remaining player on map
        If TotalPlayerOnMap(OldMap) <= 0 Then
            PlayerOnMap(OldMap) = NO
            Map(OldMap).CurWeather = Map(OldMap).StartWeather
        End If

        TempPlayer(index).MapSwitchTmr = YES
    End If

    '//Add log
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " has been warped on Map#" & MapNum & " x:" & x & " y:" & Y

    '//Update
    PlayerOnMap(MapNum) = YES
    TempPlayer(index).GettingMap = True
    SendCheckForMap index, MapNum
End Sub

Public Sub ForcePlayerMove(ByVal index As Long, ByVal Dir As Byte)
'//Exit out when error
    If index <= 0 Or index > MAX_PLAYER Then Exit Sub
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Or TempPlayer(index).UseChar > MAX_PLAYERCHAR Then Exit Sub
    If Dir < 0 Or Dir > DIR_RIGHT Then Exit Sub

    Select Case Dir
    Case DIR_UP
        If Player(index, TempPlayer(index).UseChar).Y = 0 Then Exit Sub
    Case DIR_LEFT
        If Player(index, TempPlayer(index).UseChar).x = 0 Then Exit Sub
    Case DIR_DOWN
        If Player(index, TempPlayer(index).UseChar).Y = Map(Player(index, TempPlayer(index).UseChar).Map).MaxY Then Exit Sub
    Case DIR_RIGHT
        If Player(index, TempPlayer(index).UseChar).x = Map(Player(index, TempPlayer(index).UseChar).Map).MaxX Then Exit Sub
    End Select

    PlayerMove index, Dir, True
End Sub

Public Sub PlayerMove(ByVal index As Long, ByVal Dir As Byte, Optional ByVal sendToSelf As Boolean = False)
    Dim DidMove As Boolean
    Dim OldX As Long, OldY As Long
    Dim gothealed As Boolean
    Dim i As Long, x As Byte

    '//Exit out when error
    If index <= 0 Or index > MAX_PLAYER Then Exit Sub
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Or TempPlayer(index).UseChar > MAX_PLAYERCHAR Then Exit Sub
    If Dir < 0 Or Dir > DIR_RIGHT Then Exit Sub

    DidMove = False

    With Player(index, TempPlayer(index).UseChar)
        '//Store original location in case it got desync
        OldX = .x
        OldY = .Y

        Select Case Dir
        Case DIR_UP
            .Dir = DIR_UP

            '//Check to make sure not outside of boundries
            If .Y > 0 Then
                If Not CheckDirection(.Map, DIR_UP, .x, .Y) Then
                    .Y = .Y - 1
                    DidMove = True
                End If
            Else
                '//Check Link
                If Map(.Map).LinkUp > 0 Then
                    PlayerWarp index, Map(.Map).LinkUp, .x, Map(Map(.Map).LinkUp).MaxY, .Dir
                    Exit Sub
                End If
            End If
        Case DIR_DOWN
            .Dir = DIR_DOWN

            '//Check to make sure not outside of boundries
            If .Y < Map(.Map).MaxY Then
                If Not CheckDirection(.Map, DIR_DOWN, .x, .Y) Then
                    .Y = .Y + 1
                    DidMove = True
                End If
            Else
                '//Check Link
                If Map(.Map).LinkDown > 0 Then
                    PlayerWarp index, Map(.Map).LinkDown, .x, 0, .Dir
                    Exit Sub
                End If
            End If
        Case DIR_LEFT
            .Dir = DIR_LEFT

            '//Check to make sure not outside of boundries
            If .x > 0 Then
                If Not CheckDirection(.Map, DIR_LEFT, .x, .Y) Then
                    .x = .x - 1
                    DidMove = True
                End If
            Else
                '//Check Link
                If Map(.Map).LinkLeft > 0 Then
                    PlayerWarp index, Map(.Map).LinkLeft, Map(Map(.Map).LinkLeft).MaxX, .Y, .Dir
                    Exit Sub
                End If
            End If
        Case DIR_RIGHT
            .Dir = DIR_RIGHT

            '//Check to make sure not outside of boundries
            If .x < Map(.Map).MaxX Then
                If Not CheckDirection(.Map, DIR_RIGHT, .x, .Y) Then
                    .x = .x + 1
                    DidMove = True
                End If
            Else
                '//Check Link
                If Map(.Map).LinkRight > 0 Then
                    PlayerWarp index, Map(.Map).LinkRight, 0, .Y, .Dir
                    Exit Sub
                End If
            End If
        End Select

        '//Got Desynced
        If Not DidMove Then
            .x = OldX
            .Y = OldY
            SendPlayerXY index
            SendPlayerXY index, True

            'If .Action <> 0 Then
            '    .Action = 0
            '    SendPlayerAction Index
            'End If
        Else

            '//Fish System
            If GetPlayerFishMode(index) = YES Then
                SetPlayerFishMode index, NO
                SetPlayerFishRod index, 0
                SendActionMsg GetPlayerMap(index), "Fish Down!", Player(index, TempPlayer(index).UseChar).x * 32, Player(index, TempPlayer(index).UseChar).Y * 32, BrightRed
                SendFishMode index
            End If

            TempPlayer(index).MapSwitchTmr = NO

            SendPlayerMove index, sendToSelf

            '//Check tile attribute
            Select Case Map(.Map).Tile(.x, .Y).Attribute
            Case MapAttribute.Warp
                '//Warp
                If Map(.Map).Tile(.x, .Y).Data1 > 0 Then
                    PlayerWarp index, Map(.Map).Tile(.x, .Y).Data1, Map(.Map).Tile(.x, .Y).Data2, Map(.Map).Tile(.x, .Y).Data3, Map(.Map).Tile(.x, .Y).Data4
                End If
            Case MapAttribute.Slide
                ' Slide
                'If .Action = 0 Then
                '    .Action = ACTION_SLIDE
                '    SendPlayerAction Index
                '    .ActionTmr = GetTickCount + 100
                'End If
            Case MapAttribute.HealPokemon
                '//Heal Pokemon
                gothealed = False
                For i = 1 To MAX_PLAYER_POKEMON
                    If PlayerPokemons(index).Data(i).Num > 0 Then
                        If PlayerPokemons(index).Data(i).CurHp < PlayerPokemons(index).Data(i).MaxHp Then
                            PlayerPokemons(index).Data(i).CurHp = PlayerPokemons(index).Data(i).MaxHp
                            gothealed = True
                        End If
                        If PlayerPokemons(index).Data(i).Status > 0 Then
                            PlayerPokemons(index).Data(i).Status = 0
                            gothealed = True
                        End If
                        For x = 1 To MAX_MOVESET
                            If PlayerPokemons(index).Data(i).Moveset(x).Num > 0 Then
                                If PlayerPokemons(index).Data(i).Moveset(x).CurPP < PlayerPokemons(index).Data(i).Moveset(x).TotalPP Then
                                    PlayerPokemons(index).Data(i).Moveset(x).CurPP = PlayerPokemons(index).Data(i).Moveset(x).TotalPP
                                    PlayerPokemons(index).Data(i).Moveset(x).CD = 0
                                    gothealed = True
                                End If
                            End If
                        Next
                    End If
                Next
                If Player(index, TempPlayer(index).UseChar).CurHp < GetPlayerHP(Player(index, TempPlayer(index).UseChar).Level) Then
                    Player(index, TempPlayer(index).UseChar).CurHp = GetPlayerHP(Player(index, TempPlayer(index).UseChar).Level)
                    gothealed = True
                End If
                If Player(index, TempPlayer(index).UseChar).Status > 0 Then
                    Player(index, TempPlayer(index).UseChar).Status = 0
                    Player(index, TempPlayer(index).UseChar).IsConfuse = False
                    gothealed = True
                End If
                If gothealed Then
                    Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "Pokemon HP and PP restored", White
                    Case LANG_EN: AddAlert index, "Pokemon HP and PP restored", White
                    Case LANG_ES: AddAlert index, "Pokemon HP and PP restored", White
                    End Select
                    SendPlayerPokemons index
                    SendPlayerVital index
                    SendPlayerPokemonStatus index
                    SendPlayerStatus index
                End If
            Case MapAttribute.Checkpoint
                .CheckMap = Map(.Map).Tile(.x, .Y).Data1
                .CheckX = Map(.Map).Tile(.x, .Y).Data2
                .CheckY = Map(.Map).Tile(.x, .Y).Data3
                .CheckDir = Map(.Map).Tile(.x, .Y).Data4
            Case MapAttribute.WarpCheckpoint
                If .CheckMap > 0 Then
                    PlayerWarp index, .CheckMap, .CheckX, .CheckY, .CheckDir
                End If
            End Select
        End If
    End With
End Sub

Public Sub SpawnPlayerPokemon(ByVal index As Long, ByVal PokeSlot As Byte)
    Dim MapNum As Long
    Dim statX As Byte
    Dim startPosX As Long, startPosY As Long
    Dim x As Long, Y As Long
    Dim canSpawn As Boolean
    Dim UsedBall As Byte

    If index <= 0 Or index > MAX_PLAYER Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Or TempPlayer(index).UseChar > MAX_PLAYERCHAR Then Exit Sub
    If PlayerPokemon(index).Num > 0 Then Exit Sub
    If PlayerPokemons(index).Data(PokeSlot).Num <= 0 Then Exit Sub

    MapNum = Player(index, TempPlayer(index).UseChar).Map

    '//Update Position
    With PlayerPokemon(index)
        canSpawn = False
        For x = Player(index, TempPlayer(index).UseChar).x - 1 To Player(index, TempPlayer(index).UseChar).x + 1
            For Y = Player(index, TempPlayer(index).UseChar).Y - 1 To Player(index, TempPlayer(index).UseChar).Y + 1
                If x = Player(index, TempPlayer(index).UseChar).x And Y = Player(index, TempPlayer(index).UseChar).Y Then

                Else
                    '//Check if OpenTile
                    If CheckOpenTile(MapNum, x, Y) Then
                        startPosX = x
                        startPosY = Y
                        canSpawn = True
                        Exit For
                    End If
                End If
            Next
        Next

        If canSpawn Then
            .Num = PlayerPokemons(index).Data(PokeSlot).Num
            .x = startPosX
            .Y = startPosY
            .Dir = DIR_DOWN

            .slot = PokeSlot

            .QueueMove = 0
            .QueueMoveSlot = 0
            .MoveDuration = 0
            .MoveInterval = 0
            .MoveAttackCount = 0
            .MoveCastTime = 0
            .IsConfuse = NO
            .NextCritical = NO
            .ReflectMove = 0
            .IsProtect = 0

            For statX = 1 To StatEnum.Stat_Count - 1
                .StatBuff(statX) = 0
            Next
            UsedBall = PlayerPokemons(index).Data(.slot).BallUsed

            .StatusDamage = 0
            .StatusMove = 0
        Else
            Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: AddAlert index, "Out of space", White
            Case LANG_EN: AddAlert index, "Out of space", White
            Case LANG_ES: AddAlert index, "Out of space", White
            End Select
        End If
    End With

    '//Update
    If canSpawn Then SendPlayerPokemonData index, MapNum, , YES, 0, startPosX, startPosY, UsedBall
End Sub

Public Sub ClearPlayerPokemon(ByVal index As Long)
    Dim MapNum As Long
    Dim endPosX As Long, endPosY As Long
    Dim BallUsed As Byte

    If index <= 0 Or index > MAX_PLAYER Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Or TempPlayer(index).UseChar > MAX_PLAYERCHAR Then Exit Sub
    If PlayerPokemon(index).Num <= 0 Then Exit Sub

    MapNum = Player(index, TempPlayer(index).UseChar).Map

    '//Update Position
    With PlayerPokemon(index)
        BallUsed = PlayerPokemons(index).Data(.slot).BallUsed

        .Num = 0
        endPosX = .x
        endPosY = .Y
        .x = 0
        .Y = 0
        .Dir = 0

        .slot = 0
    End With

    '//Update
    SendPlayerPokemonData index, MapNum, , YES, 1, endPosX, endPosY, BallUsed
End Sub

Public Sub PlayerPokemonWarp(ByVal index As Long, ByVal x As Long, ByVal Y As Long, ByVal Dir As Byte)
    Dim MapNum As Long

    '//Exit out when error
    If index <= 0 Or index > MAX_PLAYER Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Or TempPlayer(index).UseChar > MAX_PLAYERCHAR Then Exit Sub
    If MapNum <= 0 Or MapNum > MAX_MAP Then Exit Sub
    If PlayerPokemon(index).Num <= 0 Then Exit Sub

    '//Correct error position
    If x <= 0 Then x = 0
    If x > Map(MapNum).MaxX Then x = Map(MapNum).MaxX
    If Y <= 0 Then Y = 0
    If Y > Map(MapNum).MaxY Then Y = Map(MapNum).MaxY

    MapNum = Player(index, TempPlayer(index).UseChar).Map

    '//Update position
    With PlayerPokemon(index)
        .x = x
        .Y = Y
        .Dir = Dir
    End With

    '//Add log
    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " pokemon has been warped on Map#" & MapNum & " x:" & x & " y:" & Y
End Sub

Public Sub PlayerPokemonMove(ByVal index As Long, ByVal Dir As Byte, Optional ByVal sendToSelf As Boolean = False)
    Dim DidMove As Boolean
    Dim OldX As Long, OldY As Long
    Dim MapNum As Long
    Dim dX As Long, dY As Long

    '//Exit out when error
    If Not IsPlaying(index) Then Exit Sub
    If index <= 0 Or index > MAX_PLAYER Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Or TempPlayer(index).UseChar > MAX_PLAYERCHAR Then Exit Sub
    If Dir < 0 Or Dir > DIR_RIGHT Then Exit Sub
    If PlayerPokemon(index).Num <= 0 Then Exit Sub

    DidMove = False

    MapNum = Player(index, TempPlayer(index).UseChar).Map

    With PlayerPokemon(index)
        '//Store original location in case it got desync
        OldX = .x
        OldY = .Y

        Select Case Dir
        Case DIR_UP
            .Dir = DIR_UP

            '//Check to make sure not outside of boundries
            If .Y > 0 Then
                If Not CheckDirection(MapNum, DIR_UP, .x, .Y) Then
                    '//Check Distance
                    dX = .x - Player(index, TempPlayer(index).UseChar).x
                    dY = (.Y - 1) - Player(index, TempPlayer(index).UseChar).Y

                    '//Make sure we get a positive value
                    If dX < 0 Then dX = dX * -1
                    If dY < 0 Then dY = dY * -1

                    If Not (dX <= MAX_DISTANCE And dY <= MAX_DISTANCE) Then
                        DidMove = False
                    Else
                        .Y = .Y - 1
                        DidMove = True
                    End If
                End If
            End If
        Case DIR_DOWN
            .Dir = DIR_DOWN

            '//Check to make sure not outside of boundries
            If .Y < Map(MapNum).MaxY Then
                If Not CheckDirection(MapNum, DIR_DOWN, .x, .Y) Then
                    '//Check Distance
                    dX = .x - Player(index, TempPlayer(index).UseChar).x
                    dY = (.Y + 1) - Player(index, TempPlayer(index).UseChar).Y

                    '//Make sure we get a positive value
                    If dX < 0 Then dX = dX * -1
                    If dY < 0 Then dY = dY * -1

                    If Not (dX <= MAX_DISTANCE And dY <= MAX_DISTANCE) Then
                        DidMove = False
                    Else
                        .Y = .Y + 1
                        DidMove = True
                    End If
                End If
            End If
        Case DIR_LEFT
            .Dir = DIR_LEFT

            '//Check to make sure not outside of boundries
            If .x > 0 Then
                If Not CheckDirection(MapNum, DIR_LEFT, .x, .Y) Then
                    '//Check Distance
                    dX = (.x - 1) - Player(index, TempPlayer(index).UseChar).x
                    dY = .Y - Player(index, TempPlayer(index).UseChar).Y

                    '//Make sure we get a positive value
                    If dX < 0 Then dX = dX * -1
                    If dY < 0 Then dY = dY * -1

                    If Not (dX <= MAX_DISTANCE And dY <= MAX_DISTANCE) Then
                        DidMove = False
                    Else
                        .x = .x - 1
                        DidMove = True
                    End If
                End If
            End If
        Case DIR_RIGHT
            .Dir = DIR_RIGHT

            '//Check to make sure not outside of boundries
            If .x < Map(MapNum).MaxX Then
                If Not CheckDirection(MapNum, DIR_RIGHT, .x, .Y) Then
                    '//Check Distance
                    dX = (.x + 1) - Player(index, TempPlayer(index).UseChar).x
                    dY = .Y - Player(index, TempPlayer(index).UseChar).Y

                    '//Make sure we get a positive value
                    If dX < 0 Then dX = dX * -1
                    If dY < 0 Then dY = dY * -1

                    If Not (dX <= MAX_DISTANCE And dY <= MAX_DISTANCE) Then
                        DidMove = False
                    Else
                        .x = .x + 1
                        DidMove = True
                    End If
                End If
            End If
        End Select

        '//Got Desynced
        If Not DidMove Then
            .x = OldX
            .Y = OldY
            SendPlayerPokemonXY index
            SendPlayerPokemonXY index, True
        Else
            SendPlayerPokemonMove index, sendToSelf
        End If
    End With
End Sub

' ******************
' ** Player Logic **
' ******************
Public Sub JoinGame(ByVal index As Long, Optional ByVal CurLanguage As Byte = 0)
    Dim countOnline As Long

    '//Exit out if not connected
    If Not IsConnected(index) Then Exit Sub
    '//Exit out if already playing
    If TempPlayer(index).InGame Then Exit Sub

    frmServer.lvwInfo.ListItems(index).SubItems(1) = GetPlayerIP(index)
    frmServer.lvwInfo.ListItems(index).SubItems(2) = GetPlayerLogin(index)
    frmServer.lvwInfo.ListItems(index).SubItems(3) = Player(index, TempPlayer(index).UseChar).Name

    '//Check if staff only
    If frmServer.chkStaffOnly.Value = YES Then
        If Player(index, TempPlayer(index).UseChar).Access <= 0 Then
            Select Case CurLanguage
            Case LANG_PT: AddAlert index, "Server is available for Staff Members only", White
            Case LANG_EN: AddAlert index, "Server is available for Staff Members only", White
            Case LANG_ES: AddAlert index, "Server is available for Staff Members only", White
            End Select
            Exit Sub
        End If
    End If

    '//Load Player Pokemon
    If TempPlayer(index).UseChar > 0 Then
        LoadPlayerInv index, TempPlayer(index).UseChar
        LoadPlayerPokemons index, TempPlayer(index).UseChar
        LoadPlayerInvStorage index, TempPlayer(index).UseChar
        LoadPlayerPokemonStorage index, TempPlayer(index).UseChar
        LoadPlayerPokedex index, TempPlayer(index).UseChar
    End If

    '//Set player in-game
    TempPlayer(index).InGame = True
    TempPlayer(index).CurLanguage = CurLanguage
    TempPlayer(index).MapSwitchTmr = YES
    '//Send Data to Client

    '//Send Data
    AddAlert index, "Loading Npcs...", White, , YES
    SendNpcs index
    AddAlert index, "Loading Pokemons...", White, , YES
    SendPokemons index
    AddAlert index, "Loading Items...", White, , YES
    SendItems index
    AddAlert index, "Loading Moves...", White, , YES
    SendPokemonMoves index
    AddAlert index, "Loading Animations...", White, , YES
    SendAnimations index
    AddAlert index, "Loading Spawns...", White, , YES
    SendSpawns index
    'If Player(Index, TempPlayer(Index).UseChar).Access > ACCESS_MAPPER Then
    '   SendConversations Index
    'End If


    AddAlert index, "Loading Shop...", White, , YES
    SendShops index
    AddAlert index, "Loading Quest...", White, , YES
    SendQuests index
    AddAlert index, "Loading Inventory...", White, , YES
    SendPlayerInv index
    AddAlert index, "Loading Item Storage...", White, , YES
    SendPlayerInvStorage index
    AddAlert index, "Loading Team...", White, , YES
    SendPlayerPokemons index
    AddAlert index, "Loading Pokemon Box...", White, , YES
    SendPlayerPokemonStorage index
    AddAlert index, "Loading Pokedex...", White, , YES
    SendPlayerPokedex index
    AddAlert index, "Send Raking To Client...", White, , YES
    SendRankTo index
    AddAlert index, "Send Event Exp To Client...", White, , YES
    SendEventInfo index
    AddAlert index, "Send Vip Status To Client...", White, , YES
    CheckVipJoinGame index

    If Player(index, TempPlayer(index).UseChar).Access = ACCESS_NONE Then
        UpdateRank Trim$(Player(index, TempPlayer(index).UseChar).Name), Player(index, TempPlayer(index).UseChar).Level, Player(index, TempPlayer(index).UseChar).CurExp
    End If
    'LoadRank

    '//Send data to position
    With Player(index, TempPlayer(index).UseChar)
        PlayerWarp index, .Map, .x, .Y, .Dir

        '//Check online
        countOnline = TotalPlayerOnline

        If .Access < ACCESS_CREATOR Then
            SendMapMsg .Map, Trim$(.Name) & " has joined the game", White
        End If
        'AddLog Trim$(.Name) & " has joined the game"

        '//Send count msg
        If countOnline > 1 Then
            SendPlayerMsg index, "There are " & (countOnline - 1) & " other players online", White
        Else
            SendPlayerMsg index, "There are no other players online", White
        End If
    End With

    '//Send Message
    SendPlayerMsg index, "Welcome to " & GAME_NAME, White
    If Len(Trim$(Options.MOTD)) > 0 Then
        SendPlayerMsg index, Trim$(Options.MOTD), White
    End If
    '//Send tutorial message
    If CountPlayerPokemon(index) <= 0 Then
        '//Init Starter Pokemon
        TempPlayer(index).CurConvoNum = 1
        TempPlayer(index).CurConvoData = 1
        TempPlayer(index).CurConvoNpc = 3
        TempPlayer(index).CurConvoMapNpc = 0
        SendInitConvo index, TempPlayer(index).CurConvoNum, TempPlayer(index).CurConvoData, TempPlayer(index).CurConvoNpc
    Else
        Player(index, TempPlayer(index).UseChar).DidStart = NO
        SavePlayerData index, TempPlayer(index).UseChar
    End If
    
    '//Processa sprite temporaria do jogador
    ProcessTempSprite index

    '//Send In-Game
    SendHighIndex index
    SendPokemonHighIndex index
    SendInGame index
End Sub

Private Sub ProcessTempSprite(ByVal index As Long)
    Dim ItemNum As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub

    ItemNum = Player(index, TempPlayer(index).UseChar).KeyItemNum

    If ItemNum > 0 And ItemNum <= MAX_ITEM Then
        If Item(ItemNum).Data1 = 1 Then    '//Sprite Type
            'If Item(ItemNum).Data2 = TEMP_SPRITE_GROUP_MOUNT Then
                If Map(Player(index, TempPlayer(index).UseChar).Map).SpriteType <= TEMP_SPRITE_GROUP_NONE Then
                    ChangeTempSprite index, Item(ItemNum).Data2, ItemNum
                End If
            'End If
        End If
    End If
End Sub

Public Sub ClearTempSprite(ByVal index As Long)
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub

    Call ZeroMemory(ByVal VarPtr(TempPlayer(index).TempSprite), LenB(TempPlayer(index).TempSprite))
    Player(index, TempPlayer(index).UseChar).KeyItemNum = 0
End Sub

Public Sub LeftGame(ByVal index As Long)
    Dim sIP As String
    Dim i As Long, x As Byte, Y As Byte

    sIP = GetPlayerIP(index)

    '//Update HighIndex
    If Player_HighIndex = index Then
        Player_HighIndex = Player_HighIndex - 1
        '//Update Index to all
        SendHighIndex
    End If

    '//InGame Data
    If TempPlayer(index).InGame Then
        '//Request
        i = TempPlayer(index).PlayerRequest
        If i > 0 Then
            '//Cancel Request to index
            If IsPlaying(i) Then
                If TempPlayer(i).UseChar > 0 Then
                    If TempPlayer(i).PlayerRequest = index Then
                        If TempPlayer(index).RequestType = 1 Then  '//1 Duel
                            '//Check if already in duel
                            If TempPlayer(index).InDuel > 0 Then
                                SendActionMsg Player(i, TempPlayer(i).UseChar).Map, "Win!", Player(i, TempPlayer(i).UseChar).x * 32, Player(i, TempPlayer(i).UseChar).Y * 32, White
                                Player(i, TempPlayer(i).UseChar).Win = Player(i, TempPlayer(i).UseChar).Win + 1
                                SendPlayerPvP (i)
                                TempPlayer(i).InDuel = 0
                                TempPlayer(i).DuelTime = 0
                                TempPlayer(i).DuelTimeTmr = 0
                                TempPlayer(i).WarningTimer = 0
                                TempPlayer(i).PlayerRequest = 0
                                TempPlayer(i).RequestType = 0
                                SendRequest i
                            Else
                                '//Cancel Request to index
                                TempPlayer(i).PlayerRequest = 0
                                TempPlayer(i).RequestType = 0
                                SendRequest i
                                Select Case TempPlayer(i).CurLanguage
                                Case LANG_PT: AddAlert i, "Duel request has been cancelled", White
                                Case LANG_EN: AddAlert i, "Duel request has been cancelled", White
                                Case LANG_ES: AddAlert i, "Duel request has been cancelled", White
                                End Select
                            End If
                        ElseIf TempPlayer(index).RequestType = 2 Then    '//trade
                            '//Check if already in trade
                            If TempPlayer(index).InTrade > 0 Then
                                TempPlayer(i).InTrade = 0
                                For x = 1 To MAX_TRADE
                                    Call ZeroMemory(ByVal VarPtr(TempPlayer(i).TradeItem(x)), LenB(TempPlayer(i).TradeItem(x)))
                                Next
                                TempPlayer(i).TradeMoney = 0
                                TempPlayer(i).TradeSet = 0
                                TempPlayer(i).TradeAccept = 0
                                TempPlayer(i).PlayerRequest = 0
                                TempPlayer(i).RequestType = 0
                                Select Case TempPlayer(i).CurLanguage
                                Case LANG_PT: AddAlert i, "The trade was declined", White
                                Case LANG_EN: AddAlert i, "The trade was declined", White
                                Case LANG_ES: AddAlert i, "The trade was declined", White
                                End Select
                                SendCloseTrade i
                                SendRequest i
                            Else
                                '//Cancel Request to index
                                TempPlayer(i).PlayerRequest = 0
                                TempPlayer(i).RequestType = 0
                                SendRequest i
                                Select Case TempPlayer(i).CurLanguage
                                Case LANG_PT: AddAlert i, "Trade request has been cancelled", White
                                Case LANG_EN: AddAlert i, "Trade request has been cancelled", White
                                Case LANG_ES: AddAlert i, "Trade request has been cancelled", White
                                End Select
                            End If
                        ElseIf TempPlayer(index).RequestType = 3 Then    '//Party
                            '//Cancel Request to index
                            TempPlayer(i).PlayerRequest = 0
                            TempPlayer(i).RequestType = 0
                            SendRequest i
                            Select Case TempPlayer(i).CurLanguage
                            Case LANG_PT: AddAlert i, "Party request has been cancelled", White
                            Case LANG_EN: AddAlert i, "Party request has been cancelled", White
                            Case LANG_ES: AddAlert i, "Party request has been cancelled", White
                            End Select
                        End If
                    End If
                End If
            End If
        End If

        '//Check if already in party
        If TempPlayer(index).InParty > 0 Then
            LeaveParty index
        End If

        TempPlayer(index).InDuel = 0
        TempPlayer(index).DuelTime = 0
        TempPlayer(index).DuelTimeTmr = 0
        TempPlayer(index).WarningTimer = 0
        TempPlayer(index).PlayerRequest = 0
        TempPlayer(index).RequestType = 0
        TempPlayer(index).InTrade = 0
        For x = 1 To MAX_TRADE
            Call ZeroMemory(ByVal VarPtr(TempPlayer(index).TradeItem(x)), LenB(TempPlayer(index).TradeItem(x)))
        Next
        TempPlayer(index).TradeMoney = 0
        TempPlayer(index).TradeSet = 0
        TempPlayer(index).PlayerRequest = 0
        TempPlayer(index).RequestType = 0

        If Player(index, TempPlayer(index).UseChar).Access = ACCESS_NONE Then
            If TempPlayer(index).UseChar > 0 Then
                UpdateRank Trim$(Player(index, TempPlayer(index).UseChar).Name), Player(index, TempPlayer(index).UseChar).Level, Player(index, TempPlayer(index).UseChar).CurExp
            End If
        End If

        TempPlayer(index).InGame = False

        '//Clear In-Game Data

        '//Save Player data
        SavePlayerDatas index

        '//Left Game
        SendLeftGame index

        If TempPlayer(index).UseChar > 0 Then
            If Player(index, TempPlayer(index).UseChar).Access < ACCESS_CREATOR Then
                SendMapMsg Player(index, TempPlayer(index).UseChar).Map, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " has left the game", White
            End If
            'AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " has left the game"
        End If
    End If

    '//Clear Player Data
    ClearTempPlayer index
    ClearPlayer index
    ClearPlayerInv index
    ClearPlayerInvStorage index
    ClearPlayerPokemons index
    ClearPlayerPokemonStorage index
    ClearAccount index
    ClearPlayerPokedex index

    'AddLog "Connection from " & sIP & " has been terminated"
End Sub

Public Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If TempPlayer(i).UseChar > 0 Then
                If UCase$(Trim$(Player(i, TempPlayer(i).UseChar).Name)) = UCase$(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If
    Next

    FindPlayer = 0
End Function

Public Function FindAccount(ByVal Name As String) As Long
    Dim i As Long

    For i = 1 To Player_HighIndex
        If Len(Account(i).Username) > 0 Then
            If UCase$(Trim$(Account(i).Username)) = UCase$(Trim$(Name)) Then
                FindAccount = i
                Exit Function
            End If
        End If
    Next

    FindAccount = 0
End Function

Public Function FindSameItemSlot(ByVal index As Long, ByVal ItemNum As Long) As Byte
    Dim i As Byte

    FindSameItemSlot = 0

    If Not IsPlaying(index) Then Exit Function
    If TempPlayer(index).UseChar <= 0 Then Exit Function

    For i = 1 To MAX_PLAYER_INV
        With PlayerInv(index).Data(i)
            If .Num = ItemNum Then
                If Item(ItemNum).Stock = YES Then
                    '//add val
                    FindSameItemSlot = i
                    Exit Function
                End If
            End If
        End With
    Next
End Function

Public Function FindFreeInvSlot(ByVal index As Long, ByVal ItemNum As Long, ByRef ItemVal As Long, Optional ByRef MsgFrom As String) As Byte
    Dim i As Byte

    FindFreeInvSlot = 0

    If Not IsPlaying(index) Then Exit Function
    If TempPlayer(index).UseChar <= 0 Then Exit Function

    If Item(ItemNum).Stock = YES Then
        i = FindSameItemSlot(index, ItemNum)
        If i > 0 Then
            If CheckInvValues(index, i, ItemNum, ItemVal, MsgFrom) Then
                FindFreeInvSlot = i
                Exit Function
            Else
                Exit Function
            End If
        End If
    End If

    For i = 1 To MAX_PLAYER_INV
        With PlayerInv(index).Data(i)
            If .Locked = NO Then
                If .Num = 0 Then
                    If CheckInvValues(index, i, ItemNum, ItemVal, MsgFrom) Then
                        FindFreeInvSlot = i
                        Exit Function
                    End If
                End If
            End If
        End With
    Next
End Function

Private Function CheckInvValues(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long, ByRef ItemVal As Long, Optional ByRef MsgFrom As String) As Boolean
    CheckInvValues = True

    '//Adaptações pra não ultrapassar um limite de Amount
    If ItemVal > MAX_AMOUNT Then
        Select Case TempPlayer(index).CurLanguage
        Case LANG_PT: MsgFrom = "Quantidade Limite " & MAX_AMOUNT
        Case LANG_EN: MsgFrom = "Limit Quantity " & MAX_AMOUNT
        Case LANG_ES: MsgFrom = "Limit Quantity " & MAX_AMOUNT
        End Select

        CheckInvValues = False
        Exit Function
    End If

    If PlayerInv(index).Data(InvSlot).Value >= MAX_AMOUNT Then
        Select Case TempPlayer(index).CurLanguage
        Case LANG_PT: MsgFrom = "Quantidade Limite " & MAX_AMOUNT
        Case LANG_EN: MsgFrom = "Limit Quantity " & MAX_AMOUNT
        Case LANG_ES: MsgFrom = "Limit Quantity " & MAX_AMOUNT
        End Select

        CheckInvValues = False
        Exit Function
    End If

    If (ItemVal + PlayerInv(index).Data(InvSlot).Value) > MAX_AMOUNT Then
        '//Altera o valor pra obter apenas o que couber
        ItemVal = MAX_AMOUNT - PlayerInv(index).Data(InvSlot).Value

        Select Case TempPlayer(index).CurLanguage
        Case LANG_PT: MsgFrom = "Quantidade Excedida, você recebeu apenas (" & ItemVal & ")"
        Case LANG_EN: MsgFrom = "Quantity Exceeded, you have only received (" & ItemVal & ")"
        Case LANG_ES: MsgFrom = "Quantity Exceeded, you have only received (" & ItemVal & ")"
        End Select

        CheckInvValues = True
        Exit Function
    End If
End Function

Public Function TryGivePlayerItem(ByVal index As Long, ByVal ItemNum As Long, ByRef ItemVal As Long, Optional ByVal TmrCooldown As Long = 0) As Boolean
    Dim MsgFrom As String    '--> Utilizado como referência pra obter mensagem de outro método.

    TryGivePlayerItem = True
    If Not GiveItem(index, ItemNum, ItemVal, TmrCooldown, MsgFrom) Then

        If MsgFrom = vbNullString Then
            Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: MsgFrom = "Inventory is full"
            Case LANG_EN: MsgFrom = "Inventory is full"
            Case LANG_ES: MsgFrom = "Inventory is full"
            End Select
        End If

        AddAlert index, MsgFrom, White
        TryGivePlayerItem = False
    Else
        '//Check if there's still free slot
        If CountFreeInvSlot(index) <= 5 Then
            Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: AddAlert index, "Warning: Your inventory is almost full", White
            Case LANG_EN: AddAlert index, "Warning: Your inventory is almost full", White
            Case LANG_ES: AddAlert index, "Warning: Your inventory is almost full", White
            End Select
        End If
    End If
End Function

Public Function CountFreeInvSlot(ByVal index As Long) As Long
    Dim count As Long, i As Long

    CountFreeInvSlot = 0
    count = 0

    If Not IsPlaying(index) Then Exit Function
    If TempPlayer(index).UseChar <= 0 Then Exit Function

    For i = 1 To MAX_PLAYER_INV
        With PlayerInv(index).Data(i)
            If .Num = 0 Then
                count = count + 1
            End If
        End With
    Next

    CountFreeInvSlot = count
End Function

Public Function GiveItem(ByVal index As Long, ByVal ItemNum As Long, ByRef ItemVal As Long, Optional ByVal TmrCooldown As Long = 0, Optional ByRef MsgFrom As String) As Boolean
    Dim i As Byte

    '//Get Slot
    i = FindFreeInvSlot(index, ItemNum, ItemVal, MsgFrom)

    '//Got slot
    If i > 0 Then
        With PlayerInv(index).Data(i)
            .Num = ItemNum
            .Value = .Value + ItemVal
            .TmrCooldown = TmrCooldown
        End With
        '//Update
        SendPlayerInvSlot index, i
        GiveItem = True
    Else
        GiveItem = False
    End If
End Function

'//Player Pokemon
Public Function FindOpenPokeSlot(ByVal index As Long) As Long
    Dim i As Byte

    For i = 1 To MAX_PLAYER_POKEMON
        If PlayerPokemons(index).Data(i).Num = 0 Then
            FindOpenPokeSlot = i
            Exit Function
        End If
    Next
End Function

Public Sub GivePlayerPokemon(ByVal index As Long, ByVal PokeNum As Long, ByVal Level As Long, ByVal BallUsed As Byte, Optional ByVal IsShiny As Byte = NO, _
                             Optional ByVal IVFull As Byte = NO, Optional ByVal TheNature As Integer = -1)
    Dim i As Long, x As Byte, m As Long, s As Byte, slot As Byte, StorageSlot As Byte, gotSlot As Byte

    i = FindOpenPokeSlot(index)

    '//Got slot
    If i > 0 Then
        With PlayerPokemons(index).Data(i)
            .Num = PokeNum

            .Level = Level

            '//Nature
            If TheNature = -1 Then .Nature = Random(0, PokemonNature.NatureQuirky)
            If TheNature >= 0 Then .Nature = TheNature    'Peronalização do painel admin
            If .Nature >= PokemonNature.PokemonNature_Count - 1 Then .Nature = PokemonNature.PokemonNature_Count - 1
            .IsShiny = IsShiny    'Peronalização do painel admin
            .Status = 0


            .Happiness = 0
            .Gender = Random(GENDER_MALE, GENDER_FEMALE)
            If Not .Gender = GENDER_MALE And Not .Gender = GENDER_FEMALE Then
                .Gender = GENDER_MALE
            End If

            '//Stat
            For x = 1 To StatEnum.Stat_Count - 1
                .Stat(x).EV = 0
                .Stat(x).IV = 15    '//Default Stat
                If IVFull > 0 Then .Stat(x).IV = 31    'Peronalização do painel admin
                .Stat(x).Value = CalculatePokemonStat(x, .Num, .Level, .Stat(x).EV, .Stat(x).IV, .Nature)
            Next

            '//Vital
            .MaxHp = .Stat(StatEnum.HP).Value
            .CurHp = .MaxHp

            '//Ball Used
            .BallUsed = BallUsed

            .HeldItem = 0

            '//Moveset
            If PokeNum > 0 Then
                For m = MAX_POKEMON_MOVESET To 1 Step -1
                    '//Got Move
                    If Pokemon(PokeNum).Moveset(m).MoveNum > 0 Then
                        '//Check level
                        If .Level >= Pokemon(PokeNum).Moveset(m).MoveLevel Then
                            slot = 0
                            For s = 1 To MAX_MOVESET
                                If .Moveset(s).Num <= 0 Then
                                    slot = s
                                    Exit For
                                End If
                            Next

                            '//Add Move
                            If slot > 0 Then
                                .Moveset(slot).Num = Pokemon(PokeNum).Moveset(m).MoveNum
                                .Moveset(slot).TotalPP = PokemonMove(Pokemon(PokeNum).Moveset(m).MoveNum).PP
                                .Moveset(slot).CurPP = .Moveset(slot).TotalPP
                            End If
                        End If
                    End If
                Next
            End If

            '//Add Pokedex
            AddPlayerPokedex index, .Num, YES, YES
        End With
        '//Update
        SendPlayerPokemonSlot index, i
    Else
        For StorageSlot = 1 To MAX_STORAGE_SLOT
            gotSlot = FindFreePokeStorageSlot(index, StorageSlot)
            If gotSlot > 0 Then
                With PlayerPokemonStorage(index).slot(StorageSlot).Data(gotSlot)

                    .Num = PokeNum

                    .Level = Level

                    '//Nature
                    If TheNature = -1 Then .Nature = Random(0, PokemonNature.NatureQuirky)
                    If TheNature >= 0 Then .Nature = TheNature    'Peronalização do painel admin
                    If .Nature >= PokemonNature.PokemonNature_Count - 1 Then .Nature = PokemonNature.PokemonNature_Count - 1
                    .IsShiny = IsShiny    'Peronalização do painel admin
                    .Status = 0


                    .Happiness = 0
                    .Gender = Random(GENDER_MALE, GENDER_FEMALE)
                    If Not .Gender = GENDER_MALE And Not .Gender = GENDER_FEMALE Then
                        .Gender = GENDER_MALE
                    End If

                    If TheNature > 0 Then .Nature = TheNature    'Peronalização do painel admin

                    '//Stat
                    For x = 1 To StatEnum.Stat_Count - 1
                        .Stat(x).EV = 0
                        .Stat(x).IV = 15    '//Default Stat
                        If IVFull > 0 Then .Stat(x).IV = 31    'Peronalização do painel admin
                        .Stat(x).Value = CalculatePokemonStat(x, .Num, .Level, .Stat(x).EV, .Stat(x).IV, .Nature)
                    Next

                    '//Vital
                    .MaxHp = .Stat(StatEnum.HP).Value
                    .CurHp = .MaxHp

                    '//Ball Used
                    .BallUsed = BallUsed

                    .HeldItem = 0

                    '//Moveset
                    If PokeNum > 0 Then
                        For m = MAX_POKEMON_MOVESET To 1 Step -1
                            '//Got Move
                            If Pokemon(PokeNum).Moveset(m).MoveNum > 0 Then
                                '//Check level
                                If .Level >= Pokemon(PokeNum).Moveset(m).MoveLevel Then
                                    slot = 0
                                    For s = 1 To MAX_MOVESET
                                        If .Moveset(s).Num <= 0 Then
                                            slot = s
                                            Exit For
                                        End If
                                    Next

                                    '//Add Move
                                    If slot > 0 Then
                                        .Moveset(slot).Num = Pokemon(PokeNum).Moveset(m).MoveNum
                                        .Moveset(slot).TotalPP = PokemonMove(Pokemon(PokeNum).Moveset(m).MoveNum).PP
                                        .Moveset(slot).CurPP = .Moveset(slot).TotalPP
                                    End If
                                End If
                            End If
                        Next
                    End If

                    '//Add Pokedex
                    AddPlayerPokedex index, .Num, YES, YES

                    Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "Your pokemon has been transferred to your pokemon storage", White
                    Case LANG_EN: AddAlert index, "Your pokemon has been transferred to your pokemon storage", White
                    Case LANG_ES: AddAlert index, "Your pokemon has been transferred to your pokemon storage", White
                    End Select
                    SendPlayerPokemonStorageSlot index, StorageSlot, gotSlot
                    Exit Sub
                End With
            End If
        Next StorageSlot
    End If
End Sub

Public Sub UpdatePlayerPokemonOrder(ByVal index As Long)
    Dim i As Long

    For i = 2 To MAX_PLAYER_POKEMON
        With PlayerPokemons(index)
            '//Check if previous number is empty
            If .Data(i - 1).Num = 0 Then
                '//Move Data
                .Data(i - 1) = .Data(i)
                Call ZeroMemory(ByVal VarPtr(.Data(i)), LenB(.Data(i)))
            End If
        End With
    Next
End Sub

Public Function CountPlayerPokemon(ByVal index As Long) As Byte
    Dim i As Byte
    Dim count As Byte

    count = 0
    For i = 1 To MAX_PLAYER_POKEMON
        With PlayerPokemons(index).Data(i)
            If .Num > 0 Then
                count = count + 1
            End If
        End With
    Next
    CountPlayerPokemon = count
End Function

Public Function CountPlayerPokemonAlive(ByVal index As Long) As Byte
    Dim i As Byte
    Dim count As Byte

    count = 0
    For i = 1 To MAX_PLAYER_POKEMON
        With PlayerPokemons(index).Data(i)
            If .Num > 0 Then
                If .CurHp > 0 Then
                    count = count + 1
                End If
            End If
        End With
    Next
    CountPlayerPokemonAlive = count
End Function

'//Exp
Public Sub GivePlayerPokemonExp(ByVal index As Long, ByVal PokeSlot As Byte, ByVal Exp As Long)
    Dim TotalBonus As Long
'//Check Error
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If PokeSlot <= 0 Or PokeSlot > MAX_PLAYER_POKEMON Then Exit Sub
    If PlayerPokemons(index).Data(PokeSlot).Num <= 0 Then Exit Sub

    'Exp Rate
    If EventExp.ExpEvent Then
        Exp = Exp * EventExp.ExpMultiply
        '//Obter o bonus somado, pra entregar uma action mensagem com o total bonificado
        TotalBonus = TotalBonus + (EventExp.ExpMultiply * 100)
    End If
    'Exp Mount
    If TempPlayer(index).TempSprite.TempSpriteExp > 0 Then
        Exp = Exp + ((Exp / 100) * TempPlayer(index).TempSprite.TempSpriteExp)
        '//Obter o bonus somado, pra entregar uma action mensagem com o total bonificado
        TotalBonus = TotalBonus + (TempPlayer(index).TempSprite.TempSpriteExp)
    End If

    '//Add Exp
    With PlayerPokemons(index).Data(PokeSlot)
        '//Make sure we can give it exp based on player level
        If Player(index, TempPlayer(index).UseChar).Level + 10 <= .Level Then Exit Sub
        If .Level >= MAX_LEVEL Then Exit Sub

        .CurExp = .CurExp + Exp
        TextAdd frmServer.txtLog, "EXP: " & Exp

        '//ActionMsg
        If PlayerPokemon(index).Num > 0 Then
            If PlayerPokemon(index).slot = PokeSlot Then
                SendActionMsg Player(index, TempPlayer(index).UseChar).Map, "+" & Exp, PlayerPokemon(index).x * 32, PlayerPokemon(index).Y * 32, White
                
                If TotalBonus > 0 Then
                    SendActionMsg Player(index, TempPlayer(index).UseChar).Map, "+" & TotalBonus & "% EXP!", PlayerPokemon(index).x * 32, (PlayerPokemon(index).Y - 1) * 32, Black
                End If
            End If
        End If
    End With
    CheckPlayerPokemonLevelUp index, PokeSlot
End Sub

Public Function GivePlayerEvPowerBracer(ByVal index As Long, ByVal PokeSlot As Byte) As Boolean
    Dim CallBack As Integer
    GivePlayerEvPowerBracer = False

    With PlayerPokemons(index).Data(PokeSlot)
        If .HeldItem > 0 Then
            If Item(.HeldItem).Type = ItemTypeEnum.PowerBracer Then
                If Item(.HeldItem).Data1 >= StatEnum.HP And Item(.HeldItem).Data1 <= StatEnum.Spd Then
                    GivePlayerEvPowerBracer = True
                    CallBack = GivePlayerPokemonEVExp(index, PokeSlot, Item(.HeldItem).Data1, Item(.HeldItem).Data2)
                End If
            End If
        End If
    End With
End Function

Public Function GivePlayerPokemonEVExp(ByVal index As Long, ByVal PokeSlot As Byte, ByVal evStat As StatEnum, ByVal Exp As Long) As Integer
    Dim CountStat As Long, x As Byte, statMaxEv As Integer, Sobra As Integer

    '// Função implementada pra utilizar => Recebendo ao matar um poke,
    '                                       Ao utilizar items Barries
    '                                       Ao utilizar items Protein
    '                                       Ao utilizar Power Bracer no pokemon.

    With PlayerPokemons(index).Data(PokeSlot)
        ' Máximo de EV Total
        ' MAX_EV = 510

        ' Máximo de Ev em cada atributo
        statMaxEv = 252

        ' Faz a contagem do total de EV
        CountStat = 0
        For x = 1 To StatEnum.Stat_Count - 1
            CountStat = CountStat + PlayerPokemons(index).Data(PokeSlot).Stat(x).EV
        Next

        ' Verifica se tem a possibilidade de adicionar a exp, sem passar o máximo de EV.
        If CountStat + Exp <= MAX_EV And CountStat + Exp >= 0 Then
            If Exp > 0 Then    ' Valor Positivo
                If .Stat(evStat).EV + Exp <= statMaxEv Then
                    .Stat(evStat).EV = .Stat(evStat).EV + Exp
                    .Stat(evStat).Value = CalculatePokemonStat(evStat, .Num, .Level, .Stat(evStat).EV, .Stat(evStat).IV, .Nature)
                    GivePlayerPokemonEVExp = Exp
                Else
                    Sobra = statMaxEv - .Stat(evStat).EV
                    .Stat(evStat).EV = statMaxEv
                    .Stat(evStat).Value = CalculatePokemonStat(evStat, .Num, .Level, .Stat(evStat).EV, .Stat(evStat).IV, .Nature)
                    GivePlayerPokemonEVExp = Sobra
                End If
            ElseIf Exp < 0 Then    ' Valor Negativo
                If .Stat(evStat).EV + Exp >= 0 Then
                    .Stat(evStat).EV = .Stat(evStat).EV + Exp
                    Sobra = -Exp    ' Sobra é a quantidade retirada como um número positivo
                Else
                    Sobra = .Stat(evStat).EV
                    .Stat(evStat).EV = 0
                End If

                .Stat(evStat).Value = CalculatePokemonStat(evStat, .Num, .Level, .Stat(evStat).EV, .Stat(evStat).IV, .Nature)
                GivePlayerPokemonEVExp = -Sobra
            End If
        Else
            If Exp > 0 Then    ' Valor Positivo
                ' Obtem na variável o que falta pra chegar no MAX_EV
                Sobra = MAX_EV - CountStat

                If .Stat(evStat).EV + Sobra <= statMaxEv Then
                    .Stat(evStat).EV = .Stat(evStat).EV + Sobra
                    .Stat(evStat).Value = CalculatePokemonStat(evStat, .Num, .Level, .Stat(evStat).EV, .Stat(evStat).IV, .Nature)
                    GivePlayerPokemonEVExp = Sobra
                Else
                    .Stat(evStat).EV = statMaxEv
                    .Stat(evStat).Value = CalculatePokemonStat(evStat, .Num, .Level, .Stat(evStat).EV, .Stat(evStat).IV, .Nature)
                    GivePlayerPokemonEVExp = Sobra
                End If
            ElseIf Exp < 0 Then    ' Valor Negativo
                Sobra = .Stat(evStat).EV
                .Stat(evStat).EV = 0

                .Stat(evStat).Value = CalculatePokemonStat(evStat, .Num, .Level, .Stat(evStat).EV, .Stat(evStat).IV, .Nature)
                GivePlayerPokemonEVExp = -Sobra
            End If
        End If

        ' Atualizações se for EV tipo HP
        If evStat = HP Then
            If Not .Stat(evStat).Value = .MaxHp Then
                .MaxHp = .Stat(evStat).Value
                SendPlayerPokemonSlot index, PokeSlot
            End If
        End If

        SendPlayerPokemonsStat index, PokeSlot

    End With
End Function

Private Sub CheckPlayerPokemonLevelUp(ByVal index As Long, ByVal PokeSlot As Byte)
    Dim ExpRollover As Long
    Dim statNu As Byte
    Dim oldlevel As Long, levelcount As Long
    Dim i As Long
    Dim DidLevel As Boolean

    '//Check Error
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If PokeSlot <= 0 Or PokeSlot > MAX_PLAYER_POKEMON Then Exit Sub
    If PlayerPokemons(index).Data(PokeSlot).Num <= 0 Then Exit Sub

    '//Add Exp
    With PlayerPokemons(index).Data(PokeSlot)
        levelcount = 0
        oldlevel = .Level
        DidLevel = False
        Do While .CurExp >= GetPokemonNextExp(.Level, Pokemon(.Num).GrowthRate)
            ExpRollover = .CurExp - GetPokemonNextExp(.Level, Pokemon(.Num).GrowthRate)

            .CurExp = ExpRollover
            '//Level Up
            .Level = .Level + 1
            levelcount = levelcount + 1
            DidLevel = True

            '//Calculate new stat
            For statNu = 1 To StatEnum.Stat_Count - 1
                .Stat(statNu).Value = CalculatePokemonStat(statNu, .Num, .Level, .Stat(statNu).EV, .Stat(statNu).IV, .Nature)
            Next
            .MaxHp = .Stat(StatEnum.HP).Value
        Loop
        '//Send Update
        SendPlayerPokemonSlot index, PokeSlot

        '//Check New Move
        If levelcount > 0 Then
            SendPlaySound "levelup.wav", Player(index, TempPlayer(index).UseChar).Map
            SendPlayerPokemonVital index
            CheckNewMove index, PokeSlot
        End If

        'SendPlayerPokemonVital Index
    End With
End Sub

Public Function FindFreeMoveSlot(ByVal index As Long, ByVal PokeSlot As Byte, Optional ByVal MoveSlot As Byte = 0) As Long
    Dim i As Byte
    Dim foundsameslot As Boolean

    '//Check Error
    If Not IsPlaying(index) Then Exit Function
    If TempPlayer(index).UseChar <= 0 Then Exit Function
    If PokeSlot <= 0 Or PokeSlot > MAX_PLAYER_POKEMON Then Exit Function
    If PlayerPokemons(index).Data(PokeSlot).Num <= 0 Then Exit Function

    foundsameslot = False
    With PlayerPokemons(index).Data(PokeSlot)
        For i = 1 To MAX_MOVESET
            If .Moveset(i).Num = 0 Then
                'If MoveSlot > 0 Then
                '    If .Moveset(i).Num = MoveSlot Then
                '        foundsameslot = True
                '    End If
                '    If Not foundsameslot Then
                '        FindFreeMoveSlot = i
                '        Exit Function
                '    Else
                '        FindFreeMoveSlot = -1
                '    End If
                'Else
                FindFreeMoveSlot = i
                Exit Function
                'End If
            End If
        Next
    End With
End Function

Public Sub CheckNewMove(ByVal index As Long, ByVal PokeSlot As Byte, Optional ByVal StartIndex As Long = 1)
    Dim i As Byte, x As Byte
    Dim FoundMatch As Boolean
    Dim MoveSlot As Byte
    Dim Continue As Boolean

    '//Check Error
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If PokeSlot <= 0 Or PokeSlot > MAX_PLAYER_POKEMON Then Exit Sub
    If PlayerPokemons(index).Data(PokeSlot).Num <= 0 Then Exit Sub
    If TempPlayer(index).MoveLearnNum > 0 Then Exit Sub
    If StartIndex <= 0 Then Exit Sub

    '//Add Exp
    With PlayerPokemons(index).Data(PokeSlot)
        '//Check New Move
        For i = StartIndex To MAX_POKEMON_MOVESET
            If Pokemon(.Num).Moveset(i).MoveNum > 0 Then
                If Pokemon(.Num).Moveset(i).MoveLevel = .Level Then
                    Continue = False
                    '//Make sure move doesn't exist
                    For x = 1 To MAX_MOVESET
                        If .Moveset(x).Num = Pokemon(.Num).Moveset(i).MoveNum Then
                            Continue = True
                        End If
                    Next
                    If Not Continue Then
                        '//Check if there's available slot
                        MoveSlot = FindFreeMoveSlot(index, PokeSlot)
                        If MoveSlot >= 0 Then
                            If MoveSlot > 0 Then
                                .Moveset(MoveSlot).Num = Pokemon(.Num).Moveset(i).MoveNum
                                .Moveset(MoveSlot).TotalPP = PokemonMove(Pokemon(.Num).Moveset(i).MoveNum).PP
                                .Moveset(MoveSlot).CurPP = .Moveset(MoveSlot).TotalPP
                                SendPlayerPokemonSlot index, PokeSlot
                                '//Send Msg
                                SendPlayerMsg index, Trim$(Pokemon(.Num).Name) & " learned the move " & Trim$(PokemonMove(Pokemon(.Num).Moveset(i).MoveNum).Name), White
                            Else
                                '//Proceed to ask
                                TempPlayer(index).MoveLearnPokeSlot = PokeSlot
                                TempPlayer(index).MoveLearnNum = Pokemon(.Num).Moveset(i).MoveNum
                                TempPlayer(index).MoveLearnIndex = i + 1
                                SendNewMove index
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End With
End Sub

Private Function CheckItemCooldown(ByVal index As Long, ByVal InvSlot As Long) As Boolean
    Dim ItemNum As Long
    Dim currentTime As Long
    Dim cooldownTime As Long

    ItemNum = PlayerInv(index).Data(InvSlot).Num

    If ItemNum > 0 And ItemNum <= MAX_ITEM Then
        currentTime = GetTickCount
        cooldownTime = PlayerInv(index).Data(InvSlot).TmrCooldown

        If cooldownTime > currentTime Then
            If cooldownTime > (currentTime + Item(ItemNum).Delay) Then
                PlayerInv(index).Data(InvSlot).TmrCooldown = 0
                CheckItemCooldown = True
            End If
        Else
            CheckItemCooldown = True
        End If
    End If
End Function

Private Function GetItemCooldownSegs(ByVal index As Long, ByVal InvSlot As Long) As Long
    Dim ItemNum As Long, CD As Long, remainingTime As Long

    ItemNum = PlayerInv(index).Data(InvSlot).Num

    If ItemNum > 0 And ItemNum <= MAX_ITEM Then
        CD = PlayerInv(index).Data(InvSlot).TmrCooldown

        If CD > 0 Then
            remainingTime = (CD - GetTickCount) \ 1000

            If remainingTime >= 1 Then
                GetItemCooldownSegs = remainingTime
            End If
        End If
    End If
End Function

Public Sub PlayerUseItem(ByVal index As Long, ByVal InvSlot As Byte)
    Dim ItemNum As Long
    Dim gothealed As Boolean
    Dim x As Long
    Dim exproll As Long
    Dim Exp As Long
    Dim i As Long, CanLearn As Boolean
    Dim BerriesFunc As Integer, PokeName As String
    Dim TAKE As Boolean

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If InvSlot <= 0 Or InvSlot > MAX_PLAYER_INV Then Exit Sub
    If PlayerInv(index).Data(InvSlot).Num <= 0 Then Exit Sub
    If PlayerInv(index).Data(InvSlot).Value <= 0 Then Exit Sub
    If TempPlayer(index).InDuel > 0 Then Exit Sub
    If TempPlayer(index).InNpcDuel > 0 Then Exit Sub

    ItemNum = PlayerInv(index).Data(InvSlot).Num

    '//Verificar se o item está em cooldown
    If CheckItemCooldown(index, InvSlot) = False Then
        'Select Case TempPlayer(index).CurLanguage
        'Case LANG_PT: AddAlert index, "Item em cooldown, aguarde: " & GetItemCooldownSegs(index, InvSlot), White
        'Case LANG_EN: AddAlert index, "Item em cooldown, aguarde: " & GetItemCooldownSegs(index, InvSlot), White
        'Case LANG_ES: AddAlert index, "Item em cooldown, aguarde: " & GetItemCooldownSegs(index, InvSlot), White
        'End Select
        Call SendPlayerInvSlot(index, InvSlot)
        Exit Sub
    End If

    Select Case Item(ItemNum).Type

    Case ItemTypeEnum.pokeBall
        '//Catching
        If Map(Player(index, TempPlayer(index).UseChar).Map).Moral = 3 Then
            If Not ItemNum = 12 Then
                AddAlert index, "You cannot use this type of Pokeball here", White
                Exit Sub
            End If
        Else
            If ItemNum = 12 Then
                AddAlert index, "You cannot use this type of Pokeball here", White
                Exit Sub
            End If
        End If
        TempPlayer(index).TmpUseInvSlot = InvSlot
        SendGetData index, ItemTypeEnum.pokeBall, InvSlot
    Case ItemTypeEnum.Medicine

        '//Não pode curar em certos mapas
        If Map(GetPlayerMap(index)).NoCure = YES Then
            Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: AddAlert index, "Não permitido usar medicine aqui", White
            Case LANG_EN: AddAlert index, "Não permitido usar medicine aqui", White
            Case LANG_ES: AddAlert index, "Não permitido usar medicine aqui", White
            End Select
            Exit Sub
        End If

        Select Case Item(ItemNum).Data1    '//Type
        Case 1    '// Heal HP
            gothealed = False
            If PlayerPokemon(index).Num > 0 Then
                If PlayerPokemon(index).slot > 0 Then
                    If PlayerPokemons(index).Data(PlayerPokemon(index).slot).CurHp < PlayerPokemons(index).Data(PlayerPokemon(index).slot).MaxHp Then
                        PlayerPokemons(index).Data(PlayerPokemon(index).slot).CurHp = PlayerPokemons(index).Data(PlayerPokemon(index).slot).CurHp + Item(ItemNum).Data2
                        If PlayerPokemons(index).Data(PlayerPokemon(index).slot).CurHp > PlayerPokemons(index).Data(PlayerPokemon(index).slot).MaxHp Then
                            PlayerPokemons(index).Data(PlayerPokemon(index).slot).CurHp = PlayerPokemons(index).Data(PlayerPokemon(index).slot).MaxHp
                        End If
                        gothealed = True
                    End If
                End If
            End If
            If gothealed Then
                Select Case TempPlayer(index).CurLanguage
                Case LANG_PT: AddAlert index, "Pokemon HP restored", White
                Case LANG_EN: AddAlert index, "Pokemon HP restored", White
                Case LANG_ES: AddAlert index, "Pokemon HP restored", White
                End Select
                SendPlayerPokemonVital index

                TAKE = True
            End If

        Case 2    '// Give Exp
        Case 3    '// Heal PP
            gothealed = False
            If PlayerPokemon(index).Num > 0 Then
                If PlayerPokemon(index).slot > 0 Then
                    For x = 1 To MAX_MOVESET
                        If PlayerPokemons(index).Data(PlayerPokemon(index).slot).Moveset(x).Num > 0 Then
                            If PlayerPokemons(index).Data(PlayerPokemon(index).slot).Moveset(x).CurPP < PlayerPokemons(index).Data(PlayerPokemon(index).slot).Moveset(x).TotalPP Then
                                PlayerPokemons(index).Data(PlayerPokemon(index).slot).Moveset(x).CurPP = PlayerPokemons(index).Data(PlayerPokemon(index).slot).Moveset(x).CurPP + Item(ItemNum).Data2
                                If PlayerPokemons(index).Data(PlayerPokemon(index).slot).Moveset(x).CurPP > PlayerPokemons(index).Data(PlayerPokemon(index).slot).Moveset(x).TotalPP Then
                                    PlayerPokemons(index).Data(PlayerPokemon(index).slot).Moveset(x).CurPP = PlayerPokemons(index).Data(PlayerPokemon(index).slot).Moveset(x).TotalPP
                                End If
                                PlayerPokemons(index).Data(PlayerPokemon(index).slot).Moveset(x).CD = 0
                                gothealed = True
                            End If
                        End If
                    Next
                End If
            End If
            If gothealed Then
                Select Case TempPlayer(index).CurLanguage
                Case LANG_PT: AddAlert index, "Pokemon PP restored", White
                Case LANG_EN: AddAlert index, "Pokemon PP restored", White
                Case LANG_ES: AddAlert index, "Pokemon PP restored", White
                End Select
                For x = 1 To MAX_MOVESET
                    SendPlayerPokemonPP index, x
                Next

                TAKE = True
            End If
        Case 4    '// Revive
            TempPlayer(index).TmpUseInvSlot = InvSlot
            SendGetData index, ItemTypeEnum.Medicine, InvSlot
        Case 5    '// Cure Status
            gothealed = False
            If Item(ItemNum).Data2 > 0 Then
                If PlayerPokemon(index).Num > 0 Then
                    If PlayerPokemon(index).slot > 0 Then
                        If PlayerPokemons(index).Data(PlayerPokemon(index).slot).Status = Item(ItemNum).Data2 Then
                            PlayerPokemons(index).Data(PlayerPokemon(index).slot).Status = 0
                            gothealed = True
                        End If
                    End If
                End If
            Else
                If PlayerPokemon(index).slot > 0 Then
                    If PlayerPokemons(index).Data(PlayerPokemon(index).slot).Status > 0 Then
                        PlayerPokemons(index).Data(PlayerPokemon(index).slot).Status = 0
                        gothealed = True
                    End If
                End If
            End If
            If gothealed Then
                Select Case TempPlayer(index).CurLanguage
                Case LANG_PT: AddAlert index, "Pokemon Status removed", White
                Case LANG_EN: AddAlert index, "Pokemon Status removed", White
                Case LANG_ES: AddAlert index, "Pokemon Status removed", White
                End Select
                SendPlayerPokemonStatus index

                TAKE = True
            End If

        Case 6    '// Heal Trainer
            gothealed = False
            If Player(index, TempPlayer(index).UseChar).CurHp < GetPlayerHP(Player(index, TempPlayer(index).UseChar).Level) Then
                Player(index, TempPlayer(index).UseChar).CurHp = Player(index, TempPlayer(index).UseChar).CurHp + Item(PlayerInv(index).Data(InvSlot).Num).Data2
                If Player(index, TempPlayer(index).UseChar).CurHp > GetPlayerHP(Player(index, TempPlayer(index).UseChar).Level) Then
                    Player(index, TempPlayer(index).UseChar).CurHp = GetPlayerHP(Player(index, TempPlayer(index).UseChar).Level)
                End If
                gothealed = True
            End If

            If gothealed Then
                Select Case TempPlayer(index).CurLanguage
                Case LANG_PT: AddAlert index, "HP restored", White
                Case LANG_EN: AddAlert index, "HP restored", White
                Case LANG_ES: AddAlert index, "HP restored", White
                End Select
                SendPlayerVital index
                TAKE = True
            End If
        Case 7    '// Cure Trainer
            gothealed = False
            If Item(ItemNum).Data2 > 0 Then
                If Player(index, TempPlayer(index).UseChar).Status = Item(ItemNum).Data2 Then
                    Player(index, TempPlayer(index).UseChar).Status = 0
                    gothealed = True
                End If
            Else
                If Player(index, TempPlayer(index).UseChar).Status > 0 Then
                    Player(index, TempPlayer(index).UseChar).Status = 0
                    gothealed = True
                End If
            End If
            If gothealed Then
                Select Case TempPlayer(index).CurLanguage
                Case LANG_PT: AddAlert index, "Status was removed", White
                Case LANG_EN: AddAlert index, "Status was removed", White
                Case LANG_ES: AddAlert index, "Status was removed", White
                End Select
                SendPlayerStatus index
                TAKE = True
            End If

        End Select

        If Item(ItemNum).Data3 > 0 Then
            '//Level Up
            If PlayerPokemon(index).Num > 0 Then
                If PlayerPokemon(index).slot > 0 Then
                    exproll = GetPokemonNextExp(PlayerPokemons(index).Data(PlayerPokemon(index).slot).Level, Pokemon(PlayerPokemons(index).Data(PlayerPokemon(index).slot).Num).GrowthRate)
                    Exp = exproll - PlayerPokemons(index).Data(PlayerPokemon(index).slot).CurExp

                    If Exp > 0 Then
                        GivePlayerPokemonExp index, PlayerPokemon(index).slot, Exp
                    End If

                    TAKE = True
                End If
            End If
        End If
    Case ItemTypeEnum.Berries

        If Item(ItemNum).Data1 > 0 Then
            If PlayerPokemon(index).Num > 0 Then
                If PlayerPokemon(index).slot > 0 Then
                    For i = 1 To StatEnum.Stat_Count - 1
                        If Item(ItemNum).Data1 = i Then
                            ' Adiciona ou remove a experiência (Berries/Proteins)
                            BerriesFunc = GivePlayerPokemonEVExp(index, PlayerPokemon(index).slot, Item(ItemNum).Data1, Item(ItemNum).Data2)
                            If BerriesFunc <> 0 Then
                                TAKE = True

                                PokeName = Trim$(Pokemon(PlayerPokemon(index).Num).Name)
                                If BerriesFunc > 0 Then
                                    Select Case TempPlayer(index).CurLanguage
                                    Case LANG_PT: AddAlert index, PokeName & " aumentou " & BerriesFunc & " pontos de EV em " & GetAtributeName(Item(ItemNum).Data1), Green
                                    Case LANG_EN: AddAlert index, PokeName & " aumentou " & BerriesFunc & " pontos de EV em " & GetAtributeName(Item(ItemNum).Data1), Green
                                    Case LANG_ES: AddAlert index, PokeName & " aumentou " & BerriesFunc & " pontos de EV em " & GetAtributeName(Item(ItemNum).Data1), Green
                                    End Select
                                ElseIf BerriesFunc < 0 Then
                                    Select Case TempPlayer(index).CurLanguage
                                    Case LANG_PT: AddAlert index, PokeName & " reduziu " & Math.Abs(BerriesFunc) & " pontos de EV em " & GetAtributeName(Item(ItemNum).Data1), Grey
                                    Case LANG_EN: AddAlert index, PokeName & " reduziu " & Math.Abs(BerriesFunc) & " pontos de EV em " & GetAtributeName(Item(ItemNum).Data1), Grey
                                    Case LANG_ES: AddAlert index, PokeName & " reduziu " & Math.Abs(BerriesFunc) & " pontos de EV em " & GetAtributeName(Item(ItemNum).Data1), Grey
                                    End Select
                                End If
                            Else
                                Select Case TempPlayer(index).CurLanguage
                                Case LANG_PT: AddAlert index, PokeName & " está no limite Min/Max de EV em " & GetAtributeName(Item(ItemNum).Data1) & " " & PlayerPokemons(index).Data(PlayerPokemon(index).slot).Stat(Item(ItemNum).Data1).EV, Grey
                                Case LANG_EN: AddAlert index, PokeName & " está no limite Min/Max de EV em " & GetAtributeName(Item(ItemNum).Data1) & " " & PlayerPokemons(index).Data(PlayerPokemon(index).slot).Stat(Item(ItemNum).Data1).EV, Grey
                                Case LANG_ES: AddAlert index, PokeName & " está no limite Min/Max de EV em " & GetAtributeName(Item(ItemNum).Data1) & " " & PlayerPokemons(index).Data(PlayerPokemon(index).slot).Stat(Item(ItemNum).Data1).EV, Grey
                                End Select
                                Exit Sub
                            End If
                            Exit For
                        End If
                    Next i
                Else
                    Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "Você não está em um pokemon", White
                    Case LANG_EN: AddAlert index, "You are not in a pokemon", White
                    Case LANG_ES: AddAlert index, "No estas en un pokemon", White
                    End Select
                    Exit Sub
                End If
            Else
                Select Case TempPlayer(index).CurLanguage
                Case LANG_PT: AddAlert index, "Você não está em um pokemon", White
                Case LANG_EN: AddAlert index, "You are not in a pokemon", White
                Case LANG_ES: AddAlert index, "No estas en un pokemon", White
                End Select
                Exit Sub
            End If
        End If

    Case ItemTypeEnum.keyItems
        Select Case Item(ItemNum).Data1
        Case 1    '//Sprite Type
            If Item(ItemNum).Data2 > 0 And Item(ItemNum).Data2 < TEMP_SPRITE_GROUP_MOUNT Then
                If Map(Player(index, TempPlayer(index).UseChar).Map).SpriteType <= TEMP_SPRITE_GROUP_NONE Then
                    ChangeTempSprite index, Item(ItemNum).Data2, ItemNum
                End If
            ElseIf Item(ItemNum).Data2 > 0 And Item(ItemNum).Data2 = TEMP_SPRITE_GROUP_MOUNT Then
                If Map(Player(index, TempPlayer(index).UseChar).Map).SpriteType <= TEMP_SPRITE_GROUP_NONE Then
                    ChangeTempSprite index, Item(ItemNum).Data2, ItemNum
                End If
            ElseIf Item(ItemNum).Data2 > 0 And Item(ItemNum).Data2 = TEMP_FISH_MODE Then
                'If GetPlayerFishMode(Index) = NO Then
                Select Case GetPlayerDir(index)
                Case DIR_DOWN
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Attribute = MapAttribute.FishSpot Then
                        x = YES
                    End If
                Case DIR_UP
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Attribute = MapAttribute.FishSpot Then
                        x = YES
                    End If
                Case DIR_LEFT
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Attribute = MapAttribute.FishSpot Then
                        x = YES
                    End If
                Case DIR_RIGHT
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Attribute = MapAttribute.FishSpot Then
                        x = YES
                    End If
                End Select
                If x = YES Then
                    If GetPlayerFishMode(index) = NO Then
                        SetPlayerFishMode index, YES
                        SetPlayerFishRod index, Item(ItemNum).Data3
                        SendActionMsg GetPlayerMap(index), "Fishing Mode!", Player(index, TempPlayer(index).UseChar).x * 32, Player(index, TempPlayer(index).UseChar).Y * 32, Yellow
                        SendFishMode index
                    Else
                        SpawnPokemonIdInMap index, GetPlayerMap(index)
                        SendActionMsg GetPlayerMap(index), "Fishing!", Player(index, TempPlayer(index).UseChar).x * 32, Player(index, TempPlayer(index).UseChar).Y * 32, Green
                    End If
                Else
                    Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "Não está em frente a um local de pesca!", White
                    Case LANG_EN: AddAlert index, "Não está em frente a um local de pesca!", White
                    Case LANG_ES: AddAlert index, "Não está em frente a um local de pesca!", White
                    End Select
                    Exit Sub
                End If
            End If
        End Select
    Case ItemTypeEnum.TM_HM
        If PlayerPokemon(index).Num > 0 And PlayerPokemon(index).slot > 0 Then
            If Item(ItemNum).Data1 > 0 Then
                CanLearn = False
                For i = 1 To 110
                    If Pokemon(PlayerPokemon(index).Num).ItemMoveset(i) = Item(ItemNum).Data1 Then
                        CanLearn = True
                        Exit For
                    End If
                Next
                '//Make sure move doesn't exist
                For i = 1 To MAX_MOVESET
                    If PlayerPokemons(index).Data(PlayerPokemon(index).slot).Moveset(i).Num = Item(ItemNum).Data1 Then
                        CanLearn = False
                    End If
                Next

                If CanLearn Then
                    '//Continue
                    i = FindFreeMoveSlot(index, PlayerPokemon(index).slot)
                    If i > 0 Then
                        PlayerPokemons(index).Data(PlayerPokemon(index).slot).Moveset(i).Num = Item(ItemNum).Data1
                        PlayerPokemons(index).Data(PlayerPokemon(index).slot).Moveset(i).TotalPP = PokemonMove(Item(ItemNum).Data1).PP
                        PlayerPokemons(index).Data(PlayerPokemon(index).slot).Moveset(i).CurPP = PlayerPokemons(index).Data(PlayerPokemon(index).slot).Moveset(i).TotalPP
                        SendPlayerPokemonSlot index, PlayerPokemon(index).slot
                        '//Send Msg
                        SendPlayerMsg index, Trim$(Pokemon(PlayerPokemon(index).Num).Name) & " learned the move " & Trim$(PokemonMove(Item(ItemNum).Data1).Name), White
                    Else
                        '//Proceed to ask
                        TempPlayer(index).MoveLearnPokeSlot = PlayerPokemon(index).slot
                        TempPlayer(index).MoveLearnNum = Item(ItemNum).Data1
                        TempPlayer(index).MoveLearnIndex = 0
                        SendNewMove index
                    End If

                    If Item(ItemNum).Data2 > 0 Then
                        TAKE = True
                    End If
                Else
                    AddAlert index, "This pokemon cannot learn this move", White
                    Exit Sub
                End If
            End If
        Else
            AddAlert index, "Please spawn your pokemon first", White
            Exit Sub
        End If
    Case ItemTypeEnum.PowerBracer

    Case ItemTypeEnum.Items
        TempPlayer(index).StorageType = Item(ItemNum).Data1
        AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " enter storage"
        SendStorage index
    Case ItemTypeEnum.MysteryBox
        Dim Received As Boolean
        x = ObterItem(Item(ItemNum).Item, Item(ItemNum).ItemValue, Item(ItemNum).ItemChance)
        If x > 0 Then
            If TryGivePlayerItem(index, Item(ItemNum).Item(x), Item(ItemNum).ItemValue(x), Item(ItemNum).ItemChance(x)) Then
                Select Case TempPlayer(index).CurLanguage
                Case LANG_PT: AddAlert index, "Parabens, você recebeu " & Item(ItemNum).ItemValue(x) & "x " & Trim$(Item(Item(ItemNum).Item(x)).Name), White
                Case LANG_EN: AddAlert index, "Parabens, você recebeu " & Item(ItemNum).ItemValue(x) & "x " & Trim$(Item(Item(ItemNum).Item(x)).Name), White
                Case LANG_ES: AddAlert index, "Parabens, você recebeu " & Item(ItemNum).ItemValue(x) & "x " & Trim$(Item(Item(ItemNum).Item(x)).Name), White
                End Select

                Received = True
            End If
            '//Take Item
            TAKE = True
        End If

    Case ItemTypeEnum.Vip
        If AddVip(index, Item(ItemNum).Data1, Item(ItemNum).Data2) Then
            '//Take Item
            TAKE = True
        End If
    Case Else
        '//Not usable
        Exit Sub
    End Select

    If TAKE = True Then
        '//Take Item
        PlayerInv(index).Data(InvSlot).Value = PlayerInv(index).Data(InvSlot).Value - 1
        If PlayerInv(index).Data(InvSlot).Value <= 0 Then
            '//Clear Item
            PlayerInv(index).Data(InvSlot).Num = 0
            PlayerInv(index).Data(InvSlot).Value = 0
            PlayerInv(index).Data(InvSlot).TmrCooldown = 0
        End If
    Else
        '//Set Cooldown
        PlayerInv(index).Data(InvSlot).TmrCooldown = GetTickCount + Item(ItemNum).Delay
    End If


    SendPlayerInvSlot index, InvSlot

    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " use item " & Trim$(Item(ItemNum).Name)
End Sub

Function ObterItem(Item() As Integer, Quant() As Long, Chance() As Double) As Long
    Dim totalChances As Double
    Dim i As Integer
    
    totalChances = 0#

    ObterItem = 0

    ' Calcular o total de chances disponíveis
    For i = 1 To MAX_MYSTERY_BOX
        ' Verificar se o índice do item é maior que 0
        If Item(i) > 0 Then
            totalChances = totalChances + Chance(i)
        End If
    Next i

    ' Verificar se não há chances disponíveis
    If totalChances <= 0# Then
        Exit Function
    End If

    ' Gerar um número aleatório dentro do intervalo total de chances
    Dim numeroAleatorio As Double
    numeroAleatorio = CDbl(FormatNumber((totalChances * Rnd) + 1, 2))

    ' Percorrer os arrays para encontrar o item correspondente ao número aleatório
    Dim somaChances As Integer
    For i = 1 To MAX_MYSTERY_BOX
        ' Verificar se o índice do item é maior que 0
        If Item(i) > 0 Then
            somaChances = somaChances + Chance(i)
            
            ' Se a soma das chances for maior ou igual ao número aleatório,
            ' o jogador recebe o item correspondente
            If somaChances >= numeroAleatorio Then
                ' Faça o que precisar com o item recebido, por exemplo:
                ' Exibir uma mensagem, atualizar algum registro, etc.
                ObterItem = i
                Exit Function
            End If
        End If
    Next i
End Function

'//Count Free Pokemno slot
Public Function CountFreePokemonSlot(ByVal index As Long) As Long
    Dim count As Long
    Dim i As Byte, x As Byte

    count = 0
    For i = 1 To MAX_PLAYER_POKEMON
        If PlayerPokemons(index).Data(i).Num = 0 Then
            count = count + 1
        End If
    Next
    For i = 1 To MAX_STORAGE_SLOT
        If PlayerPokemonStorage(index).slot(i).Unlocked = YES Then
            For x = 1 To MAX_STORAGE
                If PlayerPokemonStorage(index).slot(i).Data(x).Num = 0 Then
                    count = count + 1
                End If
            Next
        End If
    Next
    CountFreePokemonSlot = count
End Function

Public Function FindSameInvStorageSlot(ByVal index As Long, ByVal StorageSlot As Byte, ByVal ItemNum As Long) As Byte
    Dim i As Byte

    FindSameInvStorageSlot = 0

    If Not IsPlaying(index) Then Exit Function
    If TempPlayer(index).UseChar <= 0 Then Exit Function

    If ItemNum <= 0 Then Exit Function

    For i = 1 To MAX_STORAGE
        With PlayerInvStorage(index).slot(StorageSlot)
            If .Unlocked = YES Then
                If .Data(i).Num = ItemNum Then
                    If Item(ItemNum).Stock = YES Then
                        '//add val
                        FindSameInvStorageSlot = i
                        Exit Function
                    End If
                End If
            End If
        End With
    Next
End Function

Private Function FindFreeInvStorageSlot(ByVal index As Long, ByVal StorageSlot As Byte, ByVal ItemNum As Long, ByRef ItemVal As Long, Optional ByRef MsgFrom As String) As Byte
    Dim i As Byte

    FindFreeInvStorageSlot = 0

    If Item(ItemNum).Stock = YES Then
        i = FindSameInvStorageSlot(index, StorageSlot, ItemNum)
        If i > 0 Then
            If CheckStorageValues(index, StorageSlot, i, ItemNum, ItemVal, MsgFrom) Then
                FindFreeInvStorageSlot = i
            End If

            Exit Function
        End If
    End If

    For i = 1 To MAX_STORAGE
        With PlayerInvStorage(index).slot(StorageSlot).Data(i)
            If .Num = 0 Then
                If CheckStorageValues(index, StorageSlot, i, ItemNum, ItemVal, MsgFrom) Then
                    FindFreeInvStorageSlot = i
                End If

                Exit Function
            End If
        End With
    Next
End Function

Public Function checkItem(ByVal index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    For i = 1 To MAX_PLAYER_INV
        If PlayerInv(index).Data(i).Num = ItemNum Then
            checkItem = i
            Exit Function
        End If
    Next
    checkItem = 0
End Function

Private Function CheckStorageValues(ByVal index As Long, ByVal StorageSlot As Long, ByVal Data As Long, ByVal ItemNum As Long, _
                                    ByRef ItemVal As Long, Optional ByRef MsgFrom As String) As Boolean

    CheckStorageValues = True

    With PlayerInvStorage(index).slot(StorageSlot).Data(Data)

        '//Adaptações pra não ultrapassar um limite de Amount
        If ItemVal > MAX_AMOUNT Then
            Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: MsgFrom = "Quantidade Limite " & MAX_AMOUNT
            Case LANG_EN: MsgFrom = "Limit Quantity " & MAX_AMOUNT
            Case LANG_ES: MsgFrom = "Limit Quantity " & MAX_AMOUNT
            End Select

            CheckStorageValues = False
            Exit Function
        End If

        If .Value >= MAX_AMOUNT Then
            Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: MsgFrom = "Quantidade Limite " & MAX_AMOUNT
            Case LANG_EN: MsgFrom = "Limit Quantity " & MAX_AMOUNT
            Case LANG_ES: MsgFrom = "Limit Quantity " & MAX_AMOUNT
            End Select

            CheckStorageValues = False
            Exit Function
        End If

        If (ItemVal + .Value) > MAX_AMOUNT Then
            '//Altera o valor pra obter apenas o que couber
            ItemVal = MAX_AMOUNT - .Value

            Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: MsgFrom = "Quantidade Excedida, você recebeu apenas (" & ItemVal & ")"
            Case LANG_EN: MsgFrom = "Quantity Exceeded, you have only received (" & ItemVal & ")"
            Case LANG_ES: MsgFrom = "Quantity Exceeded, you have only received (" & ItemVal & ")"
            End Select

            CheckStorageValues = True
            Exit Function
        End If

    End With
End Function

Public Function CountFreeStorageSlot(ByVal index As Long, ByVal StorageSlot As Long) As Long
    Dim count As Long, i As Long

    CountFreeStorageSlot = 0
    count = 0

    If Not IsPlaying(index) Then Exit Function
    If TempPlayer(index).UseChar <= 0 Then Exit Function

    For i = 1 To MAX_STORAGE
        With PlayerInvStorage(index).slot(StorageSlot).Data(i)
            If .Num = 0 Then
                count = count + 1
            End If
        End With
    Next

    CountFreeStorageSlot = count
End Function


Public Function TryGiveStorageItem(ByVal index As Long, ByVal StorageSlot As Byte, ByVal ItemNum As Long, ByRef ItemVal As Long, Optional ByVal ItemCooldown As Long = 0, Optional ByRef MsgFrom As String) As Boolean
    TryGiveStorageItem = True

    If Not GiveStorageItem(index, StorageSlot, ItemNum, ItemVal, ItemCooldown, MsgFrom) Then

        If MsgFrom = vbNullString Then
            Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: MsgFrom = "Storage is full"
            Case LANG_EN: MsgFrom = "Storage is full"
            Case LANG_ES: MsgFrom = "Storage is full"
            End Select
        End If

        AddAlert index, MsgFrom, White
        TryGiveStorageItem = False
    Else
        '//Check if there's still free slot
        If CountFreeStorageSlot(index, StorageSlot) <= 5 Then
            Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: AddAlert index, "Warning: Your storage is almost full", White
            Case LANG_EN: AddAlert index, "Warning: Your storage is almost full", White
            Case LANG_ES: AddAlert index, "Warning: Your storage is almost full", White
            End Select
        End If
    End If
End Function

Public Function GiveStorageItem(ByVal index As Long, ByVal StorageSlot As Byte, ByVal ItemNum As Long, ByRef ItemVal As Long, Optional ByVal ItemCooldown As Long = 0, Optional ByRef MsgFrom As String) As Boolean
    Dim i As Byte

    '//Verifica se tem um slot pra adicionar o item
    i = FindFreeInvStorageSlot(index, StorageSlot, ItemNum, ItemVal, MsgFrom)
    If i > 0 Then
        With PlayerInvStorage(index).slot(StorageSlot).Data(i)
            .Num = ItemNum
            .Value = .Value + ItemVal
            .TmrCooldown = ItemCooldown
            GiveStorageItem = True

            '//Update
            SendPlayerInvStorageSlot index, StorageSlot, i
        End With
    Else
        Exit Function
    End If

End Function

Public Sub ProcessConversation(ByVal index As Long, ByVal Convo As Long, ByVal ConvoData As Byte, Optional ByVal NpcNum As Long = 0, Optional ByVal tReply As Byte = 0)
    Dim i As Long, x As Long
    Dim fixData As Boolean

    fixData = False

startOver:

    If Convo <= 0 Or Convo > MAX_CONVERSATION Then Exit Sub

    If Not fixData Then
        If ConvoData <= 0 Then
            '//Initiate
            TempPlayer(index).CurConvoData = 1
        Else
            If Conversation(Convo).ConvData(ConvoData).NoReply = YES Then
                TempPlayer(index).CurConvoData = Conversation(Convo).ConvData(ConvoData).MoveNext
            Else
                If tReply > 0 And tReply <= 3 Then
                    TempPlayer(index).CurConvoData = Conversation(Convo).ConvData(ConvoData).tReplyMove(tReply)
                Else
                    TempPlayer(index).CurConvoData = 0  '//End
                End If
            End If
        End If
    End If
    ConvoData = TempPlayer(index).CurConvoData

    If ConvoData > 0 Then
        With Conversation(Convo).ConvData(ConvoData)
            '//Check for custom script
            Select Case .CustomScript
            Case CONVO_SCRIPT_INVSTORAGE    '//Inv Storage
                If TempPlayer(index).StorageType = 0 Then
                    TempPlayer(index).StorageType = 1
                    SendStorage index
                End If
                fixData = False
            Case CONVO_SCRIPT_POKESTORAGE    '//Pokemon Storage
                If TempPlayer(index).StorageType = 0 Then
                    TempPlayer(index).StorageType = 2
                    SendStorage index
                End If
                fixData = False
            Case CONVO_SCRIPT_HEAL
                '//Heal Pokemon
                For i = 1 To MAX_PLAYER_POKEMON
                    If PlayerPokemons(index).Data(i).Num > 0 Then
                        If PlayerPokemons(index).Data(i).CurHp < PlayerPokemons(index).Data(i).MaxHp Then
                            PlayerPokemons(index).Data(i).CurHp = PlayerPokemons(index).Data(i).MaxHp
                        End If
                        If PlayerPokemons(index).Data(i).Status > 0 Then
                            PlayerPokemons(index).Data(i).Status = 0
                        End If
                        For x = 1 To MAX_MOVESET
                            If PlayerPokemons(index).Data(i).Moveset(x).Num > 0 Then
                                If PlayerPokemons(index).Data(i).Moveset(x).CurPP < PlayerPokemons(index).Data(i).Moveset(x).TotalPP Then
                                    PlayerPokemons(index).Data(i).Moveset(x).CurPP = PlayerPokemons(index).Data(i).Moveset(x).TotalPP
                                    PlayerPokemons(index).Data(i).Moveset(x).CD = 0
                                End If
                            End If
                        Next
                    End If
                Next
                If Player(index, TempPlayer(index).UseChar).CurHp < GetPlayerHP(Player(index, TempPlayer(index).UseChar).Level) Then
                    Player(index, TempPlayer(index).UseChar).CurHp = GetPlayerHP(Player(index, TempPlayer(index).UseChar).Level)
                End If
                If Player(index, TempPlayer(index).UseChar).Status > 0 Then
                    Player(index, TempPlayer(index).UseChar).Status = 0
                    Player(index, TempPlayer(index).UseChar).IsConfuse = False
                End If
                Select Case TempPlayer(index).CurLanguage
                Case LANG_PT: AddAlert index, "Pokemon HP and PP restored", White
                Case LANG_EN: AddAlert index, "Pokemon HP and PP restored", White
                Case LANG_ES: AddAlert index, "Pokemon HP and PP restored", White
                End Select
                SendPlayerPokemons index
                SendPlayerVital index
                SendPlayerPokemonStatus index
                SendPlayerStatus index
                fixData = False
            Case CONVO_SCRIPT_SHOP
                If .CustomScriptData > 0 Then
                    '//Open Shop
                    If TempPlayer(index).InShop = 0 Then
                        TempPlayer(index).InShop = .CustomScriptData
                        SendOpenShop index
                    End If
                End If
                fixData = False
            Case CONVO_SCRIPT_SETSWITCH
                If .CustomScriptData > 0 Then
                    '//Open Shop
                    If IsPlaying(index) Then
                        If TempPlayer(index).UseChar > 0 Then
                            Player(index, TempPlayer(index).UseChar).Switches(.CustomScriptData) = .CustomScriptData2
                        End If
                    End If
                End If
                fixData = False
            Case CONVO_SCRIPT_GIVEPOKE
                If .CustomScriptData > 0 Then
                    If IsPlaying(index) Then
                        If TempPlayer(index).UseChar > 0 Then
                            GivePlayerPokemon index, .CustomScriptData, 5, BallEnum.b_Pokeball
                        End If
                    End If
                End If
                fixData = False
            Case CONVO_SCRIPT_GIVEITEM
                If .CustomScriptData > 0 Then
                    If IsPlaying(index) Then
                        If TempPlayer(index).UseChar > 0 Then
                            If .CustomScriptData2 > 0 Then
                                TryGivePlayerItem index, .CustomScriptData, .CustomScriptData2
                            End If
                        End If
                    End If
                End If
                fixData = False
            Case CONVO_SCRIPT_WARPTO
                If .CustomScriptData > 0 Then
                    If IsPlaying(index) Then
                        If TempPlayer(index).UseChar > 0 Then
                            PlayerWarp index, .CustomScriptData, .CustomScriptData2, .CustomScriptData3, Player(index, TempPlayer(index).UseChar).Dir
                        End If
                    End If
                End If
                fixData = False
            Case CONVO_SCRIPT_CHECKMONEY
                If .CustomScriptData > 0 Then
                    If IsPlaying(index) Then
                        If TempPlayer(index).UseChar > 0 Then
                            If Player(index, TempPlayer(index).UseChar).Money >= .CustomScriptData Then
                                '//Next
                                TempPlayer(index).CurConvoData = .CustomScriptData2
                                fixData = True
                            Else
                                TempPlayer(index).CurConvoData = .CustomScriptData3
                                fixData = True
                            End If
                        End If
                    End If
                End If
            Case CONVO_SCRIPT_TAKEMONEY
                If .CustomScriptData > 0 Then
                    If IsPlaying(index) Then
                        If TempPlayer(index).UseChar > 0 Then
                            Player(index, TempPlayer(index).UseChar).Money = Player(index, TempPlayer(index).UseChar).Money - .CustomScriptData
                            If Player(index, TempPlayer(index).UseChar).Money <= 0 Then
                                Player(index, TempPlayer(index).UseChar).Money = 0
                            End If
                            SendPlayerData index
                        End If
                    End If
                End If
                fixData = False
            Case CONVO_SCRIPT_STARTBATTLE
                If TempPlayer(index).CurConvoMapNpc > 0 Then

                    '//Npc not rebattle Option (Never Rebattle if Win)
                    If Player(index, TempPlayer(index).UseChar).NpcBattledDay(MapNpc(Player(index, TempPlayer(index).UseChar).Map, TempPlayer(index).CurConvoMapNpc).Num).Win = YES Then
                        '//Reseta o atributo caso tenha algum problema
                        If Npc(MapNpc(Player(index, TempPlayer(index).UseChar).Map, TempPlayer(index).CurConvoMapNpc).Num).Rebatle <> REBATLE_NEVER Then
                            Player(index, TempPlayer(index).UseChar).NpcBattledDay(MapNpc(Player(index, TempPlayer(index).UseChar).Map, TempPlayer(index).CurConvoMapNpc).Num).Win = NO
                            Player(index, TempPlayer(index).UseChar).NpcBattledDay(MapNpc(Player(index, TempPlayer(index).UseChar).Map, TempPlayer(index).CurConvoMapNpc).Num).NpcBattledAt = 0
                            Player(index, TempPlayer(index).UseChar).NpcBattledMonth(MapNpc(Player(index, TempPlayer(index).UseChar).Map, TempPlayer(index).CurConvoMapNpc).Num).NpcBattledAt = 0
                            Select Case TempPlayer(index).CurLanguage
                            Case LANG_PT: AddAlert index, "Tente novamente, por favor!", White
                            Case LANG_EN: AddAlert index, "Tente novamente, por favor!", White
                            Case LANG_ES: AddAlert index, "Tente novamente, por favor!", White
                            End Select
                        Else
                            Select Case TempPlayer(index).CurLanguage
                            Case LANG_PT: AddAlert index, "Você não pode lutar novamente com este TREINADOR!", White
                            Case LANG_EN: AddAlert index, "Você não pode lutar novamente com este TREINADOR!", White
                            Case LANG_ES: AddAlert index, "Você não pode lutar novamente com este TREINADOR!", White
                            End Select
                        End If
                        '//ToDo: Check if daily/monthly
                    ElseIf Not Player(index, TempPlayer(index).UseChar).NpcBattledDay(TempPlayer(index).CurConvoNpc).NpcBattledAt = Day(Now) Then
                        '// Start Npc Battle
                        If Player(index, TempPlayer(index).UseChar).Map > 0 Then
                            If MapNpc(Player(index, TempPlayer(index).UseChar).Map, TempPlayer(index).CurConvoMapNpc).InBattle <= 0 Then
                                MapNpc(Player(index, TempPlayer(index).UseChar).Map, TempPlayer(index).CurConvoMapNpc).InBattle = index
                                MapNpc(Player(index, TempPlayer(index).UseChar).Map, TempPlayer(index).CurConvoMapNpc).CurPokemon = 1
                                For i = 1 To MAX_PLAYER_POKEMON
                                    If Npc(MapNpc(Player(index, TempPlayer(index).UseChar).Map, TempPlayer(index).CurConvoMapNpc).Num).PokemonNum(i) > 0 Then
                                        MapNpc(Player(index, TempPlayer(index).UseChar).Map, TempPlayer(index).CurConvoMapNpc).PokemonAlive(i) = YES
                                    Else
                                        MapNpc(Player(index, TempPlayer(index).UseChar).Map, TempPlayer(index).CurConvoMapNpc).PokemonAlive(i) = NO
                                    End If
                                Next
                                SpawnNpcPokemon Player(index, TempPlayer(index).UseChar).Map, TempPlayer(index).CurConvoMapNpc, 1
                                TempPlayer(index).InNpcDuel = TempPlayer(index).CurConvoMapNpc
                                TempPlayer(index).DuelTime = 1
                                TempPlayer(index).DuelTimeTmr = GetTickCount + 1000
                                SendPlayerNpcDuel index
                            End If
                        End If
                    Else
                        '//Reseta o atributo caso tenha algum problema
                        If Npc(MapNpc(Player(index, TempPlayer(index).UseChar).Map, TempPlayer(index).CurConvoMapNpc).Num).Rebatle = REBATLE_NEVER Then
                            Player(index, TempPlayer(index).UseChar).NpcBattledDay(MapNpc(Player(index, TempPlayer(index).UseChar).Map, TempPlayer(index).CurConvoMapNpc).Num).Win = NO
                            Player(index, TempPlayer(index).UseChar).NpcBattledDay(MapNpc(Player(index, TempPlayer(index).UseChar).Map, TempPlayer(index).CurConvoMapNpc).Num).NpcBattledAt = 0
                            Player(index, TempPlayer(index).UseChar).NpcBattledMonth(MapNpc(Player(index, TempPlayer(index).UseChar).Map, TempPlayer(index).CurConvoMapNpc).Num).NpcBattledAt = 0
                            Select Case TempPlayer(index).CurLanguage
                            Case LANG_PT: AddAlert index, "Tente novamente, por favor!", White
                            Case LANG_EN: AddAlert index, "Tente novamente, por favor!", White
                            Case LANG_ES: AddAlert index, "Tente novamente, por favor!", White
                            End Select
                        Else
                            Select Case TempPlayer(index).CurLanguage    'AddAlert index, "You already battled this NPC", White
                            Case LANG_PT: AddAlert index, "Você já batalhou com esse npc hoje, tente novamente amanhã!", White
                            Case LANG_EN: AddAlert index, "You have already battled with this npc today, try again tomorrow!", White
                            Case LANG_ES: AddAlert index, "You have already battled with this npc today, try again tomorrow!", White
                            End Select
                        End If
                    End If
                End If
                fixData = False
            Case CONVO_SCRIPT_RELEARN
                If PlayerPokemon(index).Num > 0 And PlayerPokemon(index).slot > 0 Then
                    '//Send Relearn
                    SendRelearnMove index, PlayerPokemon(index).Num, PlayerPokemon(index).slot
                Else
                    AddAlert index, "Please spawn your pokemon", White
                End If
                fixData = False
            Case CONVO_SCRIPT_GIVEBADGE
                If .CustomScriptData > 0 And .CustomScriptData <= MAX_BADGE Then
                    Player(index, TempPlayer(index).UseChar).Badge(.CustomScriptData) = YES
                    SendPlayerData index
                End If
                fixData = False
            Case CONVO_SCRIPT_CHECKBADGE
                If .CustomScriptData > 0 And .CustomScriptData <= MAX_BADGE Then
                    If IsPlaying(index) Then
                        If TempPlayer(index).UseChar > 0 Then
                            If Player(index, TempPlayer(index).UseChar).Badge(.CustomScriptData) = YES Then
                                '//Next
                                TempPlayer(index).CurConvoData = .CustomScriptData2
                                fixData = True
                            Else
                                TempPlayer(index).CurConvoData = .CustomScriptData3
                                fixData = True
                            End If
                        End If
                    End If
                End If
            Case CONVO_SCRIPT_BEATPOKE
                If .CustomScriptData > 0 And .CustomScriptData <= MAX_GAME_POKEMON Then
                    If MapPokemon(.CustomScriptData).Num <= 0 Then
                        TempPlayer(index).CurConvoData = .CustomScriptData2
                        fixData = True
                    Else
                        TempPlayer(index).CurConvoData = .CustomScriptData3
                        fixData = True
                    End If
                End If
            Case CONVO_SCRIPT_CHECKITEM
                If .CustomScriptData > 0 And .CustomScriptData <= MAX_ITEM Then
                    If IsPlaying(index) Then
                        If TempPlayer(index).UseChar > 0 Then
                            i = checkItem(index, .CustomScriptData)
                            If i > 0 Then
                                '//Next
                                If PlayerInv(index).Data(i).Value >= .CustomScriptData2 Then
                                    TempPlayer(index).CurConvoData = .CustomScriptData3
                                    fixData = True
                                Else
                                    TempPlayer(index).CurConvoData = .MoveNext
                                    fixData = True
                                End If
                            Else
                                TempPlayer(index).CurConvoData = .MoveNext
                                fixData = True
                            End If
                        End If
                    End If
                End If
            Case CONVO_SCRIPT_TAKEITEM
                If .CustomScriptData > 0 And .CustomScriptData <= MAX_ITEM Then
                    If IsPlaying(index) Then
                        If TempPlayer(index).UseChar > 0 Then
                            i = checkItem(index, .CustomScriptData)
                            If i > 0 Then
                                '//Take Item
                                PlayerInv(index).Data(i).Value = PlayerInv(index).Data(i).Value - .CustomScriptData2
                                If PlayerInv(index).Data(i).Value <= 0 Then
                                    '//Clear Item
                                    PlayerInv(index).Data(i).Num = 0
                                    PlayerInv(index).Data(i).Value = 0
                                    PlayerInv(index).Data(i).TmrCooldown = 0
                                End If
                                SendPlayerInvSlot index, i
                            End If
                        End If
                    End If
                End If
                fixData = False
            Case CONVO_SCRIPT_RESPAWNPOKE
                If .CustomScriptData > 0 And .CustomScriptData <= MAX_GAME_POKEMON Then
                    SpawnMapPokemon .CustomScriptData, True
                End If
                fixData = False
            Case CONVO_SCRIPT_CHECKLEVEL
                If .CustomScriptData > 0 And .CustomScriptData <= MAX_LEVEL Then
                    If IsPlaying(index) Then
                        If TempPlayer(index).UseChar > 0 Then
                            If Player(index, TempPlayer(index).UseChar).Level >= (.CustomScriptData) Then
                                '//Next
                                TempPlayer(index).CurConvoData = .CustomScriptData2
                                fixData = True
                            Else
                                TempPlayer(index).CurConvoData = .CustomScriptData3
                                fixData = True
                            End If
                        End If
                    End If
                End If
            End Select

            '//Check if can init
            If .NoText = YES Then GoTo startOver
        End With
    Else
        '//End
        TempPlayer(index).CurConvoNum = 0
        TempPlayer(index).CurConvoData = 0
        TempPlayer(index).CurConvoNpc = 0
        TempPlayer(index).CurConvoMapNpc = 0
    End If

    SendInitConvo index, TempPlayer(index).CurConvoNum, TempPlayer(index).CurConvoData, NpcNum
End Sub

Public Function FindFreePokeStorageSlot(ByVal index As Long, ByVal StorageSlot As Byte) As Byte
    Dim i As Byte

    FindFreePokeStorageSlot = 0

    If Not IsPlaying(index) Then Exit Function
    If TempPlayer(index).UseChar <= 0 Then Exit Function
    If PlayerPokemonStorage(index).slot(StorageSlot).Unlocked = NO Then Exit Function

    For i = 1 To MAX_STORAGE
        With PlayerPokemonStorage(index).slot(StorageSlot).Data(i)
            If .Num = 0 Then
                FindFreePokeStorageSlot = i
                Exit Function
            End If
        End With
    Next
End Function

'//Catch
Public Function CatchMapPokemonData(ByVal index As Long, ByVal MapPokeNum As Long, ByVal UsedBall As Byte) As Boolean
    Dim StorageSlot As Byte
    Dim gotSlot As Byte
    Dim i As Long

    CatchMapPokemonData = False
    If MapPokeNum <= 0 Or MapPokeNum > MAX_GAME_POKEMON Then Exit Function
    If MapPokemon(MapPokeNum).Num <= 0 Then Exit Function

    AddLog Trim$(Player(index, TempPlayer(index).UseChar).Name) & " has caught " & Trim$(Pokemon(MapPokemon(MapPokeNum).Num).Name)

    gotSlot = FindOpenPokeSlot(index)
    '//Local Slot
    If gotSlot > 0 Then
        With PlayerPokemons(index).Data(gotSlot)
            .Num = MapPokemon(MapPokeNum).Num

            '//Stats
            .Level = MapPokemon(MapPokeNum).Level
            For i = 1 To StatEnum.Stat_Count - 1
                .Stat(i).Value = MapPokemon(MapPokeNum).Stat(i).Value
                .Stat(i).IV = MapPokemon(MapPokeNum).Stat(i).IV
                .Stat(i).EV = 0
            Next

            '//Vital
            .MaxHp = .Stat(StatEnum.HP).Value    'MapPokemon(MapPokeNum).MaxHP
            .CurHp = .MaxHp

            '//Nature
            .Nature = MapPokemon(MapPokeNum).Nature

            '//Shiny
            .IsShiny = MapPokemon(MapPokeNum).IsShiny

            '//Happiness
            .Happiness = MapPokemon(MapPokeNum).Happiness

            '//Gender
            .Gender = MapPokemon(MapPokeNum).Gender

            '//Status
            .Status = MapPokemon(MapPokeNum).Status

            '//Exp
            .CurExp = 0

            '//Moves
            For i = 1 To MAX_MOVESET
                .Moveset(i).Num = MapPokemon(MapPokeNum).Moveset(i).Num
                '//Reresh
                If .Moveset(i).Num > 0 Then
                    .Moveset(i).TotalPP = PokemonMove(.Moveset(i).Num).PP
                    .Moveset(i).CurPP = .Moveset(i).TotalPP
                    .Moveset(i).CD = 0
                End If
            Next

            '//Ball Used
            .BallUsed = UsedBall

            '//HeldItem
            .HeldItem = MapPokemon(MapPokeNum).HeldItem

            '//Add Pokedex
            AddPlayerPokedex index, .Num, YES, YES

            '//GlobalMsg IsShiny & Rarity
            If .IsShiny = YES Then
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If TempPlayer(index).UseChar > 0 Then
                            Select Case TempPlayer(index).CurLanguage
                            Case LANG_PT: SendPlayerMsg i, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " capturou um " & Trim$(Pokemon(.Num).Name) & " shiny em " & Trim$(Map(GetPlayerMap(index)).Name), Yellow
                            Case LANG_EN: SendPlayerMsg i, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " capturou um " & Trim$(Pokemon(.Num).Name) & " shiny em " & Trim$(Map(GetPlayerMap(index)).Name), Yellow
                            Case LANG_ES: SendPlayerMsg i, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " capturou um " & Trim$(Pokemon(.Num).Name) & " shiny em " & Trim$(Map(GetPlayerMap(index)).Name), Yellow
                            End Select
                        End If
                    End If
                Next i
            ElseIf Spawn(MapPokeNum).Rarity >= Options.Rarity Then
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If TempPlayer(index).UseChar > 0 Then
                            Select Case TempPlayer(index).CurLanguage
                            Case LANG_PT: SendPlayerMsg i, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " capturou um " & Trim$(Pokemon(.Num).Name) & " raro em " & Trim$(Map(GetPlayerMap(index)).Name), Yellow
                            Case LANG_EN: SendPlayerMsg i, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " capturou um " & Trim$(Pokemon(.Num).Name) & " raro em " & Trim$(Map(GetPlayerMap(index)).Name), Yellow
                            Case LANG_ES: SendPlayerMsg i, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " capturou um " & Trim$(Pokemon(.Num).Name) & " raro em " & Trim$(Map(GetPlayerMap(index)).Name), Yellow
                            End Select
                        End If
                    End If
                Next i
            End If
        End With

        UpdatePlayerPokemonOrder index
        SendPlayerPokemons index

        CatchMapPokemonData = True
        Exit Function
    Else
        '//Check Storage Slot
        For StorageSlot = 1 To MAX_STORAGE_SLOT
            gotSlot = FindFreePokeStorageSlot(index, StorageSlot)
            If gotSlot > 0 Then
                '//Give Pokemon
                With PlayerPokemonStorage(index).slot(StorageSlot).Data(gotSlot)
                    .Num = MapPokemon(MapPokeNum).Num

                    '//Stats
                    .Level = MapPokemon(MapPokeNum).Level
                    For i = 1 To StatEnum.Stat_Count - 1
                        .Stat(i).Value = MapPokemon(MapPokeNum).Stat(i).Value
                        .Stat(i).IV = MapPokemon(MapPokeNum).Stat(i).IV
                        .Stat(i).EV = 0
                    Next

                    '//Vital
                    .MaxHp = MapPokemon(MapPokeNum).MaxHp
                    .CurHp = MapPokemon(MapPokeNum).CurHp

                    '//Nature
                    .Nature = MapPokemon(MapPokeNum).Nature

                    '//Shiny
                    .IsShiny = MapPokemon(MapPokeNum).IsShiny

                    '//Happiness
                    .Happiness = MapPokemon(MapPokeNum).Happiness

                    '//Gender
                    .Gender = MapPokemon(MapPokeNum).Gender

                    '//Status
                    .Status = MapPokemon(MapPokeNum).Status

                    '//Exp
                    .CurExp = 0

                    '//Moves
                    For i = 1 To MAX_MOVESET
                        .Moveset(i).Num = MapPokemon(MapPokeNum).Moveset(i).Num
                        '//Reresh
                        If .Moveset(i).Num > 0 Then
                            .Moveset(i).TotalPP = PokemonMove(.Moveset(i).Num).PP
                            .Moveset(i).CurPP = .Moveset(i).TotalPP
                            .Moveset(i).CD = 0
                        End If
                    Next

                    '//Ball Used
                    .BallUsed = UsedBall

                    .HeldItem = 0

                    '//Add Pokedex
                    AddPlayerPokedex index, .Num, YES, YES

                    '//GlobalMsg IsShiny & Rarity
                    If .IsShiny = YES Then
                        For i = 1 To Player_HighIndex
                            If IsPlaying(i) Then
                                If TempPlayer(index).UseChar > 0 Then
                                    Select Case TempPlayer(index).CurLanguage
                                    Case LANG_PT: SendPlayerMsg i, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " capturou um " & Trim$(Pokemon(.Num).Name) & " shiny em " & Trim$(Map(GetPlayerMap(index)).Name), Yellow
                                    Case LANG_EN: SendPlayerMsg i, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " capturou um " & Trim$(Pokemon(.Num).Name) & " shiny em " & Trim$(Map(GetPlayerMap(index)).Name), Yellow
                                    Case LANG_ES: SendPlayerMsg i, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " capturou um " & Trim$(Pokemon(.Num).Name) & " shiny em " & Trim$(Map(GetPlayerMap(index)).Name), Yellow
                                    End Select
                                End If
                            End If
                        Next i
                    ElseIf Spawn(MapPokeNum).Rarity >= Options.Rarity Then
                        For i = 1 To Player_HighIndex
                            If IsPlaying(i) Then
                                If TempPlayer(index).UseChar > 0 Then
                                    Select Case TempPlayer(index).CurLanguage
                                    Case LANG_PT: SendPlayerMsg i, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " capturou um " & Trim$(Pokemon(.Num).Name) & " raro em " & Trim$(Map(GetPlayerMap(index)).Name), Yellow
                                    Case LANG_EN: SendPlayerMsg i, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " capturou um " & Trim$(Pokemon(.Num).Name) & " raro em " & Trim$(Map(GetPlayerMap(index)).Name), Yellow
                                    Case LANG_ES: SendPlayerMsg i, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " capturou um " & Trim$(Pokemon(.Num).Name) & " raro em " & Trim$(Map(GetPlayerMap(index)).Name), Yellow
                                    End Select
                                End If
                            End If
                        Next i
                    End If
                End With

                Select Case TempPlayer(index).CurLanguage
                Case LANG_PT: AddAlert index, "Your pokemon has been transferred to your pokemon storage", White
                Case LANG_EN: AddAlert index, "Your pokemon has been transferred to your pokemon storage", White
                Case LANG_ES: AddAlert index, "Your pokemon has been transferred to your pokemon storage", White
                End Select

                SendPlayerPokemonStorageSlot index, StorageSlot, gotSlot

                CatchMapPokemonData = True
                Exit Function
            End If
        Next
    End If

    CatchMapPokemonData = False
End Function

Public Function FindOpenTradeSlot(ByVal index As Long) As Long
    Dim i As Byte

    For i = 1 To MAX_TRADE
        If TempPlayer(index).TradeItem(i).Type = 0 Then
            FindOpenTradeSlot = i
            Exit Function
        End If
    Next
End Function

Public Sub AddPlayerPokedex(ByVal index As Long, ByVal PokeNum As Long, Optional ByVal Obtained As Byte = 0, Optional ByVal Scanned As Byte = 0)
    If PlayerPokedex(index).PokemonIndex(PokeNum).Obtained = 0 Then
        PlayerPokedex(index).PokemonIndex(PokeNum).Obtained = Obtained
        If Obtained = YES Then
            Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: AddAlert index, Trim$(Pokemon(PokeNum).Name) & " has been added on pokedex", White
            Case LANG_EN: AddAlert index, Trim$(Pokemon(PokeNum).Name) & " has been added on pokedex", White
            Case LANG_ES: AddAlert index, Trim$(Pokemon(PokeNum).Name) & " has been added on pokedex", White
            End Select
        End If
    End If
    If PlayerPokedex(index).PokemonIndex(PokeNum).Scanned = 0 Then
        PlayerPokedex(index).PokemonIndex(PokeNum).Scanned = Scanned
        If Scanned = YES Then
            Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: AddAlert index, Trim$(Pokemon(PokeNum).Name) & " has been scanned", White
            Case LANG_EN: AddAlert index, Trim$(Pokemon(PokeNum).Name) & " has been scanned", White
            Case LANG_ES: AddAlert index, Trim$(Pokemon(PokeNum).Name) & " has been scanned", White
            End Select
        End If
    End If
    SendPlayerPokedexSlot index, PokeNum
End Sub

Public Sub ClearMyTarget(ByVal index As Long, ByVal MapNum As Long)
    Dim i As Long

    For i = 1 To Pokemon_HighIndex
        If MapPokemon(i).Num > 0 Then
            If MapPokemon(i).Map = MapNum Then
                If MapPokemon(i).targetType = TARGET_TYPE_PLAYER Then
                    If MapPokemon(i).TargetIndex = index Then
                        MapPokemon(i).targetType = 0
                        MapPokemon(i).TargetIndex = 0
                    End If
                End If
            End If
        End If
    Next
End Sub

Public Sub ChangeTempSprite(ByVal index As Long, ByVal TempSprite As Byte, Optional ByVal ItemNum As Long = 0, Optional ByVal Forced As Boolean = False)
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub

    Select Case TempSprite
    Case TEMP_SPRITE_GROUP_NONE
        If Forced Then
            Call ClearTempSprite(index)
        Else
            If Not TempPlayer(index).TempSprite.TempSpriteType = TEMP_SPRITE_GROUP_BIKE Then
                If Not TempPlayer(index).TempSprite.TempSpriteType = TEMP_SPRITE_GROUP_MOUNT Then
                    Call ClearTempSprite(index)
                End If
            End If
        End If
    Case TEMP_SPRITE_GROUP_DIVE
        Player(index, TempPlayer(index).UseChar).KeyItemNum = ItemNum
        TempPlayer(index).TempSprite.TempSpriteType = TEMP_SPRITE_GROUP_DIVE
    Case TEMP_SPRITE_GROUP_BIKE
        If Not TempPlayer(index).TempSprite.TempSpriteType = TEMP_SPRITE_GROUP_DIVE Then
            If Not TempPlayer(index).TempSprite.TempSpriteType = TEMP_SPRITE_GROUP_BIKE Then
                Player(index, TempPlayer(index).UseChar).KeyItemNum = ItemNum
                TempPlayer(index).TempSprite.TempSpriteType = TEMP_SPRITE_GROUP_BIKE
            Else
                Call ClearTempSprite(index)
            End If
        End If
    Case TEMP_SPRITE_GROUP_MOUNT
        If Not TempPlayer(index).TempSprite.TempSpriteType = TEMP_SPRITE_GROUP_DIVE Then
            If Not TempPlayer(index).TempSprite.TempSpriteType = TEMP_SPRITE_GROUP_BIKE Then
                If Not TempPlayer(index).TempSprite.TempSpriteType = TEMP_SPRITE_GROUP_MOUNT Then
                    Player(index, TempPlayer(index).UseChar).KeyItemNum = ItemNum
                    TempPlayer(index).TempSprite.TempSpriteType = TEMP_SPRITE_GROUP_MOUNT
                    TempPlayer(index).TempSprite.TempSpriteID = Item(ItemNum).Data3
                    TempPlayer(index).TempSprite.TempSpriteExp = Item(ItemNum).Data4
                    TempPlayer(index).TempSprite.TempSpritePassiva = Item(ItemNum).Data5
                Else
                    Call ClearTempSprite(index)
                End If
            End If
        End If
        'Case TEMP_SPRITE_GROUP_SURF

    Case Else
        Call ClearTempSprite(index)
    End Select

    SendPlayerData index
End Sub

Public Function FindInvItemSlot(ByVal index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    For i = 1 To MAX_PLAYER_INV
        If PlayerInv(index).Data(i).Num = ItemNum Then
            FindInvItemSlot = i
            Exit Function
        End If
    Next
End Function

Public Sub SendPlayerPokemonFaint(ByVal index As Long)
    Dim MapNum As Long
    Dim DuelIndex As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If PlayerPokemon(index).Num <= 0 Then Exit Sub

    MapNum = Player(index, TempPlayer(index).UseChar).Map

    ClearPlayerPokemon index
    If TempPlayer(index).InDuel > 0 Then
        If IsPlaying(TempPlayer(index).InDuel) Then
            If TempPlayer(TempPlayer(index).InDuel).UseChar > 0 Then
                If CountPlayerPokemonAlive(index) <= 0 Then
                    DuelIndex = TempPlayer(index).InDuel
                    '//Player Lose
                    SendActionMsg MapNum, "Lose!", Player(index, TempPlayer(index).UseChar).x * 32, Player(index, TempPlayer(index).UseChar).Y * 32, White
                    SendActionMsg MapNum, "Win!", Player(DuelIndex, TempPlayer(DuelIndex).UseChar).x * 32, Player(DuelIndex, TempPlayer(DuelIndex).UseChar).Y * 32, White
                    Player(index, TempPlayer(index).UseChar).Lose = Player(index, TempPlayer(index).UseChar).Lose + 1
                    Player(DuelIndex, TempPlayer(DuelIndex).UseChar).Win = Player(DuelIndex, TempPlayer(DuelIndex).UseChar).Win + 1
                    SendPlayerPvP (DuelIndex)
                    SendPlayerPvP (index)
                    TempPlayer(index).InDuel = 0
                    TempPlayer(index).DuelTime = 0
                    TempPlayer(index).DuelTimeTmr = 0
                    TempPlayer(index).WarningTimer = 0
                    TempPlayer(index).PlayerRequest = 0
                    TempPlayer(index).RequestType = 0
                    TempPlayer(DuelIndex).InDuel = 0
                    TempPlayer(DuelIndex).DuelTime = 0
                    TempPlayer(DuelIndex).DuelTimeTmr = 0
                    TempPlayer(DuelIndex).WarningTimer = 0
                    TempPlayer(DuelIndex).PlayerRequest = 0
                    TempPlayer(DuelIndex).RequestType = 0
                    SendRequest DuelIndex
                    SendRequest index
                Else
                    TempPlayer(index).DuelReset = YES
                End If
            End If
        End If
    End If
    If TempPlayer(index).InNpcDuel > 0 Then
        If CountPlayerPokemonAlive(index) <= 0 Then
            '//Adicionado a apenas um método.
            PlayerLoseToNpc index, TempPlayer(index).InNpcDuel
        Else
            TempPlayer(index).DuelReset = YES
        End If
    End If
End Sub

Public Function GetLevelNextExp(ByVal Level As Long) As Long
    GetLevelNextExp = ((Level + 5) ^ 3) * (((((Level + 5) + 1) / 3) + 24) / 50)
    'GetLevelNextExp = (Level ^ 3) * (((Level / 2) + 32) / 50)
    'GetLevelNextExp = ((250 * Level) / 100) + ((10 + Level) / 2)
End Function

Public Function GetPlayerHP(ByVal Level As Long) As Long
    GetPlayerHP = ((250 * Level) / 100) + ((10 + Level) / 2)
End Function

Public Sub GivePlayerExp(ByVal index As Long, ByVal Exp As Long)
    Dim ExpRollover As Long
    Dim TotalBonus As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub

    ' Exp Rate
    Exp = Exp * Options.ExpRate

    With Player(index, TempPlayer(index).UseChar)
        If .Level >= MAX_PLAYER_LEVEL Then Exit Sub

        'Exp Rate
        If EventExp.ExpEvent Then
            Exp = Exp * EventExp.ExpMultiply
            TotalBonus = TotalBonus + (EventExp.ExpMultiply * 100)
        End If
        'Exp Mount
        If TempPlayer(index).TempSprite.TempSpriteExp > 0 Then
            Exp = Exp + ((Exp / 100) * TempPlayer(index).TempSprite.TempSpriteExp)
            TotalBonus = TotalBonus + (TempPlayer(index).TempSprite.TempSpriteExp)
        End If

        .CurExp = .CurExp + Exp

        If .CurExp > GetLevelNextExp(.Level) Then
            Do While .CurExp > GetLevelNextExp(.Level)
                ExpRollover = .CurExp - GetLevelNextExp(.Level)
                .CurExp = ExpRollover
                .Level = .Level + 1
                .CurHp = GetPlayerHP(.Level)
            Loop
            SendPlayerData index

            '//ActionMsg
            SendActionMsg Player(index, TempPlayer(index).UseChar).Map, "Level Up!", .x * 32, .Y * 32, Yellow
        End If
        SendPlayerExp index

        '//ActionMsg
        SendActionMsg Player(index, TempPlayer(index).UseChar).Map, "+" & Exp, .x * 32, .Y * 32, White

        If TotalBonus > 0 Then
            SendActionMsg Player(index, TempPlayer(index).UseChar).Map, "+" & TotalBonus & "% EXP!", .x * 32, (.Y - 1) * 32, Black
        End If
    End With
End Sub

Public Function GetExpPenalty(ByVal Level As Long)
    Dim nextExp As Long, Penalty As Long
    Dim LevelRate As Single

    nextExp = ((Level + 5) ^ 3) * (((((Level + 5) + 1) / 3) + 24) / 50)
    If Level >= 50 And Level <= 69 Then
        LevelRate = 0.1
    ElseIf Level >= 70 And Level <= 89 Then
        LevelRate = 0.2
    ElseIf Level >= 90 Then
        LevelRate = 0.4
    Else
        LevelRate = 0
    End If
    Penalty = nextExp * (0.8 + LevelRate)
    If Penalty <= 0 Then Penalty = 0
    GetExpPenalty = Penalty
End Function

Public Function GetMoneyPenalty(ByVal Level As Long, ByVal BadgeCount As Byte)
    Dim Penalty As Long

    Penalty = Level * ((BadgeCount + 1) * 120)
    If Penalty <= 0 Then Penalty = 0
    GetMoneyPenalty = Penalty
End Function

Public Sub KillPlayer(ByVal index As Long)
    Dim ExpPenalty As Long, MoneyPenalty As Long
    Dim ExpRollover As Long, CountBadge As Byte
    Dim i As Byte

    With Player(index, TempPlayer(index).UseChar)
        .CurHp = GetPlayerHP(.Level)
        If .CheckMap <= 0 Then .CheckMap = Options.StartMap
        If .CheckX <= 0 Then .CheckX = Options.startX
        If .CheckY <= 0 Then .CheckY = Options.startY
        PlayerWarp index, .CheckMap, .CheckX, .CheckY, .CheckDir

        '//Penalty
        'ExpPenalty = GetExpPenalty(.Level)
        CountBadge = 0
        For i = 1 To MAX_BADGE
            If .Badge(i) = YES Then
                CountBadge = CountBadge + 1
            End If
        Next
        MoneyPenalty = GetMoneyPenalty(.Level, CountBadge)

        'If ExpPenalty > .CurExp Then
        '    If .Level <= 1 Then
        '        .CurExp = 0
        '    Else
        '        'Do While ExpPenalty > .CurExp
        '    ExpRollover = ExpPenalty - .CurExp
        '    .Level = .Level - 1
        '    ExpPenalty = ExpRollover
        '    .CurExp = (GetLevelNextExp(.Level) - 1)
        'Loop
        '        Do While (ExpPenalty > 0)
        '            If .Level > 1 Then
        '                If ExpPenalty > .CurExp Then
        '                    ExpRollover = ExpPenalty - .CurExp
        '                    ExpPenalty = ExpRollover
        '                    .Level = .Level - 1
        '                   .CurExp = GetLevelNextExp(.Level) - 1
        '               Else
        '                   .CurExp = .CurExp - ExpPenalty
        '                   ExpPenalty = 0
        '                   ExpRollover = 0
        '               End If
        '           Else
        '                ExpPenalty = 0
        '            End If
        '        Loop
        '    End If
        '    .CurHP = GetPlayerHP(.Level)
        'Else
        '    .CurExp = .CurExp - ExpPenalty
        'End If
        If .Money <= MoneyPenalty Then
            .Money = 0
        Else
            .Money = .Money - MoneyPenalty
        End If

        'AddAlert Index, "You lose " & ExpPenalty & " trainer exp", White
        AddAlert index, "You lose $" & MoneyPenalty, White

        SendPlayerData index
    End With
End Sub

Public Sub SendWhosOnline(ByVal index As Long)
    Dim s As String
    Dim i As Long

    s = "Player Online: "
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If TempPlayer(i).UseChar > 0 Then
                s = s & Trim$(Player(i, TempPlayer(i).UseChar).Name) & ", "
            End If
        End If
    Next
    s = Left(s, Len(s) - 2)
    SendPlayerMsg index, s, White
End Sub

Public Sub CreateParty(ByVal index As Long)
    Dim i As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If TempPlayer(index).InParty > 0 Then Exit Sub

    TempPlayer(index).InParty = YES
    For i = 1 To MAX_PARTY
        TempPlayer(index).PartyIndex(i) = 0
    Next
    TempPlayer(index).PartyIndex(1) = index
    AddAlert index, "Party Created", White
    SendParty index
End Sub

Public Sub LeaveParty(ByVal index As Long)
    Dim i As Long, PartyRequest As Long, PartySlot As Byte

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If TempPlayer(index).InParty <= 0 Then Exit Sub

    TempPlayer(index).InParty = 0
    For i = 1 To MAX_PARTY
        PartyRequest = TempPlayer(index).PartyIndex(i)
        '//Remove self
        If PartyRequest = index Then
            PartySlot = i
            TempPlayer(index).PartyIndex(i) = 0
        End If
    Next
    '//Update to member
    For i = 1 To MAX_PARTY
        PartyRequest = TempPlayer(index).PartyIndex(i)
        If PartyRequest > 0 Then
            If IsPlaying(PartyRequest) Then
                If TempPlayer(PartyRequest).UseChar > 0 Then
                    If Not PartyRequest = index Then
                        TempPlayer(PartyRequest).PartyIndex(PartySlot) = 0
                        AddAlert PartyRequest, Trim$(Player(index, TempPlayer(index).UseChar).Name) & " has left the party", White
                        SendParty PartyRequest
                    End If
                End If
            End If
        End If
    Next
    AddAlert index, "You left the party", White
    SendParty index
End Sub

Public Sub JoinParty(ByVal index As Long, ByVal InviteIndex As Long)
    Dim i As Long, slot As Byte
    Dim PartyRequest As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    If TempPlayer(index).InParty <= 0 Then Exit Sub
    If Not IsPlaying(InviteIndex) Then Exit Sub
    If TempPlayer(InviteIndex).UseChar <= 0 Then Exit Sub
    If TempPlayer(InviteIndex).InParty > 0 Then Exit Sub
    slot = 0
    '//Check free slot
    For i = 1 To MAX_PARTY
        If TempPlayer(index).PartyIndex(i) <= 0 Then
            slot = i
            Exit For
        End If
    Next

    If slot > 0 Then
        For i = 1 To MAX_PARTY
            PartyRequest = TempPlayer(index).PartyIndex(i)
            If PartyRequest > 0 Then
                If IsPlaying(PartyRequest) Then
                    If TempPlayer(PartyRequest).UseChar > 0 Then
                        TempPlayer(PartyRequest).PartyIndex(slot) = InviteIndex
                        SendParty PartyRequest
                    End If
                End If
            End If
        Next

        For i = 1 To MAX_PARTY
            TempPlayer(InviteIndex).PartyIndex(i) = TempPlayer(index).PartyIndex(i)
        Next
        TempPlayer(InviteIndex).InParty = YES
        SendParty InviteIndex
    End If
End Sub

Public Function PartyCount(ByVal index As Long) As Byte
    Dim i As Long, count As Long

    count = 0
    For i = 1 To MAX_PARTY
        If TempPlayer(index).PartyIndex(i) > 0 Then
            count = count + 1
        End If
    Next
    PartyCount = count
End Function

Public Function IsPartyMember(ByVal index As Long, ByVal i As Long) As Boolean
    Dim z As Byte
    If TempPlayer(index).InParty > 0 Then
        For z = 1 To MAX_PARTY
            If TempPlayer(index).PartyIndex(z) > 0 Then
                If TempPlayer(index).PartyIndex(z) <> index Then
                    If TempPlayer(index).PartyIndex(z) = i Then
                        IsPartyMember = True
                        Exit Function
                    End If
                End If
            End If
        Next z
    End If
End Function

Public Function IsIPBanned(ByVal valIP As String) As Boolean
    Dim filename As String
    Dim f As Long
    Dim s As String

    filename = App.Path & "\data\accounts\banlist.txt"
    f = FreeFile

    '//Check if the master banlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist(filename) Then
        Open App.Path & "\data\accounts\banlist.txt" For Output As #f
        Close #f
        IsIPBanned = False
        Exit Function
    End If

    Open filename For Input As #f
    Do While Not EOF(f)
        Input #f, s

        If Trim$(LCase$(s)) = Trim$(LCase$(valIP)) Then
            IsIPBanned = True
            Close #f
            Exit Function
        End If
    Loop
    Close #f
End Function

Public Sub BanIP(ByVal valIP As String)
    Dim f As Long

    '//Append name to file
    f = FreeFile
    Open App.Path & "\data\accounts\banlist.txt" For Append As #f
    Print #f, Trim$(valIP)
    Close #f
End Sub

Public Function IsCharacterBanned(ByVal char As String) As Boolean
    Dim filename As String
    Dim f As Long
    Dim s As String

    filename = App.Path & "\data\accounts\charbanlist.txt"
    f = FreeFile

    '//Check if the master banlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist(filename) Then
        Open App.Path & "\data\accounts\charbanlist.txt" For Output As #f
        Close #f
        IsCharacterBanned = False
        Exit Function
    End If

    Open filename For Input As #f
    Do While Not EOF(f)
        Input #f, s

        If Trim$(LCase$(s)) = Trim$(LCase$(char)) Then
            IsCharacterBanned = True
            Close #f
            Exit Function
        End If
    Loop
    Close #f
End Function

Public Sub BanCharacter(ByVal char As String)
    Dim f As Long

    '//Append name to file
    f = FreeFile
    Open App.Path & "\data\accounts\charbanlist.txt" For Append As #f
    Print #f, Trim$(char)
    Close #f
End Sub

' Obtem o ID do mapa
Function GetPlayerMap(ByVal index As Long) As Long

    If index > MAX_PLAYER Then Exit Function
    GetPlayerMap = Player(index, TempPlayer(index).UseChar).Map

End Function

' Obtem o X do jogador
Function GetPlayerX(ByVal index As Long) As Long

    If index > MAX_PLAYER Then Exit Function
    GetPlayerX = Player(index, TempPlayer(index).UseChar).x
End Function

' Obtem o Y do jogador
Function GetPlayerY(ByVal index As Long) As Long

    If index > MAX_PLAYER Then Exit Function
    GetPlayerY = Player(index, TempPlayer(index).UseChar).Y
End Function

Function GetPlayerDir(ByVal index As Long) As Long
    If index > MAX_PLAYER Then Exit Function
    GetPlayerDir = Player(index, TempPlayer(index).UseChar).Dir
End Function

Function GetPlayerLogin(ByVal index As Long) As String
    GetPlayerLogin = Trim$(Player(index, TempPlayer(index).UseChar).Name)
End Function

Function HasInvItem(ByVal index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If ItemNum <= 0 Or ItemNum > MAX_ITEM Then
        HasInvItem = 0
        Exit Function
    End If

    For i = 1 To MAX_PLAYER_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then

            If Item(ItemNum).Stock > 0 Then
                HasInvItem = GetPlayerInvItemValue(index, i)
            Else
                HasInvItem = 1
            End If

            Exit Function

        End If
    Next

End Function

Function GetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long) As Long
    If index > Player_HighIndex Then Exit Function
    If InvSlot = 0 Then Exit Function

    GetPlayerInvItemNum = PlayerInv(index).Data(InvSlot).Num
End Function

Function GetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long) As Long
    If index > Player_HighIndex Then Exit Function
    If InvSlot = 0 Then Exit Function

    GetPlayerInvItemValue = PlayerInv(index).Data(InvSlot).Value
End Function

Function HasStorageItem(ByVal index As Long, ByVal StorageSlot As Byte, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If ItemNum <= 0 Or ItemNum > MAX_ITEM Then
        HasStorageItem = 0
        Exit Function
    End If

    For i = 1 To MAX_STORAGE
        ' Check to see if the player has the item
        If GetPlayerStorageItemNum(index, StorageSlot, i) = ItemNum Then

            If Item(ItemNum).Stock > 0 Then
                HasStorageItem = GetPlayerStorageItemValue(index, StorageSlot, i)
            Else
                HasStorageItem = 1
            End If

            Exit Function

        End If
    Next

End Function

Function GetPlayerStorageItemNum(ByVal index As Long, ByVal StorageSlot As Byte, ByVal StorageData As Long) As Long
    If index > Player_HighIndex Then Exit Function
    If StorageSlot = 0 Then Exit Function
    If PlayerInvStorage(index).slot(StorageSlot).Unlocked = NO Then Exit Function

    GetPlayerStorageItemNum = PlayerInvStorage(index).slot(StorageSlot).Data(StorageData).Num
End Function

Function GetPlayerStorageItemValue(ByVal index As Long, ByVal StorageSlot As Long, ByVal StorageData As Long) As Long
    If index > Player_HighIndex Then Exit Function
    If StorageSlot = 0 Then Exit Function
    If PlayerInvStorage(index).slot(StorageSlot).Unlocked = NO Then Exit Function

    GetPlayerStorageItemValue = PlayerInvStorage(index).slot(StorageSlot).Data(StorageData).Value
End Function

Sub SetPlayerFishMode(ByVal index As Long, ByVal Mode As Byte)
    If index > Player_HighIndex Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub

    Player(index, TempPlayer(index).UseChar).FishMode = Mode
End Sub

Sub SetPlayerFishRod(ByVal index As Long, ByVal Rod As Byte)
    If index > Player_HighIndex Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub

    Player(index, TempPlayer(index).UseChar).FishRod = Rod
End Sub

Function GetPlayerFishMode(ByVal index As Long) As Byte
    If index > Player_HighIndex Then Exit Function
    If TempPlayer(index).UseChar <= 0 Then Exit Function

    GetPlayerFishMode = Player(index, TempPlayer(index).UseChar).FishMode
End Function

Function GetPlayerFishRod(ByVal index As Long) As Byte
    If index > Player_HighIndex Then Exit Function
    If TempPlayer(index).UseChar <= 0 Then Exit Function

    GetPlayerFishRod = Player(index, TempPlayer(index).UseChar).FishRod
End Function

Public Function GetPlayerName(ByVal index As Long) As String
    If TempPlayer(index).UseChar <= 0 Then Exit Function
    If Not IsPlaying(index) Then Exit Function
    
    GetPlayerName = Trim$(Player(index, TempPlayer(index).UseChar).Name)
End Function
