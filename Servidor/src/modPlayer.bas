Attribute VB_Name = "modPlayer"
Option Explicit

' **********************
' ** Player Functions **
' **********************
Public Function GetPlayerIP(ByVal Index As Long) As String
    If Index <= 0 Or Index > MAX_PLAYER Then Exit Function
    GetPlayerIP = frmServer.Socket(Index).RemoteHostIP
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

Public Sub PlayerWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, ByVal Dir As Byte)
Dim OldMap As Long

    '//Exit out when error
    If Index <= 0 Or Index > MAX_PLAYER Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Or TempPlayer(Index).UseChar > MAX_PLAYERCHAR Then Exit Sub
    If MapNum <= 0 Or MapNum > MAX_MAP Then Exit Sub
    
    '//Correct error position
    If X <= 0 Then X = 0
    If X > Map(MapNum).MaxX Then X = Map(MapNum).MaxX
    If Y <= 0 Then Y = 0
    If Y > Map(MapNum).MaxY Then Y = Map(MapNum).MaxY
    
    OldMap = Player(Index, TempPlayer(Index).UseChar).Map
    
    '//Update position
    With Player(Index, TempPlayer(Index).UseChar)
        .Map = MapNum
        .X = X
        .Y = Y
        .Dir = Dir
    End With
    
    '//If map did not match
    If Not OldMap = MapNum Then
        '//Clear player data on old map
        SendLeaveMap Index, OldMap
        
        '//Clear Target
        ClearMyTarget Index, OldMap
        
        '//Check if there's still remaining player on map
        If TotalPlayerOnMap(OldMap) <= 0 Then
            PlayerOnMap(OldMap) = NO
            Map(OldMap).CurWeather = Map(OldMap).StartWeather
        End If
        
        TempPlayer(Index).MapSwitchTmr = YES
    End If
    
    '//Add log
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " has been warped on Map#" & MapNum & " x:" & X & " y:" & Y
    
    '//Update
    PlayerOnMap(MapNum) = YES
    TempPlayer(Index).GettingMap = True
    SendCheckForMap Index, MapNum
End Sub

Public Sub ForcePlayerMove(ByVal Index As Long, ByVal Dir As Byte)
    '//Exit out when error
    If Index <= 0 Or Index > MAX_PLAYER Then Exit Sub
    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Or TempPlayer(Index).UseChar > MAX_PLAYERCHAR Then Exit Sub
    If Dir < 0 Or Dir > DIR_RIGHT Then Exit Sub

    Select Case Dir
        Case DIR_UP
            If Player(Index, TempPlayer(Index).UseChar).Y = 0 Then Exit Sub
        Case DIR_LEFT
            If Player(Index, TempPlayer(Index).UseChar).X = 0 Then Exit Sub
        Case DIR_DOWN
            If Player(Index, TempPlayer(Index).UseChar).Y = Map(Player(Index, TempPlayer(Index).UseChar).Map).MaxY Then Exit Sub
        Case DIR_RIGHT
            If Player(Index, TempPlayer(Index).UseChar).X = Map(Player(Index, TempPlayer(Index).UseChar).Map).MaxX Then Exit Sub
    End Select
    
    PlayerMove Index, Dir, True
End Sub

Public Sub PlayerMove(ByVal Index As Long, ByVal Dir As Byte, Optional ByVal sendToSelf As Boolean = False)
Dim DidMove As Boolean
Dim OldX As Long, OldY As Long
Dim gothealed As Boolean
Dim i As Long, X As Byte

    '//Exit out when error
    If Index <= 0 Or Index > MAX_PLAYER Then Exit Sub
    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Or TempPlayer(Index).UseChar > MAX_PLAYERCHAR Then Exit Sub
    If Dir < 0 Or Dir > DIR_RIGHT Then Exit Sub

    DidMove = False
    
    With Player(Index, TempPlayer(Index).UseChar)
        '//Store original location in case it got desync
        OldX = .X
        OldY = .Y
        
        Select Case Dir
            Case DIR_UP
                .Dir = DIR_UP
                
                '//Check to make sure not outside of boundries
                If .Y > 0 Then
                    If Not CheckDirection(.Map, DIR_UP, .X, .Y) Then
                        .Y = .Y - 1
                        DidMove = True
                    End If
                Else
                    '//Check Link
                    If Map(.Map).LinkUp > 0 Then
                        PlayerWarp Index, Map(.Map).LinkUp, .X, Map(Map(.Map).LinkUp).MaxY, .Dir
                        Exit Sub
                    End If
                End If
            Case DIR_DOWN
                .Dir = DIR_DOWN
                
                '//Check to make sure not outside of boundries
                If .Y < Map(.Map).MaxY Then
                    If Not CheckDirection(.Map, DIR_DOWN, .X, .Y) Then
                        .Y = .Y + 1
                        DidMove = True
                    End If
                Else
                    '//Check Link
                    If Map(.Map).LinkDown > 0 Then
                        PlayerWarp Index, Map(.Map).LinkDown, .X, 0, .Dir
                        Exit Sub
                    End If
                End If
            Case DIR_LEFT
                .Dir = DIR_LEFT
                
                '//Check to make sure not outside of boundries
                If .X > 0 Then
                    If Not CheckDirection(.Map, DIR_LEFT, .X, .Y) Then
                        .X = .X - 1
                        DidMove = True
                    End If
                Else
                    '//Check Link
                    If Map(.Map).LinkLeft > 0 Then
                        PlayerWarp Index, Map(.Map).LinkLeft, Map(Map(.Map).LinkLeft).MaxX, .Y, .Dir
                        Exit Sub
                    End If
                End If
            Case DIR_RIGHT
                .Dir = DIR_RIGHT
                
                '//Check to make sure not outside of boundries
                If .X < Map(.Map).MaxX Then
                    If Not CheckDirection(.Map, DIR_RIGHT, .X, .Y) Then
                        .X = .X + 1
                        DidMove = True
                    End If
                Else
                    '//Check Link
                    If Map(.Map).LinkRight > 0 Then
                        PlayerWarp Index, Map(.Map).LinkRight, 0, .Y, .Dir
                        Exit Sub
                    End If
                End If
        End Select
    
        '//Got Desynced
        If Not DidMove Then
            .X = OldX
            .Y = OldY
            SendPlayerXY Index
            SendPlayerXY Index, True
            
            'If .Action <> 0 Then
            '    .Action = 0
            '    SendPlayerAction Index
            'End If
        Else
            TempPlayer(Index).MapSwitchTmr = NO
            
            SendPlayerMove Index, sendToSelf
            
            '//Check tile attribute
            Select Case Map(.Map).Tile(.X, .Y).Attribute
                Case MapAttribute.Warp
                    '//Warp
                    If Map(.Map).Tile(.X, .Y).Data1 > 0 Then
                        PlayerWarp Index, Map(.Map).Tile(.X, .Y).Data1, Map(.Map).Tile(.X, .Y).Data2, Map(.Map).Tile(.X, .Y).Data3, Map(.Map).Tile(.X, .Y).Data4
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
                        If PlayerPokemons(Index).Data(i).Num > 0 Then
                            If PlayerPokemons(Index).Data(i).CurHp < PlayerPokemons(Index).Data(i).MaxHp Then
                                PlayerPokemons(Index).Data(i).CurHp = PlayerPokemons(Index).Data(i).MaxHp
                                gothealed = True
                            End If
                            If PlayerPokemons(Index).Data(i).Status > 0 Then
                                PlayerPokemons(Index).Data(i).Status = 0
                                gothealed = True
                            End If
                            For X = 1 To MAX_MOVESET
                                If PlayerPokemons(Index).Data(i).Moveset(X).Num > 0 Then
                                    If PlayerPokemons(Index).Data(i).Moveset(X).CurPP < PlayerPokemons(Index).Data(i).Moveset(X).TotalPP Then
                                        PlayerPokemons(Index).Data(i).Moveset(X).CurPP = PlayerPokemons(Index).Data(i).Moveset(X).TotalPP
                                        PlayerPokemons(Index).Data(i).Moveset(X).CD = 0
                                        gothealed = True
                                    End If
                                End If
                            Next
                        End If
                    Next
                    If Player(Index, TempPlayer(Index).UseChar).CurHp < GetPlayerHP(Player(Index, TempPlayer(Index).UseChar).Level) Then
                        Player(Index, TempPlayer(Index).UseChar).CurHp = GetPlayerHP(Player(Index, TempPlayer(Index).UseChar).Level)
                        gothealed = True
                    End If
                    If Player(Index, TempPlayer(Index).UseChar).Status > 0 Then
                        Player(Index, TempPlayer(Index).UseChar).Status = 0
                        Player(Index, TempPlayer(Index).UseChar).IsConfuse = False
                        gothealed = True
                    End If
                    If gothealed Then
                        Select Case TempPlayer(Index).CurLanguage
                            Case LANG_PT: AddAlert Index, "Pokemon HP and PP restored", White
                            Case LANG_EN: AddAlert Index, "Pokemon HP and PP restored", White
                            Case LANG_ES: AddAlert Index, "Pokemon HP and PP restored", White
                        End Select
                        SendPlayerPokemons Index
                        SendPlayerVital Index
                        SendPlayerPokemonStatus Index
                        SendPlayerStatus Index
                    End If
                Case MapAttribute.Checkpoint
                    .CheckMap = Map(.Map).Tile(.X, .Y).Data1
                    .CheckX = Map(.Map).Tile(.X, .Y).Data2
                    .CheckY = Map(.Map).Tile(.X, .Y).Data3
                    .CheckDir = Map(.Map).Tile(.X, .Y).Data4
                Case MapAttribute.WarpCheckpoint
                    If .CheckMap > 0 Then
                        PlayerWarp Index, .CheckMap, .CheckX, .CheckY, .CheckDir
                    End If
            End Select
        End If
    End With
End Sub

Public Sub SpawnPlayerPokemon(ByVal Index As Long, ByVal PokeSlot As Byte)
Dim MapNum As Long
Dim statX As Byte
Dim startPosX As Long, startPosY As Long
Dim X As Long, Y As Long
Dim canSpawn As Boolean
Dim UsedBall As Byte

    If Index <= 0 Or Index > MAX_PLAYER Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Or TempPlayer(Index).UseChar > MAX_PLAYERCHAR Then Exit Sub
    If PlayerPokemon(Index).Num > 0 Then Exit Sub
    If PlayerPokemons(Index).Data(PokeSlot).Num <= 0 Then Exit Sub
    
    MapNum = Player(Index, TempPlayer(Index).UseChar).Map
    
    '//Update Position
    With PlayerPokemon(Index)
        canSpawn = False
        For X = Player(Index, TempPlayer(Index).UseChar).X - 1 To Player(Index, TempPlayer(Index).UseChar).X + 1
            For Y = Player(Index, TempPlayer(Index).UseChar).Y - 1 To Player(Index, TempPlayer(Index).UseChar).Y + 1
                If X = Player(Index, TempPlayer(Index).UseChar).X And Y = Player(Index, TempPlayer(Index).UseChar).Y Then
                    
                Else
                    '//Check if OpenTile
                    If CheckOpenTile(MapNum, X, Y) Then
                        startPosX = X
                        startPosY = Y
                        canSpawn = True
                        Exit For
                    End If
                End If
            Next
        Next
        
        If canSpawn Then
            .Num = PlayerPokemons(Index).Data(PokeSlot).Num
            .X = startPosX
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
            UsedBall = PlayerPokemons(Index).Data(.slot).BallUsed
            
            .StatusDamage = 0
            .StatusMove = 0
        Else
            Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "Out of space", White
                Case LANG_EN: AddAlert Index, "Out of space", White
                Case LANG_ES: AddAlert Index, "Out of space", White
            End Select
        End If
    End With
    
    '//Update
    If canSpawn Then SendPlayerPokemonData Index, MapNum, , YES, 0, startPosX, startPosY, UsedBall
End Sub

Public Sub ClearPlayerPokemon(ByVal Index As Long)
Dim MapNum As Long
Dim endPosX As Long, endPosY As Long
Dim BallUsed As Byte

    If Index <= 0 Or Index > MAX_PLAYER Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Or TempPlayer(Index).UseChar > MAX_PLAYERCHAR Then Exit Sub
    If PlayerPokemon(Index).Num <= 0 Then Exit Sub
    
    MapNum = Player(Index, TempPlayer(Index).UseChar).Map
    
    '//Update Position
    With PlayerPokemon(Index)
        BallUsed = PlayerPokemons(Index).Data(.slot).BallUsed
        
        .Num = 0
        endPosX = .X
        endPosY = .Y
        .X = 0
        .Y = 0
        .Dir = 0
        
        .slot = 0
    End With
    
    '//Update
    SendPlayerPokemonData Index, MapNum, , YES, 1, endPosX, endPosY, BallUsed
End Sub

Public Sub PlayerPokemonWarp(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal Dir As Byte)
Dim MapNum As Long

    '//Exit out when error
    If Index <= 0 Or Index > MAX_PLAYER Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Or TempPlayer(Index).UseChar > MAX_PLAYERCHAR Then Exit Sub
    If MapNum <= 0 Or MapNum > MAX_MAP Then Exit Sub
    If PlayerPokemon(Index).Num <= 0 Then Exit Sub
    
    '//Correct error position
    If X <= 0 Then X = 0
    If X > Map(MapNum).MaxX Then X = Map(MapNum).MaxX
    If Y <= 0 Then Y = 0
    If Y > Map(MapNum).MaxY Then Y = Map(MapNum).MaxY
    
    MapNum = Player(Index, TempPlayer(Index).UseChar).Map
    
    '//Update position
    With PlayerPokemon(Index)
        .X = X
        .Y = Y
        .Dir = Dir
    End With
    
    '//Add log
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " pokemon has been warped on Map#" & MapNum & " x:" & X & " y:" & Y
End Sub

Public Sub PlayerPokemonMove(ByVal Index As Long, ByVal Dir As Byte, Optional ByVal sendToSelf As Boolean = False)
Dim DidMove As Boolean
Dim OldX As Long, OldY As Long
Dim MapNum As Long
Dim dX As Long, dY As Long

    '//Exit out when error
    If Not IsPlaying(Index) Then Exit Sub
    If Index <= 0 Or Index > MAX_PLAYER Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Or TempPlayer(Index).UseChar > MAX_PLAYERCHAR Then Exit Sub
    If Dir < 0 Or Dir > DIR_RIGHT Then Exit Sub
    If PlayerPokemon(Index).Num <= 0 Then Exit Sub

    DidMove = False
    
    MapNum = Player(Index, TempPlayer(Index).UseChar).Map
    
    With PlayerPokemon(Index)
        '//Store original location in case it got desync
        OldX = .X
        OldY = .Y
        
        Select Case Dir
            Case DIR_UP
                .Dir = DIR_UP
                
                '//Check to make sure not outside of boundries
                If .Y > 0 Then
                    If Not CheckDirection(MapNum, DIR_UP, .X, .Y) Then
                        '//Check Distance
                        dX = .X - Player(Index, TempPlayer(Index).UseChar).X
                        dY = (.Y - 1) - Player(Index, TempPlayer(Index).UseChar).Y
                            
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
                    If Not CheckDirection(MapNum, DIR_DOWN, .X, .Y) Then
                        '//Check Distance
                        dX = .X - Player(Index, TempPlayer(Index).UseChar).X
                        dY = (.Y + 1) - Player(Index, TempPlayer(Index).UseChar).Y
                            
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
                If .X > 0 Then
                    If Not CheckDirection(MapNum, DIR_LEFT, .X, .Y) Then
                        '//Check Distance
                        dX = (.X - 1) - Player(Index, TempPlayer(Index).UseChar).X
                        dY = .Y - Player(Index, TempPlayer(Index).UseChar).Y
                            
                        '//Make sure we get a positive value
                        If dX < 0 Then dX = dX * -1
                        If dY < 0 Then dY = dY * -1
                            
                        If Not (dX <= MAX_DISTANCE And dY <= MAX_DISTANCE) Then
                            DidMove = False
                        Else
                            .X = .X - 1
                            DidMove = True
                        End If
                    End If
                End If
            Case DIR_RIGHT
                .Dir = DIR_RIGHT
                
                '//Check to make sure not outside of boundries
                If .X < Map(MapNum).MaxX Then
                    If Not CheckDirection(MapNum, DIR_RIGHT, .X, .Y) Then
                        '//Check Distance
                        dX = (.X + 1) - Player(Index, TempPlayer(Index).UseChar).X
                        dY = .Y - Player(Index, TempPlayer(Index).UseChar).Y
                            
                        '//Make sure we get a positive value
                        If dX < 0 Then dX = dX * -1
                        If dY < 0 Then dY = dY * -1
                            
                        If Not (dX <= MAX_DISTANCE And dY <= MAX_DISTANCE) Then
                            DidMove = False
                        Else
                            .X = .X + 1
                            DidMove = True
                        End If
                    End If
                End If
        End Select
    
        '//Got Desynced
        If Not DidMove Then
            .X = OldX
            .Y = OldY
            SendPlayerPokemonXY Index
            SendPlayerPokemonXY Index, True
        Else
            SendPlayerPokemonMove Index, sendToSelf
        End If
    End With
End Sub

' ******************
' ** Player Logic **
' ******************
Public Sub JoinGame(ByVal Index As Long, Optional ByVal CurLanguage As Byte = 0)
    Dim countOnline As Long

    '//Exit out if not connected
    If Not IsConnected(Index) Then Exit Sub
    '//Exit out if already playing
    If TempPlayer(Index).InGame Then Exit Sub

    frmServer.lvwInfo.ListItems(Index).SubItems(1) = GetPlayerIP(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(2) = GetPlayerLogin(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(3) = Player(Index, TempPlayer(Index).UseChar).Name

    '//Check if staff only
    If frmServer.chkStaffOnly.Value = YES Then
        If Player(Index, TempPlayer(Index).UseChar).Access <= 0 Then
            Select Case CurLanguage
            Case LANG_PT: AddAlert Index, "Server is available for Staff Members only", White
            Case LANG_EN: AddAlert Index, "Server is available for Staff Members only", White
            Case LANG_ES: AddAlert Index, "Server is available for Staff Members only", White
            End Select
            Exit Sub
        End If
    End If

    '//Load Player Pokemon
    If TempPlayer(Index).UseChar > 0 Then
        LoadPlayerInv Index, TempPlayer(Index).UseChar
        LoadPlayerPokemons Index, TempPlayer(Index).UseChar
        LoadPlayerInvStorage Index, TempPlayer(Index).UseChar
        LoadPlayerPokemonStorage Index, TempPlayer(Index).UseChar
        LoadPlayerPokedex Index, TempPlayer(Index).UseChar
    End If

    '//Set player in-game
    TempPlayer(Index).InGame = True
    TempPlayer(Index).CurLanguage = CurLanguage
    TempPlayer(Index).MapSwitchTmr = YES
    '//Send Data to Client

    '//Send Data
    AddAlert Index, "Loading Npcs...", White, , YES
    SendNpcs Index
    AddAlert Index, "Loading Pokemons...", White, , YES
    SendPokemons Index
    AddAlert Index, "Loading Items...", White, , YES
    SendItems Index
    AddAlert Index, "Loading Moves...", White, , YES
    SendPokemonMoves Index
    AddAlert Index, "Loading Animations...", White, , YES
    SendAnimations Index
    AddAlert Index, "Loading Spawns...", White, , YES
    SendSpawns Index
    'If Player(Index, TempPlayer(Index).UseChar).Access > ACCESS_MAPPER Then
    '   SendConversations Index
    'End If


    AddAlert Index, "Loading Shop...", White, , YES
    SendShops Index
    AddAlert Index, "Loading Quest...", White, , YES
    SendQuests Index
    AddAlert Index, "Loading Inventory...", White, , YES
    SendPlayerInv Index
    AddAlert Index, "Loading Item Storage...", White, , YES
    SendPlayerInvStorage Index
    AddAlert Index, "Loading Team...", White, , YES
    SendPlayerPokemons Index
    AddAlert Index, "Loading Pokemon Box...", White, , YES
    SendPlayerPokemonStorage Index
    AddAlert Index, "Loading Pokedex...", White, , YES
    SendPlayerPokedex Index
    AddAlert Index, "Send Raking To Client...", White, , YES
    SendRankTo Index
    AddAlert Index, "Send Event Exp To Client...", White, , YES
    SendEventInfo Index

    If Player(Index, TempPlayer(Index).UseChar).Access = ACCESS_NONE Then
        UpdateRank Trim$(Player(Index, TempPlayer(Index).UseChar).Name), Player(Index, TempPlayer(Index).UseChar).Level, Player(Index, TempPlayer(Index).UseChar).CurExp
    End If
    'LoadRank

    '//Send data to position
    With Player(Index, TempPlayer(Index).UseChar)
        PlayerWarp Index, .Map, .X, .Y, .Dir

        '//Check online
        countOnline = TotalPlayerOnline

        If .Access < ACCESS_CREATOR Then
            SendMapMsg .Map, Trim$(.Name) & " has joined the game", White
        End If
        'AddLog Trim$(.Name) & " has joined the game"

        '//Send count msg
        If countOnline > 1 Then
            SendPlayerMsg Index, "There are " & (countOnline - 1) & " other players online", White
        Else
            SendPlayerMsg Index, "There are no other players online", White
        End If
    End With

    '//Send Message
    SendPlayerMsg Index, "Welcome to " & GAME_NAME, White
    If Len(Trim$(Options.MOTD)) > 0 Then
        SendPlayerMsg Index, Trim$(Options.MOTD), White
    End If
    '//Send tutorial message
    If CountPlayerPokemon(Index) <= 0 Then
        '//Init Starter Pokemon
        TempPlayer(Index).CurConvoNum = 1
        TempPlayer(Index).CurConvoData = 1
        TempPlayer(Index).CurConvoNpc = 3
        TempPlayer(Index).CurConvoMapNpc = 0
        SendInitConvo Index, TempPlayer(Index).CurConvoNum, TempPlayer(Index).CurConvoData, TempPlayer(Index).CurConvoNpc
    Else
        Player(Index, TempPlayer(Index).UseChar).DidStart = NO
        SavePlayerData Index, TempPlayer(Index).UseChar
    End If

    '//Send In-Game
    SendHighIndex Index
    SendPokemonHighIndex Index
    SendInGame Index
End Sub

Public Sub LeftGame(ByVal Index As Long)
    Dim sIP As String
    Dim i As Long, X As Byte, Y As Byte

    sIP = GetPlayerIP(Index)

    '//Update HighIndex
    If Player_HighIndex = Index Then
        Player_HighIndex = Player_HighIndex - 1
        '//Update Index to all
        SendHighIndex
    End If

    '//InGame Data
    If TempPlayer(Index).InGame Then
        '//Request
        i = TempPlayer(Index).PlayerRequest
        If i > 0 Then
            '//Cancel Request to index
            If IsPlaying(i) Then
                If TempPlayer(i).UseChar > 0 Then
                    If TempPlayer(i).PlayerRequest = Index Then
                        If TempPlayer(Index).RequestType = 1 Then  '//1 Duel
                            '//Check if already in duel
                            If TempPlayer(Index).InDuel > 0 Then
                                SendActionMsg Player(i, TempPlayer(i).UseChar).Map, "Win!", Player(i, TempPlayer(i).UseChar).X * 32, Player(i, TempPlayer(i).UseChar).Y * 32, White
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
                        ElseIf TempPlayer(Index).RequestType = 2 Then    '//trade
                            '//Check if already in trade
                            If TempPlayer(Index).InTrade > 0 Then
                                TempPlayer(i).InTrade = 0
                                For X = 1 To MAX_TRADE
                                    Call ZeroMemory(ByVal VarPtr(TempPlayer(i).TradeItem(X)), LenB(TempPlayer(i).TradeItem(X)))
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
                        ElseIf TempPlayer(Index).RequestType = 3 Then    '//Party
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
        If TempPlayer(Index).InParty > 0 Then
            LeaveParty Index
        End If

        TempPlayer(Index).InDuel = 0
        TempPlayer(Index).DuelTime = 0
        TempPlayer(Index).DuelTimeTmr = 0
        TempPlayer(Index).WarningTimer = 0
        TempPlayer(Index).PlayerRequest = 0
        TempPlayer(Index).RequestType = 0
        TempPlayer(Index).InTrade = 0
        For X = 1 To MAX_TRADE
            Call ZeroMemory(ByVal VarPtr(TempPlayer(Index).TradeItem(X)), LenB(TempPlayer(Index).TradeItem(X)))
        Next
        TempPlayer(Index).TradeMoney = 0
        TempPlayer(Index).TradeSet = 0
        TempPlayer(Index).PlayerRequest = 0
        TempPlayer(Index).RequestType = 0

        If Player(Index, TempPlayer(Index).UseChar).Access = ACCESS_NONE Then
            If TempPlayer(Index).UseChar > 0 Then
                UpdateRank Trim$(Player(Index, TempPlayer(Index).UseChar).Name), Player(Index, TempPlayer(Index).UseChar).Level, Player(Index, TempPlayer(Index).UseChar).CurExp
            End If
        End If

        TempPlayer(Index).InGame = False

        '//Clear In-Game Data

        '//Save Player data
        SavePlayerDatas Index

        '//Left Game
        SendLeftGame Index

        If TempPlayer(Index).UseChar > 0 Then
            If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_CREATOR Then
                SendMapMsg Player(Index, TempPlayer(Index).UseChar).Map, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " has left the game", White
            End If
            'AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " has left the game"
        End If
    End If

    '//Clear Player Data
    ClearTempPlayer Index
    ClearPlayer Index
    ClearPlayerInv Index
    ClearPlayerInvStorage Index
    ClearPlayerPokemons Index
    ClearPlayerPokemonStorage Index
    ClearAccount Index
    ClearPlayerPokedex Index

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

Public Function FindSameItemSlot(ByVal Index As Long, ByVal ItemNum As Long) As Byte
Dim i As Byte

    FindSameItemSlot = 0
    
    If Not IsPlaying(Index) Then Exit Function
    If TempPlayer(Index).UseChar <= 0 Then Exit Function
    
    For i = 1 To MAX_PLAYER_INV
        With PlayerInv(Index).Data(i)
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

Public Function FindFreeInvSlot(ByVal Index As Long, ByVal ItemNum As Long) As Byte
    Dim i As Byte

    FindFreeInvSlot = 0

    If Not IsPlaying(Index) Then Exit Function
    If TempPlayer(Index).UseChar <= 0 Then Exit Function

    If Item(ItemNum).Stock = YES Then
        i = FindSameItemSlot(Index, ItemNum)
        If i > 0 Then
            FindFreeInvSlot = i
            Exit Function
        End If
    End If

    For i = 1 To MAX_PLAYER_INV
        With PlayerInv(Index).Data(i)
            If .Locked = NO Then
                If .Num = 0 Then
                    FindFreeInvSlot = i
                    Exit Function
                End If
            End If
        End With
    Next
End Function

Public Function TryGivePlayerItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long) As Boolean
    TryGivePlayerItem = True
    If Not GiveItem(Index, ItemNum, ItemVal) Then
        '//Error msg
        Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: AddAlert Index, "Inventory is full", White
            Case LANG_EN: AddAlert Index, "Inventory is full", White
            Case LANG_ES: AddAlert Index, "Inventory is full", White
        End Select
        TryGivePlayerItem = False
    Else
        '//Check if there's still free slot
        If CountFreeInvSlot(Index) <= 5 Then
            Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "Warning: Your inventory is almost full", White
                Case LANG_EN: AddAlert Index, "Warning: Your inventory is almost full", White
                Case LANG_ES: AddAlert Index, "Warning: Your inventory is almost full", White
            End Select
        End If
    End If
End Function

Public Function CountFreeInvSlot(ByVal Index As Long) As Long
Dim count As Long, i As Long

    CountFreeInvSlot = 0
    count = 0
    
    If Not IsPlaying(Index) Then Exit Function
    If TempPlayer(Index).UseChar <= 0 Then Exit Function
    
    For i = 1 To MAX_PLAYER_INV
        With PlayerInv(Index).Data(i)
            If .Num = 0 Then
                count = count + 1
            End If
        End With
    Next
    
    CountFreeInvSlot = count
End Function

Public Function GiveItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long) As Boolean
Dim i As Byte

    '//Get Slot
    i = FindFreeInvSlot(Index, ItemNum)
    
    '//Got slot
    If i > 0 Then
        With PlayerInv(Index).Data(i)
            .Num = ItemNum
            .Value = .Value + ItemVal
        End With
        '//Update
        SendPlayerInvSlot Index, i
        GiveItem = True
    Else
        GiveItem = False
    End If
End Function

'//Player Pokemon
Public Function FindOpenPokeSlot(ByVal Index As Long) As Long
Dim i As Byte

    For i = 1 To MAX_PLAYER_POKEMON
        If PlayerPokemons(Index).Data(i).Num = 0 Then
            FindOpenPokeSlot = i
            Exit Function
        End If
    Next
End Function

Public Sub GivePlayerPokemon(ByVal Index As Long, ByVal PokeNum As Long, ByVal Level As Long, ByVal BallUsed As Byte, Optional ByVal IsShiny As Byte = NO, _
                             Optional ByVal IVFull As Byte = NO, Optional ByVal TheNature As Integer = -1)
    Dim i As Long, X As Byte, m As Long, s As Byte, slot As Byte, storageSlot As Byte, gotSlot As Byte

    i = FindOpenPokeSlot(Index)

    '//Got slot
    If i > 0 Then
        With PlayerPokemons(Index).Data(i)
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
            For X = 1 To StatEnum.Stat_Count - 1
                .Stat(X).EV = 0
                .Stat(X).IV = 15    '//Default Stat
                If IVFull > 0 Then .Stat(X).IV = 31    'Peronalização do painel admin
                .Stat(X).Value = CalculatePokemonStat(X, .Num, .Level, .Stat(X).EV, .Stat(X).IV, .Nature)
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
            AddPlayerPokedex Index, .Num, YES, YES
        End With
        '//Update
        SendPlayerPokemonSlot Index, i
    Else
        For storageSlot = 1 To MAX_STORAGE_SLOT
            gotSlot = FindFreePokeStorageSlot(Index, storageSlot)
            If gotSlot > 0 Then
                With PlayerPokemonStorage(Index).slot(storageSlot).Data(gotSlot)

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
                    For X = 1 To StatEnum.Stat_Count - 1
                        .Stat(X).EV = 0
                        .Stat(X).IV = 15    '//Default Stat
                        If IVFull > 0 Then .Stat(X).IV = 31    'Peronalização do painel admin
                        .Stat(X).Value = CalculatePokemonStat(X, .Num, .Level, .Stat(X).EV, .Stat(X).IV, .Nature)
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
                    AddPlayerPokedex Index, .Num, YES, YES

                    Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "Your pokemon has been transferred to your pokemon storage", White
                    Case LANG_EN: AddAlert Index, "Your pokemon has been transferred to your pokemon storage", White
                    Case LANG_ES: AddAlert Index, "Your pokemon has been transferred to your pokemon storage", White
                    End Select
                    SendPlayerPokemonStorageSlot Index, storageSlot, gotSlot
                    Exit Sub
                End With
            End If
        Next storageSlot
    End If
End Sub

Public Sub UpdatePlayerPokemonOrder(ByVal Index As Long)
Dim i As Long

    For i = 2 To MAX_PLAYER_POKEMON
        With PlayerPokemons(Index)
            '//Check if previous number is empty
            If .Data(i - 1).Num = 0 Then
                '//Move Data
                .Data(i - 1) = .Data(i)
                Call ZeroMemory(ByVal VarPtr(.Data(i)), LenB(.Data(i)))
            End If
        End With
    Next
End Sub

Public Function CountPlayerPokemon(ByVal Index As Long) As Byte
Dim i As Byte
Dim count As Byte

    count = 0
    For i = 1 To MAX_PLAYER_POKEMON
        With PlayerPokemons(Index).Data(i)
            If .Num > 0 Then
                count = count + 1
            End If
        End With
    Next
    CountPlayerPokemon = count
End Function

Public Function CountPlayerPokemonAlive(ByVal Index As Long) As Byte
Dim i As Byte
Dim count As Byte

    count = 0
    For i = 1 To MAX_PLAYER_POKEMON
        With PlayerPokemons(Index).Data(i)
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
Public Sub GivePlayerPokemonExp(ByVal Index As Long, ByVal PokeSlot As Byte, ByVal Exp As Long)
    '//Check Error
    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If PokeSlot <= 0 Or PokeSlot > MAX_PLAYER_POKEMON Then Exit Sub
    If PlayerPokemons(Index).Data(PokeSlot).Num <= 0 Then Exit Sub
    
    'Exp Rate
    If EventExp.ExpEvent Then
        Exp = Exp * EventExp.ExpMultiply
    End If
    
    '//Add Exp
    With PlayerPokemons(Index).Data(PokeSlot)
        '//Make sure we can give it exp based on player level
        If Player(Index, TempPlayer(Index).UseChar).Level + 10 <= .Level Then Exit Sub
        If .Level >= MAX_LEVEL Then Exit Sub
        
        .CurExp = .CurExp + Exp
        TextAdd frmServer.txtLog, "EXP: " & Exp
        
        '//ActionMsg
        If PlayerPokemon(Index).Num > 0 Then
            If PlayerPokemon(Index).slot = PokeSlot Then
                SendActionMsg Player(Index, TempPlayer(Index).UseChar).Map, "+" & Exp, PlayerPokemon(Index).X * 32, PlayerPokemon(Index).Y * 32, White
            End If
        End If
    End With
    CheckPlayerPokemonLevelUp Index, PokeSlot
End Sub

Public Function GivePlayerEvPowerBracer(ByVal Index As Long, ByVal PokeSlot As Byte) As Boolean
    Dim CallBack As Integer
    GivePlayerEvPowerBracer = False

    With PlayerPokemons(Index).Data(PokeSlot)
        If .HeldItem > 0 Then
            If Item(.HeldItem).Type = ItemTypeEnum.PowerBracer Then
                If Item(.HeldItem).Data1 >= StatEnum.HP And Item(.HeldItem).Data1 <= StatEnum.Spd Then
                    GivePlayerEvPowerBracer = True
                    CallBack = GivePlayerPokemonEVExp(Index, PokeSlot, Item(.HeldItem).Data1, Item(.HeldItem).Data2)
                End If
            End If
        End If
    End With
End Function

Public Function GivePlayerPokemonEVExp(ByVal Index As Long, ByVal PokeSlot As Byte, ByVal evStat As StatEnum, ByVal Exp As Long) As Integer
    Dim CountStat As Long, X As Byte, statMaxEv As Integer, Sobra As Integer

    '// Função implementada pra utilizar => Recebendo ao matar um poke,
    '                                       Ao utilizar items Barries
    '                                       Ao utilizar items Protein
    '                                       Ao utilizar Power Bracer no pokemon.

    With PlayerPokemons(Index).Data(PokeSlot)
        ' Máximo de EV Total
        ' MAX_EV = 510

        ' Máximo de Ev em cada atributo
        statMaxEv = 252

        ' Faz a contagem do total de EV
        CountStat = 0
        For X = 1 To StatEnum.Stat_Count - 1
            CountStat = CountStat + PlayerPokemons(Index).Data(PokeSlot).Stat(X).EV
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
                    Sobra = -Exp ' Sobra é a quantidade retirada como um número positivo
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
                SendPlayerPokemonSlot Index, PokeSlot
            End If
        End If

        SendPlayerPokemonsStat Index, PokeSlot

    End With
End Function

Private Sub CheckPlayerPokemonLevelUp(ByVal Index As Long, ByVal PokeSlot As Byte)
Dim ExpRollover As Long
Dim statNu As Byte
Dim oldlevel As Long, levelcount As Long
Dim i As Long
Dim DidLevel As Boolean

    '//Check Error
    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If PokeSlot <= 0 Or PokeSlot > MAX_PLAYER_POKEMON Then Exit Sub
    If PlayerPokemons(Index).Data(PokeSlot).Num <= 0 Then Exit Sub
    
    '//Add Exp
    With PlayerPokemons(Index).Data(PokeSlot)
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
        SendPlayerPokemonSlot Index, PokeSlot
        
        '//Check New Move
        If levelcount > 0 Then
            SendPlaySound "levelup.wav", Player(Index, TempPlayer(Index).UseChar).Map
            SendPlayerPokemonVital Index
            CheckNewMove Index, PokeSlot
        End If
        
        'SendPlayerPokemonVital Index
    End With
End Sub

Public Function FindFreeMoveSlot(ByVal Index As Long, ByVal PokeSlot As Byte, Optional ByVal MoveSlot As Byte = 0) As Long
Dim i As Byte
Dim foundsameslot As Boolean

    '//Check Error
    If Not IsPlaying(Index) Then Exit Function
    If TempPlayer(Index).UseChar <= 0 Then Exit Function
    If PokeSlot <= 0 Or PokeSlot > MAX_PLAYER_POKEMON Then Exit Function
    If PlayerPokemons(Index).Data(PokeSlot).Num <= 0 Then Exit Function

    foundsameslot = False
    With PlayerPokemons(Index).Data(PokeSlot)
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

Public Sub CheckNewMove(ByVal Index As Long, ByVal PokeSlot As Byte, Optional ByVal StartIndex As Long = 1)
Dim i As Byte, X As Byte
Dim FoundMatch As Boolean
Dim MoveSlot As Byte
Dim Continue As Boolean

    '//Check Error
    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If PokeSlot <= 0 Or PokeSlot > MAX_PLAYER_POKEMON Then Exit Sub
    If PlayerPokemons(Index).Data(PokeSlot).Num <= 0 Then Exit Sub
    If TempPlayer(Index).MoveLearnNum > 0 Then Exit Sub
    If StartIndex <= 0 Then Exit Sub
    
    '//Add Exp
    With PlayerPokemons(Index).Data(PokeSlot)
        '//Check New Move
        For i = StartIndex To MAX_POKEMON_MOVESET
            If Pokemon(.Num).Moveset(i).MoveNum > 0 Then
                If Pokemon(.Num).Moveset(i).MoveLevel = .Level Then
                    Continue = False
                    '//Make sure move doesn't exist
                    For X = 1 To MAX_MOVESET
                        If .Moveset(X).Num = Pokemon(.Num).Moveset(i).MoveNum Then
                            Continue = True
                        End If
                    Next
                    If Not Continue Then
                        '//Check if there's available slot
                        MoveSlot = FindFreeMoveSlot(Index, PokeSlot)
                        If MoveSlot >= 0 Then
                            If MoveSlot > 0 Then
                                .Moveset(MoveSlot).Num = Pokemon(.Num).Moveset(i).MoveNum
                                .Moveset(MoveSlot).TotalPP = PokemonMove(Pokemon(.Num).Moveset(i).MoveNum).PP
                                .Moveset(MoveSlot).CurPP = .Moveset(MoveSlot).TotalPP
                                SendPlayerPokemonSlot Index, PokeSlot
                                '//Send Msg
                                SendPlayerMsg Index, Trim$(Pokemon(.Num).Name) & " learned the move " & Trim$(PokemonMove(Pokemon(.Num).Moveset(i).MoveNum).Name), White
                            Else
                                '//Proceed to ask
                                TempPlayer(Index).MoveLearnPokeSlot = PokeSlot
                                TempPlayer(Index).MoveLearnNum = Pokemon(.Num).Moveset(i).MoveNum
                                TempPlayer(Index).MoveLearnIndex = i + 1
                                SendNewMove Index
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End With
End Sub

Public Sub PlayerUseItem(ByVal Index As Long, ByVal invSlot As Byte)
    Dim ItemNum As Long
    Dim gothealed As Boolean
    Dim X As Long
    Dim exproll As Long
    Dim Exp As Long
    Dim i As Long, CanLearn As Boolean
    Dim BerriesFunc As Integer, PokeName As String

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If invSlot <= 0 Or invSlot > MAX_PLAYER_INV Then Exit Sub
    If PlayerInv(Index).Data(invSlot).Num <= 0 Then Exit Sub
    If PlayerInv(Index).Data(invSlot).Value <= 0 Then Exit Sub
    If TempPlayer(Index).InDuel > 0 Then Exit Sub
    If TempPlayer(Index).InNpcDuel > 0 Then Exit Sub

    ItemNum = PlayerInv(Index).Data(invSlot).Num

    Select Case Item(ItemNum).Type
    Case ItemTypeEnum.pokeBall
        '//Catching
        If Map(Player(Index, TempPlayer(Index).UseChar).Map).Moral = 3 Then
            If Not ItemNum = 12 Then
                AddAlert Index, "You cannot use this type of Pokeball here", White
                Exit Sub
            End If
        Else
            If ItemNum = 12 Then
                AddAlert Index, "You cannot use this type of Pokeball here", White
                Exit Sub
            End If
        End If
        TempPlayer(Index).TmpUseInvSlot = invSlot
        SendGetData Index, ItemTypeEnum.pokeBall, invSlot
    Case ItemTypeEnum.Medicine
        Select Case Item(ItemNum).Data1    '//Type
        Case 1    '// Heal HP
            gothealed = False
            If PlayerPokemon(Index).Num > 0 Then
                If PlayerPokemon(Index).slot > 0 Then
                    If PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp < PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).MaxHp Then
                        PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp = PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp + Item(ItemNum).Data2
                        If PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp > PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).MaxHp Then
                            PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp = PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).MaxHp
                        End If
                        gothealed = True
                    End If
                End If
            End If
            If gothealed Then
                Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "Pokemon HP restored", White
                Case LANG_EN: AddAlert Index, "Pokemon HP restored", White
                Case LANG_ES: AddAlert Index, "Pokemon HP restored", White
                End Select
                SendPlayerPokemonVital Index

                '//Take Item
                PlayerInv(Index).Data(invSlot).Value = PlayerInv(Index).Data(invSlot).Value - 1
                If PlayerInv(Index).Data(invSlot).Value <= 0 Then
                    '//Clear Item
                    PlayerInv(Index).Data(invSlot).Num = 0
                    PlayerInv(Index).Data(invSlot).Value = 0
                End If
                SendPlayerInvSlot Index, invSlot
            End If
        Case 2    '// Give Exp
        Case 3    '// Heal PP
            gothealed = False
            If PlayerPokemon(Index).Num > 0 Then
                If PlayerPokemon(Index).slot > 0 Then
                    For X = 1 To MAX_MOVESET
                        If PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Moveset(X).Num > 0 Then
                            If PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Moveset(X).CurPP < PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Moveset(X).TotalPP Then
                                PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Moveset(X).CurPP = PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Moveset(X).CurPP + Item(ItemNum).Data2
                                If PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Moveset(X).CurPP > PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Moveset(X).TotalPP Then
                                    PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Moveset(X).CurPP = PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Moveset(X).TotalPP
                                End If
                                PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Moveset(X).CD = 0
                                gothealed = True
                            End If
                        End If
                    Next
                End If
            End If
            If gothealed Then
                Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "Pokemon PP restored", White
                Case LANG_EN: AddAlert Index, "Pokemon PP restored", White
                Case LANG_ES: AddAlert Index, "Pokemon PP restored", White
                End Select
                For X = 1 To MAX_MOVESET
                    SendPlayerPokemonPP Index, X
                Next
                '//Take Item
                PlayerInv(Index).Data(invSlot).Value = PlayerInv(Index).Data(invSlot).Value - 1
                If PlayerInv(Index).Data(invSlot).Value <= 0 Then
                    '//Clear Item
                    PlayerInv(Index).Data(invSlot).Num = 0
                    PlayerInv(Index).Data(invSlot).Value = 0
                End If
                SendPlayerInvSlot Index, invSlot
            End If
        Case 4    '// Revive
            TempPlayer(Index).TmpUseInvSlot = invSlot
            SendGetData Index, ItemTypeEnum.Medicine, invSlot
        Case 5    '// Cure Status
            gothealed = False
            If Item(ItemNum).Data2 > 0 Then
                If PlayerPokemon(Index).Num > 0 Then
                    If PlayerPokemon(Index).slot > 0 Then
                        If PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Status = Item(ItemNum).Data2 Then
                            PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Status = 0
                            gothealed = True
                        End If
                    End If
                End If
            Else
                If PlayerPokemon(Index).slot > 0 Then
                    If PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Status > 0 Then
                        PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Status = 0
                        gothealed = True
                    End If
                End If
            End If
            If gothealed Then
                Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "Pokemon Status removed", White
                Case LANG_EN: AddAlert Index, "Pokemon Status removed", White
                Case LANG_ES: AddAlert Index, "Pokemon Status removed", White
                End Select
                SendPlayerPokemonStatus Index
                '//Take Item
                PlayerInv(Index).Data(invSlot).Value = PlayerInv(Index).Data(invSlot).Value - 1
                If PlayerInv(Index).Data(invSlot).Value <= 0 Then
                    '//Clear Item
                    PlayerInv(Index).Data(invSlot).Num = 0
                    PlayerInv(Index).Data(invSlot).Value = 0
                End If
                SendPlayerInvSlot Index, invSlot
            End If
        Case 6    '// Heal Trainer
            gothealed = False
            If Player(Index, TempPlayer(Index).UseChar).CurHp < GetPlayerHP(Player(Index, TempPlayer(Index).UseChar).Level) Then
                Player(Index, TempPlayer(Index).UseChar).CurHp = Player(Index, TempPlayer(Index).UseChar).CurHp + Item(PlayerInv(Index).Data(invSlot).Num).Data2
                If Player(Index, TempPlayer(Index).UseChar).CurHp > GetPlayerHP(Player(Index, TempPlayer(Index).UseChar).Level) Then
                    Player(Index, TempPlayer(Index).UseChar).CurHp = GetPlayerHP(Player(Index, TempPlayer(Index).UseChar).Level)
                End If
                gothealed = True
            End If

            If gothealed Then
                Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "HP restored", White
                Case LANG_EN: AddAlert Index, "HP restored", White
                Case LANG_ES: AddAlert Index, "HP restored", White
                End Select
                SendPlayerVital Index
                '//Take Item
                PlayerInv(Index).Data(invSlot).Value = PlayerInv(Index).Data(invSlot).Value - 1
                If PlayerInv(Index).Data(invSlot).Value <= 0 Then
                    '//Clear Item
                    PlayerInv(Index).Data(invSlot).Num = 0
                    PlayerInv(Index).Data(invSlot).Value = 0
                End If
                SendPlayerInvSlot Index, invSlot
            End If
        Case 7    '// Cure Trainer
            gothealed = False
            If Item(ItemNum).Data2 > 0 Then
                If Player(Index, TempPlayer(Index).UseChar).Status = Item(ItemNum).Data2 Then
                    Player(Index, TempPlayer(Index).UseChar).Status = 0
                    gothealed = True
                End If
            Else
                If Player(Index, TempPlayer(Index).UseChar).Status > 0 Then
                    Player(Index, TempPlayer(Index).UseChar).Status = 0
                    gothealed = True
                End If
            End If
            If gothealed Then
                Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "Status was removed", White
                Case LANG_EN: AddAlert Index, "Status was removed", White
                Case LANG_ES: AddAlert Index, "Status was removed", White
                End Select
                SendPlayerStatus Index
                '//Take Item
                PlayerInv(Index).Data(invSlot).Value = PlayerInv(Index).Data(invSlot).Value - 1
                If PlayerInv(Index).Data(invSlot).Value <= 0 Then
                    '//Clear Item
                    PlayerInv(Index).Data(invSlot).Num = 0
                    PlayerInv(Index).Data(invSlot).Value = 0
                End If
                SendPlayerInvSlot Index, invSlot
            End If
        End Select

        If Item(ItemNum).Data3 > 0 Then
            '//Level Up
            If PlayerPokemon(Index).Num > 0 Then
                If PlayerPokemon(Index).slot > 0 Then
                    exproll = GetPokemonNextExp(PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Level, Pokemon(PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Num).GrowthRate)
                    Exp = exproll - PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurExp

                    If Exp > 0 Then
                        GivePlayerPokemonExp Index, PlayerPokemon(Index).slot, Exp
                    End If

                    '//Take Item
                    PlayerInv(Index).Data(invSlot).Value = PlayerInv(Index).Data(invSlot).Value - 1
                    If PlayerInv(Index).Data(invSlot).Value <= 0 Then
                        '//Clear Item
                        PlayerInv(Index).Data(invSlot).Num = 0
                        PlayerInv(Index).Data(invSlot).Value = 0
                    End If
                    SendPlayerInvSlot Index, invSlot
                End If
            End If
        End If
    Case ItemTypeEnum.Berries

        If Item(ItemNum).Data1 > 0 Then
            If PlayerPokemon(Index).Num > 0 Then
                If PlayerPokemon(Index).slot > 0 Then
                    For i = 1 To StatEnum.Stat_Count - 1
                        If Item(ItemNum).Data1 = i Then
                            ' Adiciona ou remove a experiência (Berries/Proteins)
                            BerriesFunc = GivePlayerPokemonEVExp(Index, PlayerPokemon(Index).slot, Item(ItemNum).Data1, Item(ItemNum).Data2)
                            If BerriesFunc <> 0 Then
                                '//Take Item
                                PlayerInv(Index).Data(invSlot).Value = PlayerInv(Index).Data(invSlot).Value - 1
                                If PlayerInv(Index).Data(invSlot).Value <= 0 Then
                                    '//Clear Item
                                    PlayerInv(Index).Data(invSlot).Num = 0
                                    PlayerInv(Index).Data(invSlot).Value = 0
                                End If
                                SendPlayerInvSlot Index, invSlot

                                PokeName = Trim$(Pokemon(PlayerPokemon(Index).Num).Name)
                                If BerriesFunc > 0 Then
                                    Select Case TempPlayer(Index).CurLanguage
                                    Case LANG_PT: AddAlert Index, PokeName & " aumentou " & BerriesFunc & " pontos de EV em " & GetAtributeName(Item(ItemNum).Data1), Green
                                    Case LANG_EN: AddAlert Index, PokeName & " aumentou " & BerriesFunc & " pontos de EV em " & GetAtributeName(Item(ItemNum).Data1), Green
                                    Case LANG_ES: AddAlert Index, PokeName & " aumentou " & BerriesFunc & " pontos de EV em " & GetAtributeName(Item(ItemNum).Data1), Green
                                    End Select
                                ElseIf BerriesFunc < 0 Then
                                    Select Case TempPlayer(Index).CurLanguage
                                    Case LANG_PT: AddAlert Index, PokeName & " reduziu " & Math.Abs(BerriesFunc) & " pontos de EV em " & GetAtributeName(Item(ItemNum).Data1), Grey
                                    Case LANG_EN: AddAlert Index, PokeName & " reduziu " & Math.Abs(BerriesFunc) & " pontos de EV em " & GetAtributeName(Item(ItemNum).Data1), Grey
                                    Case LANG_ES: AddAlert Index, PokeName & " reduziu " & Math.Abs(BerriesFunc) & " pontos de EV em " & GetAtributeName(Item(ItemNum).Data1), Grey
                                    End Select
                                End If
                            Else
                                Select Case TempPlayer(Index).CurLanguage
                                Case LANG_PT: AddAlert Index, PokeName & " está no limite Min/Max de EV em " & GetAtributeName(Item(ItemNum).Data1) & " " & PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Stat(Item(ItemNum).Data1).EV, Grey
                                Case LANG_EN: AddAlert Index, PokeName & " está no limite Min/Max de EV em " & GetAtributeName(Item(ItemNum).Data1) & " " & PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Stat(Item(ItemNum).Data1).EV, Grey
                                Case LANG_ES: AddAlert Index, PokeName & " está no limite Min/Max de EV em " & GetAtributeName(Item(ItemNum).Data1) & " " & PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Stat(Item(ItemNum).Data1).EV, Grey
                                End Select
                            End If
                            Exit For
                        End If
                    Next i
                Else
                    Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "Você não está em um pokemon", White
                    Case LANG_EN: AddAlert Index, "You are not in a pokemon", White
                    Case LANG_ES: AddAlert Index, "No estas en un pokemon", White
                    End Select
                End If
            Else
                Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "Você não está em um pokemon", White
                Case LANG_EN: AddAlert Index, "You are not in a pokemon", White
                Case LANG_ES: AddAlert Index, "No estas en un pokemon", White
                End Select
            End If
        End If

    Case ItemTypeEnum.keyItems
        Select Case Item(ItemNum).Data1
        Case 1    '//Sprite Type
            If Item(ItemNum).Data2 > 0 Then
                If Map(Player(Index, TempPlayer(Index).UseChar).Map).SpriteType <= TEMP_SPRITE_GROUP_NONE Then
                    ChangeTempSprite Index, Item(ItemNum).Data2
                End If
            End If
        End Select
    Case ItemTypeEnum.TM_HM
        If PlayerPokemon(Index).Num > 0 And PlayerPokemon(Index).slot > 0 Then
            If Item(ItemNum).Data1 > 0 Then
                CanLearn = False
                For i = 1 To 110
                    If Pokemon(PlayerPokemon(Index).Num).ItemMoveset(i) = Item(ItemNum).Data1 Then
                        CanLearn = True
                        Exit For
                    End If
                Next
                '//Make sure move doesn't exist
                For i = 1 To MAX_MOVESET
                    If PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Moveset(i).Num = Item(ItemNum).Data1 Then
                        CanLearn = False
                    End If
                Next

                If CanLearn Then
                    '//Continue
                    i = FindFreeMoveSlot(Index, PlayerPokemon(Index).slot)
                    If i > 0 Then
                        PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Moveset(i).Num = Item(ItemNum).Data1
                        PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Moveset(i).TotalPP = PokemonMove(Item(ItemNum).Data1).PP
                        PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Moveset(i).CurPP = PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Moveset(i).TotalPP
                        SendPlayerPokemonSlot Index, PlayerPokemon(Index).slot
                        '//Send Msg
                        SendPlayerMsg Index, Trim$(Pokemon(PlayerPokemon(Index).Num).Name) & " learned the move " & Trim$(PokemonMove(Item(ItemNum).Data1).Name), White
                    Else
                        '//Proceed to ask
                        TempPlayer(Index).MoveLearnPokeSlot = PlayerPokemon(Index).slot
                        TempPlayer(Index).MoveLearnNum = Item(ItemNum).Data1
                        TempPlayer(Index).MoveLearnIndex = 0
                        SendNewMove Index
                    End If

                    If Item(ItemNum).Data2 > 0 Then
                        '//Take Item
                        PlayerInv(Index).Data(invSlot).Value = PlayerInv(Index).Data(invSlot).Value - 1
                        If PlayerInv(Index).Data(invSlot).Value <= 0 Then
                            '//Clear Item
                            PlayerInv(Index).Data(invSlot).Num = 0
                            PlayerInv(Index).Data(invSlot).Value = 0
                        End If
                        SendPlayerInvSlot Index, invSlot
                    End If
                Else
                    AddAlert Index, "This pokemon cannot learn this move", White
                    Exit Sub
                End If
            End If
        Else
            AddAlert Index, "Please spawn your pokemon first", White
            Exit Sub
        End If
    Case ItemTypeEnum.PowerBracer

    Case ItemTypeEnum.Items

    Case Else
        '//Not usable
        Exit Sub
    End Select

    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " use item " & Trim$(Item(ItemNum).Name)
End Sub

'//Count Free Pokemno slot
Public Function CountFreePokemonSlot(ByVal Index As Long) As Long
Dim count As Long
Dim i As Byte, X As Byte

    count = 0
    For i = 1 To MAX_PLAYER_POKEMON
        If PlayerPokemons(Index).Data(i).Num = 0 Then
            count = count + 1
        End If
    Next
    For i = 1 To MAX_STORAGE_SLOT
        If PlayerPokemonStorage(Index).slot(i).Unlocked = YES Then
            For X = 1 To MAX_STORAGE
                If PlayerPokemonStorage(Index).slot(i).Data(X).Num = 0 Then
                    count = count + 1
                End If
            Next
        End If
    Next
    CountFreePokemonSlot = count
End Function

Public Function DepositItem(ByVal Index As Long, ByVal storageSlot As Byte, ByVal StorageData As Byte, ByVal invSlot As Byte, ByVal Value As Long) As Boolean
    DepositItem = False
    
    If Not IsPlaying(Index) Then Exit Function
    If TempPlayer(Index).UseChar <= 0 Then Exit Function
    If storageSlot <= 0 Or storageSlot > 5 Then Exit Function
    If StorageData <= 0 Or StorageData > MAX_STORAGE Then Exit Function
    If invSlot <= 0 Or invSlot > MAX_PLAYER_INV Then Exit Function
    If PlayerInvStorage(Index).slot(storageSlot).Unlocked = False Then Exit Function

    PlayerInvStorage(Index).slot(storageSlot).Data(StorageData).Num = PlayerInv(Index).Data(invSlot).Num
    PlayerInvStorage(Index).slot(storageSlot).Data(StorageData).Value = PlayerInvStorage(Index).slot(storageSlot).Data(StorageData).Value + Value
    PlayerInv(Index).Data(invSlot).Value = PlayerInv(Index).Data(invSlot).Value - Value
    If PlayerInv(Index).Data(invSlot).Value <= 0 Then
        PlayerInv(Index).Data(invSlot).Value = 0
        PlayerInv(Index).Data(invSlot).Num = 0
    End If
    DepositItem = True
End Function

Public Function WithdrawItem(ByVal Index As Long, ByVal storageSlot As Byte, ByVal StorageData As Byte, ByVal invSlot As Byte, ByVal Value As Long) As Boolean
    WithdrawItem = False
    
    If Not IsPlaying(Index) Then Exit Function
    If TempPlayer(Index).UseChar <= 0 Then Exit Function
    If storageSlot <= 0 Or storageSlot > 5 Then Exit Function
    If StorageData <= 0 Or StorageData > MAX_STORAGE Then Exit Function
    If invSlot <= 0 Or invSlot > MAX_PLAYER_INV Then Exit Function
    If PlayerInvStorage(Index).slot(storageSlot).Unlocked = False Then Exit Function

    PlayerInv(Index).Data(invSlot).Num = PlayerInvStorage(Index).slot(storageSlot).Data(StorageData).Num
    PlayerInv(Index).Data(invSlot).Value = PlayerInv(Index).Data(invSlot).Value + Value
    PlayerInvStorage(Index).slot(storageSlot).Data(StorageData).Value = PlayerInvStorage(Index).slot(storageSlot).Data(StorageData).Value - Value
    If PlayerInvStorage(Index).slot(storageSlot).Data(StorageData).Value <= 0 Then
        PlayerInvStorage(Index).slot(storageSlot).Data(StorageData).Value = 0
        PlayerInvStorage(Index).slot(storageSlot).Data(StorageData).Num = 0
    End If
    WithdrawItem = True
End Function

Public Function FindSameInvStorageSlot(ByVal Index As Long, ByVal storageSlot As Byte, ByVal ItemNum As Long) As Byte
Dim i As Byte

    FindSameInvStorageSlot = 0
    
    If Not IsPlaying(Index) Then Exit Function
    If TempPlayer(Index).UseChar <= 0 Then Exit Function
    
    If ItemNum <= 0 Then Exit Function
    
    For i = 1 To MAX_STORAGE
        With PlayerInvStorage(Index).slot(storageSlot)
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

Private Function FindFreeInvStorageSlot(ByVal Index As Long, ByVal storageSlot As Byte, ByVal ItemNum As Long) As Byte
Dim i As Byte

    FindFreeInvStorageSlot = 0
    
    If Not IsPlaying(Index) Then Exit Function
    If TempPlayer(Index).UseChar <= 0 Then Exit Function
    If PlayerInvStorage(Index).slot(storageSlot).Unlocked = NO Then Exit Function
    If ItemNum <= 0 Then Exit Function
    
    If Item(ItemNum).Stock = YES Then
        i = FindSameInvStorageSlot(Index, storageSlot, ItemNum)
        If i > 0 Then
            FindFreeInvStorageSlot = i
            Exit Function
        End If
    End If
    
    For i = 1 To MAX_STORAGE
        With PlayerInvStorage(Index).slot(storageSlot).Data(i)
            If .Num = 0 Then
                FindFreeInvStorageSlot = i
                Exit Function
            End If
        End With
    Next
End Function

Public Function checkItem(ByVal Index As Long, ByVal ItemNum As Long) As Long
Dim i As Long

    For i = 1 To MAX_PLAYER_INV
        If PlayerInv(Index).Data(i).Num = ItemNum Then
            checkItem = i
            Exit Function
        End If
    Next
    checkItem = 0
End Function

Public Function GiveStorageItem(ByVal Index As Long, ByVal storageSlot As Byte, ByVal ItemNum As Long, ByVal ItemVal As Long) As Boolean
Dim i As Byte

    '//Get Slot
    i = FindFreeInvStorageSlot(Index, storageSlot, ItemNum)
    
    '//Got slot
    If i > 0 Then
        With PlayerInvStorage(Index).slot(storageSlot).Data(i)
            .Num = ItemNum
            .Value = .Value + ItemVal
        End With
        '//Update
        SendPlayerInvStorageSlot Index, storageSlot, i
        GiveStorageItem = True
    Else
        GiveStorageItem = False
    End If
End Function

Public Sub ProcessConversation(ByVal Index As Long, ByVal Convo As Long, ByVal ConvoData As Byte, Optional ByVal NpcNum As Long = 0, Optional ByVal tReply As Byte = 0)
    Dim i As Long, X As Long
    Dim fixData As Boolean

    fixData = False

startOver:

    If Convo <= 0 Or Convo > MAX_CONVERSATION Then Exit Sub

    If Not fixData Then
        If ConvoData <= 0 Then
            '//Initiate
            TempPlayer(Index).CurConvoData = 1
        Else
            If Conversation(Convo).ConvData(ConvoData).NoReply = YES Then
                TempPlayer(Index).CurConvoData = Conversation(Convo).ConvData(ConvoData).MoveNext
            Else
                If tReply > 0 And tReply <= 3 Then
                    TempPlayer(Index).CurConvoData = Conversation(Convo).ConvData(ConvoData).tReplyMove(tReply)
                Else
                    TempPlayer(Index).CurConvoData = 0  '//End
                End If
            End If
        End If
    End If
    ConvoData = TempPlayer(Index).CurConvoData

    If ConvoData > 0 Then
        With Conversation(Convo).ConvData(ConvoData)
            '//Check for custom script
            Select Case .CustomScript
            Case CONVO_SCRIPT_INVSTORAGE    '//Inv Storage
                If TempPlayer(Index).StorageType = 0 Then
                    TempPlayer(Index).StorageType = 1
                    SendStorage Index
                End If
                fixData = False
            Case CONVO_SCRIPT_POKESTORAGE    '//Pokemon Storage
                If TempPlayer(Index).StorageType = 0 Then
                    TempPlayer(Index).StorageType = 2
                    SendStorage Index
                End If
                fixData = False
            Case CONVO_SCRIPT_HEAL
                '//Heal Pokemon
                For i = 1 To MAX_PLAYER_POKEMON
                    If PlayerPokemons(Index).Data(i).Num > 0 Then
                        If PlayerPokemons(Index).Data(i).CurHp < PlayerPokemons(Index).Data(i).MaxHp Then
                            PlayerPokemons(Index).Data(i).CurHp = PlayerPokemons(Index).Data(i).MaxHp
                        End If
                        If PlayerPokemons(Index).Data(i).Status > 0 Then
                            PlayerPokemons(Index).Data(i).Status = 0
                        End If
                        For X = 1 To MAX_MOVESET
                            If PlayerPokemons(Index).Data(i).Moveset(X).Num > 0 Then
                                If PlayerPokemons(Index).Data(i).Moveset(X).CurPP < PlayerPokemons(Index).Data(i).Moveset(X).TotalPP Then
                                    PlayerPokemons(Index).Data(i).Moveset(X).CurPP = PlayerPokemons(Index).Data(i).Moveset(X).TotalPP
                                    PlayerPokemons(Index).Data(i).Moveset(X).CD = 0
                                End If
                            End If
                        Next
                    End If
                Next
                If Player(Index, TempPlayer(Index).UseChar).CurHp < GetPlayerHP(Player(Index, TempPlayer(Index).UseChar).Level) Then
                    Player(Index, TempPlayer(Index).UseChar).CurHp = GetPlayerHP(Player(Index, TempPlayer(Index).UseChar).Level)
                End If
                If Player(Index, TempPlayer(Index).UseChar).Status > 0 Then
                    Player(Index, TempPlayer(Index).UseChar).Status = 0
                    Player(Index, TempPlayer(Index).UseChar).IsConfuse = False
                End If
                Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "Pokemon HP and PP restored", White
                Case LANG_EN: AddAlert Index, "Pokemon HP and PP restored", White
                Case LANG_ES: AddAlert Index, "Pokemon HP and PP restored", White
                End Select
                SendPlayerPokemons Index
                SendPlayerVital Index
                SendPlayerPokemonStatus Index
                SendPlayerStatus Index
                fixData = False
            Case CONVO_SCRIPT_SHOP
                If .CustomScriptData > 0 Then
                    '//Open Shop
                    If TempPlayer(Index).InShop = 0 Then
                        TempPlayer(Index).InShop = .CustomScriptData
                        SendOpenShop Index
                    End If
                End If
                fixData = False
            Case CONVO_SCRIPT_SETSWITCH
                If .CustomScriptData > 0 Then
                    '//Open Shop
                    If IsPlaying(Index) Then
                        If TempPlayer(Index).UseChar > 0 Then
                            Player(Index, TempPlayer(Index).UseChar).Switches(.CustomScriptData) = .CustomScriptData2
                        End If
                    End If
                End If
                fixData = False
            Case CONVO_SCRIPT_GIVEPOKE
                If .CustomScriptData > 0 Then
                    If IsPlaying(Index) Then
                        If TempPlayer(Index).UseChar > 0 Then
                            GivePlayerPokemon Index, .CustomScriptData, 5, BallEnum.b_Pokeball
                        End If
                    End If
                End If
                fixData = False
            Case CONVO_SCRIPT_GIVEITEM
                If .CustomScriptData > 0 Then
                    If IsPlaying(Index) Then
                        If TempPlayer(Index).UseChar > 0 Then
                            If .CustomScriptData2 > 0 Then
                                TryGivePlayerItem Index, .CustomScriptData, .CustomScriptData2
                            End If
                        End If
                    End If
                End If
                fixData = False
            Case CONVO_SCRIPT_WARPTO
                If .CustomScriptData > 0 Then
                    If IsPlaying(Index) Then
                        If TempPlayer(Index).UseChar > 0 Then
                            PlayerWarp Index, .CustomScriptData, .CustomScriptData2, .CustomScriptData3, Player(Index, TempPlayer(Index).UseChar).Dir
                        End If
                    End If
                End If
                fixData = False
            Case CONVO_SCRIPT_CHECKMONEY
                If .CustomScriptData > 0 Then
                    If IsPlaying(Index) Then
                        If TempPlayer(Index).UseChar > 0 Then
                            If Player(Index, TempPlayer(Index).UseChar).Money >= .CustomScriptData Then
                                '//Next
                                TempPlayer(Index).CurConvoData = .CustomScriptData2
                                fixData = True
                            Else
                                TempPlayer(Index).CurConvoData = .CustomScriptData3
                                fixData = True
                            End If
                        End If
                    End If
                End If
            Case CONVO_SCRIPT_TAKEMONEY
                If .CustomScriptData > 0 Then
                    If IsPlaying(Index) Then
                        If TempPlayer(Index).UseChar > 0 Then
                            Player(Index, TempPlayer(Index).UseChar).Money = Player(Index, TempPlayer(Index).UseChar).Money - .CustomScriptData
                            If Player(Index, TempPlayer(Index).UseChar).Money <= 0 Then
                                Player(Index, TempPlayer(Index).UseChar).Money = 0
                            End If
                            SendPlayerData Index
                        End If
                    End If
                End If
                fixData = False
            Case CONVO_SCRIPT_STARTBATTLE
                If TempPlayer(Index).CurConvoMapNpc > 0 Then

                    '//Npc not rebattle Option (Never Rebattle if Win)
                    If Player(Index, TempPlayer(Index).UseChar).NpcBattledDay(MapNpc(Player(Index, TempPlayer(Index).UseChar).Map, TempPlayer(Index).CurConvoMapNpc).Num).Win = YES Then
                        '//Reseta o atributo caso tenha algum problema
                        If Npc(MapNpc(Player(Index, TempPlayer(Index).UseChar).Map, TempPlayer(Index).CurConvoMapNpc).Num).Rebatle <> REBATLE_NEVER Then
                            Player(Index, TempPlayer(Index).UseChar).NpcBattledDay(MapNpc(Player(Index, TempPlayer(Index).UseChar).Map, TempPlayer(Index).CurConvoMapNpc).Num).Win = NO
                            Player(Index, TempPlayer(Index).UseChar).NpcBattledDay(MapNpc(Player(Index, TempPlayer(Index).UseChar).Map, TempPlayer(Index).CurConvoMapNpc).Num).NpcBattledAt = 0
                            Player(Index, TempPlayer(Index).UseChar).NpcBattledMonth(MapNpc(Player(Index, TempPlayer(Index).UseChar).Map, TempPlayer(Index).CurConvoMapNpc).Num).NpcBattledAt = 0
                            Select Case TempPlayer(Index).CurLanguage
                            Case LANG_PT: AddAlert Index, "Tente novamente, por favor!", White
                            Case LANG_EN: AddAlert Index, "Tente novamente, por favor!", White
                            Case LANG_ES: AddAlert Index, "Tente novamente, por favor!", White
                            End Select
                        Else
                            Select Case TempPlayer(Index).CurLanguage
                            Case LANG_PT: AddAlert Index, "Você não pode lutar novamente com este TREINADOR!", White
                            Case LANG_EN: AddAlert Index, "Você não pode lutar novamente com este TREINADOR!", White
                            Case LANG_ES: AddAlert Index, "Você não pode lutar novamente com este TREINADOR!", White
                            End Select
                        End If
                        '//ToDo: Check if daily/monthly
                    ElseIf Not Player(Index, TempPlayer(Index).UseChar).NpcBattledDay(TempPlayer(Index).CurConvoNpc).NpcBattledAt = Day(Now) Then
                        '// Start Npc Battle
                        If Player(Index, TempPlayer(Index).UseChar).Map > 0 Then
                            If MapNpc(Player(Index, TempPlayer(Index).UseChar).Map, TempPlayer(Index).CurConvoMapNpc).InBattle <= 0 Then
                                MapNpc(Player(Index, TempPlayer(Index).UseChar).Map, TempPlayer(Index).CurConvoMapNpc).InBattle = Index
                                MapNpc(Player(Index, TempPlayer(Index).UseChar).Map, TempPlayer(Index).CurConvoMapNpc).CurPokemon = 1
                                For i = 1 To MAX_PLAYER_POKEMON
                                    If Npc(MapNpc(Player(Index, TempPlayer(Index).UseChar).Map, TempPlayer(Index).CurConvoMapNpc).Num).PokemonNum(i) > 0 Then
                                        MapNpc(Player(Index, TempPlayer(Index).UseChar).Map, TempPlayer(Index).CurConvoMapNpc).PokemonAlive(i) = YES
                                    Else
                                        MapNpc(Player(Index, TempPlayer(Index).UseChar).Map, TempPlayer(Index).CurConvoMapNpc).PokemonAlive(i) = NO
                                    End If
                                Next
                                SpawnNpcPokemon Player(Index, TempPlayer(Index).UseChar).Map, TempPlayer(Index).CurConvoMapNpc, 1
                                TempPlayer(Index).InNpcDuel = TempPlayer(Index).CurConvoMapNpc
                                TempPlayer(Index).DuelTime = 1
                                TempPlayer(Index).DuelTimeTmr = GetTickCount + 1000
                                SendPlayerNpcDuel Index
                            End If
                        End If
                    Else
                        '//Reseta o atributo caso tenha algum problema
                        If Npc(MapNpc(Player(Index, TempPlayer(Index).UseChar).Map, TempPlayer(Index).CurConvoMapNpc).Num).Rebatle = REBATLE_NEVER Then
                            Player(Index, TempPlayer(Index).UseChar).NpcBattledDay(MapNpc(Player(Index, TempPlayer(Index).UseChar).Map, TempPlayer(Index).CurConvoMapNpc).Num).Win = NO
                            Player(Index, TempPlayer(Index).UseChar).NpcBattledDay(MapNpc(Player(Index, TempPlayer(Index).UseChar).Map, TempPlayer(Index).CurConvoMapNpc).Num).NpcBattledAt = 0
                            Player(Index, TempPlayer(Index).UseChar).NpcBattledMonth(MapNpc(Player(Index, TempPlayer(Index).UseChar).Map, TempPlayer(Index).CurConvoMapNpc).Num).NpcBattledAt = 0
                            Select Case TempPlayer(Index).CurLanguage
                            Case LANG_PT: AddAlert Index, "Tente novamente, por favor!", White
                            Case LANG_EN: AddAlert Index, "Tente novamente, por favor!", White
                            Case LANG_ES: AddAlert Index, "Tente novamente, por favor!", White
                            End Select
                        Else
                            Select Case TempPlayer(Index).CurLanguage    'AddAlert index, "You already battled this NPC", White
                            Case LANG_PT: AddAlert Index, "Você já batalhou com esse npc hoje, tente novamente amanhã!", White
                            Case LANG_EN: AddAlert Index, "You have already battled with this npc today, try again tomorrow!", White
                            Case LANG_ES: AddAlert Index, "You have already battled with this npc today, try again tomorrow!", White
                            End Select
                        End If
                    End If
                End If
                fixData = False
            Case CONVO_SCRIPT_RELEARN
                If PlayerPokemon(Index).Num > 0 And PlayerPokemon(Index).slot > 0 Then
                    '//Send Relearn
                    SendRelearnMove Index, PlayerPokemon(Index).Num, PlayerPokemon(Index).slot
                Else
                    AddAlert Index, "Please spawn your pokemon", White
                End If
                fixData = False
            Case CONVO_SCRIPT_GIVEBADGE
                If .CustomScriptData > 0 And .CustomScriptData <= MAX_BADGE Then
                    Player(Index, TempPlayer(Index).UseChar).Badge(.CustomScriptData) = YES
                    SendPlayerData Index
                End If
                fixData = False
            Case CONVO_SCRIPT_CHECKBADGE
                If .CustomScriptData > 0 And .CustomScriptData <= MAX_BADGE Then
                    If IsPlaying(Index) Then
                        If TempPlayer(Index).UseChar > 0 Then
                            If Player(Index, TempPlayer(Index).UseChar).Badge(.CustomScriptData) = YES Then
                                '//Next
                                TempPlayer(Index).CurConvoData = .CustomScriptData2
                                fixData = True
                            Else
                                TempPlayer(Index).CurConvoData = .CustomScriptData3
                                fixData = True
                            End If
                        End If
                    End If
                End If
            Case CONVO_SCRIPT_BEATPOKE
                If .CustomScriptData > 0 And .CustomScriptData <= MAX_GAME_POKEMON Then
                    If MapPokemon(.CustomScriptData).Num <= 0 Then
                        TempPlayer(Index).CurConvoData = .CustomScriptData2
                        fixData = True
                    Else
                        TempPlayer(Index).CurConvoData = .CustomScriptData3
                        fixData = True
                    End If
                End If
            Case CONVO_SCRIPT_CHECKITEM
                If .CustomScriptData > 0 And .CustomScriptData <= MAX_ITEM Then
                    If IsPlaying(Index) Then
                        If TempPlayer(Index).UseChar > 0 Then
                            i = checkItem(Index, .CustomScriptData)
                            If i > 0 Then
                                '//Next
                                If PlayerInv(Index).Data(i).Value >= .CustomScriptData2 Then
                                    TempPlayer(Index).CurConvoData = .CustomScriptData3
                                    fixData = True
                                Else
                                    TempPlayer(Index).CurConvoData = .MoveNext
                                    fixData = True
                                End If
                            Else
                                TempPlayer(Index).CurConvoData = .MoveNext
                                fixData = True
                            End If
                        End If
                    End If
                End If
            Case CONVO_SCRIPT_TAKEITEM
                If .CustomScriptData > 0 And .CustomScriptData <= MAX_ITEM Then
                    If IsPlaying(Index) Then
                        If TempPlayer(Index).UseChar > 0 Then
                            i = checkItem(Index, .CustomScriptData)
                            If i > 0 Then
                                '//Take Item
                                PlayerInv(Index).Data(i).Value = PlayerInv(Index).Data(i).Value - .CustomScriptData2
                                If PlayerInv(Index).Data(i).Value <= 0 Then
                                    '//Clear Item
                                    PlayerInv(Index).Data(i).Num = 0
                                    PlayerInv(Index).Data(i).Value = 0
                                End If
                                SendPlayerInvSlot Index, i
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
                    If IsPlaying(Index) Then
                        If TempPlayer(Index).UseChar > 0 Then
                            If Player(Index, TempPlayer(Index).UseChar).Level >= (.CustomScriptData) Then
                                '//Next
                                TempPlayer(Index).CurConvoData = .CustomScriptData2
                                fixData = True
                            Else
                                TempPlayer(Index).CurConvoData = .CustomScriptData3
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
        TempPlayer(Index).CurConvoNum = 0
        TempPlayer(Index).CurConvoData = 0
        TempPlayer(Index).CurConvoNpc = 0
        TempPlayer(Index).CurConvoMapNpc = 0
    End If

    SendInitConvo Index, TempPlayer(Index).CurConvoNum, TempPlayer(Index).CurConvoData, NpcNum
End Sub

Public Function FindFreePokeStorageSlot(ByVal Index As Long, ByVal storageSlot As Byte) As Byte
Dim i As Byte

    FindFreePokeStorageSlot = 0
    
    If Not IsPlaying(Index) Then Exit Function
    If TempPlayer(Index).UseChar <= 0 Then Exit Function
    If PlayerPokemonStorage(Index).slot(storageSlot).Unlocked = NO Then Exit Function
    
    For i = 1 To MAX_STORAGE
        With PlayerPokemonStorage(Index).slot(storageSlot).Data(i)
            If .Num = 0 Then
                FindFreePokeStorageSlot = i
                Exit Function
            End If
        End With
    Next
End Function

'//Catch
Public Function CatchMapPokemonData(ByVal Index As Long, ByVal MapPokeNum As Long, ByVal UsedBall As Byte) As Boolean
    Dim storageSlot As Byte
    Dim gotSlot As Byte
    Dim i As Long

    CatchMapPokemonData = False
    If MapPokeNum <= 0 Or MapPokeNum > MAX_GAME_POKEMON Then Exit Function
    If MapPokemon(MapPokeNum).Num <= 0 Then Exit Function

    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " has caught " & Trim$(Pokemon(MapPokemon(MapPokeNum).Num).Name)

    gotSlot = FindOpenPokeSlot(Index)
    '//Local Slot
    If gotSlot > 0 Then
        With PlayerPokemons(Index).Data(gotSlot)
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
            AddPlayerPokedex Index, .Num, YES, YES

            '//GlobalMsg IsShiny & Rarity
            If .IsShiny = YES Then
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If TempPlayer(Index).UseChar > 0 Then
                            Select Case TempPlayer(Index).CurLanguage
                            Case LANG_PT: SendPlayerMsg i, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " capturou um " & Trim$(Pokemon(.Num).Name) & " shiny em " & Trim$(Map(GetPlayerMap(Index)).Name), Yellow
                            Case LANG_EN: SendPlayerMsg i, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " capturou um " & Trim$(Pokemon(.Num).Name) & " shiny em " & Trim$(Map(GetPlayerMap(Index)).Name), Yellow
                            Case LANG_ES: SendPlayerMsg i, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " capturou um " & Trim$(Pokemon(.Num).Name) & " shiny em " & Trim$(Map(GetPlayerMap(Index)).Name), Yellow
                            End Select
                        End If
                    End If
                Next i
            ElseIf Spawn(MapPokeNum).Rarity >= Options.Rarity Then
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If TempPlayer(Index).UseChar > 0 Then
                            Select Case TempPlayer(Index).CurLanguage
                            Case LANG_PT: SendPlayerMsg i, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " capturou um " & Trim$(Pokemon(.Num).Name) & " raro em " & Trim$(Map(GetPlayerMap(Index)).Name), Yellow
                            Case LANG_EN: SendPlayerMsg i, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " capturou um " & Trim$(Pokemon(.Num).Name) & " raro em " & Trim$(Map(GetPlayerMap(Index)).Name), Yellow
                            Case LANG_ES: SendPlayerMsg i, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " capturou um " & Trim$(Pokemon(.Num).Name) & " raro em " & Trim$(Map(GetPlayerMap(Index)).Name), Yellow
                            End Select
                        End If
                    End If
                Next i
            End If
        End With

        UpdatePlayerPokemonOrder Index
        SendPlayerPokemons Index

        CatchMapPokemonData = True
        Exit Function
    Else
        '//Check Storage Slot
        For storageSlot = 1 To MAX_STORAGE_SLOT
            gotSlot = FindFreePokeStorageSlot(Index, storageSlot)
            If gotSlot > 0 Then
                '//Give Pokemon
                With PlayerPokemonStorage(Index).slot(storageSlot).Data(gotSlot)
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
                    AddPlayerPokedex Index, .Num, YES, YES
                End With

                Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "Your pokemon has been transferred to your pokemon storage", White
                Case LANG_EN: AddAlert Index, "Your pokemon has been transferred to your pokemon storage", White
                Case LANG_ES: AddAlert Index, "Your pokemon has been transferred to your pokemon storage", White
                End Select

                SendPlayerPokemonStorageSlot Index, storageSlot, gotSlot

                CatchMapPokemonData = True
                Exit Function
            End If
        Next
    End If

    CatchMapPokemonData = False
End Function

Public Function FindOpenTradeSlot(ByVal Index As Long) As Long
Dim i As Byte

    For i = 1 To MAX_TRADE
        If TempPlayer(Index).TradeItem(i).Type = 0 Then
            FindOpenTradeSlot = i
            Exit Function
        End If
    Next
End Function

Public Sub AddPlayerPokedex(ByVal Index As Long, ByVal PokeNum As Long, Optional ByVal Obtained As Byte = 0, Optional ByVal Scanned As Byte = 0)
    If PlayerPokedex(Index).PokemonIndex(PokeNum).Obtained = 0 Then
        PlayerPokedex(Index).PokemonIndex(PokeNum).Obtained = Obtained
        If Obtained = YES Then
            Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, Trim$(Pokemon(PokeNum).Name) & " has been added on pokedex", White
                Case LANG_EN: AddAlert Index, Trim$(Pokemon(PokeNum).Name) & " has been added on pokedex", White
                Case LANG_ES: AddAlert Index, Trim$(Pokemon(PokeNum).Name) & " has been added on pokedex", White
            End Select
        End If
    End If
    If PlayerPokedex(Index).PokemonIndex(PokeNum).Scanned = 0 Then
        PlayerPokedex(Index).PokemonIndex(PokeNum).Scanned = Scanned
        If Scanned = YES Then
            Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, Trim$(Pokemon(PokeNum).Name) & " has been scanned", White
                Case LANG_EN: AddAlert Index, Trim$(Pokemon(PokeNum).Name) & " has been scanned", White
                Case LANG_ES: AddAlert Index, Trim$(Pokemon(PokeNum).Name) & " has been scanned", White
            End Select
        End If
    End If
    SendPlayerPokedexSlot Index, PokeNum
End Sub

Public Sub ClearMyTarget(ByVal Index As Long, ByVal MapNum As Long)
Dim i As Long

    For i = 1 To Pokemon_HighIndex
        If MapPokemon(i).Num > 0 Then
            If MapPokemon(i).Map = MapNum Then
                If MapPokemon(i).targetType = TARGET_TYPE_PLAYER Then
                    If MapPokemon(i).TargetIndex = Index Then
                        MapPokemon(i).targetType = 0
                        MapPokemon(i).TargetIndex = 0
                    End If
                End If
            End If
        End If
    Next
End Sub

Public Sub ChangeTempSprite(ByVal Index As Long, ByVal TempSprite As Byte, Optional ByVal Forced As Boolean = False)
    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub

    Select Case TempSprite
    Case TEMP_SPRITE_GROUP_NONE
        If Forced Then
            Player(Index, TempPlayer(Index).UseChar).TempSprite = 0
        Else
            If Not Player(Index, TempPlayer(Index).UseChar).TempSprite = TEMP_SPRITE_GROUP_BIKE Then
                If Not Player(Index, TempPlayer(Index).UseChar).TempSprite = TEMP_SPRITE_GROUP_MOUNT Then
                    Player(Index, TempPlayer(Index).UseChar).TempSprite = 0
                End If
            End If
        End If
    Case TEMP_SPRITE_GROUP_DIVE
        Player(Index, TempPlayer(Index).UseChar).TempSprite = TEMP_SPRITE_GROUP_DIVE
    Case TEMP_SPRITE_GROUP_BIKE
        If Not Player(Index, TempPlayer(Index).UseChar).TempSprite = TEMP_SPRITE_GROUP_DIVE Then
            If Not Player(Index, TempPlayer(Index).UseChar).TempSprite = TEMP_SPRITE_GROUP_BIKE Then
                Player(Index, TempPlayer(Index).UseChar).TempSprite = TEMP_SPRITE_GROUP_BIKE
            Else
                Player(Index, TempPlayer(Index).UseChar).TempSprite = 0
            End If
        End If
    Case TEMP_SPRITE_GROUP_MOUNT
        If Not Player(Index, TempPlayer(Index).UseChar).TempSprite = TEMP_SPRITE_GROUP_DIVE Then
            If Not Player(Index, TempPlayer(Index).UseChar).TempSprite = TEMP_SPRITE_GROUP_BIKE Then
                If Not Player(Index, TempPlayer(Index).UseChar).TempSprite = TEMP_SPRITE_GROUP_MOUNT Then
                    Player(Index, TempPlayer(Index).UseChar).TempSprite = TEMP_SPRITE_GROUP_MOUNT
                Else
                    Player(Index, TempPlayer(Index).UseChar).TempSprite = 0
                End If
            End If
        End If
        'Case TEMP_SPRITE_GROUP_SURF

    Case Else
        Player(Index, TempPlayer(Index).UseChar).TempSprite = 0
    End Select

    SendPlayerData Index
End Sub

Public Function FindInvItemSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
Dim i As Long

    For i = 1 To MAX_PLAYER_INV
        If PlayerInv(Index).Data(i).Num = ItemNum Then
            FindInvItemSlot = i
            Exit Function
        End If
    Next
End Function

Public Sub SendPlayerPokemonFaint(ByVal Index As Long)
Dim MapNum As Long
Dim DuelIndex As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If PlayerPokemon(Index).Num <= 0 Then Exit Sub
    
    MapNum = Player(Index, TempPlayer(Index).UseChar).Map
    
    ClearPlayerPokemon Index
    If TempPlayer(Index).InDuel > 0 Then
        If IsPlaying(TempPlayer(Index).InDuel) Then
            If TempPlayer(TempPlayer(Index).InDuel).UseChar > 0 Then
                If CountPlayerPokemonAlive(Index) <= 0 Then
                    DuelIndex = TempPlayer(Index).InDuel
                    '//Player Lose
                    SendActionMsg MapNum, "Lose!", Player(Index, TempPlayer(Index).UseChar).X * 32, Player(Index, TempPlayer(Index).UseChar).Y * 32, White
                    SendActionMsg MapNum, "Win!", Player(DuelIndex, TempPlayer(DuelIndex).UseChar).X * 32, Player(DuelIndex, TempPlayer(DuelIndex).UseChar).Y * 32, White
                    Player(Index, TempPlayer(Index).UseChar).Lose = Player(Index, TempPlayer(Index).UseChar).Lose + 1
                    Player(DuelIndex, TempPlayer(DuelIndex).UseChar).Win = Player(DuelIndex, TempPlayer(DuelIndex).UseChar).Win + 1
                    SendPlayerPvP (DuelIndex)
                    SendPlayerPvP (Index)
                    TempPlayer(Index).InDuel = 0
                    TempPlayer(Index).DuelTime = 0
                    TempPlayer(Index).DuelTimeTmr = 0
                    TempPlayer(Index).WarningTimer = 0
                    TempPlayer(Index).PlayerRequest = 0
                    TempPlayer(Index).RequestType = 0
                    TempPlayer(DuelIndex).InDuel = 0
                    TempPlayer(DuelIndex).DuelTime = 0
                    TempPlayer(DuelIndex).DuelTimeTmr = 0
                    TempPlayer(DuelIndex).WarningTimer = 0
                    TempPlayer(DuelIndex).PlayerRequest = 0
                    TempPlayer(DuelIndex).RequestType = 0
                    SendRequest DuelIndex
                    SendRequest Index
                Else
                    TempPlayer(Index).DuelReset = YES
                End If
            End If
        End If
    End If
    If TempPlayer(Index).InNpcDuel > 0 Then
        If CountPlayerPokemonAlive(Index) <= 0 Then
            '//Adicionado a apenas um método.
            PlayerLoseToNpc Index, TempPlayer(Index).InNpcDuel
        Else
            TempPlayer(Index).DuelReset = YES
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

Public Sub GivePlayerExp(ByVal Index As Long, ByVal Exp As Long)
Dim ExpRollover As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    
    ' Exp Rate
    Exp = Exp * Options.ExpRate
    
    With Player(Index, TempPlayer(Index).UseChar)
        If .Level >= MAX_PLAYER_LEVEL Then Exit Sub
        
        'Exp Rate
    If EventExp.ExpEvent Then
        Exp = Exp * EventExp.ExpMultiply
    End If
    
        .CurExp = .CurExp + Exp
        
        If .CurExp > GetLevelNextExp(.Level) Then
            Do While .CurExp > GetLevelNextExp(.Level)
                ExpRollover = .CurExp - GetLevelNextExp(.Level)
                .CurExp = ExpRollover
                .Level = .Level + 1
                .CurHp = GetPlayerHP(.Level)
            Loop
            SendPlayerData Index
            
            '//ActionMsg
            SendActionMsg Player(Index, TempPlayer(Index).UseChar).Map, "Level Up!", .X * 32, .Y * 32, Yellow
        End If
        SendPlayerExp Index
        
        '//ActionMsg
        SendActionMsg Player(Index, TempPlayer(Index).UseChar).Map, "+" & Exp, .X * 32, .Y * 32, White
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

Public Sub KillPlayer(ByVal Index As Long)
Dim ExpPenalty As Long, MoneyPenalty As Long
Dim ExpRollover As Long, CountBadge As Byte
Dim i As Byte

    With Player(Index, TempPlayer(Index).UseChar)
        .CurHp = GetPlayerHP(.Level)
        If .CheckMap <= 0 Then .CheckMap = Options.StartMap
        If .CheckX <= 0 Then .CheckX = Options.startX
        If .CheckY <= 0 Then .CheckY = Options.startY
        PlayerWarp Index, .CheckMap, .CheckX, .CheckY, .CheckDir
        
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
        AddAlert Index, "You lose $" & MoneyPenalty, White
        
        SendPlayerData Index
    End With
End Sub

Public Sub SendWhosOnline(ByVal Index As Long)
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
    SendPlayerMsg Index, s, White
End Sub

Public Sub CreateParty(ByVal Index As Long)
Dim i As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If TempPlayer(Index).InParty > 0 Then Exit Sub
    
    TempPlayer(Index).InParty = YES
    For i = 1 To MAX_PARTY
        TempPlayer(Index).PartyIndex(i) = 0
    Next
    TempPlayer(Index).PartyIndex(1) = Index
    AddAlert Index, "Party Created", White
    SendParty Index
End Sub

Public Sub LeaveParty(ByVal Index As Long)
Dim i As Long, PartyRequest As Long, PartySlot As Byte

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If TempPlayer(Index).InParty <= 0 Then Exit Sub
    
    TempPlayer(Index).InParty = 0
    For i = 1 To MAX_PARTY
        PartyRequest = TempPlayer(Index).PartyIndex(i)
        '//Remove self
        If PartyRequest = Index Then
            PartySlot = i
            TempPlayer(Index).PartyIndex(i) = 0
        End If
    Next
    '//Update to member
    For i = 1 To MAX_PARTY
        PartyRequest = TempPlayer(Index).PartyIndex(i)
        If PartyRequest > 0 Then
            If IsPlaying(PartyRequest) Then
                If TempPlayer(PartyRequest).UseChar > 0 Then
                    If Not PartyRequest = Index Then
                        TempPlayer(PartyRequest).PartyIndex(PartySlot) = 0
                        AddAlert PartyRequest, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " has left the party", White
                        SendParty PartyRequest
                    End If
                End If
            End If
        End If
    Next
    AddAlert Index, "You left the party", White
    SendParty Index
End Sub

Public Sub JoinParty(ByVal Index As Long, ByVal InviteIndex As Long)
Dim i As Long, slot As Byte
Dim PartyRequest As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If TempPlayer(Index).InParty <= 0 Then Exit Sub
    If Not IsPlaying(InviteIndex) Then Exit Sub
    If TempPlayer(InviteIndex).UseChar <= 0 Then Exit Sub
    If TempPlayer(InviteIndex).InParty > 0 Then Exit Sub
    slot = 0
    '//Check free slot
    For i = 1 To MAX_PARTY
        If TempPlayer(Index).PartyIndex(i) <= 0 Then
            slot = i
            Exit For
        End If
    Next
    
    If slot > 0 Then
        For i = 1 To MAX_PARTY
            PartyRequest = TempPlayer(Index).PartyIndex(i)
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
            TempPlayer(InviteIndex).PartyIndex(i) = TempPlayer(Index).PartyIndex(i)
        Next
        TempPlayer(InviteIndex).InParty = YES
        SendParty InviteIndex
    End If
End Sub

Public Function PartyCount(ByVal Index As Long) As Byte
Dim i As Long, count As Long

    count = 0
    For i = 1 To MAX_PARTY
        If TempPlayer(Index).PartyIndex(i) > 0 Then
            count = count + 1
        End If
    Next
    PartyCount = count
End Function

Public Function IsPartyMember(ByVal Index As Long, ByVal i As Long) As Boolean
    Dim z As Byte
    If TempPlayer(Index).InParty > 0 Then
        For z = 1 To MAX_PARTY
            If TempPlayer(Index).PartyIndex(z) > 0 Then
                If TempPlayer(Index).PartyIndex(z) <> Index Then
                    If TempPlayer(Index).PartyIndex(z) = i Then
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
Function GetPlayerMap(ByVal Index As Long) As Long

    If Index > MAX_PLAYER Then Exit Function
    GetPlayerMap = Player(Index, TempPlayer(Index).UseChar).Map
    
End Function

' Obtem o X do jogador
Function GetPlayerX(ByVal Index As Long) As Long

    If Index > MAX_PLAYER Then Exit Function
    GetPlayerX = Player(Index, TempPlayer(Index).UseChar).X
End Function

' Obtem o Y do jogador
Function GetPlayerY(ByVal Index As Long) As Long

    If Index > MAX_PLAYER Then Exit Function
    GetPlayerY = Player(Index, TempPlayer(Index).UseChar).Y
End Function

Function GetPlayerDir(ByVal Index As Long) As Long
    If Index > MAX_PLAYER Then Exit Function
    GetPlayerDir = Player(Index, TempPlayer(Index).UseChar).Dir
End Function

Function GetPlayerLogin(ByVal Index As Long) As String
    GetPlayerLogin = Trim$(Player(Index, TempPlayer(Index).UseChar).Name)
End Function
