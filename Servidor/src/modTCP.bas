Attribute VB_Name = "modTCP"
Option Explicit

' ******************************************
' ** Communcation to server, TCP          **
' ** Winsock Control (mswinsck.ocx)       **
' ** String packets (slow and big)        **
' ******************************************
Public Sub TcpInit()
Dim i As Byte

    '//Set the connection settings
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = Options.Port
    
    '//load all sockets
    For i = 1 To MAX_PLAYER
        ClearTempPlayer i
        ClearPlayer i
        ClearPlayerInv i
        ClearPlayerInvStorage i
        ClearPlayerPokemons i
        ClearPlayerPokemonStorage i
        ClearAccount i
        ClearPlayerPokedex i
        Load frmServer.Socket(i)
    Next
    
    '//Initiate Socket
    frmServer.Socket(0).Listen
End Sub

Public Sub DestroyTCP()
Dim i As Byte

    On Error Resume Next

    '//Unload all sockets
    For i = 1 To MAX_PLAYER
        Unload frmServer.Socket(i)
    Next
End Sub

Public Function IsConnected(ByVal index As Long) As Boolean
    ' Check for subscript out of range
    If index <= 0 Or index > MAX_PLAYER Then Exit Function
    
    If frmServer.Socket(index).State = sckConnected Then
        IsConnected = True
    End If
End Function

Public Function IsPlaying(ByVal index As Long) As Boolean
    ' Check for subscript out of range
    If index <= 0 Or index > MAX_PLAYER Then Exit Function
    
    If TempPlayer(index).InGame Then
        IsPlaying = True
    End If
End Function

Public Function FindOpenPlayerSlot() As Long
Dim i As Byte
    
    For i = 1 To MAX_PLAYER
        If Not IsConnected(i) And Not IsPlaying(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If
    Next
End Function

Public Sub AcceptConnection(ByVal index As Long, ByVal SocketId As Long)
Dim i As Long
Dim count As Long

    ' Prevent spamming
    For i = 1 To MAX_PLAYER
        If GetPlayerIP(i) = Trim$(frmServer.Socket(index).RemoteHostIP) Then
            count = count + 1
            If count > 3 Then Exit Sub
        End If
    Next
    
    ' Make sure to reject connection from banned ip
    If IsIPBanned(Trim$(frmServer.Socket(index).RemoteHostIP)) Then Exit Sub

    If (index = 0) Then
        i = FindOpenPlayerSlot

        If Not i = 0 Then
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If
End Sub

Private Sub SocketConnected(ByVal index As Long)
Dim i As Long, x As Long

    If Not index = 0 Then
        ' make sure they're not banned from ip
        TextAdd frmServer.txtLog, "Receiving connection from " & GetPlayerIP(index) & "..."
        AddIPLog "Receiving connection from " & GetPlayerIP(index)

        '//re-set the high index
        Player_HighIndex = 0
        For i = MAX_PLAYER To 1 Step -1
            If IsConnected(i) Then
                Player_HighIndex = i
                Exit For
            End If
        Next
        '//send the new highindex to all logged in players
        SendHighIndex
    End If
End Sub

Public Sub CloseSocket(ByVal index As Long)
Dim sIP As String

    If index > 0 Then
        sIP = GetPlayerIP(index)
        '//Socket Close
        frmServer.Socket(index).Close

        AddIPLog "Connection from " & sIP & " was terminated"
        TextAdd frmServer.txtLog, "Connection from " & sIP & " has been terminated..."
        LeftGame index
    End If
End Sub

Public Sub SendDataTo(ByVal index As Long, ByRef Data() As Byte)
Dim buffer As clsBuffer
Dim TempData() As Byte

    If IsConnected(index) Then
        Set buffer = New clsBuffer
        TempData = Data
        
        buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        buffer.WriteBytes TempData()
              
        frmServer.Socket(index).SendData buffer.ToArray()
    End If
End Sub

Public Sub SendDataToAll(ByRef Data() As Byte)
Dim i As Byte

    For i = 1 To Player_HighIndex
        SendDataTo i, Data()
    Next
End Sub

Public Sub SendDataToAllBut(ByVal exIndex As Long, ByRef Data() As Byte)
Dim i As Byte

    For i = 1 To Player_HighIndex
        If Not i = exIndex Then
            SendDataTo i, Data()
        End If
    Next
End Sub

Public Sub SendDataToMap(ByVal MapNum As Long, ByRef Data() As Byte)
Dim i As Byte

    For i = 1 To Player_HighIndex
        If TempPlayer(i).UseChar > 0 Then
            If Player(i, TempPlayer(i).UseChar).Map = MapNum Then
                SendDataTo i, Data()
            End If
        End If
    Next
End Sub

Public Sub SendDataToMapBut(ByVal MapNum As Long, ByVal exIndex As Long, ByRef Data() As Byte)
Dim i As Byte

    For i = 1 To Player_HighIndex
        If Not i = exIndex Then
            If TempPlayer(i).UseChar > 0 Then
                If Player(i, TempPlayer(i).UseChar).Map = MapNum Then
                    SendDataTo i, Data()
                End If
            End If
        End If
    Next
End Sub

Public Sub SendHighIndex(Optional ByVal ToSelf As Long = 0)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SHighIndex
    buffer.WriteLong Player_HighIndex
    If ToSelf > 0 Then
        SendDataTo ToSelf, buffer.ToArray()
    Else
        SendDataToAll buffer.ToArray()
    End If
    Set buffer = Nothing
End Sub

Public Sub AddAlert(ByVal index As Long, ByVal Msg As String, ByVal Color As Long, Optional ByVal pDisconnect As Byte = 0, Optional ByVal NotHideLoad As Byte = 0)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SAlertMsg
    buffer.WriteString Msg
    buffer.WriteLong Color
    buffer.WriteByte pDisconnect
    buffer.WriteByte NotHideLoad
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Debug.Print Now & Msg
End Sub

Public Sub SendLoginOk(ByVal index As Long, Optional ByVal Data1 As Byte = 0)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SLoginOk
    buffer.WriteLong index
    buffer.WriteByte Data1
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendCharacters(ByVal index As Long)
Dim buffer As clsBuffer
Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SCharacters
    For i = 1 To MAX_PLAYERCHAR
        buffer.WriteString Player(index, i).Name
        buffer.WriteLong Player(index, i).Sprite
    Next
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendInGame(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SInGame
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Function PlayerData(ByVal index As Long) As Byte()
Dim buffer As clsBuffer
Dim i As Long

    '//Check if player exist
    If TempPlayer(index).UseChar <= 0 Then Exit Function
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerData
    buffer.WriteLong index
    
    With Player(index, TempPlayer(index).UseChar)
        buffer.WriteString Trim$(.Name)
        buffer.WriteLong .Sprite
        buffer.WriteByte .Access
        buffer.WriteLong .Map
        buffer.WriteLong .x
        buffer.WriteLong .Y
        buffer.WriteByte .Dir
        buffer.WriteLong .CurHp
        buffer.WriteLong .Money
        buffer.WriteLong .TempSprite
        For i = 1 To MAX_BADGE
            buffer.WriteByte .Badge(i)
        Next
        buffer.WriteLong .Level
        buffer.WriteLong .CurExp
        For i = 1 To MAX_HOTBAR
            buffer.WriteLong .Hotbar(i)
        Next
        buffer.WriteByte .StealthMode
        buffer.WriteLong .Win
        buffer.WriteLong .Lose
        buffer.WriteLong .Tie
        
        buffer.WriteLong .Cash
        
        buffer.WriteString CStr(.Started)
        buffer.WriteLong .TimePlay
        
        buffer.WriteByte .FishMode
        buffer.WriteByte .FishRod
    End With
    
    PlayerData = buffer.ToArray()
    Set buffer = Nothing
End Function

Public Sub SendPlayerData(ByVal index As Long)
    If Not IsPlaying(index) Then Exit Sub
    
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    
    SendDataToMap Player(index, TempPlayer(index).UseChar).Map, PlayerData(index)
End Sub

Public Sub SendJoinMap(ByVal index As Long)
Dim buffer As clsBuffer
Dim i As Long

    If index <= 0 Or index > MAX_PLAYER Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Or TempPlayer(index).UseChar > MAX_PLAYERCHAR Then Exit Sub

    Set buffer = New clsBuffer
    '//Get Player data from map
    For i = 1 To Player_HighIndex
        If Not i = index Then
            If IsPlaying(i) Then
                If TempPlayer(i).UseChar > 0 Then
                    If Player(i, TempPlayer(i).UseChar).Map = Player(index, TempPlayer(index).UseChar).Map Then
                        SendDataTo index, PlayerData(i)
                        '//Check pokemon
                        If PlayerPokemon(i).Num > 0 Then
                            SendPlayerPokemonData i, 0, index
                        End If
                    Else
                        SendClearPlayer i, index
                    End If
                End If
            End If
        End If
    Next
    '//Send player data to map
    SendDataToMap Player(index, TempPlayer(index).UseChar).Map, PlayerData(index)
    
    '//Get all pokemon data on map
    For i = 1 To Pokemon_HighIndex
        If MapPokemon(i).Num > 0 Then
            If MapPokemon(i).Map = Player(index, TempPlayer(index).UseChar).Map Then
                SendPokemonData i, index
            End If
        End If
    Next
    Set buffer = Nothing
End Sub

Public Sub Create_MapCache(ByVal MapNum As Long)
Dim buffer As clsBuffer
Dim x As Long, Y As Long
Dim i As Long, a As Byte

    Set buffer = New clsBuffer
    buffer.WriteLong SMap
    buffer.WriteLong MapNum
    
    With Map(MapNum)
        '//General
        buffer.WriteLong .Revision
        buffer.WriteString Trim$(.Name)
        buffer.WriteByte .Moral
        
        '//Size
        buffer.WriteLong .MaxX
        buffer.WriteLong .MaxY
    End With
    
    '//Tiles
    For x = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY
            With Map(MapNum).Tile(x, Y)
                '//Layer
                For i = MapLayer.Ground To MapLayer.MapLayer_Count - 1
                    For a = MapLayerType.Normal To MapLayerType.Animated
                        buffer.WriteLong .Layer(i, a).Tile
                        buffer.WriteLong .Layer(i, a).TileX
                        buffer.WriteLong .Layer(i, a).TileY
                        '//Map Anim
                        buffer.WriteLong .Layer(i, a).MapAnim
                    Next
                Next
                '//Tile Data
                buffer.WriteByte .Attribute
                buffer.WriteLong .Data1
                buffer.WriteLong .Data2
                buffer.WriteLong .Data3
                buffer.WriteLong .Data4
            End With
        Next
    Next
    
    With Map(MapNum)
        '//Map Link
        buffer.WriteLong .LinkUp
        buffer.WriteLong .LinkDown
        buffer.WriteLong .LinkLeft
        buffer.WriteLong .LinkRight
        
        '//Map Data
        buffer.WriteString Trim$(.Music)
        
        '//Npc
        For i = 1 To MAX_MAP_NPC
            buffer.WriteLong .Npc(i)
        Next
        
        '//Moral
        buffer.WriteByte .KillPlayer
        buffer.WriteByte .IsCave
        buffer.WriteByte .CaveLight
        buffer.WriteByte .SpriteType
        buffer.WriteByte .StartWeather
    End With
    
    '//Input data to cache
    MapCache(MapNum).Data = buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub CacheAllMaps()
Dim i As Long

    For i = 1 To MAX_MAP
        Create_MapCache i
    Next
End Sub

Public Sub SendMap(ByVal index As Long, ByVal MapNum As Long)
    SendDataTo index, MapCache(MapNum).Data()
End Sub

Public Sub SendCheckForMap(ByVal index As Long, ByVal MapNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SCheckForMap
    buffer.WriteLong MapNum
    '//Send Revision to check if version of map are the same
    buffer.WriteLong Map(MapNum).Revision
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendMapDone(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SMapDone
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerPvP(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerPvP
    '//Send Time
    buffer.WriteLong Player(index, TempPlayer(index).UseChar).Win
    buffer.WriteLong Player(index, TempPlayer(index).UseChar).Lose
    buffer.WriteLong Player(index, TempPlayer(index).UseChar).Tie
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerCash(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerCash
    '//Send Time
    buffer.WriteLong Player(index, TempPlayer(index).UseChar).Cash
    buffer.WriteLong Player(index, TempPlayer(index).UseChar).Money
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerMove(ByVal index As Long, Optional ByVal sendToSelf As Boolean = False)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerMove
    buffer.WriteLong index
    With Player(index, TempPlayer(index).UseChar)
        buffer.WriteLong .x
        buffer.WriteLong .Y
        buffer.WriteByte .Dir
        
        If Not sendToSelf Then
            SendDataToMapBut .Map, index, buffer.ToArray()
        Else
            SendDataToMap .Map, buffer.ToArray()
        End If
    End With
    Set buffer = Nothing
End Sub

Public Sub SendPlayerXY(ByVal index As Long, Optional ByVal sendToSelf As Boolean = False)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerXY
    buffer.WriteLong index
    With Player(index, TempPlayer(index).UseChar)
        buffer.WriteLong .x
        buffer.WriteLong .Y
        
        If Not sendToSelf Then
            SendDataToMapBut index, .Map, buffer.ToArray()
        Else
            SendDataToMap .Map, buffer.ToArray()
        End If
    End With
    Set buffer = Nothing
End Sub

Public Sub SendPlayerDir(ByVal index As Long, Optional ByVal sendToSelf As Boolean = False)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerDir
    buffer.WriteLong index
    With Player(index, TempPlayer(index).UseChar)
        buffer.WriteByte .Dir
        
        If Not sendToSelf Then
            SendDataToMapBut index, .Map, buffer.ToArray()
        Else
            SendDataToMap .Map, buffer.ToArray()
        End If
    End With
    Set buffer = Nothing
End Sub

Public Sub SendLeftGame(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SLeftGame
    buffer.WriteLong index
    SendDataToAllBut index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendLeaveMap(ByVal index As Long, ByVal MapNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SLeftGame
    buffer.WriteLong index
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerMsg
    buffer.WriteString Msg
    buffer.WriteLong Color
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendGlobalMsg(ByVal Msg As String, ByVal Color As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerMsg
    buffer.WriteString Msg
    buffer.WriteLong Color
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendMapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerMsg
    buffer.WriteString Msg
    buffer.WriteLong Color
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendMapNpcData(ByVal MapNum As Long, Optional ByVal ToIndex As Long = 0)
Dim buffer As clsBuffer
Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SMapNpcData
    For i = 1 To MAX_MAP_NPC
        With MapNpc(MapNum, i)
            '//General
            buffer.WriteLong .Num
            
            '//Location
            buffer.WriteLong .x
            buffer.WriteLong .Y
            buffer.WriteByte .Dir
        End With
    Next
    If ToIndex > 0 Then
        SendDataTo ToIndex, buffer.ToArray()
    Else
        SendDataToMap MapNum, buffer.ToArray()
    End If
    Set buffer = Nothing
End Sub

Public Sub SendSpawnMapNpc(ByVal MapNum As Long, ByVal MapNpcNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SSpawnMapNpc
    buffer.WriteLong MapNpcNum
    With MapNpc(MapNum, MapNpcNum)
        '//General
        buffer.WriteLong .Num
        
        '//Location
        buffer.WriteLong .x
        buffer.WriteLong .Y
        buffer.WriteByte .Dir
    End With
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SNpcMove
    buffer.WriteLong MapNpcNum
    With MapNpc(MapNum, MapNpcNum)
        buffer.WriteLong .x
        buffer.WriteLong .Y
        buffer.WriteByte .Dir
    End With
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendNpcDir(ByVal MapNum As Long, ByVal MapNpcNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SNpcDir
    buffer.WriteLong MapNpcNum
    With MapNpc(MapNum, MapNpcNum)
        buffer.WriteByte .Dir
    End With
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

'//Pokemon
Public Sub SendPokemonData(ByVal PokemonIndex As Long, Optional ByVal ToIndex As Long = 0, Optional ByVal SetMap As Long = 0)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPokemonData
    buffer.WriteLong PokemonIndex
    With MapPokemon(PokemonIndex)
        '//General
        buffer.WriteLong .Num
        
        '//Location
        buffer.WriteLong .Map
        buffer.WriteLong .x
        buffer.WriteLong .Y
        buffer.WriteByte .Dir
        
        '//Vital
        buffer.WriteLong .CurHp
        buffer.WriteLong .MaxHp
        
        '//Shiny
        buffer.WriteByte .IsShiny
        
        '//Happiness
        buffer.WriteByte .Happiness
        
        '//Gender
        buffer.WriteByte .Gender
        
        '//Status
        buffer.WriteByte .Status
        
        If ToIndex > 0 Then
            SendDataTo ToIndex, buffer.ToArray()
        Else
            If SetMap > 0 Then
                SendDataToMap SetMap, buffer.ToArray()
            Else
                SendDataToMap .Map, buffer.ToArray()
            End If
        End If
    End With
    Set buffer = Nothing
End Sub

Public Sub SendPokemonHighIndex(Optional ByVal ToIndex As Long = 0)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPokemonHighIndex
    buffer.WriteLong Pokemon_HighIndex
    If ToIndex > 0 Then
        SendDataTo ToIndex, buffer.ToArray()
    Else
        SendDataToAll buffer.ToArray()
    End If
    Set buffer = Nothing
End Sub

Public Sub SendPokemonMove(ByVal MapPokeNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPokemonMove
    buffer.WriteLong MapPokeNum
    With MapPokemon(MapPokeNum)
        buffer.WriteLong .x
        buffer.WriteLong .Y
        buffer.WriteByte .Dir
        SendDataToMap .Map, buffer.ToArray()
    End With
    Set buffer = Nothing
End Sub

Public Sub SendPokemonDir(ByVal MapPokeNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPokemonDir
    buffer.WriteLong MapPokeNum
    With MapPokemon(MapPokeNum)
        buffer.WriteByte .Dir
        SendDataToMap .Map, buffer.ToArray()
    End With
    Set buffer = Nothing
End Sub

Public Sub SendPokemonVital(ByVal MapPokeNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPokemonVital
    buffer.WriteLong MapPokeNum
    With MapPokemon(MapPokeNum)
        buffer.WriteLong .CurHp
        buffer.WriteLong .MaxHp
        SendDataToMap .Map, buffer.ToArray()
    End With
    Set buffer = Nothing
End Sub

Public Sub SendChatbubble(ByVal MapNum As Long, ByVal Target As Long, ByVal targetType As Byte, ByVal Msg As String, ByVal Colour As Long, Optional ByVal x As Long = -1, Optional ByVal Y As Long = -1, Optional ByVal ToIndex As Long = 0)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SChatbubble
    buffer.WriteLong Target
    buffer.WriteByte targetType
    buffer.WriteString Msg
    buffer.WriteLong Colour
    buffer.WriteLong x
    buffer.WriteLong Y
    If ToIndex > 0 Then
        SendDataTo ToIndex, buffer.ToArray()
    Else
        SendDataToMap MapNum, buffer.ToArray()
    End If
    Set buffer = Nothing
End Sub

Public Sub SendPlayerPokemonData(ByVal index As Long, ByVal MapNum As Long, Optional ByVal ToIndex As Long = 0, Optional ByVal Init As Byte = 0, Optional ByVal InitState As Byte = 0, Optional ByVal BallX As Long = 0, Optional ByVal BallY As Long = 0, Optional ByVal UsedBall As Byte = 0)
Dim buffer As clsBuffer
Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerPokemonData
    buffer.WriteLong index
    With PlayerPokemon(index)
        buffer.WriteByte Init
        buffer.WriteByte InitState
        buffer.WriteLong .Num
        buffer.WriteLong .x
        buffer.WriteLong .Y
        buffer.WriteByte .Dir
        buffer.WriteByte .slot
        '//Vital
        If .slot > 0 Then
            '//Stat
            For i = 1 To StatEnum.Stat_Count - 1
                buffer.WriteLong PlayerPokemons(index).Data(.slot).Stat(i).Value
                buffer.WriteLong PlayerPokemons(index).Data(.slot).Stat(i).IV
                buffer.WriteLong PlayerPokemons(index).Data(.slot).Stat(i).EV
            Next
            buffer.WriteLong PlayerPokemons(index).Data(.slot).CurHp
            buffer.WriteLong PlayerPokemons(index).Data(.slot).MaxHp
            buffer.WriteByte PlayerPokemons(index).Data(.slot).IsShiny
            buffer.WriteByte PlayerPokemons(index).Data(.slot).Happiness
            buffer.WriteByte PlayerPokemons(index).Data(.slot).Gender
            buffer.WriteByte PlayerPokemons(index).Data(.slot).Status
            buffer.WriteLong PlayerPokemons(index).Data(.slot).HeldItem
        Else
            For i = 1 To StatEnum.Stat_Count - 1
                buffer.WriteLong 0
                buffer.WriteLong 0
                buffer.WriteLong 0
            Next
            buffer.WriteLong 0
            buffer.WriteLong 0
            buffer.WriteByte 0
            buffer.WriteByte 0
            buffer.WriteByte 0
            buffer.WriteByte 0
            buffer.WriteLong 0
        End If
        buffer.WriteByte UsedBall
        buffer.WriteLong BallX
        buffer.WriteLong BallY
        If ToIndex > 0 Then
            SendDataTo ToIndex, buffer.ToArray()
        Else
            SendDataToMap MapNum, buffer.ToArray()
        End If
    End With
    Set buffer = Nothing
End Sub

Public Sub SendPlayerPokemonMove(ByVal index As Long, Optional ByVal sendToSelf As Boolean = False)
Dim buffer As clsBuffer

    If PlayerPokemon(index).Num <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerPokemonMove
    buffer.WriteLong index
    With PlayerPokemon(index)
        buffer.WriteLong .x
        buffer.WriteLong .Y
        buffer.WriteByte .Dir
        
        If Not sendToSelf Then
            SendDataToMapBut Player(index, TempPlayer(index).UseChar).Map, index, buffer.ToArray()
        Else
            SendDataToMap Player(index, TempPlayer(index).UseChar).Map, buffer.ToArray()
        End If
    End With
    Set buffer = Nothing
End Sub

Public Sub SendPlayerPokemonXY(ByVal index As Long, Optional ByVal sendToSelf As Boolean = False)
Dim buffer As clsBuffer

    If PlayerPokemon(index).Num <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerPokemonXY
    buffer.WriteLong index
    With PlayerPokemon(index)
        buffer.WriteLong .x
        buffer.WriteLong .Y
        
        If Not sendToSelf Then
            SendDataToMapBut index, Player(index, TempPlayer(index).UseChar).Map, buffer.ToArray()
        Else
            SendDataToMap Player(index, TempPlayer(index).UseChar).Map, buffer.ToArray()
        End If
    End With
    Set buffer = Nothing
End Sub

Public Sub SendPlayerPokemonDir(ByVal index As Long, Optional ByVal sendToSelf As Boolean = False)
Dim buffer As clsBuffer

    If PlayerPokemon(index).Num <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerPokemonDir
    buffer.WriteLong index
    With PlayerPokemon(index)
        buffer.WriteByte .Dir
        
        If Not sendToSelf Then
            SendDataToMapBut index, Player(index, TempPlayer(index).UseChar).Map, buffer.ToArray()
        Else
            SendDataToMap Player(index, TempPlayer(index).UseChar).Map, buffer.ToArray()
        End If
    End With
    Set buffer = Nothing
End Sub

Public Sub SendPlayerPokemonVital(ByVal index As Long)
Dim buffer As clsBuffer

    If PlayerPokemon(index).Num <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerPokemonVital
    buffer.WriteLong index
    With PlayerPokemon(index)
        buffer.WriteLong .slot
        If .slot > 0 Then
            buffer.WriteLong PlayerPokemons(index).Data(.slot).CurHp
            buffer.WriteLong PlayerPokemons(index).Data(.slot).MaxHp
        Else
            buffer.WriteLong 0
            buffer.WriteLong 0
        End If
        SendDataToMap Player(index, TempPlayer(index).UseChar).Map, buffer.ToArray()
    End With
    Set buffer = Nothing
End Sub

Public Sub SendPlayerPokemonPP(ByVal index As Long, ByVal MoveSlot As Long)
Dim buffer As clsBuffer

    If PlayerPokemon(index).Num <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerPokemonPP
    buffer.WriteByte MoveSlot
    With PlayerPokemon(index)
        buffer.WriteLong .slot
        If .slot > 0 Then
            buffer.WriteLong PlayerPokemons(index).Data(.slot).Moveset(MoveSlot).CurPP
            buffer.WriteLong PlayerPokemons(index).Data(.slot).Moveset(MoveSlot).TotalPP
        Else
            buffer.WriteLong 0
            buffer.WriteLong 0
        End If
        SendDataTo index, buffer.ToArray()
    End With
    Set buffer = Nothing
End Sub

Public Sub SendPlayerInv(ByVal index As Long)
Dim buffer As clsBuffer
Dim i As Byte

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerInv
    With PlayerInv(index)
        For i = 1 To MAX_PLAYER_INV
            buffer.WriteLong .Data(i).Num
            buffer.WriteLong .Data(i).Value
            buffer.WriteByte .Data(i).Locked
        Next
    End With
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerInvSlot(ByVal index As Long, ByVal slot As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerInvSlot
    buffer.WriteByte slot
    With PlayerInv(index)
        buffer.WriteLong .Data(slot).Num
        buffer.WriteLong .Data(slot).Value
        buffer.WriteByte .Data(slot).Locked
    End With
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerPokemons(ByVal index As Long)
Dim buffer As clsBuffer
Dim i As Byte, x As Byte

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerPokemons
    With PlayerPokemons(index)
        For i = 1 To MAX_PLAYER_POKEMON
            buffer.WriteLong .Data(i).Num
            
            buffer.WriteByte .Data(i).Level
        
            For x = 1 To StatEnum.Stat_Count - 1
                buffer.WriteLong .Data(i).Stat(x).Value
                buffer.WriteLong .Data(i).Stat(x).IV
                buffer.WriteLong .Data(i).Stat(x).EV
            Next
            
            '//Vital
            buffer.WriteLong .Data(i).CurHp
            buffer.WriteLong .Data(i).MaxHp
            
            '//Nature
            buffer.WriteByte .Data(i).Nature
            
            '//Shiny
            buffer.WriteByte .Data(i).IsShiny
            
            '//Happiness
            buffer.WriteByte .Data(i).Happiness
            
            '//Gender
            buffer.WriteByte .Data(i).Gender
            
            '//Status
            buffer.WriteByte .Data(i).Status
            
            '//Exp
            buffer.WriteLong .Data(i).CurExp
            If .Data(i).Num > 0 Then
                buffer.WriteLong GetPokemonNextExp(.Data(i).Level, Pokemon(.Data(i).Num).GrowthRate)
            Else
                buffer.WriteLong 0
            End If
            
            '//Moveset
            For x = 1 To MAX_MOVESET
                buffer.WriteLong .Data(i).Moveset(x).Num
                buffer.WriteByte .Data(i).Moveset(x).CurPP
                buffer.WriteByte .Data(i).Moveset(x).TotalPP
            Next
            
            '//Ball Used
            buffer.WriteByte .Data(i).BallUsed
            
            buffer.WriteLong .Data(i).HeldItem
        Next
    End With
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerPokemonSlot(ByVal index As Long, ByVal slot As Byte)
Dim buffer As clsBuffer, x As Byte

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerPokemonSlot
    buffer.WriteByte slot
    With PlayerPokemons(index)
        buffer.WriteLong .Data(slot).Num
        
        buffer.WriteByte .Data(slot).Level
        
        For x = 1 To StatEnum.Stat_Count - 1
            buffer.WriteLong .Data(slot).Stat(x).Value
            buffer.WriteLong .Data(slot).Stat(x).IV
            buffer.WriteLong .Data(slot).Stat(x).EV
        Next
        
        '//Vital
        buffer.WriteLong .Data(slot).CurHp
        buffer.WriteLong .Data(slot).MaxHp
        
        '//Nature
        buffer.WriteByte .Data(slot).Nature
        
        '//Shiny
        buffer.WriteByte .Data(slot).IsShiny
        
        '//Happiness
        buffer.WriteByte .Data(slot).Happiness
        
        '//Gender
        buffer.WriteByte .Data(slot).Gender
        
        '//Status
        buffer.WriteByte .Data(slot).Status
        
        '//Exp
        buffer.WriteLong .Data(slot).CurExp
        If .Data(slot).Num > 0 Then
            buffer.WriteLong GetPokemonNextExp(.Data(slot).Level, Pokemon(.Data(slot).Num).GrowthRate)
        Else
            buffer.WriteLong 0
        End If
        
        '//Moveset
        For x = 1 To MAX_MOVESET
            buffer.WriteLong .Data(slot).Moveset(x).Num
            buffer.WriteByte .Data(slot).Moveset(x).CurPP
            buffer.WriteByte .Data(slot).Moveset(x).TotalPP
        Next
        
        '//Ball Used
        buffer.WriteByte .Data(slot).BallUsed
        
        buffer.WriteLong .Data(slot).HeldItem
    End With
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendActionMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal x As Long, ByVal Y As Long, ByVal Color As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SActionMsg
    buffer.WriteString Msg
    buffer.WriteLong Color
    buffer.WriteLong x
    buffer.WriteLong Y
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendAttack(ByVal index As Long, ByVal MapNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SAttack
    buffer.WriteLong index
    SendDataToMapBut MapNum, index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayAnimation(ByVal MapNum As Long, ByVal Anim As Long, ByVal x As Long, ByVal Y As Long, Optional ByVal OnlyTo As Long = 0)
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayAnimation
    buffer.WriteLong Anim
    buffer.WriteLong x
    buffer.WriteLong Y
    If OnlyTo > 0 Then
        SendDataTo OnlyTo, buffer.ToArray
    Else
        SendDataToMap MapNum, buffer.ToArray()
    End If
    Set buffer = Nothing
End Sub

Public Sub SendNpcAttack(ByVal MapNum As Long, ByVal MapPokemon As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SNpcAttack
    buffer.WriteLong MapPokemon
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendNewMove(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SNewMove
    buffer.WriteByte TempPlayer(index).MoveLearnPokeSlot
    buffer.WriteLong TempPlayer(index).MoveLearnNum
    buffer.WriteByte TempPlayer(index).MoveLearnIndex
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendGetData(ByVal index As Long, ByVal dataType As ItemTypeEnum, ByVal itemSlot As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SGetData
    buffer.WriteByte dataType
    buffer.WriteByte itemSlot
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendMapPokemonCatchState(ByVal MapNum As Long, ByVal PokeSlot As Long, ByVal x As Long, ByVal Y As Long, ByVal catchState As Byte, ByVal Pic As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SMapPokemonCatchState
    buffer.WriteLong PokeSlot
    buffer.WriteLong x
    buffer.WriteLong Y
    buffer.WriteByte catchState
    buffer.WriteByte Pic
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerVital(ByVal index As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerVital
    buffer.WriteLong index
    buffer.WriteLong Player(index, TempPlayer(index).UseChar).CurHp
    SendDataToMap Player(index, TempPlayer(index).UseChar).Map, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerInvStorage(ByVal index As Long)
Dim buffer As clsBuffer
Dim x As Byte, Y As Byte

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerInvStorage
    With PlayerInvStorage(index)
        For x = 1 To MAX_STORAGE_SLOT
            buffer.WriteByte .slot(x).Unlocked
            For Y = 1 To MAX_STORAGE
                buffer.WriteLong .slot(x).Data(Y).Num
                buffer.WriteLong .slot(x).Data(Y).Value
            Next
        Next
    End With
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerInvStorageSlot(ByVal index As Long, ByVal slot As Byte, ByVal Data As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerInvStorageSlot
    buffer.WriteByte slot
    buffer.WriteByte Data
    With PlayerInvStorage(index)
        buffer.WriteLong .slot(slot).Data(Data).Num
        buffer.WriteLong .slot(slot).Data(Data).Value
    End With
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerPokemonStorage(ByVal index As Long)
Dim buffer As clsBuffer
Dim x As Byte, Y As Byte, z As Byte

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerPokemonStorage
    With PlayerPokemonStorage(index)
        For x = 1 To MAX_STORAGE_SLOT
            buffer.WriteByte .slot(x).Unlocked
            For Y = 1 To MAX_STORAGE
                buffer.WriteLong .slot(x).Data(Y).Num
                
                '//Stats
                buffer.WriteByte .slot(x).Data(Y).Level
                For z = 1 To StatEnum.Stat_Count - 1
                    buffer.WriteLong .slot(x).Data(Y).Stat(z).Value
                    buffer.WriteLong .slot(x).Data(Y).Stat(z).IV
                    buffer.WriteLong .slot(x).Data(Y).Stat(z).EV
                Next
                
                '//Vital
                buffer.WriteLong .slot(x).Data(Y).CurHp
                buffer.WriteLong .slot(x).Data(Y).MaxHp
                
                '//Nature
                buffer.WriteByte .slot(x).Data(Y).Nature
                
                '//Shiny
                buffer.WriteByte .slot(x).Data(Y).IsShiny
                
                '//Happiness
                buffer.WriteByte .slot(x).Data(Y).Happiness
                
                '//Gender
                buffer.WriteByte .slot(x).Data(Y).Gender
                
                '//Status
                buffer.WriteByte .slot(x).Data(Y).Status
                
                '//Exp
                buffer.WriteLong .slot(x).Data(Y).CurExp
                If .slot(x).Data(Y).Num > 0 Then
                    buffer.WriteLong GetPokemonNextExp(.slot(x).Data(Y).Level, Pokemon(.slot(x).Data(Y).Num).GrowthRate)
                Else
                    buffer.WriteLong 0
                End If
                
                '//Moveset
                For z = 1 To MAX_MOVESET
                    buffer.WriteLong .slot(x).Data(Y).Moveset(z).Num
                    buffer.WriteLong .slot(x).Data(Y).Moveset(z).CurPP
                    buffer.WriteLong .slot(x).Data(Y).Moveset(z).TotalPP
                Next
                
                '//Ball Used
                buffer.WriteByte .slot(x).Data(Y).BallUsed
                
                '//Held Item
                buffer.WriteLong .slot(x).Data(Y).HeldItem
            Next
        Next
    End With
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerPokemonStorageSlot(ByVal index As Long, ByVal slot As Byte, ByVal Data As Byte)
Dim buffer As clsBuffer
Dim x As Byte

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerPokemonStorageSlot
    buffer.WriteByte slot
    buffer.WriteByte Data
    With PlayerPokemonStorage(index)
        buffer.WriteLong .slot(slot).Data(Data).Num
                
        '//Stats
        buffer.WriteByte .slot(slot).Data(Data).Level
        For x = 1 To StatEnum.Stat_Count - 1
            buffer.WriteLong .slot(slot).Data(Data).Stat(x).Value
            buffer.WriteLong .slot(slot).Data(Data).Stat(x).IV
            buffer.WriteLong .slot(slot).Data(Data).Stat(x).EV
        Next
                
        '//Vital
        buffer.WriteLong .slot(slot).Data(Data).CurHp
        buffer.WriteLong .slot(slot).Data(Data).MaxHp
                
        '//Nature
        buffer.WriteByte .slot(slot).Data(Data).Nature
        
        '//Shiny
        buffer.WriteByte .slot(slot).Data(Data).IsShiny
        
        '//Happiness
        buffer.WriteByte .slot(slot).Data(Data).Happiness
        
        '//Gender
        buffer.WriteByte .slot(slot).Data(Data).Gender
        
        '//Status
        buffer.WriteByte .slot(slot).Data(Data).Status
        
        '//Exp
        buffer.WriteLong .slot(slot).Data(Data).CurExp
        If .slot(slot).Data(Data).Num > 0 Then
            buffer.WriteLong GetPokemonNextExp(.slot(slot).Data(Data).Level, Pokemon(.slot(slot).Data(Data).Num).GrowthRate)
        Else
            buffer.WriteLong 0
        End If
                
        '//Moveset
        For x = 1 To MAX_MOVESET
            buffer.WriteLong .slot(slot).Data(Data).Moveset(x).Num
            buffer.WriteLong .slot(slot).Data(Data).Moveset(x).CurPP
            buffer.WriteLong .slot(slot).Data(Data).Moveset(x).TotalPP
        Next
                
        '//Ball Used
        buffer.WriteByte .slot(slot).Data(Data).BallUsed
        
        '//Held Item
        buffer.WriteLong .slot(slot).Data(Data).HeldItem
    End With
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendInitConvo(ByVal index As Long, ByVal ConvoNum As Long, ByVal ConvoData As Byte, Optional ByVal NpcNum As Long = 0)
Dim buffer As clsBuffer
Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SInitConvo
    buffer.WriteLong ConvoNum
    buffer.WriteByte ConvoData
    buffer.WriteLong NpcNum
    If ConvoNum > 0 And ConvoData > 0 Then
        buffer.WriteString Trim$(Conversation(ConvoNum).ConvData(ConvoData).TextLang(TempPlayer(index).CurLanguage + 1).Text)
        buffer.WriteByte Conversation(ConvoNum).ConvData(ConvoData).NoReply
        For i = 1 To 3
            buffer.WriteString Trim$(Conversation(ConvoNum).ConvData(ConvoData).TextLang(TempPlayer(index).CurLanguage + 1).tReply(i))
        Next
    Else
        buffer.WriteString vbNullString
        buffer.WriteByte 0
        For i = 1 To 3
            buffer.WriteString vbNullString
        Next
    End If
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendStorage(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SStorage
    buffer.WriteByte TempPlayer(index).StorageType
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendOpenShop(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SOpenShop
    buffer.WriteLong TempPlayer(index).InShop
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequest(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SRequest
    buffer.WriteLong TempPlayer(index).PlayerRequest
    buffer.WriteByte TempPlayer(index).RequestType
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlaySound(ByVal SoundName As String, ByVal MapNum As Long, Optional ByVal ToIndex As Long = 0)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlaySound
    buffer.WriteString Trim$(SoundName)
    If ToIndex > 0 Then
        SendDataTo ToIndex, buffer.ToArray()
    Else
        SendDataToMap MapNum, buffer.ToArray()
    End If
    Set buffer = Nothing
End Sub

Public Sub SendOpenTrade(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SOpenTrade
    buffer.WriteLong TempPlayer(index).InTrade
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendUpdateTradeItem(ByVal index As Long, ByVal tradeIndex As Long, ByVal TradeSlot As Byte)
Dim buffer As clsBuffer, x As Byte

    If TradeSlot <= 0 Or TradeSlot > MAX_TRADE Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SUpdateTradeItem
    If tradeIndex = index Then
        buffer.WriteByte 1
    Else
        buffer.WriteByte 0
    End If
    buffer.WriteByte TradeSlot
    With TempPlayer(tradeIndex).TradeItem(TradeSlot)
        buffer.WriteByte .Type
        
        buffer.WriteLong .Num
        buffer.WriteLong .Value
        
        buffer.WriteByte .Level
        
        For x = 1 To StatEnum.Stat_Count - 1
            buffer.WriteLong .Stat(x)
            buffer.WriteLong .StatIV(x)
            buffer.WriteLong .StatEV(x)
        Next
        
        '//Vital
        buffer.WriteLong .CurHp
        buffer.WriteLong .MaxHp
        
        '//Nature
        buffer.WriteByte .Nature
        
        '//Shiny
        buffer.WriteByte .IsShiny
        
        '//Happiness
        buffer.WriteByte .Happiness
        
        '//Gender
        buffer.WriteByte .Gender
        
        '//Status
        buffer.WriteByte .Status
        
        '//Exp
        buffer.WriteLong .CurExp
        buffer.WriteLong .nextExp
        
        '//Moveset
        For x = 1 To MAX_MOVESET
            buffer.WriteLong .Moveset(x).Num
            buffer.WriteByte .Moveset(x).CurPP
            buffer.WriteByte .Moveset(x).TotalPP
        Next
        
        '//Ball Used
        buffer.WriteByte .BallUsed
        
        '//Held Item
        buffer.WriteLong .HeldItem
        
        '//Trade Slot
        buffer.WriteByte .TradeSlot
    End With
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendTradeUpdateMoney(ByVal index As Long, ByVal TargetIndex As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong STradeUpdateMoney
    If index = TargetIndex Then
        buffer.WriteByte 1
    Else
        buffer.WriteByte 0
    End If
    buffer.WriteLong TempPlayer(TargetIndex).TradeMoney
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSetTradeState(ByVal index As Long, ByVal TargetIndex As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SSetTradeState
    If index = TargetIndex Then
        buffer.WriteByte 1
    Else
        buffer.WriteByte 0
    End If
    buffer.WriteByte TempPlayer(TargetIndex).TradeSet
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendCloseTrade(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SCloseTrade
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerPokedex(ByVal index As Long)
Dim buffer As clsBuffer
Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerPokedex
    For i = 1 To MAX_POKEMON
        With PlayerPokedex(index).PokemonIndex(i)
            buffer.WriteByte .Scanned
            buffer.WriteByte .Obtained
        End With
    Next
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerPokedexSlot(ByVal index As Long, ByVal slot As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerPokedexSlot
    buffer.WriteLong slot
    With PlayerPokedex(index).PokemonIndex(slot)
        buffer.WriteByte .Scanned
        buffer.WriteByte .Obtained
    End With
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendMapPokemonStatus(ByVal MapPokeNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPokemonStatus
    buffer.WriteLong MapPokeNum
    With MapPokemon(MapPokeNum)
        buffer.WriteByte .Status
        SendDataToMap .Map, buffer.ToArray()
    End With
    Set buffer = Nothing
End Sub

Public Sub SendMapNpcPokemonStatus(ByVal MapNum As Long, ByVal MapPokeNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SMapNpcPokemonStatus
    buffer.WriteLong MapPokeNum
    With MapNpcPokemon(MapNum, MapPokeNum)
        buffer.WriteByte .Status
    End With
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerPokemonStatus(ByVal index As Long)
Dim buffer As clsBuffer

    If PlayerPokemon(index).Num <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerPokemonStatus
    buffer.WriteLong index
    With PlayerPokemon(index)
        buffer.WriteLong .slot
        If .slot > 0 Then
            buffer.WriteByte PlayerPokemons(index).Data(.slot).Status
        Else
            buffer.WriteByte 0
        End If
        buffer.WriteByte .IsConfuse
        SendDataToMap Player(index, TempPlayer(index).UseChar).Map, buffer.ToArray()
    End With
    Set buffer = Nothing
End Sub

Public Sub SendClearPlayer(ByVal TargetIndex As Long, ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SClearPlayer
    buffer.WriteLong TargetIndex
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerPokemonsStat(ByVal index As Long, ByVal slot As Byte)
Dim buffer As clsBuffer
Dim x As Byte

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerPokemonsStat
    buffer.WriteByte slot
    For x = 1 To StatEnum.Stat_Count - 1
        buffer.WriteLong PlayerPokemons(index).Data(slot).Stat(x).Value
        buffer.WriteLong PlayerPokemons(index).Data(slot).Stat(x).IV
        buffer.WriteLong PlayerPokemons(index).Data(slot).Stat(x).EV
    Next
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerPokemonStatBuff(ByVal index As Long)
Dim buffer As clsBuffer
Dim x As Byte

    If PlayerPokemon(index).Num <= 0 Then Exit Sub
    If PlayerPokemon(index).slot <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerPokemonStatBuff
    For x = 1 To StatEnum.Stat_Count - 1
        buffer.WriteLong PlayerPokemon(index).StatBuff(x)
    Next
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerStatus(ByVal index As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerStatus
    buffer.WriteByte Player(index, TempPlayer(index).UseChar).Status
    buffer.WriteByte Player(index, TempPlayer(index).UseChar).IsConfuse
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendWeather(ByVal MapNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SWeather
    buffer.WriteByte Map(MapNum).CurWeather
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendWeatherTo(ByVal index As Long, ByVal MapNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SWeather
    buffer.WriteByte Map(MapNum).CurWeather
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

'//Pokemon
Public Sub SendNpcPokemonData(ByVal MapNum As Long, ByVal NpcIndex As Long, ByVal Init As Byte, ByVal InitState As Byte, ByVal BallX As Long, ByVal BallY As Long, Optional ByVal ToIndex As Long = 0, Optional ByVal SetMap As Long = 0)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SNpcPokemonData
    buffer.WriteLong NpcIndex
    With MapNpcPokemon(MapNum, NpcIndex)
        buffer.WriteByte Init
        buffer.WriteByte InitState
        buffer.WriteLong BallX
        buffer.WriteLong BallY
        
        '//General
        buffer.WriteLong .Num
        
        '//Location
        buffer.WriteLong .x
        buffer.WriteLong .Y
        buffer.WriteByte .Dir
        
        '//Vital
        buffer.WriteLong .CurHp
        buffer.WriteLong .MaxHp
        
        '//Shiny
        buffer.WriteByte .IsShiny
        
        '//Happiness
        buffer.WriteByte .Happiness
        
        '//Gender
        buffer.WriteByte .Gender
        
        '//Status
        buffer.WriteByte .Status
        
        If ToIndex > 0 Then
            SendDataTo ToIndex, buffer.ToArray()
        Else
            If SetMap > 0 Then
                SendDataToMap SetMap, buffer.ToArray()
            Else
                SendDataToMap MapNum, buffer.ToArray()
            End If
        End If
    End With
    Set buffer = Nothing
End Sub

Public Sub SendNpcPokemonMove(ByVal MapNum As Long, ByVal NpcIndex As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SNpcPokemonMove
    buffer.WriteLong NpcIndex
    With MapNpcPokemon(MapNum, NpcIndex)
        buffer.WriteLong .x
        buffer.WriteLong .Y
        buffer.WriteByte .Dir
        SendDataToMap MapNum, buffer.ToArray()
    End With
    Set buffer = Nothing
End Sub

Public Sub SendNpcPokemonDir(ByVal MapNum As Long, ByVal NpcIndex As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SNpcPokemonDir
    buffer.WriteLong NpcIndex
    With MapNpcPokemon(MapNum, NpcIndex)
        buffer.WriteByte .Dir
        SendDataToMap MapNum, buffer.ToArray()
    End With
    Set buffer = Nothing
End Sub

Public Sub SendNpcPokemonVital(ByVal MapNum As Long, ByVal NpcIndex As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SNpcPokemonVital
    buffer.WriteLong NpcIndex
    With MapNpcPokemon(MapNum, NpcIndex)
        buffer.WriteLong .CurHp
        buffer.WriteLong .MaxHp
        SendDataToMap MapNum, buffer.ToArray()
    End With
    Set buffer = Nothing
End Sub

Public Sub SendPlayerNpcDuel(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerNpcDuel
    buffer.WriteLong TempPlayer(index).InNpcDuel
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRelearnMove(ByVal index As Long, ByVal PokeNum As Long, ByVal PokeSlot As Byte)
Dim buffer As clsBuffer

    If PlayerPokemons(index).Data(PokeSlot).Num <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SRelearnMove
    buffer.WriteLong PokeNum
    buffer.WriteByte PokeSlot
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerAction(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerAction
    buffer.WriteByte Player(index, TempPlayer(index).UseChar).Action
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerExp(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerExp
    buffer.WriteLong Player(index, TempPlayer(index).UseChar).CurExp
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendParty(ByVal index As Long)
Dim buffer As clsBuffer
Dim PartyIndex As Long
Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SParty
    buffer.WriteByte TempPlayer(index).InParty
    For i = 1 To MAX_PARTY
        If TempPlayer(index).InParty <= 0 Then
            buffer.WriteString vbNullString
        Else
            PartyIndex = TempPlayer(index).PartyIndex(i)
            If PartyIndex > 0 Then
                If IsPlaying(PartyIndex) Then
                    If TempPlayer(PartyIndex).UseChar > 0 Then
                        buffer.WriteString Trim$(Player(PartyIndex, TempPlayer(PartyIndex).UseChar).Name)
                    End If
                End If
            End If
        End If
    Next
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ********************
' ***    EDITOR    ***
' ********************
Public Sub SendNpcs(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_NPC
        If LenB(Trim$(Npc(i).Name)) > 0 Then
            SendUpdateNpcTo index, i
        End If
    Next
End Sub

Public Sub SendUpdateNpcTo(ByVal index As Long, ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(Npc(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(Npc(xIndex)), dSize
    buffer.WriteLong SNpcs
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendUpdateNpcToAll(ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(Npc(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(Npc(xIndex)), dSize
    buffer.WriteLong SNpcs
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPokemons(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_POKEMON
        If LenB(Trim$(Pokemon(i).Name)) > 0 Then
             SendUpdatePokemonTo index, i
        End If
    Next
End Sub

Public Sub SendUpdatePokemonTo(ByVal index As Long, ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(Pokemon(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(Pokemon(xIndex)), dSize
    buffer.WriteLong SPokemons
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendUpdatePokemonToAll(ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(Pokemon(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(Pokemon(xIndex)), dSize
    buffer.WriteLong SPokemons
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendItems(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_ITEM
        If LenB(Trim$(Item(i).Name)) > 0 Then
            SendUpdateItemTo index, i
        End If
    Next
End Sub

Public Sub SendUpdateItemTo(ByVal index As Long, ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(Item(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(Item(xIndex)), dSize
    buffer.WriteLong SItems
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendUpdateItemToAll(ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(Item(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(Item(xIndex)), dSize
    buffer.WriteLong SItems
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPokemonMoves(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_POKEMON_MOVE
        If LenB(Trim$(PokemonMove(i).Name)) > 0 Then
            SendUpdatePokemonMoveTo index, i
        End If
    Next
End Sub

Public Sub SendUpdatePokemonMoveTo(ByVal index As Long, ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(PokemonMove(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(PokemonMove(xIndex)), dSize
    buffer.WriteLong SPokemonMoves
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendUpdatePokemonMoveToAll(ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(PokemonMove(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(PokemonMove(xIndex)), dSize
    buffer.WriteLong SPokemonMoves
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendAnimations(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_ANIMATION
        If LenB(Trim$(Animation(i).Name)) > 0 Then
            SendUpdateAnimationTo index, i
        End If
    Next
End Sub

Public Sub SendUpdateAnimationTo(ByVal index As Long, ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(Animation(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(Animation(xIndex)), dSize
    buffer.WriteLong SAnimation
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendUpdateAnimationToAll(ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(Animation(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(Animation(xIndex)), dSize
    buffer.WriteLong SAnimation
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSpawns(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_GAME_POKEMON
        If Spawn(i).PokeNum > 0 Then
            SendUpdateSpawnTo index, i
        End If
    Next
End Sub

Public Sub SendUpdateSpawnTo(ByVal index As Long, ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(Spawn(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(Spawn(xIndex)), dSize
    buffer.WriteLong SSpawn
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendUpdateSpawnToAll(ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(Spawn(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(Spawn(xIndex)), dSize
    buffer.WriteLong SSpawn
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendConversations(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_CONVERSATION
        If LenB(Trim$(Conversation(i).Name)) > 0 Then
            'AddAlert Index, "Loading Events [" & i & "/" & MAX_CONVERSATION & "]...", White, , YES
            SendUpdateConversationTo index, i
        End If
    Next
End Sub

Public Sub SendUpdateConversationTo(ByVal index As Long, ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

'Debug.Print Conversation(xIndex).ConvData(1).TextLang(

    Set buffer = New clsBuffer
    dSize = LenB(Conversation(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(Conversation(xIndex)), dSize
    buffer.WriteLong SConversation
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendUpdateConversationToAll(ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(Conversation(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(Conversation(xIndex)), dSize
    buffer.WriteLong SConversation
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendShops(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_SHOP
        If LenB(Trim$(Shop(i).Name)) > 0 Then
            SendUpdateShopTo index, i
        End If
    Next
End Sub

Public Sub SendUpdateShopTo(ByVal index As Long, ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(Shop(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(Shop(xIndex)), dSize
    buffer.WriteLong SShop
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendUpdateShopToAll(ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(Shop(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(Shop(xIndex)), dSize
    buffer.WriteLong SShop
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendQuests(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_QUEST
        If LenB(Trim$(Quest(i).Name)) > 0 Then
            SendUpdateQuestTo index, i
        End If
    Next
End Sub

Public Sub SendUpdateQuestTo(ByVal index As Long, ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(Quest(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(Quest(xIndex)), dSize
    buffer.WriteLong SQuest
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendUpdateQuestToAll(ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(Quest(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(Quest(xIndex)), dSize
    buffer.WriteLong SQuest
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRankToAll()
Dim buffer As clsBuffer
Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SRank
    For i = 1 To MAX_RANK
        buffer.WriteString Trim$(Rank(i).Name)
        buffer.WriteLong Rank(i).Level
        buffer.WriteLong Rank(i).Exp
    Next
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRankTo(ByVal index As Long)
Dim buffer As clsBuffer
Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SRank
    For i = 1 To MAX_RANK
        buffer.WriteString Trim$(Rank(i).Name)
        buffer.WriteLong Rank(i).Level
        buffer.WriteLong Rank(i).Exp
    Next
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendDataLimit(ByVal index As Long)
Dim buffer As clsBuffer
Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SDataLimit
    buffer.WriteInteger MAX_PLAYER
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestCash(ByVal index As Long, ByVal FindP As Integer, Optional ByVal IsCash As Boolean = True)
    Dim buffer As clsBuffer

    If Player(index, TempPlayer(index).UseChar).Access < ACCESS_CREATOR Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SRequestCash
    If IsCash Then buffer.WriteLong Player(FindP, TempPlayer(FindP).UseChar).Cash Else: buffer.WriteLong Player(FindP, TempPlayer(FindP).UseChar).Money
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendEventInfo(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEventInfo
    buffer.WriteByte EventExp.ExpMultiply
    buffer.WriteLong EventExp.ExpSecs
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestServerInfo(ByVal index As Long)
    Dim buffer As clsBuffer
    Dim sString As String, Colour As Integer

    If frmServer.chkStaffOnly.Value = YES Then
        sString = "Staff"
        Colour = Magenta
    Else
        sString = "Online"
        Colour = BrightGreen
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SRequestServerInfo
    buffer.WriteString sString
    buffer.WriteInteger TotalPlayerOnline
    buffer.WriteInteger Colour
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing

End Sub

Public Sub SendClientTimeTo(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SClientTime
    buffer.WriteByte Weekday(Date)
    buffer.WriteByte GameHour
    buffer.WriteByte GameMinute
    buffer.WriteByte GameSecs
    buffer.WriteByte GameSecs_Velocity
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendClientTimeToAll()
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SClientTime
    buffer.WriteByte Weekday(Date)
    buffer.WriteByte GameHour
    buffer.WriteByte GameMinute
    buffer.WriteByte GameSecs
    buffer.WriteByte GameSecs_Velocity
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendVirtualShopTo(ByVal index As Long)
    Dim buffer As clsBuffer
    Dim i As Long, x As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SSendVirtualShop

    For i = 1 To VirtualShopTabsRec.CountTabs - 1
        buffer.WriteLong VirtualShop(i).Max_Slots
        For x = 1 To VirtualShop(i).Max_Slots
            buffer.WriteLong VirtualShop(i).Items(x).ItemNum
            buffer.WriteLong VirtualShop(i).Items(x).ItemQuant
            buffer.WriteLong VirtualShop(i).Items(x).ItemPrice
            buffer.WriteByte VirtualShop(i).Items(x).CustomDesc
        Next x
    Next i

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendFishMode(ByVal index As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SFishMode
    buffer.WriteLong index
    buffer.WriteByte GetPlayerFishMode(index)
    buffer.WriteByte GetPlayerFishRod(index)
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendItemsCooldown(ByVal index As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Dim CD As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SItemCooldown

    For i = 1 To MAX_PLAYER_INV
        CD = PlayerInv(index).Data(i).TmrCooldown - GetTickCount
        If PlayerInv(index).Data(i).Num > 0 And CD > 0 Then
            buffer.WriteByte i
            buffer.WriteLong CD
        End If
    Next i

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendItemCooldown(ByVal index As Long, ByVal InvNum As Long)
    Dim buffer As clsBuffer
    Dim CD As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SItemCooldown

    CD = PlayerInv(index).Data(InvNum).TmrCooldown - GetTickCount
    If PlayerInv(index).Data(InvNum).Num > 0 And CD > 0 Then
        buffer.WriteByte InvNum
        buffer.WriteLong CD
    End If

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub
