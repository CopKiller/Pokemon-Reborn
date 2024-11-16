Attribute VB_Name = "modDuel"
Option Explicit

Public Const ENTITY_OPACITY_ORIGIN As Byte = 255
Public Const ENTITY_OPACITY_DUEL As Byte = 30

'//Duel
Public InDuelTargetType As Byte
Public InDuelTargetIndex As Long


Public Sub HandlePlayerDuel(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim id As Long
    Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    InDuelTargetType = buffer.ReadByte
    InDuelTargetIndex = buffer.ReadLong
    
    id = buffer.ReadLong
    
    If id > 0 Then
        MapPokemon(id).InDuel = Index
        
        For i = 1 To Pokemon_HighIndex
            id = buffer.ReadLong
            If id > 0 Then
                If MapPokemon(id).Num > 0 Then
                    If MapPokemon(id).Map = Player(Index).Map Then
                        MapPokemon(id).InDuel = True
                    Else
                        MapPokemon(id).InDuel = False
                    End If
                End If
            Else
                Exit For
            End If
        Next i
    End If
    
    Set buffer = Nothing
    
    Call ProcessPlayerDuel
End Sub

Public Sub ProcessPlayerDuel()
    Dim i As Long

    If InDuelTargetType = TARGET_TYPE_NPC Then
    
            For i = 1 To Npc_HighIndex
                If MapNpc(i).Num > 0 Then
                    If i <> InDuelTargetIndex Then
                        MapNpc(i).Opacity = ENTITY_OPACITY_DUEL
                    Else
                        MapNpc(i).Opacity = ENTITY_OPACITY_ORIGIN
                    End If
                End If
            Next i
            
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If i <> MyIndex Then
                        Player(i).Opacity = ENTITY_OPACITY_DUEL
                    Else
                        Player(i).Opacity = ENTITY_OPACITY_ORIGIN
                    End If
                End If
            Next i
            
            For i = 1 To Pokemon_HighIndex
                If MapPokemon(i).Num > 0 Then
                    MapPokemon(i).Opacity = ENTITY_OPACITY_DUEL
                End If
            Next i
        
    ElseIf InDuelTargetType = TARGET_TYPE_PLAYER Then
    
        If IsPlaying(InDuelTargetIndex) Then
            For i = 1 To Player_HighIndex
                If i <> MyIndex And i <> InDuelTargetIndex Then
                    Player(i).Opacity = ENTITY_OPACITY_DUEL
                Else
                    Player(i).Opacity = ENTITY_OPACITY_ORIGIN
                End If
            Next i
            
            For i = 1 To Npc_HighIndex
                If MapNpc(i).Num > 0 Then
                    MapNpc(i).Opacity = ENTITY_OPACITY_DUEL
                End If
            Next i
            
            For i = 1 To Pokemon_HighIndex
                If MapPokemon(i).Num > 0 Then
                    MapPokemon(i).Opacity = ENTITY_OPACITY_DUEL
                End If
            Next i
        End If

    ElseIf InDuelTargetType = TARGET_TYPE_MAPPOKEMON Then
        
            For i = 1 To Pokemon_HighIndex
                If MapPokemon(i).InDuel = True Then
                    MapPokemon(i).Opacity = ENTITY_OPACITY_ORIGIN
                Else
                    MapPokemon(i).Opacity = ENTITY_OPACITY_DUEL
                End If
            Next i

            For i = 1 To Npc_HighIndex
                If MapNpc(i).Num > 0 Then
                    MapNpc(i).Opacity = ENTITY_OPACITY_DUEL
                End If
            Next i

            If IsPlaying(InDuelTargetIndex) Then
                For i = 1 To Player_HighIndex
                    If i <> MyIndex Then
                        Player(i).Opacity = ENTITY_OPACITY_DUEL
                    Else
                        Player(i).Opacity = ENTITY_OPACITY_ORIGIN
                    End If
                Next i
            End If
    End If
End Sub


Public Sub ClearPlayerDuel()
    Dim i As Long
    
    '// Restaurar cor original dos npcs e seus pokemons caso tenha
    For i = 1 To Npc_HighIndex
        If MapNpc(i).Num > 0 Then
            MapNpc(i).Opacity = ENTITY_OPACITY_ORIGIN
        End If
    Next i
        
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            Player(i).Opacity = ENTITY_OPACITY_ORIGIN
        End If
    Next i
    
    For i = 1 To Pokemon_HighIndex
        If MapPokemon(i).Num > 0 Then
            MapPokemon(i).Opacity = ENTITY_OPACITY_ORIGIN
        End If
    Next i
    
    InDuelTargetType = 0
    InDuelTargetIndex = 0
    
End Sub
