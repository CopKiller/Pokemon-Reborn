Attribute VB_Name = "modFish"
Option Explicit

Private mapPokemonIds(MAX_MAP) As Collection

Private Sub Initialize()
    Dim i As Long
    For i = 1 To MAX_MAP
        Set mapPokemonIds(i) = New Collection
    Next i
End Sub

Public Sub AddPokemonIdToMap(ByVal mapIndex As Long, ByVal pokemonId As Long)
    mapPokemonIds(mapIndex).Add pokemonId
End Sub

Public Sub SpawnPokemonIdInMap(ByVal Index As Long, ByVal mapIndex As Long)
    Dim pokemonIds As Collection
    Dim pokemonId As Variant
    Dim Rand As Long

    Set pokemonIds = mapPokemonIds(mapIndex)

    If pokemonIds.count > 0 Then
RandomizaNovamente:
        Rand = Random(1, CLng(pokemonIds.count))
        'For Each pokemonId In pokemonIds
        'If IsWithinSpawnTime(pokemonId, GameHour) Then
            
        'End If
        'Next pokemonId
        
        'Debug.Print CLng(pokemonIds.Item(8))
        
        'Debug.Print pokemonIds.count

        If IsWithinSpawnTime(pokemonIds.Item(Rand), GameHour) Then
            Call SpawnMapPokemon(pokemonIds.Item(Rand), , , Index)
        Else
            GoTo RandomizaNovamente
        End If
        '
        
        'Debug.Print "Map Index: " & mapIndex & ", Pokemon ID: " & pokemonId
        'Next pokemonId
        'Else
        'Debug.Print "Map Index: " & mapIndex & ", No Pokemon IDs"
    End If
End Sub

Public Sub AddPokemonsFishing()
    Dim i As Long
    Initialize

    For i = 1 To MAX_GAME_POKEMON
        If Spawn(i).PokeNum > 0 And Spawn(i).Fishing And Spawn(i).MapNum > 0 Then
            AddPokemonIdToMap Spawn(i).MapNum, i
        End If
    Next i

    ' Exemplo de uso: verificar os IDs de Pokémon para cada mapa
    'Dim mapIndex As Long
    'For mapIndex = 1 To MAX_MAP
    '    SpawnPokemonIdInMap mapIndex
    'Next mapIndex
End Sub
