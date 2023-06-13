Attribute VB_Name = "modCombat"
Option Explicit

'//Ranges
Private Function IsOnLinearRange(ByVal Dir As Byte, ByVal Range As Long, ByVal startX As Long, ByVal startY As Long, ByVal CheckX As Long, ByVal CheckY As Long) As Boolean
    Select Case Dir
        Case DIR_UP
            If startX = CheckX Then
                If CheckY >= startY - Range And CheckY <= startY Then
                    IsOnLinearRange = True
                    Exit Function
                End If
            End If
        Case DIR_DOWN
            If startX = CheckX Then
                If CheckY >= startY And CheckY <= startY + Range Then
                    IsOnLinearRange = True
                    Exit Function
                End If
            End If
        Case DIR_LEFT
            If startY = CheckY Then
                If CheckX >= startX - Range And CheckX <= startX Then
                    IsOnLinearRange = True
                    Exit Function
                End If
            End If
        Case DIR_RIGHT
            If startY = CheckY Then
                If CheckX >= startX And CheckX <= startX + Range Then
                    IsOnLinearRange = True
                    Exit Function
                End If
            End If
    End Select
End Function

Private Function IsOnSprayRange(ByVal Dir As Byte, ByVal Range As Long, ByVal startX As Long, ByVal startY As Long, ByVal CheckX As Long, ByVal CheckY As Long) As Boolean
Dim checkExtra As Long

    Select Case Dir
        Case DIR_UP
            If CheckY >= startY - Range And CheckY <= startY Then
                checkExtra = startY - CheckY
                If CheckX >= startX - checkExtra And CheckX <= startX + checkExtra Then
                    IsOnSprayRange = True
                    Exit Function
                End If
            End If
        Case DIR_DOWN
            If CheckY >= startY And CheckY <= startY + Range Then
                checkExtra = CheckY - startY
                If CheckX >= startX - checkExtra And CheckX <= startX + checkExtra Then
                    IsOnSprayRange = True
                    Exit Function
                End If
            End If
        Case DIR_LEFT
            If CheckX >= startX - Range And CheckX <= startX Then
                checkExtra = startX - CheckX
                If CheckY >= startY - checkExtra And CheckY <= startY + checkExtra Then
                    IsOnSprayRange = True
                    Exit Function
                End If
            End If
        Case DIR_RIGHT
            If CheckX >= startX And CheckX <= startX + Range Then
                checkExtra = CheckX - startX
                If CheckY >= startY - checkExtra And CheckY <= startY + checkExtra Then
                    IsOnSprayRange = True
                    Exit Function
                End If
            End If
    End Select
End Function

'//AoE
Public Function IsOnAoERange(ByVal Range As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
Dim nVal As Long

    IsOnAoERange = False
    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    If nVal <= Range Then IsOnAoERange = True
End Function

'//STATS
Public Function CalculateSpeed(ByVal Spd As Long) As Long
Dim RangePercent As Long, VarPercent As Long

    On Error GoTo ErrorHandler
    
    RangePercent = Round(((Spd / 100) / (255 / 100)) * 100, 5)
    VarPercent = (100 - RangePercent)
    CalculateSpeed = (270 * (VarPercent / 100)) + 80
    
    Exit Function
ErrorHandler:
    CalculateSpeed = 255
End Function

Private Function GetNatureBoost(ByVal Nature As PokemonNature, ByVal Stat As StatEnum) As Single
    Select Case Nature
        '//Neutral
        Case PokemonNature.NatureHardy, PokemonNature.NatureDocile, PokemonNature.NatureSerious, PokemonNature.NatureBashful, PokemonNature.NatureQuirky
            GetNatureBoost = 0
            Exit Function
        Case PokemonNature.NatureLonely
            If Stat = StatEnum.Atk Then
                GetNatureBoost = 0.1
            ElseIf Stat = StatEnum.Def Then
                GetNatureBoost = -0.1
            Else
                GetNatureBoost = 0
            End If
            Exit Function
        Case PokemonNature.NatureBrave
            If Stat = StatEnum.Atk Then
                GetNatureBoost = 0.1
            ElseIf Stat = StatEnum.Spd Then
                GetNatureBoost = -0.1
            Else
                GetNatureBoost = 0
            End If
            Exit Function
        Case PokemonNature.NatureAdamant
            If Stat = StatEnum.Atk Then
                GetNatureBoost = 0.1
            ElseIf Stat = StatEnum.SpAtk Then
                GetNatureBoost = -0.1
            Else
                GetNatureBoost = 0
            End If
            Exit Function
        Case PokemonNature.NatureNaughty
            If Stat = StatEnum.Atk Then
                GetNatureBoost = 0.1
            ElseIf Stat = StatEnum.SpDef Then
                GetNatureBoost = -0.1
            Else
                GetNatureBoost = 0
            End If
            Exit Function
        Case PokemonNature.NatureBold
            If Stat = StatEnum.Def Then
                GetNatureBoost = 0.1
            ElseIf Stat = StatEnum.Atk Then
                GetNatureBoost = -0.1
            Else
                GetNatureBoost = 0
            End If
            Exit Function
        Case PokemonNature.NatureRelaxed
            If Stat = StatEnum.Def Then
                GetNatureBoost = 0.1
            ElseIf Stat = StatEnum.Spd Then
                GetNatureBoost = -0.1
            Else
                GetNatureBoost = 0
            End If
            Exit Function
        Case PokemonNature.NatureImpish
            If Stat = StatEnum.Def Then
                GetNatureBoost = 0.1
            ElseIf Stat = StatEnum.SpAtk Then
                GetNatureBoost = -0.1
            Else
                GetNatureBoost = 0
            End If
            Exit Function
        Case PokemonNature.NatureLax
            If Stat = StatEnum.Def Then
                GetNatureBoost = 0.1
            ElseIf Stat = StatEnum.SpDef Then
                GetNatureBoost = -0.1
            Else
                GetNatureBoost = 0
            End If
            Exit Function
        Case PokemonNature.NatureTimid
            If Stat = StatEnum.Spd Then
                GetNatureBoost = 0.1
            ElseIf Stat = StatEnum.Atk Then
                GetNatureBoost = -0.1
            Else
                GetNatureBoost = 0
            End If
            Exit Function
        Case PokemonNature.NatureHasty
            If Stat = StatEnum.Spd Then
                GetNatureBoost = 0.1
            ElseIf Stat = StatEnum.Def Then
                GetNatureBoost = -0.1
            Else
                GetNatureBoost = 0
            End If
            Exit Function
        Case PokemonNature.NatureJolly
            If Stat = StatEnum.Spd Then
                GetNatureBoost = 0.1
            ElseIf Stat = StatEnum.SpAtk Then
                GetNatureBoost = -0.1
            Else
                GetNatureBoost = 0
            End If
            Exit Function
        Case PokemonNature.NatureNaive
            If Stat = StatEnum.Spd Then
                GetNatureBoost = 0.1
            ElseIf Stat = StatEnum.SpDef Then
                GetNatureBoost = -0.1
            Else
                GetNatureBoost = 0
            End If
            Exit Function
        Case PokemonNature.NatureModest
            If Stat = StatEnum.SpAtk Then
                GetNatureBoost = 0.1
            ElseIf Stat = StatEnum.Atk Then
                GetNatureBoost = -0.1
            Else
                GetNatureBoost = 0
            End If
            Exit Function
        Case PokemonNature.NatureMild
            If Stat = StatEnum.SpAtk Then
                GetNatureBoost = 0.1
            ElseIf Stat = StatEnum.Def Then
                GetNatureBoost = -0.1
            Else
                GetNatureBoost = 0
            End If
            Exit Function
        Case PokemonNature.NatureQuiet
            If Stat = StatEnum.SpAtk Then
                GetNatureBoost = 0.1
            ElseIf Stat = StatEnum.Spd Then
                GetNatureBoost = -0.1
            Else
                GetNatureBoost = 0
            End If
            Exit Function
        Case PokemonNature.NatureRash
            If Stat = StatEnum.SpAtk Then
                GetNatureBoost = 0.1
            ElseIf Stat = StatEnum.SpDef Then
                GetNatureBoost = -0.1
            Else
                GetNatureBoost = 0
            End If
            Exit Function
        Case PokemonNature.NatureCalm
            If Stat = StatEnum.SpDef Then
                GetNatureBoost = 0.1
            ElseIf Stat = StatEnum.Atk Then
                GetNatureBoost = -0.1
            Else
                GetNatureBoost = 0
            End If
            Exit Function
        Case PokemonNature.NatureGentle
            If Stat = StatEnum.SpDef Then
                GetNatureBoost = 0.1
            ElseIf Stat = StatEnum.Def Then
                GetNatureBoost = -0.1
            Else
                GetNatureBoost = 0
            End If
            Exit Function
        Case PokemonNature.NatureSassy
            If Stat = StatEnum.SpDef Then
                GetNatureBoost = 0.1
            ElseIf Stat = StatEnum.Spd Then
                GetNatureBoost = -0.1
            Else
                GetNatureBoost = 0
            End If
            Exit Function
        Case PokemonNature.NatureCareful
            If Stat = StatEnum.SpDef Then
                GetNatureBoost = 0.1
            ElseIf Stat = StatEnum.SpAtk Then
                GetNatureBoost = -0.1
            Else
                GetNatureBoost = 0
            End If
            Exit Function
    End Select
End Function

Private Function GetTypeBoost(ByVal SelfType As PokemonType, ByVal targetType As PokemonType) As Single
    Select Case SelfType
        Case PokemonType.typeNormal
            Select Case targetType
                '//Null
                Case PokemonType.typeGhost
                    GetTypeBoost = 0
                    Exit Function
                '//Not Effective
                Case PokemonType.typeRock, PokemonType.typeSteel
                    GetTypeBoost = 0.5
                    Exit Function
                '//Normal
                Case Else
                    GetTypeBoost = 1
                    Exit Function
            End Select
        Case PokemonType.typeFire
            Select Case targetType
                '//Effective
                Case PokemonType.typeGrass, PokemonType.typeIce, PokemonType.typeBug, PokemonType.typeSteel
                    GetTypeBoost = 2
                    Exit Function
                '//Not Effective
                Case PokemonType.typeFire, PokemonType.typeWater, PokemonType.typeRock, PokemonType.typeDragon
                    GetTypeBoost = 0.5
                    Exit Function
                '//Normal
                Case Else
                    GetTypeBoost = 1
                    Exit Function
            End Select
        Case PokemonType.typeWater
            Select Case targetType
                '//Effective
                Case PokemonType.typeFire, PokemonType.typeGround, PokemonType.typeRock
                    GetTypeBoost = 2
                    Exit Function
                '//Not Effective
                Case PokemonType.typeWater, PokemonType.typeGrass, PokemonType.typeDragon
                    GetTypeBoost = 0.5
                    Exit Function
                '//Normal
                Case Else
                    GetTypeBoost = 1
                    Exit Function
            End Select
        Case PokemonType.typeElectric
            Select Case targetType
                '//Effective
                Case PokemonType.typeWater, PokemonType.typeFlying
                    GetTypeBoost = 2
                    Exit Function
                '//Not Effective
                Case PokemonType.typeElectric, PokemonType.typeGrass, PokemonType.typeDragon
                    GetTypeBoost = 0.5
                    Exit Function
                '//Null
                Case PokemonType.typeGround, PokemonType.typeRock
                    GetTypeBoost = 0
                    Exit Function
                '//Normal
                Case Else
                    GetTypeBoost = 1
                    Exit Function
            End Select
        Case PokemonType.typeGrass
            Select Case targetType
                '//Effective
                Case PokemonType.typeWater, PokemonType.typeGround, PokemonType.typeRock
                    GetTypeBoost = 2
                    Exit Function
                '//Not Effective
                Case PokemonType.typeFire, PokemonType.typeGrass, PokemonType.typePoison, PokemonType.typeFlying, PokemonType.typeBug, PokemonType.typeDragon, PokemonType.typeSteel
                    GetTypeBoost = 0.5
                    Exit Function
                '//Normal
                Case Else
                    GetTypeBoost = 1
                    Exit Function
            End Select
        Case PokemonType.typeIce
            Select Case targetType
                '//Effective
                Case PokemonType.typeGrass, PokemonType.typeGround, PokemonType.typeFlying, PokemonType.typeDragon
                    GetTypeBoost = 2
                    Exit Function
                '//Not Effective
                Case PokemonType.typeFire, PokemonType.typeWater, PokemonType.typeIce, PokemonType.typeSteel
                    GetTypeBoost = 0.5
                    Exit Function
                '//Normal
                Case Else
                    GetTypeBoost = 1
                    Exit Function
            End Select
        Case PokemonType.typeFighting
            Select Case targetType
                '//Effective
                Case PokemonType.typeNormal, PokemonType.typeIce, PokemonType.typeRock, PokemonType.typeDark, PokemonType.typeSteel
                    GetTypeBoost = 2
                    Exit Function
                '//Not Effective
                Case PokemonType.typePoison, PokemonType.typeFlying, PokemonType.typePsychic, PokemonType.typeBug, PokemonType.typeFairy
                    GetTypeBoost = 0.5
                    Exit Function
                '//Null
                Case PokemonType.typeGhost
                    GetTypeBoost = 0
                    Exit Function
                '//Normal
                Case Else
                    GetTypeBoost = 1
                    Exit Function
            End Select
        Case PokemonType.typePoison
            Select Case targetType
                '//Effective
                Case PokemonType.typeGrass, PokemonType.typeFairy
                    GetTypeBoost = 2
                    Exit Function
                '//Not Effective
                Case PokemonType.typePoison, PokemonType.typeGround, PokemonType.typeRock, PokemonType.typeGhost
                    GetTypeBoost = 0.5
                    Exit Function
                '//Null
                Case PokemonType.typeSteel
                    GetTypeBoost = 0
                    Exit Function
                '//Normal
                Case Else
                    GetTypeBoost = 1
                    Exit Function
            End Select
        Case PokemonType.typeGround
            Select Case targetType
                '//Effective
                Case PokemonType.typeFire, PokemonType.typeElectric, PokemonType.typePoison, PokemonType.typeRock, PokemonType.typeSteel
                    GetTypeBoost = 2
                    Exit Function
                '//Not Effective
                Case PokemonType.typeGrass, PokemonType.typeBug
                    GetTypeBoost = 0.5
                    Exit Function
                '//Null
                Case PokemonType.typeFlying
                    GetTypeBoost = 0
                    Exit Function
                '//Normal
                Case Else
                    GetTypeBoost = 1
                    Exit Function
            End Select
        Case PokemonType.typeFlying
            Select Case targetType
                '//Effective
                Case PokemonType.typeGrass, PokemonType.typeFighting, PokemonType.typeBug
                    GetTypeBoost = 2
                    Exit Function
                '//Not Effective
                Case PokemonType.typeElectric, PokemonType.typeRock, PokemonType.typeSteel
                    GetTypeBoost = 0.5
                    Exit Function
                '//Normal
                Case Else
                    GetTypeBoost = 1
                    Exit Function
            End Select
        Case PokemonType.typePsychic
            Select Case targetType
                '//Effective
                Case PokemonType.typeFighting, PokemonType.typePoison
                    GetTypeBoost = 2
                    Exit Function
                '//Not Effective
                Case PokemonType.typePsychic, PokemonType.typeSteel
                    GetTypeBoost = 0.5
                    Exit Function
                '//Null
                Case PokemonType.typeDark
                    GetTypeBoost = 0
                    Exit Function
                '//Normal
                Case Else
                    GetTypeBoost = 1
                    Exit Function
            End Select
        Case PokemonType.typeBug
            Select Case targetType
                '//Effective
                Case PokemonType.typeGrass, PokemonType.typePsychic, PokemonType.typeDark
                    GetTypeBoost = 2
                    Exit Function
                '//Not Effective
                Case PokemonType.typeFire, PokemonType.typeFighting, PokemonType.typePoison, PokemonType.typeFlying, PokemonType.typeGhost, PokemonType.typeSteel, PokemonType.typeFairy
                    GetTypeBoost = 0.5
                    Exit Function
                '//Normal
                Case Else
                    GetTypeBoost = 1
                    Exit Function
            End Select
        Case PokemonType.typeRock
            Select Case targetType
                '//Effective
                Case PokemonType.typeFire, PokemonType.typeIce, PokemonType.typeFlying, PokemonType.typeBug
                    GetTypeBoost = 2
                    Exit Function
                '//Not Effective
                Case PokemonType.typeFighting, PokemonType.typeGround, PokemonType.typeSteel
                    GetTypeBoost = 0.5
                    Exit Function
                '//Normal
                Case Else
                    GetTypeBoost = 1
                    Exit Function
            End Select
        Case PokemonType.typeGhost
            Select Case targetType
                '//Effective
                Case PokemonType.typePsychic, PokemonType.typeGhost
                    GetTypeBoost = 2
                    Exit Function
                '//Not Effective
                Case PokemonType.typeDark
                    GetTypeBoost = 0.5
                    Exit Function
                '//Null
                Case PokemonType.typeNormal
                    GetTypeBoost = 0
                    Exit Function
                '//Normal
                Case Else
                    GetTypeBoost = 1
                    Exit Function
            End Select
        Case PokemonType.typeDragon
            Select Case targetType
                '//Effective
                Case PokemonType.typeDragon
                    GetTypeBoost = 2
                    Exit Function
                '//Not Effective
                Case PokemonType.typeSteel
                    GetTypeBoost = 0.5
                    Exit Function
                '//Null
                Case PokemonType.typeFairy
                    GetTypeBoost = 0
                    Exit Function
                '//Normal
                Case Else
                    GetTypeBoost = 1
                    Exit Function
            End Select
        Case PokemonType.typeDark
            Select Case targetType
                '//Effective
                Case PokemonType.typePsychic, PokemonType.typeGhost
                    GetTypeBoost = 2
                    Exit Function
                '//Not Effective
                Case PokemonType.typeFighting, PokemonType.typeDark, PokemonType.typeFairy
                    GetTypeBoost = 0.5
                    Exit Function
                '//Normal
                Case Else
                    GetTypeBoost = 1
                    Exit Function
            End Select
        Case PokemonType.typeSteel
            Select Case targetType
                '//Effective
                Case PokemonType.typeIce, PokemonType.typeRock, PokemonType.typeFairy
                    GetTypeBoost = 2
                    Exit Function
                '//Not Effective
                Case PokemonType.typeFire, PokemonType.typeWater, PokemonType.typeElectric, PokemonType.typeSteel
                    GetTypeBoost = 0.5
                    Exit Function
                '//Normal
                Case Else
                    GetTypeBoost = 1
                    Exit Function
            End Select
        Case PokemonType.typeFairy
            Select Case targetType
                '//Effective
                Case PokemonType.typeFighting, PokemonType.typeDragon, PokemonType.typeDark
                    GetTypeBoost = 2
                    Exit Function
                '//Not Effective
                Case PokemonType.typeFire, PokemonType.typePoison, PokemonType.typeSteel
                    GetTypeBoost = 0.5
                    Exit Function
                '//Normal
                Case Else
                    GetTypeBoost = 1
                    Exit Function
            End Select
        Case Else
            GetTypeBoost = 1
            Exit Function
    End Select
End Function

Public Function CalculatePokemonStat(ByVal Stat As StatEnum, ByVal PokeNum As Long, ByVal Level As Byte, ByVal StatEV As Long, ByVal StatIV As Long, ByVal Nature As PokemonNature) As Long
Dim TotalValue As Long

    If Stat = HP Then
        '//HP calculation is different than the rest
        CalculatePokemonStat = ((StatIV + 2 * Pokemon(PokeNum).BaseStat(Stat) + (StatEV / 4)) * Level / 100) + 10 + Level
    Else
        TotalValue = (((StatIV + 2 * Pokemon(PokeNum).BaseStat(Stat) + (StatEV / 4)) * Level / 100) + 5)
        CalculatePokemonStat = TotalValue + (TotalValue * GetNatureBoost(Nature, Stat))
        '//Nature Value
        '//Increase of Stat = 110%
        '//Decrease of Stat = 90%
    End If
End Function

Public Function GetPlayerPokemonStat(ByVal Index As Long, ByVal Stat As StatEnum) As Long
    '//Check for error
    If Not IsPlaying(Index) Then Exit Function
    If TempPlayer(Index).UseChar = 0 Then Exit Function
    If PlayerPokemon(Index).Num <= 0 Then Exit Function
    If PlayerPokemon(Index).slot <= 0 Then Exit Function
    
    With PlayerPokemons(Index).Data(PlayerPokemon(Index).slot)
        '//Select Buff Stage
        Select Case PlayerPokemon(Index).StatBuff(Stat)
            Case -6: GetPlayerPokemonStat = .Stat(Stat).Value * 0.25
            Case -5: GetPlayerPokemonStat = .Stat(Stat).Value * 0.285
            Case -4: GetPlayerPokemonStat = .Stat(Stat).Value * 0.33
            Case -3: GetPlayerPokemonStat = .Stat(Stat).Value * 0.4
            Case -2: GetPlayerPokemonStat = .Stat(Stat).Value * 0.5
            Case -1: GetPlayerPokemonStat = .Stat(Stat).Value * 0.66
            Case 0: GetPlayerPokemonStat = .Stat(Stat).Value * 1
            Case 1: GetPlayerPokemonStat = .Stat(Stat).Value * 1.5
            Case 2: GetPlayerPokemonStat = .Stat(Stat).Value * 2
            Case 3: GetPlayerPokemonStat = .Stat(Stat).Value * 2.5
            Case 4: GetPlayerPokemonStat = .Stat(Stat).Value * 3
            Case 5: GetPlayerPokemonStat = .Stat(Stat).Value * 3.5
            Case 6: GetPlayerPokemonStat = .Stat(Stat).Value * 4
            Case Else:
        End Select
    End With
End Function

Public Function GetNpcPokemonStat(ByVal MapPokeNum As Long, ByVal Stat As StatEnum) As Long
    '//Check for error
    If MapPokeNum <= 0 Or MapPokeNum > MAX_GAME_POKEMON Then Exit Function
    If MapPokemon(MapPokeNum).Num <= 0 Then Exit Function
    
    With MapPokemon(MapPokeNum)
        '//Select Buff Stage
        Select Case MapPokemon(MapPokeNum).StatBuff(Stat)
            Case -6: GetNpcPokemonStat = .Stat(Stat).Value * 0.25
            Case -5: GetNpcPokemonStat = .Stat(Stat).Value * 0.285
            Case -4: GetNpcPokemonStat = .Stat(Stat).Value * 0.33
            Case -3: GetNpcPokemonStat = .Stat(Stat).Value * 0.4
            Case -2: GetNpcPokemonStat = .Stat(Stat).Value * 0.5
            Case -1: GetNpcPokemonStat = .Stat(Stat).Value * 0.66
            Case 0: GetNpcPokemonStat = .Stat(Stat).Value * 1
            Case 1: GetNpcPokemonStat = .Stat(Stat).Value * 1.5
            Case 2: GetNpcPokemonStat = .Stat(Stat).Value * 2
            Case 3: GetNpcPokemonStat = .Stat(Stat).Value * 2.5
            Case 4: GetNpcPokemonStat = .Stat(Stat).Value * 3
            Case 5: GetNpcPokemonStat = .Stat(Stat).Value * 3.5
            Case 6: GetNpcPokemonStat = .Stat(Stat).Value * 4
            Case Else:
        End Select
    End With
End Function

Public Function GetNpcTrainerPokemonStat(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Stat As StatEnum) As Long
    '//Check for error
    If MapNum <= 0 Or MapNum > MAX_MAP Then Exit Function
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPC Then Exit Function
    If MapNpcPokemon(MapNum, MapNpcNum).Num <= 0 Then Exit Function
    
    With MapNpcPokemon(MapNum, MapNpcNum)
        '//Select Buff Stage
        Select Case .StatBuff(Stat)
            Case -6: GetNpcTrainerPokemonStat = .Stat(Stat).Value * 0.25
            Case -5: GetNpcTrainerPokemonStat = .Stat(Stat).Value * 0.285
            Case -4: GetNpcTrainerPokemonStat = .Stat(Stat).Value * 0.33
            Case -3: GetNpcTrainerPokemonStat = .Stat(Stat).Value * 0.4
            Case -2: GetNpcTrainerPokemonStat = .Stat(Stat).Value * 0.5
            Case -1: GetNpcTrainerPokemonStat = .Stat(Stat).Value * 0.66
            Case 0: GetNpcTrainerPokemonStat = .Stat(Stat).Value * 1
            Case 1: GetNpcTrainerPokemonStat = .Stat(Stat).Value * 1.5
            Case 2: GetNpcTrainerPokemonStat = .Stat(Stat).Value * 2
            Case 3: GetNpcTrainerPokemonStat = .Stat(Stat).Value * 2.5
            Case 4: GetNpcTrainerPokemonStat = .Stat(Stat).Value * 3
            Case 5: GetNpcTrainerPokemonStat = .Stat(Stat).Value * 3.5
            Case 6: GetNpcTrainerPokemonStat = .Stat(Stat).Value * 4
            Case Else:
        End Select
    End With
End Function

Public Function GetMapNpcPokemonStat(ByVal MapNum As Long, ByVal MapPokeNum As Long, ByVal Stat As StatEnum) As Long
    '//Check for error
    If MapNum <= 0 Or MapNum > MAX_MAP Then Exit Function
    If MapPokeNum <= 0 Or MapPokeNum > MAX_MAP_NPC Then Exit Function
    If MapNpcPokemon(MapNum, MapPokeNum).Num <= 0 Then Exit Function
    
    With MapNpcPokemon(MapNum, MapPokeNum)
        '//Select Buff Stage
        Select Case MapNpcPokemon(MapNum, MapPokeNum).StatBuff(Stat)
            Case -6: GetMapNpcPokemonStat = .Stat(Stat).Value * 0.25
            Case -5: GetMapNpcPokemonStat = .Stat(Stat).Value * 0.285
            Case -4: GetMapNpcPokemonStat = .Stat(Stat).Value * 0.33
            Case -3: GetMapNpcPokemonStat = .Stat(Stat).Value * 0.4
            Case -2: GetMapNpcPokemonStat = .Stat(Stat).Value * 0.5
            Case -1: GetMapNpcPokemonStat = .Stat(Stat).Value * 0.66
            Case 0: GetMapNpcPokemonStat = .Stat(Stat).Value * 1
            Case 1: GetMapNpcPokemonStat = .Stat(Stat).Value * 1.5
            Case 2: GetMapNpcPokemonStat = .Stat(Stat).Value * 2
            Case 3: GetMapNpcPokemonStat = .Stat(Stat).Value * 2.5
            Case 4: GetMapNpcPokemonStat = .Stat(Stat).Value * 3
            Case 5: GetMapNpcPokemonStat = .Stat(Stat).Value * 3.5
            Case 6: GetMapNpcPokemonStat = .Stat(Stat).Value * 1
            Case Else:
        End Select
    End With
End Function

Public Function GetPokemonDamage(ByVal ownType As Long, ByVal MoveType As Long, ByVal targetType As Long, ByVal targetType2 As Long, ByVal Level As Byte, ByVal AtkStat As Long, ByVal AtkPower As Long, ByVal DefStat As Long)
Dim boostType As Single, boostType2 As Single, totalBoost As Single

    On Error GoTo ErrorHandler
    
    boostType = GetTypeBoost(MoveType, targetType)
    If targetType2 > 0 Then
        boostType2 = GetTypeBoost(MoveType, targetType2)
        totalBoost = boostType * boostType2
    Else
        totalBoost = boostType
    End If
    
    'If AtkStat <= 0 Then AtkStat = 1
    'If DefStat <= 0 Then DefStat = 1
    'If AtkPower <= 0 Then AtkPower = 1
    
    If ownType = MoveType Then
        GetPokemonDamage = ((((2 * Level / 5 + 2) * AtkStat * AtkPower / DefStat) / 50) + 2) * 1.5 * totalBoost * Random(85, 100) / 100
    Else
        GetPokemonDamage = ((((2 * Level / 5 + 2) * AtkStat * AtkPower / DefStat) / 50) + 2) * 1 * totalBoost * Random(85, 100) / 100
    End If
    
    Exit Function
ErrorHandler:
    GetPokemonDamage = 0
End Function

Private Function IsImmuneOnStatus(ByVal MoveType As Byte, ByVal PrimaryType As Byte, ByVal SecondaryType As Byte, ByVal StatusType As Byte) As Boolean
    If StatusType >= 6 Then
        IsImmuneOnStatus = False
        Exit Function
    End If
    If StatusType = StatusEnum.Sleep Then
        IsImmuneOnStatus = True
        Exit Function
    End If
    
    IsImmuneOnStatus = True
    If MoveType = PrimaryType Then
        IsImmuneOnStatus = False
        Exit Function
    End If
    If GetTypeBoost(MoveType, PrimaryType) = 0 Then
        IsImmuneOnStatus = False
        Exit Function
    End If
    If SecondaryType > 0 Then
        If MoveType = SecondaryType Then
            IsImmuneOnStatus = False
            Exit Function
        End If
        If GetTypeBoost(MoveType, SecondaryType) = 0 Then
            IsImmuneOnStatus = False
            Exit Function
        End If
    End If
End Function

'/////////////////////////
'///// Player Attack /////
'/////////////////////////
Public Sub PlayerCastMove(ByVal Index As Long, ByVal MoveNum As Long, ByVal MoveSlot As Byte, Optional ByVal DecreasePP As Boolean = True)
Dim RandomNum As Byte
Dim DuelIndex As Long

    '//Check for error
    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar = 0 Then Exit Sub
    If PlayerPokemon(Index).Num <= 0 Then Exit Sub
    If MoveNum <= 0 Or MoveNum > MAX_POKEMON_MOVE Then Exit Sub
    If PlayerPokemon(Index).slot <= 0 Then Exit Sub
    
    '//Add Queue
    With PlayerPokemon(Index)
        If Not PokemonMove(MoveNum).SelfStatusReq = StatusEnum.Sleep Then
            If PlayerPokemons(Index).Data(.slot).Status = StatusEnum.Sleep Then
                RandomNum = Random(1, 3)
                If RandomNum = 1 Then
                    '//Remove Status
                    PlayerPokemons(Index).Data(.slot).Status = 0
                    SendPlayerPokemonStatus Index
                    Select Case TempPlayer(Index).CurLanguage
                        Case LANG_PT: AddAlert Index, "Your pokemon is woke up", White
                        Case LANG_EN: AddAlert Index, "Your pokemon is woke up", White
                        Case LANG_ES: AddAlert Index, "Your pokemon is woke up", White
                    End Select
                Else
                    Select Case TempPlayer(Index).CurLanguage
                        Case LANG_PT: AddAlert Index, "Your pokemon is fast asleep", White
                        Case LANG_EN: AddAlert Index, "Your pokemon is fast asleep", White
                        Case LANG_ES: AddAlert Index, "Your pokemon is fast asleep", White
                    End Select
                    Exit Sub
                End If
            End If
        End If
        If Not PokemonMove(MoveNum).SelfStatusReq = StatusEnum.Frozen Then
            If PlayerPokemons(Index).Data(.slot).Status = StatusEnum.Frozen Then
                RandomNum = Random(1, 3)
                If RandomNum = 1 Then
                    '//Remove Status
                    PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Status = 0
                    SendPlayerPokemonStatus Index
                Else
                    Select Case TempPlayer(Index).CurLanguage
                        Case LANG_PT: AddAlert Index, "Your pokemon is frozen", White
                        Case LANG_EN: AddAlert Index, "Your pokemon is frozen", White
                        Case LANG_ES: AddAlert Index, "Your pokemon is frozen", White
                    End Select
                    Exit Sub
                End If
            End If
        End If
        
        If PokemonMove(MoveNum).SelfStatusReq > 0 Then
            If Not PlayerPokemons(Index).Data(.slot).Status = PokemonMove(MoveNum).SelfStatusReq Then
                Exit Sub
            End If
        End If
        
        '//Check PP
        If MoveSlot > 0 Then
            If PlayerPokemons(Index).Data(.slot).Moveset(MoveSlot).CurPP <= 0 Then
                Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "Out of PP", White
                    Case LANG_EN: AddAlert Index, "Out of PP", White
                    Case LANG_ES: AddAlert Index, "Out of PP", White
                End Select
                Exit Sub
            End If
            '//Check Cooldown
            If PlayerPokemons(Index).Data(.slot).Moveset(MoveSlot).CD + PokemonMove(MoveNum).Cooldown > GetTickCount Then
                Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "This move need to be recharged", White
                    Case LANG_EN: AddAlert Index, "This move need to be recharged", White
                    Case LANG_ES: AddAlert Index, "This move need to be recharged", White
                End Select
                Exit Sub
            End If
        End If
        
        '//Burn
        If PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Status = StatusEnum.Burn Then
            If PlayerPokemon(Index).StatusDamage > 0 Then
                If PlayerPokemon(Index).StatusDamage >= PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp Then
                    '//Dead
                    PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp = 0
                    SendActionMsg Player(Index, TempPlayer(Index).UseChar).Map, "-" & PlayerPokemon(Index).StatusDamage, PlayerPokemon(Index).X * 32, PlayerPokemon(Index).Y * 32, BrightRed
                    SendPlayerPokemonVital Index
                    SendPlayerPokemonFaint Index
                    Exit Sub
                Else
                    '//Reduce
                    PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp = PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp - PlayerPokemon(Index).StatusDamage
                    SendActionMsg Player(Index, TempPlayer(Index).UseChar).Map, "-" & PlayerPokemon(Index).StatusDamage, PlayerPokemon(Index).X * 32, PlayerPokemon(Index).Y * 32, BrightRed
                    '//Update
                    SendPlayerPokemonVital Index
                End If
            Else
                PlayerPokemon(Index).StatusDamage = (PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).MaxHp / 8)
            End If
        End If
        
        .QueueMove = MoveNum
        .QueueMoveSlot = MoveSlot
        
        '//Set Duration
        .MoveCastTime = GetTickCount + PokemonMove(MoveNum).CastTime
        .MoveDuration = GetTickCount + PokemonMove(MoveNum).Duration
        .MoveInterval = GetTickCount
        .MoveAttackCount = 0
        
        '//Decrease PP
        If MoveSlot > 0 Then
            If DecreasePP Then
                PlayerPokemons(Index).Data(.slot).Moveset(MoveSlot).CurPP = PlayerPokemons(Index).Data(.slot).Moveset(MoveSlot).CurPP - 1
                SendPlayerPokemonPP Index, MoveSlot
            End If
            
            '//Add ActionMsg
            SendActionMsg Player(Index, TempPlayer(Index).UseChar).Map, Trim$(PokemonMove(MoveNum).Name), .X * 32, .Y * 32, Yellow
        End If
    End With
End Sub

Public Sub ProcessPlayerMove(ByVal Index As Long, ByVal MoveNum As Long)
    Dim I As Long
    Dim Range As Long
    Dim MapNum As Long
    Dim X As Long, Y As Long
    Dim pType As Byte
    Dim Power As Long

    '//Damage Calculation
    Dim ownType As Byte
    Dim ownLevel As Byte
    Dim AtkStat As Long
    Dim Damage As Long
    Dim targetType As Byte, targetType2 As Byte
    Dim DefStat As Long

    Dim InRange As Boolean
    Dim InRange2 As Boolean
    Dim z As Byte
    Dim HealAmount As Long
    Dim statusChance As Long
    Dim statusRand As Long
    Dim DuelIndex As Long
    Dim recoil As Long
    Dim Absorbed As Long
    Dim CanAttack As Boolean
    Dim setBuff As Long
    Dim expEarn As Long

    '//Check for error
    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar = 0 Then Exit Sub
    If PlayerPokemon(Index).Num <= 0 Then Exit Sub
    If MoveNum <= 0 Or MoveNum > MAX_POKEMON_MOVE Then Exit Sub
    If PlayerPokemon(Index).slot <= 0 Then Exit Sub

    MapNum = Player(Index, TempPlayer(Index).UseChar).Map
    Range = PokemonMove(MoveNum).Range
    Power = PokemonMove(MoveNum).Power
    X = PlayerPokemon(Index).X
    Y = PlayerPokemon(Index).Y
    ownType = Pokemon(PlayerPokemon(Index).Num).PrimaryType
    ownLevel = PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Level
    pType = PokemonMove(MoveNum).Type

    '//Check Attack Category
    Select Case PokemonMove(MoveNum).Category
    Case MoveCategory.Physical
        AtkStat = GetPlayerPokemonStat(Index, Atk)
    Case MoveCategory.Special
        AtkStat = GetPlayerPokemonStat(Index, SpAtk)
    End Select

    '//Get Target
    Select Case PokemonMove(MoveNum).targetType
    Case 0    '//Self
        '//Status
        If PokemonMove(MoveNum).pStatus > 0 And PokemonMove(MoveNum).StatusToSelf = NO Then
            If PokemonMove(MoveNum).pStatus = 6 Then
                PlayerPokemon(Index).IsConfuse = YES
                SendPlayerPokemonStatus Index
                Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "Your pokemon got confused", White
                Case LANG_EN: AddAlert Index, "Your pokemon got confused", White
                Case LANG_ES: AddAlert Index, "Your pokemon got confused", White
                End Select
            Else
                If PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Status <= 0 Then
                    statusChance = (100 * (PokemonMove(MoveNum).pStatusChance / 100))

                    If IsImmuneOnStatus(PokemonMove(MoveNum).Type, Pokemon(PlayerPokemon(Index).Num).PrimaryType, Pokemon(PlayerPokemon(Index).Num).SecondaryType, PokemonMove(MoveNum).pStatus) Then
                        If statusChance > 0 Then
                            statusRand = Random(1, 100)
                            If statusRand <= statusChance Then
                                PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Status = PokemonMove(MoveNum).pStatus
                                SendPlayerPokemonStatus Index
                                Select Case PokemonMove(MoveNum).pStatus
                                Case StatusEnum.Poison
                                    Select Case TempPlayer(Index).CurLanguage
                                    Case LANG_PT: AddAlert Index, "Your pokemon got poisoned", White
                                    Case LANG_EN: AddAlert Index, "Your pokemon got poisoned", White
                                    Case LANG_ES: AddAlert Index, "Your pokemon got poisoned", White
                                    End Select
                                Case StatusEnum.Burn
                                    Select Case TempPlayer(Index).CurLanguage
                                    Case LANG_PT: AddAlert Index, "Your pokemon got burned", White
                                    Case LANG_EN: AddAlert Index, "Your pokemon got burned", White
                                    Case LANG_ES: AddAlert Index, "Your pokemon got burned", White
                                    End Select
                                Case StatusEnum.Paralize
                                    Select Case TempPlayer(Index).CurLanguage
                                    Case LANG_PT: AddAlert Index, "Your pokemon got paralized", White
                                    Case LANG_EN: AddAlert Index, "Your pokemon got paralized", White
                                    Case LANG_ES: AddAlert Index, "Your pokemon got paralized", White
                                    End Select
                                Case StatusEnum.Sleep
                                    Select Case TempPlayer(Index).CurLanguage
                                    Case LANG_PT: AddAlert Index, "Your pokemon fell asleep", White
                                    Case LANG_EN: AddAlert Index, "Your pokemon fell asleep", White
                                    Case LANG_ES: AddAlert Index, "Your pokemon fell asleep", White
                                    End Select
                                Case StatusEnum.Frozen
                                    Select Case TempPlayer(Index).CurLanguage
                                    Case LANG_PT: AddAlert Index, "Your pokemon got frozed", White
                                    Case LANG_EN: AddAlert Index, "Your pokemon got frozed", White
                                    Case LANG_ES: AddAlert Index, "Your pokemon got frozed", White
                                    End Select
                                End Select
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Select Case PokemonMove(MoveNum).AttackType
        Case 2    '//Buff/Debuff
            For z = 1 To StatEnum.Stat_Count - 1
                PlayerPokemon(Index).StatBuff(z) = PlayerPokemon(Index).StatBuff(z) + PokemonMove(MoveNum).dStat(z)
                If PlayerPokemon(Index).StatBuff(z) > 6 Then
                    PlayerPokemon(Index).StatBuff(z) = 6
                ElseIf PlayerPokemon(Index).StatBuff(z) < -6 Then
                    PlayerPokemon(Index).StatBuff(z) = -6
                End If
            Next
            SendPlayerPokemonStatBuff Index
        Case 3    '//Heal
            HealAmount = PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).MaxHp * (PokemonMove(MoveNum).Power / 100)
            PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp = PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp + HealAmount
            SendActionMsg MapNum, "+" & HealAmount, PlayerPokemon(Index).X * 32, PlayerPokemon(Index).Y * 32, BrightGreen
            If PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp >= PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).MaxHp Then
                PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp = PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).MaxHp
            End If
            SendPlayerPokemonVital Index
            If PokemonMove(MoveNum).pStatus = 7 Then
                '//Cure Status
                PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Status = 0
                Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "Your pokemon got cured", White
                Case LANG_EN: AddAlert Index, "Your pokemon got cured", White
                Case LANG_ES: AddAlert Index, "Your pokemon got frozed", White
                End Select
                SendPlayerPokemonStatus Index
            End If
        End Select
        '//Reflect
        If PokemonMove(MoveNum).ReflectType > 0 Then
            PlayerPokemon(Index).ReflectMove = PokemonMove(MoveNum).ReflectType
        End If
        If PokemonMove(MoveNum).CastProtect > 0 Then
            PlayerPokemon(Index).IsProtect = YES
        End If
    Case 1, 2, 3    '//Linear , AOE , Spray
        '//Check Target
        '//Player
        For I = 1 To Player_HighIndex
            If IsPlaying(I) Then
                If TempPlayer(I).UseChar > 0 Then
                    If Player(I, TempPlayer(I).UseChar).Map = MapNum Then
                        '//Can't kill player
                        If PlayerPokemon(I).Num > 0 Then
                            '//Check Status Req
                            If PokemonMove(MoveNum).StatusReq > 0 Then
                                If PlayerPokemons(I).Data(PlayerPokemon(I).slot).Status = PokemonMove(MoveNum).StatusReq Then
                                    CanAttack = True
                                Else
                                    CanAttack = False
                                End If
                            Else
                                CanAttack = True
                            End If
                            If PlayerPokemon(I).IsProtect > 0 Then
                                CanAttack = False
                                PlayerPokemon(I).IsProtect = NO
                                SendActionMsg MapNum, "Protected", PlayerPokemon(I).X * 32, PlayerPokemon(I).Y * 32, Yellow
                            End If

                            If CanAttack Then
                                InRange = False
                                If PokemonMove(MoveNum).targetType = 1 Then    '//AoE
                                    If IsOnAoERange(Range, X, Y, PlayerPokemon(I).X, PlayerPokemon(I).Y) Then InRange = True
                                ElseIf PokemonMove(MoveNum).targetType = 2 Then    '//Linear
                                    If IsOnLinearRange(PlayerPokemon(Index).Dir, Range, X, Y, PlayerPokemon(I).X, PlayerPokemon(I).Y) Then InRange = True
                                ElseIf PokemonMove(MoveNum).targetType = 3 Then    '//Spray
                                    If IsOnSprayRange(PlayerPokemon(Index).Dir, Range, X, Y, PlayerPokemon(I).X, PlayerPokemon(I).Y) Then InRange = True
                                Else
                                    InRange = False
                                End If

                                If InRange Then
                                    If Not I = Index Then
                                        If (TempPlayer(Index).InDuel = I And TempPlayer(I).DuelTime <= 0) Or Map(MapNum).Moral = MAP_MORAL_PVP Or (Player(I, TempPlayer(I).UseChar).Access >= ACCESS_CREATOR) Or (Player(Index, TempPlayer(Index).UseChar).Access >= ACCESS_CREATOR) Then
                                            If PlayerPokemon(I).slot > 0 Then
                                                If PokemonMove(MoveNum).pStatus = 6 Then
                                                    PlayerPokemon(I).IsConfuse = YES
                                                    SendPlayerPokemonStatus I
                                                    Select Case TempPlayer(I).CurLanguage
                                                    Case LANG_PT: AddAlert I, "Your pokemon got confused", White
                                                    Case LANG_EN: AddAlert I, "Your pokemon got confused", White
                                                    Case LANG_ES: AddAlert I, "Your pokemon got confused", White
                                                    End Select
                                                Else
                                                    '//Status
                                                    If PokemonMove(MoveNum).pStatus > 0 And PokemonMove(MoveNum).StatusToSelf = NO Then
                                                        If PlayerPokemons(I).Data(PlayerPokemon(I).slot).Status <= 0 Then
                                                            statusChance = (100 * (PokemonMove(MoveNum).pStatusChance / 100))

                                                            If IsImmuneOnStatus(PokemonMove(MoveNum).Type, Pokemon(PlayerPokemon(I).Num).PrimaryType, Pokemon(PlayerPokemon(I).Num).SecondaryType, PokemonMove(MoveNum).pStatus) Then
                                                                If statusChance > 0 Then
                                                                    statusRand = Random(1, 100)
                                                                    If statusRand <= statusChance Then
                                                                        PlayerPokemons(I).Data(PlayerPokemon(I).slot).Status = PokemonMove(MoveNum).pStatus
                                                                        SendPlayerPokemonStatus I
                                                                        Select Case PokemonMove(MoveNum).pStatus
                                                                        Case StatusEnum.Poison
                                                                            Select Case TempPlayer(I).CurLanguage
                                                                            Case LANG_PT: AddAlert I, "Your pokemon got poisoned", White
                                                                            Case LANG_EN: AddAlert I, "Your pokemon got poisoned", White
                                                                            Case LANG_ES: AddAlert I, "Your pokemon got poisoned", White
                                                                            End Select
                                                                        Case StatusEnum.Burn
                                                                            Select Case TempPlayer(I).CurLanguage
                                                                            Case LANG_PT: AddAlert I, "Your pokemon got burned", White
                                                                            Case LANG_EN: AddAlert I, "Your pokemon got burned", White
                                                                            Case LANG_ES: AddAlert I, "Your pokemon got burned", White
                                                                            End Select
                                                                        Case StatusEnum.Paralize
                                                                            Select Case TempPlayer(I).CurLanguage
                                                                            Case LANG_PT: AddAlert I, "Your pokemon got paralized", White
                                                                            Case LANG_EN: AddAlert I, "Your pokemon got paralized", White
                                                                            Case LANG_ES: AddAlert I, "Your pokemon got paralized", White
                                                                            End Select
                                                                        Case StatusEnum.Sleep
                                                                            Select Case TempPlayer(I).CurLanguage
                                                                            Case LANG_PT: AddAlert I, "Your pokemon fell asleep", White
                                                                            Case LANG_EN: AddAlert I, "Your pokemon fell asleep", White
                                                                            Case LANG_ES: AddAlert I, "Your pokemon fell asleep", White
                                                                            End Select
                                                                        Case StatusEnum.Frozen
                                                                            Select Case TempPlayer(I).CurLanguage
                                                                            Case LANG_PT: AddAlert I, "Your pokemon got frozed", White
                                                                            Case LANG_EN: AddAlert I, "Your pokemon got frozed", White
                                                                            Case LANG_ES: AddAlert I, "Your pokemon got frozed", White
                                                                            End Select
                                                                        End Select
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                    '//Check Move
                                    Select Case PokemonMove(MoveNum).AttackType
                                    Case 1    '//Damage
                                        If Not I = Index Then
                                            If (TempPlayer(Index).InDuel = I And TempPlayer(I).DuelTime <= 0) Or Map(MapNum).Moral = MAP_MORAL_PVP Or (Player(I, TempPlayer(I).UseChar).Access >= ACCESS_CREATOR) Or (Player(Index, TempPlayer(Index).UseChar).Access >= ACCESS_CREATOR) Then
                                                '//Target and Do Damage
                                                targetType = Pokemon(PlayerPokemon(I).Num).PrimaryType
                                                targetType2 = Pokemon(PlayerPokemon(I).Num).SecondaryType
                                                Select Case PokemonMove(MoveNum).Category
                                                Case MoveCategory.Physical
                                                    DefStat = GetPlayerPokemonStat(I, Def)
                                                Case MoveCategory.Special
                                                    DefStat = GetPlayerPokemonStat(I, SpDef)
                                                End Select
                                                Damage = GetPokemonDamage(ownType, pType, targetType, targetType2, ownLevel, AtkStat, Power, DefStat)
                                                '//Check Critical
                                                If PlayerPokemon(Index).NextCritical = YES Then
                                                    Damage = Damage * 2
                                                    SendActionMsg MapNum, "Critical", PlayerPokemon(Index).X * 32, PlayerPokemon(Index).Y * 32, Yellow
                                                    PlayerPokemon(Index).NextCritical = NO
                                                End If
                                                If PokemonMove(MoveNum).BoostWeather > 0 Then
                                                    If PokemonMove(MoveNum).BoostWeather = Map(MapNum).CurWeather Then
                                                        Damage = Damage * 2
                                                    End If
                                                End If
                                                If PokemonMove(MoveNum).DecreaseWeather > 0 Then
                                                    If PokemonMove(MoveNum).DecreaseWeather = Map(MapNum).CurWeather Then
                                                        Damage = Damage / 2
                                                    End If
                                                End If

                                                If Damage > 0 Then
                                                    '//Check Reflect
                                                    If PlayerPokemon(I).ReflectMove = PokemonMove(MoveNum).Category Then
                                                        If PlayerPokemon(I).ReflectMove > 0 Then
                                                            If PlayerPokemon(Index).slot > 0 Then
                                                                PlayerPokemon(I).ReflectMove = 0
                                                                SendActionMsg MapNum, "Reflected", PlayerPokemon(I).X * 32, PlayerPokemon(I).Y * 32, White

                                                                PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp = PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp - Damage
                                                                SendActionMsg MapNum, "-" & Damage, PlayerPokemon(Index).X * 32, PlayerPokemon(Index).Y * 32, BrightGreen
                                                                If PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp <= 0 Then
                                                                    PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp = 0
                                                                    SendPlayerPokemonVital Index
                                                                    SendPlayerPokemonFaint Index
                                                                Else
                                                                    SendPlayerPokemonVital Index
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        PlayerAttackPlayer Index, I, Damage

                                                        '//Absorb
                                                        If PokemonMove(MoveNum).AbsorbDamage > 0 Then
                                                            Absorbed = Damage * (PokemonMove(MoveNum).AbsorbDamage / 100)
                                                            If Absorbed > 0 Then
                                                                PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp = PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp + Absorbed
                                                                SendActionMsg MapNum, "+" & Absorbed, PlayerPokemon(Index).X * 32, PlayerPokemon(Index).Y * 32, BrightGreen
                                                                If PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp >= PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).MaxHp Then
                                                                    PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp = PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).MaxHp
                                                                End If
                                                                SendPlayerPokemonVital Index
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Case 2    '//Buff/Debuff
                                        If Not I = Index Then
                                            If (TempPlayer(Index).InDuel = I And TempPlayer(I).DuelTime <= 0) Or Map(MapNum).Moral = MAP_MORAL_PVP Or (Player(I, TempPlayer(I).UseChar).Access >= ACCESS_CREATOR) Or (Player(Index, TempPlayer(Index).UseChar).Access >= ACCESS_CREATOR) Then
                                                For z = 1 To StatEnum.Stat_Count - 1
                                                    PlayerPokemon(I).StatBuff(z) = PlayerPokemon(I).StatBuff(z) + PokemonMove(MoveNum).dStat(z)
                                                    If PlayerPokemon(I).StatBuff(z) > 6 Then
                                                        PlayerPokemon(I).StatBuff(z) = 6
                                                    ElseIf PlayerPokemon(I).StatBuff(z) < -6 Then
                                                        PlayerPokemon(I).StatBuff(z) = -6
                                                    End If
                                                Next
                                                SendPlayerPokemonStatBuff I
                                            End If
                                        End If
                                    Case 3    '//Heal
                                        If Not TempPlayer(Index).InDuel = I And Not Map(MapNum).Moral = MAP_MORAL_PVP Then
                                            If PlayerPokemon(I).slot > 0 Then
                                                HealAmount = PlayerPokemons(I).Data(PlayerPokemon(I).slot).MaxHp * (PokemonMove(MoveNum).Power / 100)
                                                PlayerPokemons(I).Data(PlayerPokemon(I).slot).CurHp = PlayerPokemons(Index).Data(PlayerPokemon(I).slot).CurHp + HealAmount
                                                SendActionMsg MapNum, "+" & HealAmount, PlayerPokemon(I).X * 32, PlayerPokemon(I).Y * 32, BrightGreen
                                                If PlayerPokemons(I).Data(PlayerPokemon(I).slot).CurHp >= PlayerPokemons(I).Data(PlayerPokemon(I).slot).MaxHp Then
                                                    PlayerPokemons(I).Data(PlayerPokemon(I).slot).CurHp = PlayerPokemons(I).Data(PlayerPokemon(I).slot).MaxHp
                                                End If
                                                SendPlayerPokemonVital I
                                                If PokemonMove(MoveNum).pStatus = 7 Then
                                                    '//Cure Status
                                                    PlayerPokemons(I).Data(PlayerPokemon(I).slot).Status = 0
                                                    Select Case TempPlayer(I).CurLanguage
                                                    Case LANG_PT: AddAlert I, "Your pokemon got cured", White
                                                    Case LANG_EN: AddAlert I, "Your pokemon got cured", White
                                                    Case LANG_ES: AddAlert I, "Your pokemon got cured", White
                                                    End Select
                                                    SendPlayerPokemonStatus I
                                                End If
                                            End If

                                            InRange2 = False
                                            If PokemonMove(MoveNum).targetType = 1 Then    '//AoE
                                                If IsOnAoERange(Range, X, Y, Player(I, TempPlayer(I).UseChar).X, Player(I, TempPlayer(I).UseChar).Y) Then InRange2 = True
                                            ElseIf PokemonMove(MoveNum).targetType = 2 Then    '//Linear
                                                If IsOnLinearRange(PlayerPokemon(Index).Dir, Range, X, Y, Player(I, TempPlayer(I).UseChar).X, Player(I, TempPlayer(I).UseChar).Y) Then InRange2 = True
                                            ElseIf PokemonMove(MoveNum).targetType = 3 Then    '//Spray
                                                If IsOnSprayRange(PlayerPokemon(Index).Dir, Range, X, Y, Player(I, TempPlayer(I).UseChar).X, Player(I, TempPlayer(I).UseChar).Y) Then InRange2 = True
                                            Else
                                                InRange2 = False
                                            End If

                                            If InRange2 Then
                                                Select Case PokemonMove(MoveNum).AttackType
                                                Case 3    '//Heal
                                                    HealAmount = GetPlayerHP(Player(I, TempPlayer(I).UseChar).Level) * (PokemonMove(MoveNum).Power / 100)
                                                    Player(I, TempPlayer(I).UseChar).CurHp = Player(I, TempPlayer(I).UseChar).CurHp + HealAmount
                                                    SendActionMsg MapNum, "+" & HealAmount, Player(I, TempPlayer(I).UseChar).X * 32, Player(I, TempPlayer(I).UseChar).Y * 32, BrightGreen
                                                    If Player(I, TempPlayer(I).UseChar).CurHp >= GetPlayerHP(Player(I, TempPlayer(I).UseChar).Level) Then
                                                        Player(I, TempPlayer(I).UseChar).CurHp = GetPlayerHP(Player(I, TempPlayer(I).UseChar).Level)
                                                    End If
                                                    SendPlayerVital I
                                                    If PokemonMove(MoveNum).pStatus = 7 Then
                                                        '//Cure Status
                                                        Player(I, TempPlayer(I).UseChar).Status = 0
                                                        Select Case TempPlayer(I).CurLanguage
                                                        Case LANG_PT: AddAlert I, "You got cured", White
                                                        Case LANG_EN: AddAlert I, "You got cured", White
                                                        Case LANG_ES: AddAlert I, "You got cured", White
                                                        End Select
                                                        SendPlayerStatus I
                                                    End If
                                                End Select
                                            End If
                                        End If
                                    End Select
                                End If
                            End If
                        ElseIf Map(MapNum).KillPlayer = YES Then
                            'Adicionado a um método, pra ser usado juntamente com o PVP
                            Call AttackPlayer(I, MoveNum, PlayerPokemon(Index).Num, X, Y, PlayerPokemon(Index).Dir, pType, ownType, ownLevel, AtkStat, Power, PlayerPokemon(Index).NextCritical, PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp, PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).MaxHp)
                        Else

                            If Not TempPlayer(Index).InDuel = I Then
                                InRange = False
                                If PokemonMove(MoveNum).targetType = 1 Then    '//AoE
                                    If IsOnAoERange(Range, X, Y, Player(I, TempPlayer(I).UseChar).X, Player(I, TempPlayer(I).UseChar).Y) Then InRange = True
                                ElseIf PokemonMove(MoveNum).targetType = 2 Then    '//Linear
                                    If IsOnLinearRange(PlayerPokemon(Index).Dir, Range, X, Y, Player(I, TempPlayer(I).UseChar).X, Player(I, TempPlayer(I).UseChar).Y) Then InRange = True
                                ElseIf PokemonMove(MoveNum).targetType = 3 Then    '//Spray
                                    If IsOnSprayRange(PlayerPokemon(Index).Dir, Range, X, Y, Player(I, TempPlayer(I).UseChar).X, Player(I, TempPlayer(I).UseChar).Y) Then InRange = True
                                Else
                                    InRange = False
                                End If

                                If InRange Then
                                    Select Case PokemonMove(MoveNum).AttackType
                                    Case 3    '//Heal
                                        HealAmount = GetPlayerHP(Player(I, TempPlayer(I).UseChar).Level) * (PokemonMove(MoveNum).Power / 100)
                                        Player(I, TempPlayer(I).UseChar).CurHp = Player(I, TempPlayer(I).UseChar).CurHp + HealAmount
                                        SendActionMsg MapNum, "+" & HealAmount, Player(I, TempPlayer(I).UseChar).X * 32, Player(I, TempPlayer(I).UseChar).Y * 32, BrightGreen
                                        If Player(I, TempPlayer(I).UseChar).CurHp >= GetPlayerHP(Player(I, TempPlayer(I).UseChar).Level) Then
                                            Player(I, TempPlayer(I).UseChar).CurHp = GetPlayerHP(Player(I, TempPlayer(I).UseChar).Level)
                                        End If
                                        SendPlayerVital I
                                        If PokemonMove(MoveNum).pStatus = 7 Then
                                            '//Cure Status
                                            Player(I, TempPlayer(I).UseChar).Status = 0
                                            Select Case TempPlayer(I).CurLanguage
                                            Case LANG_PT: AddAlert I, "You got cured", White
                                            Case LANG_EN: AddAlert I, "You got cured", White
                                            Case LANG_ES: AddAlert I, "You got cured", White
                                            End Select
                                            SendPlayerStatus I
                                        End If
                                    End Select
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
        '//Npc
        For I = 1 To Pokemon_HighIndex
            If MapPokemon(I).Num > 0 Then
                If MapPokemon(I).Map = MapNum Then
                    '//Check Status Req
                    If PokemonMove(MoveNum).StatusReq > 0 Then
                        If MapPokemon(I).Status = PokemonMove(MoveNum).StatusReq Then
                            CanAttack = True
                        Else
                            CanAttack = False
                        End If
                    Else
                        CanAttack = True
                    End If
                    If MapPokemon(I).IsProtect > 0 Then
                        CanAttack = False
                        MapPokemon(I).IsProtect = NO
                        SendActionMsg MapNum, "Protected", MapPokemon(I).X * 32, MapPokemon(I).Y * 32, Yellow
                    End If

                    If CanAttack Then
                        '//Check Location
                        '//ToDo: Must be in PvP map
                        InRange = False
                        If PokemonMove(MoveNum).targetType = 1 Then    '//AoE
                            If IsOnAoERange(Range, X, Y, MapPokemon(I).X, MapPokemon(I).Y) Then InRange = True
                        ElseIf PokemonMove(MoveNum).targetType = 2 Then    '//Linear
                            If IsOnLinearRange(PlayerPokemon(Index).Dir, Range, X, Y, MapPokemon(I).X, MapPokemon(I).Y) Then InRange = True
                        ElseIf PokemonMove(MoveNum).targetType = 3 Then    '//Spray
                            If IsOnSprayRange(PlayerPokemon(Index).Dir, Range, X, Y, MapPokemon(I).X, MapPokemon(I).Y) Then InRange = True
                        Else
                            InRange = False
                        End If

                        If InRange Then
                            '//Status
                            'If Spawn(i).pokeBuff <= 5 Then
                            If PokemonMove(MoveNum).pStatus > 0 And PokemonMove(MoveNum).StatusToSelf = NO Then
                                If PokemonMove(MoveNum).pStatus = 6 Then
                                    MapPokemon(I).IsConfuse = YES
                                Else
                                    If MapPokemon(I).Status <= 0 Then
                                        statusChance = (100 * (PokemonMove(MoveNum).pStatusChance / 100))

                                        If IsImmuneOnStatus(PokemonMove(MoveNum).Type, Pokemon(MapPokemon(I).Num).PrimaryType, Pokemon(MapPokemon(I).Num).SecondaryType, PokemonMove(MoveNum).pStatus) Then
                                            If statusChance > 0 Then
                                                statusRand = Random(1, 100)
                                                If statusRand <= statusChance Then
                                                    MapPokemon(I).Status = PokemonMove(MoveNum).pStatus
                                                    SendMapPokemonStatus I
                                                End If
                                                MapPokemon(I).targetType = TARGET_TYPE_PLAYER
                                                MapPokemon(I).TargetIndex = Index
                                                MapPokemon(I).LastAttacker = Index
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            'End If
                            '//Check Move
                            Select Case PokemonMove(MoveNum).AttackType
                            Case 1    '//Damage
                                '//Target and Do Damage
                                targetType = Pokemon(MapPokemon(I).Num).PrimaryType
                                targetType2 = Pokemon(MapPokemon(I).Num).SecondaryType
                                Select Case PokemonMove(MoveNum).Category
                                Case MoveCategory.Physical
                                    DefStat = GetNpcPokemonStat(I, Def)
                                Case MoveCategory.Special
                                    DefStat = GetNpcPokemonStat(I, SpDef)
                                End Select
                                Damage = GetPokemonDamage(ownType, pType, targetType, targetType2, ownLevel, AtkStat, Power, DefStat)
                                '//Check Critical
                                If PlayerPokemon(Index).NextCritical = YES Then
                                    Damage = Damage * 2
                                    SendActionMsg MapNum, "Critical", PlayerPokemon(Index).X * 32, PlayerPokemon(Index).Y * 32, Yellow
                                    PlayerPokemon(Index).NextCritical = NO
                                End If
                                If PokemonMove(MoveNum).BoostWeather > 0 Then
                                    If PokemonMove(MoveNum).BoostWeather = Map(MapNum).CurWeather Then
                                        Damage = Damage * 2
                                    End If
                                End If
                                If PokemonMove(MoveNum).DecreaseWeather > 0 Then
                                    If PokemonMove(MoveNum).DecreaseWeather = Map(MapNum).CurWeather Then
                                        Damage = Damage / 2
                                    End If
                                End If

                                If Damage > 0 Then
                                    '//Check Reflect
                                    If MapPokemon(I).ReflectMove = PokemonMove(MoveNum).Category Then
                                        If MapPokemon(I).ReflectMove > 0 Then
                                            If PlayerPokemon(Index).slot > 0 Then
                                                MapPokemon(I).ReflectMove = 0
                                                SendActionMsg MapNum, "Reflected", MapPokemon(I).X * 32, MapPokemon(I).Y * 32, White

                                                PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp = PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp - Damage
                                                SendActionMsg MapNum, "-" & Damage, PlayerPokemon(Index).X * 32, PlayerPokemon(Index).Y * 32, BrightGreen
                                                If PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp <= 0 Then
                                                    PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp = 0
                                                    SendPlayerPokemonVital Index
                                                    SendPlayerPokemonFaint Index
                                                Else
                                                    SendPlayerPokemonVital Index
                                                End If
                                            End If
                                        End If
                                    Else
                                        'setBuff = Spawn(i).pokeBuff
                                        'If setBuff > 0 Then
                                        '    Damage = Damage / setBuff
                                        'End If

                                        PlayerAttackNpc Index, I, Damage
                                        '//Absorb
                                        If PokemonMove(MoveNum).AbsorbDamage > 0 Then
                                            Absorbed = Damage * (PokemonMove(MoveNum).AbsorbDamage / 100)
                                            If Absorbed > 0 Then
                                                PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp = PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp + Absorbed
                                                SendActionMsg MapNum, "+" & Absorbed, PlayerPokemon(Index).X * 32, PlayerPokemon(Index).Y * 32, BrightGreen
                                                If PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp >= PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).MaxHp Then
                                                    PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp = PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).MaxHp
                                                End If
                                                SendPlayerPokemonVital Index
                                            End If
                                        End If
                                    End If
                                End If
                            Case 2    '//Buff/Debuff
                                For z = 1 To StatEnum.Stat_Count - 1
                                    MapPokemon(I).StatBuff(z) = MapPokemon(I).StatBuff(z) + PokemonMove(MoveNum).dStat(z)
                                    If MapPokemon(I).StatBuff(z) > 6 Then
                                        MapPokemon(I).StatBuff(z) = 6
                                    ElseIf MapPokemon(I).StatBuff(z) < -6 Then
                                        MapPokemon(I).StatBuff(z) = -6
                                    End If
                                Next
                                If MapPokemon(I).targetType <= 0 Then
                                    MapPokemon(I).targetType = TARGET_TYPE_PLAYER
                                    MapPokemon(I).TargetIndex = Index
                                    MapPokemon(I).LastAttacker = Index
                                End If
                                '//ToDo: Update stat to client
                            Case 3    '//Heal

                            End Select
                        End If
                    End If
                End If
            End If
        Next
        For I = 1 To MAX_MAP_NPC
            If MapNpc(MapNum, I).Num > 0 Then
                If MapNpcPokemon(MapNum, I).Num > 0 Then
                    If MapNpc(MapNum, I).InBattle = Index Then
                        '//Check Status Req
                        If PokemonMove(MoveNum).StatusReq > 0 Then
                            If MapNpcPokemon(MapNum, I).Status = PokemonMove(MoveNum).StatusReq Then
                                CanAttack = True
                            Else
                                CanAttack = False
                            End If
                        Else
                            CanAttack = True
                        End If
                        If MapNpcPokemon(MapNum, I).IsProtect > 0 Then
                            CanAttack = False
                            MapNpcPokemon(MapNum, I).IsProtect = NO
                            SendActionMsg MapNum, "Protected", MapNpcPokemon(MapNum, I).X * 32, MapNpcPokemon(MapNum, I).Y * 32, Yellow
                        End If

                        If CanAttack Then
                            InRange = False
                            If PokemonMove(MoveNum).targetType = 1 Then    '//AoE
                                If IsOnAoERange(Range, X, Y, MapNpcPokemon(MapNum, I).X, MapNpcPokemon(MapNum, I).Y) Then InRange = True
                            ElseIf PokemonMove(MoveNum).targetType = 2 Then    '//Linear
                                If IsOnLinearRange(PlayerPokemon(Index).Dir, Range, X, Y, MapNpcPokemon(MapNum, I).X, MapNpcPokemon(MapNum, I).Y) Then InRange = True
                            ElseIf PokemonMove(MoveNum).targetType = 3 Then    '//Spray
                                If IsOnSprayRange(PlayerPokemon(Index).Dir, Range, X, Y, MapNpcPokemon(MapNum, I).X, MapNpcPokemon(MapNum, I).Y) Then InRange = True
                            Else
                                InRange = False
                            End If

                            If InRange Then
                                '//Status
                                If PokemonMove(MoveNum).pStatus > 0 And PokemonMove(MoveNum).StatusToSelf = NO Then
                                    If PokemonMove(MoveNum).pStatus = 6 Then
                                        MapNpcPokemon(MapNum, I).IsConfuse = YES
                                    Else
                                        If MapNpcPokemon(MapNum, I).Status <= 0 Then
                                            statusChance = (100 * (PokemonMove(MoveNum).pStatusChance / 100))

                                            If IsImmuneOnStatus(PokemonMove(MoveNum).Type, Pokemon(MapNpcPokemon(MapNum, I).Num).PrimaryType, Pokemon(MapNpcPokemon(MapNum, I).Num).SecondaryType, PokemonMove(MoveNum).pStatus) Then
                                                If statusChance > 0 Then
                                                    statusRand = Random(1, 100)
                                                    If statusRand <= statusChance Then
                                                        MapNpcPokemon(MapNum, I).Status = PokemonMove(MoveNum).pStatus
                                                        SendMapNpcPokemonStatus MapNum, I
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                '//Check Move
                                Select Case PokemonMove(MoveNum).AttackType
                                Case 1    '//Damage
                                    '//Target and Do Damage
                                    targetType = Pokemon(MapNpcPokemon(MapNum, I).Num).PrimaryType
                                    targetType2 = Pokemon(MapNpcPokemon(MapNum, I).Num).SecondaryType
                                    Select Case PokemonMove(MoveNum).Category
                                    Case MoveCategory.Physical
                                        DefStat = GetMapNpcPokemonStat(MapNum, I, Def)
                                    Case MoveCategory.Special
                                        DefStat = GetMapNpcPokemonStat(MapNum, I, SpDef)
                                    End Select
                                    Damage = GetPokemonDamage(ownType, pType, targetType, targetType2, ownLevel, AtkStat, Power, DefStat)
                                    '//Check Critical
                                    If PlayerPokemon(Index).NextCritical = YES Then
                                        Damage = Damage * 2
                                        SendActionMsg MapNum, "Critical", PlayerPokemon(Index).X * 32, PlayerPokemon(Index).Y * 32, Yellow
                                        PlayerPokemon(Index).NextCritical = NO
                                    End If
                                    If PokemonMove(MoveNum).BoostWeather > 0 Then
                                        If PokemonMove(MoveNum).BoostWeather = Map(MapNum).CurWeather Then
                                            Damage = Damage * 2
                                        End If
                                    End If
                                    If PokemonMove(MoveNum).DecreaseWeather > 0 Then
                                        If PokemonMove(MoveNum).DecreaseWeather = Map(MapNum).CurWeather Then
                                            Damage = Damage / 2
                                        End If
                                    End If

                                    If Damage > 0 Then
                                        If TempPlayer(Index).DuelTime <= 0 Then
                                            '//Check Reflect
                                            If MapNpcPokemon(MapNum, I).ReflectMove = PokemonMove(MoveNum).Category Then
                                                If MapNpcPokemon(MapNum, I).ReflectMove > 0 Then
                                                    If PlayerPokemon(Index).slot > 0 Then
                                                        MapNpcPokemon(MapNum, I).ReflectMove = 0
                                                        SendActionMsg MapNum, "Reflected", MapNpcPokemon(MapNum, I).X * 32, MapNpcPokemon(MapNum, I).Y * 32, White

                                                        PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp = PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp - Damage
                                                        SendActionMsg MapNum, "-" & Damage, PlayerPokemon(Index).X * 32, PlayerPokemon(Index).Y * 32, BrightGreen
                                                        If PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp <= 0 Then
                                                            PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp = 0
                                                            SendPlayerPokemonVital Index
                                                            SendPlayerPokemonFaint Index
                                                        Else
                                                            SendPlayerPokemonVital Index
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                PlayerAttackNpcPokemon Index, I, Damage
                                                '//Absorb
                                                If PokemonMove(MoveNum).AbsorbDamage > 0 Then
                                                    Absorbed = Damage * (PokemonMove(MoveNum).AbsorbDamage / 100)
                                                    If Absorbed > 0 Then
                                                        PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp = PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp + Absorbed
                                                        SendActionMsg MapNum, "+" & Absorbed, PlayerPokemon(Index).X * 32, PlayerPokemon(Index).Y * 32, BrightGreen
                                                        If PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp >= PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).MaxHp Then
                                                            PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp = PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).MaxHp
                                                        End If
                                                        SendPlayerPokemonVital Index
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Case 2    '//Buff/Debuff
                                    For z = 1 To StatEnum.Stat_Count - 1
                                        MapNpcPokemon(MapNum, I).StatBuff(z) = MapNpcPokemon(MapNum, I).StatBuff(z) + PokemonMove(MoveNum).dStat(z)
                                        If MapNpcPokemon(MapNum, I).StatBuff(z) > 6 Then
                                            MapNpcPokemon(MapNum, I).StatBuff(z) = 6
                                        ElseIf MapNpcPokemon(MapNum, I).StatBuff(z) < -6 Then
                                            MapNpcPokemon(MapNum, I).StatBuff(z) = -6
                                        End If
                                    Next
                                    '//ToDo: Update stat to client
                                Case 3    '//Heal

                                End Select
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End Select

    '//Change Weather
    If PokemonMove(MoveNum).ChangeWeather > 0 Then
        If PokemonMove(MoveNum).ChangeWeather = WeatherEnum.Count_Weather Then
            '//Clear
            Map(MapNum).CurWeather = Map(MapNum).StartWeather
        Else
            Map(MapNum).CurWeather = PokemonMove(MoveNum).ChangeWeather
        End If
        SendWeather MapNum
    End If

    '//Play Animation
    If PokemonMove(MoveNum).Animation > 0 Then
        If PokemonMove(MoveNum).SelfAnim = YES Then
            SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, PlayerPokemon(Index).X, PlayerPokemon(Index).Y
        Else
            '//Check Target Type
            Select Case PokemonMove(MoveNum).targetType
            Case 0    '//Self
                SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, PlayerPokemon(Index).X, PlayerPokemon(Index).Y
            Case 1    '//AoE
                If Range > 0 Then
                    For X = PlayerPokemon(Index).X - Range To PlayerPokemon(Index).X + Range
                        For Y = PlayerPokemon(Index).Y - Range To PlayerPokemon(Index).Y + Range
                            If isValidMapPoint(MapNum, X, Y) Then
                                If IsOnAoERange(Range, PlayerPokemon(Index).X, PlayerPokemon(Index).Y, X, Y) Then
                                    SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                                End If
                            End If
                        Next
                    Next
                End If
            Case 2    '//Linear
                If Range > 0 Then
                    Select Case PlayerPokemon(Index).Dir
                    Case DIR_UP
                        For Y = PlayerPokemon(Index).Y - Range To PlayerPokemon(Index).Y - 1
                            X = PlayerPokemon(Index).X
                            SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                        Next
                    Case DIR_DOWN
                        For Y = PlayerPokemon(Index).Y + 1 To PlayerPokemon(Index).Y + Range
                            X = PlayerPokemon(Index).X
                            SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                        Next
                    Case DIR_LEFT
                        For X = PlayerPokemon(Index).X - Range To PlayerPokemon(Index).X - 1
                            Y = PlayerPokemon(Index).Y
                            SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                        Next
                    Case DIR_RIGHT
                        For X = PlayerPokemon(Index).X + 1 To PlayerPokemon(Index).X + Range
                            Y = PlayerPokemon(Index).Y
                            SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                        Next
                    End Select
                End If
            Case 3    '//Spray
                If Range > 0 Then
                    z = 1
                    Select Case PlayerPokemon(Index).Dir
                    Case DIR_UP
                        For Y = PlayerPokemon(Index).Y - 1 To PlayerPokemon(Index).Y - Range Step -1
                            For X = PlayerPokemon(Index).X - z To PlayerPokemon(Index).X + z
                                SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                            Next
                            z = z + 1
                        Next
                    Case DIR_DOWN
                        For Y = PlayerPokemon(Index).Y + 1 To PlayerPokemon(Index).Y + Range
                            For X = PlayerPokemon(Index).X - z To PlayerPokemon(Index).X + z
                                SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                            Next
                            z = z + 1
                        Next
                    Case DIR_LEFT
                        For X = PlayerPokemon(Index).X - 1 To PlayerPokemon(Index).X - Range Step -1
                            For Y = PlayerPokemon(Index).Y - z To PlayerPokemon(Index).Y + z
                                SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                            Next
                            z = z + 1
                        Next
                    Case DIR_RIGHT
                        For X = PlayerPokemon(Index).X + 1 To PlayerPokemon(Index).X + Range
                            For Y = PlayerPokemon(Index).Y - z To PlayerPokemon(Index).Y + z
                                SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                            Next
                            z = z + 1
                        Next
                    End Select
                End If
            End Select
        End If
    End If

    '//Status
    If PokemonMove(MoveNum).pStatus > 0 And PokemonMove(MoveNum).StatusToSelf = YES Then
        If PokemonMove(MoveNum).pStatus = 6 Then
            PlayerPokemon(Index).IsConfuse = YES
            SendPlayerPokemonStatus Index
            Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: AddAlert Index, "Your pokemon got confused", White
            Case LANG_EN: AddAlert Index, "Your pokemon got confused", White
            Case LANG_ES: AddAlert Index, "Your pokemon got confused", White
            End Select
        Else
            If PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Status <= 0 Then
                statusChance = (100 * (PokemonMove(MoveNum).pStatusChance / 100))

                If IsImmuneOnStatus(PokemonMove(MoveNum).Type, Pokemon(PlayerPokemon(Index).Num).PrimaryType, Pokemon(PlayerPokemon(Index).Num).SecondaryType, PokemonMove(MoveNum).pStatus) Then
                    If statusChance > 0 Then
                        statusRand = Random(1, 100)
                        If statusRand <= statusChance Then
                            PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Status = PokemonMove(MoveNum).pStatus
                            SendPlayerPokemonStatus Index
                            Select Case PokemonMove(MoveNum).pStatus
                            Case StatusEnum.Poison
                                Select Case TempPlayer(Index).CurLanguage
                                Case LANG_PT: AddAlert Index, "Your pokemon got poisoned", White
                                Case LANG_EN: AddAlert Index, "Your pokemon got poisoned", White
                                Case LANG_ES: AddAlert Index, "Your pokemon got poisoned", White
                                End Select
                            Case StatusEnum.Burn
                                Select Case TempPlayer(Index).CurLanguage
                                Case LANG_PT: AddAlert Index, "Your pokemon got burned", White
                                Case LANG_EN: AddAlert Index, "Your pokemon got burned", White
                                Case LANG_ES: AddAlert Index, "Your pokemon got burned", White
                                End Select
                            Case StatusEnum.Paralize
                                Select Case TempPlayer(Index).CurLanguage
                                Case LANG_PT: AddAlert Index, "Your pokemon got paralized", White
                                Case LANG_EN: AddAlert Index, "Your pokemon got paralized", White
                                Case LANG_ES: AddAlert Index, "Your pokemon got paralized", White
                                End Select
                            Case StatusEnum.Sleep
                                Select Case TempPlayer(Index).CurLanguage
                                Case LANG_PT: AddAlert Index, "Your pokemon fell asleep", White
                                Case LANG_EN: AddAlert Index, "Your pokemon fell asleep", White
                                Case LANG_ES: AddAlert Index, "Your pokemon fell asleep", White
                                End Select
                            Case StatusEnum.Frozen
                                Select Case TempPlayer(Index).CurLanguage
                                Case LANG_PT: AddAlert Index, "Your pokemon got frozed", White
                                Case LANG_EN: AddAlert Index, "Your pokemon got frozed", White
                                Case LANG_ES: AddAlert Index, "Your pokemon got frozed", White
                                End Select
                            End Select
                        End If
                    End If
                End If
            End If
        End If
    End If

    If PokemonMove(MoveNum).RecoilDamage > 0 Then
        recoil = PokemonMove(MoveNum).RecoilDamage
        Damage = PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).MaxHp * (recoil / 100)
        If Damage > 0 Then
            PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp = PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp - Damage
            SendActionMsg MapNum, "-" & Damage, PlayerPokemon(Index).X * 32, PlayerPokemon(Index).Y * 32, BrightRed
            If PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp <= 0 Then
                PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp = 0
                SendPlayerPokemonVital Index
                SendPlayerPokemonFaint Index
            Else
                SendPlayerPokemonVital Index
            End If
        End If
    End If

    '//Play Sound
    If Not Trim$(PokemonMove(MoveNum).Sound) = "None." Or Not Trim$(PokemonMove(MoveNum).Sound) = vbNullString Then
        SendPlaySound Trim$(PokemonMove(MoveNum).Sound), MapNum
    End If
End Sub

Public Sub NpcCastMove(ByVal MapPokemonNum As Long, ByVal MoveNum As Long, ByVal MoveSlot As Byte)
Dim RandomNum As Byte
Dim expEarn As Long
Dim I As Byte, pCount As Byte

    '//Check for error
    If MapPokemonNum <= 0 Or MapPokemonNum > MAX_GAME_POKEMON Then Exit Sub
    If MapPokemon(MapPokemonNum).Num <= 0 Then Exit Sub
    If MoveNum <= 0 Or MoveNum > MAX_POKEMON_MOVE Then Exit Sub
    
    '//Add Queue
    With MapPokemon(MapPokemonNum)
        If Not PokemonMove(MoveNum).SelfStatusReq = StatusEnum.Sleep Then
            If .Status = StatusEnum.Sleep Then
                RandomNum = Random(1, 5)
                If RandomNum = 1 Then
                    '//Remove Status
                    .Status = 0
                    SendMapPokemonStatus MapPokemonNum
                Else
                    Exit Sub
                End If
            End If
        End If
        If Not PokemonMove(MoveNum).SelfStatusReq = StatusEnum.Frozen Then
            If .Status = StatusEnum.Frozen Then
                RandomNum = Random(1, 5)
                If RandomNum = 1 Then
                    '//Remove Status
                    .Status = 0
                    SendMapPokemonStatus MapPokemonNum
                Else
                    Exit Sub
                End If
            End If
        End If
        
        If PokemonMove(MoveNum).SelfStatusReq > 0 Then
            If Not .Status = PokemonMove(MoveNum).SelfStatusReq Then
                Exit Sub
            End If
        End If
        
        '//Burn
        If .Status = StatusEnum.Burn Then
            If .StatusDamage > 0 Then
                If .StatusDamage >= .CurHp Then
                    .CurHp = 0
                    SendActionMsg .Map, "-" & .StatusDamage, .X * 32, .Y * 32, BrightRed
        
                    DefeatMapPokemon MapPokemonNum
                        
                    ClearMapPokemon MapPokemonNum
                    Exit Sub
                Else
                    .CurHp = .CurHp - .StatusDamage
                    SendActionMsg .Map, "-" & .StatusDamage, .X * 32, .Y * 32, BrightRed
                    '//Update
                    SendPokemonVital MapPokemonNum
                End If
            Else
                .StatusDamage = (.MaxHp / 8)
            End If
        End If
        
        .QueueMove = MoveNum
        .QueueMoveSlot = MoveSlot
        
        '//Set Duration
        .MoveCastTime = GetTickCount + (PokemonMove(MoveNum).CastTime)
        .MoveDuration = GetTickCount + (PokemonMove(MoveNum).Duration)
        .MoveInterval = GetTickCount
        .MoveAttackCount = 0
        
        '//Decrease PP
        '//Note: NPC Have Unlimited PP
        
        '//Add ActionMsg
        SendActionMsg .Map, Trim$(PokemonMove(MoveNum).Name), .X * 32, .Y * 32, Yellow
    End With
    
    '//Add ActionMsg
    'SendActionMsg Player(Index, TempPlayer(Index).UseChar).Map, Trim$(PokemonMove(MoveNum).Name), Player(Index, TempPlayer(Index).UseChar).x * 32, Player(Index, TempPlayer(Index).UseChar).y * 32, White
End Sub

Public Sub ProcessNpcMove(ByVal MapPokemonNum As Long, ByVal MoveNum As Long)
Dim I As Long
Dim Range As Long
Dim X As Long, Y As Long
Dim pType As Byte
Dim ownType As Byte
Dim ownLevel As Byte
Dim AtkStat As Long
Dim Power As Long
Dim MapNum As Long
Dim pi As Byte, pCount As Byte

'//Damage Calculation
Dim Damage As Long
Dim targetType As Byte, targetType2 As Byte
Dim DefStat As Long

Dim InRange As Boolean
Dim z As Byte
Dim HealAmount As Long
Dim statusChance As Long
Dim statusRand As Long
Dim recoil As Long
Dim expEarn As Long
Dim Absorbed As Long
Dim CanAttack As Boolean
Dim setBuff As Long

    '//Check for error
    If MapPokemonNum <= 0 Or MapPokemonNum > MAX_GAME_POKEMON Then Exit Sub
    If MapPokemon(MapPokemonNum).Num <= 0 Then Exit Sub
    
    Range = PokemonMove(MoveNum).Range
    Power = PokemonMove(MoveNum).Power
    X = MapPokemon(MapPokemonNum).X
    Y = MapPokemon(MapPokemonNum).Y
    ownType = Pokemon(MapPokemon(MapPokemonNum).Num).PrimaryType
    ownLevel = MapPokemon(MapPokemonNum).Level
    pType = PokemonMove(MoveNum).Type
    MapNum = MapPokemon(MapPokemonNum).Map
    
    '//Check Attack Category
    Select Case PokemonMove(MoveNum).Category
        Case MoveCategory.Physical
            AtkStat = GetNpcPokemonStat(I, Atk)
        Case MoveCategory.Special
            AtkStat = GetNpcPokemonStat(I, SpAtk)
    End Select

    Select Case PokemonMove(MoveNum).targetType
        Case 0 '//Self
            Select Case PokemonMove(MoveNum).AttackType
                Case 2 '//Buff/Debuff
                    For z = 1 To StatEnum.Stat_Count - 1
                        MapPokemon(MapPokemonNum).StatBuff(z) = MapPokemon(MapPokemonNum).StatBuff(z) + PokemonMove(MoveNum).dStat(z)
                        If MapPokemon(MapPokemonNum).StatBuff(z) > 6 Then
                            MapPokemon(MapPokemonNum).StatBuff(z) = 6
                        ElseIf MapPokemon(MapPokemonNum).StatBuff(z) < -6 Then
                            MapPokemon(MapPokemonNum).StatBuff(z) = -6
                        End If
                    Next
                    '//ToDo: Update stat to client
                Case 3 '//Heal
                    HealAmount = MapPokemon(MapPokemonNum).MaxHp * (PokemonMove(MoveNum).Power / 100)
                    MapPokemon(MapPokemonNum).CurHp = MapPokemon(MapPokemonNum).CurHp + HealAmount
                    If MapPokemon(MapPokemonNum).CurHp >= MapPokemon(MapPokemonNum).MaxHp Then
                        MapPokemon(MapPokemonNum).CurHp = MapPokemon(MapPokemonNum).MaxHp
                    End If
                    SendPokemonVital MapPokemonNum
            End Select
            '//Status
            If PokemonMove(MoveNum).pStatus > 0 And PokemonMove(MoveNum).StatusToSelf = NO Then
                If PokemonMove(MoveNum).pStatus = 6 Then
                    MapPokemon(MapPokemonNum).IsConfuse = YES
                Else
                    If MapPokemon(MapPokemonNum).Status <= 0 Then
                        statusChance = (100 * (PokemonMove(MoveNum).pStatusChance / 100))
                        
                        If IsImmuneOnStatus(PokemonMove(MoveNum).Type, Pokemon(MapPokemon(MapPokemonNum).Num).PrimaryType, Pokemon(MapPokemon(MapPokemonNum).Num).SecondaryType, PokemonMove(MoveNum).pStatus) Then
                            If statusChance > 0 Then
                                statusRand = Random(1, 100)
                                If statusRand <= statusChance Then
                                    MapPokemon(MapPokemonNum).Status = PokemonMove(MoveNum).pStatus
                                    SendMapPokemonStatus MapPokemonNum
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            '//Reflect
            If PokemonMove(MoveNum).ReflectType > 0 Then
                MapPokemon(MapPokemonNum).ReflectMove = PokemonMove(MoveNum).ReflectType
            End If
            If PokemonMove(MoveNum).CastProtect > 0 Then
                MapPokemon(MapPokemonNum).IsProtect = YES
            End If
        Case 1, 2, 3 '//AoE
            '//Check Target
            '//Player
            For I = 1 To Player_HighIndex
                If IsPlaying(I) Then
                    If TempPlayer(I).UseChar > 0 Then
                        If Player(I, TempPlayer(I).UseChar).Map = MapNum Then
                            If PlayerPokemon(I).Num > 0 Then
                                '//Check Status Req
                                If PokemonMove(MoveNum).StatusReq > 0 Then
                                    If PlayerPokemons(I).Data(PlayerPokemon(I).slot).Status = PokemonMove(MoveNum).StatusReq Then
                                        CanAttack = True
                                    Else
                                        CanAttack = False
                                    End If
                                Else
                                    CanAttack = True
                                End If
                                If PlayerPokemon(I).IsProtect > 0 Then
                                    CanAttack = False
                                    PlayerPokemon(I).IsProtect = NO
                                    SendActionMsg MapNum, "Protected", PlayerPokemon(I).X * 32, PlayerPokemon(I).Y * 32, Yellow
                                End If
                                
                                If CanAttack Then
                                    '//Check Location
                                    InRange = False
                                    If PokemonMove(MoveNum).targetType = 1 Then '//AoE
                                        If IsOnAoERange(Range, X, Y, PlayerPokemon(I).X, PlayerPokemon(I).Y) Then InRange = True
                                    ElseIf PokemonMove(MoveNum).targetType = 2 Then '//Linear
                                        If IsOnLinearRange(MapPokemon(MapPokemonNum).Dir, Range, X, Y, PlayerPokemon(I).X, PlayerPokemon(I).Y) Then InRange = True
                                    ElseIf PokemonMove(MoveNum).targetType = 3 Then '//Spray
                                        If IsOnSprayRange(MapPokemon(MapPokemonNum).Dir, Range, X, Y, PlayerPokemon(I).X, PlayerPokemon(I).Y) Then InRange = True
                                    Else
                                        InRange = False
                                    End If
                                    
                                    If InRange Then
                                        If PlayerPokemon(I).slot > 0 Then
                                            '//Check Move
                                            Select Case PokemonMove(MoveNum).AttackType
                                                Case 1 '//Damage
                                                    '//Get Target
                                                    '//Target and Do Damage
                                                    targetType = Pokemon(PlayerPokemon(I).Num).PrimaryType
                                                    targetType2 = Pokemon(PlayerPokemon(I).Num).SecondaryType
                                                    Select Case PokemonMove(MoveNum).Category
                                                        Case MoveCategory.Physical
                                                            DefStat = GetPlayerPokemonStat(I, Def)
                                                        Case MoveCategory.Special
                                                            DefStat = GetPlayerPokemonStat(I, SpDef)
                                                    End Select
                                                    Damage = GetPokemonDamage(ownType, pType, targetType, targetType2, ownLevel, AtkStat, Power, DefStat)
                                                    If MapPokemon(MapPokemonNum).NextCritical = YES Then
                                                        Damage = Damage * 2
                                                        SendActionMsg MapNum, "Critical", MapPokemon(MapPokemonNum).X * 32, MapPokemon(MapPokemonNum).Y * 32, Yellow
                                                    End If
                                                    If PokemonMove(MoveNum).BoostWeather > 0 Then
                                                        If PokemonMove(MoveNum).BoostWeather = Map(MapNum).CurWeather Then
                                                            Damage = Damage * 2
                                                        End If
                                                    End If
                                                    If PokemonMove(MoveNum).DecreaseWeather > 0 Then
                                                        If PokemonMove(MoveNum).DecreaseWeather = Map(MapNum).CurWeather Then
                                                            Damage = Damage / 2
                                                        End If
                                                    End If
                                                    
                                                    If Damage > 0 Then
                                                        '//Check Reflect
                                                        If PlayerPokemon(I).ReflectMove = PokemonMove(MoveNum).Category Then
                                                            If PlayerPokemon(I).ReflectMove > 0 Then
                                                                If PlayerPokemon(I).slot > 0 Then
                                                                    PlayerPokemon(I).ReflectMove = 0
                                                                    SendActionMsg MapNum, "Reflected", PlayerPokemon(I).X * 32, PlayerPokemon(I).Y * 32, White
    
                                                                    MapPokemon(MapPokemonNum).CurHp = MapPokemon(MapPokemonNum).CurHp - Damage
                                                                    SendActionMsg MapNum, "-" & Damage, MapPokemon(MapPokemonNum).X * 32, MapPokemon(MapPokemonNum).Y * 32, BrightGreen
                                                                    If MapPokemon(MapPokemonNum).CurHp <= 0 Then
                                                                        MapPokemon(MapPokemonNum).CurHp = 0
                                                                        '//Update
                                                                        SendPokemonVital MapPokemonNum
                                                                        
                                                                        DefeatMapPokemon MapPokemonNum

                                                                        ClearMapPokemon MapPokemonNum
                                                                    Else
                                                                        '//Update
                                                                        SendPokemonVital MapPokemonNum
                                                                    End If
                                                                End If
                                                            End If
                                                        Else
                                                            setBuff = Spawn(MapPokemonNum).pokeBuff
                                                            If setBuff > 0 Then
                                                                Damage = Damage * setBuff
                                                            End If
                                                        
                                                            NpcAttackPlayer MapNum, MapPokemonNum, I, Damage
                                                            '//Absorb
                                                            If PokemonMove(MoveNum).AbsorbDamage > 0 Then
                                                                Absorbed = Damage * (PokemonMove(MoveNum).AbsorbDamage / 100)
                                                                If Absorbed > 0 Then
                                                                    MapPokemon(MapPokemonNum).CurHp = MapPokemon(MapPokemonNum).CurHp + Absorbed
                                                                    SendActionMsg MapNum, "+" & Absorbed, MapPokemon(MapPokemonNum).X * 32, MapPokemon(MapPokemonNum).Y * 32, BrightGreen
                                                                    If MapPokemon(MapPokemonNum).CurHp >= MapPokemon(MapPokemonNum).MaxHp Then
                                                                        MapPokemon(MapPokemonNum).CurHp = MapPokemon(MapPokemonNum).MaxHp
                                                                    End If
                                                                    SendPokemonVital MapPokemonNum
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Case 2 '//Buff/Debuff
                                                    For z = 1 To StatEnum.Stat_Count - 1
                                                        PlayerPokemon(I).StatBuff(z) = PlayerPokemon(I).StatBuff(z) + PokemonMove(MoveNum).dStat(z)
                                                        If PlayerPokemon(I).StatBuff(z) > 6 Then
                                                            PlayerPokemon(I).StatBuff(z) = 6
                                                        ElseIf PlayerPokemon(I).StatBuff(z) < -6 Then
                                                            PlayerPokemon(I).StatBuff(z) = -6
                                                        End If
                                                    Next
                                                    SendPlayerPokemonStatBuff I
                                                Case 3 '//Heal
                                                    
                                            End Select
                                            '//Status
                                            If PokemonMove(MoveNum).pStatus > 0 And PokemonMove(MoveNum).StatusToSelf = NO Then
                                                If PlayerPokemon(I).slot > 0 Then
                                                    If PokemonMove(MoveNum).pStatus = 6 Then
                                                        PlayerPokemon(I).IsConfuse = YES
                                                        SendPlayerPokemonStatus I
                                                        Select Case TempPlayer(I).CurLanguage
                                                            Case LANG_PT: AddAlert I, "Your pokemon got confused", White
                                                            Case LANG_EN: AddAlert I, "Your pokemon got confused", White
                                                            Case LANG_ES: AddAlert I, "Your pokemon got confused", White
                                                        End Select
                                                    Else
                                                        If PlayerPokemons(I).Data(PlayerPokemon(I).slot).Status <= 0 Then
                                                            statusChance = (100 * (PokemonMove(MoveNum).pStatusChance / 100))
                                
                                                            If IsImmuneOnStatus(PokemonMove(MoveNum).Type, Pokemon(PlayerPokemon(I).Num).PrimaryType, Pokemon(PlayerPokemon(I).Num).SecondaryType, PokemonMove(MoveNum).pStatus) Then
                                                                If statusChance > 0 Then
                                                                    statusRand = Random(1, 100)
                                                                    If statusRand <= statusChance Then
                                                                        PlayerPokemons(I).Data(PlayerPokemon(I).slot).Status = PokemonMove(MoveNum).pStatus
                                                                        SendPlayerPokemonStatus I
                                                                        Select Case PokemonMove(MoveNum).pStatus
                                                                            Case StatusEnum.Poison
                                                                                Select Case TempPlayer(I).CurLanguage
                                                                                    Case LANG_PT: AddAlert I, "Your pokemon got poisoned", White
                                                                                    Case LANG_EN: AddAlert I, "Your pokemon got poisoned", White
                                                                                    Case LANG_ES: AddAlert I, "Your pokemon got poisoned", White
                                                                                End Select
                                                                            Case StatusEnum.Burn
                                                                                Select Case TempPlayer(I).CurLanguage
                                                                                    Case LANG_PT: AddAlert I, "Your pokemon got burned", White
                                                                                    Case LANG_EN: AddAlert I, "Your pokemon got burned", White
                                                                                    Case LANG_ES: AddAlert I, "Your pokemon got burned", White
                                                                                End Select
                                                                            Case StatusEnum.Paralize
                                                                                Select Case TempPlayer(I).CurLanguage
                                                                                    Case LANG_PT: AddAlert I, "Your pokemon got paralized", White
                                                                                    Case LANG_EN: AddAlert I, "Your pokemon got paralized", White
                                                                                    Case LANG_ES: AddAlert I, "Your pokemon got paralized", White
                                                                                End Select
                                                                            Case StatusEnum.Sleep
                                                                                Select Case TempPlayer(I).CurLanguage
                                                                                    Case LANG_PT: AddAlert I, "Your pokemon fell asleep", White
                                                                                    Case LANG_EN: AddAlert I, "Your pokemon fell asleep", White
                                                                                    Case LANG_ES: AddAlert I, "Your pokemon fell asleep", White
                                                                                End Select
                                                                            Case StatusEnum.Frozen
                                                                                Select Case TempPlayer(I).CurLanguage
                                                                                    Case LANG_PT: AddAlert I, "Your pokemon got frozed", White
                                                                                    Case LANG_EN: AddAlert I, "Your pokemon got frozed", White
                                                                                    Case LANG_ES: AddAlert I, "Your pokemon got frozed", White
                                                                                End Select
                                                                        End Select
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                'Adicionado a um método, pra ser usado juntamente com o PVP
                                Call AttackPlayer(I, MoveNum, MapPokemonNum, X, Y, MapPokemon(MapPokemonNum).Dir, pType, ownType, ownLevel, AtkStat, Power, MapPokemon(MapPokemonNum).NextCritical, MapPokemon(MapPokemonNum).CurHp, MapPokemon(MapPokemonNum).MaxHp)
                            End If
                        End If
                    End If
                End If
            Next
    End Select
    
    '//Play Animation
    If PokemonMove(MoveNum).Animation > 0 Then
        If PokemonMove(MoveNum).SelfAnim = YES Then
            SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, MapPokemon(MapPokemonNum).X, MapPokemon(MapPokemonNum).Y
        Else
            '//Check Target Type
            Select Case PokemonMove(MoveNum).targetType
                Case 0 '//Self
                    SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, MapPokemon(MapPokemonNum).X, MapPokemon(MapPokemonNum).Y
                Case 1 '//AoE
                    If Range > 0 Then
                        For X = MapPokemon(MapPokemonNum).X - Range To MapPokemon(MapPokemonNum).X + Range
                            For Y = MapPokemon(MapPokemonNum).Y - Range To MapPokemon(MapPokemonNum).Y + Range
                                If isValidMapPoint(MapNum, X, Y) Then
                                    If IsOnAoERange(Range, MapPokemon(MapPokemonNum).X, MapPokemon(MapPokemonNum).Y, X, Y) Then
                                        SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                                    End If
                                End If
                            Next
                        Next
                    End If
                Case 2 '//Linear
                    If Range > 0 Then
                        Select Case MapPokemon(MapPokemonNum).Dir
                            Case DIR_UP
                                For Y = MapPokemon(MapPokemonNum).Y - Range To MapPokemon(MapPokemonNum).Y - 1
                                    X = MapPokemon(MapPokemonNum).X
                                    SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                                Next
                            Case DIR_DOWN
                                For Y = MapPokemon(MapPokemonNum).Y + 1 To MapPokemon(MapPokemonNum).Y + Range
                                    X = MapPokemon(MapPokemonNum).X
                                    SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                                Next
                            Case DIR_LEFT
                                For X = MapPokemon(MapPokemonNum).X - Range To MapPokemon(MapPokemonNum).X - 1
                                    Y = MapPokemon(MapPokemonNum).Y
                                    SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                                Next
                            Case DIR_RIGHT
                                For X = MapPokemon(MapPokemonNum).X + 1 To MapPokemon(MapPokemonNum).X + Range
                                    Y = MapPokemon(MapPokemonNum).Y
                                    SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                                Next
                        End Select
                    End If
                Case 3 '//Spray
                    If Range > 0 Then
                        z = 1
                        Select Case MapPokemon(MapPokemonNum).Dir
                            Case DIR_UP
                                For Y = MapPokemon(MapPokemonNum).Y - 1 To MapPokemon(MapPokemonNum).Y - Range Step -1
                                    For X = MapPokemon(MapPokemonNum).X - z To MapPokemon(MapPokemonNum).X + z
                                        SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                                    Next
                                    z = z + 1
                                Next
                            Case DIR_DOWN
                                For Y = MapPokemon(MapPokemonNum).Y + 1 To MapPokemon(MapPokemonNum).Y + Range
                                    For X = MapPokemon(MapPokemonNum).X - z To MapPokemon(MapPokemonNum).X + z
                                        SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                                    Next
                                    z = z + 1
                                Next
                            Case DIR_LEFT
                                For X = MapPokemon(MapPokemonNum).X - 1 To MapPokemon(MapPokemonNum).X - Range Step -1
                                    For Y = MapPokemon(MapPokemonNum).Y - z To MapPokemon(MapPokemonNum).Y + z
                                        SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                                    Next
                                    z = z + 1
                                Next
                            Case DIR_RIGHT
                                For X = MapPokemon(MapPokemonNum).X + 1 To MapPokemon(MapPokemonNum).X + Range
                                    For Y = MapPokemon(MapPokemonNum).Y - z To MapPokemon(MapPokemonNum).Y + z
                                        SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                                    Next
                                    z = z + 1
                                Next
                        End Select
                    End If
            End Select
        End If
    End If
    
    '//Status
    If PokemonMove(MoveNum).pStatus > 0 And PokemonMove(MoveNum).StatusToSelf = YES Then
        If PokemonMove(MoveNum).pStatus = 6 Then
            MapPokemon(MapPokemonNum).IsConfuse = YES
        Else
            If MapPokemon(MapPokemonNum).Status <= 0 Then
                statusChance = (100 * (PokemonMove(MoveNum).pStatusChance / 100))
                        
                If IsImmuneOnStatus(PokemonMove(MoveNum).Type, Pokemon(MapPokemon(MapPokemonNum).Num).PrimaryType, Pokemon(MapPokemon(MapPokemonNum).Num).SecondaryType, PokemonMove(MoveNum).pStatus) Then
                    If statusChance > 0 Then
                        statusRand = Random(1, 100)
                        If statusRand <= statusChance Then
                            MapPokemon(MapPokemonNum).Status = PokemonMove(MoveNum).pStatus
                            SendMapPokemonStatus MapPokemonNum
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If PokemonMove(MoveNum).RecoilDamage > 0 Then
        recoil = PokemonMove(MoveNum).RecoilDamage
        Damage = MapPokemon(MapPokemonNum).MaxHp * (recoil / 100)
        If Damage > 0 Then
            MapPokemon(MapPokemonNum).CurHp = MapPokemon(MapPokemonNum).CurHp - Damage
            SendActionMsg MapNum, "-" & Damage, MapPokemon(MapPokemonNum).X * 32, MapPokemon(MapPokemonNum).Y * 32, BrightRed
            If MapPokemon(MapPokemonNum).CurHp <= 0 Then
                MapPokemon(MapPokemonNum).CurHp = 0
    
                DefeatMapPokemon MapPokemonNum
                
                ClearMapPokemon MapPokemonNum
            Else
                SendPokemonVital MapPokemonNum
            End If
        End If
    End If
    
    '//Play Sound
    If Not Trim$(PokemonMove(MoveNum).Sound) = "None." Or Not Trim$(PokemonMove(MoveNum).Sound) = vbNullString Then
        SendPlaySound Trim$(PokemonMove(MoveNum).Sound), MapNum
    End If
End Sub

Private Sub AttackPlayer(I As Long, MoveNum As Long, _
                        MapPokemonNum As Long, X As Long, Y As Long, Dir As Byte, _
                        pType As Byte, ownType As Byte, ownLevel As Byte, AtkStat As Long, Power As Long, _
                        NextCritical As Byte, CurHp As Long, MaxHp As Long)

'//Damage Calculation
Dim Damage As Long
Dim targetType As Byte, targetType2 As Byte
Dim DefStat As Long

Dim InRange As Boolean
Dim CanAttack As Boolean
Dim statusChance As Long
Dim statusRand As Long
Dim Absorbed As Long
Dim setBuff As Long

'//Attack Player
    With Player(I, TempPlayer(I).UseChar)
        '//Check Status Req
        If PokemonMove(MoveNum).StatusReq > 0 Then
            If .Status = PokemonMove(MoveNum).StatusReq Then
                CanAttack = True
            Else
                CanAttack = False
            End If
        Else
            CanAttack = True
        End If

        If CanAttack Then
            '//Check Location
            InRange = False
            If PokemonMove(MoveNum).targetType = 1 Then    '//AoE
                If IsOnAoERange(PokemonMove(MoveNum).Range, X, Y, .X, .Y) Then InRange = True
            ElseIf PokemonMove(MoveNum).targetType = 2 Then    '//Linear
                If IsOnLinearRange(Dir, PokemonMove(MoveNum).Range, X, Y, .X, .Y) Then InRange = True
            ElseIf PokemonMove(MoveNum).targetType = 3 Then    '//Spray
                If IsOnSprayRange(Dir, PokemonMove(MoveNum).Range, X, Y, .X, .Y) Then InRange = True
            Else
                InRange = False
            End If

            If InRange Then
                '//Check Move
                Select Case PokemonMove(MoveNum).AttackType
                Case 1    '//Damage
                    '//Get Target
                    '//Target and Do Damage
                    targetType = PokemonType.typeNormal
                    targetType2 = 0
                    DefStat = 1
                    Damage = GetPokemonDamage(ownType, pType, targetType, targetType2, ownLevel, AtkStat, Power, DefStat)
                    If NextCritical = YES Then
                        Damage = Damage * 2
                        SendActionMsg GetPlayerMap(I), "Critical", X * 32, Y * 32, Yellow
                    End If
                    If Damage > 0 Then
                        setBuff = Spawn(MapPokemonNum).pokeBuff
                        If setBuff > 0 Then
                            Damage = Damage * setBuff
                        End If

                        NpcAttackPlayerTrainer GetPlayerMap(I), I, Damage
                        '//Absorb
                        If PokemonMove(MoveNum).AbsorbDamage > 0 Then
                            Absorbed = Damage * (PokemonMove(MoveNum).AbsorbDamage / 100)
                            If Absorbed > 0 Then
                                CurHp = CurHp + Absorbed
                                SendActionMsg GetPlayerMap(I), "+" & Absorbed, X * 32, Y * 32, BrightGreen
                                If CurHp >= MaxHp Then
                                    CurHp = MaxHp
                                End If
                                SendPokemonVital MapPokemonNum
                            End If
                        End If
                    End If
                End Select
                '//Status
                If PokemonMove(MoveNum).pStatus > 0 And PokemonMove(MoveNum).StatusToSelf = NO Then
                    If PokemonMove(MoveNum).pStatus = 6 Then
                        .IsConfuse = YES
                        SendPlayerStatus I
                        Select Case TempPlayer(I).CurLanguage
                        Case LANG_PT: AddAlert I, "You got confused", White
                        Case LANG_EN: AddAlert I, "You got confused", White
                        Case LANG_ES: AddAlert I, "You got confused", White
                        End Select
                    Else
                        If .Status <= 0 Then
                            statusChance = (100 * (PokemonMove(MoveNum).pStatusChance / 100))

                            If statusChance > 0 Then
                                statusRand = Random(1, 100)
                                If statusRand <= statusChance Then
                                    .Status = PokemonMove(MoveNum).pStatus
                                    SendPlayerStatus I
                                    Select Case PokemonMove(MoveNum).pStatus
                                    Case StatusEnum.Poison
                                        Select Case TempPlayer(I).CurLanguage
                                        Case LANG_PT: AddAlert I, "You got poisoned", White
                                        Case LANG_EN: AddAlert I, "You got poisoned", White
                                        Case LANG_ES: AddAlert I, "You got poisoned", White
                                        End Select
                                    Case StatusEnum.Burn
                                        Select Case TempPlayer(I).CurLanguage
                                        Case LANG_PT: AddAlert I, "You got burned", White
                                        Case LANG_EN: AddAlert I, "You got burned", White
                                        Case LANG_ES: AddAlert I, "You got burned", White
                                        End Select
                                    Case StatusEnum.Paralize
                                        Select Case TempPlayer(I).CurLanguage
                                        Case LANG_PT: AddAlert I, "You got paralized", White
                                        Case LANG_EN: AddAlert I, "You got paralized", White
                                        Case LANG_ES: AddAlert I, "You got paralized", White
                                        End Select
                                    Case StatusEnum.Sleep
                                        Select Case TempPlayer(I).CurLanguage
                                        Case LANG_PT: AddAlert I, "You fell asleep", White
                                        Case LANG_EN: AddAlert I, "You fell asleep", White
                                        Case LANG_ES: AddAlert I, "You fell asleep", White
                                        End Select
                                    Case StatusEnum.Frozen
                                        Select Case TempPlayer(I).CurLanguage
                                        Case LANG_PT: AddAlert I, "You got frozed", White
                                        Case LANG_EN: AddAlert I, "You got frozed", White
                                        Case LANG_ES: AddAlert I, "You got frozed", White
                                        End Select
                                    End Select
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub PlayerAttackNpc(ByVal Index As Long, ByVal TargetIndex As Long, ByVal Damage As Long)
    Dim MapNum As Long
    Dim RndNum As Long
    Dim expEarn As Long
    Dim Level As Long
    Dim checkItem As Long
    Dim ChanceNum As Long
    Dim slotHaveItem As Byte

    '//Check Error
    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If PlayerPokemon(Index).Num <= 0 Then Exit Sub
    If TargetIndex <= 0 Or TargetIndex > MAX_GAME_POKEMON Then Exit Sub
    If MapPokemon(TargetIndex).Num <= 0 Then Exit Sub
    If MapPokemon(TargetIndex).InCatch = YES Then Exit Sub

    MapNum = Player(Index, TempPlayer(Index).UseChar).Map

    MapPokemon(TargetIndex).LastAttacker = Index

    If Damage >= MapPokemon(TargetIndex).CurHp Then
        ' Define a vida atual do Pokémon como 0
        MapPokemon(TargetIndex).CurHp = 0
        ' Envia uma mensagem de ação para exibir o dano na tela
        SendActionMsg MapNum, "-" & Damage, MapPokemon(TargetIndex).X * 32, MapPokemon(TargetIndex).Y * 32, BrightRed

        ' Gera um número aleatório entre 0 e 6
        RndNum = Random(0, 6)

        ' Obtém o número do slot que possui um item do NPC
        slotHaveItem = FindNpcDropSlotHaveItem(Index, TargetIndex)

        ' Verifica se não há item em nenhum slot, caso contrário, define RndNum como 0 e procede pro money
        If slotHaveItem = 0 Then RndNum = 0

        ' Verifica se RndNum está entre 1 e slotHaveItem
        If RndNum >= 1 And RndNum <= slotHaveItem Then
            ' Verifica se o Pokémon atual tem um item na posição RndNum na lista de itens
            If Pokemon(MapPokemon(TargetIndex).Num).DropNum(RndNum) > 0 Then
                ' Encontra um espaço vazio no inventário do jogador para o item
                checkItem = FindFreeInvSlot(Index, Pokemon(MapPokemon(TargetIndex).Num).DropNum(RndNum))
                If checkItem > 0 Then
                    ' Geração de um número aleatório entre 0 e 100 para determinar a chance de queda do item
                    ChanceNum = Random(0, 100)
                    If ChanceNum <= Pokemon(MapPokemon(TargetIndex).Num).DropRate(RndNum) Then
                        ' Dá o item ao jogador
                        GiveItem Index, Pokemon(MapPokemon(TargetIndex).Num).DropNum(RndNum), 1
                        AddAlert Index, "Pokemon drop a " & Trim$(Item(Pokemon(MapPokemon(TargetIndex).Num).DropNum(RndNum)).Name), White
                    End If
                End If
            End If
        Else
            ' Gera um número aleatório entre 0 e 2
            RndNum = Random(0, 2)
            If RndNum = 1 Then
                ' Geração de um número aleatório entre 1 e o dobro do nível do Pokémon
                checkItem = Random(1, MapPokemon(TargetIndex).Level * 2)
                If checkItem > 0 Then
                    ' Adiciona a quantidade de dinheiro ao jogador
                    Player(Index, TempPlayer(Index).UseChar).Money = Player(Index, TempPlayer(Index).UseChar).Money + checkItem
                    If Player(Index, TempPlayer(Index).UseChar).Money > MAX_MONEY Then
                        Player(Index, TempPlayer(Index).UseChar).Money = MAX_MONEY
                    End If
                    ' Atualiza os dados do jogador
                    SendPlayerData Index
                    AddAlert Index, "Pokemon drop $" & checkItem, White
                End If
            End If
        End If

        '//Give Exp
        '// ToDo: First 1 = If trade pokemon is 1.5 normal is 1
        '// ToDo: Second 2 = If trainer pokemon is 1.5 normal is 1
        ' Verifica se o Pokémon derrotado pode ganhar experiência
        If Spawn(TargetIndex).NoExp = NO Then
            ' Derrota o Pokémon do mapa
            DefeatMapPokemon TargetIndex
            ' Verifica se o pokemon está usando um Power Bracer e envia o EV
            GivePlayerEvPowerBracer Index, PlayerPokemon(Index).slot
            ' Adiciona a experiência do Pokémon derrotado ao Pokémon do jogador
            GivePlayerPokemonEVExp Index, PlayerPokemon(Index).slot, (Pokemon(MapPokemon(TargetIndex).Num).EvYeildType + 1), Pokemon(MapPokemon(TargetIndex).Num).EvYeildVal
        End If

        ' Limpa o Pokémon do mapa
        ClearMapPokemon TargetIndex

    Else
        MapPokemon(TargetIndex).CurHp = MapPokemon(TargetIndex).CurHp - Damage
        SendActionMsg MapNum, "-" & Damage, MapPokemon(TargetIndex).X * 32, MapPokemon(TargetIndex).Y * 32, BrightRed

        '//Update
        SendPokemonVital TargetIndex

        '//Set Target
        If MapPokemon(TargetIndex).TargetIndex = 0 Then
            MapPokemon(TargetIndex).TargetIndex = Index
            MapPokemon(TargetIndex).targetType = TARGET_TYPE_PLAYER
        Else
            If Not MapPokemon(TargetIndex).TargetIndex = Index Then
                RndNum = Random(0, 4)
                If RndNum = 1 Then
                    MapPokemon(TargetIndex).TargetIndex = Index
                    MapPokemon(TargetIndex).targetType = TARGET_TYPE_PLAYER
                End If
            End If
        End If
    End If
End Sub

Public Sub PlayerAttackPlayer(ByVal Index As Long, ByVal TargetIndex As Long, ByVal Damage As Long)
Dim MapNum As Long
Dim DuelIndex As Long

    '//Check Error
    If Not IsPlaying(Index) Then Exit Sub
    If Not IsPlaying(TargetIndex) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If TempPlayer(TargetIndex).UseChar <= 0 Then Exit Sub
    If PlayerPokemon(Index).Num <= 0 Then Exit Sub
    If PlayerPokemon(TargetIndex).Num <= 0 Then Exit Sub
    If PlayerPokemon(TargetIndex).slot <= 0 Then Exit Sub
    
    MapNum = Player(Index, TempPlayer(Index).UseChar).Map
    
    If Damage >= PlayerPokemons(TargetIndex).Data(PlayerPokemon(TargetIndex).slot).CurHp Then
        '//Defeat
        PlayerPokemons(TargetIndex).Data(PlayerPokemon(TargetIndex).slot).CurHp = 0
        SendActionMsg MapNum, "-" & Damage, PlayerPokemon(TargetIndex).X * 32, PlayerPokemon(TargetIndex).Y * 32, BrightRed
        
        '//Update Vital
        SendPlayerPokemonVital TargetIndex
        SendPlayerPokemonFaint TargetIndex
    Else
        PlayerPokemons(TargetIndex).Data(PlayerPokemon(TargetIndex).slot).CurHp = PlayerPokemons(TargetIndex).Data(PlayerPokemon(TargetIndex).slot).CurHp - Damage
        SendActionMsg MapNum, "-" & Damage, PlayerPokemon(TargetIndex).X * 32, PlayerPokemon(TargetIndex).Y * 32, BrightRed
        
        '//Update
        SendPlayerPokemonVital TargetIndex
    End If
End Sub

Public Sub NpcAttackPlayer(ByVal MapNum As Long, ByVal MapPokeNum As Long, ByVal TargetIndex As Long, ByVal Damage As Long)
Dim DuelIndex As Long

    '//Check Error
    If Not IsPlaying(TargetIndex) Then Exit Sub
    If TempPlayer(TargetIndex).UseChar <= 0 Then Exit Sub
    If PlayerPokemon(TargetIndex).Num <= 0 Then Exit Sub
    If PlayerPokemon(TargetIndex).slot <= 0 Then Exit Sub
    If MapPokemon(MapPokeNum).Num <= 0 Then Exit Sub
    
    If Damage >= PlayerPokemons(TargetIndex).Data(PlayerPokemon(TargetIndex).slot).CurHp Then
        '//Defeat
        PlayerPokemons(TargetIndex).Data(PlayerPokemon(TargetIndex).slot).CurHp = 0
        SendActionMsg MapNum, "-" & Damage, PlayerPokemon(TargetIndex).X * 32, PlayerPokemon(TargetIndex).Y * 32, BrightRed
        SendPlayerPokemonVital TargetIndex
        SendPlayerPokemonFaint TargetIndex
    Else
        PlayerPokemons(TargetIndex).Data(PlayerPokemon(TargetIndex).slot).CurHp = PlayerPokemons(TargetIndex).Data(PlayerPokemon(TargetIndex).slot).CurHp - Damage
        SendActionMsg MapNum, "-" & Damage, PlayerPokemon(TargetIndex).X * 32, PlayerPokemon(TargetIndex).Y * 32, BrightRed
        
        '//Update
        SendPlayerPokemonVital TargetIndex
    End If
End Sub

Public Sub NpcAttackPlayerTrainer(ByVal MapNum As Long, ByVal TargetIndex As Long, ByVal Damage As Long)
    '//Check Error
    If Not IsPlaying(TargetIndex) Then Exit Sub
    If TempPlayer(TargetIndex).UseChar <= 0 Then Exit Sub
    'If MapPokemon(MapPokeNum).Num <= 0 Then Exit Sub
    
    If Damage >= Player(TargetIndex, TempPlayer(TargetIndex).UseChar).CurHp Then
        '//Defeat
        Player(TargetIndex, TempPlayer(TargetIndex).UseChar).CurHp = 0
        SendActionMsg MapNum, "-" & Damage, Player(TargetIndex, TempPlayer(TargetIndex).UseChar).X * 32, Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Y * 32, BrightRed
        
        '//OnDeath
        KillPlayer TargetIndex
    Else
        Player(TargetIndex, TempPlayer(TargetIndex).UseChar).CurHp = Player(TargetIndex, TempPlayer(TargetIndex).UseChar).CurHp - Damage
        SendActionMsg MapNum, "-" & Damage, Player(TargetIndex, TempPlayer(TargetIndex).UseChar).X * 32, Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Y * 32, BrightRed
        
        '//Update
        SendPlayerVital TargetIndex
    End If
End Sub

'/////////////////////////
'///// Npc Trainer Attack /////
'/////////////////////////
Public Sub NpcPokemonCastMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal MoveNum As Long, ByVal MoveSlot As Byte, Optional ByVal DecreasePP As Boolean = True)
Dim RandomNum As Byte

    '//Check for error
    If MapNum <= 0 Or MapNum > MAX_MAP Then Exit Sub
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPC Then Exit Sub
    If MapNpcPokemon(MapNum, MapNpcNum).Num <= 0 Then Exit Sub
    If MoveNum <= 0 Or MoveNum > MAX_POKEMON_MOVE Then Exit Sub
    If MapNpc(MapNum, MapNpcNum).CurPokemon <= 0 Then Exit Sub
    
    '//Add Queue
    With MapNpcPokemon(MapNum, MapNpcNum)
        If Not PokemonMove(MoveNum).SelfStatusReq = StatusEnum.Sleep Then
            If .Status = StatusEnum.Sleep Then
                RandomNum = Random(1, 5)
                If RandomNum = 1 Then
                    '//Remove Status
                    .Status = 0
                    SendMapNpcPokemonStatus MapNum, MapNpcNum
                Else
                    Exit Sub
                End If
            End If
        End If
        If Not PokemonMove(MoveNum).SelfStatusReq = StatusEnum.Frozen Then
            If .Status = StatusEnum.Frozen Then
                RandomNum = Random(1, 3)
                If RandomNum = 1 Then
                    '//Remove Status
                    .Status = 0
                    SendMapNpcPokemonStatus MapNum, MapNpcNum
                Else
                    Exit Sub
                End If
            End If
        End If
        
        If PokemonMove(MoveNum).SelfStatusReq > 0 Then
            If Not .Status = PokemonMove(MoveNum).SelfStatusReq Then
                Exit Sub
            End If
        End If
        
        '//Check PP
        If MoveSlot > 0 Then
            If .Moveset(MoveSlot).CurPP <= 0 Then Exit Sub
            '//Check Cooldown
            If .Moveset(MoveSlot).CD + PokemonMove(MoveNum).Cooldown > GetTickCount Then Exit Sub
        End If
        
        '//Burn
        If .Status = StatusEnum.Burn Then
            If .StatusDamage > 0 Then
                If .StatusDamage >= .CurHp Then
                    '//Dead
                    .CurHp = 0
                    SendActionMsg MapNum, "-" & .StatusDamage, .X * 32, .Y * 32, BrightRed
                    SendNpcPokemonVital MapNum, MapNpcNum
                    MapNpc(MapNum, MapNpcNum).PokemonAlive(MapNpc(MapNum, MapNpcNum).CurPokemon) = NO
                    NpcPokemonCallBack MapNum, MapNpcNum
                    Exit Sub
                Else
                    '//Reduce
                    .CurHp = .CurHp - .StatusDamage
                    SendActionMsg MapNum, "-" & .StatusDamage, .X * 32, .Y * 32, BrightRed
                    '//Update
                    SendNpcPokemonVital MapNum, MapNpcNum
                End If
            Else
                .StatusDamage = (.MaxHp / 8)
            End If
        End If
        
        .QueueMove = MoveNum
        .QueueMoveSlot = MoveSlot
        
        '//Set Duration
        .MoveCastTime = GetTickCount + PokemonMove(MoveNum).CastTime
        .MoveDuration = GetTickCount + PokemonMove(MoveNum).Duration
        .MoveInterval = GetTickCount
        .MoveAttackCount = 0
        
        '//Decrease PP
        If MoveSlot > 0 Then
            If DecreasePP Then
                .Moveset(MoveSlot).CurPP = .Moveset(MoveSlot).CurPP - 1
            End If
            
            '//Add ActionMsg
            SendActionMsg MapNum, Trim$(PokemonMove(MoveNum).Name), .X * 32, .Y * 32, Yellow
        End If
    End With
End Sub
 
Public Sub ProcessNpcPokemonMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal MoveNum As Long)
Dim I As Long
Dim Range As Long
Dim X As Long, Y As Long
Dim pType As Byte
Dim Power As Long

'//Damage Calculation
Dim ownType As Byte
Dim ownLevel As Byte
Dim AtkStat As Long
Dim Damage As Long
Dim targetType As Byte, targetType2 As Byte
Dim DefStat As Long

Dim InRange As Boolean
Dim z As Byte
Dim HealAmount As Long
Dim statusChance As Long
Dim statusRand As Long
Dim DuelIndex As Long
Dim recoil As Long
Dim Absorbed As Long
Dim CanAttack As Boolean

    '//Check for error
    If MapNum <= 0 Or MapNum > MAX_MAP Then Exit Sub
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPC Then Exit Sub
    If MapNpcPokemon(MapNum, MapNpcNum).Num <= 0 Then Exit Sub
    If MoveNum <= 0 Or MoveNum > MAX_POKEMON_MOVE Then Exit Sub
    If MapNpc(MapNum, MapNpcNum).CurPokemon <= 0 Then Exit Sub
    
    Range = PokemonMove(MoveNum).Range
    Power = PokemonMove(MoveNum).Power
    X = MapNpcPokemon(MapNum, MapNpcNum).X
    Y = MapNpcPokemon(MapNum, MapNpcNum).Y
    ownType = Pokemon(MapNpcPokemon(MapNum, MapNpcNum).Num).PrimaryType
    ownLevel = MapNpcPokemon(MapNum, MapNpcNum).Level
    pType = PokemonMove(MoveNum).Type
    
    '//Check Attack Category
    Select Case PokemonMove(MoveNum).Category
        Case MoveCategory.Physical
            AtkStat = GetNpcTrainerPokemonStat(MapNum, MapNpcNum, Atk)
        Case MoveCategory.Special
            AtkStat = GetNpcTrainerPokemonStat(MapNum, MapNpcNum, SpAtk)
    End Select
    
    '//Get Target
    Select Case PokemonMove(MoveNum).targetType
        Case 0 '//Self
            '//Status
            If PokemonMove(MoveNum).pStatus > 0 And PokemonMove(MoveNum).StatusToSelf = NO Then
                If PokemonMove(MoveNum).pStatus = 6 Then
                    MapNpcPokemon(MapNum, MapNpcNum).IsConfuse = YES
                    'SendPlayerPokemonStatus Index
                Else
                    If MapNpcPokemon(MapNum, MapNpcNum).Status <= 0 Then
                        statusChance = (100 * (PokemonMove(MoveNum).pStatusChance / 100))
                        
                        If IsImmuneOnStatus(PokemonMove(MoveNum).Type, Pokemon(MapNpcPokemon(MapNum, MapNpcNum).Num).PrimaryType, Pokemon(MapNpcPokemon(MapNum, MapNpcNum).Num).SecondaryType, PokemonMove(MoveNum).pStatus) Then
                            If statusChance > 0 Then
                                statusRand = Random(1, 100)
                                If statusRand <= statusChance Then
                                    MapNpcPokemon(MapNum, MapNpcNum).Status = PokemonMove(MoveNum).pStatus
                                    'SendPlayerPokemonStatus Index
                                    SendMapNpcPokemonStatus MapNum, MapNpcNum
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            Select Case PokemonMove(MoveNum).AttackType
                Case 2 '//Buff/Debuff
                    For z = 1 To StatEnum.Stat_Count - 1
                        MapNpcPokemon(MapNum, MapNpcNum).StatBuff(z) = MapNpcPokemon(MapNum, MapNpcNum).StatBuff(z) + PokemonMove(MoveNum).dStat(z)
                        If MapNpcPokemon(MapNum, MapNpcNum).StatBuff(z) > 6 Then
                            MapNpcPokemon(MapNum, MapNpcNum).StatBuff(z) = 6
                        ElseIf MapNpcPokemon(MapNum, MapNpcNum).StatBuff(z) < -6 Then
                            MapNpcPokemon(MapNum, MapNpcNum).StatBuff(z) = -6
                        End If
                    Next
                    'SendPlayerPokemonStatBuff Index
                Case 3 '//Heal
                    HealAmount = MapNpcPokemon(MapNum, MapNpcNum).MaxHp * (PokemonMove(MoveNum).Power / 100)
                    MapNpcPokemon(MapNum, MapNpcNum).CurHp = MapNpcPokemon(MapNum, MapNpcNum).CurHp + HealAmount
                    If MapNpcPokemon(MapNum, MapNpcNum).CurHp >= MapNpcPokemon(MapNum, MapNpcNum).MaxHp Then
                        MapNpcPokemon(MapNum, MapNpcNum).CurHp = MapNpcPokemon(MapNum, MapNpcNum).MaxHp
                    End If
                    SendNpcPokemonVital MapNum, MapNpcNum
            End Select
            '//Reflect
            If PokemonMove(MoveNum).ReflectType > 0 Then
                MapNpcPokemon(MapNum, MapNpcNum).ReflectMove = PokemonMove(MoveNum).ReflectType
            End If
            If PokemonMove(MoveNum).CastProtect > 0 Then
                MapNpcPokemon(MapNum, MapNpcNum).IsProtect = YES
            End If
        Case 1, 2, 3 '//Linear , AOE , Spray
            '//Check Target
            '//Player
            For I = 1 To Player_HighIndex
                If IsPlaying(I) Then
                    If TempPlayer(I).UseChar > 0 Then
                        If Player(I, TempPlayer(I).UseChar).Map = MapNum Then
                            '//Can't kill player
                            If PlayerPokemon(I).Num > 0 Then
                                '//Check Status Req
                                If PokemonMove(MoveNum).StatusReq > 0 Then
                                    If PlayerPokemons(I).Data(PlayerPokemon(I).slot).Status = PokemonMove(MoveNum).StatusReq Then
                                        CanAttack = True
                                    Else
                                        CanAttack = False
                                    End If
                                Else
                                    CanAttack = True
                                End If
                                If PlayerPokemon(I).IsProtect > 0 Then
                                    CanAttack = False
                                    PlayerPokemon(I).IsProtect = NO
                                    SendActionMsg MapNum, "Protected", PlayerPokemon(I).X * 32, PlayerPokemon(I).Y * 32, Yellow
                                End If
                                
                                If CanAttack Then
                                    '//Check Location
                                    If MapNpc(MapNum, MapNpcNum).InBattle = I And TempPlayer(I).InNpcDuel = MapNpcNum Then
                                        If TempPlayer(I).DuelTime <= 0 Then
                                            InRange = False
                                            If PokemonMove(MoveNum).targetType = 1 Then '//AoE
                                                If IsOnAoERange(Range, X, Y, PlayerPokemon(I).X, PlayerPokemon(I).Y) Then InRange = True
                                            ElseIf PokemonMove(MoveNum).targetType = 2 Then '//Linear
                                                If IsOnLinearRange(MapNpcPokemon(MapNum, MapNpcNum).Dir, Range, X, Y, PlayerPokemon(I).X, PlayerPokemon(I).Y) Then InRange = True
                                            ElseIf PokemonMove(MoveNum).targetType = 3 Then '//Spray
                                                If IsOnSprayRange(MapNpcPokemon(MapNum, MapNpcNum).Dir, Range, X, Y, PlayerPokemon(I).X, PlayerPokemon(I).Y) Then InRange = True
                                            Else
                                                InRange = False
                                            End If
                                                
                                            If InRange Then
                                                If PlayerPokemon(I).slot > 0 Then
                                                    If PokemonMove(MoveNum).pStatus = 6 Then
                                                        PlayerPokemon(I).IsConfuse = YES
                                                        SendPlayerPokemonStatus I
                                                        Select Case TempPlayer(I).CurLanguage
                                                            Case LANG_PT: AddAlert I, "Your pokemon got confused", White
                                                            Case LANG_EN: AddAlert I, "Your pokemon got confused", White
                                                            Case LANG_ES: AddAlert I, "Your pokemon got confused", White
                                                        End Select
                                                    Else
                                                        '//Status
                                                        If PokemonMove(MoveNum).pStatus > 0 And PokemonMove(MoveNum).StatusToSelf = NO Then
                                                            If PlayerPokemons(I).Data(PlayerPokemon(I).slot).Status <= 0 Then
                                                                statusChance = (100 * (PokemonMove(MoveNum).pStatusChance / 100))
                                
                                                                If IsImmuneOnStatus(PokemonMove(MoveNum).Type, Pokemon(PlayerPokemon(I).Num).PrimaryType, Pokemon(PlayerPokemon(I).Num).SecondaryType, PokemonMove(MoveNum).pStatus) Then
                                                                    If statusChance > 0 Then
                                                                        statusRand = Random(1, 100)
                                                                        If statusRand <= statusChance Then
                                                                            PlayerPokemons(I).Data(PlayerPokemon(I).slot).Status = PokemonMove(MoveNum).pStatus
                                                                            SendPlayerPokemonStatus I
                                                                            Select Case PokemonMove(MoveNum).pStatus
                                                                                Case StatusEnum.Poison
                                                                                    Select Case TempPlayer(I).CurLanguage
                                                                                        Case LANG_PT: AddAlert I, "Your pokemon got poisoned", White
                                                                                        Case LANG_EN: AddAlert I, "Your pokemon got poisoned", White
                                                                                        Case LANG_ES: AddAlert I, "Your pokemon got poisoned", White
                                                                                    End Select
                                                                                Case StatusEnum.Burn
                                                                                    Select Case TempPlayer(I).CurLanguage
                                                                                        Case LANG_PT: AddAlert I, "Your pokemon got burned", White
                                                                                        Case LANG_EN: AddAlert I, "Your pokemon got burned", White
                                                                                        Case LANG_ES: AddAlert I, "Your pokemon got burned", White
                                                                                    End Select
                                                                                Case StatusEnum.Paralize
                                                                                    Select Case TempPlayer(I).CurLanguage
                                                                                        Case LANG_PT: AddAlert I, "Your pokemon got paralized", White
                                                                                        Case LANG_EN: AddAlert I, "Your pokemon got paralized", White
                                                                                        Case LANG_ES: AddAlert I, "Your pokemon got paralized", White
                                                                                    End Select
                                                                                Case StatusEnum.Sleep
                                                                                    Select Case TempPlayer(I).CurLanguage
                                                                                        Case LANG_PT: AddAlert I, "Your pokemon fell asleep", White
                                                                                        Case LANG_EN: AddAlert I, "Your pokemon fell asleep", White
                                                                                        Case LANG_ES: AddAlert I, "Your pokemon fell asleep", White
                                                                                    End Select
                                                                                Case StatusEnum.Frozen
                                                                                    Select Case TempPlayer(I).CurLanguage
                                                                                        Case LANG_PT: AddAlert I, "Your pokemon got frozed", White
                                                                                        Case LANG_EN: AddAlert I, "Your pokemon got frozed", White
                                                                                        Case LANG_ES: AddAlert I, "Your pokemon got frozed", White
                                                                                    End Select
                                                                            End Select
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                    '//Check Move
                                                    Select Case PokemonMove(MoveNum).AttackType
                                                        Case 1 '//Damage
                                                            '//Target and Do Damage
                                                            targetType = Pokemon(PlayerPokemon(I).Num).PrimaryType
                                                            targetType2 = Pokemon(PlayerPokemon(I).Num).SecondaryType
                                                            Select Case PokemonMove(MoveNum).Category
                                                                Case MoveCategory.Physical
                                                                    DefStat = GetPlayerPokemonStat(I, Def)
                                                                Case MoveCategory.Special
                                                                    DefStat = GetPlayerPokemonStat(I, SpDef)
                                                            End Select
                                                            Damage = GetPokemonDamage(ownType, pType, targetType, targetType2, ownLevel, AtkStat, Power, DefStat)
                                                            '//Check Critical
                                                            If MapNpcPokemon(MapNum, MapNpcNum).NextCritical = YES Then
                                                                Damage = Damage * 2
                                                                SendActionMsg MapNum, "Critical", MapNpcPokemon(MapNum, MapNpcNum).X * 32, MapNpcPokemon(MapNum, MapNpcNum).Y * 32, Yellow
                                                            End If
                                                            If Damage > 0 Then
                                                                '//Check Reflect
                                                                If PlayerPokemon(I).ReflectMove = PokemonMove(MoveNum).Category Then
                                                                    If PlayerPokemon(I).ReflectMove > 0 Then
                                                                        If PlayerPokemon(I).slot > 0 Then
                                                                            PlayerPokemon(I).ReflectMove = 0
                                                                            SendActionMsg MapNum, "Reflected", PlayerPokemon(I).X * 32, PlayerPokemon(I).Y * 32, White
            
                                                                            MapNpcPokemon(MapNum, MapNpcNum).CurHp = MapNpcPokemon(MapNum, MapNpcNum).CurHp - Damage
                                                                            SendActionMsg MapNum, "-" & Damage, MapNpcPokemon(MapNum, MapNpcNum).X * 32, MapNpcPokemon(MapNum, MapNpcNum).Y * 32, BrightGreen
                                                                            If MapNpcPokemon(MapNum, MapNpcNum).CurHp <= 0 Then
                                                                                MapNpcPokemon(MapNum, MapNpcNum).CurHp = 0
                                                                                '//Update
                                                                                SendNpcPokemonVital MapNum, MapNpcNum
                                                                                MapNpc(MapNum, MapNpcNum).PokemonAlive(MapNpc(MapNum, MapNpcNum).CurPokemon) = NO
                                                                                NpcPokemonCallBack MapNum, MapNpcNum
                                                                                TempPlayer(I).DuelTime = 3
                                                                                TempPlayer(I).DuelTimeTmr = GetTickCount + 1000
                                                                            Else
                                                                                '//Update
                                                                                SendNpcPokemonVital MapNum, MapNpcNum
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Else
                                                                    NpcPokemonAttackPlayer MapNum, MapNpcNum, I, Damage
                                                    
                                                                    '//Absorb
                                                                    If PokemonMove(MoveNum).AbsorbDamage > 0 Then
                                                                        Absorbed = Damage * (PokemonMove(MoveNum).AbsorbDamage / 100)
                                                                        If Absorbed > 0 Then
                                                                            MapNpcPokemon(MapNum, MapNpcNum).CurHp = MapNpcPokemon(MapNum, MapNpcNum).CurHp + Absorbed
                                                                            SendActionMsg MapNum, "+" & Absorbed, MapNpcPokemon(MapNum, MapNpcNum).X * 32, MapNpcPokemon(MapNum, MapNpcNum).Y * 32, BrightGreen
                                                                            If MapNpcPokemon(MapNum, MapNpcNum).CurHp >= MapNpcPokemon(MapNum, MapNpcNum).MaxHp Then
                                                                                MapNpcPokemon(MapNum, MapNpcNum).CurHp = MapNpcPokemon(MapNum, MapNpcNum).MaxHp
                                                                            End If
                                                                            SendNpcPokemonVital MapNum, MapNpcNum
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        Case 2 '//Buff/Debuff
                                                            For z = 1 To StatEnum.Stat_Count - 1
                                                                PlayerPokemon(I).StatBuff(z) = PlayerPokemon(I).StatBuff(z) + PokemonMove(MoveNum).dStat(z)
                                                                If PlayerPokemon(I).StatBuff(z) > 6 Then
                                                                    PlayerPokemon(I).StatBuff(z) = 6
                                                                ElseIf PlayerPokemon(I).StatBuff(z) < -6 Then
                                                                    PlayerPokemon(I).StatBuff(z) = -6
                                                                End If
                                                            Next
                                                            SendPlayerPokemonStatBuff I
                                                        Case 3 '//Heal
                                                                
                                                    End Select
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
    End Select
    
    '//Change Weather
    If PokemonMove(MoveNum).ChangeWeather > 0 Then
        If PokemonMove(MoveNum).ChangeWeather = WeatherEnum.Count_Weather Then
            '//Clear
            Map(MapNum).CurWeather = Map(MapNum).StartWeather
        Else
            Map(MapNum).CurWeather = PokemonMove(MoveNum).ChangeWeather
        End If
        SendWeather MapNum
    End If
    
    '//Play Animation
    If PokemonMove(MoveNum).Animation > 0 Then
        If PokemonMove(MoveNum).SelfAnim = YES Then
            SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, MapNpcPokemon(MapNum, MapNpcNum).X, MapNpcPokemon(MapNum, MapNpcNum).Y
        Else
            '//Check Target Type
            Select Case PokemonMove(MoveNum).targetType
                Case 0 '//Self
                    SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, MapNpcPokemon(MapNum, MapNpcNum).X, MapNpcPokemon(MapNum, MapNpcNum).Y
                Case 1 '//AoE
                    If Range > 0 Then
                        For X = MapNpcPokemon(MapNum, MapNpcNum).X - Range To MapNpcPokemon(MapNum, MapNpcNum).X + Range
                            For Y = MapNpcPokemon(MapNum, MapNpcNum).Y - Range To MapNpcPokemon(MapNum, MapNpcNum).Y + Range
                                If isValidMapPoint(MapNum, X, Y) Then
                                    If IsOnAoERange(Range, MapNpcPokemon(MapNum, MapNpcNum).X, MapNpcPokemon(MapNum, MapNpcNum).Y, X, Y) Then
                                        SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                                    End If
                                End If
                            Next
                        Next
                    End If
                Case 2 '//Linear
                    If Range > 0 Then
                        Select Case MapNpcPokemon(MapNum, MapNpcNum).Dir
                            Case DIR_UP
                                For Y = MapNpcPokemon(MapNum, MapNpcNum).Y - Range To MapNpcPokemon(MapNum, MapNpcNum).Y - 1
                                    X = MapNpcPokemon(MapNum, MapNpcNum).X
                                    SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                                Next
                            Case DIR_DOWN
                                For Y = MapNpcPokemon(MapNum, MapNpcNum).Y + 1 To MapNpcPokemon(MapNum, MapNpcNum).Y + Range
                                    X = MapNpcPokemon(MapNum, MapNpcNum).X
                                    SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                                Next
                            Case DIR_LEFT
                                For X = MapNpcPokemon(MapNum, MapNpcNum).X - Range To MapNpcPokemon(MapNum, MapNpcNum).X - 1
                                    Y = MapNpcPokemon(MapNum, MapNpcNum).Y
                                    SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                                Next
                            Case DIR_RIGHT
                                For X = MapNpcPokemon(MapNum, MapNpcNum).X + 1 To MapNpcPokemon(MapNum, MapNpcNum).X + Range
                                    Y = MapNpcPokemon(MapNum, MapNpcNum).Y
                                    SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                                Next
                        End Select
                    End If
                Case 3 '//Spray
                    If Range > 0 Then
                        z = 1
                        Select Case MapNpcPokemon(MapNum, MapNpcNum).Dir
                            Case DIR_UP
                                For Y = MapNpcPokemon(MapNum, MapNpcNum).Y - 1 To MapNpcPokemon(MapNum, MapNpcNum).Y - Range Step -1
                                    For X = MapNpcPokemon(MapNum, MapNpcNum).X - z To MapNpcPokemon(MapNum, MapNpcNum).X + z
                                        SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                                    Next
                                    z = z + 1
                                Next
                            Case DIR_DOWN
                                For Y = MapNpcPokemon(MapNum, MapNpcNum).Y + 1 To MapNpcPokemon(MapNum, MapNpcNum).Y + Range
                                    For X = MapNpcPokemon(MapNum, MapNpcNum).X - z To MapNpcPokemon(MapNum, MapNpcNum).X + z
                                        SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                                    Next
                                    z = z + 1
                                Next
                            Case DIR_LEFT
                                For X = MapNpcPokemon(MapNum, MapNpcNum).X - 1 To MapNpcPokemon(MapNum, MapNpcNum).X - Range Step -1
                                    For Y = MapNpcPokemon(MapNum, MapNpcNum).Y - z To MapNpcPokemon(MapNum, MapNpcNum).Y + z
                                        SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                                    Next
                                    z = z + 1
                                Next
                            Case DIR_RIGHT
                                For X = MapNpcPokemon(MapNum, MapNpcNum).X + 1 To MapNpcPokemon(MapNum, MapNpcNum).X + Range
                                    For Y = MapNpcPokemon(MapNum, MapNpcNum).Y - z To MapNpcPokemon(MapNum, MapNpcNum).Y + z
                                        SendPlayAnimation MapNum, PokemonMove(MoveNum).Animation, X, Y
                                    Next
                                    z = z + 1
                                Next
                        End Select
                    End If
            End Select
        End If
    End If
    
    '//Status
    If PokemonMove(MoveNum).pStatus > 0 And PokemonMove(MoveNum).StatusToSelf = YES Then
        If PokemonMove(MoveNum).pStatus = 6 Then
            MapNpcPokemon(MapNum, MapNpcNum).IsConfuse = YES
            'SendPlayerPokemonStatus Index
        Else
            If MapNpcPokemon(MapNum, MapNpcNum).Status <= 0 Then
                statusChance = (100 * (PokemonMove(MoveNum).pStatusChance / 100))
                        
                If IsImmuneOnStatus(PokemonMove(MoveNum).Type, Pokemon(MapNpcPokemon(MapNum, MapNpcNum).Num).PrimaryType, Pokemon(MapNpcPokemon(MapNum, MapNpcNum).Num).SecondaryType, PokemonMove(MoveNum).pStatus) Then
                    If statusChance > 0 Then
                        statusRand = Random(1, 100)
                        If statusRand <= statusChance Then
                            MapNpcPokemon(MapNum, MapNpcNum).Status = PokemonMove(MoveNum).pStatus
                            SendMapNpcPokemonStatus MapNum, MapNpcNum
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If PokemonMove(MoveNum).RecoilDamage > 0 Then
        recoil = PokemonMove(MoveNum).RecoilDamage
        Damage = MapNpcPokemon(MapNum, MapNpcNum).MaxHp * (recoil / 100)
        If Damage > 0 Then
            MapNpcPokemon(MapNum, MapNpcNum).CurHp = MapNpcPokemon(MapNum, MapNpcNum).CurHp - Damage
            SendActionMsg MapNum, "-" & Damage, MapNpcPokemon(MapNum, MapNpcNum).X * 32, MapNpcPokemon(MapNum, MapNpcNum).Y * 32, BrightRed
            If MapNpcPokemon(MapNum, MapNpcNum).CurHp <= 0 Then
                MapNpcPokemon(MapNum, MapNpcNum).CurHp = 0
                SendNpcPokemonVital MapNum, MapNpcNum
                MapNpc(MapNum, MapNpcNum).PokemonAlive(MapNpc(MapNum, MapNpcNum).CurPokemon) = NO
                NpcPokemonCallBack MapNum, MapNpcNum
            Else
                SendNpcPokemonVital MapNum, MapNpcNum
            End If
        End If
    End If
    
    '//Play Sound
    If Not Trim$(PokemonMove(MoveNum).Sound) = "None." Or Not Trim$(PokemonMove(MoveNum).Sound) = vbNullString Then
        SendPlaySound Trim$(PokemonMove(MoveNum).Sound), MapNum
    End If
End Sub

Public Sub NpcPokemonAttackPlayer(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal TargetIndex As Long, ByVal Damage As Long)
Dim DuelIndex As Long

    '//Check Error
    If MapNum <= 0 Or MapNum > MAX_MAP Then Exit Sub
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPC Then Exit Sub
    If MapNpcPokemon(MapNum, MapNpcNum).Num <= 0 Then Exit Sub
    If MapNpc(MapNum, MapNpcNum).CurPokemon <= 0 Then Exit Sub
    If Not IsPlaying(TargetIndex) Then Exit Sub
    If TempPlayer(TargetIndex).UseChar <= 0 Then Exit Sub
    If PlayerPokemon(TargetIndex).Num <= 0 Then Exit Sub
    If PlayerPokemon(TargetIndex).slot <= 0 Then Exit Sub
    
    If Damage >= PlayerPokemons(TargetIndex).Data(PlayerPokemon(TargetIndex).slot).CurHp Then
        '//Defeat
        PlayerPokemons(TargetIndex).Data(PlayerPokemon(TargetIndex).slot).CurHp = 0
        SendActionMsg MapNum, "-" & Damage, PlayerPokemon(TargetIndex).X * 32, PlayerPokemon(TargetIndex).Y * 32, BrightRed
        SendPlayerPokemonVital TargetIndex
        ClearPlayerPokemon TargetIndex
        
        If CountPlayerPokemonAlive(TargetIndex) <= 0 Then
            SendActionMsg MapNum, "Lose!", Player(TargetIndex, TempPlayer(TargetIndex).UseChar).X * 32, Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Y * 32, White
            MapNpc(MapNum, MapNpcNum).InBattle = 0
            NpcPokemonCallBack MapNum, MapNpcNum
            TempPlayer(TargetIndex).InNpcDuel = 0
            TempPlayer(TargetIndex).DuelTime = 0
            TempPlayer(TargetIndex).DuelTimeTmr = 0
            TempPlayer(TargetIndex).WarningTimer = 0
            SendPlayerNpcDuel TargetIndex
        Else
            TempPlayer(TargetIndex).DuelReset = YES
        End If
    Else
        PlayerPokemons(TargetIndex).Data(PlayerPokemon(TargetIndex).slot).CurHp = PlayerPokemons(TargetIndex).Data(PlayerPokemon(TargetIndex).slot).CurHp - Damage
        SendActionMsg MapNum, "-" & Damage, PlayerPokemon(TargetIndex).X * 32, PlayerPokemon(TargetIndex).Y * 32, BrightRed
        
        '//Update
        SendPlayerPokemonVital TargetIndex
    End If
End Sub

Public Sub PlayerAttackNpcPokemon(ByVal Index As Long, ByVal TargetIndex As Long, ByVal Damage As Long)
Dim MapNum As Long

    '//Check Error
    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If PlayerPokemon(Index).Num <= 0 Then Exit Sub
    If TargetIndex <= 0 Or TargetIndex > MAX_MAP_NPC Then Exit Sub
    MapNum = Player(Index, TempPlayer(Index).UseChar).Map
    If MapNpcPokemon(MapNum, TargetIndex).Num <= 0 Then Exit Sub
    If Not TempPlayer(Index).InNpcDuel = TargetIndex Then Exit Sub
    If Not MapNpc(MapNum, TargetIndex).InBattle = Index Then Exit Sub
    If MapNpc(MapNum, TargetIndex).CurPokemon <= 0 Then Exit Sub
    
    If Damage >= MapNpcPokemon(MapNum, TargetIndex).CurHp Then
        '//Defeat
        MapNpcPokemon(MapNum, TargetIndex).CurHp = 0
        SendActionMsg MapNum, "-" & Damage, MapNpcPokemon(MapNum, TargetIndex).X * 32, MapNpcPokemon(MapNum, TargetIndex).Y * 32, BrightRed
        
        MapNpc(MapNum, TargetIndex).PokemonAlive(MapNpc(MapNum, TargetIndex).CurPokemon) = NO
        NpcPokemonCallBack MapNum, TargetIndex
        TempPlayer(Index).DuelTime = 3
        TempPlayer(Index).DuelTimeTmr = GetTickCount + 1000
    Else
        MapNpcPokemon(MapNum, TargetIndex).CurHp = MapNpcPokemon(MapNum, TargetIndex).CurHp - Damage
        SendActionMsg MapNum, "-" & Damage, MapNpcPokemon(MapNum, TargetIndex).X * 32, MapNpcPokemon(MapNum, TargetIndex).Y * 32, BrightRed
        
        '//Update
        SendNpcPokemonVital MapNum, TargetIndex
    End If
End Sub
