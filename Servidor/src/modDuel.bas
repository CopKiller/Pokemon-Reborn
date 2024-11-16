Attribute VB_Name = "modDuel"
Option Explicit



Public Sub SendPlayerDuel(ByVal Index As Long)
    Dim buffer As clsBuffer
    Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerNpcDuel
    
    buffer.WriteByte TempPlayer(Index).InDuelTargetType
    buffer.WriteLong TempPlayer(Index).InDuel
    
    If TempPlayer(Index).InDuelTargetType = TARGET_TYPE_MAPPOKEMON Then
        
        For i = 1 To Pokemon_HighIndex
            If MapPokemon(i).targetType = TARGET_TYPE_PLAYER Or MapPokemon(i).targetType = TARGET_TYPE_PLAYERPOKEMON Then
                If MapPokemon(i).TargetIndex = Index Then
                    buffer.WriteLong i
                End If
            End If
        Next i
        
    End If
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub PlayerWinToNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long)
    Dim NpcNum As Long, MapNum As Long

    MapNum = Player(Attacker, TempPlayer(Attacker).UseChar).Map

    NpcNum = MapNpc(MapNum, MapNpcNum).Num

    SendActionMsg MapNum, "Win!", Player(Attacker, TempPlayer(Attacker).UseChar).x * 32, Player(Attacker, TempPlayer(Attacker).UseChar).Y * 32, White
    Select Case TempPlayer(Attacker).CurLanguage
    Case LANG_PT: AddAlert Attacker, "You win on a duel!", White
    Case LANG_EN: AddAlert Attacker, "You win on a duel!", White
    Case LANG_ES: AddAlert Attacker, "You win on a duel!", White
    End Select

    TempPlayer(Attacker).InDuel = 0
    TempPlayer(Attacker).InDuelTargetType = 0
    TempPlayer(Attacker).DuelTime = 0
    TempPlayer(Attacker).DuelTimeTmr = 0
    TempPlayer(Attacker).WarningTimer = 0
    SendPlayerDuel Attacker
    '//Send Reward
    If Npc(NpcNum).Reward > 0 Then
        Player(Attacker, TempPlayer(Attacker).UseChar).Money = Player(Attacker, TempPlayer(Attacker).UseChar).Money + Npc(NpcNum).Reward
        If Player(Attacker, TempPlayer(Attacker).UseChar).Money >= MAX_MONEY Then
            Player(Attacker, TempPlayer(Attacker).UseChar).Money = MAX_MONEY
        End If
        SendPlayerData Attacker
        AddAlert Attacker, "You got $" & Npc(NpcNum).Reward, White
        If Npc(NpcNum).RewardExp > 0 Then
            GivePlayerExp Attacker, Npc(NpcNum).RewardExp
        End If
    End If
    If Npc(NpcNum).WinEvent > 0 Then
        TempPlayer(Attacker).CurConvoNum = Npc(NpcNum).WinEvent
        TempPlayer(Attacker).CurConvoData = 0    '//Always start at 0
        TempPlayer(Attacker).CurConvoNpc = NpcNum
        TempPlayer(Attacker).CurConvoMapNpc = MapNpcNum
        ProcessConversation Attacker, TempPlayer(Attacker).CurConvoNum, TempPlayer(Attacker).CurConvoData, TempPlayer(Attacker).CurConvoNpc
    End If

    '//Não pode rebatalhar se ganhar.
    Player(Attacker, TempPlayer(Attacker).UseChar).NpcBattledDay(NpcNum).Win = NO
    
    If Npc(NpcNum).Rebatle = REBATLE_NONE Or Npc(NpcNum).Rebatle = REBATLE_LOSE Then    '//Não pode rebatalhar neste dia, se ganhar
            Player(Attacker, TempPlayer(Attacker).UseChar).NpcBattledDay(NpcNum).NpcBattledAt = Day(Date)
            Player(Attacker, TempPlayer(Attacker).UseChar).NpcBattledMonth(NpcNum).NpcBattledAt = Month(Date)
    ElseIf Npc(NpcNum).Rebatle = REBATLE_NEVER Then
        Player(Attacker, TempPlayer(Attacker).UseChar).NpcBattledDay(NpcNum).Win = YES
    End If
End Sub

Public Sub PlayerLoseToNpc(ByVal Victim As Long, ByVal MapNpcNum As Long)
    Dim NpcNum As Long, MapNum As Long
    
    MapNum = Player(Victim, TempPlayer(Victim).UseChar).Map

    NpcNum = MapNpc(MapNum, MapNpcNum).Num
    
    If MapNpcNum > 0 Then
        MapNpc(MapNum, MapNpcNum).InBattle = 0
        NpcPokemonCallBack MapNum, MapNpcNum
        SendActionMsg MapNum, "Lose!", Player(Victim, TempPlayer(Victim).UseChar).x * 32, Player(Victim, TempPlayer(Victim).UseChar).Y * 32, White
        TempPlayer(Victim).InDuel = 0
        TempPlayer(Victim).InDuelTargetType = 0
        TempPlayer(Victim).DuelTime = 0
        TempPlayer(Victim).DuelTimeTmr = 0
        TempPlayer(Victim).WarningTimer = 0
        SendPlayerDuel Victim

        If Npc(NpcNum).Rebatle = REBATLE_NONE Then    '//Não pode rebatalhar se perder
            Player(Victim, TempPlayer(Victim).UseChar).NpcBattledDay(NpcNum).NpcBattledAt = Day(Date)
            Player(Victim, TempPlayer(Victim).UseChar).NpcBattledMonth(NpcNum).NpcBattledAt = Month(Date)
        ElseIf Npc(NpcNum).Rebatle = REBATLE_LOSE Or Npc(NpcNum).Rebatle = REBATLE_NEVER Then    '//Pode rebatalhar se perder
            Player(Victim, TempPlayer(Victim).UseChar).NpcBattledDay(NpcNum).NpcBattledAt = 0
            Player(Victim, TempPlayer(Victim).UseChar).NpcBattledMonth(NpcNum).NpcBattledAt = 0
        End If
    End If
End Sub
