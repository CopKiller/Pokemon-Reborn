Attribute VB_Name = "modVip"
Option Explicit

Public Type PlayerVipRec
    vipType As Byte
    VipDate As Date
    VipDays As Long
End Type

Public Enum EnumVipType
    None = 0
    VipSilver
    VipGold
End Enum

Private CheckVipTimer As Long
Private Const CheckcVipInterval As Long = 3600000 ' A cada 1 hora verifica o vip de todos os jogadores

Public Function GetPlayerVipStatus(ByVal index As Long) As Byte
    GetPlayerVipStatus = Player(index, TempPlayer(index).UseChar).Vip.vipType
End Function

Private Sub SetPlayerVipStatus(ByVal index As Long, ByVal vipValue As EnumVipType)
    Player(index, TempPlayer(index).UseChar).Vip.vipType = vipValue
End Sub

Private Function GetPlayerVipDate(ByVal index As Long) As Date
    GetPlayerVipDate = Player(index, TempPlayer(index).UseChar).Vip.VipDate
End Function

Private Sub SetPlayerVipDate(ByVal index As Long, ByVal dateValue As Date)
    Player(index, TempPlayer(index).UseChar).Vip.VipDate = dateValue
End Sub

'//Vip days usado para comparações na lógica
Private Function GetPlayerVipDays(ByVal index As Long) As Long
    GetPlayerVipDays = Player(index, TempPlayer(index).UseChar).Vip.VipDays
End Function

'//Vip days para obter quantos dias faltam para acabar o vip do jogador
Public Function GetPlayerVipDaysNow(ByVal index As Long) As Long
    GetPlayerVipDaysNow = Player(index, TempPlayer(index).UseChar).Vip.VipDays - DateDiff("d", GetPlayerVipDate(index), Date)
End Function

Private Sub SetPlayerVipDays(ByVal index As Long, ByVal daysValue As Long)
    Player(index, TempPlayer(index).UseChar).Vip.VipDays = daysValue
End Sub

Public Sub CheckVipLoop()
    Dim index As Long
    If CheckVipTimer <= GetTickCount Then
        For index = 1 To Player_HighIndex
            If IsPlaying(index) Then
                ' Check Vip
                If GetPlayerVipStatus(index) > EnumVipType.None Then
                    If DateDiff("d", GetPlayerVipDate(index), Date) >= GetPlayerVipDays(index) Then
                        Call SetPlayerVipStatus(index, None)
                        Call SetPlayerVipDays(index, 0)
                        Call SendPlayerData(index)
                        Select Case TempPlayer(index).CurLanguage
                        Case LANG_PT: SendPlayerMsg index, "Seus dias de VIP acabaram...", BrightRed
                        Case LANG_EN: SendPlayerMsg index, "Seus dias de VIP acabaram...", BrightRed
                        Case LANG_ES: SendPlayerMsg index, "Seus dias de VIP acabaram...", BrightRed
                        End Select
                    End If
                End If
            End If
        Next index

        CheckVipTimer = GetTickCount + CheckcVipInterval
    End If
End Sub

Public Sub CheckVipJoinGame(ByVal index As Long)
' Check Vip
    If GetPlayerVipStatus(index) > EnumVipType.None Then
        If DateDiff("d", GetPlayerVipDate(index), Date) < GetPlayerVipDays(index) Then
        
            Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: SendPlayerMsg index, "Obrigado por adquirir seu VIP, bom jogo!", White
            Case LANG_EN: SendPlayerMsg index, "Obrigado por adquirir seu VIP, bom jogo!", White
            Case LANG_ES: SendPlayerMsg index, "Obrigado por adquirir seu VIP, bom jogo!", White
            End Select
            
        ElseIf DateDiff("d", GetPlayerVipDate(index), Date) >= GetPlayerVipDays(index) Then
            Call SetPlayerVipStatus(index, None)
            Call SetPlayerVipDays(index, 0)
            
            Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: SendPlayerMsg index, "Seus dias de VIP acabaram... bom jogo!", BrightRed
            Case LANG_EN: SendPlayerMsg index, "Seus dias de VIP acabaram... bom jogo!", BrightRed
            Case LANG_ES: SendPlayerMsg index, "Seus dias de VIP acabaram... bom jogo!", BrightRed
            End Select
        End If
    End If
End Sub

Public Function AddVip(ByVal index As Long, ByVal vipType As EnumVipType, ByVal daysValue As Long) As Boolean

    If IsPlaying(index) Then
        If GetPlayerVipStatus(index) > EnumVipType.None Then

            If GetPlayerVipStatus(index) <> vipType Then
                Select Case TempPlayer(index).CurLanguage
                Case LANG_PT: SendPlayerMsg index, "Vip diferente do atual, aguarde o seu finalizar!", White
                Case LANG_EN: SendPlayerMsg index, "Vip diferente do atual, aguarde o seu finalizar!", White
                Case LANG_ES: SendPlayerMsg index, "Vip diferente do atual, aguarde o seu finalizar!", White
                End Select
                Exit Function
            Else
                Call SetPlayerVipDays(index, GetPlayerVipDays(index) + daysValue)
                Call SendPlayerData(index)
                AddVip = True
                Select Case TempPlayer(index).CurLanguage
                Case LANG_PT: SendPlayerMsg index, "Foi acrescentado em seu vip " & daysValue & " dias!", White
                Case LANG_EN: SendPlayerMsg index, "Foi acrescentado em seu vip " & daysValue & " dias!", White
                Case LANG_ES: SendPlayerMsg index, "Foi acrescentado em seu vip " & daysValue & " dias!", White
                End Select
            End If
        Else
            Call SetPlayerVipStatus(index, vipType)
            Call SetPlayerVipDate(index, Date)
            Call SetPlayerVipDays(index, daysValue)
            Call SendPlayerData(index)
            AddVip = True
            Call SendGlobalMsg("Player " & GetPlayerName(index) & " became VIP!", Green)
        End If
    End If

    Exit Function
End Function
