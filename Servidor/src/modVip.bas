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
    
    VipCount
End Enum

'//Vip Settings
Public VipSettings(0 To 2) As VipSettingsRec

Private Type VipSettingsRec
    VipExp As Integer
    VipCoin As Integer
    VipDrop As Integer
    VipShopPrice As Integer
    VipDeathPenalty As Integer
End Type

'//Check vip Loop
Private CheckVipTimer As Long
Private Const CheckcVipInterval As Long = 3600000 ' A cada 1 hora verifica o vip de todos os jogadores

Public Function GetPlayerVipStatus(ByVal Index As Long) As Byte
    GetPlayerVipStatus = Player(Index, TempPlayer(Index).UseChar).Vip.vipType
End Function

Private Sub SetPlayerVipStatus(ByVal Index As Long, ByVal vipValue As EnumVipType)
    Player(Index, TempPlayer(Index).UseChar).Vip.vipType = vipValue
End Sub

Private Function GetPlayerVipDate(ByVal Index As Long) As Date
    GetPlayerVipDate = Player(Index, TempPlayer(Index).UseChar).Vip.VipDate
End Function

Private Sub SetPlayerVipDate(ByVal Index As Long, ByVal dateValue As Date)
    Player(Index, TempPlayer(Index).UseChar).Vip.VipDate = dateValue
End Sub

'//Vip days usado para comparações na lógica
Private Function GetPlayerVipDays(ByVal Index As Long) As Long
    GetPlayerVipDays = Player(Index, TempPlayer(Index).UseChar).Vip.VipDays
End Function

'//Vip days para obter quantos dias faltam para acabar o vip do jogador
Public Function GetPlayerVipDaysNow(ByVal Index As Long) As Long
    GetPlayerVipDaysNow = Player(Index, TempPlayer(Index).UseChar).Vip.VipDays - DateDiff("d", GetPlayerVipDate(Index), Date)
End Function

Private Sub SetPlayerVipDays(ByVal Index As Long, ByVal daysValue As Long)
    Player(Index, TempPlayer(Index).UseChar).Vip.VipDays = daysValue
End Sub




Public Function GetVipDiscountValue(ByVal Index As Long, ByVal value As Long) As Long
    GetVipDiscountValue = value

    If GetPlayerVipStatus(Index) > EnumVipType.None Then
        If value > 0 Then

            value = ((value / 100) * VipSettings(GetPlayerVipStatus(Index)).VipShopPrice)

            If value <= GetVipDiscountValue Then
                GetVipDiscountValue = value
            End If
        End If
    End If
End Function


Public Sub CheckVipLoop()
    Dim Index As Long
    If CheckVipTimer <= GetTickCount Then
        For Index = 1 To Player_HighIndex
            If IsPlaying(Index) Then
                ' Check Vip
                If GetPlayerVipStatus(Index) > EnumVipType.None Then
                    If DateDiff("d", GetPlayerVipDate(Index), Date) >= GetPlayerVipDays(Index) Then
                        Call SetPlayerVipStatus(Index, None)
                        Call SetPlayerVipDays(Index, 0)
                        Call SendPlayerData(Index)
                        Select Case TempPlayer(Index).CurLanguage
                        Case LANG_PT: SendPlayerMsg Index, "Seus dias de VIP acabaram...", BrightRed
                        Case LANG_EN: SendPlayerMsg Index, "Seus dias de VIP acabaram...", BrightRed
                        Case LANG_ES: SendPlayerMsg Index, "Seus dias de VIP acabaram...", BrightRed
                        End Select
                    End If
                End If
            End If
        Next Index

        CheckVipTimer = GetTickCount + CheckcVipInterval
    End If
End Sub

Public Sub CheckVipJoinGame(ByVal Index As Long)
' Check Vip
    If GetPlayerVipStatus(Index) > EnumVipType.None Then
        If DateDiff("d", GetPlayerVipDate(Index), Date) < GetPlayerVipDays(Index) Then
        
            Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: SendPlayerMsg Index, "Obrigado por adquirir seu VIP, bom jogo!", White
            Case LANG_EN: SendPlayerMsg Index, "Obrigado por adquirir seu VIP, bom jogo!", White
            Case LANG_ES: SendPlayerMsg Index, "Obrigado por adquirir seu VIP, bom jogo!", White
            End Select
            
        ElseIf DateDiff("d", GetPlayerVipDate(Index), Date) >= GetPlayerVipDays(Index) Then
            Call SetPlayerVipStatus(Index, None)
            Call SetPlayerVipDays(Index, 0)
            
            Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: SendPlayerMsg Index, "Seus dias de VIP acabaram... bom jogo!", BrightRed
            Case LANG_EN: SendPlayerMsg Index, "Seus dias de VIP acabaram... bom jogo!", BrightRed
            Case LANG_ES: SendPlayerMsg Index, "Seus dias de VIP acabaram... bom jogo!", BrightRed
            End Select
        End If
    End If
End Sub

Public Function AddVip(ByVal Index As Long, ByVal vipType As EnumVipType, ByVal daysValue As Long) As Boolean

    If IsPlaying(Index) Then
        If GetPlayerVipStatus(Index) > EnumVipType.None Then

            If GetPlayerVipStatus(Index) <> vipType Then
                Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: SendPlayerMsg Index, "Vip diferente do atual, aguarde o seu finalizar!", White
                Case LANG_EN: SendPlayerMsg Index, "Vip diferente do atual, aguarde o seu finalizar!", White
                Case LANG_ES: SendPlayerMsg Index, "Vip diferente do atual, aguarde o seu finalizar!", White
                End Select
                Exit Function
            Else
                Call SetPlayerVipDays(Index, GetPlayerVipDays(Index) + daysValue)
                Call SendPlayerData(Index)
                AddVip = True
                Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: SendPlayerMsg Index, "Foi acrescentado em seu vip " & daysValue & " dias!", White
                Case LANG_EN: SendPlayerMsg Index, "Foi acrescentado em seu vip " & daysValue & " dias!", White
                Case LANG_ES: SendPlayerMsg Index, "Foi acrescentado em seu vip " & daysValue & " dias!", White
                End Select
            End If
        Else
            Call SetPlayerVipStatus(Index, vipType)
            Call SetPlayerVipDate(Index, Date)
            Call SetPlayerVipDays(Index, daysValue)
            Call SendVipAdvantageTo(Index)
            Call SendPlayerData(Index)
            AddVip = True
            Call SendGlobalMsg("Player " & GetPlayerName(Index) & " became VIP!", Green)
        End If
    End If

    Exit Function
End Function

'//Vip Settings do editor do servidor.
Public Sub LoadVipSettings()
    Dim filename As String, i As Long

    filename = App.Path & "\data\vipsettings.ini"

    If Not FileExist(filename) Then

        For i = 0 To EnumVipType.VipCount - 1
            With VipSettings(i)
                .VipExp = (i * 10)
                .VipCoin = (i * 10)
                .VipDrop = (i * 10)
                .VipShopPrice = 100
                .VipDeathPenalty = 100
            End With
        Next i

        SaveVipSettings
        Exit Sub
    End If

    For i = 0 To EnumVipType.VipCount - 1
        With VipSettings(i)
            .VipExp = CInt(GetVar(filename, CStr(i), "VipExp"))
            .VipCoin = CInt(GetVar(filename, CStr(i), "VipCoin"))
            .VipDrop = CInt(GetVar(filename, CStr(i), "VipDrop"))
            .VipShopPrice = CInt(GetVar(filename, CStr(i), "VipShopPrice"))
            .VipDeathPenalty = CInt(GetVar(filename, CStr(i), "VipDeathPenalty"))
        End With
    Next i
End Sub

Public Sub SaveVipSettings()
    Dim filename As String, i As Long
    filename = App.Path & "\data\vipsettings.ini"
    
    If FileExist(filename) Then
        For i = 0 To EnumVipType.VipCount - 1
            PutVar filename, CStr(i), "VipExp", CStr(VipSettings(i).VipExp)
            PutVar filename, CStr(i), "VipCoin", CStr(VipSettings(i).VipCoin)
            PutVar filename, CStr(i), "VipDrop", CStr(VipSettings(i).VipDrop)
            PutVar filename, CStr(i), "VipShopPrice", CStr(VipSettings(i).VipShopPrice)
            PutVar filename, CStr(i), "VipDeathPenalty", CStr(VipSettings(i).VipDeathPenalty)
        Next i
    End If
End Sub


Sub SendVipAdvantageTo(ByVal Index As Long)
    Dim buffer As clsBuffer
    Dim i As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub

    If GetPlayerVipStatus(Index) > EnumVipType.None Then
    
        Set buffer = New clsBuffer
        buffer.WriteLong SVipAdvantage
        
        buffer.WriteInteger VipSettings(GetPlayerVipStatus(Index)).VipExp
        buffer.WriteInteger VipSettings(GetPlayerVipStatus(Index)).VipCoin
        buffer.WriteInteger VipSettings(GetPlayerVipStatus(Index)).VipDrop
        buffer.WriteInteger 100 - VipSettings(GetPlayerVipStatus(Index)).VipShopPrice
        buffer.WriteInteger 100 - VipSettings(GetPlayerVipStatus(Index)).VipDeathPenalty
        
        SendDataTo Index, buffer.ToArray()

        buffer.Flush: Set buffer = Nothing
    End If
End Sub

Sub SendVipAdvantageToAll()
    Dim buffer As clsBuffer
    Dim Index As Long

    For Index = 1 To Player_HighIndex
        If IsPlaying(Index) And TempPlayer(Index).UseChar > 0 Then

            If GetPlayerVipStatus(Index) > EnumVipType.None Then

                Set buffer = New clsBuffer
                buffer.WriteLong SVipAdvantage

                buffer.WriteInteger VipSettings(GetPlayerVipStatus(Index)).VipExp
                buffer.WriteInteger VipSettings(GetPlayerVipStatus(Index)).VipCoin
                buffer.WriteInteger VipSettings(GetPlayerVipStatus(Index)).VipDrop
                buffer.WriteInteger 100 - VipSettings(GetPlayerVipStatus(Index)).VipShopPrice
                buffer.WriteInteger 100 - VipSettings(GetPlayerVipStatus(Index)).VipDeathPenalty

                SendDataTo Index, buffer.ToArray()

                buffer.Flush: Set buffer = Nothing
            End If
        End If
    Next Index
End Sub
