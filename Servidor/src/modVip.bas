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
            Call SendVipAdvantageTo(index)
            Call SendPlayerData(index)
            AddVip = True
            Call SendGlobalMsg("Player " & GetPlayerName(index) & " became VIP!", Green)
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


Sub SendVipAdvantageTo(ByVal index As Long)
    Dim buffer As clsBuffer
    Dim i As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub

    If GetPlayerVipStatus(index) > EnumVipType.None Then
    
        Set buffer = New clsBuffer
        buffer.WriteLong SVipAdvantage
        
        buffer.WriteInteger VipSettings(GetPlayerVipStatus(index)).VipExp
        buffer.WriteInteger VipSettings(GetPlayerVipStatus(index)).VipCoin
        buffer.WriteInteger VipSettings(GetPlayerVipStatus(index)).VipDrop
        buffer.WriteInteger 100 - VipSettings(GetPlayerVipStatus(index)).VipShopPrice
        buffer.WriteInteger 100 - VipSettings(GetPlayerVipStatus(index)).VipDeathPenalty
        
        SendDataTo index, buffer.ToArray()

        buffer.Flush: Set buffer = Nothing
    End If
End Sub

Sub SendVipAdvantageToAll()
    Dim buffer As clsBuffer
    Dim index As Long

    For index = 1 To Player_HighIndex
        If IsPlaying(index) And TempPlayer(index).UseChar > 0 Then

            If GetPlayerVipStatus(index) > EnumVipType.None Then

                Set buffer = New clsBuffer
                buffer.WriteLong SVipAdvantage

                buffer.WriteInteger VipSettings(GetPlayerVipStatus(index)).VipExp
                buffer.WriteInteger VipSettings(GetPlayerVipStatus(index)).VipCoin
                buffer.WriteInteger VipSettings(GetPlayerVipStatus(index)).VipDrop
                buffer.WriteInteger 100 - VipSettings(GetPlayerVipStatus(index)).VipShopPrice
                buffer.WriteInteger 100 - VipSettings(GetPlayerVipStatus(index)).VipDeathPenalty

                SendDataTo index, buffer.ToArray()

                buffer.Flush: Set buffer = Nothing
            End If
        End If
    Next index
End Sub
