Attribute VB_Name = "modMapTravel"
Option Explicit

Public Type PlayerTravelRec
    Unlocked As Byte
End Type

Public Type MapTravelRec
    IsTravel As Byte
    CostValue As Long
    x As Long
    Y As Long
    BadgeReq As Byte
End Type

Private Function GetPlayerMapUnlocked(ByVal index As Long, ByVal MapNum As Long) As Byte
    If Not IsPlaying(index) Then Exit Function
    If TempPlayer(index).UseChar <= 0 Then Exit Function

    GetPlayerMapUnlocked = Player(index, TempPlayer(index).UseChar).PlayerTravel(MapNum).Unlocked
End Function

Private Sub SetPlayerMapUnlocked(ByVal index As Long, ByVal MapNum As Long, ByVal Unlocked As Byte)
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    
    Player(index, TempPlayer(index).UseChar).PlayerTravel(MapNum).Unlocked = Unlocked
End Sub

Private Function GetMapTravel(ByVal MapNum As Long) As Boolean
    If Map(MapNum).MapTravel.IsTravel = YES Then GetMapTravel = True
End Function

Public Sub CheckMapTravel(ByVal index As Long, ByVal MapNum As Long)
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub

    If GetMapTravel(MapNum) Then
        If GetPlayerMapUnlocked(index, MapNum) = NO Then
            Call SetPlayerMapUnlocked(index, MapNum, YES)
            Call SendUpdatePlayerMapTravel(index, MapNum)

            Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: AddAlert index, "Você desbloqueou " & Trim$(Map(MapNum).Name), White
            Case LANG_EN: AddAlert index, "You Unlocked " & Trim$(Map(MapNum).Name), White
            Case LANG_ES: AddAlert index, "You Unlocked " & Trim$(Map(MapNum).Name), White
            End Select
        End If
    End If
End Sub

Public Sub SendUpdatePlayerMapTravel(ByVal index As Long, ByVal MapNum As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerTravel

    If GetMapTravel(MapNum) Then
        buffer.WriteLong MapNum
        buffer.WriteByte GetPlayerMapUnlocked(index, MapNum)
        buffer.WriteString Trim$(Map(MapNum).Name)
        buffer.WriteLong Map(MapNum).MapTravel.CostValue
    End If

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendUpdatePlayerMapTravelAll(ByVal index As Long)
    Dim buffer As clsBuffer
    Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerTravel

    For i = 1 To MAX_MAP
        If GetMapTravel(i) Then
            buffer.WriteLong i
            buffer.WriteByte GetPlayerMapUnlocked(index, i)
            buffer.WriteString Trim$(Map(i).Name)
            buffer.WriteLong Map(i).MapTravel.CostValue
        End If
    Next i

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub HandlePlayerTravel(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim travelSlot As Long

    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub

    If TempPlayer(index).StorageType > 0 Then
        Exit Sub
    End If
    If TempPlayer(index).InShop > 0 Then
        Exit Sub
    End If
    If TempPlayer(index).PlayerRequest > 0 Then
        Exit Sub
    End If
    If TempPlayer(index).InDuel > 0 Then
        Exit Sub
    End If
    If TempPlayer(index).InTrade > 0 Then
        Exit Sub
    End If
    If TempPlayer(index).CurConvoNum > 0 Then
        Exit Sub
    End If
    If TempPlayer(index).InNpcDuel > 0 Then
        Exit Sub
    End If
    If Player(index, TempPlayer(index).UseChar).Action > 0 Then
        Exit Sub
    End If
    If PlayerPokemon(index).Num <> 0 Then
        Select Case TempPlayer(index).CurLanguage
        Case LANG_PT: AddAlert index, "Remova seu pokemon do mapa antes!", White
        Case LANG_EN: AddAlert index, "Remova seu pokemon do mapa antes!", White
        Case LANG_ES: AddAlert index, "Remova seu pokemon do mapa antes!", White
        End Select
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    travelSlot = buffer.ReadLong
    Set buffer = Nothing

    With Player(index, TempPlayer(index).UseChar)
        If GetMapTravel(travelSlot) Then
            If GetPlayerMapUnlocked(index, travelSlot) = YES Then

                If .Money >= Map(travelSlot).MapTravel.CostValue Then

                    If Map(travelSlot).MapTravel.BadgeReq > 0 Then
                        If .Badge(Map(travelSlot).MapTravel.BadgeReq) = NO Then
                            Select Case TempPlayer(index).CurLanguage
                            Case LANG_PT: AddAlert index, "Você precisa da insignia de " & Trim$(Map(travelSlot).Name), White
                            Case LANG_EN: AddAlert index, "Você precisa da insignia de " & Trim$(Map(travelSlot).Name), White
                            Case LANG_ES: AddAlert index, "Você precisa da insignia de " & Trim$(Map(travelSlot).Name), White
                            End Select
                            Exit Sub
                        End If
                    End If

                    .Money = .Money - Map(travelSlot).MapTravel.CostValue
                    Call PlayerWarp(index, travelSlot, Map(travelSlot).MapTravel.x, Map(travelSlot).MapTravel.Y, DIR_DOWN)
                    Call SendPlayerCash(index)

                    Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "Você foi teleportado para " & Trim$(Map(travelSlot).Name), White
                    Case LANG_EN: AddAlert index, "Você foi teleportado para " & Trim$(Map(travelSlot).Name), White
                    Case LANG_ES: AddAlert index, "Você foi teleportado para " & Trim$(Map(travelSlot).Name), White
                    End Select
                Else
                    Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "Você precisa de " & Trim$(Map(travelSlot).MapTravel.CostValue) & " para teleportar!", White
                    Case LANG_EN: AddAlert index, "Você precisa de " & Trim$(Map(travelSlot).MapTravel.CostValue) & " para teleportar!", White
                    Case LANG_ES: AddAlert index, "Você precisa de " & Trim$(Map(travelSlot).MapTravel.CostValue) & " para teleportar!", White
                    End Select
                End If

            End If
        End If
    End With

End Sub
