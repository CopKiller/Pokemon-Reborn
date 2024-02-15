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

Private Function GetPlayerMapUnlocked(ByVal Index As Long, ByVal MapNum As Long) As Byte
    If Not IsPlaying(Index) Then Exit Function
    If TempPlayer(Index).UseChar <= 0 Then Exit Function

    GetPlayerMapUnlocked = Player(Index, TempPlayer(Index).UseChar).PlayerTravel(MapNum).Unlocked
End Function

Private Sub SetPlayerMapUnlocked(ByVal Index As Long, ByVal MapNum As Long, ByVal Unlocked As Byte)
    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    
    Player(Index, TempPlayer(Index).UseChar).PlayerTravel(MapNum).Unlocked = Unlocked
End Sub

Private Function GetMapTravel(ByVal MapNum As Long) As Boolean
    If Map(MapNum).MapTravel.IsTravel = YES Then GetMapTravel = True
End Function

Public Sub CheckMapTravel(ByVal Index As Long, ByVal MapNum As Long)
    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub

    If GetMapTravel(MapNum) Then
        If GetPlayerMapUnlocked(Index, MapNum) = NO Then
            Call SetPlayerMapUnlocked(Index, MapNum, YES)
            Call SendUpdatePlayerMapTravel(Index, MapNum)

            Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: AddAlert Index, "Você desbloqueou " & Trim$(Map(MapNum).Name), White
            Case LANG_EN: AddAlert Index, "You Unlocked " & Trim$(Map(MapNum).Name), White
            Case LANG_ES: AddAlert Index, "You Unlocked " & Trim$(Map(MapNum).Name), White
            End Select
        End If
    End If
End Sub

Public Sub SendUpdatePlayerMapTravel(ByVal Index As Long, ByVal MapNum As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerTravel

    If GetMapTravel(MapNum) Then
        buffer.WriteLong MapNum
        buffer.WriteByte GetPlayerMapUnlocked(Index, MapNum)
        buffer.WriteString Trim$(Map(MapNum).Name)
        buffer.WriteLong Map(i).MapTravel.CostValue
    End If

    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendUpdatePlayerMapTravelAll(ByVal Index As Long)
    Dim buffer As clsBuffer
    Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerTravel

    For i = 1 To MAX_MAP
        If GetMapTravel(i) Then
            buffer.WriteLong i
            buffer.WriteByte GetPlayerMapUnlocked(Index, i)
            buffer.WriteString Trim$(Map(i).Name)
            buffer.WriteLong Map(i).MapTravel.CostValue
        End If
    Next i

    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub HandlePlayerTravel(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim travelSlot As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub

    If TempPlayer(Index).StorageType > 0 Then
        Exit Sub
    End If
    If TempPlayer(Index).InShop > 0 Then
        Exit Sub
    End If
    If TempPlayer(Index).PlayerRequest > 0 Then
        Exit Sub
    End If
    If TempPlayer(Index).InDuel > 0 Then
        Exit Sub
    End If
    If TempPlayer(Index).InTrade > 0 Then
        Exit Sub
    End If
    If TempPlayer(Index).CurConvoNum > 0 Then
        Exit Sub
    End If
    If TempPlayer(Index).InNpcDuel > 0 Then
        Exit Sub
    End If
    If Player(Index, TempPlayer(Index).UseChar).Action > 0 Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    travelSlot = buffer.ReadLong
    Set buffer = Nothing

    With Player(Index, TempPlayer(Index).UseChar)
        If GetMapTravel(travelSlot) Then
            If GetPlayerMapUnlocked(Index, travelSlot) = YES Then

                If .Money >= Map(travelSlot).MapTravel.CostValue Then

                    If Map(travelSlot).MapTravel.BadgeReq > 0 Then
                        If .Badge(Map(travelSlot).MapTravel.BadgeReq) = NO Then
                            Select Case TempPlayer(Index).CurLanguage
                            Case LANG_PT: AddAlert Index, "Você precisa da insignia de " & Trim$(Map(travelSlot).Name), White
                            Case LANG_EN: AddAlert Index, "Você precisa da insignia de " & Trim$(Map(travelSlot).Name), White
                            Case LANG_ES: AddAlert Index, "Você precisa da insignia de " & Trim$(Map(travelSlot).Name), White
                            End Select
                            Exit Sub
                        End If
                    End If

                    .Money = .Money - Map(travelSlot).MapTravel.CostValue
                    Call PlayerWarp(Index, travelSlot, Map(travelSlot).MapTravel.x, Map(travelSlot).MapTravel.Y, DIR_DOWN)
                    Call SendPlayerCash(Index)
                    
                    Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "Você foi teleportado para " & Trim$(Map(travelSlot).Name), White
                    Case LANG_EN: AddAlert Index, "Você foi teleportado para " & Trim$(Map(travelSlot).Name), White
                    Case LANG_ES: AddAlert Index, "Você foi teleportado para " & Trim$(Map(travelSlot).Name), White
                    End Select
                Else
                    Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "Você precisa de " & Trim$(Map(travelSlot).MapTravel.CostValue) & " para teleportar!", White
                    Case LANG_EN: AddAlert Index, "Você precisa de " & Trim$(Map(travelSlot).MapTravel.CostValue) & " para teleportar!", White
                    Case LANG_ES: AddAlert Index, "Você precisa de " & Trim$(Map(travelSlot).MapTravel.CostValue) & " para teleportar!", White
                    End Select
                End If

            End If
        End If
    End With

End Sub
