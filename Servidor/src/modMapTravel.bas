Attribute VB_Name = "modMapTravel"
Option Explicit

Public Type PlayerTravelRec
    Unlocked As Byte
End Type

Public Type MapTravelRec
    IsTravel As Byte
    CostValue As Long
    X As Long
    Y As Long
End Type

Private Function GetPlayerMapUnlocked(ByVal index As Long, ByVal mapNum As Long) As Byte
    If Not IsPlaying(index) Then Exit Function
    If TempPlayer(index).UseChar <= 0 Then Exit Function

    GetPlayerMapUnlocked = Player(index, TempPlayer(index).UseChar).PlayerTravel(mapNum).Unlocked
End Function

Private Sub SetPlayerMapUnlocked(ByVal index As Long, ByVal mapNum As Long, ByVal Unlocked As Byte)
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    
    Player(index, TempPlayer(index).UseChar).PlayerTravel(mapNum).Unlocked = Unlocked
End Sub

Private Function GetMapTravel(ByVal mapNum As Long) As Boolean
    If Map(mapNum).MapTravel.IsTravel = YES Then GetMapTravel = True
End Function

Public Sub CheckMapTravel(ByVal index As Long, ByVal mapNum As Long)
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub

    If GetMapTravel(mapNum) Then
        If GetPlayerMapUnlocked(index, mapNum) = NO Then
            Call SetPlayerMapUnlocked(index, mapNum, False)
            Call SendUpdatePlayerMapTravel(index, mapNum)

            Select Case TempPlayer(index).CurLanguage
            Case LANG_PT: AddAlert index, "Você desbloqueou " & Trim$(Map(mapNum).Name), White
            Case LANG_EN: AddAlert index, "You Unlocked " & Trim$(Map(mapNum).Name), White
            Case LANG_ES: AddAlert index, "You Unlocked " & Trim$(Map(mapNum).Name), White
            End Select
        End If
    End If
End Sub

Public Sub SendUpdatePlayerMapTravel(ByVal index As Long, ByVal mapNum As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerTravel

    If GetMapTravel(mapNum) Then
        buffer.WriteLong mapNum
        buffer.WriteByte GetPlayerMapUnlocked(index, mapNum)
        buffer.WriteString Trim$(Map(mapNum).Name)
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

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    travelSlot = buffer.ReadLong
    Set buffer = Nothing

    With Player(index, TempPlayer(index).UseChar)
        If GetMapTravel(travelSlot) Then
            If GetPlayerMapUnlocked(index, travelSlot) = YES Then

                If .Money >= Map(travelSlot).MapTravel.CostValue Then
                    .Money = .Money - Map(travelSlot).MapTravel.CostValue

                    Select Case TempPlayer(index).CurLanguage
                    Case LANG_PT: AddAlert index, "Você foi teleportado para " & Trim$(Map(travelSlot).Name), White
                    Case LANG_EN: AddAlert index, "Você foi teleportado para " & Trim$(Map(travelSlot).Name), White
                    Case LANG_ES: AddAlert index, "Você foi teleportado para " & Trim$(Map(travelSlot).Name), White
                    End Select
                    
                    Call PlayerWarp(index, travelSlot, Map(travelSlot).MapTravel.X, Map(travelSlot).MapTravel.Y, DIR_DOWN)
                    Call SendPlayerCash(index)
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
