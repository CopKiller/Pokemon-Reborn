Attribute VB_Name = "modSkins"
Option Explicit

Public Type PlayerSkinRec
    skinNum As Long
    LastPlayerSpriteNum As Long
End Type

Public Sub UseSkin(index, ByVal data1 As Long)
    If GetPlayerSkin(index) > 0 Then
        Call ClearPlayerSkin(index)
    Else
        Call SetPlayerOriginalSprite(index, GetPlayerSprite(index))
        Call SetPlayerSkin(index, data1)
    End If
    
    Call SendPlayerData(index)
End Sub

Private Function GetPlayerSprite(ByVal index As Long) As Long
    If Not IsPlaying(index) Then Exit Function
    If TempPlayer(index).UseChar <= 0 Then Exit Function
    
    GetPlayerSprite = Player(index, TempPlayer(index).UseChar).Sprite
End Function

Private Sub ClearPlayerSkin(ByVal index As Long)
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    
    Call SetPlayerSkin(index, 0)
End Sub

Private Sub SetPlayerSkin(ByVal index As Long, ByVal skinNum As Long)
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    
    Player(index, TempPlayer(index).UseChar).Skin.skinNum = skinNum
End Sub

Private Function GetPlayerSkin(ByVal index As Long) As Long
    If Not IsPlaying(index) Then Exit Function
    If TempPlayer(index).UseChar <= 0 Then Exit Function
    
    GetPlayerSkin = Player(index, TempPlayer(index).UseChar).Skin.skinNum
End Function

Public Function GetPlayerHaveSkinOrSprite(ByVal index As Long) As Long
    If GetPlayerSkin(index) > 0 Then
        GetPlayerHaveSkinOrSprite = GetPlayerSkin(index)
    Else
        GetPlayerHaveSkinOrSprite = Player(index, TempPlayer(index).UseChar).Sprite
    End If
End Function

Public Sub SetPlayerOriginalSprite(ByVal index As Long, ByVal spriteNum As Long)
    If Not IsPlaying(index) Then Exit Sub
    If TempPlayer(index).UseChar <= 0 Then Exit Sub
    
    Player(index, TempPlayer(index).UseChar).Skin.LastPlayerSpriteNum = spriteNum
End Sub

Public Function GetPlayerOriginalSprite(ByVal index As Long) As Long
    If Not IsPlaying(index) Then Exit Function
    If TempPlayer(index).UseChar <= 0 Then Exit Function
    
    GetPlayerOriginalSprite = Player(index, TempPlayer(index).UseChar).Skin.LastPlayerSpriteNum
End Function
