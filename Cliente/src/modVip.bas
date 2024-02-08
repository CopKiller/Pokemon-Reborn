Attribute VB_Name = "modVip"
Option Explicit

Public VipAdvantage As VipAdvantageRec

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

Private Type VipAdvantageRec
    ExpValue As Integer
    Exp As String
    
    CoinValue As Integer
    Coin As String
    
    DropValue As Integer
    Drop As String
    
    ShopPriceValue As Integer
    ShopPrice As String
    
    DeathPenaltyValue As Integer
    DeathPenalty As String
End Type

Public Function GetPlayerVipStatus(ByVal Index As Long) As Byte
    GetPlayerVipStatus = Player(Index).Vip.vipType
End Function

Public Sub SetPlayerVipStatus(ByVal Index As Long, ByVal vipValue As EnumVipType)
    Player(Index).Vip.vipType = vipValue
End Sub

Public Function GetPlayerVipDays(ByVal Index As Long) As Long
    GetPlayerVipDays = Player(Index).Vip.VipDays
End Function

Public Sub SetPlayerVipDays(ByVal Index As Long, ByVal daysValue As Long)
    Player(Index).Vip.VipDays = daysValue
End Sub
