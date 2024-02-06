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

Public Function GetPlayerVipStatus(ByVal index As Long) As Byte
    GetPlayerVipStatus = Player(index).Vip.vipType
End Function

Public Sub SetPlayerVipStatus(ByVal index As Long, ByVal vipValue As EnumVipType)
    Player(index).Vip.vipType = vipValue
End Sub

Public Function GetPlayerVipDays(ByVal index As Long) As Long
    GetPlayerVipDays = Player(index).Vip.VipDays
End Function

Public Sub SetPlayerVipDays(ByVal index As Long, ByVal daysValue As Long)
    Player(index).Vip.VipDays = daysValue
End Sub
