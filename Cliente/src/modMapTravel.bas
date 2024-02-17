Attribute VB_Name = "modMapTravel"
Option Explicit

Public Type PlayerTravelRec
    Unlocked As Byte
    mapName As String
    costValue As Long
    
    'Archive Data, position of icon, etc...
    DataExist As Boolean
    SrcPosX As Long
    SrcPosY As Long
    SrcWidth As Long
    SrcHeight As Long
    IconPosX As Long
    IconPosY As Long
End Type

Public Type MapTravelRec
    IsTravel As Byte
    costValue As Long
    X As Long
    y As Long
    BadgeReq As Byte
End Type

Public Function GetPlayerMapUnlocked(ByVal mapNum As Long) As Boolean
    If Player(MyIndex).PlayerTravel(mapNum).Unlocked = YES Then GetPlayerMapUnlocked = True
End Function

Private Sub SetPlayerMapUnlocked(ByVal mapNum As Long, ByVal Unlocked As Byte)
    Player(MyIndex).PlayerTravel(mapNum).Unlocked = Unlocked
End Sub

Private Sub SetPlayerMapTravelName(ByVal mapNum As Long, ByVal mapName As String)
    Player(MyIndex).PlayerTravel(mapNum).mapName = mapName
End Sub

Private Sub SetPlayerMapCostValue(ByVal mapNum As Long, ByVal costValue As Long)
    Player(MyIndex).PlayerTravel(mapNum).costValue = costValue
End Sub

Public Function GetPlayerMapCostValue(ByVal mapNum As Long) As Long
    GetPlayerMapCostValue = Player(MyIndex).PlayerTravel(mapNum).costValue
End Function

Public Sub InitPlayerMapTravel()
    Dim i As Long, FileName As String

    For i = 1 To MAX_MAP
        FileName = App.path & Texture_Path & Trim$(GameSetting.ThemePath) & "\ui\map-travel\" & i & ".ini"
        If FileExist(FileName) Then
            Player(MyIndex).PlayerTravel(i).SrcPosX = CLng(Trim$(GetVar(FileName, CStr(i), "SrcPosX")))
            Player(MyIndex).PlayerTravel(i).SrcPosY = CLng(Trim$(GetVar(FileName, CStr(i), "SrcPosY")))
            Player(MyIndex).PlayerTravel(i).SrcWidth = CLng(Trim$(GetVar(FileName, CStr(i), "SrcWidth")))
            Player(MyIndex).PlayerTravel(i).SrcHeight = CLng(Trim$(GetVar(FileName, CStr(i), "SrcHeight")))
            Player(MyIndex).PlayerTravel(i).IconPosX = CLng(Trim$(GetVar(FileName, CStr(i), "IconPosX")))
            Player(MyIndex).PlayerTravel(i).IconPosY = CLng(Trim$(GetVar(FileName, CStr(i), "IconPosY")))
            
            Player(MyIndex).PlayerTravel(i).DataExist = True
    
            DoEvents
        End If
    Next i
End Sub

Public Sub HandlePlayerTravel(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long, mapNum As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    For i = 1 To MAX_MAP
        mapNum = buffer.ReadLong

        If mapNum > 0 And mapNum <= MAX_MAP Then
            Call SetPlayerMapUnlocked(mapNum, buffer.ReadByte)
            Call SetPlayerMapTravelName(mapNum, buffer.ReadString)
            Call SetPlayerMapCostValue(mapNum, buffer.ReadLong)
        End If
    Next i

    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendPlayerTravel(ByVal TravelSlot As Long)
    Dim buffer As clsBuffer
    
    If GetPlayerMapUnlocked(TravelSlot) = False Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerTravel
    
    buffer.WriteLong TravelSlot
    
    SendData buffer.ToArray()
    
    buffer.Flush: Set buffer = Nothing
End Sub
