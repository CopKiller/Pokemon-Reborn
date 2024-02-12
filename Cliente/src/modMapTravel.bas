Attribute VB_Name = "modMapTravel"
Option Explicit

Public Type PlayerTravelRec
    Unlocked As Byte
    mapName As String
    
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
    CostValue As Long
    X As Long
    Y As Long
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

Private Function GetMapTravel(ByVal mapNum As Long) As Boolean
    If Map(mapNum).MapTravel.IsTravel = YES Then GetMapTravel = True
End Function

Public Sub InitPlayerMapTravel()
    Dim i As Long, filename As String

    For i = 1 To MAX_MAP
        filename = App.path & Texture_Path & Trim$(GameSetting.ThemePath) & "\ui\map-travel\" & i & ".ini"
        If FileExist(filename) Then
            Player(MyIndex).PlayerTravel(i).SrcPosX = CLng(Trim$(GetVar(filename, CStr(i), "SrcPosX")))
            Player(MyIndex).PlayerTravel(i).SrcPosY = CLng(Trim$(GetVar(filename, CStr(i), "SrcPosY")))
            Player(MyIndex).PlayerTravel(i).SrcWidth = CLng(Trim$(GetVar(filename, CStr(i), "SrcWidth")))
            Player(MyIndex).PlayerTravel(i).SrcHeight = CLng(Trim$(GetVar(filename, CStr(i), "SrcHeight")))
            Player(MyIndex).PlayerTravel(i).IconPosX = CLng(Trim$(GetVar(filename, CStr(i), "IconPosX")))
            Player(MyIndex).PlayerTravel(i).IconPosY = CLng(Trim$(GetVar(filename, CStr(i), "IconPosY")))
            Player(MyIndex).PlayerTravel(i).DataExist = True
    
            DoEvents
        End If
    Next i
End Sub

Public Sub HandlePlayerTravel(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long, mapNum As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    For i = 1 To MAX_MAP
        mapNum = buffer.ReadLong

        If mapNum > 0 And mapNum <= MAX_MAP Then
            Call SetPlayerMapUnlocked(mapNum, buffer.ReadByte)
            Call SetPlayerMapTravelName(mapNum, buffer.ReadString)
        End If
    Next i

    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendPlayerTravel(ByVal TravelSlot As Long)
    Dim buffer As clsBuffer
    
    If GetPlayerMapUnlocked(TravelSlot) Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerTravel
    
    buffer.WriteLong TravelSlot
    
    SendData buffer.ToArray()
    
    buffer.Flush: Set buffer = Nothing
End Sub
