Attribute VB_Name = "modEditor"
Option Explicit

'//Editor Constants
Public Const EDITOR_NONE As Byte = 0
Public Const EDITOR_MAP As Byte = 1
Public Const EDITOR_NPC As Byte = 2
Public Const EDITOR_POKEMON As Byte = 4
Public Const EDITOR_ITEM As Byte = 5
Public Const EDITOR_POKEMONMOVE As Byte = 6
Public Const EDITOR_ANIMATION As Byte = 7
Public Const EDITOR_SPAWN As Byte = 8
Public Const EDITOR_CONVERSATION As Byte = 9
Public Const EDITOR_SHOP As Byte = 10
Public Const EDITOR_QUEST As Byte = 11

'//Change
Public NpcChange(1 To MAX_NPC) As Boolean
Public PokemonChange(1 To MAX_POKEMON) As Boolean
Public ItemChange(1 To MAX_ITEM) As Boolean
Public PokemonMoveChange(1 To MAX_POKEMON_MOVE) As Boolean
Public AnimationChange(1 To MAX_ANIMATION) As Boolean
Public SpawnChange(1 To MAX_GAME_POKEMON) As Boolean
Public ConversationChange(1 To MAX_CONVERSATION) As Boolean
Public ShopChange(1 To MAX_SHOP) As Boolean
Public QuestChange(1 To MAX_QUEST) As Boolean

'//Editor data
Public Editor As Byte
Public CurTileset As Long
Public CurLayer As Byte
Public CurAttribute As Byte
Public IsAnimated As Byte
Public editorMapAnim As Long

Public EditorTileX As Long
Public EditorTileY As Long
Public EditorTileWidth As Byte
Public EditorTileHeight As Byte

Public EditorScrollY As Long
Public EditorScrollX As Long

Public EditorData1 As Long
Public EditorData2 As Long
Public EditorData3 As Long
Public EditorData4 As Long

Public EditorTmpNpc(1 To MAX_MAP_NPC) As Long
Public EditorTmpPokemon(1 To MAX_MAP_NPC) As Long

Public TileExpand As Boolean

Public EditorIndex As Long
Public EditorChange As Boolean
Public EditorStart As Boolean

' ****************
' ** Map Editor **
' ****************
Public Sub InitEditor_Map()
    '//Make sure they are not editing something on the other editors
    If Editor <> EDITOR_NONE Then Exit Sub
    If Count_Tileset <= 0 Then Exit Sub
    
    Editor = EDITOR_MAP
    
    With frmEditor_Map
        '//Set Max
        .scrlTileset.max = Count_Tileset
    
        '//Reset
        CurTileset = 1
        .scrlTileset.value = CurTileset
        
        CurLayer = MapLayer.Ground
        .optLayer(CurLayer).value = True
        
        CurAttribute = MapAttribute.Blocked
        .optAttribute(CurAttribute).value = True

        .optType(1).value = YES
        
        .fraLayers.Visible = True
        .fraAttributes.Visible = False
        
        editorMapAnim = 0
        
        TileExpand = False
        
        IsAnimated = NO
        .chkAnimated.value = IsAnimated
        
        ClearMapAttribute

        '//Open Window
        .Show
    End With
End Sub

Public Sub LoadTileset(ByVal tilesetNum As Long)
Dim X As Long, Y As Long
Dim Width As Long, Height As Long
    
    '//exit if there's no data
    If tilesetNum <= 0 Then Exit Sub
    
    With frmEditor_Map
        '//Reset data
        CurTileset = 0
        EditorTileX = 0
        EditorTileY = 0
        EditorTileWidth = 1
        EditorTileHeight = 1
        
        '//set tileset
        CurTileset = tilesetNum
        
        '//Set Scroll
        Width = GetPicWidth(Tex_Tileset(CurTileset)) * 2
        Height = GetPicHeight(Tex_Tileset(CurTileset)) * 2
        X = (Width \ TILE_X) - (.picTileset.scaleWidth \ TILE_X)
        If X >= 0 Then .scrlTileX.max = X
        Y = (Height \ TILE_Y) - (.picTileset.scaleHeight \ TILE_Y)
        If Y >= 0 Then .scrlTileY.max = Y
        
        .scrlTileX.value = 0
        .scrlTileY.value = 0
        EditorScrollX = 0
        EditorScrollY = 0
        '//horizontal scrolling
        If Width < .picTileset.scaleWidth Then
            .scrlTileX.Enabled = False
        Else
            .scrlTileX.Enabled = True
        End If
        '//vertical scrolling
        If Height < .picTileset.scaleHeight Then
            .scrlTileY.Enabled = False
        Else
            .scrlTileY.Enabled = True
        End If
    End With
End Sub

Public Sub MapEditorChooseTile(Button As Integer, X As Single, Y As Single, Optional ByVal Multiple As Boolean = False)
Dim Width As Long, Height As Long

    With frmEditor_Map
        '//Update X and Y value based on Scroll
        X = X + (EditorScrollX * TILE_X)
        Y = Y + (EditorScrollY * TILE_Y)
        
        If Button = vbLeftButton Then
            If Not Multiple Then
                '//reset
                EditorTileWidth = 1
                EditorTileHeight = 1
                
                '//set data
                EditorTileX = X \ TILE_X
                EditorTileY = Y \ TILE_Y
            Else
                '//convert the pixel number to tile number
                X = (X \ TILE_X) + 1
                Y = (Y \ TILE_Y) + 1
                
                '//check it's not out of bounds
                Width = GetPicWidth(Tex_Tileset(CurTileset)) * 2
                Height = GetPicHeight(Tex_Tileset(CurTileset)) * 2
                If X < 0 Then X = 0
                If X > Width / TILE_X Then X = Width / TILE_X
                If Y < 0 Then Y = 0
                If Y > Height / TILE_Y Then Y = Height / TILE_Y
                
                '//find out what to set the width + height of map editor to
                If X > EditorTileX Then ' drag right
                    EditorTileWidth = X - EditorTileX
                End If
                If Y > EditorTileY Then ' drag down
                    EditorTileHeight = Y - EditorTileY
                End If
            End If
        End If
    End With
End Sub

Public Sub MapEditorSetTile(ByVal X As Long, ByVal Y As Long)
Dim x2 As Long, Y2 As Long
    
    If EditorTileWidth = 1 And EditorTileHeight = 1 Then '//single
        With Map.Tile(X, Y)
            '//set layer
            .Layer(CurLayer, IsAnimated).Tile = CurTileset
            .Layer(CurLayer, IsAnimated).TileX = EditorTileX
            .Layer(CurLayer, IsAnimated).TileY = EditorTileY
            .Layer(CurLayer, IsAnimated).MapAnim = editorMapAnim
        End With
    Else '//multitile
        Y2 = 0 '//starting tile for y axis
        For Y = curTileY To curTileY + EditorTileHeight - 1
            x2 = 0 '//re-set x count every y loop
            For X = curTileX To curTileX + EditorTileWidth - 1
                If X >= 0 And X <= Map.MaxX Then
                    If Y >= 0 And Y <= Map.MaxY Then
                        With Map.Tile(X, Y)
                            .Layer(CurLayer, IsAnimated).Tile = CurTileset
                            .Layer(CurLayer, IsAnimated).TileX = EditorTileX + x2
                            .Layer(CurLayer, IsAnimated).TileY = EditorTileY + Y2
                            .Layer(CurLayer, IsAnimated).MapAnim = editorMapAnim
                        End With
                    End If
                End If
                x2 = x2 + 1
            Next
            Y2 = Y2 + 1
        Next
    End If
End Sub

Public Sub MapEditorMouseDown(ByVal Button As Integer)
Dim TempFill As LayerRec

    '//check if it's in boundary
    If Not isInBounds Then Exit Sub
    
    If Button = vbLeftButton Then
        If frmEditor_Map.optType(1).value = True Then   '//Layers
            MapEditorSetTile curTileX, curTileY
        ElseIf frmEditor_Map.optType(2).value = True Then '//Attributes
            With Map.Tile(curTileX, curTileY)
                .Attribute = CurAttribute
                .Data1 = EditorData1
                .Data2 = EditorData2
                .Data3 = EditorData3
                .Data4 = EditorData4
            End With
        End If
    ElseIf Button = vbRightButton Then
        If frmEditor_Map.optType(1).value = True Then   '//Layers
            With Map.Tile(curTileX, curTileY)
                '//clear layer
                .Layer(CurLayer, IsAnimated).Tile = 0
                .Layer(CurLayer, IsAnimated).TileX = 0
                .Layer(CurLayer, IsAnimated).TileY = 0
                .Layer(CurLayer, IsAnimated).MapAnim = 0
            End With
        ElseIf frmEditor_Map.optType(2).value = True Then '//Attributes
            With Map.Tile(curTileX, curTileY)
                '//clear attribute data
                .Attribute = 0
                .Data1 = 0
                .Data2 = 0
                .Data3 = 0
                .Data4 = 0
            End With
        End If
    End If
End Sub

Public Sub MapEditorFillLayer()
Dim X As Long, Y As Long

    If MsgBox("Are you sure that you want to fill all tiles in this layer?", vbYesNo) = vbYes Then
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                With Map.Tile(X, Y)
                    '//set layer
                    .Layer(CurLayer, IsAnimated).Tile = CurTileset
                    .Layer(CurLayer, IsAnimated).TileX = EditorTileX
                    .Layer(CurLayer, IsAnimated).TileY = EditorTileY
                    .Layer(CurLayer, IsAnimated).MapAnim = editorMapAnim
                End With
            Next
        Next
    End If
End Sub

Public Sub MapEditorClearLayer()
Dim X As Long, Y As Long

    If MsgBox("Are you sure that you want to clear all tiles in this layer?", vbYesNo) = vbYes Then
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                With Map.Tile(X, Y)
                    '//set layer
                    .Layer(CurLayer, IsAnimated).Tile = 0
                    .Layer(CurLayer, IsAnimated).TileX = 0
                    .Layer(CurLayer, IsAnimated).TileY = 0
                    .Layer(CurLayer, IsAnimated).MapAnim = 0
                End With
            Next
        Next
    End If
End Sub

Public Sub MapEditorFillAttribute()
Dim X As Long, Y As Long

    If MsgBox("Are you sure that you want to fill all attribute in this layer?", vbYesNo) = vbYes Then
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                With Map.Tile(X, Y)
                    '//set attribute
                    .Attribute = CurAttribute
                    .Data1 = EditorData1
                    .Data2 = EditorData2
                    .Data3 = EditorData3
                    .Data4 = EditorData4
                End With
            Next
        Next
    End If
End Sub

Public Sub MapEditorClearAttribute()
Dim X As Long, Y As Long

    If MsgBox("Are you sure that you want to clear all attribute in this layer?", vbYesNo) = vbYes Then
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                With Map.Tile(X, Y)
                    '//set attribute
                    .Attribute = 0
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End With
            Next
        Next
    End If
End Sub

Public Sub RandomPlaceLayer(ByVal Seed As Long)
Dim X As Long, Y As Long
Dim i As Long

    For i = 1 To Seed
        X = Rand(0, Map.MaxX)
        Y = Rand(0, Map.MaxY)
        
        If IsValidMapPoint(X, Y) Then
            With Map.Tile(X, Y)
                '//set layer
                .Layer(CurLayer, IsAnimated).Tile = CurTileset
                .Layer(CurLayer, IsAnimated).TileX = EditorTileX
                .Layer(CurLayer, IsAnimated).TileY = EditorTileY
                .Layer(CurLayer, IsAnimated).MapAnim = editorMapAnim
            End With
        End If
    Next
End Sub

Public Sub ClearMapAttribute()
Dim i As Long

    With frmEditor_Map
        .fraAttribute.Visible = False
        
        .fraNpcSpawn.Visible = False
        .cmbNpcSpawn.Clear
        For i = 1 To MAX_MAP_NPC
            If Map.Npc(i) > 0 Then
                .cmbNpcSpawn.AddItem i & ": " & Trim$(Npc(Map.Npc(i)).Name)
            Else
                .cmbNpcSpawn.AddItem i & ": None"
            End If
        Next
        
        .fraWarp.Visible = False
        .scrlWarpMap.value = 0
        .scrlWarpX.value = 0
        .scrlWarpY.value = 0
        .scrlWarpDir.value = 0
        
        .fraConvoTile.Visible = False
        .scrlConvoTileNum.value = 0

        EditorData1 = 0
        EditorData2 = 0
        EditorData3 = 0
        EditorData4 = 0
    End With
End Sub

Public Sub CloseMapEditor(ByVal RequestMap As Boolean)
Dim buffer As clsBuffer

    Editor = EDITOR_NONE
    Unload frmEditor_Map

    If RequestMap Then
        Set buffer = New clsBuffer
        buffer.WriteLong CNeedMap
        buffer.WriteByte YES
        SendData buffer.ToArray()
        Set buffer = Nothing
        
        GettingMap = True
    End If
End Sub

Public Sub MapEditorSend()
    If Player(MyIndex).Access < ACCESS_MAPPER Then Exit Sub
    CloseMapEditor False
    '//Update Revision
    Map.Revision = Map.Revision + 1
    SendMap
    GettingMap = True
End Sub

Private Function CheckSameArea(ByRef ArrayData() As TilePosRec, ByVal X As Long, ByVal Y As Long)
Dim i As Long

    For i = LBound(ArrayData) To UBound(ArrayData)
        If ArrayData(i).Used Then
            If ArrayData(i).X = X And ArrayData(i).Y = Y Then
                CheckSameArea = True
            End If
        End If
    Next
End Function

Private Sub CheckTileMatch(ByRef ArrayData() As TilePosRec, ByVal X As Long, ByVal Y As Long, ByRef CheckTile As LayerRec, ByRef TargetTile As LayerRec, ByRef Size As Long, ByRef Count As Long)
Dim TileMatched As Boolean

    ' Check matching
    If CheckTile.Tile = TargetTile.Tile And CheckTile.TileX = TargetTile.TileX And CheckTile.TileY = TargetTile.TileY And CheckTile.MapAnim = TargetTile.MapAnim Then
        TileMatched = True
    End If
    
    ' Check if we already check this part
    If Not CheckSameArea(ArrayData, X, Y) Then
        If TileMatched Then
            Count = Count + 1
            If Count >= Size Then
                Size = Size * 2
                ReDim Preserve ArrayData(Size) As TilePosRec
            End If
            ArrayData(Count).Used = True
            ArrayData(Count).Y = Y
            ArrayData(Count).X = X
        End If
    End If
End Sub

Public Sub Fill_Tile_Layer(ByVal Layer As MapLayer, ByVal LayerAnim As Byte, ByVal TileX As Integer, ByVal TileY As Integer, ByRef ReplaceTile As LayerRec)
Dim ArrayData() As TilePosRec
Dim CheckLayer As LayerRec, ConnectLayer As LayerRec
Dim Size As Long, Count As Long, LoopCount As Long
Dim CurSize As Long, Resized As Long

    ' Redim Array
    Count = 0
    Size = 1
    CurSize = Size
    Resized = 0
    ReDim ArrayData(0 To Size) As TilePosRec
    
    ' Check first tile
    CheckLayer = Map.Tile(TileX, TileY).Layer(Layer, LayerAnim)
    ' Fill the tile
    Map.Tile(TileX, TileY).Layer(Layer, LayerAnim) = ReplaceTile
    
    ' ///////////////////////
    ' //// Check Connect ////
    ' ///////////////////////
    ' Check north
    If (TileY - 1) >= 0 Then
        ' Check tile num
        ConnectLayer = Map.Tile(TileX, (TileY - 1)).Layer(Layer, LayerAnim)
        CheckTileMatch ArrayData, TileX, TileY - 1, CheckLayer, ConnectLayer, Size, Count
        If CurSize <> Size Then
            Resized = Resized + 1
            CurSize = Size
        End If
    End If
    ' Check south
    If (TileY + 1) < Map.MaxY Then
        ' Check tile num
        ConnectLayer = Map.Tile(TileX, (TileY + 1)).Layer(Layer, LayerAnim)
        CheckTileMatch ArrayData, TileX, TileY + 1, CheckLayer, ConnectLayer, Size, Count
        If CurSize <> Size Then
            Resized = Resized + 1
            CurSize = Size
        End If
    End If
    ' Check west
    If (TileX - 1) >= 0 Then
        ' Check tile num
        ConnectLayer = Map.Tile((TileX - 1), TileY).Layer(Layer, LayerAnim)
        CheckTileMatch ArrayData, TileX - 1, TileY, CheckLayer, ConnectLayer, Size, Count
        If CurSize <> Size Then
            Resized = Resized + 1
            CurSize = Size
        End If
    End If
    ' Check east
    If (TileX + 1) < Map.MaxX Then
        ' Check tile num
        ConnectLayer = Map.Tile((TileX + 1), TileY).Layer(Layer, LayerAnim)
        CheckTileMatch ArrayData, TileX + 1, TileY, CheckLayer, ConnectLayer, Size, Count
        If CurSize <> Size Then
            Resized = Resized + 1
            CurSize = Size
        End If
    End If
    
    ' //////////////////////////////////////////////
    ' //// Start the loop on all connected tile ////
    ' //////////////////////////////////////////////
    LoopCount = 0
    Do While (LoopCount <= Count)
        ' Check if array in used
        If ArrayData(LoopCount).Used Then
                ' Fill the tile
                Map.Tile(ArrayData(LoopCount).X, ArrayData(LoopCount).Y).Layer(Layer, LayerAnim) = ReplaceTile
            
                ' ///////////////////////
                ' //// Check Connect ////
                ' ///////////////////////
                ' Check north
                If (ArrayData(LoopCount).Y - 1) >= 0 Then
                    ' Check tile num
                    ConnectLayer = Map.Tile(ArrayData(LoopCount).X, (ArrayData(LoopCount).Y - 1)).Layer(Layer, LayerAnim)
                    CheckTileMatch ArrayData, ArrayData(LoopCount).X, ArrayData(LoopCount).Y - 1, CheckLayer, ConnectLayer, Size, Count
                    If CurSize <> Size Then
                        Resized = Resized + 1
                        CurSize = Size
                    End If
                End If
                ' Check south
                If (ArrayData(LoopCount).Y + 1) <= Map.MaxY Then
                    ' Check tile num
                    ConnectLayer = Map.Tile(ArrayData(LoopCount).X, (ArrayData(LoopCount).Y + 1)).Layer(Layer, LayerAnim)
                    CheckTileMatch ArrayData, ArrayData(LoopCount).X, ArrayData(LoopCount).Y + 1, CheckLayer, ConnectLayer, Size, Count
                    If CurSize <> Size Then
                        Resized = Resized + 1
                        CurSize = Size
                    End If
                End If
                ' Check west
                If (ArrayData(LoopCount).X - 1) >= 0 Then
                    ' Check tile num
                    ConnectLayer = Map.Tile((ArrayData(LoopCount).X - 1), ArrayData(LoopCount).Y).Layer(Layer, LayerAnim)
                    CheckTileMatch ArrayData, ArrayData(LoopCount).X - 1, ArrayData(LoopCount).Y, CheckLayer, ConnectLayer, Size, Count
                    If CurSize <> Size Then
                        Resized = Resized + 1
                        CurSize = Size
                    End If
                End If
                ' Check east
                If (ArrayData(LoopCount).X + 1) <= Map.MaxX Then
                    ' Check tile num
                    ConnectLayer = Map.Tile((ArrayData(LoopCount).X + 1), ArrayData(LoopCount).Y).Layer(Layer, LayerAnim)
                    CheckTileMatch ArrayData, ArrayData(LoopCount).X + 1, ArrayData(LoopCount).Y, CheckLayer, ConnectLayer, Size, Count
                    If CurSize <> Size Then
                        Resized = Resized + 1
                        CurSize = Size
                    End If
                End If
        End If
        LoopCount = LoopCount + 1
    Loop
    
    Debug.Print "Count of Resize: " & Resized
End Sub

' ****************
' ** Npc Editor **
' ****************
Public Sub InitEditor_Npc()
Dim i As Long

    '//Make sure they are not editing something on the other editors
    If Editor <> EDITOR_NONE Then Exit Sub
    Editor = EDITOR_NPC
    
    With frmEditor_Npc
        .cmbPokeNum.Clear
        .cmbPokeNum.AddItem "None"
        For i = 1 To MAX_POKEMON
            .cmbPokeNum.AddItem i & ": " & Trim$(Pokemon(i).Name)
        Next
        
        .cmbMoveset.Clear
        .cmbMoveset.AddItem "None"
        For i = 1 To MAX_POKEMON_MOVE
            .cmbMoveset.AddItem i & ": " & Trim$(PokemonMove(i).Name)
        Next
        
        '//Clear Index
        .lstIndex.Clear
        '//Add Item
        For i = 1 To MAX_NPC
            .lstIndex.AddItem i & ": " & Trim$(Npc(i).Name)
        Next
        .lstIndex.ListIndex = 0
        NpcEditorLoadIndex .lstIndex.ListIndex + 1
        
        '//No edit done
        EditorChange = False
        
        .Show
    End With
End Sub

Public Sub NpcEditorLoadIndex(ByVal xIndex As Long)
Dim X As Byte

    EditorIndex = xIndex
    
    With frmEditor_Npc
        '//General
        .txtName.Text = Trim$(Npc(xIndex).Name)
        .scrlSprite.value = Npc(xIndex).Sprite
        .cmbBehaviour.ListIndex = Npc(xIndex).Behaviour
        .scrlConvo.value = Npc(xIndex).Convo
        
        .lstPokemon.Clear
        For X = 1 To MAX_PLAYER_POKEMON
            If Npc(xIndex).PokemonNum(X) > 0 Then
                .lstPokemon.AddItem X & ": " & Trim$(Pokemon(Npc(xIndex).PokemonNum(X)).Name) & " Lv: " & Npc(xIndex).PokemonLevel(X)
            Else
                .lstPokemon.AddItem X & ": None"
            End If
        Next
        .lstPokemon.ListIndex = 0
        
        .lstMoveset.Clear
        For X = 1 To MAX_MOVESET
            If Npc(xIndex).PokemonMoveset(1, X) > 0 Then
                .lstMoveset.AddItem X & ": " & Trim$(PokemonMove(Npc(xIndex).PokemonMoveset(1, X)).Name)
            Else
                .lstMoveset.AddItem X & ": None"
            End If
        Next
        .lstMoveset.ListIndex = 0
        
        .cmbPokeNum.ListIndex = Npc(xIndex).PokemonNum(1)
        .txtLevel.Text = Npc(xIndex).PokemonLevel(1)
        .cmbMoveset.ListIndex = Npc(xIndex).PokemonMoveset(1, 1)
        
        .txtReward.Text = Npc(xIndex).Reward
        .txtRewardExp.Text = Npc(xIndex).RewardExp
        .scrlWinConvo.value = Npc(xIndex).WinEvent
    End With
    
    NpcChange(xIndex) = True
End Sub

Public Sub CloseNpcEditor()
Dim i As Long

    For i = 1 To MAX_NPC
        NpcChange(i) = False
    Next
    Editor = EDITOR_NONE
    Unload frmEditor_Npc
End Sub

' ********************
' ** Pokemon Editor **
' ********************
Public Sub InitEditor_Pokemon()
Dim i As Long

    '//Make sure they are not editing something on the other editors
    If Editor <> EDITOR_NONE Then Exit Sub
    Editor = EDITOR_POKEMON
    
    With frmEditor_Pokemon
        .cmbMoveNum.Clear
        .cmbEggMoveNum.Clear
        .cmbMoveNum.AddItem "None"
        .cmbEggMoveNum.AddItem "None"
        .cmbItemMove.AddItem "None"
        For i = 1 To MAX_POKEMON_MOVE
            .cmbMoveNum.AddItem i & ": " & Trim$(PokemonMove(i).Name)
            .cmbEggMoveNum.AddItem i & ": " & Trim$(PokemonMove(i).Name)
            .cmbItemMove.AddItem i & ": " & Trim$(PokemonMove(i).Name)
        Next
        
        .cmbItemNum.Clear
        .cmbItemNum.AddItem "None"
        For i = 1 To MAX_ITEM
            .cmbItemNum.AddItem i & ": " & Trim$(Item(i).Name)
        Next
        
        '//Clear Index
        .lstIndex.Clear
        '//Add Item
        For i = 1 To MAX_POKEMON
            .lstIndex.AddItem i & ": " & Trim$(Pokemon(i).Name)
        Next
        .lstIndex.ListIndex = 0
        PokemonEditorLoadIndex .lstIndex.ListIndex + 1
        
        '//No edit done
        EditorChange = False
        
        .Show
    End With
End Sub

Public Sub PokemonEditorLoadIndex(ByVal xIndex As Long)
Dim X As Byte

    EditorIndex = xIndex
    
    With frmEditor_Pokemon
        '//General
        .txtName.Text = Trim$(Pokemon(xIndex).Name)
        .scrlSprite.value = Pokemon(xIndex).Sprite
        .cmbBehaviour.ListIndex = Pokemon(xIndex).Behaviour
        .chkScale.value = Pokemon(xIndex).ScaleSprite
        
        '//Stats
        For X = 1 To StatEnum.Stat_Count - 1
            .txtBaseStat(X).Text = Pokemon(xIndex).BaseStat(X)
        Next
        
        '//Type
        .cmbPrimaryType.ListIndex = Pokemon(xIndex).PrimaryType
        .cmbSecondaryType.ListIndex = Pokemon(xIndex).SecondaryType
        
        '//Other
        .txtCatchRate.Text = Pokemon(xIndex).CatchRate
        .txtFemaleRate.Text = Pokemon(xIndex).FemaleRate
        .txtEggCycle.Text = Pokemon(xIndex).EggCycle
        .cmbEggGroup.ListIndex = Pokemon(xIndex).EggGroup
        .cmbEVYeildType.ListIndex = Pokemon(xIndex).EvYeildType
        .txtEVYeildVal.Text = Pokemon(xIndex).EvYeildVal
        .txtBaseExp.Text = Pokemon(xIndex).BaseExp
        .cmbGrowthRate.ListIndex = Pokemon(xIndex).GrowthRate
        .txtHeight.Text = Pokemon(xIndex).Height
        .txtWeight.Text = Pokemon(xIndex).Weight
        .txtSpecies.Text = Trim$(Pokemon(xIndex).Species)
        .txtPokedexEntry.Text = Trim$(Pokemon(xIndex).PokeDexEntry)
        
        '//Evolution
        .scrlEvolveIndex.value = 1
        
        .scrlEvolve.value = Pokemon(xIndex).evolveNum(1)
        .txtEvolveLevel.Text = Pokemon(xIndex).EvolveLevel(1)
        .scrlEvolveCondition.value = Pokemon(xIndex).EvolveCondition(1)
        .txtEvolveConditionData.Text = Pokemon(xIndex).EvolveConditionData(1)
        
        '//Moveset
        .lstMoveset.Clear
        .lstEggMove.Clear
        For X = 1 To MAX_POKEMON_MOVESET
            If Pokemon(xIndex).Moveset(X).MoveNum > 0 Then
                .lstMoveset.AddItem X & ": " & Trim$(PokemonMove(Pokemon(xIndex).Moveset(X).MoveNum).Name) & " Lv:" & Pokemon(xIndex).Moveset(X).MoveLevel
            Else
                .lstMoveset.AddItem X & ": None"
            End If
            If Pokemon(xIndex).EggMoveset(X) > 0 Then
                .lstEggMove.AddItem X & ": " & Trim$(PokemonMove(Pokemon(xIndex).EggMoveset(X)).Name)
            Else
                .lstEggMove.AddItem X & ": None"
            End If
        Next
        .lstMoveset.ListIndex = 0
        .lstEggMove.ListIndex = 0
        
        .lstItemMoveset.Clear
        For X = 1 To 110
            If Pokemon(xIndex).ItemMoveset(X) > 0 Then
                .lstItemMoveset.AddItem X & ": " & Trim$(PokemonMove(Pokemon(xIndex).ItemMoveset(X)).Name)
            Else
                .lstItemMoveset.AddItem X & ": None"
            End If
        Next
        .lstItemMoveset.ListIndex = 0
        
        .cmbMoveNum.ListIndex = Pokemon(xIndex).Moveset(1).MoveNum
        .txtMoveLevel.Text = Pokemon(xIndex).Moveset(1).MoveLevel
        
        .cmbEggMoveNum.ListIndex = Pokemon(xIndex).EggMoveset(1)
        .cmbItemMove.ListIndex = Pokemon(xIndex).ItemMoveset(1)
        
        .scrlRange.value = Pokemon(xIndex).Range
        '.scrlItemMove.value = Pokemon(xIndex).ItemMoveset(1)
        
        '//Drop
        .lstItemDrop.Clear
        For X = 1 To MAX_DROP
            If Pokemon(xIndex).DropNum(X) > 0 Then
                .lstItemDrop.AddItem X & ": " & Trim$(Item(Pokemon(xIndex).DropNum(X)).Name) & " Rate: " & Pokemon(xIndex).DropRate(X)
            Else
                .lstItemDrop.AddItem X & ": None"
            End If
        Next
        .lstItemDrop.ListIndex = 0
        
        .cmbItemNum.ListIndex = Pokemon(xIndex).DropNum(1)
        .txtItemDropRate.Text = Pokemon(xIndex).DropRate(1)
    End With
    
    PokemonChange(xIndex) = True
End Sub

Public Sub ClosePokemonEditor()
Dim i As Long

    For i = 1 To MAX_POKEMON
        PokemonChange(i) = False
    Next
    Editor = EDITOR_NONE
    Unload frmEditor_Pokemon
End Sub

' *****************
' ** Item Editor **
' *****************
Public Sub InitEditor_Item()
Dim i As Long

    '//Make sure they are not editing something on the other editors
    If Editor <> EDITOR_NONE Then Exit Sub
    Editor = EDITOR_ITEM
    
    With frmEditor_Item
        .cmbMoveList.Clear
        .cmbMoveList.AddItem "None"
        For i = 1 To MAX_POKEMON_MOVE
            .cmbMoveList.AddItem i & ": " & Trim$(PokemonMove(i).Name)
        Next
        
        '//Clear Index
        .lstIndex.Clear
        '//Add Item
        For i = 1 To MAX_ITEM
            .lstIndex.AddItem i & ": " & Trim$(Item(i).Name)
        Next
        .lstIndex.ListIndex = 0
        ItemEditorLoadIndex .lstIndex.ListIndex + 1
        
        '//No edit done
        EditorChange = False
        
        .Show
    End With
End Sub

Public Sub ItemEditorLoadIndex(ByVal xIndex As Long)
    EditorIndex = xIndex
    
    With frmEditor_Item
        '//General
        .txtName.Text = Trim$(Item(xIndex).Name)
        .scrlSprite.value = Item(xIndex).Sprite
        .chkStock.value = Item(xIndex).Stock
        .cmbType.ListIndex = Item(xIndex).Type
        
        If Item(xIndex).Type = ItemTypeEnum.Pokeball Then
            .fraPokeball.Visible = True
            .txtCatchRate.Text = Item(xIndex).Data1
            .scrlBallSprite.value = Item(xIndex).Data2
            .chkAutoCatch.value = Item(xIndex).Data3
        Else
            .fraPokeball.Visible = False
        End If
        
        If Item(xIndex).Type = ItemTypeEnum.Medicine Then
            .fraMedicine.Visible = True
            .cmbMedicineType.ListIndex = Item(xIndex).Data1
            .txtValue.Text = Item(xIndex).Data2
            .chkLevelUp.value = Item(xIndex).Data3
        Else
            .fraMedicine.Visible = False
        End If
        
        If Item(xIndex).Type = ItemTypeEnum.keyItems Then
            .fraKeyItem.Visible = True
            .cmbKeyItemType.ListIndex = Item(xIndex).Data1
            .scrlSpriteType.value = Item(xIndex).Data2
        Else
            .fraKeyItem.Visible = False
        End If
        
        If Item(xIndex).Type = ItemTypeEnum.TM_HM Then
            .fraTMHM.Visible = True
            .cmbMoveList.ListIndex = Item(xIndex).Data1
            .chkTakeItem.value = Item(xIndex).Data2
        Else
            .fraTMHM.Visible = False
        End If
        
        .txtPrice.Text = Item(xIndex).Price
        
        .txtDesc.Text = Trim$(Item(xIndex).Desc)
    End With
    
    ItemChange(xIndex) = True
End Sub

Public Sub CloseItemEditor()
Dim i As Long

    For i = 1 To MAX_ITEM
        ItemChange(i) = False
    Next
    Editor = EDITOR_NONE
    Unload frmEditor_Item
End Sub

' ************************
' ** PokemonMove Editor **
' ************************
Public Sub InitEditor_PokemonMove()
Dim i As Long

    '//Make sure they are not editing something on the other editors
    If Editor <> EDITOR_NONE Then Exit Sub
    Editor = EDITOR_POKEMONMOVE
    
    EditorStart = True
    With frmEditor_Move
        '//Sound
        .cmbSound.Clear
        .cmbSound.AddItem "None."
        For i = 1 To UBound(soundCache)
            .cmbSound.AddItem Trim$(soundCache(i))
        Next
        
        '//Clear Index
        .lstIndex.Clear
        '//Add Item
        For i = 1 To MAX_POKEMON_MOVE
            .lstIndex.AddItem i & ": " & Trim$(PokemonMove(i).Name)
        Next
        .lstIndex.ListIndex = 0
        PokemonMoveEditorLoadIndex .lstIndex.ListIndex + 1
        
        '//No edit done
        EditorChange = False
        
        .Show
    End With
    EditorStart = False
End Sub

Public Sub PokemonMoveEditorLoadIndex(ByVal xIndex As Long)
Dim X As Byte
Dim i As Long

    EditorIndex = xIndex
    
    With frmEditor_Move
        '//General
        .txtName.Text = Trim$(PokemonMove(xIndex).Name)
        .cmbType.ListIndex = PokemonMove(xIndex).Type
        .cmbCategory.ListIndex = PokemonMove(xIndex).Category
        .txtPP.Text = PokemonMove(xIndex).PP
        .txtMaxPP.Text = PokemonMove(xIndex).MaxPP
        .txtPower.Text = PokemonMove(xIndex).Power
        .scrlRange.value = PokemonMove(xIndex).Range
        .txtDescription.Text = PokemonMove(xIndex).Description
        For X = 1 To StatEnum.Stat_Count - 1
            .txtBuffDebuff(X).Text = PokemonMove(xIndex).dStat(X)
        Next
        .optTargetType(PokemonMove(xIndex).targetType).value = True
        .txtInterval.Text = PokemonMove(xIndex).Interval
        .scrlAnimation.value = PokemonMove(xIndex).Animation
        .cmbAttackType.ListIndex = PokemonMove(xIndex).AttackType
        .txtDuration.Text = PokemonMove(xIndex).Duration
        .txtCooldown.Text = PokemonMove(xIndex).Cooldown
        .txtCastTime.Text = PokemonMove(xIndex).CastTime
        .txtAmountOfAttack.Text = PokemonMove(xIndex).AmountOfAttack
        .chkPlaySelf.value = PokemonMove(xIndex).SelfAnim
        
        '//find the music we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If Trim$(.cmbSound.List(i)) = Trim$(PokemonMove(xIndex).Sound) Then
                    .cmbSound.ListIndex = i
                    Exit For
                End If
            Next
            If .cmbSound.ListIndex <= 0 Then
                .cmbSound.ListIndex = 0
            End If
        End If
        
        '//Status
        .cmbStatus.ListIndex = PokemonMove(xIndex).pStatus
        .txtStatusChance.Text = PokemonMove(xIndex).pStatusChance
        
        .txtRecoilDamage.Text = PokemonMove(xIndex).RecoilDamage
        .txtAbsorbDamage.Text = PokemonMove(xIndex).AbsorbDamage
        
        '//Weather
        .cmbWeather.ListIndex = PokemonMove(xIndex).ChangeWeather
        .cmbBoostWeather.ListIndex = PokemonMove(xIndex).BoostWeather
        .cmbStatusReq.ListIndex = PokemonMove(xIndex).StatusReq
        .cmbDecreaseWeather.ListIndex = PokemonMove(xIndex).DecreaseWeather
        .chkStatusToSelf.value = PokemonMove(xIndex).StatusToSelf
        .cmbReflectType.ListIndex = PokemonMove(xIndex).ReflectType
        .chkProtect.value = PokemonMove(xIndex).CastProtect
        .cmbSelfStatusReq.ListIndex = PokemonMove(xIndex).SelfStatusReq
    End With
    
    PokemonMoveChange(xIndex) = True
End Sub

Public Sub ClosePokemonMoveEditor()
Dim i As Long

    For i = 1 To MAX_POKEMON_MOVE
        PokemonMoveChange(i) = False
    Next
    Editor = EDITOR_NONE
    Unload frmEditor_Move
End Sub

' *****************
' ** Animation Editor **
' *****************
Public Sub InitEditor_Animation()
Dim i As Long

    '//Make sure they are not editing something on the other editors
    If Editor <> EDITOR_NONE Then Exit Sub
    Editor = EDITOR_ANIMATION
    
    With frmEditor_Animation
        '//Clear Index
        .lstIndex.Clear
        '//Add Item
        For i = 1 To MAX_ANIMATION
            .lstIndex.AddItem i & ": " & Trim$(Animation(i).Name)
        Next
        .lstIndex.ListIndex = 0
        AnimationEditorLoadIndex .lstIndex.ListIndex + 1
        
        '//No edit done
        EditorChange = False
        
        .Show
    End With
End Sub

Public Sub AnimationEditorLoadIndex(ByVal xIndex As Long)
Dim i As Long

    EditorIndex = xIndex
    
    With frmEditor_Animation
        '//General
        .txtName.Text = Trim$(Animation(xIndex).Name)
        
        For i = 0 To 1
            .scrlSprite(i).value = Animation(xIndex).Sprite(i)
            .scrlFrameCount(i).value = Animation(xIndex).Frames(i)
            .scrlLoopCount(i).value = Animation(xIndex).LoopCount(i)
            
            If Animation(xIndex).looptime(i) > 0 Then
                .scrlLoopTime(i).value = Animation(xIndex).looptime(i)
            Else
                .scrlLoopTime(i).value = 45
            End If
        Next
    End With
    
    AnimationChange(xIndex) = True
End Sub

Public Sub CloseAnimationEditor()
Dim i As Long

    For i = 1 To MAX_ANIMATION
        AnimationChange(i) = False
    Next
    Editor = EDITOR_NONE
    Unload frmEditor_Animation
End Sub

' *****************
' ** Spawn Editor **
' *****************
Public Sub InitEditor_Spawn()
Dim i As Long

    '//Make sure they are not editing something on the other editors
    If Editor <> EDITOR_NONE Then Exit Sub
    Editor = EDITOR_SPAWN
    
    With frmEditor_Spawn
        .cmbPokemonNum.Clear
        .cmbPokemonNum.AddItem "None."
        For i = 1 To MAX_POKEMON
            .cmbPokemonNum.AddItem i & ": " & Trim$(Pokemon(i).Name)
        Next
        
        '//Clear Index
        .lstMapPokemon.Clear
        '//Add Item
        For i = 1 To MAX_GAME_POKEMON
            If Spawn(i).PokeNum > 0 Then
                .lstMapPokemon.AddItem i & ": " & Trim$(Pokemon(Spawn(i).PokeNum).Name)
            Else
                .lstMapPokemon.AddItem i & ": "
            End If
        Next
        .lstMapPokemon.ListIndex = 0
        SpawnEditorLoadIndex .lstMapPokemon.ListIndex + 1
        
        '//No edit done
        EditorChange = False
        
        .Show
    End With
End Sub

Public Sub SpawnEditorLoadIndex(ByVal xIndex As Long)
Dim i As Long

    EditorIndex = xIndex
    
    With frmEditor_Spawn
        '//General
        .cmbPokemonNum.ListIndex = Spawn(xIndex).PokeNum
        
        .txtMinLevel.Text = Spawn(xIndex).MinLevel
        .txtMaxLevel.Text = Spawn(xIndex).MaxLevel
        
        .txtRespawn.Text = Spawn(xIndex).Respawn
        
        .txtSpawnMin.Text = Spawn(xIndex).SpawnTimeMin
        .txtSpawnMax.Text = Spawn(xIndex).SpawnTimeMax
        
        .txtRarity.Text = Spawn(xIndex).Rarity
        
        '//location
        .chkRandomMap.value = Spawn(xIndex).randomMap
        .chkRandomXY.value = Spawn(xIndex).randomXY
        .txtMap.Text = Spawn(xIndex).MapNum
        .txtX.Text = Spawn(xIndex).MapX
        .txtY.Text = Spawn(xIndex).MapY
        .chkCanCatch.value = Spawn(xIndex).CanCatch
        .chkNoExp.value = Spawn(xIndex).NoExp
        .scrlPokeBuff.value = Spawn(xIndex).PokeBuff
    End With
    
    SpawnChange(xIndex) = True
End Sub

Public Sub CloseSpawnEditor()
Dim i As Long

    For i = 1 To MAX_GAME_POKEMON
        SpawnChange(i) = False
    Next
    Editor = EDITOR_NONE
    Unload frmEditor_Spawn
End Sub

' *****************
' ** Conversation Editor **
' *****************
Public Sub InitEditor_Conversation()
Dim i As Long

    '//Make sure they are not editing something on the other editors
    If Editor <> EDITOR_NONE Then Exit Sub
    Editor = EDITOR_CONVERSATION
    
    With frmEditor_Conversation
        '//Clear Index
        .lstIndex.Clear
        '//Add Item
        For i = 1 To MAX_CONVERSATION
            .lstIndex.AddItem i & ": " & Trim$(Conversation(i).Name)
        Next
        .lstIndex.ListIndex = 0
        ItemEditorLoadIndex .lstIndex.ListIndex + 1
        
        '//No edit done
        EditorChange = False
        
        .Show
    End With
End Sub

Public Sub ConversationEditorLoadIndex(ByVal xIndex As Long)
Dim i As Byte

    EditorIndex = xIndex
    
    With frmEditor_Conversation
        '//General
        .txtName.Text = Trim$(Conversation(xIndex).Name)
        
        .scrlData.value = 1
        .scrlLanguage.value = 1
        For i = 1 To 3
            .txtReply(i).Text = Trim$(Conversation(xIndex).ConvData(1).TextLang(1).tReply(i))
            .txtReplyMove(i).Text = (Conversation(xIndex).ConvData(1).tReplyMove(i))
        Next
        .txtText.Text = Trim$(Conversation(xIndex).ConvData(1).TextLang(1).Text)
        .scrlCustomScript.value = Conversation(xIndex).ConvData(1).CustomScript
        .chkNoText.value = Conversation(xIndex).ConvData(1).NoText
        .chkNoReply.value = Conversation(xIndex).ConvData(1).NoReply
        .txtMoveTo.Text = Conversation(xIndex).ConvData(1).MoveNext
        .txtCustomScriptData.Text = Conversation(xIndex).ConvData(1).CustomScriptData
        .txtCustomScriptData2.Text = Conversation(xIndex).ConvData(1).CustomScriptData2
        .txtCustomScriptData3.Text = Conversation(xIndex).ConvData(1).CustomScriptData3
    End With
    
    ConversationChange(xIndex) = True
End Sub

Public Sub CloseConversationEditor()
Dim i As Long

    For i = 1 To MAX_CONVERSATION
        ConversationChange(i) = False
    Next
    Editor = EDITOR_NONE
    Unload frmEditor_Conversation
End Sub

' *****************
' ** Shop Editor **
' *****************
Public Sub InitEditor_Shop()
Dim i As Long

    '//Make sure they are not editing something on the other editors
    If Editor <> EDITOR_NONE Then Exit Sub
    Editor = EDITOR_SHOP
    
    With frmEditorShop
        '//Clear Index
        .lstIndex.Clear
        '//Add Item
        For i = 1 To MAX_SHOP
            .lstIndex.AddItem i & ": " & Trim$(Shop(i).Name)
        Next
        .lstIndex.ListIndex = 0
        ItemEditorLoadIndex .lstIndex.ListIndex + 1
        
        '//No edit done
        EditorChange = False
        
        .Show
    End With
End Sub

Public Sub ShopEditorLoadIndex(ByVal xIndex As Long)
Dim i As Byte

    EditorIndex = xIndex
    
    With frmEditorShop
        '//General
        .txtName.Text = Trim$(Shop(xIndex).Name)
        
        '//List
        .lstShopItem.Clear
        For i = 1 To MAX_SHOP_ITEM
            If Shop(xIndex).ShopItem(i).Num > 0 Then
                .lstShopItem.AddItem i & ": " & Trim$(Item(Shop(xIndex).ShopItem(i).Num).Name) & " - Price: $" & Shop(xIndex).ShopItem(i).Price
            Else
                .lstShopItem.AddItem i & ": None - Price: $" & Shop(xIndex).ShopItem(i).Price
            End If
        Next
        .lstShopItem.ListIndex = 0
        
        .scrlItemNum.value = Shop(xIndex).ShopItem(1).Num
        .txtPrice.Text = Shop(xIndex).ShopItem(1).Price
    End With
    
    ShopChange(xIndex) = True
End Sub

Public Sub CloseShopEditor()
Dim i As Long

    For i = 1 To MAX_SHOP
        ShopChange(i) = False
    Next
    Editor = EDITOR_NONE
    Unload frmEditorShop
End Sub

' *****************
' ** Quest Editor **
' *****************
Public Sub InitEditor_Quest()
Dim i As Long

    '//Make sure they are not editing something on the other editors
    If Editor <> EDITOR_NONE Then Exit Sub
    Editor = EDITOR_QUEST
    
    With frmEditor_Quest
        '//Clear Index
        .lstIndex.Clear
        '//Add Item
        For i = 1 To MAX_QUEST
            .lstIndex.AddItem i & ": " & Trim$(Quest(i).Name)
        Next
        .lstIndex.ListIndex = 0
        ItemEditorLoadIndex .lstIndex.ListIndex + 1
        
        '//No edit done
        EditorChange = False
        
        .Show
    End With
End Sub

Public Sub QuestEditorLoadIndex(ByVal xIndex As Long)
Dim i As Byte

    EditorIndex = xIndex
    
    With frmEditor_Quest
        '//General
        .txtName.Text = Trim$(Quest(xIndex).Name)
    End With
    
    QuestChange(xIndex) = True
End Sub

Public Sub CloseQuestEditor()
Dim i As Long

    For i = 1 To MAX_QUEST
        QuestChange(i) = False
    Next
    Editor = EDITOR_NONE
    Unload frmEditor_Quest
End Sub
