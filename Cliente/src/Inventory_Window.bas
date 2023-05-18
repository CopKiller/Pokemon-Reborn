Attribute VB_Name = "Inventory_Window"
Public Sub DrawInventory()
Dim i As Long
Dim DrawX As Long, DrawY As Long
Dim Sprite As Long

    With GUI(GuiEnum.GUI_INVENTORY)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height
        
        '//Buttons
        'Dim ButtonText As String, DrawText As Boolean
        For i = ButtonEnum.Inventory_Close To ButtonEnum.Inventory_Close
            If CanShowButton(i) Then
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
            End If
        Next
        
        '//Items
        For i = 1 To MAX_PLAYER_INV
            If i <> DragInvSlot Then
                If PlayerInv(i).Num > 0 Then
                    Sprite = Item(PlayerInv(i).Num).Sprite
                    
                    DrawX = .X + (7 + ((5 + TILE_X) * (((i - 1) Mod 5))))
                    DrawY = .Y + (37 + ((5 + TILE_Y) * ((i - 1) \ 5)))
                    
                    '//Draw Icon
                    If Sprite > 0 And Sprite <= Count_Item Then
                        RenderTexture Tex_Item(Sprite), DrawX + ((32 / 2) - (GetPicWidth(Tex_Item(Sprite)) / 2)), DrawY + ((32 / 2) - (GetPicHeight(Tex_Item(Sprite)) / 2)), 0, 0, GetPicWidth(Tex_Item(Sprite)), GetPicHeight(Tex_Item(Sprite)), GetPicWidth(Tex_Item(Sprite)), GetPicHeight(Tex_Item(Sprite))
                    End If
                    
                    RenderTexture Tex_System(gSystemEnum.UserInterface), DrawX, DrawY, 0, 8, TILE_X, TILE_Y, 1, 1, D3DColorARGB(20, 0, 0, 0)
                    
                    '//Count
                    If PlayerInv(i).Value > 1 Then
                        RenderText Font_Default, PlayerInv(i).Value, DrawX + 28 - (GetTextWidth(Font_Default, PlayerInv(i).Value)), DrawY + 14, White
                    End If
                End If
            End If
        Next
    End With
End Sub

' ***************
' ** Inventory **
' ***************
Public Sub InventoryMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_INVENTORY)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_INVENTORY
        
        '//Loop through all items
        For i = ButtonEnum.Inventory_Close To ButtonEnum.Inventory_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End If
            End If
        Next
        
        If Not SelMenu.Visible And InvUseSlot = 0 Then
            If Buttons = vbRightButton Then
                '//Inv
                i = IsInvItem(CursorX, CursorY)
                If i > 0 Then
                    OpenSelMenu SelMenuType.Inv, i
                End If
            Else
                '//Disable Drag when intrade
                If TradeIndex = 0 Then
                    '//Inv
                    i = IsInvItem(CursorX, CursorY)
                    If i > 0 Then
                        DragInvSlot = i
                        WindowPriority = GuiEnum.GUI_INVENTORY
                    End If
                End If
            End If
        End If
        
        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Public Sub InventoryMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim tmpX As Long, tmpY As Long
Dim i As Long

    With GUI(GuiEnum.GUI_INVENTORY)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        If CursorX >= .X And CursorX <= .X + .Width And CursorY >= .Y And CursorY <= .Y + .Height Then
        Else
            Exit Sub
        End If
        
        If DragInvSlot > 0 Or DragStorageSlot > 0 Then
            If WindowPriority = 0 Then
                WindowPriority = GuiEnum.GUI_INVENTORY
                If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_INVENTORY Then
                    UpdateGuiOrder GUI_INVENTORY
                End If
            End If
        End If
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_INVENTORY Then Exit Sub
        
        IsHovering = False
        
        '//Loop through all items
        For i = ButtonEnum.Inventory_Close To ButtonEnum.Inventory_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateNormal Then
                        Button(i).State = ButtonState.StateHover
        
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                End If
            End If
        Next
        
        '//Inv
        i = IsInvItem(CursorX, CursorY)
        If i > 0 Then
            IsHovering = True
            MouseIcon = 1 '//Select
            
            If Not InvItemDesc = i Then
                InvItemDesc = i
                InvItemDescTimer = GetTickCount
                InvItemDescShow = False
            End If
        End If

        '//Check for dragging
        If .InDrag Then
            tmpX = CursorX - .OldMouseX
            tmpY = CursorY - .OldMouseY
            
            '//Check if outbound
            If tmpX <= 0 Then tmpX = 0
            If tmpX >= Screen_Width - .Width Then tmpX = Screen_Width - .Width
            If tmpY <= 0 Then tmpY = 0
            If tmpY >= Screen_Height - .Height Then tmpY = Screen_Height - .Height
            
            .X = tmpX
            .Y = tmpY
        End If
    End With
End Sub

Public Sub InventoryMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_INVENTORY)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_INVENTORY Then Exit Sub
        
        '//Loop through all items
        For i = ButtonEnum.Inventory_Close To ButtonEnum.Inventory_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        Select Case i
                            Case ButtonEnum.Inventory_Close
                                If GUI(GuiEnum.GUI_INVENTORY).Visible Then
                                    GuiState GUI_INVENTORY, False
                                End If
                        End Select
                    End If
                End If
            End If
        Next
        
        '//Replace item
        If TradeIndex = 0 Then
            If DragInvSlot > 0 Then
                i = IsInvSlot(CursorX, CursorY)
                If i > 0 Then
                    SendSwitchInvSlot DragInvSlot, i
                End If
            End If
            DragInvSlot = 0
        End If
        
        '//Replace item
        If DragStorageSlot > 0 Then
            i = IsInvSlot(CursorX, CursorY)
            If i > 0 Then
                '//Check if value is greater than 1
                If PlayerInvStorage(InvCurSlot).Data(DragStorageSlot).Value > 1 Then
                    If Not GUI(GuiEnum.GUI_CHOICEBOX).Visible Then
                        OpenInputBox "Enter amount", IB_WITHDRAW, DragStorageSlot, i
                    End If
                Else
                    '//Send Withdraw
                    SendWithdrawItemTo InvCurSlot, DragStorageSlot, i
                End If
            End If
        End If
        DragStorageSlot = 0
        
        '//Check for dragging
        .InDrag = False
    End With
End Sub

Public Sub DrawInvItemDesc()
    Dim ItemName As String
    Dim ItemIcon As Long
    Dim DescString As String
    Dim LowBound As Long, UpBound As Long
    Dim ArrayText() As String
    Dim i As Integer
    Dim X As Long, Y As Long
    Dim yOffset As Long
    Dim SizeY As Long
    Dim ItemPrice As String

    If SelMenu.Visible Or DragInvSlot > 0 Then
        InvItemDesc = 0
        InvItemDescTimer = 0
        InvItemDescShow = False
        Exit Sub
    End If

    If InvItemDesc <= 0 Or InvItemDesc > MAX_PLAYER_INV Then Exit Sub
    If InvItemDescTimer + 400 > GetTickCount Then Exit Sub
    If PlayerInv(InvItemDesc).Num <= 0 Or PlayerInv(InvItemDesc).Num > MAX_ITEM Then Exit Sub
    InvItemDescShow = True

    ItemIcon = Item(PlayerInv(InvItemDesc).Num).Sprite
    ItemName = "~ " & Trim$(Item(PlayerInv(InvItemDesc).Num).Name) & " ~"
    DescString = Trim$(Item(PlayerInv(InvItemDesc).Num).Desc)    '"A device for catching wild Pokemon. It is thrown like a ball at the target. It is designed as a capsule system"

    If Item(PlayerInv(InvItemDesc).Num).IsCash = NO Then
        ItemPrice = "Price: " & Item(PlayerInv(InvItemDesc).Num).Price
    Else
        ItemPrice = "Price: Non Sellable"
    End If

    '//Make sure that loading text have something to draw
    If Len(DescString) < 0 Then Exit Sub

    '//Wrap the text
    WordWrap_Array Font_Default, DescString, 150, ArrayText

    '//we need these often
    LowBound = LBound(ArrayText)
    UpBound = UBound(ArrayText)

    SizeY = 25 + ((UpBound + 1) * 16)

    RenderTexture Tex_System(gSystemEnum.UserInterface), GUI(GuiEnum.GUI_INVENTORY).X + 6, GUI(GuiEnum.GUI_INVENTORY).Y + 36, 0, 8, 182, 219, 1, 1, D3DColorARGB(180, 0, 0, 0)

    RenderTexture Tex_Item(ItemIcon), GUI(GuiEnum.GUI_INVENTORY).X + GUI(GuiEnum.GUI_INVENTORY).Width / 2 - (GetPicHeight(Tex_Item(ItemIcon)) / 2), GUI(GuiEnum.GUI_INVENTORY).Y + 8 + ((219 * 0.5) - (SizeY * 0.5)), 0, 0, GetPicWidth(Tex_Item(ItemIcon)), GetPicHeight(Tex_Item(ItemIcon)), GetPicWidth(Tex_Item(ItemIcon)), GetPicHeight(Tex_Item(ItemIcon))

    RenderText Font_Default, ItemName, GUI(GuiEnum.GUI_INVENTORY).X + 6 + ((182 * 0.5) - (GetTextWidth(Font_Default, ItemName) * 0.5)), GUI(GuiEnum.GUI_INVENTORY).Y + 36 + ((219 * 0.5) - (SizeY * 0.5)), White
    
    RenderText Font_Default, ItemPrice, GUI(GuiEnum.GUI_INVENTORY).X + 6 + ((182 * 0.5) - (GetTextWidth(Font_Default, ItemName) * 0.5)), GUI(GuiEnum.GUI_INVENTORY).Y + 150 + ((219 * 0.5) - (SizeY * 0.5)), White

    '//Reset
    yOffset = 25
    '//Loop to all items
    For i = LowBound To UpBound
        '//Set Location
        '//Keep it centered
        X = GUI(GuiEnum.GUI_INVENTORY).X + 6 + ((182 * 0.5) - (GetTextWidth(Font_Default, Trim$(ArrayText(i))) * 0.5))
        Y = GUI(GuiEnum.GUI_INVENTORY).Y + 36 + ((219 * 0.5) - (SizeY * 0.5)) + yOffset

        '//Render the text
        RenderText Font_Default, Trim$(ArrayText(i)), X, Y, White

        '//Increase the location for each line
        yOffset = yOffset + 16
    Next
End Sub

