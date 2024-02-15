Attribute VB_Name = "modMapTravel_Window"
Option Explicit

Public Sub DrawPlayerTravel()
    Dim i As Long, IconX As Long, IconY As Long, SrcIconX As Long, SrcIconY As Long, SrcWidth As Long, SrcHeight As Long
    Dim Colour As Long, SString As String

    With GUI(GuiEnum.GUI_MAP)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub

        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .y, .StartX, .StartY, .Width, .Height, .Width, .Height

        ' Titulo da janela
        Select Case tmpCurLanguage
        Case LANG_PT: SString = "Kanto Região!"
        Case LANG_EN: SString = "Kanto Region!"
        Case LANG_ES: SString = "Kanto Region!"
        End Select
        RenderText Font_Default, SString, .X + 25, .y + 5, White

        i = ButtonEnum.MapTravel_Close
        If CanShowButton(i) Then
            '//Close Button
            RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .y + Button(i).y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
        End If

        For i = 1 To MAX_MAP
            If Player(MyIndex).PlayerTravel(i).DataExist And Player(MyIndex).PlayerTravel(i).mapName <> vbNullString Then
                SrcIconX = Player(MyIndex).PlayerTravel(i).SrcPosX
                SrcIconY = Player(MyIndex).PlayerTravel(i).SrcPosY
                SrcWidth = Player(MyIndex).PlayerTravel(i).SrcWidth
                SrcHeight = Player(MyIndex).PlayerTravel(i).SrcHeight
                IconX = Player(MyIndex).PlayerTravel(i).IconPosX
                IconY = Player(MyIndex).PlayerTravel(i).IconPosY

                If GetPlayerMapUnlocked(i) = False Then
                    Colour = BrightRed
                Else
                    Colour = Yellow

                    RenderTexture Tex_Gui(.Pic), .X + IconX, .y + IconY, SrcIconX, SrcIconY, SrcWidth, SrcHeight, SrcWidth, SrcHeight
                End If

                RenderText Ui_Default, Player(MyIndex).PlayerTravel(i).mapName, .X + IconX - (GetTextWidth(Ui_Default, Player(MyIndex).PlayerTravel(i).mapName) / 2) + (SrcWidth / 2), .y + IconY - 22, Colour
            End If
        Next i
    End With
End Sub

Public Sub PlayerTravelMouseDown(Buttons As Integer, Shift As Integer, X As Single, y As Single)
    Dim i As Long, IconX As Long, IconY As Long, SrcIconX As Long, SrcIconY As Long, SrcWidth As Long, SrcHeight As Long

    With GUI(GuiEnum.GUI_MAP)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub

        '//Set to top most
        UpdateGuiOrder GUI_MAP

        i = ButtonEnum.MapTravel_Close
        If CanShowButton(i) Then
            If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .y + Button(i).y And CursorY <= .y + Button(i).y + Button(i).Height Then
                If Button(i).State = ButtonState.StateHover Then
                    Button(i).State = ButtonState.StateClick
                End If
            End If
        End If

        For i = 1 To MAX_MAP
            If Player(MyIndex).PlayerTravel(i).DataExist And Player(MyIndex).PlayerTravel(i).mapName <> vbNullString Then
                If GetPlayerMapUnlocked(i) = True Then
                    SrcIconX = Player(MyIndex).PlayerTravel(i).SrcPosX
                    SrcIconY = Player(MyIndex).PlayerTravel(i).SrcPosY
                    SrcWidth = Player(MyIndex).PlayerTravel(i).SrcWidth
                    SrcHeight = Player(MyIndex).PlayerTravel(i).SrcHeight
                    IconX = Player(MyIndex).PlayerTravel(i).IconPosX
                    IconY = Player(MyIndex).PlayerTravel(i).IconPosY

                    If CursorX >= .X + IconX And CursorX <= .X + IconX + SrcWidth And CursorY >= .y + IconY And CursorY <= .y + IconY + SrcHeight Then
                        
                        'Add process to warp
                        PlayerTravelSlot = i
                        Select Case tmpCurLanguage
                        Case LANG_PT: OpenChoiceBox "Deseja teleportar até " & Player(MyIndex).PlayerTravel(i).mapName & " por " & GetPlayerMapCostValue(i) & " Moneys?", CB_TRAVEL
                        Case LANG_EN: OpenChoiceBox "Deseja teleportar até " & Player(MyIndex).PlayerTravel(i).mapName & " por " & GetPlayerMapCostValue(i) & " Moneys?", CB_TRAVEL
                        Case LANG_ES: OpenChoiceBox "Deseja teleportar até " & Player(MyIndex).PlayerTravel(i).mapName & " por " & GetPlayerMapCostValue(i) & " Moneys?", CB_TRAVEL
                        
                        End Select
                    End If
                End If
            End If
        Next i

        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Public Sub PlayerTravelMouseMove(Buttons As Integer, Shift As Integer, X As Single, y As Single)
    Dim tmpX As Long, tmpY As Long
    Dim i As Long, IconX As Long, IconY As Long, SrcIconX As Long, SrcIconY As Long, SrcWidth As Long, SrcHeight As Long

    With GUI(GuiEnum.GUI_MAP)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub

        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_MAP Then Exit Sub

        IsHovering = False

        i = ButtonEnum.MapTravel_Close
        If CanShowButton(i) Then
            If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .y + Button(i).y And CursorY <= .y + Button(i).y + Button(i).Height Then
                If Button(i).State = ButtonState.StateNormal Then
                    Button(i).State = ButtonState.StateHover

                    IsHovering = True
                    MouseIcon = 1    '//Select
                End If
            End If
        End If

        For i = 1 To MAX_MAP
            If Player(MyIndex).PlayerTravel(i).DataExist Then
                If GetPlayerMapUnlocked(i) = True Then
                    SrcIconX = Player(MyIndex).PlayerTravel(i).SrcPosX
                    SrcIconY = Player(MyIndex).PlayerTravel(i).SrcPosY
                    SrcWidth = Player(MyIndex).PlayerTravel(i).SrcWidth
                    SrcHeight = Player(MyIndex).PlayerTravel(i).SrcHeight
                    IconX = Player(MyIndex).PlayerTravel(i).IconPosX
                    IconY = Player(MyIndex).PlayerTravel(i).IconPosY

                    If CursorX >= .X + IconX And CursorX <= .X + IconX + SrcWidth And CursorY >= .y + IconY And CursorY <= .y + IconY + SrcHeight Then
                        'Add process to mousemove
                        IsHovering = True
                        MouseIcon = 1    '//Select
                    End If
                End If
            End If
        Next i



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
            .y = tmpY
        End If
    End With
End Sub

Public Sub PlayerTravelMouseUp(Buttons As Integer, Shift As Integer, X As Single, y As Single)
    Dim i As Long

    With GUI(GuiEnum.GUI_MAP)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub

        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_MAP Then Exit Sub

        i = ButtonEnum.MapTravel_Close
        If CanShowButton(i) Then
            If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .y + Button(i).y And CursorY <= .y + Button(i).y + Button(i).Height Then
                If Button(i).State = ButtonState.StateClick Then
                    Button(i).State = ButtonState.StateNormal
                    If GUI(GuiEnum.GUI_MAP).Visible = True Then
                        GuiState GUI_MAP, False
                    End If
                End If
            End If
        End If

        '//Check for dragging
        .InDrag = False
    End With
End Sub
