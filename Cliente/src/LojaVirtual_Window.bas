Attribute VB_Name = "LojaVirtual_Window"
Option Explicit

Public LojaVirtual_Skins() As LojaVirtualRec
Public LojaVirtual_Mounts() As LojaVirtualRec
Public LojaVirtual_Items() As LojaVirtualRec

Private LojaVirtual_Select As LojaVirtualRec

Private Type LojaVirtualRec
    ItemNum As Long
    ItemPrice As Long
End Type

Public Sub InitLojaVirtual()
    Dim i As Byte
    
    ReDim LojaVirtual_Skins(1 To 8)
    ReDim LojaVirtual_Mounts(1 To 8)
    ReDim LojaVirtual_Items(1 To 8)
    
    For i = LBound(LojaVirtual_Skins) To UBound(LojaVirtual_Skins)
        LojaVirtual_Skins(i).ItemNum = i
        LojaVirtual_Skins(i).ItemPrice = 20
        
        LojaVirtual_Mounts(i).ItemNum = i + UBound(LojaVirtual_Skins)
        LojaVirtual_Mounts(i).ItemPrice = 40
        
        LojaVirtual_Items(i).ItemNum = i + (UBound(LojaVirtual_Skins) * 2)
        LojaVirtual_Items(i).ItemPrice = 60
    Next i
End Sub


Private Function PlayerHaveCashValue(ByVal Price As Long) As Boolean
    PlayerHaveCashValue = False

    If Player(MyIndex).Cash >= Price Then
        PlayerHaveCashValue = True
    End If
End Function

Public Sub DrawLojaVirtual()
    Dim i As Long
    Dim tmpX As Long, tmpY As Long
    Dim ColourOpacity As Long
    Dim X As Long, Z As Long
    Dim CaptionBuy As String

    With GUI(GuiEnum.GUI_LOJAVIRTUAL)
        '//Verifica se a janela está visivel.
        If Not .Visible Then Exit Sub

        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height

        '//Define o nome dos botões conforme tradução
        Select Case tmpCurLanguage
        Case LANG_PT: CaptionBuy = "Comprar"
        Case LANG_EN: CaptionBuy = "Purchase"
        Case LANG_ES: CaptionBuy = "Purchase"
        End Select

        '//Buttons
        For i = ButtonEnum.LojaVirtual_Close To ButtonEnum.LojaVirtual_Slot8
            If CanShowButton(i) Then
                '//Renderiza o botão de compra
                If i = ButtonEnum.LojaVirtual_Buy Then
                    If LojaVirtual_Select.ItemNum > 0 And LojaVirtual_Select.ItemNum <= MAX_ITEM Then
                        '//O jogador tem o valor?
                        If PlayerHaveCashValue(LojaVirtual_Select.ItemPrice) Then
                            '//Cor normal
                            ColourOpacity = D3DColorARGB(255, 255, 255, 255)
                        Else
                            '//Cor opaca caso não tenha o valor do item
                            ColourOpacity = D3DColorARGB(255, 180, 60, 180)
                        End If

                        '//Renderiza o BackGround button
                        RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height, ColourOpacity
                        '//Renderiza o Texto do button buy
                        RenderText Font_Default, CaptionBuy, .X + Button(i).X + (Button(i).Width / 2) - (GetTextWidth(Font_Default, CaptionBuy) / 2) - 3, (.Y + Button(i).Y) + Button(i).Height / 2 - 11, White, , 255
                    End If
                ElseIf i >= ButtonEnum.LojaVirtual_Slot1 And i <= ButtonEnum.LojaVirtual_Slot8 Then
                    '//Cor normal
                    ColourOpacity = D3DColorARGB(255, 255, 255, 255)

                    '//Renderiza o BackGround do item/skin/mount
                    RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height, ColourOpacity

                    '//Obter o índice real do slot 1 ao 6 em uma variavel
                    X = i - ButtonEnum.LojaVirtual_Slot1 + 1
                    '//Renderiza o Icone do item/skin/mount
                    If Item(LojaVirtual_Items(X).ItemNum).Sprite > 0 Then
                        RenderTexture Tex_Item(Item(LojaVirtual_Items(X).ItemNum).Sprite), .X + Button(i).X + 12, .Y + Button(i).Y + 10, 0, 0, GetPicWidth(Tex_Item(Item(LojaVirtual_Items(X).ItemNum).Sprite)), GetPicHeight(Tex_Item(Item(LojaVirtual_Items(X).ItemNum).Sprite)), GetPicWidth(Tex_Item(Item(LojaVirtual_Items(X).ItemNum).Sprite)), GetPicHeight(Tex_Item(Item(LojaVirtual_Items(X).ItemNum).Sprite)), ColourOpacity
                    End If
                    '//Renderiza o nome do item
                    RenderText Font_Default, Trim$(Item(LojaVirtual_Items(X).ItemNum).Name), .X + Button(i).X + 44, (.Y + Button(i).Y) + Button(i).Height / 2 - 13, DarkGrey, , 255

                Else
                    '//Close Button
                    RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
                End If
            End If
        Next i

        '//Renderização dos icones dos items + nome
        ' For Z = 1 To X
        '//Renderiza o Icone do item/skin/mount
        '     RenderTexture Tex_Item(Item(LojaVirtual_Items(Z).ItemNum).Sprite), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height, ColourOpacity
        '//Renderiza o nome do item
        '    RenderText Font_Default, CaptionBuy, .X + Button(i).X + (Button(i).Width / 2) - (GetTextWidth(Font_Default, CaptionBuy) / 2) - 3, (.Y + Button(i).Y) + Button(i).Height / 2 - 11, White, , 255
        ' Next Z

    End With
End Sub

Public Sub LojaVirtualMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long

    With GUI(GuiEnum.GUI_LOJAVIRTUAL)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub

        '//Set to top most
        UpdateGuiOrder GUI_LOJAVIRTUAL

        '//Loop through all items
        For i = ButtonEnum.LojaVirtual_Close To ButtonEnum.LojaVirtual_Slot8
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick

                      '  Select Case i
                      '  Case ButtonEnum.Rank_ScrollUp
                      '      RankingScrollUp = True
                      '      RankingScrollDown = False
                      '      RankingScrollTimer = GetTickCount
                      '  Case ButtonEnum.Rank_ScrollDown
                      '      RankingScrollUp = False
                      '      RankingScrollDown = True
                      '      RankingScrollTimer = GetTickCount
                      '  End Select
                    End If
                End If
            End If
        Next

        '//Check for scroll
        'If CursorX >= .X + 7 And CursorX <= .X + 7 + 19 And CursorY >= .Y + RankingScrollStartY + ((RankingScrollEndY - RankingScrollSize) - RankingScrollY) And CursorY <= .Y + RankingScrollStartY + ((RankingScrollEndY - RankingScrollSize) - RankingScrollY) + RankingScrollSize Then
        '    RankingScrollHold = True
        'End If

        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Public Sub LojaVirtualMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmpX As Long, tmpY As Long
    Dim i As Long

    With GUI(GuiEnum.GUI_LOJAVIRTUAL)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub

        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_LOJAVIRTUAL Then Exit Sub

        IsHovering = False

        '//Loop through all items
        For i = ButtonEnum.LojaVirtual_Close To ButtonEnum.LojaVirtual_Slot8
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    '//Renderiza o botão de compra
                    If i = ButtonEnum.LojaVirtual_Buy Then
                        If LojaVirtual_Select.ItemNum > 0 And LojaVirtual_Select.ItemNum <= MAX_ITEM Then
                            '//O jogador tem o valor?
                            If PlayerHaveCashValue(LojaVirtual_Select.ItemPrice) Then
                                If Button(i).State = ButtonState.StateNormal Then
                                    Button(i).State = ButtonState.StateHover

                                    IsHovering = True
                                    MouseIcon = 1    '//Select
                                End If
                            End If
                        End If
                    Else
                        If Button(i).State = ButtonState.StateNormal Then
                            Button(i).State = ButtonState.StateHover

                            IsHovering = True
                            MouseIcon = 1    '//Select
                        End If
                    End If
                End If
            End If
        Next

        '//Check for scroll
        'If RankingHighIndex > RankingViewLine Then
        '    If CursorX >= .X + 7 And CursorX <= .X + 7 + 19 And CursorY >= .Y + RankingScrollStartY + ((RankingScrollEndY - RankingScrollSize) - RankingScrollY) And CursorY <= .Y + RankingScrollStartY + ((RankingScrollEndY - RankingScrollSize) - RankingScrollY) + RankingScrollSize Then
        '        IsHovering = True
        '        MouseIcon = 1    '//Select
        '    End If

        '    '//Scroll moving
        '    If RankingScrollHold Then
        '        '//Upward
        '        If CursorY < .Y + RankingScrollStartY + ((RankingScrollEndY - RankingScrollSize) - RankingScrollY) + (RankingScrollSize / 2) Then
        '            If RankingScrollY < RankingScrollEndY - RankingScrollSize Then
        '                RankingScrollY = (CursorY - (.Y + RankingScrollStartY + (RankingScrollEndY - RankingScrollSize)) - (RankingScrollSize / 2)) * -1
        '                If RankingScrollY >= RankingScrollEndY - RankingScrollSize Then RankingScrollY = RankingScrollEndY - RankingScrollSize
        '            End If
        '        End If
        '//Downward
        '        If CursorY > .Y + RankingScrollStartY + ((RankingScrollEndY - RankingScrollSize) - RankingScrollY) + RankingScrollSize - (RankingScrollSize / 2) Then
        '            If RankingScrollY > 0 Then
        '                RankingScrollY = (CursorY - (.Y + RankingScrollStartY + (RankingScrollEndY - RankingScrollSize)) - RankingScrollSize + (RankingScrollSize / 2)) * -1
        '                If RankingScrollY <= 0 Then RankingScrollY = 0
        '            End If
        '        End If

        '        RankingScrollCount = (RankingScrollLength - RankingScrollY)
        '        RankingViewCount = ((RankingScrollCount / MaxRankingViewLine) / (RankingScrollLength / MaxRankingViewLine)) * MaxRankingViewLine
        '    End If
        'End If

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

Public Sub LojaVirtualMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_LOJAVIRTUAL)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_LOJAVIRTUAL Then Exit Sub
        
        '//Loop through all items
        For i = ButtonEnum.LojaVirtual_Close To ButtonEnum.LojaVirtual_Slot8
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        Select Case i
                            Case ButtonEnum.LojaVirtual_Close
                                If GUI(GuiEnum.GUI_LOJAVIRTUAL).Visible Then
                                    GuiState GUI_LOJAVIRTUAL, False
                                End If
                        End Select
                    End If
                End If
            End If
        Next
        
        '//Ranking Scroll
        RankingScrollHold = False

        '//Check for dragging
        .InDrag = False
    End With
End Sub

