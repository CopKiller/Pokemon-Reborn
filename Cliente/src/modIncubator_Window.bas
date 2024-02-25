Attribute VB_Name = "modIncubator_Window"
Option Explicit

' Método que desenha a janela
Public Sub DrawIncubator()
    Dim i As Long
    Dim X As Long, Y As Long
    Dim SString As String

    With GUI(GuiEnum.GUI_INCUBATOR)

        ' Certifica que está visível
        If Not .Visible Then Exit Sub

        ' Desenha a janela
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height

        ' Desenha o titulo da janela
        Select Case tmpCurLanguage
        Case LANG_PT: SString = "Incubadora"
        Case LANG_EN: SString = "Incubator"
        Case LANG_ES: SString = "Incubator"
        End Select
        RenderText Font_Default, SString, .X + 25, .Y + 5, White

    End With
End Sub

' Método dos cliques na janela
Public Sub IncubatorMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Byte

    With GUI(GuiEnum.GUI_INCUBATOR)

        ' Certifica que está visível
        If Not .Visible Then Exit Sub

        ' Ordena a janela ao ser clicada
        UpdateGuiOrder GUI_INCUBATOR

        ' Verifica todos os itens
        'For i = ButtonEnum.Login_Confirm To ButtonEnum.Login_Confirm
        i = ButtonEnum.Incubator_Close
        If CanShowButton(i) Then
            If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                If Button(i).State = ButtonState.StateHover Then
                    Button(i).State = ButtonState.StateClick
                End If
            End If
        End If
        'Next

        ' Verifica se foi movido
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

' Método ao passar o mouse por cima dos itens
Public Sub IncubatorMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Byte
    Dim tmpX As Long, tmpY As Long

    With GUI(GuiEnum.GUI_INCUBATOR)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub

        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_INCUBATOR Then Exit Sub

        IsHovering = False

        '//Loop through all items
        'For i = ButtonEnum.Login_Confirm To ButtonEnum.Login_Confirm
        i = ButtonEnum.Incubator_Close
        If CanShowButton(i) Then
            If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                If Button(i).State = ButtonState.StateNormal Then
                    Button(i).State = ButtonState.StateHover

                    IsHovering = True
                    MouseIcon = 1    '//Select
                End If
            End If
        End If
        'Next

        ' Verifica se foi movido
        If .InDrag Then
            tmpX = CursorX - .OldMouseX
            tmpY = CursorY - .OldMouseY

            If tmpX <= 0 Then tmpX = 0
            If tmpX >= Screen_Width - .Width Then tmpX = Screen_Width - .Width
            If tmpY <= 0 Then tmpY = 0
            If tmpY >= Screen_Height - .Height Then tmpY = Screen_Height - .Height

            .X = tmpX
            .Y = tmpY
        End If
    End With
End Sub

Public Sub IncubatorMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Byte
    Dim FoundError As Boolean

    With GUI(GuiEnum.GUI_INCUBATOR)

        ' Certifica que está visível
        If Not .Visible Then Exit Sub

        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_INCUBATOR Then Exit Sub

        'For i = ButtonEnum.Login_Confirm To ButtonEnum.Login_Confirm
        i = ButtonEnum.Incubator_Close
        If CanShowButton(i) Then
            If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                If Button(i).State = ButtonState.StateClick Then
                    Button(i).State = ButtonState.StateNormal
                    Select Case i
                    Case ButtonEnum.Login_Confirm
                    End Select
                End If
            End If
        End If

        .InDrag = False
    End With
End Sub

