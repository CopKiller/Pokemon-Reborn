Attribute VB_Name = "modMenu_Icons"
Option Explicit

Private Const DiscordLink As String = "https://discord.com/invite/RF6T6aG8Ey"
Private Const WhatsappLink As String = "https://chat.whatsapp.com/HN7n8y80VoO30D9MLHXaYM"
Private Const SiteLink As String = "http://www.pokereborn.com.br/"

Private Enum ExternalShortcut
    Discord = 1
    Whatsapp
    Site
    
    Count_Shortcut
End Enum

Public Sub DrawMenuIcons()
    Dim i As Long
    Dim X As Long, Y As Long, width As Long, height As Long
    Dim Count_Icons As Long
    
    Count_Icons = Count_Shortcut - 1

    width = GetPicWidth(Tex_Surface(gSurfaceEnum.MenuIcons)) \ 2.5
    height = GetPicHeight(Tex_Surface(gSurfaceEnum.MenuIcons)) \ 2.5
    Y = 5
    X = (Screen_Width) - ((width * Count_Icons) / 1.9)

    ' Desenha o icone do provedor externo
    For i = 1 To Count_Icons
        RenderTexture Tex_Surface(gSurfaceEnum.MenuIcons), X + (i * 40), Y, ((i - 1) * (GetPicWidth(Tex_Surface(gSurfaceEnum.MenuIcons)) / Count_Icons)), 0, (width / Count_Icons), (height), (GetPicWidth(Tex_Surface(gSurfaceEnum.MenuIcons)) / Count_Icons), GetPicHeight(Tex_Surface(gSurfaceEnum.MenuIcons))
    Next i
End Sub

Public Sub MenuIconsMouseMove()
    Dim i As Long, width As Long, height As Long, X As Long, Y As Long
    Dim Count_Icons As Long

    IsHovering = False
    
    Count_Icons = Count_Shortcut - 1

    If MenuState <> MenuStateEnum.StateNormal Then Exit Sub

    width = GetPicWidth(Tex_Surface(gSurfaceEnum.MenuIcons)) \ 2.5
    height = GetPicHeight(Tex_Surface(gSurfaceEnum.MenuIcons)) \ 2.5
    Y = 5
    X = (Screen_Width) - ((width * Count_Icons) / 1.9)

    For i = 1 To Count_Shortcut - 1
        If CursorX >= X + (i * 40) And CursorX <= X + (i * 40) + (width / Count_Icons) And CursorY >= Y And CursorY <= Y + height Then
            IsHovering = True
            MouseIcon = 1    '//Select
        End If
    Next i
End Sub

Public Sub MenuIconsMouseUp()
    Dim i As Long, width As Long, height As Long, X As Long, Y As Long
    Dim Count_Icons As Long

    If MenuState <> MenuStateEnum.StateNormal Then Exit Sub
    
    Count_Icons = Count_Shortcut - 1

    width = GetPicWidth(Tex_Surface(gSurfaceEnum.MenuIcons)) \ 2.5
    height = GetPicHeight(Tex_Surface(gSurfaceEnum.MenuIcons)) \ 2.5
    Y = 5
    X = (Screen_Width) - ((width * Count_Icons) / 1.9)

    For i = 1 To Count_Shortcut - 1
        If CursorX >= X + (i * 40) And CursorX <= X + (i * 40) + (width / Count_Icons) And CursorY >= Y And CursorY <= Y + height Then
            Select Case i
            Case ExternalShortcut.Discord
                ShellExecute frmMain.hwnd, "open", DiscordLink, vbNullString, vbNullString, conSwNormal
            Case ExternalShortcut.Whatsapp
                ShellExecute frmMain.hwnd, "open", WhatsappLink, vbNullString, vbNullString, conSwNormal
            Case ExternalShortcut.Site
                ShellExecute frmMain.hwnd, "open", SiteLink, vbNullString, vbNullString, conSwNormal
            End Select
            Exit Sub
        End If
    Next i
End Sub
