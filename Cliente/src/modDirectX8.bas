Attribute VB_Name = "modDirectX8"
Option Explicit

'//DirectX8 Object
Private dX         As DirectX8                '//The master Object, everything comes from here
Private D3D        As Direct3D8               '//This controls all things 3D
Public D3DDevice   As Direct3DDevice8         '//This actually represents the hardware doing the rendering
Public D3DX        As D3DX8                   '//A helper library

Private DispMode   As D3DDISPLAYMODE          '//Describes our Display Mode
Private D3DWindow  As D3DPRESENT_PARAMETERS   '//Describes our Viewport

'//This is the Flexible-Vertex-Format description for a 2D vertex (Transformed and Lit)
Public Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
Public Const FVF_Size As Long = 28

'//This structure describes a transformed and lit vertex - it's identical to the DirectX7 type "D3DTLVERTEX"
Public Type TLVERTEX
    X        As Single
    Y        As Single
    z        As Single
    rhw      As Single
    Color    As Long
    tu       As Single
    tv       As Single
End Type

'//Image info holder
Private Type TextureRec
    Texture     As Direct3DTexture8
    Width       As Long
    Height      As Long
    Path        As String
    UnloadTimer As Long
    Loaded      As Boolean
End Type

'//Temporary Image info holder while loading
Private Type D3DXIMAGE_INFO_A
    Width           As Long
    Height          As Long
    Depth           As Long
    MipLevels       As Long
    Format          As CONST_D3DFORMAT
    ResourceType    As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long
End Type

Public Const DegreeToRadian As Single = 0.0174532919296
Public Const RadianToDegree As Single = 57.2958300962816

'//This is where all texture are being stored
Private GlobalTexture() As TextureRec
Private GlobalTextureCount As Long

Public CurrentTexture As Long      '//This make sure that we are not rendering the same texture every time

'//Textures
Public Tex_System() As Long
Public Tex_Surface() As Long
Public Tex_Gui() As Long
Public Tex_Character() As Long
Public Tex_PlayerSprite_N() As Long
Public Tex_PlayerSprite_D() As Long
Public Tex_PlayerSprite_B() As Long
Public Tex_Tileset() As Long
Public Tex_MapAnim() As Long
Public Tex_Pokemon() As Long
Public Tex_ShinyPokemon() As Long
Public Tex_Item() As Long
Public Tex_Misc() As Long
Public Tex_PokemonIcon() As Long
Public Tex_Animation() As Long
Public Tex_Weather() As Long
Public Tex_PokemonPortrait() As Long
Public Tex_ShinyPokemonPortrait() As Long
Public Tex_PokemonTypes() As Long
Public Tex_PokemonTypesSymbol() As Long

'//Texture Count
Public Count_System As Long
Public Count_Surface As Long
Public Count_Gui As Long
Public Count_Character As Long
Public Count_PlayerSprite_N As Long
Public Count_PlayerSprite_D As Long
Public Count_PlayerSprite_B As Long
Public Count_Tileset As Long
Public Count_MapAnim As Long
Public Count_Pokemon As Long
Public Count_ShinyPokemon As Long
Public Count_Item As Long
Public Count_Misc As Long
Public Count_PokemonIcon As Long
Public Count_Animation As Long
Public Count_Weather As Long
Public Count_PokemonPortrait As Long
Public Count_ShinyPokemonPortrait As Long
Public Count_PokemonTypes As Long
Public Count_PokemonTypesSymbol As Long

'//Texture Path
Public Const Texture_Path As String = "\data\themes\"
Public Const System_Path As String = "\textures\"
Public Const Surface_Path As String = "\textures\"
Public Const Gui_Path As String = "\textures\"
Public Const Character_Path As String = "\data\resources\character-sprites\"
Public Const PlayerSprite_Path As String = "\data\resources\player-sprites\"
Public Const Tileset_Path As String = "\data\resources\world-tiles\"
Public Const MapAnim_Path As String = "\data\resources\map-animation\"
Public Const Pokemon_Path As String = "\data\resources\pokemon\"
Public Const Item_Path As String = "\data\resources\item\"
Public Const Misc_Path As String = "\data\resources\misc\"
Public Const PokemonIcon_Path As String = "\data\resources\pokemon\"
Public Const Animation_Path As String = "\data\resources\animation\"
Public Const Weather_Path As String = "\data\resources\weather\"
Public Const PokemonPortrait_Path As String = "\data\resources\pokemon\portrait\"
Public Const PokemonTypes_Path As String = "\data\resources\poke-types\"
Public Const PokemonTypesSymbol_Path As String = "\data\resources\poke-types\icons\"

'//Global
Private Const MenuUi_Texture As Byte = 1
Private Const GameUi_Texture As Byte = 10
Private Const Hud As Byte = 12

'//Misc
Public Const Misc_Chatbubble As Byte = 1
Public Const Misc_Bar As Byte = 2
Public Const Misc_MoveSelector As Byte = 3
Public Const Misc_Pokeball As Byte = 4
Public Const Misc_Language As Byte = 5
Public Const Misc_Status As Byte = 6
Public Const Misc_PokeSelect As Byte = 7
Public Const Misc_ShinySummary As Byte = 8
Public Const Misc_CategoryTypes As Byte = 9

' ********************
' ** Initialization **
' ********************
'//Initialise DirectX
Public Sub InitDirectX()
    Set dX = New DirectX8           '//Create our Master Object
    Set D3D = dX.Direct3DCreate()   '//Make our Master Object create the Direct3D Interface
    Set D3DX = New D3DX8            '//Create our helper library..
    
    '//Check for supported hardware
    If Not EnumerateDispModes Then
        MsgBox "Could not find display. Exiting...", vbCritical
        DestroyDirectX
        End
    End If
    
    '//Update Resolution
    If Not UpdateScreenResolution Then
        MsgBox "Failed to load resolution settings. Exiting...", vbCritical
        DestroyDirectX
        End
    End If
    
    If Not InitD3DDevice(D3DCREATE_PUREDEVICE Or D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
        If Not InitD3DDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
            If Not InitD3DDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then
                If Not InitD3DDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then
                    MsgBox "Could not init D3DDevice8. Exiting...", vbCritical
                    DestroyDirectX
                    End
                End If
            End If
        End If
    End If
    
    CacheTextures
    InitRenderState
    InitFont
End Sub

Private Function EnumerateDispModes() As Boolean
Dim nModes As Integer
Dim i As Integer, X As Integer
Dim TmpResolution() As String
Dim TmpSize() As String
Dim DefaultRes As Byte

    On Error GoTo errorHandler
    
    '//Check if we already have a cache for resolution
    'If LoadResolution Then
    '    EnumerateDispModes = True
    '    Exit Function
    'End If

    '//Count how many available resolution do we have
    nModes = D3D.GetAdapterModeCount(D3DADAPTER_DEFAULT)
    
    '//Set Adapter to 32Bit
    DispMode.Format = D3DFMT_X8R8G8B8
    
    '//Set count
    ReDim TmpResolution(0 To nModes - 1)
    
    For i = 0 To nModes - 1 '//Cycle through them and collect the data...
        Call D3D.EnumAdapterModes(D3DADAPTER_DEFAULT, i, DispMode)
    
        '//Check that the device is acceptable and valid...
        If D3D.CheckDeviceType(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, DispMode.Format, False) >= 0 Then
            '//add the display size to the temp list
            '//Make sure it doesn't go down below the default size
            If DispMode.Width >= Default_ScreenWidth And DispMode.Height >= Default_ScreenHeight Then
                TmpResolution(i) = DispMode.Width & "x" & DispMode.Height
            End If
        End If
    Next i
    
    '//Remove empty string
    Call strArrRemoveNull(TmpResolution)
    
    '//Remove Duplicate Resolution
    Call strArrRemoveDuplicate(TmpResolution)
    
    '//Set count of Resolution
    ReDim Resolution.ResolutionSize(LBound(TmpResolution) To UBound(TmpResolution))
    Resolution.MaxResolution = UBound(TmpResolution)
    
    '//Update Resolution
    For i = LBound(TmpResolution) To UBound(TmpResolution)
        '//Split the resolution text to value so that we can use them
        TmpSize = Split(TmpResolution(i), "x")
        For X = LBound(TmpSize) To UBound(TmpSize)
            If X = 0 Then ' Width
                Resolution.ResolutionSize(i).Width = TmpSize(X)
            ElseIf X = 1 Then ' Height
                Resolution.ResolutionSize(i).Height = TmpSize(X)
            End If
        Next
    Next

    '//Save cache for resolution
    'SaveResolution
    
    '//We succeed
    EnumerateDispModes = True
    
    Exit Function
errorHandler:
    '//We failed
    EnumerateDispModes = False
End Function

'//Resolution
Private Function UpdateScreenResolution() As Boolean
    
    '//Set Resolution
    If GameSetting.Fullscreen = YES Then
        Screen_Width = GetSystemMetrics(SM_CXSCREEN)
        Screen_Height = GetSystemMetrics(SM_CYSCREEN)
    Else
        Screen_Width = GameSetting.Width
        Screen_Height = GameSetting.Height
    End If
    
    '//Make sure to update viewport
    ViewPortInit = False

    '//Set Window Size
    Form_Width = (Screen_Width * Screen.TwipsPerPixelX) + (frmMain.Width - (frmMain.scaleWidth * Screen.TwipsPerPixelX))
    Form_Height = (Screen_Height * Screen.TwipsPerPixelY) + (frmMain.Height - (frmMain.scaleHeight * Screen.TwipsPerPixelY))

    '//If setting is on fullscreen mode then, let's put the window on the top most to prevent clicking other program
    If GameSetting.Fullscreen = YES Then
        Call SetWindowPos(frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    End If

    '//We succeed
    UpdateScreenResolution = True
    
    Exit Function
errorHandler:
    '//We failed
    UpdateScreenResolution = False
End Function

Private Sub InitRenderState()
    D3DDevice.SetVertexShader FVF                   '//Set the vertex shader to use our vertex format

    D3DDevice.SetRenderState D3DRS_LIGHTING, YES  '//Transformed and lit vertices dont need lighting so we disable it...
    '//For Transparencies
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    D3DDevice.SetRenderState D3DRS_ZENABLE, False   '//We need to enable our Z Buffer
    
    D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    D3DDevice.SetRenderState D3DRS_ZWRITEENABLE, False
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
    D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
    D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
End Sub

'//Initialise Device, It'll return true for success, false if there was an error
Private Function InitD3DDevice(D3DCREATEFLAGS As CONST_D3DCREATEFLAGS) As Boolean
    On Error GoTo errorHandler
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode    '//Retrieve the current display Mode
    
    DispMode.Format = D3DFMT_X8R8G8B8                           '//Set Adapter to 32Bit
    DispMode.Width = Screen_Width
    DispMode.Height = Screen_Height
    
    '//Check if fullscreen
    If GameSetting.Fullscreen = YES Then
        D3DWindow.Windowed = NO                                 '//Tell it we're using Fullscreen Mode
        D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP               '//We'll refresh when the monitor does
    ElseIf GameSetting.Fullscreen = NO Then
        D3DWindow.Windowed = YES                                '//Tell it we're using Windowed Mode
        D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY               '//We'll refresh when the monitor does
    Else    '//Got Hacked
        GoTo errorHandler
    End If
    D3DWindow.BackBufferFormat = DispMode.Format            '//We'll use the format we just retrieved...
    D3DWindow.BackBufferCount = 1                           '//1 backbuffer only
    D3DWindow.BackBufferHeight = Screen_Height
    D3DWindow.BackBufferWidth = Screen_Width
    D3DWindow.hDeviceWindow = frmMain.hwnd
    
    '//We need to enable our Z Buffer
    D3DWindow.EnableAutoDepthStencil = 1
    D3DWindow.AutoDepthStencilFormat = D3DFMT_D16 '//16 bit Z-Buffer

    If Not D3DDevice Is Nothing Then Set D3DDevice = Nothing
    '//Select a appropriate hardware on what the computer can do
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hwnd, D3DCREATEFLAGS, D3DWindow)

    '//We succeeded
    InitD3DDevice = True
    
    Exit Function
errorHandler:
    '//We failed
    Set D3DDevice = Nothing
    InitD3DDevice = False
End Function

'//Unloading DirectX
Public Sub DestroyDirectX()
Dim i As Long

    If Not D3DDevice Is Nothing Then Set D3DDevice = Nothing
    If Not D3D Is Nothing Then Set D3D = Nothing
    If Not dX Is Nothing Then Set dX = Nothing
    
    '//Unload all textures
    For i = 1 To GlobalTextureCount
        Set GlobalTexture(i).Texture = Nothing
    Next
End Sub

'//We have to Cache all textures so we don't have to determine their path when rendering
Private Sub CacheTextures()
Dim i As Long
Dim TextureName As String

    ' ********************
    ' ** System Texture **
    ' ********************
    '//Let's start counting at 1
    Count_System = 1
    TextureName = "user-interface"
    Do While FileExist(App.Path & Texture_Path & Trim$(GameSetting.ThemePath) & System_Path & TextureName & GFX_EXT)
        '//If file exist, let's check the other count
        Count_System = Count_System + 1
        '//Add a proper name of the texture
        Select Case Count_System
            Case 1: TextureName = "user-interface"
            Case 2: TextureName = "cursor"
            Case 3: TextureName = "cursor-load"
            Case Else: TextureName = Count_System
        End Select
    Loop
    '//If it reach at this point, it means that the file on the certain count is missing so we must remove it
    Count_System = Count_System - 1
    '//Make sure that we have texture to set
    If Count_System > 0 Then
        '//Set texture range
        ReDim Tex_System(1 To Count_System)
        '//Set the path range to GlobalTexture
        For i = 1 To Count_System
            '//Add a proper name of the texture
            Select Case i
                Case 1: TextureName = "user-interface"
                Case 2: TextureName = "cursor"
                Case 3: TextureName = "cursor-load"
                Case Else: TextureName = i
            End Select
            Tex_System(i) = SetTexturePath(App.Path & Texture_Path & Trim$(GameSetting.ThemePath) & System_Path & TextureName & GFX_EXT)
        Next
    End If
    
    ' *********************
    ' ** Surface Texture **
    ' *********************
    '//Let's start counting at 1
    Count_Surface = 1
    TextureName = "company-name"
    Do While FileExist(App.Path & Texture_Path & Trim$(GameSetting.ThemePath) & Surface_Path & TextureName & GFX_EXT)
        '//If file exist, let's check the other count
        Count_Surface = Count_Surface + 1
        '//Add a proper name of the texture
        Select Case Count_Surface
            Case 1: TextureName = "company-name"
            Case 2: TextureName = "logo"
            Case 3: TextureName = "bg"
            Case Else: TextureName = Count_Surface
        End Select
    Loop
    '//If it reach at this point, it means that the file on the certain count is missing so we must remove it
    Count_Surface = Count_Surface - 1
    '//Make sure that we have texture to set
    If Count_Surface > 0 Then
        '//Set texture range
        ReDim Tex_Surface(1 To Count_Surface)
        '//Set the path range to GlobalTexture
        For i = 1 To Count_Surface
            '//Add a proper name of the texture
            Select Case i
                Case 1: TextureName = "company-name"
                Case 2: TextureName = "logo"
                Case 3: TextureName = "bg"
                Case Else: TextureName = i
            End Select
            Tex_Surface(i) = SetTexturePath(App.Path & Texture_Path & Trim$(GameSetting.ThemePath) & Surface_Path & TextureName & GFX_EXT)
        Next
    End If
    
    ' *****************
    ' ** Gui Texture **
    ' *****************
    '//Let's start counting at 1
    Count_Gui = 1
    TextureName = "menu-ui"
    Do While FileExist(App.Path & Texture_Path & Trim$(GameSetting.ThemePath) & Gui_Path & TextureName & GFX_EXT)
        '//If file exist, let's check the other count
        Count_Gui = Count_Gui + 1
        '//Add a proper name of the texture
        Select Case Count_Gui
            Case 1: TextureName = "menu-ui"
            Case 2: TextureName = "login-window"
            Case 3: TextureName = "register-window"
            Case 4: TextureName = "character-selection"
            Case 5: TextureName = "character-creation"
            Case 6: TextureName = "choice-box"
            Case 7: TextureName = "global-menu"
            Case 8: TextureName = "option-window"
            Case 9: TextureName = "chatbox"
            Case 10: TextureName = "game-ui"
            Case 11: TextureName = "inventory"
            Case 12: TextureName = "hud"
            Case 13: TextureName = "input-box"
            Case 14: TextureName = "move-replace"
            Case 15: TextureName = "trainer"
            Case 16: TextureName = "storage"
            Case 17: TextureName = "conv"
            Case 18: TextureName = "shop"
            Case 19: TextureName = "trade"
            Case 20: TextureName = "pokedex"
            Case 21: TextureName = "pokemon-summary"
            Case 22: TextureName = "relearn"
            Case 23: TextureName = "badge"
            Case 24: TextureName = "slot-machine"
            Case 25: TextureName = "rank"
            Case 26: TextureName = "bottom-login"
            Case Else: TextureName = Count_Gui
        End Select
    Loop
    '//If it reach at this point, it means that the file on the certain count is missing so we must remove it
    Count_Gui = Count_Gui - 1
    '//Make sure that we have texture to set
    If Count_Gui > 0 Then
        '//Set texture range
        ReDim Tex_Gui(1 To Count_Gui)
        '//Set the path range to GlobalTexture
        For i = 1 To Count_Gui
            '//Add a proper name of the texture
            Select Case i
                Case 1: TextureName = "menu-ui"
                Case 2: TextureName = "login-window"
                Case 3: TextureName = "register-window"
                Case 4: TextureName = "character-selection"
                Case 5: TextureName = "character-creation"
                Case 6: TextureName = "choice-box"
                Case 7: TextureName = "global-menu"
                Case 8: TextureName = "option-window"
                Case 9: TextureName = "chatbox"
                Case 10: TextureName = "game-ui"
                Case 11: TextureName = "inventory"
                Case 12: TextureName = "hud"
                Case 13: TextureName = "input-box"
                Case 14: TextureName = "move-replace"
                Case 15: TextureName = "trainer"
                Case 16: TextureName = "storage"
                Case 17: TextureName = "conv"
                Case 18: TextureName = "shop"
                Case 19: TextureName = "trade"
                Case 20: TextureName = "pokedex"
                Case 21: TextureName = "pokemon-summary"
                Case 22: TextureName = "relearn"
                Case 23: TextureName = "badge"
                Case 24: TextureName = "slot-machine"
                Case 25: TextureName = "rank"
                Case 26: TextureName = "bottom-login"
                Case Else: TextureName = i
            End Select
            Tex_Gui(i) = SetTexturePath(App.Path & Texture_Path & Trim$(GameSetting.ThemePath) & Gui_Path & TextureName & GFX_EXT)
        Next
    End If
    
    ' ***********************
    ' ** Character Texture **
    ' ***********************
    '//Let's start counting at 1
    Count_Character = 1
    Do While FileExist(App.Path & Character_Path & Count_Character & GFX_EXT)
        '//If file exist, let's check the other count
        Count_Character = Count_Character + 1
    Loop
    '//If it reach at this point, it means that the file on the certain count is missing so we must remove it
    Count_Character = Count_Character - 1
    '//Make sure that we have texture to set
    If Count_Character > 0 Then
        '//Set texture range
        ReDim Tex_Character(1 To Count_Character)
        '//Set the path range to GlobalTexture
        For i = 1 To Count_Character
            Tex_Character(i) = SetTexturePath(App.Path & Character_Path & i & GFX_EXT)
        Next
    End If
    
    ' ***********************
    ' ** PlayerSprite Texture **
    ' ***********************
    '//Let's start counting at 1
    Count_PlayerSprite_N = 1
    Do While FileExist(App.Path & PlayerSprite_Path & Count_PlayerSprite_N & GFX_EXT)
        '//If file exist, let's check the other count
        Count_PlayerSprite_N = Count_PlayerSprite_N + 1
    Loop
    '//If it reach at this point, it means that the file on the certain count is missing so we must remove it
    Count_PlayerSprite_N = Count_PlayerSprite_N - 1
    '//Make sure that we have texture to set
    If Count_PlayerSprite_N > 0 Then
        '//Set texture range
        ReDim Tex_PlayerSprite_N(1 To Count_PlayerSprite_N)
        '//Set the path range to GlobalTexture
        For i = 1 To Count_PlayerSprite_N
            Tex_PlayerSprite_N(i) = SetTexturePath(App.Path & PlayerSprite_Path & i & GFX_EXT)
        Next
    End If
    '//Let's start counting at 1
    Count_PlayerSprite_D = 1
    Do While FileExist(App.Path & PlayerSprite_Path & Count_PlayerSprite_D & "_d" & GFX_EXT)
        '//If file exist, let's check the other count
        Count_PlayerSprite_D = Count_PlayerSprite_D + 1
    Loop
    '//If it reach at this point, it means that the file on the certain count is missing so we must remove it
    Count_PlayerSprite_D = Count_PlayerSprite_D - 1
    '//Make sure that we have texture to set
    If Count_PlayerSprite_D > 0 Then
        '//Set texture range
        ReDim Tex_PlayerSprite_D(1 To Count_PlayerSprite_D)
        '//Set the path range to GlobalTexture
        For i = 1 To Count_PlayerSprite_D
            Tex_PlayerSprite_D(i) = SetTexturePath(App.Path & PlayerSprite_Path & i & "_d" & GFX_EXT)
        Next
    End If
    '//Let's start counting at 1
    Count_PlayerSprite_B = 1
    Do While FileExist(App.Path & PlayerSprite_Path & Count_PlayerSprite_B & "_b" & GFX_EXT)
        '//If file exist, let's check the other count
        Count_PlayerSprite_B = Count_PlayerSprite_B + 1
    Loop
    '//If it reach at this point, it means that the file on the certain count is missing so we must remove it
    Count_PlayerSprite_B = Count_PlayerSprite_B - 1
    '//Make sure that we have texture to set
    If Count_PlayerSprite_B > 0 Then
        '//Set texture range
        ReDim Tex_PlayerSprite_B(1 To Count_PlayerSprite_B)
        '//Set the path range to GlobalTexture
        For i = 1 To Count_PlayerSprite_B
            Tex_PlayerSprite_B(i) = SetTexturePath(App.Path & PlayerSprite_Path & i & "_b" & GFX_EXT)
        Next
    End If
    
    ' *********************
    ' ** Tileset Texture **
    ' *********************
    '//Let's start counting at 1
    Count_Tileset = 1
    Do While FileExist(App.Path & Tileset_Path & Count_Tileset & GFX_EXT)
        '//If file exist, let's check the other count
        Count_Tileset = Count_Tileset + 1
    Loop
    '//If it reach at this point, it means that the file on the certain count is missing so we must remove it
    Count_Tileset = Count_Tileset - 1
    '//Make sure that we have texture to set
    If Count_Tileset > 0 Then
        '//Set texture range
        ReDim Tex_Tileset(1 To Count_Tileset)
        '//Set the path range to GlobalTexture
        For i = 1 To Count_Tileset
            Tex_Tileset(i) = SetTexturePath(App.Path & Tileset_Path & i & GFX_EXT)
        Next
    End If
    
    ' *********************
    ' ** MapAnim Texture **
    ' *********************
    '//Let's start counting at 1
    Count_MapAnim = 1
    Do While FileExist(App.Path & MapAnim_Path & Count_MapAnim & GFX_EXT)
        '//If file exist, let's check the other count
        Count_MapAnim = Count_MapAnim + 1
    Loop
    '//If it reach at this point, it means that the file on the certain count is missing so we must remove it
    Count_MapAnim = Count_MapAnim - 1
    '//Make sure that we have texture to set
    If Count_MapAnim > 0 Then
        '//Set texture range
        ReDim Tex_MapAnim(1 To Count_MapAnim)
        '//Set the path range to GlobalTexture
        For i = 1 To Count_MapAnim
            Tex_MapAnim(i) = SetTexturePath(App.Path & MapAnim_Path & i & GFX_EXT)
        Next
    End If
    
    ' *********************
    ' ** Pokemon Texture **
    ' *********************
    '//Let's start counting at 1
    Count_Pokemon = 1
    Do While FileExist(App.Path & Pokemon_Path & Count_Pokemon & GFX_EXT)
        '//If file exist, let's check the other count
        Count_Pokemon = Count_Pokemon + 1
    Loop
    '//If it reach at this point, it means that the file on the certain count is missing so we must remove it
    Count_Pokemon = Count_Pokemon - 1
    '//Make sure that we have texture to set
    If Count_Pokemon > 0 Then
        '//Set texture range
        ReDim Tex_Pokemon(1 To Count_Pokemon)
        '//Set the path range to GlobalTexture
        For i = 1 To Count_Pokemon
            Tex_Pokemon(i) = SetTexturePath(App.Path & Pokemon_Path & i & GFX_EXT)
        Next
    End If
    
    ' *********************
    ' ** Shiny Pokemon Texture **
    ' *********************
    '//Let's start counting at 1
    Count_ShinyPokemon = 1
    Do While FileExist(App.Path & Pokemon_Path & Count_ShinyPokemon & "_s" & GFX_EXT)
        '//If file exist, let's check the other count
        Count_ShinyPokemon = Count_ShinyPokemon + 1
    Loop
    '//If it reach at this point, it means that the file on the certain count is missing so we must remove it
    Count_ShinyPokemon = Count_ShinyPokemon - 1
    '//Make sure that we have texture to set
    If Count_ShinyPokemon > 0 Then
        '//Set texture range
        ReDim Tex_ShinyPokemon(1 To Count_ShinyPokemon)
        '//Set the path range to GlobalTexture
        For i = 1 To Count_ShinyPokemon
            Tex_ShinyPokemon(i) = SetTexturePath(App.Path & Pokemon_Path & i & "_s" & GFX_EXT)
        Next
    End If
    
    ' ******************
    ' ** Item Texture **
    ' ******************
    '//Let's start counting at 1
    Count_Item = 1
    Do While FileExist(App.Path & Item_Path & Count_Item & GFX_EXT)
        '//If file exist, let's check the other count
        Count_Item = Count_Item + 1
    Loop
    '//If it reach at this point, it means that the file on the certain count is missing so we must remove it
    Count_Item = Count_Item - 1
    '//Make sure that we have texture to set
    If Count_Item > 0 Then
        '//Set texture range
        ReDim Tex_Item(1 To Count_Item)
        '//Set the path range to GlobalTexture
        For i = 1 To Count_Item
            Tex_Item(i) = SetTexturePath(App.Path & Item_Path & i & GFX_EXT)
        Next
    End If
    
    ' *****************
    ' ** Misc Texture **
    ' *****************
    '//Let's start counting at 1
    Count_Misc = 1
    TextureName = "chatbubble"
    Do While FileExist(App.Path & Misc_Path & TextureName & GFX_EXT)
        '//If file exist, let's check the other count
        Count_Misc = Count_Misc + 1
        '//Add a proper name of the texture
        Select Case Count_Misc
            Case 1: TextureName = "chatbubble"
            Case 2: TextureName = "bar"
            Case 3: TextureName = "move-selector"
            Case 4: TextureName = "pokeball"
            Case 5: TextureName = "language"
            Case 6: TextureName = "status"
            Case 7: TextureName = "poke-select"
            Case 8: TextureName = "shiny-summary"
            Case 9: TextureName = "category-types"
            Case Else: TextureName = Count_Misc
        End Select
    Loop
    '//If it reach at this point, it means that the file on the certain count is missing so we must remove it
    Count_Misc = Count_Misc - 1
    '//Make sure that we have texture to set
    If Count_Misc > 0 Then
        '//Set texture range
        ReDim Tex_Misc(1 To Count_Misc)
        '//Set the path range to GlobalTexture
        For i = 1 To Count_Misc
            '//Add a proper name of the texture
            Select Case i
                Case 1: TextureName = "chatbubble"
                Case 2: TextureName = "bar"
                Case 3: TextureName = "move-selector"
                Case 4: TextureName = "pokeball"
                Case 5: TextureName = "language"
                Case 6: TextureName = "status"
                Case 7: TextureName = "poke-select"
                Case 8: TextureName = "shiny-summary"
                 Case 9: TextureName = "category-types"
                Case Else: TextureName = i
            End Select
            Tex_Misc(i) = SetTexturePath(App.Path & Misc_Path & TextureName & GFX_EXT)
        Next
    End If
    
    ' *********************
    ' ** PokemonIcon Texture **
    ' *********************
    '//Let's start counting at 1
    Count_PokemonIcon = 1
    Do While FileExist(App.Path & PokemonIcon_Path & Count_PokemonIcon & "_icon" & GFX_EXT)
        '//If file exist, let's check the other count
        Count_PokemonIcon = Count_PokemonIcon + 1
    Loop
    '//If it reach at this point, it means that the file on the certain count is missing so we must remove it
    Count_PokemonIcon = Count_PokemonIcon - 1
    '//Make sure that we have texture to set
    If Count_PokemonIcon > 0 Then
        '//Set texture range
        ReDim Tex_PokemonIcon(1 To Count_PokemonIcon)
        '//Set the path range to GlobalTexture
        For i = 1 To Count_PokemonIcon
            Tex_PokemonIcon(i) = SetTexturePath(App.Path & PokemonIcon_Path & i & "_icon" & GFX_EXT)
        Next
    End If
    
    ' *********************
    ' ** Animation Texture **
    ' *********************
    '//Let's start counting at 1
    Count_Animation = 1
    Do While FileExist(App.Path & Animation_Path & Count_Animation & GFX_EXT)
        '//If file exist, let's check the other count
        Count_Animation = Count_Animation + 1
    Loop
    '//If it reach at this point, it means that the file on the certain count is missing so we must remove it
    Count_Animation = Count_Animation - 1
    '//Make sure that we have texture to set
    If Count_Animation > 0 Then
        '//Set texture range
        ReDim Tex_Animation(1 To Count_Animation)
        '//Set the path range to GlobalTexture
        For i = 1 To Count_Animation
            Tex_Animation(i) = SetTexturePath(App.Path & Animation_Path & i & GFX_EXT)
        Next
    End If
    
    ' *********************
    ' ** Weather Texture **
    ' *********************
    '//Let's start counting at 1
    Count_Weather = 1
    Do While FileExist(App.Path & Weather_Path & Count_Weather & GFX_EXT)
        '//If file exist, let's check the other count
        Count_Weather = Count_Weather + 1
    Loop
    '//If it reach at this point, it means that the file on the certain count is missing so we must remove it
    Count_Weather = Count_Weather - 1
    '//Make sure that we have texture to set
    If Count_Weather > 0 Then
        '//Set texture range
        ReDim Tex_Weather(1 To Count_Weather)
        '//Set the path range to GlobalTexture
        For i = 1 To Count_Weather
            Tex_Weather(i) = SetTexturePath(App.Path & Weather_Path & i & GFX_EXT)
        Next
    End If
    
    ' *********************
    ' ** PokemonPortrait Texture **
    ' *********************
    '//Let's start counting at 1
    Count_PokemonPortrait = 1
    Do While FileExist(App.Path & PokemonPortrait_Path & Count_PokemonPortrait & GFX_EXT)
        '//If file exist, let's check the other count
        Count_PokemonPortrait = Count_PokemonPortrait + 1
    Loop
    '//If it reach at this point, it means that the file on the certain count is missing so we must remove it
    Count_PokemonPortrait = Count_PokemonPortrait - 1
    '//Make sure that we have texture to set
    If Count_PokemonPortrait > 0 Then
        '//Set texture range
        ReDim Tex_PokemonPortrait(1 To Count_PokemonPortrait)
        '//Set the path range to GlobalTexture
        For i = 1 To Count_PokemonPortrait
            Tex_PokemonPortrait(i) = SetTexturePath(App.Path & PokemonPortrait_Path & i & GFX_EXT)
        Next
    End If
    
    ' *********************
    ' ** ShinyPokemonPortrait Texture **
    ' *********************
    '//Let's start counting at 1
    Count_ShinyPokemonPortrait = 1
    Do While FileExist(App.Path & PokemonPortrait_Path & Count_ShinyPokemonPortrait & "s" & GFX_EXT)
        '//If file exist, let's check the other count
        Count_ShinyPokemonPortrait = Count_ShinyPokemonPortrait + 1
    Loop
    '//If it reach at this point, it means that the file on the certain count is missing so we must remove it
    Count_ShinyPokemonPortrait = Count_ShinyPokemonPortrait - 1
    '//Make sure that we have texture to set
    If Count_ShinyPokemonPortrait > 0 Then
        '//Set texture range
        ReDim Tex_ShinyPokemonPortrait(1 To Count_ShinyPokemonPortrait)
        '//Set the path range to GlobalTexture
        For i = 1 To Count_ShinyPokemonPortrait
            Tex_ShinyPokemonPortrait(i) = SetTexturePath(App.Path & PokemonPortrait_Path & i & "s" & GFX_EXT)
        Next
    End If
    
    ' *********************
    ' ** PokemonTypes *****
    ' *********************
    '//Let's start counting at 1
    Count_PokemonTypes = 1
    Do While FileExist(App.Path & PokemonTypes_Path & Count_PokemonTypes & GFX_EXT)
        '//If file exist, let's check the other count
        Count_PokemonTypes = Count_PokemonTypes + 1
    Loop
    '//If it reach at this point, it means that the file on the certain count is missing so we must remove it
    Count_PokemonTypes = Count_PokemonTypes - 1
    '//Make sure that we have texture to set
    If Count_PokemonTypes > 0 Then
        '//Set texture range
        ReDim Tex_PokemonTypes(1 To Count_PokemonTypes)
        '//Set the path range to GlobalTexture
        For i = 1 To Count_PokemonTypes
            Tex_PokemonTypes(i) = SetTexturePath(App.Path & PokemonTypes_Path & i & GFX_EXT)
        Next i
    End If
    
    '//Symbol
    '//Let's start counting at 1
    Count_PokemonTypesSymbol = 1
    Do While FileExist(App.Path & PokemonTypesSymbol_Path & Count_PokemonTypesSymbol & GFX_EXT)
        '//If file exist, let's check the other count
        Count_PokemonTypesSymbol = Count_PokemonTypesSymbol + 1
    Loop
    '//If it reach at this point, it means that the file on the certain count is missing so we must remove it
    Count_PokemonTypesSymbol = Count_PokemonTypesSymbol - 1
    '//Make sure that we have texture to set
    If Count_PokemonTypesSymbol > 0 Then
        '//Set texture range
        ReDim Tex_PokemonTypesSymbol(0 To Count_PokemonTypesSymbol)
        '//Set the path range to GlobalTexture
        For i = 1 To Count_PokemonTypesSymbol
            Tex_PokemonTypesSymbol(i) = SetTexturePath(App.Path & PokemonTypesSymbol_Path & i & GFX_EXT)
        Next i
    End If
End Sub

'//Setting the count or index of the texture on GlobalTexture
Private Function SetTexturePath(ByVal Path As String) As Long
    GlobalTextureCount = GlobalTextureCount + 1                             '//Add texture count
    ReDim Preserve GlobalTexture(0 To GlobalTextureCount) As TextureRec     '//Set texture range
    GlobalTexture(GlobalTextureCount).Path = Path                           '//Set texture path
    SetTexturePath = GlobalTextureCount
    GlobalTexture(GlobalTextureCount).Loaded = False
End Function

Private Sub LoadTexture(ByVal TextureNum As Long)
Dim Tex_Info As D3DXIMAGE_INFO_A
Dim Path As String

    '//Set Texture path
    Path = GlobalTexture(TextureNum).Path
    
    '//Load the texture
    Set GlobalTexture(TextureNum).Texture = D3DX.CreateTextureFromFileEx(D3DDevice, Path, D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, RGB(255, 0, 255), Tex_Info, ByVal 0)
    '//Update the texture info size
    GlobalTexture(TextureNum).Height = Tex_Info.Height
    GlobalTexture(TextureNum).Width = Tex_Info.Width
    
    '//Reset unload timer
    GlobalTexture(TextureNum).UnloadTimer = GetTickCount
    GlobalTexture(TextureNum).Loaded = True
End Sub

Private Sub UnloadTextures()
Dim Count As Long, i As Long
Dim TexturePath As String

    If GlobalTextureCount <= 0 Then Exit Sub
    Count = UBound(GlobalTexture)
    If Count <= 0 Then Exit Sub
    
    For i = 1 To Count
        ' Make sure that the texture that we are unloading is loaded
        If GlobalTexture(i).Loaded Then
            If GetTickCount > GlobalTexture(i).UnloadTimer + 15000 Then
                '//First let's store the texture path so that we can place it again
                TexturePath = GlobalTexture(i).Path
                '//Unload and clear everything
                Set GlobalTexture(i).Texture = Nothing
                Call ZeroMemory(ByVal VarPtr(GlobalTexture(i)), LenB(GlobalTexture(i)))
                GlobalTexture(i).UnloadTimer = 0
                GlobalTexture(i).Loaded = False
                '//Let's place the texture path
                GlobalTexture(i).Path = TexturePath
            End If
        End If
    Next
End Sub

'//This set the texture of what the DrawPrimitiveUP will draw
Private Sub SetTexture(ByVal Texture As Long)
    '//Make sure that we haven't draw it yet
    If Texture <> CurrentTexture Then
        If Texture > UBound(GlobalTexture) Then Texture = UBound(GlobalTexture)
        If Texture < 0 Then Texture = 0

        '//If the texture was unloaded, then reload it
        If Not Texture = 0 Then
            If Not GlobalTexture(Texture).Loaded Then
                LoadTexture Texture
            End If
        End If
        
        '//Set the texture as the next texture to be draw on DrawPrimitiveUP
        D3DDevice.SetTexture 0, GlobalTexture(Texture).Texture
        '//Let's make sure that it won't redraw again
        CurrentTexture = Texture
    End If
End Sub

Public Function GetPicWidth(ByVal TextureNum As Long) As Long
    If TextureNum <= 0 Then Exit Function                               '//Make sure that we have Texture
    If Not GlobalTexture(TextureNum).Loaded Then SetTexture TextureNum  '//Load the texture if it's not loaded
    GetPicWidth = GlobalTexture(TextureNum).Width                       '//Send Texture Width
End Function

Public Function GetPicHeight(ByVal TextureNum As Long) As Long
    If TextureNum <= 0 Then Exit Function                               '//Make sure that we have Texture
    If Not GlobalTexture(TextureNum).Loaded Then SetTexture TextureNum  '//Load the texture if it's not loaded
    GetPicHeight = GlobalTexture(TextureNum).Height                     '//Send Texture Height
End Function

'//This make rendering of a texture more easier than doing it manually
Public Sub RenderTexture(ByVal Texture As Long, ByVal X As Long, ByVal Y As Long, ByVal pX As Long, ByVal pY As Long, ByVal sW As Long, ByVal sH As Long, ByVal rW As Long, ByVal rH As Long, Optional ByVal Colour As Long = -1, Optional ByVal Degrees As Single = 0)
Dim Box(0 To 3) As TLVERTEX
Dim Width As Long, Height As Long
Dim Des As Single

'//This is use for rotation
Dim i As Long
Dim RadAngle As Single
Dim CenterX As Single, CenterY As Single
Dim SinRad As Single, CosRad As Single
Dim NewX As Single, NewY As Single

    '//set the texture that we are using
    SetTexture Texture
    
    '//get the texture size
    Width = GetPicWidth(Texture)
    Height = GetPicHeight(Texture)
    
    '//exit out if we need to
    If Texture <= 0 Or Width <= 0 Or Height <= 0 Then Exit Sub
    
    pX = pX '+ 0.5
    pY = pY '+ 0.5
    Des = 0.000003
    '//Create the vertex of a box
    Box(0) = CreateTLVertex(X, Y, 0, 1, Colour, (pX / Width) + Des, (pY / Height) + Des)
    Box(1) = CreateTLVertex(X + sW, Y, 0, 1, Colour, ((pX + rW) / Width) + Des, (pY / Height) + Des)
    Box(2) = CreateTLVertex(X, Y + sH, 0, 1, Colour, (pX / Width) + Des, ((pY + rH) / Height) + Des)
    Box(3) = CreateTLVertex(X + sW, Y + sH, 0, 1, Colour, ((pX + rW) / Width) + Des, ((pY + rH) / Height) + Des)

    '//Check if a rotation is required
    If Degrees <> 0 And Degrees <> 360 Then
        '//Converts the angle to rotate by into radians
        RadAngle = Degrees * DegreeToRadian

        '//Set the CenterX and CenterY values
        CenterX = X + (sW * 0.5)
        CenterY = Y + (sH * 0.5)

        '//Pre-calculate the cosine and sine of the radiant
        SinRad = Sin(RadAngle)
        CosRad = Cos(RadAngle)

        '//Loops through the passed vertex buffer
        For i = 0 To 3
            '//Calculates the new X and Y co-ordinates of the vertices for the given angle around the center co-ordinates
            NewX = CenterX + (Box(i).X - CenterX) * CosRad - (Box(i).Y - CenterY) * SinRad
            NewY = CenterY + (Box(i).Y - CenterY) * CosRad + (Box(i).X - CenterX) * SinRad

            '//Applies the new co-ordinates to the buffer
            Box(i).X = NewX
            Box(i).Y = NewY
        Next
    End If
    
    '//Draw the set texture on screen
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Box(0), FVF_Size
    GlobalTexture(Texture).UnloadTimer = GetTickCount
End Sub

Public Sub RenderTextureByRects(ByVal TextureRec As Long, sRect As RECT, dRect As RECT)
    '//Render Texture using RECT's
    RenderTexture TextureRec, dRect.Left, dRect.top, sRect.Left, sRect.top _
                            , dRect.Right - dRect.Left, dRect.bottom - dRect.top _
                            , sRect.Right - sRect.Left, sRect.bottom - sRect.top _
                            , D3DColorRGBA(255, 255, 255, 255)
End Sub

'//This is just a simple wrapper function that makes filling the structures much much easier...
Private Function CreateTLVertex(X As Long, Y As Long, z As Single, rhw As Single, Color As Long, tu As Single, tv As Single) As TLVERTEX
    '//NB: whilst you can pass floating point values for the coordinates (single)
    '       there is little point - Direct3D will just approximate the coordinate by rounding
    '       which may well produce unwanted results....
    CreateTLVertex.X = X
    CreateTLVertex.Y = Y
    CreateTLVertex.z = z
    CreateTLVertex.rhw = rhw
    CreateTLVertex.Color = Color
    CreateTLVertex.tu = tu
    CreateTLVertex.tv = tv
End Function

Private Sub UpdateViewPort()
    If ViewPortInit Then Exit Sub
    
    '//Update ViewPort
    ScreenX = ((Screen_Width / TILE_X) + 1) * TILE_X
    ScreenY = ((Screen_Height / TILE_Y) + 1) * TILE_Y
    StartXValue = (((Screen_Width / TILE_X) + 1) / 2) - 1
    StartYValue = (((Screen_Height / TILE_Y) + 1) / 2) + 1
    EndXValue = ((Screen_Width / TILE_X) + 1)
    EndYValue = ((Screen_Height / TILE_Y) + 1)
    GlobalMapX = (Screen_Width / TILE_X) - 1
    GlobalMapY = (Screen_Height / TILE_Y) - 1
    ViewPortInit = True
End Sub

Private Sub UpdateCamera()
Dim offsetX As Long, offsetY As Long
Dim StartX As Long, StartY As Long
Dim EndX As Long, EndY As Long
Dim pX As Long, pY As Long
Dim pOffsetX As Long, pOffsetY As Long

    UpdateViewPort

    If MyIndex <= 0 Or MyIndex > MAX_PLAYER Then Exit Sub
    If GettingMap Then Exit Sub

    '//If player pokemon is available then switch camera
    If PlayerPokemon(MyIndex).Num > 0 Then
        pX = PlayerPokemon(MyIndex).X
        pY = PlayerPokemon(MyIndex).Y
        pOffsetX = PlayerPokemon(MyIndex).xOffset
        pOffsetY = PlayerPokemon(MyIndex).yOffset
    Else
        pX = Player(MyIndex).X
        pY = Player(MyIndex).Y
        pOffsetX = Player(MyIndex).xOffset
        pOffsetY = Player(MyIndex).yOffset
    End If
    
    If GlobalMapX <= Map.MaxX Then
        offsetX = pOffsetX + TILE_X
        offsetY = pOffsetY + TILE_Y
        StartX = pX - StartXValue - 1
        StartY = pY - StartYValue + 1
        
        If StartX < 0 Then
            offsetX = 0
            
            If StartX = -1 Then
                If pOffsetX > 0 Then
                    offsetX = pOffsetX
                End If
            End If
        
            StartX = 0
        End If
        If StartY < 0 Then
            offsetY = 0
        
            If StartY = -1 Then
                If pOffsetY > 0 Then
                    offsetY = pOffsetY
                End If
            End If
        
            StartY = 0
        End If
        
        EndX = StartX + EndXValue
        EndY = StartY + EndYValue
        
        If EndX > Map.MaxX Then
            offsetX = TILE_X
            
            If EndX = Map.MaxX + 1 Then
                If pOffsetX < 0 Then
                    offsetX = pOffsetX + TILE_X
                End If
            End If
            
            EndX = Map.MaxX
            StartX = EndX - GlobalMapX - 1
        End If
        If EndY > Map.MaxY Then
            offsetY = TILE_Y
        
            If EndY = Map.MaxY + 1 Then
                If pOffsetY < 0 Then
                    offsetY = pOffsetY + TILE_Y
                End If
            End If
        
            EndY = Map.MaxY
            StartY = EndY - GlobalMapY - 1
        End If
        
        '//Update ViewPort
        With TileView
            .top = StartY
            .bottom = EndY
            .Left = StartX
            .Right = EndX
        End With
        With Camera
            .top = offsetY
            .bottom = .top + ScreenY
            .Left = offsetX
            .Right = .Left + ScreenX
        End With
    Else
        '//Update ViewPort
        With TileView
            .top = pY - StartYValue + 1
            .bottom = .top + EndYValue + 1
            .Left = pX - StartXValue - 1
            .Right = .Left + EndXValue + 1
        End With
        With Camera
            .top = pOffsetY + TILE_Y
            .bottom = .top + ScreenY
            .Left = pOffsetX + TILE_X
            .Right = .Left + ScreenX
        End With
    End If
End Sub

'//This sub render all the stuff on the game screen
Public Sub Render_Screen()
    If ReInit Then Exit Sub
    
    ' Make sure we've got control of the form
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then
        If D3DDevice.TestCooperativeLevel = D3DERR_DEVICELOST Then Exit Sub
        D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
        D3DDevice.Reset D3DWindow
        InitRenderState
    End If
    
    '//Make sure to unload all unrequired texture
    UnloadTextures
    
    ' *****************
    ' ** Game Screen **
    ' *****************
    '//We need to clear the render device before we can draw anything
    '//This must always happen before you start rendering stuff...
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    '//Next we would render everything.
    D3DDevice.BeginScene
    '//All rendering calls go after these 'BeginScene' Line
    
    '//Check for Game State
    Select Case GameState
        Case GameStateEnum.InMenu: Render_Menu
        Case GameStateEnum.InGame: Render_Game
    End Select
    
    RenderText Font_Default, PingToDraw, 10, 10, White
    
    '//Graphics that are being rendered below here are Global Type
    '//Which means that they are not based on what current game state does the app have
    
    '//Load Screen
    DrawLoad
    
    '//Alert
    DrawOption
    DrawChoiceBox
    DrawInputBox
    DrawAlertWindow
    DrawGlobalMenu
    
    '//Fade
    If Fade Then
        RenderTexture Tex_System(gSystemEnum.UserInterface), 0, 0, 0, 8, Screen_Width, Screen_Height, 1, 1, D3DColorARGB(FadeAlpha, 0, 0, 0)
    End If
    
    '//FPS
    If GameSetting.ShowFPS = YES Then
        RenderText Font_Default, "FPS: " & GameFps, 10, 80, White
    End If
    
    '//Cursor
    If CursorTimer <= GetTickCount Then
        CanShowCursor = False
        CursorTimer = 0
    End If

    If CanShowCursor Then
        If Not InvItemDescShow Then ' and Not ShopItemDescShow and not StorageItemDescShow Then
            If MouseIcon = 1 Then
                RenderTexture Tex_System(gSystemEnum.CursorIcon), CursorX - 3, CursorY, 32 * MouseIcon, 0, 32, 32, 32, 32
            Else
                RenderTexture Tex_System(gSystemEnum.CursorIcon), CursorX, CursorY, 32 * MouseIcon, 0, 32, 32, 32, 32
            End If
            
            If IsLoading Or Fade Or GettingMap Then
                If MouseIcon = 0 Then
                    RenderTexture Tex_System(gSystemEnum.CursorLoad), CursorX + 10, CursorY + 14, 15 * CursorLoadAnim, 0, 15, 15, 15, 15
                End If
            End If
        End If
    End If
    
    '//All rendering calls go before these 'EndScene' Line
    D3DDevice.EndScene
    '//Update the frame to the screen...
    '//This is the same as the Primary.Flip method as used in DirectX 7
    '//These values below should work for almost all cases...
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
    
    '//Editor GDI
    DrawGDI
End Sub

'//This render all graphics in-game
Private Sub Render_Game()
    Dim X As Long, Y As Long
    Dim i As Long
    Dim Addy As Long

    '//Optional: If GettingMap, Show a getting map screen, Other: Just pure black
    'If GettingMap Then Exit Sub '//PURE BLACK
    If Not GettingMap Then
        '//Updating ViewPort
        UpdateCamera

        '//reset
        Addy = 0

        '//Lower Tiles
        For X = TileView.Left To TileView.Right
            For Y = TileView.top To TileView.bottom
                For i = MapLayer.Ground To MapLayer.Mask2
                    DrawMapTile i, X, Y
                Next
                '//Check Distance
                If PlayerPokemon(MyIndex).Num > 0 Then
                    If X >= Player(MyIndex).X - MAX_DISTANCE And X <= Player(MyIndex).X + MAX_DISTANCE Then
                        If Y >= Player(MyIndex).Y - MAX_DISTANCE And Y <= Player(MyIndex).Y + MAX_DISTANCE Then
                            RenderTexture Tex_System(gSystemEnum.UserInterface), ConvertMapX(X * TILE_X), ConvertMapY(Y * TILE_Y), 0, 8, TILE_X, TILE_Y, 1, 1, D3DColorARGB(50, 10, 120, 5)
                        End If
                    End If
                End If
            Next
        Next

        '//Lower Animation
        If Count_Animation > 0 Then
            For i = 1 To 255
                If AnimInstance(i).Used(0) Then
                    DrawAnimation i, 0    ' 0 = Lower Layer
                End If
            Next
        End If

        '//Sprite/Objects
        For Y = 0 To Map.MaxY
            If Player_HighIndex > 0 Then
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If Player(i).Map = Player(MyIndex).Map Then
                            If Player(i).Y = Y Then
                                DrawPlayer i
                            End If
                            If PlayerPokemon(i).Init = YES Then
                                If Player(i).StealthMode = NO Then
                                    DrawPokeball PlayerPokemon(i).BallX, PlayerPokemon(i).BallY, PlayerPokemon(i).Frame, PlayerPokemon(i).BallUsed
                                End If
                            Else
                                If PlayerPokemon(i).Num > 0 Then
                                    If PlayerPokemon(i).Y = Y Then
                                        DrawPlayerPokemon i
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            End If
            If Npc_HighIndex > 0 Then
                For i = 1 To Npc_HighIndex
                    If MapNpc(i).Num > 0 Then
                        If MapNpc(i).Y = Y Then
                            DrawNpc i
                        End If
                        If MapNpcPokemon(i).Init = YES Then
                            DrawPokeball MapNpcPokemon(i).BallX, MapNpcPokemon(i).BallY, MapNpcPokemon(i).Frame, 1
                        Else
                            If MapNpcPokemon(i).Num > 0 Then
                                If MapNpcPokemon(i).Y = Y Then
                                    DrawMapNpcPokemon i
                                End If
                            End If
                        End If
                    End If
                Next
            End If
            If Pokemon_HighIndex > 0 Then
                For i = 1 To Pokemon_HighIndex
                    If CatchBall(i).InUsed Then
                        '//drawpokeball
                        DrawPokeball CatchBall(i).X, CatchBall(i).Y, CatchBall(i).Frame, CatchBall(i).Pic
                    Else
                        If MapPokemon(i).Num > 0 Then
                            If MapPokemon(i).Y = Y Then
                                DrawPokemon i
                            End If
                        End If
                    End If
                Next
            End If
        Next

        '//Upper Tiles
        For X = TileView.Left To TileView.Right
            For Y = TileView.top To TileView.bottom
                For i = MapLayer.Fringe To MapLayer.Fringe2
                    DrawMapTile i, X, Y
                Next
            Next
        Next

        '//Upper Animation
        If Count_Animation > 0 Then
            For i = 1 To 255
                If AnimInstance(i).Used(1) Then
                    DrawAnimation i, 1    ' 1 = Upper Layer
                End If
            Next
        End If

        '//Upper Tiles
        '//Night Lights
        If Map.Sheltered = 0 Then
            '//Day And Night
            If Editor <> EDITOR_MAP Then RenderTexture Tex_System(gSystemEnum.UserInterface), 0, 0, 0, 8, Screen_Width, Screen_Height, 1, 1, DayAndNightARGB

            If ShowLights Or (Editor = EDITOR_MAP And CurLayer = MapLayer.Lights) Then
                If Editor = EDITOR_MAP Then
                    If CurLayer = MapLayer.Lights Then
                        RenderTexture Tex_System(gSystemEnum.UserInterface), 0, 0, 0, 8, Screen_Width, Screen_Height, 1, 1, D3DColorARGB(100, 0, 0, 0)
                        LightAlpha = 255
                    End If
                End If

                For X = TileView.Left To TileView.Right
                    For Y = TileView.top To TileView.bottom
                        DrawMapTile MapLayer.Lights, X, Y, LightAlpha
                    Next
                Next
            End If

            '//Weather
            DrawWeather
        End If

        '//Bar
        DrawVitalBar

        '//Name
        If GameSetting.ShowName = YES Then
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If Player(i).Map = Player(MyIndex).Map Then
                        DrawPlayerName i
                        If PlayerPokemon(i).Num > 0 And PlayerPokemon(i).Init = NO Then
                            DrawPlayerPokemonName i
                        End If
                    End If
                End If
            Next
            For i = 1 To Npc_HighIndex
                If MapNpc(i).Num > 0 Then
                    DrawNpcName i
                    If MapNpcPokemon(i).Num > 0 Then
                        DrawNpcPokemonName i
                    End If
                End If
            Next
            For i = 1 To Pokemon_HighIndex
                If Not CatchBall(i).InUsed Then
                    If MapPokemon(i).Num > 0 Then
                        If MapPokemon(i).Map = Player(MyIndex).Map Then
                            DrawPokemonName i
                        End If
                    End If
                End If
            Next
        End If

        '//Chatbubble
        For i = 1 To 255
            If chatBubble(i).active Then
                DrawChatBubble i
            End If

            DrawActionMsg i
        Next

        '//Editor
        If Editor = EDITOR_MAP Then
            DrawMapAttributes
        End If

        '//Loc
        If ShowLoc Then
            If GameSetting.ShowFPS = YES And GameSetting.ShowPing = YES Then
                Addy = Addy + 110
            ElseIf GameSetting.ShowFPS = YES Or GameSetting.ShowPing = YES Then
                Addy = Addy + 95
            Else
                Addy = Addy + 65
            End If
            'If GameSetting.ShowPing = YES Then AddY = AddY + 65

            RenderText Font_Default, "[Player Position]", 10, Addy + 10, White
            RenderText Font_Default, "Map#: " & Player(MyIndex).Map, 10, Addy + 25, White
            RenderText Font_Default, "X: " & Player(MyIndex).X & " Y: " & Player(MyIndex).Y, 10, Addy + 40, White
            RenderText Font_Default, "[Cursor Position]", 10, Addy + 55, White
            RenderText Font_Default, "Cursor X: " & CursorX & " Cursor Y: " & CursorY, 10, Addy + 70, White
            RenderText Font_Default, "Tile X: " & curTileX & " Tile Y: " & curTileY, 10, Addy + 85, White
        End If

        '//Move Selector
        DrawMoveSelector

        '//Hud
        DrawHud

        '//Buttons
        For i = ButtonEnum.Game_Pokedex To ButtonEnum.Game_Evolve
            If CanShowButton(i) Then
                Select Case i
                Case ButtonEnum.Game_Pokedex
                    If GUI(GuiEnum.GUI_POKEDEX).Visible Then
                        Button(i).State = ButtonState.StateClick
                    End If
                Case ButtonEnum.Game_Bag
                    If GUI(GuiEnum.GUI_INVENTORY).Visible Then
                        Button(i).State = ButtonState.StateClick
                    End If
                Case ButtonEnum.Game_Card
                    If GUI(GuiEnum.GUI_TRAINER).Visible Then
                        Button(i).State = ButtonState.StateClick
                    End If
                Case ButtonEnum.Game_Task
                    'If GUI(GuiEnum.GUI_RANK).Visible Then
                    '    Button(i).State = ButtonState.StateClick
                    'End If
                Case ButtonEnum.Game_Rank
                    If GUI(GuiEnum.GUI_RANK).Visible Then
                        Button(i).State = ButtonState.StateClick
                    End If
                
                Case ButtonEnum.Game_Menu
                    If GUI(GuiEnum.GUI_GLOBALMENU).Visible Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End Select
                RenderTexture Tex_Gui(GameUi_Texture), Button(i).X, Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
            End If
        Next

        '//zOrdering of gui
        If Not IsLoading Then
            If GuiVisibleCount > 0 Then
                For i = 1 To GuiVisibleCount
                    If CanShowGui(GuiZOrder(i)) Then
                        '//The first one will be rendered first
                        Select Case GuiZOrder(i)
                        Case GuiEnum.GUI_CHATBOX: DrawChatbox
                        Case GuiEnum.GUI_INVENTORY: DrawInventory
                        Case GuiEnum.GUI_MOVEREPLACE: DrawMoveReplace
                        Case GuiEnum.GUI_TRAINER: DrawTrainer
                        Case GuiEnum.GUI_INVSTORAGE: DrawInvStorage
                        Case GuiEnum.GUI_POKEMONSTORAGE: DrawPokemonStorage
                        Case GuiEnum.GUI_SHOP: DrawShop
                        Case GuiEnum.GUI_TRADE: DrawTrade
                        Case GuiEnum.GUI_POKEDEX: DrawPokedex
                        Case GuiEnum.GUI_POKEMONSUMMARY: DrawPokemonSummary
                        Case GuiEnum.GUI_RELEARN: DrawRelearn
                        Case GuiEnum.GUI_BADGE: DrawBadge
                            'Case GuiEnum.GUI_SLOTMACHINE: DrawSlotMachine
                        Case GuiEnum.GUI_RANK: DrawRank
                        End Select
                    End If
                Next
            End If
        End If

        '//Convo
        DrawConvo

        '//Icon
        DrawDragIcon

        '//SelMenu
        DrawSelMenu
        
        '//Events
        DrawEvents
    Else
        RenderTexture Tex_System(gSystemEnum.UserInterface), 0, 0, 0, 8, Screen_Width, Screen_Height, 1, 1, D3DColorARGB(255, 10, 10, 10)
        RenderText Font_Default, "Loading...", Screen_Width - GetTextWidth(Font_Default, "Loading...") - 35, Screen_Height - 25, White
        RenderTexture Tex_System(gSystemEnum.CursorLoad), Screen_Width - 25, Screen_Height - 25, 15 * CursorLoadAnim, 0, 15, 15, 15, 15
    End If

    '//reset
    Addy = 0
    If GameSetting.ShowFPS = YES Then Addy = Addy + 15

    '//Ping
    If GameSetting.ShowPing = YES Then
        RenderText Font_Default, "Ping: ", 10, Addy + 85, White
        If Ping <= 255 Then
            RenderText Font_Default, Ping & "ms", GetTextWidth(Font_Default, "Ping: ") + 10 + 3, Addy + 85, BrightGreen
        ElseIf Ping >= 256 And Ping <= 600 Then
            RenderText Font_Default, Ping & "ms", GetTextWidth(Font_Default, "Ping: ") + 10 + 3, Addy + 85, Yellow
        Else
            RenderText Font_Default, Ping & "ms", GetTextWidth(Font_Default, "Ping: ") + 10 + 3, Addy + 85, Red
        End If
    End If

    '//Usage of Item
    If InvUseSlot > 0 Then
        If PlayerInv(InvUseSlot).Num > 0 Then
            i = Item(PlayerInv(InvUseSlot).Num).Sprite

            '//Draw Icon
            If i > 0 And i <= Count_Item Then
                RenderTexture Tex_Item(i), CursorX - (GetPicWidth(Tex_Item(i)) / 2), CursorY - (GetPicHeight(Tex_Item(i)) / 2), 0, 0, GetPicWidth(Tex_Item(i)), GetPicHeight(Tex_Item(i)), GetPicWidth(Tex_Item(i)), GetPicHeight(Tex_Item(i))
            End If
        End If
    End If

    DrawInvItemDesc
    DrawShopItemDesc
    DrawStorageItemDesc
End Sub

Private Sub DrawEvents()
    Dim Addy As Integer
    If ExpMultiply > 0 Then
        Addy = 2
        RenderTexture Tex_Gui(Hud), 180, Addy, 59, 241, 165, 20, 165, 1
        RenderText Font_Default, "Evento Exp Mult: " & ExpMultiply & "x", 190, Addy, Yellow
        Addy = 21
        
        RenderTexture Tex_Gui(Hud), 180, Addy, 59, 241, 165, 20, 165, 1, D3DColorARGB(150, 255, 255, 255)
        RenderText Font_Default, "Timer: " & SecondsToHMS(ExpSecs), 190, Addy, White
    End If
End Sub

'//This render all graphics of Menu
Private Sub Render_Menu()
Dim i As Byte
Dim X As Long
Dim footer As Boolean

    '//Select state
    '//Draw all required setup on the set state
    Select Case MenuState
        Case MenuStateEnum.StateCompanyScreen:  DrawCompanyScreen
        Case MenuStateEnum.StateTitleScreen:    DrawTitleScreen
        Case MenuStateEnum.StateNormal
            DrawBackground
            
            If CreditVisible = True Then
                footer = True
            End If

            '//zOrdering of gui
            If Not IsLoading Then
                If GuiVisibleCount > 0 Then
                    For i = 1 To GuiVisibleCount
                        If CanShowGui(GuiZOrder(i)) Then
                            '//The first one will be rendered first
                            Select Case GuiZOrder(i)
                                Case GuiEnum.GUI_LOGIN:
                                    DrawLogin
                                    footer = True
                                Case GuiEnum.GUI_REGISTER:
                                    DrawRegister
                                    footer = True
                                Case GuiEnum.GUI_CHARACTERSELECT:
                                    DrawCharacterSelect
                                    footer = True
                                Case GuiEnum.GUI_CHARACTERCREATE:
                                    DrawCharacterCreate
                                    footer = True
                            End Select
                        End If
                    Next
                End If
            End If
            
            If footer = True Then
                DrawFooter
            End If
            
    End Select
End Sub

'//This render all graphics of loading screen
Private Sub DrawLoad()
Dim LowBound As Long, UpBound As Long
Dim ArrayText() As String
Dim MaxWidth As Long
Dim X As Long, Y As Long
Dim i As Integer
Dim yOffset As Long
Dim PaddingSize As Long

    '//Make sure that loading screen is visible
    If Not IsLoading Then Exit Sub
    
    '//Editable
    PaddingSize = 20

    '//Make sure that loading text have something to draw
    If Len(LoadText) > 0 Then
        '//Wrap the text
        WordWrap_Array Font_Default, LoadText, LOAD_STRING_LENGTH, ArrayText
        
        '//we need these often
        LowBound = LBound(ArrayText)
        UpBound = UBound(ArrayText)
        
        '//Check if it wrap
        If UpBound > LowBound Then
            '//Get the longest width
            MaxWidth = GetTextWidth(Font_Default, ArrayText(LowBound))
            For i = LowBound + 1 To UpBound
                If MaxWidth < GetTextWidth(Font_Default, ArrayText(i)) Then
                    MaxWidth = GetTextWidth(Font_Default, ArrayText(i))
                End If
            Next
    
            '//Draw the hud of the text
            X = (Screen_Width / 2) - (MaxWidth / 2)
            Y = (Screen_Height / 2) - ((16 * UpBound) / 2)
            RenderTexture Tex_System(gSystemEnum.UserInterface), X - PaddingSize, Y - PaddingSize, 0, 8, MaxWidth + (PaddingSize * 2), (16 * UpBound) + (PaddingSize * 2), 1, 1, D3DColorARGB(100, 0, 0, 0)
            
            '//Reset
            yOffset = 0
            '//Loop to all items
            For i = LowBound To UpBound
                '//Set Location
                '//Keep it centered
                X = (Screen_Width / 2) - (GetTextWidth(Font_Default, ArrayText(i)) / 2)
                Y = (Screen_Height / 2) - ((16 * UpBound) / 2) + yOffset
                
                '//Render the text
                RenderText Font_Default, ArrayText(i), X, Y, White
                
                '//Increase the location for each line
                yOffset = yOffset + 16
            Next
        Else
            '//Get the longest width
            MaxWidth = GetTextWidth(Font_Default, LoadText)
            
            '//Set Location
            '//Keep it centered
            X = (Screen_Width / 2) - (MaxWidth / 2)
            Y = (Screen_Height / 2) - (16 / 2)
            
            '//Draw the hud of the text
            RenderTexture Tex_System(gSystemEnum.UserInterface), X - PaddingSize, Y - PaddingSize, 0, 8, MaxWidth + (PaddingSize * 2), 16 + (PaddingSize * 2), 1, 1, D3DColorARGB(100, 0, 0, 0)
            
            '//Render the text
            RenderText Font_Default, LoadText, X, Y, White
        End If
    End If
End Sub

Private Sub DrawMapTile(ByVal Layer As MapLayer, ByVal X As Long, ByVal Y As Long, Optional ByVal Alpha As Byte = 255)
Dim MapTile As Byte
Dim AnimMapTile As Byte

    If GettingMap Then Exit Sub

    If IsValidMapPoint(X, Y) Then
        '//Check if there's a animated tile
        AnimMapTile = Map.Tile(X, Y).Layer(Layer, MapLayerType.Animated).Tile
        '//Exist
        If AnimMapTile > 0 And AnimMapTile <= Count_Tileset Then
            If MapAnim = YES Then
                With Map.Tile(X, Y).Layer(Layer, MapLayerType.Animated)
                    If .MapAnim > 0 Then
                        RenderTexture Tex_MapAnim(.MapAnim), ConvertMapX(X * TILE_X), ConvertMapY(Y * TILE_Y), PIC_X * MapFrameAnim, 0, TILE_X, TILE_Y, PIC_X, PIC_Y, D3DColorARGB(Alpha, 255, 255, 255)
                    Else
                        RenderTexture Tex_Tileset(.Tile), ConvertMapX(X * TILE_X), ConvertMapY(Y * TILE_Y), .TileX * PIC_X, .TileY * PIC_Y, TILE_X, TILE_Y, PIC_X, PIC_Y, D3DColorARGB(Alpha, 255, 255, 255)
                    End If
                End With
            Else
                MapTile = Map.Tile(X, Y).Layer(Layer, MapLayerType.Normal).Tile
                If MapTile > 0 And MapTile <= Count_Tileset Then
                    With Map.Tile(X, Y).Layer(Layer, MapLayerType.Normal)
                        If .MapAnim > 0 Then
                            RenderTexture Tex_MapAnim(.MapAnim), ConvertMapX(X * TILE_X), ConvertMapY(Y * TILE_Y), PIC_X * MapFrameAnim, 0, TILE_X, TILE_Y, PIC_X, PIC_Y, D3DColorARGB(Alpha, 255, 255, 255)
                        Else
                            RenderTexture Tex_Tileset(.Tile), ConvertMapX(X * TILE_X), ConvertMapY(Y * TILE_Y), .TileX * PIC_X, .TileY * PIC_Y, TILE_X, TILE_Y, PIC_X, PIC_Y, D3DColorARGB(Alpha, 255, 255, 255)
                        End If
                    End With
                End If
            End If
        Else
            MapTile = Map.Tile(X, Y).Layer(Layer, MapLayerType.Normal).Tile
            If MapTile > 0 And MapTile <= Count_Tileset Then
                With Map.Tile(X, Y).Layer(Layer, MapLayerType.Normal)
                    If .MapAnim > 0 Then
                        RenderTexture Tex_MapAnim(.MapAnim), ConvertMapX(X * TILE_X), ConvertMapY(Y * TILE_Y), PIC_X * MapFrameAnim, 0, TILE_X, TILE_Y, PIC_X, PIC_Y, D3DColorARGB(Alpha, 255, 255, 255)
                    Else
                        RenderTexture Tex_Tileset(.Tile), ConvertMapX(X * TILE_X), ConvertMapY(Y * TILE_Y), .TileX * PIC_X, .TileY * PIC_Y, TILE_X, TILE_Y, PIC_X, PIC_Y, D3DColorARGB(Alpha, 255, 255, 255)
                    End If
                End With
            End If
        End If
    End If
End Sub

Private Sub DrawPlayer(ByVal Index As Long)
Dim Width As Long, Height As Long
Dim oWidth As Long, oHeight As Long
Dim X As Long, Y As Long
Dim Anim As Long, rDir As Byte
Dim Sprite As Long
Dim DrawAlpha As Long

    '//Check error
    If Index <= 0 Or Index > MAX_PLAYER Then Exit Sub
    If Not IsPlaying(Index) Then Exit Sub
    
    With Player(Index)
        ' Check if Player is within screen area
        If .X < TileView.Left Or .X > TileView.Right Then Exit Sub
        If .Y < TileView.top Or .Y > TileView.bottom Then Exit Sub
        
        Sprite = .Sprite
        
        Select Case .TempSprite
            Case TEMP_SPRITE_GROUP_BIKE
                '//Empty sprite? then exit out
                If Sprite <= 0 Or Sprite > Count_PlayerSprite_B Then Exit Sub
                '//Check sprite size
                oWidth = GetPicWidth(Tex_PlayerSprite_B(Sprite)) / 3
                oHeight = GetPicHeight(Tex_PlayerSprite_B(Sprite)) / 4
            Case TEMP_SPRITE_GROUP_DIVE
                '//Empty sprite? then exit out
                If Sprite <= 0 Or Sprite > Count_PlayerSprite_D Then Exit Sub
                '//Check sprite size
                oWidth = GetPicWidth(Tex_PlayerSprite_D(Sprite)) / 3
                oHeight = GetPicHeight(Tex_PlayerSprite_D(Sprite)) / 4
            Case Else
                '//Empty sprite? then exit out
                If Sprite <= 0 Or Sprite > Count_PlayerSprite_N Then Exit Sub
                '//Check sprite size
                oWidth = GetPicWidth(Tex_PlayerSprite_N(Sprite)) / 3
                oHeight = GetPicHeight(Tex_PlayerSprite_N(Sprite)) / 4
        End Select
        Width = oWidth * 2: Height = oHeight * 2
        
        Anim = 1 '//Default Anim  "0 | >1< | 2"
        
        '//Moving
        Select Case .Dir
            Case DIR_UP
                If (.yOffset > 8) Then Anim = .Step
            Case DIR_DOWN
                If (.yOffset < -8) Then Anim = .Step
            Case DIR_LEFT
                If (.xOffset > 8) Then Anim = .Step
            Case DIR_RIGHT
                If (.xOffset < -8) Then Anim = .Step
        End Select
        
        If .Action = ACTION_SLIDE Then
            Anim = 2
        End If
        
        '//Checking Direction
        Select Case .Dir
            Case DIR_UP: rDir = 2
            Case DIR_DOWN: rDir = 0
            Case DIR_LEFT: rDir = 1
            Case DIR_RIGHT: rDir = 3
        End Select
        
        '//Set position on center of the tile
        X = (.X * TILE_X) + .xOffset - ((Width - TILE_X) / 2)
        Y = (.Y * TILE_Y) + .yOffset - (Height - TILE_Y)
        
        If .StealthMode = YES Then
            If Index <> MyIndex Then
                DrawAlpha = 0
            Else
                DrawAlpha = 70
            End If
        Else
            DrawAlpha = 255
        End If
        
        Select Case .TempSprite
            Case TEMP_SPRITE_GROUP_BIKE: RenderTexture Tex_PlayerSprite_B(Sprite), ConvertMapX(X), ConvertMapY(Y), Anim * oWidth, rDir * oHeight, Width, Height, oWidth, oHeight, D3DColorARGB(DrawAlpha, 255, 255, 255)
            Case TEMP_SPRITE_GROUP_DIVE: RenderTexture Tex_PlayerSprite_D(Sprite), ConvertMapX(X), ConvertMapY(Y), Anim * oWidth, rDir * oHeight, Width, Height, oWidth, oHeight, D3DColorARGB(DrawAlpha, 255, 255, 255)
            Case Else: RenderTexture Tex_PlayerSprite_N(Sprite), ConvertMapX(X), ConvertMapY(Y), Anim * oWidth, rDir * oHeight, Width, Height, oWidth, oHeight, D3DColorARGB(DrawAlpha, 255, 255, 255)
        End Select
    End With
End Sub

Private Sub DrawNpc(ByVal MapNpcNum As Long)
Dim Width As Long, Height As Long
Dim oWidth As Long, oHeight As Long
Dim X As Long, Y As Long
Dim Anim As Long, rDir As Byte
Dim Sprite As Long

    '//Check error
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPC Then Exit Sub
    If Map.Npc(MapNpcNum) <= 0 Then Exit Sub
    If MapNpc(MapNpcNum).Num <= 0 Then Exit Sub
    
    With MapNpc(MapNpcNum)
        ' Check if Player is within screen area
        If .X < TileView.Left Or .X > TileView.Right Then Exit Sub
        If .Y < TileView.top Or .Y > TileView.bottom Then Exit Sub
        
        '//Empty sprite? then exit out
        Sprite = Npc(.Num).Sprite
        If Sprite <= 0 Or Sprite > Count_Character Then Exit Sub
        
        Anim = 1 '//Default Anim  "0 | >1< | 2"
        
        '//Moving
        Select Case .Dir
            Case DIR_UP
                If (.yOffset > 8) Then Anim = .Step
            Case DIR_DOWN
                If (.yOffset < -8) Then Anim = .Step
            Case DIR_LEFT
                If (.xOffset > 8) Then Anim = .Step
            Case DIR_RIGHT
                If (.xOffset < -8) Then Anim = .Step
        End Select
        
        '//Check sprite size
        oWidth = GetPicWidth(Tex_Character(Sprite)) / 3
        oHeight = GetPicHeight(Tex_Character(Sprite)) / 4
        Width = oWidth * 2: Height = oHeight * 2
        
        '//Checking Direction
        Select Case .Dir
            Case DIR_UP: rDir = 2
            Case DIR_DOWN: rDir = 0
            Case DIR_LEFT: rDir = 1
            Case DIR_RIGHT: rDir = 3
        End Select
        
        '//Set position on center of the tile
        X = (.X * TILE_X) + .xOffset - ((Width - TILE_X) / 2)
        Y = (.Y * TILE_Y) + .yOffset - (Height - TILE_Y)
        
        '//Render
        RenderTexture Tex_Character(Sprite), ConvertMapX(X), ConvertMapY(Y), Anim * oWidth, rDir * oHeight, Width, Height, oWidth, oHeight
    End With
End Sub

Private Sub DrawPokemon(ByVal PokemonIndex As Long)
Dim Width As Long, Height As Long
Dim oWidth As Long, oHeight As Long
Dim X As Long, Y As Long
Dim Sprite As Long
Dim AttackSpeed As Long

Dim Anim As Long
Dim SpriteAnim As Long
Dim SpritePos As Byte
Dim Name As String

    '//Check error
    If PokemonIndex <= 0 Or PokemonIndex > MAX_GAME_POKEMON Then Exit Sub
    If MapPokemon(PokemonIndex).Num <= 0 Then Exit Sub
    
    With MapPokemon(PokemonIndex)
        ' Check if Player is within screen area
        If .X < TileView.Left Or .X > TileView.Right Then Exit Sub
        If .Y < TileView.top Or .Y > TileView.bottom Then Exit Sub
        
        '//Empty sprite? then exit out
        Sprite = Pokemon(.Num).Sprite
        If Sprite <= 0 Or Sprite > Count_Pokemon Then Exit Sub
        
        '//Check sprite size
        oWidth = GetPicWidth(Tex_Pokemon(Sprite)) / 34
        oHeight = GetPicHeight(Tex_Pokemon(Sprite))
        If Pokemon(.Num).ScaleSprite = YES Then
            Width = oWidth * 2
            Height = oHeight * 2
        Else
            Width = oWidth
            Height = oHeight
        End If
        
        AttackSpeed = 1000
        '//Check if attacking
        If .AttackTimer + (AttackSpeed / 2) > GetTickCount Then
            If .Attacking = YES Then
                SpritePos = 3 '//Attacking
            End If
        Else
            If .IdleTimer + 500 < GetTickCount Then
                If .IdleFrameTmr + 500 < GetTickCount Then
                    .IdleAnim = .IdleAnim + 1
                    If .IdleAnim > 2 Then
                        .IdleAnim = 0
                    End If
                    .IdleFrameTmr = GetTickCount
                End If
                Anim = .IdleAnim
                SpritePos = 1 '//Idle
            Else
                '//Stand
                Anim = 1 '//Default Anim  "0 | >1< | 2"
                '//Moving
                Select Case .Dir
                    Case DIR_UP
                        If (.yOffset > 8) Then Anim = .Step
                    Case DIR_DOWN
                        If (.yOffset < -8) Then Anim = .Step
                    Case DIR_LEFT
                        If (.xOffset > 8) Then Anim = .Step
                    Case DIR_RIGHT
                        If (.xOffset < -8) Then Anim = .Step
                End Select
                SpritePos = 2 '//Walking
            End If
        End If
        
        If .AttackTimer + AttackSpeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
        
        '//Checking Direction
        Select Case .Dir
            Case DIR_UP
                If SpritePos = 1 Then '//Idle
                    SpriteAnim = Anim
                ElseIf SpritePos = 2 Then '//Walking
                    SpriteAnim = 12 + Anim
                ElseIf SpritePos = 3 Then '//Attacking
                    SpriteAnim = 30
                End If
            Case DIR_DOWN
                If SpritePos = 1 Then '//Idle
                    SpriteAnim = 3 + Anim
                ElseIf SpritePos = 2 Then '//Walking
                    SpriteAnim = 12 + 3 + Anim
                ElseIf SpritePos = 3 Then '//Attacking
                    SpriteAnim = 31
                End If
            Case DIR_LEFT
                If SpritePos = 1 Then '//Idle
                    SpriteAnim = 6 + Anim
                ElseIf SpritePos = 2 Then '//Walking
                    SpriteAnim = 12 + 6 + Anim
                ElseIf SpritePos = 3 Then '//Attacking
                    SpriteAnim = 32
                End If
            Case DIR_RIGHT
                If SpritePos = 1 Then '//Idle
                    SpriteAnim = 9 + Anim
                ElseIf SpritePos = 2 Then '//Walking
                    SpriteAnim = 12 + 9 + Anim
                ElseIf SpritePos = 3 Then '//Attacking
                    SpriteAnim = 33
                End If
        End Select
        
        '//Sleeping
        If .Status = StatusEnum.Sleep Then
            SpriteAnim = 24 + MapAnim
        End If
        
        '//Set position on center of the tile
        X = (.X * TILE_X) + .xOffset - ((Width - TILE_X) / 2)
        Y = (.Y * TILE_Y) + .yOffset - (Height - TILE_Y)
        
        '//Render
        If .IsShiny = YES Then
            If Sprite > 0 And Sprite <= Count_ShinyPokemon Then
                RenderTexture Tex_ShinyPokemon(Sprite), ConvertMapX(X), ConvertMapY(Y), SpriteAnim * oWidth, oHeight, Width, Height, oWidth, oHeight
            End If
        Else
            RenderTexture Tex_Pokemon(Sprite), ConvertMapX(X), ConvertMapY(Y), SpriteAnim * oWidth, oHeight, Width, Height, oWidth, oHeight
        End If
    End With
End Sub

Private Sub DrawPlayerPokemon(ByVal Index As Long)
Dim Width As Long, Height As Long
Dim oWidth As Long, oHeight As Long
Dim X As Long, Y As Long
Dim Sprite As Long
Dim AttackSpeed As Long

Dim Anim As Long
Dim SpriteAnim As Long
Dim SpritePos As Byte

    '//Check error
    If Index <= 0 Or Index > MAX_PLAYER Then Exit Sub
    If PlayerPokemon(Index).Num <= 0 Then Exit Sub
    
    With PlayerPokemon(Index)
        ' Check if Player is within screen area
        If .X < TileView.Left Or .X > TileView.Right Then Exit Sub
        If .Y < TileView.top Or .Y > TileView.bottom Then Exit Sub
        
        '//Empty sprite? then exit out
        Sprite = Pokemon(.Num).Sprite
        If Sprite <= 0 Or Sprite > Count_Pokemon Then Exit Sub
        
        '//Check sprite size
        oWidth = GetPicWidth(Tex_Pokemon(Sprite)) / 34
        oHeight = GetPicHeight(Tex_Pokemon(Sprite))
        If Pokemon(.Num).ScaleSprite = YES Then
            Width = oWidth * 2
            Height = oHeight * 2
        Else
            Width = oWidth
            Height = oHeight
        End If
        
        AttackSpeed = 1000
        '//Check if attacking
        If .AttackTimer + (AttackSpeed / 2) > GetTickCount Then
            If .Attacking = YES Then
                SpritePos = 3 '//Attacking
            End If
        Else
            If .IdleTimer + 500 < GetTickCount Then
                If .IdleFrameTmr + 500 < GetTickCount Then
                    .IdleAnim = .IdleAnim + 1
                    If .IdleAnim > 2 Then
                        .IdleAnim = 0
                    End If
                    .IdleFrameTmr = GetTickCount
                End If
                Anim = .IdleAnim
                SpritePos = 1 '//Idle
            Else
                '//Stand
                Anim = 1 '//Default Anim  "0 | >1< | 2"
                '//Moving
                Select Case .Dir
                    Case DIR_UP
                        If (.yOffset > 8) Then Anim = .Step
                    Case DIR_DOWN
                        If (.yOffset < -8) Then Anim = .Step
                    Case DIR_LEFT
                        If (.xOffset > 8) Then Anim = .Step
                    Case DIR_RIGHT
                        If (.xOffset < -8) Then Anim = .Step
                End Select
                SpritePos = 2 '//Walking
            End If
        End If
        
        If .AttackTimer + AttackSpeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
        
        '//Checking Direction
        Select Case .Dir
            Case DIR_UP
                If SpritePos = 1 Then '//Idle
                    SpriteAnim = Anim
                ElseIf SpritePos = 2 Then '//Walking
                    SpriteAnim = 12 + Anim
                ElseIf SpritePos = 3 Then '//Attacking
                    SpriteAnim = 30
                End If
            Case DIR_DOWN
                If SpritePos = 1 Then '//Idle
                    SpriteAnim = 3 + Anim
                ElseIf SpritePos = 2 Then '//Walking
                    SpriteAnim = 12 + 3 + Anim
                ElseIf SpritePos = 3 Then '//Attacking
                    SpriteAnim = 31
                End If
            Case DIR_LEFT
                If SpritePos = 1 Then '//Idle
                    SpriteAnim = 6 + Anim
                ElseIf SpritePos = 2 Then '//Walking
                    SpriteAnim = 12 + 6 + Anim
                ElseIf SpritePos = 3 Then '//Attacking
                    SpriteAnim = 32
                End If
            Case DIR_RIGHT
                If SpritePos = 1 Then '//Idle
                    SpriteAnim = 9 + Anim
                ElseIf SpritePos = 2 Then '//Walking
                    SpriteAnim = 12 + 9 + Anim
                ElseIf SpritePos = 3 Then '//Attacking
                    SpriteAnim = 33
                End If
        End Select
        
        '//Sleeping
        If .Status = StatusEnum.Sleep Then
            SpriteAnim = 24 + MapAnim
        End If
        
        '//Set position on center of the tile
        X = (.X * TILE_X) + .xOffset - ((Width - TILE_X) / 2)
        Y = (.Y * TILE_Y) + .yOffset - (Height - TILE_Y)

        '//Render
        If .IsShiny = YES Then
            If Sprite > 0 And Sprite <= Count_ShinyPokemon Then
                RenderTexture Tex_ShinyPokemon(Sprite), ConvertMapX(X), ConvertMapY(Y), SpriteAnim * oWidth, oHeight, Width, Height, oWidth, oHeight
            End If
        Else
            RenderTexture Tex_Pokemon(Sprite), ConvertMapX(X), ConvertMapY(Y), SpriteAnim * oWidth, oHeight, Width, Height, oWidth, oHeight
        End If
    End With
End Sub

Private Sub DrawMapNpcPokemon(ByVal Index As Long)
Dim Width As Long, Height As Long
Dim oWidth As Long, oHeight As Long
Dim X As Long, Y As Long
Dim Sprite As Long
Dim AttackSpeed As Long

Dim Anim As Long
Dim SpriteAnim As Long
Dim SpritePos As Byte

    '//Check error
    If Index <= 0 Or Index > MAX_MAP_NPC Then Exit Sub
    If MapNpcPokemon(Index).Num <= 0 Then Exit Sub
    
    With MapNpcPokemon(Index)
        ' Check if Player is within screen area
        If .X < TileView.Left Or .X > TileView.Right Then Exit Sub
        If .Y < TileView.top Or .Y > TileView.bottom Then Exit Sub
        
        '//Empty sprite? then exit out
        Sprite = Pokemon(.Num).Sprite
        If Sprite <= 0 Or Sprite > Count_Pokemon Then Exit Sub
        
        '//Check sprite size
        oWidth = GetPicWidth(Tex_Pokemon(Sprite)) / 34
        oHeight = GetPicHeight(Tex_Pokemon(Sprite))
        If Pokemon(.Num).ScaleSprite = YES Then
            Width = oWidth * 2
            Height = oHeight * 2
        Else
            Width = oWidth
            Height = oHeight
        End If
        
        AttackSpeed = 1000
        '//Check if attacking
        If .AttackTimer + (AttackSpeed / 2) > GetTickCount Then
            If .Attacking = YES Then
                SpritePos = 3 '//Attacking
            End If
        Else
            If .IdleTimer + 500 < GetTickCount Then
                If .IdleFrameTmr + 500 < GetTickCount Then
                    .IdleAnim = .IdleAnim + 1
                    If .IdleAnim > 2 Then
                        .IdleAnim = 0
                    End If
                    .IdleFrameTmr = GetTickCount
                End If
                Anim = .IdleAnim
                SpritePos = 1 '//Idle
            Else
                '//Stand
                Anim = 1 '//Default Anim  "0 | >1< | 2"
                '//Moving
                Select Case .Dir
                    Case DIR_UP
                        If (.yOffset > 8) Then Anim = .Step
                    Case DIR_DOWN
                        If (.yOffset < -8) Then Anim = .Step
                    Case DIR_LEFT
                        If (.xOffset > 8) Then Anim = .Step
                    Case DIR_RIGHT
                        If (.xOffset < -8) Then Anim = .Step
                End Select
                SpritePos = 2 '//Walking
            End If
        End If
        
        If .AttackTimer + AttackSpeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
        
        '//Checking Direction
        Select Case .Dir
            Case DIR_UP
                If SpritePos = 1 Then '//Idle
                    SpriteAnim = Anim
                ElseIf SpritePos = 2 Then '//Walking
                    SpriteAnim = 12 + Anim
                ElseIf SpritePos = 3 Then '//Attacking
                    SpriteAnim = 30
                End If
            Case DIR_DOWN
                If SpritePos = 1 Then '//Idle
                    SpriteAnim = 3 + Anim
                ElseIf SpritePos = 2 Then '//Walking
                    SpriteAnim = 12 + 3 + Anim
                ElseIf SpritePos = 3 Then '//Attacking
                    SpriteAnim = 31
                End If
            Case DIR_LEFT
                If SpritePos = 1 Then '//Idle
                    SpriteAnim = 6 + Anim
                ElseIf SpritePos = 2 Then '//Walking
                    SpriteAnim = 12 + 6 + Anim
                ElseIf SpritePos = 3 Then '//Attacking
                    SpriteAnim = 32
                End If
            Case DIR_RIGHT
                If SpritePos = 1 Then '//Idle
                    SpriteAnim = 9 + Anim
                ElseIf SpritePos = 2 Then '//Walking
                    SpriteAnim = 12 + 9 + Anim
                ElseIf SpritePos = 3 Then '//Attacking
                    SpriteAnim = 33
                End If
        End Select
        
        '//Sleeping
        'If .Status = StatusEnum.Sleep Then
        '    SpriteAnim = 24 + MapAnim
        'End If
        
        '//Set position on center of the tile
        X = (.X * TILE_X) + .xOffset - ((Width - TILE_X) / 2)
        Y = (.Y * TILE_Y) + .yOffset - (Height - TILE_Y)

        '//Render
        If .IsShiny = YES Then
            If Sprite > 0 And Sprite <= Count_ShinyPokemon Then
                RenderTexture Tex_ShinyPokemon(Sprite), ConvertMapX(X), ConvertMapY(Y), SpriteAnim * oWidth, oHeight, Width, Height, oWidth, oHeight
            End If
        Else
            RenderTexture Tex_Pokemon(Sprite), ConvertMapX(X), ConvertMapY(Y), SpriteAnim * oWidth, oHeight, Width, Height, oWidth, oHeight
        End If
    End With
End Sub

Private Sub DrawActionMsg(ByVal Index As Integer)
Dim X As Long, Y As Long, i As Long
Dim Alpha As Long
Dim time As Long

    '//Exit out of there's no message
    If ActionMsg(Index).Msg = vbNullString Then Exit Sub

    '//Set the timer
    time = 1500
    If ActionMsg(Index).Y > 0 Then
        X = ActionMsg(Index).X + (TILE_X / 2) - (GetTextWidth(Font_Default, Trim$(ActionMsg(Index).Msg)) / 2)
        Y = ActionMsg(Index).Y - (TILE_Y / 2) - 2 - (ActionMsg(Index).Scroll * 0.3)
        ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
    Else
        X = ActionMsg(Index).X + (TILE_X / 2) - (GetTextWidth(Font_Default, Trim$(ActionMsg(Index).Msg)) / 2)
        Y = ActionMsg(Index).Y - (TILE_Y / 2) + 18 + (ActionMsg(Index).Scroll * 0.3)
        ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
    End If

    '//Fade while scrolling
    ActionMsg(Index).Alpha = ActionMsg(Index).Alpha - 1
    If ActionMsg(Index).Alpha <= 0 Then ClearActionMsg Index: Exit Sub
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)

    If GetTickCount < ActionMsg(Index).Created + time Then
        RenderText Font_Default, ActionMsg(Index).Msg, X, Y, ActionMsg(Index).Color, True, ActionMsg(Index).Alpha
    Else
        ClearActionMsg Index
    End If
End Sub

'//Animation
Public Sub DrawAnimation(ByVal Index As Long, ByVal Layer As Long)
Dim Sprite As Long, FrameCount As Long
Dim Width As Long, Height As Long, X As Long, Y As Long
Dim sRect As RECT

    If AnimInstance(Index).Animation = 0 Then
        ClearAnimInstance Index
        Exit Sub
    End If
    
    Sprite = Animation(AnimInstance(Index).Animation).Sprite(Layer)
    If Sprite < 1 Or Sprite > Count_Animation Then Exit Sub
    
    FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
    If FrameCount <= 0 Then Exit Sub
    
    '//total width divided by frame count
    Width = GetPicWidth(Tex_Animation(Sprite)) / FrameCount 'AnimColumns '/ FrameCount
    Height = GetPicHeight(Tex_Animation(Sprite)) 'GetPicWidth(Tex_Animation(Sprite)) '/ AnimColumns 'GetPicHeight(Tex_Animation(Sprite))
    
    With sRect
        .top = 0 '(Height * ((AnimInstance(Index).frameIndex(Layer) - 1) \ AnimColumns)) '0
        .bottom = Height
        .Left = (AnimInstance(Index).frameIndex(Layer) - 1) * Width '(Width * (((AnimInstance(Index).frameIndex(Layer) - 1) Mod AnimColumns))) '(AnimInstance(Index).frameIndex(Layer) - 1) * Width
        .Right = sRect.Left + Width 'Width 'sRect.Left + Width
    End With
    
    '//no lock, default x + y
    X = (AnimInstance(Index).X * 32) + 16 - (Width / 2)
    Y = (AnimInstance(Index).Y * 32) + 16 - (Height / 2)
    
    '//Clipping
    If Y < 0 Then
        With sRect
            .top = .top - Y
        End With
        Y = 0
    End If
    If X < 0 Then
        With sRect
            .Left = .Left - X
        End With
        X = 0
    End If
    
    'RenderTexture Tex_Animation(Sprite), ConvertMapX(X), ConvertMapY(Y), sRect.Left, sRect.top, sRect.Right, sRect.bottom, sRect.Right, sRect.bottom
    RenderTexture Tex_Animation(Sprite), ConvertMapX(X), ConvertMapY(Y), sRect.Left, sRect.top, sRect.Right - sRect.Left, sRect.bottom - sRect.top, sRect.Right - sRect.Left, sRect.bottom - sRect.top
End Sub

Public Sub DrawPokeball(ByVal X As Long, ByVal Y As Long, ByVal Frame As Byte, ByVal Pic As Byte)
Dim DrawX As Long, DrawY As Long

    DrawX = (X * TILE_X) + ((TILE_X / 2) - 10)
    DrawY = (Y * TILE_Y) + ((TILE_Y / 2) - 13)
    RenderTexture Tex_Misc(Misc_Pokeball), ConvertMapX(DrawX), ConvertMapY(DrawY), Frame * 20, Pic * 26, 20, 26, 20, 26
End Sub

' **********
' ** Misc **
' **********
Private Sub DrawCompanyScreen()
Dim Width As Long, Height As Long
Dim X As Long, Y As Long

    '//Make sure is not loading
    If IsLoading Then Exit Sub

    '//First we must turn the whole screen into black
    RenderTexture Tex_System(gSystemEnum.UserInterface), 0, 0, 0, 8, Screen_Width, Screen_Height, 1, 1, D3DColorARGB(255, 0, 0, 0)
    
    '//Render the company logo which present the game in the middle of the screen
    '//Get Size for quick access
    Width = GetPicWidth(Tex_Surface(gSurfaceEnum.CompanyScreen))
    Height = GetPicHeight(Tex_Surface(gSurfaceEnum.CompanyScreen))
    '//Set Location to center
    X = (Screen_Width / 2) - (Width / 2)
    Y = (Screen_Height / 2) - (Height / 2)
    
    RenderTexture Tex_Surface(gSurfaceEnum.CompanyScreen), X, Y, 0, 0, Width, Height, Width, Height
End Sub

Private Sub DrawTitleScreen()
Dim Width As Long, Height As Long
Dim X As Long, Y As Long

Dim i As Long

    '//Make sure is not loading
    If IsLoading Then Exit Sub

    '//First we must turn the whole screen into the color that match the background
    RenderTexture Tex_System(gSystemEnum.UserInterface), 0, 0, 0, 8, Screen_Width, Screen_Height, 1, 1, D3DColorARGB(255, 0, 93, 165)
    
    '//Get Size for quick access
    Width = GetPicWidth(Tex_Surface(gSurfaceEnum.Background))
    Height = GetPicHeight(Tex_Surface(gSurfaceEnum.Background))
    
    For i = -1 To (Screen_Width / Width) + 1
        '//Render the whole background on screen (scale it to fit the width of the screen size)
        RenderTexture Tex_Surface(gSurfaceEnum.Background), (i * Width) + BackgroundXOffset, Screen_Height - Height, 0, 0, Width, Height, Width, Height
    Next
    
    '//Render the company logo which present the game in the middle of the screen
    '//Get Size for quick access
    Width = GetPicWidth(Tex_Surface(gSurfaceEnum.TitleScreen))
    Height = GetPicHeight(Tex_Surface(gSurfaceEnum.TitleScreen))
    '//Set Location to center
    X = (Screen_Width / 2) - (Width / 2)
    Y = (Screen_Height / 2) - (Height / 2)
    
    RenderTexture Tex_Surface(gSurfaceEnum.TitleScreen), X, Y, 0, 0, Width, Height, Width, Height
End Sub

Private Sub DrawBackground()
Dim Width As Long, Height As Long
Dim textCredit As String
Dim textX As Long, textY As Long
Dim i As Long

    '//Hovering
    IsHovering = False

    '//First we must turn the whole screen into the color that match the background
    'RenderTexture Tex_System(gSystemEnum.UserInterface), 0, 0, 0, 8, Screen_Width, Screen_Height, 1, 1, D3DColorARGB(255, 0, 120, 191)
    RenderTexture Tex_System(gSystemEnum.UserInterface), 0, 0, 0, 8, Screen_Width, Screen_Height, 1, 1, D3DColorARGB(255, 0, 93, 165)
    
    '//Get Size for quick access
    Width = GetPicWidth(Tex_Surface(gSurfaceEnum.Background))
    Height = GetPicHeight(Tex_Surface(gSurfaceEnum.Background))
    
    For i = -1 To (Screen_Width / Width) + 1
        '//Render the whole background on screen (scale it to fit the width of the screen size)
        RenderTexture Tex_Surface(gSurfaceEnum.Background), (i * Width) + BackgroundXOffset, Screen_Height - Height, 0, 0, Width, Height, Width, Height
    Next

    DrawCredit
End Sub


Public Sub DrawSelMenu()
Dim i As Long, MaxHeight As Long
Dim X As Long, Y As Long

    '//Make sure we are not in editor
    If Editor > 0 Then Exit Sub
    
    With SelMenu
        If Not .Visible Then Exit Sub
        If .MaxText <= 0 Then Exit Sub

        '//ToDo: Moving Target
        X = .X
        Y = .Y
        
        IsHovering = False
        
        '//Reset Pick
        .CurPick = 0
        For i = 1 To .MaxText
            If CursorX >= X + 5 And CursorX <= X + 5 + .MaxWidth And CursorY >= Y + 5 + ((i - 1) * 18) And CursorY <= Y + 5 + ((i - 1) * 18) + 16 Then
                .CurPick = i
                IsHovering = True
                MouseIcon = 1 '//Select
            End If
        Next
        
        '//Top Left
        RenderTexture Tex_System(gSystemEnum.UserInterface), X, Y, 33, 0, 5, 5, 5, 5
        '//Top
        RenderTexture Tex_System(gSystemEnum.UserInterface), X + 5, Y, 38, 0, .MaxWidth + 5, 5, 5, 5
        '//Top Right
        RenderTexture Tex_System(gSystemEnum.UserInterface), X + .MaxWidth + 10, Y, 43, 0, 5, 5, 5, 5
        '//Left
        RenderTexture Tex_System(gSystemEnum.UserInterface), X, Y + 5, 33, 5, 5, (.MaxText * 18), 5, 5
        '//Center
        RenderTexture Tex_System(gSystemEnum.UserInterface), X + 5, Y + 5, 38, 5, .MaxWidth + 5, (.MaxText * 18), 5, 5
        '//Right
        RenderTexture Tex_System(gSystemEnum.UserInterface), X + .MaxWidth + 10, Y + 5, 43, 5, 5, (.MaxText * 18), 5, 5
        '//Bottom Left
        RenderTexture Tex_System(gSystemEnum.UserInterface), X, Y + (.MaxText * 18) + 5, 33, 8, 5, 7, 5, 7
        ' Bottom
        RenderTexture Tex_System(gSystemEnum.UserInterface), X + 5, Y + (.MaxText * 18) + 5, 38, 8, .MaxWidth + 5, 7, 5, 7
        ' Bottom Right
        RenderTexture Tex_System(gSystemEnum.UserInterface), X + .MaxWidth + 10, Y + (.MaxText * 18) + 5, 43, 8, 5, 7, 5, 7
        
        For i = 1 To .MaxText
            If .CurPick = i Then
                RenderTexture Tex_System(gSystemEnum.UserInterface), X + 4, Y + 5 + ((i - 1) * 18), 48, 0, .MaxWidth + 7, 18, 5, 5
            End If
            RenderText Font_Default, Trim$(.Text(i)), X + 5, Y + 5 + ((i - 1) * 18), White
        Next
    End With
End Sub

Private Sub DrawVitalBar()
Dim i As Long
Dim Width As Long, Height As Long
Dim MaxWidth As Long
Dim X As Long, Y As Long
Dim Color As Long

    If MyIndex <= 0 Then Exit Sub
    
    MaxWidth = GetPicWidth(Tex_Misc(Misc_Bar))
    Height = GetPicHeight(Tex_Misc(Misc_Bar)) / 2

    If Pokemon_HighIndex > 0 Then
        For i = 1 To Pokemon_HighIndex
            If Not CatchBall(i).InUsed Then
                If MapPokemon(i).Num > 0 Then
                    If MapPokemon(i).Map = Player(MyIndex).Map Then
                        If MapPokemon(i).CurHP < MapPokemon(i).MaxHP Then
                            '//get position
                            Width = (MapPokemon(i).CurHP / (MaxWidth - 6)) / (MapPokemon(i).MaxHP / (MaxWidth - 6)) * (MaxWidth - 6)
                            X = ((MapPokemon(i).X * TILE_X) + MapPokemon(i).xOffset) - ((MaxWidth / 2) - (TILE_X / 2))
                            Y = ((MapPokemon(i).Y * TILE_Y) + MapPokemon(i).yOffset) - ((Height / 2) - (TILE_Y / 2)) + 25
                        
                            '//placeholder
                            RenderTexture Tex_Misc(Misc_Bar), ConvertMapX(X), ConvertMapY(Y), 0, 0, MaxWidth, Height, MaxWidth, Height
                            
                            '//Get color
                            Select Case Width
                                Case (MaxWidth - 6) * 0.7 To (MaxWidth - 6)
                                    Color = D3DColorARGB(255, 34, 177, 76)
                                Case (MaxWidth - 6) * 0.3 To (MaxWidth - 6) * 0.7
                                    Color = D3DColorARGB(255, 255, 255, 0)
                                Case Else
                                    Color = D3DColorARGB(255, 255, 0, 0)
                            End Select
                            
                            '//Bar
                            RenderTexture Tex_Misc(Misc_Bar), ConvertMapX(X + 3), ConvertMapY(Y), 3, Height, Width, Height, (MaxWidth - 6), Height, Color
                        End If
                    End If
                End If
            End If
        Next
    End If
    If Player_HighIndex > 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If Player(i).Map = Player(MyIndex).Map Then
                    If Player(i).StealthMode = NO Then
                        If Player(i).CurHP < GetPlayerHP(Player(i).Level) Then
                            '//get position
                            Width = (Player(i).CurHP / (MaxWidth - 6)) / (GetPlayerHP(Player(i).Level) / (MaxWidth - 6)) * (MaxWidth - 6)
                            X = ((Player(i).X * TILE_X) + Player(i).xOffset) - ((MaxWidth / 2) - (TILE_X / 2))
                            Y = ((Player(i).Y * TILE_Y) + Player(i).yOffset) - ((Height / 2) - (TILE_Y / 2)) + 25
                            
                            '//placeholder
                            RenderTexture Tex_Misc(Misc_Bar), ConvertMapX(X), ConvertMapY(Y), 0, 0, MaxWidth, Height, MaxWidth, Height
                                
                            '//Get color
                            Select Case Width
                                Case (MaxWidth - 6) * 0.7 To (MaxWidth - 6)
                                    Color = D3DColorARGB(255, 34, 177, 76)
                                Case (MaxWidth - 6) * 0.3 To (MaxWidth - 6) * 0.7
                                    Color = D3DColorARGB(255, 255, 255, 0)
                                Case Else
                                    Color = D3DColorARGB(255, 255, 0, 0)
                            End Select
                                
                            '//Bar
                            RenderTexture Tex_Misc(Misc_Bar), ConvertMapX(X + 3), ConvertMapY(Y), 3, Height, Width, Height, (MaxWidth - 6), Height, Color
                        End If
                    End If
                    
                    If PlayerPokemon(i).Num > 0 And PlayerPokemon(i).Init = NO Then
                        If PlayerPokemon(i).CurHP < PlayerPokemon(i).MaxHP Then
                            '//get position
                            Width = (PlayerPokemon(i).CurHP / (MaxWidth - 6)) / (PlayerPokemon(i).MaxHP / (MaxWidth - 6)) * (MaxWidth - 6)
                            X = ((PlayerPokemon(i).X * TILE_X) + PlayerPokemon(i).xOffset) - ((MaxWidth / 2) - (TILE_X / 2))
                            Y = ((PlayerPokemon(i).Y * TILE_Y) + PlayerPokemon(i).yOffset) - ((Height / 2) - (TILE_Y / 2)) + 25
                        
                            '//placeholder
                            RenderTexture Tex_Misc(Misc_Bar), ConvertMapX(X), ConvertMapY(Y), 0, 0, MaxWidth, Height, MaxWidth, Height
                            
                            '//Get color
                            Select Case Width
                                Case (MaxWidth - 6) * 0.7 To (MaxWidth - 6)
                                    Color = D3DColorARGB(255, 34, 177, 76)
                                Case (MaxWidth - 6) * 0.3 To (MaxWidth - 6) * 0.7
                                    Color = D3DColorARGB(255, 255, 255, 0)
                                Case Else
                                    Color = D3DColorARGB(255, 255, 0, 0)
                            End Select
                            
                            '//Bar
                            RenderTexture Tex_Misc(Misc_Bar), ConvertMapX(X + 3), ConvertMapY(Y), 3, Height, Width, Height, (MaxWidth - 6), Height, Color
                        End If
                    End If
                End If
            End If
        Next
    End If
    For i = 1 To MAX_MAP_NPC
        If MapNpc(i).Num > 0 Then
            If MapNpcPokemon(i).Num > 0 And MapNpcPokemon(i).Init = NO Then
                If MapNpcPokemon(i).CurHP < MapNpcPokemon(i).MaxHP Then
                    '//get position
                    Width = (MapNpcPokemon(i).CurHP / (MaxWidth - 6)) / (MapNpcPokemon(i).MaxHP / (MaxWidth - 6)) * (MaxWidth - 6)
                    X = ((MapNpcPokemon(i).X * TILE_X) + MapNpcPokemon(i).xOffset) - ((MaxWidth / 2) - (TILE_X / 2))
                    Y = ((MapNpcPokemon(i).Y * TILE_Y) + MapNpcPokemon(i).yOffset) - ((Height / 2) - (TILE_Y / 2)) + 25
                        
                    '//placeholder
                    RenderTexture Tex_Misc(Misc_Bar), ConvertMapX(X), ConvertMapY(Y), 0, 0, MaxWidth, Height, MaxWidth, Height
                            
                    '//Get color
                    Select Case Width
                        Case (MaxWidth - 6) * 0.7 To (MaxWidth - 6)
                            Color = D3DColorARGB(255, 34, 177, 76)
                        Case (MaxWidth - 6) * 0.3 To (MaxWidth - 6) * 0.7
                            Color = D3DColorARGB(255, 255, 255, 0)
                        Case Else
                            Color = D3DColorARGB(255, 255, 0, 0)
                    End Select
                            
                    '//Bar
                    RenderTexture Tex_Misc(Misc_Bar), ConvertMapX(X + 3), ConvertMapY(Y), 3, Height, Width, Height, (MaxWidth - 6), Height, Color
                End If
            End If
        End If
    Next
End Sub

Private Sub DrawMoveSelector()
    Dim X As Long, Y As Long
    Dim mX As Long, mY As Long
    Dim Width As Long, Height As Long
    Dim MoveNum As Long
    Dim barWidth As Long
    Dim guiAlpha As Byte
    Dim MoveSlot As Byte

    '//Check if can render
    If Not chkMoveKey And Not IsTryingToSwitchAttack Then Exit Sub
    If Editor = EDITOR_MAP Then Exit Sub
    If PlayerPokemon(MyIndex).Num <= 0 Then Exit Sub
    If PlayerPokemon(MyIndex).Slot <= 0 Then Exit Sub
    If Not GameState = GameStateEnum.InGame Then Exit Sub
    If GettingMap Then Exit Sub

    '//Base Location
    X = ((PlayerPokemon(MyIndex).X * TILE_X) + PlayerPokemon(MyIndex).xOffset) - ((140 / 2) - (TILE_X / 2))
    Y = ((PlayerPokemon(MyIndex).Y * TILE_Y) + PlayerPokemon(MyIndex).yOffset) - ((140 / 2) - (TILE_Y / 2))

    '//Top, Move Index 1
    '//Set Location
    '//Check Moveset
    For MoveSlot = 1 To 4
        MoveNum = PlayerPokemons(PlayerPokemon(MyIndex).Slot).Moveset(MoveSlot).Num
        If MoveNum > 0 Then
            Select Case MoveSlot
            Case 1
                mX = X + 2
                mY = Y - 11
            Case 2
                mX = X + 2
                mY = Y + 117
            Case 3
                mX = X - 107
                mY = Y + 53
            Case 4
                mX = X + 112
                mY = Y + 53
            End Select

            '//Render
            If SetAttackMove = MoveSlot Then
                guiAlpha = 255
            Else
                guiAlpha = 100
            End If

            If SetAttackMove = MoveSlot And GameSetting.ShowPP = YES Or Ctrl_Press Then

                RenderTexture Tex_Misc(Misc_MoveSelector), ConvertMapX(mX), ConvertMapY(mY), 24, 152, 136, 32, 136, 32, D3DColorARGB(guiAlpha, 255, 255, 255)
                barWidth = ((PlayerPokemons(PlayerPokemon(MyIndex).Slot).Moveset(MoveSlot).CurPP / 115) / (PlayerPokemons(PlayerPokemon(MyIndex).Slot).Moveset(MoveSlot).TotalPP / 115)) * 115
                RenderTexture Tex_Misc(Misc_MoveSelector), ConvertMapX(mX) + 9, ConvertMapY(mY), 24, 199, barWidth, 27, 115, 27, D3DColorARGB(guiAlpha, 255, 255, 255)
                RenderText Font_Default, Trim$(PokemonMove(MoveNum).Name), ConvertMapX(mX + 17), ConvertMapY(mY + 6), White, , guiAlpha
                
                 If Ctrl_Press Then
                    ' Poke Type texture
                    If PokemonMove(MoveNum).Type > 0 Then
                        RenderTexture Tex_PokemonTypesSymbol(PokemonMove(MoveNum).Type), ConvertMapX(mX - 5), ConvertMapY(mY + 6), 0, 0, 15, 15, 20, 20
                    End If

                    ' Poke Category texture
                    If PokemonMove(MoveNum).Category > 0 Then
                        RenderTexture Tex_Misc(Misc_CategoryTypes), ConvertMapX(mX + 125), ConvertMapY(mY + 6), PokemonMove(MoveNum).Category * 20 - 20, 0, 16, 16, 20, 16
                    End If
                End If
            End If
        End If
    Next

    '//Top, Move Index 1
    '//Set Location
    mX = X + 32
    mY = Y + 3
    'if setattackmove = 1
    If Ctrl_Press Then
        If UpMoveKey Then
            RenderTexture Tex_Misc(Misc_MoveSelector), ConvertMapX(mX), ConvertMapY(mY), 76, 2, 78, 47, 78, 47, D3DColorARGB(220, 255, 255, 255)
        Else
            RenderTexture Tex_Misc(Misc_MoveSelector), ConvertMapX(mX), ConvertMapY(mY), 76, 2, 78, 47, 78, 47, D3DColorARGB(100, 255, 255, 255)
        End If
    End If

    '//Bottom, Move Index 2
    '//Set Location
    mX = X + 33
    mY = Y + 93
    If Ctrl_Press Then
        If DownMoveKey Then
            RenderTexture Tex_Misc(Misc_MoveSelector), ConvertMapX(mX), ConvertMapY(mY), 77, 76, 76, 46, 76, 46, D3DColorARGB(220, 255, 255, 255)
        Else
            RenderTexture Tex_Misc(Misc_MoveSelector), ConvertMapX(mX), ConvertMapY(mY), 77, 76, 76, 46, 76, 46, D3DColorARGB(100, 255, 255, 255)
        End If
    End If

    '//Left, Move Index 3
    '//Set Location
    mX = X + 3
    mY = Y + 33
    If Ctrl_Press Then
        If LeftMoveKey Then
            RenderTexture Tex_Misc(Misc_MoveSelector), ConvertMapX(mX), ConvertMapY(mY), 2, 12, 47, 77, 47, 77, D3DColorARGB(220, 255, 255, 255)
        Else
            RenderTexture Tex_Misc(Misc_MoveSelector), ConvertMapX(mX), ConvertMapY(mY), 2, 12, 47, 77, 47, 77, D3DColorARGB(100, 255, 255, 255)
        End If
    End If

    '//Right, Move Index 4
    '//Set Location
    mX = X + 92
    mY = Y + 33
    If Ctrl_Press Then
        If RightMoveKey Then
            RenderTexture Tex_Misc(Misc_MoveSelector), ConvertMapX(mX), ConvertMapY(mY), 182, 12, 48, 77, 48, 77, D3DColorARGB(220, 255, 255, 255)
        Else
            RenderTexture Tex_Misc(Misc_MoveSelector), ConvertMapX(mX), ConvertMapY(mY), 182, 12, 48, 77, 48, 77, D3DColorARGB(100, 255, 255, 255)
        End If
    End If
End Sub

'//Weather
Private Sub DrawWeather()
Dim i As Long
Dim chc As Long
Dim Width As Long

    If Map.Sheltered > 0 Then Exit Sub

    If Weather.Type > WeatherEnum.None Then
        '//Check which weather it is
        Select Case Weather.Type
            Case WeatherEnum.Rain
                RenderTexture Tex_System(gSystemEnum.UserInterface), 0, 0, 0, 8, Screen_Width, Screen_Height, 1, 1, D3DColorARGB(70, 0, 0, 0)
                
                If Weather.InitDrop Then
                    For i = 1 To Weather.MaxDrop
                        With Weather.Drop(i)
                            If .Pic > 0 And .Pic <= Count_Weather Then
                                '//Make sure it's on the screen
                                If .X >= -32 And .X <= Screen_Width + 32 And .Y >= -32 And .Y <= Screen_Height + 32 Then
                                    RenderTexture Tex_Weather(.Pic), .X, .Y, (GetPicWidth(Tex_Weather(.Pic)) / 4) * .PicType, 0, 32, 32, 16, 16
                                End If
                            End If
                            .X = .X - 6
                            .Y = .Y + .SpeedY
                            
                            '//If out of screen, then redraw
                            If .X <= -32 Then
                                .X = Rand(0, (Screen_Width * 2))
                                .Y = Rand((-1 * Screen_Height), -32)
                                .SpeedY = 6
                                .PicType = Rand(0, 3)
                                If .PicType < 0 Then .PicType = 0
                                If .PicType > 3 Then .PicType = 3
                            End If
                            If .Y >= Screen_Height + 32 Then
                                .X = Rand(0, (Screen_Width * 2))
                                .Y = Rand((-1 * Screen_Height), -32)
                                .SpeedY = 6
                                .PicType = Rand(0, 3)
                                If .PicType < 0 Then .PicType = 0
                                If .PicType > 3 Then .PicType = 3
                            End If
                        End With
                    Next
                End If
                Exit Sub
            Case WeatherEnum.Snow
                RenderTexture Tex_System(gSystemEnum.UserInterface), 0, 0, 0, 8, Screen_Width, Screen_Height, 1, 1, D3DColorARGB(20, 0, 0, 0)
                
                If Weather.InitDrop Then
                    For i = 1 To Weather.MaxDrop
                        With Weather.Drop(i)
                            If .Pic > 0 And .Pic <= Count_Weather Then
                                '//Make sure it's on the screen
                                If .X >= -32 And .X <= Screen_Width + 32 And .Y >= -32 And .Y <= Screen_Height + 32 Then
                                    RenderTexture Tex_Weather(.Pic), .X, .Y, (GetPicWidth(Tex_Weather(.Pic)) / 4) * .PicType, 0, 32, 32, 16, 16
                                End If
                            End If
                            .X = .X '+ 1 '(Rand(-2, 2))
                            .Y = .Y + .SpeedY
                            
                            '//If out of screen, then redraw
                            If .X <= -32 Then
                                .X = Rand(0, Screen_Width)
                                .Y = Rand((-1 * Screen_Height), -32)
                                .SpeedY = Rand(1, 3)
                                .PicType = Rand(0, 3)
                                If .PicType < 0 Then .PicType = 0
                                If .PicType > 3 Then .PicType = 3
                            End If
                            If .Y >= Screen_Height + 32 Then
                                .X = Rand(0, Screen_Width)
                                .Y = Rand((-1 * Screen_Height), -32)
                                .SpeedY = Rand(1, 3)
                                .PicType = Rand(0, 3)
                                If .PicType < 0 Then .PicType = 0
                                If .PicType > 3 Then .PicType = 3
                            End If
                        End With
                    Next
                End If
                Exit Sub
            Case WeatherEnum.SandStorm
                RenderTexture Tex_System(gSystemEnum.UserInterface), 0, 0, 0, 8, Screen_Width, Screen_Height, 1, 1, D3DColorARGB(100, 210, 187, 133)
                
                If Weather.InitDrop Then
                    For i = 1 To Weather.MaxDrop
                        With Weather.Drop(i)
                            If .Pic > 0 And .Pic <= Count_Weather Then
                                '//Make sure it's on the screen
                                If .X >= -32 And .X <= Screen_Width + 32 And .Y >= -32 And .Y <= Screen_Height + 32 Then
                                    RenderTexture Tex_Weather(.Pic), .X, .Y, (GetPicWidth(Tex_Weather(.Pic)) / 4) * .PicType, 0, 32, 32, 16, 16
                                End If
                            End If
                            .X = .X + .SpeedY
                            .Y = .Y
                            
                            '//If out of screen, then redraw
                            If .X >= Screen_Width + 32 Then
                                .X = Rand((-1 * Screen_Width), -32)
                                .Y = Rand(0, Screen_Height)
                                .SpeedY = Rand(6, 9)
                                .PicType = Rand(0, 3)
                                If .PicType < 0 Then .PicType = 0
                                If .PicType > 3 Then .PicType = 3
                            End If
                        End With
                    Next
                End If
                Exit Sub
            Case WeatherEnum.Hail
                RenderTexture Tex_System(gSystemEnum.UserInterface), 0, 0, 0, 8, Screen_Width, Screen_Height, 1, 1, D3DColorARGB(70, 0, 0, 0)
                
                If Weather.InitDrop Then
                    For i = 1 To Weather.MaxDrop
                        With Weather.Drop(i)
                            If .Pic > 0 And .Pic <= Count_Weather Then
                                '//Make sure it's on the screen
                                If .X >= -32 And .X <= Screen_Width + 32 And .Y >= -32 And .Y <= Screen_Height + 32 Then
                                    RenderTexture Tex_Weather(.Pic), .X, .Y, 0, 0, 32, 32, 16, 16
                                End If
                            End If
                            .X = .X - 2
                            .Y = .Y + .SpeedY
                            
                            '//If out of screen, then redraw
                            If .X <= -32 Then
                                .X = Rand(0, (Screen_Width * 2))
                                .Y = Rand((-1 * Screen_Height), -32)
                                .SpeedY = 9
                                If .PicType < 0 Then .PicType = 0
                                If .PicType > 3 Then .PicType = 3
                            End If
                            If .Y >= Screen_Height + 32 Then
                                .X = Rand(0, (Screen_Width * 2))
                                .Y = Rand((-1 * Screen_Height), -32)
                                .SpeedY = 9
                                If .PicType < 0 Then .PicType = 0
                                If .PicType > 3 Then .PicType = 3
                            End If
                        End With
                    Next
                End If
                Exit Sub
            Case WeatherEnum.Sunny
                If Weather.InitDrop Then
                    chc = 130
                    If WeatherAlphaState = 0 Then
                        WeatherAlpha = WeatherAlpha + 1
                        If WeatherAlpha >= 30 Then
                            WeatherAlpha = 30
                            WeatherAlphaState = 1
                        End If
                    Else
                        WeatherAlpha = WeatherAlpha - 1
                        If WeatherAlpha <= 0 Then
                            WeatherAlpha = 0
                            WeatherAlphaState = 0
                        End If
                    End If
            
                    For i = 1 To 6
                        '//242/231/0
                        RenderTexture Tex_System(gSystemEnum.UserInterface), 0, (50 * (i - 1)), 0, 8, Screen_Width, 50, 1, 1, D3DColorARGB(chc + WeatherAlpha, 255, 239, 151)
                        chc = chc - 30
                        If chc <= 0 Then chc = 0
                    Next
                End If
                Exit Sub
        End Select
    End If
End Sub

' *********
' ** GUI **
' *********
Private Sub DrawAlertWindow()
Dim LowBound As Long, UpBound As Long
Dim ArrayText() As String
Dim X As Long, Y As Long
Dim yOffset As Long
Dim i As Long
Dim W As Long

    '//Loop through all items
    For i = 1 To MAX_ALERT
        With AlertWindow(i)
            '//Make sure it is being used
            If .IsUsed Then
                '//Check timer
                If GetTickCount < .AlertTimer Then
                    '//Wrap the text
                    WordWrap_Array Font_Default, .Text, ALERT_STRING_LENGTH, ArrayText
                    
                    '//we need these often
                    LowBound = LBound(ArrayText)
                    UpBound = UBound(ArrayText)
                    
                    '//Draw the hud of the text
                    X = (Screen_Width / 2) - (.Width / 2)
                    Y = .CurYPos + 3
                    '//Top Left
                    RenderTexture Tex_System(gSystemEnum.UserInterface), X - 10, Y, 8, 0, 8, 8, 8, 8
                    '//Top
                    RenderTexture Tex_System(gSystemEnum.UserInterface), X - 2, Y, 16, 0, .Width + 4, 8, 8, 8
                    '//Top Right
                    RenderTexture Tex_System(gSystemEnum.UserInterface), X + .Width + 2, Y, 24, 0, 8, 8, 8, 8
                    '//Left
                    RenderTexture Tex_System(gSystemEnum.UserInterface), X - 10, Y + 8, 8, 8, 8, .Height - 26, 8, 8
                    '//Middle
                    RenderTexture Tex_System(gSystemEnum.UserInterface), X - 2, Y + 8, 16, 8, .Width + 4, .Height - 26, 8, 8
                    '//Right
                    RenderTexture Tex_System(gSystemEnum.UserInterface), X + .Width + 2, Y + 8, 24, 8, 8, .Height - 26, 8, 8
                    '//Bottom Left
                    RenderTexture Tex_System(gSystemEnum.UserInterface), X - 10, Y + .Height - 18, 8, 16, 8, 8, 8, 8
                    '//Bottom
                    RenderTexture Tex_System(gSystemEnum.UserInterface), X - 2, Y + .Height - 18, 16, 16, .Width + 4, 8, 8, 8
                    '//Bottom Right
                    RenderTexture Tex_System(gSystemEnum.UserInterface), X + .Width + 2, Y + .Height - 18, 24, 16, 8, 8, 8, 8

                    '//Check if it wrap
                    If UpBound > LowBound Then
                        '//Reset
                        yOffset = 0
                        '//Loop to all items
                        For W = LowBound To UpBound
                            '//Set Location
                            '//Keep it centered
                            X = (Screen_Width / 2) - (GetTextWidth(Font_Default, ArrayText(W)) / 2)
                            Y = .CurYPos + 3 + yOffset
                            
                            '//Render the text
                            RenderText Font_Default, ArrayText(W), X - 2, Y + 5, .Color
                            
                            '//Increase the location for each line
                            yOffset = yOffset + 16
                        Next
                    Else
                        '//Set Location
                        '//Keep it centered
                        X = (Screen_Width / 2) - (GetTextWidth(Font_Default, .Text) / 2)
                        Y = .CurYPos + 3
                        
                        '//Render the text
                        RenderText Font_Default, .Text, X - 2, Y + 5, .Color
                    End If
                Else
                    RemoveAlert i
                    '//Clear Alert then Update position
                End If
                
                '//Update Location
                If .CurYPos > .SetYPos Then
                    .CurYPos = .CurYPos - 5
                    If .CurYPos <= .SetYPos Then .CurYPos = .SetYPos
                ElseIf .CurYPos < .SetYPos Then
                    .CurYPos = .CurYPos + 5
                    If .CurYPos >= .SetYPos Then .CurYPos = .SetYPos
                End If
            End If
        End With
    Next
End Sub


Private Sub DrawChoiceBox()
Dim i As Long

    With GUI(GuiEnum.GUI_CHOICEBOX)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        RenderTexture Tex_System(gSystemEnum.UserInterface), 0, 0, 0, 8, Screen_Width, Screen_Height, 1, 1, D3DColorARGB(160, 0, 0, 0)
        
        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height
        
        '//Buttons
        For i = ButtonEnum.ChoiceBox_Yes To ButtonEnum.ChoiceBox_No
            If CanShowButton(i) Then
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
            End If
        Next

        '//Render Text
        RenderArrayText Font_Default, ChoiceBoxText, .X + 10, .Y + 10, 250, DarkGrey
    End With
End Sub

Private Sub DrawInputBox()
Dim i As Long

    If GettingMap Then Exit Sub

    With GUI(GuiEnum.GUI_INPUTBOX)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        RenderTexture Tex_System(gSystemEnum.UserInterface), 0, 0, 0, 8, Screen_Width, Screen_Height, 1, 1, D3DColorARGB(160, 0, 0, 0)
        
        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height
                
        '//Buttons
        For i = ButtonEnum.InputBox_Okay To ButtonEnum.InputBox_Cancel
            If CanShowButton(i) Then
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
            End If
        Next

        '//Render Text
        RenderArrayText Font_Default, InputBoxHeader, .X + 10, .Y + 10, 250, White
        
        '//Text
        RenderArrayText Font_Default, UpdateChatText(Font_Default, InputBoxText & TextLine, 210), .X + 22, .Y + 34, 250, White
    End With
End Sub

'//Hud
Private Sub DrawHud()
    Dim X As Long, Y As Long
    Dim i As Byte
    Dim Alpha As Byte
    Dim bWidth As Long
    Dim expWidth As Long
    Dim initAnim As Byte
    Dim Sprite As Long
    Dim Addy As Long

    If Editor = EDITOR_MAP Then Exit Sub

    IsHovering = False

    For i = 1 To MAX_PLAYER_POKEMON
        If PlayerPokemons(i).Num > 0 Then
            X = Screen_Width - 34 - 5 - ((i - 1) * 40)
            Y = 62    '+ 52 + ((i - 1) * 40)

            If PlayerPokemons(i).CurHP <= 0 Then
                Alpha = 150
                initAnim = 0
            Else
                Alpha = 255
                initAnim = MapAnim
                If PlayerPokemon(MyIndex).Num > 0 Then
                    If PlayerPokemon(MyIndex).Slot <> i Then
                        Alpha = 150
                    Else
                        If CursorX >= X And CursorX <= X + 34 And CursorY >= Y And CursorY <= Y + 37 Then
                            IsHovering = True
                            MouseIcon = 1    '//Select
                        End If
                    End If
                Else
                    If CursorX >= X And CursorX <= X + 34 And CursorY >= Y And CursorY <= Y + 37 Then
                        IsHovering = True
                        MouseIcon = 1    '//Select
                    End If
                End If
            End If

            '//Draw box
            RenderTexture Tex_Gui(Hud), X, Y, 203, 38, 34, 37, 34, 37, D3DColorARGB(Alpha, 255, 255, 255)

            '//Icon
            If Pokemon(PlayerPokemons(i).Num).Sprite > 0 And Pokemon(PlayerPokemons(i).Num).Sprite <= Count_PokemonIcon Then
                RenderTexture Tex_PokemonIcon(Pokemon(PlayerPokemons(i).Num).Sprite), X + 1, Y + 1, initAnim * 32, 0, 32, 32, 32, 32, D3DColorARGB(Alpha, 255, 255, 255)
                '//Poke Using item texture
                If PlayerPokemons(i).HeldItem > 0 And PlayerPokemons(i).HeldItem <= MAX_ITEM Then
                    RenderTexture Tex_Item(PokeUseHeld), X - 2, Y - 2, 0, 0, 14, 14, 24, 24, D3DColorARGB(Alpha, 255, 255, 255)
                End If
                '//Poke Type texture
                If Pokemon(PlayerPokemons(i).Num).PrimaryType > 0 Then
                    RenderTexture Tex_PokemonTypesSymbol(Pokemon(PlayerPokemons(i).Num).PrimaryType), X + 2, Y + 30, 0, 0, 15, 15, 20, 20, D3DColorARGB(Alpha, 255, 255, 255)
                    
                    If Pokemon(PlayerPokemons(i).Num).SecondaryType > 0 Then
                        RenderTexture Tex_PokemonTypesSymbol(Pokemon(PlayerPokemons(i).Num).SecondaryType), X + 17, Y + 30, 0, 0, 15, 15, 20, 20, D3DColorARGB(Alpha, 255, 255, 255)
                    End If
                End If
            End If
        End If
    Next

    Alpha = 100
    For i = 1 To MAX_HOTBAR
        X = Screen_Width - 42 - 170 - ((i - 1) * 45)
        Y = 5    '62 + 37 + 5
        RenderTexture Tex_Gui(Hud), X, Y, 5, 204, 42, 45, 42, 45, D3DColorARGB(Alpha, 255, 255, 255)

        If Player(MyIndex).Hotbar(i) > 0 Then
            '//Draw Icon
            Sprite = Item(Player(MyIndex).Hotbar(i)).Sprite

            If Sprite > 0 And Sprite <= Count_Item Then
                RenderTexture Tex_Item(Sprite), X + 9, Y + 9, 0, 0, 24, 24, 24, 24
            End If
        End If

        '//Key Preview
        RenderText Font_Default, GetKeyCodeName(ControlKey(ControlEnum.KeyHotbarSlot1 + (i - 1)).cAsciiKey), X + 5, Y + 25, White
    Next

    '//Time Stamp
    RenderTexture Tex_Gui(Hud), Screen_Width - 161 - 5, 5, 44, 134, 161, 52, 161, 52
    '//Map Name
    RenderText Ui_Default, Trim$(Map.Name), Screen_Width - 161 - 5 + 5, 8, White
    '//Server Time
    RenderText Ui_Default, KeepTwoDigit(GameHour) & ":" & KeepTwoDigit(GameMinute) & ":" & KeepTwoDigit(GameSecond), Screen_Width - 161 - 5 + 5, 16 + 8, White

    '//Icon
    If GameHour >= 5 And GameHour <= 11 Then
        '//Morning
        RenderTexture Tex_Gui(Hud), Screen_Width - 161 - 5 + 115, 5 + 2, 212, 173, 44, 44, 44, 44
    ElseIf GameHour >= 12 And GameHour <= 19 Then
        '//Day
        RenderTexture Tex_Gui(Hud), Screen_Width - 161 - 5 + 115, 5 + 2, 212, 129, 44, 44, 44, 44
    Else
        '//Night
        RenderTexture Tex_Gui(Hud), Screen_Width - 161 - 5 + 115, 5 + 2, 212, 85, 44, 44, 44, 44
    End If

    '//Pokemon Vital
    If PlayerPokemon(MyIndex).Num > 0 Then
        '//Draw Window
        RenderTexture Tex_Gui(Hud), 5, 5, 0, 3, 171, 82, 171, 82

        '//Icon
        If Pokemon(PlayerPokemon(MyIndex).Num).Sprite > 0 And Pokemon(PlayerPokemon(MyIndex).Num).Sprite < Count_PokemonIcon Then
            RenderTexture Tex_PokemonIcon(Pokemon(PlayerPokemon(MyIndex).Num).Sprite), 6, 3, MapAnim * 32, 0, 32, 32, 32, 32
        End If

        '//Name
        RenderText Font_Default, Trim$(Pokemon(PlayerPokemon(MyIndex).Num).Name), 48, 25, White

        '//Level
        If PlayerPokemon(MyIndex).Slot > 0 Then
            RenderText Font_Default, "Lv" & (PlayerPokemons(PlayerPokemon(MyIndex).Slot).Level), 135, 25, White
        End If

        '//HP
        If PlayerPokemons(PlayerPokemon(MyIndex).Slot).MaxHP > 0 Then
            bWidth = ((PlayerPokemons(PlayerPokemon(MyIndex).Slot).CurHP / 135) / (PlayerPokemons(PlayerPokemon(MyIndex).Slot).MaxHP / 135)) * 135
            RenderTexture Tex_Gui(Hud), 5 + 25, 5 + 44, 7, 97, bWidth, 13, 135, 13, D3DColorARGB(255, 31, 161, 69)
        End If

        '//Exp
        If PlayerPokemons(PlayerPokemon(MyIndex).Slot).NextExp > 0 Then
            bWidth = ((PlayerPokemons(PlayerPokemon(MyIndex).Slot).CurExp / 142) / (PlayerPokemons(PlayerPokemon(MyIndex).Slot).NextExp / 142)) * 142
            expWidth = 142 - bWidth
            RenderTexture Tex_Gui(Hud), 5 + 18, 5 + 60, 7, 111, 142 - expWidth, 7, 142 - expWidth, 7
        End If
    End If

    '//Party
    If InParty > 0 Then
        '//Party Member
        '//Render your name first
        RenderTexture Tex_Gui(Hud), 0, 159, 59, 241, 165, 20, 165, 1
        RenderText Font_Default, "Party Member", 10, 160, Yellow
        Addy = 21
        For i = 1 To MAX_PARTY
            If Len(Trim$(PartyName(i))) > 0 Then
                RenderTexture Tex_Gui(Hud), 0, 159 + Addy, 59, 241, 165, 20, 165, 1, D3DColorARGB(150, 255, 255, 255)
                RenderText Font_Default, Trim$(PartyName(i)), 10, 160 + Addy, White
                Addy = Addy + 21
            End If
        Next
    End If
End Sub

Private Sub DrawMoveReplace()
Dim i As Long
Dim moveText As String
Dim MoveSlot As Byte
Dim MoveNum As Long

    With GUI(GuiEnum.GUI_MOVEREPLACE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        If MoveLearnNum <= 0 Then Exit Sub
        If MoveLearnPokeSlot <= 0 Then Exit Sub
        If PlayerPokemons(MoveLearnPokeSlot).Num <= 0 Then Exit Sub
        
        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height
        
        '//Buttons
        Dim ButtonText As String, DrawText As Boolean
        For i = ButtonEnum.MoveReplace_Slot1 To ButtonEnum.MoveReplace_Cancel
            If CanShowButton(i) Then
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
                
                '//Render Button Text
                Select Case i
                    Case ButtonEnum.MoveReplace_Slot1 To ButtonEnum.MoveReplace_Slot4
                        MoveSlot = i - (ButtonEnum.MoveReplace_Slot1 - 1)
                        If MoveSlot > 0 Then
                            MoveNum = PlayerPokemons(MoveLearnPokeSlot).Moveset(MoveSlot).Num
                            If MoveNum > 0 Then
                                ButtonText = Trim$(PokemonMove(MoveNum).Name)
                                DrawText = True
                            End If
                        End If
                    Case ButtonEnum.MoveReplace_Cancel: ButtonText = "Cancel": DrawText = True
                    Case Else: DrawText = False
                End Select
                If DrawText Then
                    Select Case Button(i).State
                        Case ButtonState.StateNormal: RenderText Ui_Default, ButtonText, (.X + Button(i).X) + 5, (.Y + Button(i).Y) + ((Button(i).Height / 2) - (8)) - 5, D3DColorARGB(255, 229, 229, 229), False
                        Case ButtonState.StateHover: RenderText Ui_Default, ButtonText, (.X + Button(i).X) + 5, (.Y + Button(i).Y) + ((Button(i).Height / 2) - (8)) - 5, D3DColorARGB(255, 255, 255, 255), False
                        Case ButtonState.StateClick: RenderText Ui_Default, ButtonText, (.X + Button(i).X) + 5, (.Y + Button(i).Y) + ((Button(i).Height / 2) - (8)) - 5 + 3, D3DColorARGB(255, 255, 255, 255), False
                    End Select
                End If
            End If
        Next
        
        '//Draw Text
        moveText = Trim$(Pokemon(PlayerPokemons(MoveLearnPokeSlot).Num).Name) & " is trying to learn " & Trim$(PokemonMove(MoveLearnNum).Name) & ", Select a move to replace for this move"
        RenderArrayText Font_Default, moveText, .X + 16, .Y + 11, 200, White
    End With
End Sub

Private Sub DrawTrainer()
Dim i As Long, YPos As Long

    With GUI(GuiEnum.GUI_TRAINER)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height
        
        '//Buttons
        For i = ButtonEnum.Trainer_Close To ButtonEnum.Trainer_Badge
            If CanShowButton(i) Then
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
            End If
        Next
        
        '//Name
        YPos = .Y + 70
        RenderText Font_Default, Trim$(Player(MyIndex).Name), .X + 85, YPos, White
        '//Level
        YPos = YPos + 23
        RenderText Font_Default, "Lv " & Trim$(Player(MyIndex).Level), .X + 85 + (15 - (GetTextWidth(Font_Default, "Lv " & Trim$(Player(MyIndex).Level)) / 2)), YPos, White
        '//Exp
        YPos = YPos + 23
        RenderText Font_Default, Player(MyIndex).CurExp & "/" & GetLevelNextExp(Player(MyIndex).Level), .X + 85, YPos, White
        '//Money
        YPos = YPos + 23
        RenderTexture Tex_Item(IDMoney), .X + 60, YPos, 0, 0, 17, 17, 24, 24
        RenderText Font_Default, (FormatarValor(Player(MyIndex).Money)), .X + 85, YPos, White
        '//Cash
        YPos = YPos + 23
        RenderTexture Tex_Item(IDCash), .X + 60, YPos, 0, 0, 17, 17, 24, 24
        RenderText Font_Default, (FormatarValor(Player(MyIndex).Cash)), .X + 85, YPos, White
        '//Jornada Iniciada
        YPos = YPos + 23
        RenderText Font_Default, (Player(MyIndex).Started), .X + 85, YPos, White
        '//Tempo Jogado
        YPos = YPos + 23
        RenderText Font_Default, SecondsToHMS(Player(MyIndex).TimePlay), .X + 85, YPos, White
        
        '//PvP
        YPos = YPos + 53
        RenderText Font_Default, (Player(MyIndex).win), .X + 85, YPos, White
        YPos = YPos + 24
        RenderText Font_Default, (Player(MyIndex).Lose), .X + 85, YPos, White
        YPos = YPos + 24
        RenderText Font_Default, (Player(MyIndex).Tie), .X + 85, YPos, White
    End With
End Sub

Function FormatarValor(ByVal valor As Long) As String
    If Len(CStr(valor)) > 3 Then
        FormatarValor = Format(valor, "0,###")
    Else
        FormatarValor = CStr(valor)
    End If
End Function

Private Sub DrawInvStorage()
Dim i As Long
Dim slotNum As Long
Dim Sprite As Long
Dim DrawX As Long, DrawY As Long

    With GUI(GuiEnum.GUI_INVSTORAGE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height
        
        '//Buttons
        Dim ButtonText As String, DrawText As Boolean
        For i = ButtonEnum.InvStorage_Close To ButtonEnum.InvStorage_Slot5
            If CanShowButton(i) Then
                slotNum = ((i + 1) - ButtonEnum.InvStorage_Slot1)
                If InvCurSlot = slotNum Then
                    Button(i).State = ButtonState.StateClick
                End If

                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
            
                '//Render Button Text
                Select Case i
                    Case ButtonEnum.InvStorage_Slot1 To ButtonEnum.InvStorage_Slot5
                        If PlayerInvStorage(slotNum).Unlocked = YES Then
                            ButtonText = "Slot " & slotNum
                        Else
                            ButtonText = "Locked"
                        End If
                        DrawText = True
                    Case Else: DrawText = False
                End Select
                If DrawText Then
                    Select Case Button(i).State
                        Case ButtonState.StateNormal: RenderText Ui_Default, ButtonText, (.X + Button(i).X) + 5, (.Y + Button(i).Y) + ((Button(i).Height / 2) - (8)) - 5, D3DColorARGB(255, 229, 229, 229), False
                        Case ButtonState.StateHover: RenderText Ui_Default, ButtonText, (.X + Button(i).X) + 5, (.Y + Button(i).Y) + ((Button(i).Height / 2) - (8)) - 5, D3DColorARGB(255, 255, 255, 255), False
                        Case ButtonState.StateClick: RenderText Ui_Default, ButtonText, (.X + Button(i).X) + 5, (.Y + Button(i).Y) + ((Button(i).Height / 2) - (8)) - 5 + 3, D3DColorARGB(255, 255, 255, 255), False
                    End Select
                End If
            End If
        Next
        
        '//Items
        For i = 1 To MAX_STORAGE
            If i <> DragStorageSlot Then
                If PlayerInvStorage(InvCurSlot).Data(i).Num > 0 Then
                    Sprite = Item(PlayerInvStorage(InvCurSlot).Data(i).Num).Sprite
                    
                    DrawX = .X + (98 + ((5 + TILE_X) * (((i - 1) Mod 7))))
                    DrawY = .Y + (37 + ((5 + TILE_Y) * ((i - 1) \ 7)))
                    
                    '//Draw Icon
                    If Sprite > 0 And Sprite <= Count_Item Then
                        RenderTexture Tex_Item(Sprite), DrawX + ((32 / 2) - (GetPicWidth(Tex_Item(Sprite)) / 2)), DrawY + ((32 / 2) - (GetPicHeight(Tex_Item(Sprite)) / 2)), 0, 0, GetPicWidth(Tex_Item(Sprite)), GetPicHeight(Tex_Item(Sprite)), GetPicWidth(Tex_Item(Sprite)), GetPicHeight(Tex_Item(Sprite))
                    End If
                    
                    RenderTexture Tex_System(gSystemEnum.UserInterface), DrawX, DrawY, 0, 8, TILE_X, TILE_Y, 1, 1, D3DColorARGB(20, 0, 0, 0)
                    
                    '//Count
                    If PlayerInvStorage(InvCurSlot).Data(i).value > 1 Then
                        RenderText Font_Default, PlayerInvStorage(InvCurSlot).Data(i).value, DrawX + 28 - (GetTextWidth(Font_Default, PlayerInvStorage(InvCurSlot).Data(i).value)), DrawY + 14, White
                    End If
                End If
            End If
        Next
        
        '//Title
        RenderText Ui_Default, "Item Storage", .X + 10, .Y + 4, D3DColorARGB(180, 255, 255, 255), False
    End With
End Sub

Private Sub DrawPokemonStorage()
    Dim i As Long
    Dim slotNum As Long
    Dim Sprite As Long
    Dim DrawX As Long, DrawY As Long

    With GUI(GuiEnum.GUI_POKEMONSTORAGE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub

        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height

        '//Buttons
        Dim ButtonText As String, DrawText As Boolean
        For i = ButtonEnum.PokemonStorage_Close To ButtonEnum.PokemonStorage_Slot5
            If CanShowButton(i) Then
                slotNum = ((i + 1) - ButtonEnum.PokemonStorage_Slot1)
                If PokemonCurSlot = slotNum Then
                    Button(i).State = ButtonState.StateClick
                End If

                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height

                '//Render Button Text
                Select Case i
                Case ButtonEnum.PokemonStorage_Slot1 To ButtonEnum.PokemonStorage_Slot5
                    If PlayerPokemonStorage(slotNum).Unlocked = YES Then
                        ButtonText = "Slot " & slotNum
                    Else
                        ButtonText = "Locked"
                    End If
                    DrawText = True
                Case Else: DrawText = False
                End Select
                If DrawText Then
                    Select Case Button(i).State
                    Case ButtonState.StateNormal: RenderText Ui_Default, ButtonText, (.X + Button(i).X) + 5, (.Y + Button(i).Y) + ((Button(i).Height / 2) - (8)) - 5, D3DColorARGB(255, 229, 229, 229), False
                    Case ButtonState.StateHover: RenderText Ui_Default, ButtonText, (.X + Button(i).X) + 5, (.Y + Button(i).Y) + ((Button(i).Height / 2) - (8)) - 5, D3DColorARGB(255, 255, 255, 255), False
                    Case ButtonState.StateClick: RenderText Ui_Default, ButtonText, (.X + Button(i).X) + 5, (.Y + Button(i).Y) + ((Button(i).Height / 2) - (8)) - 5 + 3, D3DColorARGB(255, 255, 255, 255), False
                    End Select
                End If
            End If
        Next

        '//Pokemon
        For i = 1 To MAX_STORAGE
            If i <> DragPokeSlot Then

                If PlayerPokemonStorage(PokemonCurSlot).Data(i).Num > 0 Then
                    Sprite = Pokemon(PlayerPokemonStorage(PokemonCurSlot).Data(i).Num).Sprite

                    DrawX = .X + (98 + ((5 + TILE_X) * (((i - 1) Mod 7))))
                    DrawY = .Y + (37 + ((5 + TILE_Y) * ((i - 1) \ 7)))

                    '//Icon
                    If Sprite > 0 And Sprite < Count_PokemonIcon Then
                        RenderTexture Tex_PokemonIcon(Sprite), DrawX, DrawY, MapAnim * 32, 0, 32, 32, 32, 32
                        '//Held Item
                        If PlayerPokemonStorage(PokemonCurSlot).Data(i).HeldItem > 0 And PlayerPokemonStorage(PokemonCurSlot).Data(i).HeldItem <= MAX_ITEM Then
                            RenderTexture Tex_Item(PokeUseHeld), DrawX + 20, DrawY - 2, 0, 0, 14, 14, 24, 24
                        End If

                        '//Poke Type texture
                        'If Pokemon(PlayerPokemonStorage(PokemonCurSlot).Data(i).Num).PrimaryType > 0 Then
                        '    RenderTexture Tex_PokemonTypes(Pokemon(PlayerPokemonStorage(PokemonCurSlot).Data(i).Num).PrimaryType), DrawX + 2, DrawY + 30, 0, 0, 14, 14, 22, 23
                        '    If Pokemon(PlayerPokemonStorage(PokemonCurSlot).Data(i).Num).SecondaryType > 0 Then
                        '        RenderTexture Tex_PokemonTypes(Pokemon(PlayerPokemonStorage(PokemonCurSlot).Data(i).Num).SecondaryType), DrawX + 17, DrawY + 30, 0, 0, 14, 14, 22, 23
                        '    End If
                        'End If
                    End If

                    RenderTexture Tex_System(gSystemEnum.UserInterface), DrawX, DrawY, 0, 8, TILE_X, TILE_Y, 1, 1, D3DColorARGB(20, 0, 0, 0)
                End If
            End If

            ' Using in PokeStorage to select Pokes.
            If IsPokemonSelected(i) Then
                RenderTexture Tex_Misc(Misc_PokeSelect), GetPokemonSelectedX(i), GetPokemonSelectedY(i) - 7, 0, 0, 26, 20, 26, 20
            End If
        Next

        '//Title
        RenderText Ui_Default, "Pokemon Storage", .X + 10, .Y + 4, D3DColorARGB(180, 255, 255, 255), False
    End With
End Sub

Private Sub DrawDragIcon()
    Dim Sprite As Long

    If DragInvSlot > 0 Then
        If PlayerInv(DragInvSlot).Num > 0 Then
            Sprite = Item(PlayerInv(DragInvSlot).Num).Sprite

            '//Draw Icon
            If Sprite > 0 And Sprite <= Count_Item Then
                RenderTexture Tex_Item(Sprite), CursorX - (GetPicWidth(Tex_Item(Sprite)) / 2), CursorY - (GetPicHeight(Tex_Item(Sprite)) / 2), 0, 0, GetPicWidth(Tex_Item(Sprite)), GetPicHeight(Tex_Item(Sprite)), GetPicWidth(Tex_Item(Sprite)), GetPicHeight(Tex_Item(Sprite))
            End If
        End If
    End If
    If DragStorageSlot > 0 Then
        If PlayerInvStorage(InvCurSlot).Data(DragStorageSlot).Num > 0 Then
            Sprite = Item(PlayerInvStorage(InvCurSlot).Data(DragStorageSlot).Num).Sprite

            '//Draw Icon
            If Sprite > 0 And Sprite <= Count_Item Then
                RenderTexture Tex_Item(Sprite), CursorX - (GetPicWidth(Tex_Item(Sprite)) / 2), CursorY - (GetPicHeight(Tex_Item(Sprite)) / 2), 0, 0, GetPicWidth(Tex_Item(Sprite)), GetPicHeight(Tex_Item(Sprite)), GetPicWidth(Tex_Item(Sprite)), GetPicHeight(Tex_Item(Sprite))
            End If
        End If
    End If
    If DragPokeSlot > 0 Then
        If PlayerPokemonStorage(PokemonCurSlot).Data(DragPokeSlot).Num > 0 Then
            Sprite = Pokemon(PlayerPokemonStorage(PokemonCurSlot).Data(DragPokeSlot).Num).Sprite

            '//Draw Icon
            If Sprite > 0 And Sprite <= Count_PokemonIcon Then
                RenderTexture Tex_PokemonIcon(Sprite), CursorX - (TILE_X / 2), CursorY - (TILE_Y / 2), MapAnim * TILE_X, 0, TILE_X, TILE_Y, TILE_X, TILE_Y
            End If
        End If
    End If
End Sub



Private Sub DrawConvo()
Dim i As Long
Dim Sprite As Long
Dim spriteWidth As Long, spriteHeight As Long
Dim scaleWidth As Long, scaleHeight As Long

    With GUI(GuiEnum.GUI_CONVO)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Render black alpha
        RenderTexture Tex_System(gSystemEnum.UserInterface), 0, 0, 0, 8, Screen_Width, Screen_Height, 1, 1, D3DColorARGB(150, 0, 0, 0)
        
        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height
        
        '//Error
        If ConvoNum <= 0 Or ConvoData <= 0 Then Exit Sub
        
        '//Render Sprite
        If ConvoNpcNum > 0 Then
            Sprite = Npc(ConvoNpcNum).Sprite
            If Sprite > 0 Then
                spriteWidth = GetPicWidth(Tex_Character(Sprite)) / 3
                spriteHeight = GetPicHeight(Tex_Character(Sprite)) / 4
                scaleWidth = spriteWidth * 6
                scaleHeight = spriteHeight * 6
                RenderTexture Tex_Character(Sprite), .X + ((.Width / 2) - (scaleWidth / 2)), .Y - scaleHeight + 10, spriteWidth, 0, scaleWidth, scaleHeight, spriteWidth, spriteHeight
            End If
        End If
        
        '//Text
        RenderArrayText Font_Default, ConvoRenderText, .X + 25, .Y + 25, 400, White, , True
        
        If Len(ConvoText) > ConvoDrawTextLen Then
            RenderTexture Tex_System(gSystemEnum.CursorLoad), .X + 425, .Y + 115, 15 * CursorLoadAnim, 0, 15, 15, 15, 15
        End If
        
        '//Convo Reply
        If ConvoShowButton Then
            '//Render black alpha
            RenderTexture Tex_System(gSystemEnum.UserInterface), 0, 0, 0, 8, Screen_Width, Screen_Height, 1, 1, D3DColorARGB(150, 0, 0, 0)
        
            '//Buttons
            Dim ButtonText As String, DrawText As Boolean
            For i = ButtonEnum.Convo_Reply1 To ButtonEnum.Convo_Reply3
                If CanShowButton(i) Then
                    RenderTexture Tex_Gui(.Pic), Button(i).X, Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
                
                    '//Render Button Text
                    Select Case i
                        Case ButtonEnum.Convo_Reply1 To ButtonEnum.Convo_Reply3
                            If Len(Trim$(ConvoReply((i + 1) - ButtonEnum.Convo_Reply1))) > 0 Then
                                ButtonText = ((i + 1) - ButtonEnum.Convo_Reply1) & ": " & Trim$(ConvoReply((i + 1) - ButtonEnum.Convo_Reply1))
                                DrawText = True
                            End If
                        Case Else: DrawText = False
                    End Select
                    If DrawText Then
                        Select Case Button(i).State
                            Case ButtonState.StateNormal: RenderText Ui_Default, ButtonText, (Button(i).X) + 5, (Button(i).Y) + ((Button(i).Height / 2) - (8)) - 5, D3DColorARGB(255, 229, 229, 229), False
                            Case ButtonState.StateHover: RenderText Ui_Default, ButtonText, (Button(i).X) + 5, (Button(i).Y) + ((Button(i).Height / 2) - (8)) - 5, D3DColorARGB(255, 255, 255, 255), False
                            Case ButtonState.StateClick: RenderText Ui_Default, ButtonText, (Button(i).X) + 5, (Button(i).Y) + ((Button(i).Height / 2) - (8)) - 5 + 3, D3DColorARGB(255, 255, 255, 255), False
                        End Select
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub DrawShop()
Dim i As Long
Dim DrawX As Long, DrawY As Long
Dim Sprite As Long
Dim pricetext As String

    With GUI(GuiEnum.GUI_SHOP)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height ', D3DColorRGBA(255, 255, 255, 255)

        '//Buttons
        For i = ButtonEnum.Shop_Close To ButtonEnum.Shop_ScrollDown
            If CanShowButton(i) Then
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
            End If
        Next
        
        '//Items
        For i = ShopAddY To ShopAddY + 8
            If i > 0 And i <= MAX_SHOP_ITEM Then
                DrawX = (31 + ((4 + 127) * (((((i + 1) - ShopAddY) - 1) Mod 3))))
                DrawY = (42 + ((4 + 78) * ((((i + 1) - ShopAddY) - 1) \ 3)))
                    
                '//Check if item exist
                If Shop(ShopNum).ShopItem(i).Num > 0 Then
                    RenderTexture Tex_Gui(.Pic), .X + DrawX, .Y + DrawY, 194, 348, 127, 78, 127, 78
                    
                    '//Render icon
                    Sprite = Item(Shop(ShopNum).ShopItem(i).Num).Sprite
                    If Sprite > 0 And Sprite <= Count_Item Then
                        DrawX = DrawX
                        DrawY = DrawY
                        RenderTexture Tex_Item(Sprite), .X + DrawX + 9 + ((32 / 2) - (GetPicWidth(Tex_Item(Sprite)) / 2)), .Y + DrawY + 6 + ((32 / 2) - (GetPicHeight(Tex_Item(Sprite)) / 2)), 0, 0, GetPicWidth(Tex_Item(Sprite)), GetPicHeight(Tex_Item(Sprite)), GetPicWidth(Tex_Item(Sprite)), GetPicHeight(Tex_Item(Sprite))
                    End If
                    
                    '//Price
                    '//ToDo: Convert, 1k, 1m , etc.
                    pricetext = Item(Shop(ShopNum).ShopItem(i).Num).Price
                    Dim IDValue As Integer
                    If Item(Shop(ShopNum).ShopItem(i).Num).IsCash = YES Then
                        IDValue = IDCash
                    Else
                        IDValue = IDMoney
                    End If
                    
                    
                    '//Button
                    If ShopButtonHover = i Then
                        If ShopButtonState = 1 Then '//Hover
                            RenderTexture Tex_Gui(.Pic), .X + DrawX + 12, .Y + DrawY + 44, 33, 375, 103, 25, 103, 25
                            RenderText Font_Default, pricetext, (.X + DrawX + 12) + ((103 / 2) - (GetTextWidth(Font_Default, pricetext) / 2)), (.Y + DrawY + 44) + 1, D3DColorARGB(255, 150, 150, 255), False
                        ElseIf ShopButtonState = 2 Then '//Click
                            RenderTexture Tex_Gui(.Pic), .X + DrawX + 12, .Y + DrawY + 44, 33, 400, 103, 25, 103, 25
                            RenderText Font_Default, pricetext, (.X + DrawX + 12) + ((103 / 2) - (GetTextWidth(Font_Default, pricetext) / 2)), (.Y + DrawY + 44) + 3, White
                        Else
                            RenderTexture Tex_Gui(.Pic), .X + DrawX + 12, .Y + DrawY + 44, 33, 350, 103, 25, 103, 25 '//Normal
                            RenderText Font_Default, pricetext, (.X + DrawX + 12) + ((103 / 2) - (GetTextWidth(Font_Default, pricetext) / 2)), (.Y + DrawY + 44) + 1, White
                        End If
                    Else
                        RenderTexture Tex_Gui(.Pic), .X + DrawX + 12, .Y + DrawY + 44, 33, 350, 103, 25, 103, 25 '//Normal
                        RenderText Font_Default, pricetext, (.X + DrawX + 12) + ((103 / 2) - (GetTextWidth(Font_Default, pricetext) / 2)), (.Y + DrawY + 44) + 1, D3DColorRGBA(100, 100, 100, 255), False
                    End If
                    
                    '//Item Name
                    RenderText Font_Default, Trim$(Item(Shop(ShopNum).ShopItem(i).Num).Name), .X + DrawX + 44, .Y + DrawY + 10, D3DColorRGBA(100, 100, 100, 255), False
                     ' Render Money or Cash Icon
                    RenderTexture Tex_Item(IDValue), (.X + DrawX) + ((70 / 2) - (GetTextWidth(Font_Default, pricetext) / 2)), (.Y + DrawY + 44) + 1, 0, 0, 20, 20, 24, 24
                End If
            End If
        Next
        
        '//Title
        RenderText Ui_Default, "Shop", .X + 10, .Y + 4, D3DColorARGB(180, 255, 255, 255), False
    End With
End Sub

Private Sub DrawTrade()
Dim i As Long
Dim DrawX As Long, DrawY As Long
Dim Sprite As Long
Dim currencyText As String

    With GUI(GuiEnum.GUI_TRADE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height
        
        '//Buttons
        Dim ButtonText As String, DrawText As Boolean
        For i = ButtonEnum.Trade_Close To ButtonEnum.Trade_AddMoney
            If CanShowButton(i) Then
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
                
                '//Render Button Text
                Select Case i
                    Case ButtonEnum.Trade_Set
                        If YourTrade.TradeSet = YES Then
                            ButtonText = "Cancel"
                        Else
                            ButtonText = "Set"
                        End If
                        DrawText = True
                    Case Else: DrawText = False
                End Select
                If DrawText Then
                    Select Case Button(i).State
                        Case ButtonState.StateNormal: RenderText Font_Default, ButtonText, (.X + Button(i).X) + ((Button(i).Width / 2) - (GetTextWidth(Font_Default, ButtonText) / 2)) - 2, (.Y + Button(i).Y) + ((Button(i).Height / 2) - (8)) - 4, D3DColorARGB(255, 100, 100, 100), False
                        Case ButtonState.StateHover: RenderText Font_Default, ButtonText, (.X + Button(i).X) + ((Button(i).Width / 2) - (GetTextWidth(Font_Default, ButtonText) / 2)) - 2, (.Y + Button(i).Y) + ((Button(i).Height / 2) - (8)) - 4, D3DColorARGB(255, 100, 200, 100), False
                        Case ButtonState.StateClick: RenderText Font_Default, ButtonText, (.X + Button(i).X) + ((Button(i).Width / 2) - (GetTextWidth(Font_Default, ButtonText) / 2)) - 2, (.Y + Button(i).Y) + ((Button(i).Height / 2) - (8)) - 4 + 3, D3DColorARGB(255, 255, 255, 255), False
                    End Select
                End If
            End If
        Next
        
        '//Trade Items
        For i = 1 To MAX_TRADE
            '//Your Trade
            If YourTrade.Data(i).TradeType > 0 Then
                DrawX = .X + (12 + ((3 + 44) * ((i - 1) Mod 4)))
                DrawY = .Y + (71 + ((3 + 46) * ((i - 1) \ 4)))
                    
                If YourTrade.Data(i).Num > 0 Then
                    RenderTexture Tex_Gui(.Pic), DrawX, DrawY, 459, 395, 44, 46, 44, 46
                        
                    '//Icon
                    If YourTrade.Data(i).TradeType = 1 Then  '//Item
                        Sprite = Item(YourTrade.Data(i).Num).Sprite
                        If Sprite > 0 And Sprite <= Count_Item Then
                            RenderTexture Tex_Item(Sprite), DrawX + 7 + ((32 / 2) - (GetPicWidth(Tex_Item(Sprite)) / 2)), DrawY + 7 + ((32 / 2) - (GetPicHeight(Tex_Item(Sprite)) / 2)), 0, 0, GetPicWidth(Tex_Item(Sprite)), GetPicHeight(Tex_Item(Sprite)), GetPicWidth(Tex_Item(Sprite)), GetPicHeight(Tex_Item(Sprite))
                        End If
                        
                        '//Count
                        If YourTrade.Data(i).value > 1 Then
                            RenderText Font_Default, YourTrade.Data(i).value, DrawX + 7 + 28 - (GetTextWidth(Font_Default, YourTrade.Data(i).value)), DrawY + 7 + 14, White
                        End If
                    ElseIf YourTrade.Data(i).TradeType = 2 Then  '//Pokemon
                        Sprite = Pokemon(YourTrade.Data(i).Num).Sprite
                        If Sprite > 0 And Sprite <= Count_PokemonIcon Then
                            RenderTexture Tex_PokemonIcon(Sprite), DrawX + 7, DrawY + 7, MapAnim * TILE_X, 0, TILE_X, TILE_Y, TILE_X, TILE_Y
                        End If
                    End If
                End If
            End If
            
            '//Their Trade
            If TheirTrade.Data(i).TradeType > 0 Then
                DrawX = .X + (222 + ((3 + 44) * ((i - 1) Mod 4)))
                DrawY = .Y + (71 + ((3 + 46) * ((i - 1) \ 4)))
                    
                If TheirTrade.Data(i).Num > 0 Then
                    RenderTexture Tex_Gui(.Pic), DrawX, DrawY, 459, 395, 44, 46, 44, 46
                        
                    '//Icon
                    If TheirTrade.Data(i).TradeType = 1 Then '//Item
                        Sprite = Item(TheirTrade.Data(i).Num).Sprite
                        If Sprite > 0 And Sprite <= Count_Item Then
                            RenderTexture Tex_Item(Sprite), DrawX + 7 + ((32 / 2) - (GetPicWidth(Tex_Item(Sprite)) / 2)), DrawY + 7 + ((32 / 2) - (GetPicHeight(Tex_Item(Sprite)) / 2)), 0, 0, GetPicWidth(Tex_Item(Sprite)), GetPicHeight(Tex_Item(Sprite)), GetPicWidth(Tex_Item(Sprite)), GetPicHeight(Tex_Item(Sprite))
                        End If
                        
                        '//Count
                        If TheirTrade.Data(i).value > 1 Then
                            RenderText Font_Default, TheirTrade.Data(i).value, DrawX + 7 + 28 - (GetTextWidth(Font_Default, TheirTrade.Data(i).value)), DrawY + 7 + 14, White
                        End If
                    ElseIf TheirTrade.Data(i).TradeType = 2 Then '//Pokemon
                        Sprite = Pokemon(TheirTrade.Data(i).Num).Sprite
                        If Sprite > 0 And Sprite <= Count_PokemonIcon Then
                            RenderTexture Tex_PokemonIcon(Sprite), DrawX + 7, DrawY + 7, MapAnim * TILE_X, 0, TILE_X, TILE_Y, TILE_X, TILE_Y
                        End If
                    End If
                End If
            End If
        Next
        
        '//Set
        If YourTrade.TradeSet Then
            RenderTexture Tex_Gui(.Pic), .X + 2, .Y + 36, 12, 469, 199, 24, 199, 24
        End If
        If TheirTrade.TradeSet Then
            RenderTexture Tex_Gui(.Pic), .X + 218, .Y + 36, 12, 494, 199, 24, 199, 24
        End If
        
        '//Name
        RenderText Font_Default, Trim$(Player(MyIndex).Name) & "'s Trade", .X + 15, .Y + 39, White
        RenderText Font_Default, Trim$(Player(TradeIndex).Name) & "'s Trade", .X + 400 - (GetTextWidth(Font_Default, Trim$(Player(TradeIndex).Name) & "'s Trade")) - 4, .Y + 39, White
        
        '//Text
        If EditInputMoney Then
            currencyText = "$" & TradeInputMoney & TextLine
            RenderArrayText Font_Default, UpdateChatText(Font_Default, currencyText, 112), .X + 66, .Y + 279, 250, White
        Else
            If TradeInputMoney <> vbNullString And Val(TradeInputMoney) <> YourTrade.TradeMoney Then
                currencyText = "$" & TradeInputMoney
                RenderArrayText Font_Default, UpdateChatText(Font_Default, currencyText, 112), .X + 66, .Y + 279, 250, White
            Else
                currencyText = "$" & YourTrade.TradeMoney
                RenderArrayText Font_Default, UpdateChatText(Font_Default, currencyText, 135), .X + 66, .Y + 279, 250, White
            End If
        End If
        
        currencyText = "$" & TheirTrade.TradeMoney
        RenderArrayText Font_Default, UpdateChatText(Font_Default, currencyText, 135), .X + 276, .Y + 279, 250, White
    End With
End Sub

Private Sub DrawPokedex()
Dim i As Long
Dim DrawX As Long, DrawY As Long
Dim pokeDexIndex As Long
Dim DescText() As String
Dim MaxY As Long, PosY As Long, PosX As Long

    With GUI(GuiEnum.GUI_POKEDEX)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height
    
        '//Buttons
        For i = ButtonEnum.Pokedex_Close To ButtonEnum.Pokedex_ScrollDown
            If CanShowButton(i) Then
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
            End If
        Next
        
        '//Scroll
        RenderTexture Tex_Gui(.Pic), .X + 7, .Y + PokedexScrollStartY + ((PokedexScrollEndY - PokedexScrollSize) - PokedexScrollY), 159, 300, 19, 35, 19, 35
    
        '//Icon
        For i = (PokedexViewCount * 8) To (PokedexViewCount * 8) + 31
            If i >= 0 And i <= PokedexHighIndex Then
                pokeDexIndex = i + 1
                DrawX = (31 + ((4 + 44) * (((((i + 1) - (PokedexViewCount * 8)) - 1) Mod 8))))
                DrawY = (42 + ((4 + 46) * ((((i + 1) - (PokedexViewCount * 8)) - 1) \ 8)))
                
                RenderTexture Tex_Gui(.Pic), .X + DrawX, .Y + DrawY, 369, 290, 44, 46, 44, 46
                
                If PlayerPokedex(pokeDexIndex).Obtained = YES Then
                    '//Icon
                    If Pokemon(pokeDexIndex).Sprite > 0 And Pokemon(pokeDexIndex).Sprite < Count_PokemonIcon Then
                        RenderTexture Tex_PokemonIcon(Pokemon(pokeDexIndex).Sprite), .X + DrawX + 7, .Y + DrawY + 7, MapAnim * 32, 0, 32, 32, 32, 32
                    End If
                Else
                    If PlayerPokedex(pokeDexIndex).Scanned = YES Then
                        '//Icon
                        If Pokemon(pokeDexIndex).Sprite > 0 And Pokemon(pokeDexIndex).Sprite < Count_PokemonIcon Then
                            RenderTexture Tex_PokemonIcon(Pokemon(pokeDexIndex).Sprite), .X + DrawX + 7, .Y + DrawY + 7, MapAnim * 32, 0, 32, 32, 32, 32, D3DColorARGB(255, 50, 50, 50)
                        End If
                    Else
                        RenderTexture Tex_Gui(.Pic), .X + DrawX + 7, .Y + DrawY + 7, 92, 304, 32, 32, 32, 32
                        RenderText Font_Default, pokeDexIndex, .X + DrawX + 5, .Y + DrawY + 20, White
                    End If
                End If
            End If
        Next
        
        If PokedexInfoIndex > 0 And PokedexShowTimer <= GetTickCount Then
            If PlayerPokedex(PokedexInfoIndex).Obtained = YES Then
                WordWrap_Array Font_Default, Trim$(Pokemon(PokedexInfoIndex).PokeDexEntry), 250, DescText
                MaxY = UBound(DescText) + 2
                PosY = (.Y + 39) + ((202 * 0.5) - ((MaxY * 20) * 0.5))
                RenderTexture Tex_System(gSystemEnum.UserInterface), .X + 28, .Y + 39, 0, 8, 386, 202, 1, 1, D3DColorARGB(200, 0, 0, 0)
                
                PosX = (.X + 28) + ((386 * 0.5) - (GetTextWidth(Font_Default, Trim$(Pokemon(PokedexInfoIndex).Name)) * 0.5))
                RenderText Font_Default, Trim$(Pokemon(PokedexInfoIndex).Name), PosX, PosY, White
                For i = 1 To UBound(DescText)
                    PosX = (.X + 28) + ((386 * 0.5) - (GetTextWidth(Font_Default, Trim$(DescText(i))) * 0.5))
                    RenderText Font_Default, Trim$(DescText(i)), PosX, PosY + ((i + 1) * 20), White
                Next
            End If
        End If
    End With
End Sub

Private Sub DrawPokemonSummary()
    Dim i As Long, setStat As Byte, MoveNum As Integer

    With GUI(GuiEnum.GUI_POKEMONSUMMARY)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub

        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height

        '//Buttons
        For i = ButtonEnum.PokemonSummary_Close To ButtonEnum.PokemonSummary_Close
            If CanShowButton(i) Then
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
            End If
        Next

        '//Summary
        If SummarySlot > 0 Then
            Select Case SummaryType
            Case 1
                If PlayerPokemons(SummarySlot).Num > 0 Then
                    RenderText Font_Default, Trim$(Pokemon(PlayerPokemons(SummarySlot).Num).Name), .X + 191, .Y + 40, D3DColorARGB(180, 255, 255, 255), False
                    'If PlayerPokemons(SummarySlot).IsShiny = YES Then
                    '    RenderText Font_Default, "Yes", .X + 191, .Y + 63, D3DColorARGB(180, 255, 255, 255), False
                    'Else
                    '    RenderText Font_Default, "No", .X + 191, .Y + 63, D3DColorARGB(180, 255, 255, 255), False
                    'End If
                    RenderText Font_Default, Trim$(CheckNatureString(PlayerPokemons(SummarySlot).Nature)), .X + 191, .Y + 86, D3DColorARGB(180, 255, 255, 255), False
                    RenderText Font_Default, PlayerPokemons(SummarySlot).Level, .X + 191, .Y + 109, D3DColorARGB(180, 255, 255, 255), False
                    'RenderText Font_Default, PlayerPokemons(SummarySlot).CurHP & "/" & PlayerPokemons(SummarySlot).MaxHP, .X + 191, .Y + 143, D3DColorARGB(180, 255, 255, 255), False
                    RenderText Font_Default, PlayerPokemons(SummarySlot).CurExp & "/" & PlayerPokemons(SummarySlot).NextExp, .X + 191, .Y + 134, D3DColorARGB(180, 255, 255, 255), False

                    setStat = StatEnum.HP
                    RenderText Font_Default, PlayerPokemons(SummarySlot).Stat(setStat), .X + 191, .Y + 166, D3DColorARGB(180, 255, 255, 255), False
                    RenderText Font_Default, " (" & PlayerPokemons(SummarySlot).StatIV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, PlayerPokemons(SummarySlot).Stat(setStat)), .Y + 166, D3DColorARGB(180, 237, 233, 141), False
                    RenderText Font_Default, " (" & PlayerPokemons(SummarySlot).StatEV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, PlayerPokemons(SummarySlot).Stat(setStat)) + GetTextWidth(Font_Default, " (" & PlayerPokemons(SummarySlot).StatIV(setStat) & ")"), .Y + 166, D3DColorARGB(180, 169, 241, 163), False

                    setStat = StatEnum.Atk
                    RenderText Font_Default, PlayerPokemons(SummarySlot).Stat(setStat), .X + 191, .Y + 189, D3DColorARGB(180, 255, 255, 255), False
                    RenderText Font_Default, " (" & PlayerPokemons(SummarySlot).StatIV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, PlayerPokemons(SummarySlot).Stat(setStat)), .Y + 189, D3DColorARGB(180, 237, 233, 141), False
                    RenderText Font_Default, " (" & PlayerPokemons(SummarySlot).StatEV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, PlayerPokemons(SummarySlot).Stat(setStat)) + GetTextWidth(Font_Default, " (" & PlayerPokemons(SummarySlot).StatIV(setStat) & ")"), .Y + 189, D3DColorARGB(180, 169, 241, 163), False

                    setStat = StatEnum.Def
                    RenderText Font_Default, PlayerPokemons(SummarySlot).Stat(setStat), .X + 191, .Y + 212, D3DColorARGB(180, 255, 255, 255), False
                    RenderText Font_Default, " (" & PlayerPokemons(SummarySlot).StatIV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, PlayerPokemons(SummarySlot).Stat(setStat)), .Y + 212, D3DColorARGB(180, 237, 233, 141), False
                    RenderText Font_Default, " (" & PlayerPokemons(SummarySlot).StatEV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, PlayerPokemons(SummarySlot).Stat(setStat)) + GetTextWidth(Font_Default, " (" & PlayerPokemons(SummarySlot).StatIV(setStat) & ")"), .Y + 212, D3DColorARGB(180, 169, 241, 163), False

                    setStat = StatEnum.SpAtk
                    RenderText Font_Default, PlayerPokemons(SummarySlot).Stat(setStat), .X + 191, .Y + 235, D3DColorARGB(180, 255, 255, 255), False
                    RenderText Font_Default, " (" & PlayerPokemons(SummarySlot).StatIV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, PlayerPokemons(SummarySlot).Stat(setStat)), .Y + 235, D3DColorARGB(180, 237, 233, 141), False
                    RenderText Font_Default, " (" & PlayerPokemons(SummarySlot).StatEV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, PlayerPokemons(SummarySlot).Stat(setStat)) + GetTextWidth(Font_Default, " (" & PlayerPokemons(SummarySlot).StatIV(setStat) & ")"), .Y + 235, D3DColorARGB(180, 169, 241, 163), False

                    setStat = StatEnum.SpDef
                    RenderText Font_Default, PlayerPokemons(SummarySlot).Stat(setStat), .X + 191, .Y + 258, D3DColorARGB(180, 255, 255, 255), False
                    RenderText Font_Default, " (" & PlayerPokemons(SummarySlot).StatIV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, PlayerPokemons(SummarySlot).Stat(setStat)), .Y + 258, D3DColorARGB(180, 237, 233, 141), False
                    RenderText Font_Default, " (" & PlayerPokemons(SummarySlot).StatEV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, PlayerPokemons(SummarySlot).Stat(setStat)) + GetTextWidth(Font_Default, " (" & PlayerPokemons(SummarySlot).StatIV(setStat) & ")"), .Y + 258, D3DColorARGB(180, 169, 241, 163), False

                    setStat = StatEnum.Spd
                    RenderText Font_Default, PlayerPokemons(SummarySlot).Stat(setStat), .X + 191, .Y + 281, D3DColorARGB(180, 255, 255, 255), False
                    RenderText Font_Default, " (" & PlayerPokemons(SummarySlot).StatIV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, PlayerPokemons(SummarySlot).Stat(setStat)), .Y + 281, D3DColorARGB(180, 237, 233, 141), False
                    RenderText Font_Default, " (" & PlayerPokemons(SummarySlot).StatEV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, PlayerPokemons(SummarySlot).Stat(setStat)) + GetTextWidth(Font_Default, " (" & PlayerPokemons(SummarySlot).StatIV(setStat) & ")"), .Y + 281, D3DColorARGB(180, 169, 241, 163), False

                    '//Icon
                    If PlayerPokemons(SummarySlot).IsShiny = YES Then
                        If Pokemon(PlayerPokemons(SummarySlot).Num).Sprite > 0 And Pokemon(PlayerPokemons(SummarySlot).Num).Sprite <= Count_ShinyPokemonPortrait Then
                            RenderTexture Tex_ShinyPokemonPortrait(Pokemon(PlayerPokemons(SummarySlot).Num).Sprite), .X + 11, .Y + 43, 0, 0, 96, 96, 96, 96

                            DrawShinyStar_Summary

                        End If
                    Else
                        If Pokemon(PlayerPokemons(SummarySlot).Num).Sprite > 0 And Pokemon(PlayerPokemons(SummarySlot).Num).Sprite <= Count_PokemonPortrait Then
                            RenderTexture Tex_PokemonPortrait(Pokemon(PlayerPokemons(SummarySlot).Num).Sprite), .X + 11, .Y + 43, 0, 0, 96, 96, 96, 96
                        End If
                    End If

                    ' Held Item
                    If PlayerPokemons(SummarySlot).HeldItem > 0 Then
                        RenderText Font_Default, Trim$(Item(PlayerPokemons(SummarySlot).HeldItem).Name), .X + 10 + ((104 / 2) - (GetTextWidth(Font_Default, Trim$(Item(PlayerPokemons(SummarySlot).HeldItem).Name)) / 2)), .Y + 143, DarkBrown
                        RenderTexture Tex_Item(PokeUseHeld), .X + ((80 / 2) - (GetTextWidth(Font_Default, Trim$(Item(PlayerPokemons(SummarySlot).HeldItem).Name)) / 2)), .Y + 140, 0, 0, 22, 22, 24, 24
                    End If

                    ' Type Texture
                    If Pokemon(PlayerPokemons(SummarySlot).Num).PrimaryType > 0 Then
                        RenderTexture Tex_PokemonTypes(Pokemon(PlayerPokemons(SummarySlot).Num).PrimaryType), .X + 189, .Y + 65, 0, 0, 55, 16, 64, 14
                        If Pokemon(PlayerPokemons(SummarySlot).Num).SecondaryType > 0 Then
                            RenderTexture Tex_PokemonTypes(Pokemon(PlayerPokemons(SummarySlot).Num).SecondaryType), .X + 241, .Y + 65, 0, 0, 55, 16, 64, 14
                        End If
                    End If


                    ' Moves Name tab
                    For i = 1 To 4
                        MoveNum = PlayerPokemons(SummarySlot).Moveset(i).Num
                        If MoveNum > 0 Then
                            ' Degrade
                            RenderTexture Tex_Gui(Hud), .X + 8, .Y + 195 + (i * 30 - 32), 59, 241, 105, 28, 165, 1
                            ' Move Name
                            RenderText Ui_Default, Trim$(PokemonMove(MoveNum).Name), .X + 22, .Y + 190 + (i * 30 - 32), White
                            ' PP
                            RenderText Ui_Default, "PP: " & PlayerPokemons(SummarySlot).Moveset(i).CurPP & "/" & PokemonMove(MoveNum).PP, .X + 45, .Y + 205 + (i * 30 - 32), White

                            ' Poke Type texture
                            If PokemonMove(MoveNum).Type > 0 Then
                                RenderTexture Tex_PokemonTypesSymbol(PokemonMove(MoveNum).Type), .X + 8, .Y + 195 + (i * 30 - 32), 0, 0, 15, 15, 20, 20
                            End If

                            ' Poke Category texture
                            If PokemonMove(MoveNum).Category > 0 Then
                                RenderTexture Tex_Misc(Misc_CategoryTypes), .X + 8, .Y + 208 + (i * 30 - 32), PokemonMove(MoveNum).Category * 20 - 20, 0, 20, 16, 20, 16
                            End If
                        End If
                    Next i
                End If

            Case 2
                If SummaryData > 0 Then
                    If PlayerPokemonStorage(SummaryData).Data(SummarySlot).Num > 0 Then
                        RenderText Font_Default, Trim$(Pokemon(PlayerPokemonStorage(SummaryData).Data(SummarySlot).Num).Name), .X + 191, .Y + 40, D3DColorARGB(180, 255, 255, 255), False
                        'If PlayerPokemonStorage(SummaryData).Data(SummarySlot).IsShiny = YES Then
                        '    RenderText Font_Default, "Yes", .X + 191, .Y + 63, D3DColorARGB(180, 255, 255, 255), False
                        'Else
                        '    RenderText Font_Default, "No", .X + 191, .Y + 63, D3DColorARGB(180, 255, 255, 255), False
                        'End If
                        RenderText Font_Default, Trim$(CheckNatureString(PlayerPokemonStorage(SummaryData).Data(SummarySlot).Nature)), .X + 191, .Y + 86, D3DColorARGB(180, 255, 255, 255), False
                        RenderText Font_Default, PlayerPokemonStorage(SummaryData).Data(SummarySlot).Level, .X + 191, .Y + 109, D3DColorARGB(180, 255, 255, 255), False
                        'RenderText Font_Default, PlayerPokemonStorage(SummaryData).Data(SummarySlot).CurHP & "/" & PlayerPokemonStorage(SummaryData).Data(SummarySlot).MaxHP, .X + 191, .Y + 143, D3DColorARGB(180, 255, 255, 255), False
                        RenderText Font_Default, PlayerPokemonStorage(SummaryData).Data(SummarySlot).CurExp & "/" & PlayerPokemonStorage(SummaryData).Data(SummarySlot).NextExp, .X + 191, .Y + 134, D3DColorARGB(180, 255, 255, 255), False

                        setStat = StatEnum.HP
                        RenderText Font_Default, PlayerPokemonStorage(SummaryData).Data(SummarySlot).Stat(setStat), .X + 191, .Y + 166, D3DColorARGB(180, 255, 255, 255), False
                        RenderText Font_Default, " (" & PlayerPokemonStorage(SummaryData).Data(SummarySlot).StatIV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, PlayerPokemonStorage(SummaryData).Data(SummarySlot).Stat(setStat)), .Y + 166, D3DColorARGB(180, 237, 233, 141), False
                        RenderText Font_Default, " (" & PlayerPokemonStorage(SummaryData).Data(SummarySlot).StatEV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, PlayerPokemonStorage(SummaryData).Data(SummarySlot).Stat(setStat)) + GetTextWidth(Font_Default, " (" & PlayerPokemonStorage(SummaryData).Data(SummarySlot).StatIV(setStat) & ")"), .Y + 166, D3DColorARGB(180, 169, 241, 163), False

                        setStat = StatEnum.Atk
                        RenderText Font_Default, PlayerPokemonStorage(SummaryData).Data(SummarySlot).Stat(setStat), .X + 191, .Y + 189, D3DColorARGB(180, 255, 255, 255), False
                        RenderText Font_Default, " (" & PlayerPokemonStorage(SummaryData).Data(SummarySlot).StatIV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, PlayerPokemonStorage(SummaryData).Data(SummarySlot).Stat(setStat)), .Y + 189, D3DColorARGB(180, 237, 233, 141), False
                        RenderText Font_Default, " (" & PlayerPokemonStorage(SummaryData).Data(SummarySlot).StatEV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, PlayerPokemonStorage(SummaryData).Data(SummarySlot).Stat(setStat)) + GetTextWidth(Font_Default, " (" & PlayerPokemonStorage(SummaryData).Data(SummarySlot).StatIV(setStat) & ")"), .Y + 189, D3DColorARGB(180, 169, 241, 163), False

                        setStat = StatEnum.Def
                        RenderText Font_Default, PlayerPokemonStorage(SummaryData).Data(SummarySlot).Stat(setStat), .X + 191, .Y + 212, D3DColorARGB(180, 255, 255, 255), False
                        RenderText Font_Default, " (" & PlayerPokemonStorage(SummaryData).Data(SummarySlot).StatIV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, PlayerPokemonStorage(SummaryData).Data(SummarySlot).Stat(setStat)), .Y + 212, D3DColorARGB(180, 237, 233, 141), False
                        RenderText Font_Default, " (" & PlayerPokemonStorage(SummaryData).Data(SummarySlot).StatEV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, PlayerPokemonStorage(SummaryData).Data(SummarySlot).Stat(setStat)) + GetTextWidth(Font_Default, " (" & PlayerPokemonStorage(SummaryData).Data(SummarySlot).StatIV(setStat) & ")"), .Y + 212, D3DColorARGB(180, 169, 241, 163), False

                        setStat = StatEnum.SpAtk
                        RenderText Font_Default, PlayerPokemonStorage(SummaryData).Data(SummarySlot).Stat(setStat), .X + 191, .Y + 235, D3DColorARGB(180, 255, 255, 255), False
                        RenderText Font_Default, " (" & PlayerPokemonStorage(SummaryData).Data(SummarySlot).StatIV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, PlayerPokemonStorage(SummaryData).Data(SummarySlot).Stat(setStat)), .Y + 235, D3DColorARGB(180, 237, 233, 141), False
                        RenderText Font_Default, " (" & PlayerPokemonStorage(SummaryData).Data(SummarySlot).StatEV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, PlayerPokemonStorage(SummaryData).Data(SummarySlot).Stat(setStat)) + GetTextWidth(Font_Default, " (" & PlayerPokemonStorage(SummaryData).Data(SummarySlot).StatIV(setStat) & ")"), .Y + 235, D3DColorARGB(180, 169, 241, 163), False

                        setStat = StatEnum.SpDef
                        RenderText Font_Default, PlayerPokemonStorage(SummaryData).Data(SummarySlot).Stat(setStat), .X + 191, .Y + 258, D3DColorARGB(180, 255, 255, 255), False
                        RenderText Font_Default, " (" & PlayerPokemonStorage(SummaryData).Data(SummarySlot).StatIV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, PlayerPokemonStorage(SummaryData).Data(SummarySlot).Stat(setStat)), .Y + 258, D3DColorARGB(180, 237, 233, 141), False
                        RenderText Font_Default, " (" & PlayerPokemonStorage(SummaryData).Data(SummarySlot).StatEV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, PlayerPokemonStorage(SummaryData).Data(SummarySlot).Stat(setStat)) + GetTextWidth(Font_Default, " (" & PlayerPokemonStorage(SummaryData).Data(SummarySlot).StatIV(setStat) & ")"), .Y + 258, D3DColorARGB(180, 169, 241, 163), False

                        setStat = StatEnum.Spd
                        RenderText Font_Default, PlayerPokemonStorage(SummaryData).Data(SummarySlot).Stat(setStat), .X + 191, .Y + 281, D3DColorARGB(180, 255, 255, 255), False
                        RenderText Font_Default, " (" & PlayerPokemonStorage(SummaryData).Data(SummarySlot).StatIV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, PlayerPokemonStorage(SummaryData).Data(SummarySlot).Stat(setStat)), .Y + 281, D3DColorARGB(180, 237, 233, 141), False
                        RenderText Font_Default, " (" & PlayerPokemonStorage(SummaryData).Data(SummarySlot).StatEV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, PlayerPokemonStorage(SummaryData).Data(SummarySlot).Stat(setStat)) + GetTextWidth(Font_Default, " (" & PlayerPokemonStorage(SummaryData).Data(SummarySlot).StatIV(setStat) & ")"), .Y + 281, D3DColorARGB(180, 169, 241, 163), False

                        '//Icon
                        If PlayerPokemonStorage(SummaryData).Data(SummarySlot).IsShiny = YES Then
                            If Pokemon(PlayerPokemonStorage(SummaryData).Data(SummarySlot).Num).Sprite > 0 And Pokemon(PlayerPokemonStorage(SummaryData).Data(SummarySlot).Num).Sprite <= Count_ShinyPokemonPortrait Then
                                RenderTexture Tex_ShinyPokemonPortrait(Pokemon(PlayerPokemonStorage(SummaryData).Data(SummarySlot).Num).Sprite), .X + 11, .Y + 43, 0, 0, 96, 96, 96, 96

                                DrawShinyStar_Summary

                            End If
                        Else
                            If Pokemon(PlayerPokemonStorage(SummaryData).Data(SummarySlot).Num).Sprite > 0 And Pokemon(PlayerPokemonStorage(SummaryData).Data(SummarySlot).Num).Sprite <= Count_PokemonPortrait Then
                                RenderTexture Tex_PokemonPortrait(Pokemon(PlayerPokemonStorage(SummaryData).Data(SummarySlot).Num).Sprite), .X + 11, .Y + 43, 0, 0, 96, 96, 96, 96
                            End If
                        End If

                        ' Held Item
                        If PlayerPokemonStorage(SummaryData).Data(SummarySlot).HeldItem > 0 Then
                            RenderText Font_Default, Trim$(Item(PlayerPokemonStorage(SummaryData).Data(SummarySlot).HeldItem).Name), .X + 10 + ((104 / 2) - (GetTextWidth(Font_Default, Trim$(Item(PlayerPokemonStorage(SummaryData).Data(SummarySlot).HeldItem).Name)) / 2)), .Y + 143, DarkBrown
                            RenderTexture Tex_Item(PokeUseHeld), .X + ((80 / 2) - (GetTextWidth(Font_Default, Trim$(Item(PlayerPokemonStorage(SummaryData).Data(SummarySlot).HeldItem).Name)) / 2)), .Y + 140, 0, 0, 22, 22, 24, 24
                        End If

                        ' Type Texture
                        If Pokemon(PlayerPokemonStorage(SummaryData).Data(SummarySlot).Num).PrimaryType > 0 Then
                            RenderTexture Tex_PokemonTypes(Pokemon(PlayerPokemonStorage(SummaryData).Data(SummarySlot).Num).PrimaryType), .X + 189, .Y + 65, 0, 0, 55, 16, 64, 14
                            If Pokemon(PlayerPokemonStorage(SummaryData).Data(SummarySlot).Num).SecondaryType > 0 Then
                                RenderTexture Tex_PokemonTypes(Pokemon(PlayerPokemonStorage(SummaryData).Data(SummarySlot).Num).SecondaryType), .X + 241, .Y + 65, 0, 0, 55, 16, 64, 14
                            End If
                        End If


                        ' Moves Name tab
                        For i = 1 To 4
                            MoveNum = PlayerPokemonStorage(SummaryData).Data(SummarySlot).Moveset(i).Num
                            If MoveNum > 0 Then
                                ' Degrade
                                RenderTexture Tex_Gui(Hud), .X + 8, .Y + 195 + (i * 30 - 32), 59, 241, 105, 28, 165, 1
                                ' Move Name
                                RenderText Ui_Default, Trim$(PokemonMove(MoveNum).Name), .X + 22, .Y + 190 + (i * 30 - 32), White
                                ' PP
                                RenderText Ui_Default, "PP: " & PlayerPokemonStorage(SummaryData).Data(SummarySlot).Moveset(i).CurPP & "/" & PokemonMove(MoveNum).PP, .X + 45, .Y + 205 + (i * 30 - 32), White

                                ' Poke Type texture
                                If PokemonMove(MoveNum).Type > 0 Then
                                    RenderTexture Tex_PokemonTypesSymbol(PokemonMove(MoveNum).Type), .X + 8, .Y + 195 + (i * 30 - 32), 0, 0, 15, 15, 20, 20
                                End If

                                ' Poke Category texture
                                If PokemonMove(MoveNum).Category > 0 Then
                                    RenderTexture Tex_Misc(Misc_CategoryTypes), .X + 8, .Y + 208 + (i * 30 - 32), PokemonMove(MoveNum).Category * 20 - 20, 0, 20, 16, 20, 16
                                End If
                            End If
                        Next i
                    End If
                End If

            Case 3
                If YourTrade.Data(SummarySlot).TradeType = 2 Then    '//Pokemon
                    If YourTrade.Data(SummarySlot).Num > 0 Then
                        RenderText Font_Default, Trim$(Pokemon(YourTrade.Data(SummarySlot).Num).Name), .X + 191, .Y + 40, D3DColorARGB(180, 255, 255, 255), False
                        'If YourTrade.Data(SummarySlot).IsShiny = YES Then
                        '    RenderText Font_Default, "Yes", .X + 191, .Y + 63, D3DColorARGB(180, 255, 255, 255), False
                        'Else
                        '    RenderText Font_Default, "No", .X + 191, .Y + 63, D3DColorARGB(180, 255, 255, 255), False
                        'End If
                        RenderText Font_Default, Trim$(CheckNatureString(YourTrade.Data(SummarySlot).Nature)), .X + 191, .Y + 86, D3DColorARGB(180, 255, 255, 255), False
                        RenderText Font_Default, YourTrade.Data(SummarySlot).Level, .X + 191, .Y + 109, D3DColorARGB(180, 255, 255, 255), False
                        'RenderText Font_Default, YourTrade.Data(SummarySlot).CurHP & "/" & YourTrade.Data(SummarySlot).MaxHP, .X + 191, .Y + 143, D3DColorARGB(180, 255, 255, 255), False
                        RenderText Font_Default, YourTrade.Data(SummarySlot).CurExp & "/" & YourTrade.Data(SummarySlot).NextExp, .X + 191, .Y + 134, D3DColorARGB(180, 255, 255, 255), False

                        setStat = StatEnum.HP
                        RenderText Font_Default, YourTrade.Data(SummarySlot).Stat(setStat), .X + 191, .Y + 166, D3DColorARGB(180, 255, 255, 255), False
                        RenderText Font_Default, " (" & YourTrade.Data(SummarySlot).StatIV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, YourTrade.Data(SummarySlot).Stat(setStat)), .Y + 166, D3DColorARGB(180, 237, 233, 141), False
                        RenderText Font_Default, " (" & YourTrade.Data(SummarySlot).StatEV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, YourTrade.Data(SummarySlot).Stat(setStat)) + GetTextWidth(Font_Default, " (" & YourTrade.Data(SummarySlot).StatIV(setStat) & ")"), .Y + 166, D3DColorARGB(180, 169, 241, 163), False

                        setStat = StatEnum.Atk
                        RenderText Font_Default, YourTrade.Data(SummarySlot).Stat(setStat), .X + 191, .Y + 189, D3DColorARGB(180, 255, 255, 255), False
                        RenderText Font_Default, " (" & YourTrade.Data(SummarySlot).StatIV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, YourTrade.Data(SummarySlot).Stat(setStat)), .Y + 189, D3DColorARGB(180, 237, 233, 141), False
                        RenderText Font_Default, " (" & YourTrade.Data(SummarySlot).StatEV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, YourTrade.Data(SummarySlot).Stat(setStat)) + GetTextWidth(Font_Default, " (" & YourTrade.Data(SummarySlot).StatIV(setStat) & ")"), .Y + 189, D3DColorARGB(180, 169, 241, 163), False

                        setStat = StatEnum.Def
                        RenderText Font_Default, YourTrade.Data(SummarySlot).Stat(setStat), .X + 191, .Y + 212, D3DColorARGB(180, 255, 255, 255), False
                        RenderText Font_Default, " (" & YourTrade.Data(SummarySlot).StatIV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, YourTrade.Data(SummarySlot).Stat(setStat)), .Y + 212, D3DColorARGB(180, 237, 233, 141), False
                        RenderText Font_Default, " (" & YourTrade.Data(SummarySlot).StatEV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, YourTrade.Data(SummarySlot).Stat(setStat)) + GetTextWidth(Font_Default, " (" & YourTrade.Data(SummarySlot).StatIV(setStat) & ")"), .Y + 212, D3DColorARGB(180, 169, 241, 163), False

                        setStat = StatEnum.SpAtk
                        RenderText Font_Default, YourTrade.Data(SummarySlot).Stat(setStat), .X + 191, .Y + 235, D3DColorARGB(180, 255, 255, 255), False
                        RenderText Font_Default, " (" & YourTrade.Data(SummarySlot).StatIV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, YourTrade.Data(SummarySlot).Stat(setStat)), .Y + 235, D3DColorARGB(180, 237, 233, 141), False
                        RenderText Font_Default, " (" & YourTrade.Data(SummarySlot).StatEV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, YourTrade.Data(SummarySlot).Stat(setStat)) + GetTextWidth(Font_Default, " (" & YourTrade.Data(SummarySlot).StatIV(setStat) & ")"), .Y + 235, D3DColorARGB(180, 169, 241, 163), False

                        setStat = StatEnum.SpDef
                        RenderText Font_Default, YourTrade.Data(SummarySlot).Stat(setStat), .X + 191, .Y + 258, D3DColorARGB(180, 255, 255, 255), False
                        RenderText Font_Default, " (" & YourTrade.Data(SummarySlot).StatIV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, YourTrade.Data(SummarySlot).Stat(setStat)), .Y + 258, D3DColorARGB(180, 237, 233, 141), False
                        RenderText Font_Default, " (" & YourTrade.Data(SummarySlot).StatEV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, YourTrade.Data(SummarySlot).Stat(setStat)) + GetTextWidth(Font_Default, " (" & YourTrade.Data(SummarySlot).StatIV(setStat) & ")"), .Y + 258, D3DColorARGB(180, 169, 241, 163), False

                        setStat = StatEnum.Spd
                        RenderText Font_Default, YourTrade.Data(SummarySlot).Stat(setStat), .X + 191, .Y + 281, D3DColorARGB(180, 255, 255, 255), False
                        RenderText Font_Default, " (" & YourTrade.Data(SummarySlot).StatIV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, YourTrade.Data(SummarySlot).Stat(setStat)), .Y + 281, D3DColorARGB(180, 237, 233, 141), False
                        RenderText Font_Default, " (" & YourTrade.Data(SummarySlot).StatEV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, YourTrade.Data(SummarySlot).Stat(setStat)) + GetTextWidth(Font_Default, " (" & YourTrade.Data(SummarySlot).StatIV(setStat) & ")"), .Y + 281, D3DColorARGB(180, 169, 241, 163), False

                        '//Icon
                        If YourTrade.Data(SummarySlot).IsShiny = YES Then
                            If Pokemon(YourTrade.Data(SummarySlot).Num).Sprite > 0 And Pokemon(YourTrade.Data(SummarySlot).Num).Sprite <= Count_ShinyPokemonPortrait Then
                                RenderTexture Tex_ShinyPokemonPortrait(Pokemon(YourTrade.Data(SummarySlot).Num).Sprite), .X + 11, .Y + 43, 0, 0, 96, 96, 96, 96

                                DrawShinyStar_Summary

                            End If
                        Else
                            If Pokemon(YourTrade.Data(SummarySlot).Num).Sprite > 0 And Pokemon(YourTrade.Data(SummarySlot).Num).Sprite <= Count_PokemonPortrait Then
                                RenderTexture Tex_PokemonPortrait(Pokemon(YourTrade.Data(SummarySlot).Num).Sprite), .X + 11, .Y + 43, 0, 0, 96, 96, 96, 96
                            End If
                        End If

                        ' Held Item
                        If YourTrade.Data(SummarySlot).HeldItem > 0 Then
                            RenderText Font_Default, Trim$(Item(YourTrade.Data(SummarySlot).HeldItem).Name), .X + 10 + ((104 / 2) - (GetTextWidth(Font_Default, Trim$(Item(YourTrade.Data(SummarySlot).HeldItem).Name)) / 2)), .Y + 143, DarkBrown
                            RenderTexture Tex_Item(PokeUseHeld), .X + ((80 / 2) - (GetTextWidth(Font_Default, Trim$(Item(YourTrade.Data(SummarySlot).HeldItem).Name)) / 2)), .Y + 140, 0, 0, 22, 22, 24, 24
                        End If

                        ' Type Texture
                        If Pokemon(YourTrade.Data(SummarySlot).Num).PrimaryType > 0 Then
                            RenderTexture Tex_PokemonTypes(Pokemon(YourTrade.Data(SummarySlot).Num).PrimaryType), .X + 189, .Y + 65, 0, 0, 55, 16, 64, 14
                            If Pokemon(YourTrade.Data(SummarySlot).Num).SecondaryType > 0 Then
                                RenderTexture Tex_PokemonTypes(Pokemon(YourTrade.Data(SummarySlot).Num).SecondaryType), .X + 241, .Y + 65, 0, 0, 55, 16, 64, 14
                            End If
                        End If


                        ' Moves Name tab
                        For i = 1 To 4
                            MoveNum = YourTrade.Data(SummarySlot).Moveset(i).Num
                            If MoveNum > 0 Then
                                ' Degrade
                                RenderTexture Tex_Gui(Hud), .X + 8, .Y + 195 + (i * 30 - 32), 59, 241, 105, 28, 165, 1
                                ' Move Name
                                RenderText Ui_Default, Trim$(PokemonMove(MoveNum).Name), .X + 22, .Y + 190 + (i * 30 - 32), White
                                ' PP
                                RenderText Ui_Default, "PP: " & YourTrade.Data(SummarySlot).Moveset(i).CurPP & "/" & PokemonMove(MoveNum).PP, .X + 45, .Y + 205 + (i * 30 - 32), White

                                ' Poke Type texture
                                If PokemonMove(MoveNum).Type > 0 Then
                                    RenderTexture Tex_PokemonTypesSymbol(PokemonMove(MoveNum).Type), .X + 8, .Y + 195 + (i * 30 - 32), 0, 0, 15, 15, 20, 20
                                End If

                                ' Poke Category texture
                                If PokemonMove(MoveNum).Category > 0 Then
                                    RenderTexture Tex_Misc(Misc_CategoryTypes), .X + 8, .Y + 208 + (i * 30 - 32), PokemonMove(MoveNum).Category * 20 - 20, 0, 20, 16, 20, 16
                                End If
                            End If
                        Next i
                    End If
                End If
            Case 4
                If TheirTrade.Data(SummarySlot).TradeType = 2 Then    '//Pokemon
                    If TheirTrade.Data(SummarySlot).Num > 0 Then
                        RenderText Font_Default, Trim$(Pokemon(TheirTrade.Data(SummarySlot).Num).Name), .X + 191, .Y + 40, D3DColorARGB(180, 255, 255, 255), False
                        'If TheirTrade.Data(SummarySlot).IsShiny = YES Then
                        '    RenderText Font_Default, "Yes", .X + 191, .Y + 63, D3DColorARGB(180, 255, 255, 255), False
                        'Else
                        '    RenderText Font_Default, "No", .X + 191, .Y + 63, D3DColorARGB(180, 255, 255, 255), False
                        'End If
                        RenderText Font_Default, Trim$(CheckNatureString(TheirTrade.Data(SummarySlot).Nature)), .X + 191, .Y + 86, D3DColorARGB(180, 255, 255, 255), False
                        RenderText Font_Default, TheirTrade.Data(SummarySlot).Level, .X + 191, .Y + 109, D3DColorARGB(180, 255, 255, 255), False
                        'RenderText Font_Default, TheirTrade.Data(SummarySlot).CurHP & "/" & TheirTrade.Data(SummarySlot).MaxHP, .X + 191, .Y + 143, D3DColorARGB(180, 255, 255, 255), False
                        RenderText Font_Default, TheirTrade.Data(SummarySlot).CurExp & "/" & TheirTrade.Data(SummarySlot).NextExp, .X + 191, .Y + 134, D3DColorARGB(180, 255, 255, 255), False

                        setStat = StatEnum.HP
                        RenderText Font_Default, TheirTrade.Data(SummarySlot).Stat(setStat), .X + 191, .Y + 166, D3DColorARGB(180, 255, 255, 255), False
                        RenderText Font_Default, " (" & TheirTrade.Data(SummarySlot).StatIV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, TheirTrade.Data(SummarySlot).Stat(setStat)), .Y + 166, D3DColorARGB(180, 237, 233, 141), False
                        RenderText Font_Default, " (" & TheirTrade.Data(SummarySlot).StatEV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, TheirTrade.Data(SummarySlot).Stat(setStat)) + GetTextWidth(Font_Default, " (" & TheirTrade.Data(SummarySlot).StatIV(setStat) & ")"), .Y + 166, D3DColorARGB(180, 169, 241, 163), False

                        setStat = StatEnum.Atk
                        RenderText Font_Default, TheirTrade.Data(SummarySlot).Stat(setStat), .X + 191, .Y + 189, D3DColorARGB(180, 255, 255, 255), False
                        RenderText Font_Default, " (" & TheirTrade.Data(SummarySlot).StatIV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, TheirTrade.Data(SummarySlot).Stat(setStat)), .Y + 189, D3DColorARGB(180, 237, 233, 141), False
                        RenderText Font_Default, " (" & TheirTrade.Data(SummarySlot).StatEV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, TheirTrade.Data(SummarySlot).Stat(setStat)) + GetTextWidth(Font_Default, " (" & TheirTrade.Data(SummarySlot).StatIV(setStat) & ")"), .Y + 189, D3DColorARGB(180, 169, 241, 163), False

                        setStat = StatEnum.Def
                        RenderText Font_Default, TheirTrade.Data(SummarySlot).Stat(setStat), .X + 191, .Y + 212, D3DColorARGB(180, 255, 255, 255), False
                        RenderText Font_Default, " (" & TheirTrade.Data(SummarySlot).StatIV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, TheirTrade.Data(SummarySlot).Stat(setStat)), .Y + 212, D3DColorARGB(180, 237, 233, 141), False
                        RenderText Font_Default, " (" & TheirTrade.Data(SummarySlot).StatEV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, TheirTrade.Data(SummarySlot).Stat(setStat)) + GetTextWidth(Font_Default, " (" & TheirTrade.Data(SummarySlot).StatIV(setStat) & ")"), .Y + 212, D3DColorARGB(180, 169, 241, 163), False

                        setStat = StatEnum.SpAtk
                        RenderText Font_Default, TheirTrade.Data(SummarySlot).Stat(setStat), .X + 191, .Y + 235, D3DColorARGB(180, 255, 255, 255), False
                        RenderText Font_Default, " (" & TheirTrade.Data(SummarySlot).StatIV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, TheirTrade.Data(SummarySlot).Stat(setStat)), .Y + 235, D3DColorARGB(180, 237, 233, 141), False
                        RenderText Font_Default, " (" & TheirTrade.Data(SummarySlot).StatEV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, TheirTrade.Data(SummarySlot).Stat(setStat)) + GetTextWidth(Font_Default, " (" & TheirTrade.Data(SummarySlot).StatIV(setStat) & ")"), .Y + 235, D3DColorARGB(180, 169, 241, 163), False

                        setStat = StatEnum.SpDef
                        RenderText Font_Default, TheirTrade.Data(SummarySlot).Stat(setStat), .X + 191, .Y + 258, D3DColorARGB(180, 255, 255, 255), False
                        RenderText Font_Default, " (" & TheirTrade.Data(SummarySlot).StatIV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, TheirTrade.Data(SummarySlot).Stat(setStat)), .Y + 258, D3DColorARGB(180, 237, 233, 141), False
                        RenderText Font_Default, " (" & TheirTrade.Data(SummarySlot).StatEV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, TheirTrade.Data(SummarySlot).Stat(setStat)) + GetTextWidth(Font_Default, " (" & TheirTrade.Data(SummarySlot).StatIV(setStat) & ")"), .Y + 258, D3DColorARGB(180, 169, 241, 163), False

                        setStat = StatEnum.Spd
                        RenderText Font_Default, TheirTrade.Data(SummarySlot).Stat(setStat), .X + 191, .Y + 281, D3DColorARGB(180, 255, 255, 255), False
                        RenderText Font_Default, " (" & TheirTrade.Data(SummarySlot).StatIV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, TheirTrade.Data(SummarySlot).Stat(setStat)), .Y + 281, D3DColorARGB(180, 237, 233, 141), False
                        RenderText Font_Default, " (" & TheirTrade.Data(SummarySlot).StatEV(setStat) & ")", .X + 191 + GetTextWidth(Font_Default, TheirTrade.Data(SummarySlot).Stat(setStat)) + GetTextWidth(Font_Default, " (" & TheirTrade.Data(SummarySlot).StatIV(setStat) & ")"), .Y + 281, D3DColorARGB(180, 169, 241, 163), False

                        '//Icon
                        If TheirTrade.Data(SummarySlot).IsShiny = YES Then
                            If Pokemon(TheirTrade.Data(SummarySlot).Num).Sprite > 0 And Pokemon(TheirTrade.Data(SummarySlot).Num).Sprite <= Count_ShinyPokemonPortrait Then
                                RenderTexture Tex_ShinyPokemonPortrait(Pokemon(TheirTrade.Data(SummarySlot).Num).Sprite), .X + 11, .Y + 43, 0, 0, 96, 96, 96, 96

                                DrawShinyStar_Summary

                            End If
                        Else
                            If Pokemon(TheirTrade.Data(SummarySlot).Num).Sprite > 0 And Pokemon(TheirTrade.Data(SummarySlot).Num).Sprite <= Count_PokemonPortrait Then
                                RenderTexture Tex_PokemonPortrait(Pokemon(TheirTrade.Data(SummarySlot).Num).Sprite), .X + 11, .Y + 43, 0, 0, 96, 96, 96, 96
                            End If
                        End If

                        ' Held Item
                        If TheirTrade.Data(SummarySlot).HeldItem > 0 Then
                            RenderText Font_Default, Trim$(Item(TheirTrade.Data(SummarySlot).HeldItem).Name), .X + 10 + ((104 / 2) - (GetTextWidth(Font_Default, Trim$(Item(TheirTrade.Data(SummarySlot).HeldItem).Name)) / 2)), .Y + 143, DarkBrown
                            RenderTexture Tex_Item(PokeUseHeld), .X + ((80 / 2) - (GetTextWidth(Font_Default, Trim$(Item(TheirTrade.Data(SummarySlot).HeldItem).Name)) / 2)), .Y + 140, 0, 0, 22, 22, 24, 24
                        End If

                        ' Type Texture
                        If Pokemon(TheirTrade.Data(SummarySlot).Num).PrimaryType > 0 Then
                            RenderTexture Tex_PokemonTypes(Pokemon(PlayerPokemons(SummarySlot).Num).PrimaryType), .X + 189, .Y + 65, 0, 0, 55, 16, 64, 14
                            If Pokemon(TheirTrade.Data(SummarySlot).Num).SecondaryType > 0 Then
                                RenderTexture Tex_PokemonTypes(Pokemon(PlayerPokemons(SummarySlot).Num).SecondaryType), .X + 241, .Y + 65, 0, 0, 55, 16, 64, 14
                            End If
                        End If


                        ' Moves Name tab
                        For i = 1 To 4
                            MoveNum = TheirTrade.Data(SummarySlot).Moveset(i).Num
                            If MoveNum > 0 Then
                                ' Degrade
                                RenderTexture Tex_Gui(Hud), .X + 8, .Y + 195 + (i * 30 - 32), 59, 241, 105, 28, 165, 1
                                ' Move Name
                                RenderText Ui_Default, Trim$(PokemonMove(MoveNum).Name), .X + 22, .Y + 190 + (i * 30 - 32), White
                                ' PP
                                RenderText Ui_Default, "PP: " & TheirTrade.Data(SummarySlot).Moveset(i).CurPP & "/" & PokemonMove(MoveNum).PP, .X + 45, .Y + 205 + (i * 30 - 32), White

                                ' Poke Type texture
                                If PokemonMove(MoveNum).Type > 0 Then
                                    RenderTexture Tex_PokemonTypesSymbol(PokemonMove(MoveNum).Type), .X + 8, .Y + 195 + (i * 30 - 32), 0, 0, 15, 15, 20, 20
                                End If

                                ' Poke Category texture
                                If PokemonMove(MoveNum).Category > 0 Then
                                    RenderTexture Tex_Misc(Misc_CategoryTypes), .X + 8, .Y + 208 + (i * 30 - 32), PokemonMove(MoveNum).Category * 20 - 20, 0, 20, 16, 20, 16
                                End If
                            End If
                        Next i
                    End If
                End If
            End Select
        End If
    End With
End Sub

Private Sub DrawShinyStar_Summary()
Dim i As Long

    With GUI(GuiEnum.GUI_POKEMONSUMMARY)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Render the window
        RenderTexture Tex_Misc(Misc_ShinySummary), .X + 7, .Y + 41, ShinySummaryStep * 40, 0, 40, 39, 40, 39
    End With
End Sub

Private Sub DrawRelearn()
Dim i As Long
Dim MoveNum As Long, MN As Long
Dim X As Byte
Dim CanLearn As Boolean

    With GUI(GuiEnum.GUI_RELEARN)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height
        
        '//Buttons
        For i = ButtonEnum.Relearn_Close To ButtonEnum.Relearn_ScrollUp
            If CanShowButton(i) Then
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
            End If
        Next
        
        If MoveRelearnPokeNum > 0 Then
            For i = 1 To 5
                CanLearn = True
                MoveNum = i + MoveRelearnCurPos
                If MoveNum >= 0 And MoveNum <= MoveRelearnMaxIndex Then
                    If Pokemon(MoveRelearnPokeNum).Moveset(MoveNum).MoveNum > 0 Then
                        MN = Pokemon(MoveRelearnPokeNum).Moveset(MoveNum).MoveNum
                        '//Check if pokemon already learned the move or pokemon doesn't have enough level
                        If MoveRelearnPokeSlot > 0 Then
                            If PlayerPokemons(MoveRelearnPokeSlot).Num > 0 Then
                                For X = 1 To MAX_MOVESET
                                    If PlayerPokemons(MoveRelearnPokeSlot).Moveset(X).Num = MN Then
                                        CanLearn = False
                                    End If
                                Next
                                If PlayerPokemons(MoveRelearnPokeSlot).Level < Pokemon(MoveRelearnPokeNum).Moveset(MoveNum).MoveLevel Then
                                    CanLearn = False
                                End If
                                
                                If CanLearn Then
                                    RenderTexture Tex_Gui(.Pic), .X + 36, .Y + 46 + ((i - 1) * 48), 35, 328, 198, 42, 198, 42
                                    RenderText Font_Default, Trim$(PokemonMove(MN).Name), .X + 36 + 5, .Y + 46 + ((i - 1) * 48) + 11, White
                                    RenderText Font_Default, "Lv" & Pokemon(MoveRelearnPokeNum).Moveset(MoveNum).MoveLevel, .X + 36 + 5 + 135, .Y + 46 + ((i - 1) * 48) + 11, White
                                Else
                                    RenderTexture Tex_Gui(.Pic), .X + 36, .Y + 46 + ((i - 1) * 48), 35, 328, 198, 42, 198, 42, D3DColorARGB(150, 255, 255, 255)
                                    RenderText Font_Default, Trim$(PokemonMove(MN).Name), .X + 36 + 5, .Y + 46 + ((i - 1) * 48) + 11, White, True, 150
                                    RenderText Font_Default, "Lv" & Pokemon(MoveRelearnPokeNum).Moveset(MoveNum).MoveLevel, .X + 36 + 5 + 135, .Y + 46 + ((i - 1) * 48) + 11, White, True, 150
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub DrawRank()
    Dim i As Long
    Dim rankIndex As Long
    Dim DrawX As Integer, DrawY As Integer

    With GUI(GuiEnum.GUI_RANK)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub

        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height

        '//Buttons
        For i = ButtonEnum.Rank_Close To ButtonEnum.Rank_ScrollDown
            If CanShowButton(i) Then
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
            End If
        Next
        
        '//Scroll
        If RankingHighIndex > RankingViewLine Then
            RenderTexture Tex_Gui(.Pic), .X + 7, .Y + RankingScrollStartY + ((RankingScrollEndY - RankingScrollSize) - RankingScrollY), 28, 373, 19, 35, 19, 35
        End If
        
        '//ShowRank
        '//ID Icons Top 1, 2 and 3.
        '// 1º = 528 => 2º = 529 => 3º = 530
        For i = RankingViewCount To (RankingViewCount) + (RankingViewLine - 1)
            If i >= 0 And i < RankingHighIndex Then
                rankIndex = i + 1
                
                DrawX = .X + 30
                DrawY = .Y + (40 + (31 * ((((rankIndex) - (RankingViewCount)) - 1))))
                
                RenderTexture Tex_Gui(.Pic), DrawX, DrawY, 28, 328, 212, 28, 212, 28
                ' Draw Index
                ' Icon in 1º at 3º position
                If rankIndex >= 1 And rankIndex <= 3 Then
                    RenderTexture Tex_Item(528 - 1 + rankIndex), DrawX, DrawY, 0, 0, 24, 24, 24, 24
                    RenderText Font_Default, Trim$(Rank(rankIndex).Name) & " Lv" & Rank(rankIndex).Level, DrawX + 15, DrawY, Dark
                Else
                    RenderText Font_Default, rankIndex & ": " & Trim$(Rank(rankIndex).Name) & " Lv" & Rank(rankIndex).Level, DrawX + 5, DrawY, Dark
                End If
            End If
        Next
    End With
End Sub

Private Sub DrawBadge()
Dim i As Long
Dim PosX As Long, PosY As Long, TexX As Long, TexY As Long

    With GUI(GuiEnum.GUI_BADGE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height
        
        '//Buttons
        For i = ButtonEnum.Badge_Close To ButtonEnum.Badge_Close
            If CanShowButton(i) Then
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
            End If
        Next
        
        '//Badge
        For i = 1 To MAX_BADGE
            If Player(MyIndex).Badge(i) > 0 Then
                PosX = .X + (84 + ((1 + 20) * (((i - 1) Mod 8))))
                PosY = .Y + (42 + ((10 + 20) * ((i - 1) \ 8)))
                TexX = (37 + ((20) * (((i - 1) Mod 8))))
                TexY = (203 + ((20) * ((i - 1) \ 8)))
                
                '//Draw Icon
                RenderTexture Tex_Gui(.Pic), PosX, PosY, TexX, TexY, 20, 20, 20, 20
            End If
        Next
    End With
End Sub

Private Sub DrawSlotMachine()
Dim i As Long
Dim PosX As Long, PosY As Long, TexX As Long, TexY As Long

    With GUI(GuiEnum.GUI_SLOTMACHINE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height
        
        '//Buttons
        For i = ButtonEnum.SlotMachine_Close To ButtonEnum.SlotMachine_Close
            If CanShowButton(i) Then
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
            End If
        Next
    End With
End Sub

'//Editor
Private Sub DrawGDI()
    If frmEditor_Map.Visible And Editor = EDITOR_MAP Then
        GDITileset
    End If
    If frmEditor_Npc.Visible And Editor = EDITOR_NPC Then
        GDINpc
    End If
    If frmEditor_Pokemon.Visible And Editor = EDITOR_POKEMON Then
        GDIPokemon
    End If
    If frmEditor_Item.Visible And Editor = EDITOR_ITEM Then
        GDIItem
    End If
    If frmEditor_Animation.Visible And Editor = EDITOR_ANIMATION Then
        GDI_Animation
    End If
End Sub

Private Sub GDITileset()
Dim desRect As D3DRECT              '//Rendering Area
Dim scrlX As Long, scrlY As Long    '//Scrolling area
Dim oWidth As Long, oHeight As Long
Dim sWidth As Long, sHeight As Long

    With frmEditor_Map
        '//Exit if form is not open
        If Not .Visible Then Exit Sub
        
        If CurTileset <= 0 Or CurTileset > Count_Tileset Then Exit Sub
        
        '//Set Rendering Area
        scrlX = EditorScrollX
        scrlY = EditorScrollY
        oWidth = GetPicWidth(Tex_Tileset(CurTileset))
        oHeight = GetPicHeight(Tex_Tileset(CurTileset))
        desRect.x2 = .picTileset.scaleWidth
        desRect.Y2 = .picTileset.scaleHeight

        '//Start rendering
        D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorARGB(255, 255, 255, 255), 1#, 0
        D3DDevice.BeginScene

        If editorMapAnim > 0 And editorMapAnim <= Count_MapAnim Then
            RenderTexture Tex_MapAnim(editorMapAnim), 0, 0, PIC_X * MapFrameAnim, 0, TILE_X, TILE_Y, PIC_X, PIC_Y
        Else
            RenderTexture Tex_Tileset(CurTileset), 0, 0, scrlX * PIC_X, scrlY * PIC_Y, (oWidth * 2) - (scrlX * TILE_X), (oHeight * 2) - (scrlY * TILE_Y), oWidth - (scrlX * PIC_X), oHeight - (scrlY * PIC_Y)
            
            '//Selector
            '//Normal
            sWidth = EditorTileWidth * TILE_X
            sHeight = EditorTileHeight * TILE_Y
            RenderTexture Tex_System(gSystemEnum.UserInterface), (EditorTileX - EditorScrollX) * TILE_X, (EditorTileY - EditorScrollY) * TILE_Y, 0, 8, sWidth, sHeight, 1, 1, D3DColorARGB(100, 0, 0, 0)
        End If
        
        '//End the rendering
        D3DDevice.EndScene
        D3DDevice.Present desRect, desRect, .picTileset.hwnd, ByVal 0
    End With
End Sub

Private Sub GDINpc()
Dim desRect As D3DRECT              '//Rendering Area
Dim Sprite As Long
Dim oWidth As Long, oHeight As Long
Dim sWidth As Long, sHeight As Long

    With frmEditor_Npc
        '//Exit if form is not open
        If Not .Visible Then Exit Sub
        
        Sprite = .scrlSprite
        If Sprite <= 0 Or Sprite > Count_Character Then
            .picSprite.Cls
            Exit Sub
        End If
        
        oWidth = GetPicWidth(Tex_Character(Sprite))
        oHeight = GetPicHeight(Tex_Character(Sprite))
        desRect.x2 = .picSprite.scaleWidth
        desRect.Y2 = .picSprite.scaleHeight

        '//Start rendering
        D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorARGB(255, 240, 240, 240), 1#, 0
        D3DDevice.BeginScene

        RenderTexture Tex_Character(Sprite), 0, 0, oWidth / 3, 0, (oWidth / 3) * 2, (oHeight / 4) * 2, (oWidth / 3), (oHeight / 4)
        
        '//End the rendering
        D3DDevice.EndScene
        D3DDevice.Present desRect, desRect, .picSprite.hwnd, ByVal 0
    End With
End Sub

Private Sub GDIPokemon()
Dim desRect As D3DRECT              '//Rendering Area
Dim Sprite As Long
Dim oWidth As Long, oHeight As Long
Dim sWidth As Long, sHeight As Long
Dim Width As Long, Height As Long
Dim picX As Long, picY As Long

    With frmEditor_Pokemon
        '//Exit if form is not open
        If Not .Visible Then Exit Sub
        
        Sprite = .scrlSprite
        If Sprite <= 0 Or Sprite > Count_Pokemon Then
            .picSprite.Cls
            Exit Sub
        End If
        
        'oWidth = GetPicWidth(Tex_Pokemon(Sprite))
        'oHeight = GetPicHeight(Tex_Pokemon(Sprite))
        
        oWidth = GetPicWidth(Tex_Pokemon(Sprite))
        oHeight = GetPicHeight(Tex_Pokemon(Sprite)) / 2
        'Width = oWidth * 2: Height = oHeight * 2
        
        desRect.x2 = .picSprite.scaleWidth
        desRect.Y2 = .picSprite.scaleHeight
        
        'RenderTexture Tex_Pokemon(Sprite), ConvertMapX(X), ConvertMapY(Y), Anim * oWidth, rDir * oHeight, Width, Height, oWidth, oHeight
        
        '//Start rendering
        D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorARGB(255, 240, 240, 240), 1#, 0
        D3DDevice.BeginScene
        
        Width = (oWidth / 3) * 2
        Height = (oHeight) * 2
        
        picX = (desRect.x2 / 2) - (Width / 2)
        picY = (desRect.Y2 / 2) - (Height / 2) + 16
        RenderTexture Tex_Pokemon(Sprite), picX, picY, oWidth / 3, 0, Width, Height, oWidth / 3, oHeight
        
        'If .chkScale.value = YES Then
        '    Width = (oWidth / 3) * 2
        '    Height = (oHeight) * 2
            
        '    picX = (desRect.x2 / 2) - (Width / 2)
        '    picY = (desRect.Y2 / 2) - (Height / 2)

        '    RenderTexture Tex_Pokemon(Sprite), picX, picY, oWidth / 34, 0, Width, Height, oWidth / 34, oHeight
        'Else
        '    Width = (oWidth / 3)
        '    Height = (oHeight)
            
        '    picX = (desRect.x2 / 2) - (Width / 2)
        '    picY = (desRect.Y2 / 2) - (Height / 2)

        '    RenderTexture Tex_Pokemon(Sprite), picX, picY, oWidth / 34, 0, Width, Height, oWidth / 34, oHeight
        'End If
        
        '//End the rendering
        D3DDevice.EndScene
        D3DDevice.Present desRect, desRect, .picSprite.hwnd, ByVal 0
    End With
End Sub

Private Sub GDIItem()
Dim desRect As D3DRECT              '//Rendering Area
Dim Sprite As Long

    With frmEditor_Item
        '//Exit if form is not open
        If Not .Visible Then Exit Sub
        
        Sprite = .scrlSprite
        If Sprite <= 0 Or Sprite > Count_Item Then
            .picSprite.Cls
            Exit Sub
        End If
        
        desRect.x2 = .picSprite.scaleWidth
        desRect.Y2 = .picSprite.scaleHeight
        
        '//Start rendering
        D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorARGB(255, 240, 240, 240), 1#, 0
        D3DDevice.BeginScene
        
        RenderTexture Tex_Item(Sprite), 0, 0, 0, 0, 24, 24, 24, 24
        
        '//End the rendering
        D3DDevice.EndScene
        D3DDevice.Present desRect, desRect, .picSprite.hwnd, ByVal 0
    End With
End Sub

Private Sub GDI_Animation()
Dim AnimationNum As Long
Dim sX As Long, sY As Long
Dim i As Long
Dim Width As Long, Height As Long, destRect As D3DRECT
Dim looptime As Long
Dim FrameCount As Long
Dim ShouldRender As Boolean
    
    With frmEditor_Animation
        '//Exit if form is not open
        If Not .Visible Then Exit Sub
        
        For i = 0 To 1
            '//Set index
            AnimationNum = .scrlSprite(i).value
            If AnimationNum <= 0 Or AnimationNum > Count_Animation Then
                .picSprite(i).Cls
                GoTo continue
            End If
            
            destRect.x2 = .picSprite(i).scaleWidth
            destRect.Y2 = .picSprite(i).scaleHeight
        
            '//Start rendering
            D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorARGB(255, 240, 240, 240), 1#, 0
            D3DDevice.BeginScene

            If AnimationNum > 0 And AnimationNum <= Count_Animation Then
                looptime = .scrlLoopTime(i).value
                FrameCount = .scrlFrameCount(i).value
    
                If FrameCount > 0 Then
                    '//check if we need to render new frame
                    If AnimEditorTimer(i) + looptime <= GetTickCount Then
                        '//check if out of range
                        If AnimEditorFrame(i) >= FrameCount Then
                            AnimEditorFrame(i) = 1
                        Else
                            AnimEditorFrame(i) = AnimEditorFrame(i) + 1
                        End If
                        AnimEditorTimer(i) = GetTickCount + 25
                    End If
                
                    '//total width divided by frame count
                    Width = GetPicWidth(Tex_Animation(AnimationNum)) / frmEditor_Animation.scrlFrameCount(i).value 'AnimColumns 'GetPicWidth(Tex_Animation(AnimationNum)) '/ frmEditor_Animation.scrlFrameCount(i).value
                    Height = GetPicHeight(Tex_Animation(AnimationNum)) 'GetPicWidth(Tex_Animation(AnimationNum)) '/ AnimColumns 'GetPicHeight(Tex_Animation(AnimationNum))
        
                    sX = (AnimEditorFrame(i) - 1) * Width '(Width * (((AnimEditorFrame(i) - 1) Mod AnimColumns))) '(AnimEditorFrame(i) - 1) * Width
                    sY = 0 '(Height * ((AnimEditorFrame(i) - 1) \ AnimColumns)) '0

                    RenderTexture Tex_Animation(AnimationNum), 0, 0, sX, sY, Width, Height, Width, Height
                End If
            End If
            
            '//End the rendering
            D3DDevice.EndScene
            D3DDevice.Present destRect, destRect, .picSprite(i).hwnd, ByVal 0
continue:
        Next
    End With
End Sub

Public Sub DrawShopItemDesc()
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
    
    If ShopNum = 0 Then Exit Sub
    If ShopItemDesc <= 0 Or ShopItemDesc > MAX_SHOP_ITEM Then Exit Sub
    If ShopItemDescTimer + 400 > GetTickCount Then Exit Sub
    ShopItemDescShow = True

    ItemIcon = Item(Shop(ShopNum).ShopItem(ShopItemDesc).Num).Sprite
    ItemName = "~ " & Trim$(Item(Shop(ShopNum).ShopItem(ShopItemDesc).Num).Name) & " ~"
    DescString = Trim$(Item(Shop(ShopNum).ShopItem(ShopItemDesc).Num).Desc)    '"A device for catching wild Pokemon. It is thrown like a ball at the target. It is designed as a capsule system"

    'If Item(Shop(ShopNum).ShopItem(ShopItemDesc).Num).IsCash = NO Then
    '    ItemPrice = "Price: " & Item(PlayerInv(InvItemDesc).Num).Price
    'Else
    '    ItemPrice = "Price: Non Sellable"
    'End If

    '//Make sure that loading text have something to draw
    If Len(DescString) < 0 Then Exit Sub

    '//Wrap the text
    WordWrap_Array Font_Default, DescString, 150, ArrayText

    '//we need these often
    LowBound = LBound(ArrayText)
    UpBound = UBound(ArrayText)

    SizeY = 25 + ((UpBound + 1) * 16)

    RenderTexture Tex_System(gSystemEnum.UserInterface), CursorX, CursorY, 0, 8, 182, 219, 1, 1, D3DColorARGB(180, 0, 0, 0)

    RenderTexture Tex_Item(ItemIcon), CursorX + GUI(GuiEnum.GUI_INVENTORY).Width / 2 - (GetPicHeight(Tex_Item(ItemIcon)) / 2), CursorY + 8 + ((219 * 0.5) - (SizeY * 0.5)), 0, 0, GetPicWidth(Tex_Item(ItemIcon)), GetPicHeight(Tex_Item(ItemIcon)), GetPicWidth(Tex_Item(ItemIcon)), GetPicHeight(Tex_Item(ItemIcon))

    RenderText Font_Default, ItemName, CursorX + 6 + ((182 * 0.5) - (GetTextWidth(Font_Default, ItemName) * 0.5)), CursorY + 36 + ((219 * 0.5) - (SizeY * 0.5)), White
    
    'RenderText Font_Default, ItemPrice, GUI(GuiEnum.GUI_INVENTORY).X + 6 + ((182 * 0.5) - (GetTextWidth(Font_Default, ItemName) * 0.5)), GUI(GuiEnum.GUI_INVENTORY).Y + 150 + ((219 * 0.5) - (SizeY * 0.5)), White

    '//Reset
    yOffset = 25
    '//Loop to all items
    For i = LowBound To UpBound
        '//Set Location
        '//Keep it centered
        X = CursorX + 6 + ((182 * 0.5) - (GetTextWidth(Font_Default, Trim$(ArrayText(i))) * 0.5))
        Y = CursorY + 36 + ((219 * 0.5) - (SizeY * 0.5)) + yOffset

        '//Render the text
        RenderText Font_Default, Trim$(ArrayText(i)), X, Y, White

        '//Increase the location for each line
        yOffset = yOffset + 16
    Next
End Sub

Public Sub DrawStorageItemDesc()
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
    
    If StorageType <> 1 Then Exit Sub
    If StorageItemDesc <= 0 Or StorageItemDesc > MAX_STORAGE Then Exit Sub
    If StorageItemDescTimer + 400 > GetTickCount Then Exit Sub
    StorageItemDescShow = True

    ItemIcon = Item(PlayerInvStorage(InvCurSlot).Data(StorageItemDesc).Num).Sprite
    ItemName = "~ " & Trim$(Item(PlayerInvStorage(InvCurSlot).Data(StorageItemDesc).Num).Name) & " ~"
    DescString = Trim$(Item(PlayerInvStorage(InvCurSlot).Data(StorageItemDesc).Num).Desc)    '"A device for catching wild Pokemon. It is thrown like a ball at the target. It is designed as a capsule system"

    '//Make sure that loading text have something to draw
    If Len(DescString) < 0 Then Exit Sub

    '//Wrap the text
    WordWrap_Array Font_Default, DescString, 150, ArrayText

    '//we need these often
    LowBound = LBound(ArrayText)
    UpBound = UBound(ArrayText)

    SizeY = 45 + ((UpBound + 1) * 16)

    RenderTexture Tex_System(gSystemEnum.UserInterface), CursorX, CursorY, 0, 8, 182, 219, 1, 1, D3DColorARGB(180, 0, 0, 0)

    RenderTexture Tex_Item(ItemIcon), CursorX + GUI(GuiEnum.GUI_INVENTORY).Width / 2 - (GetPicHeight(Tex_Item(ItemIcon)) / 2), CursorY + 8 + ((219 * 0.5) - (SizeY * 0.5)), 0, 0, GetPicWidth(Tex_Item(ItemIcon)), GetPicHeight(Tex_Item(ItemIcon)), GetPicWidth(Tex_Item(ItemIcon)), GetPicHeight(Tex_Item(ItemIcon))

    RenderText Font_Default, ItemName, CursorX + 6 + ((182 * 0.5) - (GetTextWidth(Font_Default, ItemName) * 0.5)), CursorY + 36 + ((219 * 0.5) - (SizeY * 0.5)), White
    
    If Item(PlayerInvStorage(InvCurSlot).Data(StorageItemDesc).Num).IsCash = NO Then
        ItemPrice = "Price: " & Int((Item(PlayerInvStorage(InvCurSlot).Data(StorageItemDesc).Num).Price / 2))
        RenderTexture Tex_Item(IDMoney), CursorX + ((150 * 0.5) - (GetTextWidth(Font_Default, ItemPrice) * 0.5)), CursorY + 120 + ((219 * 0.5) - (SizeY * 0.5)), 0, 0, 20, 20, GetPicWidth(Tex_Item(IDMoney)), GetPicHeight(Tex_Item(IDMoney))
    Else
        ItemPrice = "Price: Non Sellable"
    End If
    
    RenderText Font_Default, ItemPrice, CursorX + 6 + ((182 * 0.5) - (GetTextWidth(Font_Default, ItemPrice) * 0.5)), CursorY + 120 + ((219 * 0.5) - (SizeY * 0.5)), White
    
    '//Reset
    yOffset = 25
    '//Loop to all items
    For i = LowBound To UpBound
        '//Set Location
        '//Keep it centered
        X = CursorX + 6 + ((182 * 0.5) - (GetTextWidth(Font_Default, Trim$(ArrayText(i))) * 0.5))
        Y = CursorY + 36 + ((219 * 0.5) - (SizeY * 0.5)) + yOffset

        '//Render the text
        RenderText Font_Default, Trim$(ArrayText(i)), X, Y, White

        '//Increase the location for each line
        yOffset = yOffset + 16
    Next
End Sub
