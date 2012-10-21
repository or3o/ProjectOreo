Attribute VB_Name = "modGraphics"
Option Explicit
' **********************
' ** Renders graphics **
' **********************
' DirectX8 Object
Private DirectX8 As DirectX8 'The master DirectX object.
Private Direct3D As Direct3D8 'Controls all things 3D.
Public Direct3D_Device As Direct3DDevice8 'Represents the hardware rendering.
Private Direct3DX As D3DX8

'The 2D (Transformed and Lit) vertex format.
Private Const FVF_TLVERTEX As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE

'The 2D (Transformed and Lit) vertex format type.
Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    RHW As Single
    color As Long
    TU As Single
    TV As Single
End Type

Private Vertex_List(3) As TLVERTEX '4 vertices will make a square.

'Some color depth constants to help make the DX constants more readable.
Private Const COLOR_DEPTH_16_BIT As Long = D3DFMT_R5G6B5
Private Const COLOR_DEPTH_24_BIT As Long = D3DFMT_A8R8G8B8
Private Const COLOR_DEPTH_32_BIT As Long = D3DFMT_X8R8G8B8

Public RenderingMode As Long

Public Direct3D_Window As D3DPRESENT_PARAMETERS 'Backbuffer and viewport description.
Private Display_Mode As D3DDISPLAYMODE

'Graphic Textures
Public Tex_Item() As DX8TextureRec ' arrays
Public Tex_Character() As DX8TextureRec
Public Tex_Paperdoll() As DX8TextureRec
Public Tex_Tileset() As DX8TextureRec
Public Tex_Resource() As DX8TextureRec
Public Tex_Animation() As DX8TextureRec
Public Tex_SpellIcon() As DX8TextureRec
Public Tex_Face() As DX8TextureRec
Public Tex_Projectile() As DX8TextureRec
Public Tex_Fog() As DX8TextureRec
Public Tex_Door As DX8TextureRec ' singes
Public Tex_Blood As DX8TextureRec
Public Tex_Misc As DX8TextureRec
Public Tex_Direction As DX8TextureRec
Public Tex_Target As DX8TextureRec
Public Tex_Bars As DX8TextureRec
Public Tex_Selection As DX8TextureRec
Public Tex_White As DX8TextureRec
Public Tex_Weather As DX8TextureRec
Public Tex_ChatBubble As DX8TextureRec
Public Tex_Fade As DX8TextureRec
Public Tex_Shadow As DX8TextureRec


' Number of graphic files
Public NumTileSets As Long
Public NumCharacters As Long
Public NumPaperdolls As Long
Public numitems As Long
Public NumResources As Long
Public NumAnimations As Long
Public NumSpellIcons As Long
Public NumFaces As Long
Public NumFogs As Long
Public NumProjectiles As Long

Public Type DX8TextureRec
    Texture As Long
    Width As Long
    Height As Long
    filepath As String
    TexWidth As Long
    TexHeight As Long
    ImageData() As Byte
    HasData As Boolean
End Type

Public Type GlobalTextureRec
    Texture As Direct3DTexture8
    TexWidth As Long
    TexHeight As Long
    timer As Long
End Type

Public Type RECT
    Top As Long
    Left As Long
    Bottom As Long
    Right As Long
End Type

Public Const SurfaceTimer As Long = 8000
Public gTexture() As GlobalTextureRec
Public NumTextures As Long

' ********************
' ** Initialization **
' ********************
Public Function InitDX8() As Boolean
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    Set DirectX8 = New DirectX8 'Creates the DirectX object.
    Set Direct3D = DirectX8.Direct3DCreate() 'Creates the Direct3D object using the DirectX object.
    Set Direct3DX = New D3DX8
    
    Direct3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display_Mode 'Use the current display mode that you
                                                                    'are already on. Incase you are confused, I'm
                                                                    'talking about your current screen resolution. ;)
    Direct3D_Window.Windowed = True 'The app will be in windowed mode.
    
    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_DISCARD 'Refresh when the monitor does.
    Direct3D_Window.BackBufferFormat = Display_Mode.Format 'Sets the format that was retrieved into the backbuffer.
    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_COPY
    Direct3D_Window.BackBufferCount = 1 '1 backbuffer only
    Direct3D_Window.BackBufferWidth = 960 ' frmMain.picScreen.ScaleWidth 'Match the backbuffer width with the display width
    Direct3D_Window.BackBufferHeight = 768 'frmMain.picScreen.ScaleHeight 'Match the backbuffer height with the display height
    Direct3D_Window.hDeviceWindow = frmMain.picScreen.hWnd 'Use frmMain as the device window.
    
    'we've already setup for Direct3D_Window.
    If TryCreateDirectX8Device = False Then
        MsgBox "Unable to initialize DirectX8. You may be missing dx8vb.dll or have incompatible hardware to use DirectX8."
        DestroyGame
    End If

    With Direct3D_Device
        .SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
    
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
        
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
        ' *** 04/09/2012 DD/MM/YYYY : Disabled MipMapping below. ***
        .SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_NONE
    End With
    
    ' Initialise the surfaces
    LoadTextures
    
    ' We're done
    InitDX8 = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "InitDX8", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
Function TryCreateDirectX8Device() As Boolean
Dim i As Long

On Error GoTo nexti

    For i = 1 To 4
        Select Case i
            Case 1
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.picScreen.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            Case 2
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.picScreen.hWnd, D3DCREATE_MIXED_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            Case 3
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.picScreen.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            Case 4
                TryCreateDirectX8Device = False
                Exit Function
        End Select
nexti:
    Next

End Function

Function GetNearestPOT(Value As Long) As Long
Dim i As Long
    Do While 2 ^ i < Value
        i = i + 1
    Loop
    GetNearestPOT = 2 ^ i
End Function

Public Sub LoadTexture(ByRef TextureRec As DX8TextureRec)
Dim SourceBitmap As cGDIpImage, ConvertedBitmap As cGDIpImage, GDIGraphics As cGDIpRenderer, GDIToken As cGDIpToken, i As Long
Dim newWidth As Long, newHeight As Long, ImageData() As Byte, fn As Long
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If TextureRec.HasData = False Then
        Set GDIToken = New cGDIpToken
        If GDIToken.Token = 0& Then MsgBox "GDI+ failed to load, exiting game!": DestroyGame
        Set SourceBitmap = New cGDIpImage
        Call SourceBitmap.LoadPicture_FileName(TextureRec.filepath, GDIToken)
        
        TextureRec.Width = SourceBitmap.Width
        TextureRec.Height = SourceBitmap.Height
        
        newWidth = GetNearestPOT(TextureRec.Width)
        newHeight = GetNearestPOT(TextureRec.Height)
        If newWidth <> SourceBitmap.Width Or newHeight <> SourceBitmap.Height Then
            Set ConvertedBitmap = New cGDIpImage
            Set GDIGraphics = New cGDIpRenderer
            i = GDIGraphics.CreateGraphicsFromImageClass(SourceBitmap)
            Call ConvertedBitmap.LoadPicture_FromNothing(newHeight, newWidth, i, GDIToken) 'I HAVE NO IDEA why this is backwards but it works.
            Call GDIGraphics.DestroyHGraphics(i)
            i = GDIGraphics.CreateGraphicsFromImageClass(ConvertedBitmap)
            Call GDIGraphics.AttachTokenClass(GDIToken)
            Call GDIGraphics.RenderImageClassToHGraphics(SourceBitmap, i)
            Call ConvertedBitmap.SaveAsPNG(ImageData)
            GDIGraphics.DestroyHGraphics (i)
            TextureRec.ImageData = ImageData
            Set ConvertedBitmap = Nothing
            Set GDIGraphics = Nothing
            Set SourceBitmap = Nothing
        Else
            Call SourceBitmap.SaveAsPNG(ImageData)
            TextureRec.ImageData = ImageData
            Set SourceBitmap = Nothing
        End If
    Else
        ImageData = TextureRec.ImageData
    End If
    
    
    Set gTexture(TextureRec.Texture).Texture = Direct3DX.CreateTextureFromFileInMemoryEx(Direct3D_Device, _
                                                    ImageData(0), _
                                                    UBound(ImageData) + 1, _
                                                    newWidth, _
                                                    newHeight, _
                                                    D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, ByVal (0), ByVal 0, ByVal 0)
    
    gTexture(TextureRec.Texture).TexWidth = newWidth
    gTexture(TextureRec.Texture).TexHeight = newHeight
    Exit Sub
errorhandler:
    HandleError "LoadTexture", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub LoadTextures()
Dim i As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    Call CheckTilesets
    Call CheckCharacters
    Call CheckPaperdolls
    Call CheckAnimations
    Call CheckItems
    Call CheckResources
    Call CheckSpellIcons
    Call CheckFaces
    Call CheckFogs
    Call CheckProjectiles
    
    NumTextures = NumTextures + 12
    
    ReDim Preserve gTexture(NumTextures)
    Tex_Fade.filepath = App.Path & "\data files\graphics\misc\fader.png"
    Tex_Fade.Texture = NumTextures - 10
    'LoadTexture Tex_Fade
    Tex_ChatBubble.filepath = App.Path & "\data files\graphics\misc\chatbubble.png"
    Tex_ChatBubble.Texture = NumTextures - 9
    'LoadTexture Tex_ChatBubble
    Tex_Weather.filepath = App.Path & "\data files\graphics\misc\weather.png"
    Tex_Weather.Texture = NumTextures - 8
    'LoadTexture Tex_Weather
    Tex_White.filepath = App.Path & "\data files\graphics\misc\white.png"
    Tex_White.Texture = NumTextures - 7
    'LoadTexture Tex_White
    Tex_Door.filepath = App.Path & "\data files\graphics\misc\door.png"
    Tex_Door.Texture = NumTextures - 6
    'LoadTexture Tex_Door
    Tex_Direction.filepath = App.Path & "\data files\graphics\misc\direction.png"
    Tex_Direction.Texture = NumTextures - 5
    'LoadTexture Tex_Direction
    Tex_Target.filepath = App.Path & "\data files\graphics\misc\target.png"
    Tex_Target.Texture = NumTextures - 4
    'LoadTexture Tex_Target
    Tex_Misc.filepath = App.Path & "\data files\graphics\misc\misc.png"
    Tex_Misc.Texture = NumTextures - 3
    'LoadTexture Tex_Misc
    Tex_Blood.filepath = App.Path & "\data files\graphics\misc\blood.png"
    Tex_Blood.Texture = NumTextures - 2
    'LoadTexture Tex_Blood
    Tex_Bars.filepath = App.Path & "\data files\graphics\misc\bars.png"
    Tex_Bars.Texture = NumTextures - 1
    'LoadTexture Tex_Bars
    Tex_Selection.filepath = App.Path & "\data files\graphics\misc\select.png"
    Tex_Selection.Texture = NumTextures
    'LoadTexture Tex_Selection
    Tex_Shadow.filepath = App.Path & "\data files\graphics\misc\shadow.png"
    Tex_Shadow.Texture = NumTextures - 11
    'LoadTexture Tex_Shadow
    EngineInitFontTextures
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadTextures", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UnloadTextures()
Dim i As Long
    
    ' If debug mode, handle error then exit out
    On Error Resume Next
    
    For i = 1 To NumTextures
        Set gTexture(i).Texture = Nothing
        ZeroMemory ByVal VarPtr(gTexture(i)), LenB(gTexture(i))
    Next
    
    ReDim gTexture(1)
    
    For i = 1 To NumTileSets
        Tex_Tileset(i).Texture = 0
    Next

    For i = 1 To numitems
        Tex_Item(i).Texture = 0
    Next

    For i = 1 To NumCharacters
        Tex_Character(i).Texture = 0
    Next
    
    For i = 1 To NumPaperdolls
        Tex_Paperdoll(i).Texture = 0
    Next
    
    For i = 1 To NumResources
        Tex_Resource(i).Texture = 0
    Next
    
    For i = 1 To NumAnimations
        Tex_Animation(i).Texture = 0
    Next
    
    For i = 1 To NumSpellIcons
        Tex_SpellIcon(i).Texture = 0
    Next
    
    For i = 1 To NumFaces
        Tex_Face(i).Texture = 0
    Next
    
    For i = 1 To NumProjectiles
        Tex_Projectile(i).Texture = 0
    Next
    
    Tex_Misc.Texture = 0
    Tex_Blood.Texture = 0
    Tex_Door.Texture = 0
    Tex_Direction.Texture = 0
    Tex_Target.Texture = 0
    Tex_Selection.Texture = 0
    Tex_Shadow.Texture = 0
    UnloadFontTextures
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UnloadTextures", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
' **************
' ** Drawing **
' **************
Public Sub RenderTexture(ByRef TextureRec As DX8TextureRec, ByVal dX As Single, ByVal dY As Single, ByVal sX As Single, ByVal sY As Single, ByVal dWidth As Single, ByVal dHeight As Single, ByVal sWidth As Single, ByVal sHeight As Single, Optional color As Long = -1)
    Dim TextureNum As Long
    Dim textureWidth As Long, textureHeight As Long, sourceX As Single, sourceY As Single, sourceWidth As Single, sourceHeight As Single
    TextureNum = TextureRec.Texture
    
    textureWidth = gTexture(TextureNum).TexWidth
    textureHeight = gTexture(TextureNum).TexHeight
    
    If sY + sHeight > textureHeight Then Exit Sub
    If sX + sWidth > textureWidth Then Exit Sub
    If sX < 0 Then Exit Sub
    If sY < 0 Then Exit Sub

    If gTexture(TextureNum).timer = 0 Then
          LoadTexture TextureRec
          AddText "Loaded texture: " & TextureNum, White
          'Sleep 100
        End If
  gTexture(TextureNum).timer = timeGetTime + SurfaceTimer
  
    sX = sX - 0.5
    sY = sY - 0.5
    dY = dY - 0.5
    dX = dX - 0.5
    sWidth = sWidth
    sHeight = sHeight
    dWidth = dWidth
    dHeight = dHeight
    If textureWidth = 0 Then
        Exit Sub
     End If
    sourceX = (sX / textureWidth)
    sourceY = (sY / textureHeight)
    sourceWidth = ((sX + sWidth) / textureWidth)
    sourceHeight = ((sY + sHeight) / textureHeight)
    
    Vertex_List(0) = Create_TLVertex(dX, dY, 0, 1, color, 0, sourceX + 0.000003, sourceY + 0.000003)
    Vertex_List(1) = Create_TLVertex(dX + dWidth, dY, 0, 1, color, 0, sourceWidth + 0.000003, sourceY + 0.000003)
    Vertex_List(2) = Create_TLVertex(dX, dY + dHeight, 0, 1, color, 0, sourceX + 0.000003, sourceHeight + 0.000003)
    Vertex_List(3) = Create_TLVertex(dX + dWidth, dY + dHeight, 0, 1, color, 0, sourceWidth + 0.000003, sourceHeight + 0.000003)
    
    Direct3D_Device.SetTexture 0, gTexture(TextureNum).Texture
    Direct3D_Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex_List(0), Len(Vertex_List(0))
End Sub

Public Sub RenderTextureByRects(TextureRec As DX8TextureRec, sRect As RECT, dRect As RECT)
    Dim TextureNum As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    ' No idea why a duplicate is actually needed here, seeing as it heads to RenderTexture, but, it just works.
    ' Leave it at that.
    TextureNum = TextureRec.Texture
    If gTexture(TextureNum).timer = 0 Then
       LoadTexture TextureRec
       AddText "Loaded texture: " & TextureNum, White
       'Sleep 100
    End If
    gTexture(TextureNum).timer = timeGetTime + SurfaceTimer
    
    RenderTexture TextureRec, dRect.Left, dRect.Top, sRect.Left, sRect.Top, dRect.Right - dRect.Left, dRect.Bottom - dRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderTextureByRects", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawDirection(ByVal X As Long, ByVal Y As Long)
Dim rec As RECT
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    ' render grid
    rec.Top = 24
    rec.Left = 0
    rec.Right = rec.Left + 32
    rec.Bottom = rec.Top + 32
    RenderTexture Tex_Direction, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' render dir blobs
    For i = 1 To 4
        rec.Left = (i - 1) * 8
        rec.Right = rec.Left + 8
        ' find out whether render blocked or not
        If Not isDirBlocked(Map.Tile(X, Y).DirBlock, CByte(i)) Then
            rec.Top = 8
        Else
            rec.Top = 16
        End If
        rec.Bottom = rec.Top + 8
        'render!
        RenderTexture Tex_Direction, ConvertMapX(X * PIC_X) + DirArrowX(i), ConvertMapY(Y * PIC_Y) + DirArrowY(i), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawDirection", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawTarget(ByVal X As Long, ByVal Y As Long)
Dim sRect As RECT
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Tex_Target.Texture = 0 Then Exit Sub
    
    Width = Tex_Target.Width / 2
    Height = Tex_Target.Height

    With sRect
        .Top = 0
        .Bottom = Height
        .Left = 0
        .Right = Width
    End With
    
    X = X - ((Width - 32) / 2)
    Y = Y - (Height / 2)
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    ' clipping
    If Y < 0 Then
        With sRect
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With sRect
            .Left = .Left - X
        End With
        X = 0
    End If
    
    RenderTexture Tex_Target, X, Y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawTarget", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawHover(ByVal tType As Long, ByVal target As Long, ByVal X As Long, ByVal Y As Long)
Dim sRect As RECT
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Tex_Target.Texture = 0 Then Exit Sub
    
    Width = Tex_Target.Width / 2
    Height = Tex_Target.Height

    With sRect
        .Top = 0
        .Bottom = Height
        .Left = Width
        .Right = .Left + Width
    End With
    
    X = X - ((Width - 32) / 2)
    Y = Y - (Height / 2)

    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    ' clipping
    If Y < 0 Then
        With sRect
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With sRect
            .Left = .Left - X
        End With
        X = 0
    End If
    
    RenderTexture Tex_Target, X, Y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawHover", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawMapTile(ByVal X As Long, ByVal Y As Long)
Dim rec As RECT
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    With Map.Tile(X, Y)
        For i = MapLayer.Ground To MapLayer.Mask2
            If Autotile(X, Y).Layer(i).renderState = RenderNormal Then
                ' Draw normally
                RenderTexture Tex_Tileset(.Layer(i).Tileset), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), .Layer(i).X * 32, .Layer(i).Y * 32, 32, 32, 32, 32, -1
              ElseIf Autotile(X, Y).Layer(i).renderState = RenderAutotile Then
                ' Draw autotiles
                DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), 1, X, Y
                DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY(Y * PIC_Y), 2, X, Y
                DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY((Y * PIC_Y) + 16), 3, X, Y
                DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY((Y * PIC_Y) + 16), 4, X, Y
 End If
            
        Next
    End With
    
    ' Error handler
    Exit Sub
    
errorhandler:
    HandleError "DrawMapTile", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawMapFringeTile(ByVal X As Long, ByVal Y As Long)
Dim rec As RECT
Dim i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    With Map.Tile(X, Y)
        For i = MapLayer.Fringe To MapLayer.Fringe2
            If Autotile(X, Y).Layer(i).renderState = RenderNormal Then
                ' Draw normally
                RenderTexture Tex_Tileset(.Layer(i).Tileset), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), .Layer(i).X * 32, .Layer(i).Y * 32, 32, 32, 32, 32, -1
            ElseIf Autotile(X, Y).Layer(i).renderState = RenderAutotile Then
                ' Draw autotiles
                DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), 1, X, Y
                DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY(Y * PIC_Y), 2, X, Y
                DrawAutoTile i, ConvertMapX(X * PIC_X), ConvertMapY((Y * PIC_Y) + 16), 3, X, Y
                DrawAutoTile i, ConvertMapX((X * PIC_X) + 16), ConvertMapY((Y * PIC_Y) + 16), 4, X, Y
            End If
        Next
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawMapFringeTile", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawDoor(ByVal X As Long, ByVal Y As Long)
Dim rec As RECT
Dim x2 As Long, y2 As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    ' sort out animation
    With TempTile(X, Y)
        If .DoorAnimate = 1 Then ' opening
            If .DoorTimer + 100 < timeGetTime Then
                If .DoorFrame < 4 Then
                    .DoorFrame = .DoorFrame + 1
                Else
                    .DoorAnimate = 2 ' set to closing
                End If
                .DoorTimer = timeGetTime
            End If
        ElseIf .DoorAnimate = 2 Then ' closing
            If .DoorTimer + 100 < timeGetTime Then
                If .DoorFrame > 1 Then
                    .DoorFrame = .DoorFrame - 1
                Else
                    .DoorAnimate = 0 ' end animation
                End If
                .DoorTimer = timeGetTime
            End If
        End If
        
        If .DoorFrame = 0 Then .DoorFrame = 1
    End With

    With rec
        .Top = 0
        .Bottom = Tex_Door.Height
        .Left = ((TempTile(X, Y).DoorFrame - 1) * (Tex_Door.Width / 4))
        .Right = .Left + (Tex_Door.Width / 4)
    End With

    x2 = (X * PIC_X)
    y2 = (Y * PIC_Y) - (Tex_Door.Height / 2) + 4
    RenderTexture Tex_Door, ConvertMapX(x2), ConvertMapY(y2), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    'Call DDS_BackBuffer.DrawFast(ConvertMapX(X2), ConvertMapY(Y2), DDS_Door, rec, DDDrawFAST_WAIT Or DDDrawFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawDoor", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawBlood(ByVal Index As Long)
Dim rec As RECT
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    'load blood then
    BloodCount = Tex_Blood.Width / 32
    
    With Blood(Index)
        ' check if we should be seeing it
        If .timer + 20000 < timeGetTime Then Exit Sub
        
        rec.Top = 0
        rec.Bottom = PIC_Y
        rec.Left = (.Sprite - 1) * PIC_X
        rec.Right = rec.Left + PIC_X
        RenderTexture Tex_Blood, ConvertMapX(.X * PIC_X), ConvertMapY(.Y * PIC_Y), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawBlood", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawAnimation(ByVal Index As Long, ByVal Layer As Long)
Dim Sprite As Long
Dim sRect As RECT
Dim dRect As RECT
Dim i As Long
Dim Width As Long, Height As Long
Dim looptime As Long
Dim FrameCount As Long
Dim X As Long, Y As Long
Dim lockindex As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If AnimInstance(Index).Animation = 0 Then
        ClearAnimInstance Index
        Exit Sub
    End If
    
    Sprite = Animation(AnimInstance(Index).Animation).Sprite(Layer)
    
    If Sprite < 1 Or Sprite > NumAnimations Then Exit Sub
    
    FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
    
    ' total width divided by frame count
    Width = Tex_Animation(Sprite).Width / FrameCount
    Height = Tex_Animation(Sprite).Height
    
    sRect.Top = 0
    sRect.Bottom = Height
    sRect.Left = (AnimInstance(Index).frameIndex(Layer) - 1) * Width
    sRect.Right = sRect.Left + Width
    
    ' change x or y if locked
    If AnimInstance(Index).LockType > TargetNone Then ' if <> none
        ' is a player
        If AnimInstance(Index).LockType = TargetPlayer Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if is ingame
            If IsPlaying(lockindex) Then
                ' check if on same map
                If GetPlayerMap(lockindex) = GetPlayerMap(MyIndex) Then
                    ' is on map, is playing, set x & y
                    X = (GetPlayerX(lockindex) * PIC_X) + 16 - (Width / 2) + Player(lockindex).xOffset
                    Y = (GetPlayerY(lockindex) * PIC_Y) + 16 - (Height / 2) + Player(lockindex).yOffset
                End If
            End If
        ElseIf AnimInstance(Index).LockType = TargetNPC Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if NPC exists
            If MapNpc(lockindex).Num > 0 Then
                ' check if alive
                If MapNpc(lockindex).Vital(Vitals.HP) > 0 Then
                    ' exists, is alive, set x & y
                    X = (MapNpc(lockindex).X * PIC_X) + 16 - (Width / 2) + MapNpc(lockindex).xOffset
                    Y = (MapNpc(lockindex).Y * PIC_Y) + 16 - (Height / 2) + MapNpc(lockindex).yOffset
                Else
                    ' npc not alive anymore, kill the animation
                    ClearAnimInstance Index
                    Exit Sub
                End If
            Else
                ' npc not alive anymore, kill the animation
                ClearAnimInstance Index
                Exit Sub
            End If
        End If
    Else
        ' no lock, default x + y
        X = (AnimInstance(Index).X * 32) + 16 - (Width / 2)
        Y = (AnimInstance(Index).Y * 32) + 16 - (Height / 2)
    End If
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)

    ' Clip to screen
    If Y < 0 Then

        With sRect
            .Top = .Top - Y
        End With

        Y = 0
    End If

    If X < 0 Then

        With sRect
            .Left = .Left - X
        End With

        X = 0
    End If
    
    RenderTexture Tex_Animation(Sprite), X, Y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawAnimation", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawItem(ByVal itemnum As Long)
Dim PicNum As Long
Dim rec As RECT
Dim MaxFrames As Byte

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    ' if it's not us then don't render
    If MapItem(itemnum).playerName <> vbNullString Then
        If MapItem(itemnum).playerName <> Trim$(GetPlayerName(MyIndex)) Then Exit Sub
    End If
    
    ' get the picture
    PicNum = Item(MapItem(itemnum).Num).Pic

    If PicNum < 1 Or PicNum > numitems Then Exit Sub

    If Tex_Item(PicNum).Width > 64 Then ' has more than 1 frame
        With rec
            .Top = 0
            .Bottom = 32
            .Left = (MapItem(itemnum).Frame * 32)
            .Right = .Left + 32
        End With
    Else
        With rec
            .Top = 0
            .Bottom = PIC_Y
            .Left = 0
            .Right = PIC_X
        End With
    End If
    
    RenderTexture Tex_Item(PicNum), ConvertMapX(MapItem(itemnum).X * PIC_X), ConvertMapY(MapItem(itemnum).Y * PIC_Y), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawItem", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Public Sub DrawMapResource(ByVal Resource_num As Long, Optional ByVal screenShot As Boolean = False)
Dim Resource_master As Long
Dim Resource_state As Long
Dim Resource_sprite As Long
Dim rec As RECT
Dim X As Long, Y As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    ' make sure it's not out of map
    If MapResource(Resource_num).X > Map.MaxX Then Exit Sub
    If MapResource(Resource_num).Y > Map.MaxY Then Exit Sub
    
    ' Get the Resource type
    Resource_master = Map.Tile(MapResource(Resource_num).X, MapResource(Resource_num).Y).Data1
    
    If Resource_master = 0 Then Exit Sub

    If Resource(Resource_master).ResourceImage = 0 Then Exit Sub
    ' Get the Resource state
    Resource_state = MapResource(Resource_num).ResourceState

    If Resource_state = 0 Then ' normal
        Resource_sprite = Resource(Resource_master).ResourceImage
    ElseIf Resource_state = 1 Then ' used
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If
    
    ' cut down everything if we're editing
    If InMapEditor Then
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If

    ' src rect
    With rec
        .Top = 0
        .Bottom = Tex_Resource(Resource_sprite).Height
        .Left = 0
        .Right = Tex_Resource(Resource_sprite).Width
    End With

    ' Set base x + y, then the offset due to size
    X = (MapResource(Resource_num).X * PIC_X) - (Tex_Resource(Resource_sprite).Width / 2) + 16
    Y = (MapResource(Resource_num).Y * PIC_Y) - Tex_Resource(Resource_sprite).Height + 32
    
    ' render it
    If Not screenShot Then
        Call DrawResource(Resource_sprite, X, Y, rec)
    Else
        Exit Sub
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawMapResource", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawResource(ByVal Resource As Long, ByVal dX As Long, dY As Long, rec As RECT)
Dim X As Long
Dim Y As Long
Dim Width As Long
Dim Height As Long
Dim destRect As RECT

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Resource < 1 Or Resource > NumResources Then Exit Sub

    X = ConvertMapX(dX)
    Y = ConvertMapY(dY)
    
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)
    
    RenderTexture Tex_Resource(Resource), X, Y, rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawResource", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub DrawBars()
Dim tmpY As Long, tmpX As Long
Dim sWidth As Long, sHeight As Long
Dim sRect As RECT
Dim barWidth As Long
Dim i As Long, npcNum As Long, partyIndex As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    ' dynamic bar calculations
    sWidth = Tex_Bars.Width
    sHeight = Tex_Bars.Height / 4
    
    ' render health bars
    For i = 1 To MAX_MAP_NPCS
        npcNum = MapNpc(i).Num
        ' exists?
        If npcNum > 0 Then
            ' alive?
            If MapNpc(i).Vital(Vitals.HP) > 0 And MapNpc(i).Vital(Vitals.HP) < NPC(npcNum).HP Then
                ' lock to npc
                tmpX = MapNpc(i).X * PIC_X + MapNpc(i).xOffset + 16 - (sWidth / 2)
                tmpY = MapNpc(i).Y * PIC_Y + MapNpc(i).yOffset + 35
                
                ' calculate the width to fill
                barWidth = ((MapNpc(i).Vital(Vitals.HP) / sWidth) / (NPC(npcNum).HP / sWidth)) * sWidth
                
                ' draw bar background
                With sRect
                    .Top = sHeight * 1 ' HP bar background
                    .Left = 0
                    .Right = .Left + sWidth
                    .Bottom = .Top + sHeight
                End With
                RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                
                ' draw the bar proper
                With sRect
                    .Top = 0 ' HP bar
                    .Left = 0
                    .Right = .Left + barWidth
                    .Bottom = .Top + sHeight
                End With
                RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
            End If
        End If
    Next

    ' check for casting time bar
    If SpellBuffer > 0 Then
        If Spell(PlayerSpells(SpellBuffer)).CastTime > 0 Then
            ' lock to player
            tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).xOffset + 16 - (sWidth / 2)
            tmpY = GetPlayerY(MyIndex) * PIC_Y + Player(MyIndex).yOffset + 35 + sHeight + 1
            
            ' calculate the width to fill
            barWidth = (timeGetTime - SpellBufferTimer) / ((Spell(PlayerSpells(SpellBuffer)).CastTime * 1000)) * sWidth
            
            ' draw bar background
            With sRect
                .Top = sHeight * 3 ' cooldown bar background
                .Left = 0
                .Right = sWidth
                .Bottom = .Top + sHeight
            End With
            RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
            
            ' draw the bar proper
            With sRect
                .Top = sHeight * 2 ' cooldown bar
                .Left = 0
                .Right = barWidth
                .Bottom = .Top + sHeight
            End With
            RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
        End If
    End If
    
    ' draw own health bar
    If GetPlayerVital(MyIndex, Vitals.HP) > 0 And GetPlayerVital(MyIndex, Vitals.HP) < GetPlayerMaxVital(MyIndex, Vitals.HP) Then
        ' lock to Player
        tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).xOffset + 16 - (sWidth / 2)
        tmpY = GetPlayerY(MyIndex) * PIC_X + Player(MyIndex).yOffset + 35
       
        ' calculate the width to fill
        barWidth = ((GetPlayerVital(MyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / sWidth)) * sWidth
       
        ' draw bar background
        With sRect
            .Top = sHeight * 1 ' HP bar background
            .Left = 0
            .Right = .Left + sWidth
            .Bottom = .Top + sHeight
        End With
        RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
       
        ' draw the bar proper
        With sRect
            .Top = 0 ' HP bar
            .Left = 0
            .Right = .Left + barWidth
            .Bottom = .Top + sHeight
        End With
        RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
    End If
    
    ' draw party health bars
    If Party.Leader > 0 Then
        For i = 1 To MAX_PARTY_MEMBERS
            partyIndex = Party.Member(i)
            If (partyIndex > 0) And (partyIndex <> MyIndex) And (GetPlayerMap(partyIndex) = GetPlayerMap(MyIndex)) Then
                ' player exists
                If GetPlayerVital(partyIndex, Vitals.HP) > 0 And GetPlayerVital(partyIndex, Vitals.HP) < GetPlayerMaxVital(partyIndex, Vitals.HP) Then
                    ' lock to Player
                    tmpX = GetPlayerX(partyIndex) * PIC_X + Player(partyIndex).xOffset + 16 - (sWidth / 2)
                    tmpY = GetPlayerY(partyIndex) * PIC_X + Player(partyIndex).yOffset + 35
                    
                    ' calculate the width to fill
                    barWidth = ((GetPlayerVital(partyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(partyIndex, Vitals.HP) / sWidth)) * sWidth
                    
                    ' draw bar background
                    With sRect
                        .Top = sHeight * 1 ' HP bar background
                        .Left = 0
                        .Right = .Left + sWidth
                        .Bottom = .Top + sHeight
                    End With
                    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                    
                    ' draw the bar proper
                    With sRect
                        .Top = 0 ' HP bar
                        .Left = 0
                        .Right = .Left + barWidth
                        .Bottom = .Top + sHeight
                    End With
                    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
            End If
        Next
    End If
                    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawBars", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawHotbar()
Dim sRect As RECT, dRect As RECT, i As Long, Num As String, n As Long, destRect As D3DRECT

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    For i = 1 To MAX_HOTBAR
    
        With dRect
            .Top = HotbarTop
            .Left = HotbarLeft + ((HotbarOffsetX + 32) * (((i - 1) Mod MAX_HOTBAR)))
            .Bottom = .Top + 32
            .Right = .Left + 32
        End With
        
        With destRect
            .y1 = HotbarTop
            .x1 = HotbarLeft + ((HotbarOffsetX + 32) * (((i - 1) Mod MAX_HOTBAR)))
            .y2 = .y1 + 32
            .x2 = .x1 + 32
        End With
        
        With sRect
            .Top = 0
            .Left = 32
            .Bottom = 32
            .Right = 64
        End With
        
        Select Case Hotbar(i).sType
            Case 1 ' inventory
                If Len(Item(Hotbar(i).Slot).Name) > 0 Then
                    If Item(Hotbar(i).Slot).Pic > 0 Then
                        If Item(Hotbar(i).Slot).Pic <= numitems Then
                            Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 255), 1#, 0
                            Direct3D_Device.BeginScene
                            RenderTextureByRects Tex_Item(Item(Hotbar(i).Slot).Pic), sRect, dRect
                            Direct3D_Device.EndScene
                            Direct3D_Device.Present destRect, destRect, frmMain.picHotbar.hWnd, ByVal (0)
                        End If
                    End If
                End If
            Case 2 ' spell
                With sRect
                    .Top = 0
                    .Left = 0
                    .Bottom = 32
                    .Right = 32
                End With
                If Len(Spell(Hotbar(i).Slot).Name) > 0 Then
                    If Spell(Hotbar(i).Slot).Icon > 0 Then
                        If Spell(Hotbar(i).Slot).Icon <= NumSpellIcons Then
                            ' check for cooldown
                            For n = 1 To MAX_PLAYER_SPELLS
                                If PlayerSpells(n) = Hotbar(i).Slot Then
                                    ' has spell
                                    If Not SpellCD(i) = 0 Then
                                        sRect.Left = 32
                                        sRect.Right = 64
                                    End If
                                End If
                            Next
                            Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 255), 1#, 0
                            Direct3D_Device.BeginScene
                            RenderTextureByRects Tex_SpellIcon(Spell(Hotbar(i).Slot).Icon), sRect, dRect
                            Direct3D_Device.EndScene
                            Direct3D_Device.Present destRect, destRect, frmMain.picHotbar.hWnd, ByVal (0)
                        End If
                    End If
                End If
        End Select
        
        ' render the letters
        Num = "F" & str(i)
        RenderText Font_Default, Num, dRect.Left + 2, dRect.Top + 16, White, 0
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawHotbar", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawPlayer(ByVal Index As Long)
Dim anim As Byte, i As Long, X As Long, Y As Long
Dim Sprite As Long, spritetop As Long
Dim rec As RECT
Dim attackspeed As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    Sprite = GetPlayerSprite(Index)

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    ' speed from weapon
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        attackspeed = Item(GetPlayerEquipment(Index, Weapon)).speed
    Else
        attackspeed = 1000
    End If

    ' Reset frame
    If Player(Index).Step = 3 Then
        anim = 0
    ElseIf Player(Index).Step = 1 Then
        anim = 2
    End If
    
    ' Check for attacking animation
    If Player(Index).AttackTimer + (attackspeed / 2) > timeGetTime Then
        If Player(Index).Attacking = 1 Then
            anim = 3
        End If
    Else
        ' If not attacking, walk normally
        Select Case GetPlayerDir(Index)
            Case DirectionUp
                If (Player(Index).yOffset > 8) Then anim = Player(Index).Step
            Case DirectionDown
                If (Player(Index).yOffset < -8) Then anim = Player(Index).Step
            Case DirectionLeft
                If (Player(Index).xOffset > 8) Then anim = Player(Index).Step
            Case DirectionRight
                If (Player(Index).xOffset < -8) Then anim = Player(Index).Step
        End Select
    End If

    ' Check to see if we want to stop making him attack
    With Player(Index)
        If .AttackTimer + attackspeed < timeGetTime Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the left
    Select Case GetPlayerDir(Index)
        Case DirectionUp
            spritetop = 3
        Case DirectionRight
            spritetop = 2
        Case DirectionDown
            spritetop = 0
        Case DirectionLeft
            spritetop = 1
    End Select

    With rec
        .Top = spritetop * (Tex_Character(Sprite).Height / 4)
        .Bottom = .Top + (Tex_Character(Sprite).Height / 4)
        .Left = anim * (Tex_Character(Sprite).Width / 4)
        .Right = .Left + (Tex_Character(Sprite).Width / 4)
    End With

    ' Calculate the X
    X = GetPlayerX(Index) * PIC_X + Player(Index).xOffset - ((Tex_Character(Sprite).Width / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (Tex_Character(Sprite).Height) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - ((Tex_Character(Sprite).Height / 4) - 32)
    Else
        ' Proceed as normal
        Y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset
    End If

    ' render the actual sprite
    Call DrawSprite(Sprite, X, Y, rec)
    If Player(MyIndex).Sprite > 2 Then
     Exit Sub
    End If
   If GetPlayerEquipment(Index, Shield) <> Player(Index).WieldDagger Then  'Get player daggers [ 1]
        If GetPlayerEquipment(Index, Armor) > 0 Then
             Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, Armor)).Paperdoll, anim, spritetop)
             End If
           If GetPlayerEquipment(Index, Helmet) > 0 Then
             Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, Helmet)).Paperdoll, anim, spritetop)
            End If
        If GetPlayerEquipment(Index, Legs) > 0 Then
              Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, Legs)).Paperdoll, anim, spritetop)
           End If
         If GetPlayerEquipment(Index, Boots) > 0 Then
              Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, Boots)).Paperdoll, anim, spritetop)
           End If
         If GetPlayerEquipment(Index, Glove) > 0 Then
              Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, Glove)).Paperdoll, anim, spritetop)
           End If
         If GetPlayerEquipment(Index, Ring) > 0 Then
              Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, Ring)).Paperdoll, anim, spritetop)
           End If
         If GetPlayerEquipment(Index, Enchant) > 0 Then
              Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, Enchant)).Paperdoll, anim, spritetop)
           End If
           
      If Item(GetPlayerEquipment(Index, Shield)).Daggerpdoll > 0 Then
         If GetPlayerEquipment(Index, Shield) > 0 Then
              Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, Shield)).Daggerpdoll, anim, spritetop)
           End If
         Else
            Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, Shield)).Paperdoll, anim, spritetop)
         End If
             If GetPlayerEquipment(Index, Weapon) > 0 Then
        Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, Weapon)).Paperdoll, anim, spritetop)
             End If
      Exit Sub
    End If


    ' check for paperdolling
    For i = 1 To UBound(PaperdollOrder)
        If GetPlayerEquipment(Index, PaperdollOrder(i)) > 0 Then
            If Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll > 0 Then
                Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll, anim, spritetop)
            End If
      End If
    Next


     
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPlayer", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawNpc(ByVal MapNpcNum As Long)
Dim anim As Byte, i As Long, X As Long, Y As Long, Sprite As Long, spritetop As Long
Dim rec As RECT
Dim attackspeed As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If MapNpc(MapNpcNum).Num = 0 Then Exit Sub ' no npc set
    
    Sprite = NPC(MapNpc(MapNpcNum).Num).Sprite

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    attackspeed = 1000

    ' Reset frame
    anim = 0
    ' Check for attacking animation
    If MapNpc(MapNpcNum).AttackTimer + (attackspeed / 2) > timeGetTime Then
        If MapNpc(MapNpcNum).Attacking = 1 Then
            anim = 3
        End If
    Else
        ' If not attacking, walk normally
        Select Case MapNpc(MapNpcNum).Dir
            Case DirectionUp
                If (MapNpc(MapNpcNum).yOffset > 8) Then anim = MapNpc(MapNpcNum).Step
            Case DirectionDown
                If (MapNpc(MapNpcNum).yOffset < -8) Then anim = MapNpc(MapNpcNum).Step
            Case DirectionLeft
                If (MapNpc(MapNpcNum).xOffset > 8) Then anim = MapNpc(MapNpcNum).Step
            Case DirectionRight
                If (MapNpc(MapNpcNum).xOffset < -8) Then anim = MapNpc(MapNpcNum).Step
        End Select
    End If

    ' Check to see if we want to stop making him attack
    With MapNpc(MapNpcNum)
        If .AttackTimer + attackspeed < timeGetTime Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the left
    Select Case MapNpc(MapNpcNum).Dir
        Case DirectionUp
            spritetop = 3
        Case DirectionRight
            spritetop = 2
        Case DirectionDown
            spritetop = 0
        Case DirectionLeft
            spritetop = 1
    End Select

    With rec
        .Top = (Tex_Character(Sprite).Height / 4) * spritetop
        .Bottom = .Top + Tex_Character(Sprite).Height / 4
        .Left = anim * (Tex_Character(Sprite).Width / 4)
        .Right = .Left + (Tex_Character(Sprite).Width / 4)
    End With

    ' Calculate the X
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).xOffset - ((Tex_Character(Sprite).Width / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (Tex_Character(Sprite).Height / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).yOffset - ((Tex_Character(Sprite).Height / 4) - 32)
    Else
        ' Proceed as normal
        Y = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).yOffset
    End If

    Call DrawSprite(Sprite, X, Y, rec)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawNpc", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawPaperdoll(ByVal x2 As Long, ByVal y2 As Long, ByVal Sprite As Long, ByVal anim As Long, ByVal spritetop As Long)
Dim rec As RECT
Dim X As Long, Y As Long
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Sprite < 1 Or Sprite > NumPaperdolls Then Exit Sub
    
    With rec
        .Top = spritetop * (Tex_Paperdoll(Sprite).Height / 4)
        .Bottom = .Top + (Tex_Paperdoll(Sprite).Height / 4)
        .Left = anim * (Tex_Paperdoll(Sprite).Width / 4)
        .Right = .Left + (Tex_Paperdoll(Sprite).Width / 4)
    End With
    
    ' clipping
    X = ConvertMapX(x2)
    Y = ConvertMapY(y2)
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)

    ' Clip to screen
    If Y < 0 Then
        With rec
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If
    
    RenderTexture Tex_Paperdoll(Sprite), X, Y, rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPaperdoll", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawSprite(ByVal Sprite As Long, ByVal x2 As Long, y2 As Long, rec As RECT)
Dim X As Long
Dim Y As Long
Dim Width As Long
Dim Height As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    X = ConvertMapX(x2)
    Y = ConvertMapY(y2)
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)
    ' render player shadow
    If frmMain.scrnopt.Value Then
    RenderTexture Tex_Shadow, X, Y + 18, 0, 0, 32, 32, 32, 32, D3DColorRGBA(255, 255, 255, 200)
    End If
    RenderTexture Tex_Character(Sprite), X, Y, rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawFog()
Dim fogNum As Long, color As Long, X As Long, Y As Long, renderState As Long

    fogNum = CurrentFog
    If fogNum <= 0 Or fogNum > NumFogs Then Exit Sub
    color = D3DColorRGBA(255, 255, 255, 255 - CurrentFogOpacity)

    renderState = 0
    ' render state
    Select Case renderState
        Case 1 ' Additive
            Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
            Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        Case 2 ' Subtractive
            Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SUBTRACT
            Direct3D_Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ZERO
            Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
    End Select
    
    For X = 0 To ((Map.MaxX * 32) / 256) + 1
        For Y = 0 To ((Map.MaxY * 32) / 256) + 1
            RenderTexture Tex_Fog(fogNum), ConvertMapX((X * 256) + fogOffsetX), ConvertMapY((Y * 256) + fogOffsetY), 0, 0, 256, 256, 256, 256, color
        Next
    Next
    
    ' reset render state
    If renderState > 0 Then
        Direct3D_Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
    End If
End Sub

Public Sub DrawTint()
Dim color As Long
    color = D3DColorRGBA(CurrentTintR, CurrentTintG, CurrentTintB, CurrentTintA)
    RenderTexture Tex_White, 0, 0, 0, 0, frmMain.picScreen.ScaleWidth, frmMain.picScreen.ScaleHeight, 32, 32, color
End Sub

Public Sub DrawWeather()
Dim color As Long, i As Long, SpriteLeft As Long
   If Map.Weather = 0 Then Exit Sub
    For i = 1 To MAX_WEATHER_PARTICLES
        If WeatherParticle(i).InUse Then
            If WeatherParticle(i).Type = WeatherStorm Then
                SpriteLeft = 0
            Else
                SpriteLeft = WeatherParticle(i).Type - 1
            End If
            RenderTexture Tex_Weather, ConvertMapX(WeatherParticle(i).X), ConvertMapY(WeatherParticle(i).Y), SpriteLeft * 32, 0, 32, 32, 32, 32, -1
        End If
    Next
End Sub

Sub DrawAnimatedInvItems()
Dim i As Long
Dim itemnum As Long, itempic As Long
Dim X As Long, Y As Long
Dim MaxFrames As Byte
Dim Amount As Long
Dim rec As RECT, rec_pos As RECT

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    ' check for map animation changes#
    For i = 1 To MAX_MAP_ITEMS

        If MapItem(i).Num > 0 Then
            itempic = Item(MapItem(i).Num).Pic

            If itempic < 1 Or itempic > numitems Then Exit Sub
            MaxFrames = (Tex_Item(itempic).Width / 2) / 32 ' Work out how many frames there are. /2 because of inventory icons as well as ingame

            If MapItem(i).Frame < MaxFrames - 1 Then
                MapItem(i).Frame = MapItem(i).Frame + 1
            Else
                MapItem(i).Frame = 1
            End If
        End If

    Next

    For i = 1 To MAX_INV
        itemnum = GetPlayerInvItemNum(MyIndex, i)

        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            itempic = Item(itemnum).Pic

            If itempic > 0 And itempic <= numitems Then
                If Tex_Item(itempic).Width > 64 Then
                    MaxFrames = (Tex_Item(itempic).Width / 2) / 32 ' Work out how many frames there are. /2 because of inventory icons as well as ingame

                    If InvItemFrame(i) < MaxFrames - 1 Then
                        InvItemFrame(i) = InvItemFrame(i) + 1
                    Else
                        InvItemFrame(i) = 1
                    End If

                    With rec
                        .Top = 0
                        .Bottom = 32
                        .Left = (Tex_Item(itempic).Width / 2) + (InvItemFrame(i) * 32) ' middle to get the start of inv gfx, then +32 for each frame
                        .Right = .Left + 32
                    End With

                    With rec_pos
                        .Top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                        .Bottom = .Top + PIC_Y
                        .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                        .Right = .Left + PIC_X
                    End With

                    ' We'll now re-Draw the item, and place the currency value over it again :P
                    RenderTextureByRects Tex_Item(itempic), rec, rec_pos

                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, i) >= 1 Then
                        Y = rec_pos.Top + 22
                        X = rec_pos.Left - 4
                        Amount = CStr(GetPlayerInvItemValue(MyIndex, i))
                        ' Draw currency but with k, m, b etc. using a convertion function
                        RenderText Font_Default, ConvertCurrency(Amount), X, Y, Yellow, 0
                        ' Check if it's gold, and update the label
                        If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ItemCurrency Then
                            frmMain.lblGold.Caption = Format$(Amount, "#,###,###,###") & "g"
                        End If
                    End If
                End If
            End If
        End If

    Next

    'frmMain.picInventory.Refresh
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawAnimatedInvItems", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DrawFace()
Dim rec As RECT, rec_pos As RECT, faceNum As Long, srcRect As D3DRECT
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If NumFaces = 0 Then Exit Sub

    faceNum = GetPlayerSprite(MyIndex)
    
    If faceNum <= 0 Or faceNum > NumFaces Then Exit Sub
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene

    With rec
        .Top = 0
        .Bottom = 100
        .Left = 0
        .Right = 100
    End With

    With rec_pos
        .Top = 0
        .Bottom = 100
        .Left = 0
        .Right = 100
    End With

    RenderTextureByRects Tex_Face(faceNum), rec, rec_pos
    With srcRect
        .x1 = 0
        .x2 = frmMain.picFace.Width
        .y1 = 0
        .y2 = frmMain.picFace.Height
    End With
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, srcRect, frmMain.picFace.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawFace", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DrawEquipment()
Dim i As Long, itemnum As Long, itempic As Long
Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT, destRect As D3DRECT
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If numitems = 0 Then Exit Sub
    
    'frmMain.picCharacter.Cls
    For i = 1 To Equipment.Equipment_Count - 1
        itemnum = GetPlayerEquipment(MyIndex, i)

        If itemnum > 0 Then
            itempic = Item(itemnum).Pic

            With rec
                .Top = 0
                .Bottom = 32
                .Left = 32
                .Right = 64
            End With

                With rec_pos
                                .Top = EqTop
                                .Bottom = .Top + PIC_Y
                                .Left = EqLeft + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))
                                .Right = .Left + PIC_X
                        End With
            Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
            Direct3D_Device.BeginScene
            RenderTextureByRects Tex_Item(itempic), rec, rec_pos
            Direct3D_Device.EndScene
            With srcRect
                .x1 = rec_pos.Left
                .x2 = rec_pos.Right
                .y1 = rec_pos.Top
                .y2 = rec_pos.Bottom
            End With
            Direct3D_Device.Present srcRect, srcRect, frmMain.picCharacter.hWnd, ByVal (0)
        End If
    Next
    

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawEquipment", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DrawInventory()
Dim i As Long, X As Long, Y As Long, itemnum As Long, itempic As Long
Dim Amount As Long
Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT, destRect As D3DRECT
Dim colour As Long
Dim tmpItem As Long, amountModifier As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene

    For i = 1 To MAX_INV
        itemnum = GetPlayerInvItemNum(MyIndex, i)

        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            itempic = Item(itemnum).Pic
            
            amountModifier = 0
            ' exit out if we're offering item in a trade.
            If InTrade > 0 Then
                For X = 1 To MAX_INV
                    tmpItem = GetPlayerInvItemNum(MyIndex, TradeYourOffer(X).Num)
                    If TradeYourOffer(X).Num = i Then
                        ' check if currency
                        If Not Item(tmpItem).Type = ItemCurrency Then
                            ' normal item, exit out
                            GoTo NextLoop
                        Else
                            ' if amount = all currency, remove from inventory
                            If TradeYourOffer(X).Value = GetPlayerInvItemValue(MyIndex, i) Then
                                GoTo NextLoop
                            Else
                                ' not all, change modifier to show change in currency count
                                amountModifier = TradeYourOffer(X).Value
                            End If
                        End If
                    End If
                Next
            End If

            If itempic > 0 And itempic <= numitems Then
                If Tex_Item(itempic).Width <= 64 Then ' more than 1 frame is handled by anim sub

                    With rec
                        .Top = 0
                        .Bottom = 32
                        .Left = 32
                        .Right = 64
                    End With

                    With rec_pos
                        .Top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                        .Bottom = .Top + PIC_Y
                        .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                        .Right = .Left + PIC_X
                    End With

                    RenderTextureByRects Tex_Item(itempic), rec, rec_pos

                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                        Y = rec_pos.Top + 22
                        X = rec_pos.Left - 4
                        
                        Amount = GetPlayerInvItemValue(MyIndex, i) - amountModifier
                        
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If Amount < 1000000 Then
                            colour = White
                        ElseIf Amount > 1000000 And Amount < 10000000 Then
                            colour = Yellow
                        ElseIf Amount > 10000000 Then
                            colour = BrightGreen
                        End If
                        RenderText Font_Default, Format$(ConvertCurrency(str(Amount)), "#,###,###,###"), X, Y, colour, 0
                        ' Check if it's gold, and update the label
                        If GetPlayerInvItemNum(MyIndex, i) = 1 Then '1 = gold :P
                            frmMain.lblGold.Caption = Format$(Amount, "#,###,###,###") & "g"
                        End If
                    End If
                End If
            End If
        End If
NextLoop:
    Next
    
    'update animated items
    DrawAnimatedInvItems
    
    With srcRect
        .x1 = 0
        .x2 = frmMain.picInventory.Width
        .y1 = 28
        .y2 = frmMain.picInventory.Height + .y1
    End With
    
    With destRect
        .x1 = 0
        .x2 = frmMain.picInventory.Width
        .y1 = 32
        .y2 = frmMain.picInventory.Height + .y1
    End With
    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, destRect, frmMain.picInventory.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawInventory", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DrawTrade()
Dim i As Long, X As Long, Y As Long, itemnum As Long, itempic As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim Amount As Long
Dim rec As RECT, rec_pos As RECT
Dim colour As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    For i = 1 To MAX_INV
        ' Draw your own offer
        itemnum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).Num)

        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            itempic = Item(itemnum).Pic

            If itempic > 0 And itempic <= numitems Then
                With rec
                    .Top = 0
                    .Bottom = 32
                    .Left = 32
                    .Right = 64
                End With

                With rec_pos
                    .Top = InvTop - 24 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    .Bottom = .Top + PIC_Y
                    .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    .Right = .Left + PIC_X
                End With

                RenderTextureByRects Tex_Item(itempic), rec, rec_pos

                ' If item is a stack - draw the amount you have
                If TradeYourOffer(i).Value > 1 Then
                    Y = rec_pos.Top + 22
                    X = rec_pos.Left - 4
                    
                    Amount = TradeYourOffer(i).Value
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If Amount < 1000000 Then
                        colour = White
                    ElseIf Amount > 1000000 And Amount < 10000000 Then
                        colour = Yellow
                    ElseIf Amount > 10000000 Then
                        colour = BrightGreen
                    End If
                    RenderText Font_Default, ConvertCurrency(str(Amount)), X, Y, colour, 0
                End If
            End If
        End If
    Next
    
    With srcRect
        .x1 = 0
        .x2 = .x1 + 193
        .y1 = 0
        .y2 = .y1 + 246
    End With
                    
    With destRect
        .x1 = 0
        .x2 = .x1 + 193
        .y1 = 0
        .y2 = 246 + .y1
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, destRect, frmMain.picYourTrade.hWnd, ByVal (0)
    
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    For i = 1 To MAX_INV
        ' Draw their offer
        itemnum = TradeTheirOffer(i).Num

        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            itempic = Item(itemnum).Pic

            If itempic > 0 And itempic <= numitems Then
                With rec
                    .Top = 0
                    .Bottom = 32
                    .Left = 32
                    .Right = 64
                End With

                With rec_pos
                    .Top = InvTop - 24 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    .Bottom = .Top + PIC_Y
                    .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    .Right = .Left + PIC_X
                End With
                
                RenderTextureByRects Tex_Item(itempic), rec, rec_pos

                ' If item is a stack - draw the amount you have
                If TradeTheirOffer(i).Value > 1 Then
                    Y = rec_pos.Top + 22
                    X = rec_pos.Left - 4
                    
                    Amount = TradeTheirOffer(i).Value
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If Amount < 1000000 Then
                        colour = White
                    ElseIf Amount > 1000000 And Amount < 10000000 Then
                        colour = Yellow
                    ElseIf Amount > 10000000 Then
                        colour = BrightGreen
                    End If
                    RenderText Font_Default, ConvertCurrency(str(Amount)), X, Y, colour, 0
                End If
            End If
        End If
    Next
    
    With srcRect
        .x1 = 0
        .x2 = .x1 + 193
        .y1 = 0
        .y2 = .y1 + 246
    End With
                    
    With destRect
        .x1 = 0
        .x2 = .x1 + 193
        .y1 = 0
        .y2 = 246 + .y1
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, destRect, frmMain.picTheirTrade.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawTrade", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DrawPlayerSpells()
Dim i As Long, X As Long, Y As Long, spellnum As Long, spellicon As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim Amount As String
Dim rec As RECT, rec_pos As RECT
Dim colour As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    'frmMain.picSpells.Cls
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene

    For i = 1 To MAX_PLAYER_SPELLS
        spellnum = PlayerSpells(i)

        If spellnum > 0 And spellnum <= MAX_SPELLS Then
            spellicon = Spell(spellnum).Icon

            If spellicon > 0 And spellicon <= NumSpellIcons Then
            
                With rec
                    .Top = 0
                    .Bottom = 32
                    .Left = 0
                    .Right = 32
                End With
                
                If Not SpellCD(i) = 0 Then
                    rec.Left = 32
                    rec.Right = 64
                End If

                With rec_pos
                    .Top = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                    .Bottom = .Top + PIC_Y
                    .Left = SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                    .Right = .Left + PIC_X
                End With

                RenderTextureByRects Tex_SpellIcon(spellicon), rec, rec_pos
            End If
        End If
    Next
    
    With srcRect
        .x1 = 0
        .x2 = frmMain.picSpells.Width
        .y1 = 28
        .y2 = frmMain.picSpells.Height + .y1
    End With
    
    With destRect
        .x1 = 0
        .x2 = frmMain.picSpells.Width
        .y1 = 32
        .y2 = frmMain.picSpells.Height + .y1
    End With
    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, destRect, frmMain.picSpells.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPlayerSpells", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DrawShop()
Dim i As Long, X As Long, Y As Long, itemnum As Long, itempic As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim Amount As String
Dim rec As RECT, rec_pos As RECT
Dim colour As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    'frmMain.picShopItems.Cls
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    For i = 1 To MAX_TRADES
        itemnum = Shop(InShop).TradeItem(i).Item 'GetPlayerInvItemNum(MyIndex, i)
        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            itempic = Item(itemnum).Pic
            If itempic > 0 And itempic <= numitems Then
            
                With rec
                    .Top = 0
                    .Bottom = 32
                    .Left = 32
                    .Right = 64
                End With
                
                With rec_pos
                    .Top = ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
                    .Bottom = .Top + PIC_Y
                    .Left = ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
                    .Right = .Left + PIC_X
                End With
                
                RenderTextureByRects Tex_Item(itempic), rec, rec_pos
                
                ' If item is a stack - draw the amount you have
                If Shop(InShop).TradeItem(i).ItemValue > 1 Then
                    Y = rec_pos.Top + 22
                    X = rec_pos.Left - 4
                    Amount = CStr(Shop(InShop).TradeItem(i).ItemValue)
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        colour = Green
                    End If
                    RenderText Font_Default, ConvertCurrency(Amount), X, Y, colour, 0
                End If
            End If
        End If
    Next
    
    With srcRect
        .x1 = ShopLeft
        .x2 = .x1 + 192
        .y1 = ShopTop
        .y2 = .y1 + 211
    End With
                
    With destRect
        .x1 = ShopLeft
        .x2 = .x1 + 192
        .y1 = ShopTop
        .y2 = 211 + .y1
    End With
                
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, destRect, frmMain.picShopItems.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawShop", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawInventoryItem(ByVal X As Long, ByVal Y As Long)
Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT, destRect As D3DRECT
Dim itemnum As Long, itempic As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    itemnum = GetPlayerInvItemNum(MyIndex, DragInvSlotNum)

    If itemnum > 0 And itemnum <= MAX_ITEMS Then
        itempic = Item(itemnum).Pic
        
        If itempic = 0 Then Exit Sub
        
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 255), 1#, 0
        Direct3D_Device.BeginScene
        
        With rec
            .Top = 0
            .Bottom = .Top + PIC_Y
            .Left = Tex_Item(itempic).Width / 2
            .Right = .Left + PIC_X
        End With

        With rec_pos
            .Top = 2
            .Bottom = .Top + PIC_Y
            .Left = 2
            .Right = .Left + PIC_X
        End With

        RenderTextureByRects Tex_Item(itempic), rec, rec_pos

        With frmMain.picTempInv
            .Top = Y
            .Left = X
            .Visible = True
            .ZOrder (0)
        End With
        With srcRect
            .x1 = 0
            .x2 = 32
            .y1 = 0
            .y2 = 32
        End With
        With destRect
            .x1 = 2
            .y1 = 2
            .y2 = .y1 + 32
            .x2 = .x1 + 32
        End With
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRect, frmMain.picTempInv.hWnd, ByVal (0)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawInventoryItem", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawDraggedSpell(ByVal X As Long, ByVal Y As Long)
Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT, destRect As D3DRECT
Dim spellnum As Long, spellpic As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    spellnum = PlayerSpells(DragSpell)

    If spellnum > 0 And spellnum <= MAX_SPELLS Then
        spellpic = Spell(spellnum).Icon
        
        If spellpic = 0 Then Exit Sub
        
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene

        With rec
            .Top = 0
            .Bottom = .Top + PIC_Y
            .Left = 0
            .Right = .Left + PIC_X
        End With

        With rec_pos
            .Top = 2
            .Bottom = .Top + PIC_Y
            .Left = 2
            .Right = .Left + PIC_X
        End With

        RenderTextureByRects Tex_SpellIcon(spellpic), rec, rec_pos

        With frmMain.picTempSpell
            .Top = Y
            .Left = X
            .Visible = True
            .ZOrder (0)
        End With
        
        With srcRect
            .x1 = 0
            .x2 = 32
            .y1 = 0
            .y2 = 32
        End With
        With destRect
            .x1 = 2
            .y1 = 2
            .y2 = .y1 + 32
            .x2 = .x1 + 32
        End With
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRect, frmMain.picTempSpell.hWnd, ByVal (0)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawInventoryItem", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawItemDesc(ByVal itemnum As Long)
Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT, destRect As D3DRECT
Dim itempic As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    'frmMain.picItemDescPic.Cls
    
    If itemnum > 0 And itemnum <= MAX_ITEMS Then
        itempic = Item(itemnum).Pic

        If itempic = 0 Then Exit Sub
        
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene

        With rec
            .Top = 0
            .Bottom = .Top + PIC_Y
            .Left = Tex_Item(itempic).Width / 2
            .Right = .Left + PIC_X
        End With

        With rec_pos
            .Top = 0
            .Bottom = 64
            .Left = 0
            .Right = 64
        End With
        RenderTextureByRects Tex_Item(itempic), rec, rec_pos

        With destRect
            .x1 = 0
            .y1 = 0
            .y2 = 64
            .x2 = 64
        End With
        
        Direct3D_Device.EndScene
        Direct3D_Device.Present destRect, destRect, frmMain.picItemDescPic.hWnd, ByVal (0)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawItemDesc", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawSpellDesc(ByVal spellnum As Long)
Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT, destRect As D3DRECT
Dim spellpic As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    'frmMain.picSpellDescPic.Cls

    If spellnum > 0 And spellnum <= MAX_SPELLS Then
        spellpic = Spell(spellnum).Icon

        If spellpic <= 0 Or spellpic > NumSpellIcons Then Exit Sub
        
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene

        With rec
            .Top = 0
            .Bottom = .Top + PIC_Y
            .Left = 0
            .Right = .Left + PIC_X
        End With

        With rec_pos
            .Top = 0
            .Bottom = 64
            .Left = 0
            .Right = 64
        End With
        RenderTextureByRects Tex_SpellIcon(spellpic), rec, rec_pos

        With destRect
            .x1 = 0
            .y1 = 0
            .y2 = 64
            .x2 = 64
        End With
        
        Direct3D_Device.EndScene
        Direct3D_Device.Present destRect, destRect, frmMain.picSpellDescPic.hWnd, ByVal (0)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawSpellDesc", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ******************
' ** Game Editors **
' ******************
Public Sub EditorMap_DrawTileset()
Dim Height As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim Width As Long
Dim Tileset As Long
Dim sRect As RECT
Dim dRect As RECT, scrlX As Long, scrlY As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    ' find tileset number
    Tileset = frmEditor_Map.scrlTileSet.Value
    
    ' exit out if doesn't exist
    If Tileset < 0 Or Tileset > NumTileSets Then Exit Sub
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    scrlX = frmEditor_Map.scrlPictureX.Value * PIC_X
    scrlY = frmEditor_Map.scrlPictureY.Value * PIC_Y
    
    Height = Tex_Tileset(Tileset).Height - scrlY
    Width = Tex_Tileset(Tileset).Width - scrlX
    
    sRect.Left = frmEditor_Map.scrlPictureX.Value * PIC_X
    sRect.Top = frmEditor_Map.scrlPictureY.Value * PIC_Y
    sRect.Right = sRect.Left + Width
    sRect.Bottom = sRect.Top + Height
    
    dRect.Top = 0
    dRect.Bottom = Height
    dRect.Left = 0
    dRect.Right = Width
    
    RenderTextureByRects Tex_Tileset(Tileset), sRect, dRect
    
    ' change selected shape for autotiles
    If frmEditor_Map.scrlAutotile.Value > 0 Then
        Select Case frmEditor_Map.scrlAutotile.Value
            Case 1 ' autotile
                EditorTileWidth = 2
                EditorTileHeight = 3
            Case 2 ' fake autotile
                EditorTileWidth = 1
                EditorTileHeight = 1
            Case 3 ' animated
                EditorTileWidth = 6
                EditorTileHeight = 3
            Case 4 ' cliff
                EditorTileWidth = 2
                EditorTileHeight = 2
            Case 5 ' waterfall
                EditorTileWidth = 2
                EditorTileHeight = 3
        End Select
    End If
    
    With destRect
        .x1 = (EditorTileX * 32) - sRect.Left
        .x2 = (EditorTileWidth * 32) + .x1
        .y1 = (EditorTileY * 32) - sRect.Top
        .y2 = (EditorTileHeight * 32) + .y1
    End With
    
    DrawSelectionBox destRect
        
    With srcRect
        .x1 = 0
        .x2 = Width
        .y1 = 0
        .y2 = Height
    End With
                    
    With destRect
        .x1 = 0
        .x2 = frmEditor_Map.picBack.ScaleWidth
        .y1 = 0
        .y2 = frmEditor_Map.picBack.ScaleHeight
    End With
    
    'Now render the selection tiles and we are done!
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Map.picBack.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_DrawTileset", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DrawSelectionBox(dRect As D3DRECT)
Dim Width As Long, Height As Long, X As Long, Y As Long
    Width = dRect.x2 - dRect.x1
    Height = dRect.y2 - dRect.y1
    X = dRect.x1
    Y = dRect.y1
    If Width > 6 And Height > 6 Then
        'Draw Box 32 by 32 at graphicselx and graphicsely
        RenderTexture Tex_Selection, X, Y, 1, 1, 2, 2, 2, 2, -1 'top left corner
        RenderTexture Tex_Selection, X + 2, Y, 3, 1, Width - 4, 2, 32 - 6, 2, -1 'top line
        RenderTexture Tex_Selection, X + 2 + (Width - 4), Y, 29, 1, 2, 2, 2, 2, -1 'top right corner
        RenderTexture Tex_Selection, X, Y + 2, 1, 3, 2, Height - 4, 2, 32 - 6, -1 'Left Line
        RenderTexture Tex_Selection, X + 2 + (Width - 4), Y + 2, 32 - 3, 3, 2, Height - 4, 2, 32 - 6, -1 'right line
        RenderTexture Tex_Selection, X, Y + 2 + (Height - 4), 1, 32 - 3, 2, 2, 2, 2, -1 'bottom left corner
        RenderTexture Tex_Selection, X + 2 + (Width - 4), Y + 2 + (Height - 4), 32 - 3, 32 - 3, 2, 2, 2, 2, -1 'bottom right corner
        RenderTexture Tex_Selection, X + 2, Y + 2 + (Height - 4), 3, 32 - 3, Width - 4, 2, 32 - 6, 2, -1 'bottom line
    End If
End Sub

Public Sub DrawTileOutline()
Dim rec As RECT
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Map.optBlock.Value Then Exit Sub

    With rec
        .Top = 0
        .Bottom = .Top + PIC_Y
        .Left = 0
        .Right = .Left + PIC_X
    End With

    RenderTexture Tex_Misc, ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawTileOutline", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub NewCharacterDrawSprite()
Dim Sprite As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim sRect As RECT
Dim dRect As RECT
Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If frmMenu.cmbClass.ListIndex = -1 Then Exit Sub
    
    If frmMenu.optMale.Value = True Then
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).MaleSprite(newCharSprite)
    Else
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).FemaleSprite(newCharSprite)
    End If
    
    If Sprite < 1 Or Sprite > NumCharacters Then
        frmMenu.picSprite.Cls
        Exit Sub
    End If
    
    Width = Tex_Character(Sprite).Width / 4
    Height = Tex_Character(Sprite).Height / 4
    
    frmMenu.picSprite.Width = Width
    frmMenu.picSprite.Height = Height
    
    sRect.Top = 0
    sRect.Bottom = sRect.Top + Height
    sRect.Left = 0
    sRect.Right = sRect.Left + Width
    
    dRect.Top = 0
    dRect.Bottom = Height
    dRect.Left = 0
    dRect.Right = Width
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    RenderTextureByRects Tex_Character(Sprite), sRect, dRect
    
    With srcRect
        .x1 = 0
        .x2 = Width
        .y1 = 0
        .y2 = Height
    End With
                    
    With destRect
        .x1 = 0
        .x2 = Width
        .y1 = 0
        .y2 = Height
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, destRect, frmMenu.picSprite.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NewCharacterDrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorMap_DrawMapItem()
Dim itemnum As Long
Dim sRect As RECT, destRect As D3DRECT
Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    itemnum = Item(frmEditor_Map.scrlMapItem.Value).Pic

    If itemnum < 1 Or itemnum > numitems Then
        frmEditor_Map.picMapItem.Cls
        Exit Sub
    End If

    sRect.Top = 0
    sRect.Bottom = PIC_Y
    sRect.Left = 0
    sRect.Right = PIC_X
    dRect.Top = 0
    dRect.Bottom = PIC_Y
    dRect.Left = 0
    dRect.Right = PIC_X
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Item(itemnum), sRect, dRect
    With destRect
        .x1 = 0
        .x2 = PIC_X
        .y1 = 0
        .y2 = PIC_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Map.picMapItem.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_DrawMapItem", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorMap_DrawKey()
Dim itemnum As Long
Dim sRect As RECT, destRect As D3DRECT
Dim dRect As RECT

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    itemnum = Item(frmEditor_Map.scrlMapKey.Value).Pic

    If itemnum < 1 Or itemnum > numitems Then
        frmEditor_Map.picMapKey.Cls
        Exit Sub
    End If
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene

    sRect.Top = 0
    sRect.Bottom = PIC_Y
    sRect.Left = 0
    sRect.Right = PIC_X
    dRect.Top = 0
    dRect.Bottom = PIC_Y
    dRect.Left = 0
    dRect.Right = PIC_X
    
    RenderTextureByRects Tex_Item(itemnum), sRect, dRect
    
    With destRect
        .x1 = 0
        .x2 = PIC_X
        .y1 = 0
        .y2 = PIC_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Map.picMapKey.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_DrawKey", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_DrawItem()
Dim itemnum As Long
Dim sRect As RECT, destRect As D3DRECT
Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    itemnum = frmEditor_Item.scrlPic.Value

    If itemnum < 1 Or itemnum > numitems Then
        frmEditor_Item.picItem.Cls
        Exit Sub
    End If


    ' rect for source
    sRect.Top = 0
    sRect.Bottom = PIC_Y
    sRect.Left = 0
    sRect.Right = PIC_X
    
    ' same for destination as source
    dRect = sRect
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Item(itemnum), sRect, dRect
    With destRect
        .x1 = 0
        .x2 = PIC_X
        .y1 = 0
        .y2 = PIC_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Item.picItem.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_DrawItem", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_DrawPaperdoll()
Dim Sprite As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim sRect As RECT
Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    'frmEditor_Item.picPaperdoll.Cls
    
    Sprite = frmEditor_Item.scrlPaperdoll.Value

    If Sprite < 1 Or Sprite > NumPaperdolls Then
        frmEditor_Item.picPaperdoll.Cls
        Exit Sub
    End If

    ' rect for source
    sRect.Top = 0
    sRect.Bottom = Tex_Paperdoll(Sprite).Height / 4
    sRect.Left = 0
    sRect.Right = Tex_Paperdoll(Sprite).Width / 4
    ' same for destination as source
    dRect = sRect
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Paperdoll(Sprite), sRect, dRect
                    
    With destRect
        .x1 = 0
        .x2 = Tex_Paperdoll(Sprite).Width / 4
        .y1 = 0
        .y2 = Tex_Paperdoll(Sprite).Height / 4
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Item.picPaperdoll.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_DrawPaperdoll", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorSpell_DrawIcon()
Dim iconnum As Long, destRect As D3DRECT
Dim sRect As RECT
Dim dRect As RECT

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    iconnum = frmEditor_Spell.scrlIcon.Value
    
    If iconnum < 1 Or iconnum > NumSpellIcons Then
        frmEditor_Spell.picSprite.Cls
        Exit Sub
    End If
    
    sRect.Top = 0
    sRect.Bottom = PIC_Y
    sRect.Left = 0
    sRect.Right = PIC_X
    dRect.Top = 0
    dRect.Bottom = PIC_Y
    dRect.Left = 0
    dRect.Right = PIC_X
    
    With destRect
        .x1 = 0
        .x2 = PIC_X
        .y1 = 0
        .y2 = PIC_Y
    End With
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_SpellIcon(iconnum), sRect, dRect
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_Spell.picSprite.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorSpell_DrawIcon", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorAnim_DrawAnim()
Dim Animationnum As Long
Dim sRect As RECT
Dim dRect As RECT
Dim i As Long
Dim Width As Long, Height As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim looptime As Long
Dim FrameCount As Long
Dim ShouldRender As Boolean
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    For i = 0 To 1
        Animationnum = frmEditor_Animation.scrlSprite(i).Value
        
        If Animationnum < 1 Or Animationnum > NumAnimations Then
            frmEditor_Animation.picSprite(i).Cls
        Else
            looptime = frmEditor_Animation.scrlLoopTime(i)
            FrameCount = frmEditor_Animation.scrlFrameCount(i)
            
            ShouldRender = False
            
            ' check if we need to render new frame
            If AnimEditorTimer(i) + looptime <= timeGetTime Then
                ' check if out of range
                If AnimEditorFrame(i) >= FrameCount Then
                    AnimEditorFrame(i) = 1
                Else
                    AnimEditorFrame(i) = AnimEditorFrame(i) + 1
                End If
                AnimEditorTimer(i) = timeGetTime
                ShouldRender = True
            End If
        
            If ShouldRender Then
                'frmEditor_Animation.picSprite(i).Cls
                
                If frmEditor_Animation.scrlFrameCount(i).Value > 0 Then
                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    ' total width divided by frame count
                    Width = Tex_Animation(Animationnum).Width / frmEditor_Animation.scrlFrameCount(i).Value
                    Height = Tex_Animation(Animationnum).Height
                    
                    sRect.Top = 0
                    sRect.Bottom = Height
                    sRect.Left = (AnimEditorFrame(i) - 1) * Width
                    sRect.Right = sRect.Left + Width
                    
                    dRect.Top = 0
                    dRect.Bottom = Height
                    dRect.Left = 0
                    dRect.Right = Width
                    
                    RenderTextureByRects Tex_Animation(Animationnum), sRect, dRect
                    
                    With srcRect
                        .x1 = 0
                        .x2 = frmEditor_Animation.picSprite(i).Width
                        .y1 = 0
                        .y2 = frmEditor_Animation.picSprite(i).Height
                    End With
                                
                    With destRect
                        .x1 = 0
                        .x2 = frmEditor_Animation.picSprite(i).Width
                        .y1 = 0
                        .y2 = frmEditor_Animation.picSprite(i).Height
                    End With
                                
                    Direct3D_Device.EndScene
                    Direct3D_Device.Present srcRect, destRect, frmEditor_Animation.picSprite(i).hWnd, ByVal (0)
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorAnim_DrawAnim", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorNpc_DrawSprite()
Dim Sprite As Long, destRect As D3DRECT
Dim sRect As RECT
Dim dRect As RECT

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    Sprite = frmEditor_NPC.scrlSprite.Value

    If Sprite < 1 Or Sprite > NumCharacters Then
        frmEditor_NPC.picSprite.Cls
        Exit Sub
    End If

    sRect.Top = 0
    sRect.Bottom = SIZE_Y
    sRect.Left = PIC_X * 3 ' facing down
    sRect.Right = sRect.Left + SIZE_X
    dRect.Top = 0
    dRect.Bottom = SIZE_Y
    dRect.Left = 0
    dRect.Right = SIZE_X
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Character(Sprite), sRect, dRect
    
    With destRect
        .x1 = 0
        .x2 = SIZE_X
        .y1 = 0
        .y2 = SIZE_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRect, destRect, frmEditor_NPC.picSprite.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorNpc_DrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorResource_DrawSprite()
Dim Sprite As Long
Dim sRect As RECT, destRect As D3DRECT, srcRect As D3DRECT
Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    ' normal sprite
    Sprite = frmEditor_Resource.scrlNormalPic.Value

    If Sprite < 1 Or Sprite > NumResources Then
        frmEditor_Resource.picNormalPic.Cls
    Else
        sRect.Top = 0
        sRect.Bottom = Tex_Resource(Sprite).Height
        sRect.Left = 0
        sRect.Right = Tex_Resource(Sprite).Width
        dRect.Top = 0
        dRect.Bottom = Tex_Resource(Sprite).Height
        dRect.Left = 0
        dRect.Right = Tex_Resource(Sprite).Width
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
        RenderTextureByRects Tex_Resource(Sprite), sRect, dRect
        With srcRect
            .x1 = 0
            .x2 = Tex_Resource(Sprite).Width
            .y1 = 0
            .y2 = Tex_Resource(Sprite).Height
        End With
        
        With destRect
            .x1 = 0
            .x2 = frmEditor_Resource.picNormalPic.ScaleWidth
            .y1 = 0
            .y2 = frmEditor_Resource.picNormalPic.ScaleHeight
        End With
                    
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRect, frmEditor_Resource.picNormalPic.hWnd, ByVal (0)
    End If

    ' exhausted sprite
    Sprite = frmEditor_Resource.scrlExhaustedPic.Value

    If Sprite < 1 Or Sprite > NumResources Then
        frmEditor_Resource.picExhaustedPic.Cls
    Else
        sRect.Top = 0
        sRect.Bottom = Tex_Resource(Sprite).Height
        sRect.Left = 0
        sRect.Right = Tex_Resource(Sprite).Width
        dRect.Top = 0
        dRect.Bottom = Tex_Resource(Sprite).Height
        dRect.Left = 0
        dRect.Right = Tex_Resource(Sprite).Width
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
        RenderTextureByRects Tex_Resource(Sprite), sRect, dRect
        
        With destRect
            .x1 = 0
            .x2 = frmEditor_Resource.picExhaustedPic.ScaleWidth
            .y1 = 0
            .y2 = frmEditor_Resource.picExhaustedPic.ScaleHeight
        End With
        
        With srcRect
            .x1 = 0
            .x2 = Tex_Resource(Sprite).Width
            .y1 = 0
            .y2 = Tex_Resource(Sprite).Height
        End With
                    
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRect, frmEditor_Resource.picExhaustedPic.hWnd, ByVal (0)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorResource_DrawSprite", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Render_Graphics()
Dim X As Long
Dim Y As Long
Dim i As Long
Dim rec As RECT
Dim rec_pos As RECT, srcRect As D3DRECT
    
    ' If debug mode, handle error then exit out
   On Error GoTo errorhandler
    
    'Check for device lost.
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then HandleDeviceLost: Exit Sub
    
    ' don't render
    If frmMain.WindowState = vbMinimized Then Exit Sub
    If GettingMap Then Exit Sub
    
    ' update the viewpoint
    UpdateCamera

    
   Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorARGB(0, 0, 0, 0), 1#, 0
        
        Direct3D_Device.BeginScene
    
            ' blit lower tiles
            If NumTileSets > 0 Then
                For X = TileView.Left To TileView.Right
                    For Y = TileView.Top To TileView.Bottom
                        If IsValidMapPoint(X, Y) Then
                            Call DrawMapTile(X, Y)
                        End If
                    Next
                Next
            End If
        
            ' render the decals
            For i = 1 To MAX_BYTE
                Call DrawBlood(i)
            Next
        
            ' Blit out the items
            If numitems > 0 Then
                For i = 1 To MAX_MAP_ITEMS
                    If MapItem(i).Num > 0 Then
                        Call DrawItem(i)
                    End If
                Next
            End If
            
            If Map.CurrentEvents > 0 Then
                For i = 1 To Map.CurrentEvents
                    If Map.MapEvents(i).Position = 0 Then
                        DrawEvent i
                    End If
                Next
            End If
            
            ' draw animations
            If NumAnimations > 0 Then
                For i = 1 To MAX_BYTE
                    If AnimInstance(i).Used(0) Then
                        DrawAnimation i, 0
                    End If
                Next
            End If
         ' draw projectiles for each player
           For i = 1 To Player_HighIndex
             For X = 1 To MAX_PLAYER_PROJECTILES
                If Player(i).ProjecTile(X).Pic > 0 Then
                           DrawProjectile i, X
                   End If
                Next
             Next
            ' Y-based render. Renders Players, Npcs and Resources based on Y-axis.
            For Y = 0 To Map.MaxY
                If NumCharacters > 0 Then
                
                    If Map.CurrentEvents > 0 Then
                        For i = 1 To Map.CurrentEvents
                            If Map.MapEvents(i).Position = 1 Then
                                If Y = Map.MapEvents(i).Y Then
                                    DrawEvent i
                                End If
                            End If
                        Next
                    End If
                    
                    ' Players
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                            If Player(i).Y = Y Then
                                Call DrawPlayer(i)
                            End If
                        End If
                    Next
                    
                    
                
                    ' Npcs
                    For i = 1 To Npc_HighIndex
                        If MapNpc(i).Y = Y Then
                            Call DrawNpc(i)
                        End If
                    Next
                End If
                
                ' Resources
                If NumResources > 0 Then
                    If Resources_Init Then
                        If Resource_Index > 0 Then
                            For i = 1 To Resource_Index
                                If MapResource(i).Y = Y Then
                                    Call DrawMapResource(i)
                                End If
                            Next
                        End If
                    End If
                End If
            Next
            
            ' animations
            If NumAnimations > 0 Then
                For i = 1 To MAX_BYTE
                    If AnimInstance(i).Used(1) Then
                        DrawAnimation i, 1
                    End If
                Next
            End If
        
            ' blit out upper tiles
            If NumTileSets > 0 Then
                For X = TileView.Left To TileView.Right
                    For Y = TileView.Top To TileView.Bottom
                        If IsValidMapPoint(X, Y) Then
                            Call DrawMapFringeTile(X, Y)
                        End If
                    Next
                Next
            End If
            
            If Map.CurrentEvents > 0 Then
                For i = 1 To Map.CurrentEvents
                    If Map.MapEvents(i).Position = 2 Then
                        DrawEvent i
                    End If
                Next
            End If
            
            DrawWeather
            DrawFog
            DrawTint
            
            ' blit out a square at mouse cursor
            If InMapEditor Then
                If frmEditor_Map.optBlock.Value = True Then
                    For X = TileView.Left To TileView.Right
                        For Y = TileView.Top To TileView.Bottom
                            If IsValidMapPoint(X, Y) Then
                                Call DrawDirection(X, Y)
                            End If
                        Next
                    Next
                End If
                Call DrawTileOutline
            End If
            
            ' Render the bars
            DrawBars
            
            ' Draw the target icon
            If myTarget > 0 Then
                If myTargetType = TargetPlayer Then
                    DrawTarget (Player(myTarget).X * 32) + Player(myTarget).xOffset, (Player(myTarget).Y * 32) + Player(myTarget).yOffset
                ElseIf myTargetType = TargetNPC Then
                    DrawTarget (MapNpc(myTarget).X * 32) + MapNpc(myTarget).xOffset, (MapNpc(myTarget).Y * 32) + MapNpc(myTarget).yOffset
                End If
            End If
            
            ' Draw the hover icon
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If Player(i).Map = Player(MyIndex).Map Then
                        If CurX = Player(i).X And CurY = Player(i).Y Then
                            If myTargetType = TargetPlayer And myTarget = i Then
                                ' dont render lol
                            Else
                                DrawHover TargetPlayer, i, (Player(i).X * 32) + Player(i).xOffset, (Player(i).Y * 32) + Player(i).yOffset
                            End If
                        End If
                    End If
                End If
            Next
            For i = 1 To Npc_HighIndex
                If MapNpc(i).Num > 0 Then
                    If CurX = MapNpc(i).X And CurY = MapNpc(i).Y Then
                        If myTargetType = TargetNPC And myTarget = i Then
                            ' dont render lol
                        Else
                            DrawHover TargetNPC, i, (MapNpc(i).X * 32) + MapNpc(i).xOffset, (MapNpc(i).Y * 32) + MapNpc(i).yOffset
                        End If
                    End If
                End If
            Next
            
            If DrawThunder > 0 Then RenderTexture Tex_White, 0, 0, 0, 0, frmMain.picScreen.ScaleWidth, frmMain.picScreen.ScaleHeight, 32, 32, D3DColorRGBA(255, 255, 255, 160): DrawThunder = DrawThunder - 1
            
            ' Get rec
            With rec
                .Top = Camera.Top
                .Bottom = .Top + ScreenY
                .Left = Camera.Left
                .Right = .Left + ScreenX
            End With
                
            ' rec_pos
            With rec_pos
                .Bottom = ScreenY
                .Right = ScreenX
            End With
                
            With srcRect
                .x1 = 0
                .x2 = frmMain.picScreen.ScaleWidth
                .y1 = 0
                .y2 = frmMain.picScreen.ScaleHeight
            End With
            
            If BFPS Then
                RenderText Font_Default, "FPS: " & CStr(GameFPS), 2, 39, Yellow, 0
            End If
            
            ' draw cursor, player X and Y locations
            If BLoc Then
                RenderText Font_Default, Trim$("cur x: " & CurX & " y: " & CurY), 2, 1, Yellow, 0
                RenderText Font_Default, Trim$("loc x: " & GetPlayerX(MyIndex) & " y: " & GetPlayerY(MyIndex)), 2, 15, Yellow, 0
                RenderText Font_Default, Trim$(" (map #" & GetPlayerMap(MyIndex) & ")"), 2, 27, Yellow, 0
            End If
            
            ' draw player names
            For i = 1 To Player_HighIndex
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    Call DrawPlayerName(i)
                End If
            Next
            
            For i = 1 To Map.CurrentEvents
                If Map.MapEvents(i).Visible = 1 Then
                    If Map.MapEvents(i).ShowName = 1 Then
                        DrawEventName (i)
                    End If
                End If
            Next
            
            ' draw npc names
            For i = 1 To Npc_HighIndex
                If MapNpc(i).Num > 0 Then
                    Call DrawNpcName(i)
                End If
            Next
            
                ' draw the messages
            For i = 1 To MAX_BYTE
                If chatBubble(i).active Then
                    DrawChatBubble i
                End If
            Next
            
            For i = 1 To Action_HighIndex
                Call DrawActionMsg(i)
            Next i
     
            ' Draw map name
            RenderText Font_Default, Map.Name, DrawMapNameX, DrawMapNameY, DrawMapNameColor, 0
            
            If InMapEditor And frmEditor_Map.optEvent.Value = True Then DrawEvents
            If InMapEditor Then Call DrawMapAttributes
            
            If FadeAmount > 0 Then RenderTexture Tex_Fade, 0, 0, 0, 0, frmMain.picScreen.ScaleWidth, frmMain.picScreen.ScaleHeight, 32, 32, D3DColorRGBA(255, 255, 255, FadeAmount)
            If FlashTimer > timeGetTime Then RenderTexture Tex_White, 0, 0, 0, 0, frmMain.picScreen.ScaleWidth, frmMain.picScreen.ScaleHeight, 32, 32, -1
    
        Direct3D_Device.EndScene
        
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost
        Exit Sub
    Else
        If InShop = False And InBank = False Then Direct3D_Device.Present srcRect, ByVal 0, 0, ByVal 0
        DrawGDI
    End If

    ' Error handler
    Exit Sub
    
errorhandler:
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost
        Exit Sub
    Else
        If options.Debug = 1 Then
            HandleError "Render_Graphics", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
            Err.Clear
        End If
        MsgBox "Unrecoverable DX8 error."
        DestroyGame
    End If
End Sub

Sub HandleDeviceLost()
Dim i As Long
'Do a loop while device is lost
   Do While Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST
       Exit Sub
   Loop
   
   UnloadTextures
   
   'Reset the device
   Direct3D_Device.Reset Direct3D_Window
   
   DirectX_ReInit
    
   LoadTextures
    For i = 1 To NumTextures
        gTexture(i).timer = 0 ' Reset all textures, ready to render again!
      Next
   ' Load tilesets again, or we'll get a very blank screen.
   LoadTilesets
End Sub
Private Function DirectX_ReInit() As Boolean

    On Error GoTo Error_Handler

    Direct3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display_Mode 'Use the current display mode that you
                                                                    'are already on. Incase you are confused, I'm
                                                                    'talking about your current screen resolution. ;)
        
    Direct3D_Window.Windowed = True 'The app will be in windowed mode.

    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_COPY 'Refresh when the monitor does.
    Direct3D_Window.BackBufferFormat = Display_Mode.Format 'Sets the format that was retrieved into the backbuffer.
    'Creates the rendering device with some useful info, along with the info
    'we've already setup for Direct3D_Window.
    'Creates the rendering device with some useful info, along with the info
    Direct3D_Window.BackBufferCount = 1 '1 backbuffer only
    Direct3D_Window.BackBufferWidth = 960 ' frmMain.picScreen.ScaleWidth 'Match the backbuffer width with the display width
    Direct3D_Window.BackBufferHeight = 768 'frmMain.picScreen.ScaleHeight 'Match the backbuffer height with the display height
    Direct3D_Window.hDeviceWindow = frmMain.picScreen.hWnd 'Use frmMain as the device window.
    
    With Direct3D_Device
        .SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
    
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False
        
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
    End With
    
    DirectX_ReInit = True

    Exit Function
    
Error_Handler:
    MsgBox "An error occured while initializing DirectX", vbCritical
    
    DestroyGame
    
    DirectX_ReInit = False
End Function
Public Sub UpdateCamera()
Dim offsetX As Long
Dim offsetY As Long
Dim StartX As Long
Dim StartY As Long
Dim EndX As Long
Dim EndY As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    offsetX = Player(MyIndex).xOffset + PIC_X
    offsetY = Player(MyIndex).yOffset + PIC_Y

    StartX = GetPlayerX(MyIndex) - StartXValue
    StartY = GetPlayerY(MyIndex) - StartYValue
    If StartX < 0 Then
        offsetX = 0
        If StartX = -1 Then
            If Player(MyIndex).xOffset > 0 Then
                offsetX = Player(MyIndex).xOffset
            End If
        End If
        StartX = 0
    End If
    If StartY < 0 Then
        offsetY = 0
        If StartY = -1 Then
            If Player(MyIndex).yOffset > 0 Then
                offsetY = Player(MyIndex).yOffset
            End If
        End If
        StartY = 0
    End If
    
    EndX = StartX + EndXValue
    EndY = StartY + EndYValue
    If EndX > Map.MaxX Then
        offsetX = 32
        If EndX = Map.MaxX + 1 Then
            If Player(MyIndex).xOffset < 0 Then
                offsetX = Player(MyIndex).xOffset + PIC_X
            End If
        End If
        EndX = Map.MaxX
        StartX = EndX - MAX_MAPX - 1
    End If
    If EndY > Map.MaxY Then
        offsetY = 32
        If EndY = Map.MaxY + 1 Then
            If Player(MyIndex).yOffset < 0 Then
                offsetY = Player(MyIndex).yOffset + PIC_Y
            End If
        End If
        EndY = Map.MaxY
        StartY = EndY - MAX_MAPY - 1
    End If

    With TileView
        .Top = StartY
        .Bottom = EndY
        .Left = StartX
        .Right = EndX
    End With

    With Camera
        .Top = offsetY
        .Bottom = .Top + ScreenY
        .Left = offsetX
        .Right = .Left + ScreenX
    End With
    
    UpdateDrawMapName

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateCamera", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function ConvertMapX(ByVal X As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    ConvertMapX = X - (TileView.Left * PIC_X) - Camera.Left
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertMapX", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ConvertMapY(ByVal Y As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    ConvertMapY = Y - (TileView.Top * PIC_Y) - Camera.Top
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertMapY", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function InViewPort(ByVal X As Long, ByVal Y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    InViewPort = False

    If X < TileView.Left Then Exit Function
    If Y < TileView.Top Then Exit Function
    If X > TileView.Right Then Exit Function
    If Y > TileView.Bottom Then Exit Function
    InViewPort = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "InViewPort", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsValidMapPoint(ByVal X As Long, ByVal Y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    IsValidMapPoint = False

    If X < 0 Then Exit Function
    If Y < 0 Then Exit Function
    If X > Map.MaxX Then Exit Function
    If Y > Map.MaxY Then Exit Function
    IsValidMapPoint = True
        
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsValidMapPoint", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub LoadTilesets()
Dim X As Long
Dim Y As Long
Dim i As Long
Dim tilesetInUse() As Boolean
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    ReDim tilesetInUse(0 To NumTileSets)
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                ' check exists
                If Map.Tile(X, Y).Layer(i).Tileset > 0 And Map.Tile(X, Y).Layer(i).Tileset <= NumTileSets Then
                    tilesetInUse(Map.Tile(X, Y).Layer(i).Tileset) = True
                End If
            Next
        Next
    Next
    
    For i = 1 To NumTileSets
        If tilesetInUse(i) Then
            LoadTexture Tex_Tileset(i)
        ' Are we currently selecting it in the map editor?
        ElseIf InMapEditor = True And frmEditor_Map.scrlTileSet.Value = i Then
            LoadTexture Tex_Tileset(i)
        ' We're not using it anywhere right now, so unload it.
        Else
            Set gTexture(Tex_Tileset(i).Texture).Texture = Nothing
            ZeroMemory ByVal VarPtr(gTexture(Tex_Tileset(i).Texture)), LenB(gTexture(Tex_Tileset(i).Texture))
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadTilesets", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DrawBank()
Dim i As Long, X As Long, Y As Long, itemnum As Long, srcRect As D3DRECT, destRect As D3DRECT
Dim Amount As String
Dim sRect As RECT, dRect As RECT
Dim Sprite As Long, colour As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If frmMain.picBank.Visible = True Then
        'frmMain.picBank.Cls
        
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
                
        For i = 1 To MAX_BANK
            itemnum = GetBankItemNum(i)
            If itemnum > 0 And itemnum <= MAX_ITEMS Then
            
                Sprite = Item(itemnum).Pic
                
                If Sprite <= 0 Or Sprite > numitems Then Exit Sub
            
                With sRect
                    .Top = 0
                    .Bottom = .Top + PIC_Y
                    .Left = Tex_Item(Sprite).Width / 2
                    .Right = .Left + PIC_X
                End With
                
                With dRect
                    .Top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                    .Bottom = .Top + PIC_Y
                    .Left = BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                    .Right = .Left + PIC_X
                End With
                
                RenderTextureByRects Tex_Item(Sprite), sRect, dRect

                ' If item is a stack - draw the amount you have
                If GetBankItemValue(i) > 1 Then
                    Y = dRect.Top + 22
                    X = dRect.Left - 4
                
                    Amount = CStr(GetBankItemValue(i))
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        colour = BrightGreen
                    End If
                    RenderText Font_Default, ConvertCurrency(Amount), X, Y, colour
                End If
            End If
        Next
        
        With srcRect
            .x1 = BankLeft
            .x2 = .x1 + 400
            .y1 = BankTop
            .y2 = .y1 + 310
        End With
                    
        With destRect
            .x1 = BankLeft
            .x2 = .x1 + 400
            .y1 = BankTop
            .y2 = 310 + .y1
        End With
                    
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRect, frmMain.picBank.hWnd, ByVal (0)
        'frmMain.picBank.Refresh
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawBank", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawBankItem(ByVal X As Long, ByVal Y As Long)
Dim sRect As RECT, dRect As RECT, srcRect As D3DRECT, destRect As D3DRECT
Dim itemnum As Long
Dim Sprite As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    itemnum = GetBankItemNum(DragBankSlotNum)
    Sprite = Item(GetBankItemNum(DragBankSlotNum)).Pic
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 255), 1#, 0
    Direct3D_Device.BeginScene
    
    If itemnum > 0 Then
        If itemnum <= MAX_ITEMS Then
            With sRect
                .Top = 0
                .Bottom = .Top + PIC_Y
                .Left = Tex_Item(Sprite).Width / 2
                .Right = .Left + PIC_X
            End With
        End If
    End If
    
    With dRect
        .Top = 2
        .Bottom = .Top + PIC_Y
        .Left = 2
        .Right = .Left + PIC_X
    End With

    RenderTextureByRects Tex_Item(Sprite), sRect, dRect
    
    With frmMain.picTempBank
        .Top = Y
        .Left = X
        .Visible = True
        .ZOrder (0)
    End With
    
    With srcRect
        .x1 = 0
        .x2 = 32
        .y1 = 0
        .y2 = 32
    End With
    With destRect
        .x1 = 2
        .y1 = 2
        .y2 = .y1 + 32
        .x2 = .x1 + 32
    End With
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, destRect, frmMain.picTempBank.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawBankItem", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawEvents()
Dim sRect As RECT
Dim Width As Long, Height As Long, i As Long, X As Long, Y As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If Map.EventCount <= 0 Then Exit Sub
    
    For i = 1 To Map.EventCount
        If Map.Events(i).pageCount <= 0 Then
                sRect.Top = 0
                sRect.Bottom = 32
                sRect.Left = 0
                sRect.Right = 32
                RenderTexture Tex_Selection, ConvertMapX(X), ConvertMapY(Y), sRect.Left, sRect.Right, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
            GoTo nextevent
        End If
        
        Width = 32
        Height = 32
    
        X = Map.Events(i).X * 32
        Y = Map.Events(i).Y * 32
        X = ConvertMapX(X)
        Y = ConvertMapY(Y)
    
        
        If i > Map.EventCount Then Exit Sub
        If 1 > Map.Events(i).pageCount Then Exit Sub
        Select Case Map.Events(i).Pages(1).GraphicType
            Case 0
                sRect.Top = 0
                sRect.Bottom = 32
                sRect.Left = 0
                sRect.Right = 32
                RenderTexture Tex_Selection, X, Y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
            Case 1
                If Map.Events(i).Pages(1).Graphic > 0 And Map.Events(i).Pages(1).Graphic <= NumCharacters Then
                    
                    sRect.Top = (Map.Events(i).Pages(1).GraphicY * (Tex_Character(Map.Events(i).Pages(1).Graphic).Height / 4))
                    sRect.Left = (Map.Events(i).Pages(1).GraphicX * (Tex_Character(Map.Events(i).Pages(1).Graphic).Width / 4))
                    sRect.Bottom = sRect.Top + 32
                    sRect.Right = sRect.Left + 32
                    RenderTexture Tex_Character(Map.Events(i).Pages(1).Graphic), X, Y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                    
                    sRect.Top = 0
                    sRect.Bottom = 32
                    sRect.Left = 0
                    sRect.Right = 32
                    RenderTexture Tex_Selection, X, Y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                Else
                    sRect.Top = 0
                    sRect.Bottom = 32
                    sRect.Left = 0
                    sRect.Right = 32
                    RenderTexture Tex_Selection, X, Y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
            Case 2
                If Map.Events(i).Pages(1).Graphic > 0 And Map.Events(i).Pages(1).Graphic < NumTileSets Then
                    sRect.Top = Map.Events(i).Pages(1).GraphicY * 32
                    sRect.Left = Map.Events(i).Pages(1).GraphicX * 32
                    sRect.Bottom = sRect.Top + 32
                    sRect.Right = sRect.Left + 32
                    RenderTexture Tex_Tileset(Map.Events(i).Pages(1).Graphic), X, Y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                    
                    sRect.Top = 0
                    sRect.Bottom = 32
                    sRect.Left = 0
                    sRect.Right = 32
                    RenderTexture Tex_Selection, X, Y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                Else
                    sRect.Top = 0
                    sRect.Bottom = 32
                    sRect.Left = 0
                    sRect.Right = 32
                    RenderTexture Tex_Selection, X, Y, sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
        End Select
nextevent:
    Next
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawEvents", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorEvent_DrawGraphic()
Dim sRect As RECT, destRect As D3DRECT, srcRect As D3DRECT
Dim dRect As RECT

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Events.picGraphicSel.Visible Then
        Select Case frmEditor_Events.cmbGraphic.ListIndex
            Case 0
                'None
                frmEditor_Events.picGraphicSel.Cls
                Exit Sub
            Case 1
                If frmEditor_Events.scrlGraphic.Value > 0 And frmEditor_Events.scrlGraphic.Value <= NumCharacters Then
                    
                    If Tex_Character(frmEditor_Events.scrlGraphic.Value).Width > 793 Then
                        sRect.Left = frmEditor_Events.hScrlGraphicSel.Value
                        sRect.Right = sRect.Left + (Tex_Character(frmEditor_Events.scrlGraphic.Value).Width - sRect.Left)
                    Else
                        sRect.Left = 0
                        sRect.Right = Tex_Character(frmEditor_Events.scrlGraphic.Value).Width
                    End If
                    
                    If Tex_Character(frmEditor_Events.scrlGraphic.Value).Height > 472 Then
                        sRect.Top = frmEditor_Events.hScrlGraphicSel.Value
                        sRect.Bottom = sRect.Top + (Tex_Character(frmEditor_Events.scrlGraphic.Value).Height - sRect.Top)
                    Else
                        sRect.Top = 0
                        sRect.Bottom = Tex_Character(frmEditor_Events.scrlGraphic.Value).Height
                    End If
                    
                    With dRect
                        .Top = 0
                        .Bottom = sRect.Bottom - sRect.Top
                        .Left = 0
                        .Right = sRect.Right - sRect.Left
                    End With
                    
                    With destRect
                        .x1 = dRect.Left
                        .x2 = dRect.Right
                        .y1 = dRect.Top
                        .y2 = dRect.Bottom
                    End With
                    
                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    RenderTextureByRects Tex_Character(frmEditor_Events.scrlGraphic.Value), sRect, dRect
                    If (GraphicSelX2 < GraphicSelX Or GraphicSelY2 < GraphicSelY) Or (GraphicSelX2 = 0 And GraphicSelY2 = 0) Then
                        With destRect
                            .x1 = (GraphicSelX * (Tex_Character(frmEditor_Events.scrlGraphic.Value).Width / 4)) - sRect.Left
                            .x2 = (Tex_Character(frmEditor_Events.scrlGraphic.Value).Width / 4) + .x1
                            .y1 = (GraphicSelY * (Tex_Character(frmEditor_Events.scrlGraphic.Value).Height / 4)) - sRect.Top
                            .y2 = (Tex_Character(frmEditor_Events.scrlGraphic.Value).Height / 4) + .y1
                        End With

                    Else
                        With destRect
                            .x1 = (GraphicSelX * 32) - sRect.Left
                            .x2 = ((GraphicSelX2 - GraphicSelX) * 32) + .x1
                            .y1 = (GraphicSelY * 32) - sRect.Top
                            .y2 = ((GraphicSelY2 - GraphicSelY) * 32) + .y1
                        End With
                    End If
                    DrawSelectionBox destRect
                    
                    With srcRect
                        .x1 = dRect.Left
                        .x2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .y1 = dRect.Top
                        .y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    With destRect
                        .x1 = 0
                        .y1 = 0
                        .x2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    Direct3D_Device.EndScene
                    Direct3D_Device.Present srcRect, destRect, frmEditor_Events.picGraphicSel.hWnd, ByVal (0)
                    
                    If GraphicSelX <= 3 And GraphicSelY <= 3 Then
                    Else
                        GraphicSelX = 0
                        GraphicSelY = 0
                    End If
                Else
                    frmEditor_Events.picGraphicSel.Cls
                    Exit Sub
                End If
            Case 2
                If frmEditor_Events.scrlGraphic.Value > 0 And frmEditor_Events.scrlGraphic.Value <= NumTileSets Then
                    
                    If Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Width > 793 Then
                        sRect.Left = frmEditor_Events.hScrlGraphicSel.Value
                        sRect.Right = sRect.Left + 800
                    Else
                        sRect.Left = 0
                        sRect.Right = Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Width
                        sRect.Left = frmEditor_Events.hScrlGraphicSel.Value = 0
                    End If
                    
                    If Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Height > 472 Then
                        sRect.Top = frmEditor_Events.vScrlGraphicSel.Value
                        sRect.Bottom = sRect.Top + 512
                    Else
                        sRect.Top = 0
                        sRect.Bottom = Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Height
                        frmEditor_Events.vScrlGraphicSel.Value = 0
                    End If
                    
                    If sRect.Left = -1 Then sRect.Left = 0
                    If sRect.Top = -1 Then sRect.Top = 0
                    
                    With dRect
                        .Top = 0
                        .Bottom = sRect.Bottom - sRect.Top
                        .Left = 0
                        .Right = sRect.Right - sRect.Left
                    End With
                    
                    
                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    RenderTextureByRects Tex_Tileset(frmEditor_Events.scrlGraphic.Value), sRect, dRect
                    If (GraphicSelX2 < GraphicSelX Or GraphicSelY2 < GraphicSelY) Or (GraphicSelX2 = 0 And GraphicSelY2 = 0) Then
                        With destRect
                            .x1 = (GraphicSelX * 32) - sRect.Left
                            .x2 = PIC_X + .x1
                            .y1 = (GraphicSelY * 32) - sRect.Top
                            .y2 = PIC_Y + .y1
                        End With

                    Else
                        With destRect
                            .x1 = (GraphicSelX * 32) - sRect.Left
                            .x2 = ((GraphicSelX2 - GraphicSelX) * 32) + .x1
                            .y1 = (GraphicSelY * 32) - sRect.Top
                            .y2 = ((GraphicSelY2 - GraphicSelY) * 32) + .y1
                        End With
                    End If
                    
                    DrawSelectionBox destRect
                    
                    With srcRect
                        .x1 = dRect.Left
                        .x2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .y1 = dRect.Top
                        .y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    With destRect
                        .x1 = 0
                        .y1 = 0
                        .x2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    Direct3D_Device.EndScene
                    Direct3D_Device.Present srcRect, destRect, frmEditor_Events.picGraphicSel.hWnd, ByVal (0)
                Else
                    frmEditor_Events.picGraphicSel.Cls
                    Exit Sub
                End If
        End Select
    Else
        Select Case tmpEvent.Pages(curPageNum).GraphicType
            Case 0
                frmEditor_Events.picGraphic.Cls
            Case 1
                If tmpEvent.Pages(curPageNum).Graphic > 0 And tmpEvent.Pages(curPageNum).Graphic <= NumCharacters Then
                    sRect.Top = tmpEvent.Pages(curPageNum).GraphicY * (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Height / 4)
                    sRect.Left = tmpEvent.Pages(curPageNum).GraphicX * (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Width / 4)
                    sRect.Bottom = sRect.Top + (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Height / 4)
                    sRect.Right = sRect.Left + (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Width / 4)
                    With dRect
                        dRect.Top = (193 / 2) - ((sRect.Bottom - sRect.Top) / 2)
                        dRect.Bottom = dRect.Top + (sRect.Bottom - sRect.Top)
                        dRect.Left = (121 / 2) - ((sRect.Right - sRect.Left) / 2)
                        dRect.Right = dRect.Left + (sRect.Right - sRect.Left)
                    End With
                    With destRect
                        .x1 = dRect.Left
                        .x2 = dRect.Right
                        .y1 = dRect.Top
                        .y2 = dRect.Bottom
                    End With
                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    RenderTextureByRects Tex_Character(frmEditor_Events.scrlGraphic.Value), sRect, dRect
                    Direct3D_Device.EndScene
                    Direct3D_Device.Present destRect, destRect, frmEditor_Events.picGraphic.hWnd, ByVal (0)
                End If
            Case 2
                If tmpEvent.Pages(curPageNum).Graphic > 0 And tmpEvent.Pages(curPageNum).Graphic <= NumTileSets Then
                    If tmpEvent.Pages(curPageNum).GraphicX2 = 0 Or tmpEvent.Pages(curPageNum).GraphicY2 = 0 Then
                        sRect.Top = tmpEvent.Pages(curPageNum).GraphicY * 32
                        sRect.Left = tmpEvent.Pages(curPageNum).GraphicX * 32
                        sRect.Bottom = sRect.Top + 32
                        sRect.Right = sRect.Left + 32
                        With dRect
                            dRect.Top = (193 / 2) - ((sRect.Bottom - sRect.Top) / 2)
                            dRect.Bottom = dRect.Top + (sRect.Bottom - sRect.Top)
                            dRect.Left = (120 / 2) - ((sRect.Right - sRect.Left) / 2)
                            dRect.Right = dRect.Left + (sRect.Right - sRect.Left)
                        End With
                        With destRect
                            .x1 = dRect.Left
                            .x2 = dRect.Right
                            .y1 = dRect.Top
                            .y2 = dRect.Bottom
                        End With
                        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                        Direct3D_Device.BeginScene
                        RenderTextureByRects Tex_Tileset(frmEditor_Events.scrlGraphic.Value), sRect, dRect
                        Direct3D_Device.EndScene
                        Direct3D_Device.Present destRect, destRect, frmEditor_Events.picGraphic.hWnd, ByVal (0)
                    Else
                        sRect.Top = tmpEvent.Pages(curPageNum).GraphicY * 32
                        sRect.Left = tmpEvent.Pages(curPageNum).GraphicX * 32
                        sRect.Bottom = sRect.Top + ((tmpEvent.Pages(curPageNum).GraphicY2 - tmpEvent.Pages(curPageNum).GraphicY) * 32)
                        sRect.Right = sRect.Left + ((tmpEvent.Pages(curPageNum).GraphicX2 - tmpEvent.Pages(curPageNum).GraphicX) * 32)
                        With dRect
                            dRect.Top = (193 / 2) - ((sRect.Bottom - sRect.Top) / 2)
                            dRect.Bottom = dRect.Top + (sRect.Bottom - sRect.Top)
                            dRect.Left = (120 / 2) - ((sRect.Right - sRect.Left) / 2)
                            dRect.Right = dRect.Left + (sRect.Right - sRect.Left)
                        End With
                        With destRect
                            .x1 = dRect.Left
                            .x2 = dRect.Right
                            .y1 = dRect.Top
                            .y2 = dRect.Bottom
                        End With
                        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                        Direct3D_Device.BeginScene
                        RenderTextureByRects Tex_Tileset(frmEditor_Events.scrlGraphic.Value), sRect, dRect
                        Direct3D_Device.EndScene
                        Direct3D_Device.Present destRect, destRect, frmEditor_Events.picGraphic.hWnd, ByVal (0)
                    End If
                End If
        End Select
    End If
    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_DrawKey", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawEvent(ID As Long)
    Dim X As Long, Y As Long, Width As Long, Height As Long, sRect As RECT, dRect As RECT, anim As Long, spritetop As Long
    If Map.MapEvents(ID).Visible = 0 Then Exit Sub
    If InMapEditor Then Exit Sub
    Select Case Map.MapEvents(ID).GraphicType
        Case 0
            Exit Sub
            
        Case 1
            If Map.MapEvents(ID).GraphicNum <= 0 Or Map.MapEvents(ID).GraphicNum > NumCharacters Then Exit Sub
            Width = Tex_Character(Map.MapEvents(ID).GraphicNum).Width / 4
            Height = Tex_Character(Map.MapEvents(ID).GraphicNum).Height / 4
            ' Reset frame
            If Map.MapEvents(ID).Step = 3 Then
                anim = 0
            ElseIf Map.MapEvents(ID).Step = 1 Then
                anim = 2
            End If
            
            Select Case Map.MapEvents(ID).Dir
                Case DirectionUp
                    If (Map.MapEvents(ID).yOffset > 8) Then anim = Map.MapEvents(ID).Step
                Case DirectionDown
                    If (Map.MapEvents(ID).yOffset < -8) Then anim = Map.MapEvents(ID).Step
                Case DirectionLeft
                    If (Map.MapEvents(ID).xOffset > 8) Then anim = Map.MapEvents(ID).Step
                Case DirectionRight
                    If (Map.MapEvents(ID).xOffset < -8) Then anim = Map.MapEvents(ID).Step
            End Select
            
            ' Set the left
            Select Case Map.MapEvents(ID).ShowDir
                Case DirectionUp
                    spritetop = 3
                Case DirectionRight
                    spritetop = 2
                Case DirectionDown
                    spritetop = 0
                Case DirectionLeft
                    spritetop = 1
            End Select
            
            If Map.MapEvents(ID).WalkAnim = 1 Then anim = 0
            
            If Map.MapEvents(ID).Moving = 0 Then anim = Map.MapEvents(ID).GraphicX
            
            With sRect
                .Top = spritetop * Height
                .Bottom = .Top + Height
                .Left = anim * Width
                .Right = .Left + Width
            End With
        
            ' Calculate the X
            X = Map.MapEvents(ID).X * PIC_X + Map.MapEvents(ID).xOffset - ((Width - 32) / 2)
        
            ' Is the player's height more than 32..?
            If (Height * 4) > 32 Then
                ' Create a 32 pixel offset for larger sprites
                Y = Map.MapEvents(ID).Y * PIC_Y + Map.MapEvents(ID).yOffset - ((Height) - 32)
            Else
                ' Proceed as normal
                Y = Map.MapEvents(ID).Y * PIC_Y + Map.MapEvents(ID).yOffset
            End If
        
            ' render the actual sprite
            Call DrawSprite(Map.MapEvents(ID).GraphicNum, X, Y, sRect)
            
        Case 2
            If Map.MapEvents(ID).GraphicNum < 1 Or Map.MapEvents(ID).GraphicNum > NumTileSets Then Exit Sub
            
            If Map.MapEvents(ID).GraphicY2 > 0 Or Map.MapEvents(ID).GraphicX2 > 0 Then
                With sRect
                    .Top = Map.MapEvents(ID).GraphicY * 32
                    .Bottom = .Top + ((Map.MapEvents(ID).GraphicY2 - Map.MapEvents(ID).GraphicY) * 32)
                    .Left = Map.MapEvents(ID).GraphicX * 32
                    .Right = .Left + ((Map.MapEvents(ID).GraphicX2 - Map.MapEvents(ID).GraphicX) * 32)
                End With
            Else
                With sRect
                    .Top = Map.MapEvents(ID).GraphicY * 32
                    .Bottom = .Top + 32
                    .Left = Map.MapEvents(ID).GraphicX * 32
                    .Right = .Left + 32
                End With
            End If
            
            X = Map.MapEvents(ID).X * 32
            Y = Map.MapEvents(ID).Y * 32
            
            X = X - ((sRect.Right - sRect.Left) / 2)
            Y = Y - (sRect.Bottom - sRect.Top) + 32
            
            
            If Map.MapEvents(ID).GraphicY2 > 0 Then
                RenderTexture Tex_Tileset(Map.MapEvents(ID).GraphicNum), ConvertMapX(Map.MapEvents(ID).X * 32), ConvertMapY((Map.MapEvents(ID).Y - ((Map.MapEvents(ID).GraphicY2 - Map.MapEvents(ID).GraphicY) - 1)) * 32), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
            Else
                RenderTexture Tex_Tileset(Map.MapEvents(ID).GraphicNum), ConvertMapX(Map.MapEvents(ID).X * 32), ConvertMapY(Map.MapEvents(ID).Y * 32), sRect.Left, sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, sRect.Right - sRect.Left, sRect.Bottom - sRect.Top, D3DColorRGBA(255, 255, 255, 255)
            End If
    End Select
End Sub

'This function will make it much easier to setup the vertices with the info it needs.
Private Function Create_TLVertex(X As Single, Y As Single, Z As Single, RHW As Single, color As Long, Specular As Long, TU As Single, TV As Single) As TLVERTEX

    Create_TLVertex.X = X
    Create_TLVertex.Y = Y
    Create_TLVertex.Z = Z
    Create_TLVertex.RHW = RHW
    Create_TLVertex.color = color
    'Create_TLVertex.Specular = Specular
    Create_TLVertex.TU = TU
    Create_TLVertex.TV = TV
    
End Function

Public Function Ceiling(dblValIn As Double, dblCeilIn As Double) As Double
' round it
Ceiling = Round(dblValIn / dblCeilIn, 0) * dblCeilIn
' if it rounded down, force it up
If Ceiling < dblValIn Then Ceiling = Ceiling + dblCeilIn
End Function

Public Sub DestroyDX8()
    UnloadTextures
    Set Direct3DX = Nothing
    Set Direct3D_Device = Nothing
    Set Direct3D = Nothing
    Set DirectX8 = Nothing
End Sub

Public Sub DrawGDI()
    'Cycle Through in-game stuff before cycling through editors
    If frmMenu.Visible Then
        If frmMenu.picCharacter.Visible Then NewCharacterDrawSprite
    End If
    
    If frmMain.Visible Then
        If frmMain.picTempInv.Visible Then DrawInventoryItem frmMain.picTempInv.Left, frmMain.picTempInv.Top
        If frmMain.picTempSpell.Visible Then DrawDraggedSpell frmMain.picTempSpell.Left, frmMain.picTempSpell.Top
        If frmMain.picSpellDesc.Visible Then DrawSpellDesc LastSpellDesc
        If frmMain.picItemDesc.Visible Then DrawItemDesc LastItemDesc
        If frmMain.picHotbar.Visible Then DrawHotbar
        If frmMain.picInventory.Visible Then DrawInventory
        If frmMain.picItemDesc.Visible Then DrawItemDesc LastItemDesc
        If frmMain.picCharacter.Visible Then DrawFace: DrawEquipment
        If frmMain.picSpells.Visible Then DrawPlayerSpells
        If frmMain.picShop.Visible Then DrawShop
        If frmMain.picTempBank.Visible Then DrawBankItem frmMain.picTempBank.Left, frmMain.picTempBank.Top
        If frmMain.picBank.Visible Then DrawBank
        If frmMain.picTrade.Visible Then DrawTrade
    End If
    
    
    If frmEditor_Animation.Visible Then
        EditorAnim_DrawAnim
    End If
    
    If frmEditor_Item.Visible Then
        EditorItem_DrawItem
        EditorItem_DrawPaperdoll
    End If
    
    If frmEditor_Map.Visible Then
        EditorMap_DrawTileset
        If frmEditor_Map.fraMapItem.Visible Then EditorMap_DrawMapItem
        If frmEditor_Map.fraMapKey.Visible Then EditorMap_DrawKey
    End If
    
    If frmEditor_NPC.Visible Then
        EditorNpc_DrawSprite
    End If
    
    If frmEditor_Resource.Visible Then
        EditorResource_DrawSprite
    End If
    
    If frmEditor_Spell.Visible Then
        EditorSpell_DrawIcon
    End If
    
    If frmEditor_Events.Visible Then
        EditorEvent_DrawGraphic
    End If
End Sub


'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'   All of this code is for auto tiles and the math behind generating them.
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Sub placeAutotile(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long, ByVal tileQuarter As Byte, ByVal autoTileLetter As String)
    With Autotile(X, Y).Layer(layerNum).QuarterTile(tileQuarter)
        Select Case autoTileLetter
            Case "a"
                .X = autoInner(1).X
                .Y = autoInner(1).Y
            Case "b"
                .X = autoInner(2).X
                .Y = autoInner(2).Y
            Case "c"
                .X = autoInner(3).X
                .Y = autoInner(3).Y
            Case "d"
                .X = autoInner(4).X
                .Y = autoInner(4).Y
            Case "e"
                .X = autoNW(1).X
                .Y = autoNW(1).Y
            Case "f"
                .X = autoNW(2).X
                .Y = autoNW(2).Y
            Case "g"
                .X = autoNW(3).X
                .Y = autoNW(3).Y
            Case "h"
                .X = autoNW(4).X
                .Y = autoNW(4).Y
            Case "i"
                .X = autoNE(1).X
                .Y = autoNE(1).Y
            Case "j"
                .X = autoNE(2).X
                .Y = autoNE(2).Y
            Case "k"
                .X = autoNE(3).X
                .Y = autoNE(3).Y
            Case "l"
                .X = autoNE(4).X
                .Y = autoNE(4).Y
            Case "m"
                .X = autoSW(1).X
                .Y = autoSW(1).Y
            Case "n"
                .X = autoSW(2).X
                .Y = autoSW(2).Y
            Case "o"
                .X = autoSW(3).X
                .Y = autoSW(3).Y
            Case "p"
                .X = autoSW(4).X
                .Y = autoSW(4).Y
            Case "q"
                .X = autoSE(1).X
                .Y = autoSE(1).Y
            Case "r"
                .X = autoSE(2).X
                .Y = autoSE(2).Y
            Case "s"
                .X = autoSE(3).X
                .Y = autoSE(3).Y
            Case "t"
                .X = autoSE(4).X
                .Y = autoSE(4).Y
        End Select
    End With
End Sub

Public Sub initAutotiles()
Dim X As Long, Y As Long, layerNum As Long
    ' Procedure used to cache autotile positions. All positioning is
    ' independant from the tileset. Calculations are convoluted and annoying.
    ' Maths is not my strong point. Luckily we're caching them so it's a one-off
    ' thing when the map is originally loaded. As such optimisation isn't an issue.
    
    ' For simplicity's sake we cache all subtile SOURCE positions in to an array.
    ' We also give letters to each subtile for easy rendering tweaks. ;]
    
    ' First, we need to re-size the array
    ReDim Autotile(0 To Map.MaxX, 0 To Map.MaxY)
    
    ' Inner tiles (Top right subtile region)
    ' NW - a
    autoInner(1).X = 32
    autoInner(1).Y = 0
    
    ' NE - b
    autoInner(2).X = 48
    autoInner(2).Y = 0
    
    ' SW - c
    autoInner(3).X = 32
    autoInner(3).Y = 16
    
    ' SE - d
    autoInner(4).X = 48
    autoInner(4).Y = 16
    
    ' Outer Tiles - NW (bottom subtile region)
    ' NW - e
    autoNW(1).X = 0
    autoNW(1).Y = 32
    
    ' NE - f
    autoNW(2).X = 16
    autoNW(2).Y = 32
    
    ' SW - g
    autoNW(3).X = 0
    autoNW(3).Y = 48
    
    ' SE - h
    autoNW(4).X = 16
    autoNW(4).Y = 48
    
    ' Outer Tiles - NE (bottom subtile region)
    ' NW - i
    autoNE(1).X = 32
    autoNE(1).Y = 32
    
    ' NE - g
    autoNE(2).X = 48
    autoNE(2).Y = 32
    
    ' SW - k
    autoNE(3).X = 32
    autoNE(3).Y = 48
    
    ' SE - l
    autoNE(4).X = 48
    autoNE(4).Y = 48
    
    ' Outer Tiles - SW (bottom subtile region)
    ' NW - m
    autoSW(1).X = 0
    autoSW(1).Y = 64
    
    ' NE - n
    autoSW(2).X = 16
    autoSW(2).Y = 64
    
    ' SW - o
    autoSW(3).X = 0
    autoSW(3).Y = 80
    
    ' SE - p
    autoSW(4).X = 16
    autoSW(4).Y = 80
    
    ' Outer Tiles - SE (bottom subtile region)
    ' NW - q
    autoSE(1).X = 32
    autoSE(1).Y = 64
    
    ' NE - r
    autoSE(2).X = 48
    autoSE(2).Y = 64
    
    ' SW - s
    autoSE(3).X = 32
    autoSE(3).Y = 80
    
    ' SE - t
    autoSE(4).X = 48
    autoSE(4).Y = 80
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For layerNum = 1 To MapLayer.Layer_Count - 1
                ' calculate the subtile positions and place them
                CalculateAutotile X, Y, layerNum
                ' cache the rendering state of the tiles and set them
                CacheRenderState X, Y, layerNum
            Next
        Next
    Next
End Sub

Public Sub CacheRenderState(ByVal X As Long, ByVal Y As Long, ByVal layerNum As Long)
Dim quarterNum As Long

    ' exit out early
    If X < 0 Or X > Map.MaxX Or Y < 0 Or Y > Map.MaxY Then Exit Sub

    With Map.Tile(X, Y)
        ' check if the tile can be rendered
        If .Layer(layerNum).Tileset <= 0 Or .Layer(layerNum).Tileset > NumTileSets Then
            Autotile(X, Y).Layer(layerNum).renderState = RenderNone
            Exit Sub
        End If
        
        ' check if it's a key - hide mask if key is closed
        If layerNum = MapLayer.Mask Then
            If .Type = TileKey Then
                If TempTile(X, Y).DoorOpen = NO Then
                    Autotile(X, Y).Layer(layerNum).renderState = RenderNone
                    Exit Sub
                Else
                    Autotile(X, Y).Layer(layerNum).renderState = RenderNormal
                    Exit Sub
                End If
            End If
        End If
        
        ' check if it needs to be rendered as an autotile
        If .Autotile(layerNum) = ATNone Or .Autotile(layerNum) = ATFake Then
            ' default to... default
            Autotile(X, Y).Layer(layerNum).renderState = RenderNormal
        Else
            Autotile(X, Y).Layer(layerNum).renderState = RenderAutotile
            ' cache tileset positioning
            For quarterNum = 1 To 4
                Autotile(X, Y).Layer(layerNum).srcX(quarterNum) = (Map.Tile(X, Y).Layer(layerNum).X * 32) + Autotile(X, Y).Layer(layerNum).QuarterTile(quarterNum).X
                Autotile(X, Y).Layer(layerNum).srcY(quarterNum) = (Map.Tile(X, Y).Layer(layerNum).Y * 32) + Autotile(X, Y).Layer(layerNum).QuarterTile(quarterNum).Y
            Next
        End If
    End With
End Sub

Public Sub CalculateAutotile(ByVal X As Long, ByVal Y As Long, ByVal layerNum As Long)
    ' Right, so we've split the tile block in to an easy to remember
    ' collection of letters. We now need to do the calculations to find
    ' out which little lettered block needs to be rendered. We do this
    ' by reading the surrounding tiles to check for matches.
    
    ' First we check to make sure an autotile situation is actually there.
    ' Then we calculate exactly which situation has arisen.
    ' The situations are "inner", "outer", "horizontal", "vertical" and "fill".
    
    ' Exit out if we don't have an auatotile
    If Map.Tile(X, Y).Autotile(layerNum) = 0 Then Exit Sub
    
    ' Okay, we have autotiling but which one?
    Select Case Map.Tile(X, Y).Autotile(layerNum)
    
        ' Normal or animated - same difference
        Case ATNormal, ATAnim
            ' North West Quarter
            CalculateNW_Normal layerNum, X, Y
            
            ' North East Quarter
            CalculateNE_Normal layerNum, X, Y
            
            ' South West Quarter
            CalculateSW_Normal layerNum, X, Y
            
            ' South East Quarter
            CalculateSE_Normal layerNum, X, Y
            
        ' Cliff
        Case ATCliff
            ' North West Quarter
            CalculateNW_Cliff layerNum, X, Y
            
            ' North East Quarter
            CalculateNE_Cliff layerNum, X, Y
            
            ' South West Quarter
            CalculateSW_Cliff layerNum, X, Y
            
            ' South East Quarter
            CalculateSE_Cliff layerNum, X, Y
            
        ' Waterfalls
        Case ATWaterfall
            ' North West Quarter
            CalculateNW_Waterfall layerNum, X, Y
            
            ' North East Quarter
            CalculateNE_Waterfall layerNum, X, Y
            
            ' South West Quarter
            CalculateSW_Waterfall layerNum, X, Y
            
            ' South East Quarter
            CalculateSE_Waterfall layerNum, X, Y
        
        ' Anything else
        Case Else
            ' Don't need to render anything... it's fake or not an autotile
    End Select
End Sub

' Normal autotiling
Public Sub CalculateNW_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North West
    If checkTileMatch(layerNum, X, Y, X - 1, Y - 1) Then tmpTile(1) = True
    
    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(2) = True
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = ATInner
    ' Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = ATHorizontal
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = ATVertical
    ' Outer
    If Not tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = ATOuter
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = ATFill
    
    ' Actually place the subtile
    Select Case situation
        Case ATInner
            placeAutotile layerNum, X, Y, 1, "e"
        Case ATOuter
            placeAutotile layerNum, X, Y, 1, "a"
        Case ATHorizontal
            placeAutotile layerNum, X, Y, 1, "i"
        Case ATVertical
            placeAutotile layerNum, X, Y, 1, "m"
        Case ATFill
            placeAutotile layerNum, X, Y, 1, "q"
    End Select
End Sub

Public Sub CalculateNE_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(1) = True
    
    ' North East
    If checkTileMatch(layerNum, X, Y, X + 1, Y - 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = ATInner
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = ATHorizontal
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = ATVertical
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = ATOuter
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = ATFill
    
    ' Actually place the subtile
    Select Case situation
        Case ATInner
            placeAutotile layerNum, X, Y, 2, "j"
        Case ATOuter
            placeAutotile layerNum, X, Y, 2, "b"
        Case ATHorizontal
            placeAutotile layerNum, X, Y, 2, "f"
        Case ATVertical
            placeAutotile layerNum, X, Y, 2, "r"
        Case ATFill
            placeAutotile layerNum, X, Y, 2, "n"
    End Select
End Sub

Public Sub CalculateSW_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(1) = True
    
    ' South West
    If checkTileMatch(layerNum, X, Y, X - 1, Y + 1) Then tmpTile(2) = True
    
    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = ATInner
    ' Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = ATHorizontal
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = ATVertical
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = ATOuter
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = ATFill
    
    ' Actually place the subtile
    Select Case situation
        Case ATInner
            placeAutotile layerNum, X, Y, 3, "o"
        Case ATOuter
            placeAutotile layerNum, X, Y, 3, "c"
        Case ATHorizontal
            placeAutotile layerNum, X, Y, 3, "s"
        Case ATVertical
            placeAutotile layerNum, X, Y, 3, "g"
        Case ATFill
            placeAutotile layerNum, X, Y, 3, "k"
    End Select
End Sub

Public Sub CalculateSE_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(1) = True
    
    ' South East
    If checkTileMatch(layerNum, X, Y, X + 1, Y + 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = ATInner
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = ATHorizontal
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = ATVertical
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = ATOuter
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = ATFill
    
    ' Actually place the subtile
    Select Case situation
        Case ATInner
            placeAutotile layerNum, X, Y, 4, "t"
        Case ATOuter
            placeAutotile layerNum, X, Y, 4, "d"
        Case ATHorizontal
            placeAutotile layerNum, X, Y, 4, "p"
        Case ATVertical
            placeAutotile layerNum, X, Y, 4, "l"
        Case ATFill
            placeAutotile layerNum, X, Y, 4, "h"
    End Select
End Sub

' Waterfall autotiling
Public Sub CalculateNW_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 1, "i"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 1, "e"
    End If
End Sub

Public Sub CalculateNE_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 2, "f"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 2, "j"
    End If
End Sub

Public Sub CalculateSW_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 3, "k"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 3, "g"
    End If
End Sub

Public Sub CalculateSE_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile As Boolean
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 4, "h"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 4, "l"
    End If
End Sub

' Cliff autotiling
Public Sub CalculateNW_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North West
    If checkTileMatch(layerNum, X, Y, X - 1, Y - 1) Then tmpTile(1) = True
    
    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(2) = True
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = ATHorizontal
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = ATVertical
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = ATFill
    ' Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = ATInner
    
    ' Actually place the subtile
    Select Case situation
        Case ATInner
            placeAutotile layerNum, X, Y, 1, "e"
        Case ATHorizontal
            placeAutotile layerNum, X, Y, 1, "i"
        Case ATVertical
            placeAutotile layerNum, X, Y, 1, "m"
        Case ATFill
            placeAutotile layerNum, X, Y, 1, "q"
    End Select
End Sub

Public Sub CalculateNE_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(1) = True
    
    ' North East
    If checkTileMatch(layerNum, X, Y, X + 1, Y - 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = ATHorizontal
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = ATVertical
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = ATFill
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = ATInner
    
    ' Actually place the subtile
    Select Case situation
        Case ATInner
            placeAutotile layerNum, X, Y, 2, "j"
        Case ATHorizontal
            placeAutotile layerNum, X, Y, 2, "f"
        Case ATVertical
            placeAutotile layerNum, X, Y, 2, "r"
        Case ATFill
            placeAutotile layerNum, X, Y, 2, "n"
    End Select
End Sub

Public Sub CalculateSW_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(1) = True
    
    ' South West
    If checkTileMatch(layerNum, X, Y, X - 1, Y + 1) Then tmpTile(2) = True
    
    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = ATHorizontal
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = ATVertical
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = ATFill
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = ATInner
    
    ' Actually place the subtile
    Select Case situation
        Case ATInner
            placeAutotile layerNum, X, Y, 3, "o"
        Case ATHorizontal
            placeAutotile layerNum, X, Y, 3, "s"
        Case ATVertical
            placeAutotile layerNum, X, Y, 3, "g"
        Case ATFill
            placeAutotile layerNum, X, Y, 3, "k"
    End Select
End Sub

Public Sub CalculateSE_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
Dim tmpTile(1 To 3) As Boolean
Dim situation As Byte

    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(1) = True
    
    ' South East
    If checkTileMatch(layerNum, X, Y, X + 1, Y + 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation -  Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = ATHorizontal
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = ATVertical
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = ATFill
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = ATInner
    
    ' Actually place the subtile
    Select Case situation
        Case ATInner
            placeAutotile layerNum, X, Y, 4, "t"
        Case ATHorizontal
            placeAutotile layerNum, X, Y, 4, "p"
        Case ATVertical
            placeAutotile layerNum, X, Y, 4, "l"
        Case ATFill
            placeAutotile layerNum, X, Y, 4, "h"
    End Select
End Sub

Public Function checkTileMatch(ByVal layerNum As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
    ' we'll exit out early if true
    checkTileMatch = True
    
    ' if it's off the map then set it as autotile and exit out early
    If x2 < 0 Or x2 > Map.MaxX Or y2 < 0 Or y2 > Map.MaxY Then
        checkTileMatch = True
        Exit Function
    End If
    
    ' fakes ALWAYS return true
    If Map.Tile(x2, y2).Autotile(layerNum) = ATFake Then
        checkTileMatch = True
        Exit Function
    End If
    
    ' check neighbour is an autotile
    If Map.Tile(x2, y2).Autotile(layerNum) = 0 Then
        checkTileMatch = False
        Exit Function
    End If
    
    ' check we're a matching
    If Map.Tile(x1, y1).Layer(layerNum).Tileset <> Map.Tile(x2, y2).Layer(layerNum).Tileset Then
        checkTileMatch = False
        Exit Function
    End If
    
    ' check tiles match
    If Map.Tile(x1, y1).Layer(layerNum).X <> Map.Tile(x2, y2).Layer(layerNum).X Then
        checkTileMatch = False
        Exit Function
    End If
        
    If Map.Tile(x1, y1).Layer(layerNum).Y <> Map.Tile(x2, y2).Layer(layerNum).Y Then
        checkTileMatch = False
        Exit Function
    End If
End Function

Public Sub DrawAutoTile(ByVal layerNum As Long, ByVal destX As Long, ByVal destY As Long, ByVal quarterNum As Long, ByVal X As Long, ByVal Y As Long)
Dim yOffset As Long, xOffset As Long

    ' calculate the offset
    Select Case Map.Tile(X, Y).Autotile(layerNum)
        Case ATWaterfall
            yOffset = (waterfallFrame - 1) * 32
        Case ATAnim
            xOffset = autoTileFrame * 64
        Case ATCliff
            yOffset = -32
    End Select
    
    ' Draw the quarter
    'EngineRenderRectangle Tex_Tileset(Map.Tile(x, y).Layer(layerNum).Tileset), destX, destY, Autotile(x, y).Layer(layerNum).srcX(quarterNum) + xOffset, Autotile(x, y).Layer(layerNum).srcY(quarterNum) + yOffset, 16, 16, 16, 16, 16, 16
    RenderTexture Tex_Tileset(Map.Tile(X, Y).Layer(layerNum).Tileset), destX, destY, Autotile(X, Y).Layer(layerNum).srcX(quarterNum) + xOffset, Autotile(X, Y).Layer(layerNum).srcY(quarterNum) + yOffset, 16, 16, 16, 16, -1
End Sub
' player Projectiles
Public Sub DrawProjectile(ByVal Index As Long, ByVal PlayerProjectile As Long)
Dim X As Long, Y As Long, PicNum As Long, i As Long
Dim rec As DxVBLib.RECT

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    ' check for subscript error
    If Index < 1 Or PlayerProjectile < 1 Or PlayerProjectile > MAX_PLAYER_PROJECTILES Then Exit Sub
    
    ' check to see if it's time to move the Projectile
    If timeGetTime > Player(Index).ProjecTile(PlayerProjectile).TravelTime Then
        With Player(Index).ProjecTile(PlayerProjectile)
            ' set next travel time and the current position and then set the actual direction based on RMXP arrow tiles.
            Select Case .direction
                ' down
                Case 0
                    .Y = .Y + 1
                    ' check if they reached maxrange
                    If .Y = (GetPlayerY(Index) + .Range) + 1 Then ClearProjectile Index, PlayerProjectile: Exit Sub
                ' up
                Case 1
                    .Y = .Y - 1
                    ' check if they reached maxrange
                    If .Y = (GetPlayerY(Index) - .Range) - 1 Then ClearProjectile Index, PlayerProjectile: Exit Sub
                ' right
                Case 2
                    .X = .X + 1
                    ' check if they reached max range
                    If .X = (GetPlayerX(Index) + .Range) + 1 Then ClearProjectile Index, PlayerProjectile: Exit Sub
                ' left
                Case 3
                    .X = .X - 1
                    ' check if they reached maxrange
                    If .X = (GetPlayerX(Index) - .Range) - 1 Then ClearProjectile Index, PlayerProjectile: Exit Sub
            End Select
            .TravelTime = timeGetTime + .speed
        End With
    End If
    
    ' set the x, y & pic values for future reference
    X = Player(Index).ProjecTile(PlayerProjectile).X
    Y = Player(Index).ProjecTile(PlayerProjectile).Y
    PicNum = Player(Index).ProjecTile(PlayerProjectile).Pic
    
    ' check if left map
    If X > Map.MaxX Or Y > Map.MaxY Or X < 0 Or Y < 0 Then
        ClearProjectile Index, PlayerProjectile
        Exit Sub
    End If
    
    ' check if we hit a block
    If Map.Tile(X, Y).Type = TileBlocked Then
        ClearProjectile Index, PlayerProjectile
        Exit Sub
    End If
    
     If Map.Tile(X, Y).Type = TileWater Then
        ClearProjectile Index, PlayerProjectile
        Exit Sub
    End If
    
    ' check for player hit
    For i = 1 To Player_HighIndex
        If X = GetPlayerX(i) And Y = GetPlayerY(i) Then
            ' they're hit, remove it
            If Not X = Player(MyIndex).X Or Not Y = GetPlayerY(MyIndex) Then
                ClearProjectile Index, PlayerProjectile
                Exit Sub
            End If
        End If
    Next
    
    ' check for npc hit
    For i = 1 To MAX_MAP_NPCS
        If X = MapNpc(i).X And Y = MapNpc(i).Y Then
            ' they're hit, remove it
            ClearProjectile Index, PlayerProjectile
            Exit Sub
        End If
    Next

    
    ' get positioning in the texture
    With rec
        .Top = 0
        .Bottom = SIZE_Y
        .Left = Player(Index).ProjecTile(PlayerProjectile).direction * SIZE_X
        .Right = .Left + SIZE_X
    End With

    ' blt the projectile
    RenderTexture Tex_Projectile(PicNum), ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    'Call Engine_BltFast(ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), DDS_Projectile(PicNum), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltProjectile", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

