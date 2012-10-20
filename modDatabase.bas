Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpFileName As String) As Long

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
Dim filename As String
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\data files\logs\errors.txt"
    Open filename For Append As #1
        Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
        Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
        Print #1, ""
    Close #1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleError", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ChkDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function FileExist(ByVal filename As String, Optional RAW As Boolean = False) As Boolean
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Not RAW Then
        If LenB(Dir(App.Path & "\" & filename)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(Dir(filename)) > 0 Then
            FileExist = True
        End If
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "FileExist", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    Call WritePrivateProfileString$(Header, Var, Value, File)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PutVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveOptions()
Dim filename As String

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\Data Files\config.ini"
    
    Call PutVar(filename, "Options", "Game_Name", Trim$(options.Game_Name))
    Call PutVar(filename, "Options", "Username", Trim$(options.Username))
    Call PutVar(filename, "Options", "Password", Trim$(options.Password))
    Call PutVar(filename, "Options", "SavePass", str(options.SavePass))
    Call PutVar(filename, "Options", "IP", options.IP)
    Call PutVar(filename, "Options", "Port", str(options.Port))
    Call PutVar(filename, "Options", "MenuMusic", Trim$(options.MenuMusic))
    Call PutVar(filename, "Options", "Music", str(options.Music))
    Call PutVar(filename, "Options", "Sound", str(options.sound))
    Call PutVar(filename, "Options", "Debug", str(options.Debug))
    Call PutVar(filename, "Options", "hq", str(options.HQ))
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadOptions()
Dim filename As String

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\Data Files\config.ini"
    
    If Not FileExist(filename, True) Then
        options.Game_Name = "Eclipse Origins"
        options.Password = vbNullString
        options.SavePass = 0
        options.Username = vbNullString
        options.IP = "127.0.0.1"
        options.Port = 7001
        options.MenuMusic = vbNullString
        options.Music = 1
        options.sound = 1
        options.Debug = 0
        options.HQ = 0
        SaveOptions
    Else
        options.Game_Name = GetVar(filename, "Options", "Game_Name")
        options.Username = GetVar(filename, "Options", "Username")
        options.Password = GetVar(filename, "Options", "Password")
        options.SavePass = Val(GetVar(filename, "Options", "SavePass"))
        options.IP = GetVar(filename, "Options", "IP")
        options.Port = Val(GetVar(filename, "Options", "Port"))
        options.MenuMusic = GetVar(filename, "Options", "MenuMusic")
        options.Music = GetVar(filename, "Options", "Music")
        options.sound = GetVar(filename, "Options", "Sound")
        options.Debug = GetVar(filename, "Options", "Debug")
        options.HQ = GetVar(filename, "Options", "hq")
    End If
    
    ' show in GUI
    If options.Music = 0 Then
        frmMain.optMOff.Value = True
    Else
        frmMain.optMOn.Value = True
    End If
    
    If options.sound = 0 Then
        frmMain.optSOff.Value = True
    Else
        frmMain.optSOn.Value = True
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveMap(ByVal MapNum As Long)
Dim filename As String
Dim f As Long
Dim X As Long
Dim Y As Long, i As Long, Z As Long, w As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT

    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Map.Name
    Put #f, , Map.Music
    Put #f, , Map.BGS
    Put #f, , Map.Revision
    Put #f, , Map.Moral
    Put #f, , Map.Up
    Put #f, , Map.Down
    Put #f, , Map.Left
    Put #f, , Map.Right
    Put #f, , Map.BootMap
    Put #f, , Map.BootX
    Put #f, , Map.BootY
    
    Put #f, , Map.Weather
    Put #f, , Map.WeatherIntensity
    
    Put #f, , Map.Fog
    Put #f, , Map.FogSpeed
    Put #f, , Map.FogOpacity
    
    Put #f, , Map.Red
    Put #f, , Map.Green
    Put #f, , Map.Blue
    Put #f, , Map.Alpha
    
    Put #f, , Map.MaxX
    Put #f, , Map.MaxY

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            Put #f, , Map.Tile(X, Y)
        Next

        DoEvents
    Next

    For X = 1 To MAX_MAP_NPCS
        Put #f, , Map.NPC(X)
        Put #f, , Map.NpcSpawnType(X)
    Next
    

    Close #f
    
    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadMap(ByVal MapNum As Long)
Dim filename As String
Dim f As Long
Dim X As Long
Dim Y As Long, i As Long, Z As Long, w As Long, p As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT
    ClearMap
    f = FreeFile
    Open filename For Binary As #f
    Get #f, , Map.Name
    Get #f, , Map.Music
    Get #f, , Map.BGS
    Get #f, , Map.Revision
    Get #f, , Map.Moral
    Get #f, , Map.Up
    Get #f, , Map.Down
    Get #f, , Map.Left
    Get #f, , Map.Right
    Get #f, , Map.BootMap
    Get #f, , Map.BootX
    Get #f, , Map.BootY
    
    Get #f, , Map.Weather
    Get #f, , Map.WeatherIntensity
        
    Get #f, , Map.Fog
    Get #f, , Map.FogSpeed
    Get #f, , Map.FogOpacity
        
    Get #f, , Map.Red
    Get #f, , Map.Green
    Get #f, , Map.Blue
    Get #f, , Map.Alpha
    
    Get #f, , Map.MaxX
    Get #f, , Map.MaxY
    ' have to set the tile()
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            Get #f, , Map.Tile(X, Y)
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Get #f, , Map.NPC(X)
        Get #f, , Map.NpcSpawnType(X)
    Next

    Close #f
    ClearTempTile
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckTilesets()
Dim i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    NumTileSets = 1
    
    ReDim Tex_Tileset(1)

    While FileExist(GFX_PATH & "tilesets\" & i & GFX_EXT)
        ReDim Preserve Tex_Tileset(NumTileSets)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Tileset(NumTileSets).filepath = App.Path & GFX_PATH & "tilesets\" & i & GFX_EXT
        Tex_Tileset(NumTileSets).Texture = NumTextures
        NumTileSets = NumTileSets + 1
        i = i + 1
    Wend
    
    NumTileSets = NumTileSets - 1
    
    If NumTileSets = 0 Then Exit Sub
     ' Error handler
     Exit Sub
errorhandler:
    HandleError "CheckTilesets", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
      Err.Clear
      Exit Sub
     End Sub

Public Sub CheckCharacters()
Dim i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    NumCharacters = 1
    
    ReDim Tex_Character(1)
    

    While FileExist(GFX_PATH & "characters\" & i & GFX_EXT)
        ReDim Preserve Tex_Character(NumCharacters)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Character(NumCharacters).filepath = App.Path & GFX_PATH & "characters\" & i & GFX_EXT
        Tex_Character(NumCharacters).Texture = NumTextures
        NumCharacters = NumCharacters + 1
        i = i + 1
    Wend
    
    NumCharacters = NumCharacters - 1
    
    If NumCharacters = 0 Then Exit Sub
    
    For i = 1 To NumCharacters
        LoadTexture Tex_Character(i)
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckCharacters", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckPaperdolls()
Dim i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    NumPaperdolls = 1
    
    ReDim Tex_Paperdoll(1)

    While FileExist(GFX_PATH & "paperdolls\" & i & GFX_EXT)
        ReDim Preserve Tex_Paperdoll(NumPaperdolls)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Paperdoll(NumPaperdolls).filepath = App.Path & GFX_PATH & "paperdolls\" & i & GFX_EXT
        Tex_Paperdoll(NumPaperdolls).Texture = NumTextures
        NumPaperdolls = NumPaperdolls + 1
        i = i + 1
    Wend
    
    NumPaperdolls = NumPaperdolls - 1
    
    If NumPaperdolls = 0 Then Exit Sub
    
    For i = 1 To NumPaperdolls
        LoadTexture Tex_Paperdoll(i)
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckPaperdolls", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAnimations()
Dim i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    NumAnimations = 1
    
    ReDim Tex_Animation(1)
    ReDim AnimationTimer(1)

    While FileExist(GFX_PATH & "animations\" & i & GFX_EXT)
        ReDim Preserve Tex_Animation(NumAnimations)
        ReDim Preserve AnimationTimer(NumAnimations)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Animation(NumAnimations).Texture = NumTextures
        Tex_Animation(NumAnimations).filepath = App.Path & GFX_PATH & "animations\" & i & GFX_EXT
        NumAnimations = NumAnimations + 1
        i = i + 1
    Wend
    
    NumAnimations = NumAnimations - 1
    
    If NumAnimations = 0 Then Exit Sub

    For i = 1 To NumAnimations
        LoadTexture Tex_Animation(i)
    Next
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckItems()
Dim i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    numitems = 1
    
    ReDim Tex_Item(1)

    While FileExist(GFX_PATH & "items\" & i & GFX_EXT)
        ReDim Preserve Tex_Item(numitems)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Item(numitems).filepath = App.Path & GFX_PATH & "items\" & i & GFX_EXT
        Tex_Item(numitems).Texture = NumTextures
        numitems = numitems + 1
        i = i + 1
    Wend
    
    numitems = numitems - 1
    
    If numitems = 0 Then Exit Sub
    
    For i = 1 To numitems
        LoadTexture Tex_Item(i)
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckResources()
Dim i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    NumResources = 1
    
    ReDim Tex_Resource(1)

    While FileExist(GFX_PATH & "resources\" & i & GFX_EXT)
        ReDim Preserve Tex_Resource(NumResources)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Resource(NumResources).filepath = App.Path & GFX_PATH & "resources\" & i & GFX_EXT
        Tex_Resource(NumResources).Texture = NumTextures
        NumResources = NumResources + 1
        i = i + 1
    Wend
    
    NumResources = NumResources - 1
    
    If NumResources = 0 Then Exit Sub
    
    For i = 1 To NumResources
        LoadTexture Tex_Resource(i)
    Next
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckSpellIcons()
Dim i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    NumSpellIcons = 1
    
    ReDim Tex_SpellIcon(1)

    While FileExist(GFX_PATH & "spellicons\" & i & GFX_EXT)
        ReDim Preserve Tex_SpellIcon(NumSpellIcons)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_SpellIcon(NumSpellIcons).filepath = App.Path & GFX_PATH & "spellicons\" & i & GFX_EXT
        Tex_SpellIcon(NumSpellIcons).Texture = NumTextures
        NumSpellIcons = NumSpellIcons + 1
        i = i + 1
    Wend
    
    NumSpellIcons = NumSpellIcons - 1
    
    If NumSpellIcons = 0 Then Exit Sub
    
    For i = 1 To NumSpellIcons
        LoadTexture Tex_SpellIcon(i)
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckSpellIcons", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckFaces()
Dim i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    NumFaces = 1
    
    ReDim Tex_Face(1)

    While FileExist(GFX_PATH & "Faces\" & i & GFX_EXT)
        ReDim Preserve Tex_Face(NumFaces)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Face(NumFaces).filepath = App.Path & GFX_PATH & "faces\" & i & GFX_EXT
        Tex_Face(NumFaces).Texture = NumTextures
        NumFaces = NumFaces + 1
        i = i + 1
    Wend
    
    NumFaces = NumFaces - 1
    
    If NumFaces = 0 Then Exit Sub
    
    For i = 1 To NumFaces
        LoadTexture Tex_Face(i)
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckFaces", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckFogs()
Dim i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    NumFogs = 1
    
    ReDim Tex_Fog(1)
    While FileExist(GFX_PATH & "fogs\" & i & GFX_EXT)
        ReDim Preserve Tex_Fog(NumFogs)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Fog(NumFogs).filepath = App.Path & GFX_PATH & "fogs\" & i & GFX_EXT
        Tex_Fog(NumFogs).Texture = NumTextures
        NumFogs = NumFogs + 1
        i = i + 1
    Wend
    
    NumFogs = NumFogs - 1
    
    If NumFogs = 0 Then Exit Sub
    
    For i = 1 To NumFogs
        LoadTexture Tex_Fog(i)
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckFogs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearPlayer(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).Name = vbNullString
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearPlayer", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearItem(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
    Item(Index).Desc = vbNullString
    Item(Index).sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearItems()
Dim i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimInstance(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(AnimInstance(Index)), LenB(AnimInstance(Index)))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimInstance", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimation(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Animation(Index).Name = vbNullString
    Animation(Index).sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimation", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimations()
Dim i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearNPC(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(NPC(Index)), LenB(NPC(Index)))
    NPC(Index).Name = vbNullString
    NPC(Index).sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearNPC", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearNpcs()
Dim i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_NPCS
        Call ClearNPC(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearSpell(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).Name = vbNullString
    Spell(Index).Desc = vbNullString
    Spell(Index).sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearSpell", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearSpells()
Dim i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearSpells", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearShop(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).Name = vbNullString
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearShop", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearShops()
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearShops", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearResource(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).Name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    Resource(Index).sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearResource", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearResources()
Dim i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapItem(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapItem(Index)), LenB(MapItem(Index)))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMap()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Map), LenB(Map))
    Map.Name = vbNullString
    Map.MaxX = MAX_MAPX
    Map.MaxY = MAX_MAPY
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)
    initAutotiles
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapItems()
Dim i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapNpc(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapNpc(Index)), LenB(MapNpc(Index)))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapNpc", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapNpcs()
Dim i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **********************
' ** Player functions **
' **********************
Function GetPlayerName(ByVal Index As Long) As String
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).Name)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Name = Name
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerClass = Player(Index).Class
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Class = ClassNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(Index).Sprite
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerSprite", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Sprite = Sprite
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerSprite", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(Index).Level
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Level = Level
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerExp(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerExp = Player(Index).EXP
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
Function GetPlayerCrafting(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerCrafting = Player(Index).Crafting
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerCrafting", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerCrafting(ByVal Index As Long, ByVal Crafting As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Crafting = Crafting
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerCrafting", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerCraftingExp(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerCraftingExp = Player(Index).CraftingExp
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerCraftingExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
Sub SetPlayerCraftingExp(ByVal Index As Long, ByVal CraftingExp As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).CraftingExp = CraftingExp
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerSwordsExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Function GetPlayerWoodcutting(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerWoodcutting = Player(Index).Woodcutting
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerWoodcutting", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
Function GetPlayerFishing(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerFishing = Player(Index).Fishing
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerFishing", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerFishing(ByVal Index As Long, ByVal Fishing As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Fishing = Fishing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerFishing", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerFishingExp(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerFishingExp = Player(Index).FishingExp
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerFishingExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerFishingExp(ByVal Index As Long, ByVal FishingExp As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).FishingExp = FishingExp
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerFishingExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SetPlayerWoodcutting(ByVal Index As Long, ByVal Woodcutting As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Woodcutting = Woodcutting
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerWoodcutting", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerWoodcuttingExp(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerWoodcuttingExp = Player(Index).WoodcuttingExp
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerWoodcuttingExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerWoodcuttingExp(ByVal Index As Long, ByVal WoodcuttingExp As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).WoodcuttingExp = WoodcuttingExp
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerWoodcuttingExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Function GetPlayerMining(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerMining = Player(Index).Mining
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerMining", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerMining(ByVal Index As Long, ByVal Mining As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Mining = Mining
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerMining", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerMiningExp(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerMiningExp = Player(Index).MiningExp
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerMiningExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerMiningExp(ByVal Index As Long, ByVal MiningExp As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).MiningExp = MiningExp
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerMiningExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Function GetPlayerNextLevel(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerNextLevel = (50 / 3) * ((GetPlayerLevel(Index) + 1) ^ 3 - (6 * (GetPlayerLevel(Index) + 1) ^ 2) + 17 * (GetPlayerLevel(Index) + 1) - 12)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerNextLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
Function GetPlayerLightArmor(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerLightArmor = Player(Index).LightArmor
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerLightArmor", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerLightArmor(ByVal Index As Long, ByVal LightArmor As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).LightArmor = LightArmor
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerLightArmor", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerLightArmorExp(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerLightArmorExp = Player(Index).LightArmorExp
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerLightArmorExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerLightArmorExp(ByVal Index As Long, ByVal LightArmorExp As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).LightArmorExp = LightArmorExp
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerLightArmorExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Function GetPlayerHeavyarmor(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerHeavyarmor = Player(Index).Heavyarmor
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerHeavyarmor", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerHeavyarmor(ByVal Index As Long, ByVal Heavyarmor As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Heavyarmor = Heavyarmor
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerHeavyarmor", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerHeavyarmorExp(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerHeavyarmorExp = Player(Index).HeavyArmorExp
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerHeavyarmorExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
   End Function
Sub SetPlayerHeavyArmorExp(ByVal Index As Long, ByVal HeavyArmorExp As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).HeavyArmorExp = HeavyArmorExp
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerHeavyArmorExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Function GetPlayerLycanthropy(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerLycanthropy = Player(Index).Lycanthropy
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerLycanthropy", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerLycanthropy(ByVal Index As Long, ByVal Lycanthropy As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Lycanthropy = Lycanthropy
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerLycanthropy", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerLycanthropyExp(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerLycanthropyExp = Player(Index).LycanthropyExp
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerLycanthropyExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerLycanthropyExp(ByVal Index As Long, ByVal LycanthropyExp As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).LycanthropyExp = LycanthropyExp
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerLycanthropyExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SetPlayerExp(ByVal Index As Long, ByVal EXP As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).EXP = EXP
    If GetPlayerLevel(Index) = MAX_LEVELS And Player(Index).EXP > GetPlayerNextLevel(Index) Then
        Player(Index).EXP = GetPlayerNextLevel(Index)
        Exit Sub
    End If
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
'prof smithy
Function GetPlayerSmith(ByVal Index As Long) As Long
' If debug mode, handle error then exit out
If options.Debug = 1 Then On Error GoTo errorhandler
If Index > MAX_PLAYERS Then Exit Function
GetPlayerSmith = Player(Index).Smith

' Error handler
Exit Function
errorhandler:
HandleError "GetPlayerSmith", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Function
End Function
Sub SetPlayerSmith(ByVal Index As Long, ByVal Smith As Long)
' If debug mode, handle error then exit out
If options.Debug = 1 Then On Error GoTo errorhandler
If Index > MAX_PLAYERS Then Exit Sub
Player(Index).Smith = Smith

' Error handler
Exit Sub
errorhandler:
HandleError "SetPlayerSmith", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
Function GetPlayerSmithExp(ByVal Index As Long) As Long
' If debug mode, handle error then exit out
If options.Debug = 1 Then On Error GoTo errorhandler
If Index > MAX_PLAYERS Then Exit Function
GetPlayerSmithExp = Player(Index).SmithExp

' Error handler
Exit Function
errorhandler:
HandleError "GetPlayerSmithExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Function
End Function
Sub SetPlayerSmithExp(ByVal Index As Long, ByVal SmithExp As Long)
' If debug mode, handle error then exit out
If options.Debug = 1 Then On Error GoTo errorhandler
If Index > MAX_PLAYERS Then Exit Sub
Player(Index).SmithExp = SmithExp

' Error handler
Exit Sub
errorhandler:
HandleError "SetPlayerSmithExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
'prof alachemy
Function GetPlayerAlchemy(ByVal Index As Long) As Long
' If debug mode, handle error then exit out
If options.Debug = 1 Then On Error GoTo errorhandler
If Index > MAX_PLAYERS Then Exit Function
GetPlayerAlchemy = Player(Index).Alchemy

' Error handler
Exit Function
errorhandler:
HandleError "GetPlayerAlchemy", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Function
End Function
Sub SetPlayerAlchemy(ByVal Index As Long, ByVal Alchemy As Long)
' If debug mode, handle error then exit out
If options.Debug = 1 Then On Error GoTo errorhandler
If Index > MAX_PLAYERS Then Exit Sub
Player(Index).Alchemy = Alchemy

' Error handler
Exit Sub
errorhandler:
HandleError "SetPlayerAlchemy", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
Function GetPlayerAlchemyExp(ByVal Index As Long) As Long
' If debug mode, handle error then exit out
If options.Debug = 1 Then On Error GoTo errorhandler
If Index > MAX_PLAYERS Then Exit Function
GetPlayerAlchemyExp = Player(Index).AlchemyExp

' Error handler
Exit Function
errorhandler:
HandleError "GetPlayerAlchemyExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Function
End Function
Sub SetPlayerAlchemyExp(ByVal Index As Long, ByVal AlchemyExp As Long)
' If debug mode, handle error then exit out
If options.Debug = 1 Then On Error GoTo errorhandler
If Index > MAX_PLAYERS Then Exit Sub
Player(Index).AlchemyExp = AlchemyExp

' Error handler
Exit Sub
errorhandler:
HandleError "SetPlayerAlchemyExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
'prof enchantor
Function GetPlayerEnchants(ByVal Index As Long) As Long
' If debug mode, handle error then exit out
If options.Debug = 1 Then On Error GoTo errorhandler
If Index > MAX_PLAYERS Then Exit Function
GetPlayerEnchants = Player(Index).Enchants

' Error handler
Exit Function
errorhandler:
HandleError "GetPlayerEnchants", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Function
End Function
Sub SetPlayerEnchants(ByVal Index As Long, ByVal Enchants As Long)
' If debug mode, handle error then exit out
If options.Debug = 1 Then On Error GoTo errorhandler
If Index > MAX_PLAYERS Then Exit Sub
Player(Index).Enchants = Enchants

' Error handler
Exit Sub
errorhandler:
HandleError "SetPlayerEnchants", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
Function GetPlayerEnchantsExp(ByVal Index As Long) As Long
' If debug mode, handle error then exit out
If options.Debug = 1 Then On Error GoTo errorhandler
If Index > MAX_PLAYERS Then Exit Function
GetPlayerEnchantsExp = Player(Index).EnchantsExp

' Error handler
Exit Function
errorhandler:
HandleError "GetPlayerEnchantsExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Function
End Function
Sub SetPlayerEnchantsExp(ByVal Index As Long, ByVal EnchantsExp As Long)
' If debug mode, handle error then exit out
If options.Debug = 1 Then On Error GoTo errorhandler
If Index > MAX_PLAYERS Then Exit Sub
Player(Index).EnchantsExp = EnchantsExp

' Error handler
Exit Sub
errorhandler:
HandleError "SetPlayerEnchantsExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
Function GetPlayerConviction(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerConviction = Player(Index).Conviction
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerConviction", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerConviction(ByVal Index As Long, ByVal Conviction As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Conviction = Conviction
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerConviction", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerConvictionExp(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerConvictionExp = Player(Index).ConvictionExp
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerConvictionExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerConvictionExp(ByVal Index As Long, ByVal ConvictionExp As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).ConvictionExp = ConvictionExp
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerConvictionExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Function GetPlayerMagic(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerMagic = Player(Index).Magic
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerMagic", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerMagic(ByVal Index As Long, ByVal Magic As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Magic = Magic
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerMagic", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerMagicExp(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerMagicExp = Player(Index).MagicExp
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerMagicExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerMagicExp(ByVal Index As Long, ByVal MagicExp As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).MagicExp = MagicExp
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerMagicExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Function GetPlayerRange(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerRange = Player(Index).Range
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerRange", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerRange(ByVal Index As Long, ByVal Range As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Range = Range
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerRange", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerRangeExp(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerRangeExp = Player(Index).RangeExp
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerRangeExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerRangeExp(ByVal Index As Long, ByVal RangeExp As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).RangeExp = RangeExp
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerRangeExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Function GetPlayerDaggers(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerDaggers = Player(Index).Daggers
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerDaggers", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerDaggers(ByVal Index As Long, ByVal Daggers As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Daggers = Daggers
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerDaggers", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerDaggersExp(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerDaggersExp = Player(Index).DaggersExp
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerDaggersExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerDaggersExp(ByVal Index As Long, ByVal DaggersExp As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).DaggersExp = DaggersExp
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerDaggersExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Function GetPlayerAxes(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerAxes = Player(Index).Axes
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerAxes", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerAxes(ByVal Index As Long, ByVal Axes As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Axes = Axes
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerAxes", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerAxesExp(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerAxesExp = Player(Index).AxesExp
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerAxesExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerAxesExp(ByVal Index As Long, ByVal AxesExp As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).AxesExp = AxesExp
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerAxesExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Function GetPlayerAccess(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(Index).Access
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerAccess", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Access = Access
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerAccess", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(Index).PK
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).PK = PK
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(Index).Vital(Vital)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Vital(Vital) = Value

    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    
    GetPlayerMaxVital = Player(Index).MaxVital(Vital)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerMaxVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetPlayerStat(ByVal Index As Long, Stat As Stats) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerStat = Player(Index).Stat(Stat)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerStat(ByVal Index As Long, Stat As Stats, ByVal Value As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    If Value <= 0 Then Value = 1
    If Value > MAX_BYTE Then Value = MAX_BYTE
    Player(Index).Stat(Stat) = Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(Index).POINTS
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).POINTS = POINTS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Or Index <= 0 Then Exit Function
    GetPlayerMap = Player(Index).Map
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Map = MapNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(Index).X
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerX", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).X = X
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerX", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(Index).Y
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerY", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Y = Y
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerY", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(Index).Dir
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Dir = Dir
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal invslot As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function
    GetPlayerInvItemNum = PlayerInv(invslot).num
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal invslot As Long, ByVal itemnum As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invslot).num = itemnum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal invslot As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = PlayerInv(invslot).Value
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal invslot As Long, ByVal ItemValue As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invslot).Value = ItemValue
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerEquipment(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerEquipment = Player(Index).Equipment(EquipmentSlot)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerEquipment", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerEquipment(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Equipment(EquipmentSlot) = InvNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerEquipment", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Function GetPlayerSwords(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerSwords = Player(Index).Swords
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerSwords", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerSwords(ByVal Index As Long, ByVal Swords As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Swords = Swords
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerSwords", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerSwordsExp(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerSwordsExp = Player(Index).SwordsExp
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerSwordsExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerSwordsExp(ByVal Index As Long, ByVal SwordsExp As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).SwordsExp = SwordsExp
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerSwordsExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
' projectiles
Public Sub CheckProjectiles()
Dim i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    i = 1
    NumProjectiles = 1
    
    ReDim Tex_Projectile(1)
    
    While FileExist(GFX_PATH & "projectiles\" & i & GFX_EXT)
        ReDim Preserve Tex_Projectile(NumProjectiles)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Projectile(NumProjectiles).filepath = App.Path & GFX_PATH & "projectiles\" & i & GFX_EXT
        Tex_Projectile(NumProjectiles).Texture = NumTextures
        NumProjectiles = NumProjectiles + 1
        i = i + 1
    Wend
    
    NumProjectiles = NumProjectiles - 1
    
    If NumProjectiles = 0 Then Exit Sub
    
    For i = 1 To NumProjectiles
        LoadTexture Tex_Projectile(i)
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearProjectile(ByVal Index As Long, ByVal PlayerProjectile As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    With Player(Index).ProjecTile(PlayerProjectile)
        .direction = 0
        .Pic = 0
        .TravelTime = 0
        .X = 0
        .Y = 0
        .Range = 0
        .Damage = 0
        .speed = 0
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearProjectile", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
