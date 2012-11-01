Attribute VB_Name = "modGameLogic"
Option Explicit


Public Sub GameLoop()
Dim FrameTime As Long
Dim tick As Long
Dim TickFPS As Long
Dim FPS As Long
Dim i As Long
Dim WalkTimer As Long
Dim tmr25 As Long
Dim tmr100 As Long
Dim tmr10000 As Long
Dim tmr500 As Long
Dim Fadetmr As Long
Dim fogtmr As Long
Dim surfTmr As Long
Dim renderTmr As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    ' *** Start GameLoop ***
    Do While InGame
        tick = timeGetTime                            ' Set the inital tick
        ElapsedTime = tick - FrameTime                 ' Set the time difference for time-based movement
        FrameTime = tick                               ' Set the time second loop time to the first.

        ' * Check surface timers *
          ' Unload unused surfaces every half-of-surface-timer seconds.
           If surfTmr < tick Then
             For i = 1 To NumTextures ' Loop!
                If gTexture(i).timer > 0 Then ' Check if it's in use!
                  If gTexture(i).timer < tick Then ' Check if it's time to unload.
                    ' Unload!
                    Set gTexture(i).Texture = Nothing
                    'ZeroMemory ByVal VarPtr(gTexture(i)), LenB(gTexture(i))
                    gTexture(i).timer = 0
               AddText "Unloaded texture: " & i, White
            End If
        End If
       DoEvents
      Next
    surfTmr = tick + (SurfaceTimer * 0.5)
  End If
                            
                            
        ' Sprites
        If tmr10000 < tick Then


            
            ' check ping
            Call GetPing
            Call DrawPing
            tmr10000 = tick + 10000
        End If

        If tmr25 < tick Then
            InGame = IsConnected
            Call CheckKeys ' Check to make sure they aren't trying to auto do anything

            If GetForegroundWindow() = frmMain.hWnd Or GetForegroundWindow() = frmEditor_Events.hWnd Then
                Call CheckInputKeys ' Check which keys were pressed
            End If
            
            ' check if we need to end the CD icon
            If NumSpellIcons > 0 Then
                For i = 1 To MAX_PLAYER_SPELLS
                    If PlayerSpells(i) > 0 Then
                        If SpellCD(i) > 0 Then
                            If SpellCD(i) + (Spell(PlayerSpells(i)).CDTime * 1000) < tick Then
                                SpellCD(i) = 0
                                frmMain.picSpells.Refresh
                                frmMain.picHotbar.Refresh
                            End If
                        End If
                    End If
                Next
            End If
            
            ' check if we need to unlock the player's spell casting restriction
            If SpellBuffer > 0 Then
                If SpellBufferTimer + (Spell(PlayerSpells(SpellBuffer)).CastTime * 1000) < tick Then
                    SpellBuffer = 0
                    SpellBufferTimer = 0
                End If
            End If

            If CanMoveNow Then
                Call CheckMovement ' Check if player is trying to move
                Call CheckAttack   ' Check to see if player is trying to attack
            End If

            ' Change map animation every 250 milliseconds
            If MapAnimTimer < tick Then
                MapAnim = Not MapAnim
                MapAnimTimer = tick + 250
            End If
            
            ' Update inv animation
            If numitems > 0 Then
                If tmr100 < tick Then
                    DrawAnimatedInvItems
                    tmr100 = tick + 100
                End If
            End If
            
            For i = 1 To MAX_BYTE
                CheckAnimInstance i
            Next
            
            tmr25 = tick + 25
        End If
        
        If tick > EventChatTimer Then
            If frmMain.lblEventChat.Visible = False Then
                If frmMain.picEventChat.Visible Then
                    frmMain.picEventChat.Visible = False
                End If
            End If
        End If
        

        ' Process input before rendering, otherwise input will be behind by 1 frame
        If WalkTimer < tick Then

            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    Call ProcessMovement(i)
                End If
            Next i

            ' Process npc movements (actually move them)
            For i = 1 To Npc_HighIndex
                If Map.NPC(i) > 0 Then
                    Call ProcessNpcMovement(i)
                End If
            Next i
            
            If Map.CurrentEvents > 0 Then
                For i = 1 To Map.CurrentEvents
                    Call ProcessEventMovement(i)
                Next i
            End If

            WalkTimer = tick + 30 ' edit this value to change WalkTimer
        End If
        
        ' fog scrolling
        If fogtmr < tick Then
            If CurrentFogSpeed > 0 Then
                ' move
                fogOffsetX = fogOffsetX - 1
                fogOffsetY = fogOffsetY - 1
                ' reset
                If fogOffsetX < -256 Then fogOffsetX = 0
                If fogOffsetY < -256 Then fogOffsetY = 0
                fogtmr = tick + 255 - CurrentFogSpeed
            End If
        End If
        
        If tmr500 < tick Then
            ' animate waterfalls
            Select Case waterfallFrame
                Case 0
                    waterfallFrame = 1
                Case 1
                    waterfallFrame = 2
                Case 2
                    waterfallFrame = 0
            End Select
            
            ' animate autotiles
            Select Case autoTileFrame
                Case 0
                    autoTileFrame = 1
                Case 1
                    autoTileFrame = 2
                Case 2
                    autoTileFrame = 0
            End Select
            tmr500 = tick + 500
        End If
        
        ProcessWeather
        
        If Fadetmr < tick Then
            If FadeType <> 2 Then
                If FadeType = 1 Then
                    If FadeAmount = 255 Then
                        
                    Else
                        FadeAmount = FadeAmount + 5
                    End If
                ElseIf FadeType = 0 Then
                    If FadeAmount = 0 Then
                    
                    Else
                        FadeAmount = FadeAmount - 5
                    End If
                End If
            End If
            Fadetmr = tick + 30
        End If

        ' *********************
        ' ** Render Graphics **
        ' *********************
        If renderTmr < tick Then
           Call Render_Graphics
            Call UpdateSounds
           renderTmr = tick + 15
        End If
        DoEvents

        ' Lock fps
        If Not FPS_Lock Then
            Do While timeGetTime < tick + 15
                DoEvents
                Sleep 1
            Loop
        End If
        
        ' Calculate fps
        If TickFPS < tick Then
            GameFPS = FPS
            TickFPS = tick + 1000
            FPS = 0
        Else
            FPS = FPS + 1
        End If

    Loop

    frmMain.Visible = False

    If isLogging Then
        isLogging = False
        frmMain.picScreen.Visible = False
        frmMenu.Visible = True
        GettingMap = True
        StopMusic
        PlayMusic options.MenuMusic
    Else
        ' Shutdown the game
        frmLoad.Visible = True
        Call SetStatus("Destroying game data...")
        Call DestroyGame
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GameLoop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ProcessMovement(ByVal Index As Long)
Dim MovementSpeed As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if player is walking, and if so process moving them over
    Select Case Player(Index).Moving
        Case PlayerWalking: MovementSpeed = RUN_SPEED '((ElapsedTime / 1000) * (RUN_SPEED * SIZE_X))
        Case PlayerRunning: MovementSpeed = WALK_SPEED '((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X) + (GetPlayerStat(Index, Stats.Agility) / 10))
        Case Else: Exit Sub
    End Select
    
    Select Case GetPlayerDir(Index)
        Case DirectionUp
            Player(Index).yOffset = Player(Index).yOffset - MovementSpeed
            If Player(Index).yOffset < 0 Then Player(Index).yOffset = 0
        Case DirectionDown
            Player(Index).yOffset = Player(Index).yOffset + MovementSpeed
            If Player(Index).yOffset > 0 Then Player(Index).yOffset = 0
        Case DirectionLeft
            Player(Index).xOffset = Player(Index).xOffset - MovementSpeed
            If Player(Index).xOffset < 0 Then Player(Index).xOffset = 0
        Case DirectionRight
            Player(Index).xOffset = Player(Index).xOffset + MovementSpeed
            If Player(Index).xOffset > 0 Then Player(Index).xOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If Player(Index).Moving > 0 Then
        If GetPlayerDir(Index) = DirectionRight Or GetPlayerDir(Index) = DirectionDown Then
            If (Player(Index).xOffset >= 0) And (Player(Index).yOffset >= 0) Then
                Player(Index).Moving = 0
                If Player(Index).Step = 1 Then
                    Player(Index).Step = 3
                Else
                    Player(Index).Step = 1
                End If
            End If
        Else
            If (Player(Index).xOffset <= 0) And (Player(Index).yOffset <= 0) Then
                Player(Index).Moving = 0
                If Player(Index).Step = 1 Then
                    Player(Index).Step = 3
                Else
                    Player(Index).Step = 1
                End If
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ProcessMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if NPC is walking, and if so process moving them over
    If MapNpc(MapNpcNum).Moving = PlayerWalking Then
        
        Select Case MapNpc(MapNpcNum).Dir
            Case DirectionUp
                MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset - ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
                If MapNpc(MapNpcNum).yOffset < 0 Then MapNpc(MapNpcNum).yOffset = 0
                
            Case DirectionDown
                MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset + ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
                If MapNpc(MapNpcNum).yOffset > 0 Then MapNpc(MapNpcNum).yOffset = 0
                
            Case DirectionLeft
                MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset - ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
                If MapNpc(MapNpcNum).xOffset < 0 Then MapNpc(MapNpcNum).xOffset = 0
                
            Case DirectionRight
                MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset + ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
                If MapNpc(MapNpcNum).xOffset > 0 Then MapNpc(MapNpcNum).xOffset = 0
                
        End Select
    
        ' Check if completed walking over to the next tile
        If MapNpc(MapNpcNum).Moving > 0 Then
            If MapNpc(MapNpcNum).Dir = DirectionRight Or MapNpc(MapNpcNum).Dir = DirectionDown Then
                If (MapNpc(MapNpcNum).xOffset >= 0) And (MapNpc(MapNpcNum).yOffset >= 0) Then
                    MapNpc(MapNpcNum).Moving = 0
                    If MapNpc(MapNpcNum).Step = 1 Then
                        MapNpc(MapNpcNum).Step = 3
                    Else
                        MapNpc(MapNpcNum).Step = 1
                    End If
                End If
            Else
                If (MapNpc(MapNpcNum).xOffset <= 0) And (MapNpc(MapNpcNum).yOffset <= 0) Then
                    MapNpc(MapNpcNum).Moving = 0
                    If MapNpc(MapNpcNum).Step = 1 Then
                        MapNpc(MapNpcNum).Step = 3
                    Else
                        MapNpc(MapNpcNum).Step = 1
                    End If
                End If
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ProcessNpcMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckMapGetItem()
Dim Buffer As New clsBuffer

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer

    If timeGetTime > Player(MyIndex).MapGetTimer + 250 Then
        If Trim$(MyText) = vbNullString Then
            Player(MyIndex).MapGetTimer = timeGetTime
            Buffer.WriteLong CMapGetItem
            SendData Buffer.ToArray()
        End If
    End If

    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckMapGetItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAttack()
Dim Buffer As clsBuffer
Dim attackspeed As Long, X As Long, Y As Long, i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If ControlDown Then
    
        Select Case Player(MyIndex).Dir
            Case DirectionUp
                X = GetPlayerX(MyIndex)
                Y = GetPlayerY(MyIndex) - 1
            Case DirectionDown
                X = GetPlayerX(MyIndex)
                Y = GetPlayerY(MyIndex) + 1
            Case DirectionLeft
                X = GetPlayerX(MyIndex) - 1
                Y = GetPlayerY(MyIndex)
            Case DirectionRight
                X = GetPlayerX(MyIndex) + 1
                Y = GetPlayerY(MyIndex)
        End Select
        
        If timeGetTime > Player(MyIndex).EventTimer Then
            For i = 1 To Map.CurrentEvents
                If Map.MapEvents(i).Visible = 1 Then
                    If Map.MapEvents(i).X = X And Map.MapEvents(i).Y = Y Then
                        Set Buffer = New clsBuffer
                        Buffer.WriteLong CEvent
                        Buffer.WriteLong i
                        SendData Buffer.ToArray()
                        Set Buffer = Nothing
                        Player(MyIndex).EventTimer = timeGetTime + 200
                    End If
                End If
            Next
        End If
    
        If SpellBuffer > 0 Then Exit Sub ' currently casting a spell, can't attack
        If StunDuration > 0 Then Exit Sub ' stunned, can't attack

        ' speed from weapon
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(MyIndex, Weapon)).speed
        Else
            attackspeed = 1000
        End If

        If Player(MyIndex).AttackTimer + attackspeed < timeGetTime Then
            If Player(MyIndex).Attacking = 0 Then

                With Player(MyIndex)
                    .Attacking = 1
                    .AttackTimer = timeGetTime
                End With

                If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
                    If Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Pic > 0 Then
                        ' projectile
                        Set Buffer = New clsBuffer
                            Buffer.WriteLong CProjecTileAttack
                            SendData Buffer.ToArray()
                            Set Buffer = Nothing
                            Exit Sub
                    End If
                End If
                        
                ' non projectile
                Set Buffer = New clsBuffer
                Buffer.WriteLong CAttack
                SendData Buffer.ToArray()
                Set Buffer = Nothing
            End If
        End If

    End If
    

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckAttack", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function IsTryingToMove() As Boolean
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If DirUp Or DirDown Or DirLeft Or DirRight Then
        IsTryingToMove = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsTryingToMove", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function CanMove() As Boolean
Dim d As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    CanMove = True

    ' Make sure they aren't trying to move when they are already moving
    If Player(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If
    
    If InEvent Then
        CanMove = False
        Exit Function
    End If

    ' Make sure they haven't just casted a spell
    If SpellBuffer > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' make sure they're not stunned
    If StunDuration > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' make sure they're not in a shop
    If InShop > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' not in bank
    If InBank Then
        'CanMove = False
        'Exit Function
        InBank = False
        frmMain.picCover.Visible = False
        frmMain.picBank.Visible = False
    End If

    d = GetPlayerDir(MyIndex)

    If DirUp Then
        Call SetPlayerDir(MyIndex, DirectionUp)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            If CheckDirection(DirectionUp) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DirectionUp Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Up > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirDown Then
        Call SetPlayerDir(MyIndex, DirectionDown)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < Map.MaxY Then
            If CheckDirection(DirectionDown) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DirectionDown Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Down > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirLeft Then
        Call SetPlayerDir(MyIndex, DirectionLeft)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DirectionLeft) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DirectionLeft Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Left > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirRight Then
        Call SetPlayerDir(MyIndex, DirectionRight)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < Map.MaxX Then
            If CheckDirection(DirectionRight) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DirectionRight Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Right > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CanMove", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function CheckDirection(ByVal direction As Byte) As Boolean
Dim X As Long
Dim Y As Long
Dim i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    CheckDirection = False
    
    ' check directional blocking
    If isDirBlocked(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).DirBlock, direction + 1) Then
        CheckDirection = True
        Exit Function
    End If

    Select Case direction
        Case DirectionUp
            X = GetPlayerX(MyIndex)
            Y = GetPlayerY(MyIndex) - 1
        Case DirectionDown
            X = GetPlayerX(MyIndex)
            Y = GetPlayerY(MyIndex) + 1
        Case DirectionLeft
            X = GetPlayerX(MyIndex) - 1
            Y = GetPlayerY(MyIndex)
        Case DirectionRight
            X = GetPlayerX(MyIndex) + 1
            Y = GetPlayerY(MyIndex)
    End Select

    ' Check to see if the map tile is blocked or not
    If Map.Tile(X, Y).Type = TileBlocked Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the map tile is blocked or not
    If Map.Tile(X, Y).Type = TileWater Then
        CheckDirection = True
        Exit Function
    End If
    
    ' Check to see if the map tile is tree or not
    If Map.Tile(X, Y).Type = TileResource Then
        CheckDirection = True
        Exit Function
    End If

    
    ' Check to see if a player is already on that tile
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
            If GetPlayerX(i) = X Then
                If GetPlayerY(i) = Y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next i

    ' Check to see if a npc is already on that tile
    For i = 1 To Npc_HighIndex
        If MapNpc(i).Num > 0 Then
            If MapNpc(i).X = X Then
                If MapNpc(i).Y = Y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next
    
    For i = 1 To Map.CurrentEvents
        If Map.MapEvents(i).Visible = 1 Then
            If Map.MapEvents(i).X = X Then
                If Map.MapEvents(i).Y = Y Then
                    If Map.MapEvents(i).WalkThrough = 0 Then
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "checkDirection", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub CheckMovement()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If IsTryingToMove Then
        If CanMove Then

            ' Check if player has the shift key down for running
            If ShiftDown Then
                Player(MyIndex).Moving = PlayerWalking
            Else
                Player(MyIndex).Moving = PlayerRunning
            End If

            Select Case GetPlayerDir(MyIndex)
                Case DirectionUp
                    Call SendPlayerMove
                    Player(MyIndex).yOffset = PIC_Y
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                Case DirectionDown
                    Call SendPlayerMove
                    Player(MyIndex).yOffset = PIC_Y * -1
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                Case DirectionLeft
                    Call SendPlayerMove
                    Player(MyIndex).xOffset = PIC_X
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                Case DirectionRight
                    Call SendPlayerMove
                    Player(MyIndex).xOffset = PIC_X * -1
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
            End Select

            If Player(MyIndex).xOffset = 0 Then
                If Player(MyIndex).yOffset = 0 Then
                    If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TileWarp Then
                        GettingMap = True
                    End If
                End If
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function isInBounds()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If (CurX >= 0) Then
        If (CurX <= Map.MaxX) Then
            If (CurY >= 0) Then
                If (CurY <= Map.MaxY) Then
                    isInBounds = True
                End If
            End If
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "isInBounds", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub UpdateDrawMapName()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    DrawMapNameX = ((MAX_MAPX + 1) * PIC_X / 2) - (getWidth(Font_Default, Trim$(Map.Name)) / 2)
    DrawMapNameY = 1

    Select Case Map.Moral
        Case moralnone
            DrawMapNameColor = QBColor(BrightRed)
        Case moralsafe
            DrawMapNameColor = QBColor(White)
        Case moralmember
            DrawMapNameColor = QBColor(Yellow)
        Case Else
            DrawMapNameColor = QBColor(White)
    End Select

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateDrawMapName", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UseItem()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check for subscript out of range
    If InventoryItemSelected < 1 Or InventoryItemSelected > MAX_INV Then
        Exit Sub
    End If

    Call SendUseItem(InventoryItemSelected)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UseItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ForgetSpell(ByVal spellslot As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check for subscript out of range
    If spellslot < 1 Or spellslot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If SpellCD(spellslot) > 0 Then
        AddText "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If SpellBuffer = spellslot Then
        AddText "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If
    
    If PlayerSpells(spellslot) > 0 Then
        Set Buffer = New clsBuffer
        Buffer.WriteLong CForgetSpell
        Buffer.WriteLong spellslot
        SendData Buffer.ToArray()
        Set Buffer = Nothing
    Else
        AddText "No spell here.", BrightRed
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ForgetSpell", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CastSpell(ByVal spellslot As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check for subscript out of range
    If spellslot < 1 Or spellslot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    If SpellCD(spellslot) > 0 Then
        AddText "Spell has not cooled down yet!", BrightRed
        Exit Sub
    End If
    
    If PlayerSpells(spellslot) = 0 Then Exit Sub

    ' Check if player has enough MP
    If GetPlayerVital(MyIndex, Vitals.MP) < Spell(PlayerSpells(spellslot)).MPCost Then
        Call AddText("Not enough MP to cast " & Trim$(Spell(PlayerSpells(spellslot)).Name) & ".", BrightRed)
        Exit Sub
    End If

    If PlayerSpells(spellslot) > 0 Then
        If timeGetTime > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Set Buffer = New clsBuffer
                Buffer.WriteLong CCast
                Buffer.WriteLong spellslot
                SendData Buffer.ToArray()
                Set Buffer = Nothing
                SpellBuffer = spellslot
                SpellBufferTimer = timeGetTime
            Else
                Call AddText("Cannot cast while walking!", BrightRed)
            End If
        End If
    Else
        Call AddText("No spell here.", BrightRed)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CastSpell", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearTempTile()
Dim X As Long
Dim Y As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    ReDim TempTile(0 To Map.MaxX, 0 To Map.MaxY)

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            TempTile(X, Y).DoorOpen = NO
        Next
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearTempTile", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DevMsg(ByVal text As String, ByVal color As Byte)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If InGame Then
        If GetPlayerAccess(MyIndex) > RankDeveloper Then
            Call AddText(text, color)
        End If
    End If

    Debug.Print text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DevMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function TwipsToPixels(ByVal twip_val As Long, ByVal XorY As Byte) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If XorY = 0 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelY
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "TwipsToPixels", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function PixelsToTwips(ByVal pixel_val As Long, ByVal XorY As Byte) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If XorY = 0 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelY
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "PixelsToTwips", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ConvertCurrency(ByVal Amount As Long) As String
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If Int(Amount) < 10000 Then
        ConvertCurrency = Amount
    ElseIf Int(Amount) < 999999 Then
        ConvertCurrency = Int(Amount / 1000) & "k"
    ElseIf Int(Amount) < 999999999 Then
        ConvertCurrency = Int(Amount / 1000000) & "m"
    Else
        ConvertCurrency = Int(Amount / 1000000000) & "b"
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertCurrency", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub DrawPing()
Dim PingToDraw As String

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    PingToDraw = Ping

    Select Case Ping
        Case -1
            PingToDraw = "Syncing"
        Case 0 To 5
            PingToDraw = "Local"
    End Select

    frmMain.lblPing.Caption = PingToDraw
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPing", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdateSpellWindow(ByVal spellnum As Long, ByVal X As Long, ByVal Y As Long)
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    ' check for off-screen
    If Y + frmMain.picSpellDesc.Height > frmMain.ScaleHeight Then
        Y = frmMain.ScaleHeight - frmMain.picSpellDesc.Height
    End If
    
    With frmMain
        .picSpellDesc.Top = Y
        .picSpellDesc.Left = X
        .picSpellDesc.Visible = True
        
        If LastSpellDesc = spellnum Then Exit Sub
        
        .lblSpellName.Caption = Trim$(Spell(spellnum).Name)
        .lblSpellDesc.Caption = Trim$(Spell(spellnum).Desc)
        frmMain.picSpellDescPic.Refresh
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdteSpellWindow", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdateDescWindow(ByVal itemnum As Long, ByVal X As Long, ByVal Y As Long)
Dim i As Long
Dim FirstLetter As String * 1
Dim Name As String
Dim ATS As Double
ATS = Trim$(Item(itemnum).speed) / 1000
Dim DMG As Long
DMG = Trim$(Item(itemnum).Data2)
'stat add
Dim str As Long
str = Trim$(Item(itemnum).Add_Stat(1))
Dim ENDu As Long
ENDu = Trim$(Item(itemnum).Add_Stat(2))
Dim AGI As Long
AGI = Trim$(Item(itemnum).Add_Stat(4))
Dim INTe As Long
INTe = Trim$(Item(itemnum).Add_Stat(3))
Dim WILL As Long
WILL = Trim$(Item(itemnum).Add_Stat(5))
Dim CRIT As Long
CRIT = Trim$(Item(itemnum).Add_Stat(6))
'stat req
Dim STRq As Long
STRq = Trim$(Item(itemnum).Stat_Req(1))
Dim ENDuq As Long
ENDuq = Trim$(Item(itemnum).Stat_Req(2))
Dim AGIq As Long
AGIq = Trim$(Item(itemnum).Stat_Req(4))
Dim INTeq As Long
INTeq = Trim$(Item(itemnum).Stat_Req(3))
Dim WILLq As Long
WILLq = Trim$(Item(itemnum).Stat_Req(5))
Dim CRITq As Long
CRITq = Trim$(Item(itemnum).Stat_Req(6))
Dim LVLq As Long
LVLq = Trim$(Item(itemnum).LevelReq)
Dim CLASSq As Long
CLASSq = Trim$(Item(itemnum).ClassReq)
'player stats
Dim Index As Long
Dim plvl As Long
Dim pstr As Long
Dim pend As Long
Dim pint As Long
Dim pagi As Long
Dim pwill As Long
Dim pCrit As Long
Dim pclass As Long


pclass = Player(MyIndex).Class

plvl = Player(MyIndex).Level
pstr = Player(MyIndex).Stat(1)
pend = Player(MyIndex).Stat(2)
pint = Player(MyIndex).Stat(3)
pagi = Player(MyIndex).Stat(4)
pwill = Player(MyIndex).Stat(5)
pCrit = Player(MyIndex).Stat(6)


'no DescWindow for currency
   If Trim$(Item(itemnum).Type) = 7 Then
    Exit Sub
   End If
   
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    FirstLetter = LCase$(Left$(Trim$(Item(itemnum).Name), 1))
   
    If FirstLetter = "$" Then
        Name = (Mid$(Trim$(Item(itemnum).Name), 2, Len(Trim$(Item(itemnum).Name)) - 1))
    Else
        Name = Trim$(Item(itemnum).Name)
    End If
    
    ' check for off-screen
    If Y + frmMain.picItemDesc.Height > frmMain.ScaleHeight Then
        Y = frmMain.ScaleHeight - frmMain.picItemDesc.Height
    End If
    
    ' set z-order
    frmMain.picItemDesc.ZOrder (0)

    With frmMain
        .picItemDesc.Top = Y
        .picItemDesc.Left = X
        .picItemDesc.Visible = True

        If LastItemDesc = itemnum Then Exit Sub ' exit out after setting x + y so we don't reset values

        ' set the name
        Select Case Item(itemnum).Rarity
            Case 0 ' white
                .lblItemName.ForeColor = RGB(255, 255, 255)
            Case 1 ' green
                .lblItemName.ForeColor = RGB(117, 198, 92)
            Case 2 ' blue
                .lblItemName.ForeColor = RGB(103, 140, 224)
            Case 3 ' maroon
                .lblItemName.ForeColor = RGB(205, 34, 0)
            Case 4 ' purple
                .lblItemName.ForeColor = RGB(193, 104, 204)
            Case 5 ' orange
                .lblItemName.ForeColor = RGB(217, 150, 64)
        End Select
        
                ' set captions
        .lblItemName.Caption = Name
        .lblItemDesc.Caption = Trim$(Item(itemnum).Desc)
        .lblPrice.Caption = "Price: " & Trim$(Item(itemnum).Price)
        .lblStat.Caption = "No Stats."
        .lblDMG = "Have no damage."
        .lblitemspeed.Caption = ""
        
        
        'to do class ----> .lblclass_req.Caption = Class(MyClass).Name
        
        
                ' Item Type
Select Case Item(itemnum).Type
            Case 0 ' None
                .lblitemType.Caption = "ETC"
            Case 1 ' Weapon
                .lblitemType.Caption = "Weapon"
                .lblitemspeed.Caption = ATS & "s"
                .lblStat.Caption = "STR: +" & str & vbCrLf & "END: +" & ENDu & vbCrLf & "INT: +" & INTe & vbCrLf & "AGI: +" & AGI & vbCrLf & "WILL: +" & WILL & vbCrLf & "CRIT%: +" & CRIT & vbCrLf
                .lblDMG = "Damage: " & DMG
            Case 2 ' Armor
                .lblitemType.Caption = "Armor"
                .lblitemspeed.Caption = ""
                .lblStat.Caption = "STR: +" & str & vbCrLf & "END: +" & ENDu & vbCrLf & "INT: +" & INTe & vbCrLf & "AGI: +" & AGI & vbCrLf & "WILL: +" & WILL & vbCrLf & "CRIT%: +" & CRIT & vbCrLf
                .lblDMG = "" 'show armor
            Case 3 ' Helmet
                .lblitemType.Caption = "Helmet"
                .lblitemspeed.Caption = ""
                .lblStat.Caption = "STR: +" & str & vbCrLf & "END: +" & ENDu & vbCrLf & "INT: +" & INTe & vbCrLf & "AGI: +" & AGI & vbCrLf & "WILL: +" & WILL & vbCrLf & "CRIT%: +" & CRIT & vbCrLf
                .lblDMG = "" 'show armor
            Case 4 ' LEGS
                .lblitemType.Caption = "Legs"
                .lblitemspeed.Caption = ""
                .lblStat.Caption = "STR: +" & str & vbCrLf & "END: +" & ENDu & vbCrLf & "INT: +" & INTe & vbCrLf & "AGI: +" & AGI & vbCrLf & "WILL: +" & WILL & vbCrLf & "CRIT%: +" & CRIT & vbCrLf
                .lblDMG = "" 'show armor
            Case 5 ' BOOTS
                .lblitemType.Caption = "Boots"
                .lblitemspeed.Caption = ""
                .lblStat.Caption = "STR: +" & str & vbCrLf & "END: +" & ENDu & vbCrLf & "INT: +" & INTe & vbCrLf & "AGI: +" & AGI & vbCrLf & "WILL: +" & WILL & vbCrLf & "CRIT%: +" & CRIT & vbCrLf
                .lblDMG = "" 'show armor
            Case 6 ' Gloves
                .lblitemType.Caption = "Gloves"
                .lblitemspeed.Caption = ""
                .lblStat.Caption = "STR: +" & str & vbCrLf & "END: +" & ENDu & vbCrLf & "INT: +" & INTe & vbCrLf & "AGI: +" & AGI & vbCrLf & "WILL: +" & WILL & vbCrLf & "CRIT%: +" & CRIT & vbCrLf
                .lblDMG = "" 'show armor
            Case 7 ' RING
                .lblitemType.Caption = "Ring"
                .lblitemspeed.Caption = ""
                .lblStat.Caption = "STR: +" & str & vbCrLf & "END: +" & ENDu & vbCrLf & "INT: +" & INTe & vbCrLf & "AGI: +" & AGI & vbCrLf & "WILL: +" & WILL & vbCrLf & "CRIT%: +" & CRIT & vbCrLf
                .lblDMG = "" 'show armor
             Case 8 ' ENCHANT
                .lblitemType.Caption = "Enchant"
                .lblitemspeed.Caption = ""
                .lblStat.Caption = "STR: +" & str & vbCrLf & "END: +" & ENDu & vbCrLf & "INT: +" & INTe & vbCrLf & "AGI: +" & AGI & vbCrLf & "WILL: +" & WILL & vbCrLf & "CRIT%: +" & CRIT & vbCrLf
                .lblDMG = "" 'show Enchant
            Case 9 ' Shield
                .lblitemType.Caption = "Shield"
                .lblitemspeed.Caption = ""
                .lblStat.Caption = "STR: +" & str & vbCrLf & "END: +" & ENDu & vbCrLf & "INT: +" & INTe & vbCrLf & "AGI: +" & AGI & vbCrLf & "WILL: +" & WILL & vbCrLf & "CRIT%: +" & CRIT & vbCrLf
                .lblDMG = "" 'show armor
            Case 10 ' orange
                .lblitemType.Caption = "Consume"
            Case 11 ' orange
                .lblitemType.Caption = "Key"
            Case 12 ' orange
                .lblitemType.Caption = "Currency"
            Case 13 ' orange
                .lblitemType.Caption = "Spell Scroll"
            Case 14 ' orange
                .lblitemType.Caption = "Recipe"
            End Select
        
        
        
        
        

        
        
        'is player meets REq:
        
        If plvl < LVLq Then
            .lbllvl_req.ForeColor = RGB(255, 0, 0)
            .lbllvl_req.Caption = "Level: " & LVLq
        Else
            .lbllvl_req.ForeColor = RGB(255, 255, 255)
            .lbllvl_req.Caption = "Level: " & LVLq
        End If
        
        If pstr < STRq Then
            .lblstr_req.ForeColor = RGB(255, 0, 0)
            .lblstr_req.Caption = "Str: " & STRq
        Else
            .lblstr_req.ForeColor = RGB(255, 255, 255)
            .lblstr_req.Caption = "Str: " & STRq
        End If
        
        If pend < ENDuq Then
            .lblend_req.ForeColor = RGB(255, 0, 0)
            .lblend_req.Caption = "End: " & ENDuq
        Else
            .lblend_req.ForeColor = RGB(255, 255, 255)
            .lblend_req.Caption = "End: " & ENDuq
        End If
        
        If pagi < AGIq Then
           .lblagi_req.ForeColor = RGB(255, 0, 0)
           .lblagi_req.Caption = "Agi: " & AGIq
        Else
           .lblagi_req.ForeColor = RGB(255, 255, 255)
           .lblagi_req.Caption = "Agi: " & AGIq
           End If
           
        If pint < INTeq Then
           .lblint_req.ForeColor = RGB(255, 0, 0)
           .lblint_req.Caption = "Int: " & INTeq
        Else
           .lblint_req.ForeColor = RGB(255, 255, 255)
           .lblint_req.Caption = "Int: " & INTeq
           End If
           
        If pwill < WILLq Then
           .lblwill_req.ForeColor = RGB(255, 0, 0)
           .lblwill_req.Caption = "Will: " & WILLq
        Else
           .lblwill_req.ForeColor = RGB(255, 255, 255)
           .lblwill_req.Caption = "Will: " & WILLq
           End If
           
  '      If pclass = CLASSq Then
   '        .lblclass_req.ForeColor = RGB(255, 255, 255)

      '  Else
      '     .lblclass_req.ForeColor = RGB(255, 0, 0)
     '      End If

        
        
        
        
        
        
        

        'Trim$(Item(itemnum).data1)
        'to do spell Details
       
        ' render the item
        frmMain.picItemDescPic.Refresh
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateDescWindow", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub CacheResources()
Dim X As Long, Y As Long, Resource_Count As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    Resource_Count = 0

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            If Map.Tile(X, Y).Type = TileResource Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve MapResource(0 To Resource_Count)
                MapResource(Resource_Count).X = X
                MapResource(Resource_Count).Y = Y
            End If
        Next
    Next

    Resource_Index = Resource_Count
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CacheResources", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CreateActionMsg(ByVal Message As String, ByVal color As Integer, ByVal MsgType As Byte, ByVal X As Long, ByVal Y As Long)
Dim i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    ActionMsgIndex = ActionMsgIndex + 1
    If ActionMsgIndex >= MAX_BYTE Then ActionMsgIndex = 1

    With ActionMsg(ActionMsgIndex)
        .Message = Message
        .color = color
        .Type = MsgType
        .Created = timeGetTime
        .Scroll = 1
        .X = X
        .Y = Y
    End With

    If ActionMsg(ActionMsgIndex).Type = ACTIONMSG_SCROLL Then
        ActionMsg(ActionMsgIndex).Y = ActionMsg(ActionMsgIndex).Y + Rand(-2, 6)
        ActionMsg(ActionMsgIndex).X = ActionMsg(ActionMsgIndex).X + Rand(-8, 8)
    End If
    
    ' find the new high index
    For i = MAX_BYTE To 1 Step -1
        If ActionMsg(i).Created > 0 Then
            Action_HighIndex = i + 1
            Exit For
        End If
    Next
    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CreateActionMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearActionMsg(ByVal Index As Byte)
Dim i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    ActionMsg(Index).Message = vbNullString
    ActionMsg(Index).Created = 0
    ActionMsg(Index).Type = 0
    ActionMsg(Index).color = 0
    ActionMsg(Index).Scroll = 0
    ActionMsg(Index).X = 0
    ActionMsg(Index).Y = 0
    
    ' find the new high index
    For i = MAX_BYTE To 1 Step -1
        If ActionMsg(i).Created > 0 Then
            Action_HighIndex = i + 1
            Exit For
        End If
    Next
    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearActionMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAnimInstance(ByVal Index As Long)
Dim looptime As Long
Dim Layer As Long
Dim FrameCount As Long
Dim lockindex As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    ' if doesn't exist then exit sub
    If AnimInstance(Index).Animation <= 0 Then Exit Sub
    If AnimInstance(Index).Animation >= MAX_ANIMATIONS Then Exit Sub
    
    For Layer = 0 To 1
        If AnimInstance(Index).Used(Layer) Then
            looptime = Animation(AnimInstance(Index).Animation).looptime(Layer)
            FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
            
            ' if zero'd then set so we don't have extra loop and/or frame
            If AnimInstance(Index).frameIndex(Layer) = 0 Then AnimInstance(Index).frameIndex(Layer) = 1
            If AnimInstance(Index).LoopIndex(Layer) = 0 Then AnimInstance(Index).LoopIndex(Layer) = 1
            
            ' check if frame timer is set, and needs to have a frame change
            If AnimInstance(Index).timer(Layer) + looptime <= timeGetTime Then
                ' check if out of range
                If AnimInstance(Index).frameIndex(Layer) >= FrameCount Then
                    AnimInstance(Index).LoopIndex(Layer) = AnimInstance(Index).LoopIndex(Layer) + 1
                    If AnimInstance(Index).LoopIndex(Layer) > Animation(AnimInstance(Index).Animation).LoopCount(Layer) Then
                        AnimInstance(Index).Used(Layer) = False
                    Else
                        AnimInstance(Index).frameIndex(Layer) = 1
                    End If
                Else
                    AnimInstance(Index).frameIndex(Layer) = AnimInstance(Index).frameIndex(Layer) + 1
                End If
                AnimInstance(Index).timer(Layer) = timeGetTime
            End If
        End If
    Next
    
    ' if neither layer is used, clear
    If AnimInstance(Index).Used(0) = False And AnimInstance(Index).Used(1) = False Then ClearAnimInstance (Index)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "checkAnimInstance", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub OpenShop(ByVal shopnum As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    InShop = shopnum
    ShopAction = 0
    frmMain.picCover.Visible = True
    frmMain.picShop.Visible = True
    frmMain.picShopItems.Refresh
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "OpenShop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetBankItemNum(ByVal bankslot As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If bankslot = 0 Then
        GetBankItemNum = 0
        Exit Function
    End If
    
    If bankslot > MAX_BANK Then
        GetBankItemNum = 0
        Exit Function
    End If
    
    GetBankItemNum = Bank.Item(bankslot).Num
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetBankItemNum", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetBankItemNum(ByVal bankslot As Long, ByVal itemnum As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    Bank.Item(bankslot).Num = itemnum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetBankItemNum", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetBankItemValue(ByVal bankslot As Long) As Long
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    GetBankItemValue = Bank.Item(bankslot).Value
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetBankItemValue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetBankItemValue(ByVal bankslot As Long, ByVal ItemValue As Long)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    Bank.Item(bankslot).Value = ItemValue
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetBankItemValue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' BitWise Operators for directional blocking
Public Sub setDirBlock(ByRef blockvar As Byte, ByRef Dir As Byte, ByVal block As Boolean)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If block Then
        blockvar = blockvar Or (2 ^ Dir)
    Else
        blockvar = blockvar And Not (2 ^ Dir)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "setDirBlock", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef Dir As Byte) As Boolean
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not blockvar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "isDirBlocked", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsHotbarSlot(ByVal X As Single, ByVal Y As Single) As Long
Dim Top As Long, Left As Long
Dim i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    IsHotbarSlot = 0

    For i = 1 To MAX_HOTBAR
        Top = HotbarTop
        Left = HotbarLeft + ((HotbarOffsetX + 32) * (((i - 1) Mod MAX_HOTBAR)))
        If X >= Left And X <= Left + PIC_X Then
            If Y >= Top And Y <= Top + PIC_Y Then
                IsHotbarSlot = i
                Exit Function
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsHotbarSlot", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub PlayMapSound(ByVal X As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim soundName As String

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If entityNum <= 0 Then Exit Sub
    
    ' find the sound
    Select Case entityType
        ' animations
        Case SoundEntity.seAnimation
            If entityNum > MAX_ANIMATIONS Then Exit Sub
            soundName = Trim$(Animation(entityNum).sound)
            
        ' items
        Case SoundEntity.seItem
            If entityNum > MAX_ITEMS Then Exit Sub
            soundName = Trim$(Item(entityNum).sound)
        ' npcs
        Case SoundEntity.seNpc
            If entityNum > MAX_NPCS Then Exit Sub
            soundName = Trim$(NPC(entityNum).sound)
        ' resources
        Case SoundEntity.seResource
            If entityNum > MAX_RESOURCES Then Exit Sub
            soundName = Trim$(Resource(entityNum).sound)
        ' spells
        Case SoundEntity.seSpell
            If entityNum > MAX_SPELLS Then Exit Sub
            soundName = Trim$(Spell(entityNum).sound)
        ' other
        Case Else
            Exit Sub
    End Select
    
    ' exit out if it's not set
    If Trim$(soundName) = "None." Then Exit Sub

    ' play the sound
    PlaySound soundName, X, Y
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayMapSound", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Dialogue(ByVal diTitle As String, ByVal diText As String, ByVal diIndex As Long, Optional ByVal isYesNo As Boolean = False, Optional ByVal Data1 As Long = 0)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    ' exit out if we've already got a dialogue open
    If dialogueIndex > 0 Then Exit Sub
    
    ' set global dialogue index
    dialogueIndex = diIndex
    
    ' set the global dialogue data
    dialogueData1 = Data1

    ' set the captions
    frmMain.lblDialogue_Title.Caption = diTitle
    frmMain.lblDialogue_Text.Caption = diText
    
    ' show/hide buttons
    If Not isYesNo Then
        frmMain.lblDialogue_Button(1).Visible = True ' Okay button
        frmMain.lblDialogue_Button(2).Visible = False ' Yes button
        frmMain.lblDialogue_Button(3).Visible = False ' No button
    Else
        frmMain.lblDialogue_Button(1).Visible = False ' Okay button
        frmMain.lblDialogue_Button(2).Visible = True ' Yes button
        frmMain.lblDialogue_Button(3).Visible = True ' No button
    End If
    
    ' show the dialogue box
    frmMain.picDialogue.Visible = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Dialogue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub dialogueHandler(ByVal Index As Long)
    ' find out which button
    If Index = 1 Then ' okay button
        ' dialogue index
        Select Case dialogueIndex
        
        End Select
    ElseIf Index = 2 Then ' yes button
        ' dialogue index
        Select Case dialogueIndex
            Case DialogueTrade
                SendAcceptTradeRequest
            Case DialogueForget
                ForgetSpell dialogueData1
            Case DialogueParty
                SendAcceptParty
        End Select
    ElseIf Index = 3 Then ' no button
        ' dialogue index
        Select Case dialogueIndex
            Case DialogueTrade
                SendDeclineTradeRequest
            Case DialogueParty
                SendDeclineParty
        End Select
    End If
End Sub

Sub ProcessEventMovement(ByVal ID As Long)

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if NPC is walking, and if so process moving them over
    If Map.MapEvents(ID).Moving = 1 Then
        
        Select Case Map.MapEvents(ID).Dir
            Case DirectionUp
                Map.MapEvents(ID).yOffset = Map.MapEvents(ID).yOffset - ((ElapsedTime / 1000) * (Map.MapEvents(ID).MovementSpeed * SIZE_X))
                If Map.MapEvents(ID).yOffset < 0 Then Map.MapEvents(ID).yOffset = 0
                
            Case DirectionDown
                Map.MapEvents(ID).yOffset = Map.MapEvents(ID).yOffset + ((ElapsedTime / 1000) * (Map.MapEvents(ID).MovementSpeed * SIZE_X))
                If Map.MapEvents(ID).yOffset > 0 Then Map.MapEvents(ID).yOffset = 0
                
            Case DirectionLeft
                Map.MapEvents(ID).xOffset = Map.MapEvents(ID).xOffset - ((ElapsedTime / 1000) * (Map.MapEvents(ID).MovementSpeed * SIZE_X))
                If Map.MapEvents(ID).xOffset < 0 Then Map.MapEvents(ID).xOffset = 0
                
            Case DirectionRight
                Map.MapEvents(ID).xOffset = Map.MapEvents(ID).xOffset + ((ElapsedTime / 1000) * (Map.MapEvents(ID).MovementSpeed * SIZE_X))
                If Map.MapEvents(ID).xOffset > 0 Then Map.MapEvents(ID).xOffset = 0
                
        End Select
    
        ' Check if completed walking over to the next tile
        If Map.MapEvents(ID).Moving > 0 Then
            If Map.MapEvents(ID).Dir = DirectionRight Or Map.MapEvents(ID).Dir = DirectionDown Then
                If (Map.MapEvents(ID).xOffset >= 0) And (Map.MapEvents(ID).yOffset >= 0) Then
                    Map.MapEvents(ID).Moving = 0
                    If Map.MapEvents(ID).Step = 1 Then
                        Map.MapEvents(ID).Step = 3
                    Else
                        Map.MapEvents(ID).Step = 1
                    End If
                End If
            Else
                If (Map.MapEvents(ID).xOffset <= 0) And (Map.MapEvents(ID).yOffset <= 0) Then
                    Map.MapEvents(ID).Moving = 0
                    If Map.MapEvents(ID).Step = 1 Then
                        Map.MapEvents(ID).Step = 3
                    Else
                        Map.MapEvents(ID).Step = 1
                    End If
                End If
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ProcessNpcMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetColorString(color As Long)
    Select Case color
        Case 0
            GetColorString = "Black"
        Case 1
            GetColorString = "Blue"
        Case 2
            GetColorString = "Green"
        Case 3
            GetColorString = "Cyan"
        Case 4
            GetColorString = "Red"
        Case 5
            GetColorString = "Magenta"
        Case 6
            GetColorString = "Brown"
        Case 7
            GetColorString = "Grey"
        Case 8
            GetColorString = "Dark Grey"
        Case 9
            GetColorString = "Bright Blue"
        Case 10
            GetColorString = "Bright Green"
        Case 11
            GetColorString = "Bright Cyan"
        Case 12
            GetColorString = "Bright Red"
        Case 13
            GetColorString = "Pink"
        Case 14
            GetColorString = "Yellow"
        Case 15
            GetColorString = "White"

    End Select
End Function

Sub ClearEventChat()
    Dim i As Long
    If AnotherChat = 1 Then
        For i = 1 To 4
            frmMain.lblChoices(i).Visible = False
        Next
        
        frmMain.lblEventChat.Caption = ""
        frmMain.lblEventChatContinue.Visible = False
    ElseIf AnotherChat = 2 Then
        For i = 1 To 4
            frmMain.lblChoices(i).Visible = False
        Next
        
        frmMain.lblEventChat.Visible = False
        frmMain.lblEventChatContinue.Visible = False
        EventChatTimer = timeGetTime + 100
    Else
        frmMain.picEventChat.Visible = False
    End If

End Sub

Public Sub MenuLoop()

    ' If debug mode, handle error then exit out
    On Error GoTo errorhandler
restartmenuloop:
    ' *** Start GameLoop ***
    Do While Not InGame


        ' *********************
        ' ** Render Graphics **
        ' *********************
        Call DrawGDI
        DoEvents
    Loop

    ' Error handler
    Exit Sub
errorhandler:
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost
        GoTo restartmenuloop
    ElseIf options.Debug = 1 Then
        HandleError "MenuLoop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
    End If
End Sub

Sub ProcessWeather()
Dim i As Long
    If CurrentWeather > 0 Then
        i = Rand(1, 101 - CurrentWeatherIntensity)
        If i = 1 Then
            'Add a new particle
            For i = 1 To MAX_WEATHER_PARTICLES
                If WeatherParticle(i).InUse = False Then
                    If Rand(1, 2) = 1 Then
                        WeatherParticle(i).InUse = True
                        WeatherParticle(i).Type = CurrentWeather
                        WeatherParticle(i).Velocity = Rand(8, 14)
                        WeatherParticle(i).X = (TileView.Left * 32) - 32
                        WeatherParticle(i).Y = (TileView.Top * 32) + Rand(-32, frmMain.picScreen.ScaleHeight)
                    Else
                        WeatherParticle(i).InUse = True
                        WeatherParticle(i).Type = CurrentWeather
                        WeatherParticle(i).Velocity = Rand(10, 15)
                        WeatherParticle(i).X = (TileView.Left * 32) + Rand(-32, frmMain.picScreen.ScaleWidth)
                        WeatherParticle(i).Y = (TileView.Top * 32) - 32
                    End If
                    Exit For
                End If
            Next
        End If
    End If
    
    If CurrentWeather = WeatherStorm Then
        i = Rand(1, 400 - CurrentWeatherIntensity)
        If i = 1 Then
            'Draw Thunder
            DrawThunder = Rand(15, 22)
            PlaySound Sound_Thunder, -1, -1
        End If
    End If
    
    For i = 1 To MAX_WEATHER_PARTICLES
        If WeatherParticle(i).InUse Then
            If WeatherParticle(i).X > TileView.Right * 32 Or WeatherParticle(i).Y > TileView.Bottom * 32 Then
                WeatherParticle(i).InUse = False
            Else
                WeatherParticle(i).X = WeatherParticle(i).X + WeatherParticle(i).Velocity
                WeatherParticle(i).Y = WeatherParticle(i).Y + WeatherParticle(i).Velocity
            End If
        End If
    Next
End Sub

Public Sub AddChatBubble(ByVal target As Long, ByVal targetType As Byte, ByVal Msg As String, ByVal colour As Long)
Dim i As Long, Index As Long

    ' set the global index
    chatBubbleIndex = chatBubbleIndex + 1
    If chatBubbleIndex < 1 Or chatBubbleIndex > MAX_BYTE Then chatBubbleIndex = 1
    
    ' default to new bubble
    Index = chatBubbleIndex
    
    ' loop through and see if that player/npc already has a chat bubble
    For i = 1 To MAX_BYTE
        If chatBubble(i).targetType = targetType Then
            If chatBubble(i).target = target Then
                ' reset master index
                If chatBubbleIndex > 1 Then chatBubbleIndex = chatBubbleIndex - 1
                ' we use this one now, yes?
                Index = i
                Exit For
            End If
        End If
    Next
    
    ' set the bubble up
    With chatBubble(Index)
        .target = target
        .targetType = targetType
        .Msg = Msg
        .colour = colour
        .timer = timeGetTime
        
        .active = True
    End With
End Sub
