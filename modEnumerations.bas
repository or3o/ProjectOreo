Attribute VB_Name = "modEnumerations"
Option Explicit

' The order of the packets must match with the server's packet enumeration

' Packets sent by server to client
Public Enum ServerPackets
    SAlertMsg = 1
    SLoginOk
    SNewCharClasses
    SClassesData
    SInGame
    SPlayerInv
    SPlayerInvUpdate
    SPlayerWornEq
    SPlayerHp
    SPlayerMp
    SPlayerStats
    SPlayerData
    SPlayerMove
    SNpcMove
    SPlayerDir
    SNpcDir
    SPlayerXY
    SPlayerXYMap
    SAttack
    SNpcAttack
    SCheckForMap
    SMapData
    SMapItemData
    SMapNpcData
    SMapDone
    SGlobalMsg
    SAdminMsg
    SPlayerMsg
    SMapMsg
    SSpawnItem
    SItemEditor
    SUpdateItem
    SREditor
    SSpawnNpc
    SNpcDead
    SNpcEditor
    SUpdateNpc
    SMapKey
    SEditMap
    SShopEditor
    SUpdateShop
    SSpellEditor
    SUpdateSpell
    SSpells
    SLeft
    SResourceCache
    SResourceEditor
    SUpdateResource
    SSendPing
    SDoorAnimation
    SActionMsg
    SPlayerEXP
    SBlood
    SAnimationEditor
    SUpdateAnimation
    SAnimation
    SMapNpcVitals
    SCooldown
    SClearSpellBuffer
    SSayMsg
    SOpenShop
    SResetShopAction
    SStunned
    SMapWornEq
    SBank
    STrade
    SCloseTrade
    STradeUpdate
    STradeStatus
    STarget
    SHotbar
    SHighIndex
    SSound
    STradeRequest
    SPartyInvite
    SPartyUpdate
    SPartyVitals
    SQuestEditor
    SUpdateQuest
    SPlayerQuest
    SQuestMessage
    SSpawnEvent
    SEventMove
    SEventDir
    SEventChat
    SEventStart
    SEventEnd
    SPlayBGM
    SPlaySound
    SFadeoutBGM
    SStopSound
    SSwitchesAndVariables
    SMapEventData
    SChatBubble
    SSpecialEffect
    SHandleProjectile
    SCraft
    ' Make sure SMSG_COUNT is below everything else
    SMSG_COUNT
End Enum

' Packets sent by client to server
Public Enum ClientPackets
    CNewAccount = 1
    CDelAccount
    CLogin
    CAddChar
    CUseChar
    CSayMsg
    CEmoteMsg
    CBroadcastMsg
    CPlayerMsg
    CPlayerMove
    CPlayerDir
    CUseItem
    CAttack
    CUseStatPoint
    CPlayerInfoRequest
    CWarpMeTo
    CWarpToMe
    CWarpTo
    CSetSprite
    CGetStats
    CRequestNewMap
    CMapData
    CNeedMap
    CMapGetItem
    CMapDropItem
    CMapRespawn
    CMapReport
    CKickPlayer
    CBanList
    CBanDestroy
    CBanPlayer
    CRequestEditMap
    CRequestEditItem
    CSaveItem
    CRequestEditNpc
    CSaveNpc
    CRequestEditShop
    CSaveShop
    CRequestEditSpell
    CSaveSpell
    CSetAccess
    CWhosOnline
    CSetMotd
    CSearch
    CSpells
    CCast
    CQuit
    CSwapInvSlots
    CRequestEditResource
    CSaveResource
    CCheckPing
    CUnequip
    CRequestPlayerData
    CRequestItems
    CRequestNPCS
    CRequestResources
    CSpawnItem
    CRequestEditAnimation
    CSaveAnimation
    CRequestAnimations
    CRequestSpells
    CRequestShops
    CRequestLevelUp
    CForgetSpell
    CCloseShop
    CBuyItem
    CSellItem
    CChangeBankSlots
    CDepositItem
    CWithdrawItem
    CCloseBank
    CAdminWarp
    CTradeRequest
    CAcceptTrade
    CDeclineTrade
    CTradeItem
    CUntradeItem
    CHotbarChange
    CHotbarUse
    CSwapSpellSlots
    CAcceptTradeRequest
    CDeclineTradeRequest
    CPartyRequest
    CAcceptParty
    CDeclineParty
    CPartyLeave
    CRequestEditQuest
    CSaveQuest
    CRequestQuests
    CPlayerHandleQuest
    CQuestLogUpdate
    CEventChatReply
    CEvent
    CSwitchesAndVariables
    CRequestSwitchesAndVariables
    CProjecTileAttack
    CCraft
    ' Make sure CMSG_COUNT is below everything else
    CMSG_COUNT
End Enum

Public HandleDataSub(SMSG_COUNT) As Long

' Stats used by Players, Npcs and Classes
Public Enum Stats
    Strength = 1
    Endurance
    Intelligence
    Agility
    Willpower
    CRIT
    ' Make sure Stat_Count is below everything else
    Stat_Count
End Enum

' Vitals used by Players, Npcs and Classes
Public Enum Vitals
    HP = 1
    MP
    ' Make sure Vital_Count is below everything else
    Vital_Count
End Enum

' Equipment used by Players
Public Enum Equipment
    Enchant = 1
    Helmet
    Ring
    Weapon
    Armor
    Shield
    Glove
    Legs
    Boots
    ' Make sure Equipment_Count is below everything else
    Equipment_Count
End Enum

' Layers in a map
Public Enum MapLayer
    Ground = 1
    Mask
    Mask2
    Fringe
    Fringe2
    ' Make sure Layer_Count is below everything else
    Layer_Count
End Enum

' Sound entities
Public Enum SoundEntity
    seAnimation = 1
    seItem
    seNpc
    seResource
    seSpell
    ' Make sure SoundEntity_Count is below everything else
    SoundEntity_Count
End Enum
' Tile Types
Public Enum TileType
    TileWalkable = 0
    TileBlocked
    TileWarp
    TileItem
    TileNPCAvoid
    TileKey
    TileKeyOPEN
    TileResource
    TileDoor
    TileNPCSpawn
    TileShop
    TileBank
    TileHeal
    TileTrap
    TileSlide
    TileSound
    TileCraft
    TileWater
End Enum

'Weather Types
Public Enum WeatherTypes
    WeatherNone = 0
    WeatherRain
    WeatherSnow
    WeatherHail
    WeatherSandStorm
    WeatherStorm
End Enum

' Player Gender
Public Enum Gender
    Male = 0
    Female
End Enum

' Item Types
Public Enum ItemType
    ItemNone = 0
    ItemWeapon
    ItemArmor
    ItemHelmet
    ItemLegs
    ItemBoots
    ItemGlove
    ItemRing
    ItemEnchant
    ItemShield
    ItemConsume
    ItemKey
    ItemCurrency
    ItemSpell
    ItemRecipe
End Enum

' Map Morals
Public Enum MapMoral
    MoralNone = 0
    MoralSafe
End Enum

' Chat Log Enumerations
Public Enum ChatLog
    ChatGlobal = 0
    ChatMap
    ChatEmote
    ChatPlayer
    ChatSystem
End Enum

' Color Enumerations
Public Enum Colors
    Black = 0
    Blue
    Green
    Cyan
    Red
    Magenta
    Brown
    Grey
    DarkGrey
    BrightBlue
    BrightGreen
    BrightCyan
    BrightRed
    Pink
    Yellow
    White
    DarkBrown
    Orange
End Enum

' Directions Enumeration
Public Enum Directions
    DirectionUp = 0
    DirectionDown
    DirectionLeft
    DirectionRight
End Enum

' Enumeration for player movement
Public Enum PlayerMoveState
    PlayerWalking = 1
    PlayerRunning
End Enum

' Admin Enumeration
Public Enum PlayerRanks
    RankNormal = 0
    RankModerator
    RankMapper
    RankDeveloper
    RankAdministrator
End Enum

' NPC Behaviour Enumeration
Public Enum NPCAI
    AIAttackOnSight = 0
    AIAttackWhenAttacked
    AIFriendly
    AIImmobile
    AIAssist
End Enum

'Buff Types
Public Enum BuffType
     BUFFNONE = 0
     BUFFADDHP
     BUFFADDMP
     BUFFADDSTR
     BUFFADDEND
     BUFFADDAGI
     BUFFADDINT
     BUFFADDWILL
     BUFFADDATK
     BUFFADDDEF
     BUFFSUBHP
     BUFFSUBMP
     BUFFSUBSTR
     BUFFSUBEND
     BUFFSUBAGI
     BUFFSUBINT
     BUFFSUBWILL
     BUFFSUBATK
     BUFFSUBDEF
End Enum
' Spell Enumeration
Public Enum SpellType
    SpellDamageHP = 0
    SpellDamageMP
    SpellHealHP
    SpellHealMP
    SpellWarp
    SpellBuff
End Enum

'Game Editors
Public Enum GameEditors
    GEditorItem = 1
    GEditorNPC
    GEditorSpell
    GEditorShop
    GEditorResource
    GEditorAnimation
End Enum

' Target type Enumeration
Public Enum targetType
    TargetNone = 0
    TargetPlayer
    TargetNPC
    TargetEvent
End Enum

' Scrolling action message Enumeration
Public Enum ActionMessages
    MsgStatic = 0
    MsgScroll
    MsgScreen
End Enum

'Dialogue Types
Public Enum DialogueTypes
    DialogueNone = 0
    DialogueTrade
    DialogueForget
    DialogueParty
End Enum

'Move Routes
Public Enum MoveRouteOpts
    MoveUp = 1
    MoveDown
    MoveLeft
    MoveRight
    MoveRandom
    MoveTowardsPlayer
    MoveAwayFromPlayer
    StepForward
    StepBack
    Wait100ms
    Wait500ms
    Wait1000ms
    TurnUp
    TurnDown
    TurnLeft
    TurnRight
    Turn90Right
    Turn90Left
    Turn180
    TurnRandom
    TurnTowardPlayer
    TurnAwayFromPlayer
    SetSpeed8xSlower
    SetSpeed4xSlower
    SetSpeed2xSlower
    SetSpeedNormal
    SetSpeed2xFaster
    SetSpeed4xFaster
    SetFreqLowest
    SetFreqLower
    SetFreqNormal
    SetFreqHigher
    SetFreqHighest
    WalkingAnimOn
    WalkingAnimOff
    DirFixOn
    DirFixOff
    WalkThroughOn
    WalkThroughOff
    PositionBelowPlayer
    PositionWithPlayer
    PositionAbovePlayer
    ChangeGraphic
End Enum

' Event Types
Public Enum EventType
    ' Message
    evAddText = 1
    evShowText
    evShowChoices
    ' Game Progression
    evPlayerVar
    evPlayerSwitch
    evSelfSwitch
    ' Flow Control
    evCondition
    evExitProcess
    ' Player
    evChangeItems
    evRestoreHP
    evRestoreMP
    evLevelUp
    evChangeLevel
    evChangeSkills
    evChangeClass
    evChangeSprite
    evChangeSex
    evChangePK
    ' Movement
    evWarpPlayer
    evSetMoveRoute
    ' Character
    evPlayAnimation
    ' Music and Sounds
    evPlayBGM
    evFadeoutBGM
    evPlaySound
    evStopSound
    'Etc...
    evCustomScript
    evSetAccess
    'Shop/Bank
    evOpenBank
    evOpenShop
    'New
    evGiveExp
    evShowChatBubble
    evLabel
    evGotoLabel
    evSpawnNpc
    evFadeIn
    evFadeOut
    evFlashWhite
    evSetFog
    evSetWeather
    evSetTint
    evWait
End Enum
'Autotile Layering
Public Enum Autotiles
    ATInner = 0
    ATOuter
    ATHorizontal
    ATVertical
    ATFill
End Enum

'Autotile Types
Public Enum AutotileTypes
    ATNone = 0
    ATNormal
    ATFake
    ATAnim
    ATCliff
    ATWaterfall
End Enum

'Rendering Types
Public Enum RenderType
    RenderNone = 0
    RenderNormal
    RenderAutotile
End Enum

'Effect Types
Public Enum EffectTypes
    EffectFadeIn = 0
    EffectFadeOut
    EffectFlash
    EffectFog
    EffectWeather
    EffectTint
End Enum

' Menu States
Public Enum MenuStates
    StateNewAccount = 0
    StateDelAccount
    StateLogin
    StateGetChars
    StateNewChar
    StateAddChar
    StateDelChar
    StateUseChar
    StateInit
End Enum

