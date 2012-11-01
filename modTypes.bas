Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map As MapRec
Public Bank As BankRec
Public TempTile() As TempTileRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public NPC(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Switches(1 To MAX_SWITCHES) As String
Public Variables(1 To MAX_VARIABLES) As String
Public MapSounds() As MapSoundRec
Public MapSoundCount As Long
Public WeatherParticle(1 To MAX_WEATHER_PARTICLES) As WeatherParticleRec
Public Autotile() As AutotileRec
Public AEditor As PlayerRec

' client-side stuff
Public ActionMsg(1 To MAX_BYTE) As ActionMsgRec
Public Blood(1 To MAX_BYTE) As BloodRec
Public AnimInstance(1 To MAX_BYTE) As AnimInstanceRec
Public MenuButton(1 To MAX_MENUBUTTONS) As ButtonRec
Public MainButton(1 To MAX_MAINBUTTONS) As ButtonRec
Public Party As PartyRec

' options
Public options As OptionsRec

' Type recs
Private Type OptionsRec
    Game_Name As String
    SavePass As Byte
    Password As String * NAME_LENGTH
    Username As String * ACCOUNT_LENGTH
    IP As String
    Port As Long
    MenuMusic As String
    Music As Byte
    sound As Byte
    Debug As Byte
    HQ As Byte
End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Public Type PlayerInvRec
    Num As Long
    Value As Long
End Type

Private Type BankRec
    Item(1 To MAX_BANK) As PlayerInvRec
End Type

Private Type SpellAnim
    spellnum As Long
    timer As Long
    FramePointer As Long
End Type

Public Type ProjectileRec
    TravelTime As Long
    direction As Long
    X As Long
    Y As Long
    Pic As Long
    Range As Long
    Damage As Long
    speed As Long
End Type

Private Type PlayerRec
    ' General
    Name As String
    Class As Byte
    Sprite As Integer
    Spriteold As Integer
    Level As Byte
    EXP As Long
    Access As Byte
    PK As Byte
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    MaxVital(1 To Vitals.Vital_Count - 1) As Long
    ' Stats
    Stat(1 To Stats.Stat_Count - 1) As Byte
    POINTS As Long
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Long
    ' Position
    Map As Integer
    X As Byte
    Y As Byte
    Dir As Byte
    ' Client use only
    xOffset As Integer
    yOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    Step As Byte
    PlayerQuest(1 To MAX_QUESTS) As PlayerQuestRec
    EventTimer As Long
    ProjecTile(1 To MAX_PLAYER_PROJECTILES) As ProjectileRec
        ' Proficiencies
    Swords As Byte
    SwordsExp As Long
    Axes As Byte
    AxesExp As Long
    Daggers As Byte
    DaggersExp As Long
    Range As Byte
    RangeExp As Long
    Magic As Byte
    MagicExp As Long
    Conviction As Byte
    ConvictionExp As Long
    Lycanthropy As Byte
    LycanthropyExp As Long
    Heavyarmor As Byte
    HeavyArmorExp As Long
    LightArmor As Byte
    LightArmorExp As Long
    Mining As Byte
    MiningExp As Long
    Woodcutting As Byte
    WoodcuttingExp As Long
    Fishing As Byte
    FishingExp As Long
    Crafting As Byte
    CraftingExp As Long
    WieldDagger As Byte
    Craft As Byte
    Smith As Byte
    SmithExp As Long
    Alchemy As Byte
    AlchemyExp As Long
    Enchants As Byte
    EnchantsExp As Long
    IsMember As Byte
    DateCount As String
End Type

Private Type TileDataRec
    X As Long
    Y As Long
    Tileset As Long
End Type

Public Type ConditionalBranchRec
    Condition As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    CommandList As Long
    ElseCommandList As Long
End Type

Public Type MoveRouteRec
    Index As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    Data5 As Long
    Data6 As Long
End Type

Public Type EventCommandRec
    Index As Long
    Text1 As String
    Text2 As String
    Text3 As String
    Text4 As String
    Text5 As String
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    Data5 As Long
    Data6 As Long
    ConditionalBranch As ConditionalBranchRec
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
End Type

Public Type CommandListRec
    CommandCount As Long
    ParentList As Long
    Commands() As EventCommandRec
End Type

Public Type EventPageRec
    'These are condition variables that decide if the event even appears to the player.
    chkVariable As Long
    VariableIndex As Long
    VariableCondition As Long
    VariableCompare As Long
    
    chkSwitch As Long
    SwitchIndex As Long
    SwitchCompare As Long
    
    chkHasItem As Long
    HasItemIndex As Long
    
    chkSelfSwitch As Long
    SelfSwitchIndex As Long
    SelfSwitchCompare As Long
    'End Conditions
    
    'Handles the Event Sprite
    GraphicType As Byte
    Graphic As Long
    GraphicX As Long
    GraphicY As Long
    GraphicX2 As Long
    GraphicY2 As Long
    
    'Handles Movement - Move Routes to come soon.
    MoveType As Byte
    MoveSpeed As Byte
    MoveFreq As Byte
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
    IgnoreMoveRoute As Long
    RepeatMoveRoute As Long
    
    'Guidelines for the event
    WalkAnim As Byte
    DirFix As Byte
    WalkThrough As Byte
    ShowName As Byte
    
    'Trigger for the event
    Trigger As Byte
    
    'Commands for the event
    CommandListCount As Long
    CommandList() As CommandListRec
    
    Position As Byte
    
    'Client Needed Only
    X As Long
    Y As Long
End Type

Public Type EventRec
    Name As String
    Global As Long
    pageCount As Long
    Pages() As EventPageRec
    X As Long
    Y As Long
End Type

Public Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Autotile(1 To MapLayer.Layer_Count - 1) As Byte
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As String
    DirBlock As Byte

End Type

Private Type MapEventRec
    Name As String
    Dir As Long
    X As Long
    Y As Long
    GraphicType As Long
    GraphicX As Long
    GraphicY As Long
    GraphicX2 As Long
    GraphicY2 As Long
    GraphicNum As Long
    Moving As Long
    MovementSpeed As Long
    Position As Long
    xOffset As Long
    yOffset As Long
    Step As Long
    Visible As Long
    WalkAnim As Long
    DirFix As Long
    ShowDir As Long
    WalkThrough As Long
    ShowName As Long
End Type

Private Type MapRec
    Name As String * NAME_LENGTH
    Music As String * MUSIC_LENGTH
    BGS As String * MUSIC_LENGTH
    
    Revision As Long
    Moral As Byte
    
    Up As Long
    Down As Long
    Left As Long
    Right As Long
    
    BootMap As Long
    BootX As Byte
    BootY As Byte
    
    Weather As Long
    WeatherIntensity As Long
    
    Fog As Long
    FogSpeed As Long
    FogOpacity As Long
    
    Red As Long
    Green As Long
    Blue As Long
    Alpha As Long
    
    MaxX As Byte
    MaxY As Byte
    
    Tile() As TileRec
    NPC(1 To MAX_MAP_NPCS) As Long
    NpcSpawnType(1 To MAX_MAP_NPCS) As Long
    EventCount As Long
    Events() As EventRec

    'Client Side Only -- Temporary
    CurrentEvents As Long
    MapEvents() As MapEventRec
End Type

Private Type ClassRec
    Name As String * NAME_LENGTH
    Stat(1 To Stats.Stat_Count - 1) As Byte
    MaleSprite() As Long
    FemaleSprite() As Long
    ' For client use
    Vital(1 To Vitals.Vital_Count - 1) As Long
End Type

Private Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    sound As String * NAME_LENGTH
    
    Pic As Integer
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    ClassReq As Byte
    AccessReq As Byte
    LevelReq As Byte
    Price As Long
    Add_Stat(1 To Stats.Stat_Count - 1) As Byte
    Rarity As Byte
    speed As Long
    BindType As Byte
    Stat_Req(1 To Stats.Stat_Count - 1) As Byte
    Animation As Long
    Paperdoll As Integer
    
    AddHP As Long
    AddMP As Long
    AddEXP As Long
    CastSpell As Long
    instaCast As Byte
    ' Proficiencies
    Swords As Boolean
    SwordsReq As Byte
    Axes As Boolean
    AxesReq As Byte
    Daggers As Boolean
    daggersReq As Byte
    Range As Boolean
    RangeReq As Byte
    Magic As Boolean
    magicReq As Byte
    Conviction As Boolean
    convictionReq As Byte
    Lycanthropy As Boolean
    lycanthropyReq As Byte
    Heavyarmor As Boolean
    HeavyArmorReq As Byte
    LightArmor As Boolean
    LightarmorReq As Byte
    Mining As Boolean
    MiningReq As Byte
    Woodcutting As Boolean
    WoodcuttingReq As Byte
    Fishing As Boolean
    FishingReq As Byte
    Crafting As Boolean
    CraftingReq As Byte
    Smithy As Boolean
    SmithReq As Byte
    Alchemist As Boolean
    AlchemyReq As Byte
    Enchanter As Boolean
    EnchantReq As Byte
    crftxp As Integer
    ProjecTile As ProjectileRec
    ammo As Byte
    ammoreq As Boolean
    isDagger As Boolean
    isTwohanded As Boolean
    Daggerpdoll As Integer
    Amount As Byte
    'crafting
    Tool As Byte
    ToolReq As Byte
    Recipe(1 To MAX_RECIPE_ITEMS) As Long
    IsMember As Byte
End Type

Private Type MapItemRec
    playerName As String
    Num As Long
    Value As Long
    Frame As Byte
    X As Byte
    Y As Byte
End Type

Private Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    sound As String * NAME_LENGTH
    
    Sprite As Long
    SpawnSecs As Long
    Behaviour As Byte
    Range As Byte
    DropChance(1 To MAX_NPC_DROPS) As Double
    DropItem(1 To MAX_NPC_DROPS) As Byte
    DropItemValue(1 To MAX_NPC_DROPS) As Integer
    Stat(1 To Stats.Stat_Count - 1) As Byte
    HP As Long
    EXP As Long
    Animation As Integer
    Damage As Integer
    Level As Byte
    Quest As Byte
    QuestNum As Long
    Skillexp As Long
End Type

Private Type MapNpcRec
    Num As Long
    target As Long
    targetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    Map As Long
    X As Byte
    Y As Byte
    Dir As Byte
    ' Client use only
    xOffset As Long
    yOffset As Long
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    Step As Byte
End Type

Private Type TradeItemRec
    Item As Integer
    ItemValue As Long
    CostItem As Long
    CostValue As Long
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

Private Type SpellRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    sound As String * NAME_LENGTH
    
    Type As Byte
    MPCost As Long
    LevelReq As Byte
    AccessReq As Byte
    ClassReq As Byte
    CastTime As Long
    CDTime As Long
    Icon As Long
    Map As Long
    X As Long
    Y As Long
    Dir As Byte
    Vital As Long
    Duration As Long
    Interval As Long
    Range As Byte
    IsAoE As Boolean
    AoE As Long
    CastAnim As Long
    SpellAnim As Long
    StunDuration As Long
    EXP As Integer
    Magic As Byte
    Lycanthropy As Byte
    Conviction As Byte
    BuffType As Byte
    Sprite As Byte
End Type

Private Type TempTileRec
    DoorOpen As Byte
    DoorFrame As Byte
    DoorTimer As Long
    DoorAnimate As Byte ' 0 = nothing| 1 = opening | 2 = closing
End Type

Public Type MapResourceRec
    X As Long
    Y As Long
    ResourceState As Byte
End Type

Private Type ResourceRec
    Name As String * NAME_LENGTH
    SuccessMessage As String * NAME_LENGTH
    EmptyMessage As String * NAME_LENGTH
    sound As String * NAME_LENGTH
    
    ResourceType As Byte
    ResourceImage As Integer
    ExhaustedImage As Integer
    ItemReward As Integer
    ToolRequired As Byte
    health As Long
    RespawnTime As Long
    WalkThrough As Boolean
    Animation As Integer
    EXP As Integer
End Type

Private Type ActionMsgRec
    Message As String
    Created As Long
    Type As Long
    color As Long
    Scroll As Long
    X As Long
    Y As Long
    timer As Long
End Type

Private Type BloodRec
    Sprite As Long
    timer As Long
    X As Long
    Y As Long
End Type

Private Type AnimationRec
    Name As String * NAME_LENGTH
    sound As String * NAME_LENGTH
    
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    looptime(0 To 1) As Long
End Type

Private Type AnimInstanceRec
    Animation As Long
    X As Long
    Y As Long
    ' used for locking to players/npcs
    lockindex As Long
    LockType As Byte
    ' timing
    timer(0 To 1) As Long
    ' rendering check
    Used(0 To 1) As Boolean
    ' counting the loop
    LoopIndex(0 To 1) As Long
    frameIndex(0 To 1) As Long
End Type

Public Type HotbarRec
    Slot As Long
    sType As Byte
End Type

Public Type ButtonRec
    FileName As String
    state As Byte
End Type

Public Type EventListRec
    CommandList As Long
    CommandNum As Long
End Type

Public Type MapSoundRec
    X As Long
    Y As Long
    SoundHandle As Long
    InUse As Boolean
    Channel As Long
End Type

Public Type WeatherParticleRec
    Type As Long
    X As Long
    Y As Long
    Velocity As Long
    InUse As Long
End Type

'Auto tiles "/
Public Type PointRec
    X As Long
    Y As Long
End Type

Public Type QuarterTileRec
    QuarterTile(1 To 4) As PointRec
    renderState As Byte
    srcX(1 To 4) As Long
    srcY(1 To 4) As Long
End Type

Public Type AutotileRec
    Layer(1 To MapLayer.Layer_Count - 1) As QuarterTileRec
End Type

Public Type ChatBubbleRec
    Msg As String
    colour As Long
    target As Long
    targetType As Byte
    timer As Long
    active As Boolean
End Type
