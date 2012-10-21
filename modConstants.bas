Attribute VB_Name = "modConstants"
Option Explicit

' API Declares
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal uDuration As Long)
' animated buttons
Public Const MAX_MENUBUTTONS As Byte = 4
Public Const MAX_MAINBUTTONS As Byte = 8
Public Const MENUBUTTON_PATH As String = "\Data Files\graphics\gui\menu\buttons\"
Public Const MAINBUTTON_PATH As String = "\Data Files\graphics\gui\main\buttons\"

' Hotbar
Public Const HotbarTop As Integer = 2
Public Const HotbarLeft As Integer = 2
Public Const HotbarOffsetX As Byte = 8

' Inventory constants
Public Const InvTop As Integer = 24
Public Const InvLeft As Integer = 12
Public Const InvOffsetY As Byte = 3
Public Const InvOffsetX As Byte = 3
Public Const InvColumns As Byte = 5

' Bank constants
Public Const BankTop As Integer = 38
Public Const BankLeft As Integer = 42
Public Const BankOffsetY As Byte = 3
Public Const BankOffsetX As Byte = 4
Public Const BankColumns As Byte = 11

' spells constants
Public Const SpellTop As Integer = 24
Public Const SpellLeft As Integer = 12
Public Const SpellOffsetY As Byte = 3
Public Const SpellOffsetX As Byte = 3
Public Const SpellColumns As Byte = 5

' shop constants
Public Const ShopTop As Integer = 6
Public Const ShopLeft As Integer = 8
Public Const ShopOffsetY As Byte = 2
Public Const ShopOffsetX As Byte = 4
Public Const ShopColumns As Byte = 5

' Character consts
Public Const EqTop As Integer = 224
Public Const EqLeft As Integer = 18
Public Const EqOffsetX As Byte = 10
Public Const EqOffsetY As Byte = 4
Public Const EqColumns As Byte = 4

' values
Public Const MAX_BYTE As Byte = 255
Public Const MAX_INTEGER As Integer = 32767
Public Const MAX_LONG As Long = 2147483647

' path constants
Public Const SOUND_PATH As String = "\Data Files\sound\"
Public Const MUSIC_PATH As String = "\Data Files\music\"

' Log Path and variables
Public Const LOG_DEBUG As String = "debug.txt"
Public Const LOG_PATH As String = "\Data Files\logs\"

' Map Path and variables
Public Const MAP_PATH As String = "\Data Files\maps\"
Public Const MAP_EXT As String = ".map"

' Gfx Path and variables
Public Const GFX_PATH As String = "\Data Files\graphics\"
Public Const GFX_EXT As String = ".png"

Public Const FONT_PATH As String = "\data files\graphics\fonts\"

' Key constants
Public Const VK_UP As Long = &H26
Public Const VK_DOWN As Long = &H28
Public Const VK_LEFT As Long = &H25
Public Const VK_RIGHT As Long = &H27
Public Const VK_SHIFT As Long = &H10
Public Const VK_RETURN As Long = &HD
Public Const VK_CONTROL As Long = &H11

' Speed moving vars
Public Const WALK_SPEED As Byte = 4
Public Const RUN_SPEED As Byte = 6

' Tile size constants
Public Const PIC_X As Long = 32
Public Const PIC_Y As Long = 32

' Sprite, item, spell size constants
Public Const SIZE_X As Byte = 32
Public Const SIZE_Y As Byte = 32

' ********************************************************
' * The values below must match with the server's values *
' ********************************************************

' General constants
Public Const MAX_PLAYERS As Byte = 70
Public Const MAX_ITEMS As Integer = 255
Public Const MAX_NPCS As Integer = 255
Public Const MAX_ANIMATIONS As Byte = 255
Public Const MAX_INV As Byte = 35
Public Const MAX_MAP_ITEMS As Byte = 255
Public Const MAX_MAP_NPCS As Byte = 30
Public Const MAX_SHOPS As Byte = 50
Public Const MAX_PLAYER_SPELLS As Byte = 35
Public Const MAX_SPELLS As Byte = 255
Public Const MAX_TRADES As Integer = 30
Public Const MAX_RESOURCES As Integer = 100
Public Const MAX_LEVELS As Byte = 100
Public Const MAX_BANK As Byte = 99
Public Const MAX_HOTBAR As Byte = 12
Public Const MAX_PARTYS As Byte = 35
Public Const MAX_PARTY_MEMBERS As Byte = 4
Public Const MAX_SWITCHES As Integer = 1000
Public Const MAX_VARIABLES As Integer = 1000
Public Const MAX_WEATHER_PARTICLES As Integer = 250
Public Const MAX_PLAYER_PROJECTILES As Byte = 20
Public Const MAX_NPC_DROPS As Byte = 10
Public Const MAX_RECIPE_ITEMS As Byte = 5
' Website
Public Const GAME_WEBSITE As String = "http://www.touchofdeathforums.com"

' Boolean constants
Public Const NO As Byte = 0
Public Const YES As Byte = 1

' String constants
Public Const NAME_LENGTH As Byte = 20
Public Const MUSIC_LENGTH As Byte = 40
Public Const ACCOUNT_LENGTH As Byte = 12

' Map constants
Public Const MAX_MAPS As Integer = 100
Public MAX_MAPX As Byte
Public MAX_MAPY As Byte

' Do Events
Public Const nLng As Long = (&H80 Or &H1 Or &H4 Or &H20) + (&H8 Or &H40)

' Scrolling action message constants
Public Const ACTIONMSG_STATIC As Long = 0
Public Const ACTIONMSG_SCROLL As Long = 1
Public Const ACTIONMSG_SCREEN As Long = 2

' stuffs
Public HalfX As Integer
Public HalfY As Integer
Public ScreenX As Integer
Public ScreenY As Integer
Public StartXValue As Integer
Public StartYValue As Integer
Public EndXValue As Integer
Public EndYValue As Integer
Public Const Half_PIC_X As Integer = PIC_X / 2
Public Const Half_PIC_Y As Integer = PIC_Y / 2

'Chatbubble
Public Const ChatBubbleWidth As Integer = 200

