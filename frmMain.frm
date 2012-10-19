VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCN.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   12765
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16260
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   851
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1084
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame FraLevels 
      Caption         =   "Levels"
      Height          =   4215
      Left            =   8040
      TabIndex        =   126
      Top             =   4200
      Visible         =   0   'False
      Width           =   3300
      Begin VB.PictureBox Picture23 
         Height          =   375
         Left            =   1680
         ScaleHeight     =   315
         ScaleWidth      =   555
         TabIndex        =   208
         Top             =   3000
         Width           =   615
      End
      Begin VB.PictureBox Picture12 
         Height          =   495
         Left            =   240
         Picture         =   "frmMain.frx":3332
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   173
         Top             =   3000
         Width           =   495
      End
      Begin VB.PictureBox Picture11 
         Height          =   495
         Left            =   2520
         Picture         =   "frmMain.frx":383C
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   172
         Top             =   2280
         Width           =   495
      End
      Begin VB.PictureBox Picture10 
         Height          =   495
         Left            =   1800
         Picture         =   "frmMain.frx":3CB5
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   171
         Top             =   2280
         Width           =   495
      End
      Begin VB.PictureBox Picture9 
         Height          =   495
         Left            =   1080
         Picture         =   "frmMain.frx":40F9
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   170
         Top             =   2280
         Width           =   495
      End
      Begin VB.PictureBox Picture7 
         Height          =   495
         Left            =   240
         Picture         =   "frmMain.frx":4564
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   168
         Top             =   2280
         Width           =   495
      End
      Begin VB.PictureBox Picture5 
         Height          =   495
         Left            =   2520
         Picture         =   "frmMain.frx":4AD5
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   164
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox piclyconthropy 
         Height          =   495
         Left            =   1800
         Picture         =   "frmMain.frx":5051
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   163
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         Height          =   495
         Left            =   1080
         Picture         =   "frmMain.frx":55F7
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   162
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox PicMagic 
         Height          =   495
         Left            =   240
         Picture         =   "frmMain.frx":5B8C
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   161
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox PicRange 
         Height          =   495
         Left            =   2520
         Picture         =   "frmMain.frx":603A
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   160
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Left            =   1800
         Picture         =   "frmMain.frx":658F
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   159
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox picaxes 
         Height          =   495
         Left            =   1080
         Picture         =   "frmMain.frx":6961
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   158
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox picswords 
         Height          =   495
         Left            =   240
         Picture         =   "frmMain.frx":6E4C
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   157
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblEnchants 
         Caption         =   "EN"
         Height          =   255
         Left            =   2520
         TabIndex        =   207
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label lblEnchantsExp 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   206
         Top             =   3840
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblAlchemyExp 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   205
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblSmithExp 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   204
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblAlchemy 
         Caption         =   "AL"
         Height          =   255
         Left            =   1800
         TabIndex        =   203
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label lblSmith 
         Caption         =   "Sm"
         Height          =   255
         Left            =   1080
         TabIndex        =   202
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label lblCraftingExp 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   156
         Top             =   3840
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label lblmagic 
         BackColor       =   &H00000000&
         Caption         =   "MA"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   155
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label lblFishingExp 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   154
         Top             =   3840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblWoodcuttingExp 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   153
         Top             =   3840
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblMiningExp 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   152
         Top             =   3840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblLightArmorExp 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   151
         Top             =   3840
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lbldaggers 
         BackColor       =   &H00000000&
         Caption         =   "DA"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   150
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblCrafting 
         BackStyle       =   0  'Transparent
         Caption         =   "Crafting"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   149
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label lblFishing 
         BackStyle       =   0  'Transparent
         Caption         =   "Fishing"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   148
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label lblWoodcutting 
         BackStyle       =   0  'Transparent
         Caption         =   "WC"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   147
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label lblHeavyarmorExp 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   146
         Top             =   3840
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblMining 
         BackStyle       =   0  'Transparent
         Caption         =   "MI"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   145
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label lblLycanthropyExp 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   144
         Top             =   3840
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblLightArmor 
         BackStyle       =   0  'Transparent
         Caption         =   "LA"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   143
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label lblConvictionExp 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   142
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblHeavyarmor 
         BackStyle       =   0  'Transparent
         Caption         =   "HA"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   141
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblMagicExp 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   140
         Top             =   3840
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblLycanthropy 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "LY"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   139
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label lblRangeExp 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   138
         Top             =   3840
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblConviction 
         BackColor       =   &H00000000&
         Caption         =   "CO"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1080
         TabIndex        =   137
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label lblDaggersExp 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   136
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblRange 
         BackColor       =   &H00000000&
         Caption         =   "RA"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   135
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lblAxesExp 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   134
         Top             =   3840
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblAxes 
         BackColor       =   &H00000000&
         Caption         =   "AX"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   133
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblswords 
         BackColor       =   &H00000000&
         Caption         =   "sw"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   127
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblswordsexp 
         BackColor       =   &H00000000&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   128
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Image Image1 
         Height          =   4815
         Left            =   0
         Picture         =   "frmMain.frx":72CB
         Top             =   0
         Width           =   3300
      End
   End
   Begin VB.PictureBox picAdmin 
      Appearance      =   0  'Flat
      BackColor       =   &H00B5B5B5&
      ForeColor       =   &H80000008&
      Height          =   7530
      Left            =   12000
      ScaleHeight     =   500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   189
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   2865
      Begin VB.CommandButton cmdLevel 
         Caption         =   "Level Up"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   7200
         Width           =   2295
      End
      Begin VB.CommandButton cmdAAnim 
         Caption         =   "Animation"
         Height          =   255
         Left            =   1440
         TabIndex        =   52
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdAAccess 
         Caption         =   "Set Access"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txtAAccess 
         Height          =   285
         Left            =   1440
         TabIndex        =   44
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtASprite 
         Height          =   285
         Left            =   2160
         TabIndex        =   42
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton cmdARespawn 
         Caption         =   "Respawn"
         Height          =   255
         Left            =   1440
         TabIndex        =   41
         Top             =   5040
         Width           =   1095
      End
      Begin VB.CommandButton cmdASprite 
         Caption         =   "Set Sprite"
         Height          =   255
         Left            =   1440
         TabIndex        =   40
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdASpawn 
         Caption         =   "Spawn Item"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   6720
         Width           =   2295
      End
      Begin VB.HScrollBar scrlAAmount 
         Height          =   255
         Left            =   240
         Min             =   1
         TabIndex        =   38
         Top             =   6360
         Value           =   1
         Width           =   2295
      End
      Begin VB.HScrollBar scrlAItem 
         Height          =   255
         Left            =   240
         Min             =   1
         TabIndex        =   36
         Top             =   5760
         Value           =   1
         Width           =   2295
      End
      Begin VB.CommandButton cmdASpell 
         Caption         =   "Spell"
         Height          =   255
         Left            =   1440
         TabIndex        =   34
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdAShop 
         Caption         =   "Shop"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdAResource 
         Caption         =   "Resource"
         Height          =   255
         Left            =   1440
         TabIndex        =   32
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdANpc 
         Caption         =   "NPC"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdAMap 
         Caption         =   "Map"
         Height          =   255
         Left            =   1440
         TabIndex        =   30
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton cmdAItem 
         Caption         =   "Item"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdADestroy 
         Caption         =   "Del Bans"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   5040
         Width           =   1095
      End
      Begin VB.CommandButton cmdAMapReport 
         Caption         =   "Map Report"
         Height          =   255
         Left            =   1440
         TabIndex        =   26
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton cmdALoc 
         Caption         =   "Loc"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton cmdAWarp 
         Caption         =   "Warp To"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtAMap 
         Height          =   285
         Left            =   960
         TabIndex        =   22
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton cmdAWarpMe2 
         Caption         =   "WarpMe2"
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdAWarp2Me 
         Caption         =   "Warp2Me"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdABan 
         Caption         =   "Ban"
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdAKick 
         Caption         =   "Kick"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtAName 
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   1095
      End
      Begin VB.Line Line5 
         X1              =   16
         X2              =   168
         Y1              =   472
         Y2              =   472
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Access:"
         Height          =   255
         Left            =   1440
         TabIndex        =   45
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite#:"
         Height          =   255
         Left            =   1440
         TabIndex        =   43
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblAAmount 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount: 1"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   6120
         Width           =   2295
      End
      Begin VB.Label lblAItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Spawn Item: None"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   5520
         Width           =   2295
      End
      Begin VB.Line Line4 
         X1              =   16
         X2              =   168
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line3 
         X1              =   16
         X2              =   168
         Y1              =   304
         Y2              =   304
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Editors:"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Line Line2 
         X1              =   16
         X2              =   168
         Y1              =   200
         Y2              =   200
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Map#:"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   16
         X2              =   168
         Y1              =   144
         Y2              =   144
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Admin Panel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   120
         Width           =   2865
      End
   End
   Begin VB.PictureBox picOptions 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8115
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   63
      Top             =   4245
      Visible         =   0   'False
      Width           =   2910
      Begin VB.OptionButton optdet 
         BackColor       =   &H00000000&
         Caption         =   "low detail"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   131
         Top             =   2400
         Width           =   975
      End
      Begin VB.OptionButton scrnopt 
         BackColor       =   &H00000000&
         Caption         =   "high detail"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1200
         MaskColor       =   &H00000000&
         TabIndex        =   130
         Top             =   2400
         Width           =   975
      End
      Begin VB.ComboBox comboresolution 
         Height          =   315
         ItemData        =   "frmMain.frx":BFCE
         Left            =   1440
         List            =   "frmMain.frx":BFD0
         TabIndex        =   129
         Text            =   "Combo1"
         Top             =   3600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   1935
         TabIndex        =   69
         Top             =   1440
         Width           =   1935
         Begin VB.OptionButton optSOn 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "On"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   71
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton optSOff 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Off"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   720
            TabIndex        =   70
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   1935
         TabIndex        =   66
         Top             =   840
         Width           =   1935
         Begin VB.OptionButton optMOff 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Off"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   720
            TabIndex        =   68
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton optMOn 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "On"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   67
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Screen Size"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   132
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sound"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   65
         Top             =   1200
         Width           =   600
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Music"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   64
         Top             =   600
         Width           =   555
      End
   End
   Begin VB.PictureBox picCharacter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8160
      Picture         =   "frmMain.frx":BFD2
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   5
      Top             =   4260
      Visible         =   0   'False
      Width           =   2910
      Begin VB.PictureBox picFace 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   840
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   85
         Top             =   5160
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   198
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   90
         Index           =   6
         Left            =   2400
         TabIndex        =   197
         Top             =   4080
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   6
         Left            =   960
         TabIndex        =   196
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label lblPoints 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   2280
         TabIndex        =   93
         Top             =   2160
         Width           =   120
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   1200
         TabIndex        =   51
         Top             =   2160
         Width           =   105
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   5
         Left            =   2400
         TabIndex        =   50
         Top             =   1920
         Width           =   105
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   1200
         TabIndex        =   49
         Top             =   1920
         Width           =   105
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   4
         Left            =   2400
         TabIndex        =   48
         Top             =   1680
         Width           =   105
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   1200
         TabIndex        =   47
         Top             =   1680
         Width           =   105
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   1080
         TabIndex        =   13
         Top             =   2160
         Width           =   120
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   5
         Left            =   2280
         TabIndex        =   12
         Top             =   1920
         Width           =   120
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   1080
         TabIndex        =   11
         Top             =   1920
         Width           =   120
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   4
         Left            =   2280
         TabIndex        =   10
         Top             =   1680
         Width           =   120
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   1080
         TabIndex        =   9
         Top             =   1680
         Width           =   120
      End
      Begin VB.Label lblCharName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empty"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   495
         Width           =   2640
      End
   End
   Begin VB.PictureBox picInventory 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3930
      Left            =   8160
      Picture         =   "frmMain.frx":FB26
      ScaleHeight     =   262
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   3
      Top             =   4260
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox Picture22 
      Height          =   495
      Left            =   1560
      Picture         =   "frmMain.frx":14F4E
      ScaleHeight     =   435
      ScaleWidth      =   1035
      TabIndex        =   184
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Fracrafting 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   1440
      TabIndex        =   165
      Top             =   1200
      Visible         =   0   'False
      Width           =   4815
      Begin VB.PictureBox Picture8 
         Height          =   495
         Left            =   4320
         Picture         =   "frmMain.frx":17825
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   169
         Top             =   0
         Width           =   495
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4095
         Left            =   0
         TabIndex        =   166
         Top             =   0
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   7223
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         BackColor       =   0
         ForeColor       =   65280
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Bronze  Items"
         TabPicture(0)   =   "frmMain.frx":19BD7
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Picture21"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Picture6"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Picture13"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Picture14"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Picture15"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Picture16"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Picture17"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Picture18"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Picture19"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).ControlCount=   9
         TabCaption(1)   =   "Tab 1"
         TabPicture(1)   =   "frmMain.frx":19BF3
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Picture20"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Tab 2"
         TabPicture(2)   =   "frmMain.frx":19C0F
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         Begin VB.PictureBox Picture20 
            Height          =   3255
            Left            =   -75000
            Picture         =   "frmMain.frx":19C2B
            ScaleHeight     =   3195
            ScaleWidth      =   4755
            TabIndex        =   181
            Top             =   360
            Width           =   4815
         End
         Begin VB.PictureBox Picture19 
            Height          =   495
            Left            =   3600
            ScaleHeight     =   435
            ScaleWidth      =   555
            TabIndex        =   180
            Top             =   2280
            Width           =   615
         End
         Begin VB.PictureBox Picture18 
            Height          =   495
            Left            =   2280
            ScaleHeight     =   435
            ScaleWidth      =   555
            TabIndex        =   179
            Top             =   2280
            Width           =   615
         End
         Begin VB.PictureBox Picture17 
            Height          =   495
            Left            =   1200
            ScaleHeight     =   435
            ScaleWidth      =   555
            TabIndex        =   178
            Top             =   2280
            Width           =   615
         End
         Begin VB.PictureBox Picture16 
            Height          =   495
            Left            =   240
            ScaleHeight     =   435
            ScaleWidth      =   435
            TabIndex        =   177
            Top             =   2280
            Width           =   495
         End
         Begin VB.PictureBox Picture15 
            Height          =   495
            Left            =   3600
            ScaleHeight     =   435
            ScaleWidth      =   435
            TabIndex        =   176
            Top             =   960
            Width           =   495
         End
         Begin VB.PictureBox Picture14 
            Height          =   495
            Left            =   2400
            Picture         =   "frmMain.frx":200FE
            ScaleHeight     =   435
            ScaleWidth      =   435
            TabIndex        =   175
            Top             =   960
            Width           =   495
         End
         Begin VB.PictureBox Picture13 
            Height          =   495
            Left            =   1320
            Picture         =   "frmMain.frx":20D40
            ScaleHeight     =   435
            ScaleWidth      =   435
            TabIndex        =   174
            Top             =   960
            Width           =   495
         End
         Begin VB.PictureBox Picture6 
            Height          =   495
            Left            =   240
            Picture         =   "frmMain.frx":21982
            ScaleHeight     =   435
            ScaleWidth      =   435
            TabIndex        =   167
            Top             =   960
            Width           =   495
         End
         Begin VB.PictureBox Picture21 
            Height          =   3255
            Left            =   0
            Picture         =   "frmMain.frx":225C4
            ScaleHeight     =   3195
            ScaleWidth      =   4755
            TabIndex        =   182
            Top             =   360
            Width           =   4815
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Lvl Req:"
               ForeColor       =   &H00FFFFFF&
               Height          =   495
               Left            =   3480
               TabIndex        =   201
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Lvl Req:"
               ForeColor       =   &H00FFFFFF&
               Height          =   495
               Left            =   1200
               TabIndex        =   200
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Lvl Req:"
               ForeColor       =   &H00FFFFFF&
               Height          =   495
               Left            =   2280
               TabIndex        =   199
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Lvl Req:"
               ForeColor       =   &H00FFFFFF&
               Height          =   495
               Left            =   120
               TabIndex        =   183
               Top             =   1080
               Width           =   735
            End
         End
      End
   End
   Begin VB.PictureBox picParty 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8115
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   102
      Top             =   4245
      Visible         =   0   'False
      Width           =   2910
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   4
         Left            =   90
         Top             =   3075
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   4
         Left            =   90
         Top             =   2940
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   3
         Left            =   90
         Top             =   2340
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   3
         Left            =   90
         Top             =   2205
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   2
         Left            =   90
         Top             =   1620
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   2
         Left            =   90
         Top             =   1485
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   1
         Left            =   90
         Top             =   870
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   1
         Left            =   90
         Top             =   735
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Label lblPartyLeave 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         TabIndex        =   108
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label lblPartyInvite 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   107
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   106
         Top             =   2670
         Width           =   2415
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   105
         Top             =   1935
         Width           =   2415
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   104
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   103
         Top             =   465
         Width           =   2415
      End
   End
   Begin VB.PictureBox picQuestLog 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8115
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   116
      Top             =   4245
      Visible         =   0   'False
      Width           =   2910
      Begin VB.TextBox txtQuestTaskLog 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   975
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   118
         Top             =   1080
         Width           =   2415
      End
      Begin VB.ListBox lstQuestLog 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   2550
         Left            =   240
         TabIndex        =   117
         Top             =   480
         Width           =   2415
      End
      Begin VB.Image imgQuestButton 
         Height          =   435
         Index           =   1
         Left            =   390
         Top             =   3480
         Width           =   315
      End
      Begin VB.Image imgQuestButton 
         Height          =   435
         Index           =   6
         Left            =   2190
         Top             =   3480
         Width           =   315
      End
      Begin VB.Image imgQuestButton 
         Height          =   435
         Index           =   5
         Left            =   1830
         Top             =   3480
         Width           =   315
      End
      Begin VB.Image imgQuestButton 
         Height          =   435
         Index           =   4
         Left            =   1470
         Top             =   3480
         Width           =   315
      End
      Begin VB.Image imgQuestButton 
         Height          =   435
         Index           =   3
         Left            =   1110
         Top             =   3480
         Width           =   315
      End
      Begin VB.Image imgQuestButton 
         Height          =   435
         Index           =   2
         Left            =   750
         Top             =   3480
         Width           =   315
      End
   End
   Begin VB.PictureBox picSpells 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8115
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   54
      Top             =   4245
      Visible         =   0   'False
      Width           =   2910
   End
   Begin VB.PictureBox picQuestDialogue 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   1440
      ScaleHeight     =   153
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   327
      TabIndex        =   119
      Top             =   1920
      Visible         =   0   'False
      Width           =   4905
      Begin VB.Label lblQuestSay 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1125
         Left            =   240
         TabIndex        =   125
         Top             =   720
         Width           =   4425
      End
      Begin VB.Label lblQuestAccept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accept Quest"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   210
         Left            =   240
         TabIndex        =   124
         Top             =   1920
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label lblQuestClose 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   210
         Left            =   4200
         TabIndex        =   123
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label lblQuestName 
         BackStyle       =   0  'Transparent
         Caption         =   "Quest Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Left            =   240
         TabIndex        =   122
         Top             =   120
         Width           =   4335
      End
      Begin VB.Label lblQuestExtra 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Extra"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   121
         Top             =   1920
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label lblQuestSubtitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Subtitle"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Left            =   240
         TabIndex        =   120
         Top             =   480
         Width           =   4335
      End
   End
   Begin VB.PictureBox picEventChat 
      BackColor       =   &H000C0E10&
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   120
      ScaleHeight     =   1800
      ScaleWidth      =   7140
      TabIndex        =   109
      Top             =   6600
      Visible         =   0   'False
      Width           =   7140
      Begin VB.Label lblChoices 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "< Choice 4 >"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   4
         Left            =   5280
         TabIndex        =   115
         Top             =   1080
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblChoices 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "< Choice 3 >"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   3
         Left            =   3600
         TabIndex        =   114
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblChoices 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "< Choice 2 >"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   2
         Left            =   1920
         TabIndex        =   113
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblChoices 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "< Choice 1 >"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   1
         Left            =   240
         TabIndex        =   112
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblEventChatContinue 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "< Continue >"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4680
         TabIndex        =   111
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblEventChat 
         BackColor       =   &H000C0E10&
         Caption         =   "This is text that appears for an event."
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   0
         TabIndex        =   110
         Top             =   1440
         Width           =   6975
      End
   End
   Begin VB.PictureBox picItemDesc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   -120
      Picture         =   "frmMain.frx":28A97
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   6
      Top             =   9120
      Visible         =   0   'False
      Width           =   3150
      Begin VB.PictureBox picItemDescPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   1920
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   87
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblwill_req 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   255
         Left            =   360
         TabIndex        =   195
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label lblagi_req 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   255
         Left            =   360
         TabIndex        =   194
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label lblint_req 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   255
         Left            =   360
         TabIndex        =   193
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label lblend_req 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   255
         Left            =   360
         TabIndex        =   192
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label lblstr_req 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   255
         Left            =   360
         TabIndex        =   191
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lbllvl_req 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   255
         Left            =   360
         TabIndex        =   190
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lblitemType 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   189
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblDMG 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   188
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblPrice 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   187
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label lblitemspeed 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   186
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblStat 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   360
         TabIndex        =   185
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblItemDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "details"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1890
         Left            =   1440
         TabIndex        =   86
         Top             =   2040
         Width           =   1440
      End
      Begin VB.Label lblItemName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   150
         TabIndex        =   7
         Top             =   240
         Width           =   2805
      End
   End
   Begin VB.PictureBox picSpellDesc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3570
      Left            =   3240
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   62
      Top             =   9120
      Visible         =   0   'False
      Width           =   3150
      Begin VB.PictureBox picSpellDescPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   1095
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   92
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblSpellDesc 
         BackStyle       =   0  'Transparent
         Caption         =   """This is an example of an item's description. It  can be quite big, so we have to keep it at a decent size."""
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1530
         Left            =   240
         TabIndex        =   91
         Top             =   1800
         Width           =   2640
      End
      Begin VB.Label lblSpellName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   150
         TabIndex        =   90
         Top             =   210
         Width           =   2805
      End
   End
   Begin VB.PictureBox picDialogue 
      Appearance      =   0  'Flat
      BackColor       =   &H000C0E10&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2085
      Left            =   6480
      ScaleHeight     =   2085
      ScaleWidth      =   7140
      TabIndex        =   96
      Top             =   11880
      Visible         =   0   'False
      Width           =   7140
      Begin VB.Label lblDialogue_Button 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Okay"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   3285
         TabIndex        =   101
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label lblDialogue_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Robin has requested a trade. Would you like to accept?"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   100
         Top             =   720
         Width           =   6615
      End
      Begin VB.Label lblDialogue_Title 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Trade Request"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   99
         Top             =   480
         Width           =   6855
      End
      Begin VB.Label lblDialogue_Button 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   3375
         TabIndex        =   98
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label lblDialogue_Button 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   3405
         TabIndex        =   97
         Top             =   1560
         Width           =   285
      End
   End
   Begin VB.PictureBox picCurrency 
      Appearance      =   0  'Flat
      BackColor       =   &H000C0E10&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2085
      Left            =   6480
      ScaleHeight     =   2085
      ScaleWidth      =   7140
      TabIndex        =   55
      Top             =   9720
      Visible         =   0   'False
      Width           =   7140
      Begin VB.TextBox txtCurrency 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   57
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label lblCurrencyCancel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   3240
         TabIndex        =   59
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblCurrencyOk 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Okay"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   3300
         TabIndex        =   58
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblCurrency 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "How many do you want to drop?"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   56
         Top             =   480
         Width           =   3855
      End
   End
   Begin VB.PictureBox picTempInv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   6480
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   4
      Top             =   9120
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picTempBank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   7080
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   73
      Top             =   9120
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picTempSpell 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   7680
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   94
      Top             =   9120
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picSSMap 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   12360
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   95
      Top             =   8280
      Width           =   255
   End
   Begin VB.PictureBox picCover 
      Appearance      =   0  'Flat
      BackColor       =   &H00181C21&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   12000
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   89
      Top             =   8280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picHotbar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   180
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   476
      TabIndex        =   88
      Top             =   5985
      Width           =   7140
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1800
      Left            =   180
      TabIndex        =   1
      Top             =   6630
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   3175
      _Version        =   393217
      BackColor       =   790032
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":2E18B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtMyChat 
      Appearance      =   0  'Flat
      BackColor       =   &H000C0E10&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   720
      TabIndex        =   2
      Top             =   8475
      Width           =   6600
   End
   Begin VB.PictureBox picBank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   120
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   72
      Top             =   120
      Visible         =   0   'False
      Width           =   7200
   End
   Begin VB.PictureBox picShop 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5115
      Left            =   1680
      ScaleHeight     =   341
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   60
      Top             =   480
      Visible         =   0   'False
      Width           =   4125
      Begin VB.PictureBox picShopItems 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3165
         Left            =   615
         ScaleHeight     =   211
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   61
         Top             =   630
         Width           =   2895
      End
      Begin VB.Image imgLeaveShop 
         Height          =   435
         Left            =   2715
         Top             =   4350
         Width           =   1035
      End
      Begin VB.Image imgShopSell 
         Height          =   435
         Left            =   1545
         Top             =   4350
         Width           =   1035
      End
      Begin VB.Image imgShopBuy 
         Height          =   435
         Left            =   375
         Top             =   4350
         Width           =   1035
      End
   End
   Begin VB.PictureBox picTrade 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   150
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   74
      Top             =   150
      Visible         =   0   'False
      Width           =   7200
      Begin VB.PictureBox picTheirTrade 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   3855
         ScaleHeight     =   247
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   76
         Top             =   465
         Width           =   2895
      End
      Begin VB.PictureBox picYourTrade 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   435
         ScaleHeight     =   247
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   75
         Top             =   465
         Width           =   2895
      End
      Begin VB.Image imgDeclineTrade 
         Height          =   435
         Left            =   3675
         Top             =   5040
         Width           =   1035
      End
      Begin VB.Image imgAcceptTrade 
         Height          =   435
         Left            =   2475
         Top             =   5040
         Width           =   1035
      End
      Begin VB.Label lblTradeStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   600
         TabIndex        =   79
         Top             =   5520
         Width           =   5895
      End
      Begin VB.Label lblTheirWorth 
         BackStyle       =   0  'Transparent
         Caption         =   "1234567890"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5160
         TabIndex        =   78
         Top             =   4500
         Width           =   1815
      End
      Begin VB.Label lblYourWorth 
         BackStyle       =   0  'Transparent
         Caption         =   "1234567890"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   77
         Top             =   4500
         Width           =   1815
      End
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00181C21&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   150
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   150
      Visible         =   0   'False
      Width           =   7200
      Begin MSWinsockLib.Winsock Socket 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   8
      Left            =   9045
      Top             =   2280
      Width           =   1035
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   7
      Left            =   10245
      Top             =   2280
      Width           =   1035
   End
   Begin VB.Label lblEXP 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   9480
      TabIndex        =   84
      Top             =   1080
      Width           =   1845
   End
   Begin VB.Label lblMP 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   9480
      TabIndex        =   83
      Top             =   750
      Width           =   1845
   End
   Begin VB.Label lblHP 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   9480
      TabIndex        =   82
      Top             =   420
      Width           =   1845
   End
   Begin VB.Image imgEXPBar 
      Height          =   240
      Left            =   7770
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Image imgMPBar 
      Height          =   240
      Left            =   7770
      Top             =   750
      Width           =   3615
   End
   Begin VB.Image imgHPBar 
      Height          =   240
      Left            =   7770
      Top             =   420
      Width           =   3615
   End
   Begin VB.Label lblPing 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   8520
      TabIndex        =   81
      Top             =   1920
      Width           =   450
   End
   Begin VB.Label lblGold 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0g"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   8520
      TabIndex        =   80
      Top             =   1515
      Width           =   225
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   6
      Left            =   10245
      Top             =   3450
      Width           =   1035
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   5
      Left            =   9045
      Top             =   3450
      Width           =   1035
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   4
      Left            =   7845
      Top             =   3450
      Width           =   1035
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   3
      Left            =   10245
      Top             =   2850
      Width           =   1035
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   2
      Left            =   9045
      Top             =   2850
      Width           =   1035
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   1
      Left            =   7845
      Top             =   2850
      Width           =   1035
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ************
' ** Events **
' ************
Private MoveForm As Boolean
Private MouseX As Long
Private MouseY As Long
Private PresentX As Long
Private PresentY As Long




Private Sub cmdAAnim_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < RankDeveloper Then
        
        Exit Sub
    End If

    SendRequestEditAnimation
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAAnim_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdLevel_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < RankDeveloper Then
        
        Exit Sub
    End If

    SendRequestLevelUp
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdLevel_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub




Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    ' move GUI
    picAdmin.Left = 544
    picCurrency.Left = txtChat.Left
    picCurrency.Top = txtChat.Top
    picDialogue.Top = txtChat.Top
    picDialogue.Left = txtChat.Left
    picCover.Top = picScreen.Top - 1
    picCover.Left = picScreen.Left - 1
    picCover.Height = picScreen.Height + 2
    picCover.Width = picScreen.Width + 2
    comboresolution.AddItem "480x384"
    comboresolution.AddItem "800x600"
    comboresolution.ListIndex = 0
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    Cancel = True
    logoutGame
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    
    ' reset all buttons
    resetButtons_Main
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgAcceptTrade_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    AcceptTrade
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgAcceptTrade_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_Click(Index As Integer)
Dim Buffer As clsBuffer
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 1
                ' show the window
                picInventory.Visible = Not picInventory.Visible
                picCharacter.Visible = False
                picSpells.Visible = False
                picOptions.Visible = False
                picParty.Visible = False
                frmMain.picInventory.Refresh
                picQuestLog.Visible = False
                FraLevels.Visible = False
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
        Case 2
                ' send packet
                Set Buffer = New clsBuffer
                Buffer.WriteLong CSpells
                SendData Buffer.ToArray()
                Set Buffer = Nothing
                ' show the window
                picSpells.Visible = Not picSpells.Visible
                picInventory.Visible = False
                picCharacter.Visible = False
                picOptions.Visible = False
                picParty.Visible = False
                picQuestLog.Visible = False
                FraLevels.Visible = False
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
        Case 3
                ' send packet
                SendRequestPlayerData
                ' show the window
                picCharacter.Visible = Not picCharacter.Visible
                picInventory.Visible = False
                picSpells.Visible = False
                picOptions.Visible = False
                picParty.Visible = False
                picQuestLog.Visible = False
                FraLevels.Visible = False
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
                ' Render
                frmMain.picCharacter.Refresh
                frmMain.picFace.Refresh
        Case 4
                ' show the window
                picCharacter.Visible = False
                picInventory.Visible = False
                picSpells.Visible = False
                picOptions.Visible = Not picOptions.Visible
                picParty.Visible = False
                picQuestLog.Visible = False
                FraLevels.Visible = False
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
        Case 5
            If myTargetType = TargetPlayer And myTarget <> MyIndex Then
                SendTradeRequest
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            Else
                AddText "Invalid trade target.", BrightRed
            End If
        Case 6
            ' show the window
            picCharacter.Visible = False
            picInventory.Visible = False
            picSpells.Visible = False
            picOptions.Visible = False
            picParty.Visible = Not picParty.Visible
            picQuestLog.Visible = False
            FraLevels.Visible = False
            ' play sound
            PlaySound Sound_ButtonClick, -1, -1
        'Alatar v1.2
        Case 7 'QuestLog
            picCharacter.Visible = False
            picInventory.Visible = False
            picSpells.Visible = False
            picOptions.Visible = False
            picParty.Visible = False
            picQuestLog.Visible = Not picQuestLog.Visible
            FraLevels.Visible = False
            UpdateQuestLog
            PlaySound Sound_ButtonClick, -1, -1
        '/Alatar v1.2
        Case 8 'levels
            picCharacter.Visible = False
            picInventory.Visible = False
            picSpells.Visible = False
            picOptions.Visible = False
            picParty.Visible = False
            picQuestLog.Visible = False
            FraLevels.Visible = Not FraLevels.Visible
            PlaySound Sound_ButtonClick, -1, -1
        '/Alatar v1.2
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    ' reset other buttons
    resetButtons_Main Index
    
    ' change the button we're hovering on
    If Not MainButton(Index).state = 2 Then ' make sure we're not clicking
        changeButtonState_Main Index, 1 ' hover
    End If
    
    ' play sound
    If Not LastButtonSound_Main = Index Then
        PlaySound Sound_ButtonHover, -1, -1
        LastButtonSound_Main = Index
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
        
    ' reset all buttons
    resetButtons_Main -1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    ' reset other buttons
    resetButtons_Main Index
    
    ' change the button we're hovering on
    changeButtonState_Main Index, 2 ' clicked
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub




Private Sub lblChoices_Click(Index As Integer)
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer
Buffer.WriteLong CEventChatReply
Buffer.WriteLong EventReplyID
Buffer.WriteLong EventReplyPage
Buffer.WriteLong Index
SendData Buffer.ToArray
Set Buffer = Nothing
ClearEventChat
InEvent = False
End Sub

Private Sub lblCurrencyCancel_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    picCurrency.Visible = False
    txtCurrency.text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblCurrencyCancel_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgDeclineTrade_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    DeclineTrade
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgDeclineTrade_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgLeaveShop_Click()
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CCloseShop
    
    SendData Buffer.ToArray()
    
    Set Buffer = Nothing
    
    picCover.Visible = False
    picShop.Visible = False
    InShop = 0
    ShopAction = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgLeaveShop_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblCurrencyOk_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If IsNumeric(txtCurrency.text) Then
        If Val(txtCurrency.text) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then txtCurrency.text = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
  Select Case CurrencyMenu
    Case 1 ' drop item
          If Val(txtCurrency.text) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then
          txtCurrency.text = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
            Else
               AddText "Please enter a valid amount.", BrightRed
                Exit Sub
           End If
          SendDropItem tmpCurrencyItem, Val(txtCurrency.text)
        Case 2 ' deposit item
           If Val(txtCurrency.text) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then
                txtCurrency.text = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
            Else
              AddText "Please enter a valid amount.", BrightRed
              Exit Sub
            End If
           DepositItem tmpCurrencyItem, Val(txtCurrency.text)
        Case 3 ' withdraw item
            If Val(txtCurrency.text) > Bank.Item(tmpCurrencyItem).Value Then
                txtCurrency.text = Bank.Item(tmpCurrencyItem).Value
            Else
               AddText "Please enter a valid amount.", BrightRed
                Exit Sub
             End If
            WithdrawItem tmpCurrencyItem, Val(txtCurrency.text)
        Case 4 ' offer trade item
            If Val(txtCurrency.text) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then
                txtCurrency.text = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
            Else
                AddText "Please enter a valid amount.", BrightRed
             Exit Sub
           End If
         TradeItem tmpCurrencyItem, Val(txtCurrency.text)
         End Select
    End If
    
    picCurrency.Visible = False
    tmpCurrencyItem = 0
    txtCurrency.text = vbNullString
    CurrencyMenu = 0 ' clear
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblCurrencyOk_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgShopBuy_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If ShopAction = 1 Then Exit Sub
    ShopAction = 1 ' buying an item
    AddText "Click on the item in the shop you wish to buy.", White
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgShopBuy_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgShopSell_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If ShopAction = 2 Then Exit Sub
    ShopAction = 2 ' selling an item
    AddText "Double-click on the item in your inventory you wish to sell.", White
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgShopSell_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblDialogue_Button_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    ' call the handler
    dialogueHandler Index
    
    picDialogue.Visible = False
    dialogueIndex = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblDialogue_Button_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub



Private Sub lblEventChat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEventChatContinue.FontBold = False
End Sub

Private Sub lblEventChatContinue_Click()
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer
Buffer.WriteLong CEventChatReply
Buffer.WriteLong EventReplyID
Buffer.WriteLong EventReplyPage
Buffer.WriteLong 0
SendData Buffer.ToArray
Set Buffer = Nothing
ClearEventChat
InEvent = False
End Sub

Private Sub lblEventChatContinue_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEventChatContinue.FontBold = False
End Sub

Private Sub lblEventChatContinue_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEventChatContinue.FontBold = True
End Sub

Private Sub lblEventChatContinue_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEventChatContinue.FontBold = True
End Sub

Private Sub lblPartyInvite_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If myTargetType = TargetPlayer And myTarget <> MyIndex Then
        SendPartyRequest
    Else
        AddText "Invalid invitation target.", BrightRed
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblPartyInvite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblPartyLeave_Click()
        ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If Party.Leader > 0 Then
        SendPartyLeave
    Else
        AddText "You are not in a party.", BrightRed
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblPartyInvite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblTrainStat_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerPOINTS(MyIndex) = 0 Then Exit Sub
    SendTrainStat Index
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblTrainStat_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optdet_Click()
If optdet.Value Then
   
    ' New Resolution BEGIN
Dim resWidth, resHeight As Long
Dim res() As String
     options.HQ = 0
      Call SaveOptions
    res() = Split(comboresolution.List(comboresolution.ListIndex), "x")


    Direct3D_Window.BackBufferWidth = res(0)
    Direct3D_Window.BackBufferHeight = res(1)

    MAX_MAPX = res(0) / 32 - 1
    MAX_MAPY = res(1) / 32 - 1
    
    HalfX = ((MAX_MAPX + 1) / 2) * PIC_X
    HalfY = ((MAX_MAPY + 1) / 2) * PIC_Y
    ScreenX = (MAX_MAPX + 1) * PIC_X
    ScreenY = (MAX_MAPY + 1) * PIC_Y
    StartXValue = ((MAX_MAPX + 1) / 2)
    StartYValue = ((MAX_MAPY + 1) / 2)
    EndXValue = (MAX_MAPX + 1) + 1
    EndYValue = (MAX_MAPY + 1) + 1
    
    ' Change Background
    frmMain.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\main" & 0 & ".jpg")
    
    picScreen.Width = 480
    picScreen.Height = 384
    
    ' Move GUI
    imgHPBar.Left = 518
    imgHPBar.Top = 28
    imgMPBar.Left = 518
    imgMPBar.Top = 50
    imgEXPBar.Left = 518
    imgEXPBar.Top = 72
    lblHP.Left = 632
    lblHP.Top = 28
    lblMP.Left = 632
    lblMP.Top = 50
    lblEXP.Left = 632
    lblEXP.Top = 72
    lblGold.Left = 568
    lblGold.Top = 101
    lblPing.Left = 568
    lblPing.Top = 128
    
    imgButton(1).Left = 523
    imgButton(1).Top = 190
    imgButton(2).Left = 603
    imgButton(2).Top = 190
    imgButton(3).Left = 683
    imgButton(3).Top = 190
    imgButton(4).Left = 523
    imgButton(4).Top = 230
    imgButton(5).Left = 603
    imgButton(5).Top = 230
    imgButton(6).Left = 683
    imgButton(6).Top = 230
    imgButton(7).Left = 683
    imgButton(7).Top = 152
    imgButton(8).Left = 603
    imgButton(8).Top = 152
    
    picParty.Left = 541
    picParty.Top = 283
    picOptions.Left = 541
    picOptions.Top = 283
    picSpells.Left = 541
    picSpells.Top = 283
    picInventory.Left = 541
    picInventory.Top = 283
    picCharacter.Left = 541
    picCharacter.Top = 283
    FraLevels.Left = 541
    FraLevels.Top = 283
    picQuestLog.Left = 541
    picQuestLog.Top = 283
    
    picSSMap.Left = 800
    picCover.Left = 824
    
    picAdmin.Left = 544
    
    picTrade.Left = picScreen.Left + (picScreen.Width / 2 - picTrade.Width / 2)
    picShop.Left = picScreen.Left + (picScreen.Width / 2 - picShop.Width / 2)
    picBank.Left = picScreen.Left + (picScreen.Width / 2 - picBank.Width / 2)
    picHotbar.Top = 399
    picEventChat.Top = 442
    txtMyChat.Top = 565
    txtChat.Top = 442
    frmMain.Width = 11865
    frmMain.Height = 9315
    comboresolution.ListIndex = 0
End If
If optdet.Value Then
scrnopt.Value = False
End If
End Sub

Private Sub optMOff_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    options.Music = 0
    ' stop music playing
    StopMusic
    ' save to config.ini
    SaveOptions
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optMOff_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optMOn_Click()
Dim MusicFile As String
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    options.Music = 1
    ' start music playing
    MusicFile = Trim$(Map.Music)
    If Not MusicFile = "None." Then
        PlayMusic MusicFile
    Else
        StopMusic
    End If
    ' save to config.ini
    SaveOptions
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optMOn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optSOff_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    StopAllSounds
    options.sound = 0
    ' save to config.ini
    SaveOptions
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optSOff_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optSOn_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    options.sound = 1
    ' save to config.ini
    SaveOptions
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optSOn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picaxes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAxesExp.Visible = True
If lblAxesExp.Visible = True Then

 lblswordsexp.Visible = False
 lblDaggersExp.Visible = False
 lblRangeExp.Visible = False
 lblMagicExp.Visible = False
 lblConvictionExp.Visible = False
 lblLycanthropyExp.Visible = False
 lblHeavyarmorExp.Visible = False
 lblLightArmorExp.Visible = False
 lblMiningExp.Visible = False
 lblWoodcuttingExp.Visible = False
 lblFishingExp.Visible = False
 lblCraftingExp.Visible = False
 lblEnchantsExp.Visible = False
 lblAlchemyExp.Visible = False
 lblSmithExp.Visible = False
 End If
End Sub

Private Sub picCover_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    
    ' reset all buttons
    resetButtons_Main
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picCover_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub picEventChat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEventChatContinue.FontBold = False
End Sub

Private Sub picHotbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim SlotNum As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    SlotNum = IsHotbarSlot(X, Y)

    If Button = 1 Then
        If SlotNum <> 0 Then
            SendHotbarUse SlotNum
        End If
    ElseIf Button = 2 Then
        If SlotNum <> 0 Then
            SendHotbarChange 0, 0, SlotNum
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picHotbar_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picHotbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim SlotNum As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    SlotNum = IsHotbarSlot(X, Y)

    If SlotNum <> 0 Then
        If Hotbar(SlotNum).sType = 1 Then ' item
            X = X + picHotbar.Left + 1
            Y = Y + picHotbar.Top - picItemDesc.Height - 1
            UpdateDescWindow Hotbar(SlotNum).Slot, X, Y
            LastItemDesc = Hotbar(SlotNum).Slot ' set it so you don't re-set values
            Exit Sub
        ElseIf Hotbar(SlotNum).sType = 2 Then ' spell
            X = X + picHotbar.Left + 1
            Y = Y + picHotbar.Top - picSpellDesc.Height - 1
            UpdateSpellWindow Hotbar(SlotNum).Slot, X, Y
            LastSpellDesc = Hotbar(SlotNum).Slot  ' set it so you don't re-set values
            Exit Sub
        End If
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    picSpellDesc.Visible = False
    LastSpellDesc = 0 ' no spell was last loaded
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picHotbar_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub




Private Sub piclyconthropy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblLycanthropyExp.Visible = True
If lblLycanthropyExp.Visible = True Then

 lblswordsexp.Visible = False
 lblAxesExp.Visible = False
 lblDaggersExp.Visible = False
 lblRangeExp.Visible = False
 lblConvictionExp.Visible = False
 lblMagicExp.Visible = False
 lblHeavyarmorExp.Visible = False
 lblLightArmorExp.Visible = False
 lblMiningExp.Visible = False
 lblWoodcuttingExp.Visible = False
 lblFishingExp.Visible = False
 lblCraftingExp.Visible = False
 lblEnchantsExp.Visible = False
 lblAlchemyExp.Visible = False
 lblSmithExp.Visible = False
 End If
End Sub

Private Sub PicMagic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblMagicExp.Visible = True
If lblMagicExp.Visible = True Then

 lblswordsexp.Visible = False
 lblAxesExp.Visible = False
 lblDaggersExp.Visible = False
 lblRangeExp.Visible = False
 lblConvictionExp.Visible = False
 lblLycanthropyExp.Visible = False
 lblHeavyarmorExp.Visible = False
 lblLightArmorExp.Visible = False
 lblMiningExp.Visible = False
 lblWoodcuttingExp.Visible = False
 lblFishingExp.Visible = False
 lblCraftingExp.Visible = False
 lblEnchantsExp.Visible = False
 lblAlchemyExp.Visible = False
 lblSmithExp.Visible = False
 End If
End Sub

Private Sub PicRange_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRangeExp.Visible = True
If lblRangeExp.Visible = True Then

 lblswordsexp.Visible = False
 lblAxesExp.Visible = False
 lblDaggersExp.Visible = False
 lblMagicExp.Visible = False
 lblConvictionExp.Visible = False
 lblLycanthropyExp.Visible = False
 lblHeavyarmorExp.Visible = False
 lblLightArmorExp.Visible = False
 lblMiningExp.Visible = False
 lblWoodcuttingExp.Visible = False
 lblFishingExp.Visible = False
 lblCraftingExp.Visible = False
 lblEnchantsExp.Visible = False
 lblAlchemyExp.Visible = False
 lblSmithExp.Visible = False
 End If
End Sub


Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If InMapEditor Then
        Call MapEditorMouseDown(Button, X, Y, False)
    Else
        ' left click
        If Button = vbLeftButton Then
            ' targetting
            Call PlayerSearch(CurX, CurY)
        ' right click
        ElseIf Button = vbRightButton Then
            If ShiftDown Then
                ' admin warp if we're pressing shift and right clicking
                If GetPlayerAccess(MyIndex) >= 2 Then AdminWarp CurX, CurY
            End If
        End If
    End If

    Call SetFocusOnChat
    
    If frmEditor_Events.Visible Then frmEditor_Events.SetFocus
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picScreen_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    If Map.Tile(Player(MyIndex).X, Player(MyIndex).Y).Type = TileCraft Then
     If Map.Tile(Player(MyIndex).X, Player(MyIndex).Y).Data1 = 2 Then
      Picture22.Visible = True
      End If
    End If
    If Map.Tile(Player(MyIndex).X, Player(MyIndex).Y).Type < TileCraft Then
      Picture22.Visible = False
    End If
    CurX = TileView.Left + ((X + Camera.Left) \ PIC_X)
    CurY = TileView.Top + ((Y + Camera.Top) \ PIC_Y)

    If InMapEditor Then
        If Button = vbLeftButton Or Button = vbRightButton Then
            Call MapEditorMouseDown(Button, X, Y)
        End If
    End If
    
    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    
    ' reset all buttons
    resetButtons_Main
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picScreen_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function IsShopItem(ByVal X As Single, ByVal Y As Single) As Long
Dim tempRec As RECT
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    IsShopItem = 0

    For i = 1 To MAX_TRADES

        If Shop(InShop).TradeItem(i).Item > 0 And Shop(InShop).TradeItem(i).Item <= MAX_ITEMS Then
            With tempRec
                .Top = ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
                .Bottom = .Top + PIC_Y
                .Left = ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsShopItem = i
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsShopItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Sub picShop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    
    ' reset all buttons
    resetButtons_Main
End Sub

Private Sub picShopItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim shopItem As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    shopItem = IsShopItem(X, Y)
    
    If shopItem > 0 Then
        Select Case ShopAction
            Case 0 ' no action, give cost
                With Shop(InShop).TradeItem(shopItem)
                    AddText "You can buy this item for " & .CostValue & " " & Trim$(Item(.CostItem).Name) & ".", White
                End With
            Case 1 ' buy item
                ' buy item code
                BuyItem shopItem
        End Select
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picShopItems_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picShopItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim shopslot As Long
Dim X2 As Long, Y2 As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    shopslot = IsShopItem(X, Y)

    If shopslot <> 0 Then
        X2 = X + picShop.Left + picShopItems.Left + 1
        Y2 = Y + picShop.Top + picShopItems.Top + 1
        UpdateDescWindow Shop(InShop).TradeItem(shopslot).Item, X2, Y2
        LastItemDesc = Shop(InShop).TradeItem(shopslot).Item
        Exit Sub
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picShopItems_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpellDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    picSpellDesc.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpellDesc_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpells_DblClick()
Dim spellnum As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    spellnum = IsPlayerSpell(SpellX, SpellY)

    If spellnum <> 0 Then
        Call CastSpell(spellnum)
        Exit Sub
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpells_DblClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpells_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim spellnum As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    spellnum = IsPlayerSpell(SpellX, SpellY)
    If Button = 1 Then ' left click
        If spellnum <> 0 Then
            DragSpell = spellnum
            frmMain.picTempSpell.Refresh
            Exit Sub
        End If
    ElseIf Button = 2 Then ' right click
        If spellnum <> 0 Then
            Dialogue "Forget Spell", "Are you sure you want to forget how to cast " & Trim$(Spell(PlayerSpells(spellnum)).Name) & "?", DialogueForget, True, spellnum
            Exit Sub
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpells_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpells_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim spellslot As Long
Dim X2 As Long, Y2 As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    SpellX = X
    SpellY = Y
    
    spellslot = IsPlayerSpell(X, Y)
    
    If DragSpell > 0 Then
        With frmMain.picTempSpell
            .Top = Y + frmMain.picSpells.Top
            .Left = X + frmMain.picSpells.Left
            .Visible = True
            .ZOrder (0)
        End With
    Else
        If spellslot <> 0 Then
            X2 = X + picSpells.Left - picSpellDesc.Width - 1
            Y2 = Y + picSpells.Top - picSpellDesc.Height - 1
            UpdateSpellWindow PlayerSpells(spellslot), X2, Y2
            LastSpellDesc = PlayerSpells(spellslot)
            Exit Sub
        End If
    End If
    
    picSpellDesc.Visible = False
    LastSpellDesc = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpells_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpells_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim rec_pos As RECT

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If DragSpell > 0 Then
        ' drag + drop
        For i = 1 To MAX_PLAYER_SPELLS
            With rec_pos
                .Top = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                .Bottom = .Top + PIC_Y
                .Left = SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                    If DragSpell <> i Then
                        SendChangeSpellSlots DragSpell, i
                        Exit For
                    End If
                End If
            End If
        Next
        ' hotbar
        For i = 1 To MAX_HOTBAR
            With rec_pos
                .Top = picHotbar.Top - picSpells.Top
                .Left = picHotbar.Left - picSpells.Left + (HotbarOffsetX * (i - 1)) + (32 * (i - 1))
                .Right = .Left + 32
                .Bottom = picHotbar.Top - picSpells.Top + 32
            End With
            
            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                    SendHotbarChange 2, DragSpell, i
                    DragSpell = 0
                    picTempSpell.Visible = False
                    Exit Sub
                End If
            End If
        Next
    End If

    DragSpell = 0
    picTempSpell.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpells_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picswords_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblswordsexp.Visible = True
If lblswordsexp.Visible = True Then

lblAxesExp.Visible = False
lblDaggersExp.Visible = False
lblRangeExp.Visible = False
lblMagicExp.Visible = False
lblConvictionExp.Visible = False
lblLycanthropyExp.Visible = False
lblHeavyarmorExp.Visible = False
lblLightArmorExp.Visible = False
lblMiningExp.Visible = False
lblWoodcuttingExp.Visible = False
lblFishingExp.Visible = False
lblCraftingExp.Visible = False
lblEnchantsExp.Visible = False
lblAlchemyExp.Visible = False
lblSmithExp.Visible = False
End If
End Sub

Private Sub picTrade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picTrade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDaggersExp.Visible = True
If lblDaggersExp.Visible = True Then

 lblswordsexp.Visible = False
 lblAxesExp.Visible = False
 lblRangeExp.Visible = False
 lblMagicExp.Visible = False
 lblConvictionExp.Visible = False
 lblLycanthropyExp.Visible = False
 lblHeavyarmorExp.Visible = False
 lblLightArmorExp.Visible = False
 lblMiningExp.Visible = False
 lblWoodcuttingExp.Visible = False
 lblFishingExp.Visible = False
 lblCraftingExp.Visible = False
 lblEnchantsExp.Visible = False
 lblAlchemyExp.Visible = False
 lblSmithExp.Visible = False
 End If
End Sub

Private Sub Picture10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblWoodcuttingExp.Visible = True
If lblWoodcuttingExp.Visible Then

 lblswordsexp.Visible = False
 lblAxesExp.Visible = False
 lblDaggersExp.Visible = False
 lblRangeExp.Visible = False
 lblConvictionExp.Visible = False
 lblMagicExp.Visible = False
 lblLycanthropyExp.Visible = False
 lblHeavyarmorExp.Visible = False
 lblLightArmorExp.Visible = False
 lblMiningExp.Visible = False
 lblFishingExp.Visible = False
 lblCraftingExp.Visible = False
 lblEnchantsExp.Visible = False
 lblAlchemyExp.Visible = False
 lblSmithExp.Visible = False
 End If
End Sub

Private Sub Picture11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblFishingExp.Visible = True
If lblFishingExp.Visible Then

 lblswordsexp.Visible = False
 lblAxesExp.Visible = False
 lblDaggersExp.Visible = False
 lblRangeExp.Visible = False
 lblConvictionExp.Visible = False
 lblMagicExp.Visible = False
 lblLycanthropyExp.Visible = False
 lblHeavyarmorExp.Visible = False
 lblLightArmorExp.Visible = False
 lblMiningExp.Visible = False
 lblWoodcuttingExp.Visible = False
 lblCraftingExp.Visible = False
 lblEnchantsExp.Visible = False
 lblAlchemyExp.Visible = False
 lblSmithExp.Visible = False
 End If
End Sub

Private Sub Picture12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCraftingExp.Visible = True
If lblCraftingExp.Visible Then

 lblswordsexp.Visible = False
 lblAxesExp.Visible = False
 lblDaggersExp.Visible = False
 lblRangeExp.Visible = False
 lblConvictionExp.Visible = False
 lblMagicExp.Visible = False
 lblLycanthropyExp.Visible = False
 lblHeavyarmorExp.Visible = False
 lblLightArmorExp.Visible = False
 lblMiningExp.Visible = False
 lblWoodcuttingExp.Visible = False
 lblFishingExp.Visible = False
 lblEnchantsExp.Visible = False
 lblAlchemyExp.Visible = False
 lblSmithExp.Visible = False
 End If
End Sub

Private Sub Picture13_Click()
       Call SendCraftin(2)
End Sub

Private Sub Picture14_Click()
        Call SendCraftin(3)
End Sub

Private Sub Picture15_Click()
        Call SendCraftin(4)
End Sub

Private Sub Picture16_Click()
         Call SendCraftin(5)
End Sub

Private Sub Picture17_Click()
         Call SendCraftin(6)
End Sub

Private Sub Picture18_Click()
         Call SendCraftin(7)
End Sub

Private Sub Picture19_Click()
      Call SendCraftin(8)
End Sub


Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblConvictionExp.Visible = True
If lblConvictionExp.Visible = True Then

 lblswordsexp.Visible = False
 lblAxesExp.Visible = False
 lblDaggersExp.Visible = False
 lblRangeExp.Visible = False
 lblMagicExp.Visible = False
 lblLycanthropyExp.Visible = False
 lblHeavyarmorExp.Visible = False
 lblLightArmorExp.Visible = False
 lblMiningExp.Visible = False
 lblWoodcuttingExp.Visible = False
 lblFishingExp.Visible = False
 lblCraftingExp.Visible = False
 lblEnchantsExp.Visible = False
 lblAlchemyExp.Visible = False
 lblSmithExp.Visible = False
 End If
End Sub

Private Sub Picture22_Click()
 Fracrafting.Visible = Not Fracrafting.Visible
Label2.Caption = "LvlReq" & Item(1).SmithReq
Label4.Caption = "LvlReq" & Item(2).SmithReq
Label5.Caption = "LvlReq" & Item(3).SmithReq
Label6.Caption = "LvlReq" & Item(4).SmithReq
End Sub


Private Sub Picture23_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblSmithExp.Visible = True
If lblSmithExp.Visible Then

 lblswordsexp.Visible = False
 lblAxesExp.Visible = False
 lblDaggersExp.Visible = False
 lblRangeExp.Visible = False
 lblConvictionExp.Visible = False
 lblMagicExp.Visible = False
 lblLycanthropyExp.Visible = False
 lblHeavyarmorExp.Visible = False
 lblLightArmorExp.Visible = False
 lblMiningExp.Visible = False
 lblWoodcuttingExp.Visible = False
 lblFishingExp.Visible = False
 lblEnchantsExp.Visible = False
 lblAlchemyExp.Visible = False
 lblCraftingExp.Visible = False
 End If
End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHeavyarmorExp.Visible = True
If lblHeavyarmorExp.Visible = True Then

 lblswordsexp.Visible = False
 lblAxesExp.Visible = False
 lblDaggersExp.Visible = False
 lblRangeExp.Visible = False
 lblConvictionExp.Visible = False
 lblMagicExp.Visible = False
 lblLycanthropyExp.Visible = False
 lblLightArmorExp.Visible = False
 lblMiningExp.Visible = False
 lblWoodcuttingExp.Visible = False
 lblFishingExp.Visible = False
 lblCraftingExp.Visible = False
 lblEnchantsExp.Visible = False
 lblAlchemyExp.Visible = False
 lblSmithExp.Visible = False
 End If
End Sub

Private Sub Picture6_Click()
        Call SendCraftin(1)
 
End Sub

Private Sub Picture7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblLightArmorExp.Visible = True
If lblLightArmorExp.Visible = True Then

 lblswordsexp.Visible = False
 lblAxesExp.Visible = False
 lblDaggersExp.Visible = False
 lblRangeExp.Visible = False
 lblConvictionExp.Visible = False
 lblMagicExp.Visible = False
 lblLycanthropyExp.Visible = False
 lblHeavyarmorExp.Visible = False
 lblMiningExp.Visible = False
 lblWoodcuttingExp.Visible = False
 lblFishingExp.Visible = False
 lblCraftingExp.Visible = False
 lblEnchantsExp.Visible = False
 lblAlchemyExp.Visible = False
 lblSmithExp.Visible = False
 End If
End Sub

Private Sub Picture8_Click()
Fracrafting.Visible = False
End Sub



Private Sub Picture9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblMiningExp.Visible = True
If lblMiningExp.Visible = True Then

 lblswordsexp.Visible = False
 lblAxesExp.Visible = False
 lblDaggersExp.Visible = False
 lblRangeExp.Visible = False
 lblConvictionExp.Visible = False
 lblMagicExp.Visible = False
 lblLycanthropyExp.Visible = False
 lblHeavyarmorExp.Visible = False
 lblLightArmorExp.Visible = False
 lblWoodcuttingExp.Visible = False
 lblFishingExp.Visible = False
 lblCraftingExp.Visible = False
 lblEnchantsExp.Visible = False
 lblAlchemyExp.Visible = False
 lblSmithExp.Visible = False
 End If
End Sub

Private Sub picYourTrade_DblClick()
Dim TradeNum As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    TradeNum = IsTradeItem(TradeX, TradeY, True)

    If TradeNum <> 0 Then
        UntradeItem TradeNum
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picYourTrade_DlbClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picYourTrade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TradeNum As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    TradeX = X
    TradeY = Y
    
    TradeNum = IsTradeItem(X, Y, True)
    
    If TradeNum <> 0 Then
        X = X + picTrade.Left + picYourTrade.Left + 4
        Y = Y + picTrade.Top + picYourTrade.Top + 4
        UpdateDescWindow GetPlayerInvItemNum(MyIndex, TradeYourOffer(TradeNum).num), X, Y
        LastItemDesc = GetPlayerInvItemNum(MyIndex, TradeYourOffer(TradeNum).num) ' set it so you don't re-set values
        Exit Sub
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picYourTrade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picTheirTrade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TradeNum As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    TradeNum = IsTradeItem(X, Y, False)
    
    If TradeNum <> 0 Then
        X = X + picTrade.Left + picTheirTrade.Left + 4
        Y = Y + picTrade.Top + picTheirTrade.Top + 4
        UpdateDescWindow TradeTheirOffer(TradeNum).num, X, Y
        LastItemDesc = TradeTheirOffer(TradeNum).num ' set it so you don't re-set values
        Exit Sub
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picTheirTrade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAAmount_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    lblAAmount.Caption = "Amount: " & scrlAAmount.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAAmount_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAItem_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    lblAItem.Caption = "Item: " & Trim$(Item(scrlAItem.Value).Name)
    If Item(scrlAItem.Value).Type = ItemCurrency Then
        scrlAAmount.Enabled = True
        Exit Sub
    End If
    scrlAAmount.Enabled = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAItem_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrnopt_Click()
If scrnopt.Value Then
' New Resolution BEGIN
Dim resWidth, resHeight As Long
Dim res() As String
    options.HQ = 1
     Call SaveOptions
    res() = Split(comboresolution.List(comboresolution.ListIndex), "x")

    Direct3D_Window.BackBufferWidth = res(0)
    Direct3D_Window.BackBufferHeight = res(1)

    MAX_MAPX = res(0) / 32 - 1
    MAX_MAPY = res(1) / 32 - 1
    
    HalfX = ((MAX_MAPX + 1) / 2) * PIC_X
    HalfY = ((MAX_MAPY + 1) / 2) * PIC_Y
    ScreenX = (MAX_MAPX + 1) * PIC_X
    ScreenY = (MAX_MAPY + 1) * PIC_Y
    StartXValue = ((MAX_MAPX + 1) / 2)
    StartYValue = ((MAX_MAPY + 1) / 2)
    EndXValue = (MAX_MAPX + 1) + 1
    EndYValue = (MAX_MAPY + 1) + 1
    
    ' Change Background
    frmMain.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\main" & 1 & ".jpg")
    
    picScreen.Width = res(0)
    picScreen.Height = res(1)
    
    ' Move GUI
    imgHPBar.Left = 523
    imgHPBar.Top = 737
    imgMPBar.Left = 523
    imgMPBar.Top = 755
    imgEXPBar.Left = 523
    imgEXPBar.Top = 781
    lblHP.Left = 680
    lblHP.Top = 736
    lblMP.Left = 680
    lblMP.Top = 755
    lblEXP.Left = 680
    lblEXP.Top = 784
    lblGold.Left = 580
    lblGold.Top = 712
    lblPing.Left = 720
    lblPing.Top = 712
    
    imgButton(1).Left = 496
    imgButton(1).Top = 632
    imgButton(2).Left = 576
    imgButton(2).Top = 632
    imgButton(3).Left = 656
    imgButton(3).Top = 632
    imgButton(4).Left = 496
    imgButton(4).Top = 672
    imgButton(5).Left = 576
    imgButton(5).Top = 672
    imgButton(6).Left = 656
    imgButton(6).Top = 672
    imgButton(7).Left = 736
    imgButton(7).Top = 632
    imgButton(8).Left = 736
    imgButton(8).Top = 672
    
    picParty.Left = 592
    picParty.Top = 320
    picOptions.Left = 592
    picOptions.Top = 320
    picSpells.Left = 592
    picSpells.Top = 320
    picInventory.Left = 592
    picInventory.Top = 320
    picCharacter.Left = 592
    picCharacter.Top = 320
    FraLevels.Left = 592
    FraLevels.Top = 320
    picQuestLog.Left = 592
    picQuestLog.Top = 320
    picSSMap.Left = 800 + (2 * 160)
    picCover.Left = 824 + (2 * 160)
    
    picAdmin.Left = 544
    
    picTrade.Left = picScreen.Left + (picScreen.Width / 2 - picTrade.Width / 2)
    picShop.Left = picScreen.Left + (picScreen.Width / 2 - picShop.Width / 2)
    picBank.Left = picScreen.Left + (picScreen.Width / 2 - picBank.Width / 2)

        picHotbar.Top = 399 + 216
        picEventChat.Top = 442 + 216
        txtMyChat.Top = 565 + 216
        txtChat.Top = 442 + 216
        frmMain.Width = 12530
        frmMain.Height = 12570
        comboresolution.ListIndex = 1
     End If
   
' New Resolution END
End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Socket_DataArrival", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    Call HandleKeyPresses(KeyAscii)

    ' prevents textbox on error ding sound
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
        KeyAscii = 0
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_KeyPress", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    Select Case KeyCode
        Case vbKeyInsert
            If Player(MyIndex).Access > 0 Then
                picAdmin.Visible = Not picAdmin.Visible
            End If
    End Select
    
    ' hotbar
    For i = 1 To MAX_HOTBAR
        If KeyCode = 111 + i Then
            SendHotbarUse i
    ElseIf KeyCode = 96 + i Then
            SendHotbarUse i
        End If
    Next
    
    ' handles delete events
    If KeyCode = vbKeyDelete Then
        If InMapEditor Then DeleteEvent CurX, CurY
    End If
    
    ' handles copy + pasting events
    If KeyCode = vbKeyC Then
        If ControlDown Then
            If InMapEditor Then
                CopyEvent_Map CurX, CurY
            End If
        End If
    End If
    If KeyCode = vbKeyV Then
        If ControlDown Then
            If InMapEditor Then
                PasteEvent_Map CurX, CurY
            End If
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_KeyUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtMyChat_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    MyText = txtMyChat
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtMyChat_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtChat_GotFocus()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    SetFocusOnChat
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtChat_GotFocus", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ***************
' ** Inventory **
' ***************
Private Sub picInventory_DblClick()
    Dim InvNum As Long
    Dim Value As Long
    Dim multiplier As Double
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    DragInvSlotNum = 0
    InvNum = IsInvItem(InvX, InvY)

    If InvNum <> 0 Then
    
        ' are we in a shop?
        If InShop > 0 Then
            Select Case ShopAction
                Case 0 ' nothing, give value
                    multiplier = Shop(InShop).BuyRate / 100
                    Value = Item(GetPlayerInvItemNum(MyIndex, InvNum)).Price * multiplier
                    If Value > 0 Then
                        AddText "You can sell this item for " & Value & " gold.", White
                    Else
                        AddText "The shop does not want this item.", BrightRed
                    End If
                Case 2 ' 2 = sell
                    SellItem InvNum
            End Select
            
            Exit Sub
        End If
        
        ' in bank?
        If InBank Then
            If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ItemCurrency Then
                CurrencyMenu = 2 ' deposit
                lblCurrency.Caption = "How many do you want to deposit?"
                tmpCurrencyItem = InvNum
                txtCurrency.text = vbNullString
                picCurrency.Visible = True
                txtCurrency.SetFocus
                Exit Sub
            End If
                
            Call DepositItem(InvNum, 0)
            Exit Sub
        End If
        
        ' in trade?
        If InTrade > 0 Then
            ' exit out if we're offering that item
            For i = 1 To MAX_INV
                If TradeYourOffer(i).num = InvNum Then
                    ' is currency?
                    If Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)).Type = ItemCurrency Then
                        ' only exit out if we're offering all of it
                        If TradeYourOffer(i).Value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(i).num) Then
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            Next
            
            If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ItemCurrency Then
                CurrencyMenu = 4 ' offer in trade
                lblCurrency.Caption = "How many do you want to trade?"
                tmpCurrencyItem = InvNum
                txtCurrency.text = vbNullString
                picCurrency.Visible = True
                txtCurrency.SetFocus
                Exit Sub
            End If
            
            Call TradeItem(InvNum, 0)
            Exit Sub
        End If
        
        ' use item if not doing anything else
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ItemNone Then Exit Sub
        Call SendUseItem(InvNum)
        Call CheckItems
        Exit Sub
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picInventory_DblClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function IsEqItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim tempRec As RECT
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    IsEqItem = 0

    For i = 1 To Equipment.Equipment_Count - 1

        If GetPlayerEquipment(MyIndex, i) > 0 And GetPlayerEquipment(MyIndex, i) <= MAX_ITEMS Then

            With tempRec
                .Top = EqTop + ((EqOffsetY + 32) * ((i - 1) \ EqColumns))
                .Bottom = .Top + PIC_Y
                .Left = EqLeft + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsEqItem = i
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsEqItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Function IsInvItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim tempRec As RECT
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    IsInvItem = 0

    For i = 1 To MAX_INV

        If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then

            With tempRec
                .Top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsInvItem = i
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsInvItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Function IsPlayerSpell(ByVal X As Single, ByVal Y As Single) As Long
    Dim tempRec As RECT
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    IsPlayerSpell = 0

    For i = 1 To MAX_PLAYER_SPELLS

        If PlayerSpells(i) > 0 And PlayerSpells(i) <= MAX_SPELLS Then

            With tempRec
                .Top = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                .Bottom = .Top + PIC_Y
                .Left = SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsPlayerSpell = i
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsPlayerSpell", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Function IsTradeItem(ByVal X As Single, ByVal Y As Single, ByVal Yours As Boolean) As Long
    Dim tempRec As RECT
    Dim i As Long
    Dim itemnum As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    IsTradeItem = 0

    For i = 1 To MAX_INV
    
        If Yours Then
            itemnum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)
        Else
            itemnum = TradeTheirOffer(i).num
        End If

        If itemnum > 0 And itemnum <= MAX_ITEMS Then

            With tempRec
                .Top = InvTop - 24 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsTradeItem = i
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsTradeItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Sub picInventory_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim InvNum As Long
    Dim i As Integer
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    InvNum = IsInvItem(X, Y)

    If Button = 1 Then
        If InvNum <> 0 Then
            If InTrade > 0 Then Exit Sub
            If InBank Or InShop Then Exit Sub
            DragInvSlotNum = InvNum
            frmMain.picTempInv.Refresh
        End If

    ElseIf Button = 2 Then
        If Not InBank And Not InShop And Not InTrade > 0 Then
            If InvNum <> 0 Then
                If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ItemCurrency Then
                    If GetPlayerInvItemValue(MyIndex, InvNum) > 0 Then
                        CurrencyMenu = 1 ' drop
                        lblCurrency.Caption = "How many do you want to drop?"
                        tmpCurrencyItem = InvNum
                        txtCurrency.text = vbNullString
                        picCurrency.Visible = True
                        txtCurrency.SetFocus
                    End If
                Else
                    Call SendDropItem(InvNum, 0)
               End If
             End If
          End If
    SetFocusOnChat
     End If
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picInventory_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picInventory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim InvNum As Long
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    InvX = X
    InvY = Y

    If DragInvSlotNum > 0 Then
        If InTrade > 0 Then Exit Sub
        If InBank Or InShop Then Exit Sub
        With frmMain.picTempInv
            .Top = Y + picInventory.Top
            .Left = X + picInventory.Left
            .Visible = True
            .ZOrder (0)
        End With
    Else
        InvNum = IsInvItem(X, Y)

        If InvNum <> 0 Then
            ' exit out if we're offering that item
            If InTrade Then
                For i = 1 To MAX_INV
                    If TradeYourOffer(i).num = InvNum Then
                        ' is currency?
                        If Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)).Type = ItemCurrency Then
                            ' only exit out if we're offering all of it
                            If TradeYourOffer(i).Value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(i).num) Then
                                Exit Sub
                            End If
                        Else
                            Exit Sub
                        End If
                    End If
                Next
            End If
            X = X + picInventory.Left - picItemDesc.Width - 1
            Y = Y + picInventory.Top - picItemDesc.Height - 1
            UpdateDescWindow GetPlayerInvItemNum(MyIndex, InvNum), X, Y
            LastItemDesc = GetPlayerInvItemNum(MyIndex, InvNum) ' set it so you don't re-set values
            Exit Sub
        End If
    End If

    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picInventory_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picInventory_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim rec_pos As RECT
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If InTrade > 0 Then Exit Sub
    If InBank Or InShop Then Exit Sub

    If DragInvSlotNum > 0 Then
        ' drag + drop
        For i = 1 To MAX_INV
            With rec_pos
                .Top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then '
                    If DragInvSlotNum <> i Then
                        SendChangeInvSlots DragInvSlotNum, i
                        Exit For
                    End If
                End If
            End If
        Next
        ' hotbar
        For i = 1 To MAX_HOTBAR
            With rec_pos
                .Top = picHotbar.Top - picInventory.Top
                .Left = picHotbar.Left - picInventory.Left + (HotbarOffsetX * (i - 1)) + (32 * (i - 1))
                .Right = .Left + 32
                .Bottom = picHotbar.Top - picInventory.Top + 32
            End With
            
            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                    SendHotbarChange 1, DragInvSlotNum, i
                    DragInvSlotNum = 0
                    picTempInv.Visible = False
                    Exit Sub
                End If
            End If
        Next
    End If

    DragInvSlotNum = 0
    picTempInv.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picInventory_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picItemDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    picItemDesc.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picItemDesc_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' *****************
' ** Char window **
' *****************

Private Sub picCharacter_Click()
    Dim EqNum As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    EqNum = IsEqItem(EqX, EqY)

    If EqNum <> 0 Then
        SendUnequip EqNum
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picCharacter_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picCharacter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim EqNum As Long
    Dim X2 As Long, Y2 As Long
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    EqX = X
    EqY = Y
    EqNum = IsEqItem(X, Y)

    If EqNum <> 0 Then
        Y2 = Y + picCharacter.Top - frmMain.picItemDesc.Height - 1
        X2 = X + picCharacter.Left - frmMain.picItemDesc.Width - 1
        UpdateDescWindow GetPlayerEquipment(MyIndex, EqNum), X2, Y2
        LastItemDesc = GetPlayerEquipment(MyIndex, EqNum) ' set it so you don't re-set values
        Exit Sub
    End If

    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picCharacter_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ****************
' ** Admin Menu **
' ****************

Private Sub cmdALoc_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < RankMapper Then
        
        Exit Sub
    End If
    
    BLoc = Not BLoc
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdALoc_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAMap_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < RankMapper Then
        
        Exit Sub
    End If
    
    SendRequestEditMap
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAMap_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarp2Me_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < RankMapper Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.text)) Then
        Exit Sub
    End If

    WarpToMe Trim$(txtAName.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAWarp2Me_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarpMe2_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < RankMapper Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.text)) Then
        Exit Sub
    End If

    WarpMeTo Trim$(txtAName.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAWarpMe2_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarp_Click()
Dim n As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < RankMapper Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAMap.text)) < 1 Then
        Exit Sub
    End If

    If Not IsNumeric(Trim$(txtAMap.text)) Then
        Exit Sub
    End If

    n = CLng(Trim$(txtAMap.text))

    ' Check to make sure its a valid map #
    If n > 0 And n <= MAX_MAPS Then
        Call WarpTo(n)
    Else
        Call AddText("Invalid map number.", Red)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAWarp_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdASprite_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < RankMapper Then
        
        Exit Sub
    End If

    If Len(Trim$(txtASprite.text)) < 1 Then
        Exit Sub
    End If

    If Not IsNumeric(Trim$(txtASprite.text)) Then
        Exit Sub
    End If

    SendSetSprite CLng(Trim$(txtASprite.text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdASprite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAMapReport_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < RankMapper Then
        
        Exit Sub
    End If

    SendMapReport
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAMapReport_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdARespawn_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < RankMapper Then
        
        Exit Sub
    End If
    
    SendMapRespawn
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdARespawn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdABan_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < RankMapper Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    SendBan Trim$(txtAName.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdABan_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAItem_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < RankDeveloper Then
        
        Exit Sub
    End If

    SendRequestEditItem
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAItem_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdANpc_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < RankDeveloper Then
        
        Exit Sub
    End If

    SendRequestEditNpc
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdANpc_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAResource_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < RankDeveloper Then
        
        Exit Sub
    End If

    SendRequestEditResource
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAResource_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAShop_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < RankDeveloper Then
        
        Exit Sub
    End If

    SendRequestEditShop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAShop_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdASpell_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < RankDeveloper Then
        
        Exit Sub
    End If

    SendRequestEditSpell
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdASpell_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAAccess_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < RankAdministrator Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 2 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.text)) Or Not IsNumeric(Trim$(txtAAccess.text)) Then
        Exit Sub
    End If

    SendSetAccess Trim$(txtAName.text), CLng(Trim$(txtAAccess.text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAAccess_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdADestroy_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < RankAdministrator Then
        
        Exit Sub
    End If

    SendBanDestroy
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdADestroy_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdASpawn_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < RankAdministrator Then
        
        Exit Sub
    End If
    
    SendSpawnItem scrlAItem.Value, scrlAAmount.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdASpawn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' bank
Private Sub picBank_DblClick()
Dim bankNum As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    DragBankSlotNum = 0

    bankNum = IsBankItem(BankX, BankY)
    If bankNum <> 0 Then
         If GetBankItemNum(bankNum) = ItemNone Then Exit Sub
         
             If Item(GetBankItemNum(bankNum)).Type = ItemCurrency Then
                CurrencyMenu = 3 ' withdraw
                lblCurrency.Caption = "How many do you want to withdraw?"
                tmpCurrencyItem = bankNum
                txtCurrency.text = vbNullString
                picCurrency.Visible = True
                txtCurrency.SetFocus
                Exit Sub
            End If
            
         WithdrawItem bankNum, 0
         Exit Sub
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBank_DlbClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picBank_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim bankNum As Long
                        
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    bankNum = IsBankItem(X, Y)
    
    If bankNum <> 0 Then
        
        If Button = 1 Then
            DragBankSlotNum = bankNum
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBank_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picBank_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim rec_pos As RECT
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

' TODO : Add sub to change bankslots client side first so there's no delay in switching
    If DragBankSlotNum > 0 Then
        For i = 1 To MAX_BANK
            With rec_pos
                .Top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                .Bottom = .Top + PIC_Y
                .Left = BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                    If DragBankSlotNum <> i Then
                        ChangeBankSlots DragBankSlotNum, i
                        Exit For
                    End If
                End If
            End If
        Next
    End If

    DragBankSlotNum = 0
    picTempBank.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBank_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picBank_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim bankNum As Long, itemnum As Long, ItemType As Long
Dim X2 As Long, Y2 As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    BankX = X
    BankY = Y
    
    If DragBankSlotNum > 0 Then
        With frmMain.picTempBank
            .Top = Y + picBank.Top
            .Left = X + picBank.Left
            .Visible = True
            .ZOrder (0)
        End With
    Else
        bankNum = IsBankItem(X, Y)
        
        If bankNum <> 0 Then
            
            X2 = X + picBank.Left + 1
            Y2 = Y + picBank.Top + 1
            UpdateDescWindow Bank.Item(bankNum).num, X2, Y2
            Exit Sub
        End If
    End If
    
    frmMain.picItemDesc.Visible = False
    LastBankDesc = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBank_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function IsBankItem(ByVal X As Single, ByVal Y As Single) As Long
Dim tempRec As RECT
Dim i As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    IsBankItem = 0
    
    For i = 1 To MAX_BANK
        If GetBankItemNum(i) > 0 And GetBankItemNum(i) <= MAX_ITEMS Then
        
            With tempRec
                .Top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                .Bottom = .Top + PIC_Y
                .Left = BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                .Right = .Left + PIC_X
            End With
            
            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    
                    IsBankItem = i
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsBankItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function


'ALATAR

'QuestDialogue:

Private Sub lblQuestAccept_Click()
    PlayerHandleQuest CLng(lblQuestAccept.Tag), 1
    picQuestDialogue.Visible = False
    lblQuestAccept.Visible = False
    lblQuestAccept.Tag = vbNullString
    lblQuestSay = "-"
    RefreshQuestLog
End Sub

Private Sub lblQuestExtra_Click()
    RunQuestDialogueExtraLabel
End Sub

Private Sub lblQuestClose_Click()
    picQuestDialogue.Visible = False
    lblQuestExtra.Visible = False
    lblQuestAccept.Visible = False
    lblQuestAccept.Tag = vbNullString
    lblQuestSay = "-"
End Sub

'QuestLog:
'Private Sub picQuestButton_Click()
'    'Need to be replaced with imgButton(X) and a proper image
'    UpdateQuestLog
'    picQuestLog.Visible = Not picQuestLog.Visible
'    PlaySound Sound_ButtonClick
'End Sub

Private Sub imgQuestButton_Click(Index As Integer)
    If Trim$(lstQuestLog.text) = vbNullString Then Exit Sub
    LoadQuestlogBox Index
End Sub

' New Resolution BEGIN
Private Sub ComboResolution_Click()
Dim resWidth, resHeight As Long
Dim res() As String

    res() = Split(comboresolution.List(comboresolution.ListIndex), "x")

    Direct3D_Window.BackBufferWidth = res(0)
    Direct3D_Window.BackBufferHeight = res(1)

    MAX_MAPX = res(0) / 32 - 1
    MAX_MAPY = res(1) / 32 - 1
    
    HalfX = ((MAX_MAPX + 1) / 2) * PIC_X
    HalfY = ((MAX_MAPY + 1) / 2) * PIC_Y
    ScreenX = (MAX_MAPX + 1) * PIC_X
    ScreenY = (MAX_MAPY + 1) * PIC_Y
    StartXValue = ((MAX_MAPX + 1) / 2)
    StartYValue = ((MAX_MAPY + 1) / 2)
    EndXValue = (MAX_MAPX + 1) + 1
    EndYValue = (MAX_MAPY + 1) + 1
    
    ' Change Background
    frmMain.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\main" & comboresolution.ListIndex & ".jpg")
    
    picScreen.Width = res(0)
    picScreen.Height = res(1)
    
    If (comboresolution.ListIndex = 0) Then
        picHotbar.Top = 399
        picEventChat.Top = 442
        txtMyChat.Top = 565
        txtChat.Top = 442
        frmMain.Width = 11865
        frmMain.Height = 9315
    ElseIf (comboresolution.ListIndex = 1) Then
        picHotbar.Top = 399 + 216
        picEventChat.Top = 442 + 216
        txtMyChat.Top = 565 + 216
        txtChat.Top = 442 + 216
        

    End If
End Sub
' New Resolution END
Private Sub txtMyChat_Click()
ChatFocus = True
End Sub
