VERSION 5.00
Begin VB.Form frmEditor_Item 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9735
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_Item.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   561
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   649
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Fraprojectile 
      Caption         =   "Projectiles"
      Height          =   1575
      Left            =   3360
      TabIndex        =   131
      Top             =   6240
      Visible         =   0   'False
      Width           =   5415
      Begin VB.HScrollBar Scrolammo 
         Height          =   255
         Left            =   3960
         TabIndex        =   141
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox ammoreq 
         Caption         =   "Req Ammo?"
         Height          =   255
         Left            =   720
         TabIndex        =   140
         Top             =   1080
         Width           =   1335
      End
      Begin VB.HScrollBar scrlProjectilePic 
         Height          =   255
         Left            =   960
         Max             =   99
         TabIndex        =   135
         Top             =   240
         Width           =   1095
      End
      Begin VB.HScrollBar scrlProjectileRange 
         Height          =   255
         Left            =   960
         Max             =   99
         TabIndex        =   134
         Top             =   720
         Width           =   1095
      End
      Begin VB.HScrollBar scrlProjectileSpeed 
         Height          =   255
         Left            =   3960
         Max             =   99
         TabIndex        =   133
         Top             =   240
         Width           =   1215
      End
      Begin VB.HScrollBar scrlProjectileDamage 
         Height          =   255
         Left            =   3960
         Max             =   99
         TabIndex        =   132
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Lblammo 
         Caption         =   "Ammo:0"
         Height          =   375
         Left            =   2520
         TabIndex        =   142
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblProjectilePic 
         Caption         =   "pic:0"
         Height          =   255
         Left            =   240
         TabIndex        =   139
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblProjectileRange 
         Caption         =   "Range:0"
         Height          =   255
         Left            =   240
         TabIndex        =   138
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblProjectileSpeed 
         Caption         =   "Speed:0"
         Height          =   255
         Left            =   2640
         TabIndex        =   137
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblProjectileDamage 
         Caption         =   "Damage:0"
         Height          =   375
         Left            =   2520
         TabIndex        =   136
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Frarecipe 
      Caption         =   "Recipe"
      Height          =   3135
      Left            =   3360
      TabIndex        =   118
      Top             =   4800
      Visible         =   0   'False
      Width           =   6375
      Begin VB.HScrollBar scrlEnchantReq 
         Height          =   255
         Left            =   4320
         Max             =   99
         TabIndex        =   161
         Top             =   2280
         Width           =   735
      End
      Begin VB.CheckBox chkEN 
         Caption         =   "Enchant"
         Height          =   255
         Left            =   5280
         TabIndex        =   160
         Top             =   2280
         Width           =   1095
      End
      Begin VB.HScrollBar scrlAlchemyReq 
         Height          =   255
         Left            =   4320
         Max             =   99
         TabIndex        =   158
         Top             =   1920
         Width           =   735
      End
      Begin VB.CheckBox chkAL 
         Caption         =   " Alchemy"
         Height          =   255
         Left            =   5280
         TabIndex        =   157
         Top             =   1920
         Width           =   1095
      End
      Begin VB.HScrollBar scrlSmithReq 
         Height          =   255
         Left            =   4320
         Max             =   99
         TabIndex        =   155
         Top             =   1560
         Width           =   735
      End
      Begin VB.CheckBox chksm 
         Caption         =   "Smith"
         Height          =   255
         Left            =   5280
         TabIndex        =   154
         Top             =   1560
         Width           =   975
      End
      Begin VB.HScrollBar Scrlamount 
         Height          =   255
         Left            =   1200
         Max             =   99
         TabIndex        =   143
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txtcrftxp 
         Height          =   375
         Left            =   4560
         TabIndex        =   125
         Top             =   1080
         Width           =   1215
      End
      Begin VB.HScrollBar scrlResult 
         Height          =   255
         Left            =   480
         TabIndex        =   122
         Top             =   2160
         Width           =   1695
      End
      Begin VB.HScrollBar scrlItemnum 
         Height          =   255
         Left            =   480
         Max             =   99
         Min             =   1
         TabIndex        =   121
         Top             =   1320
         Value           =   1
         Width           =   1815
      End
      Begin VB.HScrollBar scrlItem1 
         Height          =   255
         Left            =   480
         Max             =   99
         TabIndex        =   120
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox cmbCToolReq 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":3332
         Left            =   4080
         List            =   "frmEditor_Item.frx":3342
         TabIndex        =   119
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblEnchants 
         Caption         =   "1"
         Height          =   255
         Left            =   3120
         TabIndex        =   162
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblAlchemy 
         Caption         =   "1"
         Height          =   255
         Left            =   3120
         TabIndex        =   159
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblSmith 
         Caption         =   "1"
         Height          =   255
         Left            =   3120
         TabIndex        =   156
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Left            =   1680
         TabIndex        =   153
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblamount 
         Caption         =   "Amount:0"
         Height          =   375
         Left            =   240
         TabIndex        =   144
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Crafting Exp:"
         Height          =   255
         Left            =   3480
         TabIndex        =   128
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Tool Required:"
         Height          =   255
         Left            =   3960
         TabIndex        =   127
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblResult 
         Caption         =   "Result"
         Height          =   255
         Left            =   360
         TabIndex        =   126
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lblitemnum 
         Caption         =   "Item 2"
         Height          =   255
         Left            =   240
         TabIndex        =   124
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblItem1 
         Caption         =   "Item 1"
         Height          =   255
         Left            =   240
         TabIndex        =   123
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame fraEquipment 
      Caption         =   "Equipment Data"
      Height          =   3135
      Left            =   3360
      TabIndex        =   32
      Top             =   4680
      Visible         =   0   'False
      Width           =   6255
      Begin VB.HScrollBar ScrlDagPdoll 
         Height          =   255
         Left            =   1560
         TabIndex        =   151
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CheckBox ChkDagger 
         Caption         =   "Dagger"
         Height          =   180
         Left            =   240
         TabIndex        =   150
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CheckBox chkTwoh 
         Caption         =   "Two Handed"
         Height          =   255
         Left            =   240
         TabIndex        =   149
         Top             =   2280
         Width           =   1215
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   6
         LargeChange     =   10
         Left            =   4920
         Max             =   255
         TabIndex        =   145
         Top             =   1560
         Width           =   855
      End
      Begin VB.ComboBox cmbCTool 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":337B
         Left            =   1200
         List            =   "frmEditor_Item.frx":3388
         TabIndex        =   129
         Top             =   1920
         Width           =   1695
      End
      Begin VB.PictureBox picPaperdoll 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   5640
         ScaleHeight     =   56
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   58
         Top             =   2040
         Width           =   480
      End
      Begin VB.HScrollBar scrlPaperdoll 
         Height          =   255
         Left            =   4200
         TabIndex        =   57
         Top             =   2640
         Width           =   1335
      End
      Begin VB.HScrollBar scrlSpeed 
         Height          =   255
         LargeChange     =   100
         Left            =   4560
         Max             =   3000
         Min             =   100
         SmallChange     =   100
         TabIndex        =   40
         Top             =   840
         Value           =   100
         Width           =   1575
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   3000
         Max             =   255
         TabIndex        =   39
         Top             =   1560
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   38
         Top             =   1560
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   4680
         Max             =   255
         TabIndex        =   37
         Top             =   1200
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   3000
         Max             =   255
         TabIndex        =   36
         Top             =   1200
         Width           =   855
      End
      Begin VB.HScrollBar scrlDamage 
         Height          =   255
         LargeChange     =   10
         Left            =   1320
         Max             =   255
         TabIndex        =   35
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox cmbTool 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":33AB
         Left            =   1320
         List            =   "frmEditor_Item.frx":33BB
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   360
         Width           =   4815
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   33
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblDagPdoll 
         Caption         =   "Pdoll"
         Height          =   255
         Left            =   1560
         TabIndex        =   152
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Critl: 0"
         Height          =   180
         Index           =   6
         Left            =   4080
         TabIndex        =   146
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   690
      End
      Begin VB.Label Label7 
         Caption         =   "Crafting tool:"
         Height          =   255
         Left            =   120
         TabIndex        =   130
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblPaperdoll 
         AutoSize        =   -1  'True
         Caption         =   "Paperdoll: 0"
         Height          =   180
         Left            =   3000
         TabIndex        =   56
         Top             =   2670
         Width           =   915
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         Caption         =   "Speed: 0.1 sec"
         Height          =   180
         Left            =   3240
         TabIndex        =   48
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   1140
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Will: 0"
         Height          =   180
         Index           =   5
         Left            =   2160
         TabIndex        =   47
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   630
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   46
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   615
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Int: 0"
         Height          =   180
         Index           =   3
         Left            =   3960
         TabIndex        =   45
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   585
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ End: 0"
         Height          =   180
         Index           =   2
         Left            =   2160
         TabIndex        =   44
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   645
      End
      Begin VB.Label lblDamage 
         AutoSize        =   -1  'True
         Caption         =   "Damage: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   43
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Object Tool:"
         Height          =   180
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   585
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Info"
      Height          =   3375
      Left            =   3360
      TabIndex        =   17
      Top             =   120
      Width           =   6255
      Begin VB.CheckBox ChkProjectile 
         Alignment       =   1  'Right Justify
         Caption         =   "Projectile?"
         Height          =   255
         Left            =   2880
         TabIndex        =   163
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "skills?"
         Height          =   255
         Left            =   5280
         TabIndex        =   76
         Top             =   3000
         Width           =   855
      End
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         LargeChange     =   10
         Left            =   4200
         Max             =   99
         TabIndex        =   74
         Top             =   2760
         Width           =   1935
      End
      Begin VB.HScrollBar scrlAccessReq 
         Height          =   255
         Left            =   4200
         Max             =   5
         TabIndex        =   72
         Top             =   2400
         Width           =   1935
      End
      Begin VB.ComboBox cmbClassReq 
         Height          =   300
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   2040
         Width           =   2295
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtDesc 
         Height          =   1455
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   60
         Top             =   1800
         Width           =   2655
      End
      Begin VB.HScrollBar scrlRarity 
         Height          =   255
         Left            =   4200
         Max             =   5
         TabIndex        =   25
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox cmbBind 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":33DC
         Left            =   4200
         List            =   "frmEditor_Item.frx":33E9
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   600
         Width           =   1935
      End
      Begin VB.HScrollBar scrlPrice 
         Height          =   255
         LargeChange     =   100
         Left            =   4200
         Max             =   30000
         TabIndex        =   23
         Top             =   240
         Width           =   1935
      End
      Begin VB.HScrollBar scrlAnim 
         Height          =   255
         Left            =   5040
         Max             =   5
         TabIndex        =   22
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":3412
         Left            =   120
         List            =   "frmEditor_Item.frx":3443
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtName 
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   240
         Width           =   2055
      End
      Begin VB.HScrollBar scrlPic 
         Height          =   255
         Left            =   840
         Max             =   255
         TabIndex        =   19
         Top             =   600
         Width           =   1335
      End
      Begin VB.PictureBox picItem 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2280
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   18
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         Caption         =   "Level req: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   75
         Top             =   2760
         Width           =   900
      End
      Begin VB.Label lblAccessReq 
         AutoSize        =   -1  'True
         Caption         =   "Access Req: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   73
         Top             =   2400
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Class Req:"
         Height          =   180
         Left            =   2880
         TabIndex        =   71
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label Label4 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   2880
         TabIndex        =   68
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblRarity 
         AutoSize        =   -1  'True
         Caption         =   "Rarity: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   31
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Bind Type:"
         Height          =   180
         Left            =   2880
         TabIndex        =   30
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         Caption         =   "Price: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   29
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Anim: None"
         Height          =   180
         Left            =   2880
         TabIndex        =   28
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         Caption         =   "Pic: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Requirements"
      Height          =   975
      Left            =   3360
      TabIndex        =   6
      Top             =   3600
      Width           =   6255
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   6
         LargeChange     =   10
         Left            =   5160
         Max             =   255
         TabIndex        =   147
         Top             =   600
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   720
         Max             =   255
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   2880
         Max             =   255
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   5160
         Max             =   255
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   720
         Max             =   255
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   2880
         Max             =   255
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Crit: 0"
         Height          =   180
         Index           =   6
         Left            =   4560
         TabIndex        =   148
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "End: 0"
         Height          =   180
         Index           =   2
         Left            =   2280
         TabIndex        =   15
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Int: 0"
         Height          =   180
         Index           =   3
         Left            =   4560
         TabIndex        =   14
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   465
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Will: 0"
         Height          =   180
         Index           =   5
         Left            =   2280
         TabIndex        =   12
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   7920
      Width           =   2895
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Item List"
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.Frame Frame4 
         Caption         =   "Profiencies"
         Height          =   6855
         Left            =   120
         TabIndex        =   77
         Top             =   360
         Visible         =   0   'False
         Width           =   2895
         Begin VB.HScrollBar scrlcraftingreq 
            Height          =   255
            Left            =   1440
            Max             =   99
            TabIndex        =   117
            Top             =   6360
            Width           =   1095
         End
         Begin VB.HScrollBar scrlfishingreq 
            Height          =   255
            Left            =   1440
            Max             =   99
            TabIndex        =   116
            Top             =   5880
            Width           =   1095
         End
         Begin VB.HScrollBar scrlwoodcuttingreq 
            Height          =   255
            Left            =   1440
            Max             =   99
            TabIndex        =   115
            Top             =   5400
            Width           =   1095
         End
         Begin VB.HScrollBar scrlminingreq 
            Height          =   255
            Left            =   1440
            Max             =   99
            TabIndex        =   114
            Top             =   4920
            Width           =   1095
         End
         Begin VB.HScrollBar scrllightarmorreq 
            Height          =   255
            Left            =   1440
            Max             =   99
            TabIndex        =   113
            Top             =   4440
            Width           =   1095
         End
         Begin VB.HScrollBar scrlheavyarmorreq 
            Height          =   255
            Left            =   1440
            Max             =   99
            TabIndex        =   112
            Top             =   3840
            Width           =   1095
         End
         Begin VB.HScrollBar scrllycanthropyreq 
            Height          =   255
            Left            =   1440
            Max             =   99
            TabIndex        =   111
            Top             =   3360
            Width           =   1095
         End
         Begin VB.HScrollBar scrlconvictionreq 
            Height          =   255
            Left            =   1440
            Max             =   99
            TabIndex        =   110
            Top             =   2880
            Width           =   1095
         End
         Begin VB.HScrollBar scrlmagicreq 
            Height          =   255
            Left            =   1440
            Max             =   99
            TabIndex        =   109
            Top             =   2400
            Width           =   1095
         End
         Begin VB.HScrollBar scrlrangereq 
            Height          =   255
            Left            =   1440
            Max             =   99
            TabIndex        =   108
            Top             =   1920
            Width           =   1095
         End
         Begin VB.HScrollBar scrldaggersreq 
            Height          =   255
            Left            =   1440
            Max             =   99
            TabIndex        =   107
            Top             =   1440
            Width           =   1095
         End
         Begin VB.CheckBox chkcr 
            Caption         =   "Check2"
            Height          =   255
            Left            =   1080
            TabIndex        =   106
            Top             =   6360
            Width           =   255
         End
         Begin VB.CheckBox chkfi 
            Caption         =   "Check2"
            Height          =   255
            Left            =   1080
            TabIndex        =   105
            Top             =   5880
            Width           =   255
         End
         Begin VB.CheckBox chkwc 
            Caption         =   "Check2"
            Height          =   180
            Left            =   1080
            TabIndex        =   104
            Top             =   5400
            Width           =   135
         End
         Begin VB.CheckBox chkmi 
            Caption         =   "Check2"
            Height          =   180
            Left            =   1080
            TabIndex        =   103
            Top             =   5040
            Width           =   255
         End
         Begin VB.CheckBox chkla 
            Caption         =   "Check2"
            Height          =   255
            Left            =   1080
            TabIndex        =   102
            Top             =   4440
            Width           =   135
         End
         Begin VB.CheckBox chkha 
            Caption         =   "Check2"
            Height          =   255
            Left            =   1200
            TabIndex        =   101
            Top             =   3840
            Width           =   255
         End
         Begin VB.CheckBox chkly 
            Caption         =   "Check2"
            Height          =   255
            Left            =   1200
            TabIndex        =   100
            Top             =   3360
            Width           =   255
         End
         Begin VB.CheckBox chkco 
            Caption         =   "Check2"
            Height          =   255
            Left            =   1080
            TabIndex        =   99
            Top             =   2880
            Width           =   255
         End
         Begin VB.CheckBox chkma 
            Caption         =   "Check2"
            Height          =   255
            Left            =   1080
            TabIndex        =   98
            Top             =   2400
            Width           =   255
         End
         Begin VB.CheckBox chkra 
            Caption         =   "Check2"
            Height          =   180
            Left            =   1080
            TabIndex        =   97
            Top             =   1920
            Width           =   255
         End
         Begin VB.CheckBox chkda 
            Caption         =   "Check2"
            Height          =   180
            Left            =   1080
            TabIndex        =   96
            Top             =   1440
            Width           =   255
         End
         Begin VB.HScrollBar scrlAxesReq 
            Height          =   255
            Left            =   1440
            Max             =   99
            TabIndex        =   83
            Top             =   960
            Width           =   1095
         End
         Begin VB.CheckBox chkAX 
            Height          =   375
            Left            =   1080
            TabIndex        =   82
            Top             =   960
            Width           =   255
         End
         Begin VB.HScrollBar scrlswordreq 
            Height          =   255
            Left            =   1440
            Max             =   99
            TabIndex        =   80
            Top             =   600
            Width           =   1095
         End
         Begin VB.CheckBox chksw 
            Caption         =   "Check2"
            Height          =   255
            Left            =   1080
            TabIndex        =   78
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lblcrafting 
            Caption         =   "crafting"
            Height          =   255
            Left            =   120
            TabIndex        =   95
            Top             =   6360
            Width           =   975
         End
         Begin VB.Label lblfishing 
            Caption         =   "fishing"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   5880
            Width           =   855
         End
         Begin VB.Label lblwoodcutting 
            Caption         =   "woodcutting"
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   5400
            Width           =   1095
         End
         Begin VB.Label lblmining 
            Caption         =   "mining"
            Height          =   255
            Left            =   240
            TabIndex        =   92
            Top             =   4920
            Width           =   735
         End
         Begin VB.Label lbllightarmor 
            Caption         =   "lightarmor"
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   4440
            Width           =   975
         End
         Begin VB.Label lblheavyarmor 
            Caption         =   "Heavyarmor"
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   3840
            Width           =   975
         End
         Begin VB.Label lbllycanthropy 
            Caption         =   "lycanthropy"
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   3360
            Width           =   975
         End
         Begin VB.Label lblconviction 
            Caption         =   "conviction"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label lblmagic 
            Caption         =   "Magic"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label lblrange 
            Caption         =   "Range"
            Height          =   255
            Left            =   120
            TabIndex        =   86
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label lbldaggers 
            Caption         =   "Dagger"
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label lblaxes 
            Caption         =   "Axes"
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lbllvlreqs 
            Caption         =   "Level requirements"
            Height          =   615
            Left            =   1320
            TabIndex        =   81
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label lblswords 
            Caption         =   "Swords"
            Height          =   375
            Left            =   120
            TabIndex        =   79
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.ListBox lstIndex 
         Height          =   7260
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraVitals 
      Caption         =   "Consume Data"
      Height          =   3135
      Left            =   3360
      TabIndex        =   49
      Top             =   4680
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CheckBox chkInstant 
         Caption         =   "Instant Cast?"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   2760
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.HScrollBar scrlCastSpell 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   65
         Top             =   2400
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddExp 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   63
         Top             =   1800
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddMP 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   61
         Top             =   1200
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddHp 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   50
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label lblCastSpell 
         AutoSize        =   -1  'True
         Caption         =   "Cast Spell: None"
         Height          =   180
         Left            =   120
         TabIndex        =   66
         Top             =   2160
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblAddExp 
         AutoSize        =   -1  'True
         Caption         =   "Add Exp: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   64
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   840
      End
      Begin VB.Label lblAddMP 
         AutoSize        =   -1  'True
         Caption         =   "Add MP: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   62
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   795
      End
      Begin VB.Label lblAddHP 
         AutoSize        =   -1  'True
         Caption         =   "Add HP: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   51
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   780
      End
   End
   Begin VB.Frame fraSpell 
      Caption         =   "Spell Data"
      Height          =   1215
      Left            =   3360
      TabIndex        =   52
      Top             =   4680
      Visible         =   0   'False
      Width           =   3735
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   1080
         Max             =   255
         Min             =   1
         TabIndex        =   53
         Top             =   720
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label lblSpellName 
         AutoSize        =   -1  'True
         Caption         =   "Name: None"
         Height          =   180
         Left            =   240
         TabIndex        =   55
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lblSpell 
         AutoSize        =   -1  'True
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   240
         TabIndex        =   54
         Top             =   720
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmEditor_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CraftIndex As Long
Private LastIndex As Long
Private RecipeIndex As Long

Private Sub ammoreq_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).ammoreq = ammoreq.Value
End Sub

Private Sub Check1_Click()
    If Check1.Value Then
        Frame4.Visible = True
    Else
        Frame4.Visible = False
    End If
End Sub


Private Sub chkAX_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Axes = chkAX.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkax_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chkco_Click()
     ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Conviction = chkco.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkMA_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chkcr_Click()
     ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Crafting = chkcr.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkCR_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chkda_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Daggers = chkda.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkDA_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ChkDagger_Click()
If options.Debug = 1 Then On Error GoTo errorhandler

If ChkDagger.Value = 0 Then
Item(EditorIndex).isDagger = False
Else
Item(EditorIndex).isDagger = True
Item(EditorIndex).isTwoHanded = False
chkTwoh.Value = 0
End If

' Error handler
Exit Sub
errorhandler:
HandleError "chkTwoh", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub

Private Sub chkfi_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Fishing = chkfi.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkFI_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chkha_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Heavyarmor = chkha.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkHA_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chkla_Click()
     ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).LightArmor = chkla.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkHA_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chkly_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Lycanthropy = chkly.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkly_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chkma_Click()
     ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Magic = chkma.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkMA_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chkmi_Click()
     ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Mining = chkmi.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkHA_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ChkProjectile_Click()
    If ChkProjectile.Value Then
        Fraprojectile.Visible = True
    Else
        Fraprojectile.Visible = False
    End If
End Sub

Private Sub chkra_Click()
     ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Range = chkra.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkRA_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chksw_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Sword = chksw.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkSW_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chkTwoh_Click()
 'If debug mode, handle error then exit out
If options.Debug = 1 Then On Error GoTo errorhandler

If chkTwoh.Value = 0 Then
Item(EditorIndex).isTwoHanded = False
Else
Item(EditorIndex).isTwoHanded = True
Item(EditorIndex).isDagger = False
ChkDagger.Value = 0
End If

' Error handler
Exit Sub
errorhandler:
HandleError "chkTwoh", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub

Private Sub chkwc_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Woodcutting = chkwc.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkWC_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbBind_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).BindType = cmbBind.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbBind_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbClassReq_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).ClassReq = cmbClassReq.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbClassReq_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbCTool_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
     If cmbCTool.ListIndex >= 0 Then
    Item(EditorIndex).Tool = cmbCTool.ListIndex
        Else
        Item(EditorIndex).Tool = "None."
    End If
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Item(EditorIndex).sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Item(EditorIndex).sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbTool_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Data3 = cmbTool.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbTool_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    ClearItem EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ItemEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlPic.max = numitems
    scrlAnim.max = MAX_ANIMATIONS
    scrlPaperdoll.max = NumPaperdolls
    Scrolammo.max = MAX_ITEMS
    CraftIndex = 1
    RecipeIndex = 1
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ItemEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ItemEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    If (cmbType.ListIndex >= ItemWeapon) And (cmbType.ListIndex <= ItemShield) Then
        fraEquipment.Visible = True
        'scrlDamage_Change
    Else
        fraEquipment.Visible = False
    End If

    If cmbType.ListIndex = ItemConsume Then
        fraVitals.Visible = True
        'scrlVitalMod_Change
    Else
        fraVitals.Visible = False
    End If

    If (cmbType.ListIndex = ItemSpell) Then
        fraSpell.Visible = True
    Else
        fraSpell.Visible = False
    End If
    
    If (cmbType.ListIndex = ItemRecipe) Then
        Frarecipe.Visible = True
    Else
        Frarecipe.Visible = False
    End If
    
    Item(EditorIndex).Type = cmbType.ListIndex

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub





Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    ItemEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAccessReq_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblAccessReq.Caption = "Access Req: " & scrlAccessReq.Value
    Item(EditorIndex).AccessReq = scrlAccessReq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAccessReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddHp_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAddHP.Caption = "Add HP: " & scrlAddHp.Value
    Item(EditorIndex).AddHP = scrlAddHp.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddHP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddMp_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAddMP.Caption = "Add MP: " & scrlAddMP.Value
    Item(EditorIndex).AddMP = scrlAddMP.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddMP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddExp_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAddExp.Caption = "Add Exp: " & scrlAddExp.Value
    Item(EditorIndex).AddEXP = scrlAddExp.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddExp_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAmount_Change()
lblAmount.Caption = "Amount: " & scrlAmount.Value
Item(EditorIndex).Amount = scrlAmount.Value
End Sub

Private Sub scrlAnim_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If scrlAnim.Value = 0 Then
        sString = "None"
    Else
        sString = Trim$(Animation(scrlAnim.Value).Name)
    End If
    lblAnim.Caption = "Anim: " & sString
    Item(EditorIndex).Animation = scrlAnim.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnim_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAxesReq_Change()
  ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblAxes.Caption = "Ax: " & scrlAxesReq
    Item(EditorIndex).AxesReq = scrlAxesReq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSwordReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlconvictionreq_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblConviction.Caption = "Co: " & scrlconvictionreq
    Item(EditorIndex).convictionReq = scrlconvictionreq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlConvictionReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlcraftingreq_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblCrafting.Caption = "CR: " & scrlcraftingreq
    Item(EditorIndex).CraftingReq = scrlcraftingreq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlCraftingReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrldaggerreq_Change()
   ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lbldaggers.Caption = "Da: " & scrldaggersreq
    Item(EditorIndex).daggersReq = scrldaggersreq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDaggersReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrldaggersreq_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lbldaggers.Caption = "DA: " & scrldaggersreq
    Item(EditorIndex).daggersReq = scrldaggersreq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDaggersReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScrlDagPdoll_Change()
If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
lblDagPdoll.Caption = "Paperdoll: " & ScrlDagPdoll.Value
Item(EditorIndex).Daggerpdoll = ScrlDagPdoll.Value
End Sub

Private Sub scrlDamage_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblDamage.Caption = "Damage: " & scrlDamage.Value
    Item(EditorIndex).Data2 = scrlDamage.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDamage_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlfishingreq_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblFishing.Caption = "FI: " & scrlfishingreq
    Item(EditorIndex).FishingReq = scrlfishingreq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlFishingReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Private Sub scrlheavyarmorreq_Change()
   ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblHeavyarmor.Caption = "HA: " & scrlheavyarmorreq
    Item(EditorIndex).HeavyArmorReq = scrlheavyarmorreq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlHeavyarmorReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub
Private Sub scrlLevelReq_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblLevelReq.Caption = "Level req: " & scrlLevelReq
    Item(EditorIndex).LevelReq = scrlLevelReq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLevelReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrllightarmorreq_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblLightArmor.Caption = "LA: " & scrllightarmorreq
    Item(EditorIndex).LightarmorReq = scrllightarmorreq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLightArmorReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Private Sub scrllycanthropyreq_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblLycanthropy.Caption = "Ly: " & scrllycanthropyreq
    Item(EditorIndex).lycanthropyReq = scrllycanthropyreq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLycanthropyReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlmagicreq_Change()

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblmagic.Caption = "Ma: " & scrlmagicreq
    Item(EditorIndex).magicReq = scrlmagicreq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMagicReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Private Sub scrlminingreq_Change()
   ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblMining.Caption = "MI: " & scrlminingreq
    Item(EditorIndex).MiningReq = scrlminingreq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMiningReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Private Sub scrlPaperdoll_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPaperdoll.Caption = "Paperdoll: " & scrlPaperdoll.Value
    Item(EditorIndex).Paperdoll = scrlPaperdoll.Value
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPaperdoll_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPic_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPic.Caption = "Pic: " & scrlPic.Value
    Item(EditorIndex).Pic = scrlPic.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPrice_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPrice.Caption = "Price: " & scrlPrice.Value
    Item(EditorIndex).Price = scrlPrice.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPrice_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlrangerreq_Change()

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblRange.Caption = "Ra: " & scrlrangereq
    Item(EditorIndex).RangeReq = scrlrangereq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRangeReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Private Sub scrlrangereq_Change()
     ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblRange.Caption = "RA: " & scrlrangereq
    Item(EditorIndex).RangeReq = scrlrangereq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRangeReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRarity_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblRarity.Caption = "Rarity: " & scrlRarity.Value
    Item(EditorIndex).Rarity = scrlRarity.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRarity_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpeed_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblSpeed.Caption = "Speed: " & scrlSpeed.Value / 1000 & " sec"
    Item(EditorIndex).speed = scrlSpeed.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSpeed_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatBonus_Change(Index As Integer)
Dim text As String

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            text = "+ Str: "
        Case 2
            text = "+ End: "
        Case 3
            text = "+ Int: "
        Case 4
            text = "+ Agi: "
        Case 5
            text = "+ Will: "
        Case 6
            text = "+ CRIT: "
    End Select
            
    lblStatBonus(Index).Caption = text & scrlStatBonus(Index).Value
    Item(EditorIndex).Add_Stat(Index) = scrlStatBonus(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatBonus_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatReq_Change(Index As Integer)
    Dim text As String
    
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            text = "Str: "
        Case 2
            text = "End: "
        Case 3
            text = "Int: "
        Case 4
            text = "Agi: "
        Case 5
            text = "Will: "
        Case 6
            text = "CRIT: "
    End Select
    
    lblStatReq(Index).Caption = text & scrlStatReq(Index).Value
    Item(EditorIndex).Stat_Req(Index) = scrlStatReq(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpell_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    If Len(Trim$(Spell(scrlSpell.Value).Name)) > 0 Then
        lblSpellName.Caption = "Name: " & Trim$(Spell(scrlSpell.Value).Name)
    Else
        lblSpellName.Caption = "Name: None"
    End If
    
    lblSpell.Caption = "Spell: " & scrlSpell.Value
    
    Item(EditorIndex).Data1 = scrlSpell.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSpell_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlswordreq_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblswords.Caption = "SW: " & scrlswordreq
    Item(EditorIndex).SwordsReq = scrlswordreq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSwordReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlwoodcuttingreq_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblWoodcutting.Caption = "WC: " & scrlwoodcuttingreq
    Item(EditorIndex).WoodcuttingReq = scrlwoodcuttingreq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlWoodcuttingReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Private Sub Scrolammo_Change()
If Scrolammo.Value > 0 Then
    Lblammo.Caption = "Weapon: " + Item(Scrolammo.Value).Name
Else
    Lblammo.Caption = "Weapon: None"
End If

Item(EditorIndex).Ammo = Scrolammo.Value
End Sub

Private Sub txtcrftxp_Change()
    If Not Len(txtcrftxp.text) > 0 Then Exit Sub
    If IsNumeric(txtcrftxp.text) Then Item(EditorIndex).crftxp = Val(txtcrftxp.text)
End Sub

Private Sub txtDesc_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    Item(EditorIndex).Desc = txtDesc.text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDesc_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Item(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
' projectile
Private Sub scrlProjectileDamage_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileDamage.Caption = "Damage: " & scrlProjectileDamage.Value
    Item(EditorIndex).ProjecTile.Damage = scrlProjectileDamage.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectilePic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' projectile
Private Sub scrlProjectilePic_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectilePic.Caption = "Pic: " & scrlProjectilePic.Value
    Item(EditorIndex).ProjecTile.Pic = scrlProjectilePic.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectilePic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ProjecTile
Private Sub scrlProjectileRange_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileRange.Caption = "Range: " & scrlProjectileRange.Value
    Item(EditorIndex).ProjecTile.Range = scrlProjectileRange.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectileRange_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' projectile
Private Sub scrlProjectileSpeed_Change()
    ' If debug mode, handle error then exit out
    If options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileSpeed.Caption = "Speed: " & scrlProjectileSpeed.Value
    Item(EditorIndex).ProjecTile.speed = scrlProjectileSpeed.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRarity_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub cmbCToolReq_Click()
If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
Item(EditorIndex).ToolReq = cmbCToolReq.ListIndex
End Sub
Private Sub scrlItem1_Change()
If scrlItem1.Value > 0 Then
lblItem1.Caption = "Item: " & Trim$(Item(scrlItem1.Value).Name)
Else
lblItem1.Caption = "Item: None"
End If
Item(EditorIndex).Recipe(RecipeIndex) = scrlItem1.Value
End Sub
Private Sub scrlItemNum_Change()
RecipeIndex = scrlItemnum.Value
lblitemnum.Caption = "Item: " & RecipeIndex
scrlItem1.Value = Item(EditorIndex).Recipe(RecipeIndex)
End Sub
Private Sub scrlResult_Change()
If scrlResult.Value > 0 Then
lblResult.Caption = "Result: " & Trim$(Item(scrlResult.Value).Name)
Else
lblResult.Caption = "Result: None"
End If

Item(EditorIndex).Data3 = scrlResult.Value
End Sub
Private Sub scrlSmithReq_Change()
' If debug mode, handle error then exit out
If options.Debug = 1 Then On Error GoTo errorhandler

If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
lblSmith.Caption = "SM: " & scrlSmithReq
Item(EditorIndex).SmithReq = scrlSmithReq.Value

' Error handler
Exit Sub
errorhandler:
HandleError "scrlSmithReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
Private Sub scrlEnchantReq_Change()
' If debug mode, handle error then exit out
If options.Debug = 1 Then On Error GoTo errorhandler

If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
lblEnchants.Caption = "EN: " & scrlEnchantReq
Item(EditorIndex).EnchantReq = scrlEnchantReq.Value

' Error handler
Exit Sub
errorhandler:
HandleError "scrlSwordReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
Private Sub scrlAlchemyReq_Change()
' If debug mode, handle error then exit out
If options.Debug = 1 Then On Error GoTo errorhandler

If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
lblAlchemy.Caption = "Alchemy: " & scrlAlchemyReq
Item(EditorIndex).AlchemyReq = scrlAlchemyReq.Value

' Error handler
Exit Sub
errorhandler:
HandleError "scrlAlchemyReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
Private Sub chkAL_Click()
' If debug mode, handle error then exit out
If options.Debug = 1 Then On Error GoTo errorhandler

If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
Item(EditorIndex).Alchemist = chkAL.Value

' Error handler
Exit Sub
errorhandler:
HandleError "chkAL_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
Private Sub chkEN_Click()
' If debug mode, handle error then exit out
If options.Debug = 1 Then On Error GoTo errorhandler

If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
Item(EditorIndex).Enchanter = chkEN.Value

' Error handler
Exit Sub
errorhandler:
HandleError "chkEN_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
Private Sub chkSM_Click()
' If debug mode, handle error then exit out
If options.Debug = 1 Then On Error GoTo errorhandler

If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
Item(EditorIndex).Smithy = chksm.Value

' Error handler
Exit Sub
errorhandler:
HandleError "chkSM_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
