VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmRL 
   Caption         =   "人力资源"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10890
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7020
   ScaleWidth      =   10890
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7470
      Top             =   6510
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6660
      Top             =   6510
   End
   Begin VB.Frame frmAn 
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   8040
      TabIndex        =   151
      Top             =   6300
      Width           =   2175
      Begin VB.CommandButton cmdSave 
         Caption         =   "提交"
         Height          =   585
         Left            =   1440
         Picture         =   "frmTX.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   153
         Top             =   150
         Width           =   705
      End
      Begin VB.CommandButton cmdMod 
         Caption         =   "修改"
         Height          =   555
         Left            =   750
         Picture         =   "frmTX.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   152
         Top             =   180
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   585
      Left            =   10230
      Picture         =   "frmTX.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   150
      Top             =   6450
      Width           =   675
   End
   Begin TabDlg.SSTab tabRen 
      Height          =   7065
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   12462
      _Version        =   393216
      Tabs            =   7
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "员工基本信息"
      TabPicture(0)   =   "frmTX.frx":0A76
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtR(85)"
      Tab(0).Control(1)=   "txtR(84)"
      Tab(0).Control(2)=   "txtR(83)"
      Tab(0).Control(3)=   "txtOld"
      Tab(0).Control(4)=   "txtR(12)"
      Tab(0).Control(5)=   "txtR(6)"
      Tab(0).Control(6)=   "txtR(4)"
      Tab(0).Control(7)=   "txtR(2)"
      Tab(0).Control(8)=   "txtR(1)"
      Tab(0).Control(9)=   "txtR(0)"
      Tab(0).Control(10)=   "txtR(15)"
      Tab(0).Control(11)=   "txtR(14)"
      Tab(0).Control(12)=   "txtR(13)"
      Tab(0).Control(13)=   "txtR(11)"
      Tab(0).Control(14)=   "txtR(10)"
      Tab(0).Control(15)=   "txtR(9)"
      Tab(0).Control(16)=   "txtR(8)"
      Tab(0).Control(17)=   "txtR(7)"
      Tab(0).Control(18)=   "txtR(5)"
      Tab(0).Control(19)=   "txtR(3)"
      Tab(0).Control(20)=   "Label85"
      Tab(0).Control(21)=   "Label84"
      Tab(0).Control(22)=   "Label83"
      Tab(0).Control(23)=   "lblWid"
      Tab(0).Control(24)=   "Label15"
      Tab(0).Control(25)=   "Label14"
      Tab(0).Control(26)=   "Label13"
      Tab(0).Control(27)=   "Label12(2)"
      Tab(0).Control(28)=   "Label12(1)"
      Tab(0).Control(29)=   "Label12(0)"
      Tab(0).Control(30)=   "Label11"
      Tab(0).Control(31)=   "Label10"
      Tab(0).Control(32)=   "Label9"
      Tab(0).Control(33)=   "Label8"
      Tab(0).Control(34)=   "Label7"
      Tab(0).Control(35)=   "Label6"
      Tab(0).Control(36)=   "Label5"
      Tab(0).Control(37)=   "Label4"
      Tab(0).Control(38)=   "Label3"
      Tab(0).Control(39)=   "Label2"
      Tab(0).Control(40)=   "Label1"
      Tab(0).ControlCount=   41
      TabCaption(1)   =   "员工劳动关系信息"
      TabPicture(1)   =   "frmTX.frx":0A92
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "txtGOld"
      Tab(1).Control(2)=   "txtR(24)"
      Tab(1).Control(3)=   "txtR(23)"
      Tab(1).Control(4)=   "txtR(22)"
      Tab(1).Control(5)=   "txtR(21)"
      Tab(1).Control(6)=   "txtR(20)"
      Tab(1).Control(7)=   "txtR(19)"
      Tab(1).Control(8)=   "txtR(18)"
      Tab(1).Control(9)=   "txtR(17)"
      Tab(1).Control(10)=   "txtR(16)"
      Tab(1).Control(11)=   "Label25"
      Tab(1).Control(12)=   "Label24"
      Tab(1).Control(13)=   "Label23"
      Tab(1).Control(14)=   "Label22"
      Tab(1).Control(15)=   "Label21"
      Tab(1).Control(16)=   "Label20"
      Tab(1).Control(17)=   "Label19"
      Tab(1).Control(18)=   "Label18"
      Tab(1).Control(19)=   "Label17"
      Tab(1).Control(20)=   "Label16"
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "员工薪资福利信息"
      TabPicture(2)   =   "frmTX.frx":0AAE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtR(88)"
      Tab(2).Control(1)=   "txtR(87)"
      Tab(2).Control(2)=   "txtR(86)"
      Tab(2).Control(3)=   "txtR(62)"
      Tab(2).Control(4)=   "txtR(61)"
      Tab(2).Control(5)=   "txtR(60)"
      Tab(2).Control(6)=   "txtR(59)"
      Tab(2).Control(7)=   "txtR(58)"
      Tab(2).Control(8)=   "txtR(57)"
      Tab(2).Control(9)=   "txtR(64)"
      Tab(2).Control(10)=   "txtR(56)"
      Tab(2).Control(11)=   "txtR(55)"
      Tab(2).Control(12)=   "txtR(54)"
      Tab(2).Control(13)=   "txtR(53)"
      Tab(2).Control(14)=   "txtR(52)"
      Tab(2).Control(15)=   "txtR(63)"
      Tab(2).Control(16)=   "txtR(51)"
      Tab(2).Control(17)=   "txtR(50)"
      Tab(2).Control(18)=   "txtR(49)"
      Tab(2).Control(19)=   "txtR(48)"
      Tab(2).Control(20)=   "txtR(47)"
      Tab(2).Control(21)=   "txtR(46)"
      Tab(2).Control(22)=   "txtR(45)"
      Tab(2).Control(23)=   "txtR(44)"
      Tab(2).Control(24)=   "txtR(43)"
      Tab(2).Control(25)=   "txtR(42)"
      Tab(2).Control(26)=   "txtR(41)"
      Tab(2).Control(27)=   "txtR(40)"
      Tab(2).Control(28)=   "txtR(39)"
      Tab(2).Control(29)=   "txtR(38)"
      Tab(2).Control(30)=   "txtR(37)"
      Tab(2).Control(31)=   "txtR(36)"
      Tab(2).Control(32)=   "txtR(35)"
      Tab(2).Control(33)=   "txtR(34)"
      Tab(2).Control(34)=   "txtR(33)"
      Tab(2).Control(35)=   "txtR(32)"
      Tab(2).Control(36)=   "txtR(31)"
      Tab(2).Control(37)=   "txtR(30)"
      Tab(2).Control(38)=   "txtR(29)"
      Tab(2).Control(39)=   "txtR(28)"
      Tab(2).Control(40)=   "txtR(27)"
      Tab(2).Control(41)=   "txtR(26)"
      Tab(2).Control(42)=   "txtR(25)"
      Tab(2).Control(43)=   "Label89"
      Tab(2).Control(44)=   "Label88"
      Tab(2).Control(45)=   "Label87"
      Tab(2).Control(46)=   "Label86"
      Tab(2).Control(47)=   "Label65"
      Tab(2).Control(48)=   "Label64"
      Tab(2).Control(49)=   "Label63"
      Tab(2).Control(50)=   "Label62"
      Tab(2).Control(51)=   "Label61"
      Tab(2).Control(52)=   "Label60"
      Tab(2).Control(53)=   "Label59"
      Tab(2).Control(54)=   "Label58"
      Tab(2).Control(55)=   "Label57"
      Tab(2).Control(56)=   "Label56"
      Tab(2).Control(57)=   "Label55"
      Tab(2).Control(58)=   "Label54"
      Tab(2).Control(59)=   "Label53"
      Tab(2).Control(60)=   "Label52"
      Tab(2).Control(61)=   "Label51"
      Tab(2).Control(62)=   "Label50"
      Tab(2).Control(63)=   "Label49"
      Tab(2).Control(64)=   "Label48"
      Tab(2).Control(65)=   "Label42"
      Tab(2).Control(66)=   "Label43"
      Tab(2).Control(67)=   "Label41"
      Tab(2).Control(68)=   "Label40"
      Tab(2).Control(69)=   "Label37"
      Tab(2).Control(70)=   "Label34"
      Tab(2).Control(71)=   "Label31"
      Tab(2).Control(72)=   "Label28"
      Tab(2).Control(73)=   "Label47"
      Tab(2).Control(74)=   "Label46"
      Tab(2).Control(75)=   "Label45"
      Tab(2).Control(76)=   "Label39"
      Tab(2).Control(77)=   "Label36"
      Tab(2).Control(78)=   "Label33"
      Tab(2).Control(79)=   "Label30"
      Tab(2).Control(80)=   "Label27"
      Tab(2).Control(81)=   "Label44"
      Tab(2).Control(82)=   "Label38"
      Tab(2).Control(83)=   "Label35"
      Tab(2).Control(84)=   "Label32"
      Tab(2).Control(85)=   "Label29"
      Tab(2).Control(86)=   "Label26"
      Tab(2).ControlCount=   87
      TabCaption(3)   =   "员工资质及培训信息"
      TabPicture(3)   =   "frmTX.frx":0ACA
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label66"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label67"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label68"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label69"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label70"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label71"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Label72"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Label73"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "txtR(65)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "txtR(66)"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "txtR(67)"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "txtR(68)"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "txtR(69)"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "txtR(70)"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "txtR(71)"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "txtR(72)"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).ControlCount=   16
      TabCaption(4)   =   "员工在职期间奖励与过失记录"
      TabPicture(4)   =   "frmTX.frx":0AE6
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtR(73)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "员工评价"
      TabPicture(5)   =   "frmTX.frx":0B02
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txtR(89)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "福利资料信息"
      TabPicture(6)   =   "frmTX.frx":0B1E
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      Begin VB.TextBox txtR 
         Height          =   5280
         Index           =   89
         Left            =   -74790
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   189
         Top             =   900
         Width           =   10335
      End
      Begin VB.TextBox txtR 
         ForeColor       =   &H00008000&
         Height          =   270
         Index           =   88
         Left            =   -72960
         TabIndex        =   188
         Top             =   5760
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   87
         Left            =   -69510
         TabIndex        =   186
         Top             =   5400
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         ForeColor       =   &H00008000&
         Height          =   270
         Index           =   86
         Left            =   -72960
         TabIndex        =   185
         Top             =   5400
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   85
         Left            =   -66420
         TabIndex        =   181
         Top             =   2610
         Width           =   1965
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   84
         Left            =   -69810
         TabIndex        =   180
         Top             =   2670
         Width           =   2025
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   83
         Left            =   -73410
         TabIndex        =   177
         Top             =   3810
         Width           =   1965
      End
      Begin VB.Frame Frame1 
         Caption         =   "兼管职务"
         Height          =   2205
         Left            =   -74820
         TabIndex        =   157
         Top             =   4110
         Width           =   9945
         Begin VB.TextBox txtR 
            Height          =   315
            Index           =   82
            Left            =   7710
            TabIndex        =   175
            Top             =   1290
            Width           =   1545
         End
         Begin VB.TextBox txtR 
            Height          =   315
            Index           =   81
            Left            =   4380
            TabIndex        =   173
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox txtR 
            Height          =   285
            Index           =   80
            Left            =   1470
            TabIndex        =   171
            Top             =   1320
            Width           =   1605
         End
         Begin VB.TextBox txtR 
            Height          =   315
            Index           =   79
            Left            =   7710
            TabIndex        =   169
            Top             =   840
            Width           =   1545
         End
         Begin VB.TextBox txtR 
            Height          =   315
            Index           =   78
            Left            =   4380
            TabIndex        =   167
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtR 
            Height          =   285
            Index           =   77
            Left            =   1440
            TabIndex        =   165
            Top             =   870
            Width           =   1605
         End
         Begin VB.TextBox txtR 
            Height          =   315
            Index           =   76
            Left            =   7710
            TabIndex        =   163
            Top             =   360
            Width           =   1545
         End
         Begin VB.TextBox txtR 
            Height          =   315
            Index           =   75
            Left            =   4380
            TabIndex        =   161
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtR 
            Height          =   285
            Index           =   74
            Left            =   1470
            TabIndex        =   159
            Top             =   390
            Width           =   1605
         End
         Begin VB.Label Label82 
            Caption         =   "上级"
            Height          =   255
            Left            =   6930
            TabIndex        =   174
            Top             =   1350
            Width           =   675
         End
         Begin VB.Label Label81 
            Caption         =   "职务"
            Height          =   225
            Left            =   3450
            TabIndex        =   172
            Top             =   1350
            Width           =   645
         End
         Begin VB.Label Label80 
            Caption         =   "3 部门"
            Height          =   285
            Left            =   450
            TabIndex        =   170
            Top             =   1380
            Width           =   555
         End
         Begin VB.Label Label79 
            Caption         =   "上级"
            Height          =   255
            Left            =   6930
            TabIndex        =   168
            Top             =   900
            Width           =   675
         End
         Begin VB.Label Label78 
            Caption         =   "职务"
            Height          =   225
            Left            =   3450
            TabIndex        =   166
            Top             =   900
            Width           =   645
         End
         Begin VB.Label Label77 
            Caption         =   "2 部门"
            Height          =   285
            Left            =   450
            TabIndex        =   164
            Top             =   930
            Width           =   555
         End
         Begin VB.Label Label76 
            Caption         =   "上级"
            Height          =   255
            Left            =   6930
            TabIndex        =   162
            Top             =   420
            Width           =   675
         End
         Begin VB.Label Label75 
            Caption         =   "职务"
            Height          =   225
            Left            =   3450
            TabIndex        =   160
            Top             =   420
            Width           =   645
         End
         Begin VB.Label Label74 
            Caption         =   "1 部门"
            Height          =   285
            Left            =   450
            TabIndex        =   158
            Top             =   450
            Width           =   555
         End
      End
      Begin VB.TextBox txtGOld 
         Height          =   285
         Left            =   -70410
         Locked          =   -1  'True
         TabIndex        =   155
         Top             =   2670
         Width           =   1605
      End
      Begin VB.TextBox txtOld 
         Height          =   270
         Left            =   -69810
         Locked          =   -1  'True
         TabIndex        =   154
         Top             =   1560
         Width           =   2025
      End
      Begin VB.TextBox txtR 
         Height          =   5220
         Index           =   73
         Left            =   -74790
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   149
         Top             =   900
         Width           =   10335
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   72
         Left            =   8610
         TabIndex        =   148
         Top             =   1890
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   71
         Left            =   8610
         TabIndex        =   147
         Top             =   990
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   70
         Left            =   5310
         TabIndex        =   146
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   69
         Left            =   5310
         TabIndex        =   145
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   68
         Left            =   5310
         TabIndex        =   144
         Top             =   990
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   67
         Left            =   1950
         TabIndex        =   143
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   66
         Left            =   1950
         TabIndex        =   142
         Top             =   1950
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   65
         Left            =   1950
         TabIndex        =   141
         Top             =   990
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   62
         Left            =   -66180
         TabIndex        =   132
         Top             =   5430
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   61
         Left            =   -66180
         TabIndex        =   131
         Top             =   5076
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   60
         Left            =   -66180
         TabIndex        =   130
         Top             =   4723
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         ForeColor       =   &H00008000&
         Height          =   270
         Index           =   59
         Left            =   -66180
         TabIndex        =   129
         Top             =   4370
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   58
         Left            =   -66180
         TabIndex        =   128
         Top             =   4017
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   57
         Left            =   -66180
         TabIndex        =   127
         Top             =   3664
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         ForeColor       =   &H00008000&
         Height          =   270
         Index           =   64
         Left            =   -66180
         TabIndex        =   126
         Top             =   5760
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   56
         Left            =   -69510
         TabIndex        =   125
         Top             =   5076
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   55
         Left            =   -69510
         TabIndex        =   124
         Top             =   4723
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         ForeColor       =   &H00008000&
         Height          =   270
         Index           =   54
         Left            =   -69510
         TabIndex        =   123
         Top             =   4370
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   53
         Left            =   -69510
         TabIndex        =   122
         Top             =   4017
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   52
         Left            =   -69510
         TabIndex        =   121
         Top             =   3664
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   63
         Left            =   -69510
         TabIndex        =   120
         Top             =   5760
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   51
         Left            =   -72960
         TabIndex        =   119
         Top             =   5076
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   50
         Left            =   -72960
         TabIndex        =   118
         Top             =   4723
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         ForeColor       =   &H00008000&
         Height          =   270
         Index           =   49
         Left            =   -72960
         TabIndex        =   117
         Top             =   4370
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   48
         Left            =   -72960
         TabIndex        =   116
         Top             =   4017
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   47
         Left            =   -72960
         TabIndex        =   115
         Top             =   3664
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   46
         Left            =   -66180
         TabIndex        =   114
         Top             =   2615
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   45
         Left            =   -66180
         TabIndex        =   113
         Top             =   2260
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   44
         Left            =   -66180
         TabIndex        =   112
         Top             =   1905
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   43
         Left            =   -66180
         TabIndex        =   111
         Top             =   1550
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   42
         Left            =   -66180
         TabIndex        =   110
         Top             =   1195
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   41
         Left            =   -66180
         TabIndex        =   109
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   40
         Left            =   -69510
         TabIndex        =   108
         Top             =   3311
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   39
         Left            =   -69510
         TabIndex        =   107
         Top             =   2958
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   38
         Left            =   -69510
         TabIndex        =   106
         Top             =   2605
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   37
         Left            =   -69510
         TabIndex        =   105
         Top             =   2252
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   36
         Left            =   -69510
         TabIndex        =   104
         Top             =   1899
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   35
         Left            =   -69510
         TabIndex        =   103
         Top             =   1546
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   34
         Left            =   -69510
         TabIndex        =   102
         Top             =   1193
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   33
         Left            =   -69510
         TabIndex        =   101
         Top             =   870
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   32
         Left            =   -72960
         Locked          =   -1  'True
         TabIndex        =   100
         Top             =   3311
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   31
         Left            =   -72960
         TabIndex        =   99
         Top             =   2958
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   30
         Left            =   -72960
         TabIndex        =   98
         Top             =   2605
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   29
         Left            =   -72960
         TabIndex        =   97
         Top             =   2252
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   28
         Left            =   -72960
         Locked          =   -1  'True
         TabIndex        =   84
         Top             =   1899
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   27
         Left            =   -72960
         TabIndex        =   83
         Top             =   1546
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   26
         Left            =   -72960
         TabIndex        =   82
         Top             =   1193
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   25
         Left            =   -72960
         TabIndex        =   81
         Top             =   870
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   24
         Left            =   -70410
         TabIndex        =   52
         Top             =   3540
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   23
         Left            =   -73350
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   3540
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   22
         Left            =   -67110
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   21
         Left            =   -73350
         TabIndex        =   49
         Top             =   2670
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   20
         Left            =   -70410
         TabIndex        =   48
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   19
         Left            =   -73350
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1770
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   18
         Left            =   -67110
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   930
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   17
         Left            =   -70410
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   16
         Left            =   -73350
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   12
         Left            =   -66420
         TabIndex        =   33
         Top             =   2160
         Width           =   1965
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   6
         Left            =   -73410
         TabIndex        =   32
         Top             =   2670
         Width           =   1965
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   4
         Left            =   -66390
         TabIndex        =   31
         Top             =   1530
         Width           =   1965
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   2
         Left            =   -66390
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   900
         Width           =   1965
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   1
         Left            =   -69810
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   900
         Width           =   2025
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   0
         Left            =   -73410
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   900
         Width           =   2025
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   15
         Left            =   -69780
         TabIndex        =   27
         Top             =   5550
         Width           =   1965
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   14
         Left            =   -73410
         TabIndex        =   26
         Top             =   5580
         Width           =   1965
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   13
         Left            =   -73410
         TabIndex        =   25
         Top             =   4380
         Width           =   8985
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   11
         Left            =   -69810
         TabIndex        =   24
         Top             =   2130
         Width           =   2025
      End
      Begin VB.TextBox txtR 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Index           =   10
         Left            =   -66390
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   4890
         Width           =   1965
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   9
         Left            =   -69780
         TabIndex        =   22
         Top             =   4920
         Width           =   1965
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   8
         Left            =   -73410
         TabIndex        =   21
         Top             =   4950
         Width           =   1965
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   7
         Left            =   -73410
         TabIndex        =   20
         Top             =   3270
         Width           =   8985
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   5
         Left            =   -73410
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "310113197311041718"
         Top             =   2160
         Width           =   2025
      End
      Begin VB.TextBox txtR 
         Height          =   270
         Index           =   3
         Left            =   -73410
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1530
         Width           =   2025
      End
      Begin VB.Label Label89 
         Caption         =   "采暖基金"
         Height          =   225
         Left            =   -74190
         TabIndex        =   187
         Top             =   5850
         Width           =   885
      End
      Begin VB.Label Label88 
         Caption         =   "公积金(公司)"
         Height          =   195
         Left            =   -70800
         TabIndex        =   184
         Top             =   5460
         Width           =   1125
      End
      Begin VB.Label Label87 
         Caption         =   "Label87"
         Height          =   255
         Left            =   -70740
         TabIndex        =   183
         Top             =   6210
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label86 
         Caption         =   "公积金(个人)"
         ForeColor       =   &H00008000&
         Height          =   225
         Left            =   -74490
         TabIndex        =   182
         Top             =   5460
         Width           =   1215
      End
      Begin VB.Label Label85 
         Caption         =   "专业"
         Height          =   255
         Left            =   -67170
         TabIndex        =   179
         Top             =   2670
         Width           =   735
      End
      Begin VB.Label Label84 
         Caption         =   "毕业院校"
         Height          =   225
         Left            =   -70830
         TabIndex        =   178
         Top             =   2700
         Width           =   795
      End
      Begin VB.Label Label83 
         Alignment       =   1  'Right Justify
         Caption         =   "户别"
         Height          =   225
         Left            =   -74280
         TabIndex        =   176
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label lblWid 
         Caption         =   "lblWid"
         Height          =   315
         Left            =   -71070
         TabIndex        =   156
         Top             =   6510
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label73 
         Alignment       =   1  'Right Justify
         Caption         =   "培训协议期"
         Height          =   225
         Left            =   3810
         TabIndex        =   140
         Top             =   2940
         Width           =   1395
      End
      Begin VB.Label Label72 
         Alignment       =   1  'Right Justify
         Caption         =   "公司付费否"
         Height          =   225
         Left            =   360
         TabIndex        =   139
         Top             =   2940
         Width           =   1395
      End
      Begin VB.Label Label71 
         Alignment       =   1  'Right Justify
         Caption         =   "证书有效期"
         Height          =   225
         Left            =   6960
         TabIndex        =   138
         Top             =   1980
         Width           =   1395
      End
      Begin VB.Label Label70 
         Alignment       =   1  'Right Justify
         Caption         =   "获证日期"
         Height          =   225
         Left            =   3810
         TabIndex        =   137
         Top             =   1995
         Width           =   1395
      End
      Begin VB.Label Label69 
         Alignment       =   1  'Right Justify
         Caption         =   "发证机关/培训人"
         Height          =   225
         Left            =   360
         TabIndex        =   136
         Top             =   1995
         Width           =   1395
      End
      Begin VB.Label Label68 
         Alignment       =   1  'Right Justify
         Caption         =   "外训/内训"
         Height          =   225
         Left            =   6990
         TabIndex        =   135
         Top             =   1050
         Width           =   1395
      End
      Begin VB.Label Label67 
         Alignment       =   1  'Right Justify
         Caption         =   "证书级别"
         Height          =   225
         Left            =   3810
         TabIndex        =   134
         Top             =   1050
         Width           =   1395
      End
      Begin VB.Label Label66 
         Alignment       =   1  'Right Justify
         Caption         =   "证书名称"
         Height          =   225
         Left            =   360
         TabIndex        =   133
         Top             =   1050
         Width           =   1395
      End
      Begin VB.Label Label65 
         Alignment       =   1  'Right Justify
         Caption         =   "意外保险合同号"
         Height          =   195
         Left            =   -67620
         TabIndex        =   96
         Top             =   5490
         Width           =   1275
      End
      Begin VB.Label Label64 
         Alignment       =   1  'Right Justify
         Caption         =   "个人保险合计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   -67620
         TabIndex        =   95
         Top             =   5820
         Width           =   1275
      End
      Begin VB.Label Label63 
         Alignment       =   1  'Right Justify
         Caption         =   "公司保险合计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   -70950
         TabIndex        =   94
         Top             =   5850
         Width           =   1275
      End
      Begin VB.Label Label62 
         Alignment       =   1  'Right Justify
         Caption         =   "意外保险到期日"
         Height          =   195
         Left            =   -67620
         TabIndex        =   93
         Top             =   5136
         Width           =   1275
      End
      Begin VB.Label Label61 
         Alignment       =   1  'Right Justify
         Caption         =   "意外保险"
         Height          =   195
         Left            =   -71040
         TabIndex        =   92
         Top             =   5130
         Width           =   1275
      End
      Begin VB.Label Label60 
         Alignment       =   1  'Right Justify
         Caption         =   "综合保险"
         Height          =   195
         Left            =   -74700
         TabIndex        =   91
         Top             =   5136
         Width           =   1275
      End
      Begin VB.Label Label59 
         Alignment       =   1  'Right Justify
         Caption         =   "公司大病医疗"
         Height          =   195
         Left            =   -67620
         TabIndex        =   90
         Top             =   4782
         Width           =   1275
      End
      Begin VB.Label Label58 
         Alignment       =   1  'Right Justify
         Caption         =   "公司生育保险"
         Height          =   195
         Left            =   -71040
         TabIndex        =   89
         Top             =   4785
         Width           =   1275
      End
      Begin VB.Label Label57 
         Alignment       =   1  'Right Justify
         Caption         =   "公司工伤保险"
         Height          =   195
         Left            =   -74700
         TabIndex        =   88
         Top             =   4783
         Width           =   1275
      End
      Begin VB.Label Label56 
         Alignment       =   1  'Right Justify
         Caption         =   "个人失业保险"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   -67620
         TabIndex        =   87
         Top             =   4428
         Width           =   1275
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         Caption         =   "个人医疗保险"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   -71040
         TabIndex        =   86
         Top             =   4425
         Width           =   1275
      End
      Begin VB.Label Label54 
         Alignment       =   1  'Right Justify
         Caption         =   "个人养老保险"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   -74700
         TabIndex        =   85
         Top             =   4430
         Width           =   1275
      End
      Begin VB.Label Label53 
         Alignment       =   1  'Right Justify
         Caption         =   "公司失业保险"
         Height          =   195
         Left            =   -67620
         TabIndex        =   80
         Top             =   4074
         Width           =   1275
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         Caption         =   "公司医疗保险"
         Height          =   195
         Left            =   -71040
         TabIndex        =   79
         Top             =   4080
         Width           =   1275
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         Caption         =   "公司养老保险"
         Height          =   195
         Left            =   -74700
         TabIndex        =   78
         Top             =   4077
         Width           =   1275
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         Caption         =   "代理费"
         Height          =   195
         Left            =   -67620
         TabIndex        =   77
         Top             =   3720
         Width           =   1275
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         Caption         =   "社保基数"
         Height          =   195
         Left            =   -71040
         TabIndex        =   76
         Top             =   3724
         Width           =   1275
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         Caption         =   "交金地区"
         Height          =   195
         Left            =   -74700
         TabIndex        =   75
         Top             =   3724
         Width           =   1275
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         Caption         =   "驻外补贴"
         Height          =   195
         Left            =   -71040
         TabIndex        =   74
         Top             =   2665
         Width           =   1275
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         Caption         =   "福利合计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   -67500
         TabIndex        =   73
         Top             =   2665
         Width           =   1155
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "津贴"
         Height          =   195
         Left            =   -74700
         TabIndex        =   72
         Top             =   2665
         Width           =   1275
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         Caption         =   "其他福利"
         Height          =   195
         Left            =   -67500
         TabIndex        =   71
         Top             =   2310
         Width           =   1155
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         Caption         =   "旅游福利"
         Height          =   195
         Left            =   -67500
         TabIndex        =   70
         Top             =   1965
         Width           =   1155
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         Caption         =   "中秋节福利"
         Height          =   195
         Left            =   -67500
         TabIndex        =   69
         Top             =   1605
         Width           =   1155
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         Caption         =   "妇女节福利"
         Height          =   195
         Left            =   -67500
         TabIndex        =   68
         Top             =   1260
         Width           =   1155
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "生日福利"
         Height          =   195
         Left            =   -67500
         TabIndex        =   67
         Top             =   900
         Width           =   1155
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         Caption         =   "补贴合计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   -71040
         TabIndex        =   66
         Top             =   3375
         Width           =   1275
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         Caption         =   "工资合计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   -74700
         TabIndex        =   65
         Top             =   3371
         Width           =   1275
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         Caption         =   "特殊补贴"
         Height          =   195
         Left            =   -71040
         TabIndex        =   64
         Top             =   3015
         Width           =   1275
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         Caption         =   "住房补贴"
         Height          =   195
         Left            =   -71040
         TabIndex        =   63
         Top             =   2310
         Width           =   1275
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         Caption         =   "通信补贴"
         Height          =   195
         Left            =   -71040
         TabIndex        =   62
         Top             =   1965
         Width           =   1275
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         Caption         =   "伙食补贴"
         Height          =   195
         Left            =   -71040
         TabIndex        =   61
         Top             =   1605
         Width           =   1275
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Caption         =   "交通补贴"
         Height          =   195
         Left            =   -71040
         TabIndex        =   60
         Top             =   1260
         Width           =   1275
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Caption         =   "职位补贴"
         Height          =   195
         Left            =   -71040
         TabIndex        =   59
         Top             =   900
         Width           =   1275
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         Caption         =   "奖金"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   -74700
         TabIndex        =   58
         Top             =   3018
         Width           =   1275
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         Caption         =   "绩效奖金"
         Height          =   195
         Left            =   -74700
         TabIndex        =   57
         Top             =   2312
         Width           =   1275
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         Caption         =   "岗位工资合计"
         Height          =   195
         Left            =   -74700
         TabIndex        =   56
         Top             =   1959
         Width           =   1275
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         Caption         =   "岗位技能"
         Height          =   195
         Left            =   -74700
         TabIndex        =   55
         Top             =   1606
         Width           =   1275
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Caption         =   "基本工资"
         Height          =   195
         Left            =   -74700
         TabIndex        =   54
         Top             =   1253
         Width           =   1275
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "级别"
         Height          =   195
         Left            =   -74700
         TabIndex        =   53
         Top             =   900
         Width           =   1275
      End
      Begin VB.Label Label25 
         Caption         =   "合同到期日"
         Height          =   255
         Left            =   -71550
         TabIndex        =   43
         Top             =   3570
         Width           =   1035
      End
      Begin VB.Label Label24 
         Caption         =   "是否试用期"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74490
         TabIndex        =   42
         Top             =   3570
         Width           =   1035
      End
      Begin VB.Label Label23 
         Caption         =   "在职否"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -67980
         TabIndex        =   41
         Top             =   2710
         Width           =   1035
      End
      Begin VB.Label Label22 
         Caption         =   "工龄"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -71550
         TabIndex        =   40
         Top             =   2710
         Width           =   1035
      End
      Begin VB.Label Label21 
         Caption         =   "入职时间"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74490
         TabIndex        =   39
         Top             =   2710
         Width           =   1035
      End
      Begin VB.Label Label20 
         Caption         =   "职称"
         Height          =   255
         Left            =   -71550
         TabIndex        =   38
         Top             =   1860
         Width           =   1035
      End
      Begin VB.Label Label19 
         Caption         =   "职务"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74490
         TabIndex        =   37
         Top             =   1850
         Width           =   1035
      End
      Begin VB.Label Label18 
         Caption         =   "直属上级"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -68070
         TabIndex        =   36
         Top             =   990
         Width           =   1035
      End
      Begin VB.Label Label17 
         Caption         =   "部门"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -71550
         TabIndex        =   35
         Top             =   990
         Width           =   1035
      End
      Begin VB.Label Label16 
         Caption         =   "区域"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74490
         TabIndex        =   34
         Top             =   990
         Width           =   1035
      End
      Begin VB.Label Label15 
         Caption         =   "紧急联系电话"
         Height          =   315
         Left            =   -71130
         TabIndex        =   17
         Top             =   5610
         Width           =   1275
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "紧急联系人"
         Height          =   195
         Left            =   -74790
         TabIndex        =   16
         Top             =   5610
         Width           =   1005
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "家庭地址"
         Height          =   315
         Left            =   -75180
         TabIndex        =   15
         Top             =   4440
         Width           =   1395
      End
      Begin VB.Label Label12 
         Caption         =   "公司手机号"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   2
         Left            =   -67410
         TabIndex        =   14
         Top             =   4950
         Width           =   1005
      End
      Begin VB.Label Label12 
         Caption         =   "私人手机"
         Height          =   285
         Index           =   1
         Left            =   -71040
         TabIndex        =   13
         Top             =   4980
         Width           =   1035
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "家庭电话"
         Height          =   285
         Index           =   0
         Left            =   -74820
         TabIndex        =   12
         Top             =   5040
         Width           =   1035
      End
      Begin VB.Label Label11 
         Caption         =   "婚否"
         Height          =   285
         Left            =   -67170
         TabIndex        =   11
         Top             =   2190
         Width           =   645
      End
      Begin VB.Label Label10 
         Caption         =   "邮编"
         Height          =   315
         Left            =   -70710
         TabIndex        =   10
         Top             =   2190
         Width           =   645
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "户籍地址"
         Height          =   315
         Left            =   -75000
         TabIndex        =   9
         Top             =   3330
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "学历"
         Height          =   285
         Left            =   -74670
         TabIndex        =   8
         Top             =   2730
         Width           =   885
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "身份证号码"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -75210
         TabIndex        =   7
         Top             =   2220
         Width           =   1425
      End
      Begin VB.Label Label6 
         Caption         =   "民族"
         Height          =   255
         Left            =   -67170
         TabIndex        =   6
         Top             =   1620
         Width           =   885
      End
      Begin VB.Label Label5 
         Caption         =   "年龄"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -70710
         TabIndex        =   5
         Top             =   1620
         Width           =   585
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "出生年月日"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -75300
         TabIndex        =   4
         Top             =   1620
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "性别"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -67170
         TabIndex        =   3
         Top             =   930
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "姓名"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -70710
         TabIndex        =   2
         Top             =   930
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "员工编号"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   -74670
         TabIndex        =   1
         Top             =   930
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmRL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim timZm As Integer  '1新保存1
Private Sub cmdBack_Click()
Me.Visible = False
frmRen.Enabled = True
frmRen.ZOrder 0
End Sub

Private Sub cmdMod_Click()
cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
Dim cmd As Object
Dim tt As String

On Error Resume Next


    timZm = 1 '新保存1
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "人事档案"
        mod1.cmd.Parameters("@NBLX") = "新保存1"
        mod1.cmd.Parameters("@bh") = lblWid.Caption
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = txtR(0).Text '编号
        mod1.cmd.Parameters("@mt2") = txtR(1).Text '姓名
        mod1.cmd.Parameters("@mt3") = txtR(2).Text '性别
        mod1.cmd.Parameters("@mt4") = txtR(4).Text '
        mod1.cmd.Parameters("@mt5") = txtR(5).Text '
        mod1.cmd.Parameters("@mt6") = txtR(6).Text '
        mod1.cmd.Parameters("@mt7") = txtR(7).Text '
        mod1.cmd.Parameters("@mt8") = txtR(8).Text '
        mod1.cmd.Parameters("@mt9") = txtR(9).Text '
        mod1.cmd.Parameters("@mt10") = txtR(10).Text '
        mod1.cmd.Parameters("@mt11") = txtR(11).Text '
        mod1.cmd.Parameters("@mt12") = txtR(12).Text '
        mod1.cmd.Parameters("@mt13") = txtR(13).Text '
        mod1.cmd.Parameters("@mt14") = txtR(14).Text '
        mod1.cmd.Parameters("@mt15") = txtR(15).Text '
        mod1.cmd.Parameters("@mt16") = txtR(20).Text '职称
        mod1.cmd.Parameters("@mt17") = txtR(74).Text '部门1
        mod1.cmd.Parameters("@mt18") = txtR(75).Text
        mod1.cmd.Parameters("@mt19") = txtR(77).Text '部门2
        mod1.cmd.Parameters("@mt20") = txtR(78).Text
        mod1.cmd.Parameters("@mt21") = txtR(80).Text '部门3
        mod1.cmd.Parameters("@mt22") = txtR(81).Text
        mod1.cmd.Parameters("@mt23") = txtR(76).Text '上级人1
        mod1.cmd.Parameters("@mt24") = txtR(79).Text '上级人2
        mod1.cmd.Parameters("@mt25") = txtR(82).Text '上级人3
        mod1.cmd.Parameters("@mt26") = txtR(62).Text '意外保险合同号
        mod1.cmd.Parameters("@mt27") = txtR(47).Text '交金地区
        mod1.cmd.Parameters("@mt28") = txtR(65).Text '证书名称
        mod1.cmd.Parameters("@mt29") = txtR(66).Text
        mod1.cmd.Parameters("@mt30") = txtR(67).Text
        mod1.cmd.Parameters("@mt31") = txtR(68).Text
        mod1.cmd.Parameters("@mt32") = txtR(71).Text
        mod1.cmd.Parameters("@mt33") = ""
        mod1.cmd.Parameters("@mt34") = ""
        mod1.cmd.Parameters("@mt35") = ""
        mod1.cmd.Parameters("@mlt1") = txtR(73).Text '员工在职期间奖励与过失记录
        mod1.cmd.Parameters("@mlt2") = txtR(89).Text '员工评价
        mod1.cmd.Parameters("@mlt3") = ""
        mod1.cmd.Parameters("@mlt4") = ""
        mod1.cmd.Parameters("@mlt5") = ""
        mod1.cmd.Parameters("@mm1") = txtR(25)
        mod1.cmd.Parameters("@mm2") = txtR(26)
        mod1.cmd.Parameters("@mm3") = txtR(27)
        mod1.cmd.Parameters("@mm4") = txtR(28)
        mod1.cmd.Parameters("@mm5") = txtR(29)
        mod1.cmd.Parameters("@mm6") = txtR(30)
        mod1.cmd.Parameters("@mm7") = txtR(31)
        mod1.cmd.Parameters("@mm8") = txtR(32)
        mod1.cmd.Parameters("@mm9") = txtR(33)
        mod1.cmd.Parameters("@mm10") = txtR(34)
        mod1.cmd.Parameters("@mm11") = txtR(35)
        mod1.cmd.Parameters("@mm12") = txtR(36)
        mod1.cmd.Parameters("@mm13") = txtR(37)
        mod1.cmd.Parameters("@mm14") = txtR(38)
        mod1.cmd.Parameters("@mm15") = txtR(39)
        mod1.cmd.Parameters("@mm16") = txtR(40)
        mod1.cmd.Parameters("@mm17") = txtR(41)
        mod1.cmd.Parameters("@mm18") = txtR(42)
        mod1.cmd.Parameters("@mm19") = txtR(43)
        mod1.cmd.Parameters("@mm20") = txtR(44)
        mod1.cmd.Parameters("@mm21") = txtR(45)
        mod1.cmd.Parameters("@mm22") = txtR(46)
        'mod1.cmd.Parameters("@mm23") = txtR(47)
        mod1.cmd.Parameters("@mm24") = txtR(48)
        mod1.cmd.Parameters("@mm25") = txtR(49)
        mod1.cmd.Parameters("@mm26") = txtR(50)
        mod1.cmd.Parameters("@mm27") = txtR(51)
        mod1.cmd.Parameters("@mm28") = txtR(52)
        mod1.cmd.Parameters("@mm29") = txtR(53)
        mod1.cmd.Parameters("@mm30") = txtR(54)
        mod1.cmd.Parameters("@mm31") = txtR(55)
        mod1.cmd.Parameters("@mm32") = txtR(56)
        mod1.cmd.Parameters("@mm33") = txtR(57)
        mod1.cmd.Parameters("@mm34") = txtR(58)
        mod1.cmd.Parameters("@mm35") = txtR(59)
        mod1.cmd.Parameters("@mm36") = txtR(60)
'        mod1.cmd.Parameters("@mm37") = txtR(61)
'        mod1.cmd.Parameters("@mm38") = txtR(62)
        mod1.cmd.Parameters("@mm39") = txtR(63)
        mod1.cmd.Parameters("@mm40") = txtR(64)
        mod1.cmd.Parameters("@mm41") = txtR(86) '公积金个人
        mod1.cmd.Parameters("@mm42") = txtR(87) '公积金公司
        mod1.cmd.Parameters("@mm43") = txtR(88) '采暖基金
        mod1.cmd.Parameters("@mb1") = 0
        mod1.cmd.Parameters("@mb2") = 0
        mod1.cmd.Parameters("@mb3") = 0
        mod1.cmd.Parameters("@mb4") = 0
        mod1.cmd.Parameters("@mb5") = 0
        mod1.cmd.Parameters("@md1") = txtR(3).Text '出生年月
        mod1.cmd.Parameters("@md2") = txtR(21).Text '入职时间
        mod1.cmd.Parameters("@md3") = txtR(24).Text '合同到期日
        mod1.cmd.Parameters("@md4") = txtR(61).Text '意外保险到期日
        mod1.cmd.Parameters("@md5") = txtR(69).Text '获证日期
        mod1.cmd.Parameters("@md6") = txtR(70).Text '培训协议期
        mod1.cmd.Parameters("@md7") = txtR(72).Text '证书有效期


    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
        mod1.cmd.Execute



        mod1.Zid = mod1.cmd.Parameters("@zid").Value
        If mod1.cmd.Parameters("@errch").Value <> "成功" Then
            MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
            cmdDing.Enabled = False
            Exit Sub
        Else '提交成功,等待系统中心处理数据
            Me.Enabled = False
            frmWaitA.Visible = True
            frmWaitA.Timer2.Enabled = False
    
            frmWaitA.ZOrder 0
            frmWaitA.Timer2.Enabled = True
            timWait.Enabled = True
        End If
    
        
Set cmd = Nothing
cmdSave.Enabled = False
cmdMod.Enabled = True
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Height = 7545
Me.Width = 11010
Me.Left = 0
Me.Top = 0
End Sub

Public Sub Qing()
On Error Resume Next
Dim oo As Integer
For oo = 0 To 89
    txtR(oo).Text = ""
Next
txtOld.Text = ""
txtGOld.Text = ""
lblWid.Caption = ""
tabRen.Tab = 0
End Sub

Public Sub Bound(Uid As String)
Dim Ra, Rb, RC, RD, RE
Dim ua, ub, uc, ud, ue

Dim tt As String
Dim oo As Integer
On Error GoTo gocuO

tt = "declare @auid nvarchar(10);" & _
    "Select * from RlA where Auid='" & Uid & "';" & _
    "select @auid=bguid from rla where Auid='" & Uid & "';" & _
    "select Aren from rla where auid=@auid;" & _
    "select phox,getdate() from worker where userid='" & Uid & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workFF, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
'Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
RC = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
For oo = 0 To 89
    txtR(oo).Text = Ra(oo, 0)
Next
lblWid.Caption = Ra(90, 0)
txtGOld.Text = Int(DateDiff("yyyy", txtR(21).Text, RC(1, 0)))
txtOld.Text = Int(DateDiff("yyyy", txtR(3).Text, RC(1, 0)))
txtR(18).Text = Rb(0, 0)
txtR(10).Text = RC(0, 0) '公司小号
'是否试用期
If txtR(23).Text = "是" Then
    txtR(23).Text = "否"
Else
    txtR(23).Text = "是"
End If
Exit Sub
gocuO:
MsgBox "网络故障，退出重试！"
End
End Sub

Private Sub timQuit_Timer()
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
If timZm = 1 Then '购发编辑
'''''MsgBox "已经成功通知销售经理转移此人的项目！"

End If

timQuit.Enabled = False
End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
Dim oo As Integer
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh,mm1,mt1,mm2,mt2 from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0

   If timZm = 1 Then '新保存
                
    'txtUid.Text = mod1.WP.Fields("mt2").Value

        

        

        
    End If
    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then '未完成

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '处理失败
    timWait.Enabled = False
    ii = MsgBox("服务中心在处理您的命令时,发生如下错误:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0

    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0

    Exit Sub
End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub


Private Sub txtR_Change(Index As Integer)
txtR(28).Text = Val(txtR(26).Text) + Val(txtR(27).Text)
txtR(32).Text = Val(txtR(28).Text) + Val(txtR(29).Text) + Val(txtR(30).Text) + Val(txtR(31).Text)
txtR(40).Text = Val(txtR(33).Text) + Val(txtR(34).Text) + Val(txtR(35).Text) + Val(txtR(36).Text) + Val(txtR(37).Text) + Val(txtR(38).Text) + Val(txtR(39).Text)
txtR(46).Text = Val(txtR(41).Text) + Val(txtR(42).Text) + Val(txtR(43).Text) + Val(txtR(44).Text) + Val(txtR(45).Text)
txtR(63).Text = Val(txtR(48).Text) + Val(txtR(50).Text) + Val(txtR(51).Text) + Val(txtR(53).Text) + _
                Val(txtR(55).Text) + Val(txtR(56).Text) + Val(txtR(87).Text) + Val(txtR(57).Text) + _
                Val(txtR(58).Text) + Val(txtR(60).Text) + Val(txtR(88).Text)
txtR(64).Text = Val(txtR(49).Text) + Val(txtR(86).Text) + Val(txtR(54).Text) + Val(txtR(59).Text)
End Sub


