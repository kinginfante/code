VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form NewGzd3 
   Caption         =   "冷水机组年度检修工作报告（单）"
   ClientHeight    =   10830
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15015
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10830
   ScaleWidth      =   15015
   Begin VB.TextBox BA 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   17
      Left            =   2010
      TabIndex        =   235
      Top             =   1290
      Width           =   4065
   End
   Begin MSDataGridLib.DataGrid dtgRen 
      Height          =   8085
      Left            =   8760
      TabIndex        =   232
      Top             =   240
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   14261
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "username"
         Caption         =   "姓名"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "gzu"
         Caption         =   "组号"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         ScrollBars      =   2
         BeginProperty Column00 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   794.835
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdBack 
      Height          =   375
      Left            =   14520
      Picture         =   "NewGzd3.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   182
      ToolTipText     =   "返回"
      Top             =   10410
      Width           =   465
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   14040
      Picture         =   "NewGzd3.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   181
      ToolTipText     =   "保存"
      Top             =   10410
      Width           =   465
   End
   Begin VB.CommandButton cmdMod 
      Height          =   375
      Left            =   13560
      Picture         =   "NewGzd3.frx":076C
      Style           =   1  'Graphical
      TabIndex        =   180
      ToolTipText     =   "修改"
      Top             =   10410
      Width           =   465
   End
   Begin VB.CommandButton cmdQm 
      Caption         =   "cmdQm"
      Height          =   345
      Index           =   0
      Left            =   9900
      TabIndex        =   179
      Top             =   10410
      Width           =   945
   End
   Begin TabDlg.SSTab tabNr 
      Height          =   8565
      Left            =   -30
      TabIndex        =   17
      Top             =   1650
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   15108
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "内容1"
      TabPicture(0)   =   "NewGzd3.frx":0A76
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label15"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label19"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line59"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line58"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line57"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line56"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Line55"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line54"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line43"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line42"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Line41"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line40"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Line39"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Line38(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Line33"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Line29"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Line25"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Line24"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Line23"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Line22"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Line21"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Line20"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Line19"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Line18"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Line17"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Line16"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Line15"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Line14"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Line13"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Line12"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Line11"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Line10"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Line9"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Line8"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Line7"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Line6"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Line5"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Line4"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Line3"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Line2"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Shape1"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Label18"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Label17"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Label16"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Label14"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Label13"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Label12"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Label11"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Label10"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Label9"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "Label8"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Label7"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Label6"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Label5"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Label4"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "Label3"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "Label2"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "Label1"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "C1(155)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "C1(154)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "C1(153)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "C1(152)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "C1(151)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "C1(150)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "C1(149)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "C1(148)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "C1(147)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "C1(146)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "C1(145)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "C1(144)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "C1(143)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "C1(142)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "C1(141)"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "C1(140)"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "C1(139)"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "C1(138)"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "C1(137)"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "C1(129)"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "C1(128)"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "C1(127)"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "C1(126)"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "C1(125)"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "C1(124)"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "C1(123)"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).Control(84)=   "C1(122)"
      Tab(0).Control(84).Enabled=   0   'False
      Tab(0).Control(85)=   "C1(121)"
      Tab(0).Control(85).Enabled=   0   'False
      Tab(0).Control(86)=   "C1(120)"
      Tab(0).Control(86).Enabled=   0   'False
      Tab(0).Control(87)=   "C1(119)"
      Tab(0).Control(87).Enabled=   0   'False
      Tab(0).Control(88)=   "C1(118)"
      Tab(0).Control(88).Enabled=   0   'False
      Tab(0).Control(89)=   "C1(117)"
      Tab(0).Control(89).Enabled=   0   'False
      Tab(0).Control(90)=   "C1(116)"
      Tab(0).Control(90).Enabled=   0   'False
      Tab(0).Control(91)=   "C1(115)"
      Tab(0).Control(91).Enabled=   0   'False
      Tab(0).Control(92)=   "C1(114)"
      Tab(0).Control(92).Enabled=   0   'False
      Tab(0).Control(93)=   "C1(113)"
      Tab(0).Control(93).Enabled=   0   'False
      Tab(0).Control(94)=   "C1(112)"
      Tab(0).Control(94).Enabled=   0   'False
      Tab(0).Control(95)=   "C1(111)"
      Tab(0).Control(95).Enabled=   0   'False
      Tab(0).Control(96)=   "C1(105)"
      Tab(0).Control(96).Enabled=   0   'False
      Tab(0).Control(97)=   "C1(104)"
      Tab(0).Control(97).Enabled=   0   'False
      Tab(0).Control(98)=   "C1(103)"
      Tab(0).Control(98).Enabled=   0   'False
      Tab(0).Control(99)=   "C1(102)"
      Tab(0).Control(99).Enabled=   0   'False
      Tab(0).Control(100)=   "C1(101)"
      Tab(0).Control(100).Enabled=   0   'False
      Tab(0).Control(101)=   "C1(100)"
      Tab(0).Control(101).Enabled=   0   'False
      Tab(0).Control(102)=   "C1(99)"
      Tab(0).Control(102).Enabled=   0   'False
      Tab(0).Control(103)=   "C1(98)"
      Tab(0).Control(103).Enabled=   0   'False
      Tab(0).Control(104)=   "C1(97)"
      Tab(0).Control(104).Enabled=   0   'False
      Tab(0).Control(105)=   "C1(96)"
      Tab(0).Control(105).Enabled=   0   'False
      Tab(0).Control(106)=   "C1(95)"
      Tab(0).Control(106).Enabled=   0   'False
      Tab(0).Control(107)=   "C1(94)"
      Tab(0).Control(107).Enabled=   0   'False
      Tab(0).Control(108)=   "C1(86)"
      Tab(0).Control(108).Enabled=   0   'False
      Tab(0).Control(109)=   "C1(85)"
      Tab(0).Control(109).Enabled=   0   'False
      Tab(0).Control(110)=   "C1(84)"
      Tab(0).Control(110).Enabled=   0   'False
      Tab(0).Control(111)=   "C1(83)"
      Tab(0).Control(111).Enabled=   0   'False
      Tab(0).Control(112)=   "C1(82)"
      Tab(0).Control(112).Enabled=   0   'False
      Tab(0).Control(113)=   "C1(81)"
      Tab(0).Control(113).Enabled=   0   'False
      Tab(0).Control(114)=   "C1(80)"
      Tab(0).Control(114).Enabled=   0   'False
      Tab(0).Control(115)=   "C1(79)"
      Tab(0).Control(115).Enabled=   0   'False
      Tab(0).Control(116)=   "C1(78)"
      Tab(0).Control(116).Enabled=   0   'False
      Tab(0).Control(117)=   "C1(77)"
      Tab(0).Control(117).Enabled=   0   'False
      Tab(0).Control(118)=   "C1(76)"
      Tab(0).Control(118).Enabled=   0   'False
      Tab(0).Control(119)=   "C1(75)"
      Tab(0).Control(119).Enabled=   0   'False
      Tab(0).Control(120)=   "C1(74)"
      Tab(0).Control(120).Enabled=   0   'False
      Tab(0).Control(121)=   "C1(73)"
      Tab(0).Control(121).Enabled=   0   'False
      Tab(0).Control(122)=   "C1(72)"
      Tab(0).Control(122).Enabled=   0   'False
      Tab(0).Control(123)=   "C1(71)"
      Tab(0).Control(123).Enabled=   0   'False
      Tab(0).Control(124)=   "C1(70)"
      Tab(0).Control(124).Enabled=   0   'False
      Tab(0).Control(125)=   "C1(69)"
      Tab(0).Control(125).Enabled=   0   'False
      Tab(0).Control(126)=   "C1(68)"
      Tab(0).Control(126).Enabled=   0   'False
      Tab(0).Control(127)=   "C1(60)"
      Tab(0).Control(127).Enabled=   0   'False
      Tab(0).Control(128)=   "C1(59)"
      Tab(0).Control(128).Enabled=   0   'False
      Tab(0).Control(129)=   "C1(58)"
      Tab(0).Control(129).Enabled=   0   'False
      Tab(0).Control(130)=   "C1(57)"
      Tab(0).Control(130).Enabled=   0   'False
      Tab(0).Control(131)=   "C1(56)"
      Tab(0).Control(131).Enabled=   0   'False
      Tab(0).Control(132)=   "C1(55)"
      Tab(0).Control(132).Enabled=   0   'False
      Tab(0).Control(133)=   "C1(54)"
      Tab(0).Control(133).Enabled=   0   'False
      Tab(0).Control(134)=   "C1(53)"
      Tab(0).Control(134).Enabled=   0   'False
      Tab(0).Control(135)=   "C1(52)"
      Tab(0).Control(135).Enabled=   0   'False
      Tab(0).Control(136)=   "C1(51)"
      Tab(0).Control(136).Enabled=   0   'False
      Tab(0).Control(137)=   "C1(50)"
      Tab(0).Control(137).Enabled=   0   'False
      Tab(0).Control(138)=   "C1(49)"
      Tab(0).Control(138).Enabled=   0   'False
      Tab(0).Control(139)=   "C1(48)"
      Tab(0).Control(139).Enabled=   0   'False
      Tab(0).Control(140)=   "C1(47)"
      Tab(0).Control(140).Enabled=   0   'False
      Tab(0).Control(141)=   "C1(46)"
      Tab(0).Control(141).Enabled=   0   'False
      Tab(0).Control(142)=   "C1(45)"
      Tab(0).Control(142).Enabled=   0   'False
      Tab(0).Control(143)=   "C1(44)"
      Tab(0).Control(143).Enabled=   0   'False
      Tab(0).Control(144)=   "C1(43)"
      Tab(0).Control(144).Enabled=   0   'False
      Tab(0).Control(145)=   "C1(42)"
      Tab(0).Control(145).Enabled=   0   'False
      Tab(0).Control(146)=   "C1(34)"
      Tab(0).Control(146).Enabled=   0   'False
      Tab(0).Control(147)=   "C1(33)"
      Tab(0).Control(147).Enabled=   0   'False
      Tab(0).Control(148)=   "C1(32)"
      Tab(0).Control(148).Enabled=   0   'False
      Tab(0).Control(149)=   "C1(31)"
      Tab(0).Control(149).Enabled=   0   'False
      Tab(0).Control(150)=   "C1(30)"
      Tab(0).Control(150).Enabled=   0   'False
      Tab(0).Control(151)=   "C1(29)"
      Tab(0).Control(151).Enabled=   0   'False
      Tab(0).Control(152)=   "C1(28)"
      Tab(0).Control(152).Enabled=   0   'False
      Tab(0).Control(153)=   "C1(27)"
      Tab(0).Control(153).Enabled=   0   'False
      Tab(0).Control(154)=   "C1(26)"
      Tab(0).Control(154).Enabled=   0   'False
      Tab(0).Control(155)=   "C1(25)"
      Tab(0).Control(155).Enabled=   0   'False
      Tab(0).Control(156)=   "C1(24)"
      Tab(0).Control(156).Enabled=   0   'False
      Tab(0).Control(157)=   "C1(23)"
      Tab(0).Control(157).Enabled=   0   'False
      Tab(0).Control(158)=   "C1(22)"
      Tab(0).Control(158).Enabled=   0   'False
      Tab(0).Control(159)=   "C1(21)"
      Tab(0).Control(159).Enabled=   0   'False
      Tab(0).Control(160)=   "C1(20)"
      Tab(0).Control(160).Enabled=   0   'False
      Tab(0).Control(161)=   "C1(19)"
      Tab(0).Control(161).Enabled=   0   'False
      Tab(0).Control(162)=   "C1(18)"
      Tab(0).Control(162).Enabled=   0   'False
      Tab(0).Control(163)=   "C1(17)"
      Tab(0).Control(163).Enabled=   0   'False
      Tab(0).Control(164)=   "C1(16)"
      Tab(0).Control(164).Enabled=   0   'False
      Tab(0).Control(165)=   "C1(10)"
      Tab(0).Control(165).Enabled=   0   'False
      Tab(0).Control(166)=   "C1(9)"
      Tab(0).Control(166).Enabled=   0   'False
      Tab(0).Control(167)=   "C1(8)"
      Tab(0).Control(167).Enabled=   0   'False
      Tab(0).Control(168)=   "C1(7)"
      Tab(0).Control(168).Enabled=   0   'False
      Tab(0).Control(169)=   "C1(6)"
      Tab(0).Control(169).Enabled=   0   'False
      Tab(0).Control(170)=   "C1(5)"
      Tab(0).Control(170).Enabled=   0   'False
      Tab(0).Control(171)=   "C1(4)"
      Tab(0).Control(171).Enabled=   0   'False
      Tab(0).Control(172)=   "C1(3)"
      Tab(0).Control(172).Enabled=   0   'False
      Tab(0).Control(173)=   "C1(2)"
      Tab(0).Control(173).Enabled=   0   'False
      Tab(0).Control(174)=   "C1(1)"
      Tab(0).Control(174).Enabled=   0   'False
      Tab(0).Control(175)=   "TA(10)"
      Tab(0).Control(175).Enabled=   0   'False
      Tab(0).Control(176)=   "TA(9)"
      Tab(0).Control(176).Enabled=   0   'False
      Tab(0).Control(177)=   "TA(8)"
      Tab(0).Control(177).Enabled=   0   'False
      Tab(0).Control(178)=   "TA(7)"
      Tab(0).Control(178).Enabled=   0   'False
      Tab(0).Control(179)=   "TA(6)"
      Tab(0).Control(179).Enabled=   0   'False
      Tab(0).Control(180)=   "TA(5)"
      Tab(0).Control(180).Enabled=   0   'False
      Tab(0).Control(181)=   "TA(2)"
      Tab(0).Control(181).Enabled=   0   'False
      Tab(0).Control(182)=   "TA(1)"
      Tab(0).Control(182).Enabled=   0   'False
      Tab(0).Control(183)=   "TA(16)"
      Tab(0).Control(183).Enabled=   0   'False
      Tab(0).Control(184)=   "TA(15)"
      Tab(0).Control(184).Enabled=   0   'False
      Tab(0).Control(185)=   "TA(14)"
      Tab(0).Control(185).Enabled=   0   'False
      Tab(0).Control(186)=   "TA(13)"
      Tab(0).Control(186).Enabled=   0   'False
      Tab(0).Control(187)=   "TA(12)"
      Tab(0).Control(187).Enabled=   0   'False
      Tab(0).Control(188)=   "TA(11)"
      Tab(0).Control(188).Enabled=   0   'False
      Tab(0).Control(189)=   "TA(4)"
      Tab(0).Control(189).Enabled=   0   'False
      Tab(0).Control(190)=   "TA(3)"
      Tab(0).Control(190).Enabled=   0   'False
      Tab(0).Control(191)=   "cmdAll"
      Tab(0).Control(191).Enabled=   0   'False
      Tab(0).ControlCount=   192
      TabCaption(1)   =   "内容2"
      TabPicture(1)   =   "NewGzd3.frx":0A92
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "C1(11)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "C1(162)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "C1(161)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "C1(160)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "C1(159)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "C1(158)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "C1(157)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "C1(156)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "C1(136)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "C1(135)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "C1(134)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "C1(133)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "C1(132)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "C1(131)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "C1(130)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "C1(110)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "C1(109)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "C1(108)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "C1(107)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "C1(106)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "C1(93)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "C1(92)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "C1(91)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "C1(90)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "C1(89)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "C1(88)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "C1(87)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "C1(67)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "C1(66)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "C1(65)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "C1(64)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "C1(63)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "C1(62)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "C1(61)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "C1(41)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "C1(40)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "C1(39)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "C1(38)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "C1(37)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "C1(36)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "C1(35)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "TA(17)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "C1(15)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "C1(14)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "C1(13)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "C1(12)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "TA(18)"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "TA(19)"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "TA(20)"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "TA(21)"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "TA(22)"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "TA(23)"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "TA(24)"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "TA(25)"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "TA(26)"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "TA(27)"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "BA(8)"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).Control(57)=   "BA(9)"
      Tab(1).Control(57).Enabled=   0   'False
      Tab(1).Control(58)=   "BA(10)"
      Tab(1).Control(58).Enabled=   0   'False
      Tab(1).Control(59)=   "BA(11)"
      Tab(1).Control(59).Enabled=   0   'False
      Tab(1).Control(60)=   "Frame1"
      Tab(1).Control(60).Enabled=   0   'False
      Tab(1).Control(61)=   "BA(12)"
      Tab(1).Control(61).Enabled=   0   'False
      Tab(1).Control(62)=   "BA(13)"
      Tab(1).Control(62).Enabled=   0   'False
      Tab(1).Control(63)=   "BA(14)"
      Tab(1).Control(63).Enabled=   0   'False
      Tab(1).Control(64)=   "BA(15)"
      Tab(1).Control(64).Enabled=   0   'False
      Tab(1).Control(65)=   "BA(16)"
      Tab(1).Control(65).Enabled=   0   'False
      Tab(1).Control(66)=   "dtpC"
      Tab(1).Control(66).Enabled=   0   'False
      Tab(1).Control(67)=   "dtpB"
      Tab(1).Control(67).Enabled=   0   'False
      Tab(1).Control(68)=   "Line72"
      Tab(1).Control(68).Enabled=   0   'False
      Tab(1).Control(69)=   "Line71"
      Tab(1).Control(69).Enabled=   0   'False
      Tab(1).Control(70)=   "Line70"
      Tab(1).Control(70).Enabled=   0   'False
      Tab(1).Control(71)=   "Line69"
      Tab(1).Control(71).Enabled=   0   'False
      Tab(1).Control(72)=   "Line68"
      Tab(1).Control(72).Enabled=   0   'False
      Tab(1).Control(73)=   "Line67"
      Tab(1).Control(73).Enabled=   0   'False
      Tab(1).Control(74)=   "Line37"
      Tab(1).Control(74).Enabled=   0   'False
      Tab(1).Control(75)=   "Shape2"
      Tab(1).Control(75).Enabled=   0   'False
      Tab(1).Control(76)=   "Line51"
      Tab(1).Control(76).Enabled=   0   'False
      Tab(1).Control(77)=   "Line50"
      Tab(1).Control(77).Enabled=   0   'False
      Tab(1).Control(78)=   "Line49"
      Tab(1).Control(78).Enabled=   0   'False
      Tab(1).Control(79)=   "Line48"
      Tab(1).Control(79).Enabled=   0   'False
      Tab(1).Control(80)=   "Line47"
      Tab(1).Control(80).Enabled=   0   'False
      Tab(1).Control(81)=   "Line46"
      Tab(1).Control(81).Enabled=   0   'False
      Tab(1).Control(82)=   "Line45"
      Tab(1).Control(82).Enabled=   0   'False
      Tab(1).Control(83)=   "Line44"
      Tab(1).Control(83).Enabled=   0   'False
      Tab(1).Control(84)=   "Line1"
      Tab(1).Control(84).Enabled=   0   'False
      Tab(1).Control(85)=   "Label20"
      Tab(1).Control(85).Enabled=   0   'False
      Tab(1).Control(86)=   "Label21"
      Tab(1).Control(86).Enabled=   0   'False
      Tab(1).Control(87)=   "Label22"
      Tab(1).Control(87).Enabled=   0   'False
      Tab(1).Control(88)=   "Label23"
      Tab(1).Control(88).Enabled=   0   'False
      Tab(1).Control(89)=   "Label24"
      Tab(1).Control(89).Enabled=   0   'False
      Tab(1).Control(90)=   "Label25"
      Tab(1).Control(90).Enabled=   0   'False
      Tab(1).Control(91)=   "Label29"
      Tab(1).Control(91).Enabled=   0   'False
      Tab(1).Control(92)=   "Label30"
      Tab(1).Control(92).Enabled=   0   'False
      Tab(1).Control(93)=   "Label31"
      Tab(1).Control(93).Enabled=   0   'False
      Tab(1).Control(94)=   "Label32"
      Tab(1).Control(94).Enabled=   0   'False
      Tab(1).Control(95)=   "Label34"
      Tab(1).Control(95).Enabled=   0   'False
      Tab(1).Control(96)=   "Label36"
      Tab(1).Control(96).Enabled=   0   'False
      Tab(1).Control(97)=   "Line34"
      Tab(1).Control(97).Enabled=   0   'False
      Tab(1).Control(98)=   "Line35"
      Tab(1).Control(98).Enabled=   0   'False
      Tab(1).Control(99)=   "Line36"
      Tab(1).Control(99).Enabled=   0   'False
      Tab(1).Control(100)=   "Label38"
      Tab(1).Control(100).Enabled=   0   'False
      Tab(1).Control(101)=   "Line52"
      Tab(1).Control(101).Enabled=   0   'False
      Tab(1).Control(102)=   "Line53"
      Tab(1).Control(102).Enabled=   0   'False
      Tab(1).Control(103)=   "Line60"
      Tab(1).Control(103).Enabled=   0   'False
      Tab(1).Control(104)=   "Line61"
      Tab(1).Control(104).Enabled=   0   'False
      Tab(1).Control(105)=   "Line62"
      Tab(1).Control(105).Enabled=   0   'False
      Tab(1).Control(106)=   "Line63"
      Tab(1).Control(106).Enabled=   0   'False
      Tab(1).Control(107)=   "Line64"
      Tab(1).Control(107).Enabled=   0   'False
      Tab(1).Control(108)=   "Line65"
      Tab(1).Control(108).Enabled=   0   'False
      Tab(1).Control(109)=   "Line66"
      Tab(1).Control(109).Enabled=   0   'False
      Tab(1).Control(110)=   "Label26"
      Tab(1).Control(110).Enabled=   0   'False
      Tab(1).Control(111)=   "Label35"
      Tab(1).Control(111).Enabled=   0   'False
      Tab(1).Control(112)=   "Label37"
      Tab(1).Control(112).Enabled=   0   'False
      Tab(1).ControlCount=   113
      Begin VB.CommandButton cmdAll 
         Caption         =   "全部"
         Height          =   255
         Left            =   7980
         TabIndex        =   269
         Top             =   8010
         Width           =   885
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "水温"
         Height          =   285
         Index           =   11
         Left            =   -73260
         TabIndex        =   230
         Top             =   840
         Width           =   675
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   225
         Index           =   162
         Left            =   -60720
         TabIndex        =   229
         Top             =   2250
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   161
         Left            =   -60720
         TabIndex        =   228
         Top             =   1875
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   160
         Left            =   -60720
         TabIndex        =   227
         Top             =   1515
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   159
         Left            =   -60720
         TabIndex        =   226
         Top             =   1155
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   158
         Left            =   -60720
         TabIndex        =   225
         Top             =   795
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   157
         Left            =   -60720
         TabIndex        =   224
         Top             =   435
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   156
         Left            =   -60720
         TabIndex        =   223
         Top             =   75
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   225
         Index           =   136
         Left            =   -61590
         TabIndex        =   222
         Top             =   2250
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   135
         Left            =   -61590
         TabIndex        =   221
         Top             =   1875
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   134
         Left            =   -61590
         TabIndex        =   220
         Top             =   1515
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   133
         Left            =   -61590
         TabIndex        =   219
         Top             =   1155
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   132
         Left            =   -61590
         TabIndex        =   218
         Top             =   795
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   131
         Left            =   -61590
         TabIndex        =   217
         Top             =   435
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   130
         Left            =   -61590
         TabIndex        =   216
         Top             =   75
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件损坏"
         Height          =   285
         Index           =   110
         Left            =   -63150
         TabIndex        =   215
         Top             =   1515
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件损坏"
         Height          =   285
         Index           =   109
         Left            =   -63150
         TabIndex        =   214
         Top             =   1155
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件损坏"
         Height          =   285
         Index           =   108
         Left            =   -63150
         TabIndex        =   213
         Top             =   795
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件损坏"
         Height          =   285
         Index           =   107
         Left            =   -63150
         TabIndex        =   212
         Top             =   435
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件损坏"
         Height          =   285
         Index           =   106
         Left            =   -63150
         TabIndex        =   211
         Top             =   75
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "进行中"
         Height          =   225
         Index           =   93
         Left            =   -64410
         TabIndex        =   210
         Top             =   2250
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "进行中"
         Height          =   285
         Index           =   92
         Left            =   -64410
         TabIndex        =   209
         Top             =   1875
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已检修"
         Height          =   285
         Index           =   91
         Left            =   -64410
         TabIndex        =   208
         Top             =   1515
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已检修"
         Height          =   285
         Index           =   90
         Left            =   -64410
         TabIndex        =   207
         Top             =   1155
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已检修"
         Height          =   285
         Index           =   89
         Left            =   -64410
         TabIndex        =   206
         Top             =   795
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已调整"
         Height          =   285
         Index           =   88
         Left            =   -64410
         TabIndex        =   205
         Top             =   435
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已检修"
         Height          =   285
         Index           =   87
         Left            =   -64410
         TabIndex        =   204
         Top             =   75
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已完成"
         Height          =   225
         Index           =   67
         Left            =   -65550
         TabIndex        =   203
         Top             =   2250
         Width           =   885
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已完成"
         Height          =   285
         Index           =   66
         Left            =   -65550
         TabIndex        =   202
         Top             =   1875
         Width           =   885
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   65
         Left            =   -65340
         TabIndex        =   201
         Top             =   1515
         Width           =   675
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   64
         Left            =   -65340
         TabIndex        =   200
         Top             =   1155
         Width           =   675
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   63
         Left            =   -65340
         TabIndex        =   199
         Top             =   795
         Width           =   675
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   62
         Left            =   -65340
         TabIndex        =   198
         Top             =   435
         Width           =   675
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   61
         Left            =   -65340
         TabIndex        =   197
         Top             =   75
         Width           =   675
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   225
         Index           =   41
         Left            =   -66720
         TabIndex        =   196
         Top             =   2250
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   40
         Left            =   -66720
         TabIndex        =   195
         Top             =   1875
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   39
         Left            =   -66720
         TabIndex        =   194
         Top             =   1515
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   38
         Left            =   -66720
         TabIndex        =   193
         Top             =   1155
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   37
         Left            =   -66720
         TabIndex        =   192
         Top             =   795
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   36
         Left            =   -66720
         TabIndex        =   191
         Top             =   435
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   35
         Left            =   -66720
         TabIndex        =   190
         Top             =   75
         Width           =   255
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   17
         Left            =   -69510
         TabIndex        =   189
         Top             =   870
         Width           =   1125
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "其它"
         Height          =   285
         Index           =   15
         Left            =   -70290
         TabIndex        =   188
         Top             =   840
         Width           =   675
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "液位"
         Height          =   285
         Index           =   14
         Left            =   -71100
         TabIndex        =   187
         Top             =   840
         Width           =   675
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "低压"
         Height          =   285
         Index           =   13
         Left            =   -71790
         TabIndex        =   186
         Top             =   840
         Width           =   675
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "高压"
         Height          =   285
         Index           =   12
         Left            =   -72540
         TabIndex        =   185
         Top             =   840
         Width           =   675
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   18
         Left            =   -74760
         TabIndex        =   252
         Top             =   4800
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   19
         Left            =   -73530
         TabIndex        =   253
         Top             =   4800
         Width           =   4755
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   20
         Left            =   -68670
         TabIndex        =   254
         Top             =   4800
         Width           =   2235
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   21
         Left            =   -66210
         TabIndex        =   255
         Top             =   4800
         Width           =   4365
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   22
         Left            =   -61770
         TabIndex        =   256
         Top             =   4800
         Width           =   1605
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   23
         Left            =   -74760
         TabIndex        =   257
         Top             =   5040
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   24
         Left            =   -73530
         TabIndex        =   258
         Top             =   5040
         Width           =   4755
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   25
         Left            =   -68670
         TabIndex        =   259
         Top             =   5040
         Width           =   2235
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   26
         Left            =   -66210
         TabIndex        =   260
         Top             =   5040
         Width           =   4365
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   27
         Left            =   -61770
         TabIndex        =   261
         Top             =   5040
         Width           =   1605
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   8
         Left            =   -73530
         TabIndex        =   262
         Text            =   "的"
         Top             =   5340
         Width           =   1035
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   9
         Left            =   -70290
         TabIndex        =   263
         Text            =   "的"
         Top             =   5340
         Width           =   1035
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   10
         Left            =   -67620
         TabIndex        =   264
         Text            =   "的"
         Top             =   5340
         Width           =   1035
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   11
         Left            =   -65220
         TabIndex        =   265
         Text            =   "的"
         Top             =   5340
         Width           =   1035
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Left            =   -74910
         TabIndex        =   156
         Top             =   5580
         Width           =   10725
         Begin VB.OptionButton FPA 
            Caption         =   "优秀"
            Height          =   195
            Left            =   1350
            TabIndex        =   160
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton FPB 
            Caption         =   "满意"
            Height          =   195
            Left            =   2950
            TabIndex        =   159
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton FPC 
            Caption         =   "较满意"
            Height          =   195
            Left            =   4550
            TabIndex        =   158
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton FPD 
            Caption         =   "尚可"
            Height          =   195
            Left            =   6150
            TabIndex        =   157
            Top             =   60
            Width           =   1065
         End
         Begin VB.Label Label33 
            Caption         =   "服务评价:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            TabIndex        =   161
            Top             =   60
            Width           =   1395
         End
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   570
         Index           =   12
         Left            =   -73530
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   266
         Text            =   "NewGzd3.frx":0AAE
         Top             =   5910
         Width           =   9345
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   13
         Left            =   -64080
         TabIndex        =   267
         Top             =   5580
         Width           =   1935
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   14
         Left            =   -64110
         TabIndex        =   155
         Top             =   6150
         Width           =   1815
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   15
         Left            =   -62010
         TabIndex        =   268
         Top             =   5580
         Width           =   1845
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   16
         Left            =   -62010
         TabIndex        =   154
         Top             =   6150
         Width           =   1695
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   3
         Left            =   1710
         TabIndex        =   238
         Top             =   780
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   4
         Left            =   3150
         TabIndex        =   239
         Top             =   780
         Width           =   1065
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   11
         Left            =   5760
         TabIndex        =   241
         Top             =   780
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   12
         Left            =   7080
         TabIndex        =   243
         Top             =   780
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   13
         Left            =   8430
         TabIndex        =   245
         Top             =   780
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   14
         Left            =   9750
         TabIndex        =   247
         Top             =   780
         Width           =   1065
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   15
         Left            =   11040
         TabIndex        =   249
         Top             =   780
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   16
         Left            =   12390
         TabIndex        =   251
         Top             =   780
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   1
         Left            =   1695
         TabIndex        =   236
         Top             =   480
         Width           =   1155
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   2
         Left            =   3135
         TabIndex        =   237
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   5
         Left            =   5745
         TabIndex        =   240
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   6
         Left            =   7065
         TabIndex        =   242
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   7
         Left            =   8400
         TabIndex        =   244
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   8
         Left            =   9735
         TabIndex        =   246
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   9
         Left            =   11055
         TabIndex        =   248
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   10
         Left            =   12390
         TabIndex        =   250
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "更换蒸发器水侧端盖垫片"
         Height          =   285
         Index           =   1
         Left            =   60
         TabIndex        =   134
         Top             =   1980
         Width           =   2355
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "更换冷凝器水侧端盖垫片"
         Height          =   285
         Index           =   2
         Left            =   3000
         TabIndex        =   133
         Top             =   1980
         Width           =   2475
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "1#"
         Height          =   285
         Index           =   3
         Left            =   4290
         TabIndex        =   132
         Top             =   5550
         Width           =   495
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "2#"
         Height          =   285
         Index           =   4
         Left            =   4815
         TabIndex        =   131
         Top             =   5550
         Width           =   495
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "3#"
         Height          =   285
         Index           =   5
         Left            =   5355
         TabIndex        =   130
         Top             =   5550
         Width           =   495
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "4#"
         Height          =   285
         Index           =   6
         Left            =   5880
         TabIndex        =   129
         Top             =   5550
         Width           =   495
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "1#"
         Height          =   285
         Index           =   7
         Left            =   5010
         TabIndex        =   128
         Top             =   5910
         Width           =   495
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "2#"
         Height          =   285
         Index           =   8
         Left            =   5595
         TabIndex        =   127
         Top             =   5910
         Width           =   495
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "3#"
         Height          =   285
         Index           =   9
         Left            =   6165
         TabIndex        =   126
         Top             =   5910
         Width           =   495
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "4#"
         Height          =   285
         Index           =   10
         Left            =   6750
         TabIndex        =   125
         Top             =   5910
         Width           =   495
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   240
         Index           =   16
         Left            =   8280
         TabIndex        =   124
         Top             =   1260
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   225
         Index           =   17
         Left            =   8280
         TabIndex        =   123
         Top             =   1620
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   225
         Index           =   18
         Left            =   8280
         TabIndex        =   122
         Top             =   1980
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   225
         Index           =   19
         Left            =   8280
         TabIndex        =   121
         Top             =   2340
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   225
         Index           =   20
         Left            =   8280
         TabIndex        =   120
         Top             =   2685
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   225
         Index           =   21
         Left            =   8280
         TabIndex        =   119
         Top             =   3045
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   22
         Left            =   8280
         TabIndex        =   118
         Top             =   3405
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   8280
         TabIndex        =   117
         Top             =   3765
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   8280
         TabIndex        =   116
         Top             =   4125
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   25
         Left            =   8280
         TabIndex        =   115
         Top             =   4485
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   26
         Left            =   8280
         TabIndex        =   114
         Top             =   4845
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   27
         Left            =   8280
         TabIndex        =   113
         Top             =   5205
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   28
         Left            =   8280
         TabIndex        =   112
         Top             =   5550
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   29
         Left            =   8280
         TabIndex        =   111
         Top             =   5910
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   30
         Left            =   8280
         TabIndex        =   110
         Top             =   6270
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   31
         Left            =   8280
         TabIndex        =   109
         Top             =   6630
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   32
         Left            =   8280
         TabIndex        =   108
         Top             =   6990
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   33
         Left            =   8280
         TabIndex        =   107
         Top             =   7350
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   34
         Left            =   8280
         TabIndex        =   106
         Top             =   7710
         Width           =   255
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "化学清洗"
         Height          =   240
         Index           =   42
         Left            =   9300
         TabIndex        =   105
         Top             =   1260
         Width           =   1035
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "化学清洗"
         Height          =   225
         Index           =   43
         Left            =   9300
         TabIndex        =   104
         Top             =   1620
         Width           =   1035
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无须更换"
         Height          =   225
         Index           =   44
         Left            =   9300
         TabIndex        =   103
         Top             =   1980
         Width           =   1035
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   225
         Index           =   45
         Left            =   9660
         TabIndex        =   102
         Top             =   2340
         Width           =   675
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无漏点"
         Height          =   225
         Index           =   46
         Left            =   9480
         TabIndex        =   101
         Top             =   2685
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   225
         Index           =   47
         Left            =   9660
         TabIndex        =   100
         Top             =   3045
         Width           =   675
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   48
         Left            =   9660
         TabIndex        =   99
         Top             =   3405
         Width           =   675
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   49
         Left            =   9660
         TabIndex        =   98
         Top             =   3765
         Width           =   675
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   50
         Left            =   9660
         TabIndex        =   97
         Top             =   4125
         Width           =   675
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   51
         Left            =   9660
         TabIndex        =   96
         Top             =   4485
         Width           =   675
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   52
         Left            =   9660
         TabIndex        =   95
         Top             =   4845
         Width           =   675
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无须更换"
         Height          =   285
         Index           =   53
         Left            =   9300
         TabIndex        =   94
         Top             =   5205
         Width           =   1035
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "均正常"
         Height          =   285
         Index           =   54
         Left            =   9480
         TabIndex        =   93
         Top             =   5550
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "均正常"
         Height          =   285
         Index           =   55
         Left            =   9480
         TabIndex        =   92
         Top             =   5910
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "合格"
         Height          =   285
         Index           =   56
         Left            =   9660
         TabIndex        =   91
         Top             =   6270
         Width           =   675
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "合格"
         Height          =   285
         Index           =   57
         Left            =   9660
         TabIndex        =   90
         Top             =   6630
         Width           =   675
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已清洁"
         Height          =   285
         Index           =   58
         Left            =   9480
         TabIndex        =   89
         Top             =   6990
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   59
         Left            =   9660
         TabIndex        =   88
         Top             =   7350
         Width           =   675
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   60
         Left            =   9660
         TabIndex        =   87
         Top             =   7710
         Width           =   675
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "机械清洗"
         Height          =   240
         Index           =   68
         Left            =   10590
         TabIndex        =   86
         Top             =   1260
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "机械清洗"
         Height          =   225
         Index           =   69
         Left            =   10590
         TabIndex        =   85
         Top             =   1620
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已更换"
         Height          =   225
         Index           =   70
         Left            =   10590
         TabIndex        =   84
         Top             =   1980
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已更换"
         Height          =   225
         Index           =   71
         Left            =   10590
         TabIndex        =   83
         Top             =   2340
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "需要烧焊"
         Height          =   225
         Index           =   72
         Left            =   10590
         TabIndex        =   82
         Top             =   2685
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已排气"
         Height          =   225
         Index           =   73
         Left            =   10590
         TabIndex        =   81
         Top             =   3045
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已添加"
         Height          =   285
         Index           =   74
         Left            =   10590
         TabIndex        =   80
         Top             =   3405
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "尚可"
         Height          =   285
         Index           =   75
         Left            =   10590
         TabIndex        =   79
         Top             =   3765
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已取油样"
         Height          =   285
         Index           =   76
         Left            =   10590
         TabIndex        =   78
         Top             =   4125
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已检修"
         Height          =   285
         Index           =   77
         Left            =   10590
         TabIndex        =   77
         Top             =   4485
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已检修"
         Height          =   285
         Index           =   78
         Left            =   10590
         TabIndex        =   76
         Top             =   4845
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "垫片未换"
         Height          =   285
         Index           =   79
         Left            =   10590
         TabIndex        =   75
         Top             =   5205
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已检修"
         Height          =   285
         Index           =   80
         Left            =   10590
         TabIndex        =   74
         Top             =   5550
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已检修"
         Height          =   285
         Index           =   81
         Left            =   10590
         TabIndex        =   73
         Top             =   5910
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已检修"
         Height          =   285
         Index           =   82
         Left            =   10590
         TabIndex        =   72
         Top             =   6270
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "不合格"
         Height          =   285
         Index           =   83
         Left            =   10590
         TabIndex        =   71
         Top             =   6630
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已打磨"
         Height          =   285
         Index           =   84
         Left            =   10590
         TabIndex        =   70
         Top             =   6990
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已检修"
         Height          =   285
         Index           =   85
         Left            =   10590
         TabIndex        =   69
         Top             =   7350
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已检修"
         Height          =   285
         Index           =   86
         Left            =   10590
         TabIndex        =   68
         Top             =   7710
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "视镜损坏"
         Height          =   225
         Index           =   94
         Left            =   11850
         TabIndex        =   67
         Top             =   2340
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已放出"
         Height          =   285
         Index           =   95
         Left            =   11850
         TabIndex        =   66
         Top             =   3405
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已换油"
         Height          =   285
         Index           =   96
         Left            =   11850
         TabIndex        =   65
         Top             =   4125
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件损坏"
         Height          =   285
         Index           =   97
         Left            =   11850
         TabIndex        =   64
         Top             =   4485
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "都已更换"
         Height          =   285
         Index           =   98
         Left            =   11850
         TabIndex        =   63
         Top             =   4845
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件损坏"
         Height          =   285
         Index           =   99
         Left            =   11850
         TabIndex        =   62
         Top             =   5205
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件损坏"
         Height          =   285
         Index           =   100
         Left            =   11850
         TabIndex        =   61
         Top             =   5550
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件损坏"
         Height          =   285
         Index           =   101
         Left            =   11850
         TabIndex        =   60
         Top             =   5910
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件损坏"
         Height          =   285
         Index           =   102
         Left            =   11850
         TabIndex        =   59
         Top             =   6270
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已更换"
         Height          =   285
         Index           =   103
         Left            =   11850
         TabIndex        =   58
         Top             =   6990
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件损坏"
         Height          =   285
         Index           =   104
         Left            =   11850
         TabIndex        =   57
         Top             =   7350
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件损坏"
         Height          =   285
         Index           =   105
         Left            =   11850
         TabIndex        =   56
         Top             =   7710
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   225
         Index           =   111
         Left            =   13410
         TabIndex        =   55
         Top             =   1260
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   225
         Index           =   112
         Left            =   13410
         TabIndex        =   54
         Top             =   1620
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   225
         Index           =   113
         Left            =   13410
         TabIndex        =   53
         Top             =   1980
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   225
         Index           =   114
         Left            =   13410
         TabIndex        =   52
         Top             =   2340
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   225
         Index           =   115
         Left            =   13410
         TabIndex        =   51
         Top             =   2685
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   225
         Index           =   116
         Left            =   13410
         TabIndex        =   50
         Top             =   3045
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   117
         Left            =   13410
         TabIndex        =   49
         Top             =   3405
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   118
         Left            =   13410
         TabIndex        =   48
         Top             =   3765
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   119
         Left            =   13410
         TabIndex        =   47
         Top             =   4125
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   120
         Left            =   13410
         TabIndex        =   46
         Top             =   4485
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   121
         Left            =   13410
         TabIndex        =   45
         Top             =   4845
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   122
         Left            =   13410
         TabIndex        =   44
         Top             =   5205
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   123
         Left            =   13410
         TabIndex        =   43
         Top             =   5550
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   124
         Left            =   13410
         TabIndex        =   42
         Top             =   5910
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   125
         Left            =   13410
         TabIndex        =   41
         Top             =   6270
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   126
         Left            =   13410
         TabIndex        =   40
         Top             =   6630
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   127
         Left            =   13410
         TabIndex        =   39
         Top             =   6990
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   128
         Left            =   13410
         TabIndex        =   38
         Top             =   7350
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   129
         Left            =   13410
         TabIndex        =   37
         Top             =   7710
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   225
         Index           =   137
         Left            =   14280
         TabIndex        =   36
         Top             =   1260
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   225
         Index           =   138
         Left            =   14280
         TabIndex        =   35
         Top             =   1620
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   225
         Index           =   139
         Left            =   14280
         TabIndex        =   34
         Top             =   1980
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   225
         Index           =   140
         Left            =   14280
         TabIndex        =   33
         Top             =   2340
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   225
         Index           =   141
         Left            =   14280
         TabIndex        =   32
         Top             =   2685
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   225
         Index           =   142
         Left            =   14280
         TabIndex        =   31
         Top             =   3045
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   143
         Left            =   14280
         TabIndex        =   30
         Top             =   3405
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   144
         Left            =   14280
         TabIndex        =   29
         Top             =   3765
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   145
         Left            =   14280
         TabIndex        =   28
         Top             =   4125
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   146
         Left            =   14280
         TabIndex        =   27
         Top             =   4485
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   147
         Left            =   14280
         TabIndex        =   26
         Top             =   4845
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   148
         Left            =   14280
         TabIndex        =   25
         Top             =   5205
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   149
         Left            =   14280
         TabIndex        =   24
         Top             =   5550
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   150
         Left            =   14280
         TabIndex        =   23
         Top             =   5910
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   151
         Left            =   14280
         TabIndex        =   22
         Top             =   6270
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   152
         Left            =   14280
         TabIndex        =   21
         Top             =   6630
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   153
         Left            =   14280
         TabIndex        =   20
         Top             =   6990
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   154
         Left            =   14280
         TabIndex        =   19
         Top             =   7350
         Width           =   315
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   155
         Left            =   14280
         TabIndex        =   18
         Top             =   7710
         Width           =   315
      End
      Begin MSComCtl2.DTPicker dtpC 
         Height          =   225
         Left            =   -62010
         TabIndex        =   163
         Top             =   6150
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   397
         _Version        =   393216
         Format          =   150011905
         CurrentDate     =   38897
      End
      Begin MSComCtl2.DTPicker dtpB 
         Height          =   225
         Left            =   -64140
         TabIndex        =   162
         Top             =   6150
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   397
         _Version        =   393216
         Format          =   150011905
         CurrentDate     =   38897
      End
      Begin VB.Line Line72 
         X1              =   -60930
         X2              =   -60930
         Y1              =   60
         Y2              =   2520
      End
      Begin VB.Line Line71 
         X1              =   -61800
         X2              =   -61800
         Y1              =   60
         Y2              =   2520
      End
      Begin VB.Line Line70 
         X1              =   -63270
         X2              =   -63270
         Y1              =   60
         Y2              =   2520
      End
      Begin VB.Line Line69 
         X1              =   -64500
         X2              =   -64500
         Y1              =   60
         Y2              =   2520
      End
      Begin VB.Line Line68 
         X1              =   -65790
         X2              =   -65790
         Y1              =   60
         Y2              =   2520
      End
      Begin VB.Line Line67 
         X1              =   -67260
         X2              =   -67260
         Y1              =   60
         Y2              =   2535
      End
      Begin VB.Line Line37 
         X1              =   -62070
         X2              =   -62070
         Y1              =   5280
         Y2              =   6480
      End
      Begin VB.Shape Shape2 
         Height          =   2025
         Left            =   -74970
         Top             =   4500
         Width           =   15045
      End
      Begin VB.Line Line51 
         X1              =   -74970
         X2              =   -60030
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line50 
         X1              =   -74970
         X2              =   -60015
         Y1              =   2190
         Y2              =   2190
      End
      Begin VB.Line Line49 
         X1              =   -74970
         X2              =   -60030
         Y1              =   1830
         Y2              =   1830
      End
      Begin VB.Line Line48 
         X1              =   -74970
         X2              =   -60030
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Line Line47 
         X1              =   -74970
         X2              =   -60030
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Line Line46 
         X1              =   -74970
         X2              =   -60030
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Line Line45 
         X1              =   -74970
         X2              =   -60030
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Line Line44 
         X1              =   -74970
         X2              =   -60030
         Y1              =   60
         Y2              =   60
      End
      Begin VB.Line Line1 
         X1              =   -68400
         X2              =   -69570
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label20 
         Caption         =   "保养中发生的零配件的清单"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -74910
         TabIndex        =   178
         Top             =   4260
         Width           =   2475
      End
      Begin VB.Label Label21 
         Caption         =   "数量"
         Height          =   195
         Left            =   -74700
         TabIndex        =   177
         Top             =   4560
         Width           =   1035
      End
      Begin VB.Label Label22 
         Caption         =   "零配件或材料名称"
         Height          =   195
         Left            =   -72420
         TabIndex        =   176
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Label Label23 
         Caption         =   "零件编号"
         Height          =   195
         Left            =   -68640
         TabIndex        =   175
         Top             =   4560
         Width           =   1995
      End
      Begin VB.Label Label24 
         Caption         =   "使用情况"
         Height          =   195
         Left            =   -65880
         TabIndex        =   174
         Top             =   4560
         Width           =   1725
      End
      Begin VB.Label Label25 
         Caption         =   "供货方"
         Height          =   195
         Left            =   -61650
         TabIndex        =   173
         Top             =   4560
         Width           =   1005
      End
      Begin VB.Label Label29 
         Caption         =   "到达时间"
         Height          =   165
         Left            =   -74820
         TabIndex        =   172
         Top             =   5340
         Width           =   1035
      End
      Begin VB.Label Label30 
         Caption         =   "完成时间"
         Height          =   165
         Left            =   -71850
         TabIndex        =   171
         Top             =   5340
         Width           =   1035
      End
      Begin VB.Label Label31 
         Caption         =   "旅途时间"
         Height          =   165
         Left            =   -68730
         TabIndex        =   170
         Top             =   5340
         Width           =   1035
      End
      Begin VB.Label Label32 
         Caption         =   "加班工时"
         Height          =   165
         Left            =   -66300
         TabIndex        =   169
         Top             =   5340
         Width           =   1035
      End
      Begin VB.Label Label34 
         Caption         =   "客户意见或建议："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -74970
         TabIndex        =   168
         Top             =   5910
         Width           =   885
      End
      Begin VB.Label Label36 
         Caption         =   "日期："
         Height          =   195
         Left            =   -64080
         TabIndex        =   166
         Top             =   5910
         Width           =   945
      End
      Begin VB.Line Line34 
         X1              =   -74970
         X2              =   -60030
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Line Line35 
         X1              =   -74970
         X2              =   -60030
         Y1              =   5850
         Y2              =   5850
      End
      Begin VB.Line Line36 
         X1              =   -64170
         X2              =   -64170
         Y1              =   5280
         Y2              =   6480
      End
      Begin VB.Label Label38 
         Caption         =   "日期："
         Height          =   195
         Left            =   -62010
         TabIndex        =   164
         Top             =   5910
         Width           =   945
      End
      Begin VB.Line Line52 
         X1              =   -74970
         X2              =   -60030
         Y1              =   4500
         Y2              =   4500
      End
      Begin VB.Line Line53 
         X1              =   -74970
         X2              =   -60030
         Y1              =   4770
         Y2              =   4770
      End
      Begin VB.Line Line60 
         X1              =   -73560
         X2              =   -73560
         Y1              =   4500
         Y2              =   5265
      End
      Begin VB.Line Line61 
         X1              =   -68730
         X2              =   -68730
         Y1              =   4500
         Y2              =   5250
      End
      Begin VB.Line Line62 
         X1              =   -74970
         X2              =   -60030
         Y1              =   5010
         Y2              =   5010
      End
      Begin VB.Line Line63 
         X1              =   -74940
         X2              =   -60030
         Y1              =   5250
         Y2              =   5250
      End
      Begin VB.Line Line64 
         X1              =   -66330
         X2              =   -66330
         Y1              =   5250
         Y2              =   4500
      End
      Begin VB.Line Line65 
         X1              =   -61800
         X2              =   -61785
         Y1              =   4500
         Y2              =   4500
      End
      Begin VB.Line Line66 
         X1              =   -61800
         X2              =   -61800
         Y1              =   4500
         Y2              =   5280
      End
      Begin VB.Label Label1 
         Caption         =   "基本资料"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   153
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "类型"
         Height          =   225
         Left            =   1695
         TabIndex        =   152
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "充注量"
         Height          =   225
         Left            =   3090
         TabIndex        =   151
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "1#压缩机"
         Height          =   225
         Left            =   5760
         TabIndex        =   150
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "2#压缩机"
         Height          =   225
         Left            =   7095
         TabIndex        =   149
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label6 
         Caption         =   "3#压缩机"
         Height          =   225
         Left            =   8430
         TabIndex        =   148
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label7 
         Caption         =   "4#压缩机"
         Height          =   225
         Left            =   9765
         TabIndex        =   147
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label8 
         Caption         =   "5#压缩机"
         Height          =   225
         Left            =   11115
         TabIndex        =   146
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label9 
         Caption         =   "6#压缩机"
         Height          =   225
         Left            =   12450
         TabIndex        =   145
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label10 
         Caption         =   "制冷剂"
         Height          =   195
         Left            =   120
         TabIndex        =   144
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label Label11 
         Caption         =   "润滑油"
         Height          =   210
         Left            =   120
         TabIndex        =   143
         Top             =   780
         Width           =   1125
      End
      Begin VB.Label Label12 
         Caption         =   "压缩机型号"
         Height          =   195
         Left            =   4470
         TabIndex        =   142
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label Label13 
         Caption         =   "额定满载电流"
         Height          =   210
         Left            =   4410
         TabIndex        =   141
         Top             =   780
         Width           =   1125
      End
      Begin VB.Label Label14 
         Caption         =   "年度保养服务内容"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   120
         TabIndex        =   140
         Top             =   1050
         Width           =   2115
      End
      Begin VB.Label Label16 
         Caption         =   "今日工作内容"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   10470
         TabIndex        =   138
         Top             =   1050
         Width           =   1245
      End
      Begin VB.Label Label17 
         Caption         =   "待完成"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   13290
         TabIndex        =   137
         Top             =   1050
         Width           =   615
      End
      Begin VB.Label Label18 
         Caption         =   "无此项"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   14100
         TabIndex        =   136
         Top             =   1050
         Width           =   675
      End
      Begin VB.Shape Shape1 
         Height          =   8025
         Left            =   30
         Top             =   30
         Width           =   14955
      End
      Begin VB.Line Line2 
         X1              =   30
         X2              =   14970
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Line Line3 
         X1              =   30
         X2              =   14970
         Y1              =   390
         Y2              =   390
      End
      Begin VB.Line Line4 
         X1              =   30
         X2              =   14970
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line5 
         X1              =   1500
         X2              =   1500
         Y1              =   1020
         Y2              =   30
      End
      Begin VB.Line Line6 
         X1              =   2970
         X2              =   2970
         Y1              =   1050
         Y2              =   30
      End
      Begin VB.Line Line7 
         X1              =   4380
         X2              =   4380
         Y1              =   990
         Y2              =   30
      End
      Begin VB.Line Line8 
         X1              =   5640
         X2              =   5640
         Y1              =   990
         Y2              =   30
      End
      Begin VB.Line Line9 
         X1              =   6960
         X2              =   6960
         Y1              =   30
         Y2              =   1020
      End
      Begin VB.Line Line10 
         X1              =   8310
         X2              =   8310
         Y1              =   30
         Y2              =   1020
      End
      Begin VB.Line Line11 
         X1              =   9600
         X2              =   9600
         Y1              =   30
         Y2              =   1020
      End
      Begin VB.Line Line12 
         X1              =   10920
         X2              =   10920
         Y1              =   30
         Y2              =   1020
      End
      Begin VB.Line Line13 
         X1              =   12270
         X2              =   12270
         Y1              =   30
         Y2              =   1020
      End
      Begin VB.Line Line14 
         X1              =   13710
         X2              =   13710
         Y1              =   30
         Y2              =   1020
      End
      Begin VB.Line Line15 
         X1              =   30
         X2              =   14970
         Y1              =   1530
         Y2              =   1530
      End
      Begin VB.Line Line16 
         X1              =   30
         X2              =   14980
         Y1              =   1230
         Y2              =   1230
      End
      Begin VB.Line Line17 
         X1              =   30
         X2              =   15000
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Line Line18 
         X1              =   30
         X2              =   14970
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line19 
         X1              =   30
         X2              =   14970
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line20 
         X1              =   30
         X2              =   14970
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line21 
         X1              =   30
         X2              =   14970
         Y1              =   3330
         Y2              =   3330
      End
      Begin VB.Line Line22 
         X1              =   30
         X2              =   14970
         Y1              =   3690
         Y2              =   3690
      End
      Begin VB.Line Line23 
         X1              =   30
         X2              =   14970
         Y1              =   4050
         Y2              =   4050
      End
      Begin VB.Line Line24 
         X1              =   30
         X2              =   14970
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line25 
         X1              =   30
         X2              =   15000
         Y1              =   4770
         Y2              =   4770
      End
      Begin VB.Line Line29 
         X1              =   30
         X2              =   14970
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Line Line33 
         X1              =   30
         X2              =   14970
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Line Line38 
         Index           =   0
         X1              =   30
         X2              =   14970
         Y1              =   5850
         Y2              =   5850
      End
      Begin VB.Line Line39 
         X1              =   30
         X2              =   14970
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Line Line40 
         X1              =   30
         X2              =   14985
         Y1              =   6570
         Y2              =   6570
      End
      Begin VB.Line Line41 
         X1              =   30
         X2              =   14985
         Y1              =   6960
         Y2              =   6960
      End
      Begin VB.Line Line42 
         X1              =   30
         X2              =   14970
         Y1              =   7320
         Y2              =   7320
      End
      Begin VB.Line Line43 
         X1              =   30
         X2              =   14970
         Y1              =   7650
         Y2              =   7650
      End
      Begin VB.Line Line54 
         X1              =   7650
         X2              =   7650
         Y1              =   1020
         Y2              =   8040
      End
      Begin VB.Line Line55 
         X1              =   9060
         X2              =   9060
         Y1              =   1020
         Y2              =   8040
      End
      Begin VB.Line Line56 
         X1              =   10410
         X2              =   10410
         Y1              =   1020
         Y2              =   8040
      End
      Begin VB.Line Line57 
         X1              =   11760
         X2              =   11760
         Y1              =   1230
         Y2              =   8040
      End
      Begin VB.Line Line58 
         X1              =   13140
         X2              =   13140
         Y1              =   1230
         Y2              =   8040
      End
      Begin VB.Line Line59 
         X1              =   14040
         X2              =   14040
         Y1              =   1230
         Y2              =   8040
      End
      Begin VB.Label Label19 
         Caption         =   $"NewGzd3.frx":0AB1
         Height          =   6975
         Left            =   90
         TabIndex        =   135
         Top             =   1290
         Width           =   6135
      End
      Begin VB.Label Label26 
         Caption         =   $"NewGzd3.frx":0DC8
         Height          =   3465
         Left            =   -74880
         TabIndex        =   231
         Top             =   180
         Width           =   6915
      End
      Begin VB.Label Label15 
         Caption         =   "往日已完成"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7920
         TabIndex        =   139
         Top             =   1050
         Width           =   1425
      End
      Begin VB.Label Label35 
         Caption         =   "客户签名："
         Height          =   225
         Left            =   -64080
         TabIndex        =   167
         Top             =   5310
         Width           =   945
      End
      Begin VB.Label Label37 
         Caption         =   "质量控制签名："
         Height          =   195
         Left            =   -62010
         TabIndex        =   165
         Top             =   5340
         Width           =   1275
      End
   End
   Begin MSDataGridLib.DataGrid comHtbh 
      Height          =   1155
      Left            =   6120
      TabIndex        =   16
      Top             =   60
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2037
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "合同编号"
         Caption         =   "合同编号"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "合同金额"
         Caption         =   "合同金额"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "khdh"
         Caption         =   "khdh"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         ScrollBars      =   2
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox comXmmc 
      Height          =   300
      Left            =   1980
      TabIndex        =   15
      Top             =   750
      Width           =   4125
   End
   Begin VB.TextBox BA 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   12480
      TabIndex        =   11
      Text            =   "的"
      Top             =   90
      Visible         =   0   'False
      Width           =   4245
   End
   Begin VB.TextBox BA 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   8130
      Locked          =   -1  'True
      TabIndex        =   10
      Tag             =   "20"
      Top             =   960
      Width           =   4245
   End
   Begin VB.TextBox BA 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   8100
      TabIndex        =   9
      Top             =   570
      Width           =   4245
   End
   Begin VB.TextBox BA 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   8100
      TabIndex        =   8
      Text            =   "的"
      Top             =   150
      Width           =   4245
   End
   Begin VB.TextBox BA 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   2010
      TabIndex        =   7
      Top             =   930
      Width           =   4065
   End
   Begin VB.TextBox BA 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   2010
      TabIndex        =   6
      Top             =   510
      Width           =   4065
   End
   Begin VB.TextBox BA 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   12990
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   4065
   End
   Begin VB.TextBox BA 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   2010
      TabIndex        =   4
      Top             =   90
      Width           =   4065
   End
   Begin VB.TextBox TA 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   13800
      TabIndex        =   3
      Top             =   570
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CheckBox C1 
      Alignment       =   1  'Right Justify
      Caption         =   "1号"
      Height          =   285
      Index           =   0
      Left            =   12930
      TabIndex        =   2
      Top             =   780
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   6570
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "NewGzd3.frx":0E81
      Top             =   150
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   540
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "NewGzd3.frx":0EAA
      Top             =   90
      Width           =   1365
   End
   Begin MSComCtl2.DTPicker dtpA 
      Height          =   225
      Left            =   8100
      TabIndex        =   12
      Top             =   960
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   397
      _Version        =   393216
      Format          =   150011905
      CurrentDate     =   38897
   End
   Begin VB.Label LBLKjj 
      Caption         =   $"NewGzd3.frx":0EE1
      ForeColor       =   &H8000000D&
      Height          =   1305
      Left            =   12180
      TabIndex        =   270
      Top             =   360
      Width           =   2835
   End
   Begin VB.Line Line38 
      Index           =   1
      X1              =   1980
      X2              =   6135
      Y1              =   1530
      Y2              =   1530
   End
   Begin VB.Label Label39 
      Caption         =   "NO:"
      Height          =   255
      Left            =   12840
      TabIndex        =   234
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblBh 
      Caption         =   "Label29"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   13410
      TabIndex        =   233
      Top             =   120
      Width           =   1605
   End
   Begin VB.Label lblTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   10920
      TabIndex        =   184
      Top             =   10470
      Width           =   1905
   End
   Begin VB.Label lblQM 
      Caption         =   "签字提交"
      Height          =   225
      Index           =   0
      Left            =   8970
      TabIndex        =   183
      Top             =   10470
      Width           =   795
   End
   Begin VB.Label lblkhdh 
      Caption         =   "lblkhdh"
      Height          =   225
      Left            =   9960
      TabIndex        =   14
      Top             =   1290
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label lblGid 
      Caption         =   "lblGid"
      Height          =   225
      Left            =   8010
      TabIndex        =   13
      Top             =   1290
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Line Line32 
      X1              =   8100
      X2              =   12180
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line31 
      X1              =   8100
      X2              =   12180
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Line Line30 
      X1              =   8100
      X2              =   12180
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line28 
      X1              =   2010
      X2              =   6090
      Y1              =   1155
      Y2              =   1155
   End
   Begin VB.Line Line27 
      X1              =   2010
      X2              =   6090
      Y1              =   735
      Y2              =   735
   End
   Begin VB.Line Line26 
      X1              =   2010
      X2              =   6105
      Y1              =   330
      Y2              =   330
   End
End
Attribute VB_Name = "NewGzd3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoRen As ADODB.Recordset
Private Sub BA_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If Len(BA(Index).Text) >= BA(Index).Tag Then
    MsgBox ("字数超过限制,超过部分将不被保存!")
End If
End Sub

Private Sub cmdAll_Click()
Dim oo As Integer
If C1(16).Value = 1 Then
    For oo = 16 To 34
        C1(oo).Value = 0
    Next
Else
    For oo = 16 To 34
        C1(oo).Value = 1
    Next
End If
End Sub

Private Sub cmdBack_Click()
Me.Visible = False
frmGZDBR.Enabled = True
frmGZDBR.ZOrder 0
End Sub

Private Sub cmdQm_Click(Index As Integer)
Dim tt As String
Dim ii As Integer
On Error Resume Next
If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存!"
    Exit Sub
End If
'If lblkhdh.Caption = "" Then
'    MsgBox "请正确关联项目名称及相应的合同编号!"
'    Exit Sub
'End If
If cmdQm(0).Caption <> "" Then Exit Sub

ii = MsgBox("确认签字,此工作单将不能再修改,而且,它将传送至公司网站,供客户检阅,您确认此单已填写准确无误吗?", vbYesNo + vbInformation, "你好啊!")
If ii = vbYes Then
    tt = "update NewGzd set trq='" & DateSerial(Year(mod1.DQda), Month(mod1.DQda), Day(mod1.DQda)) & "' where gid=" & Val(lblGid.Caption)
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workBh, adOpenKeyset, adLockBatchOptimistic, adCmdText

cmdQm(0).Caption = mod1.DName
lblTm(0).Caption = mod1.DQda
frmGZDBR.adoY.Requery
Set frmGZDBR.dtgY.DataSource = frmGZDBR.adoY
frmGZDBR.adoW.Requery
Set frmGZDBR.dtgW.DataSource = frmGZDBR.adoW
End If
End Sub

Private Sub cmdSave_Click()
Dim tt As String
Dim oo As Integer
On Error Resume Next

tt = "select * from newgzd where gid=" & Val(lblGid.Caption)
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workBh, adOpenKeyset, adLockBatchOptimistic, adCmdText
For oo = 1 To 17
    mod1.HTP.Update "a" & oo, BA(oo).Text
Next
For oo = 1 To 27
    mod1.HTP.Update "mat" & oo, TA(oo).Text
Next
For oo = 1 To 162
    mod1.HTP.Update "mac" & oo, C1(oo).Value
Next
If FPA.Value = True Then
    mod1.HTP.Update "fp", 1
ElseIf FPB.Value = True Then
    mod1.HTP.Update "fp", 2
ElseIf FPC.Value = True Then
    mod1.HTP.Update "fp", 3
ElseIf FPD.Value = True Then
    mod1.HTP.Update "fp", 4
End If
mod1.HTP.Update "khdh", lblkhdh.Caption
    mod1.HTP.UpdateBatch
    cmdSave.Enabled = False
End Sub

Private Sub dtgRen_DblClick()
If dtgRen.Top = BA(4).Top Then
    BA(4).Text = adoRen.Fields("username").Value
ElseIf dtgRen.Top = BA(5).Top Then
    BA(5).Text = BA(5).Text & " " & adoRen.Fields("username").Value
ElseIf dtgRen.Top = BA(7).Top Then
    BA(7).Text = adoRen.Fields("username").Value
End If

End Sub

Private Sub Form_Load()
Me.Height = 11400
Me.Width = mod1.Gwidth
Me.Left = 0
Me.Top = 0

BA(1).Tag = 25
BA(2).Tag = 30
BA(3).Tag = 20
BA(4).Tag = 10
BA(5).Tag = 50
BA(7).Tag = 10
BA(12).Tag = 100
BA(13).Tag = 50
BA(14).Tag = 50
BA(15).Tag = 50
BA(16).Tag = 50
BA(17).Tag = 50
BA(8).Tag = 10
BA(9).Tag = 10
BA(10).Tag = 10
BA(11).Tag = 10
For oo = 1 To 27
    TA(oo).Tag = 50
Next
dtpA.Value = mod1.DQda
dtpB.Value = mod1.DQda
dtpC.Value = mod1.DQda
Dim tt As String
If mod1.comId = 0 Then
    tt = "select username,gzu from worker where zzf=1 and (bm='工程部' or bm='工程二部')  and qy ='" & mod1.Qy & "' order by gzu"
ElseIf mod1.comId = 1 Then
    tt = "select username,gzu from worker where zzf=1 and bm='广州工程部' order by gzu"
End If
Set adoRen = New ADODB.Recordset
adoRen.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set dtgRen.DataSource = adoRen
End Sub
Private Sub BA_DblClick(Index As Integer)
Dim tt As String
Dim oo As Integer
On Error Resume Next
If Index = 1 Then
    If BA(2).Text <> "" Then
        tt = "select 合同编号,合同金额,khdh from htView where 项目名称='" & BA(2).Text & "' and 状态='执行' and (合同性质='大修' or 合同性质='D. 维修合同' or 合同性质='C. 维保合同' or 合同性质='维保') order by 合同日期 desc "
        Set mod1.HTP = New ADODB.Recordset
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Set comHtbh.DataSource = mod1.HTP
        comHtbh.Visible = True
    End If
ElseIf Index = 2 Then
    If BA(2).Text <> "" Then

            'tt = "select xmmc from xmzl where ywy='" & mod1.DName & "' and xmmc like '%" & BA(2).Text & "%' order by xmmc"
            '为配合新客户打已经归到老客户下的工作单,特让新客户看到所有的项目资料,待打完后,此功能禁掉,采用上一去代码
            tt = "select xmmc from xmzl where  xmmc like '%" & BA(2).Text & "%' order by xmmc"
        If mod1.DName = "陈文珍" Then
                tt = "select 项目名称 as xmmc from xmview where comid=1 and 项目名称 like '%" & BA(2).Text & "%' order by xmmc"
        End If
        Set mod1.HTP = New ADODB.Recordset
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        For oo = 20 To 0 Step -1
            comXmmc.RemoveItem oo
        Next
        mod1.HTP.MoveFirst
        Do While Not mod1.HTP.EOF
            comXmmc.AddItem mod1.HTP.Fields("xmmc").Value
            mod1.HTP.MoveNext
        Loop
        comXmmc.Visible = True
    End If
ElseIf Index = 4 Then
    dtgRen.Top = BA(Index).Top
    dtgRen.Visible = True
ElseIf Index = 5 Then
    dtgRen.Top = BA(Index).Top
    dtgRen.Visible = True
ElseIf Index = 7 Then
    dtgRen.Top = BA(Index).Top
    dtgRen.Visible = True
End If

End Sub

Private Sub comHtbh_DblClick()
On Error Resume Next
BA(1).Text = mod1.HTP.Fields("合同编号").Value
lblkhdh.Caption = mod1.HTP.Fields("khdh").Value
End Sub

Private Sub comXmmc_Click()
BA(2).Text = comXmmc.Text
End Sub

Private Sub Form_Click()
comXmmc.Visible = False
comHtbh.Visible = False
dtgRen.Visible = False
End Sub
Private Sub TA_Click(Index As Integer)
comXmmc.Visible = False
comHtbh.Visible = False
dtgRen.Visible = False
End Sub
Private Sub BA_Click(Index As Integer)
dtgRen.Visible = False
comHtbh.Visible = False
comXmmc.Visible = False
End Sub

Private Sub cmdMod_Click()
cmdSave.Enabled = True
End Sub
Private Sub dtpA_CloseUp()
BA(6).Text = Format(dtpA.Value, "YYYY/MM/DD", vbUseSystemDayOfWeek)
End Sub



Private Sub dtpB_CloseUp()
BA(14).Text = Format(dtpB.Value, "YYYY/MM/DD", vbUseSystemDayOfWeek)
End Sub



Private Sub dtpC_CloseUp()
BA(16).Text = Format(dtpC.Value, "YYYY/MM/DD", vbUseSystemDayOfWeek)
End Sub
Private Sub TA_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If Len(TA(Index).Text) >= TA(Index).Tag Then
    MsgBox ("字数超过限制,超过部分将不被保存!")
End If
    If KeyCode = 13 Then SendKeys "{tab}"
    If Shift = 6 Then
        If KeyCode = 67 Then
            TA(Index).Text = TA(Index).Text & "℃"
        ElseIf KeyCode = 70 Then
            TA(Index).Text = TA(Index).Text & "℉"
        ElseIf KeyCode = 80 Then
            TA(Index).Text = TA(Index).Text & "psi"
        ElseIf KeyCode = 75 Then
            TA(Index).Text = TA(Index).Text & "kpa"
        ElseIf KeyCode = 71 Then
            TA(Index).Text = TA(Index).Text & "kg/cm2"
        ElseIf KeyCode = 85 Then
            TA(Index).Text = TA(Index).Text & "μf"
        ElseIf KeyCode = 79 Then
            TA(Index).Text = TA(Index).Text & "Ω"
        End If
        TA(Index).SelStart = Len(TA(Index).Text)
        TA(Index).SelLength = 1
    End If
End Sub
