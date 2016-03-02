VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form NewGzd5 
   Caption         =   "应急维修工作报告（单）"
   ClientHeight    =   10830
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15015
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
      Left            =   1920
      TabIndex        =   99
      Top             =   1320
      Width           =   4065
   End
   Begin MSDataGridLib.DataGrid dtgRen 
      Height          =   8085
      Left            =   10290
      TabIndex        =   96
      Top             =   90
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
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   14040
      Picture         =   "NewGzd5.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "保存"
      Top             =   10380
      Width           =   465
   End
   Begin TabDlg.SSTab tabNr 
      Height          =   8745
      Left            =   0
      TabIndex        =   22
      Top             =   1620
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   15425
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "内容1"
      TabPicture(0)   =   "NewGzd5.frx":066A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label8"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Shape1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line40(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line40(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line40(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Line40(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line39(3)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Line39(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Line39(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Line39(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Shape2(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Line38(0)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Line37(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label23(2)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label23(1)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label24"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label23(0)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Line36(0)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Line35(0)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Line34(0)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Line33(0)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Line25(0)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Line24(0)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Line23"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Line22"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Line21"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Line20"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Line19"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Line18"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Line17"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Line16"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Line15"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Line14"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Line13"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Line12"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Line11"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Line10"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Line9"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Line8"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Line7"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Line6"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Line5"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Line4"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Line3"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Line2"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "Label22"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Label21"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Label20"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Label19"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Label18"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "Label17"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "Label16"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "Label15"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "Label14"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "Label13"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "Label9"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "Label11"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "Label12"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "Label10"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "Label6"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "Label1"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "Line1"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "TA(103)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "TA(99)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "TA(95)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "TA(91)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "TA(87)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "TA(106)"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "TA(105)"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "TA(104)"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "TA(102)"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "TA(101)"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "TA(100)"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "TA(98)"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "TA(97)"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "TA(96)"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "TA(94)"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "TA(93)"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "TA(92)"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).Control(84)=   "TA(90)"
      Tab(0).Control(84).Enabled=   0   'False
      Tab(0).Control(85)=   "TA(89)"
      Tab(0).Control(85).Enabled=   0   'False
      Tab(0).Control(86)=   "TA(88)"
      Tab(0).Control(86).Enabled=   0   'False
      Tab(0).Control(87)=   "TA(86)"
      Tab(0).Control(87).Enabled=   0   'False
      Tab(0).Control(88)=   "TA(85)"
      Tab(0).Control(88).Enabled=   0   'False
      Tab(0).Control(89)=   "TA(84)"
      Tab(0).Control(89).Enabled=   0   'False
      Tab(0).Control(90)=   "TA(83)"
      Tab(0).Control(90).Enabled=   0   'False
      Tab(0).Control(91)=   "TA(82)"
      Tab(0).Control(91).Enabled=   0   'False
      Tab(0).Control(92)=   "TA(81)"
      Tab(0).Control(92).Enabled=   0   'False
      Tab(0).Control(93)=   "C1(20)"
      Tab(0).Control(93).Enabled=   0   'False
      Tab(0).Control(94)=   "C1(19)"
      Tab(0).Control(94).Enabled=   0   'False
      Tab(0).Control(95)=   "C1(18)"
      Tab(0).Control(95).Enabled=   0   'False
      Tab(0).Control(96)=   "C1(17)"
      Tab(0).Control(96).Enabled=   0   'False
      Tab(0).Control(97)=   "C1(16)"
      Tab(0).Control(97).Enabled=   0   'False
      Tab(0).Control(98)=   "C1(15)"
      Tab(0).Control(98).Enabled=   0   'False
      Tab(0).Control(99)=   "C1(14)"
      Tab(0).Control(99).Enabled=   0   'False
      Tab(0).Control(100)=   "C1(13)"
      Tab(0).Control(100).Enabled=   0   'False
      Tab(0).Control(101)=   "C1(12)"
      Tab(0).Control(101).Enabled=   0   'False
      Tab(0).Control(102)=   "C1(11)"
      Tab(0).Control(102).Enabled=   0   'False
      Tab(0).Control(103)=   "C1(10)"
      Tab(0).Control(103).Enabled=   0   'False
      Tab(0).Control(104)=   "C1(9)"
      Tab(0).Control(104).Enabled=   0   'False
      Tab(0).Control(105)=   "TA(80)"
      Tab(0).Control(105).Enabled=   0   'False
      Tab(0).Control(106)=   "TA(75)"
      Tab(0).Control(106).Enabled=   0   'False
      Tab(0).Control(107)=   "TA(70)"
      Tab(0).Control(107).Enabled=   0   'False
      Tab(0).Control(108)=   "TA(65)"
      Tab(0).Control(108).Enabled=   0   'False
      Tab(0).Control(109)=   "TA(60)"
      Tab(0).Control(109).Enabled=   0   'False
      Tab(0).Control(110)=   "TA(55)"
      Tab(0).Control(110).Enabled=   0   'False
      Tab(0).Control(111)=   "TA(50)"
      Tab(0).Control(111).Enabled=   0   'False
      Tab(0).Control(112)=   "TA(45)"
      Tab(0).Control(112).Enabled=   0   'False
      Tab(0).Control(113)=   "TA(40)"
      Tab(0).Control(113).Enabled=   0   'False
      Tab(0).Control(114)=   "TA(35)"
      Tab(0).Control(114).Enabled=   0   'False
      Tab(0).Control(115)=   "TA(30)"
      Tab(0).Control(115).Enabled=   0   'False
      Tab(0).Control(116)=   "TA(25)"
      Tab(0).Control(116).Enabled=   0   'False
      Tab(0).Control(117)=   "TA(78)"
      Tab(0).Control(117).Enabled=   0   'False
      Tab(0).Control(118)=   "TA(73)"
      Tab(0).Control(118).Enabled=   0   'False
      Tab(0).Control(119)=   "TA(68)"
      Tab(0).Control(119).Enabled=   0   'False
      Tab(0).Control(120)=   "TA(63)"
      Tab(0).Control(120).Enabled=   0   'False
      Tab(0).Control(121)=   "TA(58)"
      Tab(0).Control(121).Enabled=   0   'False
      Tab(0).Control(122)=   "TA(53)"
      Tab(0).Control(122).Enabled=   0   'False
      Tab(0).Control(123)=   "TA(48)"
      Tab(0).Control(123).Enabled=   0   'False
      Tab(0).Control(124)=   "TA(43)"
      Tab(0).Control(124).Enabled=   0   'False
      Tab(0).Control(125)=   "TA(38)"
      Tab(0).Control(125).Enabled=   0   'False
      Tab(0).Control(126)=   "TA(33)"
      Tab(0).Control(126).Enabled=   0   'False
      Tab(0).Control(127)=   "TA(28)"
      Tab(0).Control(127).Enabled=   0   'False
      Tab(0).Control(128)=   "TA(23)"
      Tab(0).Control(128).Enabled=   0   'False
      Tab(0).Control(129)=   "C1(7)"
      Tab(0).Control(129).Enabled=   0   'False
      Tab(0).Control(130)=   "TA(79)"
      Tab(0).Control(130).Enabled=   0   'False
      Tab(0).Control(131)=   "TA(77)"
      Tab(0).Control(131).Enabled=   0   'False
      Tab(0).Control(132)=   "TA(76)"
      Tab(0).Control(132).Enabled=   0   'False
      Tab(0).Control(133)=   "TA(74)"
      Tab(0).Control(133).Enabled=   0   'False
      Tab(0).Control(134)=   "TA(72)"
      Tab(0).Control(134).Enabled=   0   'False
      Tab(0).Control(135)=   "TA(71)"
      Tab(0).Control(135).Enabled=   0   'False
      Tab(0).Control(136)=   "TA(69)"
      Tab(0).Control(136).Enabled=   0   'False
      Tab(0).Control(137)=   "TA(67)"
      Tab(0).Control(137).Enabled=   0   'False
      Tab(0).Control(138)=   "TA(66)"
      Tab(0).Control(138).Enabled=   0   'False
      Tab(0).Control(139)=   "TA(64)"
      Tab(0).Control(139).Enabled=   0   'False
      Tab(0).Control(140)=   "TA(62)"
      Tab(0).Control(140).Enabled=   0   'False
      Tab(0).Control(141)=   "TA(61)"
      Tab(0).Control(141).Enabled=   0   'False
      Tab(0).Control(142)=   "TA(59)"
      Tab(0).Control(142).Enabled=   0   'False
      Tab(0).Control(143)=   "TA(57)"
      Tab(0).Control(143).Enabled=   0   'False
      Tab(0).Control(144)=   "TA(56)"
      Tab(0).Control(144).Enabled=   0   'False
      Tab(0).Control(145)=   "TA(54)"
      Tab(0).Control(145).Enabled=   0   'False
      Tab(0).Control(146)=   "TA(52)"
      Tab(0).Control(146).Enabled=   0   'False
      Tab(0).Control(147)=   "TA(51)"
      Tab(0).Control(147).Enabled=   0   'False
      Tab(0).Control(148)=   "TA(49)"
      Tab(0).Control(148).Enabled=   0   'False
      Tab(0).Control(149)=   "TA(47)"
      Tab(0).Control(149).Enabled=   0   'False
      Tab(0).Control(150)=   "TA(46)"
      Tab(0).Control(150).Enabled=   0   'False
      Tab(0).Control(151)=   "TA(44)"
      Tab(0).Control(151).Enabled=   0   'False
      Tab(0).Control(152)=   "TA(42)"
      Tab(0).Control(152).Enabled=   0   'False
      Tab(0).Control(153)=   "TA(41)"
      Tab(0).Control(153).Enabled=   0   'False
      Tab(0).Control(154)=   "TA(39)"
      Tab(0).Control(154).Enabled=   0   'False
      Tab(0).Control(155)=   "TA(37)"
      Tab(0).Control(155).Enabled=   0   'False
      Tab(0).Control(156)=   "TA(36)"
      Tab(0).Control(156).Enabled=   0   'False
      Tab(0).Control(157)=   "TA(34)"
      Tab(0).Control(157).Enabled=   0   'False
      Tab(0).Control(158)=   "TA(32)"
      Tab(0).Control(158).Enabled=   0   'False
      Tab(0).Control(159)=   "TA(31)"
      Tab(0).Control(159).Enabled=   0   'False
      Tab(0).Control(160)=   "TA(29)"
      Tab(0).Control(160).Enabled=   0   'False
      Tab(0).Control(161)=   "TA(27)"
      Tab(0).Control(161).Enabled=   0   'False
      Tab(0).Control(162)=   "TA(26)"
      Tab(0).Control(162).Enabled=   0   'False
      Tab(0).Control(163)=   "TA(24)"
      Tab(0).Control(163).Enabled=   0   'False
      Tab(0).Control(164)=   "TA(22)"
      Tab(0).Control(164).Enabled=   0   'False
      Tab(0).Control(165)=   "TA(21)"
      Tab(0).Control(165).Enabled=   0   'False
      Tab(0).Control(166)=   "TA(20)"
      Tab(0).Control(166).Enabled=   0   'False
      Tab(0).Control(167)=   "TA(19)"
      Tab(0).Control(167).Enabled=   0   'False
      Tab(0).Control(168)=   "TA(1)"
      Tab(0).Control(168).Enabled=   0   'False
      Tab(0).Control(169)=   "TA(2)"
      Tab(0).Control(169).Enabled=   0   'False
      Tab(0).Control(170)=   "TA(7)"
      Tab(0).Control(170).Enabled=   0   'False
      Tab(0).Control(171)=   "TA(8)"
      Tab(0).Control(171).Enabled=   0   'False
      Tab(0).Control(172)=   "TA(9)"
      Tab(0).Control(172).Enabled=   0   'False
      Tab(0).Control(173)=   "TA(10)"
      Tab(0).Control(173).Enabled=   0   'False
      Tab(0).Control(174)=   "C1(8)"
      Tab(0).Control(174).Enabled=   0   'False
      Tab(0).Control(175)=   "C1(6)"
      Tab(0).Control(175).Enabled=   0   'False
      Tab(0).Control(176)=   "C1(5)"
      Tab(0).Control(176).Enabled=   0   'False
      Tab(0).Control(177)=   "C1(4)"
      Tab(0).Control(177).Enabled=   0   'False
      Tab(0).Control(178)=   "C1(3)"
      Tab(0).Control(178).Enabled=   0   'False
      Tab(0).Control(179)=   "C1(2)"
      Tab(0).Control(179).Enabled=   0   'False
      Tab(0).Control(180)=   "C1(1)"
      Tab(0).Control(180).Enabled=   0   'False
      Tab(0).Control(181)=   "TA(18)"
      Tab(0).Control(181).Enabled=   0   'False
      Tab(0).Control(182)=   "TA(17)"
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
      Tab(0).Control(189)=   "TA(6)"
      Tab(0).Control(189).Enabled=   0   'False
      Tab(0).Control(190)=   "TA(5)"
      Tab(0).Control(190).Enabled=   0   'False
      Tab(0).Control(191)=   "TA(4)"
      Tab(0).Control(191).Enabled=   0   'False
      Tab(0).Control(192)=   "TA(3)"
      Tab(0).Control(192).Enabled=   0   'False
      Tab(0).Control(193)=   "cmdAll"
      Tab(0).Control(193).Enabled=   0   'False
      Tab(0).ControlCount=   194
      TabCaption(1)   =   "内容2"
      TabPicture(1)   =   "NewGzd5.frx":0686
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label26"
      Tab(1).Control(1)=   "Label27"
      Tab(1).Control(2)=   "Label28"
      Tab(1).Control(3)=   "Line24(1)"
      Tab(1).Control(4)=   "Line25(1)"
      Tab(1).Control(5)=   "Line33(1)"
      Tab(1).Control(6)=   "Label29"
      Tab(1).Control(7)=   "Label30"
      Tab(1).Control(8)=   "Label31"
      Tab(1).Control(9)=   "Label32"
      Tab(1).Control(10)=   "Label34"
      Tab(1).Control(11)=   "Shape2(1)"
      Tab(1).Control(12)=   "Label35"
      Tab(1).Control(13)=   "Label36"
      Tab(1).Control(14)=   "Label37"
      Tab(1).Control(15)=   "Line34(1)"
      Tab(1).Control(16)=   "Line35(1)"
      Tab(1).Control(17)=   "Line36(1)"
      Tab(1).Control(18)=   "Line37(1)"
      Tab(1).Control(19)=   "Label38"
      Tab(1).Control(20)=   "TA(107)"
      Tab(1).Control(21)=   "TA(109)"
      Tab(1).Control(22)=   "dtpC"
      Tab(1).Control(23)=   "dtpB"
      Tab(1).Control(24)=   "TA(108)"
      Tab(1).Control(25)=   "C1(21)"
      Tab(1).Control(26)=   "C1(22)"
      Tab(1).Control(27)=   "Text3"
      Tab(1).Control(28)=   "TA(110)"
      Tab(1).Control(29)=   "BA(8)"
      Tab(1).Control(30)=   "BA(9)"
      Tab(1).Control(31)=   "BA(10)"
      Tab(1).Control(32)=   "BA(11)"
      Tab(1).Control(33)=   "Frame1"
      Tab(1).Control(34)=   "BA(12)"
      Tab(1).Control(35)=   "BA(13)"
      Tab(1).Control(36)=   "BA(14)"
      Tab(1).Control(37)=   "BA(15)"
      Tab(1).Control(38)=   "BA(16)"
      Tab(1).ControlCount=   39
      Begin VB.CommandButton cmdAll 
         Caption         =   "全部"
         Height          =   225
         Left            =   5880
         TabIndex        =   216
         Top             =   6690
         Width           =   615
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   3
         Left            =   1800
         TabIndex        =   101
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   4
         Left            =   5010
         TabIndex        =   104
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   5
         Left            =   1800
         TabIndex        =   102
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   6
         Left            =   5010
         TabIndex        =   105
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   11
         Left            =   8310
         TabIndex        =   107
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   12
         Left            =   9930
         TabIndex        =   110
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   13
         Left            =   11520
         TabIndex        =   113
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   14
         Left            =   13200
         TabIndex        =   116
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   15
         Left            =   8310
         TabIndex        =   108
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   16
         Left            =   9930
         TabIndex        =   111
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   17
         Left            =   11520
         TabIndex        =   114
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   18
         Left            =   13200
         TabIndex        =   117
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "1#"
         Height          =   180
         Index           =   1
         Left            =   8760
         TabIndex        =   69
         Top             =   330
         Width           =   495
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "2#"
         Height          =   180
         Index           =   2
         Left            =   10320
         TabIndex        =   68
         Top             =   330
         Width           =   495
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "3#"
         Height          =   180
         Index           =   3
         Left            =   11940
         TabIndex        =   67
         Top             =   330
         Width           =   495
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "4#"
         Height          =   180
         Index           =   4
         Left            =   13380
         TabIndex        =   66
         Top             =   330
         Width           =   495
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "1#"
         Height          =   180
         Index           =   5
         Left            =   1830
         TabIndex        =   65
         Top             =   1680
         Width           =   495
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "2#"
         Height          =   180
         Index           =   6
         Left            =   2610
         TabIndex        =   64
         Top             =   1680
         Width           =   495
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "4#"
         Height          =   180
         Index           =   8
         Left            =   4185
         TabIndex        =   63
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   10
         Left            =   13200
         TabIndex        =   115
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   9
         Left            =   11520
         TabIndex        =   112
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   8
         Left            =   9930
         TabIndex        =   109
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   7
         Left            =   8310
         TabIndex        =   106
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   2
         Left            =   5010
         TabIndex        =   103
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   1
         Left            =   1800
         TabIndex        =   100
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   1020
         Index           =   19
         Left            =   6660
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   178
         Top             =   1950
         Width           =   8265
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   1050
         Index           =   20
         Left            =   6660
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   179
         Top             =   3390
         Width           =   8265
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   21
         Left            =   1830
         TabIndex        =   118
         Top             =   2040
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   22
         Left            =   2610
         TabIndex        =   130
         Top             =   2040
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   24
         Left            =   4185
         TabIndex        =   154
         Top             =   2040
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   26
         Left            =   1830
         TabIndex        =   119
         Top             =   2400
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   27
         Left            =   2610
         TabIndex        =   131
         Top             =   2400
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   29
         Left            =   4185
         TabIndex        =   155
         Top             =   2400
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   31
         Left            =   1830
         TabIndex        =   120
         Top             =   2760
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   32
         Left            =   2610
         TabIndex        =   132
         Top             =   2760
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   34
         Left            =   4185
         TabIndex        =   156
         Top             =   2760
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   36
         Left            =   1830
         TabIndex        =   121
         Top             =   3120
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   37
         Left            =   2610
         TabIndex        =   133
         Top             =   3120
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   39
         Left            =   4185
         TabIndex        =   157
         Top             =   3120
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   41
         Left            =   1830
         TabIndex        =   122
         Top             =   3480
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   42
         Left            =   2610
         TabIndex        =   134
         Top             =   3480
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   44
         Left            =   4185
         TabIndex        =   158
         Top             =   3480
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   46
         Left            =   1830
         TabIndex        =   123
         Top             =   3840
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   47
         Left            =   2610
         TabIndex        =   135
         Top             =   3840
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   49
         Left            =   4185
         TabIndex        =   159
         Top             =   3840
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   51
         Left            =   1830
         TabIndex        =   124
         Top             =   4200
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   52
         Left            =   2610
         TabIndex        =   136
         Top             =   4200
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   54
         Left            =   4185
         TabIndex        =   160
         Top             =   4200
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   56
         Left            =   1830
         TabIndex        =   125
         Top             =   4560
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   57
         Left            =   2610
         TabIndex        =   137
         Top             =   4560
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   59
         Left            =   4185
         TabIndex        =   161
         Top             =   4560
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   61
         Left            =   1830
         TabIndex        =   126
         Top             =   4920
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   62
         Left            =   2610
         TabIndex        =   138
         Top             =   4920
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   64
         Left            =   4185
         TabIndex        =   162
         Top             =   4920
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   66
         Left            =   1830
         TabIndex        =   127
         Top             =   5280
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   67
         Left            =   2610
         TabIndex        =   139
         Top             =   5280
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   69
         Left            =   4185
         TabIndex        =   163
         Top             =   5280
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   71
         Left            =   1830
         TabIndex        =   128
         Top             =   5640
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   72
         Left            =   2610
         TabIndex        =   140
         Top             =   5640
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   74
         Left            =   4185
         TabIndex        =   164
         Top             =   5640
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   76
         Left            =   1830
         TabIndex        =   129
         Top             =   6000
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   77
         Left            =   2610
         TabIndex        =   141
         Top             =   6000
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   79
         Left            =   4185
         TabIndex        =   165
         Top             =   6000
         Width           =   585
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "3#"
         Height          =   180
         Index           =   7
         Left            =   3405
         TabIndex        =   62
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   23
         Left            =   3405
         TabIndex        =   142
         Top             =   2040
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   28
         Left            =   3405
         TabIndex        =   143
         Top             =   2400
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   33
         Left            =   3405
         TabIndex        =   144
         Top             =   2760
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   38
         Left            =   3405
         TabIndex        =   145
         Top             =   3120
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   43
         Left            =   3405
         TabIndex        =   146
         Top             =   3480
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   48
         Left            =   3405
         TabIndex        =   147
         Top             =   3840
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   53
         Left            =   3405
         TabIndex        =   148
         Top             =   4200
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   58
         Left            =   3405
         TabIndex        =   149
         Top             =   4560
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   63
         Left            =   3405
         TabIndex        =   150
         Top             =   4920
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   68
         Left            =   3405
         TabIndex        =   151
         Top             =   5280
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   73
         Left            =   3405
         TabIndex        =   152
         Top             =   5640
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   78
         Left            =   3405
         TabIndex        =   153
         Top             =   6000
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   25
         Left            =   4980
         TabIndex        =   166
         Top             =   2025
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   30
         Left            =   4980
         TabIndex        =   167
         Top             =   2386
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   35
         Left            =   4980
         TabIndex        =   168
         Top             =   2747
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   40
         Left            =   4980
         TabIndex        =   169
         Top             =   3108
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   45
         Left            =   4980
         TabIndex        =   170
         Top             =   3469
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   50
         Left            =   4980
         TabIndex        =   171
         Top             =   3830
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   55
         Left            =   4980
         TabIndex        =   172
         Top             =   4191
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   60
         Left            =   4980
         TabIndex        =   173
         Top             =   4552
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   65
         Left            =   4980
         TabIndex        =   174
         Top             =   4913
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   70
         Left            =   4980
         TabIndex        =   175
         Top             =   5274
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   75
         Left            =   4980
         TabIndex        =   176
         Top             =   5635
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   80
         Left            =   4980
         TabIndex        =   177
         Top             =   6000
         Width           =   585
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   9
         Left            =   5970
         TabIndex        =   61
         Top             =   2025
         Width           =   285
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   10
         Left            =   5970
         TabIndex        =   60
         Top             =   2386
         Width           =   285
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   11
         Left            =   5970
         TabIndex        =   59
         Top             =   2747
         Width           =   285
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   12
         Left            =   5970
         TabIndex        =   58
         Top             =   3108
         Width           =   285
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   13
         Left            =   5970
         TabIndex        =   57
         Top             =   3469
         Width           =   285
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   14
         Left            =   5970
         TabIndex        =   56
         Top             =   3830
         Width           =   285
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   15
         Left            =   5970
         TabIndex        =   55
         Top             =   4191
         Width           =   285
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   16
         Left            =   5970
         TabIndex        =   54
         Top             =   4552
         Width           =   285
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   17
         Left            =   5970
         TabIndex        =   53
         Top             =   4913
         Width           =   285
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   18
         Left            =   5970
         TabIndex        =   52
         Top             =   5274
         Width           =   285
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   19
         Left            =   5970
         TabIndex        =   51
         Top             =   5635
         Width           =   285
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   20
         Left            =   5970
         TabIndex        =   50
         Top             =   6000
         Width           =   285
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   81
         Left            =   7920
         TabIndex        =   180
         Top             =   4573
         Width           =   2955
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   82
         Left            =   11520
         TabIndex        =   181
         Top             =   4573
         Width           =   3405
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   83
         Left            =   7920
         TabIndex        =   182
         Top             =   4929
         Width           =   2955
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   84
         Left            =   11520
         TabIndex        =   183
         Top             =   4929
         Width           =   3405
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   570
         Index           =   85
         Left            =   7530
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   184
         Top             =   5285
         Width           =   7395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   86
         Left            =   7530
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   185
         Top             =   6000
         Width           =   7395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   88
         Left            =   270
         TabIndex        =   191
         Top             =   7635
         Width           =   1455
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   89
         Left            =   270
         TabIndex        =   196
         Top             =   7905
         Width           =   1455
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   90
         Left            =   270
         TabIndex        =   201
         Top             =   8190
         Width           =   1455
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   92
         Left            =   2130
         TabIndex        =   192
         Top             =   7635
         Width           =   2865
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   93
         Left            =   2130
         TabIndex        =   197
         Top             =   7905
         Width           =   2865
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   94
         Left            =   2130
         TabIndex        =   202
         Top             =   8190
         Width           =   2865
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   96
         Left            =   5460
         TabIndex        =   193
         Top             =   7635
         Width           =   2415
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   97
         Left            =   5460
         TabIndex        =   198
         Top             =   7905
         Width           =   2415
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   98
         Left            =   5460
         TabIndex        =   203
         Top             =   8190
         Width           =   2415
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   100
         Left            =   8070
         TabIndex        =   194
         Top             =   7635
         Width           =   5055
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   101
         Left            =   8070
         TabIndex        =   199
         Top             =   7905
         Width           =   5055
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   102
         Left            =   8070
         TabIndex        =   204
         Top             =   8190
         Width           =   5055
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   104
         Left            =   13290
         TabIndex        =   195
         Top             =   7635
         Width           =   1455
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   105
         Left            =   13290
         TabIndex        =   200
         Top             =   7905
         Width           =   1455
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   106
         Left            =   13290
         TabIndex        =   205
         Top             =   8190
         Width           =   1455
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   87
         Left            =   270
         TabIndex        =   186
         Top             =   7380
         Width           =   1455
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   91
         Left            =   2130
         TabIndex        =   187
         Top             =   7380
         Width           =   2865
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   95
         Left            =   5460
         TabIndex        =   188
         Top             =   7380
         Width           =   2415
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   99
         Left            =   8070
         TabIndex        =   189
         Top             =   7380
         Width           =   5055
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   103
         Left            =   13290
         TabIndex        =   190
         Top             =   7380
         Width           =   1455
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
         Left            =   -62070
         TabIndex        =   35
         Top             =   2820
         Width           =   1725
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
         Left            =   -62040
         TabIndex        =   215
         Top             =   2250
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
         Index           =   14
         Left            =   -64110
         TabIndex        =   34
         Top             =   2820
         Width           =   1755
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
         Left            =   -64110
         TabIndex        =   214
         Top             =   2250
         Width           =   1935
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   570
         Index           =   12
         Left            =   -73560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   213
         Text            =   "NewGzd5.frx":06A2
         Top             =   2580
         Width           =   9345
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Left            =   -74970
         TabIndex        =   28
         Top             =   2250
         Width           =   10755
         Begin VB.OptionButton FPA 
            Caption         =   "优秀"
            Height          =   195
            Left            =   1350
            TabIndex        =   32
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton FPB 
            Caption         =   "满意"
            Height          =   195
            Left            =   2950
            TabIndex        =   31
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton FPC 
            Caption         =   "较满意"
            Height          =   195
            Left            =   4550
            TabIndex        =   30
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton FPD 
            Caption         =   "尚可"
            Height          =   195
            Left            =   6150
            TabIndex        =   29
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
            TabIndex        =   33
            Top             =   60
            Width           =   1395
         End
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   11
         Left            =   -65250
         TabIndex        =   212
         Text            =   "的"
         Top             =   2010
         Width           =   1035
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   10
         Left            =   -67650
         TabIndex        =   211
         Text            =   "的"
         Top             =   2010
         Width           =   1035
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   9
         Left            =   -70320
         TabIndex        =   210
         Text            =   "的"
         Top             =   2010
         Width           =   1035
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   8
         Left            =   -73560
         TabIndex        =   209
         Text            =   "的"
         Top             =   2010
         Width           =   1035
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   110
         Left            =   -61260
         TabIndex        =   27
         Top             =   1680
         Width           =   1065
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   -61890
         TabIndex        =   26
         Text            =   "复核人:"
         Top             =   1770
         Width           =   735
      End
      Begin VB.CheckBox C1 
         Caption         =   "未完成"
         Height          =   180
         Index           =   22
         Left            =   -61140
         TabIndex        =   25
         Top             =   600
         Width           =   945
      End
      Begin VB.CheckBox C1 
         Caption         =   "完成"
         Height          =   180
         Index           =   21
         Left            =   -62280
         TabIndex        =   24
         Top             =   600
         Width           =   1005
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   570
         Index           =   108
         Left            =   -73590
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   207
         Top             =   810
         Width           =   13545
      End
      Begin MSComCtl2.DTPicker dtpB 
         Height          =   225
         Left            =   -64110
         TabIndex        =   36
         Top             =   2820
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   397
         _Version        =   393216
         Format          =   149618689
         CurrentDate     =   38897
      End
      Begin MSComCtl2.DTPicker dtpC 
         Height          =   225
         Left            =   -62100
         TabIndex        =   37
         Top             =   2820
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   397
         _Version        =   393216
         Format          =   149618689
         CurrentDate     =   38897
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   510
         Index           =   109
         Left            =   -73590
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   208
         Top             =   1440
         Width           =   13545
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   107
         Left            =   -73590
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   206
         Top             =   60
         Width           =   13545
      End
      Begin VB.Line Line1 
         X1              =   60
         X2              =   15000
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label1 
         Caption         =   "应急维修服务内容与机组参数记录与（记录与压缩机对应的数据时，在压缩机编号一栏中相应编号的""□""上打""√""，若无此压缩机则打""／""）"
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
         Left            =   210
         TabIndex        =   95
         Top             =   60
         Width           =   14295
      End
      Begin VB.Label Label6 
         Caption         =   "压缩机型号"
         Height          =   195
         Left            =   6660
         TabIndex        =   89
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label Label10 
         Caption         =   "情况分析与故障判断："
         Height          =   225
         Left            =   6660
         TabIndex        =   87
         Top             =   3120
         Width           =   2115
      End
      Begin VB.Label Label12 
         Caption         =   "无此项"
         Height          =   225
         Left            =   5880
         TabIndex        =   86
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label Label11 
         Caption         =   "正常值"
         Height          =   225
         Left            =   4980
         TabIndex        =   85
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "故障现象及现场情况："
         Height          =   225
         Left            =   6660
         TabIndex        =   84
         Top             =   1680
         Width           =   2025
      End
      Begin VB.Label Label13 
         Caption         =   "故障"
         Height          =   195
         Left            =   6660
         TabIndex        =   83
         Top             =   4575
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "代码"
         Height          =   195
         Left            =   6660
         TabIndex        =   82
         Top             =   4935
         Width           =   675
      End
      Begin VB.Label Label15 
         Caption         =   "处理及结果："
         Height          =   405
         Left            =   6660
         TabIndex        =   81
         Top             =   5280
         Width           =   705
      End
      Begin VB.Label Label16 
         Caption         =   "备注："
         Height          =   225
         Left            =   6660
         TabIndex        =   80
         Top             =   6000
         Width           =   705
      End
      Begin VB.Label Label17 
         Caption         =   "在维修过程中的材料与零配件情况"
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
         Left            =   300
         TabIndex        =   79
         Top             =   6660
         Width           =   3555
      End
      Begin VB.Label Label18 
         Caption         =   "数量"
         Height          =   255
         Left            =   420
         TabIndex        =   78
         Top             =   7020
         Width           =   1065
      End
      Begin VB.Label Label19 
         Caption         =   "零配件或材料名称"
         Height          =   225
         Left            =   2280
         TabIndex        =   77
         Top             =   7020
         Width           =   2805
      End
      Begin VB.Label Label20 
         Caption         =   "零件编号"
         Height          =   255
         Left            =   5820
         TabIndex        =   76
         Top             =   7020
         Width           =   1905
      End
      Begin VB.Label Label21 
         Caption         =   "使用情况"
         Height          =   255
         Left            =   8640
         TabIndex        =   75
         Top             =   7020
         Width           =   2145
      End
      Begin VB.Label Label22 
         Caption         =   "供货方"
         Height          =   255
         Left            =   13620
         TabIndex        =   74
         Top             =   7020
         Width           =   1065
      End
      Begin VB.Line Line2 
         X1              =   60
         X2              =   14940
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Line Line3 
         X1              =   60
         X2              =   14940
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line4 
         X1              =   60
         X2              =   14955
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line5 
         X1              =   60
         X2              =   14940
         Y1              =   1530
         Y2              =   1530
      End
      Begin VB.Line Line6 
         X1              =   6540
         X2              =   6540
         Y1              =   270
         Y2              =   6330
      End
      Begin VB.Line Line7 
         X1              =   60
         X2              =   6555
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line8 
         X1              =   60
         X2              =   6540
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line9 
         X1              =   60
         X2              =   6555
         Y1              =   2610
         Y2              =   2610
      End
      Begin VB.Line Line10 
         X1              =   60
         X2              =   14940
         Y1              =   2970
         Y2              =   2970
      End
      Begin VB.Line Line11 
         X1              =   60
         X2              =   6540
         Y1              =   3330
         Y2              =   3330
      End
      Begin VB.Line Line12 
         X1              =   60
         X2              =   6540
         Y1              =   3690
         Y2              =   3690
      End
      Begin VB.Line Line13 
         X1              =   60
         X2              =   6555
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line14 
         X1              =   60
         X2              =   14940
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line15 
         X1              =   60
         X2              =   6555
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Line Line16 
         X1              =   60
         X2              =   14940
         Y1              =   5190
         Y2              =   5190
      End
      Begin VB.Line Line17 
         X1              =   60
         X2              =   6555
         Y1              =   5550
         Y2              =   5550
      End
      Begin VB.Line Line18 
         X1              =   60
         X2              =   14940
         Y1              =   5910
         Y2              =   5910
      End
      Begin VB.Line Line19 
         X1              =   1620
         X2              =   1620
         Y1              =   300
         Y2              =   6360
      End
      Begin VB.Line Line20 
         X1              =   4890
         X2              =   4890
         Y1              =   300
         Y2              =   6360
      End
      Begin VB.Line Line21 
         X1              =   3330
         X2              =   3330
         Y1              =   300
         Y2              =   6360
      End
      Begin VB.Line Line22 
         X1              =   2490
         X2              =   2490
         Y1              =   1560
         Y2              =   6360
      End
      Begin VB.Line Line23 
         X1              =   4080
         X2              =   4080
         Y1              =   1560
         Y2              =   6360
      End
      Begin VB.Line Line24 
         Index           =   0
         X1              =   5760
         X2              =   5760
         Y1              =   1560
         Y2              =   6360
      End
      Begin VB.Line Line25 
         Index           =   0
         X1              =   8130
         X2              =   8130
         Y1              =   270
         Y2              =   1530
      End
      Begin VB.Line Line33 
         Index           =   0
         X1              =   9840
         X2              =   9840
         Y1              =   270
         Y2              =   1530
      End
      Begin VB.Line Line34 
         Index           =   0
         X1              =   11430
         X2              =   11430
         Y1              =   270
         Y2              =   1530
      End
      Begin VB.Line Line35 
         Index           =   0
         X1              =   13140
         X2              =   13140
         Y1              =   270
         Y2              =   1530
      End
      Begin VB.Line Line36 
         Index           =   0
         X1              =   7410
         X2              =   7410
         Y1              =   5190
         Y2              =   4440
      End
      Begin VB.Label Label23 
         Caption         =   "1#"
         Height          =   225
         Index           =   0
         Left            =   7560
         TabIndex        =   73
         Top             =   4560
         Width           =   285
      End
      Begin VB.Label Label24 
         Caption         =   "2#"
         Height          =   225
         Left            =   11100
         TabIndex        =   72
         Top             =   4560
         Width           =   285
      End
      Begin VB.Label Label23 
         Caption         =   "3#"
         Height          =   225
         Index           =   1
         Left            =   7560
         TabIndex        =   71
         Top             =   4935
         Width           =   285
      End
      Begin VB.Label Label23 
         Caption         =   "4#"
         Height          =   225
         Index           =   2
         Left            =   11100
         TabIndex        =   70
         Top             =   4935
         Width           =   285
      End
      Begin VB.Line Line37 
         Index           =   0
         X1              =   7410
         X2              =   14940
         Y1              =   4860
         Y2              =   4860
      End
      Begin VB.Line Line38 
         Index           =   0
         X1              =   10980
         X2              =   10980
         Y1              =   5190
         Y2              =   4440
      End
      Begin VB.Shape Shape2 
         Height          =   1455
         Index           =   0
         Left            =   60
         Top             =   6960
         Width           =   14895
      End
      Begin VB.Line Line39 
         Index           =   0
         X1              =   60
         X2              =   14940
         Y1              =   7290
         Y2              =   7290
      End
      Begin VB.Line Line39 
         Index           =   1
         X1              =   60
         X2              =   14940
         Y1              =   7590
         Y2              =   7590
      End
      Begin VB.Line Line39 
         Index           =   2
         X1              =   60
         X2              =   14940
         Y1              =   7860
         Y2              =   7860
      End
      Begin VB.Line Line39 
         Index           =   3
         X1              =   60
         X2              =   14940
         Y1              =   8160
         Y2              =   8160
      End
      Begin VB.Line Line40 
         Index           =   0
         X1              =   1890
         X2              =   1890
         Y1              =   8385
         Y2              =   6960
      End
      Begin VB.Line Line40 
         Index           =   1
         X1              =   5220
         X2              =   5220
         Y1              =   8385
         Y2              =   6960
      End
      Begin VB.Line Line40 
         Index           =   2
         X1              =   7980
         X2              =   7980
         Y1              =   8385
         Y2              =   6960
      End
      Begin VB.Line Line40 
         Index           =   3
         X1              =   13200
         X2              =   13200
         Y1              =   8385
         Y2              =   6960
      End
      Begin VB.Label Label38 
         Caption         =   "日期："
         Height          =   195
         Left            =   -62040
         TabIndex        =   49
         Top             =   2580
         Width           =   945
      End
      Begin VB.Line Line37 
         Index           =   1
         X1              =   -62100
         X2              =   -62100
         Y1              =   1950
         Y2              =   3150
      End
      Begin VB.Line Line36 
         Index           =   1
         X1              =   -64200
         X2              =   -64200
         Y1              =   1950
         Y2              =   3150
      End
      Begin VB.Line Line35 
         Index           =   1
         X1              =   -75000
         X2              =   -60060
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line34 
         Index           =   1
         X1              =   -75000
         X2              =   -60060
         Y1              =   2190
         Y2              =   2190
      End
      Begin VB.Label Label37 
         Caption         =   "质量控制签名："
         Height          =   195
         Left            =   -62040
         TabIndex        =   48
         Top             =   2010
         Width           =   1275
      End
      Begin VB.Label Label36 
         Caption         =   "日期："
         Height          =   195
         Left            =   -64110
         TabIndex        =   47
         Top             =   2580
         Width           =   945
      End
      Begin VB.Label Label35 
         Caption         =   "客户签名："
         Height          =   225
         Left            =   -64110
         TabIndex        =   46
         Top             =   1980
         Width           =   945
      End
      Begin VB.Shape Shape2 
         Height          =   3165
         Index           =   1
         Left            =   -74970
         Top             =   0
         Width           =   14985
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
         Left            =   -74910
         TabIndex        =   45
         Top             =   2640
         Width           =   885
      End
      Begin VB.Label Label32 
         Caption         =   "加班工时"
         Height          =   165
         Left            =   -66330
         TabIndex        =   44
         Top             =   2010
         Width           =   1035
      End
      Begin VB.Label Label31 
         Caption         =   "旅途时间"
         Height          =   165
         Left            =   -68760
         TabIndex        =   43
         Top             =   2010
         Width           =   1035
      End
      Begin VB.Label Label30 
         Caption         =   "完成时间"
         Height          =   165
         Left            =   -71880
         TabIndex        =   42
         Top             =   2010
         Width           =   1035
      End
      Begin VB.Label Label29 
         Caption         =   "到达时间"
         Height          =   165
         Left            =   -74850
         TabIndex        =   41
         Top             =   2010
         Width           =   1035
      End
      Begin VB.Line Line33 
         Index           =   1
         X1              =   -75000
         X2              =   -60060
         Y1              =   1950
         Y2              =   1950
      End
      Begin VB.Line Line25 
         Index           =   1
         X1              =   -75000
         X2              =   -59940
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Line Line24 
         Index           =   1
         X1              =   -75000
         X2              =   -59970
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label Label28 
         Caption         =   "复核意见"
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
         Left            =   -74880
         TabIndex        =   40
         Top             =   1470
         Width           =   1035
      End
      Begin VB.Label Label27 
         Caption         =   "对机组运行的建议"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74910
         TabIndex        =   39
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label26 
         Caption         =   "工作总结"
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
         Left            =   -74910
         TabIndex        =   38
         Top             =   60
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   $"NewGzd5.frx":06A5
         Height          =   5685
         Left            =   330
         TabIndex        =   94
         Top             =   600
         Width           =   1395
      End
      Begin VB.Shape Shape1 
         Height          =   6075
         Left            =   60
         Top             =   300
         Width           =   14895
      End
      Begin VB.Label Label7 
         Caption         =   "压缩机序列号"
         Height          =   225
         Left            =   6660
         TabIndex        =   91
         Top             =   990
         Width           =   1425
      End
      Begin VB.Label Label8 
         Caption         =   "满载电流"
         Height          =   255
         Left            =   6660
         TabIndex        =   90
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "冷冻油种类"
         Height          =   255
         Left            =   3420
         TabIndex        =   93
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label Label5 
         Caption         =   "冷冻油充注量"
         Height          =   255
         Left            =   3420
         TabIndex        =   92
         Top             =   1320
         Width           =   1485
      End
      Begin VB.Label Label3 
         Caption         =   "冷媒充注量"
         Height          =   255
         Left            =   3420
         TabIndex        =   88
         Top             =   600
         Width           =   1395
      End
   End
   Begin MSDataGridLib.DataGrid comHtbh 
      Height          =   1155
      Left            =   6030
      TabIndex        =   21
      Top             =   30
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
      Left            =   1920
      TabIndex        =   20
      Top             =   750
      Width           =   4125
   End
   Begin VB.CommandButton cmdBack 
      Height          =   375
      Left            =   14520
      Picture         =   "NewGzd5.frx":076F
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "返回"
      Top             =   10380
      Width           =   465
   End
   Begin VB.CommandButton cmdMod 
      Height          =   375
      Left            =   13560
      Picture         =   "NewGzd5.frx":0871
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "修改"
      Top             =   10380
      Width           =   465
   End
   Begin VB.CommandButton cmdQm 
      Caption         =   "cmdQm"
      Height          =   345
      Index           =   0
      Left            =   9900
      TabIndex        =   12
      Top             =   10380
      Width           =   945
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
      Left            =   450
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "NewGzd5.frx":0B7B
      Top             =   90
      Width           =   1365
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
      Left            =   6480
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "NewGzd5.frx":0BB2
      Top             =   150
      Width           =   1335
   End
   Begin VB.CheckBox C1 
      Alignment       =   1  'Right Justify
      Caption         =   "1号"
      Height          =   285
      Index           =   0
      Left            =   12870
      TabIndex        =   9
      Top             =   690
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox TA 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   13710
      TabIndex        =   8
      Top             =   570
      Visible         =   0   'False
      Width           =   1395
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
      Left            =   1920
      TabIndex        =   7
      Top             =   90
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
      Left            =   12900
      TabIndex        =   6
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
      Index           =   2
      Left            =   1920
      TabIndex        =   5
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
      Height          =   210
      Index           =   3
      Left            =   1920
      TabIndex        =   4
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
      Index           =   4
      Left            =   8010
      TabIndex        =   3
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
      Index           =   5
      Left            =   8010
      TabIndex        =   2
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
      Height          =   240
      Index           =   6
      Left            =   8010
      Locked          =   -1  'True
      TabIndex        =   1
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
      Index           =   7
      Left            =   12600
      TabIndex        =   0
      Text            =   "的"
      Top             =   90
      Visible         =   0   'False
      Width           =   4245
   End
   Begin MSComCtl2.DTPicker dtpA 
      Height          =   225
      Left            =   8040
      TabIndex        =   17
      Top             =   960
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   397
      _Version        =   393216
      Format          =   149618689
      CurrentDate     =   38897
   End
   Begin VB.Label LBLKjj 
      Caption         =   $"NewGzd5.frx":0BDB
      ForeColor       =   &H8000000D&
      Height          =   1305
      Left            =   13170
      TabIndex        =   217
      Top             =   330
      Width           =   2835
   End
   Begin VB.Line Line38 
      Index           =   1
      X1              =   1890
      X2              =   6045
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label39 
      Caption         =   "NO:"
      Height          =   255
      Left            =   12750
      TabIndex        =   98
      Top             =   150
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
      Left            =   13320
      TabIndex        =   97
      Top             =   150
      Width           =   1605
   End
   Begin VB.Label lblkhdh 
      Caption         =   "lblkhdh"
      Height          =   225
      Left            =   10020
      TabIndex        =   19
      Top             =   1410
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label lblGid 
      Caption         =   "lblGid"
      Height          =   225
      Left            =   8160
      TabIndex        =   18
      Top             =   1350
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   10920
      TabIndex        =   16
      Top             =   10440
      Width           =   1905
   End
   Begin VB.Label lblQM 
      Caption         =   "签字提交"
      Height          =   225
      Index           =   0
      Left            =   8970
      TabIndex        =   15
      Top             =   10440
      Width           =   795
   End
   Begin VB.Line Line26 
      X1              =   1920
      X2              =   6015
      Y1              =   330
      Y2              =   330
   End
   Begin VB.Line Line27 
      X1              =   1920
      X2              =   6000
      Y1              =   735
      Y2              =   735
   End
   Begin VB.Line Line28 
      X1              =   1920
      X2              =   6000
      Y1              =   1155
      Y2              =   1155
   End
   Begin VB.Line Line30 
      X1              =   8010
      X2              =   12090
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line31 
      X1              =   8010
      X2              =   12090
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Line Line32 
      X1              =   8010
      X2              =   12090
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "NewGzd5"
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
If C1(9).Value = 1 Then
    For oo = 9 To 20
        C1(oo).Value = 0
    Next
Else
    For oo = 9 To 20
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
For oo = 1 To 110
    mod1.HTP.Update "mat" & oo, TA(oo).Text
Next
For oo = 1 To 22
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
Private Sub BA_Click(Index As Integer)
dtgRen.Visible = False
comHtbh.Visible = False
comXmmc.Visible = False
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Height = 11400
Me.Width = 15135

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
For oo = 1 To 110
    TA(oo).Tag = 50
Next
TA(19).Tag = 200
TA(20).Tag = 200
TA(85).Tag = 200
TA(107).Tag = 200
TA(108).Tag = 200
TA(109).Tag = 200
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
Private Sub cmdMod_Click()
cmdSave.Enabled = True
End Sub
Private Sub dtpA_CloseUp()
BA(6).Text = Format(dtpA.Value, "YYYY/MM/DD", vbUseSystemDayOfWeek)
End Sub
Private Sub TA_Click(Index As Integer)
comXmmc.Visible = False
comHtbh.Visible = False
dtgRen.Visible = False
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
