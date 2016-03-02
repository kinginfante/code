VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form NewGzd6 
   Caption         =   "机组大修工作报告（单）"
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
      Left            =   1620
      TabIndex        =   129
      Top             =   1350
      Width           =   4065
   End
   Begin MSDataGridLib.DataGrid dtgRen 
      Height          =   8085
      Left            =   10140
      TabIndex        =   126
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
   Begin TabDlg.SSTab tabNr 
      Height          =   8655
      Left            =   0
      TabIndex        =   23
      Top             =   1710
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   15266
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "内容1"
      TabPicture(0)   =   "NewGzd6.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label42"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label41"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label33(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label32(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label31(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label30(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label29(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label28(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label15"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Line16"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Line15(3)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Line14(3)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Line13(3)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Line15(2)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Line14(2)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Line13(2)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Line15(1)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Line14(1)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Line13(1)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Line15(0)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Line14(0)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Line13(0)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Line12"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Line11"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Line10"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Line7(2)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Line6(2)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Line7(1)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Line6(1)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Line9"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Line8"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Line7(0)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Line6(0)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Line5"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Line2(21)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Line4"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Line3(5)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Line3(4)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Line3(3)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Line3(2)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Line3(1)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Line3(0)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Line2(20)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Line2(19)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "Line2(18)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Line2(17)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Line2(16)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Line2(15)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Line2(14)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "Line2(13)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "Line2(12)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "Line2(11)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "Line2(10)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "Line2(9)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "Line2(8)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "Line2(7)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "Line2(6)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "Line2(5)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "Line2(4)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "Line2(3)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "Line2(2)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "Line2(1)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "Line2(0)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "Shape1"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "Label47"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "Label46"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "Label45"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "Label44"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "Label43"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "Label40"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "Label39(0)"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "Label38(0)"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "Label37(0)"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "Label36(0)"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "Label35(0)"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "Label34(0)"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "Label27"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "Label26"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).Control(84)=   "Label25"
      Tab(0).Control(84).Enabled=   0   'False
      Tab(0).Control(85)=   "Label24"
      Tab(0).Control(85).Enabled=   0   'False
      Tab(0).Control(86)=   "Label23"
      Tab(0).Control(86).Enabled=   0   'False
      Tab(0).Control(87)=   "Label22"
      Tab(0).Control(87).Enabled=   0   'False
      Tab(0).Control(88)=   "Label21"
      Tab(0).Control(88).Enabled=   0   'False
      Tab(0).Control(89)=   "Label20"
      Tab(0).Control(89).Enabled=   0   'False
      Tab(0).Control(90)=   "Label19"
      Tab(0).Control(90).Enabled=   0   'False
      Tab(0).Control(91)=   "Label18"
      Tab(0).Control(91).Enabled=   0   'False
      Tab(0).Control(92)=   "Label17"
      Tab(0).Control(92).Enabled=   0   'False
      Tab(0).Control(93)=   "Label16"
      Tab(0).Control(93).Enabled=   0   'False
      Tab(0).Control(94)=   "Label14"
      Tab(0).Control(94).Enabled=   0   'False
      Tab(0).Control(95)=   "Label13"
      Tab(0).Control(95).Enabled=   0   'False
      Tab(0).Control(96)=   "Label12"
      Tab(0).Control(96).Enabled=   0   'False
      Tab(0).Control(97)=   "Label11"
      Tab(0).Control(97).Enabled=   0   'False
      Tab(0).Control(98)=   "Label10"
      Tab(0).Control(98).Enabled=   0   'False
      Tab(0).Control(99)=   "Label7"
      Tab(0).Control(99).Enabled=   0   'False
      Tab(0).Control(100)=   "Label6"
      Tab(0).Control(100).Enabled=   0   'False
      Tab(0).Control(101)=   "Label5"
      Tab(0).Control(101).Enabled=   0   'False
      Tab(0).Control(102)=   "TA(105)"
      Tab(0).Control(102).Enabled=   0   'False
      Tab(0).Control(103)=   "TA(104)"
      Tab(0).Control(103).Enabled=   0   'False
      Tab(0).Control(104)=   "TA(103)"
      Tab(0).Control(104).Enabled=   0   'False
      Tab(0).Control(105)=   "TA(102)"
      Tab(0).Control(105).Enabled=   0   'False
      Tab(0).Control(106)=   "TA(101)"
      Tab(0).Control(106).Enabled=   0   'False
      Tab(0).Control(107)=   "TA(100)"
      Tab(0).Control(107).Enabled=   0   'False
      Tab(0).Control(108)=   "TA(99)"
      Tab(0).Control(108).Enabled=   0   'False
      Tab(0).Control(109)=   "TA(98)"
      Tab(0).Control(109).Enabled=   0   'False
      Tab(0).Control(110)=   "TA(97)"
      Tab(0).Control(110).Enabled=   0   'False
      Tab(0).Control(111)=   "TA(96)"
      Tab(0).Control(111).Enabled=   0   'False
      Tab(0).Control(112)=   "TA(95)"
      Tab(0).Control(112).Enabled=   0   'False
      Tab(0).Control(113)=   "TA(94)"
      Tab(0).Control(113).Enabled=   0   'False
      Tab(0).Control(114)=   "TA(93)"
      Tab(0).Control(114).Enabled=   0   'False
      Tab(0).Control(115)=   "TA(92)"
      Tab(0).Control(115).Enabled=   0   'False
      Tab(0).Control(116)=   "TA(91)"
      Tab(0).Control(116).Enabled=   0   'False
      Tab(0).Control(117)=   "TA(90)"
      Tab(0).Control(117).Enabled=   0   'False
      Tab(0).Control(118)=   "TA(89)"
      Tab(0).Control(118).Enabled=   0   'False
      Tab(0).Control(119)=   "TA(88)"
      Tab(0).Control(119).Enabled=   0   'False
      Tab(0).Control(120)=   "TA(87)"
      Tab(0).Control(120).Enabled=   0   'False
      Tab(0).Control(121)=   "TA(86)"
      Tab(0).Control(121).Enabled=   0   'False
      Tab(0).Control(122)=   "TA(85)"
      Tab(0).Control(122).Enabled=   0   'False
      Tab(0).Control(123)=   "TA(84)"
      Tab(0).Control(123).Enabled=   0   'False
      Tab(0).Control(124)=   "TA(83)"
      Tab(0).Control(124).Enabled=   0   'False
      Tab(0).Control(125)=   "TA(82)"
      Tab(0).Control(125).Enabled=   0   'False
      Tab(0).Control(126)=   "TA(81)"
      Tab(0).Control(126).Enabled=   0   'False
      Tab(0).Control(127)=   "TA(80)"
      Tab(0).Control(127).Enabled=   0   'False
      Tab(0).Control(128)=   "C1(32)"
      Tab(0).Control(128).Enabled=   0   'False
      Tab(0).Control(129)=   "C1(31)"
      Tab(0).Control(129).Enabled=   0   'False
      Tab(0).Control(130)=   "C1(30)"
      Tab(0).Control(130).Enabled=   0   'False
      Tab(0).Control(131)=   "C1(29)"
      Tab(0).Control(131).Enabled=   0   'False
      Tab(0).Control(132)=   "C1(28)"
      Tab(0).Control(132).Enabled=   0   'False
      Tab(0).Control(133)=   "C1(27)"
      Tab(0).Control(133).Enabled=   0   'False
      Tab(0).Control(134)=   "C1(26)"
      Tab(0).Control(134).Enabled=   0   'False
      Tab(0).Control(135)=   "C1(25)"
      Tab(0).Control(135).Enabled=   0   'False
      Tab(0).Control(136)=   "C1(24)"
      Tab(0).Control(136).Enabled=   0   'False
      Tab(0).Control(137)=   "C1(23)"
      Tab(0).Control(137).Enabled=   0   'False
      Tab(0).Control(138)=   "C1(22)"
      Tab(0).Control(138).Enabled=   0   'False
      Tab(0).Control(139)=   "C1(21)"
      Tab(0).Control(139).Enabled=   0   'False
      Tab(0).Control(140)=   "C1(20)"
      Tab(0).Control(140).Enabled=   0   'False
      Tab(0).Control(141)=   "C1(19)"
      Tab(0).Control(141).Enabled=   0   'False
      Tab(0).Control(142)=   "C1(18)"
      Tab(0).Control(142).Enabled=   0   'False
      Tab(0).Control(143)=   "C1(17)"
      Tab(0).Control(143).Enabled=   0   'False
      Tab(0).Control(144)=   "C1(16)"
      Tab(0).Control(144).Enabled=   0   'False
      Tab(0).Control(145)=   "C1(15)"
      Tab(0).Control(145).Enabled=   0   'False
      Tab(0).Control(146)=   "C1(14)"
      Tab(0).Control(146).Enabled=   0   'False
      Tab(0).Control(147)=   "C1(13)"
      Tab(0).Control(147).Enabled=   0   'False
      Tab(0).Control(148)=   "C1(12)"
      Tab(0).Control(148).Enabled=   0   'False
      Tab(0).Control(149)=   "TA(79)"
      Tab(0).Control(149).Enabled=   0   'False
      Tab(0).Control(150)=   "TA(78)"
      Tab(0).Control(150).Enabled=   0   'False
      Tab(0).Control(151)=   "TA(77)"
      Tab(0).Control(151).Enabled=   0   'False
      Tab(0).Control(152)=   "TA(76)"
      Tab(0).Control(152).Enabled=   0   'False
      Tab(0).Control(153)=   "TA(75)"
      Tab(0).Control(153).Enabled=   0   'False
      Tab(0).Control(154)=   "TA(74)"
      Tab(0).Control(154).Enabled=   0   'False
      Tab(0).Control(155)=   "TA(73)"
      Tab(0).Control(155).Enabled=   0   'False
      Tab(0).Control(156)=   "TA(72)"
      Tab(0).Control(156).Enabled=   0   'False
      Tab(0).Control(157)=   "TA(71)"
      Tab(0).Control(157).Enabled=   0   'False
      Tab(0).Control(158)=   "TA(70)"
      Tab(0).Control(158).Enabled=   0   'False
      Tab(0).Control(159)=   "TA(69)"
      Tab(0).Control(159).Enabled=   0   'False
      Tab(0).Control(160)=   "TA(68)"
      Tab(0).Control(160).Enabled=   0   'False
      Tab(0).Control(161)=   "TA(67)"
      Tab(0).Control(161).Enabled=   0   'False
      Tab(0).Control(162)=   "TA(66)"
      Tab(0).Control(162).Enabled=   0   'False
      Tab(0).Control(163)=   "C1(11)"
      Tab(0).Control(163).Enabled=   0   'False
      Tab(0).Control(164)=   "C1(10)"
      Tab(0).Control(164).Enabled=   0   'False
      Tab(0).Control(165)=   "C1(9)"
      Tab(0).Control(165).Enabled=   0   'False
      Tab(0).Control(166)=   "C1(8)"
      Tab(0).Control(166).Enabled=   0   'False
      Tab(0).Control(167)=   "C1(7)"
      Tab(0).Control(167).Enabled=   0   'False
      Tab(0).Control(168)=   "C1(6)"
      Tab(0).Control(168).Enabled=   0   'False
      Tab(0).Control(169)=   "C1(5)"
      Tab(0).Control(169).Enabled=   0   'False
      Tab(0).Control(170)=   "C1(4)"
      Tab(0).Control(170).Enabled=   0   'False
      Tab(0).Control(171)=   "C1(3)"
      Tab(0).Control(171).Enabled=   0   'False
      Tab(0).Control(172)=   "TA(65)"
      Tab(0).Control(172).Enabled=   0   'False
      Tab(0).Control(173)=   "TA(64)"
      Tab(0).Control(173).Enabled=   0   'False
      Tab(0).Control(174)=   "TA(63)"
      Tab(0).Control(174).Enabled=   0   'False
      Tab(0).Control(175)=   "TA(62)"
      Tab(0).Control(175).Enabled=   0   'False
      Tab(0).Control(176)=   "TA(61)"
      Tab(0).Control(176).Enabled=   0   'False
      Tab(0).Control(177)=   "TA(60)"
      Tab(0).Control(177).Enabled=   0   'False
      Tab(0).Control(178)=   "TA(59)"
      Tab(0).Control(178).Enabled=   0   'False
      Tab(0).Control(179)=   "TA(58)"
      Tab(0).Control(179).Enabled=   0   'False
      Tab(0).Control(180)=   "TA(57)"
      Tab(0).Control(180).Enabled=   0   'False
      Tab(0).Control(181)=   "TA(56)"
      Tab(0).Control(181).Enabled=   0   'False
      Tab(0).Control(182)=   "TA(55)"
      Tab(0).Control(182).Enabled=   0   'False
      Tab(0).Control(183)=   "TA(54)"
      Tab(0).Control(183).Enabled=   0   'False
      Tab(0).Control(184)=   "TA(53)"
      Tab(0).Control(184).Enabled=   0   'False
      Tab(0).Control(185)=   "TA(52)"
      Tab(0).Control(185).Enabled=   0   'False
      Tab(0).Control(186)=   "TA(51)"
      Tab(0).Control(186).Enabled=   0   'False
      Tab(0).Control(187)=   "TA(50)"
      Tab(0).Control(187).Enabled=   0   'False
      Tab(0).Control(188)=   "TA(49)"
      Tab(0).Control(188).Enabled=   0   'False
      Tab(0).Control(189)=   "TA(48)"
      Tab(0).Control(189).Enabled=   0   'False
      Tab(0).Control(190)=   "TA(47)"
      Tab(0).Control(190).Enabled=   0   'False
      Tab(0).Control(191)=   "TA(46)"
      Tab(0).Control(191).Enabled=   0   'False
      Tab(0).Control(192)=   "TA(45)"
      Tab(0).Control(192).Enabled=   0   'False
      Tab(0).Control(193)=   "TA(44)"
      Tab(0).Control(193).Enabled=   0   'False
      Tab(0).Control(194)=   "TA(43)"
      Tab(0).Control(194).Enabled=   0   'False
      Tab(0).Control(195)=   "TA(42)"
      Tab(0).Control(195).Enabled=   0   'False
      Tab(0).Control(196)=   "TA(41)"
      Tab(0).Control(196).Enabled=   0   'False
      Tab(0).Control(197)=   "TA(40)"
      Tab(0).Control(197).Enabled=   0   'False
      Tab(0).Control(198)=   "TA(39)"
      Tab(0).Control(198).Enabled=   0   'False
      Tab(0).Control(199)=   "TA(38)"
      Tab(0).Control(199).Enabled=   0   'False
      Tab(0).Control(200)=   "TA(37)"
      Tab(0).Control(200).Enabled=   0   'False
      Tab(0).Control(201)=   "TA(36)"
      Tab(0).Control(201).Enabled=   0   'False
      Tab(0).Control(202)=   "TA(35)"
      Tab(0).Control(202).Enabled=   0   'False
      Tab(0).Control(203)=   "TA(34)"
      Tab(0).Control(203).Enabled=   0   'False
      Tab(0).Control(204)=   "TA(33)"
      Tab(0).Control(204).Enabled=   0   'False
      Tab(0).Control(205)=   "TA(32)"
      Tab(0).Control(205).Enabled=   0   'False
      Tab(0).Control(206)=   "TA(31)"
      Tab(0).Control(206).Enabled=   0   'False
      Tab(0).Control(207)=   "TA(30)"
      Tab(0).Control(207).Enabled=   0   'False
      Tab(0).Control(208)=   "TA(29)"
      Tab(0).Control(208).Enabled=   0   'False
      Tab(0).Control(209)=   "TA(28)"
      Tab(0).Control(209).Enabled=   0   'False
      Tab(0).Control(210)=   "TA(27)"
      Tab(0).Control(210).Enabled=   0   'False
      Tab(0).Control(211)=   "TA(26)"
      Tab(0).Control(211).Enabled=   0   'False
      Tab(0).Control(212)=   "TA(25)"
      Tab(0).Control(212).Enabled=   0   'False
      Tab(0).Control(213)=   "TA(24)"
      Tab(0).Control(213).Enabled=   0   'False
      Tab(0).Control(214)=   "TA(23)"
      Tab(0).Control(214).Enabled=   0   'False
      Tab(0).Control(215)=   "TA(22)"
      Tab(0).Control(215).Enabled=   0   'False
      Tab(0).Control(216)=   "TA(21)"
      Tab(0).Control(216).Enabled=   0   'False
      Tab(0).Control(217)=   "C1(2)"
      Tab(0).Control(217).Enabled=   0   'False
      Tab(0).Control(218)=   "C1(1)"
      Tab(0).Control(218).Enabled=   0   'False
      Tab(0).Control(219)=   "TA(20)"
      Tab(0).Control(219).Enabled=   0   'False
      Tab(0).Control(220)=   "TA(19)"
      Tab(0).Control(220).Enabled=   0   'False
      Tab(0).Control(221)=   "TA(18)"
      Tab(0).Control(221).Enabled=   0   'False
      Tab(0).Control(222)=   "TA(17)"
      Tab(0).Control(222).Enabled=   0   'False
      Tab(0).Control(223)=   "TA(16)"
      Tab(0).Control(223).Enabled=   0   'False
      Tab(0).Control(224)=   "TA(15)"
      Tab(0).Control(224).Enabled=   0   'False
      Tab(0).Control(225)=   "TA(14)"
      Tab(0).Control(225).Enabled=   0   'False
      Tab(0).Control(226)=   "TA(13)"
      Tab(0).Control(226).Enabled=   0   'False
      Tab(0).Control(227)=   "TA(12)"
      Tab(0).Control(227).Enabled=   0   'False
      Tab(0).Control(228)=   "TA(11)"
      Tab(0).Control(228).Enabled=   0   'False
      Tab(0).Control(229)=   "TA(10)"
      Tab(0).Control(229).Enabled=   0   'False
      Tab(0).Control(230)=   "TA(9)"
      Tab(0).Control(230).Enabled=   0   'False
      Tab(0).Control(231)=   "TA(8)"
      Tab(0).Control(231).Enabled=   0   'False
      Tab(0).Control(232)=   "TA(7)"
      Tab(0).Control(232).Enabled=   0   'False
      Tab(0).Control(233)=   "TA(6)"
      Tab(0).Control(233).Enabled=   0   'False
      Tab(0).Control(234)=   "TA(5)"
      Tab(0).Control(234).Enabled=   0   'False
      Tab(0).Control(235)=   "TA(4)"
      Tab(0).Control(235).Enabled=   0   'False
      Tab(0).Control(236)=   "TA(3)"
      Tab(0).Control(236).Enabled=   0   'False
      Tab(0).Control(237)=   "TA(2)"
      Tab(0).Control(237).Enabled=   0   'False
      Tab(0).Control(238)=   "TA(1)"
      Tab(0).Control(238).Enabled=   0   'False
      Tab(0).Control(239)=   "cmdAll"
      Tab(0).Control(239).Enabled=   0   'False
      Tab(0).Control(240)=   "cmdA1"
      Tab(0).Control(240).Enabled=   0   'False
      Tab(0).ControlCount=   241
      TabCaption(1)   =   "内容2"
      TabPicture(1)   =   "NewGzd6.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "BA(16)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "BA(15)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "BA(14)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "BA(13)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "BA(12)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "BA(11)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "BA(10)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "BA(9)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "BA(8)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "TA(108)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Text3"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "dtpB"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "dtpC"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "TA(106)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "TA(107)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Shape2"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label38(1)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Line37"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Line36"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Line35"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Line34"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Label37(1)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Label36(1)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Label35(1)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Label34(1)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Label32(1)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Label31(1)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Label30(1)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Label29(1)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Line33"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Line25"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Label28(1)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Label48"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).ControlCount=   34
      Begin VB.CommandButton cmdA1 
         Caption         =   "全部"
         Height          =   255
         Left            =   13950
         TabIndex        =   245
         Top             =   6480
         Width           =   705
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "全部"
         Height          =   225
         Left            =   8670
         TabIndex        =   244
         Top             =   6810
         Width           =   735
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   1
         Left            =   1950
         TabIndex        =   130
         Top             =   360
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   2
         Left            =   3900
         TabIndex        =   133
         Top             =   360
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   3
         Left            =   5670
         TabIndex        =   136
         Top             =   360
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   4
         Left            =   7530
         TabIndex        =   139
         Top             =   360
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   5
         Left            =   1950
         TabIndex        =   131
         Top             =   690
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   6
         Left            =   3900
         TabIndex        =   134
         Top             =   690
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   7
         Left            =   5670
         TabIndex        =   137
         Top             =   690
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   8
         Left            =   7530
         TabIndex        =   140
         Top             =   690
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   9
         Left            =   1950
         TabIndex        =   132
         Top             =   1020
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   10
         Left            =   3900
         TabIndex        =   135
         Top             =   1020
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   11
         Left            =   5670
         TabIndex        =   138
         Top             =   1020
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   12
         Left            =   7530
         TabIndex        =   141
         Top             =   1020
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   13
         Left            =   12390
         TabIndex        =   144
         Top             =   360
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   14
         Left            =   12390
         TabIndex        =   145
         Top             =   690
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   1020
         Index           =   15
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   142
         Top             =   1590
         Width           =   8895
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   1020
         Index           =   16
         Left            =   9390
         TabIndex        =   146
         Top             =   1590
         Width           =   2505
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   1020
         Index           =   17
         Left            =   12090
         TabIndex        =   147
         Top             =   1590
         Width           =   2505
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   990
         Index           =   18
         Left            =   120
         TabIndex        =   143
         Top             =   2970
         Width           =   8895
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   990
         Index           =   19
         Left            =   9330
         TabIndex        =   148
         Top             =   2970
         Width           =   3555
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   990
         Index           =   20
         Left            =   13140
         TabIndex        =   149
         Top             =   2970
         Width           =   1455
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "大修进行中"
         Height          =   225
         Index           =   1
         Left            =   1500
         TabIndex        =   78
         Top             =   4020
         Width           =   1545
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "／ 大修已完毕"
         Height          =   225
         Index           =   2
         Left            =   3090
         TabIndex        =   77
         Top             =   4020
         Width           =   1545
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   21
         Left            =   1590
         TabIndex        =   150
         Top             =   4590
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   22
         Left            =   2850
         TabIndex        =   159
         Top             =   4590
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   23
         Left            =   4110
         TabIndex        =   168
         Top             =   4590
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   24
         Left            =   5400
         TabIndex        =   177
         Top             =   4590
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   25
         Left            =   6720
         TabIndex        =   186
         Top             =   4590
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   26
         Left            =   1590
         TabIndex        =   151
         Top             =   4860
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   27
         Left            =   2850
         TabIndex        =   160
         Top             =   4860
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   28
         Left            =   4110
         TabIndex        =   169
         Top             =   4860
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   29
         Left            =   5400
         TabIndex        =   178
         Top             =   4860
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   30
         Left            =   6720
         TabIndex        =   187
         Top             =   4860
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   31
         Left            =   1590
         TabIndex        =   152
         Top             =   5130
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   32
         Left            =   2850
         TabIndex        =   161
         Top             =   5130
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   33
         Left            =   4110
         TabIndex        =   170
         Top             =   5130
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   34
         Left            =   5400
         TabIndex        =   179
         Top             =   5130
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   35
         Left            =   6720
         TabIndex        =   188
         Top             =   5130
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   36
         Left            =   1590
         TabIndex        =   153
         Top             =   5400
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   37
         Left            =   2850
         TabIndex        =   162
         Top             =   5400
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   38
         Left            =   4110
         TabIndex        =   171
         Top             =   5400
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   39
         Left            =   5400
         TabIndex        =   180
         Top             =   5400
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   40
         Left            =   6720
         TabIndex        =   189
         Top             =   5400
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   41
         Left            =   1590
         TabIndex        =   154
         Top             =   5700
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   42
         Left            =   2850
         TabIndex        =   163
         Top             =   5700
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   43
         Left            =   4110
         TabIndex        =   172
         Top             =   5700
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   44
         Left            =   5400
         TabIndex        =   181
         Top             =   5700
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   45
         Left            =   6720
         TabIndex        =   190
         Top             =   5700
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   46
         Left            =   1590
         TabIndex        =   155
         Top             =   5970
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   47
         Left            =   2850
         TabIndex        =   164
         Top             =   5970
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   48
         Left            =   4110
         TabIndex        =   173
         Top             =   5970
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   49
         Left            =   5400
         TabIndex        =   182
         Top             =   5970
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   50
         Left            =   6720
         TabIndex        =   191
         Top             =   5970
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   51
         Left            =   1590
         TabIndex        =   156
         Top             =   6240
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   52
         Left            =   2850
         TabIndex        =   165
         Top             =   6240
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   53
         Left            =   4110
         TabIndex        =   174
         Top             =   6240
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   54
         Left            =   5400
         TabIndex        =   183
         Top             =   6240
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   55
         Left            =   6720
         TabIndex        =   192
         Top             =   6240
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   56
         Left            =   1590
         TabIndex        =   157
         Top             =   6510
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   57
         Left            =   2850
         TabIndex        =   166
         Top             =   6510
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   58
         Left            =   4110
         TabIndex        =   175
         Top             =   6510
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   59
         Left            =   5400
         TabIndex        =   184
         Top             =   6510
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   60
         Left            =   6720
         TabIndex        =   193
         Top             =   6510
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   61
         Left            =   1590
         TabIndex        =   158
         Top             =   6810
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   62
         Left            =   2850
         TabIndex        =   167
         Top             =   6810
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   63
         Left            =   4110
         TabIndex        =   176
         Top             =   6810
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   64
         Left            =   5400
         TabIndex        =   185
         Top             =   6810
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   65
         Left            =   6720
         TabIndex        =   194
         Top             =   6810
         Width           =   1125
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   195
         Index           =   3
         Left            =   8220
         TabIndex        =   76
         Top             =   4590
         Width           =   375
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   195
         Index           =   4
         Left            =   8220
         TabIndex        =   75
         Top             =   4860
         Width           =   375
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   195
         Index           =   5
         Left            =   8220
         TabIndex        =   74
         Top             =   5145
         Width           =   375
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   195
         Index           =   6
         Left            =   8220
         TabIndex        =   73
         Top             =   5415
         Width           =   375
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   195
         Index           =   7
         Left            =   8220
         TabIndex        =   72
         Top             =   5700
         Width           =   375
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   195
         Index           =   8
         Left            =   8220
         TabIndex        =   71
         Top             =   5970
         Width           =   375
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   195
         Index           =   9
         Left            =   8220
         TabIndex        =   70
         Top             =   6255
         Width           =   375
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   195
         Index           =   10
         Left            =   8220
         TabIndex        =   69
         Top             =   6525
         Width           =   375
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   195
         Index           =   11
         Left            =   8220
         TabIndex        =   68
         Top             =   6810
         Width           =   375
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   66
         Left            =   10890
         TabIndex        =   195
         Top             =   4590
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   67
         Left            =   12420
         TabIndex        =   202
         Top             =   4590
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   68
         Left            =   10890
         TabIndex        =   196
         Top             =   4860
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   69
         Left            =   12420
         TabIndex        =   203
         Top             =   4860
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   70
         Left            =   10890
         TabIndex        =   197
         Top             =   5160
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   71
         Left            =   12420
         TabIndex        =   204
         Top             =   5160
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   72
         Left            =   10890
         TabIndex        =   198
         Top             =   5430
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   73
         Left            =   12420
         TabIndex        =   205
         Top             =   5430
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   74
         Left            =   10890
         TabIndex        =   199
         Top             =   5700
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   75
         Left            =   12420
         TabIndex        =   206
         Top             =   5700
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   76
         Left            =   10890
         TabIndex        =   200
         Top             =   6000
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   77
         Left            =   12420
         TabIndex        =   207
         Top             =   6000
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   78
         Left            =   10890
         TabIndex        =   201
         Top             =   6255
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   79
         Left            =   12420
         TabIndex        =   208
         Top             =   6255
         Width           =   1125
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   195
         Index           =   12
         Left            =   13920
         TabIndex        =   67
         Top             =   4590
         Width           =   375
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   195
         Index           =   13
         Left            =   13920
         TabIndex        =   66
         Top             =   4860
         Width           =   375
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   195
         Index           =   14
         Left            =   13920
         TabIndex        =   65
         Top             =   5145
         Width           =   375
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   195
         Index           =   15
         Left            =   13920
         TabIndex        =   64
         Top             =   5415
         Width           =   375
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   195
         Index           =   16
         Left            =   13920
         TabIndex        =   63
         Top             =   5700
         Width           =   375
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   195
         Index           =   17
         Left            =   13920
         TabIndex        =   62
         Top             =   5970
         Width           =   375
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   195
         Index           =   18
         Left            =   13920
         TabIndex        =   61
         Top             =   6255
         Width           =   375
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "机组类型：本机组属于冷水机组，无需填写下列数据"
         Height          =   225
         Index           =   19
         Left            =   120
         TabIndex        =   60
         Top             =   7140
         Width           =   4665
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "/ 本机组属于热泵机组，现检查冷凝器风机后测得参数如下 "
         Height          =   225
         Index           =   20
         Left            =   4890
         TabIndex        =   59
         Top             =   7140
         Width           =   5145
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   225
         Index           =   21
         Left            =   1590
         TabIndex        =   58
         Top             =   7470
         Width           =   525
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "2"
         Height          =   225
         Index           =   22
         Left            =   2640
         TabIndex        =   57
         Top             =   7470
         Width           =   525
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "3"
         Height          =   225
         Index           =   23
         Left            =   3690
         TabIndex        =   56
         Top             =   7470
         Width           =   525
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "4"
         Height          =   225
         Index           =   24
         Left            =   4740
         TabIndex        =   55
         Top             =   7470
         Width           =   525
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "5"
         Height          =   225
         Index           =   25
         Left            =   5790
         TabIndex        =   54
         Top             =   7470
         Width           =   525
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "6"
         Height          =   225
         Index           =   26
         Left            =   6840
         TabIndex        =   53
         Top             =   7470
         Width           =   525
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "7"
         Height          =   225
         Index           =   27
         Left            =   7890
         TabIndex        =   52
         Top             =   7470
         Width           =   525
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "8"
         Height          =   225
         Index           =   28
         Left            =   8940
         TabIndex        =   51
         Top             =   7470
         Width           =   525
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "9"
         Height          =   225
         Index           =   29
         Left            =   9990
         TabIndex        =   50
         Top             =   7470
         Width           =   525
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "10"
         Height          =   225
         Index           =   30
         Left            =   11040
         TabIndex        =   49
         Top             =   7470
         Width           =   525
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "11"
         Height          =   225
         Index           =   31
         Left            =   12090
         TabIndex        =   48
         Top             =   7470
         Width           =   525
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "12"
         Height          =   225
         Index           =   32
         Left            =   13020
         TabIndex        =   47
         Top             =   7470
         Width           =   525
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   80
         Left            =   1500
         TabIndex        =   209
         Top             =   7785
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   81
         Left            =   2535
         TabIndex        =   211
         Top             =   7785
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   82
         Left            =   3585
         TabIndex        =   213
         Top             =   7785
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   83
         Left            =   4620
         TabIndex        =   215
         Top             =   7785
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   84
         Left            =   5670
         TabIndex        =   217
         Top             =   7785
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   85
         Left            =   6705
         TabIndex        =   219
         Top             =   7785
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   86
         Left            =   7740
         TabIndex        =   221
         Top             =   7785
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   87
         Left            =   8790
         TabIndex        =   223
         Top             =   7785
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   88
         Left            =   9825
         TabIndex        =   225
         Top             =   7785
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   89
         Left            =   10875
         TabIndex        =   227
         Top             =   7785
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   90
         Left            =   11910
         TabIndex        =   229
         Top             =   7785
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   91
         Left            =   12960
         TabIndex        =   231
         Top             =   7785
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   92
         Left            =   13920
         TabIndex        =   233
         Top             =   7785
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   93
         Left            =   1500
         TabIndex        =   210
         Top             =   8070
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   94
         Left            =   2535
         TabIndex        =   212
         Top             =   8070
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   95
         Left            =   3585
         TabIndex        =   214
         Top             =   8070
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   96
         Left            =   4620
         TabIndex        =   216
         Top             =   8070
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   97
         Left            =   5670
         TabIndex        =   218
         Top             =   8070
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   98
         Left            =   6705
         TabIndex        =   220
         Top             =   8070
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   99
         Left            =   7740
         TabIndex        =   222
         Top             =   8070
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   100
         Left            =   8790
         TabIndex        =   224
         Top             =   8070
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   101
         Left            =   9825
         TabIndex        =   226
         Top             =   8070
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   102
         Left            =   10875
         TabIndex        =   228
         Top             =   8070
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   103
         Left            =   11910
         TabIndex        =   230
         Top             =   8070
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   104
         Left            =   12960
         TabIndex        =   232
         Top             =   8070
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   105
         Left            =   13920
         TabIndex        =   234
         Top             =   8070
         Width           =   795
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
         Left            =   -62040
         TabIndex        =   33
         Top             =   2100
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
         Index           =   15
         Left            =   -62040
         TabIndex        =   243
         Top             =   1500
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
         Left            =   -64140
         TabIndex        =   32
         Top             =   2100
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
         TabIndex        =   242
         Top             =   1500
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
         TabIndex        =   241
         Text            =   "NewGzd6.frx":0038
         Top             =   1830
         Width           =   9345
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Left            =   -74970
         TabIndex        =   26
         Top             =   1500
         Width           =   10755
         Begin VB.OptionButton FPA 
            Caption         =   "优秀"
            Height          =   195
            Left            =   1350
            TabIndex        =   30
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton FPB 
            Caption         =   "满意"
            Height          =   195
            Left            =   2950
            TabIndex        =   29
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton FPC 
            Caption         =   "较满意"
            Height          =   195
            Left            =   4550
            TabIndex        =   28
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton FPD 
            Caption         =   "尚可"
            Height          =   195
            Left            =   6150
            TabIndex        =   27
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
            Index           =   1
            Left            =   60
            TabIndex        =   31
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
         TabIndex        =   240
         Text            =   "的"
         Top             =   1260
         Width           =   1035
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   10
         Left            =   -67650
         TabIndex        =   239
         Text            =   "的"
         Top             =   1260
         Width           =   1035
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   9
         Left            =   -70320
         TabIndex        =   238
         Text            =   "的"
         Top             =   1260
         Width           =   1035
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   8
         Left            =   -73560
         TabIndex        =   237
         Text            =   "的"
         Top             =   1260
         Width           =   1035
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   108
         Left            =   -61260
         TabIndex        =   25
         Top             =   930
         Width           =   1065
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   -61890
         TabIndex        =   24
         Text            =   "复核人:"
         Top             =   1020
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtpB 
         Height          =   225
         Left            =   -64140
         TabIndex        =   34
         Top             =   2100
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   397
         _Version        =   393216
         Format          =   149553153
         CurrentDate     =   38897
      End
      Begin MSComCtl2.DTPicker dtpC 
         Height          =   225
         Left            =   -62040
         TabIndex        =   35
         Top             =   2100
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   397
         _Version        =   393216
         Format          =   149553153
         CurrentDate     =   38897
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   570
         Index           =   106
         Left            =   -73590
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   235
         Top             =   60
         Width           =   13545
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   510
         Index           =   107
         Left            =   -73590
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   236
         Top             =   690
         Width           =   13515
      End
      Begin VB.Label Label5 
         Caption         =   "压缩机型号"
         Height          =   195
         Left            =   150
         TabIndex        =   121
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label Label6 
         Caption         =   "压缩机序列号"
         Height          =   195
         Left            =   150
         TabIndex        =   120
         Top             =   690
         Width           =   1425
      End
      Begin VB.Label Label7 
         Caption         =   "满载电流"
         Height          =   195
         Left            =   150
         TabIndex        =   119
         Top             =   1020
         Width           =   1275
      End
      Begin VB.Label Label10 
         Caption         =   "机组"
         Height          =   165
         Left            =   12720
         TabIndex        =   116
         Top             =   60
         Width           =   1005
      End
      Begin VB.Label Label11 
         Caption         =   "今日工作内容："
         Height          =   195
         Left            =   150
         TabIndex        =   115
         Top             =   1320
         Width           =   1365
      End
      Begin VB.Label Label12 
         Caption         =   "下步工作计划："
         Height          =   225
         Left            =   180
         TabIndex        =   114
         Top             =   2700
         Width           =   1605
      End
      Begin VB.Label Label13 
         Caption         =   "完成度描述"
         Height          =   195
         Left            =   9450
         TabIndex        =   113
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "进程阻滞因素描述"
         Height          =   195
         Left            =   12390
         TabIndex        =   112
         Top             =   1320
         Width           =   1905
      End
      Begin VB.Label Label16 
         Caption         =   "预计完成时间"
         Height          =   195
         Left            =   12690
         TabIndex        =   110
         Top             =   2700
         Width           =   1515
      End
      Begin VB.Label Label17 
         Caption         =   "处理及结果："
         Height          =   225
         Left            =   120
         TabIndex        =   109
         Top             =   4020
         Width           =   1245
      End
      Begin VB.Label Label18 
         Caption         =   "，调试记录如下"
         Height          =   225
         Left            =   4710
         TabIndex        =   108
         Top             =   4050
         Width           =   2655
      End
      Begin VB.Label Label19 
         Caption         =   "负载百分比"
         Height          =   165
         Left            =   150
         TabIndex        =   107
         Top             =   6255
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "电流"
         Height          =   165
         Left            =   150
         TabIndex        =   106
         Top             =   5970
         Width           =   1215
      End
      Begin VB.Label Label21 
         Caption         =   "油压或油压差"
         Height          =   165
         Left            =   150
         TabIndex        =   105
         Top             =   5700
         Width           =   1215
      End
      Begin VB.Label Label22 
         Caption         =   "排气温度"
         Height          =   165
         Left            =   150
         TabIndex        =   104
         Top             =   4860
         Width           =   1215
      End
      Begin VB.Label Label23 
         Caption         =   "吸气压力"
         Height          =   165
         Left            =   150
         TabIndex        =   103
         Top             =   4590
         Width           =   1215
      End
      Begin VB.Label Label24 
         Caption         =   "油温"
         Height          =   165
         Left            =   150
         TabIndex        =   102
         Top             =   5415
         Width           =   1215
      End
      Begin VB.Label Label25 
         Caption         =   "排气温度"
         Height          =   165
         Left            =   150
         TabIndex        =   101
         Top             =   5145
         Width           =   1215
      End
      Begin VB.Label Label26 
         Caption         =   "压缩机绝缘"
         Height          =   165
         Left            =   150
         TabIndex        =   100
         Top             =   6525
         Width           =   1215
      End
      Begin VB.Label Label27 
         Caption         =   "压缩机绕组"
         Height          =   165
         Left            =   150
         TabIndex        =   99
         Top             =   6810
         Width           =   1215
      End
      Begin VB.Label Label34 
         Caption         =   $"NewGzd6.frx":003B
         Height          =   165
         Index           =   0
         Left            =   9300
         TabIndex        =   92
         Top             =   4590
         Width           =   1275
      End
      Begin VB.Label Label35 
         Caption         =   "冷凝温度"
         Height          =   165
         Index           =   0
         Left            =   9300
         TabIndex        =   91
         Top             =   4860
         Width           =   1275
      End
      Begin VB.Label Label36 
         Caption         =   "冷却出水温度"
         Height          =   165
         Index           =   0
         Left            =   9300
         TabIndex        =   90
         Top             =   5145
         Width           =   1275
      End
      Begin VB.Label Label37 
         Caption         =   "冷却进水温度"
         Height          =   165
         Index           =   0
         Left            =   9300
         TabIndex        =   89
         Top             =   5415
         Width           =   1275
      End
      Begin VB.Label Label38 
         Caption         =   "冷冻进水温度"
         Height          =   165
         Index           =   0
         Left            =   9300
         TabIndex        =   88
         Top             =   5700
         Width           =   1275
      End
      Begin VB.Label Label39 
         Caption         =   "冷冻出水温度"
         Height          =   165
         Index           =   0
         Left            =   9300
         TabIndex        =   87
         Top             =   5970
         Width           =   1275
      End
      Begin VB.Label Label40 
         Caption         =   "电压"
         Height          =   165
         Left            =   9300
         TabIndex        =   86
         Top             =   6255
         Width           =   1275
      End
      Begin VB.Label Label43 
         Caption         =   "无此项"
         Height          =   165
         Left            =   13890
         TabIndex        =   83
         Top             =   4380
         Width           =   765
      End
      Begin VB.Label Label44 
         Caption         =   "风机编号"
         Height          =   225
         Left            =   150
         TabIndex        =   82
         Top             =   7470
         Width           =   915
      End
      Begin VB.Label Label45 
         Caption         =   "正常值"
         Height          =   225
         Left            =   13890
         TabIndex        =   81
         Top             =   7470
         Width           =   705
      End
      Begin VB.Label Label46 
         Caption         =   "电流A"
         Height          =   195
         Left            =   150
         TabIndex        =   80
         Top             =   7785
         Width           =   945
      End
      Begin VB.Label Label47 
         Caption         =   "绝缘阻值MΩ"
         Height          =   195
         Left            =   150
         TabIndex        =   79
         Top             =   8070
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   8325
         Left            =   30
         Top             =   30
         Width           =   14895
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   60
         X2              =   14955
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   60
         X2              =   14955
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Line Line2 
         Index           =   2
         X1              =   60
         X2              =   14955
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Line Line2 
         Index           =   3
         X1              =   60
         X2              =   14955
         Y1              =   1230
         Y2              =   1230
      End
      Begin VB.Line Line2 
         Index           =   4
         X1              =   60
         X2              =   14955
         Y1              =   1530
         Y2              =   1530
      End
      Begin VB.Line Line2 
         Index           =   5
         X1              =   60
         X2              =   14955
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line2 
         Index           =   6
         X1              =   60
         X2              =   14955
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Line Line2 
         Index           =   7
         X1              =   60
         X2              =   14955
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line2 
         Index           =   8
         X1              =   60
         X2              =   14955
         Y1              =   4830
         Y2              =   4830
      End
      Begin VB.Line Line2 
         Index           =   9
         X1              =   60
         X2              =   14955
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Line Line2 
         Index           =   10
         X1              =   60
         X2              =   14955
         Y1              =   5070
         Y2              =   5070
      End
      Begin VB.Line Line2 
         Index           =   11
         X1              =   60
         X2              =   14955
         Y1              =   5610
         Y2              =   5610
      End
      Begin VB.Line Line2 
         Index           =   12
         X1              =   60
         X2              =   14955
         Y1              =   5340
         Y2              =   5340
      End
      Begin VB.Line Line2 
         Index           =   13
         X1              =   60
         X2              =   14955
         Y1              =   5940
         Y2              =   5940
      End
      Begin VB.Line Line2 
         Index           =   14
         X1              =   60
         X2              =   14955
         Y1              =   6180
         Y2              =   6180
      End
      Begin VB.Line Line2 
         Index           =   15
         X1              =   60
         X2              =   14955
         Y1              =   6720
         Y2              =   6720
      End
      Begin VB.Line Line2 
         Index           =   16
         X1              =   60
         X2              =   14955
         Y1              =   6450
         Y2              =   6450
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         Index           =   17
         X1              =   60
         X2              =   14955
         Y1              =   7050
         Y2              =   7050
      End
      Begin VB.Line Line2 
         Index           =   18
         X1              =   60
         X2              =   14955
         Y1              =   7410
         Y2              =   7410
      End
      Begin VB.Line Line2 
         Index           =   19
         X1              =   60
         X2              =   14955
         Y1              =   8010
         Y2              =   8010
      End
      Begin VB.Line Line2 
         Index           =   20
         X1              =   60
         X2              =   14955
         Y1              =   7710
         Y2              =   7710
      End
      Begin VB.Line Line3 
         Index           =   0
         X1              =   1740
         X2              =   1740
         Y1              =   1230
         Y2              =   30
      End
      Begin VB.Line Line3 
         Index           =   1
         X1              =   3600
         X2              =   3600
         Y1              =   1230
         Y2              =   30
      End
      Begin VB.Line Line3 
         Index           =   2
         X1              =   5490
         X2              =   5490
         Y1              =   1230
         Y2              =   30
      End
      Begin VB.Line Line3 
         Index           =   3
         X1              =   7230
         X2              =   7230
         Y1              =   1230
         Y2              =   30
      End
      Begin VB.Line Line3 
         Index           =   4
         X1              =   9180
         X2              =   9180
         Y1              =   7020
         Y2              =   30
      End
      Begin VB.Line Line3 
         Index           =   5
         X1              =   12000
         X2              =   12000
         Y1              =   1230
         Y2              =   30
      End
      Begin VB.Line Line4 
         X1              =   12000
         X2              =   12000
         Y1              =   1530
         Y2              =   2640
      End
      Begin VB.Line Line2 
         Index           =   21
         X1              =   60
         X2              =   14955
         Y1              =   3990
         Y2              =   3990
      End
      Begin VB.Line Line5 
         X1              =   13020
         X2              =   13020
         Y1              =   2940
         Y2              =   3990
      End
      Begin VB.Line Line6 
         Index           =   0
         X1              =   1470
         X2              =   1470
         Y1              =   4320
         Y2              =   7050
      End
      Begin VB.Line Line7 
         Index           =   0
         X1              =   2790
         X2              =   2790
         Y1              =   4320
         Y2              =   7065
      End
      Begin VB.Line Line8 
         X1              =   4020
         X2              =   4020
         Y1              =   4320
         Y2              =   7050
      End
      Begin VB.Line Line9 
         X1              =   5340
         X2              =   5340
         Y1              =   4320
         Y2              =   7065
      End
      Begin VB.Line Line6 
         Index           =   1
         X1              =   6630
         X2              =   6630
         Y1              =   4320
         Y2              =   7050
      End
      Begin VB.Line Line7 
         Index           =   1
         X1              =   7950
         X2              =   7950
         Y1              =   4320
         Y2              =   7065
      End
      Begin VB.Line Line6 
         Index           =   2
         X1              =   10710
         X2              =   10710
         Y1              =   4320
         Y2              =   7050
      End
      Begin VB.Line Line7 
         Index           =   2
         X1              =   12240
         X2              =   12240
         Y1              =   4290
         Y2              =   7035
      End
      Begin VB.Line Line10 
         X1              =   13740
         X2              =   13740
         Y1              =   4320
         Y2              =   7065
      End
      Begin VB.Line Line11 
         X1              =   1470
         X2              =   90
         Y1              =   4560
         Y2              =   4320
      End
      Begin VB.Line Line12 
         X1              =   10740
         X2              =   9180
         Y1              =   4560
         Y2              =   4320
      End
      Begin VB.Line Line13 
         Index           =   0
         X1              =   1320
         X2              =   1320
         Y1              =   8340
         Y2              =   7410
      End
      Begin VB.Line Line14 
         Index           =   0
         X1              =   2430
         X2              =   2430
         Y1              =   7410
         Y2              =   8355
      End
      Begin VB.Line Line15 
         Index           =   0
         X1              =   3450
         X2              =   3450
         Y1              =   7410
         Y2              =   8340
      End
      Begin VB.Line Line13 
         Index           =   1
         X1              =   4470
         X2              =   4470
         Y1              =   8340
         Y2              =   7410
      End
      Begin VB.Line Line14 
         Index           =   1
         X1              =   5580
         X2              =   5580
         Y1              =   7410
         Y2              =   8355
      End
      Begin VB.Line Line15 
         Index           =   1
         X1              =   6600
         X2              =   6600
         Y1              =   7410
         Y2              =   8340
      End
      Begin VB.Line Line13 
         Index           =   2
         X1              =   7620
         X2              =   7620
         Y1              =   8340
         Y2              =   7410
      End
      Begin VB.Line Line14 
         Index           =   2
         X1              =   8730
         X2              =   8730
         Y1              =   7410
         Y2              =   8355
      End
      Begin VB.Line Line15 
         Index           =   2
         X1              =   9750
         X2              =   9750
         Y1              =   7410
         Y2              =   8340
      End
      Begin VB.Line Line13 
         Index           =   3
         X1              =   10770
         X2              =   10770
         Y1              =   8340
         Y2              =   7410
      End
      Begin VB.Line Line14 
         Index           =   3
         X1              =   11790
         X2              =   11790
         Y1              =   7410
         Y2              =   8355
      End
      Begin VB.Line Line15 
         Index           =   3
         X1              =   12840
         X2              =   12840
         Y1              =   7410
         Y2              =   8340
      End
      Begin VB.Line Line16 
         X1              =   13830
         X2              =   13830
         Y1              =   7410
         Y2              =   8340
      End
      Begin VB.Shape Shape2 
         Height          =   2415
         Left            =   -75000
         Top             =   0
         Width           =   14985
      End
      Begin VB.Label Label38 
         Caption         =   "日期："
         Height          =   195
         Index           =   1
         Left            =   -62040
         TabIndex        =   46
         Top             =   1830
         Width           =   945
      End
      Begin VB.Line Line37 
         X1              =   -62100
         X2              =   -62100
         Y1              =   1200
         Y2              =   2400
      End
      Begin VB.Line Line36 
         X1              =   -64200
         X2              =   -64200
         Y1              =   1200
         Y2              =   2400
      End
      Begin VB.Line Line35 
         X1              =   -75000
         X2              =   -60060
         Y1              =   1770
         Y2              =   1770
      End
      Begin VB.Line Line34 
         X1              =   -75000
         X2              =   -60060
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label37 
         Caption         =   "质量控制签名："
         Height          =   195
         Index           =   1
         Left            =   -62040
         TabIndex        =   45
         Top             =   1260
         Width           =   1275
      End
      Begin VB.Label Label36 
         Caption         =   "日期："
         Height          =   195
         Index           =   1
         Left            =   -64110
         TabIndex        =   44
         Top             =   1830
         Width           =   945
      End
      Begin VB.Label Label35 
         Caption         =   "客户签名："
         Height          =   225
         Index           =   1
         Left            =   -64110
         TabIndex        =   43
         Top             =   1230
         Width           =   945
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
         Index           =   1
         Left            =   -74910
         TabIndex        =   42
         Top             =   1890
         Width           =   885
      End
      Begin VB.Label Label32 
         Caption         =   "加班工时"
         Height          =   165
         Index           =   1
         Left            =   -66330
         TabIndex        =   41
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label Label31 
         Caption         =   "旅途时间"
         Height          =   165
         Index           =   1
         Left            =   -68760
         TabIndex        =   40
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label Label30 
         Caption         =   "完成时间"
         Height          =   165
         Index           =   1
         Left            =   -71880
         TabIndex        =   39
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label Label29 
         Caption         =   "到达时间"
         Height          =   165
         Index           =   1
         Left            =   -74850
         TabIndex        =   38
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Line Line33 
         X1              =   -75000
         X2              =   -60060
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line25 
         X1              =   -75000
         X2              =   -59940
         Y1              =   630
         Y2              =   630
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
         Index           =   1
         Left            =   -74880
         TabIndex        =   37
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label48 
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
         TabIndex        =   36
         Top             =   90
         Width           =   1125
      End
      Begin VB.Label Label15 
         Caption         =   "所需条件、工具或配件"
         Height          =   255
         Left            =   9510
         TabIndex        =   111
         Top             =   2700
         Width           =   2115
      End
      Begin VB.Label Label28 
         Caption         =   "1#压缩机"
         Height          =   195
         Index           =   0
         Left            =   1710
         TabIndex        =   98
         Top             =   4380
         Width           =   915
      End
      Begin VB.Label Label29 
         Caption         =   "2#压缩机"
         Height          =   195
         Index           =   0
         Left            =   2970
         TabIndex        =   97
         Top             =   4380
         Width           =   915
      End
      Begin VB.Label Label30 
         Caption         =   "3#压缩机"
         Height          =   195
         Index           =   0
         Left            =   4245
         TabIndex        =   96
         Top             =   4380
         Width           =   915
      End
      Begin VB.Label Label31 
         Caption         =   "4#压缩机"
         Height          =   195
         Index           =   0
         Left            =   5505
         TabIndex        =   95
         Top             =   4380
         Width           =   915
      End
      Begin VB.Label Label32 
         Caption         =   "正常值"
         Height          =   195
         Index           =   0
         Left            =   6780
         TabIndex        =   94
         Top             =   4380
         Width           =   915
      End
      Begin VB.Label Label33 
         Caption         =   "无此项"
         Height          =   195
         Index           =   0
         Left            =   8190
         TabIndex        =   93
         Top             =   4380
         Width           =   915
      End
      Begin VB.Label Label41 
         Caption         =   "机组参数"
         Height          =   225
         Left            =   10920
         TabIndex        =   85
         Top             =   4380
         Width           =   1245
      End
      Begin VB.Label Label42 
         Caption         =   "正常值"
         Height          =   225
         Left            =   12540
         TabIndex        =   84
         Top             =   4380
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "1#压缩机"
         Height          =   165
         Left            =   1950
         TabIndex        =   125
         Top             =   30
         Width           =   1755
      End
      Begin VB.Label Label2 
         Caption         =   "2#压缩机"
         Height          =   195
         Left            =   3900
         TabIndex        =   124
         Top             =   30
         Width           =   1425
      End
      Begin VB.Label Label3 
         Caption         =   "3#压缩机"
         Height          =   195
         Left            =   5670
         TabIndex        =   123
         Top             =   30
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "4#压缩机"
         Height          =   195
         Left            =   7530
         TabIndex        =   122
         Top             =   30
         Width           =   1515
      End
      Begin VB.Label Label8 
         Caption         =   "冷媒种类及充注量"
         Height          =   225
         Left            =   9600
         TabIndex        =   118
         Top             =   360
         Width           =   2025
      End
      Begin VB.Label Label9 
         Caption         =   "冷冻油种类及充注量"
         Height          =   225
         Left            =   9450
         TabIndex        =   117
         Top             =   690
         Width           =   2355
      End
   End
   Begin MSDataGridLib.DataGrid comHtbh 
      Height          =   1155
      Left            =   5730
      TabIndex        =   22
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
      Left            =   1590
      TabIndex        =   21
      Top             =   810
      Width           =   4125
   End
   Begin VB.CommandButton cmdBack 
      Height          =   375
      Left            =   14520
      Picture         =   "NewGzd6.frx":004B
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "返回"
      Top             =   10410
      Width           =   465
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   14040
      Picture         =   "NewGzd6.frx":014D
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "保存"
      Top             =   10410
      Width           =   465
   End
   Begin VB.CommandButton cmdMod 
      Height          =   375
      Left            =   13560
      Picture         =   "NewGzd6.frx":07B7
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "修改"
      Top             =   10410
      Width           =   465
   End
   Begin VB.CommandButton cmdQm 
      Caption         =   "cmdQm"
      Height          =   345
      Index           =   0
      Left            =   9900
      TabIndex        =   12
      Top             =   10410
      Width           =   945
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
      Left            =   7710
      TabIndex        =   11
      Text            =   "的"
      Top             =   630
      Width           =   4245
   End
   Begin VB.TextBox BA 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dddddd aaaa"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
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
      Index           =   6
      Left            =   7710
      Locked          =   -1  'True
      TabIndex        =   10
      Tag             =   "20"
      Top             =   1440
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
      Left            =   7710
      TabIndex        =   9
      Top             =   1020
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
      Left            =   7710
      TabIndex        =   8
      Text            =   "的"
      Top             =   210
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
      Left            =   1620
      TabIndex        =   7
      Top             =   990
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
      Left            =   1620
      TabIndex        =   6
      Top             =   570
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
      Left            =   12450
      TabIndex        =   5
      Top             =   1110
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
      Left            =   1620
      TabIndex        =   4
      Top             =   150
      Width           =   4065
   End
   Begin VB.TextBox TA 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   13260
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CheckBox C1 
      Alignment       =   1  'Right Justify
      Caption         =   "1号"
      Height          =   285
      Index           =   0
      Left            =   12330
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
      Height          =   1515
      Left            =   6180
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "NewGzd6.frx":0AC1
      Top             =   210
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
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "NewGzd6.frx":0AFA
      Top             =   150
      Width           =   1365
   End
   Begin MSComCtl2.DTPicker dtpA 
      Height          =   195
      Left            =   7710
      TabIndex        =   18
      Top             =   1440
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   344
      _Version        =   393216
      Format          =   149553153
      CurrentDate     =   38897
   End
   Begin VB.Label LBLKjj 
      Caption         =   $"NewGzd6.frx":0B31
      ForeColor       =   &H8000000D&
      Height          =   1305
      Left            =   12750
      TabIndex        =   246
      Top             =   390
      Width           =   2835
   End
   Begin VB.Line Line38 
      X1              =   1590
      X2              =   5745
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Label Label39 
      Caption         =   "NO:"
      Height          =   255
      Index           =   1
      Left            =   12450
      TabIndex        =   128
      Top             =   180
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
      Left            =   13020
      TabIndex        =   127
      Top             =   180
      Width           =   1605
   End
   Begin VB.Label lblkhdh 
      Caption         =   "lblkhdh"
      Height          =   225
      Left            =   12270
      TabIndex        =   20
      Top             =   150
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label lblGid 
      Caption         =   "lblGid"
      Height          =   225
      Left            =   13380
      TabIndex        =   19
      Top             =   150
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
      TabIndex        =   17
      Top             =   10470
      Width           =   1905
   End
   Begin VB.Label lblQM 
      Caption         =   "签字提交"
      Height          =   225
      Index           =   0
      Left            =   8970
      TabIndex        =   16
      Top             =   10470
      Width           =   795
   End
   Begin VB.Line Line1 
      X1              =   7710
      X2              =   11805
      Y1              =   870
      Y2              =   870
   End
   Begin VB.Line Line32 
      X1              =   7710
      X2              =   11790
      Y1              =   1650
      Y2              =   1650
   End
   Begin VB.Line Line31 
      X1              =   7710
      X2              =   11790
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Line Line30 
      X1              =   7710
      X2              =   11790
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Line Line28 
      X1              =   1620
      X2              =   5700
      Y1              =   1215
      Y2              =   1215
   End
   Begin VB.Line Line27 
      X1              =   1620
      X2              =   5700
      Y1              =   795
      Y2              =   795
   End
   Begin VB.Line Line26 
      X1              =   1620
      X2              =   5715
      Y1              =   390
      Y2              =   390
   End
End
Attribute VB_Name = "NewGzd6"
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

Private Sub cmdA1_Click()
Dim oo As Integer
If C1(12).Value = 1 Then
    For oo = 12 To 18
        C1(oo).Value = 0
    Next
Else
    For oo = 12 To 18
        C1(oo).Value = 1
    Next
End If
End Sub

Private Sub cmdAll_Click()
Dim oo As Integer
If C1(3).Value = 1 Then
    For oo = 3 To 11
        C1(oo).Value = 0
    Next
Else
    For oo = 3 To 11
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
For oo = 1 To 108
    mod1.HTP.Update "mat" & oo, TA(oo).Text
Next
For oo = 1 To 32
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
Me.Width = 15135
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
BA(15).Tag = 200
BA(16).Tag = 50
BA(17).Tag = 50
BA(8).Tag = 10
BA(9).Tag = 10
BA(10).Tag = 10
BA(11).Tag = 10
For oo = 1 To 108
    TA(oo).Tag = 50
Next
TA(106).Tag = 200
TA(15).Tag = 200
TA(18).Tag = 200
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

