VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form NewGzd2 
   Caption         =   "热泵机组巡视检修工作报告（单）"
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
      Left            =   1830
      TabIndex        =   162
      Top             =   1290
      Width           =   4065
   End
   Begin MSDataGridLib.DataGrid dtgRen 
      Height          =   8085
      Left            =   6330
      TabIndex        =   159
      Top             =   1680
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
      Height          =   8715
      Left            =   0
      TabIndex        =   23
      Top             =   1560
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   15372
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "内容1"
      TabPicture(0)   =   "NewGzd2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line50"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line49"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line48"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line47"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line46"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line45"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Line44"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line43"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line42"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line41"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Line40"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line39"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label41"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label40"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Line22"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Line21"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label24"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label22"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label23"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Line20"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Line19"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label21"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label20"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label2"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label3"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label4"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label5"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label6"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label7"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label8"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label9"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label10"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label11"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label12"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label13"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label14"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Label15"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Label16"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Label17"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Label18"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Label19"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Shape1"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Line2"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Line3"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Line4"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Line5"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Line6"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Line7"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Line8"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Line9"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "Line10"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Line11"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Line12"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Line13"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Line14"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "Line15"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "Line16"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "Line17"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "Line18"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "Label1"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "Line1"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "TA(73)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "TA(72)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "TA(71)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "TA(70)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "TA(69)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "TA(68)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "TA(67)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "TA(66)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "TA(65)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "TA(64)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "TA(63)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "TA(62)"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "TA(61)"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "C1(80)"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "C1(79)"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "C1(78)"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "C1(77)"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "C1(76)"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "C1(75)"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "C1(74)"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "C1(73)"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "C1(72)"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "C1(71)"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).Control(84)=   "C1(70)"
      Tab(0).Control(84).Enabled=   0   'False
      Tab(0).Control(85)=   "C1(69)"
      Tab(0).Control(85).Enabled=   0   'False
      Tab(0).Control(86)=   "C1(68)"
      Tab(0).Control(86).Enabled=   0   'False
      Tab(0).Control(87)=   "C1(60)"
      Tab(0).Control(87).Enabled=   0   'False
      Tab(0).Control(88)=   "C1(59)"
      Tab(0).Control(88).Enabled=   0   'False
      Tab(0).Control(89)=   "C1(58)"
      Tab(0).Control(89).Enabled=   0   'False
      Tab(0).Control(90)=   "C1(57)"
      Tab(0).Control(90).Enabled=   0   'False
      Tab(0).Control(91)=   "C1(56)"
      Tab(0).Control(91).Enabled=   0   'False
      Tab(0).Control(92)=   "C1(55)"
      Tab(0).Control(92).Enabled=   0   'False
      Tab(0).Control(93)=   "C1(54)"
      Tab(0).Control(93).Enabled=   0   'False
      Tab(0).Control(94)=   "TA(53)"
      Tab(0).Control(94).Enabled=   0   'False
      Tab(0).Control(95)=   "TA(52)"
      Tab(0).Control(95).Enabled=   0   'False
      Tab(0).Control(96)=   "C1(51)"
      Tab(0).Control(96).Enabled=   0   'False
      Tab(0).Control(97)=   "C1(50)"
      Tab(0).Control(97).Enabled=   0   'False
      Tab(0).Control(98)=   "C1(49)"
      Tab(0).Control(98).Enabled=   0   'False
      Tab(0).Control(99)=   "C1(48)"
      Tab(0).Control(99).Enabled=   0   'False
      Tab(0).Control(100)=   "C1(47)"
      Tab(0).Control(100).Enabled=   0   'False
      Tab(0).Control(101)=   "C1(46)"
      Tab(0).Control(101).Enabled=   0   'False
      Tab(0).Control(102)=   "C1(45)"
      Tab(0).Control(102).Enabled=   0   'False
      Tab(0).Control(103)=   "C1(44)"
      Tab(0).Control(103).Enabled=   0   'False
      Tab(0).Control(104)=   "C1(43)"
      Tab(0).Control(104).Enabled=   0   'False
      Tab(0).Control(105)=   "C1(42)"
      Tab(0).Control(105).Enabled=   0   'False
      Tab(0).Control(106)=   "C1(37)"
      Tab(0).Control(106).Enabled=   0   'False
      Tab(0).Control(107)=   "C1(36)"
      Tab(0).Control(107).Enabled=   0   'False
      Tab(0).Control(108)=   "C1(35)"
      Tab(0).Control(108).Enabled=   0   'False
      Tab(0).Control(109)=   "C1(34)"
      Tab(0).Control(109).Enabled=   0   'False
      Tab(0).Control(110)=   "C1(33)"
      Tab(0).Control(110).Enabled=   0   'False
      Tab(0).Control(111)=   "C1(32)"
      Tab(0).Control(111).Enabled=   0   'False
      Tab(0).Control(112)=   "C1(31)"
      Tab(0).Control(112).Enabled=   0   'False
      Tab(0).Control(113)=   "C1(30)"
      Tab(0).Control(113).Enabled=   0   'False
      Tab(0).Control(114)=   "C1(29)"
      Tab(0).Control(114).Enabled=   0   'False
      Tab(0).Control(115)=   "C1(28)"
      Tab(0).Control(115).Enabled=   0   'False
      Tab(0).Control(116)=   "C1(23)"
      Tab(0).Control(116).Enabled=   0   'False
      Tab(0).Control(117)=   "C1(22)"
      Tab(0).Control(117).Enabled=   0   'False
      Tab(0).Control(118)=   "C1(21)"
      Tab(0).Control(118).Enabled=   0   'False
      Tab(0).Control(119)=   "C1(20)"
      Tab(0).Control(119).Enabled=   0   'False
      Tab(0).Control(120)=   "C1(19)"
      Tab(0).Control(120).Enabled=   0   'False
      Tab(0).Control(121)=   "C1(18)"
      Tab(0).Control(121).Enabled=   0   'False
      Tab(0).Control(122)=   "C1(17)"
      Tab(0).Control(122).Enabled=   0   'False
      Tab(0).Control(123)=   "C1(16)"
      Tab(0).Control(123).Enabled=   0   'False
      Tab(0).Control(124)=   "C1(15)"
      Tab(0).Control(124).Enabled=   0   'False
      Tab(0).Control(125)=   "C1(14)"
      Tab(0).Control(125).Enabled=   0   'False
      Tab(0).Control(126)=   "C1(13)"
      Tab(0).Control(126).Enabled=   0   'False
      Tab(0).Control(127)=   "C1(1)"
      Tab(0).Control(127).Enabled=   0   'False
      Tab(0).Control(128)=   "C1(2)"
      Tab(0).Control(128).Enabled=   0   'False
      Tab(0).Control(129)=   "C1(3)"
      Tab(0).Control(129).Enabled=   0   'False
      Tab(0).Control(130)=   "C1(4)"
      Tab(0).Control(130).Enabled=   0   'False
      Tab(0).Control(131)=   "TA(1)"
      Tab(0).Control(131).Enabled=   0   'False
      Tab(0).Control(132)=   "TA(2)"
      Tab(0).Control(132).Enabled=   0   'False
      Tab(0).Control(133)=   "TA(3)"
      Tab(0).Control(133).Enabled=   0   'False
      Tab(0).Control(134)=   "TA(4)"
      Tab(0).Control(134).Enabled=   0   'False
      Tab(0).Control(135)=   "TA(5)"
      Tab(0).Control(135).Enabled=   0   'False
      Tab(0).Control(136)=   "TA(6)"
      Tab(0).Control(136).Enabled=   0   'False
      Tab(0).Control(137)=   "TA(7)"
      Tab(0).Control(137).Enabled=   0   'False
      Tab(0).Control(138)=   "TA(8)"
      Tab(0).Control(138).Enabled=   0   'False
      Tab(0).Control(139)=   "TA(9)"
      Tab(0).Control(139).Enabled=   0   'False
      Tab(0).Control(140)=   "TA(10)"
      Tab(0).Control(140).Enabled=   0   'False
      Tab(0).Control(141)=   "TA(11)"
      Tab(0).Control(141).Enabled=   0   'False
      Tab(0).Control(142)=   "TA(12)"
      Tab(0).Control(142).Enabled=   0   'False
      Tab(0).Control(143)=   "TA(13)"
      Tab(0).Control(143).Enabled=   0   'False
      Tab(0).Control(144)=   "TA(14)"
      Tab(0).Control(144).Enabled=   0   'False
      Tab(0).Control(145)=   "TA(15)"
      Tab(0).Control(145).Enabled=   0   'False
      Tab(0).Control(146)=   "TA(16)"
      Tab(0).Control(146).Enabled=   0   'False
      Tab(0).Control(147)=   "TA(17)"
      Tab(0).Control(147).Enabled=   0   'False
      Tab(0).Control(148)=   "TA(18)"
      Tab(0).Control(148).Enabled=   0   'False
      Tab(0).Control(149)=   "TA(19)"
      Tab(0).Control(149).Enabled=   0   'False
      Tab(0).Control(150)=   "TA(20)"
      Tab(0).Control(150).Enabled=   0   'False
      Tab(0).Control(151)=   "TA(21)"
      Tab(0).Control(151).Enabled=   0   'False
      Tab(0).Control(152)=   "TA(22)"
      Tab(0).Control(152).Enabled=   0   'False
      Tab(0).Control(153)=   "TA(23)"
      Tab(0).Control(153).Enabled=   0   'False
      Tab(0).Control(154)=   "TA(24)"
      Tab(0).Control(154).Enabled=   0   'False
      Tab(0).Control(155)=   "TA(25)"
      Tab(0).Control(155).Enabled=   0   'False
      Tab(0).Control(156)=   "TA(26)"
      Tab(0).Control(156).Enabled=   0   'False
      Tab(0).Control(157)=   "TA(27)"
      Tab(0).Control(157).Enabled=   0   'False
      Tab(0).Control(158)=   "TA(28)"
      Tab(0).Control(158).Enabled=   0   'False
      Tab(0).Control(159)=   "TA(29)"
      Tab(0).Control(159).Enabled=   0   'False
      Tab(0).Control(160)=   "TA(30)"
      Tab(0).Control(160).Enabled=   0   'False
      Tab(0).Control(161)=   "TA(31)"
      Tab(0).Control(161).Enabled=   0   'False
      Tab(0).Control(162)=   "TA(32)"
      Tab(0).Control(162).Enabled=   0   'False
      Tab(0).Control(163)=   "TA(33)"
      Tab(0).Control(163).Enabled=   0   'False
      Tab(0).Control(164)=   "TA(34)"
      Tab(0).Control(164).Enabled=   0   'False
      Tab(0).Control(165)=   "TA(35)"
      Tab(0).Control(165).Enabled=   0   'False
      Tab(0).Control(166)=   "TA(36)"
      Tab(0).Control(166).Enabled=   0   'False
      Tab(0).Control(167)=   "TA(37)"
      Tab(0).Control(167).Enabled=   0   'False
      Tab(0).Control(168)=   "TA(38)"
      Tab(0).Control(168).Enabled=   0   'False
      Tab(0).Control(169)=   "TA(39)"
      Tab(0).Control(169).Enabled=   0   'False
      Tab(0).Control(170)=   "TA(40)"
      Tab(0).Control(170).Enabled=   0   'False
      Tab(0).Control(171)=   "C1(5)"
      Tab(0).Control(171).Enabled=   0   'False
      Tab(0).Control(172)=   "C1(6)"
      Tab(0).Control(172).Enabled=   0   'False
      Tab(0).Control(173)=   "C1(7)"
      Tab(0).Control(173).Enabled=   0   'False
      Tab(0).Control(174)=   "C1(8)"
      Tab(0).Control(174).Enabled=   0   'False
      Tab(0).Control(175)=   "C1(9)"
      Tab(0).Control(175).Enabled=   0   'False
      Tab(0).Control(176)=   "C1(10)"
      Tab(0).Control(176).Enabled=   0   'False
      Tab(0).Control(177)=   "C1(11)"
      Tab(0).Control(177).Enabled=   0   'False
      Tab(0).Control(178)=   "C1(12)"
      Tab(0).Control(178).Enabled=   0   'False
      Tab(0).Control(179)=   "TA(41)"
      Tab(0).Control(179).Enabled=   0   'False
      Tab(0).Control(180)=   "TA(42)"
      Tab(0).Control(180).Enabled=   0   'False
      Tab(0).Control(181)=   "TA(43)"
      Tab(0).Control(181).Enabled=   0   'False
      Tab(0).Control(182)=   "TA(44)"
      Tab(0).Control(182).Enabled=   0   'False
      Tab(0).Control(183)=   "TA(45)"
      Tab(0).Control(183).Enabled=   0   'False
      Tab(0).Control(184)=   "TA(46)"
      Tab(0).Control(184).Enabled=   0   'False
      Tab(0).Control(185)=   "TA(47)"
      Tab(0).Control(185).Enabled=   0   'False
      Tab(0).Control(186)=   "TA(48)"
      Tab(0).Control(186).Enabled=   0   'False
      Tab(0).Control(187)=   "TA(49)"
      Tab(0).Control(187).Enabled=   0   'False
      Tab(0).Control(188)=   "TA(50)"
      Tab(0).Control(188).Enabled=   0   'False
      Tab(0).Control(189)=   "TA(51)"
      Tab(0).Control(189).Enabled=   0   'False
      Tab(0).Control(190)=   "cmdAll"
      Tab(0).Control(190).Enabled=   0   'False
      Tab(0).ControlCount=   191
      TabCaption(1)   =   "内容2"
      TabPicture(1)   =   "NewGzd2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TA(56)"
      Tab(1).Control(1)=   "C1(67)"
      Tab(1).Control(2)=   "C1(66)"
      Tab(1).Control(3)=   "TA(55)"
      Tab(1).Control(4)=   "C1(63)"
      Tab(1).Control(5)=   "C1(62)"
      Tab(1).Control(6)=   "C1(61)"
      Tab(1).Control(7)=   "TA(54)"
      Tab(1).Control(8)=   "C1(53)"
      Tab(1).Control(9)=   "C1(52)"
      Tab(1).Control(10)=   "C1(41)"
      Tab(1).Control(11)=   "C1(40)"
      Tab(1).Control(12)=   "C1(39)"
      Tab(1).Control(13)=   "C1(38)"
      Tab(1).Control(14)=   "C1(27)"
      Tab(1).Control(15)=   "C1(26)"
      Tab(1).Control(16)=   "C1(25)"
      Tab(1).Control(17)=   "C1(24)"
      Tab(1).Control(18)=   "BA(14)"
      Tab(1).Control(19)=   "TA(58)"
      Tab(1).Control(20)=   "BA(16)"
      Tab(1).Control(21)=   "BA(15)"
      Tab(1).Control(22)=   "BA(13)"
      Tab(1).Control(23)=   "BA(12)"
      Tab(1).Control(24)=   "Frame1"
      Tab(1).Control(25)=   "BA(11)"
      Tab(1).Control(26)=   "BA(10)"
      Tab(1).Control(27)=   "BA(9)"
      Tab(1).Control(28)=   "BA(8)"
      Tab(1).Control(29)=   "TA(60)"
      Tab(1).Control(30)=   "Text3"
      Tab(1).Control(31)=   "C1(64)"
      Tab(1).Control(32)=   "C1(65)"
      Tab(1).Control(33)=   "TA(57)"
      Tab(1).Control(34)=   "dtpC"
      Tab(1).Control(35)=   "dtpB"
      Tab(1).Control(36)=   "TA(59)"
      Tab(1).Control(37)=   "Label39(0)"
      Tab(1).Control(38)=   "Line38(0)"
      Tab(1).Control(39)=   "Label25"
      Tab(1).Control(40)=   "Line24"
      Tab(1).Control(41)=   "Line23"
      Tab(1).Control(42)=   "Label42"
      Tab(1).Control(43)=   "Label38"
      Tab(1).Control(44)=   "Line37"
      Tab(1).Control(45)=   "Line36"
      Tab(1).Control(46)=   "Line35"
      Tab(1).Control(47)=   "Line34"
      Tab(1).Control(48)=   "Label37"
      Tab(1).Control(49)=   "Label36"
      Tab(1).Control(50)=   "Label35"
      Tab(1).Control(51)=   "Shape2"
      Tab(1).Control(52)=   "Label34"
      Tab(1).Control(53)=   "Label32"
      Tab(1).Control(54)=   "Label31"
      Tab(1).Control(55)=   "Label30"
      Tab(1).Control(56)=   "Label29"
      Tab(1).Control(57)=   "Line33"
      Tab(1).Control(58)=   "Line25"
      Tab(1).Control(59)=   "Line29"
      Tab(1).Control(60)=   "Label28"
      Tab(1).Control(61)=   "Label27"
      Tab(1).Control(62)=   "Label26"
      Tab(1).ControlCount=   63
      Begin VB.CommandButton cmdAll 
         Caption         =   "全部"
         Height          =   285
         Left            =   7770
         TabIndex        =   240
         Top             =   7560
         Width           =   915
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   56
         Left            =   -71130
         TabIndex        =   229
         Top             =   1530
         Width           =   11175
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "目测检漏"
         Height          =   180
         Index           =   67
         Left            =   -74940
         TabIndex        =   156
         Top             =   1545
         Width           =   1095
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "仪器检漏"
         Height          =   180
         Index           =   66
         Left            =   -73440
         TabIndex        =   155
         Top             =   1545
         Width           =   1065
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   55
         Left            =   -62760
         TabIndex        =   228
         Top             =   915
         Width           =   2685
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   285
         Index           =   63
         Left            =   -61620
         TabIndex        =   154
         Top             =   1245
         Width           =   1005
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   285
         Index           =   62
         Left            =   -63540
         TabIndex        =   153
         Top             =   1245
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   285
         Index           =   61
         Left            =   -61620
         TabIndex        =   152
         Top             =   525
         Width           =   1005
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   54
         Left            =   -62010
         TabIndex        =   227
         Top             =   255
         Width           =   1875
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   285
         Index           =   53
         Left            =   -63540
         TabIndex        =   151
         Top             =   510
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   285
         Index           =   52
         Left            =   -63540
         TabIndex        =   150
         Top             =   150
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已调整或检修"
         Height          =   285
         Index           =   41
         Left            =   -65700
         TabIndex        =   149
         Top             =   1245
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "因故未完成清洁"
         Height          =   285
         Index           =   40
         Left            =   -65700
         TabIndex        =   148
         Top             =   870
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已检修或更换"
         Height          =   285
         Index           =   39
         Left            =   -65700
         TabIndex        =   147
         Top             =   510
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已调整或检修"
         Height          =   285
         Index           =   38
         Left            =   -65700
         TabIndex        =   146
         Top             =   150
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   27
         Left            =   -67320
         TabIndex        =   145
         Top             =   1245
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   26
         Left            =   -67320
         TabIndex        =   144
         Top             =   870
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   25
         Left            =   -67320
         TabIndex        =   143
         Top             =   510
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   24
         Left            =   -67320
         TabIndex        =   142
         Top             =   150
         Width           =   855
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
         Index           =   14
         Left            =   -64080
         TabIndex        =   127
         Top             =   6690
         Width           =   1755
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   570
         Index           =   58
         Left            =   -73590
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   231
         Top             =   4680
         Width           =   13575
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
         TabIndex        =   125
         Top             =   6690
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
         Left            =   -62010
         TabIndex        =   239
         Top             =   6120
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
         Index           =   13
         Left            =   -64080
         TabIndex        =   238
         Top             =   6120
         Width           =   1935
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   570
         Index           =   12
         Left            =   -73530
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   237
         Text            =   "NewGzd2.frx":0038
         Top             =   6450
         Width           =   9345
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Left            =   -74940
         TabIndex        =   119
         Top             =   6120
         Width           =   10755
         Begin VB.OptionButton FPD 
            Caption         =   "尚可"
            Height          =   195
            Left            =   6150
            TabIndex        =   123
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton FPC 
            Caption         =   "较满意"
            Height          =   195
            Left            =   4550
            TabIndex        =   122
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton FPB 
            Caption         =   "满意"
            Height          =   195
            Left            =   2950
            TabIndex        =   121
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton FPA 
            Caption         =   "优秀"
            Height          =   195
            Left            =   1350
            TabIndex        =   120
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
            TabIndex        =   124
            Top             =   60
            Width           =   1395
         End
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   11
         Left            =   -65220
         TabIndex        =   236
         Text            =   "的"
         Top             =   5880
         Width           =   1035
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   10
         Left            =   -67620
         TabIndex        =   235
         Text            =   "的"
         Top             =   5880
         Width           =   1035
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   9
         Left            =   -70290
         TabIndex        =   234
         Text            =   "的"
         Top             =   5880
         Width           =   1035
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   8
         Left            =   -73530
         TabIndex        =   233
         Text            =   "的"
         Top             =   5880
         Width           =   1035
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   60
         Left            =   -61230
         TabIndex        =   118
         Top             =   5550
         Width           =   1065
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   -61860
         TabIndex        =   117
         Text            =   "复核人:"
         Top             =   5640
         Width           =   735
      End
      Begin VB.CheckBox C1 
         Caption         =   "未完成"
         Height          =   180
         Index           =   64
         Left            =   -61110
         TabIndex        =   116
         Top             =   4470
         Width           =   945
      End
      Begin VB.CheckBox C1 
         Caption         =   "完成"
         Height          =   180
         Index           =   65
         Left            =   -62250
         TabIndex        =   115
         Top             =   4470
         Width           =   1005
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   57
         Left            =   -73590
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   230
         Top             =   3930
         Width           =   13575
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   780
         Index           =   51
         Left            =   12120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   213
         Top             =   1890
         Width           =   2895
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   50
         Left            =   13680
         TabIndex        =   212
         Top             =   1635
         Width           =   1335
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   49
         Left            =   12180
         TabIndex        =   207
         Top             =   1635
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   48
         Left            =   13680
         TabIndex        =   211
         Top             =   1380
         Width           =   1335
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   47
         Left            =   12180
         TabIndex        =   206
         Top             =   1380
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   46
         Left            =   13680
         TabIndex        =   210
         Top             =   1125
         Width           =   1335
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   45
         Left            =   12180
         TabIndex        =   205
         Top             =   1125
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   44
         Left            =   13680
         TabIndex        =   209
         Top             =   855
         Width           =   1335
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   43
         Left            =   12180
         TabIndex        =   204
         Top             =   855
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   42
         Left            =   13680
         TabIndex        =   208
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   41
         Left            =   12180
         TabIndex        =   203
         Top             =   600
         Width           =   1395
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   12
         Left            =   9720
         TabIndex        =   88
         Top             =   2445
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   11
         Left            =   9720
         TabIndex        =   87
         Top             =   2175
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   10
         Left            =   9720
         TabIndex        =   86
         Top             =   1905
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   9
         Left            =   9720
         TabIndex        =   85
         Top             =   1635
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   8
         Left            =   9720
         TabIndex        =   84
         Top             =   1365
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   7
         Left            =   9720
         TabIndex        =   83
         Top             =   1110
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   6
         Left            =   9720
         TabIndex        =   82
         Top             =   840
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   5
         Left            =   9720
         TabIndex        =   81
         Top             =   570
         Width           =   405
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   40
         Left            =   7800
         TabIndex        =   202
         Top             =   2445
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   39
         Left            =   6255
         TabIndex        =   194
         Top             =   2445
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   38
         Left            =   4710
         TabIndex        =   186
         Top             =   2445
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   37
         Left            =   3180
         TabIndex        =   178
         Top             =   2445
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   36
         Left            =   1635
         TabIndex        =   170
         Top             =   2445
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   35
         Left            =   7800
         TabIndex        =   201
         Top             =   2175
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   34
         Left            =   6255
         TabIndex        =   193
         Top             =   2175
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   33
         Left            =   4710
         TabIndex        =   185
         Top             =   2175
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   32
         Left            =   3180
         TabIndex        =   177
         Top             =   2190
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   31
         Left            =   1635
         TabIndex        =   169
         Top             =   2190
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   30
         Left            =   7800
         TabIndex        =   200
         Top             =   1905
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   29
         Left            =   6255
         TabIndex        =   192
         Top             =   1905
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   28
         Left            =   4710
         TabIndex        =   184
         Top             =   1905
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   27
         Left            =   3180
         TabIndex        =   176
         Top             =   1905
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   26
         Left            =   1635
         TabIndex        =   168
         Top             =   1920
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   25
         Left            =   7800
         TabIndex        =   199
         Top             =   1635
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   24
         Left            =   6255
         TabIndex        =   191
         Top             =   1635
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   23
         Left            =   4710
         TabIndex        =   183
         Top             =   1635
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   22
         Left            =   3180
         TabIndex        =   175
         Top             =   1635
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   21
         Left            =   1635
         TabIndex        =   167
         Top             =   1650
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   20
         Left            =   7800
         TabIndex        =   198
         Top             =   1365
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   19
         Left            =   6255
         TabIndex        =   190
         Top             =   1365
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   18
         Left            =   4710
         TabIndex        =   182
         Top             =   1365
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   17
         Left            =   3180
         TabIndex        =   174
         Top             =   1365
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   16
         Left            =   1635
         TabIndex        =   166
         Top             =   1380
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   15
         Left            =   7800
         TabIndex        =   197
         Top             =   1095
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   14
         Left            =   6255
         TabIndex        =   189
         Top             =   1110
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   13
         Left            =   4710
         TabIndex        =   181
         Top             =   1110
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   12
         Left            =   3180
         TabIndex        =   173
         Top             =   1110
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   11
         Left            =   1635
         TabIndex        =   165
         Top             =   1110
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   10
         Left            =   7800
         TabIndex        =   196
         Top             =   825
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   9
         Left            =   6255
         TabIndex        =   188
         Top             =   840
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   8
         Left            =   4710
         TabIndex        =   180
         Top             =   840
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   7
         Left            =   3180
         TabIndex        =   172
         Top             =   840
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   6
         Left            =   1635
         TabIndex        =   164
         Top             =   840
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   5
         Left            =   7800
         TabIndex        =   195
         Top             =   570
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   4
         Left            =   6255
         TabIndex        =   187
         Top             =   570
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   3
         Left            =   4710
         TabIndex        =   179
         Top             =   570
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   2
         Left            =   3180
         TabIndex        =   171
         Top             =   570
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   1
         Left            =   1635
         TabIndex        =   163
         Top             =   570
         Width           =   1395
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "4号"
         Height          =   180
         Index           =   4
         Left            =   6570
         TabIndex        =   80
         Top             =   300
         Width           =   585
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "3号"
         Height          =   180
         Index           =   3
         Left            =   5055
         TabIndex        =   79
         Top             =   300
         Width           =   585
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "2号"
         Height          =   180
         Index           =   2
         Left            =   3525
         TabIndex        =   78
         Top             =   300
         Width           =   585
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "1号"
         Height          =   180
         Index           =   1
         Left            =   2010
         TabIndex        =   77
         Top             =   300
         Width           =   585
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   13
         Left            =   7740
         TabIndex        =   76
         Top             =   3540
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   14
         Left            =   7740
         TabIndex        =   75
         Top             =   3900
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   15
         Left            =   7740
         TabIndex        =   74
         Top             =   4275
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   16
         Left            =   7740
         TabIndex        =   73
         Top             =   4635
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   17
         Left            =   7740
         TabIndex        =   72
         Top             =   4995
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   18
         Left            =   7740
         TabIndex        =   71
         Top             =   5355
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   19
         Left            =   7740
         TabIndex        =   70
         Top             =   5730
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   20
         Left            =   7740
         TabIndex        =   69
         Top             =   6090
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   21
         Left            =   7740
         TabIndex        =   68
         Top             =   6450
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   22
         Left            =   7740
         TabIndex        =   67
         Top             =   6810
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   23
         Left            =   7740
         TabIndex        =   66
         Top             =   7185
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已校核或检修"
         Height          =   285
         Index           =   28
         Left            =   9360
         TabIndex        =   65
         Top             =   3540
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已检修或更换"
         Height          =   285
         Index           =   29
         Left            =   9360
         TabIndex        =   64
         Top             =   4275
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "有较大波动"
         Height          =   285
         Index           =   30
         Left            =   9360
         TabIndex        =   63
         Top             =   4635
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已做相应调整"
         Height          =   285
         Index           =   31
         Left            =   9360
         TabIndex        =   62
         Top             =   4995
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已检修或更换"
         Height          =   285
         Index           =   32
         Left            =   9360
         TabIndex        =   61
         Top             =   5355
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已调整或检修"
         Height          =   285
         Index           =   33
         Left            =   9360
         TabIndex        =   60
         Top             =   5730
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已调整或检修"
         Height          =   285
         Index           =   34
         Left            =   9360
         TabIndex        =   59
         Top             =   6090
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已做相应调整"
         Height          =   285
         Index           =   35
         Left            =   9360
         TabIndex        =   58
         Top             =   6450
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已校核或检修"
         Height          =   285
         Index           =   36
         Left            =   9360
         TabIndex        =   57
         Top             =   6810
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已更换过滤器"
         Height          =   285
         Index           =   37
         Left            =   9360
         TabIndex        =   56
         Top             =   7185
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   285
         Index           =   42
         Left            =   11520
         TabIndex        =   55
         Top             =   3540
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   285
         Index           =   43
         Left            =   11520
         TabIndex        =   54
         Top             =   3930
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   180
         Index           =   44
         Left            =   11520
         TabIndex        =   53
         Top             =   4335
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   285
         Index           =   45
         Left            =   11520
         TabIndex        =   52
         Top             =   4995
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   285
         Index           =   46
         Left            =   11520
         TabIndex        =   51
         Top             =   5355
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   285
         Index           =   47
         Left            =   11520
         TabIndex        =   50
         Top             =   5730
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   285
         Index           =   48
         Left            =   11520
         TabIndex        =   49
         Top             =   6090
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   285
         Index           =   49
         Left            =   11520
         TabIndex        =   48
         Top             =   6450
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   285
         Index           =   50
         Left            =   11520
         TabIndex        =   47
         Top             =   6810
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   285
         Index           =   51
         Left            =   11520
         TabIndex        =   46
         Top             =   7185
         Width           =   1335
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   52
         Left            =   13110
         TabIndex        =   45
         Top             =   3570
         Width           =   1875
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   53
         Left            =   13080
         TabIndex        =   44
         Top             =   4290
         Width           =   1875
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   285
         Index           =   54
         Left            =   13440
         TabIndex        =   43
         Top             =   3930
         Width           =   1005
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   285
         Index           =   55
         Left            =   13440
         TabIndex        =   42
         Top             =   5355
         Width           =   1005
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   285
         Index           =   56
         Left            =   13440
         TabIndex        =   41
         Top             =   5730
         Width           =   1005
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   285
         Index           =   57
         Left            =   13440
         TabIndex        =   40
         Top             =   6090
         Width           =   1005
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   285
         Index           =   58
         Left            =   13440
         TabIndex        =   39
         Top             =   6450
         Width           =   1005
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   285
         Index           =   59
         Left            =   13440
         TabIndex        =   38
         Top             =   6810
         Width           =   1005
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   285
         Index           =   60
         Left            =   13440
         TabIndex        =   37
         Top             =   7185
         Width           =   1005
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   180
         Index           =   68
         Left            =   1800
         TabIndex        =   36
         Top             =   2700
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "2"
         Height          =   180
         Index           =   69
         Left            =   2775
         TabIndex        =   35
         Top             =   2700
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "3"
         Height          =   180
         Index           =   70
         Left            =   3750
         TabIndex        =   34
         Top             =   2700
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "4"
         Height          =   180
         Index           =   71
         Left            =   4725
         TabIndex        =   33
         Top             =   2700
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "5"
         Height          =   180
         Index           =   72
         Left            =   5685
         TabIndex        =   32
         Top             =   2700
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "6"
         Height          =   180
         Index           =   73
         Left            =   6660
         TabIndex        =   31
         Top             =   2700
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "7"
         Height          =   180
         Index           =   74
         Left            =   7635
         TabIndex        =   30
         Top             =   2700
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "8"
         Height          =   180
         Index           =   75
         Left            =   8610
         TabIndex        =   29
         Top             =   2700
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "9"
         Height          =   180
         Index           =   76
         Left            =   9585
         TabIndex        =   28
         Top             =   2700
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "10"
         Height          =   180
         Index           =   77
         Left            =   10560
         TabIndex        =   27
         Top             =   2700
         Width           =   615
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "11"
         Height          =   180
         Index           =   78
         Left            =   11745
         TabIndex        =   26
         Top             =   2700
         Width           =   525
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "12"
         Height          =   180
         Index           =   79
         Left            =   12840
         TabIndex        =   25
         Top             =   2700
         Width           =   495
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   80
         Left            =   14370
         TabIndex        =   24
         Top             =   3000
         Width           =   405
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   61
         Left            =   1620
         TabIndex        =   214
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   62
         Left            =   2550
         TabIndex        =   215
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   63
         Left            =   3480
         TabIndex        =   216
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   64
         Left            =   4470
         TabIndex        =   217
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   65
         Left            =   5430
         TabIndex        =   218
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   66
         Left            =   6390
         TabIndex        =   219
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   67
         Left            =   7380
         TabIndex        =   220
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   68
         Left            =   8370
         ScrollBars      =   2  'Vertical
         TabIndex        =   221
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   69
         Left            =   9360
         TabIndex        =   222
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   70
         Left            =   10350
         TabIndex        =   223
         Top             =   3000
         Width           =   1035
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   71
         Left            =   11520
         TabIndex        =   224
         Top             =   3000
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   72
         Left            =   12690
         TabIndex        =   225
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   73
         Left            =   13650
         TabIndex        =   226
         Top             =   3000
         Width           =   585
      End
      Begin MSComCtl2.DTPicker dtpC 
         Height          =   225
         Left            =   -62040
         TabIndex        =   126
         Top             =   6690
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   397
         _Version        =   393216
         Format          =   149094401
         CurrentDate     =   38897
      End
      Begin MSComCtl2.DTPicker dtpB 
         Height          =   225
         Left            =   -64080
         TabIndex        =   128
         Top             =   6690
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   397
         _Version        =   393216
         Format          =   149094401
         CurrentDate     =   38897
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   510
         Index           =   59
         Left            =   -73590
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   232
         Top             =   5280
         Width           =   13515
      End
      Begin VB.Label Label39 
         Caption         =   "漏点描述"
         Height          =   195
         Index           =   0
         Left            =   -72120
         TabIndex        =   158
         Top             =   1545
         Width           =   885
      End
      Begin VB.Line Line38 
         Index           =   0
         X1              =   -71100
         X2              =   -59880
         Y1              =   1725
         Y2              =   1725
      End
      Begin VB.Label Label25 
         Caption         =   "原因:"
         Height          =   210
         Left            =   -63540
         TabIndex        =   157
         Top             =   915
         Width           =   585
      End
      Begin VB.Line Line24 
         X1              =   -62790
         X2              =   -60060
         Y1              =   1095
         Y2              =   1095
      End
      Begin VB.Line Line23 
         X1              =   -62040
         X2              =   -60150
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Label Label42 
         Caption         =   $"NewGzd2.frx":003B
         Height          =   1545
         Left            =   -74820
         TabIndex        =   141
         Top             =   120
         Width           =   3705
      End
      Begin VB.Label Label38 
         Caption         =   "日期："
         Height          =   195
         Left            =   -62010
         TabIndex        =   140
         Top             =   6450
         Width           =   945
      End
      Begin VB.Line Line37 
         X1              =   -62070
         X2              =   -62070
         Y1              =   5820
         Y2              =   7020
      End
      Begin VB.Line Line36 
         X1              =   -64170
         X2              =   -64170
         Y1              =   5820
         Y2              =   7020
      End
      Begin VB.Line Line35 
         X1              =   -74970
         X2              =   -60030
         Y1              =   6390
         Y2              =   6390
      End
      Begin VB.Line Line34 
         X1              =   -74970
         X2              =   -60030
         Y1              =   6060
         Y2              =   6060
      End
      Begin VB.Label Label37 
         Caption         =   "质量控制签名："
         Height          =   195
         Left            =   -62010
         TabIndex        =   139
         Top             =   5880
         Width           =   1275
      End
      Begin VB.Label Label36 
         Caption         =   "日期："
         Height          =   195
         Left            =   -64080
         TabIndex        =   138
         Top             =   6450
         Width           =   945
      End
      Begin VB.Label Label35 
         Caption         =   "客户签名："
         Height          =   225
         Left            =   -64080
         TabIndex        =   137
         Top             =   5850
         Width           =   945
      End
      Begin VB.Shape Shape2 
         Height          =   3165
         Left            =   -74970
         Top             =   3870
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
         Left            =   -74880
         TabIndex        =   136
         Top             =   6510
         Width           =   885
      End
      Begin VB.Label Label32 
         Caption         =   "加班工时"
         Height          =   165
         Left            =   -66300
         TabIndex        =   135
         Top             =   5880
         Width           =   1035
      End
      Begin VB.Label Label31 
         Caption         =   "旅途时间"
         Height          =   165
         Left            =   -68730
         TabIndex        =   134
         Top             =   5880
         Width           =   1035
      End
      Begin VB.Label Label30 
         Caption         =   "完成时间"
         Height          =   165
         Left            =   -71850
         TabIndex        =   133
         Top             =   5880
         Width           =   1035
      End
      Begin VB.Label Label29 
         Caption         =   "到达时间"
         Height          =   165
         Left            =   -74820
         TabIndex        =   132
         Top             =   5880
         Width           =   1035
      End
      Begin VB.Line Line33 
         X1              =   -74970
         X2              =   -60030
         Y1              =   5820
         Y2              =   5820
      End
      Begin VB.Line Line25 
         X1              =   -74970
         X2              =   -59910
         Y1              =   5250
         Y2              =   5250
      End
      Begin VB.Line Line29 
         X1              =   -74970
         X2              =   -59940
         Y1              =   4650
         Y2              =   4650
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
         Left            =   -74850
         TabIndex        =   131
         Top             =   5340
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
         Left            =   -74880
         TabIndex        =   130
         Top             =   4710
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
         Left            =   -74880
         TabIndex        =   129
         Top             =   3930
         Width           =   1005
      End
      Begin VB.Line Line1 
         X1              =   60
         X2              =   15105
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label1 
         Caption         =   "常规运行参数记录（记录与压缩机或风机对应的数据时，在压缩机或风机编号一栏中相应编号的""□""上打""√""，若无此压缩机则打""／""）"
         Height          =   195
         Left            =   450
         TabIndex        =   114
         Top             =   30
         Width           =   10875
      End
      Begin VB.Line Line18 
         X1              =   13620
         X2              =   13620
         Y1              =   300
         Y2              =   3210
      End
      Begin VB.Line Line17 
         X1              =   12090
         X2              =   12090
         Y1              =   240
         Y2              =   2690
      End
      Begin VB.Line Line16 
         X1              =   10710
         X2              =   10710
         Y1              =   240
         Y2              =   2690
      End
      Begin VB.Line Line15 
         X1              =   9300
         X2              =   9300
         Y1              =   240
         Y2              =   2690
      End
      Begin VB.Line Line14 
         X1              =   7710
         X2              =   7710
         Y1              =   240
         Y2              =   2690
      End
      Begin VB.Line Line13 
         X1              =   6180
         X2              =   6180
         Y1              =   240
         Y2              =   2690
      End
      Begin VB.Line Line12 
         X1              =   4650
         X2              =   4650
         Y1              =   240
         Y2              =   2690
      End
      Begin VB.Line Line11 
         X1              =   3090
         X2              =   3090
         Y1              =   240
         Y2              =   2690
      End
      Begin VB.Line Line10 
         X1              =   1590
         X2              =   1590
         Y1              =   240
         Y2              =   3240
      End
      Begin VB.Line Line9 
         X1              =   60
         X2              =   15030
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line8 
         X1              =   60
         X2              =   15030
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line7 
         X1              =   60
         X2              =   15030
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Line Line6 
         X1              =   60
         X2              =   15030
         Y1              =   1590
         Y2              =   1590
      End
      Begin VB.Line Line5 
         X1              =   60
         X2              =   15030
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line4 
         X1              =   60
         X2              =   15030
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Line Line3 
         X1              =   60
         X2              =   15030
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Line Line2 
         X1              =   60
         X2              =   15030
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Shape Shape1 
         Height          =   2985
         Left            =   60
         Top             =   240
         Width           =   14985
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "环境温度"
         Height          =   195
         Left            =   10830
         TabIndex        =   113
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "电压"
         Height          =   195
         Left            =   10830
         TabIndex        =   112
         Top             =   1365
         Width           =   1215
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "循环水进水温度"
         Height          =   195
         Left            =   10830
         TabIndex        =   111
         Top             =   1095
         Width           =   1215
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "循环水出水温度"
         Height          =   195
         Left            =   10830
         TabIndex        =   110
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "冷凝温度"
         Height          =   195
         Left            =   10830
         TabIndex        =   109
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "正常值"
         Height          =   165
         Left            =   13800
         TabIndex        =   108
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "机组参数"
         Height          =   165
         Left            =   12270
         TabIndex        =   107
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "无此项"
         Height          =   165
         Left            =   9705
         TabIndex        =   106
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "正常值"
         Height          =   165
         Left            =   7890
         TabIndex        =   105
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "电流"
         Height          =   165
         Left            =   270
         TabIndex        =   104
         Top             =   2460
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "负载百分比"
         Height          =   165
         Left            =   270
         TabIndex        =   103
         Top             =   2190
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "非运行时油温"
         Height          =   165
         Left            =   270
         TabIndex        =   102
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "排气压力"
         Height          =   165
         Left            =   270
         TabIndex        =   101
         Top             =   1110
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "吸气温度"
         Height          =   165
         Left            =   270
         TabIndex        =   100
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "吸气压力"
         Height          =   165
         Left            =   270
         TabIndex        =   99
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "油压力"
         Height          =   165
         Left            =   270
         TabIndex        =   98
         Top             =   1650
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "排气温度"
         Height          =   165
         Left            =   270
         TabIndex        =   97
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "压缩机编号"
         Height          =   165
         Left            =   270
         TabIndex        =   96
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "风机编号"
         Height          =   165
         Left            =   270
         TabIndex        =   95
         Top             =   2730
         Width           =   1125
      End
      Begin VB.Label Label21 
         Caption         =   "电流A"
         Height          =   195
         Left            =   270
         TabIndex        =   94
         Top             =   2970
         Width           =   1125
      End
      Begin VB.Line Line19 
         X1              =   30
         X2              =   15030
         Y1              =   2670
         Y2              =   2670
      End
      Begin VB.Line Line20 
         X1              =   60
         X2              =   15000
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Label Label23 
         Caption         =   "损坏情况描述"
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
         TabIndex        =   93
         Top             =   3240
         Width           =   1395
      End
      Begin VB.Label Label22 
         Caption         =   "巡视检修工作内容如下："
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
         Left            =   240
         TabIndex        =   92
         Top             =   3240
         Width           =   2145
      End
      Begin VB.Label Label24 
         Caption         =   $"NewGzd2.frx":009D
         Height          =   4485
         Left            =   210
         TabIndex        =   91
         Top             =   3540
         Width           =   6705
      End
      Begin VB.Line Line21 
         X1              =   13110
         X2              =   15000
         Y1              =   3750
         Y2              =   3750
      End
      Begin VB.Line Line22 
         X1              =   13050
         X2              =   14940
         Y1              =   4470
         Y2              =   4470
      End
      Begin VB.Label Label40 
         Caption         =   "正常值"
         Height          =   165
         Left            =   13650
         TabIndex        =   90
         Top             =   2730
         Width           =   555
      End
      Begin VB.Label Label41 
         Caption         =   "无法测量"
         Height          =   195
         Left            =   14280
         TabIndex        =   89
         Top             =   2730
         Width           =   735
      End
      Begin VB.Line Line39 
         X1              =   14250
         X2              =   14250
         Y1              =   2670
         Y2              =   3210
      End
      Begin VB.Line Line40 
         X1              =   12660
         X2              =   12660
         Y1              =   3210
         Y2              =   2670
      End
      Begin VB.Line Line41 
         X1              =   11430
         X2              =   11430
         Y1              =   2670
         Y2              =   3210
      End
      Begin VB.Line Line42 
         X1              =   10290
         X2              =   10290
         Y1              =   2670
         Y2              =   3225
      End
      Begin VB.Line Line43 
         X1              =   9300
         X2              =   9300
         Y1              =   2670
         Y2              =   3225
      End
      Begin VB.Line Line44 
         X1              =   8310
         X2              =   8310
         Y1              =   2670
         Y2              =   3210
      End
      Begin VB.Line Line45 
         X1              =   7320
         X2              =   7320
         Y1              =   2670
         Y2              =   3210
      End
      Begin VB.Line Line46 
         X1              =   6330
         X2              =   6330
         Y1              =   2670
         Y2              =   3210
      End
      Begin VB.Line Line47 
         X1              =   5400
         X2              =   5400
         Y1              =   2670
         Y2              =   3210
      End
      Begin VB.Line Line48 
         X1              =   4410
         X2              =   4410
         Y1              =   2670
         Y2              =   3210
      End
      Begin VB.Line Line49 
         X1              =   3420
         X2              =   3420
         Y1              =   2670
         Y2              =   3225
      End
      Begin VB.Line Line50 
         X1              =   2490
         X2              =   2490
         Y1              =   2670
         Y2              =   3240
      End
   End
   Begin MSDataGridLib.DataGrid comHtbh 
      Height          =   1155
      Left            =   6300
      TabIndex        =   22
      Top             =   0
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
      Left            =   1770
      TabIndex        =   21
      Top             =   690
      Width           =   4125
   End
   Begin VB.CommandButton cmdBack 
      Height          =   360
      Left            =   14490
      Picture         =   "NewGzd2.frx":02D8
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "返回"
      Top             =   10410
      Width           =   465
   End
   Begin VB.CommandButton cmdSave 
      Height          =   360
      Left            =   14010
      Picture         =   "NewGzd2.frx":03DA
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "保存"
      Top             =   10410
      Width           =   465
   End
   Begin VB.CommandButton cmdMod 
      Height          =   360
      Left            =   13530
      Picture         =   "NewGzd2.frx":0A44
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "修改"
      Top             =   10410
      Width           =   465
   End
   Begin VB.CommandButton cmdQm 
      Caption         =   "cmdQm"
      Height          =   360
      Index           =   0
      Left            =   9870
      TabIndex        =   12
      Top             =   10410
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
      Left            =   330
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "NewGzd2.frx":0D4E
      Top             =   60
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
      Left            =   6360
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "NewGzd2.frx":0D85
      Top             =   120
      Width           =   1335
   End
   Begin VB.CheckBox C1 
      Alignment       =   1  'Right Justify
      Caption         =   "1号"
      Height          =   285
      Index           =   0
      Left            =   12660
      TabIndex        =   9
      Top             =   930
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox TA 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   13560
      TabIndex        =   8
      Top             =   540
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
      Left            =   1800
      TabIndex        =   7
      Top             =   60
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
      Left            =   12780
      TabIndex        =   6
      Top             =   1170
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
      Left            =   1800
      TabIndex        =   5
      Top             =   480
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
      Left            =   1800
      TabIndex        =   4
      Top             =   900
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
      Left            =   7890
      TabIndex        =   3
      Text            =   "的"
      Top             =   120
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
      Left            =   7890
      TabIndex        =   2
      Top             =   540
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
      Left            =   7890
      Locked          =   -1  'True
      TabIndex        =   1
      Tag             =   "20"
      Top             =   930
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
      Left            =   12480
      TabIndex        =   0
      Text            =   "的"
      Top             =   60
      Visible         =   0   'False
      Width           =   4245
   End
   Begin MSComCtl2.DTPicker dtpA 
      Height          =   225
      Left            =   7890
      TabIndex        =   18
      Top             =   930
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   397
      _Version        =   393216
      Format          =   149094401
      CurrentDate     =   38897
   End
   Begin VB.Label LBLKjj 
      Caption         =   $"NewGzd2.frx":0DAE
      ForeColor       =   &H8000000D&
      Height          =   1305
      Left            =   12420
      TabIndex        =   241
      Top             =   300
      Width           =   2835
   End
   Begin VB.Line Line38 
      Index           =   2
      X1              =   1800
      X2              =   5955
      Y1              =   1530
      Y2              =   1530
   End
   Begin VB.Label Label39 
      Caption         =   "NO:"
      Height          =   255
      Index           =   1
      Left            =   12660
      TabIndex        =   161
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
      Left            =   13230
      TabIndex        =   160
      Top             =   120
      Width           =   1605
   End
   Begin VB.Line Line38 
      Index           =   1
      X1              =   1740
      X2              =   5895
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label lblkhdh 
      Caption         =   "lblkhdh"
      Height          =   225
      Left            =   11250
      TabIndex        =   20
      Top             =   1200
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label lblGid 
      Caption         =   "lblGid"
      Height          =   225
      Left            =   9240
      TabIndex        =   19
      Top             =   1320
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   0
      Left            =   10890
      TabIndex        =   17
      Top             =   10470
      Width           =   1905
   End
   Begin VB.Label lblQM 
      Caption         =   "签字提交"
      Height          =   360
      Index           =   0
      Left            =   8940
      TabIndex        =   16
      Top             =   10470
      Width           =   795
   End
   Begin VB.Line Line26 
      X1              =   1800
      X2              =   5895
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Line Line27 
      X1              =   1800
      X2              =   5880
      Y1              =   705
      Y2              =   705
   End
   Begin VB.Line Line28 
      X1              =   1800
      X2              =   5880
      Y1              =   1125
      Y2              =   1125
   End
   Begin VB.Line Line30 
      X1              =   7890
      X2              =   11970
      Y1              =   330
      Y2              =   330
   End
   Begin VB.Line Line31 
      X1              =   7890
      X2              =   11970
      Y1              =   750
      Y2              =   750
   End
   Begin VB.Line Line32 
      X1              =   7890
      X2              =   11970
      Y1              =   1170
      Y2              =   1170
   End
End
Attribute VB_Name = "NewGzd2"
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
If C1(13).Value = 1 Then
    For oo = 13 To 23
        C1(oo).Value = 0
    Next
Else
    For oo = 13 To 23
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
For oo = 1 To 73
    mod1.HTP.Update "mat" & oo, TA(oo).Text
Next
For oo = 1 To 80
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
For oo = 1 To 73
    TA(oo).Tag = 50
Next
TA(57).Tag = 200
TA(58).Tag = 200
TA(59).Tag = 200
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

Private Sub tabNr_Click(PreviousTab As Integer)
'If tabNr.Tab = 0 Then
'    TA(56).Visible = False
'Else
'    TA(56).Visible = True
'End If
End Sub

