VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form NewGZD1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "冷水机组巡视检修工作报告（单）"
   ClientHeight    =   10860
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   15045
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10860
   ScaleWidth      =   15045
   ShowInTaskbar   =   0   'False
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
      Left            =   1710
      TabIndex        =   144
      Top             =   1260
      Width           =   4065
   End
   Begin MSDataGridLib.DataGrid dtgRen 
      Height          =   8085
      Left            =   12240
      TabIndex        =   142
      Top             =   0
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
      TabIndex        =   24
      Top             =   1560
      Width           =   15285
      _ExtentX        =   26961
      _ExtentY        =   15372
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "内容1"
      TabPicture(0)   =   "NewGZD.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line23"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label25"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line22"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line21"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line20"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label24"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Line19"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label23"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label22"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label21"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Line18"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line17"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Line16"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Line15"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Line14"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Line13"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Line12"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Line11"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Line10"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Line9"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Line8"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Line7"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Line6"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Line5"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Line4"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Line3"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Line2"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Shape1"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label20"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label19"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label18"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label17"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label16"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label15"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label14"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label13"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Label12"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Label11"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Label10"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Label9"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Label8"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Line1"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Label7"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Label6"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Label5"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Label4"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Label3"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Label2"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Label1"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "TA(56)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "C1(61)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "C1(60)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "C1(26)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "C1(39)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "TA(55)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "TA(54)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "TA(53)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "C1(59)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "C1(58)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "C1(57)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "C1(56)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "C1(55)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "C1(54)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "C1(53)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "C1(52)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "C1(51)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "C1(50)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "C1(49)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "C1(48)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "C1(47)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "C1(46)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "C1(45)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "C1(44)"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "C1(43)"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "C1(42)"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "C1(41)"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "C1(40)"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "TA(52)"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "C1(38)"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "C1(37)"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "C1(36)"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "C1(35)"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "C1(34)"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "C1(33)"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).Control(84)=   "C1(32)"
      Tab(0).Control(84).Enabled=   0   'False
      Tab(0).Control(85)=   "C1(31)"
      Tab(0).Control(85).Enabled=   0   'False
      Tab(0).Control(86)=   "C1(30)"
      Tab(0).Control(86).Enabled=   0   'False
      Tab(0).Control(87)=   "C1(29)"
      Tab(0).Control(87).Enabled=   0   'False
      Tab(0).Control(88)=   "C1(28)"
      Tab(0).Control(88).Enabled=   0   'False
      Tab(0).Control(89)=   "C1(27)"
      Tab(0).Control(89).Enabled=   0   'False
      Tab(0).Control(90)=   "C1(25)"
      Tab(0).Control(90).Enabled=   0   'False
      Tab(0).Control(91)=   "C1(24)"
      Tab(0).Control(91).Enabled=   0   'False
      Tab(0).Control(92)=   "C1(23)"
      Tab(0).Control(92).Enabled=   0   'False
      Tab(0).Control(93)=   "C1(22)"
      Tab(0).Control(93).Enabled=   0   'False
      Tab(0).Control(94)=   "C1(21)"
      Tab(0).Control(94).Enabled=   0   'False
      Tab(0).Control(95)=   "C1(20)"
      Tab(0).Control(95).Enabled=   0   'False
      Tab(0).Control(96)=   "C1(19)"
      Tab(0).Control(96).Enabled=   0   'False
      Tab(0).Control(97)=   "C1(18)"
      Tab(0).Control(97).Enabled=   0   'False
      Tab(0).Control(98)=   "C1(17)"
      Tab(0).Control(98).Enabled=   0   'False
      Tab(0).Control(99)=   "C1(16)"
      Tab(0).Control(99).Enabled=   0   'False
      Tab(0).Control(100)=   "C1(15)"
      Tab(0).Control(100).Enabled=   0   'False
      Tab(0).Control(101)=   "C1(14)"
      Tab(0).Control(101).Enabled=   0   'False
      Tab(0).Control(102)=   "C1(13)"
      Tab(0).Control(102).Enabled=   0   'False
      Tab(0).Control(103)=   "TA(51)"
      Tab(0).Control(103).Enabled=   0   'False
      Tab(0).Control(104)=   "TA(50)"
      Tab(0).Control(104).Enabled=   0   'False
      Tab(0).Control(105)=   "TA(49)"
      Tab(0).Control(105).Enabled=   0   'False
      Tab(0).Control(106)=   "TA(48)"
      Tab(0).Control(106).Enabled=   0   'False
      Tab(0).Control(107)=   "TA(47)"
      Tab(0).Control(107).Enabled=   0   'False
      Tab(0).Control(108)=   "TA(46)"
      Tab(0).Control(108).Enabled=   0   'False
      Tab(0).Control(109)=   "TA(45)"
      Tab(0).Control(109).Enabled=   0   'False
      Tab(0).Control(110)=   "TA(44)"
      Tab(0).Control(110).Enabled=   0   'False
      Tab(0).Control(111)=   "TA(43)"
      Tab(0).Control(111).Enabled=   0   'False
      Tab(0).Control(112)=   "TA(42)"
      Tab(0).Control(112).Enabled=   0   'False
      Tab(0).Control(113)=   "TA(41)"
      Tab(0).Control(113).Enabled=   0   'False
      Tab(0).Control(114)=   "C1(12)"
      Tab(0).Control(114).Enabled=   0   'False
      Tab(0).Control(115)=   "C1(11)"
      Tab(0).Control(115).Enabled=   0   'False
      Tab(0).Control(116)=   "C1(10)"
      Tab(0).Control(116).Enabled=   0   'False
      Tab(0).Control(117)=   "C1(9)"
      Tab(0).Control(117).Enabled=   0   'False
      Tab(0).Control(118)=   "C1(8)"
      Tab(0).Control(118).Enabled=   0   'False
      Tab(0).Control(119)=   "C1(7)"
      Tab(0).Control(119).Enabled=   0   'False
      Tab(0).Control(120)=   "C1(6)"
      Tab(0).Control(120).Enabled=   0   'False
      Tab(0).Control(121)=   "C1(5)"
      Tab(0).Control(121).Enabled=   0   'False
      Tab(0).Control(122)=   "TA(40)"
      Tab(0).Control(122).Enabled=   0   'False
      Tab(0).Control(123)=   "TA(39)"
      Tab(0).Control(123).Enabled=   0   'False
      Tab(0).Control(124)=   "TA(38)"
      Tab(0).Control(124).Enabled=   0   'False
      Tab(0).Control(125)=   "TA(37)"
      Tab(0).Control(125).Enabled=   0   'False
      Tab(0).Control(126)=   "TA(36)"
      Tab(0).Control(126).Enabled=   0   'False
      Tab(0).Control(127)=   "TA(35)"
      Tab(0).Control(127).Enabled=   0   'False
      Tab(0).Control(128)=   "TA(34)"
      Tab(0).Control(128).Enabled=   0   'False
      Tab(0).Control(129)=   "TA(33)"
      Tab(0).Control(129).Enabled=   0   'False
      Tab(0).Control(130)=   "TA(32)"
      Tab(0).Control(130).Enabled=   0   'False
      Tab(0).Control(131)=   "TA(31)"
      Tab(0).Control(131).Enabled=   0   'False
      Tab(0).Control(132)=   "TA(30)"
      Tab(0).Control(132).Enabled=   0   'False
      Tab(0).Control(133)=   "TA(29)"
      Tab(0).Control(133).Enabled=   0   'False
      Tab(0).Control(134)=   "TA(28)"
      Tab(0).Control(134).Enabled=   0   'False
      Tab(0).Control(135)=   "TA(27)"
      Tab(0).Control(135).Enabled=   0   'False
      Tab(0).Control(136)=   "TA(26)"
      Tab(0).Control(136).Enabled=   0   'False
      Tab(0).Control(137)=   "TA(25)"
      Tab(0).Control(137).Enabled=   0   'False
      Tab(0).Control(138)=   "TA(24)"
      Tab(0).Control(138).Enabled=   0   'False
      Tab(0).Control(139)=   "TA(23)"
      Tab(0).Control(139).Enabled=   0   'False
      Tab(0).Control(140)=   "TA(22)"
      Tab(0).Control(140).Enabled=   0   'False
      Tab(0).Control(141)=   "TA(21)"
      Tab(0).Control(141).Enabled=   0   'False
      Tab(0).Control(142)=   "TA(20)"
      Tab(0).Control(142).Enabled=   0   'False
      Tab(0).Control(143)=   "TA(19)"
      Tab(0).Control(143).Enabled=   0   'False
      Tab(0).Control(144)=   "TA(18)"
      Tab(0).Control(144).Enabled=   0   'False
      Tab(0).Control(145)=   "TA(17)"
      Tab(0).Control(145).Enabled=   0   'False
      Tab(0).Control(146)=   "TA(16)"
      Tab(0).Control(146).Enabled=   0   'False
      Tab(0).Control(147)=   "TA(15)"
      Tab(0).Control(147).Enabled=   0   'False
      Tab(0).Control(148)=   "TA(14)"
      Tab(0).Control(148).Enabled=   0   'False
      Tab(0).Control(149)=   "TA(13)"
      Tab(0).Control(149).Enabled=   0   'False
      Tab(0).Control(150)=   "TA(12)"
      Tab(0).Control(150).Enabled=   0   'False
      Tab(0).Control(151)=   "TA(11)"
      Tab(0).Control(151).Enabled=   0   'False
      Tab(0).Control(152)=   "TA(10)"
      Tab(0).Control(152).Enabled=   0   'False
      Tab(0).Control(153)=   "TA(9)"
      Tab(0).Control(153).Enabled=   0   'False
      Tab(0).Control(154)=   "TA(8)"
      Tab(0).Control(154).Enabled=   0   'False
      Tab(0).Control(155)=   "TA(7)"
      Tab(0).Control(155).Enabled=   0   'False
      Tab(0).Control(156)=   "TA(6)"
      Tab(0).Control(156).Enabled=   0   'False
      Tab(0).Control(157)=   "TA(5)"
      Tab(0).Control(157).Enabled=   0   'False
      Tab(0).Control(158)=   "TA(4)"
      Tab(0).Control(158).Enabled=   0   'False
      Tab(0).Control(159)=   "TA(3)"
      Tab(0).Control(159).Enabled=   0   'False
      Tab(0).Control(160)=   "TA(2)"
      Tab(0).Control(160).Enabled=   0   'False
      Tab(0).Control(161)=   "TA(1)"
      Tab(0).Control(161).Enabled=   0   'False
      Tab(0).Control(162)=   "C1(4)"
      Tab(0).Control(162).Enabled=   0   'False
      Tab(0).Control(163)=   "C1(3)"
      Tab(0).Control(163).Enabled=   0   'False
      Tab(0).Control(164)=   "C1(2)"
      Tab(0).Control(164).Enabled=   0   'False
      Tab(0).Control(165)=   "C1(1)"
      Tab(0).Control(165).Enabled=   0   'False
      Tab(0).Control(166)=   "cmdAll"
      Tab(0).Control(166).Enabled=   0   'False
      Tab(0).ControlCount=   167
      TabCaption(1)   =   "内容2"
      TabPicture(1)   =   "NewGZD.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "BA(14)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "TA(58)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "C1(62)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "C1(63)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "TA(60)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "BA(8)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "BA(9)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "BA(10)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "BA(11)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Frame1"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "BA(12)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "BA(13)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "BA(15)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "BA(16)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "dtpB"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "dtpC"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "TA(59)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "TA(57)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label26"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label27"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label28"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Line24"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Line25"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Line33"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Label29"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Label30"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Label31"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Label32"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Label34"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Shape2"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Label35"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Label36"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Label37"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Line34"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Line35"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "Line36"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Line37"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Label38"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).ControlCount=   39
      Begin VB.CommandButton cmdAll 
         Caption         =   "全部"
         Height          =   255
         Left            =   8010
         TabIndex        =   206
         Top             =   8160
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
         Top             =   2820
         Width           =   1725
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   570
         Index           =   58
         Left            =   -73560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   197
         Top             =   810
         Width           =   13545
      End
      Begin VB.CheckBox C1 
         Caption         =   "完成"
         Height          =   180
         Index           =   62
         Left            =   -62250
         TabIndex        =   126
         Top             =   600
         Width           =   1005
      End
      Begin VB.CheckBox C1 
         Caption         =   "未完成"
         Height          =   180
         Index           =   63
         Left            =   -61110
         TabIndex        =   125
         Top             =   600
         Width           =   945
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   -61860
         TabIndex        =   124
         Text            =   "复核人:"
         Top             =   1770
         Width           =   735
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   60
         Left            =   -61230
         TabIndex        =   123
         Top             =   1680
         Width           =   1065
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   8
         Left            =   -73530
         TabIndex        =   199
         Text            =   "的"
         Top             =   2010
         Width           =   1035
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   9
         Left            =   -70290
         TabIndex        =   200
         Text            =   "的"
         Top             =   2010
         Width           =   1035
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   10
         Left            =   -67620
         TabIndex        =   201
         Text            =   "的"
         Top             =   2010
         Width           =   1035
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   11
         Left            =   -65220
         TabIndex        =   202
         Text            =   "的"
         Top             =   2010
         Width           =   1035
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Left            =   -74940
         TabIndex        =   117
         Top             =   2250
         Width           =   10755
         Begin VB.OptionButton FPA 
            Caption         =   "优秀"
            Height          =   195
            Left            =   1350
            TabIndex        =   121
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton FPB 
            Caption         =   "满意"
            Height          =   195
            Left            =   2950
            TabIndex        =   120
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton FPC 
            Caption         =   "较满意"
            Height          =   195
            Left            =   4550
            TabIndex        =   119
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton FPD 
            Caption         =   "尚可"
            Height          =   195
            Left            =   6150
            TabIndex        =   118
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
            TabIndex        =   122
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
         TabIndex        =   203
         Text            =   "NewGZD.frx":0038
         Top             =   2580
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
         TabIndex        =   204
         Top             =   2250
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
         Index           =   15
         Left            =   -62010
         TabIndex        =   205
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
         Height          =   240
         Index           =   16
         Left            =   -62040
         TabIndex        =   116
         Top             =   2820
         Width           =   1755
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "1号"
         Height          =   180
         Index           =   1
         Left            =   2190
         TabIndex        =   90
         Top             =   360
         Width           =   585
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "2号"
         Height          =   180
         Index           =   2
         Left            =   3705
         TabIndex        =   89
         Top             =   360
         Width           =   585
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "3号"
         Height          =   180
         Index           =   3
         Left            =   5235
         TabIndex        =   88
         Top             =   360
         Width           =   585
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "4号"
         Height          =   180
         Index           =   4
         Left            =   6750
         TabIndex        =   87
         Top             =   360
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   1
         Left            =   1815
         TabIndex        =   145
         Top             =   630
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   2
         Left            =   3360
         TabIndex        =   153
         Top             =   630
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   3
         Left            =   4890
         TabIndex        =   161
         Top             =   630
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   4
         Left            =   6435
         TabIndex        =   169
         Top             =   630
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   5
         Left            =   7980
         TabIndex        =   177
         Top             =   630
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   6
         Left            =   1815
         TabIndex        =   146
         Top             =   900
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   7
         Left            =   3360
         TabIndex        =   154
         Top             =   900
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   8
         Left            =   4890
         TabIndex        =   162
         Top             =   900
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   9
         Left            =   6435
         TabIndex        =   170
         Top             =   900
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   10
         Left            =   7980
         TabIndex        =   178
         Top             =   885
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   11
         Left            =   1815
         TabIndex        =   147
         Top             =   1170
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   12
         Left            =   3360
         TabIndex        =   155
         Top             =   1170
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   13
         Left            =   4890
         TabIndex        =   163
         Top             =   1170
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   14
         Left            =   6435
         TabIndex        =   171
         Top             =   1170
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   15
         Left            =   7980
         TabIndex        =   179
         Top             =   1155
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   16
         Left            =   1815
         TabIndex        =   148
         Top             =   1440
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   17
         Left            =   3360
         TabIndex        =   156
         Top             =   1425
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   18
         Left            =   4890
         TabIndex        =   164
         Top             =   1425
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   19
         Left            =   6435
         TabIndex        =   172
         Top             =   1425
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   20
         Left            =   7980
         TabIndex        =   180
         Top             =   1425
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   21
         Left            =   1815
         TabIndex        =   149
         Top             =   1710
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   22
         Left            =   3360
         TabIndex        =   157
         Top             =   1695
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   23
         Left            =   4890
         TabIndex        =   165
         Top             =   1695
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   24
         Left            =   6435
         TabIndex        =   173
         Top             =   1695
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   25
         Left            =   7980
         TabIndex        =   181
         Top             =   1695
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   26
         Left            =   1815
         TabIndex        =   150
         Top             =   1980
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   27
         Left            =   3360
         TabIndex        =   158
         Top             =   1965
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   28
         Left            =   4890
         TabIndex        =   166
         Top             =   1965
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   29
         Left            =   6435
         TabIndex        =   174
         Top             =   1965
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   30
         Left            =   7980
         TabIndex        =   182
         Top             =   1965
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   31
         Left            =   1815
         TabIndex        =   151
         Top             =   2250
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   32
         Left            =   3360
         TabIndex        =   159
         Top             =   2235
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   33
         Left            =   4890
         TabIndex        =   167
         Top             =   2235
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   34
         Left            =   6435
         TabIndex        =   175
         Top             =   2235
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   35
         Left            =   7980
         TabIndex        =   183
         Top             =   2235
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   36
         Left            =   1815
         TabIndex        =   152
         Top             =   2505
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   37
         Left            =   3360
         TabIndex        =   160
         Top             =   2505
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   38
         Left            =   4890
         TabIndex        =   168
         Top             =   2505
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   39
         Left            =   6435
         TabIndex        =   176
         Top             =   2505
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   40
         Left            =   7980
         TabIndex        =   184
         Top             =   2505
         Width           =   1395
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   5
         Left            =   9900
         TabIndex        =   86
         Top             =   630
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   6
         Left            =   9900
         TabIndex        =   85
         Top             =   900
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   7
         Left            =   9900
         TabIndex        =   84
         Top             =   1170
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   8
         Left            =   9900
         TabIndex        =   83
         Top             =   1425
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   9
         Left            =   9900
         TabIndex        =   82
         Top             =   1695
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   10
         Left            =   9900
         TabIndex        =   81
         Top             =   1965
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   11
         Left            =   9900
         TabIndex        =   80
         Top             =   2235
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   12
         Left            =   9900
         TabIndex        =   79
         Top             =   2505
         Width           =   405
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   41
         Left            =   12360
         TabIndex        =   185
         Top             =   660
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   42
         Left            =   13860
         TabIndex        =   190
         Top             =   660
         Width           =   1335
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   43
         Left            =   12360
         TabIndex        =   186
         Top             =   915
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   44
         Left            =   13860
         TabIndex        =   191
         Top             =   915
         Width           =   1335
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   45
         Left            =   12360
         TabIndex        =   187
         Top             =   1185
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   46
         Left            =   13860
         TabIndex        =   192
         Top             =   1185
         Width           =   1335
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   47
         Left            =   12360
         TabIndex        =   188
         Top             =   1440
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   48
         Left            =   13860
         TabIndex        =   193
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   49
         Left            =   12360
         TabIndex        =   189
         Top             =   1695
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   50
         Left            =   13860
         TabIndex        =   194
         Top             =   1695
         Width           =   1335
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   840
         Index           =   51
         Left            =   12300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   195
         Top             =   1950
         Width           =   2895
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   180
         Index           =   13
         Left            =   8070
         TabIndex        =   78
         Top             =   3180
         Width           =   795
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   180
         Index           =   14
         Left            =   8070
         TabIndex        =   77
         Top             =   3540
         Width           =   795
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   180
         Index           =   15
         Left            =   8070
         TabIndex        =   76
         Top             =   3900
         Width           =   795
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   180
         Index           =   16
         Left            =   8070
         TabIndex        =   75
         Top             =   4260
         Width           =   795
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   180
         Index           =   17
         Left            =   8070
         TabIndex        =   74
         Top             =   4635
         Width           =   795
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   180
         Index           =   18
         Left            =   8070
         TabIndex        =   73
         Top             =   4995
         Width           =   795
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   180
         Index           =   19
         Left            =   8070
         TabIndex        =   72
         Top             =   5355
         Width           =   795
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   180
         Index           =   20
         Left            =   8070
         TabIndex        =   71
         Top             =   5715
         Width           =   795
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   180
         Index           =   21
         Left            =   8070
         TabIndex        =   70
         Top             =   6075
         Width           =   795
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   180
         Index           =   22
         Left            =   8070
         TabIndex        =   69
         Top             =   6435
         Width           =   795
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   180
         Index           =   23
         Left            =   8070
         TabIndex        =   68
         Top             =   6795
         Width           =   795
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   180
         Index           =   24
         Left            =   8070
         TabIndex        =   67
         Top             =   7155
         Width           =   795
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   180
         Index           =   25
         Left            =   8070
         TabIndex        =   66
         Top             =   7530
         Width           =   795
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已校核或检修"
         Height          =   180
         Index           =   27
         Left            =   9570
         TabIndex        =   65
         Top             =   3180
         Width           =   1635
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已调整或检修"
         Height          =   180
         Index           =   28
         Left            =   9570
         TabIndex        =   64
         Top             =   3900
         Width           =   1635
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已检修或更换"
         Height          =   180
         Index           =   29
         Left            =   9570
         TabIndex        =   63
         Top             =   4260
         Width           =   1635
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已检修或更换"
         Height          =   180
         Index           =   30
         Left            =   9570
         TabIndex        =   62
         Top             =   4635
         Width           =   1635
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "有较大波动"
         Height          =   180
         Index           =   31
         Left            =   9570
         TabIndex        =   61
         Top             =   4995
         Width           =   1635
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已做相应调整"
         Height          =   180
         Index           =   32
         Left            =   9570
         TabIndex        =   60
         Top             =   5355
         Width           =   1635
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已检修或更换"
         Height          =   180
         Index           =   33
         Left            =   9570
         TabIndex        =   59
         Top             =   5715
         Width           =   1635
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已调整"
         Height          =   180
         Index           =   34
         Left            =   9570
         TabIndex        =   58
         Top             =   6075
         Width           =   1635
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已校核或检修"
         Height          =   180
         Index           =   35
         Left            =   9570
         TabIndex        =   57
         Top             =   6435
         Width           =   1635
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已做相应调整"
         Height          =   180
         Index           =   36
         Left            =   9570
         TabIndex        =   56
         Top             =   6795
         Width           =   1635
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已调整或检修"
         Height          =   180
         Index           =   37
         Left            =   9570
         TabIndex        =   55
         Top             =   7155
         Width           =   1635
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "因故未完成清洁"
         Height          =   180
         Index           =   38
         Left            =   9570
         TabIndex        =   54
         Top             =   7530
         Width           =   1635
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   52
         Left            =   13290
         TabIndex        =   53
         Top             =   3090
         Width           =   2145
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   180
         Index           =   40
         Left            =   11580
         TabIndex        =   52
         Top             =   3180
         Width           =   1515
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   180
         Index           =   41
         Left            =   11580
         TabIndex        =   51
         Top             =   3540
         Width           =   1515
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   180
         Index           =   42
         Left            =   11580
         TabIndex        =   50
         Top             =   3900
         Width           =   1515
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   180
         Index           =   43
         Left            =   11580
         TabIndex        =   49
         Top             =   4260
         Width           =   1515
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   180
         Index           =   44
         Left            =   11580
         TabIndex        =   48
         Top             =   4635
         Width           =   1515
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "油色已变差"
         Height          =   180
         Index           =   45
         Left            =   11580
         TabIndex        =   47
         Top             =   5370
         Width           =   1515
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   180
         Index           =   46
         Left            =   11580
         TabIndex        =   46
         Top             =   5730
         Width           =   1515
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   180
         Index           =   47
         Left            =   11580
         TabIndex        =   45
         Top             =   6090
         Width           =   1515
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   180
         Index           =   48
         Left            =   11580
         TabIndex        =   44
         Top             =   6450
         Width           =   1515
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   180
         Index           =   49
         Left            =   11580
         TabIndex        =   43
         Top             =   6810
         Width           =   1515
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   180
         Index           =   50
         Left            =   11580
         TabIndex        =   42
         Top             =   7155
         Width           =   1515
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   180
         Index           =   51
         Left            =   13980
         TabIndex        =   41
         Top             =   3540
         Width           =   1005
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   180
         Index           =   52
         Left            =   13980
         TabIndex        =   40
         Top             =   4275
         Width           =   1005
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   180
         Index           =   53
         Left            =   13980
         TabIndex        =   39
         Top             =   5730
         Width           =   1005
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   180
         Index           =   54
         Left            =   13980
         TabIndex        =   38
         Top             =   6090
         Width           =   1005
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   180
         Index           =   55
         Left            =   13980
         TabIndex        =   37
         Top             =   6450
         Width           =   1005
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   180
         Index           =   56
         Left            =   13980
         TabIndex        =   36
         Top             =   6810
         Width           =   1005
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   180
         Index           =   57
         Left            =   13980
         TabIndex        =   35
         Top             =   7155
         Width           =   1005
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   180
         Index           =   58
         Left            =   11580
         TabIndex        =   34
         Top             =   7890
         Width           =   1515
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   180
         Index           =   59
         Left            =   13980
         TabIndex        =   33
         Top             =   7890
         Width           =   1005
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   53
         Left            =   13290
         TabIndex        =   32
         Top             =   3900
         Width           =   2145
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   54
         Left            =   13290
         TabIndex        =   31
         Top             =   4650
         Width           =   2145
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   55
         Left            =   12240
         TabIndex        =   30
         Top             =   7530
         Width           =   3135
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已调整或检修"
         Height          =   180
         Index           =   39
         Left            =   9570
         TabIndex        =   29
         Top             =   7890
         Width           =   1635
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   180
         Index           =   26
         Left            =   8070
         TabIndex        =   28
         Top             =   7890
         Width           =   795
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "目测检漏"
         Height          =   180
         Index           =   60
         Left            =   330
         TabIndex        =   27
         Top             =   8250
         Width           =   1095
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "仪器检漏"
         Height          =   180
         Index           =   61
         Left            =   1830
         TabIndex        =   26
         Top             =   8250
         Width           =   1065
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   56
         Left            =   4200
         TabIndex        =   25
         Top             =   8250
         Width           =   11175
      End
      Begin MSComCtl2.DTPicker dtpB 
         Height          =   225
         Left            =   -64080
         TabIndex        =   128
         Top             =   2820
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   397
         _Version        =   393216
         Format          =   149094401
         CurrentDate     =   38897
      End
      Begin MSComCtl2.DTPicker dtpC 
         Height          =   225
         Left            =   -62040
         TabIndex        =   129
         Top             =   2820
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
         Left            =   -73560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   198
         Top             =   1440
         Width           =   13515
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   57
         Left            =   -73560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   196
         Top             =   60
         Width           =   13545
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
         TabIndex        =   141
         Top             =   60
         Width           =   1005
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
         TabIndex        =   140
         Top             =   840
         Width           =   1125
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
         TabIndex        =   139
         Top             =   1470
         Width           =   1035
      End
      Begin VB.Line Line24 
         X1              =   -74970
         X2              =   -59940
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Line Line25 
         X1              =   -74970
         X2              =   -59910
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Line Line33 
         X1              =   -74970
         X2              =   -60030
         Y1              =   1950
         Y2              =   1950
      End
      Begin VB.Label Label29 
         Caption         =   "到达时间"
         Height          =   165
         Left            =   -74820
         TabIndex        =   138
         Top             =   2010
         Width           =   1035
      End
      Begin VB.Label Label30 
         Caption         =   "完成时间"
         Height          =   165
         Left            =   -71850
         TabIndex        =   137
         Top             =   2010
         Width           =   1035
      End
      Begin VB.Label Label31 
         Caption         =   "旅途时间"
         Height          =   165
         Left            =   -68730
         TabIndex        =   136
         Top             =   2010
         Width           =   1035
      End
      Begin VB.Label Label32 
         Caption         =   "加班工时"
         Height          =   165
         Left            =   -66300
         TabIndex        =   135
         Top             =   2010
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
         Left            =   -74880
         TabIndex        =   134
         Top             =   2640
         Width           =   885
      End
      Begin VB.Shape Shape2 
         Height          =   3165
         Left            =   -74970
         Top             =   30
         Width           =   14985
      End
      Begin VB.Label Label35 
         Caption         =   "客户签名："
         Height          =   195
         Left            =   -64080
         TabIndex        =   133
         Top             =   1980
         Width           =   945
      End
      Begin VB.Label Label36 
         Caption         =   "日期："
         Height          =   195
         Left            =   -64080
         TabIndex        =   132
         Top             =   2580
         Width           =   945
      End
      Begin VB.Label Label37 
         Caption         =   "质量控制签名："
         Height          =   195
         Left            =   -62010
         TabIndex        =   131
         Top             =   2010
         Width           =   1275
      End
      Begin VB.Line Line34 
         X1              =   -74970
         X2              =   -60030
         Y1              =   2190
         Y2              =   2190
      End
      Begin VB.Line Line35 
         X1              =   -74970
         X2              =   -60030
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line36 
         X1              =   -64170
         X2              =   -64170
         Y1              =   1950
         Y2              =   3150
      End
      Begin VB.Line Line37 
         X1              =   -62070
         X2              =   -62070
         Y1              =   1950
         Y2              =   3150
      End
      Begin VB.Label Label38 
         Caption         =   "日期："
         Height          =   195
         Left            =   -62010
         TabIndex        =   130
         Top             =   2580
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "常规运行参数记录（记录与压缩机对应的数据时，在压缩机编号一栏中相应编号的""□""上打""√""，若无此压缩机则打""×""）"
         Height          =   165
         Left            =   390
         TabIndex        =   115
         Top             =   60
         Width           =   10575
      End
      Begin VB.Label Label2 
         Caption         =   "压缩机编号"
         Height          =   165
         Left            =   450
         TabIndex        =   114
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "供液温度"
         Height          =   165
         Left            =   450
         TabIndex        =   113
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "油压差"
         Height          =   165
         Left            =   450
         TabIndex        =   112
         Top             =   1710
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "吸气压力"
         Height          =   165
         Left            =   450
         TabIndex        =   111
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "排气压力"
         Height          =   165
         Left            =   450
         TabIndex        =   110
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "排气温度"
         Height          =   165
         Left            =   450
         TabIndex        =   109
         Top             =   1170
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   30
         X2              =   15240
         Y1              =   270
         Y2              =   270
      End
      Begin VB.Label Label8 
         Caption         =   "油温"
         Height          =   165
         Left            =   450
         TabIndex        =   108
         Top             =   1980
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "负载百分比"
         Height          =   165
         Left            =   450
         TabIndex        =   107
         Top             =   2250
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "电流"
         Height          =   165
         Left            =   450
         TabIndex        =   106
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "正常值"
         Height          =   165
         Left            =   8070
         TabIndex        =   105
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "无此项"
         Height          =   165
         Left            =   9885
         TabIndex        =   104
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "机组参数"
         Height          =   165
         Left            =   12450
         TabIndex        =   103
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "正常值"
         Height          =   165
         Left            =   13980
         TabIndex        =   102
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "冷冻出水温度"
         Height          =   195
         Left            =   11010
         TabIndex        =   101
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "冷冻进水温度"
         Height          =   195
         Left            =   11010
         TabIndex        =   100
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "冷却出水温度"
         Height          =   195
         Left            =   11010
         TabIndex        =   99
         Top             =   1155
         Width           =   1215
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "冷却进水温度"
         Height          =   195
         Left            =   11010
         TabIndex        =   98
         Top             =   1425
         Width           =   1215
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "电压"
         Height          =   195
         Left            =   11010
         TabIndex        =   97
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "其它："
         Height          =   195
         Left            =   11010
         TabIndex        =   96
         Top             =   1950
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         Height          =   2505
         Left            =   240
         Top             =   300
         Width           =   14985
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   15210
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   15210
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line4 
         X1              =   240
         X2              =   15210
         Y1              =   1110
         Y2              =   1110
      End
      Begin VB.Line Line5 
         X1              =   240
         X2              =   15210
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Line Line6 
         X1              =   240
         X2              =   15210
         Y1              =   1650
         Y2              =   1650
      End
      Begin VB.Line Line7 
         X1              =   240
         X2              =   15210
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line8 
         X1              =   240
         X2              =   15210
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Line Line9 
         X1              =   240
         X2              =   15210
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Line Line10 
         X1              =   1770
         X2              =   1770
         Y1              =   300
         Y2              =   2790
      End
      Begin VB.Line Line11 
         X1              =   3270
         X2              =   3270
         Y1              =   300
         Y2              =   2790
      End
      Begin VB.Line Line12 
         X1              =   4830
         X2              =   4830
         Y1              =   300
         Y2              =   2790
      End
      Begin VB.Line Line13 
         X1              =   6360
         X2              =   6360
         Y1              =   300
         Y2              =   2790
      End
      Begin VB.Line Line14 
         X1              =   7890
         X2              =   7890
         Y1              =   300
         Y2              =   2790
      End
      Begin VB.Line Line15 
         X1              =   9480
         X2              =   9480
         Y1              =   300
         Y2              =   2790
      End
      Begin VB.Line Line16 
         X1              =   10890
         X2              =   10890
         Y1              =   300
         Y2              =   2790
      End
      Begin VB.Line Line17 
         X1              =   12270
         X2              =   12270
         Y1              =   300
         Y2              =   2790
      End
      Begin VB.Line Line18 
         X1              =   13800
         X2              =   13800
         Y1              =   300
         Y2              =   2790
      End
      Begin VB.Label Label21 
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
         Left            =   390
         TabIndex        =   95
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label Label22 
         Caption         =   $"NewGZD.frx":003B
         Height          =   4965
         Left            =   360
         TabIndex        =   94
         Top             =   3150
         Width           =   7125
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
         Left            =   13890
         TabIndex        =   93
         Top             =   2880
         Width           =   1395
      End
      Begin VB.Line Line19 
         X1              =   13290
         X2              =   15480
         Y1              =   3270
         Y2              =   3270
      End
      Begin VB.Label Label24 
         Caption         =   "原因:"
         Height          =   180
         Left            =   11580
         TabIndex        =   92
         Top             =   7530
         Width           =   585
      End
      Begin VB.Line Line20 
         X1              =   13290
         X2              =   15480
         Y1              =   4830
         Y2              =   4830
      End
      Begin VB.Line Line21 
         X1              =   13260
         X2              =   15450
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line22 
         X1              =   12210
         X2              =   15390
         Y1              =   7710
         Y2              =   7710
      End
      Begin VB.Label Label25 
         Caption         =   "漏点描述"
         Height          =   195
         Left            =   3150
         TabIndex        =   91
         Top             =   8250
         Width           =   885
      End
      Begin VB.Line Line23 
         X1              =   4170
         X2              =   15390
         Y1              =   8430
         Y2              =   8430
      End
   End
   Begin VB.ComboBox comXmmc 
      Height          =   300
      Left            =   1680
      TabIndex        =   22
      Top             =   690
      Width           =   4155
   End
   Begin MSDataGridLib.DataGrid comHtbh 
      Height          =   1155
      Left            =   5760
      TabIndex        =   23
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
   Begin VB.CommandButton cmdQm 
      Caption         =   "cmdQm"
      Height          =   345
      Index           =   0
      Left            =   9930
      TabIndex        =   16
      Top             =   10260
      Width           =   945
   End
   Begin VB.CommandButton cmdMod 
      Height          =   375
      Left            =   13590
      Picture         =   "NewGZD.frx":02B9
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "修改"
      Top             =   10260
      Width           =   465
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   14070
      Picture         =   "NewGZD.frx":05C3
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "保存"
      Top             =   10260
      Width           =   465
   End
   Begin VB.CommandButton cmdBack 
      Height          =   375
      Left            =   14550
      Picture         =   "NewGZD.frx":0C2D
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "返回"
      Top             =   10260
      Width           =   465
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
      Left            =   7590
      TabIndex        =   12
      Text            =   "的"
      Top             =   1380
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
      Height          =   210
      Index           =   5
      Left            =   7800
      TabIndex        =   9
      Top             =   510
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
      Left            =   7800
      TabIndex        =   8
      Text            =   "的"
      Top             =   90
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
      Left            =   1710
      TabIndex        =   7
      Top             =   870
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
      Left            =   1710
      TabIndex        =   6
      Top             =   450
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
      Left            =   7890
      TabIndex        =   5
      Top             =   1140
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
      Left            =   1710
      TabIndex        =   4
      Top             =   30
      Width           =   4065
   End
   Begin VB.TextBox TA 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   10200
      TabIndex        =   3
      Top             =   1170
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CheckBox C1 
      Alignment       =   1  'Right Justify
      Caption         =   "1号"
      Height          =   285
      Index           =   0
      Left            =   11310
      TabIndex        =   2
      Top             =   660
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
      Left            =   6270
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "NewGZD.frx":0D2F
      Top             =   90
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
      Height          =   1515
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "NewGZD.frx":0D58
      Top             =   30
      Width           =   1365
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
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   10
      Tag             =   "20"
      Top             =   900
      Width           =   4125
   End
   Begin MSComCtl2.DTPicker dtpA 
      Height          =   225
      Left            =   7800
      TabIndex        =   19
      Top             =   900
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   397
      _Version        =   393216
      Format          =   149094401
      CurrentDate     =   38897
   End
   Begin VB.Label LBLKjj 
      Caption         =   $"NewGZD.frx":0D8F
      ForeColor       =   &H8000000D&
      Height          =   1305
      Left            =   12480
      TabIndex        =   207
      Top             =   270
      Width           =   2835
   End
   Begin VB.Line Line38 
      X1              =   1680
      X2              =   5835
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Label Label39 
      Caption         =   "NO:"
      Height          =   255
      Left            =   12540
      TabIndex        =   143
      Top             =   30
      Width           =   495
   End
   Begin VB.Label lblkhdh 
      Caption         =   "lblkhdh"
      Height          =   225
      Left            =   8940
      TabIndex        =   21
      Top             =   1260
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label lblGid 
      Caption         =   "lblGid"
      Height          =   225
      Left            =   11160
      TabIndex        =   20
      Top             =   1290
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblQM 
      Caption         =   "签字提交"
      Height          =   225
      Index           =   0
      Left            =   9000
      TabIndex        =   18
      Top             =   10320
      Width           =   795
   End
   Begin VB.Label lblTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   10950
      TabIndex        =   17
      Top             =   10320
      Width           =   1905
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
      Left            =   13110
      TabIndex        =   11
      Top             =   30
      Width           =   1605
   End
   Begin VB.Line Line32 
      X1              =   7590
      X2              =   11670
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Line Line31 
      X1              =   7800
      X2              =   11880
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line30 
      X1              =   7800
      X2              =   11880
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Line Line29 
      X1              =   1770
      X2              =   5850
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line28 
      X1              =   1710
      X2              =   5790
      Y1              =   1095
      Y2              =   1095
   End
   Begin VB.Line Line27 
      X1              =   1710
      X2              =   5790
      Y1              =   675
      Y2              =   675
   End
   Begin VB.Line Line26 
      X1              =   1710
      X2              =   5805
      Y1              =   270
      Y2              =   270
   End
End
Attribute VB_Name = "NewGZD1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoRen As ADODB.Recordset



Private Sub BA_DblClick(Index As Integer)
Dim tt As String
Dim oo As Integer
On Error Resume Next
If Index = 1 Then
    If BA(2).Text <> "" Then
        tt = "select 合同编号,合同金额,khdh from htView where 项目名称='" & Trim(BA(2).Text) & _
        "' and 状态='执行' and (合同性质='大修' or 合同性质='D. 维修合同' or 合同性质='C. 维保合同' or 合同性质='维保') order by 合同日期 desc "
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

Private Sub BA_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
If Len(BA(Index).Text) >= BA(Index).Tag And Len(BA(Index)) > 0 And IsNull(BA(Index).Tag) = False Then
    MsgBox ("字数超过限制,超过部分将不被保存!")
End If
End Sub


Private Sub cmdAll_Click()
Dim oo As Integer
If C1(13).Value = 1 Then
    For oo = 13 To 26
        C1(oo).Value = 0
    Next
Else
    For oo = 13 To 26
        C1(oo).Value = 1
    Next
End If
End Sub

Private Sub cmdBack_Click()
Me.Visible = False
frmGZDBR.Enabled = True
frmGZDBR.ZOrder 0
End Sub

Private Sub cmdMod_Click()
cmdSave.Enabled = True
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
For oo = 1 To 60
    mod1.HTP.Update "mat" & oo, TA(oo).Text
Next
For oo = 1 To 63
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


Private Sub comHtbh_DblClick()
On Error Resume Next
BA(1).Text = mod1.HTP.Fields("合同编号").Value
lblkhdh.Caption = mod1.HTP.Fields("khdh").Value
End Sub

Private Sub comXmmc_Click()
BA(2).Text = comXmmc.Text
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

Private Sub Form_Click()
comXmmc.Visible = False
comHtbh.Visible = False
dtgRen.Visible = False
End Sub
Private Sub BA_Click(Index As Integer)
dtgRen.Visible = False
comHtbh.Visible = False
comXmmc.Visible = False
End Sub

Private Sub Ta_Change(Index As Integer)
If Len(TA(Index).Text) >= TA(Index).Tag Then
    MsgBox ("字数超过限制,超过部分将不被保存!")
End If
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

Private Sub Form_Load()
Dim oo As Integer
On Error Resume Next
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
BA(8).Tag = 10
BA(9).Tag = 10
BA(10).Tag = 10
BA(11).Tag = 10
BA(12).Tag = 100
BA(13).Tag = 50
BA(14).Tag = 50
BA(15).Tag = 50
BA(16).Tag = 50
BA(17).Tag = 50
For oo = 1 To 60
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
    tt = "select username,gzu from worker where zzf=1 and (bm='工程部' or bm='工程二部') and qy ='" & mod1.Qy & "' order by gzu"
ElseIf mod1.comId = 1 Then
    tt = "select username,gzu from worker where zzf=1 and bm='广州工程部' order by gzu"
End If
Set adoRen = New ADODB.Recordset
adoRen.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set dtgRen.DataSource = adoRen


End Sub

Private Sub TA_Click(Index As Integer)
comXmmc.Visible = False
comHtbh.Visible = False
dtgRen.Visible = False
End Sub


Private Sub TA_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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
