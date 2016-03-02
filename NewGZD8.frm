VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form NewGZD8 
   Caption         =   "工程部维修质量监督报告（单）"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15015
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9165
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
      Left            =   1470
      TabIndex        =   105
      Top             =   1260
      Width           =   4065
   End
   Begin VB.ComboBox txtDren 
      Height          =   300
      ItemData        =   "NewGZD8.frx":0000
      Left            =   7590
      List            =   "NewGZD8.frx":000A
      TabIndex        =   101
      Top             =   1290
      Width           =   4425
   End
   Begin VB.ComboBox comXmmc 
      Height          =   300
      Left            =   1470
      TabIndex        =   2
      Top             =   720
      Width           =   4125
   End
   Begin MSDataGridLib.DataGrid dtgRen 
      Height          =   8085
      Left            =   12060
      TabIndex        =   100
      Top             =   -30
      Visible         =   0   'False
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
   Begin VB.CommandButton cmdQm 
      Caption         =   "cmdQm"
      Height          =   345
      Index           =   0
      Left            =   9900
      TabIndex        =   97
      Top             =   8790
      Width           =   945
   End
   Begin VB.CommandButton cmdMod 
      Height          =   375
      Left            =   13560
      Picture         =   "NewGZD8.frx":001A
      Style           =   1  'Graphical
      TabIndex        =   96
      ToolTipText     =   "修改"
      Top             =   8790
      Width           =   465
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   14040
      Picture         =   "NewGZD8.frx":0324
      Style           =   1  'Graphical
      TabIndex        =   95
      ToolTipText     =   "保存"
      Top             =   8790
      Width           =   465
   End
   Begin VB.CommandButton cmdBack 
      Height          =   375
      Left            =   14520
      Picture         =   "NewGZD8.frx":098E
      Style           =   1  'Graphical
      TabIndex        =   94
      ToolTipText     =   "返回"
      Top             =   8790
      Width           =   465
   End
   Begin TabDlg.SSTab tabNr 
      Height          =   6675
      Left            =   0
      TabIndex        =   107
      Top             =   2070
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   11774
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "内容1"
      TabPicture(0)   =   "NewGZD8.frx":0A90
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line2(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Line4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label5"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Line2(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label11"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label13"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label15"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "C1(38)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "C1(37)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "C1(36)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "C1(35)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "C1(34)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "C1(33)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "C1(32)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "C1(31)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "C1(30)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "C1(29)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "C1(28)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "C1(27)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "C1(26)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "C1(25)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "C1(24)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "C1(23)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "C1(22)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "C1(21)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "C1(20)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "C1(19)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "TA(2)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "TA(1)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "C1(18)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "C1(17)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "C1(16)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "C1(15)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "C1(14)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "C1(13)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "C1(12)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "C1(11)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "C1(10)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "C1(9)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "C1(8)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "C1(7)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "C1(6)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "C1(5)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "C1(4)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "C1(3)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "C1(2)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "C1(1)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "cmdAll"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).ControlCount=   57
      TabCaption(1)   =   "内容2"
      TabPicture(1)   =   "NewGZD8.frx":0AAC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FPA"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FPB"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "FPC"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "FPD"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "BA(14)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "BA(8)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "BA(9)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "BA(10)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "BA(11)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "BA(12)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "BA(13)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "BA(15)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "BA(16)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "TA(62)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "TA(61)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "TA(60)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "TA(59)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "TA(58)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "TA(57)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "TA(56)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "TA(55)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "TA(54)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "TA(53)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "TA(52)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "TA(51)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "TA(50)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "TA(49)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "TA(48)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "TA(47)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "TA(46)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "TA(45)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "TA(44)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "TA(43)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "TA(42)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "TA(41)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "TA(40)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "TA(39)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "TA(38)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "TA(37)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "TA(36)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "TA(35)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "TA(34)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "TA(33)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "TA(32)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "TA(31)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "TA(30)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "TA(29)"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "TA(28)"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "TA(27)"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "TA(26)"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "TA(25)"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "TA(24)"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "TA(23)"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "TA(22)"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "TA(21)"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "TA(20)"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "TA(19)"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).Control(57)=   "TA(18)"
      Tab(1).Control(57).Enabled=   0   'False
      Tab(1).Control(58)=   "TA(17)"
      Tab(1).Control(58).Enabled=   0   'False
      Tab(1).Control(59)=   "TA(16)"
      Tab(1).Control(59).Enabled=   0   'False
      Tab(1).Control(60)=   "TA(15)"
      Tab(1).Control(60).Enabled=   0   'False
      Tab(1).Control(61)=   "TA(14)"
      Tab(1).Control(61).Enabled=   0   'False
      Tab(1).Control(62)=   "TA(13)"
      Tab(1).Control(62).Enabled=   0   'False
      Tab(1).Control(63)=   "TA(12)"
      Tab(1).Control(63).Enabled=   0   'False
      Tab(1).Control(64)=   "TA(11)"
      Tab(1).Control(64).Enabled=   0   'False
      Tab(1).Control(65)=   "TA(10)"
      Tab(1).Control(65).Enabled=   0   'False
      Tab(1).Control(66)=   "TA(9)"
      Tab(1).Control(66).Enabled=   0   'False
      Tab(1).Control(67)=   "TA(8)"
      Tab(1).Control(67).Enabled=   0   'False
      Tab(1).Control(68)=   "TA(7)"
      Tab(1).Control(68).Enabled=   0   'False
      Tab(1).Control(69)=   "TA(6)"
      Tab(1).Control(69).Enabled=   0   'False
      Tab(1).Control(70)=   "TA(5)"
      Tab(1).Control(70).Enabled=   0   'False
      Tab(1).Control(71)=   "TA(4)"
      Tab(1).Control(71).Enabled=   0   'False
      Tab(1).Control(72)=   "TA(3)"
      Tab(1).Control(72).Enabled=   0   'False
      Tab(1).Control(73)=   "dtpB"
      Tab(1).Control(73).Enabled=   0   'False
      Tab(1).Control(74)=   "dtpC"
      Tab(1).Control(74).Enabled=   0   'False
      Tab(1).Control(75)=   "Shape2"
      Tab(1).Control(75).Enabled=   0   'False
      Tab(1).Control(76)=   "Label33"
      Tab(1).Control(76).Enabled=   0   'False
      Tab(1).Control(77)=   "Label29"
      Tab(1).Control(77).Enabled=   0   'False
      Tab(1).Control(78)=   "Label30"
      Tab(1).Control(78).Enabled=   0   'False
      Tab(1).Control(79)=   "Label31"
      Tab(1).Control(79).Enabled=   0   'False
      Tab(1).Control(80)=   "Label32"
      Tab(1).Control(80).Enabled=   0   'False
      Tab(1).Control(81)=   "Label34"
      Tab(1).Control(81).Enabled=   0   'False
      Tab(1).Control(82)=   "Label35"
      Tab(1).Control(82).Enabled=   0   'False
      Tab(1).Control(83)=   "Label36"
      Tab(1).Control(83).Enabled=   0   'False
      Tab(1).Control(84)=   "Label37"
      Tab(1).Control(84).Enabled=   0   'False
      Tab(1).Control(85)=   "Line34"
      Tab(1).Control(85).Enabled=   0   'False
      Tab(1).Control(86)=   "Line35"
      Tab(1).Control(86).Enabled=   0   'False
      Tab(1).Control(87)=   "Line36"
      Tab(1).Control(87).Enabled=   0   'False
      Tab(1).Control(88)=   "Line37"
      Tab(1).Control(88).Enabled=   0   'False
      Tab(1).Control(89)=   "Label38"
      Tab(1).Control(89).Enabled=   0   'False
      Tab(1).Control(90)=   "Line11(4)"
      Tab(1).Control(90).Enabled=   0   'False
      Tab(1).Control(91)=   "Line10(4)"
      Tab(1).Control(91).Enabled=   0   'False
      Tab(1).Control(92)=   "Line11(3)"
      Tab(1).Control(92).Enabled=   0   'False
      Tab(1).Control(93)=   "Line10(3)"
      Tab(1).Control(93).Enabled=   0   'False
      Tab(1).Control(94)=   "Line11(2)"
      Tab(1).Control(94).Enabled=   0   'False
      Tab(1).Control(95)=   "Line10(2)"
      Tab(1).Control(95).Enabled=   0   'False
      Tab(1).Control(96)=   "Line11(1)"
      Tab(1).Control(96).Enabled=   0   'False
      Tab(1).Control(97)=   "Line10(1)"
      Tab(1).Control(97).Enabled=   0   'False
      Tab(1).Control(98)=   "Line11(0)"
      Tab(1).Control(98).Enabled=   0   'False
      Tab(1).Control(99)=   "Line10(0)"
      Tab(1).Control(99).Enabled=   0   'False
      Tab(1).Control(100)=   "Line9"
      Tab(1).Control(100).Enabled=   0   'False
      Tab(1).Control(101)=   "Line8"
      Tab(1).Control(101).Enabled=   0   'False
      Tab(1).Control(102)=   "Line7"
      Tab(1).Control(102).Enabled=   0   'False
      Tab(1).Control(103)=   "Line6"
      Tab(1).Control(103).Enabled=   0   'False
      Tab(1).Control(104)=   "Shape1"
      Tab(1).Control(104).Enabled=   0   'False
      Tab(1).Control(105)=   "Line5"
      Tab(1).Control(105).Enabled=   0   'False
      Tab(1).Control(106)=   "Label20"
      Tab(1).Control(106).Enabled=   0   'False
      Tab(1).Control(107)=   "Label19"
      Tab(1).Control(107).Enabled=   0   'False
      Tab(1).Control(108)=   "Label18"
      Tab(1).Control(108).Enabled=   0   'False
      Tab(1).Control(109)=   "Label17"
      Tab(1).Control(109).Enabled=   0   'False
      Tab(1).Control(110)=   "Label16"
      Tab(1).Control(110).Enabled=   0   'False
      Tab(1).Control(111)=   "Label14"
      Tab(1).Control(111).Enabled=   0   'False
      Tab(1).Control(112)=   "Label12"
      Tab(1).Control(112).Enabled=   0   'False
      Tab(1).Control(113)=   "Line2(2)"
      Tab(1).Control(113).Enabled=   0   'False
      Tab(1).ControlCount=   114
      Begin VB.CommandButton cmdAll 
         Caption         =   "全部"
         Height          =   225
         Left            =   6930
         TabIndex        =   175
         Top             =   3420
         Width           =   645
      End
      Begin VB.OptionButton FPA 
         Caption         =   "优秀"
         Height          =   195
         Left            =   -73290
         TabIndex        =   92
         Top             =   6900
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.OptionButton FPB 
         Caption         =   "满意"
         Height          =   195
         Left            =   -71685
         TabIndex        =   91
         Top             =   6900
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.OptionButton FPC 
         Caption         =   "较满意"
         Height          =   195
         Left            =   -70095
         TabIndex        =   90
         Top             =   6900
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.OptionButton FPD 
         Caption         =   "尚可"
         Height          =   195
         Left            =   -68490
         TabIndex        =   89
         Top             =   6900
         Visible         =   0   'False
         Width           =   1065
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
         Left            =   -64110
         TabIndex        =   77
         Top             =   5310
         Width           =   1725
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   8
         Left            =   -73110
         TabIndex        =   168
         Text            =   "的"
         Top             =   4500
         Width           =   1035
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   9
         Left            =   -70320
         TabIndex        =   169
         Text            =   "的"
         Top             =   4500
         Width           =   1035
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   10
         Left            =   -67650
         TabIndex        =   170
         Text            =   "的"
         Top             =   4500
         Width           =   1035
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   11
         Left            =   -65250
         TabIndex        =   171
         Text            =   "的"
         Top             =   4500
         Width           =   1035
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   840
         Index           =   12
         Left            =   -73560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   172
         Text            =   "NewGZD8.frx":0AC8
         Top             =   4800
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
         Left            =   -64110
         TabIndex        =   173
         Top             =   4740
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
         Left            =   -62040
         TabIndex        =   174
         Top             =   4740
         Width           =   1275
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
         Left            =   -62070
         TabIndex        =   76
         Top             =   5310
         Width           =   1695
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   62
         Left            =   -62670
         TabIndex        =   113
         Top             =   750
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   61
         Left            =   -62670
         TabIndex        =   119
         Top             =   1095
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   60
         Left            =   -62670
         TabIndex        =   125
         Top             =   1425
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   59
         Left            =   -62670
         TabIndex        =   131
         Top             =   1770
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   58
         Left            =   -62670
         TabIndex        =   137
         Top             =   2115
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   57
         Left            =   -62670
         TabIndex        =   143
         Top             =   2445
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   56
         Left            =   -62670
         TabIndex        =   149
         Top             =   2790
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   55
         Left            =   -62670
         TabIndex        =   155
         Top             =   3135
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   54
         Left            =   -62670
         TabIndex        =   161
         Top             =   3465
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   53
         Left            =   -62670
         TabIndex        =   167
         Top             =   3810
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   52
         Left            =   -65460
         TabIndex        =   112
         Top             =   750
         Width           =   2625
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   51
         Left            =   -65460
         TabIndex        =   118
         Top             =   1095
         Width           =   2625
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   50
         Left            =   -65460
         TabIndex        =   124
         Top             =   1425
         Width           =   2625
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   49
         Left            =   -65460
         TabIndex        =   130
         Top             =   1770
         Width           =   2625
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   48
         Left            =   -65460
         TabIndex        =   136
         Top             =   2115
         Width           =   2625
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   47
         Left            =   -65460
         TabIndex        =   142
         Top             =   2445
         Width           =   2625
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   46
         Left            =   -65460
         TabIndex        =   148
         Top             =   2790
         Width           =   2625
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   45
         Left            =   -65460
         TabIndex        =   154
         Top             =   3135
         Width           =   2625
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   44
         Left            =   -65460
         TabIndex        =   160
         Top             =   3465
         Width           =   2625
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   43
         Left            =   -65460
         TabIndex        =   166
         Top             =   3810
         Width           =   2625
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   42
         Left            =   -68280
         TabIndex        =   111
         Top             =   750
         Width           =   2745
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   41
         Left            =   -68280
         TabIndex        =   117
         Top             =   1095
         Width           =   2745
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   40
         Left            =   -68280
         TabIndex        =   123
         Top             =   1425
         Width           =   2745
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   39
         Left            =   -68280
         TabIndex        =   129
         Top             =   1770
         Width           =   2745
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   38
         Left            =   -68280
         TabIndex        =   135
         Top             =   2115
         Width           =   2745
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   37
         Left            =   -68280
         TabIndex        =   141
         Top             =   2445
         Width           =   2745
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   36
         Left            =   -68280
         TabIndex        =   147
         Top             =   2790
         Width           =   2745
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   35
         Left            =   -68280
         TabIndex        =   153
         Top             =   3135
         Width           =   2745
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   34
         Left            =   -68280
         TabIndex        =   159
         Top             =   3465
         Width           =   2745
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   33
         Left            =   -68280
         TabIndex        =   165
         Top             =   3810
         Width           =   2745
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   32
         Left            =   -71010
         TabIndex        =   110
         Top             =   750
         Width           =   2655
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   31
         Left            =   -71010
         TabIndex        =   116
         Top             =   1095
         Width           =   2655
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   30
         Left            =   -71010
         TabIndex        =   122
         Top             =   1425
         Width           =   2655
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   29
         Left            =   -71010
         TabIndex        =   128
         Top             =   1770
         Width           =   2655
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   28
         Left            =   -71010
         TabIndex        =   134
         Top             =   2115
         Width           =   2655
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   27
         Left            =   -71010
         TabIndex        =   140
         Top             =   2445
         Width           =   2655
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   26
         Left            =   -71010
         TabIndex        =   146
         Top             =   2790
         Width           =   2655
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   25
         Left            =   -71010
         TabIndex        =   152
         Top             =   3135
         Width           =   2655
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   24
         Left            =   -71010
         TabIndex        =   158
         Top             =   3465
         Width           =   2655
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   23
         Left            =   -71010
         TabIndex        =   164
         Top             =   3810
         Width           =   2655
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   22
         Left            =   -73650
         TabIndex        =   109
         Top             =   750
         Width           =   2535
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   21
         Left            =   -73650
         TabIndex        =   115
         Top             =   1095
         Width           =   2535
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   20
         Left            =   -73650
         TabIndex        =   121
         Top             =   1425
         Width           =   2535
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   19
         Left            =   -73650
         TabIndex        =   127
         Top             =   1770
         Width           =   2535
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   18
         Left            =   -73650
         TabIndex        =   133
         Top             =   2115
         Width           =   2535
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   17
         Left            =   -73650
         TabIndex        =   139
         Top             =   2445
         Width           =   2535
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   16
         Left            =   -73650
         TabIndex        =   145
         Top             =   2790
         Width           =   2535
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   15
         Left            =   -73650
         TabIndex        =   151
         Top             =   3135
         Width           =   2535
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   14
         Left            =   -73650
         TabIndex        =   157
         Top             =   3465
         Width           =   2535
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   13
         Left            =   -73650
         TabIndex        =   163
         Top             =   3810
         Width           =   2535
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   12
         Left            =   -74730
         TabIndex        =   162
         Top             =   3810
         Width           =   735
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   11
         Left            =   -74730
         TabIndex        =   156
         Top             =   3470
         Width           =   735
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   10
         Left            =   -74730
         TabIndex        =   150
         Top             =   3130
         Width           =   735
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   9
         Left            =   -74730
         TabIndex        =   144
         Top             =   2790
         Width           =   735
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   8
         Left            =   -74730
         TabIndex        =   138
         Top             =   2450
         Width           =   735
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   7
         Left            =   -74730
         TabIndex        =   132
         Top             =   2110
         Width           =   735
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   6
         Left            =   -74730
         TabIndex        =   126
         Top             =   1770
         Width           =   735
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   5
         Left            =   -74730
         TabIndex        =   120
         Top             =   1430
         Width           =   735
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   4
         Left            =   -74730
         TabIndex        =   114
         Top             =   1090
         Width           =   735
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   3
         Left            =   -74730
         TabIndex        =   108
         Top             =   750
         Width           =   735
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "预约"
         Height          =   285
         Index           =   1
         Left            =   7530
         TabIndex        =   57
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "准时"
         Height          =   285
         Index           =   2
         Left            =   7530
         TabIndex        =   56
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "热情"
         Height          =   285
         Index           =   3
         Left            =   7530
         TabIndex        =   55
         Top             =   1560
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "没有"
         Height          =   285
         Index           =   4
         Left            =   7530
         TabIndex        =   54
         Top             =   1920
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "遵守"
         Height          =   285
         Index           =   5
         Left            =   7530
         TabIndex        =   53
         Top             =   2280
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "清理"
         Height          =   285
         Index           =   6
         Left            =   7530
         TabIndex        =   52
         Top             =   2640
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "齐全"
         Height          =   285
         Index           =   7
         Left            =   7530
         TabIndex        =   51
         Top             =   3000
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "很好"
         Height          =   285
         Index           =   8
         Left            =   7530
         TabIndex        =   50
         Top             =   3360
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "未预约"
         Height          =   285
         Index           =   9
         Left            =   8790
         TabIndex        =   49
         Top             =   840
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "迟到"
         Height          =   285
         Index           =   10
         Left            =   8790
         TabIndex        =   48
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "态度冷漠"
         Height          =   285
         Index           =   11
         Left            =   8790
         TabIndex        =   47
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已及时报告"
         Height          =   285
         Index           =   12
         Left            =   8790
         TabIndex        =   46
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "有违规行为"
         Height          =   285
         Index           =   13
         Left            =   8790
         TabIndex        =   45
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "未清理干净"
         Height          =   285
         Index           =   14
         Left            =   8790
         TabIndex        =   44
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "未穿工作衣"
         Height          =   285
         Index           =   15
         Left            =   8790
         TabIndex        =   43
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "好"
         Height          =   285
         Index           =   16
         Left            =   8790
         TabIndex        =   42
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "未带工作证"
         Height          =   285
         Index           =   17
         Left            =   10530
         TabIndex        =   41
         Top             =   2940
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "可以"
         Height          =   285
         Index           =   18
         Left            =   10530
         TabIndex        =   40
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   1
         Left            =   12240
         TabIndex        =   106
         Top             =   1230
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   2
         Left            =   12210
         TabIndex        =   39
         Top             =   2340
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "）遵守安全规范？（所违反的规程：劳防用品使用"
         Height          =   285
         Index           =   19
         Left            =   2220
         TabIndex        =   38
         Top             =   4170
         Width           =   4305
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "、安全用电"
         Height          =   285
         Index           =   20
         Left            =   6480
         TabIndex        =   37
         Top             =   4170
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "、搬运吊装作业"
         Height          =   285
         Index           =   21
         Left            =   7830
         TabIndex        =   36
         Top             =   4170
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "、机械设备使用"
         Height          =   285
         Index           =   22
         Left            =   9420
         TabIndex        =   35
         Top             =   4170
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "、登高作业"
         Height          =   285
         Index           =   23
         Left            =   11040
         TabIndex        =   34
         Top             =   4170
         Width           =   1215
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "、动火作业"
         Height          =   285
         Index           =   24
         Left            =   12300
         TabIndex        =   33
         Top             =   4170
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "（是"
         Height          =   285
         Index           =   25
         Left            =   5460
         TabIndex        =   32
         Top             =   4860
         Width           =   675
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "、否"
         Height          =   285
         Index           =   26
         Left            =   6150
         TabIndex        =   31
         Top             =   4860
         Width           =   675
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "维修人员（是"
         Height          =   285
         Index           =   27
         Left            =   60
         TabIndex        =   30
         Top             =   4170
         Width           =   1425
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "维修人员（是"
         Height          =   285
         Index           =   28
         Left            =   60
         TabIndex        =   29
         Top             =   4515
         Width           =   1425
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "维修人员（是"
         Height          =   285
         Index           =   29
         Left            =   60
         TabIndex        =   28
         Top             =   4860
         Width           =   1425
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "维修人员（是"
         Height          =   285
         Index           =   30
         Left            =   60
         TabIndex        =   27
         Top             =   5190
         Width           =   1425
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "维修人员（是"
         Height          =   285
         Index           =   31
         Left            =   60
         TabIndex        =   26
         Top             =   5535
         Width           =   1425
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "维修人员（是"
         Height          =   285
         Index           =   32
         Left            =   60
         TabIndex        =   25
         Top             =   5880
         Width           =   1425
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "、否"
         Height          =   285
         Index           =   33
         Left            =   1500
         TabIndex        =   24
         Top             =   4170
         Width           =   705
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "、否"
         Height          =   285
         Index           =   34
         Left            =   1500
         TabIndex        =   23
         Top             =   4515
         Width           =   705
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "、否"
         Height          =   285
         Index           =   35
         Left            =   1500
         TabIndex        =   22
         Top             =   4860
         Width           =   705
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "、否"
         Height          =   285
         Index           =   36
         Left            =   1500
         TabIndex        =   21
         Top             =   5190
         Width           =   705
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "、否"
         Height          =   285
         Index           =   37
         Left            =   1500
         TabIndex        =   20
         Top             =   5535
         Width           =   705
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "、否"
         Height          =   285
         Index           =   38
         Left            =   1500
         TabIndex        =   19
         Top             =   5880
         Width           =   705
      End
      Begin MSComCtl2.DTPicker dtpB 
         Height          =   225
         Left            =   -64110
         TabIndex        =   78
         Top             =   5310
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   397
         _Version        =   393216
         Format          =   149159937
         CurrentDate     =   38897
      End
      Begin MSComCtl2.DTPicker dtpC 
         Height          =   225
         Left            =   -62070
         TabIndex        =   79
         Top             =   5310
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   397
         _Version        =   393216
         Format          =   149159937
         CurrentDate     =   38897
      End
      Begin VB.Label Label15 
         Caption         =   "）"
         Height          =   315
         Left            =   13650
         TabIndex        =   102
         Top             =   4200
         Width           =   315
      End
      Begin VB.Shape Shape2 
         Height          =   1275
         Left            =   -74970
         Top             =   4410
         Width           =   14955
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
         Left            =   -74580
         TabIndex        =   93
         Top             =   6900
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label Label29 
         Caption         =   "抽查人员到达时间"
         Height          =   165
         Left            =   -74850
         TabIndex        =   88
         Top             =   4500
         Width           =   1515
      End
      Begin VB.Label Label30 
         Caption         =   "完成时间"
         Height          =   165
         Left            =   -71880
         TabIndex        =   87
         Top             =   4500
         Width           =   1035
      End
      Begin VB.Label Label31 
         Caption         =   "旅途时间"
         Height          =   165
         Left            =   -68760
         TabIndex        =   86
         Top             =   4500
         Width           =   1035
      End
      Begin VB.Label Label32 
         Caption         =   "加班工时"
         Height          =   165
         Left            =   -66330
         TabIndex        =   85
         Top             =   4500
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
         Left            =   -74910
         TabIndex        =   84
         Top             =   4860
         Width           =   885
      End
      Begin VB.Label Label35 
         Caption         =   "客户签名："
         Height          =   195
         Left            =   -64110
         TabIndex        =   83
         Top             =   4470
         Width           =   945
      End
      Begin VB.Label Label36 
         Caption         =   "日期："
         Height          =   195
         Left            =   -64110
         TabIndex        =   82
         Top             =   5070
         Width           =   945
      End
      Begin VB.Label Label37 
         Caption         =   "检修主管签名："
         Height          =   165
         Left            =   -62040
         TabIndex        =   81
         Top             =   4500
         Width           =   1275
      End
      Begin VB.Line Line34 
         X1              =   -74970
         X2              =   -60030
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line35 
         X1              =   -64200
         X2              =   -60060
         Y1              =   5010
         Y2              =   5010
      End
      Begin VB.Line Line36 
         X1              =   -64200
         X2              =   -64200
         Y1              =   4440
         Y2              =   5640
      End
      Begin VB.Line Line37 
         X1              =   -62100
         X2              =   -62100
         Y1              =   4440
         Y2              =   5640
      End
      Begin VB.Label Label38 
         Caption         =   "日期："
         Height          =   195
         Left            =   -62040
         TabIndex        =   80
         Top             =   5070
         Width           =   945
      End
      Begin VB.Line Line11 
         Index           =   4
         X1              =   -74910
         X2              =   -60120
         Y1              =   3780
         Y2              =   3780
      End
      Begin VB.Line Line10 
         Index           =   4
         X1              =   -74910
         X2              =   -60120
         Y1              =   3420
         Y2              =   3420
      End
      Begin VB.Line Line11 
         Index           =   3
         X1              =   -74910
         X2              =   -60120
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line10 
         Index           =   3
         X1              =   -74910
         X2              =   -60120
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line11 
         Index           =   2
         X1              =   -74910
         X2              =   -60120
         Y1              =   2430
         Y2              =   2430
      End
      Begin VB.Line Line10 
         Index           =   2
         X1              =   -74910
         X2              =   -60120
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Line Line11 
         Index           =   1
         X1              =   -74910
         X2              =   -60120
         Y1              =   1740
         Y2              =   1740
      End
      Begin VB.Line Line10 
         Index           =   1
         X1              =   -74910
         X2              =   -60120
         Y1              =   1410
         Y2              =   1410
      End
      Begin VB.Line Line11 
         Index           =   0
         X1              =   -74910
         X2              =   -60105
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Line Line10 
         Index           =   0
         X1              =   -74910
         X2              =   -60105
         Y1              =   690
         Y2              =   690
      End
      Begin VB.Line Line9 
         X1              =   -62760
         X2              =   -62760
         Y1              =   390
         Y2              =   4200
      End
      Begin VB.Line Line8 
         X1              =   -65490
         X2              =   -65490
         Y1              =   390
         Y2              =   4200
      End
      Begin VB.Line Line7 
         X1              =   -68310
         X2              =   -68310
         Y1              =   390
         Y2              =   4170
      End
      Begin VB.Line Line6 
         X1              =   -71070
         X2              =   -71070
         Y1              =   390
         Y2              =   4170
      End
      Begin VB.Shape Shape1 
         Height          =   3795
         Left            =   -74910
         Top             =   390
         Width           =   14805
      End
      Begin VB.Line Line5 
         X1              =   -73830
         X2              =   -73830
         Y1              =   390
         Y2              =   4200
      End
      Begin VB.Label Label20 
         Caption         =   "现场工作内容名称"
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
         Left            =   -73170
         TabIndex        =   75
         Top             =   480
         Width           =   1635
      End
      Begin VB.Label Label19 
         Caption         =   "质量评价"
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
         Left            =   -62400
         TabIndex        =   74
         Top             =   480
         Width           =   1515
      End
      Begin VB.Label Label18 
         Caption         =   "目前状况描述"
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
         Left            =   -65010
         TabIndex        =   73
         Top             =   480
         Width           =   2025
      End
      Begin VB.Label Label17 
         Caption         =   "该项工作完成标准检查"
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
         Left            =   -68010
         TabIndex        =   72
         Top             =   480
         Width           =   2205
      End
      Begin VB.Label Label16 
         Caption         =   "施工情况评价"
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
         Left            =   -70290
         TabIndex        =   71
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "编号"
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
         Left            =   -74640
         TabIndex        =   70
         Top             =   480
         Width           =   705
      End
      Begin VB.Label Label13 
         Caption         =   "维修人员服务现场规范检查："
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
         Left            =   120
         TabIndex        =   69
         Top             =   3810
         Width           =   2535
      End
      Begin VB.Label Label12 
         Caption         =   "维修工作技术质量调查"
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
         Left            =   -74850
         TabIndex        =   68
         Top             =   120
         Width           =   2925
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         Index           =   2
         X1              =   -74940
         X2              =   -60810
         Y1              =   30
         Y2              =   30
      End
      Begin VB.Label Label11 
         Caption         =   "）完好？"
         Height          =   225
         Left            =   6840
         TabIndex        =   67
         Top             =   4890
         Width           =   765
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         Index           =   0
         X1              =   60
         X2              =   14190
         Y1              =   30
         Y2              =   30
      End
      Begin VB.Label Label2 
         Caption         =   $"NewGZD8.frx":0ACB
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   120
         TabIndex        =   66
         Top             =   150
         Width           =   2625
      End
      Begin VB.Label Label3 
         Caption         =   $"NewGZD8.frx":0AF6
         Height          =   2835
         Left            =   120
         TabIndex        =   65
         Top             =   810
         Width           =   6975
      End
      Begin VB.Label Label4 
         Caption         =   "迟到具体时间："
         Height          =   195
         Left            =   10530
         TabIndex        =   64
         Top             =   1230
         Width           =   1395
      End
      Begin VB.Label Label5 
         Caption         =   "违规行为："
         Height          =   225
         Left            =   10530
         TabIndex        =   63
         Top             =   2340
         Width           =   1275
      End
      Begin VB.Line Line3 
         X1              =   12240
         X2              =   14160
         Y1              =   1410
         Y2              =   1410
      End
      Begin VB.Line Line4 
         X1              =   12210
         X2              =   14130
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         Index           =   1
         X1              =   60
         X2              =   14190
         Y1              =   3690
         Y2              =   3690
      End
      Begin VB.Label Label6 
         Caption         =   "）备齐施工用的零配件、材料？"
         Height          =   225
         Left            =   2220
         TabIndex        =   62
         Top             =   4560
         Width           =   2565
      End
      Begin VB.Label Label7 
         Caption         =   "）携带工具箱或必要的工具、设备？工具"
         Height          =   225
         Left            =   2220
         TabIndex        =   61
         Top             =   4905
         Width           =   3435
      End
      Begin VB.Label Label8 
         Caption         =   "）按规范使用工具？"
         Height          =   225
         Left            =   2220
         TabIndex        =   60
         Top             =   5235
         Width           =   3045
      End
      Begin VB.Label Label9 
         Caption         =   "）按规范流程施工？"
         Height          =   255
         Left            =   2220
         TabIndex        =   59
         Top             =   5580
         Width           =   3375
      End
      Begin VB.Label Label10 
         Caption         =   "）在工作前了解该项目合同附件（技术条款）？"
         Height          =   255
         Left            =   2220
         TabIndex        =   58
         Top             =   5940
         Width           =   4005
      End
   End
   Begin MSDataGridLib.DataGrid comHtbh 
      Height          =   1155
      Left            =   5550
      TabIndex        =   0
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
   Begin VB.TextBox txtDGid 
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
      Left            =   1890
      TabIndex        =   18
      Top             =   1650
      Width           =   3675
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
      TabIndex        =   13
      Text            =   "的"
      Top             =   540
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
      Index           =   6
      Left            =   7590
      Locked          =   -1  'True
      TabIndex        =   12
      Tag             =   "20"
      Top             =   930
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
      Index           =   5
      Left            =   12420
      TabIndex        =   11
      Top             =   1410
      Visible         =   0   'False
      Width           =   2445
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
      Left            =   7590
      TabIndex        =   10
      Text            =   "的"
      Top             =   120
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
      Index           =   3
      Left            =   1500
      TabIndex        =   9
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
      Index           =   2
      Left            =   1500
      TabIndex        =   8
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
      Height          =   270
      Index           =   0
      Left            =   12330
      TabIndex        =   7
      Top             =   1020
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
      Left            =   1500
      TabIndex        =   6
      Top             =   60
      Width           =   4065
   End
   Begin VB.TextBox TA 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   13140
      TabIndex        =   5
      Top             =   390
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CheckBox C1 
      Alignment       =   1  'Right Justify
      Caption         =   "1号"
      Height          =   285
      Index           =   0
      Left            =   12210
      TabIndex        =   4
      Top             =   690
      Visible         =   0   'False
      Width           =   855
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
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "NewGZD8.frx":0C64
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
      Height          =   1875
      Left            =   6210
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "NewGZD8.frx":0C9B
      Top             =   60
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtpA 
      Height          =   195
      Left            =   7590
      TabIndex        =   14
      Top             =   930
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   344
      _Version        =   393216
      Format          =   149159937
      CurrentDate     =   38897
   End
   Begin VB.Label LBLKjj 
      Caption         =   $"NewGZD8.frx":0CD2
      ForeColor       =   &H8000000D&
      Height          =   1305
      Left            =   12210
      TabIndex        =   176
      Top             =   450
      Width           =   2835
   End
   Begin VB.Line Line12 
      X1              =   1860
      X2              =   5610
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line38 
      X1              =   1440
      X2              =   5595
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Label Label39 
      Caption         =   "NO:"
      Height          =   255
      Left            =   12300
      TabIndex        =   104
      Top             =   90
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
      Left            =   12870
      TabIndex        =   103
      Top             =   90
      Width           =   1605
   End
   Begin VB.Label lblQM 
      Caption         =   "签字提交"
      Height          =   225
      Index           =   0
      Left            =   8970
      TabIndex        =   99
      Top             =   8850
      Width           =   795
   End
   Begin VB.Label lblTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   10920
      TabIndex        =   98
      Top             =   8850
      Width           =   1905
   End
   Begin VB.Label Label1 
      Caption         =   "对应工作单编号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   90
      TabIndex        =   17
      Top             =   1650
      Width           =   1725
   End
   Begin VB.Label lblkhdh 
      Caption         =   "lblkhdh"
      Height          =   225
      Left            =   12150
      TabIndex        =   16
      Top             =   60
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label lblGid 
      Caption         =   "lblGid"
      Height          =   225
      Left            =   13260
      TabIndex        =   15
      Top             =   60
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Line Line1 
      X1              =   7590
      X2              =   11685
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Line Line31 
      X1              =   7590
      X2              =   11670
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Line Line30 
      X1              =   7590
      X2              =   11670
      Y1              =   330
      Y2              =   330
   End
   Begin VB.Line Line28 
      X1              =   1500
      X2              =   5580
      Y1              =   1125
      Y2              =   1125
   End
   Begin VB.Line Line27 
      X1              =   1500
      X2              =   5580
      Y1              =   705
      Y2              =   705
   End
   Begin VB.Line Line26 
      X1              =   1500
      X2              =   5595
      Y1              =   300
      Y2              =   300
   End
End
Attribute VB_Name = "NewGZD8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoRen As ADODB.Recordset

Private Sub BA_Click(Index As Integer)
dtgRen.Visible = False
comHtbh.Visible = False
comXmmc.Visible = False
End Sub

Private Sub BA_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
If Len(BA(Index).Text) >= BA(Index).Tag And Len(BA(Index)) > 0 And IsNull(BA(Index).Tag) = False Then
    MsgBox ("字数超过限制,超过部分将不被保存!")
End If
End Sub


Private Sub cmdAll_Click()
If C1(1).Value = 1 Then
    For oo = 1 To 8
        C1(oo).Value = 0
    Next
Else
    For oo = 1 To 8
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
For oo = 1 To 62
    mod1.HTP.Update "mat" & oo, TA(oo).Text
Next
For oo = 1 To 4
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
mod1.HTP.Update "dGid", txtDGid.Text
mod1.HTP.Update "dren", txtDren.Text
    mod1.HTP.UpdateBatch
    cmdSave.Enabled = False
End Sub

Private Sub SSTab1_DblClick()

End Sub

Private Sub dtgRen_DblClick()
If dtgRen.Top = BA(4).Top Then
    BA(4).Text = adoRen.Fields("username").Value
ElseIf dtgRen.Top = BA(5).Top Then
    BA(5).Text = BA(5).Text & " " & adoRen.Fields("username").Value
ElseIf dtgRen.Top = BA(7).Top Then
    BA(7).Text = BA(7).Text & " " & adoRen.Fields("username").Value
End If

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
Me.Height = 9735
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
BA(15).Tag = 50
BA(16).Tag = 50
BA(17).Tag = 50

BA(8).Tag = 10
BA(9).Tag = 10
BA(10).Tag = 10
BA(11).Tag = 10
For oo = 1 To 62
    TA(oo).Tag = 50
Next
TA(33).Tag = 100
TA(34).Tag = 100
TA(39).Tag = 100
TA(40).Tag = 100
TA(61).Tag = 100
TA(62).Tag = 100
'TA(63).Tag = 100
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
            tt = "select xmmc from xmzl where xmmc like '%" & BA(2).Text & "%' order by xmmc"
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
