VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form newGZD9 
   Caption         =   "风冷机组巡视检修工作报告（单）"
   ClientHeight    =   10875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15090
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10875
   ScaleWidth      =   15090
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
      Left            =   12270
      TabIndex        =   14
      Text            =   "的"
      Top             =   60
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
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   13
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
      Index           =   5
      Left            =   7680
      TabIndex        =   12
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
      Height          =   210
      Index           =   4
      Left            =   7680
      TabIndex        =   11
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
      Index           =   3
      Left            =   1590
      TabIndex        =   10
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
      Left            =   1590
      TabIndex        =   9
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
      Left            =   12570
      TabIndex        =   8
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
      Index           =   1
      Left            =   1590
      TabIndex        =   7
      Top             =   60
      Width           =   4065
   End
   Begin VB.TextBox TA 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   13350
      TabIndex        =   6
      Top             =   540
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CheckBox C1 
      Alignment       =   1  'Right Justify
      Caption         =   "1号"
      Height          =   285
      Index           =   0
      Left            =   12450
      TabIndex        =   5
      Top             =   930
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
      Left            =   6150
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "newGZD9.frx":0000
      Top             =   120
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
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "newGZD9.frx":0029
      Top             =   60
      Width           =   1365
   End
   Begin VB.ComboBox comXmmc 
      Height          =   300
      Left            =   1560
      TabIndex        =   2
      Top             =   690
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
      Height          =   240
      Index           =   17
      Left            =   1620
      TabIndex        =   0
      Top             =   1290
      Width           =   4065
   End
   Begin MSDataGridLib.DataGrid comHtbh 
      Height          =   1155
      Left            =   6090
      TabIndex        =   1
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
   Begin MSComCtl2.DTPicker dtpA 
      Height          =   225
      Left            =   7680
      TabIndex        =   15
      Top             =   930
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   397
      _Version        =   393216
      Format          =   55836673
      CurrentDate     =   38897
   End
   Begin MSDataGridLib.DataGrid dtgRen 
      Height          =   8085
      Left            =   11430
      TabIndex        =   21
      Top             =   -330
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
      TabIndex        =   22
      Top             =   1650
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   15372
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "内容1"
      TabPicture(0)   =   "newGZD9.frx":0060
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line18"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line17"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line16"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line15"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Line14"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line13"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line12"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line11"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Line10"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line9"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Line8"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Line7"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Line6"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Line5"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Line4"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Line3"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Line2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Shape1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label19"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label18"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label17"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label16"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label15"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label14"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label13"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label12"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label11"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label10"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label9"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label8"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label7"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label6"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label5"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label4"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Label3"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Label2"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Label20"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Label21"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Line19"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Line20"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Label23"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Label22"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Label24"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Line21"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Line22"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Label40"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Label41"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Line39"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "Line40"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Line41"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Line42"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Line43"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Line44"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "Line45"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "Line46"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "Line47"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "Line48"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "Line49"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "Line50"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "cmdAll"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "TA(51)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "TA(50)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "TA(49)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "TA(48)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "TA(47)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "TA(46)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "TA(45)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "TA(44)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "TA(43)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "TA(42)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "TA(41)"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "C1(12)"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "C1(11)"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "C1(10)"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "C1(9)"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "C1(8)"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "C1(7)"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "C1(6)"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "C1(5)"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "TA(40)"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "TA(39)"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "TA(38)"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).Control(84)=   "TA(37)"
      Tab(0).Control(84).Enabled=   0   'False
      Tab(0).Control(85)=   "TA(36)"
      Tab(0).Control(85).Enabled=   0   'False
      Tab(0).Control(86)=   "TA(35)"
      Tab(0).Control(86).Enabled=   0   'False
      Tab(0).Control(87)=   "TA(34)"
      Tab(0).Control(87).Enabled=   0   'False
      Tab(0).Control(88)=   "TA(33)"
      Tab(0).Control(88).Enabled=   0   'False
      Tab(0).Control(89)=   "TA(32)"
      Tab(0).Control(89).Enabled=   0   'False
      Tab(0).Control(90)=   "TA(31)"
      Tab(0).Control(90).Enabled=   0   'False
      Tab(0).Control(91)=   "TA(30)"
      Tab(0).Control(91).Enabled=   0   'False
      Tab(0).Control(92)=   "TA(29)"
      Tab(0).Control(92).Enabled=   0   'False
      Tab(0).Control(93)=   "TA(28)"
      Tab(0).Control(93).Enabled=   0   'False
      Tab(0).Control(94)=   "TA(27)"
      Tab(0).Control(94).Enabled=   0   'False
      Tab(0).Control(95)=   "TA(26)"
      Tab(0).Control(95).Enabled=   0   'False
      Tab(0).Control(96)=   "TA(25)"
      Tab(0).Control(96).Enabled=   0   'False
      Tab(0).Control(97)=   "TA(24)"
      Tab(0).Control(97).Enabled=   0   'False
      Tab(0).Control(98)=   "TA(23)"
      Tab(0).Control(98).Enabled=   0   'False
      Tab(0).Control(99)=   "TA(22)"
      Tab(0).Control(99).Enabled=   0   'False
      Tab(0).Control(100)=   "TA(21)"
      Tab(0).Control(100).Enabled=   0   'False
      Tab(0).Control(101)=   "TA(20)"
      Tab(0).Control(101).Enabled=   0   'False
      Tab(0).Control(102)=   "TA(19)"
      Tab(0).Control(102).Enabled=   0   'False
      Tab(0).Control(103)=   "TA(18)"
      Tab(0).Control(103).Enabled=   0   'False
      Tab(0).Control(104)=   "TA(17)"
      Tab(0).Control(104).Enabled=   0   'False
      Tab(0).Control(105)=   "TA(16)"
      Tab(0).Control(105).Enabled=   0   'False
      Tab(0).Control(106)=   "TA(15)"
      Tab(0).Control(106).Enabled=   0   'False
      Tab(0).Control(107)=   "TA(14)"
      Tab(0).Control(107).Enabled=   0   'False
      Tab(0).Control(108)=   "TA(13)"
      Tab(0).Control(108).Enabled=   0   'False
      Tab(0).Control(109)=   "TA(12)"
      Tab(0).Control(109).Enabled=   0   'False
      Tab(0).Control(110)=   "TA(11)"
      Tab(0).Control(110).Enabled=   0   'False
      Tab(0).Control(111)=   "TA(10)"
      Tab(0).Control(111).Enabled=   0   'False
      Tab(0).Control(112)=   "TA(9)"
      Tab(0).Control(112).Enabled=   0   'False
      Tab(0).Control(113)=   "TA(8)"
      Tab(0).Control(113).Enabled=   0   'False
      Tab(0).Control(114)=   "TA(7)"
      Tab(0).Control(114).Enabled=   0   'False
      Tab(0).Control(115)=   "TA(6)"
      Tab(0).Control(115).Enabled=   0   'False
      Tab(0).Control(116)=   "TA(5)"
      Tab(0).Control(116).Enabled=   0   'False
      Tab(0).Control(117)=   "TA(4)"
      Tab(0).Control(117).Enabled=   0   'False
      Tab(0).Control(118)=   "TA(3)"
      Tab(0).Control(118).Enabled=   0   'False
      Tab(0).Control(119)=   "TA(2)"
      Tab(0).Control(119).Enabled=   0   'False
      Tab(0).Control(120)=   "TA(1)"
      Tab(0).Control(120).Enabled=   0   'False
      Tab(0).Control(121)=   "C1(4)"
      Tab(0).Control(121).Enabled=   0   'False
      Tab(0).Control(122)=   "C1(3)"
      Tab(0).Control(122).Enabled=   0   'False
      Tab(0).Control(123)=   "C1(2)"
      Tab(0).Control(123).Enabled=   0   'False
      Tab(0).Control(124)=   "C1(1)"
      Tab(0).Control(124).Enabled=   0   'False
      Tab(0).Control(125)=   "C1(13)"
      Tab(0).Control(125).Enabled=   0   'False
      Tab(0).Control(126)=   "C1(14)"
      Tab(0).Control(126).Enabled=   0   'False
      Tab(0).Control(127)=   "C1(15)"
      Tab(0).Control(127).Enabled=   0   'False
      Tab(0).Control(128)=   "C1(16)"
      Tab(0).Control(128).Enabled=   0   'False
      Tab(0).Control(129)=   "C1(17)"
      Tab(0).Control(129).Enabled=   0   'False
      Tab(0).Control(130)=   "C1(18)"
      Tab(0).Control(130).Enabled=   0   'False
      Tab(0).Control(131)=   "C1(19)"
      Tab(0).Control(131).Enabled=   0   'False
      Tab(0).Control(132)=   "C1(20)"
      Tab(0).Control(132).Enabled=   0   'False
      Tab(0).Control(133)=   "C1(21)"
      Tab(0).Control(133).Enabled=   0   'False
      Tab(0).Control(134)=   "C1(22)"
      Tab(0).Control(134).Enabled=   0   'False
      Tab(0).Control(135)=   "C1(23)"
      Tab(0).Control(135).Enabled=   0   'False
      Tab(0).Control(136)=   "C1(28)"
      Tab(0).Control(136).Enabled=   0   'False
      Tab(0).Control(137)=   "C1(29)"
      Tab(0).Control(137).Enabled=   0   'False
      Tab(0).Control(138)=   "C1(30)"
      Tab(0).Control(138).Enabled=   0   'False
      Tab(0).Control(139)=   "C1(31)"
      Tab(0).Control(139).Enabled=   0   'False
      Tab(0).Control(140)=   "C1(32)"
      Tab(0).Control(140).Enabled=   0   'False
      Tab(0).Control(141)=   "C1(33)"
      Tab(0).Control(141).Enabled=   0   'False
      Tab(0).Control(142)=   "C1(34)"
      Tab(0).Control(142).Enabled=   0   'False
      Tab(0).Control(143)=   "C1(35)"
      Tab(0).Control(143).Enabled=   0   'False
      Tab(0).Control(144)=   "C1(36)"
      Tab(0).Control(144).Enabled=   0   'False
      Tab(0).Control(145)=   "C1(37)"
      Tab(0).Control(145).Enabled=   0   'False
      Tab(0).Control(146)=   "C1(42)"
      Tab(0).Control(146).Enabled=   0   'False
      Tab(0).Control(147)=   "C1(43)"
      Tab(0).Control(147).Enabled=   0   'False
      Tab(0).Control(148)=   "C1(44)"
      Tab(0).Control(148).Enabled=   0   'False
      Tab(0).Control(149)=   "C1(45)"
      Tab(0).Control(149).Enabled=   0   'False
      Tab(0).Control(150)=   "C1(46)"
      Tab(0).Control(150).Enabled=   0   'False
      Tab(0).Control(151)=   "C1(47)"
      Tab(0).Control(151).Enabled=   0   'False
      Tab(0).Control(152)=   "C1(48)"
      Tab(0).Control(152).Enabled=   0   'False
      Tab(0).Control(153)=   "C1(49)"
      Tab(0).Control(153).Enabled=   0   'False
      Tab(0).Control(154)=   "C1(50)"
      Tab(0).Control(154).Enabled=   0   'False
      Tab(0).Control(155)=   "C1(51)"
      Tab(0).Control(155).Enabled=   0   'False
      Tab(0).Control(156)=   "TA(52)"
      Tab(0).Control(156).Enabled=   0   'False
      Tab(0).Control(157)=   "TA(53)"
      Tab(0).Control(157).Enabled=   0   'False
      Tab(0).Control(158)=   "C1(54)"
      Tab(0).Control(158).Enabled=   0   'False
      Tab(0).Control(159)=   "C1(55)"
      Tab(0).Control(159).Enabled=   0   'False
      Tab(0).Control(160)=   "C1(56)"
      Tab(0).Control(160).Enabled=   0   'False
      Tab(0).Control(161)=   "C1(57)"
      Tab(0).Control(161).Enabled=   0   'False
      Tab(0).Control(162)=   "C1(58)"
      Tab(0).Control(162).Enabled=   0   'False
      Tab(0).Control(163)=   "C1(59)"
      Tab(0).Control(163).Enabled=   0   'False
      Tab(0).Control(164)=   "C1(60)"
      Tab(0).Control(164).Enabled=   0   'False
      Tab(0).Control(165)=   "C1(68)"
      Tab(0).Control(165).Enabled=   0   'False
      Tab(0).Control(166)=   "C1(69)"
      Tab(0).Control(166).Enabled=   0   'False
      Tab(0).Control(167)=   "C1(70)"
      Tab(0).Control(167).Enabled=   0   'False
      Tab(0).Control(168)=   "C1(71)"
      Tab(0).Control(168).Enabled=   0   'False
      Tab(0).Control(169)=   "C1(72)"
      Tab(0).Control(169).Enabled=   0   'False
      Tab(0).Control(170)=   "C1(73)"
      Tab(0).Control(170).Enabled=   0   'False
      Tab(0).Control(171)=   "C1(74)"
      Tab(0).Control(171).Enabled=   0   'False
      Tab(0).Control(172)=   "C1(75)"
      Tab(0).Control(172).Enabled=   0   'False
      Tab(0).Control(173)=   "C1(76)"
      Tab(0).Control(173).Enabled=   0   'False
      Tab(0).Control(174)=   "C1(77)"
      Tab(0).Control(174).Enabled=   0   'False
      Tab(0).Control(175)=   "C1(78)"
      Tab(0).Control(175).Enabled=   0   'False
      Tab(0).Control(176)=   "C1(79)"
      Tab(0).Control(176).Enabled=   0   'False
      Tab(0).Control(177)=   "C1(80)"
      Tab(0).Control(177).Enabled=   0   'False
      Tab(0).Control(178)=   "TA(61)"
      Tab(0).Control(178).Enabled=   0   'False
      Tab(0).Control(179)=   "TA(62)"
      Tab(0).Control(179).Enabled=   0   'False
      Tab(0).Control(180)=   "TA(63)"
      Tab(0).Control(180).Enabled=   0   'False
      Tab(0).Control(181)=   "TA(64)"
      Tab(0).Control(181).Enabled=   0   'False
      Tab(0).Control(182)=   "TA(65)"
      Tab(0).Control(182).Enabled=   0   'False
      Tab(0).Control(183)=   "TA(66)"
      Tab(0).Control(183).Enabled=   0   'False
      Tab(0).Control(184)=   "TA(67)"
      Tab(0).Control(184).Enabled=   0   'False
      Tab(0).Control(185)=   "TA(68)"
      Tab(0).Control(185).Enabled=   0   'False
      Tab(0).Control(186)=   "TA(69)"
      Tab(0).Control(186).Enabled=   0   'False
      Tab(0).Control(187)=   "TA(70)"
      Tab(0).Control(187).Enabled=   0   'False
      Tab(0).Control(188)=   "TA(71)"
      Tab(0).Control(188).Enabled=   0   'False
      Tab(0).Control(189)=   "TA(72)"
      Tab(0).Control(189).Enabled=   0   'False
      Tab(0).Control(190)=   "TA(73)"
      Tab(0).Control(190).Enabled=   0   'False
      Tab(0).ControlCount=   191
      TabCaption(1)   =   "内容2"
      TabPicture(1)   =   "newGZD9.frx":007C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TA(59)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "TA(57)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "C1(65)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "C1(64)"
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
      Tab(1).Control(15)=   "TA(58)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "BA(14)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "C1(24)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "C1(25)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "C1(26)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "C1(27)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "C1(38)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "C1(39)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "C1(40)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "C1(41)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "C1(52)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "C1(53)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "TA(54)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "C1(61)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "C1(62)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "C1(63)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "TA(55)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "C1(66)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "C1(67)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "TA(56)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "dtpC"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "dtpB"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Label26"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Label27"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "Label28"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "Line29"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "Line25"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "Line33"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "Label29"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "Label30"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "Label31"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "Label32"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "Label34"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "Shape2"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "Label35"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "Label36"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "Label37"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "Line34"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "Line35"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "Line36"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "Line37"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "Label38"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).Control(57)=   "Label42"
      Tab(1).Control(57).Enabled=   0   'False
      Tab(1).Control(58)=   "Line23"
      Tab(1).Control(58).Enabled=   0   'False
      Tab(1).Control(59)=   "Line24"
      Tab(1).Control(59).Enabled=   0   'False
      Tab(1).Control(60)=   "Label25"
      Tab(1).Control(60).Enabled=   0   'False
      Tab(1).Control(61)=   "Line38(0)"
      Tab(1).Control(61).Enabled=   0   'False
      Tab(1).Control(62)=   "Label39(0)"
      Tab(1).Control(62).Enabled=   0   'False
      Tab(1).ControlCount=   63
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   510
         Index           =   59
         Left            =   -73590
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   194
         Top             =   5280
         Width           =   13515
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   73
         Left            =   13650
         TabIndex        =   191
         Top             =   3000
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   72
         Left            =   12690
         TabIndex        =   190
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   71
         Left            =   11520
         TabIndex        =   189
         Top             =   3000
         Width           =   1125
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   70
         Left            =   10350
         TabIndex        =   188
         Top             =   3000
         Width           =   1035
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   69
         Left            =   9360
         TabIndex        =   187
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
         TabIndex        =   186
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   67
         Left            =   7380
         TabIndex        =   185
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   66
         Left            =   6390
         TabIndex        =   184
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   65
         Left            =   5430
         TabIndex        =   183
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   64
         Left            =   4470
         TabIndex        =   182
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   63
         Left            =   3480
         TabIndex        =   181
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   62
         Left            =   2550
         TabIndex        =   180
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   61
         Left            =   1620
         TabIndex        =   179
         Top             =   3000
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   80
         Left            =   14370
         TabIndex        =   178
         Top             =   3000
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "12"
         Height          =   180
         Index           =   79
         Left            =   12840
         TabIndex        =   177
         Top             =   2700
         Width           =   495
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "11"
         Height          =   180
         Index           =   78
         Left            =   11745
         TabIndex        =   176
         Top             =   2700
         Width           =   525
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "10"
         Height          =   180
         Index           =   77
         Left            =   10560
         TabIndex        =   175
         Top             =   2700
         Width           =   615
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "9"
         Height          =   180
         Index           =   76
         Left            =   9585
         TabIndex        =   174
         Top             =   2700
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "8"
         Height          =   180
         Index           =   75
         Left            =   8610
         TabIndex        =   173
         Top             =   2700
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "7"
         Height          =   180
         Index           =   74
         Left            =   7635
         TabIndex        =   172
         Top             =   2700
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "6"
         Height          =   180
         Index           =   73
         Left            =   6660
         TabIndex        =   171
         Top             =   2700
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "5"
         Height          =   180
         Index           =   72
         Left            =   5685
         TabIndex        =   170
         Top             =   2700
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "4"
         Height          =   180
         Index           =   71
         Left            =   4725
         TabIndex        =   169
         Top             =   2700
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "3"
         Height          =   180
         Index           =   70
         Left            =   3750
         TabIndex        =   168
         Top             =   2700
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "2"
         Height          =   180
         Index           =   69
         Left            =   2775
         TabIndex        =   167
         Top             =   2700
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   180
         Index           =   68
         Left            =   1800
         TabIndex        =   166
         Top             =   2700
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   285
         Index           =   60
         Left            =   13440
         TabIndex        =   165
         Top             =   7185
         Width           =   1005
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   285
         Index           =   59
         Left            =   13440
         TabIndex        =   164
         Top             =   6810
         Width           =   1005
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   285
         Index           =   58
         Left            =   13440
         TabIndex        =   163
         Top             =   6450
         Width           =   1005
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   285
         Index           =   57
         Left            =   13440
         TabIndex        =   162
         Top             =   6090
         Width           =   1005
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   285
         Index           =   56
         Left            =   13440
         TabIndex        =   161
         Top             =   5730
         Width           =   1005
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   285
         Index           =   55
         Left            =   13440
         TabIndex        =   160
         Top             =   5355
         Width           =   1005
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   285
         Index           =   54
         Left            =   13440
         TabIndex        =   159
         Top             =   3930
         Width           =   1005
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   53
         Left            =   13080
         TabIndex        =   158
         Top             =   4290
         Width           =   1875
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   52
         Left            =   13110
         TabIndex        =   157
         Top             =   3570
         Width           =   1875
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   285
         Index           =   51
         Left            =   11520
         TabIndex        =   156
         Top             =   7185
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   285
         Index           =   50
         Left            =   11520
         TabIndex        =   155
         Top             =   6810
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   285
         Index           =   49
         Left            =   11520
         TabIndex        =   154
         Top             =   6450
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   285
         Index           =   48
         Left            =   11520
         TabIndex        =   153
         Top             =   6090
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   285
         Index           =   47
         Left            =   11520
         TabIndex        =   152
         Top             =   5730
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   285
         Index           =   46
         Left            =   11520
         TabIndex        =   151
         Top             =   5355
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   285
         Index           =   45
         Left            =   11520
         TabIndex        =   150
         Top             =   4995
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   180
         Index           =   44
         Left            =   11520
         TabIndex        =   149
         Top             =   4335
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   285
         Index           =   43
         Left            =   11520
         TabIndex        =   148
         Top             =   3930
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   285
         Index           =   42
         Left            =   11520
         TabIndex        =   147
         Top             =   3540
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已更换过滤器"
         Height          =   285
         Index           =   37
         Left            =   9360
         TabIndex        =   146
         Top             =   7185
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已校核或检修"
         Height          =   285
         Index           =   36
         Left            =   9360
         TabIndex        =   145
         Top             =   6810
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已做相应调整"
         Height          =   285
         Index           =   35
         Left            =   9360
         TabIndex        =   144
         Top             =   6450
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已调整或检修"
         Height          =   285
         Index           =   34
         Left            =   9360
         TabIndex        =   143
         Top             =   6090
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已调整或检修"
         Height          =   285
         Index           =   33
         Left            =   9360
         TabIndex        =   142
         Top             =   5730
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已检修或更换"
         Height          =   285
         Index           =   32
         Left            =   9360
         TabIndex        =   141
         Top             =   5355
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已做相应调整"
         Height          =   285
         Index           =   31
         Left            =   9360
         TabIndex        =   140
         Top             =   4995
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "有较大波动"
         Height          =   285
         Index           =   30
         Left            =   9360
         TabIndex        =   139
         Top             =   4635
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已检修或更换"
         Height          =   285
         Index           =   29
         Left            =   9360
         TabIndex        =   138
         Top             =   4275
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已校核或检修"
         Height          =   285
         Index           =   28
         Left            =   9360
         TabIndex        =   137
         Top             =   3540
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   23
         Left            =   7740
         TabIndex        =   136
         Top             =   7185
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   22
         Left            =   7740
         TabIndex        =   135
         Top             =   6810
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   21
         Left            =   7740
         TabIndex        =   134
         Top             =   6450
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   20
         Left            =   7740
         TabIndex        =   133
         Top             =   6090
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   19
         Left            =   7740
         TabIndex        =   132
         Top             =   5730
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   18
         Left            =   7740
         TabIndex        =   131
         Top             =   5355
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   17
         Left            =   7740
         TabIndex        =   130
         Top             =   4995
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   16
         Left            =   7740
         TabIndex        =   129
         Top             =   4635
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   15
         Left            =   7740
         TabIndex        =   128
         Top             =   4275
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   14
         Left            =   7740
         TabIndex        =   127
         Top             =   3900
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   13
         Left            =   7740
         TabIndex        =   126
         Top             =   3540
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "1号"
         Height          =   180
         Index           =   1
         Left            =   2010
         TabIndex        =   125
         Top             =   300
         Width           =   585
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "2号"
         Height          =   180
         Index           =   2
         Left            =   3525
         TabIndex        =   124
         Top             =   300
         Width           =   585
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "3号"
         Height          =   180
         Index           =   3
         Left            =   5055
         TabIndex        =   123
         Top             =   300
         Width           =   585
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "4号"
         Height          =   180
         Index           =   4
         Left            =   6570
         TabIndex        =   122
         Top             =   300
         Width           =   585
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   1
         Left            =   1635
         TabIndex        =   121
         Top             =   570
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   2
         Left            =   3180
         TabIndex        =   120
         Top             =   570
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   3
         Left            =   4710
         TabIndex        =   119
         Top             =   570
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   4
         Left            =   6255
         TabIndex        =   118
         Top             =   570
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   5
         Left            =   7800
         TabIndex        =   117
         Top             =   570
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   6
         Left            =   1635
         TabIndex        =   116
         Top             =   840
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   7
         Left            =   3180
         TabIndex        =   115
         Top             =   840
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   8
         Left            =   4710
         TabIndex        =   114
         Top             =   840
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   9
         Left            =   6255
         TabIndex        =   113
         Top             =   840
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   10
         Left            =   7800
         TabIndex        =   112
         Top             =   825
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   11
         Left            =   1635
         TabIndex        =   111
         Top             =   1110
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   12
         Left            =   3180
         TabIndex        =   110
         Top             =   1110
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   13
         Left            =   4710
         TabIndex        =   109
         Top             =   1110
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   14
         Left            =   6255
         TabIndex        =   108
         Top             =   1110
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   15
         Left            =   7800
         TabIndex        =   107
         Top             =   1095
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   16
         Left            =   1635
         TabIndex        =   106
         Top             =   1380
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   17
         Left            =   3180
         TabIndex        =   105
         Top             =   1365
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   18
         Left            =   4710
         TabIndex        =   104
         Top             =   1365
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   19
         Left            =   6255
         TabIndex        =   103
         Top             =   1365
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   20
         Left            =   7800
         TabIndex        =   102
         Top             =   1365
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   21
         Left            =   1635
         TabIndex        =   101
         Top             =   1650
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   22
         Left            =   3180
         TabIndex        =   100
         Top             =   1635
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   23
         Left            =   4710
         TabIndex        =   99
         Top             =   1635
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   24
         Left            =   6255
         TabIndex        =   98
         Top             =   1635
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   25
         Left            =   7800
         TabIndex        =   97
         Top             =   1635
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   26
         Left            =   1635
         TabIndex        =   96
         Top             =   1920
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   27
         Left            =   3180
         TabIndex        =   95
         Top             =   1905
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   28
         Left            =   4710
         TabIndex        =   94
         Top             =   1905
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   29
         Left            =   6255
         TabIndex        =   93
         Top             =   1905
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   30
         Left            =   7800
         TabIndex        =   92
         Top             =   1905
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   31
         Left            =   1635
         TabIndex        =   91
         Top             =   2190
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   32
         Left            =   3180
         TabIndex        =   90
         Top             =   2190
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   33
         Left            =   4710
         TabIndex        =   89
         Top             =   2175
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   34
         Left            =   6255
         TabIndex        =   88
         Top             =   2175
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   35
         Left            =   7800
         TabIndex        =   87
         Top             =   2175
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   36
         Left            =   1635
         TabIndex        =   86
         Top             =   2445
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   37
         Left            =   3180
         TabIndex        =   85
         Top             =   2445
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   38
         Left            =   4710
         TabIndex        =   84
         Top             =   2445
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   39
         Left            =   6255
         TabIndex        =   83
         Top             =   2445
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   40
         Left            =   7800
         TabIndex        =   82
         Top             =   2445
         Width           =   1395
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
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   6
         Left            =   9720
         TabIndex        =   80
         Top             =   840
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   7
         Left            =   9720
         TabIndex        =   79
         Top             =   1110
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   8
         Left            =   9720
         TabIndex        =   78
         Top             =   1365
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   9
         Left            =   9720
         TabIndex        =   77
         Top             =   1635
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   10
         Left            =   9720
         TabIndex        =   76
         Top             =   1905
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   11
         Left            =   9720
         TabIndex        =   75
         Top             =   2175
         Width           =   405
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Height          =   180
         Index           =   12
         Left            =   9720
         TabIndex        =   74
         Top             =   2445
         Width           =   405
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   41
         Left            =   12180
         TabIndex        =   73
         Top             =   600
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   42
         Left            =   13680
         TabIndex        =   72
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   43
         Left            =   12180
         TabIndex        =   71
         Top             =   855
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   44
         Left            =   13680
         TabIndex        =   70
         Top             =   855
         Width           =   1335
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   45
         Left            =   12180
         TabIndex        =   69
         Top             =   1125
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   46
         Left            =   13680
         TabIndex        =   68
         Top             =   1125
         Width           =   1335
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   47
         Left            =   12180
         TabIndex        =   67
         Top             =   1380
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   48
         Left            =   13680
         TabIndex        =   66
         Top             =   1380
         Width           =   1335
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   49
         Left            =   12180
         TabIndex        =   65
         Top             =   1635
         Width           =   1395
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   50
         Left            =   13680
         TabIndex        =   64
         Top             =   1635
         Width           =   1335
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   780
         Index           =   51
         Left            =   12120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   63
         Top             =   1890
         Width           =   2895
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   57
         Left            =   -73590
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   62
         Top             =   3930
         Width           =   13575
      End
      Begin VB.CheckBox C1 
         Caption         =   "完成"
         Height          =   180
         Index           =   65
         Left            =   -62250
         TabIndex        =   61
         Top             =   4470
         Width           =   1005
      End
      Begin VB.CheckBox C1 
         Caption         =   "未完成"
         Height          =   180
         Index           =   64
         Left            =   -61110
         TabIndex        =   60
         Top             =   4470
         Width           =   945
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   -61860
         TabIndex        =   59
         Text            =   "复核人:"
         Top             =   5640
         Width           =   735
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   60
         Left            =   -61230
         TabIndex        =   58
         Top             =   5550
         Width           =   1065
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   8
         Left            =   -73530
         TabIndex        =   57
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
         TabIndex        =   56
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
         TabIndex        =   55
         Text            =   "的"
         Top             =   5880
         Width           =   1035
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   11
         Left            =   -65220
         TabIndex        =   54
         Text            =   "的"
         Top             =   5880
         Width           =   1035
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Left            =   -74940
         TabIndex        =   48
         Top             =   6120
         Width           =   10755
         Begin VB.OptionButton FPA 
            Caption         =   "优秀"
            Height          =   195
            Left            =   1350
            TabIndex        =   52
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton FPB 
            Caption         =   "满意"
            Height          =   195
            Left            =   2950
            TabIndex        =   51
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton FPC 
            Caption         =   "较满意"
            Height          =   195
            Left            =   4550
            TabIndex        =   50
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton FPD 
            Caption         =   "尚可"
            Height          =   195
            Left            =   6150
            TabIndex        =   49
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
            TabIndex        =   53
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
         TabIndex        =   47
         Text            =   "newGZD9.frx":0098
         Top             =   6450
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
         TabIndex        =   46
         Top             =   6120
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
         TabIndex        =   45
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
         Index           =   16
         Left            =   -62040
         TabIndex        =   44
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
         TabIndex        =   43
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
         Height          =   240
         Index           =   14
         Left            =   -64080
         TabIndex        =   42
         Top             =   6690
         Width           =   1755
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   24
         Left            =   -67320
         TabIndex        =   41
         Top             =   150
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   25
         Left            =   -67320
         TabIndex        =   40
         Top             =   510
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   26
         Left            =   -67320
         TabIndex        =   39
         Top             =   870
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "正常"
         Height          =   285
         Index           =   27
         Left            =   -67320
         TabIndex        =   38
         Top             =   1245
         Width           =   855
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已调整或检修"
         Height          =   285
         Index           =   38
         Left            =   -65700
         TabIndex        =   37
         Top             =   150
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已检修或更换"
         Height          =   285
         Index           =   39
         Left            =   -65700
         TabIndex        =   36
         Top             =   510
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "因故未完成清洁"
         Height          =   285
         Index           =   40
         Left            =   -65700
         TabIndex        =   35
         Top             =   870
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "已调整或检修"
         Height          =   285
         Index           =   41
         Left            =   -65700
         TabIndex        =   34
         Top             =   1245
         Width           =   1575
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   285
         Index           =   52
         Left            =   -63540
         TabIndex        =   33
         Top             =   150
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   285
         Index           =   53
         Left            =   -63540
         TabIndex        =   32
         Top             =   510
         Width           =   1335
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   54
         Left            =   -62010
         TabIndex        =   31
         Top             =   255
         Width           =   1875
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   285
         Index           =   61
         Left            =   -61620
         TabIndex        =   30
         Top             =   525
         Width           =   1005
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "部件已损坏"
         Height          =   285
         Index           =   62
         Left            =   -63540
         TabIndex        =   29
         Top             =   1245
         Width           =   1335
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无此项"
         Height          =   285
         Index           =   63
         Left            =   -61620
         TabIndex        =   28
         Top             =   1245
         Width           =   1005
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   55
         Left            =   -62760
         TabIndex        =   27
         Top             =   915
         Width           =   2685
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "仪器检漏"
         Height          =   180
         Index           =   66
         Left            =   -73440
         TabIndex        =   26
         Top             =   1545
         Width           =   1065
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "目测检漏"
         Height          =   180
         Index           =   67
         Left            =   -74940
         TabIndex        =   25
         Top             =   1545
         Width           =   1095
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   56
         Left            =   -71130
         TabIndex        =   24
         Top             =   1530
         Width           =   11175
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "全部"
         Height          =   285
         Left            =   7770
         TabIndex        =   23
         Top             =   7560
         Width           =   915
      End
      Begin MSComCtl2.DTPicker dtpC 
         Height          =   225
         Left            =   -62040
         TabIndex        =   192
         Top             =   6690
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   397
         _Version        =   393216
         Format          =   55836673
         CurrentDate     =   38897
      End
      Begin MSComCtl2.DTPicker dtpB 
         Height          =   225
         Left            =   -64080
         TabIndex        =   193
         Top             =   6690
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   397
         _Version        =   393216
         Format          =   55836673
         CurrentDate     =   38897
      End
      Begin VB.Line Line50 
         X1              =   2490
         X2              =   2490
         Y1              =   2670
         Y2              =   3240
      End
      Begin VB.Line Line49 
         X1              =   3420
         X2              =   3420
         Y1              =   2670
         Y2              =   3225
      End
      Begin VB.Line Line48 
         X1              =   4410
         X2              =   4410
         Y1              =   2670
         Y2              =   3210
      End
      Begin VB.Line Line47 
         X1              =   5400
         X2              =   5400
         Y1              =   2670
         Y2              =   3210
      End
      Begin VB.Line Line46 
         X1              =   6330
         X2              =   6330
         Y1              =   2670
         Y2              =   3210
      End
      Begin VB.Line Line45 
         X1              =   7320
         X2              =   7320
         Y1              =   2670
         Y2              =   3210
      End
      Begin VB.Line Line44 
         X1              =   8310
         X2              =   8310
         Y1              =   2670
         Y2              =   3210
      End
      Begin VB.Line Line43 
         X1              =   9300
         X2              =   9300
         Y1              =   2670
         Y2              =   3225
      End
      Begin VB.Line Line42 
         X1              =   10290
         X2              =   10290
         Y1              =   2670
         Y2              =   3225
      End
      Begin VB.Line Line41 
         X1              =   11430
         X2              =   11430
         Y1              =   2670
         Y2              =   3210
      End
      Begin VB.Line Line40 
         X1              =   12660
         X2              =   12660
         Y1              =   3210
         Y2              =   2670
      End
      Begin VB.Line Line39 
         X1              =   14250
         X2              =   14250
         Y1              =   2670
         Y2              =   3210
      End
      Begin VB.Label Label41 
         Caption         =   "无法测量"
         Height          =   195
         Left            =   14280
         TabIndex        =   235
         Top             =   2730
         Width           =   735
      End
      Begin VB.Label Label40 
         Caption         =   "正常值"
         Height          =   165
         Left            =   13650
         TabIndex        =   234
         Top             =   2730
         Width           =   555
      End
      Begin VB.Line Line22 
         X1              =   13050
         X2              =   14940
         Y1              =   4470
         Y2              =   4470
      End
      Begin VB.Line Line21 
         X1              =   13110
         X2              =   15000
         Y1              =   3750
         Y2              =   3750
      End
      Begin VB.Label Label24 
         Caption         =   $"newGZD9.frx":009B
         Height          =   4485
         Left            =   210
         TabIndex        =   233
         Top             =   3540
         Width           =   6705
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
         TabIndex        =   232
         Top             =   3240
         Width           =   2145
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
         TabIndex        =   231
         Top             =   3240
         Width           =   1395
      End
      Begin VB.Line Line20 
         X1              =   60
         X2              =   15000
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Line Line19 
         X1              =   30
         X2              =   15030
         Y1              =   2670
         Y2              =   2670
      End
      Begin VB.Label Label21 
         Caption         =   "电流A"
         Height          =   195
         Left            =   270
         TabIndex        =   230
         Top             =   2970
         Width           =   1125
      End
      Begin VB.Label Label20 
         Caption         =   "风机编号"
         Height          =   165
         Left            =   270
         TabIndex        =   229
         Top             =   2730
         Width           =   1125
      End
      Begin VB.Label Label2 
         Caption         =   "压缩机编号"
         Height          =   165
         Left            =   270
         TabIndex        =   228
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "排气温度"
         Height          =   165
         Left            =   270
         TabIndex        =   227
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "油压力"
         Height          =   165
         Left            =   270
         TabIndex        =   226
         Top             =   1650
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "吸气压力"
         Height          =   165
         Left            =   270
         TabIndex        =   225
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "吸气温度"
         Height          =   165
         Left            =   270
         TabIndex        =   224
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "排气压力"
         Height          =   165
         Left            =   270
         TabIndex        =   223
         Top             =   1110
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "非运行时油温"
         Height          =   165
         Left            =   270
         TabIndex        =   222
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "负载百分比"
         Height          =   165
         Left            =   270
         TabIndex        =   221
         Top             =   2190
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "电流"
         Height          =   165
         Left            =   270
         TabIndex        =   220
         Top             =   2460
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "正常值"
         Height          =   165
         Left            =   7890
         TabIndex        =   219
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "无此项"
         Height          =   165
         Left            =   9705
         TabIndex        =   218
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "机组参数"
         Height          =   165
         Left            =   12270
         TabIndex        =   217
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "正常值"
         Height          =   165
         Left            =   13800
         TabIndex        =   216
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "冷凝温度"
         Height          =   195
         Left            =   10830
         TabIndex        =   215
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "循环水出水温度"
         Height          =   195
         Left            =   10830
         TabIndex        =   214
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "循环水进水温度"
         Height          =   195
         Left            =   10830
         TabIndex        =   213
         Top             =   1095
         Width           =   1215
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "电压"
         Height          =   195
         Left            =   10830
         TabIndex        =   212
         Top             =   1365
         Width           =   1215
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "环境温度"
         Height          =   195
         Left            =   10830
         TabIndex        =   211
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         Height          =   2985
         Left            =   60
         Top             =   240
         Width           =   14985
      End
      Begin VB.Line Line2 
         X1              =   60
         X2              =   15030
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Line Line3 
         X1              =   60
         X2              =   15030
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Line Line4 
         X1              =   60
         X2              =   15030
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Line Line5 
         X1              =   60
         X2              =   15030
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line6 
         X1              =   60
         X2              =   15030
         Y1              =   1590
         Y2              =   1590
      End
      Begin VB.Line Line7 
         X1              =   60
         X2              =   15030
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Line Line8 
         X1              =   60
         X2              =   15030
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line9 
         X1              =   60
         X2              =   15030
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line10 
         X1              =   1590
         X2              =   1590
         Y1              =   240
         Y2              =   3240
      End
      Begin VB.Line Line11 
         X1              =   3090
         X2              =   3090
         Y1              =   240
         Y2              =   2690
      End
      Begin VB.Line Line12 
         X1              =   4650
         X2              =   4650
         Y1              =   240
         Y2              =   2690
      End
      Begin VB.Line Line13 
         X1              =   6180
         X2              =   6180
         Y1              =   240
         Y2              =   2690
      End
      Begin VB.Line Line14 
         X1              =   7710
         X2              =   7710
         Y1              =   240
         Y2              =   2690
      End
      Begin VB.Line Line15 
         X1              =   9300
         X2              =   9300
         Y1              =   240
         Y2              =   2690
      End
      Begin VB.Line Line16 
         X1              =   10710
         X2              =   10710
         Y1              =   240
         Y2              =   2690
      End
      Begin VB.Line Line17 
         X1              =   12090
         X2              =   12090
         Y1              =   240
         Y2              =   2690
      End
      Begin VB.Line Line18 
         X1              =   13620
         X2              =   13620
         Y1              =   300
         Y2              =   3210
      End
      Begin VB.Label Label1 
         Caption         =   "常规运行参数记录（记录与压缩机或风机对应的数据时，在压缩机或风机编号一栏中相应编号的""□""上打""√""，若无此压缩机则打""／""）"
         Height          =   195
         Left            =   450
         TabIndex        =   210
         Top             =   30
         Width           =   10875
      End
      Begin VB.Line Line1 
         X1              =   60
         X2              =   15105
         Y1              =   0
         Y2              =   0
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
         TabIndex        =   209
         Top             =   3930
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
         TabIndex        =   208
         Top             =   4710
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
         TabIndex        =   207
         Top             =   5340
         Width           =   1035
      End
      Begin VB.Line Line29 
         X1              =   -74970
         X2              =   -59940
         Y1              =   4650
         Y2              =   4650
      End
      Begin VB.Line Line25 
         X1              =   -74970
         X2              =   -59910
         Y1              =   5250
         Y2              =   5250
      End
      Begin VB.Line Line33 
         X1              =   -74970
         X2              =   -60030
         Y1              =   5820
         Y2              =   5820
      End
      Begin VB.Label Label29 
         Caption         =   "到达时间"
         Height          =   165
         Left            =   -74820
         TabIndex        =   206
         Top             =   5880
         Width           =   1035
      End
      Begin VB.Label Label30 
         Caption         =   "完成时间"
         Height          =   165
         Left            =   -71850
         TabIndex        =   205
         Top             =   5880
         Width           =   1035
      End
      Begin VB.Label Label31 
         Caption         =   "旅途时间"
         Height          =   165
         Left            =   -68730
         TabIndex        =   204
         Top             =   5880
         Width           =   1035
      End
      Begin VB.Label Label32 
         Caption         =   "加班工时"
         Height          =   165
         Left            =   -66300
         TabIndex        =   203
         Top             =   5880
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
         TabIndex        =   202
         Top             =   6510
         Width           =   885
      End
      Begin VB.Shape Shape2 
         Height          =   3165
         Left            =   -74970
         Top             =   3870
         Width           =   14985
      End
      Begin VB.Label Label35 
         Caption         =   "客户签名："
         Height          =   225
         Left            =   -64080
         TabIndex        =   201
         Top             =   5850
         Width           =   945
      End
      Begin VB.Label Label36 
         Caption         =   "日期："
         Height          =   195
         Left            =   -64080
         TabIndex        =   200
         Top             =   6450
         Width           =   945
      End
      Begin VB.Label Label37 
         Caption         =   "质量控制签名："
         Height          =   195
         Left            =   -62010
         TabIndex        =   199
         Top             =   5880
         Width           =   1275
      End
      Begin VB.Line Line34 
         X1              =   -74970
         X2              =   -60030
         Y1              =   6060
         Y2              =   6060
      End
      Begin VB.Line Line35 
         X1              =   -74970
         X2              =   -60030
         Y1              =   6390
         Y2              =   6390
      End
      Begin VB.Line Line36 
         X1              =   -64170
         X2              =   -64170
         Y1              =   5820
         Y2              =   7020
      End
      Begin VB.Line Line37 
         X1              =   -62070
         X2              =   -62070
         Y1              =   5820
         Y2              =   7020
      End
      Begin VB.Label Label38 
         Caption         =   "日期："
         Height          =   195
         Left            =   -62010
         TabIndex        =   198
         Top             =   6450
         Width           =   945
      End
      Begin VB.Label Label42 
         Caption         =   $"newGZD9.frx":02D6
         Height          =   1545
         Left            =   -74820
         TabIndex        =   197
         Top             =   120
         Width           =   3705
      End
      Begin VB.Line Line23 
         X1              =   -62040
         X2              =   -60150
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Line Line24 
         X1              =   -62790
         X2              =   -60060
         Y1              =   1095
         Y2              =   1095
      End
      Begin VB.Label Label25 
         Caption         =   "原因:"
         Height          =   210
         Left            =   -63540
         TabIndex        =   196
         Top             =   915
         Width           =   585
      End
      Begin VB.Line Line38 
         Index           =   0
         X1              =   -71100
         X2              =   -59880
         Y1              =   1725
         Y2              =   1725
      End
      Begin VB.Label Label39 
         Caption         =   "漏点描述"
         Height          =   195
         Index           =   0
         Left            =   -72120
         TabIndex        =   195
         Top             =   1545
         Width           =   885
      End
   End
   Begin VB.Line Line32 
      X1              =   7680
      X2              =   11760
      Y1              =   1170
      Y2              =   1170
   End
   Begin VB.Line Line31 
      X1              =   7680
      X2              =   11760
      Y1              =   750
      Y2              =   750
   End
   Begin VB.Line Line30 
      X1              =   7680
      X2              =   11760
      Y1              =   330
      Y2              =   330
   End
   Begin VB.Line Line28 
      X1              =   1590
      X2              =   5670
      Y1              =   1125
      Y2              =   1125
   End
   Begin VB.Line Line27 
      X1              =   1590
      X2              =   5670
      Y1              =   705
      Y2              =   705
   End
   Begin VB.Line Line26 
      X1              =   1590
      X2              =   5685
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Label lblGid 
      Caption         =   "lblGid"
      Height          =   225
      Left            =   9030
      TabIndex        =   20
      Top             =   1320
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblkhdh 
      Caption         =   "lblkhdh"
      Height          =   225
      Left            =   11040
      TabIndex        =   19
      Top             =   1200
      Visible         =   0   'False
      Width           =   885
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
      TabIndex        =   18
      Top             =   120
      Width           =   1605
   End
   Begin VB.Label Label39 
      Caption         =   "NO:"
      Height          =   255
      Index           =   1
      Left            =   12450
      TabIndex        =   17
      Top             =   120
      Width           =   495
   End
   Begin VB.Label LBLKjj 
      Caption         =   $"newGZD9.frx":0338
      ForeColor       =   &H8000000D&
      Height          =   1305
      Left            =   12210
      TabIndex        =   16
      Top             =   300
      Width           =   2835
   End
End
Attribute VB_Name = "newGZD9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
