VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form NewGzd7 
   Caption         =   "施工工作报告（单）"
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
      Left            =   1710
      TabIndex        =   72
      Top             =   1350
      Width           =   4065
   End
   Begin MSDataGridLib.DataGrid dtgRen 
      Height          =   8085
      Left            =   10680
      TabIndex        =   69
      Top             =   -570
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
   Begin MSDataGridLib.DataGrid comHtbh 
      Height          =   1155
      Left            =   5820
      TabIndex        =   21
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
      Left            =   6450
      MultiLine       =   -1  'True
      TabIndex        =   68
      Text            =   "NewGzd7.frx":0000
      Top             =   120
      Width           =   1335
   End
   Begin TabDlg.SSTab tabNr 
      Height          =   8565
      Left            =   0
      TabIndex        =   22
      Top             =   1650
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   15108
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "内容1"
      TabPicture(0)   =   "NewGzd7.frx":0039
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line5(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line5(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line5(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line4(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Line4(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line4(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line4(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line3(13)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Line3(12)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line3(11)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Line3(10)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Line3(9)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Line3(7)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Line3(6)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Line3(5)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Line3(4)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Line3(3)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Line3(2)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Line3(1)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Line3(0)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Shape1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label11"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label10"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label9"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label8"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label7"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label6"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label5"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label4"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label3"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label1"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Line2"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "TA(40)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "TA(39)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "TA(38)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "TA(37)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "TA(36)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "TA(35)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "C1(2)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "C1(1)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "TA(34)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "TA(33)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "TA(32)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "TA(31)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "TA(30)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "TA(29)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "TA(28)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "TA(27)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "TA(26)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "TA(25)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "TA(24)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "TA(23)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "TA(22)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "TA(21)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "TA(20)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "TA(19)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "TA(18)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "TA(17)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "TA(16)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "TA(15)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "TA(14)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "TA(13)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "TA(12)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "TA(11)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "TA(10)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "TA(9)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "TA(8)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "TA(7)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "TA(6)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "TA(5)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "TA(4)"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "TA(3)"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "TA(2)"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "TA(1)"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).ControlCount=   76
      TabCaption(1)   =   "内容2"
      TabPicture(1)   =   "NewGzd7.frx":0055
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label17"
      Tab(1).Control(1)=   "Label16"
      Tab(1).Control(2)=   "Label15"
      Tab(1).Control(3)=   "Label14"
      Tab(1).Control(4)=   "Label13"
      Tab(1).Control(5)=   "Label26"
      Tab(1).Control(6)=   "Label27"
      Tab(1).Control(7)=   "Label28"
      Tab(1).Control(8)=   "Line24"
      Tab(1).Control(9)=   "Line25"
      Tab(1).Control(10)=   "Line33"
      Tab(1).Control(11)=   "Label29"
      Tab(1).Control(12)=   "Label30"
      Tab(1).Control(13)=   "Label31"
      Tab(1).Control(14)=   "Label32"
      Tab(1).Control(15)=   "Label34"
      Tab(1).Control(16)=   "Shape2(1)"
      Tab(1).Control(17)=   "Label35"
      Tab(1).Control(18)=   "Label36"
      Tab(1).Control(19)=   "Label37"
      Tab(1).Control(20)=   "Line34"
      Tab(1).Control(21)=   "Line35"
      Tab(1).Control(22)=   "Line36"
      Tab(1).Control(23)=   "Line37"
      Tab(1).Control(24)=   "Label38"
      Tab(1).Control(25)=   "Line10"
      Tab(1).Control(26)=   "Line9"
      Tab(1).Control(27)=   "Line8"
      Tab(1).Control(28)=   "Line7"
      Tab(1).Control(29)=   "Line3(16)"
      Tab(1).Control(30)=   "Line3(15)"
      Tab(1).Control(31)=   "Line3(14)"
      Tab(1).Control(32)=   "Shape2(0)"
      Tab(1).Control(33)=   "Label12"
      Tab(1).Control(34)=   "Line3(17)"
      Tab(1).Control(35)=   "TA(61)"
      Tab(1).Control(36)=   "TA(63)"
      Tab(1).Control(37)=   "dtpC"
      Tab(1).Control(38)=   "dtpB"
      Tab(1).Control(39)=   "BA(14)"
      Tab(1).Control(40)=   "TA(62)"
      Tab(1).Control(41)=   "C1(3)"
      Tab(1).Control(42)=   "C1(4)"
      Tab(1).Control(43)=   "Text3"
      Tab(1).Control(44)=   "TA(64)"
      Tab(1).Control(45)=   "BA(8)"
      Tab(1).Control(46)=   "BA(9)"
      Tab(1).Control(47)=   "BA(10)"
      Tab(1).Control(48)=   "BA(11)"
      Tab(1).Control(49)=   "Frame1"
      Tab(1).Control(50)=   "BA(12)"
      Tab(1).Control(51)=   "BA(13)"
      Tab(1).Control(52)=   "BA(15)"
      Tab(1).Control(53)=   "BA(16)"
      Tab(1).Control(54)=   "TA(60)"
      Tab(1).Control(55)=   "TA(59)"
      Tab(1).Control(56)=   "TA(58)"
      Tab(1).Control(57)=   "TA(57)"
      Tab(1).Control(58)=   "TA(56)"
      Tab(1).Control(59)=   "TA(55)"
      Tab(1).Control(60)=   "TA(54)"
      Tab(1).Control(61)=   "TA(53)"
      Tab(1).Control(62)=   "TA(52)"
      Tab(1).Control(63)=   "TA(51)"
      Tab(1).Control(64)=   "TA(50)"
      Tab(1).Control(65)=   "TA(49)"
      Tab(1).Control(66)=   "TA(48)"
      Tab(1).Control(67)=   "TA(47)"
      Tab(1).Control(68)=   "TA(46)"
      Tab(1).Control(69)=   "TA(45)"
      Tab(1).Control(70)=   "TA(44)"
      Tab(1).Control(71)=   "TA(43)"
      Tab(1).Control(72)=   "TA(42)"
      Tab(1).Control(73)=   "TA(41)"
      Tab(1).ControlCount=   74
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   1
         Left            =   630
         TabIndex        =   73
         Top             =   930
         Width           =   2715
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   2
         Left            =   630
         TabIndex        =   77
         Top             =   1290
         Width           =   2715
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   3
         Left            =   630
         TabIndex        =   81
         Top             =   1650
         Width           =   2715
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   4
         Left            =   630
         TabIndex        =   85
         Top             =   2010
         Width           =   2715
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   5
         Left            =   630
         TabIndex        =   89
         Top             =   2370
         Width           =   2715
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   6
         Left            =   630
         TabIndex        =   93
         Top             =   2730
         Width           =   2715
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   7
         Left            =   630
         TabIndex        =   97
         Top             =   3090
         Width           =   2715
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   8
         Left            =   630
         TabIndex        =   101
         Top             =   3450
         Width           =   2715
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   9
         Left            =   3750
         TabIndex        =   74
         Top             =   930
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   10
         Left            =   3750
         TabIndex        =   78
         Top             =   1290
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   11
         Left            =   3750
         TabIndex        =   82
         Top             =   1650
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   12
         Left            =   3750
         TabIndex        =   86
         Top             =   2010
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   13
         Left            =   3750
         TabIndex        =   90
         Top             =   2370
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   14
         Left            =   3750
         TabIndex        =   94
         Top             =   2730
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   15
         Left            =   3750
         TabIndex        =   98
         Top             =   3090
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   16
         Left            =   3750
         TabIndex        =   102
         Top             =   3450
         Width           =   795
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   17
         Left            =   5010
         TabIndex        =   75
         Top             =   930
         Width           =   3195
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   18
         Left            =   5010
         TabIndex        =   79
         Top             =   1290
         Width           =   3195
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   19
         Left            =   5010
         TabIndex        =   83
         Top             =   1650
         Width           =   3195
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   20
         Left            =   5010
         TabIndex        =   87
         Top             =   2010
         Width           =   3195
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   21
         Left            =   5010
         TabIndex        =   91
         Top             =   2370
         Width           =   3195
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   22
         Left            =   5010
         TabIndex        =   95
         Top             =   2730
         Width           =   3195
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   23
         Left            =   5010
         TabIndex        =   99
         Top             =   3090
         Width           =   3195
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   24
         Left            =   5010
         TabIndex        =   103
         Top             =   3450
         Width           =   3195
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   25
         Left            =   8460
         TabIndex        =   76
         Top             =   930
         Width           =   6285
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   26
         Left            =   8460
         TabIndex        =   80
         Top             =   1290
         Width           =   6285
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   27
         Left            =   8460
         TabIndex        =   84
         Top             =   1650
         Width           =   6285
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   28
         Left            =   8460
         TabIndex        =   88
         Top             =   2010
         Width           =   6285
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   29
         Left            =   8460
         TabIndex        =   92
         Top             =   2370
         Width           =   6285
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   30
         Left            =   8460
         TabIndex        =   96
         Top             =   2730
         Width           =   6285
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   31
         Left            =   8460
         TabIndex        =   100
         Top             =   3090
         Width           =   6285
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   32
         Left            =   8460
         TabIndex        =   104
         Top             =   3450
         Width           =   6285
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   510
         Index           =   33
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   105
         Top             =   4050
         Width           =   14745
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   450
         Index           =   34
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   106
         Top             =   4860
         Width           =   14775
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "有"
         Height          =   195
         Index           =   1
         Left            =   1890
         TabIndex        =   56
         Top             =   5400
         Width           =   555
      End
      Begin VB.CheckBox C1 
         Alignment       =   1  'Right Justify
         Caption         =   "无"
         Height          =   195
         Index           =   2
         Left            =   1890
         TabIndex        =   55
         Top             =   5670
         Width           =   555
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   35
         Left            =   3180
         TabIndex        =   107
         Top             =   5400
         Width           =   5055
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   36
         Left            =   8490
         TabIndex        =   108
         Top             =   5400
         Width           =   6255
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   37
         Left            =   3180
         TabIndex        =   109
         Top             =   5670
         Width           =   5055
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   38
         Left            =   8490
         TabIndex        =   110
         Top             =   5670
         Width           =   6255
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   450
         Index           =   39
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   111
         Top             =   6150
         Width           =   14655
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   40
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   112
         Top             =   6960
         Width           =   14655
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   41
         Left            =   -74790
         TabIndex        =   113
         Top             =   570
         Width           =   915
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   42
         Left            =   -74790
         TabIndex        =   118
         Top             =   795
         Width           =   915
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   43
         Left            =   -74790
         TabIndex        =   123
         Top             =   1035
         Width           =   915
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   44
         Left            =   -74790
         TabIndex        =   128
         Top             =   1260
         Width           =   915
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   45
         Left            =   -73590
         TabIndex        =   114
         Top             =   570
         Width           =   3045
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   46
         Left            =   -73590
         TabIndex        =   119
         Top             =   795
         Width           =   3045
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   47
         Left            =   -73590
         TabIndex        =   124
         Top             =   1035
         Width           =   3045
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   48
         Left            =   -73590
         TabIndex        =   129
         Top             =   1260
         Width           =   3045
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   49
         Left            =   -70350
         TabIndex        =   115
         Top             =   570
         Width           =   3975
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   50
         Left            =   -70350
         TabIndex        =   120
         Top             =   795
         Width           =   3975
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   51
         Left            =   -70350
         TabIndex        =   125
         Top             =   1035
         Width           =   3975
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   52
         Left            =   -70350
         TabIndex        =   130
         Top             =   1260
         Width           =   3975
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   53
         Left            =   -66270
         TabIndex        =   116
         Top             =   570
         Width           =   4245
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   54
         Left            =   -66270
         TabIndex        =   121
         Top             =   795
         Width           =   4245
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   55
         Left            =   -66270
         TabIndex        =   126
         Top             =   1035
         Width           =   4245
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   56
         Left            =   -66270
         TabIndex        =   131
         Top             =   1260
         Width           =   4245
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   57
         Left            =   -61890
         TabIndex        =   117
         Top             =   570
         Width           =   1665
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   58
         Left            =   -61890
         TabIndex        =   122
         Top             =   795
         Width           =   1665
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   59
         Left            =   -61890
         TabIndex        =   127
         Top             =   1035
         Width           =   1665
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   60
         Left            =   -61890
         TabIndex        =   132
         Top             =   1260
         Width           =   1665
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
         TabIndex        =   34
         Top             =   4320
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
         TabIndex        =   142
         Top             =   3750
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
         TabIndex        =   141
         Top             =   3750
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
         TabIndex        =   140
         Text            =   "NewGzd7.frx":0071
         Top             =   4080
         Width           =   9345
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Left            =   -74940
         TabIndex        =   28
         Top             =   3750
         Width           =   10755
         Begin VB.OptionButton FPD 
            Caption         =   "尚可"
            Height          =   195
            Left            =   6150
            TabIndex        =   32
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton FPC 
            Caption         =   "较满意"
            Height          =   195
            Left            =   4550
            TabIndex        =   31
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton FPB 
            Caption         =   "满意"
            Height          =   195
            Left            =   2950
            TabIndex        =   30
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton FPA 
            Caption         =   "优秀"
            Height          =   195
            Left            =   1350
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
         Left            =   -65220
         TabIndex        =   139
         Text            =   "的"
         Top             =   3510
         Width           =   1035
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   10
         Left            =   -67620
         TabIndex        =   138
         Text            =   "的"
         Top             =   3510
         Width           =   1035
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   9
         Left            =   -70290
         TabIndex        =   137
         Text            =   "的"
         Top             =   3510
         Width           =   1035
      End
      Begin VB.TextBox BA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   8
         Left            =   -73530
         TabIndex        =   136
         Text            =   "的"
         Top             =   3510
         Width           =   1035
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   64
         Left            =   -61230
         TabIndex        =   27
         Top             =   3180
         Width           =   1065
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   -61860
         TabIndex        =   26
         Text            =   "复核人:"
         Top             =   3270
         Width           =   735
      End
      Begin VB.CheckBox C1 
         Caption         =   "未完成"
         Height          =   180
         Index           =   4
         Left            =   -61110
         TabIndex        =   25
         Top             =   2100
         Width           =   945
      End
      Begin VB.CheckBox C1 
         Caption         =   "完成"
         Height          =   180
         Index           =   3
         Left            =   -62250
         TabIndex        =   24
         Top             =   2100
         Width           =   1005
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   570
         Index           =   62
         Left            =   -73560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   134
         Top             =   2310
         Width           =   13545
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
         TabIndex        =   23
         Top             =   4320
         Width           =   1725
      End
      Begin MSComCtl2.DTPicker dtpB 
         Height          =   225
         Left            =   -64080
         TabIndex        =   35
         Top             =   4320
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
         TabIndex        =   36
         Top             =   4320
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
         Height          =   480
         Index           =   63
         Left            =   -73560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   135
         Top             =   2970
         Width           =   13515
      End
      Begin VB.TextBox TA 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   61
         Left            =   -73560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   133
         Top             =   1560
         Width           =   13545
      End
      Begin VB.Line Line3 
         Index           =   17
         X1              =   -74940
         X2              =   -60090
         Y1              =   1230
         Y2              =   1230
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   240
         X2              =   14850
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label Label1 
         Caption         =   "所维修（或安装、调试、检测、查看等）之设备、装置的详细参数："
         Height          =   165
         Left            =   90
         TabIndex        =   67
         Top             =   240
         Width           =   6945
      End
      Begin VB.Label Label3 
         Caption         =   "设备品牌及名称"
         Height          =   195
         Left            =   840
         TabIndex        =   65
         Top             =   600
         Width           =   2145
      End
      Begin VB.Label Label4 
         Caption         =   "数量"
         Height          =   195
         Left            =   3840
         TabIndex        =   64
         Top             =   600
         Width           =   555
      End
      Begin VB.Label Label5 
         Caption         =   "设备型号或序列号"
         Height          =   195
         Left            =   5700
         TabIndex        =   63
         Top             =   600
         Width           =   2205
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "具体规格参数"
         Height          =   195
         Left            =   8550
         TabIndex        =   62
         Top             =   600
         Width           =   5925
      End
      Begin VB.Label Label7 
         Caption         =   "客户要求及现场情况："
         Height          =   225
         Left            =   180
         TabIndex        =   61
         Top             =   3780
         Width           =   1995
      End
      Begin VB.Label Label8 
         Caption         =   "故障判断或施工步骤："
         Height          =   225
         Left            =   210
         TabIndex        =   60
         Top             =   4650
         Width           =   1995
      End
      Begin VB.Label Label9 
         Caption         =   "有无故障代码"
         Height          =   255
         Left            =   210
         TabIndex        =   59
         Top             =   5400
         Width           =   1185
      End
      Begin VB.Label Label10 
         Caption         =   "处理及结果："
         Height          =   195
         Left            =   210
         TabIndex        =   58
         Top             =   5910
         Width           =   1155
      End
      Begin VB.Label Label11 
         Caption         =   "备注：　"
         Height          =   255
         Left            =   210
         TabIndex        =   57
         Top             =   6690
         Width           =   825
      End
      Begin VB.Shape Shape1 
         Height          =   6945
         Left            =   60
         Top             =   480
         Width           =   14865
      End
      Begin VB.Line Line3 
         Index           =   0
         X1              =   60
         X2              =   14910
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line3 
         Index           =   1
         X1              =   60
         X2              =   14910
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Line Line3 
         Index           =   2
         X1              =   60
         X2              =   14910
         Y1              =   1530
         Y2              =   1530
      End
      Begin VB.Line Line3 
         Index           =   3
         X1              =   60
         X2              =   14910
         Y1              =   1890
         Y2              =   1890
      End
      Begin VB.Line Line3 
         Index           =   4
         X1              =   60
         X2              =   14910
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line3 
         Index           =   5
         X1              =   60
         X2              =   14910
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line3 
         Index           =   6
         X1              =   60
         X2              =   14910
         Y1              =   2970
         Y2              =   2970
      End
      Begin VB.Line Line3 
         Index           =   7
         X1              =   60
         X2              =   14910
         Y1              =   3330
         Y2              =   3330
      End
      Begin VB.Line Line3 
         Index           =   9
         X1              =   60
         X2              =   14910
         Y1              =   3660
         Y2              =   3660
      End
      Begin VB.Line Line3 
         Index           =   10
         X1              =   60
         X2              =   14910
         Y1              =   4590
         Y2              =   4590
      End
      Begin VB.Line Line3 
         Index           =   11
         X1              =   60
         X2              =   14910
         Y1              =   5340
         Y2              =   5340
      End
      Begin VB.Line Line3 
         Index           =   12
         X1              =   60
         X2              =   14910
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Line Line3 
         Index           =   13
         X1              =   60
         X2              =   14910
         Y1              =   6630
         Y2              =   6630
      End
      Begin VB.Line Line4 
         Index           =   0
         X1              =   510
         X2              =   510
         Y1              =   480
         Y2              =   3660
      End
      Begin VB.Line Line4 
         Index           =   1
         X1              =   3570
         X2              =   3570
         Y1              =   480
         Y2              =   3660
      End
      Begin VB.Line Line4 
         Index           =   2
         X1              =   4800
         X2              =   4800
         Y1              =   480
         Y2              =   3660
      End
      Begin VB.Line Line4 
         Index           =   3
         X1              =   8340
         X2              =   8340
         Y1              =   480
         Y2              =   3660
      End
      Begin VB.Line Line5 
         Index           =   0
         X1              =   1650
         X2              =   1650
         Y1              =   5880
         Y2              =   5340
      End
      Begin VB.Line Line5 
         Index           =   1
         X1              =   2880
         X2              =   2880
         Y1              =   5880
         Y2              =   5340
      End
      Begin VB.Line Line5 
         Index           =   2
         X1              =   8370
         X2              =   8370
         Y1              =   5880
         Y2              =   5340
      End
      Begin VB.Line Line6 
         X1              =   2880
         X2              =   14910
         Y1              =   5610
         Y2              =   5610
      End
      Begin VB.Label Label12 
         Caption         =   "在施工过程中的材料或零配件情况"
         Height          =   195
         Left            =   -74790
         TabIndex        =   54
         Top             =   60
         Width           =   2985
      End
      Begin VB.Shape Shape2 
         Height          =   1215
         Index           =   0
         Left            =   -74940
         Top             =   270
         Width           =   14865
      End
      Begin VB.Line Line3 
         Index           =   14
         X1              =   -74940
         X2              =   -60090
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Line Line3 
         Index           =   15
         X1              =   -74940
         X2              =   -60090
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Line Line3 
         Index           =   16
         X1              =   -74940
         X2              =   -60090
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Line Line7 
         X1              =   -73740
         X2              =   -73740
         Y1              =   1470
         Y2              =   270
      End
      Begin VB.Line Line8 
         X1              =   -70440
         X2              =   -70440
         Y1              =   1470
         Y2              =   270
      End
      Begin VB.Line Line9 
         X1              =   -66330
         X2              =   -66330
         Y1              =   1470
         Y2              =   270
      End
      Begin VB.Line Line10 
         X1              =   -61965
         X2              =   -61965
         Y1              =   1485
         Y2              =   270
      End
      Begin VB.Label Label38 
         Caption         =   "日期："
         Height          =   195
         Left            =   -62010
         TabIndex        =   48
         Top             =   4080
         Width           =   945
      End
      Begin VB.Line Line37 
         X1              =   -62070
         X2              =   -62070
         Y1              =   3450
         Y2              =   4650
      End
      Begin VB.Line Line36 
         X1              =   -64170
         X2              =   -64170
         Y1              =   3450
         Y2              =   4650
      End
      Begin VB.Line Line35 
         X1              =   -74970
         X2              =   -60030
         Y1              =   4020
         Y2              =   4020
      End
      Begin VB.Line Line34 
         X1              =   -74970
         X2              =   -60030
         Y1              =   3690
         Y2              =   3690
      End
      Begin VB.Label Label37 
         Caption         =   "质量控制签名："
         Height          =   165
         Left            =   -62010
         TabIndex        =   47
         Top             =   3510
         Width           =   1275
      End
      Begin VB.Label Label36 
         Caption         =   "日期："
         Height          =   195
         Left            =   -64080
         TabIndex        =   46
         Top             =   4080
         Width           =   945
      End
      Begin VB.Label Label35 
         Caption         =   "客户签名："
         Height          =   195
         Left            =   -64080
         TabIndex        =   45
         Top             =   3480
         Width           =   945
      End
      Begin VB.Shape Shape2 
         Height          =   3165
         Index           =   1
         Left            =   -74970
         Top             =   1500
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
         TabIndex        =   44
         Top             =   4140
         Width           =   885
      End
      Begin VB.Label Label32 
         Caption         =   "加班工时"
         Height          =   165
         Left            =   -66300
         TabIndex        =   43
         Top             =   3510
         Width           =   1035
      End
      Begin VB.Label Label31 
         Caption         =   "旅途时间"
         Height          =   165
         Left            =   -68730
         TabIndex        =   42
         Top             =   3510
         Width           =   1035
      End
      Begin VB.Label Label30 
         Caption         =   "完成时间"
         Height          =   165
         Left            =   -71850
         TabIndex        =   41
         Top             =   3510
         Width           =   1035
      End
      Begin VB.Label Label29 
         Caption         =   "到达时间"
         Height          =   165
         Left            =   -74820
         TabIndex        =   40
         Top             =   3510
         Width           =   1035
      End
      Begin VB.Line Line33 
         X1              =   -74970
         X2              =   -60030
         Y1              =   3450
         Y2              =   3450
      End
      Begin VB.Line Line25 
         X1              =   -74970
         X2              =   -59910
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line24 
         X1              =   -74970
         X2              =   -59940
         Y1              =   2280
         Y2              =   2280
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
         TabIndex        =   39
         Top             =   2970
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
         TabIndex        =   38
         Top             =   2340
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
         TabIndex        =   37
         Top             =   1560
         Width           =   1005
      End
      Begin VB.Label Label13 
         Caption         =   "数量"
         Height          =   225
         Left            =   -74790
         TabIndex        =   53
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label14 
         Caption         =   "零配件或材料名称"
         Height          =   225
         Left            =   -73140
         TabIndex        =   52
         Top             =   300
         Width           =   2505
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "零件编号或型号规格"
         Height          =   225
         Left            =   -70200
         TabIndex        =   51
         Top             =   300
         Width           =   3675
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "使用情况"
         Height          =   225
         Left            =   -66210
         TabIndex        =   50
         Top             =   300
         Width           =   4305
      End
      Begin VB.Label Label17 
         Caption         =   "供货方"
         Height          =   225
         Left            =   -61770
         TabIndex        =   49
         Top             =   300
         Width           =   1245
      End
      Begin VB.Label Label2 
         Caption         =   $"NewGzd7.frx":0074
         Height          =   2775
         Left            =   150
         TabIndex        =   66
         Top             =   930
         Width           =   195
      End
   End
   Begin VB.ComboBox comXmmc 
      Height          =   300
      Left            =   1710
      TabIndex        =   20
      Top             =   780
      Width           =   4125
   End
   Begin VB.CommandButton cmdBack 
      Height          =   375
      Left            =   14460
      Picture         =   "NewGzd7.frx":009C
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "返回"
      Top             =   10380
      Width           =   465
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   13980
      Picture         =   "NewGzd7.frx":019E
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "保存"
      Top             =   10380
      Width           =   465
   End
   Begin VB.CommandButton cmdMod 
      Height          =   375
      Left            =   13500
      Picture         =   "NewGzd7.frx":0808
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "修改"
      Top             =   10380
      Width           =   465
   End
   Begin VB.CommandButton cmdQm 
      Caption         =   "cmdQm"
      Height          =   345
      Index           =   0
      Left            =   9840
      TabIndex        =   14
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
      Left            =   330
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "NewGzd7.frx":0B12
      Top             =   120
      Width           =   1365
   End
   Begin VB.CheckBox C1 
      Alignment       =   1  'Right Justify
      Caption         =   "1号"
      Height          =   285
      Index           =   0
      Left            =   12450
      TabIndex        =   9
      Top             =   750
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox TA 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   13380
      TabIndex        =   8
      Top             =   450
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
      Left            =   1740
      TabIndex        =   7
      Top             =   120
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
      TabIndex        =   6
      Top             =   1080
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
      Left            =   1740
      TabIndex        =   5
      Top             =   540
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
      Left            =   1740
      TabIndex        =   4
      Top             =   960
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
      Left            =   7830
      TabIndex        =   3
      Text            =   "的"
      Top             =   180
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
      Left            =   7830
      TabIndex        =   2
      Top             =   990
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
      Index           =   6
      Left            =   7830
      Locked          =   -1  'True
      TabIndex        =   1
      Tag             =   "20"
      Top             =   1410
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
      Left            =   7830
      TabIndex        =   0
      Text            =   "的"
      Top             =   600
      Width           =   4245
   End
   Begin MSComCtl2.DTPicker dtpA 
      Height          =   195
      Left            =   7830
      TabIndex        =   11
      Top             =   1410
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   344
      _Version        =   393216
      Format          =   149094401
      CurrentDate     =   38897
   End
   Begin VB.Label LBLKjj 
      Caption         =   $"NewGzd7.frx":0B49
      ForeColor       =   &H8000000D&
      Height          =   1305
      Left            =   12150
      TabIndex        =   143
      Top             =   330
      Width           =   2835
   End
   Begin VB.Line Line38 
      X1              =   1680
      X2              =   5835
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Label Label39 
      Caption         =   "NO:"
      Height          =   255
      Left            =   12540
      TabIndex        =   71
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
      Left            =   13110
      TabIndex        =   70
      Top             =   180
      Width           =   1605
   End
   Begin VB.Label lblTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   10860
      TabIndex        =   19
      Top             =   10440
      Width           =   1905
   End
   Begin VB.Label lblQM 
      Caption         =   "签字提交"
      Height          =   225
      Index           =   0
      Left            =   8910
      TabIndex        =   18
      Top             =   10440
      Width           =   795
   End
   Begin VB.Line Line3 
      Index           =   8
      X1              =   0
      X2              =   14850
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line26 
      X1              =   1740
      X2              =   5835
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line27 
      X1              =   1740
      X2              =   5820
      Y1              =   765
      Y2              =   765
   End
   Begin VB.Line Line28 
      X1              =   1740
      X2              =   5820
      Y1              =   1185
      Y2              =   1185
   End
   Begin VB.Line Line30 
      X1              =   7830
      X2              =   11910
      Y1              =   390
      Y2              =   390
   End
   Begin VB.Line Line31 
      X1              =   7830
      X2              =   11910
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line32 
      X1              =   7830
      X2              =   11910
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Line Line1 
      X1              =   7830
      X2              =   11925
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblGid 
      Caption         =   "lblGid"
      Height          =   225
      Left            =   13500
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblkhdh 
      Caption         =   "lblkhdh"
      Height          =   225
      Left            =   12390
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   885
   End
End
Attribute VB_Name = "NewGzd7"
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
For oo = 1 To 64
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
    BA(7).Text = adoRen.Fields("username").Value
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
BA(15).Tag = 50
BA(16).Tag = 50
BA(17).Tag = 50
BA(8).Tag = 10
BA(9).Tag = 10
BA(10).Tag = 10
BA(11).Tag = 10
For oo = 1 To 64
    TA(oo).Tag = 50
Next
TA(33).Tag = 200
TA(34).Tag = 200
TA(39).Tag = 200
TA(40).Tag = 200
TA(61).Tag = 200
TA(62).Tag = 200
TA(63).Tag = 200
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
