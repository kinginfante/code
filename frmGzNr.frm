VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGzNr 
   Caption         =   "销售日记"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdQm 
      Caption         =   "cmdQm"
      Height          =   345
      Index           =   0
      Left            =   7680
      TabIndex        =   74
      Top             =   8190
      Width           =   945
   End
   Begin VB.Frame frmHide 
      Caption         =   "frmHid"
      Height          =   1455
      Left            =   7080
      TabIndex        =   68
      Top             =   6060
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Label lblLcou 
         Caption         =   "lblLcou"
         Height          =   255
         Left            =   3180
         TabIndex        =   78
         Top             =   690
         Width           =   1185
      End
      Begin VB.Label lblLc 
         Caption         =   "lblLc"
         Height          =   315
         Left            =   1050
         TabIndex        =   73
         Top             =   630
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label lblNlb 
         Caption         =   "lblNlb"
         Height          =   225
         Left            =   1920
         TabIndex        =   72
         Top             =   810
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label lblLcRen 
         Caption         =   "lblLcRen"
         Height          =   285
         Left            =   150
         TabIndex        =   71
         Top             =   420
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblLcUid 
         Caption         =   "lblLcUid"
         Height          =   285
         Left            =   240
         TabIndex        =   70
         Top             =   930
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblFwid 
         Caption         =   "lblFwid"
         Height          =   255
         Left            =   1860
         TabIndex        =   69
         Top             =   450
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin MSDataListLib.DataCombo comRen 
      Height          =   330
      Left            =   13230
      TabIndex        =   66
      Top             =   3840
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   582
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.ComboBox comKhmc 
      Height          =   300
      Left            =   1380
      TabIndex        =   64
      Text            =   "Combo1"
      Top             =   540
      Width           =   3495
   End
   Begin VB.CommandButton cmdRenDel 
      Caption         =   "删除"
      Height          =   285
      Left            =   14670
      TabIndex        =   62
      Top             =   3300
      Width           =   555
   End
   Begin VB.CommandButton cmdRenAdd 
      Caption         =   "确认"
      Height          =   345
      Left            =   14670
      TabIndex        =   61
      Top             =   2940
      Width           =   585
   End
   Begin VB.TextBox txtJw 
      Height          =   1065
      Left            =   1380
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   59
      Top             =   2970
      Width           =   11565
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgRen 
      Height          =   1005
      Left            =   13230
      TabIndex        =   57
      Top             =   2790
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1773
      _Version        =   393216
      BackColor       =   -2147483634
      Rows            =   3
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483634
      FocusRect       =   2
      HighLight       =   2
      FillStyle       =   1
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
      _Band(0).GridLinesBand=   3
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.TextBox txtXmFy 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   13230
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   5640
      Width           =   1635
   End
   Begin VB.TextBox txtjzDC 
      Height          =   405
      Left            =   1380
      TabIndex        =   34
      Top             =   1920
      Width           =   11535
   End
   Begin VB.TextBox txtXm 
      Height          =   735
      Left            =   1380
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   33
      Top             =   1050
      Width           =   11535
   End
   Begin VB.TextBox txtBfMd 
      Height          =   495
      Left            =   1380
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   32
      Top             =   2400
      Width           =   11535
   End
   Begin VB.CommandButton cmdFadd 
      Caption         =   "添加"
      Height          =   345
      Left            =   12150
      TabIndex        =   31
      Top             =   5880
      Width           =   585
   End
   Begin VB.CommandButton cmdFdel 
      Caption         =   "删除"
      Height          =   345
      Left            =   12150
      TabIndex        =   30
      Top             =   6210
      Width           =   585
   End
   Begin VB.CommandButton cmdTg 
      Caption         =   "费用统计"
      Height          =   345
      Left            =   11790
      TabIndex        =   28
      Top             =   6600
      Width           =   975
   End
   Begin VB.Frame frmFy 
      Caption         =   "                                  费用类别选择："
      Height          =   1485
      Left            =   1350
      TabIndex        =   27
      Top             =   3690
      Visible         =   0   'False
      Width           =   8745
      Begin VB.OptionButton optFh 
         Caption         =   "通信费"
         Height          =   195
         Left            =   6480
         TabIndex        =   51
         Top             =   750
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton optFf 
         Caption         =   "礼品费"
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   4290
         TabIndex        =   50
         Top             =   990
         Width           =   915
      End
      Begin VB.OptionButton optFe 
         Caption         =   "招待费"
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   4290
         TabIndex        =   49
         Top             =   720
         Width           =   915
      End
      Begin VB.OptionButton optFd 
         Caption         =   "餐费"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   4290
         TabIndex        =   48
         Top             =   390
         Width           =   1005
      End
      Begin VB.OptionButton optFc 
         Caption         =   "住宿费"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   1260
         TabIndex        =   47
         Top             =   990
         Width           =   885
      End
      Begin VB.OptionButton optFb 
         Caption         =   "市外交通费"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   1260
         TabIndex        =   46
         Top             =   720
         Width           =   1305
      End
      Begin VB.OptionButton optFa 
         Caption         =   "市内交通费"
         Height          =   285
         Left            =   1260
         TabIndex        =   45
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optFg 
         Caption         =   "快递费"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   6480
         TabIndex        =   44
         Top             =   450
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.OptionButton optCL 
         Caption         =   "车辆费"
         Height          =   195
         Left            =   6480
         TabIndex        =   43
         Top             =   1020
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label15 
         Caption         =   "招待费"
         ForeColor       =   &H00004000&
         Height          =   225
         Left            =   3540
         TabIndex        =   53
         Top             =   810
         Width           =   645
      End
      Begin VB.Label Label14 
         Caption         =   "差旅费"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   480
         TabIndex        =   52
         Top             =   780
         Width           =   705
      End
   End
   Begin VB.TextBox txtXBCC 
      Height          =   765
      Left            =   1380
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Top             =   7080
      Width           =   11535
   End
   Begin VB.Frame Frame1 
      Caption         =   "项目平台"
      Height          =   2505
      Left            =   13200
      TabIndex        =   21
      Top             =   30
      Width           =   2055
      Begin VB.CommandButton cmdBj 
         BackColor       =   &H008080FF&
         Caption         =   "生成报价单"
         Height          =   285
         Left            =   390
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   2160
         Width           =   1155
      End
      Begin VB.OptionButton optA 
         Caption         =   "很困难(0)"
         ForeColor       =   &H002933EF&
         Height          =   255
         Left            =   390
         TabIndex        =   25
         Top             =   420
         Width           =   1185
      End
      Begin VB.OptionButton optB 
         Caption         =   "有难度(30)"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   390
         TabIndex        =   24
         Top             =   825
         Width           =   1245
      End
      Begin VB.OptionButton optC 
         Caption         =   "有可能(60)"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   390
         TabIndex        =   23
         Top             =   1305
         Width           =   1245
      End
      Begin VB.OptionButton optD 
         Caption         =   "有把握(90)"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   390
         TabIndex        =   22
         Top             =   1770
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdKhzl 
      Caption         =   "项目资料"
      Height          =   285
      Left            =   13230
      TabIndex        =   17
      Top             =   4530
      Width           =   1665
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "修改"
      Height          =   585
      Left            =   12570
      Picture         =   "frmGzNr.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8430
      Width           =   675
   End
   Begin VB.TextBox txtzgPd 
      ForeColor       =   &H00FF0000&
      Height          =   1065
      Left            =   1380
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   7890
      Width           =   6195
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "提交"
      Height          =   585
      Left            =   13260
      Picture         =   "frmGzNr.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8430
      Width           =   675
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   585
      Left            =   14610
      Picture         =   "frmGzNr.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8430
      Width           =   585
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除"
      Enabled         =   0   'False
      Height          =   585
      Left            =   13950
      Picture         =   "frmGzNr.frx":0A76
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8430
      Width           =   645
   End
   Begin MSAdodcLib.Adodc adoFy 
      Height          =   405
      Left            =   150
      Top             =   4470
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\work\demo\work.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\work\demo\work.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "worker"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid dtgFy 
      Bindings        =   "frmGzNr.frx":0C00
      Height          =   1905
      Left            =   1380
      TabIndex        =   29
      Top             =   5160
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   3360
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "fyLB"
         Caption         =   "费用类别"
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
         DataField       =   "fyNR"
         Caption         =   "费用内容"
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
         DataField       =   "fY"
         Caption         =   "金额"
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
      BeginProperty Column03 
         DataField       =   "htbh"
         Caption         =   "归属合同编号"
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
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4004.788
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   -1  'True
            ColumnWidth     =   1500.095
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoXmgz 
      Height          =   405
      Left            =   12660
      Top             =   7380
      Visible         =   0   'False
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   714
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\work\demo\work.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\work\demo\work.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "worker"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtBfJg 
      Height          =   1005
      Left            =   1350
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   36
      Top             =   4080
      Width           =   11535
   End
   Begin VB.Label lblRen 
      Caption         =   "lblRen"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   180
      TabIndex        =   80
      Top             =   3420
      Width           =   1065
   End
   Begin VB.Label lblHtbh 
      Caption         =   "lblHtbh"
      Height          =   195
      Left            =   90
      TabIndex        =   79
      Top             =   5490
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label lblTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   0
      Left            =   7680
      TabIndex        =   76
      Top             =   8610
      Width           =   945
   End
   Begin VB.Label lblQM 
      Caption         =   "lblQm"
      Height          =   225
      Index           =   0
      Left            =   7770
      TabIndex        =   75
      Top             =   7920
      Width           =   915
   End
   Begin VB.Label lblUid 
      Caption         =   "lblUid"
      Height          =   255
      Left            =   10410
      TabIndex        =   67
      Top             =   540
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblKid 
      Caption         =   "lblKid"
      Height          =   285
      Left            =   11640
      TabIndex        =   65
      Top             =   750
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblXid 
      Caption         =   "lblxid"
      Height          =   255
      Left            =   5940
      TabIndex        =   63
      Top             =   120
      Width           =   945
   End
   Begin VB.Label Label20 
      Caption         =   "拜访客户"
      Height          =   285
      Left            =   120
      TabIndex        =   60
      Top             =   600
      Width           =   1005
   End
   Begin VB.Label lbljW 
      Caption         =   "客户交往情况:"
      Height          =   345
      Left            =   90
      TabIndex        =   58
      Top             =   3060
      Width           =   1245
   End
   Begin VB.Label Label17 
      Caption         =   "拜访客户名单"
      Height          =   1425
      Left            =   12990
      TabIndex        =   56
      Top             =   2760
      Width           =   195
   End
   Begin VB.Label lblCxmFy 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   13200
      TabIndex        =   55
      Top             =   6540
      Width           =   1665
   End
   Begin VB.Label Label2 
      Caption         =   "该项目总计费用:"
      Height          =   285
      Left            =   13200
      TabIndex        =   54
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "竞争对手："
      Height          =   285
      Left            =   90
      TabIndex        =   42
      Top             =   1980
      Width           =   1845
   End
   Begin VB.Label Label10 
      Caption         =   "项目描述："
      Height          =   285
      Left            =   90
      TabIndex        =   41
      Top             =   1110
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "项目费用："
      Height          =   195
      Left            =   13230
      TabIndex        =   40
      Top             =   5340
      Width           =   1035
   End
   Begin VB.Label Label4 
      Caption         =   "拜访目的："
      Height          =   315
      Left            =   90
      TabIndex        =   39
      Top             =   2430
      Width           =   1125
   End
   Begin VB.Label Label5 
      Caption         =   "拜访结果："
      Height          =   225
      Left            =   90
      TabIndex        =   38
      Top             =   4140
      Width           =   945
   End
   Begin VB.Label Label6 
      Caption         =   "下步措施："
      Height          =   195
      Left            =   90
      TabIndex        =   37
      Top             =   7110
      Width           =   1005
   End
   Begin VB.Label lblGid 
      Caption         =   "lblGid"
      Height          =   285
      Left            =   11850
      TabIndex        =   20
      Top             =   480
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblYwy 
      Caption         =   "lblYwy"
      Height          =   285
      Left            =   11490
      TabIndex        =   19
      Top             =   90
      Width           =   1245
   End
   Begin VB.Label Label19 
      Caption         =   "业 务 员："
      Height          =   255
      Left            =   10500
      TabIndex        =   18
      Top             =   120
      Width           =   915
   End
   Begin VB.Label Label16 
      Caption         =   "Label16"
      DataField       =   "UserId"
      DataSource      =   "adoXmgz"
      Height          =   195
      Left            =   13830
      TabIndex        =   16
      Top             =   7950
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label12 
      Caption         =   "主管评定："
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   8010
      Width           =   1125
   End
   Begin VB.Label lblZGQZ 
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   150
      TabIndex        =   13
      Top             =   8340
      Width           =   975
   End
   Begin VB.Label lblAdr 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5880
      TabIndex        =   11
      Top             =   570
      Width           =   4335
   End
   Begin VB.Label Label13 
      Caption         =   "地    址："
      Height          =   255
      Left            =   4950
      TabIndex        =   10
      Top             =   570
      Width           =   915
   End
   Begin VB.Label Label9 
      Caption         =   "星期"
      Height          =   225
      Left            =   9540
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblWeek 
      Caption         =   "五"
      Height          =   225
      Left            =   9930
      TabIndex        =   5
      Top             =   120
      Width           =   225
   End
   Begin VB.Label lblRq 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dddddd aaaa"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   195
      Left            =   8370
      TabIndex        =   4
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label Label7 
      Caption         =   "日    期："
      Height          =   285
      Left            =   7410
      TabIndex        =   3
      Top             =   120
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "项目名称："
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   90
      Width           =   945
   End
   Begin VB.Label Label3 
      Caption         =   "项目代码："
      Height          =   225
      Left            =   4950
      TabIndex        =   1
      Top             =   120
      Width           =   915
   End
   Begin VB.Label lblXmmc 
      DataField       =   "khQc"
      Height          =   285
      Left            =   1380
      TabIndex        =   0
      Top             =   90
      Width           =   3465
   End
End
Attribute VB_Name = "frmGzNr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public AdoKh As Object
Public adoLxr As Object
Public adoBlx As Object

'Public Akid(1 To 5) As Double
Private Sub cmdBack_Click()
Dim ii As Integer
Dim tt As String
On Error Resume Next

'If comKhmc.Text <> "" And cmdSave.Enabled = True Then
'    ii = MsgBox("退出是否保存?", vbYesNoCancel + vbInformation, "当心！")
'        If ii = vbYes Then
'            Call cmdSave_Click
'        ElseIf ii = vbNo Then
'
'        ElseIf ii = vbCancel Then
'
'            Exit Sub
'        End If
'
'
'
'ElseIf comKhmc.Text = "" Then
'    tt = "delete from xmgz where gid=" & Val(lblGid.Caption)
'    adoXmgz.Recordset.Close
'    adoXmgz.Recordset.Open tt, mod1.workKK, adOpenKeyset, adCmdText
'
'End If
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
ElseIf frmGzBG.Visible = True Then
    frmGzBG.Enabled = True
    frmGzBG.ZOrder 0
    
End If
frmGzNr.Visible = False

End Sub

Private Sub cmdDel_Click()
Dim tt As String
Dim ii As String
On Error Resume Next
'If (lblLc.Caption = 0 Or lblLc.Caption = 1) And lblYwy.Caption = mod1.DName Then
    ii = MsgBox("是否要删除这篇销售日记?", vbInformation + vbYesNo, "询问")
    If ii = vbYes Then
        tt = "delete from xmgz where gid=" & Val(lblGid.Caption)
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText

        tt = "delete from fytg where gid=" & Val(lblGid.Caption)
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
        frmGzNr.Visible = False
        tt = "Select atime,khqc,xmfy,NewF,gid from xmgz where ywy='" & lblYwy.Caption & "' and aTime>='" & modXmGz.FR & _
        "' and aTime <='" & modXmGz.LR & "' and lb=1 order by aTime"
        frmGzBG.adoXm.Close
        frmGzBG.adoXm.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Set frmGzBG.dtgXmgz.DataSource = frmGzBG.adoXm
    End If
'End If
End Sub

Private Sub cmdFADD_Click()
On Error Resume Next
'如果空记录,则不能添加
adoFy.Recordset.MoveLast

If adoFy.Recordset.RecordCount > 0 Then
    If IsNull(adoFy.Recordset.Fields("fy").Value) = True Then
        Exit Sub
    ElseIf adoFy.Recordset.Fields("fy").Value = 0 Then
        Exit Sub
    End If
End If
adoFy.Recordset.AddNew "gid", Val(lblGid.Caption)
adoFy.Recordset.Fields("khmc").Value = lblXmmc.Caption
adoFy.Recordset.Fields("aTime").Value = lblRq.Caption
adoFy.Recordset.Fields("yWy").Value = lblYwy.Caption
adoFy.Recordset.Fields("uid").Value = lblUid.Caption
adoFy.Recordset.Fields("htbh").Value = lblHtbh.Caption
'adoFy.Recordset.MoveLast
Set dtgFy.DataSource = adoFy
dtgFy.Columns(0).Button = True
End Sub

Private Sub cmdFDEL_Click()
Dim ii As Integer
On Error Resume Next
ii = MsgBox("是否真的要删除此费用记录？", vbYesNo + vbInformation, "???")
If ii = vbYes Then
    adoFy.Recordset.Delete adAffectCurrent
   ' Set dtgFy.DataSource = adoFy
End If
End Sub

Private Sub cmdKhzl_Click()
Dim tt As String
On Error Resume Next
Dim Kid As Long
Dim xid As Long

    'dtgKH.Col = 2
    xid = Val(lblXid.Caption)
    'dtgKH.Col = 5
'    kid = Val(dtgKH.Text)
'    dtgKH.Col = 2

    frmWait.Show
    frmWait.ZOrder 0
    
    frmWait.Refresh
    frmWait.faWait.Play
    


    
    frmGzNr.Enabled = False
    wbDN.Visible = False
    Me.MousePointer = 11
    mod1.BTZ = 1
    Call mod1.xmQing
    Call mod1.khQing
    Call mod1.xmBound(xid)
    wbDN.lblKid.Caption = wbDN.lblYZ.Tag
    Call mod1.khBound(wbDN.lblYZ.Tag, "yz")

    wbDN.frmJE.Visible = False

    wbDN.Left = 0
    wbDN.Top = 0
    wbDN.cmdMod.Enabled = False
    wbDN.cmdSave.Enabled = False
    Me.MousePointer = 0
    wbDN.tabKh.Tab = 0

    wbDN.tabKh.TabEnabled(2) = True
    wbDN.tabKh.TabEnabled(0) = True
    
    
    

    wbDN.modFi = False

    Me.MousePointer = 0
    wbDN.cmdSave.Enabled = False
    wbDN.tabKh.Enabled = True

    wbDN.khAdd = False
    '打开项目后,默认的打开客户为项目资料
    wbDN.optYz.Value = True
    wbDN.frmGL.Visible = False
    frmWait.Visible = False
    wbDN.Visible = True
    wbDN.cmdQing.Enabled = False
    wbDN.cmdNew.Enabled = False
    wbDN.cmdRadd.Enabled = False
    wbDN.cmdRdel.Enabled = False
    If wbDN.comXyxz.Text = "物业公司" Then
        wbDN.frmGL.Visible = True
    End If
    
    '更新动态签字按钮的初始设置
        For oo = 1 To 10
           wbDN.lblQM(oo).Left = wbDN.lblQM(oo - 1).Left + 1100
           wbDN.cmdQm(oo).Left = wbDN.cmdQm(oo - 1).Left + 1100
           wbDN.lblTm(oo).Left = wbDN.lblTm(oo - 1).Left + 1100
           mod1.HTP.MoveNext
        Next
End Sub

Private Sub cmdMod_Click()
On Error Resume Next
If frmGzNr.Visible = False Then Exit Sub
If lblLcRen.Caption <> mod1.DName Or Val(lblLcRen.Caption) > Val(lblLcou.Caption) Then Exit Sub
cmdMod.Enabled = False
frmGzNr.cmdSave.Enabled = True
'If adoFy.Recordset.Fields("cwQz").Value = "" Or IsNull(adoFy.Recordset.Fields("cwQz").Value) = True Then
'    cmdFadd.Enabled = True: cmdFdel.Enabled = True
'
'Else
'        cmdFadd.Enabled = False: cmdFdel.Enabled = False
'End If
If mod1.KhK = 1 And lblLc.Caption = 2 Then
    txtzgPd.Locked = False
End If
If ((lblLc.Caption = 0 Or lblLc.Caption = 1) And mod1.DName = lblYwy.Caption) Or (mod1.BmJl = True And lblLc.Caption = 2) Then
    frmGzNr.cmdRenAdd.Visible = True
    frmGzNr.cmdRenDel.Visible = True
    frmGzNr.cmdFadd.Visible = True
    frmGzNr.cmdFdel.Visible = True
    frmGzNr.cmdTg.Visible = True
    frmGzNr.cmdDel.Enabled = True
    frmGzNr.cmdFadd.Visible = True
    frmGzNr.cmdFdel.Visible = True
    frmGzNr.cmdTg.Visible = True
End If
End Sub

Private Sub cmdQm_Click(Index As Integer)
Dim tt As String
Dim Tywy As String '单子流转到下一人的姓名
Dim Tuid As String
Dim Oywy As String '原来流转人的名字
Dim Ouid As String '原来流转人的工号

On Error Resume Next

Oywy = lblLcRen.Caption
Ouid = lblLcUid.Caption
'If cmdQm(Index).Caption <> "" Then Exit Sub
If frmGzNr.Visible = False Then Exit Sub

Call cmdTg_Click
If Index = 0 And cmdSave.Enabled = True And lblLc.Caption = 0 Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If

If Val(lblLc.Caption) = 2 And cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If

If Val(lblLc.Caption) = 2 And txtzgPd.Text = "" Then
    MsgBox "请写上您的评语!"
    Exit Sub
End If

If Index + 1 <> lblLc.Caption Then '不能在不相干的位置上乱点

    Exit Sub
End If

If Trim(lblLcUid.Caption) <> mod1.DHid Then
    MsgBox "此处应由" & lblLcRen.Caption & "签字! 请您不要再点"
    Exit Sub
End If
Dim Zi As Integer
Zi = MsgBox("是否确认签字?", vbYesNo)
If Zi = vbNo Then Exit Sub

    lblLc.Caption = lblLc.Caption + 1

    
    '更新表xunjiaD中的lcRen,lcUid 字段,以及QMRZ表中的相应字段.
                Set mod1.cmd = CreateObject("adodb.command")
                mod1.cmd.ActiveConnection = mod1.cc
                mod1.cmd.CommandText = "QMRZQM"
                mod1.cmd.CommandType = adCmdStoredProc
                mod1.cmd.Parameters("@nlb") = 40 '单子(报销单)种类
                mod1.cmd.Parameters("@lc") = lblLc.Caption  '当前流程
                mod1.cmd.Parameters("@Dname") = mod1.DName
                mod1.cmd.Parameters("@uid") = mod1.DHid
                mod1.cmd.Parameters("@btz") = mod1.BTZ '业务属性
                mod1.cmd.Parameters("@zid") = cmdQm(Index).Tag '流程顺序
                mod1.cmd.Parameters("@Qdbh") = lblGid.Caption    '单子编号
                mod1.cmd.Parameters("@pje") = ""   '评审建议
                mod1.cmd.Parameters("@bm") = mod1.Bm
                mod1.cmd.Parameters("@qy") = mod1.Qy
                mod1.cmd.Parameters("@Gren") = mod1.GJR
                mod1.cmd.Parameters("@Guid") = mod1.GJId
                mod1.cmd.Parameters("@ywy") = lblYwy.Caption
                mod1.cmd.Parameters("@yid") = lblUid.Caption
                mod1.cmd.Parameters("@comid") = mod1.comId
                mod1.cmd.Execute
                Tywy = mod1.cmd.Parameters("@Tywy").Value
                Tuid = mod1.cmd.Parameters("@Tuid").Value
                Set mod1.cmd = Nothing
                cmdQm(Index).Caption = mod1.DName
                lblTm(Index).Caption = mod1.DQda
                If lblYwy.Caption = "陈思汗" Or lblYwy.Caption = "郑泉勇" Then
                    Tywy = "潘明峰"
                    Tuid = "HM087"
                End If
                If Val(lblLc.Caption) = 2 Then
                    tt = "select username,userid from worker where zzf=1 and bmjl=1 and bm='" & mod1.Bm & "'"
                    Set mod1.HTP = CreateObject("adodb.recordset")
                    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
                    Tywy = mod1.HTP.Fields("username").Value: Tuid = mod1.HTP.Fields("userid").Value
                    mod1.HTP.Close
                    Set mod1.HTP = Nothing
                End If
                lblLcRen.Caption = Tywy
                lblLcUid.Caption = Tuid
                

If Val(lblLc.Caption) > Val(lblLcou.Caption) Then
    Call mod1.EnventFinish(frmGzNr.lblFwid.Caption)

    MsgBox "请您重视领导的建议!"
Else
    '添加事务
    Call mod1.EnventAdd("销售日记", lblXmmc.Caption, lblLcRen.Caption, lblLcUid.Caption, lblGid.Caption, lblQM(Index + 1).Caption, Oywy, Ouid, lblYwy.Caption, lblUid.Caption, lblFwid.Caption, lblGid.Caption)
    If lblLc.Caption = 2 Then
        MsgBox "现在,这篇日记将交由 " & Tywy & " 来审阅!"
    ElseIf lblLc.Caption = 3 Then
        MsgBox "我敢保证," & lblYwy.Caption & "一定会看到您的批示的!"
    End If
End If

If Dialog.Visible = True Then '更新事务列表
    Call mod1.refEnvent(1)
End If
End Sub

Private Sub cmdRenAdd_Click()
Dim tt As String
On Error Resume Next
    If txtJw.Text = "" Or lblRen.Caption = "" Then
        Exit Sub
    End If
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "xmRenAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@ren") = Trim(comRen.Text)
    mod1.cmd.Parameters("@rid") = Val(comRen.BoundText)
    mod1.cmd.Parameters("@tnr") = Trim(txtJw.Text)
    mod1.cmd.Parameters("@gid") = Val(frmGzNr.lblGid.Caption)
    
    mod1.cmd.Execute

  
    Set cmd = Nothing
    
'    tt = "select ren,llid from xmren where gid=" & Val(lblGid.Caption) & " order by llid desc"
'    adoBlx.Close
'    adoBlx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    adoBlx.Requery
    Set dtgRen.DataSource = adoBlx
    dtgRen.ColWidth(2) = 0
    dtgRen.ColWidth(3) = 0
    dtgRen.ColWidth(4) = 0
    comRen.Text = ""
End Sub

Private Sub cmdRenDel_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
ii = MsgBox("确认删除此联系人的交往信息吗?", vbInformation + vbYesNo, "询问")
If ii = vbYes Then
    dtgRen.Col = 1
    tt = "delete from xmRen where llid=" & Val(dtgRen.Text)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    
'    tt = "select * from xmren where gid=" & Val(lblGid.Caption) & " order by llid desc"
'    adoBlx.Close
'    adoBlx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    adoBlx.Requery
    Set dtgRen.DataSource = adoBlx
    txtJw.Text = adoBlx.Fields("tnr").Value
End If
End Sub


Private Sub cmdSave_Click()
On Error Resume Next
If txtBfMd.Text = "" Then
   MsgBox "请输入拜访目的！"
    Exit Sub
End If

If txtBfJg.Text = "" Then
   MsgBox "请输入拜访结果！"
    Exit Sub
End If

If txtXBCC.Text = "" Then
   MsgBox "请输入下步措施！"
    Exit Sub
End If

If txtXmFy.Text = "" Then
       MsgBox "请计算项目费用！"
    Exit Sub
End If

'If txtJw.Text <> "" And adoBlx.RecordCount = 0 Then
'    MsgBox "客户联系人!"
'    Exit Sub
'End If

'如果有一条记录为空,则删除它
adoFy.Recordset.MoveFirst
Do While Not adoFy.Recordset.EOF
    If adoFy.Recordset.Fields("fy").Value = 0 Or IsNull(adoFy.Recordset.Fields("fy").Value) = True Then
        adoFy.Recordset.Delete adAffectCurrent
    End If
    adoFy.Recordset.MoveNext
Loop

Call cmdTg_Click
Call modXmGz.xmAdd
cmdSave.Enabled = False


If lblFwid.Caption = "" Then
    lblLc.Caption = 1
    tt = "update xmgz set lc=1 where gid=" & Val(lblGid.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    '添加事务
    Call mod1.EnventAdd("销售日记", lblXmmc.Caption, lblLcRen.Caption, lblLcUid.Caption, lblGid.Caption, lblQM(0).Caption, "", "", mod1.DName, mod1.DHid, 0, lblGid.Caption)
mod1.BTZ = 4
    '更新按钮
    Call modXmGz.OpenXMGZAN(True)
End If

'更新工作报告表
'tt = "Select * from xmgz where ywy like '%" & frmZu.comYwy.Text & "%' and aTime>='" & modXmGz.Fr & _
'"' and aTime <='" & modXmGz.Lr & "' and lb=1 order by aTime"
'frmGzBG.adoXmrq.Recordset.Close
'frmGzBG.adoXmrq.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText


'    tt = "Select atime,khqc,xmfy,NewF,gid from xmgz where ywy='" & lblYwy.Caption & "' and aTime>='" & modXmGz.Fr & _
'    "' and aTime <='" & modXmGz.Lr & "' and lb=1 order by aTime"
'    frmGzBG.adoXm.Close
'    frmGzBG.adoXm.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    frmGzBG.adoXm.Requery
    Set frmGzBG.dtgXmgz.DataSource = frmGzBG.adoXm
End Sub


Private Sub cmdTg_Click()
Dim tM As Double
On Error Resume Next
tM = 0
adoFy.Recordset.MoveFirst
Do While Not adoFy.Recordset.EOF
    tM = tM + adoFy.Recordset.Fields("FY").Value
    adoFy.Recordset.MoveNext
Loop
txtXmFy.Text = tM
End Sub


Private Sub comKhmc_Click()
Dim tt As String
On Error Resume Next
tt = "Select xmAdr,kid from khzl where khqc = '" & comKhmc.Text & "' order by kid desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
lblAdr.Caption = mod1.HTP.Fields("xmAdr").Value
lblKid.Caption = mod1.HTP.Fields("kid").Value
tt = "Select khman,rid from khren where kid=" & Val(lblKid.Caption)
adoLxr.Close
adoLxr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'Set dtgRen.DataSource = adoLxr
Set comRen.RowSource = adoLxr
comRen.ListField = "khman"
comRen.BoundColumn = "rid"
End Sub

Private Sub comKhmc_DropDown()
Dim oo As Integer
Dim jj As Integer
Dim tt As String
On Error Resume Next


    '设置客户名称下拉框
    jj = comKhmc.ListCount
    If jj > 0 Then
        For oo = jj - 1 To 0 Step -1
            comKhmc.RemoveItem (oo)
        Next
    End If

        tt = "Select yzmc,wymc,qt1mc,qt2mc,qt3mc,qt4mc,qt5mc from xmKhmc where xid=" & Val(lblXid.Caption)

    AdoKh.Close
    AdoKh.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    If AdoKh.RecordCount = 1 Then
        If IsNull(AdoKh.Fields("yzmc").Value) = False Then
            comKhmc.AddItem AdoKh.Fields("yzmc").Value
        End If
        If IsNull(AdoKh.Fields("wymc")) = False Then
            comKhmc.AddItem AdoKh.Fields("wymc").Value
        End If
        For oo = 1 To 5
            If IsNull(AdoKh.Fields("qt" & oo & "mc")) = False And AdoKh.Fields("qt" & oo & "mc") <> "" Then
                comKhmc.AddItem AdoKh.Fields("qt" & oo & "mc").Value
            End If
        Next
    Else  '在豪曼办公
        frmGzNr.lblXmmc.Caption = "上海豪曼制冷空调服务有限公司"
        frmGzNr.lblAdr.Caption = "办公室"
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub comRen_Click(Area As Integer)
lblRen.Caption = comRen.Text
txtJw.Text = ""
End Sub

Private Sub dtgFy_AfterColUpdate(ByVal ColIndex As Integer)
If adoFy.Recordset.Fields("fyLB").Value = "通信费" Or adoFy.Recordset.Fields("fyLB").Value = "车辆费" Then
    adoFy.Recordset.Update "khmc", adoFy.Recordset.Fields("fyNR").Value
End If
End Sub

Private Sub dtgFy_ButtonClick(ByVal ColIndex As Integer)
'If adoFy.Recordset.RecordCount > 0 Then
frmFy.Visible = True
'frmGzNr.SSTab1.TabVisible(1) = False
'End If
End Sub

Private Sub dtgRen_Click()
Dim tt As String
On Error Resume Next
dtgRen.Col = 1
tt = "select tnr,ren from xmren where llid=" & Val(dtgRen.Text)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
txtJw.Text = mod1.HTP.Fields("tnr").Value
lblRen.Caption = Trim(mod1.HTP.Fields("ren").Value)
End Sub

Private Sub dtgRen_RowColChange()
Dim tt As String
On Error Resume Next
dtgRen.Col = 1
tt = "select tnr,ren from xmren where llid=" & Val(dtgRen.Text)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
txtJw.Text = mod1.HTP.Fields("tnr").Value
lblRen.Caption = Trim(mod1.HTP.Fields("ren").Value)
End Sub


Private Sub Form_Load()
frmGzNr.Width = mod1.FWidth
frmGzNr.Height = mod1.FHeight
frmGzNr.Left = 0
frmGzNr.Top = 0
Set AdoKh = CreateObject("adodb.recordset")
Set adoLxr = CreateObject("adodb.recordset")
Set adoBlx = CreateObject("adodb.recordset")
dtgRen.ColWidth(1) = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmFy.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim tt As String
Dim ii As Integer
On Error Resume Next
If MDI.Cq = False Then
    If cmdSave.Enabled = True Then
    ii = MsgBox("退出将不保存数据！", vbYesNo + vbInformation, "当心！")
        If ii = vbNo Then Exit Sub
    End If

frmZu.Enabled = True

'If txtBfJg.Text = "" And modXmGz.Ti = True Then
'    tt = "delete from xmgz where gid=" & modXmGz.Gid
'    adoXmgz.Recordset.Close
'    adoXmgz.Recordset.Open tt, mod1.workKK, adOpenKeyset, adCmdText
''    adoXmgz.Recordset.Delete adAffectCurrent
''    adoXmgz.Recordset.UpdateBatch
'
'End If

    If Dialog.Visible = True Then
        Dialog.Enabled = True
        Dialog.ZOrder 0
    ElseIf frmGzBG.Visible = True Then
        frmGzBG.Enabled = True
        frmGzBG.ZOrder 0
        
    End If
    frmGzNr.Visible = False
    Cancel = True
End If
End Sub

Private Sub lblPd_Click()

End Sub

Private Sub frmMod_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmFy.Visible = False
End Sub

Private Sub optCL_Click()
adoFy.Recordset.Update "fyLB", optCL.Caption
adoFy.Recordset.Update "HTQF", 0
adoFy.Recordset.Update "khmc", ""
optFh.Value = False
frmFy.Visible = False
End Sub

Private Sub optFa_Click()
'On Error Resume Next
adoFy.Recordset.Update "fyLB", optFa.Caption
adoFy.Recordset.Update "HTQF", 0
adoFy.Recordset.Update "khmc", lblXmmc.Caption
optFa.Value = False
frmFy.Visible = False
End Sub

Private Sub optFah_Click()
adoFy.Recordset.Update "fyLB", optFa.Caption
adoFy.Recordset.Update "HTQF", 1
optFa.Value = False
frmFy.Visible = False
End Sub


Private Sub optFb_Click()
adoFy.Recordset.Update "fyLB", optFb.Caption
adoFy.Recordset.Update "HTQF", 0
adoFy.Recordset.Update "khmc", lblXmmc.Caption
optFb.Value = False
frmFy.Visible = False
End Sub


Private Sub optFbh_Click()
adoFy.Recordset.Update "fyLB", optFb.Caption
adoFy.Recordset.Update "HTQF", 1
optFb.Value = False
frmFy.Visible = False
End Sub


Private Sub optFc_Click()
adoFy.Recordset.Update "fyLB", optFc.Caption
adoFy.Recordset.Update "HTQF", 0
adoFy.Recordset.Update "khmc", lblXmmc.Caption
optFc.Value = False
frmFy.Visible = False
End Sub


Private Sub optFch_Click()
adoFy.Recordset.Update "fyLB", optFc.Caption
adoFy.Recordset.Update "HTQF", 1
optFc.Value = False
frmFy.Visible = False
End Sub


Private Sub optFd_Click()
adoFy.Recordset.Update "fyLB", optFd.Caption
adoFy.Recordset.Update "HTQF", 0
adoFy.Recordset.Update "khmc", lblXmmc.Caption
optFd.Value = False
frmFy.Visible = False
End Sub


Private Sub optFdh_Click()
adoFy.Recordset.Update "fyLB", optFd.Caption
adoFy.Recordset.Update "HTQF", 1
optFd.Value = False
frmFy.Visible = False
End Sub


Private Sub optFe_Click()
adoFy.Recordset.Update "fyLB", optFe.Caption
adoFy.Recordset.Update "HTQF", 0
adoFy.Recordset.Update "khmc", lblXmmc.Caption
optFe.Value = False
frmFy.Visible = False
End Sub


Private Sub optFeh_Click()
adoFy.Recordset.Update "fyLB", optFe.Caption
adoFy.Recordset.Update "HTQF", 1
optFe.Value = False
frmFy.Visible = False
End Sub


Private Sub optFf_Click()
adoFy.Recordset.Update "fyLB", optFf.Caption
adoFy.Recordset.Update "HTQF", 0
adoFy.Recordset.Update "khmc", lblXmmc.Caption
optFf.Value = False
frmFy.Visible = False
End Sub


Private Sub optFg_Click()
adoFy.Recordset.Update "fyLB", optFg.Caption

optFg.Value = False
frmFy.Visible = False
End Sub


Private Sub optFFh_Click()
adoFy.Recordset.Update "fyLB", optFf.Caption
adoFy.Recordset.Update "HTQF", 1
optFf.Value = False
frmFy.Visible = False
End Sub

Private Sub optFh_Click()
adoFy.Recordset.Update "fyLB", optFh.Caption
adoFy.Recordset.Update "HTQF", 0
adoFy.Recordset.Update "khmc", ""
optFh.Value = False
frmFy.Visible = False
End Sub


Private Sub optFhh_Click()
adoFy.Recordset.Update "fyLB", optFh.Caption
adoFy.Recordset.Update "HTQF", 1
optFh.Value = False
frmFy.Visible = False
End Sub


Private Sub SSTab1_DblClick()

End Sub

Private Sub txtJw_KeyDown(KeyCode As Integer, Shift As Integer)
If lblRen.Caption = "" Then
    MsgBox "请选择交往客户的姓名!"
    txtJw.Text = ""
    Exit Sub
End If
End Sub

Private Sub txtzgPd_Change()
'如果主管
'If txtzgPd.Text <> "" Then
    lblZGQZ.Caption = mod1.DName
    cmdSave.Enabled = True
'End If
End Sub

Private Sub txtzgPd_Click()
    If mod1.KhK = 1 Or mod1.KhK = 2 Then
    lblZGQZ.Caption = mod1.DName
    End If
End Sub
