VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form form2Htp 
   Caption         =   "合同评审"
   ClientHeight    =   9090
   ClientLeft      =   2715
   ClientTop       =   1440
   ClientWidth     =   9795
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   9795
   Visible         =   0   'False
   Begin VB.Timer timYj 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdCount 
      Caption         =   "计算"
      Height          =   315
      Left            =   8010
      TabIndex        =   162
      Top             =   6660
      Width           =   945
   End
   Begin VB.TextBox txtTcBe 
      Height          =   285
      Left            =   8220
      Locked          =   -1  'True
      TabIndex        =   159
      Text            =   "6"
      Top             =   7260
      Visible         =   0   'False
      Width           =   315
   End
   Begin MSComCtl2.UpDown UpDa 
      Height          =   315
      Left            =   8580
      TabIndex        =   160
      Top             =   7260
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtFj 
      Height          =   300
      Left            =   6330
      TabIndex        =   151
      Top             =   7650
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   5460
      TabIndex        =   150
      Text            =   "附加成本"
      Top             =   7680
      Width           =   765
   End
   Begin VB.Frame frmVCB 
      Caption         =   "请您审阅"
      Height          =   3555
      Left            =   3780
      TabIndex        =   123
      Top             =   4800
      Visible         =   0   'False
      Width           =   5565
   End
   Begin VB.Frame frmZt 
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   885
      Left            =   4920
      TabIndex        =   14
      Top             =   8160
      Width           =   1305
      Begin VB.OptionButton optG 
         Caption         =   "已 盖 章"
         Height          =   195
         Left            =   30
         TabIndex        =   114
         Top             =   270
         Width           =   1035
      End
      Begin VB.OptionButton optW 
         Caption         =   "执行完毕"
         Height          =   225
         Left            =   30
         TabIndex        =   17
         Top             =   690
         Width           =   1035
      End
      Begin VB.OptionButton optZ 
         Caption         =   "执行阶段"
         Height          =   225
         Left            =   30
         TabIndex        =   16
         Top             =   480
         Width           =   1035
      End
      Begin VB.OptionButton optP 
         BackColor       =   &H00C0FFFF&
         Caption         =   "评审阶段"
         Height          =   225
         Left            =   30
         TabIndex        =   15
         Top             =   30
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.CheckBox chkE 
      Height          =   345
      Left            =   990
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   8280
      Width           =   975
   End
   Begin MSAdodcLib.Adodc adoSale 
      Height          =   405
      Left            =   1740
      Top             =   0
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
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
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除"
      Enabled         =   0   'False
      Height          =   555
      Left            =   8430
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8520
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "修改"
      Height          =   555
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8520
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.CommandButton Command3 
      Caption         =   "返回"
      Height          =   555
      Left            =   9150
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8520
      Width           =   645
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存"
      Height          =   555
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "合同"
      Height          =   555
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8520
      Width           =   705
   End
   Begin VB.CheckBox chkD 
      Height          =   345
      Left            =   3990
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8280
      Width           =   915
   End
   Begin VB.CheckBox chkC 
      Height          =   345
      Left            =   1995
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8280
      Width           =   915
   End
   Begin VB.CheckBox chkB 
      Height          =   345
      Left            =   2970
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8280
      Width           =   945
   End
   Begin VB.CheckBox chkA 
      Height          =   345
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8280
      Width           =   945
   End
   Begin TabDlg.SSTab tabHt 
      Height          =   7455
      Left            =   30
      TabIndex        =   18
      Top             =   450
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   13150
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "评审"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdWb"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdFkQ"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frmKP"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtTcRQ"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "产品销售"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab1"
      Tab(1).Control(1)=   "cmdKuDel"
      Tab(1).Control(2)=   "cmdKuAdd"
      Tab(1).Control(3)=   "dtgCG"
      Tab(1).Control(4)=   "txtJhq"
      Tab(1).Control(5)=   "adoKu"
      Tab(1).Control(6)=   "dtPJhq"
      Tab(1).Control(7)=   "cmdDDH"
      Tab(1).Control(8)=   "dtgKu"
      Tab(1).Control(9)=   "Label43"
      Tab(1).Control(10)=   "Label42"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "派工信息"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdGzd"
      Tab(2).Control(1)=   "dtgGzb"
      Tab(2).Control(2)=   "adoLj"
      Tab(2).Control(3)=   "adoGzb"
      Tab(2).Control(4)=   "dtgLj"
      Tab(2).Control(5)=   "Label41"
      Tab(2).Control(6)=   "lblzTime"
      Tab(2).Control(7)=   "Label39"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "分包及其它明细"
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "DataGrid1"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "附加成本明细"
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "DataGrid2"
      Tab(4).ControlCount=   1
      Begin VB.TextBox txtTcRQ 
         Height          =   315
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   163
         Text            =   "提成取现日期"
         Top             =   7170
         Width           =   1845
      End
      Begin VB.Frame frmKP 
         Caption         =   "开票方式"
         Height          =   1695
         Left            =   -90
         TabIndex        =   126
         Top             =   5280
         Width           =   4215
         Begin VB.CommandButton cmdGB 
            Caption         =   "关闭"
            Height          =   315
            Left            =   3600
            TabIndex        =   133
            Top             =   1320
            Width           =   555
         End
         Begin VB.OptionButton optLE 
            Caption         =   "不开票"
            Height          =   255
            Left            =   1410
            TabIndex        =   132
            Top             =   900
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.OptionButton optLD 
            Caption         =   "其它"
            Height          =   285
            Left            =   330
            TabIndex        =   131
            Top             =   870
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.OptionButton optLa 
            Caption         =   "增值发票"
            Height          =   195
            Left            =   330
            TabIndex        =   130
            Tag             =   "17.85"
            Top             =   390
            Width           =   1065
         End
         Begin VB.OptionButton optLb 
            Caption         =   "商业发票"
            Height          =   195
            Left            =   1410
            TabIndex        =   129
            Tag             =   "17.85"
            Top             =   390
            Width           =   1065
         End
         Begin VB.OptionButton optLc 
            Caption         =   "服务发票"
            Height          =   195
            Left            =   2520
            TabIndex        =   128
            Tag             =   "5.25"
            Top             =   390
            Width           =   1065
         End
         Begin VB.CheckBox chkDzf 
            Caption         =   "到账否"
            Height          =   225
            Left            =   2550
            TabIndex        =   127
            Top             =   900
            Visible         =   0   'False
            Width           =   885
         End
      End
      Begin VB.CommandButton cmdFkQ 
         Caption         =   "付款情况"
         Height          =   465
         Left            =   9180
         TabIndex        =   125
         Top             =   1590
         Width           =   555
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   2385
         Left            =   -75000
         TabIndex        =   96
         Top             =   30
         Width           =   9825
         _ExtentX        =   17330
         _ExtentY        =   4207
         _Version        =   393216
         Tab             =   2
         TabHeight       =   520
         BackColor       =   -2147483626
         TabCaption(0)   =   "销售价"
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "dtgSale"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "预计成本价"
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "dtgYj"
         Tab(1).Control(1)=   "cmdChg"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "实际成本价"
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "dtgZj"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "cmdZhg"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).ControlCount=   2
         Begin VB.CommandButton cmdZhg 
            Caption         =   "计算实际材料成本"
            Height          =   285
            Left            =   30
            TabIndex        =   117
            Top             =   2040
            Width           =   1785
         End
         Begin VB.CommandButton cmdChg 
            Caption         =   "计算预计材料成本"
            Height          =   285
            Left            =   -74970
            TabIndex        =   116
            Top             =   2040
            Width           =   1785
         End
         Begin MSDataGridLib.DataGrid dtgSale 
            Bindings        =   "form2HtpOld.frx":0000
            Height          =   2055
            Left            =   -75000
            TabIndex        =   97
            Top             =   300
            Width           =   9795
            _ExtentX        =   17277
            _ExtentY        =   3625
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   20
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
            ColumnCount     =   12
            BeginProperty Column00 
               DataField       =   "xRq"
               Caption         =   "xRq"
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
               DataField       =   "ljMc"
               Caption         =   "产品名称"
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
               DataField       =   "phBiao"
               Caption         =   "牌号商标"
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
               DataField       =   "ljBh"
               Caption         =   "规格型号"
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
            BeginProperty Column04 
               DataField       =   "jlDw"
               Caption         =   "计量单位"
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
            BeginProperty Column05 
               DataField       =   "ljSl"
               Caption         =   "数量"
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
            BeginProperty Column06 
               DataField       =   "dj"
               Caption         =   "单价"
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
            BeginProperty Column07 
               DataField       =   "je"
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
            BeginProperty Column08 
               DataField       =   "Hg"
               Caption         =   "合计"
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
            BeginProperty Column09 
               DataField       =   "xsRy"
               Caption         =   "xsRy"
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
            BeginProperty Column10 
               DataField       =   "shFw"
               Caption         =   "shFw"
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
            BeginProperty Column11 
               DataField       =   "ID"
               Caption         =   "ID"
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
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   3000.189
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2399.811
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1200.189
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   2099.906
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   705.26
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column08 
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column09 
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1365.165
               EndProperty
               BeginProperty Column10 
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1365.165
               EndProperty
               BeginProperty Column11 
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1094.74
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dtgYj 
            Bindings        =   "form2HtpOld.frx":007C
            Height          =   2055
            Left            =   -75000
            TabIndex        =   98
            Top             =   300
            Width           =   9795
            _ExtentX        =   17277
            _ExtentY        =   3625
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   13828070
            HeadLines       =   1
            RowHeight       =   20
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
            ColumnCount     =   12
            BeginProperty Column00 
               DataField       =   "xRq"
               Caption         =   "xRq"
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
               DataField       =   "ljMc"
               Caption         =   "产品名称"
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
               DataField       =   "phBiao"
               Caption         =   "牌号商标"
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
               DataField       =   "ljBh"
               Caption         =   "规格型号"
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
            BeginProperty Column04 
               DataField       =   "jlDw"
               Caption         =   "计量单位"
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
            BeginProperty Column05 
               DataField       =   "xGSlD"
               Caption         =   "数量"
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
            BeginProperty Column06 
               DataField       =   "YJdj"
               Caption         =   "单价"
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
            BeginProperty Column07 
               DataField       =   "YJje"
               Caption         =   "金额"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column08 
               DataField       =   "Hg"
               Caption         =   "合计"
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
            BeginProperty Column09 
               DataField       =   "xsRy"
               Caption         =   "xsRy"
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
            BeginProperty Column10 
               DataField       =   "shFw"
               Caption         =   "shFw"
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
            BeginProperty Column11 
               DataField       =   "ID"
               Caption         =   "ID"
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
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   3000.189
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   2399.811
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
                  ColumnWidth     =   1200.189
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  ColumnWidth     =   2099.906
               EndProperty
               BeginProperty Column04 
                  Locked          =   -1  'True
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column05 
                  Locked          =   -1  'True
                  ColumnWidth     =   705.26
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column07 
                  Locked          =   -1  'True
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column08 
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column09 
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1365.165
               EndProperty
               BeginProperty Column10 
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1365.165
               EndProperty
               BeginProperty Column11 
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1094.74
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dtgZj 
            Bindings        =   "form2HtpOld.frx":00FA
            Height          =   2055
            Left            =   0
            TabIndex        =   99
            Top             =   300
            Width           =   9795
            _ExtentX        =   17277
            _ExtentY        =   3625
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   14673919
            HeadLines       =   1
            RowHeight       =   20
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
            ColumnCount     =   12
            BeginProperty Column00 
               DataField       =   "xRq"
               Caption         =   "xRq"
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
               DataField       =   "ljMc"
               Caption         =   "产品名称"
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
               DataField       =   "phBiao"
               Caption         =   "牌号商标"
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
               DataField       =   "ljBh"
               Caption         =   "规格型号"
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
            BeginProperty Column04 
               DataField       =   "jlDw"
               Caption         =   "计量单位"
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
            BeginProperty Column05 
               DataField       =   "xGSlD"
               Caption         =   "数量"
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
            BeginProperty Column06 
               DataField       =   "ZJdj"
               Caption         =   "单价"
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
            BeginProperty Column07 
               DataField       =   "ZJje"
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
            BeginProperty Column08 
               DataField       =   "Hg"
               Caption         =   "合计"
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
            BeginProperty Column09 
               DataField       =   "xsRy"
               Caption         =   "xsRy"
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
            BeginProperty Column10 
               DataField       =   "shFw"
               Caption         =   "shFw"
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
            BeginProperty Column11 
               DataField       =   "ID"
               Caption         =   "ID"
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
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   3000.189
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   2399.811
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
                  ColumnWidth     =   1200.189
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  ColumnWidth     =   2099.906
               EndProperty
               BeginProperty Column04 
                  Locked          =   -1  'True
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column05 
                  Locked          =   -1  'True
                  ColumnWidth     =   705.26
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column07 
                  Locked          =   -1  'True
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column08 
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column09 
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1365.165
               EndProperty
               BeginProperty Column10 
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1365.165
               EndProperty
               BeginProperty Column11 
                  ColumnAllowSizing=   0   'False
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1094.74
               EndProperty
            EndProperty
         End
      End
      Begin VB.CommandButton cmdKuDel 
         Caption         =   "删除"
         Height          =   285
         Left            =   -66390
         TabIndex        =   104
         Top             =   3810
         Width           =   675
      End
      Begin VB.CommandButton cmdKuAdd 
         Caption         =   "添加"
         Enabled         =   0   'False
         Height          =   285
         Left            =   -67140
         TabIndex        =   103
         Top             =   3810
         Width           =   735
      End
      Begin VB.CommandButton cmdWb 
         Caption         =   "客户档案"
         Height          =   795
         Left            =   3000
         TabIndex        =   95
         Top             =   1140
         Visible         =   0   'False
         Width           =   1275
      End
      Begin MSDataGridLib.DataGrid dtgCG 
         Height          =   2055
         Left            =   -75000
         TabIndex        =   92
         Top             =   4080
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   3625
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         Caption         =   "对应采购合同"
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "ljMc"
            Caption         =   "产品名称"
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
            DataField       =   "DDH"
            Caption         =   "订单号"
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
            DataField       =   "zDhQ"
            Caption         =   "到货期"
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
               ColumnWidth     =   3000.189
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1904.882
            EndProperty
            BeginProperty Column02 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtJhq 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -67830
         TabIndex        =   91
         Top             =   6210
         Visible         =   0   'False
         Width           =   1305
      End
      Begin MSAdodcLib.Adodc adoKu 
         Height          =   465
         Left            =   -71850
         Top             =   6450
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   820
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\WORK\demo\work.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\WORK\demo\work.mdb;Persist Security Info=False"
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
      Begin MSComCtl2.DTPicker dtPJhq 
         Height          =   315
         Left            =   -67860
         TabIndex        =   90
         Top             =   6180
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         Format          =   83165185
         CurrentDate     =   38180
      End
      Begin VB.CommandButton cmdDDH 
         Height          =   315
         Left            =   8250
         TabIndex        =   88
         Top             =   -360
         Width           =   1605
      End
      Begin VB.CommandButton cmdGzd 
         Height          =   345
         Left            =   -69510
         TabIndex        =   85
         Top             =   6510
         Width           =   1905
      End
      Begin MSDataGridLib.DataGrid dtgGzb 
         Bindings        =   "form2HtpOld.frx":0177
         Height          =   3825
         Left            =   -74970
         TabIndex        =   81
         Top             =   30
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   6747
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   14
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
         ColumnCount     =   21
         BeginProperty Column00 
            DataField       =   "rq"
            Caption         =   "日期"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dddddd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "gzNr"
            Caption         =   "gzNr"
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
            DataField       =   "wxWorker"
            Caption         =   "工作人员"
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
            DataField       =   "gzQk"
            Caption         =   "gzQk"
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
         BeginProperty Column04 
            DataField       =   "khMc"
            Caption         =   "khMc"
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
         BeginProperty Column05 
            DataField       =   "qy"
            Caption         =   "qy"
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
         BeginProperty Column06 
            DataField       =   "khPho"
            Caption         =   "khPho"
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
         BeginProperty Column07 
            DataField       =   "gzWf"
            Caption         =   "完工否"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "完成"
               FalseValue      =   "未完成"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "gdTime"
            Caption         =   "gdTime"
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
         BeginProperty Column09 
            DataField       =   "lkTime"
            Caption         =   "lkTime"
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
         BeginProperty Column10 
            DataField       =   "ltTime"
            Caption         =   "ltTime"
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
         BeginProperty Column11 
            DataField       =   "qtTime"
            Caption         =   "qtTime"
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
         BeginProperty Column12 
            DataField       =   "khQm"
            Caption         =   "khQm"
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
         BeginProperty Column13 
            DataField       =   "qmRq"
            Caption         =   "qmRq"
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
         BeginProperty Column14 
            DataField       =   "xzZg"
            Caption         =   "xzZg"
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
         BeginProperty Column15 
            DataField       =   "jsZj"
            Caption         =   "jsZj"
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
         BeginProperty Column16 
            DataField       =   "shFw"
            Caption         =   "shFw"
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
         BeginProperty Column17 
            DataField       =   "bhId"
            Caption         =   "工作单编号"
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
         BeginProperty Column18 
            DataField       =   "Ztime"
            Caption         =   "Ztime"
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
         BeginProperty Column19 
            DataField       =   "qtQing"
            Caption         =   "qtQing"
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
         BeginProperty Column20 
            DataField       =   "htBh"
            Caption         =   "htBh"
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
            Size            =   10
            BeginProperty Column00 
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3000.189
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column11 
               Object.Visible         =   0   'False
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column13 
               Object.Visible         =   0   'False
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   0   'False
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column16 
               Object.Visible         =   0   'False
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column17 
               Alignment       =   2
               ColumnWidth     =   1800
            EndProperty
            BeginProperty Column18 
               Object.Visible         =   0   'False
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column19 
               Object.Visible         =   0   'False
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column20 
               Object.Visible         =   0   'False
               ColumnWidth     =   2085.166
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc adoLj 
         Height          =   495
         Left            =   -65940
         Top             =   5070
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   873
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
         Caption         =   "Adodc2"
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
      Begin MSAdodcLib.Adodc adoGzb 
         Height          =   465
         Left            =   -65940
         Top             =   4380
         Visible         =   0   'False
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   820
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
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   7065
         Left            =   0
         TabIndex        =   19
         Top             =   60
         Width           =   8925
         Begin VB.TextBox txtXMMC 
            Height          =   285
            Left            =   1290
            TabIndex        =   111
            Top             =   30
            Width           =   2895
         End
         Begin VB.ComboBox comQy 
            Height          =   300
            ItemData        =   "form2HtpOld.frx":01C0
            Left            =   5730
            List            =   "form2HtpOld.frx":01C2
            Locked          =   -1  'True
            TabIndex        =   94
            Text            =   "comQy"
            Top             =   60
            Width           =   1005
         End
         Begin VB.TextBox txtKhmc 
            Height          =   315
            Left            =   1290
            TabIndex        =   70
            Top             =   330
            Width           =   2895
         End
         Begin VB.TextBox txtYwy 
            Height          =   315
            Left            =   7650
            Locked          =   -1  'True
            TabIndex        =   69
            Top             =   60
            Width           =   1275
         End
         Begin VB.TextBox txtHtbh 
            Height          =   270
            Left            =   1290
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   720
            Width           =   2895
         End
         Begin VB.OptionButton optA 
            Caption         =   "A. 零配件合同"
            Height          =   195
            Index           =   0
            Left            =   1320
            TabIndex        =   67
            Top             =   1110
            Width           =   1485
         End
         Begin VB.OptionButton optA 
            Caption         =   "B1.工程合同"
            Height          =   225
            Index           =   1
            Left            =   1320
            TabIndex        =   66
            Top             =   1710
            Width           =   1335
         End
         Begin VB.TextBox txtMon 
            Height          =   270
            Left            =   2580
            TabIndex        =   65
            Top             =   3240
            Width           =   675
         End
         Begin VB.Frame Frame4 
            Height          =   1335
            Left            =   -300
            TabIndex        =   59
            Top             =   1920
            Width           =   9855
            Begin VB.TextBox txtTian 
               Height          =   270
               Left            =   4110
               TabIndex        =   61
               Top             =   180
               Width           =   555
            End
            Begin VB.TextBox txtJhqk 
               Height          =   750
               Left            =   1530
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   60
               Top             =   480
               Width           =   7695
            End
            Begin VB.Label Label6 
               Caption         =   "*"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   1230
               TabIndex        =   108
               Top             =   540
               Width           =   195
            End
            Begin VB.Label Label32 
               Caption         =   "交货条件或工期情况："
               Height          =   435
               Left            =   330
               TabIndex        =   64
               Top             =   390
               Width           =   915
            End
            Begin VB.Label Label33 
               Caption         =   "自签订合同（并收订金）之日起 "
               Height          =   195
               Left            =   1530
               TabIndex        =   63
               Top             =   210
               Width           =   2685
            End
            Begin VB.Label Label34 
               Caption         =   "天"
               Height          =   165
               Left            =   4740
               TabIndex        =   62
               Top             =   210
               Width           =   465
            End
         End
         Begin VB.Frame Frame3 
            Height          =   3645
            Left            =   -300
            TabIndex        =   24
            Top             =   3510
            Width           =   9855
            Begin VB.TextBox txtRgf 
               Height          =   315
               Left            =   1560
               TabIndex        =   55
               Top             =   900
               Visible         =   0   'False
               Width           =   1875
            End
            Begin VB.TextBox txtClf 
               Height          =   285
               Left            =   1560
               TabIndex        =   54
               Top             =   600
               Visible         =   0   'False
               Width           =   1875
            End
            Begin VB.TextBox txtHtze 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   1
               EndProperty
               Height          =   285
               Left            =   1560
               TabIndex        =   53
               Top             =   300
               Width           =   1875
            End
            Begin VB.Frame Frame5 
               Caption         =   "由运作部负责填写"
               Height          =   5415
               Left            =   4410
               TabIndex        =   36
               Top             =   -30
               Width           =   5445
               Begin VB.TextBox txtFbje1 
                  Height          =   315
                  Left            =   1440
                  TabIndex        =   149
                  Top             =   2070
                  Width           =   1155
               End
               Begin VB.Frame frmXz 
                  BorderStyle     =   0  'None
                  Height          =   3375
                  Left            =   2610
                  TabIndex        =   137
                  Top             =   180
                  Width           =   2715
                  Begin VB.TextBox txtCBze3 
                     Height          =   285
                     Left            =   1260
                     TabIndex        =   158
                     Top             =   360
                     Width           =   945
                  End
                  Begin VB.TextBox txtFbje3 
                     Height          =   285
                     Left            =   1260
                     TabIndex        =   156
                     Top             =   1890
                     Width           =   945
                  End
                  Begin VB.TextBox txtYf3 
                     Height          =   285
                     Left            =   1260
                     TabIndex        =   155
                     Top             =   1590
                     Width           =   945
                  End
                  Begin VB.TextBox txtZXF3 
                     Height          =   285
                     Left            =   1260
                     TabIndex        =   154
                     Top             =   1260
                     Width           =   945
                  End
                  Begin VB.TextBox txtQT3 
                     Height          =   285
                     Left            =   1260
                     TabIndex        =   153
                     Top             =   960
                     Width           =   945
                  End
                  Begin VB.TextBox txtClcb3 
                     Height          =   285
                     Left            =   1260
                     TabIndex        =   152
                     Top             =   660
                     Width           =   945
                  End
                  Begin VB.TextBox txtTc2 
                     Height          =   285
                     Left            =   30
                     Locked          =   -1  'True
                     TabIndex        =   147
                     Top             =   3090
                     Width           =   1185
                  End
                  Begin VB.TextBox txtYj2 
                     Height          =   270
                     Left            =   30
                     Locked          =   -1  'True
                     TabIndex        =   146
                     Top             =   2505
                     Width           =   1185
                  End
                  Begin VB.TextBox txtLr2 
                     BeginProperty Font 
                        Name            =   "宋体"
                        Size            =   9
                        Charset         =   134
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   270
                     Left            =   30
                     Locked          =   -1  'True
                     TabIndex        =   145
                     Top             =   2820
                     Width           =   1185
                  End
                  Begin VB.TextBox txtCbze2 
                     Height          =   285
                     Left            =   30
                     Locked          =   -1  'True
                     TabIndex        =   144
                     Top             =   360
                     Width           =   1185
                  End
                  Begin VB.TextBox txtClcb2 
                     Height          =   285
                     Left            =   30
                     Locked          =   -1  'True
                     TabIndex        =   143
                     Top             =   645
                     Width           =   1185
                  End
                  Begin VB.TextBox txtYf2 
                     Height          =   270
                     Left            =   30
                     Locked          =   -1  'True
                     TabIndex        =   142
                     Top             =   1575
                     Width           =   1185
                  End
                  Begin VB.TextBox txtQt2 
                     Enabled         =   0   'False
                     Height          =   270
                     Left            =   30
                     Locked          =   -1  'True
                     TabIndex        =   141
                     ToolTipText     =   "双击此处可以看项目费用清单"
                     Top             =   945
                     Width           =   1185
                  End
                  Begin VB.TextBox txtJlr2 
                     Height          =   285
                     Left            =   30
                     Locked          =   -1  'True
                     TabIndex        =   140
                     Top             =   2220
                     Width           =   1185
                  End
                  Begin VB.TextBox txtZxF2 
                     Height          =   315
                     Left            =   30
                     Locked          =   -1  'True
                     TabIndex        =   139
                     Top             =   1260
                     Width           =   1185
                  End
                  Begin VB.TextBox txtFbje2 
                     Height          =   285
                     Left            =   30
                     Locked          =   -1  'True
                     TabIndex        =   138
                     Top             =   1890
                     Width           =   1185
                  End
                  Begin VB.Label lblTcBe 
                     Caption         =   "提成比例"
                     Height          =   195
                     Left            =   1380
                     TabIndex        =   161
                     Top             =   2880
                     Visible         =   0   'False
                     Width           =   735
                  End
                  Begin VB.Label Label18 
                     Caption         =   "备注"
                     Height          =   285
                     Left            =   1470
                     TabIndex        =   157
                     Top             =   30
                     Width           =   525
                  End
                  Begin VB.Label Label8 
                     Caption         =   "实  际"
                     Height          =   225
                     Left            =   300
                     TabIndex        =   148
                     Top             =   30
                     Width           =   585
                  End
               End
               Begin VB.TextBox txtZXF1 
                  Height          =   300
                  Left            =   1440
                  TabIndex        =   113
                  Top             =   1440
                  Width           =   1155
               End
               Begin VB.TextBox txtJlr1 
                  Height          =   270
                  Left            =   1440
                  TabIndex        =   106
                  Top             =   2400
                  Width           =   1155
               End
               Begin VB.TextBox txtTc1 
                  Height          =   285
                  Left            =   1440
                  TabIndex        =   44
                  Top             =   3270
                  Width           =   1155
               End
               Begin VB.TextBox txtYj1 
                  Height          =   270
                  Left            =   1440
                  TabIndex        =   43
                  Top             =   2685
                  Width           =   1155
               End
               Begin VB.TextBox txtLr1 
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   9
                     Charset         =   134
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Left            =   1440
                  TabIndex        =   42
                  Top             =   3000
                  Width           =   1155
               End
               Begin VB.TextBox txtCbze1 
                  Height          =   285
                  Left            =   1440
                  Locked          =   -1  'True
                  TabIndex        =   41
                  Top             =   540
                  Width           =   1155
               End
               Begin VB.TextBox txtClcb1 
                  Height          =   285
                  Left            =   1440
                  TabIndex        =   40
                  Top             =   825
                  Width           =   1155
               End
               Begin VB.TextBox txtYf1 
                  Height          =   270
                  Left            =   1440
                  TabIndex        =   39
                  Top             =   1770
                  Width           =   1155
               End
               Begin VB.TextBox txtQt1 
                  Height          =   270
                  Left            =   1440
                  TabIndex        =   38
                  Top             =   1125
                  Width           =   1155
               End
               Begin VB.CommandButton cmdJi 
                  Caption         =   "计算2"
                  Height          =   285
                  Left            =   330
                  TabIndex        =   37
                  Top             =   210
                  Width           =   675
               End
               Begin VB.Label Label27 
                  Caption         =   "分包及其他"
                  Height          =   225
                  Left            =   270
                  TabIndex        =   124
                  Top             =   2160
                  Width           =   975
               End
               Begin VB.Label Label26 
                  Caption         =   "装 卸 费"
                  Height          =   255
                  Left            =   300
                  TabIndex        =   112
                  Top             =   1500
                  Width           =   765
               End
               Begin VB.Label Label21 
                  Caption         =   "利 润 1"
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Left            =   300
                  TabIndex        =   105
                  Top             =   2490
                  Width           =   795
               End
               Begin VB.Label Label7 
                  Caption         =   "预  计"
                  Height          =   225
                  Left            =   1710
                  TabIndex        =   52
                  Top             =   210
                  Width           =   555
               End
               Begin VB.Label lblTc 
                  Caption         =   "提    成"
                  Height          =   285
                  Left            =   300
                  TabIndex        =   51
                  Top             =   3330
                  Width           =   735
               End
               Begin VB.Label lblYj 
                  Caption         =   "奖   金"
                  Height          =   225
                  Left            =   300
                  TabIndex        =   50
                  Top             =   2790
                  Width           =   975
               End
               Begin VB.Label lblLr2 
                  Caption         =   "利 润 2"
                  ForeColor       =   &H000000FF&
                  Height          =   195
                  Left            =   300
                  TabIndex        =   49
                  Top             =   3090
                  Width           =   915
               End
               Begin VB.Label Label16 
                  Caption         =   "成本总额(2)"
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
                  TabIndex        =   48
                  Top             =   600
                  Width           =   1155
               End
               Begin VB.Label Label17 
                  Caption         =   "材料成本"
                  Height          =   225
                  Left            =   300
                  TabIndex        =   47
                  Top             =   870
                  Width           =   825
               End
               Begin VB.Label Label19 
                  Caption         =   "运    费"
                  Height          =   195
                  Left            =   300
                  TabIndex        =   46
                  Top             =   1830
                  Width           =   915
               End
               Begin VB.Label Label20 
                  Caption         =   "项目费用"
                  Height          =   225
                  Left            =   300
                  TabIndex        =   45
                  Top             =   1200
                  Width           =   945
               End
            End
            Begin VB.TextBox txtAdd1 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   300
               TabIndex        =   35
               Top             =   1230
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.TextBox txtAdd2 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   300
               TabIndex        =   34
               Top             =   1590
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.TextBox txtAdd3 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   300
               TabIndex        =   33
               Top             =   1920
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.TextBox txtAz1 
               Height          =   315
               Left            =   1560
               TabIndex        =   32
               Top             =   1260
               Visible         =   0   'False
               Width           =   1875
            End
            Begin VB.TextBox txtAz2 
               Height          =   315
               Left            =   1560
               TabIndex        =   31
               Top             =   1590
               Visible         =   0   'False
               Width           =   1875
            End
            Begin VB.TextBox txtAz3 
               Height          =   315
               Left            =   1560
               TabIndex        =   30
               Top             =   1920
               Visible         =   0   'False
               Width           =   1875
            End
            Begin VB.TextBox txtAz4 
               Height          =   315
               Left            =   1560
               TabIndex        =   29
               Top             =   2250
               Visible         =   0   'False
               Width           =   1875
            End
            Begin VB.TextBox txtAz5 
               Height          =   315
               Left            =   1560
               TabIndex        =   28
               Top             =   2580
               Visible         =   0   'False
               Width           =   1875
            End
            Begin VB.TextBox txtAdd4 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   300
               TabIndex        =   27
               Top             =   2250
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.TextBox txtAdd5 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   300
               TabIndex        =   26
               Top             =   2580
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.CommandButton Command2 
               Caption         =   "计算1"
               Height          =   285
               Left            =   3570
               TabIndex        =   25
               Top             =   210
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.Label Label15 
               Caption         =   "人 工 费"
               Height          =   285
               Left            =   330
               TabIndex        =   58
               Top             =   990
               Visible         =   0   'False
               Width           =   1065
            End
            Begin VB.Label Label14 
               Caption         =   "材 料 费"
               Height          =   315
               Left            =   330
               TabIndex        =   57
               Top             =   660
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label Label13 
               Caption         =   "合同总额(1)"
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
               Left            =   330
               TabIndex        =   56
               Top             =   390
               Width           =   1425
            End
         End
         Begin VB.TextBox txtHtdate 
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dddddd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   210
            Left            =   5760
            TabIndex        =   22
            Top             =   420
            Width           =   2865
         End
         Begin VB.TextBox txtFkBz 
            Height          =   960
            Left            =   5730
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Top             =   750
            Width           =   3465
         End
         Begin VB.OptionButton optA 
            Caption         =   "E. 产品合同"
            Height          =   195
            Index           =   5
            Left            =   1320
            TabIndex        =   20
            Top             =   1410
            Width           =   1485
         End
         Begin MSComCtl2.DTPicker DT1 
            Height          =   255
            Left            =   5730
            TabIndex        =   23
            Top             =   390
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "yyyy年M月d日"
            Format          =   83165187
            CurrentDate     =   38098.7575810185
         End
         Begin VB.Label Label24 
            Caption         =   "项目名称"
            Height          =   285
            Left            =   30
            TabIndex        =   110
            Top             =   90
            Width           =   795
         End
         Begin VB.Label Label5 
            Caption         =   "*"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   5490
            TabIndex        =   107
            Top             =   810
            Width           =   195
         End
         Begin VB.Label Label44 
            Caption         =   "区    域"
            Height          =   255
            Left            =   4650
            TabIndex        =   93
            Top             =   120
            Width           =   915
         End
         Begin VB.Label lblHtxz 
            Caption         =   "lblHtxz"
            Height          =   285
            Left            =   750
            TabIndex        =   80
            Top             =   540
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "客户名称"
            Height          =   285
            Left            =   30
            TabIndex        =   79
            Top             =   420
            Width           =   1605
         End
         Begin VB.Label Label2 
            Caption         =   "业 务 员"
            Height          =   255
            Left            =   6840
            TabIndex        =   78
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "日    期"
            Height          =   255
            Left            =   4620
            TabIndex        =   77
            Top             =   450
            Width           =   915
         End
         Begin VB.Label Label4 
            Caption         =   "合同性质"
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
            Left            =   30
            TabIndex        =   76
            Top             =   1110
            Width           =   1185
         End
         Begin VB.Label Label25 
            Caption         =   "合同编号"
            Height          =   225
            Left            =   30
            TabIndex        =   75
            Top             =   780
            Width           =   945
         End
         Begin VB.Label Label31 
            Caption         =   "付款条件"
            Height          =   195
            Left            =   4650
            TabIndex        =   74
            Top             =   810
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "保 修 期："
            Height          =   225
            Left            =   30
            TabIndex        =   73
            Top             =   3300
            Width           =   945
         End
         Begin VB.Label Label11 
            Caption         =   "工作完工验收后"
            Height          =   195
            Left            =   1230
            TabIndex        =   72
            Top             =   3300
            Width           =   1365
         End
         Begin VB.Label Label12 
            Caption         =   "月"
            Height          =   255
            Left            =   3450
            TabIndex        =   71
            Top             =   3270
            Width           =   375
         End
      End
      Begin MSDataGridLib.DataGrid dtgLj 
         Height          =   2805
         Left            =   -74970
         TabIndex        =   82
         Top             =   4320
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   4948
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         TabAcrossSplits =   -1  'True
         TabAction       =   2
         WrapCellPointer =   -1  'True
         RowDividerStyle =   5
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
         Caption         =   "零配件消耗"
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "ljMc"
            Caption         =   "零件或材料名称"
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
            DataField       =   "ljBh"
            Caption         =   "零件编号"
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
            DataField       =   "sl"
            Caption         =   "数量"
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
            DataField       =   "danWei"
            Caption         =   "单位"
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
         BeginProperty Column04 
            DataField       =   "yongQing"
            Caption         =   "使用情况"
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
         BeginProperty Column05 
            DataField       =   "gongFang"
            Caption         =   "供方"
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
         BeginProperty Column06 
            DataField       =   "khMc"
            Caption         =   "khMc"
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
         BeginProperty Column07 
            DataField       =   "bhId"
            Caption         =   "bhId"
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
            ScrollBars      =   3
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            RecordSelectors =   0   'False
            BeginProperty Column00 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   1995.024
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
               ColumnWidth     =   1484.787
            EndProperty
            BeginProperty Column06 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column07 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1094.74
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dtgKu 
         Bindings        =   "form2HtpOld.frx":01C4
         Height          =   1395
         Left            =   -75000
         TabIndex        =   102
         Top             =   2400
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   2461
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   12648447
         HeadLines       =   1
         RowHeight       =   20
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
         Caption         =   "取自库存"
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "xRq"
            Caption         =   "xRq"
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
            DataField       =   "ljMc"
            Caption         =   "产品名称"
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
            DataField       =   "phBiao"
            Caption         =   "牌号商标"
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
            DataField       =   "ljBh"
            Caption         =   "规格型号"
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
         BeginProperty Column04 
            DataField       =   "jlDw"
            Caption         =   "计量单位"
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
         BeginProperty Column05 
            DataField       =   "ljSl"
            Caption         =   "数量"
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
         BeginProperty Column06 
            DataField       =   "dj"
            Caption         =   "单价"
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
         BeginProperty Column07 
            DataField       =   "je"
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
         BeginProperty Column08 
            DataField       =   "Hg"
            Caption         =   "合计"
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
         BeginProperty Column09 
            DataField       =   "xsRy"
            Caption         =   "xsRy"
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
         BeginProperty Column10 
            DataField       =   "shFw"
            Caption         =   "shFw"
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
         BeginProperty Column11 
            DataField       =   "ID"
            Caption         =   "ID"
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
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   3000.189
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   2399.811
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   2099.906
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column08 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column09 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column10 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column11 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1094.74
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   6705
         Left            =   -75000
         TabIndex        =   134
         Top             =   0
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   11827
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   20
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
         Caption         =   "分包及其他明细"
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "xRq"
            Caption         =   "xRq"
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
            DataField       =   "ljMc"
            Caption         =   "产品名称"
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
            DataField       =   "phBiao"
            Caption         =   "牌号商标"
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
            DataField       =   "ljBh"
            Caption         =   "规格型号"
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
         BeginProperty Column04 
            DataField       =   "jlDw"
            Caption         =   "计量单位"
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
         BeginProperty Column05 
            DataField       =   "ljSl"
            Caption         =   "数量"
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
         BeginProperty Column06 
            DataField       =   "dj"
            Caption         =   "单价"
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
         BeginProperty Column07 
            DataField       =   "je"
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
         BeginProperty Column08 
            DataField       =   "Hg"
            Caption         =   "合计"
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
         BeginProperty Column09 
            DataField       =   "xsRy"
            Caption         =   "xsRy"
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
         BeginProperty Column10 
            DataField       =   "shFw"
            Caption         =   "shFw"
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
         BeginProperty Column11 
            DataField       =   "ID"
            Caption         =   "ID"
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
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   3000.189
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2399.811
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2099.906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column08 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column09 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column10 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column11 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1094.74
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   6915
         Left            =   -75000
         TabIndex        =   135
         Top             =   0
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   12197
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   20
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
         Caption         =   "附加成本明细"
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "xRq"
            Caption         =   "xRq"
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
            DataField       =   "ljMc"
            Caption         =   "产品名称"
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
            DataField       =   "phBiao"
            Caption         =   "牌号商标"
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
            DataField       =   "ljBh"
            Caption         =   "规格型号"
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
         BeginProperty Column04 
            DataField       =   "jlDw"
            Caption         =   "计量单位"
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
         BeginProperty Column05 
            DataField       =   "ljSl"
            Caption         =   "数量"
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
         BeginProperty Column06 
            DataField       =   "dj"
            Caption         =   "单价"
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
         BeginProperty Column07 
            DataField       =   "je"
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
         BeginProperty Column08 
            DataField       =   "Hg"
            Caption         =   "合计"
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
         BeginProperty Column09 
            DataField       =   "xsRy"
            Caption         =   "xsRy"
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
         BeginProperty Column10 
            DataField       =   "shFw"
            Caption         =   "shFw"
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
         BeginProperty Column11 
            DataField       =   "ID"
            Caption         =   "ID"
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
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   3000.189
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2399.811
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2099.906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column08 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column09 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column10 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column11 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1094.74
            EndProperty
         EndProperty
      End
      Begin VB.Label Label43 
         Caption         =   "交货期："
         Height          =   255
         Left            =   -69180
         TabIndex        =   89
         Top             =   6270
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label42 
         Caption         =   "采购合同评审单"
         Height          =   255
         Left            =   -69240
         TabIndex        =   87
         Top             =   6750
         Width           =   1335
      End
      Begin VB.Label Label41 
         Caption         =   "工作单查询详情："
         Height          =   225
         Left            =   -69390
         TabIndex        =   86
         Top             =   6240
         Width           =   1635
      End
      Begin VB.Label lblzTime 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -66060
         TabIndex        =   84
         Top             =   6540
         Width           =   615
      End
      Begin VB.Label Label39 
         Caption         =   "总工时："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   -67290
         TabIndex        =   83
         Top             =   6510
         Width           =   1155
      End
   End
   Begin VB.Label lblHid 
      Height          =   315
      Left            =   0
      TabIndex        =   164
      Top             =   0
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label lblKhdh 
      Caption         =   "Khdh"
      Height          =   195
      Left            =   7380
      TabIndex        =   136
      Top             =   150
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblZj 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   3960
      TabIndex        =   122
      Top             =   8640
      Width           =   975
   End
   Begin VB.Label lblYz 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   1980
      TabIndex        =   121
      Top             =   8640
      Width           =   975
   End
   Begin VB.Label lblJl 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   2970
      TabIndex        =   120
      Top             =   8640
      Width           =   975
   End
   Begin VB.Label lblYw 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      TabIndex        =   119
      Top             =   8640
      Width           =   975
   End
   Begin VB.Label lblJz 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   990
      TabIndex        =   118
      Top             =   8640
      Width           =   975
   End
   Begin VB.Label lblBM 
      Caption         =   "Label27"
      Height          =   315
      Left            =   6270
      TabIndex        =   115
      Top             =   90
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label23 
      Caption         =   "注：带“*”项必须在合同中修改"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6240
      TabIndex        =   109
      Top             =   8280
      Width           =   2685
   End
   Begin VB.Label Label9 
      Caption         =   "技术支持"
      Height          =   255
      Left            =   1050
      TabIndex        =   100
      Top             =   8040
      Width           =   765
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "产品合同评审单"
      BeginProperty Font 
         Name            =   "华文彩云"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   345
      Left            =   3150
      TabIndex        =   13
      Top             =   60
      Width           =   2535
   End
   Begin VB.Label Label37 
      Caption         =   "总经理"
      Height          =   225
      Left            =   4020
      TabIndex        =   10
      Top             =   8040
      Width           =   735
   End
   Begin VB.Label Label36 
      Caption         =   "商务经理"
      Height          =   255
      Left            =   1995
      TabIndex        =   9
      Top             =   8040
      Width           =   825
   End
   Begin VB.Label Label35 
      Caption         =   "销售经理"
      Height          =   255
      Left            =   3015
      TabIndex        =   8
      Top             =   8040
      Width           =   855
   End
   Begin VB.Label Label22 
      Caption         =   "销售员"
      Height          =   225
      Left            =   90
      TabIndex        =   7
      Top             =   8040
      Width           =   675
   End
End
Attribute VB_Name = "form2Htp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmpDate As Date

Private Sub cmdChg_Click()
On Error Resume Next
Dim TQ As Double
Dim KQ As Double
TQ = 0
KQ = 0
adoSale.Recordset.MoveFirst
Do While Not adoSale.Recordset.EOF
TQ = TQ + adoSale.Recordset.Fields("yjJe").Value
adoSale.Recordset.MoveNext
Loop
adoKu.Recordset.MoveFirst
Do While Not adoKu.Recordset.EOF
KQ = KQ + adoKu.Recordset.Fields("je").Value
adoKu.Recordset.MoveNext
Loop

txtClcb1.Text = TQ + KQ

'计算预计成本总额
txtCbze1.Text = Round(Val(txtClcb1.Text) + Val(txtQt1.Text) + Val(txtYf1.Text) + Val(txtZXF1.Text), 2)
'计算预计利润1
txtJlr1.Text = Round(Val(txtHtze.Text) - Val(txtCbze1.Text), 2)
'计算利润2
txtLr1.Text = Round(Val(txtJlr1.Text) - Val(txtYj1.Text), 2)
'计算预计提成
txtTc1.Text = Round(Val(txtLr1.Text) * 0.08, 2)
End Sub

Private Sub cmdDDH_Click()
Dim tt As String
On Error Resume Next
'tt = "Select * from CGD Where DDH='" & cmdDDH.Caption & "'"
'frmCg.adoCC.Recordset.Close
'frmCg.adoCC.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
Call modCG.BoundCG(cmdDDH.Caption)
frmCg.Show

'需购产品表
tt = "Select * from htSale where  MT=0 and htF=1 and delF=1 and xGSl>0"
frmXG.adoCg.Recordset.Close
frmXG.adoCg.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
Set frmXG.dtgXG.DataSource = frmXG.adoCg
frmXG.lblLjmc.Caption = frmXG.adoCg.Recordset.Fields("ljMc").Value
frmXG.cmdHtbh.Caption = frmXG.adoCg.Recordset.Fields("htBh").Value
frmXG.txtL.Text = frmXG.adoCg.Recordset.Fields("xgSL").Value
End Sub

Private Sub cmdFkQ_Click()
On Error Resume Next
frmFuK.Visible = True
frmFuK.WindowState = 0
frmFuK.ZOrder 0
frmFuK.lblHtze.Caption = txtHtze.Text
'frmFuK.cmdYadd.Enabled = False
'frmFuK.cmdYdel.Enabled = False
'
'frmFuK.dtgYf.AllowUpdate = False
'
'Dim ft As String
'On Error Resume Next
'ft = "Select * from htPing1 where htBh='" & form2Htp.txtHtbh.Text & "' order by rq"
'frmFuK.adoHpt.Recordset.Close
'frmFuK.adoHpt.Recordset.Open ft, , adOpenKeyset, adLockBatchOptimistic, adCmdText
'Set frmFuK.dtgFk.DataSource = frmFuK.adoHpt
'
'If form2Htp.Frame1.Enabled = False Then '如果为非编辑状态
''frmFuK.frmFk.Enabled = False
'frmFuK.cmdMod1.Enabled = False
'Else
'frmFuK.cmdMod1.Enabled = True
'End If
'
''如果为执行的合同，则只有修改已收金额的权利,而且只有刘霁虹和小吴能够改
'If form2Htp.optZ.Value = True Then
'
''frmFuK.dtgFk.Columns(1).Locked = True
''frmFuK.dtgFk.Columns(2).Locked = True
''frmFuK.dtgFk.Columns(4).Locked = True
''frmFuK.dtgFk.Columns(5).Locked = True
'frmFuK.dtgFk.AllowDelete = False
'frmFuK.dtgFk.AllowAddNew = False
'
'If frmLogin.Combo1.Text = "胡颖" Or frmLogin.Combo1.Text = "刘霁虹" Then
''如果为合同执行了，则财务可以修改收款表
'
'frmFuK.cmdMod2.Enabled = True
'
'If frmLogin.Combo1.Text = "胡颖" Then frmFuK.cmdMod1.Enabled = True
'Else '如果为其他人看生成的合同，则不能修改
''frmFuK.dtgYf.Enabled = False
''frmFuK.cmdYadd.Enabled = False
''frmFuK.cmdYdel.Enabled = False
'frmFuK.cmdMod2.Enabled = False
'End If
'frmFuK.adoHpt.Recordset.MoveFirst
'Set frmFuK.dtgFk.DataSource = frmFuK.adoHpt
'
''初始化已收款表显示
'frmFuK.adoHpt.Recordset.MoveFirst
'Dim pT As String
'pT = "Select * from yiFk where htbh='" & frmFuK.adoHpt.Recordset.Fields(3).Value & _
'"' and yingRQ='" & frmFuK.adoHpt.Recordset.Fields(1).Value & "' order by YiRq"
'frmFuK.adoYf.Recordset.Close
'frmFuK.adoYf.Recordset.Open pT, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'Set frmFuK.dtgYf.DataSource = frmFuK.adoYf
'
'Else
'frmFuK.dtgFk.Columns(4).Locked = True
'frmFuK.dtgFk.Columns(5).Locked = True
'frmFuK.dtgFk.Columns(1).Locked = False
'frmFuK.dtgFk.Columns(2).Locked = False
'frmFuK.dtgFk.AllowDelete = True
'Set frmFuK.dtgYf.DataSource = Nothing
'frmFuK.dtgYf.Enabled = False
'frmFuK.cmdYadd.Enabled = False
'frmFuK.cmdYdel.Enabled = False
'frmFuK.cmdMod2.Enabled = False
'End If
'frmFuK.Visible = True
'If frmZu.frm5.Visible = True Then
'End If
'form2Htp.Enabled = False


End Sub

Private Sub cmdGB_Click()
'frmKP.Visible = False

End Sub

Private Sub cmdGzd_Click()
'On Error Resume Next
'Dim tt As String
'Dim PP As String
'Me.MousePointer = 11
'PP = cmdGzd.Caption
'tt = "Select * from gzb" & frmLogin.Combo1.Text & " where bhid='" & PP & " '"
'form2Gc.adoZong1.Recordset.Close
''adoZong.Recordset.Close
'form2Gc.adoHi.Recordset.Close
'form2Gc.adoLj.Recordset.Close
'frmWorkQ.adoWx.Recordset.Close
'form2Gc.adoZong1.Recordset.Open tt, , , adLockBatchOptimistic, adCmdText
''tt = "Select * from jizu where bhid='" & pp & " '"""
'tt = "Select * from jizu" & frmLogin.Combo1.Text & " where bhid='" & PP & " '"
'form2Gc.adoHi.Recordset.Open tt, , , adLockBatchOptimistic, adCmdText
'tt = "Select * from linjian" & frmLogin.Combo1.Text & " where bhid='" & PP & " '"
'form2Gc.adoLj.Recordset.Open tt, , , adLockBatchOptimistic, adCmdText
'tt = "Select * from workXx" & frmLogin.Combo1.Text & " where bhid='" & PP & "'"
'frmWorkQ.adoWx.Recordset.Open tt, mod1.workKK, , adLockBatchOptimistic, adCmdText
'
'Call mod1.GzdRgen
'Set form2Gc.dtgJizu.DataSource = form2Gc.adoHi
'Set form2Gc.dtgB.DataSource = form2Gc.adoHi
'Set form2Gc.dtgC.DataSource = form2Gc.adoHi
'Set form2Gc.dtgLj.DataSource = form2Gc.adoLj
'Set frmWorkQ.dtgWkXx.DataSource = frmWorkQ.adoWx
'Me.MousePointer = 0
'form2Gc.frmWork.Visible = True
'form2Gc.Visible = True
End Sub


Private Sub cmdHt_Click()
Call modHt.gxQing
Call modHt.gxBound
htgX.Show
Set htgX.dtgSale.DataSource = form2Htp.adoSale
End Sub

Private Sub cmdJi_Click()
On Error Resume Next
'计算预计成本总额
txtCbze1.Text = Round(Val(txtClcb1.Text) + Val(txtQt1.Text) + Val(txtYf1.Text) + Val(txtZXF1.Text), 2)
'计算预计利润1
txtJlr1.Text = Round(Val(txtHtze.Text) - Val(txtCbze1.Text), 2)
'计算利润2
txtLr1.Text = Round(Val(txtJlr1.Text) - Val(txtYj1.Text), 2)
'计算预计提成
txtTc1.Text = Round(Val(txtLr1.Text) * 0.08, 2)
End Sub

Private Sub cmdKuAdd_Click()
On Error Resume Next
form2Htp.adoKu.Recordset.MovePrevious
form2Htp.adoKu.Recordset.AddNew "htBh", form2Htp.adoSale.Recordset.Fields("htBh").Value
form2Htp.adoKu.Recordset.Update "hpBm", form2Htp.adoSale.Recordset.Fields("hpBm").Value
form2Htp.adoKu.Recordset.Update "ljMc", form2Htp.adoSale.Recordset.Fields("ljMc").Value
form2Htp.adoKu.Recordset.Update "phBiao", form2Htp.adoSale.Recordset.Fields("phBiao").Value
form2Htp.adoKu.Recordset.Update "ljBh", form2Htp.adoSale.Recordset.Fields("ljBh").Value
form2Htp.adoKu.Recordset.Update "hpLb", form2Htp.adoSale.Recordset.Fields("hpLb").Value
form2Htp.adoKu.Recordset.Update "jlDw", form2Htp.adoSale.Recordset.Fields("jlDw").Value
form2Htp.adoKu.Recordset.Update "khMc", form2Htp.adoSale.Recordset.Fields("khMc").Value
Set dtgKu.DataSource = form2Htp.adoKu
'cmdKuAdd.Enabled = False
End Sub

Private Sub cmdKuDel_Click()
On Error Resume Next
adoSale.Recordset.MoveFirst
Do While Not adoSale.Recordset.EOF
If adoSale.Recordset.Fields("ljMc").Value & Ko = adoKu.Recordset.Fields("ljMc").Value & Ko And _
       adoSale.Recordset.Fields("phBiao").Value & Ko = adoKu.Recordset.Fields("phBiao").Value & Ko And _
       adoSale.Recordset.Fields("ljBh").Value & Ko = adoKu.Recordset.Fields("ljBh").Value & Ko Then
adoSale.Recordset.Update "xGSl", adoSale.Recordset.Fields("ljSl").Value
adoSale.Recordset.Update "xGSlD", adoSale.Recordset.Fields("ljSl").Value
adoSale.Recordset.Fields("YJje").Value = adoSale.Recordset.Fields("YJdj").Value * adoSale.Recordset.Fields("xGSl").Value
Exit Do
End If

'MsgBox "xGSl=" & adoSale.Recordset.Fields("Xgsl").Value
adoSale.Recordset.MoveNext
Loop


adoKu.Recordset.Delete adAffectCurrent
'Set dtgKu.DataSource = adoKu
cmdKuAdd.Enabled = True
End Sub

Private Sub cmdPrint_Click()
Call modHt.gxQing
htgX.Show
Call modHt.gxBound
htgX.cmdSave.Enabled = False
form2Htp.Visible = False

'设置权限
htgX.cmdMod.Enabled = False
'If mod1.ZW = "维保销售" Or mod1.ZW = "产品销售" Or mod1.ZW = "空调销售" Or mod1.ZW = "销售助理" Then
'    If form2Htp.chkA.Value = 0 Then
'    htgX.cmdMod.Enabled = True
'    End If
'htgX.cmdPrint.Enabled = True
'
'ElseIf mod1.CJQ = True Or mod1.WJQ = True Then
'        If form2Htp.chkB.Value = 0 Or (form2Htp.txtYwy.Text = frmLogin.Combo1.Text And form2Htp.chkE.Value = 0) Then
'            htgX.cmdMod.Enabled = True
'        End If
'      If form2Htp.txtYwy.Text = frmLogin.Combo1.Text And form2Htp.chkE.Value = 0 Then
'        htgX.cmdMod.Enabled = True
'      End If
'ElseIf mod1.KXB = True And mod1.DName <> "倪旭" Then  '销售经理
'    If form2Htp.chkB.Caption = "" And form2Htp.chkC.Caption <> "" Or (form2Htp.chkB.Caption <> "" And form2Htp.chkD.Caption = "") Then
'    htgX.cmdMod.Enabled = True
'    End If
'  If form2Htp.txtYwy.Text = frmLogin.Combo1.Text And form2Htp.chkB.Value = 0 Then
'  htgX.cmdMod.Enabled = True
'  End If
'ElseIf frmLogin.Combo1.Text = "于丽" Or frmLogin.Combo1.Text = "张春华" Or mod1.DName = "倪薇" Then
'    If (form2Htp.chkE.Caption <> "" And form2Htp.chkC.Caption = "") Or (form2Htp.chkC.Caption <> "" And form2Htp.chkB.Caption = "") Then
'    htgX.cmdMod.Enabled = True
'    End If
'
'ElseIf frmLogin.Combo1.Text = "宋晓炯" Or mod1.DName = "倪旭" Then
'htgX.cmdMod.Enabled = True
'ElseIf mod1.ZW = "系统管理员" Then
'htgX.cmdMod.Enabled = True
'Else
'
'End If

'    If (Val(txtHtze.Text) < 10000 And (chkB.Caption <> "" Or chkD.Caption <> "")) Or (Val(txtHtze.Text) >= 10000 And chkD.Caption <> "") Then
'    mod1.kePrint = True
'    Else
'    mod1.kePrint = False
'    End If
    
If wbDN.Visible = True Then '如果为在客户资料中打开合同,则不能修改
    htgX.cmdMod.Enabled = False
    htgX.cmdSave.Enabled = False
    htgX.cmdPrint.Enabled = False
End If

End Sub

Private Sub cmdXAdd_Click()
On Error Resume Next
adoSale.Recordset.AddNew "htBh", form2Htp.txtHtbh.Text
'adoSale.Recordset.AddNew "htF", 0
'adoSale.Recordset.AddNew "delF", 1
Set dtgSale.DataSource = adoSale
End Sub

Private Sub cmdXDel_Click()
On Error Resume Next
adoSale.Recordset.Delete adAffectCurrent
'adoSale.Recordset.UpdateBatch
txtLj.Top = dtgSale.RowHeight * dtgSale.Bookmark - 50
End Sub

Private Sub cmdXMod1_Click()
cmdXMod1.Enabled = False
cmdXAdd.Enabled = True
cmdXDel.Enabled = True
dtgSale.AllowUpdate = True

End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdZhg_Click()
On Error Resume Next
Dim TQ As Double
Dim KQ As Double
TQ = 0
KQ = 0
adoSale.Recordset.MoveFirst
Do While Not adoSale.Recordset.EOF
TQ = TQ + adoSale.Recordset.Fields("ZjJe").Value
adoSale.Recordset.MoveNext
Loop
adoKu.Recordset.MoveFirst
Do While Not adoKu.Recordset.EOF
KQ = KQ + adoKu.Recordset.Fields("je").Value
adoKu.Recordset.MoveNext
Loop

txtClcb2.Text = TQ + KQ

''计算预计成本总额
'txtCbze1.Text = Round(Val(txtClcb1.Text) + Val(txtQt1.Text) + Val(txtYf1.Text) + Val(txtZXF1.Text), 2)
''计算预计利润1
'txtJlr1.Text = Round(Val(txtHtze.Text) - Val(txtCbze1.Text), 2)
''计算利润2
'txtLr1.Text = Round(Val(txtJlr1.Text) - Val(txtYj1.Text), 2)
''计算预计提成
'txtTc1.Text = Round(Val(txtLr1.Text) * 0.08, 2)
End Sub

Private Sub Command2_Click()
txtHtze.Text = Val(txtClf.Text) + Val(txtRgf.Text) + Val(txtAz1.Text) + _
Val(txtAz2.Text) + Val(txtAz3.Text) + Val(txtAz4.Text) + Val(txtAz5.Text)
End Sub

Private Sub Command3_Click()
Dim tt As String
On Error Resume Next

Call mod1.DelDKZ ' '退出表单时删除打开记录,以让别人能打开此单据

Call modHt.HtF '检查htping,htping1,yiFk 三表间的htF一致性

'Select Case mod1.anButton
'Case 1
'frmZu.Enabled = True
'frmZu.frm5.Enabled = True
'Case 2
'htBrow.Enabled = True
'
'Case 3
'frmZu.Enabled = True
'frmZu.frm5.Enabled = True
'End Select
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
ElseIf htBrow.Visible = True Then
    htBrow.Enabled = True
    htBrow.ZOrder 0
ElseIf htBrowG.Visible = True Then
    htBrowG.Enabled = True
    htBrowG.ZOrder 0
End If
form2Htp.Visible = False
frmFuK.Visible = False
End Sub

Private Sub Command4_Click()

End Sub

Private Sub DataGrid4_Click()

End Sub

Private Sub DT1_Change()
'txtHtbh.Text = ""
End Sub

Private Sub dt1_CloseUp()
txtHtdate.Text = Format(dt1.Value, "YYYY年M月D日")
End Sub

Private Sub dt2_CloseUp()
txtDddate.Text = Format(dt2.Value, "Long Date")
End Sub

Private Sub dt3_CloseUp()
txtHtqy.Text = Format(dt3.Value, "long Date")
End Sub

Private Sub dt4_CloseUp()
txtHtqy1.Text = Format(dt4.Value, "long Date")
End Sub

Private Sub dtgCG_Click()
On Error Resume Next
cmdDDH.Caption = frmAdo.adoTmp.Recordset.Fields("DDH").Value
End Sub

Private Sub dtgGzb_Click()
On Error Resume Next
cmdGzd.Caption = form2Htp.adoGzb.Recordset.Fields(17).Value
End Sub

Private Sub dtgGzb_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
cmdGzd.Caption = form2Htp.adoGzb.Recordset.Fields(17).Value
End Sub

Private Sub dtgKc_DblClick()
On Error Resume Next
dtgSale.Columns(1).Value = dtgKC.Columns(0).Value
'dtgLj.Columns(3).Value = ""
txtLj.Visible = False
'txtLj.Text = ""
frmLj.Visible = False
End Sub

Private Sub dtgKu_AfterColUpdate(ByVal ColIndex As Integer)
If ColIndex = 5 Or ColIndex = 6 Then
adoSale.Recordset.MoveFirst
Do While Not adoSale.Recordset.EOF
If adoSale.Recordset.Fields("ljMc").Value & Ko = adoKu.Recordset.Fields("ljMc").Value & Ko And _
       adoSale.Recordset.Fields("phBiao").Value & Ko = adoKu.Recordset.Fields("phBiao").Value & Ko And _
       adoSale.Recordset.Fields("ljBh").Value & Ko = adoKu.Recordset.Fields("ljBh").Value & Ko Then
adoSale.Recordset.Update "xGSl", adoSale.Recordset.Fields("ljSl").Value - adoKu.Recordset.Fields("ljSl").Value
adoSale.Recordset.Update "xGSlD", adoSale.Recordset.Fields("ljSl").Value - adoKu.Recordset.Fields("ljSl").Value
adoSale.Recordset.Fields("YJje").Value = adoSale.Recordset.Fields("YJdj").Value * adoSale.Recordset.Fields("xGSl").Value

Exit Do
End If

'MsgBox "xGSl=" & adoSale.Recordset.Fields("Xgsl").Value
adoSale.Recordset.MoveNext
Loop

adoKu.Recordset.Fields("je").Value = adoKu.Recordset.Fields("dj").Value * adoKu.Recordset.Fields("ljSl").Value
Set dtgYJ.DataSource = adoSale
Set dtgZj.DataSource = adoSale
End If

End Sub

Private Sub dtgKu_Click()
'Dim Ko As String
'On Error Resume Next
'Ko = ""
''点击库存表时，相应对应销售表，以方便相应产品数量的变化
'adoSale.Recordset.MoveFirst
'Do While Not adoSale.Recordset.EOF
'    If adoSale.Recordset.Fields("ljMc").Value & Ko = adoKu.Recordset.Fields("ljMc").Value & Ko And _
'        adoSale.Recordset.Fields("phBiao").Value & Ko = adoKu.Recordset.Fields("phBiao").Value & Ko And _
'        adoSale.Recordset.Fields("ljBh").Value & Ko = adoKu.Recordset.Fields("ljBh").Value & Ko Then
'        Exit Do
'    End If
'    adoSale.Recordset.MoveNext
'Loop
End Sub

Private Sub dtgKu_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'Dim Ko As String
'On Error Resume Next
'Ko = ""
''点击库存表时，相应对应销售表，以方便相应产品数量的变化
'adoSale.Recordset.MoveFirst
'Do While Not adoSale.Recordset.EOF
'    If adoSale.Recordset.Fields("ljMc").Value & Ko = adoKu.Recordset.Fields("ljMc").Value & Ko And _
'        adoSale.Recordset.Fields("phBiao").Value & Ko = adoKu.Recordset.Fields("phBiao").Value & Ko And _
'        adoSale.Recordset.Fields("ljBh").Value & Ko = adoKu.Recordset.Fields("ljBh").Value & Ko Then
'        Exit Do
'    End If
'    adoSale.Recordset.MoveNext
'Loop
End Sub

Private Sub dtgLj_Click()
On Error Resume Next
form2Htp.cmdGzd.Caption = form2Htp.adoLj.Recordset.Fields(6).Value
End Sub

Private Sub dtgLj_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
form2Htp.cmdGzd.Caption = form2Htp.adoLj.Recordset.Fields(6).Value
End Sub

Private Sub dtgYj_AfterColUpdate(ByVal ColIndex As Integer)
adoSale.Recordset.Fields("YJje").Value = adoSale.Recordset.Fields("YJdj").Value * adoSale.Recordset.Fields("xGSld").Value

End Sub

Private Sub dtgYj_Click()
On Error Resume Next
Dim Ko As String
''点击销售表时，相应对应库存表，以判断能否在库存表中添加。（如果库存表相应记录已存在，则不能添加）
'Ko = ""
'cmdKuAdd.Enabled = True
'adoKu.Recordset.MoveFirst
'Do While Not adoKu.Recordset.EOF
'    If adoSale.Recordset.Fields("ljMc").Value & Ko = adoKu.Recordset.Fields("ljMc").Value & Ko And _
'        adoSale.Recordset.Fields("phBiao").Value & Ko = adoKu.Recordset.Fields("phBiao").Value & Ko And _
'        adoSale.Recordset.Fields("ljBh").Value & Ko = adoKu.Recordset.Fields("ljBh").Value & Ko Then
'        cmdKuAdd.Enabled = False
'        Exit Do
'    End If
'    adoKu.Recordset.MoveNext
'Loop
End Sub


Private Sub dtgYj_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'On Error Resume Next
'Dim Ko As String
''点击销售表时，相应对应库存表，以判断能否在库存表中添加。（如果库存表相应记录已存在，则不能添加）
'Ko = ""
'cmdKuAdd.Enabled = True
'adoKu.Recordset.MoveFirst
'Do While Not adoKu.Recordset.EOF
'    If adoSale.Recordset.Fields("ljMc").Value & Ko = adoKu.Recordset.Fields("ljMc").Value & Ko And _
'        adoSale.Recordset.Fields("phBiao").Value & Ko = adoKu.Recordset.Fields("phBiao").Value & Ko And _
'        adoSale.Recordset.Fields("ljBh").Value & Ko = adoKu.Recordset.Fields("ljBh").Value & Ko Then
'        cmdKuAdd.Enabled = False
'        Exit Do
'    End If
'    adoKu.Recordset.MoveNext
'Loop
End Sub

Private Sub dtgZj_AfterColUpdate(ByVal ColIndex As Integer)
adoSale.Recordset.Fields("ZJje").Value = adoSale.Recordset.Fields("ZJdj").Value * adoSale.Recordset.Fields("xGSld").Value
End Sub

Private Sub dtgZj_Click()
Dim Ko As String
On Error Resume Next
Ko = ""
''点击销售表时，相应对应库存表，以判断能否在库存表中添加。（如果库存表相应记录已存在，则不能添加）
'cmdKuAdd.Enabled = True
'adoKu.Recordset.MoveFirst
'Do While Not adoKu.Recordset.EOF
'    If adoSale.Recordset.Fields("ljMc").Value & Ko = adoKu.Recordset.Fields("ljMc").Value & Ko And _
'        adoSale.Recordset.Fields("phBiao").Value & Ko = adoKu.Recordset.Fields("phBiao").Value & Ko And _
'        adoSale.Recordset.Fields("ljBh").Value & Ko = adoKu.Recordset.Fields("ljBh").Value & Ko Then
'        cmdKuAdd.Enabled = False
'        Exit Do
'    End If
'    adoKu.Recordset.MoveNext
'Loop
End Sub

Private Sub dtgZj_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'Dim Ko As String
'On Error Resume Next
'Ko = ""
''点击销售表时，相应对应库存表，以判断能否在库存表中添加。（如果库存表相应记录已存在，则不能添加）
'cmdKuAdd.Enabled = True
'adoKu.Recordset.MoveFirst
'Do While Not adoKu.Recordset.EOF
'    If adoSale.Recordset.Fields("ljMc").Value & Ko = adoKu.Recordset.Fields("ljMc").Value & Ko And _
'        adoSale.Recordset.Fields("phBiao").Value & Ko = adoKu.Recordset.Fields("phBiao").Value & Ko And _
'        adoSale.Recordset.Fields("ljBh").Value & Ko = adoKu.Recordset.Fields("ljBh").Value & Ko Then
'        cmdKuAdd.Enabled = False
'        Exit Do
'    End If
'    adoKu.Recordset.MoveNext
'Loop
End Sub

Private Sub dtPJhq_CloseUp()
txtJhq.Text = dtPJhq.Value
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'''''If Shift = 2 And KeyCode = 76 Then
'''''    form2Htp.txtYj1.Locked = True
'''''    form2Htp.txtYj2.Locked = True
'''''
'''''    'If (mod1.Kyj = True And lblBM.Caption = mod1.BMN) Or (mod1.DName = "张春华" And (optW.Value = True Or optZ.Visible = True)) Or mod1.ZW = "总经理" Or mod1.ZW = "副总经理" Then
'''''      If mod1.Kyj = True Then
'''''      If form2Htp.txtYj1.Visible = False Then
'''''            form2Htp.txtYj1.Visible = True
'''''            form2Htp.txtYj2.Visible = True
'''''            form2Htp.txtLr1.Visible = True
'''''            form2Htp.txtLr2.Visible = True
''''''            form2Htp.txtTc1.Visible = True
''''''            form2Htp.txtTc2.Visible = True
'''''            form2Htp.lblTcBe.Visible = True
'''''            form2Htp.txtTcBe.Visible = True
'''''            form2Htp.UpDa.Visible = True
'''''            form2Htp.lblYj.Visible = True
'''''            form2Htp.lblLr2.Visible = True
'''''            form2Htp.lblTcBe.Visible = True
'''''            form2Htp.txtTcBe.Visible = True
'''''            form2Htp.UpDa.Visible = True
''''''            If mod1.KY2 = True And optW.Value = False Then '小张只能修改合同末完成的实际佣金
''''''                txtYj2.Locked = False
''''''            End If
''''''            If mod1.KY1 = True Then '销售经理在老板签字后,就不能修改预计佣金
''''''                If chkD.Caption = "" Or mod1.ZW = "总经理" Or mod1.ZW = "副总经理" Then
''''''                    txtYj1.Locked = False
''''''                End If
''''''            End If
'''''    Else
'''''            form2Htp.txtYj1.Visible = False
'''''            form2Htp.txtYj2.Visible = False
'''''            form2Htp.txtLr1.Visible = False
'''''            form2Htp.txtLr2.Visible = False
'''''            form2Htp.lblTcBe.Visible = False
'''''            form2Htp.txtTcBe.Visible = False
'''''            form2Htp.UpDa.Visible = False
'''''            form2Htp.lblYj.Visible = False
'''''            form2Htp.lblLr2.Visible = False
'''''            form2Htp.lblTc.Visible = False
'''''    End If
'''''    End If
'''''
'''''
'''''
'''''
''''' End If
End Sub

Private Sub Form_Load()
Dim tt As String
Dim oo As Integer
On Error Resume Next
form2Htp.Width = 9915
form2Htp.Height = 9495
form2Htp.Top = 0
form2Htp.Left = 2000
''设置区域下拉框
'If form2Htp.comQy.ListCount > 0 Then
'    For oo = comQy.ListCount - 1 To 0 Step -1
'    comQy.RemoveItem oo
'    Next
'End If
'
'    Form2.adoHi.Recordset.Close
'    tt = "select * from yzQy"
'    Form2.adoHi.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'    Form2.adoHi.Recordset.MoveFirst
'    oo = 0
'    Do While Not Form2.adoHi.Recordset.EOF
'    comQy.AddItem Form2.adoHi.Recordset.Fields(0).Value, oo
'    oo = oo + 1
'    Form2.adoHi.Recordset.MoveNext
'    Loop
'
'    Form2.adoHi.Recordset.Close
'    DT1.Value = mod1.DQda
'End If
End Sub






Private Sub Form_Unload(Cancel As Integer)
Dim tt As String
On Error Resume Next
If MDI.Cq = False Then
Call mod1.DelDKZ ' '退出表单时删除打开记录,以让别人能打开此单据

If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
ElseIf htBrow.Visible = True Then
    htBrow.Enabled = True
    htBrow.ZOrder 0
ElseIf htBrowG.Visible = True Then
    htBrowG.Enabled = True
    htBrowG.ZOrder 0
End If
form2Htp.Visible = False
frmFuK.Visible = False


    Cancel = True
End If
End Sub



Private Sub Frame6_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub opTa_Click(Index As Integer)
lblHtxz.Caption = optA(Index).Caption
'If lblHtxz.Caption <> "" Then
'txtHtbh.Text = ""
'End If
End Sub

Private Sub optG_Click()
cmdSave.Enabled = True
End Sub

Private Sub optZ_GotFocus()
cmdSave.Enabled = True
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
'If SSTab1.Tab = 2 And frmLogin.Combo1.Text = "张春华" Then
'    dtgZj.AllowUpdate = True
'End If
End Sub

Private Sub tabHt_Click(PreviousTab As Integer)
Select Case tabHt.Tab

Case 0

Case 1
    SSTab1.Tab = 0
End Select


cmdPrint.SetFocus

End Sub

Private Sub tabHt_GotFocus()
cmdPrint.SetFocus
End Sub


Private Sub txtCbze1_LostFocus()
txtLr1.Text = Val(txtHtze.Text) - Val(txtCbze1.Text)

End Sub





Private Sub txtClcb1_DblClick()
Dim ii As Integer
ii = MsgBox("是否要按比例分配预计成本？", vbInformation + vbYesNo + vbDefaultButton2, "???")
If ii = vbYes Then
'计算预计材料成本表中的明细,按照比例分配
Bl = Val(txtClcb1.Text) / Val(txtHtze.Text)
    If adoSale.Recordset.RecordCount > 0 Then
        adoSale.Recordset.MoveFirst
        Do While Not adoSale.Recordset.EOF
            adoSale.Recordset.Fields("YJdj").Value = Round(adoSale.Recordset.Fields("dj").Value * Bl, 2)
            adoSale.Recordset.Fields("YJje").Value = Round(adoSale.Recordset.Fields("yjdj").Value * adoSale.Recordset.Fields("ljSl").Value, 2)
            adoSale.Recordset.MoveNext
        Loop
        Set dtgYJ.DataSource = adoSale
    End If
End If

'计算预计成本总额
txtCbze1.Text = Round(Val(txtClcb1.Text) + Val(txtQt1.Text) + Val(txtYf1.Text) + Val(txtZXF1.Text), 2)
'计算预计利润1
txtJlr1.Text = Round(Val(txtHtze.Text) - Val(txtCbze1.Text), 2)
'计算利润2
txtLr1.Text = Round(Val(txtJlr1.Text) - Val(txtYj1.Text), 2)
'计算预计提成
txtTc1.Text = Round(Val(txtLr1.Text) * 0.08, 2)
End Sub


Private Sub txtClcb1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
frmDgd.Visible = True
End If

End Sub






Private Sub txtCLF_DblClick()
frmClf.Show
End Sub






Private Sub txtFBQT_Click()
tabHt.Tab = 3
End Sub

Private Sub txtFj_Click()
tabHt.Tab = 4
End Sub

Private Sub txtHtbh_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
txtHtdate.SetFocus
End If
End Sub









Private Sub txtKhmc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
txtHtbh.SetFocus
End If
End Sub



Private Sub txtLj_Change()
'Dim LT As String
'On Error Resume Next
''If ColIndex = 1 Then
''If txtLj.Text <> "" Then
'If Not (mod1.anButton = 2 And form2Htp.cmdSave.Enabled = False) And txtLj.Visible = True Then
'    LT = "Select pmGg,hpBm from kc where pmGg like '%" & txtLj.Text & "%'"
'    form2Htp.adoKc.Recordset.Close
'    form2Htp.adoKc.Recordset.Open LT, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'    Set form2Htp.dtgKc.DataSource = form2Htp.adoKc
'        If form2Htp.adoKc.Recordset.RecordCount > 1 Then
'        frmLj.Top = txtLj.Top + txtLj.Height
'        frmLj.Visible = True
'        dtgKc.Height = dtgKc.RowHeight * adoKc.Recordset.RecordCount
'        frmLj.Height = dtgKc.Height
'        Else
'        frmLj.Visible = False
'        End If
'    'End If
'    dtgSale.Columns(1).Value = txtLj.Text
'    'End If
'End If
End Sub

Private Sub txtLj_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'Dim LT As String
'On Error Resume Next
''If ColIndex = 1 Then
'
'LT = "Select pmGg,hpBm from kc where pmGg like '%" & txtLj.Text & "%'"
'form2Htp.adoKc.Recordset.Close
'form2Htp.adoKc.Recordset.Open LT, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'Set form2Htp.dtgKc.DataSource = form2Htp.adoKc
' If form2Htp.adoKc.Recordset.RecordCount > 0 Then
' frmLj.Top = txtLj.Top + txtLj.Height
' frmLj.Visible = True
' End If
''End If
'End If

End Sub



















Private Sub txtQt1_Change()
''计算预计成本总额
'txtCbze1.Text = Round(Val(txtClcb1.Text) + Val(txtQt1.Text) + Val(txtYf1.Text) + Val(txtZXF1.Text), 2)
''计算预计利润1
'txtJlr1.Text = Round(Val(txtHtze.Text) - Val(txtCbze1.Text), 2)
''计算利润2
'txtLr1.Text = Round(Val(txtJlr1.Text) - Val(txtYj1.Text), 2)
''计算预计提成
'txtTc1.Text = Round(Val(txtLr1.Text) * 0.08, 2)
End Sub

Private Sub txtQt2_DblClick()
Dim tt As String
On Error Resume Next
frmXmFy.Show
tt = "Select * from fyTg where htBh='" & txtHtbh.Text & "' and BxF=1"
frmXmFy.adoFy.Recordset.Close
frmXmFy.adoFy.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmXmFy.dtgFy.DataSource = frmXmFy.adoFy
End Sub

Private Sub txtYf1_Change()
''计算预计成本总额
'txtCbze1.Text = Round(Val(txtClcb1.Text) + Val(txtQt1.Text) + Val(txtYf1.Text) + Val(txtZXF1.Text), 2)
''计算预计利润1
'txtJlr1.Text = Round(Val(txtHtze.Text) - Val(txtCbze1.Text), 2)
''计算利润2
'txtLr1.Text = Round(Val(txtJlr1.Text) - Val(txtYj1.Text), 2)
''计算预计提成
'txtTc1.Text = Round(Val(txtLr1.Text) * 0.08, 2)
End Sub

Private Sub txtYj1_Change()
''计算预计成本总额
'txtCbze1.Text = Round(Val(txtClcb1.Text) + Val(txtQt1.Text) + Val(txtYf1.Text) + Val(txtZXF1.Text), 2)
''计算预计利润1
'txtJlr1.Text = Round(Val(txtHtze.Text) - Val(txtCbze1.Text), 2)
''计算利润2
'txtLr1.Text = Round(Val(txtJlr1.Text) - Val(txtYj1.Text), 2)
''计算预计提成
'txtTc1.Text = Round(Val(txtLr1.Text) * 0.08, 2)
End Sub

Private Sub txtYj1_DblClick()

If mod1.DName = "倪旭" Then
'    If Val(txtYj2.Text) > 0 And frmYj.adoYj.Recordset.RecordCount = 0 Then
'        MsgBox "新旧版交替导致数据有误,请与马晓聪联系!"
'        Exit Sub
'    End If
    frmYj.cmdAdd.Visible = True
    frmYj.cmdDel.Visible = True
    frmYj.cmdSave.Visible = True
Else
    frmYj.cmdAdd.Visible = False
    frmYj.cmdAdd.Visible = False
    frmYj.cmdSave.Visible = False
End If
If mod1.DName = "马晓聪" Then
    frmYj.cmdAdd.Visible = True
    frmYj.cmdDel.Visible = True
    frmYj.cmdSave.Visible = True
End If
frmYj.Show
frmYj.lblHtbh.Caption = txtHtbh.Text
frmYj.lblKhmc.Caption = txtKhmc.Text
End Sub


Private Sub txtYj2_DblClick()
'frmYj.Show
'frmYj.lblHtbh.Caption = txtHtbh.Text
'frmYj.lblKhmc.Caption = txtKhmc.Text
End Sub


Private Sub txtZXF1_Change()
''计算预计成本总额
'txtCbze1.Text = Round(Val(txtClcb1.Text) + Val(txtQt1.Text) + Val(txtYf1.Text) + Val(txtZXF1.Text), 2)
''计算预计利润1
'txtJlr1.Text = Round(Val(txtHtze.Text) - Val(txtCbze1.Text), 2)
''计算利润2
'txtLr1.Text = Round(Val(txtJlr1.Text) - Val(txtYj1.Text), 2)
''计算预计提成
'txtTc1.Text = Round(Val(txtLr1.Text) * 0.08, 2)
End Sub


Private Sub UpDown1_Change()

End Sub

Private Sub UpDa_DownClick()
txtTcBe.Text = Val(txtTcBe.Text) + 1
If Val(txtTcBe.Text) = 13 Then
    txtTcBe.Text = 12
End If

End Sub


Private Sub UpDa_UpClick()
txtTcBe.Text = Val(txtTcBe.Text) - 1
If Val(txtTcBe.Text) = -1 Then
    txtTcBe.Text = 0
End If
End Sub


