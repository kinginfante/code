VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form wbMx 
   Caption         =   "维保合同评审单名细"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   4620
   ScaleWidth      =   11865
   Begin TabDlg.SSTab SSTab1 
      Height          =   4635
      Left            =   -150
      TabIndex        =   0
      Top             =   60
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   8176
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      Enabled         =   0   'False
      TabCaption(0)   =   "人工费明细"
      TabPicture(0)   =   "wbMx.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label14"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label18"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label17"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label16"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label15"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label13"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label12"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label11"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label20"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label21"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label22"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label23"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label24"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label26"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label27"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label28"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label29"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label42"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblHG"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label44"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblHG1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label25"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtDdj"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtDgT"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtDxG"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtXdj"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtXgT"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtXxG"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtJdj"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtJgT"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtJxG"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtGdj"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtGgT"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtGxG"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtDgT1"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtDxG1"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtXgT1"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtXxG1"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtJgT1"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtJxG1"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtGgT1"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtGxG1"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "cmdCou"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "dtgGzb"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "cmdGzd"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).ControlCount=   45
      TabCaption(1)   =   "差旅费明细"
      TabPicture(1)   =   "wbMx.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label10"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label9"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label8"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label7"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label6"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label4"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label3"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label1"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label30"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label31"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label32"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label33"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label34"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label35"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label36"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label37"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label40"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label41"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "lblCF"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label45"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "lblCF1"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Label47"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "txtDDXG"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "txtDDCou"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "txtDDJE"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "txtCXG"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "txtCCou"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "txtCJE"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "txtZXG"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "txtZCou"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "txtZJE"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "txtQCXG"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "txtQCCou"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "txtQCJE"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "txtHCXG"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "txtHCCou"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "txtHCJE"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "txtJPXG"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "txtJPCou"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "txtJPJE"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "txtJPXG1"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "txtHCXG1"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "txtQCXG1"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "txtZXG1"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "txtCXG1"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "txtDDXG1"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "cmdCCou"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "dtgCl"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).ControlCount=   50
      TabCaption(2)   =   "材料费明细"
      TabPicture(2)   =   "wbMx.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label19"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label39"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblChg"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "DataGrid1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "dtgSale"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "adoRGF"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtLjmc"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "dtgLjmc"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "adoLjmc"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "cmdMod1"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cmdDel"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "cmdAdd"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Text1"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "cmdJi"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "收款明细"
      TabPicture(3)   =   "wbMx.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label38"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblHtZe"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label43"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "dtgYf"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "dtgFk"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "cmdJmod"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "cmdJdel"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "cmdJadd"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "txtFkBz"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).ControlCount=   9
      Begin VB.TextBox txtFkBz 
         Height          =   1155
         Left            =   3780
         TabIndex        =   115
         Top             =   3420
         Width           =   2895
      End
      Begin VB.CommandButton cmdJi 
         Caption         =   "计算"
         Height          =   315
         Left            =   -64080
         TabIndex        =   114
         Top             =   1800
         Width           =   555
      End
      Begin VB.CommandButton cmdGzd 
         Height          =   345
         Left            =   -65730
         TabIndex        =   108
         Top             =   4260
         Width           =   1845
      End
      Begin VB.CommandButton cmdJadd 
         Caption         =   "添加"
         Height          =   375
         Left            =   8040
         TabIndex        =   105
         Top             =   600
         Width           =   555
      End
      Begin VB.CommandButton cmdJdel 
         Caption         =   "删除"
         Height          =   375
         Left            =   8040
         TabIndex        =   104
         Top             =   990
         Width           =   555
      End
      Begin VB.CommandButton cmdJmod 
         Caption         =   "修改"
         Height          =   375
         Left            =   8040
         TabIndex        =   103
         Top             =   1350
         Width           =   555
      End
      Begin MSDataGridLib.DataGrid dtgGzb 
         Height          =   3675
         Left            =   -67470
         TabIndex        =   102
         Top             =   330
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   6482
         _Version        =   393216
         BackColor       =   14680045
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
         Caption         =   "出工记录"
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "max(gzb.rq)"
            Caption         =   "日期"
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
            DataField       =   "max(gzb.wxWorker)"
            Caption         =   "维修工"
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
            DataField       =   "sum(workXX.wTime)"
            Caption         =   "工时"
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
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1305.071
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dtgCl 
         Height          =   2475
         Left            =   -66690
         TabIndex        =   101
         Top             =   870
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   4366
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   13697002
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
         Caption         =   "    详情（报销单据）"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   "日期"
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
            DataField       =   ""
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
         SplitCount      =   1
         BeginProperty Split0 
            RecordSelectors =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdCCou 
         Caption         =   "计算"
         Height          =   345
         Left            =   -73920
         TabIndex        =   99
         Top             =   3660
         Width           =   975
      End
      Begin VB.CommandButton cmdCou 
         Caption         =   "计算"
         Height          =   345
         Left            =   -73830
         TabIndex        =   94
         Top             =   3270
         Width           =   1065
      End
      Begin VB.TextBox txtDDXG1 
         Height          =   345
         Left            =   -68130
         TabIndex        =   79
         Top             =   3000
         Width           =   1365
      End
      Begin VB.TextBox txtCXG1 
         Height          =   345
         Left            =   -68130
         TabIndex        =   78
         Top             =   2670
         Width           =   1365
      End
      Begin VB.TextBox txtZXG1 
         Height          =   345
         Left            =   -68130
         TabIndex        =   77
         Top             =   2310
         Width           =   1365
      End
      Begin VB.TextBox txtQCXG1 
         Height          =   345
         Left            =   -68130
         TabIndex        =   76
         Top             =   1980
         Width           =   1365
      End
      Begin VB.TextBox txtHCXG1 
         Height          =   345
         Left            =   -68130
         TabIndex        =   75
         Top             =   1620
         Width           =   1365
      End
      Begin VB.TextBox txtJPXG1 
         Height          =   345
         Left            =   -68130
         TabIndex        =   74
         Top             =   1290
         Width           =   1365
      End
      Begin VB.TextBox txtGxG1 
         Height          =   300
         Left            =   -68520
         TabIndex        =   64
         Top             =   2040
         Width           =   930
      End
      Begin VB.TextBox txtGgT1 
         Height          =   300
         Left            =   -69450
         TabIndex        =   63
         Top             =   2040
         Width           =   930
      End
      Begin VB.TextBox txtJxG1 
         Height          =   300
         Left            =   -68520
         TabIndex        =   62
         Top             =   1740
         Width           =   930
      End
      Begin VB.TextBox txtJgT1 
         Height          =   300
         Left            =   -69450
         TabIndex        =   61
         Top             =   1740
         Width           =   930
      End
      Begin VB.TextBox txtXxG1 
         Height          =   300
         Left            =   -68520
         TabIndex        =   60
         Top             =   1440
         Width           =   930
      End
      Begin VB.TextBox txtXgT1 
         Height          =   300
         Left            =   -69450
         TabIndex        =   59
         Top             =   1440
         Width           =   930
      End
      Begin VB.TextBox txtDxG1 
         Height          =   300
         Left            =   -68520
         TabIndex        =   58
         Top             =   2340
         Width           =   930
      End
      Begin VB.TextBox txtDgT1 
         Height          =   300
         Left            =   -69450
         TabIndex        =   57
         Top             =   2340
         Width           =   930
      End
      Begin VB.TextBox txtGxG 
         Height          =   300
         Left            =   -71760
         TabIndex        =   48
         Top             =   2040
         Width           =   960
      End
      Begin VB.TextBox txtGgT 
         Height          =   300
         Left            =   -72720
         TabIndex        =   47
         Top             =   2040
         Width           =   960
      End
      Begin VB.TextBox txtGdj 
         Height          =   300
         Left            =   -73680
         TabIndex        =   46
         Top             =   2040
         Width           =   960
      End
      Begin VB.TextBox txtJxG 
         Height          =   300
         Left            =   -71760
         TabIndex        =   45
         Top             =   1740
         Width           =   960
      End
      Begin VB.TextBox txtJgT 
         Height          =   300
         Left            =   -72720
         TabIndex        =   44
         Top             =   1740
         Width           =   960
      End
      Begin VB.TextBox txtJdj 
         Height          =   300
         Left            =   -73680
         TabIndex        =   43
         Top             =   1740
         Width           =   960
      End
      Begin VB.TextBox txtXxG 
         Height          =   300
         Left            =   -71760
         TabIndex        =   42
         Top             =   1440
         Width           =   960
      End
      Begin VB.TextBox txtXgT 
         Height          =   300
         Left            =   -72720
         TabIndex        =   41
         Top             =   1440
         Width           =   960
      End
      Begin VB.TextBox txtXdj 
         Height          =   300
         Left            =   -73680
         TabIndex        =   40
         Top             =   1440
         Width           =   960
      End
      Begin VB.TextBox txtDxG 
         Height          =   300
         Left            =   -71760
         TabIndex        =   39
         Top             =   2340
         Width           =   960
      End
      Begin VB.TextBox txtDgT 
         Height          =   300
         Left            =   -72720
         TabIndex        =   38
         Top             =   2340
         Width           =   960
      End
      Begin VB.TextBox txtDdj 
         Height          =   300
         Left            =   -73680
         TabIndex        =   37
         Top             =   2340
         Width           =   960
      End
      Begin VB.TextBox Text1 
         DataField       =   "UserId"
         DataSource      =   "adoRGF"
         Height          =   285
         Left            =   -63570
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   930
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "添加"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -64080
         TabIndex        =   35
         Top             =   630
         Width           =   555
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -64080
         TabIndex        =   34
         Top             =   1020
         Width           =   555
      End
      Begin VB.CommandButton cmdMod1 
         Caption         =   "修改"
         Height          =   375
         Left            =   -64080
         TabIndex        =   33
         Top             =   1410
         Width           =   555
      End
      Begin MSAdodcLib.Adodc adoLjmc 
         Height          =   465
         Left            =   -70560
         Top             =   2490
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
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
      Begin MSDataGridLib.DataGrid dtgLjmc 
         Bindings        =   "wbMx.frx":0070
         Height          =   1395
         Left            =   -73170
         TabIndex        =   32
         Top             =   2520
         Visible         =   0   'False
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   2461
         _Version        =   393216
         AllowUpdate     =   -1  'True
         ColumnHeaders   =   0   'False
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
         ColumnCount     =   1
         BeginProperty Column00 
            DataField       =   "pmGg"
            Caption         =   "pmGg"
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
               ColumnWidth     =   2399.811
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtLjmc 
         Height          =   315
         Left            =   -73170
         TabIndex        =   31
         Top             =   2220
         Width           =   2475
      End
      Begin VB.TextBox txtJPJE 
         Height          =   345
         Left            =   -73395
         TabIndex        =   18
         Top             =   1290
         Width           =   1125
      End
      Begin VB.TextBox txtJPCou 
         Height          =   345
         Left            =   -72270
         TabIndex        =   17
         Top             =   1290
         Width           =   1125
      End
      Begin VB.TextBox txtJPXG 
         Height          =   345
         Left            =   -71145
         TabIndex        =   16
         Top             =   1290
         Width           =   1125
      End
      Begin VB.TextBox txtHCJE 
         Height          =   345
         Left            =   -73395
         TabIndex        =   15
         Top             =   1620
         Width           =   1125
      End
      Begin VB.TextBox txtHCCou 
         Height          =   345
         Left            =   -72270
         TabIndex        =   14
         Top             =   1620
         Width           =   1125
      End
      Begin VB.TextBox txtHCXG 
         Height          =   345
         Left            =   -71145
         TabIndex        =   13
         Top             =   1620
         Width           =   1125
      End
      Begin VB.TextBox txtQCJE 
         Height          =   345
         Left            =   -73395
         TabIndex        =   12
         Top             =   1980
         Width           =   1125
      End
      Begin VB.TextBox txtQCCou 
         Height          =   345
         Left            =   -72270
         TabIndex        =   11
         Top             =   1980
         Width           =   1125
      End
      Begin VB.TextBox txtQCXG 
         Height          =   345
         Left            =   -71145
         TabIndex        =   10
         Top             =   1980
         Width           =   1125
      End
      Begin VB.TextBox txtZJE 
         Height          =   345
         Left            =   -73395
         TabIndex        =   9
         Top             =   2310
         Width           =   1125
      End
      Begin VB.TextBox txtZCou 
         Height          =   345
         Left            =   -72270
         TabIndex        =   8
         Top             =   2310
         Width           =   1125
      End
      Begin VB.TextBox txtZXG 
         Height          =   345
         Left            =   -71145
         TabIndex        =   7
         Top             =   2310
         Width           =   1125
      End
      Begin VB.TextBox txtCJE 
         Height          =   345
         Left            =   -73395
         TabIndex        =   6
         Top             =   2670
         Width           =   1125
      End
      Begin VB.TextBox txtCCou 
         Height          =   345
         Left            =   -72270
         TabIndex        =   5
         Top             =   2670
         Width           =   1125
      End
      Begin VB.TextBox txtCXG 
         Height          =   345
         Left            =   -71145
         TabIndex        =   4
         Top             =   2670
         Width           =   1125
      End
      Begin VB.TextBox txtDDJE 
         Height          =   345
         Left            =   -73395
         TabIndex        =   3
         Top             =   3000
         Width           =   1125
      End
      Begin VB.TextBox txtDDCou 
         Height          =   345
         Left            =   -72270
         TabIndex        =   2
         Top             =   3000
         Width           =   1125
      End
      Begin VB.TextBox txtDDXG 
         Height          =   345
         Left            =   -71145
         TabIndex        =   1
         Top             =   3000
         Width           =   1125
      End
      Begin MSAdodcLib.Adodc adoRGF 
         Height          =   330
         Left            =   -64590
         Top             =   150
         Visible         =   0   'False
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
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
      Begin MSDataGridLib.DataGrid dtgSale 
         Bindings        =   "wbMx.frx":0086
         Height          =   1875
         Left            =   -74970
         TabIndex        =   29
         Top             =   300
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   3307
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         Caption         =   "预计"
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "hpBm"
            Caption         =   "材料编码"
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
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1995.024
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
               ColumnWidth     =   794.835
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
         Bindings        =   "wbMx.frx":009B
         Height          =   1965
         Left            =   -74910
         TabIndex        =   100
         Top             =   2430
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   3466
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   14155760
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
         Caption         =   "实际"
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "hpBm"
            Caption         =   "材料编码"
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
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1995.024
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
               ColumnWidth     =   794.835
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
      Begin MSDataGridLib.DataGrid dtgFk 
         Bindings        =   "wbMx.frx":00B1
         Height          =   1995
         Left            =   0
         TabIndex        =   106
         Top             =   300
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   3519
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
         WrapCellPointer =   -1  'True
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
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "yWy"
            Caption         =   "yWy"
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
            DataField       =   "rq"
            Caption         =   "应收日期"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dddddd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "ED"
            Caption         =   "收款额度"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   5
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "yingfJe"
            Caption         =   "应收金额"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """￥""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column04 
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
         BeginProperty Column05 
            DataField       =   "yifJe"
            Caption         =   "yifJe"
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
            DataField       =   "ZT"
            Caption         =   "状态"
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
            DataField       =   "zcF"
            Caption         =   "收到"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "OK!"
               FalseValue      =   "欠款"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "htF"
            Caption         =   "htF"
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
            DataField       =   "DelF"
            Caption         =   "DelF"
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
         SplitCount      =   1
         BeginProperty Split0 
            ScrollBars      =   2
            BeginProperty Column00 
               Object.Visible         =   0   'False
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1454.74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   1814.74
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   3000.189
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
               ColumnWidth     =   2085.166
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dtgYf 
         Bindings        =   "wbMx.frx":00C6
         Height          =   2325
         Left            =   0
         TabIndex        =   107
         Top             =   2280
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   4101
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BackColor       =   14548971
         ForeColor       =   12582912
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
         WrapCellPointer =   -1  'True
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
         ColumnCount     =   18
         BeginProperty Column00 
            DataField       =   "YiRq"
            Caption         =   "收款日期"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dddddd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "yfJe"
            Caption         =   "收款金额"
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
         BeginProperty Column03 
            DataField       =   "yingRQ"
            Caption         =   "yingRQ"
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
            DataField       =   "htF"
            Caption         =   "htF"
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
            DataField       =   "zcF"
            Caption         =   "zcF"
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
            DataField       =   "yWy"
            Caption         =   "yWy"
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
            DataField       =   "YingJe"
            Caption         =   "YingJe"
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
            DataField       =   "DelF"
            Caption         =   "DelF"
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
         BeginProperty Column10 
            DataField       =   "fkFc"
            Caption         =   "付款方式"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "银行"
               FalseValue      =   "现金"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "yinHang"
            Caption         =   "银 行"
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
            DataField       =   "qianKuan1"
            Caption         =   "qianKuan1"
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
            DataField       =   "qianKuan2"
            Caption         =   "qianKuan2"
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
            DataField       =   "qianKuan3"
            Caption         =   "qianKuan3"
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
            DataField       =   "qianKuan4"
            Caption         =   "qianKuan4"
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
            DataField       =   "qianKuan5"
            Caption         =   "qianKuan5"
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
            DataField       =   "qianKuan6"
            Caption         =   "qianKuan6"
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
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
               ColumnWidth     =   1814.74
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column10 
               Button          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column11 
               Object.Visible         =   0   'False
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column13 
               Object.Visible         =   0   'False
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   0   'False
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column16 
               Object.Visible         =   0   'False
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column17 
               Object.Visible         =   0   'False
               ColumnWidth     =   915.024
            EndProperty
         EndProperty
      End
      Begin VB.Label Label43 
         Caption         =   "付款条件备注："
         Height          =   225
         Left            =   3810
         TabIndex        =   116
         Top             =   3150
         Width           =   1275
      End
      Begin VB.Label lblChg 
         Height          =   225
         Left            =   -66900
         TabIndex        =   113
         Top             =   2280
         Width           =   1245
      End
      Begin VB.Label Label39 
         Caption         =   "合计："
         Height          =   225
         Left            =   -67590
         TabIndex        =   112
         Top             =   2280
         Width           =   555
      End
      Begin VB.Label lblHtZe 
         Height          =   255
         Left            =   4860
         TabIndex        =   111
         Top             =   2550
         Width           =   1245
      End
      Begin VB.Label Label38 
         Caption         =   "合同总额："
         Height          =   285
         Left            =   3780
         TabIndex        =   110
         Top             =   2550
         Width           =   945
      End
      Begin VB.Label Label25 
         Caption         =   "工作单查询详情："
         Height          =   225
         Left            =   -65640
         TabIndex        =   109
         Top             =   4020
         Width           =   1635
      End
      Begin VB.Label Label47 
         Caption         =   "合计："
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
         Left            =   -68880
         TabIndex        =   98
         Top             =   3750
         Width           =   735
      End
      Begin VB.Label lblCF1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -68040
         TabIndex        =   97
         Top             =   3720
         Width           =   1245
      End
      Begin VB.Label Label45 
         Caption         =   "合计："
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
         Left            =   -72480
         TabIndex        =   96
         Top             =   3690
         Width           =   885
      End
      Begin VB.Label lblCF 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -71460
         TabIndex        =   95
         Top             =   3660
         Width           =   1245
      End
      Begin VB.Label lblHG1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -69270
         TabIndex        =   93
         Top             =   3300
         Width           =   1245
      End
      Begin VB.Label Label44 
         Caption         =   "合计："
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
         Left            =   -70200
         TabIndex        =   92
         Top             =   3330
         Width           =   885
      End
      Begin VB.Label lblHG 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -71520
         TabIndex        =   91
         Top             =   3270
         Width           =   1155
      End
      Begin VB.Label Label42 
         Caption         =   "合计："
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
         Left            =   -72420
         TabIndex        =   90
         Top             =   3300
         Width           =   885
      End
      Begin VB.Label Label41 
         Caption         =   "实际"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -67830
         TabIndex        =   89
         Top             =   540
         Width           =   1125
      End
      Begin VB.Label Label40 
         Caption         =   "预计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -72630
         TabIndex        =   88
         Top             =   510
         Width           =   1155
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "金额"
         Height          =   345
         Left            =   -68130
         TabIndex        =   87
         Top             =   930
         Width           =   1365
      End
      Begin VB.Label Label36 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "当地车费"
         Height          =   345
         Left            =   -69390
         TabIndex        =   86
         Top             =   3000
         Width           =   1245
      End
      Begin VB.Label Label35 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "餐费"
         Height          =   345
         Left            =   -69390
         TabIndex        =   85
         Top             =   2655
         Width           =   1245
      End
      Begin VB.Label Label34 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "住宿费"
         Height          =   345
         Left            =   -69390
         TabIndex        =   84
         Top             =   2310
         Width           =   1245
      End
      Begin VB.Label Label33 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "往返汽车票"
         Height          =   345
         Left            =   -69390
         TabIndex        =   83
         Top             =   1965
         Width           =   1245
      End
      Begin VB.Label Label32 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "往返火车票"
         Height          =   345
         Left            =   -69390
         TabIndex        =   82
         Top             =   1620
         Width           =   1245
      End
      Begin VB.Label Label31 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "往返机票"
         Height          =   345
         Left            =   -69390
         TabIndex        =   81
         Top             =   1275
         Width           =   1245
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "项目"
         Height          =   345
         Left            =   -69390
         TabIndex        =   80
         Top             =   930
         Width           =   1245
      End
      Begin VB.Label Label29 
         Caption         =   "实际"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -69120
         TabIndex        =   73
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label28 
         Caption         =   "预计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -73230
         TabIndex        =   72
         Top             =   690
         Width           =   1155
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "金额"
         Height          =   300
         Left            =   -68520
         TabIndex        =   71
         Top             =   1140
         Width           =   930
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "工时"
         Height          =   300
         Left            =   -69450
         TabIndex        =   70
         Top             =   1140
         Width           =   930
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "工程"
         Height          =   300
         Left            =   -70380
         TabIndex        =   69
         Top             =   2040
         Width           =   930
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "急修"
         Height          =   300
         Left            =   -70380
         TabIndex        =   68
         Top             =   1740
         Width           =   930
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "巡视"
         Height          =   300
         Left            =   -70380
         TabIndex        =   67
         Top             =   1440
         Width           =   930
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "项目"
         Height          =   300
         Left            =   -70380
         TabIndex        =   66
         Top             =   1140
         Width           =   930
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "大修"
         Height          =   300
         Left            =   -70380
         TabIndex        =   65
         Top             =   2340
         Width           =   930
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "小计"
         Height          =   300
         Left            =   -71760
         TabIndex        =   56
         Top             =   1140
         Width           =   960
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "工时"
         Height          =   300
         Left            =   -72720
         TabIndex        =   55
         Top             =   1140
         Width           =   960
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "单价"
         Height          =   300
         Left            =   -73680
         TabIndex        =   54
         Top             =   1140
         Width           =   960
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "工程"
         Height          =   300
         Left            =   -74640
         TabIndex        =   53
         Top             =   2040
         Width           =   960
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "急修"
         Height          =   300
         Left            =   -74640
         TabIndex        =   52
         Top             =   1740
         Width           =   960
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "巡视"
         Height          =   300
         Left            =   -74640
         TabIndex        =   51
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "项目"
         Height          =   300
         Left            =   -74640
         TabIndex        =   50
         Top             =   1140
         Width           =   960
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "大修"
         Height          =   300
         Left            =   -74640
         TabIndex        =   49
         Top             =   2340
         Width           =   960
      End
      Begin VB.Label Label19 
         Caption         =   "产品名称："
         Height          =   375
         Left            =   -74370
         TabIndex        =   30
         Top             =   2220
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "项目"
         Height          =   345
         Left            =   -74520
         TabIndex        =   28
         Top             =   930
         Width           =   1125
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "往返机票"
         Height          =   345
         Left            =   -74520
         TabIndex        =   27
         Top             =   1275
         Width           =   1125
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "往返火车票"
         Height          =   345
         Left            =   -74520
         TabIndex        =   26
         Top             =   1620
         Width           =   1125
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "往返汽车票"
         Height          =   345
         Left            =   -74520
         TabIndex        =   25
         Top             =   1965
         Width           =   1125
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "住宿费"
         Height          =   345
         Left            =   -74520
         TabIndex        =   24
         Top             =   2310
         Width           =   1125
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "餐费"
         Height          =   345
         Left            =   -74520
         TabIndex        =   23
         Top             =   2655
         Width           =   1125
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "当地车费"
         Height          =   345
         Left            =   -74520
         TabIndex        =   22
         Top             =   3000
         Width           =   1125
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "金额"
         Height          =   345
         Left            =   -73395
         TabIndex        =   21
         Top             =   930
         Width           =   1125
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "人次"
         Height          =   345
         Left            =   -72270
         TabIndex        =   20
         Top             =   930
         Width           =   1125
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "小计"
         Height          =   345
         Left            =   -71145
         TabIndex        =   19
         Top             =   930
         Width           =   1125
      End
   End
End
Attribute VB_Name = "wbMx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
adoRGF.Recordset.AddNew "htbh", wbHTP.txtHtbh.Text
Set dtgSale.DataSource = adoRGF
End Sub



Private Sub cmdCCou_Click()
'Dim taxQ As Single
txtJPXG.Text = Val(txtJPJE.Text) * Val(txtJPCou.Text)
txtHCXG.Text = Val(txtHCJE.Text) * Val(txtHCCou.Text)
txtQCXG.Text = Val(txtQCJE.Text) * Val(txtQCCou.Text)
txtZXG.Text = Val(txtZJE.Text) * Val(txtZCou.Text)
txtCXG.Text = Val(txtCJE.Text) * Val(txtCCou.Text)
txtDDXG.Text = Val(txtDDJE.Text) * Val(txtDDCou.Text)
lblCf.Caption = Val(txtJPXG.Text) + Val(txtHCXG.Text) + Val(txtQCXG.Text) + Val(txtZXG.Text) + Val(txtCXG.Text) + _
Val(txtDDXG.Text)
wbHTP.txtCLF1.Text = lblCf.Caption


wbHTP.txtCbze1.Text = Val(wbHTP.txtClcb1.Text) + Val(wbHTP.txtRgf1.Text) + Val(wbHTP.txtCLF1.Text) + Val(wbHTP.txtFbje1.Text) + _
Val(wbHTP.txtYf1.Text) + Val(wbHTP.txtQt1.Text)
'taxQ = Val(wbHTP.txtCbze1.Text) + Val(wbHTP.txtLr1.Text)
If wbHTP.optLa.Value = True Or wbHTP.optLb.Value = True Then
        wbHTP.txtJlr1.Text = Round(Val(wbHTP.txtHtze.Text) / 1.17 - Val(wbHTP.txtCbze1.Text), 2)
        wbHTP.txtLr1.Text = Round(Val(wbHTP.txtHtze.Text) / 1.17 - Val(wbHTP.txtCbze1.Text) - Val(wbHTP.txtYj1.Text), 2)
'wbHTP.txtHtze.Text = Round(taxQ * 1.17, 0)
ElseIf wbHTP.optLc.Value = True Then
'wbHTP.txtHtze.Text = Round(taxQ * 1.06, 0)
        wbHTP.txtJlr1.Text = Round(Val(wbHTP.txtHtze.Text) / 1.06 - Val(wbHTP.txtCbze1.Text), 2)
        wbHTP.txtLr1.Text = Round(Val(wbHTP.txtHtze.Text) / 1.06 - Val(wbHTP.txtCbze1.Text) - Val(wbHTP.txtYj1.Text), 2)
End If
wbMx.lblHtze.Caption = wbHTP.txtHtze.Text

End Sub

Private Sub cmdCou_Click()
'Dim taxQ As Single
txtXxG.Text = Val(txtXdj.Text) * Val(txtXgT.Text)
txtJxG.Text = Val(txtJdj.Text) * Val(txtJgT.Text)
txtGxG.Text = Val(txtGdj.Text) * Val(txtGgT.Text)
txtDxG.Text = Val(txtDdj.Text) * Val(txtDgT.Text)
lblHG.Caption = Val(txtXxG.Text) + Val(txtJxG.Text) + Val(txtGxG.Text) + Val(txtDxG.Text)
wbHTP.txtRgf1.Text = lblHG.Caption


wbHTP.txtCbze1.Text = Val(wbHTP.txtClcb1.Text) + Val(wbHTP.txtRgf1.Text) + Val(wbHTP.txtCLF1.Text) + Val(wbHTP.txtFbje1.Text) + _
Val(wbHTP.txtYf1.Text) + Val(wbHTP.txtQt1.Text)
'taxQ = Val(wbHTP.txtCbze1.Text) + Val(wbHTP.txtLr1.Text)
If wbHTP.optLa.Value = True Or wbHTP.optLb.Value = True Then
        wbHTP.txtJlr1.Text = Round(Val(wbHTP.txtHtze.Text) / 1.17 - Val(wbHTP.txtCbze1.Text), 2)
        wbHTP.txtLr1.Text = Round(Val(wbHTP.txtHtze.Text) / 1.17 - Val(wbHTP.txtCbze1.Text) - Val(wbHTP.txtYj1.Text), 2)
'wbHTP.txtHtze.Text = Round(taxQ * 1.17, 0)
ElseIf wbHTP.optLc.Value = True Then
'wbHTP.txtHtze.Text = Round(taxQ * 1.06, 0)
        wbHTP.txtJlr1.Text = Round(Val(wbHTP.txtHtze.Text) / 1.06 - Val(wbHTP.txtCbze1.Text), 2)
        wbHTP.txtLr1.Text = Round(Val(wbHTP.txtHtze.Text) / 1.06 - Val(wbHTP.txtCbze1.Text) - Val(wbHTP.txtYj1.Text), 2)
End If
wbMx.lblHtze.Caption = wbHTP.txtHtze.Text

End Sub

Private Sub cmdDel_Click()
On Error Resume Next
adoRGF.Recordset.Delete adAffectCurrent
'adoRGF.Recordset.UpdateBatch
End Sub

Private Sub cmdJadd_Click()
frmFuK.adoHpt.Recordset.AddNew "htbh", wbHTP.txtHtbh.Text
If wbHTP.optP.Value = True Then
frmFuK.adoHpt.Recordset.Update "htF", 0
ElseIf wbHTP.optZ.Value = True Then
frmFuK.adoHpt.Recordset.Update "htF", 1
End If
frmFuK.adoHpt.Recordset.Update "delF", 1
frmFuK.adoHpt.Recordset.Update "zcF", 0
frmFuK.adoHpt.Recordset.Update "khMc", wbHTP.txtKhmc.Text
frmFuK.adoHpt.Recordset.Update "yWy", wbHTP.txtYwy.Text
frmFuK.adoHpt.Recordset.Update "yifJe", 0
Set dtgFk.DataSource = frmFuK.adoHpt
End Sub

Private Sub cmdJdel_Click()
On Error Resume Next
frmFuK.adoHpt.Recordset.Delete adAffectCurrent

End Sub

Private Sub cmdJi_Click()
On Error Resume Next
'Dim taxQ As Single
Dim tt As String
Dim ii As Single
'tt = "Select sum(je) from htSale where htbh='" & wbHTP.txtHtbh.Text & "'"
'frmAdo.adoTmp.Recordset.Close
'frmAdo.adoTmp.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'lblChg.Caption = frmAdo.adoTmp.Recordset.Fields(0).Value
ii = 0
adoRGF.Recordset.MoveFirst
Do While Not adoRGF.Recordset.EOF
ii = ii + adoRGF.Recordset.Fields("je").Value
adoRGF.Recordset.MoveNext
Loop
lblChg.Caption = ii

'If wbHTP.optP.Value = True Or wbHTP.optG.Value = True Then

    wbHTP.txtClcb1.Text = lblChg.Caption
    
    
    
    wbHTP.txtCbze1.Text = Val(wbHTP.txtClcb1.Text) + Val(wbHTP.txtRgf1.Text) + Val(wbHTP.txtCLF1.Text) + Val(wbHTP.txtFbje1.Text) + _
    Val(wbHTP.txtYf1.Text) + Val(wbHTP.txtQt1.Text)
    'taxQ = Val(wbHTP.txtCbze1.Text) + Val(wbHTP.txtLr1.Text)
    If wbHTP.optLa.Value = True Or wbHTP.optLb.Value = True Then
            wbHTP.txtJlr1.Text = Round(Val(wbHTP.txtHtze.Text) / 1.17 - Val(wbHTP.txtCbze1.Text), 2)
            wbHTP.txtLr1.Text = Round(Val(wbHTP.txtHtze.Text) / 1.17 - Val(wbHTP.txtCbze1.Text) - Val(wbHTP.txtYj1.Text), 2)
    'wbHTP.txtHtze.Text = Round(taxQ * 1.17, 0)
    ElseIf wbHTP.optLc.Value = True Then
    'wbHTP.txtHtze.Text = Round(taxQ * 1.06, 0)
            wbHTP.txtJlr1.Text = Round(Val(wbHTP.txtHtze.Text) / 1.06 - Val(wbHTP.txtCbze1.Text), 2)
            wbHTP.txtLr1.Text = Round(Val(wbHTP.txtHtze.Text) / 1.06 - Val(wbHTP.txtCbze1.Text) - Val(wbHTP.txtYj1.Text), 2)
    End If
    wbMx.lblHtze.Caption = wbHTP.txtHtze.Text
'ElseIf wbHTP.optZ.Value = True Then
'    wbHTP.txtClcb2.Text = lblChg.Caption
'End If

End Sub

Private Sub cmdJmod_Click()
dtgFk.AllowUpdate = True
cmdJadd.Enabled = True
cmdJdel.Enabled = True
cmdJmod.Enabled = False '修改按钮禁用，使得提交时能以次为改动依据更新资金流量表
End Sub

Private Sub cmdMod1_Click()
cmdAdd.Enabled = True
cmdDel.Enabled = True
dtgSale.AllowUpdate = True
txtLjmc.Enabled = True
'If adoRGF.Recordset.RecordCount = 0 Then
'adoRGF.Recordset.AddNew "htbh", wbHTP.txtHtbh.Text
'Set dtgSale.DataSource = adoRGF
'End If
End Sub



Private Sub dtgFk_AfterColEdit(ByVal ColIndex As Integer)
On Error Resume Next
If ColIndex = 2 Then
frmFuK.adoHpt.Recordset.Update "ED", frmFuK.adoHpt.Recordset.Fields("ED").Value / 100
frmFuK.adoHpt.Recordset.Update "yingfJe", Val(lblHtze.Caption) * frmFuK.adoHpt.Recordset.Fields("ED").Value
ElseIf ColIndex = 3 Then
frmFuK.adoHpt.Recordset.Update "ED", frmFuK.adoHpt.Recordset.Fields("yingfJe").Value / Val(lblHtze.Caption)
End If

End Sub

Private Sub dtgFk_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub dtgGzb_Click()
On Error Resume Next
cmdGzd.Caption = form2Htp.adoGzb.Recordset.Fields("max(gzb.htBh)").Value
End Sub

Private Sub dtgGzb_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
cmdGzd.Caption = form2Htp.adoGzb.Recordset.Fields("max(gzb.htBh)").Value
End Sub

Private Sub dtgLjmc_DblClick()
On Error Resume Next
adoRGF.Recordset.Update "hpBm", adoLjmc.Recordset.Fields("hpBm").Value '货品编码
adoRGF.Recordset.Update "ljMc", adoLjmc.Recordset.Fields("pmGg").Value '产品名称
adoRGF.Recordset.Update "phBiao", adoLjmc.Recordset.Fields("phBiao").Value '规格型号
adoRGF.Recordset.Update "hpLb", adoLjmc.Recordset.Fields("hpLb").Value '货品类别
adoRGF.Recordset.Update "jlDw", adoLjmc.Recordset.Fields("jlDw").Value '计量单位

dtgLjmc.Visible = False
End Sub

Private Sub dtgSale_AfterColEdit(ByVal ColIndex As Integer)
On Error Resume Next
If ColIndex = 6 Or ColIndex = 5 Then
adoRGF.Recordset.Update "je", adoRGF.Recordset.Fields("dj").Value * adoRGF.Recordset.Fields("ljSl").Value

End If
End Sub

Private Sub Form_Load()


wbMx.Height = 5025
wbMx.Width = 11985
dtgLjmc.Visible = False
dtgSale.AllowUpdate = False
cmdAdd.Enabled = False
cmdDel.Enabled = False




End Sub



Private Sub Form_Unload(Cancel As Integer)
If MDI.Cq = False Then
wbMx.Visible = False
Cancel = True
End If
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
dtgLjmc.Visible = False
End Sub

Private Sub txtLjmc_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Dim tt As String
If KeyCode = 13 Then
tt = "Select * from kc where pmGg like '%" & txtLjmc.Text & "%'"
adoLjmc.Recordset.Close
adoLjmc.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
Set dtgLjmc.DataSource = adoLjmc
If adoLjmc.Recordset.RecordCount > 0 Then
dtgLjmc.Visible = True
Else
dtgLjmc.Visible = False
End If
End If
End Sub

