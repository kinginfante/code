VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmPld 
   Caption         =   "配料单"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.Frame frmFc 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3015
      Left            =   8730
      TabIndex        =   90
      Top             =   5550
      Width           =   6255
      Begin VB.TextBox txtFcsl 
         Height          =   285
         Left            =   1650
         TabIndex        =   95
         Top             =   720
         Width           =   2265
      End
      Begin VB.TextBox txtFcBz 
         Height          =   1485
         Left            =   1650
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   94
         Top             =   1140
         Width           =   2265
      End
      Begin VB.CommandButton cmdFcAdd 
         Caption         =   "添加"
         Height          =   285
         Left            =   4290
         TabIndex        =   93
         Top             =   1710
         Width           =   645
      End
      Begin VB.CommandButton cmdFcDel 
         Caption         =   "删除"
         Height          =   285
         Left            =   4290
         TabIndex        =   92
         Top             =   2340
         Width           =   645
      End
      Begin VB.CommandButton cmdFcGx 
         Caption         =   "更新"
         Height          =   315
         Left            =   4290
         TabIndex        =   91
         Top             =   2010
         Width           =   645
      End
      Begin MSComCtl2.DTPicker dtpFc 
         Height          =   285
         Left            =   1650
         TabIndex        =   96
         Top             =   360
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   503
         _Version        =   393216
         Format          =   56033281
         CurrentDate     =   39312
      End
      Begin VB.Label lblFcid 
         Caption         =   "lblFcid"
         Height          =   285
         Left            =   4470
         TabIndex        =   101
         Top             =   1110
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label lblPid 
         Caption         =   "lblPid"
         Height          =   285
         Left            =   4410
         TabIndex        =   100
         Top             =   480
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label48 
         Caption         =   "日期"
         Height          =   225
         Left            =   900
         TabIndex        =   99
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label47 
         Caption         =   "数量"
         Height          =   225
         Left            =   900
         TabIndex        =   98
         Top             =   750
         Width           =   615
      End
      Begin VB.Label Label46 
         Caption         =   "备注"
         Height          =   225
         Left            =   900
         TabIndex        =   97
         Top             =   1080
         Width           =   615
      End
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6180
      Top             =   6930
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7170
      Top             =   6780
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgFc 
      Height          =   3765
      Left            =   60
      TabIndex        =   89
      Top             =   5430
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   6641
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3105
      Left            =   0
      TabIndex        =   20
      Top             =   720
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   5477
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "货品资料"
      TabPicture(0)   =   "frmPld.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dtgSale"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "备注"
      TabPicture(1)   =   "frmPld.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblBa"
      Tab(1).Control(1)=   "lblBB"
      Tab(1).Control(2)=   "lblBc"
      Tab(1).Control(3)=   "lblBd"
      Tab(1).Control(4)=   "lblBe"
      Tab(1).Control(5)=   "txtTa"
      Tab(1).Control(6)=   "txtTb"
      Tab(1).Control(7)=   "txtTc"
      Tab(1).Control(8)=   "txtTd"
      Tab(1).Control(9)=   "txtTe"
      Tab(1).ControlCount=   10
      Begin VB.TextBox txtTe 
         Height          =   2415
         Left            =   -64710
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   330
         Width           =   2415
      End
      Begin VB.TextBox txtTd 
         Height          =   2415
         Left            =   -67245
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   330
         Width           =   2415
      End
      Begin VB.TextBox txtTc 
         Height          =   2415
         Left            =   -69765
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   330
         Width           =   2415
      End
      Begin VB.TextBox txtTb 
         Height          =   2415
         Left            =   -72300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   330
         Width           =   2415
      End
      Begin VB.TextBox txtTa 
         Height          =   2415
         Left            =   -74820
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   330
         Width           =   2415
      End
      Begin MSDataGridLib.DataGrid dtgSale 
         Bindings        =   "frmPld.frx":0038
         Height          =   2745
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   4842
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   0   'False
         HeadLines       =   1
         RowHeight       =   20
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
         BeginProperty Column02 
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
         BeginProperty Column03 
            DataField       =   "jlDw"
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
            DataField       =   "ljSl"
            Caption         =   "数量"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "LLSL"
            Caption         =   "领料数量"
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
            DataField       =   "LLDJ"
            Caption         =   "领料单价"
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
            DataField       =   "LLF"
            Caption         =   "领料成本"
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
            DataField       =   "wfl"
            Caption         =   "wfl"
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
            DataField       =   "KCSl"
            Caption         =   "库存数量"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "CGSL"
            Caption         =   "采购数量"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "YRQ"
            Caption         =   "预计采购期"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "yyyy-M-d"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "DHL"
            Caption         =   "采购到货量"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column13 
            DataField       =   "DHRQ"
            Caption         =   "采购到货期"
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
            DataField       =   "GYF"
            Caption         =   "供应商"
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
            DataField       =   "YFL"
            Caption         =   "发货数量"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column16 
            DataField       =   "FRQ"
            Caption         =   "发货日期"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "yyyy-M-d"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column17 
            DataField       =   "pid"
            Caption         =   "pid"
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
            BeginProperty Column00 
               Object.Visible         =   -1  'True
               ColumnWidth     =   2399.811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2099.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column06 
               Button          =   -1  'True
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column09 
               Button          =   -1  'True
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column12 
               Button          =   -1  'True
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column14 
            EndProperty
            BeginProperty Column15 
               Button          =   -1  'True
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column17 
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
      Begin VB.Label lblBe 
         Caption         =   "配料单管理者-"
         Height          =   225
         Left            =   -64350
         TabIndex        =   31
         Tag             =   "配料单管理者-"
         Top             =   60
         Width           =   1845
      End
      Begin VB.Label lblBd 
         Caption         =   "配料单管理者-"
         Height          =   225
         Left            =   -66840
         TabIndex        =   30
         Tag             =   "配料单管理者-"
         Top             =   60
         Width           =   1845
      End
      Begin VB.Label lblBc 
         Caption         =   "配料单管理者-"
         Height          =   225
         Left            =   -69420
         TabIndex        =   29
         Tag             =   "配料单管理者-"
         Top             =   60
         Width           =   1845
      End
      Begin VB.Label lblBB 
         Caption         =   "货品确认者-"
         Height          =   225
         Left            =   -72000
         TabIndex        =   28
         Tag             =   "货品确认者-"
         Top             =   60
         Width           =   1845
      End
      Begin VB.Label lblBa 
         Caption         =   "配料单管理者-"
         Height          =   225
         Left            =   -74490
         TabIndex        =   27
         Tag             =   "配料单管理者-"
         Top             =   60
         Width           =   1845
      End
   End
   Begin VB.CommandButton cmdAD 
      Caption         =   "添加"
      Height          =   255
      Left            =   0
      TabIndex        =   88
      Top             =   1590
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdDE 
      Caption         =   "删除"
      Height          =   255
      Left            =   780
      TabIndex        =   87
      Top             =   1560
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Frame frmHide 
      Caption         =   "frmHid"
      Height          =   1575
      Left            =   4140
      TabIndex        =   75
      Top             =   3630
      Visible         =   0   'False
      Width           =   3855
      Begin VB.Label lblBm 
         Caption         =   "lblBm"
         Height          =   225
         Left            =   240
         TabIndex        =   86
         Top             =   300
         Width           =   555
      End
      Begin VB.Label lblQy 
         Caption         =   "lblQy"
         Height          =   255
         Left            =   2700
         TabIndex        =   85
         Top             =   210
         Width           =   1155
      End
      Begin VB.Label lblLc 
         Caption         =   "lblLc"
         Height          =   225
         Left            =   240
         TabIndex        =   84
         Top             =   600
         Width           =   645
      End
      Begin VB.Label lblNlb 
         Caption         =   "lblNlb"
         Height          =   225
         Left            =   1560
         TabIndex        =   83
         Top             =   630
         Width           =   645
      End
      Begin VB.Label lblLcRen 
         Caption         =   "lblLcRen"
         Height          =   195
         Left            =   300
         TabIndex        =   82
         Top             =   900
         Width           =   795
      End
      Begin VB.Label lblLcUid 
         Caption         =   "lblLcUid"
         Height          =   285
         Left            =   270
         TabIndex        =   81
         Top             =   1170
         Width           =   885
      End
      Begin VB.Label lblFwid 
         Caption         =   "lblFwid"
         Height          =   255
         Left            =   1470
         TabIndex        =   80
         Top             =   270
         Width           =   885
      End
      Begin VB.Label lblUid 
         Caption         =   "lblUid"
         Height          =   255
         Left            =   2670
         TabIndex        =   79
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblYwy 
         Caption         =   "lblYwy"
         Height          =   285
         Left            =   2610
         TabIndex        =   78
         Top             =   510
         Width           =   765
      End
      Begin VB.Label lblLcou 
         Caption         =   "lblLcou"
         Height          =   255
         Left            =   1590
         TabIndex        =   77
         Top             =   1020
         Width           =   915
      End
      Begin VB.Label lblPwf 
         Caption         =   "lblPwf"
         Height          =   225
         Left            =   2610
         TabIndex        =   76
         Top             =   1140
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdYw 
      Height          =   405
      Left            =   90
      TabIndex        =   73
      Top             =   4410
      Width           =   975
   End
   Begin VB.TextBox txtGdyy 
      Height          =   555
      Left            =   8040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   61
      Top             =   4080
      Width           =   7185
   End
   Begin VB.CommandButton cmdGd 
      BackColor       =   &H00C0FFC0&
      Caption         =   "改  单"
      Height          =   345
      Left            =   8580
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   4920
      Width           =   915
   End
   Begin VB.TextBox txtCB 
      Height          =   285
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   360
      Width           =   1365
   End
   Begin VB.CommandButton cmdCB 
      BackColor       =   &H00C0FFFF&
      Caption         =   "成本结算"
      Height          =   345
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   4920
      Width           =   885
   End
   Begin VB.CommandButton cmdDy 
      BackColor       =   &H00C0FFC0&
      Caption         =   "多余采购"
      Height          =   345
      Left            =   9570
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   4920
      Width           =   915
   End
   Begin VB.TextBox txtZfyy 
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   12960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   53
      Top             =   4080
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.CommandButton cmdZf 
      BackColor       =   &H00C0C0FF&
      Caption         =   "作废"
      Height          =   555
      Left            =   13380
      Picture         =   "frmPld.frx":004C
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   4740
      Width           =   675
   End
   Begin VB.Frame Frame1 
      Caption         =   "旧单子"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4185
      Left            =   60
      TabIndex        =   34
      Top             =   8520
      Width           =   15165
      Begin VB.CommandButton cmdRight 
         Caption         =   ">>"
         Height          =   435
         Left            =   13995
         TabIndex        =   48
         Top             =   3390
         Width           =   675
      End
      Begin VB.CommandButton cmdLeft 
         Caption         =   "<<"
         Height          =   435
         Left            =   13230
         TabIndex        =   47
         Top             =   3390
         Width           =   675
      End
      Begin VB.CommandButton cmdJa 
         Height          =   405
         Left            =   360
         TabIndex        =   39
         Top             =   2970
         Width           =   975
      End
      Begin VB.CommandButton cmdJb 
         Height          =   405
         Left            =   1605
         TabIndex        =   38
         Top             =   2970
         Width           =   975
      End
      Begin VB.CommandButton cmdJc 
         Height          =   405
         Left            =   2850
         TabIndex        =   37
         Top             =   2970
         Width           =   975
      End
      Begin VB.CommandButton cmdJd 
         Height          =   405
         Left            =   4095
         TabIndex        =   36
         Top             =   2970
         Width           =   975
      End
      Begin VB.CommandButton cmdJe 
         Height          =   405
         Left            =   5340
         TabIndex        =   35
         Top             =   2970
         Width           =   975
      End
      Begin MSAdodcLib.Adodc adoJu 
         Height          =   330
         Left            =   6750
         Top             =   2880
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
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
      Begin MSDataGridLib.DataGrid dtgJu 
         Bindings        =   "frmPld.frx":06B6
         Height          =   2805
         Left            =   0
         TabIndex        =   59
         Top             =   0
         Visible         =   0   'False
         Width           =   15165
         _ExtentX        =   26749
         _ExtentY        =   4948
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   0   'False
         HeadLines       =   1
         RowHeight       =   20
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
         ColumnCount     =   13
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
         BeginProperty Column02 
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
         BeginProperty Column03 
            DataField       =   "jlDw"
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
         BeginProperty Column05 
            DataField       =   "KCSl"
            Caption         =   "库存数量"
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
            DataField       =   "CGSL"
            Caption         =   "采购数量"
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
            DataField       =   "YRQ"
            Caption         =   "预计采购期"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "yyyy-M-d"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "DHL"
            Caption         =   "采购到货量"
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
            DataField       =   "DHRQ"
            Caption         =   "采购到货期"
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
            DataField       =   "GYF"
            Caption         =   "供应商"
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
            DataField       =   "YFL"
            Caption         =   "发货数量"
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
            DataField       =   "FRQ"
            Caption         =   "发货日期"
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
            BeginProperty Column00 
               Object.Visible         =   -1  'True
               ColumnWidth     =   2399.811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2099.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column05 
               Button          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column08 
               Button          =   -1  'True
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1200.189
            EndProperty
         EndProperty
      End
      Begin VB.Label lblOKDRQ 
         Caption         =   "kdrq"
         Height          =   375
         Left            =   10800
         TabIndex        =   66
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label lblOKD 
         Caption         =   "开单日期:"
         Height          =   255
         Left            =   9870
         TabIndex        =   65
         Top             =   3270
         Width           =   915
      End
      Begin VB.Label lblJb 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1635
         TabIndex        =   46
         Top             =   3390
         Width           =   885
      End
      Begin VB.Label lblJa 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   390
         TabIndex        =   45
         Top             =   3390
         Width           =   885
      End
      Begin VB.Label lblJd 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   4140
         TabIndex        =   44
         Top             =   3390
         Width           =   885
      End
      Begin VB.Label lblJc 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   2895
         TabIndex        =   43
         Top             =   3390
         Width           =   885
      End
      Begin VB.Label lblJe 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   5400
         TabIndex        =   42
         Top             =   3390
         Width           =   885
      End
      Begin VB.Label Label12 
         Caption         =   "配料单编号:"
         Height          =   195
         Left            =   9660
         TabIndex        =   41
         Top             =   270
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblJid 
         Caption         =   "Label6"
         Height          =   225
         Left            =   10800
         TabIndex        =   40
         Top             =   270
         Visible         =   0   'False
         Width           =   1665
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "提交"
      Height          =   555
      Left            =   12600
      Picture         =   "frmPld.frx":06CA
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4740
      Width           =   735
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "返回"
      Height          =   555
      Left            =   14100
      Picture         =   "frmPld.frx":0D34
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4740
      Width           =   645
   End
   Begin VB.CommandButton cmdQME 
      Height          =   405
      Left            =   5610
      TabIndex        =   12
      Top             =   4410
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdQMD 
      Height          =   405
      Left            =   4506
      TabIndex        =   11
      Top             =   4410
      Width           =   975
   End
   Begin VB.CommandButton cmdQMC 
      Height          =   405
      Left            =   3402
      TabIndex        =   10
      Top             =   4410
      Width           =   975
   End
   Begin VB.CommandButton cmdQMB 
      Height          =   405
      Left            =   2298
      TabIndex        =   9
      Top             =   4410
      Width           =   975
   End
   Begin VB.CommandButton cmdQMA 
      Height          =   405
      Left            =   1194
      TabIndex        =   8
      Top             =   4410
      Width           =   975
   End
   Begin MSAdodcLib.Adodc adoHp 
      Height          =   330
      Left            =   7020
      Top             =   4680
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
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
   Begin VB.TextBox txtHtze 
      Height          =   270
      Left            =   5010
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   360
      Width           =   1305
   End
   Begin VB.TextBox txtHtbh 
      Height          =   270
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   2925
   End
   Begin VB.TextBox txtXmmc 
      Height          =   270
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   2925
   End
   Begin VB.TextBox txtKhAdr 
      Height          =   270
      Left            =   5010
      TabIndex        =   3
      Top             =   0
      Width           =   3675
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   4275
      Left            =   -30
      Top             =   5310
      Width           =   15225
   End
   Begin VB.Label lblYwT 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   90
      TabIndex        =   74
      Top             =   4860
      Width           =   885
   End
   Begin VB.Label lblYw 
      Caption         =   "业务员"
      Height          =   225
      Left            =   120
      TabIndex        =   72
      Top             =   4110
      Width           =   825
   End
   Begin VB.Label lblQmE 
      Caption         =   "成本结算"
      Height          =   225
      Left            =   5640
      TabIndex        =   71
      Top             =   4110
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lblQmD 
      Caption         =   "发货"
      Height          =   225
      Left            =   4530
      TabIndex        =   70
      Top             =   4110
      Width           =   1005
   End
   Begin VB.Label lblQmC 
      Caption         =   "采购"
      Height          =   225
      Left            =   3405
      TabIndex        =   69
      Top             =   4110
      Width           =   1005
   End
   Begin VB.Label lblQmB 
      Caption         =   "库存确认"
      Height          =   225
      Left            =   2295
      TabIndex        =   68
      Top             =   4110
      Width           =   1005
   End
   Begin VB.Label lblQmA 
      Caption         =   "销售经理"
      Height          =   225
      Left            =   1170
      TabIndex        =   67
      Top             =   4110
      Width           =   1005
   End
   Begin VB.Label lblKDRQ 
      Caption         =   "kdRq"
      Height          =   375
      Left            =   10320
      TabIndex        =   64
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblKD 
      Caption         =   "开单日期:"
      Height          =   255
      Left            =   9390
      TabIndex        =   63
      Top             =   420
      Width           =   915
   End
   Begin VB.Label lblGdYY 
      Caption         =   "备注："
      Height          =   165
      Left            =   7350
      TabIndex        =   62
      Top             =   4170
      Width           =   615
   End
   Begin VB.Label lblXZ 
      Caption         =   "XZ"
      Height          =   255
      Left            =   13650
      TabIndex        =   60
      Top             =   210
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblCB 
      Caption         =   "成本总额"
      Height          =   165
      Left            =   6510
      TabIndex        =   56
      Top             =   420
      Width           =   765
   End
   Begin VB.Label lblZfyy 
      Caption         =   "作废原因"
      ForeColor       =   &H000000FF&
      Height          =   165
      Left            =   12150
      TabIndex        =   52
      Top             =   4260
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblZT 
      Caption         =   "ZT"
      ForeColor       =   &H00C000C0&
      Height          =   165
      Left            =   11760
      TabIndex        =   51
      Top             =   270
      Width           =   1905
   End
   Begin VB.Label lblGuid 
      Caption         =   "Guid"
      Height          =   255
      Left            =   13500
      TabIndex        =   49
      Top             =   390
      Width           =   1485
   End
   Begin VB.Label lblPmid 
      Caption         =   "Pmid"
      Height          =   225
      Left            =   10320
      TabIndex        =   33
      Top             =   30
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label Label5 
      Caption         =   "配料单编号:"
      Height          =   195
      Left            =   9210
      TabIndex        =   32
      Top             =   60
      Width           =   1215
   End
   Begin VB.Label lblTe 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   5670
      TabIndex        =   17
      Top             =   4860
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label lblTc 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   3426
      TabIndex        =   16
      Top             =   4860
      Width           =   885
   End
   Begin VB.Label lblTd 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   4548
      TabIndex        =   15
      Top             =   4860
      Width           =   885
   End
   Begin VB.Label lblTa 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   1182
      TabIndex        =   14
      Top             =   4860
      Width           =   885
   End
   Begin VB.Label lblTb 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   2304
      TabIndex        =   13
      Top             =   4860
      Width           =   885
   End
   Begin VB.Label Label4 
      Caption         =   "合同金额:"
      Height          =   225
      Left            =   4080
      TabIndex        =   6
      Top             =   420
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "项目地址:"
      Height          =   225
      Left            =   4080
      TabIndex        =   2
      Top             =   60
      Width           =   1245
   End
   Begin VB.Label Label2 
      Caption         =   "合同编号:"
      Height          =   255
      Left            =   150
      TabIndex        =   1
      Top             =   420
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "项目名称:"
      Height          =   285
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   945
   End
End
Attribute VB_Name = "frmPld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PZT As Integer '配料单状态 :1生成2货品确认3采购4发货5成本

Dim adoFc As ADODB.Recordset
Dim timZm As Integer  ' (1购发编辑)

Dim Lpid As Long
Private Sub cmdAD_Click()
On Error Resume Next
If adoHp.Recordset.RecordCount > 0 Then
adoHp.Recordset.MoveLast
    If (IsNull(adoHp.Recordset.Fields("ljMc").Value) = True And IsNull(adoHp.Recordset.Fields("phBiao").Value) = True And _
       IsNull(adoHp.Recordset.Fields("ljBh").Value) = True) Then
       Exit Sub
    End If
End If
adoHp.Recordset.AddNew "PmId", lblPmid.Caption
adoHp.Recordset.Update "htbh", txtHtbh.Text
Set dtgSale.DataSource = adoHp

End Sub

Private Sub cmdCB_Click()
Dim tt As String
On Error Resume Next
frmPldCB.Show
frmPld.Enabled = False
frmPldCB.ZOrder 0

tt = "PLDBoundB(" & lblPmid.Caption & ")"
frmPldCB.adoHp.Recordset.Close
frmPldCB.adoHp.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
Set frmPldCB.dtgHp.DataSource = frmPldCB.adoHp

'打开采购多余表
tt = "PLDDy(" & lblPmid.Caption & ")"
frmPldCB.adoDy.Recordset.Close
frmPldCB.adoDy.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
Set frmPldCB.dtgDy.DataSource = frmPldCB.adoDy
If lblXZ.Caption = "WB" Then
    frmPldCB.dtgDy.Columns("结帐否").Visible = True
Else
    frmPldCB.dtgDy.Columns("结帐否").Visible = False
End If
frmPldCB.lblJe.Caption = frmPld.txtCB.Text
If lblLc.Caption = 6 Then
    frmPldCB.cmdJZ.Enabled = False
Else
    frmPldCB.cmdJZ.Enabled = True
End If
End Sub


Private Sub cmdDe_Click()
On Error Resume Next
If adoHp.Recordset.Fields("jzf").Value = True Then Exit Sub '不能删除已经结帐的记录
If adoHp.Recordset.RecordCount = 1 Then Exit Sub

adoHp.Recordset.Delete adAffectCurrent
'Set dtgSale.DataSource = adoHp
End Sub


Private Sub cmdDy_Click()
frmPldDy.Show
frmPldDy.ZOrder 0
frmPld.Enabled = False
frmPldDy.cmdDel.Visible = False
frmPldDy.cmdTj.Visible = False
frmPldDy.dtgDy.Columns("产品名称").Locked = True
frmPldDy.dtgDy.Columns("牌号商标").Locked = True
frmPldDy.dtgDy.Columns("规格型号").Locked = True
frmPldDy.dtgDy.Columns("单位").Locked = True
frmPldDy.dtgDy.Columns("多余采购数量").Locked = True
frmPldDy.dtgDy.Columns("单价").Locked = True
frmPldDy.dtgDy.Columns("金额").Locked = True
If mod1.PLB = True Then
    frmPldDy.cmdDel.Visible = True
    frmPldDy.cmdTj.Visible = True
    frmPldDy.dtgDy.Columns("多余采购数量").Locked = False
    
End If
If mod1.PLE = True Then
    frmPldDy.cmdTj.Visible = True
    frmPldDy.dtgDy.Columns("单价").Locked = False
    frmPldDy.dtgDy.Columns("金额").Locked = False
End If
End Sub

Private Sub cmdFh_Click()
frmPld.Height = 6000
End Sub

Private Sub cmdFcAdd_Click()
Dim tt As String
Dim hg As Single
On Error Resume Next
If Val(txtFcsl.Text) = 0 Or Val(lblPid.Caption) = 0 Then
Exit Sub
End If
Lpid = Val(lblPid.Caption)
timZm = 1 '购发编辑
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "配料单"
    mod1.cmd.Parameters("@NBLX") = "购发编辑"
    mod1.cmd.Parameters("@bh") = lblPmid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(txtHtbh.Text)  '合同编号
    mod1.cmd.Parameters("@mt2") = lblPid.Caption
    mod1.cmd.Parameters("@mt3") = lblFcid.Caption
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
    mod1.cmd.Parameters("@mt8") = ""
    mod1.cmd.Parameters("@mt9") = ""
    mod1.cmd.Parameters("@mt10") = ""
    mod1.cmd.Parameters("@mt11") = ""
    mod1.cmd.Parameters("@mt12") = ""
    mod1.cmd.Parameters("@mt13") = ""
    mod1.cmd.Parameters("@mt14") = ""
    mod1.cmd.Parameters("@mt15") = ""
    mod1.cmd.Parameters("@mt16") = ""
    mod1.cmd.Parameters("@mt17") = ""
    mod1.cmd.Parameters("@mt18") = ""
    mod1.cmd.Parameters("@mt19") = ""
    mod1.cmd.Parameters("@mt20") = ""
    mod1.cmd.Parameters("@mt21") = ""
    mod1.cmd.Parameters("@mt22") = ""
    mod1.cmd.Parameters("@mt23") = ""
    mod1.cmd.Parameters("@mt24") = ""
    mod1.cmd.Parameters("@mt25") = ""
    mod1.cmd.Parameters("@mlt1") = txtFcBz.Text
    mod1.cmd.Parameters("@mlt2") = ""
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtFcsl.Text)
    mod1.cmd.Parameters("@mm2") = 0
    mod1.cmd.Parameters("@mm3") = 0
    mod1.cmd.Parameters("@mm4") = 0
    mod1.cmd.Parameters("@mm5") = 0
    mod1.cmd.Parameters("@mm6") = 0
    mod1.cmd.Parameters("@mm7") = 0
    mod1.cmd.Parameters("@mm8") = 0
    mod1.cmd.Parameters("@mm9") = 0
    mod1.cmd.Parameters("@mm10") = 1 '添加
    mod1.cmd.Parameters("@mm11") = 1 '发货
    mod1.cmd.Parameters("@mm12") = 0
    mod1.cmd.Parameters("@mm13") = 0
    mod1.cmd.Parameters("@mm14") = 0
    mod1.cmd.Parameters("@mm15") = 0
    mod1.cmd.Parameters("@mm16") = 0
    mod1.cmd.Parameters("@mm17") = 0
    mod1.cmd.Parameters("@mm18") = 0
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = 0
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@mb2") = 0
    mod1.cmd.Parameters("@mb3") = 0
    mod1.cmd.Parameters("@mb4") = 0
    mod1.cmd.Parameters("@mb5") = 0
    mod1.cmd.Parameters("@md1") = dtpFc.Value
    mod1.cmd.Parameters("@md2") = Null
    mod1.cmd.Parameters("@md3") = Null
    mod1.cmd.Parameters("@md4") = Null
    mod1.cmd.Parameters("@md5") = Null
    Call mod1.REV
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


Set mod1.cmd = Nothing
frmFc.Enabled = False
End Sub

Private Sub cmdFcDel_Click()
Dim tt As String
Dim ii As Integer
Dim hg As Single
On Error Resume Next
If Val(txtFcsl.Text) = 0 Or Val(lblFcid.Caption) = 0 Then
Exit Sub
End If
Lpid = Val(lblPid.Caption)
ii = MsgBox("是否确认删除？", vbYesNo)
If ii = vbNo Then
    Exit Sub
End If
timZm = 1 '添加奖金
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "配料单"
    mod1.cmd.Parameters("@NBLX") = "购发编辑"
    mod1.cmd.Parameters("@bh") = lblPmid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(txtHtbh.Text)  '合同编号
    mod1.cmd.Parameters("@mt2") = lblPid.Caption
    mod1.cmd.Parameters("@mt3") = lblFcid.Caption
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
    mod1.cmd.Parameters("@mt8") = ""
    mod1.cmd.Parameters("@mt9") = ""
    mod1.cmd.Parameters("@mt10") = ""
    mod1.cmd.Parameters("@mt11") = ""
    mod1.cmd.Parameters("@mt12") = ""
    mod1.cmd.Parameters("@mt13") = ""
    mod1.cmd.Parameters("@mt14") = ""
    mod1.cmd.Parameters("@mt15") = ""
    mod1.cmd.Parameters("@mt16") = ""
    mod1.cmd.Parameters("@mt17") = ""
    mod1.cmd.Parameters("@mt18") = ""
    mod1.cmd.Parameters("@mt19") = ""
    mod1.cmd.Parameters("@mt20") = ""
    mod1.cmd.Parameters("@mt21") = ""
    mod1.cmd.Parameters("@mt22") = ""
    mod1.cmd.Parameters("@mt23") = ""
    mod1.cmd.Parameters("@mt24") = ""
    mod1.cmd.Parameters("@mt25") = ""
    mod1.cmd.Parameters("@mlt1") = txtFcBz.Text
    mod1.cmd.Parameters("@mlt2") = ""
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtFcsl.Text)
    mod1.cmd.Parameters("@mm2") = 0
    mod1.cmd.Parameters("@mm3") = 0
    mod1.cmd.Parameters("@mm4") = 0
    mod1.cmd.Parameters("@mm5") = 0
    mod1.cmd.Parameters("@mm6") = 0
    mod1.cmd.Parameters("@mm7") = 0
    mod1.cmd.Parameters("@mm8") = 0
    mod1.cmd.Parameters("@mm9") = 0
    mod1.cmd.Parameters("@mm10") = 3 '删除
    mod1.cmd.Parameters("@mm11") = 1 '发货
    mod1.cmd.Parameters("@mm12") = 0
    mod1.cmd.Parameters("@mm13") = 0
    mod1.cmd.Parameters("@mm14") = 0
    mod1.cmd.Parameters("@mm15") = 0
    mod1.cmd.Parameters("@mm16") = 0
    mod1.cmd.Parameters("@mm17") = 0
    mod1.cmd.Parameters("@mm18") = 0
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = 0
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@mb2") = 0
    mod1.cmd.Parameters("@mb3") = 0
    mod1.cmd.Parameters("@mb4") = 0
    mod1.cmd.Parameters("@mb5") = 0
    mod1.cmd.Parameters("@md1") = dtpFc.Value
    mod1.cmd.Parameters("@md2") = Null
    mod1.cmd.Parameters("@md3") = Null
    mod1.cmd.Parameters("@md4") = Null
    mod1.cmd.Parameters("@md5") = Null
    Call mod1.REV
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


Set mod1.cmd = Nothing
frmFc.Enabled = False
End Sub

Private Sub cmdFcGx_Click()
Dim tt As String
Dim hg As Single
On Error Resume Next
If Val(txtFcsl.Text) = 0 Or Val(lblFcid.Caption) = 0 Then
Exit Sub
End If
Lpid = Val(lblPid.Caption)
timZm = 1 '添加奖金
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "配料单"
    mod1.cmd.Parameters("@NBLX") = "购发编辑"
    mod1.cmd.Parameters("@bh") = lblPmid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(txtHtbh.Text)  '合同编号
    mod1.cmd.Parameters("@mt2") = lblPid.Caption
    mod1.cmd.Parameters("@mt3") = lblFcid.Caption
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
    mod1.cmd.Parameters("@mt8") = ""
    mod1.cmd.Parameters("@mt9") = ""
    mod1.cmd.Parameters("@mt10") = ""
    mod1.cmd.Parameters("@mt11") = ""
    mod1.cmd.Parameters("@mt12") = ""
    mod1.cmd.Parameters("@mt13") = ""
    mod1.cmd.Parameters("@mt14") = ""
    mod1.cmd.Parameters("@mt15") = ""
    mod1.cmd.Parameters("@mt16") = ""
    mod1.cmd.Parameters("@mt17") = ""
    mod1.cmd.Parameters("@mt18") = ""
    mod1.cmd.Parameters("@mt19") = ""
    mod1.cmd.Parameters("@mt20") = ""
    mod1.cmd.Parameters("@mt21") = ""
    mod1.cmd.Parameters("@mt22") = ""
    mod1.cmd.Parameters("@mt23") = ""
    mod1.cmd.Parameters("@mt24") = ""
    mod1.cmd.Parameters("@mt25") = ""
    mod1.cmd.Parameters("@mlt1") = txtFcBz.Text
    mod1.cmd.Parameters("@mlt2") = ""
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtFcsl.Text)
    mod1.cmd.Parameters("@mm2") = 0
    mod1.cmd.Parameters("@mm3") = 0
    mod1.cmd.Parameters("@mm4") = 0
    mod1.cmd.Parameters("@mm5") = 0
    mod1.cmd.Parameters("@mm6") = 0
    mod1.cmd.Parameters("@mm7") = 0
    mod1.cmd.Parameters("@mm8") = 0
    mod1.cmd.Parameters("@mm9") = 0
    mod1.cmd.Parameters("@mm10") = 2 '更新
    mod1.cmd.Parameters("@mm11") = 1 '发货
    mod1.cmd.Parameters("@mm12") = 0
    mod1.cmd.Parameters("@mm13") = 0
    mod1.cmd.Parameters("@mm14") = 0
    mod1.cmd.Parameters("@mm15") = 0
    mod1.cmd.Parameters("@mm16") = 0
    mod1.cmd.Parameters("@mm17") = 0
    mod1.cmd.Parameters("@mm18") = 0
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = 0
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@mb2") = 0
    mod1.cmd.Parameters("@mb3") = 0
    mod1.cmd.Parameters("@mb4") = 0
    mod1.cmd.Parameters("@mb5") = 0
    mod1.cmd.Parameters("@md1") = dtpFc.Value
    mod1.cmd.Parameters("@md2") = Null
    mod1.cmd.Parameters("@md3") = Null
    mod1.cmd.Parameters("@md4") = Null
    mod1.cmd.Parameters("@md5") = Null
    Call mod1.REV
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


Set mod1.cmd = Nothing
frmFc.Enabled = False
End Sub


Private Sub cmdGd_Click()
Dim Pmid As Long
Dim OldPmid As Long
Dim tt As String
Dim ii As String
Dim PP As String
On Error Resume Next

If lblZT.Caption = "此单已经作废" Then Exit Sub

PP = InputBox("请输入改单原因")
If PP = "" Then
    Exit Sub
Else
    txtGdyy.Text = PP
End If

ii = MsgBox("如果改单后,此配料单将返回到最初的流程,确认改单吗?", vbYesNo + vbDefaultButton2 + vbInformation, "请注意")
If ii = vbYes Then
    

    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "PLDgd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@htbh") = Trim(txtHtbh.Text)
    mod1.cmd.Parameters("@xmmc") = Trim(txtXMMC.Text)
    mod1.cmd.Parameters("@htze") = Trim(txtHtze.Text)
    mod1.cmd.Parameters("@krq") = mod1.DQda
    mod1.cmd.Parameters("@xmADR") = Trim(txtKhAdr.Text)
    mod1.cmd.Parameters("@Pmid") = Trim(lblPmid.Caption)
    mod1.cmd.Parameters("@Guid") = Trim(lblGuid.Caption)
    mod1.cmd.Parameters("@gdyy") = Trim(PP)
    mod1.cmd.Parameters("@xz") = lblXZ.Caption
    mod1.cmd.Parameters("@Tze") = Val(txtCB.Text)
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Execute
    Set cmd = Nothing
    

    
    '复制货品列表
    tt = "select pmid from maxPld"
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Pmid = mod1.HTP.Fields("pmid").Value
    
    tt = "Select * from PLD where PMid=" & Pmid
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText

    adoHp.Recordset.MoveFirst
    Do While Not adoHp.Recordset.EOF
        mod1.HTP.AddNew "htbh", txtHtbh.Text
        mod1.HTP.Update "pmid", Pmid
        mod1.HTP.Update "hpBm", adoHp.Recordset.Fields("hpBm").Value
        mod1.HTP.Update "ljmc", adoHp.Recordset.Fields("ljmc").Value
        mod1.HTP.Update "phBiao", adoHp.Recordset.Fields("phBiao").Value
        mod1.HTP.Update "ljbh", adoHp.Recordset.Fields("ljbh").Value
        mod1.HTP.Update "hplb", adoHp.Recordset.Fields("hplb").Value
        mod1.HTP.Update "jldw", adoHp.Recordset.Fields("jldw").Value
        mod1.HTP.Update "ljsl", adoHp.Recordset.Fields("ljsl").Value
        mod1.HTP.Update "YRQ", adoHp.Recordset.Fields("YRQ").Value
        mod1.HTP.Update "KCSl", adoHp.Recordset.Fields("KCSl").Value
        mod1.HTP.Update "CGSL", adoHp.Recordset.Fields("CGSL").Value
        mod1.HTP.Update "KCDJ", adoHp.Recordset.Fields("KCDJ").Value
        mod1.HTP.Update "CGDJ", adoHp.Recordset.Fields("CGDJ").Value
        mod1.HTP.Update "DHL", adoHp.Recordset.Fields("DHL").Value
        mod1.HTP.Update "DHRQ", adoHp.Recordset.Fields("DHRQ").Value
        mod1.HTP.Update "GYF", adoHp.Recordset.Fields("GYF").Value
        mod1.HTP.Update "je", adoHp.Recordset.Fields("je").Value
        mod1.HTP.Update "YFL", adoHp.Recordset.Fields("YFL").Value
        mod1.HTP.Update "FRQ", adoHp.Recordset.Fields("FRQ").Value
        mod1.HTP.Update "WFL", adoHp.Recordset.Fields("WFL").Value
        mod1.HTP.Update "FFF", adoHp.Recordset.Fields("FFF").Value
        mod1.HTP.Update "jzF", adoHp.Recordset.Fields("jzF").Value
        mod1.HTP.UpdateBatch
        adoHp.Recordset.MoveNext
    Loop
    cmdOld.Visible = True
    
    '更新当前配料单
    OldPmid = lblPmid.Caption
    Call modPld.PLDQing
    Call modPld.PLDBound(Pmid)
    Call modPld.PldOldBound(OldPmid)
    
    
    '更新显示列表
    frmPVa.optA.Value = True
    tt = "Select * from PldnewVa"
    mod1.PldV.Close
    mod1.PldV.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmPVa.dtgPld.DataSource = mod1.PldV
    
    '刷新旧单列表
    tt = "PldOldCount(" & lblGuid.Caption & ")"
    mod1.PldO.Close
    mod1.PldO.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    mod1.PldO.MoveLast
    cmdRight.Enabled = False
    
        frmPld.dtgSale.Columns("产品名称").Locked = False
        frmPld.dtgSale.Columns("牌号商标").Locked = False
        frmPld.dtgSale.Columns("规格型号").Locked = False
        frmPld.dtgSale.Columns("单位").Locked = False
        frmPld.dtgSale.Columns("数量").Locked = False
        frmPld.cmdAD.Visible = True
        frmPld.cmdDe.Visible = True
        frmPld.cmdSave.Enabled = True
        cmdGd.Enabled = False
        
        frmPld.Height = 10305
End If


End Sub


























Private Sub cmdLeft_Click()
On Error Resume Next
'清旧单
frmPld.lblJid = ""
frmPld.cmdJa.Caption = ""
frmPld.cmdJb.Caption = ""
frmPld.cmdJc.Caption = ""
frmPld.cmdJd.Caption = ""
frmPld.cmdJe.Caption = ""
frmPld.lblJa.Caption = ""
frmPld.lblJb.Caption = ""
frmPld.lblJc.Caption = ""
frmPld.lblJd.Caption = ""
frmPld.lblJe.Caption = ""
Set frmPld.dtgJu.DataSource = Nothing

mod1.PldO.MovePrevious

mod1.PldO.MovePrevious
If mod1.PldO.BOF = True Then
    cmdLeft.Enabled = False
    mod1.PldO.MoveNext
    Call modPld.PldOldBound(mod1.PldO.Fields("Pmid").Value)
    Exit Sub
End If
mod1.PldO.MoveNext
Call modPld.PldOldBound(mod1.PldO.Fields("Pmid").Value)
cmdRight.Enabled = True
End Sub

Private Sub cmdOld_Click()
frmPld.Height = 10305
End Sub

Private Sub cmdQMA_Click()
Dim tt As String
Dim Tywy As String '单子流转到下一人的姓名
Dim Tuid As String
Dim Oywy As String '原来流转人的名字
Dim Ouid As String '原来流转人的工号

On Error Resume Next
Dim ii As Integer
If lblLc.Caption <> 2 Then Exit Sub
If lblZT.Caption = "此单已经作废" Then Exit Sub
If cmdQMA.Caption <> "" Or lblLcRen.Caption = "" Then Exit Sub
If lblLcUid.Caption <> mod1.DHid Then
    MsgBox "此处应由" & lblLcRen.Caption & "签字! 请您不要再点"
    Exit Sub
End If
If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存!"
    cmdSave.Enabled = True
    Exit Sub
End If
'对于开单者,如果其签过字,则将生成新单子,保留老单子
Oywy = lblLcRen.Caption
Ouid = lblLcUid.Caption
ii = MsgBox("签字后,这张配料单将转到下个工序,您确认吗?", vbInformation + vbYesNo, "Hello!")
If ii = vbNo Then Exit Sub


cmdQMA.Caption = mod1.DName
lblTa.Caption = mod1.DQda
cmdSave.Enabled = False
'cmdGd.Visible = True
frmPld.dtgSale.Columns("产品名称").Locked = True
frmPld.dtgSale.Columns("牌号商标").Locked = True
frmPld.dtgSale.Columns("规格型号").Locked = True
frmPld.dtgSale.Columns("单位").Locked = True
frmPld.dtgSale.Columns("数量").Locked = True
frmPld.cmdAD.Visible = False
frmPld.cmdDe.Visible = False

'lblLc.Caption = lblLc.Caption + 1
'tt = "Select username,userid from worker where plb=1 and zzf=1 and qy='" & lblQy.Caption & "'"
'Set mod1.HTP = New ADODB.Recordset
'mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'
'Tywy = mod1.HTP.Fields("username").Value
'Tuid = mod1.HTP.Fields("userid").Value
'Set mod1.HTP = New ADODB.Recordset
'lblLcRen.Caption = Tywy
'lblLcUid.Caption = Tuid
'tt = "update pldMain set lc=" & Val(lblLc.Caption) & ",lcren='" & lblLcRen.Caption & "',lcuid='" & _
'    lblLcUid.Caption & "',qma='" & cmdQMA.Caption & "',qmat='" & lblTa.Caption & "' where pmid=" & lblPmid.Caption
'Set mod1.HTP = New ADODB.Recordset
'mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'
'Call mod1.EnventAdd("配料单", txtXmmc.Text, lblLcRen.Caption, lblLcUid.Caption, lblPmid.Caption, lblQmB.Caption, Oywy, Ouid, lblYwy.Caption, lblUid.Caption, lblFwid.Caption, lblPmid.Caption)
Set mod1.cmd = New ADODB.command
mod1.cmd.ActiveConnection = mod1.CC
mod1.cmd.CommandText = "xtzxAdd"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@yid").Value = 19
mod1.cmd.Parameters("@lc").Value = lblLc.Caption
mod1.cmd.Parameters("@bh").Value = lblPmid.Caption
mod1.cmd.Parameters("@ywy").Value = mod1.DName
mod1.cmd.Parameters("@uid").Value = mod1.DHid
mod1.cmd.Parameters("@bz").Value = ""
mod1.cmd.Execute
Set cmd = Nothing
MsgBox "OK!"
lblLc.Caption = lblLc.Caption + 1
lblLcRen.Caption = ""
lblLcUid.Caption = ""

If Dialog.Visible = True Then '更新事务列表
    Call mod1.refEnvent
End If
End Sub

Private Sub cmdQMB_Click()
Dim ii As Integer
Dim tt As String
Dim Ja As Boolean
Dim Li As Single '临时多余数量
Dim Gi As Single
Dim Tywy As String '单子流转到下一人的姓名
Dim Tuid As String
Dim Oywy As String '原来流转人的名字
Dim Ouid As String '原来流转人的工号
Dim TL As Boolean '是否跳流程
On Error Resume Next
TL = False
If lblLc.Caption = 2 And cmdQMA.Caption <> "" And cmdQMB.Caption = "" Then lblLc.Caption = 3
If lblLc.Caption <> 3 Then Exit Sub
If cmdQMB.Caption <> "" Then Exit Sub

If lblZT.Caption = "此单已经作废" Then Exit Sub
If lblLcUid.Caption <> mod1.DHid Then
    MsgBox "此处应由" & lblLcRen.Caption & "签字! 请您不要再点"
    Exit Sub
End If

If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存!"
    cmdSave.Enabled = True
    Exit Sub
End If
If cmdRight.Enabled = True And frmPld.Height > 6000 Then
    MsgBox ("请将旧单调整到最后一张,以便统计多购货品")
    cmdRight.SetFocus
    Exit Sub
End If

Oywy = lblLcRen.Caption
Ouid = lblLcUid.Caption
ii = MsgBox("签字后,这张配料单将转到下个工序,您确认吗?", vbInformation + vbYesNo, "Hello!")
If ii = vbNo Then Exit Sub



'如果什么都末填,则库存数量为0,全部为采购
adoHp.Recordset.MoveFirst
Do While Not adoHp.Recordset.EOF

    If adoHp.Recordset.Fields("Kcsl").Value = 0 Then
        'adoHp.Recordset.Fields("kcsl").Value = 0
        adoHp.Recordset.Fields("CGsl").Value = adoHp.Recordset.Fields("ljsl").Value
    End If
    adoHp.Recordset.MoveNext
Loop

'如果全部为库存,则直接跳到出货
lblLc.Caption = 4
adoHp.Recordset.MoveFirst
Do While Not adoHp.Recordset.EOF
    If adoHp.Recordset.Fields("ljsl").Value > adoHp.Recordset.Fields("kcsl").Value Then
        lblLc.Caption = 3
        Exit Do
    End If
    adoHp.Recordset.MoveNext
Loop
cmdQMB.Caption = mod1.DName
lblTb.Caption = mod1.DQda

frmPld.dtgSale.Columns("库存数量").Locked = True

frmPld.cmdAD.Visible = False
frmPld.cmdDe.Visible = False

''''''''如果已经购买过货品,则计算多余采购
'''''''adoJu.Recordset.MoveFirst
'''''''Do While Not adoJu.Recordset.EOF
'''''''    Ja = False
'''''''    Li = 0
'''''''    If adoJu.Recordset.Fields("DHL").Value > 0 Then
'''''''        adoHp.Recordset.MoveFirst
'''''''        Do While Not adoHp.Recordset.EOF
'''''''            '先删除多余表中的旧记录
'''''''            Set mod1.CMD = New ADODB.command
'''''''            mod1.CMD.ActiveConnection = mod1.CC
'''''''            mod1.CMD.CommandText = "PLDDDE"
'''''''            mod1.CMD.CommandType = adCmdStoredProc
'''''''            mod1.CMD.Parameters("@pmid") = lblPmid.Caption
'''''''            mod1.CMD.Parameters("@ljmc") = adoJu.Recordset.Fields("ljmc").Value
'''''''            mod1.CMD.Parameters("@phBiao") = adoJu.Recordset.Fields("phBiao").Value
'''''''            mod1.CMD.Parameters("@ljbh") = adoJu.Recordset.Fields("ljbh").Value
'''''''            mod1.CMD.Parameters("@Omid") = lblJid.Caption
'''''''            mod1.CMD.Execute
'''''''            Set CMD = Nothing
'''''''                Gi = 0
'''''''                If adoHp.Recordset.Fields("ljMc").Value = adoJu.Recordset.Fields("ljMc").Value And _
'''''''                    adoHp.Recordset.Fields("phBiao").Value = adoJu.Recordset.Fields("phBiao").Value And _
'''''''                    adoHp.Recordset.Fields("ljBh").Value = adoJu.Recordset.Fields("ljBh").Value Then
'''''''                    Gi = adoHp.Recordset.Fields("ljsl").Value - adoJu.Recordset.Fields("kcsl").Value
'''''''                        If Gi < 0 Then Gi = 0
'''''''                    Li = adoJu.Recordset.Fields("DHL").Value - Gi
'''''''                    If Li > 0 Then
'''''''                        Ja = True
'''''''
'''''''
'''''''                    End If
'''''''                    Exit Do
'''''''                End If
'''''''
'''''''            adoHp.Recordset.MoveNext
'''''''            If adoHp.Recordset.EOF = True Then
'''''''                Ja = True
'''''''                Li = adoJu.Recordset.Fields("DHL").Value
'''''''            End If
'''''''        Loop
'''''''    End If
'''''''    If Ja = True Then '加入多余表
'''''''        Set mod1.CMD = New ADODB.command
'''''''        mod1.CMD.ActiveConnection = mod1.CC
'''''''        mod1.CMD.CommandText = "PLDduo"
'''''''        mod1.CMD.CommandType = adCmdStoredProc
'''''''        mod1.CMD.Parameters("@pmid") = lblPmid.Caption
'''''''        mod1.CMD.Parameters("@ljmc") = adoJu.Recordset.Fields("ljmc").Value
'''''''        mod1.CMD.Parameters("@phBiao") = adoJu.Recordset.Fields("phBiao").Value
'''''''        mod1.CMD.Parameters("@ljbh") = adoJu.Recordset.Fields("ljbh").Value
'''''''        mod1.CMD.Parameters("@jldw") = adoJu.Recordset.Fields("jldw").Value
'''''''        mod1.CMD.Parameters("@sl") = Li
'''''''        mod1.CMD.Execute
'''''''        Set CMD = Nothing
'''''''    End If
'''''''    adoJu.Recordset.MoveNext
'''''''Loop
'''''''        '打开采购多余表
'''''''        tt = "PLDDy(" & lblPmid.Caption & ")"
'''''''        frmPldDy.adoDy.Recordset.Close
'''''''        frmPldDy.adoDy.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
'''''''        If frmPldDy.adoDy.Recordset.RecordCount > 0 Then
'''''''            Set frmPldDy.dtgDy.DataSource = frmPldDy.adoDy
'''''''            frmPld.cmdDy.Visible = True
'''''''        Else
'''''''            frmPld.cmdDy.Visible = False
'''''''        End If
    


'If lblLc.Caption = 3 Then
''    Tywy = "张春华"
''    Tuid = "HM001"
'    Tywy = lblLcRen.Caption
'    Tuid = lblLcUid.Caption
'ElseIf lblLc.Caption = 4 Then
'    Tywy = lblLcRen.Caption
'    Tuid = lblLcUid.Caption
'End If
'
'lblLcRen.Caption = Tywy
'lblLcUid.Caption = Tuid
'tt = "update pldMain set lc=" & Val(lblLc.Caption) & ",lcren='" & lblLcRen.Caption & "',lcuid='" & _
'    lblLcUid.Caption & "',Qmb='" & cmdQMB.Caption & "',qmbt='" & lblTb.Caption & "' where pmid=" & lblPmid.Caption
'Set mod1.HTP = New ADODB.Recordset
'mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'
'If lblLc.Caption = 3 Then
'    Call mod1.EnventAdd("配料单", txtXMMC.Text, lblLcRen.Caption, lblLcUid.Caption, lblPmid.Caption, lblQmC.Caption, Oywy, Ouid, lblYwy.Caption, lblUid.Caption, lblFwid.Caption, lblPmid.Caption)
'Else
'    Call mod1.EnventAdd("配料单", txtXMMC.Text, lblLcRen.Caption, lblLcUid.Caption, lblPmid.Caption, lblQmD.Caption, Oywy, Ouid, lblYwy.Caption, lblUid.Caption, lblFwid.Caption, lblPmid.Caption)
'End If
'MsgBox "OK!"
Set mod1.cmd = New ADODB.command
mod1.cmd.ActiveConnection = mod1.CC
mod1.cmd.CommandText = "xtzxAdd"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@yid").Value = 19
mod1.cmd.Parameters("@lc").Value = lblLc.Caption
mod1.cmd.Parameters("@bh").Value = lblPmid.Caption
mod1.cmd.Parameters("@ywy").Value = mod1.DName
mod1.cmd.Parameters("@uid").Value = mod1.DHid
mod1.cmd.Parameters("@bz").Value = ""
mod1.cmd.Execute
Set cmd = Nothing
MsgBox "OK!"

lblLc.Caption = lblLc.Caption + 1
lblLcRen.Caption = ""
lblLcUid.Caption = ""

If Dialog.Visible = True Then '更新事务列表
    Call mod1.refEnvent
End If
End Sub


Private Sub cmdQMC_Click()
Dim tt As String
Dim Tywy As String '单子流转到下一人的姓名
Dim Tuid As String
Dim Oywy As String '原来流转人的名字
Dim Ouid As String '原来流转人的工号

On Error Resume Next
Dim ii As Integer
If lblLc.Caption = 3 And cmdQMB.Caption <> "" And cmdQMC.Caption = "" Then lblLc.Caption = 4
If lblLc.Caption <> 4 Then Exit Sub
If lblZT.Caption = "此单已经作废" Then Exit Sub

If cmdQMC.Caption <> "" Then Exit Sub
If lblLcUid.Caption <> mod1.DHid Then
    MsgBox "此处应由" & lblLcRen.Caption & "签字! 请您不要再点"
    Exit Sub
End If
If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存!"
     cmdSave.Enabled = True
    Exit Sub
End If
Oywy = lblLcRen.Caption
Ouid = lblLcUid.Caption


ii = MsgBox("采购任务已经完成,您确认吗?", vbInformation + vbYesNo, "Hello!")
If ii = vbNo Then Exit Sub
'lblLc.Caption = lblLc.Caption + 1

cmdQMC.Caption = mod1.DName
lblTC.Caption = mod1.DQda

frmPld.dtgSale.Columns("预计采购期").Locked = True
frmPld.dtgSale.Columns("采购到货量").Locked = True
frmPld.dtgSale.Columns("供应商").Locked = True
frmPld.dtgSale.Columns("采购到货期").Locked = True
frmPld.cmdAD.Visible = False
frmPld.cmdDe.Visible = False

Set mod1.cmd = New ADODB.command
mod1.cmd.ActiveConnection = mod1.CC
mod1.cmd.CommandText = "xtzxAdd"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@yid").Value = 19
mod1.cmd.Parameters("@lc").Value = lblLc.Caption
mod1.cmd.Parameters("@bh").Value = lblPmid.Caption
mod1.cmd.Parameters("@ywy").Value = mod1.DName
mod1.cmd.Parameters("@uid").Value = mod1.DHid
mod1.cmd.Parameters("@bz").Value = ""
mod1.cmd.Execute
Set cmd = Nothing
MsgBox "OK!"

lblLc.Caption = lblLc.Caption + 1
lblLcRen.Caption = ""
lblLcUid.Caption = ""





'tt = "Select username,userid from worker where pld=1 and qy='" & mod1.Qy & "' and zzf=1"
'Set mod1.HTP = New ADODB.Recordset
'mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'
'Tywy = mod1.HTP.Fields("username").Value
'Tuid = mod1.HTP.Fields("userid").Value
'Set mod1.HTP = New ADODB.Recordset
'lblLcRen.Caption = Tywy
'lblLcUid.Caption = Tuid
'tt = "update pldMain set lc=" & Val(lblLc.Caption) & ",lcren='" & lblLcRen.Caption & "',lcuid='" & _
'    lblLcUid.Caption & "',qmc='" & cmdQMC.Caption & "',qmct='" & lblTC.Caption & "' where pmid=" & lblPmid.Caption
'Set mod1.HTP = New ADODB.Recordset
'mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'
'Call mod1.EnventAdd("配料单", txtXMMC.Text, lblLcRen.Caption, lblLcUid.Caption, lblPmid.Caption, lblQmD.Caption, Oywy, Ouid, lblYwy.Caption, lblUid.Caption, lblFwid.Caption, lblPmid.Caption)
'MsgBox "现在,此询价单将交由 " & Tywy & " 来审阅!"


If Dialog.Visible = True Then '更新事务列表
    Call mod1.refEnvent
End If
End Sub

Private Sub cmdQMD_Click()
Dim DuoC As Single
Dim Tywy As String '单子流转到下一人的姓名
Dim Tuid As String
Dim Oywy As String '原来流转人的名字
Dim Ouid As String '原来流转人的工号
On Error Resume Next
If lblZT.Caption = "此单已经作废" Then Exit Sub

'If cmdSave.Enabled = True Then
'    MsgBox "请先将单子保存!"
'    cmdSave.Enabled = True
'    Exit Sub
'End If
'If lblLc.Caption = 4 And cmdQMC.Caption <> "" And cmdQMD.Caption = "" Then lblLc.Caption = 5
If lblLc.Caption <> 5 Then Exit Sub
If cmdQMD.Caption <> "" Then Exit Sub
If lblLcUid.Caption <> mod1.DHid Then
    MsgBox "此处应由" & lblLcRen.Caption & "签字! 请您不要再点"
    Exit Sub
End If
If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存!"
    Exit Sub
End If

Oywy = lblLcRen.Caption
Ouid = lblLcUid.Caption
ii = MsgBox("采购任务已经完成,您确认吗?", vbInformation + vbYesNo, "Hello!")
If ii = vbNo Then Exit Sub

'lblLc.Caption = lblLc.Caption + 1
'lblLc.Caption = 5

'''''''''计算实际库存发货量和实际采购发货量
''''''''adoHp.Recordset.MoveFirst
''''''''Do While Not adoHp.Recordset.EOF
''''''''    If adoHp.Recordset.Fields("kcsl").Value <= adoHp.Recordset.Fields("Yfl").Value Then
''''''''        adoHp.Recordset.Fields("kzsl").Value = adoHp.Recordset.Fields("kcsl").Value
''''''''        adoHp.Recordset.Fields("czsl").Value = adoHp.Recordset.Fields("yfl").Value - adoHp.Recordset.Fields("kcsl").Value
''''''''    ElseIf adoHp.Recordset.Fields("kcsl").Value > adoHp.Recordset.Fields("Yfl").Value Then
''''''''        adoHp.Recordset.Fields("kzsl").Value = adoHp.Recordset.Fields("yfl").Value
''''''''        adoHp.Recordset.Fields("czsl").Value = 0
''''''''    End If
''''''''    DuoC = adoHp.Recordset.Fields("DHL").Value - adoHp.Recordset.Fields("czsl").Value
''''''''    If DuoC > 0 Then   '加入多余表
''''''''
''''''''        Set mod1.CMD = New ADODB.command
''''''''        mod1.CMD.ActiveConnection = mod1.CC
''''''''        mod1.CMD.CommandText = "PLDduo"
''''''''        mod1.CMD.CommandType = adCmdStoredProc
''''''''        mod1.CMD.Parameters("@pmid") = lblPmid.Caption
''''''''        mod1.CMD.Parameters("@ljmc") = adoHp.Recordset.Fields("ljmc").Value
''''''''        mod1.CMD.Parameters("@phBiao") = adoHp.Recordset.Fields("phBiao").Value
''''''''        mod1.CMD.Parameters("@ljbh") = adoHp.Recordset.Fields("ljbh").Value
''''''''        mod1.CMD.Parameters("@jldw") = adoHp.Recordset.Fields("jldw").Value
''''''''        mod1.CMD.Parameters("@sl") = DuoC
''''''''        mod1.CMD.Execute
''''''''        Set CMD = Nothing
''''''''    End If
''''''''    adoHp.Recordset.MoveNext
''''''''Loop
cmdQMD.Caption = mod1.DName
lblTd.Caption = mod1.DQda

frmPld.dtgSale.Columns("发货数量").Locked = True
frmPld.dtgSale.Columns("发货日期").Locked = True
frmPld.cmdAD.Visible = False
frmPld.cmdDe.Visible = False

Set mod1.cmd = New ADODB.command
mod1.cmd.ActiveConnection = mod1.CC
mod1.cmd.CommandText = "xtzxAdd"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@yid").Value = 19
mod1.cmd.Parameters("@lc").Value = lblLc.Caption
mod1.cmd.Parameters("@bh").Value = lblPmid.Caption
mod1.cmd.Parameters("@ywy").Value = mod1.DName
mod1.cmd.Parameters("@uid").Value = mod1.DHid
mod1.cmd.Parameters("@bz").Value = ""
mod1.cmd.Execute
Set cmd = Nothing
MsgBox "OK!"

lblLc.Caption = lblLc.Caption + 1
lblLcRen.Caption = ""
lblLcUid.Caption = ""
'tt = "Select username,userid from worker where ple=1 and zzf=1"
'Set mod1.HTP = New ADODB.Recordset
'mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'
'Tywy = mod1.HTP.Fields("username").Value
'Tuid = mod1.HTP.Fields("userid").Value
'Set mod1.HTP = New ADODB.Recordset
'lblLcRen.Caption = Tywy
'lblLcUid.Caption = Tuid
'tt = "update pldMain set lc=" & Val(lblLc.Caption) & ",lcren='" & lblLcRen.Caption & "',lcuid='" & _
'    lblLcUid.Caption & "',QMd='" & cmdQMD.Caption & "',qmdt='" & lblTd.Caption & "' where pmid=" & lblPmid.Caption
'Set mod1.HTP = New ADODB.Recordset
'mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'
'Call mod1.EnventAdd("配料单", txtXMMC.Text, lblLcRen.Caption, lblLcUid.Caption, lblPmid.Caption, lblQmE.Caption, Oywy, Ouid, lblYwy.Caption, lblUid.Caption, lblFwid.Caption, lblPmid.Caption)
'MsgBox "现在,此询价单将交由 " & Tywy & " 来审阅!"


If Dialog.Visible = True Then '更新事务列表
    Call mod1.refEnvent
End If
End Sub

Private Sub cmdQuit_Click()
Dim tt As String
On Error Resume Next
Call mod1.DelDKZ ' '退出表单时删除打开记录,以让别人能打开此单据
frmPld.Visible = False
If adoHp.Recordset.RecordCount = 0 Then
    tt = "PLDDel(" & lblPmid.Caption & ")"
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
End If
If frmHtZX.Visible = True Then
    frmHtZX.Enabled = True
    frmHtZX.ZOrder 0
ElseIf frmHtZxG.Visible = True Then
    frmHtZxG.Enabled = True
    frmHtZxG.ZOrder 0
ElseIf Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
End If
End Sub

Private Sub cmdRight_Click()
On Error Resume Next
'清旧单
frmPld.lblJid = ""
frmPld.cmdJa.Caption = ""
frmPld.cmdJb.Caption = ""
frmPld.cmdJc.Caption = ""
frmPld.cmdJd.Caption = ""
frmPld.cmdJe.Caption = ""
frmPld.lblJa.Caption = ""
frmPld.lblJb.Caption = ""
frmPld.lblJc.Caption = ""
frmPld.lblJd.Caption = ""
frmPld.lblJe.Caption = ""
Set frmPld.dtgJu.DataSource = Nothing

mod1.PldO.MoveNext
Call modPld.PldOldBound(mod1.PldO.Fields("Pmid").Value)
mod1.PldO.MoveNext
If mod1.PldO.EOF = True Then
    cmdRight.Enabled = False
    
End If
mod1.PldO.MovePrevious
    cmdLeft.Enabled = True
End Sub

Private Sub cmdSave_Click()
Dim tt As String
On Error Resume Next

Call modPld.PldJl(lblPmid.Caption)

''更新浏览列表
'If optA.Value = True Then
'    tt = "Select * from PldnewVa"
'ElseIf optB.Value = True Then
'    tt = "Select * from PldnewVb"
'ElseIf optC.Value = True Then
'    tt = "Select * from PldnewVc"
'ElseIf optD.Value = True Then
'    tt = "Select * from PldnewVac"
'ElseIf optE.Value = True Then
'    tt = "Select * from PldnewVb"
'ElseIf optF.Value = True Then
'    tt = "Select * from PldnewVc"
'ElseIf optG.Value = True Then
'    tt = "Select * from PldnewVad"
'ElseIf optH.Value = True Then
'    tt = "Select * from PldnewVb"
'ElseIf optI.Value = True Then
'    tt = "Select * from PldnewVc"
'ElseIf optJ.Value = True Then
'    tt = "Select * from PldnewVae"
'ElseIf optK.Value = True Then
'    tt = "Select * from PldnewVb"
'ElseIf optL.Value = True Then
'    tt = "Select * from PldnewVc"
'ElseIf optM.Value = True Then
'    tt = "Select * from PldnewVab"
'ElseIf optN.Value = True Then
'    tt = "Select * from PldnewVb"
'ElseIf optO.Value = True Then
'    tt = "Select * from PldnewVc"
'ElseIf optP.Value = True Then
'    tt = "Select * from PldnewVa"
'ElseIf optQ.Value = True Then
'    tt = "Select * from PldnewVb"
'ElseIf optR.Value = True Then
'    tt = "Select * from PldnewVc"
'ElseIf optZa.Value = True Or optZb.Value = True Or optZc.Value = True Or _
'       optZd.Value = True Or optZe.Value = True Or optZf.Value = True Then                         '作废单
'    tt = "Select * from PldZfV"
'End If
'
'    mod1.PldV.Close
'    mod1.PldV.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    'Set frmPVa.dtgPld.DataSource = mod1.PldV
If lblFwid.Caption = "" Then
    lblLc.Caption = 1
    tt = "update pldMain set lc=1 where pmid=" & lblPmid.Caption
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    '添加事务
    Call mod1.EnventAdd("配料单", txtXMMC.Text, lblLcRen.Caption, lblLcUid.Caption, lblPmid.Caption, lblYw.Caption, "", "", mod1.DName, mod1.DHid, 0, lblPmid.Caption)

End If



'更新询价列表
'tt = "select * from xunjiaView where ywy='" & mod1.DName & "' and uid='" & mod1.DHid & "'"
'frmGxBiao.adoXj.Close
'frmGxBiao.adoXj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
frmGxBiao.adoXj.Requery
Set frmGxBiao.dtgXj.DataSource = frmGxBiao.adoXj

frmHtZX.adoPld.Requery
Set frmHtZX.dtgPld.DataSource = frmHtZX.adoPld
cmdSave.Enabled = False
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Command11_Click()

End Sub

Private Sub cmdYw_Click()
Dim tt As String
Dim Tywy As String '单子流转到下一人的姓名
Dim Tuid As String
Dim Oywy As String '原来流转人的名字
Dim Ouid As String '原来流转人的工号

On Error Resume Next
Dim ii As Integer

If lblLc.Caption <> 1 Then Exit Sub
If lblZT.Caption = "此单已经作废" Then Exit Sub
If cmdYw.Caption <> "" Then Exit Sub


If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存!"
    Exit Sub
End If


Oywy = lblLcRen.Caption
Ouid = lblLcUid.Caption

'If lblLcUid.Caption <> mod1.DHid Then
If lblLcRen.Caption <> mod1.DName Then
    MsgBox "此处应由" & lblLcRen.Caption & "签字! 请您不要再点"
    Exit Sub
End If

'对于开单者,如果其签过字,则将生成新单子,保留老单子


ii = MsgBox("签字后,这张配料单将转到下个工序,您确认吗?", vbInformation + vbYesNo, "Hello!")
If ii = vbNo Then Exit Sub

'lblLc.Caption = lblLc.Caption + 1
'tt = "update pldMain set lc=" & Val(lblLc.Caption) & " where pmid=" & lblPmid.Caption
'Set mod1.HTP = New ADODB.Recordset
'mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText

cmdYw.Caption = mod1.DName
lblYwT.Caption = mod1.DQda
cmdSave.Enabled = False
'cmdGd.Visible = True
frmPld.dtgSale.Columns("产品名称").Locked = True
frmPld.dtgSale.Columns("牌号商标").Locked = True
frmPld.dtgSale.Columns("规格型号").Locked = True
frmPld.dtgSale.Columns("单位").Locked = True
frmPld.dtgSale.Columns("数量").Locked = True
frmPld.cmdAD.Visible = False
frmPld.cmdDe.Visible = False


'tt = "Select username,userid from worker where khk=1 and bm='" & lblBm.Caption & "' and zzf=1"
'Set mod1.HTP = New ADODB.Recordset
'mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'
'Tywy = mod1.HTP.Fields("username").Value
'Tuid = mod1.HTP.Fields("userid").Value
'Set mod1.HTP = New ADODB.Recordset
'lblLcRen.Caption = Tywy
'lblLcUid.Caption = Tuid
'tt = "update pldMain set lc=" & Val(lblLc.Caption) & ",lcren='" & lblLcRen.Caption & "',lcuid='" & _
'    lblLcUid.Caption & "',QMyw='" & cmdYw.Caption & "',qmyt='" & lblYwT.Caption & "' where pmid=" & lblPmid.Caption
'Set mod1.HTP = New ADODB.Recordset
'mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'
'Call mod1.EnventAdd("配料单", txtXmmc.Text, lblLcRen.Caption, lblLcUid.Caption, lblPmid.Caption, lblQmA.Caption, Oywy, Ouid, lblYwy.Caption, lblUid.Caption, lblFwid.Caption, lblPmid.Caption)
Set mod1.cmd = New ADODB.command
mod1.cmd.ActiveConnection = mod1.CC
mod1.cmd.CommandText = "xtzxAdd"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@yid").Value = 19
mod1.cmd.Parameters("@lc").Value = lblLc.Caption
mod1.cmd.Parameters("@bh").Value = lblPmid.Caption
mod1.cmd.Parameters("@ywy").Value = mod1.DName
mod1.cmd.Parameters("@uid").Value = mod1.DHid
mod1.cmd.Parameters("@bz").Value = ""
mod1.cmd.Execute
Set cmd = Nothing
MsgBox "现在,此询价单将交由销售经理来审阅!"
lblLc.Caption = lblLc.Caption + 1
lblLcRen.Caption = ""
lblLcUid.Caption = ""

If Dialog.Visible = True Then '更新事务列表
    Call mod1.refEnvent
End If

End Sub

Private Sub cmdZf_Click()
Dim tt As String
On Error Resume Next
Dim ii As Integer
If txtZfyy.Visible = False Then
    lblZfyy.Visible = True
    txtZfyy.Visible = True
    MsgBox ("请输入作废的原因!")
    txtZfyy.SetFocus
ElseIf txtZfyy.Text <> "" And txtZfyy.Visible = True Then
    ii = MsgBox("单子作废后,将不能恢复,您确认要作废吗?", vbExclamation + vbYesNo + vbDefaultButton2, "警告")
    If ii = vbYes Then
        frmPld.lblZT.Caption = "此单已经作废"
        frmPld.lblZT.ForeColor = &HFF&
        cmdZf.Enabled = False
        cmdGd.Visible = False
        cmdSave.Enabled = False
        tt = "PLDZF(" & lblPmid.Caption & ",'" & txtZfyy.Text & "','" & mod1.DName & "','" & mod1.DQda & "')"
        Set mod1.HTP = New ADODB.Recordset
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
      
    Else
        txtZfyy.Text = ""
        txtZfyy.Visible = False
        lblZfyy.Visible = False
        
    End If
End If
End Sub

Private Sub dtgFc_Click()
On Error Resume Next
dtgFc.Col = 1
txtFcsl.Text = Val(dtgFc.Text)
dtgFc.Col = 2
dtpFc.Value = dtgFc.Text
dtgFc.Col = 3
txtFcBz.Text = dtgFc.Text
dtgFc.Col = 4
lblFcid.Caption = dtgFc.Text

End Sub

Private Sub dtgFc_RowColChange()
On Error Resume Next
dtgFc.Col = 1
txtFcsl.Text = Val(dtgFc.Text)
dtgFc.Col = 2
dtpFc.Value = dtgFc.Text
dtgFc.Col = 3
txtFcBz.Text = dtgFc.Text
dtgFc.Col = 4
lblFcid.Caption = dtgFc.Text
End Sub


Private Sub dtgSale_AfterColUpdate(ByVal ColIndex As Integer)
On Error Resume Next

'If ColIndex = 5 And mod1.PLB = True Then
   adoHp.Recordset.Update "CGSL", adoHp.Recordset.Fields("ljsl").Value - adoHp.Recordset.Fields("Kcsl").Value
'End If
'更新未发量
'If ColIndex = 4 Then
    adoHp.Recordset.Update "WFL", adoHp.Recordset.Fields("ljsl").Value - adoHp.Recordset.Fields("YFL").Value
    adoHp.Recordset.Update "llF", adoHp.Recordset.Fields("llsl").Value * adoHp.Recordset.Fields("lldj").Value
'End If
End Sub

Private Sub dtgSale_ButtonClick(ByVal ColIndex As Integer)
Dim tt As String
Dim Htxz As String
Dim Bid As Long
'MsgBox ColIndex
On Error Resume Next
If ColIndex = 9 And mod1.PLB = True Then
    frmPldXT.Show
    frmPldXT.ZOrder 0
    frmPld.Enabled = False
    tt = "PLDXT(" & lblPmid.Caption & ",'" & adoHp.Recordset.Fields("ljMc").Value & "','" & _
         adoHp.Recordset.Fields("phBiao").Value & "','" & adoHp.Recordset.Fields("ljBh").Value & "')"
    mod1.PlDXT.Close
    mod1.PlDXT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    Set frmPldXT.dtgXT.DataSource = mod1.PlDXT
ElseIf ColIndex = 6 And mod1.PLB = True Then
    If adoHp.Recordset.Fields("lldj").Value > 0 Then
        Exit Sub
    End If
    '查找相应询价单中的成本
    tt = "select baoid,newF,htxz from htping where htbh='" & txtHtbh.Text & "'"
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Htxz = mod1.HTP.Fields("htxz").Value
    If mod1.HTP.Fields("newf").Value = False Then
        MsgBox ("此为旧合同评审单,所以不能提取单价!")
        Exit Sub
    End If
    tt = "select bid from baojiaD where baoid=" & mod1.HTP.Fields("baoid").Value
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    If mod1.HTP.RecordCount = 0 Then
        Exit Sub
    End If
    Bid = mod1.HTP.Fields("bid").Value
    If Htxz = "产品" Or Htxz = "零配件" Then

    Else '如果为大修或维保合的询价单,则要看零配件询价cgid值
        tt = "select cgid from xunJiaD where bid=" & Bid
        Set mod1.HTP = New ADODB.Recordset
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Bid = mod1.HTP.Fields("cgid").Value
        
    End If
    tt = "select * from XunJiaMx where bid=" & Bid
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'    adoHp.Recordset.MoveFirst
'    Do While Not adoHp.Recordset.EOF
'        mod1.HTP.MoveFirst
        Do While Not mod1.HTP.EOF
            If mod1.HTP.Fields("pz").Value = "产品" Then
                If mod1.HTP.Fields("jzPb").Value = adoHp.Recordset.Fields("phBiao").Value And _
                    mod1.HTP.Fields("jzxh").Value = adoHp.Recordset.Fields("ljBh").Value And _
                    mod1.HTP.Fields("ljMc").Value = adoHp.Recordset.Fields("ljMc").Value Then
                    adoHp.Recordset.Update "lldj", mod1.HTP.Fields("dj").Value
                    Exit Do
                End If
            Else
                If mod1.HTP.Fields("pbcd").Value = adoHp.Recordset.Fields("phBiao").Value And _
                    mod1.HTP.Fields("ljBh").Value = adoHp.Recordset.Fields("ljBh").Value And _
                    mod1.HTP.Fields("ljMc").Value = adoHp.Recordset.Fields("ljMc").Value Then
                    adoHp.Recordset.Update "lldj", mod1.HTP.Fields("dj").Value
                    Set frmPld.dtgSale.DataSource = adoHp
                    Exit Do
                End If
            End If
            mod1.HTP.MoveNext
        Loop
'        adoHp.Recordset.MoveNext
        
'    Loop
ElseIf ColIndex = 15 Then
'''    lblPid.Caption = dtgSale.Columns("pid").Value
'''    tt = "select sl as 数量,rq as 日期,bz as 备注,fcid from pldfc where pid=" & lblPid.Caption & " and fc=1 order by fcid desc"
'''    adoFc.Close
'''    adoFc.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''
'''    If adoFc.RecordCount = 0 Then
'''        Set dtgFc.DataSource = adoFc
'''        dtgFc.Rows = 2
'''        dtgFc.FixedRows = 0
'''        dtgFc.FixedRows = 1
'''
'''    Else
'''        dtgFc.Rows = 2
'''        dtgFc.FixedRows = 1
'''        Set dtgFc.DataSource = adoFc
'''    End If
    Me.Height = 9645
    dtgFc.ColWidth(0) = 300
    dtgFc.ColWidth(4) = 0
End If
End Sub

Private Sub dtgSale_Click()
Dim tt As String
On Error Resume Next
    lblPid.Caption = dtgSale.Columns("pid").Value
    
    tt = "select sl as 数量,rq as 日期,bz as 备注,fcid from pldfc where pid=" & lblPid.Caption & " and fc=1 order by fcid desc"
    adoFc.Close
    adoFc.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    
    If adoFc.RecordCount = 0 Then
        Set dtgFc.DataSource = adoFc
        dtgFc.Rows = 2
        dtgFc.FixedRows = 0
        dtgFc.FixedRows = 1

    Else
        dtgFc.Rows = 2
        dtgFc.FixedRows = 1
        Set dtgFc.DataSource = adoFc
    End If
'    Me.Height = 9645
    dtgFc.ColWidth(0) = 300
    dtgFc.ColWidth(4) = 0
End Sub


Private Sub dtgSale_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim tt As String
On Error Resume Next
    lblPid.Caption = dtgSale.Columns("pid").Value
    
    tt = "select sl as 数量,rq as 日期,bz as 备注,fcid from pldfc where pid=" & lblPid.Caption & " and fc=1 order by fcid desc"
    adoFc.Close
    adoFc.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    
    If adoFc.RecordCount = 0 Then
        Set dtgFc.DataSource = adoFc
        dtgFc.Rows = 2
        dtgFc.FixedRows = 0
        dtgFc.FixedRows = 1

    Else
        dtgFc.Rows = 2
        dtgFc.FixedRows = 1
        Set dtgFc.DataSource = adoFc
    End If
'    Me.Height = 9645
    dtgFc.ColWidth(0) = 300
    dtgFc.ColWidth(4) = 0
End Sub


Private Sub Form_Load()
frmPld.Top = 300
frmPld.Height = 6000
frmPld.Width = 14925
dtpFc.Value = Date

dtgFc.ColWidth(3) = 3000
Set adoFc = New ADODB.Recordset
If mod1.DName = "王蕾" Then
    frmFc.Visible = True
Else
    frmFc.Visible = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim tt As String
On Error Resume Next
If MDI.Cq = False Then
    Call mod1.DelDKZ ' '退出表单时删除打开记录,以让别人能打开此单据
    frmPld.Visible = False
    
    Cancel = True
    If adoHp.Recordset.RecordCount = 0 Then
        tt = "PLDDel(" & lblPmid.Caption & ")"
        Set mod1.HTP = New ADODB.Recordset
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
    End If
    If frmHtZX.Visible = True Then
        frmHtZX.Enabled = True
        frmHtZX.ZOrder 0
    ElseIf frmHtZxG.Visible = True Then
        frmHtZxG.Enabled = True
        frmHtZxG.ZOrder 0
    ElseIf Dialog.Visible = True Then
        Dialog.Enabled = True
        Dialog.ZOrder 0
    End If

End If
End Sub

Private Sub Label6_Click()

End Sub

Private Sub timQuit_Timer()
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
If timZm = 1 Then '购发编辑
    frmFc.Enabled = True
    lblFcid.Caption = ""

End If
'''    Call modNewHT.NewLocked
'''    cmdSave.Enabled = False
'''    If Val(lblLc.Caption) = 0 Then
'''        lblLc.Caption = 1
'''    End If
'''ElseIf timZm = 10 Then '签字
'''    cmdDing.Enabled = True
'''    txtQM.Text = ""
'''    frmQM.Visible = False
'''    lblTX.Visible = True
'''
'''ElseIf timZm = 11 Then
'''    cmdHt.Visible = False
'''ElseIf timZm = 12 Then '删除合同
'''    Me.Visible = False
'''    If htBrow.Visible = True Then
'''        htBrow.Enabled = True
'''        htBrow.ZOrder 0
'''        htBrow.adoBr.Requery

'''    End If
'''ElseIf timZm = 13 Then '添加奖金
'''    txtFED.Text = ""
'''    txtYingFu.Text = ""
'''End If
timQuit.Enabled = False
End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
Dim oo As Integer
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh,mm1,mt1,mm2,mt2 from ml where zid=" & mod1.Zid
Set mod1.WP = New ADODB.Recordset
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0

   If timZm = 1 Then '购发编辑
                
        txtFcsl.Text = ""
        txtFcBz.Text = ""

        
        adoFc.Requery
        If adoFc.RecordCount = 0 Then
            Set dtgFc.DataSource = adoFc
            dtgFc.Rows = 2
            dtgFc.FixedRows = 0
            dtgFc.FixedRows = 1
    
        Else
            dtgFc.Rows = 2
            dtgFc.FixedRows = 1
            Set dtgFc.DataSource = adoFc
        End If
        
'        adoHp.Recordset.Requery
'        Set frmPld.dtgSale.DataSource = frmPld.adoHp
        dtgSale.Columns("发货数量") = mod1.WP.Fields("mm1").Value
        
    End If
    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then '未完成

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '处理失败
    timWait.Enabled = False
    ii = MsgBox("服务中心在处理您的命令时,发生如下错误:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        frmFc.Enabled = True

    End If
    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        frmFc.Enabled = True

    End If
    Exit Sub
End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub


