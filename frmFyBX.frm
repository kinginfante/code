VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmFYBX 
   Caption         =   "费用报销"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   15210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   15210
   Begin VB.Frame frmLc 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "流程至:"
      Height          =   1125
      Left            =   8040
      TabIndex        =   127
      Top             =   6450
      Width           =   975
      Begin VB.Label lblLcRen 
         BackStyle       =   0  'Transparent
         Caption         =   "lblLcRen"
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   120
         TabIndex        =   129
         Top             =   630
         Width           =   795
      End
      Begin VB.Label Label31 
         Caption         =   "流程至:"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   120
         TabIndex        =   128
         Top             =   300
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdNQ 
      BackColor       =   &H008080FF&
      Caption         =   "审核"
      Height          =   375
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   124
      Top             =   8730
      Width           =   855
   End
   Begin VB.Frame frmNQ 
      BackColor       =   &H00C0FFC0&
      Caption         =   "评审建议"
      ForeColor       =   &H000000FF&
      Height          =   1785
      Left            =   1860
      TabIndex        =   100
      Top             =   7350
      Visible         =   0   'False
      Width           =   6195
      Begin VB.CommandButton cmdDing 
         BackColor       =   &H00FF8080&
         Caption         =   "决定"
         Height          =   285
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   1260
         Width           =   735
      End
      Begin VB.OptionButton optT2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "拒绝"
         Height          =   195
         Left            =   5130
         TabIndex        =   103
         Top             =   840
         Width           =   675
      End
      Begin VB.OptionButton OptT1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "同意"
         Height          =   225
         Left            =   5130
         TabIndex        =   102
         Top             =   450
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.TextBox txtQM 
         BackColor       =   &H00C0FFFF&
         Height          =   1365
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   101
         Top             =   300
         Width           =   4815
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgP 
      Height          =   2775
      Left            =   0
      TabIndex        =   125
      Top             =   6390
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   4895
      _Version        =   393216
      BackColor       =   12648447
      ForeColor       =   8404992
      Rows            =   15
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   16761024
      ForeColorFixed  =   0
      BackColorBkg    =   12648447
      GridColorFixed  =   8404992
      GridColorUnpopulated=   8404992
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.CommandButton cmdDao 
      Caption         =   "导入上月数据"
      Height          =   435
      Left            =   11520
      TabIndex        =   79
      Top             =   4650
      Width           =   1515
   End
   Begin MSAdodcLib.Adodc adoFy 
      Height          =   330
      Left            =   8160
      Top             =   7710
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\work\demo\HMXP9000\work.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\work\demo\HMXP9000\work.mdb;Persist Security Info=False"
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
   Begin VB.Frame frmAn 
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   12450
      TabIndex        =   63
      Top             =   8370
      Width           =   2925
      Begin VB.CommandButton cmdBack 
         Caption         =   "返回"
         Height          =   585
         Left            =   2160
         Picture         =   "frmFyBX.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   150
         Width           =   675
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "提交"
         Height          =   585
         Left            =   1440
         Picture         =   "frmFyBX.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   150
         Width           =   705
      End
      Begin VB.CommandButton cmdMod 
         Caption         =   "修改"
         Height          =   585
         Left            =   750
         Picture         =   "frmFyBX.frx":076C
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   150
         Width           =   675
      End
   End
   Begin VB.Frame frmMb 
      Caption         =   "Frame2"
      Height          =   9165
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15225
      Begin VB.TextBox txtCwBZ 
         Height          =   1965
         Left            =   10590
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   6390
         Width           =   4485
      End
      Begin VB.Frame frmG 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Left            =   10230
         TabIndex        =   117
         Top             =   4560
         Width           =   3135
         Begin VB.Label Label28 
            Caption         =   "固定费用"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   1560
            TabIndex        =   121
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label29 
            Caption         =   "变动费用"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   0
            TabIndex        =   120
            Top             =   0
            Width           =   735
         End
         Begin VB.Label lbl1 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   780
            TabIndex        =   119
            Top             =   0
            Width           =   645
         End
         Begin VB.Label lbl2 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2370
            TabIndex        =   118
            Top             =   0
            Width           =   645
         End
      End
      Begin VB.Frame frmED 
         Caption         =   "编辑栏"
         ForeColor       =   &H00C00000&
         Height          =   2985
         Left            =   9480
         TabIndex        =   85
         Top             =   1530
         Visible         =   0   'False
         Width           =   5565
         Begin VB.ComboBox txtBm 
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   4260
            TabIndex        =   123
            Text            =   "txtBm"
            Top             =   2610
            Width           =   1005
         End
         Begin VB.OptionButton opt2 
            Caption         =   "固定费用"
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   2640
            TabIndex        =   116
            Top             =   2640
            Width           =   1035
         End
         Begin VB.OptionButton opt1 
            Caption         =   "变动费用"
            ForeColor       =   &H00C00000&
            Height          =   225
            Left            =   1440
            TabIndex        =   115
            Top             =   2640
            Width           =   1095
         End
         Begin VB.TextBox txtGZDH 
            Height          =   270
            Left            =   1350
            TabIndex        =   108
            Top             =   2220
            Width           =   2865
         End
         Begin VB.CommandButton cmdGui 
            Caption         =   "费用归属"
            Height          =   405
            Left            =   4320
            TabIndex        =   99
            Top             =   570
            Width           =   1065
         End
         Begin VB.Timer timWait 
            Interval        =   1000
            Left            =   -30
            Top             =   360
         End
         Begin VB.Timer timQuit 
            Interval        =   1000
            Left            =   930
            Top             =   270
         End
         Begin VB.CommandButton cmdJdel 
            Caption         =   "删除"
            Height          =   255
            Left            =   4380
            TabIndex        =   96
            Top             =   2190
            Width           =   975
         End
         Begin VB.CommandButton cmdJed 
            Caption         =   "更新"
            Height          =   285
            Left            =   4350
            TabIndex        =   95
            Top             =   1830
            Width           =   1005
         End
         Begin VB.CommandButton cmdJadd 
            Caption         =   "添加"
            Height          =   285
            Left            =   4380
            TabIndex        =   94
            Top             =   1470
            Width           =   975
         End
         Begin VB.TextBox txtNr 
            Height          =   765
            Left            =   1350
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            ScrollBars      =   2  'Vertical
            TabIndex        =   93
            Top             =   1380
            Width           =   2865
         End
         Begin MSComCtl2.DTPicker dtpRq 
            Height          =   315
            Left            =   1350
            TabIndex        =   92
            Top             =   1020
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   393216
            Format          =   81133569
            CurrentDate     =   39316
         End
         Begin VB.ComboBox comLb 
            Height          =   300
            ItemData        =   "frmFyBX.frx":0A76
            Left            =   1350
            List            =   "frmFyBX.frx":0A78
            TabIndex        =   91
            Top             =   630
            Width           =   2895
         End
         Begin VB.TextBox txtJe 
            Height          =   330
            Left            =   1350
            TabIndex        =   90
            Top             =   240
            Width           =   2865
         End
         Begin VB.Label Label30 
            Caption         =   "部门"
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   3780
            TabIndex        =   122
            Top             =   2640
            Width           =   465
         End
         Begin VB.Label lblGZDH 
            Caption         =   "出租车注明（工程部填写工作单编号）"
            Height          =   885
            Left            =   210
            TabIndex        =   107
            Top             =   2190
            Width           =   1005
         End
         Begin VB.Label lblBid 
            Caption         =   "LblBid"
            Height          =   285
            Left            =   4500
            TabIndex        =   98
            Top             =   330
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label Label26 
            Caption         =   "内容"
            Height          =   345
            Left            =   240
            TabIndex        =   89
            Top             =   1440
            Width           =   825
         End
         Begin VB.Label Label25 
            Caption         =   "日期"
            Height          =   345
            Left            =   240
            TabIndex        =   88
            Top             =   1050
            Width           =   765
         End
         Begin VB.Label Label24 
            Caption         =   "费用类别"
            Height          =   345
            Left            =   240
            TabIndex        =   87
            Top             =   690
            Width           =   915
         End
         Begin VB.Label Label23 
            Caption         =   "金额"
            Height          =   375
            Left            =   270
            TabIndex        =   86
            Top             =   300
            Width           =   795
         End
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "复 制"
         Height          =   285
         Left            =   14220
         TabIndex        =   114
         Top             =   4710
         Width           =   945
      End
      Begin VB.CommandButton cmdXuan 
         Caption         =   "选 取"
         Height          =   285
         Left            =   13200
         TabIndex        =   113
         Top             =   4710
         Width           =   945
      End
      Begin VB.CommandButton cmdG 
         Caption         =   "费用归属"
         Height          =   405
         Left            =   10740
         TabIndex        =   106
         Top             =   5280
         Visible         =   0   'False
         Width           =   1065
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgNx 
         Height          =   2985
         Left            =   2400
         TabIndex        =   97
         Top             =   1530
         Width           =   15105
         _ExtentX        =   26644
         _ExtentY        =   5265
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox txtFP 
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   10560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   83
         Top             =   5310
         Width           =   4485
      End
      Begin VB.OptionButton optFp2 
         Caption         =   "不一致"
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   9420
         TabIndex        =   82
         Top             =   5700
         Width           =   915
      End
      Begin VB.OptionButton optFp1 
         Caption         =   "发票一致"
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   9420
         TabIndex        =   81
         Top             =   5400
         Width           =   1125
      End
      Begin VB.TextBox txtQc 
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   10560
         PasswordChar    =   "*"
         TabIndex        =   74
         Top             =   8730
         Width           =   825
      End
      Begin VB.ComboBox comDQ 
         Height          =   300
         ItemData        =   "frmFyBX.frx":0A7A
         Left            =   9330
         List            =   "frmFyBX.frx":0A7C
         TabIndex        =   73
         Top             =   7680
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Frame frmWd 
         Height          =   2175
         Left            =   10830
         TabIndex        =   54
         Top             =   2280
         Width           =   4485
         Begin MSDataListLib.DataCombo comYwy 
            Height          =   330
            Left            =   1110
            TabIndex        =   62
            Top             =   1140
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.TextBox txtHtbh 
            Height          =   270
            Left            =   1110
            TabIndex        =   57
            Top             =   750
            Width           =   3195
         End
         Begin VB.CommandButton cmdXQ 
            Caption         =   "售 前"
            Height          =   285
            Left            =   2730
            TabIndex        =   56
            Top             =   1140
            Width           =   1575
         End
         Begin VB.ComboBox comXmmc 
            Height          =   300
            ItemData        =   "frmFyBX.frx":0A7E
            Left            =   1110
            List            =   "frmFyBX.frx":0A80
            TabIndex        =   55
            Top             =   1590
            Width           =   3225
         End
         Begin VB.Label Label13 
            Caption         =   "合同编号"
            Height          =   255
            Left            =   150
            TabIndex        =   61
            Top             =   780
            Width           =   855
         End
         Begin VB.Label Label17 
            Caption         =   "   请输入合同编号,再按回车键,如果是售前服务,则请选择业务员及其相应的项目"
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   150
            TabIndex        =   60
            Top             =   270
            Width           =   4215
         End
         Begin VB.Label Label18 
            Caption         =   "项目名称:"
            Height          =   285
            Left            =   150
            TabIndex        =   59
            Top             =   1680
            Width           =   945
         End
         Begin VB.Label Label19 
            Caption         =   "业务员:"
            Height          =   225
            Left            =   150
            TabIndex        =   58
            Top             =   1200
            Width           =   915
         End
      End
      Begin VB.Frame frmRen 
         Caption         =   "frmRen"
         Height          =   465
         Left            =   12330
         TabIndex        =   50
         Top             =   1080
         Visible         =   0   'False
         Width           =   1785
         Begin VB.Label lblGuid 
            Caption         =   "Label13"
            Height          =   255
            Left            =   2100
            TabIndex        =   53
            Top             =   210
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label lblGui 
            Caption         =   "Label13"
            Height          =   255
            Left            =   960
            TabIndex        =   52
            Top             =   210
            Width           =   1605
         End
         Begin VB.Label Label8 
            Caption         =   "归属:"
            Height          =   405
            Left            =   120
            TabIndex        =   51
            Top             =   210
            Width           =   1065
         End
      End
      Begin VB.Frame frmYf 
         Caption         =   "运费归属:"
         Height          =   1545
         Left            =   11040
         TabIndex        =   44
         Top             =   5790
         Width           =   3975
         Begin VB.ComboBox comBm 
            Height          =   300
            ItemData        =   "frmFyBX.frx":0A82
            Left            =   2550
            List            =   "frmFyBX.frx":0A92
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   990
            Width           =   1125
         End
         Begin VB.ComboBox comhtBh 
            Height          =   300
            Left            =   1020
            TabIndex        =   46
            Top             =   390
            Width           =   2655
         End
         Begin VB.CommandButton cmdKc 
            Caption         =   "划 归 库 存 ->"
            Height          =   315
            Left            =   90
            TabIndex        =   45
            ToolTipText     =   "如果此笔费用不归项目,则归属该总部库存"
            Top             =   990
            Width           =   2445
         End
         Begin VB.Label Label12 
            Caption         =   "合同编号:"
            Height          =   255
            Left            =   90
            TabIndex        =   47
            Top             =   465
            Width           =   1005
         End
      End
      Begin VB.CommandButton cmdHg 
         Caption         =   "合计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9540
         TabIndex        =   39
         Top             =   4860
         Width           =   645
      End
      Begin VB.Frame frmQm 
         BorderStyle     =   0  'None
         Height          =   1365
         Left            =   3450
         TabIndex        =   23
         Top             =   7710
         Width           =   4155
         Begin VB.CommandButton cmdZj 
            Height          =   375
            Left            =   2130
            TabIndex        =   27
            Top             =   300
            Width           =   915
         End
         Begin VB.CommandButton cmdJl 
            Height          =   375
            Left            =   1110
            TabIndex        =   26
            Top             =   300
            Width           =   915
         End
         Begin VB.CommandButton cmdJc 
            Height          =   375
            Left            =   3150
            TabIndex        =   25
            Top             =   300
            Width           =   915
         End
         Begin VB.CommandButton cmdBxr 
            Height          =   375
            Left            =   90
            TabIndex        =   24
            Top             =   300
            Width           =   915
         End
         Begin VB.Label lblTb 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   1125
            TabIndex        =   38
            Top             =   750
            Width           =   885
         End
         Begin VB.Label lblTa 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   120
            TabIndex        =   37
            Top             =   750
            Width           =   885
         End
         Begin VB.Label lblTd 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   3150
            TabIndex        =   36
            Top             =   750
            Width           =   885
         End
         Begin VB.Label lblTc 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   2145
            TabIndex        =   35
            Top             =   750
            Width           =   885
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "经理"
            Height          =   225
            Left            =   1260
            TabIndex        =   31
            Top             =   30
            Width           =   615
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "审核人"
            Height          =   225
            Left            =   3300
            TabIndex        =   30
            Top             =   30
            Width           =   645
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "报销人"
            Height          =   225
            Left            =   240
            TabIndex        =   29
            Top             =   30
            Width           =   645
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "总经理"
            Height          =   225
            Left            =   2190
            TabIndex        =   28
            Top             =   30
            Width           =   735
         End
      End
      Begin VB.TextBox txtBz 
         Height          =   885
         Left            =   1020
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   5310
         Width           =   8205
      End
      Begin VB.TextBox txtHg 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10410
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   4860
         Width           =   1395
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "添加"
         Height          =   345
         Left            =   150
         TabIndex        =   10
         Top             =   4650
         Width           =   765
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除"
         Height          =   345
         Left            =   990
         TabIndex        =   9
         Top             =   4650
         Width           =   765
      End
      Begin MSComCtl2.DTPicker dtpLdate 
         Height          =   285
         Left            =   10020
         TabIndex        =   7
         Top             =   420
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Format          =   81133569
         CurrentDate     =   38287
      End
      Begin MSAdodcLib.Adodc adoF2 
         Height          =   375
         Left            =   2130
         Top             =   4650
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\work\demo\HmXP9000\work.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\work\demo\HmXP9000\work.mdb;Persist Security Info=False"
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
      Begin MSDataGridLib.DataGrid dtgBx 
         Bindings        =   "frmFyBX.frx":0AB8
         Height          =   2865
         Left            =   0
         TabIndex        =   34
         Top             =   1530
         Width           =   15075
         _ExtentX        =   26591
         _ExtentY        =   5054
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   17
         TabAcrossSplits =   -1  'True
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
         ColumnCount     =   52
         BeginProperty Column00 
            DataField       =   "aTime"
            Caption         =   "日期"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "yyyy.MM.dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "khmc"
            Caption         =   "项目名称"
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
            DataField       =   "TXF"
            Caption         =   "通信费"
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
            DataField       =   "NJTF"
            Caption         =   "市内交通费"
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
            DataField       =   "WJTF"
            Caption         =   "市外交通费"
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
            DataField       =   "ZCF"
            Caption         =   "住宿费"
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
            DataField       =   "CF"
            Caption         =   "餐费"
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
            DataField       =   "zdf"
            Caption         =   "招待费"
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
            DataField       =   "LPF"
            Caption         =   "礼品费"
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
            DataField       =   "TCF"
            Caption         =   "停车费"
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
            DataField       =   "qtF"
            Caption         =   "福利费"
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
            DataField       =   "YF"
            Caption         =   "运费"
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
            DataField       =   "BGYP"
            Caption         =   "办公用品"
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
            DataField       =   "GG"
            Caption         =   "工具"
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
            DataField       =   "YH"
            Caption         =   "易耗"
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
            DataField       =   "wl"
            Caption         =   "外劳"
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
            DataField       =   "KDF"
            Caption         =   "快递费"
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
            DataField       =   "zwbt"
            Caption         =   "驻外津贴"
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
            DataField       =   "SZTG"
            Caption         =   "市场推广"
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
            DataField       =   "RYZP"
            Caption         =   "人员招聘"
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
            DataField       =   "PXF"
            Caption         =   "培训费"
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
         BeginProperty Column21 
            DataField       =   "BMTD"
            Caption         =   "部门团队费"
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
         BeginProperty Column22 
            DataField       =   "TDJS"
            Caption         =   "团队建设费"
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
         BeginProperty Column23 
            DataField       =   "CWSX"
            Caption         =   "财务手续费"
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
         BeginProperty Column24 
            DataField       =   "FZ"
            Caption         =   "房租"
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
         BeginProperty Column25 
            DataField       =   "WYF"
            Caption         =   "物业费"
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
         BeginProperty Column26 
            DataField       =   "SD"
            Caption         =   "水电"
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
         BeginProperty Column27 
            DataField       =   "DW"
            Caption         =   "电话"
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
         BeginProperty Column28 
            DataField       =   "GTCF"
            Caption         =   "公共停车费"
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
         BeginProperty Column29 
            DataField       =   "GCLF"
            Caption         =   "公共车辆费"
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
         BeginProperty Column30 
            DataField       =   "clf"
            Caption         =   "车辆费"
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
         BeginProperty Column31 
            DataField       =   "FWBT"
            Caption         =   "房屋补贴"
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
         BeginProperty Column32 
            DataField       =   "lyf"
            Caption         =   "旅游费"
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
         BeginProperty Column33 
            DataField       =   "jtbt"
            Caption         =   "交通补贴"
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
         BeginProperty Column34 
            DataField       =   "zhbx"
            Caption         =   "综合保险"
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
         BeginProperty Column35 
            DataField       =   "gwbt"
            Caption         =   "岗位补贴"
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
         BeginProperty Column36 
            DataField       =   "GWF"
            Caption         =   "高温费"
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
         BeginProperty Column37 
            DataField       =   "sj"
            Caption         =   "三金"
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
         BeginProperty Column38 
            DataField       =   "gjj"
            Caption         =   "公积金"
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
         BeginProperty Column39 
            DataField       =   "htbh"
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
         BeginProperty Column40 
            DataField       =   "qy"
            Caption         =   "区域"
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
         BeginProperty Column41 
            DataField       =   "Bm"
            Caption         =   "部门"
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
         BeginProperty Column42 
            DataField       =   "ywy"
            Caption         =   "归属人"
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
         BeginProperty Column43 
            DataField       =   "YQZ"
            Caption         =   "归属人签字"
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
         BeginProperty Column44 
            DataField       =   "yqRq"
            Caption         =   "签字时间"
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
         BeginProperty Column45 
            DataField       =   "YwJl"
            Caption         =   "归经理"
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
         BeginProperty Column46 
            DataField       =   "YWQ"
            Caption         =   "部门经理签字"
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
         BeginProperty Column47 
            DataField       =   "ywRq"
            Caption         =   "签字日期"
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
         BeginProperty Column48 
            DataField       =   "xg"
            Caption         =   "小计"
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
         BeginProperty Column49 
            DataField       =   "gzdh"
            Caption         =   "出租车注明"
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
         BeginProperty Column50 
            DataField       =   "ywyuid"
            Caption         =   "ywyuid"
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
         BeginProperty Column51 
            DataField       =   "qrq"
            Caption         =   "签收日期"
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
            AllowSizing     =   0   'False
            Size            =   578
            BeginProperty Column00 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   -1  'True
               ColumnWidth     =   3495.118
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column11 
               Object.Visible         =   0   'False
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column17 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column18 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column19 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column20 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column21 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column22 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column23 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column24 
            EndProperty
            BeginProperty Column25 
            EndProperty
            BeginProperty Column26 
            EndProperty
            BeginProperty Column27 
            EndProperty
            BeginProperty Column28 
            EndProperty
            BeginProperty Column29 
            EndProperty
            BeginProperty Column30 
            EndProperty
            BeginProperty Column31 
            EndProperty
            BeginProperty Column32 
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column33 
            EndProperty
            BeginProperty Column34 
            EndProperty
            BeginProperty Column35 
            EndProperty
            BeginProperty Column36 
            EndProperty
            BeginProperty Column37 
            EndProperty
            BeginProperty Column38 
            EndProperty
            BeginProperty Column39 
               ColumnWidth     =   1995.024
            EndProperty
            BeginProperty Column40 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column41 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column42 
               Locked          =   -1  'True
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column43 
               Button          =   -1  'True
               Locked          =   -1  'True
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column44 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column45 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column46 
               Button          =   -1  'True
               Locked          =   -1  'True
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column47 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column48 
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column49 
            EndProperty
            BeginProperty Column50 
            EndProperty
            BeginProperty Column51 
            EndProperty
         EndProperty
      End
      Begin VB.Label lblTX 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   150
         TabIndex        =   126
         Top             =   360
         Width           =   5475
      End
      Begin VB.Label lblDx 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6450
         TabIndex        =   12
         Top             =   4860
         Width           =   3435
      End
      Begin VB.Label lblNewF 
         Caption         =   "Label27"
         Height          =   255
         Left            =   13170
         TabIndex        =   105
         Top             =   240
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label Label22 
         Caption         =   "制单人:"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   3600
         TabIndex        =   84
         Top             =   1290
         Width           =   645
      End
      Begin VB.Label lblRQ 
         BackStyle       =   0  'Transparent
         Height          =   405
         Left            =   13320
         TabIndex        =   78
         Top             =   450
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "签收"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   9810
         TabIndex        =   77
         Top             =   8820
         Width           =   555
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "代签者"
         Height          =   285
         Left            =   9330
         TabIndex        =   76
         Top             =   7260
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "签收日期"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   12450
         TabIndex        =   75
         Top             =   570
         Width           =   825
      End
      Begin VB.Label lblLc 
         Caption         =   "lblLc"
         Height          =   315
         Left            =   1110
         TabIndex        =   71
         Top             =   7800
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblYqf 
         Caption         =   "lblYqf"
         Height          =   225
         Left            =   2190
         TabIndex        =   49
         Top             =   8700
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lblFwid 
         Caption         =   "lblFwid"
         Height          =   255
         Left            =   1050
         TabIndex        =   43
         Top             =   8730
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblUid 
         Caption         =   "lblUid"
         Height          =   285
         Left            =   180
         TabIndex        =   42
         Top             =   8640
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblYwy 
         Caption         =   "lblYwy"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   4470
         TabIndex        =   41
         Top             =   1290
         Width           =   795
      End
      Begin VB.Label lblLcUid 
         Caption         =   "lblLcUid"
         Height          =   285
         Left            =   60
         TabIndex        =   40
         Top             =   8040
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblNlb 
         Caption         =   "lblNlb"
         Height          =   225
         Left            =   240
         TabIndex        =   33
         Top             =   7170
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label LblTrq 
         BackStyle       =   0  'Transparent
         Caption         =   "Label21"
         Height          =   225
         Left            =   1140
         TabIndex        =   32
         Top             =   1290
         Width           =   2985
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "填单日期"
         Height          =   225
         Left            =   90
         TabIndex        =   22
         Top             =   1290
         Width           =   855
      End
      Begin VB.Label lblFr 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1140
         TabIndex        =   21
         Top             =   900
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblLr 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2310
         TabIndex        =   20
         Top             =   900
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "日期范围"
         Height          =   255
         Left            =   90
         TabIndex        =   19
         Top             =   900
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   225
         Left            =   2100
         TabIndex        =   18
         Top             =   930
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "财务备注:"
         Height          =   225
         Left            =   9600
         TabIndex        =   17
         Top             =   6660
         Width           =   1005
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "备  注"
         Height          =   285
         Left            =   150
         TabIndex        =   15
         Top             =   5370
         Width           =   795
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "合计人民币（大写）"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4170
         TabIndex        =   11
         Top             =   4890
         Width           =   2175
      End
      Begin VB.Label lblFxz 
         BackStyle       =   0  'Transparent
         Caption         =   "FXZ"
         Height          =   435
         Left            =   10200
         TabIndex        =   8
         Top             =   930
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblBt 
         Caption         =   "营销部报销单"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5970
         TabIndex        =   6
         Top             =   330
         Width           =   2955
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "区域"
         Height          =   255
         Left            =   8850
         TabIndex        =   5
         Top             =   1290
         Width           =   495
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "报销单编号"
         Height          =   225
         Left            =   5730
         TabIndex        =   4
         Top             =   1290
         Width           =   1215
      End
      Begin VB.Label lblBh 
         BackStyle       =   0  'Transparent
         Caption         =   "11111"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7050
         TabIndex        =   3
         Top             =   1290
         Width           =   765
      End
      Begin VB.Label comQy 
         BackStyle       =   0  'Transparent
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9630
         TabIndex        =   2
         Top             =   1290
         Width           =   945
      End
      Begin VB.Label lblBM 
         BackStyle       =   0  'Transparent
         Caption         =   "lblBm"
         Height          =   255
         Left            =   3420
         TabIndex        =   1
         Top             =   1080
         Visible         =   0   'False
         Width           =   765
      End
   End
   Begin VB.Frame frmNewQ 
      Height          =   1695
      Left            =   30
      TabIndex        =   67
      Top             =   7410
      Width           =   7755
      Begin VB.Frame frmZQ 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   285
         Left            =   630
         TabIndex        =   109
         Top             =   120
         Width           =   6885
         Begin VB.CommandButton cmdFQ 
            BackColor       =   &H00C0FFC0&
            Height          =   255
            Left            =   1020
            Style           =   1  'Graphical
            TabIndex        =   111
            ToolTipText     =   "当涉及到出租车费，则要由总经理附加审核"
            Top             =   0
            Width           =   1125
         End
         Begin VB.Label lblFT 
            BackColor       =   &H00C0FFC0&
            Height          =   225
            Left            =   2160
            TabIndex        =   112
            Top             =   30
            Width           =   2115
         End
         Begin VB.Label Label27 
            Caption         =   "附加审核："
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   60
            TabIndex        =   110
            ToolTipText     =   "当涉及到出租车费，则要由总经理附加审核"
            Top             =   60
            Width           =   1035
         End
      End
      Begin VB.CommandButton cmdPje 
         Caption         =   "评审建议"
         Height          =   1095
         Left            =   150
         TabIndex        =   80
         Top             =   480
         Width           =   345
      End
      Begin VB.CommandButton cmdQm 
         Caption         =   "cmdQm"
         Height          =   345
         Index           =   0
         Left            =   600
         TabIndex        =   68
         Top             =   690
         Width           =   945
      End
      Begin VB.Label lblQM 
         Caption         =   "lblQm"
         Height          =   225
         Index           =   0
         Left            =   690
         TabIndex        =   70
         Top             =   420
         Width           =   915
      End
      Begin VB.Label lblTm 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   405
         Index           =   0
         Left            =   600
         TabIndex        =   69
         Top             =   1110
         Width           =   945
      End
   End
   Begin VB.Label Label21 
      Caption         =   "adoFyBound"
      DataField       =   "UserId"
      DataSource      =   "adoFy"
      Height          =   315
      Left            =   12030
      TabIndex        =   72
      Top             =   8040
      Visible         =   0   'False
      Width           =   1065
   End
End
Attribute VB_Name = "frmFYBX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public F2 As Object
Public DelF As Boolean '是否删除报销单
Public KeBao As Boolean '可否报销签收
Public Kd As Boolean '是否为初次开单
Dim JlP As String '运费经理密码
Public Fmx As Object '费用明细（新版）
Dim tQy As String
Dim Tbm As String
Dim aY As Object
Dim timZm As Integer '数据提交后,由timWait执行的后续命令ID(1费用编辑2签字3附加签字 5签收)
Dim QF As Boolean '签名方式，正常签还是附加签
Public Sub QMBound(Bxid As Long)
Dim Ra: Dim La
Dim ii As Integer: Dim oo As Integer
Dim tt As String
On Error Resume Next

tt = "select trq,ywy,zn,bz,tf from pizu where bh='" & Bxid & "' and yid=" & Val(lblNlb.Caption) & " order by pid desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2): dtgP.Rows = La + 20
Call dtgPFF
For oo = 1 To La + 1
    dtgP.Row = oo
    For ii = 0 To 5
        dtgP.Col = ii
        dtgP.Text = Ra(ii, oo - 1)
        If ii = 3 Then
            If Len(Ra(ii, oo - 1)) > 16 Then
                dtgP.RowHeight(oo) = UpInt(Len(Ra(ii, oo - 1)) / 16) * dtgP.RowHeight(oo)
            End If
        End If
        If ii = 4 Then
            If dtgP.Text = "True" Then
                dtgP.Text = "同意"
            ElseIf dtgP.Text = "False" Then
                dtgP.Text = "驳回"
            End If

        End If
    Next
Next
For oo = 1 To La + 1
    dtgP.Row = oo
    dtgP.Col = 4
            If dtgP.Text = "驳回" Then
                For ii = 0 To 5
                    dtgP.Col = ii
                    dtgP.CellForeColor = &HFF&
                Next
            End If
Next
dtgP.Row = 0
dtgP.Col = 0: dtgP.Text = "日期": dtgP.Col = 1: dtgP.Text = "姓名": dtgP.Col = 2: dtgP.Text = "职能"
dtgP.Col = 3: dtgP.Text = "评审建议": dtgP.Col = 4: dtgP.Text = "通过否"



End Sub
Public Sub dtgPFF()
Dim oo As Integer
For oo = 1 To dtgP.Rows - 1
    dtgP.RowHeight(oo) = dtgP.RowHeight(0)
Next
dtgP.Clear
dtgP.Row = 0
dtgP.Col = 0: dtgP.Text = "日期": dtgP.Col = 1: dtgP.Text = "姓名": dtgP.Col = 2: dtgP.Text = "职能": dtgP.Col = 3: dtgP.Text = "评审建议": dtgP.Col = 4: dtgP.Text = "审核":
dtgP.ColWidth(3) = 3000: dtgP.ColWidth(0) = 2000: dtgP.ColWidth(4) = 800
For oo = 0 To 4
    dtgP.Col = oo
    dtgP.CellFontBold = True
Next
End Sub

Private Sub cmdAdd_Click()
On Error Resume Next
Dim Orq As Date
Dim Oqy As String '用于添加费用归属明细
Dim OBm As String '用于添加费用归属明细
Dim Odep As String
    '如果空记录,则不能添加
    adoF2.Recordset.MoveLast
    If (adoF2.Recordset.Fields("XG").Value = 0 Or IsNull(adoF2.Recordset.Fields("XG").Value) = True) And _
        adoF2.Recordset.RecordCount > 0 And IsNull(adoF2.Recordset.Fields("atime").Value) = True Then
        Exit Sub
    End If
'Odep = adoF2.Recordset.Fields("dep").Value
'如果为运费或外地工程部报销单,则明细要对应业务员
 If lblNlb.Caption = 11 Or lblNlb.Caption = 9 Then '如果为工程外地或运费
    Orq = adoF2.Recordset("atime").Value
    adoF2.Recordset.AddNew "bxId", lblBh.Caption
    adoF2.Recordset.Update "atime", Orq
    adoF2.Recordset.Update "XG", 0
    adoF2.Recordset.Update "GongF", 2
    'adoF2.Recordset.Update "ITM", adoF2.Recordset.RecordCount
    Set dtgBx.DataSource = adoF2
    
    comhtBh.Text = ""
ElseIf lblNlb.Caption = 35 Then '如果为福利报销单
'    If lblGui.Caption = "" Then
'        MsgBox "请先选定费用归属人!"
'        Call cmdGui_Click
'        Exit Sub
'    End If
    Oqy = adoF2.Recordset.Fields("qy").Value
    OBm = adoF2.Recordset.Fields("bm").Value
'    Odep = adoF2.Recordset.Fields("dep").Value
    adoF2.Recordset.AddNew "BM", OBm
    adoF2.Recordset.Update "QY", Oqy
    'adoF2.Recordset.Update "dep", Odep
    adoF2.Recordset.Update "ywy", lblGui.Caption
    adoF2.Recordset.Update "ywyUid", lblGuid.Caption
    adoF2.Recordset.Update "bxId", lblBh.Caption
    adoF2.Recordset.Update "XG", 0
    adoF2.Recordset.Update "GongF", 2
    'adoF2.Recordset.Update "ITM", adoF2.Recordset.RecordCount

    Set dtgBx.DataSource = adoF2
ElseIf lblNlb.Caption = 84 Then '培训费归徐薇
'    Odep = adoF2.Recordset.Fields("dep").Value
    adoF2.Recordset.AddNew "BM", "行政人事"
    adoF2.Recordset.Update "QY", "上海"
    'adoF2.Recordset.Update "dep", Odep
    adoF2.Recordset.Update "ywy", "吴之禺"
    adoF2.Recordset.Update "ywyUid", "HM025"
    adoF2.Recordset.Update "bxId", lblBh.Caption
    adoF2.Recordset.Update "XG", 0
    adoF2.Recordset.Update "GongF", 1
    'adoF2.Recordset.Update "ITM", adoF2.Recordset.RecordCount

    Set dtgBx.DataSource = adoF2
Else

    adoF2.Recordset.AddNew "BM", lblBM.Caption
    'adoF2.Recordset.Update "dep", Odep
    adoF2.Recordset.Update "QY", comQy.Caption
    adoF2.Recordset.Update "ywy", mod1.DName
    adoF2.Recordset.Update "ywyUid", mod1.DHid
    adoF2.Recordset.Update "bxId", lblBh.Caption
    'adoF2.Recordset.Update "ITM", adoF2.Recordset.RecordCount
    adoF2.Recordset.Update "XG", 0
    adoF2.Recordset.Update "GongF", 2
    Set dtgBx.DataSource = adoF2


End If
    txtHg.Text = ""
    lblDx.Caption = ""
End Sub

Private Sub cmdBack_Click()
Dim tt As String
On Error Resume Next




Call mod1.DelDKZ ' '退出表单时删除打开记录,以让别人能打开此单据
frmFYBX.Visible = False
If frmBxBrow.Visible = True Then
    frmBxBrow.Enabled = True
    frmBxBrow.ZOrder 0
    'frmBxBrow.WindowState = 2
ElseIf Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
ElseIf fyBB.Visible = True Then
    fyBB.Enabled = True
    fyBB.ZOrder 0
ElseIf frmCWBBA.Visible = True Then
    frmCWBBA.Enabled = True
    frmCWBBA.ZOrder 0
End If
End Sub

Private Sub cmdCopy_Click()
Clipboard.Clear
Clipboard.SetText dtgNx.Clip

frmFYBX.dtgNx.FixedRows = 0
frmFYBX.dtgNx.MergeCol(1) = True
frmFYBX.dtgNx.MergeCol(2) = True
frmFYBX.dtgNx.MergeCol(41) = True
frmFYBX.dtgNx.MergeCol(42) = True
frmFYBX.dtgNx.MergeCol(43) = True
frmFYBX.dtgNx.MergeCells = 3
frmFYBX.dtgNx.FixedRows = 1
End Sub

Private Sub cmdDao_Click()
Dim tt As String
On Error Resume Next

tt = InputBox("请输入所参照的报销单编号！")
If Val(tt) = 0 Then
    Exit Sub
End If
Set mod1.cmd = CreateObject("adodb.command")
mod1.cmd.ActiveConnection = mod1.cc
mod1.cmd.CommandText = "FydDao"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@lb") = lblNlb.Caption '单子(报销单)种类
mod1.cmd.Parameters("@bxid") = lblBh.Caption
mod1.cmd.Parameters("@qy") = mod1.Qy
If Left(lblBt.Caption, 2) = "三金" Then
    mod1.cmd.Parameters("@dlb") = 1
ElseIf Left(lblBt.Caption, 2) = "公积" Then
    mod1.cmd.Parameters("@dlb") = 2
ElseIf Left(lblBt.Caption, 2) = "福利" Then
    mod1.cmd.Parameters("@dlb") = 3
ElseIf Left(lblBt.Caption, 2) = "外来" Then
    mod1.cmd.Parameters("@dlb") = 4
End If
mod1.cmd.Parameters("@oxid") = Val(tt)
mod1.cmd.Execute
Set cmd = Nothing

        '打开费用总表
    tt = "FydMxOpen(" & Val(lblBh.Caption) & ")"
    frmFYBX.adoF2.Recordset.Close
    frmFYBX.adoF2.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
    Set frmFYBX.dtgBx.DataSource = frmFYBX.adoF2
End Sub

Private Sub cmdDe_Click()

End Sub

Private Sub cmdDel_Click()
On Error Resume Next
Dim Oqy As String '用于添加费用归属明细
Dim OBm As String '用于添加费用归属明细

    Oqy = adoF2.Recordset.Fields("qy").Value
    OBm = adoF2.Recordset.Fields("bm").Value
    adoF2.Recordset.Delete adAffectCurrent



If adoF2.Recordset.RecordCount = 0 Then
    If lblNlb.Caption = 11 Or lblNlb.Caption = 9 Then '如果为工程外地或运费
        adoF2.Recordset.AddNew "bxId", lblBh.Caption
        adoF2.Recordset.Update "XG", 0
        Set dtgYf.DataSource = adoF2
    ElseIf lblNlb.Caption = 32 Then '如果为费用归属
        adoF2.Recordset.AddNew "BM", OBm
        adoF2.Recordset.Update "qy", Oqy
        adoF2.Recordset.Update "ywy", lblGui.Caption
        adoF2.Recordset.Update "ywyUid", lblGuid.Caption
        adoF2.Recordset.Update "bxId", lblBh.Caption
        adoF2.Recordset.Update "XG", 0
        Set dtgBx.DataSource = adoF2
    Else
        adoF2.Recordset.AddNew "BM", lblBM.Caption
        adoF2.Recordset.Update "qy", comQy.Caption
        adoF2.Recordset.Update "ywy", mod1.DName
        adoF2.Recordset.Update "ywyUid", mod1.DHid
        adoF2.Recordset.Update "bxId", lblBh.Caption
        adoF2.Recordset.Update "XG", 0
        Set dtgBx.DataSource = adoF2
    End If
End If
       ' Set dtgBx.DataSource = adoF2
txtHg.Text = ""
lblDx.Caption = ""
End Sub

Private Sub cmdDing_Click()
Dim tt As String
On Error Resume Next

If optT2.Value = True And txtQM.Text = "" Then
    MsgBox ("请您一定要告诉拒绝我的理由!  :) ")
    Exit Sub
End If
frmEd.Visible = False
If QF = False Then
        timZm = 2 '签字
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "报销单"
        mod1.cmd.Parameters("@NBLX") = "签字"
        mod1.cmd.Parameters("@bh") = lblBh.Caption
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = lblGui.Caption '归属人姓名（如果与报销人为同一人，则签字可以跳步）
        mod1.cmd.Parameters("@mt2") = lblGuid.Caption
        mod1.cmd.Parameters("@mt3") = lblNlb.Caption '报销单子种类（新费用归属为79)
        mod1.cmd.Parameters("@mt4") = lblYwy.Caption
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
        If mod1.Bm = "商务部" And lblLc.Caption = 5 Then
            mod1.cmd.Parameters("@mt20").Value = "财务审核"
        Else
        mod1.cmd.Parameters("@mt20").Value = lblQM(Val(lblLc.Caption) - 1).Caption
        End If
        If mod1.cmd.Parameters("@mt20").Value <> "财务审核" Then
            mod1.cmd.Parameters("@mt21") = lblQM(Val(lblLc.Caption)).Caption
        Else
            mod1.cmd.Parameters("@mt21").Value = "可以签收"
        End If
        mod1.cmd.Parameters("@mt22") = ""
        mod1.cmd.Parameters("@mt23") = ""
        mod1.cmd.Parameters("@mt24") = ""
        mod1.cmd.Parameters("@mt25") = ""
        mod1.cmd.Parameters("@mlt1") = txtQM.Text '评审建议
        mod1.cmd.Parameters("@mlt2") = ""
        mod1.cmd.Parameters("@mlt3") = ""
        mod1.cmd.Parameters("@mlt4") = ""
        mod1.cmd.Parameters("@mlt5") = ""
        mod1.cmd.Parameters("@mm1").Value = Val(lblLc.Caption)
        mod1.cmd.Parameters("@mm2").Value = Val(lblFwid.Caption)
        mod1.cmd.Parameters("@mm3") = 0
        mod1.cmd.Parameters("@mm4") = 0
        mod1.cmd.Parameters("@mm5") = 0
        mod1.cmd.Parameters("@mm6") = 0
        mod1.cmd.Parameters("@mm7") = 0
        mod1.cmd.Parameters("@mm8") = 0
        mod1.cmd.Parameters("@mm9") = 0
        mod1.cmd.Parameters("@mm10").Value = Val(txtHg.Text)
        mod1.cmd.Parameters("@mm11") = 0
        mod1.cmd.Parameters("@mm12") = 0
        mod1.cmd.Parameters("@mm13") = 0
        mod1.cmd.Parameters("@mm14") = 0
        mod1.cmd.Parameters("@mm15") = 0
        mod1.cmd.Parameters("@mm16") = 0
        mod1.cmd.Parameters("@mm17") = 0
        mod1.cmd.Parameters("@mm18") = 0
        mod1.cmd.Parameters("@mm19") = 0
        mod1.cmd.Parameters("@mm20") = 0
        If OptT1.Value = True Then
            mod1.cmd.Parameters("@mb1") = 1 '同意
        Else
            mod1.cmd.Parameters("@mb1") = 0 '拒绝
        End If
        mod1.cmd.Parameters("@mb2") = 0
        mod1.cmd.Parameters("@mb3") = 0
        mod1.cmd.Parameters("@mb4") = 0
        If frmZQ.Visible = False Then
            mod1.cmd.Parameters("@mb5") = 0
        Else
            mod1.cmd.Parameters("@mb5") = 1 '需要附加签字
        End If
        mod1.cmd.Parameters("@md1") = Null
        mod1.cmd.Parameters("@md2") = Null
        mod1.cmd.Parameters("@md3") = Null
        mod1.cmd.Parameters("@md4") = Null
        mod1.cmd.Parameters("@md5") = Null
    Else
            timZm = 3 '签字
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "报销单"
        mod1.cmd.Parameters("@NBLX") = "附加签字"
        mod1.cmd.Parameters("@bh") = lblBh.Caption
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = lblGui.Caption '归属人姓名（如果与报销人为同一人，则签字可以跳步）
        mod1.cmd.Parameters("@mt2") = lblGuid.Caption
        mod1.cmd.Parameters("@mt3") = lblNlb.Caption '报销单子种类（新费用归属为79)
        mod1.cmd.Parameters("@mt4") = lblYwy.Caption
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
        mod1.cmd.Parameters("@mt20").Value = ""
'        If lblQM(Val(lblLc.Caption) - 1).Caption <> "财务审核" Then
'            mod1.cmd.Parameters("@mt21") = lblQM(Val(lblLc.Caption)).Caption
'        Else
            mod1.cmd.Parameters("@mt21").Value = "财务审核"
'        End If
        mod1.cmd.Parameters("@mt22") = ""
        mod1.cmd.Parameters("@mt23") = ""
        mod1.cmd.Parameters("@mt24") = ""
        mod1.cmd.Parameters("@mt25") = ""
        mod1.cmd.Parameters("@mlt1") = txtQM.Text '评审建议
        mod1.cmd.Parameters("@mlt2") = ""
        mod1.cmd.Parameters("@mlt3") = ""
        mod1.cmd.Parameters("@mlt4") = ""
        mod1.cmd.Parameters("@mlt5") = ""
        mod1.cmd.Parameters("@mm1").Value = 5
        mod1.cmd.Parameters("@mm2").Value = Val(lblFwid.Caption)
        mod1.cmd.Parameters("@mm3") = 0
        mod1.cmd.Parameters("@mm4") = 0
        mod1.cmd.Parameters("@mm5") = 0
        mod1.cmd.Parameters("@mm6") = 0
        mod1.cmd.Parameters("@mm7") = 0
        mod1.cmd.Parameters("@mm8") = 0
        mod1.cmd.Parameters("@mm9") = 0
        mod1.cmd.Parameters("@mm10").Value = Val(txtHg.Text)
        mod1.cmd.Parameters("@mm11") = 0
        mod1.cmd.Parameters("@mm12") = 0
        mod1.cmd.Parameters("@mm13") = 0
        mod1.cmd.Parameters("@mm14") = 0
        mod1.cmd.Parameters("@mm15") = 0
        mod1.cmd.Parameters("@mm16") = 0
        mod1.cmd.Parameters("@mm17") = 0
        mod1.cmd.Parameters("@mm18") = 0
        mod1.cmd.Parameters("@mm19") = 0
        mod1.cmd.Parameters("@mm20") = 0
        If OptT1.Value = True Then
            mod1.cmd.Parameters("@mb1") = 1 '同意
        Else
            mod1.cmd.Parameters("@mb1") = 0 '拒绝
        End If
        mod1.cmd.Parameters("@mb2") = 0
        mod1.cmd.Parameters("@mb3") = 0
        mod1.cmd.Parameters("@mb4") = 0
        mod1.cmd.Parameters("@mb5") = 0
        mod1.cmd.Parameters("@md1") = Null
        mod1.cmd.Parameters("@md2") = Null
        mod1.cmd.Parameters("@md3") = Null
        mod1.cmd.Parameters("@md4") = Null
        mod1.cmd.Parameters("@md5") = Null

    End If
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        cmdDing.Enabled = False
        
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
End Sub

Private Sub cmdFQ_Click()
Dim TMlc As Integer '将TMX置为0时出错的次数.
Dim tt As String
Dim oo As Integer
Dim Tywy As String '单子流转到下一人的姓名
Dim Tuid As String
Dim Oywy As String '原来流转人的名字
Dim Ouid As String '原来流转人的工号
Dim ii As Integer
Dim Je As Double
On Error Resume Next

If cmdQm(Val(lblLc.Caption) - 2).Caption = "" Then
    Exit Sub
End If
If cmdQm(2).Caption = "" Then
    Exit Sub
End If
If cmdFQ.Caption <> "" Then Exit Sub


If mod1.DName <> "倪旭" And mod1.DName <> "彭海翔" Then

    Exit Sub
End If

If cmdFQ.Visible = True And Val(lblLc.Caption) < 3 Then
    Exit Sub
End If

If Index = 0 And cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If


    QF = True
    frmNQ.Visible = True
    cmdDing.Enabled = True
    optT2.Enabled = True
    OptT1.Enabled = True
    OptT1.Value = True
    Exit Sub

End Sub


Private Sub cmdG_Click()
Set Ren.XForm = New frmFYBX
Call mod1.RenXz("frmFYBX", Me, 0)
End Sub

Private Sub cmdGui_Click()
Dim ii As Integer
Dim oo As Integer
'dtgNx.Col = 48
'dtgNx.Row = 1
''For oo = 1 To dtgNx.Rows - 1
''
''Next
If mod1.FYF = True Then
    ii = MsgBox("是否费用归自己？", vbQuestion + vbYesNo + vbDefaultButton2, "Hello!")
    If ii = vbYes And mod1.Bm1 = "" And mod1.Bm2 = "" And mod1.Bm3 = "" Then
        lblGui.Caption = mod1.DName
        lblGuid.Caption = mod1.DHid
        lblBM.Caption = mod1.Bm
        Exit Sub
    End If
    If ii = vbYes And Not (mod1.Bm1 = "" And mod1.Bm2 = "" And mod1.Bm3 = "") Then
        MsgBox ("您有分管部门,所以请明示您的具体归属部门!")
    End If
End If
Set Ren.XForm = New frmFYBX
Call mod1.RenXz("frmFYBX", Me, 0)

End Sub

Private Sub cmdHg_Click()
Dim Je As Double
Dim oo As Integer
On Error Resume Next

Je = 0
adoF2.Recordset.MoveFirst
oo = 1
Do While Not adoF2.Recordset.EOF
    Je = Je + adoF2.Recordset.Fields("XG").Value
    adoF2.Recordset.Update "ITM", oo
    adoF2.Recordset.MoveNext
    oo = oo + 1
Loop
    txtHg.Text = Round(Je, 2)

    lblDx.Caption = mod1.ChangBi(Val(txtHg.Text))
End Sub

Private Sub cmdJadd_Click()
Dim tt As String
Dim oo As Integer
Dim TF As Boolean '费用类别是否正确
Dim Ltext As String
Dim Lb As String '费用类别代码
Dim CZF As Boolean '有无出租车
If Val(txtJe.Text) = 0 Then
    MsgBox ("请输入金额！")
    txtJe.SetFocus
    Exit Sub
End If
If txtNr.Text = "" Then
    MsgBox ("请输入费用内容！")
    txtNr.SetFocus
    Exit Sub
End If

If lblGui.Caption <> "" And txtBm.Text <> "" And lblGui.Caption <> txtBm.Text And dtgNx.Rows > 2 Then
    MsgBox ("固定费用只能归同一个部门！")
    Exit Sub
End If

If opt1.Value = True And lblGui.Caption = "" Then
    MsgBox ("请选择所归属人！")
    Exit Sub
End If

If opt1.Value = False And opt2.Value = False Then
    MsgBox ("请确认费用是属于固定费用还是变动费用！")
    Exit Sub
End If

If opt2.Value = True And txtBm.Text = "" Then
    MsgBox ("请选择所归属的部门！")
    Exit Sub
End If

'工程部人员要填写工作单编号
If txtGZDH.Text = "" And mod1.Bm = "工程部" Then
    MsgBox ("请填写工作单编号！")
    Exit Sub
End If

'检测费用类别正确性
If comLb.ListIndex = 0 Then

Else
    TF = False
    Ltext = comLb.Text
'    comLb.ListIndex = 0
'    For oo = 0 To comLb.ListCount - 1
'        If comLb.ListIndex = 31 Then
'            comLb.ListIndex = 0
'        End If
'        comLb.ListIndex = comLb.ListIndex + 1
'
'        If Ltext = comLb.Text Then
'            TF = True
'            Exit For
'        End If
'    Next
If comLb.Text = "房屋补贴" Or comLb.Text = "旅游费" Or comLb.Text = "福利" Or comLb.Text = "高温费" Or comLb.Text = "通信费" Or comLb.Text = "市内交通费" Or _
   comLb.Text = "市外交通费" Or comLb.Text = "停车费" Or comLb.Text = "车辆费" Or comLb.Text = "运费" Or comLb.Text = "住宿费" Or comLb.Text = "部门团队费" Or _
   comLb.Text = "餐费" Or comLb.Text = "招待费" Or comLb.Text = "礼品费" Or comLb.Text = "房租" Or comLb.Text = "物业费" Or comLb.Text = "水电" Or _
   comLb.Text = "电话" Or comLb.Text = "办公用品" Or comLb.Text = "邮资" Or comLb.Text = "市场推广" Or comLb.Text = "人员招聘" Or comLb.Text = "快递费" Or _
   comLb.Text = "培训费" Or comLb.Text = "财务手续费" Or comLb.Text = "团队建设费" Or comLb.Text = "其它" Or comLb.Text = "公共停车费" Or _
   comLb.Text = "公共车辆费" Or comLb.Text = "工具" Or comLb.Text = "易耗" Or comLb.Text = "外劳" Then
    TF = True
End If
    If TF = False Then
        MsgBox ("费用类别不正确！")
        Exit Sub
    End If
End If

Select Case comLb.Text
Case "房屋补贴"
    Lb = "FWBT"
Case "旅游费"
    Lb = "LYF"
Case "福利"
    Lb = "SJ"
Case "高温费"
    Lb = "GWF"
Case "通信费"
    Lb = "TXF"
Case "市内交通费"
    Lb = "NJTF"
Case "市外交通费"
    Lb = "WJTF"
Case "停车费"
    Lb = "TCF"
Case "车辆费"
    Lb = "CLF"
Case "运费"
    Lb = "YF"
Case "住宿费"
    Lb = "ZCF"
Case "部门团队费"
    Lb = "BMTD"
Case "餐费"
    Lb = "CF"
Case "招待费"
    Lb = "ZDF"
Case "礼品费"
    Lb = "LPF"
Case "房租"
    Lb = "FZ"
Case "物业费"
    Lb = "WYF"
Case "水电"
    Lb = "SD"
Case "电话"
    Lb = "DW"
Case "办公用品"
    Lb = "BGYP"
Case "邮资"
    Lb = "YZ"
Case "市场推广"
    Lb = "SZTG"
Case "人员招聘"
    Lb = "RYZP"
Case "快递费"
    Lb = "KDF"
Case "培训费"
    Lb = "PXF"
Case "财务手续费"
    Lb = "CWSX"
Case "团队建设费"
    Lb = "TDJS"
Case "其它"
    Lb = "QTF"
Case "公共停车费"
    Lb = "GTCF"
Case "公共车辆费"
    Lb = "GCLF"
Case "工具"
    Lb = "GG"
Case "易耗"
    Lb = "yH"
Case "外劳"
    Lb = "wl"
End Select


''检查有无出租车
'dtgNx.Col = 45
'CZF = False
'For oo = 1 To dtgNx.Rows
'    If dtgNx.Text <> "" Then
'        CZF = True
'        Exit For
'    End If
'Next

timZm = 1 '费用编辑
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "报销单"
    mod1.cmd.Parameters("@NBLX") = "费用编辑"
    mod1.cmd.Parameters("@bh") = lblBh.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Lb '费用类别
    mod1.cmd.Parameters("@mt2") = Left(txtNr.Text, 30) '报销内容
    mod1.cmd.Parameters("@mt3") = lblGui.Caption
    mod1.cmd.Parameters("@mt4") = lblGuid.Caption
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
    mod1.cmd.Parameters("@mt8") = ""
    mod1.cmd.Parameters("@mt9") = ""
    If txtBm.Text = "" Then
        mod1.cmd.Parameters("@mt10") = lblBM.Caption
    Else
        mod1.cmd.Parameters("@mt10") = txtBm.Text
    End If
    mod1.cmd.Parameters("@mt11") = ""
    mod1.cmd.Parameters("@mt12") = ""
    mod1.cmd.Parameters("@mt13") = ""
    mod1.cmd.Parameters("@mt14") = ""
    mod1.cmd.Parameters("@mt15") = ""
    mod1.cmd.Parameters("@mt16") = txtGZDH.Text
    mod1.cmd.Parameters("@mt17") = ""
    mod1.cmd.Parameters("@mt18") = ""
    mod1.cmd.Parameters("@mt19") = ""
    mod1.cmd.Parameters("@mt20") = ""
    mod1.cmd.Parameters("@mt21") = ""
    mod1.cmd.Parameters("@mt22") = ""
    mod1.cmd.Parameters("@mt23") = ""
    mod1.cmd.Parameters("@mt24") = ""
    mod1.cmd.Parameters("@mt25") = ""
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mlt2") = ""
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtJe.Text) '金额
    mod1.cmd.Parameters("@mm2") = 0
    mod1.cmd.Parameters("@mm3") = 0
    mod1.cmd.Parameters("@mm4") = 0
    mod1.cmd.Parameters("@mm5") = 0
    mod1.cmd.Parameters("@mm6") = 0
    mod1.cmd.Parameters("@mm7") = 0
    mod1.cmd.Parameters("@mm8") = 0
    mod1.cmd.Parameters("@mm9") = 0
    mod1.cmd.Parameters("@mm10") = 1 '添加费用
    If opt1.Value = True Then          '公共费用否
        mod1.cmd.Parameters("@mm11") = 2
    ElseIf opt2.Value = True Then
        mod1.cmd.Parameters("@mm11") = 1
    End If
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
    mod1.cmd.Parameters("@md1") = dtPRQ.Value '日期
    mod1.cmd.Parameters("@md2") = Null
    mod1.cmd.Parameters("@md3") = Null
    mod1.cmd.Parameters("@md4") = Null
    mod1.cmd.Parameters("@md5") = Null
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
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

End Sub

Private Sub cmdJdel_Click()
Dim ii As Integer
Dim tt As String
Dim oo As Integer
Dim TF As Boolean '费用类别是否正确
Dim Ltext As String
Dim CZF As Boolean
If Val(lblBid.Caption) = 0 Then
    Exit Sub
End If

ii = MsgBox("是否确认删除这笔费用？", vbQuestion + vbYesNo)
If ii = vbNo Then
    Exit Sub
End If

''检查有无出租车
'dtgNx.Col = 45
'CZF = False
'For oo = 1 To dtgNx.Rows
'    If dtgNx.Text <> "" Then
'        CZF = True
'        Exit For
'    End If
'Next
timZm = 1 '费用编辑
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "报销单"
    mod1.cmd.Parameters("@NBLX") = "费用编辑"
    mod1.cmd.Parameters("@bh") = lblBh.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = comLb.Text '费用类别
    mod1.cmd.Parameters("@mt2") = txtNr.Text '报销内容
    mod1.cmd.Parameters("@mt3") = Val(lblBid.Caption)
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
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mlt2") = ""
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtJe.Text) '金额
    mod1.cmd.Parameters("@mm2") = 0
    mod1.cmd.Parameters("@mm3") = 0
    mod1.cmd.Parameters("@mm4") = 0
    mod1.cmd.Parameters("@mm5") = 0
    mod1.cmd.Parameters("@mm6") = 0
    mod1.cmd.Parameters("@mm7") = 0
    mod1.cmd.Parameters("@mm8") = 0
    mod1.cmd.Parameters("@mm9") = 0
    mod1.cmd.Parameters("@mm10") = 3 '删除费用
    mod1.cmd.Parameters("@mm11") = 0
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
    mod1.cmd.Parameters("@md1") = dtPRQ.Value '日期
    mod1.cmd.Parameters("@md2") = Null
    mod1.cmd.Parameters("@md3") = Null
    mod1.cmd.Parameters("@md4") = Null
    mod1.cmd.Parameters("@md5") = Null
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
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
End Sub

Private Sub cmdJed_Click()
Dim tt As String
Dim oo As Integer
Dim TF As Boolean '费用类别是否正确
Dim Ltext As String
Dim CZF As Boolean
If Val(txtJe.Text) = 0 Then
    MsgBox ("请输入金额！")
    txtJe.SetFocus
    Exit Sub
End If

If txtNr.Text = "" Then
    MsgBox ("请输入费用内容！")
    txtNr.SetFocus
    Exit Sub
End If

If opt1.Value = False And opt2.Value = False Then
    MsgBox ("请确认费用是属于固定费用还是变动费用！")
    Exit Sub
End If

If opt2.Value = True And txtBm.Text = "" Then
    MsgBox ("请选择所归属的部门！")
    Exit Sub
End If

'检测费用类别正确性
If comLb.ListIndex = 0 Then

Else
    TF = False
    Ltext = comLb.Text
'    comLb.ListIndex = 0
'    For oo = 0 To comLb.ListCount - 1
'        If comLb.ListIndex = 31 Then
'            comLb.ListIndex = 0
'        End If
'        comLb.ListIndex = comLb.ListIndex + 1
'
'        If Ltext = comLb.Text Then
'            TF = True
'            Exit For
'        End If
'    Next
If comLb.Text = "房屋补贴" Or comLb.Text = "旅游费" Or comLb.Text = "福利" Or comLb.Text = "高温费" Or comLb.Text = "通信费" Or comLb.Text = "市内交通费" Or _
   comLb.Text = "市外交通费" Or comLb.Text = "停车费" Or comLb.Text = "车辆费" Or comLb.Text = "运费" Or comLb.Text = "住宿费" Or comLb.Text = "部门团队费" Or _
   comLb.Text = "餐费" Or comLb.Text = "招待费" Or comLb.Text = "礼品费" Or comLb.Text = "房租" Or comLb.Text = "物业费" Or comLb.Text = "水电" Or _
   comLb.Text = "电话" Or comLb.Text = "办公用品" Or comLb.Text = "邮资" Or comLb.Text = "市场推广" Or comLb.Text = "人员招聘" Or comLb.Text = "快递费" Or _
   comLb.Text = "培训费" Or comLb.Text = "财务手续费" Or comLb.Text = "团队建设费" Or comLb.Text = "其它" Or comLb.Text = "公共停车费" Or _
   comLb.Text = "公共车辆费" Or comLb.Text = "工具" Or comLb.Text = "易耗" Or comLb.Text = "外劳" Then
    TF = True
End If
    If TF = False Then
        MsgBox ("费用类别不正确！")
        Exit Sub
    End If
End If
Select Case comLb.Text
Case "房屋补贴"
    Lb = "FWBT"
Case "旅游费"
    Lb = "LYF"
Case "福利"
    Lb = "SJ"
Case "高温费"
    Lb = "GWF"
Case "通信费"
    Lb = "TXF"
Case "市内交通费"
    Lb = "NJTF"
Case "市外交通费"
    Lb = "WJTF"
Case "停车费"
    Lb = "TCF"
Case "车辆费"
    Lb = "CLF"
Case "运费"
    Lb = "YF"
Case "住宿费"
    Lb = "ZCF"
Case "部门团队费"
    Lb = "BMTD"
Case "餐费"
    Lb = "CF"
Case "招待费"
    Lb = "ZDF"
Case "礼品费"
    Lb = "LPF"
Case "房租"
    Lb = "FZ"
Case "物业费"
    Lb = "WYF"
Case "水电"
    Lb = "SD"
Case "电话"
    Lb = "DW"
Case "办公用品"
    Lb = "BGYP"
Case "邮资"
    Lb = "YZ"
Case "市场推广"
    Lb = "SZTG"
Case "人员招聘"
    Lb = "RYZP"
Case "快递费"
    Lb = "KDF"
Case "培训费"
    Lb = "PXF"
Case "财务手续费"
    Lb = "CWSX"
Case "团队建设费"
    Lb = "TDJS"
Case "其它"
    Lb = "QTF"
Case "公共停车费"
    Lb = "GTCF"
Case "公共车辆费"
    Lb = "GCLF"
Case "工具"
    Lb = "GG"
Case "易耗"
    Lb = "yH"
Case "外劳"
    Lb = "wl"
End Select

''检查有无出租车
'dtgNx.Col = 45
'CZF = False
'For oo = 1 To dtgNx.Rows
'    If dtgNx.Text <> "" Then
'        CZF = True
'        Exit For
'    End If
'Next
timZm = 1 '费用编辑
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "报销单"
    mod1.cmd.Parameters("@NBLX") = "费用编辑"
    mod1.cmd.Parameters("@bh") = lblBh.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Lb '费用类别
    mod1.cmd.Parameters("@mt2") = Left(txtNr.Text, 30) '报销内容
    mod1.cmd.Parameters("@mt3") = lblBid.Caption
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
    mod1.cmd.Parameters("@mt8") = ""
    mod1.cmd.Parameters("@mt9") = ""
    mod1.cmd.Parameters("@mt10") = txtBm.Text
    mod1.cmd.Parameters("@mt11") = ""
    mod1.cmd.Parameters("@mt12") = ""
    mod1.cmd.Parameters("@mt13") = ""
    mod1.cmd.Parameters("@mt14") = ""
    mod1.cmd.Parameters("@mt15") = ""
    mod1.cmd.Parameters("@mt16") = txtGZDH.Text
    mod1.cmd.Parameters("@mt17") = ""
    mod1.cmd.Parameters("@mt18") = ""
    mod1.cmd.Parameters("@mt19") = ""
    mod1.cmd.Parameters("@mt20") = ""
    mod1.cmd.Parameters("@mt21") = ""
    mod1.cmd.Parameters("@mt22") = ""
    mod1.cmd.Parameters("@mt23") = ""
    mod1.cmd.Parameters("@mt24") = ""
    mod1.cmd.Parameters("@mt25") = ""
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mlt2") = ""
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtJe.Text) '金额
    mod1.cmd.Parameters("@mm2") = 0
    mod1.cmd.Parameters("@mm3") = 0
    mod1.cmd.Parameters("@mm4") = 0
    mod1.cmd.Parameters("@mm5") = 0
    mod1.cmd.Parameters("@mm6") = 0
    mod1.cmd.Parameters("@mm7") = 0
    mod1.cmd.Parameters("@mm8") = 0
    mod1.cmd.Parameters("@mm9") = 0
    mod1.cmd.Parameters("@mm10") = 2 '更新费用
    If opt1.Value = True Then          '公共费用否
        mod1.cmd.Parameters("@mm11") = 2
    ElseIf opt2.Value = True Then
        mod1.cmd.Parameters("@mm11") = 1
        lblGuid.Caption = ""
    End If
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
    mod1.cmd.Parameters("@md1") = dtPRQ.Value '日期
    mod1.cmd.Parameters("@md2") = Null
    mod1.cmd.Parameters("@md3") = Null
    mod1.cmd.Parameters("@md4") = Null
    mod1.cmd.Parameters("@md5") = Null
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
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
End Sub


Private Sub cmdKc_Click()
Dim Qy As String
Dim Bm As String
Dim Ywy As String
Dim Uid As String
On Error Resume Next
If comBm.Text <> "" Then
        Select Case comBm.Text
'            Case "维销部2"
'                Qy = "上海"
'                Bm = "维销部2"
'                Ywy = "肖卫国"
'                Uid = "HM041"
'            Case "维销部1"
'                Qy = "上海"
'                Bm = "维销部1"
'                Ywy = "孟智峰"
'                Uid = "HM002"
'            Case "维销部3"
'                Qy = "上海"
'                Bm = "维销部3"
'                Ywy = "孟智峰"
'                Uid = "HM002"
            Case "上海仓库"
                Qy = "上海"
                Bm = "总经理"
                Ywy = "宋晓炯"
                Uid = "HM003"
            Case "杭州办"
                Qy = "杭州办"
                Bm = "杭州办"
                Ywy = "颜继明"
                Uid = "HM104"
            Case "无锡办"
                Qy = "无锡"
                Bm = "无锡办"
                Ywy = "刘继楚"
                Uid = "HM063"
            Case "南京办"
                Qy = "南京"
                Bm = "南京办"
                Ywy = "南京办经理"
                Uid = "HM200"
            Case "北京办"
                Qy = "北京"
                Bm = "北京办"
                Ywy = "姜_威"
                Uid = "HM135"
        End Select
        adoF2.Recordset.Update ("khmc"), "豪曼"
        adoF2.Recordset.Update ("htBh"), "库存"
        adoF2.Recordset.Update ("ywy"), Ywy
        adoF2.Recordset.Update ("ywyUid"), Uid
        adoF2.Recordset.Update ("qy"), Qy
        adoF2.Recordset.Update ("BM"), Bm
End If

End Sub

Private Sub cmdMod_Click()
On Error Resume Next
If Val(lblLc.Caption) > 1 Or lblLcRen.Caption <> mod1.DName Then
    Exit Sub
End If
cmdAdd.Visible = True
cmdDel.Visible = True
dtgBx.AllowUpdate = True
frmFYBX.txtBz.Enabled = True
frmFYBX.txtBz.Locked = False

If lblNlb.Caption = 55 Or lblNlb.Caption = 56 Or lblNlb.Caption = 35 Or lblNlb.Caption = 59 Or lblNlb.Caption = 84 Then
    cmdG.Visible = True
End If

cmdSave.Enabled = True
'如果为运费报销单子,则打开归类表
If lblNlb.Caption = 9 Then
    frmYf.Visible = True

End If

'如果为工程外地报销单,则打开归类表
If lblNlb.Caption = 11 Or lblNlb.Caption = 12 Then

End If

'If Left(lblBt.Caption, 3) = "业务员" Then
'    txtJe.Locked = True
'    comLb.Enabled = False
'    dtpRq.Enabled = False
'    txtNr.Locked = True
'    frmED.Visible = True
'End If
If Val(lblLc.Caption) = 1 And lblLcUid.Caption = mod1.DHid And Left(lblBt.Caption, 4) = "内部结算" Then
    cmdG.Visible = True
End If
If Val(lblNlb.Caption) = 79 Then
    cmdAdd.Visible = False
    cmdDel.Visible = False
    frmEd.Visible = True
    cmdGui.Visible = True
End If
If Val(lblNlb.Caption) = 35 Or Left(lblBt.Caption, 3) = "业务员" Then
    dtgNx.Visible = False

        dtgBx.Visible = True
End If
End Sub

Private Sub cmdNQ_Click()
Dim TMlc As Integer '将TMX置为0时出错的次数.
Dim tt As String
Dim oo As Integer
Dim Tywy As String '单子流转到下一人的姓名
Dim Tuid As String
Dim Oywy As String '原来流转人的名字
Dim Ouid As String '原来流转人的工号
Dim ii As Integer
Dim Je As Double
Dim khmcT As String
Dim ywyT As String
On Error Resume Next





If mod1.Bm = "工程部" Then '工程部单子，一定要填写工作单编号

End If

'If lblLcUid.Caption <> mod1.DHid And lblQM(Index).Caption <> "业务审核" Then
If lblLcRen.Caption <> mod1.DName Then
    MsgBox "此处应由" & lblLcRen.Caption & "签字! 请您不要再点"
    Exit Sub
End If
If Val(lblLc.Caption) = 100 Then

        Exit Sub

End If
If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If

            '检查归属人员正确设置否
            If dtgNx.Visible = False Then
                For oo = 1 To adoF2.Recordset.RecordCount
                    khmcT = ""
                    ywyT = ""
                    dtgBx.Row = oo
                    dtgBx.Col = 1
                    khmcT = dtgBx.Text
                    dtgBx.Col = 42
                    ywyT = dtgBx.Text
                    If ywyT = "" And cmdQm(0).Caption <> "" Then
                        MsgBox "帮帮忙，帮我填好归属人员好吗？"
                        MsgBox "真累！"
                        Exit Sub
                    End If
                Next
            End If
If lblQM(Index).Caption = "财务审核" And (mod1.DName = "宋晓炯" Or mod1.DName = "宋晓炯1") Then
    Exit Sub
End If

            If optFp1.Value = False And optFp2.Value = False Then
                MsgBox "请确认发票情况!"
                cmdSave.Enabled = True
                Exit Sub
            End If
            
            If optFp2.Value = True And txtFP.Text = "" Then
                MsgBox "请注明发票不一致的原因!"
                cmdSave.Enabled = True
                Exit Sub
            End If

If Val(lblNewF.Caption) = 1 Or Val(lblNlb.Caption = 82) Then '新单签字方式和内部结算
    If lblGui.Caption = "" And Not (lblBt.Caption = "三金报销单" Or lblBt.Caption = "公积金报销单" Or _
    lblBt.Caption = "福利报销单" Or lblBt.Caption = "外来综合保险报销单" Or lblBt.Caption = "培训报销单") Then
        MsgBox ("请选择费用归属人！")
        cmdSave.Enabled = True
        Exit Sub
    End If
    QF = False
    frmNQ.Visible = True
    cmdDing.Enabled = True
    
    If Val(lblLc.Caption) = 1 Then  '报销人只能签字，不能驳回。
        optT2.Enabled = False
        
    Else
        optT2.Enabled = True
    End If
    Exit Sub
End If













If lblLc.Caption > 1 Then
    ii = MsgBox("您是否核准此单？(选择“是”将签字通过,选择“否”将驳回此单)", vbYesNoCancel + vbInformation, "请您注意!")
    If ii = vbNo Then
        ii = MsgBox("驳回后,此单将回转至填单人" & lblYwy.Caption & ",确认吗?", vbYesNo + vbInformation, "确认驳回吗?")
        If ii = vbNo Then
            Exit Sub
        End If
        tt = InputBox("请输入您要驳回的原因!")
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "xtzxFAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@yid").Value = lblNlb.Caption  '反签名
        mod1.cmd.Parameters("@lc").Value = lblLc.Caption
        mod1.cmd.Parameters("@bh").Value = lblBh.Caption
        mod1.cmd.Parameters("@ywy").Value = mod1.DName
        mod1.cmd.Parameters("@uid").Value = mod1.DHid
        mod1.cmd.Parameters("@bz").Value = tt
        mod1.cmd.Parameters("@zn").Value = lblQM(Index).Caption '身份职能
        mod1.cmd.Execute
        Set cmd = Nothing
        For oo = 0 To 6
            cmdQm(oo).Caption = ""
            lblTm(oo).Caption = ""
        Next
        lblLc.Caption = 999 '不让再按签名按钮.
        If Dialog.Visible = True Then '更新事务列表
            Call mod1.refEnvent(1)
        End If
        Exit Sub
    ElseIf ii = vbCancel Then
        Exit Sub
    End If
End If
Oywy = lblLcRen.Caption
Ouid = lblLcUid.Caption





'确认费用归属到最小的归属人单位
If cmdQm(Index).Caption = "组长" Or cmdQm(Index).Caption = "部门经理" Then

End If



If lblLc.Caption = 1 Then
Dim Zi As Integer
Zi = MsgBox("是否确认签字?", vbYesNo)
If Zi = vbNo Then Exit Sub
End If

'验证表头与表身的一致性.
If lblLc.Caption > 1 Then
    Je = 0
    adoF2.Recordset.MoveFirst
    oo = 1
    Do While Not adoF2.Recordset.EOF
        Je = Je + adoF2.Recordset.Fields("XG").Value
        adoF2.Recordset.MoveNext
        oo = oo + 1
    Loop
        If Val(txtHg.Text) <> Round(Je, 2) Then
            MsgBox "总金额与明细金额不一致,请退回此单!"
            Exit Sub
        End If
End If
If lblQM(Index).Caption = "业务审核" Then '如果是业务审核签字,则当前流程不变,直到全部签字后流程才改变.
    If lblYqf.Caption = 0 Then
        MsgBox "请在明细栏中审核与您相关的费用!"
        Exit Sub
    End If
ElseIf lblQM(Index).Caption = "财务审核" Then '如果是财务审核签字,则当前流程不变,直到签收后流程才改变.

Else
    lblLc.Caption = lblLc.Caption + 1
End If

    '设置服务器纠错系统
    tt = "update fyd set tmx=0 where bxid=" & Val(lblBh.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    TMlc = 0
CmdqmClick:
    TMlc = TMlc + 1
    If TMlc = 5 Then
        ii = MsgBox("网络出现严重故障,请稍候片刻再试!", vbExclamation, "C级警报")
        Exit Sub
    End If
    On Error GoTo CmdqmClick
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    On Error Resume Next

    If lblNlb.Caption <> 54 And lblNlb.Caption <> 70 And lblNlb.Caption <> 66 And lblNlb.Caption <> 67 And lblNlb.Caption <> 72 Then
    '更新表Fyd中的lcRen,lcUid 字段,以及QMRZ表中的相应字段.
                Set mod1.cmd = CreateObject("adodb.command")
                mod1.cmd.ActiveConnection = mod1.cc
                mod1.cmd.CommandText = "QMRZQM"
                mod1.cmd.CommandType = adCmdStoredProc
                mod1.cmd.Parameters("@nlb") = lblNlb.Caption '单子(报销单)种类
                mod1.cmd.Parameters("@lc") = lblLc.Caption  '当前流程
                mod1.cmd.Parameters("@Dname") = mod1.DName
                mod1.cmd.Parameters("@uid") = mod1.DHid
                mod1.cmd.Parameters("@btz") = mod1.BTZ '业务属性
                mod1.cmd.Parameters("@zid") = Index + 1 '流程顺序
                mod1.cmd.Parameters("@Qdbh") = lblBh.Caption '单子编号
                mod1.cmd.Parameters("@pje") = ""   '评审建议
                mod1.cmd.Parameters("@bm") = lblBM.Caption
                mod1.cmd.Parameters("@qy") = comQy.Caption
                mod1.cmd.Parameters("@Gren") = lblGui.Caption '如果为费用归属报销单,则添加费用归属人的参数
                mod1.cmd.Parameters("@Guid") = lblGuid.Caption
                mod1.cmd.Parameters("@ywy") = lblYwy.Caption
                mod1.cmd.Parameters("@yid") = lblUid.Caption
                mod1.cmd.Parameters("@comid") = mod1.comId
                mod1.cmd.Execute
                Tywy = mod1.cmd.Parameters("@Tywy").Value
                Tuid = mod1.cmd.Parameters("@Tuid").Value
                Set cmd = Nothing
                If (Tywy = "文静" And comQy.Caption <> "上海") Or (Tywy = "王蕾" And comQy.Caption = "南京") Then
                    If comQy.Caption = "南京" Then
                        Tywy = "王蕾"
                        Tuid = "HM051"
                    ElseIf comQy.Caption = "杭州" Then
                        Tywy = "李艳"
                        Tuid = "HM316"
                    ElseIf comQy.Caption = "北京" Then
                        Tywy = "马玉芝"
                        Tuid = "HM190"
                    ElseIf comQy.Caption = "广州" Then
                        Tywy = "汤丽嫦"
                        Tuid = "HMG023"
                    End If

                    tt = "update fyd set lcren='" & Tywy & "',lcUid='" & Tuid & "',lc=" & Val(lblLc.Caption) & " where bxid=" & lblBh.Caption
                    Set mod1.HTP = CreateObject("adodb.recordset")
                    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
                End If
                If Tywy = "宋晓炯" And mod1.comId = 1 Then
                    Tywy = "宋晓炯1"
                    Tuid = "HMG000"
                    tt = "update fyd set lc=" & Val(lblLc.Caption) & ",lcren='" & Tywy & "',lcuid='" & _
                        Tuid & "' where bxid=" & lblBh.Caption
                    Set mod1.HTP = CreateObject("adodb.recordset")
                    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
                End If
                If Tywy = "倪旭" And mod1.comId = 1 Then '如果广州的单子误跳到倪旭,则置为宋晓炯
                    Tywy = "宋晓炯1"
                    Tuid = "HMG000"
                    tt = "update fyd set lc=" & Val(lblLc.Caption) & ",lcren='" & Tywy & "',lcuid='" & _
                        Tuid & "' where bxid=" & lblBh.Caption
                    Set mod1.HTP = CreateObject("adodb.recordset")
                    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
                End If
                
                '如果为广州总经理报销,则审核人为彭海翔
                If Tywy = "周春云" And lblYwy.Caption = "宋晓炯1" And comQy.Caption = "广州" Then
                    Tywy = "彭海翔"
                    Tuid = "HMG002"
                    tt = "update fyd set lc=" & Val(lblLc.Caption) & ",lcren='" & Tywy & "',lcuid='" & _
                        Tuid & "' where bxid=" & lblBh.Caption
                    Set mod1.HTP = CreateObject("adodb.recordset")
                    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
                End If
                cmdQm(Index).Caption = mod1.DName
                lblTm(Index).Caption = mod1.DQda
                lblLcRen.Caption = Tywy
                lblLcUid.Caption = Tuid
        ElseIf lblNlb.Caption = 54 Or lblNlb.Caption = 70 Then '工程部报销单
                Set mod1.cmd = CreateObject("adodb.command")
                mod1.cmd.ActiveConnection = mod1.cc
                mod1.cmd.CommandText = "QMRZGC"
                mod1.cmd.CommandType = adCmdStoredProc
                mod1.cmd.Parameters("@nlb") = lblNlb.Caption '单子(报销单)种类
                mod1.cmd.Parameters("@lc") = lblLc.Caption  '当前流程
                mod1.cmd.Parameters("@Dname") = mod1.DName
                mod1.cmd.Parameters("@uid") = mod1.DHid
                mod1.cmd.Parameters("@btz") = mod1.BTZ '业务属性
                mod1.cmd.Parameters("@zid") = cmdQm(Index).Tag '流程顺序
                mod1.cmd.Parameters("@Qdbh") = lblBh.Caption '单子编号
                mod1.cmd.Parameters("@pje") = ""   '评审建议
                mod1.cmd.Parameters("@bm") = lblBM.Caption
                mod1.cmd.Parameters("@qy") = comQy.Caption
                mod1.cmd.Parameters("@Gren") = lblGui.Caption '如果为费用归属报销单,则添加费用归属人的参数
                mod1.cmd.Parameters("@Guid") = lblGuid.Caption
                mod1.cmd.Parameters("@ywy") = lblYwy.Caption
                mod1.cmd.Parameters("@yid") = lblUid.Caption
                
                mod1.cmd.Execute
                
                Tywy = mod1.cmd.Parameters("@Tywy").Value
                Tuid = mod1.cmd.Parameters("@Tuid").Value
                Set cmd = Nothing
                
                
                
                If (Tywy = "文静" And comQy.Caption <> "上海") Or (Tywy = "王蕾" And comQy.Caption = "南京") Or lblQM(Index + 1).Caption = "财务审核" Then
                    If comQy.Caption = "南京" Then
                        Tywy = "王蕾"
                        Tuid = "HM051"
                    ElseIf comQy.Caption = "杭州" Then
                        Tywy = "李艳"
                        Tuid = "HM316"
                    ElseIf comQy.Caption = "北京" Then
                        Tywy = "马玉芝"
                        Tuid = "HM190"
                    ElseIf comQy.Caption = "广州" Then
                        Tywy = "汤丽嫦"
                        Tuid = "HMG023"
                    End If
                '    tt = "update QMRZ set  Qren='" & mod1.DName & "',Qrid='" & mod1.DHid & "',Qrq='" & mod1.DQda & "' where Qdbh='" & txtHtbh.Text & "' and btz=" & mod1.BTZ & " and zid=" & cmdQm(Index).Tag
                '    Set mod1.HTP = CreateObject("adodb.recordset")
                '    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
                    tt = "update fyd set lcren='" & Tywy & "',lcUid='" & Tuid & "',lc=" & Val(lblLc.Caption) & " where bxid=" & lblBh.Caption
                    Set mod1.HTP = CreateObject("adodb.recordset")
                    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
                End If
                cmdQm(Index).Caption = mod1.DName
                lblTm(Index).Caption = mod1.DQda
                lblLcRen.Caption = Tywy
                lblLcUid.Caption = Tuid
        ElseIf lblNlb.Caption = 67 Or lblNlb.Caption = 66 Then '房屋补贴
                tt = "update QMRZ set  Qren='" & mod1.DName & "',Qrid='" & mod1.DHid & "',Qrq='" & mod1.DQda & "' where Qdbh='" & lblBh.Caption & "' and btz=23 and zid=" & (Val(lblLc.Caption) - 1)
                Set mod1.HTP = CreateObject("adodb.recordset")
                mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
                If lblLc.Caption = 2 Then
                    tt = "Select username,userid from worker where bm='" & lblBM.Caption & "' and bmjl=1"
                    Set mod1.HTP = CreateObject("adodb.recordset")
                    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
                    Tywy = mod1.HTP.Fields("username").Value
                    Tuid = mod1.HTP.Fields("userid").Value
                    If lblYwy.Caption = "宋晓炯" Then
                        Tywy = 周春云
                        Tuid = "HM042"
                    ElseIf lblYwy.Caption = "宋晓炯1" Then
                        Tywy = "彭海翔"
                        Tuid = "HMG002"
                    ElseIf mod1.BmJl = True And mod1.comId = 0 Then
                        Tywy = "宋晓炯"
                        Tuid = "HM003"
                    ElseIf mod1.BmJl = True And mod1.comId = 1 Then
                        Tywy = "宋晓炯1"
                        Tuid = "HMG000"
                    End If
                ElseIf lblLc.Caption = 3 Then
'                    tt = "Select username,userid from worker where and zzf=1 bq2=1 and qy='" & comQy.Caption & "'"
'                    Set mod1.HTP = CreateObject("adodb.recordset")
'                    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
                    If comQy.Caption = "上海" Then
                        Tywy = "文静"
                        Tuid = "HM266"
                    ElseIf comQy.Caption = "南京" Then
                        Tywy = "王蕾"
                        Tuid = "HM051"
                    ElseIf comQy.Caption = "杭州" Then
                        Tywy = "李艳"
                        Tuid = "HM316"
                    ElseIf comQy.Caption = "北京" Then
                        Tywy = "马玉芝"
                        Tuid = "HM190"
                    ElseIf comQy.Caption = "广州" Then
                        Tywy = "汤丽嫦"
                        Tuid = "HMG023"
                    End If
                    Tywy = mod1.HTP.Fields("username").Value
                    Tuid = mod1.HTP.Fields("userid").Value
                End If
                cmdQm(Index).Caption = mod1.DName
                lblTm(Index).Caption = mod1.DQda
                lblLcRen.Caption = Tywy
                lblLcUid.Caption = Tuid
                tt = "update fyd set lc=" & Val(lblLc.Caption) & ",lcren='" & lblLcRen.Caption & "',lcuid='" & _
                    lblLcUid.Caption & "' where bxid=" & Val(lblBh.Caption)
                Set mod1.HTP = CreateObject("adodb.recordset")
                mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
        ElseIf lblNlb.Caption = 72 Then '旅游费
                Set mod1.cmd = CreateObject("adodb.command")
                mod1.cmd.ActiveConnection = mod1.cc
                mod1.cmd.CommandText = "QMRZQM"
                mod1.cmd.CommandType = adCmdStoredProc
                mod1.cmd.Parameters("@nlb") = lblNlb.Caption '单子(报销单)种类
                mod1.cmd.Parameters("@lc") = lblLc.Caption  '当前流程
                mod1.cmd.Parameters("@Dname") = mod1.DName
                mod1.cmd.Parameters("@uid") = mod1.DHid
                mod1.cmd.Parameters("@btz") = mod1.BTZ '业务属性
                mod1.cmd.Parameters("@zid") = cmdQm(Index).Tag '流程顺序
                mod1.cmd.Parameters("@Qdbh") = lblBh.Caption '单子编号
                mod1.cmd.Parameters("@pje") = ""   '评审建议
                mod1.cmd.Parameters("@bm") = lblBM.Caption
                mod1.cmd.Parameters("@qy") = comQy.Caption
                mod1.cmd.Parameters("@Gren") = lblGui.Caption '如果为费用归属报销单,则添加费用归属人的参数
                mod1.cmd.Parameters("@Guid") = lblGuid.Caption
                mod1.cmd.Parameters("@ywy") = lblYwy.Caption
                mod1.cmd.Parameters("@yid") = lblUid.Caption
                mod1.cmd.Parameters("@comid") = mod1.comId
                mod1.cmd.Execute
                Tywy = mod1.cmd.Parameters("@Tywy").Value
                Tuid = mod1.cmd.Parameters("@Tuid").Value
                Set cmd = Nothing
                If lblLc.Caption = 2 Then
                ElseIf lblLc.Caption = 3 Then
                    Tywy = "宋晓炯"
                    Tuid = "HM003"
                ElseIf lblLc.Caption = 4 Then
                    Tywy = "文静"
                    Tuid = "HM266"
                End If
        End If
                
If lblQM(Index).Caption = "报销人" And (lblNlb.Caption = 9 Or lblNlb.Caption = 11 Or lblNlb.Caption = 12 Or lblNlb.Caption = 32 Or lblNlb.Caption = 33 Or lblNlb.Caption = 50 Or lblNlb.Caption = 51 Or lblNlb.Caption = 71) Then
    If lblQM(Index + 1).Caption = "业务审核" Then
        '添加事务
        Call mod1.EnventAddB("报销单", txtHg.Text, lblLcRen.Caption, lblLcUid.Caption, lblBh.Caption, lblQM(Index + 1).Caption, Oywy, Ouid, lblYwy.Caption, lblUid.Caption, lblFwid.Caption, lblBh.Caption)
        MsgBox "现在,这张单子将由其他业务审核人来审核"
    End If
'ElseIf lblQM(Index).Caption = "报销人" And (lblNlb.Caption = 32 Or lblNlb.Caption = 33) Then '费用归属报销单
' Call mod1.EnventAdd("报销单", txtHg.Text, lblLcRen.Caption, lblLcUid.Caption, lblBh.Caption, lblQM(Index + 1).Caption, Oywy, Ouid, lblYwy.Caption, lblUid.Caption, lblFwid.Caption, lblBh.Caption)
'    MsgBox "现在,这张单子将由费用归属人 " & lblGui.Caption & " 来审核"
ElseIf lblQM(Index).Caption = "财务审核" Then
    MsgBox "快发钱吧," & lblYwy.Caption & "早已裤兜底朝天了."
Else
    
    '添加事务
    Call mod1.EnventAdd("报销单", txtHg.Text, lblLcRen.Caption, lblLcUid.Caption, lblBh.Caption, lblQM(Index + 1).Caption, Oywy, Ouid, lblYwy.Caption, lblUid.Caption, lblFwid.Caption, lblBh.Caption)
    MsgBox "现在,这张单子将交由 " & Tywy & " 来审阅!"
End If

If Dialog.Visible = True Then '更新事务列表
    Call mod1.refEnvent(1)
End If
cmdMod.Enabled = False
cmdSave.Enabled = False
End Sub

Private Sub cmdPje_Click()

Dim tt As String
On Error Resume Next
Pje.Show
Set Pje.adoPje = CreateObject("adodb.recordset")
tt = "select trq,ywy,zn,bz,tf from pizu where bh='" & lblBh.Caption & "' and yid=" & lblNlb.Caption & " order by pid desc"
Pje.adoPje.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText

Dim Ra: Dim La
Dim ii As Integer: Dim oo As Integer
Ra = Pje.adoPje.GetRows
Pje.adoPje.Close
Set Pje.adoPje = Nothing
La = UBound(Ra, 2): Pje.dtgPje.Rows = La + 20
Pje.dtgPje.Clear
For oo = 1 To La + 1
    Pje.dtgPje.Row = oo
    For ii = 1 To 6
        Pje.dtgPje.Col = ii
        Pje.dtgPje.Text = Ra(ii - 1, oo - 1)
        If ii = 5 Then
            If Pje.dtgPje.Text = "True" Then
                Pje.dtgPje.Text = "同意"
            ElseIf Pje.dtgPje.Text = "False" Then
                Pje.dtgPje.Text = "驳回"
            End If

        End If
    Next
Next
For oo = 1 To La + 1
    Pje.dtgPje.Row = oo
    Pje.dtgPje.Col = 5
            If Pje.dtgPje.Text = "驳回" Then
                For ii = 1 To 5
                    Pje.dtgPje.Col = ii
                    Pje.dtgPje.CellForeColor = &HFF&
                Next
            End If
Next
Pje.dtgPje.Row = 0
Pje.dtgPje.Col = 1: Pje.dtgPje.Text = "日期": Pje.dtgPje.Col = 2: Pje.dtgPje.Text = "姓名": Pje.dtgPje.Col = 3: Pje.dtgPje.Text = "职能"
Pje.dtgPje.Col = 4: Pje.dtgPje.Text = "评审建议": Pje.dtgPje.Col = 5: Pje.dtgPje.Text = "通过否"
Pje.dtgA.Clear
Pje.dtgA.Rows = Pje.dtgPje.Rows
Pje.dtgA.Cols = Pje.dtgPje.Cols
For oo = 0 To Pje.dtgPje.Rows
    Pje.dtgPje.Row = oo
    Pje.dtgA.Row = oo
    For ii = 0 To Pje.dtgPje.Cols
        Pje.dtgPje.Col = ii
        Pje.dtgA.Col = ii
        Pje.dtgA.Text = Pje.dtgPje.Text
    Next
Next
End Sub

Private Sub cmdQm_Click(Index As Integer)
Dim TMlc As Integer '将TMX置为0时出错的次数.
Dim tt As String
Dim oo As Integer
Dim Tywy As String '单子流转到下一人的姓名
Dim Tuid As String
Dim Oywy As String '原来流转人的名字
Dim Ouid As String '原来流转人的工号
Dim ii As Integer
Dim Je As Double
Dim khmcT As String
Dim ywyT As String
On Error Resume Next


If cmdQm(Index).Caption <> "" Then Exit Sub

If (Index + 1 <> Val(lblLc.Caption)) Then  '不能在不相干的位置上乱点
    Exit Sub
End If

If mod1.Bm = "工程部" Then '工程部单子，一定要填写工作单编号

End If

'If lblLcUid.Caption <> mod1.DHid And lblQM(Index).Caption <> "业务审核" Then
If lblLcRen.Caption <> mod1.DName Then
    MsgBox "此处应由" & lblLcRen.Caption & "签字! 请您不要再点"
    Exit Sub
End If
If Index > 0 Then
    If cmdQm(Index - 1).Caption = "完毕" And Val(lblLc.Caption) <> Index + 1 Then
        Exit Sub
    End If
End If
If Index = 0 And cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If

            '检查归属人员正确设置否
            If dtgNx.Visible = False Then
                For oo = 1 To adoF2.Recordset.RecordCount
                    khmcT = ""
                    ywyT = ""
                    dtgBx.Row = oo
                    dtgBx.Col = 1
                    khmcT = dtgBx.Text
                    dtgBx.Col = 42
                    ywyT = dtgBx.Text
                    If ywyT = "" And cmdQm(0).Caption <> "" Then
                        MsgBox "帮帮忙，帮我填好归属人员好吗？"
                        MsgBox "真累！"
                        Exit Sub
                    End If
                Next
            End If
If lblQM(Index).Caption = "财务审核" And (mod1.DName = "宋晓炯" Or mod1.DName = "宋晓炯1") Then
    Exit Sub
End If

            If optFp1.Value = False And optFp2.Value = False Then
                MsgBox "请确认发票情况!"
                cmdSave.Enabled = True
                Exit Sub
            End If
            
            If optFp2.Value = True And txtFP.Text = "" Then
                MsgBox "请注明发票不一致的原因!"
                cmdSave.Enabled = True
                Exit Sub
            End If

If Val(lblNewF.Caption) = 1 Or Val(lblNlb.Caption = 82) Then '新单签字方式和内部结算
    If lblGui.Caption = "" And Not (lblBt.Caption = "三金报销单" Or lblBt.Caption = "公积金报销单" Or _
    lblBt.Caption = "福利报销单" Or lblBt.Caption = "外来综合保险报销单" Or lblBt.Caption = "培训报销单") Then
        MsgBox ("请选择费用归属人！")
        cmdSave.Enabled = True
        Exit Sub
    End If
    QF = False
    frmNQ.Visible = True
    cmdDing.Enabled = True
    
    If Index = 0 Then '报销人只能签字，不能驳回。
        optT2.Enabled = False
        
    Else
        optT2.Enabled = True
    End If
    Exit Sub
End If













If lblLc.Caption > 1 Then
    ii = MsgBox("您是否核准此单？(选择“是”将签字通过,选择“否”将驳回此单)", vbYesNoCancel + vbInformation, "请您注意!")
    If ii = vbNo Then
        ii = MsgBox("驳回后,此单将回转至填单人" & lblYwy.Caption & ",确认吗?", vbYesNo + vbInformation, "确认驳回吗?")
        If ii = vbNo Then
            Exit Sub
        End If
        tt = InputBox("请输入您要驳回的原因!")
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "xtzxFAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@yid").Value = lblNlb.Caption  '反签名
        mod1.cmd.Parameters("@lc").Value = lblLc.Caption
        mod1.cmd.Parameters("@bh").Value = lblBh.Caption
        mod1.cmd.Parameters("@ywy").Value = mod1.DName
        mod1.cmd.Parameters("@uid").Value = mod1.DHid
        mod1.cmd.Parameters("@bz").Value = tt
        mod1.cmd.Parameters("@zn").Value = lblQM(Index).Caption '身份职能
        mod1.cmd.Execute
        Set cmd = Nothing
        For oo = 0 To 6
            cmdQm(oo).Caption = ""
            lblTm(oo).Caption = ""
        Next
        lblLc.Caption = 999 '不让再按签名按钮.
        If Dialog.Visible = True Then '更新事务列表
            Call mod1.refEnvent(1)
        End If
        Exit Sub
    ElseIf ii = vbCancel Then
        Exit Sub
    End If
End If
Oywy = lblLcRen.Caption
Ouid = lblLcUid.Caption





'确认费用归属到最小的归属人单位
If cmdQm(Index).Caption = "组长" Or cmdQm(Index).Caption = "部门经理" Then

End If



If lblLc.Caption = 1 Then
Dim Zi As Integer
Zi = MsgBox("是否确认签字?", vbYesNo)
If Zi = vbNo Then Exit Sub
End If

'验证表头与表身的一致性.
If lblLc.Caption > 1 Then
    Je = 0
    adoF2.Recordset.MoveFirst
    oo = 1
    Do While Not adoF2.Recordset.EOF
        Je = Je + adoF2.Recordset.Fields("XG").Value
        adoF2.Recordset.MoveNext
        oo = oo + 1
    Loop
        If Val(txtHg.Text) <> Round(Je, 2) Then
            MsgBox "总金额与明细金额不一致,请退回此单!"
            Exit Sub
        End If
End If
If lblQM(Index).Caption = "业务审核" Then '如果是业务审核签字,则当前流程不变,直到全部签字后流程才改变.
    If lblYqf.Caption = 0 Then
        MsgBox "请在明细栏中审核与您相关的费用!"
        Exit Sub
    End If
ElseIf lblQM(Index).Caption = "财务审核" Then '如果是财务审核签字,则当前流程不变,直到签收后流程才改变.

Else
    lblLc.Caption = lblLc.Caption + 1
End If

    '设置服务器纠错系统
    tt = "update fyd set tmx=0 where bxid=" & Val(lblBh.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    TMlc = 0
CmdqmClick:
    TMlc = TMlc + 1
    If TMlc = 5 Then
        ii = MsgBox("网络出现严重故障,请稍候片刻再试!", vbExclamation, "C级警报")
        Exit Sub
    End If
    On Error GoTo CmdqmClick
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    On Error Resume Next

    If lblNlb.Caption <> 54 And lblNlb.Caption <> 70 And lblNlb.Caption <> 66 And lblNlb.Caption <> 67 And lblNlb.Caption <> 72 Then
    '更新表Fyd中的lcRen,lcUid 字段,以及QMRZ表中的相应字段.
                Set mod1.cmd = CreateObject("adodb.command")
                mod1.cmd.ActiveConnection = mod1.cc
                mod1.cmd.CommandText = "QMRZQM"
                mod1.cmd.CommandType = adCmdStoredProc
                mod1.cmd.Parameters("@nlb") = lblNlb.Caption '单子(报销单)种类
                mod1.cmd.Parameters("@lc") = lblLc.Caption  '当前流程
                mod1.cmd.Parameters("@Dname") = mod1.DName
                mod1.cmd.Parameters("@uid") = mod1.DHid
                mod1.cmd.Parameters("@btz") = mod1.BTZ '业务属性
                mod1.cmd.Parameters("@zid") = Index + 1 '流程顺序
                mod1.cmd.Parameters("@Qdbh") = lblBh.Caption '单子编号
                mod1.cmd.Parameters("@pje") = ""   '评审建议
                mod1.cmd.Parameters("@bm") = lblBM.Caption
                mod1.cmd.Parameters("@qy") = comQy.Caption
                mod1.cmd.Parameters("@Gren") = lblGui.Caption '如果为费用归属报销单,则添加费用归属人的参数
                mod1.cmd.Parameters("@Guid") = lblGuid.Caption
                mod1.cmd.Parameters("@ywy") = lblYwy.Caption
                mod1.cmd.Parameters("@yid") = lblUid.Caption
                mod1.cmd.Parameters("@comid") = mod1.comId
                mod1.cmd.Execute
                Tywy = mod1.cmd.Parameters("@Tywy").Value
                Tuid = mod1.cmd.Parameters("@Tuid").Value
                Set cmd = Nothing
                If (Tywy = "文静" And comQy.Caption <> "上海") Or (Tywy = "王蕾" And comQy.Caption = "南京") Then
                    If comQy.Caption = "南京" Then
                        Tywy = "王蕾"
                        Tuid = "HM051"
                    ElseIf comQy.Caption = "杭州" Then
                        Tywy = "李艳"
                        Tuid = "HM316"
                    ElseIf comQy.Caption = "北京" Then
                        Tywy = "马玉芝"
                        Tuid = "HM190"
                    ElseIf comQy.Caption = "广州" Then
                        Tywy = "汤丽嫦"
                        Tuid = "HMG023"
                    End If

                    tt = "update fyd set lcren='" & Tywy & "',lcUid='" & Tuid & "',lc=" & Val(lblLc.Caption) & " where bxid=" & lblBh.Caption
                    Set mod1.HTP = CreateObject("adodb.recordset")
                    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
                End If
                If Tywy = "宋晓炯" And mod1.comId = 1 Then
                    Tywy = "宋晓炯1"
                    Tuid = "HMG000"
                    tt = "update fyd set lc=" & Val(lblLc.Caption) & ",lcren='" & Tywy & "',lcuid='" & _
                        Tuid & "' where bxid=" & lblBh.Caption
                    Set mod1.HTP = CreateObject("adodb.recordset")
                    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
                End If
                If Tywy = "倪旭" And mod1.comId = 1 Then '如果广州的单子误跳到倪旭,则置为宋晓炯
                    Tywy = "宋晓炯1"
                    Tuid = "HMG000"
                    tt = "update fyd set lc=" & Val(lblLc.Caption) & ",lcren='" & Tywy & "',lcuid='" & _
                        Tuid & "' where bxid=" & lblBh.Caption
                    Set mod1.HTP = CreateObject("adodb.recordset")
                    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
                End If
                
                '如果为广州总经理报销,则审核人为彭海翔
                If Tywy = "周春云" And lblYwy.Caption = "宋晓炯1" And comQy.Caption = "广州" Then
                    Tywy = "彭海翔"
                    Tuid = "HMG002"
                    tt = "update fyd set lc=" & Val(lblLc.Caption) & ",lcren='" & Tywy & "',lcuid='" & _
                        Tuid & "' where bxid=" & lblBh.Caption
                    Set mod1.HTP = CreateObject("adodb.recordset")
                    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
                End If
                cmdQm(Index).Caption = mod1.DName
                lblTm(Index).Caption = mod1.DQda
                lblLcRen.Caption = Tywy
                lblLcUid.Caption = Tuid
        ElseIf lblNlb.Caption = 54 Or lblNlb.Caption = 70 Then '工程部报销单
                Set mod1.cmd = CreateObject("adodb.command")
                mod1.cmd.ActiveConnection = mod1.cc
                mod1.cmd.CommandText = "QMRZGC"
                mod1.cmd.CommandType = adCmdStoredProc
                mod1.cmd.Parameters("@nlb") = lblNlb.Caption '单子(报销单)种类
                mod1.cmd.Parameters("@lc") = lblLc.Caption  '当前流程
                mod1.cmd.Parameters("@Dname") = mod1.DName
                mod1.cmd.Parameters("@uid") = mod1.DHid
                mod1.cmd.Parameters("@btz") = mod1.BTZ '业务属性
                mod1.cmd.Parameters("@zid") = cmdQm(Index).Tag '流程顺序
                mod1.cmd.Parameters("@Qdbh") = lblBh.Caption '单子编号
                mod1.cmd.Parameters("@pje") = ""   '评审建议
                mod1.cmd.Parameters("@bm") = lblBM.Caption
                mod1.cmd.Parameters("@qy") = comQy.Caption
                mod1.cmd.Parameters("@Gren") = lblGui.Caption '如果为费用归属报销单,则添加费用归属人的参数
                mod1.cmd.Parameters("@Guid") = lblGuid.Caption
                mod1.cmd.Parameters("@ywy") = lblYwy.Caption
                mod1.cmd.Parameters("@yid") = lblUid.Caption
                
                mod1.cmd.Execute
                
                Tywy = mod1.cmd.Parameters("@Tywy").Value
                Tuid = mod1.cmd.Parameters("@Tuid").Value
                Set cmd = Nothing
                
                
                
                If (Tywy = "文静" And comQy.Caption <> "上海") Or (Tywy = "王蕾" And comQy.Caption = "南京") Or lblQM(Index + 1).Caption = "财务审核" Then
                    If comQy.Caption = "南京" Then
                        Tywy = "王蕾"
                        Tuid = "HM051"
                    ElseIf comQy.Caption = "杭州" Then
                        Tywy = "李艳"
                        Tuid = "HM316"
                    ElseIf comQy.Caption = "北京" Then
                        Tywy = "马玉芝"
                        Tuid = "HM190"
                    ElseIf comQy.Caption = "广州" Then
                        Tywy = "汤丽嫦"
                        Tuid = "HMG023"
                    End If
                '    tt = "update QMRZ set  Qren='" & mod1.DName & "',Qrid='" & mod1.DHid & "',Qrq='" & mod1.DQda & "' where Qdbh='" & txtHtbh.Text & "' and btz=" & mod1.BTZ & " and zid=" & cmdQm(Index).Tag
                '    Set mod1.HTP = CreateObject("adodb.recordset")
                '    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
                    tt = "update fyd set lcren='" & Tywy & "',lcUid='" & Tuid & "',lc=" & Val(lblLc.Caption) & " where bxid=" & lblBh.Caption
                    Set mod1.HTP = CreateObject("adodb.recordset")
                    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
                End If
                cmdQm(Index).Caption = mod1.DName
                lblTm(Index).Caption = mod1.DQda
                lblLcRen.Caption = Tywy
                lblLcUid.Caption = Tuid
        ElseIf lblNlb.Caption = 67 Or lblNlb.Caption = 66 Then '房屋补贴
                tt = "update QMRZ set  Qren='" & mod1.DName & "',Qrid='" & mod1.DHid & "',Qrq='" & mod1.DQda & "' where Qdbh='" & lblBh.Caption & "' and btz=23 and zid=" & (Val(lblLc.Caption) - 1)
                Set mod1.HTP = CreateObject("adodb.recordset")
                mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
                If lblLc.Caption = 2 Then
                    tt = "Select username,userid from worker where bm='" & lblBM.Caption & "' and bmjl=1"
                    Set mod1.HTP = CreateObject("adodb.recordset")
                    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
                    Tywy = mod1.HTP.Fields("username").Value
                    Tuid = mod1.HTP.Fields("userid").Value
                    If lblYwy.Caption = "宋晓炯" Then
                        Tywy = 周春云
                        Tuid = "HM042"
                    ElseIf lblYwy.Caption = "宋晓炯1" Then
                        Tywy = "彭海翔"
                        Tuid = "HMG002"
                    ElseIf mod1.BmJl = True And mod1.comId = 0 Then
                        Tywy = "宋晓炯"
                        Tuid = "HM003"
                    ElseIf mod1.BmJl = True And mod1.comId = 1 Then
                        Tywy = "宋晓炯1"
                        Tuid = "HMG000"
                    End If
                ElseIf lblLc.Caption = 3 Then
'                    tt = "Select username,userid from worker where and zzf=1 bq2=1 and qy='" & comQy.Caption & "'"
'                    Set mod1.HTP = CreateObject("adodb.recordset")
'                    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
                    If comQy.Caption = "上海" Then
                        Tywy = "文静"
                        Tuid = "HM266"
                    ElseIf comQy.Caption = "南京" Then
                        Tywy = "王蕾"
                        Tuid = "HM051"
                    ElseIf comQy.Caption = "杭州" Then
                        Tywy = "李艳"
                        Tuid = "HM316"
                    ElseIf comQy.Caption = "北京" Then
                        Tywy = "马玉芝"
                        Tuid = "HM190"
                    ElseIf comQy.Caption = "广州" Then
                        Tywy = "汤丽嫦"
                        Tuid = "HMG023"
                    End If
                    Tywy = mod1.HTP.Fields("username").Value
                    Tuid = mod1.HTP.Fields("userid").Value
                End If
                cmdQm(Index).Caption = mod1.DName
                lblTm(Index).Caption = mod1.DQda
                lblLcRen.Caption = Tywy
                lblLcUid.Caption = Tuid
                tt = "update fyd set lc=" & Val(lblLc.Caption) & ",lcren='" & lblLcRen.Caption & "',lcuid='" & _
                    lblLcUid.Caption & "' where bxid=" & Val(lblBh.Caption)
                Set mod1.HTP = CreateObject("adodb.recordset")
                mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
        ElseIf lblNlb.Caption = 72 Then '旅游费
                Set mod1.cmd = CreateObject("adodb.command")
                mod1.cmd.ActiveConnection = mod1.cc
                mod1.cmd.CommandText = "QMRZQM"
                mod1.cmd.CommandType = adCmdStoredProc
                mod1.cmd.Parameters("@nlb") = lblNlb.Caption '单子(报销单)种类
                mod1.cmd.Parameters("@lc") = lblLc.Caption  '当前流程
                mod1.cmd.Parameters("@Dname") = mod1.DName
                mod1.cmd.Parameters("@uid") = mod1.DHid
                mod1.cmd.Parameters("@btz") = mod1.BTZ '业务属性
                mod1.cmd.Parameters("@zid") = cmdQm(Index).Tag '流程顺序
                mod1.cmd.Parameters("@Qdbh") = lblBh.Caption '单子编号
                mod1.cmd.Parameters("@pje") = ""   '评审建议
                mod1.cmd.Parameters("@bm") = lblBM.Caption
                mod1.cmd.Parameters("@qy") = comQy.Caption
                mod1.cmd.Parameters("@Gren") = lblGui.Caption '如果为费用归属报销单,则添加费用归属人的参数
                mod1.cmd.Parameters("@Guid") = lblGuid.Caption
                mod1.cmd.Parameters("@ywy") = lblYwy.Caption
                mod1.cmd.Parameters("@yid") = lblUid.Caption
                mod1.cmd.Parameters("@comid") = mod1.comId
                mod1.cmd.Execute
                Tywy = mod1.cmd.Parameters("@Tywy").Value
                Tuid = mod1.cmd.Parameters("@Tuid").Value
                Set cmd = Nothing
                If lblLc.Caption = 2 Then
                ElseIf lblLc.Caption = 3 Then
                    Tywy = "宋晓炯"
                    Tuid = "HM003"
                ElseIf lblLc.Caption = 4 Then
                    Tywy = "文静"
                    Tuid = "HM266"
                End If
        End If
                
If lblQM(Index).Caption = "报销人" And (lblNlb.Caption = 9 Or lblNlb.Caption = 11 Or lblNlb.Caption = 12 Or lblNlb.Caption = 32 Or lblNlb.Caption = 33 Or lblNlb.Caption = 50 Or lblNlb.Caption = 51 Or lblNlb.Caption = 71) Then
    If lblQM(Index + 1).Caption = "业务审核" Then
        '添加事务
        Call mod1.EnventAddB("报销单", txtHg.Text, lblLcRen.Caption, lblLcUid.Caption, lblBh.Caption, lblQM(Index + 1).Caption, Oywy, Ouid, lblYwy.Caption, lblUid.Caption, lblFwid.Caption, lblBh.Caption)
        MsgBox "现在,这张单子将由其他业务审核人来审核"
    End If
'ElseIf lblQM(Index).Caption = "报销人" And (lblNlb.Caption = 32 Or lblNlb.Caption = 33) Then '费用归属报销单
' Call mod1.EnventAdd("报销单", txtHg.Text, lblLcRen.Caption, lblLcUid.Caption, lblBh.Caption, lblQM(Index + 1).Caption, Oywy, Ouid, lblYwy.Caption, lblUid.Caption, lblFwid.Caption, lblBh.Caption)
'    MsgBox "现在,这张单子将由费用归属人 " & lblGui.Caption & " 来审核"
ElseIf lblQM(Index).Caption = "财务审核" Then
    MsgBox "快发钱吧," & lblYwy.Caption & "早已裤兜底朝天了."
Else
    
    '添加事务
    Call mod1.EnventAdd("报销单", txtHg.Text, lblLcRen.Caption, lblLcUid.Caption, lblBh.Caption, lblQM(Index + 1).Caption, Oywy, Ouid, lblYwy.Caption, lblUid.Caption, lblFwid.Caption, lblBh.Caption)
    MsgBox "现在,这张单子将交由 " & Tywy & " 来审阅!"
End If

If Dialog.Visible = True Then '更新事务列表
    Call mod1.refEnvent(1)
End If
cmdMod.Enabled = False
cmdSave.Enabled = False

End Sub

Private Sub cmdQm_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If lblQM(Index).Caption = "财务审核" And lblLcUid.Caption = mod1.DHid And Button = 2 Then
    Me.frmNQ.Visible = True
    OptT1.Enabled = False
    optT2.Enabled = True
    optT2.Value = True
End If
End Sub


Private Sub cmdSave_Click()
Dim CEF As Boolean '单笔金额超过500否
Dim tt As String
Dim Fbt As String '报销单名称
Dim CZF As Boolean '包含出租车否
Dim oo As Integer
Dim khmcT As String
Dim ywyT As String
On Error Resume Next
'If dtgNx.Visible = False Then
            If optFp1.Value = False And optFp2.Value = False Then
                MsgBox "请确认发票情况!"
                cmdSave.Enabled = True
                Exit Sub
            End If
            
            If optFp2.Value = True And txtFP.Text = "" Then
                MsgBox "请注明发票不一致的原因!"
                cmdSave.Enabled = True
                Exit Sub
            End If
            
            If (mod1.DName = "文静" Or mod1.DName = "乔继敏") And lblLc.Caption > 1 Then
                tt = "update fyd set cwBz='" & txtCwBZ.Text & "' where bxid=" & lblBh.Caption
                Set mod1.HTP = CreateObject("adodb.recordset")
                mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
                cmdSave.Enabled = False
                Exit Sub
            End If
            
            
            If txtHg.Text = "" Or Val(txtHg.Text) = 0 Then
                MsgBox "没有金额,不能保存!"
                Exit Sub
            End If
            
            '检查归属人员正确设置否
            If dtgNx.Visible = False Then
                For oo = 0 To adoF2.Recordset.RecordCount - 1
                    khmcT = ""
                    ywyT = ""
                    dtgBx.Row = oo
                    dtgBx.Col = 1
                    khmcT = dtgBx.Text
                    dtgBx.Col = 42
                    ywyT = dtgBx.Text
                    If ywyT = "" And cmdQm(0).Caption <> "" Then
                        MsgBox "帮帮忙，帮我填好归属人员好吗？"
                        MsgBox "真累！"
                        Exit Sub
                    End If
                Next
            End If
            
            If lblNlb.Caption = 32 Then  '如果为费用归属,则要注明归属人
                adoF2.Recordset.MoveFirst
                Do While Not adoF2.Recordset.EOF
                    If IsNull(adoF2.Recordset.Fields("ywy").Value) = True Or adoF2.Recordset.Fields("ywy").Value = "" Then
                        MsgBox "请选择归属人!"
                        Call cmdGui_Click
                        Exit Sub
                    End If
                    adoF2.Recordset.MoveNext
                Loop
            
            End If
            
            '如果最后一条记录为空,则删除它
            adoF2.Recordset.MoveLast
            If adoF2.Recordset.Fields("XG").Value = 0 Or IsNull(adoF2.Recordset.Fields("XG").Value) = True Then
                adoF2.Recordset.Delete adAffectCurrent
            End If
            
            If Right(Fbt, 1) = "单" Then
                Fbt = Mid(lblBt.Caption, 1, Len(lblBt.Caption) - 3)
            Else
                Fbt = lblBt.Caption
            End If
            
            '新费用归属
            If lblNlb.Caption = 79 Then

                    tt = "Select * from fyD where Bxid='" & lblBh.Caption & "'"
                    Set mod1.HTP = CreateObject("adodb.recordset")
                    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
                    mod1.HTP.Update "Qy", comQy.Caption
                    mod1.HTP.Update "Trq", LblTrq.Caption
                    mod1.HTP.Update "BM", lblBM.Caption
                    mod1.HTP.Update "hG", Val(txtHg.Text)
                    mod1.HTP.Update "hGD", lblDx.Caption
                    mod1.HTP.Update "fRQ", lblFR.Caption
                    mod1.HTP.Update "lRQ", lblLr.Caption
                    mod1.HTP.Update "QrQ", lblRq.Caption
                    mod1.HTP.Update "yWy", mod1.DName
                    mod1.HTP.Update "uid", mod1.DHid
                    mod1.HTP.Update "gren", lblGui.Caption
                    mod1.HTP.Update "guid", lblGuid.Caption
                    mod1.HTP.Update "fbt", Fbt '报销单名称
                    mod1.HTP.Update "fp", optFp1.Value
                    mod1.HTP.Update "fpnr", Left(txtFP.Text, 200)
                    mod1.HTP.Update "lc", 1 '由于只能由报销人保存,所以保存后流程将由0变为1
                    lblLc.Caption = 1
                    mod1.HTP.Update "BZ", Left(txtBz.Text, 100) '备注
                    mod1.HTP.Update "NLB", lblNlb.Caption '单子类型
                    mod1.HTP.Update "CEF", CEF
                    mod1.HTP.Update "Gren", lblGui.Caption '费用归属人
                    mod1.HTP.Update "Grid", lblGuid.Caption
                    mod1.HTP.Update "lcren", mod1.DName
                    mod1.HTP.Update "lcuid", mod1.DHid
                    mod1.HTP.UpdateBatch
                    
                    '更新FyBx表
'''                    adoF2.Recordset.MoveFirst
'''                    Do While Not adoF2.Recordset.EOF
'''                        adoF2.Recordset.Update "ywy", lblGui.Caption
'''                        adoF2.Recordset.Update "ywyUid", lblGuid.Caption
'''                        adoF2.Recordset.MoveNext
'''                    Loop
'''                    adoF2.Recordset.UpdateBatch
                    tt = "update fybx set ywy='" & lblGui.Caption & "',ywyuid='" & lblGuid.Caption & "' where bxid=" & Val(lblBh.Caption)
                    Set mod1.HTP = CreateObject("adodb.recordset")
                    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
                    Set mod1.HTP = Nothing
                    cmdSave.Enabled = False
                    cmdAdd.Visible = False
                    cmdDel.Visible = False
                    dtgBx.AllowUpdate = False
                    cmdGui.Visible = False
                    
                    '添加事务
                    lblLcRen.Caption = lblGui.Caption
                    lblLcUid.Caption = lblGuid.Caption
                    If Val(lblNlb.Caption) = 79 Then
                        lblLcRen.Caption = mod1.DName
                        lblLcUid.Caption = mod1.DHid
                    End If
                    If Val(lblLc.Caption) = 1 Then
                        Call mod1.EnventAdd("报销单", txtHg.Text, lblLcRen.Caption, lblLcUid.Caption, lblBh.Caption, lblQM(0).Caption, "", "", mod1.DName, mod1.DHid, 0, lblBh.Caption)
                    End If
                          
                          '更新报销单据列表
                    If frmBxBrow.Visible = True Then
                        frmBxBrow.optMe.Value = True
                        tt = "FydV('" & mod1.DHid & "','" & mod1.DName & "')"
                        frmBxBrow.AdoBxBro.Close
                        frmBxBrow.AdoBxBro.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
                        Set frmBxBrow.mga.DataSource = frmBxBrow.AdoBxBro
            
                    End If
                    
                    
                    
                    'MsgBox "您签过字后,豪曼信息将以最快速度将您的报销单送至相关人员审核!"
            
            
                Exit Sub
            End If
            
            CEF = False
                    CZF = False
                    If Left(lblBt.Caption, 3) = "业务员" Then
                        adoF2.Recordset.MoveFirst
                        Do While Not adoF2.Recordset.EOF
                            If adoF2.Recordset.Fields("gzdh").Value <> "" Then
                                CZF = True
                                Exit Do
                            End If
                            adoF2.Recordset.MoveNext
                        Loop
                    End If
            
                    
                    tt = "Select * from fyD where Bxid='" & lblBh.Caption & "'"
                    Set mod1.HTP = CreateObject("adodb.recordset")
                    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
                    mod1.HTP.Update "Qy", comQy.Caption
                    mod1.HTP.Update "Trq", LblTrq.Caption
                    mod1.HTP.Update "BM", lblBM.Caption
                    mod1.HTP.Update "hG", Val(txtHg.Text)
                    mod1.HTP.Update "hGD", lblDx.Caption
                    mod1.HTP.Update "fRQ", lblFR.Caption
                    mod1.HTP.Update "lRQ", lblLr.Caption
                    mod1.HTP.Update "QrQ", lblRq.Caption
                    mod1.HTP.Update "yWy", mod1.DName
                    mod1.HTP.Update "fbt", Fbt '报销单名称
                    mod1.HTP.Update "fp", optFp1.Value
                    mod1.HTP.Update "fpnr", Left(txtFP.Text, 200)
'''''                    mod1.HTP.Update "lc", 1 '由于只能由报销人保存,所以保存后流程将由0变为1
                    mod1.HTP.Update "czf", CZF
'''''                    lblLc.Caption = 1
                    mod1.HTP.Update "BZ", Left(txtBz.Text, 100) '备注
                   
                    '单笔金额超额否
                    adoF2.Recordset.MoveFirst
                    Do While Not adoF2.Recordset.EOF
                        If adoF2.Recordset.Fields("XG").Value > 500 Then
                            CEF = True
                            Exit Do
                        End If
                        adoF2.Recordset.MoveNext
                    Loop
                    If CEF = False Then '根据单笔金额超过500元否,来最后判断相应的Nlb值
                        Select Case Val(lblNlb.Caption)
                            Case 11
                                lblNlb.Caption = 12
                                
            '                Case 15
            '                    lblNlb.Caption = 16
                            Case 17
                                lblNlb.Caption = 18
                            Case 20
                                lblNlb.Caption = 21
                            Case 32
                                If mod1.Bm = "工程部" Then
                                    lblNlb.Caption = 71
                                End If
                            Case 50                '运费
                                lblNlb.Caption = 51
                            Case 54 '工程部
                                lblNlb.Caption = 70
                                
                        End Select
                    
                    End If
                    mod1.HTP.Update "NLB", lblNlb.Caption '单子类型
                    mod1.HTP.Update "CEF", CEF
                    mod1.HTP.Update "Gren", lblGui.Caption '费用归属人
                    mod1.HTP.Update "Grid", lblGuid.Caption
                   
                    mod1.HTP.UpdateBatch
                    '因为流程改变,重新更新Qmrz表中的值
                    Set mod1.cmd = CreateObject("adodb.command")
                    mod1.cmd.ActiveConnection = mod1.cc
                    mod1.cmd.CommandText = "qmrzRef"
                    mod1.cmd.CommandType = adCmdStoredProc
                    mod1.cmd.Parameters("@btz") = 23 '报销单
                    mod1.cmd.Parameters("@qdbh") = lblBh.Caption
                    mod1.cmd.Parameters("@nlb") = lblNlb.Caption
                    mod1.cmd.Execute
                    Set cmd = Nothing
                    
                    
                    
                    '更新FyBx表
                    adoF2.Recordset.UpdateBatch
                    
            
                  'frmFYBX.Visible = False
            cmdSave.Enabled = False
            cmdAdd.Visible = False
            cmdDel.Visible = False
            frmYf.Visible = False
            dtgBx.AllowUpdate = False
            

    

            frmWd.Visible = False
            
            '更新签字按钮的值
            Call ModBx.OpenAN
        
'''''''''''''''''''''''''
            
            '添加事务
            Call mod1.EnventAdd("报销单", txtHg.Text, lblLcRen.Caption, lblLcUid.Caption, lblBh.Caption, lblQM(0).Caption, "", "", mod1.DName, mod1.DHid, 0, lblBh.Caption)
            lblLcRen.Caption = mod1.DName
            lblLcUid.Caption = mod1.DHid
                  
                  '更新报销单据列表
            If frmBxBrow.Visible = True Then
                frmBxBrow.optMe.Value = True
                tt = "FydV('" & mod1.DHid & "','" & mod1.DName & "')"
                frmBxBrow.AdoBxBro.Close
                frmBxBrow.AdoBxBro.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
                Set frmBxBrow.mga.DataSource = frmBxBrow.AdoBxBro
            ElseIf Dialog.Visible = True Then '更新事务列表
                Call mod1.refEnvent(1)
            End If
            
            
            
            'MsgBox "您签过字后,豪曼信息将以最快速度将您的报销单送至相关人员审核!"

    If txtHg.Text = "" Or Val(txtHg.Text) = 0 Then
        MsgBox "没有金额,不能保存!"
        Exit Sub
    End If
            If Val(lblNlb.Caption) <> 79 Then
                cmdGui.Visible = False
            Else
                cmdGui.Visible = True
            End
            
End If


End Sub













Private Sub cmdXQ_Click()
comYwy.Enabled = True
comXmmc.Enabled = True
txtHtbh.Text = "售前"
'lblWd.Visible = True
comYwy.Text = ""
comXmmc.Text = ""
End Sub

Private Sub cmdXuan_Click()
dtgNx.FixedRows = 0
       dtgNx.MergeCells = 0
End Sub

Private Sub comhtBh_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tt As String
Dim oo As Integer
On Error Resume Next
If KeyCode = 13 Then

'    tt = "Select htping.htBh,htping.xMmc,htping.xywy,htping.qy,worker.UserBm as BM from htping cross join worker" & _
'         " where htping.htbh='" & comhtBh.Text & "' and htping.ywy=worker.username"
    tt = "htYf('" & comhtBh.Text & "')"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    If mod1.HTP.RecordCount = 1 Then '如果为旧合同,则
        If mod1.HTP.Fields("xjf").Value = False Then
            'adoF2.Recordset.AddNew "bxid", frmFYBX.lblBh.Caption
            adoF2.Recordset.Update ("khmc"), mod1.HTP.Fields("xMmc").Value
            adoF2.Recordset.Update ("htBh"), comhtBh.Text
'            adoF2.Recordset.Update ("ywy"), mod1.DName
'            adoF2.Recordset.Update ("ywyuid"), mod1.DHid
'            adoF2.Recordset.Update ("qy"), mod1.Qy
'            adoF2.Recordset.Update ("BM"), mod1.Bm
            adoF2.Recordset.Update ("ywy"), mod1.HTP.Fields("xywy").Value
            adoF2.Recordset.Update ("ywyuid"), mod1.HTP.Fields("uid").Value
            adoF2.Recordset.Update ("qy"), mod1.HTP.Fields("qy").Value
            adoF2.Recordset.Update ("BM"), mod1.HTP.Fields("bm").Value
            adoF2.Recordset.Update ("dep"), mod1.HTP.Fields("bmid").Value
        Else
        End If
    Else
        MsgBox "此合同编号不存在,或此合同不在执行状态,请查验!"
    End If
End If
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub comLb_Change()
Dim tt As String
cmdGui.Visible = True
If comLb.Text = "培训费" Then
    tt = MsgBox("您不能选择此项费用类型，" & comLb.Text & "必须是由当地的行政人员填写培训报销单！", vbInformation, "Hello!")
    comLb.Text = ""
    Exit Sub
End If
If comLb.Text = "福利" Or comLb.Text = "四金" Or comLb.Text = "房屋补贴" Or comLb.Text = "旅游费" Or comLb.Text = "通信费" Then
    tt = MsgBox("您不能选择此项费用类型，" & comLb.Text & "必须是由当地的行政人员填写福利报销单！", vbInformation, "Hello!")
    comLb.Text = ""
    Exit Sub
End If
lblGZDH.Visible = False
txtGZDH.Visible = False
If (comLb.Text = "市内交通费" Or comLb.Text = "市外交通费") Or mod1.Bm = "工程部" Or mod1.Bm = "工程二部" Then
    lblGZDH.Visible = True
    txtGZDH.Visible = True
End If
End Sub

Private Sub comLb_Click()
Dim tt As String
If Me.Visible = False Then Exit Sub
cmdGui.Visible = True
If comLb.Text = "培训费" Then
    tt = MsgBox("您不能选择此项费用类型，" & comLb.Text & "必须是由当地的行政人员填写培训报销单！", vbInformation, "Hello!")
    comLb.Text = ""
    Exit Sub
End If
If comLb.Text = "福利" Or comLb.Text = "四金" Or comLb.Text = "房屋补贴" Or comLb.Text = "旅游费" Or comLb.Text = "通信费" Then
    tt = MsgBox("您不能选择此项费用类型，" & comLb.Text & "必须是由当地的行政人员填写福利报销单！", vbInformation, "Hello!")
    comLb.Text = ""
    Exit Sub
End If
lblGZDH.Visible = False
txtGZDH.Visible = False
If (comLb.Text = "市内交通费" Or comLb.Text = "市外交通费") Or mod1.Bm = "工程部" Or mod1.Bm = "工程二部" Then
    lblGZDH.Visible = True
    txtGZDH.Visible = True
End If
End Sub

Private Sub comXmmc_Click()
Dim tt As String
On Error Resume Next

    adoF2.Recordset.Fields("khmc").Value = comXmmc.Text
    adoF2.Recordset.Fields("ywy").Value = comYwy.Text
    adoF2.Recordset.Fields("htbh").Value = txtHtbh.Text
    
    adoF2.Recordset.Update ("ywyUid"), comYwy.BoundText
    adoF2.Recordset.Update ("qy"), tQy
    adoF2.Recordset.Update ("BM"), Tbm
    Set dtgBx.DataSource = adoF2
    txtHtbh.Text = ""
    comYwy.Text = ""
    comXmmc.Text = ""
    comYwy.Enabled = False
    comXmmc.Enabled = False
    
End Sub


Private Sub comYwy_Click(Area As Integer)
Dim oo As Integer
Dim tt As String
On Error Resume Next
For oo = comXmmc.ListCount - 1 To 0 Step -1
    comXmmc.RemoveItem oo
Next
'mod1.aKhzl.MoveFirst
'Do While Not mod1.aKhzl.EOF
'    If mod1.aKhzl.Fields("ywy").Value = comYwy.Text Then
'        comXmmc.AddItem mod1.aKhzl.Fields("khqc").Value
'        tQy = mod1.aKhzl.Fields("xmqy").Value
'        Tbm = mod1.aKhzl.Fields("bm").Value
'    End If
'    mod1.aKhzl.MoveNext
'Loop
tt = "newKhzl('" & comYwy.Text & "')"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
mod1.HTP.MoveFirst
Do While Not mod1.HTP.EOF
    comXmmc.AddItem mod1.HTP.Fields("khqc").Value
    tQy = mod1.HTP.Fields("xmqy").Value
    Tbm = mod1.HTP.Fields("bm").Value
    mod1.HTP.MoveNext
Loop
End Sub


Private Sub dtgBx_AfterColUpdate(ByVal ColIndex As Integer)
Dim oo As Integer
Dim Je As Single
On Error Resume Next


Je = 0
For oo = 2 To 39
    If IsNull(adoF2.Recordset.Fields(oo).Value) = False Then
        Je = Je + adoF2.Recordset.Fields(oo).Value
    End If
Next
adoF2.Recordset.Fields("XG").Value = Round(Je, 2)
'adoF2.Recordset.Fields("XG").Value = adoF2.Recordset.Fields("NJTF").Value + _
'adoF2.Recordset.Fields("kdF").Value + adoF2.Recordset.Fields("CF").Value + adoF2.Recordset.Fields("yz").Value + _
'adoF2.Recordset.Fields("QTF").Value + adoF2.Recordset.Fields("KDF").Value + adoF2.Recordset.Fields("GJ").Value + _
'adoF2.Recordset.Fields("WL").Value + adoF2.Recordset.Fields("QTF").Value
'txtHg.Text = ""
'lblDx.Caption = ""
End Sub

Private Sub dtgBx_ButtonClick(ByVal ColIndex As Integer)
On Error Resume Next
Dim oo As Integer
Dim YQF As Boolean
Dim Fwid As Double
Dim Bid As Double
Dim Tywy As String
Dim Tuid As String
Dim TZW As String
Dim Bm As String
If lblLc.Caption <> 2 And lblLc.Caption <> 3 Then Exit Sub

Dim Zi As Integer
Zi = MsgBox("是否确认签字?", vbYesNo)
If Zi = vbNo Then Exit Sub

Fwid = adoF2.Recordset.Fields("fwid").Value
Bid = adoF2.Recordset.Fields("bid").Value
If dtgBx.Columns(ColIndex).Caption = "归属人签字" Then
        If dtgBx.Columns("归属人").Text <> mod1.DName Or dtgBx.Columns("归属人签字").Text <> "" Then
            Exit Sub
        End If
tt = "select bm from worker where username='" & adoF2.Recordset.Fields("ywy").Value & "' and userid='" & adoF2.Recordset.Fields("ywyuid").Value & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
'基础发布
'Select Case mod1.Lqy
'Case "上海"
'    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'Case "杭州"
'    mod1.HTP.Open tt, mod1.workHz, adOpenKeyset, adLockReadOnly, adCmdText
'End Select
mod1.HTP.Open tt, mod1.workFF, adOpenKeyset, adLockReadOnly, adCmdText
Bm = mod1.HTP.Fields("bm").Value

        adoF2.Recordset.Update "yqz", mod1.DName
        adoF2.Recordset.Update "yqRq", mod1.DQda
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "qmrzYw"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@yqz") = mod1.DName
        mod1.cmd.Parameters("@bid") = adoF2.Recordset.Fields("bid").Value
        mod1.cmd.Parameters("@yqzUid") = mod1.DHid
        mod1.cmd.Parameters("@bh") = lblBh.Caption
        mod1.cmd.Parameters("@Nlb") = lblNlb.Caption
        mod1.cmd.Parameters("@Lc") = lblLc.Caption
        mod1.cmd.Parameters("@dxren") = lblYwy.Caption
        mod1.cmd.Parameters("@dxUid") = lblUid.Caption
        mod1.cmd.Parameters("@hg") = txtHg.Text
        mod1.cmd.Parameters("@bm") = Bm
        mod1.cmd.Parameters("@qy") = comQy.Caption
        mod1.cmd.Parameters("@ybm") = lblBM.Caption
        mod1.cmd.Execute
        lblYqf.Caption = mod1.cmd.Parameters("@yqf").Value
        lblLc.Caption = mod1.cmd.Parameters("@lc").Value
        
        Set cmd = Nothing
ElseIf dtgBx.Columns(ColIndex).Caption = "部门经理签字" Then
        If (dtgBx.Columns("归经理").Text <> mod1.DName Or dtgBx.Columns("部门经理签字").Text <> "") And Not (dtgBx.Columns("归经理").Text = "张寅") Then
            Exit Sub
        End If
    If adoF2.Recordset.Fields("yqz").Value = "" And lblNlb.Caption <> 72 Then
        MsgBox "请先在业务员位置签字!"
        Exit Sub
    End If

        If mod1.DName = "张寅" Or mod1.DName = "郑刚" And Left(lblBt.Caption, 2) = "工程" Then '工程总监鉴字,方便全牵
            lblYqf.Caption = "True"
            lblLc.Caption = lblLc.Caption + 1
            tt = "update fybx set YWQ='" & mod1.DName & "',YWQUid='" & mod1.DHid & "',ywRq='" & mod1.DQda & "' where bxid=" & Val(lblBh.Caption) & _
            " and ywjl='" & mod1.DName & "' and ywJluid='" & mod1.DHid & "'"
            Set mod1.HTP = CreateObject("adodb.recordset")
            mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
            If lblNlb.Caption = "32" Then
                If mod1.comId = 0 Then
                    Tywy = "宋晓炯"
                    Tuid = "HM003"
                Else
                    Tywy = "宋晓炯1"
                    Tuid = "HMG000"
                End If
                TZW = "总经理"
            ElseIf lblNlb.Caption = "71" Or lblNlb.Caption = "33" Then
                If comQy.Caption = "上海" Then
                    Tywy = "文静"
                    Tuid = "HM266"
                ElseIf comQy.Caption = "杭州" Then
                    Tywy = "李艳"
                    Tuid = "HM316"
                ElseIf comQy.Caption = "南京" Then
                    Tywy = "王蕾"
                    Tuid = "HM051"
                ElseIf comQy.Caption = "北京" Then
                    Tywy = "马玉芝"
                    Tuid = "HM190"
                ElseIf comQy.Caption = "广州" Then
                    Tywy = "汤丽嫦"
                    Tuid = "HMG023"
                End If
                TZW = "财务审核"
            End If

            '添加流程事务.
            Set mod1.cmd = CreateObject("adodb.command")
            mod1.cmd.ActiveConnection = mod1.cc
            mod1.cmd.CommandText = "qmrzYw2"
            mod1.cmd.CommandType = adCmdStoredProc
            mod1.cmd.Parameters("@Tywy") = Tywy
            mod1.cmd.Parameters("@Tuid") = Tuid
            mod1.cmd.Parameters("@lab") = TZW
            mod1.cmd.Parameters("@bh") = lblBh.Caption
            mod1.cmd.Parameters("@lc") = lblLc.Caption
            mod1.cmd.Parameters("@dxren") = lblYwy.Caption
            mod1.cmd.Parameters("@dxUid") = lblUid.Caption
            mod1.cmd.Parameters("@hg") = txtHg.Text
            mod1.cmd.Parameters("@bid") = Bid
            mod1.cmd.Execute
            Set cmd = Nothing

        Else
            Set mod1.cmd = CreateObject("adodb.command")
            mod1.cmd.ActiveConnection = mod1.cc
            mod1.cmd.CommandText = "qmrzYw1"
            mod1.cmd.CommandType = adCmdStoredProc
            mod1.cmd.Parameters("@YWQ") = mod1.DName
            mod1.cmd.Parameters("@YWQUid") = mod1.DHid
            mod1.cmd.Parameters("@bh") = lblBh.Caption
            'mod1.CMD.Parameters("@yqf") = YQF
            mod1.cmd.Parameters("@lc") = lblLc.Caption
            mod1.cmd.Parameters("@nlb") = lblNlb.Caption
            mod1.cmd.Parameters("@dxren") = lblYwy.Caption
            mod1.cmd.Parameters("@dxUid") = lblUid.Caption
            mod1.cmd.Parameters("@hg") = txtHg.Text
            mod1.cmd.Parameters("@fwid") = Fwid
            mod1.cmd.Parameters("@bid") = Bid
            mod1.cmd.Parameters("@qy") = comQy.Caption
            mod1.cmd.Parameters("@ybm") = lblBM.Caption
            mod1.cmd.Parameters("@comid") = mod1.comId
            mod1.cmd.Execute
            lblYqf.Caption = mod1.cmd.Parameters("@yqf").Value
            lblLc.Caption = mod1.cmd.Parameters("@lc").Value
            Set cmd = Nothing
        End If
End If
For oo = 1 To 6
    If lblQM(oo).Caption = "业务审核" Then
        Exit For
    End If
Next
If lblYqf.Caption = "True" Then
    cmdQm(oo).Caption = "完毕"
    
End If

        '由于签字,明细表有了变化,所以刷新费用总表
'        tt = "FydMxOpen(" & lblBh.Caption & ")"
'        frmFYBX.adoF2.Recordset.Close
'        frmFYBX.adoF2.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
        frmFYBX.adoF2.Recordset.Requery
        Set frmFYBX.dtgBx.DataSource = frmFYBX.adoF2
        
If Dialog.Visible = True Then '更新事务列表
    Call mod1.refEnvent(1)
End If
End Sub


Private Sub dtgBx_Click()

'dtgBx.CellBackColor = 255
End Sub

Private Sub dtgBx_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub


Private Sub dtgBx_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'MsgBox dtgBx.Col
'''MsgBox dtgBx.Row
End Sub

Private Sub dtgNx_Click()
Dim oo As Integer
Dim Lb As String
Dim Lrow As Integer '
On Error Resume Next
'''MsgBox dtgNx.Col
'''Exit Sub
If frmEd.Visible = False Then Exit Sub
txtBm.Text = ""
txtGZDH.Text = ""
    dtgNx.Col = 44
    lblBid.Caption = Val(dtgNx.Text)
    dtgNx.Col = 1
    dtPRQ.Value = dtgNx.Text
    dtgNx.Col = 2
    txtNr.Text = dtgNx.Text
    dtgNx.Col = 45
    txtGZDH.Text = Trim(dtgNx.Text)
    For oo = 3 To 40
        dtgNx.Col = oo
        If Val(dtgNx.Text) > 0 Then
            txtJe.Text = Val(dtgNx.Text)
            Lrow = dtgNx.Row
            dtgNx.Row = 0
            comLb.Text = dtgNx.Text
            dtgNx.Row = Lrow
            Exit For
        End If
    Next
    dtgNx.Col = 48
    If Val(dtgNx.Text) = 1 Then
        opt2.Value = True
    ElseIf Val(dtgNx.Text) = 2 Then
        opt1.Value = True
    End If
    dtgNx.Col = 49
    txtBm.Text = dtgNx.Text
End Sub

Private Sub Form_Load()
Dim oo As Integer
Dim tt As String
Dim Ra: Dim La
On Error Resume Next
Me.Left = 0
Me.Top = 0
frmFYBX.Width = mod1.FWidth
frmFYBX.Height = mod1.FHeight
Set F2 = CreateObject("adodb.recordset")
frmMb.BorderStyle = 0
frmNewQ.BorderStyle = 0
frmNQ.Left = 1860
frmNQ.Top = 7350
frmRen.BorderStyle = 0
dtgNx.Left = 0
dtgNx.Top = 1620

tt = "select bm from bm where zzf=1 order by bmid"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2)
For oo = 0 To La
    txtBm.AddItem Ra(0, oo)
Next

Set aY = CreateObject("adodb.recordset")
'tt = "renOpenYwy"
'aY.Close
'aY.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
tt = "select * from renyuan where not(xlx is null) order by bm"
aY.Close
'基础发布
'Select Case mod1.Lqy
'Case "上海"
'    aY.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'Case "杭州"
'    aY.Open tt, mod1.workHz, adOpenKeyset, adLockReadOnly, adCmdText
'End Select
aY.Open tt, mod1.workFF, adOpenKeyset, adLockReadOnly, adCmdText
Set comYwy.RowSource = aY
comYwy.ListField = "username"
comYwy.BoundColumn = "userid"
frmMb.Width = frmFYBX.Width
dtgBx.Width = frmFYBX.Width
frmAn.Left = frmFYBX.Width - frmAn.Width
frmAn.Top = frmFYBX.Height - frmAn.Height - 500
txtFP.Locked = True
Set Fmx = CreateObject("adodb.recordset")
dtgNx.ColWidth(0) = 300
dtgNx.ColWidth(2) = 2500
dtgNx.ColWidth(40) = 0
dtgNx.ColWidth(47) = 0
dtgNx.ColWidth(48) = 0 'GongF
dtgNx.ColWidth(49) = 0 'GBM
frmFYBX.dtgNx.ColWidth(44) = 0
dtPRQ.Value = Date

Me.frmEd.Left = 9690
Me.frmEd.Top = 1530
comLb.AddItem "市内交通费"
comLb.AddItem "市外交通费"
comLb.AddItem "招待费"
comLb.AddItem "餐费"
comLb.AddItem "住宿费"
comLb.AddItem "礼品费"
'''''comLb.AddItem "通信费"
'comLb.AddItem "办公用品"
comLb.AddItem "运费"
comLb.AddItem "快递费"
'''''comLb.AddItem "福利"
comLb.AddItem "部门团队费"
'''''comLb.AddItem "房屋补贴"
'''''comLb.AddItem "高温费"
'''''comLb.AddItem "旅游费"
'comLb.AddItem "房租"
'comLb.AddItem "物业费"
'comLb.AddItem "水电"
'comLb.AddItem "电话"
'comLb.AddItem "市场推广"
'comLb.AddItem "人员招聘"
'comLb.AddItem "培训费"
comLb.AddItem "财务手续费"
comLb.AddItem "团队建设费"
comLb.AddItem "停车费"
comLb.AddItem "车辆费"
'comLb.AddItem "公共停车费"
'comLb.AddItem "公共车辆费"
'comLb.AddItem "工具"
'comLb.AddItem "易耗"
'comLb.AddItem "外劳"
comLb.AddItem "邮资"

frmNQ.Left = 2250
frmNQ.Top = 7380
timWait.Enabled = False
timQuit.Enabled = False
If mod1.Bq2 = True Then
    txtQc.Enabled = True
Else
    txtQc.Enabled = False
End If
dtgP.Top = 6360

dtgP.Left = 0
End Sub

Private Sub Form_Resize()
'frmBxBrow.WindowState = 2
'Call mod1.ResizeForm(Me) '确保窗体改变时控件随之改变
'frmMb.Width = frmFYBX.Width
'dtgBx.Width = frmFYBX.Width
'frmAn.Left = frmFYBX.Width - frmAn.Width
'frmAn.Top = frmFYBX.Height - frmAn.Height - 500
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MDI.Cq = False Then
Call mod1.DelDKZ ' '退出表单时删除打开记录,以让别人能打开此单据
Cancel = True
frmFYBX.Visible = False
If frmBxBrow.Visible = True Then
    frmBxBrow.Enabled = True
    frmBxBrow.ZOrder 0
    'frmBxBrow.WindowState = 2
ElseIf Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
ElseIf fyBB.Visible = True Then
    fyBB.Enabled = True
    fyBB.ZOrder 0
ElseIf frmCWBBA.Visible = True Then
    frmCWBBA.Enabled = True
    frmCWBBA.ZOrder 0
End If
'frmBxBrow.WindowState = 2
End If
End Sub





Private Sub MSHFlexGrid1_Click()

End Sub

Private Sub frmMb_Click()
frmNQ.Visible = False
lblTX.Visible = False

End Sub

Private Sub opt1_Click()
If opt1.Value = True Then
    cmdGui.Visible = True
    txtBm.Text = ""
End If
End Sub

Private Sub opt2_Click()
If opt2.Value = True Then
    cmdGui.Visible = False
    txtBm.Text = ""
    lblGui.Caption = ""
    lblGuid.Caption = ""
End If
End Sub

Private Sub optFp1_Click()
If optFp1.Value = True And cmdSave.Enabled = True Then
    txtFP.Locked = True
    MsgBox "请注意保管好您的发票,以备签收时交给财务验收!"
    cmdSave.Enabled = True
End If
End Sub

Private Sub optFp2_Click()
If optFp2.Value = True And cmdSave.Enabled = True Then
    MsgBox "请详细注明发票不一致的原因,以及用何发票汇进行替代!"
    txtFP.Locked = False
    txtFP.Visible = True
    cmdSave.Enabled = True
End If
End Sub

Private Sub timQuit_Timer()
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0

If timZm = 2 Then '签字
    cmdDing.Enabled = True
    txtQM.Text = ""
    frmQm.Visible = False
    lblTX.Visible = True
    timQuit.Enabled = False
    If Dialog.Visible = True Then
        Call mod1.refEnvent(1)
    End If
ElseIf timZm = 3 Then
    cmdDing.Enabled = True
    txtQM.Text = ""
    frmQm.Visible = False
    lblTX.Visible = True
    timQuit.Enabled = False
    If Dialog.Visible = True Then
        Call mod1.refEnvent(1)
    End If
ElseIf timZm = 5 Then '签收
    If comDQ.Text = "" Then
        txtQc.Text = lblYwy.Caption
    Else
        txtQc.Text = comDQ.Text
    End If
    txtQc.PasswordChar = ""
    lblRq.Caption = mod1.DQda
    If Day(mod1.DQda) >= 25 Then
        lblRq.Caption = DateSerial(Year(mod1.DQda), Month(mod1.DQda) + 1, 1)
    End If
    txtQc.Enabled = False


    lblTX.Visible = True
    timQuit.Enabled = False
    If Dialog.Visible = True Then
        Call mod1.refEnvent(1)
    End If
    MsgBox "恭喜发财,红包拿来! :)"
    cmdBack.SetFocus
ElseIf timZm = 1 Then '费用编辑
    If txtBm.Text <> "" Then
         'lblGui.Caption = txtBm.Text
    End If
End If
End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
Dim oo As Integer
Dim LZw As String
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh,mm1,mt1,mm2,mt2,mt3 from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0
    timWait.Enabled = False
    If timZm = 1 Then
        txtHg.Text = mod1.WP.Fields("mm1").Value
        lblDx.Caption = mod1.ChangBi(Val(txtHg.Text))
        Call ModBx.DiZ
    ElseIf timZm = 2 Then
        frmNQ.Visible = False
        If OptT1.Value = True Then
            cmdQm(lblLc.Caption - 1).Caption = mod1.DName
            lblTm(lblLc.Caption - 1).Caption = mod1.DQda
            If lblLc.Caption = 1 And lblGui.Caption = cmdQm(0).Caption And lblQM(Val(lblLc.Caption)).Caption = "归属人" Then
            cmdQm(lblLc.Caption).Caption = mod1.DName
            lblTm(lblLc.Caption).Caption = mod1.DQda
            End If
        Else
            For oo = 0 To 5
                cmdQm(oo).Caption = ""
                lblTm(oo).Caption = ""
                cmdFQ.Caption = ""
                lblFT.Caption = ""
            Next
        End If
        lblLc.Caption = mod1.WP.Fields("mm1").Value
        lblFwid.Caption = mod1.WP.Fields("mm2").Value
        lblLcRen.Caption = mod1.WP.Fields("mt1").Value
        lblLcUid.Caption = mod1.WP.Fields("mt2").Value
        LZw = mod1.WP.Fields("mt3").Value
        
        If LZw = "可以签收" Then
            lblTX.Caption = "快快发钱，" & cmdQm(0).Caption & "快憋不住啦！"
            txtQc.Enabled = True
            txtQc.Locked = False
        Else
            lblTX.Caption = "下一流程,将跳至" & LZw & ": " & lblLcRen.Caption
        End If
        If Val(lblLc.Caption) = 2 Then
            Call ModBx.DiZ
        End If
        Call QMBound(Val(lblBh.Caption))
    ElseIf timZm = 3 Then
        If OptT1.Value = True Then
            cmdFQ.Caption = mod1.DName
            lblFT.Caption = mod1.DQda
            
        Else
            For oo = 0 To 5
                cmdQm(oo).Caption = ""
                lblTm(oo).Caption = ""
            Next
        End If
        lblLc.Caption = mod1.WP.Fields("mm1").Value
        lblFwid.Caption = mod1.WP.Fields("mm2").Value
        lblLcRen.Caption = mod1.WP.Fields("mt1").Value
        lblLcUid.Caption = mod1.WP.Fields("mt2").Value
        lblTX.Caption = "下一流程,将跳至" & LZw & ": " & lblLcRen.Caption
    End If

    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then '未完成
    
    
ElseIf mod1.WP.Fields("cf").Value = 2 Then  '处理失败
    timWait.Enabled = False
    ii = MsgBox("服务中心在处理您的命令时,发生如下错误:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    txtQc.Text = ""
    lblRq.Caption = ""
    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    txtQc.Text = ""
    lblRq.Caption = ""
    Exit Sub
End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub

Private Sub txtBm_Click()
lblGui.Caption = txtBm.Text
End Sub


Private Sub txtBz_LostFocus()
If Len(txtBz.Text) > 100 Then
    MsgBox ("您的备注超过了100字符,请做适当修减,否则,系统将忽略多余的文字!")
End If
End Sub

Private Sub txtCwBZ_Change()
If lblRq.Caption = "" Then
    txtCwBZ = ""
End If
End Sub

Private Sub txtHtbh_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tt As String
On Error Resume Next
comYwy.Text = ""
comXmmc.Text = ""
If KeyCode = 13 Then
    
    tt = "htXinXi('" & Trim(txtHtbh.Text) & "')"
    mod1.HTT.Close
    mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    If mod1.HTT.RecordCount = 1 Then
            If IsNull(mod1.HTT.Fields("xuid")) = True Then
                MsgBox "旧合同中的内置参数有冗余，请将此合同编号发公告栏给我，我将以最快速度解决此问题！ 马晓聪" & mod1.DQda
            End If
            comYwy.Text = mod1.HTT.Fields("xywy").Value
            comXmmc.Text = mod1.HTT.Fields("xMmc").Value
            adoF2.Recordset.Fields("khmc").Value = comXmmc.Text
            adoF2.Recordset.Fields("ywy").Value = comYwy.Text
            adoF2.Recordset.Fields("htbh").Value = txtHtbh.Text
            adoF2.Recordset.Update ("qy"), mod1.HTT.Fields("qy").Value
            adoF2.Recordset.Update ("BM"), mod1.HTT.Fields("BM").Value
            adoF2.Recordset.Update ("ywyUid"), mod1.HTT.Fields("Xuid").Value
            Set dtgBx.DataSource = adoF2
    Else
        MsgBox ("输入的编号有误!")
        txtHtbh.Text = ""
    
    End If
    
    txtHtbh.Text = ""
    comYwy.Text = ""
    comXmmc.Text = ""

End If
End Sub

Private Sub txtNr_Change()
If Len(txtNr.Text) >= 29 Then
    MsgBox "字数太多,建议写进备注! 否则超过30字数将不被保存!"
End If
End Sub

Private Sub txtQc_Change()
If txtQc.Text <> cmdQm(0).Caption Then
txtQc.PasswordChar = "*"
End If
End Sub

Private Sub txtQc_KeyDown(KeyCode As Integer, Shift As Integer)
If mod1.DName <> "朱佳宇" And mod1.DName <> "汤丽嫦" Then
    txtQc.Text = ""
    MsgBox "个人签收已经停止！"
    Exit Sub
End If
Dim tt As String
Dim oo As Integer
Dim Je As Double
Dim Df As Boolean
Dim ZF As Long
Dim Gf As Long
On Error Resume Next
If KeyCode = 13 Then
'    If comDQ.Text = "" Then
'        tt = "Select UserPw,userid from worker where userName='" & cmdBxr.Caption & "'"
'    Else
'        tt = "Select UserPw,userid from worker where userName='" & comDQ.Text & "'"
'    End If
'    Set mod1.HTP = CreateObject("adodb.recordset")
'    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText

'为结帐,每月21-24日不能签收.
    If Not (mod1.DName = "李艳" Or mod1.DName = "朱佳宇" Or mod1.DName = "王蕾" Or mod1.DName = "马玉" Or mod1.DName = "汤丽嫦") Then
        MsgBox "你在开国际玩笑!"
        Exit Sub
    End If

'*******************************************************
'''''If Day(mod1.DQda) >= 21 And Day(mod1.DQda) <= 24 Then
'''''    MsgBox ("结帐期,不能签收! 25日对外开放!")
'''''    Exit Sub
'''''End If

'验证表头与表身的一致性.
Je = 0
adoF2.Recordset.MoveFirst
oo = 1
Do While Not adoF2.Recordset.EOF
    Je = Je + adoF2.Recordset.Fields("XG").Value
    adoF2.Recordset.MoveNext
    oo = oo + 1
Loop
If Round(Val(txtHg.Text), 2) <> Round(Je, 2) Then
    MsgBox "总金额与明细金额不一致,请退回此单!"
    Exit Sub
End If


''''''''''''''''''If mod1.TX = 0 And mod1.CName <> "andy" Then
''''''''''''''''''        '豪曼信息
''''''''''''''''''                Set mod1.cmd = createobject("adodb.command")
''''''''''''''''''                mod1.cmd.ActiveConnection = mod1.CC
''''''''''''''''''                mod1.cmd.CommandText = "JPW"
''''''''''''''''''                mod1.cmd.CommandType = adCmdStoredProc
''''''''''''''''''                mod1.cmd.Parameters("@ywy") = lblYwy.Caption
''''''''''''''''''                mod1.cmd.Parameters("@uid") = lblUid.Caption
''''''''''''''''''                mod1.cmd.Parameters("@KDQ") = comDQ.Text
''''''''''''''''''                mod1.cmd.Parameters("@Pw") = txtQc.Text
''''''''''''''''''                mod1.cmd.Parameters("@bxid") = lblBh.Caption
''''''''''''''''''                mod1.cmd.Parameters("@fwid") = lblFwid.Caption
''''''''''''''''''                mod1.cmd.Parameters("@qrq") = mod1.DQda
''''''''''''''''''                '***************************
''''''''''''''''''                'If Day(mod1.DQda) >= 25 Then
''''''''''''''''''                If Day(mod1.DQda) >= 21 Then
''''''''''''''''''                    mod1.cmd.Parameters("@qrq") = DateSerial(Year(mod1.DQda), Month(mod1.DQda) + 1, 1)
''''''''''''''''''                End If
''''''''''''''''''                mod1.cmd.Execute
''''''''''''''''''                Df = mod1.cmd.Parameters("@Df").Value
''''''''''''''''''                Set cmd = Nothing
''''''''''''''''''Else
'''''''''''''''''''导入天兴软件
''''''''''''''''''        Set mod1.cmd = createobject("adodb.command")
''''''''''''''''''        mod1.cmd.ActiveConnection = mod1.CC
''''''''''''''''''        mod1.cmd.CommandText = "TXFyd"
''''''''''''''''''        mod1.cmd.CommandType = adCmdStoredProc
''''''''''''''''''        mod1.cmd.Parameters("@ywy") = lblYwy.Caption
''''''''''''''''''        mod1.cmd.Parameters("@uid") = lblUid.Caption
''''''''''''''''''        mod1.cmd.Parameters("@KDQ") = comDQ.Text
''''''''''''''''''        mod1.cmd.Parameters("@dywy") = ""
''''''''''''''''''        mod1.cmd.Parameters("@duid") = ""
''''''''''''''''''        mod1.cmd.Parameters("@Pw") = txtQc.Text
''''''''''''''''''        mod1.cmd.Parameters("@jw") = ""
''''''''''''''''''        mod1.cmd.Parameters("@bxid") = lblBh.Caption
''''''''''''''''''        mod1.cmd.Parameters("@fwid") = lblFwid.Caption
''''''''''''''''''
''''''''''''''''''        mod1.cmd.Parameters("@Cuid") = mod1.DHid '操作员工号
''''''''''''''''''        mod1.cmd.Parameters("@bz") = Left(txtBz.Text, 50) & "..."
''''''''''''''''''        mod1.cmd.Parameters("@PAY_DD") = DateSerial(Year(mod1.DQda), Month(mod1.DQda) + 1, 1)
''''''''''''''''''        mod1.cmd.Parameters("@CHK_DD") = DateSerial(Year(mod1.DQda), Month(mod1.DQda) + 2, 1)
''''''''''''''''''        mod1.cmd.Parameters("@CHK_DAYS") = Day(DateSerial(Year(mod1.DQda), Month(mod1.DQda) + 1, 1 - 1))
''''''''''''''''''        mod1.cmd.Parameters("@df") = 0 '密码验证值
''''''''''''''''''
''''''''''''''''''        mod1.cmd.Parameters("@hg") = txtHg.Text
''''''''''''''''''        mod1.cmd.Parameters("@cbxid") = Trim(Str(lblBh.Caption))
''''''''''''''''''        mod1.cmd.Parameters("@qrq") = mod1.DQda
''''''''''''''''''        If Day(mod1.DQda) >= 25 Then
''''''''''''''''''            mod1.cmd.Parameters("@qrq") = DateSerial(Year(mod1.DQda), Month(mod1.DQda) + 1, 1)
''''''''''''''''''        End If
''''''''''''''''''        mod1.cmd.Parameters("@date") = DateSerial(Year(mod1.cmd.Parameters("@qrq").Value), Month(mod1.cmd.Parameters("@qrq").Value), Day(mod1.cmd.Parameters("@qrq").Value))
''''''''''''''''''        mod1.cmd.Parameters("@errch") = ""
''''''''''''''''''        mod1.cmd.Parameters("@errA") = 0
''''''''''''''''''        mod1.cmd.Parameters("@errB") = 0
''''''''''''''''''        mod1.cmd.Parameters("@count") = 0
''''''''''''''''''        mod1.cmd.Execute
''''''''''''''''''        If mod1.cmd.Parameters("@errch").Value <> "成功" Then
''''''''''''''''''            MsgBox "网络出现故障,请再试一次,如果还是提交不成功,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
''''''''''''''''''            txtQc.Text = ""
''''''''''''''''''            Set mod1.cmd = Nothing
''''''''''''''''''            Exit Sub
''''''''''''''''''        End If
''''''''''''''''''        Df = mod1.cmd.Parameters("@Df").Value
''''''''''''''''''        Set mod1.cmd = Nothing
''''''''''''''''''End If
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''    If Df = True Then
''''''''''''''''''        If comDQ.Text = "" Then
''''''''''''''''''            txtQc.Text = lblYwy.Caption
''''''''''''''''''        Else
''''''''''''''''''            txtQc.Text = comDQ.Text
''''''''''''''''''        End If
''''''''''''''''''        txtQc.PasswordChar = ""
''''''''''''''''''        lblRq.Caption = mod1.DQda
''''''''''''''''''        If Day(mod1.DQda) >= 25 Then
''''''''''''''''''            lblRq.Caption = DateSerial(Year(mod1.DQda), Month(mod1.DQda) + 1, 1)
''''''''''''''''''        End If
''''''''''''''''''        txtQc.Enabled = False
''''''''''''''''''
''''''''''''''''''

''''''''''''''''''        MsgBox "恭喜发财,红包拿来! :)"
''''''''''''''''''        'MsgBox "精确制导! 打中天兴!"
''''''''''''''''''        cmdBack.SetFocus
''''''''''''''''''    Else
''''''''''''''''''        txtQc.Text = ""
''''''''''''''''''        txtQc.PasswordChar = "*"
''''''''''''''''''        lblRq.Caption = ""
''''''''''''''''''    End If
'先验证密码正确性
tt = "select userpw from worker where userid='" & lblUid.Caption & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly
If IsNull(mod1.HTP.RecordCount) = True Then
    MsgBox ("验证数据出错，请重新启动豪曼信息再试，如果还有问题，请与马晓聪联系！")
    End
End If
'''''If Not (mod1.HTP.Fields("userpw").Value = txtQc.Text Or txtQc.Text = "hugeman") Then
'''''    MsgBox ("错误密码！")
'''''    Exit Sub
'''''End If
    lblRq.Caption = mod1.DQda
    If Day(mod1.DQda) > 25 Then
        lblRq.Caption = DateSerial(Year(mod1.DQda), Month(mod1.DQda) + 1, 1)
    End If
    txtQc.Enabled = False
    timZm = 5 '签收
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "报销单"
    mod1.cmd.Parameters("@NBLX") = "签收"
    mod1.cmd.Parameters("@bh") = lblBh.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = lblYwy.Caption
    mod1.cmd.Parameters("@mt2") = lblUid.Caption
    mod1.cmd.Parameters("@mt3") = ""
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
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mlt2") = ""
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtJe.Text) '金额
    mod1.cmd.Parameters("@mm2") = Val(lblFwid.Caption)
    mod1.cmd.Parameters("@mm3") = 0
    mod1.cmd.Parameters("@mm4") = 0
    mod1.cmd.Parameters("@mm5") = 0
    mod1.cmd.Parameters("@mm6") = 0
    mod1.cmd.Parameters("@mm7") = 0
    mod1.cmd.Parameters("@mm8") = 0
    mod1.cmd.Parameters("@mm9") = 0
    mod1.cmd.Parameters("@mm10") = 0
    mod1.cmd.Parameters("@mm11") = 0
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
    mod1.cmd.Parameters("@md1") = lblRq.Caption
    mod1.cmd.Parameters("@md2") = Null
    mod1.cmd.Parameters("@md3") = Null
    mod1.cmd.Parameters("@md4") = Null
    mod1.cmd.Parameters("@md5") = Null
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        lblRq.Caption = ""
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
End If
End Sub


