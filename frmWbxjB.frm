VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmWbxjB 
   Caption         =   "报价单"
   ClientHeight    =   9180
   ClientLeft      =   1455
   ClientTop       =   1905
   ClientWidth     =   15240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.Frame frmJz 
      Caption         =   "机组信息"
      Height          =   1905
      Left            =   9150
      TabIndex        =   152
      Top             =   5610
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox txtSl 
         Height          =   285
         Left            =   3630
         Locked          =   -1  'True
         TabIndex        =   154
         Top             =   1230
         Width           =   615
      End
      Begin VB.CommandButton cmdTk 
         Caption         =   "维保条款"
         Height          =   285
         Left            =   3030
         TabIndex        =   153
         Top             =   1590
         Width           =   1245
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgA 
         Height          =   975
         Left            =   30
         TabIndex        =   155
         Top             =   210
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   1720
         _Version        =   393216
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSDataListLib.DataCombo comXh 
         Height          =   330
         Left            =   1020
         TabIndex        =   156
         Top             =   1560
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo comPb 
         Height          =   330
         Left            =   1020
         TabIndex        =   157
         Top             =   1230
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin VB.Label Label3 
         Caption         =   "数量:"
         Height          =   225
         Left            =   3060
         TabIndex        =   160
         Top             =   1290
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "机组品牌:"
         Height          =   225
         Left            =   90
         TabIndex        =   159
         Top             =   1320
         Width           =   1125
      End
      Begin VB.Label Label2 
         Caption         =   "机组型号:"
         Height          =   225
         Left            =   90
         TabIndex        =   158
         Top             =   1635
         Width           =   1095
      End
   End
   Begin VB.TextBox txtBz 
      Height          =   795
      Left            =   9840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   150
      Top             =   5700
      Width           =   3315
   End
   Begin VB.CommandButton cmdCong 
      BackColor       =   &H00C0FFC0&
      Caption         =   "重新评审"
      Height          =   1095
      Left            =   8730
      Style           =   1  'Graphical
      TabIndex        =   149
      Top             =   8100
      Width           =   375
   End
   Begin VB.TextBox txtFbnr 
      Height          =   315
      Left            =   4350
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   148
      Top             =   8430
      Width           =   1095
   End
   Begin VB.CommandButton cmdPje 
      Caption         =   "评审建议"
      Height          =   1095
      Left            =   9120
      TabIndex        =   146
      Top             =   8100
      Width           =   345
   End
   Begin VB.Frame frmYM 
      BackColor       =   &H8000000D&
      Caption         =   "奖金预计支付情况"
      Height          =   2055
      Left            =   3420
      TabIndex        =   99
      Top             =   7410
      Width           =   4665
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除"
         Height          =   285
         Left            =   3960
         TabIndex        =   103
         Top             =   1170
         Width           =   585
      End
      Begin VB.CommandButton cmdYadd 
         Caption         =   "添加"
         Height          =   315
         Left            =   3960
         TabIndex        =   102
         Top             =   810
         Width           =   585
      End
      Begin VB.TextBox txtYingFu 
         Height          =   270
         Left            =   2850
         TabIndex        =   101
         Top             =   1620
         Width           =   1035
      End
      Begin VB.TextBox txtYED 
         Height          =   285
         Left            =   930
         TabIndex        =   100
         Top             =   1620
         Width           =   645
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgYJ 
         Height          =   1275
         Left            =   90
         TabIndex        =   104
         Top             =   210
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   2249
         _Version        =   393216
         BackColorBkg    =   -2147483635
         SelectionMode   =   1
         BorderStyle     =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label31 
         BackColor       =   &H8000000D&
         Caption         =   "支付金额"
         Height          =   225
         Left            =   1980
         TabIndex        =   107
         Top             =   1650
         Width           =   915
      End
      Begin VB.Label Label30 
         BackColor       =   &H8000000D&
         Caption         =   "%"
         Height          =   255
         Left            =   1680
         TabIndex        =   106
         Top             =   1650
         Width           =   195
      End
      Begin VB.Label Label28 
         BackColor       =   &H8000000D&
         Caption         =   "收款额度"
         Height          =   255
         Left            =   90
         TabIndex        =   105
         Top             =   1650
         Width           =   825
      End
   End
   Begin VB.Frame frmGD 
      Caption         =   "项目费用分类"
      Height          =   2625
      Left            =   7710
      TabIndex        =   126
      Top             =   5880
      Visible         =   0   'False
      Width           =   6945
      Begin VB.TextBox txtGd 
         Height          =   285
         Left            =   960
         TabIndex        =   143
         Top             =   2250
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   3270
         TabIndex        =   142
         Text            =   "Text1"
         Top             =   2250
         Width           =   1515
      End
      Begin VB.Frame Frame3 
         Height          =   825
         Left            =   30
         TabIndex        =   127
         Top             =   1290
         Width           =   6885
         Begin VB.OptionButton optGDA 
            Caption         =   "中秋(月饼券)"
            Height          =   195
            Left            =   750
            TabIndex        =   135
            Top             =   180
            Width           =   1545
         End
         Begin VB.OptionButton optGDB 
            Caption         =   "春节(年会吃饭)"
            Height          =   180
            Left            =   2280
            TabIndex        =   134
            Top             =   180
            Width           =   1605
         End
         Begin VB.OptionButton optGDC 
            Caption         =   "其它"
            Height          =   195
            Left            =   3960
            TabIndex        =   133
            Top             =   180
            Width           =   795
         End
         Begin VB.TextBox txtGDNR 
            Height          =   270
            Left            =   4770
            TabIndex        =   132
            Top             =   150
            Width           =   2025
         End
         Begin VB.TextBox txtQdj 
            Height          =   270
            Left            =   750
            TabIndex        =   131
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtRl 
            Height          =   270
            Left            =   2250
            TabIndex        =   130
            Top             =   480
            Width           =   915
         End
         Begin VB.CommandButton cmdGAdd 
            Caption         =   "添加"
            Height          =   255
            Left            =   4980
            TabIndex        =   129
            Top             =   480
            Width           =   795
         End
         Begin VB.CommandButton cmdGdel 
            Caption         =   "删除"
            Height          =   255
            Left            =   5880
            TabIndex        =   128
            Top             =   480
            Width           =   885
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   255
            Left            =   3810
            TabIndex        =   136
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "yyyyy"
            Format          =   108462083
            UpDown          =   -1  'True
            CurrentDate     =   38943
         End
         Begin VB.Label Label40 
            Caption         =   "类别:"
            Height          =   195
            Left            =   120
            TabIndex        =   140
            Top             =   180
            Width           =   555
         End
         Begin VB.Label Label39 
            Caption         =   "单价:"
            Height          =   225
            Left            =   120
            TabIndex        =   139
            Top             =   540
            Width           =   585
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   6870
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Label Label38 
            Caption         =   "人数:"
            Height          =   165
            Left            =   1650
            TabIndex        =   138
            Top             =   540
            Width           =   705
         End
         Begin VB.Label Label36 
            Caption         =   "年份:"
            Height          =   195
            Left            =   3330
            TabIndex        =   137
            Top             =   510
            Width           =   525
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgGD 
         Height          =   1065
         Left            =   60
         TabIndex        =   141
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   1879
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label42 
         Caption         =   "固定费用"
         Height          =   255
         Left            =   120
         TabIndex        =   145
         Top             =   2310
         Width           =   825
      End
      Begin VB.Label Label41 
         Caption         =   "活动费用"
         Height          =   255
         Left            =   2400
         TabIndex        =   144
         Top             =   2280
         Width           =   885
      End
   End
   Begin VB.TextBox txtFbje 
      Height          =   285
      Left            =   4350
      Locked          =   -1  'True
      TabIndex        =   122
      Top             =   8820
      Width           =   1095
   End
   Begin VB.Frame frmFF 
      BackColor       =   &H00C0FFC0&
      Caption         =   "付款方式"
      Height          =   2235
      Left            =   1350
      TabIndex        =   108
      Top             =   7620
      Width           =   6585
      Begin VB.CommandButton cmdDe 
         Caption         =   "删除"
         Height          =   375
         Left            =   6000
         TabIndex        =   119
         Top             =   1830
         Width           =   525
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "添加"
         Height          =   375
         Left            =   6000
         TabIndex        =   118
         Top             =   1440
         Width           =   525
      End
      Begin VB.Frame frmFk 
         BackColor       =   &H00C0FFC0&
         Height          =   555
         Left            =   60
         TabIndex        =   109
         Top             =   1680
         Width           =   5955
         Begin VB.TextBox txtEd 
            Height          =   270
            Left            =   3570
            TabIndex        =   111
            Top             =   150
            Width           =   795
         End
         Begin VB.TextBox txtYrq 
            Height          =   300
            Left            =   990
            Locked          =   -1  'True
            TabIndex        =   110
            Top             =   150
            Width           =   1005
         End
         Begin MSComCtl2.DTPicker dtpYf 
            Height          =   315
            Left            =   990
            TabIndex        =   112
            Top             =   150
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   16711680
            CalendarTrailingForeColor=   8454016
            Format          =   108462081
            CurrentDate     =   38797
         End
         Begin VB.Label Label33 
            BackColor       =   &H00C0FFC0&
            Caption         =   "应付日期"
            Height          =   285
            Left            =   60
            TabIndex        =   116
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label34 
            BackColor       =   &H00C0FFC0&
            Caption         =   "收款额度"
            Height          =   255
            Left            =   2730
            TabIndex        =   115
            Top             =   180
            Width           =   795
         End
         Begin VB.Label Label37 
            BackColor       =   &H00C0FFC0&
            Caption         =   "%"
            Height          =   255
            Left            =   4470
            TabIndex        =   114
            Top             =   180
            Width           =   435
         End
         Begin VB.Label lblFid 
            BackColor       =   &H00C0FFC0&
            Caption         =   "lblFid"
            Height          =   165
            Left            =   4920
            TabIndex        =   113
            Top             =   240
            Visible         =   0   'False
            Width           =   945
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgFk 
         Height          =   1575
         Left            =   30
         TabIndex        =   117
         Top             =   180
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2778
         _Version        =   393216
         BackColorBkg    =   12648384
         FillStyle       =   1
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "发票类型："
      Height          =   315
      Left            =   5490
      TabIndex        =   95
      Top             =   8790
      Width           =   3465
      Begin VB.OptionButton optLc 
         Caption         =   "服务发票"
         Height          =   195
         Left            =   2190
         TabIndex        =   98
         Top             =   90
         Width           =   1065
      End
      Begin VB.OptionButton optLb 
         Caption         =   "商业发票"
         Height          =   195
         Left            =   1110
         TabIndex        =   97
         Top             =   90
         Width           =   1065
      End
      Begin VB.OptionButton optLa 
         Caption         =   "增值发票"
         Height          =   195
         Left            =   30
         TabIndex        =   96
         Top             =   90
         Width           =   1065
      End
   End
   Begin VB.Frame frmYj 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   8430
      TabIndex        =   90
      Top             =   7740
      Visible         =   0   'False
      Width           =   3105
      Begin VB.TextBox txtTcBe 
         Height          =   285
         Left            =   2670
         TabIndex        =   93
         Text            =   "6"
         Top             =   30
         Width           =   315
      End
      Begin VB.TextBox txtYj 
         Height          =   285
         Left            =   660
         TabIndex        =   91
         Top             =   30
         Width           =   1095
      End
      Begin VB.Label lblTcBe 
         Caption         =   "提成比例"
         Height          =   195
         Left            =   1890
         TabIndex        =   94
         Top             =   60
         Width           =   735
      End
      Begin VB.Label lblYj 
         Caption         =   "奖金"
         Height          =   225
         Left            =   120
         TabIndex        =   92
         Top             =   60
         Width           =   465
      End
   End
   Begin VB.ComboBox comKhmc 
      Height          =   300
      Left            =   5610
      TabIndex        =   84
      Top             =   5940
      Width           =   3390
   End
   Begin VB.CommandButton cmdHt 
      Caption         =   "合同评审单"
      Height          =   285
      Left            =   14100
      TabIndex        =   81
      Top             =   7680
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtYf 
      Height          =   300
      Left            =   1380
      TabIndex        =   80
      Top             =   8760
      Width           =   1275
   End
   Begin VB.TextBox txtXm2 
      Height          =   270
      Left            =   4350
      TabIndex        =   78
      Top             =   8100
      Width           =   1095
   End
   Begin VB.TextBox txtXm1 
      Height          =   285
      Left            =   4350
      Locked          =   -1  'True
      TabIndex        =   76
      Top             =   7770
      Width           =   1095
   End
   Begin VB.Frame frmDx 
      Height          =   375
      Left            =   3330
      TabIndex        =   65
      Top             =   6420
      Width           =   2235
      Begin VB.TextBox txtMon 
         Height          =   270
         Left            =   1200
         TabIndex        =   66
         Top             =   60
         Width           =   525
      End
      Begin VB.Label Label23 
         Caption         =   "月"
         Height          =   255
         Left            =   1950
         TabIndex        =   68
         Top             =   120
         Width           =   195
      End
      Begin VB.Label Label22 
         Caption         =   "维修保质期"
         Height          =   225
         Left            =   120
         TabIndex        =   67
         Top             =   120
         Width           =   1065
      End
   End
   Begin VB.Frame frmNb 
      Height          =   975
      Left            =   3360
      TabIndex        =   58
      Top             =   6420
      Width           =   4125
      Begin VB.TextBox txtF 
         Height          =   300
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   86
         Top             =   270
         Width           =   1455
      End
      Begin VB.TextBox txtL 
         Height          =   300
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   270
         Width           =   1305
      End
      Begin VB.TextBox txtXc 
         Height          =   270
         Left            =   3330
         TabIndex        =   60
         Top             =   660
         Width           =   405
      End
      Begin VB.TextBox txtWc 
         Height          =   270
         Left            =   1050
         TabIndex        =   59
         Top             =   660
         Width           =   495
      End
      Begin MSComCtl2.DTPicker dt4 
         Height          =   315
         Left            =   2430
         TabIndex        =   87
         Top             =   270
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Format          =   107806721
         CurrentDate     =   38098
      End
      Begin MSComCtl2.DTPicker dt3 
         Height          =   315
         Left            =   60
         TabIndex        =   88
         Top             =   270
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         _Version        =   393216
         Format          =   107806721
         CurrentDate     =   38098
      End
      Begin VB.Label Label29 
         Caption         =   "---〉"
         Height          =   225
         Left            =   1950
         TabIndex        =   89
         Top             =   330
         Width           =   375
      End
      Begin VB.Label Label21 
         Caption         =   "次"
         Height          =   225
         Left            =   3840
         TabIndex        =   64
         Top             =   690
         Width           =   315
      End
      Begin VB.Label Label20 
         Caption         =   "例检次数"
         Height          =   225
         Left            =   2430
         TabIndex        =   63
         Top             =   690
         Width           =   825
      End
      Begin VB.Label Label19 
         Caption         =   "年"
         Height          =   225
         Left            =   1650
         TabIndex        =   62
         Top             =   690
         Width           =   255
      End
      Begin VB.Label Label18 
         Caption         =   "维保年限:"
         Height          =   225
         Left            =   60
         TabIndex        =   61
         Top             =   690
         Width           =   855
      End
   End
   Begin VB.TextBox txtClcb 
      Height          =   285
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   8415
      Width           =   1275
   End
   Begin VB.TextBox txtClf 
      Height          =   285
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   8070
      Width           =   1275
   End
   Begin VB.TextBox txtRGF 
      Height          =   270
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   7740
      Width           =   1275
   End
   Begin VB.TextBox txtCb 
      Height          =   285
      Left            =   6570
      TabIndex        =   42
      Top             =   7740
      Width           =   1725
   End
   Begin VB.TextBox txtYhg 
      Height          =   285
      Left            =   6570
      TabIndex        =   39
      Top             =   8430
      Width           =   1725
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印"
      Height          =   675
      Left            =   14310
      Picture         =   "frmWbxjB.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   6300
      Width           =   765
   End
   Begin VB.CommandButton cmdXjd 
      BackColor       =   &H00C0FFC0&
      Caption         =   "询价单"
      Height          =   525
      Left            =   14280
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5700
      Width           =   795
   End
   Begin VB.TextBox comXmmc 
      Height          =   315
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   5940
      Width           =   3075
   End
   Begin VB.CommandButton cmdBack 
      Height          =   375
      Left            =   14760
      Picture         =   "frmWbxjB.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "返回"
      Top             =   8790
      Width           =   465
   End
   Begin VB.TextBox txtZu 
      Height          =   285
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   7170
      Width           =   1725
   End
   Begin VB.TextBox txtHg 
      Height          =   285
      Left            =   6570
      TabIndex        =   19
      Top             =   8085
      Width           =   1725
   End
   Begin VB.Frame frmTime 
      Enabled         =   0   'False
      Height          =   1245
      Left            =   7500
      TabIndex        =   14
      Top             =   6300
      Width           =   1485
      Begin VB.CheckBox chkBc 
         Caption         =   "2小时内到场"
         Height          =   255
         Left            =   150
         TabIndex        =   17
         Top             =   960
         Width           =   1305
      End
      Begin VB.CheckBox chkBb 
         Caption         =   "全年运转"
         Height          =   255
         Left            =   150
         TabIndex        =   16
         Top             =   360
         Width           =   1245
      End
      Begin VB.CheckBox chkBa 
         Caption         =   "24小时运转"
         Height          =   255
         Left            =   150
         TabIndex        =   15
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "时间系数:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdQm 
      Caption         =   "cmdQm"
      Height          =   345
      Index           =   0
      Left            =   9510
      TabIndex        =   13
      Top             =   8400
      Width           =   945
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   14280
      Picture         =   "frmWbxjB.frx":076C
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "保存"
      Top             =   8790
      Width           =   465
   End
   Begin VB.CommandButton cmdMod 
      Height          =   375
      Left            =   13800
      Picture         =   "frmWbxjB.frx":0DD6
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "修改"
      Top             =   8790
      Width           =   465
   End
   Begin VB.Frame frmHide 
      Caption         =   "frmHid"
      Height          =   1455
      Left            =   10260
      TabIndex        =   2
      Top             =   30
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Label lblBm 
         Caption         =   "lblBm"
         Height          =   225
         Left            =   1350
         TabIndex        =   44
         Top             =   150
         Width           =   915
      End
      Begin VB.Label lblQy 
         Caption         =   "lblQy"
         Height          =   255
         Left            =   2820
         TabIndex        =   43
         Top             =   180
         Width           =   1155
      End
      Begin VB.Label lblLc 
         Caption         =   "lblLc"
         Height          =   315
         Left            =   1050
         TabIndex        =   10
         Top             =   630
         Width           =   645
      End
      Begin VB.Label lblNlb 
         Caption         =   "lblNlb"
         Height          =   225
         Left            =   1920
         TabIndex        =   9
         Top             =   810
         Width           =   645
      End
      Begin VB.Label lblLcRen 
         Caption         =   "lblLcRen"
         Height          =   285
         Left            =   150
         TabIndex        =   8
         Top             =   420
         Width           =   795
      End
      Begin VB.Label lblLcUid 
         Caption         =   "lblLcUid"
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   930
         Width           =   885
      End
      Begin VB.Label lblFwid 
         Caption         =   "lblFwid"
         Height          =   255
         Left            =   1860
         TabIndex        =   6
         Top             =   450
         Width           =   1275
      End
      Begin VB.Label lblUid 
         Caption         =   "lblUid"
         Height          =   255
         Left            =   3750
         TabIndex        =   5
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblYwy 
         Caption         =   "lblYwy"
         Height          =   285
         Left            =   3540
         TabIndex        =   4
         Top             =   450
         Width           =   765
      End
      Begin VB.Label lblLcou 
         Caption         =   "lblLcou"
         Height          =   255
         Left            =   1860
         TabIndex        =   3
         Top             =   1080
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdLeft 
      Caption         =   "<"
      Height          =   285
      Left            =   14280
      TabIndex        =   1
      Top             =   8460
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.CommandButton cmdRight 
      Caption         =   ">"
      Height          =   285
      Left            =   14760
      TabIndex        =   0
      Top             =   8460
      Visible         =   0   'False
      Width           =   465
   End
   Begin MSAdodcLib.Adodc adoJi 
      Height          =   375
      Left            =   10980
      Top             =   8520
      Visible         =   0   'False
      Width           =   1425
      _ExtentX        =   2514
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
   Begin MSDataListLib.DataCombo comZu 
      Height          =   330
      Left            =   1380
      TabIndex        =   22
      Top             =   6757
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo2"
   End
   Begin TabDlg.SSTab tabGc 
      Height          =   5505
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   9710
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "年保"
      TabPicture(0)   =   "frmWbxjB.frx":10E0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dtgWb"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "例检"
      TabPicture(1)   =   "frmWbxjB.frx":10FC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dtgLj"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "大修"
      TabPicture(2)   =   "frmWbxjB.frx":1118
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtDxnr"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "材料"
      TabPicture(3)   =   "frmWbxjB.frx":1134
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtTl"
      Tab(3).Control(1)=   "cmdD"
      Tab(3).Control(2)=   "txtDj"
      Tab(3).Control(3)=   "cmdGx"
      Tab(3).Control(4)=   "dtgBao"
      Tab(3).Control(5)=   "dtgMa"
      Tab(3).Control(6)=   "VScroll1"
      Tab(3).Control(7)=   "Label35"
      Tab(3).Control(8)=   "Label25"
      Tab(3).Control(9)=   "Label24"
      Tab(3).ControlCount=   10
      Begin VB.TextBox txtTl 
         Height          =   315
         Left            =   -65430
         TabIndex        =   124
         Top             =   3930
         Width           =   1515
      End
      Begin VB.CommandButton cmdD 
         Caption         =   "删除"
         Height          =   315
         Left            =   -60630
         TabIndex        =   123
         Top             =   3960
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox txtDj 
         Height          =   345
         Left            =   -63000
         TabIndex        =   73
         Top             =   3930
         Width           =   1455
      End
      Begin VB.CommandButton cmdGx 
         Caption         =   "更新"
         Height          =   315
         Left            =   -61380
         TabIndex        =   72
         Top             =   3960
         Width           =   675
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBao 
         Height          =   3915
         Left            =   -75000
         TabIndex        =   69
         Top             =   0
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   6906
         _Version        =   393216
         BackColorBkg    =   -2147483627
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgMa 
         Height          =   885
         Left            =   -75000
         TabIndex        =   70
         Top             =   4290
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   1561
         _Version        =   393216
         BackColor       =   11927477
         BackColorBkg    =   -2147483627
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   30
         Left            =   -73200
         TabIndex        =   49
         Top             =   1530
         Width           =   30
      End
      Begin VB.TextBox txtDxnr 
         BackColor       =   &H80000015&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   5175
         Left            =   -75000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   46
         Top             =   0
         Width           =   15255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgLj 
         Height          =   5175
         Left            =   -75000
         TabIndex        =   48
         Top             =   0
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   9128
         _Version        =   393216
         ForeColorSel    =   -2147483646
         BackColorBkg    =   -2147483627
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgWb 
         Bindings        =   "frmWbxjB.frx":1150
         Height          =   5175
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   9128
         _Version        =   393216
         ForeColorSel    =   -2147483646
         BackColorBkg    =   -2147483627
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label35 
         Caption         =   "数量"
         Height          =   195
         Left            =   -66030
         TabIndex        =   125
         Top             =   3990
         Width           =   495
      End
      Begin VB.Label Label25 
         Caption         =   "单价"
         Height          =   285
         Left            =   -63660
         TabIndex        =   74
         Top             =   3990
         Width           =   465
      End
      Begin VB.Label Label24 
         Caption         =   "采购成本"
         Height          =   225
         Left            =   -74880
         TabIndex        =   71
         Top             =   4050
         Width           =   855
      End
   End
   Begin VB.TextBox txtZt 
      Height          =   315
      Left            =   12510
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   5160
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lblZl 
      Caption         =   "Label19"
      ForeColor       =   &H00C000C0&
      Height          =   225
      Left            =   1410
      TabIndex        =   56
      Top             =   5670
      Width           =   1155
   End
   Begin VB.Label Label44 
      Caption         =   "备注"
      Height          =   225
      Left            =   9150
      TabIndex        =   151
      Top             =   5760
      Width           =   585
   End
   Begin VB.Label Label43 
      Caption         =   "分包内容"
      Height          =   225
      Left            =   3450
      TabIndex        =   147
      Top             =   8490
      Width           =   855
   End
   Begin VB.Label Label32 
      Caption         =   "分包费用"
      Height          =   255
      Left            =   3450
      TabIndex        =   121
      Top             =   8880
      Width           =   765
   End
   Begin VB.Label lblKT 
      Caption         =   "客户名称必须与所签合同客户名称相一致"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   5640
      TabIndex        =   120
      Top             =   5670
      Width           =   3345
   End
   Begin VB.Label Label12 
      Caption         =   "客户名称"
      Height          =   225
      Left            =   4680
      TabIndex        =   83
      Top             =   5970
      Width           =   795
   End
   Begin VB.Label lblhtbh 
      Caption         =   "lblhtbh"
      Height          =   255
      Left            =   14130
      TabIndex        =   82
      Top             =   8070
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label17 
      Caption         =   "运费"
      Height          =   285
      Left            =   660
      TabIndex        =   79
      Top             =   8820
      Width           =   435
   End
   Begin VB.Label Label27 
      Caption         =   "预留项目费用"
      Height          =   225
      Left            =   3090
      TabIndex        =   77
      Top             =   8130
      Width           =   1125
   End
   Begin VB.Label Label26 
      Caption         =   "已发生项目费用"
      Height          =   225
      Left            =   2910
      TabIndex        =   75
      Top             =   7800
      Width           =   1305
   End
   Begin VB.Label lblzlZ 
      Caption         =   "性质"
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   630
      TabIndex        =   57
      Top             =   5670
      Width           =   435
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   15285
      Y1              =   7590
      Y2              =   7590
   End
   Begin VB.Label Label16 
      Caption         =   "材料成本"
      Height          =   285
      Left            =   300
      TabIndex        =   55
      Top             =   8475
      Width           =   765
   End
   Begin VB.Label Label15 
      Caption         =   "差旅费"
      Height          =   285
      Left            =   300
      TabIndex        =   54
      Top             =   8120
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "人工费"
      Height          =   285
      Left            =   300
      TabIndex        =   53
      Top             =   7770
      Width           =   765
   End
   Begin VB.Label Label14 
      Caption         =   "成本总额"
      Height          =   315
      Left            =   5670
      TabIndex        =   41
      Top             =   7770
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "优惠价"
      Height          =   315
      Left            =   5820
      TabIndex        =   40
      ToolTipText     =   "此处由工程部填入"
      Top             =   8430
      Width           =   615
   End
   Begin VB.Label lblBaoId 
      Caption         =   "lblBaoId"
      Height          =   285
      Left            =   10290
      TabIndex        =   36
      Top             =   8340
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "项目名称"
      Height          =   225
      Left            =   270
      TabIndex        =   34
      Top             =   6000
      Width           =   795
   End
   Begin VB.Label Label5 
      Caption         =   "编号"
      Height          =   225
      Left            =   630
      TabIndex        =   33
      Top             =   6390
      Width           =   435
   End
   Begin VB.Label lblBh 
      BackColor       =   &H80000014&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      Height          =   330
      Left            =   1380
      TabIndex        =   32
      Top             =   6345
      Width           =   1725
   End
   Begin VB.Label Label7 
      Caption         =   "工程部组号"
      Height          =   225
      Left            =   90
      TabIndex        =   31
      Top             =   6840
      Width           =   945
   End
   Begin VB.Label Label8 
      Caption         =   "组长"
      Height          =   225
      Left            =   630
      TabIndex        =   30
      Top             =   7230
      Width           =   465
   End
   Begin VB.Label lblBid 
      Caption         =   "lblBid"
      Height          =   255
      Left            =   11460
      TabIndex        =   29
      Top             =   8190
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label6 
      Caption         =   "对外报价"
      Height          =   315
      Left            =   5640
      TabIndex        =   28
      Top             =   8100
      Width           =   765
   End
   Begin VB.Label Label9 
      Caption         =   "总工时"
      Height          =   255
      Left            =   11850
      TabIndex        =   27
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblQM 
      Caption         =   "lblQm"
      Height          =   225
      Index           =   0
      Left            =   9510
      TabIndex        =   26
      Top             =   8130
      Width           =   975
   End
   Begin VB.Label lblTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   0
      Left            =   9510
      TabIndex        =   25
      Top             =   8820
      Width           =   945
   End
   Begin VB.Label lblOid 
      Caption         =   "lblOid"
      Height          =   285
      Left            =   12240
      TabIndex        =   24
      Top             =   8100
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmWbxjB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public adoWb As New ADODB.Recordset
Public adoLj As New ADODB.Recordset
Public adoOid As New ADODB.Recordset '计算Old单子的ADO
Public adoBx As Object '采购表
Public adoGx As Object '成本表
Public adoYj As Object '佣金表
Public adoFk As Object '付款表

Dim AdoKh As Object
Public adoA As Object
Private Sub cmdAdd_Click()
On Error Resume Next
Set mod1.cmd = CreateObject("adodb.command")
mod1.cmd.ActiveConnection = mod1.cc
mod1.cmd.CommandText = "BFkAdd"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@rq") = txtYrq.Text
mod1.cmd.Parameters("@yingfJe") = Round(Val(txtHtze.Text) * Val(txtEd.Text) / 100, 2)
mod1.cmd.Parameters("@htbh") = Trim(txtHtbh.Text)
mod1.cmd.Parameters("@ed") = Round(Val(txtEd.Text) / 100, 2)
mod1.cmd.Execute
Set cmd = Nothing

txtEd.Text = ""
adoFk.Requery
Set dtgFk.DataSource = adoFk
End Sub

Private Sub cmdBack_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
If lblLc.Caption = 0 Then
    ii = MsgBox("该单没有保存,是否将其撤消?", vbQuestion + vbYesNo, "请注意!")
    If ii = vbYes Then
        tt = "delete from baoJiaD where baoid=" & Val(lblBaoId.Caption)
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
        tt = "update xunJiaD set baoid = null where bid=" & Val(lblBid.Caption)
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Else
        Exit Sub
    End If
End If

If frmGxBiao.Visible = True Then
    frmGxBiao.Enabled = True
    frmGxBiao.ZOrder 0
    mod1.BTZ = 36
ElseIf Dialog.Visible = True Then
    Dialog.ZOrder 0
    Dialog.Enabled = True
End If

End Sub





Private Sub cmdCong_Click()
Dim ii As Integer
Dim oo As Integer
Dim tt As String
Dim Bid As Long
Dim ZL As String
Dim HtF As Integer
On Error Resume Next
'MsgBox "正在建设中!"
'Exit Sub
HtF = 88
If lblHtbh.Caption <> "" Then
    tt = "select htf from htping where htbh='" & lblHtbh.Caption & "' and delf=1"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    HtF = mod1.HTP.Fields("htf").Value
    mod1.HTP.Close
End If
Select Case HtF
Case 88
    ii = MsgBox("您的这项操作将使原先单子正在执行的流程全部撤消,是否确定执行?", vbYesNo + vbInformation, "询问")
Case 0
    ii = MsgBox("您的这项操作将使原先单子正在执行的流程全部撤消,包括与它相关尚未执行的合同评审,是否确定执行?", vbYesNo + vbInformation, "询问")
Case 9
    ii = MsgBox("您的这项操作将使原先单子正在执行的流程全部撤消,包括与它相关尚未执行的合同评审,是否确定执行?", vbYesNo + vbInformation, "询问")
Case 1
    ii = MsgBox("此合同正在执行,不要开玩笑!")
    Exit Sub
Case 2
    ii = MsgBox("此合同已经完成,不要开玩笑!")
    Exit Sub
Case Else
    ii = MsgBox("出错,请与马晓聪联系!")
    Exit Sub
End Select


If ii = vbYes Then
    tt = InputBox("请输入您要驳回的原因!")
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "xtzxFAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@yid").Value = 60 '反签名
    mod1.cmd.Parameters("@lc").Value = 2 '退回最初的流程
    mod1.cmd.Parameters("@bh").Value = lblBaoId.Caption
    mod1.cmd.Parameters("@ywy").Value = mod1.DName
    mod1.cmd.Parameters("@uid").Value = mod1.DHid
    mod1.cmd.Parameters("@bz").Value = tt
    mod1.cmd.Parameters("@zn").Value = "new" '身份职能
    mod1.cmd.Execute
    Set cmd = Nothing
    For oo = 0 To 5
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
End Sub

'Private Sub cmdD_Click()
'Dim tt As String
'Dim Lid As Long
'On Error Resume Next
'dtgBao.Col = 16
'Lid = dtgBao.Text
'tt = "delete from "
'tt = "select * from xunJiaMxView where lid=" & Lid
'adoGx.Close
'adoGx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'Set dtgMa.DataSource = adoGx
'End Sub

Private Sub cmdDe_Click()
Dim tt As String
On Error Resume Next
Dim Fid As Long
dtgFk.Col = 5
Fid = dtgFk.Text
tt = "delete from htping1 where fid=" & Fid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText

txtYed.Text = ""
adoFk.Requery
Set dtgFk.DataSource = adoFk
End Sub

Private Sub cmdDel_Click()
Dim tt As String
Dim ii As Integer
Dim hg As Single
Dim Yid As Long
'Dim Ywy As String
On Error Resume Next
'dtgYj.Col = 4
'Ywy = dtgYj.Text
dtgYJ.Col = 3
Yid = 0
Yid = dtgYJ.Text

If Yid = 0 Then
Exit Sub
End If

'If Ywy <> "" Then
'    MsgBox "此单已经激活,不能删除! 如果确定要删除,请与马晓聪联系!"
'    Exit Sub
'End If

ii = MsgBox("是否删除此记录?", vbQuestion + vbYesNo, "询问")
If ii = vbNo Then
    Exit Sub
End If

tt = "delete from byj where yid=" & Yid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
adoYj.Requery
Set dtgYJ.DataSource = adoYj

hg = 0
adoYj.MoveFirst
Do While Not adoYj.EOF
   hg = hg + adoYj.Fields("支付金额").Value
   adoYj.MoveNext
Loop
hg = hg + Val(txtYingFu.Text)

tt = "update baojiaD set yj=" & Val(txtYingFu.Text) & " where baoid=" & Val(lblBaoId.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
txtYJ.Text = hg
End Sub

Private Sub cmdGx_Click()
Dim ii As Integer
Dim CB As Long
Dim liD As Long
Dim Cgid As Long
Dim XCB As Long
On Error Resume Next
If Val(txtTl.Text) = 0 Then
    ii = MsgBox("确认数量为0吗？", vbQuestion + vbYesNo, "您好")
    If ii = vbNo Then Exit Sub
End If
dtgBao.Col = 16
liD = dtgBao.Text
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "baoJiaGx"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@dj") = Val(txtDj.Text)
    mod1.cmd.Parameters("@sl") = Val(txtTl.Text)
    mod1.cmd.Parameters("@lid") = liD
    mod1.cmd.Execute
    'txtHg.Text = Val(txtHg.Text) + mod1.CMD.Parameters("@hg").Value
    Set cmd = Nothing
    
    '获得相应询价单的cgid号
    tt = "select cgid from xunJiaD where bid=" & Val(lblBid.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Cgid = mod1.HTP.Fields("cgid").Value
    '更新相应询价明细中的数量
    tt = "update XunJiaMx set sl=" & Val(txtTl.Text) & ",hg=dj*" & Val(txtTl.Text) & " where lid=" & liD
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    '更新相应询价单中的金额
    tt = "select sum(hg) as hg from xunjiamx where bid=" & Cgid
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'    XCB = 0
'    Do While Not mod1.HTP.EOF
'        XCB = XCB + mod1.HTP.Fields("hg").Value
'        mod1.HTP.MoveNext
'    Loop
    XCB = mod1.HTP.Fields("hg").Value

    tt = "update xunjiaD set hg=" & XCB & ",yhg=" & XCB & " where bid=" & Cgid
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    
    tt = "update baojiaD set clcb=" & XCB & ",hg=" & XCB & "+clf+rgf+yf+ylxm where baoid=" & Val(lblBaoId.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    txtDj.Text = ""
    txtTl.Text = ""
    txtClcb.Text = XCB
    adoBx.Requery
    Set dtgBao.DataSource = adoBx

    txtDj.Text = ""
    txtTl.Text = ""
    
End Sub

Private Sub cmdHt_Click()
Dim tt As String
Dim ii As Integer
Dim FPLX As String
On Error Resume Next
Dim oo As Integer
Dim xZ As String
Dim Hid As Long
On Error Resume Next
tt = "select hid from htping where htbh='" & lblHtbh.Caption & "' and delf=1"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If optLa.Value = True Then
    FPLX = "增值发票"
ElseIf optLb.Value = True Then
    FPLX = "商业发票"
ElseIf optLc.Value = True Then
    FPLX = "服务发票"
End If

If mod1.HTP.RecordCount = 0 And mod1.DName = lblYwy.Caption Then
    '更新表baojiaD中的lcRen,lcUid 字段,以及QMRZ表中的相应字段.
        ii = MsgBox("是否新建合同评审单?", vbQuestion + vbYesNo, "您好!")
        If ii = vbNo Then
            Exit Sub
        End If
        mod1.BTZ = 6
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "HTAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@xmmc") = comXmmc.Text
        mod1.cmd.Parameters("@xid") = comXmmc.Tag
        mod1.cmd.Parameters("@khmc") = comKhmc.Text
        mod1.cmd.Parameters("@khdh") = comKhmc.ToolTipText
        mod1.cmd.Parameters("@ywy") = lblYwy.Caption
        mod1.cmd.Parameters("@uid") = lblUid.Caption
        mod1.cmd.Parameters("@htxz") = lblZl.Caption
        mod1.cmd.Parameters("@htze") = txtYhg.Text
        mod1.cmd.Parameters("@fbje") = txtFbje.Text
        If (lblZl.Caption = "维保" And Val(txtYhg.Text) > 50000) Or Val(txtYhg.Text) > 100000 Or Val(txtFbje.Text) > 0 Then
            mod1.cmd.Parameters("@nlb") = 62
            mod1.cmd.Parameters("@lcou") = 6
        Else
            mod1.cmd.Parameters("@nlb") = 63
            mod1.cmd.Parameters("@lcou") = 5
        End If
        mod1.cmd.Parameters("@fbje") = Val(txtFbje.Text)
        mod1.cmd.Parameters("@baoid") = Val(lblBaoId.Caption)
        mod1.cmd.Parameters("@Hrq") = Format(mod1.DQda, "yyyymmdd")
        mod1.cmd.Parameters("@lc") = 0
        mod1.cmd.Parameters("@lcren") = mod1.DName
        mod1.cmd.Parameters("@lcuid") = mod1.DHid
        mod1.cmd.Parameters("@cbze") = Val(txtCb.Text)
        mod1.cmd.Parameters("@rgf") = Val(txtRgf.Text)
        mod1.cmd.Parameters("@clf") = Val(txtClf.Text)
        mod1.cmd.Parameters("@yf") = Val(txtYf.Text)
        mod1.cmd.Parameters("@yj") = Val(txtYJ.Text)
        mod1.cmd.Parameters("@mon") = Val(txtMOn.Text)
        mod1.cmd.Parameters("@clcb") = Val(txtClcb.Text)
        mod1.cmd.Parameters("@xmfy") = Val(txtXm2.Text)
        mod1.cmd.Parameters("@qy") = mod1.Qy
        mod1.cmd.Parameters("@tcbe") = Val(txtTcBe.Text)
        mod1.cmd.Parameters("@fplx") = FPLX
        mod1.cmd.Parameters("@fbnr") = ""
        mod1.cmd.Parameters("@bz") = Trim(txtBz.Text)
        If lblZl.Caption = "维保" Then
            mod1.cmd.Parameters("@xzCh") = "WB"
            mod1.cmd.Parameters("@htqy") = txtF.Text
            mod1.cmd.Parameters("@htqy1") = txtL.Text
        ElseIf lblZl.Caption = "大修" Then
       
            mod1.cmd.Parameters("@htqy") = "1999-1-1"
            mod1.cmd.Parameters("@htqy1") = "1999-1-1"
            mod1.cmd.Parameters("@xzCh") = "DX"
        ElseIf lblZl.Caption = "工程分包" Then
       
            mod1.cmd.Parameters("@htqy") = "1999-1-1"
            mod1.cmd.Parameters("@htqy1") = "1999-1-1"
            mod1.cmd.Parameters("@xzCh") = "FB"
'        ElseIf lblZl.Caption = "购销" Then
'
'            mod1.CMD.Parameters ("@xzCh")
        End If
        mod1.cmd.Execute
        Hid = mod1.cmd.Parameters("@hid").Value
        Set cmd = Nothing
        
        Call modHt.NewQing
        
        Call modHt.NewBound(Hid)
        frmWbNew.Visible = True
            '设置流程按钮
        If (frmWbNew.lblHtxz = "维保" And Val(frmWbNew.txtHtze) > 50000) Or Val(frmWbNew.txtHtze.Text) > 100000 Or Val(txtFbje.Text) > 0 Then
            Call modHt.HtLcBut(62)
        Else
            Call modHt.HtLcBut(63)
        End If
        frmWbNew.cmdSave.Enabled = True
        frmWbNew.cmdMod.Enabled = False
Else
        mod1.BTZ = 6
        Call modHt.NewQing
        
        Call modHt.NewBound(mod1.HTP.Fields(0).Value)

        frmWbNew.Visible = True
End If

End Sub

Private Sub cmdMod_Click()
If lblYwy.Caption = mod1.DName Then
    cmdCong.Visible = True
End If
txtFbnr.Locked = False
txtFbje.Locked = False
If lblLcRen.Caption <> mod1.DName Or lblLcUid.Caption <> mod1.DHid Then
    Exit Sub
End If
'If lblLc.Caption = 2 Then
'    txtYj.Locked = False
'End If
If lblLc.Caption = 1 Then
    cmdGx.Enabled = True
    comKhmc.Locked = False
    txtYf.Locked = False
    txtXm2.Locked = False
    txtHg.Locked = False
    txtYhg.Locked = False
    txtDj.Locked = False
    dt3.Enabled = True
    dt4.Enabled = True
    txtMOn.Locked = False
    txtMOn.Enabled = True
    frmDx.Enabled = True
    txtTl.Locked = False
    cmdMod.Enabled = False
    cmdSave.Enabled = True
ElseIf lblLc.Caption = 2 Or lblLc.Caption = 3 Then
    cmdGx.Enabled = True
    txtYf.Locked = False
    txtXm2.Locked = False
    txtHg.Locked = False
    txtYhg.Locked = False
    txtDj.Locked = False
    txtTl.Locked = False
    txtFbje.Locked = False
    cmdMod.Enabled = False
    cmdSave.Enabled = True
    'txtYj.Locked = False
End If


End Sub

Private Sub cmdPje_Click()
Dim tt As String
On Error Resume Next
Pje.Show
tt = "select * from pizu where bh='" & lblBaoId.Caption & "' and yid=60 order by trq desc"
Pje.adoPje.Close
Pje.adoPje.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set Pje.dtgPje.DataSource = Pje.adoPje
Pje.txtXQ.Text = ""
End Sub

Private Sub cmdPrint_Click()
Dim tt As String
On Error Resume Next
If cmdSave.Enabled = False Then
    Me.MousePointer = 11
    Set mod1.report = mod1.crapp.OpenReport(App.Path & "\Bjdwb.rpt")
     'Set mod1.report = mod1.crapp.OpenReport(App.Path & "\tt.rpt")
    Set mod1.table = mod1.report.Database.Tables
    Set mod1.cProp = mod1.table.Item(1).ConnectionProperties
    mod1.cProp.Item("Password") = "guyonghui"
    mod1.report.SQLQueryString = "Select * from bjdwb  where baoid=" & Val(lblBaoId.Caption)
    mod1.report.ReadRecords
    frmReport.Show
    frmReport.cR1.ReportSource = mod1.report
    frmReport.cR1.ViewReport

    Me.MousePointer = 0
    frmReport.cR1.EnableExportButton = False
    frmReport.cR1.EnableExportButton = True
End If
End Sub

Private Sub cmdQm_Click(Index As Integer)
Dim tt As String
Dim Tywy As String '单子流转到下一人的姓名
Dim Tuid As String
Dim Oywy As String '原来流转人的名字
Dim Ouid As String '原来流转人的工号
Dim ERRch As String
On Error Resume Next


'If cmdQm(Index).Caption <> "" Then Exit Sub
If Val(txtYhg.Text) = 0 Then
    MsgBox ("开什么国际玩笑,对外报价是0,你想喝西北风啊!")
    Exit Sub
End If
If (txtF.Text = "" Or txtL.Text = "") And lblZl.Caption = "维保" Then
    dt3.Enabled = True
    dt4.Enabled = True
    MsgBox "请确定维保年限!"
    cmdSave.Enabled = True
    Exit Sub
End If

If optLa.Value = False And optLb.Value = False And optLc.Value = False Then
    MsgBox "请确认开票类型!"
    cmdSave.Enabled = True
    Exit Sub
End If
'If Index = 0 And cmdSave.Enabled = True And lblLc.Caption = 0 Then
If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If

If lblLc.Caption = 2 And txtTcBe.Text = "" Then
    MsgBox "请键入提成比例!"
    cmdSave.Enabled = True
    Exit Sub
End If

If Index + 1 <> lblLc.Caption Then '不能在不相干的位置上乱点

    Exit Sub
End If
If comKhmc.Text = "" Then
    MsgBox "请选择相应签约客户名称!"
    cmdSave.Enabled = True
    
    Exit Sub
End If
If lblLcUid.Caption <> mod1.DHid Then
    MsgBox "此处应由" & lblLcRen.Caption & "签字! 请您不要再点"
    Exit Sub
End If

If lblLc.Caption > 1 Then
    ii = MsgBox("您是否核准此单？(选择“是”将签字通过,选择“否”将驳回此单)", vbYesNoCancel + vbInformation, "请您注意!")
    If ii = vbNo Then
        ii = MsgBox("将驳回到报价单的初始流程!", vbYesNo + vbInformation, "确认驳回吗?")
        If ii = vbNo Then
            Exit Sub
        End If
        tt = InputBox("请输入您要驳回的原因!")
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "xtzxFAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@yid").Value = 60 '反签名
        mod1.cmd.Parameters("@lc").Value = lblLc.Caption
        mod1.cmd.Parameters("@bh").Value = lblBaoId.Caption
        mod1.cmd.Parameters("@ywy").Value = mod1.DName
        mod1.cmd.Parameters("@uid").Value = mod1.DHid
        mod1.cmd.Parameters("@bz").Value = tt
        mod1.cmd.Parameters("@zn").Value = lblQM(Index).Caption '身份职能
        mod1.cmd.Execute
        Set cmd = Nothing
        For oo = 0 To 5
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
ElseIf lblLc.Caption = 1 Then
Dim Zi As Integer
Zi = MsgBox("是否确认签字?", vbYesNo)
If Zi = vbNo Then Exit Sub
End If
'    lblLc.Caption = lblLc.Caption + 1
'    Oywy = lblLcRen.Caption
'    Ouid = lblLcUid.Caption
    
'    '更新表baojiaD中的lcRen,lcUid 字段,以及QMRZ表中的相应字段.
'                Set mod1.cmd = createobject("adodb.command")
'                mod1.cmd.ActiveConnection = mod1.CC
'                mod1.cmd.CommandText = "QMRZQM"
'                mod1.cmd.CommandType = adCmdStoredProc
'                mod1.cmd.Parameters("@nlb") = lblNlb.Caption '单子(报销单)种类
'                mod1.cmd.Parameters("@lc") = lblLc.Caption  '当前流程
'                mod1.cmd.Parameters("@Dname") = mod1.DName
'                mod1.cmd.Parameters("@uid") = mod1.DHid
'                mod1.cmd.Parameters("@btz") = mod1.BTZ '业务属性
'                mod1.cmd.Parameters("@zid") = cmdQm(Index).Tag '流程顺序
'                mod1.cmd.Parameters("@Qdbh") = lblBaoId.Caption  '单子编号
'                mod1.cmd.Parameters("@pje") = ""   '评审建议
'                mod1.cmd.Parameters("@bm") = mod1.Bm
'                mod1.cmd.Parameters("@qy") = mod1.Qy
'                mod1.cmd.Parameters("@Gren") = "" '如果为费用归属报销单,则添加费用归属人的参数
'                mod1.cmd.Parameters("@Guid") = ""
'                mod1.cmd.Parameters("@ywy") = lblYwy.Caption
'                mod1.cmd.Parameters("@yid") = lblUid.Caption
'                mod1.cmd.Parameters("@comid") = mod1.comId
'                mod1.cmd.Execute
                Set mod1.cmd = CreateObject("adodb.command")
                mod1.cmd.ActiveConnection = mod1.cc
                If lblZl.Caption = "维保" Then
                    mod1.cmd.CommandText = "QMbj"
                Else
                    mod1.cmd.CommandText = "QMbj1"
                End If
                mod1.cmd.CommandType = adCmdStoredProc
                mod1.cmd.Parameters("@btz") = mod1.BTZ '业务属性
                mod1.cmd.Parameters("@Tywy") = lblLcRen.Caption
                mod1.cmd.Parameters("@Tuid") = lblLcUid.Caption
                mod1.cmd.Parameters("@lc") = Val(lblLc.Caption)
                mod1.cmd.Parameters("@uid") = lblUid.Caption
                mod1.cmd.Parameters("@ywy") = lblYwy.Caption
                mod1.cmd.Parameters("@bh") = Val(lblBaoId.Caption)
                mod1.cmd.Parameters("@fwid") = Val(lblFwid.Caption)
                mod1.cmd.Parameters("@nr") = Trim(comXmmc.Text)
                mod1.cmd.Parameters("@lx") = "报价单"
                mod1.cmd.Parameters("@errch").Value = ""
                mod1.cmd.Execute
                ERRch = mod1.cmd.Parameters("@errch").Value
                If ERRch <> "成功" Then
                        MsgBox "网络出现故障,请再试一次,如果还是提交不成功,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
                        Exit Sub
                End If
                Tywy = mod1.cmd.Parameters("@Tywy").Value
                Tuid = mod1.cmd.Parameters("@Tuid").Value
                lblLc.Caption = mod1.cmd.Parameters("@lc").Value
                lblFwid.Caption = mod1.cmd.Parameters("@fwid").Value
                Set cmd = Nothing
                cmdQm(Index).Caption = mod1.DName
                lblTm(Index).Caption = mod1.DQda
                lblLcRen.Caption = Tywy
                lblLcUid.Caption = Tuid
                



If Val(lblLc.Caption) > Val(lblLcou.Caption) Then

    MsgBox "确认了领导的批示后,您就可以打印报价单了!"
    cmdPrint.Visible = True
    cmdHT.Visible = True

Else
''    If lblLc.Caption = 1 Then '业务员第一个签字,则询价日期等于签字日期
''
''    End If
'    '添加事务
'    Call mod1.EnventAdd("报价单", comXmmc.Text, lblLcRen.Caption, lblLcUid.Caption, lblBaoId.Caption, lblQM(Index + 1).Caption, Oywy, Ouid, lblYwy.Caption, lblUid.Caption, lblFwid.Caption, lblBaoId.Caption)
    MsgBox "现在,此询价单将交由 " & Tywy & " 来审阅!"
End If

If Dialog.Visible = True Then '更新事务列表
    Call mod1.refEnvent(1)
End If
End Sub




Private Sub cmdSave_Click()
Dim tt As String
Dim FPLX As String
On Error Resume Next
Me.Enabled = False
frmWait.Visible = True
frmWait.ZOrder 0
cmdMod.Enabled = True
cmdSave.Enabled = False

If optLa.Value = True Then
    FPLX = "增值发票"
ElseIf optLb.Value = True Then
    FPLX = "商业发票"
ElseIf optLc.Value = True Then
    FPLX = "服务发票"
End If


'先计算成本
txtCb.Text = Val(txtRgf.Text) + Val(txtClcb.Text) + Val(txtClf.Text) + Val(txtXm2.Text) + Val(txtFbje.Text) + Val(txtYf.Text) + Val(txtGd.Text)
'tt = "update baoJiaD set bhg=" & Val(txtHg.Text) & ",yhg=" & Val(txtYhg.Text) & ",yj=" & Val(txtYj.Text) & ",ylxm=" & Val(txtXm2.Text) & " where baoid=" & Val(lblBaoId.Caption)
tt = "update baoJiaD set bhg=" & Val(txtHg.Text) & ",yhg=" & Val(txtYhg.Text) & ",yj=" & Val(txtYJ.Text) & ",ylxm=" & Val(txtXm2.Text) & _
" ,hg=" & Val(txtCb.Text) & " ,yf=" & Val(txtYf.Text) & ",khmc='" & comKhmc.Text & "',khdh='" & comKhmc.ToolTipText & "',tcbe=" & _
Val(txtTcBe.Text) & ",htqy='" & txtF.Text & "',htqy1='" & txtL.Text & "',fplx='" & FPLX & "' where baoid=" & Val(lblBaoId.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'mod1.HTP.Update "xmmc", Trim(comXmmc.Text)    '项目名称
'mod1.HTP.Update "xid", comXmmc.BoundText '项目代号
'mod1.HTP.Update "bianhao", lblBh.Caption '单子编号(给用户看的)
'mod1.HTP.Update "zh", comZu.Text        '组号
'mod1.HTP.Update "Zname", Trim(txtZu.Text)     '组长
'mod1.HTP.Update "jzpb", Trim(comPb.Text)
'mod1.HTP.Update "jzxh", Trim(comXh.Text)
'mod1.HTP.Update "sl", Val(txtSl.Text)
'mod1.HTP.Update "ta", chkBa.Value   '时间系数
'mod1.HTP.Update "tb", chkBb.Value
'mod1.HTP.Update "tc", chkBc.Value
'mod1.HTP.Update "zTime", Val(txtZt.Text) '总工时
'mod1.HTP.Update "hg", Val(txtHg.Text) '总费用
'
'mod1.HTP.UpdateBatch

If lblFwid.Caption = "" Then
    lblLc.Caption = 1
    tt = "update baoJiaD set lc=1 where baoid=" & Val(lblBaoId.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    '添加事务
    Call mod1.EnventAdd("报价单", comXmmc.Text, lblLcRen.Caption, lblLcUid.Caption, lblBaoId.Caption, lblQM(0).Caption, "", "", mod1.DName, mod1.DHid, 0, lblBaoId.Caption)
    '更新按钮
    Call modBJD.OpenBJAN(1)
End If


'
''更新询价列表
'tt = "select * from xunjiaView where ywy='" & mod1.DName & "' and uid='" & mod1.DHid & "'"
'frmGxBiao.adoXj.Close
'frmGxBiao.adoXj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'Set frmGxBiao.dtgXJ.DataSource = frmGxBiao.adoXj

frmWait.Visible = False
Me.Enabled = True
Me.ZOrder 0
End Sub


Private Sub cmdTk_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
If comPb.Text = "" Or comXh.Text = "" Or Val(txtSl.Text) = 0 Then Exit Sub
'年保表
tt = "select * from xunJIaWbView where wbx='年保' and bid=" & Val(frmWbxjB.lblBid.Caption) & " and 机组品牌='" & comPb.Text & "' and 机组型号 like '%" & comXh.Text & "%'"
frmWbxjB.adoWb.Close
frmWbxjB.adoWb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmWbxjB.dtgWb.DataSource = frmWbxjB.adoWb
dtgWb.FixedRows = 0
dtgWb.MergeCol(1) = True
dtgWb.MergeCol(2) = True
dtgWb.MergeCol(3) = True
dtgWb.MergeCells = 3
dtgWb.FixedRows = 1
'例检表
tt = "select * from xunJIaWbView where wbx='例检' and bid=" & Val(frmWbxjB.lblBid.Caption) & " and 机组品牌='" & comPb.Text & "' and 机组型号 like '%" & comXh.Text & "%'"
frmWbxjB.adoLj.Close
frmWbxjB.adoLj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmWbxjB.dtgLj.DataSource = frmWbxjB.adoLj
dtgLj.FixedRows = 0
dtgLj.MergeCol(1) = True
dtgLj.MergeCol(2) = True
dtgLj.MergeCol(3) = True
dtgLj.MergeCells = 3
dtgLj.FixedRows = 1
End Sub

Private Sub cmdXjd_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
If lblLc.Caption = 0 Then
    ii = MsgBox("该单没有保存,是否将其撤消?", vbQuestion + vbYesNo, "请注意!")
    If ii = vbYes Then
        tt = "delete from baoJiaD where baoid=" & Val(lblBaoId.Caption)
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
        tt = "update xunJiaD set baoid = null where bid=" & Val(lblBid.Caption)
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Else
        Exit Sub
    End If
End If
mod1.BTZ = 36
frmWBXJ.Visible = False
If frmWBXJ.lblBid.Caption <> frmWbxjB.lblBid.Caption Then

    Call modBJD.BJDWBQing
    Call modBJD.BJDBound(Val(frmWbxjB.lblBid.Caption), 1)
    
    tt = "select bid from xunjiaOld where old=" & Val(frmWBXJ.lblOid.Caption) & " order by bid"
    frmWBXJ.adoOid.Close
    frmWBXJ.adoOid.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    If frmWBXJ.adoOid.RecordCount > 1 Then
        frmWBXJ.cmdLeft.Enabled = True
    End If
    frmWBXJ.adoOid.MoveLast
End If
frmWait.Visible = False
frmWBXJ.cmdMod.Enabled = True
frmWBXJ.cmdSave.Enabled = False
frmWbxjB.Visible = False
frmWBXJ.Visible = True
frmWBXJ.ZOrder 0
End Sub


Private Sub cmdYadd_Click()
Dim tt As String
Dim hg As Single
On Error Resume Next
If Val(txtYed.Text) = 0 Or Val(txtYingFu.Text) = 0 Then
Exit Sub
End If
hg = 0



Set mod1.cmd = CreateObject("adodb.command")
mod1.cmd.ActiveConnection = mod1.cc
mod1.cmd.CommandText = "byjAdd"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@baoid") = Val(lblBaoId.Caption)
mod1.cmd.Parameters("@YED") = Val(txtYed.Text) / 100
mod1.cmd.Parameters("@yingFu") = Val(txtYingFu.Text)
mod1.cmd.Parameters("@lcren") = mod1.DName
mod1.cmd.Parameters("@lcuid") = mod1.DHid
mod1.cmd.Execute
Set cmd = Nothing
adoYj.Requery
Set dtgYJ.DataSource = adoYj

adoYj.MoveFirst
Do While Not adoYj.EOF
   hg = hg + adoYj.Fields("支付金额").Value
   adoYj.MoveNext
Loop
'HG = HG + Val(txtYingFu.Text)
'If HG > Val(txtYj.Text) Then
'    MsgBox "填写金额有误!"
'    txtYingFu.Text = ""
'    Exit Sub
'End If
On Error Resume Next
tt = "update baojiaD set yj=" & Val(txtYingFu.Text) & " where baoid=" & Val(lblBaoId.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
On Error GoTo YJEB
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
txtYJ.Text = hg
Exit Sub

YJEB:
MsgBox "网络故障,请再试提交一次"

End Sub

Private Sub comKhmc_Change()
cmdSave.Enabled = True
End Sub

Private Sub comKhmc_Click()
Dim tt As String
On Error Resume Next
tt = "Select khdh from khzl where khqc ='" & comKhmc.Text & "'  order by kid desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
comKhmc.ToolTipText = mod1.HTP.Fields("khdh").Value
lblKT.Visible = True
cmdSave.Enabled = True
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

        tt = "Select yzmc,wymc,qt1mc,qt2mc,qt3mc,qt4mc,qt5mc from xmKhmc where xid=" & Val(comXmmc.Tag)

   
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
    End If
    AdoKh.Close
End Sub


Private Sub Command2_Click()

End Sub

Private Sub dt3_CloseUp()
txtF.Text = dt3.Value
End Sub

Private Sub dt4_CloseUp()
txtL.Text = dt4.Value
End Sub

Private Sub dtgA_Click()
On Error Resume Next
dtgA.Col = 4
JxId = dtgA.Text
dtgA.Col = 1
comPb.Text = dtgA.Text
comPb.ToolTipText = dtgA.Text
dtgA.Col = 2
comXh.Text = dtgA.Text
comXh.ToolTipText = dtgA.Text
dtgA.Col = 3
txtSl.Text = dtgA.Text
End Sub

Private Sub dtgA_RowColChange()
On Error Resume Next
dtgA.Col = 4
JxId = dtgA.Text
dtgA.Col = 1
comPb.Text = dtgA.Text
comPb.ToolTipText = dtgA.Text
dtgA.Col = 2
comXh.Text = dtgA.Text
comXh.ToolTipText = dtgA.Text
dtgA.Col = 3
txtSl.Text = dtgA.Text
End Sub

Private Sub dtgBao_Click()
Dim tt As String
Dim liD As Long
On Error Resume Next
dtgBao.Col = 11
txtTl.Text = dtgBao.Text
dtgBao.Col = 12
txtDj.Text = dtgBao.Text
dtgBao.Col = 16
liD = dtgBao.Text
tt = "select * from xunJiaMxView where lid=" & liD
adoGx.Close
adoGx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set dtgMa.DataSource = adoGx
End Sub

Private Sub dtgBao_RowColChange()
Dim tt As String
Dim liD As Long
On Error Resume Next
dtgBao.Col = 11
txtTl.Text = dtgBao.Text
dtgBao.Col = 12
txtDj.Text = dtgBao.Text
dtgBao.Col = 16
liD = dtgBao.Text
tt = "select * from xunJiaMxView where lid=" & liD
adoGx.Close
adoGx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set dtgMa.DataSource = adoGx
End Sub

Private Sub dtpYf_CloseUp()
txtYrq.Text = dtpYf.Value
End Sub


Private Sub Form_Click()
frmYm.Visible = False
frmFF.Visible = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = 76 Then
    If mod1.Kyj = True Then
        If frmYJ.Visible = False Then
            frmYJ.Visible = True
            If (lblLc.Caption = 2 Or lblLc.Caption = 3) And lblLcRen.Caption = mod1.DName Then
                'txtYj.Locked = False
                txtTcBe.Locked = False
            Else
                txtYJ.Locked = True
                txtTcBe.Locked = True
            End If
        Else
            frmYJ.Visible = False
        End If
    Else
        frmYJ.Visible = False
    End If
    
End If
End Sub

Private Sub Form_Load()
txtFbnr.Locked = True
txtFbje.Locked = True
Set adoWb = CreateObject("adodb.recordset")
Set adoLj = CreateObject("adodb.recordset")
Set adoOid = CreateObject("adodb.recordset")
Set adoBx = CreateObject("adodb.recordset")
Set adoGx = CreateObject("adodb.recordset")
Set AdoKh = CreateObject("adodb.recordset")
Set adoYj = CreateObject("adodb.recordset")
Set adoFk = CreateObject("adodb.recordset")
dtgYJ.ColWidth(0) = 300
dtgYJ.ColWidth(3) = 0
dt3.Value = mod1.DQda
dt4.Value = mod1.DQda
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0
dtgWb.ColWidth(0) = 300
dtgWb.ColWidth(4) = 3500
dtgWb.ColWidth(11) = 0
dtgWb.ColWidth(13) = 0
dtgWb.ColWidth(14) = 0
dtgWb.ColWidth(15) = 0
dtgWb.ColWidth(16) = 0
dtgWb.ColWidth(17) = 0
dtgWb.ColWidth(18) = 0
dtgWb.ColWidth(6) = 900
dtgWb.ColWidth(7) = 900
dtgWb.ColWidth(9) = 900
dtgWb.ColWidth(3) = 1815
dtgWb.ColWidth(10) = 1665
dtgWb.Left = 0
dtgWb.Top = 0
dtgFk.ColWidth(0) = 300
dtgFk.ColWidth(4) = 0
dtgFk.ColWidth(5) = 0
frmTime.BorderStyle = 0
frmDx.BorderStyle = 0
frmNb.BorderStyle = 0
dtgLj.ColWidth(0) = 300
dtgLj.ColWidth(4) = 3500
dtgLj.ColWidth(11) = 0
dtgLj.ColWidth(13) = 0
dtgLj.ColWidth(14) = 0
dtgLj.ColWidth(15) = 0
dtgLj.ColWidth(16) = 0
dtgLj.ColWidth(17) = 0
dtgLj.ColWidth(18) = 0
dtgLj.ColWidth(6) = 900
dtgLj.ColWidth(7) = 900
dtgLj.ColWidth(9) = 900
dtgLj.ColWidth(3) = 1815
dtgLj.ColWidth(10) = 1665
dtgLj.Left = 0
dtgLj.Top = 0

dtgA.ColWidth(0) = 300
dtgA.ColWidth(2) = 2000
dtgA.ColWidth(3) = 700
dtgA.ColWidth(4) = 0

dtgBao.ColWidth(0) = 300
dtgBao.ColWidth(8) = 2000
dtgBao.ColWidth(15) = 0
dtgBao.ColWidth(16) = 0
dtgBao.ColWidth(17) = 0
dtgBao.Left = 0
dtgBao.Top = 0
dtgMa.ColWidth(0) = 300
dtgMa.ColWidth(8) = 2000
dtgMa.ColWidth(15) = 0
dtgMa.ColWidth(16) = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblKT.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim tt As String
Dim ii As Integer
On Error Resume Next
If MDI.Cq = False Then
If lblLc.Caption = 0 Then
    ii = MsgBox("该单没有保存,是否将其撤消?", vbQuestion + vbYesNo, "请注意!")
    If ii = vbYes Then
        tt = "delete from baoJiaD where baoid=" & Val(lblBaoId.Caption)
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
        tt = "update xunJiaD set baoid = null where bid=" & Val(lblBid.Caption)
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Else
        Exit Sub
    End If
End If
    frmWbxjB.Visible = False
    If frmGxBiao.Visible = True Then
        frmGxBiao.Enabled = True
        frmGxBiao.ZOrder 0
        mod1.BTZ = 36
    ElseIf Dialog.Visible = True Then
        Dialog.ZOrder 0
        Dialog.Enabled = True
    End If
    Cancel = True
End If
End Sub


Private Sub opt1_Click()
dtgWb.Visible = True
dtgLj.Visible = False
End Sub


Private Sub opt2_Click()
dtgLj.Visible = True
dtgWb.Visible = False
End Sub


Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub tabGc_Click(PreviousTab As Integer)
dtgWb.Visible = False
dtgLj.Visible = False
txtDXNR.Visible = False
dtgBao.Visible = False
dtgMa.Visible = False

Select Case tabGc.Tab
Case 0
    dtgWb.Visible = True
Case 1
    dtgLj.Visible = True
Case 2
    txtDXNR.Visible = True
Case 3
    dtgBao.Visible = True
    dtgMa.Visible = True
End Select
End Sub

Private Sub txtHg_DblClick()
frmFF.Visible = True
End Sub

Private Sub txtHg_LostFocus()
txtYhg.Text = txtHg.Text
End Sub


Private Sub txtXm2_LostFocus()
txtCb.Text = Val(txtRgf.Text) + Val(txtClf.Text) + Val(txtClcb.Text) + Val(txtYf.Text) + Val(txtYJ.Text) + Val(txtXm2.Text) + Val(txtFbje.Text)
End Sub

Private Sub txtYf_LostFocus()
txtCb.Text = Val(txtRgf.Text) + Val(txtClf.Text) + Val(txtClcb.Text) + Val(txtYf.Text) + Val(txtYJ.Text) + Val(txtXm2.Text)
End Sub


Private Sub txtYJ_DblClick()
frmYm.Visible = True
End Sub

Private Sub txtYj_LostFocus()
txtCb.Text = Val(txtRgf.Text) + Val(txtClf.Text) + Val(txtClcb.Text) + Val(txtYf.Text) + Val(txtYJ.Text) + Val(txtXm2.Text)
End Sub


