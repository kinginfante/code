VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGxbjB 
   Caption         =   "购销报价单"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.TextBox txtBz 
      Height          =   285
      Left            =   9810
      TabIndex        =   114
      Top             =   7200
      Width           =   3315
   End
   Begin VB.CommandButton cmdCong 
      BackColor       =   &H00C0FFC0&
      Caption         =   "重新评审"
      Height          =   1125
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   113
      Top             =   8010
      Width           =   375
   End
   Begin VB.CommandButton cmdPje 
      Caption         =   "评审建议"
      Height          =   1095
      Left            =   9480
      TabIndex        =   112
      Top             =   8040
      Width           =   345
   End
   Begin VB.TextBox txtFbnr 
      Height          =   375
      Left            =   11760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   110
      Top             =   6720
      Width           =   3285
   End
   Begin VB.Frame frmGD 
      Caption         =   "项目费用分类"
      Height          =   2625
      Left            =   3240
      TabIndex        =   90
      Top             =   3210
      Visible         =   0   'False
      Width           =   6945
      Begin VB.Frame Frame2 
         Height          =   825
         Left            =   30
         TabIndex        =   96
         Top             =   1290
         Width           =   6885
         Begin VB.CommandButton cmdGdel 
            Caption         =   "删除"
            Height          =   255
            Left            =   5880
            TabIndex        =   109
            Top             =   480
            Width           =   885
         End
         Begin VB.CommandButton cmdGAdd 
            Caption         =   "添加"
            Height          =   255
            Left            =   4980
            TabIndex        =   108
            Top             =   480
            Width           =   795
         End
         Begin VB.TextBox txtRl 
            Height          =   270
            Left            =   2250
            TabIndex        =   105
            Top             =   480
            Width           =   915
         End
         Begin VB.TextBox txtQdj 
            Height          =   270
            Left            =   750
            TabIndex        =   103
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtGDNR 
            Height          =   270
            Left            =   4770
            TabIndex        =   101
            Top             =   150
            Width           =   2025
         End
         Begin VB.OptionButton optGDC 
            Caption         =   "其它"
            Height          =   195
            Left            =   3960
            TabIndex        =   100
            Top             =   180
            Width           =   675
         End
         Begin VB.OptionButton optGDB 
            Caption         =   "春节(年会吃饭)"
            Height          =   180
            Left            =   2280
            TabIndex        =   99
            Top             =   180
            Width           =   1605
         End
         Begin VB.OptionButton optGDA 
            Caption         =   "中秋(月饼券)"
            Height          =   195
            Left            =   750
            TabIndex        =   98
            Top             =   180
            Width           =   1545
         End
         Begin MSComCtl2.DTPicker dtpGD 
            Height          =   255
            Left            =   3810
            TabIndex        =   107
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "yyyyy"
            Format          =   108134403
            UpDown          =   -1  'True
            CurrentDate     =   38943
         End
         Begin VB.Label Label15 
            Caption         =   "年份:"
            Height          =   195
            Left            =   3330
            TabIndex        =   106
            Top             =   510
            Width           =   525
         End
         Begin VB.Label Label11 
            Caption         =   "人数:"
            Height          =   165
            Left            =   1650
            TabIndex        =   104
            Top             =   540
            Width           =   705
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   6870
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Label Label10 
            Caption         =   "单价:"
            Height          =   225
            Left            =   120
            TabIndex        =   102
            Top             =   540
            Width           =   585
         End
         Begin VB.Label Label9 
            Caption         =   "类别:"
            Height          =   195
            Left            =   120
            TabIndex        =   97
            Top             =   180
            Width           =   555
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgGD 
         Height          =   1065
         Left            =   60
         TabIndex        =   95
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   1879
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox txtXm 
         Height          =   270
         Left            =   3270
         TabIndex        =   94
         Top             =   2250
         Width           =   1515
      End
      Begin VB.TextBox txtGd 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   92
         Top             =   2250
         Width           =   1155
      End
      Begin VB.Label Label8 
         Caption         =   "活动费用"
         Height          =   255
         Left            =   2400
         TabIndex        =   93
         Top             =   2280
         Width           =   885
      End
      Begin VB.Label Label7 
         Caption         =   "固定费用"
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   2310
         Width           =   825
      End
   End
   Begin VB.Frame frmFB 
      Height          =   465
      Left            =   9060
      TabIndex        =   85
      Top             =   6690
      Width           =   1545
      Begin VB.TextBox txtFbje 
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   120
         Width           =   885
      End
      Begin VB.Label Label6 
         Caption         =   "分包"
         Height          =   255
         Left            =   150
         TabIndex        =   86
         Top             =   150
         Width           =   615
      End
   End
   Begin VB.TextBox txtSl 
      Height          =   315
      Left            =   9660
      TabIndex        =   84
      Top             =   5310
      Width           =   1515
   End
   Begin VB.CommandButton cmdD 
      Caption         =   "删除"
      Height          =   315
      Left            =   14400
      TabIndex        =   82
      Top             =   5310
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Frame frmFF 
      BackColor       =   &H00C0FFC0&
      Caption         =   "付款方式"
      Height          =   2235
      Left            =   1350
      TabIndex        =   70
      Top             =   4770
      Visible         =   0   'False
      Width           =   6585
      Begin VB.Frame frmFk 
         BackColor       =   &H00C0FFC0&
         Height          =   555
         Left            =   0
         TabIndex        =   73
         Top             =   1680
         Width           =   5955
         Begin VB.TextBox txtYrq 
            Height          =   300
            Left            =   990
            Locked          =   -1  'True
            TabIndex        =   75
            Top             =   150
            Width           =   1005
         End
         Begin VB.TextBox txtEd 
            Height          =   270
            Left            =   3570
            TabIndex        =   74
            Top             =   150
            Width           =   795
         End
         Begin MSComCtl2.DTPicker dtpYf 
            Height          =   315
            Left            =   990
            TabIndex        =   76
            Top             =   150
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   16711680
            CalendarTrailingForeColor=   8454016
            Format          =   107675649
            CurrentDate     =   38797
         End
         Begin VB.Label lblFid 
            BackColor       =   &H00C0FFC0&
            Caption         =   "lblFid"
            Height          =   165
            Left            =   4920
            TabIndex        =   80
            Top             =   240
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label Label37 
            BackColor       =   &H00C0FFC0&
            Caption         =   "%"
            Height          =   255
            Left            =   4470
            TabIndex        =   79
            Top             =   180
            Width           =   435
         End
         Begin VB.Label Label34 
            BackColor       =   &H00C0FFC0&
            Caption         =   "收款额度"
            Height          =   255
            Left            =   2730
            TabIndex        =   78
            Top             =   180
            Width           =   795
         End
         Begin VB.Label Label33 
            BackColor       =   &H00C0FFC0&
            Caption         =   "应付日期"
            Height          =   285
            Left            =   60
            TabIndex        =   77
            Top             =   180
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "添加"
         Height          =   375
         Left            =   6000
         TabIndex        =   72
         Top             =   1440
         Width           =   525
      End
      Begin VB.CommandButton cmdDe 
         Caption         =   "删除"
         Height          =   375
         Left            =   6000
         TabIndex        =   71
         Top             =   1830
         Width           =   525
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgFk 
         Height          =   1575
         Left            =   0
         TabIndex        =   81
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
      Height          =   345
      Left            =   9120
      TabIndex        =   66
      Top             =   7590
      Width           =   3465
      Begin VB.OptionButton optLa 
         Caption         =   "增值发票"
         Height          =   195
         Left            =   120
         TabIndex        =   69
         Top             =   90
         Width           =   1065
      End
      Begin VB.OptionButton optLb 
         Caption         =   "商业发票"
         Height          =   195
         Left            =   1200
         TabIndex        =   68
         Top             =   90
         Width           =   1065
      End
      Begin VB.OptionButton optLc 
         Caption         =   "服务发票"
         Height          =   195
         Left            =   2280
         TabIndex        =   67
         Top             =   90
         Width           =   1065
      End
   End
   Begin VB.Frame frmYM 
      BackColor       =   &H8000000D&
      Caption         =   "奖金预计支付情况"
      Height          =   2055
      Left            =   3210
      TabIndex        =   57
      Top             =   5610
      Width           =   4665
      Begin VB.TextBox txtYED 
         Height          =   285
         Left            =   930
         TabIndex        =   61
         Top             =   1620
         Width           =   645
      End
      Begin VB.TextBox txtYingFu 
         Height          =   270
         Left            =   2850
         TabIndex        =   60
         Top             =   1620
         Width           =   1035
      End
      Begin VB.CommandButton cmdYadd 
         Caption         =   "添加"
         Height          =   315
         Left            =   3960
         TabIndex        =   59
         Top             =   780
         Width           =   585
      End
      Begin VB.CommandButton cmdYdel 
         Caption         =   "删除"
         Height          =   285
         Left            =   3960
         TabIndex        =   58
         Top             =   1170
         Width           =   585
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgYJ 
         Height          =   1275
         Left            =   90
         TabIndex        =   62
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
      Begin VB.Label Label28 
         BackColor       =   &H8000000D&
         Caption         =   "收款额度"
         Height          =   255
         Left            =   90
         TabIndex        =   65
         Top             =   1650
         Width           =   825
      End
      Begin VB.Label Label30 
         BackColor       =   &H8000000D&
         Caption         =   "%"
         Height          =   255
         Left            =   1680
         TabIndex        =   64
         Top             =   1650
         Width           =   195
      End
      Begin VB.Label Label31 
         BackColor       =   &H8000000D&
         Caption         =   "支付金额"
         Height          =   225
         Left            =   1980
         TabIndex        =   63
         Top             =   1650
         Width           =   915
      End
   End
   Begin VB.Frame frmYj 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   4440
      TabIndex        =   52
      Top             =   8370
      Visible         =   0   'False
      Width           =   2025
      Begin VB.TextBox txtYj 
         Height          =   285
         Left            =   900
         TabIndex        =   54
         Top             =   30
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtTcBe 
         Height          =   285
         Left            =   900
         TabIndex        =   53
         Text            =   "6"
         Top             =   360
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lblYj 
         Caption         =   "奖金"
         Height          =   225
         Left            =   360
         TabIndex        =   56
         Top             =   60
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblTcBe 
         Caption         =   "提成比例"
         Height          =   195
         Left            =   0
         TabIndex        =   55
         Top             =   420
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.ComboBox comKhmc 
      Height          =   300
      Left            =   1200
      TabIndex        =   51
      Top             =   7800
      Width           =   3165
   End
   Begin VB.CommandButton cmdHt 
      Caption         =   "合同评审单"
      Height          =   285
      Left            =   14160
      TabIndex        =   49
      Top             =   7260
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtXm1 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   6780
      Width           =   1425
   End
   Begin VB.TextBox txtXm2 
      Height          =   270
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   7290
      Width           =   1425
   End
   Begin VB.TextBox txtHg 
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   8250
      Width           =   1425
   End
   Begin VB.TextBox txtYhg 
      Height          =   285
      Left            =   7560
      TabIndex        =   41
      ToolTipText     =   "此处由工程部填入"
      Top             =   8745
      Width           =   1425
   End
   Begin VB.TextBox txtCb 
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   7755
      Width           =   1425
   End
   Begin VB.TextBox txtClcb 
      Height          =   315
      Left            =   3180
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   8715
      Width           =   1125
   End
   Begin VB.TextBox txtYf 
      Height          =   315
      Left            =   1200
      TabIndex        =   36
      Top             =   8715
      Width           =   885
   End
   Begin VB.TextBox txtDj 
      Height          =   345
      Left            =   12000
      TabIndex        =   34
      Top             =   5280
      Width           =   1275
   End
   Begin VB.CommandButton cmdGx 
      Caption         =   "更新"
      Height          =   315
      Left            =   13560
      TabIndex        =   33
      Top             =   5310
      Width           =   780
   End
   Begin VB.CommandButton cmdXjd 
      BackColor       =   &H00C0FFC0&
      Caption         =   "询价单"
      Height          =   315
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7620
      Width           =   1065
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   555
      Left            =   13410
      Picture         =   "frmGxbjB.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7620
      Width           =   615
   End
   Begin VB.CommandButton cmdBack 
      Height          =   405
      Left            =   14760
      Picture         =   "frmGxbjB.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "返回"
      Top             =   8760
      Width           =   465
   End
   Begin VB.CommandButton cmdQm 
      Caption         =   "cmdQm"
      Height          =   345
      Index           =   0
      Left            =   9870
      TabIndex        =   13
      Top             =   8280
      Width           =   945
   End
   Begin VB.Frame frmHide 
      Caption         =   "frmHid"
      Height          =   1455
      Left            =   2940
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Label lblQy 
         Caption         =   "lblQy"
         Height          =   195
         Left            =   3450
         TabIndex        =   89
         Top             =   1110
         Width           =   1065
      End
      Begin VB.Label lblBm 
         Caption         =   "lblBm"
         Height          =   225
         Left            =   720
         TabIndex        =   88
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label lblLcou 
         Caption         =   "lblLcou"
         Height          =   255
         Left            =   1860
         TabIndex        =   12
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label lblYwy 
         Caption         =   "lblYwy"
         Height          =   285
         Left            =   3540
         TabIndex        =   11
         Top             =   450
         Width           =   765
      End
      Begin VB.Label lblUid 
         Caption         =   "lblUid"
         Height          =   255
         Left            =   3750
         TabIndex        =   10
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblFwid 
         Caption         =   "lblFwid"
         Height          =   255
         Left            =   1860
         TabIndex        =   9
         Top             =   450
         Width           =   1275
      End
      Begin VB.Label lblLcUid 
         Caption         =   "lblLcUid"
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   930
         Width           =   885
      End
      Begin VB.Label lblLcRen 
         Caption         =   "lblLcRen"
         Height          =   285
         Left            =   150
         TabIndex        =   7
         Top             =   420
         Width           =   795
      End
      Begin VB.Label lblNlb 
         Caption         =   "lblNlb"
         Height          =   225
         Left            =   1920
         TabIndex        =   6
         Top             =   810
         Width           =   645
      End
      Begin VB.Label lblLc 
         Caption         =   "lblLc"
         Height          =   315
         Left            =   1050
         TabIndex        =   5
         Top             =   630
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdRight 
      Caption         =   ">"
      Height          =   285
      Left            =   14760
      TabIndex        =   3
      Top             =   8250
      Width           =   465
   End
   Begin VB.CommandButton cmdLeft 
      Caption         =   "<"
      Height          =   285
      Left            =   14280
      TabIndex        =   2
      Top             =   8250
      Width           =   465
   End
   Begin VB.CommandButton cmdMod 
      Height          =   405
      Left            =   13800
      Picture         =   "frmGxbjB.frx":076C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "修改"
      Top             =   8760
      Width           =   465
   End
   Begin VB.CommandButton cmdSave 
      Height          =   405
      Left            =   14280
      Picture         =   "frmGxbjB.frx":0A76
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "保存"
      Top             =   8760
      Width           =   465
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBao 
      Height          =   5175
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   9128
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
   Begin MSDataListLib.DataCombo comXmmc 
      Height          =   330
      Left            =   1200
      TabIndex        =   16
      Top             =   7275
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   582
      _Version        =   393216
      Text            =   ""
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgMa 
      Height          =   885
      Left            =   -30
      TabIndex        =   26
      Top             =   5670
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
   Begin VB.Label Label19 
      Caption         =   "备注"
      Height          =   225
      Left            =   9120
      TabIndex        =   115
      Top             =   7260
      Width           =   585
   End
   Begin VB.Label Label18 
      Caption         =   "分包内容"
      Height          =   225
      Left            =   10800
      TabIndex        =   111
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "数量"
      Height          =   195
      Left            =   9060
      TabIndex        =   83
      Top             =   5370
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "客户名称"
      Height          =   225
      Left            =   210
      TabIndex        =   50
      Top             =   7860
      Width           =   795
   End
   Begin VB.Label lblhtbh 
      Caption         =   "lblhtbh"
      Height          =   255
      Left            =   13350
      TabIndex        =   48
      Top             =   7020
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label26 
      Caption         =   "已发生项目费用"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6000
      TabIndex        =   47
      Top             =   6810
      Width           =   1305
   End
   Begin VB.Label Label27 
      Caption         =   "预留项目费用"
      Height          =   255
      Left            =   6180
      TabIndex        =   46
      Top             =   7305
      Width           =   1125
   End
   Begin VB.Label Label12 
      Caption         =   "优惠价"
      Height          =   255
      Left            =   6720
      TabIndex        =   43
      ToolTipText     =   "此处由工程部填入"
      Top             =   8775
      Width           =   555
   End
   Begin VB.Label Label16 
      Caption         =   "材料成本"
      Height          =   225
      Left            =   2220
      TabIndex        =   39
      Top             =   8790
      Width           =   765
   End
   Begin VB.Label Label17 
      Caption         =   "运费"
      Height          =   285
      Left            =   570
      TabIndex        =   38
      Top             =   8775
      Width           =   435
   End
   Begin VB.Label Label3 
      Caption         =   "单价"
      Height          =   225
      Left            =   11340
      TabIndex        =   35
      Top             =   5340
      Width           =   525
   End
   Begin VB.Label lblZl 
      Caption         =   "Label19"
      ForeColor       =   &H00C000C0&
      Height          =   315
      Left            =   1260
      TabIndex        =   32
      Top             =   6810
      Width           =   1155
   End
   Begin VB.Label lblzlZ 
      Caption         =   "性质"
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   630
      TabIndex        =   31
      Top             =   6810
      Width           =   435
   End
   Begin VB.Label Label2 
      Caption         =   "采购成本"
      Height          =   225
      Left            =   120
      TabIndex        =   28
      Top             =   5340
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "成本总额"
      Height          =   255
      Left            =   6540
      TabIndex        =   27
      Top             =   7785
      Width           =   735
   End
   Begin VB.Label Label24 
      Caption         =   "项目名称"
      Height          =   255
      Left            =   270
      TabIndex        =   25
      Top             =   7305
      Width           =   795
   End
   Begin VB.Label lblBid 
      Caption         =   "lblBid"
      Height          =   225
      Left            =   3390
      TabIndex        =   24
      Top             =   5250
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblQM 
      Caption         =   "lblQm"
      Height          =   225
      Index           =   0
      Left            =   9930
      TabIndex        =   23
      Top             =   8010
      Width           =   915
   End
   Begin VB.Label lblTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   0
      Left            =   9840
      TabIndex        =   22
      Top             =   8700
      Width           =   945
   End
   Begin VB.Label lblBaoId 
      Caption         =   "lblBaoId"
      Height          =   285
      Left            =   6450
      TabIndex        =   21
      Top             =   5310
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label lblOid 
      Caption         =   "lblOid"
      Height          =   285
      Left            =   5130
      TabIndex        =   20
      Top             =   5310
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label13 
      Caption         =   "对外报价"
      Height          =   255
      Left            =   6540
      TabIndex        =   19
      Top             =   8280
      Width           =   735
   End
   Begin VB.Label Label14 
      Caption         =   "编号"
      Height          =   285
      Left            =   540
      TabIndex        =   18
      Top             =   8295
      Width           =   435
   End
   Begin VB.Label lblBh 
      BackColor       =   &H80000014&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      Height          =   315
      Left            =   1200
      TabIndex        =   17
      Top             =   8280
      Width           =   3135
   End
End
Attribute VB_Name = "frmGxbjB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adoOid As Object
Public adoGx As Object
Public adoBx As Object
Public adoGD As Object
Public adoHGD As Object

Public adoYj As Object '佣金表
Public adoFk As Object '付款表
Dim AdoKh As Object

Dim Lb As String
Private Sub cmdAdd_Click()
On Error Resume Next
Set mod1.cmd = CreateObject("adodb.command")
mod1.cmd.ActiveConnection = mod1.cc
mod1.cmd.CommandText = "BFkAdd"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@rq") = txtYrq.Text
mod1.cmd.Parameters("@yingfJe") = Round(Val(txtHtze.Text) * Val(txtEd.Text) / 100, 2)
mod1.cmd.Parameters("@baoid") = Val(lblBaoId.Caption)
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
Me.Visible = False
If frmGxBiao.Visible = True Then
    frmGxBiao.Enabled = True
    frmGxBiao.ZOrder 0
ElseIf Dialog.Visible = True Then
    Dialog.ZOrder 0
    Dialog.Enabled = True
End If
mod1.BTZ = 36
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

Private Sub cmdGAdd_Click()
Dim tt As String

Dim Xg As Single
On Error Resume Next
Xg = 0
If optGDA.Value = False And optGDB.Value = False And optGDC.Value = False Then
    Exit Sub
End If
If Val(txtQdj.Text) = 0 Or Val(txtRl.Text) = 0 Then
    Exit Sub
End If
Xg = Val(txtQdj.Text) * Val(txtRl.Text)
If optGDC.Value = True Then
    Lb = txtGDNR.Text
End If
'tt = "update xmgd set lb='" & LB & "',nd='" & dtpGD.Value & "',qdj=" & Val(txtQdj.Text) & ",rl=" & Val(txtRl.Text) & ",xg=" & Xg & ",baoid=" & Val(lblBaoId.Caption)
tt = "insert into xmgd (lb,nd,qdj,rl,xg,baoid) values('" & Lb & "','" & dtpGD.Value & "'," & Val(txtQdj.Text) & "," & Val(txtRl.Text) & "," & Xg & "," & Val(lblBaoId.Caption) & ")"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
adoGD.Requery
Set dtgGD.DataSource = adoGD
optGDA.Value = False
optGDB.Value = False
optGDC.Value = False
txtGDNR.Text = ""
txtQdj.Text = ""
txtRl.Text = ""


End Sub

Private Sub cmdGdel_Click()
Dim ii As Integer
Dim tt As String
Dim Gid As Long
On Error Resume Next
dtgGD.Col = 8
Gid = dtgGD.Text
ii = MsgBox("是否删除此条记录?", vbInformation + vbYesNo, "询问")
If ii = vbYes Then
    tt = "delete from xmGd where gid=" & Gid
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    adoGD.Requery
    Set dtgGD.DataSource = adoGD
End If


End Sub

Private Sub cmdGx_Click()
Dim ii As Integer
Dim CB As Long
Dim liD As Long
Dim tt As String
Dim XCB As Long
On Error Resume Next
'If Val(txtDj.Text) = 0 Then Exit Sub
If Val(txtSl.Text) = 0 Then
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
    mod1.cmd.Parameters("@sl") = Val(txtSl.Text)
    mod1.cmd.Parameters("@lid") = liD
    mod1.cmd.Execute
    'txtHg.Text = Val(txtHg.Text) + mod1.CMD.Parameters("@hg").Value
    Set cmd = Nothing
    adoBx.Requery
    Set dtgBao.DataSource = adoBx
    '计算总费用
    CB = 0
    adoBx.MoveFirst
    Do While Not adoBx.EOF
        CB = CB + adoBx.Fields("合计").Value
        adoBx.MoveNext
    Loop
    txtHg.Text = CB
    txtYhg.Text = txtHg.Text
    'txtCb.Text = Val(txtClcb.Text) + Val(txtYf.Text) + Val(txtYj.Text)
    '更新相应询价明细中的数量
    tt = "update XunJiaMx set sl=" & Val(txtSl.Text) & ",hg=dj*" & Val(txtSl.Text) & " where lid=" & liD
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    '更新相应询价单中的金额
    tt = "select sum(hg) as hg from xunjiamx where bid=" & Val(lblBid.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'    XCB = 0
'    Do While Not mod1.HTP.EOF
'        XCB = XCB + mod1.HTP.Fields("hg").Value
'        mod1.HTP.MoveNext
'    Loop
    XCB = mod1.HTP.Fields("hg").Value
    frmGXBj.txtHg.Text = XCB
    frmGXBj.txtYhg.Text = XCB
    tt = "update xunjiaD set hg=" & XCB & ",yhg=" & XCB & " where bid=" & Val(lblBid.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    tt = "update baojiaD set clcb=" & XCB & ",hg=" & XCB & "+clf+rgf+yf+ylxm where baoid=" & Val(lblBaoId.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    txtDj.Text = ""
    txtTl.Text = ""
    txtClcb.Text = XCB
    txtCb.Text = Val(txtClcb.Text) + Val(txtYf.Text)
    adoGx.Requery
    Set dtgMa.DataSource = adoGx
    Call cmdSave_Click
End Sub

Private Sub cmdHt_Click()
Dim tt As String
On Error Resume Next
Dim oo As Integer
Dim xZ As String
Dim Hid As Long
Dim FPLX As String
On Error Resume Next

If optLa.Value = True Then
    FPLX = "增值发票"
ElseIf optLb.Value = True Then
    FPLX = "商业发票"
ElseIf optLc.Value = True Then
    FPLX = "服务发票"
End If

If lblHtbh.Caption = "" Then
    Exit Sub
End If
tt = "select hid from htping where htbh='" & lblHtbh.Caption & "' and delf=1"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText

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
        mod1.cmd.Parameters("@FBNR") = txtFbnr.Text
        If Val(txtFbje.Text) > 0 Then
            mod1.cmd.Parameters("@htxz") = "工程分包"
        Else
            mod1.cmd.Parameters("@htxz") = adoBx.Fields("品种").Value
        End If
        mod1.cmd.Parameters("@fbje") = Val(txtFbje.Text)
        mod1.cmd.Parameters("@htze") = txtYhg.Text
        If (lblZl.Caption = "维保" And Val(txtYhg.Text) > 50000) Or Val(txtYhg.Text) > 100000 Or Val(txtFbje.Text) > 0 Then
            mod1.cmd.Parameters("@nlb") = 62
            mod1.cmd.Parameters("@lcou") = 6
        Else
            mod1.cmd.Parameters("@nlb") = 63
            mod1.cmd.Parameters("@lcou") = 5
        End If
        mod1.cmd.Parameters("@baoid") = Val(lblBaoId.Caption)
        mod1.cmd.Parameters("@Hrq") = Format(mod1.DQda, "yyyymmdd")
        mod1.cmd.Parameters("@lc") = 0
        mod1.cmd.Parameters("@lcren") = mod1.DName
        mod1.cmd.Parameters("@lcuid") = mod1.DHid
        mod1.cmd.Parameters("@cbze") = Val(txtCb.Text)
        mod1.cmd.Parameters("@rgf") = 0
        mod1.cmd.Parameters("@clf") = 0
        mod1.cmd.Parameters("@yf") = Val(txtYf.Text)
        mod1.cmd.Parameters("@yj") = Val(txtYJ.Text)
        mod1.cmd.Parameters("@mon") = 0
        mod1.cmd.Parameters("@clcb") = Val(txtClcb.Text)
        mod1.cmd.Parameters("@xmfy") = Val(txtXm2.Text)
        mod1.cmd.Parameters("@qy") = mod1.Qy
        mod1.cmd.Parameters("@tcbe") = Val(txtTcBe.Text)
        mod1.cmd.Parameters("@fplx") = FPLX
        mod1.cmd.Parameters("@htqy") = "1999-1-1"
        mod1.cmd.Parameters("@htqy1") = "1999-1-1"
        mod1.cmd.Parameters("@bz") = Trim(txtBz.Text)
        If adoBx.Fields("品种").Value Or adoBx.Fields("品种").Value = "购销" Then
            mod1.cmd.Parameters("@xzCh") = "LP"
        ElseIf adoBx.Fields("品种").Value = "产品" Then
            mod1.cmd.Parameters("@xzCh") = "CP"
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
    If (lblZl.Caption = "维保" And Val(txtYhg.Text) > 50000) Or Val(txtYhg.Text) > 100000 Or Val(txtFbje.Text) > 0 Then
        Call modHt.HtLcBut(62)
    Else
        Call modHt.HtLcBut(63)
    End If
        frmWbNew.cmdSave.Enabled = True
        frmWbNew.cmdMod.Enabled = False
    

Else
        mod1.BTZ = 6
        Call modHt.NewQing
    
        Call modHt.NewBound(mod1.HTP.Fields("hid").Value)
        frmWbNew.Visible = True
End If
frmGxbjB.Visible = False
End Sub

Private Sub cmdMod_Click()
If lblLcRen.Caption <> mod1.DName Or lblLcUid.Caption <> mod1.DHid Then
    Exit Sub
End If

If lblLc.Caption = 1 Then
    cmdGx.Enabled = True
    comKhmc.Locked = False
    txtYf.Locked = False
    txtXm2.Locked = False
    'txtHg.Locked = False
    txtYhg.Locked = False
    txtDj.Locked = False
    txtSl.Locked = False
    optLa.Enabled = True
    optLb.Enabled = True
    optLc.Enabled = True
    cmdMod.Enabled = False
    cmdSave.Enabled = True
ElseIf lblLc.Caption = 2 Then
    cmdGx.Enabled = True
    txtYf.Locked = False
    txtFbje.Locked = False
    txtXm2.Locked = False
    'txtHg.Locked = False
    txtYhg.Locked = False
    txtDj.Locked = False
    txtSl.Locked = False
    optLa.Enabled = True
    optLb.Enabled = True
    optLc.Enabled = True
    txtTcBe.Locked = False
    cmdMod.Enabled = False
    cmdSave.Enabled = True
End If

If comKhmc.Text = "" Then
    comKhmc.Locked = False
End If
txtFbnr.Locked = True
txtFbje.Locked = True

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
    If adoBx.Fields("品种").Value = "产品" Then
        Set mod1.report = mod1.crapp.OpenReport(App.Path & "\Bjdgx.rpt")
    ElseIf adoBx.Fields("pz").Value = "零配件" Then
        Set mod1.report = mod1.crapp.OpenReport(App.Path & "\BjdgxCP.rpt")
    End If
     'Set mod1.report = mod1.crapp.OpenReport(App.Path & "\tt.rpt")
    Set mod1.table = mod1.report.Database.Tables
    Set mod1.cProp = mod1.table.Item(1).ConnectionProperties
    mod1.cProp.Item("Password") = "guyonghui"
    mod1.report.SQLQueryString = "Select * from bjdgx  where baoid=" & Val(lblBaoId.Caption)
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
Dim oo As Integer
Dim Tywy As String '单子流转到下一人的姓名
Dim Tuid As String
Dim Oywy As String '原来流转人的名字
Dim Ouid As String '原来流转人的工号

On Error Resume Next


'If cmdQm(Index).Caption <> "" Then Exit Sub
If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If

If comKhmc.Text = "" Then
    MsgBox "请选择相应签约客户名称!"
    cmdSave.Enabled = True
    Exit Sub
End If

If optLa.Value = False And optLb.Value = False And optLc.Value = False Then
    MsgBox "请确认开票类型!"
        cmdSave.Enabled = True
    Exit Sub
End If
'If Index = 0 And cmdSave.Enabled = True And lblLc.Caption = 0 Then


If lblLc.Caption = 2 And txtTcBe.Text = "" Then
    MsgBox "请键入提成比例!"
    cmdSave.Enabled = True
    Exit Sub
End If
If Index + 1 <> lblLc.Caption Then '不能在不相干的位置上乱点

    Exit Sub
End If

If lblLcUid.Caption <> mod1.DHid Then
    MsgBox "此处应由" & lblLcRen.Caption & "签字! 请您不要再点"
    Exit Sub
End If

If comKhmc.Text = "" Then
    MsgBox "请选择相应签约客户名称!"
    cmdSave.Enabled = True
    
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
End If

Dim Zi As Integer
Zi = MsgBox("是否确认签字?", vbYesNo)
If Zi = vbNo Then Exit Sub


    
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
'先计算成本总额
txtCb.Text = Val(txtYf.Text) + Val(txtClcb.Text) + Val(txtXm2.Text) + Val(txtFbje.Text) + Val(txtGd.Text)

'tt = "update baoJiaD set hg=" & Val(txtHg.Text) & ",yhg=" & Val(txtYhg.Text) & " where baoid=" & Val(lblBaoId.Caption)
tt = "update baoJiaD set bhg=" & Val(txtHg.Text) & ",yhg=" & Val(txtYhg.Text) & ",yj=" & Val(txtYJ.Text) & ",ylxm=" & Val(txtXm2.Text) & _
", hg=" & Val(txtCb.Text) & ", yf=" & Val(txtYf.Text) & " ,khmc='" & comKhmc.Text & "',khdh='" & comKhmc.ToolTipText & _
"',tcbe=" & Val(txtTcBe.Text) & ",fplx='" & FPLX & "' where baoid=" & Val(lblBaoId.Caption)
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
    Call modBJD.OpenBJAN(0)
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


Private Sub cmdXjd_Click()
Dim tt As String
On Error Resume Next
mod1.BTZ = 36
frmGXBj.Visible = False
If frmGXBj.lblBid.Caption <> frmGxbjB.lblBid.Caption Then

    Call modBJD.BJDGXQing
    Call modBJD.BJDBound(Val(frmGxbjB.lblBid.Caption), "购销")
    
    tt = "select bid from xunjiaOld where old=" & Val(frmGXBj.lblOid.Caption) & " order by bid"
    frmGXBj.adoOid.Close
    frmGXBj.adoOid.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    frmGXBj.cmdLeft.Enabled = False
    frmGXBj.cmdRight.Enabled = False
    If frmGXBj.adoOid.RecordCount > 1 Then
        frmGXBj.cmdLeft.Enabled = True
    End If
    frmGXBj.adoOid.MoveLast
End If
frmWait.Visible = False
frmGXBj.cmdMod.Enabled = True
frmGXBj.cmdSave.Enabled = False
frmGxbjB.Visible = False
frmGXBj.Visible = True
frmGXBj.ZOrder 0
End Sub

Private Sub cmdYadd_Click()
Dim tt As String
Dim hg As Single
On Error Resume Next
If Val(txtYed.Text) = 0 Or Val(txtYingFu.Text) = 0 Then
Exit Sub
End If
hg = 0



On Error Resume Next
Set mod1.cmd = CreateObject("adodb.command")
mod1.cmd.ActiveConnection = mod1.cc
mod1.cmd.CommandText = "byjAdd"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@baoid") = Val(lblBaoId.Caption)
mod1.cmd.Parameters("@YED") = Val(txtYed.Text) / 100
mod1.cmd.Parameters("@yingFu") = Val(txtYingFu.Text)
mod1.cmd.Parameters("@lcren") = mod1.DName
mod1.cmd.Parameters("@lcuid") = mod1.DHid
On Error GoTo YJEA
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
tt = "update baojiaD set yj=" & hg & " where baoid=" & Val(lblBaoId.Caption)
On Error Resume Next
Set mod1.HTP = CreateObject("adodb.recordset")
On Error GoTo YJEA
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
txtYJ.Text = hg
Exit Sub
YJEA:
MsgBox "网络故障,请再试提交一次"

End Sub

Private Sub cmdYdel_Click()
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

'HG = 0
'adoYj.MoveFirst
'Do While Not adoYj.EOF
'   HG = HG + adoYj.Fields("支付金额").Value
'   adoYj.MoveNext
'Loop
'HG = HG + Val(txtYingFu.Text)
tt = "select sum(支付金额) from yongjin where baoid=" & Val(lblBaoId.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
hg = mod1.HTP.Fields(0).Value


tt = "update baojiaD set yj=" & Val(txtYingFu.Text) & " where baoid=" & Val(lblBaoId.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
txtYJ.Text = hg

End Sub

Private Sub comKhmc_Click()
Dim tt As String
On Error Resume Next
tt = "Select khdh from khzl where khqc ='" & comKhmc.Text & "'  order by kid desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
comKhmc.ToolTipText = mod1.HTP.Fields("khdh").Value
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

Private Sub Command1_Click()

End Sub

Private Sub dtgBao_Click()
Dim tt As String
Dim liD As Long
On Error Resume Next
dtgBao.Col = 11
txtSl.Text = dtgBao.Text
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
txtSl.Text = dtgBao.Text
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
frmGD.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = 76 Then
    If mod1.Kyj = True Then
        If frmYJ.Visible = False Then
            frmYJ.Visible = True
            lblYj.Visible = True
            txtYJ.Visible = True
            lblTcBe.Visible = True
            txtTcBe.Visible = True
            If lblLc.Caption = 2 And lblLcRen.Caption = mod1.DName Then
                txtYJ.Locked = False
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


Me.Left = 0
Me.Top = 0
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
frmFB.BorderStyle = 0

Set adoOid = CreateObject("adodb.recordset")
Set adoGx = CreateObject("adodb.recordset")
Set adoBx = CreateObject("adodb.recordset")
Set AdoKh = CreateObject("adodb.recordset")
Set adoYj = CreateObject("adodb.recordset")
Set adoFk = CreateObject("adodb.recordset")
Set adoGD = CreateObject("adodb.recordset")
Set adoHGD = CreateObject("adodb.recordset")

dtgYJ.ColWidth(0) = 300
dtgYJ.ColWidth(3) = 0
dtgMa.ColWidth(0) = 300
dtgMa.ColWidth(8) = 2000
dtgMa.ColWidth(15) = 0
dtgMa.ColWidth(16) = 0

dtgBao.ColWidth(0) = 300
dtgBao.ColWidth(8) = 2000
dtgBao.ColWidth(15) = 0
dtgBao.ColWidth(16) = 0
dtgBao.ColWidth(17) = 0
dtgBao.Left = 0
dtgBao.Top = 0
dtgFk.ColWidth(0) = 300
dtgFk.ColWidth(4) = 0
dtgFk.ColWidth(5) = 0

dtgGD.ColWidth(0) = 300
dtgGD.ColWidth(1) = 2500
dtgGD.ColWidth(2) = 800
dtgGD.ColWidth(3) = 800
dtgGD.ColWidth(6) = 0
dtgGD.ColWidth(7) = 0
dtgGD.ColWidth(8) = 0
frmFF.Top = 6930
txtFbnr.Locked = True
txtFbje.Locked = True
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
    Me.Visible = False
    If frmGxBiao.Visible = True Then
        frmGxBiao.Enabled = True
        frmGxBiao.ZOrder 0
    ElseIf Dialog.Visible = True Then
        Dialog.ZOrder 0
        Dialog.Enabled = True
    End If
    Cancel = True
    mod1.BTZ = 36
End If
End Sub


Private Sub optGDA_Click()
txtQdj.Text = 200
Lb = optGDA.Caption
End Sub

Private Sub optGDB_Click()
txtQdj.Text = 200
Lb = optGDB.Caption
End Sub


Private Sub optGDC_Click()
txtQdj.Text = ""
Lb = ""
End Sub


Private Sub txtHg_DblClick()
frmFF.Visible = True
End Sub

Private Sub txtHg_LostFocus()
txtYhg.Text = txtHg.Text
End Sub

Private Sub txtTcBe_Change()
cmdSave.Enabled = True
End Sub

Private Sub txtXm2_DblClick()
frmGD.Visible = True
End Sub

Private Sub txtXm2_LostFocus()
    txtCb.Text = Val(txtClcb.Text) + Val(txtYf.Text) + Val(txtYJ.Text) + Val(txtXm2.Text)
End Sub


Private Sub txtYf_LostFocus()
  txtCb.Text = Val(txtClcb.Text) + Val(txtYf.Text) + Val(txtYJ.Text) + Val(txtXm2.Text)
End Sub


Private Sub txtYj_Change()
cmdSave.Enabled = True

End Sub

Private Sub txtYJ_DblClick()
frmYm.Visible = True
End Sub

Private Sub txtYj_LostFocus()
  txtCb.Text = Val(txtClcb.Text) + Val(txtYf.Text) + Val(txtYJ.Text) + Val(txtXm2.Text)
End Sub


