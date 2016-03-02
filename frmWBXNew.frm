VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmWBXNew 
   Caption         =   "新版人工价格体系询价单"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15210
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15210
   Begin VB.Frame frmHide 
      Caption         =   "frmHid"
      Height          =   1455
      Left            =   5370
      TabIndex        =   8
      Top             =   60
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Label lblHLC 
         Caption         =   "lblHLC"
         Height          =   345
         Left            =   2250
         TabIndex        =   47
         Top             =   1080
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label lblHtbh 
         Caption         =   "对应合同"
         Height          =   255
         Left            =   3990
         TabIndex        =   46
         Top             =   330
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblBid 
         Caption         =   "lblBid"
         Height          =   255
         Left            =   90
         TabIndex        =   45
         Top             =   30
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblWid 
         Caption         =   "lblWid"
         Height          =   255
         Left            =   1440
         TabIndex        =   42
         Top             =   1140
         Width           =   1005
      End
      Begin VB.Label lblLc 
         Caption         =   "lblLc"
         Height          =   315
         Left            =   1050
         TabIndex        =   15
         Top             =   630
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label lblNlb 
         Caption         =   "lblNlb"
         Height          =   225
         Left            =   1920
         TabIndex        =   14
         Top             =   810
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label lblLcRen 
         Caption         =   "lblLcRen"
         Height          =   285
         Left            =   150
         TabIndex        =   13
         Top             =   420
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblLcUid 
         Caption         =   "lblLcUid"
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   930
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblFwid 
         Caption         =   "lblFwid"
         Height          =   255
         Left            =   1860
         TabIndex        =   11
         Top             =   450
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblUid 
         Caption         =   "lblUid"
         Height          =   255
         Left            =   3750
         TabIndex        =   10
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblYwy 
         Caption         =   "lblYwy"
         Height          =   285
         Left            =   2490
         TabIndex        =   9
         Top             =   780
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdQH 
      Caption         =   "切换旧版询价单"
      Height          =   1395
      Left            =   14850
      TabIndex        =   219
      Top             =   7350
      Width           =   375
   End
   Begin VB.Frame frmQm 
      BackColor       =   &H00C0FFC0&
      Caption         =   "评审建议"
      ForeColor       =   &H000000FF&
      Height          =   1785
      Left            =   6360
      TabIndex        =   0
      Top             =   6840
      Visible         =   0   'False
      Width           =   6315
      Begin VB.CommandButton cmdDing 
         BackColor       =   &H00FF8080&
         Caption         =   "决定"
         Height          =   285
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1320
         Width           =   735
      End
      Begin VB.OptionButton optT2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "拒绝"
         Height          =   195
         Left            =   5220
         TabIndex        =   3
         Top             =   870
         Width           =   675
      End
      Begin VB.OptionButton OptT1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "同意"
         Height          =   225
         Left            =   5220
         TabIndex        =   2
         Top             =   480
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.TextBox txtQM 
         BackColor       =   &H00C0FFFF&
         Height          =   1365
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   300
         Width           =   4965
      End
   End
   Begin VB.CommandButton cmdD 
      Enabled         =   0   'False
      Height          =   405
      Left            =   14250
      Picture         =   "frmWBXNew.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   218
      Top             =   8790
      Width           =   465
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   405
      Left            =   630
      TabIndex        =   169
      Top             =   6750
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   714
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin TabDlg.SSTab tabJG 
      Height          =   6105
      Left            =   -30
      TabIndex        =   48
      Top             =   690
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   10769
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "人工"
      TabPicture(0)   =   "frmWBXNew.frx":018A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmJz"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "耗材"
      TabPicture(1)   =   "frmWBXNew.frx":01A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5745
         Left            =   -75000
         TabIndex        =   170
         Top             =   330
         Width           =   15195
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgMN 
            Height          =   495
            Left            =   1710
            TabIndex        =   198
            Top             =   5310
            Visible         =   0   'False
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   873
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.TextBox txtLadr 
            Height          =   285
            Left            =   5310
            TabIndex        =   190
            Top             =   5310
            Width           =   9555
         End
         Begin VB.Frame frmLED 
            Caption         =   "编辑"
            Height          =   1305
            Left            =   0
            TabIndex        =   172
            Top             =   3720
            Width           =   15225
            Begin VB.TextBox txtJLBZ 
               Height          =   270
               Left            =   11190
               Locked          =   -1  'True
               TabIndex        =   214
               Top             =   210
               Width           =   1305
            End
            Begin VB.ComboBox comJLB 
               Height          =   300
               ItemData        =   "frmWBXNew.frx":01C2
               Left            =   8010
               List            =   "frmWBXNew.frx":01CC
               TabIndex        =   212
               Top             =   180
               Width           =   1935
            End
            Begin VB.TextBox txtL1 
               Height          =   270
               Left            =   1590
               TabIndex        =   200
               Top             =   960
               Width           =   2085
            End
            Begin VB.TextBox txtL2 
               Height          =   285
               Left            =   5280
               Locked          =   -1  'True
               TabIndex        =   199
               Top             =   930
               Width           =   1875
            End
            Begin VB.CommandButton cmdGB 
               Caption         =   "关闭"
               Height          =   285
               Left            =   14580
               TabIndex        =   191
               Top             =   540
               Width           =   615
            End
            Begin VB.CommandButton cmdLqing 
               Caption         =   "清空"
               Height          =   285
               Left            =   13800
               TabIndex        =   188
               Top             =   540
               Width           =   735
            End
            Begin VB.CommandButton cmdLDel 
               BackColor       =   &H008080FF&
               Caption         =   "删除"
               Height          =   285
               Left            =   13800
               TabIndex        =   187
               Top             =   180
               Width           =   735
            End
            Begin VB.CommandButton cmdLGx 
               Caption         =   "更新"
               Height          =   285
               Left            =   12960
               TabIndex        =   186
               Top             =   540
               Width           =   735
            End
            Begin VB.CommandButton cmdLadd 
               Caption         =   "添加"
               Height          =   285
               Left            =   12960
               TabIndex        =   185
               Top             =   180
               Width           =   735
            End
            Begin VB.TextBox txtLBz 
               Height          =   675
               Left            =   7980
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   184
               Top             =   570
               Visible         =   0   'False
               Width           =   4545
            End
            Begin VB.TextBox txtLDW 
               Height          =   285
               Left            =   6690
               TabIndex        =   182
               Top             =   570
               Width           =   465
            End
            Begin VB.TextBox txtLsl 
               Height          =   285
               Left            =   5280
               TabIndex        =   180
               Top             =   570
               Width           =   855
            End
            Begin VB.TextBox txtLjmc 
               Height          =   300
               Left            =   1590
               TabIndex        =   177
               Top             =   570
               Width           =   2085
            End
            Begin VB.TextBox txtLjbh 
               Height          =   270
               Left            =   5280
               TabIndex        =   175
               Top             =   180
               Width           =   1905
            End
            Begin VB.TextBox txtLpb 
               Height          =   300
               Left            =   1590
               TabIndex        =   174
               Top             =   180
               Width           =   2085
            End
            Begin VB.Label Label31 
               Caption         =   "基准比例"
               Height          =   255
               Left            =   10320
               TabIndex        =   213
               Top             =   240
               Width           =   825
            End
            Begin VB.Label Label30 
               Caption         =   "分类"
               Height          =   195
               Left            =   7500
               TabIndex        =   211
               Top             =   240
               Width           =   375
            End
            Begin VB.Label lblL1 
               Caption         =   "单价"
               Height          =   255
               Left            =   870
               TabIndex        =   202
               Top             =   990
               Width           =   675
            End
            Begin VB.Label lblL2 
               Caption         =   "基准单价"
               Height          =   225
               Left            =   4380
               TabIndex        =   201
               Top             =   960
               Width           =   765
            End
            Begin VB.Label Label16 
               Caption         =   "备注"
               Height          =   285
               Left            =   7500
               TabIndex        =   183
               Top             =   600
               Visible         =   0   'False
               Width           =   525
            End
            Begin VB.Label Label15 
               Caption         =   "单位"
               Height          =   255
               Left            =   6270
               TabIndex        =   181
               Top             =   600
               Width           =   405
            End
            Begin VB.Label Label14 
               Caption         =   "数量"
               Height          =   225
               Left            =   4710
               TabIndex        =   179
               Top             =   600
               Width           =   465
            End
            Begin VB.Label Label13 
               Caption         =   "耗材名称"
               Height          =   315
               Left            =   540
               TabIndex        =   178
               Top             =   615
               Width           =   945
            End
            Begin VB.Label Label12 
               Caption         =   "规格型号"
               Height          =   315
               Left            =   4350
               TabIndex        =   176
               Top             =   210
               Width           =   975
            End
            Begin VB.Label Label11 
               Caption         =   "品牌"
               Height          =   315
               Left            =   870
               TabIndex        =   173
               Top             =   240
               Width           =   1005
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgMa 
            Height          =   4725
            Left            =   30
            TabIndex        =   171
            Top             =   0
            Width           =   15225
            _ExtentX        =   26855
            _ExtentY        =   8334
            _Version        =   393216
            Rows            =   20
            Cols            =   15
            BackColorBkg    =   -2147483627
            SelectionMode   =   1
            AllowUserResizing=   1
            PictureType     =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   15
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label lblLid 
            Caption         =   "lblLid"
            Height          =   225
            Left            =   150
            TabIndex        =   193
            Top             =   5160
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label Label19 
            Caption         =   "运货地址"
            Height          =   285
            Left            =   4380
            TabIndex        =   189
            Top             =   5340
            Width           =   795
         End
      End
      Begin VB.Frame frmJz 
         BorderStyle     =   0  'None
         Caption         =   "价格体系"
         Height          =   5775
         Left            =   0
         TabIndex        =   49
         Top             =   300
         Width           =   15195
         Begin VB.TextBox txtDXNR 
            Height          =   5775
            Left            =   60
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   220
            Text            =   "frmWBXNew.frx":01D6
            Top             =   2820
            Width           =   15195
         End
         Begin VB.Frame frmM1 
            Caption         =   "主机"
            Height          =   2775
            Left            =   2130
            TabIndex        =   50
            Top             =   240
            Width           =   10395
            Begin VB.ComboBox comA13 
               Height          =   300
               ItemData        =   "frmWBXNew.frx":01DC
               Left            =   7770
               List            =   "frmWBXNew.frx":01E9
               TabIndex        =   79
               Text            =   "Combo5"
               Top             =   780
               Width           =   2325
            End
            Begin VB.Frame frmXH 
               BorderStyle     =   0  'None
               Caption         =   "Frame11"
               Height          =   1005
               Left            =   6000
               TabIndex        =   73
               Top             =   1110
               Width           =   4095
               Begin VB.TextBox txtA20 
                  Height          =   285
                  Left            =   1770
                  TabIndex        =   75
                  Top             =   510
                  Width           =   1965
               End
               Begin VB.ComboBox comA15 
                  Height          =   300
                  ItemData        =   "frmWBXNew.frx":01F7
                  Left            =   1770
                  List            =   "frmWBXNew.frx":0201
                  Style           =   2  'Dropdown List
                  TabIndex        =   74
                  Top             =   120
                  Width           =   2355
               End
               Begin VB.Label Label61 
                  Caption         =   "机组使用时间："
                  Height          =   255
                  Left            =   150
                  TabIndex        =   78
                  Top             =   570
                  Width           =   1305
               End
               Begin VB.Label lblXh 
                  Caption         =   "供热方式："
                  Height          =   225
                  Left            =   510
                  TabIndex        =   77
                  Top             =   150
                  Width           =   1065
               End
               Begin VB.Label Label1 
                  Caption         =   "年"
                  Height          =   345
                  Left            =   3870
                  TabIndex        =   76
                  Top             =   540
                  Width           =   315
               End
            End
            Begin VB.Frame Frame10 
               BorderStyle     =   0  'None
               Caption         =   "Frame10"
               Height          =   555
               Left            =   450
               TabIndex        =   69
               Top             =   2040
               Width           =   4965
               Begin VB.CheckBox chkA11 
                  Caption         =   "化学清洗"
                  Height          =   225
                  Left            =   2850
                  TabIndex        =   71
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.CheckBox chkA10 
                  Caption         =   "物理清洗"
                  Height          =   255
                  Left            =   1620
                  TabIndex        =   70
                  Top             =   210
                  Width           =   1185
               End
               Begin VB.Label Label39 
                  Caption         =   "清洗方法："
                  Height          =   225
                  Left            =   270
                  TabIndex        =   72
                  Top             =   240
                  Width           =   975
               End
            End
            Begin VB.ComboBox comA2 
               Height          =   300
               ItemData        =   "frmWBXNew.frx":0215
               Left            =   4170
               List            =   "frmWBXNew.frx":021F
               Style           =   2  'Dropdown List
               TabIndex        =   68
               Top             =   900
               Width           =   825
            End
            Begin VB.Frame Frame9 
               BorderStyle     =   0  'None
               Caption         =   "Frame9"
               Height          =   435
               Left            =   5310
               TabIndex        =   63
               Top             =   2190
               Width           =   5085
               Begin VB.OptionButton optA21c 
                  Caption         =   "大修"
                  ForeColor       =   &H00C00000&
                  Height          =   225
                  Left            =   3780
                  TabIndex        =   66
                  Top             =   150
                  Visible         =   0   'False
                  Width           =   1155
               End
               Begin VB.OptionButton optA21b 
                  Caption         =   "一次性保养"
                  Height          =   225
                  Left            =   2310
                  TabIndex        =   65
                  Top             =   150
                  Width           =   1245
               End
               Begin VB.OptionButton optA21a 
                  Caption         =   "维保"
                  Height          =   255
                  Left            =   1290
                  TabIndex        =   64
                  Top             =   120
                  Width           =   855
               End
               Begin VB.Label Label36 
                  Caption         =   "保养性质："
                  Height          =   225
                  Left            =   150
                  TabIndex        =   67
                  Top             =   150
                  Width           =   915
               End
            End
            Begin VB.TextBox txtA1 
               Height          =   270
               Left            =   2010
               TabIndex        =   62
               Top             =   930
               Width           =   1425
            End
            Begin VB.Frame Frame3 
               BorderStyle     =   0  'None
               Caption         =   "Frame3"
               Height          =   285
               Left            =   1620
               TabIndex        =   57
               Top             =   1380
               Width           =   3855
               Begin VB.TextBox txtA5 
                  Height          =   270
                  Left            =   2940
                  TabIndex        =   61
                  Text            =   "1"
                  Top             =   0
                  Width           =   405
               End
               Begin VB.TextBox txtA3 
                  Height          =   270
                  Left            =   1380
                  TabIndex        =   60
                  Text            =   "1"
                  Top             =   0
                  Width           =   435
               End
               Begin VB.CheckBox chkA7 
                  Caption         =   "冷凝器*"
                  Height          =   255
                  Left            =   2010
                  TabIndex        =   59
                  Top             =   0
                  Width           =   975
               End
               Begin VB.CheckBox chkA6 
                  Caption         =   "蒸发器*"
                  Height          =   225
                  Left            =   390
                  TabIndex        =   58
                  Top             =   30
                  Width           =   945
               End
            End
            Begin VB.Frame frmCai 
               BorderStyle     =   0  'None
               Caption         =   "Frame4"
               Height          =   315
               Left            =   600
               TabIndex        =   53
               Top             =   1800
               Width           =   5025
               Begin VB.CheckBox chkA7a 
                  Caption         =   "清洗翅片"
                  Height          =   195
                  Left            =   3750
                  TabIndex        =   217
                  Top             =   60
                  Width           =   1065
               End
               Begin VB.OptionButton optA8 
                  Caption         =   "拆一端"
                  Height          =   225
                  Left            =   1440
                  TabIndex        =   55
                  Top             =   60
                  Width           =   1095
               End
               Begin VB.OptionButton optA9 
                  Caption         =   "拆二端"
                  Height          =   240
                  Left            =   2670
                  TabIndex        =   54
                  Top             =   60
                  Width           =   1035
               End
               Begin VB.Label Label18 
                  Caption         =   "拆端盖："
                  Height          =   255
                  Left            =   300
                  TabIndex        =   56
                  Top             =   60
                  Width           =   975
               End
            End
            Begin VB.ComboBox comA0 
               Height          =   300
               ItemData        =   "frmWBXNew.frx":022D
               Left            =   2010
               List            =   "frmWBXNew.frx":0240
               Style           =   2  'Dropdown List
               TabIndex        =   52
               Top             =   360
               Width           =   1695
            End
            Begin VB.TextBox txtA12 
               Height          =   270
               Left            =   7770
               TabIndex        =   51
               Top             =   360
               Width           =   2295
            End
            Begin VB.Label Label38 
               Caption         =   "单位"
               Height          =   225
               Left            =   3630
               TabIndex        =   85
               Top             =   960
               Width           =   375
            End
            Begin VB.Label lblJa 
               Caption         =   "机组冷量："
               Height          =   285
               Index           =   0
               Left            =   720
               TabIndex        =   84
               Top             =   960
               Width           =   945
            End
            Begin VB.Label lblJa 
               Caption         =   "机组年巡视次数："
               Height          =   285
               Index           =   4
               Left            =   5970
               TabIndex        =   83
               Top             =   810
               Width           =   1515
            End
            Begin VB.Label Label6 
               Caption         =   "主机类型："
               Height          =   195
               Left            =   720
               TabIndex        =   82
               Top             =   420
               Width           =   1065
            End
            Begin VB.Label Label26 
               Caption         =   "清洗："
               Height          =   225
               Left            =   1080
               TabIndex        =   81
               Top             =   1440
               Width           =   615
            End
            Begin VB.Label Label27 
               Caption         =   "单机组压缩机数量："
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   5790
               TabIndex        =   80
               Top             =   390
               Width           =   1665
            End
         End
         Begin VB.Frame frmM5 
            Caption         =   "小机"
            Height          =   2805
            Left            =   7980
            TabIndex        =   86
            Top             =   510
            Width           =   10965
            Begin VB.Frame frmD 
               Caption         =   "空调箱"
               Height          =   2085
               Left            =   6780
               TabIndex        =   110
               Top             =   450
               Width           =   3315
               Begin VB.TextBox txtC39 
                  Height          =   285
                  Left            =   1530
                  TabIndex        =   116
                  Top             =   1680
                  Width           =   1215
               End
               Begin VB.TextBox txtC38 
                  Height          =   285
                  Left            =   1530
                  TabIndex        =   115
                  Top             =   1380
                  Width           =   1215
               End
               Begin VB.TextBox txtC52 
                  Height          =   285
                  Left            =   1530
                  TabIndex        =   114
                  Top             =   1020
                  Width           =   1215
               End
               Begin VB.CheckBox chkC51a 
                  Caption         =   "保养"
                  Height          =   225
                  Left            =   240
                  TabIndex        =   113
                  Top             =   690
                  Width           =   765
               End
               Begin VB.CheckBox chkC51b 
                  Caption         =   "巡视"
                  Height          =   195
                  Left            =   1155
                  TabIndex        =   112
                  Top             =   690
                  Width           =   705
               End
               Begin VB.CheckBox chkC51c 
                  Caption         =   "应急"
                  Height          =   225
                  Left            =   2010
                  TabIndex        =   111
                  Top             =   690
                  Width           =   675
               End
               Begin VB.Label Label21 
                  Caption         =   "3"
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   5.25
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   135
                  Left            =   780
                  TabIndex        =   192
                  Top             =   1050
                  Width           =   135
               End
               Begin VB.Label Label55 
                  Caption         =   "巡视次数："
                  Height          =   255
                  Left            =   510
                  TabIndex        =   121
                  Top             =   1740
                  Width           =   945
               End
               Begin VB.Label Label54 
                  Caption         =   "保养次数："
                  Height          =   195
                  Left            =   510
                  TabIndex        =   120
                  Top             =   1410
                  Width           =   945
               End
               Begin VB.Label Label2 
                  Caption         =   "保养性质："
                  Height          =   225
                  Left            =   210
                  TabIndex        =   119
                  Top             =   390
                  Width           =   915
               End
               Begin VB.Label Label7 
                  Caption         =   "风量（m /h）："
                  Height          =   255
                  Left            =   150
                  TabIndex        =   118
                  Top             =   1050
                  Width           =   1515
               End
               Begin VB.Label Label9 
                  Caption         =   "3"
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   5.25
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   105
                  Left            =   780
                  TabIndex        =   117
                  Top             =   1050
                  Width           =   105
               End
            End
            Begin VB.Frame frmB 
               Caption         =   "小机安装"
               Height          =   2115
               Left            =   2190
               TabIndex        =   105
               Top             =   450
               Width           =   2355
               Begin VB.TextBox txtC33 
                  Height          =   270
                  Left            =   1500
                  TabIndex        =   107
                  Top             =   870
                  Width           =   615
               End
               Begin VB.TextBox txtC32 
                  Height          =   270
                  Left            =   1500
                  TabIndex        =   106
                  Top             =   390
                  Visible         =   0   'False
                  Width           =   645
               End
               Begin VB.Label Label53 
                  Caption         =   "外机数量(>3HP)"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   109
                  Top             =   930
                  Width           =   1395
               End
               Begin VB.Label Label52 
                  Caption         =   "外机数量(<3HP)"
                  Height          =   285
                  Left            =   150
                  TabIndex        =   108
                  Top             =   420
                  Visible         =   0   'False
                  Width           =   1455
               End
            End
            Begin VB.Frame frmC 
               Caption         =   "风机盘管"
               Height          =   2085
               Left            =   4530
               TabIndex        =   97
               Top             =   450
               Width           =   2265
               Begin VB.TextBox txtC36 
                  Height          =   270
                  Left            =   1080
                  TabIndex        =   101
                  Top             =   1560
                  Width           =   675
               End
               Begin VB.TextBox txtC35 
                  Height          =   270
                  Left            =   1080
                  TabIndex        =   100
                  Top             =   1200
                  Width           =   675
               End
               Begin VB.CheckBox chkC37a 
                  Caption         =   "保养"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   99
                  Top             =   720
                  Width           =   825
               End
               Begin VB.CheckBox chkC37b 
                  Caption         =   "巡视"
                  Height          =   255
                  Left            =   1110
                  TabIndex        =   98
                  Top             =   690
                  Width           =   735
               End
               Begin VB.Label Label51 
                  Caption         =   "巡视次数："
                  Height          =   225
                  Left            =   150
                  TabIndex        =   104
                  Top             =   1560
                  Width           =   945
               End
               Begin VB.Label Label50 
                  Caption         =   "保养次数："
                  Height          =   195
                  Left            =   150
                  TabIndex        =   103
                  Top             =   1230
                  Width           =   945
               End
               Begin VB.Label Label47 
                  Caption         =   "保养性质："
                  Height          =   225
                  Left            =   150
                  TabIndex        =   102
                  Top             =   360
                  Width           =   915
               End
            End
            Begin VB.Frame frmA 
               Caption         =   "小机"
               Height          =   2115
               Left            =   120
               TabIndex        =   87
               Top             =   450
               Width           =   2085
               Begin VB.TextBox txtC30 
                  Height          =   270
                  Left            =   1140
                  TabIndex        =   93
                  Top             =   1650
                  Width           =   675
               End
               Begin VB.TextBox txtC29 
                  Height          =   270
                  Left            =   1140
                  TabIndex        =   92
                  Top             =   1290
                  Width           =   675
               End
               Begin VB.CheckBox chkC31a 
                  Caption         =   "保养"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   91
                  Top             =   570
                  Width           =   735
               End
               Begin VB.CheckBox chkC31b 
                  Caption         =   "巡视"
                  Height          =   225
                  Left            =   1170
                  TabIndex        =   90
                  Top             =   540
                  Width           =   735
               End
               Begin VB.CheckBox chkC31c 
                  Caption         =   "应急"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   89
                  Top             =   840
                  Width           =   675
               End
               Begin VB.CheckBox chkC31d 
                  Caption         =   "移机"
                  Height          =   225
                  Left            =   1170
                  TabIndex        =   88
                  Top             =   840
                  Width           =   735
               End
               Begin VB.Label Label49 
                  Caption         =   "巡视次数："
                  Height          =   225
                  Left            =   210
                  TabIndex        =   96
                  Top             =   1650
                  Width           =   945
               End
               Begin VB.Label Label48 
                  Caption         =   "保养次数："
                  Height          =   195
                  Left            =   210
                  TabIndex        =   95
                  Top             =   1320
                  Width           =   945
               End
               Begin VB.Label Label46 
                  Caption         =   "保养性质："
                  Height          =   225
                  Left            =   180
                  TabIndex        =   94
                  Top             =   270
                  Width           =   915
               End
            End
         End
         Begin VB.Frame frmM2 
            Caption         =   "水泵"
            Height          =   2745
            Left            =   8940
            TabIndex        =   122
            Top             =   1920
            Width           =   10905
            Begin VB.Frame frmJN 
               Height          =   2595
               Left            =   5460
               TabIndex        =   203
               Top             =   90
               Width           =   2625
               Begin VB.Label Label29 
                  Caption         =   "10           600"
                  ForeColor       =   &H8000000D&
                  Height          =   225
                  Left            =   270
                  TabIndex        =   210
                  Top             =   2190
                  Width           =   1845
               End
               Begin VB.Label Label28 
                  Caption         =   "8           750"
                  ForeColor       =   &H8000000D&
                  Height          =   225
                  Left            =   360
                  TabIndex        =   209
                  Top             =   1800
                  Width           =   1845
               End
               Begin VB.Label Label25 
                  Caption         =   "6          1000"
                  ForeColor       =   &H8000000D&
                  Height          =   225
                  Left            =   360
                  TabIndex        =   208
                  Top             =   1410
                  Width           =   1845
               End
               Begin VB.Label Label24 
                  Caption         =   "4          1500"
                  ForeColor       =   &H8000000D&
                  Height          =   225
                  Left            =   360
                  TabIndex        =   207
                  Top             =   1020
                  Width           =   1845
               End
               Begin VB.Label Label22 
                  Caption         =   "2          3000"
                  ForeColor       =   &H8000000D&
                  Height          =   225
                  Left            =   360
                  TabIndex        =   206
                  Top             =   630
                  Width           =   1845
               End
               Begin VB.Label Label20 
                  Caption         =   "功率"
                  ForeColor       =   &H8000000D&
                  Height          =   225
                  Left            =   1320
                  TabIndex        =   205
                  Top             =   240
                  Width           =   675
               End
               Begin VB.Label Label3 
                  Caption         =   "级数"
                  ForeColor       =   &H8000000D&
                  Height          =   195
                  Left            =   270
                  TabIndex        =   204
                  Top             =   240
                  Width           =   555
               End
            End
            Begin VB.ComboBox comB27 
               Height          =   300
               ItemData        =   "frmWBXNew.frx":0280
               Left            =   3900
               List            =   "frmWBXNew.frx":0293
               TabIndex        =   133
               Text            =   "1"
               Top             =   1560
               Width           =   1425
            End
            Begin VB.ComboBox comB26 
               Height          =   300
               ItemData        =   "frmWBXNew.frx":02A7
               Left            =   3900
               List            =   "frmWBXNew.frx":02B1
               Style           =   2  'Dropdown List
               TabIndex        =   132
               Top             =   1050
               Width           =   1425
            End
            Begin VB.TextBox txtB23 
               Height          =   285
               Left            =   3900
               TabIndex        =   131
               Top             =   480
               Width           =   1335
            End
            Begin VB.ComboBox comB25 
               Height          =   300
               ItemData        =   "frmWBXNew.frx":02C1
               Left            =   1440
               List            =   "frmWBXNew.frx":02CB
               Style           =   2  'Dropdown List
               TabIndex        =   130
               Top             =   1050
               Width           =   975
            End
            Begin VB.TextBox txtB22 
               Height          =   270
               Left            =   1380
               TabIndex        =   129
               Top             =   510
               Width           =   975
            End
            Begin VB.Frame Frame2 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   615
               Left            =   360
               TabIndex        =   123
               Top             =   1890
               Width           =   5955
               Begin VB.CheckBox chkB28a 
                  Caption         =   "保养"
                  Height          =   285
                  Left            =   1290
                  TabIndex        =   127
                  Top             =   210
                  Width           =   705
               End
               Begin VB.CheckBox chkB28b 
                  Caption         =   "巡视"
                  Height          =   285
                  Left            =   2160
                  TabIndex        =   126
                  Top             =   210
                  Width           =   705
               End
               Begin VB.CheckBox chkB28c 
                  Caption         =   "大修"
                  Height          =   285
                  Left            =   3030
                  TabIndex        =   125
                  Top             =   210
                  Width           =   705
               End
               Begin VB.CheckBox chkB28d 
                  Caption         =   "急修"
                  Height          =   285
                  Left            =   3900
                  TabIndex        =   124
                  Top             =   210
                  Width           =   705
               End
               Begin VB.Label Label42 
                  Caption         =   "保养性质："
                  Height          =   225
                  Left            =   150
                  TabIndex        =   128
                  Top             =   270
                  Width           =   915
               End
            End
            Begin VB.Label Label45 
               Caption         =   "水泵级数："
               Height          =   225
               Left            =   2910
               TabIndex        =   138
               Top             =   1620
               Width           =   915
            End
            Begin VB.Label Label44 
               Caption         =   "水泵类型："
               Height          =   195
               Left            =   2910
               TabIndex        =   137
               Top             =   1110
               Width           =   945
            End
            Begin VB.Label Label43 
               Caption         =   "巡视次数："
               Height          =   255
               Left            =   2910
               TabIndex        =   136
               Top             =   540
               Width           =   1035
            End
            Begin VB.Label Label41 
               Caption         =   "品牌："
               Height          =   195
               Left            =   780
               TabIndex        =   135
               Top             =   1110
               Width           =   555
            End
            Begin VB.Label Label40 
               Caption         =   "功率（KW）："
               Height          =   225
               Left            =   240
               TabIndex        =   134
               Top             =   570
               Width           =   1095
            End
         End
         Begin VB.Frame frmM3 
            Caption         =   "电机"
            Height          =   2745
            Left            =   6660
            TabIndex        =   139
            Top             =   1020
            Width           =   10635
            Begin VB.CheckBox Check18 
               Caption         =   "大修"
               Height          =   225
               Left            =   1320
               TabIndex        =   143
               Top             =   2070
               Width           =   795
            End
            Begin VB.CheckBox Check7 
               Caption         =   "保养"
               Height          =   195
               Left            =   1320
               TabIndex        =   142
               Top             =   2400
               Width           =   795
            End
            Begin VB.ComboBox Combo8 
               Height          =   300
               ItemData        =   "frmWBXNew.frx":02DB
               Left            =   1560
               List            =   "frmWBXNew.frx":02E5
               Style           =   2  'Dropdown List
               TabIndex        =   141
               Top             =   870
               Width           =   1515
            End
            Begin VB.TextBox Text19 
               Height          =   270
               Left            =   1560
               TabIndex        =   140
               Text            =   "Text19"
               Top             =   450
               Width           =   1485
            End
            Begin VB.Label Label59 
               Caption         =   "品牌："
               Height          =   225
               Left            =   810
               TabIndex        =   146
               Top             =   960
               Width           =   555
            End
            Begin VB.Label Label58 
               Caption         =   "功率(KW)："
               Height          =   255
               Left            =   450
               TabIndex        =   145
               Top             =   480
               Width           =   1005
            End
            Begin VB.Label Label56 
               Caption         =   "保养性质："
               Height          =   225
               Left            =   300
               TabIndex        =   144
               Top             =   2130
               Width           =   915
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgJG 
            Height          =   2775
            Left            =   60
            TabIndex        =   162
            Top             =   30
            Width           =   15165
            _ExtentX        =   26749
            _ExtentY        =   4895
            _Version        =   393216
            Rows            =   15
            SelectionMode   =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.ComboBox comDX 
            Height          =   300
            ItemData        =   "frmWBXNew.frx":02F5
            Left            =   1170
            List            =   "frmWBXNew.frx":0311
            Style           =   2  'Dropdown List
            TabIndex        =   161
            Top             =   3060
            Width           =   2565
         End
         Begin VB.TextBox txtXh 
            Height          =   315
            Left            =   1170
            TabIndex        =   160
            Top             =   3840
            Width           =   2505
         End
         Begin VB.TextBox txtXLBH 
            Height          =   270
            Left            =   1170
            TabIndex        =   159
            Top             =   4260
            Width           =   2565
         End
         Begin VB.TextBox txtSL 
            Height          =   285
            Left            =   1170
            TabIndex        =   158
            Top             =   4650
            Width           =   2565
         End
         Begin VB.TextBox txtPb 
            Height          =   270
            Left            =   1170
            TabIndex        =   157
            Top             =   3450
            Width           =   2535
         End
         Begin VB.Frame frmNewF 
            BorderStyle     =   0  'None
            Caption         =   "Frame8"
            Height          =   285
            Left            =   300
            TabIndex        =   154
            Top             =   5370
            Width           =   3165
            Begin VB.OptionButton opt15 
               Caption         =   "新签"
               Height          =   195
               Left            =   30
               TabIndex        =   156
               Top             =   90
               Width           =   855
            End
            Begin VB.OptionButton opt16 
               Caption         =   "续签"
               Height          =   255
               Left            =   900
               TabIndex        =   155
               Top             =   60
               Width           =   795
            End
         End
         Begin VB.TextBox txtWc 
            Height          =   270
            Left            =   1170
            TabIndex        =   153
            Top             =   5040
            Width           =   495
         End
         Begin VB.ComboBox comWD 
            Height          =   300
            ItemData        =   "frmWBXNew.frx":0351
            Left            =   1740
            List            =   "frmWBXNew.frx":035E
            TabIndex        =   152
            Text            =   "年"
            Top             =   5040
            Width           =   855
         End
         Begin VB.CommandButton cmdBJ 
            BackColor       =   &H00C0FFC0&
            Caption         =   "报价"
            Height          =   285
            Left            =   3750
            Style           =   1  'Graphical
            TabIndex        =   151
            Top             =   5400
            Width           =   555
         End
         Begin VB.Frame frmED 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   1275
            Left            =   3720
            TabIndex        =   147
            Top             =   4170
            Width           =   645
            Begin VB.CommandButton cmdAdd 
               Caption         =   "添加"
               Height          =   285
               Left            =   30
               TabIndex        =   150
               Top             =   150
               Width           =   585
            End
            Begin VB.CommandButton cmdDel 
               Caption         =   "删除"
               Height          =   285
               Left            =   30
               TabIndex        =   149
               Top             =   510
               Width           =   555
            End
            Begin VB.CommandButton cmdGx 
               Caption         =   "更新"
               Height          =   285
               Left            =   30
               TabIndex        =   148
               Top             =   900
               Width           =   585
            End
         End
         Begin VB.Label Label32 
            Caption         =   "保养对象："
            Height          =   285
            Left            =   180
            TabIndex        =   168
            Top             =   3060
            Width           =   1005
         End
         Begin VB.Label Label33 
            Caption         =   "品牌："
            Height          =   225
            Left            =   540
            TabIndex        =   167
            Top             =   3480
            Width           =   555
         End
         Begin VB.Label Label34 
            Caption         =   "型号："
            Height          =   225
            Left            =   540
            TabIndex        =   166
            Top             =   3870
            Width           =   675
         End
         Begin VB.Label Label35 
            Caption         =   "系列编号："
            Height          =   225
            Left            =   180
            TabIndex        =   165
            Top             =   4290
            Width           =   945
         End
         Begin VB.Label lblSl 
            Alignment       =   1  'Right Justify
            Caption         =   "数量："
            Height          =   255
            Left            =   150
            TabIndex        =   164
            Top             =   4710
            Width           =   945
         End
         Begin VB.Label Label10 
            Caption         =   "维保年限："
            Height          =   225
            Left            =   180
            TabIndex        =   163
            Top             =   5070
            Width           =   915
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00C000C0&
            BorderWidth     =   3
            DrawMode        =   10  'Mask Pen
            FillColor       =   &H00C00000&
            FillStyle       =   4  'Upward Diagonal
            Height          =   2805
            Left            =   4590
            Top             =   3060
            Width           =   10455
         End
      End
   End
   Begin VB.ComboBox txtZu 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmWBXNew.frx":0370
      Left            =   11190
      List            =   "frmWBXNew.frx":037C
      TabIndex        =   44
      Text            =   "Combo2"
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox comXmmc 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1530
      TabIndex        =   43
      Text            =   "Text2"
      Top             =   270
      Width           =   3135
   End
   Begin VB.CommandButton cmdQm 
      Height          =   345
      Index           =   2
      Left            =   11310
      TabIndex        =   39
      Top             =   8430
      Width           =   945
   End
   Begin VB.CommandButton cmdQm 
      Height          =   345
      Index           =   1
      Left            =   10230
      TabIndex        =   36
      Top             =   8430
      Width           =   945
   End
   Begin VB.Frame frmN 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   885
      Left            =   0
      TabIndex        =   31
      Top             =   8340
      Width           =   8175
      Begin VB.TextBox txtLyf 
         Height          =   270
         Left            =   6660
         TabIndex        =   215
         Top             =   570
         Width           =   1395
      End
      Begin VB.TextBox txtLJhg 
         Height          =   315
         Left            =   4050
         TabIndex        =   196
         Top             =   450
         Width           =   1695
      End
      Begin VB.TextBox txtLhg 
         Height          =   315
         Left            =   4050
         TabIndex        =   194
         Top             =   60
         Width           =   1695
      End
      Begin VB.TextBox txt1 
         Height          =   315
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   60
         Width           =   1665
      End
      Begin VB.TextBox txt2 
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   420
         Width           =   1665
      End
      Begin VB.Label Label17 
         Caption         =   "运费"
         Height          =   255
         Left            =   5940
         TabIndex        =   216
         Top             =   630
         Width           =   675
      End
      Begin VB.Label lblLJHg 
         Caption         =   "耗材基准价格"
         Height          =   375
         Left            =   3060
         TabIndex        =   197
         Top             =   450
         Width           =   855
      End
      Begin VB.Label lblLhg 
         Caption         =   "耗材成本"
         Height          =   255
         Left            =   3060
         TabIndex        =   195
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lbl1 
         Caption         =   "人工成本"
         Height          =   255
         Left            =   270
         TabIndex        =   35
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lbl2 
         Caption         =   "人工基准价格"
         Height          =   375
         Left            =   270
         TabIndex        =   34
         Top             =   450
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdBack 
      Height          =   375
      Left            =   14760
      Picture         =   "frmWBXNew.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "返回"
      Top             =   8790
      Width           =   465
   End
   Begin VB.CommandButton cmdQm 
      Height          =   345
      Index           =   0
      Left            =   9150
      TabIndex        =   19
      Top             =   8430
      Width           =   945
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   13860
      Picture         =   "frmWBXNew.frx":048C
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "保存"
      Top             =   8820
      Width           =   465
   End
   Begin VB.CommandButton cmdMod 
      Height          =   375
      Left            =   13350
      Picture         =   "frmWBXNew.frx":0AF6
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "修改"
      Top             =   8820
      Width           =   465
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印"
      Height          =   405
      Left            =   12600
      TabIndex        =   16
      Top             =   8280
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CommandButton cmdPje 
      BackColor       =   &H00FF8080&
      Caption         =   "评审建议"
      Height          =   1095
      Left            =   8610
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8100
      Width           =   465
   End
   Begin VB.TextBox txtBz 
      Height          =   1095
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   7050
      Width           =   7215
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12960
      Top             =   7000
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   13590
      Top             =   7000
   End
   Begin VB.CommandButton cmdHT 
      BackColor       =   &H00C0FFC0&
      Caption         =   "合同评审单"
      Height          =   435
      Left            =   14130
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   210
      Width           =   1065
   End
   Begin VB.OLE OLE1 
      Class           =   "Excel.Sheet.8"
      Height          =   885
      Left            =   14340
      OleObjectBlob   =   "frmWBXNew.frx":0E00
      TabIndex        =   221
      Top             =   7020
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label lblTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   2
      Left            =   11310
      TabIndex        =   41
      Top             =   8820
      Width           =   945
   End
   Begin VB.Label lblQM 
      Caption         =   "业务员确认"
      Height          =   225
      Index           =   2
      Left            =   11340
      TabIndex        =   40
      Top             =   8130
      Width           =   1005
   End
   Begin VB.Label lblTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   1
      Left            =   10230
      TabIndex        =   38
      Top             =   8820
      Width           =   945
   End
   Begin VB.Label lblQM 
      Caption         =   "商务支持"
      Height          =   225
      Index           =   1
      Left            =   10260
      TabIndex        =   37
      Top             =   8130
      Width           =   1005
   End
   Begin VB.Label Label4 
      Caption         =   "项目名称"
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
      Left            =   330
      TabIndex        =   30
      Top             =   330
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "编号"
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
      Left            =   7260
      TabIndex        =   29
      Top             =   300
      Width           =   555
   End
   Begin VB.Label lblBh 
      BackColor       =   &H80000014&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8010
      TabIndex        =   28
      Top             =   255
      Width           =   1725
   End
   Begin VB.Label Label8 
      Caption         =   "组长"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10350
      TabIndex        =   27
      Top             =   330
      Width           =   555
   End
   Begin VB.Label lblQM 
      Caption         =   "业务员"
      Height          =   225
      Index           =   0
      Left            =   9180
      TabIndex        =   26
      Top             =   8130
      Width           =   1005
   End
   Begin VB.Label lblTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   0
      Left            =   9150
      TabIndex        =   25
      Top             =   8820
      Width           =   945
   End
   Begin VB.Label lblzlZ 
      Caption         =   "性质"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   4950
      TabIndex        =   24
      Top             =   270
      Width           =   525
   End
   Begin VB.Label lblZl 
      Caption         =   "Label19"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   225
      Left            =   5730
      TabIndex        =   23
      Top             =   270
      Width           =   1155
   End
   Begin VB.Label lblBz 
      Caption         =   "备注"
      Height          =   225
      Left            =   150
      TabIndex        =   22
      Top             =   7110
      Width           =   495
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
      Height          =   345
      Left            =   9240
      TabIndex        =   21
      Top             =   7620
      Width           =   5475
   End
End
Attribute VB_Name = "frmWBXNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public JZ As Integer '基准价比例

Dim timZm As Integer '3价格体系添加 5更新 6删除 7 新保存 8 人工签字 9配件添加 10 配件更新 11配件删除18删除
Dim x2, x5, x6, x7, x8 '保养性质

Private Sub chkA6_Click()
''''If chkA6.Value = 1 Then
''''    frmCai.Visible = True
''''ElseIf Left(comA0.Text, 2) = "风冷" And chkA7.Value = 1 Then
''''    frmCai.Visible = False
''''End If
End Sub

Private Sub chkA7_Click()
frmCai.Visible = True
'''If chkA7.Value = 1 And Left(comA0.Text, 2) = "风冷" And chkA6.Value = 0 Then
'''    frmCai.Visible = False
'''    optA8.Value = True
'''End If

If chkA7.Value = 1 And Left(comA0.Text, 2) = "风冷" Then
    chkA7a.Visible = True
End If
End Sub

Private Sub cmdAdd_Click()
On Error Resume Next
Dim hg As Long
If comDX.Text = "主机" Or comDX.Text = "溴化锂" Then
If optA21a.Value = False And optA21b.Value = False And optA21c.Value = False Then
    MsgBox "请选择保养性质！"
    Exit Sub
End If
End If
If Left(comA0.Text, 2) = "风冷" And chkA7.Value = 1 And chkA6.Value = 0 Then
    optA8.Value = True
'''''    If chkA6.Value = 1 And chkA7.Value = 0 And Val(txtA3.Text) > 0 Then
        
'''''    End If
End If
If comDX.Text = "溴化锂" Then
    comA0.Text = "溴化锂"
End If
 '新版本
    timZm = 3
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单"
    mod1.cmd.Parameters("@NBLX") = "价格体系添加"
    mod1.cmd.Parameters("@bh") = lblBid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = comDX.Text '保养对象
    mod1.cmd.Parameters("@mt2") = ""
    mod1.cmd.Parameters("@mt3") = txtPb.Text  '机组品牌
    mod1.cmd.Parameters("@mt4") = txtXH.Text  '机组型号
    mod1.cmd.Parameters("@mt5") = txtXLBH.Text  '系列编号：
    If opt15.Value = True Then
        mod1.cmd.Parameters("@mt6") = "新签"
    ElseIf opt16.Value = True Then
        mod1.cmd.Parameters("@mt6") = "续签"
    Else
        mod1.cmd.Parameters("@mt6") = "新签"
    End If
    mod1.cmd.Parameters("@mt7") = lblHtbh.Caption
    If comDX.Text = "主机" Or comDX.Text = "溴化锂" Then

        If optA21a.Value = True Then
            mod1.cmd.Parameters("@mt8") = optA21a.Caption
        ElseIf optA21b.Value = True Then
            mod1.cmd.Parameters("@mt8") = optA21b.Caption
        ElseIf optA21c.Value = True Then
            mod1.cmd.Parameters("@mt8") = optA21c.Caption
        End If
        mod1.cmd.Parameters("@mt9") = ""
        mod1.cmd.Parameters("@mt10") = comA0.Text '主机类型：
        mod1.cmd.Parameters("@mt11") = comA2.Text '(机组冷量)单位
        If optA8.Value = True Then
            mod1.cmd.Parameters("@mt12") = "拆一端"
        Else
            mod1.cmd.Parameters("@mt12") = "拆二端"
        End If
        mod1.cmd.Parameters("@mt13") = comA15.Text '供热方式：
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
        mod1.cmd.Parameters("@mm1") = Val(txtSL.Text) '数量
        mod1.cmd.Parameters("@mm2") = Val(txtA3.Text) '(蒸发器*)数量
        mod1.cmd.Parameters("@mm3") = Val(txtA5.Text) '(冷凝器*)数量
        mod1.cmd.Parameters("@mm4") = Val(txtA12.Text) '单机组压缩机数量：
        mod1.cmd.Parameters("@mm5") = 0
        mod1.cmd.Parameters("@mm6") = 0
        mod1.cmd.Parameters("@mm7") = 0
        mod1.cmd.Parameters("@mm8") = 0
        mod1.cmd.Parameters("@mm9") = 0
        mod1.cmd.Parameters("@mm10") = Val(txtA1.Text) '机组冷量：
        mod1.cmd.Parameters("@mm11") = Val(comA13.Text) '机组年巡视次数：
        mod1.cmd.Parameters("@mm12") = Val(txtA20.Text) '机组使用时间：
        mod1.cmd.Parameters("@mm13") = 0
        mod1.cmd.Parameters("@mm14") = 0
        mod1.cmd.Parameters("@mm15") = 0
        mod1.cmd.Parameters("@mm16") = 0
        mod1.cmd.Parameters("@mm17") = 0
        mod1.cmd.Parameters("@mm18") = 0
        mod1.cmd.Parameters("@mm19") = 0
        mod1.cmd.Parameters("@mm20") = 0
        mod1.cmd.Parameters("@mb1") = chkA6.Value '蒸发器*
        mod1.cmd.Parameters("@mb2") = chkA7.Value '冷凝器
        mod1.cmd.Parameters("@mb3") = chkA10.Value '物理清洗
        mod1.cmd.Parameters("@mb4") = chkA11.Value '化学清洗
        mod1.cmd.Parameters("@mb5") = chkA7a.Value '清洗翅片
        mod1.cmd.Parameters("@md1") = Null
        mod1.cmd.Parameters("@md2") = Null
        mod1.cmd.Parameters("@md3") = Null
        mod1.cmd.Parameters("@md4") = Null
        mod1.cmd.Parameters("@md5") = Null
    Else
        If comDX.Text = "水泵" Then
                x2 = ""
                If chkB28a.Value = 1 Then
                    x2 = "保养"
                End If
                If chkB28b.Value = 1 Then
                    If x2 <> "" Then
                        x2 = x2 & "+" & chkB28b.Caption
                    Else
                        x2 = chkB28b.Caption
                    End If
                End If
                If chkB28c.Value = 1 Then
                    If x2 <> "" Then
                        x2 = x2 & "+" & chkB28c.Caption
                    Else
                        x2 = chkB28c.Caption
                    End If
                End If
                If chkB28d.Value = 1 Then
                    If x2 <> "" Then
                        x2 = x2 & "+" & chkB28d.Caption
                    Else
                        x2 = chkB28d.Caption
                    End If
                End If
                mod1.cmd.Parameters("@mt8") = x2
        ElseIf comDX.Text = "小机" Then
            
                x5 = ""
                If chkC31a.Value = 1 Then
                    x5 = "保养"
                End If
                If chkC31b.Value = 1 Then
                    If x5 <> "" Then
                        x5 = x5 & "+" & chkC31b.Caption
                    Else
                        x5 = chkC31b.Caption
                    End If
                End If
                If chkC31c.Value = 1 Then
                    If x5 <> "" Then
                        x5 = x5 & "+" & chkC31c.Caption
                    Else
                        x5 = chkC31c.Caption
                    End If
                End If
                If chkC31d.Value = 1 Then
                    If x5 <> "" Then
                        x5 = x5 & "+" & chkC31d.Caption
                    Else
                        x5 = chkC31d.Caption
                    End If
                End If
                mod1.cmd.Parameters("@mt8") = x5
        ElseIf comDX.Text = "小机安装" Then
                mod1.cmd.Parameters("@mt8") = "安装"
        ElseIf comDX.Text = "风机盘管" Then
                x7 = ""
                If chkC37a.Value = 1 Then
                    x7 = "保养"
                End If
                If chkC37b.Value = 1 Then
                    If x7 <> "" Then
                        x7 = x7 & "+" & chkC37b.Caption
                    Else
                        x7 = chkC37b.Caption
                    End If
                End If

                mod1.cmd.Parameters("@mt8") = x7
            
        ElseIf comDX.Text = "空调箱" Then
                x8 = ""
                If chkC51a.Value = 1 Then
                    x8 = "保养"
                End If
                If chkC51b.Value = 1 Then
                    If x8 <> "" Then
                        x8 = x8 & "+" & chkC51b.Caption
                    Else
                        x8 = chkC51b.Caption
                    End If
                End If
                If chkC51c.Value = 1 Then
                    If x8 <> "" Then
                        x8 = x8 & "+" & chkC51c.Caption
                    Else
                        x8 = chkC51c.Caption
                    End If
                End If
                mod1.cmd.Parameters("@mt8") = x8
        End If
        mod1.cmd.Parameters("@mt9") = comB25.Text
        mod1.cmd.Parameters("@mt10") = comB26.Text
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
        mod1.cmd.Parameters("@mm1") = Val(txtB22.Text)
        mod1.cmd.Parameters("@mm2") = Val(txtB23.Text)
        mod1.cmd.Parameters("@mm3") = Val(comB27.Text)
        mod1.cmd.Parameters("@mm4") = Val(txtC29.Text)
        mod1.cmd.Parameters("@mm5") = Val(txtC30.Text)
        mod1.cmd.Parameters("@mm6") = Val(txtC32.Text)
        mod1.cmd.Parameters("@mm7") = Val(txtC33.Text)
        mod1.cmd.Parameters("@mm8") = Val(txtC35.Text)
        mod1.cmd.Parameters("@mm9") = Val(txtC36.Text)
        mod1.cmd.Parameters("@mm10") = Val(txtC38.Text)
        mod1.cmd.Parameters("@mm11") = Val(txtC39.Text)
        mod1.cmd.Parameters("@mm12") = Val(txtC52.Text)
        mod1.cmd.Parameters("@mm13") = 0
        mod1.cmd.Parameters("@mm14") = 0
        mod1.cmd.Parameters("@mm15") = 0
        mod1.cmd.Parameters("@mm16") = 0
        mod1.cmd.Parameters("@mm17") = 0
        mod1.cmd.Parameters("@mm18") = 0
        mod1.cmd.Parameters("@mm19") = 0
        mod1.cmd.Parameters("@mm20") = Val(txtSL.Text)
        mod1.cmd.Parameters("@mb1") = 0
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

Private Sub cmdBack_Click()

Dim tt As String
Dim ii As Integer
On Error Resume Next
Me.Visible = False
If frmGxBiao.Visible = True Then
    frmGxBiao.Enabled = True
    frmGxBiao.ZOrder 0
ElseIf Dialog.Visible = True Then
    Dialog.ZOrder 0
    Dialog.Enabled = True
ElseIf FMXC.Visible = True Then
    FMXC.ZOrder 0
    FMXC.Enabled = True
    mod1.BTZ = 6
End If
End Sub

Private Sub cmdBJ_Click()
On Error Resume Next
Dim hg: Dim Lhg
Dim oo As Integer: Dim ii As Integer
Dim Odx As String: Dim aL As Integer
Dim N1, N2, N3, N4, N5, N6, N7, N8, N9, N10, CP1, CP2, CP3, CP5, CP6, CP7, CP8, CP81, CP82, CP83, CP31, CP32, CP33, CP34, CP51, CP52, CP53, CP54, CP71, CP72
Dim D9, Z9, Z10, Z11, Z12, Z13, Z14, D11, D15, D18, Z18, DD18, ZZ18, D19, Z8, D4, Z4, L4, CP32A
Dim C1, C2, C3
Dim LZhi, OCol
dtgJG.Visible = False
'价格单元清空
For oo = 1 To dtgJG.Rows
    dtgJG.Col = 41
    dtgJG.Row = oo
    dtgJG.Text = ""
Next
dtgJG.Row = 1
Odx = dtgJG.Text
D9 = 0: N2 = 0: N3 = 0: N4 = 0: N5 = 0: D18 = 0: Z18 = 0: Y18 = 0: DD18 = 0: ZZ18 = 0: D19 = 0: N8 = 0: N9 = 0: Z9 = 0: N10 = 0: N4 = 1: D4 = "": Z4 = "": L4 = 1

'水泵巡视算法特殊，先算出总的机组数量
For oo = 1 To dtgN.Rows
    dtgN.Row = oo
        dtgN.Col = 6
    If InStr(1, dtgN.Text, "巡视") > 0 Then
        dtgN.Col = 5
        CP32A = CP32A + Val(dtgN.Text)
    End If
Next

For oo = 1 To dtgJG.Rows
    dtgJG.Col = 1
   dtgJG.Row = oo
    CP31 = 0
    If dtgJG.Row = dtgJG.Rows - 5 Then '如果对象为空，则停止计算
        Exit For
    ElseIf Odx = "" And dtgJG.Text <> "" Then
        Odx = dtgJG.Text
    ElseIf dtgJG.Text <> Odx Then
            Select Case Odx
            Case "主机"
                N1 = Round(D9 / aL, 0)
                N1 = Int(N1 / 200)
                If N1 > 3 And N1 <= 4 Then N1 = 3
                If N1 > 4 Then N1 = 4
                N5 = mod1.UpInt(DD18 / ZZ18)
                N5 = 1 + 0.1 * (N5 - 1)
                N6 = D19
                N7 = ZZ18
                N8 = Round(N8 / N7, 3)
                
                
'''''''                            MsgBox ("N1=" & N1 & ",N2=" & N2 & "  N3=" & N3 & "  N5=" & N5 & " N6=" & N6 & _
'''''''                   "  N7=" & N7 & "  N8=" & N8 & "  N9=" & N9 & "  N10=" & N10 & "  N4=" & N4) & _
'''''''                   Chr(13) & "  CP1=(200 * (N1 + N2 * N3 * N4) * N7 + 600 * mod1.UpInt(N7 / 3)) * N5 + 350 * (N6 + mod1.UpInt(N7 / 3) - 1) + 600 * (2 + N7) * N8 + 600 * mod1.UpInt(N10 / 3) * N9=" & CP1 & _
'''''''                   Chr(13) & "CP2=(200 * (N1 + 2 * N2 * N3 * N4) * N7 + 600 * mod1.UpInt(N7 / 3)) + 350 * (N6 + mod1.UpInt(N7 / 3) - 1) + 600 * (2 + N7) * N8 + 600 * mod1.UpInt(N10 / 3) * N9=" & CP2

                '主机维保
                CP1 = (200 * (N1 + N2 * N3 * N4) * N7 + 600 * mod1.UpInt(N7 / 3)) * N5 + 350 * (N6 + mod1.UpInt(N7 / 3) - 1) + 600 * (2 + N7) * N8 + 600 * mod1.UpInt(N10 / 3) * N9
                dtgJG.Row = dtgJG.Row - 1: OCol = dtgJG.Col: dtgJG.Col = 6
                If dtgJG.Text = "一次性保养" Then
                    CP1 = CP1 * 0.7
                ElseIf dtgJG.Text = "大修" Then
                    CP1 = "人工报价"
                End If
                dtgJG.Row = dtgJG.Row + 1: dtgJG.Col = OCol
                dtgJG.Row = dtgJG.Row - 1
                OCol = dtgJG.Col: dtgJG.Col = 41
                dtgJG.Text = CP1
                dtgJG.Row = dtgJG.Row + 1: dtgJG.Col = OCol
            Case "溴化锂"
                N1 = Round(D9 / aL, 0)
                N1 = Int(N1 / 200)
                If N1 > 3 And N1 <= 4 Then N1 = 3
                If N1 > 4 Then N1 = 4
                N5 = mod1.UpInt(DD18 / ZZ18)
                N5 = 1 + 0.1 * (N5 - 1)
                N6 = D19
                N7 = ZZ18
                N8 = Round(N8 / N7, 3)
                '溴化锂
                CP2 = (200 * (N1 + 2 * N2 * N3 * N4) * N7 + 600 * mod1.UpInt(N7 / 3)) + 350 * (N6 + mod1.UpInt(N7 / 3) - 1) + 600 * (2 + N7) * N8 + 600 * mod1.UpInt(N10 / 3) * N9
                If dtgJG.Text = "一次性保养" Then
                    CP2 = CP2 * 0.7
                End If
                dtgJG.Row = dtgJG.Row + 1: dtgJG.Col = OCol
                dtgJG.Row = dtgJG.Row - 2
                OCol = dtgJG.Col: dtgJG.Col = 41
                dtgJG.Text = CP2
                dtgJG.Row = dtgJG.Row + 1: dtgJG.Col = OCol
            Case "空调箱"
                N1 = UpInt(C1 / 3): N4 = UpInt(C1 / 5): N5 = N5 / LZhi
                N5 = UpInt(N5 / 50000)
                If N5 > 2 Then
                    If mod1.Bm = "配送中心" Then
                        If N5 > 2 Then
                            N5 = Val(InputBox("请商务支持手工键入风量系数（N5）："))
                        End If
                    Else
                        MsgBox "由于空调箱过大，须查看现场后报价！"
                    End If
                End If
                dtgJG.Col = 6: dtgJG.Row = dtgJG.Row - 1
                If InStr(1, dtgJG.Text, "保养") > 0 Then
                    CP81 = (250 + 150 * N1) * N2 * N5
                End If
                If InStr(1, dtgJG.Text, "巡视") > 0 Then
                    CP82 = (N1 * N3 * 150) * N5
                End If
                If InStr(1, dtgJG.Text, "应急") > 0 Then
                    CP83 = N4 * 400
                End If
                    CP8 = CP81 + CP82 + CP83
                OCol = dtgJG.Col: dtgJG.Col = 41
                dtgJG.Text = CP8
                dtgJG.Row = dtgJG.Row + 1: dtgJG.Col = OCol
            Case "水泵"
                If N1 <= 45 Then
                    N6 = 1
                Else
                    N6 = 1.5
                End If
                If N1 < 45 Then
                    N1 = 1
                Else
                    N1 = 1.5
                End If
                dtgJG.Col = 6: dtgJG.Row = dtgJG.Row - 1
                If InStr(1, dtgJG.Text, "保养") > 0 Then
                    CP31 = 100 * N1 * N2 * N3
                    If CP31 <= 400 Then
                        CP31 = 400
                    End If
                End If
                If InStr(1, dtgJG.Text, "巡视") > 0 Then
''''''                    N5 = UpInt(N3 / 20)
''''''                    CP32 = 150 * N4 * N5
                    N5 = UpInt(CP32A / 20)
                    CP32 = Round(150 * N4 * N5 * N3 / CP32A, 0)
                End If
                If InStr(1, dtgJG.Text, "大修") > 0 Then
                    CP33 = 400 * N6 * N7 * (N8 + 1) / 2 + 100 * N3
                End If
                If InStr(1, dtgJG.Text, "急修") > 0 Then
                    CP34 = (400 * N6 * N7 * (N8 + 1) / 2 + 100 * N3) / 5
                End If

                    CP3 = CP31 + CP32 + CP33 + CP34
                 OCol = dtgJG.Col: dtgJG.Col = 41
                dtgJG.Text = CP3
                dtgJG.Row = dtgJG.Row + 1: dtgJG.Col = OCol
            Case "小机"
                dtgJG.Col = 6: dtgJG.Row = dtgJG.Row - 1
                If InStr(1, dtgJG.Text, "保养") > 0 Then
                    CP51 = (200 * (1 + (N3 - 1) / 4)) * N1
                End If
                If InStr(1, dtgJG.Text, "巡视") > 0 Then
                    CP52 = (150 * (1 + (N3 - 1) / 100)) * N2
                End If
                If InStr(1, dtgJG.Text, "应急") > 0 Then
                    CP53 = 100 * (1 + (N3 - 1) / 8)
                End If
                If InStr(1, dtgJG.Text, "移机") > 0 Then
                    CP54 = 0
                End If

                    CP5 = CP51 + CP52 + CP53 + CP54
                OCol = dtgJG.Col: dtgJG.Col = 41
                dtgJG.Text = CP5
                dtgJG.Row = dtgJG.Row + 1: dtgJG.Col = OCol
            Case "小机安装"
                CP6 = 420 * (1 + (N1 - 1) / 3) + N2 * 100
                 OCol = dtgJG.Col: dtgJG.Col = 6
                OCol = dtgJG.Col: dtgJG.Col = 41
                dtgJG.Text = CP6
                dtgJG.Row = dtgJG.Row + 1: dtgJG.Col = OCol
            Case "风机盘管"
                dtgJG.Col = 6: dtgJG.Row = dtgJG.Row - 1
                If InStr(1, dtgJG.Text, "保养") > 0 Then
                    CP71 = (200 * (1 + (N3 - 1) / 4)) * N1
                End If
                If InStr(1, dtgJG.Text, "巡视") > 0 Then
                    CP72 = (250 * (1 + (N3 - 1) / 100)) * N2
                End If
                    CP7 = Round((CP71 + CP72) / 2, 0)
                 OCol = dtgJG.Col: dtgJG.Col = 6
                OCol = dtgJG.Col: dtgJG.Col = 41
                dtgJG.Text = CP7
                dtgJG.Row = dtgJG.Row + 1: dtgJG.Col = OCol
           End Select


''            MsgBox ("N1=" & N1 & "   N2=" & N2 & " N3=" & N3 & "   N4=" & N4 & "   N5=" & N5 & _
''                    "  CP8=" & CP8)
'            MsgBox ("N1=" & N1 & "  N2=" & N2 & "   N3=" & N3 & "    N4=" & N4 & "   N6=" & N6 & "   N7=" & N7 & "   N8=" & N8 & Chr(13) & "CP3=" & CP3)
            Odx = dtgJG.Text
            CP81 = 0: CP82 = 0: CP83 = 0
            D9 = 0: N1 = 0: N2 = 0: N3 = 0: N4 = 0: N5 = 0: D18 = 0: Z18 = 0: Y18 = 0: DD18 = 0: ZZ18 = 0: D19 = 0
            N8 = 0: N9 = 0: Z9 = 0: N10 = 0: N4 = 1: D4 = "": Z4 = "": L4 = 1: Z11 = 0: N6 = 0: N7 = 0: CP1 = 0: CP2 = 0: CP3 = 0: CP5 = 0: cp4 = 0: CP5 = 0: CP6 = 0: CP7 = 0: CP8 = 0
            CP81 = 0: CP82 = 0: CP83 = 0: C1 = 0
    End If



    
        If Odx = "主机" Or Odx = "溴化锂" Then
            'N1
            dtgJG.Col = 9
            Z9 = Val(dtgJG.Text)
            dtgJG.Col = 5
            Z9 = Z9 * Val(dtgJG.Text)
            dtgJG.Col = 10
            Z10 = dtgJG.Text
            If Z10 <> "USRT" Then
                Z9 = Z9 / 3.516
            End If
            D9 = D9 + Z9
            dtgJG.Col = 5
            aL = aL + Val(dtgJG.Text)
            'N2
            dtgJG.Col = 11
            If dtgJG.Text = "True" Then
                Z11 = 1
            Else
                Z11 = 0
            End If
            dtgJG.Col = 12
            Z12 = Val(dtgJG.Text)
            D11 = Z11 * Z12
            dtgJG.Col = 13
            If dtgJG.Text = "True" Then
                Z13 = 1
            Else
                Z13 = 0
            End If
            dtgJG.Col = 14
            Z14 = Val(dtgJG.Text)
            D11 = D11 + Z13 * Z14
            If N2 < D11 Then
                N2 = D11
            End If
            'N2
            dtgJG.Col = 15
            If dtgJG.Text = "拆一端" Then
                D15 = 1
            Else
                D15 = 2
            End If
            dtgJG.Col = 8
            If Left(dtgJG.Text, 2) = "风冷" Then
                D15 = 1
            End If
            If N3 < D15 Then
                N3 = D15
            End If
            'N5
            dtgJG.Col = 5
            Z18 = Val(dtgJG.Text)
            dtgJG.Col = 18
            D18 = Val(dtgJG.Text)
            DD18 = DD18 + D18 * Z18
            ZZ18 = ZZ18 + Z18

            'N6
            dtgJG.Col = 19
            If D19 < Val(dtgJG.Text) Then
                D19 = Val(dtgJG.Text)
            End If
            'N8
            dtgJG.Col = 8
            Select Case dtgJG.Text
            Case "水冷机组"
                Z8 = 1
            Case "风冷热泵机组"
                Z8 = 2
            Case "地源热泵机组"
                Z8 = 1.5
            Case "风冷单冷机组"
                Z8 = 1
            Case "溴化锂"
                dtgJG.Col = 20
                If dtgJG.Text = "直燃式" Then
                    Z8 = 2
                ElseIf dtgJG.Text = "蒸气式" Then
                    Z8 = 1
                End If
            End Select
            dtgJG.Col = 5
            Z8 = Z8 * Val(dtgJG.Text)
            N8 = N8 + Z8
            'N9
            dtgJG.Col = 7
            If dtgJG.Text = "新签" Then
                Z9 = 1
                dtgJG.Col = 5
                N10 = N10 + Val(dtgJG.Text)
            Else
                Z9 = 0
            End If
            If N9 < Z9 Then
                N9 = Z9
            End If
            'N4
            dtgJG.Col = 16
            D4 = dtgJG.Text
            dtgJG.Col = 17
            Z4 = dtgJG.Text
            If Not (D4 = "True" And Z4 = "True") Then
                L4 = 1
            Else
                L4 = 2
            End If
            If N4 < L4 Then
                N4 = L4
            End If
        ElseIf Odx = "空调箱" Then
            dtgJG.Col = 5 '数量
            C1 = C1 + Val(dtgJG.Text)
            dtgJG.Col = 33 '保养次数
            C2 = Val(dtgJG.Text)
            If N2 < C2 Then
                N2 = C2
            End If
            dtgJG.Col = 34 '询视次数
            C3 = Val(dtgJG.Text)
            If N3 < C3 Then
                N3 = C3
            End If
            dtgJG.Col = 36 '风量
            LZhi = LZhi + 1
            N5 = N5 + Val(dtgJG.Text)
        ElseIf Odx = "水泵" Then
            dtgJG.Col = 22
            N1 = Val(dtgJG.Text)
            dtgJG.Col = 24
            If dtgJG.Text = "国产" Then
                N2 = 1
            Else
                N2 = 2
            End If
            dtgJG.Col = 5
            N3 = Val(dtgJG.Text)
            dtgJG.Col = 23
            N4 = Val(dtgJG.Text)
            dtgJG.Col = 25
            If dtgJG.Text = "卧式" Then
                N7 = 1
            Else
                N7 = 1.5
            End If
            dtgJG.Col = 26
            N8 = Val(dtgJG.Text)
        ElseIf Odx = "小机" Then
            dtgJG.Col = 27
            N1 = Val(dtgJG.Text)
            dtgJG.Col = 28
            N2 = Val(dtgJG.Text)
            dtgJG.Col = 5
            N3 = Val(dtgJG.Text)
        ElseIf Odx = "小机安装" Then
            dtgJG.Col = 5
            N1 = Val(dtgJG.Text)
            dtgJG.Col = 30
            N2 = Val(dtgJG.Text)
        ElseIf Odx = "风机盘管" Then
             dtgJG.Col = 5
            N3 = Val(dtgJG.Text)
            dtgJG.Col = 31
            N1 = Val(dtgJG.Text)
            dtgJG.Col = 32
            N2 = Val(dtgJG.Text)
        End If

Next


'2倍基准价(小机相反，）
'''''''''hg = 0: Lhg = 0 '人工成本总合
'''''''''For oo = 1 To dtgJG.Rows
'''''''''    dtgJG.Col = 1
'''''''''    If Trim(dtgJG.Text) = "小机" Then
'''''''''        dtgJG.Col = 41: dtgJG.Row = oo
'''''''''        'dtgJG.Text = Round(Val(dtgJG.Text) / mod1.JiZ1, 2)
'''''''''        Lhg = Lhg + Round(Val(dtgJG.Text) / 2, 2)
'''''''''        If dtgJG.Text = 0 Then
'''''''''            dtgJG.Text = ""
'''''''''        End If
'''''''''        hg = Round(hg + Val(dtgJG.Text), 2)
'''''''''    Else
'''''''''        dtgJG.Col = 41: dtgJG.Row = oo
'''''''''        Lhg = Lhg + Round(Val(dtgJG.Text), 2)
'''''''''        dtgJG.Text = Round(Val(dtgJG.Text) / mod1.JiZ1, 2)
'''''''''        If dtgJG.Text = 0 Then
'''''''''            dtgJG.Text = ""
'''''''''        End If
'''''''''        hg = Round(hg + Val(dtgJG.Text), 2)
'''''''''
'''''''''    End If
'''''''''Next
dtgJG.Visible = True
txt2.Text = hg
txt1.Text = Lhg
End Sub

Private Sub cmdD_Click()
Dim ii As Integer
Dim tt As String
On Error Resume Next
tt = "select htbh from htping where hid=" & Val(lblHtbh.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
If mod1.HTP.Fields(0).Value <> "HMNEW" Then
    Exit Sub
End If
If lblYwy.Caption <> mod1.DName Then Exit Sub
ii = MsgBox("是否删除此询价单？", vbYesNo + vbQuestion, "Hello")
If ii = vbNo Then
    Exit Sub
End If
timZm = 18 '删除合同
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单"
    mod1.cmd.Parameters("@NBLX") = "删除"
    mod1.cmd.Parameters("@bh") = lblBid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(lblZl.Caption)
    mod1.cmd.Parameters("@mt2") = ""
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
    mod1.cmd.Parameters("@mm1") = 0
    mod1.cmd.Parameters("@mm2") = Val(lblHtbh.Caption)
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
    mod1.cmd.Parameters("@md1") = Null
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
End Sub

Private Sub cmdDel_Click()
Dim ii As Integer
On Error Resume Next

ii = MsgBox("是否删除此条记录?", vbInformation + vbYesNo, "询问")
If ii = vbNo Then Exit Sub

 '新版本
    timZm = 6
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单"
    mod1.cmd.Parameters("@NBLX") = "价格体系删除"
    mod1.cmd.Parameters("@bh") = lblWid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = ""
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
    mod1.cmd.Parameters("@mm1") = 0
    mod1.cmd.Parameters("@mm2") = 0
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
    mod1.cmd.Parameters("@md1") = Null
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

Private Sub cmdDing_Click()
Dim tt As String
On Error Resume Next
If OptT1.Value = False And optT2.Value = False Then
    Exit Sub
End If
If optT2.Value = True And txtQM.Text = "" Then
    MsgBox ("请您一定要告诉拒绝我的理由!  :) ")
    Exit Sub
End If
timZm = 8 '人工签字
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单"
    mod1.cmd.Parameters("@NBLX") = "人工签字"
    mod1.cmd.Parameters("@bh") = Val(lblBid.Caption)
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(lblYwy.Caption)
    mod1.cmd.Parameters("@mt2") = Trim(lblUid.Caption)
    mod1.cmd.Parameters("@mt3") = Trim(comXmmc.Text)
    mod1.cmd.Parameters("@mt4") = Trim(lblHtbh.Caption)
    mod1.cmd.Parameters("@mt5") = Trim(lblZl.Caption)
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
    mod1.cmd.Parameters("@mt8") = ""
    mod1.cmd.Parameters("@mt9") = ""
    mod1.cmd.Parameters("@mt10") = ""
    mod1.cmd.Parameters("@mt11") = ""
    mod1.cmd.Parameters("@mt12") = ""
    mod1.cmd.Parameters("@mt13") = Trim(lblZl.Caption) '性质
    mod1.cmd.Parameters("@mt14") = lblFwid.Caption
    mod1.cmd.Parameters("@mt15") = ""
    mod1.cmd.Parameters("@mt16") = ""
    mod1.cmd.Parameters("@mt17") = ""
    mod1.cmd.Parameters("@mt18") = ""
    mod1.cmd.Parameters("@mt19") = ""
    If optT2.Value = True Then
        lblLc.Caption = 3
    End If
    mod1.cmd.Parameters("@mt20") = lblQM(Val(lblLc.Caption) - 1).Caption
    mod1.cmd.Parameters("@mt21") = ""
    
    mod1.cmd.Parameters("@mt22") = ""
    mod1.cmd.Parameters("@mt23") = ""
    mod1.cmd.Parameters("@mt24") = ""
    mod1.cmd.Parameters("@mt25") = ""
    mod1.cmd.Parameters("@mlt1") = txtQM.Text '评审建议
    mod1.cmd.Parameters("@mlt2") = ""
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(lblLc.Caption)
    mod1.cmd.Parameters("@mm2") = Val(lblFwid.Caption)
    mod1.cmd.Parameters("@mm3") = 0
    mod1.cmd.Parameters("@mm4") = 0
    mod1.cmd.Parameters("@mm5") = 0
    mod1.cmd.Parameters("@mm6") = 0
    mod1.cmd.Parameters("@mm7") = 0
    mod1.cmd.Parameters("@mm8") = 0
    mod1.cmd.Parameters("@mm9") = 0
    mod1.cmd.Parameters("@mm10") = Val(txt1.Text) '人工价格
    mod1.cmd.Parameters("@mm11") = 0
    mod1.cmd.Parameters("@mm12") = 0
    mod1.cmd.Parameters("@mm13") = 0
    mod1.cmd.Parameters("@mm14") = 0
    mod1.cmd.Parameters("@mm15") = 0
    mod1.cmd.Parameters("@mm16") = Val(txt2.Text) + Val(txtLyf.Text) '基准价格+运费=人工合计
    mod1.cmd.Parameters("@mm17") = Val(txtLJhg.Text) '材料基准价
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



Private Sub cmdEx_Click()
Dim tt As String
Dim Mid As Long
On Error Resume Next
''''''txtNew.Col = 0
''''''Mid = Val(txtNew.Text)
''''''tt = "update bjxt set bid=" & Val(lblBid.Caption) & ",uid='" & mod1.DHid & "',mid=" & Mid
''''''Set mod1.HTP = CreateObject("adodb.recordset")
''''''mod1.HTP.Open tt, mod1.workBD, adOpenForwardOnly, adLockReadOnly, adCmdText
''''''mod1.HTP.Close
''''''Set mod1.HTP = Nothing
''''''
''''''    OLE1.SourceDoc = "c:\work\demo\hmxp9000\" & "bjxt.xls"
''''''    OLE1.Action = 1
''''''    OLE1.DoVerb (-2)
If comLx.Text = "小机末端空调箱保养" Then
    frmWBXT2.Show
    frmWBXT2.ZOrder 0
    Call frmWBXT2.Qing
Else
frmWBXT.Show
frmWBXT.ZOrder 0
Call frmWBXT.Qing
End If
End Sub

Private Sub cmdGB_Click()
frmLED.Visible = False
End Sub

Private Sub cmdGx_Click()
On Error Resume Next
Dim hg As Long
If Left(comA0.Text, 2) = "风冷" And chkA7.Value = 1 And chkA6.Value = 0 Then
    optA8.Value = True
End If

 '新版本
    timZm = 5
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单"
    mod1.cmd.Parameters("@NBLX") = "价格体系更新"
    mod1.cmd.Parameters("@bh") = lblWid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = comDX.Text '保养对象
    mod1.cmd.Parameters("@mt2") = ""
    mod1.cmd.Parameters("@mt3") = txtPb.Text  '机组品牌
    mod1.cmd.Parameters("@mt4") = txtXH.Text  '机组型号
    mod1.cmd.Parameters("@mt5") = txtXLBH.Text  '系列编号：
    If opt15.Value = True Then
        mod1.cmd.Parameters("@mt6") = "新签"
    ElseIf opt16.Value = True Then
        mod1.cmd.Parameters("@mt6") = "续签"
    Else
        mod1.cmd.Parameters("@mt6") = "新签"
    End If
    mod1.cmd.Parameters("@mt7") = lblHtbh.Caption
    If comDX.Text = "主机" Or comDX.Text = "溴化锂" Then

        If optA21a.Value = True Then
            mod1.cmd.Parameters("@mt8") = optA21a.Caption
        ElseIf optA21b.Value = True Then
            mod1.cmd.Parameters("@mt8") = optA21b.Caption
        ElseIf optA21c.Value = True Then
            mod1.cmd.Parameters("@mt8") = optA21c.Caption
        End If
        mod1.cmd.Parameters("@mt9") = ""
        mod1.cmd.Parameters("@mt10") = comA0.Text '主机类型：
        mod1.cmd.Parameters("@mt11") = comA2.Text '(机组冷量)单位
        If optA8.Value = True Then
            mod1.cmd.Parameters("@mt12") = "拆一端"
        Else
            mod1.cmd.Parameters("@mt12") = "拆二端"
        End If
        mod1.cmd.Parameters("@mt13") = comA15.Text '供热方式：
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
        mod1.cmd.Parameters("@mm1") = Val(txtSL.Text) '数量
        mod1.cmd.Parameters("@mm2") = Val(txtA3.Text) '(蒸发器*)数量
        mod1.cmd.Parameters("@mm3") = Val(txtA5.Text) '(冷凝器*)数量
        mod1.cmd.Parameters("@mm4") = Val(txtA12.Text) '单机组压缩机数量：
        mod1.cmd.Parameters("@mm5") = 0
        mod1.cmd.Parameters("@mm6") = 0
        mod1.cmd.Parameters("@mm7") = 0
        mod1.cmd.Parameters("@mm8") = 0
        mod1.cmd.Parameters("@mm9") = 0
        mod1.cmd.Parameters("@mm10") = Val(txtA1.Text) '机组冷量：
        mod1.cmd.Parameters("@mm11") = Val(comA13.Text) '机组年巡视次数：
        mod1.cmd.Parameters("@mm12") = Val(txtA20.Text) '机组使用时间：
        mod1.cmd.Parameters("@mm13") = 0
        mod1.cmd.Parameters("@mm14") = 0
        mod1.cmd.Parameters("@mm15") = 0
        mod1.cmd.Parameters("@mm16") = 0
        mod1.cmd.Parameters("@mm17") = 0
        mod1.cmd.Parameters("@mm18") = 0
        mod1.cmd.Parameters("@mm19") = 0
        mod1.cmd.Parameters("@mm20") = 0
        mod1.cmd.Parameters("@mb1") = chkA6.Value '蒸发器*
        mod1.cmd.Parameters("@mb2") = chkA7.Value '冷凝器
        mod1.cmd.Parameters("@mb3") = chkA10.Value '物理清洗
        mod1.cmd.Parameters("@mb4") = chkA11.Value '化学清洗
        mod1.cmd.Parameters("@mb5") = chkA7a.Value '清洗翅片
        mod1.cmd.Parameters("@md1") = Null
        mod1.cmd.Parameters("@md2") = Null
        mod1.cmd.Parameters("@md3") = Null
        mod1.cmd.Parameters("@md4") = Null
        mod1.cmd.Parameters("@md5") = Null
    Else
        If comDX.Text = "水泵" Then
                x2 = ""
                If chkB28a.Value = 1 Then
                    x2 = "保养"
                End If
                If chkB28b.Value = 1 Then
                    If x2 <> "" Then
                        x2 = x2 & "+" & chkB28b.Caption
                    Else
                        x2 = chkB28b.Caption
                    End If
                End If
                If chkB28c.Value = 1 Then
                    If x2 <> "" Then
                        x2 = x2 & "+" & chkB28c.Caption
                    Else
                        x2 = chkB28c.Caption
                    End If
                End If
                If chkB28d.Value = 1 Then
                    If x2 <> "" Then
                        x2 = x2 & "+" & chkB28d.Caption
                    Else
                        x2 = chkB28d.Caption
                    End If
                End If
                mod1.cmd.Parameters("@mt8") = x2
        ElseIf comDX.Text = "小机" Then
                x5 = ""
                If chkC31a.Value = 1 Then
                    x5 = "保养"
                End If
                If chkC31b.Value = 1 Then
                    If x5 <> "" Then
                        x5 = x5 & "+" & chkC31b.Caption
                    Else
                        x5 = chkC31b.Caption
                    End If
                End If
                If chkC31c.Value = 1 Then
                    If x5 <> "" Then
                        x5 = x5 & "+" & chkC31c.Caption
                    Else
                        x5 = chkC31c.Caption
                    End If
                End If
                If chkC31d.Value = 1 Then
                    If x5 <> "" Then
                        x5 = x5 & "+" & chkC31d.Caption
                    Else
                        x5 = chkC31d.Caption
                    End If
                End If
                mod1.cmd.Parameters("@mt8") = x5
        ElseIf comDX.Text = "小机安装" Then
                mod1.cmd.Parameters("@mt8") = "安装"
        ElseIf comDX.Text = "风机盘管" Then
                x7 = ""
                If chkC37a.Value = 1 Then
                    x7 = "保养"
                End If
                If chkC37b.Value = 1 Then
                    If x7 <> "" Then
                        x7 = x7 & "+" & chkC37b.Caption
                    Else
                        x7 = chkC37b.Caption
                    End If
                End If

                mod1.cmd.Parameters("@mt8") = x7
        ElseIf comDX.Text = "空调箱" Then
                x8 = ""
                If chkC51a.Value = 1 Then
                    x8 = "保养"
                End If
                If chkC51b.Value = 1 Then
                    If x8 <> "" Then
                        x8 = x8 & "+" & chkC51b.Caption
                    Else
                        x8 = chkC51b.Caption
                    End If
                End If
                If chkC51c.Value = 1 Then
                    If x8 <> "" Then
                        x8 = x8 & "+" & chkC51c.Caption
                    Else
                        x8 = chkC51c.Caption
                    End If
                End If
                mod1.cmd.Parameters("@mt8") = x8
                
        End If
        mod1.cmd.Parameters("@mt9") = comB25.Text
        mod1.cmd.Parameters("@mt10") = comB26.Text
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
        mod1.cmd.Parameters("@mm1") = Val(txtB22.Text)
        mod1.cmd.Parameters("@mm2") = Val(txtB23.Text)
        mod1.cmd.Parameters("@mm3") = Val(comB27.Text)
        mod1.cmd.Parameters("@mm4") = Val(txtC29.Text)
        mod1.cmd.Parameters("@mm5") = Val(txtC30.Text)
        mod1.cmd.Parameters("@mm6") = Val(txtC32.Text)
        mod1.cmd.Parameters("@mm7") = Val(txtC33.Text)
        mod1.cmd.Parameters("@mm8") = Val(txtC35.Text)
        mod1.cmd.Parameters("@mm9") = Val(txtC36.Text)
        mod1.cmd.Parameters("@mm10") = Val(txtC38.Text)
        mod1.cmd.Parameters("@mm11") = Val(txtC39.Text)
        mod1.cmd.Parameters("@mm12") = Val(txtC52.Text)
        mod1.cmd.Parameters("@mm13") = 0
        mod1.cmd.Parameters("@mm14") = 0
        mod1.cmd.Parameters("@mm15") = 0
        mod1.cmd.Parameters("@mm16") = 0
        mod1.cmd.Parameters("@mm17") = 0
        mod1.cmd.Parameters("@mm18") = 0
        mod1.cmd.Parameters("@mm19") = 0
        mod1.cmd.Parameters("@mm20") = Val(txtSL.Text)
        mod1.cmd.Parameters("@mb1") = 0
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

Private Sub cmdHt_Click()

If mod1.DName = "张砚纯" Then
    Exit Sub
End If
If mod1.DName = "彭海翔" And lblYwy.Caption <> mod1.DName Then '彭海翔只能打开自己的合同
    MsgBox "哈哈！"
    MsgBox "你想干嘛？"
    Exit Sub
End If
mod1.BTZ = 6
If FMXC.Visible = True And Val(FMXC.lblMHid.Caption) = Val(lblHtbh.Caption) Then
    Me.Visible = False
    FMXC.Enabled = True
    FMXC.ZOrder 0
Else

        Call modNewHT.NewMQing
        
        Call modNewHT.NewMBound(Val(lblHtbh.Caption))
        If FMXC.Visible = True Then '如果打开成功,则隐藏自己.
            Me.Visible = False
            FMXC.ZOrder 0
        End If
End If
    FMXC.cmdMQm(0).Visible = True
    FMXC.lblMQM(0).Visible = True
    FMXC.lblMTm(0).Visible = True
    FMXC.ZOrder 0
End Sub

Private Sub cmdLadd_Click()
On Error Resume Next
Dim hg As Long

If Val(txtLsl.Text) = 0 Then
    MsgBox "请确认数量!"
    txtSL.SetFocus
    Exit Sub
End If
If comJLB.Text <> "A" And comJLB.Text <> "B" Or txtJLBZ.Text = "" Then
    MsgBox "请选择产品的类型"
    Exit Sub
End If


                                   '新版本
    timZm = 9
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单"
    mod1.cmd.Parameters("@NBLX") = "配件添加"
    mod1.cmd.Parameters("@bh") = lblHtbh.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = lblBid.Caption
    mod1.cmd.Parameters("@mt2") = "配件"
    mod1.cmd.Parameters("@mt3") = txtLpb.Text  '机组品牌
    mod1.cmd.Parameters("@mt4") = txtLjbh.Text  '规格型号
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
    mod1.cmd.Parameters("@mt8") = txtLjmc.Text '零件名称
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
    mod1.cmd.Parameters("@mt20") = comJLB.Text
    mod1.cmd.Parameters("@mt21") = ""
    mod1.cmd.Parameters("@mt22") = ""
    mod1.cmd.Parameters("@mt23") = ""
    mod1.cmd.Parameters("@mt24") = ""
    mod1.cmd.Parameters("@mt25") = txtLDW.Text '单位
    mod1.cmd.Parameters("@mlt1") = txtLBz.Text
    mod1.cmd.Parameters("@mlt2") = ""
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtLsl.Text) '数量
    mod1.cmd.Parameters("@mm2") = Val(txtJLBZ.Text) '基准值
    mod1.cmd.Parameters("@mm3") = 0
    mod1.cmd.Parameters("@mm4") = 0
    mod1.cmd.Parameters("@mm5") = 0
    mod1.cmd.Parameters("@mm6") = 0
    mod1.cmd.Parameters("@mm7") = 0
    mod1.cmd.Parameters("@mm8") = 0
    mod1.cmd.Parameters("@mm9") = 0
    mod1.cmd.Parameters("@mm10") = 0
    mod1.cmd.Parameters("@mm11") = Val(txtL1.Text)
    mod1.cmd.Parameters("@mm12") = Val(txtL2.Text)
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
    mod1.cmd.Parameters("@md1") = Null
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
        cmdAdd.Enabled = False
        cmdJG.Enabled = False
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


    Set mod1.cmd = Nothing





End Sub

Private Sub cmdLDel_Click()
Dim ii As Integer
On Error Resume Next
If mod1.VLP = 2 Or mod1.VLP = 3 And mod1.DName <> "马晓聪" Then
    MsgBox "You are a Pig!"
    End
End If
ii = MsgBox("是否删除此条记录?", vbInformation + vbYesNo, "询问")
If ii = vbYes Then
   
     timZm = 11
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单"
    mod1.cmd.Parameters("@NBLX") = "配件删除"
    mod1.cmd.Parameters("@bh") = lblHtbh.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = lblBid.Caption
    mod1.cmd.Parameters("@mt2") = "配件"
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
    mod1.cmd.Parameters("@mm1") = Val(lblLid.Caption)
    mod1.cmd.Parameters("@mm2") = 0
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
    mod1.cmd.Parameters("@md1") = Null
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
        cmdDel.Enabled = False
        cmdJG.Enabled = False
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

Private Sub cmdLGx_Click()
On Error Resume Next

If Val(txtLsl.Text) = 0 Then
    Exit Sub
End If

    timZm = 10
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单"
    mod1.cmd.Parameters("@NBLX") = "配件更新"
    mod1.cmd.Parameters("@bh") = lblHtbh.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = lblBid.Caption
    mod1.cmd.Parameters("@mt2") = "配件"
    mod1.cmd.Parameters("@mt3") = txtLpb.Text  '机组品牌
    mod1.cmd.Parameters("@mt4") = txtLjbh.Text   '机组型号
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
    mod1.cmd.Parameters("@mt8") = txtLjmc.Text '零件名称
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
    mod1.cmd.Parameters("@mt20") = comJLB.Text
    mod1.cmd.Parameters("@mt21") = ""
    mod1.cmd.Parameters("@mt22") = ""
    mod1.cmd.Parameters("@mt23") = ""
    mod1.cmd.Parameters("@mt24") = ""
    mod1.cmd.Parameters("@mt25") = ""
    mod1.cmd.Parameters("@mlt1") = txtLBz.Text
    mod1.cmd.Parameters("@mlt2") = ""
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtLsl.Text) '数量
    mod1.cmd.Parameters("@mm2") = Val(lblLid.Caption)
    mod1.cmd.Parameters("@mm3") = Val(txtJLBZ.Text)
    mod1.cmd.Parameters("@mm4") = 0
    mod1.cmd.Parameters("@mm5") = 0
    mod1.cmd.Parameters("@mm6") = 0
    mod1.cmd.Parameters("@mm7") = 0
    mod1.cmd.Parameters("@mm8") = 0
    mod1.cmd.Parameters("@mm9") = 0
    mod1.cmd.Parameters("@mm10") = 0
    mod1.cmd.Parameters("@mm11") = Val(txtL1.Text)
    mod1.cmd.Parameters("@mm12") = Val(txtL2.Text)
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
    mod1.cmd.Parameters("@md1") = Null
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
        cmdAdd.Enabled = False
        cmdJG.Enabled = False
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If
    cmdJgx.Enabled = False

    Set mod1.cmd = Nothing

End Sub


Private Sub cmdLqing_Click()
txtLpb.Text = ""
txtLjbh.Text = ""
txtLjmc.Text = ""
txtLsl.Text = ""
txtLDW.Text = ""
txtLBz.Text = ""
'''txtLyf.Text = ""
'''txtLadr.Text = ""
txtL1.Text = ""
txtL2.Text = ""
'''txtLhg.Text = ""
'''txtLJhg.Text = ""
lblLid.Caption = ""
End Sub

Private Sub cmdMod_Click()
If mod1.DName = "倪旭" Or mod1.DName = "张砚纯" Then
    cmdSave.Enabled = True
    txtZu.Locked = False
End If
If lblLcUid.Caption = mod1.DHid And Val(lblLc.Caption) < 3 Then
    cmdSave.Enabled = True
    cmdD.Enabled = True
    
    If Val(lblLc.Caption) = 1 Then
        frmED.Visible = True
        txtBz.Locked = False

    End If
    If Val(lblLc.Caption) = 2 Then
        frmLED.Visible = True
        txtLyf.Locked = False
        txtLadr.Locked = False
    End If
End If
If mod1.DName = "" And lblLcRen.Caption = "吴金荣" Or mod1.DName = "马晓聪" Then
    txt2.Locked = False
    txt1.Locked = False
    frmLED.Visible = True
    cmdSave.Enabled = True
    txtBz.Locked = False
End If
If mod1.DName = "马晓聪" Then '马晓聪可以修改成本，并将成本导入合同
    lblLc.Caption = 3
End If
End Sub

Private Sub cmdPje_Click()
Dim tt As String
On Error Resume Next
Pje.Show
Set Pje.adoPje = CreateObject("adodb.recordset")
tt = "select trq,ywy,zn,bz,tf from pizu where bh='" & lblBid.Caption & "' and yid=43 order by pid desc"
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

Private Sub cmdQH_Click()
Me.Visible = False
        Call modBJD.BJDWBQing
        Call modBJD.BJDBound(lblBid.Caption, lblZl.Caption)
        Call modBJD.wbxjLocked
        frmWBXJ.Show
        frmWBXJ.lblLcUid.Caption = FMXC.txtXYwy.ToolTipText
        frmWBXJ.lblLcRen.Caption = FMXC.txtXYwy.Text
End Sub

Private Sub cmdQm_Click(Index As Integer)
Dim ii As String
Dim tt As String
Dim Tywy As String '单子流转到下一人的姓名
Dim Tuid As String
Dim Oywy As String '原来流转人的名字
Dim Ouid As String '原来流转人的工号

On Error Resume Next

If Index + 1 <> lblLc.Caption Then '不能在不相干的位置上乱点
    Exit Sub
End If
'If Index = 0 And cmdSave.Enabled = True And lblLc.Caption = 0 Then
If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If
If lblLcUid.Caption <> mod1.DHid Then
    MsgBox "此处应由" & lblLcRen.Caption & "签字! 请您不要再点"
    Exit Sub
End If

'''''If txtZu.Text = "" Then
'''''    cmdSave.Enabled = True
'''''    MsgBox "没有选择工程部组长！"
'''''    Exit Sub
'''''End If

frmQm.Visible = True
If lblLc.Caption = 1 Then
    optT2.Enabled = False
    OptT1.Value = True
Else
    optT2.Enabled = True
    OptT1.Enabled = True
    OptT1.Value = False
    optT2.Value = False
End If
Exit Sub

           


End Sub

Private Sub cmdQm_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tt As Integer
On Error Resume Next
If Button = 2 And lblQM(Index).Caption = "业务员确认" And Val(lblLc.Caption) = 100 And lblYwy.Caption = mod1.DName Then
'''''''''    tt = "select lc from htping where hid=" & Val(lblHtbh.Caption)
'''''''''    Set mod1.HTP = CreateObject("adodb.recordset")
'''''''''    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''''''    If IsNull(mod1.HTP.Fields("lc").Value) = True Then
'''''''''        Exit Sub
'''''''''    End If
    If Val(lblHLC.Caption) < 2 Then
        Me.frmQm.Visible = True
        Me.OptT1.Enabled = False
        Me.optT2.Enabled = True
        Me.optT2.Value = True
        lblLc.Caption = 4
        cmdDing.Enabled = True
    End If
End If
End Sub


Private Sub cmdSave_Click()
On Error Resume Next
'Call cmdBJ_Click


 '新版本
    timZm = 7
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单"
    mod1.cmd.Parameters("@NBLX") = "新保存"
    mod1.cmd.Parameters("@bh") = lblBid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = ""
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
    mod1.cmd.Parameters("@mt8") = ""
    mod1.cmd.Parameters("@mt9") = ""
    mod1.cmd.Parameters("@mt10") = ""
    mod1.cmd.Parameters("@mt11") = comWD.Text '维保年限
    mod1.cmd.Parameters("@mt12") = txtZu.Text '组长
    mod1.cmd.Parameters("@mt13") = txtZu.ToolTipText
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
    mod1.cmd.Parameters("@mt25") = Trim(txtLadr.Text)
    mod1.cmd.Parameters("@mlt1") = txtBz.Text
    mod1.cmd.Parameters("@mlt2") = txtDxnr.Text
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtWc.Text) '维保年限
    mod1.cmd.Parameters("@mm2") = Val(txt1.Text) '人工价
    mod1.cmd.Parameters("@mm3") = Val(txt2.Text) '基准价
    mod1.cmd.Parameters("@mm4") = Val(txtLhg.Text) '材料单价
    mod1.cmd.Parameters("@mm5") = Val(txtLJhg.Text) '材料基准价
    mod1.cmd.Parameters("@mm6") = Val(txtLyf.Text) '运费
    mod1.cmd.Parameters("@mm7") = Val(lblFwid.Caption)
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
    mod1.cmd.Parameters("@md1") = Null
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

Private Sub comA0_Click()
frmCai.Visible = True
If Left(comA0.Text, 2) = "风冷" Then
    chkA10.Caption = "清水冲洗"
    If chkA7.Value = 1 And chkA6.Value = 0 Then
        '''''frmCai.Visible = False
        'optA8.Value = True
        chkA7a.Visible = True
    End If
Else
    chkA10.Caption = "物理清洗"
End If

''''''''''If chkA7.Value = 1 And Left(comA0.Text, 2) = "风冷" And chkA6.Value = 0 Then
''''''''''    frmCai.Visible = False
''''''''''    optA8.Value = True
''''''''''End If
'''''''
'''''''If chkA7.Value = 1 And Left(comA0.Text, 2) = "风冷" Then
'''''''    chkA7a.Visible = True
'''''''End If
End Sub


Private Sub comB27_Click()
frmJN.Visible = True
End Sub

Private Sub comDX_Click()
Call Jqing
frmM1.Visible = False
frmM2.Visible = False
frmM3.Visible = False
frmM5.Visible = False
frmNewF.Visible = False
frmA.Enabled = False
frmB.Enabled = False
frmC.Enabled = False
frmD.Enabled = False
lblSl.Caption = "数量"
Select Case comDX.Text
Case "主机"
    frmM1.Visible = True
    frmXH.Visible = False
    frmNewF.Visible = True
Case "溴化锂"
    frmM1.Visible = True
    frmXH.Visible = True
    frmNewF.Visible = True
    comA0.Text = "溴化锂"
Case "水泵"
    frmM2.Visible = True
Case "电机"
    frmM3.Visible = True
Case "小机"
    frmM5.Visible = True
    lblSl.Caption = "内机数量"
    frmA.Enabled = True
Case "小机安装"
    frmM5.Visible = True
    lblSl.Caption = "内机数量"
    frmB.Enabled = True
Case "风机盘管"
    frmM5.Visible = True
    frmC.Enabled = True
Case "空调箱"
    frmM5.Visible = True
    frmD.Enabled = True
End Select
End Sub




'''''Private Sub comJLB_Click()
'''''If comJLB.Text = "A" Then
'''''    txtJLBZ.Text = Format(mod1.JiZ5A, "0.00")
'''''ElseIf comJLB.Text = "B" Then
'''''    txtJLBZ.Text = Format(mod1.JiZ5B, "0.00")
'''''End If
'''''txtL2.Text = Round(Val(txtL1.Text) / Val(txtJLBZ.Text), 2)
'''''End Sub


Private Sub dtgJG_Click()
Dim OCol As Integer
On Error Resume Next

dtgN.Col = dtgJG.Col: dtgN.Row = dtgJG.Row

Call Jqing
lblSl.Caption = "数量"
frmM1.Visible = False
frmM2.Visible = False
frmM3.Visible = False
frmM5.Visible = False
frmA.Enabled = False
frmB.Enabled = False
frmC.Enabled = False
frmD.Enabled = False
OCol = dtgN.Col
dtgN.Col = 1: comDX.Text = dtgN.Text
dtgN.Col = 2: txtPb.Text = dtgN.Text
dtgN.Col = 3: txtXH.Text = dtgN.Text
dtgN.Col = 4: txtXLBH.Text = dtgN.Text
dtgN.Col = 5: txtSL.Text = Val(dtgN.Text)
dtgN.Col = 6
Select Case comDX.Text
Case "主机"
    If dtgN.Text = "维保" Then
        optA21a.Value = True
    ElseIf dtgN.Text = "一次性保养" Then
        optA21b.Value = True
    ElseIf dtgN.Text = "大修" Then
        optA21c.Value = True
    End If
    frmM1.Visible = True
Case "溴化锂"
    If dtgN.Text = "维保" Then
        optA21a.Value = True
    ElseIf dtgN.Text = "一次性保养" Then
        optA21b.Value = True
    ElseIf dtgN.Text = "大修" Then
        optA21c.Value = True
    End If
    frmM1.Visible = True
Case "水泵"
        If InStr(1, dtgN.Text, "保养") > 0 Then
            chkB28a.Value = 1
        End If
        If InStr(1, dtgN.Text, "巡视") > 0 Then
            chkB28b.Value = 1
        End If
        If InStr(1, dtgN.Text, "大修") > 0 Then
            chkB28c.Value = 1
        End If
        If InStr(1, dtgN.Text, "急修") > 0 Then
            chkB28d.Value = 1
        End If
    frmM2.Visible = True
Case "小机"
    frmM5.Visible = True
    
        If InStr(1, dtgN.Text, "保养") > 0 Then
            chkC31a.Value = 1
        End If
        If InStr(1, dtgN.Text, "巡视") > 0 Then
            chkC31b.Value = 1
        End If
        If InStr(1, dtgN.Text, "应急") > 0 Then
            chkC31c.Value = 1
        End If
        If InStr(1, dtgN.Text, "移机") > 0 Then
            chkC31d.Value = 1
        End If
        lblSl.Caption = "内机数量"
        frmA.Enabled = True
Case "小机安装"
    frmM5.Visible = True
        lblSl.Caption = "内机数量"
        frmB.Enabled = True
Case "风机盘管"
        If InStr(1, dtgN.Text, "保养") > 0 Then
            chkC37a.Value = 1
        End If
        If InStr(1, dtgN.Text, "巡视") > 0 Then
            chkC37b.Value = 1
        End If
        frmC.Enabled = True
Case "空调箱"
        If InStr(1, dtgN.Text, "保养") > 0 Then
            chkC51a.Value = 1
        End If
        If InStr(1, dtgN.Text, "巡视") > 0 Then
            chkC51b.Value = 1
        End If
        If InStr(1, dtgN.Text, "应急") > 0 Then
            chkC51c.Value = 1
        End If
        frmD.Enabled = True
End Select
dtgN.Col = 7
If dtgN.Text = "新签" Then
    opt15.Value = True
Else
    opt16.Value = True
End If
dtgN.Col = 8: comA0.Text = dtgN.Text
If Left(comA0.Text, 2) = "风冷" Then
    chkA10.Caption = "清水冲洗"
Else
    chkA10.Caption = "物理清洗"
End If
dtgN.Col = 9: txtA1.Text = dtgN.Text
dtgN.Col = 10: comA2.Text = dtgN.Text
dtgN.Col = 11
If dtgN.Text = "True" Then
    chkA6.Value = 1
Else
    chkA6.Value = 0
End If
dtgN.Col = 12: txtA3.Text = Val(dtgN.Text)
dtgN.Col = 13
If dtgN.Text = "True" Then
    chkA7.Value = 1
'''    If Left(comA0.Text, 2) = "风冷" Then
'''        frmCai.Visible = False
'''    Else
'''        frmCai.Visible = True
'''    End If
    If Left(comA0.Text, 2) = "风冷" Then
        chkA7a.Visible = True
    Else
        chkA7a.Visible = False
    End If
Else
    chkA7.Value = 0
End If
dtgN.Col = 14: txtA5.Text = Val(dtgN.Text)
dtgN.Col = 15
If dtgN.Text = "拆一端" Then
    optA8.Value = True
ElseIf dtgN.Text = "拆二端" Then
    optA9.Value = True
End If
dtgN.Col = 16
If dtgN.Text = "True" Then
    chkA10.Value = 1
Else
    chkA10.Value = 0
End If
dtgN.Col = 17
If dtgN.Text = "True" Then
    chkA11.Value = 1
Else
    chkA11.Value = 0
End If
dtgN.Col = 18: txtA12.Text = Val(dtgN.Text)
dtgN.Col = 19: comA13.Text = Val(dtgN.Text)
dtgN.Col = 20: comA15.Text = dtgN.Text
dtgN.Col = 21: txtA20.Text = Val(dtgN.Text)

dtgN.Col = 22: txtB22.Text = Val(dtgN.Text)
dtgN.Col = 23: txtB23.Text = Val(dtgN.Text)
dtgN.Col = 24: comB25.Text = dtgN.Text
dtgN.Col = 25: comB26.Text = dtgN.Text
dtgN.Col = 26: comB27.Text = Val(dtgN.Text)

dtgN.Col = 27: txtC29.Text = Val(dtgN.Text)
dtgN.Col = 28: txtC30.Text = Val(dtgN.Text)
dtgN.Col = 29: txtC32.Text = Val(dtgN.Text)
dtgN.Col = 30: txtC33.Text = Val(dtgN.Text)
dtgN.Col = 31: txtC35.Text = Val(dtgN.Text)
dtgN.Col = 32: txtC36.Text = Val(dtgN.Text)
dtgN.Col = 33: txtC38.Text = Val(dtgN.Text)
dtgN.Col = 34: txtC39.Text = Val(dtgN.Text)
dtgN.Col = 36: txtC52.Text = Val(dtgN.Text)
dtgN.Col = 37:
If dtgN.Text = "True" Then
    chkA7a.Value = 1
Else
    chkA7a.Value = 0
End If

dtgN.Col = 35
lblWid.Caption = Val(dtgN.Text)

End Sub

Private Sub dtgMa_Click()
On Error Resume Next
dtgMn.Row = dtgMa.Row

tt = "SELECT 机组品牌 , 机组型号, 零件名称, 数量, ldw, 基准单价, 基准合计, liD, 成本单价, 合计 From XunJiaMxView where bid=" & Bid & " order by lid"
dtgMn.Col = 1
txtLpb.Text = Trim(dtgMn.Text)
dtgMn.Col = 2
txtLjbh.Text = Trim(dtgMn.Text)
dtgMn.Col = 3
txtLjmc.Text = Trim(dtgMn.Text)
dtgMn.Col = 4
txtLsl.Text = Val(dtgMn.Text)
dtgMn.Col = 5
txtLDW.Text = Trim(dtgMn.Text)
dtgMn.Col = 6
txtL2.Text = Val(dtgMn.Text) '基准单价
dtgMn.Col = 8
lblLid.Caption = Val(dtgMn.Text)
dtgMn.Col = 9
txtL1.Text = Val(dtgMn.Text)
dtgMn.Col = 11
comJLB.Text = Trim(dtgMn.Text)
dtgMn.Col = 12
txtJLBZ.Text = Format(Val(dtgMn.Text), "0.00")
End Sub

Private Sub dtgNew_DblClick()
Call frmWBXT.Qing
frmWBXT.Show
frmWBXT.ZOrder 0
End Sub


Private Sub Form_Click()
frmQm.Visible = False

End Sub

Private Sub Form_Load()
Dim ii As Integer
Dim Ra
Dim La
Dim oo As Integer

On Error Resume Next
Me.Left = 0
Me.Top = 0
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
frmM1.Left = 4620: frmM2.Left = 4620: frmM3.Left = 4620: frmM5.Left = 4620
frmM1.Top = 2940: frmM2.Top = 2940: frmM3.Top = 2940: frmM5.Top = 2940
frmM1.Width = 10395: frmM2.Width = 10395: frmM3.Width = 10395: frmM5.Width = 10395
dtgJG.ColWidth(0) = 300: dtgJG.Cols = 50: dtgJG.Rows = 10
dtgJG.Row = 0: dtgJG.Col = 1: dtgJG.Text = "保养对象": dtgJG.Col = 2: dtgJG.Text = "品牌名称": dtgJG.Col = 3: dtgJG.Text = "型号"
dtgJG.Col = 4: dtgJG.Text = "系列编号": dtgJG.Col = 6: dtgJG.Text = "保养性质": dtgJG.Col = 5: dtgJG.Text = "数量": dtgJG.Col = 7: dtgJG.Text = "新签否"
dtgJG.ColWidth(1) = 2000: dtgJG.ColWidth(2) = 2000: dtgJG.ColWidth(3) = 2500: dtgJG.ColWidth(4) = 2500: dtgJG.ColWidth(6) = 2000
For ii = 8 To 40
    dtgJG.ColWidth(ii) = 0
Next
'''''For ii = 0 To LM
'''''    txtZu.AddItem RM(0, ii)
frmQm.Top = 7410
frmQm.Left = 6420
'''''Next
dtgMa.ColWidth(0) = 300: dtgMa.ColWidth(2) = 2500: dtgMa.ColWidth(3) = 2000: dtgMa.ColWidth(8) = 5900
dtgMa.Row = 0: dtgMa.Col = 1: dtgMa.Text = "品牌": dtgMa.Col = 2: dtgMa.Text = "规格": dtgMa.Col = 3: dtgMa.Text = "耗材名称":
dtgMa.Col = 4: dtgMa.Text = "数量": dtgMa.Col = 5: dtgMa.Text = "单位": dtgMa.Col = 6: dtgMa.Text = "单价(基准)": dtgMa.Col = 7: dtgMa.Text = "小计"
dtgMa.Col = 8: dtgMa.Text = "备注"

tt = "select username from worker where zuf=1 and zzf=1"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2)
For oo = 0 To La
    txtZu.AddItem Ra(0, oo)
Next
txtDxnr.Left = 30: txtDxnr.Top = 30
'txtNew.ColWidth(0) = 0
'txtNew.ColWidth(1) = 15000
dtgNew.Row = 0: dtgNew.Col = 0: dtgNew.Text = "操作对象"
dtgNew.Col = 1: dtgNew.Text = "基准金额"
dtgNew.Col = 2: dtgNew.Text = "交通差旅费"
dtgNew.Col = 3: dtgNew.Text = " 承接人"
dtgNew.Col = 4: dtgNew.Text = " 备注"
If mod1.DName = "谢雪梅" Then
    dtgNew.Visible = True
Else
    dtgNew.Visible = False
End If
dtgNew.ColWidth(0) = 6000
dtgNew.ColWidth(4) = 6500
dtgNew.Rows = 30
End Sub


Public Sub Qing()
Dim tt As String
On Error Resume Next
lblZl.Caption = ""
comXmmc.Tag = ""
comXmmc.Text = ""
lblBid.Caption = ""
lblBh.Caption = ""
txtZu.Text = ""
txtZu.ToolTipText = ""
txt1.Text = ""
txt2.Text = ""
lblYwy.Caption = ""
lblUid.Caption = ""
txtWc.Text = ""
comWD.Text = "年"
lblLc.Caption = ""
lblLcRen.Caption = ""
lblLcUid.Caption = ""
lblFwid.Caption = ""
lblNlb.Caption = ""

Call Jqing


txtBz.Text = ""
lblHtbh.Caption = ""
JZ = 0
lblHLC.Caption = ""
lblWid.Caption = ""


frmM1.Visible = False
frmM2.Visible = False
frmM3.Visible = False
frmM5.Visible = False
frmED.Visible = False

txt1.Locked = True
txt2.Locked = True
lblTX.Caption = ""
lblTX.Visible = False

tabJG.Tab = 0
frmLED.Visible = False

txtLyf.Text = ""
txtLadr.Text = ""
txtLhg.Text = ""
txtLJhg.Text = ""
tabJG.Tab = 0
dtgMa.Clear: dtgMa.Cols = 2: dtgMa.FixedCols = 1
comJLB.Text = ""
txtJLBZ.Text = ""
txtDxnr.Text = ""
cmdD.Enabled = False
cmdDing.Enabled = True
Call cmdLqing_Click
If mod1.DName = "谢雪梅" Then
    txtNew.Visible = True
    cmdEx.Visible = True
Else
    txtNew.Visible = False
    cmdEx.Visible = False
End If
frmAdd.Visible = False
End Sub

Public Sub Bound(Bid As Long)
Dim tt As String
Dim oo As Integer: Dim ii As Integer
Dim Ra, Rb, RC, RD, RE
Dim La, Lb, Lc, Ld, Le
On Error GoTo EPP
mod1.BTZ = 36
Call Locked
tt = "declare @hid int;" & _
    "select @hid=cast(htbh as int) from xunjiaD where bid=" & Bid & ";" & _
    "select jz,zl,xid,xmmc,bid,bianhao,zname,hg,jhg,ywy,uid,wc,wd,lc,lcren,lcuid,fwid,nlb,bz,htbh,zuid,chg,cjhg,yf,yfadr,dxnr from XunJiaD where bid=" & Bid & ";" & _
    "select lc from htping where hid=@hid;" & _
    "select dx,jzpb,jzxh,xlbh,sl as 数量,xz as 保养性质,nqf as 新签否,a00,a01,a02,a06,a03,a07,a05,a08,a10,a11,a12,a13,a15,a20," & _
    "b22,b23,b25,b26,b27,c29,c30,c32,c33,c35,c36,c38,c39,wid,c52,a07a from wbView where bid=" & Bid & " order by zid,jzxh,xz,wid;" & _
    "select * from QMRZ where btz=36 and qdbh='" & Bid & "' order by zid;" & _
    "SELECT 机组品牌 , 机组型号, 零件名称, 数量, ldw, 基准单价, 基准合计, liD, 成本单价, 合计,jlb,jlbz From XunJiaMxView where bid=" & Bid & " order by lid"
 
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    RC = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    RD = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    RE = mod1.HTP.GetRows
End If
mod1.HTP.Close
Set mod1.HTP = Nothing
On Error Resume Next
La = UBound(Ra, 2) + 1
Lc = UBound(RC, 2) + 1

Ld = UBound(RD, 2) + 1
Le = UBound(RE, 2) + 1

JZ = Ra(0, 0)
lblZl.Caption = Ra(1, 0)
comXmmc.Tag = Ra(2, 0)
comXmmc.Text = Ra(3, 0)
lblBid.Caption = Ra(4, 0)
lblBh.Caption = Ra(5, 0)
txtZu.Text = Ra(6, 0)
txt1.Text = Ra(7, 0)
txt2.Text = Ra(8, 0)
lblYwy.Caption = Ra(9, 0)
lblUid.Caption = Ra(10, 0)
txtWc.Text = Ra(11, 0)
comWD.Text = Ra(12, 0)
lblLc.Caption = Ra(13, 0)
lblLcRen.Caption = Ra(14, 0)
lblLcUid.Caption = Ra(15, 0)
lblFwid.Caption = Ra(16, 0)
lblNlb.Caption = Ra(17, 0)
txtBz.Text = Ra(18, 0)
lblHtbh.Caption = Ra(19, 0)
txtZu.ToolTipText = Ra(20, 0)
txtLhg.Text = Ra(21, 0)
txtLJhg.Text = Ra(22, 0)
txtLyf.Text = Ra(23, 0)
txtLadr.Text = Ra(24, 0)
txtDxnr.Text = Ra(25, 0)

lblHLC.Caption = Rb(0, 0) '对应合同的流程

'列表明细
dtgJG.Clear
dtgJG.Row = 0: dtgJG.Col = 1: dtgJG.Text = "保养对象": dtgJG.Col = 2: dtgJG.Text = "品牌名称": dtgJG.Col = 3: dtgJG.Text = "型号"
dtgJG.Col = 4: dtgJG.Text = "系列编号": dtgJG.Col = 6: dtgJG.Text = "保养性质": dtgJG.Col = 5: dtgJG.Text = "数量": dtgJG.Col = 7: dtgJG.Text = "新签否"
dtgJG.Col = 41: dtgJG.Text = "价格"
If Lc > 0 Then

    dtgJG.Visible = False
dtgJG.Rows = Lc + 10
Dim Odx As String: Dim Oxh As String: Dim Nrow As Integer: Dim Oxz As String
Odx = RC(0, 0): Nrow = 1: Oxh = RC(2, 0): Oxz = RC(5, 0)
    For oo = 1 To Lc + 20
        dtgJG.Row = Nrow
        If RC(0, oo - 1) = "水泵" And Oxh = RC(2, oo - 1) And Oxh <> "" And Odx <> "" Or _
         RC(0, oo - 1) = "主机" And Oxz = RC(5, oo - 1) And Odx = RC(0, oo - 1) And Odx <> "" Or _
         RC(0, oo - 1) = "溴化锂" And Oxz = RC(5, oo - 1) And Odx = RC(0, oo - 1) And Odx <> "" Or dtgJG.Row = 1 Or _
        (RC(0, oo - 1) = "空调箱" Or RC(0, oo - 1) = "小机" Or RC(0, oo - 1) = "小机安装" Or RC(0, oo - 1) = "风机盘管") And Odx = RC(0, oo - 1) Then
:           'Oxh = Rc(2, oo - 1): Oxz = Rc(5, oo - 1)
            For ii = 1 To 51
                dtgJG.Col = ii
                dtgJG.Text = RC(ii - 1, oo - 1)
            Next
            Nrow = Nrow + 1

        Else
            dtgJG.Row = dtgJG.Row + 1
            For ii = 1 To 51
                dtgJG.Col = ii
                dtgJG.Text = RC(ii - 1, oo - 1)
            Next
            Nrow = Nrow + 2
            
            Odx = RC(0, oo - 1)
            Oxh = RC(2, oo - 1)
            Oxz = RC(5, oo - 1)
        End If
    Next
''''''''''    dtgJG.MergeCol(1) = True
''''''''''    dtgJG.MergeCells = 3
    dtgJG.Visible = True
    
    '复制到内表
    dtgN.Clear: dtgN.Rows = dtgJG.Rows: dtgN.Cols = dtgJG.Cols
    For oo = 0 To dtgN.Rows
        dtgN.Row = oo: dtgJG.Row = oo
        For ii = 0 To dtgN.Cols
            dtgN.Col = ii: dtgJG.Col = ii
            dtgN.Text = dtgJG.Text
        Next
    Next
End If
If mod1.Bm = "配送中心" Then
    lbl1.Visible = True: txt1.Visible = True
    lbl2.Visible = True: txt2.Visible = True
    lblLhg.Visible = True: txtLhg.Visible = True
    lblLJHg.Visible = True: txtLJhg.Visible = True
Else
    lbl1.Visible = False: txt1.Visible = False
    lbl2.Visible = True: txt2.Visible = True
    lblLhg.Visible = False: txtLhg.Visible = False
    lblLJHg.Visible = True: txtLJhg.Visible = True
End If

lblWid.Caption = ""
''''''        Call modBJD.OpenXJAN(LX)
''''''        If Val(lblBid.Caption) >= 6794 Then
''''''            lblQM(2).Caption = "商务支持"
''''''        End If


dtgJG.Row = 1
Call dtgJG_Click

'签字按钮
For oo = 0 To 2
cmdQm(oo).Caption = ""
lblTm(oo).Caption = ""
Next
 For oo = 0 To Ld - 1
    If RD(9, oo) = True Then
       cmdQm(oo).Caption = RD(1, oo)
       lblTm(oo).Caption = RD(4, oo)
    End If
   cmdQm(oo).Tag = RD(8, oo)
Next

'材料列表
dtgMa.Row = 0: dtgMa.Col = 1: dtgMa.Text = "品牌": dtgMa.Col = 2: dtgMa.Text = "规格": dtgMa.Col = 3: dtgMa.Text = "耗材名称":
dtgMa.Col = 4: dtgMa.Text = "数量": dtgMa.Col = 5: dtgMa.Text = "单位": dtgMa.Col = 6: dtgMa.Text = "单价(基准)": dtgMa.Col = 7: dtgMa.Text = "基准小计"
For oo = 1 To Le
    dtgMa.Row = oo
    For ii = 1 To 15
        dtgMa.Col = ii
        dtgMa.Text = RE(ii - 1, oo - 1)
    Next
    
Next
'复制内表
dtgMn.Clear: dtgMn.Rows = dtgMa.Rows: dtgMn.Cols = dtgMa.Cols
For oo = 0 To dtgMa.Rows
    dtgMa.Row = oo: dtgMn.Row = oo
    For ii = 0 To 15
        dtgMa.Col = ii: dtgMn.Col = ii
        dtgMn.Text = dtgMa.Text
    Next
Next

'Call cmdBJ_Click

cmdSave.Enabled = False
If lblYwy.Caption = mod1.DName Then '业务员看不到报价按钮,表格中的基价
    cmdBJ.Visible = False
    dtgJG.ColWidth(41) = 0
Else
    cmdBJ.Visible = True
    dtgJG.ColWidth(41) = 1000
End If
If lblZl.Caption = "维保" Then
    'txtBz.Left = 720: txtBz.Top = 7050
    txtDxnr.Visible = False
Else
    txtDxnr.Visible = True
End If

Exit Sub
EPP:
MsgBox ("网络故障，请退出后重试！")
End
End Sub

Public Sub Jqing()
On Error Resume Next
comDX.Text = ""
txtPb.Text = ""
txtXH.Text = ""
txtXLBH.Text = ""
txtSL.Text = ""
opt15.Value = False
opt16.Value = False

comA0.Text = ""
txtA1.Text = ""
comA2.Text = "USRT"
chkA6.Value = 0
txtA3.Text = 1
chkA7.Value = 0
txtA5.Text = 1
optA8.Value = False
optA9.Value = False
chkA10.Value = 0
chkA11.Value = 0
txtA12.Text = ""
comA13.Text = ""
comA15.Text = ""
txtA20.Text = ""
optA21a.Value = False
optA21b.Value = False
optA21c.Value = False
optC51c.Value = False
txtC52.Text = ""

txtB22.Text = ""
txtB23.Text = ""
comB25.Text = ""
comB26.Text = ""
comB27.Text = ""
chkB28a.Value = False: chkB28b.Value = False: chkB28c.Value = False: chkB28d.Value = False

txtC29.Text = ""
txtC30.Text = ""
chkC31a.Value = False: chkC31b.Value = False: chkC31c.Value = False: chkC31d.Value = False
txtC32.Text = ""
txtC33.Text = ""
txtC35.Text = ""
txtC36.Text = ""
chkC37a.Value = False: chkC37b.Value = False
txtC38.Text = ""
txtC39.Text = ""
lblWid.Caption = ""
chkC51a.Value = 0: chkC51b.Value = 0: chkC51c.Value = 0

frmM1.Visible = False
frmM2.Visible = False
frmM3.Visible = False
frmM5.Visible = False
chkA7a.Value = 0

End Sub

Public Sub MXBound(Bid As Long)
Dim tt As String
Dim oo As Integer: Dim ii As Integer
Dim Ra
Dim La
On Error GoTo EPP
tt = "select dx,jzpb,jzxh,xlbh,sl as 数量,xz as 保养性质,nqf as 新签否,a00,a01,a02,a06,a03,a07,a05,a08,a10,a11,a12,a13,a15,a20," & _
    "b22,b23,b25,b26,b27,c29,c30,c32,c33,c35,c36,c38,c39,wid,c52,a07a from wbView where bid=" & Bid & " order by zid,jzxh,xz,wid"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
If mod1.HTP.BOF = False Then
    Ra = mod1.HTP.GetRows
    La = UBound(Ra, 2) + 1
End If
mod1.HTP.Close
Set mod1.HTP = Nothing
On Error Resume Next

dtgJG.Visible = False
dtgJG.Clear
dtgJG.Rows = La + 20
''''''''''For oo = 1 To la
''''''''''    dtgJG.Row = oo
''''''''''    For ii = 1 To 50
''''''''''        dtgJG.Col = ii
''''''''''        dtgJG.Text = Ra(ii - 1, oo - 1)
''''''''''    Next
''''''''''Next

Dim Odx As String: Dim Oxh As String: Dim Nrow As Integer: Dim Oxz As String
Odx = Ra(0, 0): Nrow = 1: Oxh = Ra(2, 0): Oxz = Ra(5, 0)
    For oo = 1 To La + 20
        dtgJG.Row = Nrow
        If Ra(0, oo - 1) = "水泵" And Oxh = Ra(2, oo - 1) And Oxh <> "" And Odx <> "" Or _
         Ra(0, oo - 1) = "主机" And Oxz = Ra(5, oo - 1) And Odx = Ra(0, oo - 1) And Odx <> "" Or _
         Ra(0, oo - 1) = "溴化锂" And Oxz = Ra(5, oo - 1) And Odx = Ra(0, oo - 1) And Odx <> "" Or dtgJG.Row = 1 Or _
         (Ra(0, oo - 1) = "空调箱" Or Ra(0, oo - 1) = "小机" Or Ra(0, oo - 1) = "小机安装" Or Ra(0, oo - 1) = "风机盘管") And Odx = Ra(0, oo - 1) Then
:           'Oxh = ra(2, oo - 1): Oxz = ra(5, oo - 1)
            For ii = 1 To 50
                dtgJG.Col = ii
                dtgJG.Text = Ra(ii - 1, oo - 1)
            Next
            Nrow = Nrow + 1

        Else
            dtgJG.Row = dtgJG.Row + 1
            For ii = 1 To 50
                dtgJG.Col = ii
                dtgJG.Text = Ra(ii - 1, oo - 1)
            Next
            Nrow = Nrow + 2
            
            Odx = Ra(0, oo - 1)
            Oxh = Ra(2, oo - 1)
            Oxz = Ra(5, oo - 1)
        End If
    Next
''''''''''    dtgJG.MergeCol(1) = True
''''''''''    dtgJG.MergeCells = 3
    dtgJG.Visible = True

dtgJG.Row = 0: dtgJG.Col = 1: dtgJG.Text = "保养对象": dtgJG.Col = 2: dtgJG.Text = "品牌名称": dtgJG.Col = 3: dtgJG.Text = "型号"
dtgJG.Col = 4: dtgJG.Text = "系列编号": dtgJG.Col = 6: dtgJG.Text = "保养性质": dtgJG.Col = 5: dtgJG.Text = "数量": dtgJG.Col = 7: dtgJG.Text = "新签否"
dtgJG.Col = 41: dtgJG.Text = "价格"
dtgJG.Visible = True

    '复制到内表
    dtgN.Clear: dtgN.Rows = dtgJG.Rows: dtgN.Cols = dtgJG.Cols
    For oo = 0 To dtgN.Rows
        dtgN.Row = oo: dtgJG.Row = oo
        For ii = 0 To dtgN.Cols
            dtgN.Col = ii: dtgJG.Col = ii
            dtgN.Text = dtgJG.Text
        Next
    Next
    dtgJG.Row = 1
    Call dtgJG_Click
    

Exit Sub
EPP:
MsgBox ("网络故障，请退出后再试！")
End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmJN.Visible = False
End Sub

Private Sub frmM2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmJN.Visible = False
End Sub

Private Sub lbl2_DblClick()
If mod1.Bm = "商务部" Or mod1.DName = "倪旭" Or mod1.DName = "宋晓炯" Or mod1.DName = "宋晓炯1" Or mod1.DName = "马晓聪" Or mod1.DName = "周春云" Or mod1.DName = "" Or Ywy = "吴金荣" Then
    If lbl1.Visible = False Then
        lbl1.Visible = True
        txt1.Visible = True
        lblLhg.Visible = True
        txtLhg.Visible = True
    Else
        lbl1.Visible = False
        txt1.Visible = False
        lblLhg.Visible = False
        txtLhg.Visible = False
    End If
End If
End Sub

Private Sub lblLJHg_DblClick()
If mod1.Bm = "商务部" Or mod1.DName = "倪旭" Or mod1.DName = "宋晓炯" Or mod1.DName = "宋晓炯1" Or mod1.DName = "马晓聪" Or mod1.DName = "周春云" Or mod1.DName = "" Or Ywy = "吴金荣" Then
    If lbl1.Visible = False Then
        lbl1.Visible = True
        txt1.Visible = True
        lblLhg.Visible = True
        txtLhg.Visible = True
    Else
        lbl1.Visible = False
        txt1.Visible = False
        lblLhg.Visible = False
        txtLhg.Visible = False
    End If
End If
End Sub

Private Sub timQuit_Timer()
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0


If timZm = 7 Then    '新保存
    If mod1.DName = "倪旭" Or mod1.DName = "张砚纯" Then
        Call mod1.refEnvent(1)
    End If
ElseIf timZm = 8 Then '签字
    cmdDing.Enabled = True
    txtQM.Text = ""
    frmQm.Visible = False
    lblTX.Visible = True
    If Dialog.Visible = True Then '更新事务列表
        Call mod1.refEnvent(1)
    End If
    If cmdQm(2).Caption <> "" And FMXC.Visible = True Then '业务员确认后，修改合同上的成本
        If lblZl.Caption = "维保" Then
            FMXC.txtH1.Text = Val(txt2.Text)
        ElseIf lblZl.Caption = "大修" Then
            FMXC.txtH2.Text = Val(txt2.Text)
        ElseIf lblZl.Caption = "工程分包" Then
            FMXC.txtW3.Text = Val(txt2.Text)
        ElseIf lblZl.Caption = "水处理" Then
            FMXC.txtW4.Text = Val(txt2.Text)
        End If
        FMXC.txtFC.Text = Val(FMXC.txtFC.Text) + Val(txtLJhg.Text) '辅材合计
    End If
ElseIf timZm = 9 Or timZm = 10 Or timZm = 11 Then
    Call cmdLqing_Click
    Call Cbound(Val(lblBid.Caption))
ElseIf timZm = 18 Then '删除
    Me.Visible = False
'''''''''    If FMXC.Visible = True Then
'''''''''        If lblZl.Caption = "维保" Then
'''''''''            FMXC.cmdW1.ToolTipText = ""
'''''''''        ElseIf lblZl.Caption = "大修" Then
'''''''''            FMXC.cmdW2.ToolTipText = ""
'''''''''        ElseIf lblZl.Caption = "工程分包" Then
'''''''''            FMXC.cmdW3.ToolTipText = ""
'''''''''        ElseIf lblZl.Caption = "水处理" Then
'''''''''            FMXC.cmdW4.ToolTipText = ""
'''''''''        End If
'''''''''    End If
End If
timQuit.Enabled = False
Me.Enabled = True
End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
On Error Resume Next
timWait.Enabled = False
Me.Enabled = False
tt = "select cf,bz,bh,mm1,mm2,mt1,mt2 from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0
    If timZm = 3 Then '价格体系添加
        Call MXBound(Val(lblBid.Caption))
        Call Jqing
    ElseIf timZm = 5 Then '价格体系更新
        Call MXBound(Val(lblBid.Caption))
        Call Jqing
    ElseIf timZm = 6 Then '价格体系删除
        Call MXBound(Val(lblBid.Caption))
        Call Jqing
    ElseIf timZm = 7 Then '新保存
        cmdSave.Enabled = False
    ElseIf timZm = 8 Then '签名
        If OptT1.Value = True Then
            cmdQm(lblLc.Caption - 1).Caption = mod1.DName
            lblTm(lblLc.Caption - 1).Caption = mod1.DQda
        Else
            For ii = 0 To 2
                cmdQm(ii).Caption = ""
                lblTm(ii).Caption = ""
            Next
        End If
        lblLc.Caption = mod1.WP.Fields("mm1").Value
        lblFwid.Caption = mod1.WP.Fields("mm2").Value
        lblLcRen.Caption = mod1.WP.Fields("mt1").Value
        lblLcUid.Caption = mod1.WP.Fields("mt2").Value
        lblTX.Caption = "下一流程,将跳至" & lblQM(Val(lblLc.Caption) - 1).Caption & ": " & lblLcRen.Caption
    End If

    timWait.Enabled = False

    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then '未完成

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '处理失败
    ii = MsgBox("服务中心在处理您的命令时,发生如下错误:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdJG.Enabled = False
    End If
    timWait.Enabled = False
    Me.Enabled = True
    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdJG.Enabled = False
    End If
    Me.Enabled = True
    Exit Sub

End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub


Private Sub txtL1_LostFocus()
On Error Resume Next
txtL2.Text = Round(Val(txtL1.Text) / Val(txtJLBZ.Text), 2)
End Sub


Private Sub txtZu_Click()
If mod1.DName <> "倪旭" And mod1.DName <> "张砚纯" Then
    txtZu.Text = ""
    txtZu.ToolTipText = ""
    MsgBox "组长必须由工程部经理来指定！"
    Exit Sub
End If
txtZu.ToolTipText = mod1.ZuId(1, txtZu.ListIndex)
End Sub



Public Sub Cbound(Bid As Long)
Dim tt As String
Dim oo As Integer: Dim ii As Integer
Dim hg: Dim jhg
Dim Ra
Dim La
On Error GoTo EPPP
tt = "SELECT 机组品牌 , 机组型号, 零件名称, 数量, ldw, 基准单价, 基准合计, liD, 成本单价, 合计,jlb,jlbz From XunJiaMxView where bid=" & Bid & " order by lid"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
La = UBound(Ra, 2) + 1
dtgMa.Clear: dtgMa.Cols = 2: dtgMa.FixedCols = 1
'材料列表
dtgMa.Row = 0: dtgMa.Col = 1: dtgMa.Text = "品牌": dtgMa.Col = 2: dtgMa.Text = "规格": dtgMa.Col = 3: dtgMa.Text = "耗材名称":
dtgMa.Col = 4: dtgMa.Text = "数量": dtgMa.Col = 5: dtgMa.Text = "单位": dtgMa.Col = 6: dtgMa.Text = "单价(基准)": dtgMa.Col = 7: dtgMa.Text = "基准小计"
For oo = 1 To La
    dtgMa.Row = oo
    For ii = 1 To 15
        dtgMa.Col = ii
        dtgMa.Text = Ra(ii - 1, oo - 1)
    Next
    
Next

'复制内表
dtgMn.Clear: dtgMn.Rows = dtgMa.Rows: dtgMn.Cols = dtgMa.Cols
For oo = 0 To dtgMa.Rows
    dtgMa.Row = oo: dtgMn.Row = oo
    For ii = 0 To 15
        dtgMa.Col = ii: dtgMn.Col = ii
        dtgMn.Text = dtgMa.Text
        If ii = 7 Then
            jhg = jhg + Val(dtgMn.Text)
        End If
        If ii = 10 Then
            hg = hg + Val(dtgMn.Text)
        End If
    Next
Next
txtLhg.Text = hg + Val(txtLyf.Text)
'txtLJhg.Text = Round(Val(txtLhg.Text) / Val(txtJLBZ.Text), 2) + Val(txtLyf.Text)
txtLJhg.Text = jhg + Val(txtLyf.Text)

Exit Sub
EPPP:
MsgBox ("网络故障，请退出后再试！")
End
End Sub

Public Sub Locked()
txtLyf.Locked = True
txtLadr.Locked = True
frmED.Visible = False
frmLED.Visible = False
txtBz.Locked = True
End Sub
