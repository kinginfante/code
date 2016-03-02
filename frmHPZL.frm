VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmHPZL 
   BackColor       =   &H00C0FFC0&
   Caption         =   "货品资料"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15255
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15255
   Begin VB.CommandButton cmdBTD 
      BackColor       =   &H00FFFFC0&
      Caption         =   "替代"
      Height          =   765
      Left            =   10920
      Picture         =   "frmHPZL.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   134
      Top             =   8280
      Width           =   735
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgNN 
      Height          =   675
      Left            =   10230
      TabIndex        =   123
      Top             =   6120
      Visible         =   0   'False
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1191
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdBr 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "浏览"
      Height          =   765
      Left            =   10230
      Picture         =   "frmHPZL.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   119
      Top             =   8280
      Width           =   705
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgL1 
      Height          =   2475
      Left            =   11730
      TabIndex        =   0
      Top             =   5430
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   4366
      _Version        =   393216
      BackColor       =   12648384
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   12648384
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame frm1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5625
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   15225
      Begin VB.CommandButton cmdZ 
         Caption         =   "转"
         Height          =   375
         Left            =   3240
         TabIndex        =   135
         Top             =   80
         Width           =   375
      End
      Begin VB.ComboBox txtL3 
         Height          =   300
         Left            =   4860
         TabIndex        =   126
         Text            =   "Combo3"
         Top             =   2820
         Width           =   2475
      End
      Begin VB.ComboBox txtL2 
         Height          =   300
         Left            =   4860
         TabIndex        =   125
         Text            =   "Combo2"
         Top             =   2310
         Width           =   2475
      End
      Begin VB.ComboBox txtL1 
         Height          =   300
         Left            =   4860
         TabIndex        =   124
         Text            =   "Combo1"
         Top             =   1770
         Width           =   2475
      End
      Begin VB.Frame frmName 
         BackColor       =   &H00FFFFC0&
         Caption         =   "请在列表中选择相应的货品名称"
         Height          =   3375
         Left            =   12150
         TabIndex        =   120
         Top             =   3060
         Width           =   3765
         Begin VB.CommandButton cmdAll 
            Caption         =   "全部显示"
            Height          =   315
            Left            =   2580
            TabIndex        =   127
            Top             =   2970
            Width           =   975
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   120
            TabIndex        =   122
            Top             =   2970
            Width           =   2295
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgName 
            Height          =   2565
            Left            =   90
            TabIndex        =   121
            Top             =   270
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   4524
            _Version        =   393216
            BackColor       =   16777152
            Rows            =   300
            FixedCols       =   0
            BackColorFixed  =   12648384
            BackColorBkg    =   16777152
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.TextBox txtTD 
         Height          =   675
         Left            =   8610
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   118
         Text            =   "frmHPZL.frx":0544
         Top             =   1470
         Width           =   2445
      End
      Begin VB.TextBox txtBm3 
         Height          =   270
         Left            =   1140
         TabIndex        =   113
         Text            =   "Text6"
         Top             =   2840
         Width           =   2445
      End
      Begin VB.TextBox txtBm2 
         Height          =   270
         Left            =   1140
         TabIndex        =   112
         Text            =   "Text5"
         Top             =   2290
         Width           =   2445
      End
      Begin VB.TextBox txtBm1 
         Height          =   270
         Left            =   1140
         TabIndex        =   111
         Text            =   "Text4"
         Top             =   1740
         Width           =   2445
      End
      Begin VB.CheckBox chkJYF 
         BackColor       =   &H00C0FFC0&
         Caption         =   "禁用"
         Height          =   180
         Left            =   8550
         TabIndex        =   106
         Top             =   2730
         Width           =   1995
      End
      Begin VB.TextBox txtTdbh 
         Height          =   270
         Left            =   8580
         TabIndex        =   105
         Text            =   "Text1"
         Top             =   2280
         Width           =   2445
      End
      Begin VB.CommandButton cmdTd 
         Caption         =   "替代"
         Height          =   315
         Left            =   7560
         TabIndex        =   104
         Top             =   2280
         Width           =   765
      End
      Begin VB.TextBox txtYpb 
         Height          =   270
         Left            =   1140
         TabIndex        =   64
         Text            =   "Text1"
         Top             =   3390
         Width           =   2445
      End
      Begin VB.TextBox txtBz 
         Height          =   915
         Left            =   1140
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Text            =   "frmHPZL.frx":054A
         Top             =   4050
         Width           =   6225
      End
      Begin VB.TextBox txtOname 
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   1140
         TabIndex        =   28
         Text            =   "Text13"
         Top             =   640
         Width           =   2445
      End
      Begin VB.TextBox txtBH 
         Height          =   270
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   90
         Width           =   2085
      End
      Begin VB.TextBox txtPartName 
         Height          =   270
         Left            =   1140
         TabIndex        =   26
         Text            =   "Text3"
         Top             =   1190
         Width           =   2445
      End
      Begin VB.TextBox txtEngName 
         Height          =   285
         Left            =   4860
         TabIndex        =   25
         Text            =   "Text4"
         Top             =   105
         Width           =   2445
      End
      Begin VB.TextBox txtGG 
         Height          =   285
         Left            =   4860
         TabIndex        =   24
         Text            =   "Text7"
         Top             =   654
         Width           =   2445
      End
      Begin VB.TextBox txtXN 
         Height          =   285
         Left            =   4860
         TabIndex        =   23
         Text            =   "Text8"
         Top             =   1203
         Width           =   2445
      End
      Begin VB.TextBox txtFF 
         Height          =   285
         Left            =   7710
         TabIndex        =   22
         Text            =   "Text9"
         Top             =   3960
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.TextBox txtPb 
         Height          =   525
         Left            =   8610
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Text            =   "frmHPZL.frx":0551
         Top             =   120
         Width           =   2445
      End
      Begin VB.TextBox txtJz 
         Height          =   555
         Left            =   8610
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Text            =   "frmHPZL.frx":0558
         Top             =   735
         Width           =   2445
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         Height          =   435
         Index           =   2
         Left            =   3780
         Top             =   1140
         Width           =   3585
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         Height          =   435
         Index           =   1
         Left            =   60
         Top             =   540
         Width           =   3585
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         Height          =   435
         Index           =   0
         Left            =   60
         Top             =   1110
         Width           =   3585
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "替代编号"
         Height          =   255
         Left            =   7620
         TabIndex        =   117
         Top             =   1500
         Width           =   885
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "类别3"
         Height          =   315
         Left            =   4110
         TabIndex        =   116
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "类别2"
         Height          =   315
         Left            =   4110
         TabIndex        =   115
         Top             =   2340
         Width           =   825
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "类别1"
         Height          =   315
         Left            =   4110
         TabIndex        =   114
         Top             =   1800
         Width           =   645
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "别名3"
         Height          =   315
         Left            =   330
         TabIndex        =   110
         Top             =   2895
         Width           =   465
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "别名2"
         Height          =   315
         Left            =   330
         TabIndex        =   109
         Top             =   2340
         Width           =   765
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "别名1"
         Height          =   315
         Left            =   330
         TabIndex        =   108
         Top             =   1785
         Width           =   795
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "原厂品牌"
         Height          =   315
         Left            =   240
         TabIndex        =   63
         Top             =   3450
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "编码"
         Height          =   315
         Left            =   450
         TabIndex        =   39
         Top             =   120
         Width           =   1125
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "货品名称"
         Height          =   315
         Left            =   120
         TabIndex        =   38
         Top             =   1230
         Width           =   1125
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "英文名"
         Height          =   315
         Left            =   4020
         TabIndex        =   37
         Top             =   165
         Width           =   1125
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "包装规格"
         Height          =   315
         Left            =   3840
         TabIndex        =   36
         Top             =   705
         Width           =   885
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "产品型号"
         Height          =   315
         Left            =   3840
         TabIndex        =   35
         Top             =   1245
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "使用方法"
         Height          =   315
         Left            =   7740
         TabIndex        =   34
         Top             =   3600
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "适用品牌"
         Height          =   315
         Left            =   7590
         TabIndex        =   33
         Top             =   150
         Width           =   1125
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "适用机组"
         Height          =   315
         Left            =   7590
         TabIndex        =   32
         Top             =   735
         Width           =   1125
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "备注"
         Height          =   315
         Left            =   390
         TabIndex        =   31
         Top             =   4140
         Width           =   645
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "原厂编号"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   675
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFF00&
      Caption         =   "确认"
      Height          =   765
      Left            =   12660
      Picture         =   "frmHPZL.frx":055F
      Style           =   1  'Graphical
      TabIndex        =   107
      Top             =   7530
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Frame frmBr 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5655
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   10845
      Begin VB.CommandButton cmdJQ 
         Caption         =   "近期录入"
         Height          =   285
         Left            =   5250
         TabIndex        =   92
         Top             =   5250
         Width           =   945
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgLP 
         Height          =   5085
         Left            =   60
         TabIndex        =   14
         Top             =   60
         Visible         =   0   'False
         Width           =   10845
         _ExtentX        =   19129
         _ExtentY        =   8969
         _Version        =   393216
         BackColor       =   16777152
         FixedCols       =   0
         BackColorFixed  =   16777152
         BackColorBkg    =   16777152
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame frmXT 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   345
         Left            =   6240
         TabIndex        =   87
         Top             =   5190
         Width           =   3135
         Begin VB.CommandButton cmdXT 
            Caption         =   "查询相同记录"
            Height          =   315
            Left            =   1470
            TabIndex        =   89
            Top             =   30
            Width           =   1395
         End
         Begin VB.ComboBox comMLx 
            Height          =   300
            ItemData        =   "frmHPZL.frx":09A1
            Left            =   90
            List            =   "frmHPZL.frx":09AB
            TabIndex        =   88
            Text            =   "货品名称"
            Top             =   60
            Width           =   1185
         End
      End
      Begin VB.CommandButton cmdT 
         Caption         =   "替代"
         Height          =   285
         Left            =   4320
         TabIndex        =   62
         Top             =   5250
         Width           =   795
      End
      Begin VB.CommandButton cmdGB 
         Caption         =   "关闭"
         Height          =   315
         Left            =   9480
         TabIndex        =   61
         Top             =   5190
         Width           =   735
      End
      Begin VB.CommandButton cmdC 
         Caption         =   "查询"
         Height          =   285
         Left            =   3300
         TabIndex        =   18
         Top             =   5280
         Width           =   945
      End
      Begin VB.ComboBox comLx 
         Height          =   300
         ItemData        =   "frmHPZL.frx":09C3
         Left            =   930
         List            =   "frmHPZL.frx":09D6
         TabIndex        =   17
         Text            =   "货品"
         Top             =   5250
         Width           =   1095
      End
      Begin VB.TextBox txtZ 
         Height          =   285
         Left            =   2040
         TabIndex        =   16
         Top             =   5250
         Width           =   1185
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "查询方式"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   5280
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "复制"
      Height          =   255
      Left            =   9240
      TabIndex        =   103
      Top             =   5700
      Width           =   645
   End
   Begin VB.Frame frmBm 
      BackColor       =   &H00FFFFC0&
      Caption         =   "编辑"
      Height          =   975
      Left            =   30
      TabIndex        =   93
      Top             =   4590
      Visible         =   0   'False
      Width           =   8955
      Begin VB.CommandButton Command3 
         Caption         =   "添加"
         Height          =   315
         Left            =   6180
         TabIndex        =   102
         Top             =   600
         Width           =   915
      End
      Begin VB.CommandButton Command2 
         Caption         =   "添加"
         Height          =   315
         Left            =   2940
         TabIndex        =   101
         Top             =   600
         Width           =   915
      End
      Begin VB.CommandButton Command1 
         Caption         =   "添加"
         Height          =   315
         Left            =   150
         TabIndex        =   100
         Top             =   600
         Width           =   915
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   7200
         TabIndex        =   99
         Top             =   240
         Width           =   1545
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3870
         TabIndex        =   97
         Top             =   240
         Width           =   1785
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1050
         TabIndex        =   95
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "三级编码"
         Height          =   285
         Left            =   6270
         TabIndex        =   98
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "二级编码"
         Height          =   285
         Left            =   2970
         TabIndex        =   96
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "一级编码"
         Height          =   285
         Left            =   120
         TabIndex        =   94
         Top             =   270
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdSC 
      Caption         =   "市场价未录"
      Height          =   315
      Left            =   6300
      TabIndex        =   91
      Top             =   5190
      Width           =   1245
   End
   Begin VB.Frame frmQm 
      BackColor       =   &H00C0FFC0&
      Caption         =   "评审建议"
      ForeColor       =   &H000000FF&
      Height          =   1785
      Left            =   -30
      TabIndex        =   81
      Top             =   7320
      Visible         =   0   'False
      Width           =   6315
      Begin VB.TextBox txtQM 
         BackColor       =   &H00C0FFFF&
         Height          =   1365
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   85
         Top             =   300
         Width           =   4965
      End
      Begin VB.OptionButton OptT1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "同意"
         Height          =   225
         Left            =   5220
         TabIndex        =   84
         Top             =   480
         Width           =   705
      End
      Begin VB.OptionButton OptT2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "拒绝"
         Height          =   195
         Left            =   5220
         TabIndex        =   83
         Top             =   870
         Width           =   675
      End
      Begin VB.CommandButton cmdDing 
         BackColor       =   &H00FF8080&
         Caption         =   "决定"
         Height          =   285
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.Frame frm2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   6375
      Left            =   11040
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   4005
      Begin VB.TextBox txtCb 
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   990
         TabIndex        =   133
         Text            =   "Text4"
         Top             =   3690
         Width           =   2865
      End
      Begin VB.TextBox txtPartName1 
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   990
         TabIndex        =   74
         Text            =   "Text1"
         Top             =   557
         Width           =   2865
      End
      Begin VB.TextBox txtBh1 
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   990
         TabIndex        =   73
         Text            =   "Text1"
         Top             =   90
         Width           =   2865
      End
      Begin VB.TextBox txtGy 
         Height          =   315
         Left            =   30
         TabIndex        =   70
         Top             =   5790
         Width           =   3855
      End
      Begin VB.TextBox txtGy1 
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   47
         Text            =   "Text14"
         Top             =   1024
         Width           =   2865
      End
      Begin VB.TextBox txtJJ1 
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   990
         TabIndex        =   46
         Text            =   "Text15"
         Top             =   1476
         Width           =   2865
      End
      Begin VB.TextBox txtGy2 
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "Text16"
         Top             =   1928
         Width           =   2865
      End
      Begin VB.TextBox txtJJ2 
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   990
         TabIndex        =   44
         Text            =   "Text17"
         Top             =   2380
         Width           =   2865
      End
      Begin VB.TextBox txtGy3 
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "Text18"
         Top             =   2832
         Width           =   2865
      End
      Begin VB.TextBox txtJJ3 
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   990
         TabIndex        =   42
         Text            =   "Text19"
         Top             =   3285
         Width           =   2865
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
         Height          =   1425
         Left            =   30
         TabIndex        =   69
         Top             =   4110
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   2514
         _Version        =   393216
         BackColor       =   12648384
         Rows            =   50
         FixedRows       =   0
         FixedCols       =   0
         BackColorFixed  =   12648384
         BackColorBkg    =   16777152
         WordWrap        =   -1  'True
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         PictureType     =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "成本价"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   30
         TabIndex        =   132
         Top             =   3720
         Width           =   705
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "货品名称"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   72
         Top             =   608
         Width           =   765
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "编号"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   71
         Top             =   150
         Width           =   675
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "供应商1"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   0
         TabIndex        =   53
         Top             =   1066
         Width           =   915
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "价格1"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   0
         TabIndex        =   52
         Top             =   1524
         Width           =   915
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "供应商2"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   0
         TabIndex        =   51
         Top             =   1982
         Width           =   915
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "价格2"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   0
         TabIndex        =   50
         Top             =   2440
         Width           =   915
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "供应商3"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   0
         TabIndex        =   49
         Top             =   2898
         Width           =   915
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "价格3"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   0
         TabIndex        =   48
         Top             =   3360
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdPic 
      Caption         =   "图片"
      Height          =   345
      Left            =   13980
      TabIndex        =   60
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame frm3 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   11070
      TabIndex        =   41
      Top             =   0
      Width           =   3945
      Begin VB.TextBox txtPb1 
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   990
         TabIndex        =   131
         Text            =   "Text4"
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox txtDj 
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   990
         TabIndex        =   129
         Text            =   "Text4"
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox txtPartName2 
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   80
         Text            =   "Text1"
         Top             =   420
         Width           =   2865
      End
      Begin VB.TextBox txtOname2 
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   78
         Text            =   "Text2"
         Top             =   750
         Width           =   2865
      End
      Begin VB.TextBox txtBh2 
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   77
         Text            =   "Text1"
         Top             =   30
         Width           =   2865
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgPic 
         Height          =   975
         Left            =   0
         TabIndex        =   58
         Top             =   2850
         Visible         =   0   'False
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   1720
         _Version        =   393216
         BackColor       =   12648384
         Rows            =   5
         FixedCols       =   0
         BackColorFixed  =   16777152
         ForeColorFixed  =   16576
         BackColorBkg    =   12648384
         SelectionMode   =   1
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox txtMj 
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   990
         TabIndex        =   55
         Text            =   "Text20"
         Top             =   1200
         Width           =   2865
      End
      Begin VB.TextBox txtListPrice 
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   990
         TabIndex        =   54
         Text            =   "Text21"
         Top             =   1545
         Width           =   2865
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "品牌"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   540
         TabIndex        =   130
         Top             =   2340
         Width           =   375
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "单价"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   540
         TabIndex        =   128
         Top             =   1980
         Width           =   405
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "零件名称"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   90
         TabIndex        =   79
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "原厂编号"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   60
         TabIndex        =   76
         Top             =   810
         Width           =   915
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "编号"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   330
         TabIndex        =   75
         Top             =   90
         Width           =   435
      End
      Begin VB.Label lblPic 
         BackStyle       =   0  'Transparent
         Caption         =   "已经正确关联图片"
         ForeColor       =   &H000040C0&
         Height          =   345
         Left            =   30
         TabIndex        =   59
         Top             =   2610
         Visible         =   0   'False
         Width           =   3765
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "面价"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   420
         TabIndex        =   57
         Top             =   1260
         Width           =   405
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "市场指导价"
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   0
         TabIndex        =   56
         Top             =   1635
         Width           =   945
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgLN 
      Height          =   675
      Left            =   8940
      TabIndex        =   12
      Top             =   6480
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1191
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   90
   End
   Begin VB.CommandButton cmdCreate 
      BackColor       =   &H00FFFFC0&
      Caption         =   "新建"
      Height          =   765
      Left            =   12360
      Picture         =   "frmHPZL.frx":0A04
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8280
      Width           =   675
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "删除"
      Height          =   765
      Left            =   13740
      Picture         =   "frmHPZL.frx":0E46
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8280
      Width           =   645
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFC0&
      Caption         =   "返回"
      Height          =   765
      Left            =   14400
      Picture         =   "frmHPZL.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8280
      Width           =   585
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0FFC0&
      Caption         =   "提交"
      Height          =   765
      Left            =   13050
      Picture         =   "frmHPZL.frx":10D2
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8280
      Width           =   675
   End
   Begin VB.CommandButton cmdNQ 
      BackColor       =   &H008080FF&
      Caption         =   "审核"
      Height          =   765
      Left            =   9540
      Picture         =   "frmHPZL.frx":173C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8280
      Width           =   675
   End
   Begin VB.CommandButton cmdMod 
      BackColor       =   &H00C0FFC0&
      Caption         =   "修改"
      Height          =   765
      Left            =   11670
      Picture         =   "frmHPZL.frx":1B7E
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "修改"
      Top             =   8280
      Width           =   675
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN3 
      Height          =   2055
      Left            =   8550
      TabIndex        =   5
      Top             =   7740
      Visible         =   0   'False
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   3625
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN2 
      Height          =   2145
      Left            =   7170
      TabIndex        =   4
      Top             =   7890
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   3784
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN1 
      Height          =   2355
      Left            =   9870
      TabIndex        =   3
      Top             =   7770
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   4154
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgL2 
      Height          =   7275
      Left            =   2820
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   12832
      _Version        =   393216
      BackColor       =   12648384
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   12648384
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgL3 
      Height          =   7275
      Left            =   6000
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   12832
      _Version        =   393216
      BackColor       =   12648384
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   12648384
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtLb1 
      Height          =   285
      Left            =   8430
      TabIndex        =   65
      Text            =   "Text5"
      Top             =   2160
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.TextBox txtLb2 
      Height          =   285
      Left            =   8430
      TabIndex        =   66
      Text            =   "Text6"
      Top             =   2640
      Visible         =   0   'False
      Width           =   2445
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgP 
      Height          =   3405
      Left            =   0
      TabIndex        =   86
      Top             =   5670
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   6006
      _Version        =   393216
      BackColor       =   15728356
      ForeColor       =   8404992
      Rows            =   15
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   16777152
      ForeColorFixed  =   0
      BackColorBkg    =   15728356
      GridColorFixed  =   8404992
      GridColorUnpopulated=   8404992
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.Label lblTx 
      BackStyle       =   0  'Transparent
      Caption         =   "Label15"
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
      Height          =   315
      Left            =   11010
      TabIndex        =   90
      Top             =   7620
      Width           =   4275
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "类别2(网站)"
      Height          =   315
      Left            =   7410
      TabIndex        =   68
      Top             =   2670
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "类别1(网站)"
      Height          =   315
      Left            =   7410
      TabIndex        =   67
      Top             =   2160
      Visible         =   0   'False
      Width           =   1125
   End
End
Attribute VB_Name = "frmHPZL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim B1id As Long
Dim B2id As Long
Dim B3id As Long
Dim Bm1 As String
Dim Bm2 As String
Dim Bm3 As String
Dim Bm As String
Dim HY1 As String
Dim HY2 As String
Dim HY3 As String
Public Pid As Long
Dim timZm As Integer
Public GyId As Integer '确定双击哪个供应商
Dim Jpid As Long '近期翻页ID
Dim frId As Integer '全权限查看时的内容ID

Dim LL As String '录入者
Dim LLUid As String
Dim LCRen As String
Dim LCUid As String
Dim Lc As Integer
Dim Fwid As Long
Dim NPF As Integer '倪工审核否

Dim Lbh As String '自动编号变量
Dim TmpBh As String '当3字头更新编号时,原编号的临时变量



Private Sub chkJYF_Click()
Dim ii As Integer
Dim NR As String
If chkJYF.Value = 1 And (frmHPZL.Visible = True And frmHPBR.Visible = False) Then
    ii = MsgBox("是否确认禁用？", vbYesNo + vbQuestion, "请确认")
    If ii = vbNo Then
        chkJYF.Value = 0
        Exit Sub
    End If
    NR = InputBox("请输入禁用的原因！")
    If NR = "" Then Exit Sub
    
    timZm = 7 '禁用
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "新货品资料"
        mod1.cmd.Parameters("@NBLX") = "禁用"
        mod1.cmd.Parameters("@bh") = txtBh.Text
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = ""
        mod1.cmd.Parameters("@mlt1") = NR
        mod1.cmd.Parameters("@mm1") = 0
        mod1.cmd.Parameters("@mb1") = Null
        mod1.cmd.Parameters("@md1") = Null
        Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
        mod1.cmd.Execute
        mod1.Zid = mod1.cmd.Parameters("@zid").Value
        If mod1.cmd.Parameters("@errch").Value <> "成功" Then
            MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
            If timZm = 2 Then '保存
                cmdSave.Enabled = False
            End If
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
    Exit Sub
End If

      

End Sub

Private Sub cmdBack_Click()
Me.Visible = False
    frmZu.Enabled = True
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0

End If
End Sub

Private Sub cmdBr_Click()
Me.Enabled = False
frmHPBR.Visible = True
frmHPBR.WindowState = 0
frmHPBR.ZOrder 0
'Call frmHPBR.dtgLPFF

End Sub

Private Sub cmdBTD_Click()
If mod1.DName <> "倪东海" Then
    Exit Sub
End If
frmHPTD.Show
frmHPTD.ZOrder 0
Call frmHPTD.Bound

End Sub

Private Sub cmdC_Click()
Dim tt As String
Dim LT1 As String
Dim LT2 As String
Dim LT3 As String

Select Case comLx.Text
Case "货品"
    tt = "select bh,partname,'原厂编号:'+oname+' '+gg+' '+xn+' '+ff+' 适用机组:'+jz,pid from nlpmxc where (partname like '%" & _
    Replace(Replace(txtZ.Text, vbCrLf, "", 1), vbCrLf, "", 1) & "%' or oname like '%" & Replace(Replace(txtZ.Text, vbCrLf, "", 1), vbCrLf, "", 1) & "%' or bh='" & Replace(Replace(txtZ.Text, vbCrLf, "", 1), vbCrLf, "", 1) & "') and delf=1 order by pid desc"
Case "编号"
    If Len(Replace(Replace(txtZ.Text, vbCrLf, "", 1), vbCrLf, "", 1)) = 1 And Val(Replace(Replace(txtZ.Text, vbCrLf, "", 1), vbCrLf, "", 1)) > 0 Then
        tt = "select bh,partname,'原厂编号:'+oname+' '+gg+' '+xn+' '+ff+' 适用机组:'+jz,pid from nlpmxc where left(bh,1)='" & Replace(Replace(txtZ.Text, vbCrLf, "", 1), vbCrLf, "", 1) & "' and delf=1  order by pid desc"
    ElseIf Len(Replace(Replace(txtZ.Text, vbCrLf, "", 1), vbCrLf, "", 1)) = 2 And Val(Replace(Replace(txtZ.Text, vbCrLf, "", 1), vbCrLf, "", 1)) > 0 Then
        tt = "select bh,partname,'原厂编号:'+oname+' '+gg+' '+xn+' '+ff+' 适用机组:'+jz,pid from nlpmxc where left(bh,2)='" & Replace(Replace(txtZ.Text, vbCrLf, "", 1), vbCrLf, "", 1) & "' and delf=1  order by pid desc"
    ElseIf Len(Replace(Replace(txtZ.Text, vbCrLf, "", 1), vbCrLf, "", 1)) = 3 And Val(Replace(txtZ.Text, vbCrLf, "", 1)) > 0 Then
        tt = "select bh,partname,'原厂编号:'+oname+' '+gg+' '+xn+' '+ff+' 适用机组:'+jz,pid from nlpmxc where left(bh,3)='" & Replace(txtZ.Text, vbCrLf, "", 1) & "' and delf=1  order by pid desc"
    ElseIf Len(Replace(txtZ.Text, vbCrLf, "", 1)) = 5 And Val(Replace(txtZ.Text, vbCrLf, "", 1)) > 0 Then
        tt = "select bh,partname,'原厂编号:'+oname+' '+gg+' '+xn+' '+ff+' 适用机组:'+jz,pid from nlpmxc where bh='" & Replace(txtZ.Text, vbCrLf, "", 1) & "'"
    End If
Case "原厂编号"
    tt = "select bh,partname,'原厂编号:'+oname+' '+gg+' '+xn+' '+ff+' 适用机组:'+jz,pid from nlpmxc where oname='" & Replace(txtZ.Text, vbCrLf, "", 1) & "' and delf=1 "
Case "适用品牌"
    tt = "select bh,partname,'原厂编号:'+oname+' '+gg+' '+xn+' '+ff+' 适用机组:'+jz,pid from nlpmxc where pb='" & Replace(txtZ.Text, vbCrLf, "", 1) & "' and delf=1  order by pid desc"
Case "适用机组"
    tt = "select bh,partname,'原厂编号:'+oname+' '+gg+' '+xn+' '+ff+' 适用机组:'+jz,pid from nlpmxc where jz like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' and delf=1  order by pid desc"
Case "分类"
    tt = "select bh,partname,'原厂编号:'+oname+' '+gg+' '+xn+' '+ff,pid from nlpmxc where (lb1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or lb2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%')  and delf=1  order by pid desc"
End Select
If tt = "" Then Exit Sub
Call dtgLPBound(tt)
End Sub

Private Sub cmdCopy_Click()
Clipboard.Clear
Clipboard.SetText dtgP.Clip
End Sub

Private Sub cmdCreate_Click()
Dim tt As String
Dim Ra
Dim Rb
Dim LT As String
If Not (mod1.DName = "货品录入员" Or mod1.DName = "李午阳" Or mod1.DName = "马晓聪") Then Exit Sub
'''''If Len(txtBH.Text) <> 3 Then
'''''    MsgBox "请正确选择货品分类!"
'''''    Exit Sub
'''''End If
'''''
'''''Call Qing
'''''Bm = Bm1 & Bm2 & Bm3
'''''tt = "select top 1 bh from nlpmxc where left(bh,3)='" & Bm & "' order by pid desc"
'''''Set mod1.HTP = CreateObject("adodb.recordset")
'''''mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
'''''If mod1.HTP.BOF = True Then
'''''    mod1.HTP.Close
'''''    Set mod1.HTP = Nothing
'''''    txtBH.Text = Bm1 & Bm2 & Bm3 & "00"
'''''Else
'''''    Ra = mod1.HTP.GetRows
'''''    mod1.HTP.Close
'''''    Set mod1.HTP = Nothing
'''''    tt = "declare @id int;" & _
'''''        "select @id=id from newid where bm='" & Right(Ra(0, 0), 2) & "';" & _
'''''        "select bm from newid where id=@id+1"
'''''    Set mod1.HTP = CreateObject("adodb.recordset")
'''''    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
'''''
'''''    Ra = mod1.HTP.GetRows
'''''    mod1.HTP.Close
'''''    Set mod1.HTP = Nothing
'''''
'''''
'''''''''''     tt = "select bm from newid where id=" & (Ra(0, 0) + 1)
'''''''''''    Set mod1.HTP = CreateObject("adodb.recordset")
'''''''''''    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
'''''''''''    Ra = mod1.HTP.GetRows
'''''''''''    mod1.HTP.Close
'''''''''''    Set mod1.HTP = Nothing
'''''    txtBH.Text = Bm1 & Bm2 & Bm3 & Ra(0, 0)
'''''End If
'''''
''''''先检测是否会重复编号
'''''tt = "select count(bh) from nlpmxc where bh='" & txtBH.Text & "'"
'''''Set mod1.HTP = CreateObject("adodb.recordset")
'''''mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
'''''Rb = mod1.HTP.GetRows
'''''mod1.HTP.Close
'''''Set mod1.HTP = Nothing
'''''If Rb(0, 0) = 1 Then
'''''    MsgBox ("出错，请与马晓聪联系!")
'''''    txtBH.Text = ""
'''''    Exit Sub
'''''End If
'''''
'''''txtLb1.Text = HY2
'''''txtLb2.Text = HY3
'''''If Bm1 = "9" Then
'''''    txtPb.Text = HY2
'''''    txtYpb.Text = HY2
'''''    txtJz.Locked = False
'''''Else
'''''    txtJz.Locked = True
'''''End If
Call Qing
dtgL1.Visible = True
    MsgBox "请正确选择货品分类!"
    Exit Sub
    
If Len(txtBh.Text) <> 1 Then
    MsgBox "请正确选择货品分类!"
    Exit Sub
End If

Call Qing
Bm = Bm1
tt = "select max(bh) from nlpmxc where left(bh,1)='" & Bm & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Rb = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
Bm = Rb(0, 0)
txtBh.Text = Bm
cmdCreate.Visible = False
cmdOK.Visible = True
txtBh.Locked = False
frm1.Enabled = True

Exit Sub


'先检测是否会重复编号
tt = "select count(bh) from nlpmxc where bh='" & txtBh.Text & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Rb = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
If Rb(0, 0) = 1 Then
    MsgBox ("出错，请与马晓聪联系!")
    txtBh.Text = ""
    Exit Sub
End If

txtLb1.Text = HY2
txtLb2.Text = HY3
If Bm1 = "9" Then
    txtPb.Text = HY2
    txtYpb.Text = HY2
    txtJz.Locked = False
Else
    txtJz.Locked = True
End If

Pid = 0
frm1.Enabled = True
frm2.Enabled = False
frm3.Enabled = False
cmdSave.Enabled = True
txtPartName.Locked = False
txtOname.Locked = False
txtEngName.Locked = False
txtBm1.Locked = False
txtBm2.Locked = False
txtBm3.Locked = False
txtL1.Locked = False
txtL2.Locked = False
txtL3.Locked = False
txtGG.Locked = False
txtXN.Locked = False
txtFF.Locked = False
txtYpb.Locked = False

End Sub

Private Sub cmdDel_Click()
Dim ii As Integer
'禁用的单子，如果在审核状态，业务员不能做修改和删除
If chkJYF.Value = 1 And Lc > 1 And mod1.DName <> "倪东海" Then
    Exit Sub
End If
If NPF = 1 And chkJYF.Value = 0 Then
    MsgBox "此单倪工已经审核，不能删除，只能禁用！"
    Exit Sub
End If
If mod1.DName <> "货品录入员" And mod1.DName <> "倪东海" And mod1.DName <> "李午阳" Then
     Exit Sub
End If

If mod1.DName = "货品录入员" Then
    If Not ((Left(txtBh.Text, 1) = "3" Or Left(txtBh.Text, 1) = "A")) Then
        Exit Sub
    End If
End If
If mod1.DName = "李午阳" Then
    If Not ((Left(txtBh.Text, 1) = "H" Or Left(txtBh.Text, 1) = "9" Or Left(txtBh.Text, 1) = "B" Or Left(txtBh.Text, 1) = "8" Or Left(txtBh.Text, 1) = "1")) Then
        Exit Sub
    End If
End If

'''''If chkJYF.Value = 0 And NPF = 1 Then
'''''    MsgBox "此单未禁用，不能删除！"
'''''    Exit Sub
'''''End If
ii = MsgBox("是否删除此货品资料？", vbYesNo + vbQuestion, "Hello")
If ii = vbNo Then
    Exit Sub
End If
timZm = 12 '删除
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "新货品资料"
    mod1.cmd.Parameters("@NBLX") = "删除"
    mod1.cmd.Parameters("@bh") = txtBh.Text
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = 0
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = Null
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

Private Sub cmdDing_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next

If optT2.Value = True And Trim(txtQM.Text) = "" Then
    MsgBox ("请您一定要告诉拒绝我的理由!  :) ")
    Exit Sub
End If
If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If

timZm = 6 '签字
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "新货品资料"
    mod1.cmd.Parameters("@NBLX") = "签字"
    mod1.cmd.Parameters("@bh") = Trim(txtBh.Text)
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = LL
    mod1.cmd.Parameters("@mt2") = LLUid
    mod1.cmd.Parameters("@mt3") = txtPartName.Text
    mod1.cmd.Parameters("@mt4") = Left(txtBh.Text, 1)
    If txtJz.ToolTipText <> "" And txtJz.ToolTipText <> txtJz.Text Then
        mod1.cmd.Parameters("@mt5") = "适用机组：" & txtJz.ToolTipText '缓存记录
    End If
    If txtBz.ToolTipText <> "" And txtBz.ToolTipText <> txtBz.Text Then
        mod1.cmd.Parameters("@mt5") = mod1.cmd.Parameters("@mt5") & Chr(13) & Chr(10) & "备注" & txtBz.Text
    End If
    If Len(mod1.cmd.Parameters("@mt5").Value) > 50 Then
         mod1.cmd.Parameters("@mt5") = ""
    End If
    mod1.cmd.Parameters("@mlt1") = txtQM.Text & Chr(13) & Chr(10) & mod1.cmd.Parameters("@mt5")   '评审建议
    mod1.cmd.Parameters("@mm1") = Lc
    mod1.cmd.Parameters("@mm2") = Fwid
    If OptT1.Value = True Then
        mod1.cmd.Parameters("@mb1") = 1 '同意
    Else
        mod1.cmd.Parameters("@mb1") = 0 '拒绝
    End If
    mod1.cmd.Parameters("@md1") = Null
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
frmQm.Visible = False
End Sub

Private Sub cmdGB_Click()
frmBr.Visible = False
End Sub

Private Sub cmdJQ_Click()
Dim tt As String

If Jpid = 0 Then
    tt = "select top 50 bh,partname,'原厂编号:'+oname+' '+gg+' '+xn+' '+ff+'适用机组:'+jz,pid from nlpmxc where delf=1  order by pid desc"
Else
    tt = "select top 50 bh,partname,'原厂编号:'+oname+' '+gg+' '+xn+' '+ff+'适用机组:'+jz,pid from nlpmxc where delf=1 and pid<=" & Jpid & " order by pid desc"
End If
Call dtgLPBound(tt)
End Sub

Private Sub cmdMod_Click()
Dim Xg As Integer '修改 1为技术员本人，2 为经理驳回(经过经理审核过的单子，驳回后，能修改的内容有限）

'禁用的单子，如果在审核状态，业务员不能做修改和删除
If chkJYF.Value = 1 And Lc > 1 And mod1.DName <> "倪东海" And mod1.DName <> "邹晨" Then
    Exit Sub
End If
If mod1.DName = "货品录入员" Then
    If Not ((Left(txtBh.Text, 1) = "3" Or Left(txtBh.Text, 1) = "A")) Then
        Exit Sub
    End If
End If
If mod1.DName = "李午阳" Then
    If Not ((Left(txtBh.Text, 1) = "H" Or Left(txtBh.Text, 1) = "9" Or Left(txtBh.Text, 1) = "B" Or Left(txtBh.Text, 1) = "8" Or Left(txtBh.Text, 1) = "1" Or Left(txtBh.Text, 1) = "2" Or Left(txtBh.Text, 1) = "4" Or Left(txtBh.Text, 1) = "5" Or Left(txtBh.Text, 1) = "6" Or Left(txtBh.Text, 1) = "7")) Then
        Exit Sub
    End If
End If
frm1.Enabled = False
If mod1.DName = "货品录入员" Or mod1.DName = "李午阳" Or mod1.DName = "倪东海" Then
    frm1.Enabled = True
    cmdSave.Enabled = True
'''''''    If Left(txtBH.Text, 1) = 9 Or Left(txtBH.Text, 1) = 8 Or Left(txtBH.Text, 1) = 3 Or Left(txtBH.Text, 1) = "B" Or Left(txtBH.Text, 1) = "A" Then
'''''''
'''''''        If Lc = 1 Then  '内存暂存历史数据
'''''''            txtJz.ToolTipText = txtJz.Text
'''''''            txtJz.Locked = False
'''''''            txtPartName.Locked = False
'''''''            txtBm1.Locked = False
'''''''            txtBm2.Locked = False
'''''''            txtBm3.Locked = False
'''''''            txtL1.Locked = False
'''''''            txtL2.Locked = False
'''''''            txtL3.Locked = False
'''''''        End If
'''''''    Else
'''''''        txtJz.Locked = True
'''''''        txtJz.Locked = True
'''''''        txtOname.Locked = False
'''''''        txtXN.Locked = False
'''''''    End If
'''''''
'''''''    '只能添加，不能修改
'''''''    If txtEngName.Text = "" Or Lc = 1 And Lc Then txtEngName.Locked = False
'''''''    If txtGG.Text = "" Or Lc = 1 Then txtGG.Locked = False
'''''''    If txtFF.Text = "" Or Lc = 1 Then txtFF.Locked = False
'''''''
''''''''''''    If Lc = 1 Then txtPartName.Locked = False
''''''''''''    If Lc = 1 Then txtOname.Locked = False
'''''''    If NPF = 0 Then
'''''''        txtOname.Locked = False
'''''''        txtEngName.Locked = False
'''''''        txtYpb.Locked = False
'''''''        txtGG.Locked = False
'''''''        txtXN.Locked = False
'''''''        txtL1.Locked = False
'''''''        txtL2.Locked = False
'''''''        txtL3.Locked = False
'''''''        txtBm1.Locked = False
'''''''        txtBm2.Locked = False
'''''''        txtBm3.Locked = False
'''''''        If Left(txtBH.Text, 1) = "9" Or Left(txtBH.Text, 1) = "8" Or Left(txtBH.Text, 1) = "3" Or Left(txtBH.Text, 1) = "B" Then
'''''''
'''''''            If cmdSave.Enabled = False And Lc = 1 Then '内存暂存历史数据
'''''''                txtJz.ToolTipText = txtJz.Text
'''''''                txtJz.Locked = False
'''''''            End If
'''''''        Else
'''''''            txtJz.Locked = True
'''''''            txtJz.Locked = True
'''''''        End If
'''''''    Else '倪工审核过，有些内容就不能修改
'''''''        txtEngName.Locked = False
'''''''        txtYpb.Locked = False
'''''''        txtGG.Locked = False
'''''''        txtL1.Locked = False
'''''''        txtL2.Locked = False
'''''''        txtL3.Locked = False
'''''''        txtBm1.Locked = False
'''''''        txtBm2.Locked = False
'''''''        txtBm3.Locked = False
'''''''        '如果项目为空，则允许添加
'''''''        If Lc = 1 And txtPartName.Text = "" Then
'''''''            txtPartName.Locked = False
'''''''        End If
'''''''        If Lc = 1 And txtOname.Text = "" Then
'''''''            txtOname.Locked = False
'''''''        End If
'''''''        If Lc = 1 And txtXN.Text = "" Then
'''''''            txtXN.Locked = False
'''''''        End If
'''''''        If Left(txtBH.Text, 1) = "9" Or Left(txtBH.Text, 1) = "8" Or Left(txtBH.Text, 1) = "3" Or Left(txtBH.Text, 1) = "B" Then
'''''''
'''''''            If cmdSave.Enabled = False And Lc = 1 Then '内存暂存历史数据
'''''''                txtJz.ToolTipText = txtJz.Text
'''''''                txtJz.Locked = False
'''''''            End If
'''''''        Else
'''''''            txtJz.Locked = True
'''''''            txtJz.Locked = True
'''''''        End If
'''''''    End If
'''''''
    If mod1.DName <> "倪东海" Then '录入员一经修改，流程跳至录入员
        Lc = 1: LCRen = mod1.DName: LCUid = mod1.DHid
    End If
        chkJYF.Enabled = True
    dtgP.Row = 1: dtgP.Col = 1
    txtPartName.ToolTipText = txtPartName.Text
    If (dtgP.Text = mod1.DName Or dtgP.Text = "") And Lc = 1 Then
        Xg = 1
        If NPF = 0 Then
            txtOname.ToolTipText = txtOname.Text: txtOname.Locked = False
            txtPartName.ToolTipText = txtPartName.Text: txtPartName.Locked = False
            txtXN.ToolTipText = txtXN.Text: txtXN.Locked = False
            txtGG.ToolTipText = txtGG.Text: txtGG.Locked = False
            txtEngName.ToolTipText = txtEngName.Text: txtEngName.Locked = False
            txtYpb.ToolTipText = txtYpb.Text: txtYpb.Locked = False
            If Left(txtBh.Text, 1) = "9" Then
                txtJz.Locked = False
            End If
            If Left(txtBh.Text, 1) = "B" Or Left(txtBh.Text, 1) = "H" Then
                txtJz.Locked = False
                txtPb.Locked = False
            End If
        Else
            If txtOname.Text = "" Then
                txtOname.Locked = False
                txtOname.ToolTipText = txtOname.Text
            End If
            If txtPartName.Text = "" Then
                txtPartName.Locked = False
                txtPartName.ToolTipText = txtPartName.Text
            End If
            If txtXN.Text = "" Then
                txtXN.Locked = False
                txtXN.ToolTipText = txtXN.Text
            End If
        End If
        txtBm1.ToolTipText = txtBm1.Text: txtBm1.Locked = False
        txtBm2.ToolTipText = txtBm2.Text: txtBm2.Locked = False
        txtBm3.ToolTipText = txtBm3.Text: txtBm3.Locked = False
        txtL1.ToolTipText = txtL1.Text: txtL1.Locked = False
        txtL2.ToolTipText = txtL2.Text: txtL2.Locked = False
        txtL3.ToolTipText = txtL3.Text: txtL3.Locked = False
        txtBz.ToolTipText = txtBz.Text
        If txtYpb.Text = "" Then txtYpb.Locked = False
        If txtGG.Text = "" Then txtGG.Locked = False
        If txtPb.Text = "" Then
            If (Left(txtBh.Text, 1) = "9" Or Left(txtBh.Text, 1) = "8" Or Left(txtBh.Text, 1) = "3" Or Left(txtBh.Text, 1) = "B") Then
                txtPb.Locked = False
                cmdTD.Enabled = False
            Else
                cmdTD.Enabled = True
            End If
        End If
        If txtJz.Text = "" Then
            If (Left(txtBh.Text, 1) = "9" Or Left(txtBh.Text, 1) = "8" Or Left(txtBh.Text, 1) = "3" Or Left(txtBh.Text, 1) = "B") Then
                txtJz.Locked = False
                cmdTD.Enabled = False
            Else
                cmdTD.Enabled = True
            End If
        End If
    Else
        Xg = 2
        txtBm1.ToolTipText = txtBm1.Text: txtBm1.Locked = False
        txtBm2.ToolTipText = txtBm2.Text: txtBm2.Locked = False
        txtBm3.ToolTipText = txtBm3.Text: txtBm3.Locked = False
        txtL1.ToolTipText = txtL1.Text: txtL1.Locked = False
        txtL2.ToolTipText = txtL2.Text: txtL2.Locked = False
        txtL3.ToolTipText = txtL3.Text: txtL3.Locked = False
        txtBz.ToolTipText = txtBz.Text
        txtYpb.ToolTipText = txtYpb.Text: txtYpb.Locked = False
        txtGG.ToolTipText = txtGG.Text: txtGG.Locked = False
        txtPb.ToolTipText = txtPb.Text
            If (Left(txtBh.Text, 1) = "A" Or Left(txtBh.Text, 1) = "3" Or Left(txtBh.Text, 1) = "B") Then
                txtPb.Locked = False
                cmdTD.Enabled = False
            Else
                cmdTD.Enabled = True
            End If

            txtJz.ToolTipText = txtJz.Text
            If (Left(txtBh.Text, 1) = "A" Or Left(txtBh.Text, 1) = "3" Or Left(txtBh.Text, 1) = "B") Then
                txtJz.Locked = False
                cmdTD.Enabled = False
            Else
                cmdTD.Enabled = True
            End If
    End If
    If mod1.DName = "倪东海" Then '暂时给经理开通所有权限
        Call Me.CreateQuan(Left(txtBh.Text, 1))
        txtOname.Locked = False
        txtPartName.Locked = False
        txtXN.Locked = False
    End If
    txtBz.Locked = False
ElseIf mod1.DName = "" Or Ywy = "吴金荣" Then
    frm2.Enabled = True
    cmdSave.Enabled = True
    txtJJ1.Locked = False
    txtJJ2.Locked = False
    txtJJ3.Locked = False
    txtCb.Locked = False
ElseIf mod1.DName = "邹晨" Then
    cmdSave.Enabled = True
    frm3.Enabled = True
    'txtOname2.SetFocus
    txtOname2.SelStart = 0
    txtOname2.SelLength = Len(txtOname2.Text)
    Clipboard.Clear
    Clipboard.SetText txtOname2.SelText
ElseIf mod1.DName = "马晓聪" Then
    frm1.Enabled = True
    frm2.Enabled = True
    frm3.Enabled = True
    cmdSave.Enabled = True
End If

End Sub

Private Sub cmdNQ_Click()
Dim ii As Integer
Dim oo As Integer
On Error Resume Next
'禁用的单子，如果在审核状态，业务员不能做修改和删除
If chkJYF.Value = 1 And Lc > 1 And mod1.DName <> "倪东海" Then
    Exit Sub
End If
If mod1.DName = "货品录入员" Then
    If Not ((Left(txtBh.Text, 1) = "3" Or Left(txtBh.Text, 1) = "A")) Then
        Exit Sub
    End If
End If
If mod1.DName = "李午阳" Then
    If Not ((Left(txtBh.Text, 1) = "H" Or Left(txtBh.Text, 1) = "9" Or Left(txtBh.Text, 1) = "B" Or Left(txtBh.Text, 1) = "8" Or Left(txtBh.Text, 1) = "1")) Then
        Exit Sub
    End If
End If

'''''If lblTX.Caption = "审核完毕!" Then Exit Sub
If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If

If LCUid <> mod1.DHid And mod1.DName <> "倪东海" And LCRen = "导入" Then
        MsgBox "此处应由" & LCRen & "签字! 请您不要再点"
        Exit Sub
End If

frmQm.Visible = True
If Lc = 1 Then
    optT2.Enabled = False
    OptT1.Value = True
    
Else
    OptT1.Enabled = True
    optT2.Enabled = True
    OptT1.Value = False
    optT2.Value = False
End If
cmdDing.Enabled = True
End Sub

Private Sub cmdNQ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'禁用的单子，如果在审核状态，业务员不能做修改和删除
If chkJYF.Value = 1 And Lc > 1 And mod1.DName <> "倪东海" Then
    Exit Sub
End If
If Left(txtBh.Text, 1) = "3" Or Left(txtBh.Text, 1) = "A" Then
    If mod1.DName <> "货品录入员" Then
        Exit Sub
    End If
End If

If (Left(txtBh.Text, 1) = "H" Or Left(txtBh.Text, 1) = "B" Or Left(txtBh.Text, 1) = "9" Or Left(txtBh.Text, 1) = "8" Or Left(txtBh.Text, 1) = "1") And mod1.DName <> "李午阳" Then
    Exit Sub
End If


    frmQm.Visible = True
    OptT1.Enabled = False
    optT2.Value = True

End Sub


Private Sub cmdOK_Click()
Dim tt As String
Dim Rb

'先检测是否会重复编号
tt = "select count(bh) from nlpmxc where bh='" & txtBh.Text & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Rb = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
If Rb(0, 0) > 0 Then
    MsgBox ("有重复编号，请确认!")
    Exit Sub
End If
cmdOK.Visible = False
cmdCreate.Visible = True
txtLb1.Text = HY2
txtLb2.Text = HY3
If Bm1 = "9" Then
    txtPb.Text = HY2
    txtYpb.Text = HY2
    txtJz.Locked = False
Else
    txtJz.Locked = True
End If

Pid = 0

frm2.Enabled = False
frm3.Enabled = False
cmdSave.Enabled = True
txtPartName.Locked = False
txtOname.Locked = False
txtEngName.Locked = False
txtGG.Locked = False
txtXN.Locked = False
txtFF.Locked = False
txtYpb.Locked = False
txtBh.Locked = True
dtgL1.Visible = False
txtBm1.Locked = False
txtBm2.Locked = False
txtBm3.Locked = False
txtL1.Locked = False
txtL2.Locked = False
txtL3.Locked = False

'倪工修改
''''''frmName.Visible = True
txtPartName.Locked = False


txtName.Text = ""
If Left(txtBh, 1) = "9" Or Left(txtBh, 1) = "8" Or Left(txtBh, 1) = "3" Or Left(txtBh.Text, 1) = "B" Then
    txtJz.Locked = False
End If

Call CreateQuan(Left(txtBh.Text, 1)) '新建字段权限

End Sub

Private Sub cmdPic_Click()
'''''Dim d As String
'''''d = Dir("C:\Documents and Settings\Administrator\桌面\数据对应图片\*" & txtOname.Text & "*.jpg")
'''''Do Until d = ""
'''''   MsgBox "找到一个图片文件：" & d
'''''
'''''   d = Dir
'''''Loop
Dim picL As Integer
picL = 1
Dim d As String
d = Dir("C:\Documents and Settings\Administrator\桌面\数据对应图片\*instr(text1.txt,vbcrlf)" & txtOname.Text & ".jpg")
Do Until d = ""
   MsgBox "找到一个图片文件：" & d
   d = Dir
   picL = picL + 1
   If picL = 3 Then Exit Do
Loop

End Sub

Private Sub cmdSave_Click()
Dim tt As String
Dim oo As Integer
Dim ModfiTT As String
Dim ii As Integer
Dim Ra
Dim Rb
Dim LPB '替代品牌之和
Dim LJZ '替代机组之和
Dim La As Integer
If Len(txtBh.Text) <> 5 Then Exit Sub
If txtPartName.Text = "" Then Exit Sub
'修改记录
If txtPartName.Text <> txtPartName.ToolTipText Then ModfiTT = ModfiTT & "货品名称：" & txtPartName.ToolTipText & " "
If txtOname.Text <> txtOname.ToolTipText Then ModfiTT = ModfiTT & "原厂品牌：" & txtOname.ToolTipText & " "
If txtXN.Text <> txtXN.ToolTipText Then ModfiTT = ModfiTT & "产品型号：" & txtXN.ToolTipText & " "
If txtBm1.Text <> txtBm1.ToolTipText Then ModfiTT = ModfiTT & "别名1：" & txtBm1.ToolTipText & " "
If txtBm2.Text <> txtBm2.ToolTipText Then ModfiTT = ModfiTT & "别名2：" & txtBm2.ToolTipText & " "
If txtBm3.Text <> txtBm3.ToolTipText Then ModfiTT = ModfiTT & "别名3：" & txtBm3.ToolTipText & " "
If txtEngName.Text <> txtEngName.ToolTipText Then ModfiTT = ModfiTT & "英文名：" & txtEngName.ToolTipText & " "
If txtL1.Text <> txtL1.ToolTipText Then ModfiTT = ModfiTT & "类别1：" & txtL1.ToolTipText & " "
If txtL2.Text <> txtL2.ToolTipText Then ModfiTT = ModfiTT & "类别2：" & txtL2.ToolTipText & " "
If txtL3.Text <> txtL3.ToolTipText Then ModfiTT = ModfiTT & "类别3：" & txtL3.ToolTipText & " "
If txtBz.Text <> txtBz.ToolTipText Then ModfiTT = ModfiTT & "备注：" & txtBz.ToolTipText & " "
If txtYpb.Text <> txtYpb.ToolTipText Then ModfiTT = ModfiTT & "原厂品牌：" & txtYpb.ToolTipText & " "
If txtGG.Text <> txtGG.ToolTipText Then ModfiTT = ModfiTT & "包装规格：" & txtGG.ToolTipText & " "
If txtPb.Text <> txtPb.ToolTipText Then ModfiTT = ModfiTT & "适用品牌：" & txtPb.ToolTipText & " "
If txtJz.Text <> txtJz.ToolTipText Then ModfiTT = ModfiTT & "适用机组：" & txtJz.ToolTipText & " "
If frm1.Enabled = True Then
'倪工修改
'''''''    If Pid = 0 And txtOname.Text <> "" Then  '检测原厂编号是否有重复
'''''''        tt = "select pid from nlpmxc where oname='" & txtOname.Text & "' and jyf=1 and delf=1"
'''''''        Set mod1.HTP = CreateObject("adodb.recordset")
'''''''        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
'''''''        If mod1.HTP.BOF = False Then
'''''''            Ra = mod1.HTP.GetRows
'''''''            mod1.HTP.Close
'''''''            Set mod1.HTP = Nothing
'''''''            ii = MsgBox("检测到有相同原厂编号!,不能添加新货品,是否返回到以前存在的货品?", vbYesNo + vbQuestion, mod1.chenHu)
'''''''            If ii = vbYes Then
'''''''                Call Bound(Val(Ra(0, 0)))
'''''''            End If
'''''''            Exit Sub
'''''''        Else
'''''''            Set mod1.HTP = Nothing
'''''''        End If
'''''''    End If
Set Ra = Nothing

    timZm = 1 '保存
        '保存前检测是否有重复编号
        tt = "select pid from nlpmxc where bh='" & txtBh.Text & "'"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        If mod1.HTP.BOF = False Then
            Ra = mod1.HTP.GetRows
            Pid = Ra(0, 0)
            mod1.HTP.Close
        End If
        
        '1,2,4,5,6,7开头的货品，适用品牌＝原厂品牌＋所替代的品牌之和
        If Left(txtBh.Text, 1) = "1" Or Left(txtBh.Text, 1) = "2" Or Left(txtBh.Text, 1) = "4" Or Left(txtBh.Text, 1) = "5" Or _
         Left(txtBh.Text, 1) = "6" Or Left(txtBh.Text, 1) = "7" Then
            tt = "select ypb,jz from nlptdpb where tdbh='" & txtBh.Text & "'"
            Set mod1.HTP = CreateObject("adodb.recordset")
            mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
            If mod1.HTP.BOF = True Then
                Set mod1.HTP = Nothing
                LPB = ""
            Else
                Ra = mod1.HTP.GetRows
                mod1.HTP.Close
                Set mod1.HTP = Nothing
                La = UBound(Ra, 2) + 1
                For oo = 0 To La - 1
                    If InStr(1, LPB, Ra(0, oo)) > 0 Then
                    Else
                        LPB = LPB & " " & Ra(0, oo)
                    End If
                    If InStr(1, LJZ, Ra(1, oo)) > 0 Then
                    Else
                        LJZ = LJZ & " " & Ra(1, oo)
                    End If
                Next
            End If
            
            txtPb.Text = txtYpb.Text & LPB
            txtJz.Text = Trim(LJZ)
        End If
        
        Set mod1.HTP = Nothing
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "新货品资料"
        mod1.cmd.Parameters("@NBLX") = "保存"
        mod1.cmd.Parameters("@bh") = Str(Pid)
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = Replace(txtBh.Text, vbCrLf, "", 1)
        mod1.cmd.Parameters("@mt2") = Replace(txtPartName.Text, vbCrLf, "", 1)
        mod1.cmd.Parameters("@mt3") = txtEngName.Text
        mod1.cmd.Parameters("@mt4") = txtLb1.Text
        mod1.cmd.Parameters("@mt5") = txtLb2.Text
        mod1.cmd.Parameters("@mt6") = txtGG.Text
        mod1.cmd.Parameters("@mt7") = txtXN.Text '产品型号
        mod1.cmd.Parameters("@mt8") = txtFF.Text
        mod1.cmd.Parameters("@mt9") = txtPb.Text
        'If Left(txtBh.Text, 1) = 9 Or Left(txtBh.Text, 1) = 8 Or Left(txtBh.Text, 1) = 3 Then
        If Left(txtBh.Text, 1) = 9 Or Left(txtBh.Text, 1) = 8 Then
            mod1.cmd.Parameters("@mt9") = txtYpb.Text
        End If
        mod1.cmd.Parameters("@mt10") = txtJz.Text
        mod1.cmd.Parameters("@mt11") = Replace(txtOname.Text, vbCrLf, "", 1)
        mod1.cmd.Parameters("@mt12") = txtYpb.Text
        mod1.cmd.Parameters("@mt13") = txtBm1.Text
        mod1.cmd.Parameters("@mt14") = txtBm2.Text
        mod1.cmd.Parameters("@mt15") = txtBm3.Text
        mod1.cmd.Parameters("@mt16") = txtL1.Text
        mod1.cmd.Parameters("@mt17") = txtL2.Text
        mod1.cmd.Parameters("@mt18") = txtL3.Text
        mod1.cmd.Parameters("@mt19") = txtBh.ToolTipText
        
        mod1.cmd.Parameters("@mlt1") = txtBz.Text '备注
        mod1.cmd.Parameters("@mlt2") = ModfiTT '修改
        mod1.cmd.Parameters("@mm1") = 0
        If chkJYF.Value = 1 Then
            mod1.cmd.Parameters("@mb1") = 0
        Else
            mod1.cmd.Parameters("@mb1") = 1
        End If
        mod1.cmd.Parameters("@md1") = Null
        Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
        mod1.cmd.Execute
        mod1.Zid = mod1.cmd.Parameters("@zid").Value
        If mod1.cmd.Parameters("@errch").Value <> "成功" Then
            MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
                cmdSave.Enabled = False
            Exit Sub
        Else '提交成功,等待系统中心处理数据
            Me.Enabled = False
            frmWaitA.Visible = True
            frmWaitA.Timer2.Enabled = False
    
            frmWaitA.ZOrder 0
            frmWaitA.Timer2.Enabled = True
            timWait.Enabled = True
            cmdSave.Enabled = False
        End If
    Set mod1.cmd = Nothing
ElseIf txtJJ1.Locked = False Then
    timZm = 2 '保存
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "新货品资料"
        mod1.cmd.Parameters("@NBLX") = "保存2"
        mod1.cmd.Parameters("@bh") = Str(Pid)
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = txtBh.Text
        mod1.cmd.Parameters("@mlt1") = ""
        mod1.cmd.Parameters("@mm1") = Val(txtGy1.ToolTipText)
        mod1.cmd.Parameters("@mm2") = Val(txtGy2.ToolTipText)
        mod1.cmd.Parameters("@mm3") = Val(txtGY3.ToolTipText)
        mod1.cmd.Parameters("@mm4") = Val(txtJJ1.Text)
        mod1.cmd.Parameters("@mm5") = Val(txtJJ2.Text)
        mod1.cmd.Parameters("@mm6") = Val(txtJJ3.Text)
        mod1.cmd.Parameters("@mb1") = Null
        mod1.cmd.Parameters("@md1") = Null
        Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
        mod1.cmd.Execute
        mod1.Zid = mod1.cmd.Parameters("@zid").Value
        If mod1.cmd.Parameters("@errch").Value <> "成功" Then
            MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
                cmdSave.Enabled = False
            Exit Sub
        Else '提交成功,等待系统中心处理数据
            Me.Enabled = False
            frmWaitA.Visible = True
            frmWaitA.Timer2.Enabled = False
    
            frmWaitA.ZOrder 0
            frmWaitA.Timer2.Enabled = True
            timWait.Enabled = True
            cmdSave.Enabled = False
        End If
    Set mod1.cmd = Nothing
ElseIf frm3.Enabled = True Then
    timZm = 3 '保存
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "新货品资料"
        mod1.cmd.Parameters("@NBLX") = "保存3"
        mod1.cmd.Parameters("@bh") = Str(Pid)
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = txtBh.Text
        mod1.cmd.Parameters("@mlt1") = ""
        mod1.cmd.Parameters("@mm1") = Val(txtMj.Text)
        mod1.cmd.Parameters("@mm2") = Val(txtListPrice.Text)
        mod1.cmd.Parameters("@mm3") = Val(txtDj.Text)
        Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
        mod1.cmd.Execute
        mod1.Zid = mod1.cmd.Parameters("@zid").Value
        If mod1.cmd.Parameters("@errch").Value <> "成功" Then
            MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
                cmdSave.Enabled = False
            Exit Sub
        Else '提交成功,等待系统中心处理数据
            Me.Enabled = False
            frmWaitA.Visible = True
            frmWaitA.Timer2.Enabled = False
    
            frmWaitA.ZOrder 0
            frmWaitA.Timer2.Enabled = True
            timWait.Enabled = True
            cmdSave.Enabled = False
        End If
    Set mod1.cmd = Nothing
End If
End Sub

Private Sub cmdSC_Click()
Dim tt As String
Dim JT As String
JT = ",oname,gg,xn,pb,jz,ypb,bm1,bm2,bm3,l1,l2,l3"

    tt = "select bh,partname,pid" & JT & " from nlpmxc where delf=1 and (left(bh,1)='9' or left(bh,1)='8') and (listprice is null or listprice=0 or listprice=1000) and oname<>'' order by bh desc"
Call frmHPBR.dtgLPBound(tt)

frmHPBR.Show
frmHPBR.ZOrder 0
End Sub

Private Sub cmdT_Click()
If cmdT.Caption = "替代" Then
    Call TD(txtBh.Text)
    cmdT.Caption = "被替代"
Else
    Call TDB(txtBh.Text)
    cmdT.Caption = "替代"
End If
End Sub

Private Sub cmdTD_Click()
Dim tt As String
Dim ii As Integer
Dim Ra
If txtTdbh.Visible = False Then
    txtTdbh.Visible = True
    Exit Sub
End If
If Len(txtBh.Text) <> 5 Then Exit Sub
If txtTdbh.Visible = True And txtTdbh.Text <> "" Then
    '先查有无重复
    tt = "select tdbh from nlptd where bh='" & txtBh.Text & "' and tdbh='" & txtTdbh.Text & "'"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    If mod1.HTP.BOF = False Then
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        MsgBox "有重复"
        Exit Sub
    End If
    Set mod1.HTP = Nothing
    
    timZm = 5 '替代
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "新货品资料"
        mod1.cmd.Parameters("@NBLX") = "替代"
        mod1.cmd.Parameters("@bh") = txtBh.Text
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = txtTdbh.Text
        mod1.cmd.Parameters("@mlt1") = ""
        mod1.cmd.Parameters("@mm1") = 0
        mod1.cmd.Parameters("@mb1") = Null
        mod1.cmd.Parameters("@md1") = Null
        Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
        mod1.cmd.Execute
        mod1.Zid = mod1.cmd.Parameters("@zid").Value
        If mod1.cmd.Parameters("@errch").Value <> "成功" Then
            MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
            If timZm = 2 Then '保存
                cmdSave.Enabled = False
            End If
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
    Exit Sub
End If
If txtTdbh.Visible = True And txtTdbh.Text = "" Then
    txtTdbh.Visible = False
    Exit Sub
End If
End Sub

Private Sub cmdXT_Click()
Dim tt As String
Select Case comMLx.Text
Case "货品名称"
    tt = "select bh,partname,'原厂编号:'+oname+' '+gg+' '+xn+' '+ff+' 适用机组:'+jz,pid from nlpmxc" & _
    " where partname in (select partname from nlpmxc group by partname having(count(*))>1)"
    
Case "适用机组"
    tt = "select bh,partname,'原厂编号:'+oname+' '+gg+' '+xn+' '+ff+' 适用机组:'+jz,pid from nlpmxc" & _
    " where jz in (select jz from nlpmxc group by jz having(count(*))>1)"
Case ""
    tt = "select bh,partname,'原厂编号:'+oname+' '+gg+' '+xn+' '+ff+' 适用机组:'+jz,pid from nlpmxc order by pb desc,partname,jz"
End Select
Call dtgLPBound(tt)
End Sub

Private Sub cmdZ_Click()
Dim tt As String
Dim Tbh As String
Dim Ra
Dim Rb
Dim LT As String
Dim ii As Integer
'如果已经有更新,则不能再添加
tt = "select bh from nlpmxc where gxbh='" & txtBh.Text & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
If Not (mod1.HTP.EOF) = True Then
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing

        ii = MsgBox("此号已经被更新过,更新号为" & Ra(0, 0), vbOKOnly, "不行啦")
        Exit Sub
End If
    mod1.HTP.Close
    Set mod1.HTP = Nothing

If Not (mod1.DName = "货品录入员" Or mod1.DName = "李午阳" Or mod1.DName = "马晓聪") Then Exit Sub
Tbh = txtBh.Text
Call Qing
cmdZ.Visible = True
txtBh.ToolTipText = Tbh

dtgL1.Visible = True
    MsgBox "请正确选择货品分类!"
    Exit Sub
    

End Sub

Private Sub dtgBr_DblClick()
On Error Resume Next
If dtgBr.Row = 0 Then Exit Sub
If GyId = 0 Then GyId = 1
If GyId = 1 Then
    dtgBr.Col = 0: txtGy1.Text = dtgBr.Text
    dtgBr.Col = 1: txtGy1.ToolTipText = dtgBr.Text
ElseIf GyId = 2 Then
    dtgBr.Col = 0: txtGy2.Text = dtgBr.Text
    dtgBr.Col = 1: txtGy2.ToolTipText = dtgBr.Text
ElseIf GyId = 3 Then
    dtgBr.Col = 0: txtGY3.Text = dtgBr.Text
    dtgBr.Col = 1: txtGY3.ToolTipText = dtgBr.Text
End If
End Sub


Private Sub dtgL1_Click()
dtgN1.Row = dtgL1.Row
dtgN1.Col = 2
'''''''If dtgN1.Text = "" Then
'''''''    Call Me.dtgL2FF
'''''''    Call Me.dtgL3FF
'''''''    txtBH.Text = ""
'''''''    Exit Sub
'''''''End If
'''''''
'''''''B1id = Val(dtgN1.Text)
'''''''Call BoundL2
'''''''dtgN1.Col = 0
'''''''Bm1 = Trim(dtgN1.Text)
'''''''Call Qing
'''''''txtBH.Text = Bm1
'''''''Call dtgL3FF
'''''''dtgN1.Col = 1
'''''''HY1 = dtgN1.Text
'''''''
'''''''Dim tt As String
'''''''Dim Ra
'''''''Dim Rb
'''''''Dim LT As String
'''''''
'''''''If Len(txtBH.Text) <> 1 Then
'''''''    MsgBox "请正确选择货品分类!"
'''''''    Exit Sub
'''''''End If
'''''''
'''''''Call Qing
'''''''Bm = Bm1
'''''''On Error Resume Next
'''''''tt = "select max(bh) from nlpmxc where left(bh,1)='" & Bm & "'"
'''''''Set mod1.HTP = CreateObject("adodb.recordset")
'''''''mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
'''''''Rb = mod1.HTP.GetRows
'''''''mod1.HTP.Close
'''''''Set mod1.HTP = Nothing
'''''''Bm = Rb(0, 0)
'''''''txtBH.Text = Bm
dtgN1.Col = 0
If (dtgN1.Text = "" Or txtBh.Text <> "") And cmdZ.Visible = False Then Exit Sub
If (Trim(dtgN1.Text) = "3" Or Trim(dtgN1.Text) = "A") Then
    If mod1.DName <> "货品录入员" Then
        Exit Sub
    End If
End If
If (Trim(dtgN1.Text) = "H" Or Trim(dtgN1.Text) = "B" Or Trim(dtgN1.Text) = "9" Or Trim(dtgN1.Text) = "8" Or Trim(dtgN1.Text) = "1") And mod1.DName <> "李午阳" Then
    Exit Sub
End If
If cmdZ.Visible = True And Trim(dtgN1.Text) = "3" Then
    Exit Sub
End If
Call NewBh(dtgN1.Text)
cmdCreate.Visible = False

    cmdOK.Visible = True

'txtBH.Locked = False
frm1.Enabled = True
End Sub


Private Sub dtgL1_DblClick()
frmBr.Visible = True
dtgLP.Visible = True
End Sub

Private Sub dtgL2_Click()
dtgN2.Row = dtgL2.Row
dtgN2.Col = 2
If dtgN2.Text = "" Then
    Call Me.dtgL3FF
    txtBh.Text = ""
    Exit Sub
End If
B2id = Val(dtgN2.Text)
Call BoundL3
dtgN2.Col = 0
Bm2 = Trim(dtgN2.Text)
Call Qing
txtBh.Text = Bm1 + Bm2
dtgN2.Col = 1
HY2 = dtgN2.Text
End Sub


Private Sub dtgL3_Click()
dtgN3.Row = dtgL3.Row
dtgN3.Col = 2
If dtgN3.Text = "" Then Exit Sub
B3id = Val(dtgN3.Text)
'Call BoundL3
dtgN3.Col = 0
Bm3 = Trim(dtgN3.Text)
Call Qing
txtBh.Text = Bm1 + Bm2 + Bm3
dtgN3.Col = 1
HY3 = dtgN3.Text

End Sub


Private Sub dtgL3_DblClick()
Dim tt As String
If B3id > 0 Then

    Bm = Bm1 & Bm2 & Bm3
    tt = "select bh,partname,gg+' '+xn+' '+ff,pid from nlpmxc where left(bh,3)='" & Bm & "' and delf=1"
    Call dtgLPBound(tt)
Else
    Call dtgLPFF
End If
frmBr.Visible = True
dtgLP.Visible = True
End Sub


Private Sub dtgLP_Click()
Dim Bh As String
On Error Resume Next
dtgLN.Row = dtgLP.Row
dtgLN.Col = 3
Pid = Val(dtgLN.Text)
If Pid = 0 Then Exit Sub
If txtTdbh.Visible = False Then
    Call Qing
    Call Bound(Pid)
Else
    dtgLN.Col = 0
    Bh = dtgLN.Text
    If Bh <> txtBh.Text Then
        txtTdbh.Text = Bh
    End If
End If
End Sub

Private Sub dtgName_DblClick()
If dtgName.Row <> 0 Then
    dtgNN.Row = dtgName.Row
    dtgNN.Col = 1
    txtPartName.Text = dtgNN.Text
    frmName.Visible = False
End If
End Sub


Private Sub dtgPic_Click()
If dtgPic.Row = 0 Then Exit Sub
dtgPic.Col = 1
If dtgPic.Text = "" Then Exit Sub
Dim icmd As Object
Set icmd = CreateObject("wscript.shell")
icmd.Run "command.com /c rename C:\Documents and Settings\Administrator\桌面\数据对应图片\" & txtOname.Text & ".jpg" & txtBh.Text & ".jpg", 0, True
dtgPic.Visible = False
End Sub

Private Sub Form_Click()
'frmBr.Visible = False
Me.txtTdbh.Visible = False
dtgL1.Visible = False
frmHPBR.Visible = False
frmQm.Visible = False
End Sub

Public Sub dtgPFF()
Dim oo As Integer
For oo = 1 To dtgP.Rows - 1
    dtgP.RowHeight(oo) = dtgP.RowHeight(0) * 2
Next
dtgP.Clear
dtgP.Row = 0
dtgP.Col = 0: dtgP.Text = "日期": dtgP.Col = 1: dtgP.Text = "姓名": dtgP.Col = 2: dtgP.Text = "职能": dtgP.Col = 3: dtgP.Text = "评审建议": dtgP.Col = 4: dtgP.Text = "审核":
dtgP.ColWidth(0) = 1665
dtgP.ColWidth(1) = 1005
dtgP.ColWidth(2) = 0
 dtgP.ColWidth(3) = 4290: dtgP.ColWidth(4) = 1035
For oo = 0 To 4
    dtgP.Col = oo
    dtgP.CellFontBold = True
Next
End Sub
Public Sub QMBound(Rz, Lz As Integer)
Dim ii As Integer: Dim oo As Integer
On Error Resume Next
Call dtgPFF
dtgP.Rows = Lz + 20

For oo = 1 To Lz + 1
    dtgP.Row = oo
    For ii = 0 To 5
        dtgP.Col = ii
        dtgP.Text = Rz(ii, oo - 1)
        If ii = 3 Then
            If Len(Rz(ii, oo - 1)) > 16 Then
                dtgP.RowHeight(oo) = UpInt(Len(Rz(ii, oo - 1)) / 16) * dtgP.RowHeight(oo)
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
For oo = 1 To Lz + 1
    dtgP.Row = oo
    dtgP.Col = 4
            If dtgP.Text = "驳回" Then
                For ii = 0 To 5
                    dtgP.Col = ii
                    dtgP.CellForeColor = &HFF&
                Next
            End If
Next
End Sub
Private Sub Form_DblClick()
If mod1.DName = "宋晓炯" Or mod1.DName = "沈维" Or mod1.DName = "马晓聪" Then
    frm1.Visible = False: frm3.Visible = False
    frm2.Visible = False
    If frId = 1 Then
        frm1.Visible = True
        frId = 2
    ElseIf frId = 2 Then
        frm2.Visible = True
        frId = 3
    ElseIf frId = 3 Then
        frm3.Visible = True
        frId = 1
    End If
ElseIf mod1.DName = "倪东海" Then
'''    If frmBm.Visible = True Then
'''        frmBm.Visible = False
'''    Else
'''        frmBm.Visible = True
'''    End If

End If

End Sub


Private Sub Form_Load()
Dim Ra, Rb
Dim La As Long
Dim Lb As Integer
Dim oo As Long
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0
frmBr.Top = 0
frmBr.Left = 30
dtgLP.Top = 0
dtgLP.Left = 0
cmdOK.Left = cmdCreate.Left
cmdOK.Top = cmdCreate.Top

frId = 2
If mod1.DName <> "倪东海" And mod1.DName <> "马晓聪" Then
    frmXT.Visible = False
    frmBm.Visible = False
Else
    frmXT.Visible = True
'''''    frmBm.Visible = True
End If

If mod1.DName = "邹晨" Then
    cmdSC.Visible = True
Else
    cmdSC.Visible = False
End If

dtgName.Row = 0: dtgName.Col = 1
dtgName.Text = "双击列表选择货品名称": dtgName.CellFontBold = True
dtgName.Cols = 2
dtgName.ColWidth(0) = 0
dtgName.ColWidth(1) = 2000
dtgNN.Cols = 2
dtgNN.ColWidth(0) = 0
dtgNN.ColWidth(1) = 2000
tt = "select lid,partname,lb1,lb2,lb3,bm1,bm2,bm3 from nlplb order by lid;" & _
    "select Lb from nlpLBN"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
Lb = UBound(Rb, 2) + 1
dtgName.Rows = La + 10
dtgNN.Rows = La + 10
On Error Resume Next
For oo = 1 To La
    dtgName.Row = oo
    dtgName.Col = 0: dtgName.Text = Ra(0, oo - 1)
    dtgName.Col = 1: dtgName.Text = Ra(1, oo - 1)
'''    dtgName.Col = 2: dtgName.Text = Ra(2, oo - 1)
'''    dtgName.Col = 3: dtgName.Text = Ra(3, oo - 1)
'''    dtgName.Col = 4: dtgName.Text = Ra(4, oo - 1)
'''    dtgName.Col = 5: dtgName.Text = Ra(5, oo - 1)
'''    dtgName.Col = 6: dtgName.Text = Ra(6, oo - 1)
'''    dtgName.Col = 7: dtgName.Text = Ra(7, oo - 1)
    dtgNN.Row = oo
    dtgNN.Col = 0: dtgNN.Text = Ra(0, oo - 1)
    dtgNN.Col = 1: dtgNN.Text = Ra(1, oo - 1)
'''    dtgNN.Col = 2: dtgNN.Text = Ra(2, oo - 1)
'''    dtgNN.Col = 3: dtgNN.Text = Ra(3, oo - 1)
'''    dtgNN.Col = 4: dtgNN.Text = Ra(4, oo - 1)
'''    dtgNN.Col = 5: dtgNN.Text = Ra(5, oo - 1)
'''    dtgNN.Col = 6: dtgNN.Text = Ra(6, oo - 1)
'''    dtgNN.Col = 7: dtgNN.Text = Ra(7, oo - 1)
Next
On Error Resume Next
For oo = 50 To 0 Step -1
    txtL1.RemoveItem oo
    txtL2.RemoveItem oo
    txtL3.RemoveItem oo
Next
For oo = 0 To Lb - 1
    txtL1.AddItem Rb(0, oo)
    txtL2.AddItem Rb(0, oo)
    txtL3.AddItem Rb(0, oo)
Next
End Sub

Public Sub dtgL1FF()
dtgL1.Clear
dtgL1.Cols = 3
dtgL1.Rows = 30
dtgL1.Row = 0:
dtgL1.Col = 0: dtgL1.Text = "编码": dtgL1.CellFontBold = True
dtgL1.Col = 1: dtgL1.Text = "一级编码含义": dtgL1.CellFontBold = True
dtgL1.ColWidth(2) = 0
dtgL1.ColWidth(0) = 500
dtgL1.ColWidth(1) = 1590

dtgN1.Clear
dtgN1.Cols = 3
dtgN1.Rows = 30

End Sub


Public Sub BoundL1()
Dim tt As String
Dim oo As Long
Dim Ra
Dim La As Long
Call dtgL1FF
tt = "select bm,hy,b1id from L1 order by bm desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
For oo = 1 To La
    dtgL1.Row = oo
    dtgL1.Col = 0: dtgL1.Text = Ra(0, oo - 1)
    dtgL1.Col = 1: dtgL1.Text = Ra(1, oo - 1)
    dtgL1.Col = 2: dtgL1.Text = Ra(2, oo - 1)
    dtgN1.Row = oo
    dtgN1.Col = 0: dtgN1.Text = Ra(0, oo - 1)
    dtgN1.Col = 1: dtgN1.Text = Ra(1, oo - 1)
    dtgN1.Col = 2: dtgN1.Text = Ra(2, oo - 1)
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
    frmZu.Enabled = True
    Cancel = True
End Sub

Public Sub BoundL2()
Dim tt As String
Dim oo As Long
Dim Ra
Dim La As Long
Call dtgL2FF
tt = "select bm,hy,b2id from L2 where b1id=" & B1id & " order by bm desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
For oo = 1 To La
    dtgL2.Row = oo
    dtgL2.Col = 0: dtgL2.Text = Ra(0, oo - 1)
    dtgL2.Col = 1: dtgL2.Text = Ra(1, oo - 1)
    dtgL2.Col = 2: dtgL2.Text = Ra(2, oo - 1)
    dtgN2.Row = oo
    dtgN2.Col = 0: dtgN2.Text = Ra(0, oo - 1)
    dtgN2.Col = 1: dtgN2.Text = Ra(1, oo - 1)
    dtgN2.Col = 2: dtgN2.Text = Ra(2, oo - 1)
Next
End Sub

Public Sub BoundL3()
Dim tt As String
Dim oo As Long
Dim Ra
Dim La As Long
Call dtgL3FF
tt = "select bm,hy,b3id from L3 where b2id=" & B2id & " order by bm desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
For oo = 1 To La
    dtgL3.Row = oo
    dtgL3.Col = 0: dtgL3.Text = Ra(0, oo - 1)
    dtgL3.Col = 1: dtgL3.Text = Ra(1, oo - 1)
    dtgL3.Col = 2: dtgL3.Text = Ra(2, oo - 1)
    dtgN3.Row = oo
    dtgN3.Col = 0: dtgN3.Text = Ra(0, oo - 1)
    dtgN3.Col = 1: dtgN3.Text = Ra(1, oo - 1)
    dtgN3.Col = 2: dtgN3.Text = Ra(2, oo - 1)
Next
End Sub
Public Sub dtgL2FF()
dtgL2.Clear
dtgL2.Cols = 3
dtgL2.Rows = 30
dtgL2.Row = 0:
dtgL2.Col = 0: dtgL2.Text = "编码": dtgL2.CellFontBold = True
dtgL2.Col = 1: dtgL2.Text = "二级编码含义": dtgL2.CellFontBold = True
dtgL2.ColWidth(2) = 0
dtgL2.ColWidth(0) = 500
dtgL2.ColWidth(1) = 1590

dtgN2.Clear
dtgN2.Cols = 3
dtgN2.Rows = 30

End Sub

Public Sub dtgL3FF()
dtgL3.Clear
dtgL3.Cols = 3
dtgL3.Rows = 80
dtgL3.Row = 0:
dtgL3.Col = 0: dtgL3.Text = "编码": dtgL3.CellFontBold = True
dtgL3.Col = 1: dtgL3.Text = "三级编码含义": dtgL3.CellFontBold = True
dtgL3.ColWidth(2) = 0
dtgL3.ColWidth(0) = 500
dtgL3.ColWidth(1) = 1590

dtgN3.Clear
dtgN3.Cols = 3
dtgN3.Rows = 80

End Sub

Public Sub Qing()
    frm2.Visible = False
txtBh.Text = ""
txtBh.ToolTipText = ""
TmpBh = ""
Pid = 0
txtBh1.Text = "": txtPartName1.Text = ""
txtBh2.Text = "": txtPartName2.Text = ""
txtPartName.Text = "": txtPartName.ToolTipText = ""
txtEngName.Text = "": txtEngName.ToolTipText = ""
txtLb1.Text = "": txtLb1.ToolTipText = ""
txtLb2.Text = "": txtLb2.ToolTipText = ""
txtGG.Text = "": txtGG.ToolTipText = ""
txtXN.Text = "": txtXN.ToolTipText = ""
txtFF.Text = "": txtFF.ToolTipText = ""
txtPb.Text = "": txtPb1.Text = "": txtPb.ToolTipText = ""
txtJz.Text = "": txtJz.ToolTipText = ""
txtBz.Text = "": txtBz.ToolTipText = ""
txtOname.Text = "": txtOname.ToolTipText = ""
txtOname2.Text = ""
txtGy1.Text = "": txtGy1.ToolTipText = ""
txtGy2.Text = "": txtGy2.ToolTipText = ""
txtGY3.Text = "": txtGY3.ToolTipText = ""
txtJJ1.Text = ""
txtJJ2.Text = ""
txtJJ3.Text = ""
txtJJ1.Locked = True
txtJJ2.Locked = True
txtJJ3.Locked = True
txtCb.Text = ""
txtMj.Text = ""
txtListPrice.Text = ""
frm1.Enabled = False
txtCb.Locked = True
txtCb.Text = ""
'frm2.Enabled = False
frm3.Enabled = False
cmdSave.Enabled = False
txtTdbh.Text = ""
txtTdbh.Visible = False
cmdTD.Visible = False
txtYpb.Text = "": txtYpb.ToolTipText = ""
txtBm1.Text = "": txtBm1.ToolTipText = ""
txtBm2.Text = "": txtBm2.ToolTipText = ""
txtBm3.Text = "": txtBm3.ToolTipText = ""
txtL1.Text = "": txtL1.ToolTipText = ""
txtL2.Text = "": txtL2.ToolTipText = ""
txtL3.Text = "": txtL3.ToolTipText = ""
txtDj.Text = ""
Call Me.dtgbrFF
dtgL1.Visible = False
If mod1.DName = "货品录入员" Or mod1.DName = "待定" Then
frm1.Visible = False: frm2.Visible = False: frm3.Visible = False
    cmdTD.Visible = True
    frm1.Visible = True
ElseIf mod1.Bm = "配送中心" And mod1.DName <> "邹晨" Or mod1.Bm = "供应链" Then
frm1.Visible = False: frm2.Visible = False: frm3.Visible = False
    frm2.Visible = True
ElseIf mod1.DName = "邹晨" Then
frm1.Visible = False: frm2.Visible = False: frm3.Visible = True
    frm3.Visible = True
ElseIf mod1.DName = "马晓聪" Or mod1.DName = "宋晓炯" Or mod1.DName = "沈维" Then
'''''    frm1.Visible = True
End If


frmQm.Visible = False

LL = ""
LLUid = ""
LCRen = ""
LCUid = ""
Lc = 1
Fwid = 0
frmQm.Visible = False
txtQM.Text = ""
OptT1.Value = False
optT2.Value = False
lblTX.Caption = ""
txtTD.Text = "" '替代编号
txtTD.Locked = True
Call dtgPFF

    txtBh.Locked = True
    txtPartName.Locked = True
    txtOname.Locked = True
    txtEngName.Locked = True
    txtGG.Locked = True
    txtXN.Locked = True
    txtFF.Locked = True
    txtPb.Locked = True
    txtJz.Locked = True
    txtYpb.Locked = True
    chkJYF.Enabled = False
    txtL1.Locked = True
    txtL2.Locked = True
    txtL3.Locked = True
    txtBm1.Locked = True
    txtBm2.Locked = True
    txtBm3.Locked = True
    txtJz.ToolTipText = ""
    txtBz.ToolTipText = ""
    frmName.Visible = False

    NPF = 0 '倪工审核否
    txtBz.Locked = True
    
    cmdZ.Visible = False
    
End Sub

Private Sub frm1_DblClick()
If mod1.DName = "宋晓炯" Or mod1.DName = "沈维" Or mod1.DName = "马晓聪" Then
    frm2.Visible = True
End If
End Sub

Private Sub timQuit_Timer()
Dim Rz
Dim Lz As Integer
Dim Rb
Dim Lb As Integer
Dim Ra: Dim La As Integer
Dim RD
Dim Ld As Integer
Dim oo As Integer
Dim tt As String
Dim LPB As String: Dim LJZ As String
On Error Resume Next
Dim ii As Integer
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
If timZm = 1 Then '保存
    '更新相应的替代适用机组和品牌
     If Left(txtBh.Text, 1) = 9 Or Left(txtBh.Text, 1) = 8 Or Left(txtBh.Text, 1) = 3 Or Left(txtBh.Text, 1) = "B" Then
        tt = "select tdbh from nlptd where bh='" & txtBh.Text & "'"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        Ra = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
     End If
     La = UBound(Ra, 2) + 1
     For oo = 0 To La - 1
        tt = "select pb,jz from nlptdsave where tdbh='" & Ra(0, oo) & "'"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        Rb = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        Lb = UBound(Rb, 2) + 1
        For ii = 0 To Lb - 1
            If InStr(1, LPB, Rb(0, ii)) > 0 Then
            Else
                LPB = LPB & " " & Rb(0, ii)
            End If
            If InStr(1, LJZ, Rb(1, ii)) > 0 Then
            Else
                LJZ = LJZ & " " & Rb(1, ii)
            End If
        Next
        tt = "update nlpmxc set pb='" & LPB & "',jz='" & LJZ & "' where bh='" & Ra(0, oo) & "'"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        Set mod1.HTP = Nothing
     Next
    tt = "select trq,ywy,zn,bz,tf from pizu where bh='" & txtBh.Text & "' and yid=95 order by pid desc"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Rz = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    Lz = UBound(Rz, 2) + 1
    Call QMBound(Rz, Lz)
    '保存后流程返回业务员本人，待经验审核确认
    Lc = 1: LCRen = mod1.DName: LCUid = mod1.DHid
    lblTX.Caption = "流程至：" & LCRen
        
'''''    If NPF = 1 Then
'''''        lblTX.Caption = lblTX.Caption & ",倪工已经审核！"
'''''    End If
    dtgL1.Visible = False
ElseIf timZm = 2 Then

ElseIf timZm = 5 Then '替代
    txtTdbh.Text = ""
    Call TD(txtBh.Text)
    txtPb.Text = "": txtJz.Text = ""
    tt = "select pb,jz,tdbh from nlpmxctd where ybh='" & txtBh.Text & "' order by pb"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Rb = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    Lb = UBound(Rb, 2) + 1
    For oo = 0 To Lb - 1
        txtPb.Text = txtPb.Text & Rb(0, oo) & " "
        txtJz.Text = txtJz.Text & "(" & Rb(0, oo) & ")" & Rb(1, oo) & " "
        txtTD.Text = txtTD.Text & " " & Rb(2, oo)
    Next
    '更新适用品牌和机组和替代编号
    tt = "update nlpmxc set pb='" & txtPb.Text & "',jz='" & txtJz.Text & "',tdbh='" & txtTD.Text & "' where bh='" & txtBh.Text & "'"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set mod1.HTP = Nothing
ElseIf timZm = 6 Then '签字
    tt = "select trq,ywy,zn,bz,tf from pizu where bh='" & txtBh.Text & "' and yid=95 order by pid desc"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Rz = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    Lz = UBound(Rz, 2) + 1
    Call QMBound(Rz, Lz)
ElseIf timZm = 12 Then
    Call Qing
End If
timQuit.Enabled = False
End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
Dim Bid As Long
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh,mm1,mm2,mt2,mt1,mt3,mt4 from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    mod1.Ti = 5
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    timWait.Enabled = False
    If timZm = 1 Then
        Pid = mod1.WP.Fields("mm1").Value
        If LCUid = "" Then
            LCUid = mod1.DHid
            LCRen = mod1.DName
        End If
    ElseIf timZm = 6 Or timZm = 7 Then
                Lc = mod1.WP.Fields("mm1").Value
                Fwid = mod1.WP.Fields("mm2").Value
                LCRen = mod1.WP.Fields("mt1").Value
                LCUid = mod1.WP.Fields("mt2").Value
                lblTX.Caption = "下一流程,将跳至" & mod1.WP.Fields("mt3").Value & ": " & LCRen
                If Lc = 100 Then lblTX.Caption = "审核完毕!"
                
                

    End If
    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then '未完成

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '处理失败
    ii = MsgBox("服务中心在处理您的命令时,发生如下错误:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    Exit Sub
'''''    If timZm = 1 Then
'''''        NiceButton1.Enabled = False
'''''    End If
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
'''''    If timZm = 1 Then
'''''        NiceButton1.Enabled = False
'''''    End If
    Exit Sub
End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub



Public Sub dtgLPFF()
Dim oo As Long
dtgLP.Clear
dtgLP.Rows = 300
dtgLP.Cols = 4
dtgLP.Row = 0
dtgLP.Col = 0: dtgLP.Text = "编号": dtgLP.CellFontBold = True
dtgLP.Col = 1: dtgLP.Text = "货品名称": dtgLP.CellFontBold = True
dtgLP.Col = 2: dtgLP.Text = "描述": dtgLP.CellFontBold = True
dtgLP.Col = 3: dtgLP.Text = Pid: dtgLP.CellFontBold = True

dtgLN.Clear
dtgLN.Rows = 300
dtgLN.Cols = 4
dtgLP.ColWidth(3) = 0
dtgLP.ColWidth(1) = 1860
dtgLP.ColWidth(2) = 7815
For oo = 1 To 299
    dtgLP.RowHeight(oo) = dtgLP.RowHeight(0) * 2
Next
End Sub

Public Sub dtgLPBound(tt As String)
Dim Ra
Dim La


Dim oo As Long
dtgLP.Visible = False
Call dtgLPFF

Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
dtgLP.Rows = La + 30
dtgLN.Rows = La + 30

For oo = 1 To La
    dtgLP.Row = oo
    dtgLP.Col = 0: dtgLP.Text = Ra(0, oo - 1)
    dtgLP.Col = 1: dtgLP.Text = Ra(1, oo - 1)
    dtgLP.Col = 2: dtgLP.Text = Ra(2, oo - 1)
    dtgLP.Col = 3: dtgLP.Text = Ra(3, oo - 1)
    dtgLN.Row = oo
    dtgLN.Col = 0: dtgLN.Text = Ra(0, oo - 1)
    dtgLN.Col = 1: dtgLN.Text = Ra(1, oo - 1)
    dtgLN.Col = 2: dtgLN.Text = Ra(2, oo - 1)
    dtgLN.Col = 3: dtgLN.Text = Ra(3, oo - 1)
    If oo = La Then
        Jpid = Ra(3, oo - 1)
    End If
    If Jpid < 10 Then
        Jpid = 0
    End If
Next
dtgLP.Visible = True
End Sub

Public Sub Bound(Pid As Long)
Dim tt As String
Dim Ra
Dim Rb
Dim RC
Dim RD
Dim RE
Dim Rz
Dim d As String
Dim di As Integer
Dim picL As Integer: Dim Lb As Integer: Dim oo As Integer: Dim Lz As Integer
picL = 1

tt = "declare @bh varchar(10),@gid1 int,@gid2 int,@gid3 int;" & _
    "select @bh=bh,@gid1=gid1,@gid2=gid2,@gid3=gid3 from nlpmxc where pid=" & Pid & ";" & _
    "select bh,partname,engname,lb1,lb2,gg,xn,ff,pb,jz,bz,oname,mj,listprice,ypb,pic,gid1,gid2,gid3,dj1,dj2,dj3,ll,lluid,lcren,lcuid,lc,fwid,jyf,bm1,bm2,bm3,l1,l2,l3,tdbh,npf,dj,cb,gxbh from nlpmxc where pid=" & Pid & ";" & _
    "select pb,jz from nlpmxctd where ybh=@bh order by pb;" & _
    "select mc from gymxc where gid=@gid1;select mc from gymxc where gid=@gid2;select mc from gymxc where gid=@gid3;" & _
        "select trq,ywy,zn,bz,tf from pizu where bh=@bh and yid=95 order by pid desc"

Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText

On Error Resume Next
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
RC = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
RD = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
RE = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rz = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
Lb = UBound(Rb, 2) + 1

txtBh.Text = Ra(0, 0): txtBh1.Text = Ra(0, 0): txtBh2.Text = Ra(0, 0): TmpBh = Ra(0, 0)
txtPartName.Text = Ra(1, 0): txtPartName1.Text = Ra(1, 0): txtPartName2.Text = Ra(1, 0)
txtEngName.Text = Ra(2, 0)
txtLb1.Text = Ra(3, 0)
txtLb2.Text = Ra(4, 0)
txtGG.Text = Ra(5, 0)
txtXN.Text = Ra(6, 0)
txtFF.Text = Ra(7, 0)
txtPb.Text = Ra(8, 0): txtPb1.Text = Ra(8, 0)
txtJz.Text = Ra(9, 0)
txtBz.Text = Ra(10, 0)

txtOname.Text = Ra(11, 0): txtOname2.Text = Ra(11, 0)
txtOname.Text = Replace(txtOname.Text, vbCrLf, "", 1)
txtOname2.Text = Replace(txtOname2.Text, vbCrLf, "", 1)
txtJJ1.Text = Ra(19, 0)
txtJJ2.Text = Ra(20, 0)
txtJJ3.Text = Ra(21, 0)
txtMj.Text = Ra(12, 0)
txtListPrice.Text = Ra(13, 0)
txtYpb.Text = Ra(14, 0)

''''If Left(txtBh.Text, 1) <> "9" Or Left(txtBh.Text, 1) <> "8" Then
''''    txtPb.Text = "": txtJz.Text = ""
''''    For oo = 0 To Lb - 1
''''        txtPb.Text = txtPb.Text & Rb(0, oo) & " "
''''        txtJz.Text = txtJz.Text & "(" & Rb(0, oo) & ")" & Rb(1, oo) & " "
''''    Next
''''End If

Call dtgPicFF
If Ra(15, 0) = True Then
    dtgPic.Visible = False
    
Else

    d = Dir("C:\Documents and Settings\Administrator\桌面\数据对应图片\*" & txtOname.Text & "*.jpg")
    Do Until d = ""
       'MsgBox "找到一个图片文件：" & d

       dtgPic.Row = picL
       dtgPic.Text = d
       d = Dir
        picL = picL + 1
    Loop
    d = Dir("C:\Documents and Settings\Administrator\桌面\数据对应图片\" & txtBh.Text & ".jpg")
    Do Until d = ""
       'MsgBox "找到一个图片文件：" & d

       dtgPic.Row = picL
       dtgPic.Text = d
       d = Dir
        picL = picL + 1
    Loop
    dtgPic.Visible = True
End If
txtGy1.ToolTipText = Ra(16, 0)
txtGy2.ToolTipText = Ra(17, 0)
txtGY3.ToolTipText = Ra(18, 0)
txtGy1.Text = RC(0, 0)
txtGy2.Text = RD(0, 0)
txtGY3.Text = RE(0, 0)

LL = Ra(22, 0)
LLUid = Ra(23, 0)
LCRen = Ra(24, 0)
LCUid = Ra(25, 0)
Lc = Ra(26, 0)
        lblTX.Caption = "流程至：" & LCRen
        If Lc = 100 Then lblTX.Caption = "审核完毕!"
Fwid = Ra(27, 0)
If Ra(28, 0) = True Then
    chkJYF.Value = 0
Else
    chkJYF.Value = 1
End If
txtBm1.Text = Ra(29, 0)
txtBm2.Text = Ra(30, 0)
txtBm3.Text = Ra(31, 0)
txtL1.Text = Ra(32, 0)
txtL2.Text = Ra(33, 0)
txtL3.Text = Ra(34, 0)
txtTD.Text = Ra(35, 0)
NPF = Ra(36, 0) '倪工审核否
txtDj.Text = Ra(37, 0) '单价
txtBh.ToolTipText = Ra(39, 0)
''''''If NPF = 1 Then
''''''    lblTX.Caption = lblTX.Caption & ",倪工已经审核！"
''''''End If
Lz = UBound(Rz, 2) + 1
Call Me.QMBound(Rz, Lz)
Me.Pid = Pid
'frm2.Visible = False
frm1.Enabled = True
txtBz.Enabled = True
If Left(txtBh.Text, 1) = "3" And (mod1.DName = "李午阳" Or mod1.DName = "货品录入员") Then
    cmdZ.Visible = True
Else
    cmdZ.Visible = False
End If
End Sub

Private Sub txtGy_Change()
Dim tt As String
Dim Ra
Dim La As Long
Dim oo As Long
If Len(txtGy.Text) < 2 Then Exit Sub
'tt = "select mc,gid from gymxc where mc like '%" & txtGy.Text & "%' and delf=1 and lc=100"
tt = "select mc,gid from gymxc where mc like '%" & txtGy.Text & "%' and delf=1"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
Call Me.dtgbrFF
For oo = 1 To La
    dtgBr.Row = oo
    dtgBr.Col = 0: dtgBr.Text = Ra(0, oo - 1)
    dtgBr.Col = 1: dtgBr.Text = Ra(1, oo - 1)
Next
End Sub

Private Sub txtGy1_Click()
GyId = 1
End Sub

Private Sub txtGy1_DblClick()
Me.GyId = 1
Call frmGY.dtgBFF
frmGY.Show
End Sub


Private Sub txtGy2_Click()
GyId = 2
End Sub

Private Sub txtGy2_DblClick()
Me.GyId = 2
Call frmGY.dtgBFF
frmGY.Show
End Sub


Private Sub txtGy3_Click()
GyId = 3
End Sub

Private Sub txtGy3_DblClick()
Me.GyId = 3
Call frmGY.dtgBFF
frmGY.Show
End Sub



Public Sub dtgPicFF()
dtgPic.Clear
dtgPic.Row = 0
dtgPic.Col = 1: dtgPic.Text = "关联的图片文件名(双击则正式关联!)": dtgPic.CellFontBold = True
dtgPic.ColWidth(1) = 3540
dtgPic.ColWidth(0) = 0
End Sub

Private Sub txtName_Change()
Dim tt As String
Dim Ra
Dim La

End Sub

Private Sub txtOname_LostFocus()
'倪工修改
Dim Ra

    If txtOname.Text <> "" Then  '检测原厂编号是否有重复
        tt = "select pid from nlpmxc where oname='" & txtOname.Text & "' and jyf=1 and delf=1 and bh<>'" & _
        txtBh.Text & "' and bh<>'" & txtBh.ToolTipText & "'"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        If mod1.HTP.BOF = False Then
            Ra = mod1.HTP.GetRows
            mod1.HTP.Close
            Set mod1.HTP = Nothing
            ii = MsgBox("检测到有相同原厂编号!,不能添加新货品,是否返回到以前存在的货品?", vbYesNo + vbQuestion, mod1.chenHu)
            If ii = vbYes Then
                Call Bound(Val(Ra(0, 0)))
            Else
                txtOname.Text = ""
            End If
            Exit Sub
        Else
            Set mod1.HTP = Nothing
        End If
    End If
End Sub


Private Sub txtTdbh_DblClick()
txtTdbh.Text = ""
End Sub


Public Sub TD(Bh As String)
Dim tt As String
Dim Ra, Rb
Dim La, Lb
Dim oo As Long
Dim R1, R2, R3, R4, R5, R6, R7
tt = "select bh,partname,'原厂编号:'+oname+' '+gg+' '+xn+' '+ff+'适用机组:'+jz,pid from nlpmxc where bh='" & Bh & "' and delf=1 ;" & _
    "select bh,partname,detail,pid from nlpmxcTd where ybh='" & Bh & "' and delf=1  order by tid desc"
Call dtgLPFF
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
Lb = UBound(Rb, 2) + 1
dtgLP.Rows = Lb + 30
dtgLN.Rows = Lb + 30
dtgLP.Visible = False
dtgLP.Row = 1
    dtgLP.Col = 0: dtgLP.Text = Ra(0, 0)
    dtgLP.Col = 1: dtgLP.Text = Ra(1, 0)
    dtgLP.Col = 2: dtgLP.Text = Ra(2, 0)
    dtgLP.Col = 3: dtgLP.Text = Ra(3, 0)
    dtgLN.Row = 1
    dtgLN.Col = 0: dtgLN.Text = Ra(0, 0)
    dtgLN.Col = 1: dtgLN.Text = Ra(1, 0)
    dtgLN.Col = 2: dtgLN.Text = Ra(2, 0)
    dtgLN.Col = 3: dtgLN.Text = Ra(3, 0)
For oo = 2 To Lb + 1
    dtgLP.Row = oo
    dtgLP.Col = 0: dtgLP.Text = Rb(0, oo - 2)
    dtgLP.Col = 1: dtgLP.Text = Rb(1, oo - 2)
    dtgLP.Col = 2: dtgLP.Text = Rb(2, oo - 2)
    dtgLP.Col = 3: dtgLP.Text = Rb(3, oo - 2)
    dtgLN.Row = oo
    dtgLN.Col = 0: dtgLN.Text = Rb(0, oo - 2)
    dtgLN.Col = 1: dtgLN.Text = Rb(1, oo - 2)
    dtgLN.Col = 2: dtgLN.Text = Rb(2, oo - 2)
    dtgLN.Col = 3: dtgLN.Text = Rb(3, oo - 2)
'''''    If oo = Lb Then
'''''        Jpid = Rb(3, oo - 1)
'''''    End If
'''''    If Jpid < 10 Then
'''''        Jpid = 0
'''''    End If
Next
dtgLP.Visible = True
End Sub
Public Sub TDB(Bh As String)
Dim tt As String
Dim Ra, Rb
Dim La, Lb
Dim oo As Long
Dim R1, R2, R3, R4, R5, R6, R7
tt = "select bh,partname,'原厂编号:'+oname+' '+gg+' '+xn+' '+ff+'适用机组:'+jz,pid from nlpmxc where bh='" & Bh & "' and delf=1;" & _
    "select bh,partname,detail,pid from nlpmxcTdb where ybh='" & Bh & "' and delf=1 order by tid desc"
Call dtgLPFF
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
Lb = UBound(Rb, 2) + 1
dtgLP.Rows = Lb + 30
dtgLN.Rows = Lb + 30
dtgLP.Visible = False
dtgLP.Row = 1
    dtgLP.Col = 0: dtgLP.Text = Ra(0, 0)
    dtgLP.Col = 1: dtgLP.Text = Ra(1, 0)
    dtgLP.Col = 2: dtgLP.Text = Ra(2, 0)
    dtgLP.Col = 3: dtgLP.Text = Ra(3, 0)
    dtgLN.Row = 1
    dtgLN.Col = 0: dtgLN.Text = Ra(0, 0)
    dtgLN.Col = 1: dtgLN.Text = Ra(1, 0)
    dtgLN.Col = 2: dtgLN.Text = Ra(2, 0)
    dtgLN.Col = 3: dtgLN.Text = Ra(3, 0)
For oo = 2 To Lb + 1
    dtgLP.Row = oo
    dtgLP.Col = 0: dtgLP.Text = Rb(0, oo - 2)
    dtgLP.Col = 1: dtgLP.Text = Rb(1, oo - 2)
    dtgLP.Col = 2: dtgLP.Text = Rb(2, oo - 2)
    dtgLP.Col = 3: dtgLP.Text = Rb(3, oo - 2)
    dtgLN.Row = oo
    dtgLN.Col = 0: dtgLN.Text = Rb(0, oo - 2)
    dtgLN.Col = 1: dtgLN.Text = Rb(1, oo - 2)
    dtgLN.Col = 2: dtgLN.Text = Rb(2, oo - 2)
    dtgLN.Col = 3: dtgLN.Text = Rb(3, oo - 2)
'''''    If oo = Lb Then
'''''        Jpid = Rb(3, oo - 1)
'''''    End If
'''''    If Jpid < 10 Then
'''''        Jpid = 0
'''''    End If
Next
dtgLP.Visible = True
End Sub



Public Sub dtgbrFF()
On Error Resume Next
dtgBr.Clear
dtgBr.Rows = 50
dtgBr.Cols = 2
dtgBr.Row = 0
dtgBr.Col = 0: dtgBr.Text = "供应商名称（鼠标双击选择）": dtgBr.CellFontBold = True
dtgBr.ColWidth(1) = 0
dtgBr.ColWidth(0) = 3000

End Sub

Public Sub NewBh(B1 As String)
Dim tt As String: Dim Bh As String
Dim Tbh As String
Dim Ra
Dim L1 As String: Dim L2 As String: Dim L3 As String: Dim L4 As String: Dim l5 As String
tt = "select mbh from l1 where bm='" & B1 & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
Bh = Ra(0, 0)

Jbh:
l5 = Right(Bh, 1)
L1 = Left(Bh, 1)
L2 = Mid(Bh, 2, 1)
L3 = Mid(Bh, 3, 1)
L4 = Mid(Bh, 4, 1)

If Val(l5) > 0 And Val(l5) < 9 Or l5 = "0" Then
    l5 = Val(l5) + 1
ElseIf Val(l5) = 9 Then
    l5 = "A"
ElseIf UCase(l5) = "A" Then
    l5 = "B"
ElseIf UCase(l5) = "B" Then
    l5 = "C"
ElseIf UCase(l5) = "C" Then
    l5 = "D"
ElseIf UCase(l5) = "D" Then
    l5 = "E"
ElseIf UCase(l5) = "E" Then
    l5 = "F"
ElseIf UCase(l5) = "F" Then
    l5 = "G"
ElseIf UCase(l5) = "G" Then
    l5 = "H"
ElseIf UCase(l5) = "H" Then
    l5 = "I"
ElseIf UCase(l5) = "I" Then
    l5 = "J"
ElseIf UCase(l5) = "J" Then
    l5 = "K"
ElseIf UCase(l5) = "K" Then
    l5 = "L"
ElseIf UCase(l5) = "L" Then
    l5 = "M"
ElseIf UCase(l5) = "M" Then
    l5 = "N"
ElseIf UCase(l5) = "N" Then
    l5 = "P"
ElseIf UCase(l5) = "O" Then
    l5 = "P"
ElseIf UCase(l5) = "P" Then
    l5 = "Q"
ElseIf UCase(l5) = "Q" Then
    l5 = "R"
ElseIf UCase(l5) = "R" Then
    l5 = "S"
ElseIf UCase(l5) = "S" Then
    l5 = "T"
ElseIf UCase(l5) = "T" Then
    l5 = "U"
ElseIf UCase(l5) = "U" Then
    l5 = "V"
ElseIf UCase(l5) = "V" Then
    l5 = "W"
ElseIf UCase(l5) = "W" Then
    l5 = "X"
ElseIf UCase(l5) = "X" Then
    l5 = "Y"
ElseIf UCase(l5) = "Y" Then
    l5 = "Z"
ElseIf UCase(l5) = "Z" Then
    '进位'
    l5 = "0"
    Call NewBh4(L4, L3, L2, L1)
End If
Lbh = L1 & L2 & L3 & L4 & l5
tt = "select bh from nlpmxc where bh='" & Lbh & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
If mod1.HTP.BOF = True Then
    Set mod1.HTP = Nothing
    If cmdZ.Visible = False Then
        txtBh.Text = Lbh
    Else
'''        Tbh = TmpBh
'''        Call Me.Qing
'''        txtBh.ToolTipText = Tbh
        txtBh.Text = Lbh
        Me.Pid = 0
        'cmdSave.Enabled = True
    End If
    tt = "update l1 set mbh='" & Lbh & "' where bm='" & L1 & "'"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set mod1.HTP = Nothing
Else
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    tt = "update l1 set mbh='" & Lbh & "' where bm='" & L1 & "'"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    'Exit Sub
    Set mod1.HTP = Nothing
    Bh = Lbh
    GoTo Jbh
End If
End Sub

Public Sub NewBh4(L4 As String, L3 As String, L2 As String, L1 As String)
Dim tt As String

If Val(L4) > 0 And Val(L4) < 9 Or L4 = "0" Then
    L4 = Val(L4) + 1
ElseIf Val(L4) = 9 Then
    L4 = "A"
ElseIf UCase(L4) = "A" Then
    L4 = "B"
ElseIf UCase(L4) = "B" Then
    L4 = "C"
ElseIf UCase(L4) = "C" Then
    L4 = "D"
ElseIf UCase(L4) = "D" Then
    L4 = "E"
ElseIf UCase(L4) = "E" Then
    L4 = "F"
ElseIf UCase(L4) = "F" Then
    L4 = "G"
ElseIf UCase(L4) = "G" Then
    L4 = "H"
ElseIf UCase(L4) = "H" Then
    L4 = "I"
ElseIf UCase(L4) = "I" Then
    L4 = "J"
ElseIf UCase(L4) = "J" Then
    L4 = "K"
ElseIf UCase(L4) = "K" Then
    L4 = "L"
ElseIf UCase(L4) = "L" Then
    L4 = "M"
ElseIf UCase(L4) = "M" Then
    L4 = "N"
ElseIf UCase(L4) = "N" Then
    L4 = "P"
ElseIf UCase(L4) = "O" Then
    L4 = "P"
ElseIf UCase(L4) = "P" Then
    L4 = "Q"
ElseIf UCase(L4) = "Q" Then
    L4 = "R"
ElseIf UCase(L4) = "R" Then
    L4 = "S"
ElseIf UCase(L4) = "S" Then
    L4 = "T"
ElseIf UCase(L4) = "T" Then
    L4 = "U"
ElseIf UCase(L4) = "U" Then
    L4 = "V"
ElseIf UCase(L4) = "V" Then
    L4 = "W"
ElseIf UCase(L4) = "W" Then
    L4 = "X"
ElseIf UCase(L4) = "X" Then
    L4 = "Y"
ElseIf UCase(L4) = "Y" Then
    L4 = "Z"
ElseIf UCase(L4) = "Z" Then
    '进位'
    L4 = "0"
    Call NewBh3(L4, L3, L2, L1)
End If

End Sub

Public Sub NewBh3(L4 As String, L3 As String, L2 As String, L1 As String)
Dim tt As String

If Val(L3) > 0 And Val(L3) < 9 Or L3 = "0" Then
    L3 = Val(L3) + 1
ElseIf Val(L3) = 9 Then
    L3 = "A"
ElseIf UCase(L3) = "A" Then
    L3 = "B"
ElseIf UCase(L3) = "B" Then
    L3 = "C"
ElseIf UCase(L3) = "C" Then
    L3 = "D"
ElseIf UCase(L3) = "D" Then
    L3 = "E"
ElseIf UCase(L3) = "E" Then
    L3 = "F"
ElseIf UCase(L3) = "F" Then
    L3 = "G"
ElseIf UCase(L3) = "G" Then
    L3 = "H"
ElseIf UCase(L3) = "H" Then
    L3 = "I"
ElseIf UCase(L3) = "I" Then
    L3 = "J"
ElseIf UCase(L3) = "J" Then
    L3 = "K"
ElseIf UCase(L3) = "K" Then
    L3 = "L"
ElseIf UCase(L3) = "L" Then
    L3 = "M"
ElseIf UCase(L3) = "M" Then
    L3 = "N"
ElseIf UCase(L3) = "N" Then
    L3 = "P"
ElseIf UCase(L3) = "O" Then
    L3 = "P"
ElseIf UCase(L3) = "P" Then
    L3 = "Q"
ElseIf UCase(L3) = "Q" Then
    L3 = "R"
ElseIf UCase(L3) = "R" Then
    L3 = "S"
ElseIf UCase(L3) = "S" Then
    L3 = "T"
ElseIf UCase(L3) = "T" Then
    L3 = "U"
ElseIf UCase(L3) = "U" Then
    L3 = "V"
ElseIf UCase(L3) = "V" Then
    L3 = "W"
ElseIf UCase(L3) = "W" Then
    L3 = "X"
ElseIf UCase(L3) = "X" Then
    L3 = "Y"
ElseIf UCase(L3) = "Y" Then
    L3 = "Z"
ElseIf UCase(L3) = "Z" Then
    '进位'
    L3 = "0"
    Call NewBh2(L4, L3, L2, L1)
End If

End Sub

Public Sub NewBh2(L4 As String, L3 As String, L2 As String, L1 As String)
Dim tt As String

If (Val(L2) > 0 Or L2 = "0") And Val(L2) < 9 Then
    L2 = Val(L2) + 1
ElseIf L2 = "9" Then
    L2 = "A"
ElseIf UCase(L2) = "A" Then
    L2 = "B"
ElseIf UCase(L2) = "B" Then
    L2 = "C"
ElseIf UCase(L2) = "C" Then
    L2 = "D"
ElseIf UCase(L2) = "D" Then
    L2 = "E"
ElseIf UCase(L2) = "E" Then
    L2 = "F"
ElseIf UCase(L2) = "F" Then
    L2 = "G"
ElseIf UCase(L2) = "G" Then
    L2 = "H"
ElseIf UCase(L2) = "H" Then
    L2 = "I"
ElseIf UCase(L2) = "I" Then
    L2 = "J"
ElseIf UCase(L2) = "J" Then
    L2 = "K"
ElseIf UCase(L2) = "K" Then
    L2 = "L"
ElseIf UCase(L2) = "L" Then
    L2 = "M"
ElseIf UCase(L2) = "M" Then
    L2 = "N"
ElseIf UCase(L2) = "N" Then
    L2 = "P"
ElseIf UCase(L2) = "O" Then
    L2 = "P"
ElseIf UCase(L2) = "P" Then
    L2 = "Q"
ElseIf UCase(L2) = "Q" Then
    L2 = "R"
ElseIf UCase(L2) = "R" Then
    L2 = "S"
ElseIf UCase(L2) = "S" Then
    L2 = "T"
ElseIf UCase(L2) = "T" Then
    L2 = "U"
ElseIf UCase(L2) = "U" Then
    L2 = "V"
ElseIf UCase(L2) = "V" Then
    L2 = "W"
ElseIf UCase(L2) = "W" Then
    L2 = "X"
ElseIf UCase(L2) = "X" Then
    L2 = "Y"
ElseIf UCase(L2) = "Y" Then
    L2 = "Z"
ElseIf UCase(L2) = "Z" Then
    '进位'
    MsgBox "编号已经达到极限，需重设第一位编码，请速与马晓聪联系！"
    Exit Sub
End If

End Sub

Public Sub CreateQuan(ft As String)
Select Case ft
    Case "3"
        txtOname.Locked = False: txtPartName.Locked = False: txtBm1.Locked = False: txtBm2.Locked = False: txtBm3.Locked = False
        txtYpb.Locked = False: txtEngName.Locked = False: txtGG.Locked = False: txtXN.Locked = False
        txtL1.Locked = False: txtL2.Locked = False: txtL3.Locked = False
        txtPb.Locked = False: txtJz.Locked = False: txtBz.Locked = False
    Case "A"
        txtOname.Locked = False: txtPartName.Locked = False: txtBm1.Locked = False: txtBm2.Locked = False: txtBm3.Locked = False
        txtYpb.Locked = False: txtEngName.Locked = False: txtGG.Locked = False: txtXN.Locked = False
        txtL1.Locked = False: txtL2.Locked = False: txtL3.Locked = False
        txtPb.Locked = False: txtJz.Locked = False: txtBz.Locked = False
    Case "H"
        txtOname.Locked = False: txtPartName.Locked = False: txtBm1.Locked = False: txtBm2.Locked = False: txtBm3.Locked = False
        txtYpb.Locked = False: txtEngName.Locked = False: txtGG.Locked = False: txtXN.Locked = False
        txtL1.Locked = False: txtL2.Locked = False: txtL3.Locked = False
        txtPb.Locked = True: txtJz.Locked = True '替代生成
        txtBz.Locked = False
    Case "B"
        txtOname.Locked = False: txtPartName.Locked = False: txtBm1.Locked = False: txtBm2.Locked = False: txtBm3.Locked = False
        txtYpb.Locked = False: txtEngName.Locked = False: txtGG.Locked = False: txtXN.Locked = False
        txtL1.Locked = False: txtL2.Locked = False: txtL3.Locked = False
        txtPb.Locked = False: txtJz.Locked = False: txtBz.Locked = False
    Case "9"
        txtOname.Locked = False: txtPartName.Locked = False: txtBm1.Locked = False: txtBm2.Locked = False: txtBm3.Locked = False
        txtYpb.Locked = False: txtEngName.Locked = False: txtGG.Locked = False: txtXN.Locked = False
        txtL1.Locked = False: txtL2.Locked = False: txtL3.Locked = False
        txtPb.Locked = True '按原厂品牌自动生成
        txtJz.Locked = False: txtBz.Locked = False
    Case "8"
        txtOname.Locked = False: txtPartName.Locked = False: txtBm1.Locked = False: txtBm2.Locked = False: txtBm3.Locked = False
        txtYpb.Locked = False: txtEngName.Locked = False: txtGG.Locked = False: txtXN.Locked = False
        txtL1.Locked = False: txtL2.Locked = False: txtL3.Locked = False
        txtPb.Locked = True '按原厂品牌自动生成
        txtJz.Locked = False: txtBz.Locked = False
    Case "1"
        txtOname.Locked = False: txtPartName.Locked = False: txtBm1.Locked = False: txtBm2.Locked = False: txtBm3.Locked = False
        txtYpb.Locked = False: txtEngName.Locked = False: txtGG.Locked = False: txtXN.Locked = False
        txtL1.Locked = False: txtL2.Locked = False: txtL3.Locked = False
        txtPb.Locked = True '按原厂品牌自动生成＋替代生成
        txtJz.Locked = True '替代生成
        txtBz.Locked = False
End Select
End Sub

Private Sub txtXN_LostFocus()
'倪工修改
Dim Ra

    If txtOname.Text <> "" Then  '检测原厂编号是否有重复
        tt = "select pid from nlpmxc where xn='" & txtXN.Text & "' and jyf=1 and delf=1 and bh<>'" & _
        txtBh.Text & "' and bh<>'" & txtBh.ToolTipText & "'"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        If mod1.HTP.BOF = False Then
            Ra = mod1.HTP.GetRows
            mod1.HTP.Close
            Set mod1.HTP = Nothing
            ii = MsgBox("检测到有相同产品型号！", vbInformation + vbOKOnly, mod1.chenHu)
'''            If ii = vbYes Then
'''                Call Bound(Val(Ra(0, 0)))
'''            Else
'''                txtOname.Text = ""
'''            End If
            Exit Sub
        Else
            Set mod1.HTP = Nothing
        End If
    End If
End Sub


Private Sub txtYpb_LostFocus()
''''''''If txtYpb.Text <> "" And (Left(txtBh.Text, 1) = "9" Or Left(txtBh.Text, 1) = "1") Then
''''''''    txtPb.Text = txtYpb.Text
''''''''End If
End Sub

