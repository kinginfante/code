VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGXBj 
   Caption         =   "购销询价单"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdNQ 
      BackColor       =   &H008080FF&
      Caption         =   "审核"
      Height          =   765
      Left            =   7980
      Picture         =   "frmGXBj.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   155
      Top             =   8400
      Width           =   675
   End
   Begin VB.CommandButton cmdGy 
      BackColor       =   &H00FF8080&
      Caption         =   "供应商资料"
      Height          =   285
      Left            =   3450
      Style           =   1  'Graphical
      TabIndex        =   112
      Top             =   5940
      Width           =   1035
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "导入至EXCEL"
      Height          =   345
      Left            =   13830
      TabIndex        =   154
      Top             =   4290
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Frame frmQm 
      BackColor       =   &H00C0FFC0&
      Caption         =   "评审建议"
      ForeColor       =   &H000000FF&
      Height          =   1785
      Left            =   9570
      TabIndex        =   121
      Top             =   5520
      Visible         =   0   'False
      Width           =   6315
      Begin VB.CommandButton cmdDing 
         BackColor       =   &H00FF8080&
         Caption         =   "决定"
         Height          =   285
         Left            =   5250
         Style           =   1  'Graphical
         TabIndex        =   125
         Top             =   1350
         Width           =   735
      End
      Begin VB.OptionButton optT2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "拒绝"
         Height          =   195
         Left            =   5220
         TabIndex        =   124
         Top             =   870
         Width           =   675
      End
      Begin VB.OptionButton OptT1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "同意"
         Height          =   225
         Left            =   5220
         TabIndex        =   123
         Top             =   480
         Width           =   705
      End
      Begin VB.TextBox txtQM 
         BackColor       =   &H00C0FFFF&
         Height          =   1365
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   122
         Top             =   300
         Width           =   4965
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   585
      Left            =   9960
      TabIndex        =   147
      Top             =   4620
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1032
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame frmSd 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   435
      Left            =   60
      TabIndex        =   141
      Top             =   4260
      Visible         =   0   'False
      Width           =   9465
      Begin VB.TextBox txtJzxh 
         Height          =   270
         Left            =   3870
         TabIndex        =   151
         Top             =   30
         Width           =   1335
      End
      Begin VB.CommandButton cmdDao 
         BackColor       =   &H0000C000&
         Caption         =   "货品添加"
         Height          =   345
         Left            =   8430
         Style           =   1  'Graphical
         TabIndex        =   146
         Top             =   -30
         Width           =   1005
      End
      Begin VB.CommandButton cmdNGx 
         BackColor       =   &H00C0FFC0&
         Caption         =   "更新"
         Height          =   315
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   144
         Top             =   0
         Width           =   825
      End
      Begin VB.TextBox txtNsl 
         Height          =   270
         Left            =   5790
         TabIndex        =   143
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdNDel 
         BackColor       =   &H008080FF&
         Caption         =   "删除"
         Height          =   345
         Left            =   7470
         Style           =   1  'Graphical
         TabIndex        =   142
         Top             =   -30
         Width           =   885
      End
      Begin MSDataListLib.DataCombo comJzPb1 
         Height          =   330
         Left            =   990
         TabIndex        =   148
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "机组型号"
         Height          =   255
         Left            =   3030
         TabIndex        =   150
         Top             =   60
         Width           =   915
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "机组品牌"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   150
         TabIndex        =   149
         ToolTipText     =   "机组品牌"
         Top             =   60
         Width           =   855
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "数量"
         Height          =   225
         Left            =   5340
         TabIndex        =   145
         Top             =   60
         Width           =   375
      End
   End
   Begin VB.Frame frmWai 
      BackColor       =   &H00FF8080&
      Caption         =   "外购资料"
      Height          =   3075
      Left            =   3150
      TabIndex        =   97
      Top             =   6510
      Visible         =   0   'False
      Width           =   4875
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgGy 
         Height          =   1455
         Left            =   1650
         TabIndex        =   113
         Top             =   1020
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   2566
         _Version        =   393216
         BackColor       =   16777215
         FixedRows       =   0
         BackColorFixed  =   8454016
         BackColorBkg    =   8454016
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton cmdOpen 
         BackColor       =   &H00FF8080&
         Caption         =   "..."
         Height          =   255
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   690
         Width           =   465
      End
      Begin VB.TextBox txtGyid 
         Height          =   270
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   105
         Top             =   270
         Width           =   2595
      End
      Begin VB.TextBox txtGYmc 
         Height          =   345
         Left            =   1680
         TabIndex        =   104
         Top             =   660
         Width           =   2595
      End
      Begin VB.TextBox txtGyman 
         Height          =   315
         Left            =   1680
         TabIndex        =   103
         Top             =   1140
         Width           =   2595
      End
      Begin VB.TextBox txtGyAdr 
         Height          =   285
         Left            =   1680
         TabIndex        =   102
         Top             =   1650
         Width           =   2595
      End
      Begin VB.CommandButton cmdGB 
         Caption         =   "关闭"
         Height          =   345
         Left            =   4020
         TabIndex        =   101
         Top             =   2640
         Width           =   825
      End
      Begin VB.TextBox txtGYPho 
         Height          =   315
         Left            =   1680
         TabIndex        =   100
         Top             =   2070
         Width           =   2625
      End
      Begin VB.CommandButton cmdGadd 
         Caption         =   "清空"
         Height          =   285
         Left            =   120
         TabIndex        =   99
         Top             =   2670
         Width           =   825
      End
      Begin VB.CommandButton cmdGsave 
         Caption         =   "保存"
         Height          =   285
         Left            =   1020
         TabIndex        =   98
         Top             =   2670
         Width           =   705
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "供应商名称"
         Height          =   225
         Left            =   210
         TabIndex        =   110
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "联系人"
         Height          =   285
         Left            =   210
         TabIndex        =   109
         Top             =   1170
         Width           =   885
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "编号:"
         Height          =   255
         Left            =   210
         TabIndex        =   108
         Top             =   330
         Width           =   1095
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   "地址:"
         Height          =   195
         Left            =   240
         TabIndex        =   107
         Top             =   1710
         Width           =   855
      End
      Begin VB.Label Label57 
         BackStyle       =   0  'Transparent
         Caption         =   "电话:"
         Height          =   285
         Left            =   240
         TabIndex        =   106
         Top             =   2130
         Width           =   705
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgGM 
      Height          =   1455
      Left            =   2550
      TabIndex        =   120
      Top             =   6600
      Visible         =   0   'False
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   2566
      _Version        =   393216
      BackColor       =   16777215
      FixedRows       =   0
      BackColorFixed  =   8454016
      BackColorBkg    =   16744576
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgP 
      Height          =   2865
      Left            =   -60
      TabIndex        =   140
      Top             =   7620
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   5054
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
   Begin VB.Frame frmCg 
      Height          =   1545
      Left            =   30
      TabIndex        =   49
      Top             =   4890
      Width           =   9135
      Begin VB.TextBox txtZBQ 
         Height          =   270
         Left            =   7380
         TabIndex        =   153
         Top             =   750
         Width           =   1665
      End
      Begin VB.Frame frmJ 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   195
         Left            =   1890
         TabIndex        =   128
         Top             =   360
         Width           =   2235
         Begin VB.TextBox txtJdj 
            Height          =   270
            Left            =   930
            TabIndex        =   129
            Text            =   "Text1"
            Top             =   0
            Width           =   1155
         End
         Begin VB.Label Label16 
            Caption         =   "基准单价"
            Height          =   255
            Left            =   120
            TabIndex        =   130
            Top             =   30
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdGyOpen 
         BackColor       =   &H00FF8080&
         Caption         =   "..."
         Height          =   285
         Left            =   2820
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   1230
         Width           =   675
      End
      Begin VB.TextBox txtGybz 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   5280
         TabIndex        =   118
         Top             =   1230
         Width           =   3105
      End
      Begin VB.TextBox txtGM 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   116
         Top             =   1200
         Width           =   1905
      End
      Begin VB.ComboBox txtAdr 
         Height          =   300
         ItemData        =   "frmGXBj.frx":0442
         Left            =   2820
         List            =   "frmGXBj.frx":044F
         TabIndex        =   81
         Text            =   "Combo1"
         Top             =   780
         Width           =   3555
      End
      Begin VB.TextBox txtYf 
         Height          =   285
         Left            =   810
         TabIndex        =   79
         Text            =   "Text1"
         Top             =   780
         Width           =   1065
      End
      Begin VB.Frame frmZ 
         Height          =   405
         Left            =   -8310
         TabIndex        =   70
         Top             =   690
         Width           =   8295
      End
      Begin VB.TextBox txtDrq 
         Height          =   330
         Left            =   4830
         TabIndex        =   67
         Top             =   360
         Width           =   1485
      End
      Begin VB.TextBox txtMj 
         Height          =   270
         Left            =   810
         TabIndex        =   66
         Top             =   390
         Width           =   1065
      End
      Begin VB.CommandButton cmdGx 
         BackColor       =   &H008080FF&
         Caption         =   "更新"
         Height          =   315
         Left            =   8490
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   1230
         Width           =   645
      End
      Begin VB.TextBox txtBrq 
         Height          =   315
         Left            =   7380
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   330
         Width           =   1365
      End
      Begin MSComCtl2.DTPicker dtpBrq 
         Height          =   315
         Left            =   7380
         TabIndex        =   52
         Top             =   330
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   8454016
         CalendarTitleBackColor=   16711808
         CalendarTrailingForeColor=   -2147483635
         Format          =   50659329
         CurrentDate     =   38797
      End
      Begin VB.TextBox txtDj 
         Height          =   270
         Left            =   2820
         TabIndex        =   68
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label27 
         Caption         =   "质保期"
         Height          =   255
         Left            =   6780
         TabIndex        =   152
         Top             =   810
         Width           =   615
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "备注:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4710
         TabIndex        =   117
         Top             =   1260
         Width           =   525
      End
      Begin VB.Label Label21 
         Caption         =   "供应商:"
         ForeColor       =   &H00FF0000&
         Height          =   165
         Left            =   120
         TabIndex        =   115
         Top             =   1260
         Width           =   645
      End
      Begin VB.Label Label18 
         Caption         =   "送货地址"
         Height          =   255
         Left            =   2010
         TabIndex        =   80
         Top             =   855
         Width           =   825
      End
      Begin VB.Label lblYf 
         Caption         =   "运费"
         Height          =   225
         Left            =   300
         TabIndex        =   78
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label15 
         Caption         =   "到货期"
         Height          =   255
         Left            =   4170
         TabIndex        =   69
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "市场价"
         Height          =   315
         Left            =   120
         TabIndex        =   65
         Top             =   420
         Width           =   795
      End
      Begin VB.Label Label17 
         Caption         =   "报价有效期"
         Height          =   315
         Left            =   6420
         TabIndex        =   51
         Top             =   390
         Width           =   1065
      End
      Begin VB.Label lblDj 
         Caption         =   "成本单价"
         Height          =   225
         Left            =   2040
         TabIndex        =   50
         Top             =   420
         Width           =   765
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00008000&
      Caption         =   "速达金额"
      Height          =   315
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   138
      ToolTipText     =   "根据合同金额自动分配,点击将刷新"
      Top             =   4830
      Width           =   945
   End
   Begin VB.TextBox txtFJ 
      Height          =   285
      Left            =   13860
      TabIndex        =   134
      ToolTipText     =   "此处由商务支持部添加"
      Top             =   7050
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdD 
      Enabled         =   0   'False
      Height          =   405
      Left            =   14280
      Picture         =   "frmGXBj.frx":046C
      Style           =   1  'Graphical
      TabIndex        =   132
      Top             =   8760
      Width           =   465
   End
   Begin VB.CommandButton cmdFk 
      BackColor       =   &H00FFC0C0&
      Caption         =   "付款条件参考"
      Height          =   435
      Left            =   13950
      Style           =   1  'Graphical
      TabIndex        =   131
      Top             =   330
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmdBjd 
      BackColor       =   &H00C0C0FF&
      Caption         =   "报价单"
      Height          =   315
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   900
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmdHT 
      BackColor       =   &H00C0FFC0&
      Caption         =   "合同评审单"
      Height          =   405
      Left            =   13980
      Style           =   1  'Graphical
      TabIndex        =   114
      Top             =   7350
      Width           =   1305
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11520
      Top             =   7980
   End
   Begin VB.Timer timQuit 
      Interval        =   1000
      Left            =   12270
      Top             =   8130
   End
   Begin VB.Frame frmNew 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   11040
      TabIndex        =   85
      Top             =   2100
      Visible         =   0   'False
      Width           =   4305
      Begin VB.CommandButton cmdWb 
         Caption         =   "人工询价"
         Height          =   315
         Left            =   330
         TabIndex        =   137
         Top             =   210
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txtHtbh 
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   136
         Top             =   0
         Visible         =   0   'False
         Width           =   3315
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   585
         Left            =   150
         TabIndex        =   88
         Top             =   330
         Width           =   705
         Begin VB.OptionButton opt1 
            Caption         =   "内部"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   30
            TabIndex        =   90
            Top             =   0
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.OptionButton opt2 
            Caption         =   "外部"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   30
            TabIndex        =   89
            Top             =   300
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdJG 
         BackColor       =   &H008080FF&
         Caption         =   "选购决定"
         Height          =   495
         Left            =   1170
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   360
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label20 
         Caption         =   "从何处购买"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   180
         TabIndex        =   86
         Top             =   30
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin VB.TextBox txtBz 
      Height          =   1185
      Left            =   10260
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   77
      Text            =   "frmGXBj.frx":05F6
      Top             =   5640
      Width           =   4875
   End
   Begin VB.Frame frmCT 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1275
      Left            =   11070
      TabIndex        =   71
      Top             =   7230
      Width           =   1215
      Begin VB.CommandButton cmdCT 
         Caption         =   "cmdQm"
         Height          =   345
         Left            =   150
         TabIndex        =   72
         Top             =   420
         Width           =   945
      End
      Begin VB.Label lblCCC 
         Caption         =   "产品采购"
         Height          =   225
         Left            =   240
         TabIndex        =   74
         Top             =   120
         Width           =   915
      End
      Begin VB.Label lblCT 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   150
         TabIndex        =   73
         Top             =   840
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdSave 
      Height          =   405
      Left            =   13800
      Picture         =   "frmGXBj.frx":05FC
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "保存"
      Top             =   8760
      Width           =   465
   End
   Begin VB.CommandButton cmdMod 
      Height          =   405
      Left            =   13320
      Picture         =   "frmGXBj.frx":0C66
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "修改"
      Top             =   8760
      Width           =   465
   End
   Begin VB.CommandButton cmdLeft 
      Caption         =   "<"
      Height          =   285
      Left            =   14280
      TabIndex        =   43
      Top             =   8250
      Width           =   465
   End
   Begin VB.CommandButton cmdRight 
      Caption         =   ">"
      Height          =   285
      Left            =   14760
      TabIndex        =   42
      Top             =   8250
      Width           =   465
   End
   Begin VB.TextBox txtHg 
      Height          =   270
      Left            =   10230
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   7020
      Width           =   4905
   End
   Begin VB.TextBox txtYhg 
      Height          =   270
      Left            =   13380
      Locked          =   -1  'True
      TabIndex        =   38
      ToolTipText     =   "此处由工程部填入"
      Top             =   7980
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frmHide 
      Caption         =   "frmHid"
      Height          =   1455
      Left            =   3090
      TabIndex        =   23
      Top             =   540
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Label LBLwhG 
         Height          =   255
         Left            =   1080
         TabIndex        =   93
         Top             =   1170
         Width           =   915
      End
      Begin VB.Label LBLyHG 
         Caption         =   "LBLyHG"
         Height          =   255
         Left            =   90
         TabIndex        =   92
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label LBLhG 
         Height          =   225
         Left            =   180
         TabIndex        =   91
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label lblQy 
         Caption         =   "lblQy"
         Height          =   255
         Left            =   2490
         TabIndex        =   61
         Top             =   180
         Width           =   1155
      End
      Begin VB.Label lblBm 
         Caption         =   "lblBm"
         Height          =   225
         Left            =   1020
         TabIndex        =   60
         Top             =   150
         Width           =   915
      End
      Begin VB.Label lblPwf 
         Caption         =   "lblPwf"
         Height          =   255
         Left            =   3510
         TabIndex        =   59
         Top             =   1110
         Width           =   1035
      End
      Begin VB.Label lblLc 
         Caption         =   "lblLc"
         Height          =   315
         Left            =   1260
         TabIndex        =   31
         Top             =   330
         Width           =   645
      End
      Begin VB.Label lblNlb 
         Caption         =   "lblNlb"
         Height          =   225
         Left            =   1920
         TabIndex        =   30
         Top             =   810
         Width           =   645
      End
      Begin VB.Label lblLcRen 
         Caption         =   "lblLcRen"
         Height          =   285
         Left            =   150
         TabIndex        =   29
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lblLcUid 
         Caption         =   "lblLcUid"
         Height          =   285
         Left            =   240
         TabIndex        =   28
         Top             =   600
         Width           =   885
      End
      Begin VB.Label lblFwid 
         Caption         =   "lblFwid"
         Height          =   255
         Left            =   1860
         TabIndex        =   27
         Top             =   450
         Width           =   1275
      End
      Begin VB.Label lblUid 
         Caption         =   "lblUid"
         Height          =   255
         Left            =   3750
         TabIndex        =   26
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblYwy 
         Caption         =   "lblYwy"
         Height          =   285
         Left            =   3540
         TabIndex        =   25
         Top             =   450
         Width           =   765
      End
      Begin VB.Label lblLcou 
         Caption         =   "lblLcou"
         Height          =   255
         Left            =   1860
         TabIndex        =   24
         Top             =   1080
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdBack 
      Height          =   405
      Left            =   14760
      Picture         =   "frmGXBj.frx":0F70
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "返回"
      Top             =   8760
      Width           =   465
   End
   Begin VB.Frame frmYw 
      Caption         =   "业务员填写"
      Height          =   2865
      Left            =   180
      TabIndex        =   3
      Top             =   5010
      Width           =   9135
      Begin VB.OptionButton OPTN 
         Caption         =   "杰升价"
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   1140
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.OptionButton OPTW 
         Caption         =   "自定价"
         Height          =   195
         Left            =   120
         TabIndex        =   94
         Top             =   1470
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton cmdQing 
         Caption         =   "清空"
         Height          =   375
         Left            =   8370
         TabIndex        =   58
         Top             =   1260
         Width           =   735
      End
      Begin MSDataListLib.DataCombo comJzxh 
         Height          =   330
         Left            =   6270
         TabIndex        =   48
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo comJzpb 
         Height          =   330
         Left            =   3000
         TabIndex        =   47
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.CommandButton cmdJgx 
         Caption         =   "更新"
         Height          =   315
         Left            =   8370
         TabIndex        =   35
         Top             =   2490
         Width           =   735
      End
      Begin VB.TextBox txtSl 
         Height          =   330
         Left            =   6270
         TabIndex        =   34
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除"
         Height          =   315
         Left            =   8370
         TabIndex        =   22
         Top             =   2100
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "添加"
         Height          =   345
         Left            =   8370
         TabIndex        =   21
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtCd 
         Height          =   330
         Left            =   6270
         TabIndex        =   19
         Top             =   1890
         Width           =   1935
      End
      Begin VB.TextBox txtLjmc 
         Height          =   330
         Left            =   6270
         TabIndex        =   18
         Top             =   1390
         Width           =   1935
      End
      Begin VB.TextBox txtCbh 
         Height          =   330
         Left            =   6270
         TabIndex        =   17
         Top             =   890
         Width           =   1935
      End
      Begin VB.TextBox txtLjbh 
         Height          =   270
         Left            =   3000
         TabIndex        =   16
         Top             =   1920
         Width           =   1905
      End
      Begin VB.TextBox txtXlh 
         Height          =   270
         Left            =   3000
         TabIndex        =   15
         Top             =   1410
         Width           =   1905
      End
      Begin VB.TextBox txtYxh 
         Height          =   270
         Left            =   3000
         TabIndex        =   14
         Top             =   900
         Width           =   1905
      End
      Begin VB.ComboBox comLx 
         Height          =   300
         ItemData        =   "frmGXBj.frx":1072
         Left            =   150
         List            =   "frmGXBj.frx":107C
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2640
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "查询比较"
         Height          =   255
         Left            =   180
         TabIndex        =   96
         Top             =   870
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblzlZ 
         Caption         =   "性质"
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   150
         TabIndex        =   64
         Top             =   390
         Width           =   435
      End
      Begin VB.Label lblZl 
         Caption         =   "Label19"
         ForeColor       =   &H00C000C0&
         Height          =   225
         Left            =   720
         TabIndex        =   63
         Top             =   390
         Width           =   945
      End
      Begin VB.Label lblLid 
         Caption         =   "lblLid"
         Height          =   315
         Left            =   390
         TabIndex        =   56
         Top             =   1860
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label11 
         Caption         =   "数   量"
         Height          =   255
         Left            =   5190
         TabIndex        =   33
         Top             =   2460
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "零件规格编号"
         Height          =   315
         Left            =   1830
         TabIndex        =   12
         Top             =   1950
         Width           =   1125
      End
      Begin VB.Label Label9 
         Caption         =   "品牌及产地"
         Height          =   315
         Left            =   5160
         TabIndex        =   11
         Top             =   1950
         Width           =   1035
      End
      Begin VB.Label Label8 
         Caption         =   "机组序列号"
         Height          =   315
         Left            =   1830
         TabIndex        =   10
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label Label7 
         Caption         =   "零件名称"
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   5160
         TabIndex        =   9
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "出厂编号"
         Height          =   315
         Left            =   5160
         TabIndex        =   8
         Top             =   930
         Width           =   1065
      End
      Begin VB.Label Label5 
         Caption         =   "机组品牌"
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1860
         TabIndex        =   7
         ToolTipText     =   "机组品牌"
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "压缩机型号"
         Height          =   315
         Left            =   1830
         TabIndex        =   6
         Top             =   930
         Width           =   1125
      End
      Begin VB.Label Label3 
         Caption         =   "机组型号"
         Height          =   315
         Left            =   5160
         TabIndex        =   5
         Top             =   420
         Width           =   1035
      End
      Begin VB.Label lblPz 
         Caption         =   "品种"
         Height          =   255
         Left            =   390
         TabIndex        =   4
         Top             =   2160
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin MSDataListLib.DataCombo comXmmc 
      Height          =   330
      Left            =   10260
      TabIndex        =   32
      Top             =   5220
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   582
      _Version        =   393216
      Text            =   ""
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgMa 
      Height          =   4695
      Left            =   -60
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   8281
      _Version        =   393216
      BackColor       =   12648384
      Cols            =   29
      BackColorFixed  =   12648384
      BackColorBkg    =   12648447
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      PictureType     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   29
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgNew 
      Height          =   4695
      Left            =   30
      TabIndex        =   139
      Top             =   -690
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   8281
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblSDJE 
      Caption         =   "Label23"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   13740
      TabIndex        =   135
      Top             =   4890
      Width           =   1395
   End
   Begin VB.Label Label22 
      Caption         =   "添加成本"
      Height          =   285
      Left            =   12810
      TabIndex        =   133
      ToolTipText     =   "此处由商务支持部添加"
      Top             =   7050
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblHLC 
      Caption         =   "lblHLC"
      Height          =   345
      Left            =   12150
      TabIndex        =   127
      Top             =   6930
      Visible         =   0   'False
      Width           =   1665
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
      ForeColor       =   &H00C00000&
      Height          =   345
      Left            =   9300
      TabIndex        =   126
      Top             =   7800
      Width           =   5475
   End
   Begin VB.Label lblCfwid 
      Caption         =   "lblCfwid"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   10200
      TabIndex        =   84
      Top             =   6510
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblZT 
      Caption         =   "Label21"
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
      Height          =   285
      Left            =   8970
      TabIndex        =   83
      Top             =   6420
      Width           =   1665
   End
   Begin VB.Label lblZ 
      Caption         =   "张春华"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   585
      Left            =   11520
      TabIndex        =   82
      Top             =   6930
      Width           =   1785
   End
   Begin VB.Label Label19 
      Caption         =   "备注"
      Height          =   225
      Left            =   9570
      TabIndex        =   76
      Top             =   5670
      Width           =   585
   End
   Begin VB.Label lblHtbh 
      Caption         =   "对应合同"
      Height          =   255
      Left            =   9210
      TabIndex        =   75
      Top             =   6000
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lblWbid 
      Caption         =   "lblWbid"
      Height          =   315
      Left            =   9690
      TabIndex        =   62
      Top             =   7860
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lblBh 
      BackColor       =   &H80000014&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      Height          =   285
      Left            =   10260
      TabIndex        =   55
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label14 
      Caption         =   "编号"
      Height          =   285
      Left            =   9570
      TabIndex        =   46
      Top             =   4950
      Width           =   435
   End
   Begin VB.Label Label13 
      Caption         =   "总费用"
      Height          =   225
      Left            =   9480
      TabIndex        =   41
      Top             =   7020
      Width           =   765
   End
   Begin VB.Label Label12 
      Caption         =   "优惠价"
      Height          =   255
      Left            =   12600
      TabIndex        =   40
      ToolTipText     =   "此处由工程部填入"
      Top             =   8010
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOid 
      Caption         =   "lblOid"
      Height          =   285
      Left            =   10650
      TabIndex        =   37
      Top             =   6750
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblBaoId 
      Caption         =   "lblBaoId"
      Height          =   285
      Left            =   9240
      TabIndex        =   36
      Top             =   6060
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblBid 
      Caption         =   "lblBid"
      Height          =   345
      Left            =   9330
      TabIndex        =   2
      Top             =   6022
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label24 
      Caption         =   "项目名称"
      Height          =   285
      Left            =   9210
      TabIndex        =   1
      Top             =   5280
      Width           =   795
   End
End
Attribute VB_Name = "frmGXBj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public adoPb As Object
Public adoJz As Object
Public adoXm As Object
Public adoOid As Object  '计算Old单子的ADO
Public adoGx As Object
Public CTF As Boolean '需要产品采购确认否(黄嘉琦,用于大修中的购销询价)
Public adoGy As Object '显示供应商的ado
Public FB As Boolean '是否为分包询价
Public ZF As Boolean '是否回驳过单据
Public JZ As Integer '基准价比例

Dim timZm As Integer '数据提交后,由timWait执行的后续命令ID(1决定何处购买,2单价更新,3配件添加,4配件更新5.供应商更新6配件删除7配件签字8删除)
Private Sub cmdAdd_Click()
On Error Resume Next
Dim hg As Long
'If comLx.Text = "" Then Exit Sub
If lblZl.Caption = "配件" Then
    comLx.Text = "零配件"
Else
    comLx.Text = "产品"
End If
''''''If Val(txtSl.Text) = 0 Then
''''''    MsgBox "请确认数量!"
''''''    txtSl.SetFocus
''''''    Exit Sub
''''''End If
If txtDRQ.Text = "" Then
    txtDRQ.Text = Date
End If
If txtBrq.Text = "" Then
    txtBrq.Text = Date
End If

''''''''''''If Val(lblHtbh.Caption) = 0 Then '老版本
''''''''''    Set mod1.cmd = createobject("adodb.command")
''''''''''        mod1.cmd.ActiveConnection = mod1.CC
''''''''''        mod1.cmd.CommandText = "gxAdd"
''''''''''        mod1.cmd.CommandType = adCmdStoredProc
''''''''''        mod1.cmd.Parameters("@pz") = Trim(comLx.Text)
''''''''''        mod1.cmd.Parameters("@jzpb") = Trim(comJzpb.Text)
''''''''''        mod1.cmd.Parameters("@jzxh") = Trim(comJzxh.Text)
''''''''''        mod1.cmd.Parameters("@yxh") = Trim(txtYxh.Text)
''''''''''        mod1.cmd.Parameters("@ccbh") = Trim(txtCbh.Text)
''''''''''        mod1.cmd.Parameters("@jzbh") = Trim(txtXlh.Text)
''''''''''        mod1.cmd.Parameters("@ljbh") = Trim(txtLjbh.Text)
''''''''''        mod1.cmd.Parameters("@ljmc") = Trim(txtLjmc.Text)
''''''''''        mod1.cmd.Parameters("@pbcd") = Trim(txtCd.Text)
''''''''''
''''''''''        mod1.cmd.Parameters("@sl") = Val(txtSl.Text)
''''''''''        mod1.cmd.Parameters("@mj") = Val(txtMj.Text)
''''''''''        mod1.cmd.Parameters("@dj") = Val(txtDj.Text)
''''''''''        mod1.cmd.Parameters("@hg") = Val(txtDj.Text) * Val(txtSl.Text)
''''''''''        mod1.cmd.Parameters("@drq") = txtDrq.Text
''''''''''        mod1.cmd.Parameters("@brq") = txtBrq.Text
''''''''''        mod1.cmd.Parameters("@bid") = Trim(lblBid.Caption)
''''''''''        mod1.cmd.Execute
''''''''''        Set cmd = Nothing
''''''''''''Else                                    '新版本
    timZm = 3
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
    mod1.cmd.Parameters("@mt2") = lblZl.Caption
    mod1.cmd.Parameters("@mt3") = comJzpb.Text '机组品牌
    mod1.cmd.Parameters("@mt4") = comJzXh.Text '机组型号
    mod1.cmd.Parameters("@mt5") = txtYxh.Text '压缩机型号
    mod1.cmd.Parameters("@mt6") = txtCbh.Text '出厂编号
    mod1.cmd.Parameters("@mt7") = txtXlh.Text '机组序列号
    mod1.cmd.Parameters("@mt8") = txtLjmc.Text '零件名称
    mod1.cmd.Parameters("@mt9") = txtLjbh.Text '零件规格号
    mod1.cmd.Parameters("@mt10") = txtCd.Text '品牌及产地
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
    mod1.cmd.Parameters("@mm1") = Val(txtSL.Text) '数量
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
    If mod1.Bm = "配送中心" Then
    mod1.cmd.Parameters("@mb5") = 1
    Else
    mod1.cmd.Parameters("@mb5") = 0
    End If
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





'''''''Hg = 0
'''''''adoGx.MoveFirst
'''''''Do While Not adoGx.EOF
'''''''    Hg = Hg + adoGx.Fields("合计").Value
'''''''    adoGx.MoveNext
'''''''Loop


'''''''''txtHg.Text = Hg
'''''''''txtYhg.Text = txtHg.Text
'comLx.Text = ""


End Sub

Private Sub cmdBack_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next

'frmGXBj.Visible = False
'If frmGxBiao.Visible = True Then
'    frmGxBiao.Enabled = True
'ElseIf Dialog.Visible = True Then
'    Dialog.Enabled = True
'End If

frmGxbjSD.Visible = False

frmGXBj.Visible = False
Call modBJD.BJDGXQing

If frmGxBiao.Visible = True Then
    frmGxBiao.Enabled = True
    frmGxBiao.ZOrder 0
ElseIf FMXC.Visible = True Then

    FMXC.Enabled = True
    FMXC.ZOrder 0
'''''    FMXC.cmdW5.Enabled = True
'''''    FMXC.cmdW6.Enabled = True
ElseIf Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0

End If


End Sub

Private Sub cmdBJ_Click()

End Sub

Private Sub cmdBjd_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
Call modBJD.BaoJDWBQing

If Val(lblBaoId.Caption) = 0 Then
    MsgBox ("这是旧版询价单才有的功能!")
    Exit Sub
    ii = MsgBox("是否生成新报价单?", vbQuestion + vbYesNo, "您辛苦了!")
    If ii = vbNo Then
        Exit Sub
    End If
    If cmdRight.Enabled = True Then
        MsgBox "当前记录不是最终有效询价单,故不能生成新报价单"
        Exit Sub ''如果不是最终有效询价单,则不能生成新报价单
    End If
    If lblYwy.Caption <> mod1.DName Then
        MsgBox "必须由业务员亲自生成报价单!"
        Exit Sub
    End If

    frmGxbjB.Visible = False
    mod1.BTZ = 37
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "BJDadd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@xmmc") = Trim(comXmmc.Text)
    mod1.cmd.Parameters("@xid") = comXmmc.Tag
    mod1.cmd.Parameters("@ywy") = Trim(lblYwy.Caption)
    mod1.cmd.Parameters("@uid") = Trim(lblUid.Caption)
    mod1.cmd.Parameters("@lx") = 0
    mod1.cmd.Parameters("@zh") = 0
    mod1.cmd.Parameters("@zName") = ""
'    mod1.CMD.Parameters("@jzPb") = ""
'    mod1.CMD.Parameters("@jzxh") = ""
'    mod1.CMD.Parameters("@sl") = 0
    mod1.cmd.Parameters("@ta") = 0
    mod1.cmd.Parameters("@tb") = 0
    mod1.cmd.Parameters("@tc") = 0
    mod1.cmd.Parameters("@ztime") = 0
    mod1.cmd.Parameters("@yhg") = Val(txtYhg.Text)
    mod1.cmd.Parameters("@nlb") = 60
    mod1.cmd.Parameters("@lcou") = 3
    mod1.cmd.Parameters("@bid") = Val(lblBid.Caption)
    mod1.cmd.Parameters("@clcb") = 0
    mod1.cmd.Parameters("@zl") = lblZl.Caption
    mod1.cmd.Parameters("@clf") = 0
    mod1.cmd.Parameters("@rgf") = 0
    mod1.cmd.Parameters("@mon") = 0
    mod1.cmd.Parameters("@dxnr") = ""
    mod1.cmd.Parameters("@wc") = 0
    mod1.cmd.Parameters("@xc") = 0
    mod1.cmd.Parameters("@cgid") = 0
    mod1.cmd.Parameters("@bz") = Trim(txtBz.Text)
    mod1.cmd.Parameters("@fbje") = 0
    mod1.cmd.Parameters("@fbnr") = ""
    mod1.cmd.Parameters("@yf") = Val(txtYf.Text)

    'mod1.CMD.Parameters("
    mod1.cmd.Execute

    lblBaoId.Caption = mod1.cmd.Parameters("@baoid").Value
    frmGxbjB.lblBaoId.Caption = mod1.cmd.Parameters("@baoid").Value
    Set cmd = Nothing
    Call modBJD.BaoJDGXQing
    Call modBJD.BaoJDBound(Val(lblBaoId.Caption), "购销")



    tt = "select * from baojiaOld where old=" & Val(frmGxbjB.lblOid.Caption) & " order by baoid"
    frmGxbjB.adoOid.Close
    frmGxbjB.adoOid.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    If frmGxbjB.adoOid.RecordCount > 1 Then
        frmGxbjB.cmdLeft.Enabled = True
    End If
    frmGxbjB.adoOid.MoveLast
        '设置流程按钮
    Call modBJD.BjGXLcBut(60)


    frmGxbjB.Visible = True
    frmGxbjB.cmdPrint.Visible = False
    Call modBJD.gxbjbLocked
    frmGxbjB.txtSL.Locked = False
    frmGxbjB.comKhmc.Locked = False
    frmGxbjB.txtYf.Locked = False
    frmGxbjB.txtXm2.Locked = False
    frmGxbjB.txtHg.Locked = False
    frmGxbjB.txtYhg.Locked = False
    frmGxbjB.cmdGx.Enabled = True
    frmGxbjB.txtDj.Locked = False
    frmGxbjB.cmdMod.Enabled = False
    frmGxbjB.cmdSave.Enabled = True
    frmGxbjB.lblZl.Caption = "购销"
    frmGxbjB.txtCb.Text = txtYhg.Text
Else
    mod1.BTZ = 37
    Call modBJD.BaoJDGXQing
    Call modBJD.BaoJDBound(Val(lblBaoId.Caption), "购销")
    frmGxbjB.Visible = True
    Call modBJD.gxbjbLocked
    frmGxbjB.cmdSave.Enabled = False
    frmGxbjB.cmdMod.Enabled = True
End If

frmGxbjB.optLa.Enabled = True
frmGxbjB.optLb.Enabled = True
frmGxbjB.optLc.Enabled = True



frmGXBj.Visible = False
End Sub

Private Sub cmdCong_Click()
Dim ii As Integer
Dim oo As Integer
Dim tt As String
Dim Bid As Long
Dim ZL As String
On Error Resume Next
'MsgBox "正在建设中!"
'Exit Sub
'If Val(lblBaoId.Caption) > 0 Then
'    Exit Sub
'End If
ii = MsgBox("您的这项操作将使原先单子正在执行的流程全部撤消,是否确定执行?", vbYesNo + vbInformation, "询问")
If ii = vbYes Then
    tt = InputBox("请输入您要驳回的原因!")
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "xtzxFAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@yid").Value = 43 '反签名
    mod1.cmd.Parameters("@lc").Value = 2 '退回最初的流程
    mod1.cmd.Parameters("@bh").Value = lblBid.Caption
    mod1.cmd.Parameters("@ywy").Value = mod1.DName
    mod1.cmd.Parameters("@uid").Value = mod1.DHid
    mod1.cmd.Parameters("@bz").Value = tt
    mod1.cmd.Parameters("@zn").Value = "new" '身份职能
    mod1.cmd.Execute
    If Left(mod1.cmd.Parameters("@jch").Value, 6) = "合同已经生效" Then
        MsgBox mod1.cmd.Parameters("@jch").Value
        Set cmd = Nothing
        Exit Sub
    End If
    Set cmd = Nothing
'''''    For oo = 0 To 5
'''''        cmdQm(oo).Caption = ""
'''''        lblTm(oo).Caption = ""
'''''    Next
    lblLc.Caption = 999 '不让再按签名按钮.
    If Dialog.Visible = True Then '更新事务列表
        Call mod1.refEnvent(1)
    End If
    cmdBjd.Visible = False
    Exit Sub
ElseIf ii = vbCancel Then
    Exit Sub
End If


End Sub

Private Sub cmdD_Click()
Dim ii As Integer
Dim tt As String
On Error Resume Next
tt = "select htbh from htping where hid=" & Val(lblHtbh.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
If mod1.HTP.Fields(0).Value <> "HMNEW" And mod1.DName <> "马晓聪" Then
    Exit Sub
End If
If lblYwy.Caption <> mod1.DName And mod1.DName <> "马晓聪" Then Exit Sub

ii = MsgBox("是否删除此询价单？", vbYesNo + vbQuestion, "Hello")
If ii = vbNo Then
    Exit Sub
End If
timZm = 8 '删除合同
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

Private Sub cmdDao_Click()
'Set frmLingjian.LpXh = CreateObject("adodb.recordset")
Dim tt As String
Dim oo As Integer: Dim ii As Integer
Dim Ra, La
On Error Resume Next
tt = "SELECT top 100 dbo.l_goods.code, dbo.l_goods.name, dbo.l_goods.specs, dbo.l_goodstype.name AS goodtypename, dbo.l_goodsunit.unitname,dbo.l_goods.goodsid" & _
    " FROM dbo.l_goods LEFT OUTER JOIN dbo.l_goodsunit ON dbo.l_goods.goodsid = dbo.l_goodsunit.goodsid LEFT OUTER JOIN dbo.l_goodstype ON dbo.l_goods.gdtypeid = dbo.l_goodstype.gdtypeid where  dbo.l_goods.closed=0"
'''''Set mod1.HTP = CreateObject("adodb.recordset")
'''''mod1.HTP.Open tt, mod1.workSD, adOpenForwardOnly, adLockReadOnly, adCmdText
'''''Ra = mod1.HTP.GetRows
'''''La = UBound(Ra, 2)
'''''frmGxbjSD.dtgHP.Rows = La + 10
'''''mod1.HTP.Close
'''''Set mod1.HTP = Nothing
'''''On Error GoTo GXBJERR
'''''For oo = 1 To La + 1
''''''''        If oo = 50 Then
''''''''            ii = ii
''''''''        End If
'''''    frmGxbjSD.dtgHP.Row = oo
'''''    For ii = 1 To 6
'''''        frmGxbjSD.dtgHP.Col = ii
'''''
'''''        If IsNull(Ra(ii - 1, oo - 1)) = False Then
'''''            frmGxbjSD.dtgHP.Text = Ra(ii - 1, oo - 1)
'''''        End If
'''''    Next
'''''Next
Call frmGxbjSD.dtgFF
Call frmGxbjSD.CX(tt)

frmGxbjSD.Show
frmGxbjSD.ZOrder 0
''''''''''''If comJzpb.Text <> "" Then
''''''''''''    frmLingjian.Caption = comJzpb.Text
''''''''''''    frmLingjian.Show
''''''''''''
''''''''''''    For oo = frmLingjian.comJzxh.ListCount - 1 To 0 Step -1
''''''''''''        frmLingjian.comJzxh.RemoveItem oo
''''''''''''    Next
''''''''''''
''''''''''''    tt = "LPG_jzXhP('" & frmLingjian.Caption & "')"
''''''''''''    frmLingjian.LpXh.Close
''''''''''''    frmLingjian.LpXh.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
''''''''''''    If comJzpb.Text = "开利" Or comJzpb.Text = "约克" Then
''''''''''''        frmLingjian.dtgView.Columns(5).Caption = "伏斯价"
''''''''''''
''''''''''''            frmLingjian.dtgView.Columns("伏斯价").Visible = False
''''''''''''            If mod1.DName = "张春华" Or mod1.DName = "邹晨" Or mod1.DName = "" Then
''''''''''''                frmLingjian.dtgView.Columns("伏斯价").Visible = True
''''''''''''            End If
''''''''''''
''''''''''''    Else
''''''''''''        frmLingjian.dtgView.Columns(5).Caption = "库存价"
''''''''''''        frmLingjian.dtgView.Columns("库存价").Visible = False
''''''''''''        If mod1.DName = "张春华" Or mod1.DName = "邹晨" Or mod1.DName = "" Then
''''''''''''            frmLingjian.dtgView.Columns("库存价").Visible = True
''''''''''''        End If
''''''''''''    End If
''''''''''''    Set frmLingjian.dtgView.DataSource = Nothing
''''''''''''    cmdGx.Enabled = False
''''''''''''ElseIf txtLjmc.Text = 1 Then
''''''''''''    frmLingjian.Caption = "制冷剂"
''''''''''''    frmLingjian.Show
''''''''''''
''''''''''''    For oo = frmLingjian.comJzxh.ListCount - 1 To 0 Step -1
''''''''''''        frmLingjian.comJzxh.RemoveItem oo
''''''''''''    Next
''''''''''''
''''''''''''    tt = "LPG_jzXhP('" & frmLingjian.Caption & "')"
''''''''''''    frmLingjian.LpXh.Close
''''''''''''    frmLingjian.LpXh.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
''''''''''''        frmLingjian.dtgView.Columns(5).Caption = "库存价"
''''''''''''
''''''''''''            frmLingjian.dtgView.Columns("库存价").Visible = False
''''''''''''            If mod1.DName = "张春华" Or mod1.DName = "邹晨" Or mod1.DName = "" Then
''''''''''''                frmLingjian.dtgView.Columns("库存价").Visible = True
''''''''''''            End If
''''''''''''
''''''''''''    Set frmLingjian.dtgView.DataSource = Nothing
''''''''''''    cmdGx.Enabled = False
''''''''''''End If
Exit Sub
GXBJERR:
MsgBox "ok" & oo

End Sub

Private Sub cmdDel_Click()
Dim ii As Integer
On Error Resume Next
If mod1.VLP = 2 Or mod1.VLP = 3 And mod1.DName <> "马晓聪" Then
    MsgBox "You are a Pig!"
    End
End If
ii = MsgBox("是否删除此条记录?", vbInformation + vbYesNo, "询问")
If ii = vbYes Then
'''''    Set mod1.cmd = createobject("adodb.command")
'''''        mod1.cmd.ActiveConnection = mod1.CC
'''''        mod1.cmd.CommandText = "gxDel"
'''''        mod1.cmd.CommandType = adCmdStoredProc
'''''        mod1.cmd.Parameters("@lid") = Val(lblLId.Caption)
'''''        mod1.cmd.Execute
'''''    Set cmd = Nothing
    
     timZm = 6
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
    mod1.cmd.Parameters("@mt2") = lblZl.Caption
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
   
    
    
    
    
    
    
    

''''' adoGx.Requery
'''''Set dtgMa.DataSource = adoGx
'''''If adoGx.RecordCount > 0 Then
'''''    dtgMa.FixedRows = 0
'''''    dtgMa.MergeCol(1) = True
'''''    dtgMa.MergeCol(2) = True
'''''    dtgMa.MergeCol(10) = True
'''''    dtgMa.MergeCol(14) = True
'''''    dtgMa.MergeCells = 3
'''''    dtgMa.FixedRows = 1
'''''End If
'''''
''''''comLx.Text = ""
'''''comJzpb.Text = ""
'''''comJzxh.Text = ""
'''''txtYxh.Text = ""
'''''txtCbh.Text = ""
'''''txtXlh.Text = ""
'''''txtLjbh.Text = ""
'''''txtLjmc.Text = ""
'''''txtCd.Text = ""
'''''txtDrq.Text = ""
'''''txtSl.Text = ""
End If
End Sub

Private Sub cmdDing_Click()
Dim tt As String
On Error Resume Next

If OptT1.Value = True And Val(lblLc.Caption) = 4 Then '业务员确认时,检查是否选择供应商
    dtgN.Col = 23
    For oo = 1 To adoGx.RecordCount
        dtgN.Row = oo
        If dtgN.Text = "" Then
            frmQm.Visible = False
            MsgBox "请选确认供应商!"
            Exit Sub
        End If

    Next
End If
If optT2.Value = True And txtQM.Text = "" Then
    MsgBox ("请您一定要告诉拒绝我的理由!  :) ")
    Exit Sub
End If
timZm = 7 '配件签字
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单"
    mod1.cmd.Parameters("@NBLX") = "配件签字"
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



Private Sub cmdExcel_Click()
'If mod1.DName <> "马晓聪" Then Exit Sub
dtgMa.FixedRows = 0
dtgMa.Col = 0
dtgMa.Row = 0
dtgMa.ColSel = 11
    dtgMa.RowSel = adoGx.RecordCount

Clipboard.Clear
Clipboard.SetText dtgMa.Clip


dtgMa.FixedRows = 1
End Sub

Private Sub cmdGAdd_Click()
Call GyQing

End Sub

Private Sub cmdGB_Click()
frmWai.Visible = False
End Sub

Private Sub cmdGsave_Click()
On Error Resume Next
If Val(txtGyid.Text) = 1 Then Exit Sub '不能修改零件事业部资料

''''''''''Set mod1.cmd = createobject("adodb.command")
''''''''''    mod1.cmd.ActiveConnection = mod1.workKK
''''''''''    mod1.cmd.CommandText = "GYUpdate"
''''''''''    mod1.cmd.CommandType = adCmdStoredProc
''''''''''    mod1.cmd.Parameters("@gyid") = Val(txtGyid.Text)
''''''''''    mod1.cmd.Parameters("@gymc") = Trim(txtGYmc.Text)
''''''''''    mod1.cmd.Parameters("@gyman") = Trim(txtGyman.Text)
''''''''''    mod1.cmd.Parameters("@gyadr") = Trim(txtGyAdr.Text)
''''''''''    mod1.cmd.Parameters("@gyPho") = Trim(txtGYPho.Text)
''''''''''    mod1.cmd.Parameters("@ywy") = mod1.DName
''''''''''    mod1.cmd.Parameters("@uid") = mod1.DHid
''''''''''    mod1.cmd.Parameters("@gyBz") = Trim(txtGybz.Text)
''''''''''    mod1.cmd.Parameters("@errch") = ""
''''''''''    mod1.cmd.Parameters("@lid") = Val(lblLid.Caption)
''''''''''    mod1.cmd.Parameters("@bid") = Val(lblBid.Caption)
''''''''''    mod1.cmd.Parameters("@hg") = 0
''''''''''    mod1.cmd.Execute
''''''''''    txtHg.Text = mod1.cmd.Parameters("@hg").Value
''''''''''    txtYhg.Text = mod1.cmd.Parameters("@hg").Value
''''''''''Set cmd = Nothing
''''''''''adoGx.Requery
''''''''''    Set dtgMa.DataSource = adoGx
    timZm = 5
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.workKK
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单"
    mod1.cmd.Parameters("@NBLX") = "供应商更新"
    mod1.cmd.Parameters("@bh") = lblHtbh.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = lblBid.Caption
    mod1.cmd.Parameters("@mt2") = lblZl.Caption
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = txtGymc.Text
    mod1.cmd.Parameters("@mt5") = txtGyman.Text
    mod1.cmd.Parameters("@mt6") = txtGyAdr.Text
    mod1.cmd.Parameters("@mt7") = txtGYPho.Text
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
    mod1.cmd.Parameters("@mlt1") = txtGybz.Text
    mod1.cmd.Parameters("@mlt2") = ""
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtGyid.Text)
    mod1.cmd.Parameters("@mm2") = Val(lblLid.Caption)
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
        cmdAdd.Enabled = False
        cmdJG.Enabled = False
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If
    cmdGsave.Enabled = False

    Set mod1.cmd = Nothing

End Sub

Private Sub cmdGx_Click()
On Error Resume Next
Dim tt As String
Dim CHD As Object
Dim hg As Long
Dim Err As String
If lblLc.Caption = 1 Then Exit Sub
'''''If Val(lblLid.Caption) = 0 And mod1.DName <> "马晓聪" Then Exit Sub
'''''If mod1.BM = "零件事业部" Or mod1.BM = "行政人事" Or mod1.DName = "" Or mod1.DName = "杨燕" Or mod1.DName = "马晓聪" Or mod1.BM = "技术部" Then
'''''    If (Val(txtDj.Text) = 0 Or Val(txtMj.Text) = 0 Or txtDrq.Text = "") And mod1.DName <> "马晓聪" Then Exit Sub
'''''Else '业务员填写外部价
'''''    If lblYwy.Caption <> mod1.DName Then Exit Sub
'''''    If Val(txtDj.Text) = 0 Then
'''''    MsgBox ("请确认单价!")
'''''    txtDj.SetFocus
'''''    Exit Sub
'''''    End If
'''''    If txtGM.Text = "" Or Val(txtGM.ToolTipText) = 0 Then
'''''        MsgBox ("请确认供应商!")
'''''    Exit Sub
'''''    End If
'''''End If

If mod1.DName = "" Or Ywy = "吴金荣" Or mod1.DName = "杨燕" Then
    frmJ.Visible = False
    If Val(txtJdj.Text) = 0 Then
        MsgBox "请键入基准价"
        frmJ.Visible = True
        txtJdj.SetFocus
        Exit Sub
    End If

End If

'只有当前合同执行人可以修改成本单价供应商
If (FMXC.lblLcRen = mod1.DName) And FMXC.Visible = True Then

tt = ""
End If

    timZm = 2
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单"
    mod1.cmd.Parameters("@NBLX") = "单价更新"
    mod1.cmd.Parameters("@bh") = lblBid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = lblZl.Caption
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = lblZl.Caption
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
    mod1.cmd.Parameters("@mt8") = ""
    mod1.cmd.Parameters("@mt9") = ""
    mod1.cmd.Parameters("@mt10") = ""
    mod1.cmd.Parameters("@mt11") = txtZBQ.Text
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
    mod1.cmd.Parameters("@mt23") = txtDRQ.Text
    mod1.cmd.Parameters("@mt24") = ""
    mod1.cmd.Parameters("@mt25") = lblZl.Caption
    mod1.cmd.Parameters("@mlt1") = txtGybz.Text
    mod1.cmd.Parameters("@mlt2") = ""
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtMj.Text)
    mod1.cmd.Parameters("@mm2") = Val(txtDj.Text)
    mod1.cmd.Parameters("@mm3") = Val(txtYf.Text)
    mod1.cmd.Parameters("@mm4") = 0
    mod1.cmd.Parameters("@mm5") = Val(lblLid.Caption)
    mod1.cmd.Parameters("@mm6") = Val(txtGM.ToolTipText)
    mod1.cmd.Parameters("@mm7") = Val(txtJdj.Text)
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
    mod1.cmd.Parameters("@mm18") = Val(lblHtbh.Caption)
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = 0
    If mod1.Bm = "市场营销部" Or (mod1.Bm = "北京配送中心" And Val(lblLc.Caption) = 2) Or mod1.Bm = "技术部" Then    '判断是更新内部价还是外部价
        mod1.cmd.Parameters("@mb1") = 0
    Else
        mod1.cmd.Parameters("@mb1") = 1
    End If
    mod1.cmd.Parameters("@mb2") = 0
    mod1.cmd.Parameters("@mb3") = 0
    mod1.cmd.Parameters("@mb4") = 0
    mod1.cmd.Parameters("@mb5") = 0
    'If (mod1.BM = "零件事业部" Or mod1.BM = "配送中心" Or mod1.BM = "行政人事") And mod1.DName <> "马晓聪" Or (mod1.DName = "徐瑛" And Val(lblLc.Caption) >= 2) Then
    If mod1.Bm = "市场营销部" Or mod1.Bm = "行政人事" Or mod1.Bm = "技术部" Or (mod1.Bm = "北京配送中心" And Val(lblLc.Caption) >= 2) Then
        mod1.cmd.Parameters("@md1") = Null
        mod1.cmd.Parameters("@md2") = txtBrq.Text
    Else
        mod1.cmd.Parameters("@md1") = Null
        mod1.cmd.Parameters("@md2") = Null
    End If
    mod1.cmd.Parameters("@md3") = Null
    mod1.cmd.Parameters("@md4") = Null
    mod1.cmd.Parameters("@md5") = Null
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
     mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        If timZm = 1 Then '决定从何处购买
            cmdJG.Enabled = False
        End If
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        cmdJG.Enabled = False
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


    Set mod1.cmd = Nothing

'End If

comLx.Text = ""
comJzpb.Text = ""
comJzXh.Text = ""
txtYxh.Text = ""
txtCbh.Text = ""
txtXlh.Text = ""
txtLjbh.Text = ""
txtLjmc.Text = ""
txtCd.Text = ""
txtDRQ.Text = ""
txtSL.Text = ""
txtDj.Text = ""
txtBrq.Text = ""
txtMj.Text = ""
cmdSave.Enabled = True
If lblLc.Caption < 2 Then
    cmdGx.Enabled = False
End If
End Sub

Private Sub cmdGy_Click()
Dim tt As String
On Error Resume Next
If mod1.Bm = "零件事业部" Then
    Exit Sub
End If
txtGyid.Text = ""
txtGymc.Text = ""
txtGyman.Text = ""
txtGyAdr.Text = ""
txtGYPho.Text = ""

If Val(cmdGy.ToolTipText) > 0 Then
tt = "select * from gynew where gyid=" & cmdGy.ToolTipText
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
txtGyid.Text = mod1.HTP.Fields("gyid").Value
txtGymc.Text = mod1.HTP.Fields("gymc").Value
txtGyman.Text = mod1.HTP.Fields("gyman").Value
txtGyAdr.Text = mod1.HTP.Fields("gyadr").Value
txtGYPho.Text = mod1.HTP.Fields("gypho").Value
End If
frmWai.Visible = True
dtgGy.Visible = False

End Sub

Private Sub cmdGyOpen_Click()
Dim tt As String
On Error Resume Next
'''''If mod1.BM = "零件事业部" Then Exit Sub
tt = "select gyid,gymc from gyNew where uid='" & mod1.DHid & "' or gyid=1"
adoGy.Close
adoGy.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set dtgGM.DataSource = adoGy
dtgGM.Visible = True

End Sub

Private Sub cmdHt_Click()
Dim Ra
Dim tt As String
tt = "select newf from htping where hid=" & Val(lblHtbh.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
If Ra(0, 0) = 6 Then
    Call FmxcNew.Bound(Val(lblHtbh.Caption))
    FmxcNew.Show
    FmxcNew.ZOrder 0
    Me.Visible = False
    Exit Sub
End If
If mod1.Bm = "零件事业部" Then
    MsgBox "哈哈！"
    MsgBox "你想干嘛？"
    Exit Sub
End If
mod1.BTZ = 6

If FMXC.Visible = True And Val(FMXC.lblMHid.Caption) = Val(lblHtbh.Caption) Then
    Me.Visible = False
    FMXC.Enabled = True
    FMXC.ZOrder 0
ElseIf Val(lblHtbh.Caption) < 19345 Then

        Call modNewHT.NewMQing
        
        Call modNewHT.NewMBound(Val(lblHtbh.Caption))
        If FMXC.Visible = True Then '如果打开成功,则隐藏自己.
            Me.Visible = False
            FMXC.ZOrder 0
        End If
Else
        Call modNewHT.NewMQing
        
        Call modNewHT.NewB(Val(lblHtbh.Caption))
        If FMXC.Visible = True Then '如果打开成功,则隐藏自己.
            Me.Visible = False
            FMXC.ZOrder 0
        End If
End If
    FMXC.cmdMQm(0).Visible = True
    FMXC.lblMQM(0).Visible = True
    FMXC.lblMTm(0).Visible = True
End Sub

Private Sub cmdJG_Click()
timZm = 1
If mod1.DName <> "张春华" Then
    Exit Sub
End If
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单"
    mod1.cmd.Parameters("@NBLX") = "选购决定"
    mod1.cmd.Parameters("@bh") = lblBid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = lblHtbh.Caption
    mod1.cmd.Parameters("@mt2") = lblZl.Caption
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
    mod1.cmd.Parameters("@mm1") = Val(txtYhg.Text)
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
    If opt1.Value = True Then              '决定从何购买,同时影响合同评审成本
        mod1.cmd.Parameters("@md1") = 0
    ElseIf opt2.Value = True Then
        mod1.cmd.Parameters("@md1") = 1
    End If
    mod1.cmd.Parameters("@md2") = Null
    mod1.cmd.Parameters("@md3") = Null
    mod1.cmd.Parameters("@md4") = Null
    mod1.cmd.Parameters("@md5") = Null
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        If timZm = 1 Then '决定从何处购买
            cmdJG.Enabled = False
        End If
        Exit Sub
    Else '提交成功,等待系统中心处理数据
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

Private Sub cmdJgx_Click()
On Error Resume Next

''''''''Set mod1.cmd = createobject("adodb.command")
''''''''    mod1.cmd.ActiveConnection = mod1.CC
''''''''    mod1.cmd.CommandText = "gxUpdate"
''''''''    mod1.cmd.CommandType = adCmdStoredProc
''''''''    mod1.cmd.Parameters("@pz") = Trim(comLx.Text)
''''''''    mod1.cmd.Parameters("@jzpb") = Trim(comJzpb.Text)
''''''''    mod1.cmd.Parameters("@jzxh") = Trim(comJzxh.Text)
''''''''    mod1.cmd.Parameters("@yxh") = Trim(txtYxh.Text)
''''''''    mod1.cmd.Parameters("@ccbh") = Trim(txtCbh.Text)
''''''''    mod1.cmd.Parameters("@jzbh") = Trim(txtXlh.Text)
''''''''    mod1.cmd.Parameters("@ljbh") = Trim(txtLjbh.Text)
''''''''    mod1.cmd.Parameters("@ljmc") = Trim(txtLjmc.Text)
''''''''    mod1.cmd.Parameters("@pbcd") = Trim(txtCd.Text)
''''''''
''''''''    mod1.cmd.Parameters("@sl") = Val(txtSl.Text)
''''''''    'mod1.CMD.Parameters("@dj") = Val(txtDj.Text)
''''''''    'mod1.CMD.Parameters("@hg") = Val(txtDj.Text) * Val(txtSl.Text)
''''''''    'mod1.CMD.Parameters("@brq") = txtBrq.Text
''''''''    mod1.cmd.Parameters("@bid") = Val(lblBid.Caption)
''''''''    mod1.cmd.Parameters("@lid") = Val(lblLid.Caption)
''''''''    mod1.cmd.Execute
''''''''Set cmd = Nothing
''''''If Val(txtSl.Text) = 0 Then
''''''    Exit Sub
''''''End If



    timZm = 4
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
    mod1.cmd.Parameters("@mt2") = lblZl.Caption
    mod1.cmd.Parameters("@mt3") = comJzpb.Text '机组品牌
    mod1.cmd.Parameters("@mt4") = comJzXh.Text '机组型号
    mod1.cmd.Parameters("@mt5") = txtYxh.Text '压缩机型号
    mod1.cmd.Parameters("@mt6") = txtCbh.Text '出厂编号
    mod1.cmd.Parameters("@mt7") = txtXlh.Text '机组序列号
    mod1.cmd.Parameters("@mt8") = txtLjmc.Text '零件名称
    mod1.cmd.Parameters("@mt9") = txtLjbh.Text '零件规格号
    mod1.cmd.Parameters("@mt10") = txtCd.Text '品牌及产地
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
    mod1.cmd.Parameters("@mm1") = Val(txtSL.Text) '数量
    mod1.cmd.Parameters("@mm2") = Val(lblLid.Caption)
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

Private Sub cmdLeft_Click()
Dim tt As String
Dim ZL As String
On Error Resume Next
If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存!"
    Exit Sub
End If
Me.Enabled = False
frmWait.Show
frmWait.ZOrder
frmWait.Refresh
frmGXBj.adoOid.MovePrevious
ZL = lblZl.Caption
'打开新建单
Call modBJD.BJDGXQing
Call modBJD.BJDGDBound(frmGXBj.adoOid.Fields("bid").Value)
Call modBJD.gxbjLocked
frmGXBj.cmdRight.Enabled = True
frmGXBj.cmdBjd.Visible = False
'frmGXBj.cmdCong.Visible = False
frmGXBj.cmdWb.Visible = False
cmdMod.Enabled = False
cmdSave.Enabled = False
frmGXBj.lblZl.ForeColor = &H80000012
frmGXBj.lblzlZ.ForeColor = &H80000012
frmGXBj.adoOid.MovePrevious
If frmGXBj.adoOid.BOF = True Then
    cmdLeft.Enabled = False
Else
    cmdLeft.Enabled = True
End If
frmGXBj.adoOid.MoveNext
frmWait.Visible = False
Me.Enabled = True
Me.ZOrder 0
End Sub

Private Sub cmdMod_Click()
Dim tt As String
Dim Lc As Integer
Dim HTZE As Single
Dim HtF As Integer
On Error Resume Next
'如果合同已经盖章执行，则不能修改询价单
If mod1.DName = "马晓聪" Then
        frmZ.Visible = True
        

        comXmmc.Locked = True
        
        txtYhg.Locked = False
        txtMj.Locked = False
        cmdGx.Enabled = True
        frmCg.Enabled = True
        txtDj.Locked = False
        Call modBJD.gxbjUnLocked
        txtYf.Locked = False
        txtADR.Locked = False
        cmdSave.Enabled = True
        frmSd.Visible = True
        Exit Sub
End If
If mod1.Bm <> "市场营销部" And mod1.Bm <> "配送中心" Then
    tt = "select LC,htze,htF from htping where hid=" & Val(lblHtbh.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    If IsNull(mod1.HTP.RecordCount) = True Then
        MsgBox ("出错，请关闭程序再试一次!")
        Exit Sub
    End If
    Lc = mod1.HTP.Fields("lc").Value
    HTZE = mod1.HTP.Fields("htze").Value
    HtF = mod1.HTP.Fields("htf").Value
    If Lc > 1 And HTZE >= 15000 Or HTZE < 15000 And (HtF = 1 Or HtF = 2) Then
       MsgBox ("您的合同已经签字，不能修改成本！")
       Exit Sub
    Else
        Call modBJD.gxbjUnLocked
        cmdSave.Enabled = True
        cmdGx.Enabled = True
        
    End If
End If

If mod1.Bm = "零件事业部" Then
    lblLcRen.Caption = mod1.DName
    lblLcUid.Caption = mod1.DHid
End If
If mod1.DName = "张春华" And (lblLcRen.Caption = "邹晨" Or mod1.DName = "" Or Ywy = "吴金荣") Then
    lblLcRen.Caption = "张春华"
    lblLcUid.Caption = "HM001"
End If
Call modBJD.gxbjLocked

If mod1.DName = lblLcRen.Caption Or mod1.Bm = "配送中心" Or lblLc.Caption = 100 Then

    


    frmSd.Visible = True
    Call modBJD.gxbjUnLocked
    frmCg.Enabled = True
    comXmmc.Locked = True
    cmdSave.Enabled = True
    cmdD.Enabled = True
    cmdGsave.Enabled = True
    If txtGM.Text = "零件事业部" Or mod1.DName = "" Or Ywy = "吴金荣" Then
    txtDj.Locked = True
    Else
    txtDj.Locked = False
    End If
End If

If (lblLc.Caption = 2) And lblLcRen.Caption = mod1.DName Then
        If mod1.DName = "张春华" Or mod1.DName = "邹晨" Or mod1.DName = "" Or Ywy = "吴金荣" Then
            frmZ.Visible = True
        End If
    'If mod1.VLP = 2 Or mod1.VLP = 3 Then '采购可以改优惠价,业务员则不行.
        Call modBJD.gxbjUnLocked
        comXmmc.Locked = True
        
        txtYhg.Locked = False
        txtMj.Locked = False
        cmdGx.Enabled = True
        frmCg.Enabled = True
        txtDj.Locked = False
        Call modBJD.gxbjUnLocked
        comXmmc.Locked = True
        cmdSave.Enabled = True
    'End If
End If

If (mod1.Bm = "零件事业部") And Val(lblLc.Caption) > 1 And Val(lblLc.Caption) = 2 Then

        frmZ.Visible = True
        

        Call modBJD.gxbjUnLocked
        comXmmc.Locked = True
        
        txtYhg.Locked = False
        txtMj.Locked = False
        cmdGx.Enabled = True
        frmCg.Enabled = True
        txtDj.Locked = False
        Call modBJD.gxbjUnLocked
        txtYf.Locked = False
        txtADR.Locked = False
        comXmmc.Locked = True
        cmdSave.Enabled = True
End If

End Sub


Private Sub cmdNDel_Click()
Call cmdDel_Click
End Sub

Private Sub cmdNGx_Click()
'''''If Val(txtNsl.Text) = 0 Then
'''''    MsgBox "请键入数量"
'''''    txtNsl.SetFocus
'''''End If
If comJzpb.Text = "" Then
    MsgBox "请确认机组品牌!"
    comJzpb.SetFocus
    Exit Sub
End If
If txtJzxh.Text = "" Then
    MsgBox "请确认机组型号!"
    txtJzxh.SetFocus
    Exit Sub
End If
txtSL.Text = txtNsl.Text
comJzpb.Text = comJzPb1.Text
txtJzxh.Text = comJzXh.Text
Call cmdJgx_Click
End Sub

Private Sub cmdNQ_Click()
Dim ii As Integer
Dim tt As String
Dim Ra
Dim Tywy As String '单子流转到下一人的姓名
Dim Tuid As String
Dim Oywy As String '原来流转人的名字
Dim Ouid As String '原来流转人的工号
Dim CGF As Boolean '是否需要采购确认价格
Dim oo As Integer
On Error Resume Next
If Val(lblBaoId.Caption) = 0 Then
lblBaoId.Caption = ""
ElseIf lblBaoId.Caption <> "" Or lblBaoId.Caption <> 0 Then
'    MsgBox "已经生成报价单,不能签字!"
'    Exit Sub
End If

If Val(lblLc.Caption) = 0 Then lblLc.Caption = 1

If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If

If lblLc.Caption = 1 And Right(mod1.Bm, 3) = "工程部" And txtHtbh.Text = "" Then
    MsgBox "请正确关联大修合同编号!"
    cmdSave.Enabled = True
    Exit Sub
End If

If mod1.Bm = "零件事业部" And mod1.DName <> "张春华" Then
    lblLcRen.Caption = mod1.DName
    lblLcUid.Caption = mod1.DHid
End If

If lblLcUid.Caption <> mod1.DHid Then
    tt = "select xuid from htping where hid=" & Val(lblHtbh.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    If Ra(0, 0) <> lblLcUid.Caption Then
        MsgBox "此处应由" & lblLcRen.Caption & "签字! 请您不要再点"
        Exit Sub
    End If
End If

frmQm.Visible = True
If lblLc.Caption = 1 Then
    optT2.Enabled = False
    OptT1.Value = True
Else
    optT2.Enabled = True
    OptT1.Value = False
    optT2.Value = False
End If
If mod1.Bm = "零件事业部" Then
    optT2.Caption = "驳回"
Else
    optT2.Caption = "增补"
End If
End Sub

Private Sub cmdNQ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Val(lblLc.Caption) = 100 And Button = 2 And lblYwy.Caption = mod1.DName Then
    frmQm.Visible = True
    OptT1.Enabled = False
    optT2.Value = True
End If
End Sub

Private Sub cmdOpen_Click()
Dim tt As String
On Error Resume Next
tt = "select gyid,gymc,gyman,gyadr,gypho from gyNew where uid='" & mod1.DHid & "' or gyid=1"
adoGy.Close
adoGy.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set dtgGy.DataSource = adoGy
dtgGy.Visible = True

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

Private Sub cmdQing_Click()
'comLx.Text = ""
comJzpb.Text = ""
comJzXh.Text = ""
txtYxh.Text = ""
txtCbh.Text = ""
txtXlh.Text = ""
txtLjbh.Text = ""
txtLjmc.Text = ""
txtCd.Text = ""
txtDRQ.Text = ""
txtSL.Text = ""
End Sub





Private Sub cmdRight_Click()
Dim tt As String
Dim ZL As String
On Error Resume Next
Me.Enabled = False
frmWait.Show
frmWait.ZOrder
frmWait.Refresh
frmGXBj.adoOid.MoveNext
ZL = lblZl.Caption
'打开新建单
Call modBJD.BJDGXQing
Call modBJD.BJDGDBound(frmGXBj.adoOid.Fields("bid").Value)
Call modBJD.gxbjLocked
frmGXBj.cmdLeft.Enabled = True
frmGXBj.cmdBjd.Visible = False
'frmGXBj.cmdCong.Visible = False
frmGXBj.cmdWb.Visible = False
cmdMod.Enabled = False
cmdSave.Enabled = False

frmGXBj.adoOid.MoveNext
If frmGXBj.adoOid.EOF = True Then
    frmGXBj.lblZl.ForeColor = &H80000012
    frmGXBj.lblzlZ.ForeColor = &H80000012
    cmdMod.Enabled = True
    cmdRight.Enabled = False
    If (mod1.Bm = lblBM.Caption And mod1.BmJl = True Or mod1.DName = lblYwy.Caption Or (mod1.DName = "宋晓炯" Or mod1.DName = "宋晓炯1" Or mod1.DName = "倪旭")) And lblZl.Caption <> "购销" Then
        cmdWb.Visible = True
    Else
        cmdWb.Visible = False
    End If
    If mod1.DName = lblYwy.Caption Then
        If lblPwf.Caption = "True" Then
            cmdBjd.Visible = True
        End If
    If mod1.DName = lblYwy.Caption Then
        cmdCong.Visible = True
    End If
    End If
Else
    cmdRight.Enabled = True
End If

frmGXBj.adoOid.MovePrevious

frmWait.Visible = False
Me.Enabled = True
Me.ZOrder 0
End Sub

Private Sub cmdSave_Click()
Dim tt As String
Dim hg As Single
Dim Ra
On Error Resume Next
If comXmmc.Text = "" Then
    MsgBox "请输入项目名称!"
    Exit Sub
End If
'将自定价导入合同
tt = "select sum(jhg) from xunjiamx where bid=" & Val(lblBid.Caption) & " And gyid > 0"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
hg = mod1.HTP.Fields(0).Value
mod1.HTP.Close
Set mod1.HTP = Nothing
tt = "select htrow from xunjiaD where bid=" & Bid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
If Ra(0, 0) > 0 Then
    If FmxcNew.Visible = True Then
        FmxcNew.dtgLx.Col = 2
        FmxcNew.dtgLx.Row = Ra(0, 0)
        FmxcNew.dtgLx.Text = hg
    End If
Else
    If lblZl.Caption = "配件" Or lblZl.Caption = "配件询价单" Then
        tt = "update htping set w55=" & hg & " where hid=" & Val(lblHtbh.Caption)
    Else
        tt = "update htping set w66=" & hg & " where hid=" & Val(lblHtbh.Caption)
    End If
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    If FMXC.Visible = True Then
        FMXC.dtgFL.Col = 2
            If lblZl.Caption = "配件" Or lblZl.Caption = "配件询价单" Then
                FMXC.txtH5.Text = hg
                FMXC.dtgFL.Row = 6
                FMXC.dtgFL.Text = hg
            Else
                FMXC.dtgFL.Row = 7
                FMXC.txtH6.Text = hg
                FMXC.dtgFL.Text = hg
            End If
    End If
End If
Me.Enabled = False
frmWait.Visible = True
frmWait.ZOrder 0
cmdMod.Enabled = True
cmdSave.Enabled = False

''''''''''''Hg = 0
''''''''''''adoGx.MoveFirst
''''''''''''Do While Not adoGx.EOF
''''''''''''    Hg = Hg + adoGx.Fields("合计").Value
''''''''''''    adoGx.MoveNext
''''''''''''Loop
''''''''''''
''''''''''''Hg = Hg + Val(txtYf.Text)
''''''''''''txtHg.Text = Hg

tt = "select * from XunJiaD where bid=" & Val(lblBid.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
mod1.HTP.Update "xmmc", Trim(comXmmc.Text)    '项目名称
mod1.HTP.Update "xid", comXmmc.Tag  '项目代号
mod1.HTP.Update "bianhao", lblBh.Caption '单子编号(给用户看的)
mod1.HTP.Update "yhg", Val(txtYhg.Text) '小张优惠价
mod1.HTP.Update "bz", Trim(txtBz.Text)
mod1.HTP.Update "yf", Val(txtYf.Text)
mod1.HTP.Update "yfadr", Trim(txtADR.Text)
mod1.HTP.UpdateBatch



If lblFwid.Caption = "" Then
    lblLc.Caption = 1
    tt = "update xunJiaD set lc=1 where bid=" & Val(lblBid.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
End If
'''''    '添加事务
'''''    Call mod1.EnventAdd("询价单", comXmmc.Text, lblLcRen.Caption, lblLcUid.Caption, lblBid.Caption, lblQM(0).Caption, "", "", mod1.DName, mod1.DHid, 0, lblBid.Caption)
'''''    '更新按钮
'''''    Call modBJD.OpenXJAN(0)
'''''End If



'更新询价列表
'tt = "select * from xunjiaView where ywy='" & mod1.DName & "' and uid='" & mod1.DHid & "'"
'frmGxBiao.adoXj.Close
'frmGxBiao.adoXj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
frmGxBiao.adoXj.Requery
Set frmGxBiao.dtgXj.DataSource = frmGxBiao.adoXj

frmWait.Visible = False
Me.Enabled = True
Me.ZOrder 0
If txtYhg.Text = "" And FB = True Then '如果为分包,且没有产品,则直接跳至报价单
    cmdBjd.Visible = True
End If
End Sub

Private Sub cmdWb_Click()
Dim tt As String
On Error Resume Next
frmWBXJ.Visible = False
If frmWBXJ.comXmmc.Text = "" Then
    Call modBJD.BJDWBQing
    Call modBJD.BJDWDBound(Val(lblWbid.Caption))
    Call modBJD.wbxjLocked
    tt = "select bid from xunjiaOld where oid=" & Val(frmWBXJ.lblOid.Caption) & " order by bid"
    frmWBXJ.adoOid.Close
    frmWBXJ.adoOid.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText

    If frmWBXJ.adoOid.RecordCount > 1 Then
        frmWBXJ.cmdRight.Enabled = False
        frmWBXJ.cmdLeft.Enabled = True
    Else
        frmWBXJ.cmdRight.Enabled = False
        frmWBXJ.cmdRight.Enabled = False
    End If
    frmWBXJ.Visible = True
    frmGXBj.Visible = False
    frmWBXJ.adoOid.MoveLast
Else
    frmWBXJ.Visible = True
End If
'frmGXBj.Visible = False
frmWBXJ.lblZl.ForeColor = &HC000C0
frmWBXJ.lblzlZ.ForeColor = &HC000C0
frmWBXJ.ZOrder 0
frmGXBj.Visible = False

End Sub

Private Sub cmdZ_Click()

End Sub

Private Sub comJzpb_Click(Area As Integer)
Dim tt As String
On Error Resume Next

'''''If frmGXBj.Visible = False Then Exit Sub
'''''
'''''    tt = "select * from bjxt_jzxh where pbid='" & frmGXBj.comJzpb.BoundText & "'"
'''''    frmGXBj.adoJz.Close
'''''    frmGXBj.adoJz.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''    Set frmGXBj.comJzxh.RowSource = frmGXBj.adoJz
'''''    frmGXBj.comJzxh.ListField = "jzxh"
'''''    frmGXBj.comJzxh.BoundColumn = "xhid"
'''''    frmGXBj.adoJz.MoveFirst
'''''    frmGXBj.comJzxh.Text = frmGXBj.adoJz.Fields("jzxh").Value
'''''txtCd.Text = comJzpb.Text
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()

End Sub

Private Sub comXmmc_Change()
If Me.Visible = False Then Exit Sub
Dim tt As String
Dim adoHH As Object
On Error Resume Next
Set adoHH = CreateObject("adodb.recordset")
    comXmmc.Tag = comXmmc.BoundText
If comXmmc.Text <> "" And Right(mod1.Bm, 3) = "工程部" Then
    tt = "select htbh from htping where delf=1 and htf=1 and htxz='大修' and xid=" & comXmmc.Tag
    adoHH.Close
    adoHH.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    If adoHH.RecordCount > 0 Then
        txtHtbh.Text = adoHH.Fields("htbh").Value
    Else
        txtHtbh.Text = ""
    End If
End If
End Sub

Private Sub comXmmc_Click(Area As Integer)
'Dim tt As String
'On Error Resume Next
    comXmmc.Tag = comXmmc.BoundText

End Sub

Private Sub dtgGM_Click()
On Error Resume Next
txtGybz.Text = ""
dtgGM.Col = 1
txtGM.ToolTipText = dtgGM.Text
dtgGM.Col = 2
txtGM.Text = dtgGM.Text
''''''dtgGy.Col = 3
''''''txtGyman.Text = dtgGy.Text
''''''dtgGy.Col = 4
''''''txtGyAdr.Text = dtgGy.Text
''''''dtgGy.Col = 5
''''''txtGYPho.Text = dtgGy.Text
dtgGM.Visible = False
If txtGM.Text = "零件事业部" Or mod1.DName = "徐瑛" Or mod1.DName = "马晓聪" Then
    txtDj.Locked = True
'    If OPTN.Value = True Then
        dtgMa.Col = 13
'    Else
'        dtgMa.Col = 14
'    End If
    txtDj.Text = dtgMa.Text
    If Val(txtDj.Text) = 0 Then
        dtgMa.Col = 14
        txtDj.Text = Val(dtgMa.Text)
    End If
    lblDj.Visible = False
    txtDj.Visible = False
    frmJ.Visible = True
Else
    lblDj.Visible = True
    txtDj.Visible = True
    txtDj.Locked = False
    frmJ.Visible = False
End If
End Sub

Private Sub dtgGy_Click()
On Error Resume Next

dtgGy.Col = 1
txtGyid.Text = dtgGy.Text
dtgGy.Col = 2
txtGymc.Text = dtgGy.Text
dtgGy.Col = 3
txtGyman.Text = dtgGy.Text
dtgGy.Col = 4
txtGyAdr.Text = dtgGy.Text
dtgGy.Col = 5
txtGYPho.Text = dtgGy.Text
dtgGy.Visible = False
End Sub

Private Sub dtgMa_Click()
On Error Resume Next
frmWai.Visible = False
'''''MsgBox dtgMa.Col
'''''Exit Sub
If adoGx.RecordCount = 0 Then Exit Sub
If frmGXBj.Visible = False Then Exit Sub
comLx.Text = ""
comJzpb.Text = ""
comJzXh.Text = ""
txtYxh.Text = ""
txtCbh.Text = ""
txtXlh.Text = ""
txtLjbh.Text = ""
txtLjmc.Text = ""
txtCd.Text = ""
txtDRQ.Text = ""
txtSL.Text = ""
txtNsl.Text = ""
txtDj.Text = ""
txtBrq.Text = ""
lblLid.Caption = ""
comJzPb1.Text = ""
txtJzxh.Text = ""
Call GyQing

dtgN.Row = dtgMa.Row
dtgN.Col = 1
comLx.Text = dtgN.Text
dtgN.Col = 2
comJzpb.Text = dtgN.Text
comJzPb1.Text = dtgN.Text
dtgN.Col = 3
comJzXh.Text = dtgN.Text
txtJzxh.Text = dtgN.Text
dtgN.Col = 4
txtYxh.Text = dtgN.Text
dtgN.Col = 5
txtCbh.Text = dtgN.Text
dtgN.Col = 6
txtXlh.Text = dtgN.Text
dtgN.Col = 7
txtLjbh.Text = dtgN.Text
dtgN.Col = 8
txtLjmc.Text = dtgN.Text
dtgN.Col = 9
txtCd.Text = dtgN.Text
dtgN.Col = 10
txtDRQ.Text = dtgN.Text
dtgN.Col = 11
txtSL.Text = dtgN.Text
txtNsl.Text = dtgN.Text
dtgN.Col = 12
txtMj.Text = dtgN.Text
'If OPTN.Value = True Then
    dtgN.Col = 13
'Else
'    dtgn.Col = 14
'End If
txtDj.Text = Val(dtgN.Text)
If Val(txtDj.Text) = 0 Then
    dtgN.Col = 14
    txtDj.Text = Val(dtgN.Text)
End If
dtgN.Col = 15

txtJdj.Text = dtgN.Text



dtgN.Col = 19
txtBrq.Text = dtgN.Text
dtgN.Col = 20
txtZBQ.Text = dtgN.Text '质保期
dtgN.Col = 21
lblLid.Caption = dtgN.Text
If optW.Value = True Then
    dtgN.Col = 22
    txtGM.ToolTipText = dtgN.Text
    cmdGy.ToolTipText = dtgN.Text
    dtgN.Col = 23
    
    txtGM.Text = dtgN.Text
    dtgN.Col = 24
    txtGybz.Text = dtgN.Text
Else
    txtGM.ToolTipText = 1
    txtGM.Text = "零件事业部"
    dtgN.Col = 24
    txtGybz.Text = dtgN.Text

End If



End Sub




Private Sub dtgMa_RowColChange()
On Error Resume Next
frmWai.Visible = False
'''''MsgBox dtgMa.Col
'''''Exit Sub
If adoGx.RecordCount = 0 Then Exit Sub
If frmGXBj.Visible = False Then Exit Sub
comLx.Text = ""
comJzpb.Text = ""
comJzXh.Text = ""
txtYxh.Text = ""
txtCbh.Text = ""
txtXlh.Text = ""
txtLjbh.Text = ""
txtLjmc.Text = ""
txtCd.Text = ""
txtDRQ.Text = ""
txtSL.Text = ""
txtNsl.Text = ""
txtDj.Text = ""
txtBrq.Text = ""
lblLid.Caption = ""
comJzPb1.Text = ""
txtJzxh.Text = ""
Call GyQing

dtgN.Row = dtgMa.Row
dtgN.Col = 1
comLx.Text = dtgN.Text
dtgN.Col = 2
comJzpb.Text = dtgN.Text
comJzPb1.Text = dtgN.Text
dtgN.Col = 3
comJzXh.Text = dtgN.Text
txtJzxh.Text = dtgN.Text
dtgN.Col = 4
txtYxh.Text = dtgN.Text
dtgN.Col = 5
txtCbh.Text = dtgN.Text
dtgN.Col = 6
txtXlh.Text = dtgN.Text
dtgN.Col = 7
txtLjbh.Text = dtgN.Text
dtgN.Col = 8
txtLjmc.Text = dtgN.Text
dtgN.Col = 9
txtCd.Text = dtgN.Text
dtgN.Col = 10
txtDRQ.Text = dtgN.Text
dtgN.Col = 11
txtSL.Text = dtgN.Text
txtNsl.Text = dtgN.Text
dtgN.Col = 12
txtMj.Text = dtgN.Text
'If OPTN.Value = True Then
    dtgN.Col = 13
'Else
'    dtgn.Col = 14
'End If
txtDj.Text = Val(dtgN.Text)
If Val(txtDj.Text) = 0 Then
    dtgN.Col = 14
    txtDj.Text = Val(dtgN.Text)
End If
dtgN.Col = 15

txtJdj.Text = dtgN.Text



dtgN.Col = 19
txtBrq.Text = dtgN.Text
dtgN.Col = 20
txtZBQ.Text = dtgN.Text '质保期
dtgN.Col = 21
lblLid.Caption = dtgN.Text
If optW.Value = True Then
    dtgN.Col = 22
    txtGM.ToolTipText = dtgN.Text
    cmdGy.ToolTipText = dtgN.Text
    dtgN.Col = 23
    
    txtGM.Text = dtgN.Text
    dtgN.Col = 24
    txtGybz.Text = dtgN.Text
Else
    txtGM.ToolTipText = 1
    txtGM.Text = "零件事业部"
    dtgN.Col = 24
    txtGybz.Text = dtgN.Text

End If

End Sub


Private Sub dtpBrq_CloseUp()
txtBrq.Text = dtpBrq.Value
End Sub


Private Sub dtpDrq_CloseUp()
txtDRQ.Text = dtpDrq.Value
End Sub


Private Sub Form_Click()
dtgGy.Visible = False
dtgGM.Visible = False
frmQm.Visible = False
lblTX.Visible = False

End Sub
Public Sub QMBound(Bid As Long)
Dim Ra: Dim La
Dim ii As Integer: Dim oo As Integer
Dim tt As String
On Error Resume Next

tt = "select trq,ywy,zn,bz,tf from pizu where bh='" & Bid & "' and yid=43 order by pid desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2): dtgP.Rows = La + 20
dtgP.Clear
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

lblTX.Caption = "流程至:" & lblLcRen.Caption
lblTX.Visible = True

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
Private Sub Form_DblClick()
'''''''Dim tt As String
'''''''On Error Resume Next
'''''''If mod1.DName = "张春华" And lblZ.Visible = False And (cmdQm(1).Caption = "邹晨" Or cmdQm(1).Caption = "") Then
'''''''    Set mod1.cmd = createobject("adodb.command")
'''''''    mod1.cmd.ActiveConnection = mod1.cc
'''''''    mod1.cmd.CommandText = "CHX"
'''''''    mod1.cmd.CommandType = adCmdStoredProc
'''''''    mod1.cmd.Parameters("@Cfwid") = Val(lblCfwid.Caption)
'''''''    mod1.cmd.Parameters("@errch") = ""
'''''''    mod1.cmd.Parameters("@bh") = lblBid.Caption
'''''''    mod1.cmd.Execute
'''''''    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
'''''''            MsgBox "网络出现故障,请再试一次,如果还是提交不成功,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
'''''''            Exit Sub
'''''''    End If
'''''''    Set mod1.cmd = Nothing
'''''''    lblZ.Visible = True
'''''''    lblZT.Visible = True
'''''''    lblZT.Caption = mod1.DQda
'''''''    If Dialog.Visible = True Then '更新事务列表
'''''''        Call mod1.refEnvent(1)
'''''''    End If
'''''''End If
End Sub

Private Sub Form_Load()
Dim tt As String
On Error Resume Next
tt = "select jzpb,pbid from bjxt_jzpb"
frmGXBj.adoPb.Close
Set frmGXBj.adoPb = CreateObject("adodb.recordset")
frmGXBj.adoPb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGXBj.comJzPb1.RowSource = frmGXBj.adoPb

frmGXBj.comJzPb1.ListField = "jzpb"
frmGXBj.comJzPb1.BoundColumn = "pbid"
'frmSD.Left = 1710
'frmSD.Top = 270
frmYw.Left = 0
frmYw.Top = 4740
FB = False
Me.Left = 0
Me.Top = 0
frmCg.Top = 4740
frmGXBj.Width = mod1.FWidth
frmGXBj.Height = mod1.FHeight
Set adoPb = CreateObject("adodb.recordset")
Set adoJz = CreateObject("adodb.recordset")
Set adoXm = CreateObject("adodb.recordset")
Set adoOid = CreateObject("adodb.recordset")
Set adoGx = CreateObject("adodb.recordset")
Set adpgu = CreateObject("adodb.recordset")
Set adoGy = CreateObject("adodb.recordset")
dtpDrq.Value = Date
dtpBrq.Value = Date

''''''dtgMa.ColWidth(0) = 0
''''''dtgMa.ColWidth(1) = 0
''''''dtgMa.ColWidth(4) = 0
''''''dtgMa.ColWidth(5) = 0
''''''dtgMa.ColWidth(8) = 2000
''''''dtgMa.ColWidth(17) = 1000 '报价有效期
''''''dtgMa.ColWidth(16) = 0 '外包合计
''''''dtgMa.ColWidth(14) = 0 '外包单价
''''''dtgMa.ColWidth(18) = 0
''''''dtgMa.ColWidth(19) = 0
''''''dtgMa.ColWidth(gyid) = 0 '供应商编号
''''''dtgMa.ColWidth(gybz) = 0 '供应商备注
If mod1.Bm = "零件事业部" Then
    dtgMa.ColWidth(21) = 0

End If
    cmdExcel.Visible = True
If mod1.Mname <> "马晓聪" Then
    frmNew.Visible = False
End If
dtgGy.Left = 1680
dtgGy.Top = 990
frmWai.Top = 6090
frmWai.Left = 2460
dtgGM.Visible = False
dtgGy.ColWidth(0) = 0
dtgGy.ColWidth(1) = 0
dtgGy.ColWidth(2) = 25000
dtgGy.ColWidth(3) = 0
dtgGy.ColWidth(4) = 0
dtgGy.ColWidth(5) = 0
dtgGy.ColWidth(6) = 0
dtgGM.ColWidth(0) = 0
dtgGM.ColWidth(1) = 0
dtgGM.ColWidth(2) = 2500
OptT1.Value = True
frmQm.Left = 9000
frmQm.Top = 7440

dtgMa.ColWidth(25) = 0
dtgMa.ColWidth(26) = 0
dtgMa.ColWidth(27) = 0
dtgMa.ColWidth(28) = 0 'FJ
dtgMa.ColWidth(4) = 0
dtgNew.Left = 0
dtgNew.Top = 0
dtgP.Top = 6270
dtgP.Left = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim tt As String
Dim ii As Integer
On Error Resume Next
frmGxbjSD.Visible = False
If MDI.Cq = False Then
If cmdMod.Enabled = False And cmdSave.Enabled = True Then
    ii = MsgBox("新建单子没有保存,您确认要退出吗?", vbInformation + vbYesNo, "询问")
    If ii = vbYes Then
        tt = "delete from xunjiaD where bid=" & Val(lblBid.Caption)
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
       If cmdLeft.Enabled = True Then '将原先的作作废单子恢复。
            adoOid.MovePrevious
            tt = "update xunjiaD set xj=1 where bid=" & adoOid.Fields(0).Value
            Set mod1.HTP = CreateObject("adodb.recordset")
            mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
       End If
    Else
        Exit Sub
    End If
End If
Call modBJD.BJDGXQing
If frmGxBiao.Visible = True Then
    frmGxBiao.Enabled = True
    frmGxBiao.ZOrder 0
ElseIf FMXC.Visible = True Then
    FMXC.Enabled = True
    FMXC.ZOrder 0

ElseIf Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0

End If
End If




End Sub


Private Sub Option3_Click()

End Sub


Private Sub frmWai_Click()
dtgGy.Visible = False
End Sub

Private Sub Label16_DblClick()
If mod1.DName = "宋晓炯" Or mod1.DName = "邹晨" Or mod1.DName = "马晓聪" Or mod1.DName = "郑刚" Or mod1.DName = "乔继敏" Or mod1.DName = "王全红" Or mod1.DName = "杨燕" Or mod1.DName = "" Or Ywy = "吴金荣" Then
    frmJ.Visible = False
    lblDj.Visible = True
    txtDj.Visible = True
End If
End Sub

Private Sub lblDj_DblClick()
If mod1.DName = "宋晓炯" Or mod1.DName = "邹晨" Or mod1.DName = "马晓聪" Or mod1.DName = "郑刚" Or mod1.DName = "乔继敏" Or mod1.DName = "王全红" Or mod1.DName = "杨燕" Or mod1.DName = "" Or Ywy = "吴金荣" Then
frmJ.Visible = True
End If
End Sub

Private Sub OPTN_Click()
Dim oo As Integer
''''''dtgMa.ColWidth(16) = 0 '外包合计
''''''dtgMa.ColWidth(14) = 0 '外包单价
''''''dtgMa.ColWidth(13) = 1000
''''''dtgMa.ColWidth(15) = 1000
''''''dtgMa.ColWidth(17) = 1000
''''''dtgMa.ColWidth(21) = 0
On Error Resume Next
dtgMa.Row = 0
dtgMa.ColWidth(0) = 0
For oo = 0 To dtgMa.Cols - 1
    dtgMa.Col = oo
    If dtgMa.Text = "机组型号" Or dtgMa.Text = "零件编号" Or dtgMa.Text = "零件名称" Then
        dtgMa.ColWidth(oo) = 2000
    End If
    If dtgMa.Text = "到货期" Or dtgMa.Text = "报价有效期" Then
        dtgMa.ColWidth(oo) = 1500
    End If
    
    If dtgMa.Text = "压缩机型号" Or dtgMa.Text = "出厂编号" Or dtgMa.Text = "机组序列号" Or dtgMa.Text = "市场价" Or _
    dtgMa.Text = "bid" Or dtgMa.Text = "Lid" Or dtgMa.Text = "gyId" Or dtgMa.Text = "gyBZ" Or dtgMa.Text = "品种" Or dtgMa.Text = "外包单价" Or dtgMa.Text = "外包合计" Then
        dtgMa.ColWidth(oo) = 0
    End If
    If lblUid.Caption = mod1.DHid Then  '业务员，只显示基准价
        If dtgMa.Text = "成本单价" Or dtgMa.Text = "合计" Then
            dtgMa.ColWidth(oo) = 0
        End If
        If dtgMa.Text = "基准单价" Or dtgMa.Text = "基准合计" Then
            dtgMa.ColWidth(oo) = 1000
        End If
    ElseIf mod1.Bm = "零件事业部" Then
        If dtgMa.Text = "成本单价" Or dtgMa.Text = "合计" Then
            dtgMa.ColWidth(oo) = 1000
        End If
        If dtgMa.Text = "基准单价" Or dtgMa.Text = "基准合计" Then
            dtgMa.ColWidth(oo) = 0
        End If
    Else '其他人员都能看到
        If dtgMa.Text = "成本单价" Or dtgMa.Text = "合计" Or dtgMa.Text = "基准单价" Or dtgMa.Text = "基准合计" Then
            dtgMa.ColWidth(oo) = 1000
        End If
    End If
Next
txtHg.Text = Val(LBLhG.Caption)
txtYhg.Text = Val(LBLyHG.Caption)

cmdGx.Enabled = False

End Sub


Private Sub optW_Click()
If mod1.Bm = "零件事业部" Then Exit Sub
'''''''dtgMa.ColWidth(16) = 1000 '外包合计
'''''''dtgMa.ColWidth(14) = 1000 '外包单价
'''''''dtgMa.ColWidth(13) = 0
'''''''dtgMa.ColWidth(15) = 0
'''''''dtgMa.ColWidth(17) = 0
txtHg.Text = Val(lblWhg.Caption)
txtYhg.Text = Val(lblWhg.Caption)
dtgMa.Row = 0
dtgMa.ColWidth(0) = 0
For oo = 0 To dtgMa.Cols - 1
    dtgMa.Col = oo
    If dtgMa.Text = "机组型号" Or dtgMa.Text = "零件编号" Or dtgMa.Text = "零件名称" Then
        dtgMa.ColWidth(oo) = 2000
    End If
    If dtgMa.Text = "到货期" Or dtgMa.Text = "报价有效期" Then
        dtgMa.ColWidth(oo) = 1500
    End If
    If dtgMa.Text = "压缩机型号 " Or dtgMa.Text = "出厂编号" Or dtgMa.Text = "机组序列号" Or dtgMa.Text = "品牌产地" Or dtgMa.Text = "市场价" Or _
    dtgMa.Text = "bid" Or dtgMa.Text = "Lid" Or dtgMa.Text = "gyId" Or dtgMa.Text = "gyBZ" Or dtgMa.Text = "品种" Then
        dtgMa.ColWidth(oo) = 0
    End If
    If lblUid.Caption = mod1.DHid Then  '业务员，只显示基准价
        If dtgMa.Text = "成本单价" Or dtgMa.Text = "合计" Then
            dtgMa.ColWidth(oo) = 0
        End If
        If dtgMa.Text = "基准单价" Or dtgMa.Text = "基准合计" Then
            dtgMa.ColWidth(oo) = 1000
        End If
    ElseIf mod1.Bm = "零件事业部" Then
        If dtgMa.Text = "成本单价" Or dtgMa.Text = "合计" Then
            dtgMa.ColWidth(oo) = 1000
        End If
        If dtgMa.Text = "基准单价" Or dtgMa.Text = "基准合计" Then
            dtgMa.ColWidth(oo) = 0
        End If
    ElseIf mod1.Bm = "商务部" Then '其他人员都能看到
        If dtgMa.Text = "成本单价" Or dtgMa.Text = "合计" Or dtgMa.Text = "基准单价" Or dtgMa.Text = "基准合计" Then
            dtgMa.ColWidth(oo) = 1000
        End If
    Else
        If dtgMa.Text = "成本单价" Or dtgMa.Text = "合计" Then
            dtgMa.ColWidth(oo) = 0
        End If
        If dtgMa.Text = "基准单价" Or dtgMa.Text = "基准合计" Then
            dtgMa.ColWidth(oo) = 1000
        End If
    
    End If
Next
cmdGx.Enabled = False
End Sub


Private Sub timQuit_Timer()
On Error Resume Next
Dim oo As Integer
Dim jj As Integer
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
If timZm = 1 Then '如果为"决定何处购买"
    cmdJG.Enabled = False
ElseIf timZm = 2 Then
    cmdGx.Enabled = True
    If mod1.Bm <> "零件事业部" And mod1.DName <> "徐瑛" Then
        optW.Value = True
    End If
ElseIf timZm = 3 Or timZm = 6 Then       '添加配件,配件删除
    adoGx.Requery
    dtgMa.Visible = False
                frmGXBj.dtgMa.FixedCols = 1
    Set dtgMa.DataSource = adoGx
    Call dtgMaFF
    dtgMa.Visible = True
    '显示商务支持添加的产品（变色）
    For oo = 1 To frmGXBj.dtgMa.Rows
        frmGXBj.dtgMa.Col = 28
        frmGXBj.dtgMa.Row = oo
        If frmGXBj.dtgMa.Text = "True" Then
            For jj = 1 To 25
                frmGXBj.dtgMa.Col = jj
                frmGXBj.dtgMa.CellForeColor = &H8000000D
            Next
        End If
    Next
    If mod1.Bm = "配送中心" And timZm = 3 Then '让配送中心人可以签字
'''''        lblQM(0).Caption = ""
'''''        lblQM(1).Caption = ""
'''''        cmdQm(0).Caption = ""
'''''        cmdQm(1).Caption = ""
'''''        lblTm(0).Caption = ""
'''''        lblTm(1).Caption = ""
        lblLc.Caption = 1
        lblLcRen.Caption = mod1.DName
        lblLcUid.Caption = mod1.DHid
    End If
    
'''''    If adoGx.RecordCount > 1 Then
'''''    dtgMa.FixedRows = 0
'''''    dtgMa.MergeCol(1) = True
'''''    dtgMa.MergeCol(2) = True
'''''    dtgMa.MergeCol(10) = True
'''''    dtgMa.MergeCol(14) = True
'''''    dtgMa.MergeCells = 3
'''''    dtgMa.FixedRows = 1
'''''    End If
    comJzpb.Text = ""
    comJzXh.Text = ""
    txtYxh.Text = ""
    txtCbh.Text = ""
    txtXlh.Text = ""
    txtLjbh.Text = ""
    txtLjmc.Text = ""
    txtCd.Text = ""
    txtDRQ.Text = ""
    txtSL.Text = ""
    txtMj.Text = ""
    txtDj.Text = ""
    txtBrq.Text = ""
    cmdAdd.Enabled = True
    cmdDel.Enabled = True
    
   
ElseIf timZm = 4 Then      '配件更新
    adoGx.Requery
    dtgMa.Visible = False
                frmGXBj.dtgMa.FixedCols = 1
    Set dtgMa.DataSource = adoGx
    Call dtgMaFF
    dtgMa.Visible = True
    'comLx.Text = ""
    comJzpb.Text = ""
    comJzXh.Text = ""
    txtYxh.Text = ""
    txtCbh.Text = ""
    txtXlh.Text = ""
    txtLjbh.Text = ""
    txtLjmc.Text = ""
    txtCd.Text = ""
    txtDRQ.Text = ""
    txtSL.Text = ""
ElseIf timZm = 5 Then '供应商更新
    cmdGsave.Enabled = True
    txtGyid.Text = ""
    txtGymc.Text = ""
    txtGyman.Text = ""
    txtGyAdr.Text = ""
    txtGYPho.Text = ""
ElseIf timZm = 7 Then '签字
    cmdDing.Enabled = True
    txtQM.Text = ""
    frmQm.Visible = False
    'If cmdQm(2).Caption = "" Then
    lblTX.Visible = True
    'End If
    If Dialog.Visible = True Then '更新事务列表
        Call mod1.refEnvent(1)
    End If
    Call QMBound(Val(lblBid.Caption))
ElseIf timZm = 8 Then '删除
    Me.Visible = False
    If FMXC.Visible = True Then
        If lblZl.Caption = "零配件" Or lblZl.Caption = "配件询价单" Then
            FMXC.dtgFL.Row = 6
            FMXC.dtgFL.Col = 2
            FMXC.dtgFL.Text = ""
        ElseIf lblZl.Caption = "产品" Then
            FMXC.dtgFL.Col = 2
            FMXC.dtgFL.Row = 7
            FMXC.dtgFL.Text = ""
        End If
    End If
End If
timQuit.Enabled = False
End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh,mm1,mm2,mt1,mt2,mt3 from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0
    If timZm = 2 Then
        If mod1.Bm = "零件事业部" Or mod1.DName = "徐瑛" Then
            txtHg.Text = mod1.WP.Fields("mm1").Value
            txtYhg.Text = txtHg.Text
            LBLhG.Caption = txtHg.Text
            LBLyHG.Caption = txtHg.Text
        Else
            txtHg.Text = mod1.WP.Fields("mm2").Value
            txtYhg.Text = txtHg.Text
            lblWhg.Caption = txtHg.Text
            
        End If
        adoGx.Requery
        dtgMa.Visible = False
        dtgMa.Clear: dtgN.Clear
                    frmGXBj.dtgMa.FixedCols = 1
        Set dtgMa.DataSource = adoGx
            Call dtgMaFF
            dtgMa.Visible = True
    ElseIf timZm = 7 Then '签名
'''                If OptT1.Value = True Then
'''                    cmdQm(lblLc.Caption - 1).Caption = mod1.DName
'''                    lblTm(lblLc.Caption - 1).Caption = mod1.DQda
'''                Else
'''                    cmdQm(0).Caption = ""
'''                    lblTm(0).Caption = ""
'''                    cmdQm(1).Caption = ""
'''                    lblTm(1).Caption = ""
'''                    cmdQm(2).Caption = ""
'''                    lblTm(2).Caption = ""
'''                End If
                lblLc.Caption = mod1.WP.Fields("mm1").Value
                lblFwid.Caption = mod1.WP.Fields("mm2").Value
                lblLcRen.Caption = mod1.WP.Fields("mt1").Value
                lblLcUid.Caption = mod1.WP.Fields("mt2").Value
                lblTX.Caption = "下一流程,将跳至" & mod1.WP.Fields("mt3").Value & ": " & lblLcRen.Caption
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
    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdJG.Enabled = False
    End If
    Exit Sub

End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub



Public Sub GyQing()
txtGyid.Text = ""
txtGymc.Text = ""
txtGyman.Text = ""
txtGyAdr.Text = ""
txtGYPho.Text = ""
txtGybz.Text = ""
End Sub


Public Sub SDJE(Je As Double) '分摊速达金额
Dim CB As Single
Dim Lhg As Single
Dim oo As Integer
Dim LXG As Single
Dim GY As String
Dim LLG As Single
For oo = 1 To dtgMa.Rows - 1

    dtgMa.Row = oo
    

    dtgMa.Col = 23
    GY = Trim(dtgMa.Text)
    If GY <> "" Then
        dtgMa.Col = 18
        CB = CB + Val(dtgMa.Text)
    End If
    dtgMa.Col = 18
    If dtgMa.Text = "" Then
        Exit For
    End If
Next

''''''If mod1.BM = "零件事业部" Or mod1.DName = "徐瑛" Then
''''''    CB = Val(LBLwhG.Caption)
''''''Else
''''''    CB = Val(txtHg.Text)
''''''End If
If CB = 0 Then Exit Sub
    frmGXBj.lblSDJE.Caption = Je
    dtgMa.Rows = dtgMa.Rows + 20
    '分摊速达金额
    frmGXBj.dtgMa.Cols = frmGXBj.dtgMa.Cols + 1
    frmGXBj.dtgMa.Row = 0: frmGXBj.dtgMa.Col = 29: frmGXBj.dtgMa.Text = "速达小计": frmGXBj.dtgMa.CellFontBold = True: frmGXBj.dtgMa.CellForeColor = &H8000&
    dtgMa.ColWidth(29) = 900
    For oo = 1 To dtgMa.Rows + 1
        dtgMa.Row = oo
        dtgMa.Col = 23
        GY = Trim(dtgMa.Text)
        dtgMa.Col = 18
        LXG = Val(dtgMa.Text)
''''''        If dtgMa.Text = "" Then
''''''            Exit For
''''''        End If
        dtgMa.Col = 29
        If Round(Je * LXG / CB, 2) > 0 And GY <> "" Then
            dtgMa.Text = Round(Je * LXG / CB, 2)
            LLG = Lhg
            Lhg = Lhg + Val(dtgMa.Text)
            frmGXBj.dtgMa.CellForeColor = &H8000&
        Else
            dtgMa.Text = ""
        End If
        dtgMa.Row = oo + 1: dtgMa.Col = 18 '最后一行时,值为差值,确保没有1误差
        If dtgMa.Text = "" Then
             If Je <> Lhg Then
                dtgMa.Col = 29
                dtgMa.Row = oo
'''''                If Je > Lhg Then
'''''                    dtgMa.Text = Val(dtgMa.Text) + 1
'''''                Else
'''''                    dtgMa.Text = Val(dtgMa.Text) - 1
'''''                End If
                dtgMa.Text = Je - LLG
             End If
            Exit For
        End If
        
    Next
End Sub

Public Sub dtgMaFF()
On Error Resume Next
Dim oo As Integer
Dim jj As Integer

frmGXBj.dtgMa.Rows = frmGXBj.dtgMa.Rows + 20
frmGXBj.dtgN.Rows = frmGXBj.dtgMa.Rows
frmGXBj.dtgN.Cols = frmGXBj.dtgMa.Cols
    
    For oo = 0 To frmGXBj.dtgMa.Cols - 1
        frmGXBj.dtgMa.Col = oo
        frmGXBj.dtgMa.Row = 0
        If frmGXBj.dtgMa.Text = "机组型号" Or frmGXBj.dtgMa.Text = "零件编号" Or frmGXBj.dtgMa.Text = "零件名称" Then
            
            frmGXBj.dtgMa.ColWidth(oo) = 2000

        End If

        If frmGXBj.dtgMa.Text = "到货期" Or frmGXBj.dtgMa.Text = "报价有效期" Then
            frmGXBj.dtgMa.ColWidth(oo) = 1500
        End If
        If frmGXBj.dtgMa.Text = "压缩机型号 " Or frmGXBj.dtgMa.Text = "出厂编号" Or frmGXBj.dtgMa.Text = "机组序列号" Or frmGXBj.dtgMa.Text = "品牌产地" Or frmGXBj.dtgMa.Text = "市场价" Or _
        frmGXBj.dtgMa.Text = "bid" Or frmGXBj.dtgMa.Text = "Lid" Or frmGXBj.dtgMa.Text = "gyId" Or frmGXBj.dtgMa.Text = "gyBZ" Or frmGXBj.dtgMa.Text = "品种" Then
            frmGXBj.dtgMa.ColWidth(oo) = 0
        End If
            If frmGXBj.lblYwy = "谢雪梅" Or Val(lblBid.Caption) > 10058 Then
                If frmGXBj.dtgMa.Text = "压缩机型号" Then
                    frmGXBj.dtgMa.Text = "单位"
                    frmGXBj.dtgMa.ColWidth(oo) = 500
                ElseIf frmGXBj.dtgMa.Text = "机组型号" Then
                    frmGXBj.dtgMa.ColWidth(oo) = 1500
                ElseIf frmGXBj.dtgMa.Text = "零件编号" Then
                    frmGXBj.dtgMa.ColWidth(oo) = 1000
                    frmGXBj.dtgMa.Text = "货品编码"
                ElseIf frmGXBj.dtgMa.Text = "品牌产地" Then
                    frmGXBj.dtgMa.Text = "规格"
                    frmGXBj.dtgMa.ColWidth(oo) = 2500
                ElseIf frmGXBj.dtgMa.Text = "零件名称" Then

                    frmGXBj.dtgMa.Text = "货品名称"
                ElseIf frmGXBj.dtgMa.Text = "质保期" Then
                    frmGXBj.dtgMa.ColWidth(oo) = 1000
                End If
                
            End If
        If lblUid.Caption = mod1.DHid Then  '业务员，只显示基准价
            If frmGXBj.dtgMa.Text = "成本单价" Or frmGXBj.dtgMa.Text = "合计" Or frmGXBj.dtgMa.Text = "外包单价" Or frmGXBj.dtgMa.Text = "外包合计" Then
                frmGXBj.dtgMa.ColWidth(oo) = 0
            End If
            If frmGXBj.dtgMa.Text = "基准单价" Or frmGXBj.dtgMa.Text = "基准合计" Then
                frmGXBj.dtgMa.ColWidth(oo) = 1000
            End If
        ElseIf mod1.Bm = "零件事业部" Then
            If frmGXBj.dtgMa.Text = "成本单价" Or frmGXBj.dtgMa.Text = "合计" Then
                frmGXBj.dtgMa.ColWidth(oo) = 1000
            End If
            If frmGXBj.dtgMa.Text = "基准单价" Or frmGXBj.dtgMa.Text = "基准合计" Then
                frmGXBj.dtgMa.ColWidth(oo) = 1000
            End If
        ElseIf mod1.Bm = "商务部" Then '其他人员都能看到
            If frmGXBj.dtgMa.Text = "成本单价" Or frmGXBj.dtgMa.Text = "合计" Or frmGXBj.dtgMa.Text = "基准单价" Or frmGXBj.dtgMa.Text = "基准合计" Then
                frmGXBj.dtgMa.ColWidth(oo) = 1000
            End If
        Else
            If frmGXBj.dtgMa.Text = "成本单价" Or frmGXBj.dtgMa.Text = "合计" Then
                frmGXBj.dtgMa.ColWidth(oo) = 0
            End If
            If frmGXBj.dtgMa.Text = "基准单价" Or frmGXBj.dtgMa.Text = "基准合计" Then
                frmGXBj.dtgMa.ColWidth(oo) = 1000
            End If
        
        End If
    Next
        Set frmGXBj.dtgMa.DataSource = Nothing
        
        
    '显示商务支持添加的产品（变色）

    For oo = 1 To frmGXBj.dtgMa.Rows + 1
        frmGXBj.dtgMa.Col = 28
        frmGXBj.dtgMa.Row = oo
        frmGXBj.dtgN.Row = oo
        If frmGXBj.dtgMa.Text = "True" Then
            For jj = 1 To 25
                frmGXBj.dtgMa.Col = jj
                frmGXBj.dtgMa.CellForeColor = &HFF0000
            Next
        End If
        For jj = 1 To 25
            frmGXBj.dtgMa.Col = jj
            frmGXBj.dtgN.Col = jj
            frmGXBj.dtgN.Text = frmGXBj.dtgMa.Text
            If jj = 8 Or jj = 9 Or jj = 3 Or jj = 10 Then
                frmGXBj.dtgMa.Text = StrConv(frmGXBj.dtgMa.Text, vbWide)
                frmGXBj.dtgMa.CellFontWidth = 0

                If Len(frmGXBj.dtgMa.Text) > 10 Then
                    frmGXBj.dtgMa.RowHeight(oo) = 255 * (UpInt(Len(frmGXBj.dtgMa.Text) / 15) + 1)
                    'Exit For
                End If
            End If
            If jj = 10 Or jj = 19 Then
                frmGXBj.dtgMa.Text = Format(frmGXBj.dtgMa.Text, "YYYY-MM-DD")
            End If
        Next
    Next
    frmGXBj.dtgMa.FixedCols = 10
End Sub

