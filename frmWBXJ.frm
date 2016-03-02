VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmWBXJ 
   Caption         =   "维保询价单-人工"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Visible         =   0   'False
   Begin VB.ComboBox txtZu 
      Height          =   300
      Left            =   1380
      TabIndex        =   242
      Text            =   "Combo4"
      Top             =   6360
      Width           =   1725
   End
   Begin VB.Frame frmRG 
      Caption         =   "人工费表"
      Height          =   2235
      Left            =   60
      TabIndex        =   74
      Top             =   7140
      Visible         =   0   'False
      Width           =   5805
      Begin VB.TextBox txtFbnr 
         Height          =   1005
         Left            =   3330
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   90
         Top             =   750
         Width           =   2325
      End
      Begin VB.TextBox txtFBje 
         Height          =   270
         Left            =   1500
         TabIndex        =   88
         Top             =   1860
         Width           =   1635
      End
      Begin VB.TextBox txtF4 
         ForeColor       =   &H00C000C0&
         Height          =   285
         Left            =   1500
         TabIndex        =   84
         Top             =   1500
         Width           =   1635
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定"
         Height          =   315
         Left            =   5190
         TabIndex        =   81
         Top             =   1860
         Width           =   525
      End
      Begin VB.TextBox txtF3 
         Height          =   270
         Left            =   1500
         TabIndex        =   80
         Top             =   1140
         Width           =   1635
      End
      Begin VB.TextBox txtF2 
         Height          =   285
         Left            =   1500
         TabIndex        =   78
         Top             =   750
         Width           =   1635
      End
      Begin VB.TextBox txtF1 
         Height          =   270
         Left            =   1500
         TabIndex        =   76
         Top             =   330
         Width           =   1635
      End
      Begin VB.Label Label25 
         Caption         =   "分包内容"
         Height          =   195
         Left            =   3390
         TabIndex        =   89
         Top             =   390
         Width           =   735
      End
      Begin VB.Label Label24 
         Caption         =   "分包金额"
         Height          =   255
         Left            =   420
         TabIndex        =   87
         Top             =   1920
         Width           =   825
      End
      Begin VB.Label Label22 
         Caption         =   "风险基金"
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   420
         TabIndex        =   83
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "部门费用"
         Height          =   225
         Left            =   420
         TabIndex        =   79
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label Label20 
         Caption         =   "区域费用"
         Height          =   195
         Left            =   420
         TabIndex        =   77
         Top             =   810
         Width           =   825
      End
      Begin VB.Label Label19 
         Caption         =   "小组费用"
         Height          =   225
         Left            =   390
         TabIndex        =   75
         Top             =   390
         Width           =   1035
      End
   End
   Begin VB.Frame frmQm 
      BackColor       =   &H00C0FFC0&
      Caption         =   "评审建议"
      ForeColor       =   &H000000FF&
      Height          =   1785
      Left            =   8430
      TabIndex        =   92
      Top             =   6930
      Visible         =   0   'False
      Width           =   6315
      Begin VB.TextBox txtQM 
         BackColor       =   &H00C0FFFF&
         Height          =   1365
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   96
         Top             =   300
         Width           =   4965
      End
      Begin VB.OptionButton OptT1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "同意"
         Height          =   225
         Left            =   5220
         TabIndex        =   95
         Top             =   480
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.OptionButton optT2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "拒绝"
         Height          =   195
         Left            =   5220
         TabIndex        =   94
         Top             =   870
         Width           =   675
      End
      Begin VB.CommandButton cmdDing 
         BackColor       =   &H00FF8080&
         Caption         =   "决定"
         Height          =   285
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdDel 
      Enabled         =   0   'False
      Height          =   405
      Left            =   14280
      Picture         =   "frmWBXJ.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   241
      Top             =   8790
      Width           =   465
   End
   Begin VB.Frame frmJz 
      Caption         =   "价格体系"
      Height          =   4395
      Left            =   1890
      TabIndex        =   134
      Top             =   1080
      Visible         =   0   'False
      Width           =   15195
      Begin VB.Frame frmNewF 
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   675
         Left            =   300
         TabIndex        =   238
         Top             =   3600
         Width           =   3165
         Begin VB.OptionButton opt16 
            Caption         =   "续签"
            Height          =   255
            Left            =   1140
            TabIndex        =   240
            Top             =   300
            Width           =   1335
         End
         Begin VB.OptionButton opt15 
            Caption         =   "新签"
            Height          =   195
            Left            =   0
            TabIndex        =   239
            Top             =   330
            Width           =   855
         End
      End
      Begin VB.Frame frmM1 
         Caption         =   "主机"
         Height          =   2805
         Left            =   3930
         TabIndex        =   135
         Top             =   1620
         Width           =   10905
         Begin VB.CommandButton cmdBJ 
            Caption         =   "报价"
            Height          =   285
            Left            =   10260
            TabIndex        =   237
            Top             =   2280
            Width           =   555
         End
         Begin VB.ComboBox com13 
            Height          =   300
            ItemData        =   "frmWBXJ.frx":018A
            Left            =   7770
            List            =   "frmWBXJ.frx":0197
            TabIndex        =   234
            Text            =   "Combo5"
            Top             =   780
            Width           =   2325
         End
         Begin VB.Frame frmXH 
            Caption         =   "Frame11"
            Height          =   1005
            Left            =   6000
            TabIndex        =   172
            Top             =   1140
            Width           =   4095
            Begin VB.TextBox Text20 
               Height          =   285
               Left            =   1740
               TabIndex        =   236
               Top             =   510
               Width           =   2295
            End
            Begin VB.ComboBox comCxh 
               Height          =   300
               ItemData        =   "frmWBXJ.frx":01A5
               Left            =   1740
               List            =   "frmWBXJ.frx":01AF
               Style           =   2  'Dropdown List
               TabIndex        =   231
               Top             =   120
               Width           =   2355
            End
            Begin VB.Label Label61 
               Caption         =   "机组使用时间："
               Height          =   255
               Left            =   150
               TabIndex        =   235
               Top             =   570
               Width           =   1305
            End
            Begin VB.Label lblXh 
               Caption         =   "供热方式："
               Height          =   225
               Left            =   510
               TabIndex        =   173
               Top             =   150
               Width           =   1065
            End
         End
         Begin VB.Frame Frame10 
            BorderStyle     =   0  'None
            Caption         =   "Frame10"
            Height          =   555
            Left            =   120
            TabIndex        =   168
            Top             =   2100
            Width           =   4965
            Begin VB.CheckBox chk11 
               Caption         =   "化学清洁"
               Height          =   225
               Left            =   3150
               TabIndex        =   171
               Top             =   240
               Width           =   1335
            End
            Begin VB.CheckBox chk10 
               Caption         =   "物理清洗"
               Height          =   255
               Left            =   1800
               TabIndex        =   170
               Top             =   210
               Width           =   1305
            End
            Begin VB.Label Label39 
               Caption         =   "清洗方法："
               Height          =   225
               Left            =   600
               TabIndex        =   169
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.ComboBox comJ2 
            Height          =   300
            ItemData        =   "frmWBXJ.frx":01C3
            Left            =   4170
            List            =   "frmWBXJ.frx":01CD
            TabIndex        =   164
            Text            =   "USRT"
            Top             =   900
            Width           =   825
         End
         Begin VB.Frame Frame9 
            BorderStyle     =   0  'None
            Caption         =   "Frame9"
            Height          =   435
            Left            =   5310
            TabIndex        =   159
            Top             =   2190
            Width           =   5085
            Begin VB.OptionButton Option11 
               Caption         =   "大修"
               ForeColor       =   &H00C00000&
               Height          =   225
               Left            =   3780
               TabIndex        =   200
               Top             =   150
               Width           =   1155
            End
            Begin VB.OptionButton Option10 
               Caption         =   "一次性保养"
               Height          =   225
               Left            =   2310
               TabIndex        =   199
               Top             =   150
               Width           =   1245
            End
            Begin VB.OptionButton Option9 
               Caption         =   "维保"
               Height          =   255
               Left            =   1290
               TabIndex        =   198
               Top             =   120
               Width           =   855
            End
            Begin VB.Label Label36 
               Caption         =   "保养性质："
               Height          =   225
               Left            =   150
               TabIndex        =   167
               Top             =   150
               Width           =   915
            End
         End
         Begin VB.TextBox txtJ1 
            Height          =   270
            Left            =   2010
            TabIndex        =   144
            Top             =   930
            Width           =   1425
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   285
            Left            =   1620
            TabIndex        =   141
            Top             =   1380
            Width           =   3855
            Begin VB.TextBox txtJ5 
               Height          =   270
               Left            =   2940
               TabIndex        =   166
               Text            =   "1"
               Top             =   0
               Width           =   405
            End
            Begin VB.TextBox txtJ3 
               Height          =   270
               Left            =   1350
               TabIndex        =   165
               Text            =   "1"
               Top             =   0
               Width           =   375
            End
            Begin VB.CheckBox chkJ7 
               Caption         =   "冷凝器*"
               Height          =   255
               Left            =   2010
               TabIndex        =   143
               Top             =   0
               Width           =   975
            End
            Begin VB.CheckBox chkJ6 
               Caption         =   "蒸发器*"
               Height          =   225
               Left            =   390
               TabIndex        =   142
               Top             =   30
               Width           =   1035
            End
         End
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   255
            Left            =   1920
            TabIndex        =   138
            Top             =   1800
            Width           =   2865
            Begin VB.OptionButton optJ8 
               Caption         =   "拆一端"
               Height          =   225
               Left            =   120
               TabIndex        =   140
               Top             =   30
               Width           =   1095
            End
            Begin VB.OptionButton optJ9 
               Caption         =   "拆两端"
               Height          =   195
               Left            =   1350
               TabIndex        =   139
               Top             =   30
               Width           =   1035
            End
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            ItemData        =   "frmWBXJ.frx":01DB
            Left            =   2040
            List            =   "frmWBXJ.frx":01EB
            Style           =   2  'Dropdown List
            TabIndex        =   137
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txt12 
            Height          =   270
            Left            =   7770
            TabIndex        =   136
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label38 
            Caption         =   "单位"
            Height          =   225
            Left            =   3630
            TabIndex        =   163
            Top             =   960
            Width           =   375
         End
         Begin VB.Label lblJa 
            Caption         =   "机组冷量："
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   150
            Top             =   960
            Width           =   945
         End
         Begin VB.Label lblJa 
            Caption         =   "机组年巡视次数："
            Height          =   285
            Index           =   4
            Left            =   5970
            TabIndex        =   149
            Top             =   810
            Width           =   1515
         End
         Begin VB.Label Label6 
            Caption         =   "主机类型："
            Height          =   195
            Left            =   720
            TabIndex        =   148
            Top             =   420
            Width           =   1065
         End
         Begin VB.Label Label18 
            Caption         =   "拆端盖："
            Height          =   255
            Left            =   900
            TabIndex        =   147
            Top             =   1890
            Width           =   975
         End
         Begin VB.Label Label26 
            Caption         =   "清洗："
            Height          =   225
            Left            =   1080
            TabIndex        =   146
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label27 
            Caption         =   "单机组压缩机数量："
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5790
            TabIndex        =   145
            Top             =   390
            Width           =   1665
         End
      End
      Begin VB.Frame frmM3 
         Caption         =   "电机"
         Height          =   2715
         Left            =   5850
         TabIndex        =   223
         Top             =   3180
         Width           =   10635
         Begin VB.CheckBox Check18 
            Caption         =   "大修"
            Height          =   225
            Left            =   1320
            TabIndex        =   230
            Top             =   2070
            Width           =   795
         End
         Begin VB.CheckBox Check7 
            Caption         =   "保养"
            Height          =   195
            Left            =   1320
            TabIndex        =   229
            Top             =   2400
            Width           =   795
         End
         Begin VB.ComboBox Combo8 
            Height          =   300
            ItemData        =   "frmWBXJ.frx":0223
            Left            =   1560
            List            =   "frmWBXJ.frx":022D
            Style           =   2  'Dropdown List
            TabIndex        =   228
            Top             =   870
            Width           =   1515
         End
         Begin VB.TextBox Text19 
            Height          =   270
            Left            =   1560
            TabIndex        =   226
            Text            =   "Text19"
            Top             =   450
            Width           =   1485
         End
         Begin VB.Label Label59 
            Caption         =   "品牌："
            Height          =   225
            Left            =   810
            TabIndex        =   227
            Top             =   960
            Width           =   555
         End
         Begin VB.Label Label58 
            Caption         =   "功率(KW)："
            Height          =   255
            Left            =   450
            TabIndex        =   225
            Top             =   480
            Width           =   1005
         End
         Begin VB.Label Label56 
            Caption         =   "保养性质："
            Height          =   225
            Left            =   300
            TabIndex        =   224
            Top             =   2130
            Width           =   915
         End
      End
      Begin VB.Frame frmM5 
         Caption         =   "小机"
         Height          =   2805
         Left            =   2640
         TabIndex        =   191
         Top             =   630
         Width           =   10965
         Begin VB.Frame Frame16 
            Caption         =   "空调箱"
            Height          =   2085
            Left            =   6780
            TabIndex        =   218
            Top             =   450
            Width           =   2295
            Begin VB.TextBox Text18 
               Height          =   285
               Left            =   1140
               TabIndex        =   222
               Top             =   900
               Width           =   855
            End
            Begin VB.TextBox Text17 
               Height          =   270
               Left            =   1140
               TabIndex        =   220
               Top             =   330
               Width           =   825
            End
            Begin VB.Label Label55 
               Caption         =   "巡视次数："
               Height          =   255
               Left            =   120
               TabIndex        =   221
               Top             =   960
               Width           =   945
            End
            Begin VB.Label Label54 
               Caption         =   "保养次数："
               Height          =   195
               Left            =   120
               TabIndex        =   219
               Top             =   390
               Width           =   945
            End
         End
         Begin VB.Frame Frame15 
            Caption         =   "小机安装"
            Height          =   2115
            Left            =   2190
            TabIndex        =   213
            Top             =   450
            Width           =   2355
            Begin VB.TextBox Text16 
               Height          =   270
               Left            =   1500
               TabIndex        =   217
               Top             =   870
               Width           =   615
            End
            Begin VB.TextBox Text15 
               Height          =   270
               Left            =   1500
               TabIndex        =   215
               Top             =   390
               Width           =   645
            End
            Begin VB.Label Label53 
               Caption         =   "外机数量(>3HP)"
               Height          =   195
               Left            =   180
               TabIndex        =   216
               Top             =   930
               Width           =   1395
            End
            Begin VB.Label Label52 
               Caption         =   "外机数量(<3HP)"
               Height          =   285
               Left            =   150
               TabIndex        =   214
               Top             =   420
               Width           =   1455
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "风机盘管"
            Height          =   2085
            Left            =   4530
            TabIndex        =   201
            Top             =   450
            Width           =   2265
            Begin VB.TextBox Text14 
               Height          =   270
               Left            =   1080
               TabIndex        =   212
               Top             =   720
               Width           =   675
            End
            Begin VB.TextBox Text13 
               Height          =   270
               Left            =   1080
               TabIndex        =   211
               Top             =   360
               Width           =   675
            End
            Begin VB.CheckBox Check6 
               Caption         =   "巡视"
               Height          =   225
               Left            =   1020
               TabIndex        =   208
               Top             =   1680
               Width           =   825
            End
            Begin VB.CheckBox Check5 
               Caption         =   "保养"
               Height          =   195
               Left            =   180
               TabIndex        =   207
               Top             =   1710
               Width           =   795
            End
            Begin VB.Label Label51 
               Caption         =   "巡视次数："
               Height          =   225
               Left            =   150
               TabIndex        =   210
               Top             =   720
               Width           =   945
            End
            Begin VB.Label Label50 
               Caption         =   "保养次数："
               Height          =   195
               Left            =   150
               TabIndex        =   209
               Top             =   390
               Width           =   945
            End
            Begin VB.Label Label47 
               Caption         =   "保养性质："
               Height          =   225
               Left            =   150
               TabIndex        =   202
               Top             =   1140
               Width           =   915
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   "小机"
            Height          =   2115
            Left            =   120
            TabIndex        =   192
            Top             =   450
            Width           =   2085
            Begin VB.TextBox Text12 
               Height          =   270
               Left            =   1110
               TabIndex        =   206
               Top             =   780
               Width           =   675
            End
            Begin VB.TextBox Text11 
               Height          =   270
               Left            =   1110
               TabIndex        =   205
               Top             =   420
               Width           =   675
            End
            Begin VB.CheckBox Check17 
               Caption         =   "保养"
               Height          =   195
               Left            =   210
               TabIndex        =   196
               Top             =   1470
               Width           =   795
            End
            Begin VB.CheckBox Check16 
               Caption         =   "巡视"
               Height          =   225
               Left            =   1050
               TabIndex        =   195
               Top             =   1440
               Width           =   825
            End
            Begin VB.CheckBox Check15 
               Caption         =   "应急"
               Height          =   225
               Left            =   210
               TabIndex        =   194
               Top             =   1800
               Width           =   795
            End
            Begin VB.CheckBox Check14 
               Caption         =   "移机"
               ForeColor       =   &H00C00000&
               Height          =   255
               Left            =   1050
               TabIndex        =   193
               Top             =   1740
               Width           =   825
            End
            Begin VB.Label Label49 
               Caption         =   "巡视次数："
               Height          =   225
               Left            =   180
               TabIndex        =   204
               Top             =   780
               Width           =   945
            End
            Begin VB.Label Label48 
               Caption         =   "保养次数："
               Height          =   195
               Left            =   180
               TabIndex        =   203
               Top             =   450
               Width           =   945
            End
            Begin VB.Label Label46 
               Caption         =   "保养性质："
               Height          =   225
               Left            =   180
               TabIndex        =   197
               Top             =   1140
               Width           =   915
            End
         End
      End
      Begin VB.Frame frmM2 
         Caption         =   "水泵"
         Height          =   2745
         Left            =   6720
         TabIndex        =   174
         Top             =   2430
         Width           =   10905
         Begin VB.ComboBox Combo7 
            Height          =   300
            ItemData        =   "frmWBXJ.frx":023D
            Left            =   3900
            List            =   "frmWBXJ.frx":025F
            TabIndex        =   190
            Text            =   "1"
            Top             =   1560
            Width           =   1425
         End
         Begin VB.ComboBox Combo6 
            Height          =   300
            ItemData        =   "frmWBXJ.frx":0282
            Left            =   3900
            List            =   "frmWBXJ.frx":028C
            Style           =   2  'Dropdown List
            TabIndex        =   188
            Top             =   1050
            Width           =   1425
         End
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   3900
            TabIndex        =   186
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox Combo3 
            Height          =   300
            ItemData        =   "frmWBXJ.frx":029C
            Left            =   1410
            List            =   "frmWBXJ.frx":02A6
            Style           =   2  'Dropdown List
            TabIndex        =   179
            Top             =   1050
            Width           =   975
         End
         Begin VB.TextBox Text9 
            Height          =   270
            Left            =   1380
            TabIndex        =   177
            Top             =   510
            Width           =   975
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   615
            Left            =   360
            TabIndex        =   175
            Top             =   1890
            Width           =   5955
            Begin VB.CheckBox Check13 
               Caption         =   "急修"
               Height          =   255
               Left            =   4080
               TabIndex        =   184
               Top             =   270
               Width           =   825
            End
            Begin VB.CheckBox Check12 
               Caption         =   "大修"
               Height          =   225
               Left            =   3180
               TabIndex        =   183
               Top             =   270
               Width           =   795
            End
            Begin VB.CheckBox Check11 
               Caption         =   "巡视"
               Height          =   225
               Left            =   2250
               TabIndex        =   182
               Top             =   270
               Width           =   825
            End
            Begin VB.CheckBox Check10 
               Caption         =   "保养"
               Height          =   195
               Left            =   1320
               TabIndex        =   181
               Top             =   270
               Width           =   795
            End
            Begin VB.Label Label42 
               Caption         =   "保养性质："
               Height          =   225
               Left            =   150
               TabIndex        =   180
               Top             =   270
               Width           =   915
            End
         End
         Begin VB.Label Label45 
            Caption         =   "水泵级数："
            Height          =   225
            Left            =   2910
            TabIndex        =   189
            Top             =   1620
            Width           =   915
         End
         Begin VB.Label Label44 
            Caption         =   "水泵类型："
            Height          =   195
            Left            =   2910
            TabIndex        =   187
            Top             =   1110
            Width           =   945
         End
         Begin VB.Label Label43 
            Caption         =   "巡视次数："
            Height          =   255
            Left            =   2910
            TabIndex        =   185
            Top             =   540
            Width           =   1035
         End
         Begin VB.Label Label41 
            Caption         =   "品牌："
            Height          =   195
            Left            =   780
            TabIndex        =   178
            Top             =   1110
            Width           =   555
         End
         Begin VB.Label Label40 
            Caption         =   "功率（KW）："
            Height          =   225
            Left            =   240
            TabIndex        =   176
            Top             =   570
            Width           =   1095
         End
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   1170
         TabIndex        =   162
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox txtJSL 
         Height          =   285
         Left            =   1170
         TabIndex        =   161
         Top             =   3270
         Width           =   2565
      End
      Begin VB.TextBox txtXLBH 
         Height          =   270
         Left            =   1170
         TabIndex        =   158
         Top             =   2850
         Width           =   2565
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   1170
         TabIndex        =   156
         Top             =   2430
         Width           =   2505
      End
      Begin VB.ComboBox comDX 
         Height          =   300
         ItemData        =   "frmWBXJ.frx":02B6
         Left            =   1170
         List            =   "frmWBXJ.frx":02CF
         Style           =   2  'Dropdown List
         TabIndex        =   153
         Top             =   1620
         Width           =   2565
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgJG 
         Height          =   1305
         Left            =   0
         TabIndex        =   151
         Top             =   210
         Width           =   15165
         _ExtentX        =   26749
         _ExtentY        =   2302
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label60 
         Caption         =   "欢迎使用："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   645
         Left            =   5940
         TabIndex        =   233
         Top             =   2310
         Width           =   2625
      End
      Begin VB.Label Label57 
         Caption         =   "超白金核价体系！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   9120
         TabIndex        =   232
         Top             =   2340
         Width           =   4245
      End
      Begin VB.Label Label37 
         Caption         =   "数量："
         Height          =   255
         Left            =   540
         TabIndex        =   160
         Top             =   3330
         Width           =   705
      End
      Begin VB.Label Label35 
         Caption         =   "序列编号："
         Height          =   225
         Left            =   180
         TabIndex        =   157
         Top             =   2880
         Width           =   945
      End
      Begin VB.Label Label34 
         Caption         =   "型号："
         Height          =   225
         Left            =   540
         TabIndex        =   155
         Top             =   2460
         Width           =   675
      End
      Begin VB.Label Label33 
         Caption         =   "品牌："
         Height          =   225
         Left            =   540
         TabIndex        =   154
         Top             =   2070
         Width           =   555
      End
      Begin VB.Label Label32 
         Caption         =   "保养对象："
         Height          =   285
         Left            =   150
         TabIndex        =   152
         Top             =   1680
         Width           =   1005
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3675
      Left            =   840
      TabIndex        =   112
      Top             =   180
      Visible         =   0   'False
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   6482
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "主机维保"
      TabPicture(0)   =   "frmWBXJ.frx":0307
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "溴化锂"
      TabPicture(1)   =   "frmWBXJ.frx":0323
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "水泵保养 巡视和大修"
      TabPicture(2)   =   "frmWBXJ.frx":033F
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "水泵电机保养 大修"
      TabPicture(3)   =   "frmWBXJ.frx":035B
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "小机 风机盘管的保养 巡视 急修"
      TabPicture(4)   =   "frmWBXJ.frx":0377
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "多联机 传奇的安装"
      TabPicture(5)   =   "frmWBXJ.frx":0393
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "空调箱维保"
      TabPicture(6)   =   "frmWBXJ.frx":03AF
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      Begin VB.Frame Frame5 
         Height          =   3255
         Left            =   -74910
         TabIndex        =   113
         Top             =   630
         Width           =   14865
         Begin VB.TextBox txtJa 
            Height          =   270
            Index           =   3
            Left            =   2340
            TabIndex        =   126
            Top             =   1050
            Width           =   1905
         End
         Begin VB.TextBox txtJa 
            Height          =   270
            Index           =   2
            Left            =   2310
            TabIndex        =   125
            Top             =   2790
            Width           =   2325
         End
         Begin VB.TextBox txtJa 
            Height          =   270
            Index           =   1
            Left            =   11460
            TabIndex        =   124
            Top             =   480
            Width           =   1515
         End
         Begin VB.Frame Frame7 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   285
            Left            =   1950
            TabIndex        =   121
            Top             =   1410
            Width           =   2625
            Begin VB.CheckBox Check4 
               Caption         =   "冷凝器"
               Height          =   255
               Left            =   1650
               TabIndex        =   123
               Top             =   0
               Width           =   1005
            End
            Begin VB.CheckBox Check3 
               Caption         =   "蒸发器"
               Height          =   225
               Left            =   390
               TabIndex        =   122
               Top             =   30
               Width           =   1035
            End
         End
         Begin VB.Frame Frame6 
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   255
            Left            =   2250
            TabIndex        =   118
            Top             =   1830
            Width           =   2865
            Begin VB.OptionButton Option8 
               Caption         =   "只拆一端"
               Height          =   225
               Left            =   120
               TabIndex        =   120
               Top             =   30
               Width           =   1095
            End
            Begin VB.OptionButton Option7 
               Caption         =   "两者都清洗"
               Height          =   195
               Left            =   1350
               TabIndex        =   119
               Top             =   30
               Width           =   1335
            End
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            ItemData        =   "frmWBXJ.frx":03CB
            Left            =   2370
            List            =   "frmWBXJ.frx":03DB
            Style           =   2  'Dropdown List
            TabIndex        =   117
            Top             =   390
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Height          =   270
            Left            =   2310
            TabIndex        =   116
            Top             =   2310
            Width           =   2295
         End
         Begin VB.OptionButton Option6 
            Caption         =   "新签"
            Height          =   195
            Left            =   9930
            TabIndex        =   115
            Top             =   1380
            Width           =   855
         End
         Begin VB.OptionButton Option5 
            Caption         =   "续签"
            Height          =   255
            Left            =   11460
            TabIndex        =   114
            Top             =   1350
            Width           =   1335
         End
         Begin VB.Label lblJa 
            Caption         =   "机组冷量(USRT)"
            Height          =   285
            Index           =   3
            Left            =   630
            TabIndex        =   133
            Top             =   1080
            Width           =   2025
         End
         Begin VB.Label lblJa 
            Caption         =   "机组年巡视次数："
            Height          =   285
            Index           =   2
            Left            =   510
            TabIndex        =   132
            Top             =   2820
            Width           =   1605
         End
         Begin VB.Label lblJa 
            Caption         =   "机组数量："
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   1
            Left            =   9690
            TabIndex        =   131
            Top             =   510
            Width           =   1605
         End
         Begin VB.Label Label31 
            Caption         =   "机组类型"
            Height          =   195
            Left            =   1140
            TabIndex        =   130
            Top             =   450
            Width           =   765
         End
         Begin VB.Label Label30 
            Caption         =   "清洗时拆换热器："
            Height          =   255
            Left            =   510
            TabIndex        =   129
            Top             =   1890
            Width           =   1485
         End
         Begin VB.Label Label29 
            Caption         =   "清洗："
            Height          =   225
            Left            =   1410
            TabIndex        =   128
            Top             =   1470
            Width           =   615
         End
         Begin VB.Label Label28 
            Caption         =   "单机组压缩机数量："
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   330
            TabIndex        =   127
            Top             =   2340
            Width           =   1665
         End
      End
   End
   Begin VB.Frame frmN 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1515
      Left            =   2880
      TabIndex        =   106
      Top             =   7740
      Width           =   3675
      Begin VB.TextBox txt2 
         Height          =   315
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   110
         Top             =   900
         Width           =   2115
      End
      Begin VB.TextBox txt1 
         Height          =   315
         Left            =   1260
         TabIndex        =   108
         Top             =   360
         Width           =   2085
      End
      Begin VB.Label lbl2 
         Caption         =   "基准价格"
         Height          =   285
         Left            =   270
         TabIndex        =   109
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lbl1 
         Caption         =   "人工成本"
         Height          =   255
         Left            =   270
         TabIndex        =   107
         Top             =   420
         Width           =   855
      End
   End
   Begin VB.Frame frmOld 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   945
      Left            =   4770
      TabIndex        =   99
      Top             =   8160
      Width           =   3615
      Begin VB.TextBox txtHg 
         Height          =   285
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   102
         Top             =   30
         Width           =   975
      End
      Begin VB.TextBox txtYhg 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   101
         ToolTipText     =   "此处由工程部填入"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtClf 
         Height          =   285
         Left            =   2520
         TabIndex        =   100
         Top             =   0
         Width           =   975
      End
      Begin VB.Label lblHG 
         Caption         =   "人工费"
         Height          =   255
         Left            =   0
         TabIndex        =   105
         Top             =   30
         Width           =   675
      End
      Begin VB.Label Label11 
         Caption         =   "总费用"
         Height          =   255
         Left            =   1830
         TabIndex        =   104
         ToolTipText     =   "此处由工程部填入"
         Top             =   390
         Width           =   615
      End
      Begin VB.Label lblClf 
         Caption         =   "差旅费"
         Height          =   255
         Left            =   1830
         TabIndex        =   103
         Top             =   30
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdHT 
      BackColor       =   &H00C0FFC0&
      Caption         =   "合同评审单"
      Height          =   525
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   6900
      Width           =   1065
   End
   Begin VB.Timer timQuit 
      Interval        =   1000
      Left            =   13590
      Top             =   5730
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12960
      Top             =   5640
   End
   Begin VB.TextBox txtBz 
      Height          =   945
      Left            =   5160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   86
      Top             =   6630
      Width           =   8775
   End
   Begin VB.CommandButton cmdPje 
      Caption         =   "评审建议"
      Height          =   1035
      Left            =   8760
      TabIndex        =   82
      Top             =   8100
      Width           =   345
   End
   Begin VB.Frame Frame1 
      Caption         =   "机组信息"
      Height          =   2025
      Left            =   4890
      TabIndex        =   60
      Top             =   4620
      Width           =   9465
      Begin VB.Frame frmNew 
         Caption         =   "Frame2"
         Height          =   1845
         Left            =   90
         TabIndex        =   67
         Top             =   180
         Width           =   5655
         Begin VB.CommandButton cmdJdel 
            Caption         =   "删  除"
            Height          =   375
            Left            =   4530
            TabIndex        =   70
            Top             =   600
            Width           =   1035
         End
         Begin VB.CommandButton cmdJadd 
            Caption         =   "添  加"
            Height          =   375
            Left            =   4530
            TabIndex        =   69
            Top             =   210
            Width           =   1035
         End
         Begin VB.CommandButton cmdJgx 
            Caption         =   "更  新"
            Height          =   345
            Left            =   4530
            TabIndex        =   68
            Top             =   990
            Width           =   1035
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgA 
            Height          =   1695
            Left            =   0
            TabIndex        =   71
            Top             =   0
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   2990
            _Version        =   393216
            SelectionMode   =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.TextBox txtSL 
         Height          =   285
         Left            =   7140
         TabIndex        =   61
         Top             =   1620
         Width           =   2235
      End
      Begin MSDataListLib.DataCombo comXh 
         Height          =   330
         Left            =   7140
         TabIndex        =   62
         Top             =   1230
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo comPb 
         Height          =   330
         Left            =   7140
         TabIndex        =   63
         Top             =   780
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label2 
         Caption         =   "机组型号:"
         Height          =   225
         Left            =   5880
         TabIndex        =   66
         Top             =   1245
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "机组品牌:"
         Height          =   225
         Left            =   5850
         TabIndex        =   65
         Top             =   855
         Width           =   1125
      End
      Begin VB.Label Label3 
         Caption         =   "数量:"
         Height          =   225
         Left            =   6240
         TabIndex        =   64
         Top             =   1650
         Width           =   675
      End
   End
   Begin VB.Frame frmHide 
      Caption         =   "frmHid"
      Height          =   1455
      Left            =   10050
      TabIndex        =   24
      Top             =   4380
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Label lblBm 
         Caption         =   "lblBm"
         Height          =   225
         Left            =   960
         TabIndex        =   57
         Top             =   150
         Width           =   915
      End
      Begin VB.Label lblQy 
         Caption         =   "lblQy"
         Height          =   255
         Left            =   2430
         TabIndex        =   56
         Top             =   180
         Width           =   1155
      End
      Begin VB.Label lblPwf 
         Caption         =   "lblPwf"
         Height          =   255
         Left            =   3600
         TabIndex        =   55
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label lblLcou 
         Caption         =   "lblLcou"
         Height          =   255
         Left            =   1860
         TabIndex        =   34
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label lblYwy 
         Caption         =   "lblYwy"
         Height          =   285
         Left            =   3540
         TabIndex        =   31
         Top             =   450
         Width           =   765
      End
      Begin VB.Label lblUid 
         Caption         =   "lblUid"
         Height          =   255
         Left            =   3750
         TabIndex        =   30
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblFwid 
         Caption         =   "lblFwid"
         Height          =   255
         Left            =   1860
         TabIndex        =   29
         Top             =   450
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblLcUid 
         Caption         =   "lblLcUid"
         Height          =   285
         Left            =   240
         TabIndex        =   28
         Top             =   930
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblLcRen 
         Caption         =   "lblLcRen"
         Height          =   285
         Left            =   150
         TabIndex        =   27
         Top             =   420
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblNlb 
         Caption         =   "lblNlb"
         Height          =   225
         Left            =   1920
         TabIndex        =   26
         Top             =   810
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label lblLc 
         Caption         =   "lblLc"
         Height          =   315
         Left            =   1050
         TabIndex        =   25
         Top             =   630
         Visible         =   0   'False
         Width           =   645
      End
   End
   Begin VB.Frame frmNb 
      Height          =   405
      Left            =   330
      TabIndex        =   48
      Top             =   6720
      Width           =   4125
      Begin VB.TextBox txtWc 
         Height          =   270
         Left            =   1050
         TabIndex        =   50
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox txtXc 
         Height          =   270
         Left            =   3330
         TabIndex        =   49
         Top             =   120
         Width           =   405
      End
      Begin VB.Label Label10 
         Caption         =   "维保年限:"
         Height          =   225
         Left            =   60
         TabIndex        =   54
         Top             =   150
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "年"
         Height          =   225
         Left            =   1650
         TabIndex        =   53
         Top             =   150
         Width           =   255
      End
      Begin VB.Label Label14 
         Caption         =   "例检次数"
         Height          =   225
         Left            =   2430
         TabIndex        =   52
         Top             =   150
         Width           =   825
      End
      Begin VB.Label Label15 
         Caption         =   "次"
         Height          =   195
         Left            =   3810
         TabIndex        =   51
         Top             =   180
         Width           =   255
      End
   End
   Begin VB.Frame frmDx 
      Height          =   375
      Left            =   240
      TabIndex        =   44
      Top             =   6750
      Width           =   2235
      Begin VB.TextBox txtMon 
         Height          =   270
         Left            =   1290
         TabIndex        =   45
         Top             =   90
         Width           =   525
      End
      Begin VB.Label Label17 
         Caption         =   "维修保质期"
         Height          =   225
         Left            =   120
         TabIndex        =   47
         Top             =   120
         Width           =   1065
      End
      Begin VB.Label Label16 
         Caption         =   "月"
         Height          =   255
         Left            =   1950
         TabIndex        =   46
         Top             =   120
         Width           =   195
      End
   End
   Begin TabDlg.SSTab tabGc 
      Height          =   4395
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   7752
      _Version        =   393216
      TabOrientation  =   1
      TabHeight       =   520
      TabCaption(0)   =   "年保"
      TabPicture(0)   =   "frmWBXJ.frx":0413
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dtgWb"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "例检"
      TabPicture(1)   =   "frmWBXJ.frx":042F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dtgLj"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "大修"
      TabPicture(2)   =   "frmWBXJ.frx":044B
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtDxnr"
      Tab(2).ControlCount=   1
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
         Height          =   4065
         Left            =   -74970
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Top             =   30
         Width           =   15165
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgWb 
         Bindings        =   "frmWBXJ.frx":0467
         Height          =   3645
         Left            =   90
         TabIndex        =   39
         Top             =   120
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   6429
         _Version        =   393216
         Cols            =   5
         ForeColorSel    =   -2147483646
         BackColorBkg    =   -2147483627
         WordWrap        =   -1  'True
         FillStyle       =   1
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
         _Band(0)._NumMapCols=   4
         _Band(0)._MapCol(0)._Name=   "UserId"
         _Band(0)._MapCol(0)._RSIndex=   0
         _Band(0)._MapCol(1)._Name=   "UserName"
         _Band(0)._MapCol(1)._RSIndex=   1
         _Band(0)._MapCol(2)._Name=   "UserPw"
         _Band(0)._MapCol(2)._RSIndex=   2
         _Band(0)._MapCol(3)._Name=   "ywy"
         _Band(0)._MapCol(3)._RSIndex=   3
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgLj 
         Height          =   3975
         Left            =   -75000
         TabIndex        =   40
         Top             =   0
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   7011
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
   End
   Begin VB.CommandButton cmdBjd 
      BackColor       =   &H00C0C0FF&
      Caption         =   "报价单"
      Height          =   345
      Left            =   14280
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6150
      Width           =   915
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印"
      Height          =   405
      Left            =   12600
      TabIndex        =   35
      Top             =   8280
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CommandButton cmdRight 
      Caption         =   ">"
      Height          =   285
      Left            =   14760
      TabIndex        =   33
      Top             =   8460
      Width           =   465
   End
   Begin VB.CommandButton cmdLeft 
      Caption         =   "<"
      Height          =   285
      Left            =   14280
      TabIndex        =   32
      Top             =   8460
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
   Begin VB.CommandButton cmdMod 
      Height          =   375
      Left            =   13320
      Picture         =   "frmWBXJ.frx":047B
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "修改"
      Top             =   8790
      Width           =   465
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   13800
      Picture         =   "frmWBXJ.frx":0785
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "保存"
      Top             =   8790
      Width           =   465
   End
   Begin VB.CommandButton cmdQm 
      Caption         =   "cmdQm"
      Height          =   345
      Index           =   0
      Left            =   9150
      TabIndex        =   18
      Top             =   8430
      Width           =   945
   End
   Begin VB.CommandButton cmdD 
      Caption         =   "删除"
      Height          =   315
      Left            =   13140
      TabIndex        =   17
      Top             =   4770
      Width           =   975
   End
   Begin VB.Frame frmTime 
      Height          =   1185
      Left            =   270
      TabIndex        =   12
      Top             =   7980
      Width           =   2955
      Begin VB.CheckBox chkBa 
         Caption         =   "24小时运转"
         Height          =   255
         Left            =   150
         TabIndex        =   15
         Top             =   330
         Width           =   1215
      End
      Begin VB.CheckBox chkBb 
         Caption         =   "全年运转"
         Height          =   255
         Left            =   150
         TabIndex        =   14
         Top             =   645
         Width           =   1845
      End
      Begin VB.CheckBox chkBc 
         Caption         =   "2小时内到场"
         Height          =   255
         Left            =   150
         TabIndex        =   13
         Top             =   960
         Width           =   1845
      End
      Begin VB.Label Label13 
         Caption         =   "时间系数:"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   90
         Width           =   1155
      End
   End
   Begin VB.TextBox txtZt 
      Height          =   315
      Left            =   810
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4350
      Visible         =   0   'False
      Width           =   1005
   End
   Begin MSDataListLib.DataCombo comZu 
      Height          =   330
      Left            =   1380
      TabIndex        =   7
      Top             =   5925
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   582
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo comXmmc 
      Height          =   330
      Left            =   1380
      TabIndex        =   3
      Top             =   5100
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   582
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.CommandButton cmdJi 
      Caption         =   "计算"
      Height          =   375
      Left            =   14250
      TabIndex        =   1
      Top             =   8070
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Height          =   375
      Left            =   14790
      Picture         =   "frmWBXJ.frx":0DEF
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "返回"
      Top             =   8790
      Width           =   465
   End
   Begin VB.CommandButton cmdBjxt 
      BackColor       =   &H00FFC0C0&
      Caption         =   "询价系统"
      Height          =   375
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   4590
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdTk 
      Caption         =   "维保条款"
      Height          =   345
      Left            =   14220
      TabIndex        =   73
      Top             =   2970
      Width           =   975
   End
   Begin VB.CommandButton cmdCg 
      Caption         =   "采购询价"
      Height          =   345
      Left            =   14280
      TabIndex        =   41
      Top             =   3990
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblHLC 
      Caption         =   "lblHLC"
      Height          =   345
      Left            =   0
      TabIndex        =   111
      Top             =   0
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
      Height          =   345
      Left            =   8490
      TabIndex        =   97
      Top             =   7620
      Width           =   5475
   End
   Begin VB.Label lblHtbh 
      Caption         =   "对应合同"
      Height          =   255
      Left            =   0
      TabIndex        =   91
      Top             =   0
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label23 
      Caption         =   "备注"
      Height          =   225
      Left            =   4710
      TabIndex        =   85
      Top             =   6660
      Width           =   495
   End
   Begin VB.Label lblZl 
      Caption         =   "Label19"
      ForeColor       =   &H00C000C0&
      Height          =   225
      Left            =   1410
      TabIndex        =   59
      Top             =   4800
      Width           =   1155
   End
   Begin VB.Label lblzlZ 
      Caption         =   "性质"
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   630
      TabIndex        =   58
      Top             =   4800
      Width           =   435
   End
   Begin VB.Label lblCgid 
      Caption         =   "lblCgid"
      Height          =   285
      Left            =   12630
      TabIndex        =   42
      Top             =   8850
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblBaoId 
      Caption         =   "lblBaoId"
      Height          =   285
      Left            =   10560
      TabIndex        =   37
      Top             =   8190
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblOid 
      Caption         =   "lblOid"
      Height          =   285
      Left            =   10500
      TabIndex        =   23
      Top             =   8880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   0
      Left            =   9150
      TabIndex        =   20
      Top             =   8820
      Width           =   945
   End
   Begin VB.Label lblQM 
      Caption         =   "lblQm"
      Height          =   225
      Index           =   0
      Left            =   9180
      TabIndex        =   19
      Top             =   8130
      Width           =   1005
   End
   Begin VB.Label Label9 
      Caption         =   "总工时"
      Height          =   255
      Left            =   150
      TabIndex        =   10
      Top             =   4410
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblBid 
      Caption         =   "lblBid"
      Height          =   255
      Left            =   11490
      TabIndex        =   9
      Top             =   8910
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label8 
      Caption         =   "组长"
      Height          =   225
      Left            =   660
      TabIndex        =   8
      Top             =   6420
      Width           =   465
   End
   Begin VB.Label Label7 
      Caption         =   "工程部组号"
      Height          =   225
      Left            =   90
      TabIndex        =   6
      Top             =   6000
      Width           =   945
   End
   Begin VB.Label lblBh 
      BackColor       =   &H80000014&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      Height          =   330
      Left            =   1380
      TabIndex        =   5
      Top             =   5505
      Width           =   1725
   End
   Begin VB.Label Label5 
      Caption         =   "编号"
      Height          =   225
      Left            =   630
      TabIndex        =   4
      Top             =   5550
      Width           =   435
   End
   Begin VB.Label Label4 
      Caption         =   "项目名称"
      Height          =   225
      Left            =   270
      TabIndex        =   2
      Top             =   5160
      Width           =   975
   End
End
Attribute VB_Name = "frmWBXJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adoPb As New ADODB.Recordset
Public adoJz As New ADODB.Recordset
Public adoWb As New ADODB.Recordset
Public adoLj As New ADODB.Recordset
Public adoXm As New ADODB.Recordset
Public adoZu As New ADODB.Recordset
Public adoOid As New ADODB.Recordset '计算Old单子的ADO
Public JZ As Integer '基准价比例


Dim Jall As Boolean  '选择全部还是逐条选择(关系到计算方式)
Public adoA As Object '机组的ADO
Dim JxId As Long
Public ZF As Boolean '是否回驳过单据

Dim timZm As Integer '(8删除)

Private Sub chkJ6_Click()
If chkJ6.Value = 0 Then
    txtJ3.Text = ""
End If
End Sub

Private Sub chkJ7_Click()
If chkJ7.Value = 0 Then
    txtJ5.Text = ""
End If
End Sub


Private Sub cmdBack_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next

Call modBJD.BJDWBQing
frmWBXJ.Visible = False
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
Dim TP, TL
Dim N1: Dim N2: Dim N3: Dim N4: Dim N5: Dim N6: Dim N7: Dim N8: Dim N9: Dim N10
'n1
If comJ2.Text = "USRT" Then
    N1 = Int(Val(txtJ1.Text) / 200)
    If Val(txtJ1.Text) > 1000 Then
        N1 = 5
    End If
ElseIf comJ2.Text = "KW" Then
    N1 = Int(Val(txtJ1.Text) / (200 * 3.516))
    If Val(txtJ1.Text) > (1000 * 3.516) Then
        N1 = 5
    End If
End If
'n2
If chkJ6.Value = 1 Then
    TP = 1
Else
    TP = 0
End If
If chkJ7.Value = 1 Then
    TL = 1
Else
    TL = 0
End If
    N2 = Val(txtJ3.Text) * TP + Val(txtJ5.Text) * TL
 'n3
 If optJ8.Value = True Then
    N3 = 1
 Else
    N3 = 2
 End If
 'n4
 If chk10.Value = 1 And chk11.Value = 0 Or chk10.Value = 0 And chk11.Value = 1 Or chk10.Value = 0 And chk11.Value = 0 Then
    N4 = 1
Else
    N4 = 2
End If
'n5
N5 = Val(txt12.Text)
'n6
N6 = Val(com13.Text)
'n7
N7 = Val(txtJSL.Text)
'n9
If opt15.Value = True Then
    N9 = 1
Else
    N9 = 0
End If
End Sub

Private Sub cmdBjd_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
Call modBJD.BaoJDWBQing
'检查有无报价单,若没有则添加,若有则打开.
tt = "Select bid from BaoJIaD where bid =" & Val(lblBid.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.HTP.RecordCount = 0 Then
'    Exit Sub
'    If cmdRight.Enabled = True Then
'        MsgBox "当前记录不是最终有效询价单,故不能生成新报价单"
'        Exit Sub ''如果不是最终有效询价单,则不能生成新报价单
'    End If
    If lblYwy.Caption <> mod1.DName Then
        MsgBox "必须由业务员亲自生成报价单!"
        Exit Sub
    End If
    ii = MsgBox("是否生成新报价单?", vbQuestion + vbYesNo, "您辛苦了!")
    If ii = vbNo Then
        Exit Sub
    End If
    frmWbxjB.Visible = False
    mod1.BTZ = 37
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "BJDadd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@xmmc") = Trim(comXmmc.Text)
    mod1.cmd.Parameters("@xid") = Trim(comXmmc.Tag)
    mod1.cmd.Parameters("@ywy") = Trim(lblYwy.Caption)
    mod1.cmd.Parameters("@uid") = Trim(lblUid.Caption)
    mod1.cmd.Parameters("@lx") = 1
    mod1.cmd.Parameters("@zh") = comZu.Text
    mod1.cmd.Parameters("@zName") = Trim(txtZu.Text)
    If lblZl.Caption = "工程分包" Then
        mod1.cmd.Parameters("@zh") = 100
        mod1.cmd.Parameters("@zName") = "无"
    End If
'    mod1.CMD.Parameters("@jzPb") = Trim(comPb.Text)
'    mod1.CMD.Parameters("@jzxh") = Trim(comXh.Text)
'    mod1.CMD.Parameters("@sl") = Val(txtSl.Text)
    mod1.cmd.Parameters("@ta") = chkBa.Value
    mod1.cmd.Parameters("@tb") = chkBb.Value
    mod1.cmd.Parameters("@tc") = chkBc.Value
    mod1.cmd.Parameters("@ztime") = Val(txtZt.Text)
    mod1.cmd.Parameters("@yhg") = Val(txtYhg.Text)
'    mod1.CMD.Parameters("@nlb") = cmdBjd.Tag
'    mod1.CMD.Parameters("@lcou") = Right(cmdBjd.ToolTipText, 1)
    If lblZl.Caption = "维保" Then
        mod1.cmd.Parameters("@nlb") = 52
        mod1.cmd.Parameters("@lcou") = 4
    Else
        mod1.cmd.Parameters("@nlb") = 61
        mod1.cmd.Parameters("@lcou") = 3
    End If
    mod1.cmd.Parameters("@bid") = Val(lblBid.Caption)
    mod1.cmd.Parameters("@zl") = lblZl.Caption
    mod1.cmd.Parameters("@clf") = Val(txtClf.Text)
    mod1.cmd.Parameters("@rgf") = Val(txtYhg.Text) - Val(txtClf.Text)
    'mod1.CMD.Parameters("@clcb") = Val(frmGXBj.txtYhg.Text)
    mod1.cmd.Parameters("@mon") = Val(txtMon.Text)
    mod1.cmd.Parameters("@dxnr") = Trim(txtDxnr.Text)
    mod1.cmd.Parameters("@wc") = Val(txtWc.Text)
    mod1.cmd.Parameters("@xc") = Val(txtXc.Text)
    mod1.cmd.Parameters("@cgid") = Val(lblCgid.Caption)
    mod1.cmd.Parameters("@bz") = Trim(txtBz.Text)
    mod1.cmd.Parameters("@fbje") = Val(txtFbje.Text)
    mod1.cmd.Parameters("@fbnr") = Trim(txtFbnr.Text)
    mod1.cmd.Parameters("@yf") = 0
    mod1.cmd.Execute
 
          
    lblBaoId.Caption = mod1.cmd.Parameters("@baoid").Value
    frmWbxjB.lblBaoId.Caption = mod1.cmd.Parameters("@baoid").Value
    Set cmd = Nothing
    Call modBJD.BaoJDWBQing
    Call modBJD.BaoJDBound(Val(lblBaoId.Caption), lblZl.Caption)
    tt = "select * from baojiaOld where old=" & Val(frmWbxjB.lblOid.Caption) & " order by baoid"
    frmWbxjB.adoOid.Close
    frmWbxjB.adoOid.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    If frmWbxjB.adoOid.RecordCount > 1 Then
        frmWbxjB.cmdLeft.Enabled = True
    End If
    frmWbxjB.adoOid.MoveLast
    If frmWbxjB.lblZl.Caption = "维保" Then
        '设置流程按钮
        Call modBJD.BjWBLcBut(52)
    Else
        Call modBJD.BjWBLcBut(61)
    End If
    frmWbxjB.txtYf.Locked = False
    frmWbxjB.txtXm2.Locked = False
    frmWbxjB.txtHg.Locked = False
    frmWbxjB.txtYhg.Locked = False
    frmWbxjB.cmdCong.Visible = False
    frmWbxjB.Visible = True
    frmWbxjB.cmdPrint.Visible = False
    frmWbxjB.cmdSave.Enabled = True
    frmWbxjB.cmdMod.Enabled = False
    frmWbxjB.frmNb.Enabled = True
    frmWbxjB.dt3.Enabled = True
    frmWbxjB.dt4.Enabled = True
    
    If adoA.RecordCount = 0 Then
        frmWbxjB.comPb.Text = comPb.Text
        frmWbxjB.comXh.Text = comXh.Text
        frmWbxjB.txtSL.Text = txtSL.Text
        tt = "update baojiaD set jzpb='" & comPb.Text & "',jzxh='" & comXh.Text & "',sl=" & Val(txtSL.Text) & " where baoid=" & Val(frmWbxjB.lblBaoId.Caption)
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    End If
    frmWbxjB.txtSL.Locked = False
ElseIf mod1.HTP.RecordCount = 1 Then
    mod1.BTZ = 37
    Call modBJD.BaoJDWBQing
    Call modBJD.BaoJDBound(Val(lblBaoId.Caption), lblZl.Caption)
    frmWbxjB.Visible = True
    frmWbxjB.cmdSave.Enabled = False
    frmWbxjB.cmdMod.Enabled = True
End If
    frmWbxjB.frmYj.Visible = False
    frmWbxjB.frmYm.Visible = False

'frmWbxjB.txtHg.Locked = False
'frmWbxjB.txtCb.Text = txtYhg.Text '成本


frmWBXJ.Visible = False
frmWbxjB.txtFbje.Locked = False
frmWbxjB.txtFbnr.Locked = False
'frmWbxjB.cmdCong.Visible = True
End Sub

Private Sub cmdBjxt_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next

If comPb.Text = "" Or comXh.Text = "" Or txtSL.Text = "" Then
    MsgBox ("请选择机组信息")
    Exit Sub
End If
frmWbBj.Visible = False
ii = MsgBox("采用全部选择维保条款还是逐条添加?", vbInformation + vbYesNo, "全选否")
If ii = vbYes Then
    frmWBXJ.Enabled = False
    frmWait.Visible = True
    frmWait.ZOrder 0
    frmWait.Refresh
    Jall = True
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "xunJiaWbAddAll"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@xhid") = Trim(comXh.BoundText)
    mod1.cmd.Parameters("@bid") = Val(lblBid.Caption)
    mod1.cmd.Execute
    Set cmd = Nothing
    '年保表
    tt = "select * from xunJIaWbView where wbx='年保' and bid=" & Val(frmWBXJ.lblBid.Caption)
    frmWBXJ.adoWb.Close
    frmWBXJ.adoWb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmWBXJ.dtgWb.DataSource = frmWBXJ.adoWb
    frmWBXJ.dtgWb.FixedRows = 0
    frmWBXJ.dtgWb.MergeCol(1) = True
    frmWBXJ.dtgWb.MergeCol(2) = True
    frmWBXJ.dtgWb.MergeCol(3) = True
    frmWBXJ.dtgWb.MergeCells = 3
    frmWBXJ.dtgWb.FixedRows = 1
    '例检表
    tt = "select * from xunJIaWbView where wbx='例检' and bid=" & Val(frmWBXJ.lblBid.Caption)
    frmWBXJ.adoLj.Close
    frmWBXJ.adoLj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmWBXJ.dtgLj.DataSource = frmWBXJ.adoLj
    frmWBXJ.dtgLj.FixedRows = 0
    frmWBXJ.dtgLj.MergeCol(1) = True
    frmWBXJ.dtgLj.MergeCol(2) = True
    frmWBXJ.dtgLj.MergeCol(3) = True
    frmWBXJ.dtgLj.MergeCells = 3
    frmWBXJ.dtgLj.FixedRows = 1
    frmWait.Visible = False
    frmWBXJ.Enabled = True
    frmWBXJ.ZOrder 0
Else
    Jall = False
    tt = "select xt,xtid from bjxt_xt where xhid='" & comXh.BoundText & "' order by xtid"
    frmWbBj.adoBj.Close
    frmWbBj.adoBj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmWbBj.comLb.RowSource = frmWbBj.adoBj
    frmWbBj.comLb.ListField = "xt"
    frmWbBj.comLb.BoundColumn = "xtid"
    'frmWbBj.adoBj.MoveFirst
    'frmWbBj.comLb.Text = frmWbBj.adoBj.Fields("xt").Value
    frmWbBj.comLb.Text = ""

    frmWbBj.ZOrder 0
    frmWbBj.comPb.Text = comPb.Text
    frmWbBj.comXh.Text = comXh.Text
    frmWbBj.comPb.Enabled = False
    frmWbBj.comXh.Enabled = False
    frmWbBj.Show
End If
cmdSave.Enabled = True
cmdMod.Enabled = False
End Sub


Private Sub Command1_Click()

End Sub

Private Sub cmdCg_Click()
'再新建采购询价
'Dim tt As String
Dim LX As Boolean
Dim ii As Integer
'On Error Resume Next
'    frmGXBj.Show
'    frmGxBiao.Enabled = False
Dim tt As String
On Error Resume Next

If lblZl.Caption = "" Then Exit Sub
frmGXBj.Visible = False

frmWait.Show
frmWait.ZOrder 0
frmWait.Refresh

If Val(lblCgid.Caption) = 0 And lblYwy.Caption = mod1.DName Then            '如果没有建立采购询价单,而且为制单者本人,则新建
    Exit Sub
    ii = MsgBox("尚无采购询价记录,是否新建?", vbInformation + vbYesNo, "询问")
    If ii = vbNo Then
        frmWait.Visible = False
        Exit Sub
    End If

    '取得购销询价单的流程参数
    If frmGxBiao.cmdCreat.ToolTipText = "" Then
        tt = "xunJiaBut('" & mod1.DName & "','" & mod1.DHid & "','购销')"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
        frmGxBiao.cmdCreat.Tag = mod1.HTP.Fields("nlb").Value
        frmGxBiao.cmdCreat.ToolTipText = "流程总数为:" & mod1.HTP.Fields("lcou").Value
    End If
    Call modBJD.BJDGXQing
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "xunJiaAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@Lx") = 0
    mod1.cmd.Parameters("@zl") = lblZl.Caption
    mod1.cmd.Parameters("@Lcou") = Right(frmGxBiao.cmdCreat.ToolTipText, 1) '流程总数
    mod1.cmd.Parameters("@Lc") = 0 '当前流程
    mod1.cmd.Parameters("@lcRen") = mod1.DName
    mod1.cmd.Parameters("@lcUid") = mod1.DHid
    mod1.cmd.Parameters("@nLb") = frmGxBiao.cmdCreat.Tag
    mod1.cmd.Execute
    frmGXBj.lblBid.Caption = mod1.cmd.Parameters("@bid").Value
    'frmGXBj.lblBh.Caption = "XJD" & mod1.CMD.Parameters("@bid").Value
    frmGXBj.lblBh.Caption = frmWBXJ.lblBh.Caption  '在维保询价中,采购编号等于维保询价单编号
    frmGXBj.lblOid.Caption = mod1.cmd.Parameters("@bid").Value
    frmGXBj.lblLcou.Caption = Right(frmGxBiao.cmdCreat.ToolTipText, 1) '流程总数
    frmGXBj.lblLc.Caption = 0
    frmGXBj.lblLcRen.Caption = mod1.DName
    frmGXBj.lblLcUid.Caption = mod1.DHid
    frmGXBj.lblNlb.Caption = frmGxBiao.cmdCreat.Tag
    frmGXBj.lblYwy.Caption = mod1.DName
    frmGXBj.lblUid.Caption = mod1.DHid
    lblCgid.Caption = mod1.cmd.Parameters("@bid").Value    '更新原表的cgid值
    Set cmd = Nothing
    
    '更新原表的cgid值
    tt = "update xunjiaD set cgid=" & Val(frmGXBj.lblBid.Caption) & " where bid=" & Val(frmWBXJ.lblBid.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    '更新新建记录的wbid,xmmc,xid,bianhao值
    tt = "update xunjiaD set wbid=" & Val(frmWBXJ.lblBid.Caption) & ",xmmc='" & frmWBXJ.comXmmc.Text & "',xid=" & frmWBXJ.comXmmc.Tag & _
    ",bianhao='" & frmWBXJ.lblBh.Caption & "' where bid=" & Val(frmGXBj.lblBid.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    
    '设置项目名称信息
    frmGXBj.comXmmc.Text = comXmmc.Text
    frmGXBj.comXmmc.Tag = comXmmc.Tag
    frmGXBj.lblZl.Caption = lblZl.Caption
    tt = "update xunjiaD set xmmc='" & comXmmc.Text & "',xid=" & comXmmc.Tag & " where bid=" & Val(lblCgid.Caption)
    
    tt = "select jzpb,pbid from bjxt_jzpb"
    frmGXBj.adoPb.Close
    frmGXBj.adoPb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmGXBj.comJzpb.RowSource = frmGXBj.adoPb
    frmGXBj.comJzpb.ListField = "jzpb"
    frmGXBj.comJzpb.BoundColumn = "pbid"
    frmGXBj.txtHg.Locked = True
    frmGXBj.txtYhg.Locked = True
    
        '设置流程按钮
        Call modBJD.XJGXLcBut(43)
        

    '刷新购销列表
    tt = "select * from xunJIamxView where bid=" & Val(frmGXBj.lblBid.Caption)
        frmGXBj.adoGx.Close
        frmGXBj.adoGx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Set frmGXBj.dtgMa.DataSource = frmGXBj.adoGx
    
    frmGXBj.cmdSave.Enabled = True
    frmGxBiao.Enabled = False
    'frmGXBj.cmdBjd.Visible = False
    frmGXBj.txtYhg.Locked = True
ElseIf lblCgid.Caption <> "" And frmGXBj.comXmmc.Text = "" Then '打开已有的采购询价单
            Call modBJD.BJDGXQing


            Call modBJD.BJDGDBound(lblCgid.Caption)
            tt = "select bid from xunjiaOld where oid=" & Val(frmGXBj.lblOid.Caption) & " order by bid"
            frmGXBj.adoOid.Close
            frmGXBj.adoOid.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
            frmGXBj.adoOid.MoveLast
            If frmGXBj.adoOid.RecordCount > 1 Then
                frmGXBj.cmdRight.Enabled = False
                frmGXBj.cmdLeft.Enabled = True
            Else
                frmGXBj.cmdRight.Enabled = False
                frmGXBj.cmdLeft.Enabled = False
            End If
ElseIf lblCgid.Caption <> "" And frmGXBj.comXmmc.Text <> "" Then

End If
frmWBXJ.Visible = False
frmWait.Visible = False
frmGXBj.Visible = True
frmGXBj.frmCg.Enabled = False
frmGXBj.comXmmc.Locked = True
frmGXBj.lblZl.ForeColor = &HC000C0
frmGXBj.lblzlZ.ForeColor = &HC000C0
frmGXBj.comLx.Text = "零配件"
frmGXBj.comLx.Locked = False
End Sub

Private Sub cmdCong_Click()
Dim ii As Integer
Dim tt As String
Dim Bid As Long
Dim ZL As String
On Error Resume Next
'MsgBox "正在建设中!"
'Exit Sub



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

'


'


End Sub


Private Sub cmdD_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
If tabGc.Tab = 0 Then
    dtgWb.Col = 12
    If dtgWb.Text = "" Then
        MsgBox ("此为基本选项,不能删除!")
        Exit Sub
    End If
    dtgWb.Col = 18
ElseIf tabGc.Tab = 1 Then
    dtgLj.Col = 12
    If dtgLj.Text = "" Then
        MsgBox ("此为基本选项,不能删除!")
        Exit Sub
    End If
    dtgLj.Col = 18
End If
ii = MsgBox("是否确定删除此记录?", vbInformation + vbYesNo, "询问")
If ii = vbNo Then Exit Sub

If tabGc.Tab = 0 Then
    tt = "delete from xunjiawb where lid=" & Val(dtgWb.Text)
ElseIf tabGc.Tab = 1 Then
    tt = "delete from xunjiawb where lid=" & Val(dtgLj.Text)
End If
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText

If tabGc.Tab = 0 Then  '刷新年保表

    adoWb.Requery
    Set frmWBXJ.dtgWb.DataSource = frmWBXJ.adoWb

ElseIf tabGc.Tab = 1 Then                          '刷新例检表

        adoLj.Requery
    Set frmWBXJ.dtgLj.DataSource = frmWBXJ.adoLj
End If
    
End Sub

Private Sub cmdDel_Click()
Dim ii As Integer
Dim tt As String
On Error Resume Next
tt = "select htbh from htping where hid=" & Val(lblHtbh.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
If mod1.HTP.Fields(0).Value <> "HMNEW" Then
    Exit Sub
End If
mod1.HTP.Close
Set mod1.HTP = Nothing
If lblYwy.Caption <> mod1.DName Then Exit Sub
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
timZm = 7 '人工签字
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
    mod1.cmd.Parameters("@mt6") = comZu.Text  '组号
    mod1.cmd.Parameters("@mt7") = txtZu.Text '组长
    mod1.cmd.Parameters("@mt8") = ""
    mod1.cmd.Parameters("@mt9") = ""
    mod1.cmd.Parameters("@mt10") = ""
    mod1.cmd.Parameters("@mt11") = ""
    mod1.cmd.Parameters("@mt12") = ""
    mod1.cmd.Parameters("@mt13") = Trim(lblZl.Caption) '性质
    mod1.cmd.Parameters("@mt14") = Trim(lblFwid.Caption)
    mod1.cmd.Parameters("@mt15") = ""
    mod1.cmd.Parameters("@mt16") = ""
    mod1.cmd.Parameters("@mt17") = ""
    mod1.cmd.Parameters("@mt18") = ""
    mod1.cmd.Parameters("@mt19") = mod1.Qy
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
    mod1.cmd.Parameters("@mm2") = 0
    mod1.cmd.Parameters("@mm3") = 0
    mod1.cmd.Parameters("@mm4") = 0
    mod1.cmd.Parameters("@mm5") = 0
    mod1.cmd.Parameters("@mm6") = 0
    mod1.cmd.Parameters("@mm7") = 0
    mod1.cmd.Parameters("@mm8") = 0
    mod1.cmd.Parameters("@mm9") = 0
    mod1.cmd.Parameters("@mm10") = Val(txtHg.Text)
    mod1.cmd.Parameters("@mm11") = Val(txtClf.Text)
    mod1.cmd.Parameters("@mm12") = 0
    mod1.cmd.Parameters("@mm13") = 0
    mod1.cmd.Parameters("@mm14") = 0
    mod1.cmd.Parameters("@mm15") = 0
    mod1.cmd.Parameters("@mm16") = Val(txt2.Text) '基准价格
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

Private Sub cmdHt_Click()
If mod1.Bm = "工程部" Or mod1.Bm = "工程二部" Then
    MsgBox "哈哈！"
    MsgBox "你想干嘛？"
    Exit Sub
End If

If mod1.DName = "彭海翔" And lblYwy.Caption <> mod1.DName Then '彭海翔只能打开自己的合同
    MsgBox "哈哈！"
    MsgBox "你想干嘛？"
    Exit Sub
End If

If mod1.DName = "张砚纯" Then
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
End Sub

Private Sub cmdJadd_Click()
Dim tt As String
On Error Resume Next
If comPb.Text = "" Or comXh.Text = "" Or Val(txtSL.Text) = 0 Then Exit Sub
tt = "insert into wbjb (jzpb,jzxh,sl,bid) values ('" & comPb.Text & "','" & comXh.Text & "'," & Val(txtSL.Text) & "," & Val(lblBid.Caption) & ")"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
adoA.Requery
Set dtgA.DataSource = adoA


''''''''''''''添加维保条款
'''''''''''''Set mod1.cmd = createobject("adodb.command")
'''''''''''''mod1.cmd.ActiveConnection = mod1.CC
'''''''''''''mod1.cmd.CommandText = "xunJiaWbAddAll"
'''''''''''''mod1.cmd.CommandType = adCmdStoredProc
'''''''''''''If comPb.Text = "约克" Or comPb.Text = "顿汉布什" Or comPb.Text = "开利" Or comPb.Text = "麦克威尔" Or comPb.Text = "特灵" Then
'''''''''''''    mod1.cmd.Parameters("@xhid") = Trim(comXh.BoundText)
'''''''''''''Else
'''''''''''''    mod1.cmd.Parameters("@xhid") = 1
'''''''''''''End If
'''''''''''''mod1.cmd.Parameters("@bid") = Val(lblBid.Caption)
'''''''''''''mod1.cmd.Parameters("@jzpb") = comPb.Text
'''''''''''''mod1.cmd.Parameters("@jzxh") = comXh.Text
'''''''''''''mod1.cmd.Execute
'''''''''''''Set cmd = Nothing

'''''''''''''年保表
''''''''''''tt = "select * from xunJIaWbView where wbx='年保' and bid=" & Val(frmWBXJ.lblBid.Caption) & " and 机组品牌='" & comPb.Text & "' and 机组型号 like'%" & comXh.Text & "%'"
''''''''''''frmWBXJ.adoWb.Close
''''''''''''frmWBXJ.adoWb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
''''''''''''Set frmWBXJ.dtgWb.DataSource = frmWBXJ.adoWb
''''''''''''frmWBXJ.dtgWb.FixedRows = 0
''''''''''''frmWBXJ.dtgWb.MergeCol(1) = True
''''''''''''frmWBXJ.dtgWb.MergeCol(2) = True
''''''''''''frmWBXJ.dtgWb.MergeCol(3) = True
''''''''''''frmWBXJ.dtgWb.MergeCells = 3
''''''''''''frmWBXJ.dtgWb.FixedRows = 1
'''''''''''''例检表
''''''''''''tt = "select * from xunJIaWbView where wbx='例检' and bid=" & Val(frmWBXJ.lblBid.Caption) & " and 机组品牌='" & comPb.Text & "' and 机组型号 like '%" & comXh.Text & "%'"
''''''''''''frmWBXJ.adoLj.Close
''''''''''''frmWBXJ.adoLj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
''''''''''''Set frmWBXJ.dtgLj.DataSource = frmWBXJ.adoLj
''''''''''''frmWBXJ.dtgLj.FixedRows = 0
''''''''''''frmWBXJ.dtgLj.MergeCol(1) = True
''''''''''''frmWBXJ.dtgLj.MergeCol(2) = True
''''''''''''frmWBXJ.dtgLj.MergeCol(3) = True
''''''''''''frmWBXJ.dtgLj.MergeCells = 3
''''''''''''frmWBXJ.dtgLj.FixedRows = 1

End Sub

Private Sub cmdJdel_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
ii = MsgBox("是否确定删除此记录?", vbQuestion + vbYesNo, "您好")
If ii = vbYes Then
    tt = "delete from wbjb  where wid=" & JxId
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    adoA.Requery
    Set dtgA.DataSource = adoA

    
    tt = "delete from XunJiaMx where bid=" & Val(lblBid.Caption) & " and jzpb='" & comPb.Text & "' and jzxh='" & comXh.Text & "'"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    
    comPb.Text = ""
    comXh.Text = ""
    txtSL.Text = ""
End If
End Sub


Private Sub cmdJgx_Click()
Dim tt As String
On Error Resume Next
If comPb.Text <> comPb.ToolTipText Or comXh.Text <> comXh.ToolTipText Then
    MsgBox "不能更改机组品牌或机组型号,请先删除,再重新添加!"
    Exit Sub
End If
tt = "update wbjb set jzpb='" & comPb.Text & "',jzxh='" & comXh.Text & "',sl=" & Val(txtSL.Text) & " where wid=" & JxId
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
adoA.Requery
Set dtgA.DataSource = adoA
End Sub


Private Sub cmdJi_Click()
Dim oo As Integer
Dim ii As Integer
Dim tt As String
On Error Resume Next
Dim Zhg As Single '总总费用
Dim hg As Single '总费用
Dim ztT As Single '维保总正常
Dim ZT As Single '维保可选总工时
Dim ZBF As Single '总设备费
Dim XX As Single '时间系数
Dim JX As Single '机组系数
Dim LT As Single '例检总工时
'dtgWb.Col = 12
'dtgLj.Col = 12

If Val(txtWc.Text) = 0 Then
    MsgBox ("请输入维保年限!")
    txtWc.SetFocus
    Exit Sub
End If

If Val(txtXc.Text) = 0 Then
    MsgBox ("请输入例检次数!")
    txtXc.SetFocus
End If


frmWBXJ.Enabled = False
frmWait.Visible = True
frmWait.ZOrder 0
frmWait.Refresh
Zhg = 0
adoA.MoveFirst
Do While Not adoA.EOF
    comPb.Text = adoA.Fields("机组品牌").Value
    comXh.Text = adoA.Fields("机组型号").Value
    txtSL.Text = adoA.Fields("数量").Value
    '年保表
    tt = "select * from xunJIaWbView where wbx='年保' and bid=" & Val(frmWBXJ.lblBid.Caption) & " and 机组品牌='" & comPb.Text & "' and 机组型号 like '%" & comXh.Text & "%'"
    frmWBXJ.adoWb.Close
    frmWBXJ.adoWb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmWBXJ.dtgWb.DataSource = frmWBXJ.adoWb
    '例检表
    tt = "select * from xunJIaWbView where wbx='例检' and bid=" & Val(frmWBXJ.lblBid.Caption) & " and 机组品牌='" & comPb.Text & "' and 机组型号 like '%" & comXh.Text & "%'"
    frmWBXJ.adoLj.Close
    frmWBXJ.adoLj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmWBXJ.dtgLj.DataSource = frmWBXJ.adoLj
    
    
    hg = 0
    JX = 1
    If Val(txtSL.Text) = 2 Then
        JX = 0.9
    ElseIf Val(txtSL.Text) = 3 Then
        JX = 0.8
    ElseIf Val(txtSL.Text) = 4 Then
        JX = 0.8
    ElseIf Val(txtSL.Text) > 4 Then
        JX = 0.7
    End If
    XX = 1
    If chkBa.Value = 1 Then XX = XX + 0.05
    If chkBb.Value = 1 Then XX = XX + 0.05
    If chkBc.Value = 1 Then XX = XX + 0.05
    If XX > 1.2 Then XX = 1.2
    ztT = 0
    LT = 0
    ZT = 0
    ZBF = 0
    hg = 0

        
    adoWb.MoveFirst
    Do While Not adoWb.EOF
        If adoWb.Fields("可选项").Value = "可选" Then
            ZT = ZT + adoWb.Fields("工时").Value
            ZBF = ZBF + adoWb.Fields("设备费").Value
            If adoWb.Fields("fjl").Value = 0 Then
                ZT = ZT + adoWb.Fields("附加工时").Value
            End If
        Else
            ztT = ztT + adoWb.Fields("工时").Value
        End If
        
        adoWb.MoveNext
    Loop
        If ztT > 16 Then ztT = 16
    adoLj.MoveFirst
    Do While Not adoLj.EOF
        LT = LT + adoLj.Fields("工时").Value
        adoLj.MoveNext
    Loop
        If LT > 8 Then LT = 8
    hg = LT * Val(txtSL.Text) * Val(txtXc.Text) * JX * 30 + (ztT + ZT) * Val(txtSL.Text) * 1.5 * 30 + ZBF * Val(txtSL.Text)
    
   'txtZt.Text = ZT * XX
    Zhg = Zhg + hg
    adoA.MoveNext
Loop
txtHg.Text = Zhg
frmWait.Visible = False
frmWBXJ.Enabled = True
frmWBXJ.ZOrder 0
cmdSave.Enabled = True
cmdMod.Enabled = False
End Sub

Private Sub cmdLeft_Click()
Dim tt As String
On Error Resume Next
If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存!"
    Exit Sub
End If
Me.Enabled = False
frmWait.Show
frmWait.ZOrder
frmWait.Refresh
frmWBXJ.adoOid.MovePrevious
'打开新建单
Call modBJD.BJDWBQing
Call modBJD.BJDBound(frmWBXJ.adoOid.Fields("bid").Value, lblZl.Caption)
Call modBJD.wbxjLocked
frmWBXJ.cmdRight.Enabled = True
frmWBXJ.cmdBjd.Visible = False
'frmWBXJ.cmdCong.Visible = False
frmWBXJ.cmdCg.Visible = False
cmdMod.Enabled = False
cmdSave.Enabled = False
frmWBXJ.lblZl.ForeColor = &H80000012
frmWBXJ.lblzlZ.ForeColor = &H80000012
frmWBXJ.adoOid.MovePrevious
If frmWBXJ.adoOid.BOF = True Then
    cmdLeft.Enabled = False
Else
    cmdLeft.Enabled = True
End If
frmWBXJ.adoOid.MoveNext
frmWait.Visible = False
Me.Enabled = True
Me.ZOrder 0
End Sub

Private Sub cmdMod_Click()
If mod1.DName = "马晓聪" Then
    frmWBXJ.cmdD.Visible = True
    frmWBXJ.comPb.Locked = False
    frmWBXJ.comXh.Locked = False
    frmWBXJ.txtSL.Locked = False
    frmWBXJ.cmdJi.Visible = True
    frmWBXJ.txtMon.Locked = False
    frmWBXJ.cmdJadd.Enabled = True
    frmWBXJ.cmdJdel.Enabled = True
    frmWBXJ.cmdJgx.Enabled = True
    frmWBXJ.txtWc.Locked = False
    frmWBXJ.txtXc.Locked = False
    frmWBXJ.comZu.Locked = False
    frmWBXJ.txtDxnr.Locked = True
    frmWBXJ.txtBz.Locked = False
    cmdSave.Enabled = True
    txt1.Locked = False
    txt2.Locked = False
End If
If mod1.DName = "倪旭" Or mod1.DName = "张砚纯" Then
    cmdSave.Enabled = True
    txtZu.Locked = False
End If
If Val(lblLc.Caption) = 100 Then
    Exit Sub
End If
'cmdMod.Enabled = False
cmdSave.Enabled = True
If mod1.DName = lblYwy.Caption Then
    cmdDel.Enabled = True
End If
If lblLcRen.Caption = mod1.DName And lblLc.Caption = 1 Then
    frmWBXJ.cmdD.Visible = True
    frmWBXJ.comPb.Locked = False
    frmWBXJ.comXh.Locked = False
    frmWBXJ.txtSL.Locked = False
    frmWBXJ.cmdJi.Visible = True
    frmWBXJ.txtMon.Locked = False
    frmWBXJ.cmdJadd.Enabled = True
    frmWBXJ.cmdJdel.Enabled = True
    frmWBXJ.cmdJgx.Enabled = True
    frmWBXJ.txtWc.Locked = False
    frmWBXJ.txtXc.Locked = False
    frmWBXJ.comZu.Locked = False
    frmWBXJ.txtDxnr.Locked = True
    frmWBXJ.txtBz.Locked = False

ElseIf lblLcRen.Caption = mod1.DName And Val(lblLc.Caption) = 2 Then
    frmWBXJ.txtYhg.Locked = False
    frmWBXJ.txtClf.Locked = False
    frmWBXJ.txtFbje.Locked = False
    frmWBXJ.txtZu.Locked = False
'''    If lblZl.Caption = "维保" Then
'''        frmWBXJ.txtHg.Locked = False
'''        txtF1.Locked = False
'''
'''        If mod1.DName = "郑刚" Then
'''            txtF2.Locked = False
'''            txtF3.Locked = False
'''        ElseIf mod1.DName = "张寅" Then
'''            txtF2.Locked = False
'''            txtF3.Locked = False
'''        End If
'''    Else
'''        frmWBXJ.txtHg.Locked = True
'''        frmWBXJ.txtDxnr.Locked = False
'''        txtF1.Locked = False
'''        If mod1.DName = "郑刚" Then
'''            txtF2.Locked = False
'''            txtF3.Locked = False
'''            txtF4.Locked = False
'''        ElseIf mod1.DName = "张寅" Then
'''            txtF2.Locked = False
'''            txtF3.Locked = False
'''            txtF4.Locked = False
'''        End If
'''    End If
    frmWBXJ.txtHg.Locked = True
    frmWBXJ.txtF1.Locked = False
    frmWBXJ.txtF2.Locked = False
    frmWBXJ.txtF3.Locked = False
    frmWBXJ.txtF4.Locked = False
    frmWBXJ.txtDxnr.Locked = False
    frmWBXJ.txt1.Locked = False
    frmWBXJ.txt2.Locked = False
ElseIf Val(lblLc.Caption) = 3 Then
    frmWBXJ.txtZu.Locked = False
    frmWBXJ.txt1.Locked = False
    frmWBXJ.txt2.Locked = False
    txtDxnr.Locked = False
    txtBz.Locked = False
Else


End If
txtFbje.Locked = False
txtFbnr.Locked = False
txtZu.Locked = False
End Sub


Private Sub cmdOK_Click()
'If lblZl.Caption = "大修" Then
    txtHg.Text = Val(txtF1.Text) + Val(txtF2.Text) + Val(txtF3.Text) + Val(txtF4.Text) + Val(txtFbje.Text)
    txtYhg.Text = Val(txtHg.Text) + Val(txtClf.Text)
'End If
frmRG.Visible = False
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

Private Sub cmdQm_Click(Index As Integer)
Dim ii As String
Dim tt As String
Dim Tywy As String '单子流转到下一人的姓名
Dim Tuid As String
Dim Oywy As String '原来流转人的名字
Dim Ouid As String '原来流转人的工号

On Error Resume Next
cmdDing.Enabled = True
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

'''If txtZu.Text = "" And Val(lblLc.Caption) > 1 Then
'''    cmdSave.Enabled = True
'''    MsgBox "没有选择工程部组长！"
'''    Exit Sub
'''End If

frmQm.Visible = True
If lblLc.Caption = 1 Then
    optT2.Enabled = False
    OptT1.Value = True
Else
    optT2.Enabled = True
    OptT1.Value = False
    optT2.Value = False
End If
Exit Sub



If lblLc.Caption > 1 Then
    ii = MsgBox("您是否核准此单？(选择“是”将签字通过,选择“否”将驳回此单)", vbYesNoCancel + vbInformation, "请您注意!")
    If ii = vbNo Then
        ii = MsgBox("您是否要驳回前一位" & cmdQm(Index - 1).Caption & "的签名?", vbYesNo + vbInformation, "确认驳回吗?")
        If ii = vbNo Then
            Exit Sub
        End If
        tt = InputBox("请输入您要驳回的原因!")
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "xtzxFAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@yid").Value = 43 '反签名
        mod1.cmd.Parameters("@lc").Value = lblLc.Caption
        mod1.cmd.Parameters("@bh").Value = lblBid.Caption
        mod1.cmd.Parameters("@ywy").Value = mod1.DName
        mod1.cmd.Parameters("@uid").Value = mod1.DHid
        mod1.cmd.Parameters("@bz").Value = tt
        mod1.cmd.Parameters("@zn").Value = lblQM(Index).Caption '身份职能
        mod1.cmd.Execute
        Zid = mod1.cmd.Parameters("@Zid").Value
        Set cmd = Nothing
        cmdQm(Index - 1).Caption = ""
        lblTm(Index - 1).Caption = ""
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
'If cmdQm(Index).Caption <> "" Then Exit Sub
If txtHg.Text = "" And Val(lblLc.Caption) = 1 And lblZl.Caption = "维保" Then
    MsgBox "请确认金额!"
    Exit Sub
End If

If Val(txtYhg.Text) = 0 And Val(lblLc.Caption) > 1 And lblZl.Caption = "维保" Then
    MsgBox "请确认金额!"
    Exit Sub
End If
If Val(txtHg.Text) = 0 And Val(lblLc.Caption) = 2 And lblZl.Caption = "大修" Then
    MsgBox "请确认金额!"
    Exit Sub
End If

If Val(txtYhg.Text) = 0 And Val(lblLc.Caption) > 2 And lblZl.Caption = "大修" Then
    MsgBox "请确认金额!"
    Exit Sub
End If



    lblLc.Caption = lblLc.Caption + 1
If lblZl.Caption = "工程分包" Then
    lblLc.Caption = 5
'    comZu.Text = ""
'    txtZu.Text = ""
End If
    
    '更新表xunjiaD中的lcRen,lcUid 字段,以及QMRZ表中的相应字段.
                Set mod1.cmd = CreateObject("adodb.command")
                mod1.cmd.ActiveConnection = mod1.cc
                mod1.cmd.CommandText = "QMRZXJ"
                mod1.cmd.CommandType = adCmdStoredProc
                mod1.cmd.Parameters("@nlb") = 44 '单子(报销单)种类
                mod1.cmd.Parameters("@lc") = lblLc.Caption  '当前流程
                mod1.cmd.Parameters("@Dname") = mod1.DName
                mod1.cmd.Parameters("@uid") = mod1.DHid
                mod1.cmd.Parameters("@btz") = mod1.BTZ '业务属性
                mod1.cmd.Parameters("@zid") = cmdQm(Index).Tag '流程顺序
                mod1.cmd.Parameters("@Qdbh") = lblBid.Caption   '单子编号
                mod1.cmd.Parameters("@pje") = ""   '评审建议
                mod1.cmd.Parameters("@bm") = mod1.Bm
                mod1.cmd.Parameters("@ZH") = comZu.Text  '组号
                mod1.cmd.Parameters("@Zname") = txtZu.Text '组长
                mod1.cmd.Parameters("@ywy") = lblYwy.Caption
                mod1.cmd.Parameters("@comId") = mod1.comId '公司
                mod1.cmd.Execute
                Tywy = mod1.cmd.Parameters("@Tywy").Value
                Tuid = mod1.cmd.Parameters("@Tuid").Value
                Set mod1.cmd = Nothing
                cmdQm(Index).Caption = mod1.DName
                lblTm(Index).Caption = mod1.DQda
                lblLcRen.Caption = Tywy
                lblLcUid.Caption = Tuid
                

If Val(lblLc.Caption) > Val(lblLcou.Caption) Then
    Call mod1.EnventFinish(frmWBXJ.lblFwid.Caption)
    tt = "update xunJiaD set Pwf=1 where bid=" & Val(lblBid.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    MsgBox "确认了这份询价单后,您的实际人工成本就产生了!"

If lblZl.Caption = "维保" Then
    tt = "update htping set w11=" & Val(txtYhg.Text) & " where hid=" & Val(lblHtbh.Caption)
    FMXC.txtH1.Text = txtYhg.Text
ElseIf lblZl.Caption = "大修" Then
    tt = "update htping set w22=" & Val(txtYhg.Text) & " where hid=" & Val(lblHtbh.Caption)
        FMXC.txtH2.Text = txtYhg.Text
ElseIf lblZl.Caption = "工程分包" Then
    tt = "update htping set w33=" & Val(txtYhg.Text) & " where hid=" & Val(lblHtbh.Caption)
End If
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText


Else
'    If lblLc.Caption = 1 Then '业务员第一个签字,则询价日期等于签字日期
'
'    End If
    '添加事务
    Call mod1.EnventAdd("询价单", comXmmc.Text, lblLcRen.Caption, lblLcUid.Caption, lblBid.Caption, lblQM(Index + 1).Caption, Oywy, Ouid, lblYwy.Caption, lblUid.Caption, lblFwid.Caption, lblBid.Caption)
    MsgBox "现在,此询价单将交由 " & Tywy & " 来审阅!"
End If

If Dialog.Visible = True Then '更新事务列表
    Call mod1.refEnvent(1)
End If
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

Private Sub cmdRight_Click()
Dim tt As String
On Error Resume Next
Me.Enabled = False
frmWait.Show
frmWait.ZOrder
frmWait.Refresh
frmWBXJ.adoOid.MoveNext
'打开新建单
Call modBJD.BJDWBQing
Call modBJD.BJDBound(frmWBXJ.adoOid.Fields("bid").Value, lblZl.Caption)
Call modBJD.wbxjLocked
frmWBXJ.cmdLeft.Enabled = True
cmdMod.Enabled = False
cmdSave.Enabled = False
cmdBjd.Visible = False
cmdCong.Visible = False
cmdCg.Visible = False
frmWBXJ.adoOid.MoveNext
If frmWBXJ.adoOid.EOF = True Then
    cmdRight.Enabled = False
    frmWBXJ.lblZl.ForeColor = &HC000C0
    frmWBXJ.lblzlZ.ForeColor = &HC000C0
    If mod1.Bm = lblBM.Caption And mod1.BmJl = True Or mod1.DName = lblYwy.Caption Or (mod1.DName = "宋晓炯" Or mod1.DName = "宋晓炯1" Or mod1.DName = "倪旭") Then
        frmWBXJ.cmdCg.Visible = True
        If mod1.DName = lblYwy.Caption And lblPwf.Caption = 1 Then
            'cmdBjd.Visible = True
        End If
    Else
        frmWBXJ.cmdCg.Visible = False
    End If
    frmWBXJ.cmdMod.Enabled = True
    If mod1.DName = lblYwy.Caption Then
        cmdCong.Visible = True
    End If
Else
    cmdRight.Enabled = True
End If
frmWBXJ.adoOid.MovePrevious
frmWait.Visible = False
Me.Enabled = True
Me.ZOrder 0
End Sub

Private Sub cmdSave_Click()
Dim tt As String
On Error Resume Next
'''''If lblZl.Caption = "维保" And txtF1.Text = "" Then
'''''    txtF1.Text = Val(txtHg.Text) * 0.55
'''''    txtF2.Text = Val(txtHg.Text) * 0.225
'''''    txtF3.Text = Val(txtHg.Text) * 0.225
'''''Else
    txtHg.Text = Val(txtF1.Text) + Val(txtF2.Text) + Val(txtF3.Text) + Val(txtF4.Text) + Val(txtFbje.Text)
    txtYhg.Text = Val(txtHg.Text) + Val(txtClf.Text)
'''''End If
'''''''''''''If Val(lblLc.Caption) = 1 Then
'''''''''''''    If lblZl.Caption = "维保" Or lblZl.Caption = "大修" Then
'''''''''''''            If mod1.Qy = "上海" Or mod1.Qy = "广州" Then
'''''''''''''                txtZu.Text = "张寅"
'''''''''''''            ElseIf mod1.Qy = "杭州" Or mod1.Qy = "南京" Then
'''''''''''''                txtZu.Text = "吴胜明"
'''''''''''''            ElseIf mod1.Qy = "北京" Then
'''''''''''''                txtZu.Text = "曹松"
'''''''''''''            End If
'''''''''''''    Else
'''''''''''''        txtZu.Text = "潘明峰"
'''''''''''''    End If
'''''''''''''End If
'''''''If mod1.DName = "倪旭" Or mod1.DName = "张砚纯" Then
'''''''    tt = "update newfuwu set cf=1 where fwid=" & Val(lblFwid.Caption)
'''''''    Set mod1.HTP = CreateObject("adodb.recordset")
'''''''    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'''''''    Call mod1.refEnvent(1)
'''''''End If

Me.Enabled = False
frmWait.Visible = True
frmWait.ZOrder 0
cmdMod.Enabled = True
cmdSave.Enabled = False
tt = "select * from XunJiaD where bid=" & Val(lblBid.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
mod1.HTP.Update "xmmc", Trim(comXmmc.Text)    '项目名称
mod1.HTP.Update "xid", comXmmc.Tag '项目代号
mod1.HTP.Update "bianhao", lblBh.Caption '单子编号(给用户看的)
mod1.HTP.Update "zh", comZu.Text        '组号
mod1.HTP.Update "Zname", Trim(txtZu.Text)     '组长
'mod1.HTP.Update "jzpb", Trim(comPb.Text)
'mod1.HTP.Update "jzxh", Trim(comXh.Text)
'mod1.HTP.Update "sl", Val(txtSl.Text)
mod1.HTP.Update "ta", chkBa.Value   '时间系数
mod1.HTP.Update "tb", chkBb.Value
mod1.HTP.Update "tc", chkBc.Value
mod1.HTP.Update "zTime", Val(txtZt.Text) '总工时
mod1.HTP.Update "hg", Val(txtHg.Text) '总费用
mod1.HTP.Update "yhg", Val(txtYhg.Text) '优惠价
mod1.HTP.Update "clf", Val(txtClf.Text) '差旅费
mod1.HTP.Update "wc", Val(txtWc.Text)
mod1.HTP.Update "xc", Val(txtXc.Text)
mod1.HTP.Update "dxnr", Trim(txtDxnr.Text)
mod1.HTP.Update "mon", Val(txtMon.Text)
mod1.HTP.Update "f1", Val(txtF1.Text)
mod1.HTP.Update "f2", Val(txtF2.Text)
mod1.HTP.Update "f3", Val(txtF3.Text)
mod1.HTP.Update "bz", Trim(txtBz.Text)
mod1.HTP.Update "fbje", Val(txtFbje.Text)
mod1.HTP.Update "fbnr", Trim(txtFbnr.Text)
If Val(lblBid.Caption) >= 6794 Then
    mod1.HTP.Update "jhg", Val(txt2.Text)
    mod1.HTP.Update "hg", Val(txt1.Text) '总费用
    mod1.HTP.Update "yhg", Val(txt1.Text) '优惠价
End If
mod1.HTP.UpdateBatch

If lblFwid.Caption = "" Then
    lblLc.Caption = 1
    tt = "update xunJiaD set lc=1 where bid=" & Val(lblBid.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    '添加事务
    Call mod1.EnventAdd("询价单", comXmmc.Text, lblLcRen.Caption, lblLcUid.Caption, lblBid.Caption, lblQM(0).Caption, "", "", mod1.DName, mod1.DHid, 0, lblBid.Caption)
'    '更新按钮
'    Call modBJD.OpenXJAN(1)
End If




'更新询价列表
'tt = "select * from xunjiaView where ywy='" & mod1.DName & "' and uid='" & mod1.DHid & "'"
'frmGxBiao.adoXj.Close
'frmGxBiao.adoXj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
frmGxBiao.adoXj.Requery
Set frmGxBiao.dtgXj.DataSource = frmGxBiao.adoXj

frmWait.Visible = False
Me.Enabled = True
Me.ZOrder 0

cmdCg.Enabled = True






End Sub

Private Sub cmdTK_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
If comPb.Text = "" Or comXh.Text = "" Or Val(txtSL.Text) = 0 Then Exit Sub
'年保表
tt = "select * from xunJIaWbView where wbx='年保' and bid=" & Val(frmWBXJ.lblBid.Caption) & " and 机组品牌='" & comPb.Text & "' and 机组型号 like '%" & comXh.Text & "%'"
frmWBXJ.adoWb.Close
frmWBXJ.adoWb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmWBXJ.dtgWb.DataSource = frmWBXJ.adoWb
frmWBXJ.dtgWb.FixedRows = 0
frmWBXJ.dtgWb.MergeCol(1) = True
frmWBXJ.dtgWb.MergeCol(2) = True
frmWBXJ.dtgWb.MergeCol(3) = True
frmWBXJ.dtgWb.MergeCells = 3
frmWBXJ.dtgWb.FixedRows = 1
'例检表
tt = "select * from xunJIaWbView where wbx='例检' and bid=" & Val(frmWBXJ.lblBid.Caption) & " and 机组品牌='" & comPb.Text & "' and 机组型号 like '%" & comXh.Text & "%'"
frmWBXJ.adoLj.Close
frmWBXJ.adoLj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmWBXJ.dtgLj.DataSource = frmWBXJ.adoLj
frmWBXJ.dtgLj.FixedRows = 0
frmWBXJ.dtgLj.MergeCol(1) = True
frmWBXJ.dtgLj.MergeCol(2) = True
frmWBXJ.dtgLj.MergeCol(3) = True
frmWBXJ.dtgLj.MergeCells = 3
frmWBXJ.dtgLj.FixedRows = 1
End Sub

Private Sub comDX_Click()

frmM1.Visible = False
frmM2.Visible = False
frmM3.Visible = False
frmM5.Visible = False
frmNewF.Visible = False
Select Case comDX.Text
Case "主机"
    frmM1.Visible = True
    frmXH.Visible = False
frmNewF.Visible = True
Case "水泵"
    frmM2.Visible = True
Case "电机"
    frmM3.Visible = True
Case "小机"
    frmM5.Visible = True
Case "小机安装"
    frmM5.Visible = True
Case "风机盘管"
    frmM5.Visible = True
Case "空调箱"
    frmM5.Visible = True
End Select








End Sub


Private Sub comPb_Change()
Dim tt As String
On Error Resume Next

If frmWBXJ.Visible = False Then Exit Sub
    If comPb.Text = "约克" Or comPb.Text = "顿汉布什" Or comPb.Text = "开利" Or comPb.Text = "麦克威尔" Or comPb.Text = "特灵" Then
        tt = "select * from bjxt_jzxh where pbid='" & frmWBXJ.comPb.BoundText & "'"
    Else
        tt = "select * from bjxt_jzxh where pbid=1"
    End If
    frmWBXJ.adoJz.Close
    frmWBXJ.adoJz.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmWBXJ.comXh.RowSource = frmWBXJ.adoJz
    frmWBXJ.comXh.ListField = "jzxh"
    frmWBXJ.comXh.BoundColumn = "xhid"
    frmWBXJ.adoJz.MoveFirst
    frmWBXJ.comXh.Text = frmWBXJ.adoJz.Fields("jzxh").Value
End Sub

Private Sub comXmmc_Click(Area As Integer)
comXmmc.Tag = comXmmc.BoundText
End Sub

Private Sub comZu_Change()
'If frmWBXJ.Visible = True Then
'If comZu.Text = "100" And lblZl.Caption <> "工程分包" Or comZu.Text <> "100" And lblZl.Caption = "工程分包" Then
'    comZu.Text = ""
'    Exit Sub
'
'End If
'    txtZu.Text = comZu.BoundText
'End If
'If txtZu.Text = "冷群辉" Then
'    txtZu.Text = "吴胜明"
'End If
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
txtSL.Text = dtgA.Text
End Sub

Private Sub dtgA_RowColChange()
On Error Resume Next
dtgA.Col = 4
JxId = dtgA.Text
dtgA.Col = 1
comPb.Text = dtgA.Text
dtgA.Col = 2
comXh.Text = dtgA.Text
dtgA.Col = 3
txtSL.Text = dtgA.Text
End Sub


Private Sub dtgLj_DblClick()
Dim Wnr As String
Dim Dl As Single
Dim FJl As Boolean
Dim liD As Long
Dim tt As String
Dim ii As Single
On Error Resume Next
dtgLj.Col = 4
Wnr = dtgLj.Text
dtgLj.Col = 7
Dl = dtgLj.Text
dtgLj.Col = 14
FJl = dtgLj.Text
dtgLj.Col = 18
liD = dtgLj.Text

If FJl = True And (Dl = 0 Or IsNull(Dl) = True) Then
    ii = InputBox("请输入" & adoWb.Fields("wnr").Value & "条款中的附加量:")
    frmWBXJ.Enabled = False
    frmWait.Visible = True
    frmWait.ZOrder 0
    frmWait.Refresh
    tt = "update xunjiaWb set dl=" & ii & " where lid=" & liD
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    '例检表
'    tt = "select * from xunJIaWbView where wbx='例检' and bid=" & Val(frmWBXJ.lblBid.Caption)
'    frmWBXJ.adoLj.Close
'    frmWBXJ.adoLj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    adoLj.Requery
    Set frmWBXJ.dtgLj.DataSource = frmWBXJ.adoLj
'    frmWBXJ.dtgLj.FixedRows = 0
'    frmWBXJ.dtgLj.MergeCol(1) = True
'    frmWBXJ.dtgLj.MergeCol(2) = True
'    frmWBXJ.dtgLj.MergeCol(3) = True
'    frmWBXJ.dtgLj.MergeCells = 3
'    frmWBXJ.dtgLj.FixedRows = 1
    frmWait.Visible = False
    frmWBXJ.Enabled = True
End If
End Sub



Private Sub dtgWb_DblClick()
'Dim Orow As Long
'Dim OCol As Long
Dim Wnr As String
Dim Dl As Single
Dim FJl As Boolean
Dim liD As Long
Dim tt As String
Dim ii As Single
On Error Resume Next
dtgWb.Col = 4
Wnr = dtgWb.Text
dtgWb.Col = 7
Dl = dtgWb.Text
dtgWb.Col = 14
FJl = dtgWb.Text
dtgWb.Col = 18
liD = dtgWb.Text
'Orow = dtgWb.RowSel
'OCol = dtgWb.ColSel

If FJl = True And (Dl = 0 Or IsNull(Dl) = True) Then
    ii = InputBox("请输入" & Wnr & "条款中的附加量:")
    frmWBXJ.Enabled = False
    frmWait.Visible = True
    frmWait.ZOrder 0
    frmWait.Refresh
    tt = "update xunjiaWb set dl=" & ii & " where lid=" & liD
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'    '例检表
'    tt = "select * from xunJIaWbView where wbx='年保' and bid=" & Val(frmWBXJ.lblBid.Caption)
'    frmWBXJ.adoWb.Close
'    frmWBXJ.adoWb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    frmWBXJ.adoWb.Requery
    Set frmWBXJ.dtgWb.DataSource = frmWBXJ.adoWb
'    frmWBXJ.dtgWb.FixedRows = 0
'    frmWBXJ.dtgWb.MergeCol(1) = True
'    frmWBXJ.dtgWb.MergeCol(2) = True
'    frmWBXJ.dtgWb.MergeCol(3) = True
'    frmWBXJ.dtgWb.MergeCells = 3
'    frmWBXJ.dtgWb.FixedRows = 1
    frmWait.Visible = False
    frmWBXJ.Enabled = True
'    dtgWb.RowSel = Orow
'    dtgWb.ColSel = OCol
    
End If

End Sub


Private Sub Form_Click()
frmRG.Visible = False
End Sub

Private Sub Form_Load()
Dim tt As String
Dim Ra
Dim La
Dim oo As Integer
On Error Resume Next

Me.Left = 0
Me.Top = 0
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Set adoPb = CreateObject("adodb.recordset")
Set adoJz = CreateObject("adodb.recordset")
Set adoWb = CreateObject("adodb.recordset")
Set adoLj = CreateObject("adodb.recordset")
Set adoXm = CreateObject("adodb.recordset")
Set adoZu = CreateObject("adodb.recordset")
Set adoOid = CreateObject("adodb.recordset")
Set adoA = CreateObject("adodb.recordset")
frmNew.BorderStyle = 0
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
dtgA.ColWidth(0) = 300
dtgA.ColWidth(2) = 2000
dtgA.ColWidth(3) = 700
dtgA.ColWidth(4) = 0

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
frmDx.BorderStyle = 0
OptT1.Value = True
frmNb.BorderStyle = 0

''''If mod1.comId = 0 Then
''''    'tt = "select username,gzu from worker_gcz where zuf=1 or (username='郑刚') order by gzu"
''''    tt = "select username,gzu from worker_gcz where gzu< 10 order by gzu"
''''ElseIf mod1.comId = 1 Then
''''    tt = "select username,gzu from worker_gcz where zuf=1 and comid=1 order by gzu"
''''End If
''''adoZu.Close
''''
''''adoZu.Open tt, mod1.workFF, adOpenKeyset, adLockReadOnly, adCmdText
''''Set comZu.RowSource = adoZu
''''comZu.ListField = "gzu"
''''comZu.BoundColumn = "username"


tt = "select jzpb,pbid from bjxt_jzpb"
frmWBXJ.adoPb.Close
'基础发布
'Select Case mod1.Lqy
'Case "上海"
'    frmWBXJ.adoPb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'Case "杭州"
'    frmWBXJ.adoPb.Open tt, mod1.workHz, adOpenKeyset, adLockReadOnly, adCmdText
'End Select
frmWBXJ.adoPb.Open tt, mod1.workFF, adOpenKeyset, adLockReadOnly, adCmdText
Set frmWBXJ.comPb.RowSource = frmWBXJ.adoPb
frmWBXJ.comPb.ListField = "jzpb"
frmWBXJ.comPb.BoundColumn = "pbid"
frmTime.BorderStyle = 0
frmQm.Left = 7470
frmQm.Top = 7500


frmM1.Left = 4260: frmM2.Left = 4260: frmM3.Left = 4260: frmM5.Left = 4260
frmM1.Top = 1590: frmM2.Top = 1590: frmM3.Top = 1590: frmM5.Top = 1590
dtgJG.ColWidth(0) = 300: dtgJG.Cols = 12: dtgJG.Rows = 5
dtgJG.Row = 0: dtgJG.Col = 1: dtgJG.Text = "保养对象": dtgJG.Col = 2: dtgJG.Text = "品牌名称": dtgJG.Col = 3: dtgJG.Text = "型号"
dtgJG.Col = 4: dtgJG.Text = "系列编号": dtgJG.Col = 5: dtgJG.Text = "保养性质": dtgJG.Col = 6: dtgJG.Text = "人工价": dtgJG.Col = 7: dtgJG.Text = "基准价"
dtgJG.ColWidth(1) = 2000: dtgJG.ColWidth(2) = 2000: dtgJG.ColWidth(3) = 2500: dtgJG.ColWidth(4) = 2500:

tt = "select username from worker where zuf=1 and zzf=1"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2)
For oo = 0 To La
    frmWBXJ.txtZu.AddItem Ra(0, oo)
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim tt As String
Dim ii As Integer
On Error Resume Next
If MDI.Cq = False Then
If cmdSave.Enabled = True And cmdMod.Enabled = False Then
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
Cancel = True
Call modBJD.BJDWBQing
frmWBXJ.Visible = False
If frmGxBiao.Visible = True Then
    frmGxBiao.Enabled = True
    frmGxBiao.ZOrder 0
ElseIf Dialog.Visible = True Then
    Dialog.ZOrder 0
    Dialog.Enabled = True
End If
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




Private Sub Frame1_Click()
frmRG.Visible = False
End Sub

Private Sub lbl2_Click()
If mod1.Bm = "商务部" Or mod1.DName = "倪旭" Or mod1.DName = "宋晓炯" Or mod1.DName = "宋晓炯1" Or mod1.DName = "马晓聪" Or mod1.DName = "周春云" Then
    If lbl1.Visible = False Then
        lbl1.Visible = True
        txt1.Visible = True
    Else
        lbl1.Visible = False
        txt1.Visible = False
    End If
End If
End Sub

Private Sub tabGc_Click(PreviousTab As Integer)
dtgWb.Visible = False
dtgLj.Visible = False
If tabGc.Tab = 0 Then
    dtgWb.Visible = True
ElseIf tabGc.Tab = 1 Then
    dtgLj.Visible = True
End If
End Sub

Private Sub timQuit_Timer()
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
If timZm = 1 Then '如果为"决定何处购买"
    cmdJG.Enabled = False
ElseIf timZm = 2 Then
    cmdGx.Enabled = True
    If mod1.Bm <> "零件事业部" Then
        optW.Value = True
    End If
ElseIf timZm = 3 Or timZm = 6 Then       '添加配件,配件删除
    adoGx.Requery
    Set dtgMa.DataSource = adoGx
    If adoGx.RecordCount > 1 Then
    dtgMa.FixedRows = 0
    dtgMa.MergeCol(1) = True
    dtgMa.MergeCol(2) = True
    dtgMa.MergeCol(10) = True
    dtgMa.MergeCol(14) = True
    dtgMa.MergeCells = 3
    dtgMa.FixedRows = 1
    End If
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
    Set dtgMa.DataSource = adoGx
    
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
    txtGYmc.Text = ""
    txtGyman.Text = ""
    txtGyAdr.Text = ""
    txtGYPho.Text = ""
ElseIf timZm = 7 Then '签字
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
    End If
ElseIf timZm = 8 Then '删除
    Me.Visible = False
    If FMXC.Visible = True Then
''''''        If lblZl.Caption = "维保" Then
''''''            FMXC.cmdW1.ToolTipText = ""
''''''        ElseIf lblZl.Caption = "大修" Then
''''''            FMXC.cmdW2.ToolTipText = ""
''''''        ElseIf lblZl.Caption = "工程分包" Then
''''''            FMXC.cmdW3.ToolTipText = ""
''''''        ElseIf lblZl.Caption = "水处理" Then
''''''            FMXC.cmdW4.ToolTipText = ""
''''''        End If
    End If
    If Dialog.Visible = True Then
        Dialog.Enabled = True
        Dialog.ZOrder 0
    End If
End If
timQuit.Enabled = False
End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh,mm1,mm2,mt1,mt2 from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0
    If timZm = 2 Then
        If mod1.Bm = "零件事业部" Then
            txtHg.Text = mod1.WP.Fields("mm1").Value
            txtYhg.Text = txtHg.Text
            LBLhG.Caption = txtHg.Text
            LBLyHG.Caption = txtHg.Text
        Else
            txtHg.Text = mod1.WP.Fields("mm1").Value
            txtYhg.Text = txtHg.Text
            lblWhg.Caption = txtHg.Text
        End If
        adoGx.Requery
        Set dtgMa.DataSource = adoGx
    ElseIf timZm = 7 Then '签名
                If OptT1.Value = True Then
                    cmdQm(lblLc.Caption - 1).Caption = mod1.DName
                    lblTm(lblLc.Caption - 1).Caption = mod1.DQda
                Else
                    cmdQm(lblLc.Caption - 2).Caption = ""
                    lblTm(lblLc.Caption - 2).Caption = ""
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

Private Sub txt1_DblClick()
If (mod1.Bm = "工程部" Or mod1.DName = "倪旭") And Val(txtF1.Text) > 0 Then
    frmRG.Visible = True
End If
End Sub

Private Sub txt1_LostFocus()
'''''If lblZl.Caption = "维保" Then
'''''    txt2.Text = Val(txt1.Text) / (1 - mod1.JiZ1 / 100)
'''''ElseIf lblZl.Caption = "大修" Then
'''''    txt2.Text = Val(txt1.Text) / (1 - mod1.JiZ2 / 100)
'''''ElseIf lblZl.Caption = "工程分包" Then
'''''    txt2.Text = Val(txt1.Text) / (1 - mod1.JiZ3 / 100)
'''''ElseIf lblZl.Caption = "水处理" Then
'''''    txt2.Text = Val(txt1.Text) / (1 - mod1.JiZ4 / 100)
'''''End If
End Sub


Private Sub txtClf_LostFocus()
'If txtYhg.Text = "" Then
    txtYhg.Text = Val(txtHg.Text) + Val(txtClf.Text)
'End If
End Sub


Private Sub txtDxnr_Change()
If frmWBXJ.Visible = False Then Exit Sub
If Len(txtDxnr.Text) = 500 Then
    MsgBox ("您编辑文字数将超过此项目的最大容纳字数,多余文字将不被保存!")
End If
End Sub


Private Sub txtHg_DblClick()
If (mod1.Bm = "工程部" Or mod1.DName = "倪旭") And Val(txtF1.Text) > 0 Then
    frmRG.Visible = True
End If
End Sub

Private Sub txtHg_LostFocus()
txtYhg.Text = Val(txtHg.Text) + Val(txtClf.Text)
If lblZl.Caption = "维保" Then
    txtF1.Text = Val(txtHg.Text) * 0.55
    txtF2.Text = Val(txtHg.Text) * 0.225
    txtF3.Text = Val(txtHg.Text) * 0.225
End If
End Sub


