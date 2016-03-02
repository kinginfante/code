VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{EF977422-E047-42A7-A004-1C0695C81FCF}#1.0#0"; "NiceForm.ocx"
Begin VB.Form FmxcXJ 
   BackColor       =   &H00C0FFC0&
   Caption         =   "询价单"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   15210
   Begin VB.TextBox txtBJ 
      BackColor       =   &H00C0FFC0&
      Height          =   270
      Left            =   10140
      TabIndex        =   113
      Text            =   "Text1"
      Top             =   7320
      Width           =   1425
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgDW 
      Height          =   1815
      Left            =   4680
      TabIndex        =   109
      Top             =   720
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3201
      _Version        =   393216
      BackColor       =   16777152
      BackColorFixed  =   16777152
      BackColorBkg    =   16777152
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtBhg 
      BackColor       =   &H00FFFFC0&
      Height          =   270
      Left            =   10140
      TabIndex        =   101
      Text            =   "Text1"
      Top             =   7680
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   285
      Left            =   3090
      TabIndex        =   92
      Top             =   780
      Width           =   1605
      Begin VB.OptionButton Option1 
         Caption         =   "收缩"
         Height          =   255
         Left            =   750
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   0
         Width           =   855
      End
      Begin VB.OptionButton optV1 
         Caption         =   "展开"
         Height          =   255
         Left            =   -30
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   0
         Value           =   -1  'True
         Width           =   765
      End
   End
   Begin VB.Frame frmQm 
      BackColor       =   &H00C0FFC0&
      Caption         =   "评审建议"
      ForeColor       =   &H000000FF&
      Height          =   1785
      Left            =   2940
      TabIndex        =   0
      Top             =   7380
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
   Begin MSAdodcLib.Adodc adoFile 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Connect         =   "Provider=SQLOLEDB.1;Password=hugemanzou;Persist Security Info=True;User ID=zou;Initial Catalog=HMZou;Data Source=10.128.123.10"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=hugemanzou;Persist Security Info=True;User ID=zou;Initial Catalog=HMZou;Data Source=10.128.123.10"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "HMFile"
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
   Begin VB.CommandButton cmdTK 
      BackColor       =   &H00FFFFC0&
      Caption         =   "条款"
      Height          =   765
      Left            =   10200
      Picture         =   "FmxcXJ.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   8400
      Width           =   795
   End
   Begin VB.Frame frmCGRZ 
      BackColor       =   &H00C0FFC0&
      Caption         =   "采购日志"
      Height          =   2355
      Left            =   9720
      TabIndex        =   85
      Top             =   5760
      Width           =   5925
      Begin VB.CommandButton cmdCSAVE 
         BackColor       =   &H00C0FFC0&
         Caption         =   "提交"
         Height          =   435
         Left            =   5340
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   1890
         Width           =   525
      End
      Begin VB.TextBox txtCED 
         Height          =   375
         Left            =   30
         TabIndex        =   87
         Text            =   "Text1"
         Top             =   1890
         Width           =   5235
      End
      Begin VB.TextBox txtCBZ 
         BackColor       =   &H00C0FFC0&
         Height          =   1545
         Left            =   30
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   86
         Top             =   300
         Width           =   5895
      End
   End
   Begin VB.CommandButton cmdCreate 
      BackColor       =   &H00FFFFC0&
      Caption         =   "采购"
      Height          =   765
      Left            =   10980
      Picture         =   "FmxcXJ.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   8400
      Width           =   675
   End
   Begin VB.TextBox txtXQ 
      BackColor       =   &H00C0FFC0&
      Height          =   3105
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   83
      Text            =   "FmxcXJ.frx":0884
      Top             =   5880
      Width           =   9195
   End
   Begin VB.Frame frmGY 
      BackColor       =   &H00C0FFC0&
      Caption         =   "供应商价格"
      Height          =   1995
      Left            =   5160
      TabIndex        =   68
      Top             =   3720
      Visible         =   0   'False
      Width           =   10005
      Begin VB.TextBox txtGy 
         Height          =   315
         Left            =   6180
         TabIndex        =   81
         Top             =   1560
         Width           =   3735
      End
      Begin VB.OptionButton optGy3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "供应商3"
         Height          =   285
         Left            =   180
         TabIndex        =   80
         Top             =   1230
         Width           =   975
      End
      Begin VB.OptionButton optGy2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "供应商2"
         Height          =   285
         Left            =   180
         TabIndex        =   79
         Top             =   810
         Width           =   975
      End
      Begin VB.OptionButton optGy1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "供应商1"
         Height          =   285
         Left            =   180
         TabIndex        =   78
         Top             =   390
         Width           =   975
      End
      Begin VB.TextBox txtGdj3 
         Height          =   285
         Left            =   5280
         TabIndex        =   77
         Text            =   "Text3"
         Top             =   1230
         Width           =   765
      End
      Begin VB.TextBox txtGdj2 
         Height          =   285
         Left            =   5280
         TabIndex        =   76
         Text            =   "Text2"
         Top             =   787
         Width           =   765
      End
      Begin VB.TextBox txtGdj1 
         Height          =   270
         Left            =   5280
         TabIndex        =   75
         Text            =   "Text1"
         Top             =   390
         Width           =   765
      End
      Begin VB.TextBox txtGY3 
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   72
         Text            =   "Text1"
         Top             =   1260
         Width           =   3195
      End
      Begin VB.TextBox txtGy2 
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "Text1"
         Top             =   825
         Width           =   3195
      End
      Begin VB.TextBox txtGy1 
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   "Text1"
         Top             =   390
         Width           =   3195
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgGy 
         Height          =   1335
         Left            =   6180
         TabIndex        =   82
         Top             =   120
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   2355
         _Version        =   393216
         BackColor       =   12648384
         Rows            =   50
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
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "价格3"
         Height          =   255
         Left            =   4680
         TabIndex        =   74
         Top             =   1260
         Width           =   525
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "价格2"
         Height          =   255
         Left            =   4680
         TabIndex        =   73
         Top             =   825
         Width           =   525
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "价格1"
         Height          =   255
         Left            =   4680
         TabIndex        =   70
         Top             =   420
         Width           =   525
      End
   End
   Begin VB.Frame frmWB 
      BackColor       =   &H00C0FFC0&
      Caption         =   "编辑项"
      Height          =   2385
      Left            =   -360
      TabIndex        =   52
      Top             =   1080
      Width           =   8655
      Begin VB.Frame frmWBJ 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   495
         Left            =   2160
         TabIndex        =   102
         Top             =   1800
         Width           =   3615
         Begin VB.TextBox txtWBJe 
            Height          =   270
            Left            =   2700
            TabIndex        =   106
            Text            =   "Text1"
            Top             =   120
            Width           =   615
         End
         Begin VB.TextBox txtWBdj 
            Height          =   270
            Left            =   960
            TabIndex        =   104
            Text            =   "Text1"
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "对外金额"
            Height          =   255
            Left            =   1900
            TabIndex        =   105
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "对外单价"
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.TextBox txtWcdj 
         Height          =   270
         Left            =   1140
         TabIndex        =   66
         Text            =   "Text1"
         Top             =   1890
         Width           =   915
      End
      Begin VB.CommandButton cmdWadd 
         BackColor       =   &H00FFFF00&
         Caption         =   "添加"
         Height          =   345
         Left            =   7410
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   1860
         Width           =   765
      End
      Begin VB.CommandButton cmdWGx 
         BackColor       =   &H00FF8080&
         Caption         =   "更新"
         Height          =   345
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   1860
         Width           =   765
      End
      Begin VB.CommandButton cmdWdel 
         BackColor       =   &H008080FF&
         Caption         =   "作废"
         Height          =   345
         Left            =   6570
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   1860
         Width           =   765
      End
      Begin VB.TextBox txtWDJ 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   1890
         Width           =   915
      End
      Begin VB.TextBox txtNr 
         BackColor       =   &H00FFFFC0&
         Height          =   1305
         Left            =   1140
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   54
         Top             =   390
         Width           =   7305
      End
      Begin VB.Label lblWcdj 
         BackStyle       =   0  'Transparent
         Caption         =   "成本单价"
         Height          =   345
         Left            =   270
         TabIndex        =   65
         Top             =   1950
         Width           =   825
      End
      Begin VB.Label lblWdj 
         BackStyle       =   0  'Transparent
         Caption         =   "基准价"
         Height          =   255
         Left            =   450
         TabIndex        =   55
         Top             =   1950
         Width           =   585
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "业务内容"
         Height          =   255
         Left            =   210
         TabIndex        =   53
         Top             =   420
         Width           =   795
      End
   End
   Begin VB.TextBox txtBrq 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   10290
      Locked          =   -1  'True
      TabIndex        =   49
      Top             =   420
      Width           =   1485
   End
   Begin VB.TextBox txtYfadr 
      BackColor       =   &H00FFFFC0&
      Height          =   270
      Left            =   4710
      Locked          =   -1  'True
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   420
      Width           =   4335
   End
   Begin VB.TextBox txtBz 
      BackColor       =   &H00FFFFC0&
      Height          =   1035
      Left            =   9810
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   30
      Text            =   "FmxcXJ.frx":088A
      Top             =   6000
      Width           =   5325
   End
   Begin VB.TextBox txtXmmc 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   4710
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   60
      Width           =   4335
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFC0&
      Caption         =   "返回"
      Height          =   765
      Left            =   14550
      Picture         =   "FmxcXJ.frx":0890
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "返回"
      Top             =   8400
      Width           =   675
   End
   Begin VB.CommandButton cmdMod 
      BackColor       =   &H00C0FFC0&
      Caption         =   "修改"
      Height          =   765
      Left            =   12420
      Picture         =   "FmxcXJ.frx":0992
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "修改"
      Top             =   8400
      Width           =   675
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0FFC0&
      Caption         =   "保存"
      Height          =   765
      Left            =   13140
      Picture         =   "FmxcXJ.frx":0C9C
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "保存"
      Top             =   8400
      Width           =   675
   End
   Begin VB.CommandButton cmdD 
      BackColor       =   &H00C0FFC0&
      Caption         =   "删除"
      Enabled         =   0   'False
      Height          =   765
      Left            =   13830
      Picture         =   "FmxcXJ.frx":1306
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8400
      Width           =   675
   End
   Begin VB.CommandButton cmdNQ 
      BackColor       =   &H008080FF&
      Caption         =   "审核"
      Height          =   765
      Left            =   11700
      Picture         =   "FmxcXJ.frx":1490
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8400
      Width           =   675
   End
   Begin VB.Frame frmSd 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1275
      Left            =   30
      TabIndex        =   19
      Top             =   4680
      Width           =   5145
      Begin VB.TextBox txtLx 
         Height          =   270
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   108
         Top             =   900
         Width           =   3375
      End
      Begin VB.Frame frmBJ 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   495
         Left            =   120
         TabIndex        =   95
         Top             =   0
         Width           =   5295
         Begin VB.TextBox txtYH 
            Height          =   270
            Left            =   4080
            TabIndex        =   111
            Text            =   "Text1"
            Top             =   120
            Width           =   855
         End
         Begin VB.TextBox txtBje 
            Height          =   270
            Left            =   2520
            TabIndex        =   99
            Text            =   "Text2"
            Top             =   120
            Width           =   855
         End
         Begin VB.TextBox txtBdj 
            Height          =   270
            Left            =   840
            TabIndex        =   98
            Text            =   "Text1"
            Top             =   120
            Width           =   765
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "优惠价"
            Height          =   255
            Left            =   3480
            TabIndex        =   110
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "对外金额"
            Height          =   255
            Left            =   1680
            TabIndex        =   97
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "对外单价"
            Height          =   255
            Left            =   0
            TabIndex        =   96
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.TextBox txtSL 
         Height          =   270
         Left            =   990
         TabIndex        =   60
         Top             =   510
         Width           =   1125
      End
      Begin VB.CommandButton cmdNDel 
         BackColor       =   &H008080FF&
         Caption         =   "作废"
         Height          =   345
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   450
         Width           =   855
      End
      Begin VB.CommandButton cmdNGx 
         BackColor       =   &H00FF8080&
         Caption         =   "更新"
         Height          =   345
         Left            =   2220
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   450
         Width           =   855
      End
      Begin VB.CommandButton cmdDao 
         BackColor       =   &H00FFFF00&
         Caption         =   "业务添加"
         Height          =   345
         Left            =   4050
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   450
         Width           =   915
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "业务类型"
         Height          =   255
         Left            =   120
         TabIndex        =   107
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "数量"
         Height          =   225
         Left            =   270
         TabIndex        =   61
         Top             =   540
         Width           =   375
      End
   End
   Begin VB.TextBox txtJHg 
      BackColor       =   &H00FFFFC0&
      Height          =   270
      Left            =   10140
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   8010
      Width           =   1425
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9450
      Top             =   8670
   End
   Begin VB.Timer timQuit 
      Interval        =   1000
      Left            =   9960
      Top             =   8730
   End
   Begin VB.Frame frmCg 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "采购部填写"
      Height          =   1395
      Left            =   0
      TabIndex        =   6
      Top             =   3300
      Width           =   5175
      Begin VB.CommandButton cmdGy 
         BackColor       =   &H00C0E0FF&
         Caption         =   "供应商"
         Height          =   315
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   1020
         Width           =   885
      End
      Begin VB.Frame frmJ 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   285
         Left            =   2160
         TabIndex        =   11
         Top             =   210
         Width           =   2235
         Begin VB.TextBox txtJdj 
            Height          =   270
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   30
            Width           =   1155
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "基准单价"
            Height          =   255
            Left            =   210
            TabIndex        =   13
            Top             =   60
            Width           =   855
         End
      End
      Begin VB.TextBox txtDj 
         Height          =   270
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   240
         Width           =   1155
      End
      Begin VB.Frame frmZ 
         Height          =   405
         Left            =   -8310
         TabIndex        =   10
         Top             =   690
         Width           =   8295
      End
      Begin VB.TextBox txtDrq 
         Height          =   270
         Left            =   990
         TabIndex        =   9
         Top             =   1020
         Width           =   1125
      End
      Begin VB.TextBox txtMj 
         Height          =   270
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtZBQ 
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   990
         TabIndex        =   7
         Top             =   600
         Width           =   3165
      End
      Begin VB.Label lblLid 
         Caption         =   "lblLid"
         Height          =   255
         Left            =   2670
         TabIndex        =   62
         Top             =   990
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "到货期"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "市场指导价"
         Height          =   315
         Left            =   3480
         TabIndex        =   16
         Top             =   840
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lblDj 
         BackStyle       =   0  'Transparent
         Caption         =   "成本单价"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   270
         Width           =   765
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "质保期"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   630
         Width           =   615
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   495
      Left            =   13020
      TabIndex        =   5
      Top             =   7140
      Visible         =   0   'False
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   873
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   5055
      Left            =   0
      TabIndex        =   31
      Top             =   780
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   8916
      _Version        =   393216
      BackColor       =   16777152
      BackColorFixed  =   15728356
      BackColorBkg    =   16777152
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   3
      PictureType     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgP 
      Height          =   3135
      Left            =   30
      TabIndex        =   32
      Top             =   6000
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   5530
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
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin MSComCtl2.DTPicker dtpBrq 
      Height          =   315
      Left            =   10350
      TabIndex        =   50
      Top             =   420
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   8454016
      CalendarTitleBackColor=   16711808
      CalendarTrailingForeColor=   -2147483635
      Format          =   100597761
      CurrentDate     =   38797
   End
   Begin VB.CommandButton cmdHT 
      BackColor       =   &H00C0FFC0&
      Caption         =   "合同评审单"
      Height          =   345
      Left            =   12210
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   60
      Width           =   2745
   End
   Begin NiceFormControl.NiceButton cmdDht 
      Height          =   345
      Left            =   12210
      TabIndex        =   64
      Top             =   30
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   609
      BTYPE           =   3
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   255
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FmxcXJ.frx":18D2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      Style           =   9
      Caption         =   "导入合同"
   End
   Begin VB.Label lblBj 
      BackStyle       =   0  'Transparent
      Caption         =   "对外报价"
      Height          =   255
      Left            =   9300
      TabIndex        =   112
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label lblBhg 
      BackStyle       =   0  'Transparent
      Caption         =   "实际报价"
      Height          =   255
      Left            =   9300
      TabIndex        =   100
      Top             =   7680
      Width           =   735
   End
   Begin VB.OLE OLE2 
      Class           =   "Excel.Sheet.8"
      Height          =   585
      Left            =   150
      OleObjectBlob   =   "FmxcXJ.frx":18EE
      TabIndex        =   91
      Top             =   780
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      DataField       =   "FName"
      DataSource      =   "adoFile"
      Height          =   255
      Left            =   9420
      TabIndex        =   90
      Top             =   8910
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "报价有效期"
      Height          =   315
      Left            =   9330
      TabIndex        =   51
      Top             =   450
      Width           =   1065
   End
   Begin VB.Label lblBid 
      Caption         =   "lblBid"
      Height          =   285
      Left            =   9720
      TabIndex        =   48
      Top             =   8280
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "地址"
      Height          =   285
      Left            =   4020
      TabIndex        =   46
      Top             =   480
      Width           =   555
   End
   Begin VB.Label lblWhg 
      Caption         =   "Whg"
      Height          =   225
      Left            =   11580
      TabIndex        =   45
      Top             =   5370
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblYwy 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1080
      TabIndex        =   44
      Top             =   420
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "业务员"
      Height          =   255
      Left            =   360
      TabIndex        =   43
      Top             =   480
      Width           =   645
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "备注"
      Height          =   225
      Left            =   9300
      TabIndex        =   42
      Top             =   6120
      Width           =   435
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "项目名称"
      Height          =   285
      Left            =   3660
      TabIndex        =   41
      Top             =   120
      Width           =   795
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "编号"
      Height          =   285
      Left            =   390
      TabIndex        =   40
      Top             =   90
      Width           =   435
   End
   Begin VB.Label lblBh 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      Height          =   285
      Left            =   1080
      TabIndex        =   39
      Top             =   60
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "性质"
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   12270
      TabIndex        =   38
      Top             =   510
      Width           =   585
   End
   Begin VB.Label lblZl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   12750
      TabIndex        =   37
      Top             =   510
      Width           =   2115
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "基准合计"
      Height          =   255
      Left            =   9300
      TabIndex        =   36
      Top             =   8070
      Width           =   765
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "询价日期"
      Height          =   195
      Left            =   9450
      TabIndex        =   35
      Top             =   120
      Width           =   885
   End
   Begin VB.Label lblRq 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
      Height          =   285
      Left            =   10290
      TabIndex        =   34
      Top             =   60
      Width           =   1485
   End
   Begin VB.Label lblTX 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   11760
      TabIndex        =   33
      Top             =   8010
      Width           =   3945
   End
End
Attribute VB_Name = "FmxcXJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lc As Integer
Dim Fwid As Long
Dim LCRen As String
Dim LCUid As String
Dim timZm As Integer
Dim htRow As Single
Dim GyId As Integer

Dim THid As Long

Dim Bh As String

Private Sub cmdBack_Click()
Me.Visible = False
If Dialog.Visible = True And FmxcNew.Visible = False Then
    Call mod1.refEnvent(1)
    Dialog.ZOrder 0
    Dialog.Enabled = True
 
ElseIf frmGxBiao.Visible = True Then
    If frmGxBNew.Visible = True Then
        frmGxBNew.Show
        frmGxBNew.ZOrder 0
        Exit Sub
    End If
    frmGxBiao.Enabled = True
    frmGxBiao.ZOrder 0
ElseIf FmxcNew.Visible = True Then
    FmxcNew.Show
    FmxcNew.ZOrder 0

End If
End Sub

Private Sub cmdCreate_Click()
If frmCGRZ.Visible = False Then
    frmCGRZ.Visible = True
    txtCED.Visible = False
    cmdCSAVE.Visible = False
    If mod1.Bm = "市场营销部" Then
        txtCED.Visible = True
        txtCED.Locked = False
        cmdCSAVE.Visible = True
    End If
Else
    frmCGRZ.Visible = False
End If
End Sub

Private Sub cmdCSAVE_Click()
Dim tt As String
On Error Resume Next

If LCRen <> mod1.DName And mod1.DName <> "马晓聪" Then Exit Sub


timZm = 8 '采购提交
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单2011"
    mod1.cmd.Parameters("@NBLX") = "采购提交"
    mod1.cmd.Parameters("@bh") = lblBid.ToolTipText
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtCED.Text
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
cmdSave.Enabled = False
End Sub

Private Sub cmdD_Click()
Dim ii As Integer
Dim tt As String
On Error Resume Next
tt = "select lc from htping where hid=" & Val(cmdHT.ToolTipText)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
If mod1.HTP.BOF = False Then
    If mod1.HTP.Fields(0).Value > 1 And mod1.Mname <> "马晓聪" Then
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        Exit Sub
    End If
End If
Set mod1.HTP = Nothing
If lblYwy.Caption <> mod1.DName And mod1.DName <> "马晓聪" Then Exit Sub

ii = MsgBox("是否删除此询价单？", vbYesNo + vbQuestion, "Hello")
If ii = vbNo Then
    Exit Sub
End If
timZm = 3 '删除合同
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单2011"
    mod1.cmd.Parameters("@NBLX") = "删除"
    mod1.cmd.Parameters("@bh") = lblBid.ToolTipText
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(lblZl.Caption)
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = 0
    mod1.cmd.Parameters("@mm2") = Val(cmdHT.ToolTipText)
    mod1.cmd.Parameters("@mm3") = htRow
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
If frmGxBiao.Visible = True Then
    frmGxBiao.adoXj.Requery
    Set frmGxBiao.dtgXj.DataSource = frmGxBiao.adoXj
End If
End Sub

Private Sub cmdDao_Click()
'Set frmLingjian.LpXh = CreateObject("adodb.recordset")
Dim tt As String
Dim oo As Integer: Dim ii As Integer
Dim Ra, La
On Error Resume Next
If txtLx.Text = "" And lblZl.Caption = "询价指令" Then
    MsgBox "请确定业务类型！"
    Exit Sub
End If
    FmxcXjHp.cmdDao.Caption = "导入"
'''''If mod1.Mname = "马晓聪" Then
If Val(Right(lblBh.Caption, 5)) > 14504 Then
'''    tt = "select top 50 bh,partname,'原厂编号:'+oname+' '+gg+' '+xn+' '+ff+' 适用机组:'+jz from nlpmxc where lc>1 and jyf=1 and delf=1 order by bh desc"
'''    Call FmxcXjHp.Bound(tt)
'''    FmxcXjHp.Show
'''    FmxcXjHp.ZOrder 0
    Call frmHPBR.dtgLPFF
    Call frmHPBR.dtgFF
    frmHPBR.frmZX.Visible = True
    frmHPBR.Show
    frmHPBR.ZOrder 0
'''    If mod1.Qy = "上海" Then
'''        frmHPBR.frmZX.Visible = False
'''    Else
        frmHPBR.frmZX.Visible = True
'''    End If
    Exit Sub
End If
tt = "SELECT top 100 dbo.l_goods.code, dbo.l_goods.name, dbo.l_goods.specs, dbo.l_goodstype.name AS goodtypename, dbo.l_goodsunit.unitname,dbo.l_goods.goodsid" & _
    " FROM dbo.l_goods LEFT OUTER JOIN dbo.l_goodsunit ON dbo.l_goods.goodsid = dbo.l_goodsunit.goodsid LEFT OUTER JOIN dbo.l_goodstype ON dbo.l_goods.gdtypeid = dbo.l_goodstype.gdtypeid where  dbo.l_goods.closed=0"
Call frmGxbjSD.dtgFF
Call frmGxbjSD.CX(tt)

frmGxbjSD.Show
frmGxbjSD.ZOrder 0

End Sub

Private Sub cmdDht_Click()
Dim Hid As Long
Hid = Val(cmdHT.ToolTipText)
If Hid = 0 Then Exit Sub

Dim ii As Integer
Dim tt As String
On Error Resume Next
tt = "select htbh from htping where hid=" & Val(cmdHT.ToolTipText)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
If mod1.HTP.BOF = False Then
If mod1.HTP.Fields(0).Value <> "HMNEW" And mod1.DName <> "马晓聪" And mod1.Mname <> "马晓聪" Then
    Exit Sub
End If
End If
If lblYwy.Caption <> mod1.DName And mod1.DName <> "马晓聪" Then Exit Sub


timZm = 7 '导入合同
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单2011"
    mod1.cmd.Parameters("@NBLX") = "导入合同"
    mod1.cmd.Parameters("@bh") = Hid
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = lblYwy.Caption
    mod1.cmd.Parameters("@mt2") = lblYwy.ToolTipText
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(lblBid.ToolTipText)
    mod1.cmd.Parameters("@mm2") = htRow
    mod1.cmd.Parameters("@mm3") = Lc
    mod1.cmd.Parameters("@mm4") = Val(txtJHg.Text)
    If FmxcNew.NewId = 0 Then Exit Sub
    mod1.cmd.Parameters("@mm11") = FmxcNew.NewId '根据实际合同中的性质,调整询价单的类型
    FmxcNew.dtgLx.Col = 2: FmxcNew.dtgLx.Row = FmxcNew.NewId
    mod1.cmd.Parameters("@mt3") = FmxcNew.dtgLx.Text
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = txtBrq.Text
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
cmdSave.Enabled = False
End Sub

Private Sub cmdDing_Click()
Dim tt As String
Dim ii As Integer
Dim Ra
On Error Resume Next
cmdDing.Enabled = False
'''''''If lblLc.Caption = 1 Then
'''''''    dtgN.Row = 1
'''''''    dtgN.Col = 1
'''''''    If dtgN.Text = "" Then
'''''''        ii = MsgBox("您没有在业务明细表中添加项目内容,是否现在添加?", vbQuestion + vbYesNo + vbDefaultButton1, "请您注意!")
'''''''        If ii = vbYes Then
'''''''            Call cmdDao_Click
'''''''        End If
'''''''        Exit Sub
'''''''    End If
'''''''End If
If optT2.Value = True And Val(cmdHT.ToolTipText) > 0 Then
    tt = "select htf from htping where hid=" & Val(cmdHT.ToolTipText)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    If Ra(0, 0) <> 0 Then
        MsgBox "合同不在执行状态，不能够驳回！"
        Exit Sub
    End If
End If

If Lc = 1 And Val(lblBid.ToolTipText) < 17662 Then   '2012的单子，不能走流程
'    MsgBox "此为旧版的合同，请与马晓聪联系"
'    Exit Sub
End If
If Lc = 2 And OptT1.Value = True And Val(txtJHg.Text) = 0 Then
    MsgBox ("合计为空,不能够签字,请重新更新基准价!")
    Exit Sub
End If
If optT2.Value = True And txtQM.Text = "" Then
    MsgBox ("请您一定要告诉拒绝我的理由!  :) ")
    Exit Sub
End If
If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If
If Lc = 2 And LCUid = mod1.DHid And txtBrq.Text = "" And Me.OptT1.Value = True Then
    MsgBox "请确认报价有效期"
    dtpBrq.Visible = True
    cmdSave.Enabled = True
    Exit Sub
End If
timZm = 6 '签字
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单2013"
    mod1.cmd.Parameters("@NBLX") = "签字"
    mod1.cmd.Parameters("@bh") = Val(lblBid.ToolTipText)
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(lblYwy.Caption)
    mod1.cmd.Parameters("@mt2") = Trim(lblYwy.ToolTipText)
    mod1.cmd.Parameters("@mt3") = Trim(txtXmmc.Text)
    mod1.cmd.Parameters("@mt4") = Trim(cmdHT.ToolTipText)
    mod1.cmd.Parameters("@mt5") = Trim(lblZl.Caption)
    mod1.cmd.Parameters("@mt7") = LCRen

    mod1.cmd.Parameters("@mlt1") = txtQM.Text '评审建议
    If mod1.Qy <> "上海" And lblZl.Caption = "询价指令" Then Lc = 100
    mod1.cmd.Parameters("@mm1") = Lc
    mod1.cmd.Parameters("@mm2") = Fwid
    mod1.cmd.Parameters("@mm3") = htRow
    mod1.cmd.Parameters("@mm5") = Val(txtJHg.Text)
    mod1.cmd.Parameters("@mm6") = Val(cmdHT.ToolTipText)
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

Private Sub cmdGy_Click()
Dim tt As String
Dim Bh As String
Dim Mc1: Dim Mc2: Dim Mc3: Dim Ra
If frmGY.Visible = True Then
    frmGY.Visible = False
Else
    frmGY.Visible = True
    dtgN.Col = 0: Bh = dtgN.Text
    '调出此货品的原来供应商及成本价和市场指导价
    If Bh <> "" Then
'''        If txtGy1.ToolTipText = 0 And txtGy2.ToolTipText = 0 And txtGy3.ToolTipText = 0 Then
            tt = "declare @gid1 int, @gid2 int,@gid3 int;" & _
                "select @gid1=gid1,@gid2=gid2,@gid3=gid3 from nlpmxc where bh='" & Bh & "';" & _
                "select gid1,gid2,gid3,dj1,dj2,dj3,listprice from nlpmxc where bh='" & Bh & "';" & _
                "select mc from gymxc where gid=@gid1;" & _
                "select mc from gymxc where gid=@gid2;" & _
                "select mc from gymxc where gid=@gid3;"
                Set mod1.HTP = CreateObject("adodb.recordset")
                mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
                On Error Resume Next
                Ra = mod1.HTP.GetRows
                Set mod1.HTP = mod1.HTP.NextRecordset
                Mc1 = mod1.HTP.GetRows
                Set mod1.HTP = mod1.HTP.NextRecordset
                Mc2 = mod1.HTP.GetRows
                Set mod1.HTP = mod1.HTP.NextRecordset
                Mc3 = mod1.HTP.GetRows
                mod1.HTP.Close
                Set mod1.HTP = Nothing
                txtGy1.ToolTipText = Ra(0, 0)
                txtGy2.ToolTipText = Ra(1, 0)
                txtGY3.ToolTipText = Ra(2, 0)
                txtGdj1.Text = Ra(3, 0)
                txtGdj2.Text = Ra(4, 0)
                txtGdj3.Text = Ra(5, 0)
                txtMj.Text = Ra(6, 0)
                txtGy1.Text = Mc1(0, 0)
                txtGy2.Text = Mc2(0, 0)
                txtGY3.Text = Mc3(0, 0)
                optGy1.Value = False: optGy2.Value = False: optGy3.Value = False
                optGy1.ForeColor = &H80000008: txtGy1.ForeColor = &H80000008: txtGdj1.ForeColor = &H80000008
                optGy2.ForeColor = &H80000008: txtGy2.ForeColor = &H80000008: txtGdj2.ForeColor = &H80000008
                optGy3.ForeColor = &H80000008: txtGY3.ForeColor = &H80000008: txtGdj3.ForeColor = &H80000008
'''        End If
    End If
End If
End Sub

Private Sub cmdHt_Click()
Dim Bh As String
Dim tt As String
Dim ii As Integer
Dim Hid As Integer
Dim Ra
If Val(cmdHT.ToolTipText) = 0 Then
    If FmxcXJ.lblYwy <> mod1.DName Then
        Exit Sub
    End If
    If txtBrq.Text = "" Then
        MsgBox ("没有报价有效期，不能关联合同！")
        Exit Sub
    End If
    If txtBrq.Text < mod1.DQda Then
        If Val(lblBid.ToolTipText) < 17662 Then
            MsgBox "这是去年的询价单，不能关联合同!"
            Exit Sub
        End If
        ii = MsgBox("已经超过报价有效期！,如果关联合同，此询价单将重新走流程", vbQuestion + vbYesNo, "请确认")
        If ii = vbNo Then
            Exit Sub
        End If
    End If
    
    Bh = InputBox("请输入关联的合同编号或合同序列号")
    If Val(Bh) = 0 Then
        tt = "select lc,xuid,hid,uid from htping where delf=1 and htbh='" & Bh & "'"
    Else
        tt = "select lc,xuid,hid,uid from htping where delf=1 and hid=" & Bh
    End If
    If Bh = "" Then Exit Sub
    On Error Resume Next
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    If IsNull(Ra(0, 0)) = True Then
        MsgBox "您输入了不正确的合同号！"
        Exit Sub
    End If
    If Ra(1, 0) <> mod1.DHid And Ra(3, 0) <> mod1.DHid Then
        MsgBox "此合同不是你管理！"
        Exit Sub
    End If
    If Ra(0, 0) > 1 Then
        MsgBox "此合同不是在编辑状态，不能导入此询价单！"
        Exit Sub
    End If
    Hid = Ra(2, 0)
    THid = Ra(2, 0)
    timZm = 9 '导入合同2013
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单2011"
    mod1.cmd.Parameters("@NBLX") = "导入合同"
    mod1.cmd.Parameters("@bh") = Hid
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = lblYwy.Caption
    mod1.cmd.Parameters("@mt2") = lblYwy.ToolTipText
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(lblBid.ToolTipText)
    mod1.cmd.Parameters("@mm2") = htRow
    mod1.cmd.Parameters("@mm3") = Lc
    mod1.cmd.Parameters("@mm4") = Val(txtJHg.Text)
   ' If FmxcNew.NewId = 0 Then Exit Sub
    mod1.cmd.Parameters("@mm11") = FmxcNew.NewId '根据实际合同中的性质,调整询价单的类型
    FmxcNew.dtgLx.Col = 2: FmxcNew.dtgLx.Row = FmxcNew.NewId
    mod1.cmd.Parameters("@mt3") = FmxcNew.dtgLx.Text
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = txtBrq.Text
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
cmdSave.Enabled = False
    
    
Else
    Call FmxcNew.Bound(Val(cmdHT.ToolTipText))
    FmxcNew.Show
    FmxcNew.ZOrder
    Me.Visible = False
    If mod1.DName = "朱婷婷" Or mod1.DName = "汪燕明" Or mod1.DName = "吴金荣" Or mod1.DName = "吴金荣" Then
    Call FmxcNew.Xian
End If
End If
End Sub

Private Sub cmdMod_Click()
Dim tt As String
Dim Ra
frmCg.Visible = False
If LCRen = "吴金荣" And mod1.DName = "吴金荣" Then
    LCRen = "吴金荣": LCUid = "HM804"
End If
If lblZl.Caption = "询价指令" Then
    frmCg.Visible = True
    frmBJ.Visible = True
    If Lc = 2 Then
        frmSd.Visible = False
    Else
        frmSd.Visible = True
    End If
    frmSd.Visible = True
    cmdSave.Enabled = True
    cmdD.Enabled = True
    Exit Sub
End If
If Lc = 1 And LCRen = mod1.DName Then
    cmdSave.Enabled = True
    cmdD.Enabled = True
    'If lblZl.Caption = "询价指令" And mod1.Qy = "上海" Then Exit Sub
    frmCg.Visible = True: txtMj.Locked = False
    
    If htRow = 1 Or htRow = 2 Or htRow = 3 Or htRow = 4 Or htRow = 5 Or htRow = 6 Or htRow = 7 Or htRow = 7.14 And lblZl.Caption <> "三菱" Or htRow = 8 And lblZl.Caption <> "松下" Or htRow >= 20 Or InStr(1, "维保大修其他人工压缩机维修保养中介业务分包运费吊装费工程人工", lblZl.Caption) > 0 Or _
     (Val(lblBid.ToolTipText) >= 20512 And lblZl.ToolTipText = True) Then
        frmWB.Visible = True
        txtWDJ.Locked = True
        If mod1.GxName = "报价功能" And mod1.GXF = True And lblZl.Caption = "询价指令" Then
            frmWBJ.Visible = True
            txtWBdj.Text = ""
            txtWBJe.Text = ""
            cmdWadd.Enabled = True
        Else
            frmWBJ.Visible = False
        End If
    Else
        frmSd.Visible = True
        cmdDao.Enabled = True
        cmdNGx.Enabled = True
        If mod1.GxName = "报价功能" And mod1.GXF = True And lblZl.Caption = "询价指令" Then
            frmBJ.Visible = True
            txtBdj.Text = ""
            txtBje.Text = ""
            If mod1.Qy <> "上海" Then
                cmdGy.Visible = False
            Else
                cmdGy.Visible = True
            End If
            cmdGy.Visible = True
        Else
            frmBJ.Visible = False
        End If
    End If
ElseIf Lc < 5 And Lc > 1 And LCRen = mod1.DName Or mod1.DName = "马晓聪" Then
    If htRow = 1 Or htRow = 2 Or htRow = 3 Or htRow = 4 Or htRow = 5 Or htRow = 6 Or htRow = 7 Or htRow = 7.14 And lblZl.Caption <> "三菱" Or htRow = 8 And lblZl.Caption <> "松下" Or htRow >= 20 Or InStr(1, "维保大修其他人工压缩机维修保养中介业务分包运费吊装费工程人工", lblZl.Caption) > 0 Or _
     (Val(lblBid.ToolTipText) >= 20512 And lblZl.ToolTipText = True) Then
        dtpBrq.Visible = True
        frmWB.Visible = True
        txtWDJ.Locked = False
        lblWdj.Visible = False
        txtWDJ.Visible = False
        lblWcdj.Visible = True
        txtWcdj.Visible = True
    Else
        dtpBrq.Visible = True
        frmSd.Visible = True
        cmdDao.Enabled = True
        cmdNGx.Enabled = True
        frmCg.Visible = True
        frmJ.Visible = False
        lblDj.Visible = True
        txtDj.Visible = True
    End If
        cmdSave.Enabled = True
ElseIf Lc = 5 Or Lc = 100 And lblYwy.ToolTipText = mod1.DHid Then
    tt = "select htbh from htping where hid=" & Val(cmdHT.ToolTipText)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    If mod1.HTP.BOF = False Then
        If mod1.HTP.Fields(0).Value <> "HMNEW" And mod1.DName <> "马晓聪" Then
            Exit Sub
        End If
    End If
    If lblZl.ToolTipText = "False" Then
        frmSd.Visible = True
        cmdDao.Enabled = False
        cmdNGx.Enabled = True
    Else
        frmWB.Visible = True
        cmdWadd.Enabled = False
        cmdWdel.Enabled = True
        cmdWGx.Enabled = True
    End If

End If
End Sub


Private Sub cmdNDel_Click()
Dim ii As Integer
On Error Resume Next
If Val(lblLid.Caption) = 0 Then Exit Sub
ii = MsgBox("是否作废此条记录?", vbInformation + vbYesNo, "询问")
If ii = vbYes Then
   
     timZm = 1
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单2011"
    mod1.cmd.Parameters("@NBLX") = "配件删除"
    mod1.cmd.Parameters("@bh") = cmdHT.ToolTipText
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = lblBid.ToolTipText
    mod1.cmd.Parameters("@mt2") = lblZl.Caption
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(lblLid.Caption)
    mod1.cmd.Parameters("@mm2") = Val(cmdHT.ToolTipText)
    mod1.cmd.Parameters("@mm3") = htRow
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = Null
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

Private Sub cmdNGx_Click()
Dim ii As Integer
Dim Bh As String
On Error Resume Next
If Val(lblLid.Caption) = 0 Then Exit Sub
If txtLx.Text = "" And mod1.Qy <> "上海" Then
    MsgBox "请确定业务类型！"
    Exit Sub
End If
If txtDrq.Text = "" And Lc > 1 Then
    MsgBox "请填入到货期！"
    Exit Sub
End If
If optGy1.Value = False And optGy2.Value = False And optGy3.Value = False And Lc > 1 Then
    MsgBox "请确定向哪家供应商购买！"
    Exit Sub
End If
dtgN.Col = 0: Bh = dtgN.Text
'If Bh = "" Then Exit Sub
     timZm = 2
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单2011"
    mod1.cmd.Parameters("@NBLX") = "项目更新"
    mod1.cmd.Parameters("@bh") = cmdHT.ToolTipText
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = lblBid.ToolTipText
    mod1.cmd.Parameters("@mt2") = lblZl.Caption
    If htRow = 1 Or htRow = 2 Or htRow = 3 Or htRow = 4 Or htRow = 5 Or htRow = 6 Or htRow = 7 And lblZl.Caption <> "三菱" Or htRow = 8 And lblZl.Caption <> "松下" Or htRow >= 20 Or InStr(1, "维保大修其他人工压缩机维修保养中介业务分包运费吊装费工程人工", lblZl.Caption) > 0 Or _
     (Val(lblBid.ToolTipText) >= 20512 And lblZl.ToolTipText = True) Then
        mod1.cmd.Parameters("@mt3") = txtNr.Text
    Else
        mod1.cmd.Parameters("@mt3") = txtZBQ.Text
    End If
    mod1.cmd.Parameters("@mt4") = txtDrq.Text
    mod1.cmd.Parameters("@mt5") = Trim(cmdHT.ToolTipText)
    mod1.cmd.Parameters("@mt21") = Bh '货品编号，更新相应货品表中的供应商，单价和市场指导价
    mod1.cmd.Parameters("@mt22") = txtLx.Text '业务类型
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(lblLid.Caption)
    mod1.cmd.Parameters("@mm2") = Val(cmdHT.ToolTipText)
    mod1.cmd.Parameters("@mm3") = htRow
    mod1.cmd.Parameters("@mm5") = Val(txtMj.Text)
    If htRow = 1 Or htRow = 2 Or htRow = 3 Or htRow = 4 Or htRow = 5 Or htRow = 6 Or htRow = 7 And lblZl.Caption <> "三菱" Or htRow = 8 And lblZl.Caption <> "松下" Or htRow >= 20 Or InStr(1, "维保大修其他人工压缩机维修保养中介业务分包运费吊装费工程人工", lblZl.Caption) > 0 Or _
     (Val(lblBid.ToolTipText) >= 20512 And lblZl.ToolTipText = True) Then
        mod1.cmd.Parameters("@mm6") = Val(txtWcdj.Text)
        mod1.cmd.Parameters("@mm7") = Val(txtWDJ.Text)
        mod1.cmd.Parameters("@mm8") = 1
    Else
        mod1.cmd.Parameters("@mm6") = Val(txtDj.Text)
        mod1.cmd.Parameters("@mm7") = Val(txtJdj.Text)
        mod1.cmd.Parameters("@mm8") = Val(txtSL.Text)
        mod1.cmd.Parameters("@mm11") = Val(txtGy1.ToolTipText)
        mod1.cmd.Parameters("@mm12") = Val(txtGy2.ToolTipText)
        mod1.cmd.Parameters("@mm13") = Val(txtGY3.ToolTipText)
        mod1.cmd.Parameters("@mm14") = Val(txtGdj1.Text)
        mod1.cmd.Parameters("@mm15") = Val(txtGdj2.Text)
        mod1.cmd.Parameters("@mm16") = Val(txtGdj3.Text)
        
        If optGy1.Value = True Then
            mod1.cmd.Parameters("@mm17") = txtGy1.ToolTipText
        ElseIf optGy2.Value = True Then
            mod1.cmd.Parameters("@mm17") = txtGy2.ToolTipText
        ElseIf optGy3.Value = True Then
            mod1.cmd.Parameters("@mm17") = txtGY3.ToolTipText
        End If
        mod1.cmd.Parameters("@mm18") = Val(txtBdj.Text)
        mod1.cmd.Parameters("@mm19") = Val(txtBje.Text)
        mod1.cmd.Parameters("@mm20") = Val(txtYH.Text)
    End If
    mod1.cmd.Parameters("@mm9") = Lc
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = Null
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

End Sub


Private Sub cmdNQ_Click()
Dim ii As Integer
Dim tt As String
Dim Ra
Dim Tywy As String '单子流转到下一人的姓名
Dim Tuid As String
Dim Oywy As String '原来流转人的名字
Dim Ouid As String '原来流转人的工号

Dim oo As Integer
On Error Resume Next


If lblTX.Caption = "审核完毕!" Then Exit Sub
If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If


If LCUid <> mod1.DHid Then
    tt = "select xuid from htping where hid=" & Val(cmdHT.ToolTipText)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing

    If Ra(0, 0) <> mod1.DHid Or Lc > 1 Then
        MsgBox "此处应由" & LCRen & "签字! 请您不要再点"
        Exit Sub
    End If
End If

frmQm.Visible = True
cmdDing.Enabled = True
If Lc = 1 Then
    optT2.Enabled = False
    OptT1.Value = True
    
Else
    OptT1.Enabled = True
    optT2.Enabled = True
    OptT1.Value = False
    optT2.Value = False
End If
If Lc = 2 Then
    optT2.Caption = "驳回"
Else
    optT2.Caption = "增补"
End If
End Sub

Private Sub cmdNQ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lc = 100 And Button = 2 And lblYwy.Caption = mod1.DName Then
    frmQm.Visible = True
    OptT1.Enabled = False
    optT2.Value = True
    cmdDing.Enabled = True
End If
End Sub


Private Sub cmdSave_Click()
Dim ii As Integer
Dim tt As String
On Error Resume Next
''''''tt = "select htbh from htping where hid=" & Val(cmdHT.ToolTipText)
''''''Set mod1.HTP = CreateObject("adodb.recordset")
''''''mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
''''''If mod1.HTP.BOF = False Then
''''''If mod1.HTP.Fields(0).Value <> "HMNEW" And mod1.DName <> "马晓聪" Then
''''''    Exit Sub
''''''End If
''''''End If
If LCRen <> mod1.DName And mod1.DName <> "马晓聪" Then Exit Sub

frmWB.Visible = False
frmCg.Visible = False
timZm = 4 '保存合同
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单2011"
    mod1.cmd.Parameters("@NBLX") = "保存"
    mod1.cmd.Parameters("@bh") = lblBid.ToolTipText
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtYfadr.Text
    mod1.cmd.Parameters("@mt2") = txtXmmc.Text
    mod1.cmd.Parameters("@mlt1") = txtBz.Text
    mod1.cmd.Parameters("@mm1") = Val(txtXmmc.ToolTipText)
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = txtBrq.Text
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
cmdSave.Enabled = False
End Sub

Private Sub cmdTK_Click()
Dim tt As String
Dim bt() As Byte
Dim cc As String
cc = "Provider=SQLOLEDB.1;Password=hugemanzou;Persist Security Info=True;User ID=zou;Initial Catalog=HMZou;Data Source=10.128.123.10"
On Error Resume Next
tt = "select Nfile,fsize from hmzou.dbo.hmfile "
adoFile.Recordset.Close
adoFile.Recordset.Open tt, cc, adOpenKeyset, adLockReadOnly, adCmdText
ReDim bt(adoFile.Recordset.Fields("Fsize").Value) As Byte
bt() = adoFile.Recordset.Fields("Nfile").GetChunk(adoFile.Recordset.Fields("Fsize").Value + 1)


Open ("c:\work\" & "技术模板.xls") For Binary As #3
Put #3, , bt()
Close #3

'Me.Visible = False

    OLE2.SourceDoc = "c:\work\技术模板.xls"
    OLE2.Action = 1
    OLE2.DoVerb (-2)

 
End Sub

Private Sub cmdWadd_Click()
Dim ii As Integer
On Error Resume Next

   
     timZm = 5
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单2011"
    mod1.cmd.Parameters("@NBLX") = "人工添加"
    mod1.cmd.Parameters("@bh") = cmdHT.ToolTipText
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = lblBid.ToolTipText
    mod1.cmd.Parameters("@mt2") = lblZl.Caption
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mlt1") = txtNr.Text '人工内容
    mod1.cmd.Parameters("@mm1") = Val(txtWDJ.Text)
    mod1.cmd.Parameters("@mm2") = 0
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = Null
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
   

End Sub

Private Sub cmdWdel_Click()
Dim ii As Integer
On Error Resume Next
If Val(lblLid.Caption) = 0 Then Exit Sub
ii = MsgBox("是否作废此条记录?", vbInformation + vbYesNo, "询问")
If ii = vbYes Then
   
     timZm = 1
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单2011"
    mod1.cmd.Parameters("@NBLX") = "配件删除"
    mod1.cmd.Parameters("@bh") = cmdHT.ToolTipText
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = lblBid.ToolTipText
    mod1.cmd.Parameters("@mt2") = lblZl.Caption
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(lblLid.Caption)
    mod1.cmd.Parameters("@mm2") = Val(cmdHT.ToolTipText)
    mod1.cmd.Parameters("@mm3") = htRow
    mod1.cmd.Parameters("@mm9") = Lc
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = Null
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

Private Sub cmdWGx_Click()
Dim ii As Integer
On Error Resume Next
If Val(lblLid.Caption) = 0 Then Exit Sub
   
     timZm = 2
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单2011"
    mod1.cmd.Parameters("@NBLX") = "项目更新"
    mod1.cmd.Parameters("@bh") = cmdHT.ToolTipText
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = lblBid.ToolTipText
    mod1.cmd.Parameters("@mt2") = lblZl.Caption
    If htRow = 1 Or htRow = 2 Or htRow = 3 Or htRow = 4 Or htRow = 5 Or htRow = 6 Or htRow = 7 And lblZl.Caption <> "三菱" Or htRow = 8 And lblZl.Caption <> "松下" Or htRow >= 20 Or InStr(1, "维保大修其他人工压缩机维修保养中介业务分包运费吊装费工程人工", lblZl.Caption) > 0 Or _
     (Val(lblBid.ToolTipText) >= 20512 And lblZl.ToolTipText = True) Then
        mod1.cmd.Parameters("@mt3") = txtNr.Text
    Else
        mod1.cmd.Parameters("@mt3") = txtZBQ.Text
    End If
    mod1.cmd.Parameters("@mt4") = txtDrq.Text
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(lblLid.Caption)
    mod1.cmd.Parameters("@mm2") = Val(cmdHT.ToolTipText)
    mod1.cmd.Parameters("@mm3") = htRow
    mod1.cmd.Parameters("@mm5") = Val(txtMj.Text)
    If htRow = 1 Or htRow = 2 Or htRow = 3 Or htRow = 4 Or htRow = 5 Or htRow = 6 Or htRow = 7 And lblZl.Caption <> "三菱" Or htRow = 8 And lblZl.Caption <> "松下" Or htRow >= 20 Or InStr(1, "维保大修其他人工压缩机维修保养中介业务分包运费吊装费工程人工", lblZl.Caption) > 0 Or _
     (Val(lblBid.ToolTipText) >= 20512 And lblZl.ToolTipText = True) Then
        mod1.cmd.Parameters("@mm6") = Val(txtWcdj.Text)
        mod1.cmd.Parameters("@mm7") = Val(txtWDJ.Text)
        mod1.cmd.Parameters("@mm8") = 1
    Else
        mod1.cmd.Parameters("@mm6") = Val(txtDj.Text)
        mod1.cmd.Parameters("@mm7") = Val(txtJdj.Text)
        mod1.cmd.Parameters("@mm8") = Val(txtSL.Text)
    End If
    mod1.cmd.Parameters("@mm9") = Lc
    
    mod1.cmd.Parameters("@mm18") = Val(txtWBdj.Text)
    mod1.cmd.Parameters("@mm19") = Val(txtWBJe.Text)
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = Null
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
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Combo1_DblClick()

End Sub

Private Sub dtgBr_Click()

Dim NR As String
On Error Resume Next
dtgN.Row = dtgBr.Row
If htRow = 1 Or htRow = 2 Or htRow = 3 Or htRow = 4 Or htRow = 5 Or htRow = 6 Or htRow = 7 And lblZl.Caption <> "三菱" Or htRow = 8 And lblZl.Caption <> "松下" Or htRow >= 20 Or InStr(1, "维保大修其他人工压缩机维修保养中介业务分包运费吊装费工程人工", lblZl.Caption) > 0 Or _
     (Val(lblBid.ToolTipText) >= 20512 And lblZl.ToolTipText = True) Then
    dtgN.Col = 4: lblLid.Caption = dtgN.Text
    dtgN.Col = 0: txtNr.Text = dtgN.Text
    dtgN.Col = 1: txtWDJ.Text = dtgN.Text
    dtgN.Col = 6: txtWcdj.Text = dtgN.Text
    dtgN.Col = 7: txtWBdj.Text = dtgN.Text
    dtgN.Col = 8: txtWBJe.Text = dtgN.Text
Else
    optGy1.Value = False: optGy2.Value = False: optGy3.Value = False
    txtGy1.Text = "": txtGy2.Text = "": txtGY3.Text = ""
    txtGdj1.Text = "": txtGdj2.Text = "": txtGdj3.Text = ""
    optGy1.ForeColor = &H80000008: txtGy1.ForeColor = &H80000008: txtGdj1.ForeColor = &H80000008
    optGy2.ForeColor = &H80000008: txtGy2.ForeColor = &H80000008: txtGdj2.ForeColor = &H80000008
    optGy3.ForeColor = &H80000008: txtGY3.ForeColor = &H80000008: txtGdj3.ForeColor = &H80000008
    dtgN.Col = 10:
    lblLid.Caption = dtgN.Text
    dtgN.Col = 2: txtMj.Text = Val(dtgN.Text)
    dtgN.Col = 3: txtDj.Text = dtgN.Text
    dtgN.Col = 4: txtJdj.Text = Val(dtgN.Text)
    dtgN.Col = 5: txtSL.Text = Val(dtgN.Text)
    dtgN.Col = 7: txtDrq.Text = dtgN.Text
    dtgN.Col = 8: txtZBQ.Text = dtgN.Text
    dtgN.Col = 12: txtGy1.ToolTipText = Val(dtgN.Text)
    dtgN.Col = 13: txtGy2.ToolTipText = Val(dtgN.Text)
    dtgN.Col = 14: txtGY3.ToolTipText = Val(dtgN.Text)
    dtgN.Col = 15: txtGdj1.Text = Val(dtgN.Text)
    dtgN.Col = 16: txtGdj2.Text = Val(dtgN.Text)
    dtgN.Col = 17: txtGdj3.Text = Val(dtgN.Text)
    dtgN.Col = 18: txtGy1.Text = dtgN.Text
    dtgN.Col = 19: txtGy2.Text = dtgN.Text
    dtgN.Col = 20: txtGY3.Text = dtgN.Text
    dtgN.Col = 21
    If Val(dtgN.Text) <> 0 Then
        If Val(dtgN.Text) = txtGy1.ToolTipText Then
            optGy1.Value = True: optGy1.ForeColor = &HC00000: txtGy1.ForeColor = &HC00000: txtGdj1.ForeColor = &HC00000
        ElseIf Val(dtgN.Text) = txtGy2.ToolTipText Then
            optGy2.Value = True: optGy2.ForeColor = &HC00000: txtGy2.ForeColor = &HC00000: txtGdj2.ForeColor = &HC00000
        ElseIf Val(dtgN.Text) = txtGY3.ToolTipText Then
            optGy3.Value = True: optGy3.ForeColor = &HC00000: txtGY3.ForeColor = &HC00000: txtGdj3.ForeColor = &HC00000
        End If
    End If
    dtgN.Col = 22: txtBdj.Text = dtgN.Text
    dtgN.Col = 23: txtBje.Text = dtgN.Text
    dtgN.Col = 24: txtYH.Text = dtgN.Text
    dtgN.Col = 25: txtLx.Text = dtgN.Text
    dtgN.Col = 0: Bh = dtgN.Text
    dtgN.Col = 1: NR = dtgN.Text
    If Bh = "" And InStr(1, NR, "分包") > 0 Then
        'frmJ.Visible = False
        cmdGy.Visible = False
    Else
        frmJ.Visible = True
        cmdGy.Visible = True
    End If
End If
cmdGy.Visible = True
End Sub

Private Sub dtgBr_DblClick()
Dim Bh As String
Dim Ra
dtgN.Row = dtgBr.Row
If htRow = 1 Or htRow = 2 Or htRow = 3 Or htRow = 4 Or htRow = 5 Or htRow = 6 Or htRow = 7 And lblZl.Caption <> "三菱" Or htRow = 8 And lblZl.Caption <> "松下" Or htRow >= 20 Or InStr(1, "维保大修其他人工压缩机维修保养中介业务分包运费吊装费工程人工", lblZl.Caption) > 0 Or _
     (Val(lblBid.ToolTipText) >= 20512 And lblZl.ToolTipText = True) Then

Else
    If mod1.Bm = "配送中心" Or mod1.DName = "马晓聪" Or mod1.DName = "宋晓炯" Or mod1.Bm = "商务部" Or mod1.DName = "" Or Ywy = "吴金荣" Or mod1.DName = "朱婷婷" Or mod1.DName = "邹晨" Or mod1.DName = "沈维" Then
'''''        optGy1.Value = False: optGy2.Value = False: optGy3.Value = False
'''''        txtGy1.Text = "": txtGy2.Text = "": txtGY3.Text = ""
'''''        txtGdj1.Text = "": txtGdj2.Text = "": txtGdj3.Text = ""
'''''        optGy1.ForeColor = &H80000008: txtGy1.ForeColor = &H80000008: txtGdj1.ForeColor = &H80000008
'''''        optGy2.ForeColor = &H80000008: txtGy2.ForeColor = &H80000008: txtGdj2.ForeColor = &H80000008
'''''        optGy3.ForeColor = &H80000008: txtGY3.ForeColor = &H80000008: txtGdj3.ForeColor = &H80000008
'''''        dtgN.Col = 10: lblLid.Caption = dtgN.Text
'''''        dtgN.Col = 2: txtMj.Text = Val(dtgN.Text)
'''''        dtgN.Col = 3: txtDj.Text = Val(dtgN.Text)
'''''        dtgN.Col = 4: txtJdj.Text = Val(dtgN.Text)
'''''        dtgN.Col = 5: txtSL.Text = Val(dtgN.Text)
'''''        dtgN.Col = 7: txtDrq.Text = dtgN.Text
'''''        dtgN.Col = 8: txtZBQ.Text = dtgN.Text
'''''        dtgN.Col = 12: txtGy1.ToolTipText = Val(dtgN.Text)
'''''        dtgN.Col = 13: txtGy2.ToolTipText = Val(dtgN.Text)
'''''        dtgN.Col = 14: txtGY3.ToolTipText = Val(dtgN.Text)
'''''        dtgN.Col = 15: txtGdj1.Text = Val(dtgN.Text)
'''''        dtgN.Col = 16: txtGdj2.Text = Val(dtgN.Text)
'''''        dtgN.Col = 17: txtGdj3.Text = Val(dtgN.Text)
'''''        dtgN.Col = 18: txtGy1.Text = dtgN.Text
'''''        dtgN.Col = 19: txtGy2.Text = dtgN.Text
'''''        dtgN.Col = 20: txtGY3.Text = dtgN.Text
'''''        dtgN.Col = 21
'''''        If Val(dtgN.Text) <> 0 Then
'''''            If Val(dtgN.Text) = txtGy1.ToolTipText Then
'''''                optGy1.Value = True: optGy1.ForeColor = &HC00000: txtGy1.ForeColor = &HC00000: txtGdj1.ForeColor = &HC00000
'''''            ElseIf Val(dtgN.Text) = txtGy2.ToolTipText Then
'''''                optGy2.Value = True: optGy2.ForeColor = &HC00000: txtGy2.ForeColor = &HC00000: txtGdj2.ForeColor = &HC00000
'''''            ElseIf Val(dtgN.Text) = txtGY3.ToolTipText Then
'''''                optGy3.Value = True: optGy3.ForeColor = &HC00000: txtGY3.ForeColor = &HC00000: txtGdj3.ForeColor = &HC00000
'''''            End If
'''''        End If
            dtgN.Col = 0: Bh = dtgN.Text
            tt = "select oname from nlpmxc where bh='" & Bh & "'"
            Set mod1.HTP = CreateObject("adodb.recordset")
            mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
            On Error Resume Next
            Ra = mod1.HTP.GetRows
            mod1.HTP.Close
            Set mod1.HTP = Nothing
            txtXQ.Text = ""
            txtXQ.Text = "原厂编号：" & Ra(0, 0) & Chr(13) & Chr(10)
            txtXQ.Text = txtXQ.Text & "供应商1:" & txtGy1.Text & " 单价：" & txtGdj1.Text & Chr(13) & Chr(10) & _
                                    "供应商2:" & txtGy2.Text & " 单价：" & txtGdj2.Text & Chr(13) & Chr(10) & _
                                    "供应商3:" & txtGY3.Text & " 单价：" & txtGdj3.Text
            txtXQ.Visible = True
            
    End If
End If
End Sub

Private Sub dtgDW_DblClick()
dtgDW.Col = 0: txtXmmc.Text = dtgDW.Text
dtgDW.Col = 1: txtXmmc.ToolTipText = dtgDW.Text
dtgDW.Col = 2: txtYfadr.Text = dtgDW.Text
dtgDW.Visible = False
End Sub

Private Sub dtgGy_Click()
On Error Resume Next
If dtgGy.Row = 0 Then Exit Sub
If GyId = 0 Then GyId = 1
If GyId = 1 Then
    dtgGy.Col = 0: txtGy1.Text = dtgGy.Text
    dtgGy.Col = 1: txtGy1.ToolTipText = dtgGy.Text
    txtGdj1.Text = ""
ElseIf GyId = 2 Then
    dtgGy.Col = 0: txtGy2.Text = dtgGy.Text
    dtgGy.Col = 1: txtGy2.ToolTipText = dtgGy.Text
    txtGdj2.Text = ""
ElseIf GyId = 3 Then
    dtgGy.Col = 0: txtGY3.Text = dtgGy.Text
    dtgGy.Col = 1: txtGY3.ToolTipText = dtgGy.Text
    txtGdj3.Text = ""
End If
End Sub

Private Sub dtpBrq_CloseUp()
txtBrq.Text = dtpBrq.Value
cmdSave.Enabled = True
End Sub


Private Sub Form_Click()
frmCg.Visible = False
frmSd.Visible = False
frmWB.Visible = False
frmQm.Visible = False
frmYj.Visible = False
frmGY.Visible = False
txtXQ.Visible = False
dtgDW.Visible = False
End Sub
Public Sub dtgGYFF()
dtgGy.Clear
dtgGy.Rows = 50
dtgGy.Cols = 2
dtgGy.Row = 0
dtgGy.Col = 0: dtgGy.Text = "供应商名称（鼠标双击选择）": dtgGy.CellFontBold = True
dtgGy.ColWidth(1) = 0
dtgGy.ColWidth(0) = 3480

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
 dtgP.ColWidth(3) = 5100: dtgP.ColWidth(4) = 1035
For oo = 0 To 4
    dtgP.Col = oo
    dtgP.CellFontBold = True
Next
End Sub
Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
frmWB.Left = 0: frmWB.Top = 3660
frmCg.Left = 0: frmCg.Top = 4140
frmSd.Left = 0: frmSd.Top = 5550
frmJ.Left = 1920: frmJ.Top = 240
Me.dtpBrq.Value = Date
Me.txtXQ.Visible = False
End Sub
Public Sub Qing()
On Error Resume Next
FmxcXJ.Visible = True
txtXmmc.Text = "": txtXmmc.ToolTipText = ""
lblYwy.Caption = "": lblYwy.ToolTipText = ""
lblRq.Caption = ""
lblWhg.Caption = ""
lblZl.Caption = ""
lblZl.ToolTipText = ""
Lc = 1
LCRen = ""
LCUid = ""
Fwid = 0
txtBz.Text = ""
cmdHT.ToolTipText = ""
htRow = 0
txtYfadr.Text = ""
txtJHg.Text = ""
lblBid.Caption = "": lblBid.ToolTipText = ""
txtBrq.Text = ""
dtpBrq.Visible = False
frmCg.Visible = False
frmWB.Visible = False
frmSd.Visible = False
Call MXQing
cmdSave.Enabled = False
cmdD.Enabled = False
dtgBr.Clear
dtgN.Clear
Call dtgPFF
frmJ.Visible = True
'''lblDj.Visible = False
'''txtDj.Visible = False
cmdDht.Visible = False
lblWcdj.Visible = False
txtWcdj.Visible = False
lblWdj.Visible = True
txtWDJ.Visible = True
optGy1.Value = False: optGy2.Value = False: optGy3.Value = False
txtGy1.Text = "": txtGy2.Text = "": txtGY3.Text = ""
txtGdj1.Text = "": txtGdj2.Text = "": txtGdj3.Text = ""
optGy1.ForeColor = &H80000008: txtGy1.ForeColor = &H80000008: txtGdj1.ForeColor = &H80000008
optGy2.ForeColor = &H80000008: txtGy2.ForeColor = &H80000008: txtGdj2.ForeColor = &H80000008
optGy3.ForeColor = &H80000008: txtGY3.ForeColor = &H80000008: txtGdj3.ForeColor = &H80000008
frmGY.Visible = False
txtXQ.Text = ""
txtXQ.Visible = False
txtCBZ.Text = ""
txtCBZ.Locked = True
frmCGRZ.Visible = False
txtCED.Text = "": txtCED.Locked = True
txtBhg.Text = ""
txtBJ.Text = ""
Call Me.dtgGYFF
End Sub

Public Sub dtgbrFF()
Dim oo As Integer
On Error Resume Next
If htRow = 1 Or htRow = 2 Or htRow = 3 Or htRow = 4 Or htRow = 5 Or htRow = 6 Or htRow = 7 And lblZl.Caption <> "三菱" Or _
htRow = 8 And lblZl.Caption <> "松下" Or htRow >= 20 Or htRow = 7.14 Or _
     InStr(1, "维保大修其他人工压缩机维修保养中介业务分包运费吊装费工程人工", lblZl.Caption) > 0 Or _
     (Val(lblBid.ToolTipText) >= 20512 And lblZl.ToolTipText = True) Then
    dtgBr.Cols = 9
    dtgBr.Row = 0
    dtgBr.Col = 0: dtgBr.Text = "业务内容(" & lblZl.Caption & ")": dtgBr.CellFontBold = True
    If InStr(1, lblZl.Caption, "预估") > 0 Then
        dtgBr.Col = 1: dtgBr.Text = "速达金额"
    Else
        dtgBr.Col = 1: dtgBr.Text = "基准价"
    End If
    
    dtgBr.CellFontBold = True
    dtgBr.Col = 2: dtgBr.Text = "备注": dtgBr.CellFontBold = True
    dtgBr.Col = 3: dtgBr.Text = "有效": dtgBr.CellFontBold = True
    dtgBr.Col = 4: dtgBr.Text = "Lid"
    dtgBr.Col = 5: dtgBr.Text = "速达小计": dtgBr.CellFontBold = True:: dtgBr.CellForeColor = &H8000&
    dtgBr.Col = 6: dtgBr.Text = "单价"
    dtgBr.ColWidth(0) = 11760
    dtgBr.ColWidth(1) = -1
    dtgBr.ColWidth(2) = 0
    dtgBr.ColWidth(3) = -1 '备注字段不需要
    dtgBr.ColWidth(4) = 0
    dtgN.Cols = 9
    For oo = 1 To dtgBr.Rows + 1
        dtgBr.RowHeight(oo) = dtgBr.RowHeight(0) * 2
    Next
    If mod1.GxName = "报价功能" And mod1.GXF = True And lblZl.Caption = "询价指令" Then
        dtgBr.ColWidth(0) = 10710
        dtgBr.ColWidth(5) = 0
        dtgBr.ColWidth(6) = 0
        dtgBr.Col = 7: dtgBr.Text = "对外单价": dtgBr.CellFontBold = True: dtgBr.CellForeColor = &H8000&
        dtgBr.Col = 8: dtgBr.Text = "对外金额": dtgBr.CellFontBold = True: dtgBr.CellForeColor = &H8000&
        dtgBr.ColWidth(7) = -1
        dtgBr.ColWidth(8) = -1
    Else
        dtgBr.ColWidth(0) = 11760
        dtgBr.ColWidth(5) = -1
        dtgBr.ColWidth(6) = -1
        dtgBr.ColWidth(7) = 0
        dtgBr.ColWidth(8) = 0
    End If
Else
    dtgBr.Cols = 26
    dtgBr.Row = 0
    dtgBr.Col = 0: dtgBr.Text = "编号": dtgBr.CellFontBold = True
    dtgBr.Col = 1: dtgBr.Text = "货品(" & lblZl.Caption & ")": dtgBr.CellFontBold = True
    dtgBr.Col = 2: dtgBr.Text = "市场价": dtgBr.CellFontBold = True
    dtgBr.Col = 3: dtgBr.Text = "单价":  dtgBr.CellFontBold = True
    dtgBr.Col = 4: dtgBr.Text = "基准单价": dtgBr.CellFontBold = True
    dtgBr.Col = 5: dtgBr.Text = "数量": dtgBr.CellFontBold = True
    dtgBr.Col = 6: dtgBr.Text = "小计": dtgBr.CellFontBold = True
    dtgBr.Col = 7: dtgBr.Text = "到货期": dtgBr.CellFontBold = True
    dtgBr.Col = 8: dtgBr.Text = "质保期": dtgBr.CellFontBold = True
    dtgBr.Col = 9: dtgBr.Text = "有效": dtgBr.CellFontBold = True
    dtgBr.Col = 10: dtgBr.Text = "Lid": dtgBr.CellFontBold = True
    dtgBr.Col = 11: dtgBr.Text = "速达金额": dtgBr.CellFontBold = True:: dtgBr.CellForeColor = &H8000&
    dtgBr.ColWidth(10) = 0
    dtgBr.ColWidth(0) = -1
    dtgBr.ColWidth(2) = 0
    dtgBr.ColWidth(3) = 0
    dtgBr.ColWidth(7) = 1500
    dtgBr.ColWidth(8) = 1500
    dtgBr.ColWidth(1) = 5700
    dtgBr.ColWidth(4) = -1
    dtgBr.ColWidth(5) = -1
    dtgBr.ColWidth(6) = -1
    dtgBr.ColWidth(9) = -1
'''    dtgBr.ColWidth(8) = 0
'''    dtgBr.ColWidth(7) = 0
    dtgBr.ColWidth(12) = 0
    dtgBr.ColWidth(13) = 0
    dtgBr.ColWidth(14) = 0
    dtgBr.ColWidth(15) = 0
    dtgBr.ColWidth(16) = 0
    dtgBr.ColWidth(17) = 0
    dtgBr.ColWidth(18) = 0
    dtgBr.ColWidth(19) = 0
    dtgBr.ColWidth(20) = 0
    dtgBr.ColWidth(21) = 0
    dtgN.Cols = 26
    For oo = 1 To dtgBr.Rows + 1
        dtgBr.RowHeight(oo) = dtgBr.RowHeight(0)
    Next
    
    If mod1.GxName = "报价功能" And mod1.GXF = True And lblZl.Caption = "询价指令" Then
        dtgBr.ColWidth(1) = 4685
        dtgBr.ColWidth(22) = -1
        dtgBr.ColWidth(23) = -1
        dtgBr.ColWidth(24) = -1
        dtgBr.ColWidth(25) = 3300
        dtgBr.ColWidth(11) = 0
        dtgBr.Col = 22: dtgBr.Text = "对外单价": dtgBr.CellFontBold = True: dtgBr.CellForeColor = &H8000&
        dtgBr.Col = 23: dtgBr.Text = "对外金额": dtgBr.CellFontBold = True: dtgBr.CellForeColor = &H8000&
        dtgBr.Col = 24: dtgBr.Text = "优惠价": dtgBr.CellFontBold = True: dtgBr.CellForeColor = &H8000&
        dtgBr.Col = 25: dtgBr.Text = "业务类型": dtgBr.CellFontBold = True: dtgBr.CellForeColor = &H8000&
    Else
        dtgBr.ColWidth(1) = 5700
        dtgBr.ColWidth(22) = 0
        dtgBr.ColWidth(23) = 0
        dtgBr.ColWidth(24) = 0
        dtgBr.ColWidth(25) = 0
        dtgBr.ColWidth(11) = -1
        dtgBr.Col = 11: dtgBr.Text = "速达金额": dtgBr.CellFontBold = True: dtgBr.CellForeColor = &H8000&
    End If
End If
dtgBr.Rows = 50: dtgN.Rows = 50
End Sub

Public Sub Bound(Bid As Long)
Dim tt As String
Dim Ra, Rb, Rz, RC, RD, RE
Dim Lb As Integer
Dim Lz As Integer
Call Qing
tt = "select xmmc,xid,ywy,uid,rq,whg,zl,lc,lcren,lcuid,fwid,bz,htbh,htrow,yfadr,jhg,brq,bid,lx,sdje,bje from xunjiaD where bid=" & Bid & ";" & _
    "select ljbh,detail,mj,dj,jdj,sl,jhg,drq,zbq,delf,lid,ljmc,gyid1,gyid2,gyid3,gdj1,gdj2,gdj3,mc1,mc2,mc3,gyid,sddj,sdxg,sdyh,ywlx  from XJDetail where bid=" & Bid & " order by delf desc,lid desc;" & _
    "select trq,ywy,zn,bz,tf from pizu where bh='" & Bid & "' and yid=43 order by pid desc;" & _
    "select rq,nr,cname from xunjiaCN where bid=" & Bid & " order by cid desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rz = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
RC = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing

Lb = UBound(Rb, 2) + 1
Lz = UBound(Rz, 2) + 1
txtXmmc.Text = Ra(0, 0)
txtXmmc.ToolTipText = Ra(1, 0)
lblYwy.Caption = Ra(2, 0)
lblYwy.ToolTipText = Ra(3, 0)
lblRq.Caption = DateSerial(Year(Ra(4, 0)), Month(Ra(4, 0)), Day(Ra(4, 0)))
lblWhg.Caption = Ra(5, 0)
lblZl.Caption = Ra(6, 0)
If lblZl.Caption = "" And FmxcNew.Visible = True Then
    lblZl.Caption = FmxcNew.XJZL
End If
Lc = Ra(7, 0)
LCRen = Ra(8, 0)
lblTX.Caption = "流程至: " & LCRen
If Lc = 100 Then lblTX.Caption = "审核完毕!"
LCUid = Ra(9, 0)
Fwid = Ra(10, 0)
txtBz.Text = Ra(11, 0)
cmdHT.ToolTipText = Ra(12, 0)
htRow = Ra(13, 0)
txtYfadr.Text = Ra(14, 0)
txtJHg.Text = Ra(15, 0)
lblBh.Caption = "XJD" & Bid
lblBh.ToolTipText = Bid
txtBrq.Text = Ra(16, 0)
lblBid.ToolTipText = Ra(17, 0): lblBid.Caption = "XJD" & Ra(17, 0)
lblZl.ToolTipText = Ra(18, 0)
Call CGBound(RC)

Call dtgBrBound(Rb, Lb)
Call QMBound(Bid, Rz, Lz)

If mod1.GxName = "报价功能" And mod1.GXF = True And lblZl.Caption = "询价指令" Then
    lblBj.Visible = True: txtBJ.Visible = True
    lblBhg.Visible = True
    txtBhg.Visible = True
    txtBJ.Text = Ra(20, 0)
    txtBhg.Text = Ra(19, 0)
Else
    lblBhg.Visible = False
    txtBhg.Visible = False
    lblBj.Visible = False
    txtBJ.Visible = False
End If

If mod1.DName = "朱婷婷" Or mod1.DName = "汪燕明" Or mod1.DName = "吴金荣" Or mod1.DName = "吴金荣" Then
    Call Me.Xian
End If
End Sub


Public Sub dtgBrBound(Rb, Lb As Integer)
Dim oo As Integer
On Error Resume Next
dtgBr.Clear: dtgBr.Visible = False: dtgN.Clear
Call dtgbrFF
dtgBr.Rows = Lb + 50
dtgN.Rows = dtgBr.Rows
If (htRow = 1 Or htRow = 2 Or htRow = 3 Or htRow = 4 Or htRow = 5 Or htRow = 6 Or htRow = 7 Or htRow = 7.14 And lblZl.Caption <> "三菱" Or htRow = 8 And lblZl.Caption <> "松下" Or htRow >= 20 Or InStr(1, "维保大修其他人工压缩机维修保养中介业务分包运费吊装费工程人工", lblZl.Caption) > 0 Or _
     (Val(lblBid.ToolTipText) >= 20512 And lblZl.ToolTipText = True)) And Not (lblZl = "分包->工程人工" And Val(lblBid.ToolTipText) > 22211 And Val(lblBid.ToolTipText) < 22670) Then
    For oo = 1 To Lb + 1
        dtgBr.Row = oo: dtgBr.RowHeight(oo) = dtgBr.RowHeight(0) * 2
        dtgBr.Col = 0
        dtgN.Row = oo
        If IsNull(Rb(11, oo - 1)) = False And Rb(11, oo - 1) <> "" Then '兼顾版本,内容可能在ljmc与zbq两者之一的字段
            dtgBr.Text = Rb(11, oo - 1): dtgN.Col = 0: dtgN.Text = dtgBr.Text
        Else
            dtgBr.Text = Rb(8, oo - 1): dtgN.Col = 0: dtgN.Text = dtgBr.Text
        End If
        dtgBr.Col = 1: dtgBr.Text = Rb(4, oo - 1)

        dtgBr.Col = 2: dtgBr.Text = Rb(8, oo - 1)
        dtgBr.Col = 3: dtgBr.Text = Rb(9, oo - 1)
        dtgBr.Col = 4: dtgBr.Text = Rb(10, oo - 1)
        dtgBr.Col = 6: dtgBr.Text = Rb(3, oo - 1)
        dtgBr.RowHeight(oo) = dtgBr.RowHeight(0) * 2

        
        dtgN.Col = 1: dtgN.Text = Rb(4, oo - 1)
        dtgN.Col = 2: dtgN.Text = Rb(8, oo - 1)
        dtgN.Col = 3: dtgN.Text = Rb(9, oo - 1)
        dtgN.Col = 4: dtgN.Text = Rb(10, oo - 1)
        dtgN.Col = 6: dtgN.Text = Rb(3, oo - 1)
        If mod1.GxName = "报价功能" And mod1.GXF = True And lblZl.Caption = "询价指令" Then
            dtgBr.Col = 7: dtgBr.Text = Rb(22, oo - 1): dtgN.Col = 7: dtgN.Text = Rb(22, oo - 1): dtgBr.CellForeColor = &H8000&
            dtgBr.Col = 8: dtgBr.Text = Rb(23, oo - 1): dtgN.Col = 8: dtgN.Text = Rb(23, oo - 1): dtgBr.CellForeColor = &H8000&
        End If
        '检查作废为红色字
        dtgBr.Col = 3
        If dtgBr.Text = "False" Then
            dtgBr.Col = 0: dtgBr.CellForeColor = &HFF&
            dtgBr.Col = 1: dtgBr.CellForeColor = &HFF&
            dtgBr.Col = 2: dtgBr.CellForeColor = &HFF&
            dtgBr.Col = 3: dtgBr.CellForeColor = &HFF&
            dtgBr.Col = 4: dtgBr.CellForeColor = &HFF&
        Else
            dtgBr.Col = 0: dtgBr.CellForeColor = &H0&
            dtgBr.Col = 1: dtgBr.CellForeColor = &H0&
            dtgBr.Col = 2: dtgBr.CellForeColor = &H0&
            dtgBr.Col = 3: dtgBr.CellForeColor = &H0&
            dtgBr.Col = 4: dtgBr.CellForeColor = &H0&
        End If
    Next
    
Else
    For oo = 1 To Lb + 1
        dtgBr.Row = oo
        dtgBr.Col = 0: dtgBr.Text = Rb(0, oo - 1)
        dtgBr.Col = 1: dtgBr.Text = Rb(1, oo - 1)
        frmZu.lblDtg.Caption = dtgBr.Text
        dtgBr.RowHeight(oo) = frmZu.lblDtg.Height
        dtgBr.Col = 2: dtgBr.Text = Rb(2, oo - 1)
        dtgBr.Col = 3: dtgBr.Text = Rb(3, oo - 1)
        dtgBr.Col = 4: dtgBr.Text = Rb(4, oo - 1)
        dtgBr.Col = 5: dtgBr.Text = Rb(5, oo - 1)
        dtgBr.Col = 6: dtgBr.Text = Rb(6, oo - 1)
        dtgBr.Col = 7: dtgBr.Text = Rb(7, oo - 1)
        dtgBr.Col = 8: dtgBr.Text = Rb(8, oo - 1)
        dtgBr.Col = 9: dtgBr.Text = Rb(9, oo - 1)
        dtgBr.Col = 10: dtgBr.Text = Rb(10, oo - 1)
        dtgBr.Col = 12: dtgBr.Text = Rb(12, oo - 1)
        dtgBr.Col = 13: dtgBr.Text = Rb(13, oo - 1)
        dtgBr.Col = 14: dtgBr.Text = Rb(14, oo - 1)
        dtgBr.Col = 15: dtgBr.Text = Rb(15, oo - 1)
        dtgBr.Col = 16: dtgBr.Text = Rb(16, oo - 1)
        dtgBr.Col = 17: dtgBr.Text = Rb(17, oo - 1)
        dtgBr.Col = 18: dtgBr.Text = Rb(18, oo - 1)
        dtgBr.Col = 19: dtgBr.Text = Rb(19, oo - 1)
        dtgBr.Col = 20: dtgBr.Text = Rb(20, oo - 1)
        dtgBr.Col = 21: dtgBr.Text = Rb(21, oo - 1)
        dtgN.Row = oo
        dtgN.Col = 0: dtgN.Text = Rb(0, oo - 1)
        dtgN.Col = 1: dtgN.Text = Rb(1, oo - 1)
        dtgN.Col = 2: dtgN.Text = Rb(2, oo - 1)
        dtgN.Col = 3: dtgN.Text = Rb(3, oo - 1)
        dtgN.Col = 4: dtgN.Text = Rb(4, oo - 1)
        dtgN.Col = 5: dtgN.Text = Rb(5, oo - 1)
        dtgN.Col = 6: dtgN.Text = Rb(6, oo - 1)
        dtgN.Col = 7: dtgN.Text = Rb(7, oo - 1)
        dtgN.Col = 8: dtgN.Text = Rb(8, oo - 1)
        dtgN.Col = 9: dtgN.Text = Rb(9, oo - 1)
        dtgN.Col = 10: dtgN.Text = Rb(10, oo - 1)
        dtgN.Col = 12: dtgN.Text = Rb(12, oo - 1)
        dtgN.Col = 13: dtgN.Text = Rb(13, oo - 1)
        dtgN.Col = 14: dtgN.Text = Rb(14, oo - 1)
        dtgN.Col = 15: dtgN.Text = Rb(15, oo - 1)
        dtgN.Col = 16: dtgN.Text = Rb(16, oo - 1)
        dtgN.Col = 17: dtgN.Text = Rb(17, oo - 1)
        dtgN.Col = 18: dtgN.Text = Rb(18, oo - 1)
        dtgN.Col = 19: dtgN.Text = Rb(19, oo - 1)
        dtgN.Col = 20: dtgN.Text = Rb(20, oo - 1)
        dtgN.Col = 21: dtgN.Text = Rb(21, oo - 1)
        If mod1.GxName = "报价功能" And mod1.GXF = True And lblZl.Caption = "询价指令" Then
            dtgBr.Col = 22: dtgBr.Text = Rb(22, oo - 1): dtgN.Col = 22: dtgN.Text = Rb(22, oo - 1): dtgBr.CellForeColor = &H8000&
            dtgBr.Col = 23: dtgBr.Text = Rb(23, oo - 1): dtgN.Col = 23: dtgN.Text = Rb(23, oo - 1): dtgBr.CellForeColor = &H8000&
            dtgBr.Col = 24: dtgBr.Text = Rb(24, oo - 1): dtgN.Col = 24: dtgN.Text = Rb(24, oo - 1): dtgBr.CellForeColor = &H8000&
            dtgBr.Col = 25: dtgBr.Text = Rb(25, oo - 1): dtgN.Col = 25: dtgN.Text = Rb(25, oo - 1)
            
        End If
''''        dtgBr.RowHeight(oo) = dtgBr.RowHeight(0) * 2
        '检查作废为红色字
        dtgBr.Col = 9
        If dtgBr.Text = "False" Then
            dtgBr.Col = 0: dtgBr.CellForeColor = &HFF&
            dtgBr.Col = 1: dtgBr.CellForeColor = &HFF&
            dtgBr.Col = 2: dtgBr.CellForeColor = &HFF&
            dtgBr.Col = 3: dtgBr.CellForeColor = &HFF&
            dtgBr.Col = 4: dtgBr.CellForeColor = &HFF&
            dtgBr.Col = 5: dtgBr.CellForeColor = &HFF&
            dtgBr.Col = 6: dtgBr.CellForeColor = &HFF&
            dtgBr.Col = 7: dtgBr.CellForeColor = &HFF&
            dtgBr.Col = 8: dtgBr.CellForeColor = &HFF&
            dtgBr.Col = 9: dtgBr.CellForeColor = &HFF&
        Else
            dtgBr.Col = 0: dtgBr.CellForeColor = &H0&
            dtgBr.Col = 1: dtgBr.CellForeColor = &H0&
            dtgBr.Col = 2: dtgBr.CellForeColor = &H0&
            dtgBr.Col = 3: dtgBr.CellForeColor = &H0&
            dtgBr.Col = 4: dtgBr.CellForeColor = &H0&
            dtgBr.Col = 5: dtgBr.CellForeColor = &H0&
            dtgBr.Col = 6: dtgBr.CellForeColor = &H0&
            dtgBr.Col = 7: dtgBr.CellForeColor = &H0&
            dtgBr.Col = 8: dtgBr.CellForeColor = &H0&
            dtgBr.Col = 9: dtgBr.CellForeColor = &H0&
        End If
    Next
End If
dtgBr.Visible = True
If mod1.DName = "朱婷婷" Or mod1.DName = "汪燕明" Or mod1.DName = "吴金荣" Or mod1.DName = "吴金荣" Then
    Call Me.Xian
End If
End Sub

Public Sub QMBound(Bid As Long, Rz, Lz As Integer)
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

Private Sub Form_Unload(Cancel As Integer)

Me.Visible = False
If Dialog.Visible = True Then
    Call mod1.refEnvent(1)
    Dialog.ZOrder 0
    Dialog.Enabled = True
 
ElseIf frmGxBiao.Visible = True Then
    If frmGxBNew.Visible = True Then
        frmGxBNew.Show
        frmGxBNew.ZOrder 0
        Exit Sub
    End If
    frmGxBiao.Enabled = True
    frmGxBiao.ZOrder 0
ElseIf FmxcNew.Visible = True Then
    FmxcNew.Show
    FmxcNew.ZOrder 0

End If
Cancel = True
End Sub

Private Sub Label16_DblClick()
If mod1.DName = "宋晓炯" Or mod1.DName = "马晓聪" Or mod1.DName = "乔继敏" Or mod1.DName = "王全红" Then
    frmJ.Visible = False
    lblDj.Visible = True
    txtDj.Visible = True
End If
End Sub


Private Sub lblDj_Click()
If mod1.DName = "宋晓炯" Or mod1.DName = "马晓聪" Or mod1.DName = "乔继敏" Or mod1.DName = "王全红" Then
        frmJ.Visible = True
        lblDj.Visible = False
        txtDj.Visible = False
End If
End Sub


Private Sub lblWcdj_Click()
If mod1.DName = "宋晓炯" Or mod1.DName = "马晓聪" Or mod1.DName = "乔继敏" Or mod1.DName = "王全红" Then
    lblWcdj.Visible = False
    txtWcdj.Visible = False
    lblWdj.Visible = True
    txtWDJ.Visible = True
End If
End Sub

Private Sub lblWdj_Click()
If mod1.DName = "宋晓炯" Or mod1.DName = "马晓聪" Or mod1.DName = "乔继敏" Or mod1.DName = "王全红" Then
    lblWdj.Visible = False
    txtWDJ.Visible = False
    lblWcdj.Visible = True
    txtWcdj.Visible = True
End If
End Sub

Private Sub optGy1_Click()
optGy1.ForeColor = &H80000008: txtGy1.ForeColor = &H80000008: txtGdj1.ForeColor = &H80000008
optGy2.ForeColor = &H80000008: txtGy2.ForeColor = &H80000008: txtGdj2.ForeColor = &H80000008
optGy3.ForeColor = &H80000008: txtGY3.ForeColor = &H80000008: txtGdj3.ForeColor = &H80000008
If optGy1.Value = True Then
    optGy1.ForeColor = &HC00000: txtGy1.ForeColor = &HC00000: txtGdj1.ForeColor = &HC00000
    txtDj.Text = txtGdj1.Text
    txtJdj.Text = Round(Val(txtDj.Text) * mod1.JZ, 2)
End If
End Sub

Private Sub optGy2_Click()
optGy1.ForeColor = &H80000008: txtGy1.ForeColor = &H80000008: txtGdj1.ForeColor = &H80000008
optGy2.ForeColor = &H80000008: txtGy2.ForeColor = &H80000008: txtGdj2.ForeColor = &H80000008
optGy3.ForeColor = &H80000008: txtGY3.ForeColor = &H80000008: txtGdj3.ForeColor = &H80000008
If optGy2.Value = True Then
    optGy2.ForeColor = &HC00000: txtGy2.ForeColor = &HC00000: txtGdj2.ForeColor = &HC00000
    txtDj.Text = txtGdj2.Text
    txtJdj.Text = Round(Val(txtDj.Text) * mod1.JZ, 2)
End If
End Sub


Private Sub optGy3_Click()
optGy1.ForeColor = &H80000008: txtGy1.ForeColor = &H80000008: txtGdj1.ForeColor = &H80000008
optGy2.ForeColor = &H80000008: txtGy2.ForeColor = &H80000008: txtGdj2.ForeColor = &H80000008
optGy3.ForeColor = &H80000008: txtGY3.ForeColor = &H80000008: txtGdj3.ForeColor = &H80000008
If optGy3.Value = True Then
    optGy3.ForeColor = &HC00000: txtGY3.ForeColor = &HC00000: txtGdj3.ForeColor = &HC00000
    txtDj.Text = txtGdj3.Text
    txtJdj.Text = Round(Val(txtDj.Text) * mod1.JZ, 2)
End If
End Sub


Private Sub Option1_Click()
Call dtgXZ(False)
End Sub

Private Sub optV1_Click()
Call dtgXZ(True)
End Sub

Private Sub timQuit_Timer()
Dim Rb, Rz, Rf
Dim Lb As Integer
Dim Lz As Integer
On Error Resume Next
Dim oo As Integer
Dim jj As Integer
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
If timZm = 1 Or timZm = 2 Or timZm = 5 Then '配件删除或更新或人工添加
    tt = "select ljbh,detail,mj,dj,jdj,sl,jhg,drq,zbq,delf,lid,ljmc,gyid1,gyid2,gyid3,gdj1,gdj2,gdj3,mc1,mc2,mc3,gyid,sddj,sdxg,sdyh,ywlx  from XJDetail  where bid=" & Val(FmxcXJ.lblBid.ToolTipText) & " order by delf desc,lid desc"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdTex
    Rb = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    Lb = UBound(Rb, 2)
    Call FmxcXJ.dtgBrBound(Rb, Lb)
    Call MXQing
    If FmxcNew.Visible = True Then
        Call FmxcNew.LXBound(Rf, Rg)
    End If
ElseIf timZm = 3 Then '询价单删除
    Me.Visible = False
    If FmxcNew.Visible = True Then
        FmxcNew.dtgLx.Row = htRow
        FmxcNew.dtgLx.Col = 2: FmxcNew.dtgLx.Text = ""
        FmxcNew.dtgLx.Col = 3: FmxcNew.dtgLx.Text = ""
        FmxcNew.dtgLx.Col = 4: FmxcNew.dtgLx.Text = ""
        Call FmxcNew.Cale
        Call FmxcNew.LXBound(Rf, Rg)
    End If
    Call Me.Qing
    Me.Visible = False
ElseIf timZm = 6 Then '签字
    tt = "select trq,ywy,zn,bz,tf from pizu where bh='" & lblBid.ToolTipText & "' and yid=43 order by pid desc"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Rz = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    Lz = UBound(Rz, 2) + 1
    Call QMBound(Val(lblBid.ToolTipText), Rz, Lz)
    If FmxcNew.Visible = True Then
        Call FmxcNew.LXBound(Rf, Rg)
    End If
ElseIf timZm = 7 Then '导入合同
    cmdDht.Visible = False
    If Lc = 1 Then
        LCRen = lblYwy.Caption
        LCUid = lblYwy.ToolTipText
        MsgBox ("关联合同成功！此询价单将重新审核！")
    Else
        MsgBox ("关联合同成功！")
    End If
    If FmxcNew.Visible = True Then
            FmxcNew.dtgLx.Row = htRow
            FmxcNew.dtgLx.Col = 4
            FmxcNew.dtgLx.Text = "XJD" & Trim(lblBid.ToolTipText)
    End If
ElseIf timZm = 9 Then '导入合同2013
    cmdDht.Visible = False
    If Lc = 1 Then
        LCRen = lblYwy.Caption
        LCUid = lblYwy.ToolTipText
        MsgBox ("关联合同成功！此询价单将重新审核！")
    Else
        MsgBox ("关联合同成功！")
    End If
End If
timQuit.Enabled = False
End Sub



Public Sub MXQing()
txtMj.Text = ""
txtSL.Text = ""
txtDj.Text = ""
txtJdj.Text = ""
txtDrq.Text = ""
txtZBQ.Text = ""
lblLid.Caption = ""
txtNr.Text = ""
txtWDJ.Text = ""
txtWcdj.Text = ""
txtBdj.Text = ""
txtBje.Text = ""
txtWBdj.Text = ""
txtWBJe.Text = ""
txtYH.Text = ""
End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
Dim Rf
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh,mm1,mm2,mt1,mt2,mt3 from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Or timZm = 2 Then
        txtJHg.Text = mod1.WP.Fields("mm1").Value
'''''        txtBhg.Text = mod1.WP.Fields("mm2").Value
'''''        txtBje.Text = mod1.WP.Fields("mm3").Value
        txtBJ.Text = mod1.WP.Fields("mm3").Value
        txtBhg.Text = mod1.WP.Fields("mm2").Value
    ElseIf timZm = 6 Then '签名

                Lc = mod1.WP.Fields("mm1").Value
                Fwid = mod1.WP.Fields("mm2").Value
                LCRen = mod1.WP.Fields("mt1").Value
                LCUid = mod1.WP.Fields("mt2").Value
                lblTX.Caption = "下一流程,将跳至:" & LCRen
                If Lc = 100 Then lblTX.Caption = "审核完毕!"
                If FmxcNew.Visible = True Then
                    FmxcNew.dtgLx.Row = htRow
                    FmxcNew.dtgLx.Col = 2
                    FmxcNew.dtgLx.Text = txtJHg.Text
                End If
    ElseIf timZm = 7 Then '导入合同
        Lc = mod1.WP.Fields("mm1").Value
        If Lc = 100 Then
            If FmxcNew.Visible = True Then
                FmxcNew.dtgLx.Col = 2
                FmxcNew.dtgLx.Row = FmxcNew.NewId
                FmxcNew.dtgLx.Text = txtJHg.Text
            End If
        End If
    ElseIf timZm = 9 Then ' 导入合同2013
        Lc = mod1.WP.Fields("mm1").Value
        If Lc = 100 Then
            If FmxcNew.Visible = True Then
                Call FmxcNew.LXBound(Rf, Rg)
            End If
        End If
        cmdHT.ToolTipText = THid
    ElseIf timZm = 8 Then
        tt = "select rq,nr,cname from xunjiaCN where bid=" & Val(lblBid.ToolTipText) & " order by cid desc"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        On Error Resume Next
        RC = mod1.HTP.GetRows
        Set mod1.HTP = mod1.HTP.NextRecordset
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        Call CGBound(RC)
    
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


Private Sub txtBdj_Change()
txtBje.Text = Val(txtBdj.Text) * Val(txtSL.Text)
txtYH.Text = txtBje.Text
End Sub

Private Sub txtBdj_LostFocus()
txtBje.Text = Val(txtBdj.Text) * Val(txtSL.Text)
txtYH.Text = txtBje.Text
End Sub

Private Sub txtDj_Change()
txtJdj.Text = Round(Val(txtDj.Text) * mod1.JZ, 2)
End Sub


Private Sub txtDj_DblClick()
Dim tt As String
Dim Ra
If Left(Bh, 1) = "3" Then Exit Sub
If Val(txtDj.Text) > 0 Then Exit Sub
If Bh = "" Then Exit Sub
tt = "select dj from xunjiamx where ljbh='" & Bh & "' and dj>0 and delf=1 and bid<>" & Val(lblBid.ToolTipText) & " order by lid desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
If IsNull(Ra(0, 0)) = True Then
    txtDj.Text = "无历史记录"
Else
    txtDj.Text = Ra(0, 0)
End If

End Sub


Private Sub txtGdj1_Change()
If optGy1.Value = True Then
    txtDj.Text = txtGdj1.Text
    txtJdj.Text = Round(Val(txtDj.Text) * mod1.JZ, 2)
End If
End Sub

Private Sub txtGdj2_Change()
If optGy2.Value = True Then
    txtDj.Text = txtGdj2.Text
    txtJdj.Text = Round(Val(txtDj.Text) * mod1.JZ, 2)
End If
End Sub


Private Sub txtGdj3_Change()
If optGy3.Value = True Then
    txtDj.Text = txtGdj3.Text
    txtJdj.Text = Round(Val(txtDj.Text) * mod1.JZ, 2)
End If
End Sub


Private Sub txtGy_Change()
Dim tt As String
Dim Ra
Dim La As Long
Dim oo As Long
If Len(txtGy.Text) < 2 Then Exit Sub
'tt = "select mc,gid from gymxc where mc like '%" & txtGy.Text & "%' and delf=1 and lc=100"
tt = "select mc,gid from gymxc where mc like '%" & txtGy.Text & "%' and delf=1 and lc>=2"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
Call Me.dtgGYFF
For oo = 1 To La
    dtgGy.Row = oo
    dtgGy.Col = 0: dtgGy.Text = Ra(0, oo - 1)
    dtgGy.Col = 1: dtgGy.Text = Ra(1, oo - 1)
Next
End Sub

Private Sub txtGy1_Click()
GyId = 1
End Sub

Private Sub txtGy1_DblClick()
On Error Resume Next
Dim Gid As Long

Gid = Val(txtGy1.ToolTipText)
'If Gid = 0 Then Exit Sub
Call frmGyDetail.Qing
Call frmGyDetail.Bound(Gid)
frmGyDetail.cmdSave.Enabled = False
frmGyDetail.Show
frmGyDetail.ZOrder 0
End Sub


Private Sub txtGy2_Click()
GyId = 2
End Sub

Private Sub txtGy2_DblClick()
On Error Resume Next
Dim Gid As Long

Gid = Val(txtGy2.ToolTipText)
'If Gid = 0 Then Exit Sub
Call frmGyDetail.Qing
Call frmGyDetail.Bound(Gid)
frmGyDetail.cmdSave.Enabled = False
frmGyDetail.Show
frmGyDetail.ZOrder 0
End Sub


Private Sub txtGy3_Click()
GyId = 3
End Sub

Private Sub txtGy3_DblClick()
On Error Resume Next
Dim Gid As Long

Gid = Val(txtGY3.ToolTipText)
'If Gid = 0 Then Exit Sub
Call frmGyDetail.Qing
Call frmGyDetail.Bound(Gid)
frmGyDetail.cmdSave.Enabled = False
frmGyDetail.Show
frmGyDetail.ZOrder 0
End Sub


Private Sub txtLx_DblClick()
FmxcLxNew.Show
FmxcLxNew.cmdNew.Visible = False
End Sub

Private Sub txtSL_Change()
txtBje.Text = Val(txtBdj.Text) * Val(txtSL.Text)
txtYH.Text = txtBje.Text
End Sub

Private Sub txtWBdj_Change()
txtWBJe.Text = txtWBdj.Text
End Sub

Private Sub txtWcdj_Change()
If InStr(1, lblZl.Caption, "预估") > 0 Then
    txtWDJ.Text = txtWcdj.Text
Else
    txtWDJ.Text = Round(Val(txtWcdj.Text) * mod1.JZ, 2)
End If
End Sub
Public Sub SDJE(Je As Double) '分摊速达金额
Dim CB As Double
Dim Lhg As Double
Dim oo As Integer
Dim LXG As Double
Dim GY As String
Dim LLG As Double
On Error Resume Next
If Je > 0 Then
For oo = 1 To dtgBr.Rows + 1
    dtgBr.Row = oo
    dtgBr.Col = 3
    If dtgBr.Text = "" Then Exit For
        dtgBr.Col = 7
    If htRow = 1 Or htRow = 2 Or htRow = 3 Or htRow = 4 Or htRow = 5 Or htRow = 6 Or htRow = 7 Or htRow = 7.14 And lblZl.Caption <> "三菱" Or htRow = 8 And lblZl.Caption <> "松下" Or htRow >= 20 Or InStr(1, "维保大修其他人工压缩机维修保养中介业务分包运费吊装费工程人工", lblZl.Caption) > 0 Or _
     (Val(lblBid.ToolTipText) >= 20512 And lblZl.ToolTipText = True) And Not (lblZl = "分包->工程人工" And Val(lblBid.ToolTipText) > 22211 And Val(lblBid.ToolTipText) < 22670) Then
        dtgBr.Col = 3
        If dtgBr.Text = "True" Then
            dtgBr.Col = 1: CB = Val(dtgBr.Text): Lhg = Lhg + CB
            dtgBr.Col = 5: dtgBr.Text = Round(Je * CB / Val(txtJHg.Text), 2): dtgBr.CellForeColor = &H8000&
        End If
    Else
        dtgBr.Col = 9
        If dtgBr.Text = "True" Then
            dtgBr.Col = 6
            CB = Val(dtgBr.Text)
            Lhg = Lhg + CB
            dtgBr.Col = 11: dtgBr.Text = Round(Je * CB / Val(txtJHg.Text), 2): dtgBr.CellForeColor = &H8000&
        End If
    End If
Next

If Lhg <> Val(txtJHg.Text) Then
    dtgBr.Row = dtgBr.Row - 1
    If htRow = 1 Or htRow = 2 Or htRow = 3 Or htRow = 4 Or htRow = 5 Or htRow = 6 Or htRow = 7 And lblZl.Caption <> "三菱" Or htRow = 8 And lblZl.Caption <> "松下" Or htRow >= 20 Or InStr(1, "维保大修其他人工压缩机维修保养中介业务分包运费吊装费工程人工", lblZl.Caption) > 0 Or _
     (Val(lblBid.ToolTipText) >= 20512 And lblZl.ToolTipText = True) Then
        dtgBr.Col = 5: LXG = Val(dtgBr.Text): Lhg = Lhg - LXG
        dtgBr.Text = Je - Lhg
    Else
        dtgBr.Col = 11: LXG = Val(dtgBr.Text): Lhg = Je - LXG
        dtgBr.Text = Je - Lhg
    End If
End If

End If

End Sub

Public Sub CGBound(RC)
Dim Lc As Integer
Dim oo As Integer
On Error Resume Next
txtCBZ.Text = ""
Lc = UBound(RC, 2)
oo = 0
Do While oo <= Lc
    txtCBZ.Text = txtCBZ.Text & RC(2, oo) & " " & RC(0, oo) & " " & RC(1, oo) & Chr(13) & Chr(10)
    oo = oo + 1
Loop
End Sub

Public Sub dtgXZ(ZX As Boolean)
Dim oo As Integer
On Error Resume Next
dtgBr.Visible = False
dtgBr.Row = 0
    For oo = 1 To dtgBr.Rows - 1
        dtgBr.Row = oo
        dtgBr.Col = 1
        If dtgBr.Text = "" Then Exit For
        frmZu.lblDtg.Caption = dtgBr.Text
        If ZX = True Then
            dtgBr.RowHeight(oo) = frmZu.lblDtg.Height
        Else
            dtgBr.RowHeight(oo) = dtgBr.RowHeight(0)
        End If
    Next

dtgBr.Visible = True
End Sub

Public Sub Xian()
Dim oo As Integer
lblBhg.Visible = False
txtBhg.Visible = False
On Error Resume Next
If dtgBr.Cols = 9 Then
    dtgBr.ColWidth(7) = 0
    dtgBr.ColWidth(8) = 0
    For oo = 1 To 100
        dtgBr.Row = oo
        dtgBr.Col = 7: dtgBr.Text = "我早料到了"
        dtgBr.Col = 8: dtgBr.Text = "哈哈"
    Next
    frmWBJ.Visible = False
Else
    dtgBr.ColWidth(22) = 0
    dtgBr.ColWidth(23) = 0
        dtgBr.Row = oo
        dtgBr.Col = 22: dtgBr.Text = "我早料到了"
        dtgBr.Col = 23: dtgBr.Text = "哈哈"
    frmBJ.Visible = False
End If
End Sub

Public Sub dtgDWFF()
dtgDW.Clear
dtgDW.Cols = 3
dtgDW.ColWidth(1) = 0
dtgDW.ColWidth(2) = 0
dtgDW.ColWidth(0) = 3300
End Sub

Public Sub dtgDWBound(tt As String)
Dim Ra
Dim La As Long
Call dtgDWFF

End Sub

Private Sub txtXmmc_DblClick()
Dim tt As String
Dim oo As Integer
Dim Ra
Dim La As Integer
If Not (Len(txtXmmc.Text) >= 2) Then Exit Sub
tt = "select xmmc,xid,xmadr from xmzl where xmmc like '%" & txtXmmc.Text & "%'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
Call Me.dtgDWFF
dtgDW.Rows = La + 10
For oo = 0 To La - 1
    dtgDW.Row = oo
    dtgDW.Col = 0: dtgDW.Text = Ra(0, oo)
    dtgDW.Col = 1: dtgDW.Text = Ra(1, oo)
    dtgDW.Col = 2: dtgDW.Text = Ra(2, oo)
    
Next
dtgDW.Visible = True
dtgDW.Top = txtXmmc.Top + txtXmmc.Height
End Sub


