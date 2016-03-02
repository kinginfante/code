VERSION 5.00
Begin VB.Form frmWBXT2 
   Caption         =   "小机、末端、空调箱保养定额价格表"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11370
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9120
   ScaleWidth      =   11370
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   2010
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   0
   End
   Begin VB.TextBox txtCJR 
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
      Height          =   360
      Left            =   1140
      TabIndex        =   100
      Top             =   8580
      Width           =   2445
   End
   Begin VB.Frame frmJT 
      BackColor       =   &H00C0FFFF&
      Height          =   2475
      Left            =   7380
      TabIndex        =   86
      Top             =   6000
      Width           =   3975
      Begin VB.TextBox J21 
         Height          =   270
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   99
         Top             =   2100
         Width           =   1305
      End
      Begin VB.TextBox J20 
         Height          =   270
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   98
         Top             =   1776
         Width           =   1305
      End
      Begin VB.TextBox J19 
         Height          =   270
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   97
         Top             =   1452
         Width           =   1305
      End
      Begin VB.TextBox J18 
         Height          =   270
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   96
         Top             =   1128
         Width           =   1305
      End
      Begin VB.TextBox J17 
         Height          =   270
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   95
         Top             =   804
         Width           =   1305
      End
      Begin VB.TextBox J16 
         Height          =   270
         Left            =   1830
         TabIndex        =   94
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label Label25 
         Caption         =   $"frmWBXT2.frx":0000
         Height          =   195
         Left            =   2130
         TabIndex        =   93
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWBXT2.frx":000C
         Height          =   165
         Left            =   270
         TabIndex        =   92
         Top             =   2130
         Width           =   1095
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "急修"
         Height          =   165
         Left            =   270
         TabIndex        =   91
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "巡视"
         Height          =   165
         Left            =   270
         TabIndex        =   90
         Top             =   1485
         Width           =   1095
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "年保"
         Height          =   165
         Left            =   270
         TabIndex        =   89
         Top             =   1155
         Width           =   1095
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "单次金额"
         Height          =   165
         Left            =   270
         TabIndex        =   88
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "与内环距离"
         Height          =   165
         Left            =   270
         TabIndex        =   87
         Top             =   510
         Width           =   1095
      End
   End
   Begin VB.ComboBox C24 
      Height          =   300
      ItemData        =   "frmWBXT2.frx":0016
      Left            =   2340
      List            =   "frmWBXT2.frx":0020
      TabIndex        =   85
      Text            =   "Combo1"
      Top             =   7740
      Width           =   1275
   End
   Begin VB.TextBox txtBz 
      Height          =   5385
      Left            =   7380
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   82
      Text            =   "frmWBXT2.frx":0032
      Top             =   570
      Width           =   4035
   End
   Begin VB.ComboBox C23 
      Height          =   300
      ItemData        =   "frmWBXT2.frx":0038
      Left            =   2340
      List            =   "frmWBXT2.frx":0042
      Style           =   2  'Dropdown List
      TabIndex        =   79
      Top             =   7380
      Width           =   1275
   End
   Begin VB.ComboBox C22 
      Height          =   300
      ItemData        =   "frmWBXT2.frx":0054
      Left            =   2340
      List            =   "frmWBXT2.frx":005E
      Style           =   2  'Dropdown List
      TabIndex        =   78
      Top             =   7050
      Width           =   1275
   End
   Begin VB.ComboBox C21 
      Height          =   300
      ItemData        =   "frmWBXT2.frx":0070
      Left            =   2340
      List            =   "frmWBXT2.frx":009B
      Style           =   2  'Dropdown List
      TabIndex        =   77
      Top             =   6690
      Width           =   1275
   End
   Begin VB.CommandButton cmdJi 
      Caption         =   "计算"
      Height          =   375
      Left            =   9300
      TabIndex        =   76
      Top             =   8700
      Width           =   495
   End
   Begin VB.CommandButton cmdBack 
      Height          =   375
      Left            =   10830
      Picture         =   "frmWBXT2.frx":00C9
      Style           =   1  'Graphical
      TabIndex        =   75
      ToolTipText     =   "返回"
      Top             =   8700
      Width           =   495
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   10290
      Picture         =   "frmWBXT2.frx":01CB
      Style           =   1  'Graphical
      TabIndex        =   74
      ToolTipText     =   "保存"
      Top             =   8700
      Width           =   495
   End
   Begin VB.CommandButton cmdMod 
      Height          =   375
      Left            =   9795
      Picture         =   "frmWBXT2.frx":0835
      Style           =   1  'Graphical
      TabIndex        =   73
      ToolTipText     =   "修改"
      Top             =   8700
      Width           =   495
   End
   Begin VB.ComboBox C19 
      Height          =   300
      ItemData        =   "frmWBXT2.frx":0B3F
      Left            =   2370
      List            =   "frmWBXT2.frx":0B49
      Style           =   2  'Dropdown List
      TabIndex        =   72
      Top             =   5760
      Width           =   1245
   End
   Begin VB.ComboBox C18 
      Height          =   300
      ItemData        =   "frmWBXT2.frx":0B53
      Left            =   2370
      List            =   "frmWBXT2.frx":0B75
      Style           =   2  'Dropdown List
      TabIndex        =   71
      Top             =   5415
      Width           =   1245
   End
   Begin VB.TextBox C17 
      Height          =   270
      Left            =   2370
      TabIndex        =   70
      Top             =   5100
      Width           =   1215
   End
   Begin VB.ComboBox C16 
      Height          =   300
      ItemData        =   "frmWBXT2.frx":0B9E
      Left            =   2370
      List            =   "frmWBXT2.frx":0BA8
      Style           =   2  'Dropdown List
      TabIndex        =   69
      Top             =   4770
      Width           =   1245
   End
   Begin VB.ComboBox C15 
      Height          =   300
      ItemData        =   "frmWBXT2.frx":0BB2
      Left            =   2370
      List            =   "frmWBXT2.frx":0BD4
      Style           =   2  'Dropdown List
      TabIndex        =   68
      Top             =   4425
      Width           =   1245
   End
   Begin VB.TextBox C14 
      Height          =   270
      Left            =   2370
      TabIndex        =   67
      Top             =   4110
      Width           =   1215
   End
   Begin VB.ComboBox C13 
      Height          =   300
      ItemData        =   "frmWBXT2.frx":0BFD
      Left            =   2340
      List            =   "frmWBXT2.frx":0C07
      Style           =   2  'Dropdown List
      TabIndex        =   66
      Top             =   3780
      Width           =   1245
   End
   Begin VB.ComboBox C12 
      Height          =   300
      ItemData        =   "frmWBXT2.frx":0C11
      Left            =   2340
      List            =   "frmWBXT2.frx":0C33
      Style           =   2  'Dropdown List
      TabIndex        =   65
      Top             =   3435
      Width           =   1245
   End
   Begin VB.TextBox C11 
      Height          =   270
      Left            =   2340
      TabIndex        =   64
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ComboBox C10 
      Height          =   300
      ItemData        =   "frmWBXT2.frx":0C5C
      Left            =   2340
      List            =   "frmWBXT2.frx":0C66
      Style           =   2  'Dropdown List
      TabIndex        =   63
      Top             =   2790
      Width           =   1245
   End
   Begin VB.ComboBox C9 
      Height          =   300
      ItemData        =   "frmWBXT2.frx":0C70
      Left            =   2340
      List            =   "frmWBXT2.frx":0C92
      Style           =   2  'Dropdown List
      TabIndex        =   62
      Top             =   2445
      Width           =   1245
   End
   Begin VB.TextBox C8 
      Height          =   270
      Left            =   2340
      TabIndex        =   61
      Top             =   2130
      Width           =   1215
   End
   Begin VB.ComboBox C7 
      Height          =   300
      ItemData        =   "frmWBXT2.frx":0CBB
      Left            =   2340
      List            =   "frmWBXT2.frx":0CC5
      Style           =   2  'Dropdown List
      TabIndex        =   60
      Top             =   1770
      Width           =   1245
   End
   Begin VB.ComboBox C6 
      Height          =   300
      ItemData        =   "frmWBXT2.frx":0CCF
      Left            =   2340
      List            =   "frmWBXT2.frx":0CF1
      Style           =   2  'Dropdown List
      TabIndex        =   59
      Top             =   1425
      Width           =   1245
   End
   Begin VB.TextBox C5 
      Height          =   270
      Left            =   2340
      TabIndex        =   58
      Top             =   1110
      Width           =   1215
   End
   Begin VB.ComboBox C4 
      Height          =   300
      ItemData        =   "frmWBXT2.frx":0D1A
      Left            =   2340
      List            =   "frmWBXT2.frx":0D24
      Style           =   2  'Dropdown List
      TabIndex        =   57
      Top             =   780
      Width           =   1245
   End
   Begin VB.ComboBox C3 
      Height          =   300
      ItemData        =   "frmWBXT2.frx":0D2E
      Left            =   2340
      List            =   "frmWBXT2.frx":0D50
      Style           =   2  'Dropdown List
      TabIndex        =   56
      Top             =   435
      Width           =   1245
   End
   Begin VB.TextBox D24 
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
      Left            =   4050
      Locked          =   -1  'True
      TabIndex        =   55
      Top             =   7710
      Width           =   2775
   End
   Begin VB.TextBox D23 
      Height          =   270
      Left            =   4050
      Locked          =   -1  'True
      TabIndex        =   54
      Top             =   7380
      Width           =   945
   End
   Begin VB.TextBox D22 
      Height          =   270
      Left            =   4050
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   7050
      Width           =   945
   End
   Begin VB.TextBox D21 
      Height          =   270
      Left            =   4050
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   6690
      Width           =   945
   End
   Begin VB.TextBox D20 
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
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   6150
      Width           =   2775
   End
   Begin VB.TextBox F19 
      Height          =   270
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   5760
      Width           =   855
   End
   Begin VB.TextBox E19 
      Height          =   270
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   5760
      Width           =   855
   End
   Begin VB.TextBox D19 
      Height          =   270
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   5760
      Width           =   855
   End
   Begin VB.TextBox F16 
      Height          =   270
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   4770
      Width           =   855
   End
   Begin VB.TextBox E16 
      Height          =   270
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   4770
      Width           =   855
   End
   Begin VB.TextBox D16 
      Height          =   270
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   4770
      Width           =   855
   End
   Begin VB.TextBox F13 
      Height          =   270
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   3780
      Width           =   855
   End
   Begin VB.TextBox E13 
      Height          =   270
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   3780
      Width           =   855
   End
   Begin VB.TextBox D13 
      Height          =   270
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   3780
      Width           =   855
   End
   Begin VB.TextBox F10 
      Height          =   270
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox E10 
      Height          =   270
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox D10 
      Height          =   270
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox F7 
      Height          =   270
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox E7 
      Height          =   270
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox D7 
      Height          =   270
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox F4 
      Height          =   270
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   780
      Width           =   855
   End
   Begin VB.TextBox E4 
      Height          =   270
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   780
      Width           =   855
   End
   Begin VB.TextBox D4 
      Height          =   270
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   780
      Width           =   855
   End
   Begin VB.TextBox C2 
      Height          =   270
      Left            =   2340
      TabIndex        =   24
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblMid 
      Caption         =   "lblMid"
      Height          =   315
      Left            =   90
      TabIndex        =   83
      Top             =   330
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label lblBid 
      Caption         =   "lblBid"
      Height          =   195
      Left            =   0
      TabIndex        =   101
      Top             =   0
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Label16 
      Caption         =   "全包"
      Height          =   225
      Left            =   1200
      TabIndex        =   84
      Top             =   7800
      Width           =   945
   End
   Begin VB.Label Label18 
      Caption         =   "备注"
      Height          =   285
      Left            =   7470
      TabIndex        =   81
      Top             =   150
      Width           =   1725
   End
   Begin VB.Label Label17 
      Caption         =   "承接人"
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
      Left            =   240
      TabIndex        =   80
      Top             =   8670
      Width           =   735
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   0
      X2              =   7365
      Y1              =   6090
      Y2              =   6090
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   0
      X2              =   7365
      Y1              =   5070
      Y2              =   5070
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   0
      X2              =   7365
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   0
      X2              =   7365
      Y1              =   3090
      Y2              =   3090
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   0
      X2              =   7365
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   0
      X2              =   7365
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label15 
      Caption         =   "材料全包"
      Height          =   255
      Left            =   1140
      TabIndex        =   51
      Top             =   7410
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "含大修"
      Height          =   225
      Left            =   1140
      TabIndex        =   50
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "巡视次数"
      Height          =   285
      Left            =   1140
      TabIndex        =   49
      Top             =   6720
      Width           =   915
   End
   Begin VB.Label Label12 
      Caption         =   "基本保养"
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
      Left            =   1140
      TabIndex        =   47
      Top             =   6180
      Width           =   1125
   End
   Begin VB.Label Label11 
      Caption         =   "材料"
      Height          =   225
      Left            =   6060
      TabIndex        =   28
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label10 
      Caption         =   "大修"
      Height          =   225
      Left            =   5070
      TabIndex        =   27
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "维保"
      Height          =   225
      Left            =   4080
      TabIndex        =   26
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "金额"
      Height          =   225
      Left            =   4110
      TabIndex        =   25
      Top             =   150
      Width           =   2475
   End
   Begin VB.Label Label7 
      Caption         =   "类型6"
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
      Index           =   5
      Left            =   210
      TabIndex        =   23
      Top             =   5280
      Width           =   645
   End
   Begin VB.Label Label7 
      Caption         =   "类型5"
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
      Index           =   4
      Left            =   210
      TabIndex        =   22
      Top             =   4290
      Width           =   645
   End
   Begin VB.Label Label7 
      Caption         =   "类型4"
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
      Index           =   3
      Left            =   210
      TabIndex        =   21
      Top             =   3312
      Width           =   645
   End
   Begin VB.Label Label7 
      Caption         =   "类型3"
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
      Index           =   2
      Left            =   210
      TabIndex        =   20
      Top             =   2358
      Width           =   645
   End
   Begin VB.Label Label7 
      Caption         =   "类型2"
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
      Index           =   1
      Left            =   210
      TabIndex        =   19
      Top             =   1404
      Width           =   645
   End
   Begin VB.Label Label7 
      Caption         =   "类型1"
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
      Index           =   0
      Left            =   210
      TabIndex        =   18
      Top             =   450
      Width           =   645
   End
   Begin VB.Label Label3 
      Caption         =   "保养次数"
      Height          =   195
      Index           =   4
      Left            =   1170
      TabIndex        =   17
      Top             =   5790
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "大小"
      Height          =   195
      Index           =   4
      Left            =   1170
      TabIndex        =   16
      Top             =   5475
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "数量"
      Height          =   195
      Index           =   4
      Left            =   1170
      TabIndex        =   15
      Top             =   5160
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "保养次数"
      Height          =   195
      Index           =   3
      Left            =   1170
      TabIndex        =   14
      Top             =   4830
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "大小"
      Height          =   195
      Index           =   3
      Left            =   1170
      TabIndex        =   13
      Top             =   4515
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "数量"
      Height          =   195
      Index           =   3
      Left            =   1170
      TabIndex        =   12
      Top             =   4200
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "保养次数"
      Height          =   195
      Index           =   2
      Left            =   1170
      TabIndex        =   11
      Top             =   3840
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "大小"
      Height          =   195
      Index           =   2
      Left            =   1170
      TabIndex        =   10
      Top             =   3525
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "数量"
      Height          =   195
      Index           =   2
      Left            =   1170
      TabIndex        =   9
      Top             =   3210
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "保养次数"
      Height          =   195
      Index           =   1
      Left            =   1170
      TabIndex        =   8
      Top             =   2850
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "大小"
      Height          =   195
      Index           =   1
      Left            =   1170
      TabIndex        =   7
      Top             =   2535
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "数量"
      Height          =   195
      Index           =   1
      Left            =   1170
      TabIndex        =   6
      Top             =   2220
      Width           =   1035
   End
   Begin VB.Label Label6 
      Caption         =   "保养次数"
      Height          =   195
      Left            =   1170
      TabIndex        =   5
      Top             =   1830
      Width           =   1035
   End
   Begin VB.Label Label5 
      Caption         =   "大小"
      Height          =   195
      Left            =   1170
      TabIndex        =   4
      Top             =   1515
      Width           =   1035
   End
   Begin VB.Label Label4 
      Caption         =   "数量"
      Height          =   195
      Left            =   1170
      TabIndex        =   3
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "保养次数"
      Height          =   195
      Index           =   0
      Left            =   1170
      TabIndex        =   2
      Top             =   810
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "大小"
      Height          =   195
      Index           =   0
      Left            =   1170
      TabIndex        =   1
      Top             =   495
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "数量"
      Height          =   195
      Index           =   0
      Left            =   1170
      TabIndex        =   0
      Top             =   180
      Width           =   1035
   End
End
Attribute VB_Name = "frmWBXT2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim timZm As Integer '(1保存 )

Private Sub cmdBack_Click()
Me.Visible = False
End Sub

Private Sub cmdJi_Click()

If Me.Visible = False Then Exit Sub
If Val(C2.Text) = 0 And Val(C5.Text) = 0 And Val(C8.Text) = 0 And Val(C11.Text) = 0 And Val(C14.Text) = 0 And Val(C17.Text) = 0 Then
    MsgBox "请确定机组数量!"
    Exit Sub
End If
Call J1


End Sub

Public Sub Qing()
C2.Text = ""
C3.Text = 1
C4.Text = 1
C5.Text = ""
C6.Text = 1
C7.Text = 1
C8.Text = ""
C9.Text = 1
C10.Text = 1
C11.Text = ""
C12.Text = 1
C13.Text = 1
C14.Text = ""
C15.Text = 1
C16.Text = 1
C17.Text = ""
C18.Text = 1
C19.Text = 1
C21.Text = 0
C22.Text = "不包含"
C23.Text = "不包含"
D4.Text = ""
E4.Text = ""
F4.Text = ""
D7.Text = ""
E7.Text = ""
F7.Text = ""
D10.Text = ""
E10.Text = ""
F10.Text = ""
D13.Text = ""
E13.Text = ""
F13.Text = ""
D16.Text = ""
E16.Text = ""
F16.Text = ""
D19.Text = ""
E19.Text = ""
F19.Text = ""
D20.Text = ""
D21.Text = ""
D22.Text = ""
D23.Text = ""
D24.Text = ""
C24 = "非全包"
txtBz.Text = ""
lblBid.Caption = ""
lblMid.Caption = ""
End Sub

Private Sub cmdMod_Click()
If Val(frmWBXX.lblLc.Caption) > 1 And mod1.DName <> "" Then
    Exit Sub

End If
cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
On Error Resume Next

Call cmdJi_Click

 '保存
    timZm = 1
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "人工业务"
    mod1.cmd.Parameters("@NBLX") = "保存"
    mod1.cmd.Parameters("@bh") = lblMid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = ""
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = txtCJR.Text '承接人
    mod1.cmd.Parameters("@mt6") = C22
    mod1.cmd.Parameters("@mt7") = C23
    mod1.cmd.Parameters("@mt8") = C24
    mod1.cmd.Parameters("@mlt1") = txtBz.Text '备注

    mod1.cmd.Parameters("@mm1") = Val(D24.Text) '基准价
    mod1.cmd.Parameters("@mm2") = Val(J21.Text) '交通费
    mod1.cmd.Parameters("@mm3") = 0
    mod1.cmd.Parameters("@mm4") = 0
    mod1.cmd.Parameters("@mm5") = Val(lblBid.Caption) '询价单号
    mod1.cmd.Parameters("@mm6") = Val(C2.Text)
    mod1.cmd.Parameters("@mm7") = Val(C3.Text)
    mod1.cmd.Parameters("@mm8") = Val(C4.Text)
    mod1.cmd.Parameters("@mm9") = Val(C5.Text)
    mod1.cmd.Parameters("@mm10") = Val(C6.Text)
    mod1.cmd.Parameters("@mm11") = Val(C7.Text)
    mod1.cmd.Parameters("@mm12") = Val(C8.Text)
    mod1.cmd.Parameters("@mm13") = Val(C9.Text)
    mod1.cmd.Parameters("@mm14") = Val(C10.Text)
    mod1.cmd.Parameters("@mm15") = Val(C11.Text)
    mod1.cmd.Parameters("@mm16") = Val(C12.Text)
    mod1.cmd.Parameters("@mm17") = Val(C13.Text)
    mod1.cmd.Parameters("@mm18") = Val(C14.Text)
    mod1.cmd.Parameters("@mm19") = Val(C15.Text)
    mod1.cmd.Parameters("@mm20") = Val(C16.Text)
    mod1.cmd.Parameters("@mm21") = Val(C17.Text)
    mod1.cmd.Parameters("@mm22") = Val(C18.Text)
    mod1.cmd.Parameters("@mm23") = Val(C19.Text)
    mod1.cmd.Parameters("@mm24") = Val(C21.Text)
    mod1.cmd.Parameters("@mm25") = Val(J16.Text)
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

Private Sub Form_Load()
Me.Width = 11490
Me.Height = 9630
Me.Left = 0
Me.Top = 0
End Sub



Public Sub J1()
Dim tt As String
Dim Ra: Dim La
Dim oo As Integer
On Error GoTo frmWBXT2ERR
If Val(C2.Text) > 0 Then
If C4 = 1 Then
    tt = "select B1 from Z3 where jdx=" & Val(C3.Text)
Else
    tt = "select B2 from Z3 where jdx=" & Val(C3.Text)
End If
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
D4 = Ra(0, 0) * Val(C2.Text)

tt = "select dxrg from Z3 where jdx=" & Val(C3.Text)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
E4.Text = Ra(0, 0) * Val(C2.Text)

tt = "select clqb from Z3 where jdx=" & Val(C3.Text)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
F4.Text = Ra(0, 0) * Val(C2.Text)
End If

If Val(C5.Text) > 0 Then
If C7 = 1 Then
    tt = "select B1 from Z3 where jdx=" & Val(C6.Text)
Else
    tt = "select B2 from Z3 where jdx=" & Val(C6.Text)
End If
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
D7 = Ra(0, 0) * Val(C5.Text)

tt = "select dxrg from Z3 where jdx=" & Val(C6.Text)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
E7.Text = Ra(0, 0) * Val(C5.Text)

tt = "select clqb from Z3 where jdx=" & Val(C6.Text)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
F7.Text = Ra(0, 0) * Val(C5.Text)
End If

If Val(C8.Text) > 0 Then
If C10 = 1 Then
    tt = "select B1 from Z3 where jdx=" & Val(C9.Text)
Else
    tt = "select B2 from Z3 where jdx=" & Val(C9.Text)
End If
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
D10 = Ra(0, 0) * Val(C8.Text)

tt = "select dxrg from Z3 where jdx=" & Val(C9.Text)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
E10.Text = Ra(0, 0) * Val(C8.Text)

tt = "select clqb from Z3 where jdx=" & Val(C9.Text)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
F10.Text = Ra(0, 0) * Val(C8.Text)
End If

If Val(C11.Text) > 0 Then
If C13 = 1 Then
    tt = "select B1 from Z3 where jdx=" & Val(C12.Text)
Else
    tt = "select B2 from Z3 where jdx=" & Val(C12.Text)
End If
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
D13 = Ra(0, 0) * Val(C11.Text)

tt = "select dxrg from Z3 where jdx=" & Val(C12.Text)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
E13.Text = Ra(0, 0) * Val(C11.Text)

tt = "select clqb from Z3 where jdx=" & Val(C12.Text)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
F13.Text = Ra(0, 0) * Val(C11.Text)
End If

If Val(C14.Text) > 0 Then
If C16 = 1 Then
    tt = "select B1 from Z3 where jdx=" & Val(C15.Text)
Else
    tt = "select B2 from Z3 where jdx=" & Val(C15.Text)
End If
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
D16 = Ra(0, 0) * Val(C14.Text)

tt = "select dxrg from Z3 where jdx=" & Val(C15.Text)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
E16.Text = Ra(0, 0) * Val(C14.Text)

tt = "select clqb from Z3 where jdx=" & Val(C15.Text)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
F16.Text = Ra(0, 0) * Val(C14.Text)
End If

If Val(C17.Text) > 0 Then
If C19 = 1 Then
    tt = "select B1 from Z3 where jdx=" & Val(C18.Text)
Else
    tt = "select B2 from Z3 where jdx=" & Val(C18.Text)
End If
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
D19 = Ra(0, 0) * Val(C17.Text)

tt = "select dxrg from Z3 where jdx=" & Val(C18.Text)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
E19.Text = Ra(0, 0) * Val(C17.Text)

tt = "select clqb from Z3 where jdx=" & Val(C18.Text)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
F19.Text = Ra(0, 0) * Val(C17.Text)
End If
'd20=SUM(D2:D19)*(1-(C2+C5+C8+C11+C14+C17-50)/(C2+C5+C8+C11+C14+C17+100)/2)*1.2
D20 = Round((Val(D4.Text) + Val(D7.Text) + Val(D10.Text) + Val(D13.Text) + Val(D16.Text) + Val(D19.Text)) * _
        (1 - (Val(C2.Text) + Val(C5.Text) + Val(C8.Text) + Val(C11.Text) + Val(C14.Text) + Val(C17.Text) - 50) / _
        (Val(C2.Text) + Val(C5.Text) + Val(C8.Text) + Val(C11.Text) + Val(C14.Text) + Val(C17.Text) + 100) / 2) * 1.2, 0)
'd21=IF(C21=0,0,250*(C21+SUM(C2,C5,C8,C11,C14,C17)/30))
If C21.Text = 0 Then
    D21 = 0
Else
    D21 = Round(250 * (Val(C21.Text) + (Val(C2.Text) + Val(C5.Text) + Val(C8.Text) + Val(C11.Text) + Val(C14.Text) + Val(C17.Text)) / 30), 0)
End If
'd22=IF(C22="包含",SUM(E4:E19),0)
If C22.Text = "包含" Then
    D22 = Round(Val(E4.Text) + Val(E7.Text) + Val(E10.Text) + Val(E13.Text) + Val(E16.Text) + Val(E19.Text), 0)
Else
    D22 = 0
End If
'd23=IF(C23="包含",SUM(F4:F19),0)
If C23.Text = "包含" Then
    D23 = Round(Val(F4.Text) + Val(F7.Text) + Val(F10.Text) + Val(F13.Text) + Val(F16.Text) + Val(F19.Text), 0)
Else
    D23 = 0
End If
If C24.Text = "全包" Then
    D24 = Val(D20.Text) + Val(D21.Text) + Val(D22.Text) + Val(D23.Text)
Else
    D24 = Val(D20) + Val(D21)
End If

'交通
J17 = Round(30 + 2 * Val(J16), 0)
J18 = Round(Val(J17) * (1 + (Val(C2) + Val(C5) + Val(C8) + Val(C11) + Val(C14) + Val(C17) - 1) / 30), 0)
If Val(C21) = 0 Then
    J19 = 0
Else
    J19 = Round(Val(J17) * (Val(C21) + ((Val(C2) + Val(C5) + Val(C8) + Val(C11) + Val(C14) + Val(C17) - 1) / 30)), 0)
End If
J20 = Round(Val(J17) * (1 + (Val(C2) + Val(C5) + Val(C8) + Val(C11) + Val(C14) + Val(C17) - 1) / 20), 0)
J21 = Round(Val(J18) + Val(J19) + Val(J20), 2)
'j21=SUM(J18:J20)
Exit Sub
frmWBXT2ERR:
MsgBox "出错!"
End Sub

Private Sub timQuit_Timer()
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0


If timZm = 1 Then    '保存
    cmdSave.Enabled = False
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
    If timZm = 1 Then '保存
        Call frmWBXX.MXBound(Val(frmWBXX.lblBid.Caption))
        frmWBXX.txt2.Text = mod1.WP.Fields("mm2").Value
        'Call frmWBXX.ji
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



Public Sub Bound(Mid As Long)
Dim JZ As Single
Dim JT As Single
Dim tt As String
Dim Ra
On Error GoTo frmWBXT2
tt = "select mt1,mt2,mt3,mt4,mt5,mt6,mt7,mt8,mt9,mt10,mt11,mt12,mt13,mt14,mt15,mt16,mt17,mt18,mt19,mt20," & _
    "mt21,mt22,mt23,mt24,mt25,mt26,mt27,mt28,mt29,mt30,mt31,mt32,mt33,mt34,mt35,mt36,mt37,mt38,mt39,mt40,mt41,mt42,mlt1,mlt2,mlt3,mlt4,mlt5," & _
    "mm1,mm2,mm3,mm4,mm5,mm6,mm7,mm8,mm9,mm10,mm11,mm12,mm13,mm14,mm15,mm16,mm17,mm18,mm19,mm20," & _
    "mm21,mm22,mm23,mm24,mm25,mm26,mm27,mm28,mm29,mm30,mb1,mb2,mb3,mb4,mb5,bid,mid" & _
    " from MlMX where mid=" & Mid
'19
'46
'66
'83
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
On Error Resume Next
txtCJR.Text = Ra(4, 0) '承接人
C22.Text = Ra(5, 0)
C23.Text = Ra(6, 0)
C24.Text = Ra(7, 0)
txtBz.Text = Ra(42, 0) '备注
D24.Text = Ra(47, 0): JZ = Val(D24.Text) '基准价
J21.Text = Ra(48, 0): JT = Val(J21.Text) '差旅
C2.Text = Ra(52, 0)
C3.Text = Ra(53, 0)
C4.Text = Ra(54, 0)
C5.Text = Ra(55, 0)
C6.Text = Ra(56, 0)
C7.Text = Ra(57, 0)
C8.Text = Ra(58, 0)
C9.Text = Ra(59, 0)
C10.Text = Ra(60, 0)
C11.Text = Ra(61, 0)
C12.Text = Ra(62, 0)
C13.Text = Ra(63, 0)
C14.Text = Ra(64, 0)
C15.Text = Ra(65, 0)
C16.Text = Ra(66, 0)
C17.Text = Ra(67, 0)
C18.Text = Ra(68, 0)
C19.Text = Ra(69, 0)
C21.Text = Ra(70, 0)
J16.Text = Ra(71, 0)

  

 lblMid.Caption = Ra(83, 0)
 lblBid.Caption = Ra(82, 0)
Call J1
Exit Sub
frmWBXT2:
MsgBox "出错!"
End


End Sub
