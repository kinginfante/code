VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form b2 
   Caption         =   "上海豪曼制冷空调服务有限公司"
   ClientHeight    =   9150
   ClientLeft      =   5865
   ClientTop       =   2955
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   15210
   Begin VB.CommandButton cmdMQm 
      Height          =   345
      Index           =   4
      Left            =   11580
      TabIndex        =   117
      Top             =   8280
      Width           =   1245
   End
   Begin VB.CommandButton cmdMQm 
      Height          =   345
      Index           =   3
      Left            =   10290
      TabIndex        =   114
      Top             =   8280
      Width           =   1245
   End
   Begin VB.CommandButton cmdMQm 
      Height          =   345
      Index           =   2
      Left            =   8970
      TabIndex        =   111
      Top             =   8280
      Width           =   1245
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "修改"
      Height          =   585
      Left            =   13260
      Picture         =   "b2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   8550
      Width           =   645
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   585
      Left            =   14610
      Picture         =   "b2.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   8550
      Width           =   585
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "提交"
      Enabled         =   0   'False
      Height          =   585
      Left            =   13920
      Picture         =   "b2.frx":040C
      Style           =   1  'Graphical
      TabIndex        =   107
      Top             =   8550
      Width           =   675
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12300
      Top             =   5700
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11520
      Top             =   5700
   End
   Begin VB.Frame frmQm 
      BackColor       =   &H00C0FFC0&
      Caption         =   "评审建议"
      ForeColor       =   &H000000FF&
      Height          =   1785
      Left            =   -60
      TabIndex        =   100
      Top             =   7380
      Visible         =   0   'False
      Width           =   6315
      Begin VB.TextBox txtQM 
         BackColor       =   &H00C0FFFF&
         Height          =   1365
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   104
         Top             =   300
         Width           =   4965
      End
      Begin VB.OptionButton OptT1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "同意"
         Height          =   225
         Left            =   5220
         TabIndex        =   103
         Top             =   480
         Width           =   705
      End
      Begin VB.OptionButton optT2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "拒绝"
         Height          =   195
         Left            =   5220
         TabIndex        =   102
         Top             =   870
         Width           =   675
      End
      Begin VB.CommandButton cmdDing 
         BackColor       =   &H00FF8080&
         Caption         =   "决定"
         Height          =   285
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   1365
      Left            =   870
      TabIndex        =   96
      Top             =   -660
      Visible         =   0   'False
      Width           =   2505
      Begin VB.Label lblLc 
         Caption         =   "lblLc"
         Height          =   315
         Left            =   0
         TabIndex        =   99
         Top             =   390
         Width           =   645
      End
      Begin VB.Label lblLcUid 
         Caption         =   "lblLcUid"
         Height          =   285
         Left            =   30
         TabIndex        =   98
         Top             =   810
         Width           =   885
      End
      Begin VB.Label lblFwid 
         Caption         =   "lblFwid"
         Height          =   255
         Left            =   1200
         TabIndex        =   97
         Top             =   0
         Width           =   885
      End
   End
   Begin VB.TextBox txtBmp 
      Height          =   585
      Left            =   1710
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   95
      Text            =   "b2.frx":0A76
      Top             =   8490
      Width           =   4395
   End
   Begin VB.TextBox txtZjp 
      Height          =   615
      Left            =   1710
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   94
      Text            =   "b2.frx":0A7C
      Top             =   7800
      Width           =   4395
   End
   Begin VB.CommandButton cmdZuan 
      Caption         =   "->"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14820
      TabIndex        =   93
      Top             =   0
      Width           =   435
   End
   Begin VB.CommandButton cmdMQm 
      Height          =   345
      Index           =   0
      Left            =   6240
      TabIndex        =   88
      Top             =   8280
      Width           =   1245
   End
   Begin VB.CommandButton cmdMQm 
      Height          =   345
      Index           =   1
      Left            =   7590
      TabIndex        =   87
      Top             =   8280
      Width           =   1305
   End
   Begin VB.TextBox txtJ5 
      Height          =   270
      Left            =   14070
      TabIndex        =   84
      Text            =   "Text55"
      Top             =   6810
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txtI5 
      Height          =   270
      Left            =   14070
      TabIndex        =   83
      Text            =   "Text54"
      Top             =   6420
      Width           =   795
   End
   Begin VB.TextBox txtJ4 
      Height          =   270
      Left            =   12420
      TabIndex        =   82
      Text            =   "Text53"
      Top             =   6810
      Width           =   795
   End
   Begin VB.TextBox txtI4 
      Height          =   270
      Left            =   12420
      TabIndex        =   81
      Text            =   "Text52"
      Top             =   6420
      Width           =   795
   End
   Begin VB.TextBox txtJ3 
      Height          =   270
      Left            =   11310
      TabIndex        =   80
      Text            =   "Text51"
      Top             =   6810
      Width           =   795
   End
   Begin VB.TextBox txtI3 
      Height          =   270
      Left            =   11310
      TabIndex        =   79
      Text            =   "Text50"
      Top             =   6420
      Width           =   795
   End
   Begin VB.TextBox txtJ2 
      Height          =   270
      Left            =   9900
      TabIndex        =   78
      Text            =   "Text49"
      Top             =   6810
      Width           =   795
   End
   Begin VB.TextBox txtI2 
      Height          =   270
      Left            =   9900
      TabIndex        =   77
      Text            =   "Text48"
      Top             =   6420
      Width           =   795
   End
   Begin VB.TextBox txtJ1 
      Height          =   270
      Left            =   2850
      TabIndex        =   76
      Text            =   "Text47"
      Top             =   6810
      Width           =   6645
   End
   Begin VB.TextBox txtI1 
      Height          =   285
      Left            =   2850
      TabIndex        =   75
      Text            =   "Text46"
      Top             =   6450
      Width           =   6675
   End
   Begin VB.TextBox txtH5 
      Height          =   270
      Left            =   14070
      TabIndex        =   69
      Text            =   "Text26"
      Top             =   6120
      Width           =   795
   End
   Begin VB.TextBox txtG5 
      Height          =   270
      Left            =   14070
      TabIndex        =   68
      Text            =   "Text25"
      Top             =   5730
      Width           =   795
   End
   Begin VB.TextBox txtF5 
      Height          =   270
      Left            =   14070
      TabIndex        =   67
      Text            =   "Text24"
      Top             =   5415
      Width           =   795
   End
   Begin VB.TextBox txtE5 
      Height          =   270
      Left            =   14070
      TabIndex        =   66
      Text            =   "Text23"
      Top             =   5085
      Width           =   795
   End
   Begin VB.TextBox txtD5 
      Height          =   270
      Left            =   14070
      TabIndex        =   65
      Text            =   "Text22"
      Top             =   4740
      Width           =   795
   End
   Begin VB.TextBox txtC5 
      Height          =   270
      Left            =   14070
      TabIndex        =   64
      Text            =   "Text21"
      Top             =   4410
      Width           =   795
   End
   Begin VB.TextBox txtH4 
      Height          =   270
      Left            =   12420
      TabIndex        =   63
      Text            =   "Text26"
      Top             =   6120
      Width           =   795
   End
   Begin VB.TextBox txtG4 
      Height          =   270
      Left            =   12420
      TabIndex        =   62
      Text            =   "Text25"
      Top             =   5730
      Width           =   795
   End
   Begin VB.TextBox txtF4 
      Height          =   270
      Left            =   12420
      TabIndex        =   61
      Text            =   "Text24"
      Top             =   5415
      Width           =   795
   End
   Begin VB.TextBox txtE4 
      Height          =   270
      Left            =   12420
      TabIndex        =   60
      Text            =   "Text23"
      Top             =   5085
      Width           =   795
   End
   Begin VB.TextBox txtD4 
      Height          =   270
      Left            =   12420
      TabIndex        =   59
      Text            =   "Text22"
      Top             =   4740
      Width           =   795
   End
   Begin VB.TextBox txtC4 
      Height          =   270
      Left            =   12420
      TabIndex        =   58
      Text            =   "Text21"
      Top             =   4410
      Width           =   795
   End
   Begin VB.TextBox txtH3 
      Height          =   270
      Left            =   11310
      TabIndex        =   57
      Text            =   "Text26"
      Top             =   6120
      Width           =   795
   End
   Begin VB.TextBox txtG3 
      Height          =   270
      Left            =   11310
      TabIndex        =   56
      Text            =   "Text25"
      Top             =   5730
      Width           =   795
   End
   Begin VB.TextBox txtF3 
      Height          =   270
      Left            =   11310
      TabIndex        =   55
      Text            =   "Text24"
      Top             =   5415
      Width           =   795
   End
   Begin VB.TextBox txtE3 
      Height          =   270
      Left            =   11310
      TabIndex        =   54
      Text            =   "Text23"
      Top             =   5085
      Width           =   795
   End
   Begin VB.TextBox txtD3 
      Height          =   270
      Left            =   11310
      TabIndex        =   53
      Text            =   "Text22"
      Top             =   4740
      Width           =   795
   End
   Begin VB.TextBox txtC3 
      Height          =   270
      Left            =   11310
      TabIndex        =   52
      Text            =   "Text21"
      Top             =   4410
      Width           =   795
   End
   Begin VB.TextBox txtH2 
      Height          =   270
      Left            =   9900
      TabIndex        =   51
      Text            =   "Text26"
      Top             =   6120
      Width           =   795
   End
   Begin VB.TextBox txtG2 
      Height          =   270
      Left            =   9900
      TabIndex        =   50
      Text            =   "Text25"
      Top             =   5730
      Width           =   795
   End
   Begin VB.TextBox txtF2 
      Height          =   270
      Left            =   9900
      TabIndex        =   49
      Text            =   "Text24"
      Top             =   5415
      Width           =   795
   End
   Begin VB.TextBox txtE2 
      Height          =   270
      Left            =   9900
      TabIndex        =   48
      Text            =   "Text23"
      Top             =   5085
      Width           =   795
   End
   Begin VB.TextBox txtD2 
      Height          =   270
      Left            =   9900
      TabIndex        =   47
      Text            =   "Text22"
      Top             =   4740
      Width           =   795
   End
   Begin VB.TextBox txtC2 
      Height          =   270
      Left            =   9900
      TabIndex        =   46
      Text            =   "Text21"
      Top             =   4410
      Width           =   795
   End
   Begin VB.TextBox txtH1 
      Height          =   270
      Left            =   2820
      TabIndex        =   45
      Text            =   "Text20"
      Top             =   6120
      Width           =   6735
   End
   Begin VB.TextBox txtG1 
      Height          =   270
      Left            =   2820
      TabIndex        =   44
      Text            =   "Text19"
      Top             =   5784
      Width           =   6735
   End
   Begin VB.TextBox txtF1 
      Height          =   270
      Left            =   2820
      TabIndex        =   43
      Text            =   "Text18"
      Top             =   5448
      Width           =   6735
   End
   Begin VB.TextBox txtE1 
      Height          =   270
      Left            =   2820
      TabIndex        =   42
      Text            =   "Text17"
      Top             =   5112
      Width           =   6735
   End
   Begin VB.TextBox txtD1 
      Height          =   270
      Left            =   2820
      TabIndex        =   41
      Text            =   "Text16"
      Top             =   4776
      Width           =   6735
   End
   Begin VB.TextBox txtA1 
      Height          =   825
      Left            =   2760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   40
      Text            =   "b2.frx":0A82
      Top             =   2250
      Width           =   6915
   End
   Begin VB.TextBox txtC1 
      Height          =   270
      Left            =   2820
      TabIndex        =   39
      Text            =   "Text14"
      Top             =   4440
      Width           =   6735
   End
   Begin VB.TextBox txtB5 
      Height          =   645
      Left            =   14040
      TabIndex        =   28
      Text            =   "Text13"
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txtB4 
      Height          =   645
      Left            =   12330
      TabIndex        =   27
      Text            =   "Text12"
      Top             =   3120
      Width           =   915
   End
   Begin VB.TextBox txtA5 
      Height          =   1545
      Left            =   14040
      TabIndex        =   26
      Text            =   "Text11"
      Top             =   1410
      Width           =   705
   End
   Begin VB.TextBox txtA4 
      Height          =   1545
      Left            =   12390
      TabIndex        =   25
      Text            =   "Text10"
      Top             =   1410
      Width           =   705
   End
   Begin VB.TextBox txtB3 
      Height          =   645
      Left            =   11220
      TabIndex        =   24
      Text            =   "Text9"
      Top             =   3120
      Width           =   825
   End
   Begin VB.TextBox txtA3 
      Height          =   1545
      Left            =   11250
      TabIndex        =   23
      Text            =   "Text8"
      Top             =   1410
      Width           =   705
   End
   Begin VB.TextBox txtB2 
      Height          =   645
      Left            =   9840
      TabIndex        =   22
      Text            =   "30%"
      Top             =   3120
      Width           =   765
   End
   Begin VB.TextBox txtA2 
      Height          =   1545
      Left            =   9870
      TabIndex        =   21
      Text            =   "50%"
      Top             =   1410
      Width           =   705
   End
   Begin VB.TextBox txtB1 
      Height          =   645
      Left            =   2760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Text            =   "b2.frx":0A89
      Top             =   3120
      Width           =   6915
   End
   Begin VB.TextBox txtA 
      Height          =   825
      Left            =   2760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Text            =   "b2.frx":0A8F
      Top             =   1380
      Width           =   6915
   End
   Begin VB.TextBox txtYwy 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txtBm 
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
      Left            =   11250
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   480
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker txtM 
      Height          =   345
      Left            =   5340
      TabIndex        =   120
      Top             =   450
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy年MM月"
      Format          =   55967747
      CurrentDate     =   39415
   End
   Begin VB.Label lblZ5 
      Caption         =   "lblZ5"
      ForeColor       =   &H00004080&
      Height          =   225
      Left            =   14100
      TabIndex        =   123
      Top             =   7230
      Width           =   675
   End
   Begin VB.Label lblZ4 
      Caption         =   "Label36"
      ForeColor       =   &H00004080&
      Height          =   225
      Left            =   12450
      TabIndex        =   122
      Top             =   7230
      Width           =   675
   End
   Begin VB.Label lblZ3 
      Caption         =   "Label35"
      ForeColor       =   &H00004080&
      Height          =   225
      Left            =   11370
      TabIndex        =   121
      Top             =   7230
      Width           =   675
   End
   Begin VB.Label lblMTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   4
      Left            =   11580
      TabIndex        =   119
      Top             =   8700
      Width           =   1245
   End
   Begin VB.Label lblMQM 
      Caption         =   "被考核员工"
      Height          =   225
      Index           =   4
      Left            =   11640
      TabIndex        =   118
      Top             =   8010
      Width           =   1095
   End
   Begin VB.Label lblMTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   3
      Left            =   10290
      TabIndex        =   116
      Top             =   8700
      Width           =   1245
   End
   Begin VB.Label lblMQM 
      Caption         =   "人事"
      Height          =   225
      Index           =   3
      Left            =   9150
      TabIndex        =   115
      Top             =   8010
      Width           =   495
   End
   Begin VB.Label lblMTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   2
      Left            =   8970
      TabIndex        =   113
      Top             =   8700
      Width           =   1245
   End
   Begin VB.Label lblMQM 
      Caption         =   "部门主管"
      Height          =   225
      Index           =   2
      Left            =   10380
      TabIndex        =   112
      Top             =   8010
      Width           =   885
   End
   Begin VB.Line Line13 
      X1              =   750
      X2              =   750
      Y1              =   420
      Y2              =   6810
   End
   Begin VB.Label lblKid 
      Caption         =   "lblKid"
      Height          =   225
      Left            =   0
      TabIndex        =   110
      Top             =   0
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblZF 
      Caption         =   "Label35"
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
      Left            =   4860
      TabIndex        =   106
      Top             =   7410
      Width           =   1215
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
      Left            =   6570
      TabIndex        =   105
      Top             =   7500
      Width           =   5475
   End
   Begin VB.Shape Shape1 
      Height          =   6795
      Left            =   60
      Top             =   420
      Width           =   14985
   End
   Begin VB.Line Line12 
      X1              =   750
      X2              =   15030
      Y1              =   6780
      Y2              =   6780
   End
   Begin VB.Line Line11 
      Index           =   1
      X1              =   750
      X2              =   15000
      Y1              =   6390
      Y2              =   6390
   End
   Begin VB.Line Line11 
      Index           =   0
      X1              =   750
      X2              =   15030
      Y1              =   6090
      Y2              =   6090
   End
   Begin VB.Line Line10 
      X1              =   750
      X2              =   14580
      Y1              =   5730
      Y2              =   5730
   End
   Begin VB.Line Line9 
      X1              =   750
      X2              =   15030
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line8 
      X1              =   780
      X2              =   15030
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line7 
      X1              =   750
      X2              =   15030
      Y1              =   4740
      Y2              =   4740
   End
   Begin VB.Line Line6 
      X1              =   90
      X2              =   15000
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label lblMQM 
      Caption         =   "被考核员工"
      Height          =   225
      Index           =   0
      Left            =   6300
      TabIndex        =   92
      Top             =   8010
      Width           =   1095
   End
   Begin VB.Label lblMTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   0
      Left            =   6240
      TabIndex        =   91
      Top             =   8700
      Width           =   1245
   End
   Begin VB.Label lblMTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   1
      Left            =   7590
      TabIndex        =   90
      Top             =   8700
      Width           =   1305
   End
   Begin VB.Label lblMQM 
      Caption         =   "直接主管"
      Height          =   225
      Index           =   1
      Left            =   7710
      TabIndex        =   89
      Top             =   8010
      Width           =   945
   End
   Begin VB.Label Label34 
      Caption         =   "部门主管评语"
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
      Left            =   30
      TabIndex        =   86
      Top             =   8550
      Width           =   1605
   End
   Begin VB.Label Label33 
      Caption         =   "直接主管评语"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   30
      TabIndex        =   85
      Top             =   7860
      Width           =   1605
   End
   Begin VB.Label Label32 
      Caption         =   "创新及自主学习"
      Height          =   225
      Left            =   1380
      TabIndex        =   74
      Top             =   6870
      Width           =   1365
   End
   Begin VB.Label Label31 
      Caption         =   "员工考勤"
      Height          =   255
      Left            =   1380
      TabIndex        =   73
      Top             =   6480
      Width           =   885
   End
   Begin VB.Label Label30 
      Caption         =   "总得分=工作业绩+工作能力和态度+加分项="
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
      Height          =   285
      Left            =   120
      TabIndex        =   72
      Top             =   7410
      Width           =   4665
   End
   Begin VB.Label Label29 
      Caption         =   "加分项目"
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
      Left            =   120
      TabIndex        =   71
      Top             =   6900
      Width           =   975
   End
   Begin VB.Label Label28 
      Caption         =   "出勤5%"
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
      Left            =   120
      TabIndex        =   70
      Top             =   6420
      Width           =   735
   End
   Begin VB.Label Label27 
      Caption         =   "敬业精神"
      Height          =   195
      Left            =   1350
      TabIndex        =   38
      Top             =   6150
      Width           =   915
   End
   Begin VB.Label Label26 
      Caption         =   "团队协作"
      Height          =   195
      Left            =   1350
      TabIndex        =   37
      Top             =   5820
      Width           =   915
   End
   Begin VB.Label Label25 
      Caption         =   "反应速度"
      Height          =   195
      Left            =   1350
      TabIndex        =   36
      Top             =   5490
      Width           =   915
   End
   Begin VB.Label Label24 
      Caption         =   "执行力"
      Height          =   195
      Left            =   1350
      TabIndex        =   35
      Top             =   5160
      Width           =   915
   End
   Begin VB.Label Label23 
      Caption         =   "沟通能力"
      Height          =   195
      Left            =   1350
      TabIndex        =   34
      Top             =   4830
      Width           =   915
   End
   Begin VB.Label Label22 
      Caption         =   "知识技能"
      Height          =   195
      Left            =   1350
      TabIndex        =   33
      Top             =   4500
      Width           =   915
   End
   Begin VB.Label Label21 
      Caption         =   "部门主管"
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
      Left            =   13920
      TabIndex        =   32
      Top             =   3900
      Width           =   1035
   End
   Begin VB.Label Label20 
      Caption         =   "直接主管"
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
      Left            =   12360
      TabIndex        =   31
      Top             =   3900
      Width           =   1035
   End
   Begin VB.Label Label19 
      Caption         =   "自评"
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
      Left            =   11340
      TabIndex        =   30
      Top             =   3900
      Width           =   585
   End
   Begin VB.Label Label18 
      Caption         =   "权重"
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
      Left            =   9930
      TabIndex        =   29
      Top             =   3900
      Width           =   675
   End
   Begin VB.Label Label17 
      Caption         =   "工作能力和态度15%"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   180
      TabIndex        =   20
      Top             =   4410
      Width           =   555
   End
   Begin VB.Label Label4 
      Caption         =   "考核项目"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1050
      Width           =   525
   End
   Begin VB.Label Label16 
      Caption         =   "工作业绩80%"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   120
      TabIndex        =   19
      Top             =   2340
      Width           =   675
   End
   Begin VB.Line Line5 
      X1              =   9780
      X2              =   9780
      Y1              =   420
      Y2              =   7200
   End
   Begin VB.Line Line4 
      X1              =   11130
      X2              =   11130
      Y1              =   420
      Y2              =   7200
   End
   Begin VB.Line Line3 
      X1              =   90
      X2              =   15030
      Y1              =   3870
      Y2              =   3870
   End
   Begin VB.Label Label15 
      Caption         =   "考核细则"
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
      Left            =   3360
      TabIndex        =   18
      Top             =   3900
      Width           =   1035
   End
   Begin VB.Label Label14 
      Caption         =   "考核内容"
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
      Left            =   900
      TabIndex        =   17
      Top             =   3900
      Width           =   1035
   End
   Begin VB.Line Line2 
      X1              =   780
      X2              =   15030
      Y1              =   3090
      Y2              =   3090
   End
   Begin VB.Line Line1 
      X1              =   750
      X2              =   15030
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label13 
      Caption         =   "日常工作达标情况"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1470
      TabIndex        =   14
      Top             =   3180
      Width           =   1035
   End
   Begin VB.Label Label12 
      Caption         =   "专项工作"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   1590
      TabIndex        =   13
      Top             =   1470
      Width           =   255
   End
   Begin VB.Label Label11 
      Caption         =   "客户(20%)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14010
      TabIndex        =   12
      Top             =   990
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "直接主管(60%)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12360
      TabIndex        =   11
      Top             =   990
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "自评(20%)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11160
      TabIndex        =   10
      Top             =   990
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "重要性系数"
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
      Left            =   9870
      TabIndex        =   9
      Top             =   990
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "完成情况简述"
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
      Left            =   3810
      TabIndex        =   8
      Top             =   990
      Width           =   1545
   End
   Begin VB.Label Label6 
      Caption         =   "考核重点："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1500
      TabIndex        =   7
      Top             =   990
      Width           =   1155
   End
   Begin VB.Label Label2 
      Caption         =   "姓名："
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
      Left            =   90
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "部门"
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
      Left            =   10020
      TabIndex        =   4
      Top             =   480
      Width           =   585
   End
   Begin VB.Label Label5 
      Caption         =   "月份"
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
      Left            =   4140
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "员工月度考核表"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4710
      TabIndex        =   0
      Top             =   60
      Width           =   2745
   End
End
Attribute VB_Name = "b2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim timZm As Integer '1保存（直接主管）2保存（员工）3保存（人事）5签字

Private Sub cmdBack_Click()
Me.Visible = False
frmZu.TBa.Buttons(7).Value = tbrUnpressed
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
End If
End Sub

Private Sub cmdDing_Click()
Dim tt As String
On Error Resume Next

If optT2.Value = True And txtQM.Text = "" Then
    MsgBox ("请您一定要告诉拒绝我的理由!  :) ")
    Exit Sub
End If
timZm = 5 '签字
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "员工表2"
    mod1.cmd.Parameters("@NBLX") = "签字"
    mod1.cmd.Parameters("@bh") = lblKid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtYwy.Text
    mod1.cmd.Parameters("@mt2") = txtYwy.ToolTipText
    mod1.cmd.Parameters("@mt3") = txtBm.Text
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
    mod1.cmd.Parameters("@mt20") = lblMQM(Val(lblLc.Caption) - 1).Caption
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
    Call mod1.REV
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

Private Sub cmdMod_Click()
If Trim(lblLcUid.Caption) <> mod1.DHid Then Exit Sub
If Val(lblLc.Caption) = 0 Then '主管初设

    txtC1.Locked = False
    txtD1.Locked = False
    txtE1.Locked = False
    txtF1.Locked = False
    txtG1.Locked = False
    txtH1.Locked = False
    txtI1.Locked = False

    txtC2.Locked = False
    txtD2.Locked = False
    txtE2.Locked = False
    txtF2.Locked = False
    txtG2.Locked = False
    txtH2.Locked = False
    txtI2.Locked = False

    cmdSave.Enabled = True
ElseIf Val(lblLc.Caption) = 1 Then '自评
    If b1.cmdMQm(1).Caption = "" Then Exit Sub
    txtA3.Locked = False
    txtB3.Locked = False
    txtC3.Locked = False
    txtD3.Locked = False
    txtE3.Locked = False
    txtF3.Locked = False
    txtG3.Locked = False
    txtH3.Locked = False
    txtI3.Locked = False
    txtJ1.Locked = False
    txtJ2.Locked = False
    txtJ3.Locked = False
    cmdSave.Enabled = True
ElseIf Val(lblLc.Caption) = 2 Then '直接主管
    txtA1.Locked = False
    txtB1.Locked = False
    txtA4.Locked = False
    txtB4.Locked = False
    txtC4.Locked = False
    txtD4.Locked = False
    txtE4.Locked = False
    txtF4.Locked = False
    txtG4.Locked = False
    txtH4.Locked = False
    txtI4.Locked = False
    txtJ4.Locked = False
    txtJ1.Locked = False
    txtJ2.Locked = False
    txtZjp.Locked = False
    cmdSave.Enabled = True
ElseIf Val(lblLc.Caption) = 3 Then '人事
    txtA5.Locked = False
    txtB5.Locked = False
    txtC5.Locked = False
    txtD5.Locked = False
    txtE5.Locked = False
    txtF5.Locked = False
    txtG5.Locked = False
    txtH5.Locked = False
    txtI5.Locked = False
    txtJ5.Locked = False
    cmdSave.Enabled = True
ElseIf Val(lblLc.Caption) = 4 Then '部门主管
    txtA5.Locked = False
    txtB5.Locked = False
    txtC5.Locked = False
    txtD5.Locked = False
    txtE5.Locked = False
    txtF5.Locked = False
    txtG5.Locked = False
    txtH5.Locked = False
    txtI5.Locked = False
    txtJ5.Locked = False
    txtA1.Locked = False
    txtB1.Locked = False
    txtA4.Locked = False
    txtB4.Locked = False
    txtC4.Locked = False
    txtD4.Locked = False
    txtE4.Locked = False
    txtF4.Locked = False
    txtG4.Locked = False
    txtH4.Locked = False
    txtI4.Locked = False
    txtJ4.Locked = False
    txtJ1.Locked = False
    txtJ2.Locked = False
    txtBmp.Locked = False
    cmdSave.Enabled = True
End If
End Sub

Private Sub cmdMQm_Click(Index As Integer)
Dim QZ As Integer
Dim oo As Integer
On Error Resume Next
If Me.Visible = False Then Exit Sub
'先检测权重是否超出100%
QZ = 0
If b1.cmdMQm(1).Caption = "" Then Exit Sub
If Trim(lblLcUid.Caption) <> mod1.DHid Then
    MsgBox "此处应由" & lblLcUid.ToolTipText & "签字! 请您不要再点"
    Exit Sub
End If

If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If
If Index + 1 <> lblLc.Caption Then '不能在不相干的位置上乱点
    Exit Sub
End If


If Index = 0 Then '初次只能签字，不能驳回。
    optT2.Enabled = False
Else
    optT2.Enabled = True
End If
OptT1.Value = True

frmQm.Visible = True
End Sub

Private Sub cmdSave_Click()
Dim tt As String
On Error Resume Next
b2.lblZ3.Caption = Val(b2.txtA3.Text) * Val(b2.txtA2.Text) / 100 + Val(b2.txtB3.Text) * Val(b2.txtB2.Text) / 100 + Val(b2.txtC3.Text) * Val(b2.txtC2.Text) / 100 + _
Val(b2.txtD3.Text) * Val(b2.txtD2.Text) / 100 + Val(b2.txtE3.Text) * Val(b2.txtE2.Text) / 100 + Val(b2.txtF3.Text) * Val(b2.txtF2.Text) / 100 + Val(b2.txtG3.Text) * Val(b2.txtG2.Text) / 100 + _
Val(b2.txtH3.Text) * Val(b2.txtH2.Text) / 100 + Val(b2.txtI3.Text) * Val(b2.txtI2.Text) / 100
b2.lblZ4.Caption = Val(b2.txtA4.Text) * Val(b2.txtA2.Text) / 100 + Val(b2.txtB4.Text) * Val(b2.txtB2.Text) / 100 + Val(b2.txtC4.Text) * Val(b2.txtC2.Text) / 100 + _
Val(b2.txtD4.Text) * Val(b2.txtD2.Text) / 100 + Val(b2.txtE4.Text) * Val(b2.txtE2.Text) / 100 + Val(b2.txtF4.Text) * Val(b2.txtF2.Text) / 100 + Val(b2.txtG4.Text) * Val(b2.txtG2.Text) / 100 + _
Val(b2.txtH4.Text) * Val(b2.txtH2.Text) / 100 + Val(b2.txtI4.Text) * Val(b2.txtI2.Text) / 100
b2.lblZ5.Caption = Val(b2.txtA5.Text) * Val(b2.txtA2.Text) / 100 + Val(b2.txtB5.Text) * Val(b2.txtB2.Text) / 100 + Val(b2.txtC5.Text) * Val(b2.txtC2.Text) / 100 + _
Val(b2.txtD5.Text) * Val(b2.txtD2.Text) / 100 + Val(b2.txtE5.Text) * Val(b2.txtE2.Text) / 100 + Val(b2.txtF5.Text) * Val(b2.txtF2.Text) / 100 + Val(b2.txtG5.Text) * Val(b2.txtG2.Text) / 100 + _
Val(b2.txtH5.Text) * Val(b2.txtH2.Text) / 100 + Val(b2.txtI5.Text) * Val(b2.txtI2.Text) / 100
b2.lblZF.Caption = Val(lblZ3.Caption) * 0.2 + Val(lblZ4.Caption) * 0.6 + Val(lblZ5.Caption) * 0.2 + Val(txtJ4.Text)

txtC2.Text = Val(txtC2.Text) & "%"
txtD2.Text = Val(txtD2.Text) & "%"
txtE2.Text = Val(txtE2.Text) & "%"
txtF2.Text = Val(txtF2.Text) & "%"
txtG2.Text = Val(txtG2.Text) & "%"
txtH2.Text = Val(txtH2.Text) & "%"
txtI2.Text = Val(txtI2.Text) & "%"
txtJ2.Text = Val(txtJ2.Text) & "%"


timZm = 1 '员工表1保存
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
If Val(lblLc.Caption) = 0 Then '直接主管
    mod1.cmd.Parameters("@NB") = "员工表2"
    mod1.cmd.Parameters("@NBLX") = "保存1"
    mod1.cmd.Parameters("@bh") = lblKid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtA1.Text
    mod1.cmd.Parameters("@mt2") = txtB1.Text
    mod1.cmd.Parameters("@mt3") = txtC1.Text
    mod1.cmd.Parameters("@mt4") = txtD1.Text
    mod1.cmd.Parameters("@mt5") = txtE1.Text
    mod1.cmd.Parameters("@mt6") = txtF1.Text
    mod1.cmd.Parameters("@mt7") = txtG1.Text
    mod1.cmd.Parameters("@mt8") = txtH1.Text
    mod1.cmd.Parameters("@mt9") = txtI1.Text
    mod1.cmd.Parameters("@mt10") = txtJ1.Text
    mod1.cmd.Parameters("@mt11") = ""
    mod1.cmd.Parameters("@mt12") = txtZjp.Text
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
    mod1.cmd.Parameters("@mm1") = Val(txtA2.Text)
    mod1.cmd.Parameters("@mm2") = Val(txtB2.Text)
    mod1.cmd.Parameters("@mm3") = Val(txtC2.Text)
    mod1.cmd.Parameters("@mm4") = Val(txtD2.Text)
    mod1.cmd.Parameters("@mm5") = Val(txtE2.Text)
    mod1.cmd.Parameters("@mm6") = Val(txtF2.Text)
    mod1.cmd.Parameters("@mm7") = Val(txtG2.Text)
    mod1.cmd.Parameters("@mm8") = Val(txtH2.Text)
    mod1.cmd.Parameters("@mm9") = Val(txtI2.Text)
    mod1.cmd.Parameters("@mm10") = Val(txtJ2.Text)
    mod1.cmd.Parameters("@mm11") = Val(txtA4.Text)
    mod1.cmd.Parameters("@mm12") = Val(txtB4.Text)
    mod1.cmd.Parameters("@mm13") = Val(txtC4.Text)
    mod1.cmd.Parameters("@mm14") = Val(txtD4.Text)
    mod1.cmd.Parameters("@mm15") = Val(txtE4.Text)
    mod1.cmd.Parameters("@mm16") = Val(txtF4.Text)
    mod1.cmd.Parameters("@mm17") = Val(txtG4.Text)
    mod1.cmd.Parameters("@mm18") = Val(txtH4.Text)
    mod1.cmd.Parameters("@mm19") = Val(txtI4.Text)
    mod1.cmd.Parameters("@mm20") = Val(txtJ4.Text)
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
ElseIf Val(lblLc.Caption) = 1 Then               '员工保存
    mod1.cmd.Parameters("@NB") = "员工表2"
    mod1.cmd.Parameters("@NBLX") = "保存2"
    mod1.cmd.Parameters("@bh") = lblKid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtJ1.Text
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
    mod1.cmd.Parameters("@mm1") = Val(txtA3.Text)
    mod1.cmd.Parameters("@mm2") = Val(txtB3.Text)
    mod1.cmd.Parameters("@mm3") = Val(txtC3.Text)
    mod1.cmd.Parameters("@mm4") = Val(txtD3.Text)
    mod1.cmd.Parameters("@mm5") = Val(txtE3.Text)
    mod1.cmd.Parameters("@mm6") = Val(txtF3.Text)
    mod1.cmd.Parameters("@mm7") = Val(txtG3.Text)
    mod1.cmd.Parameters("@mm8") = Val(txtH3.Text)
    mod1.cmd.Parameters("@mm9") = Val(txtI3.Text)
    mod1.cmd.Parameters("@mm10") = Val(txtJ3.Text)
    mod1.cmd.Parameters("@mm11") = Val(txtJ2.Text)
    mod1.cmd.Parameters("@mm12") = 0
    mod1.cmd.Parameters("@mm13") = 0
    mod1.cmd.Parameters("@mm14") = 0
    mod1.cmd.Parameters("@mm15") = 0
    mod1.cmd.Parameters("@mm16") = 0
    mod1.cmd.Parameters("@mm17") = 0
    mod1.cmd.Parameters("@mm18") = 0
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = Val(lblZF.Caption)
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
ElseIf Val(lblLc.Caption) = 2 Then               '直接主管
    mod1.cmd.Parameters("@NB") = "员工表2"
    mod1.cmd.Parameters("@NBLX") = "保存3"
    mod1.cmd.Parameters("@bh") = lblKid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtJ1.Text
    mod1.cmd.Parameters("@mt2") = txtA1.Text
    mod1.cmd.Parameters("@mt3") = txtB1.Text
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = txtZjp.Text
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
    mod1.cmd.Parameters("@mm1") = Val(txtA4.Text)
    mod1.cmd.Parameters("@mm2") = Val(txtB4.Text)
    mod1.cmd.Parameters("@mm3") = Val(txtC4.Text)
    mod1.cmd.Parameters("@mm4") = Val(txtD4.Text)
    mod1.cmd.Parameters("@mm5") = Val(txtE4.Text)
    mod1.cmd.Parameters("@mm6") = Val(txtF4.Text)
    mod1.cmd.Parameters("@mm7") = Val(txtG4.Text)
    mod1.cmd.Parameters("@mm8") = Val(txtH4.Text)
    mod1.cmd.Parameters("@mm9") = Val(txtI4.Text)
    mod1.cmd.Parameters("@mm10") = Val(txtJ4.Text)
    mod1.cmd.Parameters("@mm11") = Val(txtJ2.Text)
    mod1.cmd.Parameters("@mm12") = 0
    mod1.cmd.Parameters("@mm13") = 0
    mod1.cmd.Parameters("@mm14") = 0
    mod1.cmd.Parameters("@mm15") = 0
    mod1.cmd.Parameters("@mm16") = 0
    mod1.cmd.Parameters("@mm17") = 0
    mod1.cmd.Parameters("@mm18") = 0
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = Val(lblZF.Caption)
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
ElseIf Val(lblLc.Caption) = 3 Then               '人事
    mod1.cmd.Parameters("@NB") = "员工表2"
    mod1.cmd.Parameters("@NBLX") = "保存4"
    mod1.cmd.Parameters("@bh") = lblKid.Caption
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
    mod1.cmd.Parameters("@mm1") = Val(txtA5.Text)
    mod1.cmd.Parameters("@mm2") = Val(txtB5.Text)
    mod1.cmd.Parameters("@mm3") = Val(txtC5.Text)
    mod1.cmd.Parameters("@mm4") = Val(txtD5.Text)
    mod1.cmd.Parameters("@mm5") = Val(txtE5.Text)
    mod1.cmd.Parameters("@mm6") = Val(txtF5.Text)
    mod1.cmd.Parameters("@mm7") = Val(txtG5.Text)
    mod1.cmd.Parameters("@mm8") = Val(txtH5.Text)
    mod1.cmd.Parameters("@mm9") = Val(txtI5.Text)
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
    mod1.cmd.Parameters("@mm20") = Val(lblZF.Caption)
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
ElseIf Val(lblLc.Caption) = 4 Then               '部门主管
    mod1.cmd.Parameters("@NB") = "员工表2"
    mod1.cmd.Parameters("@NBLX") = "保存5"
    mod1.cmd.Parameters("@bh") = lblKid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtJ1.Text
    mod1.cmd.Parameters("@mt2") = txtA1.Text
    mod1.cmd.Parameters("@mt3") = txtB1.Text
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = txtBmp.Text
    mod1.cmd.Parameters("@mt6") = Val(txtJ2.Text)
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
    mod1.cmd.Parameters("@mm1") = Val(txtA4.Text)
    mod1.cmd.Parameters("@mm2") = Val(txtB4.Text)
    mod1.cmd.Parameters("@mm3") = Val(txtC4.Text)
    mod1.cmd.Parameters("@mm4") = Val(txtD4.Text)
    mod1.cmd.Parameters("@mm5") = Val(txtE4.Text)
    mod1.cmd.Parameters("@mm6") = Val(txtF4.Text)
    mod1.cmd.Parameters("@mm7") = Val(txtG4.Text)
    mod1.cmd.Parameters("@mm8") = Val(txtH4.Text)
    mod1.cmd.Parameters("@mm9") = Val(txtI4.Text)
    mod1.cmd.Parameters("@mm10") = Val(txtJ4.Text)
    mod1.cmd.Parameters("@mm11") = Val(txtA5.Text)
    mod1.cmd.Parameters("@mm12") = Val(txtB5.Text)
    mod1.cmd.Parameters("@mm13") = Val(txtC5.Text)
    mod1.cmd.Parameters("@mm14") = Val(txtD5.Text)
    mod1.cmd.Parameters("@mm15") = Val(txtE5.Text)
    mod1.cmd.Parameters("@mm16") = Val(txtF5.Text)
    mod1.cmd.Parameters("@mm17") = Val(txtG5.Text)
    mod1.cmd.Parameters("@mm18") = Val(txtH5.Text)
    mod1.cmd.Parameters("@mm19") = Val(txtI5.Text)
    mod1.cmd.Parameters("@mm20") = Val(lblZF.Caption)
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
    Call mod1.REV
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

Private Sub cmdZuan_Click()
If b1.Visible = True Then
    b2.Visible = True
    b1.Visible = False
ElseIf b2.Visible = True Then
    'b3.Visible = True
    b2.Visible = False
    b1.Visible = True
'    b3.Visible = False
'ElseIf b3.Visible = True Then
'    b1.Visible = True
'    b3.Visible = False
End If
End Sub

Private Sub Form_Click()
frmQm.Visible = False
End Sub

Private Sub Form_Load()
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0
frmQm.Left = 0
frmQm.Top = 7380
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmZu.TBa.Buttons(7).Value = tbrUnpressed
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
End If
End Sub

Public Sub b2Locked()
txtA1.Locked = True
txtB1.Locked = True
txtC1.Locked = True
txtD1.Locked = True
txtE1.Locked = True
txtF1.Locked = True
txtG1.Locked = True
txtH1.Locked = True
txtI1.Locked = True
txtJ1.Locked = True

txtA2.Locked = True
txtB2.Locked = True
txtC2.Locked = True
txtD2.Locked = True
txtE2.Locked = True
txtF2.Locked = True
txtG2.Locked = True
txtH2.Locked = True
txtI2.Locked = True
txtJ2.Locked = True

txtA3.Locked = True
txtB3.Locked = True
txtC3.Locked = True
txtD3.Locked = True
txtE3.Locked = True
txtF3.Locked = True
txtG3.Locked = True
txtH3.Locked = True
txtI3.Locked = True
txtJ3.Locked = True

txtA4.Locked = True
txtB4.Locked = True
txtC4.Locked = True
txtD4.Locked = True
txtE4.Locked = True
txtF4.Locked = True
txtG4.Locked = True
txtH4.Locked = True
txtI4.Locked = True
txtJ4.Locked = True

txtA5.Locked = True
txtB5.Locked = True
txtC5.Locked = True
txtD5.Locked = True
txtE5.Locked = True
txtF5.Locked = True
txtG5.Locked = True
txtH5.Locked = True
txtI5.Locked = True
txtJ5.Locked = True

txtZjp.Locked = True
txtBmp.Locked = True

End Sub

Private Sub timQuit_Timer()
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0


If timZm = 1 Then
    Call Me.b2Locked
    cmdSave.Enabled = False
ElseIf timZm = 5 Then '签字
    cmdDing.Enabled = True
    txtQM.Text = ""
    frmQm.Visible = False
    lblTX.Visible = True
    Call mod1.refEnvent
End If
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
    If timZm = 1 Then

    ElseIf timZm = 5 Then '签名
        If OptT1.Value = True Then
            cmdMQm(lblLc.Caption - 1).Caption = mod1.DName
            lblMTm(lblLc.Caption - 1).Caption = mod1.DQda
        Else
            For oo = 0 To 5
                cmdMQm(oo).Caption = ""
                lblMTm(oo).Caption = ""
            Next
        End If
        lblLc.Caption = mod1.WP.Fields("mm1").Value
        lblFwid.Caption = mod1.WP.Fields("mm2").Value
        lblLcUid.ToolTipText = mod1.WP.Fields("mt1").Value
        lblLcUid.Caption = mod1.WP.Fields("mt2").Value
        lblTX.Caption = "下一流程,将跳至" & lblMQM(Val(lblLc.Caption) - 1).Caption & ": " & lblLcUid.ToolTipText
        lblTX.Visible = True
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
        cmdNew.Enabled = False
    End If
    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdNew.Enabled = False
    End If
    Exit Sub
End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub


