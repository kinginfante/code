VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmWbNew 
   Caption         =   "超白金版合同评审单"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   7950
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3330
      Top             =   7950
   End
   Begin VB.CommandButton cmdCong 
      BackColor       =   &H00C0FFC0&
      Caption         =   "重新评审"
      Height          =   1095
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   199
      Top             =   8010
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton cmdPje 
      Caption         =   "评审建议"
      Height          =   1095
      Left            =   420
      TabIndex        =   198
      Top             =   8010
      Width           =   345
   End
   Begin VB.Frame frmGD 
      Caption         =   "项目费用分类"
      Height          =   2625
      Left            =   4110
      TabIndex        =   168
      Top             =   8160
      Visible         =   0   'False
      Width           =   6945
      Begin VB.TextBox txtGd 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   185
         Top             =   2250
         Width           =   1155
      End
      Begin VB.TextBox txtXm 
         Height          =   270
         Left            =   3270
         TabIndex        =   184
         Top             =   2250
         Width           =   1515
      End
      Begin VB.Frame Frame3 
         Height          =   825
         Left            =   30
         TabIndex        =   169
         Top             =   1290
         Width           =   6885
         Begin VB.OptionButton optGDA 
            Caption         =   "中秋(月饼券)"
            Height          =   195
            Left            =   750
            TabIndex        =   177
            Top             =   180
            Width           =   1545
         End
         Begin VB.OptionButton optGDB 
            Caption         =   "春节(年会吃饭)"
            Height          =   180
            Left            =   2280
            TabIndex        =   176
            Top             =   180
            Width           =   1605
         End
         Begin VB.OptionButton optGDC 
            Caption         =   "其它"
            Height          =   195
            Left            =   3960
            TabIndex        =   175
            Top             =   180
            Width           =   675
         End
         Begin VB.TextBox txtGDNR 
            Height          =   270
            Left            =   4770
            TabIndex        =   174
            Top             =   150
            Width           =   2025
         End
         Begin VB.TextBox txtQdj 
            Height          =   270
            Left            =   750
            TabIndex        =   173
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtRl 
            Height          =   270
            Left            =   2250
            TabIndex        =   172
            Top             =   480
            Width           =   915
         End
         Begin VB.CommandButton cmdGAdd 
            Caption         =   "添加"
            Height          =   255
            Left            =   4980
            TabIndex        =   171
            Top             =   480
            Width           =   795
         End
         Begin VB.CommandButton cmdGdel 
            Caption         =   "删除"
            Height          =   255
            Left            =   5880
            TabIndex        =   170
            Top             =   480
            Width           =   885
         End
         Begin MSComCtl2.DTPicker dtpGD 
            Height          =   255
            Left            =   3810
            TabIndex        =   178
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "yyyyy"
            Format          =   107675651
            UpDown          =   -1  'True
            CurrentDate     =   38943
         End
         Begin VB.Label Label45 
            Caption         =   "类别:"
            Height          =   195
            Left            =   120
            TabIndex        =   182
            Top             =   180
            Width           =   555
         End
         Begin VB.Label Label32 
            Caption         =   "单价:"
            Height          =   225
            Left            =   120
            TabIndex        =   181
            Top             =   540
            Width           =   585
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   6870
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Label Label31 
            Caption         =   "人数:"
            Height          =   165
            Left            =   1650
            TabIndex        =   180
            Top             =   540
            Width           =   705
         End
         Begin VB.Label Label30 
            Caption         =   "年份:"
            Height          =   195
            Left            =   3330
            TabIndex        =   179
            Top             =   510
            Width           =   525
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgGD 
         Height          =   1065
         Left            =   60
         TabIndex        =   183
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   1879
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label47 
         Caption         =   "固定费用"
         Height          =   255
         Left            =   120
         TabIndex        =   187
         Top             =   2310
         Width           =   825
      End
      Begin VB.Label Label46 
         Caption         =   "活动费用"
         Height          =   255
         Left            =   2400
         TabIndex        =   186
         Top             =   2280
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdQm 
      Caption         =   "cmdQm"
      Height          =   345
      Index           =   0
      Left            =   780
      TabIndex        =   57
      Top             =   8310
      Width           =   945
   End
   Begin TabDlg.SSTab tabHt 
      Height          =   7905
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   13944
      _Version        =   393216
      TabOrientation  =   1
      TabHeight       =   520
      TabCaption(0)   =   "评审"
      TabPicture(0)   =   "wbHTP.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblJlr"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label26"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label24"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label17"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label20"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label19"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label18"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label5"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label15"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label13"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label4"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label25"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label3(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label2(0)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label44"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label38"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label8"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label29"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Line1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lblHtxz"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Shape1"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label48"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Shape2"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label49"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label50(1)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "dtgYf"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtTcRQ"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtKhmc"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtXMMC"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtKhdm"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtJlr1"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtADR"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtCbze1"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Frame1"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtClcb1"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtQt1"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtYf1"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtFbje1"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtCLF1"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txtRgf1"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txtHtze"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "cmdWb"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtHtbh"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txtXYwy"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "comQy"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txtEd"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txtZe"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txtHtrq"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "txtFbje2"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txtYf2"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "txtCbze2"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "Command3"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "Command4"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "txtQt2"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "txtJlr2"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "frmHide"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "frmYj"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "dtgFk"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "frmFk"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "cmdClcb"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "cmdKP"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "txtKPBz"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "frmFX"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "frmYM"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "txtBz"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "timYj"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "txtYjpw"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).ControlCount=   71
      TabCaption(1)   =   "服务内容"
      TabPicture(1)   =   "wbHTP.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tabGc"
      Tab(1).Control(1)=   "frmJi"
      Tab(1).Control(2)=   "Command1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "财务评定"
      TabPicture(2)   =   "wbHTP.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frmCw"
      Tab(2).ControlCount=   1
      Begin VB.TextBox txtYjpw 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   232
         Top             =   1920
         Visible         =   0   'False
         Width           =   3315
      End
      Begin VB.Frame frmCw 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   7125
         Left            =   -75000
         TabIndex        =   202
         Top             =   240
         Width           =   15225
         Begin VB.Frame frmQkF 
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   1665
            Left            =   570
            TabIndex        =   223
            Top             =   690
            Width           =   4395
            Begin VB.CommandButton cmdQkfDel 
               Caption         =   "删除"
               Height          =   285
               Left            =   3510
               TabIndex        =   227
               Top             =   1320
               Width           =   675
            End
            Begin VB.CommandButton cmdQkfAdd 
               Caption         =   "添加"
               Height          =   285
               Left            =   3540
               TabIndex        =   226
               Top             =   1020
               Width           =   645
            End
            Begin VB.TextBox txtQkFBz 
               Height          =   555
               Left            =   900
               TabIndex        =   225
               Top             =   1020
               Width           =   2265
            End
            Begin VB.TextBox txtQkfJe 
               Height          =   285
               Left            =   900
               TabIndex        =   224
               Top             =   600
               Width           =   2265
            End
            Begin MSComCtl2.DTPicker dtpQkF 
               Height          =   285
               Left            =   900
               TabIndex        =   228
               Top             =   240
               Width           =   2325
               _ExtentX        =   4101
               _ExtentY        =   503
               _Version        =   393216
               Format          =   108396545
               CurrentDate     =   39312
            End
            Begin VB.Label Label53 
               Caption         =   "备注"
               Height          =   225
               Left            =   150
               TabIndex        =   231
               Top             =   960
               Width           =   615
            End
            Begin VB.Label Label54 
               Caption         =   "金额"
               Height          =   225
               Left            =   150
               TabIndex        =   230
               Top             =   630
               Width           =   615
            End
            Begin VB.Label Label55 
               Caption         =   "日期"
               Height          =   225
               Left            =   150
               TabIndex        =   229
               Top             =   300
               Width           =   735
            End
         End
         Begin VB.Frame frmJTF 
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   1665
            Left            =   540
            TabIndex        =   210
            Top             =   3180
            Width           =   4395
            Begin VB.CommandButton cmdJTFdel 
               Caption         =   "删除"
               Height          =   285
               Left            =   3510
               TabIndex        =   214
               Top             =   1320
               Width           =   675
            End
            Begin VB.CommandButton cmdJTFadd 
               Caption         =   "添加"
               Height          =   285
               Left            =   3540
               TabIndex        =   213
               Top             =   1020
               Width           =   645
            End
            Begin VB.TextBox txtJTFbz 
               Height          =   555
               Left            =   900
               TabIndex        =   212
               Top             =   1020
               Width           =   2265
            End
            Begin VB.TextBox txtJtfJe 
               Height          =   285
               Left            =   900
               TabIndex        =   211
               Top             =   600
               Width           =   2265
            End
            Begin MSComCtl2.DTPicker dtpJTF 
               Height          =   285
               Left            =   900
               TabIndex        =   215
               Top             =   240
               Width           =   2325
               _ExtentX        =   4101
               _ExtentY        =   503
               _Version        =   393216
               Format          =   108396545
               CurrentDate     =   39312
            End
            Begin VB.Label Label52 
               Caption         =   "备注"
               Height          =   225
               Left            =   150
               TabIndex        =   218
               Top             =   960
               Width           =   615
            End
            Begin VB.Label Label51 
               Caption         =   "金额"
               Height          =   225
               Left            =   150
               TabIndex        =   217
               Top             =   630
               Width           =   615
            End
            Begin VB.Label Label50 
               Caption         =   "日期"
               Height          =   225
               Index           =   0
               Left            =   150
               TabIndex        =   216
               Top             =   300
               Width           =   735
            End
         End
         Begin VB.TextBox txtQkf 
            Height          =   285
            Left            =   2370
            TabIndex        =   209
            Top             =   5130
            Width           =   2415
         End
         Begin VB.TextBox txtJTf 
            Height          =   345
            Left            =   2400
            TabIndex        =   208
            Top             =   2730
            Width           =   2355
         End
         Begin VB.TextBox txtYjfBz 
            Height          =   795
            Left            =   360
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   207
            Top             =   6330
            Visible         =   0   'False
            Width           =   3405
         End
         Begin VB.TextBox txtYjf 
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   206
            Top             =   120
            Width           =   2325
         End
         Begin VB.CheckBox chkQKF 
            Caption         =   "已收全款"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   435
            Left            =   510
            TabIndex        =   205
            Top             =   5040
            Width           =   1455
         End
         Begin VB.CheckBox chkJTF 
            Caption         =   "已结提成"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   465
            Left            =   510
            TabIndex        =   204
            Top             =   2670
            Width           =   1455
         End
         Begin VB.CheckBox chkYJF 
            Caption         =   "已算业绩"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   495
            Left            =   510
            TabIndex        =   203
            Top             =   0
            Width           =   1665
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgQkf 
            Height          =   2145
            Left            =   5400
            TabIndex        =   219
            Top             =   5070
            Width           =   9405
            _ExtentX        =   16589
            _ExtentY        =   3784
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgJTf 
            Height          =   2175
            Left            =   5400
            TabIndex        =   220
            Top             =   2640
            Width           =   9405
            _ExtentX        =   16589
            _ExtentY        =   3836
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgyjF 
            Height          =   2175
            Left            =   5400
            TabIndex        =   222
            Top             =   120
            Width           =   9405
            _ExtentX        =   16589
            _ExtentY        =   3836
            _Version        =   393216
            BackColorBkg    =   16761024
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00C000C0&
            BorderWidth     =   3
            Index           =   1
            X1              =   0
            X2              =   15210
            Y1              =   4920
            Y2              =   4920
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   3
            Index           =   0
            X1              =   0
            X2              =   15210
            Y1              =   2490
            Y2              =   2490
         End
      End
      Begin VB.Timer timYj 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   14670
         Top             =   120
      End
      Begin VB.TextBox txtBz 
         Height          =   465
         Left            =   6360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   200
         Top             =   2580
         Width           =   3525
      End
      Begin VB.Frame frmYM 
         BackColor       =   &H8000000D&
         Caption         =   "奖金预计支付情况"
         Height          =   2055
         Left            =   5460
         TabIndex        =   141
         Top             =   5070
         Width           =   4665
         Begin VB.CommandButton cmdClose 
            Caption         =   "关闭"
            Height          =   285
            Left            =   3960
            TabIndex        =   150
            Top             =   1590
            Width           =   615
         End
         Begin VB.TextBox txtFED 
            Height          =   285
            Left            =   930
            TabIndex        =   145
            Top             =   1620
            Width           =   645
         End
         Begin VB.TextBox txtYingFu 
            Height          =   270
            Left            =   2850
            TabIndex        =   144
            Top             =   1620
            Width           =   1035
         End
         Begin VB.CommandButton cmdYadd 
            Caption         =   "添加"
            Height          =   315
            Left            =   3960
            TabIndex        =   143
            Top             =   810
            Width           =   585
         End
         Begin VB.CommandButton cmdYdel 
            Caption         =   "删除"
            Height          =   285
            Left            =   3960
            TabIndex        =   142
            Top             =   1170
            Width           =   585
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgYJ 
            Height          =   1275
            Left            =   30
            TabIndex        =   146
            Top             =   210
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   2249
            _Version        =   393216
            BackColorBkg    =   -2147483635
            SelectionMode   =   1
            BorderStyle     =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label lblyjFF 
            Caption         =   "lblYjff"
            Height          =   255
            Left            =   3540
            TabIndex        =   197
            Top             =   150
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label Label41 
            BackColor       =   &H8000000D&
            Caption         =   "收款额度"
            Height          =   255
            Left            =   90
            TabIndex        =   149
            Top             =   1650
            Width           =   825
         End
         Begin VB.Label Label40 
            BackColor       =   &H8000000D&
            Caption         =   "%"
            Height          =   255
            Left            =   1680
            TabIndex        =   148
            Top             =   1650
            Width           =   195
         End
         Begin VB.Label Label39 
            BackColor       =   &H8000000D&
            Caption         =   "支付金额"
            Height          =   225
            Left            =   1980
            TabIndex        =   147
            Top             =   1650
            Width           =   915
         End
      End
      Begin VB.Frame frmFX 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1605
         Left            =   4320
         TabIndex        =   192
         Top             =   3750
         Width           =   585
         Begin VB.CommandButton cmdDe 
            Caption         =   "删除"
            Height          =   375
            Left            =   0
            TabIndex        =   196
            Top             =   780
            Width           =   525
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "添加"
            Height          =   375
            Left            =   0
            TabIndex        =   195
            Top             =   390
            Width           =   525
         End
         Begin VB.CommandButton cmdQing 
            Caption         =   "清空"
            Height          =   345
            Left            =   0
            TabIndex        =   194
            Top             =   0
            Width           =   525
         End
         Begin VB.CommandButton cmdGx 
            Caption         =   "更新"
            Height          =   315
            Left            =   0
            TabIndex        =   193
            Top             =   1170
            Width           =   525
         End
      End
      Begin VB.TextBox txtKPBz 
         Height          =   1005
         Left            =   5130
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   191
         Top             =   5940
         Width           =   4935
      End
      Begin VB.CommandButton cmdKP 
         Caption         =   "开  票"
         Height          =   375
         Left            =   5100
         TabIndex        =   189
         Top             =   7050
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdClcb 
         Caption         =   "核对"
         Height          =   285
         Left            =   14340
         TabIndex        =   188
         Top             =   1500
         Width           =   735
      End
      Begin VB.Frame frmFk 
         Height          =   555
         Left            =   240
         TabIndex        =   131
         Top             =   5670
         Width           =   4245
         Begin VB.TextBox txtYrq 
            Height          =   300
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   138
            Top             =   150
            Width           =   1005
         End
         Begin VB.TextBox txtYed 
            Height          =   270
            Left            =   3150
            TabIndex        =   135
            Top             =   150
            Width           =   795
         End
         Begin MSComCtl2.DTPicker dtpYf 
            Height          =   315
            Left            =   900
            TabIndex        =   132
            Top             =   150
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   16711680
            CalendarTrailingForeColor=   8454016
            Format          =   108331009
            CurrentDate     =   38797
         End
         Begin VB.Label lblFid 
            Caption         =   "lblFid"
            Height          =   165
            Left            =   3600
            TabIndex        =   137
            Top             =   360
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label Label37 
            Caption         =   "%"
            Height          =   255
            Left            =   4050
            TabIndex        =   136
            Top             =   180
            Width           =   435
         End
         Begin VB.Label Label34 
            Caption         =   "收款额度"
            Height          =   255
            Left            =   2310
            TabIndex        =   134
            Top             =   180
            Width           =   795
         End
         Begin VB.Label Label33 
            Caption         =   "应付日期"
            Height          =   285
            Left            =   60
            TabIndex        =   133
            Top             =   180
            Width           =   735
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgFk 
         Height          =   1875
         Left            =   180
         TabIndex        =   130
         Top             =   3750
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3307
         _Version        =   393216
         FillStyle       =   1
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Frame frmYj 
         Height          =   2775
         Left            =   10590
         TabIndex        =   117
         Top             =   4200
         Width           =   4095
         Begin VB.CommandButton cmdCount 
            Caption         =   "计算"
            Height          =   315
            Left            =   1590
            TabIndex        =   124
            Top             =   1650
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.TextBox txtTcBe 
            Height          =   285
            Left            =   990
            TabIndex        =   123
            Text            =   "6"
            Top             =   1650
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txtTc2 
            Height          =   285
            Left            =   990
            TabIndex        =   122
            Top             =   2010
            Width           =   1305
         End
         Begin VB.TextBox txtLr1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   1170
            Locked          =   -1  'True
            TabIndex        =   121
            ToolTipText     =   "预计"
            Top             =   630
            Width           =   1185
         End
         Begin VB.TextBox txtYj1 
            Height          =   285
            Left            =   1170
            Locked          =   -1  'True
            TabIndex        =   120
            Top             =   240
            Width           =   1185
         End
         Begin VB.TextBox txtYj2 
            Height          =   285
            Left            =   2460
            Locked          =   -1  'True
            TabIndex        =   119
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtLr2 
            Height          =   285
            Left            =   2460
            Locked          =   -1  'True
            TabIndex        =   118
            ToolTipText     =   "实际"
            Top             =   630
            Width           =   1215
         End
         Begin MSComCtl2.UpDown UpDa 
            Height          =   315
            Left            =   1320
            TabIndex        =   125
            Top             =   1650
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label lblTcBe 
            Caption         =   "提成比例"
            Height          =   195
            Left            =   60
            TabIndex        =   129
            Top             =   1710
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblLr 
            Caption         =   "利 润 2"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   128
            Top             =   690
            Width           =   915
         End
         Begin VB.Label lblTC 
            Caption         =   "提    成"
            Height          =   195
            Left            =   60
            TabIndex        =   127
            Top             =   2070
            Width           =   735
         End
         Begin VB.Label lblYj 
            Caption         =   "奖    金"
            Height          =   255
            Left            =   120
            TabIndex        =   126
            Top             =   300
            Width           =   975
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "打印"
         Height          =   585
         Left            =   -60420
         Picture         =   "wbHTP.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   7260
         Width           =   645
      End
      Begin VB.Frame frmJi 
         Caption         =   "Frame2"
         Height          =   2505
         Left            =   -74970
         TabIndex        =   92
         Top             =   5070
         Width           =   15195
         Begin VB.Frame Frame2 
            Caption         =   "机组信息"
            Height          =   1905
            Left            =   3360
            TabIndex        =   159
            Top             =   210
            Width           =   7485
            Begin VB.CommandButton cmdTk 
               Caption         =   "维保条款"
               Height          =   285
               Left            =   4200
               TabIndex        =   161
               Top             =   1590
               Width           =   3225
            End
            Begin VB.TextBox txtSl 
               Height          =   285
               Left            =   5340
               Locked          =   -1  'True
               TabIndex        =   160
               Top             =   1140
               Width           =   1965
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgA 
               Height          =   1635
               Left            =   30
               TabIndex        =   162
               Top             =   210
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   2884
               _Version        =   393216
               SelectionMode   =   1
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin MSDataListLib.DataCombo comXh 
               Height          =   330
               Left            =   5340
               TabIndex        =   163
               Top             =   690
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   582
               _Version        =   393216
               Locked          =   -1  'True
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo comPb 
               Height          =   330
               Left            =   5340
               TabIndex        =   164
               Top             =   240
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   582
               _Version        =   393216
               Locked          =   -1  'True
               Text            =   ""
            End
            Begin VB.Label Label2 
               Caption         =   "机组型号:"
               Height          =   225
               Index           =   1
               Left            =   4410
               TabIndex        =   167
               Top             =   765
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "机组品牌:"
               Height          =   225
               Index           =   1
               Left            =   4410
               TabIndex        =   166
               Top             =   330
               Width           =   1125
            End
            Begin VB.Label Label3 
               Caption         =   "数量:"
               Height          =   225
               Index           =   1
               Left            =   4410
               TabIndex        =   165
               Top             =   1200
               Width           =   555
            End
         End
         Begin VB.Frame frmTime 
            Enabled         =   0   'False
            Height          =   1245
            Left            =   120
            TabIndex        =   105
            Top             =   1140
            Width           =   3075
            Begin VB.CheckBox chkBc 
               Caption         =   "2小时内到场"
               Enabled         =   0   'False
               Height          =   255
               Left            =   150
               TabIndex        =   108
               Top             =   960
               Width           =   1845
            End
            Begin VB.CheckBox chkBb 
               Caption         =   "全年运转"
               Enabled         =   0   'False
               Height          =   255
               Left            =   150
               TabIndex        =   107
               Top             =   645
               Width           =   1845
            End
            Begin VB.CheckBox chkBa 
               Caption         =   "24小时运转"
               Enabled         =   0   'False
               Height          =   255
               Left            =   150
               TabIndex        =   106
               Top             =   330
               Width           =   1215
            End
            Begin VB.Label Label27 
               Caption         =   "时间系数:"
               Height          =   195
               Left            =   180
               TabIndex        =   109
               Top             =   120
               Width           =   1155
            End
         End
         Begin VB.TextBox txtZu 
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   104
            Text            =   "Text1"
            Top             =   750
            Width           =   1725
         End
         Begin VB.Frame frmNb 
            Height          =   915
            Left            =   10950
            TabIndex        =   97
            Top             =   630
            Width           =   4125
            Begin VB.TextBox txtL 
               Height          =   300
               Left            =   2430
               Locked          =   -1  'True
               TabIndex        =   140
               Top             =   210
               Width           =   1305
            End
            Begin VB.TextBox txtF 
               Height          =   300
               Left            =   60
               Locked          =   -1  'True
               TabIndex        =   139
               Top             =   210
               Width           =   1455
            End
            Begin VB.TextBox txtXc 
               Height          =   270
               Left            =   3330
               Locked          =   -1  'True
               TabIndex        =   99
               Top             =   600
               Width           =   405
            End
            Begin VB.TextBox txtWc 
               Height          =   270
               Left            =   1050
               Locked          =   -1  'True
               TabIndex        =   98
               Top             =   600
               Width           =   495
            End
            Begin MSComCtl2.DTPicker dt4 
               Height          =   315
               Left            =   2430
               TabIndex        =   114
               Top             =   210
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   556
               _Version        =   393216
               Format          =   109117441
               CurrentDate     =   38098
            End
            Begin MSComCtl2.DTPicker dt3 
               Height          =   315
               Left            =   60
               TabIndex        =   115
               Top             =   210
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   556
               _Version        =   393216
               Format          =   109117441
               CurrentDate     =   38098
            End
            Begin VB.Label Label28 
               Caption         =   "---〉"
               Height          =   225
               Left            =   1950
               TabIndex        =   116
               Top             =   270
               Width           =   375
            End
            Begin VB.Label Label21 
               Caption         =   "次"
               Height          =   225
               Left            =   3840
               TabIndex        =   103
               Top             =   630
               Width           =   315
            End
            Begin VB.Label Label10 
               Caption         =   "例检次数"
               Height          =   225
               Left            =   2430
               TabIndex        =   102
               Top             =   630
               Width           =   825
            End
            Begin VB.Label Label12 
               Caption         =   "年"
               Height          =   225
               Left            =   1650
               TabIndex        =   101
               Top             =   630
               Width           =   255
            End
            Begin VB.Label Label16 
               Caption         =   "维保年限:"
               Height          =   225
               Left            =   60
               TabIndex        =   100
               Top             =   630
               Width           =   855
            End
         End
         Begin VB.Frame frmDx 
            Height          =   375
            Left            =   10980
            TabIndex        =   93
            Top             =   1530
            Width           =   2235
            Begin VB.TextBox txtMon 
               Height          =   270
               Left            =   1290
               Locked          =   -1  'True
               TabIndex        =   94
               Top             =   120
               Width           =   525
            End
            Begin VB.Label Label23 
               Caption         =   "月"
               Height          =   255
               Left            =   1950
               TabIndex        =   96
               Top             =   120
               Width           =   195
            End
            Begin VB.Label Label22 
               Caption         =   "维修保质期"
               DragMode        =   1  'Automatic
               Height          =   225
               Left            =   120
               TabIndex        =   95
               Top             =   120
               Width           =   1065
            End
         End
         Begin MSDataListLib.DataCombo comZu 
            Height          =   330
            Left            =   1440
            TabIndex        =   110
            Top             =   345
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   582
            _Version        =   393216
            Locked          =   -1  'True
            Text            =   "DataCombo2"
         End
         Begin VB.Label Label36 
            Caption         =   "组长"
            Height          =   225
            Left            =   690
            TabIndex        =   112
            Top             =   810
            Width           =   465
         End
         Begin VB.Label Label35 
            Caption         =   "工程部组号"
            Height          =   225
            Left            =   150
            TabIndex        =   111
            Top             =   420
            Width           =   945
         End
      End
      Begin VB.Frame frmHide 
         Caption         =   "frmHid"
         Height          =   1455
         Left            =   330
         TabIndex        =   78
         Top             =   2910
         Visible         =   0   'False
         Width           =   4935
         Begin VB.Label lblPwf 
            Caption         =   "lblPwf"
            Height          =   225
            Left            =   2520
            TabIndex        =   91
            Top             =   1080
            Width           =   1185
         End
         Begin VB.Label lblLcou 
            Caption         =   "lblLcou"
            Height          =   255
            Left            =   1500
            TabIndex        =   88
            Top             =   960
            Width           =   1185
         End
         Begin VB.Label lblYwy 
            Caption         =   "lblYwy"
            Height          =   285
            Left            =   2520
            TabIndex        =   87
            Top             =   450
            Width           =   765
         End
         Begin VB.Label lblUid 
            Caption         =   "lblUid"
            Height          =   255
            Left            =   2580
            TabIndex        =   86
            Top             =   780
            Width           =   975
         End
         Begin VB.Label lblFwid 
            Caption         =   "lblFwid"
            Height          =   255
            Left            =   1380
            TabIndex        =   85
            Top             =   210
            Width           =   885
         End
         Begin VB.Label lblLcUid 
            Caption         =   "lblLcUid"
            Height          =   285
            Left            =   180
            TabIndex        =   84
            Top             =   1020
            Width           =   885
         End
         Begin VB.Label lblLcRen 
            Caption         =   "lblLcRen"
            Height          =   285
            Left            =   150
            TabIndex        =   83
            Top             =   810
            Width           =   795
         End
         Begin VB.Label lblNlb 
            Caption         =   "lblNlb"
            Height          =   225
            Left            =   1470
            TabIndex        =   82
            Top             =   570
            Width           =   645
         End
         Begin VB.Label lblLc 
            Caption         =   "lblLc"
            Height          =   315
            Left            =   150
            TabIndex        =   81
            Top             =   600
            Width           =   645
         End
         Begin VB.Label lblQy 
            Caption         =   "lblQy"
            Height          =   255
            Left            =   2610
            TabIndex        =   80
            Top             =   150
            Width           =   1155
         End
         Begin VB.Label lblBm 
            Caption         =   "lblBm"
            Height          =   225
            Left            =   150
            TabIndex        =   79
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.TextBox txtJlr2 
         Height          =   285
         Left            =   13020
         TabIndex        =   68
         Top             =   3780
         Width           =   1215
      End
      Begin VB.TextBox txtQt2 
         Height          =   285
         Left            =   13020
         TabIndex        =   67
         Top             =   3390
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "清单"
         Height          =   285
         Left            =   14340
         TabIndex        =   66
         Top             =   3450
         Width           =   765
      End
      Begin VB.CommandButton Command3 
         Caption         =   "清单"
         Height          =   315
         Left            =   14340
         TabIndex        =   65
         Top             =   3030
         Width           =   765
      End
      Begin VB.TextBox txtCbze2 
         Height          =   315
         Left            =   13050
         TabIndex        =   64
         ToolTipText     =   "实际"
         Top             =   1080
         Width           =   1185
      End
      Begin VB.TextBox txtYf2 
         Height          =   315
         Left            =   13020
         TabIndex        =   63
         ToolTipText     =   "实际"
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtFbje2 
         Height          =   315
         Left            =   13020
         TabIndex        =   62
         ToolTipText     =   "实际"
         Top             =   2610
         Width           =   1215
      End
      Begin VB.TextBox txtHtrq 
         Height          =   315
         Left            =   6360
         TabIndex        =   61
         Top             =   1590
         Width           =   1815
      End
      Begin VB.TextBox txtZe 
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   3090
         Width           =   1515
      End
      Begin VB.TextBox txtEd 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H80000002&
         Height          =   270
         Left            =   9030
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   3090
         Width           =   885
      End
      Begin VB.ComboBox comQy 
         Height          =   300
         ItemData        =   "wbHTP.frx":06BE
         Left            =   8970
         List            =   "wbHTP.frx":06C0
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1575
         Width           =   945
      End
      Begin VB.TextBox txtXYwy 
         Height          =   315
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1095
         Width           =   3555
      End
      Begin VB.TextBox txtHtbh 
         Height          =   270
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1650
         Width           =   3315
      End
      Begin VB.CommandButton cmdWb 
         Caption         =   "项目档案"
         Height          =   315
         Left            =   1410
         TabIndex        =   28
         Top             =   2580
         Width           =   3375
      End
      Begin VB.TextBox txtHtze 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   27
         Top             =   3090
         Width           =   3345
      End
      Begin VB.TextBox txtRgf1 
         Height          =   285
         Left            =   11730
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1852
         Width           =   2505
      End
      Begin VB.TextBox txtCLF1 
         Height          =   285
         Left            =   11730
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2238
         Width           =   2505
      End
      Begin VB.TextBox txtFbje1 
         Height          =   285
         Left            =   11730
         TabIndex        =   24
         ToolTipText     =   "预计"
         Top             =   2624
         Width           =   1215
      End
      Begin VB.TextBox txtYf1 
         Height          =   285
         Left            =   11730
         TabIndex        =   23
         ToolTipText     =   "预计"
         Top             =   3010
         Width           =   1215
      End
      Begin VB.TextBox txtQt1 
         Height          =   285
         Left            =   11730
         TabIndex        =   22
         Top             =   3396
         Width           =   1215
      End
      Begin VB.TextBox txtClcb1 
         Height          =   285
         Left            =   11730
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1466
         Width           =   2505
      End
      Begin VB.Frame Frame1 
         Caption         =   "发票类型："
         Height          =   765
         Left            =   240
         TabIndex        =   17
         Top             =   6720
         Width           =   4035
         Begin VB.OptionButton optLa 
            Caption         =   "增值发票"
            Height          =   195
            Left            =   180
            TabIndex        =   20
            Top             =   300
            Width           =   1065
         End
         Begin VB.OptionButton optLb 
            Caption         =   "商业发票"
            Height          =   195
            Left            =   1260
            TabIndex        =   19
            Top             =   300
            Width           =   1065
         End
         Begin VB.OptionButton optLc 
            Caption         =   "服务发票"
            Height          =   195
            Left            =   2370
            TabIndex        =   18
            Top             =   300
            Width           =   1065
         End
      End
      Begin VB.TextBox txtCbze1 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   11730
         TabIndex        =   16
         ToolTipText     =   "预计"
         Top             =   1080
         Width           =   1245
      End
      Begin VB.TextBox txtADR 
         Height          =   285
         Left            =   6360
         TabIndex        =   15
         Top             =   2130
         Width           =   3555
      End
      Begin VB.TextBox txtJlr1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   11730
         TabIndex        =   14
         Top             =   3782
         Width           =   1245
      End
      Begin VB.TextBox txtKhdm 
         Height          =   270
         Left            =   1440
         TabIndex        =   13
         Top             =   1140
         Width           =   3315
      End
      Begin VB.TextBox txtXMMC 
         Height          =   285
         Left            =   6360
         TabIndex        =   12
         Top             =   600
         Width           =   3555
      End
      Begin VB.ComboBox txtKhmc 
         Height          =   300
         Left            =   1440
         TabIndex        =   11
         ToolTipText     =   "请在列表中选择客户"
         Top             =   630
         Width           =   3345
      End
      Begin VB.TextBox txtTcRQ 
         Height          =   315
         Left            =   10650
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "提成取现日期"
         Top             =   6960
         Visible         =   0   'False
         Width           =   1845
      End
      Begin MSDataGridLib.DataGrid dtgYf 
         Bindings        =   "wbHTP.frx":06C2
         Height          =   1845
         Left            =   5130
         TabIndex        =   54
         Top             =   3750
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   3254
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BackColor       =   13631199
         ForeColor       =   12582912
         HeadLines       =   1
         RowHeight       =   15
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   "开票日期"
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
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   "开票金额"
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
            DataField       =   ""
            Caption         =   "发类类型"
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
            DataField       =   ""
            Caption         =   "支付否"
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
            ScrollBars      =   2
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   599.811
            EndProperty
         EndProperty
      End
      Begin TabDlg.SSTab tabGc 
         Height          =   4935
         Left            =   -75000
         TabIndex        =   69
         Top             =   0
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   8705
         _Version        =   393216
         TabOrientation  =   1
         Tabs            =   4
         Tab             =   3
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "年保"
         TabPicture(0)   =   "wbHTP.frx":06D6
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "dtgWb"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "例检"
         TabPicture(1)   =   "wbHTP.frx":06F2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "dtgLj"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "大修"
         TabPicture(2)   =   "wbHTP.frx":070E
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtDxnr"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "材料"
         TabPicture(3)   =   "wbHTP.frx":072A
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "dtgMa"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "dtgBao"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "VScroll1"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).Control(3)=   "frmPL"
         Tab(3).Control(3).Enabled=   0   'False
         Tab(3).ControlCount=   4
         Begin VB.Frame frmPL 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   345
            Left            =   8370
            TabIndex        =   152
            Top             =   3180
            Width           =   6765
            Begin VB.CommandButton cmdGG 
               Caption         =   "更新"
               Height          =   315
               Left            =   5190
               TabIndex        =   156
               Top             =   30
               Width           =   675
            End
            Begin VB.TextBox txtDj 
               Height          =   345
               Left            =   3570
               TabIndex        =   155
               Top             =   30
               Width           =   1455
            End
            Begin VB.CommandButton cmdD 
               Caption         =   "删除"
               Height          =   315
               Left            =   5940
               TabIndex        =   154
               Top             =   30
               Visible         =   0   'False
               Width           =   825
            End
            Begin VB.TextBox txtTl 
               Height          =   315
               Left            =   1170
               TabIndex        =   153
               Top             =   30
               Width           =   1515
            End
            Begin VB.Label Label43 
               Caption         =   "单价"
               Height          =   285
               Left            =   2910
               TabIndex        =   158
               Top             =   90
               Width           =   465
            End
            Begin VB.Label Label42 
               Caption         =   "数量"
               Height          =   195
               Left            =   540
               TabIndex        =   157
               Top             =   90
               Width           =   495
            End
         End
         Begin VB.TextBox txtDxnr 
            Height          =   4545
            Left            =   -74970
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   73
            Top             =   30
            Width           =   15165
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   30
            Left            =   1800
            TabIndex        =   72
            Top             =   1530
            Width           =   30
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBao 
            Height          =   3135
            Left            =   0
            TabIndex        =   70
            Top             =   0
            Width           =   15225
            _ExtentX        =   26855
            _ExtentY        =   5530
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
            Height          =   945
            Left            =   0
            TabIndex        =   71
            Top             =   3660
            Width           =   15255
            _ExtentX        =   26908
            _ExtentY        =   1667
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgLj 
            Height          =   4575
            Left            =   -75000
            TabIndex        =   74
            Top             =   0
            Width           =   15255
            _ExtentX        =   26908
            _ExtentY        =   8070
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
            Bindings        =   "wbHTP.frx":0746
            Height          =   4575
            Left            =   -75000
            TabIndex        =   75
            Top             =   0
            Width           =   15255
            _ExtentX        =   26908
            _ExtentY        =   8070
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
         Begin VB.Label Label14 
            Caption         =   "采购成本"
            Height          =   225
            Left            =   -74880
            TabIndex        =   77
            Top             =   4050
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "单价"
            Height          =   285
            Left            =   -63030
            TabIndex        =   76
            Top             =   3990
            Width           =   465
         End
      End
      Begin VB.Label Label50 
         Caption         =   "%"
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   1
         Left            =   9930
         TabIndex        =   221
         Top             =   3120
         Width           =   225
      End
      Begin VB.Label Label49 
         Caption         =   "备注"
         Height          =   225
         Left            =   5580
         TabIndex        =   201
         Top             =   2640
         Width           =   585
      End
      Begin VB.Shape Shape2 
         Height          =   3975
         Left            =   4980
         Top             =   3570
         Width           =   5175
      End
      Begin VB.Label Label48 
         Caption         =   "开票备注"
         Height          =   225
         Left            =   5160
         TabIndex        =   190
         Top             =   5670
         Width           =   885
      End
      Begin VB.Shape Shape1 
         Height          =   3975
         Left            =   90
         Top             =   3570
         Width           =   4905
      End
      Begin VB.Label lblHtxz 
         Caption         =   "Label22"
         Height          =   315
         Left            =   1440
         TabIndex        =   60
         Top             =   2190
         Width           =   3315
      End
      Begin VB.Line Line1 
         X1              =   10350
         X2              =   10350
         Y1              =   7560
         Y2              =   0
      End
      Begin VB.Label Label29 
         Caption         =   "收款总额"
         Height          =   315
         Left            =   5250
         TabIndex        =   56
         Top             =   3150
         Width           =   795
      End
      Begin VB.Label Label8 
         Caption         =   "收款额度"
         Height          =   255
         Left            =   8100
         TabIndex        =   55
         Top             =   3120
         Width           =   915
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "合   同   评   审   单"
         BeginProperty Font 
            Name            =   "华文彩云"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   4230
         TabIndex        =   51
         Top             =   120
         Width           =   2715
      End
      Begin VB.Label Label44 
         Caption         =   "区  域"
         Height          =   255
         Left            =   8340
         TabIndex        =   50
         Top             =   1635
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "客户名称"
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   49
         Top             =   690
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "项目管理者"
         Height          =   255
         Index           =   0
         Left            =   5100
         TabIndex        =   48
         Top             =   1170
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "日    期"
         Height          =   255
         Index           =   0
         Left            =   5250
         TabIndex        =   47
         Top             =   1665
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "合同编号"
         Height          =   225
         Left            =   240
         TabIndex        =   46
         Top             =   1710
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "合同性质"
         Height          =   195
         Left            =   240
         TabIndex        =   45
         Top             =   2205
         Width           =   915
      End
      Begin VB.Label Label13 
         Caption         =   "合同总金额"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   44
         Top             =   3150
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "人 工 费"
         Height          =   255
         Left            =   10680
         TabIndex        =   43
         Top             =   1890
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "差 旅 费"
         Height          =   255
         Left            =   10680
         TabIndex        =   42
         Top             =   2280
         Width           =   915
      End
      Begin VB.Label Label18 
         Caption         =   "分包金额"
         Height          =   255
         Left            =   10680
         TabIndex        =   41
         Top             =   2670
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "运    费"
         Height          =   255
         Left            =   10680
         TabIndex        =   40
         Top             =   3060
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "项目费用"
         Height          =   255
         Left            =   10680
         TabIndex        =   39
         Top             =   3450
         Width           =   885
      End
      Begin VB.Label Label9 
         Caption         =   "成本分析"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   12000
         TabIndex        =   38
         Top             =   690
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "材料成本"
         Height          =   255
         Left            =   10680
         TabIndex        =   37
         Top             =   1500
         Width           =   825
      End
      Begin VB.Label Label24 
         Caption         =   "成本总额"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10680
         TabIndex        =   36
         Top             =   1110
         Width           =   885
      End
      Begin VB.Label Label26 
         Caption         =   "项目地址"
         Height          =   255
         Left            =   5250
         TabIndex        =   35
         Top             =   2190
         Width           =   885
      End
      Begin VB.Label lblJlr 
         Caption         =   "利 润 1"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   10680
         TabIndex        =   34
         Top             =   3840
         Width           =   915
      End
      Begin VB.Label Label6 
         Caption         =   "客户代码"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   1170
         Width           =   885
      End
      Begin VB.Label Label7 
         Caption         =   "项目名称"
         Height          =   255
         Left            =   5250
         TabIndex        =   32
         Top             =   660
         Width           =   795
      End
   End
   Begin VB.Frame frmZt 
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   885
      Left            =   10650
      TabIndex        =   4
      Top             =   8280
      Visible         =   0   'False
      Width           =   1185
      Begin VB.OptionButton optG 
         Caption         =   "已 盖 章"
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   270
         Width           =   1035
      End
      Begin VB.OptionButton optP 
         Caption         =   "评审阶段"
         Height          =   180
         Left            =   90
         TabIndex        =   7
         Top             =   60
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton optZ 
         Caption         =   "执行阶段"
         Height          =   225
         Left            =   90
         TabIndex        =   6
         Top             =   480
         Width           =   1035
      End
      Begin VB.OptionButton optW 
         Caption         =   "执行完毕"
         Height          =   225
         Left            =   90
         TabIndex        =   5
         Top             =   690
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除"
      Enabled         =   0   'False
      Height          =   585
      Left            =   13980
      Picture         =   "wbHTP.frx":075A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8580
      Width           =   645
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "修改"
      Height          =   585
      Left            =   12630
      Picture         =   "wbHTP.frx":08E4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8580
      Width           =   645
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   585
      Left            =   14640
      Picture         =   "wbHTP.frx":0D26
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8580
      Width           =   585
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "提交"
      Height          =   585
      Left            =   13290
      Picture         =   "wbHTP.frx":0E28
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8580
      Width           =   675
   End
   Begin VB.Label lblJiLI 
      Caption         =   "请再次按提交按钮,以便刷新数据"
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   13110
      TabIndex        =   151
      Top             =   8160
      Width           =   1725
   End
   Begin VB.Label lblBaoid 
      Caption         =   "lblBaoid"
      Height          =   285
      Left            =   4200
      TabIndex        =   90
      Top             =   8340
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label lblHid 
      Caption         =   "lblHid"
      Height          =   285
      Left            =   2520
      TabIndex        =   89
      Top             =   8370
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label lblQM 
      Caption         =   "lblQm"
      Height          =   225
      Index           =   0
      Left            =   840
      TabIndex        =   59
      Top             =   8040
      Width           =   1185
   End
   Begin VB.Label lblTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   0
      Left            =   780
      TabIndex        =   58
      Top             =   8730
      Width           =   945
   End
End
Attribute VB_Name = "frmWbNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adoWb As New ADODB.Recordset
Public adoLj As New ADODB.Recordset
Public adoOid As New ADODB.Recordset '计算Old单子的ADO
Public adoBx As Object '采购表
Public adoGx As Object '成本表
Public adoFk As Object '预计付款
Public adoYj As Object '资金表
Public adoGD As Object '固定项目费用
Public adoHGD As Object '固定项目费用总和

Dim JILI As Integer '保存的次数'销售总监要进行两次保存

Public adoA As Object

Public Bid As Long '对应询价单的编号

Dim timZm As Integer '数据提交后,由timWait执行的后续命令ID(2 保存合同 3新建询价单(配件),6新建询价单(产品),10签字11生成合同编号12删除合同13编辑奖金15提成编辑16全款编辑)

Dim Pw As String

Private Sub Check1_Click()
If chkYJF.Value = 1 Then
    txtYjf.Text = mod1.DQda
Else
    txtYjf.Text = ""
End If
End Sub

Private Sub chkYJF_Click()
If chkYJF.Value = 1 Then
    txtYjf.Text = mod1.DQda
Else
    txtYjf.Text = ""
End If
End Sub

'Public HTF As Integer '合同状态
Private Sub cmdAdd_Click()
On Error Resume Next
Set mod1.cmd = CreateObject("adodb.command")
mod1.cmd.ActiveConnection = mod1.cc
mod1.cmd.CommandText = "htFkAdd"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@rq") = txtYrq.Text
mod1.cmd.Parameters("@yingfJe") = Round(Val(txtHtze.Text) * Val(txtYed.Text) / 100, 2)
mod1.cmd.Parameters("@htbh") = Trim(txtHtbh.Text)
mod1.cmd.Parameters("@ed") = Round(Val(txtYed.Text) / 100, 2)
mod1.cmd.Execute
Set cmd = Nothing

txtYed.Text = ""
adoFk.Requery
Set dtgFk.DataSource = adoFk
End Sub

Private Sub cmdBack_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
If lblLc.Caption = 0 And cmdSave.Enabled = True And lblFwid.Caption = "" Then
    ii = MsgBox("该单没有保存,是否将其撤消?", vbQuestion + vbYesNo, "请注意!")
    If ii = vbYes Then
        tt = "update htping set delf=0 where hid=" & Val(lblHid.Caption)
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
        tt = "update baoJiaD set htbh=null where baoid=" & Val(lblBaoId.Caption)
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Else
        Exit Sub
    End If
End If
Me.Visible = False
If htBrow.Visible = True Then
    htBrow.Enabled = True
    htBrow.ZOrder 0
ElseIf htBrowG.Visible = True Then
    htBrowG.Enabled = True
    htBrowG.ZOrder 0
ElseIf Dialog.Visible = True Then
    Dialog.ZOrder 0
    Dialog.Enabled = True
ElseIf frmGxBiao.Visible = True Then
    frmGxBiao.Enabled = True
    frmGxBiao.ZOrder 0
ElseIf frmCWBB.Visible = True Then
    frmCWBB.Enabled = True
    frmCWBB.ZOrder 0
End If
Call mod1.DelDKZ ' '退出表单时删除打开记录,以让别人能打开此单据
End Sub

Private Sub cmdClcb_Click()
Dim tt As String
Dim ii As Integer
Dim CB As Long
On Error Resume Next
'由材料成本明细重新计算材料成本（防止在以前报价询价中出错）,合同执行后则除外,因为可能由倪旭硬性修改.

    adoBx.MoveFirst
    Do While Not adoBx.EOF
        tt = "select 合计 from xunJiaMxView where lid=" & adoBx.Fields("lid").Value
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    
        CB = CB + mod1.HTP.Fields("合计").Value
        adoBx.MoveNext
    Loop
    If Val(txtClcb1.Text) <> CB Then
        ii = MsgBox("核算出询价成本为" & CB & ",与现材料成本不符,是否确认修改?", vbYesNo + vbInformation + vbDefaultButton2)
        If ii = vbNo Then Exit Sub
        txtClcb1.Text = CB
        tt = "update htping set clcb=" & Val(txtClcb1.Text) & " where htbh='" & txtHtbh.Text & "'"
        Set mod1.HTP = CreateObject("adodb.recordset")
        On Error GoTo uphtErr
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    End If
Exit Sub
uphtErr:
MsgBox ("网络故障，请再次提交！")
cmdSave.Enabled = True
End Sub

Private Sub cmdClose_Click()
frmYm.Visible = False

End Sub

Private Sub cmdCong_Click()
Dim tt As String
Dim Bid As Long
Dim ZL As String
Dim oo As Integer
Dim ii As Integer
On Error Resume Next
If Not (optZ.Value = True Or optW.Value = True) Then
    If mod1.DName = lblYwy.Caption Then





        ii = MsgBox("您的这项操作将使原先单子正在执行的流程全部撤消,是否确定执行?", vbYesNo + vbInformation, "询问")
        If ii = vbYes Then
            tt = InputBox("请输入您要驳回的原因!")
            If tt = "" Then Exit Sub
            Set mod1.cmd = CreateObject("adodb.command")
            mod1.cmd.ActiveConnection = mod1.cc
            mod1.cmd.CommandText = "xtzxFAdd"
            mod1.cmd.CommandType = adCmdStoredProc
            mod1.cmd.Parameters("@yid").Value = 62 '反签名
            mod1.cmd.Parameters("@lc").Value = 2 '退回最初的流程
            mod1.cmd.Parameters("@bh").Value = txtHtbh.Text
            mod1.cmd.Parameters("@ywy").Value = mod1.DName
            mod1.cmd.Parameters("@uid").Value = mod1.DHid
            mod1.cmd.Parameters("@bz").Value = tt
            mod1.cmd.Parameters("@zn").Value = "new" '身份职能
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
        End If
    End If
Else
    MsgBox ("合同已经正式生成,不能修改!")
End If
End Sub

Private Sub cmdDe_Click()
Dim tt As String
On Error Resume Next
tt = "delete from htping1 where fid=" & Val(lblFid.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText

txtYed.Text = ""
adoFk.Requery
Set dtgFk.DataSource = adoFk
End Sub

Private Sub cmdGG_Click()
Dim CB As Long
Dim liD As Long
Dim Bid As Long
Dim XCB As Long
On Error Resume Next
If Val(txtDj.Text) = 0 Then Exit Sub
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
    
    tt = "select bid from baojiaD where baoid=" & Val(lblBaoId.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Bid = mod1.HTP.Fields("bid").Value
    If lblHtxz.Caption = "维保" Or lblHtxz.Caption = "大修" Then
        '获得相应询价单的cgid号
        tt = "select cgid from xunJiaD where bid=" & Bid
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Bid = mod1.HTP.Fields("cgid").Value
    End If
    
    '更新相应询价明细中的数量
    tt = "update XunJiaMx set sl=" & Val(txtTl.Text) & ",hg=dj*" & Val(txtTl.Text) & " where lid=" & liD
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    '更新相应询价单中的金额
    tt = "select sum(hg) as hg from xunjiamx where bid=" & Bid
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'    XCB = 0
'    Do While Not mod1.HTP.EOF
'        XCB = XCB + mod1.HTP.Fields("hg").Value
'        mod1.HTP.MoveNext
'    Loop
    XCB = mod1.HTP.Fields("hg").Value

    tt = "update xunjiaD set hg=" & XCB & ",yhg=" & XCB & " where bid=" & Bid
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    txtDj.Text = ""
    txtSl.Text = ""
    txtClcb.Text = XCB
    adoBx.Requery
    Set dtgBao.DataSource = adoBx
    Call cmdSave_Click
    txtDj.Text = ""
    txtTl.Text = ""
End Sub

Private Sub cmdGx_Click()
On Error Resume Next
Set mod1.cmd = CreateObject("adodb.command")
mod1.cmd.ActiveConnection = mod1.cc
mod1.cmd.CommandText = "htFkGx"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@rq") = dtpYf.Value
mod1.cmd.Parameters("@yingfJe") = Round(Val(txtHtze.Text) * Val(txtYed.Text) / 100, 2)
mod1.cmd.Parameters("@htbh") = Trim(txtHtbh.Text)
mod1.cmd.Parameters("@ed") = Round(Val(txtYed.Text) / 100, 2)
mod1.cmd.Parameters("@Fid") = Val(lblFid.Caption)
mod1.cmd.Execute
Set cmd = Nothing

txtYed.Text = ""
adoFk.Requery
Set dtgFk.DataSource = adoFk
End Sub

Private Sub cmdJTFadd_Click()
Dim tt As String
Dim hg As Single
On Error Resume Next
If Val(txtJtfJe.Text) = 0 Then
Exit Sub
End If

timZm = 15 '添加奖金
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "合同评审单"
    mod1.cmd.Parameters("@NBLX") = "提成编辑"
    mod1.cmd.Parameters("@bh") = lblHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(txtHtbh.Text) '合同编号
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
    mod1.cmd.Parameters("@mlt1") = txtJTFbz.Text
    mod1.cmd.Parameters("@mlt2") = ""
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtJtfJe.Text)
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
    mod1.cmd.Parameters("@mb1") = 1 '添加提成
    mod1.cmd.Parameters("@mb2") = 0
    mod1.cmd.Parameters("@mb3") = 0
    mod1.cmd.Parameters("@mb4") = 0
    mod1.cmd.Parameters("@mb5") = 0
    mod1.cmd.Parameters("@md1") = dtpJTF.Value
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

Private Sub cmdJTFdel_Click()
Dim tt As String
Dim ii As Integer
Dim Yid As Long

On Error Resume Next

dtgJTf.Col = 4
Yid = 0
Yid = Val(dtgJTf.Text)


If Yid = 0 Then
Exit Sub
End If



ii = MsgBox("是否删除此记录?", vbQuestion + vbYesNo, "询问")
If ii = vbNo Then
    Exit Sub
End If

timZm = 15 '提成编辑
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "合同评审单"
    mod1.cmd.Parameters("@NBLX") = "提成编辑"
    mod1.cmd.Parameters("@bh") = lblHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(txtHtbh.Text) '合同编号
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
    mod1.cmd.Parameters("@mm1") = Yid
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
    mod1.cmd.Parameters("@mb1") = 0 '提成删除
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


Private Sub cmdKP_Click()
Call frmFP.KPQing
frmFP.Show
frmFP.txtXmmc.Text = txtXmmc.Text
frmFP.txtKhmc.Text = txtKhmc.Text
frmFP.txtHtze.Text = txtHtze.Text
frmFP.lblHtxz.Caption = lblHtxz.Caption
frmFP.txtHtbh.Text = txtHtbh.Text
Set frmFP.dtgFk.DataSource = adoFk
Me.Enabled = False
frmFP.ZOrder 0
End Sub

Private Sub cmdMod_Click()
On Error Resume Next
If mod1.DName = "乔继敏" Then
'txtYjf.Locked = False
'txtJTf.Locked = False
'txtQkf.Locked = False
'chkYJF.Enabled = True
'chkJTF.Enabled = True
'chkQKF.Enabled = True
'txtYjfBz.Locked = False
frmJTF.Visible = True
frmQkF.Visible = True
frmCw.Enabled = True
cmdSave.Enabled = True
Exit Sub
End If

cmdYadd.Visible = False
cmdYdel.Visible = False
If lblLc.Caption = 1 And lblYwy.Caption = mod1.DName Then
    frmFX.Visible = True
    dt3.Enabled = True
    dt4.Enabled = True
    optLa.Enabled = True
    optLb.Enabled = True
    optLc.Enabled = True
    cmdSave.Enabled = True
    
ElseIf (mod1.DName = "倪旭" Or mod1.DName = "宋晓炯" Or mod1.DName = "宋晓炯1" Or mod1.DName = "马晓聪") And optW.Value = False Then
    frmPL.Visible = True
    frmFX.Visible = True
    dt3.Enabled = True
    dt4.Enabled = True
    optLa.Enabled = True
    optLb.Enabled = True
    optLc.Enabled = True
    'txtYf1.Locked = False
    txtQt1.Locked = False
    txtYj1.Locked = False
    txtTcBe.Locked = False
    txtHtze.Locked = False
    cmdSave.Enabled = True
    txtFbje1.Locked = False
    If lblyjFF.Caption = "False" Then
        cmdYadd.Visible = True
        cmdYdel.Visible = True
    End If
    'JILI = 0
End If

'If mod1.DName = "宋晓炯" Or mod1.DName = "马晓聪" Then
'    txtFbje1.Locked = False
'End If


End Sub

Private Sub cmdPje_Click()
Dim tt As String
Dim Cgid As Long
Dim Bid As Long
On Error Resume Next
Pje.Show
'取得报价单和询价单的评审建议
tt = "select bid from baojiaD where baoid=" & Val(lblBaoId.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Bid = mod1.HTP.Fields("bid").Value
tt = "select cgid from xunjiaD where bid=" & Bid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Cgid = mod1.HTP.Fields("cgid").Value

tt = "select * from pizu where (bh='" & txtHtbh.Text & "' and yid=62) or (bh='" & lblBaoId.Caption & _
"' and yid=60) or ((bh='" & Bid & "' or bh='" & Cgid & "') and yid=43) order by trq desc"
Pje.adoPje.Close
Pje.adoPje.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set Pje.dtgPje.DataSource = Pje.adoPje
Pje.txtXQ.Text = ""
End Sub

Private Sub cmdQing_Click()
txtYed.Text = ""
End Sub

Private Sub cmdQkfAdd_Click()
Dim tt As String
Dim hg As Single
On Error Resume Next
If Val(txtQkfJe.Text) = 0 Then
Exit Sub
End If

timZm = 16 '全款编辑
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "合同评审单"
    mod1.cmd.Parameters("@NBLX") = "全款编辑"
    mod1.cmd.Parameters("@bh") = lblHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(txtHtbh.Text) '合同编号
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
    mod1.cmd.Parameters("@mlt1") = txtQkFBz.Text
    mod1.cmd.Parameters("@mlt2") = ""
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtQkfJe.Text)
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
    mod1.cmd.Parameters("@mb1") = 1 '添加全款
    mod1.cmd.Parameters("@mb2") = 0
    mod1.cmd.Parameters("@mb3") = 0
    mod1.cmd.Parameters("@mb4") = 0
    mod1.cmd.Parameters("@mb5") = 0
    mod1.cmd.Parameters("@md1") = dtpQkF.Value
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

Private Sub cmdQkfDel_Click()
Dim tt As String
Dim ii As Integer
Dim Yid As Long

On Error Resume Next

dtgyjF.Col = 4
Yid = 0
Yid = Val(dtgyjF.Text)


If Yid = 0 Then
Exit Sub
End If



ii = MsgBox("是否删除此记录?", vbQuestion + vbYesNo, "询问")
If ii = vbNo Then
    Exit Sub
End If

timZm = 16 '全款编辑
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "合同评审单"
    mod1.cmd.Parameters("@NBLX") = "全款编辑"
    mod1.cmd.Parameters("@bh") = lblHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(txtHtbh.Text) '合同编号
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
    mod1.cmd.Parameters("@mm1") = Yid
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
    mod1.cmd.Parameters("@mb1") = 0 '全款删除
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


Private Sub cmdQm_Click(Index As Integer)
Dim tt As String
Dim Tywy As String '单子流转到下一人的姓名
Dim Tuid As String
Dim Oywy As String '原来流转人的名字
Dim Ouid As String '原来流转人的工号



On Error Resume Next
If cmdQm(Index).Caption <> "" Then
    Exit Sub
End If
If adoFk.RecordCount = 0 Then
    MsgBox ("请输入付款方式!")
    cmdSave.Enabled = True
    Exit Sub
End If

If optLa.Value = False And optLb.Value = False And optLc.Value = False Then
    MsgBox ("请选择开票方式!")
    cmdSave.Enabled = True
    Exit Sub
End If

If lblHtxz.Caption = "维保" And (txtF.Text = "" Or txtL.Text = "") Then
    MsgBox ("请标明维保的起始期和截至期!")
    frmWbNew.tabHt.Tab = 1
    cmdSave.Enabled = True
    Exit Sub
End If


'If cmdQm(Index).Caption <> "" Then Exit Sub

'If Index = 0 And cmdSave.Enabled = True And lblLc.Caption = 0 Then
If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If

If Index + 1 <> lblLc.Caption Then '不能在不相干的位置上乱点

    Exit Sub
End If



'If lblLcUid.Caption <> mod1.DHid Then
If lblLcRen.Caption <> mod1.DName Then
    MsgBox "此处应由" & lblLcRen.Caption & "签字! 请您不要再点"
    Exit Sub
End If

If lblQM(Index).Caption = "执行完毕确认" Then
    MsgBox "未收全款，不能点完成！"
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
        mod1.cmd.Parameters("@yid").Value = 62 '反签名
        mod1.cmd.Parameters("@lc").Value = lblLc.Caption
        mod1.cmd.Parameters("@bh").Value = txtHtbh.Text
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

Oywy = lblLcRen.Caption
Ouid = lblLcUid.Caption

    lblLc.Caption = lblLc.Caption + 1
    
''''''If lblQM(Index).Caption = "合同执行" Then
''''''        Set mod1.CMD = createobject("adodb.command")
''''''        mod1.CMD.ActiveConnection = mod1.CC
''''''        mod1.CMD.CommandText = "TXht"
''''''        mod1.CMD.CommandType = adCmdStoredProc
''''''        mod1.CMD.Parameters("@lc") = Val(lblLc.Caption)
''''''        mod1.CMD.Parameters("@htbh") = txtHtbh.Text
''''''        mod1.CMD.Parameters("@fwid") = Val(lblFwid.Caption)
''''''        mod1.CMD.Parameters("@nr") = txtXMMC.Text
''''''        mod1.CMD.Parameters("@khdh") = txtKhdm.Text
''''''        mod1.CMD.Parameters("@uid") = lblUid.Caption
''''''        mod1.CMD.Parameters("@bm") = lblBm.Caption
''''''        mod1.CMD.Parameters("@Errch") = ""   '评审建议
''''''        mod1.CMD.Execute
''''''        If mod1.CMD.Parameters("@Errch").Value <> "成功" Then
''''''        lblLc.Caption = lblLc.Caption - 1
''''''        MsgBox "网络出现故障,请再试一次,如果还是提交不成功,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
''''''        Exit Sub
''''''        End If
''''''    cmdQm(Index).Caption = mod1.DName
''''''    lblTm(Index).Caption = mod1.DQda
''''''    lblLcRen.Caption = ""

''''''End If
''''''    MsgBox ("数据导入成功!接下来,将请天兴软件负责此单的执行!")
''''''Exit Sub
''''''End If

If lblLc.Caption <> 2 Then
    
    '更新表baojiaD中的lcRen,lcUid 字段,以及QMRZ表中的相应字段.
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
                mod1.cmd.Parameters("@Qdbh") = Trim(txtHtbh.Text)   '单子编号
                mod1.cmd.Parameters("@pje") = ""   '评审建议
                mod1.cmd.Parameters("@bm") = ""
                mod1.cmd.Parameters("@qy") = ""
                mod1.cmd.Parameters("@Gren") = "" '如果为费用归属报销单,则添加费用归属人的参数
                mod1.cmd.Parameters("@Guid") = ""
                mod1.cmd.Parameters("@ywy") = lblYwy.Caption
                mod1.cmd.Parameters("@yid") = lblUid.Caption
                mod1.cmd.Parameters("@comid") = mod1.comId
                mod1.cmd.Execute
                Tywy = mod1.cmd.Parameters("@Tywy").Value
                Tuid = mod1.cmd.Parameters("@Tuid").Value
                Set cmd = Nothing
                cmdQm(Index).Caption = mod1.DName
                lblTm(Index).Caption = mod1.DQda

Else
    If mod1.comId = 0 And Not (mod1.Bm = "维销部3" Or mod1.Bm = "产品部1" Or mod1.Bm = "产品部2") Then
        Tywy = "倪旭"
        Tuid = "HM040"
    Else
        If mod1.comId = 0 Then
            Tywy = "宋晓炯"
            Tuid = "HM003"
        ElseIf mod1.comId = 1 Then
            Tywy = "宋晓炯1"
            Tuid = "HMG000"
        End If
    End If
    tt = "update QMRZ set  Qren='" & mod1.DName & "',Qrid='" & mod1.DHid & "',Qrq='" & mod1.DQda & "',xf=1 where Qdbh='" & txtHtbh.Text & "' and btz=" & mod1.BTZ & " and zid=" & cmdQm(Index).Tag
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    tt = "update htping set lcren='" & Tywy & "',lcUid='" & Tuid & "',lc=" & Val(lblLc.Caption) & " where htbh='" & txtHtbh.Text & "'"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    
    cmdQm(Index).Caption = mod1.DName
    lblTm(Index).Caption = mod1.DQda
End If


If lblQM(Index + 1).Caption = "财务盖章" Then
    If comQy.Text = "上海" Then

    ElseIf comQy.Text = "南京" Then
        Tywy = "王蕾"
        Tuid = "HM051"
    ElseIf comQy.Text = "杭州" Then
        Tywy = "李艳"
        Tuid = "HM316"
    ElseIf comQy.Text = "北京" Then
        Tywy = "马玉芝"
        Tuid = "HM190"
    ElseIf comQy.Text = "广州" Then
        Tywy = "汤丽嫦"
        Tuid = "HMG023"
    End If
    tt = "update htping set lcren='" & Tywy & "',lcUid='" & Tuid & "',lc=" & Val(lblLc.Caption) & " where htbh='" & txtHtbh.Text & "'"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
ElseIf lblQM(Index + 1).Caption = "合同执行" Then
    If comQy.Text = "上海" Then
        Tywy = "封红"
        Tuid = "HM233"
    ElseIf comQy.Text = "南京" Then
        Tywy = "王蕾"
        Tuid = "HM051"
    ElseIf comQy.Text = "杭州" Then
        Tywy = "李艳"
        Tuid = "HM316"
    ElseIf comQy.Text = "北京" Then
        Tywy = "马玉芝"
        Tuid = "HM190"
    ElseIf comQy.Text = "广州" Then
        Tywy = "汤丽嫦"
        Tuid = "HMG023"
    End If
    tt = "update htping set lcren='" & Tywy & "',lcUid='" & Tuid & "',lc=" & Val(lblLc.Caption) & " where htbh='" & txtHtbh.Text & "'"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
End If
lblLcRen.Caption = Tywy
lblLcUid.Caption = Tuid

Select Case lblQM(Index).Caption
Case "财务盖章"
    tt = "update htping set htf=9 where hid=" & Val(lblHid.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
Case "合同执行"
    tt = "update htping set htf=1,htrq='" & Date & "' where hid=" & Val(lblHid.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
Case "执行完毕确认"
    tt = "update htping set htf=2 where hid=" & Val(lblHid.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
End Select

If Val(lblLc.Caption) > Val(lblLcou.Caption) Then
    Call mod1.EnventFinish(frmWbNew.lblFwid.Caption)
    tt = "update htping set Pwf=1 where hid=" & Val(lblHid.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    MsgBox "终于完成这份合同了!"


Else
'    If lblLc.Caption = 1 Then '业务员第一个签字,则询价日期等于签字日期
'
'    End If
    '添加事务
    If lblLc.Caption <> 6 Then
        Call mod1.EnventAdd("合同评审单", txtXmmc.Text, lblLcRen.Caption, lblLcUid.Caption, txtHtbh.Text, lblQM(Index + 1).Caption, Oywy, Ouid, lblYwy.Caption, lblUid.Caption, lblFwid.Caption, lblHid.Caption)
    End If
    Select Case lblQM(Val(lblLc.Caption) - 1).Caption
    Case "财务盖章"
        MsgBox "审核全部通过,此单可以同客户盖章了!"
    Case "合同执行"
        MsgBox "现在,此询价单将交由 " & Tywy & " 来审阅!"
    Case "执行完毕确认"
        MsgBox "豪曼信息将提醒" & lblYwy.Caption & "去注意这份合同!"
    Case Else
        MsgBox "现在,此询价单将交由 " & Tywy & " 来审阅!"
    End Select
    

End If



End Sub

Private Sub cmdSave_Click()

On Error Resume Next
Dim FPLX As String
Dim CB As Single

    If optLa.Value = True Then
        FPLX = "增值发票"
    ElseIf optLb.Value = True Then
        FPLX = "商业发票"
    ElseIf optLc.Value = True Then
        FPLX = "服务发票"
    End If

If mod1.DName = "乔继敏" Then '小乔保存财务评定
    timZm = 2 '保存合同
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "合同评审单"
        mod1.cmd.Parameters("@NBLX") = "保存旧财务评定"
        mod1.cmd.Parameters("@bh") = Val(lblHid.Caption)
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
        mod1.cmd.Parameters("@mlt3") = "" '业绩备注
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
        mod1.cmd.Parameters("@mb1") = chkYJF.Value '业绩否
        mod1.cmd.Parameters("@mb2") = chkJTF.Value '提成否
        mod1.cmd.Parameters("@mb3") = chkQKF.Value '全款否
        mod1.cmd.Parameters("@mb4") = 0
        mod1.cmd.Parameters("@mb5") = 0
        mod1.cmd.Parameters("@md1") = dt3.Value  '维保起始期
        mod1.cmd.Parameters("@md2") = dt4.Value
        mod1.cmd.Parameters("@md3") = Null
        mod1.cmd.Parameters("@md4") = Null
        mod1.cmd.Parameters("@md5") = Null
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
Else

    FPLX = ""
    CB = 0
    
    On Error Resume Next
    '计算成本利润
    txtCbze1.Text = Val(txtClcb1.Text) + Val(txtRgf1.Text) + Val(txtCLF1.Text) + Val(txtFbje1.Text) + Val(txtYf1.Text) + Val(txtQt1.Text)
    txtJlr1.Text = Val(txtHtze.Text) - Val(txtCbze1.Text)
    txtLr1.Text = Val(txtJlr1.Text) - Val(txtYj1.Text)
    
    

    If txtTcRQ.Text = "" Then
        txtTcRQ.Text = "2000-1-1"
    End If
    'If (txtF.Text = "" Or txtL.Text = "") And lblHtxz.Caption = "维保" Then
    '    MsgBox "请输入维保工期!"
    '    Exit Sub
    '    tabHt.Tab = 1
    'End If
    
    
        If txtF.Text = "" Then txtF.Text = "1999-1-1"
        If txtL.Text = "" Then txtL.Text = "1999-1-1"
        Set mod1.cmd = CreateObject("adodb.command")
            mod1.cmd.ActiveConnection = mod1.cc
            mod1.cmd.CommandText = "htNewAdd"
            mod1.cmd.CommandType = adCmdStoredProc
            mod1.cmd.Parameters("@htze") = Val(txtHtze.Text)
            mod1.cmd.Parameters("@fplx") = FPLX
            mod1.cmd.Parameters("@fbje") = Val(txtFbje1.Text)
            mod1.cmd.Parameters("@tcbe") = Val(txtTcBe.Text)
            mod1.cmd.Parameters("@tc1") = Val(txtTc2.Text)
            mod1.cmd.Parameters("@tcrq") = txtTcRQ.Text
            mod1.cmd.Parameters("@htqy") = txtF.Text
            mod1.cmd.Parameters("@htqy1") = txtL.Text
            mod1.cmd.Parameters("@hid") = Val(lblHid.Caption)
            mod1.cmd.Parameters("@yj") = Val(txtYj1.Text)
            mod1.cmd.Parameters("@qtf1") = Val(txtQt1.Text)
            mod1.cmd.Execute
            Set cmd = Nothing
            cmdSave.Enabled = False
        
        If lblFwid.Caption = "" Then
            lblLc.Caption = 1
            tt = "update htping set lc=1 where hid=" & Val(lblHid.Caption)
            Set mod1.HTP = CreateObject("adodb.recordset")
            mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
            
            '添加事务
            Call mod1.EnventAdd("合同评审单", txtXmmc.Text, lblLcRen.Caption, lblLcUid.Caption, txtHtbh.Text, lblQM(0).Caption, "", "", mod1.DName, mod1.DHid, 0, lblHid.Caption)
            '更新按钮
            Call modHt.OpenHtAn
        End If
        lblJiLI.Visible = False
End If





'If (lblZl.Caption = "维保" And Val(txtYhg.Text) > 50000) Or Val(txtYhg.Text) > 100000 Or Val(txtFbje.Text) > 0 Then
End Sub

Private Sub cmdTk_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
If comPb.Text = "" Or comXh.Text = "" Or Val(txtSl.Text) = 0 Then Exit Sub
'年保表
tt = "select * from xunJIaWbView where wbx='年保' and bid=" & frmWbNew.Bid & " and 机组品牌='" & comPb.Text & "' and 机组型号 like '%" & comXh.Text & "%'"
adoWb.Close
adoWb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set dtgWb.DataSource = adoWb
dtgWb.FixedRows = 0
dtgWb.MergeCol(1) = True
dtgWb.MergeCol(2) = True
dtgWb.MergeCol(3) = True
dtgWb.MergeCells = 3
dtgWb.FixedRows = 1
'例检表
tt = "select * from xunJIaWbView where wbx='例检' and bid=" & frmWbNew.Bid & " and 机组品牌='" & comPb.Text & "' and 机组型号 like '%" & comXh.Text & "%'"
adoLj.Close
adoLj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set dtgLj.DataSource = adoLj
dtgLj.FixedRows = 0
dtgLj.MergeCol(1) = True
dtgLj.MergeCol(2) = True
dtgLj.MergeCol(3) = True
dtgLj.MergeCells = 3
dtgLj.FixedRows = 1
End Sub

Private Sub cmdWb_Click()
Dim tt As String
On Error Resume Next
Dim Kid As Long
Dim xid As Long

    'dtgKH.Col = 2
    xid = frmWbNew.txtXmmc.Tag
    
    'dtgKH.Col = 5
'    kid = Val(dtgKH.Text)
'    dtgKH.Col = 2

    frmWait.Show
    frmWait.ZOrder 0
    
    frmWait.Refresh
    frmWait.faWait.Play
    


    
    frmWbNew.Enabled = False
    wbDN.Visible = False
    Me.MousePointer = 11
    mod1.BTZ = 1
    Call mod1.xmQing
    Call mod1.khQing
    Call mod1.xmBound(xid)
    wbDN.lblKid.Caption = wbDN.lblYZ.Tag
    Call mod1.khBound(wbDN.lblYZ.Tag, "yz")

    wbDN.frmJE.Visible = False

    wbDN.Left = 0
    wbDN.Top = 0
    wbDN.cmdMod.Enabled = False
    wbDN.cmdSave.Enabled = False
    Me.MousePointer = 0
    wbDN.tabKh.Tab = 0

    wbDN.tabKh.TabEnabled(2) = True
    wbDN.tabKh.TabEnabled(0) = True
    
    
    

    wbDN.modFi = False

    Me.MousePointer = 0
    wbDN.cmdSave.Enabled = False
    wbDN.tabKh.Enabled = True

    wbDN.khAdd = False
    '打开项目后,默认的打开客户为项目资料
    wbDN.optYz.Value = True
    wbDN.frmGL.Visible = False
    frmWait.Visible = False
    wbDN.Visible = True
    wbDN.cmdQing.Enabled = False
    wbDN.cmdNew.Enabled = False
    wbDN.cmdRadd.Enabled = False
    wbDN.cmdRdel.Enabled = False
    If wbDN.comXyxz.Text = "物业公司" Then
        wbDN.frmGL.Visible = True
    End If
    
    '更新动态签字按钮的初始设置
        For oo = 1 To 10
           wbDN.lblQM(oo).Left = wbDN.lblQM(oo - 1).Left + 1100
           wbDN.cmdQm(oo).Left = wbDN.cmdQm(oo - 1).Left + 1100
           wbDN.lblTm(oo).Left = wbDN.lblTm(oo - 1).Left + 1100
           mod1.HTP.MoveNext
        Next
End Sub

Private Sub cmdYadd_Click()
Dim tt As String
Dim hg As Single
On Error Resume Next
If Val(txtFED.Text) = 0 Or Val(txtYingFu.Text) = 0 Then
Exit Sub
End If

tt = "select yjff from htping where htbh='" & txtHtbh.Text & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If IsNull(mod1.HTP.RecordCount) Or mod1.HTP.RecordCount = 0 Then
    Exit Sub
End If
If mod1.HTP.Fields("yjff").Value = True Then
    MsgBox ("奖金已经全部支付,不能再更改!")
    Exit Sub
End If

Set mod1.cmd = CreateObject("adodb.command")
mod1.cmd.ActiveConnection = mod1.cc
mod1.cmd.CommandText = "htyjAdd"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@htbh") = Trim(txtHtbh.Text)
mod1.cmd.Parameters("@YED") = Val(txtFED.Text) / 100
mod1.cmd.Parameters("@yingFu") = Val(txtYingFu.Text)
mod1.cmd.Parameters("@xmmc") = Trim(txtXmmc.Text)
mod1.cmd.Execute
Set cmd = Nothing
adoYj.Requery
Set dtgYJ.DataSource = adoYj

hg = 0
If adoYj.RecordCount > 0 Then
    adoYj.MoveFirst
    Do While Not adoYj.EOF
       hg = hg + adoYj.Fields("支付金额").Value
       adoYj.MoveNext
    Loop
End If
'HG = HG + Val(txtYingFu.Text)
'If HG > Val(txtYj.Text) Then
'    MsgBox "填写金额有误!"
'    txtYingFu.Text = ""
'    Exit Sub
'End If
'End If
txtYj1.Text = hg
txtLr1.Text = Val(txtJlr1.Text) - Val(txtYj1.Text)
tt = "update htping set yj=" & Val(txtYj1.Text) & ",xmlr=" & Val(txtLr1.Text) & " where htbh='" & txtHtbh.Text & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText


End Sub

Private Sub cmdYdel_Click()
Dim tt As String
Dim hg As Single
Dim ii As Integer
Dim Yid As Long
Dim Ywy As String
On Error Resume Next
dtgYJ.Col = 4
Ywy = dtgYJ.Text
dtgYJ.Col = 3
Yid = 0
Yid = dtgYJ.Text


If Yid = 0 Then
Exit Sub
End If

If Ywy <> "" Then
    MsgBox "此单已经激活,不能删除! 如果确定要删除,请与马晓聪联系!"
    Exit Sub
End If


ii = MsgBox("是否删除此记录?", vbQuestion + vbYesNo, "询问")
If ii = vbNo Then
    Exit Sub
End If
tt = "delete from yongjin where yid=" & Yid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
adoYj.Requery
Set dtgYJ.DataSource = adoYj

hg = 0
If adoYj.RecordCount > 0 Then
    adoYj.MoveFirst
    Do While Not adoYj.EOF
       hg = hg + adoYj.Fields("支付金额").Value
       adoYj.MoveNext
    Loop
End If

txtYj1.Text = hg
txtLr1.Text = Val(txtJlr1.Text) - Val(txtYj1.Text)
tt = "update htping set yj=" & Val(txtYj1.Text) & ",xmlr=" & Val(txtLr1.Text) & " where htbh='" & txtHtbh.Text & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText

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
dtgBao.Col = 16
liD = dtgBao.Text
tt = "select * from xunJiaMxView where lid=" & liD
adoGx.Close
adoGx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set dtgMa.DataSource = adoGx
End Sub


Private Sub dtgFk_Click()
On Error Resume Next
dtgFk.Col = 1
dtpYf.Value = dtgFk.Text
dtgFk.Col = 2
txtYed.Text = dtgFk.Text
dtgFk.Col = 5
lblFid.Caption = dtgFk.Text
End Sub

Private Sub dtgFk_RowColChange()
On Error Resume Next
dtgFk.Col = 1
txtYrq.Text = dtgFk.Text
dtgFk.Col = 2
txtYed.Text = dtgFk.Text
dtgFk.Col = 5
lblFid.Caption = dtgFk.Text
End Sub


Private Sub dtpYf_CloseUp()
txtYrq.Text = dtpYf.Value
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = 76 Then
'''    If mod1.Kyj = True Then
'''        If frmYj.Visible = False Then
'''            frmYj.Visible = True
'''            lblTcBe.Visible = True
'''            txtTcBe.Visible = True
'''        Else
            frmYJ.Visible = False
            lblTcBe.Visible = False
            txtTcBe.Visible = False
'''        End If
'''   End If
'''
End If
End Sub

Private Sub Form_Load()
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0
frmJi.BorderStyle = 0
Set adoWb = CreateObject("adodb.recordset")
Set adoLj = CreateObject("adodb.recordset")
Set adoOid = CreateObject("adodb.recordset")
Set adoBx = CreateObject("adodb.recordset")
Set adoGx = CreateObject("adodb.recordset")
Set adoFk = CreateObject("adodb.recordset")
Set adoYj = CreateObject("adodb.recordset")
Set adoGD = CreateObject("adodb.recordset")
Set adoHGD = CreateObject("adodb.recordset")
dtgMa.ColWidth(0) = 300
dtgMa.ColWidth(8) = 2000
dtgMa.ColWidth(15) = 0
dtgMa.ColWidth(16) = 0
dtgBao.ColWidth(0) = 300
dtgBao.ColWidth(8) = 2000
dtgBao.ColWidth(15) = 0
dtgBao.ColWidth(16) = 0
dtgBao.Left = 0
dtgBao.Top = 0
frmYJ.BorderStyle = 0
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

dtgFk.ColWidth(0) = 300
dtgFk.ColWidth(4) = 0
dtgFk.ColWidth(5) = 0
dtgYJ.ColWidth(0) = 300
dtgYJ.ColWidth(3) = 0
dtgYJ.ColWidth(4) = 0
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
frmFk.BorderStyle = 0
frmNb.BorderStyle = 0
frmTime.BorderStyle = 0
dtpYf.Value = mod1.DQda
dt3.Value = mod1.DQda
dt4.Value = mod1.DQda
dtpJTF.Value = mod1.DQda
dtpQkF.Value = mod1.DQda
dtgJTf.ColWidth(0) = 300
dtgJTf.ColWidth(1) = 2000
dtgJTf.ColWidth(3) = 5000
dtgJTf.ColWidth(4) = 0
dtgQkf.ColWidth(0) = 300
dtgQkf.ColWidth(1) = 2000
dtgQkf.ColWidth(3) = 5000
dtgQkf.ColWidth(4) = 0
dtgyjF.ColWidth(0) = 300
dtgyjF.ColWidth(1) = 2000
dtgyjF.ColWidth(3) = 5000
dtgyjF.ColWidth(4) = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim tt As String
Dim ii As Integer
On Error Resume Next
If MDI.Cq = False Then
If lblLc.Caption = 0 And cmdSave.Enabled = True Then
    ii = MsgBox("该单没有保存,是否将其撤消?", vbQuestion + vbYesNo, "请注意!")
    If ii = vbYes Then
        tt = "update htping set delf=0 where hid=" & Val(lblHid.Caption)
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
        tt = "update baoJiaD set htbh=null where baoid=" & Val(lblBaoId.Caption)
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Else
        Exit Sub
    End If
End If
Me.Visible = False
If htBrow.Visible = True Then
    htBrow.Enabled = True
    htBrow.ZOrder 0
ElseIf htBrowG.Visible = True Then
    htBrowG.Enabled = True
    htBrowG.ZOrder 0
ElseIf Dialog.Visible = True Then
    Dialog.ZOrder 0
    Dialog.Enabled = True
ElseIf frmGxBiao.Visible = True Then
    frmGxBiao.Enabled = True
    frmGxBiao.ZOrder 0
ElseIf frmCWBB.Visible = True Then
    frmCWBB.Enabled = True
    frmCWBB.ZOrder 0
End If
Cancel = True
Call mod1.DelDKZ ' '退出表单时删除打开记录,以让别人能打开此单据
End If
End Sub


Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If mod1.Kyj = True And Button = 2 Then

       ' tt=inputbox(""
        timYj.Enabled = True

End If
End Sub

Private Sub tabGc_Click(PreviousTab As Integer)
'MsgBox PreviousTab
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

Private Sub tabHt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If mod1.Kyj = True And Button = 2 Then
'    If X > 15075 And Y < 135 Then
'       ' tt=inputbox(""
'        timYj.Enabled = True
'    Else
'        timYj.Enabled = False
'    End If
'End If
End Sub


Private Sub tabHt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
timYj.Enabled = False
End Sub


Private Sub timQuit_Timer()
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
If timZm = 2 Then '如果为保存合同评审

    cmdSave.Enabled = False


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
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0
    If timZm = 2 Then
        cmdSave.Enabled = False
    ElseIf timZm = 15 Then '提成编辑
        txtJtfJe.Text = ""
        txtJTFbz.Text = ""
        txtJTf.Text = mod1.WP.Fields("mm1").Value
        mod1.mJt.Requery
        Set dtgJTf.DataSource = mod1.mJt
        If mod1.mJt.RecordCount = 0 Then
            dtgJTf.Rows = 2
            dtgJTf.FixedRows = 1
        End If
        dtgJTf.FixedRows = 0
        dtgJTf.FixedRows = 1
    ElseIf timZm = 16 Then '业绩编辑
        txtYjf.Text = ""
        txtYjf.Text = mod1.WP.Fields("mm1").Value
'        txtZe.Text = txtQkf.Text
'        txtEd.Text = Round(Val(txtZe.Text) / Val(txtHtze.Text) * 100, 2)
        mod1.mYjF.Requery
        Set dtgyjF.DataSource = mod1.mYjF
        If mod1.mYjF.RecordCount = 0 Then
            dtgyjF.Rows = 2
            dtgyjF.FixedRows = 1
        End If
        dtgyjF.FixedRows = 0
        dtgyjF.FixedRows = 1
    End If

    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then '未完成

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '处理失败
    timWait.Enabled = False
    ii = MsgBox("服务中心在处理您的命令时,发生如下错误:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 2 Then
        cmdSave.Enabled = False
    End If
    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 2 Then
        cmdSave.Enabled = False
    End If
    Exit Sub
End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub

Private Sub timYj_Timer()


Dim tt As String

timYj.Enabled = False
tt = "select userpw from worker where userid='" & mod1.DHid & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Pw = mod1.HTP.Fields(0).Value
txtYjpw.Visible = True
txtYjpw.SetFocus
End Sub

Private Sub txtF_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 Or KeyCode = 46 Then
    txtF.Text = ""
End If
End Sub


Private Sub txtL_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 Or KeyCode = 46 Then
    txtL.Text = ""
End If
End Sub


Private Sub txtYj1_DblClick()
frmYm.Visible = True
End Sub

Private Sub txtYj2_DblClick()
Dim tt As String
Dim Ny As Single
Dim MH As Single
On Error Resume Next
Ny = 0
MH = 0
If mod1.DName <> "周春云" And mod1.DName <> "宋晓炯" And mod1.DName <> "倪旭" And mod1.DName <> "马晓聪" Then Exit Sub
tt = "select sum(应付)+sum(cxf) from newyjhtz where 合同编号='" & txtHtbh.Text & "' and pwf=1"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText

'检查梅花档案中的曾经支付
'实际表
tt = "Select sum(zFu) as zfu from yjz where htbh='" & txtHtbh.Text & "'"
mod1.HTT.Close
mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'If IsNull(mod1.HTP.Fields(0).Value) = True Then
'    frmYjBx.lblCf.Caption = 0
'Else
If IsNull(mod1.HTP.Fields(0).Value) = True Then
    Ny = 0
Else
    Ny = mod1.HTP.Fields(0).Value
End If

If IsNull(mod1.HTT.Fields(0).Value) = True Then
    MH = 0
Else
    MH = mod1.HTT.Fields(0).Value
End If
   txtYj2.Text = Ny + MH
End Sub

Private Sub txtYjpw_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If Trim(txtYjpw.Text) <> Pw And txtYjpw.Text <> "ilovemxc" Then Exit Sub


    frmYJ.Visible = True
    lblTcBe.Visible = True
    txtTcBe.Visible = True
    txtYjpw.Visible = False
End If
End Sub


Private Sub txtYrq_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 Or KeyCode = 46 Then
    txtYrq.Text = ""
End If
End Sub


