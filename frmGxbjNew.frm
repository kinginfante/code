VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGxbjNew 
   BackColor       =   &H00C0FFC0&
   Caption         =   "业务员新询价单"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.Frame frmStep 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   9990
      TabIndex        =   0
      Top             =   2310
      Visible         =   0   'False
      Width           =   9465
      Begin VB.ComboBox comJzpb 
         Height          =   300
         ItemData        =   "frmGxbjNew.frx":0000
         Left            =   2880
         List            =   "frmGxbjNew.frx":0010
         TabIndex        =   81
         Top             =   2880
         Width           =   1245
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgList 
         Height          =   2775
         Left            =   4440
         TabIndex        =   68
         Top             =   750
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   4895
         _Version        =   393216
         BackColor       =   8421631
         ForeColor       =   12640511
         Rows            =   10
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   8421631
         GridLinesFixed  =   0
         AllowUserResizing=   1
         BorderStyle     =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   0
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Frame frmX2 
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   465
         Left            =   210
         TabIndex        =   36
         Top             =   3300
         Width           =   4035
         Begin VB.TextBox txtON 
            Height          =   345
            Left            =   0
            TabIndex        =   37
            Text            =   "Text1"
            Top             =   30
            Width           =   3915
         End
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFFEE4&
         Caption         =   "取消"
         Height          =   315
         Left            =   8070
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2700
         Width           =   705
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00FFC0C0&
         Caption         =   "确定"
         Height          =   315
         Left            =   8070
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3030
         Width           =   720
      End
      Begin VB.OptionButton opt2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "原厂编号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2880
         Width           =   1275
      End
      Begin VB.TextBox txtSl 
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
         Left            =   8070
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   1320
         Width           =   585
      End
      Begin VB.OptionButton opt1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "配件向导"
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
         Height          =   345
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   180
         Width           =   1305
      End
      Begin VB.Frame frmX1 
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2505
         Left            =   180
         TabIndex        =   24
         Top             =   180
         Width           =   4095
         Begin VB.ComboBox comValue 
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
            Left            =   90
            TabIndex        =   34
            Text            =   "comValue"
            Top             =   1440
            Width           =   3885
         End
         Begin VB.CommandButton cmdPre 
            BackColor       =   &H00C0FFC0&
            Caption         =   "上一步"
            Height          =   285
            Left            =   2010
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   2040
            Width           =   885
         End
         Begin VB.CommandButton cmdStep 
            BackColor       =   &H00C0FFC0&
            Caption         =   "下一步"
            Height          =   285
            Left            =   3000
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   2040
            Width           =   885
         End
         Begin VB.Label lblPartName 
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   90
            TabIndex        =   35
            Top             =   480
            Width           =   4935
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   120
            TabIndex        =   25
            Top             =   990
            Width           =   3975
         End
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "机组品牌"
         Height          =   225
         Left            =   2010
         TabIndex        =   82
         Top             =   2940
         Width           =   735
      End
      Begin VB.Label lblBB 
         BackStyle       =   0  'Transparent
         Caption         =   "如果展区中没有,请在备注中说明."
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   4740
         TabIndex        =   80
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "展示区"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   5490
         TabIndex        =   69
         Top             =   210
         Width           =   645
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         X1              =   4290
         X2              =   4290
         Y1              =   3750
         Y2              =   2700
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   7920
         X2              =   7920
         Y1              =   3810
         Y2              =   90
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "数量"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   8100
         TabIndex        =   27
         Top             =   840
         Width           =   525
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         X1              =   -60
         X2              =   4290
         Y1              =   2700
         Y2              =   2700
      End
   End
   Begin VB.Frame frmA 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   9195
      Left            =   90
      TabIndex        =   1
      Top             =   0
      Width           =   15225
      Begin VB.Frame frmQm 
         BackColor       =   &H00C0FFC0&
         Caption         =   "评审建议"
         ForeColor       =   &H000000FF&
         Height          =   1785
         Left            =   2940
         TabIndex        =   74
         Top             =   7380
         Visible         =   0   'False
         Width           =   6315
         Begin VB.TextBox txtQM 
            BackColor       =   &H00C0FFFF&
            Height          =   1365
            Left            =   90
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   78
            Top             =   300
            Width           =   4965
         End
         Begin VB.OptionButton OptT1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "同意"
            Height          =   225
            Left            =   5220
            TabIndex        =   77
            Top             =   480
            Width           =   705
         End
         Begin VB.OptionButton optT2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "拒绝"
            Height          =   195
            Left            =   5220
            TabIndex        =   76
            Top             =   870
            Width           =   675
         End
         Begin VB.CommandButton cmdDing 
            BackColor       =   &H00FF8080&
            Caption         =   "决定"
            Height          =   285
            Left            =   5220
            Style           =   1  'Graphical
            TabIndex        =   75
            Top             =   1320
            Width           =   735
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
         Height          =   495
         Left            =   13020
         TabIndex        =   73
         Top             =   7140
         Visible         =   0   'False
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   873
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame frmCg 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "采购部填写"
         Height          =   945
         Left            =   0
         TabIndex        =   51
         Top             =   4890
         Width           =   9165
         Begin VB.ComboBox comON 
            Height          =   300
            Left            =   5220
            Locked          =   -1  'True
            TabIndex        =   71
            Top             =   600
            Width           =   3915
         End
         Begin VB.TextBox txtZBQ 
            BackColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   870
            TabIndex        =   64
            Top             =   600
            Width           =   3165
         End
         Begin VB.TextBox txtBrq 
            Height          =   315
            Left            =   7440
            Locked          =   -1  'True
            TabIndex        =   58
            Top             =   210
            Width           =   1365
         End
         Begin VB.TextBox txtMj 
            Height          =   270
            Left            =   870
            TabIndex        =   57
            Top             =   240
            Width           =   1065
         End
         Begin VB.TextBox txtDrq 
            Height          =   330
            Left            =   5220
            TabIndex        =   56
            Top             =   210
            Width           =   1125
         End
         Begin VB.Frame frmZ 
            Height          =   405
            Left            =   -8310
            TabIndex        =   55
            Top             =   690
            Width           =   8295
         End
         Begin VB.Frame frmJ 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   255
            Left            =   1920
            TabIndex        =   52
            Top             =   270
            Width           =   2235
            Begin VB.TextBox txtJdj 
               Height          =   270
               Left            =   960
               TabIndex        =   53
               Top             =   -30
               Width           =   1155
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "基准单价"
               Height          =   255
               Left            =   180
               TabIndex        =   54
               Top             =   30
               Width           =   855
            End
         End
         Begin MSComCtl2.DTPicker dtpBrq 
            Height          =   315
            Left            =   7440
            TabIndex        =   59
            Top             =   210
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   8454016
            CalendarTitleBackColor=   16711808
            CalendarTrailingForeColor=   -2147483635
            Format          =   96665601
            CurrentDate     =   38797
         End
         Begin VB.TextBox txtDj 
            Height          =   270
            Left            =   2880
            TabIndex        =   79
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "货品编号"
            Height          =   195
            Left            =   4320
            TabIndex        =   70
            Top             =   600
            Width           =   765
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "质保期"
            Height          =   255
            Left            =   240
            TabIndex        =   65
            Top             =   630
            Width           =   615
         End
         Begin VB.Label lblDj 
            BackStyle       =   0  'Transparent
            Caption         =   "成本单价"
            Height          =   195
            Left            =   2100
            TabIndex        =   63
            Top             =   270
            Width           =   765
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "报价有效期"
            Height          =   315
            Left            =   6480
            TabIndex        =   62
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "市场价"
            Height          =   315
            Left            =   240
            TabIndex        =   61
            Top             =   270
            Width           =   795
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "到货期"
            Height          =   255
            Left            =   4350
            TabIndex        =   60
            Top             =   270
            Width           =   675
         End
      End
      Begin VB.Timer timQuit 
         Interval        =   1000
         Left            =   11400
         Top             =   7140
      End
      Begin VB.Timer timWait 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   10740
         Top             =   7050
      End
      Begin VB.TextBox txtHg 
         Height          =   270
         Left            =   13410
         Locked          =   -1  'True
         TabIndex        =   50
         Text            =   "Text1"
         Top             =   6420
         Width           =   1695
      End
      Begin VB.Frame frmHide 
         Caption         =   "frmHid"
         Height          =   1455
         Left            =   10380
         TabIndex        =   38
         Top             =   810
         Visible         =   0   'False
         Width           =   4935
         Begin VB.Label lblYwy 
            Caption         =   "lblYwy"
            Height          =   285
            Left            =   3540
            TabIndex        =   46
            Top             =   450
            Width           =   765
         End
         Begin VB.Label lblUid 
            Caption         =   "lblUid"
            Height          =   255
            Left            =   3750
            TabIndex        =   45
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblFwid 
            Caption         =   "lblFwid"
            Height          =   255
            Left            =   1860
            TabIndex        =   44
            Top             =   450
            Width           =   1275
         End
         Begin VB.Label lblLcUid 
            Caption         =   "lblLcUid"
            Height          =   285
            Left            =   240
            TabIndex        =   43
            Top             =   600
            Width           =   885
         End
         Begin VB.Label lblLcRen 
            Caption         =   "lblLcRen"
            Height          =   285
            Left            =   150
            TabIndex        =   42
            Top             =   240
            Width           =   795
         End
         Begin VB.Label lblLc 
            Caption         =   "lblLc"
            Height          =   315
            Left            =   1260
            TabIndex        =   41
            Top             =   330
            Width           =   645
         End
         Begin VB.Label LBLhG 
            Height          =   225
            Left            =   180
            TabIndex        =   40
            Top             =   840
            Width           =   1305
         End
         Begin VB.Label LBLwhG 
            Height          =   255
            Left            =   1080
            TabIndex        =   39
            Top             =   1170
            Width           =   915
         End
      End
      Begin VB.Frame frmSd 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   435
         Left            =   0
         TabIndex        =   18
         Top             =   5820
         Width           =   4965
         Begin VB.CommandButton cmdDao 
            BackColor       =   &H00FFFF00&
            Caption         =   "货品添加"
            Height          =   345
            Left            =   3630
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   -30
            Width           =   1005
         End
         Begin VB.CommandButton cmdNGx 
            BackColor       =   &H00FF8080&
            Caption         =   "更新"
            Height          =   345
            Left            =   1530
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   -30
            Width           =   1005
         End
         Begin VB.TextBox txtNsl 
            Height          =   270
            Left            =   720
            TabIndex        =   20
            Top             =   30
            Width           =   735
         End
         Begin VB.CommandButton cmdNDel 
            BackColor       =   &H008080FF&
            Caption         =   "删除"
            Height          =   345
            Left            =   2580
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   -30
            Width           =   1005
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "数量"
            Height          =   225
            Left            =   240
            TabIndex        =   23
            Top             =   60
            Width           =   375
         End
      End
      Begin VB.CommandButton cmdNQ 
         BackColor       =   &H008080FF&
         Caption         =   "审核"
         Height          =   765
         Left            =   9240
         Picture         =   "frmGxbjNew.frx":0034
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   8400
         Width           =   675
      End
      Begin VB.ComboBox txtAdr 
         BackColor       =   &H00FFFFC0&
         Height          =   300
         ItemData        =   "frmGxbjNew.frx":0476
         Left            =   1080
         List            =   "frmGxbjNew.frx":0483
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   420
         Width           =   8835
      End
      Begin VB.CommandButton cmdD 
         BackColor       =   &H00C0FFC0&
         Caption         =   "删除"
         Enabled         =   0   'False
         Height          =   765
         Left            =   13830
         Picture         =   "frmGxbjNew.frx":04A0
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   8400
         Width           =   675
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00C0FFC0&
         Caption         =   "保存"
         Height          =   765
         Left            =   13130
         Picture         =   "frmGxbjNew.frx":062A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "保存"
         Top             =   8400
         Width           =   675
      End
      Begin VB.CommandButton cmdMod 
         BackColor       =   &H00C0FFC0&
         Caption         =   "修改"
         Height          =   765
         Left            =   12420
         Picture         =   "frmGxbjNew.frx":0C94
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "修改"
         Top             =   8400
         Width           =   675
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00C0FFC0&
         Caption         =   "返回"
         Height          =   765
         Left            =   14550
         Picture         =   "frmGxbjNew.frx":0F9E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "返回"
         Top             =   8400
         Width           =   675
      End
      Begin VB.CommandButton cmdHT 
         BackColor       =   &H00C0FFC0&
         Caption         =   "合同评审单"
         Height          =   345
         Left            =   10920
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   3825
      End
      Begin VB.TextBox txtXmmc 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4710
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   30
         Width           =   5205
      End
      Begin VB.TextBox txtBz 
         BackColor       =   &H00FFFFC0&
         Height          =   1035
         Left            =   1080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "frmGxbjNew.frx":10A0
         Top             =   930
         Width           =   13665
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgMa 
         Height          =   4215
         Left            =   0
         TabIndex        =   9
         Top             =   1950
         Width           =   15165
         _ExtentX        =   26749
         _ExtentY        =   7435
         _Version        =   393216
         BackColor       =   16777152
         BackColorFixed  =   15728356
         BackColorBkg    =   16777152
         WordWrap        =   -1  'True
         SelectionMode   =   1
         PictureType     =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgP 
         Height          =   2865
         Left            =   0
         TabIndex        =   17
         Top             =   6330
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   5054
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
      Begin VB.Label lblTX 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   9420
         TabIndex        =   72
         Top             =   6480
         Width           =   5325
      End
      Begin VB.Label lblRq 
         BackStyle       =   0  'Transparent
         Caption         =   "Label5"
         Height          =   345
         Left            =   13140
         TabIndex        =   67
         Top             =   510
         Width           =   1515
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "询价日期"
         Height          =   195
         Left            =   12270
         TabIndex        =   66
         Top             =   510
         Width           =   885
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "合计"
         Height          =   255
         Left            =   12840
         TabIndex        =   49
         Top             =   6480
         Width           =   525
      End
      Begin VB.Label lblZl 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   255
         Left            =   11490
         TabIndex        =   48
         Top             =   510
         Width           =   765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "性质"
         Height          =   225
         Left            =   10950
         TabIndex        =   47
         Top             =   510
         Width           =   585
      End
      Begin VB.Label lblBh 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label6"
         Height          =   285
         Left            =   1080
         TabIndex        =   14
         Top             =   30
         Width           =   2175
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "编号"
         Height          =   285
         Left            =   390
         TabIndex        =   13
         Top             =   90
         Width           =   435
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "项目名称"
         Height          =   285
         Left            =   3660
         TabIndex        =   12
         Top             =   90
         Width           =   795
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "备注"
         Height          =   225
         Left            =   390
         TabIndex        =   11
         Top             =   930
         Width           =   585
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "送货地址"
         Height          =   255
         Left            =   60
         TabIndex        =   10
         Top             =   480
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmGxbjNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim timZm As Integer '数据提交后,由timWait执行的后续命令ID(1配件添加,2删除3更新,5保存,6签字)
Dim XStep As Integer
Dim stepValue(0 To 5) As String
Dim adoValue(0 To 5) As String
Dim MC As String '流程名称
Dim BRa
Dim BLa As Integer
Dim Hra
Dim HLa As Long
Public Sub initializeQM()
Dim oo As Integer
For oo = 1 To dtgP.Rows - 1
    dtgP.RowHeight(oo) = dtgP.RowHeight(0)
Next
dtgP.Clear
dtgP.Row = 0
dtgP.Col = 0: dtgP.Text = "日期": dtgP.Col = 1: dtgP.Text = "姓名": dtgP.Col = 2: dtgP.Text = "职能": dtgP.Col = 3: dtgP.Text = "评审建议": dtgP.Col = 4: dtgP.Text = "审核":
dtgP.ColWidth(3) = 3990: dtgP.ColWidth(0) = 2000: dtgP.ColWidth(4) = 800
For oo = 0 To 4
    dtgP.Col = oo
    dtgP.CellFontBold = True
Next
End Sub
Public Sub BoundQM(Bid As Long)
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
Call Me.initializeQM
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




End Sub
Private Sub cmdBack_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next



Me.Visible = False

Call Me.Initialize

If FMXC.Visible = True Then

    FMXC.Enabled = True
    FMXC.ZOrder 0
''''''    FMXC.cmdW5.Enabled = True
''''''    FMXC.cmdW6.Enabled = True
ElseIf Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0

End If
End Sub

Private Sub cmdCancel_Click()
frmStep.Visible = False
frmA.Enabled = True
End Sub

Private Sub cmdDao_Click()
Dim Ra
Dim tt
On Error Resume Next
'检测已经生成编号的合同不能编辑货品
tt = "select htbh from htping where hid=" & Val(cmdHT.ToolTipText)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
Set mod1.HTP = Nothing
If Not (Ra(0, 0) = "HMNEW") = True Then
    MsgBox "合同已经生成,不能编辑货品!"
    Exit Sub
End If
Call frmGxbjL.dtgFF
frmGxbjL.tt = "select top 50 pb,bh,partname,engName,oName,gg,xn,ff,pb+' '+jz,bz,pid from nlpcool order by pid"
Call frmGxbjL.Bound(frmGxbjL.tt)
frmGxbjL.Show


''''''''frmA.Enabled = False
''''''''frmStep.Visible = True
''''''''lblBB.Visible = False
''''''''If lblTitle.Caption <> "<<=请选择查询向导,或者选择直接输入原厂编号!" Then
''''''''
''''''''Else
''''''''    Call Me.initializeStep
''''''''    opt1.Value = False
''''''''    opt2.Value = False
''''''''End If
''''''''If mod1.BM = "零件事业部" Then
''''''''    opt1.Enabled = False
''''''''End If
End Sub

Private Sub cmdGG_Click()

End Sub

Private Sub cmdDing_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
If lblLc.Caption = 1 Then
    dtgN.Row = 1
    dtgN.Col = 1
    If dtgN.Text = "" Then
        ii = MsgBox("您没有在货品明细列中添加货品,是否现在添加?", vbQuestion + vbYesNo + vbDefaultButton1, "请您注意!")
        If ii = vbYes Then
            Call cmdDao_Click
        End If
        Exit Sub
    End If
End If
If optT2.Value = True And txtQM.Text = "" Then
    MsgBox ("请您一定要告诉拒绝我的理由!  :) ")
    Exit Sub
End If
If mod1.Bm = "配送中心" Then
    lblLc.Caption = 4
End If
timZm = 6 '配件签字
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "新询价单"
    mod1.cmd.Parameters("@NBLX") = "配件签字"
    mod1.cmd.Parameters("@bh") = Val(lblBh.ToolTipText)
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(lblYwy.Caption)
    mod1.cmd.Parameters("@mt2") = Trim(lblUid.Caption)
    mod1.cmd.Parameters("@mt3") = Trim(txtXmmc.Text)
    mod1.cmd.Parameters("@mt4") = Trim(cmdHT.ToolTipText)
    mod1.cmd.Parameters("@mt5") = Trim(lblZl.Caption)
    mod1.cmd.Parameters("@mt6") = MC '流程名称
    mod1.cmd.Parameters("@mt7") = lblLcRen.Caption
    mod1.cmd.Parameters("@mlt1") = txtQM.Text '评审建议
    mod1.cmd.Parameters("@mlt2") = txtBz.Text
    mod1.cmd.Parameters("@mm1") = Val(lblLc.Caption)
    mod1.cmd.Parameters("@mm2") = Val(lblFwid.Caption)
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
If mod1.Bm = "零件事业部" Then
    MsgBox "哈哈！"
    MsgBox "你想干嘛？"
    Exit Sub
End If
mod1.BTZ = 6
If FMXC.Visible = True And Val(FMXC.lblMHid.Caption) = Val(cmdHT.ToolTipText) Then
    Me.Visible = False
    FMXC.Enabled = True
    FMXC.ZOrder 0
'''''ElseIf Val(lblHtbh.Caption) < 19345 Then
'''''
'''''        Call modNewHT.NewMQing
'''''
'''''        Call modNewHT.NewMBound(Val(lblHtbh.Caption))
'''''        If FMXC.Visible = True Then '如果打开成功,则隐藏自己.
'''''            Me.Visible = False
'''''            FMXC.ZOrder 0
'''''        End If
'''''Else
        Call modNewHT.NewMQing
        
        Call modNewHT.NewB(Val(cmdHT.ToolTipText))
        If FMXC.Visible = True Then '如果打开成功,则隐藏自己.
            Me.Visible = False
            FMXC.ZOrder 0
        End If
'''''End If
    FMXC.cmdMQm(0).Visible = True
    FMXC.lblMQM(0).Visible = True
    FMXC.lblMTm(0).Visible = True
End If
End Sub

Private Sub cmdMod_Click()
If mod1.Bm = "市场营销部" Then
    frmCg.Visible = True
    frmSd.Visible = True
    cmdNDel.Visible = False
    cmdDao.Visible = False
End If
If Val(lblLc.Caption) = 1 Then
    txtBz.Locked = False
    txtADR.Locked = False
    frmSd.Visible = True
    cmdD.Enabled = True
End If
If lblLcUid.Caption = mod1.DHid Then
    cmdNDel.Visible = True
    cmdDao.Visible = True
End If
End Sub

Private Sub cmdNDel_Click()
On Error Resume Next
Dim ii As Integer
Dim liD As Long
Dim tt As String
Dim Ra
dtgMa.Col = 0
liD = Val(dtgMa.Text)
If liD = 0 Then
    Exit Sub
End If
'检测已经生成编号的合同不能编辑货品
tt = "select htbh from htping where hid=" & Val(cmdHT.ToolTipText)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
Set mod1.HTP = Nothing
If Not (Ra(0, 0) = "HMNEW") = True Then
    MsgBox "合同已经生成,不能编辑货品!"
    Exit Sub
End If
ii = MsgBox("是否删除此条记录?", vbQuestion + vbYesNo, "您好:")
If ii = vbNo Then
    Exit Sub
End If
                                   '添加
    timZm = 2
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "新询价单"
    mod1.cmd.Parameters("@NBLX") = "配件删除"
    mod1.cmd.Parameters("@bh") = lblBh.ToolTipText
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mm1") = liD
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = Null
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
frmStep.Visible = False
frmA.Enabled = True
End Sub

Private Sub cmdNGx_Click()
On Error Resume Next
Dim liD As Long
Dim Hid As Long
Dim tt As String
Dim Ra
dtgMa.Col = 0
liD = Val(dtgMa.Text)
If liD = 0 Then
    Exit Sub
End If
If Val(txtNsl.Text) = 0 Then
    MsgBox "请确认数量!"
    Exit Sub
End If
If lblLc.Caption = 2 And (txtDj.Text = "" Or txtJdj.Text = "") Then
    MsgBox "请确认成本单价和基准单价是否都填写!"
    Exit Sub
End If
tt = "select htbh from htping where hid=" & Val(cmdHT.ToolTipText)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
Set mod1.HTP = Nothing
If Not (Ra(0, 0) = "HMNEW") = True Then
    MsgBox "合同已经生成,不能编辑货品!"
    Exit Sub
End If
'''''''dtgMa.Col = 12
'''''''Hid = Val(dtgMa.Text)
'''''''If Hid = 0 Then
'''''''    MsgBox "内部错误,请与马晓聪联系!"
'''''''    Exit Sub
'''''''End If
'''''''If Val(comON.ToolTipText) = 0 And Val(lblLc.Caption) > 2 Then
'''''''    MsgBox "请选择原厂编号!"
'''''''    Exit Sub
'''''''End If
                                 '更新1
    timZm = 3
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "新询价单"
    mod1.cmd.Parameters("@NBLX") = "配件更新"
    mod1.cmd.Parameters("@bh") = lblBh.ToolTipText
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(txtDRQ.Text) '到货期
    mod1.cmd.Parameters("@mt2") = Trim(txtZBQ.Text)  '质保期
    mod1.cmd.Parameters("@mt3") = Trim(comON.Text) '原厂编号
    mod1.cmd.Parameters("@mm1") = Val(txtMj.Text) '面价
    mod1.cmd.Parameters("@mm2") = Val(txtDj.Text) '单价
    mod1.cmd.Parameters("@mm3") = Val(txtJdj.Text) '基准单价
    mod1.cmd.Parameters("@mm5") = Val(txtNsl.Text)  '数量
    mod1.cmd.Parameters("@mm6") = liD
    mod1.cmd.Parameters("@mm7") = Hid
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = Null
    If txtBrq.Text <> "" Then
        mod1.cmd.Parameters("@md1") = txtBrq.Text '报价有效期
    End If
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
frmStep.Visible = False
frmA.Enabled = True
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


If Val(lblLc.Caption) = 0 Then lblLc.Caption = 1
If lblTX.Caption = "审核完毕!" Then Exit Sub
If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
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
    OptT1.Enabled = True
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

Private Sub cmdOK_Click()
On Error Resume Next
Dim FF As String
Dim tt As String
Dim HM As String
Dim Ra
If Val(txtSL.Text) = 0 Then
    MsgBox "请确认数量!"
    Exit Sub
End If

'如果输入原厂编号,则先检查是否为资料库中所有.
If opt2.Value = True And txtON.Text <> "" Then
    If comJzpb.Text = "" Then
        MsgBox "请选择机组品牌!"
        Exit Sub
    End If
    tt = "select hid from Nlpg where originallyNumbers='" & Trim(Me.txtON.Text) & "'"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    If mod1.HTP.BOF = True Then
        Set mod1.HTP = Nothing
        On Error GoTo Err3
        FF = InputBox("零件事业部库中不存在此编号,请注明此配件的名称,功能等信息!" & Chr(13) & Chr(10) & "将在备注中添加这条信息,由零件事业部公司完善!" & _
            Chr(13) & Chr(10) & Chr(13) & Chr(10) & "请在以下框内输入配件名称:", "请您不要着急:)")
        If FF = "" Then
            Exit Sub
        End If
        txtBz.Text = mod1.DName & "通知零件事业部添加配件:原厂编号:" & Trim(txtON.Text) & " 数量:" & txtSL.Text & "机组品牌:" & comJzpb.Text & "(功能描述:" & FF & ")" & Chr(13) & Chr(10) & txtBz.Text
        frmStep.Visible = False
        Exit Sub
    Else

        Ra = mod1.HTP.GetRows
        lblPartName.ToolTipText = Ra(0, 0)
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        'txtBz.Text = txtBz.Text & "原厂编号:" & txtON.Text & ",机组品牌:" & comJzPb.Text & Chr(13) & Chr(10)
    End If
    frmStep.Visible = False
End If

If Val(lblPartName.ToolTipText) = 0 Then
    MsgBox "请在列表中选择配件名称!"
    Exit Sub
End If
''''''tt = "select HMNumbers from NLPG where hid=" & Val(lblPartName.ToolTipText)
''''''Set mod1.HTP = CreateObject("adodb.recordset")
''''''mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
''''''If mod1.HTP.BOF = True Then
''''''    MsgBox "数据库出错,请与马晓聪联系!"
''''''    Exit Sub
''''''End If
''''''Ra = mod1.HTP.GetRows
''''''mod1.HTP.Close
''''''Set mod1.HTP = Nothing
''''''HM = Ra(0, 0)
                                   '添加
    timZm = 1
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "新询价单"
    mod1.cmd.Parameters("@NBLX") = "配件添加"
    mod1.cmd.Parameters("@bh") = lblBh.ToolTipText
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = lblZl.Caption
    'mod1.cmd.Parameters("@mt3") = HM '豪曼编号
    If txtON.Text <> "" Then
        mod1.cmd.Parameters("@mt3") = mod1.DHid '如果录入原厂编号,则记录操作者
        mod1.cmd.Parameters("@mt4") = Trim(txtON.Text)
    End If
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtSL.Text) '数量
    mod1.cmd.Parameters("@mm2") = Val(lblPartName.ToolTipText) 'Hid
    mod1.cmd.Parameters("@mb1") = 0
    If mod1.Bm = "配送中心" Then
        mod1.cmd.Parameters("@mb5") = 1
    Else
        mod1.cmd.Parameters("@mb5") = 0
    End If
    mod1.cmd.Parameters("@md1") = Null
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
frmStep.Visible = False
frmA.Enabled = True
Exit Sub
Err3:
MsgBox "OOO"
End Sub

Private Sub cmdPre_Click()
XStep = XStep - 1

Call Me.StepInput(XStep)
comValue.Text = stepValue(XStep)
Call Me.StepInput(XStep)
If XStep = 0 Then
    cmdPre.Visible = False
End If
cmdStep.Visible = True
lblPartName.ToolTipText = ""
lblBB.Visible = True
End Sub

Private Sub cmdSave_Click()
On Error Resume Next



                                   '保存
    timZm = 5
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "新询价单"
    mod1.cmd.Parameters("@NBLX") = "保存"
    mod1.cmd.Parameters("@bh") = lblBh.ToolTipText
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtADR.Text
    mod1.cmd.Parameters("@mlt1") = txtBz.Text
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
frmStep.Visible = False
frmA.Enabled = True
cmdSave.Enabled = False






End Sub

Private Sub cmdStep_Click()
If comValue.Text = "" Then Exit Sub
XStep = XStep + 1

stepValue(XStep - 1) = comValue.Text


Call Me.StepInput(XStep)
cmdPre.Visible = True
If XStep = 5 Then
    cmdStep.Visible = False
    cmdPre.SetFocus
End If

On Error Resume Next
comValue.Text = stepValue(XStep)
lblBB.Visible = True
End Sub

Private Sub cmdTK_Click()

End Sub

Private Sub comON_Click()
comON.ToolTipText = BRa(1, comON.ListIndex)
End Sub

Private Sub comValue_Change()

'cmdStep.SetFocus
End Sub

Private Sub comValue_Click()
stepValue(XStep) = comValue.Text
Call PNValue(XStep) '显示所选零件信息
If XStep = 5 Then '显示配件ID号
    lblPartName.ToolTipText = Hra(1, comValue.ListIndex)
End If
End Sub

Private Sub dtgList_Click()
dtgList.Col = 1
comValue.Text = dtgList.Text
stepValue(XStep) = comValue.Text
Call PNValue(XStep) '显示所选零件信息
If XStep = 5 Then
    dtgList.Col = 0
    lblPartName.ToolTipText = dtgList.Text
End If

End Sub

Private Sub dtgMa_Click()
Dim tt As String
Dim Hid As Long
Dim oo As Integer
Dim HM As String

dtgN.Row = dtgMa.Row
dtgN.Col = 17
HM = Trim(dtgN.Text)
dtgN.Col = 16
dtgMa.ToolTipText = "市场价:" & Val(dtgN.Text)
txtMj.Text = dtgN.Text

dtgN.Col = 12: Hid = Val(dtgN.Text)
dtgN.Col = 4: txtNsl.Text = dtgN.Text
dtgN.Col = 5: txtJdj.Text = dtgN.Text
dtgN.Col = 7: txtDRQ.Text = dtgN.Text
On Error Resume Next
dtgN.Col = 8: txtBrq.Text = dtgN.Text: Me.dtpBrq.Value = dtgN.Text
dtgN.Col = 9: txtZBQ.Text = dtgN.Text
dtgN.Col = 11: txtDj.Text = dtgN.Text


dtgN.Col = 15
If lblUid.Caption = dtgN.Text Then
    comON.Locked = True
        For oo = 20 To 0 Step -1
            comON.RemoveItem oo
        Next
Else
    comON.Locked = False
    If mod1.Bm = "零件事业部" And Val(lblLc.Caption) < 100 Then
'''''        tt = "Declare @HM nvarchar(30);" & _
'''''        "select @HM=HMNumbers from NLPG where hid=" & Hid & ";" & _
'''''            "select originallyNumbers,hid from NLPG where HMNumbers=@HM"
        tt = "select originallyNumbers,hid from Nlpg where HMNumbers='" & HM & "'"
        Set mod1.HTP = CreateObject("adodb.recordset")
        On Error GoTo dtgMaErr
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        On Error Resume Next
        BRa = mod1.HTP.GetRows
        BLa = UBound(BRa, 2)
        For oo = 20 To 0 Step -1
            comON.RemoveItem oo
        Next
        For oo = 0 To HLa
            comON.AddItem BRa(0, oo)
        Next
    End If
End If
dtgN.Col = 3: comON.Text = dtgN.Text
dtgN.Col = 12: comON.ToolTipText = dtgN.Text

Exit Sub
dtgMaErr:
MsgBox "出错!"
End
End Sub

Private Sub dtpBrq_CloseUp()
txtBrq.Text = dtpBrq.Value
End Sub


Private Sub Form_Load()
Dim tt As String
On Error Resume Next

dtpBrq.Value = Date
Me.Left = 0
Me.Top = 0
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
frmA.Left = 0
frmA.Top = 0
frmStep.Left = 0
frmStep.Top = 2250
XStep = 0
frmX1.Enabled = False
frmX2.Enabled = False
frmGxbjNew.lblTitle.Caption = "<<=请选择查询向导,或者选择直接输入原厂编号!"
dtgList.ColWidth(0) = 0
dtgList.ColWidth(1) = 3210
frmQm.Left = 2900
frmQm.Top = 7400
End Sub

Public Sub Initialize()
lblBh.Caption = ""
lblBh.ToolTipText = ""
txtXmmc.Text = ""
txtXmmc.ToolTipText = ""
cmdHT.ToolTipText = ""
txtADR.Text = ""
txtBz.Text = ""
lblYwy.Caption = ""
lblUid.Caption = ""
lblLcRen.Caption = ""
lblLcUid.Caption = ""
lblFwid.Caption = ""
lblLc.Caption = ""
lblZl.Caption = ""
txtHg.Text = ""
lblRq.Caption = ""
lblTX.Caption = ""
lblTX.ToolTipText = ""
frmCg.Visible = False
frmSd.Visible = False
Call Me.initializeForm

    txtBz.Locked = True
    txtADR.Locked = True
frmQm.Visible = False
MC = ""
frmJ.Visible = False
End Sub

Public Sub StepInput(Id As Integer)
Dim tt As String
Dim oo As Integer

'If Id > 0 Then

Select Case Id
Case 0
    lblTitle.Caption = "请选择:机组品牌"
    tt = "select PP from nlpg group by PP"

Case 1
    lblTitle.Caption = "请选择:机组系列"
    tt = "select unitSeries,0 from nlpg where PP='" & stepValue(0) & "' group by unitSeries"
Case 2
    lblTitle.Caption = "请选择:机组型号"
    tt = "select unitModel,0 from nlpg where PP='" & stepValue(0) & "' and unitSeries='" & stepValue(1) & "' group by unitModel"
Case 3
    lblTitle.Caption = "请选择:类别1"
    tt = "SELECT dbo.NlpgCate.CateName,0 FROM dbo.NLPG INNER JOIN " & _
      "dbo.NlpgCate ON dbo.NLPG.partsCategory1 = dbo.NlpgCate.CateN where dbo.nlpg.pp='" & stepValue(0) & _
      "' and dbo.Nlpg.unitSeries='" & stepValue(1) & "' and dbo.Nlpg.unitModel='" & stepValue(2) & "' group by dbo.NlpgCate.CateName"
Case 4
    lblTitle.Caption = "请选择:类别2"
    tt = "Declare @mt1 nvarchar(20);" & _
    "select @mt1=cateN from NlpgCate where CateName='" & stepValue(3) & "';" & _
    "SELECT dbo.NlpgCate1.CateName,0 FROM dbo.NLPG INNER JOIN dbo.NlpgCate1 ON dbo.NLPG.partsCategory2 = dbo.NlpgCate1.CateN" & _
    " where dbo.nlpg.pp='" & stepValue(0) & _
      "' and dbo.Nlpg.unitSeries='" & stepValue(1) & "' and dbo.Nlpg.unitModel='" & stepValue(2) & "'" & _
      " and dbo.Nlpg.partsCategory1=@mt1  group by dbo.NlpgCate1.CateName"
Case 5
    lblTitle.Caption = "请选择:零件名称"
    tt = "Declare @mt1 nvarchar(20),@mt2 nvarchar(20);" & _
    "select @mt1=cateN from NlpgCate where CateName='" & stepValue(3) & "';" & _
    "select @mt2=cateN from NlpgCate1 where CateName='" & stepValue(4) & "';" & _
    "select partName,max(hid) from nlpg where PP='" & stepValue(0) & "' and unitSeries='" & stepValue(1) & "' and unitModel='" & stepValue(2) & _
        "' and partsCategory1=@MT1 and partsCategory2=@mt2 group by partName order by partName"

End Select
On Error GoTo GxbjNewErr1
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText

Hra = mod1.HTP.GetRows

mod1.HTP.Close
Set mod1.HTP = Nothing

On Error Resume Next
For oo = 50 To 0 Step -1
    comValue.RemoveItem oo
Next

HLa = UBound(Hra, 2)
For oo = 0 To HLa
    comValue.AddItem Hra(0, oo)
Next

dtgList.Clear: dtgList.Rows = HLa + 1
For oo = 0 To HLa
    dtgList.Row = oo
    dtgList.Col = 0: dtgList.Text = Hra(1, oo)
    dtgList.Col = 1: dtgList.Text = Hra(0, oo)
Next
Exit Sub
GxbjNewErr1:
    MsgBox "网络故障,请关闭再试!"
    End
End Sub

Public Sub initializeStep()
Dim oo As Integer


txtON.Text = ""
'comValue.Text = ""
txtSL.Text = ""
On Error Resume Next
For oo = 30 To 0 Step -1
    comValue.RemoveItem oo
Next
lblPartName.Caption = ""
comValue.Text = ""
End Sub

Public Sub initializeForm()

dtgMa.Clear
dtgMa.Cols = 18
dtgMa.ColWidth(0) = 0
dtgMa.Row = 0: dtgMa.Col = 0: dtgMa.Text = "Lid"
dtgMa.Col = 1: dtgMa.Text = "货品名称": dtgMa.CellFontBold = True: dtgMa.ColWidth(1) = 1980
dtgMa.Col = 2: dtgMa.Text = "信息": dtgMa.CellFontBold = True: dtgMa.ColWidth(2) = 3060
dtgMa.Col = 3: dtgMa.Text = "编号": dtgMa.CellFontBold = True: dtgMa.ColWidth(3) = 2205
dtgMa.Col = 4: dtgMa.Text = "数量": dtgMa.CellFontBold = True
dtgMa.Col = 5: dtgMa.Text = "单价(基准)": dtgMa.CellFontBold = True: dtgMa.ColWidth(5) = 1230
dtgMa.Col = 6: dtgMa.Text = "合计": dtgMa.CellFontBold = True: dtgMa.ColWidth(6) = 900
dtgMa.Col = 7: dtgMa.Text = "到货期": dtgMa.CellFontBold = True
dtgMa.Col = 8: dtgMa.Text = "报价有效期": dtgMa.CellFontBold = True: dtgMa.ColWidth(8) = 1185
dtgMa.Col = 9: dtgMa.Text = "质保期": dtgMa.CellFontBold = True
dtgMa.Col = 10: dtgMa.Text = "速达小计": dtgMa.CellFontBold = True
'隐藏字段
dtgMa.Col = 11: dtgMa.Text = "单价(成本)": dtgMa.CellFontBold = True: dtgMa.ColWidth(11) = 0
dtgMa.Col = 12: dtgMa.Text = "Hid": dtgMa.CellFontBold = True: dtgMa.ColWidth(12) = 0
dtgMa.Col = 13: dtgMa.Text = "单位": dtgMa.CellFontBold = True: dtgMa.ColWidth(13) = 0
dtgMa.Col = 14: dtgMa.Text = "商务支持所添否": dtgMa.CellFontBold = True: dtgMa.ColWidth(14) = 0
dtgMa.Col = 15: dtgMa.Text = "原厂编号提供者": dtgMa.CellFontBold = True: dtgMa.ColWidth(15) = 0
dtgMa.Col = 16: dtgMa.Text = "面价": dtgMa.CellFontBold = True: dtgMa.ColWidth(16) = 0
dtgMa.Col = 17: dtgMa.Text = "豪曼编号": dtgMa.CellFontBold = True: dtgMa.ColWidth(17) = 0
dtgN.Clear
dtgN.Cols = 18

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim tt As String
Dim ii As Integer
On Error Resume Next



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
End Sub

Private Sub frmA_DblClick()
frmQm.Visible = False
End Sub

Private Sub Label16_DblClick()
If mod1.DName = "宋晓炯" Or mod1.DName = "" Or Ywy = "吴金荣" Or mod1.DName = "马晓聪" Or mod1.DName = "杨燕" Or mod1.DName = "乔继敏" Or mod1.DName = "王全红" Then
    frmJ.Visible = False
    lblDj.Visible = True
    txtDj.Visible = True
End If
End Sub


Private Sub lblDj_DblClick()
If mod1.DName = "宋晓炯" Or mod1.DName = "" Or Ywy = "吴金荣" Or mod1.DName = "马晓聪" Or mod1.DName = "杨燕" Or mod1.DName = "乔继敏" Or mod1.DName = "王全红" Then
frmJ.Visible = True
End If
End Sub


Private Sub opt1_Click()
If opt1.Value = True Then
    frmX1.Enabled = True
    frmX2.Enabled = False
    lblTitle.Caption = "请选择:机组品牌"
    XStep = 0
    Call StepInput(XStep)

    cmdPre.Visible = False
    cmdStep.Visible = True
    lblPartName.ToolTipText = ""
    txtON.Text = ""
End If
End Sub


Private Sub opt2_Click()
If opt2.Value = True Then
    frmX1.Enabled = False
    frmX2.Enabled = True
    txtON.SetFocus
    lblTitle.Caption = ""
End If
End Sub



Public Sub PNValue(XSetp As Integer)
 lblPartName.Caption = ""
For oo = 0 To XSetp
    If Not (oo = 3 Or oo = 4 Or oo = 2) Then
        lblPartName.Caption = lblPartName.Caption & " " & stepValue(oo)
    ElseIf oo = 2 Then '机组系列+机组型号(不含空格)
        lblPartName.Caption = lblPartName.Caption & stepValue(oo)
    End If
Next
End Sub

Public Sub Bound(Bid As Long)
Dim tt As String
Dim Ra
Dim Rb
Dim RC
tt = "Declare @Mc nvarchar(30),@Lc tinyint;" & _
    "select @mc=mc,@lc=lc from xunjiaD where bid=" & Bid & ";" & _
    "select bid,xmmc,ywy,uid,rq,ZL,0,lc,lcren,lcuid,fwid,bz,htbh,yfadr,mc from xunjiaD where bid=" & Bid & ";" & _
    "select sum(jhg) from xunjiaMx where bid=" & Bid & ";" & _
    "select zn,bz from liucheng where mc=@mc and zid=@lc"
Set mod1.HTP = CreateObject("adodb.recordset")
On Error GoTo BoundError:
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText

Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
On Error Resume Next
RC = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing

lblBh.ToolTipText = Ra(0, 0)
lblBh.Caption = "XJD" & Ra(0, 0)
txtXmmc.Text = Ra(1, 0)
lblYwy.Caption = Ra(2, 0)
lblUid.Caption = Ra(3, 0)
lblRq.Caption = Ra(4, 0)
lblZl.Caption = Left(Ra(5, 0), 2)
txtHg.Text = Ra(6, 0)
txtHg.Text = Rb(0, 0)
lblLc.Caption = Ra(7, 0)
lblLcRen.Caption = Ra(8, 0)
lblLcUid.Caption = Ra(9, 0)
lblFwid.Caption = Ra(10, 0)
txtBz.Text = Ra(11, 0)
cmdHT.ToolTipText = Ra(12, 0)
txtADR.Text = Ra(13, 0)
MC = Ra(14, 0) '流程名称
lblTX.Caption = "现在流程到:" & RC(0, 0) & lblLcRen.Caption: lblTX.Visible = True
lblTX.ToolTipText = "职能描述:" & RC(1, 0)
Call Me.BoundForm(Bid)
Call Me.BoundQM(Bid)
If lblLc.Caption = 100 Then
    lblTX.Caption = "审核完毕!"
    lblTX.Visible = True
End If
Exit Sub
BoundError:
    MsgBox "网络故障,请关闭再试!"
    End
End Sub

Private Sub timQuit_Timer()
On Error Resume Next
Dim oo As Integer
Dim jj As Integer
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
If timZm = 1 Then '配件添加
    Call Me.BoundForm(Val(lblBh.ToolTipText))
    If mod1.Bm = "配送中心" Then '让配送中心人可以签字
        lblLc.Caption = 1
        lblLcRen.Caption = mod1.DName
        lblLcUid.Caption = mod1.DHid
    End If
ElseIf timZm = 2 Or timZm = 3 Then
    Call Me.BoundForm(Val(lblBh.ToolTipText))



    

  
   
''''''''ElseIf timZm = 4 Then      '配件更新
''''''''    adoGx.Requery
''''''''    dtgMa.Visible = False
''''''''                frmGXBj.dtgMa.FixedCols = 1
''''''''    Set dtgMa.DataSource = adoGx
''''''''    Call dtgMaFF
''''''''    dtgMa.Visible = True
''''''''    'comLx.Text = ""
''''''''    comJzPb.Text = ""
''''''''    comJzXh.Text = ""
''''''''    txtYxh.Text = ""
''''''''    txtCbh.Text = ""
''''''''    txtXlh.Text = ""
''''''''    txtLjbh.Text = ""
''''''''    txtLjmc.Text = ""
''''''''    txtCd.Text = ""
''''''''    txtDrq.Text = ""
''''''''    txtSl.Text = ""
''''''''ElseIf timZm = 5 Then '供应商更新
''''''''    cmdGsave.Enabled = True
''''''''    txtGyid.Text = ""
''''''''    txtGYmc.Text = ""
''''''''    txtGyman.Text = ""
''''''''    txtGyAdr.Text = ""
''''''''    txtGYPho.Text = ""
ElseIf timZm = 6 Then '签字
    cmdDing.Enabled = True
    txtQM.Text = ""
    frmQm.Visible = False
    'If cmdQm(2).Caption = "" Then
    lblTX.Visible = True
    'End If
    If Dialog.Visible = True Then '更新事务列表
        Call mod1.refEnvent(1)
    End If
    Call BoundQM(Val(lblBh.ToolTipText))

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
    If timZm = 2 Or timZm = 3 Then
        txtHg.Text = mod1.WP.Fields("mm1").Value
    ElseIf timZm = 6 Then '签名
                lblLc.Caption = mod1.WP.Fields("mm1").Value
                lblFwid.Caption = mod1.WP.Fields("mm2").Value
                lblLcRen.Caption = mod1.WP.Fields("mt1").Value
                lblLcUid.Caption = mod1.WP.Fields("mt2").Value
                lblTX.Caption = "下一流程,将跳至" & mod1.WP.Fields("mt3").Value & ": " & lblLcRen.Caption
                frmQm.Visible = False
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



Public Sub BoundForm(Bid As Long)
Dim tt As String
Dim Ra
Dim La
Dim oo As Integer
Dim ii As Integer
Dim jj As Integer

Call Me.initializeForm


'''''tt = "select lid,ljmc,jzpb+' '+jzxh,ljbh,sl,jdj,jhg,drq,brq,zbq,'',dj,hid,ldw,fj,uid,mj,ccbh from xunjiaMx where bid=" & Bid
tt = "select lid,partname,GG, bh,sl,jdj,jhg,drq,brq,zbq,'',dj,pid,ldw,fj,uid,mj,ccbh from xunjiaMlpg where bid=" & Bid
Set mod1.HTP = CreateObject("adodb.recordset")
On Error GoTo BoundFormERR:
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows

dtgMa.Visible = False
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
dtgMa.Rows = La + 20: dtgN.Rows = dtgMa.Rows
For oo = 1 To La
    dtgMa.Row = oo: dtgN.Row = oo
    For ii = 0 To 17
        dtgMa.Col = ii: dtgN.Col = ii
        dtgMa.Text = Ra(ii, oo - 1): dtgN.Text = dtgMa.Text: dtgMa.CellForeColor = &H80000012
'''''''''''        If ii = 15 Then
'''''''''''            If dtgMa.Text = "" And Val(lblLc.Caption) = 1 Then '如果未确定设定编号者,则不显示原厂编号
'''''''''''                dtgMa.Col = 3: dtgMa.Text = "": dtgN.Col = 3: dtgN.Text = ""
'''''''''''            ElseIf dtgMa.Text = lblUid.Caption Then '如果为业务员自己设定,则红色字显示
'''''''''''                dtgMa.Col = 3: dtgMa.CellForeColor = &HFF&
'''''''''''            Else
'''''''''''                dtgMa.Col = 3: dtgMa.CellForeColor = &H8000000D  '零件事业部设定的原厂编号,蓝色显示.
'''''''''''            End If
'''''''''''        End If
    Next
Next
dtgMa.Visible = True

'''''''''''''''显示商务支持添加的产品（变色）
''''''''''''''For oo = 1 To dtgMa.Rows
''''''''''''''    dtgMa.Col = 14
''''''''''''''    dtgMa.Row = oo
''''''''''''''    If dtgMa.Text = "True" Then
''''''''''''''        For jj = 1 To 16
''''''''''''''            dtgMa.Col = jj
''''''''''''''            dtgMa.CellForeColor = &H8000000D
''''''''''''''        Next
''''''''''''''    End If
''''''''''''''Next

Exit Sub
BoundFormERR:
    MsgBox "网络故障,请关闭再试!"
    End
End Sub

Private Sub txtAdr_Change()
cmdSave.Visible = True
End Sub

Private Sub txtBz_Change()
cmdSave.Enabled = True
End Sub


