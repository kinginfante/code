VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form htgX 
   Caption         =   "购销合同"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10455
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   10455
   Begin VB.ComboBox comKhmc 
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   6240
      TabIndex        =   100
      ToolTipText     =   "请在列表中选择客户"
      Top             =   0
      Width           =   4005
   End
   Begin VB.ComboBox txtT4 
      Height          =   300
      ItemData        =   "htgX.frx":0000
      Left            =   3780
      List            =   "htgX.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   98
      Top             =   3360
      Width           =   6405
   End
   Begin VB.CommandButton cmdPin 
      Caption         =   "评审单"
      Height          =   585
      Left            =   6240
      Picture         =   "htgX.frx":002E
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   10530
      Width           =   675
   End
   Begin VB.CommandButton cmdHg 
      Caption         =   "合计"
      Height          =   315
      Left            =   6390
      TabIndex        =   94
      Top             =   2310
      Width           =   705
   End
   Begin VB.TextBox txtHg 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   7110
      TabIndex        =   93
      Top             =   2280
      Width           =   1395
   End
   Begin VB.CommandButton cmdXMod1 
      Caption         =   "修改"
      Height          =   375
      Left            =   9690
      TabIndex        =   92
      Top             =   2250
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdXDel 
      Caption         =   "删除"
      Height          =   375
      Left            =   9150
      TabIndex        =   91
      Top             =   2250
      Width           =   495
   End
   Begin VB.CommandButton cmdXAdd 
      Caption         =   "添加"
      Height          =   375
      Left            =   8580
      TabIndex        =   90
      Top             =   2250
      Width           =   525
   End
   Begin MSComCtl2.DTPicker DTPQdDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy""年""M""月""d""日"""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   8550
      TabIndex        =   88
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      CalendarTitleBackColor=   16711680
      Format          =   109248513
      CurrentDate     =   38208
   End
   Begin VB.TextBox txtXGyzBM 
      Height          =   270
      Left            =   6810
      TabIndex        =   87
      Top             =   10230
      Width           =   3345
   End
   Begin VB.TextBox txtGyzBM 
      Height          =   270
      Left            =   1800
      TabIndex        =   86
      Top             =   10230
      Width           =   3105
   End
   Begin VB.TextBox txtZBQ 
      Height          =   270
      Left            =   9000
      TabIndex        =   82
      Top             =   3030
      Width           =   765
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印"
      Height          =   585
      Left            =   6930
      Picture         =   "htgX.frx":0470
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   10530
      Width           =   645
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "提交"
      Height          =   585
      Left            =   8250
      Picture         =   "htgX.frx":0ADA
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   10530
      Width           =   675
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   585
      Left            =   9600
      Picture         =   "htgX.frx":1144
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   10530
      Width           =   585
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "修改"
      Height          =   585
      Left            =   7590
      Picture         =   "htgX.frx":1246
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   10530
      Width           =   645
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除"
      Enabled         =   0   'False
      Height          =   585
      Left            =   8940
      Picture         =   "htgX.frx":1688
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   10530
      Width           =   645
   End
   Begin MSComCtl2.DTPicker dtpYXQ 
      Height          =   285
      Left            =   1980
      TabIndex        =   75
      Top             =   10650
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   503
      _Version        =   393216
      Format          =   109182977
      CurrentDate     =   38205
   End
   Begin VB.TextBox txtXGZH 
      Height          =   270
      Left            =   6810
      TabIndex        =   73
      Top             =   9930
      Width           =   3345
   End
   Begin VB.TextBox txtXGkhYY 
      Height          =   270
      Left            =   6810
      TabIndex        =   72
      Top             =   9630
      Width           =   3345
   End
   Begin VB.TextBox txtXGFH 
      Height          =   300
      Left            =   6810
      TabIndex        =   71
      Top             =   9300
      Width           =   3345
   End
   Begin VB.TextBox txtXGFX 
      Height          =   270
      Left            =   6810
      TabIndex        =   70
      Top             =   9000
      Width           =   3345
   End
   Begin VB.TextBox txtXGdW 
      Height          =   270
      Left            =   6810
      TabIndex        =   69
      Top             =   8730
      Width           =   3345
   End
   Begin VB.TextBox txtXGwiTo 
      Height          =   270
      Left            =   6810
      TabIndex        =   68
      Top             =   8460
      Width           =   3345
   End
   Begin VB.TextBox txtXGfdBr 
      Height          =   285
      Left            =   6810
      TabIndex        =   61
      Top             =   8160
      Width           =   3345
   End
   Begin VB.TextBox txtXdwAdr 
      Height          =   285
      Left            =   6810
      TabIndex        =   59
      Top             =   7890
      Width           =   3345
   End
   Begin VB.TextBox txtXdwMc 
      Height          =   285
      Left            =   6810
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   7620
      Width           =   3345
   End
   Begin VB.TextBox txtGZH 
      Height          =   270
      Left            =   1800
      TabIndex        =   54
      Top             =   9930
      Width           =   3105
   End
   Begin VB.TextBox txtGkhYY 
      Height          =   270
      Left            =   1800
      TabIndex        =   53
      Top             =   9630
      Width           =   3105
   End
   Begin VB.TextBox txtGFH 
      Height          =   270
      Left            =   1800
      TabIndex        =   52
      Top             =   9330
      Width           =   3105
   End
   Begin VB.TextBox txtGFX 
      Height          =   270
      Left            =   1800
      TabIndex        =   51
      Top             =   9030
      Width           =   3105
   End
   Begin VB.TextBox txtGdW 
      Height          =   270
      Left            =   1800
      TabIndex        =   50
      Top             =   8730
      Width           =   3105
   End
   Begin VB.TextBox txtGwiTo 
      Height          =   270
      Left            =   1800
      TabIndex        =   49
      Top             =   8460
      Width           =   3105
   End
   Begin VB.TextBox txtGfdBr 
      Height          =   285
      Left            =   1800
      TabIndex        =   41
      Top             =   8160
      Width           =   3105
   End
   Begin VB.TextBox txtGdwAdr 
      Height          =   285
      Left            =   1800
      TabIndex        =   39
      Top             =   7890
      Width           =   3105
   End
   Begin VB.TextBox txtGdwMc 
      Height          =   285
      Left            =   1800
      TabIndex        =   37
      Top             =   7620
      Width           =   3105
   End
   Begin VB.TextBox txtT14 
      BackColor       =   &H80000009&
      Height          =   270
      Left            =   3750
      TabIndex        =   34
      Top             =   6960
      Width           =   6405
   End
   Begin VB.TextBox txtT13 
      BackColor       =   &H80000009&
      Height          =   270
      Left            =   3750
      TabIndex        =   33
      Top             =   6600
      Width           =   6405
   End
   Begin VB.TextBox txtT12 
      BackColor       =   &H80000009&
      Height          =   270
      Left            =   3750
      TabIndex        =   32
      Top             =   6240
      Width           =   6405
   End
   Begin VB.TextBox txtT11 
      BackColor       =   &H80000009&
      Height          =   270
      Left            =   5070
      TabIndex        =   31
      Top             =   5880
      Width           =   5085
   End
   Begin VB.TextBox txtT10 
      BackColor       =   &H80000009&
      Height          =   270
      Left            =   3750
      TabIndex        =   30
      Top             =   5520
      Width           =   6405
   End
   Begin VB.TextBox txtT9 
      BackColor       =   &H80000009&
      Height          =   270
      Left            =   3750
      TabIndex        =   29
      Top             =   5160
      Width           =   6405
   End
   Begin VB.TextBox txtT8 
      BackColor       =   &H80000009&
      Height          =   270
      Left            =   3750
      TabIndex        =   28
      Top             =   4800
      Width           =   6405
   End
   Begin VB.TextBox txtT7 
      BackColor       =   &H80000009&
      Height          =   270
      Left            =   4380
      TabIndex        =   27
      Top             =   4440
      Width           =   5775
   End
   Begin VB.TextBox txtT6 
      BackColor       =   &H80000009&
      Height          =   270
      Left            =   3780
      TabIndex        =   26
      Top             =   4080
      Width           =   6375
   End
   Begin VB.TextBox txtT5 
      BackColor       =   &H80000009&
      Height          =   270
      Left            =   3780
      TabIndex        =   25
      Top             =   3720
      Width           =   6375
   End
   Begin VB.TextBox txtT3 
      BackColor       =   &H80000009&
      Height          =   270
      Left            =   3780
      TabIndex        =   24
      Top             =   3000
      Width           =   4365
   End
   Begin VB.TextBox txtT2 
      BackColor       =   &H80000009&
      Height          =   270
      Left            =   3780
      TabIndex        =   23
      Top             =   2640
      Width           =   6375
   End
   Begin VB.TextBox txtQyDD 
      Height          =   270
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   600
      Width           =   1125
   End
   Begin VB.TextBox txtHtbh 
      Height          =   270
      Left            =   6240
      TabIndex        =   21
      Top             =   300
      Width           =   4005
   End
   Begin VB.TextBox txtXF 
      Height          =   270
      Left            =   1350
      TabIndex        =   20
      Top             =   300
      Width           =   2925
   End
   Begin VB.TextBox txtGF 
      Height          =   270
      Left            =   1350
      TabIndex        =   19
      Top             =   0
      Width           =   2925
   End
   Begin MSDataGridLib.DataGrid dtgSale 
      Bindings        =   "htgX.frx":1812
      Height          =   1335
      Left            =   0
      TabIndex        =   89
      Top             =   900
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   2355
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   20
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "ljMc"
         Caption         =   "产品名称"
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
         DataField       =   "phBiao"
         Caption         =   "牌号商标"
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
         DataField       =   "ljBh"
         Caption         =   "规格型号"
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
         DataField       =   "jlDw"
         Caption         =   "单位"
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
      BeginProperty Column04 
         DataField       =   "ljSl"
         Caption         =   "数量"
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
      BeginProperty Column05 
         DataField       =   "dj"
         Caption         =   "单价"
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
      BeginProperty Column06 
         DataField       =   "je"
         Caption         =   "金额"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "JHT"
         Caption         =   "交货期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "款到n天内"
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
         RecordSelectors =   0   'False
         BeginProperty Column00 
            Object.Visible         =   -1  'True
            ColumnWidth     =   2399.811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2099.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1395.213
         EndProperty
      EndProperty
   End
   Begin VB.Label lblKhdh 
      Caption         =   "khdh"
      Height          =   225
      Left            =   4380
      TabIndex        =   101
      Top             =   150
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label47 
      Caption         =   "客户名称："
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5190
      TabIndex        =   99
      Top             =   30
      Width           =   945
   End
   Begin VB.Label lblDx 
      Height          =   345
      Left            =   3150
      TabIndex        =   96
      Top             =   2250
      Width           =   3105
   End
   Begin VB.Label Label46 
      Caption         =   "合计人民币（大写）"
      Height          =   255
      Left            =   1320
      TabIndex        =   95
      Top             =   2280
      Width           =   1665
   End
   Begin VB.Line Line1 
      X1              =   5040
      X2              =   5040
      Y1              =   7320
      Y2              =   10545
   End
   Begin VB.Label Label45 
      Caption         =   "邮政编码："
      Height          =   195
      Left            =   5730
      TabIndex        =   85
      Top             =   10260
      Width           =   1755
   End
   Begin VB.Label Label44 
      Caption         =   "邮政编码："
      Height          =   195
      Left            =   720
      TabIndex        =   84
      Top             =   10260
      Width           =   945
   End
   Begin VB.Label Label43 
      Caption         =   "月"
      Height          =   225
      Left            =   9870
      TabIndex        =   83
      Top             =   3060
      Width           =   285
   End
   Begin VB.Label Label42 
      Caption         =   "质保期"
      Height          =   225
      Left            =   8250
      TabIndex        =   81
      Top             =   3060
      Width           =   675
   End
   Begin VB.Label Label41 
      Caption         =   "有效期限：至"
      Height          =   285
      Left            =   660
      TabIndex        =   74
      Top             =   10680
      Width           =   1245
   End
   Begin VB.Label Label40 
      Caption         =   "账号："
      Height          =   225
      Left            =   6090
      TabIndex        =   67
      Top             =   9990
      Width           =   1455
   End
   Begin VB.Label Label39 
      Caption         =   "开户银行："
      Height          =   270
      Left            =   5730
      TabIndex        =   66
      Top             =   9660
      Width           =   1035
   End
   Begin VB.Label Label38 
      Caption         =   "国税号："
      Height          =   255
      Left            =   5910
      TabIndex        =   65
      Top             =   9330
      Width           =   1635
   End
   Begin VB.Label Label37 
      Caption         =   "传真："
      Height          =   225
      Left            =   6090
      TabIndex        =   64
      Top             =   9030
      Width           =   1425
   End
   Begin VB.Label Label36 
      Caption         =   "电话："
      Height          =   225
      Left            =   6120
      TabIndex        =   63
      Top             =   8760
      Width           =   1365
   End
   Begin VB.Label Label35 
      Caption         =   "委托代理人："
      Height          =   225
      Left            =   5580
      TabIndex        =   62
      Top             =   8460
      Width           =   1905
   End
   Begin VB.Label Label34 
      Caption         =   "法定代表人："
      Height          =   225
      Left            =   5580
      TabIndex        =   60
      Top             =   8190
      Width           =   1935
   End
   Begin VB.Label Label33 
      Caption         =   "单位地址："
      Height          =   225
      Left            =   5760
      TabIndex        =   58
      Top             =   7950
      Width           =   1755
   End
   Begin VB.Label Label32 
      Caption         =   "单位名称（章）："
      Height          =   255
      Left            =   5220
      TabIndex        =   56
      Top             =   7650
      Width           =   2295
   End
   Begin VB.Label Label31 
      Caption         =   "需    方"
      Height          =   195
      Left            =   7470
      TabIndex        =   55
      Top             =   7380
      Width           =   1695
   End
   Begin VB.Label Label30 
      Caption         =   "账号："
      Height          =   225
      Left            =   1080
      TabIndex        =   48
      Top             =   9990
      Width           =   645
   End
   Begin VB.Label Label29 
      Caption         =   "开户银行："
      Height          =   270
      Left            =   720
      TabIndex        =   47
      Top             =   9660
      Width           =   975
   End
   Begin VB.Label Label28 
      Caption         =   "开户银行："
      Height          =   375
      Left            =   1530
      TabIndex        =   46
      Top             =   9660
      Width           =   45
   End
   Begin VB.Label Label27 
      Caption         =   "国税号："
      Height          =   255
      Left            =   900
      TabIndex        =   45
      Top             =   9330
      Width           =   825
   End
   Begin VB.Label Label26 
      Caption         =   "传真："
      Height          =   225
      Left            =   1080
      TabIndex        =   44
      Top             =   9030
      Width           =   615
   End
   Begin VB.Label Label25 
      Caption         =   "电话："
      Height          =   225
      Left            =   1110
      TabIndex        =   43
      Top             =   8760
      Width           =   555
   End
   Begin VB.Label Label24 
      Caption         =   "委托代理人："
      Height          =   225
      Left            =   570
      TabIndex        =   42
      Top             =   8460
      Width           =   1095
   End
   Begin VB.Label Label23 
      Caption         =   "法定代表人："
      Height          =   225
      Left            =   570
      TabIndex        =   40
      Top             =   8190
      Width           =   1125
   End
   Begin VB.Label Label22 
      Caption         =   "单位地址："
      Height          =   225
      Left            =   750
      TabIndex        =   38
      Top             =   7950
      Width           =   945
   End
   Begin VB.Label Label21 
      Caption         =   "单位名称（章）："
      Height          =   255
      Left            =   210
      TabIndex        =   36
      Top             =   7650
      Width           =   1485
   End
   Begin VB.Label Label20 
      Caption         =   "供    方"
      Height          =   195
      Left            =   2430
      TabIndex        =   35
      Top             =   7380
      Width           =   885
   End
   Begin VB.Shape Shape1 
      Height          =   3225
      Left            =   150
      Top             =   7320
      Width           =   10125
   End
   Begin VB.Label Label19 
      Caption         =   "十四、其它约定事项"
      Height          =   285
      Left            =   240
      TabIndex        =   18
      Top             =   6990
      Width           =   2115
   End
   Begin VB.Label Label18 
      Caption         =   "十三、解决合同纠纷的方式	"
      Height          =   285
      Left            =   240
      TabIndex        =   17
      Top             =   6630
      Width           =   2505
   End
   Begin VB.Label Label17 
      Caption         =   "十二、违约责任"
      Height          =   285
      Left            =   240
      TabIndex        =   16
      Top             =   6270
      Width           =   1635
   End
   Begin VB.Label Label16 
      Caption         =   "十一、如需提供担保，另立合同担保书，作为本合同附件"
      Height          =   285
      Left            =   240
      TabIndex        =   15
      Top             =   5910
      Width           =   4725
   End
   Begin VB.Label Label15 
      Caption         =   "十、结算方式及期限"
      Height          =   285
      Left            =   240
      TabIndex        =   14
      Top             =   5550
      Width           =   2385
   End
   Begin VB.Label Label14 
      Caption         =   "九、合同生效期"
      Height          =   285
      Left            =   240
      TabIndex        =   13
      Top             =   5190
      Width           =   1605
   End
   Begin VB.Label Label13 
      Caption         =   "八、验收方式及提出异议期限"
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Top             =   4830
      Width           =   2685
   End
   Begin VB.Label Label12 
      Caption         =   "七、包装标准、包装物的供应与回收和费用负担"
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Top             =   4470
      Width           =   4035
   End
   Begin VB.Label Label11 
      Caption         =   "六、合理损耗计算方法"
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   4110
      Width           =   1905
   End
   Begin VB.Label Label10 
      Caption         =   "五、运输方式及到达站（港）的费用负担"
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Top             =   3750
      Width           =   3435
   End
   Begin VB.Label Label9 
      Caption         =   "四、交(提)货方式"
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   3390
      Width           =   2655
   End
   Begin VB.Label Label8 
      Caption         =   "三、供方对质量负责的条件和期限"
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   3030
      Width           =   2835
   End
   Begin VB.Label Label7 
      Caption         =   "二、质量要求技术标准"
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   2670
      Width           =   2115
   End
   Begin VB.Label Label6 
      Caption         =   "一、产品名称、商标、型号、数量、金额、供货时间及数量 "
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   630
      Width           =   4725
   End
   Begin VB.Label Label5 
      Caption         =   "签约时间："
      Height          =   225
      Left            =   7470
      TabIndex        =   4
      Top             =   630
      Width           =   1035
   End
   Begin VB.Label Label4 
      Caption         =   "签约地点："
      Height          =   195
      Left            =   5190
      TabIndex        =   3
      Top             =   630
      Width           =   945
   End
   Begin VB.Label Label3 
      Caption         =   "合同编号："
      Height          =   225
      Left            =   5190
      TabIndex        =   2
      Top             =   330
      Width           =   945
   End
   Begin VB.Label Label2 
      Caption         =   "需方："
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   330
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "供方："
      Height          =   225
      Left            =   240
      TabIndex        =   0
      Top             =   30
      Width           =   915
   End
End
Attribute VB_Name = "htgX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
Dim tt As String
On Error Resume Next

htgX.Visible = False
htBrow.Enabled = True
frmZu.Enabled = True
Call mod1.DelDKZ  '退出表单时删除打开记录,以让别人能打开此单据
htBrow.ZOrder 0
htBrow.Enabled = True
End Sub

Private Sub cmdDel_Click()
MsgBox "暂时无此功能！"
End Sub

Private Sub cmdDown_Click()

End Sub

Private Sub cmdHg_Click()
On Error Resume Next
Dim hg As Double
hg = 0
form2Htp.adoSale.Recordset.MoveFirst
Do While Not form2Htp.adoSale.Recordset.EOF
hg = hg + form2Htp.adoSale.Recordset.Fields("je").Value
form2Htp.adoSale.Recordset.MoveNext
Loop
hg = Round(hg, 2)
txtHg.Text = hg
lblDx.Caption = mod1.ChangBi(hg)
End Sub

Private Sub cmdMod_Click()
cmdSave.Enabled = True

cmdXMod1.Enabled = False
cmdXAdd.Enabled = True
cmdXDel.Enabled = True
dtgSale.AllowUpdate = True
If form2Htp.adoSale.Recordset.RecordCount = 0 Then
form2Htp.adoSale.Recordset.AddNew "htBh", form2Htp.txtHtbh.Text
End If
End Sub

Private Sub cmdPin_Click()
Dim tt As String
On Error Resume Next

tt = "Select * from htPing where htBh='" & txtHtbh.Text & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
form2Htp.lblHtxz.Caption = ""

Call modHt.htQing
Call modHt.htBound '绑定合同评审单字段
form2Htp.Show

'如果维修合同，则计算总工时，并列出出工列表
'If form2Htp.optA(1).Value = True Or form2Htp.optA(3).Value = True Or form2Htp.optA(4).Value = True Then

    
    
    '显示出工列表和零配件列表
    tt = "Select * from gzb where htbh='" & txtHtbh.Text & "' order by rq"
    form2Htp.adoGzb.Recordset.Close
    form2Htp.adoGzb.Recordset.Open tt, , adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set form2Htp.dtgGzb.DataSource = form2Htp.adoGzb
    tt = "Select * from linjian where gongFang='豪曼' and  bhid=any (select bhid from gzb where  htbh='" _
    & txtHtbh.Text & "')"
    form2Htp.adoLj.Recordset.Close
    form2Htp.adoLj.Recordset.Open tt, , adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set form2Htp.dtgLj.DataSource = form2Htp.adoLj
    
    form2Htp.cmdGzd.Caption = form2Htp.adoGzb.Recordset.Fields("bhId").Value
    '计算总工时
    tt = "Select sum(wTime) from workXx where bhid=any (Select bhid from gzb where htbh='" _
    & txtHtbh.Text & "')"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    form2Htp.lblZtime.Caption = Val(mod1.HTP.Fields(0).Value)
    mod1.HTP.Close
'End If

'设置权限
form2Htp.chkA.Enabled = False
form2Htp.chkB.Enabled = False
form2Htp.chkC.Enabled = False
form2Htp.chkD.Enabled = False
form2Htp.chkE.Enabled = False
form2Htp.cmdMod.Enabled = False
form2Htp.cmdSave.Enabled = False


'打开收款表
Dim ft As String
frmFuK.WindowState = 1
frmFuK.Visible = True

ft = "Select * from htPing1 where htBh='" & txtHtbh.Text & "' order by rq"
frmFuK.adoHpt.Recordset.Close
frmFuK.adoHpt.Recordset.Open ft, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText

'如果为没有记录，则添加一条空记录，以便用户编辑（用户可能提交时未填付款表，以后要填时用）
'If frmFuK.adoHpt.Recordset.RecordCount = 0 Then
'frmFuK.adoHpt.Recordset.AddNew
'frmFuK.adoHpt.Recordset.UpdateBatch
'End If
Set frmFuK.dtgFk.DataSource = frmFuK.adoHpt

ft = "Select * from yiFk Where htBh='" & frmFuK.adoHpt.Recordset.Fields("htBh").Value & _
"' and yingRQ='" & frmFuK.adoHpt.Recordset.Fields("rq").Value & "' order by yiRq"
frmFuK.adoYf.Recordset.Close
frmFuK.adoYf.Recordset.Open ft, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
Set frmFuK.dtgYf.DataSource = frmFuK.adoYf

'打开产品表
ft = ""
ft = "Select * from htSale Where htBh='" & txtHtbh.Text & "'"
form2Htp.adoSale.Recordset.Close
form2Htp.adoSale.Recordset.Open ft, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
Set form2Htp.dtgSale.DataSource = form2Htp.adoSale
Set form2Htp.dtgYJ.DataSource = form2Htp.adoSale
Set form2Htp.dtgZj.DataSource = form2Htp.adoSale

'ft = "Select * from CG Where htbh='" & txtHtbh.Text & "' and khmc<>'库存'"
'frmAdo.adoTmp.Recordset.Close
'frmAdo.adoTmp.Recordset.Open ft, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'Set form2Htp.dtgCG.DataSource = frmAdo.adoTmp

'打开佣金表
tt = "Select * from Yongjin where htBh='" & txtHtbh.Text & "' order by yId"
frmYj.adoYj.Recordset.Close
frmYj.adoYj.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
Set frmYj.dtgYJ.DataSource = frmYj.adoYj

''打开“取自库存表”
'tt = "Select * from kcJa where htBh='" & txtHtbh.Text & "'"
'form2Htp.adoKu.Recordset.Close
'form2Htp.adoKu.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'Set form2Htp.dtgKu.DataSource = form2Htp.adoKu


    '佣金、利润2、提成不显示
    form2Htp.txtYj1.Visible = False
    form2Htp.txtYj2.Visible = False
    form2Htp.txtLr1.Visible = False
    form2Htp.txtLr2.Visible = False
    form2Htp.txtTc1.Visible = False
    form2Htp.txtTc2.Visible = False
    form2Htp.lblYj.Visible = False
    form2Htp.lblLr2.Visible = False
    form2Htp.lblTC.Visible = False
    
form2Htp.tabHt.TabEnabled(1) = True
form2Htp.tabHt.TabEnabled(2) = True
'End If

'form2Htp.cmdXAdd.Enabled = False
'form2Htp.cmdXDel.Enabled = False
'form2Htp.cmdXMod1.Enabled = False
'form2Htp.cmdXMod1.Enabled = True
form2Htp.dtgSale.AllowUpdate = False
'form2Htp.txtLj.Visible = False
form2Htp.tabHt.Tab = 0
htgX.Visible = False
frmFuK.cmdAdd.Enabled = False
frmFuK.cmdDel.Enabled = False

If wbDN.Visible = True Then '如果为在客户资料中打开合同,则不能修改
    form2Htp.cmdMod.Enabled = False
    form2Htp.cmdSave.Enabled = False
End If
End Sub

Private Sub cmdPrint_Click()
'''''Dim tt As String
'''''On Error Resume Next
'''''If cmdSave.Enabled = False Then
'''''    Me.MousePointer = 11
'''''    Set mod1.report = mod1.crapp.OpenReport(App.Path & "\gxht.rpt")
'''''     'Set mod1.report = mod1.crapp.OpenReport(App.Path & "\tt.rpt")
'''''    Set mod1.table = mod1.report.Database.Tables
'''''    Set mod1.cProp = mod1.table.Item(1).ConnectionProperties
'''''    mod1.cProp.Item("Password") = "guyonghui"
'''''    mod1.report.SQLQueryString = "Select gxht.*,htSale.* from gxht cross join htSale where gxht.htbh='" & txtHtbh.Text & "' and gxht.htbh=htSale.htBh"
'''''    mod1.report.ReadRecords
'''''    frmReport.Show
'''''    frmReport.cR1.ReportSource = mod1.report
'''''    frmReport.cR1.ViewReport
'''''
'''''    Me.MousePointer = 0
'''''    frmReport.cR1.EnableExportButton = False
'''''    frmReport.cR1.EnableExportButton = True
'''''End If
End Sub

Private Sub cmdSave_Click()
Dim tt As String
If txtXF.Text = "" Then
MsgBox "请正确填写“需方”"
txtXF.SetFocus
Exit Sub
End If

If txtHtbh.Text = "" Then
MsgBox "请正确填写“合同编号”"
txtHtbh.SetFocus
Exit Sub
End If

If txtT9.Text = "" Then
MsgBox "请正确填写“合同生效期”"
txtT9.SetFocus
Exit Sub
End If

If txtXGwiTo.Text = "" Then
MsgBox "请填写“委托代理人”"
txtXGwiTo.SetFocus
Exit Sub
End If

If txtT4.Text = "" Then
MsgBox "请选择交提货方式"
Exit Sub
End If

Call modHt.gxAdd ''销售合同添加数据
cmdSave.Enabled = False

cmdPin.Enabled = True
End Sub

Private Sub cmdXAdd_Click()
On Error Resume Next
form2Htp.adoSale.Recordset.AddNew "htBh", htgX.txtHtbh.Text
form2Htp.adoSale.Recordset.Update "htF", 0
form2Htp.adoSale.Recordset.Update "delF", 1
form2Htp.adoSale.Recordset.Update "MT", 0 '购买状态
Set dtgSale.DataSource = form2Htp.adoSale
End Sub

Private Sub cmdXDel_Click()
On Error Resume Next
Dim tt As String
form2Htp.adoSale.Recordset.Delete adAffectCurrent
'form2Htp.adoSale.Recordset.UpdateBatch
'tt = "Select * from htSale where htbh='" & htgX.txtHtbh.Text & "'"
'form2Htp.adoSale.Recordset.Clone
'form2Htp.adoSale.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'Set dtgSale.DataSource = form2Htp.adoSale
End Sub

Private Sub Command1_Click()

End Sub

Private Sub dtgKhmc_DblClick()


End Sub


Private Sub dtgSale_AfterColUpdate(ByVal ColIndex As Integer)
form2Htp.adoSale.Recordset.Fields("je").Value = form2Htp.adoSale.Recordset.Fields("dj").Value * form2Htp.adoSale.Recordset.Fields("ljSl").Value
Set dtgSale.DataSource = form2Htp.adoSale
'form2Htp.adoSale.Recordset.Fields("fkSl").Value = form2Htp.adoSale.Recordset.Fields("ljSL").Value
form2Htp.adoSale.Recordset.Fields("xgSl").Value = form2Htp.adoSale.Recordset.Fields("ljSL").Value
form2Htp.adoSale.Recordset.Fields("xgSlD").Value = form2Htp.adoSale.Recordset.Fields("ljSL").Value
form2Htp.adoSale.Recordset.Fields("khmc").Value = txtXF.Text
End Sub

Private Sub dtgSale_GotFocus()
If txtXdwMc.Text = "" Then
    MsgBox "客户资料中不存在此客户，请在客户档案中正确添加！"
    txtXF.Text = ""
End If
End Sub

Private Sub dtgSale_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub



Private Sub Form_Load()
htgX.Height = 11475
htgX.Width = 10575
htgX.Top = 0
htgX.Left = 3000
'dtgKhmc.Visible = False
End Sub


Private Sub Text21_Change()

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim tt As String
On Error Resume Next
If MDI.Cq = False Then


htgX.Visible = False
frmZu.Enabled = True

htBrow.ZOrder 0
htBrow.Enabled = True
Call mod1.DelDKZ  '退出表单时删除打开记录,以让别人能打开此单据
Cancel = True
End If
End Sub

Private Sub txtHg_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
lblDx.Caption = mod1.ChangBi(txtHg.Text)
End If
End Sub

Private Sub txtXF_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim tt As String
'On Error Resume Next
'If KeyCode = 13 And txtXF.Text <> "" Then

End Sub

Private Sub txtZBQ_LostFocus()
If txtZBQ.Text <> "" Then
txtT3.Text = ""
txtT3.Text = "质保期" & txtZBQ.Text & "个月" & txtT3.Text
End If
End Sub
